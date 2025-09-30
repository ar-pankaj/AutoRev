# -*- coding: utf-8 -*-
# Align viewports on sheets by snapping the intersection of the bottom-most and left-most visible grids
# to a user-picked point + (X, Y) offset on the sheet.
#
# UI: Point selection runs before the main dialog. Offset units are fixed to mm.
#
# Author: Pankaj Prabhakar (Modified by Gemini)
# Revit: 2020+ (uses legacy viewport alignment for versions < 2022)
# Env: pyRevit / RevitPythonShell (IronPython)

from __future__ import division
import sys
import math

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup,
    Viewport, ViewType, XYZ, ViewSheet, ElementId
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogIcon
# Import for point selection and exception handling
from Autodesk.Revit.UI.Selection import ObjectSnapTypes
from Autodesk.Revit.Exceptions import OperationCanceledException

# Added pyrevit import for output handling
from pyrevit import script

# ----------------- UI (Windows Forms) -----------------
clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')

from System.Drawing import Point, Size
from System.Windows.Forms import (
    Form, Label, TextBox, ComboBox, Button, CheckBox, DialogResult, FormBorderStyle,
    FormStartPosition, ComboBoxStyle, GroupBox, RadioButton, ListBox, SelectionMode,
    TabControl, TabPage, AnchorStyles, AutoScaleMode
)

# ----------------- Helpers -----------------

def mm_to_ft(mm):
    """Metric to Revit internal feet."""
    return float(mm) * 0.00328083989501312  # 1 mm = 0.0032808399 ft

def in_to_ft(inches):
    return float(inches) / 12.0

def parse_float(text, default=0.0):
    try:
        return float(text)
    except:
        return float(default)

def to_view_xy(view, p):
    """Project a model point to the view's 2D basis (Right=X, Up=Y)."""
    origin = view.Origin
    vx = view.RightDirection
    vy = view.UpDirection
    v = p - origin
    return (v.DotProduct(vx), v.DotProduct(vy))

def from_view_xy(view, x, y):
    """Lift a view-plane (x,y) back to 3D model XYZ."""
    origin = view.Origin
    vx = view.RightDirection
    vy = view.UpDirection
    return origin + vx.Multiply(x) + vy.Multiply(y)

def curve_points_in_view_xy(view, curve, samples=11):
    """Sample curve in view-plane coordinates."""
    pts = []
    # Try tessellation first
    try:
        tess = curve.Tessellate()
        if tess and len(tess) >= 2:
            for p in tess:
                pts.append(to_view_xy(view, p))
            return pts
    except:
        pass
    # Fallback param sampling
    try:
        t0 = curve.GetEndParameter(0)
        t1 = curve.GetEndParameter(1)
        if samples < 2:
            samples = 2
        for i in range(samples):
            t = t0 + (t1 - t0) * i / float(samples - 1)
            p = curve.Evaluate(t, True)
            pts.append(to_view_xy(view, p))
        return pts
    except:
        pass
    # Endpoints as last resort
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        pts.append(to_view_xy(view, p0))
        pts.append(to_view_xy(view, p1))
    except:
        pass
    return pts

def classify_grid_orientation(view, curve, axis_tol):
    """
    Return ('vertical'|'horizontal'|None, x_metric, y_metric)
    Based on average direction vs X/Y axes in the view plane.
    """
    pts = curve_points_in_view_xy(view, curve)
    if not pts:
        return (None, None, None)

    xs = [xy[0] for xy in pts]
    ys = [xy[1] for xy in pts]
    x_avg = sum(xs) / float(len(xs))
    y_avg = sum(ys) / float(len(ys))

    if len(pts) >= 2:
        (x0, y0), (x1, y1) = pts[0], pts[-1]
        dx, dy = (x1 - x0), (y1 - y0)
        mag = math.hypot(dx, dy)
        if mag < 1e-9:
            return (None, x_avg, y_avg)
        ux, uy = dx / mag, dy / mag
    else:
        return (None, x_avg, y_avg)

    ax = abs(ux)  # alignment with X
    ay = abs(uy)  # alignment with Y

    if ay > ax and (ay - ax) > axis_tol:
        return ('vertical', x_avg, y_avg)
    elif ax > ay and (ax - ay) > axis_tol:
        return ('horizontal', x_avg, y_avg)
    else:
        return (None, x_avg, y_avg)

def find_bottom_left_grid_intersection(doc, view, axis_tol):
    """
    Find the model XYZ of (left-most vertical grid) x (bottom-most horizontal grid)
    based on visible grids in the given view.
    """
    grids = list(FilteredElementCollector(doc, view.Id)
                 .OfCategory(BuiltInCategory.OST_Grids)
                 .WhereElementIsNotElementType())
    if not grids:
        return None

    vertical_candidates = []
    horizontal_candidates = []

    for g in grids:
        crv = getattr(g, 'Curve', None)
        if crv is None:
            continue
        orient, x_avg, y_avg = classify_grid_orientation(view, crv, axis_tol)
        if orient == 'vertical':
            vertical_candidates.append((g, x_avg, y_avg))
        elif orient == 'horizontal':
            horizontal_candidates.append((g, x_avg, y_avg))

    if not vertical_candidates or not horizontal_candidates:
        return None

    # Left-most vertical grid (min x), bottom-most horizontal grid (min y)
    _, x_left, _ = min(vertical_candidates, key=lambda t: t[1])
    _, _, y_bot  = min(horizontal_candidates, key=lambda t: t[2])

    return from_view_xy(view, x_left, y_bot)

def get_titleblock_on_sheet(doc, sheet, strict_one_titleblock):
    """Check for title block presence based on strictness rule."""
    tblocks = list(FilteredElementCollector(doc, sheet.Id)
                      .OfCategory(BuiltInCategory.OST_TitleBlocks)
                      .WhereElementIsNotElementType())
    if not tblocks:
        return None
    if strict_one_titleblock and len(tblocks) != 1:
        return None
    return tblocks[0]

def get_sheet_point_from_model_point_legacy(view, viewport, model_point):
    """
    Calculates the sheet coordinate for a given model point.
    This is a fallback for Revit versions < 2022.
    It assumes the viewport is not rotated on the sheet.
    """
    sheet_center_pt = viewport.GetBoxCenter()
    crop_box = view.CropBox
    crop_center_2d_view_coords = (crop_box.Min + crop_box.Max) / 2.0
    anchor_2d_view_coords = to_view_xy(view, model_point)
    delta_x_model = anchor_2d_view_coords[0] - crop_center_2d_view_coords.X
    delta_y_model = anchor_2d_view_coords[1] - crop_center_2d_view_coords.Y
    scale = view.Scale
    delta_x_sheet = delta_x_model / scale
    delta_y_sheet = delta_y_model / scale
    return sheet_center_pt + XYZ(delta_x_sheet, delta_y_sheet, 0)

def get_allowed_plan_types():
    """Build a list of plan-like ViewType enum values that exist in this Revit version."""
    allowed = [ViewType.FloorPlan, ViewType.CeilingPlan]
    eng = getattr(ViewType, 'EngineeringPlan', None)
    if eng is not None:
        allowed.append(eng)
    structp = getattr(ViewType, 'StructuralPlan', None)
    if structp is not None:
        allowed.append(structp)
    return tuple(allowed)

ALLOWED_PLAN_TYPES = get_allowed_plan_types()

# ----------------- Windows Form -----------------

class AlignViewsForm(Form):
    def __init__(self, doc):
        self._doc = doc
        self.Text = "Align Viewports to Grid Intersection"
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterScreen
        self.ClientSize = Size(700, 580) # Reduced height
        self.MinimumSize = Size(700, 580)
        self.AutoScaleMode = AutoScaleMode.Font
        
        # ---- GroupBox for Alignment Settings ----
        gbSettings = GroupBox()
        gbSettings.Text = "Alignment Settings"
        gbSettings.Location = Point(10, 10)
        gbSettings.Size = Size(self.ClientSize.Width - 20, 150) # Reduced height
        gbSettings.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right

        y = 25
        pad = 30
        
        # REMOVED: Title Block Corner dropdown is no longer needed.

        lbl_offx = Label(Text="Offset X (+→) from Picked Point (mm):", Location=Point(15, y), AutoSize=True)
        self.tbOffX = TextBox(Text="0", Location=Point(260, y-3), Size=Size(140, 22))
        self.tbOffX.Anchor = AnchorStyles.Top | AnchorStyles.Left
        y += pad

        lbl_offy = Label(Text="Offset Y (+↑) from Picked Point (mm):", Location=Point(15, y), AutoSize=True)
        self.tbOffY = TextBox(Text="0", Location=Point(260, y-3), Size=Size(140, 22))
        self.tbOffY.Anchor = AnchorStyles.Top | AnchorStyles.Left
        y += pad + 5

        self.chkPlanOnly = CheckBox(Text="Process plan views only", Checked=True, Location=Point(15, y), AutoSize=True)
        y += 25
        self.chkStrictTB = CheckBox(Text="Process only sheets with exactly one title block", Checked=True, Location=Point(15, y), AutoSize=True)

        for c in [lbl_offx, self.tbOffX, lbl_offy, self.tbOffY,
                    self.chkPlanOnly, self.chkStrictTB]:
            gbSettings.Controls.Add(c)
        
        self.Controls.Add(gbSettings)

        # ---- GroupBox for Sheet Selection ----
        gbSheets = GroupBox()
        gbSheets.Text = "Select Sheets to Process"
        gbSheets.Location = Point(10, gbSettings.Bottom + 10)
        gbSheets.Size = Size(self.ClientSize.Width - 20, 350)
        gbSheets.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom

        yy = 25
        
        lbl_filter = Label(Text="Filter:", Location=Point(15, yy+3), AutoSize=True)
        self.tbFilter = TextBox(Location=Point(65, yy), Size=Size(275, 22))
        self.tbFilter.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.btnSelectAll = Button(Text="Select All", Location=Point(350, yy-1), Size=Size(90, 26))
        self.btnSelectAll.Anchor = AnchorStyles.Top | AnchorStyles.Right
        self.btnSelectNone = Button(Text="Select None", Location=Point(450, yy-1), Size=Size(100, 26))
        self.btnSelectNone.Anchor = AnchorStyles.Top | AnchorStyles.Right
        yy += 34

        self.lbSheets = ListBox(Location=Point(15, yy), Size=Size(gbSheets.Width - 30, 280))
        self.lbSheets.SelectionMode = SelectionMode.MultiExtended
        self.lbSheets.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom

        # Events
        self.tbFilter.TextChanged += self._apply_filter
        self.btnSelectAll.Click += self._select_all
        self.btnSelectNone.Click += self._select_none
        
        for c in [lbl_filter, self.tbFilter, self.btnSelectAll, self.btnSelectNone, self.lbSheets]:
            gbSheets.Controls.Add(c)
        
        self.Controls.Add(gbSheets)
        
        # populate sheet list
        self._all_sheets = list(FilteredElementCollector(self._doc).OfClass(ViewSheet))
        self._all_sheets.sort(key=lambda s: (s.SheetNumber or "").upper())
        self._sheet_items = []  # (label, ElementId)
        for sh in self._all_sheets:
            label = "{} - {}".format(sh.SheetNumber, sh.Name)
            self._sheet_items.append((label, sh.Id))
            self.lbSheets.Items.Add(label)
        self._sheet_items_filtered = [lbl for (lbl, _) in self._sheet_items]

        # ---- Buttons (bottom-right) ----
        self.btnOK = Button(Text="Run", Size=Size(90, 28))
        self.btnCancel = Button(Text="Cancel", Size=Size(90, 28))
        self.btnOK.Location = Point(self.ClientSize.Width - 200, self.ClientSize.Height - 50)
        self.btnCancel.Location = Point(self.ClientSize.Width - 100, self.ClientSize.Height - 50)
        self.btnOK.Anchor = AnchorStyles.Right | AnchorStyles.Bottom
        self.btnCancel.Anchor = AnchorStyles.Right | AnchorStyles.Bottom
        self.btnOK.Click += self.on_ok
        self.btnCancel.Click += self.on_cancel
        self.Controls.Add(self.btnOK)
        self.Controls.Add(self.btnCancel)

        self.Values = None

    def _apply_filter(self, sender, args):
        query = (self.tbFilter.Text or "").strip().lower()
        self.lbSheets.BeginUpdate()
        self.lbSheets.Items.Clear()
        if not query:
            self._sheet_items_filtered = [lbl for (lbl, _) in self._sheet_items]
        else:
            self._sheet_items_filtered = [lbl for (lbl, _) in self._sheet_items if query in lbl.lower()]
        for lbl in self._sheet_items_filtered:
            self.lbSheets.Items.Add(lbl)
        self.lbSheets.EndUpdate()

    def _select_all(self, sender, args):
        for i in range(self.lbSheets.Items.Count):
            self.lbSheets.SetSelected(i, True)

    def _select_none(self, sender, args):
        self.lbSheets.ClearSelected()

    def on_ok(self, sender, args):
        offx = parse_float(self.tbOffX.Text, 0.0)
        offy = parse_float(self.tbOffY.Text, 0.0)
        plan_only = bool(self.chkPlanOnly.Checked)
        strict_tb = bool(self.chkStrictTB.Checked)

        # Convert from fixed mm unit to feet
        offset_x_ft, offset_y_ft = mm_to_ft(offx), mm_to_ft(offy)

        selected_labels = list(self.lbSheets.SelectedItems)
        if not selected_labels:
            TaskDialog.Show("Align Viewports", "Please select at least one sheet from the list to process.")
            return
        label_to_id = dict(self._sheet_items)
        picklist_ids = [label_to_id[lbl] for lbl in selected_labels if lbl in label_to_id]

        self.Values = {
            "offset_x_ft": offset_x_ft, "offset_y_ft": offset_y_ft,
            "picklist_ids": picklist_ids, "plan_only": plan_only,
            "strict_tb": strict_tb, "axis_tol": 0.15 # Hardcoded value
        }
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

# ----------------- Main -----------------

if __name__ == '__main__':
    # pyRevit injection
    uiapp = __revit__
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document
    
    # --- MODIFIED: Point selection happens BEFORE the form is shown ---
    # 1. Check if the active view is a sheet
    if not isinstance(doc.ActiveView, ViewSheet):
        TaskDialog.Show(
            "Align Viewports",
            "You must run this script from a sheet view. Please open a sheet and try again.",
            TaskDialogIcon.TaskDialogIconWarning
        )
        sys.exit()

    # 2. Prompt user to pick the base point
    picked_point = None
    try:
        picked_point = uidoc.Selection.PickPoint(
            ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections,
            "Select the base alignment point on the sheet"
        )
    except OperationCanceledException:
        print("Script cancelled by user during point selection.")
        sys.exit()

    if not picked_point:
        print("ERROR: No point was selected.")
        sys.exit()

    # 3. Show the main form
    form = AlignViewsForm(doc)
    if form.ShowDialog() != DialogResult.OK or not form.Values:
        sys.exit()

    vals = form.Values
    
    # --- MODIFIED: Define the final target point using the picked point and offsets ---
    target_pt = picked_point + XYZ(vals["offset_x_ft"], vals["offset_y_ft"], 0.0)

    # Sheet collection is from the picklist
    picklist_ids = vals["picklist_ids"]
    sheets = [doc.GetElement(eid) for eid in picklist_ids if isinstance(eid, ElementId)]
    sheets = [s for s in sheets if isinstance(s, ViewSheet)]
    if not sheets:
        print("ERROR: No valid sheets were selected from the list.")
        sys.exit()

    processed_views = []
    skipped = []

    tg = TransactionGroup(doc, "Align Viewports to Grid Intersection")
    tg.Start()
    try:
        for sheet in sheets:
            # The check for title blocks is now only for filtering, not for positioning
            if vals["strict_tb"]:
                if not get_titleblock_on_sheet(doc, sheet, vals["strict_tb"]):
                    skipped.append((sheet.SheetNumber, "Missing or multiple title blocks"))
                    continue

            vports = list(FilteredElementCollector(doc, sheet.Id).OfClass(Viewport))
            if not vports:
                skipped.append((sheet.SheetNumber, "No viewports"))
                continue
            
            # REMOVED: No longer need to calculate base_corner per sheet. target_pt is global.

            t = Transaction(doc, "Align viewports on sheet {}".format(sheet.SheetNumber))
            t.Start()
            try:
                aligned_any = False
                for vp in vports:
                    view = doc.GetElement(vp.ViewId)
                    if not view or (vals["plan_only"] and view.ViewType not in ALLOWED_PLAN_TYPES):
                        continue

                    model_anchor = find_bottom_left_grid_intersection(doc, view, vals["axis_tol"])
                    if model_anchor is None:
                        continue
                    
                    current_sheet_pt = None
                    try:
                        if hasattr(vp, "GetTransform"): # Revit 2022+
                            transform = vp.GetTransform()
                            current_sheet_pt = transform.OfPoint(model_anchor)
                        else: # Legacy
                            current_sheet_pt = get_sheet_point_from_model_point_legacy(view, vp, model_anchor)
                    except Exception as e:
                        skipped.append((sheet.SheetNumber, "Error calculating transform: {}".format(e)))
                        continue
                    
                    if not current_sheet_pt:
                        continue

                    delta = target_pt - current_sheet_pt
                    new_center = vp.GetBoxCenter().Add(delta)
                    vp.SetBoxCenter(new_center)
                    aligned_any = True
                    processed_views.append(
                        (view.Name, "Aligned", "{} - {}".format(sheet.SheetNumber, sheet.Name))
                    )

                if aligned_any:
                    t.Commit()
                else:
                    t.RollBack()
                    skipped.append((sheet.SheetNumber, "No eligible viewports found (check grids/filters)"))
            except Exception as e:
                t.RollBack()
                skipped.append((sheet.SheetNumber, "Runtime Error: {}".format(e)))
        tg.Assimilate()
    except Exception as e:
        tg.RollBack()
        print("FATAL ERROR: An unexpected error occurred: {}".format(e))

    # --- Final Report ---
    output = script.get_output()
    output.set_title("Align Viewports Report")

    # --- Aligned Views Table ---
    if processed_views:
        sorted_processed = sorted(processed_views, key=lambda x: (x[2], x[0]))
        
        aligned_table_data = []
        for view_name, status, sheet_info in sorted_processed:
            styled_status = '<div style="color:green; font-weight:bold;">{}</div>'.format(status)
            aligned_table_data.append([sheet_info, view_name, styled_status])

        output.print_table(
            table_data=aligned_table_data,
            title="Aligned Views ({})".format(len(aligned_table_data)),
            columns=["Sheets", "View Name", "Status"]
        )

    # --- Skipped Sheets Table ---
    if skipped:
        sorted_skipped = sorted(list(set(skipped)))
        output.print_table(
            table_data=sorted_skipped,
            title="Skipped Sheets ({})".format(len(sorted_skipped)),
            columns=["Sheet", "Reason for Skipping"]
        )

    if not processed_views and not skipped:
        print("No sheets were selected or processed.")