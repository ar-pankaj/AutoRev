# -*- coding: utf-8 -*-

from __future__ import division

__title__ = "Align View\nin Sheets"

__author__ = "Pankaj Prabhakar"

__doc__ = """Align viewports on sheets by snapping the intersection of the bottom-most and left-most visible grids
to a user-picked point + (X, Y) offset on the sheet.

Drift-proof method:
  • Enter TVP (Temporary View Properties) to temporarily hide annotations/datums & disable annotation crop.
  • Use vp.GetBoxCenter() (CLEAN) as the sheet crop-center; map (model→crop-local) vector to sheet via scale & rotation.
  • Move viewport by delta = target_on_sheet - anchor_on_sheet.
  • Verify and apply a tiny residual nudge to eliminate 0.01 ft (~3.048 mm) quantization offsets.

Revit: 2020+  | Env: pyRevit / RevitPythonShell (IronPython)"""

import sys
import math
import clr

# Revit API
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup,
    Viewport, ViewType, XYZ, ViewSheet, ElementId,
    BuiltInParameter, Category, CategoryType, TemporaryViewMode,
    ViewportRotation
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogIcon
from Autodesk.Revit.UI.Selection import ObjectSnapTypes
from Autodesk.Revit.Exceptions import OperationCanceledException

# pyRevit
from pyrevit import script

# --------------------------- UI (Windows Forms) ---------------------------
clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System.Drawing import Point, Size
from System.Windows.Forms import (
    Form, Label, TextBox, Button, CheckBox, DialogResult, FormBorderStyle,
    FormStartPosition, GroupBox, ListBox, SelectionMode, AnchorStyles, AutoScaleMode
)

# --------------------------- Helpers: units & view-plane geometry ---------------------------
def mm_to_ft(mm): return float(mm) * 0.00328083989501312
def parse_float(text, default=0.0):
    try: return float(text)
    except: return float(default)

def to_view_xy(view, p):
    """Project a model point to the view's 2D basis (Right=X, Up=Y)."""
    origin = view.Origin
    vx = view.RightDirection
    vy = view.UpDirection
    v = p - origin
    return (v.DotProduct(vx), v.DotProduct(vy))

def from_view_xy(view, x, y):
    """Lift a view-plane (x,y) back to 3D model XYZ."""
    origin, vx, vy = view.Origin, view.RightDirection, view.UpDirection
    return origin + vx.Multiply(x) + vy.Multiply(y)

def curve_points_in_view_xy(view, curve, samples=11):
    pts = []
    try:
        tess = curve.Tessellate()
        if tess and len(tess) >= 2:
            for p in tess: pts.append(to_view_xy(view, p))
            return pts
    except: pass
    try:
        t0 = curve.GetEndParameter(0); t1 = curve.GetEndParameter(1)
        if samples < 2: samples = 2
        for i in range(samples):
            t = t0 + (t1 - t0) * i / float(samples - 1)
            p = curve.Evaluate(t, True)
            pts.append(to_view_xy(view, p))
        return pts
    except: pass
    try:
        p0 = curve.GetEndPoint(0); p1 = curve.GetEndPoint(1)
        pts.append(to_view_xy(view, p0)); pts.append(to_view_xy(view, p1))
    except: pass
    return pts

def classify_grid_orientation(view, curve, axis_tol):
    """
    Return ('vertical'|'horizontal'|None, x_metric, y_metric)
    Based on average direction vs X/Y axes in the view plane.
    """
    pts = curve_points_in_view_xy(view, curve)
    if not pts: return (None, None, None)
    xs = [xy[0] for xy in pts]; ys = [xy[1] for xy in pts]
    x_avg = sum(xs) / float(len(xs)); y_avg = sum(ys) / float(len(ys))
    if len(pts) >= 2:
        (x0,y0),(x1,y1) = pts[0], pts[-1]
        dx,dy = (x1-x0),(y1-y0); mag = math.hypot(dx,dy)
        if mag < 1e-9: return (None, x_avg, y_avg)
        ux,uy = dx/mag, dy/mag
    else:
        return (None, x_avg, y_avg)
    ax, ay = abs(ux), abs(uy)
    if ay > ax and (ay - ax) > axis_tol:   return ('vertical',   x_avg, y_avg)
    if ax > ay and (ax - ay) > axis_tol:   return ('horizontal', x_avg, y_avg)
    return (None, x_avg, y_avg)

def find_bottom_left_grid_intersection(doc, view, axis_tol):
    """Find model XYZ of (left-most vertical grid) x (bottom-most horizontal grid) visible in the given view."""
    grids = list(FilteredElementCollector(doc, view.Id)
                 .OfCategory(BuiltInCategory.OST_Grids)
                 .WhereElementIsNotElementType())
    if not grids: return None
    vertical, horizontal = [], []
    for g in grids:
        crv = getattr(g, 'Curve', None)
        if crv is None: continue
        orient, x_avg, y_avg = classify_grid_orientation(view, crv, axis_tol)
        if orient == 'vertical':     vertical.append((g, x_avg, y_avg))
        elif orient == 'horizontal': horizontal.append((g, x_avg, y_avg))
    if not vertical or not horizontal: return None
    _, x_left, _ = min(vertical,   key=lambda t: t[1])   # min x
    _, _, y_bot = min(horizontal,  key=lambda t: t[2])   # min y
    return from_view_xy(view, x_left, y_bot)

# --------------------- Drift-proof: TVP clean-up and crop→sheet mapping ---------------------
def can_hide_category(view, cat):
    try: return (cat is not None and cat.AllowsVisibilityControl and view.CanCategoryBeHidden(cat.Id))
    except: return False

def get_annotation_and_datum_categories(doc, view):
    """Collect annotation categories + common datums (robust across Revit versions)."""
    cats = []
    settings = doc.Settings
    for cat in settings.Categories:
        if cat is None: continue
        try:
            if cat.CategoryType == CategoryType.Annotation and can_hide_category(view, cat):
                cats.append(cat)
        except: pass
    datum_names = [
        "OST_Grids", "OST_Levels",
        "OST_ReferencePlanes", "OST_ReferenceLines",
        "OST_Sections", "OST_SectionHeads", "OST_SectionLine",
        "OST_ElevationMarks"
    ]
    for nm in datum_names:
        try:
            bic = getattr(BuiltInCategory, nm, None)
            if bic is None: continue
            dcat = Category.GetCategory(doc, bic)
            if can_hide_category(view, dcat) and dcat not in cats:
                cats.append(dcat)
        except: pass
    return cats

def hide_categories(view, cats, hidden=True):
    states = {}
    for cat in cats:
        try:
            cid = cat.Id
            cur = view.GetCategoryHidden(cid)
            states[cid] = cur
            if cur != hidden: view.SetCategoryHidden(cid, hidden)
        except: pass
    return states

def restore_categories(view, states):
    for cid, was_hidden in states.items():
        try:
            cur = view.GetCategoryHidden(cid)
            if cur != was_hidden: view.SetCategoryHidden(cid, was_hidden)
        except: pass

def try_hide_viewport_title(vp, want_hidden=True):
    """Hide/show viewport title by instance parameter if present. Returns (supported, original_value)."""
    try:
        p = vp.get_Parameter(BuiltInParameter.VIEWPORT_LABEL_VISIBLE)
        if p and not p.IsReadOnly:
            orig = p.AsInteger()
            newv = 0 if want_hidden else 1
            if orig != newv: p.Set(newv)
            return True, orig
    except: pass
    return False, None

def toggle_annotation_crop(view, want_active):
    """Toggle Annotation Crop if available. Returns (supported, original_value)."""
    try:
        p = view.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE)
        if p and not p.IsReadOnly:
            orig = p.AsInteger()
            newv = 1 if want_active else 0
            if orig != newv: p.Set(newv)
            return True, orig
    except: pass
    return False, None

def rot_apply(v_sheet_unrot, rotation_enum):
    """Rotate a 2D vector in sheet XY based on viewport rotation."""
    s = str(rotation_enum)
    if rotation_enum == ViewportRotation.None: return XYZ(v_sheet_unrot.X, v_sheet_unrot.Y, 0)
    if ("Ninety" in s and "Counter" in s) or ("NinetyDegrees" in s and "Clockwise" not in s):
        return XYZ(-v_sheet_unrot.Y,  v_sheet_unrot.X, 0)   # +90° CCW
    if "Clockwise" in s:
        return XYZ( v_sheet_unrot.Y, -v_sheet_unrot.X, 0)   # -90° CW
    if "OneHundredEighty" in s or "Halfway" in s or "180" in s:
        return XYZ(-v_sheet_unrot.X, -v_sheet_unrot.Y, 0)   # 180°
    return XYZ(v_sheet_unrot.X, v_sheet_unrot.Y, 0)

def local_vec_to_sheet(v_local, view_scale, vp_rotation):
    """Convert a crop-local vector (model feet) to sheet vector (feet), including rotation."""
    scaled = XYZ(v_local.X / view_scale, v_local.Y / view_scale, 0)
    return rot_apply(scaled, vp_rotation)

def compute_anchor_sheet_point_via_clean_crop(doc, view, vp, model_xyz):
    """
    Compute the sheet coordinate of a model XYZ robustly:
      - enable TVP (when available), hide annotations/datums & viewport title, disable annotation crop
      - use vp.GetBoxCenter() (CLEAN) as the sheet crop-center (avoids Outline min/max quantization)
      - map model -> crop-local via crop Transform; then local vector -> sheet via scale+rotation
      - restore temporary state
    """
    # Try to enable TVP
    tvp_enabled = False
    try:
        if hasattr(view, "EnableTemporaryViewPropertiesMode"):
            tvp_enabled = view.EnableTemporaryViewPropertiesMode(view.Id)
    except:
        tvp_enabled = False

    cats_to_hide = get_annotation_and_datum_categories(doc, view)
    saved_states = hide_categories(view, cats_to_hide, hidden=True)
    title_supported, title_orig = try_hide_viewport_title(vp, want_hidden=True)
    ac_supported, ac_orig = toggle_annotation_crop(view, want_active=False)

    try:
        # Recompute extents after temp visibility changes
        try: doc.Regenerate()
        except: pass

        # Cleaned inputs
        clean_center = vp.GetBoxCenter()  # sheet crop-center in the cleaned state
        cb = view.CropBox
        if cb is None:
            raise Exception("View has no CropBox.")
        T = cb.Transform
        view_scale = view.Scale
        vp_rot = vp.Rotation

        crop_ctr_local = (cb.Min + cb.Max) / 2.0

        # Model -> crop-local (use view plane z)
        pt_local = T.Inverse.OfPoint(XYZ(model_xyz.X, model_xyz.Y, T.Origin.Z))

        # Vector from local crop center to local point -> map into sheet
        v_local = pt_local - crop_ctr_local
        v_sheet = local_vec_to_sheet(v_local, view_scale, vp_rot)

        # Anchor on sheet
        return clean_center + v_sheet

    finally:
        # Restore everything
        if ac_supported and ac_orig is not None:
            try: toggle_annotation_crop(view, want_active=(ac_orig == 1))
            except: pass
        try: restore_categories(view, saved_states)
        except: pass
        if tvp_enabled:
            try: view.DisableTemporaryViewMode(TemporaryViewMode.TemporaryViewProperties)
            except: pass
        if title_supported and title_orig is not None:
            try: try_hide_viewport_title(vp, want_hidden=(title_orig == 0))
            except: pass

# --------------------------- Allowed view types ---------------------------
def get_allowed_plan_types():
    allowed = [ViewType.FloorPlan, ViewType.CeilingPlan]
    eng = getattr(ViewType, 'EngineeringPlan', None)
    if eng is not None: allowed.append(eng)
    structp = getattr(ViewType, 'StructuralPlan', None)
    if structp is not None: allowed.append(structp)
    return tuple(allowed)
ALLOWED_PLAN_TYPES = get_allowed_plan_types()

def get_titleblock_on_sheet(doc, sheet, strict_one_titleblock):
    tblocks = list(FilteredElementCollector(doc, sheet.Id)
                   .OfCategory(BuiltInCategory.OST_TitleBlocks)
                   .WhereElementIsNotElementType())
    if not tblocks: return None
    if strict_one_titleblock and len(tblocks) != 1: return None
    return tblocks[0]

# --------------------------- Windows Form ---------------------------
class AlignViewsForm(Form):
    def __init__(self, doc):
        self._doc = doc
        self.Text = "Align Viewports to Grid Intersection"
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterScreen
        self.ClientSize = Size(700, 580)
        self.MinimumSize = Size(700, 580)
        self.AutoScaleMode = AutoScaleMode.Font

        gbSettings = GroupBox()
        gbSettings.Text = "Alignment Settings"
        gbSettings.Location = Point(10, 10)
        gbSettings.Size = Size(self.ClientSize.Width - 20, 150)
        gbSettings.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        y = 25; pad = 30

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

        for c in [lbl_offx, self.tbOffX, lbl_offy, self.tbOffY, self.chkPlanOnly, self.chkStrictTB]:
            gbSettings.Controls.Add(c)
        self.Controls.Add(gbSettings)

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

        self.tbFilter.TextChanged += self._apply_filter
        self.btnSelectAll.Click += self._select_all
        self.btnSelectNone.Click += self._select_none

        for c in [lbl_filter, self.tbFilter, self.btnSelectAll, self.btnSelectNone, self.lbSheets]:
            gbSheets.Controls.Add(c)
        self.Controls.Add(gbSheets)

        # populate sheet list
        self._all_sheets = list(FilteredElementCollector(self._doc).OfClass(ViewSheet))
        self._all_sheets.sort(key=lambda s: (s.SheetNumber or "").upper())
        self._sheet_items = []
        for sh in self._all_sheets:
            label = "{} - {}".format(sh.SheetNumber, sh.Name)
            self._sheet_items.append((label, sh.Id))
            self.lbSheets.Items.Add(label)
        self._sheet_items_filtered = [lbl for (lbl, _) in self._sheet_items]

        self.btnOK = Button(Text="Run", Size=Size(90, 28))
        self.btnCancel = Button(Text="Cancel", Size=Size(90, 28))
        self.btnOK.Location = Point(self.ClientSize.Width - 200, self.ClientSize.Height - 50)
        self.btnCancel.Location = Point(self.ClientSize.Width - 100, self.ClientSize.Height - 50)
        self.btnOK.Anchor = AnchorStyles.Right | AnchorStyles.Bottom
        self.btnCancel.Anchor = AnchorStyles.Right | AnchorStyles.Bottom
        self.btnOK.Click += self.on_ok
        self.btnCancel.Click += self.on_cancel
        self.Controls.Add(self.btnOK); self.Controls.Add(self.btnCancel)

        self.Values = None

    def _apply_filter(self, sender, args):
        query = (self.tbFilter.Text or "").strip().lower()
        self.lbSheets.BeginUpdate(); self.lbSheets.Items.Clear()
        if not query:
            self._sheet_items_filtered = [lbl for (lbl, _) in self._sheet_items]
        else:
            self._sheet_items_filtered = [lbl for (lbl, _) in self._sheet_items if query in lbl.lower()]
        for lbl in self._sheet_items_filtered: self.lbSheets.Items.Add(lbl)
        self.lbSheets.EndUpdate()

    def _select_all(self, sender, args):
        for i in range(self.lbSheets.Items.Count): self.lbSheets.SetSelected(i, True)

    def _select_none(self, sender, args): self.lbSheets.ClearSelected()

    def on_ok(self, sender, args):
        offx = parse_float(self.tbOffX.Text, 0.0)
        offy = parse_float(self.tbOffY.Text, 0.0)
        plan_only = bool(self.chkPlanOnly.Checked)
        strict_tb = bool(self.chkStrictTB.Checked)
        selected_labels = list(self.lbSheets.SelectedItems)
        if not selected_labels:
            TaskDialog.Show("Align Viewports", "Please select at least one sheet from the list to process.")
            return
        label_to_id = dict(self._sheet_items)
        picklist_ids = [label_to_id[lbl] for lbl in selected_labels if lbl in label_to_id]
        self.Values = {
            "offset_x_ft": mm_to_ft(offx), "offset_y_ft": mm_to_ft(offy),
            "picklist_ids": picklist_ids, "plan_only": plan_only,
            "strict_tb": strict_tb, "axis_tol": 0.15
        }
        self.DialogResult = DialogResult.OK; self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel; self.Close()

# --------------------------- Main ---------------------------
if __name__ == '__main__':
    uiapp = __revit__
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document

    # Must be run from a sheet to pick a sheet point
    if not isinstance(doc.ActiveView, ViewSheet):
        TaskDialog.Show("Align Viewports",
                        "You must run this script from a sheet view. Please open a sheet and try again.",
                        TaskDialogIcon.TaskDialogIconWarning)
        sys.exit()

    # Pick base alignment point on the sheet
    try:
        picked_point = uidoc.Selection.PickPoint(
            ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections,
            "Select the base alignment point on the sheet"
        )
    except OperationCanceledException:
        print("Script cancelled by user during point selection.")
        sys.exit()

    if not picked_point:
        print("ERROR: No point was selected."); sys.exit()

    # Show options dialog
    form = AlignViewsForm(doc)
    if form.ShowDialog() != DialogResult.OK or not form.Values:
        sys.exit()
    vals = form.Values
    target_pt = picked_point + XYZ(vals["offset_x_ft"], vals["offset_y_ft"], 0.0)

    # Resolve selected sheets
    picklist_ids = vals["picklist_ids"]
    sheets = [doc.GetElement(eid) for eid in picklist_ids if isinstance(eid, ElementId)]
    sheets = [s for s in sheets if isinstance(s, ViewSheet)]
    if not sheets:
        print("ERROR: No valid sheets were selected from the list."); sys.exit()

    processed_views, skipped = [], []

    tg = TransactionGroup(doc, "Align Viewports to Grid Intersection (drift-proof)")
    tg.Start()
    try:
        for sheet in sheets:
            if vals["strict_tb"]:
                if not get_titleblock_on_sheet(doc, sheet, vals["strict_tb"]):
                    skipped.append((sheet.SheetNumber, "Missing or multiple title blocks")); continue

            vports = list(FilteredElementCollector(doc, sheet.Id).OfClass(Viewport))
            if not vports:
                skipped.append((sheet.SheetNumber, "No viewports")); continue

            t = Transaction(doc, "Align viewports on sheet {}".format(sheet.SheetNumber))
            t.Start()
            try:
                aligned_any = False
                for vp in vports:
                    view = doc.GetElement(vp.ViewId)
                    if not view or (vals["plan_only"] and view.ViewType not in ALLOWED_PLAN_TYPES):
                        continue
                    if not getattr(view, "CropBoxActive", False):
                        continue

                    # Model anchor: bottom-left visible grid intersection in this view
                    model_anchor = find_bottom_left_grid_intersection(doc, view, vals["axis_tol"])
                    if model_anchor is None:
                        continue

                    try:
                        # ---- Drift-proof anchor position on SHEET (TVP clean-up + center-based) ----
                        sheet_anchor = compute_anchor_sheet_point_via_clean_crop(doc, view, vp, model_anchor)

                        # Move vector that hits the picked target
                        delta = target_pt - sheet_anchor

                        # 1st move
                        vp.SetBoxCenter(vp.GetBoxCenter() + delta)

                        # Optional: verify & nudge once (eliminates sub-centimeter residuals like 0.01 ft)
                        try: doc.Regenerate()
                        except: pass
                        rez = target_pt - compute_anchor_sheet_point_via_clean_crop(doc, view, vp, model_anchor)
                        if abs(rez.X) > 1e-5 or abs(rez.Y) > 1e-5:
                            vp.SetBoxCenter(vp.GetBoxCenter() + XYZ(rez.X, rez.Y, 0.0))

                        aligned_any = True
                        processed_views.append((view.Name, "Aligned", "{} - {}".format(sheet.SheetNumber, sheet.Name)))

                    except Exception as ee:
                        skipped.append((sheet.SheetNumber, "Error aligning viewport: {}".format(ee)))
                        continue

                if aligned_any: t.Commit()
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

    if processed_views:
        sorted_processed = sorted(processed_views, key=lambda x: (x[2], x[0]))
        aligned_table_data = []
        for view_name, status, sheet_info in sorted_processed:
            styled_status = '<div style="color:green; font-weight:bold;">{}</div>'.format(status)
            aligned_table_data.append([sheet_info, view_name, styled_status])
        output.print_table(table_data=aligned_table_data,
                           title="Aligned Views ({})".format(len(aligned_table_data)),
                           columns=["Sheets", "View Name", "Status"])
    if skipped:
        sorted_skipped = sorted(list(set(skipped)))
        output.print_table(table_data=sorted_skipped,
                           title="Skipped Sheets ({})".format(len(sorted_skipped)),
                           columns=["Sheet", "Reason for Skipping"])

    # (This is the line you said it stopped at; leaving the last print in place)
    if not processed_views and not skipped:
        print("No sheets were selected or processed.")