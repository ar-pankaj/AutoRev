# -*- coding: utf-8 -*-
# Align viewports on sheets by snapping the intersection of the bottom-most and left-most visible grids
# to a point defined by a title block corner + (X, Y) offset on the sheet.
#
# UI: Tabbed dialog (Settings | Sheets) with multi-select sheet picker and filter.
#
# Author: M365 Copilot (for Pankaj Prabhakar)
# Revit: 2020+ (requires Viewport.GetTransform)
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
from Autodesk.Revit.UI import TaskDialog

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

def get_titleblock_bbox_on_sheet(doc, sheet, strict_one_titleblock):
    """Return (bbmin, bbmax) in sheet coordinates."""
    tblocks = list(FilteredElementCollector(doc, sheet.Id)
                   .OfCategory(BuiltInCategory.OST_TitleBlocks)
                   .WhereElementIsNotElementType())
    if not tblocks:
        return None
    if strict_one_titleblock and len(tblocks) != 1:
        return None
    tb = tblocks[0]
    bb = tb.get_BoundingBox(sheet)
    if not bb:
        return None
    return (bb.Min, bb.Max)

def corner_point_from_bbox(corner_name, bbmin, bbmax):
    if corner_name == "Bottom-Left":
        return XYZ(bbmin.X, bbmin.Y, 0.0)
    if corner_name == "Bottom-Right":
        return XYZ(bbmax.X, bbmin.Y, 0.0)
    if corner_name == "Top-Left":
        return XYZ(bbmin.X, bbmax.Y, 0.0)
    if corner_name == "Top-Right":
        return XYZ(bbmax.X, bbmax.Y, 0.0)
    return None

def viewport_transform_safe(vp):
    if hasattr(vp, "GetTransform"):
        return vp.GetTransform()
    raise NotImplementedError("Viewport.GetTransform() not available in this Revit version.")

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
        self.ClientSize = Size(700, 560)
        self.MinimumSize = Size(700, 560)
        self.AutoScaleMode = AutoScaleMode.Font

        # ---- Tab control ----
        self.tabs = TabControl()
        self.tabs.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
        self.tabs.Location = Point(10, 10)
        self.tabs.Size = Size(self.ClientSize.Width - 20, self.ClientSize.Height - 80)

        self.tabSettings = TabPage("Settings")
        self.tabSettings.AutoScroll = True
        self.tabSheets = TabPage("Sheets")
        self.tabSheets.AutoScroll = True

        self.tabs.TabPages.Add(self.tabSettings)
        self.tabs.TabPages.Add(self.tabSheets)
        self.Controls.Add(self.tabs)

        # ---- Settings tab ----
        y = 15
        pad = 30

        lbl_corner = Label(Text="Title Block Corner:", Location=Point(15, y), AutoSize=True)
        self.cbCorner = ComboBox(Location=Point(200, y-3), Size=Size(200, 22), DropDownStyle=ComboBoxStyle.DropDownList)
        for s in ["Bottom-Left", "Bottom-Right", "Top-Left", "Top-Right"]:
            self.cbCorner.Items.Add(s)
        self.cbCorner.SelectedIndex = 0
        self.cbCorner.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        y += pad

        lbl_units = Label(Text="Offset Units:", Location=Point(15, y), AutoSize=True)
        self.cbUnits = ComboBox(Location=Point(200, y-3), Size=Size(200, 22), DropDownStyle=ComboBoxStyle.DropDownList)
        for s in ["mm", "inches"]:
            self.cbUnits.Items.Add(s)
        self.cbUnits.SelectedIndex = 0
        self.cbUnits.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        y += pad

        lbl_offx = Label(Text="Offset X (+→):", Location=Point(15, y), AutoSize=True)
        self.tbOffX = TextBox(Text="20", Location=Point(200, y-3), Size=Size(200, 22))
        self.tbOffX.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        y += pad

        lbl_offy = Label(Text="Offset Y (+↑):", Location=Point(15, y), AutoSize=True)
        self.tbOffY = TextBox(Text="20", Location=Point(200, y-3), Size=Size(200, 22))
        self.tbOffY.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        y += pad

        self.chkPlanOnly = CheckBox(Text="Plan views only", Checked=True, Location=Point(15, y))
        y += 24
        self.chkStrictTB = CheckBox(Text="Strict: exactly 1 title block", Checked=True, Location=Point(15, y))
        y += 24

        lbl_tol = Label(Text="Axis tolerance (0..1):", Location=Point(15, y), AutoSize=True)
        self.tbTol = TextBox(Text="0.15", Location=Point(200, y-3), Size=Size(200, 22))
        self.tbTol.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        y += pad

        # add to Settings tab
        for c in [lbl_corner, self.cbCorner, lbl_units, self.cbUnits,
                  lbl_offx, self.tbOffX, lbl_offy, self.tbOffY,
                  self.chkPlanOnly, self.chkStrictTB, lbl_tol, self.tbTol]:
            self.tabSettings.Controls.Add(c)

        # ---- Sheets tab ----
        yy = 15

        self.rbAll = RadioButton(Text="All sheets", Location=Point(15, yy), Checked=True)
        yy += 24
        self.rbSelected = RadioButton(Text="Selected sheets (current selection)", Location=Point(15, yy))
        yy += 24
        self.rbPickList = RadioButton(Text="Pick from list", Location=Point(15, yy))
        yy += 30

        lbl_filter = Label(Text="Filter:", Location=Point(30, yy+3), AutoSize=True)
        self.tbFilter = TextBox(Location=Point(80, yy), Size=Size(260, 22))
        self.btnSelectAll = Button(Text="Select All", Location=Point(350, yy-1), Size=Size(90, 26))
        self.btnSelectNone = Button(Text="Select None", Location=Point(450, yy-1), Size=Size(100, 26))
        yy += 34

        self.lbSheets = ListBox(Location=Point(30, yy), Size=Size(520, 300))
        self.lbSheets.SelectionMode = SelectionMode.MultiExtended
        self.lbSheets.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom

        # Events
        self.rbAll.CheckedChanged += self._scope_changed
        self.rbSelected.CheckedChanged += self._scope_changed
        self.rbPickList.CheckedChanged += self._scope_changed
        self.tbFilter.TextChanged += self._apply_filter
        self.btnSelectAll.Click += self._select_all
        self.btnSelectNone.Click += self._select_none

        # add to Sheets tab
        for c in [self.rbAll, self.rbSelected, self.rbPickList,
                  lbl_filter, self.tbFilter, self.btnSelectAll, self.btnSelectNone, self.lbSheets]:
            self.tabSheets.Controls.Add(c)

        # populate sheet list
        self._all_sheets = list(FilteredElementCollector(self._doc).OfClass(ViewSheet))
        self._all_sheets.sort(key=lambda s: (s.SheetNumber or "").upper())
        self._sheet_items = []  # (label, ElementId)
        for sh in self._all_sheets:
            label = "{} - {}".format(sh.SheetNumber, sh.Name)
            self._sheet_items.append((label, sh.Id))
            self.lbSheets.Items.Add(label)
        self._sheet_items_filtered = [lbl for (lbl, _) in self._sheet_items]

        # Initially disable picklist controls
        self._enable_picklist(False)

        # ---- Buttons (bottom-right) ----
        self.btnOK = Button(Text="Run", Size=Size(90, 28))
        self.btnCancel = Button(Text="Cancel", Size=Size(90, 28))
        # position relative to form size
        self.btnOK.Location = Point(self.ClientSize.Width - 200, self.ClientSize.Height - 50)
        self.btnCancel.Location = Point(self.ClientSize.Width - 100, self.ClientSize.Height - 50)
        self.btnOK.Anchor = AnchorStyles.Right | AnchorStyles.Bottom
        self.btnCancel.Anchor = AnchorStyles.Right | AnchorStyles.Bottom

        self.btnOK.Click += self.on_ok
        self.btnCancel.Click += self.on_cancel

        self.Controls.Add(self.btnOK)
        self.Controls.Add(self.btnCancel)

        self.Values = None

    # ---- Scope helpers ----
    def _enable_picklist(self, enabled):
        self.lbSheets.Enabled = enabled
        self.tbFilter.Enabled = enabled
        self.btnSelectAll.Enabled = enabled
        self.btnSelectNone.Enabled = enabled

    def _scope_changed(self, sender, args):
        self._enable_picklist(self.rbPickList.Checked)

    def _apply_filter(self, sender, args):
        query = (self.tbFilter.Text or "").strip().lower()
        self.lbSheets.Items.Clear()
        if not query:
            self._sheet_items_filtered = [lbl for (lbl, _) in self._sheet_items]
        else:
            self._sheet_items_filtered = [lbl for (lbl, _) in self._sheet_items if query in lbl.lower()]
        for lbl in self._sheet_items_filtered:
            self.lbSheets.Items.Add(lbl)

    def _select_all(self, sender, args):
        for i in range(self.lbSheets.Items.Count):
            self.lbSheets.SetSelected(i, True)

    def _select_none(self, sender, args):
        for i in range(self.lbSheets.Items.Count):
            self.lbSheets.SetSelected(i, False)

    # ---- OK / Cancel ----
    def on_ok(self, sender, args):
        corner = self.cbCorner.SelectedItem
        units = self.cbUnits.SelectedItem
        offx = parse_float(self.tbOffX.Text, 0.0)
        offy = parse_float(self.tbOffY.Text, 0.0)
        plan_only = bool(self.chkPlanOnly.Checked)
        strict_tb = bool(self.chkStrictTB.Checked)
        axis_tol = parse_float(self.tbTol.Text, 0.15)

        # Convert to feet
        if units == "mm":
            offset_x_ft = mm_to_ft(offx)
            offset_y_ft = mm_to_ft(offy)
        else:
            offset_x_ft = in_to_ft(offx)
            offset_y_ft = in_to_ft(offy)

        # Determine scope
        if self.rbSelected.Checked:
            scope_mode = "selected"
            picklist_ids = None
        elif self.rbPickList.Checked:
            scope_mode = "picklist"
            # map selected labels back to ElementIds
            selected_labels = [self.lbSheets.Items[i] for i in range(self.lbSheets.Items.Count) if self.lbSheets.GetSelected(i)]
            if not selected_labels:
                TaskDialog.Show("Align Viewports", "Please select at least one sheet from the list or choose a different scope.")
                return
            label_to_id = dict(self._sheet_items)  # label -> ElementId
            picklist_ids = [label_to_id[lbl] for lbl in selected_labels if lbl in label_to_id]
        else:
            scope_mode = "all"
            picklist_ids = None

        self.Values = {
            "corner": corner,
            "offset_x_ft": offset_x_ft,
            "offset_y_ft": offset_y_ft,
            "scope_mode": scope_mode,
            "picklist_ids": picklist_ids,
            "plan_only": plan_only,
            "strict_tb": strict_tb,
            "axis_tol": axis_tol
        }
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

# ----------------- Main -----------------

# pyRevit injection
uiapp = __revit__
uidoc = uiapp.ActiveUIDocument
doc = uidoc.Document

# Show dialog
form = AlignViewsForm(doc)
res = form.ShowDialog()

if res != DialogResult.OK or not form.Values:
    sys.exit()

corner = form.Values["corner"]
offset_x_ft = form.Values["offset_x_ft"]
offset_y_ft = form.Values["offset_y_ft"]
scope_mode = form.Values["scope_mode"]           # "all" | "selected" | "picklist"
picklist_ids = form.Values["picklist_ids"]       # list[ElementId] or None
plan_only = form.Values["plan_only"]
strict_tb = form.Values["strict_tb"]
axis_tol = form.Values["axis_tol"]

# Collect sheets per scope
if scope_mode == "selected":
    sel_ids = list(uidoc.Selection.GetElementIds())
    sheets = []
    for eid in sel_ids:
        el = doc.GetElement(eid)
        if not el:
            continue
        if isinstance(el, ViewSheet):
            if el not in sheets:
                sheets.append(el)
            continue
        # Resolve owner sheet if an element on a sheet is selected
        try:
            owner_view_id = el.OwnerViewId
            if owner_view_id and owner_view_id.IntegerValue != -1:
                vw = doc.GetElement(owner_view_id)
                if isinstance(vw, ViewSheet) and vw not in sheets:
                    sheets.append(vw)
        except:
            pass
    if not sheets:
        TaskDialog.Show("Align Viewports", "No sheets found in current selection.")
        sys.exit()
elif scope_mode == "picklist":
    sheets = [doc.GetElement(eid) for eid in picklist_ids if isinstance(eid, ElementId)]
    sheets = [s for s in sheets if isinstance(s, ViewSheet)]
    if not sheets:
        TaskDialog.Show("Align Viewports", "No valid sheets selected from the list.")
        sys.exit()
else:
    sheets = list(FilteredElementCollector(doc).OfClass(ViewSheet))

processed = 0
skipped = []

tg = TransactionGroup(doc, "Align Viewports to Grid Intersection")
tg.Start()

try:
    for sheet in sheets:
        # Viewports on this sheet only (scoped collector)
        vports = list(FilteredElementCollector(doc, sheet.Id).OfClass(Viewport))
        if not vports:
            skipped.append((sheet.SheetNumber, "No viewports"))
            continue

        bb = get_titleblock_bbox_on_sheet(doc, sheet, strict_tb)
        if not bb:
            skipped.append((sheet.SheetNumber, "Missing/Multiple title blocks (strict mode) or invalid bbox"))
            continue

        bbmin, bbmax = bb
        base_corner = corner_point_from_bbox(corner, bbmin, bbmax)
        if not base_corner:
            skipped.append((sheet.SheetNumber, "Invalid corner selection"))
            continue

        target_pt = XYZ(base_corner.X + offset_x_ft, base_corner.Y + offset_y_ft, 0.0)

        t = Transaction(doc, "Align viewports on sheet {}".format(sheet.SheetNumber))
        t.Start()
        try:
            aligned_any = False

            for vp in vports:
                view = doc.GetElement(vp.ViewId)
                if not view:
                    continue
                if plan_only and (view.ViewType not in ALLOWED_PLAN_TYPES):
                    continue

                # Find grid intersection in model space
                model_anchor = find_bottom_left_grid_intersection(doc, view, axis_tol)
                if model_anchor is None:
                    continue

                # Map that model point to current sheet coords via viewport transform
                try:
                    xf = viewport_transform_safe(vp)
                except NotImplementedError:
                    skipped.append((sheet.SheetNumber, "Viewport transform unsupported"))
                    continue

                current_sheet_pt = xf.OfPoint(model_anchor)

                # Compute delta and move viewport (via box center)
                delta = XYZ(target_pt.X - current_sheet_pt.X, target_pt.Y - current_sheet_pt.Y, 0.0)
                new_center = vp.GetBoxCenter().Add(delta)
                vp.SetBoxCenter(new_center)
                aligned_any = True

            if aligned_any:
                t.Commit()
                processed += 1
            else:
                t.RollBack()
                skipped.append((sheet.SheetNumber, "No eligible viewports (no grids / filtered by options)"))
        except Exception as e:
            t.RollBack()
            skipped.append((sheet.SheetNumber, "Error: {}".format(e)))

    tg.Assimilate()
except Exception as e:
    tg.RollBack()
    raise

# Report
msg = "Done.\nSheets processed (committed): {}\n".format(processed)
if skipped:
    msg += "Skipped:\n" + "\n".join([" - {}: {}".format(s[0], s[1]) for s in skipped])
TaskDialog.Show("Align Viewports", msg)
print(msg)
