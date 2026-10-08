# -*- coding: utf-8 -*-

__title__ = "Grid\nDimensions"
__author__ = "Pankaj"
__doc__ = """Grid-to-Grid Top & Left Dimensions (Chain + Optional Overall)
Stable UI (no SplitContainer), DPI-aware, scrollable settings, bottom command bar (OK/Cancel always visible).
Features:
- Search + multi-select views
- Select All / Clear buttons for views (aligned, left-anchored)
- SHIFT + click range selection for views
- Dimension Type selection (Linear-only)
- Baseline: Nearest Parallel Grids (Top offsets from HORIZONTAL; Left from VERTICAL) or Crop Edge
- Chain Offset (mm) — independent Top/Left with "Same for both"
- Overall Dimension (Top & Left) with independent extra offsets and "Same for both"
- Option to select all created dimensions and save them as a Selection Set named after the first view
- NEW: Option to keep created dimensions visible ONLY in their owner view (hide from all other selected views)"""

import clr
import System

# Revit API
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewPlan, View, Grid, XYZ, Line, ReferenceArray, Reference,
    Transaction, TransactionGroup, DatumExtentType, UnitUtils, ElementId, BuiltInParameter,
    SelectionFilterElement, DimensionType, Dimension  # <-- Added Dimension
)

# Units compatibility (pre-/post-2021)
try:
    from Autodesk.Revit.DB import UnitTypeId
    HAS_UNITTYPEID = True
except:
    from Autodesk.Revit.DB import DisplayUnitType
    HAS_UNITTYPEID = False

# Dimension types + style enum (when available)
try:
    from Autodesk.Revit.DB import DimensionStyleType
    HAS_DIMSTYLETYPE = True
except:
    HAS_DIMSTYLETYPE = False

# WinForms
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
import System.Windows.Forms as WinForms
from System.Drawing import Size, Point
from System import Decimal, IntPtr
from System.Windows.Forms import (
    Form, Label, TextBox, CheckedListBox, Button, ComboBox, ComboBoxStyle,
    GroupBox, RadioButton, NumericUpDown, MessageBox, DialogResult,
    IWin32Window, AutoScaleMode, AnchorStyles, CheckBox, Panel,
    TableLayoutPanel, FlowLayoutPanel, DockStyle, Padding, Keys, CheckState
)

# pyRevit handles
__doc__ = 'Add grid-to-grid dimensions (Top & Left) in multiple views: chain + optional overall, with search, linear-only dim types, and robust baseline options.'
__author__ = 'Pankaj'
uiapp = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document


# -------------------------
# Helpers
# -------------------------
def mm_to_internal(mm):
    """Convert millimeters to internal feet."""
    if HAS_UNITTYPEID:
        return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)
    else:
        return UnitUtils.ConvertToInternalUnits(float(mm), DisplayUnitType.DUT_MILLIMETERS)


class JtWindowHandle(IWin32Window):
    """Wrap Revit main window handle so ShowDialog has the correct owner."""
    def __init__(self, hwnd):
        self._handle = IntPtr(hwnd)

    @property
    def Handle(self):
        return self._handle


def safe_view_name(v):
    """Reliable view name; fallback to BuiltInParameter to avoid IronPython AttributeError."""
    try:
        nm = v.Name
        if nm:
            return nm
    except:
        pass
    try:
        p = v.get_Parameter(BuiltInParameter.VIEW_NAME)
        if p:
            nm = p.AsString()
            if nm:
                return nm
    except:
        pass
    return u"View {}".format(v.Id.IntegerValue)


def safe_type_name(dt):
    """Reliable type name; fallback to SYMBOL_NAME_PARAM."""
    try:
        nm = dt.Name
        if nm:
            return nm
    except:
        pass
    try:
        p = dt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            nm = p.AsString()
            if nm:
                return nm
    except:
        pass
    return u"DimType {}".format(dt.Id.IntegerValue)


def collect_plan_views(document):
    """Collect non-template, printable plan views."""
    views = (FilteredElementCollector(document)
             .OfClass(ViewPlan)
             .ToElements())
    result = [v for v in views if (not v.IsTemplate and v.CanBePrinted)]
    result.sort(key=lambda v: safe_view_name(v))
    return result


def collect_linear_dimension_types(document):
    """Collect only Linear dimension types (robust across versions)."""
    all_dt = list(FilteredElementCollector(document).OfClass(DimensionType))
    linear = []
    for dt in all_dt:
        try:
            if HAS_DIMSTYLETYPE and hasattr(dt, 'StyleType'):
                if str(dt.StyleType) == str(DimensionStyleType.Linear):
                    linear.append(dt)
                continue
        except:
            pass
        nm = (safe_type_name(dt) or "").lower()
        if ("linear" in nm) or ("aligned" in nm):
            linear.append(dt)
    seen = set(); out = []
    for dt in linear:
        if dt.Id.IntegerValue not in seen:
            out.append(dt); seen.add(dt.Id.IntegerValue)
    out.sort(key=lambda x: safe_type_name(x))
    return out


def get_grid_curve_in_view(grid, view):
    """Prefer curve as shown in view; fallback to model extents if needed."""
    curves = grid.GetCurvesInView(DatumExtentType.ViewSpecific, view)
    if curves and len(curves) > 0:
        return curves[0]
    curves = grid.GetCurvesInView(DatumExtentType.Model, view)
    if curves and len(curves) > 0:
        return curves[0]
    return None


def classify_grids_by_orientation(grids, view):
    """Return (vertical_grids, horizontal_grids) in the view basis."""
    vertical_grids = []
    horizontal_grids = []
    up = view.UpDirection.Normalize()
    right = view.RightDirection.Normalize()
    for g in grids:
        c = get_grid_curve_in_view(g, view)
        if c is None:
            continue
        dirv = (c.GetEndPoint(1) - c.GetEndPoint(0))
        if dirv.IsZeroLength():
            continue
        dirv = dirv.Normalize()
        if abs(dirv.DotProduct(up)) >= abs(dirv.DotProduct(right)):
            vertical_grids.append(g)
        else:
            horizontal_grids.append(g)

    def _uniq(lst):
        seen = set(); out = []
        for g in lst:
            k = g.Id.IntegerValue
            if k not in seen:
                out.append(g); seen.add(k)
        return out

    return _uniq(vertical_grids), _uniq(horizontal_grids)


def extents_in_view_coords(grids, view):
    """Return (minX, maxX, minY, maxY) for given grids in the view coordinate frame."""
    if not grids:
        return None
    cb = view.CropBox
    T = cb.Transform
    Ti = T.Inverse
    xs = []; ys = []
    for g in grids:
        c = get_grid_curve_in_view(g, view)
        if c is None:
            continue
        p0v = Ti.OfPoint(c.GetEndPoint(0))
        p1v = Ti.OfPoint(c.GetEndPoint(1))
        xs.extend([p0v.X, p1v.X]); ys.extend([p0v.Y, p1v.Y])
    if not xs or not ys:
        return None
    return (min(xs), max(xs), min(ys), max(ys))


def sort_grids_by_view_axis(grids, view, axis='x'):
    """Sort grids by midpoint along X or Y in view coordinates (for overall first/last)."""
    if not grids:
        return []
    cb = view.CropBox
    Ti = cb.Transform.Inverse

    def keyfn(g):
        c = get_grid_curve_in_view(g, view)
        if c is None:
            return 0.0
        p0v = Ti.OfPoint(c.GetEndPoint(0))
        p1v = Ti.OfPoint(c.GetEndPoint(1))
        midx = 0.5*(p0v.X + p1v.X)
        midy = 0.5*(p0v.Y + p1v.Y)
        return midx if axis == 'x' else midy

    return sorted(grids, key=keyfn)


# -------------------------
# Baselines (separate Top/Left offsets)
# -------------------------
def baseline_from_crop(view, off_top_mm, off_left_mm):
    cb = view.CropBox
    if not view.CropBoxActive or cb is None:
        raise Exception("View has no active crop region. Turn it on before running.")
    T = cb.Transform
    xMin, xMax = cb.Min.X, cb.Max.X
    yMin, yMax = cb.Min.Y, cb.Max.Y
    zPlan = cb.Min.Z
    off_top = mm_to_internal(off_top_mm)
    off_left = mm_to_internal(off_left_mm)

    # Top = horizontal line above
    topP1 = T.OfPoint(XYZ(xMin, yMax + off_top, zPlan))
    topP2 = T.OfPoint(XYZ(xMax, yMax + off_top, zPlan))
    # Left = vertical line left of
    leftP1 = T.OfPoint(XYZ(xMin - off_left, yMin, zPlan))
    leftP2 = T.OfPoint(XYZ(xMin - off_left, yMax, zPlan))
    return Line.CreateBound(topP1, topP2), Line.CreateBound(leftP1, leftP2)


def baseline_from_nearest_parallel(view, vertical_grids, horizontal_grids, off_top_mm, off_left_mm):
    """
    Nearest Parallel Grids logic:
    - TOP baseline offsets from topmost HORIZONTAL grid (parallel) and spans VERTICAL grids.
    - LEFT baseline offsets from leftmost VERTICAL grid (parallel) and spans HORIZONTAL grids.
    """
    cb = view.CropBox
    if not view.CropBoxActive or cb is None:
        raise Exception("View has no active crop region. Turn it on before running.")
    T = cb.Transform
    zPlan = cb.Min.Z
    off_top = mm_to_internal(off_top_mm)
    off_left = mm_to_internal(off_left_mm)

    v_ext = extents_in_view_coords(vertical_grids, view)
    h_ext = extents_in_view_coords(horizontal_grids, view)

    # TOP (horizontal)
    yTop = (h_ext[3] + off_top) if h_ext else (cb.Max.Y + off_top)
    xMin_top, xMax_top = (v_ext[0], v_ext[1]) if v_ext else (cb.Min.X, cb.Max.X)
    topP1 = T.OfPoint(XYZ(xMin_top, yTop, zPlan))
    topP2 = T.OfPoint(XYZ(xMax_top, yTop, zPlan))
    lineTop = Line.CreateBound(topP1, topP2)

    # LEFT (vertical)
    xLeft = (v_ext[0] - off_left) if v_ext else (cb.Min.X - off_left)
    yMin_left, yMax_left = (h_ext[2], h_ext[3]) if h_ext else (cb.Min.Y, cb.Max.Y)
    leftP1 = T.OfPoint(XYZ(xLeft, yMin_left, zPlan))
    leftP2 = T.OfPoint(XYZ(xLeft, yMax_left, zPlan))
    lineLeft = Line.CreateBound(leftP1, leftP2)

    return lineTop, lineLeft


# -------------------------
# Dimensions
# -------------------------
def new_dimension(doc, view, line, refs, dimtype):
    """Create dimension and return the created element (or None)."""
    if refs and refs.Size >= 2:
        if dimtype:
            return doc.Create.NewDimension(view, line, refs, dimtype)
        else:
            return doc.Create.NewDimension(view, line, refs)
    return None


def add_grid_dimensions_in_view(doc, view,
                                off_top_mm=120.0,
                                off_left_mm=120.0,
                                baseline_mode='grids',
                                dimtype=None,
                                add_overall=False,
                                overall_top_mm=120.0,
                                overall_left_mm=120.0):
    """Create chain (and optional overall) grid dimensions in a view and return list of created ElementIds."""
    created = []
    grids = (FilteredElementCollector(doc, view.Id)
             .OfClass(Grid)
             .ToElements())
    if len(grids) < 2:
        return created

    vertical_grids, horizontal_grids = classify_grids_by_orientation(grids, view)

    refsTop = ReferenceArray()
    for g in vertical_grids:
        refsTop.Append(Reference(g))
    refsLeft = ReferenceArray()
    for g in horizontal_grids:
        refsLeft.Append(Reference(g))

    if baseline_mode == 'crop':
        lineTop_chain, lineLeft_chain = baseline_from_crop(view, off_top_mm, off_left_mm)
    else:
        lineTop_chain, lineLeft_chain = baseline_from_nearest_parallel(view, vertical_grids, horizontal_grids, off_top_mm, off_left_mm)

    # Chain
    d1 = new_dimension(doc, view, lineTop_chain, refsTop, dimtype)
    if d1: created.append(d1.Id)
    d2 = new_dimension(doc, view, lineLeft_chain, refsLeft, dimtype)
    if d2: created.append(d2.Id)

    # Overall (optional)
    if add_overall:
        top_combo = float(off_top_mm) + float(overall_top_mm)
        left_combo = float(off_left_mm) + float(overall_left_mm)
        if baseline_mode == 'crop':
            lineTop_ov, lineLeft_ov = baseline_from_crop(view, top_combo, left_combo)
        else:
            lineTop_ov, lineLeft_ov = baseline_from_nearest_parallel(view, vertical_grids, horizontal_grids, top_combo, left_combo)

        v_sorted = sort_grids_by_view_axis(vertical_grids, view, axis='x')
        if len(v_sorted) >= 2:
            refsTopOverall = ReferenceArray()
            refsTopOverall.Append(Reference(v_sorted[0]))
            refsTopOverall.Append(Reference(v_sorted[-1]))
            d3 = new_dimension(doc, view, lineTop_ov, refsTopOverall, dimtype)
            if d3: created.append(d3.Id)

        h_sorted = sort_grids_by_view_axis(horizontal_grids, view, axis='y')
        if len(h_sorted) >= 2:
            refsLeftOverall = ReferenceArray()
            refsLeftOverall.Append(Reference(h_sorted[0]))
            refsLeftOverall.Append(Reference(h_sorted[-1]))
            d4 = new_dimension(doc, view, lineLeft_ov, refsLeftOverall, dimtype)
            if d4: created.append(d4.Id)

    doc.Regenerate()
    return created


# ---- robust helper — typed .NET Array[ElementId] via Array.CreateInstance ----
from System import Array  # use the Array type directly
def to_elementid_array(ids_iterable):
    """
    Build a System.Array[Autodesk.Revit.DB.ElementId] without using generic subscript syntax.
    This avoids IronPython's generic parsing quirks.
    """
    buf = []
    for eid in ids_iterable:
        if isinstance(eid, ElementId):
            buf.append(eid)
        else:
            try:
                buf.append(ElementId(int(eid)))
            except:
                pass
    eid_type = clr.GetClrType(ElementId)  # System.Type for Autodesk.Revit.DB.ElementId
    arr = Array.CreateInstance(eid_type, len(buf))
    for i, val in enumerate(buf):
        arr[i] = val
    return arr


# ---- NEW: Isolate created dimensions to their owner view during this run ----
def isolate_created_dims_to_owner_view(doc, owner_view, all_other_views, created_ids):
    """
    Hide 'created_ids' from all views in 'all_other_views' (except owner_view),
    so they remain visible ONLY in their owner view. Safe-checks CanBeHidden per view.
    """
    if not created_ids:
        return
    # Fetch element objects once
    created_elems = []
    for eid in created_ids:
        try:
            el = doc.GetElement(eid)
            if el:
                created_elems.append(el)
        except:
            pass

    for vw in all_other_views:
        if not isinstance(vw, View):
            continue
        if vw.Id == owner_view.Id:
            continue
        to_hide = []
        for el in created_elems:
            try:
                # If the element category can be hidden in this view, queue it
                if el.CanBeHidden(vw):
                    to_hide.append(el.Id)
            except:
                # Some rare elements may not support CanBeHidden(view)
                pass
        if to_hide:
            vw.HideElements(to_elementid_array(to_hide))


# -------------------------
# WinForms UI
# -------------------------
class ViewPickerForm(Form):
    def __init__(self, viewlist, dimtypes):
        Form.__init__(self)

        # Form base
        self.Text = "Grid Top+Left Dimensions"
        self.StartPosition = WinForms.FormStartPosition.CenterParent
        self.AutoScaleMode = AutoScaleMode.Dpi
        self.MinimizeBox = False
        self.MaximizeBox = True
        self.Padding = Padding(8)
        self.ClientSize = Size(900, 700)
        self.MinimumSize = Size(720, 560)

        # Data maps
        self.view_map = {}
        self.dimtype_map = {}

        # Views (safe names + Id)
        self._all_items = []
        for v in viewlist:
            nm = safe_view_name(v)
            disp = u"{0} (Id {1})".format(nm, v.Id.IntegerValue)
            self._all_items.append(disp)
            self.view_map[disp] = v.Id
        self._filtered_items = list(self._all_items)
        self._checked_ids = set()  # default: select none

        # Dimension types (Linear-only; list provided by main())
        self.dimtype_list = []
        for dt in dimtypes:
            disp = u"{0} (Id {1})".format(safe_type_name(dt), dt.Id.IntegerValue)
            self.dimtype_list.append(disp)
            self.dimtype_map[disp] = dt.Id

        # SHIFT-range selection state
        self._lastClickedIndex = -1
        self._mouseDownIndex = -1
        self._isBatchChecking = False

        # Root layout: 3 rows (Search, Content, Buttons)
        root = TableLayoutPanel()
        root.Dock = DockStyle.Fill
        root.Padding = Padding(0)
        root.Margin = Padding(0)
        root.ColumnCount = 1
        root.RowCount = 3
        root.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 100))
        root.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.AutoSize))
        root.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 100))
        root.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.AutoSize))
        self.Controls.Add(root)

        # Row 0: Search
        panelSearch = Panel()
        panelSearch.Dock = DockStyle.Top
        panelSearch.Height = 36
        lblSearch = Label(); lblSearch.Text = "Search Views:"; lblSearch.AutoSize = True; lblSearch.Location = Point(0, 10)
        self.txtSearch = TextBox()
        self.txtSearch.Location = Point(110, 6)
        self.txtSearch.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.txtSearch.Width = 700
        self.txtSearch.TextChanged += self._on_search
        panelSearch.Controls.Add(lblSearch)
        panelSearch.Controls.Add(self.txtSearch)
        root.Controls.Add(panelSearch, 0, 0)

        # Row 1: Content (list + settings)
        content = TableLayoutPanel()
        content.Dock = DockStyle.Fill
        content.Padding = Padding(0)
        content.Margin = Padding(0)
        content.ColumnCount = 1
        content.RowCount = 2
        content.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 100))
        content.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 60))
        content.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 40))
        root.Controls.Add(content, 0, 1)

        # List area
        listArea = TableLayoutPanel()
        listArea.Dock = DockStyle.Fill
        listArea.Padding = Padding(0)
        listArea.Margin = Padding(0)
        listArea.ColumnCount = 1
        listArea.RowCount = 2
        listArea.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 100))
        listArea.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.AutoSize))
        listArea.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 100))
        content.Controls.Add(listArea, 0, 0)

        toolbar = FlowLayoutPanel()
        toolbar.Dock = DockStyle.Top
        toolbar.AutoSize = True
        toolbar.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink
        toolbar.WrapContents = False
        toolbar.Padding = Padding(0, 0, 0, 6)
        toolbar.FlowDirection = WinForms.FlowDirection.LeftToRight

        self.btnAllViews = Button(); self.btnAllViews.Text = "Select All"; self.btnAllViews.Width = 100
        self.btnAllViews.Margin = Padding(0, 0, 8, 0)
        self.btnAllViews.Click += self._on_all

        self.btnNoneViews = Button(); self.btnNoneViews.Text = "Clear"; self.btnNoneViews.Width = 100
        self.btnNoneViews.Margin = Padding(0, 0, 16, 0)
        self.btnNoneViews.Click += self._on_none

        lblHint = Label()
        lblHint.Text = "Tip: Shift-click to toggle a range"
        lblHint.AutoSize = True
        lblHint.Margin = Padding(0, 6, 0, 0)

        toolbar.Controls.Add(self.btnAllViews)
        toolbar.Controls.Add(self.btnNoneViews)
        toolbar.Controls.Add(lblHint)
        listArea.Controls.Add(toolbar, 0, 0)

        panelList = Panel(); panelList.Dock = DockStyle.Fill
        self.chkViews = CheckedListBox()
        self.chkViews.Dock = DockStyle.Fill
        self.chkViews.CheckOnClick = True
        self.chkViews.IntegralHeight = False
        panelList.Controls.Add(self.chkViews)
        listArea.Controls.Add(panelList, 0, 1)

        # SHIFT-range events
        self.chkViews.MouseDown += self._on_list_mouse_down
        self.chkViews.ItemCheck += self._on_item_check

        # Settings area
        panelSettings = Panel(); panelSettings.Dock = DockStyle.Fill; panelSettings.AutoScroll = True
        content.Controls.Add(panelSettings, 0, 1)

        settings = TableLayoutPanel()
        settings.Parent = panelSettings
        settings.Dock = DockStyle.Top
        settings.AutoSize = True
        settings.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink
        settings.Padding = Padding(0)
        settings.Margin = Padding(0)
        settings.ColumnCount = 2
        settings.RowCount = 0
        settings.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.AutoSize))
        settings.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 100))

        # Row: Dimension Type (Linear-only)
        lblDimType = Label(); lblDimType.Text = "Dimension Type (Linear only):"; lblDimType.AutoSize = True
        self.cmbDimType = ComboBox()
        self.cmbDimType.DropDownStyle = ComboBoxStyle.DropDownList
        self.cmbDimType.Width = 420
        self.cmbDimType.Items.Add("<Use Current Default>")
        for disp in [u"{} (Id {})".format(safe_type_name(dt), dt.Id.IntegerValue) for dt in dimtypes]:
            self.cmbDimType.Items.Add(disp)
        self.cmbDimType.SelectedIndex = 0
        settings.Controls.Add(lblDimType, 0, settings.RowCount)
        settings.Controls.Add(self.cmbDimType, 1, settings.RowCount)
        settings.RowCount += 1

        # Row: Baseline group (Chain)
        grpBase = GroupBox()
        grpBase.Text = "Baseline Reference (for Chain)"
        grpBase.AutoSize = True
        grpBase.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink
        grpBase.Dock = DockStyle.Top

        baseFlow = FlowLayoutPanel()
        baseFlow.Dock = DockStyle.Top
        baseFlow.AutoSize = True
        baseFlow.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink
        baseFlow.WrapContents = False
        baseFlow.Padding = Padding(8, 6, 8, 6)
        baseFlow.Margin = Padding(8, 4, 8, 6)
        baseFlow.FlowDirection = WinForms.FlowDirection.LeftToRight

        self.rbGrids = RadioButton()
        self.rbGrids.Text = "Nearest Parallel Grids"
        self.rbGrids.AutoSize = True
        self.rbGrids.Checked = True
        self.rbGrids.Margin = Padding(4, 4, 16, 4)

        self.rbCrop = RadioButton()
        self.rbCrop.Text = "Crop Edge"
        self.rbCrop.AutoSize = True
        self.rbCrop.Margin = Padding(0, 4, 24, 4)

        # Same for both + inputs
        self.chkSameChain = CheckBox()
        self.chkSameChain.Text = "Same offset for Top & Left"
        self.chkSameChain.AutoSize = True
        self.chkSameChain.Checked = True
        self.chkSameChain.Margin = Padding(0, 4, 24, 4)

        lblTop = Label(); lblTop.Text = "Top Offset (mm):"; lblTop.AutoSize = True; lblTop.Margin = Padding(0, 7, 6, 4)
        self.numOffsetTop = NumericUpDown()
        self.numOffsetTop.Minimum = Decimal(0)
        self.numOffsetTop.Maximum = Decimal(35000)
        self.numOffsetTop.DecimalPlaces = 0
        self.numOffsetTop.Value = Decimal(0)
        self.numOffsetTop.Width = 80
        self.numOffsetTop.Margin = Padding(0, 4, 16, 4)

        lblLeft = Label(); lblLeft.Text = "Left Offset (mm):"; lblLeft.AutoSize = True; lblLeft.Margin = Padding(0, 7, 6, 4)
        self.numOffsetLeft = NumericUpDown()
        self.numOffsetLeft.Minimum = Decimal(0)
        self.numOffsetLeft.Maximum = Decimal(35000)
        self.numOffsetLeft.DecimalPlaces = 0
        self.numOffsetLeft.Value = Decimal(0)
        self.numOffsetLeft.Width = 80
        self.numOffsetLeft.Margin = Padding(0, 4, 4, 4)

        baseFlow.Controls.Add(self.rbGrids)
        baseFlow.Controls.Add(self.rbCrop)
        baseFlow.Controls.Add(self.chkSameChain)
        baseFlow.Controls.Add(lblTop)
        baseFlow.Controls.Add(self.numOffsetTop)
        baseFlow.Controls.Add(lblLeft)
        baseFlow.Controls.Add(self.numOffsetLeft)
        grpBase.Controls.Add(baseFlow)

        settings.Controls.Add(grpBase, 0, settings.RowCount)
        settings.SetColumnSpan(grpBase, 2)
        settings.RowCount += 1

        # Row: Overall group
        grpOverall = GroupBox()
        grpOverall.Text = "Overall Dimension (extra line farther out)"
        grpOverall.AutoSize = True
        grpOverall.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink
        grpOverall.Dock = DockStyle.Top

        overallFlow = FlowLayoutPanel()
        overallFlow.Dock = DockStyle.Top
        overallFlow.AutoSize = True
        overallFlow.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink
        overallFlow.WrapContents = False
        overallFlow.Padding = Padding(8, 6, 8, 6)
        overallFlow.Margin = Padding(8, 4, 8, 6)
        overallFlow.FlowDirection = WinForms.FlowDirection.LeftToRight

        self.chkOverall = CheckBox()
        self.chkOverall.Text = "Add Overall Dimension (Top & Left)"
        self.chkOverall.AutoSize = True
        self.chkOverall.Margin = Padding(4, 4, 16, 4)

        self.chkSameOverall = CheckBox()
        self.chkSameOverall.Text = "Same extra offset for Top & Left"
        self.chkSameOverall.AutoSize = True
        self.chkSameOverall.Checked = True
        self.chkSameOverall.Margin = Padding(0, 4, 24, 4)

        lblOverallTop = Label(); lblOverallTop.Text = "Top Extra Offset (mm):"; lblOverallTop.AutoSize = True; lblOverallTop.Margin = Padding(0, 7, 6, 4)
        self.numOverallTop = NumericUpDown()
        self.numOverallTop.Minimum = Decimal(0)
        self.numOverallTop.Maximum = Decimal(35000)
        self.numOverallTop.DecimalPlaces = 0
        self.numOverallTop.Value = Decimal(0)
        self.numOverallTop.Width = 80
        self.numOverallTop.Margin = Padding(0, 4, 16, 4)

        lblOverallLeft = Label(); lblOverallLeft.Text = "Left Extra Offset (mm):"; lblOverallLeft.AutoSize = True; lblOverallLeft.Margin = Padding(0, 7, 6, 4)
        self.numOverallLeft = NumericUpDown()
        self.numOverallLeft.Minimum = Decimal(0)
        self.numOverallLeft.Maximum = Decimal(35000)
        self.numOverallLeft.DecimalPlaces = 0
        self.numOverallLeft.Value = Decimal(0)
        self.numOverallLeft.Width = 80
        self.numOverallLeft.Margin = Padding(0, 4, 4, 4)

        overallFlow.Controls.Add(self.chkOverall)
        overallFlow.Controls.Add(self.chkSameOverall)
        overallFlow.Controls.Add(lblOverallTop)
        overallFlow.Controls.Add(self.numOverallTop)
        overallFlow.Controls.Add(lblOverallLeft)
        overallFlow.Controls.Add(self.numOverallLeft)
        grpOverall.Controls.Add(overallFlow)

        settings.Controls.Add(grpOverall, 0, settings.RowCount)
        settings.SetColumnSpan(grpOverall, 2)
        settings.RowCount += 1

        # Row: Select & Save created dimensions as a Selection Set
        self.chkSelectSave = CheckBox()
        self.chkSelectSave.Text = "Select all created dimensions and create a Selection Set (named after first selected view)"
        self.chkSelectSave.AutoSize = True
        settings.Controls.Add(self.chkSelectSave, 0, settings.RowCount)
        settings.SetColumnSpan(self.chkSelectSave, 2)
        settings.RowCount += 1

        # --- NEW Row: Isolate created dims to their owner view (this run)
        self.chkIsolateOwnerDims = CheckBox()
        self.chkIsolateOwnerDims.Text = "Keep created dimensions only in their owner view (hide from other selected views)"
        self.chkIsolateOwnerDims.AutoSize = True
        self.chkIsolateOwnerDims.Checked = True  # default ON per your requirement
        settings.Controls.Add(self.chkIsolateOwnerDims, 0, settings.RowCount)
        settings.SetColumnSpan(self.chkIsolateOwnerDims, 2)
        settings.RowCount += 1

        # Bottom buttons
        panelButtons = Panel()
        panelButtons.Dock = DockStyle.Bottom
        panelButtons.Height = 52

        flow = FlowLayoutPanel()
        flow.Parent = panelButtons
        flow.Dock = DockStyle.Right
        flow.FlowDirection = WinForms.FlowDirection.LeftToRight
        flow.WrapContents = False
        flow.Padding = Padding(0, 10, 8, 10)
        flow.AutoSize = True
        flow.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink

        self.btnOK = Button(); self.btnOK.Text = "OK"; self.btnOK.Width = 96; self.btnOK.DialogResult = DialogResult.OK
        self.btnOK.Click += self._on_ok
        self.btnCancel = Button(); self.btnCancel.Text = "Cancel"; self.btnCancel.Width = 96; self.btnCancel.DialogResult = DialogResult.Cancel

        flow.Controls.Add(self.btnOK); flow.Controls.Add(self.btnCancel)
        root.Controls.Add(panelButtons, 0, 2)

        # Default buttons
        self.AcceptButton = self.btnOK
        self.CancelButton = self.btnCancel

        # Populate list (default: none selected)
        self._populate(self._filtered_items, default_checked=False)

        # Outputs (defaults)
        self.SelectedViewIds = []
        self.SelectedDimTypeId = None
        self.BaselineMode = 'grids'
        self.OffsetTopMm = 120.0
        self.OffsetLeftMm = 120.0
        self.AddOverall = False
        self.OverallTopMm = 120.0
        self.OverallLeftMm = 120.0
        self.SelectAndSave = False
        self.IsolateOwnerDims = True  # <-- NEW output flag

        # Attach settings panel
        panelSettings.Controls.Add(settings)

        # Events
        self.chkSameChain.CheckedChanged += self._toggle_chain_same
        self.numOffsetTop.ValueChanged += self._bind_chain_top
        self.chkOverall.CheckedChanged += self._toggle_overall_enabled
        self.chkSameOverall.CheckedChanged += self._toggle_overall_same
        self.numOverallTop.ValueChanged += self._bind_overall_top

        # Initialize interactive states
        self._toggle_chain_same(None, None)
        self._toggle_overall_enabled(None, None)
        self._toggle_overall_same(None, None)

    # ----- List interaction helpers -----
    def _capture_checked_ids(self):
        ids = set()
        for i in range(self.chkViews.Items.Count):
            if self.chkViews.GetItemChecked(i):
                disp = self.chkViews.Items[i]
                ids.add(self.view_map[disp].IntegerValue)
        self._checked_ids = ids

    def _populate(self, items, default_checked=False):
        self.chkViews.BeginUpdate()
        self.chkViews.Items.Clear()
        for disp in items:
            idx = self.chkViews.Items.Add(disp)
            vid = self.view_map[disp].IntegerValue
            self.chkViews.SetItemChecked(idx, (vid in self._checked_ids))
        self.chkViews.EndUpdate()

    def _on_search(self, sender, args):
        self._capture_checked_ids()
        q = (self.txtSearch.Text or "").strip().lower()
        if not q:
            self._filtered_items = list(self._all_items)
        else:
            self._filtered_items = [s for s in self._all_items if q in s.lower()]
        self._populate(self._filtered_items, default_checked=False)

    # Select All / Clear current filtered list
    def _on_all(self, sender, args):
        self.chkViews.BeginUpdate()
        for i in range(self.chkViews.Items.Count):
            self.chkViews.SetItemCheckState(i, CheckState.Checked)
        self.chkViews.EndUpdate()
        self._capture_checked_ids()

    def _on_none(self, sender, args):
        self.chkViews.BeginUpdate()
        for i in range(self.chkViews.Items.Count):
            self.chkViews.SetItemCheckState(i, CheckState.Unchecked)
        self.chkViews.EndUpdate()
        self._capture_checked_ids()

    # SHIFT-range selection support
    def _on_list_mouse_down(self, sender, e):
        try:
            idx = self.chkViews.IndexFromPoint(Point(e.X, e.Y))
        except:
            idx = -1
        self._mouseDownIndex = idx

    def _on_item_check(self, sender, e):
        if self._isBatchChecking:
            return
        shift = (WinForms.Control.ModifierKeys & Keys.Shift) == Keys.Shift
        if shift and self._lastClickedIndex >= 0:
            start = min(self._lastClickedIndex, e.Index if self._mouseDownIndex < 0 else self._mouseDownIndex)
            end = max(self._lastClickedIndex, e.Index if self._mouseDownIndex < 0 else self._mouseDownIndex)
            desired = e.NewValue
            try:
                self._isBatchChecking = True
                self.chkViews.BeginUpdate()
                for i in range(start, end + 1):
                    self.chkViews.SetItemCheckState(i, desired)
            finally:
                self.chkViews.EndUpdate()
                self._isBatchChecking = False
        self._capture_checked_ids()
        self._lastClickedIndex = e.Index

    def _on_ok(self, sender, args):
        self._capture_checked_ids()
        self.SelectedViewIds = []
        for disp in self._all_items:
            vid = self.view_map[disp]
            if vid.IntegerValue in self._checked_ids:
                self.SelectedViewIds.append(vid)

        sel = self.cmbDimType.SelectedItem
        if sel and sel != "<Use Current Default>":
            self.SelectedDimTypeId = self.dimtype_map.get(sel, None)
        else:
            self.SelectedDimTypeId = None

        self.BaselineMode = 'grids' if self.rbGrids.Checked else 'crop'

        # Read chain values
        self.OffsetTopMm = float(self.numOffsetTop.Value)
        self.OffsetLeftMm = float(self.numOffsetLeft.Value) if not bool(self.chkSameChain.Checked) else float(self.numOffsetTop.Value)

        # Read overall values
        self.AddOverall = bool(self.chkOverall.Checked)
        if self.AddOverall:
            self.OverallTopMm = float(self.numOverallTop.Value)
            self.OverallLeftMm = float(self.numOverallLeft.Value) if not bool(self.chkSameOverall.Checked) else float(self.numOverallTop.Value)

        self.SelectAndSave = bool(self.chkSelectSave.Checked)
        self.IsolateOwnerDims = bool(self.chkIsolateOwnerDims.Checked)

    # --- UI helpers for Same-for-both behaviour ---
    def _toggle_chain_same(self, sender, args):
        same = bool(self.chkSameChain.Checked)
        self.numOffsetLeft.Enabled = not same
        if same:
            # mirror current Top value into Left
            self.numOffsetLeft.Value = self.numOffsetTop.Value

    def _bind_chain_top(self, sender, args):
        # When same-for-both is on, keep Left = Top as user changes Top
        if bool(self.chkSameChain.Checked):
            self.numOffsetLeft.Value = self.numOffsetTop.Value

    def _toggle_overall_enabled(self, sender, args):
        on = bool(self.chkOverall.Checked)
        self.chkSameOverall.Enabled = on
        self.numOverallTop.Enabled = on
        self.numOverallLeft.Enabled = on and (not bool(self.chkSameOverall.Checked))

    def _toggle_overall_same(self, sender, args):
        same = bool(self.chkSameOverall.Checked)
        self.numOverallLeft.Enabled = bool(self.chkOverall.Checked) and (not same)
        if same:
            self.numOverallLeft.Value = self.numOverallTop.Value

    def _bind_overall_top(self, sender, args):
        if bool(self.chkSameOverall.Checked):
            self.numOverallLeft.Value = self.numOverallTop.Value


# -------------------------
# Entry
# -------------------------
def main():
    plan_views = collect_plan_views(doc)
    if not plan_views:
        MessageBox.Show("No eligible plan views found.", "Grid Dims")
        return

    # Collect only Linear dimension types
    lin_dimtypes = collect_linear_dimension_types(doc)

    form = ViewPickerForm(plan_views, lin_dimtypes)
    try:
        hwnd = uiapp.MainWindowHandle
        owner = JtWindowHandle(hwnd)
        result = form.ShowDialog(owner)
    except:
        result = form.ShowDialog()

    if result != DialogResult.OK or not form.SelectedViewIds:
        return

    dimtype_elem = None
    if form.SelectedDimTypeId:
        dimtype_elem = doc.GetElement(form.SelectedDimTypeId)

    # Resolve selected views (objects) once
    selected_views = []
    for vid in form.SelectedViewIds:
        v = doc.GetElement(vid)
        if isinstance(v, View) and not v.IsTemplate:
            selected_views.append(v)

    tg = TransactionGroup(doc, "Add grid dimensions (Chain + Overall)")
    tg.Start()
    errors = []
    all_dim_ids = []

    for v in selected_views:
        t = Transaction(doc, "Grid dims in {}".format(safe_view_name(v)))
        t.Start()
        try:
            created_here = add_grid_dimensions_in_view(
                doc, v,
                off_top_mm=form.OffsetTopMm,
                off_left_mm=form.OffsetLeftMm,
                baseline_mode=('grids' if form.BaselineMode == 'grids' else 'crop'),
                dimtype=dimtype_elem,
                add_overall=form.AddOverall,
                overall_top_mm=(form.OverallTopMm if form.AddOverall else 0.0),
                overall_left_mm=(form.OverallLeftMm if form.AddOverall else 0.0)
            )
            if created_here:
                all_dim_ids.extend(created_here)

                # NEW: immediately hide these new dims from all OTHER selected views
                if form.IsolateOwnerDims:
                    other_views = [ov for ov in selected_views if ov.Id != v.Id]
                    isolate_created_dims_to_owner_view(doc, v, other_views, created_here)

        except Exception as ex:
            errors.append(u"View '{}': {}".format(safe_view_name(v), ex))
        t.Commit()

    tg.Assimilate()

    # Create Selection Set and (try to) select in UI
    if form.SelectAndSave and all_dim_ids:
        first_view = selected_views[0] if selected_views else None
        base_name = safe_view_name(first_view) if first_view else "Grid Dims"
        name_to_use = base_name

        tsel = Transaction(doc, "Create selection set for created dimensions")
        tsel.Start()
        sfe = None
        attempt = 1
        while True:
            try:
                sfe = SelectionFilterElement.Create(doc, name_to_use)
                break
            except:
                attempt += 1
                name_to_use = u"{} ({})".format(base_name, attempt)

        # Use typed .NET Array[ElementId]
        id_arr = to_elementid_array(all_dim_ids)
        sfe.SetElementIds(id_arr)
        tsel.Commit()

        # Also select them in the UI (non-fatal if it fails)
        try:
            uidoc.Selection.SetElementIds(id_arr)
            if errors:
                MessageBox.Show(
                    "Created Selection Set '{}' with {} dimensions and selected them.\n\n"
                    "Completed with some issues:\n\n{}".format(name_to_use, id_arr.Length, "\n".join(errors)),
                    "Grid Dims"
                )
            else:
                MessageBox.Show(
                    "Created Selection Set '{}' with {} dimensions and selected them.".format(name_to_use, id_arr.Length),
                    "Grid Dims"
                )
        except Exception as sel_ex:
            if errors:
                MessageBox.Show(
                    "Created Selection Set '{}' with {} dimensions.\n"
                    "Could not select them in UI:\n{}\n\nCompleted with some issues:\n\n{}"
                    .format(name_to_use, id_arr.Length, sel_ex, "\n".join(errors)),
                    "Grid Dims"
                )
            else:
                MessageBox.Show(
                    "Created Selection Set '{}' with {} dimensions.\n"
                    "Could not select them in UI:\n{}".format(name_to_use, id_arr.Length, sel_ex),
                    "Grid Dims"
                )
        return

    # Final message
    if errors:
        MessageBox.Show("Completed with some issues:\n\n" + "\n".join(errors), "Grid Dims")
    else:
        MessageBox.Show("Chain + Overall dimensions added to selected views.", "Grid Dims")


if __name__ == "__main__":
    main()