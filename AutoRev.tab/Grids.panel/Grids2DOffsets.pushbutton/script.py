# -*- coding: utf-8 -*-
from __future__ import print_function

"""
Set 2D Grid Extents & Bubbles — pyRevit (Revit 2024, IronPython 2.7)
- Per-Side Offsets: Top / Bottom / Left / Right (or one distance for all four sides)
- Per-Side Modes: Top / Bottom / Left / Right = outside | inside (or one mode for all four sides)
- Set grid bubbles (End 1 / End 2 / Both / None)
- Works on Floor Plans and Structural Plans
- WinForms Input UI + WPF Results Dialog (inline XAML via XamlReader)
Author: Pankaj Prabhakar (pyRevit adaptation)
"""

import clr, math, datetime

# ------------------------------ Revit API --------------------------------
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import (
    FilteredElementCollector, View, ViewPlan, ViewType, Grid, ElementId,
    Line, XYZ, BoundingBoxXYZ, DatumExtentType, BuiltInParameter, DatumEnds, Transaction
)

# ------------------------------ pyRevit -----------------------------------
from pyrevit import revit

# ------------------------------ WinForms (input) --------------------------
clr.AddReference("System")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
from System import String
from System.Drawing import Size, Point, SystemFonts
from System.Windows.Forms import (
    Application as WinFormsApp, Form, Label, Button, ComboBox, CheckBox, DialogResult,
    AnchorStyles, FormStartPosition, TextBox, Keys, ComboBoxStyle, Control,
    ListView, View as WinView, ColumnHeaderStyle, CheckState,
    MessageBox
)

# ------------------------------ WPF (results) -----------------------------
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
from System.Windows import Clipboard as WpfClipboard
from Microsoft.Win32 import SaveFileDialog

# Load XAML from string safely (avoid treating it as a path)
clr.AddReference('System.IO')
clr.AddReference('System.Xml')
from System.IO import StringReader
from System.Xml import XmlReader
from System.Windows.Markup import XamlReader

clr.AddReference('System.Data')
from System.Data import DataTable

# Enable WinForms themed rendering
try:
    WinFormsApp.EnableVisualStyles()
except:
    pass

# Revit document handles (pyRevit)
doc  = revit.doc
uidoc = revit.uidoc


# ============================== Helpers ==============================
def to_feet_mm(val_mm):
    """Convert mm to ft."""
    try:
        return float(val_mm) / 304.8
    except:
        return 0.0

def get_plan_views(document):
    """Return Floor Plan + Structural Plan views (non-templates), sorted by Name."""
    views = []
    for v in FilteredElementCollector(document).OfClass(View).ToElements():
        try:
            if v.IsTemplate:
                continue
            if v.ViewType in (ViewType.FloorPlan, ViewType.EngineeringPlan):
                views.append(v)
        except:
            pass
    views.sort(key=lambda x: x.Name)
    return views

def get_selected_grids(document, uidocument):
    sel_ids = uidocument.Selection.GetElementIds()
    grids = []
    if sel_ids:
        for eid in sel_ids:
            el = document.GetElement(eid)
            if isinstance(el, Grid):
                grids.append(el)
    return grids

def get_scope_or_crop_bbox(doc, view):
    """Return (BoundingBoxXYZ, 'scope'/'crop') if available."""
    try:
        sb_param = view.get_Parameter(BuiltInParameter.VIEWER_SCOPE_BOX)
        if sb_param:
            sb_id = sb_param.AsElementId()
            if sb_id and sb_id != ElementId.InvalidElementId:
                sb_el = doc.GetElement(sb_id)
                if sb_el:
                    sbbox = sb_el.get_BoundingBox(None)
                    if sbbox and isinstance(sbbox, BoundingBoxXYZ):
                        return sbbox, "scope"
    except:
        pass
    try:
        cb = view.CropBox
        if isinstance(cb, BoundingBoxXYZ):
            return cb, "crop"
    except:
        pass
    return None, None

def bbox_corners_2d(bbox):
    """Return (corners [p00,p10,p11,p01], center) honoring bbox.Transform.
       p00=(minX,minY), p10=(maxX,minY), p11=(maxX,maxY), p01=(minX,maxY) """
    T = bbox.Transform
    bb_min, bb_max = bbox.Min, bbox.Max
    z_mid = 0.5 * (bb_min.Z + bb_max.Z)
    p00_l = XYZ(bb_min.X, bb_min.Y, z_mid)
    p10_l = XYZ(bb_max.X, bb_min.Y, z_mid)
    p11_l = XYZ(bb_max.X, bb_max.Y, z_mid)
    p01_l = XYZ(bb_min.X, bb_max.Y, z_mid)
    c_l   = XYZ(0.5*(bb_min.X+bb_max.X), 0.5*(bb_min.Y+bb_max.Y), z_mid)
    return [T.OfPoint(p00_l), T.OfPoint(p10_l), T.OfPoint(p11_l), T.OfPoint(p01_l)], T.OfPoint(c_l)

def line_seg_intersection_2d(p, v, a, b, tol=1e-9):
    px, py = p.X, p.Y
    vx, vy = v.X, v.Y
    ax, ay = a.X, a.Y
    bx, by = b.X, b.Y
    sx, sy = (bx - ax), (by - ay)
    denom = (vx * sy - vy * sx)
    if abs(denom) < tol:
        return (False, None, None, None)
    dx, dy = (ax - px), (ay - py)
    t = (dx * sy - dy * sx) / denom
    u = (dx * vy - dy * vx) / denom
    if u < -tol or u > 1 + tol:
        return (False, None, None, None)
    ix = px + t * vx
    iy = py + t * vy
    ip = XYZ(ix, iy, a.Z)
    return (True, ip, t, u)

def grid_line_from_grid(grid):
    try:
        crv = grid.Curve
        if isinstance(crv, Line):
            return crv
    except:
        pass
    return None

def ensure_2d_in_view(grid, view):
    try:
        has2d = grid.IsCurveInView(DatumExtentType.ViewSpecific, view)
    except:
        has2d = False
    if not has2d:
        try:
            model_enum = grid.GetCurvesInView(DatumExtentType.Model, view)
            model_list = list(model_enum) if model_enum else []
            if model_list:
                grid.SetCurveInView(DatumExtentType.ViewSpecific, view, model_list[0])
                return
        except:
            pass
    gl = grid_line_from_grid(grid)
    if isinstance(gl, Line):
        grid.SetCurveInView(DatumExtentType.ViewSpecific, view, gl)

def project_point_onto_line_3d(point, line_origin, line_dir_unit):
    w = XYZ(point.X - line_origin.X, point.Y - line_origin.Y, point.Z - line_origin.Z)
    t = w.X * line_dir_unit.X + w.Y * line_dir_unit.Y + w.Z * line_dir_unit.Z
    return XYZ(line_origin.X + t * line_dir_unit.X,
               line_origin.Y + t * line_dir_unit.Y,
               line_origin.Z + t * line_dir_unit.Z), t

def dist2_xy(a, b):
    dx, dy = (a.X - b.X), (a.Y - b.Y)
    return dx*dx + dy*dy

def _get_basis_curve_for_orientation(grid, view):
    """Prefer current 2D curve in view > model curve in view > overall curve."""
    try:
        cur2d_enum = grid.GetCurvesInView(DatumExtentType.ViewSpecific, view)
        cur2d_list = list(cur2d_enum) if cur2d_enum else []
        if cur2d_list and isinstance(cur2d_list[0], Line):
            return cur2d_list[0]
    except:
        pass
    try:
        model_enum = grid.GetCurvesInView(DatumExtentType.Model, view)
        model_list = list(model_enum) if model_enum else []
        if model_list and isinstance(model_list[0], Line):
            return model_list[0]
    except:
        pass
    return grid_line_from_grid(grid)

def _map_endpoints_to_basis(p1, p2, basis_line):
    """Preserve End0/End1 identity by mapping new endpoints to nearest basis endpoints (XY)."""
    if not isinstance(basis_line, Line):
        return p1, p2
    b0 = basis_line.GetEndPoint(0)
    b1 = basis_line.GetEndPoint(1)
    s_keep = dist2_xy(p1, b0) + dist2_xy(p2, b1)
    s_swap = dist2_xy(p1, b1) + dist2_xy(p2, b0)
    if s_swap < s_keep:
        return p2, p1
    return p1, p2

def set_grid_2d_extents_in_view(doc, grid, view, rect_corners, rect_center,
                                 offsets_ft, modes_by_side,
                                 tol=1e-7, min_len_ft=1e-4):
    """
    offsets_ft: dict with keys 'top','bottom','left','right' -> float (feet)
    modes_by_side: dict with keys 'top','bottom','left','right' -> 'outside'|'inside'
    """
    gl = grid_line_from_grid(grid)
    if not isinstance(gl, Line):
        return (False, "Skipped (curved grid).")
    o3 = gl.GetEndPoint(0)
    d3 = gl.GetEndPoint(1) - gl.GetEndPoint(0)
    d3_len = math.sqrt(d3.X*d3.X + d3.Y*d3.Y + d3.Z*d3.Z)
    if d3_len < tol:
        return (False, "Skipped (degenerate grid line).")
    v3 = XYZ(d3.X/d3_len, d3.Y/d3_len, d3.Z/d3_len)
    v2 = XYZ(v3.X, v3.Y, 0.0)
    v2_len = math.sqrt(v2.X*v2.X + v2.Y*v2.Y)
    if v2_len < tol:
        return (False, "Skipped (grid unsuitable in plan).")
    v2 = XYZ(v2.X/v2_len, v2.Y/v2_len, 0.0)

    p00, p10, p11, p01 = rect_corners
    # Edge order: bottom, right, top, left (counter-clockwise)
    edges = [(p00, p10, 'bottom'), (p10, p11, 'right'),
             (p11, p01, 'top'),    (p01, p00, 'left')]

    origin2 = XYZ(o3.X, o3.Y, p00.Z)
    hits = []
    for (a, b, side) in edges:
        hit, ip, t_on_line, u = line_seg_intersection_2d(origin2, v2, a, b)
        if hit:
            hits.append((ip, t_on_line, side))

    # Deduplicate coincident hits (e.g., exact corner)
    unique = []
    for (pt, t, side) in hits:
        if not any((pt.DistanceTo(q[0]) < 1e-6) for q in unique):
            unique.append((pt, t, side))
    hits = unique
    if len(hits) < 2:
        return (False, "Skipped (grid does not intersect box twice).")

    hits.sort(key=lambda x: x[1])
    (pA_xy, _, sideA) = hits[0]
    (pB_xy, _, sideB) = hits[-1]

    pA_on, _ = project_point_onto_line_3d(pA_xy, o3, v3)
    pB_on, _ = project_point_onto_line_3d(pB_xy, o3, v3)

    # Per-end outside/inside according to the side it hits
    modeA = (modes_by_side.get(sideA, 'outside') or 'outside').strip().lower()
    modeB = (modes_by_side.get(sideB, 'outside') or 'outside').strip().lower()
    outsideA = (modeA == 'outside')
    outsideB = (modeB == 'outside')

    def choose_endpoint(p_on, offset_mag, is_outside):
        q_plus  = XYZ(p_on.X + offset_mag*v3.X, p_on.Y + offset_mag*v3.Y, p_on.Z + offset_mag*v3.Z)
        q_minus = XYZ(p_on.X - offset_mag*v3.X, p_on.Y - offset_mag*v3.Y, p_on.Z - offset_mag*v3.Z)
        d_plus, d_minus = dist2_xy(q_plus, rect_center), dist2_xy(q_minus, rect_center)
        return q_plus if (d_plus >= d_minus if is_outside else d_plus <= d_minus) else q_minus

    offA = float(offsets_ft.get(sideA, 0.0))
    offB = float(offsets_ft.get(sideB, 0.0))
    p1 = choose_endpoint(pA_on, offA, outsideA)
    p2 = choose_endpoint(pB_on, offB, outsideB)

    if p1.DistanceTo(p2) < min_len_ft:
        return (False, "Skipped (new 2D length too small).")

    ensure_2d_in_view(grid, view)
    basis = _get_basis_curve_for_orientation(grid, view)
    p1, p2 = _map_endpoints_to_basis(p1, p2, basis)

    try:
        new_line = Line.CreateBound(p1, p2)
        grid.SetCurveInView(DatumExtentType.ViewSpecific, view, new_line)
        return (True, "Updated.")
    except Exception as ex1:
        # Fallback with 2D basis if available
        try:
            cur2d_enum = grid.GetCurvesInView(DatumExtentType.ViewSpecific, view)
            cur2d_list = list(cur2d_enum) if cur2d_enum else []
            cur2d = cur2d_list[0] if cur2d_list else None
            if isinstance(cur2d, Line):
                o2 = cur2d.GetEndPoint(0)
                d2v = cur2d.GetEndPoint(1) - cur2d.GetEndPoint(0)
                d2l = math.sqrt(d2v.X*d2v.X + d2v.Y*d2v.Y + d2v.Z*d2v.Z)
                if d2l > tol:
                    v2d = XYZ(d2v.X/d2l, d2v.Y/d2l, d2v.Z/d2l)
                    pA_on2, _ = project_point_onto_line_3d(pA_xy, o2, v2d)
                    pB_on2, _ = project_point_onto_line_3d(pB_xy, o2, v2d)

                    def choose2(p_on, offset_mag, is_outside):
                        q_plus  = XYZ(p_on.X + offset_mag*v2d.X, p_on.Y + offset_mag*v2d.Y, p_on.Z + offset_mag*v2d.Z)
                        q_minus = XYZ(p_on.X - offset_mag*v2d.X, p_on.Y - offset_mag*v2d.Y, p_on.Z - offset_mag*v2d.Z)
                        d_plus  = dist2_xy(q_plus, rect_center)
                        d_minus = dist2_xy(q_minus, rect_center)
                        return q_plus if (d_plus >= d_minus if is_outside else d_plus <= d_minus) else q_minus

                    q1 = choose2(pA_on2, offA, outsideA)
                    q2 = choose2(pB_on2, offB, outsideB)
                    q1, q2 = _map_endpoints_to_basis(q1, q2, cur2d)
                    if q1.DistanceTo(q2) >= min_len_ft:
                        grid.SetCurveInView(DatumExtentType.ViewSpecific, view, Line.CreateBound(q1, q2))
                        return (True, "Updated (fallback 2D basis).")
        except Exception as ex2:
            return (False, "ERROR (fallback failed): {0}".format(ex2))
        return (False, "ERROR: {0}".format(ex1))

END_CHOICES = ["End 1 Only", "End 2 Only", "Both Ends", "None"]
END_MAP = {
    "End 1 Only": (True, False),
    "End 2 Only": (False, True),
    "Both Ends" : (True, True),
    "None"      : (False, False),
}

def apply_grid_end_choice(grid, view, choice):
    show0, show1 = END_MAP.get(choice, (True, True))
    try:
        (grid.ShowBubbleInView if show0 else grid.HideBubbleInView)(DatumEnds.End0, view)
    except:
        pass
    try:
        (grid.ShowBubbleInView if show1 else grid.HideBubbleInView)(DatumEnds.End1, view)
    except:
        pass


# ============================ WPF Results (inline XAML) ============================
def _parse_result_line(line):
    try:
        s = line
        v1 = s.find("View '")
        if v1 < 0:
            return (None, None, line)
        v2 = s.find("':", v1+6)
        view = s[v1+6:v2]
        g1 = s.find("Grid '", v2)
        if g1 < 0:
            return (view, None, s[v2+2:].strip())
        g2 = s.find("'", g1+6)
        grid = s[g1+6:g2]
        mpos = s.find("->", g2)
        msg = s[mpos+2:].strip() if mpos >= 0 else ""
        return (view, grid, msg)
    except:
        return (None, None, line)

def _fmt_offsets(p):
    same = p.get('same_all', True)
    if same:
        dmm = p.get('distance_mm', None) or p.get('distances_mm', {}).get('top', 0.0)
        dft = p.get('offset_ft', None)   or p.get('offsets_ft', {}).get('top', 0.0)
        return "Distance: {0} mm ({1:.6f} ft) [same on all sides]".format(dmm, dft)
    else:
        dmm = p.get('distances_mm', {})
        dft = p.get('offsets_ft', {})
        return ("Distances (mm / ft): "
                "Top {0} / {1:.6f}, "
                "Bottom {2} / {3:.6f}, "
                "Left {4} / {5:.6f}, "
                "Right {6} / {7:.6f}").format(
                    dmm.get('top',0), dft.get('top',0.0),
                    dmm.get('bottom',0), dft.get('bottom',0.0),
                    dmm.get('left',0), dft.get('left',0.0),
                    dmm.get('right',0), dft.get('right',0.0))

def _fmt_modes(p):
    same_mode = p.get('same_mode', True)
    if same_mode:
        m = p.get('mode', None) or p.get('modes', {}).get('top', 'outside')
        return str(m)
    else:
        m = p.get('modes', {})
        return "Top {0}, Bottom {1}, Left {2}, Right {3}".format(
            m.get('top','outside'), m.get('bottom','outside'),
            m.get('left','outside'), m.get('right','outside')
        )

RESULTS_XAML = r"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Results — Grid Extents + Bubbles (pyRevit)"
        Width="920" Height="720"
        WindowStartupLocation="CenterScreen"
        MinWidth="820" MinHeight="560"
        FontFamily="{x:Static SystemFonts.MessageFontFamily}">
  <DockPanel LastChildFill="True">
    <!-- Header (top) -->
    <StackPanel DockPanel.Dock="Top" Margin="12" Orientation="Vertical">
      <TextBlock Text="Run Summary" FontWeight="Bold" FontSize="14"/>
      <TextBlock x:Name="lblParams1" Margin="0,6,0,0"/>
      <TextBlock x:Name="lblParams2" Margin="0,2,0,0"/>
      <TextBlock x:Name="lblCounts"  Margin="0,8,0,6"/>
    </StackPanel>

    <!-- Bottom buttons -->
    <Grid DockPanel.Dock="Bottom" Margin="12,0,12,12">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="8"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="8"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <Button x:Name="btnClose" Grid.Column="5" Width="100" Height="28" Content="Close" />
      <Button x:Name="btnSave"  Grid.Column="3" Width="100" Height="28" Content="Save…"/>
      <Button x:Name="btnCopy"  Grid.Column="1" Width="120" Height="28" Content="Copy Summary"/>
    </Grid>

    <!-- Tabs (fill) -->
    <TabControl x:Name="tabs" Margin="12" >
      <TabItem Header="Modified">
        <ListView x:Name="lvMod">
          <ListView.View>
            <GridView>
              <GridViewColumn Header="View" Width="220" DisplayMemberBinding="{Binding Path=View}" />
              <GridViewColumn Header="Grid" Width="120" DisplayMemberBinding="{Binding Path=Grid}" />
              <GridViewColumn Header="Action" Width="520" DisplayMemberBinding="{Binding Path=Action}" />
            </GridView>
          </ListView.View>
        </ListView>
      </TabItem>

      <TabItem Header="Skipped">
        <ListView x:Name="lvSkip">
          <ListView.View>
            <GridView>
              <GridViewColumn Header="Line" Width="860" DisplayMemberBinding="{Binding Path=Line}" />
            </GridView>
          </ListView.View>
        </ListView>
      </TabItem>

      <TabItem Header="Errors">
        <ListView x:Name="lvErr">
          <ListView.View>
            <GridView>
              <GridViewColumn Header="Line" Width="860" DisplayMemberBinding="{Binding Path=Line}" />
            </GridView>
          </ListView.View>
        </ListView>
      </TabItem>
    </TabControl>
  </DockPanel>
</Window>
"""

class ResultsWindow(object):
    def __init__(self, summary):
        # Build Window from XAML string safely
        sr = StringReader(RESULTS_XAML)
        xr = XmlReader.Create(sr)
        self.win = XamlReader.Load(xr)

        # --- Header strings
        ts = summary.get("timestamp") or ""
        p  = summary.get("params", {}) or {}
        c  = summary.get("counts", {}) or {}

        params1 = _fmt_offsets(p)
        modes_str = _fmt_modes(p)
        params2 = "Modes: {0}   Ends: {1}   Selection only: {2}   Timestamp: {3}".format(
            modes_str, p.get('end_choice'), p.get('selection_only'), ts
        )
        counts  = "Views selected: {0}   Views processed: {1}   Grids modified: {2}   Skipped lines: {3}   Errors: {4}".format(
            c.get("views_selected"), c.get("views_processed"), c.get("grids_modified"),
            c.get("skipped"), c.get("errors")
        )

        self._find("lblParams1").Text = params1
        self._find("lblParams2").Text = params2
        self._find("lblCounts").Text  = counts

        # --- Build tables for tabs
        dt_mod = DataTable("Modified")
        dt_mod.Columns.Add("View",  String)
        dt_mod.Columns.Add("Grid",  String)
        dt_mod.Columns.Add("Action",String)
        for s in (summary.get("modified", []) or []):
            v, g, m = _parse_result_line(s)
            row = dt_mod.NewRow()
            row["View"] = v or ""
            row["Grid"] = g or ""
            row["Action"] = m or s
            dt_mod.Rows.Add(row)

        dt_skip = DataTable("Skipped")
        dt_skip.Columns.Add("Line", String)
        for s in (summary.get("skipped", []) or []):
            row = dt_skip.NewRow()
            row["Line"] = s
            dt_skip.Rows.Add(row)

        dt_err = DataTable("Errors")
        dt_err.Columns.Add("Line", String)
        for s in (summary.get("errors", []) or []):
            row = dt_err.NewRow()
            row["Line"] = s
            dt_err.Rows.Add(row)

        # Bind
        self._find("lvMod").ItemsSource  = dt_mod.DefaultView
        self._find("lvSkip").ItemsSource = dt_skip.DefaultView
        self._find("lvErr").ItemsSource  = dt_err.DefaultView

        # Full summary text (for Copy/Save)
        self._summary_text = self._build_summary_text(summary)

        # Wire buttons
        self._find("btnCopy").Click  += self.on_copy
        self._find("btnSave").Click  += self.on_save
        self._find("btnClose").Click += self.on_close

    def _find(self, name):
        return self.win.FindName(name)

    def _build_summary_text(self, summary):
        p = summary.get("params", {}) or {}
        c = summary.get("counts", {}) or {}
        lines = []
        lines.append("=== Grid Extents + Bubbles — Results ===")
        lines.append("Timestamp: {0}".format(summary.get("timestamp")))
        lines.append("")
        lines.append("Parameters:")
        lines.append(" " + _fmt_offsets(p))
        lines.append(" Modes: {0}".format(_fmt_modes(p)))
        lines.append(" Ends: {0}".format(p.get("end_choice")))
        lines.append(" Selection only: {0}".format(p.get("selection_only")))
        lines.append("")
        lines.append("Counts:")
        lines.append(" Views selected: {0}".format(c.get("views_selected")))
        lines.append(" Views processed: {0}".format(c.get("views_processed")))
        lines.append(" Grids modified: {0}".format(c.get("grids_modified")))
        lines.append(" Skipped lines: {0}".format(c.get("skipped")))
        lines.append(" Errors: {0}".format(c.get("errors")))
        lines.append("")
        lines.append("---- Modified ----")
        for s in (summary.get("modified", []) or []):
            lines.append(" " + s)
        lines.append("")
        lines.append("---- Skipped ----")
        for s in (summary.get("skipped", []) or []):
            lines.append(" " + s)
        lines.append("")
        lines.append("---- Errors ----")
        for s in (summary.get("errors", []) or []):
            lines.append(" " + s)
        return "\n".join(lines)

    def on_copy(self, sender, args):
        try:
            WpfClipboard.Clear()
            WpfClipboard.SetText(self._summary_text)
            MessageBox.Show("Summary copied to clipboard.")
        except Exception as ex:
            MessageBox.Show("Copy failed: {0}".format(ex))

    def on_save(self, sender, args):
        try:
            dlg = SaveFileDialog()
            dlg.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        # ShowDialog returns Nullable<bool>; in IronPython truthiness works
            if dlg.ShowDialog():
                with open(dlg.FileName, 'w') as f:
                    f.write(self._summary_text)
                MessageBox.Show("Saved: {0}".format(dlg.FileName))
        except Exception as ex:
            MessageBox.Show("Save failed: {0}".format(ex))

    def on_close(self, sender, args):
        self.win.Close()

    def ShowDialog(self):
        return self.win.ShowDialog()


# ============================ WinForms Input ============================
END_CHOICES = ["End 1 Only", "End 2 Only", "Both Ends", "None"]

class GridBubbleForm(Form):
    def __init__(self, document, uidocument):
        self.doc = document
        self.uidoc = uidocument
        self.views_all = get_plan_views(document)
        self.views_filtered = list(self.views_all)
        self._checks = {v.Id.IntegerValue: False for v in self.views_all}

        Form.__init__(self)
        self.Text = "Set 2D Grid Extents & Bubbles (pyRevit)"
        self.Width = 960
        self.Height = 820
        self.StartPosition = FormStartPosition.CenterScreen
        try:
            self.Font = SystemFonts.MessageBoxFont
        except:
            pass

        margin, spacing = 12, 8
        btn_h = 28

        # Filter row
        self.lblFilter = Label(Text="Search views:")
        self.lblFilter.AutoSize = True
        self.lblFilter.Location = Point(margin, margin)
        self.Controls.Add(self.lblFilter)

        self.txtFilter = TextBox()
        self.txtFilter.Location = Point(self.lblFilter.Right + spacing, self.lblFilter.Top - 2)
        self.txtFilter.Size = Size(260, 24)
        self.txtFilter.TextChanged += self.on_filter_changed
        self.Controls.Add(self.txtFilter)

        self.btnSelectAll = Button(Text="Select All")
        self.btnSelectAll.Size = Size(90, btn_h)
        self.btnSelectAll.Location = Point(self.txtFilter.Right + spacing, self.txtFilter.Top - 2)
        self.btnSelectAll.Anchor = AnchorStyles.Top | AnchorStyles.Right
        self.btnSelectAll.Click += self.on_select_all
        self.Controls.Add(self.btnSelectAll)

        self.btnSelectNone = Button(Text="Select None")
        self.btnSelectNone.Size = Size(100, btn_h)
        self.btnSelectNone.Location = Point(self.btnSelectAll.Right + spacing, self.txtFilter.Top - 2)
        self.btnSelectNone.Anchor = AnchorStyles.Top | AnchorStyles.Right
        self.btnSelectNone.Click += self.on_select_none
        self.Controls.Add(self.btnSelectNone)

        # Distance / Mode / Ends / Selection
        y2 = self.lblFilter.Bottom + spacing

        self.chkSameAll = CheckBox(Text="Use one distance for all 4 sides")
        self.chkSameAll.AutoSize = True
        self.chkSameAll.Checked = True
        self.chkSameAll.Location = Point(margin, y2)
        self.chkSameAll.CheckedChanged += self.on_sameall_changed
        self.Controls.Add(self.chkSameAll)

        self.lblDist = Label(Text="Distance (mm):")
        self.lblDist.AutoSize = True
        self.lblDist.Location = Point(self.chkSameAll.Right + spacing*2, y2)
        self.Controls.Add(self.lblDist)

        self.txtDist = TextBox(Text="300")
        self.txtDist.Location = Point(self.lblDist.Right + spacing, y2 - 2)
        self.txtDist.Size = Size(80, 24)
               # self.txtDist.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.Controls.Add(self.txtDist)

        # Global Mode + "same mode" toggle
        self.lblMode = Label(Text="Offset mode:")
        self.lblMode.AutoSize = True
        self.lblMode.Location = Point(self.txtDist.Right + 2*spacing, y2)
        self.Controls.Add(self.lblMode)

        self.cmbMode = ComboBox()
        self.cmbMode.DropDownStyle = ComboBoxStyle.DropDownList
        self.cmbMode.Items.Add("outside")
        self.cmbMode.Items.Add("inside")
        self.cmbMode.SelectedIndex = 0
        self.cmbMode.Location = Point(self.lblMode.Right + spacing, y2 - 2)
        self.cmbMode.Size = Size(110, 24)
        self.Controls.Add(self.cmbMode)

        self.chkSameMode = CheckBox(Text="Use one mode for all 4 sides")
        self.chkSameMode.AutoSize = True
        self.chkSameMode.Checked = True
        self.chkSameMode.Location = Point(self.cmbMode.Right + 2*spacing, y2 - 2)
        self.chkSameMode.CheckedChanged += self.on_samemode_changed
        self.Controls.Add(self.chkSameMode)

        self.lblEnd = Label(Text="Grid End choice:")
        self.lblEnd.AutoSize = True
        self.lblEnd.Location = Point(self.chkSameMode.Right + 2*spacing, y2)
        self.Controls.Add(self.lblEnd)

        self.cmbEnd = ComboBox()
        self.cmbEnd.DropDownStyle = ComboBoxStyle.DropDownList
        for s in END_CHOICES:
            self.cmbEnd.Items.Add(s)
        self.cmbEnd.SelectedIndex = 2   # Both Ends
        self.cmbEnd.Location = Point(self.lblEnd.Right + spacing, y2 - 2)
        self.cmbEnd.Size = Size(160, 24)
        self.Controls.Add(self.cmbEnd)

        self.chkSelectionOnly = CheckBox(Text="Use only selected Grids in Revit")
        self.chkSelectionOnly.AutoSize = True
        self.chkSelectionOnly.Location = Point(self.cmbEnd.Right + 2*spacing, y2 - 2)
        self.Controls.Add(self.chkSelectionOnly)

        # Per-side distances
        y3 = self.lblDist.Bottom + spacing*2
        self.lblPerSide = Label(Text="Per-side distances (mm):")
        self.lblPerSide.AutoSize = True
        self.lblPerSide.Location = Point(margin, y3)
        self.Controls.Add(self.lblPerSide)

        y4 = self.lblPerSide.Bottom + spacing
        self.lblTop = Label(Text="Top:");     self.lblTop.AutoSize = True;     self.lblTop.Location = Point(margin, y4); self.Controls.Add(self.lblTop)
        self.txtTop = TextBox(Text="300");    self.txtTop.Location = Point(self.lblTop.Right + spacing, y4 - 2); self.txtTop.Size = Size(70, 24); self.Controls.Add(self.txtTop)

        self.lblBottom = Label(Text="Bottom:"); self.lblBottom.AutoSize = True; self.lblBottom.Location = Point(self.txtTop.Right + spacing*4, y4); self.Controls.Add(self.lblBottom)
        self.txtBottom = TextBox(Text="300");   self.txtBottom.Location = Point(self.lblBottom.Right + spacing, y4 - 2); self.txtBottom.Size = Size(70, 24); self.Controls.Add(self.txtBottom)

        self.lblLeft = Label(Text="Left:");   self.lblLeft.AutoSize = True;   self.lblLeft.Location = Point(self.txtBottom.Right + spacing*4, y4); self.Controls.Add(self.lblLeft)
        self.txtLeft = TextBox(Text="300");   self.txtLeft.Location = Point(self.lblLeft.Right + spacing, y4 - 2); self.txtLeft.Size = Size(70, 24); self.Controls.Add(self.txtLeft)

        self.lblRight = Label(Text="Right:"); self.lblRight.AutoSize = True;  self.lblRight.Location = Point(self.txtLeft.Right + spacing*4, y4); self.Controls.Add(self.lblRight)
        self.txtRight = TextBox(Text="300");  self.txtRight.Location = Point(self.lblRight.Right + spacing, y4 - 2); self.txtRight.Size = Size(70, 24); self.Controls.Add(self.txtRight)

        # Per-side modes (below distances)
        y5 = self.txtRight.Bottom + spacing*2
        self.lblPerSideMode = Label(Text="Per-side offset modes:")
        self.lblPerSideMode.AutoSize = True
        self.lblPerSideMode.Location = Point(margin, y5)
        self.Controls.Add(self.lblPerSideMode)

        y6 = self.lblPerSideMode.Bottom + spacing

        # Helper to create a mode combo
        def _make_mode_combo(x, y):
            cb = ComboBox()
            cb.DropDownStyle = ComboBoxStyle.DropDownList
            cb.Items.Add("outside")
            cb.Items.Add("inside")
            cb.SelectedIndex = 0
            cb.Location = Point(x, y - 2)
            cb.Size = Size(90, 24)
            return cb

        self.lblTopM = Label(Text="Top:");     self.lblTopM.AutoSize = True;     self.lblTopM.Location = Point(margin, y6); self.Controls.Add(self.lblTopM)
        self.cmbTopMode = _make_mode_combo(self.lblTopM.Right + spacing, y6);     self.Controls.Add(self.cmbTopMode)

        self.lblBottomM = Label(Text="Bottom:"); self.lblBottomM.AutoSize = True; self.lblBottomM.Location = Point(self.cmbTopMode.Right + spacing*4, y6); self.Controls.Add(self.lblBottomM)
        self.cmbBottomMode = _make_mode_combo(self.lblBottomM.Right + spacing, y6); self.Controls.Add(self.cmbBottomMode)

        self.lblLeftM = Label(Text="Left:");   self.lblLeftM.AutoSize = True;    self.lblLeftM.Location = Point(self.cmbBottomMode.Right + spacing*4, y6); self.Controls.Add(self.lblLeftM)
        self.cmbLeftMode = _make_mode_combo(self.lblLeftM.Right + spacing, y6);   self.Controls.Add(self.cmbLeftMode)

        self.lblRightM = Label(Text="Right:"); self.lblRightM.AutoSize = True;   self.lblRightM.Location = Point(self.cmbLeftMode.Right + spacing*4, y6); self.Controls.Add(self.lblRightM)
        self.cmbRightMode = _make_mode_combo(self.lblRightM.Right + spacing, y6); self.Controls.Add(self.cmbRightMode)

        # Initialize per-side enable state
        self.on_sameall_changed(None, None)
        self.on_samemode_changed(None, None)

        # Views list
        self.lblViews = Label(Text="Select Floor/Structural Plan Views:")
        self.lblViews.AutoSize = True
        self.lblViews.Location = Point(margin, self.cmbRightMode.Bottom + spacing*2)
        self.Controls.Add(self.lblViews)

        self.lvViews = ListView()
        self.lvViews.View = WinView.Details
        self.lvViews.CheckBoxes = True
        self.lvViews.FullRowSelect = True
        self.lvViews.MultiSelect = True
        self.lvViews.HeaderStyle = ColumnHeaderStyle.Nonclickable
        self.lvViews.Location = Point(margin, self.lblViews.Bottom + spacing)
        self.lvViews.Size = Size(self.ClientSize.Width - 2*margin, 420)
        self.lvViews.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.lvViews.Columns.Add("View Name", self.lvViews.Width - 20)
        self.lvViews.ItemCheck += self.on_lv_itemcheck
        self.lvViews.KeyDown   += self.on_lv_keydown
        self.lvViews.MouseDown += self.on_lv_mousedown
        self.Controls.Add(self.lvViews)

        self._batchChecking = False
        self._last_index = None
        self._shiftDown = False
        self.populate_list(self.views_filtered)

        # Counter
        self.lblCount = Label(Text="Checked: 0 of 0 (visible: 0 of 0)")
        self.lblCount.AutoSize = True
        self.lblCount.Location = Point(margin, self.lvViews.Bottom + spacing)
        self.Controls.Add(self.lblCount)

        # OK/Cancel
        self.btnOK = Button(Text="OK")
        self.btnOK.Size = Size(100, btn_h)
        self.btnOK.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btnOK.Location = Point(self.ClientSize.Width - (100*2 + spacing + margin), self.ClientSize.Height - btn_h - margin)
        self.btnOK.Click += self.on_ok
        self.Controls.Add(self.btnOK)

        self.btnCancel = Button(Text="Cancel")
        self.btnCancel.Size = Size(100, btn_h)
        self.btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btnCancel.Location = Point(self.btnOK.Right + spacing, self.btnOK.Top)
        self.btnCancel.Click += self.on_cancel
        self.Controls.Add(self.btnCancel)

        self.update_count()

    # ---- List handling ----
    def populate_list(self, views):
        self.lvViews.BeginUpdate()
        try:
            self.lvViews.Items.Clear()
            for v in views:
                txt = v.Name or ""
                item = self.lvViews.Items.Add(txt)
                item.Tag = v
                try:
                    item.Checked = bool(self._checks.get(v.Id.IntegerValue, False))
                except:
                    item.Checked = False
        finally:
            self.lvViews.EndUpdate()

    def on_filter_changed(self, sender, e):
        txt_l = (self.txtFilter.Text or "").lower()
        self.views_filtered = [v for v in self.views_all if txt_l in (v.Name or "").lower()]
        self.populate_list(self.views_filtered)
        self.update_count()

    def on_select_all(self, sender, e):
        self._batchChecking = True
        try:
            for i in range(self.lvViews.Items.Count):
                it = self.lvViews.Items[i]
                if not it.Checked:
                    it.Checked = True
                v = it.Tag
                try:
                    self._checks[v.Id.IntegerValue] = True
                except:
                    pass
        finally:
            self._batchChecking = False
            self.update_count()

    def on_select_none(self, sender, e):
        self._batchChecking = True
        try:
            for i in range(self.lvViews.Items.Count):
                it = self.lvViews.Items[i]
                if it.Checked:
                    it.Checked = False
                v = it.Tag
                try:
                    self._checks[v.Id.IntegerValue] = False
                except:
                    pass
        finally:
            self._batchChecking = False
            self.update_count()

    def on_lv_mousedown(self, sender, e):
        try:
            self._shiftDown = ((Control.ModifierKeys & Keys.Shift) == Keys.Shift)
        except:
            self._shiftDown = False
        hit = sender.HitTest(e.X, e.Y)
        if hit.Item is not None:
            self._last_index = hit.Item.Index

    def on_lv_keydown(self, sender, e):
        if e.KeyCode == Keys.Space:
            sel = list(sender.SelectedIndices)
            if not sel and sender.FocusedItem is not None:
                sel = [sender.FocusedItem.Index]
            if sel:
                any_unchecked = any((not sender.Items[i].Checked) for i in sel)
                target = True if any_unchecked else False
                self._batchChecking = True
                try:
                    for i in sel:
                        it = sender.Items[i]
                        if it.Checked != target:
                            it.Checked = target
                        v = it.Tag
                        try:
                            self._checks[v.Id.IntegerValue] = target
                        except:
                            pass
                finally:
                    self._batchChecking = False
                    self.update_count()

    def on_lv_itemcheck(self, sender, e):
        if self._batchChecking:
            return
        try:
            it = sender.Items[e.Index]
            v = it.Tag
            self._checks[v.Id.IntegerValue] = (e.NewValue == CheckState.Checked)
            if self._shiftDown and self._last_index is not None and self._last_index != e.Index:
                lo, hi = (self._last_index, e.Index)
                if lo > hi:
                    lo, hi = hi, lo
                target = (e.NewValue == CheckState.Checked)
                self._batchChecking = True
                try:
                    for i in range(lo, hi+1):
                        if i == e.Index:
                            continue
                        item_i = sender.Items[i]
                        if item_i.Checked != target:
                            item_i.Checked = target
                        v_i = item_i.Tag
                        try:
                            self._checks[v_i.Id.IntegerValue] = target
                        except:
                            pass
                finally:
                    self._batchChecking = False
        finally:
            self.update_count()

    # ---- Toggles ----
    def on_sameall_changed(self, sender, e):
        same = self.chkSameAll.Checked
        for ctrl in [self.lblPerSide, self.lblTop, self.txtTop, self.lblBottom, self.txtBottom, self.lblLeft, self.txtLeft, self.lblRight, self.txtRight]:
            ctrl.Enabled = (not same)
        if same:
            try:
                v = (self.txtDist.Text or "").strip()
                for tb in [self.txtTop, self.txtBottom, self.txtLeft, self.txtRight]:
                    tb.Text = v
            except:
                pass

    def on_samemode_changed(self, sender, e):
        same = self.chkSameMode.Checked
        self.lblMode.Enabled = same
        self.cmbMode.Enabled = same
        for ctrl in [self.lblPerSideMode, self.lblTopM, self.cmbTopMode, self.lblBottomM, self.cmbBottomMode,
                     self.lblLeftM, self.cmbLeftMode, self.lblRightM, self.cmbRightMode]:
            ctrl.Enabled = (not same)
        if same:
            try:
                m = str(self.cmbMode.SelectedItem) or "outside"
                for cb in [self.cmbTopMode, self.cmbBottomMode, self.cmbLeftMode, self.cmbRightMode]:
                    cb.SelectedIndex = 0 if m == "outside" else 1
            except:
                pass

    # ---- Count, OK/Cancel, and properties ----
    def update_count(self):
        total_checked = 0
        for k in self._checks:
            try:
                if self._checks[k]:
                    total_checked += 1
            except:
                pass
        total_all = len(self._checks)
        visible_all = self.lvViews.Items.Count
        visible_checked = 0
        for i in range(visible_all):
            if self.lvViews.Items[i].Checked:
                visible_checked += 1
        self.lblCount.Text = "Checked: {0} of {1} (visible: {2} of {3})".format(total_checked, total_all, visible_checked, visible_all)

    def on_ok(self, sender, e):
        if self.same_all:
            try:
                val = float((self.txtDist.Text or "0").strip())
                if val <= 0:
                    raise ValueError
            except:
                MessageBox.Show("Please enter a positive numeric Distance in mm.")
                return
        else:
            try:
                vals = [float((tb.Text or "0").strip()) for tb in [self.txtTop, self.txtBottom, self.txtLeft, self.txtRight]]
                if any(v <= 0 for v in vals):
                    raise ValueError
            except:
                MessageBox.Show("Please enter positive numeric distances (mm) for Top/Bottom/Left/Right.")
                return

        if not self.selected_views:
            MessageBox.Show("Please check at least one Floor/Structural Plan view.")
            return
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_cancel(self, sender, e):
        self.DialogResult = DialogResult.Cancel
        self.Close()

    # ---- Accessors ----
    @property
    def selected_views(self):
        checked_ids = set(k for k,v in self._checks.items() if v)
        result = []
        for v in self.views_all:
            try:
                if v.Id.IntegerValue in checked_ids:
                    result.append(v)
            except:
                pass
        return result

    @property
    def end_choice(self):
        s = self.cmbEnd.SelectedItem
        return str(s) if s else "Both Ends"

    @property
    def selection_only(self):
        return bool(self.chkSelectionOnly.Checked)

    @property
    def same_all(self):
        return bool(self.chkSameAll.Checked)

    @property
    def same_mode(self):
        return bool(self.chkSameMode.Checked)

    @property
    def distances_mm(self):
        if self.same_all:
            try:
                d = float((self.txtDist.Text or "0").strip())
            except:
                d = 0.0
            return { 'top': d, 'bottom': d, 'left': d, 'right': d }
        else:
            def _f(tb):
                try:
                    return float((tb.Text or "0").strip())
                except:
                    return 0.0
            return {
                'top': _f(self.txtTop),
                'bottom': _f(self.txtBottom),
                'left': _f(self.txtLeft),
                'right': _f(self.txtRight),
            }

    @property
    def modes_choice(self):
        if self.same_mode:
            m = str(self.cmbMode.SelectedItem) if self.cmbMode.SelectedItem else "outside"
            return {'top': m, 'bottom': m, 'left': m, 'right': m}
        else:
            def _m(cb):
                try:
                    s = cb.SelectedItem
                    return str(s) if s else "outside"
                except:
                    return "outside"
            return {
                'top': _m(self.cmbTopMode),
                'bottom': _m(self.cmbBottomMode),
                'left': _m(self.cmbLeftMode),
                'right': _m(self.cmbRightMode),
            }


# ================================ MAIN ================================
results, skipped, errors = [], [], []

form = GridBubbleForm(doc, uidoc)
res = form.ShowDialog()
if res != DialogResult.OK:
    MessageBox.Show("User cancelled.")
else:
    views = form.selected_views
    end_choice = form.end_choice
    selection_only = form.selection_only
    distances_mm = form.distances_mm  # dict with top/bottom/left/right
    modes_in = form.modes_choice      # dict with top/bottom/left/right

    if not views or any((distances_mm[k] <= 0 for k in ['top','bottom','left','right'])):
        MessageBox.Show("No views selected or invalid distances.")
    else:
        offsets_ft = {k: to_feet_mm(distances_mm.get(k, 0.0)) for k in ['top','bottom','left','right']}
        views_processed = 0
        grids_modified = 0

        t = Transaction(doc, "Set 2D Grid Extents (pyRevit)")
        t.Start()
        try:
            selected_grids_global = get_selected_grids(doc, uidoc) if selection_only else None

            for v in views:
                try:
                    if v.ViewType not in (ViewType.FloorPlan, ViewType.EngineeringPlan):
                        skipped.append("View '{0}' is not a plan; skipped.".format(v.Name))
                        continue
                except:
                    skipped.append("View '{0}' is not a plan; skipped.".format(v.Name))
                    continue

                bbox, source = get_scope_or_crop_bbox(doc, v)
                if not bbox:
                    skipped.append("View '{0}' has no Scope/Crop Box; skipped.".format(v.Name))
                    continue

                views_processed += 1
                rect_corners, rect_center = bbox_corners_2d(bbox)

                if selection_only:
                    grids_in_view = []
                    visible_ids = set(e.Id.IntegerValue for e in FilteredElementCollector(doc, v.Id).OfClass(Grid))
                    if selected_grids_global:
                        for g in selected_grids_global:
                            try:
                                if g.Id.IntegerValue in visible_ids:
                                    grids_in_view.append(g)
                            except:
                                pass
                    else:
                        skipped.append("View '{0}': selection_only True but no grids selected; skipped.".format(v.Name))
                        continue
                else:
                    grids_in_view = list(FilteredElementCollector(doc, v.Id).OfClass(Grid))

                for g in grids_in_view:
                    try:
                        ok, msg = set_grid_2d_extents_in_view(doc, g, v, rect_corners, rect_center,
                                                              offsets_ft, modes_in)
                        (results if ok else skipped).append("View '{0}': Grid '{1}' -> {2}".format(v.Name, g.Name, msg))
                        if ok:
                            grids_modified += 1
                        apply_grid_end_choice(g, v, end_choice)
                        results.append("View '{0}': Grid '{1}' -> Ends: {2}".format(v.Name, g.Name, end_choice))
                    except Exception as exg:
                        errors.append("View '{0}': Grid '{1}' -> ERROR: {2}".format(v.Name, g.Name, exg))
        finally:
            t.Commit()

        # Regenerate & repaint
        try: doc.Regenerate()
        except: pass
        try: uidoc.RefreshActiveView()
        except: pass
        try: WinFormsApp.DoEvents()
        except: pass

        # Build summary for WPF results
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        same_all = (abs(distances_mm['top'] - distances_mm['bottom']) < 1e-9 and
                    abs(distances_mm['top'] - distances_mm['left'])   < 1e-9 and
                    abs(distances_mm['top'] - distances_mm['right'])  < 1e-9)
        same_mode = (modes_in['top'] == modes_in['bottom'] == modes_in['left'] == modes_in['right'])
        summary = {
            "timestamp": ts,
            "params": {
                "same_all": same_all,
                "distance_mm": distances_mm['top'] if same_all else None,
                "distances_mm": distances_mm,
                "offset_ft": offsets_ft['top'] if same_all else None,
                "offsets_ft": offsets_ft,
                "same_mode": same_mode,
                "mode": modes_in['top'] if same_mode else None,
                "modes": modes_in,
                "end_choice": end_choice,
                "selection_only": selection_only
            },
            "counts": {
                "views_selected": len(views),
                "views_processed": views_processed,
                "grids_modified": grids_modified,
                "skipped": len(skipped),
                "errors": len(errors)
            },
            "modified": results,
            "skipped": skipped,
            "errors": errors
        }

        # Show WPF Results dialog (from string XAML)
        ResultsWindow(summary).ShowDialog()
