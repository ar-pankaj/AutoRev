# -*- coding: utf-8 -*-

__author__  = "Pankaj Prabhakar"
__doc__     = """
_____________________________________________________________________
Set Grid Extents in Elevations and Sections
_____________________________________________________________________

Description:

Adjusts the view-specific 2D extents of grids in selected elevation
and section views using the nearest visible levels.

How it works:

- Select the required elevation and section views.
- Enter the top and bottom grid offsets.
- Choose inside or outside positioning for each grid end.
- Select which grid bubbles should be visible.
- The tool updates straight grids visible in the selected views.
- Modified, skipped, and failed items are shown in a results window.
_____________________________________________________________________
"""

from __future__ import print_function
import clr, math, datetime
# ------------------------------ Revit API ------------------------------
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import (
    FilteredElementCollector, View, ViewType, Grid, Level, ElementId,
    Line, XYZ, DatumExtentType, BuiltInParameter, DatumEnds, Transaction
)

# ------------------------------- pyRevit -------------------------------
from pyrevit import revit

# ------------------------- WinForms (input) ----------------------------
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

# -------------------------- WPF (results) ------------------------------
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
from System.Windows import Clipboard as WpfClipboard
from Microsoft.Win32 import SaveFileDialog
clr.AddReference('System.IO')
clr.AddReference('System.Xml')
from System.IO import StringReader
from System.Xml import XmlReader
from System.Windows.Markup import XamlReader
clr.AddReference('System.Data')
from System.Data import DataTable

try:
    WinFormsApp.EnableVisualStyles()
except:
    pass

doc  = revit.doc
uidoc = revit.uidoc

# ============================== Helpers ===============================

def to_feet_mm(val_mm):
    try:
        return float(val_mm) / 304.8
    except:
        return 0.0

def get_elev_section_views(document):
    """Return Elevation + Section views (non-templates), sorted by Name."""
    views = []
    for v in FilteredElementCollector(document).OfClass(View).ToElements():
        try:
            if v.IsTemplate:
                continue
            if v.ViewType in (ViewType.Elevation, ViewType.Section):
                views.append(v)
        except:
            pass
    views.sort(key=lambda x: x.Name)
    return views

# ---- Datum helpers (reused/adapted) ----

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
        try:
            gl = grid.Curve
            if isinstance(gl, Line):
                grid.SetCurveInView(DatumExtentType.ViewSpecific, view, gl)
        except:
            pass


def _get_basis_curve_for_view(grid, view):
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
    try:
        crv = grid.Curve
        if isinstance(crv, Line):
            return crv
    except:
        pass
    return None


def _map_endpoints_to_basis(p1, p2, basis_line):
    """Preserve End0/End1 identity: map new endpoints to nearest basis endpoints by Z then XY."""
    if not isinstance(basis_line, Line):
        return p1, p2
    b0 = basis_line.GetEndPoint(0)
    b1 = basis_line.GetEndPoint(1)
    s_keep = (abs(p1.Z-b0.Z) + abs(p2.Z-b1.Z), (p1.X-b0.X)**2 + (p1.Y-b0.Y)**2 + (p2.X-b1.X)**2 + (p2.Y-b1.Y)**2)
    s_swap = (abs(p1.Z-b1.Z) + abs(p2.Z-b0.Z), (p1.X-b1.X)**2 + (p1.Y-b1.Y)**2 + (p2.X-b0.X)**2 + (p2.Y-b0.Y)**2)
    if s_swap < s_keep:
        return p2, p1
    return p1, p2


def _point_on_line_at_z(line, z_target, tol=1e-12):
    """Return point on 'line' where it crosses the horizontal plane Z=z_target. None if nearly horizontal."""
    p0 = line.GetEndPoint(0); p1 = line.GetEndPoint(1)
    dz = (p1.Z - p0.Z)
    if abs(dz) < tol:
        return None
    t = (z_target - p0.Z) / dz
    return XYZ(p0.X + t*(p1.X - p0.X), p0.Y + t*(p1.Y - p0.Y), p0.Z + t*(p1.Z - p0.Z))


def _visible_levels_in_view(doc, view):
    """Return list of (level, elevation_ft) that are visible in this view."""
    lvls = []
    try:
        for lvl in FilteredElementCollector(doc, view.Id).OfClass(Level):
            try:
                lvls.append((lvl, float(lvl.Elevation)))
            except:
                pass
    except:
        pass
    # Ensure unique by Id and sort by elevation
    seen = set()
    uniq = []
    for l, z in lvls:
        if l.Id.IntegerValue in seen:
            continue
        seen.add(l.Id.IntegerValue)
        uniq.append((l, z))
    uniq.sort(key=lambda t: t[1])
    return uniq


# ---------- CORE: GRID-TO-LEVEL OFFSET LOGIC FOR ELEV/SECTION ----------

def set_grid_2d_extents_in_vertical_view(doc, grid, view,
                                         offsets_ft, modes_by_side,
                                         levels_sorted,
                                         tol=1e-7, min_len_ft=1e-4):
    """
    Set 2D extents by referencing nearest visible Level for Top and Bottom ends.

    offsets_ft: dict {'top','bottom'} -> float in feet
    modes_by_side: dict {'top','bottom'} -> 'outside'|'inside'
    levels_sorted: list of (Level, elevation_ft), ascending by elevation
    """
    gl = _get_basis_curve_for_view(grid, view)
    if not isinstance(gl, Line):
        return (False, "Skipped (curved or unsupported grid).")

    p0 = gl.GetEndPoint(0); p1 = gl.GetEndPoint(1)
    # Identify top/bottom indices by Z
    top_idx  = 0 if p0.Z >= p1.Z else 1
    bot_idx  = 1 - top_idx

    def _nearest_level_z(z):
        best = None
        best_d = 1e99
        for (_lvl, elev) in levels_sorted:
            d = abs(elev - z)
            if d < best_d:
                best_d = d
                best = elev
        return best

    new_pts = {0: None, 1: None}

    def _solve_side(side_name):
        idx  = top_idx if side_name == 'top' else bot_idx
        ep   = gl.GetEndPoint(idx)
        mode = (modes_by_side.get(side_name, 'outside') or 'outside').strip().lower()
        off  = float(offsets_ft.get(side_name, 0.0))
        # Sign convention in Elev/Section along +Z
        if side_name == 'top':
            sgn = +1 if mode == 'outside' else -1
        else:  # bottom
            sgn = -1 if mode == 'outside' else +1

        ref_z = _nearest_level_z(ep.Z)
        if ref_z is None:
            return False
        z_target = ref_z + sgn * off
        pt = _point_on_line_at_z(gl, z_target)
        if pt is None:
            # Fallback: move along line direction by offset projected on Z
            d3 = p1 - p0
            dlen = math.sqrt(d3.X*d3.X + d3.Y*d3.Y + d3.Z*d3.Z)
            if dlen < tol:
                return False
            v3 = XYZ(d3.X/dlen, d3.Y/dlen, d3.Z/dlen)
            dz = sgn * off
            if abs(v3.Z) < tol:
                return False
            t = dz / v3.Z
            pt = XYZ(ep.X + t*v3.X, ep.Y + t*v3.Y, ep.Z + t*v3.Z)
        new_pts[idx] = pt
        return True

    solved_any = False
    solved_any = _solve_side('top') or solved_any
    solved_any = _solve_side('bottom') or solved_any

    if not solved_any:
        return (False, "Skipped (no valid endpoints solved).")

    if new_pts[0] is None: new_pts[0] = p0
    if new_pts[1] is None: new_pts[1] = p1

    if new_pts[0].DistanceTo(new_pts[1]) < min_len_ft:
        return (False, "Skipped (new 2D length too small).")

    # Ensure 2D container
    ensure_2d_in_view(grid, view)

    # Preserve End0/End1
    basis = _get_basis_curve_for_view(grid, view)
    q1, q2 = _map_endpoints_to_basis(new_pts[0], new_pts[1], basis)

    try:
        grid.SetCurveInView(DatumExtentType.ViewSpecific, view, Line.CreateBound(q1, q2))
        return (True, "Updated by nearest visible Level(s).")
    except Exception as ex:
        return (False, "ERROR: {0}".format(ex))


# Bubble controls: Top/Bottom
END_CHOICES = ["Top Only", "Bottom Only", "Both", "None"]
END_MAP = {
    "Top Only"   : (True,  False),
    "Bottom Only": (False, True),
    "Both"       : (True,  True),
    "None"       : (False, False),
}

def apply_grid_end_choice_top_bottom(grid, view, choice):
    show_top, show_bot = END_MAP.get(choice, (True, True))
    # Map to End0/End1 by Z
    try:
        ln = _get_basis_curve_for_view(grid, view)
        if not isinstance(ln, Line):
            return
        p0 = ln.GetEndPoint(0); p1 = ln.GetEndPoint(1)
        top_is_end0 = (p0.Z >= p1.Z)
        top_end = DatumEnds.End0 if top_is_end0 else DatumEnds.End1
        bot_end = DatumEnds.End1 if top_is_end0 else DatumEnds.End0
        (grid.ShowBubbleInView if show_top else grid.HideBubbleInView)(top_end, view)
        (grid.ShowBubbleInView if show_bot else grid.HideBubbleInView)(bot_end, view)
    except:
        pass


# ============================ WPF Results (inline XAML) ============================
RESULTS_XAML = r"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Results — Grid Extents (Elev/Section)"
        Width="900" Height="680"
        WindowStartupLocation="CenterScreen"
        MinWidth="800" MinHeight="520"
        FontFamily="{x:Static SystemFonts.MessageFontFamily}">
  <DockPanel LastChildFill="True">
    <StackPanel DockPanel.Dock="Top" Margin="12" Orientation="Vertical">
      <TextBlock Text="Run Summary" FontWeight="Bold" FontSize="14"/>
      <TextBlock x:Name="lblParams1" Margin="0,6,0,0"/>
      <TextBlock x:Name="lblParams2" Margin="0,2,0,0"/>
      <TextBlock x:Name="lblCounts"  Margin="0,8,0,6"/>
    </StackPanel>

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

    <TabControl x:Name="tabs" Margin="12" >
      <TabItem Header="Modified">
        <ListView x:Name="lvMod">
          <ListView.View>
            <GridView>
              <GridViewColumn Header="View" Width="220" DisplayMemberBinding="{Binding Path=View}" />
              <GridViewColumn Header="Grid" Width="120" DisplayMemberBinding="{Binding Path=Grid}" />
              <GridViewColumn Header="Action" Width="500" DisplayMemberBinding="{Binding Path=Action}" />
            </GridView>
          </ListView.View>
        </ListView>
      </TabItem>
      <TabItem Header="Skipped">
        <ListView x:Name="lvSkip">
          <ListView.View>
            <GridView>
              <GridViewColumn Header="Line" Width="840" DisplayMemberBinding="{Binding Path=Line}" />
            </GridView>
          </ListView.View>
        </ListView>
      </TabItem>
      <TabItem Header="Errors">
        <ListView x:Name="lvErr">
          <ListView.View>
            <GridView>
              <GridViewColumn Header="Line" Width="840" DisplayMemberBinding="{Binding Path=Line}" />
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
        sr = StringReader(RESULTS_XAML)
        xr = XmlReader.Create(sr)
        self.win = XamlReader.Load(xr)
        ts = summary.get("timestamp") or ""
        p  = summary.get("params", {}) or {}
        c  = summary.get("counts", {}) or {}
        params1 = "Distance Top/Bottom: {0} mm / {1} mm (ft: {2:.6f} / {3:.6f})".format(
            p.get('dist_mm',{}).get('top',0), p.get('dist_mm',{}).get('bottom',0),
            p.get('off_ft',{}).get('top',0.0), p.get('off_ft',{}).get('bottom',0.0))
        modes_str = "Modes: Top {0}, Bottom {1}".format(
            p.get('modes',{}).get('top','outside'), p.get('modes',{}).get('bottom','outside'))
        params2 = "{0} | Ends: {1} | Selection only: {2} | {3}".format(modes_str, p.get('end_choice'), p.get('selection_only'), ts)
        counts  = "Views selected: {0} | Views processed: {1} | Grids modified: {2} | Skipped: {3} | Errors: {4}".format(
            c.get('views_selected'), c.get('views_processed'), c.get('grids_modified'), c.get('skipped'), c.get('errors'))
        self._find("lblParams1").Text = params1
        self._find("lblParams2").Text = params2
        self._find("lblCounts").Text  = counts

        dt_mod = DataTable("Modified"); dt_mod.Columns.Add("View", String); dt_mod.Columns.Add("Grid", String); dt_mod.Columns.Add("Action", String)
        for s in (summary.get("modified", []) or []):
            # allow dicts or strings
            if isinstance(s, dict):
                row = dt_mod.NewRow(); row["View"] = s.get('view',''); row["Grid"] = s.get('grid',''); row["Action"] = s.get('msg',''); dt_mod.Rows.Add(row)
            else:
                row = dt_mod.NewRow(); row["View"] = ""; row["Grid"] = ""; row["Action"] = str(s); dt_mod.Rows.Add(row)
        dt_skip = DataTable("Skipped"); dt_skip.Columns.Add("Line", String)
        for s in (summary.get("skipped", []) or []):
            row = dt_skip.NewRow(); row["Line"] = s; dt_skip.Rows.Add(row)
        dt_err = DataTable("Errors"); dt_err.Columns.Add("Line", String)
        for s in (summary.get("errors", []) or []):
            row = dt_err.NewRow(); row["Line"] = s; dt_err.Rows.Add(row)
        self._find("lvMod").ItemsSource  = dt_mod.DefaultView
        self._find("lvSkip").ItemsSource = dt_skip.DefaultView
        self._find("lvErr").ItemsSource  = dt_err.DefaultView

        self._summary_text = self._build_summary_text(summary)
        self._find("btnCopy").Click += self.on_copy
        self._find("btnSave").Click += self.on_save
        self._find("btnClose").Click += self.on_close

    def _find(self, name):
        return self.win.FindName(name)

    def _build_summary_text(self, summary):
        p = summary.get("params", {}) or {}
        c = summary.get("counts", {}) or {}
        lines = []
        lines.append("=== Grid Extents (Elev/Section) — Results ===")
        lines.append("Timestamp: {0}".format(summary.get("timestamp")))
        lines.append("")
        lines.append("Parameters:")
        lines.append("  Distances (mm): Top {0}, Bottom {1}".format(p.get('dist_mm',{}).get('top',0), p.get('dist_mm',{}).get('bottom',0)))
        lines.append("  Offsets (ft): Top {0:.6f}, Bottom {1:.6f}".format(p.get('off_ft',{}).get('top',0.0), p.get('off_ft',{}).get('bottom',0.0)))
        lines.append("  Modes: Top {0}, Bottom {1}".format(p.get('modes',{}).get('top','outside'), p.get('modes',{}).get('bottom','outside')))
        lines.append("  Ends: {0}".format(p.get('end_choice')))
        lines.append("  Selection only: {0}".format(p.get('selection_only')))
        lines.append("")
        lines.append("Counts:")
        lines.append("  Views selected: {0}".format(c.get('views_selected')))
        lines.append("  Views processed: {0}".format(c.get('views_processed')))
        lines.append("  Grids modified: {0}".format(c.get('grids_modified')))
        lines.append("  Skipped: {0}".format(c.get('skipped')))
        lines.append("  Errors: {0}".format(c.get('errors')))
        lines.append("")
        lines.append("---- Modified ----")
        for s in (summary.get("modified", []) or []):
            if isinstance(s, dict):
                lines.append("  View '{0}': Grid '{1}' -> {2}".format(s.get('view',''), s.get('grid',''), s.get('msg','')))
            else:
                lines.append("  " + str(s))
        lines.append("")
        lines.append("---- Skipped ----")
        for s in (summary.get("skipped", []) or []):
            lines.append("  " + s)
        lines.append("")
        lines.append("---- Errors ----")
        for s in (summary.get("errors", []) or []):
            lines.append("  " + s)
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
END_CHOICES = ["Top Only", "Bottom Only", "Both", "None"]

class ElevSectionGridForm(Form):
    def __init__(self, document, uidocument):
        self.doc   = document
        self.uidoc = uidocument
        self.views_all      = get_elev_section_views(document)
        self.views_filtered = list(self.views_all)
        self._checks = {v.Id.IntegerValue: False for v in self.views_all}

        Form.__init__(self)
        self.Text = "Set 2D Grid Extents (Elev/Section) — by Nearest Level"
        self.Width = 920
        self.Height = 760
        self.StartPosition = FormStartPosition.CenterScreen
        try: self.Font = SystemFonts.MessageBoxFont
        except: pass

        margin, spacing, btn_h = 12, 8, 28

        # Filter row
        self.lblFilter = Label(Text="Search views:"); self.lblFilter.AutoSize = True; self.lblFilter.Location = Point(margin, margin); self.Controls.Add(self.lblFilter)
        self.txtFilter = TextBox(); self.txtFilter.Location = Point(self.lblFilter.Right + spacing, self.lblFilter.Top - 2); self.txtFilter.Size = Size(260, 24); self.txtFilter.TextChanged += self.on_filter_changed; self.Controls.Add(self.txtFilter)
        self.btnSelectAll  = Button(Text="Select All");  self.btnSelectAll.Size  = Size(90, btn_h); self.btnSelectAll.Location  = Point(self.txtFilter.Right + spacing, self.txtFilter.Top - 2); self.btnSelectAll.Click  += self.on_select_all;  self.Controls.Add(self.btnSelectAll)
        self.btnSelectNone = Button(Text="Select None"); self.btnSelectNone.Size = Size(100,btn_h); self.btnSelectNone.Location = Point(self.btnSelectAll.Right + spacing, self.txtFilter.Top - 2); self.btnSelectNone.Click += self.on_select_none; self.Controls.Add(self.btnSelectNone)

        # Distances and modes
        y2 = self.lblFilter.Bottom + spacing
        self.lblTop = Label(Text="Top Offset (mm):"); self.lblTop.AutoSize = True; self.lblTop.Location = Point(margin, y2); self.Controls.Add(self.lblTop)
        self.txtTop = TextBox(Text="300"); self.txtTop.Location = Point(self.lblTop.Right + spacing, y2 - 2); self.txtTop.Size = Size(80,24); self.Controls.Add(self.txtTop)
        self.lblTopM = Label(Text="Mode:"); self.lblTopM.AutoSize = True; self.lblTopM.Location = Point(self.txtTop.Right + spacing*2, y2); self.Controls.Add(self.lblTopM)
        self.cmbTopMode = ComboBox(); self.cmbTopMode.DropDownStyle = ComboBoxStyle.DropDownList; self.cmbTopMode.Items.Add("outside"); self.cmbTopMode.Items.Add("inside"); self.cmbTopMode.SelectedIndex = 0; self.cmbTopMode.Location = Point(self.lblTopM.Right + spacing, y2 - 2); self.cmbTopMode.Size = Size(90,24); self.Controls.Add(self.cmbTopMode)

        self.lblBot = Label(Text="Bottom Offset (mm):"); self.lblBot.AutoSize = True; self.lblBot.Location = Point(self.cmbTopMode.Right + spacing*4, y2); self.Controls.Add(self.lblBot)
        self.txtBot = TextBox(Text="300"); self.txtBot.Location = Point(self.lblBot.Right + spacing, y2 - 2); self.txtBot.Size = Size(80,24); self.Controls.Add(self.txtBot)
        self.lblBotM = Label(Text="Mode:"); self.lblBotM.AutoSize = True; self.lblBotM.Location = Point(self.txtBot.Right + spacing*2, y2); self.Controls.Add(self.lblBotM)
        self.cmbBotMode = ComboBox(); self.cmbBotMode.DropDownStyle = ComboBoxStyle.DropDownList; self.cmbBotMode.Items.Add("outside"); self.cmbBotMode.Items.Add("inside"); self.cmbBotMode.SelectedIndex = 0; self.cmbBotMode.Location = Point(self.lblBotM.Right + spacing, y2 - 2); self.cmbBotMode.Size = Size(90,24); self.Controls.Add(self.cmbBotMode)

        # Head (bubble) choice and selection toggle
        y3 = self.lblTop.Bottom + spacing*3
        self.lblEnd = Label(Text="Bubble Heads:"); self.lblEnd.AutoSize = True; self.lblEnd.Location = Point(margin, y3); self.Controls.Add(self.lblEnd)
        self.cmbEnd = ComboBox(); self.cmbEnd.DropDownStyle = ComboBoxStyle.DropDownList
        for s in END_CHOICES: self.cmbEnd.Items.Add(s)
        self.cmbEnd.SelectedIndex = 2  # Both
        self.cmbEnd.Location = Point(self.lblEnd.Right + spacing, y3 - 2); self.cmbEnd.Size = Size(160,24); self.Controls.Add(self.cmbEnd)

        self.chkSelectionOnly = CheckBox(Text="Use only selected Grids in Revit"); self.chkSelectionOnly.AutoSize = True; self.chkSelectionOnly.Location = Point(self.cmbEnd.Right + spacing*3, y3 - 2); self.Controls.Add(self.chkSelectionOnly)

        # Views list
        self.lblViews = Label(Text="Select Elevation / Section Views:"); self.lblViews.AutoSize = True; self.lblViews.Location = Point(margin, self.cmbEnd.Bottom + spacing*2); self.Controls.Add(self.lblViews)
        self.lvViews = ListView(); self.lvViews.View = WinView.Details; self.lvViews.CheckBoxes = True; self.lvViews.FullRowSelect = True; self.lvViews.MultiSelect = True; self.lvViews.HeaderStyle = ColumnHeaderStyle.Nonclickable
        self.lvViews.Location = Point(margin, self.lblViews.Bottom + spacing); self.lvViews.Size = Size(self.ClientSize.Width - 2*margin, 420)
        self.lvViews.Columns.Add("View Name", self.lvViews.Width - 20)
        self.lvViews.ItemCheck += self.on_lv_itemcheck
        self.lvViews.KeyDown   += self.on_lv_keydown
        self.lvViews.MouseDown += self.on_lv_mousedown
        self.Controls.Add(self.lvViews)

        self._batchChecking = False
        self._last_index    = None
        self._shiftDown     = False
        self.populate_list(self.views_filtered)

        # Counter
        self.lblCount = Label(Text="Checked: 0 of 0 (visible: 0 of 0)"); self.lblCount.AutoSize = True; self.lblCount.Location = Point(margin, self.lvViews.Bottom + spacing); self.Controls.Add(self.lblCount)

        # OK/Cancel
        self.btnOK = Button(Text="OK"); self.btnOK.Size = Size(100, btn_h); self.btnOK.Location = Point(self.ClientSize.Width - (100*2 + spacing + margin), self.ClientSize.Height - btn_h - margin); self.btnOK.Anchor = AnchorStyles.Bottom | AnchorStyles.Right; self.btnOK.Click += self.on_ok; self.Controls.Add(self.btnOK)
        self.btnCancel = Button(Text="Cancel"); self.btnCancel.Size = Size(100, btn_h); self.btnCancel.Location = Point(self.btnOK.Right + spacing, self.btnOK.Top); self.btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right; self.btnCancel.Click += self.on_cancel; self.Controls.Add(self.btnCancel)

        self.update_count()

    # ---- Confirmation handlers (added) ----
    def on_ok(self, sender, e):
        # Basic validation: positive numeric offsets and at least one view checked
        def _is_pos(x):
            try:
                return float((x or '0').strip()) > 0
            except:
                return False
        if not self.selected_views:
            MessageBox.Show('Please check at least one Elevation/Section view.')
            return
        if not (_is_pos(self.txtTop.Text) and _is_pos(self.txtBot.Text)):
            MessageBox.Show('Please enter positive numeric offsets (mm) for Top and Bottom.')
            return
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_cancel(self, sender, e):
        self.DialogResult = DialogResult.Cancel
        self.Close()

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
                try: self._checks[v.Id.IntegerValue] = True
                except: pass
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
                try: self._checks[v.Id.IntegerValue] = False
                except: pass
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
                        try: self._checks[v.Id.IntegerValue] = target
                        except: pass
                finally:
                    self._batchChecking = False
                    self.update_count()

    def on_lv_itemcheck(self, sender, e):
        if self._batchChecking:
            return
        try:
            it = sender.Items[e.Index]
            v  = it.Tag
            self._checks[v.Id.IntegerValue] = (e.NewValue == CheckState.Checked)
            if self._shiftDown and self._last_index is not None and self._last_index != e.Index:
                lo, hi = (self._last_index, e.Index)
                if lo > hi: lo, hi = hi, lo
                target = (e.NewValue == CheckState.Checked)
                self._batchChecking = True
                try:
                    for i in range(lo, hi+1):
                        if i == e.Index: continue
                        item_i = sender.Items[i]
                        if item_i.Checked != target:
                            item_i.Checked = target
                        v_i = item_i.Tag
                        try: self._checks[v_i.Id.IntegerValue] = target
                        except: pass
                finally:
                    self._batchChecking = False
        finally:
            self.update_count()

    def update_count(self):
        total_all = len(self._checks)
        total_checked = sum(1 for k,v in self._checks.items() if v)
        visible_all = self.lvViews.Items.Count
        visible_checked = sum(1 for i in range(visible_all) if self.lvViews.Items[i].Checked)
        self.lblCount.Text = "Checked: {0} of {1} (visible: {2} of {3})".format(total_checked, total_all, visible_checked, visible_all)

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
    def selection_only(self):
        return bool(self.chkSelectionOnly.Checked)

    @property
    def end_choice(self):
        s = self.cmbEnd.SelectedItem
        return str(s) if s else "Both"

    @property
    def distances_mm(self):
        def _f(tb):
            try: return float((tb.Text or "0").strip())
            except: return 0.0
        return {'top': _f(self.txtTop), 'bottom': _f(self.txtBot)}

    @property
    def modes_choice(self):
        def _m(cb):
            try:
                s = cb.SelectedItem
                return str(s) if s else "outside"
            except:
                return "outside"
        return {'top': _m(self.cmbTopMode), 'bottom': _m(self.cmbBotMode)}


# ================================ MAIN ================================
results, skipped, errors = [], [], []
form = ElevSectionGridForm(doc, uidoc)
res = form.ShowDialog()

if res != DialogResult.OK:
    MessageBox.Show("User cancelled.")
else:
    views          = form.selected_views
    end_choice     = form.end_choice
    selection_only = form.selection_only
    distances_mm   = form.distances_mm  # {'top','bottom'}
    modes_in       = form.modes_choice  # {'top','bottom'}

    if (not views) or any((distances_mm[k] <= 0 for k in ['top','bottom'])):
        MessageBox.Show("No views selected or invalid distances.")
    else:
        offsets_ft = {k: to_feet_mm(distances_mm.get(k, 0.0)) for k in ['top','bottom']}
        views_processed = 0
        grids_modified  = 0

        t = Transaction(doc, "Set 2D Grid Extents (Elev/Section) — by Level")
        t.Start()
        try:
            selected_grids_global = None
            if selection_only:
                # Only grids that are currently selected in Revit
                sel_ids = uidoc.Selection.GetElementIds()
                if sel_ids:
                    selected_grids_global = []
                    for eid in sel_ids:
                        el = doc.GetElement(eid)
                        if isinstance(el, Grid):
                            selected_grids_global.append(el)
                else:
                    skipped.append("Selection-only was True but no grids are selected in Revit.")

            for v in views:
                try:
                    if v.ViewType not in (ViewType.Elevation, ViewType.Section):
                        skipped.append("View '{0}' is not Elevation/Section; skipped.".format(v.Name))
                        continue
                except:
                    skipped.append("View '{0}' is not Elevation/Section; skipped.".format(v.Name))
                    continue

                levels_sorted = _visible_levels_in_view(doc, v)
                if not levels_sorted:
                    skipped.append("View '{0}': no visible Levels found; skipped.".format(v.Name))
                    continue

                views_processed += 1

                # Decide which grids to process in this view
                if selection_only and selected_grids_global is not None:
                    grids_in_view = []
                    visible_ids = set(e.Id.IntegerValue for e in FilteredElementCollector(doc, v.Id).OfClass(Grid))
                    for g in selected_grids_global:
                        try:
                            if g.Id.IntegerValue in visible_ids:
                                grids_in_view.append(g)
                        except:
                            pass
                else:
                    grids_in_view = list(FilteredElementCollector(doc, v.Id).OfClass(Grid))

                for g in grids_in_view:
                    try:
                        ok, msg = set_grid_2d_extents_in_vertical_view(
                            doc, g, v,
                            offsets_ft, modes_in,
                            levels_sorted
                        )
                        (results if ok else skipped).append({'view': v.Name, 'grid': g.Name, 'msg': msg})
                        if ok:
                            grids_modified += 1
                            # Apply bubble ends
                            apply_grid_end_choice_top_bottom(g, v, end_choice)
                            results.append({'view': v.Name, 'grid': g.Name, 'msg': "Heads: {0}".format(end_choice)})
                    except Exception as exg:
                        errors.append("View '{0}': Grid '{1}' -> ERROR: {2}".format(v.Name, g.Name, exg))
        finally:
            t.Commit()
            try: doc.Regenerate()
            except: pass
            try: uidoc.RefreshActiveView()
            except: pass
            try: WinFormsApp.DoEvents()
            except: pass

        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        summary = {
            'timestamp': ts,
            'params': {
                'dist_mm': distances_mm,
                'off_ft' : offsets_ft,
                'modes'  : modes_in,
                'end_choice': end_choice,
                'selection_only': selection_only,
            },
            'counts': {
                'views_selected': len(views),
                'views_processed': views_processed,
                'grids_modified' : grids_modified,
                'skipped': len(skipped),
                'errors' : len(errors),
            },
            'modified': results,
            'skipped' : skipped,
            'errors'  : errors,
        }
        ResultsWindow(summary).ShowDialog()