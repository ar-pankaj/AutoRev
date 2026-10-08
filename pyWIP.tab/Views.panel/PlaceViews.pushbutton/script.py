# -*- coding: utf-8 -*-

__title__ = "List Views\nSheets"

__author__ = "Pankaj Prabhakar"

__doc__ = """Place multiple views on multiple sheets (IronPython + UI)
- Two-column pickers with clickable sort and live counters
- No prompts for margins/spacing/columns (safe defaults internally)
- Correct viewport title type handling via ElementType
- Patched output summary table (reliable in pyRevit output)"""

from __future__ import print_function

# pyRevit
from pyrevit import forms, revit, script

# Revit API
from Autodesk.Revit.DB import (
    FilteredElementCollector, View, ViewType, ViewSchedule, ViewSheet, Viewport,
    ScheduleSheetInstance, XYZ, ElementType, BuiltInParameter
)

# .NET WinForms for custom pickers
import clr
clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System.Drawing import Size, Point
from System.Windows.Forms import (
    Form, ListView, ColumnHeader, Button, Label, DialogResult,
    View as WFView, AnchorStyles, FormStartPosition, ListViewItem
)

# ---------------------------------------------------------------------------
# Globals & Defaults
# ---------------------------------------------------------------------------
doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

FT_PER_MM = 1.0 / 304.8
def mm_to_ft(mm): return mm * FT_PER_MM

# Internal layout defaults (no prompts shown to user)
DEFAULT_MARGIN_FT  = mm_to_ft(15.0)   # safe margin from sheet edges
DEFAULT_SPACING_FT = mm_to_ft(10.0)   # spacing between tiles
DEFAULT_COLS       = None             # auto square-ish grid
DEFAULT_BATCH_N    = 6                # for "Batch N views per sheet" mode

# ---------------------------------------------------------------------------
# Utility: Collectors & Core Helpers
# ---------------------------------------------------------------------------
def collect_candidate_views():
    """Return printable, non-template views + schedules (exclude templates/internal)."""
    all_views = list(FilteredElementCollector(doc).OfClass(View))
    good = []
    for v in all_views:
        try:
            if v.IsTemplate:
                continue
            if v.ViewType in (ViewType.ProjectBrowser, ViewType.Internal):
                continue
            if not v.CanBePrinted:
                continue
            if isinstance(v, ViewSheet):
                continue
            good.append(v)
        except:
            continue
    return good


def collect_sheets():
    """Collect all non-placeholder sheets."""
    sheets = list(FilteredElementCollector(doc).OfClass(ViewSheet))
    return [s for s in sheets if not s.IsPlaceholder]


def _etype_name(et):
    """Robust ElementType name getter for IronPython."""
    try:
        return et.Name
    except:
        p = et.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        return p.AsString() if p else str(et.Id.IntegerValue)


def get_viewport_types():
    """
    Collect viewport title 'types' as ElementTypes (language-independent filter):
    any ElementType exposing VIEWPORT_ATTR_SHOW_EXTENSION_LINE belongs to Viewport titles.
    """
    fec = (FilteredElementCollector(doc)
           .WhereElementIsElementType()
           .OfClass(ElementType))
    return [et for et in fec
            if et.get_Parameter(BuiltInParameter.VIEWPORT_ATTR_SHOW_EXTENSION_LINE) is not None]


def sheet_outline_ft(sheet):
    # Outline in UV, but we only need values; XYZ z=0 for placement
    ol = sheet.Outline
    return ol.Min.U, ol.Min.V, ol.Max.U, ol.Max.V


def grid_centers(min_u, min_v, max_u, max_v, n_items,
                 margin_ft=DEFAULT_MARGIN_FT,
                 spacing_ft=DEFAULT_SPACING_FT,
                 cols=DEFAULT_COLS):
    """Compute centers arranged in a grid within margins; auto columns if None."""
    left = min_u + margin_ft
    right = max_u - margin_ft
    bottom = min_v + margin_ft
    top = max_v - margin_ft

    avail_w = max(0.0001, right - left)
    avail_h = max(0.0001, top - bottom)

    # auto columns (square-ish)
    if cols is None or cols <= 0:
        import math
        cols = int(math.ceil(math.sqrt(n_items if n_items > 0 else 1)))
    rows = (n_items + cols - 1) // cols if cols > 0 else 0

    # cell sizes
    if cols > 1:
        cell_w = (avail_w - spacing_ft * (cols - 1)) / float(cols)
    else:
        cell_w = avail_w
    if rows > 1:
        cell_h = (avail_h - spacing_ft * (rows - 1)) / float(rows)
    else:
        cell_h = avail_h

    centers = []
    for i in range(n_items):
        r = i // cols
        c = i % cols
        x = left + c * (cell_w + spacing_ft) + cell_w / 2.0
        y = top  - r * (cell_h + spacing_ft) - cell_h / 2.0
        centers.append(XYZ(x, y, 0))
    return centers


def can_add_view_to_sheet(sheet, view):
    try:
        return Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id)
    except:
        return False


def has_schedule_on_sheet(sheet, vsched):
    try:
        insts = list(FilteredElementCollector(doc, sheet.Id).OfClass(ScheduleSheetInstance))
        for si in insts:
            if si.ScheduleId == vsched.Id:
                return True
        return False
    except:
        return False


def place_on_sheet(sheet, views, centers, vp_type=None, avoid_dupe_legend=True):
    """Place given views on a sheet at given XYZ centers. Return (placed, skipped) report."""
    placed, skipped = [], []

    # Track any view already on this sheet (for legend dedupe)
    existing_view_ids_on_sheet = set()
    try:
        for vpId in sheet.GetAllViewports():
            vp = doc.GetElement(vpId)
            if vp:
                existing_view_ids_on_sheet.add(vp.ViewId)
    except:
        pass

    for view, center in zip(views, centers):
        try:
            # Schedules
            if isinstance(view, ViewSchedule):
                if has_schedule_on_sheet(sheet, view):
                    skipped.append((view, "Already placed schedule on sheet"))
                    continue
                with revit.Transaction("Place Schedule on Sheet"):
                    ScheduleSheetInstance.Create(doc, sheet.Id, view.Id, center)
                placed.append((view, "Schedule"))
                continue

            # Legends: allow on many sheets but avoid duplicates on same sheet
            if view.ViewType == ViewType.Legend and avoid_dupe_legend:
                if view.Id in existing_view_ids_on_sheet:
                    skipped.append((view, "Legend already on sheet"))
                    continue

            if not can_add_view_to_sheet(sheet, view):
                skipped.append((view, "Cannot add view to sheet"))
                continue

            with revit.Transaction("Place View on Sheet"):
                vp = Viewport.Create(doc, sheet.Id, view.Id, center)
                if vp_type:
                    try:
                        vp.ChangeTypeId(vp_type.Id)
                    except:
                        pass
            placed.append((view, "Viewport"))

            if view.ViewType == ViewType.Legend:
                existing_view_ids_on_sheet.add(view.Id)

        except Exception as ex:
            skipped.append((view, "Error: {}".format(ex)))
    return placed, skipped

# ---------------------------------------------------------------------------
# Output table rendering (PATCH)
# ---------------------------------------------------------------------------
def _as_text(x):
    """Return a safe unicode string for output table cells (avoid .NET strings)."""
    try:
        return unicode(x)        # IronPython 2
    except:
        try:
            return str(x)
        except:
            return ""


def render_summary_table(report_rows):
    """
    Renders a nice table in the pyRevit output window.
    Expects report_rows as a list of 5-tuples:
      (Sheet No., Sheet Name, View Name, View Type, Result)
    """
    columns = ["Sheet No.", "Sheet Name", "View", "Type", "Result"]
    # Ensure every cell is a plain Python string
    table_data = [[_as_text(c) for c in row] for row in report_rows]

    # Preferred: pyRevit's HTML table
    try:
        output.print_table(table_data=table_data, columns=columns, title="Placement Summary")
        return
    except Exception:
        pass

    # Fallback: Markdown table (always works)
    lines = []
    lines.append("| " + " | ".join(columns) + " |")
    lines.append("| " + " | ".join(["---"] * len(columns)) + " |")
    for row in table_data:
        lines.append("| " + " | ".join(row) + " |")
    output.print_md("\n".join(lines))

# ---------------------------------------------------------------------------
# WinForms: Two-column pickers with clickable sort and live counters
# ---------------------------------------------------------------------------
class SheetPickerForm(Form):
    """Two-column (Sheet Number | Sheet Name) picker with clickable headers and live counter."""
    def __init__(self, sheets):
        Form.__init__(self)
        self.Text = "Select Sheets"
        self.StartPosition = FormStartPosition.CenterScreen
        self.MinimumSize = Size(700, 500)
        self.Size = Size(900, 600)
        self._sort_col = 0
        self._ascending = True
        self.SelectedSheets = []

        # Instruction label
        self.lbl = Label()
        self.lbl.Text = "Click column headers to sort. Use Ctrl / Shift for multi-select."
        self.lbl.AutoSize = True
        self.lbl.Location = Point(12, 12)
        self.Controls.Add(self.lbl)

        # Live counter label (top-right)
        self.counter_lbl = Label()
        self.counter_lbl.AutoSize = True
        self.counter_lbl.Location = Point(self.ClientSize.Width - 220, 12)
        self.Controls.Add(self.counter_lbl)

        # ListView
        self.lv = ListView()
        self.lv.View = WFView.Details
        self.lv.FullRowSelect = True
        self.lv.MultiSelect = True
        self.lv.HideSelection = False
        self.lv.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.lv.Location = Point(12, 36)
        self.lv.Size = Size(self.ClientSize.Width - 24, self.ClientSize.Height - 100)

        # Columns
        ch_no = ColumnHeader();  ch_no.Text = "Sheet Number"; ch_no.Width = 180
        ch_nm = ColumnHeader();  ch_nm.Text = "Sheet Name";   ch_nm.Width = 600
        self.lv.Columns.Add(ch_no); self.lv.Columns.Add(ch_nm)

        # Initial sort by Sheet Number then Name
        def key_sheet(s): return ((s.SheetNumber or "").lower(), (s.Name or "").lower())
        sheets_sorted = sorted(sheets, key=key_sheet)

        for s in sheets_sorted:
            item = ListViewItem(s.SheetNumber or "")
            item.SubItems.Add(s.Name or "")
            item.Tag = s
            self.lv.Items.Add(item)

        # Sorting on header click
        def on_col_click(sender, e):
            if e.Column == self._sort_col:
                self._ascending = not self._ascending
            else:
                self._sort_col, self._ascending = e.Column, True
            items = [self.lv.Items[i] for i in range(self.lv.Items.Count)]
            def key_for(it):
                t1 = (it.SubItems[self._sort_col].Text or "").lower()
                t2 = (it.SubItems[1 - self._sort_col].Text or "").lower()
                return (t1, t2)
            items_sorted = sorted(items, key=key_for, reverse=not self._ascending)
            self.lv.BeginUpdate(); self.lv.Items.Clear()
            for it in items_sorted: self.lv.Items.Add(it)
            self.lv.EndUpdate()
            update_counter()

        self.lv.ColumnClick += on_col_click

        # Live counter updates
        def update_counter(*args):
            self.counter_lbl.Text = "Selected: {} / {}".format(self.lv.SelectedItems.Count, self.lv.Items.Count)
        self.lv.SelectedIndexChanged += update_counter

        self.Controls.Add(self.lv)

        # Buttons
        self.btn_ok = Button();     self.btn_ok.Text = "OK"
        self.btn_cancel = Button(); self.btn_cancel.Text = "Cancel"
        for btn in (self.btn_ok, self.btn_cancel):
            btn.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            btn.Size = Size(100, 28)
        self.btn_ok.Location = Point(self.ClientSize.Width - 212, self.ClientSize.Height - 48)
        self.btn_cancel.Location = Point(self.ClientSize.Width - 106, self.ClientSize.Height - 48)
        self.btn_ok.Click += self._on_ok
        self.btn_cancel.Click += self._on_cancel
        self.Controls.Add(self.btn_ok); self.Controls.Add(self.btn_cancel)

        # Resize handler
        def on_resize(sender, args):
            self.lv.Size = Size(self.ClientSize.Width - 24, self.ClientSize.Height - 100)
            self.btn_ok.Location = Point(self.ClientSize.Width - 212, self.ClientSize.Height - 48)
            self.btn_cancel.Location = Point(self.ClientSize.Width - 106, self.ClientSize.Height - 48)
            self.counter_lbl.Location = Point(self.ClientSize.Width - 220, 12)
        self.Resize += on_resize

        # Initialize counter
        update_counter()

    def _on_ok(self, sender, args):
        self.SelectedSheets = [it.Tag for it in self.lv.SelectedItems]
        if not self.SelectedSheets:
            forms.alert("Select at least one sheet.", warn_icon=True)
            return
        self.DialogResult = DialogResult.OK
        self.Close()

    def _on_cancel(self, sender, args):
        self.SelectedSheets = []
        self.DialogResult = DialogResult.Cancel
        self.Close()


class ViewPickerForm(Form):
    """Two-column (View Type | View Name) picker with clickable headers and live counter."""
    def __init__(self, views):
        Form.__init__(self)
        self.Text = "Select Views"
        self.StartPosition = FormStartPosition.CenterScreen
        self.MinimumSize = Size(700, 500)
        self.Size = Size(900, 600)
        self._sort_col = 0
        self._ascending = True
        self.SelectedViews = []

        # Instruction label
        self.lbl = Label()
        self.lbl.Text = "Click column headers to sort. Use Ctrl / Shift for multi-select."
        self.lbl.AutoSize = True
        self.lbl.Location = Point(12, 12)
        self.Controls.Add(self.lbl)

        # Live counter label (top-right)
        self.counter_lbl = Label()
        self.counter_lbl.AutoSize = True
        self.counter_lbl.Location = Point(self.ClientSize.Width - 220, 12)
        self.Controls.Add(self.counter_lbl)

        # ListView
        self.lv = ListView()
        self.lv.View = WFView.Details
        self.lv.FullRowSelect = True
        self.lv.MultiSelect = True
        self.lv.HideSelection = False
        self.lv.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.lv.Location = Point(12, 36)
        self.lv.Size = Size(self.ClientSize.Width - 24, self.ClientSize.Height - 100)

        # Columns
        ch_type = ColumnHeader(); ch_type.Text = "View Type"; ch_type.Width = 250
        ch_name = ColumnHeader(); ch_name.Text = "View Name"; ch_name.Width = 550
        self.lv.Columns.Add(ch_type); self.lv.Columns.Add(ch_name)

        # Initial sort by type then name
        def vtype_str(v):
            try: return v.ViewType.ToString()
            except: return ""
        views_sorted = sorted(views, key=lambda v: (vtype_str(v).lower(), (v.Name or "").lower()))

        for v in views_sorted:
            vt = vtype_str(v)
            name = v.Name or ""
            item = ListViewItem(vt); item.SubItems.Add(name); item.Tag = v
            self.lv.Items.Add(item)

        # Sorting on header click
        def on_col_click(sender, e):
            if e.Column == self._sort_col:
                self._ascending = not self._ascending
            else:
                self._sort_col, self._ascending = e.Column, True
            items = [self.lv.Items[i] for i in range(self.lv.Items.Count)]
            def key_for(it):
                t1 = (it.SubItems[self._sort_col].Text or "").lower()
                t2 = (it.SubItems[1 - self._sort_col].Text or "").lower()
                return (t1, t2)
            items_sorted = sorted(items, key=key_for, reverse=not self._ascending)
            self.lv.BeginUpdate(); self.lv.Items.Clear()
            for it in items_sorted: self.lv.Items.Add(it)
            self.lv.EndUpdate()
            update_counter()

        self.lv.ColumnClick += on_col_click

        # Live counter updates
        def update_counter(*args):
            self.counter_lbl.Text = "Selected: {} / {}".format(self.lv.SelectedItems.Count, self.lv.Items.Count)
        self.lv.SelectedIndexChanged += update_counter

        self.Controls.Add(self.lv)

        # Buttons
        self.btn_ok = Button();     self.btn_ok.Text = "OK"
        self.btn_cancel = Button(); self.btn_cancel.Text = "Cancel"
        for btn in (self.btn_ok, self.btn_cancel):
            btn.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            btn.Size = Size(100, 28)
        self.btn_ok.Location = Point(self.ClientSize.Width - 212, self.ClientSize.Height - 48)
        self.btn_cancel.Location = Point(self.ClientSize.Width - 106, self.ClientSize.Height - 48)
        self.btn_ok.Click += self._on_ok
        self.btn_cancel.Click += self._on_cancel
        self.Controls.Add(self.btn_ok); self.Controls.Add(self.btn_cancel)

        # Resize handler
        def on_resize(sender, args):
            self.lv.Size = Size(self.ClientSize.Width - 24, self.ClientSize.Height - 100)
            self.btn_ok.Location = Point(self.ClientSize.Width - 212, self.ClientSize.Height - 48)
            self.btn_cancel.Location = Point(self.ClientSize.Width - 106, self.ClientSize.Height - 48)
            self.counter_lbl.Location = Point(self.ClientSize.Width - 220, 12)
        self.Resize += on_resize

        # Initialize counter
        update_counter()

    def _on_ok(self, sender, args):
        self.SelectedViews = [it.Tag for it in self.lv.SelectedItems]
        if not self.SelectedViews:
            forms.alert("Select at least one view.", warn_icon=True)
            return
        self.DialogResult = DialogResult.OK
        self.Close()

    def _on_cancel(self, sender, args):
        self.SelectedViews = []
        self.DialogResult = DialogResult.Cancel
        self.Close()

# ---------------------------------------------------------------------------
# UI Flow: pickers & placement options
# ---------------------------------------------------------------------------
def pick_sheets_and_views():
    # Sheets
    sheets = collect_sheets()
    if not sheets:
        forms.alert("No sheets found.", exitscript=True)

    sform = SheetPickerForm(sheets)
    if sform.ShowDialog() != DialogResult.OK or not sform.SelectedSheets:
        script.exit()
    sel_sheets = sform.SelectedSheets

    # Views
    views = collect_candidate_views()
    if not views:
        forms.alert("No candidate views found.", exitscript=True)

    vform = ViewPickerForm(views)
    if vform.ShowDialog() != DialogResult.OK or not vform.SelectedViews:
        script.exit()
    sel_views = vform.SelectedViews

    return sel_sheets, sel_views


def pick_mode():
    modes = [
        "One-to-one (ordered)",
        "Batch N views per sheet",
        "All-to-all (tile all selected views on every selected sheet)",
    ]
    res = forms.CommandSwitchWindow.show(modes, message="Placement mode")
    if not res:
        script.exit()
    return res


def pick_viewport_type():
    vptypes = get_viewport_types()
    if not vptypes:
        return None
    choices = { _etype_name(vpt): vpt for vpt in vptypes }
    name = forms.SelectFromList.show(sorted(choices.keys()),
                                     title="Select Viewport Type (optional)",
                                     multiselect=False)
    return choices[name] if name else None

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def run():
    sheets, views = pick_sheets_and_views()
    mode = pick_mode()
    vp_type = pick_viewport_type()
    avoid_dupe_legend = True

    report_rows = []

    if mode == "One-to-one (ordered)":
        if len(views) != len(sheets):
            forms.alert("Counts differ: {} views vs {} sheets.\nPlacing up to the smaller count."
                        .format(len(views), len(sheets)))
        count = min(len(views), len(sheets))
        for i in range(count):
            sheet = sheets[i]
            view = views[i]
            min_u, min_v, max_u, max_v = sheet_outline_ft(sheet)
            centers = grid_centers(min_u, min_v, max_u, max_v, 1,
                                   DEFAULT_MARGIN_FT, DEFAULT_SPACING_FT, cols=1)
            placed, skipped = place_on_sheet(sheet, [view], centers,
                                             vp_type=vp_type, avoid_dupe_legend=avoid_dupe_legend)
            for (v, kind) in placed:
                report_rows.append((sheet.SheetNumber, sheet.Name, v.Name, v.ViewType.ToString(),
                                    "Placed ({})".format(kind)))
            for (v, msg) in skipped:
                report_rows.append((sheet.SheetNumber, sheet.Name, v.Name, v.ViewType.ToString(),
                                    "Skipped: {}".format(msg)))

    elif mode == "Batch N views per sheet":
        v_idx = 0
        for sheet in sheets:
            if v_idx >= len(views):
                break
            batch = views[v_idx:v_idx + DEFAULT_BATCH_N]
            v_idx += DEFAULT_BATCH_N
            min_u, min_v, max_u, max_v = sheet_outline_ft(sheet)
            centers = grid_centers(min_u, min_v, max_u, max_v, len(batch),
                                   DEFAULT_MARGIN_FT, DEFAULT_SPACING_FT, cols=DEFAULT_COLS)
            placed, skipped = place_on_sheet(sheet, batch, centers,
                                             vp_type=vp_type, avoid_dupe_legend=avoid_dupe_legend)
            for (v, kind) in placed:
                report_rows.append((sheet.SheetNumber, sheet.Name, v.Name, v.ViewType.ToString(),
                                    "Placed ({})".format(kind)))
            for (v, msg) in skipped:
                report_rows.append((sheet.SheetNumber, sheet.Name, v.Name, v.ViewType.ToString(),
                                    "Skipped: {}".format(msg)))

    elif mode == "All-to-all (tile all selected views on every selected sheet)":
        for sheet in sheets:
            min_u, min_v, max_u, max_v = sheet_outline_ft(sheet)
            centers = grid_centers(min_u, min_v, max_u, max_v, len(views),
                                   DEFAULT_MARGIN_FT, DEFAULT_SPACING_FT, cols=DEFAULT_COLS)
            placed, skipped = place_on_sheet(sheet, views, centers,
                                             vp_type=vp_type, avoid_dupe_legend=avoid_dupe_legend)
            for (v, kind) in placed:
                report_rows.append((sheet.SheetNumber, sheet.Name, v.Name, v.ViewType.ToString(),
                                    "Placed ({})".format(kind)))
            for (v, msg) in skipped:
                report_rows.append((sheet.SheetNumber, sheet.Name, v.Name, v.ViewType.ToString(),
                                    "Skipped: {}".format(msg)))

    # Placement Summary (table) — patched
    if report_rows:
        render_summary_table(report_rows)     # <— uses robust columns + safe strings
    else:
        forms.alert("Nothing placed.")

if __name__ == "__main__":
    run()
