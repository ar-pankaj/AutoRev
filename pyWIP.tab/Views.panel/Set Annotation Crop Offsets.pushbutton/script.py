# -*- coding: utf-8 -*-
__title__ = "Set Annotation\nCrop Offsets"
__author__  = "Pankaj Prabhakar"
__doc__     = "Set Annotation Crop Offsets - Live-filter views & shift-select multiple views; set Top/Bottom/Left/Right Annotation Crop Offsets (inputs in mm)."

# -----------------------------
# Imports (Revit & .NET)
# -----------------------------
import clr
from System import Decimal

clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import (
    FilteredElementCollector, View, BuiltInParameter, Transaction, StorageType
)

from System.Windows.Forms import (
    Form, Label, ListBox, Button, ComboBox, TextBox, NumericUpDown, CheckBox,
    SelectionMode, DialogResult, MessageBox, MessageBoxButtons, FormBorderStyle,
    FormStartPosition, ComboBoxStyle, AutoScaleMode, HorizontalAlignment
)

from System.Drawing import Size, Point

# pyRevit globals
uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document


# -----------------------------
# Unit helpers & constraints
# -----------------------------
INCH_TO_FT = 1.0 / 12.0
MM_TO_FT   = 1.0 / 304.8
MIN_FT     = (1.0 / 8.0) * INCH_TO_FT    # Revit API minimum: 1/8" (paper units)

def clamp_nonneg_min(ft_value):
    """Revit requires non-negative and >= 1/8\" (paper units)."""
    return ft_value if ft_value >= MIN_FT else MIN_FT

def mm_to_ft(mm_value):
    return float(mm_value) * MM_TO_FT

def ft_to_mm(ft_value):
    return float(ft_value) / MM_TO_FT


# -----------------------------
# Data collection
# -----------------------------
def eligible_views(document):
    """Collect non-template views that support annotation crops."""
    views = []
    for v in FilteredElementCollector(document).OfClass(View):
        try:
            if v.IsTemplate:
                continue
            mgr = v.GetCropRegionShapeManager()
            if mgr and mgr.CanHaveAnnotationCrop:
                views.append(v)
        except:
            continue
    views.sort(key=lambda x: (str(x.ViewType), x.Name))
    return views


# -----------------------------
# WinForms UI
# -----------------------------
class OffsetsForm(Form):
    def __init__(self, all_views):
        # Window
        self.Text = "Adjust Annotation Crop Offsets (Multiple Views)"
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MinimizeBox = False
        self.MaximizeBox = False
        self.StartPosition = FormStartPosition.CenterScreen
        self.AutoScaleMode = AutoScaleMode.Font
        self.ClientSize = Size(820, 640)

        # Data
        self.all_views = all_views[:]      # unfiltered source
        self.filtered_views = all_views[:] # current working list

        # ------------- Filter Bar (live) -------------
        y = 10
        self.lblFilter = Label(Text="Filter: View Type | Name contains")
        self.lblFilter.Location = Point(10, y)
        self.lblFilter.Size = Size(400, 20)
        self.Controls.Add(self.lblFilter)

        # View Type (defaults to FloorPlan if available)
        self.cmbType = ComboBox()
        self.cmbType.Location = Point(10, y + 24)
        self.cmbType.Size = Size(240, 26)
        self.cmbType.DropDownStyle = ComboBoxStyle.DropDownList

        types = sorted({str(v.ViewType) for v in self.all_views})
        self.cmbType.Items.Add("All")
        for t in types:
            self.cmbType.Items.Add(t)

        # Default to FloorPlan if present, else All
        default_type = "FloorPlan"
        if default_type in [self.cmbType.Items[i] for i in range(self.cmbType.Items.Count)]:
            self.cmbType.SelectedItem = default_type
        else:
            self.cmbType.SelectedIndex = 0

        self.cmbType.SelectedIndexChanged += self._on_filter_changed
        self.Controls.Add(self.cmbType)

        # Name contains (live)
        self.txtName = TextBox()
        self.txtName.Location = Point(260, y + 24)
        self.txtName.Size = Size(320, 26)
        self.txtName.TextChanged += self._on_filter_changed
        self.Controls.Add(self.txtName)

        # Count
        self.lblCount = Label(Text="")
        self.lblCount.Location = Point(10, y + 56)
        self.lblCount.Size = Size(350, 20)
        self.Controls.Add(self.lblCount)

        # ------------- Views List -------------
        self.lb = ListBox()
        self.lb.Location = Point(10, y + 78)
        self.lb.Size = Size(800, 380)
        self.lb.SelectionMode = SelectionMode.MultiExtended
        self.lb.IntegralHeight = False
        self.Controls.Add(self.lb)

        # Initial filter & populate
        self._filter_views()
        self._refresh_listbox()

        # ------------- Selection Utils -------------
        self.btnAll  = Button(Text="Select All")
        self.btnNone = Button(Text="Select None")
        self.btnRead = Button(Text="Read From First Selected")
        self.btnAll.Location  = Point(10,  y + 468)
        self.btnNone.Location = Point(110, y + 468)
        self.btnRead.Location = Point(220, y + 468)
        self.btnAll.Size  = Size(90, 28)
        self.btnNone.Size = Size(100, 28)
        self.btnRead.Size = Size(200, 28)
        self.btnAll.Click  += self._select_all
        self.btnNone.Click += self._select_none
        self.btnRead.Click += self._read_from_first
        self.Controls.Add(self.btnAll)
        self.Controls.Add(self.btnNone)
        self.Controls.Add(self.btnRead)

        # ------------- Offsets (mm, wide controls) -------------
        base_y = y + 506
        self.lblTop    = Label(Text="Top (mm):");    self.lblTop.Location    = Point(10, base_y)
        self.lblBottom = Label(Text="Bottom (mm):"); self.lblBottom.Location = Point(10, base_y + 38)
        self.lblLeft   = Label(Text="Left (mm):");   self.lblLeft.Location   = Point(290, base_y)
        self.lblRight  = Label(Text="Right (mm):");  self.lblRight.Location  = Point(290, base_y + 38)
        for l in (self.lblTop, self.lblBottom, self.lblLeft, self.lblRight):
            l.Size = Size(110, 24)
            self.Controls.Add(l)

        def make_num(x, y, default_val):
            n = NumericUpDown()
            n.Location = Point(x, y)
            n.Size = Size(140, 28)
            n.DecimalPlaces = 1
            n.Increment = Decimal(1)
            n.Minimum = Decimal(0)
            n.Maximum = Decimal(100000)
            n.Value   = Decimal(default_val)
            n.TextAlign = HorizontalAlignment.Right
            return n

        # Default 25.0 mm ≈ 1 inch
        self.numTop    = make_num(120, base_y,       25.0)
        self.numBottom = make_num(120, base_y + 36,  25.0)
        self.numLeft   = make_num(400, base_y,       25.0)
        self.numRight  = make_num(400, base_y + 36,  25.0)
        self.Controls.Add(self.numTop); self.Controls.Add(self.numBottom)
        self.Controls.Add(self.numLeft); self.Controls.Add(self.numRight)

        self.chkUniform = CheckBox(Text="Use same offset for all sides")
        self.chkUniform.Location = Point(600, base_y)
        self.chkUniform.Size = Size(210, 24)
        self.chkUniform.Checked = False
        self.chkUniform.CheckedChanged += self._toggle_uniform
        self.Controls.Add(self.chkUniform)

        # ------------- Apply / Cancel -------------
        self.btnApply  = Button(Text="Apply")
        self.btnCancel = Button(Text="Cancel")
        self.btnApply.Location  = Point(620, base_y + 70)
        self.btnCancel.Location = Point(715, base_y + 70)
        self.btnApply.Size  = Size(80, 28)
        self.btnCancel.Size = Size(80, 28)
        self.btnApply.Click  += self._apply_click
        self.btnCancel.Click += self._cancel_click
        self.Controls.Add(self.btnApply)
        self.Controls.Add(self.btnCancel)

    # ---------- Live filter handlers ----------
    def _on_filter_changed(self, sender, args):
        self._filter_views()
        self._refresh_listbox()

    def _filter_views(self):
        type_sel  = self.cmbType.SelectedItem.ToString() if self.cmbType.SelectedItem else "All"
        name_text = self.txtName.Text.strip().lower()

        out = []
        for v in self.all_views:
            if type_sel != "All" and str(v.ViewType) != type_sel:
                continue
            if name_text and (name_text not in v.Name.lower()):
                continue
            out.append(v)
        self.filtered_views = out

    def _refresh_listbox(self):
        self.lb.BeginUpdate()
        try:
            self.lb.Items.Clear()
            for v in self.filtered_views:
                self.lb.Items.Add("{} | {}".format(str(v.ViewType), v.Name))
            self.lblCount.Text = "Showing {} of {} eligible views".format(
                len(self.filtered_views), len(self.all_views)
            )
        finally:
            self.lb.EndUpdate()

    # ---------- Selection helpers ----------
    def _select_all(self, sender, args):
        self.lb.BeginUpdate()
        try:
            self.lb.ClearSelected()
            for i in range(self.lb.Items.Count):
                self.lb.SetSelected(i, True)
        finally:
            self.lb.EndUpdate()

    def _select_none(self, sender, args):
        self.lb.ClearSelected()

    def _read_from_first(self, sender, args):
        sel = self.get_selected_views()
        if not sel:
            MessageBox.Show("Select at least one view first.", "pyRevit")
            return
        v = sel[0]
        mgr = v.GetCropRegionShapeManager()
        if mgr is None or not mgr.CanHaveAnnotationCrop:
            MessageBox.Show("Selected view does not support Annotation Crop.", "pyRevit")
            return

        # Read offsets (paper-units feet) and display in mm
        t = ft_to_mm(mgr.TopAnnotationCropOffset)
        b = ft_to_mm(mgr.BottomAnnotationCropOffset)
        l = ft_to_mm(mgr.LeftAnnotationCropOffset)
        r = ft_to_mm(mgr.RightAnnotationCropOffset)

        self.numTop.Value    = Decimal(round(t, 1))
        if self.chkUniform.Checked:
            self.numBottom.Value = self.numTop.Value
            self.numLeft.Value   = self.numTop.Value
            self.numRight.Value  = self.numTop.Value
        else:
            self.numBottom.Value = Decimal(round(b, 1))
            self.numLeft.Value   = Decimal(round(l, 1))
            self.numRight.Value  = Decimal(round(r, 1))

    # ---------- Uniform toggle ----------
    def _toggle_uniform(self, sender, args):
        if self.chkUniform.Checked:
            t = self.numTop.Value
            self.numBottom.Enabled = False
            self.numLeft.Enabled   = False
            self.numRight.Enabled  = False
            self.numBottom.Value   = t
            self.numLeft.Value     = t
            self.numRight.Value    = t
        else:
            self.numBottom.Enabled = True
            self.numLeft.Enabled   = True
            self.numRight.Enabled  = True

    # ---------- Apply / Cancel ----------
    def _apply_click(self, sender, args):
        if self.lb.SelectedIndices.Count == 0:
            MessageBox.Show("Please select at least one view.", "No Selection", MessageBoxButtons.OK)
            return
        self.DialogResult = DialogResult.OK
        self.Close()

    def _cancel_click(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

    # ---------- Getters ----------
    def get_selected_views(self):
        idxs = [i for i in self.lb.SelectedIndices]
        return [self.filtered_views[i] for i in idxs]

    def get_offsets_feet(self):
        # Values are in mm; convert to feet and clamp to API minimum
        t = clamp_nonneg_min(mm_to_ft(self.numTop.Value))
        b = clamp_nonneg_min(mm_to_ft(self.numBottom.Value))
        l = clamp_nonneg_min(mm_to_ft(self.numLeft.Value))
        r = clamp_nonneg_min(mm_to_ft(self.numRight.Value))
        return (t, b, l, r)


# -----------------------------
# Execution
# -----------------------------
all_views = eligible_views(doc)
if not all_views:
    MessageBox.Show("No eligible views found that support Annotation Crop.", "pyRevit")
else:
    form = OffsetsForm(all_views)
    if form.ShowDialog() == DialogResult.OK:
        sel_views = form.get_selected_views()
        off_top, off_bot, off_left, off_right = form.get_offsets_feet()

        t = Transaction(doc, "Set Annotation Crop Offsets (multiple)")
        t.Start()
        try:
            for v in sel_views:
                # Ensure Crop is active
                if not v.CropBoxActive:
                    v.CropBoxActive = True

                # Ensure Annotation Crop is ON
                p = v.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE)
                if p and p.StorageType == StorageType.Integer and p.AsInteger() != 1:
                    p.Set(1)

                # Set offsets (paper-units feet)
                mgr = v.GetCropRegionShapeManager()
                if mgr is None or not mgr.CanHaveAnnotationCrop:
                    continue

                mgr.TopAnnotationCropOffset    = off_top
                mgr.BottomAnnotationCropOffset = off_bot
                mgr.LeftAnnotationCropOffset   = off_left
                mgr.RightAnnotationCropOffset  = off_right

            t.Commit()
            MessageBox.Show("Annotation crop offsets applied to {} view(s).".format(len(sel_views)), "pyRevit")
        except Exception as ex:
            t.RollBack()