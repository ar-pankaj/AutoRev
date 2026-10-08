# -*- coding: utf-8 -*-
__title__ = "Model Mate"

import os
from Autodesk.Revit.DB import *
from pyrevit.forms import WPFWindow
from System.Collections.ObjectModel import ObservableCollection
from collections import defaultdict
from System.Windows.Controls import CheckBox, TextBlock
from System.Windows import Thickness, FontWeights

doc = __revit__.ActiveUIDocument.Document


# ---------------------------------------------------
# PATH
# ---------------------------------------------------
xamlfile = xamlfile = os.path.join(os.path.dirname(__file__), "ModelMateView.xaml")


# ===================================================
# MODELS
# ===================================================
class MaterialItem(object):
    def __init__(self, category, type_name, material, count, has_conflict, element_ids):
        self.Category = category
        self.TypeName = type_name
        self.Material = material
        self.Count = count
        self.HasConflict = has_conflict
        self.Select = False
        self.ElementIds = element_ids


class ReportItem(object):
    def __init__(self, name, vtype, status, template_name, elem):
        self.Select = False
        self.Name = name
        self.Type = vtype
        self.HasTemplate = status
        self.TemplateName = template_name
        self.ElementObj = elem


# ===================================================
# HELPERS
# ===================================================
def get_type_name(e):
    t = doc.GetElement(e.GetTypeId())
    if t:
        p = t.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            return p.AsString()
        return t.Name
    return "Unknown"


def get_material_name(e):

    # -------------------------
    # 1. INSTANCE
    # -------------------------
    p = e.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)

    if p:
        m = doc.GetElement(p.AsElementId())
        if m:
            return m.Name

    # -------------------------
    # 2. TYPE PARAMETER
    # -------------------------
    t = doc.GetElement(e.GetTypeId())

    if t:

        p_type = t.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)

        if p_type:
            m = doc.GetElement(p_type.AsElementId())
            if m:
                return m.Name

        # -------------------------
        # 3. COMPOUND STRUCTURE
        # -------------------------
        if hasattr(t, "GetCompoundStructure"):

            cs = t.GetCompoundStructure()

            if cs:

                for i in range(cs.LayerCount):

                    if cs.GetLayerFunction(i) == MaterialFunctionAssignment.Structure:

                        mat_id = cs.GetMaterialId(i)

                        m = doc.GetElement(mat_id)

                        if m:
                            return m.Name

    return "None"


# ===================================================
# MATERIAL DATA
# ===================================================
def collect_material_data():

    categories = [
        (BuiltInCategory.OST_StructuralFraming, "Beams"),
        (BuiltInCategory.OST_StructuralColumns, "Columns"),
        (BuiltInCategory.OST_StructuralFoundation, "Foundations"),
        (BuiltInCategory.OST_Walls, "Walls"),
        (BuiltInCategory.OST_Floors, "Floors"),
    ]

    data = {}
    element_map = {}

    for bic, cat_name in categories:

        elems = FilteredElementCollector(doc)\
            .OfCategory(bic)\
            .WhereElementIsNotElementType()\
            .ToElements()

        for e in elems:

            key = (cat_name, get_type_name(e), get_material_name(e))

            element_map.setdefault(key, []).append(e.Id)
            data[key] = data.get(key, 0) + 1

    return data, sorted(set(c for _, c in categories)), element_map


# ===================================================
# TEMPLATES DATA
# ===================================================
def get_views():
    return [
        v for v in FilteredElementCollector(doc).OfClass(View)
        if not v.IsTemplate
    ]


def collect_templates():
    return {
        t.Name: t
        for t in FilteredElementCollector(doc).OfClass(View)
        if t.IsTemplate
    }


# ===================================================
# UI
# ===================================================
def show_window():

    window = WPFWindow(xamlfile)

    # ---------------- MATERIAL TAB ----------------
    material_items = ObservableCollection[MaterialItem]()
    window.ReportGridMaterials.ItemsSource = material_items


    state = {
        "raw_data": None,
        "categories": None,
        "element_map": None
    }

    state["raw_data"], state["categories"], state["element_map"] = (
        collect_material_data()
    )

    materials = sorted(
        m.Name
        for m in FilteredElementCollector(doc).OfClass(Material)
        if m and m.Name
    )

    window.CategoryFilter.ItemsSource = state["categories"]
    window.MaterialCombo.ItemsSource = materials

    window.CategoryFilter.SelectedIndex = (
        0 if state["categories"] else -1
    )

    window.MaterialCombo.SelectedIndex = (
        0 if materials else -1
    )

    # ---------------- TEMPLATE TAB ----------------
    template_items = ObservableCollection[ReportItem]()
    window.ReportGridTemplates.ItemsSource = template_items

    checkbox_map = {}

    views = get_views()
    sheet_map = {}

    for vp in FilteredElementCollector(doc).OfClass(Viewport):

        view = doc.GetElement(vp.ViewId)
        sheet = doc.GetElement(vp.SheetId)

        if view and sheet:
            sheet_map[view.Id] = sheet


    templates = collect_templates()

    window.GlobalTemplateCombo.ItemsSource = sorted(templates.keys())

    if templates:
        window.GlobalTemplateCombo.SelectedIndex = 0


    # ===================================================
    # MATERIAL LOAD
    # ===================================================
    def load_material(selected_category=None):

        material_items.Clear()

        type_map = defaultdict(set)

        for (cat, typ, mat) in state["raw_data"].keys():
            type_map[(cat, typ)].add(mat)

        for (cat, typ, mat), count in sorted(
            state["raw_data"].items()
        ):

            if selected_category and cat != selected_category:
                continue

            material_items.Add(

                MaterialItem(
                    cat,
                    typ,
                    mat,
                    count,
                    len(type_map[(cat, typ)]) > 1,
                    state["element_map"][(cat, typ, mat)]
                )

            )

    def populate_views(filter_sheets=False):

        window.ViewListPanel.Items.Clear()
        checkbox_map.clear()

        if filter_sheets:

            grouped = {}

            for v in views:

                if v.Id in sheet_map:

                    sheet = sheet_map[v.Id]

                    key = "{} - {}".format(
                        sheet.SheetNumber,
                        sheet.Name
                    )

                    grouped.setdefault(key, []).append(v)

            for key in sorted(grouped.keys()):

                header = TextBlock()
                header.Text = key
                header.FontWeight = FontWeights.Bold
                header.Margin = Thickness(4, 8, 0, 2)

                window.ViewListPanel.Items.Add(header)

                for v in sorted(grouped[key], key=lambda x: x.Name):

                    cb = CheckBox()
                    cb.Content = v.Name
                    cb.Margin = Thickness(20, 2, 0, 2)

                    window.ViewListPanel.Items.Add(cb)

                    checkbox_map[cb] = v

        else:

            for v in sorted(views, key=lambda x: x.Name):

                cb = CheckBox()
                cb.Content = v.Name
                cb.Margin = Thickness(2)

                window.ViewListPanel.Items.Add(cb)

                checkbox_map[cb] = v

    def select_all(sender, args):

        for cb in checkbox_map:
            cb.IsChecked = True


    def unselect_all(sender, args):

        for cb in checkbox_map:
            cb.IsChecked = False

    def run_template_check(sender=None, args=None):

        template_items.Clear()

        selected_views = [
            checkbox_map[c]
            for c in checkbox_map
            if c.IsChecked
        ]

        if not selected_views:

            window.ReportHeader.Text = "No Views Selected"
            return

        missing = 0

        only_missing = (
            window.ChkOnlyMissingTemplates.IsChecked
        )

        for view in selected_views:

            has_tpl = (
                view.ViewTemplateId != ElementId(-1)
            )

            status = (
                "YES ✔"
                if has_tpl
                else "NO ❌"
            )

            if not has_tpl:

                missing += 1
                template_name = "None"

            else:

                tpl = doc.GetElement(
                    view.ViewTemplateId
                )

                template_name = (
                    tpl.Name
                    if tpl
                    else "None"
                )

            if only_missing and has_tpl:
                continue

            template_items.Add(

                ReportItem(
                    view.Name,
                    str(view.ViewType),
                    status,
                    template_name,
                    view
                )

            )

        window.ReportHeader.Text = (
            "Views: {} | Without Template: {}"
            .format(
                len(selected_views),
                missing
            )
        )

    def toggle_sheets(sender, args):

        populate_views(
            window.ChkPlacedOnSheets.IsChecked
        )


    def open_view(sender, args):

        item = window.ReportGridTemplates.SelectedItem

        if item:

            __revit__.ActiveUIDocument.ActiveView = (
                item.ElementObj
            )

    # ---------------------------------------------------
    # ENABLE/DISABLE CATEGORY FILTER UI
    # ---------------------------------------------------
    def toggle_filter_ui():
        window.CategoryFilter.IsEnabled = window.FilterByCategory.IsChecked

    # ---------------------------------------------------
    # FILTER
    # ---------------------------------------------------
    def toggle(sender, args):
        toggle_filter_ui()

        load_material(
            window.CategoryFilter.SelectedItem
            if window.FilterByCategory.IsChecked
            else None
        )

    def material_filter(sender, args):

        text = window.MaterialSearch.Text.lower()

        filtered = (
            [m for m in materials if text in m.lower()]
            if text else materials
        )

        window.MaterialCombo.ItemsSource = filtered

        # -------------------------
        # AUTO SELECT FIRST ITEM
        # -------------------------
        if filtered:
            window.MaterialCombo.SelectedIndex = 0
        else:
            window.MaterialCombo.SelectedIndex = -1
    # ===================================================
    # APPLY MATERIAL
    # ===================================================
    def apply_material(sender, args):

        selected = window.MaterialCombo.SelectedItem

        selected_rows = [
            i for i in material_items
            if i.Select
        ]

        if not selected or not selected_rows:
            return

        mat = next(
            (
                m
                for m in FilteredElementCollector(doc)
                .OfClass(Material)
                if m.Name == selected
            ),
            None
        )

        if not mat:
            return

        processed_types = set()

        t = Transaction(doc, "Apply Material")
        t.Start()

        for row in selected_rows:

            for eid in row.ElementIds:

                el = doc.GetElement(eid)

                if not el:
                    continue

                # -------------------------
                # Try instance parameter
                # -------------------------
                p = el.get_Parameter(
                    BuiltInParameter.STRUCTURAL_MATERIAL_PARAM
                )

                if p and not p.IsReadOnly:

                    p.Set(mat.Id)
                    continue

                # -------------------------
                # Try compound structure
                # -------------------------
                type_id = el.GetTypeId()

                if type_id in processed_types:
                    continue

                processed_types.add(type_id)

                typ = doc.GetElement(type_id)

                if not typ:
                    continue

                if hasattr(typ, "GetCompoundStructure"):

                    cs = typ.GetCompoundStructure()

                    if cs:

                        changed = False

                        for i in range(cs.LayerCount):

                            if (
                                cs.GetLayerFunction(i)
                                == MaterialFunctionAssignment.Structure
                            ):

                                cs.SetMaterialId(i, mat.Id)
                                changed = True

                        if changed:

                            typ.SetCompoundStructure(cs)
                            continue

                # -------------------------
                # Try type parameter
                # -------------------------
                p_type = typ.get_Parameter(
                    BuiltInParameter.STRUCTURAL_MATERIAL_PARAM
                )

                if p_type and not p_type.IsReadOnly:
                    p_type.Set(mat.Id)

        t.Commit()

        # -------------------------
        # FULL REFRESH 
        # -------------------------
        state["raw_data"], \
        state["categories"], \
        state["element_map"] = collect_material_data()

        materials = sorted(
            m.Name
            for m in FilteredElementCollector(doc).OfClass(Material)
            if m and m.Name
        )

        window.CategoryFilter.ItemsSource = state["categories"]
        window.MaterialCombo.ItemsSource = materials

        load_material(
            window.CategoryFilter.SelectedItem
            if window.FilterByCategory.IsChecked
            else None
        )

        window.ReportGridMaterials.Items.Refresh()

        window.Title = "Material Applied"


    # ===================================================
    # APPLY TEMPLATE
    # ===================================================
    def apply_template(sender, args):

        name = window.GlobalTemplateCombo.SelectedItem
        tpl = templates.get(name)

        if not tpl:
            return

        selected = [i for i in template_items if i.Select]

        t = Transaction(doc, "Apply Template")
        t.Start()

        for i in selected:
            i.ElementObj.ViewTemplateId = tpl.Id

        t.Commit()

        run_template_check(None, None)


    # ===================================================
    # EVENTS
    # ===================================================
    window.BtnApplyMaterial.Click += apply_material
    window.BtnRunTemplate.Click += run_template_check

    window.BtnApplyTemplate.Click += apply_template

    window.BtnSelectAll.Click += select_all
    window.BtnUnselectAll.Click += unselect_all

    window.ChkPlacedOnSheets.Checked += toggle_sheets
    window.ChkPlacedOnSheets.Unchecked += toggle_sheets

    window.ChkOnlyMissingTemplates.Checked += run_template_check
    window.ChkOnlyMissingTemplates.Unchecked += run_template_check

    window.ReportGridTemplates.MouseDoubleClick += open_view

    window.FilterByCategory.Checked += toggle
    window.FilterByCategory.Unchecked += toggle

    window.CategoryFilter.SelectionChanged += (
        lambda s, e: toggle(None, None)
    )

    window.MaterialSearch.TextChanged += material_filter

    populate_views(False)
    load_material()
    run_template_check()

    window.ShowDialog()


show_window()