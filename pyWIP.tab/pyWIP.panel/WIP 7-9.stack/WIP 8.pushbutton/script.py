# -*- coding: utf-8 -*-

__author__ = "Pankaj Prabhakar"
__version__ = "Version 2.4"

__doc__ = """
Load View Templates From Link

Read-only comparison workflow:
1. Select a loaded Revit Link.
2. Select one or more linked View Templates.
3. Compare source templates directly with matching destination templates.
4. Show a pyRevit HTML comparison report.
5. Copy only after Continue is selected.
6. Copy new templates and, when Override is enabled, replace changed templates.
7. Leave unchanged, partially compared, and failed templates untouched.

Compared areas:
- Requested View Template parameter values and Include states
- Model Categories
- Annotation Categories
- Filters
- Worksets
- Revit Links, best effort and name-mapped

Required XAML file:
CopyViewTemplate.xaml
"""

import os
import traceback
from collections import defaultdict

from Autodesk.Revit.DB import (
    Category,
    CopyPasteOptions,
    DuplicateTypeAction,
    ElementId,
    ElementTransformUtils,
    FilteredElementCollector,
    IDuplicateTypeNamesHandler,
    ParameterFilterElement,
    RevitLinkInstance,
    RevitLinkType,
    SelectionFilterElement,
    Transaction,
    TransactionGroup,
    Transform,
    View
)

try:
    from Autodesk.Revit.DB import FilteredWorksetCollector, WorksetKind
except Exception:
    FilteredWorksetCollector = None
    WorksetKind = None

from pyrevit import forms, script
from System.Collections.Generic import List
from System.Windows import Visibility, RoutedEventHandler
from System.Windows.Controls.Primitives import ToggleButton

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

WINDOW_TITLE = "Load View Templates From Link"
XAML_FILE = os.path.join(os.path.dirname(__file__), "CopyViewTemplate.xaml")

COMPARE_MODEL_CATEGORIES = True
COMPARE_ANNOTATION_CATEGORIES = True
COMPARE_FILTERS = True
COMPARE_WORKSETS = True
COMPARE_REVIT_LINKS = True

REQUESTED_PARAMETER_NAMES = {
    "view scale",
    "scale value 1:",
    "scale value 1",
    "detail level",
    "parts visibility",
    "model display",
    "shadows",
    "sketchy lines",
    "lighting",
    "photographic exposure",
    "background",
    "phase filter",
    "discipline",
    "show hidden lines",
    "rendering settings",
    "bim fabrication release",
    "detail name",
    "future build",
    "not in contract",
    "package name",
    "rfi number",
    "view sub-type",
    "view subtype",
    "view type"
}


class UseDestinationTypesHandler(IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DuplicateTypeAction.UseDestinationTypes


DUPLICATE_TYPE_HANDLER = UseDestinationTypesHandler()


def safe_element_name(element):
    try:
        return element.Name
    except Exception:
        return "<Unnamed Element>"


def element_id_value(element_id):
    if element_id is None:
        return None
    try:
        return element_id.Value
    except Exception:
        pass
    try:
        return element_id.IntegerValue
    except Exception:
        return str(element_id)


def is_invalid_element_id(element_id):
    if element_id is None:
        return True
    try:
        return element_id == ElementId.InvalidElementId
    except Exception:
        return element_id_value(element_id) == -1


def safe_call(obj, method_name, default=None):
    if obj is None:
        return default
    try:
        method = getattr(obj, method_name, None)
        return method() if method is not None else default
    except Exception:
        return default


def safe_property(obj, property_name, default=None):
    if obj is None:
        return default
    try:
        return getattr(obj, property_name)
    except Exception:
        return default


def normalize_name(value):
    return "" if value is None else str(value).strip().lower()


def normalize_value(value):
    return "" if value is None else " ".join(str(value).strip().lower().split())


def values_equal(first, second):
    return normalize_value(first) == normalize_value(second)


def collection_id_values(element_ids):
    values = set()
    if element_ids is not None:
        for element_id in element_ids:
            values.add(element_id_value(element_id))
    return values


def unique_rows(rows):
    result = []
    seen = set()
    for row in rows:
        key = tuple(str(value) for value in row)
        if key not in seen:
            seen.add(key)
            result.append(row)
    return result


def html_escape(value):
    if value is None:
        return ""
    return (str(value).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;").replace("'", "&#39;"))


def html_cell(value):
    return html_escape(value).replace("\n", "<br/>")


def add_report_styles(output):
    css = """
    .vt-report{width:100%;margin:14px 0 26px;color:#172033;font-family:'Segoe UI',Arial,sans-serif}
    .vt-report *{box-sizing:border-box}.vt-banner{margin:8px 0 16px;padding:10px 13px;border-left:5px solid #2589bd;background:#eaf4fa;line-height:1.45}
    .vt-template{margin:22px 0 30px;border:1px solid #aab4bd;border-radius:6px;background:#fff}.vt-template-head{padding:11px 13px;border-bottom:1px solid #aab4bd;background:#243c50;color:#fff}
    .vt-template-head h2{margin:0;color:#fff;font-size:17px}.vt-status{margin-top:5px;font-size:12px}.vt-section{padding:13px;border-bottom:1px solid #ccd2d7}.vt-section:last-child{border-bottom:0}
    .vt-section h3{margin:0 0 9px;color:#17364a;font-size:14px}.vt-table-wrap{width:100%;overflow-x:auto;border:1px solid #aab4bd;background:#fff}
    .vt-table{width:100%;min-width:980px;table-layout:fixed;border-collapse:collapse}.vt-item{width:25%}.vt-property{width:23%}.vt-existing,.vt-linked{width:26%}
    .vt-table th{padding:8px 10px;border-right:1px solid #576e80;border-bottom:1px solid #576e80;background:#385568;color:#fff;text-align:left;vertical-align:top;font-size:11px;white-space:normal;overflow-wrap:break-word}
    .vt-table td{padding:8px 10px;border-right:1px solid #c7cdd2;border-bottom:1px solid #bec5ca;color:#172033;text-align:left;vertical-align:top;font-family:Consolas,'Courier New',monospace;font-size:10px;line-height:1.45;white-space:normal;overflow-wrap:break-word}
    .vt-table th:last-child,.vt-table td:last-child{border-right:0}.vt-table tbody tr:last-child td{border-bottom:0}.vt-table tbody tr:nth-child(even) td{background:#f3f5f7}
    .vt-table td:first-child,.vt-table td:nth-child(2){font-family:'Segoe UI',Arial,sans-serif}.vt-new{color:#145da0;font-weight:600}.vt-changed{color:#0b6b37;font-weight:600}
    .vt-unchanged{color:#69747e}.vt-partial{color:#806000;font-weight:600}.vt-failed{color:#a11a1a;font-weight:600}.vt-warning{margin:8px 0 0;padding:8px 10px;border-left:4px solid #c28b00;background:#fff7df;color:#674b00;font-size:11px;line-height:1.4}
    """
    try:
        output.add_style(css)
    except Exception:
        output.print_html("<style>{}</style>".format(css))


def status_css_class(status):
    return {
        "new": "vt-new",
        "changed": "vt-changed",
        "unchanged": "vt-unchanged",
        "partially compared": "vt-partial",
        "failed": "vt-failed"
    }.get(normalize_name(status), "")


def print_section_html(output, section):
    rows = section.get("rows", [])
    warnings = section.get("warnings", [])
    parts = ['<div class="vt-section"><h3>{}</h3>'.format(html_escape(section.get("section", "Comparison")))]
    if rows:
        parts.extend(['<div class="vt-table-wrap"><table class="vt-table"><colgroup><col class="vt-item"/><col class="vt-property"/><col class="vt-existing"/><col class="vt-linked"/></colgroup>',
                      '<thead><tr><th>Item</th><th>Property</th><th>Current Project</th><th>Linked Template</th></tr></thead><tbody>'])
        for row in rows:
            values = list(row) + ["", "", "", ""]
            parts.append("<tr><td>{}</td><td>{}</td><td>{}</td><td>{}</td></tr>".format(
                html_cell(values[0]), html_cell(values[1]), html_cell(values[2]), html_cell(values[3])))
        parts.append("</tbody></table></div>")
    else:
        parts.append("<p>No differences detected in this section.</p>")
    for warning in warnings:
        parts.append('<div class="vt-warning">{}</div>'.format(html_cell(warning)))
    parts.append("</div>")
    output.print_html("".join(parts))


def print_template_comparison_report(linked_document, results):
    output = script.get_output()
    try:
        output.set_title("View Template Comparison Report")
    except Exception:
        pass
    add_report_styles(output)
    counts = defaultdict(int)
    for result in results:
        counts[result["status"]] += 1
    output.print_md("# View Template Comparison Report")
    output.print_html('<div class="vt-banner"><strong>No project changes have been made.</strong><br/>Templates were compared directly. Copying starts only after Continue is selected.</div>')
    output.print_table(
        table_data=[
            ["Source link", linked_document.Title],
            ["Selected templates", str(len(results))],
            ["New", str(counts["New"])],
            ["Changed", str(counts["Changed"])],
            ["Unchanged", str(counts["Unchanged"])],
            ["Partially compared", str(counts["Partially Compared"])],
            ["Failed", str(counts["Failed"])]
        ],
        columns=["Item", "Result"],
        title="Comparison Summary"
    )
    for result in sorted(results, key=lambda item: item["template_name"].lower()):
        changed_text = ", ".join(result.get("changed_sections", [])) or "None"
        output.print_html(
            '<div class="vt-template"><div class="vt-template-head"><h2>{}</h2><div class="vt-status">Status: <span class="{}">{}</span><br/>Changed sections: {}</div></div>'.format(
                html_escape(result["template_name"]), status_css_class(result["status"]), html_escape(result["status"]), html_escape(changed_text)))
        sections = result.get("sections", [])
        if sections:
            for section in sections:
                print_section_html(output, section)
        else:
            message = "The entire template will be copied." if result["status"] == "New" else "No comparison sections were generated."
            output.print_html('<div class="vt-section">{}</div>'.format(html_cell(message)))
        for warning in result.get("warnings", []):
            output.print_html('<div class="vt-section"><div class="vt-warning">{}</div></div>'.format(html_cell(warning)))
        output.print_html("</div>")


def collect_loaded_links():
    result = {}
    instances = []
    for link_instance in (FilteredElementCollector(doc).OfClass(RevitLinkInstance)
                          .WhereElementIsNotElementType().ToElements()):
        try:
            link_document = link_instance.GetLinkDocument()
        except Exception:
            link_document = None
        if link_document is None:
            continue
        try:
            if link_document.IsFamilyDocument:
                continue
        except Exception:
            continue
        try:
            base_name = link_instance.Name
        except Exception:
            base_name = link_document.Title
        instances.append((base_name, link_instance, link_document))
    counts = defaultdict(int)
    for base_name, _, _ in instances:
        counts[base_name] += 1
    for base_name, instance, link_document in instances:
        display_name = ("{} | Instance ID {}".format(base_name, element_id_value(instance.Id))
                        if counts[base_name] > 1 else base_name)
        result[display_name] = link_document
    return result


def collect_view_templates(owner_document):
    templates = []
    for view in (FilteredElementCollector(owner_document).OfClass(View)
                 .WhereElementIsNotElementType().ToElements()):
        try:
            if view.IsTemplate:
                templates.append(view)
        except Exception:
            continue
    return sorted(templates, key=lambda item: safe_element_name(item).lower())


def collect_destination_templates():
    return collect_view_templates(doc)


def find_destination_template(template_name, source_template=None):
    fallback = None
    for destination in collect_destination_templates():
        if safe_element_name(destination) != template_name:
            continue
        if source_template is None:
            return destination
        try:
            if destination.ViewType == source_template.ViewType:
                return destination
        except Exception:
            fallback = destination
    return fallback


def get_parameter_display_name(parameter):
    if parameter is None:
        return "<Parameter Not Found>"
    try:
        return parameter.Definition.Name
    except Exception:
        return "<Unknown Parameter>"


def parameter_identity(parameter):
    if parameter is None:
        return None
    try:
        if parameter.IsShared and parameter.GUID is not None:
            return ("shared", str(parameter.GUID).lower())
    except Exception:
        pass
    try:
        parameter_id = element_id_value(parameter.Id)
        if isinstance(parameter_id, int) and parameter_id < 0:
            return ("builtin", parameter_id)
    except Exception:
        pass
    return ("definition", normalize_name(get_parameter_display_name(parameter)),
            str(safe_property(parameter, "StorageType", "")))


def template_parameters_by_identity(template):
    result = {}
    try:
        parameters = template.Parameters
    except Exception:
        parameters = []
    for parameter in parameters:
        identity = parameter_identity(parameter)
        if identity is not None:
            result[identity] = parameter
    return result


def parameter_element_name(owner_document, element_id):
    if is_invalid_element_id(element_id):
        return "<None>"
    try:
        element = owner_document.GetElement(element_id)
        if element is not None:
            return safe_element_name(element)
    except Exception:
        pass
    return "Element ID {}".format(element_id_value(element_id))


def parameter_display_value(owner_document, parameter):
    if parameter is None:
        return "<Parameter Not Found>"
    try:
        if not parameter.HasValue:
            return "<No Value>"
    except Exception:
        pass
    try:
        value_string = parameter.AsValueString()
        if value_string not in (None, ""):
            return value_string
    except Exception:
        pass
    storage = str(safe_property(parameter, "StorageType", ""))
    try:
        if "String" in storage:
            return parameter.AsString() or ""
        if "Integer" in storage:
            return str(parameter.AsInteger())
        if "Double" in storage:
            return str(parameter.AsDouble())
        if "ElementId" in storage:
            return parameter_element_name(owner_document, parameter.AsElementId())
    except Exception:
        return "<Unreadable Value>"
    return "<Unreadable Value>"


def non_controlled_parameter_keys(template):
    non_controlled_ids = collection_id_values(
        safe_call(template, "GetNonControlledTemplateParameterIds", []))
    keys = set()
    try:
        parameters = template.Parameters
    except Exception:
        parameters = []
    for parameter in parameters:
        try:
            if element_id_value(parameter.Id) in non_controlled_ids:
                identity = parameter_identity(parameter)
                if identity is not None:
                    keys.add(identity)
        except Exception:
            continue
    return keys


def compare_requested_parameters(existing_template, linked_template, existing_document, linked_document):
    rows = []
    existing_parameters = template_parameters_by_identity(existing_template)
    linked_parameters = template_parameters_by_identity(linked_template)
    existing_non_controlled = non_controlled_parameter_keys(existing_template)
    linked_non_controlled = non_controlled_parameter_keys(linked_template)
    keys = set(existing_parameters.keys()) | set(linked_parameters.keys())
    for key in sorted(keys, key=lambda value: str(value)):
        existing_parameter = existing_parameters.get(key)
        linked_parameter = linked_parameters.get(key)
        representative = linked_parameter if linked_parameter is not None else existing_parameter
        name = get_parameter_display_name(representative)
        if normalize_name(name) not in REQUESTED_PARAMETER_NAMES:
            continue
        existing_value = parameter_display_value(existing_document, existing_parameter)
        linked_value = parameter_display_value(linked_document, linked_parameter)
        if not values_equal(existing_value, linked_value):
            rows.append([name, "Value", existing_value, linked_value])
        existing_included = existing_parameter is not None and key not in existing_non_controlled
        linked_included = linked_parameter is not None and key not in linked_non_controlled
        if existing_included != linked_included:
            rows.append([name, "Include", "Included" if existing_included else "Excluded",
                         "Included" if linked_included else "Excluded"])
    return {"section": "Requested Parameters", "rows": unique_rows(rows), "warnings": []}


def category_path(category):
    if category is None:
        return "<Unknown Category>"
    names = []
    current = category
    while current is not None:
        try:
            names.insert(0, current.Name)
            current = current.Parent
        except Exception:
            break
    return " > ".join(names)


def root_category(category):
    current = category
    while current is not None:
        try:
            parent = current.Parent
        except Exception:
            parent = None
        if parent is None:
            return current
        current = parent
    return category


def is_imported_category(category):
    """Identify dynamic CAD/import category trees shown on Imported Categories."""
    root = root_category(category)
    if root is None:
        return False
    try:
        root_id = element_id_value(root.Id)
        # Built-in Revit category ids are negative. Imported file roots are
        # document-specific positive ids. Test only the root so ordinary custom
        # subcategories below a built-in category are not excluded.
        if isinstance(root_id, int) and root_id > 0:
            return True
    except Exception:
        pass
    try:
        name = str(root.Name).lower()
        return name.endswith((".dwg", ".dxf", ".dgn", ".sat", ".skp"))
    except Exception:
        return False


def category_identity(category):
    """Use stable built-in root ids, plus subcategory names, across documents."""
    root = root_category(category)
    root_id = element_id_value(root.Id) if root is not None else None
    descendants = []
    current = category
    while current is not None and current != root:
        try:
            descendants.insert(0, normalize_name(current.Name))
            current = current.Parent
        except Exception:
            break
    if isinstance(root_id, int) and root_id < 0:
        return ("builtin", root_id, tuple(descendants))
    return ("path", normalize_name(category_path(category)))


def iter_category_tree(category):
    yield category
    try:
        children = list(category.SubCategories)
    except Exception:
        children = []
    for child in children:
        for nested in iter_category_tree(child):
            yield nested


def category_lookup(owner_document, category_type_name):
    result = {}
    skipped_import_roots = set()
    try:
        categories = owner_document.Settings.Categories
    except Exception:
        return result, skipped_import_roots
    for top_category in categories:
        try:
            if category_type_name.lower() not in str(top_category.CategoryType).lower():
                continue
        except Exception:
            continue
        if is_imported_category(top_category):
            skipped_import_roots.add(category_path(root_category(top_category)))
            continue
        for category in iter_category_tree(top_category):
            result[category_identity(category)] = category
    return result, skipped_import_roots


def color_text(color):
    if color is None:
        return "By Object Styles"
    try:
        if not color.IsValid:
            return "By Object Styles"
    except Exception:
        pass
    try:
        return "RGB({}, {}, {})".format(color.Red, color.Green, color.Blue)
    except Exception:
        return str(color)


def referenced_element_name(owner_document, element_id, default_text="By Object Styles"):
    if is_invalid_element_id(element_id):
        return default_text
    try:
        element = owner_document.GetElement(element_id)
        if element is not None:
            return safe_element_name(element)
    except Exception:
        pass
    return "Element ID {}".format(element_id_value(element_id))


def readable_line_weight(value):
    if value in (None, -1):
        return "By Object Styles"
    return str(value)


def readable_detail_level(value):
    if value is None:
        return "By View"
    text = str(value)
    if normalize_name(text) in ("undefined", "invalid", "-1", ""):
        return "By View"
    return text


def override_settings_dictionary(owner_document, settings):
    """Return all category V/G columns exposed by OverrideGraphicSettings."""
    if settings is None:
        return {}
    return {
        "Projection/Surface Lines - Weight": readable_line_weight(
            safe_property(settings, "ProjectionLineWeight", None)),
        "Projection/Surface Lines - Color": color_text(
            safe_property(settings, "ProjectionLineColor", None)),
        "Projection/Surface Lines - Pattern": referenced_element_name(
            owner_document, safe_property(settings, "ProjectionLinePatternId", None)),
        "Projection/Surface Pattern - Foreground": referenced_element_name(
            owner_document, safe_property(settings, "SurfaceForegroundPatternId", None), "None"),
        "Projection/Surface Pattern - Foreground Color": color_text(
            safe_property(settings, "SurfaceForegroundPatternColor", None)),
        "Projection/Surface Pattern - Foreground Visible": str(
            safe_property(settings, "IsSurfaceForegroundPatternVisible", False)),
        "Projection/Surface Pattern - Background": referenced_element_name(
            owner_document, safe_property(settings, "SurfaceBackgroundPatternId", None), "None"),
        "Projection/Surface Pattern - Background Color": color_text(
            safe_property(settings, "SurfaceBackgroundPatternColor", None)),
        "Projection/Surface Pattern - Background Visible": str(
            safe_property(settings, "IsSurfaceBackgroundPatternVisible", False)),
        "Projection/Surface - Transparency": str(
            safe_property(settings, "Transparency", 0)),
        "Cut Lines - Weight": readable_line_weight(
            safe_property(settings, "CutLineWeight", None)),
        "Cut Lines - Color": color_text(
            safe_property(settings, "CutLineColor", None)),
        "Cut Lines - Pattern": referenced_element_name(
            owner_document, safe_property(settings, "CutLinePatternId", None)),
        "Cut Pattern - Foreground": referenced_element_name(
            owner_document, safe_property(settings, "CutForegroundPatternId", None), "None"),
        "Cut Pattern - Foreground Color": color_text(
            safe_property(settings, "CutForegroundPatternColor", None)),
        "Cut Pattern - Foreground Visible": str(
            safe_property(settings, "IsCutForegroundPatternVisible", False)),
        "Cut Pattern - Background": referenced_element_name(
            owner_document, safe_property(settings, "CutBackgroundPatternId", None), "None"),
        "Cut Pattern - Background Color": color_text(
            safe_property(settings, "CutBackgroundPatternColor", None)),
        "Cut Pattern - Background Visible": str(
            safe_property(settings, "IsCutBackgroundPatternVisible", False)),
        "Halftone": str(safe_property(settings, "Halftone", False)),
        "Detail Level": readable_detail_level(
            safe_property(settings, "DetailLevel", None))
    }


def compare_category_section(existing_template, linked_template, existing_document,
                             linked_document, category_type_name, section_name):
    rows = []
    warnings = []
    existing_categories, existing_imports = category_lookup(
        existing_document, category_type_name)
    linked_categories, linked_imports = category_lookup(
        linked_document, category_type_name)

    if existing_imports or linked_imports:
        warnings.append(
            "Imported CAD category trees were excluded from {}: {} existing, {} linked. "
            "They belong to Revit's Imported Categories tab and use document-specific "
            "category/layer ids, so reporting every CAD layer as a missing model category "
            "would create false differences.".format(
                section_name, len(existing_imports), len(linked_imports)))

    for key in sorted(set(existing_categories.keys()) | set(linked_categories.keys()),
                      key=lambda value: str(value)):
        existing_category = existing_categories.get(key)
        linked_category = linked_categories.get(key)
        representative = linked_category if linked_category is not None else existing_category
        name = category_path(representative)

        if existing_category is None:
            rows.append([name, "Category", "Not present", "Present"])
            continue
        if linked_category is None:
            rows.append([name, "Category", "Present", "Not present"])
            continue

        try:
            existing_hidden = existing_template.GetCategoryHidden(existing_category.Id)
        except Exception:
            existing_hidden = "<Not supported>"
        try:
            linked_hidden = linked_template.GetCategoryHidden(linked_category.Id)
        except Exception:
            linked_hidden = "<Not supported>"
        existing_visibility = ("Hidden" if existing_hidden is True else
                               "Visible" if existing_hidden is False else str(existing_hidden))
        linked_visibility = ("Hidden" if linked_hidden is True else
                             "Visible" if linked_hidden is False else str(linked_hidden))
        if not values_equal(existing_visibility, linked_visibility):
            rows.append([name, "Visibility", existing_visibility, linked_visibility])

        try:
            existing_ogs = existing_template.GetCategoryOverrides(existing_category.Id)
            existing_values = override_settings_dictionary(existing_document, existing_ogs)
        except Exception:
            existing_values = {}
        try:
            linked_ogs = linked_template.GetCategoryOverrides(linked_category.Id)
            linked_values = override_settings_dictionary(linked_document, linked_ogs)
        except Exception:
            linked_values = {}

        for property_name in sorted(set(existing_values.keys()) | set(linked_values.keys())):
            existing_value = existing_values.get(property_name, "<Not supported>")
            linked_value = linked_values.get(property_name, "<Not supported>")
            if not values_equal(existing_value, linked_value):
                rows.append([name, property_name, existing_value, linked_value])

    return {"section": section_name, "rows": unique_rows(rows), "warnings": warnings}

def applied_filter_data(owner_document, template):
    result = {}
    filter_ids = safe_call(template, "GetOrderedFilters", []) or safe_call(template, "GetFilters", [])
    for order, filter_id in enumerate(filter_ids, start=1):
        try:
            filter_element = owner_document.GetElement(filter_id)
        except Exception:
            filter_element = None
        name = safe_element_name(filter_element) if filter_element is not None else "Filter ID {}".format(element_id_value(filter_id))
        filter_type = filter_element.GetType().Name if filter_element is not None else "<Unknown>"
        try:
            visibility = template.GetFilterVisibility(filter_id)
        except Exception:
            visibility = "<Unavailable>"
        try:
            enabled = template.GetIsFilterEnabled(filter_id)
        except Exception:
            enabled = "<Unavailable>"
        try:
            overrides = template.GetFilterOverrides(filter_id)
        except Exception:
            overrides = None
        result[(normalize_name(name), normalize_name(filter_type))] = {
            "name": name,
            "order": order,
            "visibility": visibility,
            "enabled": enabled,
            "overrides": override_settings_dictionary(owner_document, overrides)
        }
    return result


def compare_applied_filters(existing_template, linked_template, existing_document, linked_document):
    rows = []
    existing_filters = applied_filter_data(existing_document, existing_template)
    linked_filters = applied_filter_data(linked_document, linked_template)
    for key in sorted(set(existing_filters.keys()) | set(linked_filters.keys()), key=lambda value: str(value)):
        existing_data = existing_filters.get(key)
        linked_data = linked_filters.get(key)
        representative = linked_data if linked_data is not None else existing_data
        name = representative["name"]
        if existing_data is None:
            rows.append([name, "Applied", "No", "Yes"])
            continue
        if linked_data is None:
            rows.append([name, "Applied", "Yes", "No"])
            continue
        for field, label in (("order", "Order"), ("visibility", "Visibility"), ("enabled", "Enabled")):
            if existing_data[field] != linked_data[field]:
                rows.append([name, label, str(existing_data[field]), str(linked_data[field])])
        for prop in sorted(set(existing_data["overrides"].keys()) | set(linked_data["overrides"].keys())):
            existing_value = existing_data["overrides"].get(prop, "<Unavailable>")
            linked_value = linked_data["overrides"].get(prop, "<Unavailable>")
            if not values_equal(existing_value, linked_value):
                rows.append([name, prop, existing_value, linked_value])
    return {"section": "Filters", "rows": unique_rows(rows), "warnings": []}


def collect_user_worksets(owner_document):
    result = {}
    if FilteredWorksetCollector is None:
        return result
    try:
        collector = FilteredWorksetCollector(owner_document)
        if WorksetKind is not None:
            collector = collector.OfKind(WorksetKind.UserWorkset)
        for workset in collector:
            result[normalize_name(workset.Name)] = workset
    except Exception:
        pass
    return result


def compare_worksets(existing_template, linked_template, existing_document, linked_document):
    rows = []
    warnings = []
    if FilteredWorksetCollector is None:
        warnings.append("Workset comparison is unavailable in this Revit API environment.")
        return {"section": "Worksets", "rows": rows, "warnings": warnings}
    existing_worksets = collect_user_worksets(existing_document)
    linked_worksets = collect_user_worksets(linked_document)
    for key in sorted(set(existing_worksets.keys()) | set(linked_worksets.keys())):
        existing_workset = existing_worksets.get(key)
        linked_workset = linked_worksets.get(key)
        representative = linked_workset if linked_workset is not None else existing_workset
        name = representative.Name
        if existing_workset is None:
            rows.append([name, "Workset availability", "Missing", "Available"])
            continue
        if linked_workset is None:
            rows.append([name, "Workset availability", "Available", "Missing"])
            continue
        try:
            existing_visibility = str(existing_template.GetWorksetVisibility(existing_workset.Id))
        except Exception:
            existing_visibility = "<Unavailable>"
        try:
            linked_visibility = str(linked_template.GetWorksetVisibility(linked_workset.Id))
        except Exception:
            linked_visibility = "<Unavailable>"
        if not values_equal(existing_visibility, linked_visibility):
            rows.append([name, "Visibility", existing_visibility, linked_visibility])
    return {"section": "Worksets", "rows": unique_rows(rows), "warnings": warnings}


def collect_link_types(owner_document):
    result = {}
    try:
        elements = (FilteredElementCollector(owner_document).OfClass(RevitLinkType)
                    .WhereElementIsElementType().ToElements())
    except Exception:
        elements = []
    for element in elements:
        result[normalize_name(safe_element_name(element))] = element
    return result


def compare_revit_links(existing_template, linked_template, existing_document, linked_document):
    rows = []
    existing_links = collect_link_types(existing_document)
    linked_links = collect_link_types(linked_document)
    for key in sorted(set(existing_links.keys()) | set(linked_links.keys())):
        existing_link = existing_links.get(key)
        linked_link = linked_links.get(key)
        representative = linked_link if linked_link is not None else existing_link
        name = safe_element_name(representative)
        if existing_link is None:
            rows.append([name, "Link availability", "Missing", "Available"])
        elif linked_link is None:
            rows.append([name, "Link availability", "Available", "Missing"])
        else:
            # Public API support for custom RVT-link display differs between releases.
            try:
                existing_hidden = existing_template.GetCategoryHidden(existing_link.Id)
            except Exception:
                existing_hidden = "<Unavailable>"
            try:
                linked_hidden = linked_template.GetCategoryHidden(linked_link.Id)
            except Exception:
                linked_hidden = "<Unavailable>"
            if existing_hidden != linked_hidden:
                rows.append([name, "Visibility", str(existing_hidden), str(linked_hidden)])
    return {
        "section": "Revit Links",
        "rows": unique_rows(rows),
        "warnings": ["Revit Link comparison is best-effort and name-mapped. Custom linked-view settings are reported only when exposed by the active Revit API version."]
    }


def compare_view_templates_directly(existing_template, linked_template, linked_document):
    sections = [compare_requested_parameters(existing_template, linked_template, doc, linked_document)]
    if COMPARE_MODEL_CATEGORIES:
        sections.append(compare_category_section(existing_template, linked_template, doc, linked_document, "Model", "Model Categories"))
    if COMPARE_ANNOTATION_CATEGORIES:
        sections.append(compare_category_section(existing_template, linked_template, doc, linked_document, "Annotation", "Annotation Categories"))
    if COMPARE_FILTERS:
        sections.append(compare_applied_filters(existing_template, linked_template, doc, linked_document))
    if COMPARE_WORKSETS:
        sections.append(compare_worksets(existing_template, linked_template, doc, linked_document))
    if COMPARE_REVIT_LINKS:
        sections.append(compare_revit_links(existing_template, linked_template, doc, linked_document))
    changed_sections = [section["section"] for section in sections if section.get("rows")]
    warnings = []
    for section in sections:
        warnings.extend(section.get("warnings", []))
    status = "Changed" if changed_sections else "Partially Compared" if warnings else "Unchanged"
    return {
        "template_name": safe_element_name(linked_template),
        "status": status,
        "changed_sections": changed_sections,
        "sections": sections,
        "warnings": warnings
    }


def run_direct_template_comparison(linked_document, selected_templates):
    results = []
    for linked_template in selected_templates:
        name = safe_element_name(linked_template)
        existing_template = find_destination_template(name, linked_template)
        if existing_template is None:
            results.append({"template_name": name, "status": "New", "changed_sections": ["Entire Template"], "sections": [], "warnings": []})
            continue
        try:
            results.append(compare_view_templates_directly(existing_template, linked_template, linked_document))
        except Exception:
            results.append({"template_name": name, "status": "Failed", "changed_sections": [], "sections": [], "warnings": [traceback.format_exc()]})
    return results


class ViewTemplateItem(object):
    def __init__(self, element):
        self.element = element
        self.Name = safe_element_name(element)
        self.IsChecked = False


def record_template_assignments(template_names):
    selected_names = set(template_names)
    assignments = defaultdict(list)
    templates_by_id = {element_id_value(template.Id): template for template in collect_destination_templates()}
    for view in (FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType().ToElements()):
        try:
            if view.IsTemplate or is_invalid_element_id(view.ViewTemplateId):
                continue
            template = templates_by_id.get(element_id_value(view.ViewTemplateId))
            if template is not None and safe_element_name(template) in selected_names:
                assignments[safe_element_name(template)].append(view.Id)
        except Exception:
            continue
    return assignments


def delete_matching_destination_templates(template_names):
    names = set(template_names)
    deleted = []
    for template in collect_destination_templates():
        name = safe_element_name(template)
        if name in names:
            doc.Delete(template.Id)
            deleted.append(name)
    return deleted


def reassign_replacement_templates(assignments):
    replacements = {safe_element_name(template): template for template in collect_destination_templates()}
    warnings = []
    for name, view_ids in assignments.items():
        replacement = replacements.get(name)
        if replacement is None:
            warnings.append("Replacement View Template was not found: {}".format(name))
            continue
        for view_id in view_ids:
            view = doc.GetElement(view_id)
            if view is None:
                continue
            try:
                view.ViewTemplateId = replacement.Id
            except Exception as error:
                warnings.append("Could not assign '{}' to view ID {}: {}".format(name, element_id_value(view_id), error))
    return warnings


def create_element_id_list(elements):
    ids = List[ElementId]()
    for element in elements:
        ids.Add(element.Id)
    return ids


def copy_view_templates(source_document, source_templates):
    options = CopyPasteOptions()
    options.SetDuplicateTypeNamesHandler(DUPLICATE_TYPE_HANDLER)
    return ElementTransformUtils.CopyElements(
        source_document, create_element_id_list(source_templates), doc,
        Transform.Identity, options)


def apply_template_transfer(source_document, templates_to_process, names_to_override):
    assignments = {}
    warnings = []
    copied_ids = None
    group = TransactionGroup(doc, "Load Selected View Templates From Link")
    started = False
    try:
        group.Start()
        started = True
        if names_to_override:
            assignments = record_template_assignments(names_to_override)
            transaction = Transaction(doc, "Remove Changed View Templates")
            transaction.Start()
            try:
                delete_matching_destination_templates(names_to_override)
                transaction.Commit()
            except Exception:
                transaction.RollBack()
                raise
        transaction = Transaction(doc, "Copy View Templates From Link")
        transaction.Start()
        try:
            copied_ids = copy_view_templates(source_document, templates_to_process)
            transaction.Commit()
        except Exception:
            transaction.RollBack()
            raise
        if assignments:
            transaction = Transaction(doc, "Reassign Replacement View Templates")
            transaction.Start()
            try:
                warnings = reassign_replacement_templates(assignments)
                transaction.Commit()
            except Exception:
                transaction.RollBack()
                raise
        group.Assimilate()
        started = False
    except Exception:
        if started:
            try:
                group.RollBack()
            except Exception:
                pass
        raise
    return {"copied_ids": copied_ids, "warnings": warnings}


class LoadViewTemplatesWindow(forms.WPFWindow):
    def __init__(self):
        if not os.path.exists(XAML_FILE):
            forms.alert("The UI file was not found:\n\n{}".format(XAML_FILE), title=WINDOW_TITLE, exitscript=True)
        forms.WPFWindow.__init__(self, XAML_FILE)
        self.links = collect_loaded_links()
        self.source_doc = None
        self.all_template_items = []
        try:
            self.main_title.Text = WINDOW_TITLE
            self.footer_version.Text = __version__
        except Exception:
            pass
        self.UI_Stack_ViewTemplates.Visibility = Visibility.Collapsed
        self.UI_btn_Run.IsEnabled = False

        # The XAML binds CheckBox.IsChecked to ViewTemplateItem.IsChecked.
        # Listen to routed CheckBox events so manually clicking a template
        # immediately updates the Compare/Transfer button state.
        self._checkbox_changed_handler = RoutedEventHandler(
            self.UIe_template_checkbox_changed
        )
        self.UI_ListBox_ViewTemplates.AddHandler(
            ToggleButton.CheckedEvent,
            self._checkbox_changed_handler
        )
        self.UI_ListBox_ViewTemplates.AddHandler(
            ToggleButton.UncheckedEvent,
            self._checkbox_changed_handler
        )
        if not self.links:
            forms.alert("No loaded Revit Links were found.", title=WINDOW_TITLE, exitscript=True)
        for link_name in sorted(self.links.keys(), key=lambda value: value.lower()):
            self.UI_LinkSource.Items.Add(link_name)
        if self.UI_LinkSource.Items.Count == 1:
            self.UI_LinkSource.SelectedIndex = 0
        self.ShowDialog()

    def header_drag(self, sender, event_args):
        try:
            self.DragMove()
        except Exception:
            pass

    def button_close(self, sender, event_args):
        self.Close()

    def show_template_interface(self):
        self.UI_Stack_ViewTemplates.Visibility = Visibility.Visible
        self.update_run_button()

    def hide_template_interface(self):
        self.UI_Stack_ViewTemplates.Visibility = Visibility.Collapsed
        self.UI_btn_Run.IsEnabled = False

    def update_run_button(self):
        self.UI_btn_Run.IsEnabled = self.selected_template_count() > 0

    def UIe_ComboBox_Changed(self, sender, event_args):
        selected = self.UI_LinkSource.SelectedItem
        if selected is None:
            self.source_doc = None
            self.all_template_items = []
            self.UI_ListBox_ViewTemplates.ItemsSource = None
            self.hide_template_interface()
            return
        self.source_doc = self.links.get(str(selected))
        if self.source_doc is None:
            self.hide_template_interface()
            forms.alert("The selected Revit Link is unavailable.", title=WINDOW_TITLE)
            return
        self.all_template_items = [ViewTemplateItem(template) for template in collect_view_templates(self.source_doc)]
        self.UI_TextBox_Filter.Text = ""
        self.rebuild_template_list()
        if self.all_template_items:
            self.show_template_interface()
        else:
            self.hide_template_interface()
            forms.alert("No View Templates were found in the selected Revit Link.", title=WINDOW_TITLE)

    def UIe_text_filter_updated(self, sender, event_args):
        self.rebuild_template_list()

    def rebuild_template_list(self):
        if not hasattr(self, "UI_ListBox_ViewTemplates") or not hasattr(self, "UI_TextBox_Filter"):
            return
        keyword = (self.UI_TextBox_Filter.Text or "").strip().lower()
        visible = [item for item in self.all_template_items if not keyword or keyword in item.Name.lower()]
        self.UI_ListBox_ViewTemplates.ItemsSource = None
        self.UI_ListBox_ViewTemplates.ItemsSource = visible
        self.update_run_button()

    def selected_template_count(self):
        return sum(1 for item in self.all_template_items if item.IsChecked)

    def selected_templates(self):
        return [item.element for item in self.all_template_items if item.IsChecked]

    def get_visible_template_items(self):
        source = self.UI_ListBox_ViewTemplates.ItemsSource
        return list(source) if source is not None else []

    def refresh_visible_template_items(self, items):
        self.UI_ListBox_ViewTemplates.ItemsSource = None
        self.UI_ListBox_ViewTemplates.ItemsSource = items
        self.update_run_button()

    def UIe_template_checkbox_changed(self, sender, event_args):
        """Synchronize a clicked template CheckBox and the Run button."""
        try:
            checkbox = event_args.OriginalSource
            template_item = checkbox.DataContext
            if isinstance(template_item, ViewTemplateItem):
                template_item.IsChecked = bool(checkbox.IsChecked)
        except Exception:
            # Routed events from other ToggleButton-derived controls are ignored.
            pass
        self.update_run_button()

    def UIe_btn_select_all(self, sender, event_args):
        items = self.get_visible_template_items()
        for item in items:
            item.IsChecked = True
        self.refresh_visible_template_items(items)

    def UIe_btn_select_none(self, sender, event_args):
        items = self.get_visible_template_items()
        for item in items:
            item.IsChecked = False
        self.refresh_visible_template_items(items)

    def UIe_btn_run(self, sender, event_args):
        if self.source_doc is None:
            forms.alert("Select a source Revit Link first.", title=WINDOW_TITLE)
            return
        selected_templates = self.selected_templates()
        if not selected_templates:
            forms.alert("Select at least one View Template.", title=WINDOW_TITLE)
            return
        self.UI_btn_Run.IsEnabled = False
        try:
            results = run_direct_template_comparison(self.source_doc, selected_templates)
        except Exception:
            self.UI_btn_Run.IsEnabled = True
            forms.alert("Comparison failed. No changes were made.\n\n{}".format(traceback.format_exc()), title=WINDOW_TITLE)
            return
        print_template_comparison_report(self.source_doc, results)
        names_by_status = defaultdict(set)
        for result in results:
            names_by_status[result["status"]].add(result["template_name"])
        new_names = names_by_status["New"]
        changed_names = names_by_status["Changed"]
        unchanged_names = names_by_status["Unchanged"]
        partial_names = names_by_status["Partially Compared"]
        failed_names = names_by_status["Failed"]
        override_existing = bool(self.UI_check_override.IsChecked)
        process_names = set(new_names)
        override_names = set()
        if changed_names and override_existing:
            process_names.update(changed_names)
            override_names.update(changed_names)
        elif changed_names:
            proceed = forms.alert(
                "{} changed templates exist, but Override is disabled.\n\nOnly new templates will be copied. Continue?".format(len(changed_names)),
                title=WINDOW_TITLE, yes=True, no=True)
            if not proceed:
                self.UI_btn_Run.IsEnabled = True
                return
        templates_to_process = [template for template in selected_templates if safe_element_name(template) in process_names]
        if not templates_to_process:
            self.UI_btn_Run.IsEnabled = True
            forms.alert("No new or changed View Templates require transfer.", title=WINDOW_TITLE)
            return
        confirmed = forms.alert(
            "Comparison completed. No project changes have been made.\n\nNew: {}\nChanged to override: {}\nUnchanged: {}\nPartially compared: {}\nFailed: {}\n\nReview the pyRevit report. Continue?".format(
                len(new_names), len(override_names), len(unchanged_names), len(partial_names), len(failed_names)),
            title="Continue View Template Transfer?", yes=True, no=True)
        if not confirmed:
            self.UI_btn_Run.IsEnabled = True
            return
        try:
            transfer_result = apply_template_transfer(self.source_doc, templates_to_process, override_names)
        except Exception:
            self.UI_btn_Run.IsEnabled = True
            forms.alert("Transfer failed and was rolled back.\n\n{}".format(traceback.format_exc()), title=WINDOW_TITLE)
            return
        self.Close()
        copied_ids = transfer_result["copied_ids"]
        try:
            copied_count = copied_ids.Count
        except Exception:
            try:
                copied_count = len(copied_ids)
            except Exception:
                copied_count = 0
        lines = [
            "View Template transfer completed.", "",
            "New templates copied: {}".format(len(new_names)),
            "Changed templates overridden: {}".format(len(override_names)),
            "Unchanged templates skipped: {}".format(len(unchanged_names)),
            "Partially compared templates skipped: {}".format(len(partial_names)),
            "Failed templates skipped: {}".format(len(failed_names)),
            "Revit copied element IDs: {}".format(copied_count)
        ]
        for warning in transfer_result["warnings"]:
            lines.append("Warning: {}".format(warning))
        try:
            uidoc.RefreshActiveView()
        except Exception:
            pass
        forms.alert("\n".join(lines), title=WINDOW_TITLE)


try:
    LoadViewTemplatesWindow()
except Exception:
    forms.alert("The command encountered an unexpected error.\n\n{}".format(traceback.format_exc()), title=WINDOW_TITLE)