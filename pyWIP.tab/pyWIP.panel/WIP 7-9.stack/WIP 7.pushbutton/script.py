# -*- coding: utf-8 -*-

__author__ = "Pankaj Prabhakar"
__version__ = "Version 1.0"

__doc__ = """
Load selected view filters from a loaded Revit link.

Features:
- Native Shift-click and Ctrl-click filter selection
- Search without losing hidden selections
- Dry-run comparison with transaction rollback
- Confirmation before applying changes
- Copy filters missing from the current project
- Update only parameter filters that have changed
- Category comparison
- Existing and linked rule comparison
- Fixed-layout HTML output report
- Four-column report:
    Filter Name
    What Changed
    Existing Rules
    Linked Rules

Required XAML:
LoadViewFiltersFromLink.xaml"""

# ==========================================================================================
# IMPORTS
# ==========================================================================================

import os
import traceback

from Autodesk.Revit.DB import (
    BuiltInParameter,
    Category,
    CopyPasteOptions,
    DuplicateTypeAction,
    ElementId,
    ElementTransformUtils,
    FilteredElementCollector,
    IDuplicateTypeNamesHandler,
    LabelUtils,
    ParameterFilterElement,
    RevitLinkInstance,
    SelectionFilterElement,
    Transaction,
    Transform
)

from pyrevit import forms, script

from System import Enum
from System.Collections.Generic import List
from System.Windows.Controls import CheckBox, ListBoxItem


# ==========================================================================================
# REVIT CONTEXT
# ==========================================================================================

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

WINDOW_TITLE = "Load View Filters From Link"

XAML_FILE = os.path.join(
    os.path.dirname(__file__),
    "LoadViewFiltersFromLink.xaml"
)


# ==========================================================================================
# COPY HANDLER
# ==========================================================================================

class UseDestinationTypesHandler(IDuplicateTypeNamesHandler):
    """Use destination types when duplicate dependency names are found."""

    def OnDuplicateTypeNamesFound(self, args):
        return DuplicateTypeAction.UseDestinationTypes


# ==========================================================================================
# GENERAL HELPERS
# ==========================================================================================

def safe_element_name(element):
    """Return an element name without raising an exception."""

    try:
        return element.Name
    except Exception:
        return "<Unnamed Element>"


def element_id_value(element_id):
    """Return the numeric value of an ElementId."""

    if element_id is None:
        return None

    try:
        return element_id.Value
    except Exception:
        return element_id.IntegerValue


def safe_call(obj, method_name, default=None):
    """Safely call a parameterless method."""

    if obj is None:
        return default

    try:
        method = getattr(obj, method_name, None)
        return method() if method is not None else default
    except Exception:
        return default


def safe_get_property(obj, property_name, default=None):
    """Safely read an object property."""

    if obj is None:
        return default

    try:
        return getattr(obj, property_name)
    except Exception:
        return default


def is_parameter_filter(element):
    return isinstance(element, ParameterFilterElement)


def is_selection_filter(element):
    return isinstance(element, SelectionFilterElement)


def filter_type_name(element):

    if is_parameter_filter(element):
        return "Parameter Filter"

    if is_selection_filter(element):
        return "Selection Filter"

    return "View Filter"


def truncate_text(text, maximum_length=5000):
    """Limit report-cell text length."""

    text = "" if text is None else str(text)

    if len(text) <= maximum_length:
        return text

    return text[:maximum_length] + "..."


def deduplicate_rows(rows):
    """Remove duplicate rows while preserving their order."""

    unique_rows = []
    seen_rows = set()

    for row in rows:

        row_key = tuple(
            str(value)
            for value in row
        )

        if row_key in seen_rows:
            continue

        seen_rows.add(row_key)
        unique_rows.append(row)

    return unique_rows


# ==========================================================================================
# HTML REPORT HELPERS
# ==========================================================================================

def html_escape(value):
    """
    Escape text before inserting it into HTML.

    This prevents rule characters such as <, >, &, quotes, and pipes
    from damaging the table structure.
    """

    if value is None:
        return ""

    text = str(value)

    return (
        text
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&#39;")
    )


def html_cell(value):
    """
    Prepare a value for an HTML table cell.

    Newline characters become HTML line breaks. Pipe characters remain
    ordinary text and do not create extra columns.
    """

    return html_escape(value).replace(
        "\n",
        "<br/>"
    )


def report_status_class(change_text):
    """Return the CSS class for a report row."""

    text = str(change_text).lower()

    if "failed" in text:
        return "vf-failed"

    if "no differences" in text:
        return "vf-unchanged"

    if "new filter" in text:
        return "vf-new"

    if (
        "filter rules changed" in text or
        "categories added" in text or
        "categories removed" in text
    ):
        return "vf-changed"

    if "preserved" in text:
        return "vf-preserved"

    return ""


def add_filter_report_styles(output):
    """Add fixed-layout report styles to the pyRevit output window."""

    css = """
    .vf-report {
        box-sizing: border-box;
        width: 100%;
        margin: 16px 0 26px 0;
        color: #172033;
        font-family: "Segoe UI", Arial, sans-serif;
    }

    .vf-report * {
        box-sizing: border-box;
    }

    .vf-report h2 {
        margin: 0 0 10px 0;
        color: #172033;
        font-family: "Segoe UI", Arial, sans-serif;
        font-size: 18px;
        font-weight: 600;
    }

    .vf-table-wrap {
        box-sizing: border-box;
        display: block;
        width: 100%;
        overflow-x: auto;
        overflow-y: visible;
        border: 1px solid #9ba8b3;
        border-radius: 5px;
        background: #ffffff;
    }

    .vf-table {
        box-sizing: border-box;
        width: 100%;
        min-width: 1200px;
        margin: 0;
        table-layout: fixed;
        border-collapse: collapse;
        border-spacing: 0;
        background: #ffffff;
    }

    .vf-table .vf-col-name {
        width: 16%;
    }

    .vf-table .vf-col-change {
        width: 28%;
    }

    .vf-table .vf-col-existing {
        width: 28%;
    }

    .vf-table .vf-col-linked {
        width: 28%;
    }

    .vf-table thead,
    .vf-table tbody,
    .vf-table tr {
        width: 100%;
    }

    .vf-table th {
        padding: 10px 11px;
        border-right: 1px solid #51687b;
        border-bottom: 1px solid #51687b;
        background: #243c50;
        color: #ffffff;
        text-align: left;
        vertical-align: top;
        font-family: "Segoe UI", Arial, sans-serif;
        font-size: 12px;
        font-weight: 600;
        line-height: 1.35;
        white-space: normal;
        overflow-wrap: break-word;
        word-wrap: break-word;
    }

    .vf-table th:last-child {
        border-right: 0;
    }

    .vf-table td {
        padding: 9px 11px;
        border-right: 1px solid #c4cbd1;
        border-bottom: 1px solid #b5bdc5;
        background: #ffffff;
        color: #172033;
        text-align: left;
        vertical-align: top;
        font-family: Consolas, "Courier New", monospace;
        font-size: 11px;
        font-weight: normal;
        line-height: 1.45;
        white-space: normal;
        overflow-wrap: break-word;
        word-wrap: break-word;
        word-break: normal;
    }

    .vf-table td:last-child {
        border-right: 0;
    }

    .vf-table tbody tr:last-child td {
        border-bottom: 0;
    }

    .vf-table tbody tr:nth-child(even) td {
        background: #f3f5f7;
    }

    .vf-table tbody tr:nth-child(odd) td {
        background: #ffffff;
    }

    .vf-table td.vf-name {
        color: #0f3f66;
        font-family: "Segoe UI", Arial, sans-serif;
        font-weight: 600;
        overflow-wrap: break-word;
        word-wrap: break-word;
    }

    .vf-table td.vf-change {
        font-family: "Segoe UI", Arial, sans-serif;
        font-size: 11px;
    }

    .vf-table td.vf-rules {
        color: #172033;
        font-family: Consolas, "Courier New", monospace;
    }

    .vf-table tr.vf-new td.vf-change {
        color: #145da0;
        font-weight: 600;
    }

    .vf-table tr.vf-changed td.vf-change {
        color: #0b6b37;
        font-weight: 600;
    }

    .vf-table tr.vf-unchanged td.vf-change {
        color: #68737d;
    }

    .vf-table tr.vf-preserved td.vf-change {
        color: #775a00;
    }

    .vf-table tr.vf-failed td {
        background: #fff0f0;
    }

    .vf-table tr.vf-failed td.vf-change {
        color: #a11a1a;
        font-weight: 600;
    }
    """

    try:
        output.add_style(css)

    except Exception:
        output.print_html(
            "<style>{}</style>".format(css)
        )


def print_filter_html_table(
        output,
        rows,
        title,
        change_heading,
        existing_heading,
        linked_heading):
    """
    Print a fixed-layout HTML comparison table.

    Columns:
    1. Filter Name
    2. Proposed Change or What Changed
    3. Existing or Previous Rules
    4. Linked or Current Rules
    """

    if not rows:

        output.print_md(
            "No filter comparison results were generated."
        )

        return

    sorted_rows = sorted(
        rows,
        key=lambda row: str(row[0]).lower()
    )

    html_parts = [
        '<div class="vf-report">',
        '<h2>{}</h2>'.format(
            html_escape(title)
        ),
        '<div class="vf-table-wrap">',
        '<table class="vf-table">',
        '<colgroup>',
        '<col class="vf-col-name"/>',
        '<col class="vf-col-change"/>',
        '<col class="vf-col-existing"/>',
        '<col class="vf-col-linked"/>',
        '</colgroup>',
        '<thead>',
        '<tr>',
        '<th>Filter Name</th>',
        '<th>{}</th>'.format(
            html_escape(change_heading)
        ),
        '<th>{}</th>'.format(
            html_escape(existing_heading)
        ),
        '<th>{}</th>'.format(
            html_escape(linked_heading)
        ),
        '</tr>',
        '</thead>',
        '<tbody>'
    ]

    for row in sorted_rows:

        filter_name = (
            row[0]
            if len(row) > 0
            else ""
        )

        change_text = (
            row[1]
            if len(row) > 1
            else ""
        )

        existing_rules = (
            row[2]
            if len(row) > 2
            else ""
        )

        linked_rules = (
            row[3]
            if len(row) > 3
            else ""
        )

        status_class = report_status_class(
            change_text
        )

        html_parts.extend(
            [
                '<tr class="{}">'.format(
                    status_class
                ),
                '<td class="vf-name">{}</td>'.format(
                    html_cell(filter_name)
                ),
                '<td class="vf-change">{}</td>'.format(
                    html_cell(change_text)
                ),
                '<td class="vf-rules">{}</td>'.format(
                    html_cell(existing_rules)
                ),
                '<td class="vf-rules">{}</td>'.format(
                    html_cell(linked_rules)
                ),
                '</tr>'
            ]
        )

    html_parts.extend(
        [
            '</tbody>',
            '</table>',
            '</div>',
            '</div>'
        ]
    )

    output.print_html(
        "".join(html_parts)
    )


# ==========================================================================================
# LOADED REVIT LINKS
# ==========================================================================================

def collect_loaded_links():
    """Collect all loaded Revit link documents."""

    links = {}

    link_instances = (
        FilteredElementCollector(doc)
        .OfClass(RevitLinkInstance)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    for link_instance in link_instances:

        try:
            link_document = (
                link_instance.GetLinkDocument()
            )

            if link_document is None:
                continue

            display_name = link_instance.Name

            if display_name in links:

                display_name = "{} | Instance ID {}".format(
                    display_name,
                    element_id_value(link_instance.Id)
                )

            links[display_name] = link_document

        except Exception:
            continue

    return links


# ==========================================================================================
# FILTER COLLECTION
# ==========================================================================================

def collect_filters_by_class(
        source_document,
        filter_class):

    try:
        return list(
            FilteredElementCollector(source_document)
            .OfClass(filter_class)
            .WhereElementIsNotElementType()
            .ToElements()
        )

    except Exception:
        return []


def collect_all_filters(source_document):
    """Collect parameter and selection filters from the linked document."""

    source_filters = []

    source_filters.extend(
        collect_filters_by_class(
            source_document,
            ParameterFilterElement
        )
    )

    source_filters.extend(
        collect_filters_by_class(
            source_document,
            SelectionFilterElement
        )
    )

    return sorted(
        source_filters,
        key=lambda element: (
            safe_element_name(element).lower()
        )
    )


# ==========================================================================================
# EXISTING FILTER LOOKUP
# ==========================================================================================

def find_existing_filter(
        filter_class,
        filter_name):
    """Find a host filter having the same class and name."""

    try:
        host_filters = (
            FilteredElementCollector(doc)
            .OfClass(filter_class)
            .WhereElementIsNotElementType()
            .ToElements()
        )

        for host_filter in host_filters:

            if (
                safe_element_name(host_filter) ==
                filter_name
            ):
                return host_filter

    except Exception:
        pass

    return None


# ==========================================================================================
# CATEGORY COMPARISON
# ==========================================================================================

def category_id_set(parameter_filter):

    return set(
        element_id_value(category_id)
        for category_id
        in parameter_filter.GetCategories()
    )


def category_name(category_id):

    try:
        category = Category.GetCategory(
            doc,
            category_id
        )

        if category is not None:
            return category.Name

    except Exception:
        pass

    return "Category ID {}".format(
        element_id_value(category_id)
    )


def compare_filter_categories(
        existing_filter,
        translated_filter):
    """
    Compare the categories admitted by two host-document filters.
    """

    existing_values = category_id_set(
        existing_filter
    )

    translated_values = category_id_set(
        translated_filter
    )

    if existing_values == translated_values:
        return False, []

    existing_ids = (
        existing_filter.GetCategories()
    )

    translated_ids = (
        translated_filter.GetCategories()
    )

    added_categories = sorted(
        category_name(category_id)
        for category_id in translated_ids
        if (
            element_id_value(category_id)
            not in existing_values
        )
    )

    removed_categories = sorted(
        category_name(category_id)
        for category_id in existing_ids
        if (
            element_id_value(category_id)
            not in translated_values
        )
    )

    changes = []

    if added_categories:

        changes.append(
            "Categories added: {}".format(
                ", ".join(added_categories)
            )
        )

    if removed_categories:

        changes.append(
            "Categories removed: {}".format(
                ", ".join(removed_categories)
            )
        )

    return True, changes


# ==========================================================================================
# PARAMETER AND ELEMENT NAMES
# ==========================================================================================

def parameter_name(parameter_id):
    """Resolve project, shared, and built-in parameter names."""

    if parameter_id is None:
        return "<Unknown Parameter>"

    try:
        parameter_element = doc.GetElement(
            parameter_id
        )

        if parameter_element is not None:

            name = safe_element_name(
                parameter_element
            )

            if name:
                return name

    except Exception:
        pass

    try:
        built_in_parameter = Enum.ToObject(
            BuiltInParameter,
            element_id_value(parameter_id)
        )

        label = LabelUtils.GetLabelFor(
            built_in_parameter
        )

        if label:
            return label

    except Exception:
        pass

    return "Parameter ID {}".format(
        element_id_value(parameter_id)
    )


def element_name_from_id(element_id):
    """Resolve an ElementId rule value to a readable element name."""

    if element_id is None:
        return "<None>"

    try:
        element = doc.GetElement(element_id)

        if element is not None:
            return safe_element_name(element)

    except Exception:
        pass

    return "Element ID {}".format(
        element_id_value(element_id)
    )


# ==========================================================================================
# RULE VALUE EXTRACTION
# ==========================================================================================

def format_rule_value(value):
    """Convert a rule value into readable report text."""

    if value is None:
        return "<No Value>"

    if isinstance(value, ElementId):
        return element_name_from_id(value)

    if isinstance(value, bool):
        return "True" if value else "False"

    try:
        if isinstance(value, float):
            return "{:.6g}".format(value)

    except Exception:
        pass

    try:
        return str(value)

    except Exception:
        return "<Unreadable Value>"


def get_rule_parameter_id(rule):

    parameter_id = safe_call(
        rule,
        "GetRuleParameter"
    )

    if parameter_id is not None:
        return parameter_id

    return safe_get_property(
        rule,
        "RuleParameter"
    )


def get_rule_value(rule):
    """Extract a value from different FilterRule subclasses."""

    property_names = (
        "RuleString",
        "RuleValue",
        "Value",
        "StringValue",
        "IntegerValue",
        "DoubleValue",
        "ElementIdValue"
    )

    for property_name in property_names:

        value = safe_get_property(
            rule,
            property_name
        )

        if value is not None:
            return value

    method_names = (
        "GetRuleString",
        "GetRuleValue",
        "GetValue",
        "GetGlobalParameterId"
    )

    for method_name in method_names:

        value = safe_call(
            rule,
            method_name
        )

        if value is not None:
            return value

    return None


def get_rule_tolerance(rule):

    for property_name in (
        "Epsilon",
        "Tolerance"
    ):

        value = safe_get_property(
            rule,
            property_name
        )

        if value is not None:
            return value

    for method_name in (
        "GetTolerance",
        "GetEpsilon"
    ):

        value = safe_call(
            rule,
            method_name
        )

        if value is not None:
            return value

    return None


# ==========================================================================================
# RULE OPERATOR
# ==========================================================================================

def get_rule_operator(rule):
    """Return a readable rule operator."""

    evaluator = safe_call(
        rule,
        "GetEvaluator"
    )

    if evaluator is None:

        evaluator = safe_get_property(
            rule,
            "RuleEvaluator"
        )

    if evaluator is not None:

        try:
            evaluator_name = (
                evaluator.GetType().Name
            )

        except Exception:
            evaluator_name = str(evaluator)

    else:

        try:
            evaluator_name = (
                rule.GetType().Name
            )

        except Exception:
            evaluator_name = "Unknown Rule"

    operator_names = {
        "FilterStringEquals":
            "equals",

        "FilterStringNotEquals":
            "does not equal",

        "FilterStringContains":
            "contains",

        "FilterStringNotContains":
            "does not contain",

        "FilterStringBeginsWith":
            "begins with",

        "FilterStringNotBeginsWith":
            "does not begin with",

        "FilterStringEndsWith":
            "ends with",

        "FilterStringNotEndsWith":
            "does not end with",

        "FilterNumericEquals":
            "equals",

        "FilterNumericNotEquals":
            "does not equal",

        "FilterNumericGreater":
            "is greater than",

        "FilterNumericGreaterOrEqual":
            "is greater than or equal to",

        "FilterNumericLess":
            "is less than",

        "FilterNumericLessOrEqual":
            "is less than or equal to",

        "FilterElementIdEquals":
            "equals",

        "FilterElementIdNotEquals":
            "does not equal",

        "FilterHasValue":
            "has a value",

        "FilterHasNoValue":
            "has no value",

        "HasValueFilterRule":
            "has a value",

        "HasNoValueFilterRule":
            "has no value",

        "FilterGlobalParameterAssociation":
            "is associated with",

        "FilterGlobalParameterAssociationRule":
            "is associated with"
    }

    return operator_names.get(
        evaluator_name,
        evaluator_name
    )


# ==========================================================================================
# RULE DESCRIPTIONS
# ==========================================================================================

def describe_filter_rule(rule):
    """Describe one FilterRule using separate readable lines."""

    if rule is None:
        return "<Null Rule>"

    try:
        rule_type = rule.GetType().Name

    except Exception:
        rule_type = "Unknown Rule"

    if rule_type == "FilterInverseRule":

        inner_rule = safe_call(
            rule,
            "GetInnerRule"
        )

        if inner_rule is None:

            inner_rule = safe_get_property(
                rule,
                "InnerRule"
            )

        if inner_rule is not None:

            return "NOT ({})".format(
                describe_filter_rule(inner_rule)
            )

    rule_parameter_name = parameter_name(
        get_rule_parameter_id(rule)
    )

    operator = get_rule_operator(rule)
    value = get_rule_value(rule)

    if (
        value is None and
        operator in (
            "has a value",
            "has no value"
        )
    ):

        description = (
            "Parameter: {}\n"
            "Operator: {}"
        ).format(
            rule_parameter_name,
            operator
        )

    else:

        description = (
            "Parameter: {}\n"
            "Operator: {}\n"
            "Value: {}"
        ).format(
            rule_parameter_name,
            operator,
            format_rule_value(value)
        )

    tolerance = get_rule_tolerance(rule)

    if tolerance is not None:

        try:
            tolerance_text = (
                "{:.6g}".format(tolerance)
            )

        except Exception:
            tolerance_text = str(tolerance)

        description += (
            "\nTolerance: {}"
            .format(tolerance_text)
        )

    return description


def describe_element_filter(element_filter):
    """Recursively describe a Revit ElementFilter."""

    if element_filter is None:
        return ["No rules"]

    try:
        filter_type = (
            element_filter.GetType().Name
        )

    except Exception:
        return ["Unknown ElementFilter"]

    if filter_type == "ElementParameterFilter":

        rules = safe_call(
            element_filter,
            "GetRules",
            []
        )

        descriptions = [
            describe_filter_rule(rule)
            for rule in rules
        ]

        if not descriptions:
            descriptions = ["No rules"]

        if safe_get_property(
                element_filter,
                "Inverted",
                False):

            descriptions = [
                "NOT ({})".format(description)
                for description in descriptions
            ]

        return descriptions

    if filter_type in (
        "LogicalAndFilter",
        "LogicalOrFilter"
    ):

        group_operator = (
            "AND"
            if filter_type == "LogicalAndFilter"
            else "OR"
        )

        child_descriptions = []

        for child_filter in safe_call(
                element_filter,
                "GetFilters",
                []):

            child_rules = describe_element_filter(
                child_filter
            )

            child_descriptions.append(
                "\n{}\n".format(
                    group_operator
                ).join(child_rules)
            )

        if child_descriptions:
            return child_descriptions

        return ["No rules"]

    return [filter_type]


def normalized_rule_descriptions(element_filter):
    """Normalize descriptions for comparison and reporting."""

    return sorted(
        str(description).strip()
        for description
        in describe_element_filter(element_filter)
        if description
    )


def format_rule_list(descriptions):
    """Format each rule as a separate block."""

    if not descriptions:
        return "No rules"

    formatted_rules = []

    for index, description in enumerate(
            descriptions,
            start=1):

        formatted_rules.append(
            "Rule {}\n{}".format(
                index,
                description
            )
        )

    return truncate_text(
        "\n\n".join(formatted_rules)
    )


# ==========================================================================================
# FILTER COPY
# ==========================================================================================

def copy_filter_element(
        source_document,
        source_filter):
    """Copy one filter and return the copied filter element."""

    source_class = source_filter.GetType()

    source_name = safe_element_name(
        source_filter
    )

    source_ids = List[ElementId]()
    source_ids.Add(source_filter.Id)

    copy_options = CopyPasteOptions()

    copy_options.SetDuplicateTypeNamesHandler(
        UseDestinationTypesHandler()
    )

    copied_ids = ElementTransformUtils.CopyElements(
        source_document,
        source_ids,
        doc,
        Transform.Identity,
        copy_options
    )

    candidates = []

    for copied_id in copied_ids:

        copied_element = doc.GetElement(
            copied_id
        )

        if copied_element is None:
            continue

        try:
            if (
                copied_element.GetType() ==
                source_class
            ):

                candidates.append(
                    copied_element
                )

        except Exception:
            continue

    for candidate in candidates:

        if (
            safe_element_name(candidate) ==
            source_name
        ):

            return candidate

    if candidates:
        return candidates[0]

    raise Exception(
        "Revit did not return a copied filter for '{}'."
        .format(source_name)
    )


def copy_new_filter(
        source_document,
        source_filter):
    """Copy a filter missing from the current project."""

    source_name = safe_element_name(
        source_filter
    )

    copied_filter = copy_filter_element(
        source_document,
        source_filter
    )

    if (
        safe_element_name(copied_filter) !=
        source_name
    ):

        copied_filter.Name = source_name

    return copied_filter


# ==========================================================================================
# PARAMETER FILTER COMPARISON
# ==========================================================================================

def compare_and_update_parameter_filter(
        source_document,
        source_filter,
        existing_filter):
    """
    Compare and conditionally update an existing parameter filter.

    Returns:
        destination filter
        changed
        change descriptions
        previous rules
        linked rules
    """

    temporary_filter = None
    changes = []

    try:
        temporary_filter = copy_filter_element(
            source_document,
            source_filter
        )

        if temporary_filter is None:

            raise Exception(
                "A temporary host-compatible filter "
                "was not created."
            )

        if (
            temporary_filter.Id ==
            existing_filter.Id
        ):

            raise Exception(
                "Revit returned the existing filter instead "
                "of creating a temporary filter."
            )

        if not isinstance(
                temporary_filter,
                ParameterFilterElement):

            raise Exception(
                "The temporary element is not a "
                "ParameterFilterElement."
            )

        existing_element_filter = (
            existing_filter.GetElementFilter()
        )

        linked_element_filter = (
            temporary_filter.GetElementFilter()
        )

        existing_rules = (
            normalized_rule_descriptions(
                existing_element_filter
            )
        )

        linked_rules = (
            normalized_rule_descriptions(
                linked_element_filter
            )
        )

        (
            categories_changed,
            category_changes
        ) = compare_filter_categories(
            existing_filter,
            temporary_filter
        )

        if categories_changed:

            existing_filter.SetCategories(
                temporary_filter.GetCategories()
            )

            changes.extend(
                category_changes
            )

        rules_changed = False

        if linked_element_filter is None:

            if existing_element_filter is not None:

                existing_filter.ClearRules()
                rules_changed = True

        else:

            rules_changed = (
                existing_filter.SetElementFilter(
                    linked_element_filter
                )
            )

        if rules_changed:

            changes.append(
                "Filter rules changed"
            )

        elif existing_rules != linked_rules:

            changes.append(
                "Rule structure differs, but Revit considers "
                "the rules logically equivalent"
            )

        doc.Delete(
            temporary_filter.Id
        )

        temporary_filter = None

        return (
            existing_filter,
            categories_changed or rules_changed,
            changes,
            existing_rules,
            linked_rules
        )

    except Exception:

        if temporary_filter is not None:

            try:
                if (
                    temporary_filter.Id !=
                    existing_filter.Id
                ):

                    doc.Delete(
                        temporary_filter.Id
                    )

            except Exception:
                pass

        raise


def copy_or_update_filter(
        source_document,
        source_filter):
    """Copy a new filter or compare an existing filter."""

    source_name = safe_element_name(
        source_filter
    )

    source_class = source_filter.GetType()

    existing_filter = find_existing_filter(
        source_class,
        source_name
    )

    if existing_filter is None:

        copied_filter = copy_new_filter(
            source_document,
            source_filter
        )

        linked_rules = [
            "Not applicable"
        ]

        if is_parameter_filter(copied_filter):

            linked_rules = (
                normalized_rule_descriptions(
                    copied_filter.GetElementFilter()
                )
            )

        return (
            copied_filter,
            "copied",
            ["New filter copied"],
            ["Not applicable"],
            linked_rules
        )

    if is_parameter_filter(source_filter):

        (
            destination_filter,
            changed,
            changes,
            existing_rules,
            linked_rules
        ) = compare_and_update_parameter_filter(
            source_document,
            source_filter,
            existing_filter
        )

        if changed:

            return (
                destination_filter,
                "updated",
                changes,
                existing_rules,
                linked_rules
            )

        return (
            destination_filter,
            "unchanged",
            ["No differences detected"],
            existing_rules,
            linked_rules
        )

    if is_selection_filter(source_filter):

        return (
            existing_filter,
            "selection_preserved",
            [
                "Selection filter preserved. Linked element IDs "
                "cannot be mapped automatically."
            ],
            ["Not applicable"],
            ["Not applicable"]
        )

    raise Exception(
        "Unsupported filter type: {}".format(
            source_class.Name
        )
    )


# ==========================================================================================
# FILTER PROCESSING
# ==========================================================================================

def empty_result():

    return {
        "copied": [],
        "updated": [],
        "unchanged": [],
        "selection_preserved": [],
        "failed": [],
        "report": []
    }


def process_filters(
        source_document,
        selected_filters,
        dry_run=False):
    """Process filters inside the caller's active transaction."""

    result = empty_result()

    for source_filter in selected_filters:

        filter_name = safe_element_name(
            source_filter
        )

        try:
            (
                destination_filter,
                status,
                changes,
                existing_rules,
                linked_rules
            ) = copy_or_update_filter(
                source_document,
                source_filter
            )

            result[status].append(
                filter_name
            )

            if status == "copied":

                changes = [
                    (
                        "New filter will be copied"
                        if dry_run
                        else "New filter copied"
                    )
                ]

            result["report"].append(
                [
                    filter_name,
                    truncate_text(
                        "; ".join(changes),
                        2500
                    ),
                    format_rule_list(
                        existing_rules
                    ),
                    format_rule_list(
                        linked_rules
                    )
                ]
            )

        except Exception:

            result["failed"].append(
                {
                    "name": filter_name,
                    "type": filter_type_name(
                        source_filter
                    ),
                    "error": traceback.format_exc()
                }
            )

            result["report"].append(
                [
                    filter_name,
                    "Failed to compare or process",
                    "Unavailable",
                    "Unavailable"
                ]
            )

    result["report"] = deduplicate_rows(
        result["report"]
    )

    return result


def run_dry_run(
        source_document,
        selected_filters):
    """Run the comparison and roll back all model changes."""

    transaction = Transaction(
        doc,
        "Dry Run - Compare View Filters From Link"
    )

    try:
        transaction.Start()

        result = process_filters(
            source_document,
            selected_filters,
            dry_run=True
        )

        transaction.RollBack()

        return result

    except Exception:

        try:
            transaction.RollBack()
        except Exception:
            pass

        raise


def apply_changes(
        source_document,
        selected_filters):
    """Apply and commit approved filter changes."""

    transaction = Transaction(
        doc,
        "Load Selected View Filters From Link"
    )

    try:
        transaction.Start()

        result = process_filters(
            source_document,
            selected_filters,
            dry_run=False
        )

        transaction.Commit()

        return result

    except Exception:

        try:
            transaction.RollBack()
        except Exception:
            pass

        raise


# ==========================================================================================
# UI DATA
# ==========================================================================================

class FilterItem(object):

    def __init__(self, element):

        self.element = element
        self.name = safe_element_name(
            element
        )

        self.is_selected = False
        self.checkbox = None
        self.listbox_item = None


# ==========================================================================================
# WPF WINDOW
# ==========================================================================================

class LoadFiltersWindow(forms.WPFWindow):

    def __init__(self):

        if not os.path.exists(XAML_FILE):

            forms.alert(
                "The UI file was not found."
                "\n\n{}".format(XAML_FILE),
                title=WINDOW_TITLE,
                exitscript=True
            )

        forms.WPFWindow.__init__(
            self,
            XAML_FILE
        )

        self.links = collect_loaded_links()
        self.source_doc = None
        self.all_filter_items = []
        self.is_syncing_selection = False

        if not self.links:

            forms.alert(
                "No loaded Revit links were found."
                "\n\nOpen Manage Links and make sure at least "
                "one Revit link is loaded.",
                title=WINDOW_TITLE,
                exitscript=True
            )

        for link_name in sorted(
                self.links.keys()):

            self.combo_links.Items.Add(
                link_name
            )

        if self.combo_links.Items.Count == 1:

            self.combo_links.SelectedIndex = 0

        self.ShowDialog()


    # ======================================================================================
    # WINDOW EVENTS
    # ======================================================================================

    def header_drag(
            self,
            sender,
            event_args):

        try:
            self.DragMove()
        except Exception:
            pass


    def button_close(
            self,
            sender,
            event_args):

        self.Close()


    # ======================================================================================
    # LINK AND SEARCH
    # ======================================================================================

    def link_selection_changed(
            self,
            sender,
            event_args):

        selected_link = (
            self.combo_links.SelectedItem
        )

        if selected_link is None:
            return

        self.source_doc = self.links[
            str(selected_link)
        ]

        self.all_filter_items = [
            FilterItem(filter_element)
            for filter_element
            in collect_all_filters(
                self.source_doc
            )
        ]

        self.textbox_filter.Text = ""

        self.rebuild_filter_list()


    def filter_text_changed(
            self,
            sender,
            event_args):

        self.rebuild_filter_list()


    # ======================================================================================
    # FILTER LIST
    # ======================================================================================

    def rebuild_filter_list(self):
        """
        Rebuild the visible ListBox while preserving selections hidden
        by the current search.
        """

        if not hasattr(
                self,
                "filters_list"):

            return

        self.is_syncing_selection = True

        try:
            self.filters_list.Items.Clear()

            keyword = (
                self.textbox_filter.Text
                .strip()
                .lower()
            )

            visible_count = 0

            for item in self.all_filter_items:

                item.checkbox = None
                item.listbox_item = None

                if (
                    keyword and
                    keyword not in item.name.lower()
                ):
                    continue

                checkbox = CheckBox()
                checkbox.Content = item.name
                checkbox.Tag = item
                checkbox.IsChecked = (
                    item.is_selected
                )

                # The ListBoxItem handles normal, Ctrl, and Shift clicks.
                checkbox.IsHitTestVisible = False
                checkbox.Focusable = False

                try:
                    checkbox.Style = self.FindResource(
                        "FilterCheckBoxStyle"
                    )

                except Exception:
                    pass

                listbox_item = ListBoxItem()
                listbox_item.Content = checkbox
                listbox_item.Tag = item
                listbox_item.IsSelected = (
                    item.is_selected
                )

                item.checkbox = checkbox
                item.listbox_item = listbox_item

                self.filters_list.Items.Add(
                    listbox_item
                )

                visible_count += 1

        finally:
            self.is_syncing_selection = False

        self.update_filter_count(
            visible_count
        )


    def filters_selection_changed(
            self,
            sender,
            event_args):
        """
        Synchronize Shift, Ctrl, keyboard, and ordinary ListBox
        selections with the filter objects and checkbox indicators.
        """

        if self.is_syncing_selection:
            return

        self.is_syncing_selection = True

        try:
            for listbox_item in event_args.AddedItems:

                filter_item = listbox_item.Tag

                if filter_item is None:
                    continue

                filter_item.is_selected = True

                if filter_item.checkbox is not None:

                    filter_item.checkbox.IsChecked = True

            for listbox_item in event_args.RemovedItems:

                filter_item = listbox_item.Tag

                if filter_item is None:
                    continue

                filter_item.is_selected = False

                if filter_item.checkbox is not None:

                    filter_item.checkbox.IsChecked = False

        finally:
            self.is_syncing_selection = False

        self.update_filter_count(
            self.filters_list.Items.Count
        )


    def selected_filter_count(self):

        return sum(
            1
            for item in self.all_filter_items
            if item.is_selected
        )


    def selected_filters(self):

        return [
            item.element
            for item in self.all_filter_items
            if item.is_selected
        ]


    def update_filter_count(
            self,
            visible_count):

        selected_count = (
            self.selected_filter_count()
        )

        self.filter_count_text.Text = (
            "{} shown | {} total | {} selected"
            .format(
                visible_count,
                len(self.all_filter_items),
                selected_count
            )
        )

        self.button_load.IsEnabled = (
            selected_count > 0
        )


    # ======================================================================================
    # SELECT ALL AND NONE
    # ======================================================================================

    def button_select_all(
            self,
            sender,
            event_args):

        self.is_syncing_selection = True

        try:
            self.filters_list.SelectAll()

            for listbox_item in self.filters_list.Items:

                filter_item = listbox_item.Tag

                if filter_item is None:
                    continue

                filter_item.is_selected = True
                listbox_item.IsSelected = True

                if filter_item.checkbox is not None:

                    filter_item.checkbox.IsChecked = True

        finally:
            self.is_syncing_selection = False

        self.update_filter_count(
            self.filters_list.Items.Count
        )


    def button_select_none(
            self,
            sender,
            event_args):

        self.is_syncing_selection = True

        try:
            self.filters_list.UnselectAll()

            for listbox_item in self.filters_list.Items:

                filter_item = listbox_item.Tag

                if filter_item is None:
                    continue

                filter_item.is_selected = False
                listbox_item.IsSelected = False

                if filter_item.checkbox is not None:

                    filter_item.checkbox.IsChecked = False

        finally:
            self.is_syncing_selection = False

        self.update_filter_count(
            self.filters_list.Items.Count
        )


    # ======================================================================================
    # OUTPUT REPORT
    # ======================================================================================

    def print_report(
            self,
            result,
            dry_run):
        """
        Print the summary and fixed-layout HTML comparison table.
        """

        output = script.get_output()

        report_title = (
            "View Filter Dry-Run Report"
            if dry_run
            else "View Filter Change Report"
        )

        try:
            output.set_title(
                report_title
            )
        except Exception:
            pass

        output.print_md(
            "# {}".format(report_title)
        )

        if dry_run:

            output.print_md(
                "**No model changes have been committed. "
                "The dry-run transaction was rolled back.**"
            )

        else:

            output.print_md(
                "**Approved changes have been committed "
                "to the current project.**"
            )

        summary_data = [
            [
                "Source link",
                self.source_doc.Title
            ],
            [
                "New filters",
                str(len(result["copied"]))
            ],
            [
                "Changed filters",
                str(len(result["updated"]))
            ],
            [
                "Unchanged filters",
                str(len(result["unchanged"]))
            ],
            [
                "Selection filters preserved",
                str(
                    len(
                        result[
                            "selection_preserved"
                        ]
                    )
                )
            ],
            [
                "Failures",
                str(len(result["failed"]))
            ]
        ]

        # Summary data does not contain rule pipes, so the standard
        # pyRevit table remains safe here.
        output.print_table(
            table_data=summary_data,
            columns=[
                "Item",
                "Result"
            ],
            title="Summary"
        )

        add_filter_report_styles(
            output
        )

        print_filter_html_table(
            output=output,
            rows=result["report"],
            title=(
                "Proposed Filter Changes"
                if dry_run
                else "Final Filter Change Report"
            ),
            change_heading=(
                "Proposed Change"
                if dry_run
                else "What Changed"
            ),
            existing_heading=(
                "Existing Rules"
                if dry_run
                else "Previous Rules"
            ),
            linked_heading=(
                "Linked Rules"
                if dry_run
                else "Current Rules"
            )
        )

        if result["failed"]:

            output.print_md(
                "## Failure Details"
            )

            for failure in result["failed"]:

                output.print_md(
                    "### {}: {}".format(
                        failure["type"],
                        failure["name"]
                    )
                )

                output.print_md(
                    "```text\n{}\n```".format(
                        failure["error"]
                    )
                )


    # ======================================================================================
    # PROCESS BUTTON
    # ======================================================================================

    def button_run(
            self,
            sender,
            event_args):

        if self.source_doc is None:

            forms.alert(
                "Select a source Revit link first.",
                title=WINDOW_TITLE
            )

            return

        selected_filters = (
            self.selected_filters()
        )

        if not selected_filters:

            forms.alert(
                "Select at least one filter.",
                title=WINDOW_TITLE
            )

            return

        self.button_load.IsEnabled = False

        # ------------------------------------------------------------------
        # Dry run
        # ------------------------------------------------------------------

        try:
            dry_result = run_dry_run(
                self.source_doc,
                selected_filters
            )

        except Exception:

            self.button_load.IsEnabled = True

            forms.alert(
                "The dry run failed."
                "\n\nNo changes were applied."
                "\n\n{}".format(
                    traceback.format_exc()
                ),
                title=WINDOW_TITLE
            )

            return

        self.print_report(
            dry_result,
            True
        )

        proposed_count = (
            len(dry_result["copied"]) +
            len(dry_result["updated"])
        )

        if proposed_count == 0:

            self.button_load.IsEnabled = True

            forms.alert(
                "Dry run completed."
                "\n\nNo new or changed filters were detected."
                "\n\nUnchanged filters: {}"
                "\nSelection filters preserved: {}"
                "\nFailed comparisons: {}"
                .format(
                    len(dry_result["unchanged"]),
                    len(
                        dry_result[
                            "selection_preserved"
                        ]
                    ),
                    len(dry_result["failed"])
                ),
                title="Dry Run Complete"
            )

            return

        # ------------------------------------------------------------------
        # User confirmation
        # ------------------------------------------------------------------

        confirmed = forms.alert(
            "Dry run completed."
            "\n\nNo model changes have been made."
            "\n\nNew filters to copy: {}"
            "\nExisting filters to update: {}"
            "\nUnchanged filters: {}"
            "\nSelection filters preserved: {}"
            "\nFailed comparisons: {}"
            "\n\nReview the pyRevit output report."
            "\n\nApply the proposed changes?"
            .format(
                len(dry_result["copied"]),
                len(dry_result["updated"]),
                len(dry_result["unchanged"]),
                len(
                    dry_result[
                        "selection_preserved"
                    ]
                ),
                len(dry_result["failed"])
            ),
            title="Apply View Filter Changes?",
            yes=True,
            no=True
        )

        if not confirmed:

            self.button_load.IsEnabled = True

            return

        # ------------------------------------------------------------------
        # Apply changes
        # ------------------------------------------------------------------

        try:
            final_result = apply_changes(
                self.source_doc,
                selected_filters
            )

        except Exception:

            self.button_load.IsEnabled = True

            forms.alert(
                "The apply operation failed."
                "\n\nThe transaction was rolled back."
                "\n\n{}".format(
                    traceback.format_exc()
                ),
                title=WINDOW_TITLE
            )

            return

        self.Close()

        self.print_report(
            final_result,
            False
        )

        forms.alert(
            "View filter processing completed."
            "\n\nNew filters copied: {}"
            "\nChanged filters updated: {}"
            "\nUnchanged filters: {}"
            "\nSelection filters preserved: {}"
            "\nFailed filters: {}"
            .format(
                len(final_result["copied"]),
                len(final_result["updated"]),
                len(final_result["unchanged"]),
                len(
                    final_result[
                        "selection_preserved"
                    ]
                ),
                len(final_result["failed"])
            ),
            title=WINDOW_TITLE
        )


# ==========================================================================================
# ENTRY POINT
# ==========================================================================================

try:
    LoadFiltersWindow()

except Exception:

    forms.alert(
        "The command encountered an unexpected error."
        "\n\n{}".format(
            traceback.format_exc()
        ),
        title=WINDOW_TITLE
    )