# -*- coding: utf-8 -*-

__author__ = "Pankaj Prabhakar"
__version__ = "Version 1.0"

__doc__ = """
Load View Templates From Link

Copy selected View Templates from a loaded Revit Link into the
current Revit project.

Workflow:
1. Select a loaded Revit Link.
2. Select one or more View Templates from the linked project.
3. Choose whether matching View Templates should be overridden.
4. Transfer the selected View Templates into the current project.

When Override Existing View Templates is enabled:
- Existing templates with matching names are replaced.
- Views using the existing templates are recorded.
- The replacement templates are reassigned to those views.

This script uses a standalone WPF XAML file.
No EF Tools Python libraries are required.
"""


# ==================================================
# IMPORTS
# ==================================================

import os
import traceback
from collections import defaultdict

from Autodesk.Revit.DB import (
    CopyPasteOptions,
    DuplicateTypeAction,
    ElementId,
    ElementTransformUtils,
    FilteredElementCollector,
    IDuplicateTypeNamesHandler,
    RevitLinkInstance,
    Transaction,
    TransactionGroup,
    Transform,
    View
)

from pyrevit import forms

from System.Collections.Generic import List
from System.Windows import Visibility


# ==================================================
# REVIT VARIABLES
# ==================================================

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

WINDOW_TITLE = "Load View Templates From Link"

XAML_FILE = os.path.join(
    os.path.dirname(__file__),
    "CopyViewTemplate.xaml"
)


# ==================================================
# DUPLICATE TYPE HANDLER
# ==================================================

class UseDestinationTypesHandler(IDuplicateTypeNamesHandler):
    """Use existing destination types when duplicate types are found."""

    def OnDuplicateTypeNamesFound(self, args):
        return DuplicateTypeAction.UseDestinationTypes


# Keep a module-level reference to prevent garbage collection while
# Revit is executing ElementTransformUtils.CopyElements.
DUPLICATE_TYPE_HANDLER = UseDestinationTypesHandler()


# ==================================================
# GENERAL HELPERS
# ==================================================

def safe_element_name(element):
    """Return the element name without raising an exception."""

    try:
        return element.Name
    except Exception:
        return "<Unnamed View Template>"


def element_id_value(element_id):
    """
    Return the numerical value of an ElementId.

    Revit 2024 and earlier normally expose IntegerValue.
    Newer versions may expose Value.
    """

    if element_id is None:
        return None

    try:
        return element_id.Value
    except Exception:
        pass

    try:
        return element_id.IntegerValue
    except Exception:
        pass

    return str(element_id)


def is_invalid_element_id(element_id):
    """Return True when an ElementId is invalid."""

    if element_id is None:
        return True

    try:
        return element_id == ElementId.InvalidElementId
    except Exception:
        return element_id_value(element_id) == -1


# ==================================================
# REVIT LINK COLLECTION
# ==================================================

def collect_loaded_links():
    """
    Collect loaded Revit Links in the current project.

    Returns:
        {
            "Link display name": linked_document
        }
    """

    result = {}
    collected_instances = []

    link_instances = (
        FilteredElementCollector(doc)
        .OfClass(RevitLinkInstance)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    for link_instance in link_instances:

        try:
            link_doc = link_instance.GetLinkDocument()
        except Exception:
            link_doc = None

        # GetLinkDocument returns None when the link is unloaded.
        if link_doc is None:
            continue

        try:
            if link_doc.IsFamilyDocument:
                continue
        except Exception:
            continue

        try:
            base_name = link_instance.Name
        except Exception:
            base_name = link_doc.Title

        collected_instances.append(
            (
                base_name,
                link_instance,
                link_doc
            )
        )

    name_counts = defaultdict(int)

    for base_name, link_instance, link_doc in collected_instances:
        name_counts[base_name] += 1

    for base_name, link_instance, link_doc in collected_instances:

        if name_counts[base_name] > 1:

            display_name = "{} | Instance ID {}".format(
                base_name,
                element_id_value(link_instance.Id)
            )

        else:
            display_name = base_name

        result[display_name] = link_doc

    return result


# ==================================================
# VIEW TEMPLATE COLLECTION
# ==================================================

def collect_view_templates(source_doc):
    """Collect View Templates from the supplied document."""

    templates = []

    views = (
        FilteredElementCollector(source_doc)
        .OfClass(View)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    for view in views:

        try:
            if view.IsTemplate:
                templates.append(view)
        except Exception:
            continue

    return sorted(
        templates,
        key=lambda item: safe_element_name(item).lower()
    )


def collect_destination_templates():
    """Collect View Templates from the current project."""

    return collect_view_templates(doc)


def find_destination_template(template_name):
    """
    Find a View Template in the current project by exact name.
    """

    for view_template in collect_destination_templates():

        if safe_element_name(view_template) == template_name:
            return view_template

    return None


# ==================================================
# VIEW TEMPLATE LIST ITEM
# ==================================================

class ViewTemplateItem(object):
    """
    Wrapper used by the XAML ListBox.

    The following properties correspond to the XAML bindings:

        IsChecked="{Binding IsChecked}"
        Text="{Binding Name}"
    """

    def __init__(self, element):

        self.element = element
        self.Name = safe_element_name(element)
        self.IsChecked = False


# ==================================================
# OVERRIDE FUNCTIONS
# ==================================================

def record_template_assignments(template_names):
    """
    Record views using destination View Templates whose names match
    the selected linked templates.

    Returns:
        {
            "Template Name": [ViewId, ViewId]
        }
    """

    selected_names = set(template_names)
    assignments = defaultdict(list)

    destination_templates_by_id = {}

    for destination_template in collect_destination_templates():

        destination_templates_by_id[
            element_id_value(destination_template.Id)
        ] = destination_template

    destination_views = (
        FilteredElementCollector(doc)
        .OfClass(View)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    for destination_view in destination_views:

        try:
            if destination_view.IsTemplate:
                continue
        except Exception:
            continue

        try:
            template_id = destination_view.ViewTemplateId
        except Exception:
            continue

        if is_invalid_element_id(template_id):
            continue

        destination_template = destination_templates_by_id.get(
            element_id_value(template_id)
        )

        if destination_template is None:
            continue

        template_name = safe_element_name(
            destination_template
        )

        if template_name in selected_names:
            assignments[template_name].append(
                destination_view.Id
            )

    return assignments


def delete_matching_destination_templates(template_names):
    """
    Delete View Templates in the current project whose names match
    the selected source View Templates.

    Must be called inside an active transaction.
    """

    selected_names = set(template_names)
    deleted_names = []

    for destination_template in collect_destination_templates():

        template_name = safe_element_name(
            destination_template
        )

        if template_name not in selected_names:
            continue

        doc.Delete(destination_template.Id)
        deleted_names.append(template_name)

    return deleted_names


def reassign_replacement_templates(template_assignments):
    """
    Assign newly copied replacement View Templates to views that used
    the deleted destination templates.

    Must be called inside an active transaction.
    """

    replacements_by_name = {}

    for destination_template in collect_destination_templates():

        replacements_by_name[
            safe_element_name(destination_template)
        ] = destination_template

    warnings = []

    for template_name, view_ids in template_assignments.items():

        replacement_template = replacements_by_name.get(
            template_name
        )

        if replacement_template is None:

            warnings.append(
                "Replacement View Template was not found: {}".format(
                    template_name
                )
            )

            continue

        for view_id in view_ids:

            destination_view = doc.GetElement(view_id)

            if destination_view is None:
                continue

            try:
                destination_view.ViewTemplateId = (
                    replacement_template.Id
                )

            except Exception as assignment_error:

                warnings.append(
                    "Could not assign '{}' to view ID {}: {}".format(
                        template_name,
                        element_id_value(view_id),
                        assignment_error
                    )
                )

    return warnings


# ==================================================
# COPY FUNCTIONS
# ==================================================

def create_element_id_list(elements):
    """
    Create a .NET List[ElementId] from Revit elements.
    """

    element_ids = List[ElementId]()

    for element in elements:
        element_ids.Add(element.Id)

    return element_ids


def copy_view_templates(source_doc, source_templates):
    """
    Copy selected View Templates from the linked document into the
    current project.

    Must be called inside an active transaction.
    """

    source_ids = create_element_id_list(
        source_templates
    )

    copy_options = CopyPasteOptions()

    copy_options.SetDuplicateTypeNamesHandler(
        DUPLICATE_TYPE_HANDLER
    )

    copied_ids = ElementTransformUtils.CopyElements(
        source_doc,
        source_ids,
        doc,
        Transform.Identity,
        copy_options
    )

    return copied_ids


# ==================================================
# WPF WINDOW
# ==================================================

class LoadViewTemplatesWindow(forms.WPFWindow):

    def __init__(self):

        # ------------------------------------------
        # VALIDATE XAML
        # ------------------------------------------

        if not os.path.exists(XAML_FILE):

            forms.alert(
                "The UI file was not found:\n\n{}".format(
                    XAML_FILE
                ),
                title=WINDOW_TITLE,
                exitscript=True
            )

        # Load the standalone XAML.
        forms.WPFWindow.__init__(
            self,
            XAML_FILE
        )

        # ------------------------------------------
        # WINDOW DATA
        # ------------------------------------------

        self.links = collect_loaded_links()
        self.source_doc = None

        # Master list containing all View Templates.
        self.all_template_items = []

        # ------------------------------------------
        # INITIAL XAML VALUES
        # ------------------------------------------

        try:
            self.main_title.Text = WINDOW_TITLE
        except Exception:
            pass

        try:
            self.footer_version.Text = __version__
        except Exception:
            pass

        self.UI_Stack_ViewTemplates.Visibility = (
            Visibility.Collapsed
        )

        self.UI_btn_Run.IsEnabled = False

        # ------------------------------------------
        # VALIDATE LINKS
        # ------------------------------------------

        if not self.links:

            forms.alert(
                "No loaded Revit Links were found.\n\n"
                "Open Manage Links and make sure at least one "
                "Revit Link is loaded.",
                title=WINDOW_TITLE,
                exitscript=True
            )

        # ------------------------------------------
        # POPULATE LINK COMBOBOX
        # ------------------------------------------

        for link_name in sorted(
                self.links.keys(),
                key=lambda value: value.lower()):

            # Add strings directly, matching the working filter tool.
            self.UI_LinkSource.Items.Add(link_name)

        # Automatically select the only loaded link.
        if self.UI_LinkSource.Items.Count == 1:
            self.UI_LinkSource.SelectedIndex = 0

        self.ShowDialog()

    # ==================================================
    # WINDOW EVENTS
    # ==================================================

    def header_drag(self, sender, event_args):
        """
        Move the borderless WPF window.

        Matches:
        MouseDown="header_drag"
        """

        try:
            self.DragMove()
        except Exception:
            pass

    def button_close(self, sender, event_args):
        """
        Close the interface.

        Matches:
        Click="button_close"
        """

        self.Close()

    # ==================================================
    # INTERFACE STATE
    # ==================================================

    def show_template_interface(self):
        """Show and enable the View Template selection interface."""

        self.UI_Stack_ViewTemplates.Visibility = (
            Visibility.Visible
        )

        self.update_run_button()

    def hide_template_interface(self):
        """Hide and disable the View Template selection interface."""

        self.UI_Stack_ViewTemplates.Visibility = (
            Visibility.Collapsed
        )

        self.UI_btn_Run.IsEnabled = False

    def update_run_button(self):
        """
        Enable the transfer button when at least one template is
        selected.
        """

        self.UI_btn_Run.IsEnabled = (
            self.selected_template_count() > 0
        )

    # ==================================================
    # LINK SELECTION
    # ==================================================

    def UIe_ComboBox_Changed(self, sender, event_args):
        """
        Load View Templates when the source link changes.

        Matches:
        SelectionChanged="UIe_ComboBox_Changed"
        """

        selected_link = self.UI_LinkSource.SelectedItem

        if selected_link is None:

            self.source_doc = None
            self.all_template_items = []

            self.UI_ListBox_ViewTemplates.ItemsSource = None
            self.hide_template_interface()

            return

        selected_link_name = str(selected_link)

        self.source_doc = self.links.get(
            selected_link_name
        )

        if self.source_doc is None:

            self.all_template_items = []

            self.UI_ListBox_ViewTemplates.ItemsSource = None
            self.hide_template_interface()

            forms.alert(
                "The selected Revit Link is unloaded or unavailable.",
                title=WINDOW_TITLE
            )

            return

        source_templates = collect_view_templates(
            self.source_doc
        )

        self.all_template_items = [
            ViewTemplateItem(view_template)
            for view_template in source_templates
        ]

        self.UI_TextBox_Filter.Text = ""

        self.rebuild_template_list()

        if not self.all_template_items:

            self.hide_template_interface()

            forms.alert(
                "No View Templates were found in the selected "
                "Revit Link.",
                title=WINDOW_TITLE
            )

            return

        self.show_template_interface()

    # ==================================================
    # SEARCH FILTER
    # ==================================================

    def UIe_text_filter_updated(self, sender, event_args):
        """
        Filter the displayed View Template list.

        Matches:
        TextChanged="UIe_text_filter_updated"
        """

        self.rebuild_template_list()

    def rebuild_template_list(self):
        """
        Rebuild the ListBox using the current search keyword.

        ViewTemplateItem objects are retained in the master list, so
        their checked states remain available while filtering.
        """

        if not hasattr(self, "UI_ListBox_ViewTemplates"):
            return

        if not hasattr(self, "UI_TextBox_Filter"):
            return

        filter_text = self.UI_TextBox_Filter.Text

        if filter_text is None:
            filter_text = ""

        filter_text = filter_text.strip().lower()

        visible_items = []

        for template_item in self.all_template_items:

            if filter_text:

                if filter_text not in template_item.Name.lower():
                    continue

            visible_items.append(template_item)

        self.UI_ListBox_ViewTemplates.ItemsSource = None
        self.UI_ListBox_ViewTemplates.ItemsSource = visible_items

        self.update_run_button()

    # ==================================================
    # SELECTION COUNT
    # ==================================================

    def selected_template_count(self):
        """Return the total number of checked View Templates."""

        count = 0

        for template_item in self.all_template_items:

            if template_item.IsChecked:
                count += 1

        return count

    def selected_templates(self):
        """Return selected source View Template elements."""

        selected = []

        for template_item in self.all_template_items:

            if template_item.IsChecked:
                selected.append(
                    template_item.element
                )

        return selected

    # ==================================================
    # SELECT ALL / SELECT NONE
    # ==================================================

    def get_visible_template_items(self):
        """
        Return the ViewTemplateItem objects currently displayed in
        the ListBox.
        """

        visible_items = []

        items_source = (
            self.UI_ListBox_ViewTemplates.ItemsSource
        )

        if items_source is None:
            return visible_items

        for template_item in items_source:
            visible_items.append(template_item)

        return visible_items

    def refresh_visible_template_items(self, visible_items):
        """
        Refresh the ListBox after programmatically changing checked
        values.
        """

        self.UI_ListBox_ViewTemplates.ItemsSource = None
        self.UI_ListBox_ViewTemplates.ItemsSource = visible_items

        self.update_run_button()

    def UIe_btn_select_all(self, sender, event_args):
        """
        Select every currently visible View Template.

        Matches:
        Click="UIe_btn_select_all"
        """

        visible_items = self.get_visible_template_items()

        for template_item in visible_items:
            template_item.IsChecked = True

        self.refresh_visible_template_items(
            visible_items
        )

    def UIe_btn_select_none(self, sender, event_args):
        """
        Deselect every currently visible View Template.

        Matches:
        Click="UIe_btn_select_none"
        """

        visible_items = self.get_visible_template_items()

        for template_item in visible_items:
            template_item.IsChecked = False

        self.refresh_visible_template_items(
            visible_items
        )

    # ==================================================
    # CHECKBOX CLICK SUPPORT
    # ==================================================

    def synchronize_checkbox_states(self):
        """
        Read the current ListBox containers before starting transfer.

        Normally the two-way binding updates IsChecked automatically.
        This method provides an additional synchronization step.
        """

        items_source = (
            self.UI_ListBox_ViewTemplates.ItemsSource
        )

        if items_source is None:
            return

        for template_item in items_source:

            try:
                container = (
                    self.UI_ListBox_ViewTemplates
                    .ItemContainerGenerator
                    .ContainerFromItem(template_item)
                )

                if container is None:
                    continue

                content = container.Content

                # The content normally remains the ViewTemplateItem.
                # Binding updates template_item.IsChecked directly.
                if content is not None:
                    pass

            except Exception:
                continue

    # ==================================================
    # TRANSFER BUTTON
    # ==================================================

    def UIe_btn_run(self, sender, event_args):
        """
        Transfer selected View Templates.

        Matches:
        Click="UIe_btn_run"
        """

        if self.source_doc is None:

            forms.alert(
                "Select a source Revit Link first.",
                title=WINDOW_TITLE
            )

            return

        self.synchronize_checkbox_states()

        selected_templates = self.selected_templates()

        if not selected_templates:

            forms.alert(
                "Select at least one View Template.",
                title=WINDOW_TITLE
            )

            return

        selected_template_names = [
            safe_element_name(view_template)
            for view_template in selected_templates
        ]

        override_existing = bool(
            self.UI_check_override.IsChecked
        )

        # Existing destination templates that match selected names.
        existing_template_names = []

        for template_name in selected_template_names:

            existing_template = find_destination_template(
                template_name
            )

            if existing_template is not None:
                existing_template_names.append(
                    template_name
                )

        template_assignments = {}
        deleted_template_names = []
        reassignment_warnings = []
        copied_element_ids = None

        transaction_group = TransactionGroup(
            doc,
            "Load Selected View Templates From Link"
        )

        group_started = False

        try:

            transaction_group.Start()
            group_started = True

            # --------------------------------------
            # OVERRIDE MATCHING TEMPLATES
            # --------------------------------------

            if override_existing and existing_template_names:

                template_assignments = (
                    record_template_assignments(
                        existing_template_names
                    )
                )

                delete_transaction = Transaction(
                    doc,
                    "Remove Matching View Templates"
                )

                delete_transaction.Start()

                try:

                    deleted_template_names = (
                        delete_matching_destination_templates(
                            existing_template_names
                        )
                    )

                    delete_transaction.Commit()

                except Exception:

                    try:
                        delete_transaction.RollBack()
                    except Exception:
                        pass

                    raise

            # --------------------------------------
            # COPY VIEW TEMPLATES
            # --------------------------------------

            copy_transaction = Transaction(
                doc,
                "Copy View Templates From Link"
            )

            copy_transaction.Start()

            try:

                copied_element_ids = copy_view_templates(
                    self.source_doc,
                    selected_templates
                )

                copy_transaction.Commit()

            except Exception:

                try:
                    copy_transaction.RollBack()
                except Exception:
                    pass

                raise

            # --------------------------------------
            # REASSIGN REPLACED TEMPLATES
            # --------------------------------------

            if override_existing and template_assignments:

                assignment_transaction = Transaction(
                    doc,
                    "Reassign View Templates"
                )

                assignment_transaction.Start()

                try:

                    reassignment_warnings = (
                        reassign_replacement_templates(
                            template_assignments
                        )
                    )

                    assignment_transaction.Commit()

                except Exception:

                    try:
                        assignment_transaction.RollBack()
                    except Exception:
                        pass

                    raise

            transaction_group.Assimilate()
            group_started = False

        except Exception:

            if group_started:

                try:
                    transaction_group.RollBack()
                except Exception:
                    pass

            forms.alert(
                "The View Template transfer failed and all changes "
                "were rolled back.\n\n{}".format(
                    traceback.format_exc()
                ),
                title=WINDOW_TITLE
            )

            return

        # Close only after the complete operation succeeds.
        self.Close()

        # ==================================================
        # RESULT REPORT
        # ==================================================

        updated_names = set(
            deleted_template_names
        )

        added_names = []

        for template_name in selected_template_names:

            if template_name not in updated_names:
                added_names.append(template_name)

        try:
            copied_element_count = copied_element_ids.Count
        except Exception:

            try:
                copied_element_count = len(copied_element_ids)
            except Exception:
                copied_element_count = 0

        result_lines = [
            "Selected View Templates were processed.",
            "",
            "Source link: {}".format(self.source_doc.Title),
            "Selected templates: {}".format(
                len(selected_templates)
            ),
            "New templates copied: {}".format(
                len(added_names)
            ),
            "Existing templates overridden: {}".format(
                len(updated_names)
            ),
            "Revit copied element IDs: {}".format(
                copied_element_count
            )
        ]

        if not override_existing and existing_template_names:

            result_lines.extend([
                "",
                "Override was disabled.",
                "Revit may have created numbered copies for templates "
                "whose names already existed."
            ])

        if added_names:

            result_lines.extend([
                "",
                "New or separately copied templates:"
            ])

            for template_name in added_names:
                result_lines.append(
                    "  - {}".format(template_name)
                )

        if updated_names:

            result_lines.extend([
                "",
                "Overridden templates:"
            ])

            for template_name in sorted(updated_names):
                result_lines.append(
                    "  - {}".format(template_name)
                )

        if reassignment_warnings:

            result_lines.extend([
                "",
                "Reassignment warnings:"
            ])

            for warning in reassignment_warnings:
                result_lines.append(
                    "  - {}".format(warning)
                )

        try:
            uidoc.RefreshActiveView()
        except Exception:
            pass

        forms.alert(
            "\n".join(result_lines),
            title=WINDOW_TITLE
        )


# ==================================================
# COMMAND EXECUTION
# ==================================================

try:

    LoadViewTemplatesWindow()

except Exception:

    forms.alert(
        "The command encountered an unexpected error.\n\n{}".format(
            traceback.format_exc()
        ),
        title=WINDOW_TITLE
    )