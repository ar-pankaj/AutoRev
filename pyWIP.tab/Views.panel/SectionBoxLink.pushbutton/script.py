# -*- coding: utf-8 -*-

__title__ = "SectionBox\nRevitLink"

__author__ = "Pankaj Prabhakar"

__doc__ = """
Creates a section box around selected elements inside Revit links.

Normal click:
    Uses currently selected linked elements.

    If no valid linked elements are selected, the command prompts
    the user to select linked elements.

    The section box is applied to the 3D view selected through
    Shift-click configuration.

Shift-click:
    Select the target 3D view, configure the section-box offset,
    or review the current settings."""


from pyrevit import revit, DB, UI, forms, script, EXEC_PARAMS


# ----------------------------------------------------------------------
# REVIT CONTEXT
# ----------------------------------------------------------------------

doc = revit.doc
uidoc = revit.uidoc

DEFAULT_OFFSET_MM = 100.0

logger = script.get_logger()
config = script.get_config()


# ----------------------------------------------------------------------
# GENERAL HELPERS
# ----------------------------------------------------------------------

def element_id_value(element_id):
    """
    Return a Revit ElementId as a Python integer.

    Revit 2024 and newer use ElementId.Value.
    Older Revit versions use ElementId.IntegerValue.
    """

    try:
        return int(element_id.Value)
    except:
        return int(element_id.IntegerValue)


def get_element_id_from_value(value):
    """
    Create a Revit ElementId from a saved integer value.

    Python's int supports both normal and large integer values.
    """

    return DB.ElementId(int(value))


def millimetres_to_internal_units(value_mm):
    """Convert millimetres to Revit internal length units."""

    try:
        # Revit 2021 and newer
        return DB.UnitUtils.ConvertToInternalUnits(
            float(value_mm),
            DB.UnitTypeId.Millimeters
        )
    except:
        # Compatibility fallback for older Revit versions
        return float(value_mm) / 304.8


# ----------------------------------------------------------------------
# LINKED-ELEMENT SELECTION VALIDATION
# ----------------------------------------------------------------------

def is_valid_linked_reference(reference):
    """
    Check whether a reference represents an element inside a loaded
    Revit link.
    """

    if reference is None:
        return False

    try:
        linked_element_id = reference.LinkedElementId

        if linked_element_id is None:
            return False

        if linked_element_id == DB.ElementId.InvalidElementId:
            return False

        link_instance = doc.GetElement(reference.ElementId)

        if not isinstance(link_instance, DB.RevitLinkInstance):
            return False

        link_document = link_instance.GetLinkDocument()

        if link_document is None:
            return False

        linked_element = link_document.GetElement(linked_element_id)

        return linked_element is not None

    except:
        return False


# ----------------------------------------------------------------------
# SELECTION FILTER
# ----------------------------------------------------------------------

class LinkedElementSelectionFilter(UI.Selection.ISelectionFilter):
    """Allow selection only from loaded Revit links."""

    def AllowElement(self, element):
        return isinstance(element, DB.RevitLinkInstance)

    def AllowReference(self, reference, position):
        return is_valid_linked_reference(reference)


# ----------------------------------------------------------------------
# LINKED-ELEMENT SELECTION
# ----------------------------------------------------------------------

def get_current_linked_selection():
    """
    Return linked-element references from the current Revit selection.

    If Revit stores only the link instance instead of the linked
    sub-element reference, this function returns an empty list and the
    command automatically opens the linked-element selection prompt.
    """

    linked_references = []

    try:
        current_references = uidoc.Selection.GetReferences()
    except:
        current_references = []

    for reference in current_references:
        if is_valid_linked_reference(reference):
            linked_references.append(reference)

    return linked_references


def pick_linked_elements():
    """
    Prompt the user to select linked elements.

    This function is used only when no valid linked-element references
    are found in the current selection.
    """

    selection_filter = LinkedElementSelectionFilter()

    try:
        with forms.WarningBar(
            title="Select linked elements, then click Finish"
        ):
            picked_references = uidoc.Selection.PickObjects(
                UI.Selection.ObjectType.LinkedElement,
                selection_filter,
                "Select linked elements"
            )

        valid_references = []

        for reference in picked_references:
            if is_valid_linked_reference(reference):
                valid_references.append(reference)

        return valid_references

    except UI.Exceptions.OperationCanceledException:
        return []

    except Exception as exception:
        logger.exception(
            "Could not select linked elements: {0}".format(
                exception
            )
        )
        return []


def get_linked_element_references():
    """
    Use currently selected linked elements first.

    If no valid linked elements are currently selected, open the
    linked-element selection prompt.
    """

    references = get_current_linked_selection()

    if references:
        logger.info(
            "Using {0} linked element(s) from the current selection."
            .format(len(references))
        )
        return references

    return pick_linked_elements()


# ----------------------------------------------------------------------
# 3D VIEW FUNCTIONS
# ----------------------------------------------------------------------

def is_valid_section_box_view(view):
    """
    Return True when the view is a usable orthographic 3D view.
    """

    if view is None:
        return False

    if not isinstance(view, DB.View3D):
        return False

    if view.IsTemplate:
        return False

    if view.IsPerspective:
        return False

    return True


def get_available_3d_views():
    """
    Return all usable orthographic 3D views sorted by view name.
    """

    available_views = []

    collector = (
        DB.FilteredElementCollector(doc)
        .OfClass(DB.View3D)
    )

    for view in collector:
        if is_valid_section_box_view(view):
            available_views.append(view)

    available_views.sort(
        key=lambda item: item.Name.lower()
    )

    return available_views


def get_saved_3d_view():
    """
    Return the 3D view stored in the command configuration.

    The saved ElementId is checked first. The saved view name is used
    as a fallback if the ElementId is no longer valid.
    """

    saved_view_id = None
    saved_view_name = None

    try:
        saved_view_id = int(config.target_3d_view_id)
    except:
        saved_view_id = None

    try:
        saved_view_name = str(config.target_3d_view_name)
    except:
        saved_view_name = None

    # First try the saved ElementId.
    if saved_view_id is not None:
        try:
            view_id = get_element_id_from_value(saved_view_id)
            saved_view = doc.GetElement(view_id)

            if is_valid_section_box_view(saved_view):
                return saved_view
        except:
            pass

    # If the ElementId is invalid, try the saved view name.
    if saved_view_name:
        saved_name_lower = saved_view_name.lower()

        for view in get_available_3d_views():
            if view.Name.lower() == saved_name_lower:
                return view

    return None


def save_target_3d_view(view):
    """Save the selected target 3D view."""

    config.target_3d_view_id = element_id_value(view.Id)
    config.target_3d_view_name = view.Name

    script.save_config()


def select_target_3d_view(show_confirmation=True):
    """
    Display all available orthographic 3D views and save the selected
    view as the target for this command.
    """

    available_views = get_available_3d_views()

    if not available_views:
        forms.alert(
            "No usable orthographic 3D views were found.\n\n"
            "Create a non-perspective 3D view and run the command again.",
            title="Select Target 3D View",
            warn_icon=True
        )
        return None

    view_names = []

    for view in available_views:
        view_names.append(view.Name)

    selected_name = forms.SelectFromList.show(
        view_names,
        title="Select Section Box 3D View",
        button_name="Save 3D View",
        multiselect=False,
        width=500,
        height=600
    )

    if not selected_name:
        return None

    selected_view = None

    for view in available_views:
        if view.Name == selected_name:
            selected_view = view
            break

    if selected_view is None:
        return None

    save_target_3d_view(selected_view)

    if show_confirmation:
        forms.alert(
            "Target 3D view saved:\n\n{0}".format(
                selected_view.Name
            ),
            title="Section Box Configuration"
        )

    return selected_view


def get_target_3d_view():
    """
    Return the target 3D view.

    Priority:
        1. 3D view saved through Shift-click.
        2. Active orthographic 3D view.
        3. Ask the user to select a target 3D view.

    The command does not use a random existing 3D view.
    """

    saved_view = get_saved_3d_view()

    if saved_view is not None:
        return saved_view

    active_view = doc.ActiveView

    if is_valid_section_box_view(active_view):
        save_target_3d_view(active_view)
        return active_view

    return select_target_3d_view(
        show_confirmation=False
    )


# ----------------------------------------------------------------------
# OFFSET CONFIGURATION
# ----------------------------------------------------------------------

def get_saved_offset_mm():
    """Return the saved offset or the default offset."""

    try:
        return float(config.section_box_offset_mm)
    except:
        return DEFAULT_OFFSET_MM


def configure_offset():
    """
    Ask for a section-box offset and save it in pyRevit configuration.
    """

    current_offset = get_saved_offset_mm()

    value = forms.ask_for_string(
        default="{0:g}".format(current_offset),
        prompt=(
            "Enter the section box offset in millimetres.\n\n"
            "The offset will be added outside all six sides of the "
            "selected linked elements."
        ),
        title="Section Box Offset"
    )

    if value is None:
        return False

    try:
        value = value.strip().replace(",", ".")
        offset_mm = float(value)

        if offset_mm < 0:
            forms.alert(
                "Offset must be zero or greater.",
                title="Invalid Offset",
                warn_icon=True
            )
            return False

        config.section_box_offset_mm = offset_mm
        script.save_config()

        return True

    except ValueError:
        forms.alert(
            "Please enter a valid numerical value.",
            title="Invalid Offset",
            warn_icon=True
        )
        return False


# ----------------------------------------------------------------------
# SHIFT-CLICK CONFIGURATION
# ----------------------------------------------------------------------

def show_current_settings():
    """Display the currently saved 3D view and offset."""

    saved_view = get_saved_3d_view()
    offset_mm = get_saved_offset_mm()

    if saved_view is not None:
        view_name = saved_view.Name
    else:
        view_name = "Not configured"

    forms.alert(
        "Target 3D view:\n"
        "{0}\n\n"
        "Section box offset:\n"
        "{1:g} mm\n\n"
        "Shift-click the tool to change these settings.".format(
            view_name,
            offset_mm
        ),
        title="Section Box Settings"
    )


def configure_tool():
    """
    Open the Shift-click configuration menu.

    Options:
        - Select target 3D view
        - Set section-box offset
        - View current settings
    """

    selected_option = forms.CommandSwitchWindow.show(
        [
            "Select Target 3D View",
            "Set Section Box Offset",
            "View Current Settings"
        ],
        message="Configure SectionBox RevitLink"
    )

    if not selected_option:
        return

    if selected_option == "Select Target 3D View":
        select_target_3d_view(
            show_confirmation=True
        )

    elif selected_option == "Set Section Box Offset":
        if configure_offset():
            forms.alert(
                "Section box offset saved as {0:g} mm.".format(
                    get_saved_offset_mm()
                ),
                title="Section Box Configuration"
            )

    elif selected_option == "View Current Settings":
        show_current_settings()


# ----------------------------------------------------------------------
# BOUNDING BOX CALCULATION
# ----------------------------------------------------------------------

def get_transformed_bbox_points(bounding_box, link_transform):
    """
    Return all eight bounding-box corners in host-model coordinates.

    All eight corners are transformed so rotated or mirrored Revit
    links are handled correctly.
    """

    minimum = bounding_box.Min
    maximum = bounding_box.Max

    bbox_transform = bounding_box.Transform

    coordinates = (
        (minimum.X, minimum.Y, minimum.Z),
        (minimum.X, minimum.Y, maximum.Z),
        (minimum.X, maximum.Y, minimum.Z),
        (minimum.X, maximum.Y, maximum.Z),
        (maximum.X, minimum.Y, minimum.Z),
        (maximum.X, minimum.Y, maximum.Z),
        (maximum.X, maximum.Y, minimum.Z),
        (maximum.X, maximum.Y, maximum.Z),
    )

    transformed_points = []

    for x, y, z in coordinates:
        point = DB.XYZ(x, y, z)

        # Bounding-box-local coordinates to linked-document coordinates.
        point = bbox_transform.OfPoint(point)

        # Linked-document coordinates to host-document coordinates.
        point = link_transform.OfPoint(point)

        transformed_points.append(point)

    return transformed_points


def get_section_box_from_references(references, offset_internal):
    """
    Create one combined BoundingBoxXYZ around all selected linked
    elements.
    """

    min_x = float("inf")
    min_y = float("inf")
    min_z = float("inf")

    max_x = float("-inf")
    max_y = float("-inf")
    max_z = float("-inf")

    valid_element_count = 0

    # Cache link documents and transforms for faster processing.
    link_cache = {}

    for reference in references:
        link_instance_id = reference.ElementId
        cache_key = element_id_value(link_instance_id)

        if cache_key not in link_cache:
            link_instance = doc.GetElement(link_instance_id)

            if not isinstance(link_instance, DB.RevitLinkInstance):
                continue

            link_document = link_instance.GetLinkDocument()

            if link_document is None:
                continue

            link_cache[cache_key] = (
                link_document,
                link_instance.GetTotalTransform()
            )

        link_document, link_transform = link_cache[cache_key]

        linked_element = link_document.GetElement(
            reference.LinkedElementId
        )

        if linked_element is None:
            continue

        # Element bounding boxes are faster than complete geometry.
        bounding_box = linked_element.get_BoundingBox(None)

        if bounding_box is None:
            continue

        transformed_points = get_transformed_bbox_points(
            bounding_box,
            link_transform
        )

        for point in transformed_points:
            min_x = min(min_x, point.X)
            min_y = min(min_y, point.Y)
            min_z = min(min_z, point.Z)

            max_x = max(max_x, point.X)
            max_y = max(max_y, point.Y)
            max_z = max(max_z, point.Z)

        valid_element_count += 1

    if valid_element_count == 0:
        return None, 0

    section_box = DB.BoundingBoxXYZ()
    section_box.Transform = DB.Transform.Identity

    section_box.Min = DB.XYZ(
        min_x - offset_internal,
        min_y - offset_internal,
        min_z - offset_internal
    )

    section_box.Max = DB.XYZ(
        max_x + offset_internal,
        max_y + offset_internal,
        max_z + offset_internal
    )

    return section_box, valid_element_count


# ----------------------------------------------------------------------
# VIEW ACTIVATION AND ZOOM
# ----------------------------------------------------------------------

def activate_and_zoom(view):
    """Activate the configured 3D view and zoom to fit."""

    try:
        if uidoc.ActiveView.Id != view.Id:
            uidoc.ActiveView = view

    except:
        try:
            uidoc.RequestViewChange(view)

        except Exception as exception:
            logger.warning(
                "Could not activate 3D view: {0}".format(
                    exception
                )
            )

    try:
        uidoc.RefreshActiveView()
    except:
        pass

    try:
        for ui_view in uidoc.GetOpenUIViews():
            if ui_view.ViewId == view.Id:
                ui_view.ZoomToFit()
                break

    except Exception as exception:
        logger.warning(
            "Could not zoom the 3D view to fit: {0}".format(
                exception
            )
        )


# ----------------------------------------------------------------------
# MAIN COMMAND
# ----------------------------------------------------------------------

def main():
    """Run the linked-element section-box command."""

    # First use linked elements from the current Revit selection.
    # If none are available, open the selection prompt.
    references = get_linked_element_references()

    if not references:
        return

    offset_mm = get_saved_offset_mm()
    offset_internal = millimetres_to_internal_units(offset_mm)

    section_box, valid_count = get_section_box_from_references(
        references,
        offset_internal
    )

    if section_box is None:
        forms.alert(
            "A valid bounding box could not be obtained from the "
            "selected linked elements.\n\n"
            "The elements may not have model geometry, or their "
            "Revit link may not be loaded.",
            title="Section Box",
            warn_icon=True
        )
        return

    target_view = get_target_3d_view()

    if target_view is None:
        forms.alert(
            "A target 3D view has not been configured.\n\n"
            "Shift-click this tool and choose "
            "'Select Target 3D View'.",
            title="Section Box",
            warn_icon=True
        )
        return

    try:
        with revit.Transaction("Section Box - Revit Link"):
            target_view.SetSectionBox(section_box)
            target_view.IsSectionBoxActive = True

    except Exception as exception:
        logger.exception(
            "Could not apply section box: {0}".format(
                exception
            )
        )

        forms.alert(
            "The section box could not be applied to:\n\n"
            "{0}\n\n"
            "The view may be controlled by a view template, deleted, "
            "locked, or unsuitable for a section box.".format(
                target_view.Name
            ),
            title="Section Box",
            warn_icon=True
        )
        return

    activate_and_zoom(target_view)

    logger.info(
        "Section box applied to '{0}' around {1} linked element(s), "
        "using an offset of {2:g} mm.".format(
            target_view.Name,
            valid_count,
            offset_mm
        )
    )


# ----------------------------------------------------------------------
# COMMAND ENTRY
# ----------------------------------------------------------------------

# Do not add a separate config.py file to this pushbutton bundle.
# Shift-click is handled directly through EXEC_PARAMS.config_mode.

if EXEC_PARAMS.config_mode:
    configure_tool()
else:
    main()