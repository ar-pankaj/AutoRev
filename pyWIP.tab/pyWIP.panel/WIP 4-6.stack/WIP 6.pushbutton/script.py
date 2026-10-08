# -*- coding: utf-8 -*-
__author__ = "Pankaj Prabhakar"
__doc__ = """
Change Window Level Without Changing Model Location
Compatible with Revit 2024 and pyRevit IronPython.

Description:
    Changes the associated level of selected windows while preserving
    their original location in the Revit model.

Method:
    1. Store the original insertion-point elevation.
    2. Store the original absolute sill elevation.
    3. Change the window's level.
    4. Recalculate the sill height relative to the target level.
    5. Verify and correct any small elevation difference.
"""

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    Level,
    LocationPoint,
    StorageType,
    SubTransaction,
    Transaction
)

from Autodesk.Revit.UI.Selection import (
    ISelectionFilter,
    ObjectType
)

from pyrevit import revit, forms, script


# ------------------------------------------------------------
# Revit and pyRevit references
# ------------------------------------------------------------

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


# ------------------------------------------------------------
# Selection filter
# ------------------------------------------------------------

class WindowSelectionFilter(ISelectionFilter):
    """Allow the user to select only window elements."""

    def AllowElement(self, element):
        if element is None:
            return False

        if element.Category is None:
            return False

        return (
            element.Category.Id.IntegerValue
            == int(BuiltInCategory.OST_Windows)
        )

    def AllowReference(self, reference, position):
        return False


# ------------------------------------------------------------
# Level-list display class
# ------------------------------------------------------------

class LevelOption(object):
    """Wrapper for displaying Revit levels in the pyRevit dialog."""

    def __init__(self, level):
        self.level = level
        self.name = level.Name
        self.elevation = level.Elevation

    def __str__(self):
        return self.name


# ------------------------------------------------------------
# Helper functions
# ------------------------------------------------------------

def is_window(element):
    """Check whether an element belongs to the Windows category."""

    if element is None:
        return False

    if element.Category is None:
        return False

    return (
        element.Category.Id.IntegerValue
        == int(BuiltInCategory.OST_Windows)
    )


def get_selected_windows():
    """
    Get windows from the current Revit selection.

    If no windows are currently selected, prompt the user
    to select windows manually.
    """

    windows = []

    selected_ids = uidoc.Selection.GetElementIds()

    for element_id in selected_ids:
        element = doc.GetElement(element_id)

        if is_window(element):
            windows.append(element)

    if windows:
        return windows

    try:
        references = uidoc.Selection.PickObjects(
            ObjectType.Element,
            WindowSelectionFilter(),
            "Select windows whose level should be changed"
        )

        for reference in references:
            element = doc.GetElement(reference.ElementId)

            if is_window(element):
                windows.append(element)

    except Exception:
        # The user probably cancelled the selection.
        return []

    return windows


def get_location_point(element):
    """
    Return the insertion point of an element.

    Standard wall-hosted window family instances normally
    use LocationPoint.
    """

    location = element.Location

    if location is None:
        return None

    if isinstance(location, LocationPoint):
        return location.Point

    if hasattr(location, "Point"):
        return location.Point

    return None


def get_level_parameter(element):
    """
    Find an editable level parameter for the window.

    FAMILY_LEVEL_PARAM is normally used for standard
    wall-hosted window family instances.
    """

    possible_parameters = [
        BuiltInParameter.FAMILY_LEVEL_PARAM,
        BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
        BuiltInParameter.SCHEDULE_LEVEL_PARAM
    ]

    for parameter_id in possible_parameters:
        parameter = element.get_Parameter(parameter_id)

        if parameter is None:
            continue

        if parameter.IsReadOnly:
            continue

        if parameter.StorageType != StorageType.ElementId:
            continue

        return parameter

    return None


def get_sill_height_parameter(element):
    """Return the editable sill-height parameter of a window."""

    parameter = element.get_Parameter(
        BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM
    )

    if parameter is None:
        return None

    if parameter.IsReadOnly:
        return None

    if parameter.StorageType != StorageType.Double:
        return None

    return parameter


def get_level_from_parameter(level_parameter):
    """
    Return the Revit Level referenced by a level parameter.
    """

    if level_parameter is None:
        return None

    if level_parameter.StorageType != StorageType.ElementId:
        return None

    level_id = level_parameter.AsElementId()

    if level_id is None:
        return None

    if level_id == ElementId.InvalidElementId:
        return None

    level = doc.GetElement(level_id)

    if isinstance(level, Level):
        return level

    return None


def get_window_description(window):
    """Return a readable description of a window element."""

    element_id = window.Id.IntegerValue
    family_name = "Unknown family"
    type_name = "Unknown type"

    try:
        symbol = window.Symbol

        if symbol is not None:
            type_name = symbol.Name

            if symbol.Family is not None:
                family_name = symbol.Family.Name

    except Exception:
        pass

    return "{} : {} | ID {}".format(
        family_name,
        type_name,
        element_id
    )


# ------------------------------------------------------------
# Get selected windows
# ------------------------------------------------------------

windows = get_selected_windows()

if not windows:
    forms.alert(
        "No windows were selected.\n\n"
        "Select one or more windows and run the tool again.",
        title="Change Window Level",
        exitscript=True
    )


# ------------------------------------------------------------
# Collect levels
# ------------------------------------------------------------

levels = list(
    FilteredElementCollector(doc)
    .OfClass(Level)
    .WhereElementIsNotElementType()
)

if not levels:
    forms.alert(
        "No levels were found in the current Revit model.",
        title="Change Window Level",
        exitscript=True
    )


# Sort levels by elevation, from lowest to highest.
levels.sort(key=lambda level: level.Elevation)

level_options = []

for level in levels:
    level_options.append(LevelOption(level))


# ------------------------------------------------------------
# Ask the user to select the target level
# ------------------------------------------------------------

selected_level_option = forms.SelectFromList.show(
    level_options,
    title="Select New Window Level",
    button_name="Change Window Level",
    multiselect=False,
    name_attr="name",
    width=500,
    height=600
)

if selected_level_option is None:
    script.exit()

target_level = selected_level_option.level


# ------------------------------------------------------------
# Result containers
# ------------------------------------------------------------

successful = []
skipped = []
failed = []


# ------------------------------------------------------------
# Main transaction
# ------------------------------------------------------------

transaction = Transaction(
    doc,
    "Change Window Level Without Moving"
)

transaction.Start()

try:

    for window in windows:

        # A subtransaction allows one failed window to be rolled
        # back without cancelling all successfully changed windows.
        subtransaction = SubTransaction(doc)
        subtransaction.Start()

        try:
            window_description = get_window_description(window)

            # ----------------------------------------------------
            # Check whether the element can be modified
            # ----------------------------------------------------

            if window.GroupId != ElementId.InvalidElementId:
                skipped.append(
                    (
                        window.Id,
                        window_description,
                        "Window belongs to a model group"
                    )
                )

                subtransaction.RollBack()
                continue

            if window.Document.IsLinked:
                skipped.append(
                    (
                        window.Id,
                        window_description,
                        "Window belongs to a linked model"
                    )
                )

                subtransaction.RollBack()
                continue

            # ----------------------------------------------------
            # Get required parameters
            # ----------------------------------------------------

            level_parameter = get_level_parameter(window)
            sill_parameter = get_sill_height_parameter(window)

            if level_parameter is None:
                skipped.append(
                    (
                        window.Id,
                        window_description,
                        "No editable level parameter was found"
                    )
                )

                subtransaction.RollBack()
                continue

            if sill_parameter is None:
                skipped.append(
                    (
                        window.Id,
                        window_description,
                        "No editable sill-height parameter was found"
                    )
                )

                subtransaction.RollBack()
                continue

            old_level = get_level_from_parameter(level_parameter)

            if old_level is None:
                skipped.append(
                    (
                        window.Id,
                        window_description,
                        "The current window level could not be identified"
                    )
                )

                subtransaction.RollBack()
                continue

            # ----------------------------------------------------
            # Store the original location
            # ----------------------------------------------------

            original_point = get_location_point(window)

            if original_point is None:
                skipped.append(
                    (
                        window.Id,
                        window_description,
                        "The window insertion point could not be identified"
                    )
                )

                subtransaction.RollBack()
                continue

            original_x = original_point.X
            original_y = original_point.Y
            original_z = original_point.Z

            # ----------------------------------------------------
            # Calculate the absolute sill elevation
            # ----------------------------------------------------

            old_sill_height = sill_parameter.AsDouble()

            absolute_sill_elevation = (
                old_level.Elevation + old_sill_height
            )

            # Calculate the sill-height offset relative to the
            # selected target level.
            new_sill_height = (
                absolute_sill_elevation - target_level.Elevation
            )

            # ----------------------------------------------------
            # Change the level and sill height
            # ----------------------------------------------------

            level_parameter.Set(target_level.Id)
            sill_parameter.Set(new_sill_height)

            doc.Regenerate()

            # ----------------------------------------------------
            # Verify that the window did not move vertically
            # ----------------------------------------------------

            updated_point = get_location_point(window)

            if updated_point is not None:
                elevation_difference = (
                    original_z - updated_point.Z
                )

                tolerance = 0.000001

                if abs(elevation_difference) > tolerance:
                    corrected_sill_height = (
                        sill_parameter.AsDouble()
                        + elevation_difference
                    )

                    sill_parameter.Set(corrected_sill_height)
                    doc.Regenerate()

            # ----------------------------------------------------
            # Final position check
            # ----------------------------------------------------

            final_point = get_location_point(window)

            if final_point is None:
                raise Exception(
                    "The final window location could not be verified"
                )

            x_difference = abs(final_point.X - original_x)
            y_difference = abs(final_point.Y - original_y)
            z_difference = abs(final_point.Z - original_z)

            final_tolerance = 0.0001

            if (
                x_difference > final_tolerance
                or y_difference > final_tolerance
                or z_difference > final_tolerance
            ):
                raise Exception(
                    "The window position changed more than the "
                    "permitted tolerance"
                )

            # ----------------------------------------------------
            # Commit this window
            # ----------------------------------------------------

            subtransaction.Commit()

            successful.append(
                (
                    window.Id,
                    window_description,
                    old_level.Name,
                    target_level.Name
                )
            )

        except Exception as window_error:

            if subtransaction.HasStarted():
                subtransaction.RollBack()

            failed.append(
                (
                    window.Id,
                    get_window_description(window),
                    str(window_error)
                )
            )

    transaction.Commit()

except Exception as main_error:

    if transaction.HasStarted():
        transaction.RollBack()

    forms.alert(
        "The transaction could not be completed.\n\n"
        "Error:\n{}".format(str(main_error)),
        title="Change Window Level",
        exitscript=True
    )


# ------------------------------------------------------------
# Output report
# ------------------------------------------------------------

output.print_md("# Change Window Level Report")

output.print_md(
    "**Target level:** `{}`".format(target_level.Name)
)

output.print_md(
    "**Windows selected:** {}".format(len(windows))
)

output.print_md(
    "**Successfully changed:** {}".format(len(successful))
)

output.print_md(
    "**Skipped:** {}".format(len(skipped))
)

output.print_md(
    "**Failed:** {}".format(len(failed))
)


# ------------------------------------------------------------
# Successful windows
# ------------------------------------------------------------

if successful:

    output.print_md("## Successfully Changed")

    for (
        element_id,
        description,
        old_level_name,
        new_level_name
    ) in successful:

        output.print_md(
            "- {} | {} → {} | {}".format(
                description,
                old_level_name,
                new_level_name,
                output.linkify(element_id)
            )
        )


# ------------------------------------------------------------
# Skipped windows
# ------------------------------------------------------------

if skipped:

    output.print_md("## Skipped")

    for element_id, description, reason in skipped:

        output.print_md(
            "- {} | Reason: {} | {}".format(
                description,
                reason,
                output.linkify(element_id)
            )
        )


# ------------------------------------------------------------
# Failed windows
# ------------------------------------------------------------

if failed:

    output.print_md("## Failed")

    for element_id, description, error_message in failed:

        output.print_md(
            "- {} | Error: {} | {}".format(
                description,
                error_message,
                output.linkify(element_id)
            )
        )


# ------------------------------------------------------------
# Final message
# ------------------------------------------------------------

message = (
    "Target level: {}\n\n"
    "Successfully changed: {}\n"
    "Skipped: {}\n"
    "Failed: {}"
).format(
    target_level.Name,
    len(successful),
    len(skipped),
    len(failed)
)

forms.alert(
    message,
    title="Change Window Level"
)