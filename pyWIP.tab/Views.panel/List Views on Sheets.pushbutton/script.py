# -*- coding: utf-8 -*-

__title__ = "List Views\nSheets"

__author__ = "Pankaj Prabhakar"

__doc__ = """
List & Select Views Placed on Selected Sheets

Normal Click:
- Shows a list of all placed views (via Viewports) and Schedules (via ScheduleSheetInstance)
  on chosen sheets, with clickable links + a "Select All" link.

Shift+Click:
- Skips the sheet picker.
- Uses currently selected sheets in the Project Browser.
- Selects all placed instances (Viewports + ScheduleSheetInstances) for batch editing.

Compatible: pyRevit (IronPython) | Revit 2024
"""
# pylint: disable=import-error,invalid-name,broad-except,superfluous-parens

__title__ = "List Views\non Sheets"

from pyrevit import revit, forms, script
from Autodesk.Revit import DB

output = script.get_output()
logger = script.get_logger()
doc = revit.doc


def _is_placeholder(sheet):
    try:
        return getattr(sheet, "IsPlaceholder", False)
    except Exception:
        return False


def _get_selected_sheets_from_browser():
    """Return ViewSheet elements from current selection (Project Browser)."""
    sel = revit.get_selection()
    sheets = []
    for el in sel.elements:
        if isinstance(el, DB.ViewSheet) and not _is_placeholder(el):
            sheets.append(el)
    return sheets


def _pick_sheets():
    """Open sheet picker (excludes placeholders) and return selected sheets."""
    sheet_elements = forms.select_sheets(
        button_name="List Placed Views",
        use_selection=True,            # respects current selection if they are sheets
        include_placeholder=False
    )
    if not sheet_elements:
        script.exit()
    return sheet_elements


def _get_viewports_on_sheet(sheet):
    """Return list[DB.Viewport] on the sheet."""
    vport_ids = sheet.GetAllViewports()  # ICollection[ElementId]
    return [doc.GetElement(eid) for eid in vport_ids]


def _get_schedules_on_sheet(sheet):
    """Return list[DB.ScheduleSheetInstance] on the sheet (placed schedules)."""
    return list(DB.FilteredElementCollector(doc, sheet.Id)
                .OfClass(DB.ScheduleSheetInstance)
                .ToElements())


def _format_view_desc(view):
    """Format view descriptor like 'Floor Plan: Level 1'."""
    try:
        vtype = str(view.ViewType)
    except Exception:
        vtype = "View"
    try:
        name = view.Name
    except Exception:
        name = "<Unnamed>"
    return "{}: {}".format(vtype, name)


def _print_views_for_sheet(sheet):
    """Print all placed items for a sheet, return list of instance ids (ElementId)."""
    all_instance_ids = []

    # Viewports (graphical views, legends, drafting, sections, elevations, etc.)
    vports = _get_viewports_on_sheet(sheet)
    for vp in vports:
        v = doc.GetElement(vp.ViewId)
        desc = _format_view_desc(v) if isinstance(v, DB.View) else "Viewport"
        print(
            "SHEET: {0} - {1}\t\tVIEW: {2} {3}".format(
                sheet.SheetNumber,
                sheet.Name,
                desc,
                output.linkify(vp.Id)
            )
        )
        all_instance_ids.append(vp.Id)

    # Schedules (tabular views)
    sched_insts = _get_schedules_on_sheet(sheet)
    for ssi in sched_insts:
        try:
            sched = doc.GetElement(ssi.ScheduleId)
            sname = sched.Name if sched else "Schedule"
        except Exception:
            sname = "Schedule"
        print(
            "SHEET: {0} - {1}\t\tSCHEDULE: {2} {3}".format(
                sheet.SheetNumber,
                sheet.Name,
                sname,
                output.linkify(ssi.Id)
            )
        )
        all_instance_ids.append(ssi.Id)

    return all_instance_ids


def list_views_on_sheets(sheets):
    """Print all placed views & schedules for given sheets and a 'Select All' link."""
    all_ids = []
    for sheet in sheets:
        if _is_placeholder(sheet):
            # Just ignore placeholders silently
            continue
        all_ids.extend(_print_views_for_sheet(sheet))

    if all_ids:
        print(
            "{}".format(
                output.linkify(all_ids, title="Select All Placed Views & Schedules")
            )
        )
    else:
        forms.alert(
            "No placed views or schedules found on the selected sheets.",
            title="Nothing to List",
            warn_icon=True
        )


def select_views_on_sheets(sheets):
    """Select all placed view/schedule instances on given sheets."""
    all_ids = []
    for sheet in sheets:
        if _is_placeholder(sheet):
            continue
        # Collect viewports and schedules
        vports = _get_viewports_on_sheet(sheet)
        scheds = _get_schedules_on_sheet(sheet)
        all_ids.extend([vp.Id for vp in vports])
        all_ids.extend([ssi.Id for ssi in scheds])

    if all_ids:
        revit.get_selection().set_to(all_ids)
    else:
        forms.alert(
            "No placed views or schedules found on the selected sheets.",
            title="Nothing to Select",
            warn_icon=True
        )


# -------------------- Orchestrate --------------------
try:
    if __shiftclick__:
        # Skip picker; use currently selected sheets in Project Browser
        sheets = _get_selected_sheets_from_browser()
        if not sheets:
            forms.alert(
                "Shift+Click mode:\nSelect one or more sheets in the Project Browser and try again.",
                title="No Sheets Selected",
                warn_icon=True
            )
            script.exit()
        select_views_on_sheets(sheets)
    else:
        # Normal click: show list with linkify + Select All link
        list_views_on_sheets(_pick_sheets())

except Exception as ex:
    logger.error("Failed to list/select views on sheets.")
    logger.exception(ex)
    forms.alert(
        "An error occurred. Check pyRevit console for details.",
        title="Error",
        warn_icon=True
    )
