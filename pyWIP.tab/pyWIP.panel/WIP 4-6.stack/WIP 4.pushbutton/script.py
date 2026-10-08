# -*- coding: utf-8 -*-

__author__  = "Pankaj Prabhakar"
__doc__ = """
_________________________________________
Align Legends
_____________________________________________________________________
Description:

Aligns matching legend viewports across multiple selected sheets using
one selected sheet as the reference.

How it works:

- Select at least two sheets in the Project Browser.
- Run the tool and choose a reference sheet.
- Legends are matched by their view names.
- Matching legends on other sheets are moved to the reference position.
- Unmatched legends and non-legend viewports are skipped.
_____________________________________________________________________
"""

from Autodesk.Revit.DB import (
    ViewSheet,
    Viewport,
    ViewType,
    Transaction,
    TransactionGroup
)

from pyrevit import forms, script

# --------------------------------------------------------------------------------------
# VARIABLES
# --------------------------------------------------------------------------------------
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
logger = script.get_logger()
output = script.get_output()


# --------------------------------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------------------------------
def get_selected_sheets():
    """Return selected ViewSheet elements from current selection."""
    sel_ids = uidoc.Selection.GetElementIds()
    sheets = []

    for elid in sel_ids:
        el = doc.GetElement(elid)
        if isinstance(el, ViewSheet):
            sheets.append(el)

    return sheets


def get_legend_viewports(sheet):
    """
    Return dictionary of legend name -> viewport
    from a given sheet.
    """
    legends = {}

    for vp_id in sheet.GetAllViewports():
        vp = doc.GetElement(vp_id)
        if not vp:
            continue

        view = doc.GetElement(vp.ViewId)
        if not view:
            continue

        if view.ViewType == ViewType.Legend:
            legend_name = view.Name

            # Store the viewport by legend name
            # If duplicates somehow exist, keep the first one found
            if legend_name not in legends:
                legends[legend_name] = vp

    return legends


def sheet_display_name(sheet):
    return "{} - {}".format(sheet.SheetNumber, sheet.Name)


class SheetOption(forms.TemplateListItem):
    """Wrapper for SelectFromList."""
    @property
    def name(self):
        return "{} - {}".format(self.item.SheetNumber, self.item.Name)


# --------------------------------------------------------------------------------------
# MAIN
# --------------------------------------------------------------------------------------
def main():
    selected_sheets = get_selected_sheets()

    if len(selected_sheets) < 2:
        forms.alert(
            "Please select at least 2 sheets in the Project Browser and run again.",
            title=__title__,
            exitscript=True
        )

    # Ask user to choose reference sheet
    ref_sheet = forms.SelectFromList.show(
        [SheetOption(x) for x in selected_sheets],
        title="Select Reference Sheet",
        button_name="Use as Reference",
        multiselect=False
    )

    if not ref_sheet:
        forms.alert("No reference sheet selected.", title=__title__, exitscript=True)

    other_sheets = [s for s in selected_sheets if s.Id != ref_sheet.Id]

    ref_legends = get_legend_viewports(ref_sheet)

    if not ref_legends:
        forms.alert(
            "No legends found on the reference sheet:\n{}".format(sheet_display_name(ref_sheet)),
            title=__title__,
            exitscript=True
        )

    print("_" * 100)
    print("Reference Sheet: {}".format(sheet_display_name(ref_sheet)))
    print("Legends found on reference sheet:")
    for lname in sorted(ref_legends.keys()):
        print(" - {}".format(lname))
    print("_" * 100)

    moved_count = 0
    skipped_count = 0

    tg = TransactionGroup(doc, __title__)
    tg.Start()

    try:
        t = Transaction(doc, "Align Legends on Sheets")
        t.Start()

        for sheet in other_sheets:
            print("\nProcessing Sheet: {}".format(sheet_display_name(sheet)))
            other_legends = get_legend_viewports(sheet)

            if not other_legends:
                print("  No legends found on this sheet.")
                skipped_count += 1
                continue

            matched_any = False

            for legend_name, ref_vp in ref_legends.items():
                if legend_name in other_legends:
                    target_vp = other_legends[legend_name]

                    try:
                        ref_center = ref_vp.GetBoxCenter()
                        target_vp.SetBoxCenter(ref_center)

                        print("  Aligned legend: {}".format(legend_name))
                        moved_count += 1
                        matched_any = True
                    except Exception as ex:
                        print("  Failed to align legend '{}': {}".format(legend_name, ex))
                else:
                    print("  Legend not found on sheet: {}".format(legend_name))

            if not matched_any:
                skipped_count += 1

        t.Commit()
        tg.Assimilate()

    except Exception as ex:
        import traceback
        print(traceback.format_exc())
        try:
            t.RollBack()
        except:
            pass
        try:
            tg.RollBack()
        except:
            pass

        forms.alert(
            "An error occurred:\n{}".format(str(ex)),
            title=__title__,
            exitscript=True
        )

    print("\n" + "_" * 100)
    print("Finished.")
    print("Total legends aligned: {}".format(moved_count))
    print("Sheets skipped / no matches: {}".format(skipped_count))
    print("_" * 100)


if __name__ == "__main__":
    main()