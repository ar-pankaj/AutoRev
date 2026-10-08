# -*- coding: utf-8 -*-

__title__ = "Zoom\nSelected"

__author__ = "Pankaj Prabhakar"

__doc__ = """
Zoom to Selected

Fast pyRevit command for Revit 2022-2026.
Designed for IronPython 2.7.

No transaction is required.
"""


from pyrevit import revit, forms


TITLE = "Zoom Selected"


def main():
    uidoc = revit.uidoc

    if uidoc is None:
        forms.alert(
            "Open a Revit project or family document first.",
            title=TITLE,
            warn_icon=True
        )
        return

    # Revit already returns ICollection<ElementId>.
    # Pass it directly to ShowElements without conversion.
    selected_ids = uidoc.Selection.GetElementIds()

    if selected_ids.Count == 0:
        forms.alert(
            "Select one or more elements, then run the command again.",
            title=TITLE,
            warn_icon=True
        )
        return

    try:
        uidoc.ShowElements(selected_ids)

    except Exception as error:
        forms.alert(
            "Revit could not zoom to the selected element(s).\n\n"
            "The selection may not be displayable in a suitable "
            "graphical view.\n\n"
            "Error:\n{0}".format(error),
            title=TITLE,
            warn_icon=True
        )


if __name__ == "__main__":
    main()