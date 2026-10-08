# -*- coding: utf-8 -*-

__author__ = "Pankaj Prabhakar"

__doc__ = """
_____________________________________________________________________
Select Curtain Panels
_____________________________________________________________________

Description:

Selects multiple curtain wall panels while filtering out mullions,
curtain grids, and other elements.

How it works:

- Run the tool and click the required curtain panels.
- Press Finish on the Options Bar when done.
- The selected panels are added to the active Revit selection.
- Press Esc to cancel the operation.
_____________________________________________________________________
"""

import clr
from Autodesk.Revit.DB import BuiltInCategory, ElementId
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Collections.Generic import List

# Try to use native Revit TaskDialog; fall back to pyRevit forms if not available
try:
    from Autodesk.Revit.UI import TaskDialog
    USE_TASKDIALOG = True
except:
    USE_TASKDIALOG = False
    from pyrevit import forms

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document


class CurtainPanelSelectionFilter(ISelectionFilter):
    """Allow only curtain wall panels (no mullions, grids, or other elements)."""

    def AllowElement(self, element):
        try:
            cat = element.Category
            return bool(cat) and cat.Id.IntegerValue == int(BuiltInCategory.OST_CurtainWallPanels)
        except:
            return False

    def AllowReference(self, reference, point):
        # For ObjectType.Element picks, returning True is safe and avoids blocking the click.
        return True


def main():
    prompt = ("Click curtain panels to add to selection.\n"
              "Press Finish (green check) on the Options Bar when done, or Esc to cancel.")
    filt = CurtainPanelSelectionFilter()

    try:
        # Multi-pick with a visible Finish option
        refs = uidoc.Selection.PickObjects(ObjectType.Element, filt, prompt)
    except OperationCanceledException:
        # User pressed Esc or canceled
        return

    if not refs:
        return

    # Convert References -> ElementIds (more robust than relying on Reference.ElementId)
    id_list = [doc.GetElement(r).Id for r in refs]
    ids = List[ElementId](id_list)

    # Apply selection to the UI
    uidoc.Selection.SetElementIds(ids)

    # Notify
    msg = "Finished. Selected {} curtain panel(s).".format(ids.Count)
    if USE_TASKDIALOG:
        TaskDialog.Show("Curtain Panels", msg)
    else:
        # Fallback if TaskDialog import is not available (e.g., some environments)
        forms.alert(msg, title="Curtain Panels", warn_icon=False)


if __name__ == "__main__":
    main()