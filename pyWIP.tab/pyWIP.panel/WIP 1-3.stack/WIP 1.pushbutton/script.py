# -*- coding: utf-8 -*-

import clr
import System

__author__ = "Pankaj Prabhakar"
__doc__ = """
_____________________________________________________________________
Select Hosted Elements
_____________________________________________________________________

Description:

Finds and selects elements hosted by a picked Revit element.

How it works:

- Run the tool and pick a host element.
- Select the required hosted elements from the displayed list.
- Click Select to add them to the active Revit selection.
- Switch to a model view before running the tool.
_____________________________________________________________________
"""

from pyrevit import revit, DB, forms
from Autodesk.Revit.UI import Selection

uidoc = revit.uidoc
doc = revit.doc


def safe_cat_name(el):
    try:
        return el.Category.Name if el.Category else "-"
    except:
        return "-"


def display_name(el):
    try:
        if isinstance(el, DB.FamilyInstance):
            fam = ""
            typ = ""
            if el.Symbol:
                typ = el.Symbol.Name or ""
                try:
                    fam = el.Symbol.Family.Name or ""
                except:
                    fam = ""
            if fam or typ:
                return "{0} : {1}".format(fam, typ)
            return el.Name or safe_cat_name(el)
        else:
            return el.Name or safe_cat_name(el)
    except:
        return safe_cat_name(el)


def find_hosted_elements(doc, host_el):
    """
    Return a list of elements hosted by 'host_el'.
    Strategy:
      1) For FamilyInstance: check .Host.Id
      2) For all: check BuiltInParameter.HOST_ID_PARAM
    """
    hosted = []
    all_elems = DB.FilteredElementCollector(doc).WhereElementIsNotElementType()

    for el in all_elems:
        try:
            if el.Id == host_el.Id:
                continue

            # 1) FamilyInstance.Host
            if isinstance(el, DB.FamilyInstance):
                h = el.Host
                if h and h.Id == host_el.Id:
                    hosted.append(el)
                    continue

            # 2) HOST_ID_PARAM
            p = el.get_Parameter(DB.BuiltInParameter.HOST_ID_PARAM)
            if p:
                hid = p.AsElementId()
                if hid and hid.IntegerValue == host_el.Id.IntegerValue:
                    hosted.append(el)
                    continue
        except:
            # keep scanning robustly
            pass

    return hosted


def make_label(el):
    cat = safe_cat_name(el)
    name = display_name(el)
    # include id in the label to guarantee uniqueness
    return "[{0}] {1} (Id:{2})".format(cat, name, el.Id.IntegerValue)


def to_net_list_elementid(py_ids):
    """
    Build a real System.Collections.Generic.List<ElementId> using reflection.
    No [] generic syntax; avoids editor mangling.
    Returns None if it can't build the list.
    """
    try:
        list_generic = System.Type.GetType("System.Collections.Generic.List`1")
        eid_type = clr.GetClrType(DB.ElementId)
        types_arr = System.Array.CreateInstance(clr.GetClrType(System.Type), 1)
        types_arr[0] = eid_type
        net_list_type = list_generic.MakeGenericType(types_arr)
        net_list = System.Activator.CreateInstance(net_list_type)

        for val in py_ids:
            if isinstance(val, DB.ElementId):
                net_list.Add(val)
            else:
                try:
                    net_list.Add(DB.ElementId(int(val)))
                except:
                    pass
        return net_list
    except:
        return None


def can_active_view_accept_selection():
    """Some views (Schedule/Sheet/Report) can't accept model selection."""
    try:
        v = doc.ActiveView
        if v is None:
            return False
        bad_types = [DB.ViewType.Schedule, DB.ViewType.DrawingSheet, DB.ViewType.Report]
        return v.ViewType not in bad_types
    except:
        return True


def main():
    # Quick view sanity check (helps avoid selection API failures)
    if not can_active_view_accept_selection():
        forms.alert("Active view cannot accept selection (e.g., Schedule/Sheet). Switch to a model view and retry.", ok=True)
        return

    # Pick a host element
    try:
        ref = uidoc.Selection.PickObject(Selection.ObjectType.Element, "Pick a host element")
    except:
        forms.alert("Host picking cancelled.", ok=True)
        return

    host_el = doc.GetElement(ref.ElementId)
    if host_el is None:
        forms.alert("Could not resolve the picked element.", ok=True)
        return

    hosted = find_hosted_elements(doc, host_el)
    if not hosted:
        forms.alert("No hosted elements found for: {0}".format(display_name(host_el)), ok=True)
        return

    # Build plain string labels and a mapping back to ElementId (no generics in values)
    labels = []
    label_to_id = {}
    for el in hosted:
        label = make_label(el)
        labels.append(label)
        label_to_id[label] = el.Id  # store ElementId

    # Show selection list with only strings
    selected_labels = forms.SelectFromList.show(
        labels,
        multiselect=True,
        name="Hosted Elements | {0}".format(display_name(host_el)),
        button_name="Select"
    )

    if not selected_labels:
        forms.alert("Nothing selected.", ok=True)
        return

    # Convert selected labels to ElementIds using mapping (primary path)
    ids = []
    for lbl in selected_labels:
        try:
            eid = label_to_id.get(lbl, None)
            if isinstance(eid, DB.ElementId):
                ids.append(eid)
            elif eid is not None:
                ids.append(DB.ElementId(int(eid)))
        except:
            pass

    # Fallback: parse "(Id:123)" at the end of each label
    if not ids:
        try:
            for lbl in selected_labels:
                start = lbl.rfind("(Id:")
                end = lbl.rfind(")")
                if start != -1 and end != -1:
                    num = lbl[start + 4:end].strip()
                    ids.append(DB.ElementId(int(num)))
        except:
            pass

    if not ids:
        forms.alert("Selection conversion failed. Try reloading pyRevit or re-running.", ok=True)
        return

    # Try to set selection using plain Python list
    try:
        uidoc.Selection.SetElementIds(ids)
        return
    except:
        pass

    # Fallback: build real List<ElementId> via reflection and use it
    net_ids = to_net_list_elementid(ids)
    if net_ids:
        try:
            uidoc.Selection.SetElementIds(net_ids)
            return
        except:
            pass

    # Final fallback: set the first element only
    try:
        uidoc.Selection.SetElementIds([ids[0]])
        return
    except:
        # If even this fails, surface a helpful message
        msg = "Could not set selection in Revit.\n" \
              "Possible causes:\n" \
              " - Active view is a Schedule/Sheet or disallows model selection.\n" \
              " - Elements are not selectable in the current context (e.g., family editor).\n" \
              " - Values did not convert to ElementId in your environment."
        forms.alert(msg, ok=True)
        return


if __name__ == "__main__":
    main()
    