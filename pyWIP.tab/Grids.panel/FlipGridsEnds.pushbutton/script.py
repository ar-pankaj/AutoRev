# -*- coding: utf-8 -*-

__title__ = "Flip Grid\nEnds"

__doc__ = """Flip grids so End0 becomes End1 and vice versa (handles MultiSegmentGrid).
            Flip Grid Ends (End0 ↔ End1) – Revit 2024, IronPython (pyRevit)
            - Handles Grid and MultiSegmentGrid
            - Uses GetCurvesInView/SetCurveInView for robust updates
            - Temporarily unpins pinned elements
            Author: Pankaj Prabhakar"""

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import revit, DB, forms

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document


# ---------- Helpers ----------
class GridSelectionFilter(ISelectionFilter):
    def AllowElement(self, e):
        return isinstance(e, (DB.Grid, DB.MultiSegmentGrid))
    def AllowReference(self, ref, pt):
        return False


def get_grids_from_selection():
    sel = [doc.GetElement(eid) for eid in uidoc.Selection.GetElementIds()]
    return [e for e in sel if isinstance(e, (DB.Grid, DB.MultiSegmentGrid))]


def pick_grids():
    refs = uidoc.Selection.PickObjects(ObjectType.Element, GridSelectionFilter(),
                                       "Pick grids to flip End0/End1")
    return [doc.GetElement(r.ElementId) for r in refs]


def grids_in_active_view():
    view = doc.ActiveView
    # Only visible Grid elements in active view
    collector = FilteredElementCollector(doc, view.Id).OfClass(DB.Grid)
    return list(collector)


def reverse_curve(crv):
    # Revit API curves have CreateReversed()
    try:
        return crv.CreateReversed()
    except:
        if hasattr(crv, "Reverse"):
            return crv.Reverse()
        return None


def get_model_curve_in_view(g, view):
    """
    Return a curve for the grid as seen in 'view', preferring Model extents.
    If Model not available, fall back to ViewSpecific in this view.
    """
    # For standard Grid
    if isinstance(g, DB.Grid):
        # Try model curve in view
        crvs = list(g.GetCurvesInView(DatumExtentType.Model, view))
        if crvs:
            return crvs[0], DatumExtentType.Model

        # Fallback: view-specific curve in this view
        crvs = list(g.GetCurvesInView(DatumExtentType.ViewSpecific, view))
        if crvs:
            return crvs[0], DatumExtentType.ViewSpecific

        # Last resort (older API compatibility): try Grid.Curve (may be read-only)
        try:
            if getattr(g, "Curve", None):
                return g.Curve, None
        except:
            pass
        return None, None

    return None, None


def set_curve_in_view(g, view, datum_extent_type, crv):
    """
    Safely set the reversed curve via SetCurveInView.
    Will handle both Model and ViewSpecific, if requested.
    """
    if datum_extent_type in (DatumExtentType.Model, DatumExtentType.ViewSpecific):
        g.SetCurveInView(datum_extent_type, view, crv)
        return True

    # If we only had a raw 'Curve' (legacy), try assigning for compatibility
    try:
        g.Curve = crv
        return True
    except:
        return False


def flip_single_grid(g, view, notes):
    """
    Flip a single Grid (not a MultiSegmentGrid). Returns True/False.
    Adds info to 'notes' dict for diagnostics.
    """
    # Unpin temporarily if pinned
    was_pinned = False
    try:
        was_pinned = g.Pinned
        if was_pinned:
            g.Pinned = False
    except:
        pass

    try:
        src_curve, src_ext = get_model_curve_in_view(g, view)
        if not src_curve:
            notes[g.Id.IntegerValue] = "No curve found in active view."
            return False

        rev_curve = reverse_curve(src_curve)
        if not rev_curve:
            notes[g.Id.IntegerValue] = "Failed to reverse curve."
            return False

        # First try to set Model extent curve (global)
        ok_model = False
        try:
            ok_model = set_curve_in_view(g, view, DatumExtentType.Model, rev_curve)
        except Exception as ex:
            # Common Revit guard: new curve must be coincident with the original datum line
            notes[g.Id.IntegerValue] = "SetCurveInView(Model) failed: {0}".format(ex)

        # If the grid uses ViewSpecific extents in this view, set that as well
        try:
            ext0 = g.GetDatumExtentTypeInView(DatumEnds.End0, view)
            ext1 = g.GetDatumExtentTypeInView(DatumEnds.End1, view)
            if ext0 == DatumExtentType.ViewSpecific or ext1 == DatumExtentType.ViewSpecific:
                try:
                    set_curve_in_view(g, view, DatumExtentType.ViewSpecific, rev_curve)
                except Exception as ex_vs:
                    notes[g.Id.IntegerValue] = "SetCurveInView(ViewSpecific) failed: {0}".format(ex_vs)
        except:
            # Older API or unexpected behavior – ignore
            pass

        # If we succeeded at least once, call it flipped
        if ok_model or (src_ext == DatumExtentType.ViewSpecific):
            return True
        else:
            # Last-chance legacy: try setting Grid.Curve directly
            try:
                if getattr(g, "Curve", None):
                    g.Curve = rev_curve
                    return True
            except Exception as ex_legacy:
                notes[g.Id.IntegerValue] = "Legacy Grid.Curve set failed: {0}".format(ex_legacy)
            return False

    finally:
        # Restore pin state
        try:
            if was_pinned:
                g.Pinned = True
        except:
            pass


def flip_multisegment_grid(msg, view, notes):
    """
    Flip a MultiSegmentGrid by flipping each child Grid in the chain.
    """
    flipped_any = False
    try:
        seg_ids = list(msg.GetGridIds())
    except Exception as ex:
        notes[msg.Id.IntegerValue] = "GetGridIds() failed: {0}".format(ex)
        return False

    for gid in seg_ids:
        sg = doc.GetElement(gid)
        if isinstance(sg, DB.Grid):
            if flip_single_grid(sg, view, notes):
                flipped_any = True
    return flipped_any


# ---------- Main ----------
choice = forms.CommandSwitchWindow.show(
    ['Use current selection', 'Pick grids', 'All grids in active view'],
    message='Flip Grid Ends (End0 ↔ End1): choose source'
)

if not choice:
    forms.alert("Cancelled.", exitscript=True)

if choice == 'Use current selection':
    grids = get_grids_from_selection()
    if not grids:
        forms.alert("No grids in current selection.", exitscript=True)

elif choice == 'Pick grids':
    try:
        grids = pick_grids()
    except:
        forms.alert("Nothing picked.", exitscript=True)

elif choice == 'All grids in active view':
    grids = grids_in_active_view()
    if not grids:
        forms.alert("No grids found in active view.", exitscript=True)

# De-duplicate and ensure types
unique_ids = set([g.Id for g in grids if isinstance(g, (DB.Grid, DB.MultiSegmentGrid))])
grids = [doc.GetElement(eid) for eid in unique_ids]

if not grids:
    forms.alert("No valid grids found.", exitscript=True)

# Flip
t = Transaction(doc, "Flip Grids End0↔End1")
t.Start()
flipped = 0
skipped = 0
notes = {}  # elementId(int) -> reason

active_view = doc.ActiveView

for g in grids:
    try:
        if isinstance(g, DB.Grid):
            ok = flip_single_grid(g, active_view, notes)
        elif isinstance(g, DB.MultiSegmentGrid):
            ok = flip_multisegment_grid(g, active_view, notes)
        else:
            ok = False

        if ok:
            flipped += 1
        else:
            skipped += 1
            if g.Id.IntegerValue not in notes:
                notes[g.Id.IntegerValue] = "Unknown reason."
    except Exception as ex:
        skipped += 1
        notes[g.Id.IntegerValue] = "Exception: {0}".format(ex)

t.Commit()

# Summarise & show details if there were skips
msg = "Flipped: {0}\nSkipped: {1}".format(flipped, skipped)
if skipped > 0:
    msg += "\n\nDetails:\n"
    # Show up to 10 reasons inline (avoid massive message box)
    shown = 0
    for eid, reason in notes.items():
        msg += " - Id {0}: {1}\n".format(eid, reason)
        shown += 1
        if shown >= 10 and len(notes) > 10:
            msg += " - (…{0} more)".format(len(notes) - 10)
            break

forms.alert(msg, title="Flip Grid Ends", warn_icon=(skipped > 0))
