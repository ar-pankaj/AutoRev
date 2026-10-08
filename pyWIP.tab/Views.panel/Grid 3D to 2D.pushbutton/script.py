# -*- coding: utf-8 -*-

__title__ = "Grids\n3D→2D"

__author__ = "Pankaj Prabhakar"

__doc__ = """pyRevit: Convert selected grids from 3D to 2D in the Active View
Revit 2024 | Works with pyRevit (IronPython)

What it does:
- Finds grids from current selection
- For each grid, sets both ends (End0 & End1) to ViewSpecific (2D) in the active view
- Skips any grid not visible in the view (records reason)

Notes:
- Revit API allows each datum end to have its own extent mode per view.
- SetDatumExtentType throws if the datum can't be visible in the given view, so we catch & report."""

from pyrevit import revit, DB, script

doc   = revit.doc
uidoc = revit.uidoc
view  = doc.ActiveView

# Collect selected elements that are grids
selected_elements = list(revit.get_selection().elements)
grids = [e for e in selected_elements
         if e and e.Category
         and e.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_Grids)]

if not grids:
    script.exit("Select one or more grids, then run again.")

converted = 0
skipped   = []  # (grid_name, reason)

t = DB.Transaction(doc, "Grids: 3D → 2D in '{}'".format(view.Name))
t.Start()
for g in grids:
    ok = True
    for end in (DB.DatumEnds.End0, DB.DatumEnds.End1):
        try:
            # Flip each end to 2D (view-specific) in the active view
            g.SetDatumExtentType(end, view, DB.DatumExtentType.ViewSpecific)
        except Exception as ex:
            ok = False
            skipped.append((getattr(g, "Name", "<unnamed>"), str(ex)))
            break
    if ok:
        converted += 1
t.Commit()

# Report
out = script.get_output()
out.print_md("### ✅ Converted {} grid(s) to **2D** in view **{}**".format(converted, view.Name))
if skipped:
    out.print_md("### ⚠️ Skipped (not visible or other constraint)")
    for name, reason in skipped:
        out.print_md("- `{}` — {}".format(name, reason))