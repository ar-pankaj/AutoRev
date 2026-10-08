# -*- coding: utf-8 -*-
__title__   = "Paint Peeler"
__doc__     = """Version = 2.0
Date    = 23.06.2026
________________________________________________________________
Description:
Find the painted surface.
________________________________________________________________
How-To:
1. Step 1 #⛔Active View must be 3D
2. Step 2 #⛔Select Architecture Categories to Check
3. Step 3 #⛔Get elements from selected categories
4. Step 4 #⛔Get the Element geometry
5. Step 5 #⛔Find Painted surface
6. Step 6 #⛔Create Red Graphic Override
7. Step 7 #⛔Override Painted elements 
8. Step 8 #⛔Painted Element Report
________________________________________________________________
To-Do:
1.Create a flex form to select elements User wish to run the tool. (imp for big projects)
2.Detect painted split faces, and override the area that is painted.
3.Override only the surface that is painted.( Revit API needs some work around)
________________________________________________________________
Last Updates:
- [23.06.2026] v1.0 Change Description
________________________________________________________________
Author: Saurabh Tulsankar under guidance of Erik Frits (from LearnRevitAPI.com)"""
# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
from Autodesk.Revit.DB import *

#pyRevit
from pyrevit import forms, script, revit

#.NET Imports
import clr
from rpw.db import curve

clr.AddReference('System')
from System.Collections.Generic import List


# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
doc    = __revit__.ActiveUIDocument.Document #type:Document
uidoc  = __revit__.ActiveUIDocument          # __revit__ is internal variable in pyRevit
app    = __revit__.Application
output = script.get_output()                 # pyRevit Output Menu


# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░

# 1️⃣ Active View must be 3D
active_view = doc.ActiveView

allowed_view_type = [ViewType.ThreeD]
#🚨Ensure active view is 3D
if active_view.ViewType not in allowed_view_type:
    forms.alert("Current Active View is not supported. Please try again in 3D VIEW.", exitscript=True)


# 2️⃣ Select Architecture Categories to Check
cat_options = {
    "Walls": BuiltInCategory.OST_Walls,
    "Floors": BuiltInCategory.OST_Floors,
    "Ceilings": BuiltInCategory.OST_Ceilings,
    "Roofs": BuiltInCategory.OST_Roofs,
    "Columns": BuiltInCategory.OST_Columns,
    "Generic Models": BuiltInCategory.OST_GenericModel}

selected_cat_names = forms.SelectFromList.show(
    sorted(cat_options.keys()),
    title="Select Categories to Check for Painted Faces",
    multiselect=True,
    button_name="Find Painted Elements"
)

if not selected_cat_names:
    forms.alert("No categories selected.", exitscript=True)

# 3️⃣ Get elements from selected categories
all_elems = []

for cat_name in selected_cat_names:
    cat = cat_options[cat_name]

    elems = FilteredElementCollector(doc, active_view.Id)\
            .OfCategory(cat)\
            .WhereElementIsNotElementType()\
            .ToElements()

    all_elems.extend(elems)

#❗For Small Project to check all elements at once.❗
# all_elems = FilteredElementCollector(doc, active_view.Id)\
#             .WhereElementIsNotElementType()\
#             .ToElements()

#🚨Ensure Element in the project
if not all_elems:
    forms.alert("No elements found in active view.", exitscript=True)


#4️⃣Get the Element geometry ( need to do this since paint is on face not element )
opt = Options()
opt.View = active_view

#5️⃣Find Painted surface
painted_elems = []
painted_elem_data = []

for elem in all_elems:
    # Skip element types / invalid elements
    if not elem.Category:
        continue

    geo = elem.get_Geometry(opt)
    if not geo:
        continue

    elem_materials = []
    painted_face_count = 0

    for geo_obj in [g for g in geo if isinstance(g, Solid) and g.Faces.Size > 0]:
        for face in geo_obj.Faces:
            if doc.IsPainted(elem.Id, face):
                painted_face_count += 1

                mat_id = doc.GetPaintedMaterial(elem.Id, face)
                mat = doc.GetElement(mat_id)
                mat_name = Element.Name.GetValue(mat)

                if mat_name not in elem_materials:
                    elem_materials.append(mat_name)

    if elem_materials:
        painted_elems.append(elem)
        painted_elem_data.append([elem, elem_materials])

if not painted_elems:
    forms.alert("No painted Elements found.", exitscript=True)



#6️⃣Create Red Graphic Override
red = Color(255, 0, 0)  #🎨Color

# from tool num: 5 BIMpressionList painter
all_patterns = FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements()
solid_pattern = [i for i in all_patterns if i.GetFillPattern().IsSolidFill][0]

ovg = OverrideGraphicSettings()
ovg.SetProjectionLineColor(red)
ovg.SetSurfaceForegroundPatternColor(red)
ovg.SetSurfaceForegroundPatternId(solid_pattern.Id)

# 7️⃣ Override Painted Surface in Red
t = Transaction(doc, "Highlight Painted Elements")
t.Start()


#🚨Ensure to reset the override if the surface is unpainted
painted_elem_ids = set([elem.Id for elem in painted_elems])
for elem in all_elems:
    if elem.Id not in painted_elem_ids:
        active_view.SetElementOverrides(elem.Id, OverrideGraphicSettings())

# paint the element
for elem in painted_elems:
    active_view.SetElementOverrides(elem.Id, ovg)

t.Commit()


# 8️⃣Painted Element Report
# Create table
table_data = []

for elem, materials in painted_elem_data:
    elem_type =  Element.Name.GetValue(doc.GetElement(elem.GetTypeId()))
    elem_link = output.linkify(elem.Id,"{} : {}".format(elem.Category.Name, elem_type))


    table_data.append([elem_link,", ".join(materials)])

output.print_md("## Paint Peeler Report")
output.print_table(table_data=table_data,
                    columns=["Element", "Paint Material"])  # print table

output.print_md("---")
output.print_md("**Total Painted Elements:** {}".format(len(painted_elems)))



#╔═╗╦═╗╔═╗╔═╗╔═╗  ╔═╗╔═╗  ╔═╗╔═╗╔╗╔╔═╗╔═╗╔═╗╔╦╗
#╠═╝╠╦╝║ ║║ ║╠╣   ║ ║╠╣   ║  ║ ║║║║║  ║╣ ╠═╝ ║
#╩  ╩╚═╚═╝╚═╝╚    ╚═╝╚    ╚═╝╚═╝╝╚╝╚═╝╚═╝╩   ╩
#░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░


# #1️⃣Active view to be 3D
# active_view = doc.ActiveView
# allowed_view_type = [ViewType.ThreeD]
#
# #2️⃣Get all Walls
# all_walls = FilteredElementCollector(doc, active_view.Id).OfClass(Wall).WhereElementIsNotElementType().ToElements()
#
# #3️⃣Get the Wall geometry ( need to do this since paint is on face not element )
# # Element
# #    │
# # get_Geometry(Options)
# #    │
# # GeometryElement
# #    │
# #    ├── Solid
# #    │      ├── Face
# #    │      └── Edge
# #    │
# #    └── GeometryInstance
#
# opt = Options()
# opt.View = active_view
#
# #4️⃣Find Painted surface
# painted_wall_data = []
# painted_walls = []
#
# for wall in all_walls:
#     geo = wall.get_Geometry(opt)
#     if not geo:
#         continue
#
#     wall_materials = []
#
#     for geo_obj in geo:
#         if isinstance(geo_obj, Solid) and geo_obj.Faces.Size > 0:
#            for face in geo_obj.Faces:
#
#                 if doc.IsPainted(wall.Id, face):
#                     mat_id = doc.GetPaintedMaterial(wall.Id, face)
#                     # print(wall.Id,mat_id)
#                     if wall not in painted_walls:
#                         painted_walls.append(wall)
#                     break
#
# #5️⃣Create Red Graphic Override
# # Color
# red = Color(255, 0, 0)
#
# # from tool num: 5 BIMpressionList painter
# all_patterns = FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements()
# solid_pattern = [i for i in all_patterns if i.GetFillPattern().IsSolidFill][0]
#
# ovg = OverrideGraphicSettings()
# ovg.SetProjectionLineColor(red)
# ovg.SetSurfaceForegroundPatternColor(red)
# ovg.SetSurfaceForegroundPatternId(solid_pattern.Id)
#
#
# #6️⃣Make the change
# t = Transaction(doc, "Highlight walls")
# t.Start()
#
# for wall in painted_walls:
#     active_view.SetElementOverrides(wall.Id, ovg)
#
# t.Commit()
# # It works and paints wall Red