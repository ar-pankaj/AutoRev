# -*- coding: utf-8 -*-
__title__   = "Export Sheets to pdf/dwg"
__doc__     = """Version = 1.1
Date    = 2026-01-19
______________________________________________________________
Description:

Exports selected Sheets to PDF and/or DWG.

- User selects formats
- User selects export path or model path
- Automatic overwrite detection
- Export only new sheets option
- Automatic folder creation
- Filename generation based on parameters
______________________________________________________________
Author: Kees Groenendijk
"""

# ------------------------------------------------------------
# IMPORTS
# ------------------------------------------------------------

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog

from pyrevit import forms, DB, script

import os
import clr

clr.AddReference('System')
from System.Collections.Generic import List


# ------------------------------------------------------------
# VARIABLES
# ------------------------------------------------------------

doc    = __revit__.ActiveUIDocument.Document #type:Document
uidoc  = __revit__.ActiveUIDocument          # __revit__ is internal variable in pyRevit
app    = __revit__.Application
output = script.get_output()                 # pyRevit Output Menu

# Collect all title blocks
titleblocks = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_TitleBlocks)\
    .WhereElementIsNotElementType()\
    .ToElements()

tb_dict = {tb.OwnerViewId: tb for tb in titleblocks}

# ------------------------------------------------------------
# PARAMETER HELPERS
# ------------------------------------------------------------

def get_param(elem, name, default= ""): # returns the value of a parameter, or default if not found or no value

    p = elem.LookupParameter(name)
    if not p:
        return default

    if not p.HasValue:
        return default

    if p.StorageType == StorageType.String:
        return p.AsString()

    if p.StorageType == StorageType.Integer:
        return p.AsInteger()

    if p.StorageType == StorageType.Double:
        return p.AsDouble()

    if p.StorageType == StorageType.ElementId:
        return p.AsElementId()

    return default


def sanitize_filename(name): # removes invalid characters from filename and replaces spaces with underscores
    for c in '<>:"/\\|?*':
        name = name.replace(c, "_")
    return name.replace(" ", "_")

rvt_year = int(app.VersionNumber)

# def get_titleblock(sheet):
#     col = (
#         FilteredElementCollector(doc, sheet.Id)
#         .OfCategory(BuiltInCategory.OST_TitleBlocks)
#         .WhereElementIsNotElementType()
#     )
#     return col.FirstElement()

def sheets_to_id_list(sheets): # converts a list of sheets to a .NET List[ElementId]
    """Convert sheets to .NET List[ElementId]"""
    ids = List[ElementId]()
    for s in sheets:
        ids.Add(s.Id)
    return ids


def get_titleblock(sheet): # returns the titleblock element of a sheet
    return tb_dict.get(sheet.Id)

# ------------------------------------------------------------
# FILENAME GENERATOR
# ------------------------------------------------------------

project_num = doc.ProjectInformation.Number or "Pxxxxxx"

def generate_filename(sheet): # parameters are hardcoded in the function for convenience, can be changed if needed, This is specified to our titlelock family, if you use a different titleblock family, you may need to adjust the parameter names accordingly.

    parts = [project_num]

    parts.append(get_param(sheet, "HA_ALG_fase"))
    parts.append(get_param(sheet, "HA_ALG_discipline"))

    bwdl = get_param(sheet, "HA_ALG_bouwdeel")

    tb = get_titleblock(sheet)
    if tb: # if titleblock exists, check if the builiding part parameter is set to 1, if so, add the "HA_ALG_bouwdeel" parameter to the filename
        visible = tb.LookupParameter("HA_ALG_bouwdeel_visible")
        if visible and visible.HasValue:
            vis = visible.AsInteger()
            if vis == 1:
                parts.append(bwdl)

    parts.append(get_param(sheet, "HA_ALG_documenttype"))

    man_nr = get_param(sheet, "HA_ALG_tekeningnummer") #depends on your titleblock family, if you use a different titleblock family, you may need to adjust the parameter name accordingly.
    parts.append(man_nr or get_param(sheet, "Sheet Number"))

    rev = get_param(sheet, "Current Revision")
    if rev:
        parts.append(rev)

    return sanitize_filename("_".join(parts))


# ------------------------------------------------------------
# EXISTING FILE CHECK
# ------------------------------------------------------------

def collect_existing(folder, extensions): # returns a set of existing filenames (without extensions) in the given folder with the given extensions

    existing = set()

    if not os.path.exists(folder):
        return existing

    for f in os.listdir(folder):
        name, ext = os.path.splitext(f)
        if ext.lower() in extensions:
            existing.add(name)

    return existing


# ------------------------------------------------------------
# USER INPUT
# ------------------------------------------------------------

all_sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()

# Build dictionary for all sheets with their collection as key (or "Uncategorized" if no collection)
sheet_dict = {}
if rvt_year > 2024:
    for sheet in all_sheets:
        sheet_collection_id = sheet.SheetCollectionId if sheet.SheetCollectionId != ElementId.InvalidElementId else None
        sheet_collection = doc.GetElement(sheet_collection_id).Name if sheet_collection_id else "No Sheet Collection"
        key = "{} - {}_{}".format(sheet_collection, sheet.SheetNumber, sheet.Name)
        sheet_dict[key] = sheet

else:
    for sheet in all_sheets:
        key = "{}_{}".format(sheet.SheetNumber, sheet.Name)
        sheet_dict[key] = sheet

selected_sheet_keys = forms.SelectFromList.show(
    sorted(sheet_dict.keys()),
    title="Select Sheets to Export",
    multiselect=True
)

selected_sheets = [sheet_dict[key] for key in selected_sheet_keys] if selected_sheet_keys else []


if not selected_sheets:
    script.exit()

formats = forms.SelectFromList.show(
    ["PDF", "DWG", "Combine Sheets"],
    title="Select formats",
    multiselect=True,
    button_name="Continue"
)


if not formats:
    script.exit()

sheet_combine = False
if "Combine Sheets" in formats:
    sheet_combine = True

path_choice = forms.ask_for_one_item(
    ['Kies een bestandspad', 'Gebruik bestandspad Model'],
    default='Gebruik bestandspad Model',
    title='Bestandspad'
)

# ------------------------------------------------------------
# EXPORT PATHS
# ------------------------------------------------------------

if path_choice == 'Kies een bestandspad':
    base_dir = forms.pick_folder("Select export folder")
    if not base_dir:
        script.exit()
else:
    base_dir = os.path.dirname(doc.PathName)

pdf_dir = base_dir # you can change this to a different folder if needed, by default it will export to the same folder as the model, or the user selected folder
dwg_dir = os.path.join(base_dir, "dwg")


# ------------------------------------------------------------
# OPTIONS
# ------------------------------------------------------------

# set PDF export options

pdf_options = PDFExportOptions()
pdf_options.Combine = True
pdf_options.HideCropBoundaries = True
pdf_options.HideReferencePlane = True
pdf_options.HideScopeBoxes = True
pdf_options.HideUnreferencedViewTags = True
pdf_options.ExportQuality = PDFExportQualityType.DPI300
pdf_options.ZoomType = ZoomType.Zoom
pdf_options.ZoomPercentage = 100
pdf_options.PaperOrientation = PageOrientationType.Auto
#pdf_options.RasterQuality.High

dwg_options = DWGExportOptions.GetPredefinedOptions(doc, "CAS30") # 'CAS30' is the default export setup in Revit, you can change this to another setup if needed
dwg_options.FileVersion = ACADVersion.R2013


# ------------------------------------------------------------
# CONFLICT CHECK
# ------------------------------------------------------------

existing_conflicts = set()

if "PDF" in formats:
    pdf_existing = collect_existing(pdf_dir, [".pdf"])
else:
    pdf_existing = set()

if "DWG" in formats:
    dwg_existing = collect_existing(dwg_dir, [".dwg"])
else:
    dwg_existing = set()

if sheet_combine and not "DWG" in formats:
    fname = generate_filename(selected_sheets[0])
    if fname in pdf_existing:
        existing_conflicts.add(fname)
else:
    for sheet in selected_sheets:
        fname = generate_filename(sheet)

        if fname in pdf_existing or fname in dwg_existing:
            existing_conflicts.add(fname)


if existing_conflicts:

    proceed = forms.alert(
        "{}\n\nBestanden bestaan al.\nWat wil je doen?".format(
            "\n".join(existing_conflicts)
        ),
        options=[
            "Overwrite",
            "Export only new sheets",
            "Cancel"
        ]
    )

    if proceed == "Cancel":
        script.exit()

    if proceed == "Export only new sheets":

        filtered = []
        for sheet in selected_sheets:
            if generate_filename(sheet) not in existing_conflicts:
                filtered.append(sheet)

        selected_sheets = filtered


# ------------------------------------------------------------
# CREATE FOLDERS
# ------------------------------------------------------------

if "PDF" in formats and not os.path.exists(pdf_dir):
    os.makedirs(pdf_dir)

if "DWG" in formats and not os.path.exists(dwg_dir):
    os.makedirs(dwg_dir)


# ------------------------------------------------------------
# EXPORT
# ------------------------------------------------------------

success = 0
failed  = []



if "PDF" in formats and sheet_combine:
    try:
        ids = sheets_to_id_list(selected_sheets)
        pdf_options.FileName = generate_filename(selected_sheets[0])
        doc.Export(pdf_dir, ids, pdf_options)
        success += len(selected_sheets)

    except Exception as e:
        failed.append("{} -> {}".format(sheet.Name, e))

for sheet in selected_sheets:

    fname = generate_filename(sheet)

    ids = List[ElementId]()
    ids.Add(sheet.Id)

    pdf_options.FileName = fname

    try:

        if "PDF" in formats and not sheet_combine:
            doc.Export(pdf_dir, ids, pdf_options)

        if "DWG" in formats:
            doc.Export(dwg_dir, fname, ids, dwg_options)

        success += 1

    except Exception as e:
        failed.append("{} -> {}".format(sheet.Name, e))



# ------------------------------------------------------------
# RESULT
# ------------------------------------------------------------

TaskDialog.Show(
    "Export Result",
    "Succesvol geëxporteerd: {}\n\nFouten:\n{}".format(
        success,
        "\n".join(failed) if failed else "-"
    )
)

