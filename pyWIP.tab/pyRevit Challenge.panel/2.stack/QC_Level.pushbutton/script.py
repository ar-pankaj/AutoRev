# -*- coding: utf-8 -*-

__tittle__='Check Levels'
__doc__="""Version = 1.1
Date = 02.22.2026
Author: Mateo Lopez"""

__author__          ='Mateo Lopez'
__helpurl__         ='https://www.linkedin.com/in/mateo-lopez/'
__min_revit_ver__   = 2022
__max_revit_ver__   = 2026
__highlight__       = 'new'

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

#pyrevit
from pyrevit import forms, revit, script

# .NET Imports
import clr
clr.AddReference('System')
from System.Collections.Generic import List

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝

doc     = __revit__.ActiveUIDocument.Document   #type: Document
uidoc   = __revit__.ActiveUIDocument            #type: UIDocument
app     = __revit__.Application                 # Application class
output  = script.get_output()

rvt_year     = int(app.VersionNumber)
qc_name_key  = 'QC_ML_Check Levels'

# ╔═╗╦ ╦╔╗╔╔═╗╔╦╗╦╔═╗╔╗╔╔═╗
# ╠╣ ║ ║║║║║   ║ ║║ ║║║║╚═╗
# ╚  ╚═╝╝╚╝╚═╝ ╩ ╩╚═╝╝╚╝╚═╝

def get_level_id(elem):
    level_params = [
        BuiltInParameter.SCHEDULE_LEVEL_PARAM,
        BuiltInParameter.LEVEL_PARAM,
        BuiltInParameter.RBS_START_LEVEL_PARAM
    ]

    for bip in level_params:
        p = elem.get_Parameter(bip)
        if p and p.StorageType == StorageType.ElementId:
            return p.AsElementId()
    return None

def chunk_list(lst, size):
    for i in range(0, len(lst), size):
        yield lst[i:i + size]

# ╔═╗╦  ╔═╗╔═╗╔═╗╔═╗╔═╗
# ║  ║  ╠═╣╚═╗╚═╗║╣ ╚═╗
# ╚═╝╩═╝╩ ╩╚═╝╚═╝╚═╝╚═╝

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
# Get active view
active_view = doc.ActiveView

# Ensure is executed in a 3D view
if not isinstance(active_view, View3D):
    forms.alert(
        'Please run this script from a 3D view.',
        title='Invalid View',
        warn_icon=True,
        exitscript=True
    )

# Collect all elements in view
all_elem = []
collector = FilteredElementCollector(doc, active_view.Id)\
    .WhereElementIsNotElementType()\
    .ToElements()

# Categories excluded — elements with Base/Top Level logic
EXCLUDED_CATEGORIES = [
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_StackedWalls,
    BuiltInCategory.OST_Curtain_Systems,
    BuiltInCategory.OST_CurtainWallPanels,
    BuiltInCategory.OST_CurtainWallMullions,
    BuiltInCategory.OST_Columns,
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_StructuralTruss,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_Roofs,
    BuiltInCategory.OST_Stairs,
    BuiltInCategory.OST_Ramps,
    BuiltInCategory.OST_Planting,
    BuiltInCategory.OST_Entourage,
    BuiltInCategory.OST_Mass,
    BuiltInCategory.OST_Parts,
    BuiltInCategory.OST_Site,
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_RailingSupport,
    BuiltInCategory.OST_VerticalCirculation,
]

excluded_cat_ids = [ElementId(cat) for cat in EXCLUDED_CATEGORIES]

# Filter accurate elements
for elem in collector:
    if elem.Category is None:
        continue
    # Skip elements inside groups
    if elem.GroupId != ElementId.InvalidElementId:
        continue
    # Skip Model In-Place
    if isinstance(elem, FamilyInstance) and elem.Symbol.Family.IsInPlace:
        continue
    # Skip excluded categories
    if elem.Category.Id in excluded_cat_ids:
        continue
    # Only Model Categories
    if elem.Category.CategoryType != CategoryType.Model:
        continue
    # Only Scheduleable Categories
    if not elem.Category.AllowsBoundParameters:
        continue
    # Only Elements with Level parameter
    if get_level_id(elem) is None:
        continue
    all_elem.append(elem)

# Safety check
if not all_elem:
    forms.alert(
        'No elements with level parameter found in active view.',
        title='No elements found',
        warn_icon=True,
        exitscript=True
    )

# All levels
all_levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
if not all_levels:
    forms.alert('No levels found in document.', exitscript=True)
lvl_elev = [round(lvl.Elevation,2) for lvl in all_levels]
min_lvl = all_levels[lvl_elev.index(min(lvl_elev))]

#Compare levels and Z
# ------------------------------------------------------------
correct_lvl = {}

for elem in all_elem:
    elem_bb = elem.get_BoundingBox(None)
    if elem_bb is None:
        continue
    max_bb = elem_bb.Max
    min_bb = elem_bb.Min
    centroid_bb = XYZ(
        (max_bb.X + min_bb.X) / 2.0,
        (max_bb.Y + min_bb.Y) / 2.0,
        (max_bb.Z + min_bb.Z) / 2.0
    )
    elem_z = round(centroid_bb.Z,2)
    # 1. level_elevation <= element_Z
    valid_lvl = [(lvl, elev) for lvl, elev in zip(all_levels, lvl_elev) if elev <= elem_z]
    # 2. If any valid level → pick the highest one
    if valid_lvl:
        assigned_lvl = max(valid_lvl, key=lambda x: x[1])[0]
    else:
        # 3. Else → lowest level
        assigned_lvl = min_lvl
    correct_lvl[elem.Id] = assigned_lvl

#Ensure multi-category schedule doesn't exist
# ------------------------------------------------------------
sch_name = qc_name_key

# Check if schedule already exists
existing_sch = None
existing_schedules  = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()
for v in existing_schedules:
    if v.Name == sch_name:
        existing_sch = v
        break

if existing_sch:
    forms.alert(
        '{} already exists.\nValues will be updated.'.format(sch_name),
        title='Schedule Updated',
    )

# Compare actual level and correct level
# ------------------------------------------------------------
report = {}
report_elems = []

t = Transaction(doc, 'AUBIMAT: Check levels')
t.Start()

try:
    # Clear Comments on all elements of the categories
    for elem in all_elem:
        p_comments = elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if p_comments and p_comments.AsString():
            p_comments.Set('')

    # Set Comments on incorrect elements
    for elem in all_elem:
        if elem.Id not in correct_lvl:
            continue
        lvl = correct_lvl[elem.Id]
        lvl_id = lvl.Id
        level_param_id = get_level_id(elem)
        if level_param_id is None:
            continue  # element has no level parameter; skip
        if lvl_id != level_param_id:
            report_elems.append(elem)
            #Set correct level in comments
            p_comments = elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
            if p_comments:
                p_comments.Set(lvl.Name)
            # --------------------------------------------
            # Data
            cat_name = elem.Category.Name if elem.Category else '-'
            lvl_name = lvl.Name
            if cat_name not in report:
                report[cat_name] = {}
            if lvl_name not in report[cat_name]:
                report[cat_name][lvl_name] = []
            report[cat_name][lvl_name].append(elem)

    # Only create schedule if it doesn't exist
    if not existing_sch:
        #Create multi-category schedule
        # ------------------------------------------------------------
        invalid_id = ElementId.InvalidElementId
        multi_schedule = ViewSchedule.CreateSchedule(doc, invalid_id)
        multi_schedule.Name = sch_name

        # Set Phase Filter to None
        # ------------------------------------------------------------
        phase_filter_param = multi_schedule.get_Parameter(BuiltInParameter.VIEW_PHASE_FILTER)
        if phase_filter_param:
            phase_filter_param.Set(ElementId.InvalidElementId)

        #Add schedule fields
        # ------------------------------------------------------------
        sched_def   = multi_schedule.Definition
        schedulable = sched_def.GetSchedulableFields()

        # BIP map — language independent field identification
        field_bip_map = [
            (BuiltInParameter.ELEM_CATEGORY_PARAM,        'Category'),
            (BuiltInParameter.ELEM_FAMILY_PARAM,           'Family'),
            (BuiltInParameter.ELEM_TYPE_PARAM,             'Type'),
            (BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, 'Comments'),
        ]

        field_category = None
        field_comments = None

        for bip, col_name in field_bip_map:
            bip_id = ElementId(bip)
            for f in schedulable:
                if f.ParameterId == bip_id:
                    added = sched_def.AddField(f)
                    if bip == BuiltInParameter.ELEM_CATEGORY_PARAM:
                        field_category = added
                    elif bip == BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS:
                        field_comments = added
                    break

        # Safety check — if fields were not found, abort
        if field_category is None or field_comments is None:
            raise Exception('Could not find required schedule fields.')

        #Add schedule filter
        # ------------------------------------------------------------
        filter_comments = ScheduleFilter(field_comments.FieldId,ScheduleFilterType.GreaterThan,'')
        sched_def.AddFilter(filter_comments)

        # Clear existing group/sort
        # ------------------------------------------------------------
        sched_def.ClearSortGroupFields()

        # Group & Sort by Category
        cat_sort = ScheduleSortGroupField(field_category.FieldId)
        cat_sort.SortOrder = ScheduleSortOrder.Ascending
        cat_sort.ShowHeader = False
        cat_sort.ShowBlankLine = True
        cat_sort.ShowFooter = True
        cat_sort.ShowFooterTitle = True
        cat_sort.ShowFooterCount = True

        sched_def.AddSortGroupField(cat_sort)

        # Sort by Comments
        # ------------------------------------------------------------
        comments_sort = ScheduleSortGroupField(field_comments.FieldId)
        comments_sort.SortOrder = ScheduleSortOrder.Ascending
        comments_sort.ShowHeader = False
        comments_sort.ShowBlankLine = True
        comments_sort.ShowFooter = False

        sched_def.AddSortGroupField(comments_sort)

        # Grand Total
        # ------------------------------------------------------------
        sched_def.ShowGrandTotal = True
        sched_def.ShowGrandTotalTitle = True
        sched_def.ShowGrandTotalCount = True

    # Create QC 3D view
    # ------------------------------------------------------------
    view_family_types = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
    view_3d_type = None
    for vft in view_family_types:
        if vft.ViewFamily == ViewFamily.ThreeDimensional:
            view_3d_type = vft
            break

    if view_3d_type and report_elems:
        # Check if QC view already exists
        qc_view = None
        existing_3d_views = FilteredElementCollector(doc).OfClass(View3D).ToElements()
        for v in existing_3d_views:
            if not v.IsTemplate and v.Name == qc_name_key:
                qc_view = v
                break

        # Only create if it doesn't exist
        if qc_view is None:
            qc_view = View3D.CreateIsometric(doc, view_3d_type.Id)
            qc_view.Name = qc_name_key
            qc_view.DetailLevel = ViewDetailLevel.Fine
            qc_view.DisplayStyle = DisplayStyle.FlatColors
        # Remove view template
        qc_view.ViewTemplateId = ElementId.InvalidElementId
        # Set Phase Filter to None
        phase_filter_param = qc_view.get_Parameter(BuiltInParameter.VIEW_PHASE_FILTER)
        if phase_filter_param:
            phase_filter_param.Set(ElementId.InvalidElementId)

        # Hide ALL categories (model + annotation)
        all_categories = doc.Settings.Categories
        for cat in all_categories:
            if qc_view.CanCategoryBeHidden(cat.Id):
                try:
                    qc_view.SetCategoryHidden(cat.Id, True)
                except:
                    pass
            # Subcategories
            for subcat in cat.SubCategories:
                if qc_view.CanCategoryBeHidden(subcat.Id):
                    try:
                        qc_view.SetCategoryHidden(subcat.Id, True)
                    except:
                        pass

        # Unhide only categories of incorrect elements
        incorrect_cat_ids = set([elem.Category.Id for elem in report_elems if elem.Category])
        for cat in all_categories:
            if cat.Id in incorrect_cat_ids:
                try:
                    qc_view.SetCategoryHidden(cat.Id, False)
                except:
                    pass
                # Also unhide all subcategories of this category
                for subcat in cat.SubCategories:
                    if qc_view.CanCategoryBeHidden(subcat.Id):
                        try:
                            qc_view.SetCategoryHidden(subcat.Id, False)
                        except:
                            pass

        # Unhide levels specifically
        levels_cat_id = ElementId(BuiltInCategory.OST_Levels)
        try:
            qc_view.SetCategoryHidden(levels_cat_id, False)
        except:
            pass

        # Isolate only incorrect elements + levels within their visible categories
        level_collector = FilteredElementCollector(doc).OfClass(Level).ToElementIds()

        all_ids_to_isolate = List[ElementId]([elem.Id for elem in report_elems])
        for level_id in level_collector:
            all_ids_to_isolate.Add(level_id)

        qc_view.IsolateElementsTemporary(all_ids_to_isolate)

    t.Commit()

except Exception as e:
    if t.HasStarted():
        t.RollBack()
    forms.alert('Error: {}'.format(str(e)), title='Script Failed — No Changes Made', exitscript=True)

# ------------------------------------------------------------
# Print grouped report
# ------------------------------------------------------------
total = 0
CHUNK_SIZE = 90

for cat_name in sorted(report.keys()):
    print('=' * 50)
    print('CATEGORY: {}'.format(cat_name))
    for lvl_name in sorted(report[cat_name].keys()):
        elems    = report[cat_name][lvl_name]
        ids      = [e.Id for e in elems]
        if len(ids) <= CHUNK_SIZE:
            print('  -> Level: {} ({} elements)  {}'.format(
                lvl_name,
                len(elems),
                output.linkify(ids, title='Select {} {}'.format(len(elems), cat_name))
            ))
        else:
            for chunk in chunk_list(ids, CHUNK_SIZE):
                print('  -> Level: {} ({} elements)  {}'.format(
                    lvl_name,
                    len(chunk),
                    output.linkify(chunk, title='Select {} {}'.format(len(chunk), cat_name))
                ))

        total += len(elems)

print('=' * 50)

if total == 0:
    print('Congrats! All elements are assigned to the correct level.')
else:
    print('Total elements to fix: {}'.format(total))
#----------------------------------------------------
print('-'*50)
print('Its finished')
print('Script has been developed by Mateo Lopez')