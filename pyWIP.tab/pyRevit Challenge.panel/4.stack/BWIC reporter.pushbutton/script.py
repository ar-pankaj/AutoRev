# -*- coding: utf-8 -*-
__title__ = 'BWIC reporter'
__author__ = 'Kamila Milewska'
__doc__ = """Version = 5.0
Date = 24.06.2026
________________________________________________________________
Description:
Generates a linkify and csv report of MEP BWIC (buildersworks) and their host data. Host needs to be in the current document.

Tip: You can save your linkify report and reopen it later.
Tip: Open linkify report in MEP model to select/zoom BWIC elements.
Tip: Bind structural or cladding links, then run the tool to maximize host finding.
________________________________________________________________
How-To:
1. Pick MEP model.
2. Select BWIC category.
3. Select Reference parameter.
4. Select csv export path.
5. Enjoy your reports.
________________________________________________________________
V 2.0 Modified for walls and floors and mechanical equipment BWIC, added CSV export
V 3.0 Added MEP link selection, csv export path selection and progressbar
V 4.0 Added BWIC category and reference parameter selection
V 5.0 Added project type parameters and binding checks"""
#----------------------------------------------------------------------
#outline
#----------------------------------------------------------------------
#pick MEP model
#select BWIC category
#check project parameters and binding
#get main elements - bwic
#create intersection filters
#get intersecting elements and properties
#report linkify and csv
#----------------------------------------------------------------------
#imports
#----------------------------------------------------------------------
import clr
clr.AddReference('System')
from Autodesk.Revit.DB import *
from pyrevit import script, forms, revit
#----------------------------------------------------------------------
#variables
#----------------------------------------------------------------------
doc     = __revit__.ActiveUIDocument.Document
uidoc   = __revit__.ActiveUIDocument
app     = __revit__.Application
output  = script.get_output()  #pyrevit output menu

#parameter names
acoustic_rating_param = 'Acoustic Rating'
floor_fire_rating_param = 'Floor Fire Rating'
#----------------------------------------------------------------------
#FUNCTIONS
#----------------------------------------------------------------------
# Get linked document by id
def get_linked_document_by_id(doc, link_id):
    # Get the RevitLinkInstance element by id
    link_instance = doc.GetElement(link_id)
    if link_instance and isinstance(link_instance, RevitLinkInstance):
        linked_doc = link_instance.GetLinkDocument()
        return linked_doc
    return None

def get_bwic_reference(elem):
    ref = 'No reference'
    try:
        param = elem.LookupParameter(bwic_reference_param)
        if param:
            if param.StorageType == StorageType.String:
                ref = param.AsString()
            elif param.StorageType == StorageType.Double:
                ref = round((UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters)), 2)
            elif param.StorageType == StorageType.Integer:
                ref = param.AsInteger()
            elif param.StorageType == StorageType.ElementId:
                ref = str(param.AsElementId())
    except:
        pass
    return ref

def get_parameter_binding(target_category, param_name):
    binding_map = doc.ParameterBindings
    iterator = binding_map.ForwardIterator()
    iterator.Reset()
    param_found = False
    is_bound_to_category = False
    while iterator.MoveNext():
        definition = iterator.Key
        binding = iterator.Current
        if definition.Name == param_name:
            param_found = True
            if isinstance(binding, TypeBinding):
                bound_categories = binding.Categories
                for category in bound_categories:
                    if int(str(category.Id)) == int(target_category):
                        is_bound_to_category = True
                        break
            break
    if param_found:
        if is_bound_to_category:
            return True
        else:
            forms.alert('Type parameter {} not found for category {}.\nIt will be skipped.'.format(param_name, target_category))
            return False
    else:
        forms.alert("Parameter '{}' not found in project.\nIt will be skipped.".format(param_name))
        return False

def get_wall_fire_rating(host):
    try:
        fire_rating = host.WallType.get_Parameter(BuiltInParameter.FIRE_RATING).AsString()
    except:
        fire_rating = 'Unknown'
    return fire_rating

def get_floor_fire_rating(host, fr_param):
    if fr_param is False:
        fire_rating = '-'
    else:
        try:
            fire_rating = host.FloorType.LookupParameter(floor_fire_rating_param).AsString()
        except:
            fire_rating = 'Unknown'
    return fire_rating

def get_wall_acou_rating(host, ac_param):
    if ac_param is False:
        acou_rating = '-'
    else:
       try:
            acou_rating = host.WallType.LookupParameter(acoustic_rating_param).AsString()
       except:
            acou_rating = 'Unknown'
    return acou_rating

def get_floor_acou_rating(host, ac_param):
    if ac_param is False:
        acou_rating = '-'
    else:
        try:
            acou_rating = host.FloorType.LookupParameter(acoustic_rating_param).AsString()
        except:
            acou_rating = 'Unknown'
    return acou_rating
#----------------------------------------------------------------------
#collectors
#----------------------------------------------------------------------
#pick MEP model
with forms.WarningBar(title='Select MEP link'):
    pick_link_mep = revit.pick_element_by_category(BuiltInCategory.OST_RvtLinks, message='Select MEP link')
if not pick_link_mep:
    forms.alert('No MEP model picked. \nExiting...', exitscript=True)
linked_mep = get_linked_document_by_id(doc, pick_link_mep.Id)
#select bwic category
dict_cats = {'Mechanical Equipment' : BuiltInCategory.OST_MechanicalEquipment,
             'Generic Models' : BuiltInCategory.OST_GenericModel}
from rpw.ui.forms import (FlexForm, Label, ComboBox, Separator, Button)
components = [
                Label('CATEGORY:'), ComboBox('category', dict_cats),
                Separator(),
                Button('Select')]
form = FlexForm('Select BWIC category', components)
form.show()
if not form.values:
    forms.alert('No category selected.', exitscript=True)
input = form.values
cat = input['category']
all_elems = FilteredElementCollector(linked_mep).OfCategory(cat).WhereElementIsNotElementType().ToElements()
if not all_elems:
    forms.alert('No elements of selected category found. \nExiting...', exitscript=True)
#select bwic reference parameter
sel_reference_param = forms.select_parameters(all_elems[0], title='Select BWIC Reference Parameter', button_name='Go!', multiple=False, filterfunc=None, include_instance=True, include_type=False, exclude_readonly=True)
if not sel_reference_param:
    forms.alert('Cancelled!', exitscript=True)
bwic_reference_param = sel_reference_param[0] #[2] is param definition
#sort list by reference
all_elems = sorted(all_elems, key=get_bwic_reference)

#get parameter binding for walls and floors
parameter_ac_w = get_parameter_binding(BuiltInCategory.OST_Walls, acoustic_rating_param)
#wall fire rating is built-in
parameter_ac_f = get_parameter_binding(BuiltInCategory.OST_Floors, acoustic_rating_param)
parameter_fr_f = get_parameter_binding(BuiltInCategory.OST_Floors, floor_fire_rating_param)
#----------------------------------------------------------------------
#main
#----------------------------------------------------------------------
data =[] #csv
linkify_data = []
headers = ['Element Id', 'BWIC Type', 'BWIC reference','Host','Host Type', 'Host Description', 'Host Thickness', 'Fire Rating', 'Acoustic Rating']
data.append(headers)
linkify_headers = ['Element Id', 'BWIC Type', 'BWIC reference','Host Id', 'Host','Host Type', 'Host Description','Host Thickness', 'Fire Rating', 'Acoustic Rating']
linkify_data.append(linkify_headers)

#Progressbar
with forms.ProgressBar(cancellable=True) as pb:  # context for progress bar
    max_value = len(all_elems)

    for count, elem in enumerate(all_elems):
        try:
            BB_elem = elem.get_BoundingBox(None) #default bb for 3d view in coarse mode
            if not BB_elem:
                continue
            outline   = Outline(BB_elem.Min, BB_elem.Max)
            BB_filter = BoundingBoxIntersectsFilter(outline)

            intersect_filter = ElementIntersectsElementFilter(elem)

            intersect_walls = FilteredElementCollector(doc)\
                .OfCategory(BuiltInCategory.OST_Walls)\
                .WhereElementIsNotElementType()\
                .WherePasses(BB_filter)\
                .WherePasses(intersect_filter)\
                .ToElements()

            intersect_floors = FilteredElementCollector(doc)\
                .OfCategory(BuiltInCategory.OST_Floors)\
                .WhereElementIsNotElementType()\
                .WherePasses(BB_filter)\
                .WherePasses(intersect_filter)\
                .ToElements()

            if intersect_walls:
                for wall in intersect_walls:
                    elem_id = elem.Id
                    bwic_type = elem.Name
                    bwic_ref = get_bwic_reference(elem)
                    host = wall.Category.Name
                    host_type = wall.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK).AsString()
                    host_description = wall.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION).AsString()
                    host_thickness = round(UnitUtils.ConvertFromInternalUnits(wall.WallType.Width, UnitTypeId.Millimeters),1)
                    fire_rating = get_wall_fire_rating(wall)
                    acou_rating = get_wall_acou_rating(wall, parameter_ac_w)
                    row = [elem_id, bwic_type, bwic_ref, host, host_type, host_description, host_thickness, fire_rating, acou_rating]
                    data.append(row)

                    #linkify
                    link_elem = output.linkify(elem_id)
                    link_host = output.linkify(wall.Id)
                    linkify_row = [link_elem, bwic_type, bwic_ref, host, link_host, host_type, host_description, host_thickness, fire_rating, acou_rating]
                    linkify_data.append(linkify_row)

            if intersect_floors:
                for floor in intersect_floors:
                    elem_id = elem.Id
                    bwic_type = elem.Name
                    bwic_ref = get_bwic_reference(elem)
                    host = floor.Category.Name
                    host_type = floor.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK).AsString()
                    host_description = floor.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION).AsString()
                    host_thickness = round(UnitUtils.ConvertFromInternalUnits(floor.FloorType.get_Parameter(BuiltInParameter.FLOOR_ATTR_DEFAULT_THICKNESS_PARAM).AsDouble(), UnitTypeId.Millimeters),1)
                    fire_rating = get_floor_fire_rating(floor, parameter_fr_f)
                    acou_rating = get_floor_acou_rating(floor, parameter_ac_f)
                    row = [elem_id, bwic_type, bwic_ref, host, host_type, host_description, host_thickness, fire_rating, acou_rating]
                    data.append(row)

                    # linkify
                    link_elem = output.linkify(elem_id)
                    link_host = output.linkify(floor.Id)
                    linkify_row = [link_elem, bwic_type, bwic_ref, host, link_host, host_type, host_description, host_thickness, fire_rating, acou_rating]
                    linkify_data.append(linkify_row)

            if not intersect_walls and not intersect_floors:
                elem_id = elem.Id
                bwic_type = elem.Name
                bwic_ref = get_bwic_reference(elem)
                host = 'No host found'
                host_type = 'No host found'
                host_description = 'No host found'
                host_thickness = 'No host found'
                fire_rating = 'No host found'
                acou_rating = 'No host found'
                row = [elem_id, bwic_type, bwic_ref, host, host_type, host_description,host_thickness, fire_rating, acou_rating]
                data.append(row)
                # linkify
                link_elem = output.linkify(elem_id)
                link_host = 'No host found'
                linkify_row = [link_elem, bwic_type, bwic_ref, host, link_host, host_type, host_description, host_thickness, fire_rating, acou_rating]
                linkify_data.append(linkify_row)

        except:
            print('Error encountered:')
            import traceback
            print(traceback.print_exc()) #gives error message
            print('Contact developer.')

        #progress bar
        if pb.cancelled:
            forms.alert('Task Cancelled!', exitscript=True)
            break
        else:
            pb.update_progress(count, max_value)
#----------------------------------------------------------------------
#output
#----------------------------------------------------------------------
output.print_table(table_data=linkify_data[1:], title="BWIC Report", columns=linkify_headers)
#and
#WRITE CSV
import csv, os, datetime
picked_path = forms.pick_folder(title='Select base folder for BWIC reports')
if not picked_path:
    forms.alert('CSV report cancelled!', exitscript=True)
timestamp   = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
filename    = 'BWIC_report_{}.csv'.format(timestamp)
path_folder = os.path.join(picked_path, 'BWIC reports')
path_file = os.path.join(path_folder, filename)
# Ensure Folder Exists
if not os.path.exists(path_folder):
    os.makedirs(path_folder)
with open(path_file, 'wb') as csvfile:       # 'wb' for IronPython 2.7
    writer = csv.writer(csvfile, delimiter=',')
    writer.writerows(data)
#print('Exported to: {}'.format(path_file))
#----------------------------------------------------------------------