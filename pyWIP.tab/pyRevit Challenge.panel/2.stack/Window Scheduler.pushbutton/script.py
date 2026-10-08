# -*- coding: utf-8 -*-
__title__   = "Window Scheduler"
__doc__     = """Version = 1.1
Date    = 27.06.2026
________________________________________________________________
Description:
Produces 'correct' Commonwealth grid-style window schedules

________________________________________________________________
How-To:
1. Choose Phase, Scale, Layout, Titleblock, and Viewport Style
2. Choose Windows to schedule
3. Window Scheduler will generate a series of plan and elevation views of each window, tag them and place them on a sheet

________________________________________________________________
Author: Peter le Roux"""

from Autodesk.Revit.DB import *
from pyrevit import forms, script, DB
from rpw.ui.forms import FlexForm, Label, ComboBox, Separator, Button
import clr
clr.AddReference('System')
from System.Collections.Generic import List

doc    = __revit__.ActiveUIDocument.Document 
uidoc  = __revit__.ActiveUIDocument          
app    = __revit__.Application
output = script.get_output()                 

# =========================================================
# GLOBAL OPTIMIZATIONS & CACHING
# =========================================================

# Cache all existing view and sheet names once at startup to prevent O(N²) filter operations
existing_view_names = set()
all_views_collector = DB.FilteredElementCollector(doc).OfClass(DB.View)
for v in all_views_collector:
    try:
        if v.Name:
            existing_view_names.add(v.Name)
    except:
        pass

# =========================================================
# HELPERS
# =========================================================

def get_safe_type_name(element_type):
    name_param = element_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
    return name_param.AsString() if name_param else "Unknown Type"

def get_unique_view_name(desired_name, names_set):
    unique_name = desired_name
    counter = 1
    
    while unique_name in names_set:
        unique_name = "{} ({})".format(desired_name, counter)
        counter += 1
        
    names_set.add(unique_name) # Dynamic allocation updates track internal uniqueness instantly
    return unique_name

# =========================================================
# 1 & 2: UI AND SELECTION 
# =========================================================

# --- Collect Project Phases ---
phase_options = {phase.Name: phase.Name for phase in doc.Phases}
if not phase_options:
    print("No phases found in this document.")
    import sys; sys.exit()

# --- Collect Viewport Types Safely ---
# --- Collect Viewport Types Safely ---
vp_options = {}

all_types = DB.FilteredElementCollector(doc).WhereElementIsElementType().ToElements()

for t in all_types:
    # Check if it possesses the Viewport Title parameter
    has_title_param = t.get_Parameter(DB.BuiltInParameter.VIEWPORT_ATTR_SHOW_LABEL)
    
    if has_title_param is not None:
        type_name = "Unknown"
        name_param = t.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
        
        if name_param and name_param.AsString():
            type_name = name_param.AsString()
        else:
            sym_param = t.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if sym_param and sym_param.AsString():
                type_name = sym_param.AsString()
            else:
                try:
                    if t.Name: type_name = t.Name
                except:
                    type_name = "Viewport Style " + str(t.Id)
        
        vp_options[type_name] = type_name

if not vp_options:
    vp_options = {"No Viewport Types Loaded": None}

# --- Collect Window Tags ---
tag_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_WindowTags).WhereElementIsElementType().ToElements()
if tag_collector:
    tag_options = {"{} - {}".format(t.FamilyName, get_safe_type_name(t)): t for t in tag_collector}
else:
    tag_options = {"No Window Tags Loaded": None}

# --- Collect TitleBlocks ---
tb_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()
if tb_collector:
    tb_options = {"{} - {}".format(tb.FamilyName, get_safe_type_name(tb)): tb for tb in tb_collector}
else:
    tb_options = {"No TitleBlocks Loaded": None}

# --- Build the UI Form ---
components = [
    Label('Select Project Phase:'),
    ComboBox('phase_name', phase_options),
    Separator(),
    
    Label('Select View Scale:'),
    ComboBox('scale', {'1:10': 10, '1:20': 20, '1:50': 50}),
    
    Label('Select Sheet Layout:'),
    ComboBox('layout', {'Single window per sheet': 'single', 'Multiple windows per sheet': 'multiple'}),
    
    Label('Select Viewport Style (No Title):'),
    ComboBox('viewport_type', vp_options),
    
    Label('Select Window Tag Type:'),
    ComboBox('window_tag', tag_options),
    
    Label('Select TitleBlock:'),
    ComboBox('titleblock', tb_options),
    
    Separator(),
    Button('Continue to Window Selection')
]

form = FlexForm('Schedule Settings', components)
form.show()

# --- Extract values and resolve strings to IDs ---
if form.values:
    chosen_phase_name = form.values['phase_name']
    chosen_scale = form.values['scale']
    chosen_layout = form.values['layout']
    chosen_vp_name = form.values['viewport_type']
    chosen_tag = form.values['window_tag']
    chosen_titleblock = form.values['titleblock'] 
    
    if not chosen_phase_name or not chosen_titleblock or not chosen_vp_name:
        print("Error: Missing required UI selections.")
        import sys; sys.exit()
        
    selected_phase = next((p for p in doc.Phases if p.Name == chosen_phase_name), None)
    selected_vp_type_id = None
    for t in all_types:
        if t.get_Parameter(DB.BuiltInParameter.VIEWPORT_ATTR_SHOW_LABEL) is not None:
            name_param = t.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
            safe_name = name_param.AsString() if name_param and name_param.AsString() else ""
            if not safe_name:
                try: safe_name = t.Name
                except: pass
                
            if safe_name == chosen_vp_name:
                selected_vp_type_id = t.Id
                break
    
    if not selected_phase or not selected_vp_type_id:
        print("Error: Could not resolve elements from UI.")
        import sys; sys.exit()
else:
    print("Script cancelled by user.")
    import sys; sys.exit()

# ---------------------------------------------------------
# Filter Windows & Ask User to Select
# ---------------------------------------------------------
param_id = DB.ElementId(DB.BuiltInParameter.PHASE_CREATED)
rule = DB.ParameterFilterRuleFactory.CreateEqualsRule(param_id, selected_phase.Id)
phase_filter = DB.ElementParameterFilter(rule)

windows_in_phase = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_Windows)\
    .WhereElementIsNotElementType()\
    .WherePasses(phase_filter)\
    .ToElements()

if not windows_in_phase:
    print("No windows were created in the '{}' phase.".format(selected_phase.Name))
    import sys; sys.exit()

unique_window_types = {}
for w in windows_in_phase:
    type_id = w.GetTypeId()
    if type_id not in unique_window_types:
        unique_window_types[type_id] = w

class UniqueWindowOption(object):
    def __init__(self, window):
        self.window = window
        type_name = window.Name
        window_type = doc.GetElement(window.GetTypeId())
        family_name = window_type.FamilyName
        
        type_mark_param = window_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK)
        type_mark = type_mark_param.AsString() if type_mark_param and type_mark_param.AsString() else "No mark"
        self.display_name = "[{}] {} - {}".format(type_mark, family_name, type_name)

window_options = [UniqueWindowOption(w) for w in unique_window_types.values()]
window_options.sort(key=lambda x: x.display_name)

selected_options = forms.SelectFromList.show(
    window_options,
    name_attr='display_name',
    title='Select windows to schedule',
    button_name='Generate Schedules',
    multiselect=True
)

if selected_options:
    selected_windows = [opt.window for opt in selected_options]
else:
    print("Window selection cancelled.")
    import sys; sys.exit()

# =========================================================
# TRANSACTION 1: GENERATE VIEWS
# =========================================================

view_family_types = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType).ToElements()
detail_view_type = next((v for v in view_family_types if v.ViewFamily == DB.ViewFamily.Detail), None)
floor_plan_type = next((v for v in view_family_types if v.ViewFamily == DB.ViewFamily.FloorPlan), None)

if not detail_view_type or not floor_plan_type:
    print("Error: Could not find dynamic View Family Type configurations in the model.")
    import sys; sys.exit()

# --- Pre-Collect Levels for Slab Heights ---
all_levels = DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements()
sorted_levels = sorted(all_levels, key=lambda l: l.Elevation)

view_pairs = []
t = DB.Transaction(doc, "Generate Window Schedule Views")
t.Start()

for window in selected_windows:
    wall = window.Host
    if not wall:
        continue
        
    window_type = doc.GetElement(window.GetTypeId())
    width_param = window_type.get_Parameter(DB.BuiltInParameter.WINDOW_WIDTH)
    height_param = window_type.get_Parameter(DB.BuiltInParameter.WINDOW_HEIGHT)
    
    w_width = width_param.AsDouble() if width_param else 3.0
    w_height = height_param.AsDouble() if height_param else 3.0
    
    mark_param = window_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK)
    window_mark = mark_param.AsString() if mark_param and mark_param.AsString() else "W_Unknown"

    loc_pt = window.Location.Point
    ext_normal = wall.Orientation 
    inside_normal = ext_normal.Multiply(-1)
    offset = 1.5 

    # --- CREATE INSIDE ELEVATION ---
    elev_t = DB.Transform.Identity
    elev_t.Origin = loc_pt
    elev_t.BasisZ = inside_normal                 
    elev_t.BasisY = DB.XYZ.BasisZ                 
    elev_t.BasisX = elev_t.BasisY.CrossProduct(elev_t.BasisZ).Normalize() 

    level_id = window.LevelId if window.LevelId != DB.ElementId.InvalidElementId else wall.LevelId
    base_level = doc.GetElement(level_id)
    base_elev_z = base_level.Elevation if base_level else loc_pt.Z

    top_elev_z = None
    for lvl in sorted_levels:
        if lvl.Elevation > base_elev_z + 0.1:
            top_elev_z = lvl.Elevation
            break
            
    if top_elev_z is None:
        top_elev_z = base_elev_z + 10.0 

    y_min = (base_elev_z - loc_pt.Z) - 0.5
    y_max = (top_elev_z - loc_pt.Z) + 0.5

    elev_bb = DB.BoundingBoxXYZ()
    elev_bb.Transform = elev_t
    elev_bb.Min = DB.XYZ(-w_width/2 - offset, y_min, -offset)
    elev_bb.Max = DB.XYZ(w_width/2 + offset, y_max, offset)

    elev_view = DB.ViewSection.CreateDetail(doc, detail_view_type.Id, elev_bb)
    elev_view.Scale = chosen_scale
    elev_view.DetailLevel = DB.ViewDetailLevel.Fine
    
    # --- CREATE TRUE FLOOR PLAN ---
    plan_view = DB.ViewPlan.Create(doc, floor_plan_type.Id, level_id)
    plan_view.Scale = chosen_scale
    plan_view.DetailLevel = DB.ViewDetailLevel.Fine

    view_range = plan_view.GetViewRange()
    sill_param = window.get_Parameter(DB.BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM)
    sill_height = sill_param.AsDouble() if sill_param else 0.0
    cut_offset = sill_height + (w_height / 2.0)

    view_range.SetOffset(DB.PlanViewPlane.CutPlane, cut_offset)
    view_range.SetOffset(DB.PlanViewPlane.TopClipPlane, cut_offset + 1.0)
    view_range.SetOffset(DB.PlanViewPlane.BottomClipPlane, sill_height - 0.5)
    view_range.SetOffset(DB.PlanViewPlane.ViewDepthPlane, sill_height - 0.5)
    plan_view.SetViewRange(view_range)

    plan_bb = DB.BoundingBoxXYZ()
    plan_bb.Min = DB.XYZ(loc_pt.X - w_width/2 - offset, loc_pt.Y - offset, loc_pt.Z - 10.0)
    plan_bb.Max = DB.XYZ(loc_pt.X + w_width/2 + offset, loc_pt.Y + offset, loc_pt.Z + 10.0)

    plan_view.CropBoxActive = True
    plan_view.CropBox = plan_bb
    plan_view.CropBoxVisible = False
    
    doc.Regenerate()
    visible_without_crop = DB.FilteredElementCollector(doc, plan_view.Id).ToElementIds()
    
    plan_view.CropBoxVisible = True
    doc.Regenerate()
    collector = DB.FilteredElementCollector(doc, plan_view.Id)
    if visible_without_crop.Count > 0:
        collector.Excluding(visible_without_crop)
        
    crop_element = collector.FirstElement()
    if crop_element:
        angle = DB.XYZ.BasisY.AngleTo(ext_normal)
        cross = DB.XYZ.BasisY.CrossProduct(ext_normal)
        if cross.Z < 0: angle = -angle
        if angle != 0.0:
            axis = DB.Line.CreateBound(loc_pt, loc_pt.Add(DB.XYZ.BasisZ))
            DB.ElementTransformUtils.RotateElement(doc, crop_element.Id, axis, angle)

    plan_view.CropBoxVisible = False

    if chosen_tag:
        window_ref = DB.Reference(window)
        tag_offset = 2.0 
        tag_position = loc_pt.Add(inside_normal.Multiply(tag_offset))
        
        # --- Cross-Version API Compatibility (Revit 2022 - 2026) ---
        try:
            # Revit 2024+ (6 arguments: TagMode removed)
            new_tag = DB.IndependentTag.Create(doc, plan_view.Id, window_ref, False, DB.TagOrientation.Horizontal, tag_position)
        except TypeError:
            # Revit 2023 and older (7 arguments: Requires TagMode)
            new_tag = DB.IndependentTag.Create(doc, plan_view.Id, window_ref, False, DB.TagMode.TM_ADDBY_CATEGORY, DB.TagOrientation.Horizontal, tag_position)

        if new_tag:
            new_tag.ChangeTypeId(chosen_tag.Id)
            
    # Standardized Naming Conventions using fast set verification
    base_elev_name = "SCHEDULE {} ELEVATION".format(window_mark)
    elev_view.Name = get_unique_view_name(base_elev_name, existing_view_names)
        
    base_plan_name = "SCHEDULE {} PLAN".format(window_mark)
    plan_view.Name = get_unique_view_name(base_plan_name, existing_view_names)
        
    view_pairs.append({
        'elev_view': elev_view,
        'plan_view':  plan_view,
        'mark':       window_mark
    })

t.Commit()
print("Successfully generated {} view pairs.".format(len(view_pairs)))

# =========================================================
# TRANSACTION 2: PLACE VIEWS ON SHEETS
# =========================================================

H_GAP   = 0.08   
COL_GAP = 0.12   
ROW_GAP = 0.15   
MARGIN  = 0.15   

t2 = DB.Transaction(doc, "Place Window Schedule on Sheets")
t2.Start()

real_sheets = []

phase_filters = DB.FilteredElementCollector(doc).OfClass(DB.PhaseFilter).ToElements()
show_complete_filter = next((pf for pf in phase_filters if "complete" in pf.Name.lower()), None)

for vp_data in view_pairs:
    for view in [vp_data['elev_view'], vp_data['plan_view']]:
        try:
            view.get_Parameter(DB.BuiltInParameter.VIEW_PHASE).Set(selected_phase.Id)
            if show_complete_filter:
                view.get_Parameter(DB.BuiltInParameter.VIEW_PHASE_FILTER).Set(show_complete_filter.Id)
        except:
            pass 

if chosen_layout == 'single':
    for vp_data in view_pairs:
        elev_view  = vp_data['elev_view']
        plan_view  = vp_data['plan_view']
        window_mark = vp_data['mark']

        sheet = DB.ViewSheet.Create(doc, chosen_titleblock.Id)
        real_sheets.append(sheet)
        
        base_sheet_name = "Window Schedule - Type {}".format(window_mark)
        sheet.Name = get_unique_view_name(base_sheet_name, existing_view_names)

        vp_elev = DB.Viewport.Create(doc, sheet.Id, elev_view.Id, DB.XYZ.Zero)
        vp_elev.ChangeTypeId(selected_vp_type_id)
            
        vp_plan = DB.Viewport.Create(doc, sheet.Id, plan_view.Id,  DB.XYZ.Zero)
        vp_plan.ChangeTypeId(selected_vp_type_id)
            
        doc.Regenerate()

        elev_outline = vp_elev.GetBoxOutline()
        plan_outline = vp_plan.GetBoxOutline()

        elev_h = elev_outline.MaximumPoint.Y - elev_outline.MinimumPoint.Y
        plan_h = plan_outline.MaximumPoint.Y  - plan_outline.MinimumPoint.Y

        tb_inst = DB.FilteredElementCollector(doc, sheet.Id).OfCategory(DB.BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().FirstElement()
        if tb_inst:
            tb_bb  = tb_inst.get_BoundingBox(sheet)
            center_u = (tb_bb.Max.X + tb_bb.Min.X) / 2.0
            center_v = (tb_bb.Max.Y + tb_bb.Min.Y) / 2.0
        else:
            center_u, center_v = 1.39, 0.97

        total_h   = elev_h + H_GAP + plan_h
        group_top = center_v + total_h / 2.0

        vp_elev.SetBoxCenter(DB.XYZ(center_u, group_top - elev_h / 2.0, 0))
        vp_plan.SetBoxCenter(DB.XYZ(center_u, group_top - elev_h - H_GAP - plan_h / 2.0, 0))

elif chosen_layout == 'multiple':
    ref_sheet = DB.ViewSheet.Create(doc, chosen_titleblock.Id)
    ref_sheet.Name = "_TEMP MEASURE - DELETE ME"

    tb_inst = DB.FilteredElementCollector(doc, ref_sheet.Id).OfCategory(DB.BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().FirstElement()
    if tb_inst:
        tb_bb  = tb_inst.get_BoundingBox(ref_sheet)
        bounds = {
            'left':   tb_bb.Min.X + MARGIN,
            'right':  tb_bb.Max.X - MARGIN,
            'top':    tb_bb.Max.Y - MARGIN,
            'bottom': tb_bb.Min.Y + MARGIN,
        }
    else:
        bounds = {'left': 0.15, 'right': 2.78, 'top': 1.63, 'bottom': 0.15}

    temp_vps = []
    for vp_data in view_pairs:
        vp_e = DB.Viewport.Create(doc, ref_sheet.Id, vp_data['elev_view'].Id, DB.XYZ.Zero)
        vp_p = DB.Viewport.Create(doc, ref_sheet.Id, vp_data['plan_view'].Id,  DB.XYZ.Zero)
        temp_vps.append((vp_e, vp_p))
        
    doc.Regenerate()  

    cell_data = []
    for i, (vp_e, vp_p) in enumerate(temp_vps):
        eo = vp_e.GetBoxOutline()
        po = vp_p.GetBoxOutline()

        ew = eo.MaximumPoint.X - eo.MinimumPoint.X
        eh = eo.MaximumPoint.Y - eo.MinimumPoint.Y
        pw = po.MaximumPoint.X - po.MinimumPoint.X
        ph = po.MaximumPoint.Y - po.MinimumPoint.Y

        doc.Delete(vp_e.Id)
        doc.Delete(vp_p.Id)

        cell_data.append({
            'elev_view': view_pairs[i]['elev_view'],
            'plan_view':  view_pairs[i]['plan_view'],
            'mark':       view_pairs[i]['mark'],
            'ew': ew, 'eh': eh,
            'pw': pw, 'ph': ph,
            'cell_w': max(ew, pw),          
            'cell_h': eh + H_GAP + ph,      
        })

    doc.Delete(ref_sheet.Id)  
    doc.Regenerate()

    rows = []
    current_row = []
    cur_x = bounds['left']
    
    for cell in cell_data:
        cw = cell['cell_w']
        if cur_x + cw > bounds['right'] and current_row:
            rows.append(current_row)
            current_row = [cell]
            cur_x = bounds['left'] + cw + COL_GAP
        else:
            current_row.append(cell)
            cur_x += cw + COL_GAP
            
    if current_row:
        rows.append(current_row)

    assignments  = []   
    sheet_count  = 0
    cur_y        = bounds['top']
    
    for row in rows:
        max_eh = max([c['eh'] for c in row])
        max_ph = max([c['ph'] for c in row])
        row_h = max_eh + H_GAP + max_ph
        
        if cur_y - row_h < bounds['bottom']:
            sheet_count += 1
            cur_y = bounds['top']
            
        cur_x = bounds['left']
        
        for cell in row:
            cw = cell['cell_w']
            center_x = cur_x + cw / 2.0
            
            elev_bottom_y = cur_y - max_eh
            elev_cy = elev_bottom_y + cell['eh'] / 2.0
            
            plan_top_y = elev_bottom_y - H_GAP
            plan_cy = plan_top_y - cell['ph'] / 2.0
            
            assignments.append((sheet_count, cell, center_x, elev_cy, plan_cy))
            cur_x += cw + COL_GAP
            
        cur_y -= row_h + ROW_GAP

    for i in range(sheet_count + 1):
        s = DB.ViewSheet.Create(doc, chosen_titleblock.Id)
        base_sheet_name = "Window Schedule - Sheet {}".format(i + 1)
        s.Name = get_unique_view_name(base_sheet_name, existing_view_names)
        real_sheets.append(s)

    for sid, cell, cx, elev_cy, plan_cy in assignments:
        target = real_sheets[sid]
        try:
            new_vp_elev = DB.Viewport.Create(doc, target.Id, cell['elev_view'].Id, DB.XYZ(cx, elev_cy, 0))
            new_vp_elev.ChangeTypeId(selected_vp_type_id)
                
            new_vp_plan = DB.Viewport.Create(doc, target.Id, cell['plan_view'].Id,  DB.XYZ(cx, plan_cy,  0))
            new_vp_plan.ChangeTypeId(selected_vp_type_id)
                
        except Exception as e:
            print("Could not place '{}' on sheet. Error: {}".format(cell['mark'], e))

t2.Commit()

# =========================================================
# 7 FINAL REPORT
# =========================================================
output.print_md("## Window Schedule Complete")
output.print_md("**Phase:** {}".format(selected_phase.Name))
output.print_md("**Window types scheduled:** {}".format(len(view_pairs)))
output.print_md("**Sheets created:** {}".format(len(real_sheets)))
output.print_md("**Scale:** 1:{}".format(chosen_scale))
output.print_md("**Layout:** {}".format(chosen_layout))

output.print_md("---")
output.print_md("### Generated Sheets Reference")

table_data = []
for sheet in real_sheets:
    sheet_number = sheet.SheetNumber if sheet.SheetNumber else "---"
    clickable_link = output.linkify(sheet.Id, title=sheet_number)
    table_data.append([clickable_link, sheet.Name])

output.print_table(
    table_data=table_data,
    columns=["Sheet Number (Click to Open)", "Sheet Name"]
)

print("\nSuccessfully generated layout tracking grids natively!")