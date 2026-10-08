# -*- coding: utf-8 -*-
__title__ = "23-Magic Dims"
__doc__ = """Version = 1.7
Date    = 28.06.2026
________________________________________________________________
Description:
Auto Dimension Tool for custom selections (Pipes, Cable Trays, Conduits).
Flattens geometry to 2D to support elements at different elevations.

PRO TIP: Holds Shift key down while clicking the tool to change/reset 
the remembered dimension style.
________________________________________________________________
Author: Katarzyna Lipka-Sidor"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
# ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
from Autodesk.Revit.DB import *

# pyRevit
from pyrevit import forms, script

# .NET Imports
import clr

clr.AddReference('System')
clr.AddReference('PresentationCore')  # Required for Keyboard inputs
from System.Collections.Generic import List
from System.Windows.Input import Keyboard, Key

import math
import os
import json

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
# ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
doc = __revit__.ActiveUIDocument.Document  # type:Document
uidoc = __revit__.ActiveUIDocument  # __revit__ is internal variable in pyRevit
app = __revit__.Application
output = script.get_output()  # pyRevit Output Menu

VALID_VIEW_TYPES = [ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.Section, ViewType.Elevation,
                    ViewType.EngineeringPlan]
MODE_AXIS = "Dimensions to the nearest Grids"
MODE_SPACING = "Continuous dimensions between selected elements"

GRID_SEARCH_RADIUS = 30.0  # feet (~9.1m)
MIN_SPACING_DISTANCE = 0.033  # feet (~10mm)
MAX_SPACING_DISTANCE = 50.0  # feet (~15.2m)
DIM_LINE_MARGIN = 1.5  # feet (~450mm) - offset distance for dimension line
ENDPOINT_TOLERANCE = 0.01  # feet
MIN_MEASURABLE_DISTANCE = 0.01  # feet
ANGLE_TOLERANCE_COS = math.cos(math.radians(0.1))

STORAGE_FILE = os.path.join(os.environ["USERPROFILE"], "pyrevit_magic_wand_config.json")

# FUNCTIONS

class DimensionTypeWrapper(object):
    def __init__(self, item, name):
        self.item = item
        self.name = name

def get_curve_midpoint(curve):
    return curve.Evaluate(0.5, True)

def get_element_curve(element):
    loc = element.Location
    if loc and hasattr(loc, "Curve"):
        return loc.Curve
    return None

def flatten_to_2d(point, is_plan_view=True):
    """Flattens a 3D point into a 2D plane based on view type to ignore elevation differences"""
    if is_plan_view:
        return XYZ(point.X, point.Y, 0.0)
    return point

def normalize_2d_vector(vector):
    """Safely normalizes a 2D projected vector"""
    v_2d = XYZ(vector.X, vector.Y, 0.0)
    length = v_2d.GetLength()
    if length > 0.0001:
        return v_2d.Normalize()
    return XYZ.Zero

def are_parallel(c1, c2):
    if not isinstance(c1, Line) or not isinstance(c2, Line):
        return False
    d1 = normalize_2d_vector(c1.Direction)
    d2 = normalize_2d_vector(c2.Direction)
    if d1.IsAlmostEqualTo(XYZ.Zero) or d2.IsAlmostEqualTo(XYZ.Zero):
        return False
    return abs(d1.DotProduct(d2)) > ANGLE_TOLERANCE_COS

def project_point_to_line_2d(point, line):
    """Projects a point onto a line in 2D horizontal space"""
    p2d = flatten_to_2d(point)
    origin_2d = flatten_to_2d(line.GetEndPoint(0))
    dir_2d = normalize_2d_vector(line.Direction)
    unbound = Line.CreateUnbound(origin_2d, dir_2d)
    return unbound.Project(p2d).XYZPoint

def build_dimension_line(p1, p2, target_z):
    """Builds a flat dimension line matching the Viewplane Z level"""
    p1_flat = XYZ(p1.X, p1.Y, target_z)
    p2_flat = XYZ(p2.X, p2.Y, target_z)
    vector = p2_flat - p1_flat
    if vector.GetLength() < MIN_MEASURABLE_DISTANCE:
        return None
    direction = vector.Normalize()
    return Line.CreateBound(p1_flat - direction * DIM_LINE_MARGIN, p2_flat + direction * DIM_LINE_MARGIN)

# SELECTION & USER INTERFACE

def get_selected_mep_elements():
    """Retrieves specific straight MEP elements from active UI selection"""
    selected_ids = uidoc.Selection.GetElementIds()
    if not selected_ids:
        return []
    valid_categories = {int(BuiltInCategory.OST_DuctCurves), int(BuiltInCategory.OST_PipeCurves), int(BuiltInCategory.OST_CableTray), int(BuiltInCategory.OST_Conduit)}
    elements = []
    for e_id in selected_ids:
        el = doc.GetElement(e_id)
        if el and el.Category and int(el.Category.BuiltInCategory) in valid_categories:
            curve = get_element_curve(el)
            if isinstance(curve, Line):
                # STRESS TEST FIX: Odrzucenie elementów pionowych (Z-driven)
                if abs(curve.Direction.Z) > 0.999:
                    continue
                elements.append(el)
    return elements

def get_linear_dimension_types():
    """Gathers all linear DimensionTypes sorted by name"""
    dim_types = FilteredElementCollector(doc).OfClass(DimensionType).ToElements()
    linear_types = [dt for dt in dim_types if dt.StyleType == DimensionStyleType.Linear]
    items = []
    for dt in linear_types:
        p_name = dt.Parameter[BuiltInParameter.SYMBOL_NAME_PARAM]
        name_str = p_name.AsString() if p_name else "Unknown Style"
        items.append(DimensionTypeWrapper(dt, name_str))
    return sorted(items, key=lambda x: x.name)

# CONFIGURATION PERSISTENCE HANDLERS (NATIVE JSON BACKEND)

def save_saved_style_name(name_str):
    try:
        with open(STORAGE_FILE, 'w') as f:
            json.dump({"last_style": name_str}, f)
    except Exception:
        pass

def load_saved_style_name():
    if os.path.exists(STORAGE_FILE):
        try:
            with open(STORAGE_FILE, 'r') as f:
                data = json.load(f)
                return data.get("last_style", "")
        except Exception:
            return ""
    return ""

# GEOMETRY REFERENCE ENGINE (FOR CABLE TRAYS & PIPING)

def get_centerline_reference(element, opt):
    loc_curve = get_element_curve(element)
    if not loc_curve:
        return None

    geo = element.get_Geometry(opt)
    if geo:
        ref = _find_ref(geo, loc_curve)
        if ref: return ref

    opt_model = Options()
    opt_model.ComputeReferences = True
    opt_model.DetailLevel = ViewDetailLevel.Fine

    geo_model = element.get_Geometry(opt_model)
    if geo_model:
        return _find_ref(geo_model, loc_curve)
    return None


def _find_ref(geometry_element, loc_curve):
    p0 = loc_curve.GetEndPoint(0)
    p1 = loc_curve.GetEndPoint(1)

    for g in geometry_element:
        if isinstance(g, GeometryInstance):
            for ig in g.GetInstanceGeometry():
                if isinstance(ig, Line) and ig.Reference:
                    g0 = ig.GetEndPoint(0)
                    g1 = ig.GetEndPoint(1)
                    if (g0.IsAlmostEqualTo(p0, ENDPOINT_TOLERANCE) and g1.IsAlmostEqualTo(p1, ENDPOINT_TOLERANCE)) or \
                            (g0.IsAlmostEqualTo(p1, ENDPOINT_TOLERANCE) and g1.IsAlmostEqualTo(p0, ENDPOINT_TOLERANCE)):
                        return ig.Reference
        elif isinstance(g, Line) and g.Reference:
            g0 = g.GetEndPoint(0)
            g1 = g.GetEndPoint(1)
            if (g0.IsAlmostEqualTo(p0, ENDPOINT_TOLERANCE) and g1.IsAlmostEqualTo(p1, ENDPOINT_TOLERANCE)) or \
                    (g0.IsAlmostEqualTo(p1, ENDPOINT_TOLERANCE) and g1.IsAlmostEqualTo(p0, ENDPOINT_TOLERANCE)):
                return g.Reference
    return None

# DIMENSIONING PROCESSING FUNCTIONS

def run_axis_dimensioning(view, elements, opt, dim_type):
    grids = list(FilteredElementCollector(doc, view.Id).OfClass(Grid))

    if not grids:
        return 0, [("Global", "No Grids visible in the active view to dimension to.")]

    created, skipped = 0, []
    target_z = view.Origin.Z

    for e in elements:
        curve = get_element_curve(e)
        mid_2d = flatten_to_2d(get_curve_midpoint(curve))
        nearest_grid, nearest_dist, nearest_point_2d = None, None, None

        for g in grids:
            gcurve = g.Curve
            if not isinstance(gcurve, Line) or not are_parallel(curve, gcurve):
                continue

            proj_2d = project_point_to_line_2d(mid_2d, gcurve)
            dist = mid_2d.DistanceTo(proj_2d)
            if dist > GRID_SEARCH_RADIUS:
                continue

            if nearest_dist is None or dist < nearest_dist:
                nearest_dist, nearest_grid, nearest_point_2d = dist, g, proj_2d

        if not nearest_grid:
            skipped.append((output.linkify(e.Id), "No parallel grid within range"))
            continue

        eref = get_centerline_reference(e, opt)
        if not eref:
            skipped.append((output.linkify(e.Id), "No centerline reference found"))
            continue

        line = build_dimension_line(mid_2d, nearest_point_2d, target_z)
        if not line:
            continue

        ref_array = ReferenceArray()
        ref_array.Append(eref)
        ref_array.Append(Reference(nearest_grid))

        try:
            new_dim = doc.Create.NewDimension(view, line, ref_array, dim_type)
            created += 1
        except Exception as ex:
            skipped.append((output.linkify(e.Id), str(ex)))

    return created, skipped

def run_spacing_dimensioning(view, elements, opt, dim_type):
    if len(elements) < 2:
        return 0, [("Selection", "At least two elements are required to calculate spacing")]

    base_element = elements[0]
    base_curve = get_element_curve(base_element)

    base_dir_2d = normalize_2d_vector(base_curve.Direction)
    if base_dir_2d.IsAlmostEqualTo(XYZ.Zero):
        return 0, [("Geometry Error", "Invalid element directional vector on 2D view plane")]

    perp_dir_2d = base_dir_2d.CrossProduct(XYZ.BasisZ).Normalize()

    valid_sorted_data = []
    skipped = []
    target_z = view.Origin.Z

    for e in elements:
        curve = get_element_curve(e)
        if not are_parallel(base_curve, curve):
            skipped.append((output.linkify(e.Id), "Element is not parallel to the selection group context"))
            continue

        eref = get_centerline_reference(e, opt)
        if not eref:
            skipped.append((output.linkify(e.Id), "No valid geometric centerline reference found"))
            continue

        midpoint = get_curve_midpoint(curve)
        midpoint_2d = flatten_to_2d(midpoint)

        sort_scalar = midpoint_2d.DotProduct(perp_dir_2d)
        valid_sorted_data.append((sort_scalar, midpoint_2d, eref, e.Id))

    valid_sorted_data.sort(key=lambda x: x[0])

    if len(valid_sorted_data) < 2:
        return 0, skipped

    ref_array = ReferenceArray()
    start_pt_2d = valid_sorted_data[0][1]
    end_pt_2d = valid_sorted_data[-1][1]

    for data in valid_sorted_data:
        ref_array.Append(data[2])

    line = build_dimension_line(start_pt_2d, end_pt_2d, target_z)
    if not line:
        return 0, [("Geometry Block", "Failed to build valid flat 2D dimension track lines")]

    try:
        new_dim = doc.Create.NewDimension(view, line, ref_array, dim_type)
        return 1, skipped
    except Exception as ex:
        return 0, [("Revit API Creation Engine Fault", str(ex))]

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░

def run_script():
    view = doc.ActiveView
    if not view or view.ViewType not in VALID_VIEW_TYPES:
        forms.alert("Please switch to a valid 2D Plan or Section View.", title="Invalid View")
        return

    elements = get_selected_mep_elements()
    if not elements:
        forms.alert("No valid elements selected! Please select Pipes, Conduits, or Cable Trays first.",
                    title="Selection Missing")
        return

    mode_options = [MODE_AXIS, MODE_SPACING]
    selected_mode = forms.SelectFromList.show(mode_options, multiselect=False, title="Choose Dimensioning Target")
    if not selected_mode:
        return

    all_styles = get_linear_dimension_types()
    if not all_styles:
        forms.alert("No Linear Dimension Styles found in this project.", title="Styles Missing")
        return

    saved_style_name = load_saved_style_name()
    selected_dim_type = None
    chosen_style_name = ""
    force_reset = Keyboard.IsKeyDown(Key.LeftShift) or Keyboard.IsKeyDown(Key.RightShift)

    if saved_style_name and not force_reset:
        for s in all_styles:
            if s.name == saved_style_name:
                selected_dim_type = s.item
                chosen_style_name = s.name
                break

    if not selected_dim_type:
        selected_style_wrapper = forms.SelectFromList.show(all_styles, multiselect=False,
                                                           title="Choose Dimension Style")
        if not selected_style_wrapper:
            return
        selected_dim_type = selected_style_wrapper.item
        chosen_style_name = selected_style_wrapper.name
        save_saved_style_name(chosen_style_name)

    can_change_detail = True
    template_id = view.ViewTemplateId
    if template_id != ElementId.InvalidElementId:
        template = doc.GetElement(template_id)
        p = template.Parameter[BuiltInParameter.VIEW_DETAIL_LEVEL]
        if p and p.IsReadOnly:
            can_change_detail = False

# Context Performance Optimization: Single TransactionGroup wrapper
    tg = TransactionGroup(doc, "Magic Wand Group")
    tg.Start()

    original_detail = view.DetailLevel
    if original_detail != ViewDetailLevel.Fine and can_change_detail:
        with Transaction(doc, "Prepare View for Dimensions") as t_detail:
            t_detail.Start()
            view.DetailLevel = ViewDetailLevel.Fine
            doc.Regenerate()
            t_detail.Commit()
    elif original_detail != ViewDetailLevel.Fine and not can_change_detail:
        # Informacja dla użytkownika (non-blocking alert)
        forms.toast("View Template controls Detail Level. Geometric links might be less accurate.",
                    title="Template Lock")

    opt = Options()
    opt.View = view
    opt.ComputeReferences = True
    created, skipped = 0, []

    with Transaction(doc, "Magic Wand - Custom Selection") as t:
        t.Start()
        try:
            if selected_mode == MODE_AXIS:
                created, skipped = run_axis_dimensioning(view, elements, opt, selected_dim_type)
            else:
                created, skipped = run_spacing_dimensioning(view, elements, opt, selected_dim_type)

            if created:
                t.Commit()
            else:
                t.RollBack()
        except Exception as ex:
            t.RollBack()
            forms.alert(str(ex), title="Execution Error")

    if original_detail != ViewDetailLevel.Fine:
        with Transaction(doc, "Restore view settings") as t_restore:
            t_restore.Start()
            view.DetailLevel = original_detail
            t_restore.Commit()

    tg.Assimilate()

# Output report
    output.print_md(
        "## MAGIC WAND SUMMARY\n**Mode:** {}\n**Style Applied:** {}\n**Created Chains:** {}\n**Skipped:** {}".format(
            selected_mode, chosen_style_name, created, len(skipped)))

    if skipped:
        output.print_md("### Details on skipped items:")
        for ident, reason in skipped:
            output.print_md("* {} : {}".format(ident, reason))

if __name__ == "__main__":
    run_script()