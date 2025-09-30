# -*- coding: utf-8 -*-
"""
Places a text note at the intersection of the leftmost and bottom-most
grids on the selected sheet(s).

This script is designed to be used with pyRevit in Autodesk Revit.

"""
# Import necessary Revit API classes and pyRevit modules
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ViewSheet,
    Grid,
    TextNote,
    TextNoteType,
    Line,
    XYZ,
    SetComparisonResult,
    IntersectionResultArray,
    View,
    ViewportRotation
)
from pyrevit import revit, forms
import clr # Common Language Runtime for .NET interop

# Get the current Revit document and selection
doc = revit.doc
selection = revit.get_selection()

# --- Pre-checks ---

# 1. Filter the selection to get only ViewSheet elements
selected_sheets = [s for s in selection if isinstance(s, ViewSheet)]

if not selected_sheets:
    forms.alert("No sheets selected. Please select one or more sheets in a project view or from the Project Browser.", exitscript=True)

# 2. Find a default TextNoteType to use for creating the text
# We need a TextNoteType Id to create a TextNote element.
default_text_type = FilteredElementCollector(doc)\
                    .OfClass(TextNoteType)\
                    .FirstElement()

if not default_text_type:
    forms.alert("No Text Note Types found in the project. Cannot create text.", exitscript=True)


# --- Main Logic ---

# Group all model changes into a single transaction for better performance and undo capability
with revit.Transaction('Place Text at Grid Intersections'):
    print("Processing {} sheet(s)...".format(len(selected_sheets)))
    
    # Iterate over each selected sheet
    for sheet in selected_sheets:
        sheet_name = revit.query.get_name(sheet)
        print("\nProcessing Sheet: {}".format(sheet_name))

        # --- Find Grids on the Sheet ---
        
        # Collect all unique grids from all views placed on the current sheet
        grids_on_sheet = {} # Use a dictionary to easily store unique grids by their Id
        placed_view_ids = sheet.GetAllPlacedViews()
        
        for view_id in placed_view_ids:
            # Get the view element from its Id
            view = doc.GetElement(view_id)
            # Ensure the view can actually display grids (e.g., plan views, detail views)
            if view and isinstance(view, View) and view.CanBePrinted:
                view_grids = FilteredElementCollector(doc, view_id)\
                             .OfClass(Grid)\
                             .ToElements()
                
                # Add unique grids to our dictionary
                for grid in view_grids:
                    if grid.Id not in grids_on_sheet:
                        grids_on_sheet[grid.Id] = grid
        
        if not grids_on_sheet:
            print(" -> No grids found on this sheet.")
            continue # Skip to the next sheet

        # --- Identify Leftmost and Bottom-most Grids ---

        leftmost_grid = None
        bottommost_grid = None
        # Use extreme values to ensure the first grid checked becomes the initial reference
        min_x_coord = float('inf') 
        min_y_coord = float('inf')

        for grid in grids_on_sheet.values():
            curve = grid.Curve
            
            # We are only considering grids that are straight lines
            if isinstance(curve, Line):
                direction = curve.Direction.Normalize()
                origin = curve.GetEndPoint(0)

                # Check if grid is vertical (its direction is parallel to the Y-axis)
                if direction.IsAlmostEqualTo(XYZ.BasisY) or direction.IsAlmostEqualTo(-XYZ.BasisY):
                    if origin.X < min_x_coord:
                        min_x_coord = origin.X
                        leftmost_grid = grid
                
                # Check if grid is horizontal (its direction is parallel to the X-axis)
                elif direction.IsAlmostEqualTo(XYZ.BasisX) or direction.IsAlmostEqualTo(-XYZ.BasisX):
                    if origin.Y < min_y_coord:
                        min_y_coord = origin.Y
                        bottommost_grid = grid

        # --- Calculate Intersection and Place Text ---

        if leftmost_grid and bottommost_grid:
            curve_left = leftmost_grid.Curve
            curve_bottom = bottommost_grid.Curve
            
            # The Intersect method requires an 'out' parameter in C#, which is handled
            # using a clr.Reference in IronPython.
            result_array_ref = clr.Reference[IntersectionResultArray]()
            result = curve_left.Intersect(curve_bottom, result_array_ref)
            
            # Check if the curves overlap/intersect
            if result == SetComparisonResult.Overlap:
                intersection_point = result_array_ref.Value[0].XYZPoint
                
                # --- Find a suitable view to place the text in ---
                # A suitable view must have the intersection point inside its crop region.
                suitable_view = None
                
                # Re-iterate through the placed views to find one that can host the text note
                for view_id in placed_view_ids:
                    view = doc.GetElement(view_id)
                    if not (view and isinstance(view, View) and view.CanBePrinted):
                        continue

                    # Check if the view is cropped and if the point is inside the crop box
                    if view.CropBoxActive:
                        crop_box = view.CropBox
                        # Check if the point is within the X and Y bounds of the crop box
                        # This works for plan views. We ignore Z for this check.
                        if (crop_box.Min.X <= intersection_point.X <= crop_box.Max.X and
                            crop_box.Min.Y <= intersection_point.Y <= crop_box.Max.Y):
                            suitable_view = view
                            break # Found a good view, no need to check others
                    else:
                        # If the view is not cropped, it's a suitable candidate
                        suitable_view = view
                        break
                
                if suitable_view:
                    # --- Transform model point to sheet point ---
                    target_viewport = None
                    # Find the viewport associated with the suitable view on this sheet
                    for vp_id in sheet.GetAllViewports():
                        vp = doc.GetElement(vp_id)
                        if vp.ViewId == suitable_view.Id:
                            target_viewport = vp
                            break
                    
                    if target_viewport:
                        # Get the center of the viewport on the sheet
                        viewport_center_sheet = target_viewport.GetBoxCenter()
                        
                        # Get the center of the view's crop box in model coordinates
                        crop_box_model = suitable_view.CropBox
                        crop_box_center_model = (crop_box_model.Min + crop_box_model.Max) / 2.0
                        
                        # Calculate the offset vector from the view center to the intersection point
                        offset_vector_model = intersection_point - crop_box_center_model

                        # Get the view scale
                        view_scale = suitable_view.Scale

                        # Scale the offset vector by the view scale
                        scaled_offset_vector = offset_vector_model / view_scale
                        
                        # Handle the viewport rotation on the sheet
                        # FIXED: Use the .Rotation property instead of GetRotationOnSheet() method
                        rotation = target_viewport.Rotation
                        
                        offset_vector_sheet = None
                        if rotation == ViewportRotation.None:
                            offset_vector_sheet = XYZ(scaled_offset_vector.X, scaled_offset_vector.Y, 0)
                        elif rotation == ViewportRotation.NinetyDegreesCounterclockwise:
                            offset_vector_sheet = XYZ(-scaled_offset_vector.Y, scaled_offset_vector.X, 0)
                        elif rotation == ViewportRotation.Clockwise:
                            offset_vector_sheet = XYZ(scaled_offset_vector.Y, -scaled_offset_vector.X, 0)
                        # In Revit 2022+, Halfway was renamed to OneHundredEightyDegrees
                        elif rotation == ViewportRotation.Halfway or str(rotation) == "OneHundredEightyDegrees":
                             offset_vector_sheet = XYZ(-scaled_offset_vector.X, -scaled_offset_vector.Y, 0)

                        if offset_vector_sheet:
                            # Calculate the final position on the sheet
                            text_location_sheet = viewport_center_sheet + offset_vector_sheet
                        
                            # Create the text note directly on the sheet
                            TextNote.Create(doc, sheet.Id, text_location_sheet, "X", default_text_type.Id)
                            
                            print(" -> Success: Placed text on sheet at intersection of grids '{}' and '{}'.".format(
                                revit.query.get_name(leftmost_grid), 
                                revit.query.get_name(bottommost_grid)
                            ))
                        else:
                            print(" -> Warning: Unsupported viewport rotation found. Could not place text on sheet.")
                    else:
                        print(" -> Warning: Could not find the viewport for the suitable view. Text not placed on sheet.")

                else:
                    # If no suitable view was found
                    print(" -> Warning: Grid intersection is outside the crop region of all views on this sheet. Text not placed.")

            else:
                print(" -> Grids '{}' and '{}' do not intersect.".format(
                    revit.query.get_name(leftmost_grid), 
                    revit.query.get_name(bottommost_grid)
                ))
        else:
            # Provide feedback if one or both grids could not be found
            if not leftmost_grid:
                print(" -> Warning: Could not determine the leftmost vertical grid.")
            if not bottommost_grid:
                print(" -> Warning: Could not determine the bottom-most horizontal grid.")

print("\nScript finished.")