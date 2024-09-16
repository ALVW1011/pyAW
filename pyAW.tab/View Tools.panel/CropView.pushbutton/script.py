# -*- coding: utf-8 -*-
import sys
import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.Exceptions import *
from pyrevit import forms

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
view = doc.ActiveView

# List of view types that can be cropped
croppable_view_types = [
    ViewType.FloorPlan,
    ViewType.CeilingPlan,
    ViewType.EngineeringPlan,
    ViewType.AreaPlan,
    ViewType.Section,
    ViewType.Elevation,
    ViewType.Detail,
    ViewType.DraftingView
]

# Verify that the view supports cropping
if view.ViewType not in croppable_view_types:
    forms.alert("This view cannot be cropped.", exitscript=True)

# Explicitly enable cropping for the view if it supports it
if not view.CropBoxActive:
    with Transaction(doc, "Enable Cropping") as t:
        t.Start()
        view.CropBoxActive = True  # Ensure cropping is enabled
        view.CropBoxVisible = True  # Make the crop region visible
        t.Commit()

# Progress tracking using pyRevit forms
with forms.ProgressBar(title='Processing Selection...') as pb:
    pb.update_progress(0, 20)

    # Prompt the user to select a room or an element
    try:
        sel = uidoc.Selection
        ref = sel.PickObject(ObjectType.Element, "Select a room or an element within the room.")
        selected_element = doc.GetElement(ref.ElementId)
        pb.update_progress(20, 40)

    except OperationCanceledException:
        forms.alert("Selection Cancelled", exitscript=True)
    except Exception as e:
        forms.alert("Error during selection: {0}".format(str(e)), exitscript=True)

    # Initialize variables for the summary
    family_name = ""
    type_name = ""
    room_name = ""
    room_number = ""
    crop_successful = False

    # Get family and type names of the selected element
    if isinstance(selected_element, FamilyInstance):
        family_name = selected_element.Symbol.Family.Name
        type_name = selected_element.Name
    elif isinstance(selected_element, Element):
        family_name = selected_element.Category.Name
        type_name = selected_element.Name

    # Prompt the user for an offset (in mm)
    crop_offset_input = forms.ask_for_string(
        title="Crop Boundary Offset",
        prompt="Enter an offset for the crop boundary (in mm):",
        default="100"
)
    if crop_offset_input is None:
        forms.alert("Offset input cancelled.", exitscript=True)

    # Convert input to float
    try:
        crop_offset_mm = float(crop_offset_input)
    except ValueError:
        forms.alert("Invalid number entered for offset.", exitscript=True)

    # Convert mm to feet (since Revit API uses feet)
    crop_offset_feet = crop_offset_mm / 304.8


    pb.update_progress(40, 60)

    # Determine if the selected element is a room
    room = None
    if isinstance(selected_element, SpatialElement) and selected_element.Category.Id.IntegerValue == int(BuiltInCategory.OST_Rooms):
        room = selected_element
    else:
        # Find the room containing the element
        location = selected_element.Location
        if location and hasattr(location, 'Point'):
            point = location.Point
            room = doc.GetRoomAtPoint(point)

    # Get room name and number if a room is found
    if room:
        room_name = room.LookupParameter("Name").AsString()
        room_number = room.LookupParameter("Number").AsString()

    # Function to get room boundary lines
    def get_room_boundary_lines(room):
        options = SpatialElementBoundaryOptions()
        boundary_segments = room.GetBoundarySegments(options)
        boundary_curves = []

        if boundary_segments:
            for boundary_list in boundary_segments:
                for segment in boundary_list:
                    curve = segment.GetCurve()
                    boundary_curves.append(curve)
        return boundary_curves

    # Get room boundary lines
    boundary_lines = get_room_boundary_lines(room)

    if not boundary_lines:
        forms.alert("No boundary lines found for the room.", exitscript=True)

    pb.update_progress(60, 80)

    # Calculate combined bounding box from room boundary lines and apply the offset
    def get_combined_bounding_box_from_curves(curves, offset):
        min_x = float('inf')
        min_y = float('inf')
        max_x = float('-inf')
        max_y = float('-inf')

        for curve in curves:
            start = curve.GetEndPoint(0)
            end = curve.GetEndPoint(1)
            min_x = min(min_x, start.X, end.X)
            min_y = min(min_y, start.Y, end.Y)
            max_x = max(max_x, start.X, end.X)
            max_y = max(max_y, start.Y, end.Y)

        # Apply the offset to the bounding box
        bbox = BoundingBoxXYZ()
        bbox.Min = XYZ(min_x - offset, min_y - offset, view.GenLevel.Elevation)  # Use the elevation of the view level
        bbox.Max = XYZ(max_x + offset, max_y + offset, view.GenLevel.Elevation)

        return bbox

    # Create crop region from room boundary lines
    combined_bbox = get_combined_bounding_box_from_curves(boundary_lines, crop_offset_feet)

    if combined_bbox:
        with Transaction(doc, "Adjust Crop Region") as t:
            try:
                t.Start()
                view.CropBoxActive = True
                view.CropBoxVisible = True
                view.CropBox = combined_bbox
                t.Commit()
                crop_successful = True
                pb.update_progress(80, 100)
            except Exception as e:
                t.RollBack()
                forms.alert("Error committing crop region transaction: {0}".format(str(e)), exitscript=True)
    else:
        forms.alert("Could not determine the crop region based on the room's boundaries.", exitscript=True)

    # Display final summary using a pyRevit form
    summary_message = (
        "Selected Element:\n"
        "Family: {0}\n"
        "Type: {1}\n\n"
        "Room Information:\n"
        "Room Name: {2}\n"
        "Room Number: {3}\n\n"
        "Crop Successful: {4}\n"
        "Crop Offset: {5} mm"
    ).format(
        family_name,
        type_name,
        room_name,
        room_number,
        'Yes' if crop_successful else 'No',
        crop_offset_mm
    )
    
    forms.alert(summary_message)
