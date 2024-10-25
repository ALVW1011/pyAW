from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    FamilyInstance,
    Transaction,
    Level,
    BuiltInParameter,
    XYZ,
    ViewType,
    SketchPlane,
    Plane
)
import clr

# Initialize an empty list to collect errors
error_messages = []

# Helper Functions

def feet_to_meters(feet):
    """Convert feet to meters."""
    return feet * 0.3048

def get_categories(doc):
    """Retrieve all unique category names in the document."""
    categories = set()
    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
    for element in collector:
        if element.Category and element.Category.Name:
            categories.add(element.Category.Name)
    return sorted(list(categories))

def get_families(doc, category_name):
    """Retrieve all unique family names within a specific category."""
    families = set()
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance)
    for fam_instance in collector:
        if fam_instance.Category and fam_instance.Category.Name == category_name:
            families.add(fam_instance.Symbol.Family.Name)
    return sorted(list(families))

def get_family_types(doc, category_name, family_names):
    """Retrieve all unique family types within selected families."""
    family_types = set()
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance)
    for fam_instance in collector:
        if (fam_instance.Category and fam_instance.Category.Name == category_name and
            fam_instance.Symbol.Family.Name in family_names):
            family_types.add(fam_instance.Name)
    return sorted(list(family_types))

def get_family_instances(doc, category_name, family_names, type_names):
    """Retrieve all family instances based on selected category, families, and types."""
    instances = []
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance)
    for fam_instance in collector:
        if (fam_instance.Category and fam_instance.Category.Name == category_name and
            fam_instance.Symbol.Family.Name in family_names and
            fam_instance.Name in type_names):
            instances.append(fam_instance)
    return instances

def get_current_level(doc, instance):
    """
    Retrieve the current level name of a family instance.
    
    Args:
        doc: The active Revit document.
        instance: The FamilyInstance element.
    
    Returns:
        The name of the current level or "None assigned" if not found.
    """
    # Attempt to retrieve using INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM
    schedule_level_param = instance.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM)
    if schedule_level_param:
        level_id = schedule_level_param.AsElementId()
        if level_id and level_id != DB.ElementId.InvalidElementId:
            level = doc.GetElement(level_id)
            if level:
                return level.Name

    # Attempt to retrieve using LevelId
    level_id = instance.LevelId
    if level_id and level_id != DB.ElementId.InvalidElementId:
        level = doc.GetElement(level_id)
        if level:
            return level.Name

    # Attempt to retrieve using 'Reference Level' or 'Work Plane' parameters
    param_names = ['Reference Level', 'Work Plane']
    for param_name in param_names:
        param = instance.LookupParameter(param_name)
        if param:
            if param.StorageType == DB.StorageType.ElementId:
                level_id = param.AsElementId()
                if level_id and level_id != DB.ElementId.InvalidElementId:
                    level = doc.GetElement(level_id)
                    if level:
                        return level.Name
            elif param.StorageType == DB.StorageType.String:
                level_name = param.AsString()
                if level_name:
                    return level_name
    return "None assigned"

def get_nearest_level(doc, instance, direction, selected_levels):
    """
    Determine the nearest level based on direction and selected levels.
    
    Args:
        doc: The active Revit document.
        instance: The FamilyInstance element.
        direction: "Below", "Above", or "Either".
        selected_levels: List of Level elements to consider.
    
    Returns:
        The nearest Level element or None if not found.
    """
    if hasattr(instance.Location, 'Point'):
        instance_location = instance.Location.Point
        nearest_level = None
        nearest_distance = float('inf')
        for level in selected_levels:
            level_elevation = level.Elevation
            if direction == "Below":
                distance = instance_location.Z - level_elevation
                if 0 <= distance < nearest_distance:
                    nearest_distance = distance
                    nearest_level = level
            elif direction == "Above":
                distance = level_elevation - instance_location.Z
                if 0 <= distance < nearest_distance:
                    nearest_distance = distance
                    nearest_level = level
            elif direction == "Either":
                distance = abs(instance_location.Z - level_elevation)
                if distance < nearest_distance:
                    nearest_distance = distance
                    nearest_level = level
        return nearest_level
    else:
        return None

def get_element_z_position(element):
    """
    Retrieve the Z-axis position of the element.
    
    Args:
        element: The FamilyInstance element.
    
    Returns:
        The Z-axis coordinate as a float.
    
    Raises:
        Exception: If the element does not have a Location Point.
    """
    location = element.Location
    if not location or not hasattr(location, 'Point'):
        raise Exception("Element '{0}' does not have a movable location point.".format(element.Name))
    return location.Point.Z

def get_offset_parameter(element):
    """
    Retrieve the appropriate offset parameter of the element.
    
    Args:
        element: The FamilyInstance element.
    
    Returns:
        The Parameter object if found, else None.
    """
    # List of potential offset parameter names
    offset_param_names = [
        "Offset from Host",
        "Elevation from Level",
        "Base Offset",
        "Offset",
        "Work Plane Offset",
        "Offset Elevation",
        "Height Offset from Level",
        "Shaft Offset",
        "Offset from Base Level"
        # Add more names as needed
    ]
    
    for param_name in offset_param_names:
        param = element.LookupParameter(param_name)
        if param:
            return param
    return None

def calculate_current_offset(doc, element):
    """
    Calculate the current offset of the element relative to its level.
    
    Args:
        doc: The active Revit document.
        element: The FamilyInstance element.
    
    Returns:
        The offset as a float or an error message.
    """
    try:
        current_z = get_element_z_position(element)
        current_level_name = get_current_level(doc, element)
        current_level = get_level_by_name(doc, current_level_name)
        
        if not current_level:
            return 0.0  # Default offset if level not found
        
        current_elevation = current_level.Elevation
        offset = current_z - current_elevation
        return offset
    except Exception as e:
        error_messages.append("Error calculating current offset for '{0}': {1}".format(element.Name, e))
        return "Error"

def calculate_new_offset(target_level_elevation, current_z):
    """
    Calculate the new offset to maintain the element's Z-axis position.
    
    Args:
        target_level_elevation: The elevation of the target level.
        current_z: The current Z-axis position of the element.
    
    Returns:
        The new offset value as a float.
    """
    new_offset = current_z - target_level_elevation
    return new_offset

def get_level_by_name(doc, level_name):
    """
    Retrieve a Level element by its name.
    
    Args:
        doc: The active Revit document.
        level_name: The name of the level.
    
    Returns:
        The Level element if found, else None.
    """
    if level_name == "None assigned":
        return None
    collector = FilteredElementCollector(doc).OfClass(Level)
    for level in collector:
        if level.Name == level_name:
            return level
    return None

def set_level_parameters(element, level):
    """
    Set all level-related parameters of an element to the specified level.
    
    Args:
        element: The FamilyInstance element.
        level: The Level element to set.
    """
    try:
        # Attempt to set INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM
        schedule_level_param = element.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM)
        if schedule_level_param and not schedule_level_param.IsReadOnly:
            schedule_level_param.Set(level.Id)
        
        # Attempt to set FAMILY_LEVEL_PARAM
        family_level_param = element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
        if family_level_param and not family_level_param.IsReadOnly:
            family_level_param.Set(level.Id)
        
        # Set all other Level-related parameters dynamically
        for param in element.Parameters:
            if param.StorageType == DB.StorageType.ElementId and "level" in param.Definition.Name.lower():
                # Avoid setting parameters already set above
                if param.Definition.Name not in [
                    "INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM",
                    "FAMILY_LEVEL_PARAM"
                ]:
                    if not param.IsReadOnly:
                        try:
                            param.Set(level.Id)
                        except Exception as e:
                            error_messages.append("Failed to set parameter '{0}' for '{1}': {2}".format(param.Definition.Name, element.Name, e))
    except Exception as e:
        error_messages.append("Error setting level parameters for '{0}': {1}".format(element.Name, e))

def edit_work_plane(element, target_level, doc):
    """
    Edit the Work Plane of an element to align with the target level.
    
    Args:
        element: The FamilyInstance element.
        target_level: The Level element to align the Work Plane with.
        doc: The active Revit document.
    """
    try:
        # Check if the element has a Work Plane property
        if hasattr(element, "GetWorkPlane") and hasattr(element, "SetWorkPlane"):
            current_wp = element.GetWorkPlane()
            if current_wp:
                # Define a horizontal plane at the target level's elevation aligned with the element's location
                location_point = element.Location.Point
                plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(location_point.X, location_point.Y, target_level.Elevation))
                sketch_plane = SketchPlane.Create(doc, plane)
                
                # Set the new Work Plane
                element.SetWorkPlane(sketch_plane)
    except Exception as e:
        error_messages.append("Failed to edit Work Plane for '{0}': {1}".format(element.Name, e))

def verify_level_assignment(doc, instance, desired_level_name):
    """
    Verify if the instance's level has been successfully updated.
    
    Args:
        doc: The active Revit document.
        instance: The FamilyInstance element.
        desired_level_name: The name of the desired level.
    
    Returns:
        True if the level matches the desired level, False otherwise.
    """
    current_level_name = get_current_level(doc, instance)
    return current_level_name == desired_level_name

def assign_levels(doc, instances_to_update, direction):
    """
    Assign the nearest levels to the selected family instances by modifying them in place.
    
    Args:
        doc: The active Revit document.
        instances_to_update: List of dictionaries with instance data.
        direction: Direction for level assignment ("Below", "Above", "Either").
    
    Returns:
        List of tuples with updated instance information.
    """
    updated_instances = []
    with Transaction(doc, "Assign Nearest Level") as t:
        try:
            t.Start()
            for data in instances_to_update:
                instance = data['instance']
                nearest_level = data['desired_level']
                current_level = data['current_level']
                current_offset = data['current_offset']
                
                if not nearest_level:
                    error_messages.append("No nearest level found for '{0}'.".format(instance.Name))
                    continue  # Skip to next instance
                
                try:
                    # Set all level-related parameters to the nearest level
                    set_level_parameters(instance, nearest_level)
                    
                    # Edit Work Plane if applicable
                    edit_work_plane(instance, nearest_level, doc)
                    
                    # Calculate and set the new offset
                    new_offset = calculate_new_offset(nearest_level.Elevation, get_element_z_position(instance))
                    offset_param = get_offset_parameter(instance)
                    if offset_param and not offset_param.IsReadOnly:
                        offset_param.Set(new_offset)
                        updated_instances.append((instance, current_level, nearest_level.Name, current_offset, new_offset))
                    else:
                        # If no offset parameter, append with "N/A"
                        updated_instances.append((instance, current_level, nearest_level.Name, current_offset, "N/A"))
                        error_messages.append("No editable Offset parameter for '{0}'. Z-position may have changed.".format(instance.Name))
                    
                    # Verification Step
                    if verify_level_assignment(doc, instance, nearest_level.Name):
                        # Level assignment successful
                        pass
                    else:
                        error_messages.append("Verification failed: Level for '{0}' was not updated correctly.".format(instance.Name))
                    
                except Exception as e:
                    updated_instances.append((instance, current_level, nearest_level.Name, current_offset, "N/A"))
                    error_messages.append("Failed to adjust parameters for '{0}': {1}".format(instance.Name, e))
            t.Commit()
        except Exception as e:
            t.RollBack()
            error_messages.append("Transaction failed: {0}".format(e))
    return updated_instances

def select_direction():
    """Prompt user to select level assignment direction."""
    directions = ["Below", "Above", "Either"]
    selected_direction = forms.SelectFromList.show(
        directions,
        title="Select Level Assignment Direction",
        multiselect=False
    )
    if not selected_direction:
        forms.alert("No direction selected.", exitscript=True)
    return selected_direction

def select_levels(doc):
    """Prompt user to select levels to use for assignment."""
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    level_names = [level.Name for level in levels]
    selected_level_names = forms.SelectFromList.show(
        level_names,
        title="Select Levels to Use",
        multiselect=True
    )
    if not selected_level_names:
        forms.alert("No levels selected.", exitscript=True)
    selected_levels = [level for level in levels if level.Name in selected_level_names]
    return selected_levels

def get_current_selection(doc):
    """
    Retrieve currently selected FamilyInstance elements.
    
    Args:
        doc: The active Revit document.
    
    Returns:
        List of selected FamilyInstance elements.
    """
    selection = revit.get_selection()
    existing_instances = [el for el in selection if isinstance(el, FamilyInstance)]
    return existing_instances

def select_elements_directly(doc):
    """
    Allow users to directly select elements from a graphical view.
    
    Args:
        doc: The active Revit document.
    
    Returns:
        Combined list of pre-selected and newly selected FamilyInstance elements.
    """
    active_view = doc.ActiveView
    graphical_view_types = [
        ViewType.FloorPlan,
        ViewType.CeilingPlan,
        ViewType.ThreeD,
        ViewType.EngineeringPlan,
        ViewType.Section,
        ViewType.Elevation,
        ViewType.Detail,
        ViewType.DrawingSheet,
        ViewType.Internal
    ]

    if active_view.ViewType not in graphical_view_types:
        forms.alert(
            "Direct Selection requires an active graphical view.\n"
            "Please switch to a Floor Plan, Ceiling Plan, Section, Elevation, or 3D view and try again.",
            title="Non-Graphical View Active",
            exitscript=True
        )

    # Retrieve existing selection
    existing_instances = get_current_selection(doc)
    existing_count = len(existing_instances)
    
    if existing_count > 0:
        forms.alert(
            "Found {0} pre-selected family instance(s). You can add more by selecting additional elements.".format(existing_count),
            title="Existing Selection",
            warn_icon=True
        )

    # Prompt user to select additional elements
    try:
        new_selected_elements = revit.pick_elements(
            message="Select additional elements from the model:"
        )
    except Exception as e:
        forms.alert(
            "An error occurred during element selection:\n{0}".format(e),
            title="Selection Error",
            exitscript=True
        )

    if not new_selected_elements:
        if existing_count == 0:
            forms.alert("No elements selected.", exitscript=True)
        else:
            # Only existing elements are considered
            return existing_instances

    # Filter only FamilyInstance elements
    new_instances = [el for el in new_selected_elements if isinstance(el, FamilyInstance)]
    if not new_instances:
        if existing_count == 0:
            forms.alert("No family instances selected in the new selection.", exitscript=True)
        else:
            # Only existing elements are considered
            return existing_instances

    # Combine existing and new instances, avoiding duplicates based on ElementId
    existing_ids = set(el.Id for el in existing_instances)
    combined_instances = existing_instances + [el for el in new_instances if el.Id not in existing_ids]

    total_combined = len(combined_instances)
    forms.alert(
        "Total of {0} family instance(s) selected for processing.".format(total_combined),
        title="Combined Selection",
        warn_icon=True
    )

    if not combined_instances:
        forms.alert("No family instances available for processing.", exitscript=True)

    return combined_instances

def main():
    doc = revit.doc

    # Step 0: Choose selection mode
    selection_modes = ["Filter Selection", "Direct Selection"]
    selected_mode = forms.SelectFromList.show(
        selection_modes,
        title="Choose Element Selection Method",
        multiselect=False
    )
    if not selected_mode:
        forms.alert("No selection mode chosen.", exitscript=True)

    if selected_mode == "Filter Selection":
        # Filter Selection Process
        # Step 1: Select category
        categories = get_categories(doc)
        selected_category = forms.SelectFromList.show(categories, title="Select Category")
        if not selected_category:
            forms.alert("No category selected.", exitscript=True)

        # Step 2: Select families within the selected category
        families = get_families(doc, selected_category)
        selected_families = forms.SelectFromList.show(
            families,
            title="Select Families in {0}".format(selected_category),
            multiselect=True
        )
        if not selected_families:
            forms.alert("No families selected.", exitscript=True)

        # Step 3: Select family types within the selected families
        family_types = get_family_types(doc, selected_category, selected_families)
        selected_family_types = forms.SelectFromList.show(
            family_types,
            title="Select Family Types in Selected Families",
            multiselect=True
        )
        if not selected_family_types:
            forms.alert("No family types selected.", exitscript=True)

        # Step 4: Retrieve family instances based on selections
        instances = get_family_instances(doc, selected_category, selected_families, selected_family_types)
        if not instances:
            forms.alert("No instances found for the selected family types.", exitscript=True)
    else:
        # Direct Selection Process
        instances = select_elements_directly(doc)

    if not instances:
        forms.alert("No family instances selected for processing.", exitscript=True)

    # Step 2: Select levels to use
    selected_levels = select_levels(doc)

    # Step 3: Select level assignment direction
    direction = select_direction()

    # Step 4: Prepare data for each instance
    instances_data = []
    for instance in instances:
        current_level = get_current_level(doc, instance)
        nearest_level = get_nearest_level(doc, instance, direction, selected_levels)
        desired_level_name = nearest_level.Name if nearest_level else "None found"
        current_offset = calculate_current_offset(doc, instance)
        instances_data.append({
            'instance': instance,
            'current_level': current_level,
            'desired_level': nearest_level,
            'desired_level_name': desired_level_name,
            'current_offset': current_offset
        })

    # Step 5: Display the list and confirm
    results = ["Instance Name | Current Level | Current Offset (m) | Desired Level | Desired Offset (m)"]
    for data in instances_data:
        current_offset = data['current_offset']
        if isinstance(current_offset, float):
            offset_display = "{0:.3f}".format(feet_to_meters(current_offset))
        else:
            offset_display = current_offset  # e.g., Error message

        # Calculate desired offset if possible
        if data['desired_level'] and isinstance(current_offset, float):
            try:
                desired_offset = calculate_new_offset(data['desired_level'].Elevation, get_element_z_position(data['instance']))
                desired_offset_m = feet_to_meters(desired_offset)
                desired_offset_display = "{0:.3f}".format(desired_offset_m)
            except Exception as e:
                desired_offset_display = "Error"
                error_messages.append("Error calculating desired offset for '{0}': {1}".format(data['instance'].Name, e))
        else:
            desired_offset_display = "N/A"

        results.append("{0} | {1} | {2} | {3} | {4}".format(
            data['instance'].Name,
            data['current_level'],
            offset_display,
            data['desired_level_name'],
            desired_offset_display
        ))

    message = "\n".join(results)
    proceed = forms.alert(
        message,
        title="Instances to Update",
        yes=True,
        no=True,
        exitscript=True
    )
    if not proceed:
        forms.alert("Operation cancelled by the user.", exitscript=True)

    # Step 6: Assign nearest levels by modifying elements in place
    updated_instances = assign_levels(doc, instances_data, direction)

    # Step 7: Display the results
    if updated_instances:
        results = ["Instance Name | Previous Level | Previous Offset (m) | New Level | New Offset (m)"]
        for instance, prev_level, new_level, prev_offset, new_offset in updated_instances:
            if isinstance(prev_offset, float):
                prev_offset_m = feet_to_meters(prev_offset)
                prev_offset_display = "{0:.3f}".format(prev_offset_m)
            else:
                prev_offset_display = prev_offset  # e.g., Error message

            if isinstance(new_offset, float):
                new_offset_m = feet_to_meters(new_offset)
                new_offset_display = "{0:.3f}".format(new_offset_m)
            else:
                new_offset_display = new_offset  # e.g., "N/A" or Error message

            results.append("{0} | {1} | {2} | {3} | {4}".format(
                instance.Name,
                prev_level,
                prev_offset_display,
                new_level,
                new_offset_display
            ))

        # Append error messages if any
        if error_messages:
            results.append("\nErrors Encountered:")
            for error in error_messages:
                results.append("- {0}".format(error))
        else:
            results.append("\nErrors Encountered:")
            results.append("None")

        final_message = "\n".join(results)
        forms.alert(final_message, title="Updated Instances")
    else:
        forms.alert("No instances were updated.", title="Result")

if __name__ == "__main__":
    main()
