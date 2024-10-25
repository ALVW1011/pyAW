from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    FamilyInstance,
    Transaction,
    Level,
    BuiltInParameter,
    XYZ,
    Plane,
    SketchPlane,
    ViewType
)

# Helper functions remain the same as before...

def get_categories(doc):
    categories = set()
    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
    for element in collector:
        if element.Category and element.Category.Name:
            categories.add(element.Category.Name)
    return sorted(list(categories))

def get_families(doc, category_name):
    families = set()
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance)
    for fam_instance in collector:
        if fam_instance.Category.Name == category_name:
            families.add(fam_instance.Symbol.Family.Name)
    return sorted(list(families))

def get_family_types(doc, category_name, family_names):
    family_types = set()
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance)
    for fam_instance in collector:
        if (fam_instance.Category.Name == category_name and
            fam_instance.Symbol.Family.Name in family_names):
            family_types.add(fam_instance.Name)
    return sorted(list(family_types))

def get_family_instances(doc, category_name, family_names, type_names):
    instances = []
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance)
    for fam_instance in collector:
        if (fam_instance.Category.Name == category_name and
            fam_instance.Symbol.Family.Name in family_names and
            fam_instance.Name in type_names):
            instances.append(fam_instance)
    return instances

def get_current_level(doc, instance):
    # Try INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM
    schedule_level_param = instance.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM)
    if schedule_level_param:
        level_id = schedule_level_param.AsElementId()
        if level_id and level_id != DB.ElementId.InvalidElementId:
            level = doc.GetElement(level_id)
            if level:
                return level.Name

    # Try LevelId
    level_id = instance.LevelId
    if level_id and level_id != DB.ElementId.InvalidElementId:
        level = doc.GetElement(level_id)
        if level:
            return level.Name

    # Check 'Reference Level' or 'Work Plane' parameters
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
    if hasattr(instance.Location, 'Point'):
        instance_location = instance.Location.Point
        nearest_level = None
        nearest_distance = float('inf')
        for level in selected_levels:
            level_elevation = level.Elevation
            level_name = level.Name

            if direction == "Below":
                distance = instance_location.Z - level_elevation
                if distance >= 0 and distance < nearest_distance:
                    nearest_distance = distance
                    nearest_level = level
            elif direction == "Above":
                distance = level_elevation - instance_location.Z
                if distance >= 0 and distance < nearest_distance:
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

def assign_nearest_levels(doc, instances_to_update, direction):
    updated_instances = []
    with Transaction(doc, "Assign Nearest Level") as t:
        try:
            t.Start()
            for instance_data in instances_to_update:
                instance = instance_data['instance']
                nearest_level = instance_data['desired_level']
                current_level_name = instance_data['current_level']

                if nearest_level:
                    # Attempt to set INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM
                    schedule_level_param = instance.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM)
                    if schedule_level_param and not schedule_level_param.IsReadOnly:
                        schedule_level_param.Set(nearest_level.Id)
                        updated_instances.append((instance, current_level_name, nearest_level.Name))
                    else:
                        # As a fallback, try to set FAMILY_LEVEL_PARAM
                        family_level_param = instance.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
                        if family_level_param and not family_level_param.IsReadOnly:
                            family_level_param.Set(nearest_level.Id)
                            updated_instances.append((instance, current_level_name, nearest_level.Name))
                        else:
                            # Unable to set level
                            pass
            t.Commit()
        except Exception as e:
            t.RollBack()
            forms.alert("An error occurred: {}".format(e), title="Error")
    return updated_instances

def select_direction():
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
    Retrieves the currently selected FamilyInstance elements in Revit.
    
    Args:
        doc: The active Revit document.
    
    Returns:
        A list of currently selected FamilyInstance elements.
    """
    # Use pyRevit's helper function to get the current selection
    selection = revit.get_selection()
    # Filter only FamilyInstance elements
    existing_instances = [el for el in selection if isinstance(el, FamilyInstance)]
    return existing_instances

def select_elements_directly(doc):
    """
    Allows users to directly select elements from a graphical view, 
    adding to any elements that were already selected before running the tool.
    
    Args:
        doc: The active Revit document.
    
    Returns:
        A combined list of pre-selected and newly selected FamilyInstance elements.
    """
    active_view = doc.ActiveView
    # Define which ViewTypes are considered graphical
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

    # Step 1: Retrieve existing selection
    existing_instances = get_current_selection(doc)
    existing_count = len(existing_instances)
    
    if existing_count > 0:
        forms.alert(
            "Found {0} pre-selected family instance(s). You can add more by selecting additional elements.".format(existing_count),
            title="Existing Selection",
            warn_icon=True
        )

    # Step 2: Prompt user to select additional elements
    try:
        new_selected_elements = revit.pick_elements(
            message="Select additional elements from the model:"
        )
    except Exception as e:
        forms.alert(
            "An error occurred during element selection:\n{}".format(e),
            title="Selection Error",
            exitscript=True
        )

    if not new_selected_elements:
        if existing_count == 0:
            forms.alert("No elements selected.", exitscript=True)
        else:
            # Only existing elements are considered
            return existing_instances

    # Step 3: Filter only FamilyInstance elements
    new_instances = [el for el in new_selected_elements if isinstance(el, FamilyInstance)]
    if not new_instances:
        if existing_count == 0:
            forms.alert("No family instances selected in the new selection.", exitscript=True)
        else:
            # Only existing elements are considered
            return existing_instances

    # Step 4: Combine existing and new instances, avoiding duplicates based on ElementId
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
        # Existing filter selection process
        # Step 1: Select elements via filters
        categories = get_categories(doc)
        selected_category = forms.SelectFromList.show(categories, title="Select Category")
        if not selected_category:
            forms.alert("No category selected.", exitscript=True)

        families = get_families(doc, selected_category)
        selected_families = forms.SelectFromList.show(
            families,
            title="Select Families in {}".format(selected_category),
            multiselect=True
        )
        if not selected_families:
            forms.alert("No families selected.", exitscript=True)

        family_types = get_family_types(doc, selected_category, selected_families)
        selected_family_types = forms.SelectFromList.show(
            family_types,
            title="Select Family Types in Selected Families",
            multiselect=True
        )
        if not selected_family_types:
            forms.alert("No family types selected.", exitscript=True)

        instances = get_family_instances(doc, selected_category, selected_families, selected_family_types)
        if not instances:
            forms.alert("No instances found for the selected family types.", exitscript=True)
    else:
        # New direct selection process with enhanced functionality
        instances = select_elements_directly(doc)

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
        instances_data.append({
            'instance': instance,
            'current_level': current_level,
            'desired_level': nearest_level,
            'desired_level_name': desired_level_name
        })

    # Step 5: Display the list and confirm
    results = ["Instance Name | Current Level | Desired Level"]
    for data in instances_data:
        results.append("{0} | {1} | {2}".format(
            data['instance'].Name,
            data['current_level'],
            data['desired_level_name']
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

    # Step 6: Assign nearest levels
    updated_instances = assign_nearest_levels(doc, instances_data, direction)

    # Step 7: Display the results
    if updated_instances:
        results = ["Instance Name | Previous Level | New Level"]
        for instance, prev_level, new_level in updated_instances:
            results.append("{0} | {1} | {2}".format(instance.Name, prev_level, new_level))
        forms.alert("\n".join(results), title="Updated Instances")
    else:
        forms.alert("No instances were updated.", title="Result")

if __name__ == "__main__":
    main()
