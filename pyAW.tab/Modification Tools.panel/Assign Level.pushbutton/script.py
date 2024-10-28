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
    """Retrieve the current level name of a family instance."""
    schedule_level_param = instance.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM)
    if schedule_level_param:
        level_id = schedule_level_param.AsElementId()
        if level_id and level_id != DB.ElementId.InvalidElementId:
            level = doc.GetElement(level_id)
            if level:
                return level.Name

    level_id = instance.LevelId
    if level_id and level_id != DB.ElementId.InvalidElementId:
        level = doc.GetElement(level_id)
        if level:
            return level.Name
    return "None assigned"

def get_nearest_level(doc, instance, direction, selected_levels):
    """Determine the nearest level based on direction and selected levels."""
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
    """Retrieve the Z-axis position of the element."""
    location = element.Location
    if not location or not hasattr(location, 'Point'):
        raise Exception("Element '{0}' does not have a movable location point.".format(element.Name))
    return location.Point.Z

def calculate_new_elevation_from_level(element, target_level_elevation):
    """Calculate the new Elevation from Level to maintain Z-coordinate position."""
    current_z = get_element_z_position(element)
    new_elevation = current_z - target_level_elevation
    return new_elevation

def get_level_by_name(doc, level_name):
    """Retrieve a Level element by its name."""
    collector = FilteredElementCollector(doc).OfClass(Level)
    for level in collector:
        if level.Name == level_name:
            return level
    return None

def set_level_parameters(element, level):
    """Set all level-related parameters of an element to the specified level."""
    schedule_level_param = element.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM)
    if schedule_level_param and not schedule_level_param.IsReadOnly:
        schedule_level_param.Set(level.Id)

def assign_schedule_level_and_elevation(doc, instances, target_level):
    """Assign schedule level and recalculate Elevation from Level for elements."""
    updated_instances = []
    with Transaction(doc, "Assign Schedule Level and Adjust Elevation") as t:
        t.Start()
        for instance in instances:
            # Set Schedule Level
            set_level_parameters(instance, target_level)

            # Calculate and set new Elevation from Level
            new_elevation_from_level = calculate_new_elevation_from_level(instance, target_level.Elevation)
            elevation_param = instance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
            if elevation_param and not elevation_param.IsReadOnly:
                elevation_param.Set(new_elevation_from_level)
                updated_instances.append((instance.Name, target_level.Name, feet_to_meters(new_elevation_from_level)))
            else:
                error_messages.append("No editable Elevation from Level for '{0}'.".format(instance.Name))
        t.Commit()
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

def select_target_level(doc):
    """Prompt user to select a target level."""
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    level_names = [level.Name for level in levels]
    selected_level_name = forms.SelectFromList.show(level_names, title="Select Target Level", multiselect=False)
    if not selected_level_name:
        forms.alert("No level selected.", exitscript=True)
    return get_level_by_name(doc, selected_level_name)

def get_current_selection(doc):
    """Retrieve currently selected FamilyInstance elements."""
    selection = revit.get_selection()
    existing_instances = [el for el in selection if isinstance(el, FamilyInstance)]
    return existing_instances

def assign_levels_and_elevations(doc, instances_to_update, direction, selected_levels):
    """Assign levels and recalculate Elevation from Level for elements."""
    updated_instances = []
    with Transaction(doc, "Assign Nearest Level and Adjust Elevation") as t:
        t.Start()
        for data in instances_to_update:
            instance = data['instance']
            nearest_level = data['desired_level']
            current_level = data['current_level']
            current_offset = data['current_offset']
            
            if not nearest_level:
                error_messages.append("No nearest level found for '{0}'.".format(instance.Name))
                continue

            try:
                # Set Schedule Level
                set_level_parameters(instance, nearest_level)

                # Calculate and set new Elevation from Level
                new_elevation_from_level = calculate_new_elevation_from_level(instance, nearest_level.Elevation)
                elevation_param = instance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
                if elevation_param and not elevation_param.IsReadOnly:
                    elevation_param.Set(new_elevation_from_level)
                    updated_instances.append((instance, current_level, nearest_level.Name, current_offset, feet_to_meters(new_elevation_from_level)))
                else:
                    error_messages.append("No editable Elevation from Level for '{0}'.".format(instance.Name))
            except Exception as e:
                updated_instances.append((instance, current_level, nearest_level.Name, current_offset, "N/A"))
                error_messages.append("Failed to adjust parameters for '{0}': {1}".format(instance.Name, e))
        t.Commit()
    return updated_instances

def main():
    doc = revit.doc

    # Step 1: Choose selection mode
    selection_modes = ["Filter Selection", "Direct Selection"]
    selected_mode = forms.SelectFromList.show(
        selection_modes,
        title="Choose Element Selection Method",
        multiselect=False
    )
    if not selected_mode:
        forms.alert("No selection mode chosen.", exitscript=True)

    if selected_mode == "Filter Selection":
        categories = get_categories(doc)
        selected_category = forms.SelectFromList.show(categories, title="Select Category")
        if not selected_category:
            forms.alert("No category selected.", exitscript=True)

        families = get_families(doc, selected_category)
        selected_families = forms.SelectFromList.show(
            families,
            title="Select Families in {0}".format(selected_category),
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
        instances = get_current_selection(doc)
        if not instances:
            forms.alert("No family instances selected.", exitscript=True)

    selected_levels = select_levels(doc)
    direction = select_direction()

    instances_data = []
    for instance in instances:
        current_level = get_current_level(doc, instance)
        nearest_level = get_nearest_level(doc, instance, direction, selected_levels)
        current_offset = calculate_new_elevation_from_level(instance, nearest_level.Elevation) if nearest_level else "N/A"
        instances_data.append({
            'instance': instance,
            'current_level': current_level,
            'desired_level': nearest_level,
            'desired_level_name': nearest_level.Name if nearest_level else "None found",
            'current_offset': current_offset
        })

    updated_instances = assign_levels_and_elevations(doc, instances_data, direction, selected_levels)

    # Display the results
    results = ["Instance Name | Previous Level | Previous Offset (m) | New Level | New Elevation from Level (m)"]
    for instance, prev_level, new_level, prev_offset, new_offset in updated_instances:
        results.append("{0} | {1} | {2} | {3} | {4}".format(
            instance.Name,
            prev_level,
            prev_offset,
            new_level,
            new_offset
        ))

    if error_messages:
        results.append("\nErrors Encountered:")
        for error in error_messages:
            results.append("- {0}".format(error))

    forms.alert("\n".join(results), title="Updated Instances")

if __name__ == "__main__":
    main()
