from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    FamilyInstance,
    Transaction,
    Level,
    BuiltInParameter,
    XYZ,
    Plane,
    SketchPlane
)

# Helper functions remain the same as before...

def get_categories(doc):
    # ... (same as before)
    categories = set()
    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
    for element in collector:
        if element.Category and element.Category.Name:
            categories.add(element.Category.Name)
    return sorted(list(categories))

def get_families(doc, category_name):
    # ... (same as before)
    families = set()
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance)
    for fam_instance in collector:
        if fam_instance.Category.Name == category_name:
            families.add(fam_instance.Symbol.Family.Name)
    return sorted(list(families))

def get_family_types(doc, category_name, family_names):
    # ... (same as before)
    family_types = set()
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance)
    for fam_instance in collector:
        if (fam_instance.Category.Name == category_name and
            fam_instance.Symbol.Family.Name in family_names):
            family_types.add(fam_instance.Name)
    return sorted(list(family_types))

def get_family_instances(doc, category_name, family_names, type_names):
    # ... (same as before)
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
    # Modified to filter by selected levels
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

def main():
    doc = revit.doc

    # Step 1: Select elements
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