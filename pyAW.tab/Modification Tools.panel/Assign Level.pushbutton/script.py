from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import FilteredElementCollector, FamilyInstance, Transaction, Level

# Helper function to get all categories in the entire project (not limited to view)
def get_categories(doc):
    categories = set()
    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()  # Not restricted to view
    
    for element in collector:
        if element.Category and element.Category.Name:
            categories.add(element.Category.Name)
    
    # Sort categories alphabetically
    return sorted(list(categories))

# Helper function to get all families in the selected category across the project
def get_families(doc, category_name):
    families = set()
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType()
    
    for fam_instance in collector:
        if fam_instance.Category.Name == category_name:
            families.add(fam_instance.Symbol.Family.Name)
    
    # Sort families alphabetically
    return sorted(list(families))

# Helper function to get all family types in the selected families across the project
def get_family_types(doc, category_name, family_names):
    family_types = set()  # Using a set to store unique family types
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType()
    
    for fam_instance in collector:
        if (fam_instance.Category.Name == category_name 
            and fam_instance.Symbol.Family.Name in family_names):
            family_types.add(fam_instance.Name)  # Add only the family type (ignoring IDs here)
    
    # Sort family types alphabetically
    return sorted(list(family_types))

# Helper function to collect all instances of selected family types from selected families
def get_family_instances(doc, category_name, family_names, type_names):
    instances = []
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType()
    
    for fam_instance in collector:
        if (fam_instance.Category.Name == category_name 
            and fam_instance.Symbol.Family.Name in family_names
            and fam_instance.Name in type_names):
            instances.append(fam_instance)
    
    return instances

# Function to get current level of an instance
def get_current_level(doc, instance):
    schedule_level_param = instance.get_Parameter(DB.BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM)
    if schedule_level_param:
        level_id = schedule_level_param.AsElementId()
        if level_id != DB.ElementId.InvalidElementId:
            level = doc.GetElement(level_id)
            return level.Name
    return "None assigned"

# Function to get nearest level below the instance, excluding levels to ignore
def get_nearest_level(doc, instance, levels_to_ignore):
    if hasattr(instance.Location, 'Point'):
        instance_location = instance.Location.Point
        levels = FilteredElementCollector(doc).OfClass(Level)

        # Find the nearest level below the instance
        nearest_level = None
        nearest_distance = float('inf')
        
        for level in levels:
            level_name = level.Name
            if level_name in levels_to_ignore:
                continue  # Skip levels to ignore
            level_elevation = level.Elevation
            if instance_location.Z > level_elevation:
                distance = instance_location.Z - level_elevation
                if distance < nearest_distance:
                    nearest_distance = distance
                    nearest_level = level
        
        return nearest_level
    else:
        # If the family instance has no location, we can't determine the level
        return None

# Function to assign the nearest level to the selected instances
def assign_nearest_levels(doc, selected_instances, levels_to_ignore):
    updated_instances = []
    
    with Transaction(doc, "Assign Nearest Level") as t:
        t.Start()
        
        for instance in selected_instances:
            # Get current schedule level parameter
            schedule_level_param = instance.get_Parameter(DB.BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM)
            
            if schedule_level_param:
                current_level = schedule_level_param.AsElementId()
                current_level_name = "None assigned" if current_level == DB.ElementId.InvalidElementId else doc.GetElement(current_level).Name
                
                # Find nearest level
                nearest_level = get_nearest_level(doc, instance, levels_to_ignore)
                if nearest_level:
                    schedule_level_param.Set(nearest_level.Id)
                    updated_instances.append((instance, current_level_name, nearest_level.Name))
        
        t.Commit()
    
    return updated_instances

# Step-by-step element selection
def select_elements(doc):
    # Step 1: Select Category
    categories = get_categories(doc)
    selected_category = forms.SelectFromList.show(categories, title="Select Category")
    if not selected_category:
        forms.alert("No category selected.", exitscript=True)
    
    # Step 2: Select Families within the chosen category (allowing multiple selections)
    families = get_families(doc, selected_category)
    selected_families = forms.SelectFromList.show(
        families,
        title="Select Families in {0}".format(selected_category),
        multiselect=True
    )
    if not selected_families:
        forms.alert("No families selected.", exitscript=True)
    
    # Step 3: Select Family Types within the chosen families (allowing multiple selections)
    family_types = get_family_types(doc, selected_category, selected_families)
    selected_family_types = forms.SelectFromList.show(
        family_types,
        title="Select Family Types in Selected Families",
        multiselect=True
    )
    if not selected_family_types:
        forms.alert("No family types selected.", exitscript=True)

    # Step 4: Select Instances within the chosen family types
    instances = get_family_instances(doc, selected_category, selected_families, selected_family_types)
    if not instances:
        forms.alert("No instances found for the selected family types.", exitscript=True)

    # Return the selected instances
    return instances

# Main function
def main():
    doc = revit.doc
    
    # Step 1: Select elements (no model selection)
    primary_elements = select_elements(doc)
    
    if not primary_elements:
        forms.alert("No primary elements selected.", exitscript=True)
    
    # Step 2: Generate a checklist for user selection with current levels
    element_checklist = []
    for instance in primary_elements:
        current_level = get_current_level(doc, instance)
        element_checklist.append("{0} (Current Level: {1})".format(instance.Name, current_level))
    
    # Let the user select which instances to update
    selected_instance_names = forms.SelectFromList.show(
        element_checklist,
        title="Select Instances to Update",
        multiselect=True
    )
    
    if not selected_instance_names:
        forms.alert("No instances selected for updating.", exitscript=True)
    
    # Filter the selected instances based on the names selected in the checklist
    selected_instances = [
        instance for instance in primary_elements
        if "{0} (Current Level: {1})".format(instance.Name, get_current_level(doc, instance)) in selected_instance_names
    ]
    
    # Step 3: Let the user select levels to ignore
    levels = FilteredElementCollector(doc).OfClass(Level)
    level_names = [level.Name for level in levels]
    levels_to_ignore = forms.SelectFromList.show(
        level_names,
        title="Select Levels to Ignore",
        multiselect=True
    )
    if levels_to_ignore is None:
        levels_to_ignore = []
    
    # Step 4: Assign nearest levels to the selected instances
    updated_instances = assign_nearest_levels(doc, selected_instances, levels_to_ignore)
    
    # Step 5: Display the results
    if updated_instances:
        results = ["Instance Name | Current Level | New Level"]
        for instance, current_level, new_level in updated_instances:
            results.append("{0} | {1} | {2}".format(instance.Name, current_level, new_level))
        
        forms.alert("\n".join(results))
    else:
        forms.alert("No instances were updated.")

# Execute the script
if __name__ == "__main__":
    main()
