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

# Function to get all parameters from the selected instances
def get_all_instance_parameters(instances):
    param_names = set()
    for instance in instances:
        params = instance.Parameters
        for param in params:
            if param.Definition and param.Definition.Name:
                param_names.add(param.Definition.Name)
    return sorted(list(param_names))

# Function to copy parameter values
def copy_parameter_values(doc, instances, source_param_name, target_param_name):
    updated_instances = []
    
    with Transaction(doc, "Copy Parameter Values") as t:
        t.Start()
        for instance in instances:
            source_param = instance.LookupParameter(source_param_name)
            target_param = instance.LookupParameter(target_param_name)
            
            if source_param and target_param:
                source_storage_type = source_param.StorageType
                target_storage_type = target_param.StorageType
                
                if target_storage_type == DB.StorageType.String:
                    # Get the value as formatted string
                    value = source_param.AsValueString()
                    if value is None:
                        # If AsValueString() returns None, fallback to other methods
                        if source_storage_type == DB.StorageType.String:
                            value = source_param.AsString()
                        elif source_storage_type == DB.StorageType.Integer:
                            value = str(source_param.AsInteger())
                        elif source_storage_type == DB.StorageType.Double:
                            value = str(source_param.AsDouble())
                        elif source_storage_type == DB.StorageType.ElementId:
                            eid = source_param.AsElementId()
                            value = str(eid.IntegerValue)
                        else:
                            value = ''
                    target_param.Set(value)
                    updated_instances.append(instance)
                elif source_storage_type == target_storage_type:
                    # Copy the value directly
                    if source_storage_type == DB.StorageType.String:
                        value = source_param.AsString()
                        target_param.Set(value)
                    elif source_storage_type == DB.StorageType.Integer:
                        value = source_param.AsInteger()
                        target_param.Set(value)
                    elif source_storage_type == DB.StorageType.Double:
                        value = source_param.AsDouble()
                        target_param.Set(value)
                    elif source_storage_type == DB.StorageType.ElementId:
                        value = source_param.AsElementId()
                        target_param.Set(value)
                    updated_instances.append(instance)
                else:
                    forms.alert("Cannot assign value to target parameter '{}' on instance '{}'. Storage types do not match.".format(
                        target_param_name, instance.Name))
            else:
                forms.alert("Source or target parameter not found on instance '{}'.".format(instance.Name))
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
    instances = select_elements(doc)

    if not instances:
        forms.alert("No instances selected.", exitscript=True)

    # Step 2: Get all available parameters from the selected instances
    param_names = get_all_instance_parameters(instances)

    # Step 3: Prompt the user to select the source parameter
    source_param_name = forms.SelectFromList.show(
        param_names,
        title="Select Source Parameter",
        multiselect=False
    )
    if not source_param_name:
        forms.alert("No source parameter selected.", exitscript=True)

    # Step 4: Prompt the user to select the target parameter
    target_param_name = forms.SelectFromList.show(
        param_names,
        title="Select Target Parameter",
        multiselect=False
    )
    if not target_param_name:
        forms.alert("No target parameter selected.", exitscript=True)

    # Step 5: Copy parameter values
    updated_instances = copy_parameter_values(doc, instances, source_param_name, target_param_name)

    # Step 6: Display the results
    if updated_instances:
        forms.alert("Parameter values copied from '{}' to '{}' for {} instances.".format(
            source_param_name, target_param_name, len(updated_instances)))
    else:
        forms.alert("No instances were updated.")

# Execute the script
if __name__ == "__main__":
    main()
