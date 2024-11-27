from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import FilteredElementCollector, FamilyInstance, Transaction

# Initialize document
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# Helper function to get categories
def get_categories(doc):
    categories = set()
    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
    
    for element in collector:
        if element.Category and element.Category.Name:
            categories.add(element.Category.Name)
    
    return sorted(list(categories))

# Helper function to get families in selected categories
def get_families(doc, category_names):
    families = set()
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType()
    
    for fam_instance in collector:
        if fam_instance.Category.Name in category_names:
            families.add(fam_instance.Symbol.Family.Name)
    
    return sorted(list(families))

# Helper function to get family types in the selected families
def get_family_types(doc, category_names, family_names):
    family_types = set()
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType()
    
    for fam_instance in collector:
        if fam_instance.Category.Name in category_names and fam_instance.Symbol.Family.Name in family_names:
            family_types.add(fam_instance.Name)
    
    return sorted(list(family_types))

# Helper function to get all instances of selected family types
def get_family_instances(doc, category_names, family_names, type_names):
    instances = []
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType()
    
    for fam_instance in collector:
        if (fam_instance.Category.Name in category_names 
            and fam_instance.Symbol.Family.Name in family_names 
            and fam_instance.Name in type_names):
            instances.append(fam_instance)
    
    return instances

# Helper function to get integer parameters of the selected element
def get_integer_parameters(element):
    integer_parameters = []
    for param in element.Parameters:
        if param.StorageType == DB.StorageType.Integer:
            integer_parameters.append(param.Definition.Name)
    return integer_parameters

# Main function
def main():
    # Step 1: Select Categories
    categories = get_categories(doc)
    selected_categories = forms.SelectFromList.show(categories, title="Select Categories", multiselect=True)
    if not selected_categories:
        forms.alert("No categories selected.", exitscript=True)
    
    # Step 2: Select Families within the chosen categories
    families = get_families(doc, selected_categories)
    selected_families = forms.SelectFromList.show(families, title="Select Families", multiselect=True)
    if not selected_families:
        forms.alert("No families selected.", exitscript=True)
    
    # Step 3: Select Family Types within the chosen families
    family_types = get_family_types(doc, selected_categories, selected_families)
    selected_family_types = forms.SelectFromList.show(family_types, title="Select Family Types", multiselect=True)
    if not selected_family_types:
        forms.alert("No family types selected.", exitscript=True)

    # Step 4: Get instances of the selected family types
    instances = get_family_instances(doc, selected_categories, selected_families, selected_family_types)
    if not instances:
        forms.alert("No instances found for the selected family types.", exitscript=True)

    # Step 5: Present checklist of instances with Element IDs only
    instance_checklist = []
    instance_dict = {}
    for instance in instances:
        display_text = "Element ID: {0}".format(instance.Id.IntegerValue)
        instance_checklist.append(display_text)
        instance_dict[display_text] = instance

    selected_instances = forms.SelectFromList.show(instance_checklist, title="Select Instances", multiselect=True)
    if not selected_instances:
        forms.alert("No instances selected.", exitscript=True)
    
    selected_elements = [instance_dict[inst] for inst in selected_instances]

    # Step 6: Select integer parameter
    param_options = get_integer_parameters(selected_elements[0])
    selected_param_name = forms.SelectFromList.show(param_options, title="Select Integer Parameter")
    if not selected_param_name:
        forms.alert("No parameter selected.", exitscript=True)

    # Step 7: Input Element ID into the selected integer parameter
    with Transaction(doc, "Input Element ID into Integer Parameter") as t:
        t.Start()
        for element in selected_elements:
            param = element.LookupParameter(selected_param_name)
            if param and param.StorageType == DB.StorageType.Integer:
                param.Set(element.Id.IntegerValue)
        t.Commit()

    forms.alert("Element IDs have been input into the selected parameter.")

# Execute the script
if __name__ == "__main__":
    main()
