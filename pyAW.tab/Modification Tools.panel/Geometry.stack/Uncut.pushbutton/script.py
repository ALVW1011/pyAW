from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import FamilyInstance, Transaction, InstanceVoidCutUtils

# Initialize lists for error messages and successful uncut operations
error_messages = []
successful_uncut_instances = []

def get_categories(doc):
    """Retrieve all unique category names in the document."""
    categories = set()
    collector = DB.FilteredElementCollector(doc).WhereElementIsNotElementType()
    for element in collector:
        if element.Category and element.Category.Name:
            categories.add(element.Category.Name)
    return sorted(list(categories))

def get_families(doc, category_name):
    """Retrieve all unique family names within a specific category."""
    families = set()
    collector = DB.FilteredElementCollector(doc).OfClass(FamilyInstance)
    for fam_instance in collector:
        if fam_instance.Category and fam_instance.Category.Name == category_name:
            families.add(fam_instance.Symbol.Family.Name)
    return sorted(list(families))

def get_family_types(doc, category_name, family_names):
    """Retrieve all unique family types within selected families."""
    family_types = set()
    collector = DB.FilteredElementCollector(doc).OfClass(FamilyInstance)
    for fam_instance in collector:
        if (fam_instance.Category and fam_instance.Category.Name == category_name and
            fam_instance.Symbol.Family.Name in family_names):
            family_types.add(fam_instance.Name)
    return sorted(list(family_types))

def get_family_instances(doc, category_name, family_names, type_names):
    """Retrieve all family instances based on selected category, families, and types."""
    instances = []
    collector = DB.FilteredElementCollector(doc).OfClass(FamilyInstance)
    for fam_instance in collector:
        if (fam_instance.Category and fam_instance.Category.Name == category_name and
            fam_instance.Symbol.Family.Name in family_names and
            fam_instance.Name in type_names):
            instances.append(fam_instance)
    return instances

def get_current_selection():
    """Retrieve currently selected elements."""
    selection = revit.get_selection()
    # Ensure that only FamilyInstances are processed
    instances = [elem for elem in selection if isinstance(elem, FamilyInstance)]
    return instances

def get_cut_elements(doc, instance):
    """Retrieve elements that the instance is cutting."""
    try:
        if InstanceVoidCutUtils.IsVoidInstanceCuttingElement(instance):
            cut_element_ids = InstanceVoidCutUtils.GetElementsBeingCut(instance)
            cut_elements = [doc.GetElement(id) for id in cut_element_ids]
            return cut_elements
        else:
            return []
    except Exception as e:
        # Log or handle the exception
        return []

def uncut_geometry(doc, instance):
    """Attempts to uncut geometry from the selected instance."""
    # Use InstanceVoidCutUtils.GetElementsBeingCut to find the elements being cut by the instance
    cut_elements = get_cut_elements(doc, instance)

    for cut_element in cut_elements:
        try:
            # Attempt to remove the cut relationship
            InstanceVoidCutUtils.RemoveInstanceVoidCut(doc, instance, cut_element)
            successful_uncut_instances.append((instance, cut_element))
        except Exception as e:
            error_messages.append("Failed to uncut geometry between {0} and {1}: {2}".format(
                instance.Id, cut_element.Id, str(e)))

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
        # Select Category
        categories = get_categories(doc)
        selected_category = forms.SelectFromList.show(categories, title="Select Category")
        if not selected_category:
            forms.alert("No category selected.", exitscript=True)

        # Select Families
        families = get_families(doc, selected_category)
        selected_families = forms.SelectFromList.show(
            families,
            title="Select Families in {0}".format(selected_category),
            multiselect=True
        )
        if not selected_families:
            forms.alert("No families selected.", exitscript=True)

        # Select Family Types
        family_types = get_family_types(doc, selected_category, selected_families)
        selected_family_types = forms.SelectFromList.show(
            family_types,
            title="Select Family Types in Selected Families",
            multiselect=True
        )
        if not selected_family_types:
            forms.alert("No family types selected.", exitscript=True)

        # Retrieve Instances
        instances = get_family_instances(doc, selected_category, selected_families, selected_family_types)
        if not instances:
            forms.alert("No instances found for the selected family types.", exitscript=True)
    else:
        # Direct Selection
        instances = get_current_selection()
        if not instances:
            forms.alert("No family instances selected.", exitscript=True)

    # Step 2: Debugging - List selected instances and corresponding cut elements
    instance_info = []
    for instance in instances:
        cut_elements = get_cut_elements(doc, instance)
        cut_element_info = ["ID: {0}, Name: {1}".format(cut_elem.Id, cut_elem.Name) for cut_elem in cut_elements]
        instance_info.append("Instance ID: {0}, Name: {1}, Cutting Elements: \n  {2}".format(
            instance.Id, instance.Name, "\n  ".join(cut_element_info) if cut_element_info else "None"
        ))
    instance_info_message = "The following instances were selected along with their cutting relationships:\n\n" + "\n\n".join(instance_info)
    instance_info_message += "\n\nTotal selected instances: {0}".format(len(instances))
    
    # Confirm to proceed
    if not forms.alert(instance_info_message, title="Confirm Instances to Uncut", ok=True, cancel=True):
        forms.alert("Process canceled by user.", exitscript=True)

    # Step 3: Auto-Uncut Geometry
    with Transaction(doc, "Auto-Uncut Geometry") as t:
        t.Start()
        for instance in instances:
            uncut_geometry(doc, instance)
        t.Commit()

    # Step 4: Reporting
    # Report successful uncut operations
    if successful_uncut_instances:
        success_info = [
            "Uncut between Instance ID: {0}, Name: {1} and Cut Element ID: {2}, Name: {3}".format(
                instance.Id, instance.Name, cut_elem.Id, cut_elem.Name
            ) for (instance, cut_elem) in successful_uncut_instances
        ]
        success_info_message = "The following uncut operations were successful:\n\n" + "\n".join(success_info)
        success_info_message += "\n\nTotal successfully uncut operations: {0}".format(len(successful_uncut_instances))
        forms.alert(success_info_message, title="Uncut Operations Completed", exitscript=False)
    else:
        forms.alert("No uncut operations were performed.", title="Uncut Operations Completed", exitscript=False)

    # Report errors, if any
    if error_messages:
        forms.alert("\n".join(error_messages), title="Errors Encountered", exitscript=False)

if __name__ == "__main__":
    main()
