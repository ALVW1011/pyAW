# -*- coding: utf-8 -*-
"""
PyRevit Tool: Copy Family Instances from Linked Models to Host Model with Step-by-Step Selection

This script allows users to:
1. Select linked Revit models.
2. Select categories within the selected linked models.
3. Select family types within the selected categories.
4. Select instances within the selected family types.
5. Copy and paste the selected instances in place into the host model.

Compatible with Revit 2020 and PyRevit.
"""

from pyrevit import revit, DB, forms
import clr
import sys
from System.Collections.Generic import List as Clist  # Import .NET List for collections

# Initialize the active document and UI document
uidoc = revit.uidoc
doc = revit.doc

def get_linked_models(doc):
    """
    Retrieves all linked Revit models in the current project.
    Returns a list of tuples: (Link Name, RevitLinkInstance)
    """
    links = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance).ToElements()
    link_info = []
    for link in links:
        link_name = link.Name
        if link_name:
            link_info.append((link_name, link))
    return link_info

def get_categories_from_links(link_instances):
    """
    Retrieves all unique categories from the selected linked models.
    Returns a sorted list of category names.
    """
    categories = set()
    for link in link_instances:
        linked_doc = link.GetLinkDocument()
        if not linked_doc:
            continue
        collector = DB.FilteredElementCollector(linked_doc).WhereElementIsNotElementType()
        for element in collector:
            if element.Category and element.Category.Name:
                categories.add(element.Category.Name)
    return sorted(categories)

def get_family_symbols_from_links(link_instances, selected_categories):
    """
    Retrieves all unique family symbols (types) from the selected linked models and categories.
    Returns a dictionary mapping display names to FamilySymbol elements.
    """
    family_symbols = {}
    for link in link_instances:
        linked_doc = link.GetLinkDocument()
        if not linked_doc:
            continue
        collector = DB.FilteredElementCollector(linked_doc).OfClass(DB.FamilyInstance).WhereElementIsNotElementType()
        for inst in collector:
            try:
                cat = inst.Category
                if cat and cat.Name in selected_categories:
                    symbol = inst.Symbol
                    if symbol:
                        # Safely get family name
                        if symbol.Family and hasattr(symbol.Family, 'Name'):
                            family_name = symbol.Family.Name
                        else:
                            family_name = "Unknown Family"

                        # Safely get type name
                        if hasattr(symbol, 'Name'):
                            type_name = symbol.Name
                        else:
                            type_name = "Unknown Type"

                        display_name = "{0} : {1}".format(family_name, type_name)
                        family_symbols[display_name] = symbol
            except Exception as e:
                # Log the error and continue
                print("Error processing symbol in linked document '{}': {}".format(linked_doc.Title, e))
    return family_symbols

def get_family_instances_from_links(link_instances, selected_categories, selected_family_symbols):
    """
    Retrieves all family instances from the selected linked models, categories, and family symbols.
    Returns a list of FamilyInstance objects.
    """
    instances = []
    symbol_ids = [symbol.Id for symbol in selected_family_symbols.values()]
    for link in link_instances:
        linked_doc = link.GetLinkDocument()
        if not linked_doc:
            continue
        collector = DB.FilteredElementCollector(linked_doc).OfClass(DB.FamilyInstance).WhereElementIsNotElementType()
        for inst in collector:
            try:
                cat = inst.Category
                if cat and cat.Name in selected_categories and inst.Symbol:
                    if inst.Symbol.Id in symbol_ids:
                        instances.append(inst)
            except Exception as e:
                # Log the error and continue
                print("Error processing instance in linked document '{}': {}".format(linked_doc.Title, e))
    return instances

def present_selection(title, options, multiselect=True):
    """
    Presents a selection dialog to the user.
    Returns the list of selected items or None if canceled.
    """
    return forms.SelectFromList.show(
        options,
        title=title,
        button_name="Select",
        multiselect=multiselect
    )

def main():
    # Step 1: Select Linked Models
    link_info = get_linked_models(doc)
    if not link_info:
        forms.alert("No linked models found in the current project.", exitscript=True)

    link_names = [name for name, _ in link_info]
    selected_link_names = present_selection(
        title="Select Linked Models",
        options=link_names,
        multiselect=True
    )
    if not selected_link_names:
        forms.alert("No linked models selected.", exitscript=True)

    # Map selected link names to their instances
    selected_links = []
    for name in selected_link_names:
        for lname, link in link_info:
            if lname == name:
                selected_links.append(link)
                break

    # Step 2: Select Categories within the selected linked models
    categories = get_categories_from_links(selected_links)
    if not categories:
        forms.alert("No categories found in the selected linked models.", exitscript=True)

    selected_categories = present_selection(
        title="Select Categories",
        options=categories,
        multiselect=True
    )
    if not selected_categories:
        forms.alert("No categories selected.", exitscript=True)

    # Step 3: Select Family Types within the selected categories
    family_symbols = get_family_symbols_from_links(selected_links, selected_categories)
    if not family_symbols:
        forms.alert("No family types found in the selected categories.", exitscript=True)

    family_symbol_names = sorted(family_symbols.keys())
    selected_family_symbol_names = present_selection(
        title="Select Family Types",
        options=family_symbol_names,
        multiselect=True
    )
    if not selected_family_symbol_names:
        forms.alert("No family types selected.", exitscript=True)

    # Map selected names to FamilySymbol elements
    selected_family_symbols = {name: family_symbols[name] for name in selected_family_symbol_names}

    # Step 4: Select Family Instances within the selected family types
    selected_instances = get_family_instances_from_links(selected_links, selected_categories, selected_family_symbols)
    if not selected_instances:
        forms.alert("No family instances found for the selected criteria.", exitscript=True)

    # Prepare display list for instances
    instance_display = []
    instance_dict = {}
    for inst in selected_instances:
        try:
            # Attempt to find the link name for better identification
            link_name = "Unknown Link"
            for link in selected_links:
                linked_doc = link.GetLinkDocument()
                if linked_doc and linked_doc.Equals(inst.Document):
                    link_name = link.Name
                    break

            # Safely get family name
            if inst.Symbol and inst.Symbol.Family and hasattr(inst.Symbol.Family, 'Name'):
                family_name = inst.Symbol.Family.Name
            else:
                family_name = "Unknown Family"

            # Safely get type name
            if inst.Symbol and hasattr(inst.Symbol, 'Name'):
                type_name = inst.Symbol.Name
            else:
                type_name = "Unknown Type"

            display_str = "Link: {0} | Element ID: {1} | Family: {2} | Type: {3}".format(
                link_name,
                inst.Id.IntegerValue,
                family_name,
                type_name
            )
            instance_display.append(display_str)
            instance_dict[display_str] = inst
        except Exception as e:
            # Log the error and continue
            print("Error preparing instance display: {}".format(e))

    # Present checklist of instances
    selected_instances_display = present_selection(
        title="Select Instances to Copy",
        options=instance_display,
        multiselect=True
    )
    if not selected_instances_display:
        forms.alert("No instances selected.", exitscript=True)

    # Retrieve the actual family instances
    final_selected_instances = [instance_dict[disp] for disp in selected_instances_display]

    # Step 5: Copy and Paste Instances into Host Model
    try:
        # Group elements by their linked document
        link_to_elements = {}
        for inst in final_selected_instances:
            try:
                # Identify which link this instance belongs to
                source_link = None
                for link in selected_links:
                    linked_doc = link.GetLinkDocument()
                    if linked_doc and linked_doc.Equals(inst.Document):
                        source_link = link
                        break
                if source_link:
                    if source_link not in link_to_elements:
                        link_to_elements[source_link] = []
                    link_to_elements[source_link].append(inst.Id)
                else:
                    # Handle elements not from any link (unlikely in this context)
                    msg = "Warning: Element ID {} does not belong to any selected linked model.".format(inst.Id.IntegerValue)
                    forms.alert(msg, title="Warning", warn_icon=True)
            except Exception as e:
                print("Error grouping elements by linked document: {}".format(e))

        # Initialize a counter for copied elements
        total_copied = 0

        # Start a transaction
        with revit.Transaction("Copy Family Instances from Link"):
            for link, ids in link_to_elements.items():
                # Get the source document
                source_doc = link.GetLinkDocument()
                if not source_doc:
                    forms.alert("Failed to get the linked document: {}".format(link.Name))
                    continue

                # Prepare the copy paste options
                options = DB.CopyPasteOptions()

                # Convert ids to .NET ICollection[ElementId]
                ids_col = Clist[DB.ElementId](ids)

                # Perform the copy
                try:
                    copied_elements = DB.ElementTransformUtils.CopyElements(
                        source_doc,
                        ids_col,
                        doc,
                        DB.Transform.Identity,
                        options
                    )

                    if copied_elements:
                        total_copied += len(copied_elements)
                except Exception as e:
                    forms.alert("Error copying elements from link '{}': {}".format(link.Name, e))

        if total_copied == 0:
            forms.alert("No elements were copied. Please check the selection criteria.", exitscript=True)
            return

        forms.alert("Successfully copied and pasted {} family instances.".format(total_copied))

    except Exception as e:
        # Capture any unexpected errors and display them
        error_message = "An error occurred during the copy-paste process: {}".format(e)
        forms.alert(error_message, title="Error", warn_icon=True)

# Execute the script
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # Catch any unexpected exceptions and provide a user-friendly message
        error_message = "An unexpected error occurred: {}".format(e)
        forms.alert(error_message, title="Unexpected Error", warn_icon=True)