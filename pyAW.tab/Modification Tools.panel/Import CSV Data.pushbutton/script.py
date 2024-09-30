import Autodesk.Revit
from pyrevit import forms, script, revit, DB
import os
import csv



# Initialize the output log for showing results
output = script.get_output()
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# Step 1: Let the user select a CSV file
file_path = forms.pick_file(file_ext='csv', title="Select a CSV Clash Report File")

if not file_path:
    forms.alert("No file selected. Exiting script.", exitscript=True)

# Step 2: Read the CSV file to extract the headers
try:
    # Use csv.Sniffer to detect the dialect and ensure correct parsing
    with open(file_path, 'r') as f:
        sample = f.read(1024)
        f.seek(0)
        dialect = csv.Sniffer().sniff(sample)
        reader = csv.reader(f, dialect)
        headers = next(reader)
except Exception as e:
    forms.alert("Error reading CSV file: {}".format(str(e)), exitscript=True)

# Step 3: Let the user select headers for Primary and Intersecting Elements
primary_header = forms.SelectFromList.show(headers, title="Select Primary Element ID Column")
if not primary_header:
    forms.alert("No Primary Element ID column selected. Exiting script.", exitscript=True)

intersecting_header = forms.SelectFromList.show(headers, title="Select Intersecting Element ID Column")
if not intersecting_header:
    forms.alert("No Intersecting Element ID column selected. Exiting script.", exitscript=True)

# Step 3a: Let the user select the document (host or linked) for each column
# Collect list of documents: host and linked documents
documents = []
document_names = []

# Add the host document
documents.append(doc)
document_names.append("Host Document ({})".format(doc.Title))

# Collect linked documents
link_instances = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance)
for link_instance in link_instances:
    link_doc = link_instance.GetLinkDocument()
    if link_doc:
        documents.append(link_doc)
        # Use the link instance name or link document title as display name
        link_name = link_instance.Name  # This includes the linked file name and instance name
        document_names.append("Linked Document ({})".format(link_name))

# Verify that documents and document_names are aligned
output.print_md("Available Documents:")
for idx, doc_name in enumerate(document_names):
    output.print_md("  {}: {}".format(idx, doc_name))

# Ask the user to select the document for primary elements
primary_doc_choice = forms.SelectFromList.show(document_names, title="Select Document for Primary Elements")
if not primary_doc_choice:
    forms.alert("No document selected for Primary Elements. Exiting script.", exitscript=True)
primary_doc_index = document_names.index(primary_doc_choice)
primary_doc = documents[primary_doc_index]
output.print_md("Selected primary document: {}".format(primary_doc_choice))

# Ask the user to select the document for intersecting elements
intersecting_doc_choice = forms.SelectFromList.show(document_names, title="Select Document for Intersecting Elements")
if not intersecting_doc_choice:
    forms.alert("No document selected for Intersecting Elements. Exiting script.", exitscript=True)
intersecting_doc_index = document_names.index(intersecting_doc_choice)
intersecting_doc = documents[intersecting_doc_index]
output.print_md("Selected intersecting document: {}".format(intersecting_doc_choice))

# Step 4: Process the data from the CSV
data = []
try:
    # Use the same dialect for DictReader
    with open(file_path, 'r') as f:
        reader = csv.DictReader(f, dialect=dialect)
        for row in reader:
            if row[primary_header] and row[intersecting_header]:
                data.append((row[primary_header], row[intersecting_header]))
except Exception as e:
    forms.alert("Error processing CSV data: {}".format(str(e)), exitscript=True)

# Output total number of rows processed
output.print_md("Total rows processed from CSV: {}".format(len(data)))

# Step 5: Group by primary elements and combine intersecting elements
grouped_data = {}
for primary_id_str, intersecting_id_str in data:
    if primary_id_str not in grouped_data:
        grouped_data[primary_id_str] = set()
    grouped_data[primary_id_str].add(intersecting_id_str)

# Step 6: Detect the primary elements in the selected document based on Element IDs
primary_element_ids = list(grouped_data.keys())
primary_elements = []
for element_id_str in primary_element_ids:
    try:
        element_id = DB.ElementId(int(element_id_str))
        element = primary_doc.GetElement(element_id)
        if element:
            primary_elements.append(element)
        else:
            output.print_md("Element ID {} not found in the selected primary document.".format(element_id_str))
    except Exception as e:
        output.print_md("Error processing Element ID {}: {}".format(element_id_str, str(e)))

# Output total number of primary elements found
output.print_md("Total primary elements found: {}".format(len(primary_elements)))

# If no valid primary elements were detected, exit the script
if not primary_elements:
    forms.alert("No valid primary elements found in the selected primary document. Exiting script.", exitscript=True)

# Step 7: Show intermediate output for primary elements and their intersecting elements
processed_data = []
for primary_element in primary_elements:
    primary_element_id = str(primary_element.Id.IntegerValue)
    intersecting_element_ids = list(grouped_data[primary_element_id])
    processed_data.append([primary_element_id, ", ".join(intersecting_element_ids)])

# Add feedback before printing table
output.print_md("Processing complete. Showing the intermediate output for primary and intersecting elements.")

try:
    output.print_table(processed_data, title="Primary Elements and Intersecting Elements", columns=["Primary Element ID", "Intersecting Element IDs"])
except Exception as e:
    output.print_md("Error displaying table: {}".format(str(e)))
    forms.alert("Failed to display table. Exiting script.", exitscript=True)

# Step 8: Pause for user confirmation before proceeding
proceed_options = ['Yes', 'No']
proceed_choice = forms.CommandSwitchWindow.show(proceed_options, message="Proceed with data collection?", title="Confirmation")

if proceed_choice != 'Yes':
    forms.alert("Script cancelled.", exitscript=True)

# Step 9: Allow the user to select the parameter in primary elements to update
def get_text_or_multiline_parameters(element):
    params = []
    for param in element.Parameters:
        if param.StorageType == DB.StorageType.String and param.Definition.ParameterType in [
            DB.ParameterType.Text, DB.ParameterType.MultilineText]:
            if not param.IsReadOnly:
                params.append(param.Definition.Name)
    return params

# Getting parameters from the first primary element (assuming the same for all)
available_primary_params = get_text_or_multiline_parameters(primary_elements[0])
if not available_primary_params:
    forms.alert("No writable text or multiline text parameters found in primary elements. Exiting script.", exitscript=True)

selected_primary_param = forms.SelectFromList.show(
    available_primary_params,
    title="Select parameter to update in Primary Elements"
)

if not selected_primary_param:
    forms.alert("No parameter selected from primary elements. Exiting script.", exitscript=True)

# Step 10: Collect parameters from the intersecting elements
intersecting_param_names = set()

# For the first intersecting element (for simplicity), collect parameters
first_primary_element_id = str(primary_elements[0].Id.IntegerValue)
first_intersecting_element_id = list(grouped_data[first_primary_element_id])[0]
intersecting_element = None

# Try to get the intersecting element from the selected intersecting document
try:
    element_id = DB.ElementId(int(first_intersecting_element_id))
    intersecting_element = intersecting_doc.GetElement(element_id)
except Exception as e:
    output.print_md("Error getting intersecting element ID {}: {}".format(first_intersecting_element_id, str(e)))

if intersecting_element is None:
    forms.alert("First intersecting element not found in the selected intersecting document. Exiting script.", exitscript=True)

# Collect instance parameters
for param in intersecting_element.Parameters:
    if param.Definition.Name and param.HasValue:
        intersecting_param_names.add(param.Definition.Name)

# If the element is a FamilyInstance, collect type parameters
if isinstance(intersecting_element, DB.FamilyInstance):
    family_symbol = intersecting_element.Symbol
    for param in family_symbol.Parameters:
        if param.Definition.Name and param.HasValue:
            intersecting_param_names.add(param.Definition.Name)

# Convert the set to a sorted list
intersecting_param_names = sorted(intersecting_param_names)

# Step 11: Allow the user to select parameters from intersecting elements to retrieve
selected_intersecting_params = forms.SelectFromList.show(
    intersecting_param_names,
    title="Select parameters to retrieve from Intersecting Elements (multiple selection allowed)",
    multiselect=True
)

if not selected_intersecting_params:
    forms.alert("No parameters selected from intersecting elements. Exiting script.", exitscript=True)

# Step 11a: Ask the user if they want to include parameter names as prefixes
include_param_names_options = ['Yes', 'No']
include_param_names_choice = forms.CommandSwitchWindow.show(
    include_param_names_options,
    message="Include parameter names as prefixes in the output?",
    title="Parameter Name Prefix"
)

if include_param_names_choice == 'Yes':
    include_param_names = True
else:
    include_param_names = False

# Step 12: Collect data from intersecting elements and update primary elements
with revit.Transaction("Update Primary Elements with Intersecting Element Data"):
    for primary_element in primary_elements:
        primary_element_id_str = str(primary_element.Id.IntegerValue)
        intersecting_element_ids = grouped_data[primary_element_id_str]
        # Initialize a dictionary to store parameter values and counts
        param_values_counts = {param_name: {} for param_name in selected_intersecting_params}

        for intersecting_element_id_str in intersecting_element_ids:
            intersecting_element = None
            # Try to get the intersecting element from the selected intersecting document
            try:
                element_id = DB.ElementId(int(intersecting_element_id_str))
                intersecting_element = intersecting_doc.GetElement(element_id)
            except Exception as e:
                output.print_md("Error getting intersecting element ID {}: {}".format(intersecting_element_id_str, str(e)))

            if intersecting_element is None:
                # If still None, skip this intersecting element
                continue

            # Retrieve the selected parameters from the intersecting element
            for param_name in selected_intersecting_params:
                # Try instance parameter
                param = intersecting_element.LookupParameter(param_name)
                param_value = None
                if param and param.HasValue:
                    param_value = param.AsString() or param.AsValueString() or ""
                else:
                    # If not found on instance, check type parameters for FamilyInstances
                    if isinstance(intersecting_element, DB.FamilyInstance):
                        family_symbol = intersecting_element.Symbol
                        param = family_symbol.LookupParameter(param_name)
                        if param and param.HasValue:
                            param_value = param.AsString() or param.AsValueString() or ""

                if param_value is not None:
                    # Aggregate the values and counts
                    if param_value in param_values_counts[param_name]:
                        param_values_counts[param_name][param_value] += 1
                    else:
                        param_values_counts[param_name][param_value] = 1
                else:
                    # Handle cases where the parameter value is not found
                    if "N/A" in param_values_counts[param_name]:
                        param_values_counts[param_name]["N/A"] += 1
                    else:
                        param_values_counts[param_name]["N/A"] = 1

        # Prepare the final data to set in the primary element's parameter
        combined_data_parts = []
        for param_name in selected_intersecting_params:
            values_counts = param_values_counts[param_name]
            # Format the values with counts
            formatted_values = []
            for value, count in values_counts.items():
                if count > 1:
                    formatted_value = "{} ({})".format(value, count)
                else:
                    formatted_value = value
                formatted_values.append(formatted_value)

            # Remove duplicates
            unique_formatted_values = list(set(formatted_values))

            # Combine values into a single string
            if include_param_names:
                data_part = "{0}: {1}".format(param_name, ", ".join(unique_formatted_values))
            else:
                data_part = ", ".join(unique_formatted_values)

            combined_data_parts.append(data_part)

        # Combine all data parts into the final value
        final_value = "\n".join(combined_data_parts)

        # Set the final value to the primary element's parameter
        primary_param = primary_element.LookupParameter(selected_primary_param)
        if primary_param and not primary_param.IsReadOnly:
            try:
                primary_param.Set(final_value)
            except Exception as e:
                output.print_md("Error setting parameter '{}' on element {}: {}".format(selected_primary_param, primary_element.Id.IntegerValue, str(e)))
        else:
            output.print_md("Parameter '{}' not found or read-only on primary element.".format(selected_primary_param))

# Step 13: Output success message
output.print_md("Successfully updated primary elements with selected data of intersecting elements.")
