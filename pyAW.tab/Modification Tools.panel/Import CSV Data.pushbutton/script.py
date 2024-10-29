from pyrevit import forms, script, revit, DB
import os
import csv
from collections import Counter

# Initialize the output log for showing results
output = script.get_output()
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# Helper Functions

def format_parameter_info(param, is_type_param):
    """
    Formats the parameter name with flags indicating if it's read-only, built-in, shared, and type or instance.
    """
    flags = []
    if param.IsReadOnly:
        flags.append("Read-Only")
    if param.Definition.BuiltInParameter != DB.BuiltInParameter.INVALID:
        flags.append("Built-In")
    else:
        flags.append("Shared")
    flags.append("Type" if is_type_param else "Instance")
    flags_str = ", ".join(flags)
    return "{0} [{1}]".format(param.Definition.Name, flags_str)

def get_relevant_parameters(element, include_read_only=False):
    """
    Retrieves relevant parameters from an element, including both instance and type parameters.
    """
    params = []

    # Collect instance parameters
    for param in element.Parameters:
        if param.StorageType != DB.StorageType.None and (include_read_only or not param.IsReadOnly):
            formatted_name = format_parameter_info(param, is_type_param=False)
            params.append((formatted_name, param.Definition.Name))

    # Collect type parameters
    type_id = element.GetTypeId()
    if type_id != DB.ElementId.InvalidElementId:
        type_element = element.Document.GetElement(type_id)
        if type_element:
            for param in type_element.Parameters:
                if param.StorageType != DB.StorageType.None and (include_read_only or not param.IsReadOnly):
                    formatted_name = format_parameter_info(param, is_type_param=True)
                    params.append((formatted_name, param.Definition.Name))

    return params

def are_parameter_types_compatible(source_param, target_param):
    """
    Determines if the source and target parameters have compatible types.
    """
    if source_param.StorageType == DB.StorageType.ElementId and target_param.StorageType == DB.StorageType.String:
        return True

    compatible_pairs = {
        DB.StorageType.String: [DB.StorageType.String],
        DB.StorageType.Integer: [DB.StorageType.Integer],
        DB.StorageType.Double: [DB.StorageType.Double],
        DB.StorageType.ElementId: [DB.StorageType.ElementId]
    }

    source_type = source_param.StorageType
    target_type = target_param.StorageType

    return target_type in compatible_pairs.get(source_type, [])

def convert_parameter_value(source_param, target_param, value):
    """
    Converts the parameter value from source type to target type.
    """
    source_type = source_param.StorageType
    target_type = target_param.StorageType

    if source_type == DB.StorageType.ElementId and target_type == DB.StorageType.String:
        return (value, True)
    elif source_type == DB.StorageType.Double and target_type == DB.StorageType.String:
        return (str(value), True)
    elif source_type == DB.StorageType.String and target_type == DB.StorageType.Double:
        try:
            numeric_value = float(value)
            return (numeric_value, True)
        except ValueError:
            return (None, False)
    elif source_type == DB.StorageType.Integer and target_type == DB.StorageType.String:
        return (str(value), True)
    elif source_type == DB.StorageType.String and target_type == DB.StorageType.Integer:
        try:
            numeric_value = int(value)
            return (numeric_value, True)
        except ValueError:
            return (None, False)
    else:
        return (None, False)

def get_parameter_value(param, doc):
    """
    Retrieves the value of a parameter based on its StorageType.
    """
    if param.HasValue:
        value_str = param.AsValueString()
        if value_str:
            return value_str
        else:
            if param.StorageType == DB.StorageType.String:
                return param.AsString()
            elif param.StorageType == DB.StorageType.Integer:
                return param.AsInteger()
            elif param.StorageType == DB.StorageType.Double:
                return param.AsDouble()
            elif param.StorageType == DB.StorageType.ElementId:
                element_id = param.AsElementId()
                if element_id != DB.ElementId.InvalidElementId and element_id != DB.ElementId(-1):
                    element = doc.GetElement(element_id)
                    if element:
                        return element.Name
                    else:
                        category = DB.Category.GetCategory(doc, element_id)
                        if category:
                            return category.Name
                        else:
                            return "ElementId: {}".format(element_id.IntegerValue)
                else:
                    return "None"
            else:
                return None
    else:
        return None

# Built-In Parameters
STRUCTURAL_MATERIAL_PARAM = DB.BuiltInParameter.STRUCTURAL_MATERIAL_PARAM
TYPE_MARK_PARAM = DB.BuiltInParameter.ALL_MODEL_TYPE_MARK

# Step 1: Let the user select a CSV file
file_path = forms.pick_file(file_ext='csv', title="Select a CSV Clash Report File")

if not file_path:
    forms.alert("No file selected. Exiting script.", exitscript=True)

# Step 2: Read the CSV file to extract the headers
try:
    with open(file_path, 'r') as f:
        sample = f.read(1024)
        f.seek(0)
        dialect = csv.Sniffer().sniff(sample)
        reader = csv.reader(f, dialect)
        headers = next(reader)
except Exception as e:
    forms.alert("Error reading CSV file: {0}".format(str(e)), exitscript=True)

# Step 3: Let the user select headers for Primary and Intersecting Elements
primary_header = forms.SelectFromList.show(headers, title="Select Primary Element ID Column")
if not primary_header:
    forms.alert("No Primary Element ID column selected. Exiting script.", exitscript=True)

intersecting_header = forms.SelectFromList.show(headers, title="Select Intersecting Element ID Column")
if not intersecting_header:
    forms.alert("No Intersecting Element ID column selected. Exiting script.", exitscript=True)

# Step 3a: Let the user select the document (host or linked) for each column
documents = []
document_names = []

# Add the host document
documents.append(doc)
document_names.append("Host Document ({0})".format(doc.Title))

# Collect linked documents
link_instances = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance)
for link_instance in link_instances:
    link_doc = link_instance.GetLinkDocument()
    if link_doc:
        documents.append(link_doc)
        link_name = link_instance.Name
        document_names.append("Linked Document ({0})".format(link_name))

# Output available documents
output.print_md("### Available Documents:")
for idx, doc_name in enumerate(document_names):
    output.print_md("  {0}: {1}".format(idx, doc_name))

# Select documents for primary and intersecting elements
primary_doc_choice = forms.SelectFromList.show(document_names, title="Select Document for Primary Elements")
if not primary_doc_choice:
    forms.alert("No document selected for Primary Elements. Exiting script.", exitscript=True)
primary_doc_index = document_names.index(primary_doc_choice)
primary_doc = documents[primary_doc_index]
output.print_md("**Selected primary document:** {0}".format(primary_doc_choice))

intersecting_doc_choice = forms.SelectFromList.show(document_names, title="Select Document for Intersecting Elements")
if not intersecting_doc_choice:
    forms.alert("No document selected for Intersecting Elements. Exiting script.", exitscript=True)
intersecting_doc_index = document_names.index(intersecting_doc_choice)
intersecting_doc = documents[intersecting_doc_index]
output.print_md("**Selected intersecting document:** {0}".format(intersecting_doc_choice))

# Step 4: Process the data from the CSV
data = []
try:
    with open(file_path, 'r') as f:
        reader = csv.DictReader(f, dialect=dialect)
        for row in reader:
            if row[primary_header] and row[intersecting_header]:
                data.append((row[primary_header], row[intersecting_header]))
except Exception as e:
    forms.alert("Error processing CSV data: {0}".format(str(e)), exitscript=True)

# Output total number of rows processed
output.print_md("**Total rows processed from CSV:** {0}".format(len(data)))

# Step 5: Group by primary elements and combine intersecting elements
grouped_data = {}
for primary_id_str, intersecting_id_str in data:
    if primary_id_str not in grouped_data:
        grouped_data[primary_id_str] = set()
    grouped_data[primary_id_str].add(intersecting_id_str)

# Step 6: Detect the primary elements in the selected document
primary_element_ids = list(grouped_data.keys())
primary_elements = []
for element_id_str in primary_element_ids:
    try:
        element_id = DB.ElementId(int(element_id_str))
        element = primary_doc.GetElement(element_id)
        if element:
            primary_elements.append(element)
        else:
            output.print_md("Element ID **{0}** not found in the selected primary document.".format(element_id_str))
    except Exception as e:
        output.print_md("Error processing Element ID **{0}**: {1}".format(element_id_str, str(e)))

# Output total number of primary elements found
output.print_md("**Total primary elements found:** {0}".format(len(primary_elements)))

# If no valid primary elements were detected, exit the script
if not primary_elements:
    forms.alert("No valid primary elements found in the selected primary document. Exiting script.", exitscript=True)

# Step 7: Show intermediate output for primary elements and their intersecting elements
processed_data = []
for primary_element in primary_elements:
    primary_element_id = str(primary_element.Id.IntegerValue)
    intersecting_element_ids = list(grouped_data[primary_element_id])
    processed_data.append([primary_element_id, ", ".join(intersecting_element_ids)])

# Output the intermediate data
output.print_md("### Intermediate Output: Primary and Intersecting Elements")

try:
    output.print_table(processed_data, title="Primary Elements and Intersecting Elements", columns=["Primary Element ID", "Intersecting Element IDs"])
except Exception as e:
    output.print_md("Error displaying table: {0}".format(str(e)))
    forms.alert("Failed to display table. Exiting script.", exitscript=True)

# Step 9: Allow the user to select the parameter in primary elements to update
available_primary_params = get_relevant_parameters(primary_elements[0], include_read_only=False)
if not available_primary_params:
    forms.alert("No writable parameters found in primary elements. Exiting script.", exitscript=True)

# Extract formatted names for display
formatted_primary_params = [param[0] for param in available_primary_params]

# User selects from formatted names
selected_formatted_param = forms.SelectFromList.show(
    formatted_primary_params,
    title="Select parameter to update in Primary Elements"
)

if not selected_formatted_param:
    forms.alert("No parameter selected from primary elements. Exiting script.", exitscript=True)

# Map back to the actual parameter name
selected_primary_param = None
for formatted_name, actual_name in available_primary_params:
    if formatted_name == selected_formatted_param:
        selected_primary_param = actual_name
        break

if not selected_primary_param:
    forms.alert("Selected parameter could not be mapped. Exiting script.", exitscript=True)

# Step 10: Collect parameters from the intersecting elements
intersecting_params = {}

# Iterate through intersecting elements to collect parameters
for primary_element in primary_elements:
    primary_element_id_str = str(primary_element.Id.IntegerValue)
    intersecting_element_ids = grouped_data[primary_element_id_str]
    for intersecting_element_id_str in intersecting_element_ids:
        try:
            element_id = DB.ElementId(int(intersecting_element_id_str))
            intersecting_element = intersecting_doc.GetElement(element_id)
        except Exception as e:
            output.print_md("Error getting intersecting element ID **{0}**: {1}".format(intersecting_element_id_str, str(e)))
            continue

        if intersecting_element is None:
            continue

        # Retrieve relevant parameters, including read-only ones
        relevant_params = get_relevant_parameters(intersecting_element, include_read_only=True)
        for formatted_name, actual_name in relevant_params:
            intersecting_params[formatted_name] = actual_name

# Convert the dictionary keys to a sorted list for display
formatted_intersecting_params = sorted(intersecting_params.keys())

# User selects from formatted names
selected_formatted_intersecting_params = forms.SelectFromList.show(
    formatted_intersecting_params,
    title="Select parameters to retrieve from Intersecting Elements (multiple selection allowed)",
    multiselect=True
)

if not selected_formatted_intersecting_params:
    forms.alert("No parameters selected from intersecting elements. Exiting script.", exitscript=True)

# Map back to actual parameter names
selected_intersecting_params = [intersecting_params[formatted_name] for formatted_name in selected_formatted_intersecting_params]

# Step 11a: Ask the user if they want to include parameter names as prefixes
include_param_names_choice = forms.CommandSwitchWindow.show(
    ['Yes', 'No'],
    message="Include parameter names as prefixes in the output?",
    title="Parameter Name Prefix"
)

include_param_names = include_param_names_choice == 'Yes'

# Step 11b: Aggregate all unique parameter values across all primary elements
unique_param_values = {param_name: set() for param_name in selected_intersecting_params}

# Collect unique parameter values
for primary_element in primary_elements:
    primary_element_id_str = str(primary_element.Id.IntegerValue)
    intersecting_element_ids = grouped_data[primary_element_id_str]
    for intersecting_element_id_str in intersecting_element_ids:
        try:
            element_id = DB.ElementId(int(intersecting_element_id_str))
            intersecting_element = intersecting_doc.GetElement(element_id)
        except Exception as e:
            output.print_md("Error getting intersecting element ID **{0}**: {1}".format(intersecting_element_id_str, str(e)))
            continue

        if intersecting_element is None:
            continue

        for param_name in selected_intersecting_params:
            source_value = None

            # Check if the selected parameter is "Structural Material"
            if param_name == "Structural Material":
                type_id = intersecting_element.GetTypeId()
                if type_id != DB.ElementId.InvalidElementId:
                    type_element = intersecting_element.Document.GetElement(type_id)
                    if type_element:
                        source_param = type_element.get_Parameter(STRUCTURAL_MATERIAL_PARAM)
                        if source_param and source_param.HasValue:
                            source_value = get_parameter_value(source_param, intersecting_doc)
            elif param_name == "Type Mark":
                type_id = intersecting_element.GetTypeId()
                if type_id != DB.ElementId.InvalidElementId:
                    type_element = intersecting_element.Document.GetElement(type_id)
                    if type_element:
                        source_param = type_element.get_Parameter(TYPE_MARK_PARAM)
                        if source_param and source_param.HasValue:
                            source_value = get_parameter_value(source_param, intersecting_doc)
            else:
                source_param = intersecting_element.LookupParameter(param_name)

                if source_param is None or not source_param.HasValue:
                    type_id = intersecting_element.GetTypeId()
                    if type_id != DB.ElementId.InvalidElementId:
                        type_element = intersecting_element.Document.GetElement(type_id)
                        if type_element:
                            source_param = type_element.LookupParameter(param_name)
                            if source_param and source_param.HasValue:
                                source_value = get_parameter_value(source_param, intersecting_doc)
                else:
                    source_value = get_parameter_value(source_param, intersecting_doc)

            source_value_str = "N/A" if source_value is None else str(source_value)
            unique_param_values[param_name].add(source_value_str)

# Step 11c: Let the user select which unique values to include per parameter
selected_values_per_param = {}
for param_name, values in unique_param_values.items():
    sorted_values = sorted(values)
    selected_values = forms.SelectFromList.show(
        sorted_values,
        title="Select values to include for parameter '{0}'".format(param_name),
        multiselect=True
    )
    if selected_values:
        selected_values_per_param[param_name] = set(selected_values)
    else:
        selected_values_per_param[param_name] = set()

# Step 11d: Ask the user to input a delimiter if more than one intersecting parameter is chosen
delimiter = "\n"
if len(selected_intersecting_params) > 1:
    delimiter_input = forms.ask_for_string(
        prompt="Enter a delimiter to separate different parameters' values within each group (e.g., ';', '|', ','):",
        default=";",
        title="Input Delimiter"
    )
    if delimiter_input is not None:
        delimiter = delimiter_input
    else:
        delimiter = ";"

# Step 12: Collect data from intersecting elements and update primary elements
with revit.Transaction("Update Primary Elements with Intersecting Element Data"):
    for primary_element in primary_elements:
        primary_element_id_str = str(primary_element.Id.IntegerValue)
        intersecting_element_ids = grouped_data[primary_element_id_str]
        # Initialize a list to store grouped parameter values
        grouped_param_values = []

        # Initialize variables to keep track of user decisions for batch operations
        apply_to_all = False
        user_decision = None  # Stores 'Convert and Transfer', 'Cancel Transfer', 'Cancel Script'

        for intersecting_element_id_str in intersecting_element_ids:
            try:
                element_id = DB.ElementId(int(intersecting_element_id_str))
                intersecting_element = intersecting_doc.GetElement(element_id)
            except Exception as e:
                output.print_md("Error getting intersecting element ID **{0}**: {1}".format(intersecting_element_id_str, str(e)))
                continue

            if intersecting_element is None:
                continue

            # Collect parameter values for this intersecting element
            param_values = []
            for param_name in selected_intersecting_params:
                source_value = None

                # Check if the selected parameter is "Structural Material"
                if param_name == "Structural Material":
                    type_id = intersecting_element.GetTypeId()
                    if type_id != DB.ElementId.InvalidElementId:
                        type_element = intersecting_element.Document.GetElement(type_id)
                        if type_element:
                            source_param = type_element.get_Parameter(STRUCTURAL_MATERIAL_PARAM)
                            if source_param and source_param.HasValue:
                                source_value = get_parameter_value(source_param, intersecting_doc)
                elif param_name == "Type Mark":
                    type_id = intersecting_element.GetTypeId()
                    if type_id != DB.ElementId.InvalidElementId:
                        type_element = intersecting_element.Document.GetElement(type_id)
                        if type_element:
                            source_param = type_element.get_Parameter(TYPE_MARK_PARAM)
                            if source_param and source_param.HasValue:
                                source_value = get_parameter_value(source_param, intersecting_doc)
                else:
                    source_param = intersecting_element.LookupParameter(param_name)

                    if source_param is None or not source_param.HasValue:
                        type_id = intersecting_element.GetTypeId()
                        if type_id != DB.ElementId.InvalidElementId:
                            type_element = intersecting_element.Document.GetElement(type_id)
                            if type_element:
                                source_param = type_element.LookupParameter(param_name)
                                if source_param and source_param.HasValue:
                                    source_value = get_parameter_value(source_param, intersecting_doc)
                    else:
                        source_value = get_parameter_value(source_param, intersecting_doc)

                source_value_str = "N/A" if source_value is None else str(source_value)

                # If the source value is not selected by the user, skip
                if source_value_str not in selected_values_per_param.get(param_name, set()):
                    continue

                # Lookup target parameter in the primary element
                target_param = primary_element.LookupParameter(selected_primary_param)
                if target_param is None:
                    output.print_md("Parameter **'{0}'** not found on primary element ID **{1}**.".format(selected_primary_param, primary_element.Id.IntegerValue))
                    continue

                # Check type compatibility
                if not are_parameter_types_compatible(source_param, target_param):
                    if apply_to_all and user_decision:
                        user_choice = user_decision
                    else:
                        message = "Parameter **'{0}'** has incompatible types.\nSource Type: **{1}**\nTarget Type: **{2}**\nChoose an action:".format(
                            param_name,
                            source_param.StorageType,
                            target_param.StorageType
                        )
                        user_choice = forms.CommandSwitchWindow.show(
                            ['Convert and Transfer', 'Convert and Transfer All', 'Cancel Transfer', 'Cancel Transfer All', 'Cancel Script'],
                            message=message,
                            title="Type Mismatch Detected"
                        )

                        if user_choice in ['Convert and Transfer All', 'Cancel Transfer All']:
                            apply_to_all = True
                            user_decision = user_choice

                    if user_choice == 'Convert and Transfer' or user_choice == 'Convert and Transfer All':
                        converted_value, success = convert_parameter_value(source_param, target_param, source_value)
                        if success:
                            source_value = converted_value
                        else:
                            output.print_md("**Failed to convert value '{0}' for parameter '{1}'. Skipping transfer.**".format(source_value, param_name))
                            continue
                    elif user_choice == 'Cancel Transfer' or user_choice == 'Cancel Transfer All':
                        output.print_md("**Transfer of parameter '{0}' cancelled by user.**".format(param_name))
                        continue
                    elif user_choice == 'Cancel Script':
                        forms.alert("Script execution cancelled by user.", exitscript=True)
                    else:
                        continue

                param_values.append(str(source_value))

            # Only proceed if param_values has all selected parameters
            if len(param_values) == len(selected_intersecting_params):
                # Join parameter values with the delimiter
                grouped_value = delimiter.join(param_values)
                grouped_param_values.append(grouped_value)

        # Count duplicates and format values with counts
        value_counts = Counter(grouped_param_values)
        formatted_values = []
        for value, count in value_counts.items():
            if count > 1:
                formatted_values.append("{0} ({1})".format(value, count))
            else:
                formatted_values.append(value)

        # Optionally include parameter names as prefixes
        if include_param_names:
            new_formatted_values = []
            for item in formatted_values:
                # Extract count if present
                if " (" in item and item.endswith(")"):
                    value_part, count_part = item.rsplit(" (", 1)
                    count_part = "(" + count_part
                else:
                    value_part = item
                    count_part = ""

                values = value_part.split(delimiter)
                if len(values) != len(selected_intersecting_params):
                    continue
                paired_values = ["{0}: {1}".format(param_name, val) for param_name, val in zip(selected_intersecting_params, values)]
                new_value = delimiter.join(paired_values) + count_part
                new_formatted_values.append(new_value)
            final_value = ", ".join(new_formatted_values)
        else:
            final_value = ", ".join(formatted_values)

        # Set the final value to the primary element's parameter
        primary_param = primary_element.LookupParameter(selected_primary_param)
        if primary_param and not primary_param.IsReadOnly:
            try:
                primary_param.Set(final_value)
            except Exception as e:
                output.print_md("**Error setting parameter '{0}' on element {1}: {2}**".format(selected_primary_param, primary_element.Id.IntegerValue, str(e)))
        else:
            output.print_md("**Parameter '{0}' not found or is read-only on primary element.**".format(selected_primary_param))

# Step 13: Output success message
output.print_md("**Successfully updated primary elements with selected data from intersecting elements.**")
