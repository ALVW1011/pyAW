from pyrevit import revit, DB, forms
import csv
import os

# Current Revit Document
doc = revit.doc

def is_file_locked(filepath):
    """Check if a file is locked by another process."""
    if not os.path.exists(filepath):
        return False
    try:
        with open(filepath, 'a'):
            pass
        return False
    except IOError:
        return True

def get_shared_parameter_names():
    """Retrieve all shared parameter names from the project."""
    shared_param_names = set()
    binding_map = doc.ParameterBindings
    iterator = binding_map.ForwardIterator()
    iterator.Reset()
    while iterator.MoveNext():
        definition = iterator.Key
        if definition.ParameterType != DB.ParameterType.Invalid and definition.ParameterGroup != DB.BuiltInParameterGroup.INVALID:
            shared_param_names.add(definition.Name)
    return sorted(shared_param_names)

def export_schedule_to_csv():
    """Export a selected schedule to a CSV file, ensuring headers are not duplicated."""
    # Collect all non-template schedules in the project
    schedules = [view for view in DB.FilteredElementCollector(doc)
                 .OfClass(DB.ViewSchedule) if not view.IsTemplate]

    if not schedules:
        forms.alert("No schedules found in the project.", title="Export Aborted")
        return

    # Create a mapping from schedule names to schedule objects
    schedule_dict = {schedule.Name: schedule for schedule in schedules}
    schedule_names = sorted(schedule_dict.keys())

    # Prompt user to select a schedule
    selected_schedule_name = forms.SelectFromList.show(
        schedule_names,
        title="Select Schedule to Export",
        button_name="Export",
        multiselect=False
    )

    if not selected_schedule_name:
        forms.alert("No schedule selected. Aborting export.")
        return

    schedule = schedule_dict[selected_schedule_name]

    # Collect all visible fields
    definition = schedule.Definition
    fields = []
    headers = []
    for i in range(definition.GetFieldCount()):
        field = definition.GetField(i)
        if not field.IsHidden:
            fields.append(field)
            headers.append(field.GetName())

    if len(fields) == 0:
        forms.alert("The selected schedule has no visible fields. Aborting export.")
        return

    # Prompt user to save CSV file
    csv_file_path = forms.save_file(file_ext='csv')
    if not csv_file_path:
        forms.alert("No file path selected. Aborting export.")
        return

    # Check if the file is locked and prompt user to close it
    while is_file_locked(csv_file_path):
        retry = forms.CommandSwitchWindow.show(
            ["Retry", "Abort"],
            message="The file is currently in use. Please close the file and choose an action."
        )
        if retry == "Abort" or not retry:
            forms.alert("Export aborted by user.", title="Export Aborted")
            return  # User chose to abort the operation

    # Get schedule data
    table_data = schedule.GetTableData()
    body_section = table_data.GetSectionData(DB.SectionType.Body)

    # Open CSV file for writing
    with open(csv_file_path, 'wb') as file:
        writer = csv.writer(file, lineterminator='\n')
        # Write headers in a single row
        writer.writerow([h.encode('utf-8') for h in headers])

        # Determine the start and end rows
        start_row = body_section.FirstRowNumber
        end_row = body_section.LastRowNumber

        # Check if the first row contains headers and skip it if necessary
        first_row_values = [schedule.GetCellText(DB.SectionType.Body, start_row, col).strip() for col in range(len(fields))]
        header_values = [h.strip() for h in headers]
        if first_row_values == header_values:
            start_row += 1  # Skip the first row as it contains headers

        # Iterate over each row, skipping group headers
        for row in range(start_row, end_row + 1):
            try:
                row_data = []

                # Check if the row corresponds to an element (i.e., the first cell is not empty)
                first_cell_text = schedule.GetCellText(DB.SectionType.Body, row, 0).strip()
                if not first_cell_text:
                    continue  # Skip group headers or empty rows

                # Retrieve field values
                for col_index in range(len(fields)):
                    cell_text = schedule.GetCellText(DB.SectionType.Body, row, col_index).strip()
                    row_data.append(cell_text.encode('utf-8'))

                # Write the row to CSV
                writer.writerow(row_data)

            except Exception as e:
                # Handle errors gracefully
                user_choice = forms.CommandSwitchWindow.show(
                    ["Continue", "Abort"],
                    message="Error processing row {0}: {1}\nDo you want to continue or abort?".format(row, str(e))
                )
                if user_choice == "Abort" or not user_choice:
                    forms.alert("Export aborted by user.", title="Export Aborted")
                    return  # Exit the loop and abort export
                else:
                    continue  # Proceed to next row

    forms.toast(
        "Schedule '{0}' exported successfully to: {1}".format(schedule.Name, csv_file_path),
        title="Export Complete"
    )

def import_csv_to_revit():
    """Import data from a CSV file to update Revit elements, skipping unchanged values."""
    # Prompt user to select CSV file
    csv_file_path = forms.pick_file(file_ext='csv')
    if not csv_file_path:
        forms.alert("No file selected. Aborting import.")
        return

    # Get available shared parameter names
    shared_param_names = get_shared_parameter_names()

    if not shared_param_names:
        forms.alert("No shared parameters found in the project.", title="Import Aborted")
        return

    # Prompt user to select the unique identifier parameter
    unique_id_param_name = forms.SelectFromList.show(
        shared_param_names,
        title="Select Unique Identifier Parameter",
        button_name="Select",
        multiselect=False
    )

    if not unique_id_param_name:
        forms.alert("No unique identifier parameter selected. Aborting import.")
        return

    # Get the parameter definition
    param_definition = None
    binding_map = doc.ParameterBindings
    iterator = binding_map.ForwardIterator()
    iterator.Reset()
    while iterator.MoveNext():
        definition = iterator.Key
        if definition.Name == unique_id_param_name:
            param_definition = definition
            break

    if not param_definition:
        forms.alert("Parameter '{0}' not found in the project.".format(unique_id_param_name))
        return

    # Read CSV data
    with open(csv_file_path, 'r') as file:
        reader = csv.DictReader(file)
        t = DB.Transaction(doc, "Import CSV Data")
        t.Start()
        try:
            user_wants_to_abort = False  # Flag to determine if the user wants to abort

            # Prompt user about updating type parameters
            update_type_params = forms.CommandSwitchWindow.show(
                ["Yes", "No"],
                message="Do you want to update Type Parameters during import?"
            ) == "Yes"

            for row in reader:
                if user_wants_to_abort:
                    break  # Exit the loop if user chose to abort
                try:
                    # Retrieve unique identifier from the row
                    unique_id_value = row.get(unique_id_param_name, "").strip()
                    if not unique_id_value:
                        continue  # Skip rows with invalid unique identifier

                    # Prepare the filter based on the parameter storage type
                    provider = DB.ParameterValueProvider(param_definition.Id)
                    param_storage_type = param_definition.ParameterType
                    if param_storage_type == DB.ParameterType.Text:
                        # Use string comparison
                        evaluator = DB.FilterStringEquals()
                        rule = DB.FilterStringRule(provider, evaluator, unique_id_value, False)
                        element_filter = DB.ElementParameterFilter(rule)
                    elif param_storage_type == DB.ParameterType.Integer:
                        # Convert unique_id_value to integer
                        try:
                            unique_id_int = int(unique_id_value)
                        except ValueError:
                            # Invalid integer value
                            continue  # Skip this row
                        evaluator = DB.FilterNumericEquals()
                        rule = DB.FilterIntegerRule(provider, evaluator, unique_id_int)
                        element_filter = DB.ElementParameterFilter(rule)
                    else:
                        # Unsupported parameter type
                        forms.alert("Unsupported unique identifier parameter type.")
                        user_wants_to_abort = True
                        break

                    # Find the element with this unique identifier
                    collector = DB.FilteredElementCollector(doc).WhereElementIsNotElementType()
                    elements = collector.WherePasses(element_filter).ToElements()
                    if elements:
                        element = elements[0]
                    else:
                        # Handle element not found
                        user_choice = forms.CommandSwitchWindow.show(
                            ["Continue", "Abort"],
                            message="Element with '{0}' = '{1}' not found.\nDo you want to continue or abort?".format(unique_id_param_name, unique_id_value)
                        )
                        if user_choice == "Abort" or not user_choice:
                            user_wants_to_abort = True
                            break
                        else:
                            continue  # Skip to next row

                    # Iterate through each key-value pair in the row
                    for key, value in row.items():
                        if key == unique_id_param_name:
                            continue  # Skip unique identifier

                        param = element.LookupParameter(key)
                        if not param and hasattr(element, "Symbol") and element.Symbol:
                            param = element.Symbol.LookupParameter(key)
                        if not param:
                            continue  # Parameter not found, skip

                        if param.IsReadOnly:
                            continue  # Skip read-only parameters

                        # Determine if the parameter is a type parameter
                        is_type_param = param.Element.Id != element.Id
                        if is_type_param and not update_type_params:
                            continue  # Skip updating this parameter

                        # Get current parameter value
                        current_value = None
                        if param.StorageType == DB.StorageType.String:
                            current_value = param.AsString() or ""
                        elif param.StorageType == DB.StorageType.Double:
                            # Get parameter's definition
                            param_def = param.Definition
                            # Get parameter's UnitType
                            param_unit_type = param_def.UnitType
                            # Get the format options for this UnitType from project units
                            format_options = doc.GetUnits().GetFormatOptions(param_unit_type)
                            display_unit = format_options.DisplayUnits
                            # Convert internal value to display units for comparison
                            internal_value = param.AsDouble()
                            current_value = DB.UnitUtils.ConvertFromInternalUnits(internal_value, display_unit)
                        elif param.StorageType == DB.StorageType.Integer:
                            current_value = param.AsInteger()
                        elif param.StorageType == DB.StorageType.ElementId:
                            current_value = param.AsElementId().IntegerValue

                        # Compare the current value with the new value
                        new_value = value.strip()
                        if param.StorageType == DB.StorageType.String:
                            if current_value == new_value:
                                continue  # Skip if value hasn't changed
                            param.Set(new_value)
                        elif param.StorageType == DB.StorageType.Double:
                            try:
                                # Parse the value from the CSV (assumed to be in display units)
                                display_value = float(new_value)
                                # Convert to internal units (feet)
                                internal_value = DB.UnitUtils.ConvertToInternalUnits(display_value, display_unit)
                                if abs(param.AsDouble() - internal_value) < 1e-6:
                                    continue  # Skip if value hasn't changed
                                param.Set(internal_value)
                            except Exception as e:
                                # Handle exceptions if needed
                                pass
                        elif param.StorageType == DB.StorageType.Integer:
                            try:
                                new_int_value = int(new_value)
                                if current_value == new_int_value:
                                    continue  # Skip if value hasn't changed
                                param.Set(new_int_value)
                            except:
                                pass  # Invalid integer, skip
                        elif param.StorageType == DB.StorageType.ElementId:
                            try:
                                new_elem_id = DB.ElementId(int(new_value))
                                if current_value == new_elem_id.IntegerValue:
                                    continue  # Skip if value hasn't changed
                                param.Set(new_elem_id)
                            except:
                                pass  # Invalid ElementId, skip

                except Exception as e:
                    # Handle errors gracefully
                    user_choice = forms.CommandSwitchWindow.show(
                        ["Continue", "Abort"],
                        message="Error updating element with '{0}' = '{1}': {2}\nDo you want to continue or abort?".format(
                            unique_id_param_name, unique_id_value, str(e))
                    )
                    if user_choice == "Abort" or not user_choice:
                        forms.alert("Import aborted by user.", title="Import Aborted")
                        t.RollBack()
                        return  # Exit the loop and abort import
                    else:
                        continue  # Proceed to next row
            # Commit the transaction if all updates succeed
            if not user_wants_to_abort:
                t.Commit()
                forms.toast("Data imported successfully.", title="Import Complete")
            else:
                t.RollBack()
                forms.alert("Import aborted by user.", title="Import Aborted")

        except Exception as e:
            # Rollback transaction in case of unexpected errors
            t.RollBack()
            forms.alert("Error during import: {0}".format(str(e)), title="Import Aborted")
            return

def schedule_export_import():
    """Provide a menu to choose between exporting or importing schedule data."""
    options = {
        "Export Schedule to CSV": export_schedule_to_csv,
        "Import CSV to Revit": import_csv_to_revit
    }

    selected_option = forms.CommandSwitchWindow.show(
        options.keys(),
        message="Choose an action:",
        title="Schedule Export/Import",
        button_name="Run"
    )

    if selected_option:
        options[selected_option]()
    else:
        forms.alert("No action selected. Operation cancelled.", title="Operation Cancelled")

# Run the combined export/import function
schedule_export_import()
