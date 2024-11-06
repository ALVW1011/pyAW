# -*- coding: utf-8 -*-

__title__ = 'Dimension Selected Elements'
__author__ = 'Your Name'

from pyrevit import forms
from pyrevit import script
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ISelectionFilter
from System.Collections.Generic import List

# Get the active Revit application and documents
uiapp = __revit__
uidoc = uiapp.ActiveUIDocument
doc = uidoc.Document

import sys

# Initialize output for debugging
output = script.get_output()

# Ask user how to select views
view_selection_options = ['Use Current View', 'Select Views from List', 'Select Viewports on Sheet']
selected_view_selection = forms.SelectFromList.show(
    view_selection_options,
    multiselect=False,
    title='How would you like to select the views for dimension placement?'
)

if not selected_view_selection:
    forms.alert('No view selection method chosen.')
    script.exit()

selected_views = []

if selected_view_selection == 'Use Current View':
    selected_views = [uidoc.ActiveView]
elif selected_view_selection == 'Select Views from List':
    # Collect all views in the project
    all_views = FilteredElementCollector(doc)\
        .OfClass(View)\
        .WhereElementIsNotElementType()\
        .ToElements()

    # Filter views to include only specific view types
    valid_view_types = [
        ViewType.FloorPlan,
        ViewType.CeilingPlan,
        ViewType.Elevation,
        ViewType.Section,
        ViewType.Detail,
        ViewType.ThreeD
    ]

    available_views = [v for v in all_views if not v.IsTemplate and v.ViewType in valid_view_types]

    # Allow user to select views
    view_choices = ['{} ({})'.format(v.Name, v.ViewType) for v in available_views]
    selected_view_names = forms.SelectFromList.show(
        view_choices,
        multiselect=True,
        title='Select Views for Dimension Placement'
    )

    if not selected_view_names:
        forms.alert('No views selected.')
        script.exit()

    # Map selected view names to view objects
    view_name_map = { '{} ({})'.format(v.Name, v.ViewType): v for v in available_views }
    for view_name in selected_view_names:
        view = view_name_map.get(view_name)
        if view:
            selected_views.append(view)

    if not selected_views:
        forms.alert('No valid views selected.')
        script.exit()

elif selected_view_selection == 'Select Viewports on Sheet':
    # Allow the user to select viewports on a sheet
    class ViewportSelectionFilter(ISelectionFilter):
        def AllowElement(self, element):
            if isinstance(element, Viewport):
                return True
            return False
        def AllowReference(self, reference, position):
            return False

    selection = uidoc.Selection
    prompt_message = 'Please select the viewports on the sheet.'
    try:
        selected_refs = selection.PickObjects(
            ObjectType.Element,
            ViewportSelectionFilter(),
            prompt_message
        )
    except Autodesk.Revit.Exceptions.OperationCanceledException:
        forms.alert('Selection canceled.')
        script.exit()
        selected_refs = None

    if not selected_refs or len(selected_refs) == 0:
        forms.alert('No viewports selected.')
        script.exit()

    # Get the views from the selected viewports
    for sel_ref in selected_refs:
        element = doc.GetElement(sel_ref.ElementId)
        if isinstance(element, Viewport):
            view_id = element.ViewId
            view = doc.GetElement(view_id)
            if view:
                selected_views.append(view)
    if not selected_views:
        forms.alert('No valid views selected from the viewports.')
        script.exit()
else:
    forms.alert('Invalid selection.')
    script.exit()

# Proceed with the rest of the script...

# Collect all categories that allow bound parameters
all_categories = sorted(
    [c for c in doc.Settings.Categories if c.AllowsBoundParameters],
    key=lambda x: x.Name
)

# Allow user to select categories
selected_categories = forms.SelectFromList.show(
    [c.Name for c in all_categories],
    multiselect=True,
    title='Select Categories'
)

if not selected_categories:
    forms.alert('No categories selected.')
    script.exit()

# Map category names to category objects
category_map = {c.Name: c for c in all_categories}
selected_category_ids = [category_map[name].Id for name in selected_categories]

# Collect all family symbols (types) in the selected categories from the host document
all_family_symbols = FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()
host_family_symbols = [fs for fs in all_family_symbols if fs.Category and fs.Category.Id in selected_category_ids]

# Collect family symbols from linked documents
linked_family_symbols = []
for link_instance in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
    link_doc = link_instance.GetLinkDocument()
    if link_doc:
        link_symbols = FilteredElementCollector(link_doc).OfClass(FamilySymbol).ToElements()
        symbols = [fs for fs in link_symbols if fs.Category and fs.Category.Id in selected_category_ids]
        linked_family_symbols.extend([(fs, link_instance) for fs in symbols])

# Prepare type choices from host and linked types
type_choices = []

# For host family symbols
host_type_map = {}
for fs in host_family_symbols:
    family_name = fs.Family.Name
    symbol_param = fs.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    symbol_name = symbol_param.AsString() if symbol_param else 'Unnamed Symbol'
    display_name = 'Host: {} : {}'.format(family_name, symbol_name)
    type_choices.append(display_name)
    host_type_map[display_name] = fs

# For linked family symbols
linked_type_map = {}
for fs, link_instance in linked_family_symbols:
    link_doc = link_instance.GetLinkDocument()
    if link_doc:
        link_name = link_doc.Title
    else:
        link_name = 'Unloaded Link'
    family_name = fs.Family.Name
    symbol_param = fs.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    symbol_name = symbol_param.AsString() if symbol_param else 'Unnamed Symbol'
    display_name = 'Linked ({}) : {} : {}'.format(link_name, family_name, symbol_name)
    type_choices.append(display_name)
    linked_type_map[display_name] = (fs, link_instance)

# Allow user to select types
selected_types = forms.SelectFromList.show(
    type_choices,
    multiselect=True,
    title='Select Family Types'
)

if not selected_types:
    forms.alert('No family types selected.')
    script.exit()

# Map selected type names to their corresponding FamilySymbol objects
selected_host_type_ids = [host_type_map[name].Id for name in selected_types if name in host_type_map]
selected_linked_types = [linked_type_map[name] for name in selected_types if name in linked_type_map]

# Dimension reference options
dimension_reference_options = ['Outer Edge/Face', 'Center Point']
selected_dimension_reference = forms.SelectFromList.show(
    dimension_reference_options,
    multiselect=False,
    title='Select Dimension Reference for Selected Instances'
)

if not selected_dimension_reference:
    forms.alert('No dimension reference selected.')
    script.exit()

# Dimension targets
dimension_target_options = ['Nearest Gridlines', 'Nearest Wall Faces', 'Other Instances']
selected_dimension_targets = forms.SelectFromList.show(
    dimension_target_options,
    multiselect=True,
    title='Select Dimension Targets'
)

if not selected_dimension_targets:
    forms.alert('No dimension targets selected.')
    script.exit()

# Collect dimension styles (types)
dimension_types = FilteredElementCollector(doc)\
    .OfClass(DimensionType)\
    .ToElements()

# Prepare dimension style choices
dimension_style_choices = []
dimension_type_map = {}

for dt in dimension_types:
    dt_param = dt.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
    dt_name = dt_param.AsString() if dt_param else 'Unnamed Dimension Style'
    dimension_style_choices.append(dt_name)
    dimension_type_map[dt_name] = dt

# Allow user to select dimension style
selected_dimension_style_name = forms.SelectFromList.show(
    dimension_style_choices,
    multiselect=False,
    title='Select Dimension Style'
)

if not selected_dimension_style_name:
    forms.alert('No dimension style selected.')
    script.exit()

# Get the selected dimension type
selected_dimension_type = dimension_type_map.get(selected_dimension_style_name, None)

if not selected_dimension_type:
    forms.alert('Selected dimension style not found.')
    script.exit()

# Begin transaction
t = Transaction(doc, 'Dimension Selected Elements')
t.Start()

try:
    # Function to get references based on user selection
    def get_references(element, ref_option, view):
        refs = []
        options = Options()
        options.ComputeReferences = True
        options.IncludeNonVisibleObjects = False
        options.View = view  # Use the specific view to get the correct references
        geom_elem = element.get_Geometry(options)
        if geom_elem is None:
            output.print_md('No geometry found for element ID {}'.format(element.Id))
            return refs
        for geom_obj in geom_elem:
            output.print_md('Processing geometry object of type {}'.format(type(geom_obj)))
            if isinstance(geom_obj, Solid):
                if ref_option == 'Outer Edge/Face':
                    for face in geom_obj.Faces:
                        if face.Reference:
                            refs.append(face.Reference)
                elif ref_option == 'Center Point':
                    for edge in geom_obj.Edges:
                        if edge.Reference:
                            refs.append(edge.Reference)
        return refs

    # Function to find nearest elements (gridlines or walls)
    def find_nearest_elements(element, elements_list):
        min_dist = None
        nearest_element = None
        elem_location = element.Location
        if not hasattr(elem_location, 'Point'):
            output.print_md('Element ID {} does not have a valid location point.'.format(element.Id))
            return None
        elem_point = elem_location.Point
        for elem in elements_list:
            elem_curve = None
            if isinstance(elem, Grid):
                elem_curve = elem.Curve
            elif isinstance(elem, Wall):
                loc_curve = elem.Location
                if isinstance(loc_curve, LocationCurve):
                    elem_curve = loc_curve.Curve
            else:
                continue
            if elem_curve:
                project_result = elem_curve.Project(elem_point)
                if project_result is None:
                    continue
                closest_point = project_result.XYZPoint
                dist = elem_point.DistanceTo(closest_point)
                if min_dist is None or dist < min_dist:
                    min_dist = dist
                    nearest_element = elem
        return nearest_element

    # Prepare a list to collect all dimension references
    all_dimension_references = []

    # Iterate over selected views
    for view in selected_views:
        output.print_md('**Processing View:** {}'.format(view.Name))

        # Collect gridlines and walls visible in the view
        if 'Nearest Gridlines' in selected_dimension_targets:
            gridlines_in_view = FilteredElementCollector(doc, view.Id)\
                .OfClass(Grid)\
                .WhereElementIsNotElementType()\
                .ToElements()
        else:
            gridlines_in_view = []

        if 'Nearest Wall Faces' in selected_dimension_targets:
            walls_in_view = FilteredElementCollector(doc, view.Id)\
                .OfClass(Wall)\
                .WhereElementIsNotElementType()\
                .ToElements()
        else:
            walls_in_view = []

        # Collect host instances visible in the view
        host_instances_in_view = []
        if selected_host_type_ids:
            visible_filter = VisibleInViewFilter(doc, view.Id)
            collector = FilteredElementCollector(doc, view.Id)\
                .WhereElementIsNotElementType()\
                .WherePasses(visible_filter)\
                .OfClass(FamilyInstance)\
                .ToElements()
            host_instances_in_view = [e for e in collector if e.GetTypeId() in selected_host_type_ids]

        # Collect linked instances visible in the view
        linked_instances_in_view = []
        for fs, link_instance in selected_linked_types:
            link_doc = link_instance.GetLinkDocument()
            if link_doc:
                # Get the transform from the linked document to the host document
                transform = link_instance.GetTotalTransform()
                # Collect instances of the selected type
                collector = FilteredElementCollector(link_doc)\
                    .OfClass(FamilyInstance)\
                    .WhereElementIsNotElementType()
                instances = [e for e in collector if e.GetTypeId() == fs.Id]
                # Filter instances that are visible in the view
                for inst in instances:
                    # Transform the instance's bounding box to host coordinates
                    bbox = inst.get_BoundingBox(None)
                    if bbox:
                        transformed_bbox = BoundingBoxXYZ()
                        transformed_bbox.Min = transform.OfPoint(bbox.Min)
                        transformed_bbox.Max = transform.OfPoint(bbox.Max)
                        # Check if the transformed bounding box intersects the view's crop box
                        view_bbox = view.CropBox
                        if view_bbox.Contains(transformed_bbox.Min) or view_bbox.Contains(transformed_bbox.Max):
                            linked_instances_in_view.append((inst, link_instance))

        # Dimension between elements and targets
        for inst in host_instances_in_view:
            inst_refs = get_references(inst, selected_dimension_reference, view)
            if not inst_refs:
                output.print_md('No references found for element ID {}'.format(inst.Id))
                continue
            else:
                output.print_md('Found {} references for element ID {}'.format(len(inst_refs), inst.Id))

            dimension_references = []
            # Dimension to nearest gridline
            if 'Nearest Gridlines' in selected_dimension_targets and gridlines_in_view:
                nearest_grid = find_nearest_elements(inst, gridlines_in_view)
                if nearest_grid:
                    grid_ref = Reference(nearest_grid)
                    dimension_references.append((inst_refs[0], grid_ref))

            # Dimension to nearest wall face
            if 'Nearest Wall Faces' in selected_dimension_targets and walls_in_view:
                nearest_wall = find_nearest_elements(inst, walls_in_view)
                if nearest_wall:
                    wall_refs = get_references(nearest_wall, 'Outer Edge/Face', view)
                    if wall_refs:
                        dimension_references.append((inst_refs[0], wall_refs[0]))

            # Dimension to other selected instances
            if 'Other Instances' in selected_dimension_targets:
                for other_inst in host_instances_in_view:
                    if other_inst.Id == inst.Id:
                        continue
                    other_refs = get_references(other_inst, selected_dimension_reference, view)
                    if other_refs:
                        dimension_references.append((inst_refs[0], other_refs[0]))
                for linked_inst, link_instance in linked_instances_in_view:
                    other_refs = get_references(linked_inst, selected_dimension_reference, view)
                    if other_refs:
                        transformed_ref = other_refs[0].CreateLinkReference(link_instance)
                        dimension_references.append((inst_refs[0], transformed_ref))

            # Collect dimension references
            if dimension_references:
                for ref_pair in dimension_references:
                    all_dimension_references.append((view, inst, ref_pair))

    # After collecting all_dimension_references
    if not all_dimension_references:
        output.print_md('No dimension references found.')
        t.RollBack()
    else:
        # Output the list of dimension references
        output.print_md('**Dimension References Found:**')
        for idx, (view, inst, ref_pair) in enumerate(all_dimension_references):
            ref1, ref2 = ref_pair
            output.print_md('{}. View: {}, Element ID: {}, Ref1 ID: {}, Ref2 ID: {}'.format(
                idx + 1, view.Name, inst.Id, ref1.ElementId, ref2.ElementId))

        # Confirm before placing dimensions
        proceed = forms.alert('Found {} dimension references. Do you want to proceed with placing dimensions?'.format(len(all_dimension_references)),
                              yes=True, no=True)
        if not proceed:
            t.RollBack()
        else:
            # Create dimensions
            for idx, (view, inst, ref_pair) in enumerate(all_dimension_references):
                ref1, ref2 = ref_pair
                try:
                    # Get points from the references
                    try:
                        ref1_point = ref1.GlobalPoint
                    except:
                        ref1_point = inst.Location.Point
                    try:
                        ref2_point = ref2.GlobalPoint
                    except:
                        ref2_point = None
                        if isinstance(ref2.ElementId, ElementId):
                            ref2_element = doc.GetElement(ref2.ElementId)
                            if ref2_element:
                                ref2_location = ref2_element.Location
                                if hasattr(ref2_location, 'Point'):
                                    ref2_point = ref2_location.Point
                    if not ref1_point or not ref2_point:
                        output.print_md('Cannot create dimension without valid points for elements {} and {}.'.format(ref1.ElementId, ref2.ElementId))
                        continue  # Cannot create dimension without valid points

                    # Create a dimension line between the two points
                    dim_line = Line.CreateBound(ref1_point, ref2_point)

                    # Create the dimension
                    dim = doc.Create.NewDimension(view, dim_line, List[Reference]([ref1, ref2]), selected_dimension_type)
                    if dim:
                        output.print_md('Dimension created in view "{}" between elements {} and {}.'.format(
                            view.Name, ref1.ElementId, ref2.ElementId))
                    else:
                        output.print_md('Failed to create dimension in view "{}" between elements {} and {}.'.format(
                            view.Name, ref1.ElementId, ref2.ElementId))
                except Exception as e:
                    output.print_md('Error creating dimension in view "{}": {}'.format(view.Name, str(e)))
                    continue
            t.Commit()
            forms.alert('Dimensions have been created successfully.')

except Exception as e:
    if t.HasStarted():
        t.RollBack()
    forms.alert('An error occurred: {}'.format(str(e)))
    raise
