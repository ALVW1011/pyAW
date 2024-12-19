# -*- coding: utf-8 -*-
"""
Dimension Selected Elements (Incorporating Geometry Extraction from Snippet)
----------------------------------------------------------------------------
This script treats all chosen elements equally (both host and linked) and attempts
to dimension them based on selected criteria. It incorporates more robust geometry
extraction logic, inspired by the snippet you provided, to unwrap geometry instances
and find solids, faces, or edges for dimension references.

Works for plan, elevation, and section views, assuming elements have 3D geometry at
the chosen detail level.
"""

__title__ = 'Dimension Selected Elements (Improved Geometry)'
__author__ = 'Your Name'

from pyrevit import forms
from pyrevit import script
from Autodesk.Revit.DB import (
    View, ViewType, FamilySymbol, Wall, Grid, DimensionType, RevitLinkInstance,
    FamilyInstance, FilteredElementCollector, BuiltInParameter, Transaction,
    ElementId, Options, Solid, LocationCurve, BoundingBoxXYZ, XYZ,
    VisibleInViewFilter, Reference, Line, GeometryInstance, UnitUtils, UnitTypeId
)
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ISelectionFilter
from System.Collections.Generic import List
import Autodesk

uiapp = __revit__
uidoc = uiapp.ActiveUIDocument
doc = uidoc.Document
output = script.get_output()

def bounding_box_contains_point(bbox, point):
    return (bbox.Min.X <= point.X <= bbox.Max.X and
            bbox.Min.Y <= point.Y <= bbox.Max.Y and
            bbox.Min.Z <= point.Z <= bbox.Max.Z)

def get_selected_views():
    view_selection_options = ['Use Current View', 'Select Views from List', 'Select Viewports on Sheet']
    selected_view_selection = forms.SelectFromList.show(
        view_selection_options,
        multiselect=False,
        title='How would you like to select the views for dimension placement?'
    )

    if not selected_view_selection:
        forms.alert('No view selection method chosen.')
        return None

    if selected_view_selection == 'Use Current View':
        return [uidoc.ActiveView]

    elif selected_view_selection == 'Select Views from List':
        all_views = FilteredElementCollector(doc)\
            .OfClass(View)\
            .WhereElementIsNotElementType()\
            .ToElements()

        valid_view_types = [
            ViewType.FloorPlan,
            ViewType.CeilingPlan,
            ViewType.Elevation,
            ViewType.Section,
            ViewType.Detail,
            ViewType.ThreeD
        ]
        available_views = [v for v in all_views if (not v.IsTemplate and v.ViewType in valid_view_types)]

        view_choices = ['{} ({})'.format(v.Name, v.ViewType) for v in available_views]
        selected_view_names = forms.SelectFromList.show(
            view_choices,
            multiselect=True,
            title='Select Views for Dimension Placement'
        )

        if not selected_view_names:
            forms.alert('No views selected.')
            return None

        view_name_map = { '{} ({})'.format(v.Name, v.ViewType): v for v in available_views }
        selected_views = []
        for view_name in selected_view_names:
            view = view_name_map.get(view_name)
            if view:
                selected_views.append(view)

        if not selected_views:
            forms.alert('No valid views selected.')
            return None
        return selected_views

    elif selected_view_selection == 'Select Viewports on Sheet':
        class ViewportSelectionFilter(ISelectionFilter):
            def AllowElement(self, element):
                return isinstance(element, Viewport)
            def AllowReference(self, reference, position):
                return False

        selection = uidoc.Selection
        prompt_message = 'Please select the viewports on the sheet.'
        try:
            selected_refs = selection.PickObjects(ObjectType.Element, ViewportSelectionFilter(), prompt_message)
        except Autodesk.Revit.Exceptions.OperationCanceledException:
            forms.alert('Selection canceled.')
            return None

        if not selected_refs:
            forms.alert('No viewports selected.')
            return None

        selected_views = []
        for sel_ref in selected_refs:
            element = doc.GetElement(sel_ref.ElementId)
            if isinstance(element, Viewport):
                view_id = element.ViewId
                view = doc.GetElement(view_id)
                if view:
                    selected_views.append(view)

        if not selected_views:
            forms.alert('No valid views selected from the viewports.')
            return None
        return selected_views
    else:
        forms.alert('Invalid selection.')
        return None

def select_categories():
    all_categories = sorted(
        [c for c in doc.Settings.Categories if c.AllowsBoundParameters],
        key=lambda x: x.Name
    )

    selected_categories = forms.SelectFromList.show(
        [c.Name for c in all_categories],
        multiselect=True,
        title='Select Categories'
    )

    if not selected_categories:
        forms.alert('No categories selected.')
        return None
    category_map = {c.Name: c for c in all_categories}
    return [category_map[name].Id for name in selected_categories]

def select_family_types(selected_category_ids):
    all_family_symbols = FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()
    host_family_symbols = [fs for fs in all_family_symbols if fs.Category and fs.Category.Id in selected_category_ids]

    linked_family_symbols = []
    for link_instance in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        link_doc = link_instance.GetLinkDocument()
        if link_doc:
            link_symbols = FilteredElementCollector(link_doc).OfClass(FamilySymbol).ToElements()
            symbols = [fs for fs in link_symbols if fs.Category and fs.Category.Id in selected_category_ids]
            linked_family_symbols.extend([(fs, link_instance) for fs in symbols])

    type_choices = []
    host_type_map = {}
    for fs in host_family_symbols:
        family_name = fs.Family.Name
        symbol_param = fs.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        symbol_name = symbol_param.AsString() if symbol_param else 'Unnamed Symbol'
        display_name = 'Host: {} : {}'.format(family_name, symbol_name)
        type_choices.append(display_name)
        host_type_map[display_name] = fs

    linked_type_map = {}
    for fs, link_instance in linked_family_symbols:
        link_doc = link_instance.GetLinkDocument()
        link_name = link_doc.Title if link_doc else 'Unloaded Link'
        family_name = fs.Family.Name
        symbol_param = fs.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        symbol_name = symbol_param.AsString() if symbol_param else 'Unnamed Symbol'
        display_name = 'Linked ({}) : {} : {}'.format(link_name, family_name, symbol_name)
        type_choices.append(display_name)
        linked_type_map[display_name] = (fs, link_instance)

    selected_types = forms.SelectFromList.show(
        type_choices,
        multiselect=True,
        title='Select Family Types'
    )

    if not selected_types:
        forms.alert('No family types selected.')
        return None, None

    selected_host_type_ids = [host_type_map[name].Id for name in selected_types if name in host_type_map]
    selected_linked_types = [linked_type_map[name] for name in selected_types if name in linked_type_map]

    output.print_md("Selected Host Type IDs: {}".format(selected_host_type_ids))
    output.print_md("Selected Linked Types: {}".format([(lt[0].Id.IntegerValue, lt[1].Name) for lt in selected_linked_types]))

    return selected_host_type_ids, selected_linked_types

def select_dimension_preferences():
    dimension_reference_options = ['Outer Edge/Face', 'Center Point']
    selected_dimension_reference = forms.SelectFromList.show(
        dimension_reference_options,
        multiselect=False,
        title='Select Dimension Reference for Selected Instances'
    )

    if not selected_dimension_reference:
        forms.alert('No dimension reference selected.')
        return None, None, None

    dimension_target_options = ['Nearest Gridlines', 'Nearest Wall Faces', 'Other Instances']
    selected_dimension_targets = forms.SelectFromList.show(
        dimension_target_options,
        multiselect=True,
        title='Select Dimension Targets'
    )

    if not selected_dimension_targets:
        forms.alert('No dimension targets selected.')
        return None, None, None

    dimension_types = FilteredElementCollector(doc).OfClass(DimensionType).ToElements()
    dimension_style_choices = []
    dimension_type_map = {}
    for dt in dimension_types:
        dt_param = dt.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        dt_name = dt_param.AsString() if dt_param else 'Unnamed Dimension Style'
        dimension_style_choices.append(dt_name)
        dimension_type_map[dt_name] = dt

    selected_dimension_style_name = forms.SelectFromList.show(
        dimension_style_choices,
        multiselect=False,
        title='Select Dimension Style'
    )

    if not selected_dimension_style_name:
        forms.alert('No dimension style selected.')
        return None, None, None

    selected_dimension_type = dimension_type_map.get(selected_dimension_style_name, None)
    if not selected_dimension_type:
        forms.alert('Selected dimension style not found.')
        return None, None, None

    output.print_md("Selected Dimension Reference: {}".format(selected_dimension_reference))
    output.print_md("Selected Dimension Targets: {}".format(selected_dimension_targets))
    output.print_md("Selected Dimension Style: {}".format(selected_dimension_style_name))

    return selected_dimension_reference, selected_dimension_targets, selected_dimension_type

def get_refs_from_solid(solid, ref_option):
    refs = []
    if ref_option == 'Outer Edge/Face':
        # Extract faces references
        for face in solid.Faces:
            if face.Reference:
                refs.append(face.Reference)
    elif ref_option == 'Center Point':
        # Extract edge references
        for edge in solid.Edges:
            if edge.Reference:
                refs.append(edge.Reference)
    return refs

def get_element_solids(element, view):
    """Attempt to retrieve all solids from the element geometry."""
    options = Options()
    options.ComputeReferences = True
    options.IncludeNonVisibleObjects = False
    options.View = view

    geom_elem = element.get_Geometry(options)
    if geom_elem is None:
        output.print_md('No geometry for element ID {} in view "{}"'.format(element.Id, view.Name))
        return []

    solids = []

    def process_geometry(geom):
        local_solids = []
        for g in geom:
            if isinstance(g, GeometryInstance):
                inst_geom = g.GetInstanceGeometry()
                local_solids.extend(process_geometry(inst_geom))
            elif isinstance(g, Solid) and g.Volume > 1e-6:
                local_solids.append(g)
            # Ignore other geometry objects that aren't solids
        return local_solids

    solids = process_geometry(geom_elem)
    return solids

def extract_element_geometry(element, view, ref_option):
    """Extract references (faces/edges) from element geometry."""
    solids = get_element_solids(element, view)
    if not solids:
        output.print_md('No solid geometry found for element ID {} in view "{}". Might be symbolic or detail level insufficient.'.format(element.Id, view.Name))
        return []

    all_refs = []
    for solid in solids:
        refs_from_solid = get_refs_from_solid(solid, ref_option)
        all_refs.extend(refs_from_solid)

    if not all_refs:
        output.print_md('No references extracted for element ID {} with option "{}" in view "{}"'.format(element.Id, ref_option, view.Name))

    return all_refs

def find_nearest_elements(element, elements_list):
    min_dist = None
    nearest_element = None
    elem_location = element.Location

    if not hasattr(elem_location, 'Point'):
        output.print_md('Element ID {} does not have a valid location point.'.format(element.Id))
        return None

    elem_point = elem_location.Point
    for elem in elements_list:
        if isinstance(elem, Grid):
            elem_curve = elem.Curve
        elif isinstance(elem, Wall):
            loc_curve = elem.Location
            elem_curve = loc_curve.Curve if isinstance(loc_curve, LocationCurve) else None
        else:
            continue

        if elem_curve:
            project_result = elem_curve.Project(elem_point)
            if project_result:
                dist = elem_point.DistanceTo(project_result.XYZPoint)
                if min_dist is None or dist < min_dist:
                    min_dist = dist
                    nearest_element = elem

    if nearest_element:
        output.print_md('Nearest element to ID {} is {} at distance {}'.format(element.Id, nearest_element.Id, min_dist))
    else:
        output.print_md('No nearest element found for ID {} among given elements.'.format(element.Id))

    return nearest_element

def collect_all_elements_in_view(view, category_ids, host_type_ids, linked_types):
    vis_filter = VisibleInViewFilter(doc, view.Id)
    collector = FilteredElementCollector(doc, view.Id)\
        .WhereElementIsNotElementType()\
        .WherePasses(vis_filter)

    chosen_host_elements = []
    for e in collector:
        if e.Category and e.Category.Id in category_ids:
            # If host_type_ids is empty, that means no host types specifically chosen, allow all from that category
            if (not host_type_ids) or (e.GetTypeId() in host_type_ids):
                chosen_host_elements.append((e, False, None))

    output.print_md('Host chosen elements in view "{}": {}'.format(view.Name, len(chosen_host_elements)))

    chosen_linked_elements = []
    view_bbox = view.CropBox
    for (fs, link_instance) in linked_types:
        link_doc = link_instance.GetLinkDocument()
        if not link_doc:
            continue
        link_collector = FilteredElementCollector(link_doc)\
            .OfClass(FamilyInstance)\
            .WhereElementIsNotElementType()
        instances = [inst for inst in link_collector if inst.GetTypeId() == fs.Id]

        count_visible = 0
        for inst in instances:
            bbox = inst.get_BoundingBox(None)
            if bbox:
                transform = link_instance.GetTotalTransform()
                transformed_min = transform.OfPoint(bbox.Min)
                transformed_max = transform.OfPoint(bbox.Max)
                if (bounding_box_contains_point(view_bbox, transformed_min) or
                    bounding_box_contains_point(view_bbox, transformed_max)):
                    chosen_linked_elements.append((inst, True, link_instance))
                    count_visible += 1

        output.print_md('Checking {} linked instances of type ID {} in view "{}". Found {} visible.'.format(
            len(instances), fs.Id, view.Name, count_visible))

    total_elements = chosen_host_elements + chosen_linked_elements
    output.print_md('Total chosen elements in view "{}": {}'.format(view.Name, len(total_elements)))
    return total_elements

def create_dimensions(doc, dimension_type, dimension_references):
    for idx, (vw, elem, ref_pair) in enumerate(dimension_references):
        ref1, ref2 = ref_pair
        try:
            ref1_point = None
            ref2_point = None
            try:
                ref1_point = ref1.GlobalPoint
            except:
                if hasattr(elem.Location, 'Point'):
                    ref1_point = elem.Location.Point
            try:
                ref2_point = ref2.GlobalPoint
            except:
                if isinstance(ref2.ElementId, ElementId):
                    ref2_element = doc.GetElement(ref2.ElementId)
                    if ref2_element and hasattr(ref2_element.Location, 'Point'):
                        ref2_point = ref2_element.Location.Point

            if not ref1_point or not ref2_point:
                output.print_md('Cannot create dimension without valid points for references {} and {}.'.format(ref1.ElementId, ref2.ElementId))
                continue

            dim_line = Line.CreateBound(ref1_point, ref2_point)
            dim = doc.Create.NewDimension(vw, dim_line, List[Reference]([ref1, ref2]), dimension_type)

            if dim:
                output.print_md('Dimension created in view "{}" between elements {} and {}.'.format(
                    vw.Name, ref1.ElementId, ref2.ElementId))
            else:
                output.print_md('Failed to create dimension in view "{}" between elements {} and {}.'.format(
                    vw.Name, ref1.ElementId, ref2.ElementId))
        except Exception as e:
            output.print_md('Error creating dimension in view "{}": {}'.format(vw.Name, str(e)))
            continue

def main():
    selected_views = get_selected_views()
    if not selected_views:
        return

    selected_category_ids = select_categories()
    if not selected_category_ids:
        return

    selected_host_type_ids, selected_linked_types = select_family_types(selected_category_ids)
    if selected_host_type_ids is None:
        return

    (selected_dimension_reference, selected_dimension_targets,
     selected_dimension_type) = select_dimension_preferences()
    if not all([selected_dimension_reference, selected_dimension_targets, selected_dimension_type]):
        return

    t = Transaction(doc, 'Dimension Selected Elements')
    t.Start()
    try:
        all_dimension_references = []

        for view in selected_views:
            output.print_md('**Processing View:** {}'.format(view.Name))

            chosen_elements = collect_all_elements_in_view(view, selected_category_ids, selected_host_type_ids, selected_linked_types)

            # Collect targets
            if 'Nearest Gridlines' in selected_dimension_targets:
                gridlines_in_view = FilteredElementCollector(doc, view.Id)\
                    .OfClass(Grid)\
                    .WhereElementIsNotElementType()\
                    .ToElements()
                output.print_md('Gridlines in view "{}": {}'.format(view.Name, len(gridlines_in_view)))
            else:
                gridlines_in_view = []

            if 'Nearest Wall Faces' in selected_dimension_targets:
                walls_in_view = FilteredElementCollector(doc, view.Id)\
                    .OfClass(Wall)\
                    .WhereElementIsNotElementType()\
                    .ToElements()
                output.print_md('Walls in view "{}": {}'.format(view.Name, len(walls_in_view)))
            else:
                walls_in_view = []

            # "Other Instances" targets are just the chosen elements themselves, dimension between them.

            for (elem, is_linked, link_inst) in chosen_elements:
                refs = extract_element_geometry(elem, view, selected_dimension_reference)
                if not refs:
                    continue  # No references from this element, skip dimensioning

                dimension_references = []
                # Dimension to nearest gridline
                if 'Nearest Gridlines' in selected_dimension_targets and gridlines_in_view:
                    nearest_grid = find_nearest_elements(elem, gridlines_in_view)
                    if nearest_grid:
                        dimension_references.append((refs[0], Reference(nearest_grid)))

                # Dimension to nearest wall face
                if 'Nearest Wall Faces' in selected_dimension_targets and walls_in_view:
                    nearest_wall = find_nearest_elements(elem, walls_in_view)
                    if nearest_wall:
                        wall_refs = extract_element_geometry(nearest_wall, view, 'Outer Edge/Face')
                        if wall_refs:
                            dimension_references.append((refs[0], wall_refs[0]))

                # Dimension to other instances
                if 'Other Instances' in selected_dimension_targets:
                    for (other_elem, other_is_linked, other_link_inst) in chosen_elements:
                        if other_elem.Id == elem.Id:
                            continue
                        other_refs = extract_element_geometry(other_elem, view, selected_dimension_reference)
                        if other_refs:
                            # If needed, handle linked references. For now, assume direct references are okay.
                            dimension_references.append((refs[0], other_refs[0]))

                if dimension_references:
                    for ref_pair in dimension_references:
                        all_dimension_references.append((view, elem, ref_pair))
                else:
                    output.print_md('Element ID {}: No dimension reference pairs created.'.format(elem.Id))

        if not all_dimension_references:
            output.print_md('No dimension references found.')
            output.print_md('Possible reasons:')
            output.print_md('- Elements do not show 3D geometry in this view or detail level.')
            output.print_md('- Selected reference option ("Outer Edge/Face" or "Center Point") not compatible with element geometry.')
            output.print_md('- No suitable targets found nearby (walls, gridlines, or other instances).')
            output.print_md('Try adjusting detail level, categories, or test with known 3D model elements (e.g., walls in a plan view at Fine detail).')
            t.RollBack()
        else:
            output.print_md('**Dimension References Found:**')
            for idx, (vw, inst, ref_pair) in enumerate(all_dimension_references):
                ref1, ref2 = ref_pair
                output.print_md('{}. View: {}, Element ID: {}, Ref1 ID: {}, Ref2 ID: {}'.format(
                    idx + 1, vw.Name, inst.Id, ref1.ElementId, ref2.ElementId))

            proceed = forms.alert('Found {} dimension references. Proceed?'.format(len(all_dimension_references)),
                                  yes=True, no=True)
            if not proceed:
                t.RollBack()
            else:
                create_dimensions(doc, selected_dimension_type, all_dimension_references)
                t.Commit()
                forms.alert('Dimensions have been created successfully.')

    except Exception as e:
        if t.HasStarted():
            t.RollBack()
        forms.alert('An error occurred: {}'.format(str(e)))
        raise

if __name__ == "__main__":
    main()
