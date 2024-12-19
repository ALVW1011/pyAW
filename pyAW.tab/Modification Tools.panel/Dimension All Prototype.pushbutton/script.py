# -*- coding: utf-8 -*-
"""
Simplified Wall Dimensioning Script
-----------------------------------
This script dimensions walls in selected views, with options for using wall location lines,
nearest gridlines, or other specified references. The functionality is tailored to simplify
the process while maintaining flexibility.
"""

__title__ = 'Simplified Wall Dimensioning'
__author__ = 'Your Name'

from pyrevit import forms
from pyrevit import script
from Autodesk.Revit.DB import (
    View, ViewType, Wall, Grid, DimensionType, FilteredElementCollector,
    BuiltInParameter, Transaction, Line, ReferenceArray
)
from Autodesk.Revit.UI import *

uiapp = __revit__
uidoc = uiapp.ActiveUIDocument
doc = uidoc.Document
output = script.get_output()

# Utility Functions
def get_selected_views():
    view_selection_options = ['Use Current View', 'Select Views from List']
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
        all_views = FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType().ToElements()
        valid_view_types = [
            ViewType.FloorPlan,
            ViewType.CeilingPlan,
            ViewType.Elevation,
            ViewType.Section
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
        return [view_name_map[name] for name in selected_view_names]

# Collect Walls and Filter by Type
def collect_and_filter_walls(view):
    walls = FilteredElementCollector(doc, view.Id).OfClass(Wall).WhereElementIsNotElementType().ToElements()

    wall_types = {}
    for wall in walls:
        try:
            wall_type = wall.WallType
            if wall_type and wall_type.Name not in wall_types:
                wall_types[wall_type.Name] = wall_type
        except AttributeError as e:
            output.print_md(f'Error accessing WallType for wall ID {wall.Id}: {e}')

    if not wall_types:
        forms.alert('No wall types found in the selected view. Aborting.')
        return None

    selected_wall_types = forms.SelectFromList.show(
        sorted(wall_types.keys()),
        multiselect=True,
        title='Select Wall Types to Include in Dimensioning'
    )

    if not selected_wall_types:
        forms.alert('No wall types selected. Aborting.')
        return None

    filtered_walls = []
    for wall in walls:
        try:
            if wall.WallType and wall.WallType.Name in selected_wall_types:
                filtered_walls.append(wall)
        except AttributeError as e:
            output.print_md(f'Error filtering wall ID {wall.Id}: {e}')

    return filtered_walls

# Dimension Walls
def dimension_walls(view, walls, include_gridlines, location_line_option, dimension_type):
    grids = FilteredElementCollector(doc, view.Id).OfClass(Grid).WhereElementIsNotElementType().ToElements() if include_gridlines else []

    dimension_references = []

    for wall in walls:
        location_curve = getattr(wall.Location, 'Curve', None)
        if not location_curve:
            output.print_md('Wall ID {} has no valid location curve.'.format(wall.Id))
            continue

        line = Line.CreateBound(location_curve.GetEndPoint(0), location_curve.GetEndPoint(1))
        ref_array = ReferenceArray()
        ref_array.Append(location_curve.Reference)

        if include_gridlines:
            for grid in grids:
                grid_curve = grid.Curve
                projection = grid_curve.Project(location_curve.GetEndPoint(0))
                if projection and projection.XYZPoint:
                    ref_array.Append(grid_curve.Reference)

        dimension_references.append((view, ref_array, line))

    for vw, ref_array, line in dimension_references:
        try:
            dim = doc.Create.NewDimension(vw, line, ref_array, dimension_type)
            output.print_md('Dimension created: {}'.format(dim.Id))
        except Exception as e:
            output.print_md('Error creating dimension: {}'.format(e))

# Main Function
def main():
    # Step 1: Select Views
    selected_views = get_selected_views()
    if not selected_views:
        return

    for view in selected_views:
        # Step 2: Collect and Filter Walls
        walls = collect_and_filter_walls(view)
        if not walls:
            continue

        # Step 3: Include or Omit Gridlines
        include_gridlines = forms.alert('Include Gridlines in Dimensioning?', yes=True, no=True)

        # Step 4: Select Wall Location Line
        location_line_option = forms.SelectFromList.show(
            ['Wall Centerline', 'Core Centerline', 'Finish Face Exterior', 'Finish Face Interior'],
            multiselect=False,
            title='Select Wall Location Line'
        )
        if not location_line_option:
            forms.alert('No location line option selected. Aborting.')
            continue

        # Step 5: Select Dimension Type
        dimension_type_name = forms.SelectFromList.show(
            [dt.Name for dt in FilteredElementCollector(doc).OfClass(DimensionType)],
            multiselect=False,
            title='Select Dimension Style'
        )
        if not dimension_type_name:
            forms.alert('No dimension style selected. Aborting.')
            continue

        dimension_type = None
        for dt in FilteredElementCollector(doc).OfClass(DimensionType):
            if dt.Name == dimension_type_name:
                dimension_type = dt
                break

        if not dimension_type:
            forms.alert('Selected dimension style not found. Aborting.')
            continue

        # Step 6: Execute Dimensioning
        t = Transaction(doc, 'Dimension Walls')
        t.Start()
        try:
            dimension_walls(view, walls, include_gridlines, location_line_option, dimension_type)
            t.Commit()
        except Exception as e:
            output.print_md('Error: {}'.format(e))
            t.RollBack()

if __name__ == '__main__':
    main()
