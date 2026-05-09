# -*- coding: utf-8 -*-
__title__  = "_filters"
__author__  = "Seed43"
__doc__     = """
VERSION 260507
_____________________________________________________________________
Description:
Shared helper functions for creating and applying Revit View Filters.

Used by all Filter Manager tools. Import the functions you need
rather than importing the whole module.

Example:
    from Snippets._filters import (
        find_existing_filter,
        get_unique_filter_name,
        view_template_controls_filters,
        apply_filter_to_target,
        create_parameter_filter,
    )
_____________________________________________________________________
Functions:
find_existing_filter(filter_name)
    Search the project for a filter by name.

get_unique_filter_name(base_name)
    Return a unique filter name, appending a counter if needed.

view_template_controls_filters(view)
    Check if the view template controls the filter setting.

apply_filter_to_target(param_filter, doc, revit)
    Add and hide a filter on the active view or its template.

create_parameter_filter(filter_name, category, param_id,
                        filter_value, doc)
    Create a string-equals ParameterFilterElement.

create_category_filter(filter_name, category, doc)
    Create a category-only filter that matches every element.
_____________________________________________________________________
Last update:
- Initial release, extracted from Filter Manager scripts
_____________________________________________________________________
"""

from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ParameterFilterElement,
    ParameterValueProvider,
    FilterStringEquals,
    FilterStringRule,
    FilterNumericGreater,
    FilterIntegerRule,
    ElementParameterFilter,
    ElementId,
    BuiltInParameter,
    View,
)


# ── FILTER LOOKUP ─────────────────────────────────────────────────────────────

def find_existing_filter(filter_name, doc):
    """Return a ParameterFilterElement by name, or None if not found."""
    for f in FilteredElementCollector(doc).OfClass(ParameterFilterElement):
        if f.Name == filter_name:
            return f
    return None


def get_unique_filter_name(base_name, doc):
    """Return a unique filter name by appending a counter if needed."""
    if not find_existing_filter(base_name, doc):
        return base_name
    i = 2
    while True:
        candidate = "{} ({})".format(base_name, i)
        if not find_existing_filter(candidate, doc):
            return candidate
        i += 1


# ── VIEW TEMPLATE ─────────────────────────────────────────────────────────────

def view_template_controls_filters(view, doc):
    """Return True if the view's template is controlling the Filters setting."""
    try:
        if not view.ViewTemplateId \
                or view.ViewTemplateId == ElementId.InvalidElementId:
            return False
        template = doc.GetElement(view.ViewTemplateId)
        if not template or not isinstance(template, View) \
                or not template.IsTemplate:
            return False
        control_ids      = list(template.GetTemplateParameterIds())
        filters_param_id = ElementId(BuiltInParameter.VIS_GRAPHICS_FILTERS)
        return filters_param_id in control_ids
    except Exception:
        return False


# ── APPLY FILTER ──────────────────────────────────────────────────────────────

def apply_filter_to_target(param_filter, doc, revit):
    """
    Add a filter to the active view or its template and hide it.

    If the view has a template that controls filters, the filter is
    applied to the template instead of the view directly.
    Returns True on success, False if the active view is not a View.
    """
    active_view = doc.ActiveView
    if not isinstance(active_view, View):
        return False

    target_view = active_view
    if active_view.ViewTemplateId != ElementId.InvalidElementId:
        template = doc.GetElement(active_view.ViewTemplateId)
        if template and isinstance(template, View) and template.IsTemplate:
            if view_template_controls_filters(active_view, doc):
                target_view = template

    with revit.Transaction("Apply Filter to View or Template"):
        if not target_view.IsFilterApplied(param_filter.Id):
            target_view.AddFilter(param_filter.Id)
        target_view.SetFilterVisibility(param_filter.Id, False)

    return True


# ── CREATE FILTERS ────────────────────────────────────────────────────────────

def create_parameter_filter(filter_name, category, param_id, filter_value, doc):
    """
    Create a ParameterFilterElement using a string equals rule.

    Args:
        filter_name:  Name for the new filter.
        category:     Revit Category object to filter on.
        param_id:     ElementId of the parameter to filter by.
        filter_value: String value to match.
        doc:          Current Revit document.

    Returns a ParameterFilterElement.
    """
    category_ids   = List[ElementId]([category.Id])
    provider       = ParameterValueProvider(param_id)
    evaluator      = FilterStringEquals()
    rule           = FilterStringRule(provider, evaluator, filter_value)
    element_filter = ElementParameterFilter(rule)
    return ParameterFilterElement.Create(doc, filter_name, category_ids, element_filter)


def create_category_filter(filter_name, category, doc):
    """
    Create a ParameterFilterElement that matches every element in the category.

    Uses ID_PARAM greater than 0, which is always valid. Use this as a
    fallback when no suitable string parameter can be found.
    """
    category_ids   = List[ElementId]([category.Id])
    param_id       = ElementId(BuiltInParameter.ID_PARAM)
    provider       = ParameterValueProvider(param_id)
    evaluator      = FilterNumericGreater()
    rule           = FilterIntegerRule(provider, evaluator, 0)
    element_filter = ElementParameterFilter(rule)
    return ParameterFilterElement.Create(doc, filter_name, category_ids, element_filter)
