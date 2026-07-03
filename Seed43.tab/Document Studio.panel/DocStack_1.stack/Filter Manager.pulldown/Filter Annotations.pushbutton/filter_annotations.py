# -*- coding: utf-8 -*-
# filter_annotations.py
# pylint: disable=import-error,invalid-name,broad-except

from System import Int64
from System.Collections.Generic import List
from pyrevit import revit, DB, UI, forms, script
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

# ── [LIB] Snippets/_filters.py ───────────────────────────────────────────────
from Snippets._filters import (
    find_existing_filter,
    get_unique_filter_name,
    apply_filter_to_target,
)

# ── [LIB] Snippets/_selection.py ─────────────────────────────────────────────
from Snippets._selection import get_element_type, get_element_category

doc    = revit.doc
uidoc  = revit.uidoc
logger = script.get_logger()

FILTER_NAME_FORMAT     = "Annotation - {0} ({1} {2}{3})"
FILTER_NAME_FORMAT_ALL = "Annotation - {0} (All)"
NONE_OPTION            = "(none - hide ALL elements of this category)"

# Condition lists shown to the user, matched to Revit's own dialog per storage type.
STRING_OPS = [
    "equals", "does not equal",
    "contains", "does not contain",
    "begins with", "does not begin with",
    "ends with", "does not end with",
    "has a value", "has no value",
]
NUMERIC_OPS = [
    "equals", "does not equal",
    "is greater than", "is greater than or equal to",
    "is less than", "is less than or equal to",
    "has a value", "has no value",
]
ELEMENTID_OPS = [
    "equals", "does not equal",
    "has a value", "has no value",
]

# ── HELPERS ───────────────────────────────────────────────────────────────────

def _safe_bic(name):
    try:
        return getattr(DB.BuiltInCategory, name)
    except AttributeError:
        return None

_HOST_BICS = set(filter(None, [
    _safe_bic("OST_Walls"),             _safe_bic("OST_Floors"),
    _safe_bic("OST_Roofs"),             _safe_bic("OST_Columns"),
    _safe_bic("OST_StructuralColumns"), _safe_bic("OST_Doors"),
    _safe_bic("OST_Windows"),           _safe_bic("OST_Furniture"),
    _safe_bic("OST_GenericModel"),      _safe_bic("OST_MechanicalEquipment"),
    _safe_bic("OST_ElectricalEquipment"), _safe_bic("OST_PlumbingFixtures"),
    _safe_bic("OST_StructuralFraming"), _safe_bic("OST_Stairs"),
    _safe_bic("OST_Ramps"),             _safe_bic("OST_Railings"),
    _safe_bic("OST_Ceilings"),          _safe_bic("OST_CurtainWallPanels"),
    _safe_bic("OST_CurtainWallMullions"),
]))

def is_host_model(category):
    try:
        return category.BuiltInCategory in _HOST_BICS
    except Exception:
        return False

def _first_valid_bic(*names):
    for n in names:
        bic = _safe_bic(n)
        if bic is not None:
            return bic
    return None

VIEW_MARKER_CATEGORIES = {}
for _label, _bic in (
    ("Sections",   _first_valid_bic("OST_Sections")),
    ("Elevations", _first_valid_bic("OST_Elev", "OST_Elevations", "OST_ElevationMarks")),
    ("Callouts",   _first_valid_bic("OST_Callouts", "OST_CalloutHeads", "OST_CalloutBoundary")),
    ("Views",      _first_valid_bic("OST_Viewers")),
):
    if _bic is not None:
        VIEW_MARKER_CATEGORIES[_label] = _bic


def guess_view_marker_category_name(view):
    """A Detail View type can carry both a Section Tag and a Callout Tag,
    so which annotation category its marker actually belongs to isn't
    something the View object itself reliably tells you. ViewType gives
    a reasonable guess, but the user gets the final say, same as they
    would picking a category by hand in the native Filters dialog."""
    try:
        vt = view.ViewType
    except Exception:
        vt = None
    if vt == DB.ViewType.Section:
        return "Sections"
    if vt == DB.ViewType.Elevation:
        return "Elevations"
    if vt == DB.ViewType.Detail:
        return "Callouts"
    return "Views"


PICKER_CANCELLED = object()  # sentinel: user backed out of the marker-category picker


def get_correct_category(element, doc):
    """
    View markers (section heads, elevation tags, callout heads) belong to
    one of four distinct annotation categories that don't map cleanly
    from the View element itself, Revit's own Filters dialog requires a
    manual category pick for exactly this reason. We offer a best guess
    first, sourced from ViewType, but let the user confirm or override.

    isinstance(element, DB.View) doesn't work here, the picked object's
    .NET type reports as plain Element rather than ViewSection in this
    Revit/IronPython build (confirmed by direct testing), so instead we
    detect it by category name, which we already know resolves correctly.
    """
    category = get_element_category(element, doc)
    if category and category.Name == "Views":
        guess = guess_view_marker_category_name(element)
        names = list(VIEW_MARKER_CATEGORIES.keys())
        if guess in names:
            names.remove(guess)
            names.insert(0, guess)
        else:
            names.sort()
        choice = forms.SelectFromList.show(
            names,
            title="Which marker category? (guessed: {})".format(guess),
            button_name="Select Category"
        )
        if choice is None:
            return PICKER_CANCELLED
        return DB.Category.GetCategory(doc, VIEW_MARKER_CATEGORIES[choice])
    return category

class AnnotationSelFilter(ISelectionFilter):
    def AllowElement(self, element):
        return not isinstance(element, DB.RevitLinkInstance)
    def AllowReference(self, reference, xyz):
        return False

# ── REAL FILTERABLE PARAMETERS (this is what image 1 comes from) ─────────────

def _safe_viewtype(name):
    return getattr(DB.ViewType, name, None)

# Revit's own Family-and-Type value list for view markers excludes view
# types that don't really have a normal "type" the way plans/sections/
# details do (Schedules, Sheets, 3D Views, Legends), but it does NOT
# restrict to only the chosen marker category, e.g. "Sections" still
# shows Floor Plans and Structural Plans as valid options. This matches
# that behavior instead of over-filtering to an exact category match.
EXCLUDED_VIEWTYPES_FOR_MARKER_SCAN = set(filter(None, [
    _safe_viewtype("ThreeD"),          _safe_viewtype("Schedule"),
    _safe_viewtype("DrawingSheet"),    _safe_viewtype("Legend"),
    _safe_viewtype("Internal"),        _safe_viewtype("ProjectBrowser"),
    _safe_viewtype("SystemBrowser"),   _safe_viewtype("Report"),
    _safe_viewtype("Walkthrough"),     _safe_viewtype("ColumnSchedule"),
    _safe_viewtype("PanelSchedule"),   _safe_viewtype("Rendering"),
    _safe_viewtype("Undefined"),
]))


def get_filterable_params(category, doc):
    """
    Returns {display_name: ElementId} for every parameter Revit itself
    considers valid for filtering this category. This is the exact same
    list the native 'Filters' dialog uses, so no more guessing.
    """
    cat_ids = List[DB.ElementId]([category.Id])
    try:
        param_ids = DB.ParameterFilterUtilities.GetFilterableParametersInCommon(doc, cat_ids)
    except Exception as ex:
        logger.debug("GetFilterableParametersInCommon failed: {}".format(str(ex)))
        return {}

    result = {}
    for pid in param_ids:
        name = None
        if pid.Value < 0:
            # Built-in parameter -> get its localized UI label
            try:
                bip  = DB.BuiltInParameter(pid.Value)
                name = DB.LabelUtils.GetLabelFor(bip)
            except Exception:
                name = None
        else:
            # Shared / project parameter -> it's a real element in the doc
            elem = doc.GetElement(pid)
            if elem:
                name = elem.Name
        if name:
            result[name] = pid
    return result


# ParameterFilterUtilities.GetFilterableParametersInCommon does not return
# the full parameter set for the "Views" category (sections/elevations/
# callouts) the way it does for ordinary model categories -- it's a real
# gap in that API, not a category-lookup mistake. These ids were captured
# directly from Fred's own element inspector against a live section view,
# so they're confirmed-real rather than guessed. If Create() rejects one,
# it fails with a normal catchable error, it won't crash Revit.
VIEW_CATEGORY_PARAMS = {
    "Detail Level":                 -1011002,
    "Discipline":                   -1005163,
    "Family":                       -1002051,
    "Family and Type":              -1002052,
    "Phase":                        -1012102,
    "Phase Filter":                 -1012103,
    "Referencing Detail":           -1005171,
    "Referencing Sheet":            -1005170,
    "Referencing Sheet Collection": -1005222,
    "Scale Value":                  -1005150,
    "Sheet Collection":             -1005224,
    "Sheet Name":                   -1005223,
    "Sheet Number":                 -1006601,
    "Title on Sheet":               -1005114,
    "View Name":                    -1005112,
}


def find_shared_param_id(doc, name):
    """'Folder' and similar project/shared parameters aren't built-in, so
    they have to be looked up by name instead of a fixed id."""
    for sp in DB.FilteredElementCollector(doc).OfClass(DB.SharedParameterElement):
        if sp.Name == name:
            return sp.Id
    return None


def get_filterable_params_for_category(category, doc):
    marker_cat_ids = set()
    for bic in VIEW_MARKER_CATEGORIES.values():
        cat = DB.Category.GetCategory(doc, bic)
        if cat:
            marker_cat_ids.add(cat.Id.Value)

    if category.Id.Value in marker_cat_ids:
        result = {}
        for name, pid_int in VIEW_CATEGORY_PARAMS.items():
            result[name] = DB.ElementId(Int64(pid_int))
        # Merge in whatever the API DOES confirm (this is where Workset comes from)
        result.update(get_filterable_params(category, doc))
        folder_id = find_shared_param_id(doc, "Folder")
        if folder_id:
            result["Folder"] = folder_id
        return result
    return get_filterable_params(category, doc)


CUSTOM_VALUE_ENTRY = "— Type a custom value —"


def get_distinct_param_values(param_id, category, doc, storage_type, limit=40):
    """
    Scans existing elements of this category for the real values already
    in use for this parameter, so the user can pick from options that
    actually exist instead of typing blind. Essential for enum-style
    Integer params (Discipline='Structural' is really stored as some
    meaningless index like 3), and just plain more convenient for String
    params like Sheet Name or View Name. Falls back to nothing (caller
    then asks for raw typed input) if there are too many distinct values
    to be a useful picker.
    """
    marker_cat_ids = set()
    for bic in VIEW_MARKER_CATEGORIES.values():
        cat = DB.Category.GetCategory(doc, bic)
        if cat:
            marker_cat_ids.add(cat.Id.Value)

    if category.Id.Value in marker_cat_ids:
        # OfCategoryId finds graphical content placed *within* a view
        # (walls, tags, etc), it does not retrieve View objects
        # themselves, so scanning Sections/Elevations/Callouts/Views
        # needs a plain View collector instead.
        all_views = [v for v in DB.FilteredElementCollector(doc).OfClass(DB.View)
                     if not v.IsTemplate]
        collector = [v for v in all_views
                     if v.ViewType not in EXCLUDED_VIEWTYPES_FOR_MARKER_SCAN]
    else:
        collector = DB.FilteredElementCollector(doc) \
            .OfCategoryId(category.Id).WhereElementIsNotElementType()
    seen = {}
    is_builtin = param_id.Value < 0
    bip = None
    if is_builtin:
        try:
            bip = DB.BuiltInParameter(param_id.Value)
        except Exception:
            is_builtin = False

    for el in collector:
        try:
            p = None
            if is_builtin:
                p = el.get_Parameter(bip)
            if p is None:
                for cand in el.Parameters:
                    if cand.Id == param_id:
                        p = cand
                        break
            if p is None or not p.HasValue:
                continue
            if storage_type == DB.StorageType.String:
                raw = p.AsString() or p.AsValueString()
                if not raw:
                    continue
                disp = raw
            elif storage_type == DB.StorageType.ElementId:
                raw = p.AsElementId()
                if raw is None or raw.Value == -1:
                    continue
                disp = p.AsValueString() or str(raw.Value)
            else:
                raw  = p.AsInteger()
                disp = p.AsValueString() or str(raw)
            seen[raw] = disp
            if len(seen) > limit:
                return {}
        except Exception:
            continue
    return seen


def get_param_on_host(param_id, element, element_type):
    """
    Find the Parameter object (and which host it lives on) so we can
    read its StorageType and current value.

    Element.get_Parameter() only accepts BuiltInParameter, Definition, or
    Guid, there is no overload for a raw ElementId, even though that's
    exactly what GetFilterableParametersInCommon hands back. Passing an
    ElementId straight into get_Parameter() throws "expected
    BuiltInParameter, got ElementId". Matching by .Id across the host's
    own parameter list sidesteps that entirely and works the same way
    regardless of whether the parameter is built-in, shared, or project.
    """
    for host in (element, element_type):
        if host is None:
            continue
        p = None
        if param_id.Value < 0:
            try:
                p = host.get_Parameter(DB.BuiltInParameter(param_id.Value))
            except Exception:
                p = None
        if p is None:
            try:
                for cand in host.Parameters:
                    if cand.Id == param_id:
                        p = cand
                        break
            except Exception:
                pass
        if p is not None:
            return p, host
    return None, None

# ── VALUE COLLECTION ──────────────────────────────────────────────────────────

def get_filter_value(param, storage_type, operator, param_id=None, category=None, doc=None):
    """
    Prompts the user for a value appropriate to the parameter's storage type.
    Returns (value_for_rule, display_string_for_filter_name).
    Returns ("SKIP", None) if the user cancels or gives a bad value.
    """
    if operator in ("has a value", "has no value"):
        return None, None

    if storage_type == DB.StorageType.ElementId:
        distinct = {}
        if param_id is not None and category is not None and doc is not None:
            distinct = get_distinct_param_values(param_id, category, doc, storage_type)

        if len(distinct) >= 1:
            labels = sorted(set(distinct.values()))
            choice = forms.SelectFromList.show(
                labels, title="Filter Value", button_name="Select")
            if choice is None:
                return "SKIP", None
            raw = next(k for k, v in distinct.items() if v == choice)
            return raw, choice

        # Fallback: scanning found nothing usable, use the selected
        # element's own current value instead of leaving the user stuck.
        try:
            eid = param.AsElementId()
        except Exception:
            eid = None
        if eid is None or eid.Value == -1:
            forms.alert(
                "The selected element has no value set for this parameter.\n"
                "Pick a different parameter, or select an element that has one.",
                title="No Value")
            return "SKIP", None
        display = param.AsValueString() or str(eid.Value)
        return eid, display

    if storage_type == DB.StorageType.String:
        current = param.AsString() or param.AsValueString() or ""
        distinct = {}
        if param_id is not None and category is not None and doc is not None:
            distinct = get_distinct_param_values(param_id, category, doc, storage_type)

        if len(distinct) >= 1:
            labels = sorted(set(distinct.values())) + [CUSTOM_VALUE_ENTRY]
            choice = forms.SelectFromList.show(
                labels, title="Filter Value", button_name="Select")
            if choice is None:
                return "SKIP", None
            if choice != CUSTOM_VALUE_ENTRY:
                raw = next(k for k, v in distinct.items() if v == choice)
                return raw, choice
            # falls through to typed input below

        typed = forms.ask_for_string(
            default=current, prompt="Enter value to match:", title="Filter Value")
        if typed is None:
            return "SKIP", None
        return typed, typed

    if storage_type == DB.StorageType.Integer:
        # Handle Yes/No parameters with a clean picker instead of raw 0/1 typing
        is_yesno = False
        try:
            is_yesno = (param.Definition.GetDataType() == DB.SpecTypeId.Boolean.YesNo)
        except Exception:
            pass
        if is_yesno:
            choice = forms.SelectFromList.show(
                ["Yes", "No"], title="Filter Value", button_name="Select")
            if choice is None:
                return "SKIP", None
            return (1 if choice == "Yes" else 0), choice

        # Enum-style integers (Discipline, Detail Level, etc): the raw stored
        # number is meaningless on its own, so offer real options if we can
        # find them among existing elements of this category.
        distinct = {}
        if param_id is not None and category is not None and doc is not None:
            distinct = get_distinct_param_values(param_id, category, doc, storage_type)

        if len(distinct) >= 1:
            labels = sorted(set(distinct.values())) + [CUSTOM_VALUE_ENTRY]
            choice = forms.SelectFromList.show(
                labels, title="Filter Value", button_name="Select")
            if choice is None:
                return "SKIP", None
            if choice != CUSTOM_VALUE_ENTRY:
                raw = next(k for k, v in distinct.items() if v == choice)
                return raw, choice
            # falls through to typed input below

        current = param.AsValueString() or str(param.AsInteger())
        typed = forms.ask_for_string(
            default=current, prompt="Enter integer value:", title="Filter Value")
        if typed is None:
            return "SKIP", None
        try:
            return int(typed.strip()), typed
        except Exception:
            forms.alert("'{}' is not a valid integer.".format(typed), title="Invalid Value")
            return "SKIP", None

    if storage_type == DB.StorageType.Double:
        current = param.AsValueString() or ""
        typed = forms.ask_for_string(
            default=current,
            prompt="Enter value (in project display units):",
            title="Filter Value")
        if typed is None:
            return "SKIP", None
        try:
            raw = float(typed.strip())
        except Exception:
            forms.alert("'{}' is not a valid number.".format(typed), title="Invalid Value")
            return "SKIP", None
        try:
            unit_id  = param.GetUnitTypeId()
            internal = DB.UnitUtils.ConvertToInternalUnits(raw, unit_id)
        except Exception:
            internal = raw  # unitless double (ratio, etc.)
        return internal, typed

    return "SKIP", None

# ── FILTER RULE CONSTRUCTION ──────────────────────────────────────────────────

def build_filter_rule(param_id, storage_type, operator, value):
    """Returns (FilterRule, inverted_bool) or (None, False) if unsupported."""
    provider = DB.ParameterValueProvider(param_id)

    if operator in ("has a value", "has no value"):
        rule = DB.HasValueFilterRule(param_id)
        return rule, (operator == "has no value")

    if storage_type == DB.StorageType.String:
        inverted = operator.startswith("does not")
        base_op  = operator.replace("does not ", "")
        evaluator = {
            "equals":      DB.FilterStringEquals(),
            "contains":    DB.FilterStringContains(),
            "begins with": DB.FilterStringBeginsWith(),
            "ends with":   DB.FilterStringEndsWith(),
        }[base_op]
        rule = DB.FilterStringRule(provider, evaluator, value)
        return rule, inverted

    if storage_type in (DB.StorageType.Integer, DB.StorageType.Double):
        inverted = (operator == "does not equal")
        base_op  = "equals" if inverted else operator
        evaluator = {
            "equals":                      DB.FilterNumericEquals(),
            "is greater than":             DB.FilterNumericGreater(),
            "is greater than or equal to": DB.FilterNumericGreaterOrEqual(),
            "is less than":                DB.FilterNumericLess(),
            "is less than or equal to":    DB.FilterNumericLessOrEqual(),
        }[base_op]
        if storage_type == DB.StorageType.Integer:
            rule = DB.FilterIntegerRule(provider, evaluator, value)
        else:
            rule = DB.FilterDoubleRule(provider, evaluator, value, 1e-6)
        return rule, inverted

    if storage_type == DB.StorageType.ElementId:
        inverted  = (operator == "does not equal")
        evaluator = DB.FilterNumericEquals()
        rule = DB.FilterElementIdRule(provider, evaluator, value)
        return rule, inverted

    return None, False


def create_param_filter(filter_name, category, element_filter, doc):
    cat_ids = List[DB.ElementId]([category.Id])
    return DB.ParameterFilterElement.Create(doc, filter_name, cat_ids, element_filter)


def hide_category_in_view(view, category, doc):
    """
    'Hide all elements of this category' — done the way Revit itself does
    it, by toggling category visibility directly in the view. No
    ParameterFilterElement is created, so there's nothing for the Filters
    dialog to choke on. My earlier attempts (a bare Create() call with no
    elementFilter, then a 'match everything' LogicalOrFilter) both crashed
    Revit on categories the API doesn't officially support filter rules
    for. This sidesteps that entirely.
    """
    try:
        if not category.get_AllowsVisibilityControl(view):
            forms.alert(
                "'{}' visibility can't be controlled in this view.".format(
                    category.Name),
                title="Not Allowed")
            return False
    except Exception:
        pass

    with revit.Transaction("Hide Category"):
        view.SetCategoryHidden(category.Id, True)
    return True

# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    try:
        with forms.WarningBar(title="Filter Annotation: Select element. ESC to cancel."):
            while True:
                try:
                    picked_ref = uidoc.Selection.PickObject(
                        ObjectType.Element,
                        AnnotationSelFilter(),
                        "Select annotation or datum element"
                    )
                    element = doc.GetElement(picked_ref.ElementId)

                    if isinstance(element, DB.RevitLinkInstance):
                        forms.alert("Please select a host model element.",
                                    title="Invalid Selection")
                        continue

                    category = get_correct_category(element, doc)
                    if category is PICKER_CANCELLED:
                        continue
                    if not category:
                        forms.alert("Element has no valid category.", title="Error")
                        continue

                    if is_host_model(category):
                        opt = forms.alert(
                            "This looks like a host model element ({}).\n\n"
                            "Use 'Filter Element (Host)' instead.\n\n"
                            "Continue anyway?".format(category.Name),
                            options=["Continue", "Cancel"],
                            title="Wrong Script?"
                        )
                        if opt != "Continue":
                            continue

                    element_type = get_element_type(element, doc)

                    # ── Step 1: pick a real, Revit-confirmed filterable parameter ──
                    filterable = get_filterable_params_for_category(category, doc)
                    options = [NONE_OPTION] + sorted(filterable.keys())

                    param_choice = forms.SelectFromList.show(
                        options,
                        title="Filter '{}' by parameter".format(category.Name),
                        button_name="Select Parameter"
                    )
                    if param_choice is None:
                        continue

                    use_category_only = (param_choice == NONE_OPTION)

                    if use_category_only:
                        hide_category_in_view(uidoc.ActiveView, category, doc)
                        continue

                    param_id = filterable[param_choice]
                    param, host = get_param_on_host(param_id, element, element_type)
                    if param is None:
                        forms.alert(
                            "Couldn't read '{}' from the selected element.".format(
                                param_choice),
                            title="Error")
                        continue

                    storage_type = param.StorageType
                    if storage_type == DB.StorageType.String:
                        ops = STRING_OPS
                    elif storage_type in (DB.StorageType.Integer, DB.StorageType.Double):
                        ops = NUMERIC_OPS
                    elif storage_type == DB.StorageType.ElementId:
                        ops = ELEMENTID_OPS
                    else:
                        forms.alert("Unsupported parameter type.", title="Error")
                        continue

                    # ── Step 2: pick a condition, scoped to this parameter's type ──
                    operator = forms.SelectFromList.show(
                        ops,
                        title="Condition for '{}'".format(param_choice),
                        button_name="Select Condition"
                    )
                    if operator is None:
                        continue

                    # ── Step 3: value (skipped for has/has-no-value) ──
                    value, display_value = get_filter_value(
                        param, storage_type, operator,
                        param_id=param_id, category=category, doc=doc)
                    if value == "SKIP":
                        continue

                    rule, inverted = build_filter_rule(
                        param_id, storage_type, operator, value)
                    if rule is None:
                        forms.alert("Could not build a filter rule for that choice.",
                                    title="Error")
                        continue

                    element_filter = DB.ElementParameterFilter(rule, inverted)

                    val_label = " '{}'".format(display_value) if display_value else ""
                    filter_name = FILTER_NAME_FORMAT.format(
                        category.Name, param_choice, operator, val_label)

                    existing = find_existing_filter(filter_name, doc)
                    if existing:
                        opt = forms.alert(
                            "Filter '{}' already exists.\n\n"
                            "What do you want to do?".format(filter_name),
                            options=["Use Existing", "Create New", "Skip"],
                            title="Filter Already Exists"
                        )
                        if opt == "Skip":
                            continue
                        elif opt == "Use Existing":
                            apply_filter_to_target(existing, doc, revit)
                            continue
                        elif opt == "Create New":
                            filter_name = get_unique_filter_name(filter_name, doc)

                    with revit.Transaction("Create Annotation Filter"):
                        param_filter = create_param_filter(
                            filter_name, category, element_filter, doc)

                        if param_filter:
                            apply_filter_to_target(param_filter, doc, revit)
                        else:
                            forms.alert("Failed to create filter.", title="Error")

                except Exception as ex:
                    ex_str = str(ex).lower()
                    if any(w in ex_str for w in
                           ("cancelled", "aborted", "operation was cancelled")):
                        break
                    logger.error("Error: {}".format(str(ex)))
                    forms.alert(
                        "An error occurred:\n\n{}\n\nCheck the pyRevit console.".format(
                            str(ex)),
                        title="Error"
                    )
                    break

    except Exception as ex:
        ex_str = str(ex).lower()
        if any(w in ex_str for w in ("cancelled", "aborted")):
            return
        logger.error("Error: {}".format(str(ex)))
        forms.alert("Error occurred. Check pyRevit console.", title="Error")

if __name__ == "__main__":
    main()
