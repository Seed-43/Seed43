# -*- coding: utf-8 -*-
"""
Shared helpers for reading/writing Revit parameters and collecting
categories/schedules. Promoted out of pyTable so any tool touching
parameters, model categories, or schedules can reuse them.
"""

from Autodesk.Revit.DB import FilteredElementCollector, StorageType, ViewSchedule, ElementId

try:
    from pyrevit import script as _script
    _logger = _script.get_logger()
except Exception:
    _logger = None


def _log_info(msg):
    if _logger is not None:
        try:
            _logger.info(msg)
        except Exception:
            pass


# ── PARAMETER VALUES ────────────────────────────────────────────────────────

def get_param_display_value(param):
    """
    Return a parameter's value as display text, matching what a user would
    see/type in a spreadsheet or form.

    AsValueString() applies the parameter's own unit formatting (via its
    ForgeTypeId) and resolves ElementId-backed params to their display name,
    so no separate UnitUtils conversion is needed on the read side. Falls
    back per StorageType where AsValueString() returns None.
    """
    val_str = param.AsValueString()
    if val_str is not None:
        return val_str

    storage = param.StorageType
    if storage == StorageType.String:
        return param.AsString() or u""
    if storage == StorageType.Integer:
        return unicode(param.AsInteger())
    if storage == StorageType.Double:
        return unicode(param.AsDouble())
    return u""


def set_param_from_display_value(param, text):
    """
    Write a display-text value back to a parameter.

    SetValueString() is the primary path: it parses the string as Revit's own
    UI would, so it is unit-aware (a length param handles "3200" vs "3.2 m"
    vs "10' - 6\"" per project units) and resolves named ElementId references.

    Not every StorageType accepts it, so String and Integer fall back to a
    typed Set(). Double deliberately has NO fallback: a raw float(text) Set()
    would write internal (feet-based) units, silently misreading whatever
    unit the display string was in. Raises instead, so the caller can skip it
    rather than corrupt a real dimension.

    Raises ValueError if the value could not be set.
    """
    try:
        if param.SetValueString(text):
            return
    except Exception:
        pass

    storage = param.StorageType
    if storage == StorageType.String:
        param.Set(text)
    elif storage == StorageType.Integer:
        param.Set(int(round(float(text))))
    elif storage == StorageType.Double:
        raise ValueError(
            "Could not set parameter '{0}' - SetValueString failed and no "
            "safe unit-aware fallback exists for a Double value of "
            "'{1}'".format(param.Definition.Name, text))
    else:
        raise ValueError(
            "Could not set parameter '{0}' - unsupported StorageType for "
            "value '{1}'".format(param.Definition.Name, text))


# ── CATEGORIES ──────────────────────────────────────────────────────────────

def get_categories_with_elements(doc):
    """Return every Category in the document that has at least one placed
    (non-type) element, as a list of Category objects."""
    result = []
    seen = set()
    for cat in doc.Settings.Categories:
        if cat is None or cat.Name in seen:
            continue
        collector = FilteredElementCollector(doc).OfCategoryId(cat.Id).WhereElementIsNotElementType()
        if collector.GetElementCount() == 0:
            continue
        seen.add(cat.Name)
        result.append(cat)
    return result


def get_elements_by_category(doc, category_id):
    """Return all instances (not types) in the given category."""
    return list(FilteredElementCollector(doc).OfCategoryId(category_id).WhereElementIsNotElementType())


# ── SCHEDULES ───────────────────────────────────────────────────────────────

def get_schedules(doc):
    """Return every non-template ViewSchedule in the document."""
    schedules = FilteredElementCollector(doc).OfClass(ViewSchedule).WhereElementIsNotElementType()
    return [s for s in schedules if not s.IsTemplate]


# ── DROPDOWN / VALID-VALUE DISCOVERY ────────────────────────────────────────
#
# Every lookup below is wrapped defensively: a wrong guess about an API
# detail costs that one column its dropdown, not the whole export.

# BuiltInParameters backed by a fixed .NET enum with a stable value set.
# BuiltInParameter and enum names differ across Revit versions, so entries
# are registered ONE AT A TIME - a bad guess skips only that entry.
_ENUM_PARAM_TABLE = {}
try:
    from Autodesk.Revit.DB import BuiltInParameter, LabelUtils
    import Autodesk.Revit.DB as _DB
except Exception:
    BuiltInParameter = None
    LabelUtils = None
    _DB = None


def _try_register_enum(bip_name, enum_type_name):
    if BuiltInParameter is None or _DB is None:
        _log_info("pyTable dropdown table: BuiltInParameter/Autodesk.Revit.DB import failed entirely - "
                   "no enum-based dropdowns will work at all.")
        return
    try:
        bip = getattr(BuiltInParameter, bip_name)
    except AttributeError:
        _log_info("pyTable dropdown table: BuiltInParameter.{0} does not exist on this Revit "
                   "version - skipped.".format(bip_name))
        return
    try:
        enum_type = getattr(_DB, enum_type_name)
    except AttributeError:
        _log_info("pyTable dropdown table: Autodesk.Revit.DB.{0} does not exist on this Revit "
                   "version - skipped.".format(enum_type_name))
        return
    _ENUM_PARAM_TABLE[bip] = enum_type
    _log_info("pyTable dropdown table: registered {0} -> {1}".format(bip_name, enum_type_name))


# All confirmed live against a real pyRevit log - the enum names below are
# what Revit actually resolved to, not guesses.
_try_register_enum("VIEW_DETAIL_LEVEL", "ViewDetailLevel")
_try_register_enum("VIEW_DISCIPLINE", "ViewDiscipline")
_try_register_enum("MODEL_GRAPHICS_STYLE", "DisplayStyle")               # "Visual Style" (some contexts)
_try_register_enum("MODEL_GRAPHICS_STYLE_ANON_DRAFT", "DisplayStyle")    # "Visual Style" (views)
_try_register_enum("VIEW_PARTS_VISIBILITY", "PartsVisibility")           # "Parts Visibility"
_try_register_enum("VIEW_MODEL_DISPLAY_MODE", "DisplayModel")            # "Display Model"
_try_register_enum("COLOR_SCHEME_LOCATION", "ColorSchemeLocation")       # "Color Scheme Location"


def get_param_dropdown_options(param, doc):
    """
    Return a list of valid display-string options for a constrained
    parameter, or None if it's free text/number with no fixed set of
    values. Three cases, most to least reliable:

      1. Yes/No parameters - always exactly ["Yes", "No"].
      2. A small hardcoded table of BuiltInParameters backed by a fixed
         .NET enum (Detail Level, View Discipline, ...) via
         LabelUtils.GetLabelFor().
      3. ElementId-storage parameters (Level, Phase, Material, ...) -
         collects every live-document element sharing the category of the
         parameter's CURRENT value, so the list matches what is actually
         pickable. Returns None if the parameter is empty, since there is
         then no category to work from.
    """
    try:
        if _is_yes_no(param):
            return [u"Yes", u"No"]
    except Exception:
        pass

    try:
        options = _get_known_enum_options(param)
        if options:
            return options
    except Exception:
        pass

    try:
        if param.StorageType == StorageType.ElementId:
            return _get_element_id_options(param, doc)
    except Exception:
        pass

    return None


def _is_yes_no(param):
    if param.StorageType != StorageType.Integer:
        return False
    try:
        from Autodesk.Revit.DB import SpecTypeId
        return param.Definition.GetDataType() == SpecTypeId.Boolean.YesNo
    except Exception:
        return False


def _get_built_in_parameter(param):
    try:
        from Autodesk.Revit.DB import InternalDefinition
        definition = param.Definition
        if isinstance(definition, InternalDefinition):
            return definition.BuiltInParameter
        _log_info("pyTable dropdown: '{0}'.Definition is not an InternalDefinition ({1}) - "
                   "can't check the enum table for it.".format(
                       getattr(definition, "Name", "?"), type(definition).__name__))
    except Exception as ex:
        _log_info("pyTable dropdown: _get_built_in_parameter failed: {0}".format(ex))
    return None


_CAMEL_SPLIT_RE = None
try:
    import re as _re
    _CAMEL_SPLIT_RE = _re.compile(r'([a-z0-9])([A-Z])')
except Exception:
    pass

_SENTINEL_MEMBER_NAMES = set([u"undefined", u"invalid", u"none", u"unknown"])


def _camel_to_title(name):
    """'ShowPartsOnly' -> 'Show Parts Only'. Falls back to the raw name if
    the regex module is somehow unavailable."""
    if _CAMEL_SPLIT_RE is None:
        return name
    return _CAMEL_SPLIT_RE.sub(r'\1 \2', name)


def _get_known_enum_options(param):
    name = getattr(param.Definition, "Name", "?")
    bip = _get_built_in_parameter(param)
    if bip is None:
        _log_info("pyTable dropdown '{0}': not resolvable to a BuiltInParameter.".format(name))
        return None
    if bip not in _ENUM_PARAM_TABLE:
        _log_info("pyTable dropdown '{0}': BuiltInParameter is {1}, not in the registered "
                   "enum table.".format(name, bip))
        return None
    enum_type = _ENUM_PARAM_TABLE[bip]
    from System import Enum
    options = []
    try:
        for member in Enum.GetValues(enum_type):
            member_name = unicode(member)
            if member_name.lower() in _SENTINEL_MEMBER_NAMES:
                continue  # e.g. an "Undefined"/"Invalid" placeholder member, not a real choice
            label = None
            if LabelUtils is not None:
                try:
                    label = LabelUtils.GetLabelFor(member)
                except Exception:
                    # GetLabelFor is BuiltInCategory-only, so most enums land
                    # here - expected, not worth logging.
                    pass
            options.append(label if label else _camel_to_title(member_name))
    except Exception as ex:
        _log_info("pyTable dropdown '{0}': found in table ({1}), but Enum.GetValues "
                   "failed: {2}".format(name, enum_type, ex))
        return None
    if not options:
        _log_info("pyTable dropdown '{0}': enum table entry found but produced zero options.".format(name))
    return options or None


def _get_element_id_options(param, doc):
    current_id = param.AsElementId()
    if current_id is None or current_id == ElementId.InvalidElementId:
        return None
    current_el = doc.GetElement(current_id)
    if current_el is None or current_el.Category is None:
        return None

    category_id = current_el.Category.Id
    names = set()
    for el in FilteredElementCollector(doc).OfCategoryId(category_id):
        try:
            name = el.Name
        except Exception:
            continue
        if name:
            names.add(name)
    return sorted(names) if names else None
