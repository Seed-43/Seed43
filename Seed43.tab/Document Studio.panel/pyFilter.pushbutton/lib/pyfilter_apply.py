# -*- coding: utf-8 -*-
# pyfilter_apply.py
# Seed43 Filter Manager - Apply logic
# pylint: disable=import-error,invalid-name,broad-except

from System import Int64
from System.Collections.Generic import List
from pyrevit import revit, DB, forms
from Snippets._revisions import safe_str

doc = revit.doc

SUPPORTED_VIEW_TYPES = {
    DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan,
    DB.ViewType.Elevation,  DB.ViewType.Section,
    DB.ViewType.Detail,     DB.ViewType.ThreeD,
    DB.ViewType.DraftingView, DB.ViewType.EngineeringPlan,
    DB.ViewType.AreaPlan,   DB.ViewType.Walkthrough,
}

# ── HELPERS ───────────────────────────────────────────────────────────────────

def element_ids_equal(a, b):
    try:
        return a.Equals(b)
    except Exception:
        return False


def element_id_value(eid):
    if hasattr(eid, "Value"):
        return eid.Value
    if hasattr(eid, "IntegerValue"):
        return eid.IntegerValue
    return None


def build_category_id_list(cat_ids_raw):
    cat_lookup = {}

    def _add_cat(cat):
        try:
            val = element_id_value(cat.Id)
            if val is not None:
                cat_lookup[int(val)] = cat.Id
        except Exception:
            pass
        # Also traverse subcategories
        try:
            for sub in cat.SubCategories:
                _add_cat(sub)
        except Exception:
            pass

    try:
        for cat in doc.Settings.Categories:
            _add_cat(cat)
    except Exception:
        pass

    result = List[DB.ElementId]()
    for cid in cat_ids_raw:
        try:
            int_val = int(str(cid))
        except Exception:
            continue
        if int_val in cat_lookup:
            result.Add(cat_lookup[int_val])
        else:
            # Directly construct the ElementId — built-in categories always
            # have valid negative IDs regardless of document contents.
            try:
                result.Add(DB.ElementId(Int64(int_val)))
            except Exception:
                pass
    return result


def str_to_element_id(val):
    if val is None:
        return None
    try:
        return DB.ElementId(Int64(int(val)))
    except Exception:
        return None


def hex_to_color(hex_str):
    if not hex_str or not hex_str.startswith("#"):
        return None
    try:
        r = int(hex_str[1:3], 16)
        g = int(hex_str[3:5], 16)
        b = int(hex_str[5:7], 16)
        return DB.Color(r, g, b)
    except Exception:
        return None


def apply_color(ogs, setter, val):
    c = hex_to_color(val)
    if c:
        try:
            getattr(ogs, setter)(c)
        except Exception:
            pass


def _resolve_pattern_id(val, is_fill=True):
    """Resolve a pattern value to an ElementId in the current document.
    For fill patterns, only returns Drafting patterns — Revit rejects Model
    patterns in view filter overrides with 'Fill pattern must be a drafting pattern'.
    """
    if val is None:
        return None

    name_hint = None
    if isinstance(val, dict):
        name_hint = val.get("name")
        val = val.get("id")

    if val is None:
        return None

    try:
        int_val = int(str(val))
    except Exception:
        return None

    if int_val == -1:
        return None

    eid = DB.ElementId(Int64(int_val))

    def _is_valid_fill(el):
        """True if el is a FillPatternElement with a Drafting pattern."""
        if not isinstance(el, DB.FillPatternElement):
            return False
        try:
            return el.GetFillPattern().Target == DB.FillPatternTarget.Drafting
        except Exception:
            return True  # if Target enum unavailable, allow it

    # Try the stored ID in this document first
    try:
        el = doc.GetElement(eid)
        if el is not None:
            if is_fill:
                if _is_valid_fill(el):
                    return eid
                # ID found but it's a Model pattern — fall through to name lookup
            else:
                if isinstance(el, DB.LinePatternElement):
                    return eid
    except Exception:
        pass

    # ID not found or wrong type — try name lookup
    if name_hint:
        try:
            if is_fill:
                for el in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
                    try:
                        if safe_str(el.Name) == name_hint and _is_valid_fill(el):
                            return el.Id
                    except Exception:
                        pass
            else:
                for el in DB.FilteredElementCollector(doc).OfClass(DB.LinePatternElement):
                    try:
                        if safe_str(el.Name) == name_hint:
                            return el.Id
                    except Exception:
                        pass
        except Exception:
            pass

    return None


def _get_solid_fill_id():
    """Find the solid fill Drafting pattern ElementId in the current document.
    Revit requires Drafting patterns for view filter overrides — Model patterns
    throw 'Fill pattern must be a drafting pattern'."""
    try:
        for fp in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
            pat = fp.GetFillPattern()
            if pat.IsSolidFill and pat.Target == DB.FillPatternTarget.Drafting:
                return fp.Id
    except Exception:
        pass
    # Fallback: any solid fill (in case Target enum differs by version)
    try:
        for fp in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
            if fp.GetFillPattern().IsSolidFill:
                return fp.Id
    except Exception:
        pass
    return None


def apply_eid(ogs, setter, val):
    # Determine if this is a fill or line pattern setter
    is_fill = "Surface" in setter or "Cut" in setter
    eid = _resolve_pattern_id(val, is_fill=is_fill)
    if eid is None:
        return
    try:
        getattr(ogs, setter)(eid)
    except Exception:
        pass


def apply_color_with_solid_fill(ogs, color_setter, pattern_setter, color_val):
    """Apply a surface/cut colour override and ensure a valid solid fill
    pattern is set alongside it so Revit renders the colour."""
    c = hex_to_color(color_val)
    if not c:
        return
    try:
        getattr(ogs, color_setter)(c)
    except Exception:
        pass
    # Always set the solid fill pattern when a colour override is present.
    # Without a pattern Revit silently ignores the colour.
    try:
        solid_id = _get_solid_fill_id()
        if solid_id:
            getattr(ogs, pattern_setter)(solid_id)
    except Exception:
        pass


def apply_bool(ogs, setter, val):
    if val is not None:
        try:
            getattr(ogs, setter)(bool(val))
        except Exception:
            pass


def deserialise_overrides(s):
    ogs = DB.OverrideGraphicSettings()
    if not s:
        return ogs

    apply_color(ogs, "SetProjectionLineColor",     s.get("proj_line_color"))
    apply_eid(ogs,   "SetProjectionLinePatternId", s.get("proj_line_pattern_id"))
    w = s.get("proj_line_weight")
    if w is not None and w != -1:
        try: ogs.SetProjectionLineWeight(int(w))
        except Exception: pass

    apply_color(ogs, "SetCutLineColor",     s.get("cut_line_color"))
    apply_eid(ogs,   "SetCutLinePatternId", s.get("cut_line_pattern_id"))
    w = s.get("cut_line_weight")
    if w is not None and w != -1:
        try: ogs.SetCutLineWeight(int(w))
        except Exception: pass

    # Surface and cut fill colours need a valid fill pattern in 2026.
    # Order: colour first, then try stored pattern ID, then solid-fill fallback.
    apply_color_with_solid_fill(ogs,
        "SetSurfaceForegroundPatternColor",
        "SetSurfaceForegroundPatternId",
        s.get("surf_fg_color"))
    apply_eid(ogs,   "SetSurfaceForegroundPatternId",    s.get("surf_fg_pat"))
    apply_bool(ogs,  "SetSurfaceForegroundPatternVisible", s.get("surf_fg_visible"))

    apply_color_with_solid_fill(ogs,
        "SetSurfaceBackgroundPatternColor",
        "SetSurfaceBackgroundPatternId",
        s.get("surf_bg_color"))
    apply_eid(ogs,   "SetSurfaceBackgroundPatternId",    s.get("surf_bg_pat"))
    apply_bool(ogs,  "SetSurfaceBackgroundPatternVisible", s.get("surf_bg_visible"))

    apply_color_with_solid_fill(ogs,
        "SetCutForegroundPatternColor",
        "SetCutForegroundPatternId",
        s.get("cut_fg_color"))
    apply_eid(ogs,   "SetCutForegroundPatternId",    s.get("cut_fg_pat"))
    apply_bool(ogs,  "SetCutForegroundPatternVisible", s.get("cut_fg_visible"))

    apply_color_with_solid_fill(ogs,
        "SetCutBackgroundPatternColor",
        "SetCutBackgroundPatternId",
        s.get("cut_bg_color"))
    apply_eid(ogs,   "SetCutBackgroundPatternId",    s.get("cut_bg_pat"))
    apply_bool(ogs,  "SetCutBackgroundPatternVisible", s.get("cut_bg_visible"))

    t = s.get("transparency")
    if t is not None:
        try: ogs.SetSurfaceTransparency(int(t))
        except Exception:
            try: ogs.Transparency = int(t)
            except Exception: pass

    h = s.get("halftone")
    if h is not None:
        try: ogs.SetHalftone(bool(h))
        except Exception:
            try: ogs.Halftone = bool(h)
            except Exception: pass

    return ogs


def get_live_views():
    result = {}
    for v in DB.FilteredElementCollector(doc).OfClass(DB.View):
        if v.IsTemplate:
            result["[Template] " + safe_str(v.Name)] = v
            continue
        try:
            if v.ViewType in SUPPORTED_VIEW_TYPES:
                result[safe_str(v.Name)] = v
        except Exception:
            continue
    return result


def apply_filter_to_view(fid, settings, v, warnings):
    label = safe_str(v.Name)
    try:
        existing = list(v.GetFilters())
        if not any(element_ids_equal(eid, fid) for eid in existing):
            v.AddFilter(fid)
    except Exception as ex:
        warnings.append("Add filter to '{}': {}".format(label, type(ex).__name__))
        return False

    enabled = settings.get("enabled")
    if enabled is not None:
        try: v.SetIsFilterEnabled(fid, bool(enabled))
        except Exception: pass

    visible = settings.get("visible")
    if visible is not None:
        try: v.SetFilterVisibility(fid, bool(visible))
        except Exception: pass

    try:
        v.SetFilterOverrides(fid, deserialise_overrides(settings))
    except Exception as ex:
        warnings.append("Set overrides on '{}': {}".format(label, type(ex).__name__))

    return True

# ── MAIN APPLY FLOW ───────────────────────────────────────────────────────────

def run_apply(template, status_callback=None):
    """
    Apply a loaded template dict to user-selected views.
    status_callback(msg) is called with progress updates if provided.
    Returns a summary string.
    """
    def status(msg):
        if status_callback:
            status_callback(msg)

    tpl_filters = template.get("filters", [])
    if not tpl_filters:
        return "Template contains no filters."

    live_views = get_live_views()
    chosen_views = forms.SelectFromList.show(
        sorted(live_views.keys()),
        title="Apply '{}' to Views".format(template.get("name", "Template")),
        button_name="Apply",
        multiselect=True,
    )
    if not chosen_views:
        return None

    target_views = [live_views[n] for n in chosen_views]
    live_filters = {
        f.Name: f
        for f in DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement)
    }

    to_create    = []
    to_skip      = set()

    for fdata in tpl_filters:
        name = fdata["name"]
        if name not in live_filters:
            to_create.append(fdata)
        else:
            answer = forms.alert(
                "Filter '{}' already exists.\n\nOverwrite, use existing, or skip?".format(name),
                title="Filter Exists",
                warn_icon=True,
                options=["Overwrite", "Use Existing", "Skip"]
            )
            if answer == "Overwrite":
                overwrite_entry = dict(fdata)
                overwrite_entry["_overwrite"] = True
                to_create.append(overwrite_entry)
            elif answer == "Use Existing":
                pass  # keep existing filter, still apply settings
            else:
                to_skip.add(name)

    created  = 0
    applied  = 0
    warnings = []

    with revit.Transaction("Apply Filter Template: {}".format(
            template.get("name", ""))):

        # Delete overwrite targets
        for fdata in to_create:
            if fdata.get("_overwrite") and fdata["name"] in live_filters:
                try:
                    doc.Delete(live_filters[fdata["name"]].Id)
                except Exception:
                    warnings.append("{} - could not delete for overwrite".format(fdata["name"]))
                    to_skip.add(fdata["name"])

        live_filters = {
            f.Name: f
            for f in DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement)
        }

        # Create missing
        for fdata in to_create:
            name = fdata["name"]
            if name in to_skip:
                continue
            defn = fdata.get("definition", {})
            cats = build_category_id_list(defn.get("cats", []))
            if cats.Count == 0:
                warnings.append("{} - no category data".format(name))
                to_skip.add(name)
                continue
            try:
                DB.ParameterFilterElement.Create(doc, name, cats)
                created += 1
                status("Created filter: {}".format(name))
            except Exception as ex:
                warnings.append("{} - create failed: {}".format(name, type(ex).__name__))
                to_skip.add(name)

        live_filters = {
            f.Name: f
            for f in DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement)
        }

        # Apply to views
        for fdata in tpl_filters:
            name = fdata["name"]
            if name in to_skip:
                continue
            f = live_filters.get(name)
            if not f:
                warnings.append("{} - not found after create".format(name))
                continue
            settings = fdata.get("settings", {})
            for v in target_views:
                if apply_filter_to_view(f.Id, settings, v, warnings):
                    applied += 1

    lines = ["Template '{}' applied.".format(template.get("name", ""))]
    if created:
        lines.append("Created {} new filter{}.".format(created, "s" if created != 1 else ""))
    lines.append("Applied to {} assignment{}.".format(applied, "s" if applied != 1 else ""))
    if to_skip:
        lines.append("Skipped: {}".format(", ".join(to_skip)))
    if warnings:
        lines.append("\nWarnings:\n" + "\n".join(warnings))
    lines.append("\nNote: filter rules must be reconfigured manually.")
    return "\n".join(lines)
