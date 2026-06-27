# -*- coding: utf-8 -*-
# view_assign.py
# Seed43 Filter Manager - assign a saved template's filters to live
# Views / View Templates.
# pylint: disable=import-error,invalid-name,broad-except

from pyrevit import revit, DB
from Snippets._revisions import safe_str

import pyfilter_save as fs
import pyfilter_apply as fa

doc = revit.doc

# View types that can carry filters and are worth listing.
SUPPORTED_VIEW_TYPES = {
    DB.ViewType.FloorPlan,  DB.ViewType.CeilingPlan,
    DB.ViewType.Elevation,  DB.ViewType.Section,
    DB.ViewType.Detail,     DB.ViewType.ThreeD,
    DB.ViewType.DraftingView, DB.ViewType.EngineeringPlan,
    DB.ViewType.AreaPlan,
}

# ── HELPERS ───────────────────────────────────────────────────────────────────

def fid_value(eid):
    if hasattr(eid, "Value"):
        return eid.Value
    if hasattr(eid, "IntegerValue"):
        return eid.IntegerValue
    return None


def _overrides_allowed(v):
    try:
        return v.AreGraphicsOverridesAllowed()
    except Exception:
        return True


def _ids_equal(a, b):
    try:
        return a.Equals(b)
    except Exception:
        return False


def live_filter_map():
    return {
        f.Name: f
        for f in DB.FilteredElementCollector(doc)
                   .OfClass(DB.ParameterFilterElement)
    }


def filters_on_view(view):
    try:
        return set(fid_value(fid) for fid in view.GetFilters())
    except Exception:
        return set()

# ── VIEW LISTING ──────────────────────────────────────────────────────────────

def get_live_views(templates):
    """
    Return {display_name: view}. If templates is True, only view templates;
    otherwise only real views of a supported type. Filtered to views that
    actually allow graphic overrides (i.e. can carry filters).
    """
    result = {}
    for v in DB.FilteredElementCollector(doc).OfClass(DB.View):
        try:
            if v.IsTemplate != bool(templates):
                continue
            if not templates and v.ViewType not in SUPPORTED_VIEW_TYPES:
                continue
            if not _overrides_allowed(v):
                continue
            result[safe_str(v.Name)] = v
        except Exception:
            continue
    return result

# ── PER-VIEW STATE ────────────────────────────────────────────────────────────

def default_assignment_for_template(view, template_filters):
    """
    Build the per-view tick state for one view from a saved template's filter
    list:
        { fname: {"assigned": bool, "settings": {...}, "definition": {...}} }

    'assigned' defaults to whether a filter of that name is already present on
    the view, so existing assignments show up ticked. 'settings' and
    'definition' come straight from the template, so applying reuses the
    template's saved override look.
    """
    live        = live_filter_map()
    on_view_ids = filters_on_view(view)
    out = {}
    for fdata in template_filters:
        fname = fdata.get("name")
        if not fname:
            continue
        existing = live.get(fname)
        already  = False
        if existing is not None:
            already = fid_value(existing.Id) in on_view_ids
        out[fname] = {
            "assigned":   already,
            "settings":   dict(fdata.get("settings", {})),
            "definition": fdata.get("definition", {}),
        }
    return out

# ── COMMIT ────────────────────────────────────────────────────────────────────

def commit_assignments(view_state, status_callback=None, additive=False):
    """
    Push the ticked template filters onto each view in a single transaction.

    view_state is keyed by view UniqueId:
        { uid: {"view": <View>, "template": <name>,
                "filters": {fname: {"assigned": bool,
                                    "settings": {...},
                                    "definition": {...}}}} }

    Ticked filters are added (created from their definition first if they do not
    exist in the model) and given the template's saved overrides. When additive
    is False, unticked filters that are present on the view are removed; when
    additive is True they are left untouched. Returns a summary string.
    """
    def status(msg):
        if status_callback:
            status_callback(msg)

    if not view_state:
        return "No views selected. Pick a view on the left and tick filters."

    live_filters = live_filter_map()

    # Collect filters that need creating (ticked anywhere but missing).
    to_create = {}
    for entry in view_state.values():
        for fname, st in entry.get("filters", {}).items():
            if (st.get("assigned")
                    and fname not in live_filters
                    and fname not in to_create):
                to_create[fname] = st.get("definition", {})

    created = added = removed = touched_views = 0
    warnings = []

    with revit.Transaction("Assign Template Filters to Views"):

        # Create any missing filters from their stored category list.
        for fname, defn in to_create.items():
            cats = fa.build_category_id_list(defn.get("cats", []))
            if cats.Count == 0:
                warnings.append("{} - no category data, not created".format(fname))
                continue
            try:
                DB.ParameterFilterElement.Create(doc, fname, cats)
                created += 1
                status("Created filter: {}".format(fname))
            except Exception as ex:
                warnings.append("{} - create failed: {}".format(
                    fname, type(ex).__name__))

        live_filters = live_filter_map()

        # Apply ticks per view.
        for uid, entry in view_state.items():
            view = entry.get("view")
            fmap = entry.get("filters", {})
            if view is None:
                continue

            label = safe_str(view.Name)
            try:
                current_ids = list(view.GetFilters())
            except Exception:
                current_ids = []

            view_changed = False

            for fname, st in fmap.items():
                want = bool(st.get("assigned"))
                if additive and not want:
                    continue  # additive push never removes
                f = live_filters.get(fname)
                if not f:
                    continue
                on_view = any(_ids_equal(eid, f.Id) for eid in current_ids)

                if want and not on_view:
                    try:
                        view.AddFilter(f.Id)
                        added += 1
                        view_changed = True
                    except Exception as ex:
                        warnings.append("Add '{}' to '{}': {}".format(
                            fname, label, type(ex).__name__))
                        continue

                elif not want and on_view:
                    try:
                        view.RemoveFilter(f.Id)
                        removed += 1
                        view_changed = True
                    except Exception as ex:
                        warnings.append("Remove '{}' from '{}': {}".format(
                            fname, label, type(ex).__name__))
                    continue  # nothing more to set on a removed filter

                if want:
                    settings = st.get("settings", {})
                    enabled  = settings.get("enabled")
                    if enabled is not None:
                        try: view.SetIsFilterEnabled(f.Id, bool(enabled))
                        except Exception: pass
                    visible = settings.get("visible")
                    if visible is not None:
                        try: view.SetFilterVisibility(f.Id, bool(visible))
                        except Exception: pass
                    try:
                        view.SetFilterOverrides(
                            f.Id, fa.deserialise_overrides(settings))
                    except Exception as ex:
                        import traceback
                        warnings.append("Overrides '{}' on '{}': {} -- {}".format(
                            fname, label,
                            type(ex).__name__, str(ex)))

            if view_changed:
                touched_views += 1
                status("Updated: {}".format(label))

    lines = ["Template filters applied."]
    if created:
        lines.append("Created {} new filter{}.".format(
            created, "s" if created != 1 else ""))
    lines.append("Views changed: {}".format(touched_views))
    lines.append("Filters added: {}, removed: {}".format(added, removed))
    if warnings:
        lines.append("\nWarnings:\n" + "\n".join(warnings))
    return "\n".join(lines)

# ── PULL (view -> template) ───────────────────────────────────────────────────

def pull_view_filters(view):
    """
    Read the filters currently applied to a view (or view template) and return
    template-ready rows:
        [ {"name", "settings", "definition"}, ... ]
    Settings capture the live VG overrides; definition captures the filter's
    categories so the template can recreate it elsewhere.
    """
    id_to_f = {fid_value(f.Id): f for f in live_filter_map().values()}
    rows = []
    try:
        fids = list(view.GetFilters())
    except Exception:
        fids = []
    for fid in fids:
        f = id_to_f.get(fid_value(fid))
        if not f:
            continue
        try:
            settings = fs.get_filter_settings_from_view(f, view)
        except Exception:
            settings = {}
        try:
            definition = fs.serialise_filter_def(f)
        except Exception:
            definition = {}
        rows.append({
            "name":       f.Name,
            "settings":   settings,
            "definition": definition,
        })
    return rows
