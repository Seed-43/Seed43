# -*- coding: utf-8 -*-
"""Reversible blanket overrides of a view's filter checkboxes.

The Filters tab of Visibility/Graphics has two checkbox columns per filter -
Enable Filter and Visibility - and no way to switch a whole view's worth off
and back on again without losing which ones the user had already unticked.
This module does that round trip.

WHAT IS REMEMBERED
    Every applied filter's current value for the column being overridden,
    not just the ones that get changed. Restoring then writes each recorded
    value straight back, so a filter the user had already unticked is put
    back unticked rather than being swept up in a blanket re-tick. That is
    the whole point of the tool and the reason a record exists at all.

WHAT IT TOUCHES
    The active view, only, even when a view template controls its Filters
    setting. Revit's UI greys that tab out, but SetIsFilterEnabled and
    SetFilterVisibility write straight to the view and the value survives
    the transaction with the template untouched and every other view using
    it unaffected. Reapplying the template does clear the override, which
    is worth telling the user and is not worth guarding against.

WHERE THE RECORD LIVES
    An Extensible Storage entity on the view itself, holding one JSON
    string, for the same reasons as _dimoverrides.py: delete the view and
    the record goes with it, every read starts from a live element so a
    record cannot dangle, and it survives save and reopen where a session
    variable would not.

    The two columns are recorded independently under separate keys, so
    "filters off" and "filters shown" can be active at the same time and
    reset in either order.

THE ON-SCREEN WARNING
    A blanket override is invisible - the view just looks wrong - so it is
    flagged the way Revit flags its own temporary states, with a coloured
    frame round the view and a caption in the corner.

    Revit will only draw that frame for Temporary View Properties mode, and
    only colours it when a custom title is set: setting CustomTitle and
    CustomColor on their own does nothing, IsCustomized() stays False.
    EnableTemporaryViewPropertiesMode(view.Id) turns the mode on without
    applying any template, which is what makes it usable as a pure notice.

    The catch is that changes made WHILE the mode is on are discarded when
    it is turned off. Filter writes therefore have to happen with the notice
    down - hence notice_off(), then the work, then notice_on() - and callers
    must keep that order. Anything the user changes in the view while the
    notice is up is Revit's temporary-mode behaviour, not ours.

HOW LOUD THE TOOLS ARE
    The wording the tools show is an explanation of the round trip, not a
    status line, so report() shows it once per column per direction and then
    stays quiet - four popups over the life of the install, recorded in
    .user/ViewFilters/guide_shown.json. Anything that went wrong is shown
    every time and never counts as one of the four. forget_reports() puts
    them all back for anyone who wants the explanation again.

THE SAFETY NET
    An override must never be silently carried into a model somebody else
    opens, and it must never be silently destroyed either. Two hooks, doing
    different jobs:

      On close, the Close command hook offers reset_document() - remove the
      overrides, optionally saving so the file on disk is clean. It is bound
      to the command rather than the doc-closing event because
      DocumentClosing is a Revit pre-event and the API forbids modifying a
      document inside one, so a doc-closing hook could see an override but
      never undo it.

      On open, restore_notices() puts the amber frame back on whatever is
      still overridden. The record survives a save and reopen; the frame,
      being session state, does not. Re-raising it is what makes "leave them
      for now" an honest answer to the close prompt rather than a way to
      lose track of an override.

    Both check overridden_views() first and do nothing at all - no
    transaction, no modified flag - on a document with no override on it.

snippets.yaml entry:
  _viewfilters.py:
    description: >
      Blanket on/off overrides of a view's Enable Filter and Visibility
      checkboxes, with the previous per-filter state recorded into the
      Revit file so the reset restores exactly what was there.
"""

import clr
import json
import os

from Autodesk.Revit.DB import (Color, ElementId, FilteredElementCollector,
                               TemporaryViewMode, View, ViewType)
from Autodesk.Revit.DB.ExtensibleStorage import (AccessLevel, Entity,
                                                 ExtensibleStorageFilter,
                                                 Schema, SchemaBuilder)
from System import Guid, String

from Snippets import _userdata

__all__ = ["ENABLE", "VISIBLE", "KINDS", "KIND_LABEL", "NOTICE_PREFIX",
           "schema", "eid", "supports_filters", "applied_filters",
           "filter_name", "read_flag", "write_flag",
           "load_record", "store_record", "forget_record", "active_kinds",
           "last_error", "apply_override", "restore_override",
           "notice_on", "notice_off", "notice_state",
           "controlling_template", "overridden_views", "overridden_names",
           "restore_notices", "reset_document",
           "toggle", "summarise", "needs_attention", "report_due",
           "mark_reported", "forget_reports", "report"]


# ── SCHEMA ──────────────────────────────────────────────────────────────────

# Fixed for the life of the tool. A schema's field set is frozen once the
# GUID has been used, so the single field is an opaque JSON string and all
# future shape changes happen inside it, keyed off "version".
SCHEMA_GUID = Guid("1c919ee2-7d64-4e6e-84ee-849bc4c7a799")
SCHEMA_NAME = "Seed43ViewFilterOverrides"
VENDOR_ID = "SEED43"
FIELD_NAME = "record"
RECORD_VERSION = 1


def schema():
    """Return the Seed43 filter override schema, registering it if this
    session has not seen it yet.

    Lookup first: a schema GUID can only be built once per Revit session,
    and opening any file that already carries records registers it before
    this tool ever runs.
    """
    found = Schema.Lookup(SCHEMA_GUID)
    if found is not None:
        return found

    builder = SchemaBuilder(SCHEMA_GUID)
    builder.SetSchemaName(SCHEMA_NAME)
    builder.SetVendorId(VENDOR_ID)
    # Public both ways: a later Seed43 version, or a plain audit script,
    # should be able to read and clear these without impersonating us.
    builder.SetReadAccessLevel(AccessLevel.Public)
    builder.SetWriteAccessLevel(AccessLevel.Public)
    builder.AddSimpleField(FIELD_NAME, clr.GetClrType(String))
    return builder.Finish()


# ── THE TWO COLUMNS ─────────────────────────────────────────────────────────

ENABLE = "enabled"
VISIBLE = "visibility"

KINDS = (ENABLE, VISIBLE)

# What every filter is forced to while the override is active.
KIND_TARGET = {ENABLE: False, VISIBLE: True}

# Caption fragment shown in the view frame.
KIND_LABEL = {ENABLE: "FILTERS OFF", VISIBLE: "FILTERS SHOWN"}

# How the tools describe each column to the user.
KIND_COLUMN = {ENABLE: "Enable Filter", VISIBLE: "Visibility"}


# ── ELEMENT IDS ─────────────────────────────────────────────────────────────

def eid(element_id):
    """The integer behind an ElementId, on old and new Revit alike.

    IntegerValue was removed in 2026 and Value did not arrive until 2024, so
    neither name alone spans the versions these tools support.

    The int() is not decoration: Value hands back a .NET Int64, and these
    numbers are used as dict keys against ids that came back out of JSON as
    plain Python ints. Coercing here keeps the two comparable.
    """
    try:
        raw = element_id.Value
    except AttributeError:
        raw = element_id.IntegerValue
    return int(raw)


# NOTE: the same shim exists as _connections.eid. Left duplicated rather than
# imported - a filter module reaching into the steel connections module for a
# five-line version guard is worse coupling than the copy.


# ── FILTER STATE ────────────────────────────────────────────────────────────

def supports_filters(view):
    """True for a view whose filter checkboxes this module can read.

    Sheets, schedules, legends and templates all reach the tool sooner or
    later - a sheet is simply what is open when someone reaches for the
    ribbon - and GetFilters raises on most of them, so the type test comes
    first and the call is guarded anyway.
    """
    if not isinstance(view, View) or view.IsTemplate:
        return False
    if view.ViewType in (ViewType.DrawingSheet, ViewType.Schedule,
                         ViewType.ProjectBrowser, ViewType.SystemBrowser,
                         ViewType.Undefined, ViewType.Internal):
        return False
    try:
        view.GetFilters()
        return True
    except Exception:
        return False


def applied_filters(view):
    """The ElementIds of the filters applied to view, or an empty list."""
    try:
        return list(view.GetFilters())
    except Exception:
        return []


def filter_name(doc, filter_id):
    """A filter's name, or its id as text if the element has gone."""
    try:
        element = doc.GetElement(filter_id)
        if element is not None:
            return element.Name
    except Exception:
        pass
    return u"<id {}>".format(eid(filter_id))


def read_flag(view, kind, filter_id):
    """Read one filter's checkbox for the given column."""
    if kind == ENABLE:
        return view.GetIsFilterEnabled(filter_id)
    return view.GetFilterVisibility(filter_id)


def write_flag(view, kind, filter_id, value):
    """Write one filter's checkbox for the given column."""
    if kind == ENABLE:
        view.SetIsFilterEnabled(filter_id, value)
    else:
        view.SetFilterVisibility(filter_id, value)


# ── RECORD ──────────────────────────────────────────────────────────────────

_LAST_ERROR = [None]


def last_error():
    """Why the most recent store_record failed, or None.

    store_record returns False rather than raising so a partial run can
    still report itself, but a bare False leaves nothing to diagnose. This
    carries the reason out.
    """
    return _LAST_ERROR[0]


def load_record(view):
    """The stored record for view as a dict of kind -> entry list.

    Never raises. A view carrying no entity, an entity written by a future
    version, or a corrupt payload all read as "nothing overridden" - the
    worst case is that the user resets by hand, which is where they were
    before the tool existed.
    """
    try:
        entity = view.GetEntity(schema())
        if entity is None or not entity.IsValid():
            return {}
        try:
            raw = entity.Get[String](FIELD_NAME)
        except TypeError:
            raw = entity.Get(FIELD_NAME)   # see the note in store_record
    except Exception:
        return {}

    try:
        data = json.loads(raw)
    except Exception:
        return {}

    if not isinstance(data, dict):
        return {}
    if data.get("version", 0) > RECORD_VERSION:
        return {}

    clean = {}
    for kind in KINDS:
        entries = data.get(kind)
        if not isinstance(entries, list) or not entries:
            continue
        rows = []
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            try:
                filter_id = int(entry["id"])
            except (KeyError, TypeError, ValueError):
                continue
            rows.append({"id": filter_id,
                         "name": entry.get("name", ""),
                         "was": bool(entry.get("was", True))})
        if rows:
            clean[kind] = rows
    return clean


def store_record(view, record):
    """Stamp view with record, or clear it when record is empty.

    Must be called inside an open transaction. Returns True on success; on
    failure returns False and leaves the reason in last_error().
    """
    if not record:
        return forget_record(view)

    payload = {"version": RECORD_VERSION}
    for kind in KINDS:
        if record.get(kind):
            payload[kind] = record[kind]

    # NOTE: ensure_ascii=False for the same reason as _dimoverrides.py -
    # IronPython 2 makes str and unicode the same .NET String, so CPython's
    # ascii encoder tries to utf-8 decode any text holding U+0080..U+00FF
    # and .NET throws. Filter names carry exactly that, Ø and ² especially.
    text = json.dumps(payload, sort_keys=True, ensure_ascii=False)
    try:
        entity = Entity(schema())
        try:
            entity.Set[String](FIELD_NAME, text)
        except TypeError:
            # Entity.Set has several overloads and IronPython's generic
            # binder can fail to pick Set<T>(string, T) from the explicit
            # form. Letting it infer T from the argument resolves the same
            # call.
            entity.Set(FIELD_NAME, text)
        view.SetEntity(entity)
        return True
    except Exception as ex:
        _LAST_ERROR[0] = u"{}: {}".format(type(ex).__name__, ex)
        return False


def forget_record(view):
    """Drop view's record. Must be called inside an open transaction."""
    try:
        view.DeleteEntity(schema())
        return True
    except Exception:
        return False


def active_kinds(view):
    """Which columns currently have an override recorded on view, in order."""
    record = load_record(view)
    return [kind for kind in KINDS if record.get(kind)]


# ── OVERRIDES ───────────────────────────────────────────────────────────────

def apply_override(view, kind):
    """Force every applied filter's kind checkbox to the override value.

    Returns (entries, changed, failed). entries is the record to store: one
    row per filter INCLUDING the ones already sitting at the override value,
    because a restore that only knew about the ones it changed would have no
    way to leave those alone.

    A filter that cannot be read is skipped entirely rather than recorded
    with a guessed value - an unrecorded filter is left untouched by the
    restore, which is the safe failure.
    """
    entries = []
    changed = 0
    failed = 0
    target = KIND_TARGET[kind]
    doc = view.Document

    for filter_id in applied_filters(view):
        try:
            was = read_flag(view, kind, filter_id)
        except Exception:
            failed += 1
            continue
        entries.append({"id": eid(filter_id),
                        "name": filter_name(doc, filter_id),
                        "was": bool(was)})
        if bool(was) == target:
            continue
        try:
            write_flag(view, kind, filter_id, target)
            changed += 1
        except Exception:
            failed += 1
    return entries, changed, failed


def restore_override(view, kind, entries):
    """Put each recorded filter's kind checkbox back to what it was.

    Returns (restored, gone, failed). gone counts recorded filters that are
    no longer applied to the view; a filter added while the override was up
    has no record and is deliberately left as the user set it.
    """
    wanted = dict((entry["id"], entry) for entry in entries)
    restored = 0
    failed = 0

    for filter_id in applied_filters(view):
        entry = wanted.pop(eid(filter_id), None)
        if entry is None:
            continue
        try:
            if bool(read_flag(view, kind, filter_id)) != entry["was"]:
                write_flag(view, kind, filter_id, entry["was"])
                restored += 1
        except Exception:
            failed += 1
    return restored, len(wanted), failed


# ── NOTICE ──────────────────────────────────────────────────────────────────

NOTICE_PREFIX = u"SEED43"

# Amber. Loud enough to read as a warning beside Revit's own cyan temporary
# hide/isolate and red reveal-hidden frames, without being mistaken for one.
NOTICE_RGB = (255, 176, 0)

_NOTICE_SEPARATOR = u"  |  "


def _modes(view):
    try:
        return view.TemporaryViewModes
    except Exception:
        return None


def notice_state(view):
    """What is drawing the view frame right now.

    Returns "none", "seed43" for a frame this module put up, or "foreign"
    for a Temporary View Properties mode somebody else started - a real
    temporary view template, most likely. Foreign frames are never taken
    over or torn down; the user put it there on purpose.
    """
    modes = _modes(view)
    if modes is None:
        return "none"
    try:
        if not modes.IsModeActive(TemporaryViewMode.TemporaryViewProperties):
            return "none"
        title = modes.CustomTitle or u""
    except Exception:
        return "none"
    return "seed43" if title.startswith(NOTICE_PREFIX) else "foreign"


def notice_on(view, kinds):
    """Raise the amber frame on view, captioned for the active columns.

    Must be called inside an open transaction, and AFTER every filter write
    of the run - see the module docstring. Returns True if the frame is up.
    """
    if not kinds:
        return False
    if notice_state(view) == "foreign":
        return False

    modes = _modes(view)
    if modes is None:
        return False
    try:
        if not view.CanEnableTemporaryViewPropertiesMode():
            return False
        labels = [KIND_LABEL[kind] for kind in KINDS if kind in kinds]
        modes.CustomTitle = NOTICE_PREFIX + _NOTICE_SEPARATOR \
            + u" + ".join(labels)
        modes.CustomColor = Color(*NOTICE_RGB)
        view.EnableTemporaryViewPropertiesMode(view.Id)
        return True
    except Exception:
        return False


def notice_off(view):
    """Take our own frame back down, leaving anyone else's alone.

    Must be called inside an open transaction, and BEFORE any filter write
    of the run, because Revit discards changes made while Temporary View
    Properties mode is active.
    """
    if notice_state(view) != "seed43":
        return False
    modes = _modes(view)
    try:
        view.EnableTemporaryViewPropertiesMode(ElementId.InvalidElementId)
        if modes is not None:
            modes.RemoveCustomization()
        return True
    except Exception:
        return False


# ── RUN ─────────────────────────────────────────────────────────────────────

def controlling_template(doc, view):
    """The view template that owns this view's Filters setting, or None.

    Not a routing decision - the override always goes on the view itself,
    which Revit permits even here. It is only ever a warning: reapplying
    that template puts its own filter settings back and takes the override
    with them.
    """
    try:
        from Snippets import _filters
        if _filters.view_template_controls_filters(view, doc):
            return doc.GetElement(view.ViewTemplateId)
    except Exception:
        pass
    return None


def overridden_views(doc):
    """Every view in doc carrying an override record.

    ExtensibleStorageFilter is a quick filter keyed on the schema GUID, so
    this stays cheap on a model with thousands of views where testing each
    one with GetEntity would not. Cheap matters: the hooks run it on every
    document open and close purely to find out there is nothing to do.
    """
    try:
        found = (FilteredElementCollector(doc)
                 .WherePasses(ExtensibleStorageFilter(SCHEMA_GUID))
                 .ToElements())
    except Exception:
        return []
    return [element for element in found if isinstance(element, View)]


def restore_notices(doc):
    """Put the amber frame back on every view still carrying an override.

    Temporary View Properties mode is session state. The override itself is
    written into the file and survives a save and reopen, but the frame does
    not - which leaves a view silently overridden, the one state these tools
    must never produce. The open hook calls this so the marker outlives the
    file being closed.

    Deliberately touches nothing but the frame: a record the user chose to
    keep is theirs to keep, and re-marking it is not the same as undoing it.

    Returns (raised, pending). pending counts views that still want a frame
    and did not get one, which is the caller's cue to try again inside a
    transaction - see the note in hooks/doc-opened.py about why the first
    attempt is made without one.
    """
    raised = 0
    pending = 0
    for view in overridden_views(doc):
        record = load_record(view)
        kinds = [kind for kind in KINDS if record.get(kind)]
        if not kinds:
            continue
        state = notice_state(view)
        if state != "none":
            continue        # ours already, or somebody else's to leave alone
        if notice_on(view, kinds):
            raised += 1
        else:
            pending += 1
    return raised, pending


def overridden_names(doc, limit=8):
    """The names of views in doc still carrying an override, for a dialog.

    Capped with an "and N more" tail: a model where someone has overridden
    forty views is exactly the model where the list must not fill the
    screen.
    """
    names = []
    for view in overridden_views(doc):
        try:
            names.append(view.Name)
        except Exception:
            continue
    names.sort()
    if len(names) <= limit:
        return names
    return names[:limit] + [u"and {} more".format(len(names) - limit)]


def reset_document(doc):
    """Put every overridden view in doc back, and drop its record.

    The safety net behind the hooks, so a forgotten override cannot become
    a permanent part of the model. Must be called inside an open
    transaction, and only when overridden_views() has already found
    something - opening a transaction on a document nobody has touched
    marks it modified for nothing.

    Returns (views, filters): how many views were reset, and how many
    checkboxes were written.
    """
    views = 0
    filters = 0
    for view in overridden_views(doc):
        record = load_record(view)
        if not record:
            continue
        # Frame down first, for the same reason toggle() does it: anything
        # written while it is up is discarded when it comes down.
        notice_off(view)
        for kind in KINDS:
            if not record.get(kind):
                continue
            restored, _gone, _failed = restore_override(view, kind,
                                                        record[kind])
            filters += restored
        forget_record(view)
        views += 1
    return views, filters


def toggle(view, kind):
    """Apply the kind override to view, or reset it if it is already on.

    Everything happens on this one view: the checkboxes, the record, and
    the frame. Must be called inside an open transaction. The three phases
    run in the order the frame demands: take the frame down, do the filter
    work, put the frame back up for whatever is still overridden.

    A view already sitting in somebody else's Temporary View Properties mode
    is refused with action "blocked". Every filter write would be discarded
    the moment they left that mode, which would look exactly like the tool
    silently doing nothing.

    Returns a dict:
        action   "applied", "restored" or "blocked"
        kind     the column acted on
        total    filters recorded (applied) or looked for (restored)
        changed  checkboxes actually written
        gone     recorded filters no longer applied to the view (restore)
        failed   filters that would not read or write
        stored   False if the record could not be written - see last_error()
        notice   "seed43", "foreign" or "none", the frame state afterwards
        kinds    columns still overridden after this run
    """
    if notice_state(view) == "foreign":
        return {"action": "blocked", "kind": kind, "total": 0, "changed": 0,
                "gone": 0, "failed": 0, "stored": True, "notice": "foreign",
                "kinds": active_kinds(view)}

    record = load_record(view)
    resetting = bool(record.get(kind))

    notice_off(view)

    if resetting:
        restored, gone, failed = restore_override(view, kind, record[kind])
        total = len(record[kind])
        changed = restored
        action = "restored"
        record.pop(kind, None)
    else:
        entries, changed, failed = apply_override(view, kind)
        total = len(entries)
        gone = 0
        action = "applied"
        record[kind] = entries

    stored = store_record(view, record)
    remaining = [name for name in KINDS if record.get(name)]
    notice_on(view, remaining)

    return {"action": action,
            "kind": kind,
            "total": total,
            "changed": changed,
            "gone": gone,
            "failed": failed,
            "stored": stored,
            "notice": notice_state(view),
            "kinds": remaining}


def summarise(result, template_name=None):
    """The user-facing wording for a toggle() result.

    One place rather than two, so the pair of tools cannot drift apart in
    how they describe the same outcome.
    """
    kind = result["kind"]
    column = KIND_COLUMN[kind]

    if result["action"] == "blocked":
        return (u"This view is already in Temporary View Properties mode.\n\n"
                u"Restore its view properties first, otherwise Revit would "
                u"discard the override the moment that mode ends.")

    if result["action"] == "applied":
        if kind == ENABLE:
            head = u"{} of {} filters switched off.".format(
                result["changed"], result["total"])
        else:
            head = u"{} of {} filters made visible.".format(
                result["changed"], result["total"])
        tail = u"Run the tool again to put the {} column back exactly as " \
               u"it was.".format(column)
    else:
        head = u"{} of {} filters put back.".format(
            result["changed"], result["total"])
        tail = u"The {} column is as it was before the override.".format(
            column)

    lines = [head]
    if result["gone"]:
        lines.append(u"{} recorded filters are no longer applied to this "
                     u"view and were skipped.".format(result["gone"]))
    if result["failed"]:
        lines.append(u"{} filters would not read or write and were left "
                     u"alone.".format(result["failed"]))
    if template_name and result["action"] == "applied":
        lines.append(u"This view's filters come from view template \"{}\". "
                     u"The override is on this view alone, so reapplying "
                     u"that template will clear it.".format(template_name))
    if not result["stored"]:
        lines.append(u"WARNING: the previous state could not be recorded, so "
                     u"the reset will not be able to restore it. {}".format(
                         last_error() or u""))
    elif result["kinds"] and result["notice"] != "seed43":
        lines.append(u"This view type will not show the override warning "
                     u"frame, so nothing on screen marks it as overridden.")
    elif result["action"] == "applied":
        lines.append(u"The amber frame marks the view until you reset it, "
                     u"and is redrawn each time the model is reopened. "
                     u"Closing the model offers to clear anything still on.")
    lines.append(u"")
    lines.append(tail)
    return u"\n".join(lines)


# ── FIRST RUN ───────────────────────────────────────────────────────────────

# The wording summarise() produces is a guide to how the pair of tools works,
# not a status line: worth reading the first time each direction runs, and
# noise every time after. This records which of the four combinations of
# column and direction the user has already been shown.
GUIDE_TOOL = "ViewFilters"
GUIDE_FILE = "guide_shown.json"


def _guide_path():
    return _userdata.user_path(GUIDE_TOOL, GUIDE_FILE)


def _guide_key(result):
    return "{}.{}".format(result["kind"], result["action"])


def _load_shown():
    """Which guides have been shown. Never raises; a missing or unreadable
    file simply means the user has not seen any of them."""
    try:
        path = _guide_path()
        if not os.path.isfile(path):
            return {}
        with open(path, "r") as handle:
            data = json.load(handle)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _save_shown(shown):
    """Record which guides have been shown. Never raises - failing to write
    it costs the user one repeated popup, which is not worth an error."""
    try:
        with open(_guide_path(), "w") as handle:
            json.dump(shown, handle, indent=2, sort_keys=True)
        return True
    except Exception:
        return False


def needs_attention(result):
    """True when the run has something the user has to be told about.

    Kept separate from the guide so a warning is never swallowed by "you
    have seen this before", and never spends the guide either - a first run
    that could not record its state has not explained anything.
    """
    if result["action"] == "blocked":
        return True
    if not result["stored"] or result["failed"] or result["gone"]:
        return True
    return bool(result["kinds"]) and result["notice"] != "seed43"


def report_due(result):
    """True if the guide for this column and direction has not been shown."""
    return not _load_shown().get(_guide_key(result))


def mark_reported(result):
    """Record that the guide for this column and direction has been shown."""
    shown = _load_shown()
    shown[_guide_key(result)] = True
    return _save_shown(shown)


def forget_reports():
    """Drop the record of which guides have been shown, so they show again.

    The way back for anyone who wants the explanation a second time, and the
    reason the file is one flat dict rather than four scattered markers.
    """
    try:
        path = _guide_path()
        if os.path.isfile(path):
            os.remove(path)
        return True
    except Exception:
        return False


def report(result, template_name=None):
    """The message to put in front of the user, or None to stay quiet.

    All the popup policy in one place, so the two tools cannot disagree
    about when they nag. Marks the guide as shown on the way out, so a
    caller only has to act on what comes back.
    """
    if needs_attention(result):
        return summarise(result, template_name)
    if not report_due(result):
        return None
    mark_reported(result)
    return summarise(result, template_name) + \
        u"\n\nThis note only appears the first time. From now on the tool " \
        u"runs without a popup."
