# -*- coding: utf-8 -*-
# "Batch Upgrade / schedule_io"
# "Seed43"
# """
# The Batch-Upgrade side of one-time scheduled runs. Reuses the shared
# engine in lib/Snippets/_schedule.py - the same one pySheets' scheduler
# uses - rather than reinventing due-time/grace-window logic, but keeps
# its own separate armed file: Batch Upgrade entries aren't tied to an
# open document the way pySheets' are, so they don't belong mixed in
# with pySheets' own scheduled_print.json.
#
# Only ever one entry armed at a time - there's no "manage several"
# schedule UI here, just one run, one time. Arming a new one replaces
# whatever was armed before.
#
# Batch Upgrade only ever arms a "run once" entry - see startup.py's
# Batch Upgrade schedule handler, which drops an entry unconditionally
# once it's due (fired or missed its grace window), so there's no
# repeat-specific logic needed on this side at all.
# """

# ── IMPORTS ────────────────────────────────────────────────────────────────

import os
import sys
import json

__all__ = ["schedule_path", "schedule_mod", "read_schedule",
           "write_schedule", "armed_entry", "arm", "disarm",
           "job_snapshot_path", "write_snapshot", "read_snapshot",
           "clear_snapshot"]


# ── CONSTANTS ──────────────────────────────────────────────────────────────

# tools/ -> Batch Upgrade.pushbutton -> Seed43.panel -> Seed43.tab ->
# Seed43.extension -> lib. Same walk as job_io.py/templates.py, extended
# one level further to reach the shared Snippets package.
_TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))
_PUSHBUTTON_DIR = os.path.dirname(_TOOLS_DIR)
_PANEL_DIR = os.path.dirname(_PUSHBUTTON_DIR)
_TAB_DIR = os.path.dirname(_PANEL_DIR)
_EXTENSION_DIR = os.path.dirname(_TAB_DIR)
_LIB_DIR = os.path.join(_EXTENSION_DIR, "lib")

_SCHEDULE_FILE = os.path.join(
    _EXTENSION_DIR, ".user", "BatchUpgrade", "settings", "scheduled_run.json")

# The frozen settings a scheduled run executes. Written when the schedule is
# armed, deleted when it fires or is cancelled.
_JOB_SNAPSHOT = os.path.join(
    _EXTENSION_DIR, ".user", "BatchUpgrade", "settings", "scheduled_job.json")

_JOB_LABEL = "Scheduled batch upgrade"

# Not a real document path - Batch Upgrade jobs aren't tied to one, but
# Snippets._schedule.due_entries requires a truthy document_path on every
# entry (it doubles as the "is this row even filled in" guard), so this
# is a fixed placeholder rather than an actual file path.
_PSEUDO_DOC = "batch-upgrade"


# ── CORE LOGIC ─────────────────────────────────────────────────────────────

def schedule_mod():
    """The shared schedule engine, imported defensively.

    This file can be loaded from startup.py's headless schedule path,
    where lib/ isn't guaranteed to already be on sys.path (unlike a
    normal pyRevit script invocation, which pyRevit sets up for you) -
    same reasoning as startup.py's own _schedule_mod().
    """
    try:
        if os.path.isdir(_LIB_DIR) and _LIB_DIR not in sys.path:
            sys.path.insert(0, _LIB_DIR)
        from Snippets import _schedule
        return _schedule
    except Exception:
        return None


def schedule_path():
    return _SCHEDULE_FILE


def read_schedule():
    sched = schedule_mod()
    if not sched:
        return None
    return sched.read_armed_file(_SCHEDULE_FILE)


def write_schedule(data):
    sched = schedule_mod()
    if not sched:
        return False
    return sched.write_armed_file(_SCHEDULE_FILE, data)


def armed_entry():
    """The single currently-armed entry, or None if nothing is armed."""
    data = read_schedule()
    if not data:
        return None
    entries = data.get("entries") or []
    return entries[0] if entries else None


def job_snapshot_path():
    """The one-off job a scheduled run executes.

    A batch upgrade is a one-shot thing, so arming snapshots the window's
    current settings straight to this file rather than pointing at a saved
    template. Nothing else reads it and it's deleted once the run is over.
    """
    return _JOB_SNAPSHOT


def write_snapshot(files, out_dir, targets, audit=False, compact=True):
    """Freeze the current settings as the job a scheduled run will execute."""
    folder = os.path.dirname(_JOB_SNAPSHOT)
    if folder and not os.path.isdir(folder):
        try:
            os.makedirs(folder)
        except OSError:
            return False
    payload = {
        "files": files,
        "out_dir": out_dir,
        "targets": sorted(targets),
        "audit": bool(audit),
        "compact": bool(compact),
    }
    try:
        with open(_JOB_SNAPSHOT, "w") as handle:
            json.dump(payload, handle, indent=2)
        return True
    except Exception:
        return False


def read_snapshot(path=None):
    """Read a scheduled run's frozen settings, or None if unreadable."""
    path = path or _JOB_SNAPSHOT
    if not os.path.isfile(path):
        return None
    try:
        with open(path, "r") as handle:
            return json.load(handle)
    except Exception:
        return None


def clear_snapshot(path=None):
    """Delete the one-off job file. Safe to call when it isn't there."""
    path = path or _JOB_SNAPSHOT
    try:
        if os.path.isfile(path):
            os.remove(path)
    except OSError:
        pass


def arm(when, files, out_dir, targets, audit=False, compact=True):
    """Arm a one-time run of the given settings at datetime `when`.

    Replaces any existing armed entry and its snapshot.
    """
    sched = schedule_mod()
    if not sched:
        return False
    if not write_snapshot(files, out_dir, targets, audit, compact):
        return False
    entry = {
        # profile_name/profile_path are the shared engine's field names; here
        # they describe the frozen job rather than a saved template.
        "profile_name": _JOB_LABEL,
        "profile_path": _JOB_SNAPSHOT,
        "document_path": _PSEUDO_DOC,
        "document_title": _JOB_LABEL,
        "next_run": when.strftime(sched.TS_FMT),
    }
    data = read_schedule() or {"version": 2, "entries": []}
    data["entries"] = [entry]
    return write_schedule(data)


def disarm():
    """Cancel whatever is currently armed, and bin its snapshot with it."""
    clear_snapshot()
    data = read_schedule() or {"version": 2, "entries": []}
    data["entries"] = []
    return write_schedule(data)
