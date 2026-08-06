# -*- coding: utf-8 -*-
"""Shared schedule state for pySheets' scheduled printing.

Two very different callers need the same rules, which is why this lives in
lib rather than beside the tool:

  * pySheets.py     - arms, disarms and runs cards from its Schedule tab.
  * startup.py      - fires armed cards from a session-level Idling handler,
                      long after the pySheets window has closed.

A "card" is one saved profile with a schedule attached. Its timing (hour,
minute, repeat mode, weekdays, start date) lives in the profile's own JSON
under a "schedule" key, so it travels with the profile. What lives in the
armed file here is only runtime state: which cards are armed, the document
each was armed against, and when each next comes due.

Deliberately imports nothing from pyrevit: startup.py runs on every Revit
launch and must stay cheap (pyrevit.forms alone costs 2 to 4 seconds).
"""

import os
import json
import os.path as op
from datetime import datetime, timedelta

__all__ = [
    'TS_FMT', 'DATE_FMT', 'GRACE_MINUTES', 'HEARTBEAT_STALE_SECONDS',
    'parse_ts', 'read_armed_file', 'write_armed_file',
    'compute_next_run', 'window_owns', 'due_entries', 'sort_entries',
    'is_stale', 'advance_entry',
]


# ── CONSTANTS ────────────────────────────────────────────────────────────────

TS_FMT = '%Y-%m-%dT%H:%M:%S'    # timestamps in the armed file
DATE_FMT = '%Y-%m-%d'           # start_date on a profile's schedule block

# How late an overdue run may still fire. Past this it is skipped and rolled
# on to its next occurrence - so a card armed for 18:00 doesn't ambush you
# with a print run the next morning because Revit was closed overnight.
GRACE_MINUTES = 60

# An open pySheets window stamps the armed file every 20 seconds to claim the
# cards belonging to its own document. Once the stamp is this stale, the
# background handler takes them back - which also covers the window crashing.
HEARTBEAT_STALE_SECONDS = 60

_EMPTY = {
    'version':       2,
    'grace_minutes': GRACE_MINUTES,
    'heartbeat':     None,
    'heartbeat_doc': None,
    'entries':       [],
}


# ── TIME ─────────────────────────────────────────────────────────────────────

def parse_ts(raw):
    """Parse an armed-file timestamp, or None if it is missing/unusable."""
    if not raw:
        return None
    try:
        return datetime.strptime(raw, TS_FMT)
    except Exception:
        return None


def compute_next_run(block, now=None):
    """Next datetime a schedule block should fire, or None if it never will.

    `block` is a profile's "schedule" dict. Repeat mode needs at least one
    weekday ticked; without one there is no valid run time and None comes
    back, which is what greys the card out in the UI."""
    if not block:
        return None
    hour, minute = block.get('hour'), block.get('minute')
    if hour is None or minute is None:
        return None

    now = now or datetime.now()
    if block.get('repeat_mode') != 'Repeat':
        nxt = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        if nxt <= now:
            nxt += timedelta(days=1)
        return nxt

    days = set(block.get('days') or [])
    if not days:
        return None
    start = now.date()
    raw = block.get('start_date')
    if raw:
        try:
            start = max(datetime.strptime(raw, DATE_FMT).date(), start)
        except Exception:
            pass
    # Eight days covers any weekly pattern, whichever day it starts on.
    for i in range(8):
        d = start + timedelta(days=i)
        cand = datetime(d.year, d.month, d.day, hour, minute)
        if cand.weekday() in days and cand > now:
            return cand
    return None


# ── ARMED FILE ───────────────────────────────────────────────────────────────

def read_armed_file(path):
    """The armed-cards file as a dict, with every key defaulted.

    Returns the empty shape rather than None, so no caller has to guard, and
    upgrades a v1 file (one flat schedule) to the v2 list on the way through
    - a schedule armed before multi-card support is not silently lost."""
    data = None
    try:
        if op.isfile(path):
            with open(path, 'r') as f:
                data = json.load(f)
    except Exception:
        data = None
    if not isinstance(data, dict):
        return dict(_EMPTY, entries=[])

    if 'entries' not in data:
        entry = dict((k, data.get(k)) for k in
                     ('profile_name', 'document_path', 'document_title',
                      'next_run'))
        data = dict(_EMPTY,
                    entries=[entry] if data.get('enabled') else [])
    for key, default in _EMPTY.items():
        data.setdefault(key, default)
    if not isinstance(data.get('entries'), list):
        data['entries'] = []
    return data


def write_armed_file(path, data):
    """Write the armed-cards file. Raises nothing the caller must handle -
    scheduling should never take the UI down over a disk error - so callers
    that care about failure should check the file themselves."""
    try:
        folder = op.dirname(path)
        if folder and not op.isdir(folder):
            os.makedirs(folder)
        with open(path, 'w') as f:
            json.dump(data, f, indent=2)
        return True
    except Exception:
        return False


# ── DUE CARDS ────────────────────────────────────────────────────────────────

def sort_entries(entries):
    """Due time first, then profile name — the order the dropdowns list them
    in, so several cards armed for the same minute run in a predictable
    order rather than whatever the file happened to hold."""
    return sorted(entries, key=lambda e: (e.get('next_run') or '',
                                          (e.get('profile_name') or '').lower()))


def window_owns(data, entry, now=None):
    """True while an open pySheets window has claimed this card's document.

    Ownership is per document, not per file: a window can only run cards for
    the project it was opened on, so cards armed against other projects stay
    with the background handler even while a window is open."""
    now = now or datetime.now()
    beat = parse_ts(data.get('heartbeat'))
    if not beat or (now - beat).total_seconds() > HEARTBEAT_STALE_SECONDS:
        return False
    owned = data.get('heartbeat_doc') or ''
    path = entry.get('document_path') or ''
    return bool(owned) and op.normcase(owned) == op.normcase(path)


def due_entries(data, now=None, doc_path=None, respect_heartbeat=True):
    """Armed cards whose time has come, sorted.

    Includes cards that are past their grace window: the caller decides
    whether to run or retire those, and it has to see them either way or a
    missed card would sit stale forever. Use `is_stale` to tell them apart.

    doc_path          - only cards armed against this document.
    respect_heartbeat - skip cards an open window has claimed."""
    now = now or datetime.now()
    out = []
    for entry in data.get('entries') or []:
        if not entry.get('profile_name') or not entry.get('document_path'):
            continue
        if doc_path and (op.normcase(entry['document_path'])
                         != op.normcase(doc_path)):
            continue
        when = parse_ts(entry.get('next_run'))
        if not when or now < when:
            continue
        if respect_heartbeat and window_owns(data, entry, now):
            continue
        out.append(entry)
    return sort_entries(out)


def is_stale(data, entry, now=None):
    """True when a due card missed its slot by more than the grace window."""
    now = now or datetime.now()
    when = parse_ts(entry.get('next_run'))
    if not when:
        return True
    grace = data.get('grace_minutes') or GRACE_MINUTES
    return (now - when).total_seconds() > grace * 60


def advance_entry(entry, now=None):
    """Roll one armed card on to its next occurrence.

    Returns True if the entry now holds a future run time, False if the card
    is spent (a "Once" that has fired, or a profile that has gone) and should
    be dropped from the list. Reads the timing straight off the profile the
    entry points at, so a card can be retired without opening pySheets - a
    schedule missed while Revit was shut would otherwise sit stale forever,
    past its grace window and never able to fire again."""
    path = entry.get('profile_path')
    block = None
    if path:
        try:
            with open(path, 'r') as f:
                block = json.load(f).get('schedule')
        except Exception:
            block = None
    if not isinstance(block, dict) or block.get('repeat_mode') != 'Repeat':
        return False
    nxt = compute_next_run(block, now)
    if not nxt:
        return False
    entry['next_run'] = nxt.strftime(TS_FMT)
    return True
