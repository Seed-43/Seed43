# -*- coding: utf-8 -*-
"""Dimension override records, stored in the Revit file itself.

Revit drops a dimension's override text when the thing it measures is
adjusted or re-referenced. The element survives, its text does not, and
there is nothing left in the model saying what it used to read. This module
keeps that record.

WHERE THE RECORD LIVES
    An Extensible Storage entity on the Dimension element itself, holding
    one JSON string. Not a project-wide index element, for three reasons:

      - Housekeeping is free. Delete the dimension and the record goes with
        it, so a redundant entry cannot outlive the thing it describes. An
        index would have to be pruned, and pruning is where stale data
        creeps in.
      - The record cannot dangle. Every read starts from a live element, so
        "the id is missing" is not a state that can occur.
      - It rides with the element through copy, paste and group edits, and
        it does not serialise every user through one owned element if a
        model is ever workshared.

WHY THE RECORD IS PER SEGMENT
    A multi-segment dimension - a string of dims - is ONE element with ONE
    ElementId carrying several independent texts. An id alone therefore
    cannot address a text. The record holds a list, one entry per segment
    in index order, and single-segment dimensions are simply a list of one.

    Segment index is only trustworthy while the segment count holds. Adding
    or removing a reference reshuffles the indices, which is the very edit
    that drops overrides in the first place - so the count is stored
    alongside and a mismatch is reported for review rather than restored
    blind.

snippets.yaml entry:
  _dimoverrides.py:
    description: >
      Per-dimension override text recorded into the Revit file as an
      Extensible Storage entity, so a dropped override can be identified
      and put back.
"""

import clr
import json

from Autodesk.Revit.DB import Dimension, FilteredElementCollector, SpotDimension
from Autodesk.Revit.DB.ExtensibleStorage import (AccessLevel, Entity,
                                                 ExtensibleStorageFilter,
                                                 Schema, SchemaBuilder)
from System import Guid, String

__all__ = ["FIELDS", "IN_SYNC", "DROPPED", "CHANGED", "RESEGMENTED",
           "UNTRACKED", "schema", "is_target", "text_carriers", "read_current",
           "write_uniform", "write_indexed", "write_segments", "has_text",
           "summarise",
           "load_record", "store_record", "forget_record", "last_error",
           "tracked_dimensions", "compare"]


# ── SCHEMA ──────────────────────────────────────────────────────────────────

# Fixed for the life of the tool. A schema's field set is frozen once the
# GUID has been used, so the single field is an opaque JSON string and all
# future shape changes happen inside it, keyed off "version".
SCHEMA_GUID = Guid("bbf47014-8ef1-4f8c-b18b-6994f2a4a6e9")
SCHEMA_NAME = "Seed43DimensionOverrides"
VENDOR_ID = "SEED43"
FIELD_NAME = "record"
RECORD_VERSION = 1

FIELDS = ("above", "prefix", "value", "suffix", "below")


def schema():
    """Return the Seed43 override schema, registering it if this session
    has not seen it yet.

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
    # should be able to read and repair these without impersonating us.
    builder.SetReadAccessLevel(AccessLevel.Public)
    builder.SetWriteAccessLevel(AccessLevel.Public)
    builder.AddSimpleField(FIELD_NAME, clr.GetClrType(String))
    return builder.Finish()


# ── DIMENSION ACCESS ────────────────────────────────────────────────────────

def is_target(elem):
    """True for a dimension whose text this module can read and write.

    SpotDimension subclasses Dimension but carries its text in type-driven
    parameters instead, so an isinstance test alone would let a spot
    elevation reach a writer that cannot serve it.
    """
    return isinstance(elem, Dimension) and not isinstance(elem, SpotDimension)


def text_carriers(dim):
    """The objects that actually hold dim's text fields, in index order.

    A multi-segment dimension refuses every one of Above/Below/Prefix/
    Suffix/ValueOverride on the Dimension itself - "Cannot access this
    property if this dimension has more than one segment" - and hands them
    to its DimensionSegments. NumberOfSegments reports 0, not 1, for an
    ordinary single-segment dimension, hence the > 1 test rather than a
    truth test.
    """
    if dim.NumberOfSegments > 1:
        return list(dim.Segments)
    return [dim]


def read_current(dim):
    """dim's text fields as they stand now: one dict per carrier."""
    out = []
    for carrier in text_carriers(dim):
        out.append({"above":  carrier.Above or "",
                    "prefix": carrier.Prefix or "",
                    "value":  carrier.ValueOverride or "",
                    "suffix": carrier.Suffix or "",
                    "below":  carrier.Below or ""})
    return out


def write_uniform(dim, data):
    """Write one set of fields onto every carrier of dim.

    An empty string is not a no-op: it is how an existing override is
    cleared and the measured value handed back, so blanks are written
    through rather than skipped.
    """
    for carrier in text_carriers(dim):
        _write_one(carrier, data)


def write_indexed(dim, indices, data):
    """Write one set of fields onto just the listed carriers of dim.

    This is what makes a single segment of a dimension string writable. The
    API hands text to the segment, not to the dimension, so addressing one
    segment is simply a matter of writing that carrier and leaving its
    neighbours alone - no read-modify-write of the whole string, and no risk
    of an untouched segment being rewritten with a stale value.

    Indices outside the current carrier list are ignored, so an index that
    has gone stale degrades to a partial write rather than an exception.
    """
    carriers = text_carriers(dim)
    for index in indices:
        if 0 <= index < len(carriers):
            _write_one(carriers[index], data)


def write_segments(dim, segments):
    """Write a per-carrier list of field dicts back onto dim.

    Extra recorded segments are ignored and missing ones left alone, so a
    length mismatch degrades to a partial restore instead of an exception.
    Callers that care should test compare() first.
    """
    carriers = text_carriers(dim)
    for index, carrier in enumerate(carriers):
        if index >= len(segments):
            break
        _write_one(carrier, segments[index])


def _write_one(carrier, data):
    carrier.Above = data.get("above", "")
    carrier.Prefix = data.get("prefix", "")
    carrier.ValueOverride = data.get("value", "")
    carrier.Suffix = data.get("suffix", "")
    carrier.Below = data.get("below", "")


def has_text(segments):
    """True if any field of any segment carries text."""
    for seg in segments:
        for name in FIELDS:
            if seg.get(name):
                return True
    return False


def summarise(segments):
    """A short one-line rendering of a segment list, for a report column."""
    parts = []
    for seg in segments:
        bits = [seg.get(name, "") for name in FIELDS]
        text = u" ".join(b for b in bits if b)
        parts.append(text if text else u"(blank)")
    return u" | ".join(parts)


# ── RECORD STORAGE ──────────────────────────────────────────────────────────

def load_record(dim):
    """The stored record for dim as a list of field dicts, or None.

    Never raises. A dimension carrying no entity, an entity written by a
    future version, or a corrupt payload all read as "not tracked" - the
    tracker's job is to report what it can restore, not to fall over on
    what it cannot.
    """
    try:
        entity = dim.GetEntity(schema())
        if entity is None or not entity.IsValid():
            return None
        try:
            raw = entity.Get[String](FIELD_NAME)
        except TypeError:
            raw = entity.Get(FIELD_NAME)   # see the note in store_record
    except Exception:
        return None

    try:
        data = json.loads(raw)
    except Exception:
        return None

    if not isinstance(data, dict):
        return None
    if data.get("version", 0) > RECORD_VERSION:
        return None

    segments = data.get("segments")
    if not isinstance(segments, list) or not segments:
        return None

    clean = []
    for seg in segments:
        if not isinstance(seg, dict):
            return None
        clean.append(dict((name, seg.get(name, "") or "") for name in FIELDS))
    return clean


_LAST_ERROR = [None]


def last_error():
    """Why the most recent store_record failed, or None.

    store_record has to keep returning False rather than raising - a scan
    over thousands of dimensions must not stop dead on one bad element -
    but a silent False that a caller only reports as a count leaves nothing
    to diagnose. This carries the reason out.
    """
    return _LAST_ERROR[0]


def store_record(dim, segments=None):
    """Stamp dim with segments, defaulting to its current text.

    Must be called inside an open transaction. Returns True on success; on
    failure returns False and leaves the reason in last_error().
    """
    if segments is None:
        segments = read_current(dim)

    # NOTE: ensure_ascii=False is required, not cosmetic. IronPython 2 makes
    # str and unicode the same .NET String, so CPython 2's ascii encoder
    # takes a branch meant for byte strings:
    #
    #   if isinstance(s, str) and HAS_UTF8.search(s):   # [\x80-\xff]
    #       s = s.decode('utf-8')
    #
    # isinstance() is True for every IronPython string, so any text holding
    # U+0080..U+00FF gets byte-decoded and .NET throws. Dimension overrides
    # are full of exactly that: Ø for diameter, ² for square metres. The
    # non-ascii encoder does no decoding, and the payload lands in a .NET
    # UTF-16 string field anyway, so escaping bought nothing to begin with.
    payload = json.dumps({"version": RECORD_VERSION, "segments": segments},
                         sort_keys=True, ensure_ascii=False)
    try:
        entity = Entity(schema())
        try:
            entity.Set[String](FIELD_NAME, payload)
        except TypeError:
            # Entity.Set has several overloads and IronPython's generic
            # binder can fail to pick Set<T>(string, T) from the explicit
            # form. Letting it infer T from the argument resolves the same
            # call. Belt and braces: a silent failure here would look
            # exactly like "the tool does nothing".
            entity.Set(FIELD_NAME, payload)
        dim.SetEntity(entity)
        return True
    except Exception as ex:
        _LAST_ERROR[0] = u"{}: {}".format(type(ex).__name__, ex)
        return False


def forget_record(dim):
    """Drop dim's record. Must be called inside an open transaction."""
    try:
        dim.DeleteEntity(schema())
        return True
    except Exception:
        return False


def tracked_dimensions(doc):
    """Every dimension in doc carrying a record.

    ExtensibleStorageFilter is a quick filter, so this stays cheap on a
    model with thousands of dimensions where testing each one with
    GetEntity would not.
    """
    return list(FilteredElementCollector(doc)
                .OfClass(Dimension)
                .WherePasses(ExtensibleStorageFilter(SCHEMA_GUID))
                .ToElements())


# ── COMPARISON ──────────────────────────────────────────────────────────────

IN_SYNC = "in_sync"          # record matches what the dimension reads now
DROPPED = "dropped"          # recorded text is gone, dimension is now blank
CHANGED = "changed"          # dimension carries different text than recorded
RESEGMENTED = "resegmented"  # segment count moved, indices no longer line up
UNTRACKED = "untracked"      # no record on this dimension


def compare(dim, record=None):
    """Classify dim against its record. Returns (status, record, current).

    DROPPED is the state worth acting on: the record held text and the
    dimension now holds none, which is what an adjusted or re-referenced
    dimension looks like after Revit resets it. CHANGED means someone
    deliberately typed something else and is reported but never restored by
    default.
    """
    if record is None:
        record = load_record(dim)
    current = read_current(dim)

    if record is None:
        return UNTRACKED, None, current
    if len(record) != len(current):
        return RESEGMENTED, record, current
    if record == current:
        return IN_SYNC, record, current
    if has_text(record) and not has_text(current):
        return DROPPED, record, current
    return CHANGED, record, current
