# -*- coding: utf-8 -*-
# script.py
from pyrevit import revit, DB, forms, script
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List

doc   = revit.doc
uidoc = revit.uidoc

# ── PICK TEXT NOTES ───────────────────────────────────────────────────────────

try:
    with forms.WarningBar(title="Select Text Notes to split (ESC to Cancel):"):
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element, "Select text notes to split")
except Exception:
    script.exit()

text_notes = []
for ref in refs:
    el = doc.GetElement(ref.ElementId)
    if isinstance(el, DB.TextNote):
        text_notes.append(el)

if not text_notes:
    forms.alert("No TextNotes selected.", title="Split Text", exitscript=True)

# ── FORMAT RUN EXTRACTION ─────────────────────────────────────────────────────
# Read formatting character by character and collapse into runs.
# Each run is a dict: {start, length, bold, italic, underline, allcaps}

def get_format_runs(formatted_text, text_length):
    runs   = []
    if text_length == 0:
        return runs

    def get_flags(i):
        r = DB.TextRange(i, 1)
        # FormatStatus: 0=None, 1=True, -1=Mixed
        # Use int() to avoid IronPython 2 keyword clash (.True is reserved)
        return (
            int(formatted_text.GetBoldStatus(r))      == 1,
            int(formatted_text.GetItalicStatus(r))    == 1,
            int(formatted_text.GetUnderlineStatus(r)) == 1,
            int(formatted_text.GetAllCapsStatus(r))   == 1,
        )

    prev_flags = get_flags(0)
    run_start  = 0

    for i in range(1, text_length):
        flags = get_flags(i)
        if flags != prev_flags:
            runs.append({
                "start":   run_start,
                "length":  i - run_start,
                "bold":    prev_flags[0],
                "italic":  prev_flags[1],
                "underline": prev_flags[2],
                "allcaps": prev_flags[3],
            })
            run_start  = i
            prev_flags = flags

    runs.append({
        "start":   run_start,
        "length":  text_length - run_start,
        "bold":    prev_flags[0],
        "italic":  prev_flags[1],
        "underline": prev_flags[2],
        "allcaps": prev_flags[3],
    })
    return runs


def apply_format_runs(new_note, runs, line_start_in_orig):
    """Apply stored format runs to a new note, offset by line_start_in_orig."""
    fmt        = new_note.GetFormattedText()
    line_len   = len(new_note.Text)

    for run in runs:
        # Intersect run with this line's character range
        run_end  = run["start"] + run["length"]
        line_end = line_start_in_orig + line_len

        isect_start = max(run["start"], line_start_in_orig)
        isect_end   = min(run_end, line_end)

        if isect_end <= isect_start:
            continue

        # Translate to local offsets within the new note
        local_start  = isect_start - line_start_in_orig
        local_length = isect_end - isect_start
        tr = DB.TextRange(local_start, local_length)

        if run["bold"]:
            fmt.SetBoldStatus(tr, True)
        if run["italic"]:
            fmt.SetItalicStatus(tr, True)
        if run["underline"]:
            fmt.SetUnderlineStatus(tr, True)
        if run["allcaps"]:
            fmt.SetAllCapsStatus(tr, True)

    new_note.SetFormattedText(fmt)


# ── WORD WRAP ─────────────────────────────────────────────────────────────────

def wrap_segment(text, box_width, char_width):
    if char_width <= 0:
        return [text]
    max_chars = int(box_width / char_width)
    if max_chars <= 0:
        return [text]

    words   = text.split(" ")
    lines   = []
    current = ""

    for word in words:
        test = (current + " " + word).strip()
        if len(test) <= max_chars:
            current = test
        else:
            if current:
                lines.append(current)
            current = word

    if current:
        lines.append(current)

    return lines if lines else [text]


def split_note_to_lines(note):
    """Return list of (line_text, start_index_in_original) tuples."""
    text_type   = doc.GetElement(note.GetTypeId())
    text_size   = text_type.get_Parameter(
        DB.BuiltInParameter.TEXT_SIZE).AsDouble()
    width_param = text_type.get_Parameter(
        DB.BuiltInParameter.TEXT_WIDTH_SCALE)
    width_scale = width_param.AsDouble() if width_param else 1.0
    box_width   = note.Width
    char_width  = text_size * width_scale * 0.6

    raw      = note.Text.replace("\r\n", "\n").replace("\r", "\n")
    segments = raw.split("\n")

    result      = []
    orig_offset = 0  # character position in original string

    for seg in segments:
        if not seg.strip():
            result.append((None, orig_offset))
            orig_offset += len(seg) + 1  # +1 for the \n
            continue

        wrapped      = wrap_segment(seg, box_width, char_width)
        seg_offset   = 0

        for line in wrapped:
            result.append((line, orig_offset + seg_offset))
            seg_offset += len(line) + 1  # +1 for the space between wrapped words

        orig_offset += len(seg) + 1  # +1 for the \n

    return result


# ── SPLIT AND CREATE ──────────────────────────────────────────────────────────

new_ids = List[DB.ElementId]()

with revit.Transaction("Seed43 - Split Text Notes"):
    for note in text_notes:
        line_data = split_note_to_lines(note)
        real_lines = [(t, o) for t, o in line_data if t is not None]

        if len(real_lines) <= 1:
            continue

        # Read formatting before deletion
        fmt_text   = note.GetFormattedText()
        text_len   = len(note.Text)
        runs       = get_format_runs(fmt_text, text_len)

        text_type  = doc.GetElement(note.GetTypeId())
        text_size  = text_type.get_Parameter(
            DB.BuiltInParameter.TEXT_SIZE).AsDouble()
        line_height  = text_size * 1.5
        text_type_id = note.GetTypeId()
        origin       = note.Coord

        new_notes = []
        for i, (line_text, orig_offset) in enumerate(line_data):
            if line_text is None:
                continue
            point = DB.XYZ(
                origin.X,
                origin.Y - i * line_height,
                origin.Z
            )
            new_note = DB.TextNote.Create(
                doc,
                doc.ActiveView.Id,
                point,
                line_text,
                text_type_id
            )
            new_notes.append((new_note, orig_offset))
            new_ids.Add(new_note.Id)

        doc.Delete(note.Id)

        # Apply formatting to each new note
        for new_note, orig_offset in new_notes:
            any_format = any(
                r["bold"] or r["italic"] or r["underline"] or r["allcaps"]
                for r in runs
            )
            if any_format:
                apply_format_runs(new_note, runs, orig_offset)

if new_ids.Count == 0:
    forms.alert(
        "Nothing to split. Selected notes are already single lines.",
        title="Split Text"
    )
else:
    uidoc.Selection.SetElementIds(new_ids)
    forms.alert(
        "Done. {} lines created.".format(new_ids.Count),
        title="Split Text"
    )
