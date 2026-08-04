# -*- coding: utf-8 -*-
# script.py
from pyrevit import revit, DB, forms, script
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List

doc   = revit.doc
uidoc = revit.uidoc

# ── PICK TEXT NOTES ───────────────────────────────────────────────────────────

try:
    with forms.WarningBar(title="Select Text Notes to merge (ESC to Cancel):"):
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element, "Select text notes to merge")
except Exception:
    script.exit()

text_notes = []
for ref in refs:
    el = doc.GetElement(ref.ElementId)
    if isinstance(el, DB.TextNote):
        text_notes.append(el)

if not text_notes:
    forms.alert("No TextNotes selected.", title="MText", exitscript=True)

# ── SORT TOP TO BOTTOM, LEFT TO RIGHT ────────────────────────────────────────

def get_bb(note):
    return note.get_BoundingBox(doc.ActiveView)

text_notes.sort(key=lambda n: (-get_bb(n).Max.Y, n.Coord.X))

# ── TOUCHING CHECK ────────────────────────────────────────────────────────────

def boxes_touch(upper, lower):
    bb_upper = get_bb(upper)
    bb_lower = get_bb(lower)
    if bb_upper is None or bb_lower is None:
        return False
    return bb_lower.Max.Y >= bb_upper.Min.Y

# ── STRIP NOTE TEXT ───────────────────────────────────────────────────────────

def strip_text(note):
    return note.Text.replace("\r\n", " ").replace("\r", " ").replace("\n", " ").strip()

# ── CHAR WIDTH ESTIMATE ───────────────────────────────────────────────────────

def get_char_width(note):
    text_type   = doc.GetElement(note.GetTypeId())
    text_size   = text_type.get_Parameter(
        DB.BuiltInParameter.TEXT_SIZE).AsDouble()
    width_param = text_type.get_Parameter(
        DB.BuiltInParameter.TEXT_WIDTH_SCALE)
    width_scale = width_param.AsDouble() if width_param else 1.0
    return text_size * width_scale * 0.6

# ── BUILD MERGED STRING ───────────────────────────────────────────────────────
# Initially join all touching notes with \r\n and gaps with \r\n\r\n.
# We will post-process to collapse \r\n between wrapped lines.

parts       = []   # (stripped_text, source_note)
merged_text = ""

for i, note in enumerate(text_notes):
    text = strip_text(note)
    if not text:
        continue

    if i == 0:
        joiner = ""
    elif boxes_touch(text_notes[i - 1], note):
        joiner = "\r\n"
    else:
        joiner = "\r\n\r\n"

    parts.append((text, note))
    merged_text += joiner + text

# ── POST-PROCESS: COLLAPSE WRAPPED LINES ─────────────────────────────────────
# Split on single \r\n and estimate each line's fill % of the merged box
# width. A "full" line (above threshold) followed by another is a wrap, so
# that \r\n becomes a space; a short line ends a paragraph and keeps its
# \r\n. Double \r\n\r\n paragraph breaks are always preserved.

FULL_LINE_THRESHOLD = 0.72  # line must be >= 72% of box width to be "full"

base_note  = text_notes[0]
widest     = max(note.Width for note in text_notes)
char_width = get_char_width(base_note)
max_chars  = widest / char_width if char_width > 0 else 9999

def line_fill(line):
    """Fraction of box width this line occupies."""
    if max_chars <= 0:
        return 0
    return len(line.strip()) / max_chars

# Split on double breaks first to preserve paragraph gaps
paragraphs = merged_text.split("\r\n\r\n")

collapsed_paragraphs = []
for para in paragraphs:
    lines = para.split("\r\n")
    if len(lines) <= 1:
        collapsed_paragraphs.append(para)
        continue

    # Walk lines: if current line is "full", collapse the \r\n after it
    result = lines[0]
    for j in range(1, len(lines)):
        prev_line      = lines[j - 1].strip()
        ends_sentence  = prev_line.endswith(".")
        is_full        = line_fill(prev_line) >= FULL_LINE_THRESHOLD

        if is_full and not ends_sentence:
            result += " " + lines[j]     # wrapped continuation, join with space
        else:
            result += "\r\n" + lines[j]  # end of sentence or short line, keep break

    collapsed_paragraphs.append(result)

merged_text = "\r\n\r\n".join(collapsed_paragraphs)

# ── CREATE MERGED NOTE ────────────────────────────────────────────────────────

with revit.Transaction("Seed43 - MText Merge"):
    opts     = DB.TextNoteOptions(base_note.GetTypeId())
    new_note = DB.TextNote.Create(
        doc,
        doc.ActiveView.Id,
        base_note.Coord,
        merged_text,
        opts
    )
    new_note.Width = widest

    # ── APPLY FORMATTING ─────────────────────────────────────────────────────

    merged_fmt   = new_note.GetFormattedText()
    search_start = 0

    for text, note in parts:
        src_fmt  = note.GetFormattedText()
        src_orig = note.Text

        real_idx = [i for i, ch in enumerate(src_orig)
                    if ch not in ("\r", "\n")]

        found = merged_fmt.Find(text, search_start, False, False)
        if found.Length == 0:
            search_start += len(text)
            continue

        base_offset  = found.Start
        search_start = base_offset + len(text)

        def flags_at(orig_i):
            r = DB.TextRange(orig_i, 1)
            return (
                int(src_fmt.GetBoldStatus(r))      == 1,
                int(src_fmt.GetItalicStatus(r))    == 1,
                int(src_fmt.GetUnderlineStatus(r)) == 1,
                int(src_fmt.GetAllCapsStatus(r))   == 1,
            )

        n = len(real_idx)
        if n == 0:
            continue

        prev  = flags_at(real_idx[0])
        start = 0

        for j in range(1, n):
            curr = flags_at(real_idx[j])
            if curr != prev:
                if any(prev):
                    tr = DB.TextRange(base_offset + start, j - start)
                    if prev[0]: merged_fmt.SetBoldStatus(tr, True)
                    if prev[1]: merged_fmt.SetItalicStatus(tr, True)
                    if prev[2]: merged_fmt.SetUnderlineStatus(tr, True)
                    if prev[3]: merged_fmt.SetAllCapsStatus(tr, True)
                start = j
                prev  = curr

        if any(prev):
            tr = DB.TextRange(base_offset + start, n - start)
            if prev[0]: merged_fmt.SetBoldStatus(tr, True)
            if prev[1]: merged_fmt.SetItalicStatus(tr, True)
            if prev[2]: merged_fmt.SetUnderlineStatus(tr, True)
            if prev[3]: merged_fmt.SetAllCapsStatus(tr, True)

    new_note.SetFormattedText(merged_fmt)
    new_id = new_note.Id

uidoc.Selection.SetElementIds(List[DB.ElementId]([new_id]))
forms.alert(
    "Done. Merged note is selected. Move it as needed.",
    title="MText"
)
