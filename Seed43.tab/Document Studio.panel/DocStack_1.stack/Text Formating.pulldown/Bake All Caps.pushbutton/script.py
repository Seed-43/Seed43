# -*- coding: utf-8 -*-
# "Bake All Caps"
# "Seed43"
# """
# Strips the All Caps override (Ctrl+Shift+A) off Text Notes and rewrites the
# stored text as typed capitals, so each note reads exactly the same on the
# sheet but no longer depends on the formatting flag to look uppercase.
# """

# ── IMPORTS ───────────────────────────────────────────────────────────────────

from System.Collections.Generic import List

from pyrevit import revit, DB, forms, script

try:
    from Snippets import _dialogs as sdlg
except Exception:
    sdlg = None

# ── CONSTANTS ─────────────────────────────────────────────────────────────────

doc    = revit.doc
uidoc  = revit.uidoc
output = script.get_output()

TITLE = "Bake All Caps"

# NOTE: FormatStatus.None can't be written as an attribute - `None` is a
# reserved word in Python, so the enum member has to come out via getattr.
CAPS_NONE = getattr(DB.FormatStatus, "None")
CAPS_ALL  = DB.FormatStatus.All

SCOPE_SELECTION = "selection"
SCOPE_VIEW      = "view"
SCOPE_MODEL     = "model"

SCOPE_LABELS = {
    SCOPE_SELECTION: "Selection",
    SCOPE_VIEW:      "Active View",
    SCOPE_MODEL:     "Whole Model",
}

PREVIEW_ROWS = 6

# ── DIALOGS ───────────────────────────────────────────────────────────────────

def _alert(message, title=TITLE, exitscript=False):
    """Themed popup via the shared Snippets dialog lib, falls back to
    pyRevit's default forms.alert if the shared lib isn't available."""
    if sdlg:
        sdlg.message(message, title=title)
    else:
        forms.alert(message, title=title)
    if exitscript:
        script.exit()


def _choose(message, options, title=TITLE, detail_text=''):
    """Themed multi-button choice, falls back to pyRevit's switch window.

    Args:
        options (list of tuples): (key, label) pairs, last one is the primary
            button. Returns the chosen key, or None if dismissed.
    """
    if sdlg:
        return sdlg.choice(message, options, title=title,
                           detail_text=detail_text)

    picked = forms.CommandSwitchWindow.show(
        [label for _, label in options], message=message)
    for key, label in options:
        if label == picked:
            return key
    return None


def _clip(text, width=44):
    """Flatten a note to one short single-line string for the preview panel."""
    flat = u" ".join(text.split())
    if len(flat) <= width:
        return flat
    return flat[:width - 1] + u"…"

# ── CORE LOGIC ────────────────────────────────────────────────────────────────

# --- Collection ---

def notes_in_scope(scope):
    """Return every Text Note the chosen scope covers."""
    if scope == SCOPE_SELECTION:
        picked = [doc.GetElement(eid)
                  for eid in uidoc.Selection.GetElementIds()]
        return [el for el in picked if isinstance(el, DB.TextNote)]

    if scope == SCOPE_VIEW:
        collector = DB.FilteredElementCollector(doc, doc.ActiveView.Id)
    else:
        collector = DB.FilteredElementCollector(doc)
    return list(collector.OfClass(DB.TextNote))


def carrying_override(notes):
    """Return only the notes that actually have the All Caps flag set."""
    found = []
    for note in notes:
        try:
            if note.GetFormattedText().GetAllCapsStatus() != CAPS_NONE:
                found.append(note)
        except Exception:
            pass
    return found


def checkout(notes):
    """Make the notes editable in a workshared model, no-op otherwise."""
    if not doc.IsWorkshared:
        return
    try:
        DB.WorksharingUtils.CheckoutElements(
            doc, List[DB.ElementId]([n.Id for n in notes]))
    except Exception:
        pass

# --- Inspection ---

def caps_flags(formatted, plain):
    """Return one boolean per character, True where All Caps is active.

    Revit only reports an aggregate status (None/All/Mixed) for a range, so
    the only way to find the exact spans is to ask character by character.
    """
    return [formatted.GetAllCapsStatus(DB.TextRange(i, 1)) == CAPS_ALL
            for i in range(len(plain))]


def rendered_text(plain, flags):
    """Return what the note currently shows on the sheet - the stored text
    with the flagged characters uppercased."""
    return u"".join(ch.upper() if flags[i] else ch
                    for i, ch in enumerate(plain))


def capped_runs(flags):
    """Return (start, length) for each contiguous run of flagged characters."""
    runs  = []
    start = None
    for i, is_capped in enumerate(flags):
        if is_capped and start is None:
            start = i
        elif not is_capped and start is not None:
            runs.append((start, i - start))
            start = None
    if start is not None:
        runs.append((start, len(flags) - start))
    return runs

# --- Rewriting ---

def bake_note(note):
    """Rewrite one note's flagged spans as typed capitals, clear the override.

    Returns (status, displayed_before, stored_after, problem) where status is
    one of "rewritten", "flagonly", "skipped" or "suspect". Nothing is written
    until every span has been checked, so a note is never half-converted.
    """
    formatted = note.GetFormattedText()
    plain     = formatted.GetPlainText()

    if formatted.GetAllCapsStatus() == CAPS_NONE:
        return ("skipped", plain, plain, "override already clear")

    flags  = caps_flags(formatted, plain)
    before = rendered_text(plain, flags)

    # NOTE: every replacement is worked out before any is applied. A few
    # characters grow when uppercased (German 'ß' -> 'SS'), which would shift
    # each index after them - better to skip the note than half-convert it.
    edits = []
    for start, length in capped_runs(flags):
        source = plain[start:start + length]
        upper  = source.upper()
        if len(upper) != len(source):
            return ("skipped", before, plain, "uppercasing changes text length")
        if upper != source:
            edits.append((start, length, upper))

    for start, length, upper in edits:
        formatted.SetPlainText(DB.TextRange(start, length), upper)

    formatted.SetAllCapsStatus(formatted.AsTextRange(), False)
    note.SetFormattedText(formatted)

    after = note.GetFormattedText().GetPlainText()
    # The whole point is that the sheet reads identically afterwards, so the
    # new stored text must equal what was displayed before. Anything else is
    # a silent wording change - surface it instead of reporting a clean run.
    if after != before:
        return ("suspect", before, after, "result differs from what was shown")

    return ("rewritten" if after != plain else "flagonly", before, after, None)

# ── UI / ENTRY POINT ──────────────────────────────────────────────────────────

scope = _choose(
    "Which Text Notes should be checked for the All Caps override?",
    [(SCOPE_SELECTION, "Selection"),
     (SCOPE_VIEW,      "Active View"),
     (SCOPE_MODEL,     "Whole Model")])

if not scope:
    script.exit()

notes = notes_in_scope(scope)
if not notes:
    _alert("No Text Notes found in the {} scope."
           .format(SCOPE_LABELS[scope].lower()), exitscript=True)

targets = carrying_override(notes)
if not targets:
    _alert("Checked {} Text Note(s) - none carry the All Caps override."
           .format(len(notes)), exitscript=True)

# --- Confirm, with a sample of what will actually change ---

preview = []
for note in targets:
    if len(preview) >= PREVIEW_ROWS:
        break
    formatted = note.GetFormattedText()
    plain     = formatted.GetPlainText()
    shown     = rendered_text(plain, caps_flags(formatted, plain))
    if _clip(plain) != _clip(shown):
        preview.append(u"{}\n   → {}".format(_clip(plain), _clip(shown)))

answer = _choose(
    "{} of the {} Text Note(s) checked carry the All Caps override.\n\n"
    "Their stored text will be rewritten as typed capitals and the override "
    "cleared. What each note shows on the sheet stays exactly the same."
    .format(len(targets), len(notes)),
    [("cancel", "Cancel"), ("run", "Bake Caps")],
    detail_text=u"\n".join(preview))

if answer != "run":
    script.exit()

# --- Run ---

checkout(targets)

rewritten = []
flagonly  = []
skipped   = []
suspect   = []
cancelled = False

with revit.Transaction(TITLE):
    with forms.ProgressBar(title="Baking All Caps  ({value} of {max_value})",
                           cancellable=True) as pb:
        for i, note in enumerate(targets):
            if pb.cancelled:
                cancelled = True
                break
            try:
                status, before, after, problem = bake_note(note)
            except Exception as err:
                skipped.append((note, str(err)))
                pb.update_progress(i + 1, len(targets))
                continue

            if status == "suspect":
                suspect.append((note, before, after))
            elif problem:
                skipped.append((note, problem))
            elif status == "rewritten":
                rewritten.append((note, before))
            else:
                flagonly.append(note)
            pb.update_progress(i + 1, len(targets))

# --- Report ---

output.print_md("# {}".format(TITLE))
output.print_md(
    "**Scope:** {}  &nbsp;|&nbsp;  **Text Notes checked:** {}  "
    "&nbsp;|&nbsp;  **carried the override:** {}"
    .format(SCOPE_LABELS[scope], len(notes), len(targets)))

if cancelled:
    output.print_md(
        "> **Cancelled part way through.** The {} note(s) already processed "
        "keep their change - re-run the tool to finish the rest."
        .format(len(rewritten) + len(flagonly)))

output.print_md(
    "- **{}** rewritten as typed capitals\n"
    "- **{}** already typed uppercase, override cleared only\n"
    "- **{}** left untouched\n"
    "- **{}** need checking"
    .format(len(rewritten), len(flagonly), len(skipped), len(suspect)))

if suspect:
    output.print_md("### Need checking")
    output.print_md(
        "These notes were converted, but the result no longer matches what "
        "was on the sheet beforehand. Review them and undo if wrong.")
    rows = [[output.linkify(note.Id), _clip(before, 46), _clip(after, 46)]
            for note, before, after in suspect]
    output.print_table(rows, columns=["Note", "Was shown", "Now reads"])

if rewritten:
    output.print_md("### Rewritten")
    rows = [[output.linkify(note.Id), _clip(before, 70)]
            for note, before in rewritten]
    output.print_table(rows, columns=["Note", "Text"])

if skipped:
    output.print_md("### Left untouched")
    rows = [[output.linkify(note.Id), reason] for note, reason in skipped]
    output.print_table(rows, columns=["Note", "Reason"])

output.print_md(
    "---\n*Nothing has been saved. Check the drawings, then save the model "
    "yourself if you are happy with the result.*")
