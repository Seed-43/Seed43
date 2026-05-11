# -*- coding: utf-8 -*-
# match_width.py
from pyrevit import revit, DB, forms, script

doc = revit.doc

# ── FUNCTIONS ─────────────────────────────────────────────────────────────────

def copy_text_width(from_text, to_texts):
    """Copy the Width value from the reference text box to each target."""
    ref_width = from_text.Width

    with revit.Transaction("Match Text Width"):
        changed = 0
        for text in to_texts:
            if text.Width != ref_width:
                text.Width = ref_width
                changed += 1
        return changed

# ── PICK REFERENCE ────────────────────────────────────────────────────────────

with forms.WarningBar(title="Pick REFERENCE Text Note (source width):"):
    source_text = revit.pick_element()

if not source_text or not isinstance(source_text, DB.TextNote):
    forms.alert("Please select a valid Text Note.", exitscript=True)

# ── PICK TARGETS ──────────────────────────────────────────────────────────────

with forms.WarningBar(title="Pick TARGET Text Notes (ESC to finish):"):
    while True:
        target_texts = revit.pick_elements()
        if not target_texts:
            break

        target_text_notes = [
            el for el in target_texts
            if isinstance(el, DB.TextNote)
        ]

        if target_text_notes:
            copy_text_width(source_text, target_text_notes)
