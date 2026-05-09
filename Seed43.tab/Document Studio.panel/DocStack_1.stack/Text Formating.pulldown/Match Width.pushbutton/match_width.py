# -*- coding: utf-8 -*-
__title__  = "Match Text Width"
__author__  = "Seed43"
__doc__     = """
𝐕𝐄𝐑𝐒𝐈𝐎𝐍 𝟐𝟔𝟎𝟓𝟎𝟏
_____________________________________________________________________
Description:
Copies the text box width from a reference Text Note and applies it
to one or more target Text Notes, ensuring a consistent annotation
layout across views.

Only the physical width of the text box is changed. Text content
and formatting are not affected.
_____________________________________________________________________
How-to:
-> Run the tool
-> Click the REFERENCE Text Note (the one with the width you want)
-> Click the TARGET Text Notes one at a time (or pick a group)
-> Press ESC when finished

Width is applied instantly to each selection batch.
_____________________________________________________________________
Notes:
- Only works on Text Note elements
- Some view contexts may restrict editing the width of certain notes
- Multiple targets can be selected in one session
- No changes are made to text content or formatting
_____________________________________________________________________
Last update:
- Initial release
_____________________________________________________________________
"""

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
