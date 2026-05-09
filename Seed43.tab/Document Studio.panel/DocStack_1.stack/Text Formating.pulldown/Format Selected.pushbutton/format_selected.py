# -*- coding: utf-8 -*-
__title__  = "Format Selected Text Notes"
__author__  = "Seed43"
__doc__     = """
𝐕𝐄𝐑𝐒𝐈𝐎𝐍 𝟐𝟔𝟎𝟓𝟎𝟏
_____________________________________________________________________
Description:
Interactive formatting tool for Text Notes in Revit.

- Apply formatting instantly to selected Text Notes
- Supports:
    - Bold
    - Italic
    - Underline
- Works in continuous selection mode for fast repetitive editing
- Updates formatting immediately as elements are picked in the model

Designed for fast manual annotation formatting and documentation
cleanup.
_____________________________________________________________________
How-to:
-> Run the tool
-> Select formatting options (Bold, Italic, Underline)
-> Pick Text Notes in the model
-> Continue selecting multiple elements
-> Press ESC to finish
_____________________________________________________________________
Notes:
- Only Text Notes are affected
- Formatting applies to the full text content of each note
- Ideal for rapid annotation checking and standardisation
_____________________________________________________________________
Last update:
- Initial release
_____________________________________________________________________
"""

from pyrevit import revit, DB, forms, script

doc = revit.doc


# ── SELECT FORMATTING OPTIONS ─────────────────────────────────────────────────

options = ["Bold", "Italic", "Underline"]

selected_formats = forms.SelectFromList.show(
    options,
    title="Select Text Formatting:",
    multiselect=True
)

if not selected_formats:
    script.exit()

do_bold      = "Bold"      in selected_formats
do_italic    = "Italic"    in selected_formats
do_underline = "Underline" in selected_formats


# ── PICK AND FORMAT TEXTS ─────────────────────────────────────────────────────

with forms.WarningBar(title="Pick Text Notes to format:"):
    while True:
        target_texts = revit.pick_elements()
        if not target_texts:
            break

        target_text_notes = [
            el for el in target_texts if isinstance(el, DB.TextNote)]
        if not target_text_notes:
            continue

        with revit.Transaction("Format Text"):
            for text in target_text_notes:
                formatted_text = text.GetFormattedText()
                full_range     = DB.TextRange(0, len(text.Text))

                if do_bold:
                    formatted_text.SetBoldStatus(full_range, True)
                if do_italic:
                    formatted_text.SetItalicStatus(full_range, True)
                if do_underline:
                    formatted_text.SetUnderlineStatus(full_range, True)

                text.SetFormattedText(formatted_text)
