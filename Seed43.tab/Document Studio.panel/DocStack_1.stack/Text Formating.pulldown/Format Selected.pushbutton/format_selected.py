# -*- coding: utf-8 -*-
# format_selected.py
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
