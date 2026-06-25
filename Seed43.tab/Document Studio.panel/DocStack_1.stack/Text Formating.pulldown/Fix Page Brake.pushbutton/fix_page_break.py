# -*- coding: utf-8 -*-
# fix_page_break.py
from pyrevit import revit, DB, forms, script

doc   = revit.doc
uidoc = revit.uidoc

# ── PICK TEXT NOTES ───────────────────────────────────────────────────────────

with forms.WarningBar(title="Pick Text Notes to fix (ESC when done):"):
    selected = revit.pick_elements()

if not selected:
    script.exit()

textnotes = [el for el in selected if isinstance(el, DB.TextNote)]

if not textnotes:
    forms.alert("No Text Notes selected.", exitscript=True)

# ── APPLY FIX ─────────────────────────────────────────────────────────────────
# Work directly on FormattedText — formatting is never touched.
# Walk backwards to keep indices valid after each replacement.
# Double breaks (\r\r, \x0b\r, \x0b\x0b) are preserved.
# Single breaks (\r, \x0b) are replaced with a space.

fixed = 0

with revit.Transaction("Seed43 - Fix Page Breaks"):
    for tn in textnotes:
        fmt  = tn.GetFormattedText()
        text = tn.Text
        i    = len(text) - 1
        changed = False

        while i >= 0:
            ch = text[i]
            if ch == "\r" and i > 0 and text[i - 1] == "\r":
                i -= 2
                continue
            if ch == "\r" and i > 0 and text[i - 1] == "\x0b":
                i -= 2
                continue
            if ch == "\x0b" and i > 0 and text[i - 1] == "\x0b":
                i -= 2
                continue
            if ch == "\r" or ch == "\x0b":
                fmt.SetPlainText(DB.TextRange(i, 1), " ")
                changed = True
            i -= 1

        if changed:
            tn.SetFormattedText(fmt)
            fixed += 1

forms.alert(
    "Done. {} note(s) fixed.".format(fixed),
    title="Fix Page Breaks"
)
