# -*- coding: utf-8 -*-
# leader_alignment_match.py
from pyrevit import revit, DB, forms, script

doc = revit.doc

# ── FUNCTIONS ─────────────────────────────────────────────────────────────────

def copy_text_alignment(from_text, to_texts):
    """Copy horizontal and vertical alignment and leader attachment points."""
    ref_halign       = from_text.HorizontalAlignment
    ref_valign       = from_text.VerticalAlignment
    ref_left_attach  = from_text.LeaderLeftAttachment
    ref_right_attach = from_text.LeaderRightAttachment

    with revit.Transaction("Match Text Alignment"):
        changed = 0
        for text in to_texts:
            updated = False

            if text.HorizontalAlignment != ref_halign:
                text.HorizontalAlignment = ref_halign
                updated = True

            if text.VerticalAlignment != ref_valign:
                text.VerticalAlignment = ref_valign
                updated = True

            if text.LeaderLeftAttachment != ref_left_attach:
                text.LeaderLeftAttachment = ref_left_attach
                updated = True

            if text.LeaderRightAttachment != ref_right_attach:
                text.LeaderRightAttachment = ref_right_attach
                updated = True

            if updated:
                changed += 1

        return changed

# ── PICK REFERENCE ────────────────────────────────────────────────────────────

with forms.WarningBar(title="Pick REFERENCE Text Note (alignment to copy):"):
    source_text = revit.pick_element()

    if source_text and isinstance(source_text, DB.TextNote):

        # ── PICK TARGETS ──────────────────────────────────────────────────────

        with forms.WarningBar(title="Pick TARGET Text Notes to match alignment:"):
            while True:
                target_texts = revit.pick_elements()
                if not target_texts:
                    break

                target_text_notes = [
                    el for el in target_texts
                    if isinstance(el, DB.TextNote)
                ]

                if target_text_notes:
                    copy_text_alignment(source_text, target_text_notes)
