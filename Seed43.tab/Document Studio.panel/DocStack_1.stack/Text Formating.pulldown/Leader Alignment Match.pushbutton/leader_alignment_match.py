# -*- coding: utf-8 -*-
__title__  = "Match Text Alignment"
__author__  = "Seed43"
__doc__     = """
𝐕𝐄𝐑𝐒𝐈𝐎𝐍 𝟐𝟔𝟎𝟓𝟎𝟏
_____________________________________________________________________
Description:
Matches paragraph alignment settings between Text Notes in the model.

Transfers alignment properties from a reference Text Note to selected
target Text Notes, ensuring consistent annotation presentation.

Includes:
- Horizontal Alignment (Left, Centre, Right)
- Vertical Alignment (Top, Middle, Bottom)
- Leader Attachment Points (Left, Right)

Does NOT modify text content or leaders themselves, only alignment
settings.
_____________________________________________________________________
How-to:
-> Run the tool
-> Select a REFERENCE Text Note (the one with the alignment to copy)
-> Select TARGET Text Notes to update
-> Repeat selection as needed (ESC to finish)

-> Alignment updates instantly on each selection
_____________________________________________________________________
Notes:
- Only works on Text Note elements
- Multiple targets can be selected in one run
- No changes are made to text content or leader geometry
- Press ESC at any time to exit safely
_____________________________________________________________________
Last update:
- Initial release
_____________________________________________________________________
"""

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
