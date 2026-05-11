# -*- coding: utf-8 -*-
# match_properties.py
from pyrevit import revit, DB, forms, script

doc = revit.doc

# ── FUNCTIONS ─────────────────────────────────────────────────────────────────

def match_properties(source, targets, match_style, match_width,
                     match_halign, match_valign):
    """Match selected properties from source to each target element."""
    with revit.Transaction("Match Label Properties"):
        changed = 0
        for target in targets:
            updated = False

            if match_style:
                try:
                    if target.GetTypeId() != source.GetTypeId():
                        target.ChangeTypeId(source.GetTypeId())
                        updated = True
                except Exception:
                    pass

            if match_width and hasattr(target, "Width"):
                try:
                    if hasattr(target, "KeepReadable"):
                        target.KeepReadable = False
                    if abs(target.Width - source.Width) > 0.001:
                        target.Width = source.Width
                        updated = True
                except Exception:
                    pass

            if match_halign and hasattr(target, "HorizontalAlignment"):
                try:
                    if target.HorizontalAlignment != source.HorizontalAlignment:
                        target.HorizontalAlignment = source.HorizontalAlignment
                        updated = True
                except Exception:
                    pass

            if match_valign and hasattr(target, "VerticalAlignment"):
                try:
                    if target.VerticalAlignment != source.VerticalAlignment:
                        target.VerticalAlignment = source.VerticalAlignment
                        updated = True
                except Exception:
                    pass

            if updated:
                changed += 1

        return changed

# ── SELECT PROPERTIES TO MATCH ────────────────────────────────────────────────

options = [
    "Text Style (Type)",
    "Width",
    "Horizontal Alignment",
    "Vertical Alignment"
]

selected = forms.SelectFromList.show(
    options,
    title="What do you want to match?",
    multiselect=True,
    button_name="Match Selected"
)

if not selected:
    script.exit()

match_style  = "Text Style (Type)"   in selected
match_width  = "Width"               in selected
match_halign = "Horizontal Alignment" in selected
match_valign = "Vertical Alignment"  in selected

# ── PICK REFERENCE ────────────────────────────────────────────────────────────

with forms.WarningBar(title="Pick REFERENCE Label or Text:"):
    source = revit.pick_element()

if not source or not isinstance(source, (DB.TextNote, DB.TextElement)):
    forms.alert(
        "Please select a Text Note or Annotation Label as reference.",
        exitscript=True
    )

# ── PICK TARGETS ──────────────────────────────────────────────────────────────

total_changed = 0

with forms.WarningBar(title="Pick TARGET Labels (ESC to finish):"):
    while True:
        targets = revit.pick_elements()
        if not targets:
            break

        valid_targets = [
            el for el in targets
            if isinstance(el, (DB.TextNote, DB.TextElement))
        ]

        if valid_targets:
            changed = match_properties(
                source, valid_targets,
                match_style, match_width, match_halign, match_valign
            )
            total_changed += changed

            if changed > 0:
                forms.toast("Updated {} label(s)".format(changed), title="Success")
            else:
                forms.toast(
                    "No changes applied (some properties may be locked)",
                    title="Match Properties"
                )

# ── RESULT ────────────────────────────────────────────────────────────────────

forms.alert(
    "Operation finished.\n\n"
    "Total labels updated: {}\n\n"
    "Properties attempted:\n"
    "- Text Style       : {}\n"
    "- Width            : {}\n"
    "- Horizontal Align : {}\n"
    "- Vertical Align   : {}".format(
        total_changed,
        "Yes" if match_style  else "No",
        "Yes" if match_width  else "No",
        "Yes" if match_halign else "No",
        "Yes" if match_valign else "No"
    ),
    title="Done"
)
