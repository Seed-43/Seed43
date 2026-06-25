# -*- coding: utf-8 -*-
# find_replace.py
from pyrevit import revit, DB, forms, script

doc    = revit.doc
uidoc  = revit.uidoc
output = script.get_output()

# ── SELECT SCOPE ──────────────────────────────────────────────────────────────

scope_options = [
    "Selected Text Notes",
    "Active View Only",
    "Entire Project",
]

scope = forms.SelectFromList.show(
    scope_options,
    title="Find and Replace - Scope",
    multiselect=False,
    button_name="Next"
)

if not scope:
    script.exit()

# ── SELECT CASE SENSITIVITY ───────────────────────────────────────────────────

case_options = [
    "Case-Sensitive",
    "Case-Insensitive",
]

case_choice = forms.SelectFromList.show(
    case_options,
    title="Find and Replace - Match Mode",
    multiselect=False,
    button_name="Next"
)

if not case_choice:
    script.exit()

case_sensitive = case_choice == "Case-Sensitive"

# ── FIND AND REPLACE STRINGS ──────────────────────────────────────────────────

find_str = forms.ask_for_string(
    prompt="Text to find:",
    title="Find and Replace"
)

if find_str is None:
    script.exit()

if find_str == "":
    forms.alert("Find string cannot be empty.", exitscript=True)

replace_str = forms.ask_for_string(
    prompt="Replace with (leave blank to delete):",
    title="Find and Replace",
    default=""
)

if replace_str is None:
    script.exit()

# ── COLLECT TEXT NOTES ────────────────────────────────────────────────────────

if scope == "Selected Text Notes":
    with forms.WarningBar(title="Pick Text Notes to search (ESC when done):"):
        selected = revit.pick_elements()

    if not selected:
        script.exit()

    all_textnotes = [el for el in selected if isinstance(el, DB.TextNote)]

    if not all_textnotes:
        forms.alert(
            "None of the selected elements are Text Notes.",
            exitscript=True
        )

elif scope == "Active View Only":
    all_textnotes = list(
        DB.FilteredElementCollector(doc, doc.ActiveView.Id)
          .OfClass(DB.TextNote)
          .ToElements()
    )

else:
    all_textnotes = list(
        DB.FilteredElementCollector(doc)
          .OfClass(DB.TextNote)
          .ToElements()
    )

# ── FIND MATCHES ──────────────────────────────────────────────────────────────

def contains_match(text, find, sensitive):
    if sensitive:
        return find in text
    return find.lower() in text.lower()

def do_replace(text, find, replace, sensitive):
    if sensitive:
        return text.replace(find, replace)
    # Case-insensitive replace preserving original casing of non-matched text
    import re
    return re.sub(re.escape(find), replace, text, flags=re.IGNORECASE)

matches = [
    tn for tn in all_textnotes
    if contains_match(tn.Text or "", find_str, case_sensitive)
]

if not matches:
    forms.alert(
        "No Text Notes containing \"{}\" were found.\n\n"
        "Scope      : {}\n"
        "Match mode : {}".format(
            find_str, scope, case_choice),
        title="No Matches"
    )
    script.exit()

# ── CONFIRM ───────────────────────────────────────────────────────────────────

confirm = forms.alert(
    "Found {} Text Note(s) matching \"{}\".\n\n"
    "Find        : {}\n"
    "Replace     : {}\n"
    "Scope       : {}\n"
    "Match mode  : {}\n\n"
    "Continue?".format(
        len(matches),
        find_str,
        find_str,
        replace_str if replace_str != "" else "(delete)",
        scope,
        case_choice
    ),
    title="Find and Replace",
    yes=True,
    no=True
)

if not confirm:
    script.exit()

# ── APPLY REPLACEMENTS ────────────────────────────────────────────────────────

fixed   = 0
skipped = 0

with revit.Transaction("Find and Replace Text"):
    for tn in matches:
        try:
            original = tn.Text
            cleaned  = do_replace(original, find_str, replace_str, case_sensitive)

            if cleaned != original:
                tn.Text = cleaned
                fixed += 1
        except Exception as e:
            output.print_md("Skipped element `{}`: {}".format(tn.Id, str(e)))
            skipped += 1

# ── RESULT ────────────────────────────────────────────────────────────────────

forms.alert(
    "Operation complete.\n\n"
    "Matched  : {}\n"
    "Fixed    : {}\n"
    "Skipped  : {}".format(len(matches), fixed, skipped),
    title="Done"
)
