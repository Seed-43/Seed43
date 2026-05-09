# -*- coding: utf-8 -*-
__title__  = "Parameter Manager"
__author__ = "Seed43"
__doc__    = """
𝐕𝐄𝐑𝐒𝐈𝐎𝐍 𝟐𝟔𝟎𝟓𝟎𝟏
_____________________________________________________________________
Description:

Checks which Seed43 shared parameters are installed in the current
project and lets you add any that are missing. Parameters are bound
to the correct Revit categories automatically, no manual setup needed.

Tick marks show what's already installed. Select any missing ones
and click Create Parameters to add them.
_____________________________________________________________________
"""

import os
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    Transaction, SharedParameterElement,
    BuiltInCategory, CategorySet, InstanceBinding
)

# BuiltInParameterGroup was moved in Revit 2024 — handle both versions
try:
    from Autodesk.Revit.DB import BuiltInParameterGroup
except ImportError:
    from Autodesk.Revit.DB import GroupTypeId as BuiltInParameterGroup

doc = revit.doc
app = __revit__.Application

# ── Category mapping ─────────────────────────────────────────────────────────
# Maps the GROUP NAME in the shared parameter file to the Revit categories
# that parameter should be bound to.
GROUP_CATEGORIES = {
    "Sheets":  [BuiltInCategory.OST_Sheets, BuiltInCategory.OST_Views],
    "Project": [BuiltInCategory.OST_ProjectInformation],
    "Data":    [BuiltInCategory.OST_ProjectInformation],
}

# ── Shared parameter file ────────────────────────────────────────────────────
script_dir = os.path.dirname(__file__)
sp_path    = os.path.join(script_dir, "Seed43SharedParameters.txt")

if not os.path.isfile(sp_path):
    forms.alert(
        "Shared parameter file not found:\n{}".format(sp_path),
        exitscript=True
    )

# Save and restore the user's existing shared param file path
original_sp_path = app.SharedParametersFilename
try:
    app.SharedParametersFilename = sp_path
    sp_file = app.OpenSharedParameterFile()
finally:
    # Always restore — even if OpenSharedParameterFile throws
    if original_sp_path:
        app.SharedParametersFilename = original_sp_path

if not sp_file:
    forms.alert("Cannot open shared parameter file:\n{}".format(sp_path), exitscript=True)


# ── Build param list from file ───────────────────────────────────────────────
params = []
for group in sp_file.Groups:
    group_name = group.Name
    categories = GROUP_CATEGORIES.get(group_name, [BuiltInCategory.OST_ProjectInformation])
    for definition in group.Definitions:
        params.append({
            "name":       definition.Name,
            "guid":       definition.GUID,
            "definition": definition,
            "group_name": group_name,
            "categories": categories,
        })


# ── Detection — is this param already bound to any of its target categories? ─
def is_installed(param):
    """
    Return True if a parameter with this GUID OR this name is already
    bound in the document. GUID match is exact; name match catches params
    that were created from a different shared parameter file.
    """
    # 1. Exact GUID match
    if SharedParameterElement.Lookup(doc, param["guid"]):
        binding_map = doc.ParameterBindings
        it = binding_map.ForwardIterator()
        it.Reset()
        while it.MoveNext():
            if it.Key and it.Key.Name == param["name"]:
                return True

    # 2. Name-only fallback — catches same-name params from a different file
    binding_map = doc.ParameterBindings
    it = binding_map.ForwardIterator()
    it.Reset()
    while it.MoveNext():
        if it.Key and it.Key.Name == param["name"]:
            return True

    return False


# ── Build UI list ─────────────────────────────────────────────────────────────
# Separate installed from missing so installed ones are shown but not selectable
installed_items = []
missing_items   = []

for p in sorted(params, key=lambda x: x["name"]):
    if is_installed(p):
        installed_items.append("✔  " + p["name"])
    else:
        missing_items.append("❌  " + p["name"])

# Show missing first, then installed (greyed out context)
ui_list = missing_items + (["─── Already installed ───"] if installed_items else []) + installed_items

selected_display = forms.SelectFromList.show(
    ui_list,
    title="Seed43 Parameter Manager",
    multiselect=True,
    button_name="Create Parameters",
    width=700,
    height=750,
    info="Select missing parameters (❌) to install them.\n"
         "Parameters marked ✔ are already in this project."
)

if not selected_display:
    script.exit()

# Filter out divider and already-installed selections
to_create = []
for item in selected_display:
    if item.startswith("─"):
        continue
    if item.startswith("✔"):
        continue
    name  = item.replace("❌  ", "").strip()
    param = next((p for p in params if p["name"] == name), None)
    if param:
        to_create.append(param)

if not to_create:
    forms.alert("All selected parameters are already installed.", exitscript=True)


# ── Create parameters ─────────────────────────────────────────────────────────
created = []
skipped = []
errors  = []

# Re-open the shared param file for the transaction (restore path temporarily)
original_sp_path = app.SharedParametersFilename
app.SharedParametersFilename = sp_path
sp_file_for_create = app.OpenSharedParameterFile()

t = Transaction(doc, "Seed43 — Create Shared Parameters")
t.Start()

try:
    binding_map = doc.ParameterBindings

    for param in to_create:
        name = param["name"]

        # Find the live definition from the re-opened file
        live_definition = None
        for grp in sp_file_for_create.Groups:
            for defn in grp.Definitions:
                if defn.GUID == param["guid"]:
                    live_definition = defn
                    break
            if live_definition:
                break

        if not live_definition:
            errors.append("{} — definition not found in file".format(name))
            continue

        # Register the shared parameter element in the document if needed
        try:
            existing_sp = SharedParameterElement.Lookup(doc, param["guid"])
            if not existing_sp:
                SharedParameterElement.Create(doc, live_definition)
        except Exception as e:
            errors.append("{} — registration failed: {}".format(name, str(e)))
            continue

        # Build category set for this parameter's target categories
        cat_set = app.Create.NewCategorySet()
        for bic in param["categories"]:
            try:
                cat = doc.Settings.Categories.get_Item(bic)
                if cat:
                    cat_set.Insert(cat)
            except Exception:
                pass

        if cat_set.IsEmpty:
            errors.append("{} — no valid categories resolved".format(name))
            continue

        # Bind
        binding = app.Create.NewInstanceBinding(cat_set)
        try:
            if not binding_map.Contains(live_definition):
                try:
                    pg = BuiltInParameterGroup.PG_DATA      # Revit 2023 and earlier
                except AttributeError:
                    pg = BuiltInParameterGroup.Data         # Revit 2024+
                binding_map.Insert(live_definition, binding, pg)
                created.append(name)
            else:
                skipped.append("{} (already bound)".format(name))
        except Exception as e:
            errors.append("{} — binding failed: {}".format(name, str(e)))

    t.Commit()

except Exception as e:
    t.RollBack()
    forms.alert("Transaction failed and was rolled back:\n{}".format(str(e)), exitscript=True)

finally:
    # Always restore the user's original shared param file
    if original_sp_path:
        app.SharedParametersFilename = original_sp_path


# ── Result summary ────────────────────────────────────────────────────────────
lines = []
if created:
    lines.append("CREATED ({}):\n{}".format(len(created), "\n".join("  • " + n for n in created)))
if skipped:
    lines.append("SKIPPED ({}):\n{}".format(len(skipped), "\n".join("  • " + n for n in skipped)))
if errors:
    lines.append("ERRORS ({}):\n{}".format(len(errors),  "\n".join("  • " + n for n in errors)))

forms.alert(
    "\n\n".join(lines) if lines else "No changes made.",
    title="Parameter Manager — Done"
)
