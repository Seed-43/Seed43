# -*- coding: utf-8 -*-
# Project Units Manager
# Saves all project unit settings to a JSON file under .user, or restores
# them from it.
# Targets Revit 2024-2026 (ForgeTypeId API only).

import os
import json
from pyrevit import revit, DB, forms, script
from Snippets import _userdata

doc = revit.doc

# Settings live in .user so the updater can never overwrite them. Any file
# saved by an older version, back when it sat beside this script, is moved
# across on first run.
JSON_PATH = _userdata.migrate(
    os.path.join(os.path.dirname(__file__), "project_units.json"),
    _userdata.user_path("Units", "project_units.json"))

# ── Helpers ────────────────────────────────────────────────────────────────────

def forge_id_str(forge_type_id):
    try:
        return forge_type_id.TypeId
    except Exception:
        return str(forge_type_id)

def forge_id_from_str(s):
    return DB.ForgeTypeId(s)

def format_options_to_dict(fo):
    d = {}
    try:
        d["unit_type_id"] = forge_id_str(fo.GetUnitTypeId())
    except Exception:
        d["unit_type_id"] = None
    try:
        d["accuracy"] = fo.Accuracy
    except Exception:
        d["accuracy"] = None
    try:
        d["use_grouping"] = fo.UseGrouping
    except Exception:
        d["use_grouping"] = False
    try:
        d["suppress_leading_zeros"] = fo.SuppressLeadingZeros
    except Exception:
        d["suppress_leading_zeros"] = False
    try:
        d["suppress_trailing_zeros"] = fo.SuppressTrailingZeros
    except Exception:
        d["suppress_trailing_zeros"] = False
    try:
        d["show_measurement_line_below"] = fo.ShowMeasurementLineBelow
    except Exception:
        d["show_measurement_line_below"] = False
    try:
        d["prefix"] = fo.Prefix
    except Exception:
        d["prefix"] = ""
    try:
        d["suffix"] = fo.Suffix
    except Exception:
        d["suffix"] = ""
    return d

# ── Action picker ──────────────────────────────────────────────────────────────

action = forms.CommandSwitchWindow.show(
    ["Update Units from JSON", "Save Units to JSON"],
    message="What would you like to do?\n\n"
            "Update  - apply units from project_units.json to this project\n"
            "Save    - write current project units to project_units.json",
    title="Project Units Manager"
)

if not action:
    script.exit()

# ══════════════════════════════════════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════════════════════════════════════
if action == "Save Units to JSON":
    units     = doc.GetUnits()
    all_specs = DB.UnitUtils.GetAllMeasurableSpecs()
    export    = {}

    for spec_id in all_specs:
        try:
            fo = units.GetFormatOptions(spec_id)
            export[forge_id_str(spec_id)] = format_options_to_dict(fo)
        except Exception:
            pass

    with open(JSON_PATH, "w") as f:
        json.dump(export, f, indent=2)

# ══════════════════════════════════════════════════════════════════════════════
# UPDATE
# ══════════════════════════════════════════════════════════════════════════════
else:
    if not os.path.isfile(JSON_PATH):
        script.exit()

    with open(JSON_PATH, "r") as f:
        import_data = json.load(f)

    with revit.Transaction("Project Units Manager - apply units"):
        units = doc.GetUnits()

        for spec_str, fo_dict in import_data.items():
            try:
                spec_id = forge_id_from_str(spec_str)

                # Get the existing FormatOptions for this spec from the Units
                # object - mutate it and feed it back. Invalid specs will
                # raise here and be caught by the outer except.
                fo = units.GetFormatOptions(spec_id)

                if fo_dict.get("unit_type_id"):
                    fo.SetUnitTypeId(forge_id_from_str(fo_dict["unit_type_id"]))

                if fo_dict.get("accuracy") is not None:
                    try:
                        fo.Accuracy = fo_dict["accuracy"]
                    except Exception:
                        pass

                try:
                    fo.UseGrouping = fo_dict.get("use_grouping", False)
                except Exception:
                    pass
                try:
                    fo.SuppressLeadingZeros = fo_dict.get(
                        "suppress_leading_zeros", False)
                except Exception:
                    pass
                try:
                    fo.SuppressTrailingZeros = fo_dict.get(
                        "suppress_trailing_zeros", False)
                except Exception:
                    pass
                try:
                    fo.ShowMeasurementLineBelow = fo_dict.get(
                        "show_measurement_line_below", False)
                except Exception:
                    pass
                if fo_dict.get("prefix"):
                    try:
                        fo.Prefix = fo_dict["prefix"]
                    except Exception:
                        pass
                if fo_dict.get("suffix"):
                    try:
                        fo.Suffix = fo_dict["suffix"]
                    except Exception:
                        pass

                units.SetFormatOptions(spec_id, fo)

            except Exception:
                pass

        doc.SetUnits(units)
