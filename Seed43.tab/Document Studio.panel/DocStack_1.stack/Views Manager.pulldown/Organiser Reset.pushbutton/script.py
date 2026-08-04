# -*- coding: utf-8 -*-
import os
import json
from pyrevit import revit, DB, forms, script

# ── LOCATE CONFIG ─────────────────────────────────────────────────────────────
# view_organiser_config.json lives in the sibling View Organiser.pushbutton,
# so walk up two levels to the pulldown both buttons share:
#   .../Views Manager.pulldown/Organiser Reset.pushbutton/script.py
_split      = os.path.dirname(os.path.dirname(__file__))
CONFIG_PATH = os.path.join(
    _split, "View Organiser.pushbutton", "view_organiser_config.json")

NONE_LABEL = "- None -"

# ── HELPERS ───────────────────────────────────────────────────────────────────

def get_text_param_names(elements):
    names = set()
    for element in elements:
        try:
            for p in element.Parameters:
                try:
                    if (p.StorageType == DB.StorageType.String
                            and not p.IsReadOnly
                            and p.Definition
                            and p.Definition.Name):
                        names.add(p.Definition.Name)
                except Exception:
                    pass
        except Exception:
            pass
    return sorted(names)

def pick_folder_params(all_sheets, all_views):
    sample_views = [
        v for v in all_views
        if not v.IsTemplate
        and v.ViewType not in [DB.ViewType.DrawingSheet]
    ][:20]

    sheet_params = [NONE_LABEL] + get_text_param_names(all_sheets[:10])
    view_params  = [NONE_LABEL] + get_text_param_names(sample_views)

    sheet_choice = forms.ask_for_one_item(
        sheet_params,
        default=NONE_LABEL,
        prompt="Which parameter on SHEETS holds the folder grouping?\n"
               "(This value will be read from the sheet and copied to the view.)\n\n"
               "Select None to skip folder syncing.",
        title="View Organiser Config - Sheet Folder Parameter (source)"
    )
    if sheet_choice is None:
        script.exit()

    view_choice = forms.ask_for_one_item(
        view_params,
        default=NONE_LABEL,
        prompt="Which parameter on VIEWS should receive the folder value?\n\n"
               "Select None to skip folder syncing.",
        title="View Organiser Config - View Folder Parameter (destination)"
    )
    if view_choice is None:
        script.exit()

    CASE_OPTIONS = [
        "ALL CAPS  (e.g. GROUND FLOOR PLAN)",
        "Title Case  (e.g. Ground Floor Plan)",
        "Ignore  (leave as-is)",
    ]
    CASE_MAP = {
        CASE_OPTIONS[0]: "upper",
        CASE_OPTIONS[1]: "title",
        CASE_OPTIONS[2]: "ignore",
    }
    title_case_choice = forms.ask_for_one_item(
        CASE_OPTIONS,
        default=CASE_OPTIONS[0],
        prompt="How should 'Title on Sheet' values be cased?\n\n"
               "ALL CAPS   - forces every title to uppercase\n"
               "Title Case - capitalises the first letter of each word\n"
               "Ignore     - no case changes applied",
        title="View Organiser Config - Title on Sheet Casing"
    )
    if title_case_choice is None:
        script.exit()

    sheet_name_case_choice = forms.ask_for_one_item(
        CASE_OPTIONS,
        default=CASE_OPTIONS[0],
        prompt="How should Sheet Name values be cased?\n\n"
               "ALL CAPS   - forces every sheet name to uppercase\n"
               "Title Case - capitalises the first letter of each word\n"
               "Ignore     - no case changes applied",
        title="View Organiser Config - Sheet Name Casing"
    )
    if sheet_name_case_choice is None:
        script.exit()

    return (
        None if sheet_choice == NONE_LABEL else sheet_choice,
        None if view_choice  == NONE_LABEL else view_choice,
        CASE_MAP[title_case_choice],
        CASE_MAP[sheet_name_case_choice],
    )

# ── READ CURRENT CONFIG ───────────────────────────────────────────────────────

current = None
if os.path.isfile(CONFIG_PATH):
    try:
        with open(CONFIG_PATH, "r") as f:
            current = json.load(f)
    except Exception:
        pass

# ── CHOOSE ACTION ─────────────────────────────────────────────────────────────

if current:
    sheet_p       = current.get("sheet_folder_param") or "None"
    view_p        = current.get("view_folder_param")  or "None"
    title_case    = current.get("title_on_sheet_case", "upper")
    sname_case    = current.get("sheet_name_case", "upper")
    CASE_LABELS   = {"upper": "ALL CAPS", "title": "Title Case", "ignore": "Ignore"}
    summary = (
        "Current config:\n"
        "  Sheet param:        {}\n"
        "  View param:         {}\n"
        "  Title on Sheet:     {}\n"
        "  Sheet name casing:  {}\n\n"
        "What would you like to do?"
    ).format(sheet_p, view_p, CASE_LABELS.get(title_case, title_case), CASE_LABELS.get(sname_case, sname_case))
else:
    summary = "No config file found.\n\nWhat would you like to do?"

action = forms.CommandSwitchWindow.show(
    ["Reconfigure", "Delete"],
    message=summary,
    title="View Organiser Config"
)

if not action:
    script.exit()

# ── DELETE ────────────────────────────────────────────────────────────────────

if action == "Delete":
    if not os.path.isfile(CONFIG_PATH):
        forms.alert("No config file to delete.", title="View Organiser Config")
        script.exit()
    if not forms.alert(
            "Delete the saved config?\n\n"
            "View Organiser will prompt for parameters on next run.",
            title="Confirm Delete",
            yes=True,
            no=True):
        script.exit()
    try:
        os.remove(CONFIG_PATH)
        forms.alert("Config deleted.", title="View Organiser Config")
    except Exception as e:
        forms.alert(
            "Could not delete config:\n\n{}".format(str(e)),
            title="Error")
    script.exit()

# ── RECONFIGURE ───────────────────────────────────────────────────────────────

all_views = list(
    DB.FilteredElementCollector(revit.doc)
    .OfClass(DB.View)
    .WhereElementIsNotElementType()
    .ToElements()
)
all_sheets = list(
    DB.FilteredElementCollector(revit.doc)
    .OfClass(DB.ViewSheet)
    .ToElements()
)

sheet_param, view_param, title_on_sheet_case, sheet_name_case = pick_folder_params(all_sheets, all_views)

data = {
    "sheet_folder_param":  sheet_param,
    "view_folder_param":   view_param,
    "title_on_sheet_case": title_on_sheet_case,
    "sheet_name_case":     sheet_name_case,
}

CASE_LABELS = {"upper": "ALL CAPS", "title": "Title Case", "ignore": "Ignore"}

try:
    with open(CONFIG_PATH, "w") as f:
        json.dump(data, f, indent=2)
    forms.alert(
        "Config saved.\n\n"
        "  Sheet param:        {}\n"
        "  View param:         {}\n"
        "  Title on Sheet:     {}\n"
        "  Sheet name casing:  {}".format(
            sheet_param or "None",
            view_param  or "None",
            CASE_LABELS.get(title_on_sheet_case, title_on_sheet_case),
            CASE_LABELS.get(sheet_name_case, sheet_name_case)),
        title="View Organiser Config")
except Exception as e:
    forms.alert(
        "Could not save config:\n\n{}".format(str(e)),
        title="Error")
