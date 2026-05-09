# -*- coding: utf-8 -*-
__title__  = "Set Rebar to Workset"
__author__  = "Fred da Silveira"
__doc__     = """
𝐕𝐄𝐑𝐒𝐈𝐎𝐍 𝟐𝟔𝟎𝟓𝟎𝟏
_____________________________________________________________________
Description:
Assigns every rebar element in the model to a chosen workset.

The tool is smart about finding the right workset:
- If exactly one workset with "rebar" in its name exists, it is used
  automatically with no prompt
- If several rebar worksets exist, you are shown a list of them plus
  an Other option to pick from all worksets
- If no rebar worksets exist, all worksets are listed with an option
  to create a new one on the spot
_____________________________________________________________________
How-to:
-> Run the tool in a workshared Revit model
-> Choose the target workset (auto-detected if only one rebar workset
   is found)
-> All rebar elements are moved to the selected workset
-> A message confirms how many elements were updated
_____________________________________________________________________
Notes:
- The model must be workshared (using Worksets) for this tool to work
- Creating a new workset requires a valid name
- If no rebar is found in the model, the tool will notify you and stop
_____________________________________________________________________
Last update:
- Initial release
_____________________________________________________________________
"""

from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import FilteredWorksetCollector, WorksetKind
import sys

doc   = revit.doc
uidoc = revit.uidoc


# ── GET WORKSETS ──────────────────────────────────────────────────────────────

worksets       = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
rebar_worksets = [ws for ws in worksets if "rebar" in ws.Name.lower()]


# ── FUNCTIONS ─────────────────────────────────────────────────────────────────

def select_or_create_workset():
    """Show all user worksets plus a Create New option."""
    options       = [ws.Name for ws in worksets] + ["Create New"]
    selected_name = forms.SelectFromList.show(
        options, button_name="Select or Create Workset")

    if not selected_name:
        forms.alert("Operation cancelled.")
        return None

    if selected_name == "Create New":
        new_name = forms.ask_for_string(
            prompt="Enter name for new Rebar workset:",
            default="Rebar"
        )
        if not new_name:
            forms.alert("Invalid name. Operation cancelled.")
            return None
        try:
            transaction = DB.Transaction(doc, "Create Rebar Workset")
            transaction.Start()
            new_ws = DB.Workset.Create(doc, new_name)
            transaction.Commit()
            return new_ws
        except Exception as e:
            if transaction.HasStarted():
                transaction.RollBack()
            forms.alert("Failed to create workset: {}".format(str(e)))
            return None

    return next(ws for ws in worksets if ws.Name == selected_name)


def get_target_workset():
    """Determine which workset to use based on what already exists."""
    num_rebar = len(rebar_worksets)

    if num_rebar == 1:
        return rebar_worksets[0]

    if num_rebar > 1:
        options       = [ws.Name for ws in rebar_worksets] + ["Other"]
        selected_name = forms.SelectFromList.show(
            options, button_name="Select Rebar Workset")
        if not selected_name:
            forms.alert("Operation cancelled.")
            return None
        if selected_name == "Other":
            return select_or_create_workset()
        return next(ws for ws in rebar_worksets if ws.Name == selected_name)

    # No rebar worksets found, go straight to full list
    return select_or_create_workset()


# ── GET TARGET WORKSET ────────────────────────────────────────────────────────

target_ws = get_target_workset()
if not target_ws:
    sys.exit()


# ── COLLECT REBAR ─────────────────────────────────────────────────────────────

rebar_elements = list(
    DB.FilteredElementCollector(doc)
      .OfClass(DB.Structure.Rebar)
      .WhereElementIsNotElementType()
)

if not rebar_elements:
    forms.alert("No rebar elements found in the model.")
    sys.exit()


# ── SET WORKSET ───────────────────────────────────────────────────────────────

count       = 0
transaction = DB.Transaction(doc, "Set Rebar to Workset")
transaction.Start()

try:
    for elem in rebar_elements:
        param = elem.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
        if param and not param.IsReadOnly:
            param.Set(target_ws.Id.IntegerValue)
            count += 1
    transaction.Commit()
    forms.alert(
        "Successfully set {} rebar elements to \"{}\" workset.".format(
            count, target_ws.Name)
    )
except Exception as e:
    if transaction.HasStarted():
        transaction.RollBack()
    forms.alert("Error: {}".format(str(e)))
