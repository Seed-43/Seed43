# -*- coding: utf-8 -*-
__title__  = "Toggle Rebar Visibility"
__author__  = "Fred da Silveira"
__doc__     = """
𝐕𝐄𝐑𝐒𝐈𝐎𝐍 𝟐𝟔𝟎𝟓𝟎𝟏
_____________________________________________________________________
Description:
Toggles rebar visibility between obscured and unobscured in the
current view.

The tool detects the current state of all visible rebar and decides
what to do:
- If some are obscured and some are not, all are made unobscured
- If all are already unobscured, all are made obscured
- If all are already obscured, all are made unobscured
_____________________________________________________________________
How-to:
-> Open a view that contains rebar
-> Run the tool
-> Rebar visibility switches automatically based on current state
-> A confirmation message shows what action was applied
_____________________________________________________________________
Notes:
- Works on the active view only
- Only rebar visible in the current view is affected
- Obscured means the rebar is shown as hidden behind concrete
  (dashed lines), unobscured means it is shown as solid lines
_____________________________________________________________________
Last update:
- Initial release
_____________________________________________________________________
"""

import clr
from Autodesk.Revit import DB
from pyrevit import forms, script

doc         = __revit__.ActiveUIDocument.Document
active_view = doc.ActiveView


# ── FUNCTIONS ─────────────────────────────────────────────────────────────────

def get_rebars_in_current_view():
    """Retrieve all rebar elements visible in the current view."""
    return (
        DB.FilteredElementCollector(doc, active_view.Id)
          .OfCategory(DB.BuiltInCategory.OST_Rebar)
          .WhereElementIsNotElementType()
          .ToElements()
    )


def check_rebar_states(rebars):
    """Return a count of obscured and unobscured rebar in the view."""
    obscured_count   = 0
    unobscured_count = 0
    for rebar in rebars:
        if rebar.IsUnobscuredInView(active_view):
            unobscured_count += 1
        else:
            obscured_count += 1
    return obscured_count, unobscured_count


def set_rebar_visibility(rebars, make_unobscured):
    """Set all rebar elements to obscured or unobscured."""
    for rebar in rebars:
        rebar.SetUnobscuredInView(active_view, make_unobscured)


# ── GET REBAR ─────────────────────────────────────────────────────────────────

rebars_in_view = get_rebars_in_current_view()

if not rebars_in_view:
    forms.alert("No rebars found in the current view.", title="Error")
    script.exit()


# ── DETERMINE ACTION ──────────────────────────────────────────────────────────

obscured_count, unobscured_count = check_rebar_states(rebars_in_view)

if obscured_count > 0 and unobscured_count > 0:
    # Mixed state, make all unobscured
    make_unobscured    = True
    action_description = "Set All Rebar Unobscured (Mixed State)"
elif unobscured_count == len(rebars_in_view):
    # All unobscured, make all obscured
    make_unobscured    = False
    action_description = "Set All Rebar Obscured"
else:
    # All obscured, make all unobscured
    make_unobscured    = True
    action_description = "Set All Rebar Unobscured"


# ── APPLY ─────────────────────────────────────────────────────────────────────

t = DB.Transaction(doc, action_description)
t.Start()

try:
    set_rebar_visibility(rebars_in_view, make_unobscured)
    t.Commit()
except Exception as e:
    t.RollBack()
    forms.alert("Error: {}".format(str(e)), title="Transaction Failed")
