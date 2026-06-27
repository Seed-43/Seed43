# -*- coding: utf-8 -*-
# toggle_concealment.py
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
