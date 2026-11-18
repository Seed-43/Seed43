# -*- coding: utf-8 -*-
# toggle_concealment.py
#
# Toggles rebar in the current view between obscured and unobscured.

from pyrevit import revit, DB, forms, script

doc         = revit.doc
active_view = doc.ActiveView

try:
    from Snippets import _dialogs as sdlg
except Exception:
    sdlg = None

TITLE = "Toggle Rebar Visibility"

# ── DIALOGS ───────────────────────────────────────────────────────────────────

def _alert(message, title=TITLE, exitscript=False):
    """Themed popup via the shared Snippets dialog lib, falls back to
    pyRevit's default forms.alert if the shared lib isn't available."""
    if sdlg:
        sdlg.message(message, title=title)
    else:
        forms.alert(message, title=title)
    if exitscript:
        script.exit()

# ── FUNCTIONS ─────────────────────────────────────────────────────────────────

def get_rebars_in_current_view():
    """Retrieve all rebar elements visible in the current view."""
    return list(
        DB.FilteredElementCollector(doc, active_view.Id)
          .OfCategory(DB.BuiltInCategory.OST_Rebar)
          .WhereElementIsNotElementType()
    )


def sort_rebar_states(rebars):
    """Split rebar by whether it can report its obscured state.

    OST_Rebar covers Rebar, RebarInSystem and RebarContainer, which do not all
    expose IsUnobscuredInView. An AttributeError here used to kill the tool
    before it touched anything, so whatever cannot answer is counted and left
    alone rather than taking the run down with it.

    Returns (unobscured_count, supported, unsupported_count).
    """
    supported   = []
    unobscured  = 0
    unsupported = 0
    for rebar in rebars:
        try:
            if rebar.IsUnobscuredInView(active_view):
                unobscured += 1
            supported.append(rebar)
        except Exception:
            unsupported += 1
    return unobscured, supported, unsupported


def set_rebar_visibility(rebars, make_unobscured):
    """Set each rebar obscured or unobscured, returning (changed, failed).

    Per element, so one locked or unsupported bar cannot roll back the run.
    """
    changed = 0
    failed  = 0
    for rebar in rebars:
        try:
            rebar.SetUnobscuredInView(active_view, make_unobscured)
            changed += 1
        except Exception:
            failed += 1
    return changed, failed

# ── GET REBAR ─────────────────────────────────────────────────────────────────

rebars_in_view = get_rebars_in_current_view()

if not rebars_in_view:
    _alert("No rebars found in the current view.", exitscript=True)

unobscured_count, supported, unsupported_count = sort_rebar_states(
    rebars_in_view)

if not supported:
    _alert("None of the {} rebar element(s) in this view support the "
           "obscured/unobscured setting.".format(len(rebars_in_view)),
           exitscript=True)

obscured_count = len(supported) - unobscured_count

# ── DETERMINE ACTION ──────────────────────────────────────────────────────────

if obscured_count > 0 and unobscured_count > 0:
    # Mixed state, make all unobscured
    make_unobscured    = True
    action_description = "Set All Rebar Unobscured (Mixed State)"
elif unobscured_count == len(supported):
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
    changed, failed = set_rebar_visibility(supported, make_unobscured)
    t.Commit()
except Exception as e:
    # Guarded: if Commit itself threw, the transaction has already ended and
    # a blind RollBack throws a second exception over the top of the real one.
    if t.HasStarted():
        t.RollBack()
    _alert("Error: {}".format(str(e)), title="Transaction Failed",
           exitscript=True)

# ── REPORT ────────────────────────────────────────────────────────────────────

# The view redraw is the feedback on a clean run -- only speak up when part of
# the selection did not make it, which previously passed in silence.
if failed or unsupported_count:
    summary = ["Set {} rebar element(s) {}.".format(
        changed, "unobscured" if make_unobscured else "obscured")]
    if failed:
        summary.append(
            "{} could not be changed - most likely owned by another "
            "user.".format(failed))
    if unsupported_count:
        summary.append(
            "{} skipped - this rebar type does not support the "
            "obscured/unobscured setting.".format(unsupported_count))
    _alert("\n\n".join(summary))
