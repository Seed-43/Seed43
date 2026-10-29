# -*- coding: utf-8 -*-
# "Sheet Scale"
# "Seed43"
# """
# Writes the view scales of each As indicated sheet into Scale_Override and
# turns the title block's Override Scale on; clears it and turns the switch
# off on every other sheet. Title blocks get the label and switch from Set Up
# Title Blocks (title_blocks.py). Based on the XRT Scale Finder.
# """

# ── IMPORTS ────────────────────────────────────────────────────────────────

from collections import defaultdict

from Autodesk.Revit.DB import (
    BuiltInCategory, BuiltInParameter, FilteredElementCollector,
    StorageType, Transaction, ViewSheet, ViewType, Viewport
)
from pyrevit import forms, script

# ── CONSTANTS ──────────────────────────────────────────────────────────────

OVERRIDE_TEXT = "Scale_Override"   # sheet parameter the title block shows
TOGGLE        = "Override Scale"   # title block switch (see title_blocks.py)
NO_SCALE_TEXT = "N.T.S"
SEPARATOR     = " "
# NOTE: a viewport type whose name holds any of these shows no scale on the
# sheet (View Title - A3 Plan No Scale, A3 Auto Line (N.T.S), No Title,
# View Title - None), so its view is left out.
SKIP_VP_TEXT  = ("no scale", "n.t.s", "no title", "none")
SKIP_VIEWS    = (ViewType.Legend,)

AS_INDICATED  = "as indicated"  # Revit's sheet Scale text when views differ
NOT_SET_UP    = "title block not set up - run Set Up Title Blocks"

# ── HELPERS ────────────────────────────────────────────────────────────────

uidoc  = __revit__.ActiveUIDocument
doc    = uidoc.Document
output = script.get_output()


# --- Scales ---
def scale_string(scales):
    """Join scales small to big as '1:10 1:20', or N.T.S when there are none."""
    if not scales:
        return NO_SCALE_TEXT
    return SEPARATOR.join("1:{}".format(s) for s in sorted(scales))


def viewport_type_name(vp, cache):
    type_id = vp.GetTypeId()
    if type_id not in cache:
        vpt = doc.GetElement(type_id)
        p = vpt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM) if vpt else None
        cache[type_id] = (p.AsString() or "") if p else ""
    return cache[type_id]


def sheet_scales(viewports, vp_type_cache):
    """Distinct view scales on one sheet, legends and no-scale viewports left out."""
    scales = set()
    for vp in viewports:
        view = doc.GetElement(vp.ViewId)
        if view is None or view.ViewType in SKIP_VIEWS:
            continue
        vp_name = viewport_type_name(vp, vp_type_cache).lower()
        if any(text in vp_name for text in SKIP_VP_TEXT):
            continue
        try:
            scale = view.Scale
        except Exception:
            continue
        if scale and scale > 0:
            scales.add(scale)
    return scales


def sheet_revit_scale(sheet):
    """Revit's own sheet Scale text, e.g. '1 : 10' or 'As indicated'."""
    p = sheet.get_Parameter(BuiltInParameter.SHEET_SCALE)
    return ((p.AsString() if p else "") or "").strip()


# --- Title block switch ---
def set_override(titleblock, on):
    """Set the title block's Override Scale. Returns (label for the report,
    changed), or None when the title block is not set up."""
    p = titleblock.LookupParameter(TOGGLE) if titleblock else None
    if p is None or p.IsReadOnly or p.StorageType != StorageType.Integer:
        return None
    want = 1 if on else 0
    changed = p.AsInteger() != want
    if changed:
        p.Set(want)
    return "{} = {}".format(TOGGLE, "Yes" if on else "No"), changed


def set_text(element, name, value):
    """Set a text parameter. Returns True when the value changed."""
    p = element.LookupParameter(name)
    if p is None or p.IsReadOnly or p.StorageType != StorageType.String:
        return False
    if (p.AsString() or "") == value:
        return False
    p.Set(value)
    return True

# ── CORE LOGIC ─────────────────────────────────────────────────────────────

def pick_sheets(choice):
    if choice == "All Sheets":
        return [s for s in FilteredElementCollector(doc).OfClass(ViewSheet)
                if not s.IsPlaceholder]
    picked = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    picked = [s for s in picked if isinstance(s, ViewSheet)]
    if picked:
        return picked
    if isinstance(doc.ActiveView, ViewSheet):
        return [doc.ActiveView]
    return []


def update_sheets(sheets):
    """Write the scale parameters and toggles. Returns report rows."""
    # NOTE: everything is collected up front - nesting collectors inside the
    # sheet loop silently cuts the loop short.
    vps_by_sheet = defaultdict(list)
    for vp in FilteredElementCollector(doc).OfClass(Viewport).ToElements():
        vps_by_sheet[vp.SheetId].append(vp)
    tb_by_sheet = {}
    for tb in (FilteredElementCollector(doc)
               .OfCategory(BuiltInCategory.OST_TitleBlocks)
               .WhereElementIsNotElementType().ToElements()):
        tb_by_sheet.setdefault(tb.OwnerViewId, tb)

    vp_type_cache = {}
    rows = []
    t = Transaction(doc, "Sheet Scale")
    t.Start()
    try:
        for sheet in sheets:
            tb     = tb_by_sheet.get(sheet.Id)
            scales = sheet_scales(vps_by_sheet.get(sheet.Id, []), vp_type_cache)

            # NOTE: only sheets where Revit's own Scale reads As indicated get
            # the scales and Override Scale on - every other sheet is cleared
            # and the switch goes off, so the title block shows Revit's Scale.
            # The family can't do this itself: its formulas can't read sheet
            # parameters or test text.
            revit_scale  = sheet_revit_scale(sheet)
            as_indicated = AS_INDICATED in revit_scale.lower()
            text = scale_string(scales) if as_indicated else ""
            override = set_override(tb, as_indicated)
            if override is None:
                rows.append([sheet.SheetNumber, sheet.Name, revit_scale,
                             "-", NOT_SET_UP, ""])
                continue
            toggle_text, changed = override
            changed |= set_text(sheet, OVERRIDE_TEXT, text)
            rows.append([sheet.SheetNumber, sheet.Name, revit_scale,
                         text or "-", toggle_text, "Yes" if changed else ""])
        t.Commit()
    except Exception:
        t.RollBack()
        raise
    return rows

# ── UI / ENTRY POINT ───────────────────────────────────────────────────────

SET_UP = "Set Up Title Blocks"


def run_setup():
    from title_blocks import DONOR_VERSION, set_up_title_blocks
    if int(doc.Application.VersionNumber) < DONOR_VERSION:
        forms.alert("Setting up title blocks needs Revit {} or later: the "
                    "Scale_Override label is copied from Sheet Scale.rfa, "
                    "saved in {}.".format(DONOR_VERSION, DONOR_VERSION))
        return
    if not forms.alert(
            "Every title block family placed on a sheet gets a Scale_Override "
            "label on its scale label and an Override Scale switch, and is "
            "loaded back into the project (the families are not saved). "
            "Scale_Override is added to the project's sheets if missing.\n\n"
            "Save the project first. Continue?", yes=True, no=True):
        return
    # NOTE: no progress bar - opening and closing family documents loses
    # pyRevit's handle on the Revit window, and its updates then fail.
    rows, added = set_up_title_blocks(doc)
    output.print_table(
        [list(r) for r in rows], columns=["Title block", "Result"],
        title="Sheet Scale - title blocks{}".format(
            ", Scale_Override added to sheets" if added else ""))


if __name__ == "__main__":
    choice = forms.CommandSwitchWindow.show(
        ["Current Sheet", "All Sheets", SET_UP], message="Sheet Scale")
    if choice == SET_UP:
        run_setup()
    elif choice:
        sheets = pick_sheets(choice)
        if not sheets:
            forms.alert("Open a sheet or select sheets in the Project Browser first.")
        else:
            rows = update_sheets(sorted(sheets, key=lambda s: s.SheetNumber))
            output.print_table(
                rows,
                columns=["Sheet", "Name", "Revit Scale", "Scale_Override",
                         "Override Scale", "Changed"],
                title="Sheet Scale - {} sheet(s), {} changed".format(
                    len(rows), sum(1 for r in rows if r[5])))
