# -*- coding: utf-8 -*-
# "Sheet Scale"
# "Seed43"
# """
# Writes the view scales of each As indicated sheet into A1_Scale / A3_Scale
# and turns the title block's scale override on; clears them and turns it off
# on every other sheet. Based on the XRT Scale Finder, fixed for title blocks
# that are natively A3.
# """

# ── IMPORTS ────────────────────────────────────────────────────────────────

import re
from collections import defaultdict

from Autodesk.Revit.DB import (
    BuiltInCategory, BuiltInParameter, FilteredElementCollector,
    StorageType, Transaction, ViewSheet, ViewType, Viewport
)
from pyrevit import forms, script

# ── CONSTANTS ──────────────────────────────────────────────────────────────

A1_PARAM      = "A1_Scale"
A3_PARAM      = "A3_Scale"
NO_SCALE_TEXT = "N.T.S"
SEPARATOR     = " "
# NOTE: a viewport type whose name holds any of these shows no scale on the
# sheet (View Title - A3 Plan No Scale, A3 Auto Line (N.T.S), No Title,
# View Title - None), so its view is left out.
SKIP_VP_TEXT  = ("no scale", "n.t.s", "no title", "none")
SKIP_VIEWS    = (ViewType.Legend,)

FT_TO_MM      = 304.8
A3_LONG_MM    = 420.0
A1_LONG_MM    = 841.0
SIZE_TOL_MM   = 15.0            # borders are rarely drawn to the exact ISO size

AS_INDICATED  = "as indicated"  # Revit's sheet Scale text when views differ

# NOTE: matches "A3 Scale Overide" (the spelling in the Nagel A3 border) as well
# as "A1 Scale Override", so either family's toggle is found.
TOGGLE_RE = re.compile(r"^\s*(A\d)\s*Scale\s*Overr?ide\s*$", re.IGNORECASE)

# ── HELPERS ────────────────────────────────────────────────────────────────

uidoc  = __revit__.ActiveUIDocument
doc    = uidoc.Document
output = script.get_output()


# --- Sheet size ---
def sheet_size_name(titleblock):
    """Return 'A1', 'A3' or '' from the title block's sheet width/height."""
    if titleblock is None:
        return ""
    w = titleblock.get_Parameter(BuiltInParameter.SHEET_WIDTH)
    h = titleblock.get_Parameter(BuiltInParameter.SHEET_HEIGHT)
    if w is None or h is None:
        return ""
    long_mm = max(w.AsDouble(), h.AsDouble()) * FT_TO_MM
    if abs(long_mm - A3_LONG_MM) <= SIZE_TOL_MM:
        return "A3"
    if abs(long_mm - A1_LONG_MM) <= SIZE_TOL_MM:
        return "A1"
    return ""


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


# --- Title block toggle ---
def find_toggle(titleblock, size):
    """The Yes/No scale override toggle, preferring the one named for this size."""
    fallback = None
    for p in titleblock.Parameters:
        if p.StorageType != StorageType.Integer or p.IsReadOnly:
            continue
        m = TOGGLE_RE.match(p.Definition.Name)
        if not m:
            continue
        if m.group(1).upper() == size:
            return p
        fallback = fallback or p
    return fallback


def set_toggle(titleblock, size, on):
    """Set the override toggle. Returns (label for the report, changed)."""
    p = find_toggle(titleblock, size) if titleblock else None
    if p is None:
        return "-", False
    want = 1 if on else 0
    changed = p.AsInteger() != want
    if changed:
        p.Set(want)
    return "{} = {}".format(p.Definition.Name, "Yes" if on else "No"), changed


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
            size   = sheet_size_name(tb)
            scales = sheet_scales(vps_by_sheet.get(sheet.Id, []), vp_type_cache)

            # NOTE: only sheets where Revit's own Scale reads As indicated get a
            # value and the override on - every other sheet is cleared and the
            # override goes off, so the title block shows Revit's Scale. The
            # family can't do this itself: its formulas can't read sheet
            # parameters or test text. Scales are written 1:1 into the
            # parameter for the sheet's own size; a sheet of unknown size gets
            # both. The other size's parameter is cleared - old XRT runs left
            # stale A1_Scale values (e.g. 1:20 on 1:10 A3 sheets).
            revit_scale  = sheet_revit_scale(sheet)
            as_indicated = AS_INDICATED in revit_scale.lower()
            text = scale_string(scales) if as_indicated else ""
            changed = False
            changed |= set_text(sheet, A1_PARAM, "" if size == "A3" else text)
            changed |= set_text(sheet, A3_PARAM, "" if size == "A1" else text)
            toggle_text, toggled = set_toggle(tb, size, as_indicated)
            changed |= toggled

            rows.append([sheet.SheetNumber, sheet.Name, size or "?", revit_scale,
                         text or "-", toggle_text, "Yes" if changed else ""])
        t.Commit()
    except Exception:
        t.RollBack()
        raise
    return rows

# ── UI / ENTRY POINT ───────────────────────────────────────────────────────

if __name__ == "__main__":
    choice = forms.CommandSwitchWindow.show(["Current Sheet", "All Sheets"],
                                            message="Sheet Scale")
    if choice:
        sheets = pick_sheets(choice)
        if not sheets:
            forms.alert("Open a sheet or select sheets in the Project Browser first.")
        else:
            rows = update_sheets(sorted(sheets, key=lambda s: s.SheetNumber))
            output.print_table(
                rows,
                columns=["Sheet", "Name", "Size", "Revit Scale", "Written",
                         "Override", "Changed"],
                title="Sheet Scale - {} sheet(s), {} changed".format(
                    len(rows), sum(1 for r in rows if r[6])))
