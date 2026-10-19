# -*- coding: utf-8 -*-
"""Section box a thin slice of the model at a level or along a grid.

One window lists "Select in model", every level and every grid; type to
filter (a level or grid name) and Enter, or pick the first entry and click
a level or grid line. A level gives a horizontal slice, a grid a vertical
one. Replaces the old Model Slice pulldown (Level Slice + Grid Slice).
"""
import re

from pyrevit import revit, DB, forms, script
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

# ── [LIB] Snippets/_sectionbox.py ───────────────────────────────────────────
from Snippets._sectionbox import (
    mm_to_ft, require_3d_view, datum_extents, apply_section_box,
)

doc = revit.doc
uidoc = revit.uidoc

TITLE = "Model Slice"

# ── CONSTANTS ───────────────────────────────────────────────────────────────

# Cut this far either side of the level or grid line. Enough to catch a
# floor build-up and the slab under it without the storey above.
OFFSET_MM = 500.0

# The box runs this far past the outermost grids, below the lowest level,
# above the highest, and past each end of a grid, so edge members, footings
# and roof steel stay in.
PAD_MM = 1000.0

PICK_IN_MODEL = u"Select in model..."
LEVEL_LABEL = u"Level  {}"
GRID_LABEL = u"Grid  {}"


# ── DATUMS ──────────────────────────────────────────────────────────────────

def _natural_key(name):
    """Sort Level 2 before Level 10, and grid A2 before A10."""
    return [int(p) if p.isdigit() else p.lower()
            for p in re.split(r"(\d+)", unicode(name))]


def get_levels():
    """Every level, lowest first."""
    return sorted(DB.FilteredElementCollector(doc).OfClass(DB.Level),
                  key=lambda lv: lv.ProjectElevation)


def get_grids():
    """Every grid, in the order the bubbles read."""
    grids = (DB.FilteredElementCollector(doc).OfClass(DB.Grid)
             .WhereElementIsNotElementType())
    return sorted(grids, key=lambda g: _natural_key(g.Name))


class DatumFilter(ISelectionFilter):
    """Levels and grids only."""

    def AllowElement(self, element):
        return isinstance(element, (DB.Level, DB.Grid))

    def AllowReference(self, reference, point):
        return False


def pick_in_model():
    """A clicked level or grid, or None if cancelled or none is clickable.

    Levels and grids are hidden in most 3D views unless turned on in
    Visibility/Graphics, so a click is not always possible. Typing the name
    in the window, or selecting it in a plan first, always works.
    """
    prompt = u"{}: click a level or grid line. ESC to cancel.".format(TITLE)
    try:
        with forms.WarningBar(title=prompt):
            ref = uidoc.Selection.PickObject(
                ObjectType.Element, DatumFilter(), prompt)
        return doc.GetElement(ref.ElementId)
    except OperationCanceledException:
        return None


def choose_datum(levels, grids):
    """The level or grid to slice at, or None if cancelled.

    A single level or grid already selected is used with no window at all;
    selecting a grid in plan before switching to 3D is the quick route.
    """
    selected = [el for el in revit.get_selection()
                if isinstance(el, (DB.Level, DB.Grid))]
    if len(selected) == 1:
        return selected[0]

    by_label = {}
    for lv in levels:
        by_label[LEVEL_LABEL.format(lv.Name)] = lv
    for g in grids:
        by_label[GRID_LABEL.format(g.Name)] = g
    options = ([PICK_IN_MODEL]
               + [LEVEL_LABEL.format(lv.Name) for lv in levels]
               + [GRID_LABEL.format(g.Name) for g in grids])

    chosen = forms.CommandSwitchWindow.show(
        options, message=u"{}: type a level or grid name, or pick one"
                         .format(TITLE))
    if not chosen:
        return None
    if chosen != PICK_IN_MODEL:
        return by_label.get(chosen)

    datum = pick_in_model()
    if datum is None:
        forms.alert("Nothing picked. Levels and grids are often hidden in "
                    "3D views: run it again and type the name instead, or "
                    "select the grid in a plan view first.", title=TITLE)
    return datum


# ── BOX CONSTRUCTION ────────────────────────────────────────────────────────

def level_box(level, extents, offset):
    """Plan from the grids (or model), height the level plus/minus offset.

    ProjectElevation, not Elevation: with a Survey Point elevation base,
    Elevation is an RL and the box lands tens of metres above the model.
    """
    z = level.ProjectElevation
    box = DB.BoundingBoxXYZ()
    box.Min = DB.XYZ(extents.Min.X, extents.Min.Y, z - offset)
    box.Max = DB.XYZ(extents.Max.X, extents.Max.Y, z + offset)
    return box


def straight_grid_box(grid, extents, offset, pad):
    """A box rotated to lie along a straight grid, pad past each end.

    Built in the grid's own coordinates (X along, Y across, Z up) and a
    Transform carries it back into model space. Min and Max are relative to
    that Transform's origin, NOT world coordinates.
    """
    start = grid.Curve.GetEndPoint(0)
    end = grid.Curve.GetEndPoint(1)

    along = (end - start).Normalize()
    across = DB.XYZ.BasisZ.CrossProduct(along).Normalize()

    half_length = start.DistanceTo(end) / 2.0 + pad
    half_height = (extents.Max.Z - extents.Min.Z) / 2.0
    mid_z = (extents.Max.Z + extents.Min.Z) / 2.0

    transform = DB.Transform.Identity
    transform.Origin = DB.XYZ((start.X + end.X) / 2.0,
                              (start.Y + end.Y) / 2.0, mid_z)
    transform.BasisX = along
    transform.BasisY = across
    transform.BasisZ = along.CrossProduct(across)

    box = DB.BoundingBoxXYZ()
    box.Transform = transform
    box.Min = DB.XYZ(-half_length, -offset, -half_height)
    box.Max = DB.XYZ(half_length, offset, half_height)
    return box


def curved_grid_box(grid, extents, offset):
    """Axis-aligned fallback for an arc grid: the sampled arc, padded.

    A rotated slice has no meaning along a curve. Sampled rather than
    boxed, because an arc's tessellated box is not always tight.
    """
    xs, ys = [], []
    for i in range(11):
        p = grid.Curve.Evaluate(i / 10.0, True)
        xs.append(p.X)
        ys.append(p.Y)
    box = DB.BoundingBoxXYZ()
    box.Min = DB.XYZ(min(xs) - offset, min(ys) - offset, extents.Min.Z)
    box.Max = DB.XYZ(max(xs) + offset, max(ys) + offset, extents.Max.Z)
    return box


# ── MAIN ────────────────────────────────────────────────────────────────────

def main():
    view = require_3d_view(doc)
    if view is None:
        forms.alert("Section boxes only exist in 3D views.\n\n"
                    "Open a 3D view and try again.", title=TITLE)
        script.exit()

    levels = get_levels()
    grids = get_grids()
    if not levels and not grids:
        forms.alert("No levels or grids found in this model.", title=TITLE)
        script.exit()

    datum = choose_datum(levels, grids)
    if datum is None:
        script.exit()

    # From the grids and levels, not the view (which may already be boxed,
    # shrinking the result every run) and not the whole model.
    pad = mm_to_ft(PAD_MM)
    extents, notes = datum_extents(doc, pad)
    if extents is None:
        forms.alert("Nothing with geometry was found to size the box by.",
                    title=TITLE)
        script.exit()

    offset = mm_to_ft(OFFSET_MM)
    shape = ""
    if isinstance(datum, DB.Level):
        box = level_box(datum, extents, offset)
        what = datum.Name
        # the height is the level's own, only the plan fallback matters
        notes = [n for n in notes if "grids" in n]
    else:
        if isinstance(datum.Curve, DB.Line):
            box = straight_grid_box(datum, extents, offset, pad)
        else:
            box = curved_grid_box(datum, extents, offset)
            shape = ("\n\nThis grid is curved, so the box is aligned to the "
                     "model axes around it rather than rotated along it.")
        what = u"grid {}".format(datum.Name)
        # the plan is the grid's own, only the height fallback matters
        notes = [n for n in notes if "levels" in n]

    if not apply_section_box(view, box, TITLE):
        forms.alert("Could not set the section box on this view.",
                    title=TITLE)
        script.exit()

    forms.alert(u"Section box set {:.0f}mm either side of {}.{}{}".format(
        OFFSET_MM, what, shape, u"".join(u"\n\n" + n for n in notes)),
        title=TITLE)


if __name__ == "__main__":
    main()
