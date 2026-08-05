# -*- coding: utf-8 -*-
# selection_box_grid.py
# "Selection Box: Grid"
# "Seed43"
# """
# Section-box a thin vertical slice of the model along a chosen grid.
# """
from pyrevit import revit, DB, forms, script

# ── [LIB] Snippets/_selection.py, Snippets/_sectionbox.py ───────────────────
from Snippets._selection import choose_datum
from Snippets._sectionbox import (
    mm_to_ft, require_3d_view, model_extents, apply_section_box,
)

doc = revit.doc
uidoc = revit.uidoc

# ── CONSTANTS ───────────────────────────────────────────────────────────────

# Cut this far either side of the grid line.
OFFSET_MM = 500.0


# ── GRIDS ───────────────────────────────────────────────────────────────────

def get_grids():
    """Every grid in the model, sorted the way the grid bubbles read."""
    grids = list(DB.FilteredElementCollector(doc)
                 .OfClass(DB.Grid)
                 .WhereElementIsNotElementType())
    return sorted(grids, key=lambda g: _natural_key(g.Name))


def _natural_key(name):
    """Sort A1, A2, A10 in that order rather than A1, A10, A2."""
    import re
    return [int(p) if p.isdigit() else p.lower()
            for p in re.split(r'(\d+)', str(name))]


# ── BOX CONSTRUCTION ────────────────────────────────────────────────────────

def straight_grid_box(grid, extents, offset):
    """A box rotated to lie along a straight grid.

    The box is built in the grid's own coordinates - X along the line, Y
    across it, Z vertical - and a Transform carries it back into model space.
    Min and Max are relative to that Transform's origin, NOT world coordinates,
    which is the part that catches people out.
    """
    curve = grid.Curve
    start = curve.GetEndPoint(0)
    end = curve.GetEndPoint(1)

    along = (end - start).Normalize()
    across = DB.XYZ.BasisZ.CrossProduct(along).Normalize()

    half_length = start.DistanceTo(end) / 2.0
    half_height = (extents.Max.Z - extents.Min.Z) / 2.0
    mid_z = (extents.Max.Z + extents.Min.Z) / 2.0

    transform = DB.Transform.Identity
    transform.Origin = DB.XYZ((start.X + end.X) / 2.0,
                              (start.Y + end.Y) / 2.0,
                              mid_z)
    transform.BasisX = along
    transform.BasisY = across
    transform.BasisZ = along.CrossProduct(across)

    box = DB.BoundingBoxXYZ()
    box.Transform = transform
    box.Min = DB.XYZ(-half_length, -offset, -half_height)
    box.Max = DB.XYZ(half_length, offset, half_height)
    return box


def curved_grid_box(grid, extents, offset):
    """Axis-aligned fallback for an arc grid.

    A rotated slice has no meaning along a curve, so this boxes the arc's own
    extents plus the offset. Wider than the straight case, but still far less
    than the whole model.
    """
    curve = grid.Curve
    xs, ys = [], []
    # Sample the curve rather than trusting a bounding box, which an arc's
    # tessellation does not always give tightly.
    for i in range(11):
        p = curve.Evaluate(i / 10.0, True)
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
                    "Open a 3D view and try again.",
                    title="Selection Box: Grid")
        script.exit()

    grids = get_grids()
    if not grids:
        forms.alert("No grids found in this model.",
                    title="Selection Box: Grid")
        script.exit()

    # Clicking is the natural route for grids - the name rarely tells you
    # which one you want, the position does. Selecting one in a plan view
    # before switching to 3D also works, since the selection survives.
    grid = choose_datum(uidoc, doc, revit, forms, DB.Grid, grids,
                        title="Section box around which grid?",
                        pick_prompt="Click a grid line")
    if grid is None:
        script.exit()

    # Against the model, not the view - see the Level tool for why.
    extents = model_extents(doc)
    if extents is None:
        forms.alert("Could not measure the model extents - nothing with "
                    "geometry was found.", title="Selection Box: Grid")
        script.exit()

    offset = mm_to_ft(OFFSET_MM)
    if isinstance(grid.Curve, DB.Line):
        box = straight_grid_box(grid, extents, offset)
        shape = ""
    else:
        box = curved_grid_box(grid, extents, offset)
        shape = "\n\nThis grid is curved, so the box is aligned to the model " \
                "axes around it rather than rotated along it."

    if not apply_section_box(view, box, "Selection Box: Grid"):
        forms.alert("Could not set the section box on this view.",
                    title="Selection Box: Grid")
        script.exit()

    forms.alert("Section box set {:.0f}mm either side of grid {}.{}".format(
        OFFSET_MM, grid.Name, shape), title="Selection Box: Grid")


if __name__ == "__main__":
    main()
