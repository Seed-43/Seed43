# -*- coding: utf-8 -*-
"""Shared helpers for tools that drive a 3D view's section box.

snippets.yaml entry:
  _sectionbox.py:
    description: >
      Helpers for tools that drive a 3D view's section box - model extents,
      the 3D-view guard, and applying a box. Used by Model Slice.

    functions:
      mm_to_ft:        Convert millimetres to Revit's internal feet.
      require_3d_view: Return the active view if it can take a section box, else None.
      model_extents:   Return a BoundingBoxXYZ enclosing every model element.
      datum_extents:   Plan extent of the grids and height range of the levels, padded.
      apply_section_box: Turn the section box on and set it, in one transaction.
"""

from Autodesk.Revit.DB import (
    BoundingBoxXYZ, BuiltInCategory, CategoryType, ElementId,
    FilteredElementCollector, Grid,
    ImportInstance, Level, Line, Transaction, View, View3D, XYZ,
)

__all__ = ["mm_to_ft", "require_3d_view", "model_extents", "datum_extents",
           "apply_section_box"]

_MM_PER_FT = 304.8
_CAMERAS = ElementId(BuiltInCategory.OST_Cameras)


def mm_to_ft(mm):
    """Millimetres to Revit's internal feet."""
    return mm / _MM_PER_FT


# ── VIEW GUARD ──────────────────────────────────────────────────────────────

def require_3d_view(doc):
    """Return the active view if it can take a section box, else None.

    Section boxes are a View3D feature, and a template cannot be modified
    directly, so both are rejected before anything else runs.
    """
    view = doc.ActiveView
    if not isinstance(view, View3D):
        return None
    if view.IsTemplate:
        return None
    return view


# ── MODEL EXTENTS ───────────────────────────────────────────────────────────

def _is_model_element(element):
    """True for a physical model element worth measuring the building by.

    Everything else that has a bounding box is out. Measured on 4286 A Block
    Remedial (2026-10-01), where counting everything gave a 609 x 2023 m
    box around a building under 100 m across:
      - uncategorised internal elements, eight of them 1.5 km off site
      - cameras (the {3D} camera alone spanned 609 x 419 x 257 m) and
        views, which have a category but are not model geometry
      - view-specific detail items and lines, drawn at drafting-view
        coordinates that have nothing to do with the model
      - CAD imports, which are often a whole site plan
    """
    if element.ViewSpecific:
        return False
    if isinstance(element, (View, ImportInstance)):
        return False
    cat = element.Category
    if cat is None or cat.CategoryType != CategoryType.Model:
        return False
    # Cameras are filed as a Model category, and a 3D view's camera box
    # spans its whole frustum: 609 x 419 x 257 m on its own in that model.
    return not cat.Id.Equals(_CAMERAS)

def model_extents(doc, view=None):
    """Return a BoundingBoxXYZ enclosing every model element, or None.

    Walks elements rather than trusting any single call, because there is no
    reliable "whole model bounding box" in the API. Only physical model
    elements count (see _is_model_element): plenty of other things carry a
    bounding box, and one of them far away swamps the rest.

    Pass a view to measure only what that view shows; omit it to measure the
    model. Note a section box already applied to the view will clip the
    per-view answer, which is why callers wanting true extents pass nothing.
    """
    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()

    min_x = min_y = min_z = None
    max_x = max_y = max_z = None

    for element in collector:
        if not _is_model_element(element):
            continue
        try:
            box = element.get_BoundingBox(view)
        except Exception:
            box = None
        if box is None:
            continue
        if min_x is None:
            min_x, min_y, min_z = box.Min.X, box.Min.Y, box.Min.Z
            max_x, max_y, max_z = box.Max.X, box.Max.Y, box.Max.Z
            continue
        min_x = min(min_x, box.Min.X)
        min_y = min(min_y, box.Min.Y)
        min_z = min(min_z, box.Min.Z)
        max_x = max(max_x, box.Max.X)
        max_y = max(max_y, box.Max.Y)
        max_z = max(max_z, box.Max.Z)

    if min_x is None:
        return None

    extents = BoundingBoxXYZ()
    extents.Min = XYZ(min_x, min_y, min_z)
    extents.Max = XYZ(max_x, max_y, max_z)
    return extents


# ── DATUM EXTENTS ───────────────────────────────────────────────────────────

def _grid_points(grid):
    """Points along a grid: the two ends of a line, 11 samples of an arc."""
    curve = grid.Curve
    if isinstance(curve, Line):
        return [curve.GetEndPoint(0), curve.GetEndPoint(1)]
    return [curve.Evaluate(i / 10.0, True) for i in range(11)]


def datum_extents(doc, pad_ft):
    """The building as the grids and levels define it, padded by pad_ft.

    Returns (BoundingBoxXYZ, notes). X and Y come from the grids, Z from
    the lowest to the highest level. model_extents takes in everything,
    so one stray survey point or old import far off site blows the box up
    to the size of the county; the datums are where the building is.

    Falls back to model_extents per axis when there is nothing to measure
    by: plan when the model has no grids, height when it has fewer than two
    levels. Each fallback adds a line to notes so the caller can say so.
    Returns (None, notes) if even the fallback finds no geometry.
    """
    notes = []

    xs, ys = [], []
    for grid in FilteredElementCollector(doc).OfClass(Grid):
        try:
            for p in _grid_points(grid):
                xs.append(p.X)
                ys.append(p.Y)
        except Exception:
            continue  # a grid without a readable curve adds nothing

    # ProjectElevation, never Elevation: Elevation follows the level type's
    # Elevation Base, so on a Survey Point base it is an RL (55.4 m for a
    # ground floor at 0.0 on 4286 A Block) while boxes are in project
    # coordinates.
    elevations = [lv.ProjectElevation for lv in
                  FilteredElementCollector(doc).OfClass(Level)]

    model = None
    if not xs or len(elevations) < 2:
        model = model_extents(doc)
        if model is None:
            return None, notes

    if xs:
        min_x, max_x = min(xs) - pad_ft, max(xs) + pad_ft
        min_y, max_y = min(ys) - pad_ft, max(ys) + pad_ft
    else:
        min_x, max_x = model.Min.X, model.Max.X
        min_y, max_y = model.Min.Y, model.Max.Y
        notes.append("No grids in this model, so the plan size is the "
                     "whole model.")

    if len(elevations) >= 2:
        min_z, max_z = min(elevations) - pad_ft, max(elevations) + pad_ft
    else:
        min_z, max_z = model.Min.Z, model.Max.Z
        notes.append("Fewer than two levels, so the height is the whole "
                     "model.")

    box = BoundingBoxXYZ()
    box.Min = XYZ(min_x, min_y, min_z)
    box.Max = XYZ(max_x, max_y, max_z)
    return box, notes


# ── APPLY ───────────────────────────────────────────────────────────────────

def apply_section_box(view, box, transaction_name="Set Section Box"):
    """Turn the section box on and set it, in one transaction.

    IsSectionBoxActive is set first: setting the box on a view whose section
    box is off leaves it stored but invisible, which looks like the tool did
    nothing.
    """
    t = Transaction(view.Document, transaction_name)
    t.Start()
    try:
        view.IsSectionBoxActive = True
        view.SetSectionBox(box)
        t.Commit()
        return True
    except Exception:
        if t.HasStarted():
            t.RollBack()
        return False
