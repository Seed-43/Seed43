# -*- coding: utf-8 -*-
"""Shared helpers for tools that drive a 3D view's section box.

snippets.yaml entry:
  _sectionbox.py:
    description: >
      Helpers for tools that drive a 3D view's section box - model extents,
      the 3D-view guard, and applying a box. Used by Grid Slice and
      Level Slice.
    functions:
      mm_to_ft:        Convert millimetres to Revit's internal feet.
      require_3d_view: Return the active view if it can take a section box, else None.
      model_extents:   Return a BoundingBoxXYZ enclosing every model element.
      apply_section_box: Turn the section box on and set it, in one transaction.
"""

from Autodesk.Revit.DB import (
    BoundingBoxXYZ, FilteredElementCollector, Transaction, View3D, XYZ,
)

__all__ = ["mm_to_ft", "require_3d_view", "model_extents", "apply_section_box"]

_MM_PER_FT = 304.8


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

def model_extents(doc, view=None):
    """Return a BoundingBoxXYZ enclosing every model element, or None.

    Walks elements rather than trusting any single call, because there is no
    reliable "whole model bounding box" in the API. Elements with no geometry
    (view-specific annotation, groups) simply return no box and are skipped.

    Pass a view to measure only what that view shows; omit it to measure the
    model. Note a section box already applied to the view will clip the
    per-view answer, which is why callers wanting true extents pass nothing.
    """
    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()

    min_x = min_y = min_z = None
    max_x = max_y = max_z = None

    for element in collector:
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
