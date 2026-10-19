# -*- coding: utf-8 -*-
"""Rotate an annotation to match the direction of another element.

Pick the tag, then pick the element it should line up with. The tag spins
about its own head point - it does not move - until it is parallel to the
picked element as that element reads in the current view. Repeat as often
as you like; Escape at either pick ends the tool.

Worked example (what this was built from): tag 4485630 "MULTIBRACE" sitting
horizontal on view S1.31 - 1 - ROOF, aligned to generic model 4484002, a
line-based MULTIBRACE running at 49.69 deg in plan.

Direction is read from the picked element in this order:

  LocationCurve            walls, beams, detail lines, line-based families
  Grid / CurveElement      .Curve / .GeometryCurve
  LocationPoint family     the family's own local X axis (hand direction)

The angle is measured in VIEW space - the element direction projected onto
the view's right and up vectors - so this works in plans, sections and
drafting views alike, not just plan.

Tags are forced to AnyModelDirection orientation first; a Horizontal or
Vertical tag physically cannot hold an arbitrary angle.

Runs silently. Only a skip or a failure raises a dialog.

Target: Revit 2026, IronPython 2.
"""

import math

from Autodesk.Revit.DB import (CategoryType, CurveElement,
                               ElementTransformUtils, FamilyInstance,
                               IndependentTag, Line, LocationCurve,
                               LocationPoint, TagOrientation, TextNote,
                               Transaction)
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from pyrevit import revit, forms

doc = revit.doc
uidoc = revit.uidoc

# --------------------------------------------------------------- settings
KEEP_UPRIGHT = True     # fold the angle into (-90, 90] so text never reads
                        # upside down. False = match the element exactly,
                        # including its direction of travel.
ANGLE_OFFSET_DEG = 0.0  # added after alignment. 90.0 = run the tag square
                        # to the element instead of along it.

TOL = 1e-9


# ---------------------------------------------------------------- helpers
def safe(v):
    """Unicode-safe conversion. Revit strings can carry non-ASCII."""
    try:
        return unicode(v)
    except Exception:
        return str(v)


def eid(elem):
    """Element id as a number, across the 2024 ElementId 64-bit change."""
    try:
        return elem.Id.Value
    except Exception:
        return elem.Id.IntegerValue


def describe(elem):
    """Short 'Category name (id)' label for warning dialogs."""
    cat = elem.Category.Name if elem.Category else elem.GetType().Name
    return u"{} {} ({})".format(safe(cat), safe(elem.Name), eid(elem))


def anchor_of(elem):
    """The point an annotation should spin about, or None if not rotatable.

    Doubles as the test for 'is this thing a tag we can drive' - the pick
    filter calls it, so anything that returns a point here is selectable.
    """
    if isinstance(elem, IndependentTag):
        return elem.TagHeadPosition
    if isinstance(elem, TextNote):
        return elem.Coord

    loc = elem.Location
    if isinstance(loc, LocationPoint):
        # generic annotations, detail items, room/space tags - but NOT model
        # elements. Plenty of columns and framing members have a location
        # point too, and picking one as "the tag" would silently rotate a
        # piece of the building instead of a label.
        #
        # ViewSpecific carries this test, not CategoryType: Detail Items
        # report CategoryType.Model even though they are 2D view-specific
        # linework, so a category-only test throws them out wrongly.
        cat = elem.Category
        if elem.ViewSpecific or (cat is not None
                                 and cat.CategoryType == CategoryType.Annotation):
            return loc.Point
    return None


def direction_of(elem):
    """World-space direction vector of the element to align to, or None."""
    loc = elem.Location
    if isinstance(loc, LocationCurve):
        crv = loc.Curve
        return crv.GetEndPoint(1) - crv.GetEndPoint(0)

    if isinstance(elem, CurveElement):
        crv = elem.GeometryCurve
        return crv.GetEndPoint(1) - crv.GetEndPoint(0)

    # grids and levels carry a curve but no LocationCurve. Some elements
    # raise on these properties rather than simply not having them, so the
    # access is guarded rather than a bare getattr.
    for prop in ("Curve", "GeometryCurve"):
        try:
            crv = getattr(elem, prop, None)
        except Exception:
            crv = None
        if crv is not None:
            return crv.GetEndPoint(1) - crv.GetEndPoint(0)

    if isinstance(elem, FamilyInstance) and isinstance(loc, LocationPoint):
        return elem.GetTransform().BasisX

    return None


def view_angle(vec, view):
    """Angle of a world vector as it reads on screen, in radians.

    Projected onto the view's own right/up axes, so the result is what the
    user sees whether the view is a plan, a section or a drafting view.
    """
    x = vec.DotProduct(view.RightDirection)
    y = vec.DotProduct(view.UpDirection)
    if abs(x) < TOL and abs(y) < TOL:
        return None  # element runs straight into the screen - no readable angle
    return math.atan2(y, x)


def current_angle(elem, view):
    """The annotation's present on-screen angle, in radians.

    Every branch must report a real angle. Falling back to a blanket 0.0
    for an annotation that is ALREADY rotated makes the delta wrong by
    exactly its current rotation, which lands it at the wrong angle
    without any error - so room/space tags read LocationPoint.Rotation
    rather than being lumped in with the unknowns.
    """
    if isinstance(elem, IndependentTag):
        return elem.RotationAngle  # already measured from view right
    if isinstance(elem, TextNote):
        return view_angle(elem.BaseDirection, view) or 0.0
    if isinstance(elem, FamilyInstance):
        return view_angle(elem.GetTransform().BasisX, view) or 0.0

    loc = elem.Location
    if isinstance(loc, LocationPoint):
        try:
            return loc.Rotation
        except Exception:
            return 0.0
    return 0.0


def fold(angle):
    """Bring an angle into (-90, 90] so the text stays the right way up."""
    half = math.pi / 2.0
    while angle > half + TOL:
        angle -= math.pi
    while angle <= -half - TOL:
        angle += math.pi
    return angle


class TagFilter(ISelectionFilter):
    """Only let the first pick land on something we can actually rotate."""

    def AllowElement(self, elem):
        try:
            return anchor_of(elem) is not None
        except Exception:
            return False

    def AllowReference(self, ref, point):
        return False


# ------------------------------------------------------------------- work
def rotate_one(tag, target, view):
    """Spin tag to match target.

    Returns None when it worked, or a message saying why it did not. Success
    is deliberately silent - the tool prints no run report.
    """
    anchor = anchor_of(tag)
    vec = direction_of(target)
    if vec is None:
        return u"No direction could be read from {}.\n\nPick a wall, beam, " \
               u"grid, detail line, or any line-based family.".format(
                   describe(target))

    target_ang = view_angle(vec, view)
    if target_ang is None:
        return u"{} runs straight into the screen in this view, so it has " \
               u"no angle to match.".format(describe(target))

    target_ang += math.radians(ANGLE_OFFSET_DEG)

    was_pinned = tag.Pinned
    before = current_angle(tag, view)

    t = Transaction(doc, "Rotate tag to element")
    t.Start()
    try:
        if was_pinned:
            tag.Pinned = False

        # a Horizontal/Vertical tag cannot hold a free angle - free it first,
        # and re-read the angle afterwards because the switch can move it
        if isinstance(tag, IndependentTag):
            if tag.TagOrientation != TagOrientation.AnyModelDirection:
                tag.TagOrientation = TagOrientation.AnyModelDirection
                doc.Regenerate()
                before = current_angle(tag, view)

        if KEEP_UPRIGHT:
            target_ang = fold(target_ang)

        delta = target_ang - before
        if KEEP_UPRIGHT:
            delta = fold(delta)

        if abs(delta) > 1e-6:
            axis = Line.CreateBound(anchor, anchor + view.ViewDirection)
            ElementTransformUtils.RotateElement(doc, tag.Id, axis, delta)
            doc.Regenerate()

        if was_pinned:
            tag.Pinned = True
        t.Commit()
    except Exception as ex:
        t.RollBack()
        return u"Could not rotate {}.\n\n{}".format(describe(tag), safe(ex))

    return None


def main():
    view = doc.ActiveView
    try:
        view.ViewDirection  # sheets and schedules throw here
    except Exception:
        forms.alert("Run this from a plan, section, elevation or drafting "
                    "view - not a sheet or schedule.",
                    title="Rotate To Element", exitscript=True)

    sel = uidoc.Selection
    tag_filter = TagFilter()

    while True:
        try:
            with forms.WarningBar(title="Rotate To Element: Pick the TAG or "
                                        "TEXT to rotate. ESC to finish."):
                ref = sel.PickObject(ObjectType.Element, tag_filter,
                                     "Pick the annotation to rotate")
            tag = doc.GetElement(ref.ElementId)

            with forms.WarningBar(title="Rotate To Element: Now pick the "
                                        "ELEMENT to line it up with. "
                                        "ESC to finish."):
                ref2 = sel.PickObject(ObjectType.Element,
                                      "Pick the element to align to")
            target = doc.GetElement(ref2.ElementId)
        except OperationCanceledException:
            break

        # a tag is view-specific; rotate it in the view that owns it
        owner = doc.GetElement(tag.OwnerViewId) if tag.OwnerViewId else None
        problem = rotate_one(tag, target, owner or view)
        if problem:
            forms.alert(problem, title="Rotate To Element")


main()
