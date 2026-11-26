# -*- coding: utf-8 -*-
# "Center In View"
# "Seed43"
# """
# Move the elements of a drafting view so a point you pick lands on the
# view's centre point (its origin), then reset the view title on any sheet
# it sits on to its default position. Pick the point on the elements, done.
# """
#
# Target: Revit 2022-2026, IronPython 2.

# ── IMPORTS ────────────────────────────────────────────────────────────────
from System.Collections.Generic import List
from Autodesk.Revit.DB import (ElementId, ElementTransformUtils, Transaction,
                               Viewport, ViewDrafting, ViewType, XYZ,
                               FilteredElementCollector)
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ObjectSnapTypes
from pyrevit import revit, forms

doc = revit.doc
uidoc = revit.uidoc

# ── CONSTANTS ──────────────────────────────────────────────────────────────
TITLE = "Center In View"
CENTER = XYZ(0.0, 0.0, 0.0)     # a drafting view's centre point is its origin
SNAPS = (ObjectSnapTypes.Endpoints | ObjectSnapTypes.Midpoints |
         ObjectSnapTypes.Intersections | ObjectSnapTypes.Centers)


# ── HELPERS (private to this file) ─────────────────────────────────────────
def safe(v):
    """Unicode-safe conversion. Revit strings can carry non-ASCII."""
    try:
        return unicode(v)
    except Exception:
        return str(v)


def is_2d_view(view):
    """True for drafting views and legends - the views that have an origin."""
    return isinstance(view, ViewDrafting) or view.ViewType == ViewType.Legend


# ── CORE LOGIC ─────────────────────────────────────────────────────────────
# --- Which elements move ---
def elements_to_move(view):
    """Selected elements if there is a selection, else every element drawn
    in this view. Model-less views only hold view-specific elements, so the
    owner test just drops invisible bookkeeping elements."""
    picked = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    picked = [e for e in picked if e is not None and e.OwnerViewId == view.Id]
    if picked:
        return picked
    found = FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType()
    return [e for e in found if e.OwnerViewId == view.Id]


# --- View title ---
def reset_view_titles(view):
    """Put the title of every viewport showing this view back at its default
    position, as if the viewport had just been placed. A zero LabelOffset is
    that default. Pinned viewports are unpinned for the change and pinned
    again."""
    for vp in FilteredElementCollector(doc).OfClass(Viewport):
        if vp.ViewId != view.Id:
            continue
        was_pinned = vp.Pinned
        if was_pinned:
            vp.Pinned = False
        try:
            vp.LabelOffset = XYZ(0.0, 0.0, 0.0)
        except Exception:
            pass    # NOTE: a title that cannot move is left as it was
        if was_pinned:
            vp.Pinned = True


# --- The move ---
def move_to_center(view, elems, from_point):
    """Move elems so from_point lands on the view centre.

    Pinned elements are unpinned for the move and pinned again afterwards.
    Returns the number of elements that could not be moved.
    """
    vector = CENTER - XYZ(from_point.X, from_point.Y, 0.0)
    t = Transaction(doc, TITLE)
    t.Start()
    repin = []
    try:
        for e in elems:
            if e.Pinned:
                e.Pinned = False
                repin.append(e)

        ids = List[ElementId]([e.Id for e in elems])
        failed = 0
        try:
            ElementTransformUtils.MoveElements(doc, ids, vector)
        except Exception:
            # NOTE: one unmovable element rejects the whole batch, so fall
            # back to one at a time and count the ones that refuse.
            failed = 0
            for e in elems:
                try:
                    ElementTransformUtils.MoveElement(doc, e.Id, vector)
                except Exception:
                    failed += 1

        for e in repin:
            e.Pinned = True

        # the content has moved under the viewport, so put the view title
        # back where Revit would place it on a fresh viewport
        doc.Regenerate()
        reset_view_titles(view)
        t.Commit()
        return failed
    except Exception:
        t.RollBack()
        raise


# ── UI / ENTRY POINT ───────────────────────────────────────────────────────
def main():
    view = doc.ActiveView
    if not is_2d_view(view):
        forms.alert("Run this in a drafting view or legend.",
                    title=TITLE, exitscript=True)

    elems = elements_to_move(view)
    if not elems:
        forms.alert("There is nothing in this view to move.",
                    title=TITLE, exitscript=True)

    try:
        with forms.WarningBar(title="Center In View: Pick the point that "
                                    "should land on the view centre. "
                                    "ESC to cancel."):
            point = uidoc.Selection.PickPoint(SNAPS, "Pick the reference point")
    except OperationCanceledException:
        return

    try:
        failed = move_to_center(view, elems, point)
    except Exception as ex:
        forms.alert(u"Could not move the elements.\n\n{}".format(safe(ex)),
                    title=TITLE)
        return

    if failed:
        forms.alert(u"{} element(s) could not be moved.".format(failed),
                    title=TITLE)


if __name__ == "__main__":
    main()
