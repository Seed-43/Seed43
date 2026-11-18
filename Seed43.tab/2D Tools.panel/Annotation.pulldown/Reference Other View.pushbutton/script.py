# -*- coding: utf-8 -*-
"""Reference Other View.

Replace a section, detail callout or elevation with a "Reference Other View"
marker that points at a view which already exists.

Revit cannot convert a marker in place, so the swap is: create the reference
marker where the old one sits, then delete the old view (which takes its
marker with it). Both happen in ONE transaction - if either step fails,
nothing changes.

Target: Revit 2022-2026, IronPython 2.
"""

# ── IMPORTS ─────────────────────────────────────────────────────────────────

from Autodesk.Revit.DB import (BuiltInCategory, BuiltInParameter,
                               ElementCategoryFilter, ElementTransformUtils,
                               FilteredElementCollector, Plane, Transaction, View,
                               ViewSection, ViewType, Viewport, XYZ)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import forms, revit

doc = revit.doc
uidoc = revit.uidoc

# ── CONSTANTS ───────────────────────────────────────────────────────────────

TOOL_NAME = "Reference Other View"

# What a reference SECTION may point at, and what a reference CALLOUT may.
SECTION_TARGETS = ("Section", "Detail", "DraftingView")
CALLOUT_TARGETS = ("Section", "Detail", "DraftingView", "Elevation",
                   "FloorPlan", "CeilingPlan", "EngineeringPlan")

# Categories a picked marker can belong to.
MARKER_CATEGORIES = ("OST_Viewers", "OST_Callouts", "OST_Elev",
                     "OST_Sections")


# ── HELPERS ─────────────────────────────────────────────────────────────────

def eid(element_id):
    """ElementId as a plain number, across the 2024 64-bit change."""
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def _bic_ids(names):
    """Category ids (as numbers) for whichever of these this build has."""
    found = set()
    for name in names:
        member = getattr(BuiltInCategory, name, None)
        if member is not None:
            found.add(int(member))
    return found


def _view_types(names):
    """The named ViewType members that this Revit build actually has."""
    found = []
    for name in names:
        member = getattr(ViewType, name, None)
        if member is not None:
            found.append(member)
    return found


MARKER_CAT_IDS = _bic_ids(MARKER_CATEGORIES)


def is_callout(view):
    """True when the view is a callout cut from a parent view."""
    try:
        return eid(view.GetCalloutParentId()) > 0
    except Exception:
        return False


def sheet_lookup():
    """{view id number: sheet number} for every view placed on a sheet."""
    lookup = {}
    for vp in FilteredElementCollector(doc).OfClass(Viewport):
        sheet = doc.GetElement(vp.SheetId)
        if sheet is not None:
            lookup[eid(vp.ViewId)] = sheet.SheetNumber
    return lookup


# ── PICKING THE VIEW TO REPLACE ─────────────────────────────────────────────

class MarkerFilter(ISelectionFilter):
    """Lets the user pick only section, callout and elevation markers."""

    def AllowElement(self, element):
        if element is None or element.Category is None:
            return False
        return eid(element.Category.Id) in MARKER_CAT_IDS

    def AllowReference(self, reference, position):
        return False


def views_for_marker(marker):
    """The section / detail / elevation views a marker element belongs to."""
    found = []
    cat_filter = ElementCategoryFilter(marker.Category.Id)
    marker_id = eid(marker.Id)
    for view in FilteredElementCollector(doc).OfClass(ViewSection):
        if view.IsTemplate:
            continue
        try:
            deps = view.GetDependentElements(cat_filter)
        except Exception:
            continue
        for dep in deps:
            if eid(dep) == marker_id:
                found.append(view)
                break
    return found


class Choice(object):
    """One row in a pick list."""

    def __init__(self, view, sheets):
        self.view = view
        sheet = sheets.get(eid(view.Id)) or u"no sheet"
        number = u"-"
        try:
            param = view.get_Parameter(BuiltInParameter.VIEWER_DETAIL_NUMBER)
            if param is not None and param.AsString():
                number = param.AsString()
        except Exception:
            pass
        # Detail number and sheet lead, so "X4" or "S2.02" typed in the
        # filter finds the view the way the marker on the drawing reads.
        self.name = u"{} / {}   {}   [{}]".format(
            number, sheet, view.Name, view.ViewType)


def pick_view_to_replace():
    """Return (old view, parent view, marker) - Nones if cancelled."""
    sheets = sheet_lookup()
    marker = None
    old_view = None

    # Pre-selection first: a marker or a view picked before the click.
    for element_id in uidoc.Selection.GetElementIds():
        element = doc.GetElement(element_id)
        if isinstance(element, ViewSection):
            old_view = element
            break
        if (element is not None and element.Category is not None
                and eid(element.Category.Id) in MARKER_CAT_IDS):
            marker = element
            break

    if old_view is None and marker is None:
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element, MarkerFilter(),
                "Pick the section, detail or elevation marker to replace")
        except OperationCanceledException:
            return None, None, None
        marker = doc.GetElement(ref.ElementId)

    parent = doc.ActiveView
    if marker is not None:
        views = views_for_marker(marker)
        if not views:
            forms.alert("Could not find the view behind that marker.\n\n"
                        "Select the view in the Project Browser instead, "
                        "then run the tool again.", title=TOOL_NAME)
            return None, None, None
        if len(views) == 1:
            old_view = views[0]
        else:
            rows = [Choice(v, sheets) for v in views]
            picked = forms.SelectFromList.show(
                rows, name_attr="name", title="Which view on this marker?",
                button_name="Replace this one", multiselect=False)
            if not picked:
                return None, None, None
            old_view = picked.view
        owner = doc.GetElement(marker.OwnerViewId)
        if isinstance(owner, View):
            parent = owner

    if is_callout(old_view):
        parent = doc.GetElement(old_view.GetCalloutParentId())
    return old_view, parent, marker


# ── PICKING THE VIEW TO REFERENCE ───────────────────────────────────────────

def pick_target(old_view, parent, callout):
    """List the views that can be referenced; typing filters the list."""
    allowed = _view_types(CALLOUT_TARGETS if callout else SECTION_TARGETS)
    sheets = sheet_lookup()

    mode = forms.CommandSwitchWindow.show(
        ["Pick marker in view", "Choose from list"],
        message="Reference which view?")
    if not mode:
        return None
    if mode == "Pick marker in view":
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element, MarkerFilter(),
                "Pick the marker of the view to reference")
        except OperationCanceledException:
            return None
        views = [v for v in views_for_marker(doc.GetElement(ref.ElementId))
                 if v.ViewType in allowed
                 and eid(v.Id) not in (eid(old_view.Id), eid(parent.Id))]
        if not views:
            forms.alert("That marker's view cannot be referenced from here.",
                        title=TOOL_NAME)
            return None
        if len(views) == 1:
            return views[0]
        picked = forms.SelectFromList.show(
            [Choice(v, sheets) for v in views], name_attr="name",
            title="Which view on this marker?", button_name="Reference this",
            multiselect=False)
        return picked.view if picked else None

    rows = []
    for view in FilteredElementCollector(doc).OfClass(View):
        if view.IsTemplate or view.ViewType not in allowed:
            continue
        if eid(view.Id) in (eid(old_view.Id), eid(parent.Id)):
            continue
        rows.append(Choice(view, sheets))
    if not rows:
        forms.alert("No views can be referenced from here.", title=TOOL_NAME)
        return None
    rows.sort(key=lambda r: r.name.lower())
    picked = forms.SelectFromList.show(
        rows, name_attr="name", multiselect=False,
        title="Reference which view? (type to filter)",
        button_name="Reference this view")
    return picked.view if picked else None


# ── GEOMETRY ────────────────────────────────────────────────────────────────

def span_along(element, view, axis):
    """(low, high) of an element's bounding box along a world X or Y axis.

    Used to compare how far two markers reach, whatever their bubbles and
    stubs add beyond the line itself. At any other angle an axis-aligned box
    would over-read, so None is returned and the caller skips the check.
    """
    if element is None:
        return None
    box = element.get_BoundingBox(view)
    if box is None:
        return None
    if abs(axis.Y) > 0.9999:
        sign, lo, hi = axis.Y, box.Min.Y, box.Max.Y
    elif abs(axis.X) > 0.9999:
        sign, lo, hi = axis.X, box.Min.X, box.Max.X
    else:
        return None
    return min(sign * lo, sign * hi), max(sign * lo, sign * hi)


def find_marker(old_view, parent):
    """The marker element of old_view that is drawn in the parent view."""
    for name in MARKER_CATEGORIES:
        cat = getattr(BuiltInCategory, name, None)
        if cat is None:
            continue
        try:
            ids = old_view.GetDependentElements(ElementCategoryFilter(cat))
        except Exception:
            continue
        for dep in ids:
            el = doc.GetElement(dep)
            if el is not None and eid(el.OwnerViewId) == eid(parent.Id):
                return el
    return None


def marker_points(old_view, callout):
    """Model-space points that reproduce the old marker's position.

    A section / elevation marker is a line, so head and tail are its two
    ends, taken from the view's crop width. A callout is the rectangle of its crop
    region, so the points are two opposite corners.
    """
    box = old_view.CropBox
    t = box.Transform
    if callout:
        return (t.OfPoint(XYZ(box.Min.X, box.Min.Y, 0.0)),
                t.OfPoint(XYZ(box.Max.X, box.Max.Y, 0.0)))
    # The bubble (head) sits at the high end of the old marker.
    return (t.Origin + t.BasisX.Multiply(box.Max.X),
            t.Origin + t.BasisX.Multiply(box.Min.X))


def needs_flip(old_view, parent, head, tail):
    """True when head -> tail would make the arrow face the wrong way.

    A reference marker's arrow follows its head -> tail direction, so the
    bubble end and the arrow side can only both match the old marker by
    flipping the new one afterwards.
    """
    direction = (tail - head).Normalize()
    looks = parent.ViewDirection.CrossProduct(direction)
    wanted = old_view.ViewDirection.Negate()
    return looks.DotProduct(wanted) < 0


def flip_marker(parent, element_id, head, tail):
    """Mirror a new reference marker about its own line. True on success."""
    try:
        if not ElementTransformUtils.CanMirrorElement(doc, element_id):
            return False
        direction = (tail - head).Normalize()
        normal = parent.ViewDirection.CrossProduct(direction).Normalize()
        before = set(eid(i) for i in parent.GetReferenceSections())
        ElementTransformUtils.MirrorElement(
            doc, element_id, Plane.CreateByNormalAndOrigin(normal, head))
        doc.Regenerate()
        # Mirroring leaves a second marker behind the flipped one. Keep the
        # one we mirrored and delete any extra, so only one marker remains.
        for extra in list(parent.GetReferenceSections()):
            if eid(extra) not in before and eid(extra) != eid(element_id):
                doc.Delete(extra)
        return True
    except Exception:
        return False


def create_reference_section(parent, target, head, tail):
    """Create a reference section and return its new element id."""
    before = set(eid(i) for i in parent.GetReferenceSections())
    ViewSection.CreateReferenceSection(doc, parent.Id, target.Id, head, tail)
    doc.Regenerate()
    new = [i for i in parent.GetReferenceSections() if eid(i) not in before]
    return new[0] if new else None


def calibrated_points(parent, new_id, old_span, axis, head, tail):
    """Head and tail moved so the new marker reaches as far as the old one.

    A reference marker draws its own bubble and stub beyond head and tail,
    and those are not the same length as the old marker's, so the ends are
    measured after the first attempt and moved by the difference.
    """
    new_span = span_along(doc.GetElement(new_id), parent, axis)
    if not old_span or not new_span:
        return None
    d_hi = old_span[1] - new_span[1]
    d_lo = old_span[0] - new_span[0]
    if abs(d_hi) < 1e-3 and abs(d_lo) < 1e-3:
        return None
    new_head = head + axis.Multiply(d_hi)
    new_tail = tail + axis.Multiply(d_lo)
    if (new_head - new_tail).DotProduct(axis) < 0.1:
        return None
    return new_head, new_tail


# ── MAIN ────────────────────────────────────────────────────────────────────

def main():
    old_view, parent, marker = pick_view_to_replace()
    if old_view is None:
        return
    if parent is None or not isinstance(parent, View):
        forms.alert("Could not work out which view the marker sits in.",
                    title=TOOL_NAME)
        return
    if eid(old_view.Id) == eid(doc.ActiveView.Id):
        forms.alert("The view being replaced is open as the active view. "
                    "Open the view the marker is drawn in and try again.",
                    title=TOOL_NAME)
        return

    callout = is_callout(old_view)
    target = pick_target(old_view, parent, callout)
    if target is None:
        return

    sheets = sheet_lookup()
    on_sheet = sheets.get(eid(old_view.Id))
    try:
        count = (FilteredElementCollector(doc, old_view.Id)
                 .WhereElementIsNotElementType().GetElementCount())
    except Exception:
        count = 0
    lost = u"{} view-specific elements".format(count) if count else \
        u"everything drawn in it"
    note = u"\nIt is on sheet {} and will come off it.".format(on_sheet) \
        if on_sheet else u""
    msg = (u"'{old}' will be DELETED together with {lost}.{note}\n\n"
           u"A reference marker to '{new}' replaces it in '{par}'.\n\n"
           u"Continue?").format(old=old_view.Name, lost=lost, note=note,
                                new=target.Name, par=parent.Name)
    if not forms.alert(msg, title=TOOL_NAME, yes=True, no=True):
        return

    try:
        head, tail = marker_points(old_view, callout)
    except Exception as exc:
        forms.alert(u"Could not read the old marker's position:\n{}"
                    .format(exc), title=TOOL_NAME)
        return

    old_name = old_view.Name
    axis = old_view.CropBox.Transform.BasisX
    old_span = None
    if not callout:
        old_span = span_along(marker or find_marker(old_view, parent),
                              parent, axis)
    flip = (not callout) and needs_flip(old_view, parent, head, tail)
    flipped = True
    with Transaction(doc, "Reference Other View") as tx:
        tx.Start()
        try:
            # Create first: if this fails nothing has been deleted yet.
            if callout:
                ViewSection.CreateReferenceCallout(
                    doc, parent.Id, target.Id, head, tail)
            else:
                new_id = create_reference_section(parent, target, head, tail)
                fixed = calibrated_points(parent, new_id, old_span, axis,
                                          head, tail) if new_id else None
                if fixed:
                    try:
                        second = create_reference_section(
                            parent, target, fixed[0], fixed[1])
                    except Exception:
                        second = None
                    if second:
                        doc.Delete(new_id)
                        new_id = second
                        head, tail = fixed
                if flip and new_id:
                    flipped = flip_marker(parent, new_id, head, tail)
                elif flip:
                    flipped = False
            doc.Delete(old_view.Id)
        except Exception as exc:
            tx.RollBack()
            forms.alert(u"Nothing was changed.\n\n{}".format(exc),
                        title=TOOL_NAME)
            return
        tx.Commit()

    forms.toast(u"'{}' replaced with a reference to '{}'"
                .format(old_name, target.Name), title=TOOL_NAME)
    if flip and not flipped:
        forms.alert(u"The new marker's arrow faces the other way. Select it "
                    u"and click its flip arrow to turn it round.",
                    title=TOOL_NAME)


main()
