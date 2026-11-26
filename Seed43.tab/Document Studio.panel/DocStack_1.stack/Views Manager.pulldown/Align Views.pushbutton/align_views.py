# -*- coding: utf-8 -*-
"""
Align Views - line up drafting views, legends and schedules on sheets, the
way model views line up by their coordinates.

Two modes:
  View origin   Every chosen viewport is moved so its view's 0,0 sits on
                its sheet exactly where the reference view's 0,0 sits on
                the reference sheet. Works across sheets, no picking.
  Pick points   On the open sheet: pick the reference viewport; the tool
                goes into that view to pick the point (snapping to its
                lines), then back out to the sheet. Then, for each view to
                move: pick the viewport, pick its matching point inside it,
                and it lines up. Esc to finish.

  Schedules      Schedules have no Saved Position, so they line up by their
                top-left corner (a schedule's insert point). Reference: a
                drafting view or legend (its 0,0, which is where a table
                built by pyLink starts) or another schedule. Every chosen
                schedule moves so its top-left sits there.

Both can line up horizontally, vertically or both. Uses
Viewport.GetPositionAtViewAnchor (Revit 2026+) for the origin mode; picking works in any version.
"""
from System.Collections.Generic import List
from pyrevit import revit, DB, forms, script
from pyrevit.forms import WarningBar
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType, ObjectSnapTypes

doc = revit.doc
uidoc = revit.uidoc

ORIGIN = 'Line up by view origin'
PICK = 'Line up by picked points'
SCHEDULES = 'Line up schedules (top-left)'

BOTH = 'Both'
HORIZONTAL = 'Horizontal only (same height)'
VERTICAL = 'Vertical only (same left/right)'

# Views whose own 0,0 means something to line up on: they have no model
# coordinates, so their origin is the only shared reference.
ORIGIN_VIEW_TYPES = (DB.ViewType.DraftingView, DB.ViewType.Legend)

SNAPS = (ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections
         | ObjectSnapTypes.Midpoints | ObjectSnapTypes.Nearest
         | ObjectSnapTypes.Perpendicular | ObjectSnapTypes.Centers)


# ── HELPERS ──────────────────────────────────────────────────────────────────

class _ViewportFilter(ISelectionFilter):
    def AllowElement(self, el):
        return isinstance(el, DB.Viewport)

    def AllowReference(self, ref, pt):
        return False


def _origin_on_sheet(vp):
    """Where the view's 0,0 lands on its sheet, or None if Revit can't say.

    GetPositionAtViewAnchor (2026+) works for every view. The projection
    transform only exists for model views: drafting views and legends
    report "no transforms", so it is just the fallback."""
    try:
        return vp.GetPositionAtViewAnchor(DB.ViewAnchor.ViewOrigin)
    except Exception:
        pass
    try:
        if vp.HasViewportTransforms():
            return vp.GetProjectionToSheetTransform().OfPoint(DB.XYZ.Zero)
    except Exception:
        pass
    return None


def _move(vp, delta, axis):
    """Shift a viewport on its sheet by delta, limited to the axis."""
    dx = 0.0 if axis == HORIZONTAL else delta.X
    dy = 0.0 if axis == VERTICAL else delta.Y
    if abs(dx) < 1e-9 and abs(dy) < 1e-9:
        return False
    vp.SetBoxCenter(vp.GetBoxCenter() + DB.XYZ(dx, dy, 0))
    return True


def _label(vp):
    sheet = doc.GetElement(vp.SheetId)
    view = doc.GetElement(vp.ViewId)
    return u'{}  {}   [{}]'.format(sheet.SheetNumber, sheet.Name, view.Name)


def _pick_viewport(prompt):
    with WarningBar(title=prompt + ' Esc to finish.'):
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element,
                                             _ViewportFilter(), prompt)
        except Exception:
            return None
    return doc.GetElement(ref.ElementId)


def _pick_point(prompt):
    with WarningBar(title=prompt + ' Esc to finish.'):
        try:
            return uidoc.Selection.PickPoint(SNAPS, prompt)
        except Exception:
            return None


def _ask_axis():
    return forms.CommandSwitchWindow.show(
        [BOTH, HORIZONTAL, VERTICAL], message='Line up:')


# ── MODE: VIEW ORIGIN ────────────────────────────────────────────────────────

class VpItem(forms.TemplateListItem):
    @property
    def name(self):
        return _label(self.item)


def _origin_viewports():
    out = []
    for vp in DB.FilteredElementCollector(doc).OfClass(DB.Viewport):
        view = doc.GetElement(vp.ViewId)
        if view is not None and view.ViewType in ORIGIN_VIEW_TYPES:
            out.append(vp)
    return out


def run_origin():
    candidates = _origin_viewports()
    if len(candidates) < 2:
        forms.alert('Needs at least two drafting views or legends on sheets.',
                    exitscript=True)

    # Reference: the selected viewport, or one picked from the list.
    candidate_ids = set(v.Id for v in candidates)
    sel = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    sel = [e for e in sel if isinstance(e, DB.Viewport) and e.Id in candidate_ids]
    reference = sel[0] if len(sel) == 1 else None
    items = sorted((VpItem(v) for v in candidates), key=lambda i: i.name)
    if reference is None:
        picked = forms.SelectFromList.show(
            items, title='Reference view (stays where it is)',
            button_name='Use as reference', multiselect=False)
        if not picked:
            script.exit()
        reference = picked

    targets = forms.SelectFromList.show(
        [i for i in items if i.item.Id != reference.Id],
        title=u'Line these up with {}'.format(
            doc.GetElement(reference.ViewId).Name),
        button_name='Line up', multiselect=True)
    if not targets:
        script.exit()
    axis = _ask_axis()
    if not axis:
        script.exit()

    goal = _origin_on_sheet(reference)
    if goal is None:
        forms.alert('This Revit version cannot report where a drafting '
                    "view's origin sits on the sheet (needs Revit 2026). "
                    'Use Line up by picked points instead.', exitscript=True)
    moved = 0
    skipped = 0
    with revit.Transaction('Align Views: by view origin'):
        for vp in targets:
            here = _origin_on_sheet(vp)
            if here is None:
                skipped += 1
            elif _move(vp, goal - here, axis):
                moved += 1
    msg = '{} of {} views moved.'.format(moved, len(targets))
    if skipped:
        msg += ' {} could not be read and were left.'.format(skipped)
    forms.alert(msg, title='Align Views', warn_icon=False)


# ── MODE: SCHEDULES ──────────────────────────────────────────────────────────

class SchedItem(forms.TemplateListItem):
    """A drafting view / legend viewport or a schedule on a sheet."""
    @property
    def name(self):
        it = self.item
        if isinstance(it, DB.Viewport):
            return u'View:  ' + _label(it)
        sheet = doc.GetElement(it.OwnerViewId)
        sched = doc.GetElement(it.ScheduleId)
        return u'Schedule:  {}  {}   [{}]'.format(
            sheet.SheetNumber, sheet.Name, sched.Name)


def _schedule_instances():
    out = []
    for inst in DB.FilteredElementCollector(doc).OfClass(
            DB.ScheduleSheetInstance):
        try:
            if inst.IsTitleblockRevisionSchedule:
                continue
        except Exception:
            pass
        out.append(inst)
    return out


def _schedule_goal(ref):
    """Sheet point the schedules line up to: a view's origin, or a
    schedule's own top-left."""
    if isinstance(ref, DB.Viewport):
        return _origin_on_sheet(ref)
    return ref.Point


def run_schedules():
    scheds = _schedule_instances()
    if not scheds:
        forms.alert('There are no schedules on sheets.', exitscript=True)
    refs = list(_origin_viewports()) + scheds
    items = sorted((SchedItem(v) for v in refs), key=lambda i: i.name)

    reference = None
    selected = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    ref_ids = set(v.Id for v in refs)
    selected = [e for e in selected if e is not None and e.Id in ref_ids]
    if len(selected) == 1:
        reference = selected[0]
    if reference is None:
        reference = forms.SelectFromList.show(
            items, title='Reference (a view origin or a schedule)',
            button_name='Use as reference', multiselect=False)
        if not reference:
            script.exit()
    goal = _schedule_goal(reference)
    if goal is None:
        forms.alert("Can't read where that view's origin sits on its sheet "
                    '(needs Revit 2026). Pick a schedule as the reference.',
                    exitscript=True)

    targets = forms.SelectFromList.show(
        [i for i in items if isinstance(i.item, DB.ScheduleSheetInstance)
         and i.item.Id != reference.Id],
        title='Line these schedules up', button_name='Line up',
        multiselect=True)
    if not targets:
        script.exit()
    axis = _ask_axis()
    if not axis:
        script.exit()

    moved = 0
    with revit.Transaction('Align Views: schedules'):
        for inst in targets:
            delta = goal - inst.Point
            dx = 0.0 if axis == HORIZONTAL else delta.X
            dy = 0.0 if axis == VERTICAL else delta.Y
            if abs(dx) < 1e-9 and abs(dy) < 1e-9:
                continue
            try:
                DB.ElementTransformUtils.MoveElement(
                    doc, inst.Id, DB.XYZ(dx, dy, 0))
                moved += 1
            except Exception:
                pass
    forms.alert('{} of {} schedules moved.'.format(moved, len(targets)),
                title='Align Views', warn_icon=False)


# ── MODE: PICK POINTS ────────────────────────────────────────────────────────

def _view_point_to_sheet(vp, point):
    """A point picked inside a drafting view or legend, as a sheet point.

    Those views have no projection transform, but Revit 2026 reports where
    the view's 0,0 sits on the sheet; from there it is the view scale and
    the viewport's quarter-turn rotation."""
    origin = _origin_on_sheet(vp)
    if origin is None:
        return None
    scale = float(doc.GetElement(vp.ViewId).Scale or 1)
    x, y = point.X / scale, point.Y / scale
    if vp.Rotation == DB.ViewportRotation.Clockwise:
        x, y = y, -x
    elif vp.Rotation == DB.ViewportRotation.Counterclockwise:
        x, y = -y, x
    return origin + DB.XYZ(x, y, 0)


def _pick_point_in_view(vp, sheet, prompt):
    """Go into the viewport's view, pick a point there (where the snaps to
    its lines work), come back out to the sheet, and return the point in
    sheet coordinates. A view tab that wasn't open before is closed again.
    None on Esc."""
    view = doc.GetElement(vp.ViewId)
    if view.ViewType not in ORIGIN_VIEW_TYPES:
        # A model view: snaps to its geometry work from the sheet already.
        return _pick_point(prompt)
    was_open = any(uv.ViewId == view.Id for uv in uidoc.GetOpenUIViews())
    uidoc.ActiveView = view
    try:
        point = _pick_point(prompt)
    finally:
        uidoc.ActiveView = sheet
        if not was_open:
            for uv in uidoc.GetOpenUIViews():
                if uv.ViewId == view.Id:
                    try:
                        uv.Close()
                    except Exception:
                        pass
    if point is None:
        return None
    return _view_point_to_sheet(vp, point)


def run_pick():
    sheet = revit.active_view
    if sheet.ViewType != DB.ViewType.DrawingSheet:
        forms.alert('Open the sheet with the views to line up first.',
                    exitscript=True)
    axis = _ask_axis()
    if not axis:
        script.exit()
    ref = _pick_viewport('Pick the REFERENCE view (stays where it is).')
    if ref is None:
        script.exit()
    goal = _pick_point_in_view(
        ref, sheet, u'In "{}": pick the point to line up TO.'.format(
            doc.GetElement(ref.ViewId).Name))
    if goal is None:
        script.exit()

    moved = 0
    while True:
        vp = _pick_viewport('Pick the next view to move.')
        if vp is None:
            break
        point = _pick_point_in_view(
            vp, sheet, u'In "{}": pick the point that should line up.'.format(
                doc.GetElement(vp.ViewId).Name))
        if point is None:
            break
        with revit.Transaction('Align Views: by picked points'):
            if _move(vp, goal - point, axis):
                moved += 1
    if moved:
        forms.alert('{} view(s) lined up.'.format(moved), title='Align Views',
                    warn_icon=False)


# ── RUN ──────────────────────────────────────────────────────────────────────

mode = forms.CommandSwitchWindow.show([ORIGIN, PICK, SCHEDULES],
                                      message='Align views:')
if mode == ORIGIN:
    run_origin()
elif mode == PICK:
    run_pick()
elif mode == SCHEDULES:
    run_schedules()
