# -*- coding: utf-8 -*-
"""
View Order - bring views and schedules on a sheet to the front or send them
to the back.

Revit has no draw-order setting for viewports or schedules: it draws them in
the order they were placed, the newest on top. So an item is brought to the
front by placing it again, and sent to the back by placing again everything
drawn above it (views and schedules alike). Each one goes back exactly where
it was; a viewport keeps its viewport type, detail number, title position,
rotation and instance parameters, a schedule keeps its position, segment and
pin.
"""
from System.Collections.Generic import List
from pyrevit import revit, DB, forms, script
from pyrevit.forms import WarningBar
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

doc = revit.doc
uidoc = revit.uidoc

FRONT = 'Bring to Front'
BACK = 'Send to Back'


# ── PICK ─────────────────────────────────────────────────────────────────────

def _is_sheet_item(el):
    """A viewport or a schedule on a sheet. A schedule that is part of a
    title block (the revision schedule) is not one of the sheet's items."""
    if isinstance(el, DB.Viewport):
        return True
    if isinstance(el, DB.ScheduleSheetInstance):
        try:
            return not el.IsTitleblockRevisionSchedule
        except Exception:
            return True
    return False


class _SheetItemFilter(ISelectionFilter):
    def AllowElement(self, el):
        return _is_sheet_item(el)

    def AllowReference(self, ref, pt):
        return False


def _sheet_id(item):
    """Id of the sheet a viewport or a schedule instance sits on."""
    if isinstance(item, DB.Viewport):
        return item.SheetId
    return item.OwnerViewId


def _picked():
    """Viewports and schedules selected before running, or picked on the
    active sheet."""
    sel = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    sel = [e for e in sel if e is not None and _is_sheet_item(e)]
    if sel:
        return sel
    if revit.active_view.ViewType != DB.ViewType.DrawingSheet:
        forms.alert('Open a sheet, or select views or schedules on one, '
                    'first.', exitscript=True)
    with WarningBar(title='Pick the views and schedules to move. '
                          'Click Finish when done.'):
        try:
            refs = uidoc.Selection.PickObjects(ObjectType.Element,
                                               _SheetItemFilter(),
                                               'Pick the views to move')
        except Exception:
            script.exit()
    return [doc.GetElement(r.ElementId) for r in refs]


def _id_int(eid):
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue


# ── RE-PLACE ─────────────────────────────────────────────────────────────────

_SKIP_PARAMS = (DB.BuiltInParameter.VIEWPORT_DETAIL_NUMBER,
                DB.BuiltInParameter.VIEWPORT_SHEET_NUMBER,
                DB.BuiltInParameter.VIEWPORT_SHEET_NAME)


def _instance_values(vp):
    """Writable instance parameters other than the ones set explicitly."""
    skip = set(int(b) for b in _SKIP_PARAMS)
    out = []
    for p in vp.Parameters:
        if p.IsReadOnly:
            continue
        try:
            if p.Definition.BuiltInParameter != DB.BuiltInParameter.INVALID \
                    and int(p.Definition.BuiltInParameter) in skip:
                continue
        except Exception:
            pass
        st = p.StorageType
        if st == DB.StorageType.String:
            out.append((p.Definition.Name, st, p.AsString()))
        elif st == DB.StorageType.Integer:
            out.append((p.Definition.Name, st, p.AsInteger()))
        elif st == DB.StorageType.Double:
            out.append((p.Definition.Name, st, p.AsDouble()))
        elif st == DB.StorageType.ElementId:
            out.append((p.Definition.Name, st, p.AsElementId()))
    return out


def _replace(vp):
    """Place vp's view again where it is, as the newest viewport on the
    sheet, and return the new viewport."""
    sheet = doc.GetElement(vp.SheetId)
    view_id = vp.ViewId
    center = vp.GetBoxCenter()
    type_id = vp.GetTypeId()
    rotation = vp.Rotation
    detail = vp.get_Parameter(
        DB.BuiltInParameter.VIEWPORT_DETAIL_NUMBER).AsString()
    try:
        label_offset, label_len = vp.LabelOffset, vp.LabelLineLength
    except Exception:
        label_offset = label_len = None
    values = _instance_values(vp)

    sheet.DeleteViewport(vp)
    new = DB.Viewport.Create(doc, sheet.Id, view_id, center)
    new.ChangeTypeId(type_id)
    if rotation != getattr(DB.ViewportRotation, 'None'):
        new.Rotation = rotation
    new.SetBoxCenter(center)
    if detail:
        try:
            new.get_Parameter(
                DB.BuiltInParameter.VIEWPORT_DETAIL_NUMBER).Set(detail)
        except Exception:
            pass
    try:
        if label_offset is not None:
            new.LabelOffset = label_offset
        if label_len:
            new.LabelLineLength = label_len
    except Exception:
        pass
    for name, st, val in values:
        p = new.LookupParameter(name)
        if p is None or p.IsReadOnly or val is None:
            continue
        try:
            p.Set(val)
        except Exception:
            pass
    return new


def _replace_schedule(inst):
    """Place a schedule (or one segment of a split schedule) on its sheet
    again, where it is, as the newest item on the sheet."""
    sheet_id = inst.OwnerViewId
    sched_id = inst.ScheduleId
    point = inst.Point
    pinned = inst.Pinned
    try:
        segment = inst.SegmentIndex
    except Exception:
        segment = -1
    sched = doc.GetElement(sched_id)
    split = False
    try:
        split = sched.IsSplit()
    except Exception:
        pass

    if pinned:
        inst.Pinned = False
    doc.Delete(inst.Id)
    if split and segment is not None and segment >= 0:
        new = DB.ScheduleSheetInstance.Create(doc, sheet_id, sched_id, point,
                                              segment)
    else:
        new = DB.ScheduleSheetInstance.Create(doc, sheet_id, sched_id, point)
    if pinned:
        new.Pinned = True
    return new


def _replace_item(item):
    if isinstance(item, DB.Viewport):
        return _replace(item)
    return _replace_schedule(item)


def _sheet_items(sheet):
    """Everything on a sheet that has a draw order here: its viewports and
    its schedules."""
    items = [doc.GetElement(i) for i in sheet.GetAllViewports()]
    for inst in list(DB.FilteredElementCollector(doc, sheet.Id)
                     .OfClass(DB.ScheduleSheetInstance).ToElements()):
        if _is_sheet_item(inst):
            items.append(inst)
    return items


# ── RUN ──────────────────────────────────────────────────────────────────────

picked = _picked()
if not picked:
    script.exit()
mode = forms.CommandSwitchWindow.show([FRONT, BACK],
                                      message='Move the picked views:')
if not mode:
    script.exit()

result_ids = []
with revit.Transaction('View Order: ' + mode):
    by_sheet = {}
    for it in picked:
        by_sheet.setdefault(_id_int(_sheet_id(it)), []).append(it)
    for sheet_key, vps in by_sheet.items():
        sheet = doc.GetElement(_sheet_id(vps[0]))
        if mode == FRONT:
            # Picked items, kept in their order among themselves.
            to_move = sorted(vps, key=lambda v: _id_int(v.Id))
        else:
            # Every other view and schedule on the sheet is placed again, in
            # its own order, so the picked ones end up the oldest: at the
            # back.
            picked_ids = set(_id_int(v.Id) for v in vps)
            to_move = sorted(
                (it for it in _sheet_items(sheet)
                 if _id_int(it.Id) not in picked_ids),
                key=lambda v: _id_int(v.Id))
        for it in to_move:
            new = _replace_item(it)
            if mode == FRONT:
                result_ids.append(new.Id)
        if mode == BACK:
            result_ids += [v.Id for v in vps]

uidoc.Selection.SetElementIds(List[DB.ElementId](result_ids))
