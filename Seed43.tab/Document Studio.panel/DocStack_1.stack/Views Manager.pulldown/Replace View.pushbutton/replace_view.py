# -*- coding: utf-8 -*-
"""
Replace View - swap a view on a sheet for another one, in the same place.

Pick a viewport (or schedule) on the current sheet, or have one selected
before running, then choose the view to put there instead. The new view
keeps the old one's centre, viewport type, detail number, title position
and rotation; a schedule keeps its insertion point and rotation.
"""
from pyrevit import revit, DB, forms, script
from pyrevit.forms import WarningBar
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

doc = revit.doc
uidoc = revit.uidoc


# ── PICK THE VIEWPORT ────────────────────────────────────────────────────────

class _SheetItemFilter(ISelectionFilter):
    """Only viewports and schedules on the sheet can be picked."""
    def AllowElement(self, el):
        return isinstance(el, (DB.Viewport, DB.ScheduleSheetInstance)) \
            and not getattr(el, 'IsTitleblockRevisionSchedule', False)

    def AllowReference(self, ref, pt):
        return False


def _picked_item():
    """The viewport or schedule to replace: the current selection if it is
    one, otherwise picked on the active sheet."""
    for eid in uidoc.Selection.GetElementIds():
        el = doc.GetElement(eid)
        if _SheetItemFilter().AllowElement(el):
            return el
    if revit.active_view.ViewType != DB.ViewType.DrawingSheet:
        forms.alert('Open a sheet, or select a viewport on one, first.',
                    exitscript=True)
    with WarningBar(title='Pick the view to replace. Esc to cancel.'):
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element,
                                             _SheetItemFilter(),
                                             'Pick the view to replace')
        except Exception:
            script.exit()
    return doc.GetElement(ref.ElementId)


old = _picked_item()
sheet = doc.GetElement(old.OwnerViewId)
is_schedule = isinstance(old, DB.ScheduleSheetInstance)
old_view = doc.GetElement(old.ScheduleId if is_schedule else old.ViewId)


# ── CHOOSE THE REPLACEMENT ───────────────────────────────────────────────────

_KIND = {
    DB.ViewType.FloorPlan: 'Floor Plan', DB.ViewType.CeilingPlan: 'Ceiling Plan',
    DB.ViewType.EngineeringPlan: 'Structural Plan', DB.ViewType.AreaPlan: 'Area Plan',
    DB.ViewType.Elevation: 'Elevation', DB.ViewType.Section: 'Section',
    DB.ViewType.Detail: 'Detail', DB.ViewType.DraftingView: 'Drafting',
    DB.ViewType.Legend: 'Legend', DB.ViewType.ThreeD: '3D',
    DB.ViewType.Rendering: 'Rendering', DB.ViewType.Schedule: 'Schedule',
    DB.ViewType.PanelSchedule: 'Panel Schedule',
}


class ViewItem(forms.TemplateListItem):
    @property
    def name(self):
        return u'[{}]  {}'.format(_KIND.get(self.item.ViewType,
                                            str(self.item.ViewType)),
                                  self.item.Name)


def _candidates():
    """Views that can go where the old one is. A schedule is swapped for
    another schedule; a viewport for any view Revit will let onto this
    sheet (unplaced views, plus legends, which can sit on many sheets)."""
    out = []
    for v in DB.FilteredElementCollector(doc).OfClass(DB.View):
        if v.IsTemplate or v.Id == old_view.Id:
            continue
        if is_schedule:
            if (isinstance(v, DB.ViewSchedule)
                    and not v.IsTitleblockRevisionSchedule
                    and not v.IsInternalKeynoteSchedule):
                out.append(v)
        elif v.ViewType not in (DB.ViewType.Schedule, DB.ViewType.DrawingSheet,
                                DB.ViewType.ProjectBrowser,
                                DB.ViewType.SystemBrowser, DB.ViewType.Internal,
                                DB.ViewType.Undefined):
            try:
                if DB.Viewport.CanAddViewToSheet(doc, sheet.Id, v.Id):
                    out.append(v)
            except Exception:
                pass
    return sorted((ViewItem(v) for v in out), key=lambda i: i.name.lower())


items = _candidates()
if not items:
    forms.alert('No view can be placed on this sheet in its place.',
                exitscript=True)
new_view = forms.SelectFromList.show(
    items,
    title=u'Replace "{}" with'.format(old_view.Name),
    button_name='Replace',
    multiselect=False,
)
if not new_view:
    script.exit()


# ── REPLACE ──────────────────────────────────────────────────────────────────

with revit.Transaction('Replace View'):
    if is_schedule:
        point, rotation = old.Point, old.Rotation
        doc.Delete(old.Id)
        new = DB.ScheduleSheetInstance.Create(doc, sheet.Id, new_view.Id, point)
        try:
            new.Rotation = rotation
        except Exception:
            pass
    else:
        center = old.GetBoxCenter()
        type_id = old.GetTypeId()
        rotation = old.Rotation
        p = old.get_Parameter(DB.BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
        detail = p.AsString() if p else None
        try:
            label_offset, label_len = old.LabelOffset, old.LabelLineLength
        except Exception:
            label_offset = label_len = None
        sheet.DeleteViewport(old)
        new = DB.Viewport.Create(doc, sheet.Id, new_view.Id, center)
        new.ChangeTypeId(type_id)
        if rotation != getattr(DB.ViewportRotation, "None"):
            new.Rotation = rotation
        # The box can change size with rotation: put the centre back.
        new.SetBoxCenter(center)
        if detail:
            try:
                new.get_Parameter(
                    DB.BuiltInParameter.VIEWPORT_DETAIL_NUMBER).Set(detail)
            except Exception:
                pass  # the number is taken by another view on the sheet
        try:
            if label_offset is not None:
                new.LabelOffset = label_offset
            if label_len:
                new.LabelLineLength = label_len
        except Exception:
            pass

uidoc.Selection.SetElementIds(
    __import__('System').Collections.Generic.List[DB.ElementId]([new.Id]))
