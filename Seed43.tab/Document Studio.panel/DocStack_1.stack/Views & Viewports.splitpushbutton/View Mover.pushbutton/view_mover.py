# -*- coding: utf-8 -*-
# view_mover.py
from pyrevit import revit, DB, forms, script
from pyrevit.forms import WarningBar

doc = revit.doc

# ── VALIDATE ACTIVE SHEET ─────────────────────────────────────────────────────

cursheet = revit.active_view
if cursheet.ViewType != DB.ViewType.DrawingSheet:
    forms.alert("Please run this tool from a sheet view.", exitscript=True)

# ── SELECT TARGET SHEET ───────────────────────────────────────────────────────

dest_sheet = forms.select_sheets(
    title="Select Target Sheet",
    button_name="Select Sheet",
    multiple=False,
    include_placeholder=False,
    use_selection=True
)

if not dest_sheet:
    forms.alert("You must select a target sheet.", exitscript=True)

# ── SELECT VIEWPORTS ──────────────────────────────────────────────────────────

with WarningBar(title="Select viewports or schedules to move. Click Finish when done."):
    selected_elements = revit.pick_elements()

if not selected_elements:
    script.exit()

selected_vports = [
    el for el in selected_elements
    if isinstance(el, (DB.Viewport, DB.ScheduleSheetInstance))
]

if not selected_vports:
    forms.alert("At least one viewport or schedule must be selected.", exitscript=True)

# ── FUNCTIONS ─────────────────────────────────────────────────────────────────

def get_existing_data(sheet):
    """Return existing detail numbers and view IDs on the given sheet."""
    existing_numbers  = set()
    existing_view_ids = set()
    for vp_id in sheet.GetAllViewports():
        vp           = doc.GetElement(vp_id)
        detail_param = vp.LookupParameter("Detail Number")
        if detail_param:
            val = detail_param.AsString()
            if val:
                existing_numbers.add(val)
        existing_view_ids.add(vp.ViewId)
    return existing_numbers, existing_view_ids

# ── GET EXISTING DATA ─────────────────────────────────────────────────────────

existing_detail_numbers, existing_view_ids = get_existing_data(dest_sheet)

# ── MOVE ELEMENTS ─────────────────────────────────────────────────────────────

skipped_duplicates      = []
auto_prefixed_viewports = []
failed_moves            = []

t = DB.Transaction(doc, "Move Viewports Preserve Layout")
t.Start()

try:
    for vp in selected_vports:

        # ── Move viewport ─────────────────────────────────────────────────────

        if isinstance(vp, DB.Viewport):
            try:
                view_id   = vp.ViewId
                view_name = doc.GetElement(view_id).Name

                if view_id in existing_view_ids:
                    skipped_duplicates.append(
                        "{} (already on target sheet)".format(view_name))
                    continue

                vp_center  = vp.GetBoxCenter()
                vp_type_id = vp.GetTypeId()

                detail_param    = vp.LookupParameter("Detail Number")
                detail_num      = detail_param.AsString() if detail_param else None
                original_detail = detail_num

                if detail_num and detail_num in existing_detail_numbers:
                    detail_num = "**{}".format(detail_num)
                    counter    = 1
                    while detail_num in existing_detail_numbers:
                        detail_num = "**{}{}".format("*" * counter, original_detail)
                        counter += 1
                    auto_prefixed_viewports.append(
                        "{} to {}".format(view_name, detail_num))
                    existing_detail_numbers.add(detail_num)

                label_offset      = None
                label_line_length = None
                try:
                    label_offset      = vp.LabelOffset
                    label_line_length = vp.LabelLineLength
                except Exception:
                    pass

                doc.GetElement(vp.SheetId).DeleteViewport(vp)

                new_vp = DB.Viewport.Create(doc, dest_sheet.Id, view_id, vp_center)
                new_vp.ChangeTypeId(vp_type_id)

                if detail_num:
                    param = new_vp.LookupParameter("Detail Number")
                    if param and param.StorageType == DB.StorageType.String:
                        param.Set(detail_num)

                try:
                    if label_offset:
                        new_vp.LabelOffset = label_offset
                    if label_line_length:
                        new_vp.LabelLineLength = label_line_length
                except Exception:
                    pass

            except Exception as e:
                failed_moves.append("{}: {}".format(view_name, str(e)))

        # ── Move schedule ─────────────────────────────────────────────────────

        elif isinstance(vp, DB.ScheduleSheetInstance):
            try:
                schedule_id = vp.ScheduleId
                point       = vp.Point
                DB.ScheduleSheetInstance.Create(doc, dest_sheet.Id, schedule_id, point)
                doc.Delete(vp.Id)
            except Exception as e:
                try:
                    name = doc.GetElement(schedule_id).Name
                except Exception:
                    name = "Unknown Schedule"
                failed_moves.append("{}: {}".format(name, str(e)))

    t.Commit()

except Exception as e:
    t.RollBack()
    forms.alert("Transaction failed: {}".format(str(e)), title="Error")
    script.exit()

# ── RESULT ────────────────────────────────────────────────────────────────────

success_count = len(selected_vports) - len(skipped_duplicates) - len(failed_moves)
msg           = "Operation Completed."

if success_count > 0:
    msg += "\n\nMOVED: {} item(s)".format(success_count)

if auto_prefixed_viewports:
    msg += "\n\nAUTO-PREFIXED:\n" + "\n".join(
        ["- " + x for x in auto_prefixed_viewports])

if skipped_duplicates:
    msg += "\n\nSKIPPED:\n" + "\n".join(
        ["- " + x for x in skipped_duplicates])

if failed_moves:
    msg += "\n\nFAILED:\n" + "\n".join(
        ["- " + x for x in failed_moves])

forms.alert(msg, title="Completed")
