# -*- coding: utf-8 -*-
__title__  = "Write Cloud Parameter"
__author__  = "Seed43"
__doc__     = """
𝐕𝐄𝐑𝐒𝐈𝐎𝐍 𝟐𝟔𝟎𝟓𝟎𝟏
_____________________________________________________________________
Description:
Writes sheet and revision data directly to the built-in Revision Cloud
parameters so clouds carry their context wherever they appear.

Writes: Sheet Number + Revision Number to the Mark parameter
Example result: "S202 - C"

Works for:
- Clouds placed directly on sheets
- Clouds inside views that are placed on sheets
_____________________________________________________________________
How-to:
-> Run the tool
-> No selection required, runs automatically
-> The tool will:
    - Collect all revision clouds in the model
    - Resolve each cloud's parent sheet
    - Write the sheet number and revision number to the Mark parameter
-> Result:
    - All reachable clouds are updated
    - Skipped clouds are logged with the reason
    - A summary is printed on completion
_____________________________________________________________________
Notes:
- Uses the built-in Mark parameter, no project parameters required
- Mark is written as "Sheet - RevNumber" (for example "S202 - C")
- The revision number shown is the label visible on the sheet,
  not the internal sequence number
- Clouds on views that are not placed on any sheet are skipped
- Clouds whose workset is not editable are logged and skipped
- All writes happen in a single transaction and can be undone
_____________________________________________________________________
Last update:
- Initial release
_____________________________________________________________________
"""

from pyrevit import revit, DB, script

doc    = revit.doc
output = script.get_output()

INVALID_ID = DB.ElementId.InvalidElementId
BIP_MARK   = DB.BuiltInParameter.ALL_MODEL_MARK


# ── UI STYLE ──────────────────────────────────────────────────────────────────

output.add_style("""
body {
    background-color: #232933;
    color: #F4FAFF;
    font-family: Consolas, Courier New, monospace;
    padding: 20px;
}
.header { color: #2B933F; font-weight: bold; font-size: 1.2em; }
.sheet  { color: #F4FAFF; padding-left: 15px; }
.rev    { color: #8B9199; padding-left: 30px; }
.warn   { color: #E0A040; padding-left: 30px; }
.line   { color: #3B4553; }
""")

output.print_html("<div class='header'>WRITE CLOUD PARAMETER ENGINE</div>")
output.print_html("<div class='line'>------------------------------------</div>")


# ── COLLECT CLOUDS ────────────────────────────────────────────────────────────
# OfClass(DB.RevisionCloud) is not valid, clouds are AnnotationSymbol elements
# filtered by the OST_RevisionClouds category.

clouds = (
    DB.FilteredElementCollector(doc)
      .OfCategory(DB.BuiltInCategory.OST_RevisionClouds)
      .WhereElementIsNotElementType()
      .ToElements()
)

if not clouds:
    output.print_html("<div class='warn'>No revision clouds found in the project.</div>")
    script.exit()

output.print_html("<div class='sheet'>Found {} revision clouds</div>".format(len(clouds)))
output.print_html("<div class='sheet'>Writing to: Mark (sheet number - revision number)</div>")
output.print_html("<div class='line'>------------------------------------</div>")


# ── PRE-CACHE VIEW TO SHEET MAP ───────────────────────────────────────────────
# Build ViewId to SheetId lookup in one pass so the cloud loop
# does not run a collector per cloud.

view_to_sheet = {}

for vp in (
    DB.FilteredElementCollector(doc)
      .OfClass(DB.Viewport)
      .ToElements()
):
    try:
        if vp.SheetId != INVALID_ID:
            view_to_sheet[vp.ViewId] = vp.SheetId
    except Exception:
        pass

output.print_html(
    "<div class='sheet'>Mapped {} views to sheets</div>".format(len(view_to_sheet)))


# ── PRE-CACHE REVISION NUMBERS ────────────────────────────────────────────────
# PROJECT_REVISION_REVISION_NUM is the user-facing revision number shown on
# the sheet (for example "C" or "3"), not the internal sequence number.

rev_num_cache = {}

for rev in DB.FilteredElementCollector(doc).OfClass(DB.Revision).ToElements():
    try:
        param = rev.get_Parameter(DB.BuiltInParameter.PROJECT_REVISION_REVISION_NUM)
        if param:
            rev_num_cache[rev.Id] = param.AsString() or ""
    except Exception:
        pass


# ── APPLY ─────────────────────────────────────────────────────────────────────

updated          = 0
skipped_no_sheet = 0
skipped_readonly = 0
skipped_other    = 0

with revit.Transaction("Write Sheet and Sequence to Revision Clouds"):

    for cloud in clouds:
        try:
            owner_id = cloud.OwnerViewId
            if owner_id == INVALID_ID:
                skipped_other += 1
                continue

            owner = doc.GetElement(owner_id)
            if not owner:
                skipped_other += 1
                continue

            # Resolve the parent sheet
            sheet = None
            if isinstance(owner, DB.ViewSheet):
                sheet = owner
            else:
                sheet_id = view_to_sheet.get(owner_id)
                if sheet_id:
                    sheet = doc.GetElement(sheet_id)

            if not sheet:
                skipped_no_sheet += 1
                continue

            sheet_num = sheet.SheetNumber or ""
            rev_num   = rev_num_cache.get(cloud.RevisionId, "")

            # Combine into a single Mark value: "S202 - C"
            mark_value = "{} - {}".format(sheet_num, rev_num) if rev_num else sheet_num

            param_mark = cloud.get_Parameter(BIP_MARK)
            changed    = False

            if param_mark and mark_value and not param_mark.IsReadOnly:
                param_mark.Set(mark_value)
                changed = True

            if changed:
                updated += 1
                output.print_html("<div class='sheet'>Sheet {}</div>".format(sheet_num))
                output.print_html("<div class='rev'>Mark: {}</div>".format(mark_value))
            else:
                skipped_other += 1

        except Exception as ex:
            skipped_readonly += 1
            ws_name = "N/A"
            try:
                if doc.IsWorkshared:
                    ws = doc.GetWorksetTable().GetWorkset(cloud.WorksetId)
                    ws_name = ws.Name if ws else "???"
            except Exception:
                pass
            output.print_html(
                "<div class='warn'>WARNING: Cloud {} | Workset: {} | {}</div>".format(
                    cloud.Id, ws_name, str(ex)))


# ── SUMMARY ───────────────────────────────────────────────────────────────────

output.print_html("<div class='line'>------------------------------------</div>")
output.print_html("<div class='header'>COMPLETE</div>")
output.print_html(
    "<div class='sheet'>Clouds updated:              {}</div>".format(updated))
output.print_html(
    "<div class='sheet'>Skipped (no sheet):          {}</div>".format(skipped_no_sheet))
output.print_html(
    "<div class='sheet'>Skipped (read-only):         {}</div>".format(skipped_readonly))
output.print_html(
    "<div class='sheet'>Skipped (other):             {}</div>".format(skipped_other))
