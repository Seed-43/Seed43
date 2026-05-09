# -*- coding: utf-8 -*-
__title__  = "Persistent Revisions"
__author__  = "Seed43"
__doc__     = """
𝐕𝐄𝐑𝐒𝐈𝐎𝐍 𝟐𝟔𝟎𝟓𝟎𝟏
_____________________________________________________________________
Description:
Permanently links revision clouds to their parent sheets, ensuring
revision history is never lost regardless of where the cloud is placed.

- Scans ALL revision clouds in the model
- Detects whether each cloud lives on a sheet or a hosted view
- Propagates revisions to the correct sheet automatically
- Never removes existing sheet revisions (additive only)
- Skips placeholder sheets safely
_____________________________________________________________________
How-to:
-> Run the tool
-> No selection required, runs automatically
-> The tool will:
    - Collect all revision clouds across the model
    - Map each cloud to its parent sheet
    - Apply any missing revisions to the sheet revision list
-> Result:
    - All sheets updated with their correct revisions
    - A live log shows every sheet and revision linked
    - A summary is printed on completion
_____________________________________________________________________
Notes:
- Revision clouds on views hosted on sheets are correctly propagated
- Placeholder sheets are skipped as they cannot hold revisions
- Existing sheet revisions are preserved, nothing is removed
- All changes are made in a single transaction and can be undone
_____________________________________________________________________
Last update:
- Initial release
_____________________________________________________________________
"""

from pyrevit import DB, script
from collections import defaultdict
from System.Collections.Generic import List

# ── [LIB] Snippets/_revisions.py ─────────────────────────────────────────────
from Snippets._revisions import get_revision_description

doc    = __revit__.ActiveUIDocument.Document
output = script.get_output()

INVALID_ID = DB.ElementId.InvalidElementId


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

output.print_html("<div class='header'>PERSISTENT REVISION ENGINE INITIALISING</div>")
output.print_html("<div class='line'>------------------------------------</div>")


# ── PRE-CACHE REVISIONS ───────────────────────────────────────────────────────

rev_name_cache = {
    r.Id: get_revision_description(r)
    for r in DB.FilteredElementCollector(doc).OfClass(DB.Revision).ToElements()
}


# ── VIEW TO SHEET MAP ─────────────────────────────────────────────────────────

view_to_sheets = defaultdict(list)

for vp in DB.FilteredElementCollector(doc).OfClass(DB.Viewport).ToElements():
    try:
        if vp.SheetId != INVALID_ID:
            view_to_sheets[vp.ViewId].append(vp.SheetId)
    except Exception as e:
        output.print_html(
            "<div class='warn'>WARNING (viewport mapping): {}</div>".format(str(e)))


# ── COLLECT REVISION CLOUDS ───────────────────────────────────────────────────

sheet_to_revs = defaultdict(set)

for c in (
    DB.FilteredElementCollector(doc)
      .OfCategory(DB.BuiltInCategory.OST_RevisionClouds)
      .WhereElementIsNotElementType()
      .ToElements()
):
    try:
        rev_id = c.RevisionId
        if rev_id == INVALID_ID:
            continue

        owner_id = c.OwnerViewId
        owner    = doc.GetElement(owner_id)

        if isinstance(owner, DB.ViewSheet):
            sheet_to_revs[owner.Id].add(rev_id)
        elif owner_id in view_to_sheets:
            for sid in view_to_sheets[owner_id]:
                sheet_to_revs[sid].add(rev_id)

    except Exception as e:
        output.print_html(
            "<div class='warn'>WARNING (cloud collection): {}</div>".format(str(e)))


# ── APPLY ─────────────────────────────────────────────────────────────────────

updated_sheets = 0
total_added    = 0
log_entries    = []

all_sheets = (
    DB.FilteredElementCollector(doc)
      .OfClass(DB.ViewSheet)
      .ToElements()
)

t = DB.Transaction(doc, "Persistent Revisions Engine")
t.Start()

try:
    for sheet in all_sheets:
        if sheet.IsPlaceholder:
            continue

        current  = set(sheet.GetAdditionalRevisionIds())
        detected = sheet_to_revs.get(sheet.Id, set())
        missing  = detected - current

        if not missing:
            continue

        sheet.SetAdditionalRevisionIds(List[DB.ElementId](list(current | missing)))

        rev_names = [rev_name_cache.get(rid, "?") for rid in missing]
        log_entries.append((sheet.SheetNumber, rev_names))

        updated_sheets += 1
        total_added    += len(missing)

    t.Commit()

except Exception as e:
    t.RollBack()
    output.print_html(
        "<div class='warn'>ERROR (transaction rolled back): {}</div>".format(str(e)))


# ── LIVE LOG ──────────────────────────────────────────────────────────────────

for sheet_number, rev_names in log_entries:
    output.print_html("<div class='sheet'>Sheet {}</div>".format(sheet_number))
    for name in rev_names:
        output.print_html("<div class='rev'>Linked Rev: {}</div>".format(name))


# ── SUMMARY ───────────────────────────────────────────────────────────────────

output.print_html("<div class='line'>------------------------------------</div>")
output.print_html("<div class='header'>COMPLETE</div>")
output.print_html(
    "<div class='sheet'>Sheets Updated: {}</div>".format(updated_sheets))
output.print_html(
    "<div class='sheet'>Revisions Linked to Sheets: {}</div>".format(total_added))
