# -*- coding: utf-8 -*-
__title__  = "Revision Status Colour"
__author__  = "Seed43"
__doc__     = """
𝐕𝐄𝐑𝐒𝐈𝐎𝐍 𝟐𝟔𝟎𝟓𝟎𝟏
_____________________________________________________________________
Description:
Colours revision clouds and their tags in all views based on whether
the parent revision has been marked as issued.

- Issued revisions:     clouds and tags shown in light gray
- Not issued revisions: clouds and tags shown in red

Works automatically across every view that contains revision clouds.
_____________________________________________________________________
How-to:
-> Run the tool
-> No selection required, runs automatically
-> The tool will:
    - Collect all revision clouds and tags across the model
    - Check whether each cloud belongs to an issued or not-issued
      revision
    - Apply graphic overrides to clouds and their tags per view
-> Result:
    - Issued clouds and tags shown in gray
    - Not-issued clouds and tags shown in red
    - A summary is printed on completion
_____________________________________________________________________
Notes:
- The tool checks whether a revision has been officially marked as
  issued, not by looking at its name
- You cannot change the text colour of a tag individually in Revit,
  any colour change affects the whole tag appearance
- Works on all views containing clouds, not limited to the active view
- All changes are made in a single transaction and can be undone
_____________________________________________________________________
Last update:
- Initial release
_____________________________________________________________________
"""

from pyrevit import revit, DB, script

# ── [LIB] Snippets/_revisions.py ─────────────────────────────────────────────
from Snippets._revisions import get_revision_description

doc    = revit.doc
output = script.get_output()


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

output.print_html("<div class='header'>REVISION STATUS COLOUR ENGINE INITIALISING</div>")
output.print_html("<div class='line'>------------------------------------</div>")


# ── PRE-CACHE REVISION NAMES ──────────────────────────────────────────────────

rev_name_cache = {
    r.Id: get_revision_description(r)
    for r in DB.FilteredElementCollector(doc).OfClass(DB.Revision).ToElements()
}


# ── COLLECT AND CLASSIFY CLOUDS ───────────────────────────────────────────────

clouds = (
    DB.FilteredElementCollector(doc)
      .OfCategory(DB.BuiltInCategory.OST_RevisionClouds)
      .WhereElementIsNotElementType()
      .ToElements()
)

tags = (
    DB.FilteredElementCollector(doc)
      .OfCategory(DB.BuiltInCategory.OST_RevisionCloudTags)
      .WhereElementIsNotElementType()
      .ToElements()
)

cloud_issued = {}

for cloud in clouds:
    try:
        rev_id = cloud.RevisionId
        if rev_id == DB.ElementId.InvalidElementId:
            cloud_issued[cloud.Id] = False
            continue
        rev = doc.GetElement(rev_id)
        cloud_issued[cloud.Id] = bool(rev.Issued) if rev else False
    except Exception as e:
        output.print_html(
            "<div class='warn'>WARNING (cloud classification): {}</div>".format(str(e)))
        cloud_issued[cloud.Id] = False


# ── GRAPHIC OVERRIDES ─────────────────────────────────────────────────────────

issued_color     = DB.Color(192, 192, 192)  # Light gray
not_issued_color = DB.Color(255, 0,   0)    # Red

ogs_issued = DB.OverrideGraphicSettings()
ogs_issued.SetProjectionLineColor(issued_color)

ogs_not_issued = DB.OverrideGraphicSettings()
ogs_not_issued.SetProjectionLineColor(not_issued_color)


# ── APPLY OVERRIDES ───────────────────────────────────────────────────────────

issued_count     = 0
not_issued_count = 0
logged_revs      = set()
last_sheet_name  = None

with revit.Transaction("Colour Revision Clouds"):

    for cloud in clouds:
        try:
            view_id = cloud.OwnerViewId
            if view_id == DB.ElementId.InvalidElementId:
                continue
            view = doc.GetElement(view_id)
            if not view:
                continue

            is_issued = cloud_issued.get(cloud.Id, False)
            ogs       = ogs_issued if is_issued else ogs_not_issued
            view.SetElementOverrides(cloud.Id, ogs)

            rev      = doc.GetElement(cloud.RevisionId)
            rev_name = rev_name_cache.get(cloud.RevisionId, "?") if rev else "?"

            key = (view.Name, rev_name)
            if key not in logged_revs:
                logged_revs.add(key)
                if view.Name != last_sheet_name:
                    output.print_html(
                        "<div class='sheet'>Sheet {}</div>".format(view.Name))
                    last_sheet_name = view.Name
                if is_issued:
                    output.print_html(
                        "<div class='rev'>Issued (gray) Rev: {}</div>".format(rev_name))
                else:
                    output.print_html(
                        "<div class='rev'>Not Issued (red) Rev: {}</div>".format(rev_name))

            if is_issued:
                issued_count += 1
            else:
                not_issued_count += 1

        except Exception as e:
            output.print_html(
                "<div class='warn'>WARNING (cloud override): {}</div>".format(str(e)))

    # Note: GetTaggedLocalElementIds() may be deprecated in Revit 2026,
    # check against the SDK if tags stop being coloured correctly.
    for tag in tags:
        try:
            view_id = tag.OwnerViewId
            if view_id == DB.ElementId.InvalidElementId:
                continue
            view = doc.GetElement(view_id)
            if not view:
                continue

            tagged_ids = tag.GetTaggedLocalElementIds()
            if not tagged_ids:
                continue

            tagged_cloud_id = list(tagged_ids)[0]
            is_issued       = cloud_issued.get(tagged_cloud_id, False)
            ogs             = ogs_issued if is_issued else ogs_not_issued
            view.SetElementOverrides(tag.Id, ogs)

        except Exception as e:
            output.print_html(
                "<div class='warn'>WARNING (tag override): {}</div>".format(str(e)))


# ── SUMMARY ───────────────────────────────────────────────────────────────────

output.print_html("<div class='line'>------------------------------------</div>")
output.print_html("<div class='header'>COMPLETE</div>")
output.print_html(
    "<div class='sheet'>Issued Clouds Coloured (gray): {}</div>".format(issued_count))
output.print_html(
    "<div class='sheet'>Not-Issued Clouds Coloured (red): {}</div>".format(not_issued_count))
