# -*- coding: utf-8 -*-
# find_type_in_views.py
from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List

doc   = revit.doc
uidoc = revit.uidoc

# ── GET TEXT NOTE TYPES ───────────────────────────────────────────────────────

tntypes    = list(DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType))
type_names = sorted(set(
    tn.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    for tn in tntypes
))

# ── SELECT TYPE TO SEARCH FOR ─────────────────────────────────────────────────

selected_type_name = forms.SelectFromList.show(
    type_names, title="Select Text Type to Find", multiselect=False)
if not selected_type_name:
    script.exit()

selected_type_ids = {
    tn.Id for tn in tntypes
    if tn.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == selected_type_name
}

# ── FIND MATCHING TEXT NOTES ──────────────────────────────────────────────────

all_textnotes  = DB.FilteredElementCollector(doc).OfClass(DB.TextNote).ToElements()
matching_texts = [tn for tn in all_textnotes if tn.TextNoteType.Id in selected_type_ids]

if not matching_texts:
    forms.alert(
        "No text notes found with type: {}".format(selected_type_name),
        exitscript=True
    )

# ── MAP TEXTS TO VIEWS ────────────────────────────────────────────────────────

view_dict = {}  # {view_id: {name, type, count, view}}

for text in matching_texts:
    view_id = text.OwnerViewId
    view    = doc.GetElement(view_id)
    if view:
        if view_id not in view_dict:
            view_type = (
                "Legend"
                if isinstance(view, DB.View) and view.ViewType == DB.ViewType.Legend
                else "View"
            )
            view_dict[view_id] = {
                "name":  view.Name,
                "type":  view_type,
                "count": 0,
                "view":  view
            }
        view_dict[view_id]["count"] += 1

# ── BUILD SELECTION LIST ──────────────────────────────────────────────────────

selection_items = [
    {
        "label":   "{}: {}".format(info["type"], info["name"]),
        "id":      view_id,
        "element": info["view"]
    }
    for view_id, info in sorted(view_dict.items(), key=lambda x: x[1]["name"])
]

if not selection_items:
    forms.alert(
        "No views found containing text type: {}".format(selected_type_name),
        exitscript=True
    )

# ── PICK VIEW TO OPEN ─────────────────────────────────────────────────────────

selected_label = forms.SelectFromList.show(
    [item["label"] for item in selection_items],
    title="Select View to Open ({} total locations)".format(len(selection_items)),
    multiselect=False
)

if not selected_label:
    script.exit()

# ── OPEN AND SELECT ───────────────────────────────────────────────────────────

selected_item = next(
    item for item in selection_items if item["label"] == selected_label)

uidoc.ActiveView = selected_item["element"]

view_id       = selected_item["id"]
texts_in_view = [t.Id for t in matching_texts if t.OwnerViewId == view_id]
element_ids   = List[DB.ElementId](texts_in_view)
uidoc.Selection.SetElementIds(element_ids)

script.exit()
