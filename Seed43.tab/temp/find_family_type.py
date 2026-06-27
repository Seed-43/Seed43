# -*- coding: utf-8 -*-
# find_family_type.py
from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List

doc   = revit.doc
uidoc = revit.uidoc

# ── GET ALL LOADED FAMILIES ───────────────────────────────────────────────────

all_symbols = list(
    DB.FilteredElementCollector(doc)
    .OfClass(DB.FamilySymbol)
    .ToElements()
)

# Build a dict: family name -> sorted list of type names
family_map = {}
for sym in all_symbols:
    fam_name  = sym.Family.Name
    type_name = sym.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if fam_name not in family_map:
        family_map[fam_name] = []
    if type_name not in family_map[fam_name]:
        family_map[fam_name].append(type_name)

for fam_name in family_map:
    family_map[fam_name] = sorted(family_map[fam_name])

# ── SELECT FAMILY ─────────────────────────────────────────────────────────────

sorted_family_names = sorted(family_map.keys())

selected_family_name = forms.SelectFromList.show(
    sorted_family_names,
    title="Step 1 of 2 — Select Family",
    multiselect=False
)
if not selected_family_name:
    script.exit()

# ── SELECT TYPE WITHIN THAT FAMILY ───────────────────────────────────────────

type_choices = ["(All Types)"] + family_map[selected_family_name]

selected_type_name = forms.SelectFromList.show(
    type_choices,
    title="Step 2 of 2 — Select Type  [Family: {}]".format(selected_family_name),
    multiselect=False
)
if not selected_type_name:
    script.exit()

# ── RESOLVE MATCHING SYMBOL IDs ───────────────────────────────────────────────

if selected_type_name == "(All Types)":
    matching_symbol_ids = {
        sym.Id for sym in all_symbols
        if sym.Family.Name == selected_family_name
    }
else:
    matching_symbol_ids = {
        sym.Id for sym in all_symbols
        if sym.Family.Name == selected_family_name
        and sym.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == selected_type_name
    }

# ── FIND INSTANCES ────────────────────────────────────────────────────────────

all_instances = (
    DB.FilteredElementCollector(doc)
    .OfClass(DB.FamilyInstance)
    .ToElements()
)

matching_instances = [
    fi for fi in all_instances
    if fi.GetTypeId() in matching_symbol_ids
]

if not matching_instances:
    label = (
        selected_family_name
        if selected_type_name == "(All Types)"
        else "{} : {}".format(selected_family_name, selected_type_name)
    )
    forms.alert(
        "No instances found for: {}".format(label),
        exitscript=True
    )

# ── MAP INSTANCES TO VIEWS ────────────────────────────────────────────────────

view_dict = {}  # {view_id: {name, type, count, view, instances}}

for fi in matching_instances:
    view_id = fi.OwnerViewId

    # FamilyInstances placed in 3D or not owned by a view return InvalidElementId
    if view_id == DB.ElementId.InvalidElementId:
        view_id_key = DB.ElementId.InvalidElementId
        if view_id_key not in view_dict:
            view_dict[view_id_key] = {
                "name":      "(Model — not view-specific)",
                "type":      "Model",
                "count":     0,
                "view":      None,
                "instances": []
            }
        view_dict[view_id_key]["count"] += 1
        view_dict[view_id_key]["instances"].append(fi)
        continue

    view = doc.GetElement(view_id)
    if view is None:
        continue

    if view_id not in view_dict:
        if isinstance(view, DB.View) and view.ViewType == DB.ViewType.Legend:
            view_type = "Legend"
        elif isinstance(view, DB.View) and view.ViewType == DB.ViewType.DraftingView:
            view_type = "Drafting"
        elif isinstance(view, DB.ViewSheet):
            view_type = "Sheet"
        else:
            view_type = "View"

        view_dict[view_id] = {
            "name":      view.Name,
            "type":      view_type,
            "count":     0,
            "view":      view,
            "instances": []
        }

    view_dict[view_id]["count"]     += 1
    view_dict[view_id]["instances"].append(fi)

# ── BUILD SELECTION LIST ──────────────────────────────────────────────────────

# Model-level instances (not view-owned) go at the top
selection_items = []

if DB.ElementId.InvalidElementId in view_dict:
    info = view_dict[DB.ElementId.InvalidElementId]
    selection_items.append({
        "label":     "{} ({} instance{})".format(
            info["name"], info["count"], "s" if info["count"] != 1 else ""),
        "id":        DB.ElementId.InvalidElementId,
        "view":      None,
        "instances": info["instances"]
    })

for view_id, info in sorted(view_dict.items(), key=lambda x: x[1]["name"]):
    if view_id == DB.ElementId.InvalidElementId:
        continue
    selection_items.append({
        "label":     "{}: {} ({} instance{})".format(
            info["type"], info["name"],
            info["count"], "s" if info["count"] != 1 else ""),
        "id":        view_id,
        "view":      info["view"],
        "instances": info["instances"]
    })

if not selection_items:
    forms.alert(
        "No views found containing the selected family type.",
        exitscript=True
    )

total_instances = sum(info["count"] for info in view_dict.values())

# ── PICK VIEW TO OPEN ─────────────────────────────────────────────────────────

selected_label = forms.SelectFromList.show(
    [item["label"] for item in selection_items],
    title="Select View to Open  [{} total instances]".format(total_instances),
    multiselect=False
)

if not selected_label:
    script.exit()

# ── OPEN VIEW AND SELECT INSTANCES ───────────────────────────────────────────

selected_item = next(
    item for item in selection_items if item["label"] == selected_label)

if selected_item["view"] is not None:
    uidoc.ActiveView = selected_item["view"]

instance_ids  = List[DB.ElementId]([fi.Id for fi in selected_item["instances"]])
uidoc.Selection.SetElementIds(instance_ids)

script.exit()
