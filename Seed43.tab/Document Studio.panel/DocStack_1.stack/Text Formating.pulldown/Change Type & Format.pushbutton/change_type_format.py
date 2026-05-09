# -*- coding: utf-8 -*-
__title__  = "Change Text Type and Format"
__author__  = "Seed43"
__doc__     = """
𝐕𝐄𝐑𝐒𝐈𝐎𝐍 𝟐𝟔𝟎𝟓𝟎𝟏
_____________________________________________________________________
Description:
Batch updates Text Notes across the project by changing their Text
Type and applying formatting.

- Change Text Type (optional)
- Apply formatting:
    - Bold
    - Italic
    - Underline
- Works across the entire project

Worksharing support:
- Automatically checks out required elements
- Handles worksets safely
- Optional sync with central

Useful for:
- Global text standardisation
- QA cleanup
- Fixing inconsistent annotation styles
_____________________________________________________________________
How-to:
-> Run the tool
-> Select SOURCE Text Type (the type you want to change FROM)
-> Select TARGET Text Type (optional, the type to change TO)
-> Select formatting options (optional)

-> Tool will:
    - Update all matching Text Notes

Note:
- Formatting can be applied without changing type
- Type can be changed without applying formatting
_____________________________________________________________________
Last update:
- Initial release
_____________________________________________________________________
"""

from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List

doc    = revit.doc
uidoc  = revit.uidoc
output = script.get_output()


# ── GET TEXT NOTE TYPES ───────────────────────────────────────────────────────

tntypes    = list(DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType))
type_names = sorted(set(
    tn.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    for tn in tntypes
))


# ── SELECT SOURCE TYPE ────────────────────────────────────────────────────────

source_type_name = forms.SelectFromList.show(
    type_names, title="Pick SOURCE Text Type", multiselect=False)
if not source_type_name:
    script.exit()

source_type_ids = {
    tn.Id for tn in tntypes
    if tn.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == source_type_name
}


# ── SELECT TARGET TYPE ────────────────────────────────────────────────────────

target_type_name = forms.SelectFromList.show(
    type_names, title="Pick TARGET Text Type", multiselect=False)


# ── SELECT FORMATTING OPTIONS ─────────────────────────────────────────────────

format_options   = ["Bold", "Italic", "Underline"]
selected_formats = forms.SelectFromList.show(
    format_options, title="Select Formatting (optional):", multiselect=True)

if not target_type_name and not selected_formats:
    script.exit()

target_type_id = (
    next((tn.Id for tn in tntypes
          if tn.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == target_type_name),
         None)
    if target_type_name else None
)


# ── FIND SOURCE TEXTS ─────────────────────────────────────────────────────────

all_textnotes = DB.FilteredElementCollector(doc).OfClass(DB.TextNote).ToElements()
source_texts  = [tn for tn in all_textnotes if tn.TextNoteType.Id in source_type_ids]
total_texts   = len(source_texts)

if total_texts == 0:
    output.print_md("# No texts found with source type")
    script.exit()


# ── CHECKOUT FOR WORKSHARED MODELS ────────────────────────────────────────────

if doc.IsWorkshared:
    worksets_to_edit = set()
    for text in source_texts:
        workset_id = text.WorksetId
        workset    = doc.GetWorksetTable().GetWorkset(workset_id)
        if not workset.IsEditable:
            worksets_to_edit.add(workset_id)

    if worksets_to_edit:
        try:
            for ws_id in worksets_to_edit:
                DB.WorksharingUtils.CheckoutWorksets(
                    doc, List[DB.WorksetId]([ws_id]))
            import time
            time.sleep(0.5)
        except Exception as e:
            output.print_md("# Failed to make worksets editable: {}".format(str(e)))
            script.exit()

    element_ids_list = List[DB.ElementId]([t.Id for t in source_texts])
    try:
        DB.WorksharingUtils.CheckoutElements(doc, element_ids_list)
    except Exception:
        pass


# ── APPLY CHANGES ─────────────────────────────────────────────────────────────

if doc.IsWorkshared:
    tg = DB.TransactionGroup(doc, "Change Type and Format")
    tg.Start()

t = DB.Transaction(doc, "Change Type and Format")
t.Start()

try:
    for text in source_texts:
        if target_type_id and text.TextNoteType.Id != target_type_id:
            try:
                text.ChangeTypeId(target_type_id)
            except Exception:
                pass
            try:
                type_param = text.get_Parameter(DB.BuiltInParameter.SYMBOL_ID_PARAM)
                if type_param and not type_param.IsReadOnly:
                    type_param.Set(target_type_id)
            except Exception:
                pass

        if selected_formats:
            formatted_text = text.GetFormattedText()
            full_range     = DB.TextRange(0, len(text.Text))

            if "Bold" in selected_formats:
                formatted_text.SetBoldStatus(full_range, True)
            if "Italic" in selected_formats:
                formatted_text.SetItalicStatus(full_range, True)
            if "Underline" in selected_formats:
                formatted_text.SetUnderlineStatus(full_range, True)

            text.SetFormattedText(formatted_text)

    t.Commit()

    if doc.IsWorkshared:
        tg.Assimilate()
        try:
            sync_options  = DB.SynchronizeWithCentralOptions()
            trans_options = DB.TransactWithCentralOptions()
            doc.SynchronizeWithCentral(trans_options, sync_options)
        except Exception:
            pass
    else:
        try:
            doc.Save()
        except Exception:
            pass

except Exception as e:
    t.RollBack()
    if doc.IsWorkshared and "tg" in locals():
        tg.RollBack()
    output.print_md("# ERROR: {}".format(str(e)))
    import traceback
    output.print_md("```\n{}\n```".format(traceback.format_exc()))
    script.exit()

script.exit()
