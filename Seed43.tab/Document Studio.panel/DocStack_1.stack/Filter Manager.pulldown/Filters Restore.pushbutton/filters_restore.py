# -*- coding: utf-8 -*-
# filters_restore.py
# pylint: disable=import-error,invalid-name,broad-except

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System.Collections.Generic import List
from System.IO import File
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ParameterFilterElement,
    ElementId,
)
from pyrevit import revit, forms, script

# ── [LIB] Snippets/_revisions.py ─────────────────────────────────────────────
from Snippets._revisions import get_backup_path, load_backup

doc    = revit.doc
logger = script.get_logger()

# ── FUNCTIONS ─────────────────────────────────────────────────────────────────

def get_live_filters():
    """Return a dict of {filter_name: ParameterFilterElement} from the project."""
    return {
        f.Name: f
        for f in FilteredElementCollector(doc).OfClass(ParameterFilterElement)
    }

def restore_filter(name, entry):
    """
    Recreate a ParameterFilterElement from a backup entry.

    Since only category IDs and a rules summary were stored, not full
    rule data, a minimal filter is created with the correct categories
    but no rules. Rules must be reconfigured manually after restore.

    Returns (True, None) on success, (False, reason_string) on failure.
    """
    try:
        cat_ids_raw = entry.get("data", {}).get("cats", [])
        if not cat_ids_raw:
            return False, "No categories in backup entry"

        cat_id_list = List[ElementId]()
        for cid in cat_ids_raw:
            try:
                cat_id_list.Add(ElementId(int(cid)))
            except Exception:
                pass

        if cat_id_list.Count == 0:
            return False, "Could not parse category IDs"

        ParameterFilterElement.Create(doc, name, cat_id_list)
        return True, None

    except Exception as ex:
        return False, type(ex).__name__

# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    backup_path = get_backup_path(doc)

    if not backup_path or not File.Exists(backup_path):
        forms.alert(
            "No backup file found next to this Revit model.\n\n"
            "Run Filter Delete and Backup first to create one.",
            title="Restore Filters"
        )
        return

    backup = load_backup(backup_path)
    if not backup:
        forms.alert("Backup file is empty or unreadable.", title="Restore Filters")
        return

    live = get_live_filters()

    # Build display list with status labels
    options = []
    for name in sorted(backup.keys()):
        date = backup[name].get("date", "unknown")
        if name not in live:
            options.append("{} [{}]  --  MISSING from project".format(name, date))
        else:
            options.append("{} [{}]  --  already exists".format(name, date))

    selected = forms.SelectFromList.show(
        options,
        title="Restore Filters - Select to Restore",
        button_name="Restore Selected",
        multiselect=True,
    )

    if not selected:
        return

    # Map display string back to filter name (name is everything before " [date]")
    to_restore           = []
    overwrite_candidates = []
    for display in selected:
        name = display.split(" [")[0]
        if name in live:
            overwrite_candidates.append(name)
        else:
            to_restore.append(name)

    if overwrite_candidates:
        overwrite = forms.alert(
            "{} filter{} already exist in the project:\n\n{}\n\nOverwrite?".format(
                len(overwrite_candidates),
                "s" if len(overwrite_candidates) != 1 else "",
                "\n".join(overwrite_candidates)
            ),
            yes=True,
            no=True,
            title="Overwrite Existing?"
        )
        if overwrite:
            to_restore.extend(overwrite_candidates)

    if not to_restore:
        return

    restored = 0
    skipped  = 0
    failed   = []

    with revit.Transaction("Restore Filters"):
        for name in to_restore:
            if name in live:
                try:
                    doc.Delete(live[name].Id)
                except Exception:
                    failed.append("{} - could not delete existing".format(name))
                    continue

            ok, err = restore_filter(name, backup[name])
            if ok:
                restored += 1
            else:
                failed.append("{} - {}".format(name, err or "unknown error"))

    msg = "Restored {} filter{}.".format(restored, "s" if restored != 1 else "")
    if skipped:
        msg += "\nSkipped: {}".format(skipped)
    if failed:
        msg += "\n\nFailed:\n" + "\n".join(failed)
    msg += (
        "\n\nNote: restored filters match original categories but rules "
        "may need to be reconfigured manually."
    )

    forms.alert(msg, title="Restore Filters")

if __name__ == "__main__":
    main()
