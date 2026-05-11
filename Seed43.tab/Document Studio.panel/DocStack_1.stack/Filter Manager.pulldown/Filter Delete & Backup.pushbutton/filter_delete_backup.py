# -*- coding: utf-8 -*-
# filter_delete_backup.py
# pylint: disable=import-error,invalid-name,broad-except

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

import datetime

from Autodesk.Revit.DB import FilteredElementCollector, ParameterFilterElement
from pyrevit import revit, forms, script

# ── [LIB] Snippets/_revisions.py ─────────────────────────────────────────────
from Snippets._revisions import (
    safe_str,
    get_backup_path,
    load_backup,
    save_backup,
)

doc    = revit.doc
logger = script.get_logger()

TODAY = datetime.datetime.now().strftime("%Y-%m-%d")

# ── FUNCTIONS ─────────────────────────────────────────────────────────────────

def serialise_filter(f):
    """Convert a filter to a plain dict for JSON storage."""
    try:
        cat_ids = sorted([
            str(getattr(cid, "Value", None) or getattr(cid, "IntegerValue", 0))
            for cid in f.GetCategories()
        ])
    except Exception:
        cat_ids = []
    try:
        rules_str = safe_str(f.GetElementFilter().ToString())
    except Exception:
        rules_str = ""
    return {
        "name":  safe_str(f.Name),
        "cats":  cat_ids,
        "rules": rules_str,
    }

def sync_backup(backup, live_filters):
    """Update the backup dict with any new or changed filters."""
    updated = False
    for f in live_filters:
        name    = safe_str(f.Name)
        current = serialise_filter(f)

        if name not in backup:
            backup[name] = {"date": TODAY, "data": current}
            updated = True
        else:
            stored = backup[name]["data"]
            same   = (
                stored["cats"]  == current["cats"] and
                stored["rules"] == current["rules"]
            )
            if not same:
                old_date    = backup[name]["date"]
                archive_key = "{} [{}]".format(name, old_date)
                i           = 1
                base        = archive_key
                while archive_key in backup:
                    archive_key = "{} ({})".format(base, i)
                    i += 1
                backup[archive_key] = {"date": old_date, "data": stored}
                backup[name]        = {"date": TODAY,    "data": current}
                updated = True

    return backup, updated

def get_all_filters():
    """Return all ParameterFilterElement objects in the project."""
    return list(FilteredElementCollector(doc).OfClass(ParameterFilterElement))

# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    filters = get_all_filters()

    if not filters:
        forms.alert("No View Filters found in this project.", title="Manage Filters")
        return

    filter_map  = {f.Name: f for f in sorted(filters, key=lambda x: x.Name)}
    backup_path = get_backup_path(doc)

    if backup_path:
        backup = load_backup(backup_path)
        backup, _ = sync_backup(backup, filters)
        save_backup(backup_path, backup, logger)

    selected_names = forms.SelectFromList.show(
        sorted(filter_map.keys()),
        title="Manage Filters - Select to Delete",
        button_name="Delete Selected",
        multiselect=True,
    )

    if not selected_names:
        return

    backup_note = (
        "\n\nBackup saved next to your Revit file." if backup_path
        else "\n\nWarning: model not saved, no backup created."
    )

    confirm = forms.alert(
        "Delete {} filter{}?{}\n\n{}".format(
            len(selected_names),
            "s" if len(selected_names) != 1 else "",
            backup_note,
            "\n".join(selected_names)
        ),
        yes=True,
        no=True,
        title="Confirm Delete"
    )

    if not confirm:
        return

    deleted = 0
    failed  = []

    with revit.Transaction("Delete Filters"):
        for name in selected_names:
            try:
                doc.Delete(filter_map[name].Id)
                deleted += 1
            except Exception as ex:
                failed.append("{} - {}".format(name, type(ex).__name__))

    if backup_path:
        backup = load_backup(backup_path)
        backup, _ = sync_backup(backup, get_all_filters())
        save_backup(backup_path, backup, logger)

    msg = "Deleted {} filter{}.".format(deleted, "s" if deleted != 1 else "")
    if backup_path:
        msg += "\n\nBackup: {}".format(backup_path)
    if failed:
        msg += "\n\nFailed:\n" + "\n".join(failed)

    forms.alert(msg, title="Manage Filters")

if __name__ == "__main__":
    main()
