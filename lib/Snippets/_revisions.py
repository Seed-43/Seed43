# -*- coding: utf-8 -*-
__title__  = "_revisions"
__author__  = "Seed43"
__doc__     = """
VERSION 260507
_____________________________________________________________________
Description:
Shared helper functions for working with Revit Revisions and the
filter backup system.

Used by Persistent Revisions, Revision Status Colour, Filter Delete
and Backup, and Filters Restore. Import the functions you need.

Example:
    from Snippets._revisions import get_revision_description
    from Snippets._revisions import safe_str, get_backup_path, load_backup, save_backup
_____________________________________________________________________
Functions:
get_revision_description(rev)
    Return the user-visible description of a revision.

safe_str(net_string)
    Convert a .NET string to a safe Python string.

get_backup_path(doc)
    Return the path for the filter backup JSON file next to the model.

load_backup(path)
    Load a backup JSON file and return it as a dict.

save_backup(path, data, logger)
    Write a dict to disk as a JSON backup file.
_____________________________________________________________________
Last update:
- Initial release, extracted from Persistent Revisions, Revision
  Status Colour, Filter Delete and Backup, and Filters Restore
_____________________________________________________________________
"""

import json
from Autodesk.Revit.DB import BuiltInParameter
from System.IO import File, Directory, Path
from System.Text import Encoding


# ── REVISION HELPERS ──────────────────────────────────────────────────────────

def get_revision_description(rev):
    """
    Return the user-visible description of a revision.

    Reads the PROJECT_REVISION_REVISION_DESCRIPTION parameter first.
    Falls back to the revision Name if the description is empty.
    Returns "?" if both fail.
    """
    try:
        param = rev.get_Parameter(
            BuiltInParameter.PROJECT_REVISION_REVISION_DESCRIPTION)
        if param and param.AsString():
            return param.AsString()
    except Exception:
        pass
    try:
        return rev.Name
    except Exception:
        return "?"


# ── BACKUP HELPERS ────────────────────────────────────────────────────────────

def safe_str(net_string):
    """
    Convert a .NET string to a safe Python string via UTF-8 bytes.

    Replaces any characters that cannot be represented in ASCII with
    a placeholder. Returns "unknown" if conversion fails entirely.
    """
    try:
        raw = Encoding.UTF8.GetBytes(net_string)
        return (
            Encoding.UTF8.GetString(raw)
                         .encode("utf-8", "replace")
                         .decode("ascii", "replace")
        )
    except Exception:
        return "unknown"


def get_backup_path(doc):
    """
    Return the full path for the filter backup JSON file.

    The file is placed next to the Revit model and named:
    [ModelName]_filters_backup.json

    Returns None if the model has not been saved yet.
    """
    try:
        model_path = doc.PathName
        if not model_path:
            return None
        folder     = Path.GetDirectoryName(model_path)
        model_name = Path.GetFileNameWithoutExtension(model_path)
        return Path.Combine(
            folder,
            "{}_filters_backup.json".format(safe_str(model_name))
        )
    except Exception:
        return None


def load_backup(path):
    """
    Load a filter backup JSON file and return it as a dict.

    Returns an empty dict if the file does not exist or cannot be read.
    """
    try:
        if not File.Exists(path):
            return {}
        raw = File.ReadAllText(path, Encoding.UTF8)
        return json.loads(raw)
    except Exception:
        return {}


def save_backup(path, data, logger=None):
    """
    Write a backup dict to disk as a JSON file.

    Creates the folder if it does not exist. Logs an error via logger
    if the write fails (logger is optional).
    """
    try:
        folder = Path.GetDirectoryName(path)
        if folder and not Directory.Exists(folder):
            Directory.CreateDirectory(folder)
        content = json.dumps(data, indent=2, ensure_ascii=True)
        File.WriteAllText(path, content, Encoding.UTF8)
    except Exception as ex:
        if logger:
            logger.error("Could not save backup: {}".format(type(ex).__name__))
