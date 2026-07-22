# -*- coding: utf-8 -*-
import json
from Autodesk.Revit.DB import BuiltInParameter
from System.IO import File, Directory, Path
from System.Text import Encoding


# ── REVISION HELPERS ──────────────────────────────────────────────────────────

def get_revision_description(rev):
    """
    Return the user-visible description of a revision.

    Tries the direct Description property first, since that's the standard
    Revit API property and doesn't need a parameter lookup. Falls back to
    the PROJECT_REVISION_REVISION_DESCRIPTION parameter, then the revision
    Name, in case Description is ever unavailable. Returns "?" if all fail.
    """
    try:
        if rev.Description:
            return rev.Description
    except Exception:
        pass
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
