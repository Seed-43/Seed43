# -*- coding: utf-8 -*-
"""pyRecent history, settings and document tracking.

Shared by the pyRecent button (reads and edits the list) and startup.py
(records opens, saves and closes as they happen). Everything lives under
.user/pyRecent/ so the updater never touches it:

    .user/pyRecent/history.json     the list of files and their opens
    .user/pyRecent/settings.json    the Options tab
    .user/pyRecent/thumbnails/*.png one picture per file

Event handlers go through a small trampoline (see subscribe) because pyRevit
runs startup.py again on every reload, in a fresh engine, while the handlers
Revit already holds stay subscribed. The trampoline is subscribed once per
Revit session and always calls the newest copy of this module.
"""

import io
import os
import json
import hashlib
from datetime import datetime

from System import AppDomain

from Snippets import _userdata

__all__ = [
    "TOOL", "load_settings", "save_settings", "in_tracked_folders",
    "load_history", "save_history", "find_project", "record_open",
    "record_close", "set_reference", "set_thumbnail", "thumbnail_path",
    "remove_project", "remove_where", "set_skip", "skipping", "subscribe",
]

# ── CONSTANTS ───────────────────────────────────────────────────────────────

TOOL = "pyRecent"
MAX_OPENS = 200                 # open-history lines kept per file

DEFAULT_SETTINGS = {
    "filter_enabled": False,    # only track files inside `folders`
    "folders": [],
    "auto_reference": True,     # opened and closed without a save -> Reference
    "thumb_on_open": True,      # make a thumbnail the first time a file opens
}

# AppDomain slots are shared by every pyRevit engine in this Revit session,
# unlike module globals, which each engine gets its own copy of.
_SKIP_KEY = "Seed43.pyRecent.SkipRecording"
_IMPL_KEY = "Seed43.pyRecent.Impl"
_SUBSCRIBED_KEY = "Seed43.pyRecent.Subscribed"


# ── HELPERS (private to this file) ──────────────────────────────────────────

def _dir():
    return _userdata.user_dir(TOOL)


def thumb_dir():
    """Folder holding the thumbnails."""
    return _userdata.user_dir(TOOL, "thumbnails")


def _now():
    return datetime.now().strftime("%Y-%m-%dT%H:%M:%S")


def _read_json(path, default):
    try:
        with io.open(path, "r", encoding="utf-8") as f:
            return json.loads(f.read())
    except Exception:
        return default


def _write_json(path, data):
    text = json.dumps(data, indent=2)
    tmp = path + ".tmp"
    with io.open(tmp, "w", encoding="utf-8") as f:
        f.write(unicode(text))
    if os.path.exists(path):
        os.remove(path)
    os.rename(tmp, path)


def path_key(path):
    """Stable short id for a file path (thumbnail names)."""
    return hashlib.md5(path.lower().encode("utf-8")).hexdigest()


def _norm(path):
    return os.path.normpath(path or u"").replace(u"/", u"\\").lower().rstrip(u"\\")


# ── SETTINGS ────────────────────────────────────────────────────────────────

def load_settings():
    """The Options tab settings, with defaults for anything missing."""
    data = _read_json(os.path.join(_dir(), "settings.json"), {})
    out = dict(DEFAULT_SETTINGS)
    if isinstance(data, dict):
        out.update(data)
    return out


def save_settings(settings):
    _write_json(os.path.join(_dir(), "settings.json"), settings)


# --- Folder structures ---
# A tracked folder is a folder STRUCTURE: a full path in which any folder can
# be "*", meaning any one folder. Built from an example folder, e.g.
#   N:\(N) NAGEL CONSULTANTS LTD\#JOB\*\*\DRAWINGS\STRUCTURAL
# tracks DRAWINGS\STRUCTURAL in every job, and the folders under it.

ANY = u"*"


def split_folder(path):
    """Folder path -> its parts: drive (or \\\\server) first, then folders."""
    p = os.path.normpath(path or u"").replace(u"/", u"\\").rstrip(u"\\")
    if p.startswith(u"\\\\"):
        parts = p[2:].split(u"\\")
        parts[0] = u"\\\\" + parts[0]
        return parts
    return [x for x in p.split(u"\\") if x]


def suggest_any(parts):
    """Guess which parts change from job to job: folders with a number in
    them (job numbers, job ranges), never the drive, server or share."""
    keep = 2 if parts and parts[0].startswith(u"\\\\") else 1
    return [i >= keep and any(c.isdigit() for c in part)
            for i, part in enumerate(parts)]


def build_structure(parts, any_flags):
    """Join parts back into a structure string, '*' where any_flags is set."""
    out = [ANY if flag else part for part, flag in zip(parts, any_flags)]
    if out and out[0].startswith(u"\\\\"):
        return u"\\".join(out)
    return u"\\".join(out) if len(out) > 1 else out[0] + u"\\"


def matches_structure(folder, structure):
    """True if `folder` is the structure's folder, or inside it."""
    want = split_folder(structure)
    have = split_folder(folder)
    if not want or len(have) < len(want):
        return False
    for w, h in zip(want, have):
        if w != ANY and w.lower() != h.lower():
            return False
    return True


def in_tracked_folders(path, settings=None):
    """True if file `path` should be tracked under the folder filter."""
    settings = settings or load_settings()
    structures = [f for f in settings.get("folders") or [] if f and f.strip()]
    if not settings.get("filter_enabled") or not structures:
        return True
    folder = os.path.dirname(path)
    for structure in structures:
        structure = structure.strip()
        if structure[1:2] == u":" or structure.startswith(u"\\\\"):
            if matches_structure(folder, structure):
                return True
        # an older partial entry (DRAWINGS\STRUCTURAL): that run of folders anywhere
        elif (u"\\" + _norm(structure).strip(u"\\") + u"\\") in (u"\\" + _norm(folder) + u"\\"):
            return True
    return False


# ── HISTORY ─────────────────────────────────────────────────────────────────

def load_history():
    data = _read_json(os.path.join(_dir(), "history.json"), None)
    if not isinstance(data, dict) or "projects" not in data:
        data = {"projects": []}
    return data


def save_history(data):
    _write_json(os.path.join(_dir(), "history.json"), data)


def find_project(data, path):
    low = (path or u"").lower()
    for proj in data["projects"]:
        if proj.get("path", u"").lower() == low:
            return proj
    return None


def record_open(path, name=None, revit_version=None):
    """Record an open, newest first. Opens on the same day share one line."""
    if skipping() or not path:
        return None
    data = load_history()
    proj = find_project(data, path)
    ts = _now()
    if proj is None:
        proj = {
            "path": path,
            "name": name or os.path.splitext(os.path.basename(path))[0],
            "revit_version": revit_version or u"",
            "first_opened": ts,
            "last_opened": ts,
            "last_saved": u"",
            "is_reference": False,
            "thumbnail": u"",
            "opens": [{"timestamp": ts}],
        }
    else:
        data["projects"].remove(proj)
        proj["last_opened"] = ts
        if name:
            proj["name"] = name
        if revit_version:
            proj["revit_version"] = revit_version
        opens = proj.get("opens") or []
        if opens and opens[-1].get("timestamp", u"")[:10] == ts[:10]:
            opens[-1] = {"timestamp": ts}
        else:
            opens.append({"timestamp": ts})
        proj["opens"] = opens[-MAX_OPENS:]
    data["projects"].insert(0, proj)
    save_history(data)
    return proj


def record_close(path, saved=None, auto_reference=True):
    """Stamp the open's duration and, if asked, the Reference flag.

    saved is True / False for whether the file was saved while it was open,
    or None when that is unknown (it was open before tracking started).
    """
    if skipping() or not path:
        return
    data = load_history()
    proj = find_project(data, path)
    if not proj:
        return
    opens = proj.get("opens") or []
    if opens and opens[-1].get("duration_sec") is None:
        try:
            opened = datetime.strptime(opens[-1]["timestamp"], "%Y-%m-%dT%H:%M:%S")
            seconds = int((datetime.now() - opened).total_seconds())
            if 0 <= seconds <= 7 * 24 * 3600:
                opens[-1]["duration_sec"] = seconds
        except Exception:
            pass
    if saved:
        proj["last_saved"] = _now()
    if auto_reference and saved is not None:
        proj["is_reference"] = not saved
    save_history(data)


def _update(path, **fields):
    data = load_history()
    proj = find_project(data, path)
    if not proj:
        return False
    proj.update(fields)
    save_history(data)
    return True


def set_reference(path, is_reference):
    return _update(path, is_reference=bool(is_reference))


def set_thumbnail(path, filename):
    """Point a file at a new thumbnail, deleting the one it replaces."""
    data = load_history()
    proj = find_project(data, path)
    if not proj:
        return False
    old = proj.get("thumbnail")
    if old and old != filename:
        try:
            os.remove(os.path.join(thumb_dir(), old))
        except Exception:
            pass
    proj["thumbnail"] = filename
    save_history(data)
    return True


def thumbnail_path(proj):
    """Full path of a file's thumbnail, or None if it has none."""
    name = proj.get("thumbnail")
    if not name:
        return None
    full = os.path.join(thumb_dir(), name)
    return full if os.path.exists(full) else None


def remove_where(predicate):
    """Remove every file for which predicate(proj) is true. Returns count."""
    data = load_history()
    keep, gone = [], []
    for proj in data["projects"]:
        (gone if predicate(proj) else keep).append(proj)
    for proj in gone:
        if proj.get("thumbnail"):
            try:
                os.remove(os.path.join(thumb_dir(), proj["thumbnail"]))
            except Exception:
                pass
    if gone:
        data["projects"] = keep
        save_history(data)
    return len(gone)


def remove_project(path):
    low = path.lower()
    return remove_where(lambda p: p.get("path", u"").lower() == low) > 0


# ── SKIP FLAG ───────────────────────────────────────────────────────────────

def set_skip(value):
    """Stop recording while pyRecent opens a file itself (thumbnail regen)."""
    AppDomain.CurrentDomain.SetData(_SKIP_KEY, bool(value))


def skipping():
    try:
        return bool(AppDomain.CurrentDomain.GetData(_SKIP_KEY))
    except Exception:
        return False


# ── DOCUMENT EVENTS ─────────────────────────────────────────────────────────

# Per-session state, kept in this module because every event reaches it.
# _session: lower-case path -> {"saved": bool}
# _closing: document id -> path, between DocumentClosing and DocumentClosed
_session = {}
_closing = {}


def is_host_document(doc):
    """A saved project file, not a family, link or new unsaved model."""
    try:
        return bool(doc and not doc.IsFamilyDocument and not doc.IsLinked
                    and doc.PathName)
    except Exception:
        return False


def on_opened(sender, args):
    doc = args.Document
    if skipping() or not is_host_document(doc):
        return
    settings = load_settings()
    path = doc.PathName
    if not in_tracked_folders(path, settings):
        return
    _session[path.lower()] = {"saved": False}
    proj = record_open(path, doc.Title, doc.Application.VersionNumber)
    if proj and settings.get("thumb_on_open", True) and not thumbnail_path(proj):
        from Snippets import _recent_thumb
        _recent_thumb.generate(doc, path)


def on_saving(sender, args):
    """A save only counts as work if there was something to save."""
    doc = args.Document
    if not is_host_document(doc):
        return
    state = _session.get(doc.PathName.lower())
    if state is not None and doc.IsModified:
        state["saved"] = True


def on_saved_as(sender, args):
    doc = args.Document
    if skipping() or not is_host_document(doc):
        return
    if in_tracked_folders(doc.PathName):
        _session[doc.PathName.lower()] = {"saved": True}
        record_open(doc.PathName, doc.Title, doc.Application.VersionNumber)


def on_closing(sender, args):
    doc = args.Document
    if is_host_document(doc):
        _closing[args.DocumentId] = doc.PathName


def on_closed(sender, args):
    """Settle the Reference flag once the close has really happened.

    Left until DocumentClosed because the 'save changes?' answer is given
    after DocumentClosing fires, and that save has to count.
    """
    path = _closing.pop(args.DocumentId, None)
    if not path or skipping():
        return
    if str(args.Status) != "Succeeded":
        return          # close was cancelled, the file is still open
    settings = load_settings()
    if not in_tracked_folders(path, settings):
        return
    state = _session.pop(path.lower(), None)
    record_close(path, state["saved"] if state else None,
                 settings.get("auto_reference", True))


def _trampoline(name):
    def handler(sender, args):
        impl = AppDomain.CurrentDomain.GetData(_IMPL_KEY)
        if impl is None:
            return
        try:
            getattr(impl, name)(sender, args)
        except Exception:
            pass        # never let tracking get in the way of opening a file
    return handler


def subscribe(app):
    """Hook the document events once per Revit session.

    Each call (every startup.py run, so every pyRevit reload) makes this copy
    of the module the one the handlers call, so edits take effect on reload.
    """
    import sys
    AppDomain.CurrentDomain.SetData(_IMPL_KEY, sys.modules[__name__])
    if AppDomain.CurrentDomain.GetData(_SUBSCRIBED_KEY):
        return False
    app.DocumentOpened += _trampoline("on_opened")
    app.DocumentSaving += _trampoline("on_saving")
    app.DocumentSavedAs += _trampoline("on_saved_as")
    app.DocumentSynchronizingWithCentral += _trampoline("on_saving")
    app.DocumentClosing += _trampoline("on_closing")
    app.DocumentClosed += _trampoline("on_closed")
    AppDomain.CurrentDomain.SetData(_SUBSCRIBED_KEY, True)
    return True
