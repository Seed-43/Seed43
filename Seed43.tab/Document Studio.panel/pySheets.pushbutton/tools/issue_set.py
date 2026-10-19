# -*- coding: utf-8 -*-
# issue_set.py
"""Issue set mode: archive superseded exports, then rebuild the combined set
from the files on disk.

Each export folder keeps a register (REGISTER_NAME, issue register.pysheets) mapping every exported
item, by Revit UniqueId, to the files it last produced, its revision and the
day it was printed. Filenames alone cannot do this: the revision is usually
in the name, so the Rev A and Rev B files of one sheet share nothing a
lookup could match on.

The rules, per format folder (<export>\\<FMT>\\):
- An item with a revision is judged by the Revit revision itself, not its
  label, and by where it sits in the project's revision sequence:
    same revision, same label      overwrite
    same revision, new label       archive (Draft 01 reused as Rev A)
    old revision deleted           archive
    old revision earlier in list   archive (Rev 1 after A, B, C)
    old revision later in list     blocked: printing an older revision
                                   than the one issued is refused, unless
                                   the user overrides it (then archive)
  Entries without revision ids (none yet) fall back to comparing labels.
- An item with no revision (3D views, schedules) is archived when the last
  export was on an earlier day. The same day overwrites.
- Superseded files go to <FMT>\\Archived\\. A same-named file already there
  is replaced; undated items get the print date added so days don't collide.
- The first run on a folder with no register archives everything already
  in it, so the register starts from a known state.

Set files (the combined PDF and the single XLS workbook) carry the date in
their name. A new name means a new day, and the previous set is archived.

Nothing here touches Revit, so the rules can be tested outside it.
"""
import io
import json
import os
import os.path as op
import re
import shutil
from datetime import datetime


# ── CONSTANTS ──────────────────────────────────────────────────────────────
REGISTER_NAME = 'issue register.pysheets'   # JSON inside
LEGACY_REGISTER_NAME = '_pySheets issue register.json'   # read once, then removed
ARCHIVE_DIR   = 'Archived'
DATE_FMT      = '%Y-%m-%d'   # same as pyrevit coreutils.current_date()

ORDER_NUMBER  = 'number'
ORDER_BROWSER = 'browser'
ORDER_LIST    = 'list'
ORDER_LABELS  = [(ORDER_NUMBER,  'Sheet number'),
                 (ORDER_BROWSER, 'Project Browser order'),
                 (ORDER_LIST,    'pySheets list order')]

DEFAULT_RULES = {'order': ORDER_NUMBER}

PDFSHARP_DLL = op.join('PdfSharp', 'PdfSharp.dll')   # under Seed43 lib/


# ── HELPERS ────────────────────────────────────────────────────────────────
def today():
    """Today's date in the register's format."""
    return datetime.now().strftime(DATE_FMT)


def dated_name(name, day=None):
    """`name` with the date on the end, unless it already carries it.

    The default combined naming format starts with {current_date}, so most
    names arrive dated already and pass through untouched."""
    day = day or today()
    if day in name:
        return name
    return u'{}_{}'.format(name, day)


def natural_key(text):
    """Sort key that puts S1.2 before S1.10."""
    parts = re.split(r'(\d+)', text or u'')
    return [int(p) if p.isdigit() else p.lower() for p in parts]


def archive_dir_for(path):
    """Create `path` if needed and return it."""
    if not op.isdir(path):
        os.makedirs(path)
    return path


def archive_dir(folder):
    """The Archived folder for one format folder, created on demand."""
    return archive_dir_for(op.join(folder, ARCHIVE_DIR))


def snapshot(folder):
    """Top-level files in `folder` as {name: (mtime, size)}."""
    found = {}
    if not op.isdir(folder):
        return found
    for name in os.listdir(folder):
        path = op.join(folder, name)
        if op.isfile(path):
            st = os.stat(path)
            found[name] = (st.st_mtime, st.st_size)
    return found


def changed_files(before, folder):
    """Names in `folder` that are new or rewritten since `before`."""
    now = snapshot(folder)
    return sorted(n for n, sig in now.items() if before.get(n) != sig), now


def is_locked(path):
    """True when another program holds `path` open (a PDF open in a viewer).

    Opening for update is refused while a viewer has the file, which is
    exactly the case that would otherwise fail halfway through a move."""
    if not op.isfile(path):
        return False
    try:
        with open(path, 'r+b'):
            pass
        return False
    except (IOError, OSError):
        return True


def _move(src, dst):
    """Move one file, replacing whatever is at `dst`."""
    if op.isfile(dst):
        os.remove(dst)
    shutil.move(src, dst)


# ── REGISTER ───────────────────────────────────────────────────────────────
class IssueRegister(object):
    """The JSON register kept in the export folder.

    Layout: {'version': 1,
             'items': {fmt: {uid: entry}},
             'sets':  {fmt: entry}}
    entry: {'files': [names], 'rev': str or None, 'rev_uid': str or None,
            'date': 'YYYY-MM-DD', 'number': str, 'name': str}
    A format key existing in 'items' (even empty) means that format folder
    has been taken over, so its first-run sweep has already happened."""

    # --- construction ---
    def __init__(self, base_folder):
        self.path   = op.join(base_folder, REGISTER_NAME)
        self.legacy = op.join(base_folder, LEGACY_REGISTER_NAME)
        self.data = {'version': 1, 'items': {}, 'sets': {}}
        source = (self.path if op.isfile(self.path) else
                  self.legacy if op.isfile(self.legacy) else None)
        if source:
            with io.open(source, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
            self.data['items'] = loaded.get('items') or {}
            self.data['sets']  = loaded.get('sets') or {}

    # --- public methods ---
    def knows_format(self, fmt):
        return fmt in self.data['items']

    def adopt_format(self, fmt):
        self.data['items'].setdefault(fmt, {})

    def item(self, fmt, uid):
        return self.data['items'].get(fmt, {}).get(uid)

    def items(self, fmt):
        return self.data['items'].get(fmt, {})

    def set_item(self, fmt, uid, entry):
        self.data['items'].setdefault(fmt, {})[uid] = entry

    def set_file(self, fmt):
        return self.data['sets'].get(fmt)

    def set_set_file(self, fmt, entry):
        self.data['sets'][fmt] = entry

    def save(self):
        """Write via a temp file so a failed write can't truncate the
        register the next run depends on."""
        tmp = self.path + '.tmp'
        # ensure_ascii keeps the output plain str, which sidesteps IronPython
        # 2's str/unicode split when a sheet name has a non-ASCII character.
        with open(tmp, 'w') as f:
            f.write(json.dumps(self.data, indent=2, sort_keys=True))
        if op.isfile(self.path):
            os.remove(self.path)
        os.rename(tmp, self.path)
        # Its contents now live in REGISTER_NAME, so the old name can go.
        if op.isfile(self.legacy):
            os.remove(self.legacy)


# ── RUN ────────────────────────────────────────────────────────────────────
class IssueItem(object):
    """What the rules need to know about one queued export, kept free of
    pySheets classes so this module stays testable outside Revit."""

    def __init__(self, uid, rev, number, name, token=None, rev_uid=None):
        self.uid     = uid
        self.rev     = rev or None     # None = item carries no revision
        self.rev_uid = rev_uid or None # UniqueId of the Revit Revision
        self.number  = number or u''
        self.name    = name or u''
        self.token   = token           # caller's own object (the QueueItem)
        self.files   = []              # filled in as the export runs
        self.moved   = []              # (archived path, original path)
        self.blocked = None            # register entry it would go back past
        self.override = False          # user chose to issue it anyway


class IssueRun(object):
    """One export run in issue set mode.

    Call order: plan() for every format, locked() to check nothing is held
    open, archive() to move the superseded files out, then during export
    attribute() as each item finishes, finish_format() after each format,
    finish_set() for set files, and save() at the end."""

    # --- construction ---
    def __init__(self, base_folder, rules=None, log=None, revisions=None):
        self.base   = base_folder
        # {Revision UniqueId: position in the project's revision sequence},
        # read live each run because revisions get reordered and deleted.
        self.revisions = revisions or {}
        self.rules  = dict(DEFAULT_RULES, **(rules or {}))
        self.day    = today()
        self.reg    = IssueRegister(base_folder)
        self.log    = log
        self.swept  = {}      # fmt -> number of files archived on first run
        self._plan  = {}      # fmt -> (folder, [IssueItem])
        self._sets  = {}      # fmt -> dict(folder, name, archive_folder, ...)
        self._snap  = {}      # fmt -> folder snapshot for attribution

    # --- public methods ---
    def plan(self, fmt, folder, items):
        """Register the items about to export into `folder` for `fmt`."""
        self._plan[fmt] = (folder, list(items))

    def plan_set(self, fmt, folder, filename, archive_folder):
        """Register a set file (combined PDF, single workbook) about to be
        written as `folder`\\`filename`. Its archive is `archive_folder`."""
        self._sets[fmt] = {'folder': folder, 'name': filename,
                           'archive': archive_folder, 'moved': []}

    def planned(self, fmt):
        """The IssueItems planned for `fmt`."""
        return self._plan.get(fmt, (None, []))[1]

    def set_target(self, fmt):
        """Full path the set file for `fmt` will be written to, or None."""
        s = self._sets.get(fmt)
        return op.join(s['folder'], s['name']) if s else None

    def locked(self):
        """Files this run must move or overwrite that are held open."""
        found = []
        for fmt, (folder, items) in self._plan.items():
            if not self.reg.knows_format(fmt):
                found.extend(op.join(folder, n) for n in snapshot(folder))
                continue
            for it in items:
                prev = self.reg.item(fmt, it.uid)
                for n in (prev or {}).get('files', []):
                    found.append(op.join(folder, n))
        for fmt, s in self._sets.items():
            found.append(op.join(s['folder'], s['name']))
            prev = self.reg.set_file(fmt)
            if prev:
                found.extend(op.join(prev.get('folder', s['folder']), n)
                             for n in prev.get('files', []))
        return sorted(set(p for p in found if is_locked(p)))

    def archive(self):
        """Sweep untracked folders, then move superseded files out."""
        for fmt, (folder, items) in self._plan.items():
            if not op.isdir(folder):
                os.makedirs(folder)
            if not self.reg.knows_format(fmt):
                self.swept[fmt] = self._sweep(folder)
                self.reg.adopt_format(fmt)
            for it in items:
                prev = self.reg.item(fmt, it.uid)
                verdict = self._compare(prev, it)
                if verdict == 'older' and not it.override:
                    it.blocked = prev
                elif verdict in ('newer', 'older'):
                    it.moved = self._archive_entry(folder, prev)
            self._snap[fmt] = snapshot(folder)
        for fmt, s in self._sets.items():
            prev = self.reg.set_file(fmt)
            if not prev:
                continue
            prev_folder = prev.get('folder', s['folder'])
            for n in prev.get('files', []):
                src = op.join(prev_folder, n)
                same = (op.normcase(src) ==
                        op.normcase(op.join(s['folder'], s['name'])))
                if op.isfile(src) and not same:
                    dst = op.join(archive_dir_for(s['archive']), n)
                    _move(src, dst)
                    s['moved'].append((dst, src))

    def attribute(self, fmt, item, done):
        """Credit the files written since the last call to `item`.

        Exporters name their output in ways the caller can't predict (DWG
        xrefs, image exports append the view name), so the folder itself is
        the record of what each item produced."""
        if fmt not in self._snap:
            return
        folder = self._plan[fmt][0]
        names, now = changed_files(self._snap[fmt], folder)
        self._snap[fmt] = now
        if done:
            item.files = names

    def finish_format(self, fmt, done_items, register_items=True):
        """Settle the register for `fmt` once its exports have run.

        Items that failed get their archived files put back, so a failed
        Rev B never leaves the sheet missing from the current set."""
        folder, items = self._plan.get(fmt, (None, []))
        done_ids = set(id(i) for i in done_items)
        for it in items:
            if id(it) not in done_ids:
                self._restore(it.moved)
                continue
            if not register_items:
                continue
            prev = self.reg.item(fmt, it.uid)
            # A same-revision export under a new filename (the sheet was
            # renamed, or the naming format changed) leaves the old file
            # behind. Archive it rather than leave two copies current.
            for n in (prev or {}).get('files', []):
                if n not in it.files and op.isfile(op.join(folder, n)):
                    self._archive_entry(folder, dict(prev, files=[n]))
            self.reg.set_item(fmt, it.uid, {
                'files': it.files, 'rev': it.rev, 'rev_uid': it.rev_uid,
                'date': self.day, 'number': it.number, 'name': it.name})

    def finish_set(self, fmt, ok):
        """Record the set file, or put the previous one back on failure."""
        s = self._sets.get(fmt)
        if not s:
            return
        if not ok:
            self._restore(s['moved'])
            return
        self.reg.set_set_file(fmt, {'folder': s['folder'],
                                    'files': [s['name']], 'date': self.day})

    def current_files(self, fmt, order_map=None):
        """Paths of `fmt`'s current files in combined-set order.

        Built from the register but filtered by what is on disk, so a file
        deleted by hand drops out of the next set."""
        folder = self._plan.get(fmt, (op.join(self.base, fmt), []))[0]
        mode = self.rules.get('order', ORDER_NUMBER)
        order_map = order_map or {}

        def key(pair):
            uid, entry = pair
            pos = order_map.get(uid) if mode != ORDER_NUMBER else None
            return (pos is None, pos, natural_key(entry.get('number')),
                    natural_key(entry.get('name')))

        paths = []
        for uid, entry in sorted(self.reg.items(fmt).items(), key=key):
            for n in sorted(entry.get('files', []), key=natural_key):
                p = op.join(folder, n)
                if op.isfile(p):
                    paths.append(p)
        return paths

    def save(self):
        self.reg.save()

    # --- private helpers ---
    def _compare(self, prev, it):
        """How `it` stands against what was last issued: 'same' (overwrite),
        'newer' (archive the old file) or 'older' (refuse the export)."""
        if not prev:
            return 'same'
        if not it.rev:
            return 'same' if prev.get('date') == self.day else 'newer'
        prev_uid = prev.get('rev_uid')
        if not prev.get('rev'):
            return 'newer'                 # first revision on this sheet
        if not prev_uid or not it.rev_uid:
            # Registered before revision ids were kept: labels are all
            # there is to go on.
            return 'same' if prev.get('rev') == it.rev else 'newer'
        if prev_uid == it.rev_uid:
            # One revision, relabelled: Draft 01 reused as Rev A.
            return 'same' if prev.get('rev') == it.rev else 'newer'
        prev_pos = self.revisions.get(prev_uid)
        cur_pos  = self.revisions.get(it.rev_uid)
        if prev_pos is None or cur_pos is None:
            return 'newer'                 # old revision deleted from Revit
        return 'newer' if prev_pos < cur_pos else 'older'

    def find_older(self):
        """(fmt, IssueItem, issued entry) for every item that is an earlier
        revision than the one already issued. Moves nothing, so the caller
        can ask before archive() runs and set .override on the ones to
        issue anyway."""
        found = []
        for fmt, (_, items) in sorted(self._plan.items()):
            for it in items:
                prev = self.reg.item(fmt, it.uid)
                if self._compare(prev, it) == 'older':
                    found.append((fmt, it, prev))
        return found

    def blocked(self):
        """(fmt, IssueItem) for every item refused as an older revision."""
        return [(fmt, it) for fmt, (_, items) in sorted(self._plan.items())
                for it in items if it.blocked]

    def _archive_entry(self, folder, prev):
        """Move one register entry's files into Archived. Returns the moves
        made, as (archived path, original path), so they can be undone."""
        moved = []
        dated = not prev.get('rev')
        for n in prev.get('files', []):
            src = op.join(folder, n)
            if not op.isfile(src):
                continue
            if dated:
                stem, ext = op.splitext(n)
                n = u'{}_{}{}'.format(stem, prev.get('date') or self.day, ext)
            dst = op.join(archive_dir(folder), n)
            _move(src, dst)
            moved.append((dst, src))
        return moved

    def _sweep(self, folder):
        """First run on a folder: archive whatever is already there."""
        count = 0
        for n in sorted(snapshot(folder)):
            _move(op.join(folder, n), op.join(archive_dir(folder), n))
            count += 1
        return count

    def _restore(self, moves):
        for dst, src in reversed(moves or []):
            try:
                if op.isfile(dst) and not op.exists(src):
                    shutil.move(dst, src)
            except Exception as ex:
                if self.log:
                    self.log.warning('Could not restore %s: %s', src, ex)


# ── PDF MERGE ──────────────────────────────────────────────────────────────
_pdfsharp = None


def _load_pdfsharp(lib_dir):
    """Load PDFsharp 1.50 once per session.

    1.50 rather than 6.x: 6.x needs Microsoft.Extensions.Logging.Abstractions
    8.0, and Revit already has 2.2 (2024) or 6.0 (2025) loaded, so it can
    never bind. 1.50 has no dependencies at all."""
    global _pdfsharp
    if _pdfsharp is not None:
        return _pdfsharp
    import clr
    clr.AddReferenceToFileAndPath(op.join(lib_dir, PDFSHARP_DLL))
    # HACK: .NET 8 (Revit 2025+) ships without the legacy code pages and
    # PDFsharp 1.50 writes in Windows-1252, so Save() throws "No data is
    # available for encoding 1252" until they are registered. The provider
    # does not exist on .NET Framework (Revit 2024), which has them built in.
    try:
        from System.Text import Encoding, CodePagesEncodingProvider
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
    except Exception:
        pass
    from PdfSharp.Pdf import PdfDocument
    from PdfSharp.Pdf.IO import PdfReader, PdfDocumentOpenMode
    _pdfsharp = (PdfDocument, PdfReader, PdfDocumentOpenMode)
    return _pdfsharp


def merge_pdfs(paths, dest, lib_dir):
    """Merge `paths` in order into `dest`. Returns the page count.

    Written to a temp file first and swapped in, so a failure part-way
    leaves the previous set intact rather than a truncated one."""
    PdfDocument, PdfReader, PdfDocumentOpenMode = _load_pdfsharp(lib_dir)
    out = PdfDocument()
    pages = 0
    try:
        for p in paths:
            src = PdfReader.Open(p, PdfDocumentOpenMode.Import)
            try:
                for i in range(src.PageCount):
                    out.AddPage(src.Pages[i])
                    pages += 1
            finally:
                src.Dispose()
        tmp = dest + '.tmp'
        out.Save(tmp)
    finally:
        out.Dispose()
    if op.isfile(dest):
        os.remove(dest)
    os.rename(tmp, dest)
    return pages
