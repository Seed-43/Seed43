# -*- coding: utf-8 -*-
# "Shared Params"
# "Seed43"
# """
# Edit a Revit shared parameter file as a table.
#
# WHY IT IS NOT JUST A TEXT EDITOR
#     A shared parameter file is the only place a GUID is written down. A
#     family that bound to a parameter finds it again by GUID and by nothing
#     else, so a GUID that gets duplicated, retyped or dropped silently
#     unbinds every family that ever used it, and the damage shows up much
#     later as a schedule column that will not fill in.
#
#     So the table is the easy half. The work is in refusing to write a file
#     that would do that: duplicate GUIDs, malformed GUIDs, empty names,
#     group ids with no group, and tabs or line breaks pasted into a field,
#     which would shift every column after them by one.
#
# THE SAVE CANNOT LEAVE A HALF WRITTEN FILE
#     The old file is copied into a _backups folder beside it, the new one
#     is written to a temp file, and only then is the temp moved into place.
#     A crash part way through therefore leaves either the old file whole or
#     the new one whole, never a truncated one. This is the specific failure
#     this tool exists to avoid.
#
# NOTHING IS FLATTENED ON THE WAY THROUGH
#     Values are carried as raw strings from load to save. The column order
#     is read from the '*PARAM' header rather than assumed, so a file
#     written by an older Revit, with no USERMODIFIABLE and no
#     HIDEWHENNOVALUE column, parses into the right fields instead of
#     sliding the description into the visibility slot.
#
#     A field this tool does not recognise survives untouched unless that
#     cell is edited. Fred's own file carries VISIBLE=5 on two rows, which
#     is neither 0 nor 1; a tool that coerced it to a bool would rewrite
#     those rows without being asked.
#
# THE MODEL IS NEVER TOUCHED
#     This edits a text file on disk. It reads which file Revit is pointed
#     at, and can point Revit at another one, and that is the whole of its
#     contact with Revit. It starts no transaction, so there is nothing to
#     undo in the model. Undo here steps back through edits to the table.
#
# Target: Revit 2022-2026, IronPython 2.
# """

# ── IMPORTS ─────────────────────────────────────────────────────────────────

import copy
import os
import sys

from pyrevit import HOST_APP, forms

from System import Action
from System.Collections.Generic import List
from System.ComponentModel import (GroupDescription, INotifyPropertyChanged,
                                   PropertyChangedEventArgs)
from System.Windows.Controls import ComboBox, GroupStyle, TextBox
from System.Windows.Data import ListCollectionView
from System.Windows.Input import Key, Keyboard, ModifierKeys
from System.Windows.Threading import DispatcherPriority

from Snippets._dialogs import choice, confirm, message, save_file_as
from Snippets._userdata import user_path
from Snippets.seed43_theme import (apply_seed43_dimensions,
                                   apply_seed43_palette)

# The parse and write core is a sibling module rather than a lib/Snippets
# entry: nothing else consumes it, and a single consumer does not belong in
# the shared folder. pyRevit puts the bundle on sys.path already; the append
# is here so the module also imports when the file is run for its tests.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.append(SCRIPT_DIR)

import sp_file  # noqa: E402


# ── CONSTANTS ───────────────────────────────────────────────────────────────

TOOL_NAME = "Shared Params"
GUIDE_FILE = user_path(TOOL_NAME, "guide_shown.json")
LAST_FILE = user_path(TOOL_NAME, "last_file.txt")

UNDO_DEPTH = 40
ALL_GROUPS = u"All groups"

# Column positions in SharedParams.xaml. Resolving the columns by index
# rather than by x:Name keeps this working whatever the name scope does with
# DataGridColumn, which is not part of the visual tree.
COL_NAME = 0
COL_TYPE = 1
COL_GROUP = 2
COL_DESC = 3
COL_USERMOD = 4
COL_VISIBLE = 5
COL_GUID = 6
COL_HIDE = 7

YES = u"Yes"
NO = u"No"

# What an empty cell means, per field. An old file has no USERMODIFIABLE
# column at all, and Revit treats a parameter with no entry as modifiable,
# so showing "No" there would misreport the file.
BLANK_MEANS = {
    "VISIBLE": True,
    "USERMODIFIABLE": True,
    "HIDEWHENNOVALUE": False,
}

GUIDE = (
    u"Shared Params edits a Revit shared parameter file as a table.\n\n"
    u"Open a file, or use Revit file to load the one Revit is currently "
    u"pointed at. Edit any cell in place. Select rows to add, duplicate, "
    u"delete, rename or move them between groups.\n\n"
    u"Save checks the file before writing it. Duplicate or malformed GUIDs, "
    u"empty names and groups that do not exist stop the save, because Revit "
    u"finds a parameter by its GUID and nothing else: reuse one and every "
    u"family bound to it points at the wrong parameter.\n\n"
    u"New GUID is the one button to be careful with. A parameter that "
    u"already exists in a family or a project is matched by GUID, so giving "
    u"it a new one breaks that link. It is for fixing a duplicate, not for "
    u"tidying.\n\n"
    u"The old file is copied into a _backups folder beside it on every "
    u"save, and the new file is written to one side and moved into place, "
    u"so an interrupted save cannot leave you with half a file.\n\n"
    u"Clicking a column header sorts what is on screen. Sort reorders the "
    u"rows in the file itself."
)


# ── DATA TYPE PICKER ────────────────────────────────────────────────────────
# The file stores tokens (LINEAR_FORCE, HVAC_DUCTSIZE). The picker shows them
# the way Revit's own dialog does: readable names under discipline headings.

# Group prefixes to drop from the label, since the heading already says it.
_TYPE_PREFIXES = ("HVAC_", "ELECTRICAL_", "PIPING_")

# Tokens the word rule below cannot split or case correctly.
_TYPE_LABELS = {
    "YESNO": u"Yes/No",
    "MULTILINETEXT": u"Multiline Text",
    "FAMILYTYPE": u"Family Type",
    "URL": u"URL",
    "HVAC_DUCTSIZE": u"Duct Size",
    "HVAC_CROSSSECTION": u"Cross Section",
    "HVAC_HEATGAIN": u"Heat Gain",
    "ELECTRICAL_CABLETRAYSIZE": u"Cable Tray Size",
    "ELECTRICAL_CONDUITSIZE": u"Conduit Size",
    "CROSSSECTIONALAREA": u"Cross Sectional Area",
    "MASSPERUNITAREA": u"Mass per Unit Area",
    "NUMBEROFPOLES": u"Number of Poles",
}

FILE_ONLY_GROUP = u"In this file"


def type_label(token):
    """LINEAR_FORCE -> Linear Force. Unknown tokens come out readable too."""
    if token in _TYPE_LABELS:
        return _TYPE_LABELS[token]
    text = token
    for prefix in _TYPE_PREFIXES:
        if text.startswith(prefix):
            text = text[len(prefix):]
            break
    words = [w.capitalize() for w in text.split(u"_") if w]
    words = [w.lower() if w in (u"Per", u"Of") and i else w
             for i, w in enumerate(words)]
    return u" ".join(words) or token


class TypeOption(object):
    """One entry in the data type dropdown. Token is what gets written."""

    def __init__(self, token, group):
        self._token = token
        self._group = group
        self._label = type_label(token)

    @property
    def Token(self):
        return self._token

    @property
    def Label(self):
        return self._label

    @property
    def Group(self):
        return self._group

    def __str__(self):
        return self._label


class _ByTypeGroup(GroupDescription):
    """Groups TypeOptions under their discipline heading.

    A GroupDescription subclass rather than PropertyGroupDescription("Group"):
    it reads the attribute directly, so it does not depend on WPF reflecting
    over an IronPython object.
    """

    def GroupNameFromItem(self, item, level, culture):
        return getattr(item, "Group", FILE_ONLY_GROUP)


def build_type_options(params):
    """Curated types by discipline, then any token only the file uses."""
    options = []
    known = set()
    for group, tokens in sp_file.DATA_TYPE_GROUPS:
        for token in sorted(tokens, key=type_label):
            options.append(TypeOption(token, group))
            known.add(token)
    extra = set()
    for p in params:
        t = p.get("DATATYPE", u"").strip()
        if t and t not in known:
            extra.add(t)
    for token in sorted(extra, key=type_label):
        options.append(TypeOption(token, FILE_ONLY_GROUP))
    return options


# ── REMEMBERING THE LAST FILE ───────────────────────────────────────────────

def remember_file(path):
    """Note which file was last edited, so the tool reopens on it.

    Written as UTF-8 rather than handed straight to a text mode write: a
    path with a non-ASCII character in it is unicode by the time it comes
    back from the file dialog, and IronPython 2 would try to encode that as
    ASCII and throw.
    """
    try:
        folder = os.path.dirname(LAST_FILE)
        if not os.path.isdir(folder):
            os.makedirs(folder)
        with open(LAST_FILE, "wb") as fh:
            fh.write(path.encode("utf-8"))
    except Exception:
        pass


def recall_file():
    try:
        if os.path.isfile(LAST_FILE):
            with open(LAST_FILE, "rb") as fh:
                path = fh.read().decode("utf-8").strip()
            if path and os.path.isfile(path):
                return path
    except Exception:
        pass
    return None


# ── ROW MODEL ───────────────────────────────────────────────────────────────

class ParamRow(INotifyPropertyChanged):
    """One grid row, backed by the dict that will be written to the file.

    Every property reads and writes ``self.data`` in place, so the file
    object is always current and there is no separate step where the grid
    is copied back into it and could be missed.

    INotifyPropertyChanged is implemented by hand because IronPython cannot
    declare a .NET event; the add_/remove_ pair is what WPF actually binds
    against. Same shape as pyTransmit's RecipientRecord.
    """

    def __init__(self, data, owner):
        self.data = data
        self.owner = owner          # the window, for the group name lookup
        self._problem = u""
        self._state = u""
        self._handlers = []

    # --- the .NET event ---

    def add_PropertyChanged(self, handler):
        self._handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._handlers:
            self._handlers.remove(handler)

    def _notify(self, name):
        args = PropertyChangedEventArgs(name)
        for handler in self._handlers:
            handler(self, args)

    # --- helpers ---

    @staticmethod
    def _clean(value):
        """A tab or a line break in a field would shift every column after
        it, so they are turned into spaces on the way in rather than
        rejected later."""
        if value is None:
            return u""
        return (unicode(value).replace(u"\t", u" ")
                .replace(u"\r", u" ").replace(u"\n", u" "))

    def _get(self, key):
        return self.data.get(key, u"")

    def _set(self, key, value, prop):
        value = self._clean(value)
        if self.data.get(key, u"") != value:
            self.data[key] = value
            self.owner.mark_dirty()
            self._notify(prop)

    def _flag(self, key):
        raw = self._get(key).strip()
        if not raw:
            return YES if BLANK_MEANS.get(key, False) else NO
        return NO if raw == u"0" else YES

    def _set_flag(self, key, value, prop):
        # Only write when the displayed state actually moves. A row holding
        # VISIBLE=5 shows as Yes; picking Yes again must leave the 5 alone.
        if value == self._flag(key):
            return
        self.data[key] = u"1" if value == YES else u"0"
        self.owner.mark_dirty()
        self._notify(prop)

    # --- bound properties ---

    @property
    def Name(self):
        return self._get("NAME")

    @Name.setter
    def Name(self, value):
        self._set("NAME", value, "Name")

    @property
    def DataType(self):
        return self._get("DATATYPE")

    @DataType.setter
    def DataType(self, value):
        self._set("DATATYPE", value, "DataType")

    @property
    def GroupName(self):
        return self.owner.group_name_for(self._get("GROUP"))

    @GroupName.setter
    def GroupName(self, value):
        group_id = self.owner.group_id_for(value)
        if group_id and self.data.get("GROUP", u"") != group_id:
            self.data["GROUP"] = group_id
            self.owner.mark_dirty()
            self._notify("GroupName")

    @property
    def Description(self):
        return self._get("DESCRIPTION")

    @Description.setter
    def Description(self, value):
        self._set("DESCRIPTION", value, "Description")

    @property
    def Guid(self):
        return self._get("GUID")

    @Guid.setter
    def Guid(self, value):
        self._set("GUID", (value or u"").strip(), "Guid")

    @property
    def DataCategory(self):
        return self._get("DATACATEGORY")

    @DataCategory.setter
    def DataCategory(self, value):
        self._set("DATACATEGORY", value, "DataCategory")

    @property
    def Visible(self):
        return self._flag("VISIBLE")

    @Visible.setter
    def Visible(self, value):
        self._set_flag("VISIBLE", value, "Visible")

    @property
    def UserModifiable(self):
        return self._flag("USERMODIFIABLE")

    @UserModifiable.setter
    def UserModifiable(self, value):
        self._set_flag("USERMODIFIABLE", value, "UserModifiable")

    @property
    def HideWhenNoValue(self):
        return self._flag("HIDEWHENNOVALUE")

    @HideWhenNoValue.setter
    def HideWhenNoValue(self, value):
        self._set_flag("HIDEWHENNOVALUE", value, "HideWhenNoValue")

    # --- read only, drives the cell colour and the row tooltip ---

    @property
    def RowState(self):
        return self._state

    @property
    def Problem(self):
        return self._problem or None

    def flag_problem(self, text):
        self._problem = text
        self._state = u"error" if text else u""
        self._notify("RowState")
        self._notify("Problem")


# ── GROUPS DIALOG ───────────────────────────────────────────────────────────

class GroupsWindow(forms.WPFWindow):
    """Add, rename and remove the parameter groups of the open file."""

    def __init__(self, owner, spf):
        forms.WPFWindow.__init__(self, "Groups.xaml")
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)
        self.Owner = owner
        self.spf = spf
        self.changed = False
        self.fill()

    def fill(self, select_id=None):
        items = List[object]()
        for g in self.spf.groups:
            used = self.spf.group_usage(g.get("ID", u""))
            items.Add(u"{0}  ({1})   {2}".format(
                g.get("NAME", u""), used,
                u"" if used else u"unused"))
        self.group_list.ItemsSource = items
        self.hint_tb.Text = (
            u"{0} groups. The number in brackets is how many parameters "
            u"sit in each.".format(len(self.spf.groups)))
        if select_id is not None:
            for i, g in enumerate(self.spf.groups):
                if g.get("ID") == select_id:
                    self.group_list.SelectedIndex = i
                    break

    def selected(self):
        i = self.group_list.SelectedIndex
        if i < 0 or i >= len(self.spf.groups):
            return None
        return self.spf.groups[i]

    def selection_changed(self, sender, args):
        pass

    def add_clicked(self, sender, args):
        name = (self.new_tb.Text or u"").strip()
        if not name:
            return
        if name in [g.get("NAME", u"") for g in self.spf.groups]:
            message(u"There is already a group called {0}.".format(name),
                    title=u"Groups")
            return
        row = self.spf.add_group(name)
        self.new_tb.Text = u""
        self.changed = True
        self.fill(select_id=row["ID"])

    def rename_clicked(self, sender, args):
        g = self.selected()
        if not g:
            message(u"Pick a group to rename.", title=u"Groups")
            return
        from Snippets._dialogs import ask_string
        new = ask_string(u"New name for {0}".format(g.get("NAME", u"")),
                         title=u"Rename group", default=g.get("NAME", u""))
        if not new or new == g.get("NAME", u""):
            return
        g["NAME"] = new.replace(u"\t", u" ").strip()
        self.changed = True
        self.fill(select_id=g["ID"])

    def del_clicked(self, sender, args):
        g = self.selected()
        if not g:
            message(u"Pick a group to delete.", title=u"Groups")
            return
        gid = g.get("ID", u"")
        used = self.spf.group_usage(gid)
        if used:
            # Deleting a group out from under its parameters would leave
            # every one of them pointing at an id that is no longer there,
            # which is a file Revit will not read.
            message(
                u"{0} parameters are in {1}. Move them to another group "
                u"first, with Set group on the main window.".format(
                    used, g.get("NAME", u"")),
                title=u"Groups")
            return
        if len(self.spf.groups) == 1:
            message(u"A shared parameter file needs at least one group.",
                    title=u"Groups")
            return
        if not confirm(u"Delete the group {0}?".format(g.get("NAME", u"")),
                       title=u"Groups", yes=u"Delete"):
            return
        self.spf.groups.remove(g)
        self.changed = True
        self.fill()

    def ok_clicked(self, sender, args):
        self.Close()


# ── RENAME DIALOG ───────────────────────────────────────────────────────────

class RenameWindow(forms.WPFWindow):
    """Prefix, suffix and find/replace across the selected parameter names."""

    def __init__(self, owner, names):
        forms.WPFWindow.__init__(self, "Rename.xaml")
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)
        self.Owner = owner
        self.names = names
        self.result = None
        self.title_tb.Text = u"Rename {0} parameter{1}".format(
            len(names), u"" if len(names) == 1 else u"s")
        self.refresh_preview()

    def apply_to(self, name):
        find = self.find_tb.Text or u""
        with_ = self.with_tb.Text or u""
        out = name
        if find:
            out = out.replace(find, with_)
        return (self.prefix_tb.Text or u"") + out + (self.suffix_tb.Text or u"")

    def refresh_preview(self):
        lines = []
        for name in self.names[:6]:
            new = self.apply_to(name)
            lines.append(u"{0}\n   -> {1}".format(name, new))
        if len(self.names) > 6:
            lines.append(u"... and {0} more".format(len(self.names) - 6))
        self.preview_tb.Text = u"\n".join(lines)

        blank = [n for n in self.names if not self.apply_to(n).strip()]
        self.warn_tb.Text = (
            u"{0} would end up with an empty name.".format(len(blank))
            if blank else u"")
        self.ok_btn.IsEnabled = not blank

    def preview_changed(self, sender, args):
        self.refresh_preview()

    def ok_clicked(self, sender, args):
        self.result = (self.prefix_tb.Text or u"",
                       self.suffix_tb.Text or u"",
                       self.find_tb.Text or u"",
                       self.with_tb.Text or u"")
        self.Close()

    def cancel_clicked(self, sender, args):
        self.result = None
        self.Close()


# ── MAIN WINDOW ─────────────────────────────────────────────────────────────

class SharedParamsWindow(forms.WPFWindow):
    """The parameter table."""

    # --- construction ---

    def __init__(self, xaml_name, start_path=None):
        forms.WPFWindow.__init__(self, xaml_name)
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)

        self.spf = sp_file.SharedParameterFile()
        self.dirty = False
        self.undo_stack = []
        self._loading = False
        self._group_names = {}
        self._group_ids = {}

        # Yes/No never changes, so it is filled once.
        for index in (COL_VISIBLE, COL_USERMOD, COL_HIDE):
            self.grid.Columns[index].ItemsSource = List[str]([YES, NO])

        if start_path and os.path.isfile(start_path):
            self.load_file(start_path)
        else:
            self.new_file(confirm_first=False)

    # --- the file ---

    def load_file(self, path):
        try:
            spf = sp_file.SharedParameterFile.load(path)
        except sp_file.ParseError as ex:
            message(unicode(ex), title=u"Shared Params")
            return False
        except Exception as ex:
            message(u"Could not read that file.\n\n{0}".format(ex),
                    title=u"Shared Params")
            return False
        self.spf = spf
        self.dirty = False
        self.undo_stack = []
        self.after_file_change()
        self.set_status(
            u"{0} parameters in {1} groups.".format(
                len(spf.params), len(spf.groups)))
        remember_file(path)
        return True

    def new_file(self, confirm_first=True):
        if confirm_first and not self.check_discard():
            return
        spf = sp_file.SharedParameterFile()
        spf.add_group(u"Default Group")
        self.spf = spf
        self.dirty = False
        self.undo_stack = []
        self.after_file_change()
        self.set_status(u"New file. Add a parameter to get going.")

    def after_file_change(self):
        """Rebuild everything that depends on which file is open."""
        self._group_names = self.spf.group_names()
        self._group_ids = {}
        for gid, name in self._group_names.items():
            self._group_ids[name] = gid

        # The type list is the curated one plus whatever the file already
        # uses, so a token this tool has never heard of still shows in its
        # own rows and can be picked for others. Grouped by discipline; the
        # column's SelectedValuePath maps each option back to its token.
        self._type_options = build_type_options(self.spf.params)
        view = ListCollectionView(List[object](self._type_options))
        view.GroupDescriptions.Add(_ByTypeGroup())
        self.grid.Columns[COL_TYPE].ItemsSource = view

        names = sorted(self._group_ids.keys())
        self.grid.Columns[COL_GROUP].ItemsSource = List[str](names)

        self._loading = True
        try:
            items = List[object]()
            items.Add(ALL_GROUPS)
            for name in names:
                items.Add(name)
            self.filter_group_cb.ItemsSource = items
            self.filter_group_cb.SelectedIndex = 0
        finally:
            self._loading = False

        self.refresh_rows()
        self.refresh_title()

    def refresh_title(self):
        path = self.spf.path
        star = u"  *" if self.dirty else u""
        self.subtitle_tb.Text = (path or u"unsaved file") + star

    def mark_dirty(self):
        if not self.dirty:
            self.dirty = True
            self.refresh_title()

    # --- group lookup, used by every row ---

    def group_name_for(self, group_id):
        name = self._group_names.get(group_id)
        if name is None:
            # An id with no group is a real fault in the file rather than
            # something to hide, so it is shown as what it is.
            return u"<missing {0}>".format(group_id)
        return name

    def group_id_for(self, name):
        return self._group_ids.get(name)

    # --- the table ---

    def visible_params(self):
        """The params passing the search box and the group filter."""
        needle = (self.search_tb.Text or u"").strip().lower()
        chosen = self.filter_group_cb.SelectedItem
        group_id = None
        if chosen and chosen != ALL_GROUPS:
            group_id = self._group_ids.get(chosen)

        out = []
        for p in self.spf.params:
            if group_id is not None and p.get("GROUP", u"") != group_id:
                continue
            if needle:
                blob = u" ".join([
                    p.get("NAME", u""), p.get("DESCRIPTION", u""),
                    p.get("DATATYPE", u""), p.get("GUID", u"")]).lower()
                if needle not in blob:
                    continue
            out.append(p)
        return out

    def refresh_rows(self, keep_selection=True):
        """Rebuild the grid rows from the file.

        Assigning one fresh list is deliberate: adding into a live
        ObservableCollection raises a change notification per row, which on
        a few thousand parameters is what turns a filter keystroke into a
        visible stall.
        """
        selected_guids = set()
        if keep_selection:
            for row in self.grid.SelectedItems:
                selected_guids.add(row.data.get("GUID", u""))

        rows = List[object]()
        first = None
        for p in self.visible_params():
            row = ParamRow(p, self)
            rows.Add(row)
            if first is None and p.get("GUID", u"") in selected_guids:
                first = row
        self.grid.ItemsSource = rows

        if keep_selection and selected_guids:
            self.grid.SelectedItems.Clear()
            for row in rows:
                if row.data.get("GUID", u"") in selected_guids:
                    self.grid.SelectedItems.Add(row)
            if first is not None:
                self.grid.ScrollIntoView(first)

        self.refresh_count()

    def refresh_count(self):
        shown = self.grid.Items.Count
        total = len(self.spf.params)
        if shown == total:
            text = u"{0} shared parameters in {1} groups.".format(
                total, len(self.spf.groups))
        else:
            text = u"{0} of {1} shared parameters shown.".format(shown, total)
        if self.dirty:
            text += u"  Unsaved changes."
        self.set_status(text)

    def set_status(self, text):
        self.status_tb.Text = text

    def say_later(self, text, title):
        """Show a dialog once the current WPF event has unwound."""
        def show():
            message(text, title=title)
        self.Dispatcher.BeginInvoke(DispatcherPriority.Background,
                                    Action(show))

    def selected_rows(self):
        return list(self.grid.SelectedItems)

    def need_selection(self):
        rows = self.selected_rows()
        if not rows:
            message(u"Select one or more rows first.", title=u"Shared Params")
            return None
        return rows

    # --- undo ---

    def snapshot(self):
        self.undo_stack.append((copy.deepcopy(self.spf.params),
                                copy.deepcopy(self.spf.groups)))
        if len(self.undo_stack) > UNDO_DEPTH:
            self.undo_stack.pop(0)

    def undo_clicked(self, sender, args):
        if not self.undo_stack:
            message(u"Nothing to undo.", title=u"Shared Params")
            return
        params, groups = self.undo_stack.pop()
        self.spf.params = params
        self.spf.groups = groups
        self.mark_dirty()
        self.after_file_change()
        self.set_status(u"Undone. {0} steps left.".format(len(self.undo_stack)))

    # --- file buttons ---

    def open_clicked(self, sender, args):
        if not self.check_discard():
            return
        path = forms.pick_file(file_ext="txt",
                               title="Pick a shared parameter file")
        if path:
            self.load_file(path)

    def revit_clicked(self, sender, args):
        path = None
        try:
            path = HOST_APP.app.SharedParametersFilename
        except Exception:
            pass
        if not path or not os.path.isfile(path):
            message(
                u"Revit is not pointed at a shared parameter file, or the "
                u"file it names is not there.\n\n{0}".format(path or u""),
                title=u"Shared Params")
            return
        if not self.check_discard():
            return
        self.load_file(path)

    def new_clicked(self, sender, args):
        self.new_file()

    def saveas_clicked(self, sender, args):
        current = self.spf.path or u""
        path = save_file_as(
            u"Save shared parameters as",
            os.path.basename(current) or u"Shared Parameters.txt",
            "txt",
            initial_folder=os.path.dirname(current) if current else None)
        if path:
            self.write_file(path)

    def set_revit_clicked(self, sender, args):
        if not self.spf.path:
            message(u"Save the file first, then Revit can be pointed at it.",
                    title=u"Shared Params")
            return
        if self.dirty and not confirm(
                u"There are unsaved changes. Point Revit at the file as it "
                u"is on disk anyway?", title=u"Shared Params", yes=u"Point"):
            return
        try:
            HOST_APP.app.SharedParametersFilename = self.spf.path
            self.set_status(u"Revit now reads shared parameters from this "
                            u"file.")
        except Exception as ex:
            message(u"Revit would not take that path.\n\n{0}".format(ex),
                    title=u"Shared Params")

    # --- row buttons ---

    def add_clicked(self, sender, args):
        if not self.spf.groups:
            self.spf.add_group(u"Default Group")
            self.after_file_change()
        self.snapshot()
        group_id = None
        chosen = self.filter_group_cb.SelectedItem
        if chosen and chosen != ALL_GROUPS:
            group_id = self._group_ids.get(chosen)
        row = self.spf.blank_param(self.spf.unique_name(), group_id or u"")
        self.spf.params.append(row)
        self.mark_dirty()
        self.refresh_rows(keep_selection=False)
        self.select_by_guid([row.get("GUID")])
        self.set_status(u"Added {0}.".format(row.get("NAME")))

    def dup_clicked(self, sender, args):
        rows = self.need_selection()
        if not rows:
            return
        self.snapshot()
        made = []
        for row in rows:
            clone = dict(row.data)
            clone["GUID"] = sp_file.new_guid()
            clone["NAME"] = self.unique_copy_name(clone.get("NAME", u""))
            self.spf.params.append(clone)
            made.append(clone["GUID"])
        self.mark_dirty()
        self.refresh_rows(keep_selection=False)
        self.select_by_guid(made)
        self.set_status(u"Duplicated {0} parameter{1}, each with a fresh "
                        u"GUID.".format(len(made), u"" if len(made) == 1
                                        else u"s"))

    def unique_copy_name(self, name):
        taken = set([p.get("NAME", u"") for p in self.spf.params])
        base = name or u"Parameter"
        candidate = base + u" copy"
        n = 2
        while candidate in taken:
            candidate = u"{0} copy {1}".format(base, n)
            n += 1
        return candidate

    def guid_clicked(self, sender, args):
        rows = self.need_selection()
        if not rows:
            return
        if not confirm(
                u"Give {0} parameter{1} a new GUID?\n\n"
                u"Revit matches a shared parameter by its GUID, so any "
                u"family or project already bound to the old one loses that "
                u"link and the values with it. This is for fixing a "
                u"duplicate, not for tidying.".format(
                    len(rows), u"" if len(rows) == 1 else u"s"),
                title=u"New GUID", yes=u"Replace"):
            return
        self.snapshot()
        touched = []
        for row in rows:
            row.data["GUID"] = sp_file.new_guid()
            touched.append(row.data["GUID"])
        self.mark_dirty()
        self.refresh_rows(keep_selection=False)
        self.select_by_guid(touched)
        self.set_status(u"{0} GUID{1} replaced.".format(
            len(touched), u"" if len(touched) == 1 else u"s"))

    def del_clicked(self, sender, args):
        rows = self.need_selection()
        if not rows:
            return
        names = [r.data.get("NAME", u"") for r in rows]
        preview = u"\n".join(names[:12])
        if len(names) > 12:
            preview += u"\n... and {0} more".format(len(names) - 12)
        if not confirm(
                u"Delete {0} parameter{1} from the file?\n\n{2}".format(
                    len(rows), u"" if len(rows) == 1 else u"s", preview),
                title=u"Delete", yes=u"Delete"):
            return
        self.snapshot()
        doomed = set([id(r.data) for r in rows])
        self.spf.params = [p for p in self.spf.params if id(p) not in doomed]
        self.mark_dirty()
        self.refresh_rows(keep_selection=False)
        self.set_status(u"Deleted {0}. Undo puts them back.".format(
            len(rows)))

    def select_by_guid(self, guids):
        wanted = set(guids)
        self.grid.SelectedItems.Clear()
        last = None
        for row in self.grid.Items:
            if row.data.get("GUID", u"") in wanted:
                self.grid.SelectedItems.Add(row)
                last = row
        if last is not None:
            self.grid.ScrollIntoView(last)

    # --- batch edits ---

    def rename_clicked(self, sender, args):
        rows = self.need_selection()
        if not rows:
            return
        dialog = RenameWindow(self, [r.data.get("NAME", u"") for r in rows])
        dialog.ShowDialog()
        if not dialog.result:
            return
        prefix, suffix, find, with_ = dialog.result
        if not (prefix or suffix or find):
            return
        self.snapshot()
        for row in rows:
            name = row.data.get("NAME", u"")
            if find:
                name = name.replace(find, with_)
            row.data["NAME"] = prefix + name + suffix
        self.mark_dirty()
        self.refresh_rows()
        self.set_status(u"Renamed {0} parameter{1}.".format(
            len(rows), u"" if len(rows) == 1 else u"s"))

    def type_clicked(self, sender, args):
        rows = self.need_selection()
        if not rows:
            return
        # Keyed on discipline and name together: Power, Temperature, Slope
        # and others exist in more than one discipline with different
        # tokens, so the name alone would write the wrong one.
        by_label = {}
        order = []
        for opt in self._type_options:
            key = u"{0}: {1}".format(opt.Group, opt.Label)
            if key not in by_label:
                by_label[key] = opt.Token
                order.append(key)
        picked = forms.SelectFromList.show(
            order, title="Set data type",
            button_name="Apply", multiselect=False)
        if not picked:
            return
        self.snapshot()
        for row in rows:
            row.data["DATATYPE"] = by_label[picked]
        self.mark_dirty()
        self.refresh_rows()
        self.set_status(u"{0} row{1} set to {2}.".format(
            len(rows), u"" if len(rows) == 1 else u"s", picked))

    def movegroup_clicked(self, sender, args):
        rows = self.need_selection()
        if not rows:
            return
        names = sorted(self._group_ids.keys())
        if not names:
            message(u"The file has no groups yet.", title=u"Shared Params")
            return
        picked = forms.SelectFromList.show(
            names, title="Move to group", button_name="Move",
            multiselect=False)
        if not picked:
            return
        group_id = self._group_ids.get(picked)
        if not group_id:
            return
        self.snapshot()
        for row in rows:
            row.data["GROUP"] = group_id
        self.mark_dirty()
        self.refresh_rows()
        self.set_status(u"{0} row{1} moved into {2}.".format(
            len(rows), u"" if len(rows) == 1 else u"s", picked))

    def groups_clicked(self, sender, args):
        self.snapshot()
        dialog = GroupsWindow(self, self.spf)
        dialog.ShowDialog()
        if dialog.changed:
            self.mark_dirty()
            self.after_file_change()
            self.set_status(u"{0} groups.".format(len(self.spf.groups)))
        else:
            self.undo_stack.pop()

    def sort_clicked(self, sender, args):
        picked = choice(
            u"Reorder the rows in the file itself. Clicking a column header "
            u"only sorts what is on screen.",
            [("name", u"By name"),
             ("group", u"By group, then name"),
             ("type", u"By type, then name")],
            title=u"Sort the file")
        if not picked:
            return
        self.snapshot()
        if picked == "name":
            key = lambda p: p.get("NAME", u"").lower()
        elif picked == "group":
            key = lambda p: (self.group_name_for(p.get("GROUP", u"")).lower(),
                             p.get("NAME", u"").lower())
        else:
            key = lambda p: (p.get("DATATYPE", u"").lower(),
                             p.get("NAME", u"").lower())
        self.spf.params.sort(key=key)
        self.mark_dirty()
        self.refresh_rows()
        self.set_status(u"Rows reordered in the file. Save to keep it.")

    # --- filters ---

    def search_changed(self, sender, args):
        if not self._loading:
            self.refresh_rows(keep_selection=False)

    def filter_changed(self, sender, args):
        if not self._loading:
            self.refresh_rows(keep_selection=False)

    def clear_filters(self):
        self._loading = True
        try:
            self.search_tb.Text = u""
            self.filter_group_cb.SelectedIndex = 0
        finally:
            self._loading = False
        self.refresh_rows(keep_selection=False)

    def clear_clicked(self, sender, args):
        self.clear_filters()

    # --- grid events ---

    def cell_edit_ending(self, sender, args):
        """Guard the edit, and snapshot the state it is about to leave.

        The commit happens after this handler returns, so the deepcopy here
        catches the value as it was, which is what Undo needs.
        """
        try:
            column = args.Column
            editor = args.EditingElement
            text = getattr(editor, "Text", None)

            if column is self.grid.Columns[COL_GUID] and text is not None:
                new = (text or u"").strip()
                complaint = None
                if not sp_file.is_guid(new):
                    complaint = (
                        u"{0} is not a GUID.\n\nRevit wants the form "
                        u"8-4-4-4-12, for example {1}".format(
                            new or u"An empty value", sp_file.new_guid()))
                else:
                    clash = [p for p in self.spf.params
                             if p.get("GUID", u"").lower() == new.lower()
                             and p is not args.Row.Item.data]
                    if clash:
                        complaint = (
                            u"That GUID is already on {0}.\n\nTwo parameters "
                            u"sharing a GUID are one parameter as far as "
                            u"Revit is concerned.".format(
                                clash[0].get("NAME", u"another row")))
                if complaint:
                    args.Cancel = True
                    # The dialog waits until the grid has finished unwinding
                    # this edit. Opening a modal window from inside
                    # CellEditEnding re-enters the grid while the cell is
                    # still committing.
                    self.say_later(complaint, u"GUID")
                    return

            self.snapshot()
            self.mark_dirty()
            # The value has not landed yet, so the count and the row colours
            # are refreshed once the commit is through.
            self.Dispatcher.BeginInvoke(
                DispatcherPriority.Background, Action(self.refresh_count))
        except Exception:
            # A guard that throws must not take the edit down with it.
            pass

    def preparing_cell_for_edit(self, sender, args):
        """Give the data type dropdown its discipline headings.

        GroupStyle is a plain collection on ComboBox, not a dependency
        property, so no Style can set it. The column builds a fresh editing
        ComboBox for each edit; it gets the heading style from the window's
        resources here.
        """
        editor = args.EditingElement
        if (args.Column is self.grid.Columns[COL_TYPE]
                and isinstance(editor, ComboBox)
                and editor.GroupStyle.Count == 0):
            style = self.TryFindResource("TypeGroupStyle")
            if isinstance(style, GroupStyle):
                editor.GroupStyle.Add(style)

    def grid_key_down(self, sender, args):
        if args.Key == Key.Delete:
            focused = Keyboard.FocusedElement
            # Delete inside a cell editor is a text edit, not a row delete.
            if isinstance(focused, TextBox):
                return
            if self.selected_rows():
                args.Handled = True
                self.del_clicked(sender, args)
        elif (args.Key == Key.Z
              and Keyboard.Modifiers == ModifierKeys.Control):
            args.Handled = True
            self.undo_clicked(sender, args)

    # --- check and save ---

    def check_clicked(self, sender, args):
        errors, warnings = self.run_check()
        if not errors and not warnings:
            message(u"No problems found in {0} parameters.".format(
                len(self.spf.params)), title=u"Check")
            return
        message(self.problem_text(errors, warnings), title=u"Check")

    def run_check(self):
        """Validate, and paint the offending rows red.

        The row mapping comes back keyed by position rather than by name,
        because two rows can share a name and the one at fault is often the
        duplicate itself.
        """
        errors, warnings, row_errors = self.spf.validate_rows()
        faulty = {}
        for index, messages in row_errors.items():
            faulty[id(self.spf.params[index])] = u"\n".join(messages)

        # A row the filters are hiding would be named in the report and then
        # never marked, which reads as the tool complaining about nothing.
        # So the filters come off when the fault is somewhere off screen.
        if faulty:
            shown = set([id(row.data) for row in self.grid.Items])
            if [key for key in faulty if key not in shown]:
                self.clear_filters()

        for row in self.grid.Items:
            row.flag_problem(faulty.get(id(row.data), u""))
        return errors, warnings

    @staticmethod
    def problem_text(errors, warnings):
        parts = []
        if errors:
            shown = errors[:12]
            parts.append(u"{0} problem{1} that must be fixed:\n\n{2}".format(
                len(errors), u"" if len(errors) == 1 else u"s",
                u"\n".join([u"  " + e for e in shown])))
            if len(errors) > 12:
                parts.append(u"  ... and {0} more".format(len(errors) - 12))
        if warnings:
            shown = warnings[:8]
            parts.append(u"{0} thing{1} worth a look:\n\n{2}".format(
                len(warnings), u"" if len(warnings) == 1 else u"s",
                u"\n".join([u"  " + w for w in shown])))
            if len(warnings) > 8:
                parts.append(u"  ... and {0} more".format(len(warnings) - 8))
        return u"\n\n".join(parts)

    def save_clicked(self, sender, args):
        self.write_file(self.spf.path)

    def write_file(self, path):
        self.grid.CommitEdit()
        errors, warnings = self.run_check()
        if errors:
            message(
                self.problem_text(errors, [])
                + u"\n\nNothing has been written. The rows at fault are "
                  u"marked in red.",
                title=u"Cannot save yet")
            return False
        if warnings and not confirm(
                self.problem_text([], warnings) + u"\n\nSave anyway?",
                title=u"Check", yes=u"Save"):
            return False

        if not path:
            current = self.spf.path or u""
            path = save_file_as(
                u"Save shared parameters as",
                os.path.basename(current) or u"Shared Parameters.txt", "txt",
                initial_folder=os.path.dirname(current) if current else None)
            if not path:
                return False

        try:
            written = self.spf.save(path)
        except Exception as ex:
            message(u"The file was not written, and the copy on disk is "
                    u"untouched.\n\n{0}".format(ex), title=u"Save failed")
            return False

        self.dirty = False
        self.refresh_title()
        self.set_status(u"Saved {0} parameters to {1}".format(
            len(self.spf.params), written))
        remember_file(written)
        return True

    # --- closing ---

    def check_discard(self):
        """True when it is safe to throw the current edits away."""
        if not self.dirty:
            return True
        picked = choice(
            u"There are unsaved changes to this parameter file.",
            [("cancel", u"Cancel"), ("discard", u"Discard"), ("save", u"Save")],
            title=u"Unsaved changes")
        if picked == "save":
            return self.write_file(self.spf.path)
        return picked == "discard"

    def close_clicked(self, sender, args):
        if self.check_discard():
            self.Close()


# ── ENTRY ───────────────────────────────────────────────────────────────────

def first_run_guide():
    """Explain the tool once, then stay quiet about it."""
    if os.path.isfile(GUIDE_FILE):
        return
    message(GUIDE, title=u"Shared Params")
    try:
        folder = os.path.dirname(GUIDE_FILE)
        if not os.path.isdir(folder):
            os.makedirs(folder)
        with open(GUIDE_FILE, "w") as fh:
            fh.write("shown")
    except Exception:
        pass


def opening_path():
    """The file to land on: the last one edited, else Revit's own."""
    path = recall_file()
    if path:
        return path
    try:
        path = HOST_APP.app.SharedParametersFilename
        if path and os.path.isfile(path):
            return path
    except Exception:
        pass
    return None


def main():
    first_run_guide()
    window = SharedParamsWindow("SharedParams.xaml", opening_path())
    window.ShowDialog()


if __name__ == "__main__":
    main()
