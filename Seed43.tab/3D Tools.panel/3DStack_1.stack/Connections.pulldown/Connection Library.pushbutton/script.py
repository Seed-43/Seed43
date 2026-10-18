# -*- coding: utf-8 -*-
# "Connection Library"
# "Seed43"
# """
# A shelf of steel connections kept in a Revit file, so a connection set up
# once can be used on every job after it.
#
# WHY THE LIBRARY IS AN .RVT AND NOT A JSON FILE
#     The Modify Parameters values - plate sizes, bolt grades, layouts - are
#     the whole point of saving a connection, and they live in the Advance
#     Steel object store. No property, no parameter and no extensible
#     storage entity reaches them, so a tool cannot read them out to a file
#     of its own and cannot write them back. The only thing that moves them
#     is Revit's own CopyElements, and that needs a real document at both
#     ends. The library is therefore an ordinary .rvt holding nothing but
#     connection types.
#
#     This is also why the library cannot be built out of Create: Create
#     names the family but resets the settings. See _connections.py.
#
# VERSIONS: FORWARD ONLY
#     A .rvt opens in its own version or a newer one, never an older one, and
#     saving in a newer Revit upgrades the file for good. So the library
#     belongs on the OLDEST Revit still in use, and this tool will not write
#     into an older library without saying so first: one silent save would
#     lock out everyone who has not upgraded. A library from a NEWER Revit
#     cannot be opened at all, and is reported as such rather than failing
#     somewhere less obvious.
#
# Target: Revit 2022-2026, IronPython 2.
# """

# ── IMPORTS ─────────────────────────────────────────────────────────────────

import json
import os

from Autodesk.Revit.DB import BasicFileInfo, SaveAsOptions, Transaction
from pyrevit import forms, revit
from pyrevit.framework import Windows

from Snippets import _userdata
from Snippets._connections import (connection_types, eid, element_name,
                                   family_name, placed_counts, transfer)
from Snippets.seed43_theme import (apply_seed43_dimensions,
                                   apply_seed43_palette, get_color)

doc = revit.doc
app = doc.Application

# ── CONSTANTS ───────────────────────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TOOL_NAME = "Connection Library"

# The library path is one user's own, not a shipped default: it lives in
# .user so an update cannot overwrite it. migrate() carries across a copy
# left beside the script by an earlier version.
SETTINGS_FILE = _userdata.migrate(
    os.path.join(SCRIPT_DIR, "Settings", "library.json"),
    _userdata.user_path(TOOL_NAME, "library.json"))

DEFAULT_LIBRARY = os.path.join(
    os.path.expanduser("~"), "Documents", "Seed43 Connection Library.rvt")

# Column widths, mirrored in ConnectionLibrary.xaml's two header rows. The
# two must move together or the headings drift off the data.
COL_FAMILY = 170
COL_PLACED = 62


# ── SETTINGS ────────────────────────────────────────────────────────────────

def load_path():
    """The library path last used, or the default if there is none yet."""
    try:
        with open(SETTINGS_FILE, "r") as handle:
            data = json.load(handle)
        path = data.get("path")
        if path:
            return path
    except Exception:
        pass
    return DEFAULT_LIBRARY


def save_path(path):
    """Remember the library path. Never raises: this is a convenience."""
    try:
        settings_dir = os.path.dirname(SETTINGS_FILE)
        if not os.path.isdir(settings_dir):
            os.makedirs(settings_dir)
        with open(SETTINGS_FILE, "w") as handle:
            json.dump({"path": path}, handle, indent=2)
    except Exception:
        pass


# ── THE LIBRARY FILE ────────────────────────────────────────────────────────

class LibraryError(Exception):
    """A library that cannot be used, with a sentence saying why."""


def inspect(path):
    """(format, is_later, is_current) for a library file on disk.

    Read without opening the document, so a library from a newer Revit can
    be reported as unusable instead of throwing somewhere less obvious.
    """
    info = BasicFileInfo.Extract(path)
    return (info.Format, info.IsSavedInLaterVersion,
            info.IsSavedInCurrentVersion)


class Library(object):
    """An open library document, and the connection types it holds."""

    # --- construction ---
    def __init__(self, path):
        self.path = path
        self.document = None
        self.format = u""
        self.older = False      # written by an older Revit than this one
        self.dirty = False      # has unsaved changes from this session
        self._open()

    def _open(self):
        if not os.path.isfile(self.path):
            raise LibraryError(
                u"No library yet. Press New to make one, Browse to point at "
                u"an existing one, or just highlight connections and press "
                u"Save and it will offer to make one.")
        try:
            self.format, later, current = inspect(self.path)
        except Exception as ex:
            raise LibraryError(
                u"{} is not a readable Revit file: {}".format(self.path, ex))
        if later:
            raise LibraryError(
                u"This library was saved by Revit {} and cannot be opened by "
                u"this one. Revit files only open forward.".format(
                    self.format))
        self.older = not current
        try:
            self.document = app.OpenDocumentFile(self.path)
        except Exception as ex:
            raise LibraryError(
                u"Could not open the library: {}. If it is open in another "
                u"Revit session or another window, close it there "
                u"first.".format(ex))

    # --- public methods ---
    def rows(self):
        """Every connection type in the library."""
        return connection_types(self.document)

    def save(self):
        """Write the library back to disk if this session changed it."""
        if not self.dirty:
            return
        self.document.Save()
        self.dirty = False

    def close(self):
        """Save and close. Safe to call twice."""
        if self.document is None:
            return
        try:
            self.save()
        finally:
            try:
                self.document.Close(False)
            except Exception:
                pass
            self.document = None


def create_library(path):
    """Make an empty library at path.

    Built from the default project template rather than from nothing, since
    the API has no way to make a bare document. Its stock connection types
    come along, which is harmless: they are the same ones every model
    already has, so they simply show as already present.
    """
    fresh = app.NewProjectDocument(app.DefaultProjectTemplate)
    options = SaveAsOptions()
    options.OverwriteExistingFile = False
    try:
        fresh.SaveAs(path, options)
    finally:
        fresh.Close(False)


# ── UI ──────────────────────────────────────────────────────────────────────

def brush(key, fallback="#FFFFFF"):
    """A resolved brush for a palette key, for text built in Python."""
    return Windows.Media.SolidColorBrush(
        Windows.Media.ColorConverter.ConvertFromString(
            get_color(SCRIPT_DIR, key, fallback=fallback)))


class LibraryWindow(forms.WPFWindow):
    """This model on the left, the library on the right, arrows between."""

    # --- construction ---
    def __init__(self, xaml_name):
        forms.WPFWindow.__init__(self, xaml_name)
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)

        # Resolved after the palette is applied, never before. A brush built
        # from an unresolved lookup comes back transparent, and the row text
        # then renders invisible with no error to explain it.
        self._muted = brush("text_muted", "#9CA3AF")
        self._primary = brush("text_primary", "#FFFFFF")
        self._green = brush("primary_green", "#208A3C")

        self.library = None
        self._warned_upgrade = False
        self._model_rows = []
        self._lib_rows = []

        self.path_tb.Text = load_path()
        self.open_library(load_path(), quiet=True)

    # --- public methods ---
    def open_library(self, path, quiet=False):
        """Point at a library file and read it, or report why not."""
        self.close_library()
        try:
            self.library = Library(path)
        except LibraryError as ex:
            self.library = None
            self.refresh(unicode(ex))
            return False
        except Exception as ex:
            self.library = None
            self.refresh(u"Could not open the library: {}".format(ex))
            return False

        save_path(path)
        self.path_tb.Text = path
        self.refresh(None if quiet else u"Opened {}.".format(
            os.path.basename(path)))
        return True

    def close_library(self):
        if self.library is not None:
            self.library.close()
            self.library = None

    def refresh(self, note=None):
        """Re-read both sides and rebuild the lists."""
        counts = placed_counts(doc)
        self._model_rows = connection_types(doc)
        self._lib_rows = self.library.rows() if self.library else []

        in_library = set((family_name(s), element_name(s))
                         for s in self._lib_rows)

        self.model_lb.Items.Clear()
        for symbol in self._model_rows:
            self.model_lb.Items.Add(self._model_item(symbol, counts,
                                                     in_library))
        self.lib_lb.Items.Clear()
        for symbol in self._lib_rows:
            self.lib_lb.Items.Add(self._library_item(symbol))

        self.model_head_tb.Text = u"This model  ({})".format(
            len(self._model_rows))
        self.lib_head_tb.Text = (u"Library  ({})".format(len(self._lib_rows))
                                 if self.library else u"Library  (not open)")
        self.subtitle_tb.Text = (
            u"|  {}".format(os.path.basename(self.library.path))
            if self.library else u"|  no library")

        if self.library is None:
            self.path_tb.Text = u"(no library open)"

        self.explain_tb.Text = self._explain()
        self._sync_buttons()
        if note:
            self.status_tb.Text = note

    # --- event handlers ---
    def selection_changed(self, sender, args):
        self._sync_buttons()

    def browse_clicked(self, sender, args):
        picked = forms.pick_file(file_ext="rvt",
                                 title="Pick the connection library")
        if picked:
            self.open_library(picked)

    def new_clicked(self, sender, args):
        """Create an empty library and switch to it."""
        self.make_library()

    def make_library(self):
        """Ask for a path, make the library there, and open it.

        A path that already holds a file is opened rather than refused. The
        save dialog has already asked about overwriting by that point, and
        somebody who picks an existing library plainly means to use it;
        sending them back to Browse to do the same thing twice would be the
        tool being pedantic about which button was pressed.
        """
        picked = forms.save_file(file_ext="rvt",
                                 default_name="Seed43 Connection Library",
                                 title="Create a connection library")
        if not picked:
            return False
        if not picked.lower().endswith(".rvt"):
            picked += ".rvt"
        if os.path.isfile(picked):
            return self.open_library(picked)
        self.close_library()
        try:
            create_library(picked)
        except Exception as ex:
            forms.alert(u"The library was not created:\n\n{}".format(ex),
                        title=TOOL_NAME)
            self.refresh()
            return False
        return self.open_library(picked)

    def push_clicked(self, sender, args):
        """Copy the highlighted model connections into the library."""
        chosen = self._chosen(self.model_lb)
        if not chosen:
            self.status_tb.Text = u"Highlight connections in this model first."
            return
        # Deliberately after the selection check and not behind _ready(). On
        # a machine with no library yet, Save is the button somebody reaches
        # for, and a greyed one with the reason in a status line is a dead
        # end: offer to make the library from here instead.
        if self.library is None and not self._offer_library():
            return
        if not self._allow_write():
            return
        self._run(doc, chosen, self.library.document,
                  u"into the library", mark_dirty=True)

    def pull_clicked(self, sender, args):
        """Copy the highlighted library connections into this model."""
        if not self._ready():
            return
        chosen = self._chosen(self.lib_lb)
        if not chosen:
            self.status_tb.Text = u"Highlight connections in the library first."
            return
        self._run(self.library.document, chosen, doc,
                  u"into this model", mark_dirty=True)

    def delete_clicked(self, sender, args):
        """Remove the highlighted connections from the library."""
        if not self._ready():
            return
        chosen = self._chosen(self.lib_lb)
        if not chosen:
            self.status_tb.Text = u"Highlight connections in the library first."
            return
        if not self._allow_write():
            return

        names = [u"{} : {}".format(family_name(s), element_name(s))
                 for s in chosen]
        if not forms.alert(
                u"Delete {} connection(s) from the library?\n\n{}\n\nThis "
                u"cannot be undone once the library is saved.".format(
                    len(names), u"\n".join(names[:12])),
                title=TOOL_NAME, yes=True, no=True):
            return

        removed, failed = 0, []
        t = Transaction(self.library.document, "Delete connection types")
        t.Start()
        try:
            for symbol in chosen:
                label = u"{} : {}".format(family_name(symbol),
                                          element_name(symbol))
                try:
                    self.library.document.Delete(symbol.Id)
                    removed += 1
                except Exception as ex:
                    failed.append(u"{} - {}".format(label, ex))
            t.Commit()
        except Exception as ex:
            t.RollBack()
            forms.alert(u"Nothing was deleted:\n\n{}".format(ex),
                        title=TOOL_NAME)
            self.refresh()
            return

        self.library.dirty = True
        note = u"Deleted {} from the library.".format(removed)
        if failed:
            forms.alert(u"{} could not be deleted:\n\n{}".format(
                len(failed), u"\n".join(failed[:10])), title=TOOL_NAME)
        self.refresh(note)

    def close_clicked(self, sender, args):
        self.Close()

    def window_closed(self, sender, args):
        self.close_library()

    # --- private helpers ---
    def _ready(self):
        if self.library is None:
            self.status_tb.Text = (
                u"No library is open. Use Browse or New first.")
            return False
        return True

    def _offer_library(self):
        """No library yet: offer to make or find one. True if there is one now."""
        answer = forms.alert(
            u"There is no connection library open yet.\n\nA library is an "
            u"ordinary Revit file that holds nothing but connection types. "
            u"Make one now, or point at one that already exists?",
            title=TOOL_NAME, options=["Make a new library",
                                      "Find an existing one", "Cancel"])
        if answer == "Make a new library":
            return self.make_library()
        if answer == "Find an existing one":
            picked = forms.pick_file(file_ext="rvt",
                                     title="Pick the connection library")
            return bool(picked) and self.open_library(picked)
        return False

    def _allow_write(self):
        """Ask once before the first save into an older library.

        Saving upgrades the file for good, and a library that has quietly
        become a 2026 file is one nobody on 2024 can open again. Once per
        session is enough: after the first yes the decision has been made.
        """
        if not self.library.older or self._warned_upgrade:
            return True
        if forms.alert(
                u"This library was saved by Revit {}. Writing to it will "
                u"upgrade the file to Revit {}, and it will no longer open "
                u"in any older version.\n\nGo ahead?".format(
                    self.library.format, app.VersionNumber),
                title=TOOL_NAME, yes=True, no=True):
            self._warned_upgrade = True
            return True
        return False

    def _run(self, source_doc, symbols, dest_doc, where, mark_dirty):
        """Do one transfer and report it."""
        try:
            report = transfer(source_doc, symbols, dest_doc)
        except Exception as ex:
            forms.alert(u"The transfer stopped:\n\n{}".format(ex),
                        title=TOOL_NAME)
            self.refresh()
            return

        # Loud, and before the lists are rebuilt. A stranded source is the
        # one outcome where this tool has left a model worse than it found
        # it: copying renames the original, and if the name could not be put
        # back the type is sitting there called "Ref".
        if report.stranded:
            forms.alert(
                u"{} connection(s) could not have their name put back after "
                u"being copied, and are now named \"Ref\" with no Type "
                u"Mark:\n\n{}\n\nRename these by hand before saving.".format(
                    len(report.stranded),
                    u"\n".join(u"{} : {}".format(f, n)
                               for f, n in report.stranded[:10])),
                title=TOOL_NAME)

        if mark_dirty and report.copied and self.library is not None:
            self.library.dirty = True

        parts = [u"Copied {} {}.".format(len(report.copied), where)]
        if report.present:
            parts.append(u"{} already there.".format(len(report.present)))
        if report.reused:
            parts.append(u"{} matched an existing type and were left "
                         u"alone.".format(len(report.reused)))
        if report.failed:
            parts.append(u"{} failed.".format(len(report.failed)))
            forms.alert(u"These did not copy:\n\n{}".format(
                u"\n".join(report.failed[:10])), title=TOOL_NAME)
        self.refresh(u" ".join(parts))

    def _explain(self):
        if self.library is None:
            return (u"A library is an ordinary Revit file holding nothing but "
                    u"connection types. Copying is the only thing that keeps "
                    u"the Modify Parameters values, so the shelf has to be a "
                    u".rvt rather than a file of this tool's own.")
        if self.library.older:
            return (u"Saved by Revit {}, and this is Revit {}. Writing to it "
                    u"upgrades the file for good and older Revits will no "
                    u"longer open it, so keep the library on the oldest "
                    u"version still in use.".format(self.library.format,
                                                    app.VersionNumber))
        return (u"Revit {}, same as this session. Save copies connections "
                u"out, Load copies them in, and both keep the Modify "
                u"Parameters values that a clone would reset.".format(
                    self.library.format))

    def _sync_buttons(self):
        has_library = self.library is not None
        model_picked = len(self.model_lb.SelectedItems)
        lib_picked = len(self.lib_lb.SelectedItems)

        # Save stays live without a library on purpose: pressing it is how
        # somebody with no library yet gets one. Load and Delete cannot be,
        # since there is nothing on that side to act on.
        self.push_btn.IsEnabled = bool(model_picked)
        self.pull_btn.IsEnabled = has_library and bool(lib_picked)
        self.delete_btn.IsEnabled = has_library and bool(lib_picked)
        self.push_btn.Content = (u"»»  Save {}".format(model_picked)
                                 if model_picked else u"»»  Save")
        self.pull_btn.Content = (u"Load {}  ««".format(lib_picked)
                                 if lib_picked else u"Load  ««")

    def _chosen(self, listbox):
        return [item.Tag for item in listbox.SelectedItems]

    def _model_item(self, symbol, counts, in_library):
        placed = counts.get(eid(symbol.Id), 0)
        panel = Windows.Controls.StackPanel()
        panel.Orientation = Windows.Controls.Orientation.Horizontal
        panel.Children.Add(self._cell(family_name(symbol), COL_FAMILY,
                                      self._muted))
        panel.Children.Add(self._cell(
            unicode(placed) if placed else u"-", COL_PLACED,
            self._green if placed else self._muted))
        # Green marks a type the library already holds, so a second pass
        # over a model does not mean reading every name to find what is new.
        already = (family_name(symbol), element_name(symbol)) in in_library
        panel.Children.Add(self._cell(
            element_name(symbol), 0,
            self._green if already else self._primary, bold=True))
        return self._item(panel, symbol)

    def _library_item(self, symbol):
        panel = Windows.Controls.StackPanel()
        panel.Orientation = Windows.Controls.Orientation.Horizontal
        panel.Children.Add(self._cell(family_name(symbol), COL_FAMILY,
                                      self._muted))
        panel.Children.Add(self._cell(element_name(symbol), 0, self._primary,
                                      bold=True))
        return self._item(panel, symbol)

    def _item(self, panel, symbol):
        item = Windows.Controls.ListBoxItem()
        item.Content = panel
        item.Tag = symbol
        return item

    def _cell(self, text, width, colour=None, bold=False):
        block = Windows.Controls.TextBlock()
        block.Text = text or u""
        block.TextTrimming = Windows.TextTrimming.CharacterEllipsis
        block.VerticalAlignment = Windows.VerticalAlignment.Center
        block.Margin = Windows.Thickness(0, 0, 10, 0)
        block.ToolTip = text or None
        if width:
            block.Width = width
        if colour is not None:
            block.Foreground = colour
        if bold:
            block.FontWeight = Windows.FontWeights.SemiBold
        return block


# ── ENTRY POINT ─────────────────────────────────────────────────────────────

def main():
    window = LibraryWindow("ConnectionLibrary.xaml")
    try:
        window.ShowDialog()
    finally:
        # Belt and braces over the Closed handler. A library left open holds
        # a Revit lock on the file, so it must close even if the window came
        # down some way the handler did not see.
        window.close_library()


main()
