# -*- coding: utf-8 -*-
# "pyRecent"
# "Seed43"
# """
# Recently opened project files with thumbnails, an open history per file and
# a Reference flag. Opens, saves and closes are recorded by startup.py through
# Snippets._recent; this window only reads and edits that list.
# """
import os
import os.path as op

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Windows.Forms")

import System
from System.IO import FileStream, FileMode, FileAccess
from System.Windows import (Thickness, HorizontalAlignment, VerticalAlignment,
                            GridLength, GridUnitType, TextTrimming, FontWeights,
                            CornerRadius, Visibility, FrameworkElement)
from System.Windows.Controls import (Grid, ColumnDefinition, StackPanel,
                                     TextBlock, Image, Border, Orientation,
                                     CheckBox, Button)
from System.Windows.Markup import XamlReader
from System.Windows.Media import Stretch, VisualTreeHelper
from System.Windows.Media.Imaging import (BitmapImage, BitmapCacheOption,
                                          BitmapCreateOptions)

from pyrevit import HOST_APP, forms

from Snippets import _recent
from Snippets import _dialogs as dlg
from Snippets._icons import set_header_icon
from Snippets._support import github_issue_url, open_url, support_mailto
from Snippets.seed43_theme import apply_seed43_palette, apply_seed43_dimensions

# ── CONSTANTS ───────────────────────────────────────────────────────────────

SCRIPT_DIR = op.dirname(__file__)
TITLE = u"pyRecent"
STYLES_XAML = op.join(op.dirname(op.abspath(_recent.__file__)), "Seed43Styles.xaml")


# ── HELPERS (private to this file) ──────────────────────────────────────────

# --- Formatting ---
def _fmt_ts(ts):
    """2026-11-17T09:12:03 -> 2026-11-17  09:12"""
    if not ts:
        return u"-"
    return ts.replace(u"T", u"  ")[:17]


def _fmt_duration(sec):
    if sec is None:
        return u""
    sec = int(sec)
    if sec < 60:
        return u"{0}s".format(sec)
    mins = sec // 60
    if mins < 60:
        return u"{0}m".format(mins)
    hours, mins = divmod(mins, 60)
    if hours < 24:
        return u"{0}h {1}m".format(hours, mins) if mins else u"{0}h".format(hours)
    days, hours = divmod(hours, 24)
    return u"{0}d {1}h".format(days, hours) if hours else u"{0}d".format(days)


def _year(version):
    """'2026' / 'Revit 2026' / '26' -> '2026', or None."""
    digits = u"".join(c for c in unicode(version or u"") if c.isdigit())
    if len(digits) >= 4:
        return digits[:4]
    if len(digits) == 2:
        return u"20" + digits
    return None


# --- Opening files ---
def _find_revit_exe(year):
    """Revit.exe for a given year on this machine, or None."""
    for root in (r"C:\Program Files\Autodesk", r"D:\Program Files\Autodesk",
                 r"C:\Program Files (x86)\Autodesk"):
        exe = op.join(root, u"Revit {0}".format(year), u"Revit.exe")
        if op.isfile(exe):
            return exe
    return None


def open_file(path, version):
    """Open `path`, in the Revit version it was last used in if that differs.

    Called after the window has closed. Calling OpenAndActivateDocument from
    inside the window's own message loop is what gave "The old active
    document still has transaction phases open".
    """
    target = _year(version)
    current = _year(HOST_APP.version)
    if target and current and target != current:
        exe = _find_revit_exe(target)
        if exe:
            System.Diagnostics.Process.Start(exe, u'"{0}"'.format(path))
            return
        dlg.message(u"Revit {0} is not installed here, so the file opens in "
                    u"Revit {1}. Saving it will upgrade it.".format(target, current),
                    title=TITLE)
    HOST_APP.uiapp.OpenAndActivateDocument(path)


def _open_document(path):
    for doc in HOST_APP.app.Documents:
        try:
            if doc.PathName and doc.PathName.lower() == path.lower():
                return doc
        except Exception:
            pass
    return None


def regenerate_thumbnail(path):
    """New thumbnail. A closed file is opened in the background, unrecorded.

    Background (OpenDocumentFile) rather than OpenAndActivateDocument: no
    view opens, nothing flickers, and it can run from the window. A
    workshared file is detached so its central model is never touched.
    """
    from Autodesk.Revit import DB
    from Snippets import _recent_thumb
    doc = _open_document(path)
    if doc is not None:
        return _recent_thumb.generate(doc, path)
    _recent.set_skip(True)
    opened = None
    try:
        options = DB.OpenOptions()
        try:
            if DB.BasicFileInfo.Extract(path).IsWorkshared:
                options.DetachFromCentralOption = \
                    DB.DetachFromCentralOption.DetachAndPreserveWorksets
        except Exception:
            pass
        model_path = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(path)
        opened = HOST_APP.app.OpenDocumentFile(model_path, options)
        return _recent_thumb.generate(opened, path)
    finally:
        if opened is not None:
            try:
                opened.Close(False)
            except Exception:
                pass
        _recent.set_skip(False)


# ── UI / ENTRY POINT ────────────────────────────────────────────────────────

class RecentWindow(forms.WPFWindow):
    """The pyRecent window: the Files page, and Options from the header menu."""

    # --- construction ---
    def __init__(self):
        forms.WPFWindow.__init__(self, op.join(SCRIPT_DIR, "pyRecent.xaml"))
        self._load_styles()
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)
        set_header_icon(self, SCRIPT_DIR)

        self.open_request = None        # (path, version), acted on after close
        self._selected = None           # path of the selected file
        self._expanded = None           # path whose open history is showing
        self._example = None            # (parts, any_flags) being edited
        self._multi = False             # multi-select mode for Remove
        self._ticked = set()            # lower-case paths ticked in that mode
        self._loading = True
        self._settings = _recent.load_settings()
        self._load_options()
        self._loading = False

        self._show_page("files")
        self._refresh()

    # --- pages and the header menu ---
    def _show_page(self, name):
        files = name == "files"
        self.page_files.Visibility = Visibility.Visible if files else Visibility.Collapsed
        self.page_options.Visibility = Visibility.Collapsed if files else Visibility.Visible
        self.header_subtitle_tb.Text = u"| Files" if files else u"| Options"
        self.header_normal_btns.Visibility = Visibility.Visible if files else Visibility.Collapsed
        self.options_header_btns.Visibility = Visibility.Collapsed if files else Visibility.Visible

    def _close_menu(self):
        self.menu_popup.IsOpen = False
        self.menu_toggle_btn.IsChecked = False

    def menu_toggle_clicked(self, sender, args):
        self.menu_popup.IsOpen = bool(self.menu_toggle_btn.IsChecked)

    def window_preview_mouse_down(self, sender, args):
        """Close the menu on a click anywhere outside it (see MENU in
        Seed43Styles.xaml for why the Popup cannot do this itself)."""
        if not self.menu_popup.IsOpen:
            return
        node = args.OriginalSource
        while node is not None:
            if node is self.menu_toggle_btn or node is self.menu_border:
                return
            try:
                node = VisualTreeHelper.GetParent(node)
            except Exception:
                node = None
        self._close_menu()

    def options_back_clicked(self, sender, args):
        self.cancel_structure_clicked(None, None)
        self._refresh()
        self._show_page("files")

    def menu_options_clicked(self, sender, args):
        self._close_menu()
        self._show_page("options")

    def help_clicked(self, sender, args):
        self._close_menu()
        open_url(support_mailto(TITLE, SCRIPT_DIR), window=self)

    def issue_clicked(self, sender, args):
        self._close_menu()
        open_url(github_issue_url(TITLE, SCRIPT_DIR), window=self)

    def win_close_clicked(self, sender, args):
        self.Close()

    # --- files list ---
    def _visible_projects(self):
        text = (self.search_tb.Text or u"").lower().strip()
        show_ref = bool(self.show_ref_cb.IsChecked)
        settings = self._settings
        out = []
        for proj in _recent.load_history()["projects"]:
            path = proj.get("path", u"")
            if not show_ref and proj.get("is_reference"):
                continue
            if not _recent.in_tracked_folders(path, settings):
                continue
            if text and text not in proj.get("name", u"").lower() and text not in path.lower():
                continue
            out.append(proj)
        return out

    def _refresh(self):
        self.list_panel.Children.Clear()
        projects = self._visible_projects()
        for proj in projects:
            self.list_panel.Children.Add(self._card(proj))
        total = len(_recent.load_history()["projects"])
        note = u"   \u00b7   folder filter on" if self._settings.get("filter_enabled") else u""
        self.status_tb.Text = u"{0} shown   \u00b7   {1} in the list{2}".format(
            len(projects), total, note)
        if self._multi:
            shown = set(p.get("path", u"").lower() for p in projects)
            self.multi_count_tb.Text = u"{0} ticked".format(len(self._ticked))
            self.tick_all_cb.IsChecked = bool(shown) and shown <= self._ticked

    def _text(self, text, size=12, muted=False, bold=False):
        tb = TextBlock()
        tb.Text = text
        tb.FontSize = size
        tb.TextTrimming = TextTrimming.CharacterEllipsis
        tb.SetResourceReference(TextBlock.ForegroundProperty, "BrushTextPrimary")
        if muted:
            tb.Opacity = 0.65
        if bold:
            tb.FontWeight = FontWeights.SemiBold
        return tb

    def _card(self, proj):
        path = proj.get("path", u"")
        ticked = self._multi and path.lower() in self._ticked
        selected = ticked or (not self._multi and self._selected
                              and self._selected.lower() == path.lower())
        expanded = (not self._multi and self._expanded
                    and self._expanded.lower() == path.lower())

        outer = StackPanel()
        outer.Margin = Thickness(0, 0, 8, 8)

        card = Border()
        card.SetResourceReference(Border.BackgroundProperty, "BrushCardBg")
        card.SetResourceReference(Border.BorderBrushProperty,
                                  "BrushPrimaryGreen" if selected else "BrushBorderMuted")
        card.BorderThickness = Thickness(2 if selected else 1)
        card.CornerRadius = CornerRadius(8, 8, 0 if expanded else 8, 0 if expanded else 8)
        card.Padding = Thickness(12, 10, 12, 10)
        card.Cursor = System.Windows.Input.Cursors.Hand
        card.Tag = path
        card.MouseLeftButtonDown += self._card_clicked
        if not op.exists(path):
            card.Opacity = 0.55
            card.ToolTip = u"This file no longer exists"

        grid = Grid()
        for width in (GridLength.Auto, GridLength(96),
                      GridLength(1, GridUnitType.Star), GridLength.Auto):
            col = ColumnDefinition()
            col.Width = width
            grid.ColumnDefinitions.Add(col)
        card.Child = grid

        # tick box, multi-select mode only; clicking anywhere on the card
        # toggles it too (see _card_clicked)
        if self._multi:
            tick = CheckBox()
            tick.IsChecked = ticked
            tick.IsHitTestVisible = False
            tick.VerticalAlignment = VerticalAlignment.Center
            tick.Margin = Thickness(0, 0, 12, 0)
            tick.SetResourceReference(FrameworkElement.StyleProperty, "CheckBoxStyle")
            grid.Children.Add(tick)

        # thumbnail
        frame = Border()
        frame.Width, frame.Height = 88, 66
        frame.CornerRadius = CornerRadius(6)
        frame.ClipToBounds = True
        frame.SetResourceReference(Border.BackgroundProperty, "BrushInputBg")
        thumb = _recent.thumbnail_path(proj)
        if thumb:
            image = Image()
            image.Stretch = Stretch.UniformToFill
            bitmap = BitmapImage()
            bitmap.BeginInit()
            bitmap.UriSource = System.Uri(thumb)
            bitmap.CacheOption = BitmapCacheOption.OnLoad
            bitmap.CreateOptions = BitmapCreateOptions.IgnoreImageCache
            bitmap.DecodePixelWidth = 176
            bitmap.EndInit()
            image.Source = bitmap
            frame.Child = image
        Grid.SetColumn(frame, 1)
        grid.Children.Add(frame)

        # name, version badge, path, details
        info = StackPanel()
        info.Margin = Thickness(14, 0, 10, 0)
        info.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(info, 2)
        grid.Children.Add(info)

        title = StackPanel()
        title.Orientation = Orientation.Horizontal
        title.Children.Add(self._text(proj.get("name", u"-"), 14, bold=True))
        year = _year(proj.get("revit_version"))
        if year:
            badge = Border()
            badge.SetResourceReference(Border.BackgroundProperty, "BrushPrimaryGreen")
            badge.CornerRadius = CornerRadius(4)
            badge.Padding = Thickness(6, 1, 6, 1)
            badge.Margin = Thickness(10, 0, 0, 0)
            badge.VerticalAlignment = VerticalAlignment.Center
            label = self._text(u"R" + year[2:], 11, bold=True)
            label.Foreground = System.Windows.Media.Brushes.White
            badge.Child = label
            title.Children.Add(badge)
        info.Children.Add(title)

        path_tb = self._text(path, 11, muted=True)
        path_tb.Margin = Thickness(0, 3, 0, 0)
        info.Children.Add(path_tb)

        opens = len(proj.get("opens") or [])
        details = u"Opened {0}   \u00b7   {1} day{2} opened".format(
            _fmt_ts(proj.get("last_opened")), opens, u"" if opens == 1 else u"s")
        if proj.get("last_saved"):
            details += u"   \u00b7   last saved {0}".format(_fmt_ts(proj["last_saved"]))
        meta = self._text(details, 11, muted=True)
        meta.Margin = Thickness(0, 4, 0, 0)
        info.Children.Add(meta)

        # reference flag
        right = StackPanel()
        right.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(right, 3)
        grid.Children.Add(right)
        ref = CheckBox()
        ref.Content = u"Reference"
        ref.IsChecked = bool(proj.get("is_reference"))
        ref.Tag = path
        ref.SetResourceReference(FrameworkElement.StyleProperty, "CheckBoxStyle")
        ref.ToolTip = u"A file you only look at or link, not one you work in"
        ref.Click += self._reference_clicked
        right.Children.Add(ref)
        if self._multi:
            hint_text = u"click to tick" if not ticked else u"click to untick"
        else:
            hint_text = u"click to hide history" if expanded else u"click for history"
        hint = self._text(hint_text, 10, muted=True)
        hint.HorizontalAlignment = HorizontalAlignment.Right
        hint.Margin = Thickness(0, 6, 0, 0)
        right.Children.Add(hint)

        outer.Children.Add(card)
        if expanded:
            outer.Children.Add(self._history_panel(proj))
        return outer

    def _history_panel(self, proj):
        panel = Border()
        panel.SetResourceReference(Border.BackgroundProperty, "BrushInputBg")
        panel.SetResourceReference(Border.BorderBrushProperty, "BrushBorderMuted")
        panel.BorderThickness = Thickness(1, 0, 1, 1)
        panel.CornerRadius = CornerRadius(0, 0, 8, 8)
        panel.Padding = Thickness(114, 10, 16, 12)
        stack = StackPanel()
        panel.Child = stack
        heading = self._text(u"OPEN HISTORY", 11, bold=True)
        heading.SetResourceReference(TextBlock.ForegroundProperty, "BrushPrimaryGreen")
        heading.Margin = Thickness(0, 0, 0, 6)
        stack.Children.Add(heading)
        opens = proj.get("opens") or []
        if not opens:
            stack.Children.Add(self._text(u"No recorded opens yet", muted=True))
        for entry in reversed(opens):
            line = _fmt_ts(entry.get("timestamp"))
            took = _fmt_duration(entry.get("duration_sec"))
            if took:
                line += u"   \u00b7   open {0}".format(took)
            row = self._text(line)
            row.Margin = Thickness(0, 2, 0, 2)
            stack.Children.Add(row)
        return panel

    # --- files events ---
    def _card_clicked(self, sender, args):
        path = sender.Tag
        if self._multi:
            key = path.lower()
            if key in self._ticked:
                self._ticked.discard(key)
            else:
                self._ticked.add(key)
            self._refresh()
            return
        if args.ClickCount == 2:
            self._selected = path
            self.open_clicked(None, None)
            return
        self._selected = path
        self._expanded = None if (self._expanded or u"").lower() == path.lower() else path
        self._refresh()

    def _reference_clicked(self, sender, args):
        _recent.set_reference(sender.Tag, bool(sender.IsChecked))
        if not self.show_ref_cb.IsChecked:
            self._refresh()

    def _selected_project(self):
        if not self._selected:
            dlg.message(u"Select a file first.", title=TITLE)
            return None
        return _recent.find_project(_recent.load_history(), self._selected)

    def open_clicked(self, sender, args):
        proj = self._selected_project()
        if not proj:
            return
        if not op.exists(proj["path"]):
            dlg.message(u"This file no longer exists:\n{0}".format(proj["path"]), title=TITLE)
            return
        self.open_request = (proj["path"], proj.get("revit_version"))
        self.Close()

    # --- removing ---
    def remove_clicked(self, sender, args):
        """One file selected: confirm it, or switch to ticking several.
        Nothing selected: straight to ticking."""
        proj = None
        if self._selected:
            proj = _recent.find_project(_recent.load_history(), self._selected)
        if proj is None:
            self._set_multi(True)
            return
        key = dlg.choice(
            u"Remove {0} from the list?\nThe file itself is not touched.".format(
                proj.get("name")),
            [("cancel", u"Cancel"), ("multi", u"Select multiple"),
             ("remove", u"Remove")],
            title=u"Remove selected file")
        if key == "remove":
            _recent.remove_project(proj["path"])
            self._selected = self._expanded = None
            self._refresh()
        elif key == "multi":
            self._ticked = set([proj["path"].lower()])
            self._set_multi(True)

    def _set_multi(self, on):
        self._multi = on
        if not on:
            self._ticked = set()
        self.normal_tools.Visibility = Visibility.Collapsed if on else Visibility.Visible
        self.multi_tools.Visibility = Visibility.Visible if on else Visibility.Collapsed
        self._refresh()

    def multi_done_clicked(self, sender, args):
        if not self._ticked:
            dlg.message(u"Tick the files to remove first, or press Cancel.",
                        title=TITLE)
            return
        count = len(self._ticked)
        key = dlg.choice(
            u"Remove the {0} ticked file{1} from the list?\nThe files "
            u"themselves are not touched.".format(count, u"" if count == 1 else u"s"),
            [("cancel", u"Cancel"), ("more", u"Select more"),
             ("remove", u"Remove {0}".format(count))],
            title=u"Remove selected files")
        if key == "remove":
            ticked = set(self._ticked)
            _recent.remove_where(lambda p: p.get("path", u"").lower() in ticked)
            self._selected = self._expanded = None
            self._set_multi(False)
        elif key == "cancel":
            self._set_multi(False)
        # "more" (or closing the popup) keeps the ticks and stays in ticking

    def multi_cancel_clicked(self, sender, args):
        self._set_multi(False)

    def tick_all_clicked(self, sender, args):
        if self.tick_all_cb.IsChecked:
            self._ticked = set(p.get("path", u"").lower()
                               for p in self._visible_projects())
        else:
            self._ticked = set()
        self._refresh()

    def regen_clicked(self, sender, args):
        proj = self._selected_project()
        if not proj:
            return
        if not op.exists(proj["path"]):
            dlg.message(u"This file no longer exists:\n{0}".format(proj["path"]), title=TITLE)
            return
        self.status_tb.Text = u"Making a thumbnail of {0}...".format(proj.get("name"))
        self.Cursor = System.Windows.Input.Cursors.Wait
        try:
            done = regenerate_thumbnail(proj["path"])
        except Exception as err:
            done = None
            dlg.message(u"Could not make a thumbnail:\n{0}".format(err), title=TITLE)
        finally:
            self.Cursor = None
        if done is None:
            dlg.message(u"No thumbnail was made. The model may have nothing "
                        u"to show in 3D.", title=TITLE)
        self._refresh()

    def show_ref_clicked(self, sender, args):
        self._refresh()

    def search_changed(self, sender, args):
        self._refresh()

    # --- options ---
    def _load_options(self):
        s = self._settings
        self.filter_cb.IsChecked = bool(s.get("filter_enabled"))
        self.auto_ref_cb.IsChecked = bool(s.get("auto_reference", True))
        self.thumb_cb.IsChecked = bool(s.get("thumb_on_open", True))
        self.folders_lb.Items.Clear()
        for folder in s.get("folders") or []:
            self.folders_lb.Items.Add(folder)

    def _save_options(self):
        if self._loading:
            return
        s = self._settings
        s["filter_enabled"] = bool(self.filter_cb.IsChecked)
        s["auto_reference"] = bool(self.auto_ref_cb.IsChecked)
        s["thumb_on_open"] = bool(self.thumb_cb.IsChecked)
        s["folders"] = [unicode(item) for item in self.folders_lb.Items]
        _recent.save_settings(s)

    def filter_clicked(self, sender, args):
        if self.filter_cb.IsChecked and not self.folders_lb.Items.Count:
            dlg.message(u"Browse to an example folder to add a structure; "
                        u"until then every file is still tracked.", title=TITLE)
        self._save_options()

    def option_clicked(self, sender, args):
        self._save_options()

    # --- folder structures ---
    def browse_folder_clicked(self, sender, args):
        folder = forms.pick_folder(title=u"Pick one job's folder as an example")
        if not folder:
            return
        parts = _recent.split_folder(folder)
        self._example = (parts, _recent.suggest_any(parts))
        self._show_structure()

    def _show_structure(self):
        """Draw the example folder as one pill per folder, 'any' ones lit."""
        parts, flags = self._example
        self.segments_panel.Children.Clear()
        for i, (part, is_any) in enumerate(zip(parts, flags)):
            pill = Button()
            pill.SetResourceReference(FrameworkElement.StyleProperty, "PillButtonStyle")
            pill.Margin = Thickness(0, 0, 6, 6)
            pill.Tag = u"Viewing" if is_any else u"Enabled"
            pill.Content = u"any folder" if is_any else part
            pill.ToolTip = (u"{0}\nClick to keep this folder as it is".format(part)
                            if is_any else u"Click if this folder changes from job to job")
            pill.Click += self._make_pill_handler(i)
            self.segments_panel.Children.Add(pill)
        structure = _recent.build_structure(parts, flags)
        self.pattern_tb.Text = structure
        files = _recent.load_history()["projects"]
        hits = sum(1 for p in files if _recent.matches_structure(
            op.dirname(p.get("path", u"")), structure))
        self.pattern_hint_tb.Text = (
            u"Tracks every file in this structure and the folders under it. "
            u"{0} of the {1} files in the list match.".format(hits, len(files)))
        self.structure_card.Visibility = Visibility.Visible

    def _make_pill_handler(self, index):
        def handler(sender, args):
            parts, flags = self._example
            flags = list(flags)
            flags[index] = not flags[index]
            self._example = (parts, flags)
            self._show_structure()
        return handler

    def add_structure_clicked(self, sender, args):
        if not self._example:
            return
        structure = _recent.build_structure(*self._example)
        existing = [unicode(item).lower() for item in self.folders_lb.Items]
        if structure.lower() not in existing:
            self.folders_lb.Items.Add(structure)
        if not self.filter_cb.IsChecked:
            self.filter_cb.IsChecked = True
        self.cancel_structure_clicked(None, None)
        self._save_options()

    def cancel_structure_clicked(self, sender, args):
        self._example = None
        self.segments_panel.Children.Clear()
        self.structure_card.Visibility = Visibility.Collapsed

    def remove_folder_clicked(self, sender, args):
        item = self.folders_lb.SelectedItem
        if item is not None:
            self.folders_lb.Items.Remove(item)
            self._save_options()

    def remove_missing_clicked(self, sender, args):
        self._close_menu()
        count = _recent.remove_where(lambda p: not op.exists(p.get("path", u"")))
        dlg.message(u"{0} missing file{1} taken off the list.".format(
            count, u"" if count == 1 else u"s"), title=TITLE)

        self._refresh()

    def remove_untracked_clicked(self, sender, args):
        self._close_menu()
        settings = self._settings
        if not settings.get("filter_enabled") or not settings.get("folders"):
            dlg.message(u"Turn on the folder filter in Options and add a folder "
                        u"structure first.", title=TITLE)
            return
        count = _recent.remove_where(
            lambda p: not _recent.in_tracked_folders(p.get("path", u""), settings))
        dlg.message(u"{0} file{1} outside the tracked folders taken off the "
                    u"list.".format(count, u"" if count == 1 else u"s"), title=TITLE)
        self._refresh()

    # --- private helpers ---
    def _load_styles(self):
        """Merge the shared Seed43 control styles into this window."""
        stream = FileStream(STYLES_XAML, FileMode.Open, FileAccess.Read)
        try:
            self.Resources.MergedDictionaries.Add(XamlReader.Load(stream))
        finally:
            stream.Close()


def main():
    window = RecentWindow()
    window.ShowDialog()
    if window.open_request:
        path, version = window.open_request
        try:
            open_file(path, version)
        except Exception as err:
            dlg.message(u"Could not open the file:\n{0}\n\n{1}".format(path, err),
                        title=TITLE)


if __name__ == "__main__":
    main()
