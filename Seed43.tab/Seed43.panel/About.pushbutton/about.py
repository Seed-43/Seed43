# -*- coding: utf-8 -*-
# about.py
import os
import clr
import shutil
import zipfile

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("System.Net")

from System.Windows.Markup import XamlReader
from System.Windows import (
    MessageBox, MessageBoxButton, MessageBoxImage, Visibility, Thickness,
    Duration, CornerRadius, FrameworkElement, HorizontalAlignment, VerticalAlignment
)
from System.Windows.Controls import (
    StackPanel, Border, TextBlock, DockPanel, Dock
)
from System.Windows.Input import Cursors, MouseButtonState
from System.Windows.Media import SolidColorBrush, ColorConverter, Brushes
from System.Windows.Media.Animation import (
    ThicknessAnimation, ColorAnimation, CubicEase, EasingMode
)
from System.Net import WebClient
from System.IO import File, StreamReader
from System.Threading import Thread, ThreadStart
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind, Action, TimeSpan
from threading import Lock

# ── VARIABLES ─────────────────────────────────────────────────────────────────
GITHUB_ORG    = "Seed-43"
MAIN_REPO     = "Seed43"
BRANCH        = "main"

APPDATA       = os.environ.get("APPDATA", "")
EXTENSION_DIR = os.path.join(APPDATA, "pyRevit", "Extensions", "Seed43.extension")
VERSION_FILE  = os.path.join(EXTENSION_DIR, "version.txt")
VERSION_URL   = "https://raw.githubusercontent.com/{}/{}/{}/version.txt".format(
    GITHUB_ORG, MAIN_REPO, BRANCH)
REPO_ZIP_URL  = "https://github.com/{}/{}/archive/refs/heads/{}.zip".format(
    GITHUB_ORG, MAIN_REPO, BRANCH)

# ── Load XAML ─────────────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(__file__)
XAML_PATH  = os.path.join(SCRIPT_DIR, "About.xaml")
ICON_PATH  = os.path.join(SCRIPT_DIR, "icon.png")

# Walk up to find the enclosing .tab folder for the tool scanner
TAB_DIR  = None
_current = SCRIPT_DIR
for _ in range(10):
    if _current.endswith('.tab'):
        TAB_DIR = _current
        break
    _parent = os.path.dirname(_current)
    if _parent == _current:
        break
    _current = _parent

# ── Helpers ───────────────────────────────────────────────────────────────────

def load_xaml(path):
    reader = StreamReader(path)
    window = XamlReader.Load(reader.BaseStream)
    reader.Close()
    return window

def read_local_version():
    try:
        if not File.Exists(VERSION_FILE):
            return "0.0.0"
        reader  = StreamReader(VERSION_FILE)
        content = reader.ReadToEnd().strip()
        reader.Close()
        for line in content.splitlines():
            line = line.strip()
            if line:
                return line
    except Exception:
        pass
    return "0.0.0"

def read_last_update():
    try:
        if not File.Exists(VERSION_FILE):
            return []
        reader  = StreamReader(VERSION_FILE)
        content = reader.ReadToEnd()
        reader.Close()
        notes      = []
        in_section = False
        for line in content.splitlines():
            stripped = line.strip()
            if stripped == "Last update:":
                in_section = True
                continue
            if in_section:
                if stripped.startswith("___"):
                    break
                if not stripped:
                    continue
                if stripped.startswith("-"):
                    notes.append(u"  " + stripped)
                else:
                    notes.append(stripped)
        return notes
    except Exception:
        return []

def fetch_remote_version():
    try:
        client = WebClient()
        raw    = client.DownloadString(VERSION_URL).strip()
        for line in raw.splitlines():
            line = line.strip()
            if line:
                return line
    except Exception:
        return None

def version_tuple(version_str):
    try:
        return tuple(int(x) for x in version_str.strip().split("."))
    except Exception:
        return (0, 0, 0)

def dispatch(window, fn):
    window.Dispatcher.Invoke(Action(fn))

# ── Migration helpers ─────────────────────────────────────────────────────────

def read_migrations(extracted_root):
    """
    Parse seed43_migrations.yaml from the extracted repo root.
    Returns a list of migration dicts:
      [{'from': 'rel/path', 'to': 'rel/path',
        'subfolders': [{'from': 'x', 'to': 'y'}, ...]}, ...]
    Paths are relative to Seed43.tab.
    Returns [] if file not found or parse fails.
    """
    yaml_path = os.path.join(extracted_root, "seed43_migrations.yaml")
    if not os.path.exists(yaml_path):
        return []
    migrations = []
    current    = None
    try:
        with open(yaml_path, "r") as f:
            lines = f.readlines()
        for line in lines:
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            # Top-level migration entry
            if stripped.startswith("- from:"):
                if current:
                    migrations.append(current)
                current = {
                    'from':       stripped[len("- from:"):].strip(),
                    'to':         None,
                    'subfolders': []
                }
            elif stripped.startswith("to:") and current and current['to'] is None:
                current['to'] = stripped[len("to:"):].strip()
            elif stripped.startswith("subfolders:"):
                pass  # section marker, nothing to capture
            # Subfolder entry (indented with - from:)
            elif stripped.startswith("- from:") and current:
                # Already handled above as new top-level entry check
                # but subfolders use same syntax indented
                pass
            # Handle subfolder entries: lines like "  - from: x"
        # Subfolder lines have extra indent — re-parse properly
        current    = None
        migrations = []
        in_sub     = False
        sub_from   = None
        for line in lines:
            stripped = line.strip()
            indent   = len(line) - len(line.lstrip())
            if not stripped or stripped.startswith("#"):
                continue
            if stripped.startswith("migrations:"):
                continue
            # Top-level migration (indent 2, starts with "- from:")
            if indent <= 4 and stripped.startswith("- from:"):
                if current:
                    migrations.append(current)
                current  = {'from': stripped[len("- from:"):].strip(),
                            'to': None, 'subfolders': []}
                in_sub   = False
                sub_from = None
            elif indent <= 4 and stripped.startswith("to:") and current and current['to'] is None:
                current['to'] = stripped[len("to:"):].strip()
            elif stripped == "subfolders:":
                in_sub = True
            elif in_sub and stripped.startswith("- from:"):
                sub_from = stripped[len("- from:"):].strip()
            elif in_sub and stripped.startswith("to:") and sub_from:
                current['subfolders'].append({
                    'from': sub_from,
                    'to':   stripped[len("to:"):].strip()
                })
                sub_from = None
        if current:
            migrations.append(current)
    except Exception:
        return []
    return [m for m in migrations if m.get('from') and m.get('to')]


def apply_migrations(migrations, tab_dst):
    """
    For each migration entry, copy .json files from the old folder path
    to the new folder path inside tab_dst, then for each subfolder do the same.
    Paths in the migration are relative to Seed43.tab.
    Old folders are left in place for sync_tree to clean up.
    """
    for m in migrations:
        old_dir = os.path.join(tab_dst, m['from'])
        new_dir = os.path.join(tab_dst, m['to'])

        if not os.path.isdir(old_dir):
            continue  # nothing to migrate

        # Ensure destination exists
        if not os.path.isdir(new_dir):
            os.makedirs(new_dir)

        # Copy json files from old root to new root
        for fname in os.listdir(old_dir):
            if fname.lower().endswith(".json"):
                src = os.path.join(old_dir, fname)
                dst = os.path.join(new_dir, fname)
                if not os.path.exists(dst):
                    shutil.copy2(src, dst)

        # Handle subfolders
        for sub in m.get('subfolders', []):
            old_sub = os.path.join(old_dir, sub['from'])
            new_sub = os.path.join(new_dir, sub['to'])
            if not os.path.isdir(old_sub):
                continue
            if not os.path.isdir(new_sub):
                os.makedirs(new_sub)
            for fname in os.listdir(old_sub):
                if fname.lower().endswith(".json"):
                    src = os.path.join(old_sub, fname)
                    dst = os.path.join(new_sub, fname)
                    if not os.path.exists(dst):
                        shutil.copy2(src, dst)

# ── Sync helper ───────────────────────────────────────────────────────────────

def sync_tree(src, dst, skipped=None):
    """Sync src into dst file by file.
    Tries to copy/update all files. If a file is locked by Revit the
    exception is caught silently and the file is added to the skipped list.
    Removes dst files and dirs not in src, skipping locked ones silently."""
    if skipped is None:
        skipped = []
    if not os.path.isdir(dst):
        os.makedirs(dst)
    src_names = set(os.listdir(src))
    dst_names = set(os.listdir(dst))
    for name in src_names:
        s = os.path.join(src, name)
        d = os.path.join(dst, name)
        if os.path.isdir(s):
            sync_tree(s, d, skipped)
        else:
            try:
                shutil.copy2(s, d)
            except Exception:
                skipped.append(d)
    for name in dst_names:
        if name not in src_names:
            d = os.path.join(dst, name)
            if os.path.isfile(d):
                try:
                    os.remove(d)
                except Exception:
                    skipped.append(d)
            elif os.path.isdir(d):
                try:
                    shutil.rmtree(d)
                except Exception:
                    skipped.append(d)
    return skipped

# ── YAML order helpers ────────────────────────────────────────────────────────

def read_yaml_layout(folder_path):
    yaml_path = os.path.join(folder_path, "bundle.yaml")
    try:
        if not os.path.exists(yaml_path):
            return []
        with open(yaml_path, "r") as f:
            lines = f.readlines()
        in_layout = False
        order     = []
        for line in lines:
            stripped = line.strip()
            if stripped.startswith("layout:"):
                in_layout = True
                continue
            if in_layout:
                if stripped.startswith("- "):
                    order.append(stripped[2:].strip())
                elif stripped and not stripped.startswith("#"):
                    in_layout = False
        return order
    except Exception:
        return []

def apply_yaml_order(items, folder_path):
    order = read_yaml_layout(folder_path)
    if not order:
        return items
    index   = {name: i for i, name in enumerate(order)}
    known   = [it for it in items if it['name'] in index]
    unknown = [it for it in items if it['name'] not in index]
    known.sort(key=lambda it: index[it['name']])
    return known + unknown

# ── Tool scanner helpers ───────────────────────────────────────────────────────

def folder_ext(name):
    base = name[:-4] if name.lower().endswith('.off') else name
    idx  = base.rfind('.')
    return base[idx:].lower() if idx != -1 else None

def strip_ext(name, ext):
    base = name[:-4] if name.lower().endswith('.off') else name
    cut  = base.lower().rfind(ext)
    return base[:cut] if cut != -1 else base

def is_panel_folder(name):
    return folder_ext(name) == '.panel'

def has_script(folder_path):
    try:
        return any(
            f.lower().endswith('.py') or f.lower().endswith('.xaml')
            for f in os.listdir(folder_path)
        )
    except Exception:
        return False

def scan_pushbuttons(folder_path):
    buttons = []
    try:
        entries = sorted(os.listdir(folder_path))
    except Exception:
        return buttons
    for name in entries:
        ext  = folder_ext(name)
        path = os.path.join(folder_path, name)
        if not os.path.isdir(path):
            continue
        if ext == '.pushbutton' and has_script(path):
            buttons.append({'type': 'button', 'name': strip_ext(name, '.pushbutton'), 'path': path})
    return buttons

def scan_panel(folder_path):
    items = []
    try:
        children = sorted(os.listdir(folder_path))
    except Exception:
        return items
    for child_name in children:
        child_path = os.path.join(folder_path, child_name)
        if not os.path.isdir(child_path):
            continue
        ext = folder_ext(child_name)
        if ext == '.pushbutton':
            if has_script(child_path):
                items.append({'type': 'button',
                              'name': strip_ext(child_name, '.pushbutton'),
                              'path': child_path})
        elif ext == '.pulldown':
            items.append({'type': 'pulldown',
                          'name': strip_ext(child_name, '.pulldown'),
                          'path': child_path,
                          'children': scan_pushbuttons(child_path)})
        elif ext == '.splitpushbutton':
            items.append({'type': 'splitpushbutton',
                          'name': strip_ext(child_name, '.splitpushbutton'),
                          'path': child_path,
                          'children': scan_pushbuttons(child_path)})
        elif ext == '.stack':
            stack_items = []
            try:
                stack_children = sorted(os.listdir(child_path))
            except Exception:
                continue
            for sc_name in stack_children:
                sc_path = os.path.join(child_path, sc_name)
                if not os.path.isdir(sc_path):
                    continue
                sc_ext = folder_ext(sc_name)
                if sc_ext == '.pushbutton' and has_script(sc_path):
                    stack_items.append({'type': 'button',
                                        'name': strip_ext(sc_name, '.pushbutton'),
                                        'path': sc_path})
                elif sc_ext == '.pulldown':
                    stack_items.append({'type': 'pulldown',
                                        'name': strip_ext(sc_name, '.pulldown'),
                                        'path': sc_path,
                                        'children': scan_pushbuttons(sc_path)})
                elif sc_ext == '.splitpushbutton':
                    stack_items.append({'type': 'splitpushbutton',
                                        'name': strip_ext(sc_name, '.splitpushbutton'),
                                        'path': sc_path,
                                        'children': scan_pushbuttons(sc_path)})
            if stack_items:
                items.append({'type': 'stack',
                              'name': strip_ext(child_name, '.stack'),
                              'path': child_path,
                              'children': stack_items})
    return items

# ── Folder toggle logic ────────────────────────────────────────────────────────

class FolderRenamer(object):
    def __init__(self, folder_path, parent=None):
        self.folder_path = folder_path
        self.handlers    = []
        self.parent      = parent
        self._lock       = Lock()

    def sync(self):
        with self._lock:
            self._do_sync()

    def _do_sync(self):
        any_on        = any(h.is_on for h in self.handlers)
        currently_off = self.folder_path.lower().endswith('.off')
        if any_on and currently_off:
            new_path = self.folder_path[:-4]
        elif not any_on and not currently_off:
            new_path = self.folder_path + '.off'
        else:
            if self.parent:
                self.parent.sync()
            return
        try:
            if os.path.exists(new_path):
                shutil.rmtree(new_path)
            os.rename(self.folder_path, new_path)
            old_path         = self.folder_path
            self.folder_path = new_path
            for h in self.handlers:
                h.path = h.path.replace(old_path, new_path, 1)
        except Exception:
            pass
        if self.parent:
            self.parent.sync()

class FolderHandler(object):
    ON_COLOR  = "#208A3C"
    OFF_COLOR = "#A0AABB"

    def __init__(self, window, path, renamer):
        self.window  = window
        self.path    = path
        self.is_on   = not path.lower().endswith('.off')
        self.switch  = None
        self.knob    = None
        self.busy    = False
        self.renamer = renamer

    def animate(self, turn_on):
        duration                 = Duration(TimeSpan.FromMilliseconds(140))
        ease                     = CubicEase()
        ease.EasingMode          = EasingMode.EaseOut
        knob_anim                = ThicknessAnimation()
        knob_anim.Duration       = duration
        knob_anim.To             = Thickness(22, 2, 0, 2) if turn_on else Thickness(2, 2, 0, 2)
        knob_anim.EasingFunction = ease
        self.knob.BeginAnimation(FrameworkElement.MarginProperty, knob_anim)
        color_anim          = ColorAnimation()
        color_anim.Duration = duration
        color_anim.To       = ColorConverter.ConvertFromString(self.ON_COLOR if turn_on else self.OFF_COLOR)
        self.switch.Background.BeginAnimation(SolidColorBrush.ColorProperty, color_anim)

    def toggle(self, sender, args):
        if self.busy:
            return
        self.busy = True
        def worker():
            try:
                new_path = self.path + '.off' if self.is_on else self.path[:-4]
                if os.path.exists(new_path):
                    shutil.rmtree(new_path)
                os.rename(self.path, new_path)
                self.path  = new_path
                self.is_on = not self.is_on
                self.renamer.sync()
                def done():
                    self.animate(self.is_on)
                    self.busy = False
                dispatch(self.window, done)
            except Exception as e:
                def fail():
                    self.busy = False
                    MessageBox.Show(str(e))
                dispatch(self.window, fail)
        Thread(ThreadStart(worker)).Start()

# ── Tool UI builder ───────────────────────────────────────────────────────────

class ToolManager(object):

    def __init__(self, window):
        self.window    = window
        self.container = window.FindName("tools_container")

    def build(self):
        if not self.container:
            return
        self.container.Children.Clear()
        for panel in self._scan():
            self.container.Children.Add(
                self._panel_ui(panel['name'], panel['path'], panel['items'])
            )

    def _scan(self):
        panels = []
        if not TAB_DIR or not os.path.isdir(TAB_DIR):
            return panels
        try:
            entries = sorted(os.listdir(TAB_DIR))
        except Exception:
            return panels
        for name in entries:
            path = os.path.join(TAB_DIR, name)
            if not os.path.isdir(path) or not is_panel_folder(name):
                continue
            display = strip_ext(name, '.panel')
            if display.lower() in ('seed43', 'about'):
                continue
            panels.append({'name': display, 'path': path, 'items': scan_panel(path)})
        return apply_yaml_order(panels, TAB_DIR)

    def _panel_ui(self, name, panel_path, items):
        renamer = FolderRenamer(panel_path)
        body    = StackPanel()

        for item in apply_yaml_order(items, panel_path):
            if item['type'] == 'button':
                body.Children.Add(self._tool_row(item, renamer))
            elif item['type'] in ('pulldown', 'splitpushbutton'):
                body.Children.Add(self._tool_row(item, renamer))
            elif item['type'] == 'stack':
                stack_renamer = FolderRenamer(item['path'], parent=renamer)
                for child in apply_yaml_order(item['children'], item['path']):
                    if child['type'] == 'button':
                        body.Children.Add(self._tool_row(child, stack_renamer))
                    elif child['type'] in ('pulldown', 'splitpushbutton'):
                        body.Children.Add(self._tool_row(child, stack_renamer))

        header = self._make_collapsible_header(name, body)

        outer = StackPanel()
        outer.Children.Add(header)
        outer.Children.Add(body)

        card       = Border()
        card.Style = self.window.FindResource("Card")
        card.Child = outer
        return card

    def _make_collapsible_header(self, label_text, body):
        body.Visibility   = Visibility.Collapsed
        header            = Border()
        header.Padding    = Thickness(6, 6, 10, 6)
        header.Background = Brushes.Transparent
        header.Cursor     = Cursors.Hand

        dock = DockPanel()

        title       = TextBlock()
        title.Text  = label_text
        title.Style = self.window.FindResource("Title")

        arrow        = TextBlock()
        arrow.Text   = u"\u25BC"
        arrow.Margin = Thickness(6, 0, 0, 0)
        arrow.Style  = self.window.FindResource("Title")

        DockPanel.SetDock(arrow, Dock.Right)
        dock.Children.Add(arrow)
        dock.Children.Add(title)
        header.Child = dock

        def toggle(s, e):
            if body.Visibility == Visibility.Collapsed:
                body.Visibility = Visibility.Visible
                arrow.Text      = u"\u25B2"
            else:
                body.Visibility = Visibility.Collapsed
                arrow.Text      = u"\u25BC"

        header.MouseLeftButtonUp += toggle
        return header

    def _tool_row(self, item, renamer):
        path  = item['path']
        name  = item['name']
        is_on = not path.lower().endswith('.off')

        label                   = TextBlock()
        label.Text              = name
        label.Style             = self.window.FindResource("ToolText")
        label.VerticalAlignment = VerticalAlignment.Center

        switch              = Border()
        switch.Width        = 40
        switch.Height       = 20
        switch.CornerRadius = CornerRadius(10)
        switch.Cursor       = Cursors.Hand
        switch.Background   = SolidColorBrush(
            ColorConverter.ConvertFromString(
                FolderHandler.ON_COLOR if is_on else FolderHandler.OFF_COLOR))

        knob                     = Border()
        knob.Width               = 16
        knob.Height              = 16
        knob.CornerRadius        = CornerRadius(8)
        knob.Background          = Brushes.White
        knob.HorizontalAlignment = HorizontalAlignment.Left
        knob.Margin              = Thickness(22, 2, 0, 2) if is_on else Thickness(2, 2, 0, 2)
        switch.Child             = knob

        handler        = FolderHandler(self.window, path, renamer)
        handler.switch = switch
        handler.knob   = knob
        switch.MouseLeftButtonUp += handler.toggle
        renamer.handlers.append(handler)

        row        = DockPanel()
        row.Margin = Thickness(6, 4, 6, 4)
        DockPanel.SetDock(switch, Dock.Right)
        row.Children.Add(switch)
        row.Children.Add(label)
        return row

# ── Main dialog ───────────────────────────────────────────────────────────────

class Seed43Dialog(object):

    def __init__(self):
        self.window = load_xaml(XAML_PATH)

        if os.path.exists(ICON_PATH):
            img           = self.window.FindName("header_icon")
            bmp           = BitmapImage()
            bmp.BeginInit()
            bmp.UriSource = Uri(ICON_PATH, UriKind.Absolute)
            bmp.EndInit()
            img.Source    = bmp

        self._bind()
        self._init_tools()
        self._check_versions()

    def _bind(self):
        self.window.FindName("footer_reload").Click             += self._on_reload
        self.window.FindName("header_update_btn").Click         += self._on_s43_update
        self.window.FindName("update_ribbon").MouseLeftButtonUp += self._on_s43_update
        self.window.FindName("issues_btn").Click                += self._on_issues

    def _init_tools(self):
        self._tool_manager = ToolManager(self.window)
        self._tool_manager.build()

    def _on_reload(self, sender, args):
        self.window.Close()
        try:
            from pyrevit.loader import sessionmgr
            sessionmgr.reload_pyrevit()
        except Exception as ex:
            MessageBox.Show(
                "Could not reload PyRevit:\n\n" + str(ex),
                "Reload Failed",
                MessageBoxButton.OK,
                MessageBoxImage.Warning
            )

    def _check_versions(self):
        def worker():
            local  = read_local_version()
            notes  = read_last_update()
            remote = fetch_remote_version()
            dispatch(self.window, lambda: self._update_s43_ui(local, notes, remote))
        t = Thread(ThreadStart(worker))
        t.IsBackground = True
        t.Start()

    def _render_changelog(self, notes):
        from System.Windows import Thickness, FontWeights, TextWrapping
        container = self.window.FindName("s43_changelog_panel")
        if not container:
            self.window.FindName("s43_changelog").Text = "\n".join(notes)
            return
        container.Children.Clear()
        for note in notes:
            stripped = note.strip()
            if stripped.startswith("-"):
                from System.Windows.Controls import Grid, ColumnDefinition
                from System.Windows import GridLength, GridUnitType
                g    = Grid()
                c0   = ColumnDefinition()
                c0.Width = GridLength(16)
                c1   = ColumnDefinition()
                c1.Width = GridLength(1, GridUnitType.Star)
                g.ColumnDefinitions.Add(c0)
                g.ColumnDefinitions.Add(c1)
                g.Margin = Thickness(8, 1, 0, 1)
                dash = TextBlock()
                dash.Text       = u"-"
                dash.FontSize   = 12
                dash.Foreground = SolidColorBrush(ColorConverter.ConvertFromString("#F4FAFF"))
                dash.Opacity    = 0.9
                Grid.SetColumn(dash, 0)
                g.Children.Add(dash)
                body = TextBlock()
                body.Text        = stripped[1:].strip()
                body.FontSize    = 12
                body.TextWrapping = TextWrapping.Wrap
                body.Foreground  = SolidColorBrush(ColorConverter.ConvertFromString("#F4FAFF"))
                body.Opacity     = 0.9
                Grid.SetColumn(body, 1)
                g.Children.Add(body)
                container.Children.Add(g)
            else:
                tb            = TextBlock()
                tb.Text       = stripped
                tb.FontSize   = 12
                tb.Foreground = SolidColorBrush(ColorConverter.ConvertFromString("#208A3C"))
                tb.FontWeight = FontWeights.SemiBold
                tb.Margin     = Thickness(0, 6, 0, 2)
                container.Children.Add(tb)

    def _update_s43_ui(self, local, notes, remote):
        self.window.FindName("s43_title").Text = u"\u25CF  Installed  v{}".format(local) if local else "Version unknown"
        self._render_changelog(notes)
        if remote and version_tuple(remote) > version_tuple(local):
            self._remote_s43_version = remote
            self.window.FindName("update_ribbon_version").Text = \
                u"v{}  \u2192  v{}".format(local, remote)
            self.window.FindName("update_ribbon").Visibility = Visibility.Visible
        elif not remote:
            extra = [u"Could not reach GitHub to check for updates."]
            self._render_changelog(notes + extra)

    def _show_card(self, name):
        for panel in ("card_normal", "card_confirm", "card_done"):
            p = self.window.FindName(panel)
            if p:
                p.Visibility = Visibility.Visible if panel == name else Visibility.Collapsed

    def _start_spinner(self):
        from System.Windows.Threading import DispatcherTimer
        from System import TimeSpan
        self._spinner_dots = 0
        self._spinner_timer = DispatcherTimer()
        self._spinner_timer.Interval = TimeSpan.FromMilliseconds(400)
        def tick(s, e):
            self._spinner_dots = (self._spinner_dots + 1) % 4
            dots = u"." * self._spinner_dots
            container = self.window.FindName("s43_changelog_panel")
            if container:
                container.Children.Clear()
                from System.Windows import Thickness
                tb = TextBlock()
                tb.Text       = u"Updating" + dots
                tb.FontSize   = 12
                tb.Foreground = SolidColorBrush(ColorConverter.ConvertFromString("#208A3C"))
                tb.Margin     = Thickness(0, 2, 0, 2)
                container.Children.Add(tb)
        self._spinner_timer.Tick += tick
        self._spinner_timer.Start()

    def _stop_spinner(self):
        try:
            self._spinner_timer.Stop()
        except Exception:
            pass

    def _on_s43_update(self, sender, args):
        self._show_card("card_confirm")

        def on_cancel(s, e):
            self._show_card("card_normal")

        def on_yes(s, e):
            self._show_card("card_normal")
            self._do_update()

        cancel_btn = self.window.FindName("confirm_cancel_btn")
        yes_btn    = self.window.FindName("confirm_yes_btn")
        try:
            cancel_btn.Click -= on_cancel
            yes_btn.Click    -= on_yes
        except Exception:
            pass
        cancel_btn.Click += on_cancel
        yes_btn.Click    += on_yes

    def _do_update(self):
        self.window.FindName("update_ribbon").Visibility = Visibility.Collapsed

        EXTENSIONS_DIR = os.path.join(os.environ.get("APPDATA", ""), "pyRevit", "Extensions")
        S43_INSTALL    = os.path.join(EXTENSIONS_DIR, "Seed43.extension")
        TAB_DIR_DEST   = os.path.join(S43_INSTALL, "Seed43.tab")
        TEMP_ZIP       = os.path.join(os.environ.get("TEMP", ""), "seed43_update.zip")
        TEMP_DIR       = os.path.join(os.environ.get("TEMP", ""), "seed43_update_extracted")

        self._start_spinner()

        def log(msg):
            dispatch(self.window, lambda: setattr(
                self.window.FindName("s43_changelog"), "Text", msg))

        def worker():
            try:
                log("Downloading...")
                wc = WebClient()
                wc.Headers.Add("Cache-Control", "no-cache, no-store")
                wc.DownloadFile(REPO_ZIP_URL, TEMP_ZIP)

                log("Extracting...")
                if os.path.exists(TEMP_DIR):
                    shutil.rmtree(TEMP_DIR)
                os.makedirs(TEMP_DIR)
                with zipfile.ZipFile(TEMP_ZIP, "r") as z:
                    z.extractall(TEMP_DIR)

                extracted_root = None
                for item in os.listdir(TEMP_DIR):
                    full = os.path.join(TEMP_DIR, item)
                    if os.path.isdir(full):
                        extracted_root = full
                        break
                if not extracted_root:
                    raise Exception("Could not find extracted folder.")

                new_tab = os.path.join(extracted_root, "Seed43.tab")
                if not os.path.exists(new_tab):
                    raise Exception("Seed43.tab not found in download.")

                # ── Apply migrations before syncing ───────────────────────────
                log("Checking migrations...")
                migrations = read_migrations(extracted_root)
                if migrations:
                    apply_migrations(migrations, TAB_DIR_DEST)

                log("Installing...")
                skipped = sync_tree(new_tab, TAB_DIR_DEST)

                # ── Sync root files (startup.py, extension.json, etc.) ────────
                ROOT_SKIP = {
                    "Seed43.tab", "lib", "UI",
                    ".git", ".gitignore", "README.md", "LICENSE",
                    "install.bat", "sync-start.bat", "sync-end.bat",
                }
                for fname in os.listdir(extracted_root):
                    if fname in ROOT_SKIP:
                        continue
                    src = os.path.join(extracted_root, fname)
                    dst = os.path.join(S43_INSTALL, fname)
                    if os.path.isfile(src):
                        try:
                            shutil.copy2(src, dst)
                        except Exception:
                            skipped.append(dst)

                new_version_file = os.path.join(extracted_root, "version.txt")
                if os.path.isfile(new_version_file):
                    shutil.copy2(new_version_file, VERSION_FILE)

                version = fetch_remote_version() or "unknown"
                log("Done, v{0}".format(version))

                try:
                    if os.path.exists(TEMP_ZIP):
                        os.remove(TEMP_ZIP)
                except Exception:
                    pass
                try:
                    if os.path.exists(TEMP_DIR):
                        shutil.rmtree(TEMP_DIR)
                except Exception:
                    pass

                _skipped = skipped[:]
                dispatch(self.window, lambda: self._stop_spinner())
                dispatch(self.window, lambda: self._on_s43_update_done(version, _skipped))

            except Exception as ex:
                dispatch(self.window, lambda: self._stop_spinner())
                dispatch(self.window, lambda: self._on_error(str(ex)))

        t = Thread(ThreadStart(worker))
        t.IsBackground = True
        t.Start()

    def _on_s43_update_done(self, version, skipped=None):
        from System.Windows import Thickness, FontWeights, TextWrapping
        self.window.FindName("s43_title").Text = u"\u25CF  Installed  v{0}".format(version)
        title = self.window.FindName("card_done_title")
        if title:
            title.Text = u"Seed43 Updated to v{0}.".format(version)

        # Build done message as separate TextBlocks so newlines render correctly
        panel = self.window.FindName("card_done_panel")
        if panel:
            panel.Children.Clear()

            def add_line(text, dim=False):
                tb = TextBlock()
                tb.Text        = text
                tb.FontSize    = 12
                tb.TextWrapping = TextWrapping.Wrap
                tb.Foreground  = SolidColorBrush(ColorConverter.ConvertFromString(
                    "#A0AABB" if dim else "#F4FAFF"))
                tb.Opacity     = 0.7 if dim else 0.9
                tb.Margin      = Thickness(0, 1, 0, 1)
                panel.Children.Add(tb)

            if skipped:
                names = [os.path.basename(p) for p in skipped]
                tb = TextBlock()
                tb.Margin = Thickness(0, 8, 0, 2)
                tb.Text = u"Revit has {0} file{1} open that could not be updated:".format(
                    len(names), "s" if len(names) != 1 else "")
                tb.FontSize    = 12
                tb.TextWrapping = TextWrapping.Wrap
                tb.Foreground  = SolidColorBrush(ColorConverter.ConvertFromString("#D4720A"))
                panel.Children.Add(tb)
                for n in names:
                    add_line(u"  - " + n, dim=True)
                tb2 = TextBlock()
                tb2.Text      = u"Reload PyRevit to apply the remaining changes."
                tb2.FontSize  = 12
                tb2.TextWrapping = TextWrapping.Wrap
                tb2.Foreground = SolidColorBrush(ColorConverter.ConvertFromString("#F4FAFF"))
                tb2.Opacity   = 0.9
                tb2.Margin    = Thickness(0, 8, 0, 0)
                panel.Children.Add(tb2)


        self._show_card("card_done")

        def on_ok(s, e):
            self.window.Close()
            try:
                from pyrevit.loader import sessionmgr
                sessionmgr.reload_pyrevit()
            except Exception as ex:
                MessageBox.Show(
                    "Please reload PyRevit manually.\n\n" + str(ex),
                    "Reload Required",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                )

        ok_btn = self.window.FindName("done_ok_btn")
        try:
            ok_btn.Click -= on_ok
        except Exception:
            pass
        ok_btn.Click += on_ok

    def _on_issues(self, sender, args):
        import System
        psi = System.Diagnostics.ProcessStartInfo("https://github.com/Seed-43/Seed43/issues")
        psi.UseShellExecute = True
        System.Diagnostics.Process.Start(psi)

    def _on_error(self, msg):
        MessageBox.Show("Operation failed:\n\n" + msg, "Seed43",
                        MessageBoxButton.OK, MessageBoxImage.Error)

    def show(self):
        self.window.ShowDialog()

# ── Entry point ───────────────────────────────────────────────────────────────
dialog = Seed43Dialog()
dialog.show()
