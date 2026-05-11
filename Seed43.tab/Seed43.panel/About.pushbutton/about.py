# -*- coding: utf-8 -*-
# about.py
import os
import clr
import json
import shutil
import zipfile

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("System.Net")

import System
from System.Windows.Markup import XamlReader
from System.Windows import (
    MessageBox, MessageBoxButton, MessageBoxImage, Visibility, Thickness, Duration
)
from System.Windows.Controls import StackPanel, Border, TextBlock, DockPanel
from System.Windows.Media import SolidColorBrush, ColorConverter
from System.Windows.Media.Animation import (
    ThicknessAnimation, ColorAnimation, CubicEase, EasingMode
)
import System.Net
import System.IO
from System.Net import WebClient
from System.IO import File, Directory, Path, StreamReader
from System.Threading import Thread, ThreadStart
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind
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
    """Read the installed version string from the extension root version.txt.
    Returns only the first non-empty line as the version number."""
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
    """Read the Last update notes from the extension root version.txt.
    Returns a list of bullet point strings, or an empty list if not found."""
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
                if stripped.startswith("-"):
                    notes.append(stripped)
        return notes
    except Exception:
        return []

def fetch_remote_version():
    """Download version.txt from GitHub and return the version number string.
    Returns None if the request fails."""
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
    """Convert a version string like 1.2.3 to a tuple (1, 2, 3) for comparison."""
    try:
        return tuple(int(x) for x in version_str.strip().split("."))
    except Exception:
        return (0, 0, 0)

def dispatch(window, fn):
    window.Dispatcher.Invoke(System.Action(fn))

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
    """Return button dicts for every .pushbutton directly inside folder_path."""
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

def scan_panel(panel_path):
    """
    Return a structured list of items found inside a .panel folder.

    Supported types:
        .pushbutton     -> {'type': 'button', ...}
        .pulldown       -> {'type': 'pulldown',  'children': [buttons]}
        .splitpushbutton-> {'type': 'splitpushbutton', 'children': [buttons]}
        .stack          -> {'type': 'stack', 'name': str, 'children': [any of the above]}
    """
    items = []
    try:
        children = sorted(os.listdir(panel_path))
    except Exception:
        return items

    for child_name in children:
        child_path = os.path.join(panel_path, child_name)
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
        duration            = Duration(System.TimeSpan.FromMilliseconds(140))
        ease                = CubicEase()
        ease.EasingMode     = EasingMode.EaseOut
        knob_anim               = ThicknessAnimation()
        knob_anim.Duration      = duration
        knob_anim.To            = Thickness(22, 2, 0, 2) if turn_on else Thickness(2, 2, 0, 2)
        knob_anim.EasingFunction = ease
        self.knob.BeginAnimation(System.Windows.FrameworkElement.MarginProperty, knob_anim)
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

# ── Order persistence ─────────────────────────────────────────────────────────

ORDER_FILE = os.path.join(SCRIPT_DIR, "tool_order.json")

def load_order_data():
    """
    Load the full order JSON.
    Structure:
    {
      "panels": ["PanelA", "PanelB", ...],
      "groups": {
        "C:\\path\\to\\pulldown": ["ToolX", "ToolY"],
        ...
      }
    }
    """
    try:
        if os.path.exists(ORDER_FILE):
            with open(ORDER_FILE, "r") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
    except Exception:
        pass
    return {"panels": [], "groups": {}}

def save_order_data(data):
    """Persist the full order dict to JSON."""
    try:
        with open(ORDER_FILE, "w") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass

def load_order():
    """Return saved panel name order list."""
    return load_order_data().get("panels", [])

def save_order(names):
    """Update just the panels list in the JSON."""
    data = load_order_data()
    data["panels"] = names
    save_order_data(data)

def write_bundle_yaml(names):
    """
    Update the layout: list in the bundle.yaml that sits in TAB_DIR.
    Seed43 is always written first, followed by the user-ordered panels.
    Preserves all other keys in the file.
    """
    if not TAB_DIR:
        return False
    yaml_path = os.path.join(TAB_DIR, "bundle.yaml")
    try:
        if os.path.exists(yaml_path):
            with open(yaml_path, "r") as f:
                lines = f.readlines()
        else:
            lines = []

        # Strip out any existing layout block
        new_lines = []
        skip = False
        for line in lines:
            stripped = line.strip()
            if stripped.startswith("layout:"):
                skip = True
                continue
            if skip:
                if stripped.startswith("- "):
                    continue
                else:
                    skip = False
            new_lines.append(line)

        # Remove trailing blank lines
        while new_lines and new_lines[-1].strip() == "":
            new_lines.pop()

        # Build ordered list: Seed43 pinned first, then user order (excluding Seed43 if present)
        ordered = ["Seed43"] + [n for n in names if n.lower() not in ("seed43", "about")]

        new_lines.append("\nlayout:\n")
        for name in ordered:
            new_lines.append("  - {}\n".format(name))

        with open(yaml_path, "w") as f:
            f.writelines(new_lines)
        return True
    except Exception:
        return False

def load_group_order(group_path):
    """
    Read child order for a group — checks JSON first, then falls back
    to the bundle.yaml inside the group folder.
    """
    # Check JSON store first
    data = load_order_data()
    groups = data.get("groups", {})
    if group_path in groups and groups[group_path]:
        return groups[group_path]

    # Fallback: read from bundle.yaml inside the folder
    yaml_path = os.path.join(group_path, "bundle.yaml")
    try:
        if os.path.exists(yaml_path):
            with open(yaml_path, "r") as f:
                lines = f.readlines()
            in_layout = False
            order = []
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
        pass
    return []

def write_group_yaml(group_path, names):
    """
    Write (or create) bundle.yaml inside a pulldown/split folder.
    Preserves existing keys; only replaces the layout block.
    """
    yaml_path = os.path.join(group_path, "bundle.yaml")
    try:
        if os.path.exists(yaml_path):
            with open(yaml_path, "r") as f:
                lines = f.readlines()
        else:
            # Create a minimal yaml with just the layout
            lines = []

        new_lines = []
        skip = False
        for line in lines:
            stripped = line.strip()
            if stripped.startswith("layout:"):
                skip = True
                continue
            if skip:
                if stripped.startswith("- "):
                    continue
                else:
                    skip = False
            new_lines.append(line)

        while new_lines and new_lines[-1].strip() == "":
            new_lines.pop()

        new_lines.append("\nlayout:\n")
        for name in names:
            new_lines.append("  - {}\n".format(name))

        with open(yaml_path, "w") as f:
            f.writelines(new_lines)
        return True
    except Exception:
        return False

def apply_group_order(children, group_path):
    """Sort child items by the saved layout order in the group's bundle.yaml."""
    saved = load_group_order(group_path)
    if not saved:
        return children
    index   = {name: i for i, name in enumerate(saved)}
    known   = [c for c in children if c['name'] in index]
    unknown = [c for c in children if c['name'] not in index]
    known.sort(key=lambda c: index[c['name']])
    return known + unknown

def apply_order(panels, saved_order):
    """
    Sort panels according to saved_order.
    Any panels not in saved_order are appended at the end in their original order.
    """
    if not saved_order:
        return panels
    index = {name: i for i, name in enumerate(saved_order)}
    known   = [p for p in panels if p['name'] in index]
    unknown = [p for p in panels if p['name'] not in index]
    known.sort(key=lambda p: index[p['name']])
    return known + unknown

# ── Tool UI builder ───────────────────────────────────────────────────────────

class ToolManager(object):
    """Builds and populates the tools_container inside the existing window."""

    # Data format passed during a drag operation
    DRAG_FORMAT = "Seed43PanelCard"

    def __init__(self, window):
        self.window          = window
        self.container       = window.FindName("tools_container")
        self._drag_source    = None
        self._group_registry = {}  # path -> StackPanel (the droppable body)

    def collect_group_orders(self):
        """
        Walk _group_registry and read the current child Tag order from each
        live StackPanel. Returns {path: [name, ...]} for every registered group.
        """
        result = {}
        for path, sp in self._group_registry.items():
            try:
                names = [c.Tag for c in sp.Children if c.Tag]
                if names:
                    result[path] = names
            except Exception:
                pass
        return result

    def build(self):
        if not self.container:
            return
        self.container.Children.Clear()
        panels = apply_order(self._scan(), load_order())
        for panel in panels:
            self.container.Children.Add(
                self._panel_ui(panel['name'], panel['path'], panel['items'])
            )
        self._wire_drop_target()

    # ── Drop target on the container ──────────────────────────────────────────

    def _wire_drop_target(self):
        """Allow the container StackPanel to receive drops and reorder cards."""
        self.container.AllowDrop = True

        def on_drag_over(sender, e):
            if e.Data.GetDataPresent(self.DRAG_FORMAT):
                e.Effects = System.Windows.DragDropEffects.Move
            else:
                e.Effects = System.Windows.DragDropEffects.None
            e.Handled = True

        def on_drop(sender, e):
            if not e.Data.GetDataPresent(self.DRAG_FORMAT):
                return
            dragged = e.Data.GetData(self.DRAG_FORMAT)  # the card Border

            # Find which card the cursor is over (before removal)
            pos      = e.GetPosition(self.container)
            children = list(self.container.Children)
            target   = None
            for child in children:
                pt  = child.TranslatePoint(System.Windows.Point(0, 0), self.container)
                if pos.Y < pt.Y + child.ActualHeight:
                    target = child
                    break

            if target is None or target is dragged:
                return

            src_idx = self.container.Children.IndexOf(dragged)
            dst_idx = self.container.Children.IndexOf(target)

            self.container.Children.Remove(dragged)

            # After removal, if we dragged downward the target index shifts by -1
            insert_at = self.container.Children.IndexOf(target)
            if src_idx < dst_idx:
                insert_at += 1
            self.container.Children.Insert(insert_at, dragged)

            # Persist new order
            names = [c.Tag for c in self.container.Children if c.Tag]
            save_order(names)

        self.container.DragOver += on_drag_over
        self.container.Drop     += on_drop

    # ── Scan ──────────────────────────────────────────────────────────────────

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
        return panels

    # ── Collapsible header ────────────────────────────────────────────────────

    def _make_collapsible_header(self, label_text, body, style_key, arrow_style_key=None,
                                  grip=None):
        body.Visibility   = Visibility.Collapsed
        header            = Border()
        header.Padding    = Thickness(6, 6, 10, 6)
        header.Background = System.Windows.Media.Brushes.Transparent
        header.Cursor     = System.Windows.Input.Cursors.Hand

        dock = DockPanel()

        # Grip handle on the far left (only for panel-level headers)
        if grip is not None:
            DockPanel.SetDock(grip, System.Windows.Controls.Dock.Left)
            dock.Children.Add(grip)

        title       = TextBlock()
        title.Text  = label_text
        title.Style = self.window.FindResource(style_key)

        arrow        = TextBlock()
        arrow.Text   = u"\u25BC"
        arrow.Margin = Thickness(6, 0, 0, 0)
        if arrow_style_key:
            arrow.Style = self.window.FindResource(arrow_style_key)
        else:
            arrow.Foreground = System.Windows.Media.Brushes.Gray

        DockPanel.SetDock(arrow, System.Windows.Controls.Dock.Right)
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

    # ── Panel card (draggable) ────────────────────────────────────────────────

    def _make_grip(self):
        """Build the ⠿ drag handle TextBlock."""
        grip                   = TextBlock()
        grip.Text              = u"\u22EE\u22EE"   # ⋮⋮
        grip.Foreground        = System.Windows.Media.Brushes.Gray
        grip.FontSize          = 11
        grip.Cursor            = System.Windows.Input.Cursors.SizeAll
        grip.VerticalAlignment = System.Windows.VerticalAlignment.Center
        grip.Margin            = Thickness(2, 0, 6, 0)
        grip.ToolTip           = "Drag to reorder"
        return grip

    def _panel_ui(self, name, panel_path, items):
        panel_renamer = FolderRenamer(panel_path, parent=None)

        PANEL_CHILD_FORMAT = "Seed43PanelChild_" + panel_path

        body           = StackPanel()
        body.AllowDrop = True

        # Apply saved order from the panel's own bundle.yaml
        ordered_items = apply_group_order(items, panel_path)

        for item in ordered_items:
            child_card = self._panel_child_ui(item, panel_renamer, PANEL_CHILD_FORMAT, body, panel_path)
            if child_card is not None:
                body.Children.Add(child_card)

        # Drop target for reordering items within the panel
        def on_drag_over(sender, e):
            if e.Data.GetDataPresent(PANEL_CHILD_FORMAT):
                e.Effects = System.Windows.DragDropEffects.Move
            else:
                e.Effects = System.Windows.DragDropEffects.None
            e.Handled = True

        def on_drop(sender, e):
            if not e.Data.GetDataPresent(PANEL_CHILD_FORMAT):
                return
            dragged = e.Data.GetData(PANEL_CHILD_FORMAT)
            pos     = e.GetPosition(body)
            target  = None
            for child in list(body.Children):
                pt = child.TranslatePoint(System.Windows.Point(0, 0), body)
                if pos.Y < pt.Y + child.ActualHeight:
                    target = child
                    break
            if target is None or target is dragged:
                return
            src_idx   = body.Children.IndexOf(dragged)
            dst_idx   = body.Children.IndexOf(target)
            body.Children.Remove(dragged)
            insert_at = body.Children.IndexOf(target)
            if src_idx < dst_idx:
                insert_at += 1
            body.Children.Insert(insert_at, dragged)
            names = [c.Tag for c in body.Children if c.Tag]
            write_group_yaml(panel_path, names)

        body.DragOver += on_drag_over
        body.Drop     += on_drop

        # Register so collect_group_orders can read it
        self._group_registry[panel_path] = body

        grip   = self._make_grip()
        header = self._make_collapsible_header(name, body, style_key="Title", grip=grip)

        outer = StackPanel()
        outer.Children.Add(header)
        outer.Children.Add(body)

        card       = Border()
        card.Style = self.window.FindResource("Card")
        card.Child = outer
        card.Tag   = name

        mgr = self

        def on_grip_mouse_move(sender, e):
            if e.LeftButton == System.Windows.Input.MouseButtonState.Pressed:
                mgr._drag_source = card
                System.Windows.DragDrop.DoDragDrop(
                    card,
                    System.Windows.DataObject(mgr.DRAG_FORMAT, card),
                    System.Windows.DragDropEffects.Move
                )
                mgr._drag_source = None

        grip.MouseMove += on_grip_mouse_move

        return card

    def _panel_child_ui(self, item, panel_renamer, drag_format, container, panel_path):
        """
        Wrap a panel-level item (button, pulldown, splitpushbutton, stack) in a
        draggable wrapper with a grip on the left so it can be reordered within the panel.
        """
        if item['type'] == 'button':
            inner = self._tool_ui(item, panel_renamer, standalone=True)
        elif item['type'] == 'pulldown':
            inner = self._pulldown_ui(item, panel_renamer)
        elif item['type'] == 'splitpushbutton':
            inner = self._splitpushbutton_ui(item, panel_renamer)
        elif item['type'] == 'stack':
            inner = self._stack_ui(item, panel_renamer)
        else:
            return None

    def _make_item_grip(self):
        grip                   = TextBlock()
        grip.Text              = u"\u22EE\u22EE"
        grip.FontSize          = 9
        grip.Foreground        = System.Windows.Media.Brushes.White
        grip.Opacity           = 0.6
        grip.Cursor            = System.Windows.Input.Cursors.SizeAll
        grip.VerticalAlignment = System.Windows.VerticalAlignment.Center
        grip.HorizontalAlignment = System.Windows.HorizontalAlignment.Center
        grip.ToolTip           = "Drag to reorder"
        return grip

    def _coloured_bar_wrapper(self, inner, item_name, bar_color, drag_format):
        """
        Wrap `inner` with a coloured left bar that contains the ⋮⋮ grip.
        Returns (wrapper_border, grip) so the caller can wire MouseMove.
        """
        grip = self._make_item_grip()

        # Coloured left tab
        bar                      = Border()
        bar.Width                = 18
        bar.Background           = SolidColorBrush(
            ColorConverter.ConvertFromString(bar_color))
        bar.Child                = grip

        # Grid: col0 = bar, col1 = content
        grid = System.Windows.Controls.Grid()
        col0                   = System.Windows.Controls.ColumnDefinition()
        col0.Width             = System.Windows.GridLength(18)
        col1                   = System.Windows.Controls.ColumnDefinition()
        col1.Width             = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
        grid.ColumnDefinitions.Add(col0)
        grid.ColumnDefinitions.Add(col1)

        System.Windows.Controls.Grid.SetColumn(bar,   0)
        System.Windows.Controls.Grid.SetColumn(inner, 1)
        grid.Children.Add(bar)
        grid.Children.Add(inner)

        wrapper                  = Border()
        wrapper.Tag              = item_name
        wrapper.CornerRadius     = System.Windows.CornerRadius(0, 4, 4, 0)
        wrapper.Margin           = Thickness(0, 3, 4, 3)
        wrapper.ClipToBounds     = True
        wrapper.Child            = grid

        fmt = drag_format
        def on_mouse_move(sender, e):
            if e.LeftButton == System.Windows.Input.MouseButtonState.Pressed:
                System.Windows.DragDrop.DoDragDrop(
                    wrapper,
                    System.Windows.DataObject(fmt, wrapper),
                    System.Windows.DragDropEffects.Move
                )
        grip.MouseMove += on_mouse_move

        return wrapper

    def _panel_child_ui(self, item, panel_renamer, drag_format, container, panel_path):
        """
        Wrap a panel-level item in a coloured-bar card whose left block IS the grip.
        Colour key:
          button        -> muted (#4A5568)
          pulldown      -> green  (#208A3C)
          splitpushbtn  -> green  (#208A3C)
          stack         -> teal   (#2E7D52)
        """
        if item['type'] == 'button':
            inner    = self._tool_ui(item, panel_renamer, standalone=False)
            bar_color = "#4A5568"
        elif item['type'] == 'pulldown':
            inner    = self._pulldown_ui(item, panel_renamer)
            bar_color = "#208A3C"
        elif item['type'] == 'splitpushbutton':
            inner    = self._splitpushbutton_ui(item, panel_renamer)
            bar_color = "#208A3C"
        elif item['type'] == 'stack':
            inner    = self._stack_ui(item, panel_renamer)
            bar_color = "#2E7D52"
        else:
            return None

        return self._coloured_bar_wrapper(inner, item['name'], bar_color, drag_format)

    # ── Child drag container ──────────────────────────────────────────────────

    def _make_child_container(self, group_path, children, renamer):
        """
        Build a droppable StackPanel for button children inside a group.
        Each row gets a small grip. Dropping saves the new order to the
        group's bundle.yaml (creating it if needed).
        """
        CHILD_FORMAT = "Seed43ChildRow_" + group_path  # unique per group

        sp           = StackPanel()
        sp.AllowDrop = True

        # Apply saved order first
        ordered = apply_group_order(children, group_path)

        for child in ordered:
            row = self._tool_row_with_grip(child, renamer, CHILD_FORMAT, sp, group_path)
            sp.Children.Add(row)

        def on_drag_over(sender, e):
            if e.Data.GetDataPresent(CHILD_FORMAT):
                e.Effects = System.Windows.DragDropEffects.Move
            else:
                e.Effects = System.Windows.DragDropEffects.None
            e.Handled = True

        def on_drop(sender, e):
            if not e.Data.GetDataPresent(CHILD_FORMAT):
                return
            dragged = e.Data.GetData(CHILD_FORMAT)
            pos     = e.GetPosition(sp)
            target  = None
            for child in list(sp.Children):
                pt = child.TranslatePoint(System.Windows.Point(0, 0), sp)
                if pos.Y < pt.Y + child.ActualHeight:
                    target = child
                    break
            if target is None or target is dragged:
                return
            src_idx = sp.Children.IndexOf(dragged)
            dst_idx = sp.Children.IndexOf(target)
            sp.Children.Remove(dragged)
            insert_at = sp.Children.IndexOf(target)
            if src_idx < dst_idx:
                insert_at += 1
            sp.Children.Insert(insert_at, dragged)
            # Persist
            names = [c.Tag for c in sp.Children if c.Tag]
            write_group_yaml(group_path, names)

        sp.DragOver += on_drag_over
        sp.Drop     += on_drop

        # Register for collect_group_orders
        self._group_registry[group_path] = sp

        return sp

    def _tool_row_with_grip(self, item, renamer, drag_format, container, group_path):
        """A tool toggle row with the grip inside a coloured left bar."""
        path  = item['path']
        name  = item['name']
        is_on = not path.lower().endswith('.off')

        label                   = TextBlock()
        label.Text              = name
        label.Style             = self.window.FindResource("ToolText")
        label.VerticalAlignment = System.Windows.VerticalAlignment.Center

        switch              = Border()
        switch.Width        = 40
        switch.Height       = 20
        switch.CornerRadius = System.Windows.CornerRadius(10)
        switch.Cursor       = System.Windows.Input.Cursors.Hand
        switch.Background   = SolidColorBrush(
            ColorConverter.ConvertFromString(
                FolderHandler.ON_COLOR if is_on else FolderHandler.OFF_COLOR))

        knob                     = Border()
        knob.Width               = 16
        knob.Height              = 16
        knob.CornerRadius        = System.Windows.CornerRadius(8)
        knob.Background          = System.Windows.Media.Brushes.White
        knob.HorizontalAlignment = System.Windows.HorizontalAlignment.Left
        knob.Margin              = Thickness(22, 2, 0, 2) if is_on else Thickness(2, 2, 0, 2)
        switch.Child             = knob

        handler        = FolderHandler(self.window, path, renamer)
        handler.switch = switch
        handler.knob   = knob
        switch.MouseLeftButtonUp += handler.toggle
        renamer.handlers.append(handler)

        content        = DockPanel()
        content.Margin = Thickness(6, 4, 6, 4)
        DockPanel.SetDock(switch, System.Windows.Controls.Dock.Right)
        content.Children.Add(switch)
        content.Children.Add(label)

        # Grip lives in the coloured left bar
        grip = self._make_item_grip()

        bar            = Border()
        bar.Width      = 18
        bar.Background = SolidColorBrush(
            ColorConverter.ConvertFromString("#4A5568"))
        bar.Child      = grip

        grid = System.Windows.Controls.Grid()
        col0       = System.Windows.Controls.ColumnDefinition()
        col0.Width = System.Windows.GridLength(18)
        col1       = System.Windows.Controls.ColumnDefinition()
        col1.Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
        grid.ColumnDefinitions.Add(col0)
        grid.ColumnDefinitions.Add(col1)
        System.Windows.Controls.Grid.SetColumn(bar,     0)
        System.Windows.Controls.Grid.SetColumn(content, 1)
        grid.Children.Add(bar)
        grid.Children.Add(content)

        row              = Border()
        row.Tag          = name
        row.CornerRadius = System.Windows.CornerRadius(0, 4, 4, 0)
        row.Margin       = Thickness(0, 2, 4, 2)
        row.ClipToBounds = True
        row.Child        = grid

        fmt = drag_format
        def on_mouse_move(sender, e):
            if e.LeftButton == System.Windows.Input.MouseButtonState.Pressed:
                System.Windows.DragDrop.DoDragDrop(
                    row,
                    System.Windows.DataObject(fmt, row),
                    System.Windows.DragDropEffects.Move
                )
        grip.MouseMove += on_mouse_move

        return row

    # ── Pulldown group ────────────────────────────────────────────────────────

    def _pulldown_ui(self, item, panel_renamer):
        pulldown_renamer = FolderRenamer(item['path'], parent=panel_renamer)
        body = self._make_child_container(item['path'], item['children'], pulldown_renamer)
        header = self._make_collapsible_header(
            item['name'], body,
            style_key="PulldownHeader", arrow_style_key="PulldownHeader"
        )
        inner = StackPanel()
        inner.Children.Add(header)
        inner.Children.Add(body)
        card       = Border()
        card.Style = self.window.FindResource('PulldownCard')
        card.Child = inner
        return card

    def _splitpushbutton_ui(self, item, panel_renamer):
        split_renamer = FolderRenamer(item['path'], parent=panel_renamer)
        body = self._make_child_container(item['path'], item['children'], split_renamer)
        header = self._make_collapsible_header(
            item['name'], body,
            style_key="PulldownHeader", arrow_style_key="PulldownHeader"
        )
        tag_lbl             = TextBlock()
        tag_lbl.Text        = u"SPLIT"
        tag_lbl.FontSize    = 9
        tag_lbl.Foreground  = System.Windows.Media.Brushes.Gray
        tag_lbl.Margin      = Thickness(8, 0, 0, 0)
        tag_lbl.VerticalAlignment = System.Windows.VerticalAlignment.Center
        try:
            dock = header.Child
            dock.Children.Add(tag_lbl)
        except Exception:
            pass
        inner = StackPanel()
        inner.Children.Add(header)
        inner.Children.Add(body)
        card       = Border()
        card.Style = self.window.FindResource('PulldownCard')
        card.Child = inner
        return card

    def _stack_ui(self, item, panel_renamer):
        stack_renamer = FolderRenamer(item['path'], parent=panel_renamer)
        STACK_FORMAT  = "Seed43StackRow_" + item['path']
        stack_path    = item['path']

        body           = StackPanel()
        body.AllowDrop = True

        # Apply saved order
        ordered_children = apply_group_order(item['children'], stack_path)

        for child in ordered_children:
            if child['type'] == 'button':
                inner     = self._tool_ui(child, stack_renamer, standalone=False)
                bar_color = "#4A5568"
            elif child['type'] == 'pulldown':
                inner     = self._pulldown_ui(child, stack_renamer)
                bar_color = "#208A3C"
            elif child['type'] == 'splitpushbutton':
                inner     = self._splitpushbutton_ui(child, stack_renamer)
                bar_color = "#208A3C"
            else:
                continue
            wrapped = self._coloured_bar_wrapper(inner, child['name'], bar_color, STACK_FORMAT)
            body.Children.Add(wrapped)

        def on_drag_over(sender, e):
            if e.Data.GetDataPresent(STACK_FORMAT):
                e.Effects = System.Windows.DragDropEffects.Move
            else:
                e.Effects = System.Windows.DragDropEffects.None
            e.Handled = True

        def on_drop(sender, e):
            if not e.Data.GetDataPresent(STACK_FORMAT):
                return
            dragged = e.Data.GetData(STACK_FORMAT)
            pos     = e.GetPosition(body)
            target  = None
            for child in list(body.Children):
                pt = child.TranslatePoint(System.Windows.Point(0, 0), body)
                if pos.Y < pt.Y + child.ActualHeight:
                    target = child
                    break
            if target is None or target is dragged:
                return
            src_idx   = body.Children.IndexOf(dragged)
            dst_idx   = body.Children.IndexOf(target)
            body.Children.Remove(dragged)
            insert_at = body.Children.IndexOf(target)
            if src_idx < dst_idx:
                insert_at += 1
            body.Children.Insert(insert_at, dragged)
            names = [c.Tag for c in body.Children if c.Tag]
            write_group_yaml(stack_path, names)

        body.DragOver += on_drag_over
        body.Drop     += on_drop

        # Register for collect_group_orders
        self._group_registry[stack_path] = body

        header = self._make_collapsible_header(
            item['name'], body,
            style_key="PulldownHeader", arrow_style_key="PulldownHeader"
        )
        tag_lbl                   = TextBlock()
        tag_lbl.Text              = u"STACK"
        tag_lbl.FontSize          = 9
        tag_lbl.Foreground        = System.Windows.Media.Brushes.Gray
        tag_lbl.Margin            = Thickness(8, 0, 0, 0)
        tag_lbl.VerticalAlignment = System.Windows.VerticalAlignment.Center
        try:
            header.Child.Children.Add(tag_lbl)
        except Exception:
            pass

        inner_sp = StackPanel()
        inner_sp.Children.Add(header)
        inner_sp.Children.Add(body)

        card                 = Border()
        card.Background      = SolidColorBrush(
            ColorConverter.ConvertFromString("#1E2733"))
        card.BorderBrush     = SolidColorBrush(
            ColorConverter.ConvertFromString("#2E7D52"))
        card.BorderThickness = Thickness(2, 0, 0, 0)
        card.CornerRadius    = System.Windows.CornerRadius(0, 4, 4, 0)
        card.Margin          = Thickness(12, 3, 4, 3)
        card.Child           = inner_sp
        return card

    # ── Individual tool row ───────────────────────────────────────────────────

    def _tool_ui(self, item, renamer, standalone=False):
        path  = item['path']
        name  = item['name']
        is_on = not path.lower().endswith('.off')

        label                   = TextBlock()
        label.Text              = name
        label.Style             = self.window.FindResource("ToolText")
        label.VerticalAlignment = System.Windows.VerticalAlignment.Center

        switch              = Border()
        switch.Width        = 40
        switch.Height       = 20
        switch.CornerRadius = System.Windows.CornerRadius(10)
        switch.Cursor       = System.Windows.Input.Cursors.Hand
        switch.Background   = SolidColorBrush(
            ColorConverter.ConvertFromString(
                FolderHandler.ON_COLOR if is_on else FolderHandler.OFF_COLOR))

        knob                     = Border()
        knob.Width               = 16
        knob.Height              = 16
        knob.CornerRadius        = System.Windows.CornerRadius(8)
        knob.Background          = System.Windows.Media.Brushes.White
        knob.HorizontalAlignment = System.Windows.HorizontalAlignment.Left
        knob.Margin              = Thickness(22, 2, 0, 2) if is_on else Thickness(2, 2, 0, 2)
        switch.Child             = knob

        handler        = FolderHandler(self.window, path, renamer)
        handler.switch = switch
        handler.knob   = knob
        switch.MouseLeftButtonUp += handler.toggle
        renamer.handlers.append(handler)

        row        = DockPanel()
        row.Margin = Thickness(6, 4, 6, 4)
        DockPanel.SetDock(switch, System.Windows.Controls.Dock.Right)
        row.Children.Add(switch)
        row.Children.Add(label)

        if standalone:
            card       = Border()
            card.Style = self.window.FindResource("StandaloneCard")
            card.Child = row
            return card
        return row

# ── Main dialog ───────────────────────────────────────────────────────────────

class Seed43Dialog(object):

    def __init__(self):
        self.window = load_xaml(XAML_PATH)

        # ── Load icon ─────────────────────────────────────────────────────────
        if os.path.exists(ICON_PATH):
            img       = self.window.FindName("header_icon")
            bmp       = BitmapImage()
            bmp.BeginInit()
            bmp.UriSource = Uri(ICON_PATH, UriKind.Absolute)
            bmp.EndInit()
            img.Source = bmp

        self._bind()
        self._init_tools()
        self._check_versions()

    def _bind(self):
        self.window.FindName("footer_reload").Click             += self._on_reload
        self.window.FindName("update_ribbon").MouseLeftButtonUp += self._on_s43_update
        self.window.FindName("apply_order_btn").Click           += self._on_apply_order

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
            local   = read_local_version()
            notes   = read_last_update()
            remote  = fetch_remote_version()
            dispatch(self.window, lambda: self._update_s43_ui(local, notes, remote))
        t = Thread(ThreadStart(worker))
        t.IsBackground = True
        t.Start()

    def _update_s43_ui(self, local, notes, remote):
        # ── Set version as card title ─────────────────────────────────────────
        self.window.FindName("s43_title").Text = u"\u25CF  Installed  v{}".format(local) if local else "Version unknown"

        # ── Show Last update notes from version.txt ───────────────────────────
        if notes:
            self.window.FindName("s43_changelog").Text = "\n".join(notes)
        else:
            self.window.FindName("s43_changelog").Text = ""

        # ── Show update ribbon if newer version available on GitHub ───────────
        if remote and version_tuple(remote) > version_tuple(local):
            self._remote_s43_version = remote
            self.window.FindName("update_ribbon_version").Text = \
                u"v{}  \u2192  v{}".format(local, remote)
            self.window.FindName("update_ribbon").Visibility = Visibility.Visible
        elif not remote:
            self.window.FindName("s43_changelog").Text = (
                "\n".join(notes) + "\n\nCould not reach GitHub to check for updates."
                if notes else "Could not reach GitHub to check for updates."
            )

    def _on_s43_update(self, sender, args):
        result = MessageBox.Show(
            "Update Seed43 extension to v{0}?\n\nThe extension will be re-downloaded from GitHub.\nReload PyRevit in Revit after updating.".format(
                getattr(self, "_remote_s43_version", "latest")),
            "Update Seed43",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question
        )
        if str(result) != "Yes":
            return

        self.window.FindName("update_ribbon").Visibility = Visibility.Collapsed

        EXTENSIONS_DIR = os.path.join(os.environ.get("APPDATA", ""), "pyRevit", "Extensions")
        S43_INSTALL    = os.path.join(EXTENSIONS_DIR, "Seed43.extension")
        TAB_DIR_DEST   = os.path.join(S43_INSTALL, "Seed43.tab")
        TEMP_ZIP       = os.path.join(os.environ.get("TEMP", ""), "seed43_update.zip")
        TEMP_DIR       = os.path.join(os.environ.get("TEMP", ""), "seed43_update_extracted")

        # File extensions to skip during sync, preserving local config files
        SKIP_EXTENSIONS = (".yaml", ".json")

        def log(msg):
            dispatch(self.window, lambda: setattr(
                self.window.FindName("s43_changelog"), "Text", msg))

        def worker():
            try:
                log("Downloading update...")
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

                log("Installing update...")
                new_tab = os.path.join(extracted_root, "Seed43.tab")
                if not os.path.exists(new_tab):
                    raise Exception("Seed43.tab not found in download.")

                # ── Replace Seed43.tab, skipping yaml and json files ──────────
                if os.path.isdir(TAB_DIR_DEST):
                    shutil.rmtree(TAB_DIR_DEST)
                shutil.copytree(
                    new_tab,
                    TAB_DIR_DEST,
                    ignore=shutil.ignore_patterns(*["*" + ext for ext in SKIP_EXTENSIONS])
                )

                # ── Update local version.txt ──────────────────────────────────
                new_version_file = os.path.join(extracted_root, "version.txt")
                if os.path.isfile(new_version_file):
                    shutil.copy2(new_version_file, VERSION_FILE)

                version = fetch_remote_version() or "unknown"

                log("Done, v{0}".format(version))
                if os.path.exists(TEMP_ZIP):
                    os.remove(TEMP_ZIP)
                if os.path.exists(TEMP_DIR):
                    shutil.rmtree(TEMP_DIR)

                dispatch(self.window, lambda: self._on_s43_update_done(version))

            except Exception as ex:
                dispatch(self.window, lambda: self._on_error(str(ex)))

        t = Thread(ThreadStart(worker))
        t.IsBackground = True
        t.Start()

    def _on_s43_update_done(self, version):
        self._local_s43_version = version
        self.window.FindName("update_ribbon").Visibility = Visibility.Collapsed
        self.window.FindName("s43_version").Text = u"\u25CF  Installed  v{0}".format(version)
        self.window.FindName("s43_changelog").Text = u"Updated to v{0}, reloading PyRevit...".format(version)
        MessageBox.Show(
            "Seed43 updated to v{0}.\n\nPyRevit will now reload to apply the update.".format(version),
            "Seed43 Updated",
            MessageBoxButton.OK,
            MessageBoxImage.Information
        )
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

    def _on_error(self, msg):
        MessageBox.Show("Operation failed:\n\n" + msg, "Seed43",
                        MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_apply_order(self, sender, args):
        container = self.window.FindName("tools_container")

        # ── Collect panel order ───────────────────────────────────────────────
        panel_names = [c.Tag for c in container.Children if c.Tag]

        # ── Walk the UI tree to collect every group's current child order ─────
        # We stored the path→children mapping in ToolManager during build.
        # Re-read it from the live UI by traversing the tree using the
        # _group_paths registry we'll maintain on ToolManager.
        group_orders = {}
        if hasattr(self, '_tool_manager') and self._tool_manager:
            group_orders = self._tool_manager.collect_group_orders()

        # ── Save JSON (single source of truth) ───────────────────────────────
        data = {"panels": panel_names, "groups": group_orders}
        save_order_data(data)

        # ── Write every bundle.yaml from the JSON data ────────────────────────
        errors = []

        # Tab-level bundle.yaml (panel order, Seed43 pinned first)
        if not write_bundle_yaml(panel_names):
            errors.append("tab bundle.yaml")

        # Each group's bundle.yaml
        for path, names in group_orders.items():
            if not write_group_yaml(path, names):
                errors.append(os.path.basename(path))

        if not errors:
            MessageBox.Show(
                "Panel order saved to bundle.yaml.\n\nThe new ribbon order will take effect after Revit restarts.",
                "Order Applied",
                MessageBoxButton.OK,
                MessageBoxImage.Information
            )
        else:
            MessageBox.Show(
                "Some files could not be written:\n\n" + "\n".join(errors) +
                "\n\nCheck that the files are not read-only.",
                "Partial Write",
                MessageBoxButton.OK,
                MessageBoxImage.Warning
            )

    def show(self):
        self.window.ShowDialog()

# ── Entry point ───────────────────────────────────────────────────────────────
dialog = Seed43Dialog()
dialog.show()
