# -*- coding: utf-8 -*-
import os
import clr
import shutil
import zipfile

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("System.Net")

from System.IO import File, Directory, StreamReader
from System.Net import WebClient
from System.Windows.Markup import XamlReader
from System.Windows import Visibility
from System.Windows.Media.Imaging import BitmapImage
from System.Windows.Threading import Dispatcher
from System.Threading import Thread, ThreadStart
from System import Uri, UriKind, Action

# NOTE: pyrevit.forms is intentionally NOT imported here. It is a heavy module
# (2 to 4 seconds to import) and was the entire cause of slow Seed43 startup.
# It is imported lazily below, only on the rare path where an update is applied.


# ── XAML ─────────────────────────────────────────────────────────────────────

WINDOW_XAML = """
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Seed43 Update"
    Width="420"
    SizeToContent="Height"
    ResizeMode="NoResize"
    WindowStartupLocation="CenterScreen"
    Background="{StaticResource LocalWindowBg}"
    TextElement.FontFamily="Segoe UI">

    <Window.Resources>

        <SolidColorBrush x:Key="LocalWindowBg"    Color="#3B4553"/>
        <SolidColorBrush x:Key="LocalHeaderBg"    Color="#232933"/>
        <SolidColorBrush x:Key="LocalAccent"        Color="#208A3C"/>
        <SolidColorBrush x:Key="LocalAccentHover"   Color="#2B933F"/>
        <SolidColorBrush x:Key="LocalAccentPressed" Color="#1A6E2E"/>
        <SolidColorBrush x:Key="LocalLightText"   Color="#F4FAFF"/>
        <SolidColorBrush x:Key="LocalCardBg"      Color="#F4FAFF"/>
        <SolidColorBrush x:Key="LocalCardText"    Color="#2B3340"/>
        <SolidColorBrush x:Key="LocalSecondaryBg"        Color="#404553"/>
        <SolidColorBrush x:Key="LocalSecondaryHover"     Color="#4E5566"/>
        <SolidColorBrush x:Key="LocalSecondaryPressed"   Color="#333B48"/>
        <SolidColorBrush x:Key="LocalDisabledBg"  Color="#555555"/>
        <SolidColorBrush x:Key="LocalDisabledFg"  Color="#888888"/>
        <SolidColorBrush x:Key="LocalDangerHover"   Color="#FF5555"/>
        <SolidColorBrush x:Key="LocalDangerPressed" Color="#CC4444"/>
        <SolidColorBrush x:Key="LocalProgressTrackBg" Color="#D0D8E0"/>

        <DropShadowEffect x:Key="HeaderShadow"
                          Color="Black" Opacity="0.3"
                          ShadowDepth="2" BlurRadius="6"/>

        <DropShadowEffect x:Key="CardShadow"
                          Color="Black" Opacity="0.2"
                          ShadowDepth="2" BlurRadius="4"/>

        <Style x:Key="SmallButtonStyle" TargetType="Button">
            <Setter Property="Background"      Value="{StaticResource LocalAccent}"/>
            <Setter Property="Foreground"      Value="{StaticResource LocalLightText}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding"         Value="12,4"/>
            <Setter Property="FontSize"        Value="11"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Height"          Value="24"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Bd"
                                Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{StaticResource LocalAccentHover}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{StaticResource LocalAccentPressed}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Bd" Property="Background" Value="{StaticResource LocalDisabledBg}"/>
                                <Setter Property="Foreground"                  Value="{StaticResource LocalDisabledFg}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="SecondaryButtonStyle" TargetType="Button">
            <Setter Property="Background"      Value="{StaticResource LocalSecondaryBg}"/>
            <Setter Property="Foreground"      Value="{StaticResource LocalLightText}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding"         Value="16,8"/>
            <Setter Property="FontSize"        Value="12"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Bd"
                                Background="{TemplateBinding Background}"
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{StaticResource LocalSecondaryHover}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{StaticResource LocalSecondaryPressed}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Bd" Property="Background" Value="{StaticResource LocalDisabledBg}"/>
                                <Setter Property="Foreground"                  Value="{StaticResource LocalDisabledFg}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="CloseButtonStyle" TargetType="Button">
            <Setter Property="Background"      Value="Transparent"/>
            <Setter Property="Foreground"      Value="{StaticResource LocalLightText}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontSize"        Value="14"/>
            <Setter Property="Width"           Value="30"/>
            <Setter Property="Height"          Value="30"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Bd"
                                Background="{TemplateBinding Background}"
                                CornerRadius="15">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{StaticResource LocalDangerHover}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{StaticResource LocalDangerPressed}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

    </Window.Resources>

    <Grid>

        <!-- Header -->
        <Border Height="70"
                VerticalAlignment="Top"
                Background="{StaticResource LocalHeaderBg}"
                CornerRadius="0,0,12,12"
                Effect="{StaticResource HeaderShadow}"
                Panel.ZIndex="10">
            <Grid Margin="24,0,24,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0"
                            Orientation="Horizontal"
                            VerticalAlignment="Center">
                    <Image x:Name="header_icon"
                           Width="32" Height="32"
                           VerticalAlignment="Center"
                           Margin="0,0,10,0"/>
                    <TextBlock FontSize="32" FontWeight="Bold"
                               Foreground="{StaticResource LocalAccent}" Text="SEED"/>
                    <TextBlock FontSize="32" FontWeight="SemiBold"
                               Foreground="{StaticResource LocalLightText}" Text="43"/>
                    <TextBlock FontSize="20" FontWeight="SemiBold"
                               Foreground="{StaticResource LocalLightText}" Opacity="0.75"
                               VerticalAlignment="Bottom"
                               Margin="10,0,0,5"
                               Text="  |  Update"/>
                </StackPanel>
                <Button x:Name="header_close_btn"
                        Grid.Column="1"
                        Style="{StaticResource CloseButtonStyle}"
                        Content="&#x2716;"
                        VerticalAlignment="Center"/>
            </Grid>
        </Border>

        <!-- Update card -->
        <Border Margin="24,90,24,24"
                Background="{StaticResource LocalCardBg}"
                BorderBrush="{StaticResource LocalAccent}"
                BorderThickness="1"
                CornerRadius="6"
                Padding="24"
                Effect="{StaticResource CardShadow}">
            <StackPanel>

                <TextBlock x:Name="update_title_lbl"
                           Text="Update Available"
                           Foreground="{StaticResource LocalAccent}"
                           FontSize="14"
                           FontWeight="SemiBold"
                           Margin="0,0,0,8"/>

                <TextBlock x:Name="update_msg_lbl"
                           Foreground="{StaticResource LocalCardText}"
                           FontSize="12"
                           TextWrapping="Wrap"
                           Margin="0,0,0,20"/>

                <ProgressBar x:Name="update_progress"
                             Height="4"
                             Margin="0,0,0,12"
                             Minimum="0"
                             Maximum="100"
                             Value="0"
                             Foreground="{StaticResource LocalAccent}"
                             Background="{StaticResource LocalProgressTrackBg}"
                             BorderThickness="0"
                             Visibility="Collapsed"/>

                <TextBlock x:Name="status_lbl"
                           Foreground="{StaticResource LocalAccent}"
                           FontSize="11"
                           Margin="0,0,0,12"
                           Visibility="Collapsed"/>

                <StackPanel Orientation="Horizontal"
                            HorizontalAlignment="Right">
                    <Button x:Name="skip_btn"
                            Content="Not Now"
                            Style="{StaticResource SecondaryButtonStyle}"
                            Padding="12,4"
                            Height="24"
                            FontSize="11"
                            Margin="0,0,8,0"/>
                    <Button x:Name="update_btn"
                            Content="Update Now"
                            Style="{StaticResource SmallButtonStyle}"
                            Width="100"/>
                </StackPanel>

            </StackPanel>
        </Border>

    </Grid>
</Window>
"""


# ── VARIABLES ─────────────────────────────────────────────────────────────────

GITHUB_USER   = "Seed-43"
GITHUB_REPO   = "Seed43"
GITHUB_BRANCH = "main"

VERSION_URL  = "https://raw.githubusercontent.com/{}/{}/{}/version.txt".format(
    GITHUB_USER, GITHUB_REPO, GITHUB_BRANCH)

REPO_ZIP_URL = "https://github.com/{}/{}/archive/refs/heads/{}.zip".format(
    GITHUB_USER, GITHUB_REPO, GITHUB_BRANCH)

SCRIPT_DIR    = os.path.dirname(__file__)
APPDATA       = os.environ.get("APPDATA", "")
EXTENSION_DIR = os.path.join(APPDATA, "pyRevit", "Extensions", "Seed43.extension")
TAB_DIR       = os.path.join(EXTENSION_DIR, "Seed43.tab")
VERSION_FILE  = os.path.join(EXTENSION_DIR, "version.txt")
ICON_PATH     = os.path.join(SCRIPT_DIR, "icon.png")


# ── FUNCTIONS ─────────────────────────────────────────────────────────────────

def read_local_version():
    """Read the installed version string from the extension root version.txt.
    Returns only the first non-empty line as the version number."""
    try:
        if not File.Exists(VERSION_FILE):
            return "0.0.0"
        reader  = StreamReader(VERSION_FILE)
        content = reader.ReadToEnd()
        reader.Close()
        for line in content.splitlines():
            line = line.strip()
            if line:
                return line
    except Exception:
        pass
    return "0.0.0"


def fetch_remote_version():
    """Download version.txt from GitHub and return only the version number.
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


def sync_tree(src, dst):
    """Sync src into dst file by file.
    - Copies all files from src, skipping .json files
    - Deletes files in dst that are not in src, but keeps .json files
    - Removes dirs in dst not in src only if they contain no .json files
    - Recurses into all subdirs
    """
    if not os.path.isdir(dst):
        os.makedirs(dst)

    src_names = set(os.listdir(src))
    dst_names = set(os.listdir(dst))

    # Copy/overwrite everything from src except .json
    for name in src_names:
        s = os.path.join(src, name)
        d = os.path.join(dst, name)
        if os.path.isdir(s):
            sync_tree(s, d)
        elif not name.lower().endswith(".json"):
            shutil.copy2(s, d)

    # Delete dst files/dirs not in src, but preserve .json files
    for name in dst_names:
        if name not in src_names:
            d = os.path.join(dst, name)
            if os.path.isfile(d):
                if not name.lower().endswith(".json"):
                    os.remove(d)
            elif os.path.isdir(d):
                has_json = any(
                    f.lower().endswith(".json")
                    for _, _, files in os.walk(d)
                    for f in files
                )
                if not has_json:
                    shutil.rmtree(d)


def read_migrations(extracted_root):
    """Parse seed43_migrations.yaml from the extracted repo root.
    Returns a list of migration dicts with from/to/subfolders keys."""
    yaml_path = os.path.join(extracted_root, "seed43_migrations.yaml")
    if not os.path.exists(yaml_path):
        return []
    migrations = []
    current    = None
    in_sub     = False
    sub_from   = None
    try:
        with open(yaml_path, "r") as f:
            lines = f.readlines()
        for line in lines:
            stripped = line.strip()
            indent   = len(line) - len(line.lstrip())
            if not stripped or stripped.startswith("#"):
                continue
            if stripped.startswith("migrations:"):
                continue
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
    """Copy .json files from old folder paths to new folder paths before syncing."""
    for m in migrations:
        old_dir = os.path.join(tab_dst, m['from'])
        new_dir = os.path.join(tab_dst, m['to'])
        if not os.path.isdir(old_dir):
            continue
        if not os.path.isdir(new_dir):
            os.makedirs(new_dir)
        for fname in os.listdir(old_dir):
            if fname.lower().endswith(".json"):
                src = os.path.join(old_dir, fname)
                dst = os.path.join(new_dir, fname)
                if not os.path.exists(dst):
                    shutil.copy2(src, dst)
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


def download_and_apply_update(status_lbl, progress_bar):
    """Download the repo zip, sync the new Seed43.tab into place, update version.txt.
    Preserves local .json config files. Copies .yaml files from the new version.
    Returns True on success, False on failure."""
    tmp_zip = os.path.join(EXTENSION_DIR, "_seed43_update.zip")
    tmp_dir = os.path.join(EXTENSION_DIR, "_seed43_update_tmp")

    try:
        # ── Download ──────────────────────────────────────────────────────────

        def on_progress(sender, e):
            progress_bar.Visibility = Visibility.Visible
            progress_bar.Value      = e.ProgressPercentage

        status_lbl.Visibility = Visibility.Visible
        status_lbl.Text       = "Downloading update..."

        client = WebClient()
        client.DownloadProgressChanged += on_progress
        client.DownloadFile(REPO_ZIP_URL, tmp_zip)

        # ── Extract ───────────────────────────────────────────────────────────

        status_lbl.Text = "Extracting files..."

        if Directory.Exists(tmp_dir):
            shutil.rmtree(tmp_dir)
        Directory.CreateDirectory(tmp_dir)

        with zipfile.ZipFile(tmp_zip, "r") as zf:
            zf.extractall(tmp_dir)

        extracted_root = os.path.join(tmp_dir, "{}-{}".format(GITHUB_REPO, GITHUB_BRANCH))
        new_tab        = os.path.join(extracted_root, "Seed43.tab")

        if not os.path.isdir(new_tab):
            status_lbl.Text = "Update failed: Seed43.tab not found in download."
            return False

        # ── Sync Seed43.tab, keeping json, updating yaml ──────────────────────

        status_lbl.Text = "Checking migrations..."
        migrations = read_migrations(extracted_root)
        if migrations:
            apply_migrations(migrations, TAB_DIR)

        status_lbl.Text = "Applying update..."
        sync_tree(new_tab, TAB_DIR)

        # ── Sync root files (startup.py, extension.json, etc.) ────────────────
        ROOT_SKIP = {
            "Seed43.tab", "lib", "UI",
            ".git", ".gitignore", "README.md", "LICENSE",
            "install.bat", "sync-start.bat", "sync-end.bat",
        }
        for fname in os.listdir(extracted_root):
            if fname in ROOT_SKIP:
                continue
            src = os.path.join(extracted_root, fname)
            dst = os.path.join(EXTENSION_DIR, fname)
            if os.path.isfile(src):
                if not fname.lower().endswith(".json") and not fname.lower().endswith(".png"):
                    shutil.copy2(src, dst)

        # ── Update local version.txt ──────────────────────────────────────────

        new_version_file = os.path.join(extracted_root, "version.txt")
        if os.path.isfile(new_version_file):
            shutil.copy2(new_version_file, VERSION_FILE)

        # ── Cleanup ───────────────────────────────────────────────────────────

        shutil.rmtree(tmp_dir, ignore_errors=True)
        if os.path.isfile(tmp_zip):
            os.remove(tmp_zip)

        progress_bar.Value = 100
        status_lbl.Text    = "Update complete. Please restart Revit."
        return True

    except Exception as ex:
        status_lbl.Visibility = Visibility.Visible
        status_lbl.Text       = "Update failed: {}".format(str(ex))
        shutil.rmtree(tmp_dir, ignore_errors=True)
        try:
            if os.path.isfile(tmp_zip):
                os.remove(tmp_zip)
        except Exception:
            pass
        return False


# ── WINDOW CONTROLLER ─────────────────────────────────────────────────────────

class UpdateWindow(object):

    def __init__(self, local_version, remote_version):
        self.window = XamlReader.Parse(WINDOW_XAML)

        # ── Load icon ─────────────────────────────────────────────────────────
        if os.path.exists(ICON_PATH):
            img           = self.window.FindName("header_icon")
            bmp           = BitmapImage()
            bmp.BeginInit()
            bmp.UriSource = Uri(ICON_PATH, UriKind.Absolute)
            bmp.EndInit()
            img.Source    = bmp

        self.title_lbl    = self.window.FindName("update_title_lbl")
        self.msg_lbl      = self.window.FindName("update_msg_lbl")
        self.progress_bar = self.window.FindName("update_progress")
        self.status_lbl   = self.window.FindName("status_lbl")
        self.skip_btn     = self.window.FindName("skip_btn")
        self.update_btn   = self.window.FindName("update_btn")
        self.close_btn    = self.window.FindName("header_close_btn")

        self._updated = False

        self.msg_lbl.Text = (
            "A new version of Seed43 is available.\n\n"
            "Your version:    {}\n"
            "Latest version:  {}\n\n"
            "Would you like to update now?"
        ).format(local_version, remote_version)

        self._bind_events()

    def _bind_events(self):
        self.skip_btn.Click   += self._on_skip
        self.close_btn.Click  += self._on_skip
        self.update_btn.Click += self._on_update

    def _on_skip(self, sender, e):
        self.window.Close()

    def _on_update(self, sender, e):
        self.update_btn.IsEnabled = False
        self.skip_btn.IsEnabled   = False

        success       = download_and_apply_update(self.status_lbl, self.progress_bar)
        self._updated = success

        if success:
            self.title_lbl.Text       = "Update Complete"
            self.update_btn.Content   = "Done"
            self.update_btn.IsEnabled = True
            self.update_btn.Click    -= self._on_update
            self.update_btn.Click    += self._on_skip
        else:
            self.skip_btn.IsEnabled   = True
            self.update_btn.IsEnabled = True

    def show(self):
        self.window.ShowDialog()
        return self._updated


# ── MAIN ──────────────────────────────────────────────────────────────────────

def _check_and_notify(ui_dispatcher):
    """Run silently in the background, only show the update window if needed."""
    try:
        local_version  = read_local_version()
        remote_version = fetch_remote_version()

        # ── No connection or no update needed, do nothing ─────────────────────
        if remote_version is None:
            return
        if version_tuple(remote_version) <= version_tuple(local_version):
            return

        # ── Update available, show the window on the UI thread ────────────────
        def show_window():
            window  = UpdateWindow(local_version, remote_version)
            updated = window.show()
            if updated:
                from pyrevit import forms
                forms.alert(
                    "Update applied successfully.\n\n"
                    "Please restart Revit for the changes to take effect.",
                    title="Seed43 Update"
                )

        ui_dispatcher.Invoke(Action(show_window))

    except Exception:
        pass


# ── SCHEDULED PRINT (pySheets, works even if pySheets isn't open) ─────────────
#
# pySheets' own in-window scheduler (a DispatcherTimer) only runs while its
# window is open, closing the window kills it. This registers a session-level
# Application.Idling handler instead, so an armed schedule survives the
# window closing, as long as Revit and the target document stay open.
#
# V1 only: if the target document isn't open when the schedule comes due, it
# is skipped rather than opened automatically. Auto-opening the document is a
# planned V2 addition, not built yet.

def _find_pysheets_dir():
    """Locate PySheets.pushbutton under Seed43.tab. Returns the folder
    path, or None if it can't be found (extension reorganized, tool
    removed, etc, in which case the scheduler just does nothing)."""
    try:
        for root, _dirs, _files in os.walk(TAB_DIR):
            if os.path.basename(root) == "PySheets.pushbutton":
                return root
    except Exception:
        pass
    return None


_PYSHEETS_DIR = _find_pysheets_dir()
_SCHEDULE_FILE = (
    os.path.join(_PYSHEETS_DIR, "userdata", "settings", "scheduled_print.json")
    if _PYSHEETS_DIR else None
)
_LIB_DIR = os.path.join(EXTENSION_DIR, "lib")

_last_schedule_check = [0.0]


def _read_schedule():
    if not _SCHEDULE_FILE:
        return None
    try:
        if not File.Exists(_SCHEDULE_FILE):
            return None
        import json
        reader  = StreamReader(_SCHEDULE_FILE)
        content = reader.ReadToEnd()
        reader.Close()
        return json.loads(content)
    except Exception:
        return None


def _launch_pysheets_schedule(sched):
    """Import pySheets fresh and hand it the due schedule. Imported lazily
    here, not at module load, to keep it out of every Revit startup."""
    import sys as _sys
    if _PYSHEETS_DIR and _PYSHEETS_DIR not in _sys.path:
        _sys.path.insert(0, _PYSHEETS_DIR)
    if _LIB_DIR and os.path.isdir(_LIB_DIR) and _LIB_DIR not in _sys.path:
        _sys.path.insert(0, _LIB_DIR)
    import pySheets
    pySheets.launch_scheduled(sched)


from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
from Autodesk.Revit.DB import ModelPathUtils


class _PySheetsScheduleHandler(IExternalEventHandler):
    """Runs on Revit's own API thread (that's the whole point of
    ExternalEvent), safely deferred by Revit itself until no command is
    active, so this never interrupts something the user is mid-way
    through doing."""

    def Execute(self, uiapp):
        try:
            sched = _read_schedule()
            if not sched or not sched.get("enabled"):
                return
            doc_path = sched.get("document_path")
            if not doc_path:
                return

            target_doc = None
            for d in uiapp.Application.Documents:
                try:
                    if d.PathName and os.path.normcase(d.PathName) == os.path.normcase(doc_path):
                        target_doc = d
                        break
                except Exception:
                    continue
            if target_doc is None:
                return  # V1: not open, skip rather than open it

            try:
                model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(doc_path)
                uiapp.OpenAndActivateDocument(model_path)
            except Exception:
                pass  # already open+active, activation failing here is not fatal

            _launch_pysheets_schedule(sched)
        except Exception:
            pass

    def GetName(self):
        return "Seed43 pySheets Scheduled Print"


_pysheets_schedule_event = ExternalEvent.Create(_PySheetsScheduleHandler())


def _on_idling(sender, args):
    """Cheap periodic check (throttled to roughly every 20 seconds) for
    whether an armed pySheets schedule has come due."""
    try:
        import time
        now = time.time()
        if now - _last_schedule_check[0] < 20:
            return
        _last_schedule_check[0] = now

        sched = _read_schedule()
        if not sched or not sched.get("enabled"):
            return
        next_run_str = sched.get("next_run")
        if not next_run_str:
            return

        from datetime import datetime
        try:
            next_run = datetime.strptime(next_run_str, "%Y-%m-%dT%H:%M:%S")
        except Exception:
            return
        if datetime.now() < next_run:
            return

        _pysheets_schedule_event.Raise()
    except Exception:
        pass


def main():
    ui_dispatcher = Dispatcher.CurrentDispatcher

    def worker():
        _check_and_notify(ui_dispatcher)

    t = Thread(ThreadStart(worker))
    t.IsBackground = True
    t.Start()

    try:
        __revit__.Application.Idling += _on_idling
    except Exception:
        pass


if __name__ == "__main__":
    main()
