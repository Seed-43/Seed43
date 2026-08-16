# -*- coding: utf-8 -*-
import os
import clr
import shutil
import traceback
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
    xmlns:shell="clr-namespace:System.Windows.Shell;assembly=PresentationFramework"
    Title="Seed43 Update"
    Width="420"
    SizeToContent="Height"
    ResizeMode="NoResize"
    WindowStyle="None"
    WindowStartupLocation="CenterScreen"
    Background="{DynamicResource LocalWindowBg}"
    TextElement.FontFamily="Segoe UI">

    <shell:WindowChrome.WindowChrome>
        <shell:WindowChrome CaptionHeight="70"
                             CornerRadius="10"
                             GlassFrameThickness="0"
                             ResizeBorderThickness="0"
                             UseAeroCaptionButtons="False"/>
    </shell:WindowChrome.WindowChrome>

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
                        VerticalAlignment="Center"
                        shell:WindowChrome.IsHitTestVisibleInChrome="True"/>
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
_LOG_FILE     = os.path.join(SCRIPT_DIR, "startup_error.log")


_LOG_RETENTION_DAYS = 28

# The date the log was last pruned in this session. A Revit left open for
# weeks would otherwise never prune, so the check is per-day rather than
# once at startup - but still only one pass a day, not one per write.
_log_pruned_on = [None]


def _prune_log(force=False):
    """Drop log entries older than _LOG_RETENTION_DAYS, at most once a day.

    Both entry shapes lead with their date - "[YYYY-MM-DD ...]" for a note,
    "-- context [YYYY-MM-DD ...] --" for a traceback - so an entry's header
    dates every line under it until the next header.

    Undated entries are treated as older than anything dated and dropped:
    the only ones that exist predate timestamping, so they are by definition
    beyond the window.

    Swallows everything. Pruning the log must never be the reason startup
    fails, and it cannot call _log_error without risking a loop.
    """
    try:
        import datetime as _dt
        today = _dt.date.today()
        if not force and _log_pruned_on[0] == today:
            return
        _log_pruned_on[0] = today

        if not os.path.isfile(_LOG_FILE):
            return

        cutoff = today - _dt.timedelta(days=_LOG_RETENTION_DAYS)

        import re as _re
        header = _re.compile(r"^(?:\[|--\s.*\s\[)(\d{4})-(\d{2})-(\d{2})")

        with open(_LOG_FILE, "r") as handle:
            lines = handle.readlines()

        kept = []
        keeping = False
        for line in lines:
            match = header.match(line)
            if match:
                stamped = _dt.date(int(match.group(1)),
                                   int(match.group(2)),
                                   int(match.group(3)))
                keeping = stamped >= cutoff
            elif line.startswith("--") or line.startswith("["):
                # A header shape that carries no parsable date - legacy entry.
                keeping = False
            if keeping:
                kept.append(line)

        # Only rewrite when something actually aged out, so the common case
        # is a read and nothing more.
        if len(kept) != len(lines):
            with open(_LOG_FILE, "w") as handle:
                handle.writelines(kept)
    except Exception:
        pass


def _log_note(message):
    """A stated condition in the same log as _log_error, without a traceback.

    For the case that isn't an exception but still leaves the scheduler dead:
    the Idling subscription failing outright. Nothing routine is logged here
    - the handler runs every 20 seconds, and a note per pass would bury the
    real errors."""
    try:
        _prune_log()
        import datetime as _dt
        with open(_LOG_FILE, "a") as f:
            f.write("[{}] {}\n".format(
                _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), message))
    except Exception:
        pass


def _log_error(context):
    """Silent-by-design elsewhere in this file, but a swallowed exception
    here means the update popup just never appears with no way to tell
    why - write the traceback to a local file instead of losing it.

    Timestamped like _log_note. Without the stamp there is no way to tell a
    traceback written seconds ago from one left by a Revit session that has
    been running since before the code was last changed - which is exactly
    the question worth answering when the scheduler goes quiet.
    """
    try:
        _prune_log()
        import datetime as _dt
        with open(_LOG_FILE, "a") as f:
            f.write("-- {} [{}] --\n{}\n".format(
                context,
                _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                traceback.format_exc()))
    except Exception:
        pass


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
        _log_error("fetch_remote_version")
        return None


def version_tuple(version_str):
    """Convert a version string like 1.2.3 to a tuple (1, 2, 3) for comparison."""
    try:
        return tuple(int(x) for x in version_str.strip().split("."))
    except Exception:
        return (0, 0, 0)


def sync_tree(src, dst, keep_json=True):
    """Sync src into dst file by file.
    - Copies all files from src, skipping .json files
    - Deletes files in dst that are not in src, but keeps .json files
    - Removes dirs in dst not in src only if they contain no .json files
    - Recurses into all subdirs

    keep_json=False syncs everything including .json. Used for lib/, which is
    shipped code rather than user settings - the .json-is-user-data rule holds
    in Seed43.tab but not there.
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
            sync_tree(s, d, keep_json)
        elif not (keep_json and name.lower().endswith(".json")):
            shutil.copy2(s, d)

    # Delete dst files/dirs not in src, but preserve .json files
    for name in dst_names:
        if name not in src_names:
            d = os.path.join(dst, name)
            if os.path.isfile(d):
                if not (keep_json and name.lower().endswith(".json")):
                    os.remove(d)
            elif os.path.isdir(d):
                has_json = keep_json and any(
                    f.lower().endswith(".json")
                    for _, _, files in os.walk(d)
                    for f in files
                )
                if not has_json:
                    shutil.rmtree(d)


def _load_migrations():
    """Import Snippets._migrations lazily, returning (read, apply) or (None, None).

    Deliberately not a module-level import: startup.py keeps its import cost
    near zero (see the pyrevit.forms note at the top), and lib/ is not on
    sys.path this early. Returns None on failure so a missing/broken shared
    module degrades to skipping migrations rather than blocking the update.
    """
    import sys as _sys
    lib_dir = os.path.join(EXTENSION_DIR, "lib")
    if os.path.isdir(lib_dir) and lib_dir not in _sys.path:
        _sys.path.insert(0, lib_dir)
    try:
        from Snippets._migrations import read_migrations, apply_migrations
        return read_migrations, apply_migrations
    except Exception:
        return None, None


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
        read_migrations, apply_migrations = _load_migrations()
        if read_migrations:
            migrations = read_migrations(extracted_root)
            if migrations:
                apply_migrations(migrations, TAB_DIR)

        status_lbl.Text = "Applying update..."
        sync_tree(new_tab, TAB_DIR)

        # ── Sync lib/ (shared Snippets modules) ───────────────────────────────
        # Kept out of the root-file loop below, which only handles files. lib
        # is code, not settings, so it syncs with keep_json=False - otherwise a
        # new _icons.json or palette key would never reach an existing install.
        new_lib = os.path.join(extracted_root, "lib")
        if os.path.isdir(new_lib):
            sync_tree(new_lib, os.path.join(EXTENSION_DIR, "lib"), keep_json=False)

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
        _log_error("_check_and_notify")


# ── SCHEDULED PRINT (pySheets, works even if pySheets isn't open) ─────────────
#
# pySheets' own in-window scheduler (a DispatcherTimer) only runs while its
# window is open, closing the window kills it. This registers a session-level
# Application.Idling handler instead, so armed schedules survive the window
# closing, as long as Revit and the target document stay open.
#
# Several profiles ("punch cards") can be armed at once. Cards that come due
# together run one at a time, grouped by document: a pySheets window is built
# around whichever project was active when it opened, so each document needs
# its own window, and the document that was active beforehand is restored at
# the end.
#
# V1 only: if the target document isn't open when the schedule comes due, it
# is skipped rather than opened automatically. Auto-opening the document is a
# planned V2 addition, not built yet.

def _find_pysheets_dir():
    """Locate pySheets.pushbutton under Seed43.tab. Returns the folder
    path, or None if it can't be found (extension reorganized, tool
    removed, etc, in which case the scheduler just does nothing)."""
    try:
        for root, _dirs, _files in os.walk(TAB_DIR):
            # Case-insensitive: Windows treats pySheets and pySheets as the
            # same folder, but this comparison would not, so a rename that
            # looks like a no-op silently kills the scheduler.
            if os.path.basename(root).lower() == "pysheets.pushbutton":
                return root
    except Exception:
        pass
    return None


_PYSHEETS_DIR = _find_pysheets_dir()

# pySheets keeps its settings in .user now (see Snippets/_userdata.py). The
# path is built here rather than imported, to keep startup's import cost at
# zero - if the layout there ever changes, this has to change with it.
#
# The legacy location is still checked as a fallback: a schedule armed before
# pySheets first migrated sits beside the tool, and the move only happens when
# the user next opens the window - which may be after this fires.
_SCHEDULE_FILE_NEW = os.path.join(
    EXTENSION_DIR, ".user", "pySheets", "settings", "scheduled_print.json")
_SCHEDULE_FILE_OLD = (
    os.path.join(_PYSHEETS_DIR, "userdata", "settings", "scheduled_print.json")
    if _PYSHEETS_DIR else None
)


def _schedule_file():
    """Return whichever schedule file exists, preferring the .user one."""
    try:
        if os.path.isfile(_SCHEDULE_FILE_NEW):
            return _SCHEDULE_FILE_NEW
        if _SCHEDULE_FILE_OLD and os.path.isfile(_SCHEDULE_FILE_OLD):
            return _SCHEDULE_FILE_OLD
    except Exception:
        pass
    return None
_LIB_DIR = os.path.join(EXTENSION_DIR, "lib")

_last_schedule_check = [0.0]

_UNSET = object()
_sched_mod_cache = [_UNSET]


def _schedule_mod():
    """The shared schedule rules from lib/Snippets/_schedule.py.

    Imported lazily and cached: it costs nothing but json/datetime, but there
    is no reason to touch the disk on a Revit launch where nothing is armed.
    Returns None if lib is missing, in which case the scheduler stands down."""
    if _sched_mod_cache[0] is _UNSET:
        _sched_mod_cache[0] = None
        try:
            import sys as _sys
            if _LIB_DIR and os.path.isdir(_LIB_DIR) and _LIB_DIR not in _sys.path:
                _sys.path.insert(0, _LIB_DIR)
            from Snippets import _schedule
            _sched_mod_cache[0] = _schedule
        except Exception:
            _log_error("_schedule_mod")
    return _sched_mod_cache[0]


def _read_schedule():
    """The armed cards file, already normalised (and upgraded from the old
    single-schedule format if that is what is still on disk)."""
    sched = _schedule_mod()
    path  = _schedule_file()
    if not sched or not path:
        return None
    try:
        return sched.read_armed_file(path)
    except Exception:
        return None


def _write_schedule(data):
    """Write the armed cards back — used when retiring a card that missed
    its slot, so it doesn't sit stale forever."""
    sched = _schedule_mod()
    path  = _schedule_file() or _SCHEDULE_FILE_NEW
    if not sched or not path:
        return
    try:
        sched.write_armed_file(path, data)
    except Exception:
        _log_error("_write_schedule")


def _launch_pysheets_schedule(entries):
    """Import pySheets fresh and hand it the cards due for the document that
    is active right now. Imported lazily here, not at module load, to keep it
    out of every Revit startup."""
    import sys as _sys
    if _PYSHEETS_DIR and _PYSHEETS_DIR not in _sys.path:
        _sys.path.insert(0, _PYSHEETS_DIR)
    if _LIB_DIR and os.path.isdir(_LIB_DIR) and _LIB_DIR not in _sys.path:
        _sys.path.insert(0, _LIB_DIR)
    import pySheets
    pySheets.launch_scheduled(entries)


from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
from Autodesk.Revit.DB import ModelPathUtils


class _PySheetsScheduleHandler(IExternalEventHandler):
    """Runs on Revit's own API thread (that's the whole point of
    ExternalEvent), safely deferred by Revit itself until no command is
    active, so this never interrupts something the user is mid-way
    through doing."""

    def Execute(self, uiapp):
        try:
            sched_mod = _schedule_mod()
            sched     = _read_schedule()
            if not sched_mod or not sched:
                return
            due = sched_mod.due_entries(sched)
            if not due:
                return

            # Cards that missed their slot by more than the grace window are
            # rolled straight on to their next occurrence here - no window
            # opens, nothing prints. Doing it in this pass matters: otherwise
            # a card missed over a weekend stays stuck in the past and can
            # never fire again.
            runnable, retired = [], False
            for entry in due:
                if not sched_mod.is_stale(sched, entry):
                    runnable.append(entry)
                    continue
                retired = True
                if not sched_mod.advance_entry(entry):
                    sched["entries"] = [e for e in sched.get("entries") or []
                                        if e is not entry]
            if retired:
                _write_schedule(sched)
            if not runnable:
                return
            due = runnable

            # Whatever the user was working in comes back at the end - a
            # scheduled print should not quietly move them to another model.
            previous_path = None
            try:
                previous_path = uiapp.ActiveUIDocument.Document.PathName
            except Exception:
                pass

            open_paths = {}
            for d in uiapp.Application.Documents:
                try:
                    if d.PathName and not d.IsLinked:
                        open_paths[os.path.normcase(d.PathName)] = d.PathName
                except Exception:
                    continue

            # One window per document, cards grouped so each window runs every
            # card belonging to its own project before the next one opens.
            by_doc = []
            for entry in due:
                key = os.path.normcase(entry.get("document_path") or "")
                if key not in open_paths:
                    continue        # V1: not open, skip rather than open it
                for path, entries in by_doc:
                    if path == key:
                        entries.append(entry)
                        break
                else:
                    by_doc.append((key, [entry]))

            for key, entries in by_doc:
                try:
                    model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(
                        open_paths[key])
                    uiapp.OpenAndActivateDocument(model_path)
                except Exception:
                    pass  # already open+active, activation failing is not fatal
                try:
                    _launch_pysheets_schedule(entries)
                except Exception:
                    _log_error("_launch_pysheets_schedule")

            if previous_path and by_doc:
                try:
                    uiapp.OpenAndActivateDocument(
                        ModelPathUtils.ConvertUserVisiblePathToModelPath(
                            previous_path))
                except Exception:
                    pass
        except Exception:
            _log_error("_PySheetsScheduleHandler.Execute")

    def GetName(self):
        return "Seed43 pySheets Scheduled Print"


try:
    _pysheets_schedule_event = ExternalEvent.Create(_PySheetsScheduleHandler())
except Exception:
    # Unwrapped, a failure here kills the whole module before the update
    # check or the Idling subscription is reached, with nothing to say why.
    _pysheets_schedule_event = None
    _log_error("ExternalEvent.Create")


# ── BATCH UPGRADE HEADLESS PICKUP ───────────────────────────────────────────
#
# Batch Upgrade drives other Revit versions by launching them directly and
# leaving a job file for this hook to find - see
# Batch Upgrade.pushbutton/tools/job_io.py and tools/headless_batch.py for
# the full story on why (short version: the pyRevit CLI's own runner addin
# can end up version-mismatched against a clone's engine assemblies; a
# normal Revit launch with pyRevit attached, which is what this is, never
# goes through that addin at all).
#
# Cheap by design, same discipline as the pySheets check below: an
# os.path.isfile() on every Idling tick when there's nothing to do, which
# is the normal case on every ordinary Revit launch. Only imports anything
# - and only then walks the tab tree to find the pushbutton folder - on
# the rare tick a job is actually waiting.

def _find_batch_upgrade_dir():
    """Locate Batch Upgrade.pushbutton under Seed43.tab. Returns the folder
    path, or None if it can't be found (extension reorganized, tool
    removed, etc, in which case this hook just does nothing), same
    pattern as _find_pysheets_dir above."""
    try:
        for root, _dirs, _files in os.walk(TAB_DIR):
            if os.path.basename(root).lower() == "batch upgrade.pushbutton":
                return root
    except Exception:
        pass
    return None


_BATCH_UPGRADE_DIR = _find_batch_upgrade_dir()

# Batch Upgrade's job files live under .user, same convention pySheets
# uses for its own settings (see _SCHEDULE_FILE_NEW above). Kept
# independent of _BATCH_UPGRADE_DIR (which can be None) since this path
# doesn't depend on where the pushbutton itself lives - must also be kept
# in step with tools/job_io.py's own _ROOT, which derives the same path
# a different way (relative to its own file location, since it can't
# import this module).
_BATCH_UPGRADE_JOB_DIR = os.path.join(
    EXTENSION_DIR, ".user", "BatchUpgrade", "settings")


def _batch_upgrade_job_path(year):
    return os.path.join(_BATCH_UPGRADE_JOB_DIR, "job_{}.json".format(year))


def _check_batch_upgrade():
    """If a job is queued for this Revit year, hand off to headless_batch
    and let it close Revit when done.

    Safe to call on every Idling tick: headless_batch deletes the job file
    the instant it's picked up, before doing any of the actual work, so a
    second tick before Revit has actually finished exiting just finds
    nothing there and returns immediately.
    """
    if not _BATCH_UPGRADE_DIR:
        return
    try:
        year = int(str(__revit__.Application.VersionNumber)[:4])
    except Exception:
        return
    if not os.path.isfile(_batch_upgrade_job_path(year)):
        return

    try:
        import sys as _sys
        if _BATCH_UPGRADE_DIR not in _sys.path:
            _sys.path.insert(0, _BATCH_UPGRADE_DIR)
        from tools import headless_batch
        # PostCommand (used to close Revit once the job's done) needs a
        # real UIApplication, not the UIControlledApplication __revit__
        # sometimes is here - same wrapping _subscribe_idling already
        # falls back to below, just done unconditionally rather than only
        # on failure, since headless_batch always needs the real thing.
        from Autodesk.Revit.UI import UIApplication
        uiapp = UIApplication(__revit__.Application)
        headless_batch.check_and_run(__revit__.Application, uiapp)
    except Exception:
        _log_error("_check_batch_upgrade")


# ── BATCH UPGRADE SCHEDULED RUN ─────────────────────────────────────────────
#
# A one-time "run this batch at this date and time" - the timer reuses the
# same Snippets/_schedule.py engine and the same 20-second-throttled Idling
# check pySheets' own scheduler uses (see _on_idling below), just with its
# own separate armed file: Batch Upgrade entries aren't tied to an open
# document the way pySheets' cards are, so they don't belong mixed in with
# pySheets' scheduled_print.json.
#
# Only ever one entry armed at a time, and always dropped the instant it's
# due - fired or missed its grace window, either way there is no second
# occurrence to roll forward to. That is what makes this "once" rather than
# "repeat": nothing here ever calls Snippets._schedule.advance_entry, since
# that function's repeat-handling branch would never apply to a Batch
# Upgrade entry in the first place.

_BATCH_UPGRADE_SCHEDULE_FILE = os.path.join(
    EXTENSION_DIR, ".user", "BatchUpgrade", "settings", "scheduled_run.json")


def _read_batch_upgrade_schedule():
    sched = _schedule_mod()
    if not sched:
        return None
    try:
        return sched.read_armed_file(_BATCH_UPGRADE_SCHEDULE_FILE)
    except Exception:
        return None


def _write_batch_upgrade_schedule(data):
    sched = _schedule_mod()
    if not sched:
        return
    try:
        sched.write_armed_file(_BATCH_UPGRADE_SCHEDULE_FILE, data)
    except Exception:
        _log_error("_write_batch_upgrade_schedule")


def _clear_batch_upgrade_snapshot(path):
    """Delete a spent scheduled job's frozen settings, if it left any."""
    if not _BATCH_UPGRADE_DIR:
        return
    try:
        import sys as _sys
        if _BATCH_UPGRADE_DIR not in _sys.path:
            _sys.path.insert(0, _BATCH_UPGRADE_DIR)
        from tools import schedule_io
        schedule_io.clear_snapshot(path)
    except Exception:
        _log_error("_clear_batch_upgrade_snapshot")


def _run_batch_upgrade_scheduled(entry):
    """Run the settings a due entry froze, headlessly - no progress window,
    nobody is watching.

    The entry points at a one-off snapshot written when the schedule was
    armed, not a saved template: a batch upgrade is a one-shot job, so the
    snapshot is binned once it has run.
    """
    if not _BATCH_UPGRADE_DIR:
        return
    snapshot_path = entry.get("profile_path")
    try:
        import sys as _sys
        if _BATCH_UPGRADE_DIR not in _sys.path:
            _sys.path.insert(0, _BATCH_UPGRADE_DIR)
        from tools import schedule_io, batch_runner
        from Autodesk.Revit.UI import UIApplication

        data = schedule_io.read_snapshot(snapshot_path)
        if not data:
            return

        app = __revit__.Application
        uiapp = UIApplication(app)
        host_year = int(str(app.VersionNumber)[:4])

        by_year = batch_runner.run_batch(
            data.get("files") or [], data.get("out_dir") or "",
            data.get("targets") or [], bool(data.get("audit")),
            bool(data.get("compact", True)),
            app, uiapp, host_year, progress=batch_runner.NullProgress())
        batch_runner.write_report(by_year, data.get("out_dir") or "",
                                  show_output=False)
    except Exception:
        _log_error("_run_batch_upgrade_scheduled")
    finally:
        # The armed entry is already gone by the time this runs; drop the
        # snapshot too so a spent one-shot job leaves nothing behind. In
        # `finally` because a run that threw halfway is still spent - the
        # entry is never coming back to retry it.
        try:
            from tools import schedule_io as _sched_io
            _sched_io.clear_snapshot(snapshot_path)
        except Exception:
            pass


class _BatchUpgradeScheduleHandler(IExternalEventHandler):
    """Runs on Revit's own API thread, same reasoning as
    _PySheetsScheduleHandler above: safely deferred until no command is
    active, so a scheduled run never interrupts something the user is
    mid-way through doing."""

    def Execute(self, uiapp):
        try:
            sched_mod = _schedule_mod()
            sched     = _read_batch_upgrade_schedule()
            if not sched_mod or not sched:
                return
            due = sched_mod.due_entries(sched, respect_heartbeat=False)
            if not due:
                return

            # Only one entry is ever armed, but handled as a list for
            # symmetry with pySheets' handler above. Dropped unconditionally
            # once due - a one-time entry gets exactly one chance, whether
            # it fires now or turns out to be too stale to run at all.
            entry = due[0]
            stale = sched_mod.is_stale(sched, entry)
            sched["entries"] = [e for e in sched.get("entries") or []
                                if e is not entry]
            _write_batch_upgrade_schedule(sched)

            if stale:
                # Too late to run, but the entry is spent either way - bin its
                # snapshot as well, or a missed schedule leaves the frozen job
                # file behind forever.
                _clear_batch_upgrade_snapshot(entry.get("profile_path"))
                return
            _run_batch_upgrade_scheduled(entry)
        except Exception:
            _log_error("_BatchUpgradeScheduleHandler.Execute")

    def GetName(self):
        return "Seed43 Batch Upgrade Scheduled Run"


try:
    _batch_upgrade_schedule_event = ExternalEvent.Create(
        _BatchUpgradeScheduleHandler())
except Exception:
    _batch_upgrade_schedule_event = None
    _log_error("ExternalEvent.Create (Batch Upgrade)")


def _on_idling(sender, args):
    """Cheap periodic check (throttled to roughly every 20 seconds) for
    whether any armed pySheets card, or the one possible armed Batch
    Upgrade schedule, has come due. Also checks, unthrottled since it's
    cheap, for a waiting Batch Upgrade job - see _check_batch_upgrade
    above, a separate feature from the scheduled-run timer here."""
    _check_batch_upgrade()

    try:
        import time
        now = time.time()
        if now - _last_schedule_check[0] < 20:
            return
        _last_schedule_check[0] = now

        if _batch_upgrade_schedule_event is not None:
            bsched = _read_batch_upgrade_schedule()
            if bsched and bsched.get("entries"):
                sched_mod = _schedule_mod()
                if sched_mod and sched_mod.due_entries(
                        bsched, respect_heartbeat=False):
                    _batch_upgrade_schedule_event.Raise()

        if _pysheets_schedule_event is None:
            return

        sched_mod = _schedule_mod()
        sched     = _read_schedule()
        if not sched_mod or not sched or not sched.get("entries"):
            return
        if not sched_mod.due_entries(sched):
            return

        _pysheets_schedule_event.Raise()
    except Exception:
        _log_error("_on_idling")


_uiapp_ref = [None]


def main():
    # Ages the log out on every Revit launch, so it shrinks even in a stretch
    # where nothing is ever logged to trigger the per-write check.
    _prune_log()

    ui_dispatcher = Dispatcher.CurrentDispatcher

    def worker():
        _check_and_notify(ui_dispatcher)

    t = Thread(ThreadStart(worker))
    t.IsBackground = True
    t.Start()

    _subscribe_idling()
    _rebuild_stale_icons()


def _rebuild_stale_icons():
    """Redraw any button icon whose SVG or the palette is newer than its PNG.

    About rebuilds on close when the accent changes, but Revit holds on to
    the icon files it has already loaded, so a write can be refused or simply
    not picked up until the next session. Checking again here closes that
    gap. Deliberately on the UI thread rather than the worker above: WPF
    rendering needs an STA thread, and this one is.

    Costs a handful of stat calls when nothing has changed, which is the
    normal case - nothing is rendered unless a source is actually newer."""
    try:
        import sys as _sys
        if _LIB_DIR and os.path.isdir(_LIB_DIR) and _LIB_DIR not in _sys.path:
            _sys.path.insert(0, _LIB_DIR)
        from Snippets import _svg_icons
        written, problems = _svg_icons.rebuild_all(EXTENSION_DIR, only_stale=True)
        if problems:
            _log_note("icon rebuild: {} written, {} skipped".format(
                len(written), len(problems)))
            for problem in problems:
                _log_note("  " + problem)
    except Exception:
        _log_error("_rebuild_stale_icons")


def _subscribe_idling():
    """Subscribe _on_idling to Revit's Idling event.

    Idling lives on UIControlledApplication / UIApplication, NOT on the DB
    Application - `__revit__.Application.Idling` raises AttributeError, which
    is what silently kept this scheduler from ever running. What `__revit__`
    itself is depends on the host, so try it directly first and fall back to
    wrapping the DB application, recording which one worked."""
    try:
        __revit__.Idling += _on_idling
        return True
    except Exception:
        _log_error("Idling subscribe on __revit__")

    try:
        from Autodesk.Revit.UI import UIApplication
        # Kept on the module so the wrapper cannot be collected while it is
        # holding the only reference to our subscription.
        _uiapp_ref[0] = UIApplication(__revit__.Application)
        _uiapp_ref[0].Idling += _on_idling
        return True
    except Exception:
        _log_error("Idling subscribe on UIApplication")

    _log_note("Idling could not be subscribed — scheduler is off")
    return False


if __name__ == "__main__":
    main()