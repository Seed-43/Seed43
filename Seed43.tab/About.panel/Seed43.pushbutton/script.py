# -*- coding: utf-8 -*-
__title__  = "Check for Updates"
__author__  = "Seed43"
__doc__     = """
𝐕𝐄𝐑𝐒𝐈𝐎𝐍 𝟐𝟔𝟎𝟓𝟎𝟏
_____________________________________________________________________
Description:
Checks whether a newer version of Seed43 is available on GitHub and
offers to download and apply the update automatically.

The version is read from a version.txt file stored on GitHub. If a
newer version is found, a window appears showing your current version
and the available version. Clicking Update Now downloads the latest
Seed43.tab folder from GitHub and replaces the one on your machine.
Your settings and config files are not affected.
_____________________________________________________________________
How-to:
-> Run the tool from the Seed43 tab
-> If an update is available, a window will appear
-> Click Update Now to apply it, or Not Now to skip
-> If already up to date, a brief message confirms this
_____________________________________________________________________
Notes:
- An internet connection is required
- The update replaces only the Seed43.tab folder
- Your settings and config files are not affected
- Revit must be restarted after updating for changes to take effect
_____________________________________________________________________
Last update:
- Initial release
_____________________________________________________________________
"""

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

from pyrevit import forms, script


# ── XAML ──────────────────────────────────────────────────────────────────────

WINDOW_XAML = """
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Seed43 Update"
    Width="420"
    SizeToContent="Height"
    ResizeMode="NoResize"
    WindowStartupLocation="CenterScreen"
    Background="#3B4553"
    TextElement.FontFamily="Segoe UI">

    <Window.Resources>

        <DropShadowEffect x:Key="HeaderShadow"
                          Color="Black" Opacity="0.3"
                          ShadowDepth="2" BlurRadius="6"/>

        <DropShadowEffect x:Key="CardShadow"
                          Color="Black" Opacity="0.2"
                          ShadowDepth="2" BlurRadius="4"/>

        <Style x:Key="SmallButtonStyle" TargetType="Button">
            <Setter Property="Background"      Value="#208A3C"/>
            <Setter Property="Foreground"      Value="#F4FAFF"/>
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
                                <Setter TargetName="Bd" Property="Background" Value="#2B933F"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#1A6E2E"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Bd" Property="Background" Value="#555555"/>
                                <Setter Property="Foreground"                  Value="#888888"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="SecondaryButtonStyle" TargetType="Button">
            <Setter Property="Background"      Value="#404553"/>
            <Setter Property="Foreground"      Value="#F4FAFF"/>
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
                                <Setter TargetName="Bd" Property="Background" Value="#4E5566"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#333B48"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Bd" Property="Background" Value="#555555"/>
                                <Setter Property="Foreground"                  Value="#888888"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="CloseButtonStyle" TargetType="Button">
            <Setter Property="Background"      Value="Transparent"/>
            <Setter Property="Foreground"      Value="#F4FAFF"/>
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
                                <Setter TargetName="Bd" Property="Background" Value="#FF5555"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#CC4444"/>
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
                Background="#232933"
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
                               Foreground="#208A3C" Text="SEED"/>
                    <TextBlock FontSize="32" FontWeight="SemiBold"
                               Foreground="#F4FAFF" Text="43"/>
                    <TextBlock FontSize="20" FontWeight="SemiBold"
                               Foreground="#F4FAFF" Opacity="0.75"
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
                Background="#F4FAFF"
                BorderBrush="#208A3C"
                BorderThickness="1"
                CornerRadius="6"
                Padding="24"
                Effect="{StaticResource CardShadow}">
            <StackPanel>

                <TextBlock x:Name="update_title_lbl"
                           Text="Update Available"
                           Foreground="#208A3C"
                           FontSize="14"
                           FontWeight="SemiBold"
                           Margin="0,0,0,8"/>

                <TextBlock x:Name="update_msg_lbl"
                           Foreground="#2B3340"
                           FontSize="12"
                           TextWrapping="Wrap"
                           Margin="0,0,0,20"/>

                <ProgressBar x:Name="update_progress"
                             Height="4"
                             Margin="0,0,0,12"
                             Minimum="0"
                             Maximum="100"
                             Value="0"
                             Foreground="#208A3C"
                             Background="#D0D8E0"
                             BorderThickness="0"
                             Visibility="Collapsed"/>

                <TextBlock x:Name="status_lbl"
                           Foreground="#208A3C"
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

# File extensions to skip during update, preserving local config files
SKIP_EXTENSIONS = (".yaml", ".json")


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


def download_and_apply_update(status_lbl, progress_bar):
    """Download the repo zip, swap in the new Seed43.tab, update version.txt.
    Skips yaml and json files to preserve local config.
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

        # ── Replace Seed43.tab, skipping yaml and json files ─────────────────

        status_lbl.Text = "Applying update..."

        if os.path.isdir(TAB_DIR):
            shutil.rmtree(TAB_DIR)
        shutil.copytree(
            new_tab,
            TAB_DIR,
            ignore=shutil.ignore_patterns(*["*" + ext for ext in SKIP_EXTENSIONS])
        )

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
            img       = self.window.FindName("header_icon")
            bmp       = BitmapImage()
            bmp.BeginInit()
            bmp.UriSource = Uri(ICON_PATH, UriKind.Absolute)
            bmp.EndInit()
            img.Source = bmp

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
                forms.alert(
                    "Update applied successfully.\n\n"
                    "Please restart Revit for the changes to take effect.",
                    title="Seed43 Update"
                )

        ui_dispatcher.Invoke(Action(show_window))

    except Exception:
        pass


def main():
    ui_dispatcher = Dispatcher.CurrentDispatcher

    def worker():
        _check_and_notify(ui_dispatcher)

    t = Thread(ThreadStart(worker))
    t.IsBackground = True
    t.Start()


if __name__ == "__main__":
    main()
