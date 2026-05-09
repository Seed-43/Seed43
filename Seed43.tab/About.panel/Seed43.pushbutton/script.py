# -*- coding: utf-8 -*-
"""
Seed43 Startup migration script.
Runs automatically when Revit starts via pyRevit.
Shows an update popup and replaces the entire Seed43.extension folder from GitHub.
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

from pyrevit import forms


# -- VARIABLES -----------------------------------------------------------------

GITHUB_USER   = "Seed-43"
GITHUB_REPO   = "Seed43"
GITHUB_BRANCH = "main"

REPO_ZIP_URL  = "https://github.com/{}/{}/archive/refs/heads/{}.zip".format(
    GITHUB_USER, GITHUB_REPO, GITHUB_BRANCH)

APPDATA       = os.environ.get("APPDATA", "")
EXTENSION_DIR = os.path.join(APPDATA, "pyRevit", "Extensions", "Seed43.extension")

# icon sits next to this script before the update wipes it
SCRIPT_DIR    = os.path.dirname(__file__)
ICON_PATH     = os.path.join(SCRIPT_DIR, "icon.png")


# -- XAML ----------------------------------------------------------------------

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
                           Text="A new version of Seed43 is available.&#x0a;&#x0a;Click Update Now to install the latest version, or Not Now to skip."
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


# -- UPDATE LOGIC --------------------------------------------------------------

def download_and_apply_update(window, status_lbl, progress_bar):
    """
    Download zip from GitHub, wipe Seed43.extension, copy everything in.
    Runs on a background thread; dispatches UI updates back to the window.
    Skips existing .json and .yaml files to preserve local config.
    """
    tmp_zip = os.path.join(os.environ.get("TEMP", APPDATA), "seed43_startup_update.zip")
    tmp_dir = os.path.join(os.environ.get("TEMP", APPDATA), "seed43_startup_update_tmp")

    SKIP_EXT = (".json", ".yaml")

    def log(msg):
        window.Dispatcher.Invoke(Action(lambda: setattr(status_lbl, "Text", msg)))

    def set_progress(val):
        window.Dispatcher.Invoke(Action(lambda: setattr(progress_bar, "Value", val)))

    def show_progress():
        window.Dispatcher.Invoke(Action(lambda: setattr(progress_bar, "Visibility", Visibility.Visible)))

    def show_status():
        window.Dispatcher.Invoke(Action(lambda: setattr(status_lbl, "Visibility", Visibility.Visible)))

    try:
        show_status()
        show_progress()
        log("Connecting to GitHub...")

        wc = WebClient()
        wc.Headers.Add("Cache-Control", "no-cache, no-store")
        wc.DownloadFile(REPO_ZIP_URL, tmp_zip)
        set_progress(40)

        log("Extracting...")
        if os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir)
        os.makedirs(tmp_dir)
        with zipfile.ZipFile(tmp_zip, "r") as zf:
            zf.extractall(tmp_dir)
        set_progress(60)

        extracted_root = None
        for item in os.listdir(tmp_dir):
            full = os.path.join(tmp_dir, item)
            if os.path.isdir(full):
                extracted_root = full
                break
        if not extracted_root:
            raise Exception("Could not find extracted folder.")

        log("Installing...")

        # Wipe and replace Seed43.extension
        if os.path.isdir(EXTENSION_DIR):
            shutil.rmtree(EXTENSION_DIR)

        def safe_copy_tree(src, dst):
            if not os.path.exists(dst):
                os.makedirs(dst)
            for item in os.listdir(src):
                s = os.path.join(src, item)
                d = os.path.join(dst, item)
                if os.path.isdir(s):
                    safe_copy_tree(s, d)
                else:
                    ext = os.path.splitext(item)[1].lower()
                    if ext in SKIP_EXT and os.path.exists(d):
                        continue
                    shutil.copy2(s, d)

        safe_copy_tree(extracted_root, EXTENSION_DIR)
        set_progress(90)

        # Cleanup
        if os.path.exists(tmp_zip):
            os.remove(tmp_zip)
        if os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir)

        set_progress(100)
        log("Update complete. Please restart Revit.")
        return True

    except Exception as ex:
        log("Update failed: {}".format(str(ex)))
        try:
            if os.path.exists(tmp_zip):
                os.remove(tmp_zip)
            if os.path.exists(tmp_dir):
                shutil.rmtree(tmp_dir)
        except Exception:
            pass
        return False


# -- WINDOW --------------------------------------------------------------------

class UpdateWindow(object):

    def __init__(self):
        self.window       = XamlReader.Parse(WINDOW_XAML)
        self.title_lbl    = self.window.FindName("update_title_lbl")
        self.msg_lbl      = self.window.FindName("update_msg_lbl")
        self.progress_bar = self.window.FindName("update_progress")
        self.status_lbl   = self.window.FindName("status_lbl")
        self.skip_btn     = self.window.FindName("skip_btn")
        self.update_btn   = self.window.FindName("update_btn")
        self.close_btn    = self.window.FindName("header_close_btn")

        if os.path.exists(ICON_PATH):
            img           = self.window.FindName("header_icon")
            bmp           = BitmapImage()
            bmp.BeginInit()
            bmp.UriSource = Uri(ICON_PATH, UriKind.Absolute)
            bmp.EndInit()
            img.Source    = bmp

        self.close_btn.Click  += lambda s, e: self.window.Close()
        self.skip_btn.Click   += lambda s, e: self.window.Close()
        self.update_btn.Click += self._on_update

    def _on_update(self, sender, e):
        self.update_btn.IsEnabled = False
        self.skip_btn.IsEnabled   = False
        self.close_btn.IsEnabled  = False

        window      = self.window
        status_lbl  = self.status_lbl
        progress_bar = self.progress_bar
        title_lbl   = self.title_lbl
        update_btn  = self.update_btn
        skip_btn    = self.skip_btn
        close_btn   = self.close_btn

        def worker():
            success = download_and_apply_update(window, status_lbl, progress_bar)

            def done():
                if success:
                    title_lbl.Text          = "Update Complete"
                    update_btn.Content      = "Done"
                    update_btn.IsEnabled    = True
                    update_btn.Click       -= self._on_update
                    update_btn.Click       += lambda s, e: window.Close()
                else:
                    skip_btn.IsEnabled  = True
                    update_btn.IsEnabled = True
                close_btn.IsEnabled = True

            window.Dispatcher.Invoke(Action(done))

        t = Thread(ThreadStart(worker))
        t.IsBackground = True
        t.Start()

    def show(self):
        self.window.ShowDialog()


# -- ENTRY POINT ---------------------------------------------------------------

def main():
    ui_dispatcher = Dispatcher.CurrentDispatcher

    def show():
        UpdateWindow().show()

    ui_dispatcher.Invoke(Action(show))


if __name__ == "__main__":
    main()
