# -*- coding: utf-8 -*-
"""
Seed43 About / Install dialog
PyRevit pushbutton script - IronPython + WPF
Location: Seed43.extension/Seed43.tab/About.panel/Stack01.stack/Seed43.pushbutton/script.py
Loads UI from: seed43.xaml (same folder)
"""

import os
import clr
import json

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("System.Net")
clr.AddReference("System.IO.Compression")
clr.AddReference("System.IO.Compression.FileSystem")

import System
from System.Windows.Markup import XamlReader
from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage, Visibility
import System.Net
import System.IO
import System.IO.Compression
from System.Net import WebClient
from System.IO import File, Directory, Path, StreamReader
from System.IO.Compression import ZipFile
from System.Threading import Thread, ThreadStart

# ── Constants ─────────────────────────────────────────────────────────────────
GITHUB_ORG   = "Seed-43"
MAIN_REPO    = "Seed43"
BRANCH       = "main"

CHANGELOG_URL = "https://raw.githubusercontent.com/{o}/{r}/{b}/changelog.json".format(
    o=GITHUB_ORG, r=MAIN_REPO, b=BRANCH)

APPDATA = os.environ.get("APPDATA", "")

TOOLS = [
    {
        "id":            "pytransmit",
        "name":          "PyTransmit",
        "repo":          "Seed43-PyTransmit",
        "install_dir":   os.path.join(APPDATA, "pyRevit", "Extensions", "PyTransmit.extension"),
        "version_file":  os.path.join(APPDATA, "pyRevit", "Extensions", "PyTransmit.extension", "version.txt"),
        "changelog_url": "https://raw.githubusercontent.com/{o}/{r}/{b}/changelog.json".format(
                             o=GITHUB_ORG, r="Seed43-PyTransmit", b=BRANCH),
        "zip_url":       "https://github.com/{o}/{r}/archive/refs/heads/{b}.zip".format(
                             o=GITHUB_ORG, r="Seed43-PyTransmit", b=BRANCH),
    }
]

S43_VERSION_FILE = os.path.join(APPDATA, "pyRevit", "Extensions", "seed43_version.txt")

# ── Load XAML from file ───────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(__file__)
XAML_PATH  = os.path.join(SCRIPT_DIR, "seed43.xaml")

# ── Helpers ───────────────────────────────────────────────────────────────────

def load_xaml(path):
    reader = StreamReader(path)
    window = XamlReader.Load(reader.BaseStream)
    reader.Close()
    return window

def fetch_json(url):
    try:
        return json.loads(WebClient().DownloadString(url))
    except Exception:
        return None

def get_local_version(version_file):
    if File.Exists(version_file):
        return File.ReadAllText(version_file).strip()
    return None

def write_local_version(version_file, version):
    Directory.CreateDirectory(Path.GetDirectoryName(version_file))
    File.WriteAllText(version_file, version)

def dispatch(window, fn):
    window.Dispatcher.Invoke(System.Action(fn))


# ── Dialog ────────────────────────────────────────────────────────────────────

class Seed43Dialog(object):

    def __init__(self):
        self.window     = load_xaml(XAML_PATH)
        self.tool       = TOOLS[0]
        self._expanded  = False
        self._installed = False
        self._busy      = False
        self._bind()
        self._check_versions()

    # ── Bind events ───────────────────────────────────────────────────────────

    def _bind(self):
        self.window.FindName("header_close").Click              += lambda s, e: self.window.Close()
        self.window.FindName("footer_close").Click              += lambda s, e: self.window.Close()
        self.window.FindName("update_ribbon").MouseLeftButtonUp += self._on_s43_update
        self.window.FindName("pt_header").MouseLeftButtonUp     += self._on_toggle
        self.window.FindName("pt_action_btn").Click             += self._on_action

    # ── Toggle expand / collapse ──────────────────────────────────────────────

    def _on_toggle(self, sender, args):
        if self._busy:
            return
        self._expanded = not self._expanded
        self.window.FindName("pt_body").Visibility = (
            Visibility.Visible if self._expanded else Visibility.Collapsed)
        self.window.FindName("pt_chevron").Text = (
            u"\u25B2" if self._expanded else u"\u25BC")

    # ── Action button (Install / Uninstall) ───────────────────────────────────

    def _on_action(self, sender, args):
        if self._busy:
            return
        if self._installed:
            self._run_uninstall()
        else:
            self._run_install()

    # ── Install ───────────────────────────────────────────────────────────────

    def _run_install(self):
        self._busy = True
        self._set_busy_ui("Connecting to GitHub...")
        tool = self.tool

        def log(msg):
            dispatch(self.window, lambda: setattr(
                self.window.FindName("pt_progress_lbl"), "Text", msg))

        def worker():
            try:
                tmp_zip = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(), tool["repo"] + "_install.zip")
                tmp_dir = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(), tool["repo"] + "_extracted")

                log("Downloading files...")
                WebClient().DownloadFile(tool["zip_url"], tmp_zip)

                log("Extracting...")
                if Directory.Exists(tmp_dir):
                    Directory.Delete(tmp_dir, True)
                ZipFile.ExtractToDirectory(tmp_zip, tmp_dir)

                extracted_root = None
                for d in Directory.GetDirectories(tmp_dir):
                    extracted_root = d
                    break
                if not extracted_root:
                    raise System.Exception("Could not find extracted folder.")

                log("Installing...")
                if Directory.Exists(tool["install_dir"]):
                    Directory.Delete(tool["install_dir"], True)
                Directory.Move(extracted_root, tool["install_dir"])

                if File.Exists(tmp_zip):
                    File.Delete(tmp_zip)
                if Directory.Exists(tmp_dir):
                    Directory.Delete(tmp_dir, True)

                remote  = fetch_json(tool["changelog_url"])
                version = remote.get("version", "unknown") if remote else "unknown"
                write_local_version(tool["version_file"], version)

                dispatch(self.window, lambda: self._on_install_done(version, remote))

            except System.Exception as ex:
                dispatch(self.window, lambda: self._on_error(str(ex)))

        t = Thread(ThreadStart(worker))
        t.IsBackground = True
        t.Start()

    # ── Uninstall ─────────────────────────────────────────────────────────────

    def _run_uninstall(self):
        result = MessageBox.Show(
            "Uninstall PyTransmit?\n\nThis will delete all files from:\n" + self.tool["install_dir"],
            "Uninstall PyTransmit",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question
        )
        if str(result) != "Yes":
            return

        self._busy = True
        self._set_busy_ui("Removing files...")
        tool = self.tool

        def log(msg):
            dispatch(self.window, lambda: setattr(
                self.window.FindName("pt_progress_lbl"), "Text", msg))

        def worker():
            try:
                log("Removing files...")
                if Directory.Exists(tool["install_dir"]):
                    Directory.Delete(tool["install_dir"], True)
                log("Cleaning up...")
                dispatch(self.window, lambda: self._on_uninstall_done())
            except System.Exception as ex:
                dispatch(self.window, lambda: self._on_error(str(ex)))

        t = Thread(ThreadStart(worker))
        t.IsBackground = True
        t.Start()

    # ── State helpers ─────────────────────────────────────────────────────────

    def _set_busy_ui(self, label):
        self.window.FindName("pt_action_btn").IsEnabled       = False
        self.window.FindName("pt_progress_wrap").Visibility   = Visibility.Visible
        self.window.FindName("pt_progress_lbl").Text          = label

    def _on_install_done(self, version, remote):
        self._busy      = False
        self._installed = True
        self.window.FindName("pt_progress_wrap").Visibility = Visibility.Collapsed
        btn           = self.window.FindName("pt_action_btn")
        btn.IsEnabled = True
        btn.Content   = "Uninstall"
        btn.Style     = self.window.FindResource("DangerBtn")
        self.window.FindName("pt_dot").Fill = \
            System.Windows.Media.BrushConverter().ConvertFrom("#208A3C")
        self.window.FindName("pt_status_lbl").Text = \
            u"Installed  v{0}  \u2192  {1}".format(version, self.tool["install_dir"])
        if remote:
            ver     = remote.get("version", "")
            changes = remote.get("changes", [])
            self.window.FindName("pt_changelog").Text = \
                u"Latest: v{0}\n{1}".format(ver, u"\n".join([u"\u203a " + c for c in changes]))

    def _on_uninstall_done(self):
        self._busy      = False
        self._installed = False
        self.window.FindName("pt_progress_wrap").Visibility = Visibility.Collapsed
        btn           = self.window.FindName("pt_action_btn")
        btn.IsEnabled = True
        btn.Content   = "Install"
        btn.Style     = self.window.FindResource("SmallBtn")
        self.window.FindName("pt_dot").Fill = \
            System.Windows.Media.BrushConverter().ConvertFrom("#A0AABB")
        self.window.FindName("pt_status_lbl").Text = "Not installed"
        self.window.FindName("pt_changelog").Text  = ""

    def _on_error(self, msg):
        self._busy = False
        self.window.FindName("pt_progress_wrap").Visibility = Visibility.Collapsed
        self.window.FindName("pt_action_btn").IsEnabled     = True
        MessageBox.Show("Operation failed:\n\n" + msg, "Seed43",
                        MessageBoxButton.OK, MessageBoxImage.Error)

    # ── Version checks ────────────────────────────────────────────────────────

    def _check_versions(self):
        def worker():
            local_s43  = get_local_version(S43_VERSION_FILE)
            remote_s43 = fetch_json(CHANGELOG_URL)
            dispatch(self.window, lambda: self._update_s43_ui(local_s43, remote_s43))
            local_pt  = get_local_version(self.tool["version_file"])
            remote_pt = fetch_json(self.tool["changelog_url"])
            dispatch(self.window, lambda: self._update_pt_ui(local_pt, remote_pt))

        t = Thread(ThreadStart(worker))
        t.IsBackground = True
        t.Start()

    def _update_s43_ui(self, local, remote):
        self.window.FindName("s43_version").Text = (
            u"\u25CF  Installed  v{0}".format(local) if local else "Version unknown")
        if remote:
            ver     = remote.get("version", "")
            changes = remote.get("changes", [])
            self.window.FindName("s43_changelog").Text = \
                u"Latest: v{0}   {1}".format(ver, u"  \u2022  ".join(changes[:2]))
            # Show orange ribbon if update available
            if local and ver and ver != local:
                self._remote_s43_version = ver
                self.window.FindName("update_ribbon_version").Text = \
                    u"v{0}  \u2192  v{1}".format(local, ver)
                self.window.FindName("update_ribbon").Visibility = Visibility.Visible
        else:
            self.window.FindName("s43_changelog").Text = "Could not reach GitHub"

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

        ribbon = self.window.FindName("update_ribbon")
        ribbon.Visibility = Visibility.Collapsed

        import urllib.request as _ur
        import zipfile as _zf

        EXTENSIONS_DIR = os.path.join(os.environ.get("APPDATA", ""), "pyRevit", "Extensions")
        S43_INSTALL    = os.path.join(EXTENSIONS_DIR, "Seed43.extension")
        ZIP_URL        = "https://github.com/{o}/{r}/archive/refs/heads/{b}.zip".format(
                            o=GITHUB_ORG, r=MAIN_REPO, b=BRANCH)
        TEMP_ZIP       = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "seed43_update.zip")
        TEMP_DIR       = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "seed43_update_extracted")

        def log(msg):
            dispatch(self.window, lambda: setattr(
                self.window.FindName("s43_changelog"), "Text", msg))

        def worker():
            try:
                log("Downloading update...")
                WebClient().DownloadFile(ZIP_URL, TEMP_ZIP)
                log("Extracting...")
                if Directory.Exists(TEMP_DIR):
                    Directory.Delete(TEMP_DIR, True)
                ZipFile.ExtractToDirectory(TEMP_ZIP, TEMP_DIR)
                extracted_root = None
                for d in Directory.GetDirectories(TEMP_DIR):
                    extracted_root = d
                    break
                if not extracted_root:
                    raise System.Exception("Could not find extracted folder.")
                src = System.IO.Path.Combine(extracted_root, "Seed43.extension")
                if not Directory.Exists(src):
                    raise System.Exception("Seed43.extension not found in ZIP.")
                log("Installing update...")
                if Directory.Exists(S43_INSTALL):
                    Directory.Delete(S43_INSTALL, True)
                import shutil as _sh
                _sh.copytree(src, S43_INSTALL)
                remote = fetch_json(CHANGELOG_URL)
                version = remote.get("version", "unknown") if remote else "unknown"
                # Write version.txt into the newly installed extension
                new_version_file = System.IO.Path.Combine(S43_INSTALL, "version.txt")
                File.WriteAllText(new_version_file, version)
                # Also write a backup version file one level up so it survives overwrites
                backup_version_file = System.IO.Path.Combine(
                    os.path.join(os.environ.get("APPDATA", ""), "pyRevit", "Extensions"),
                    "seed43_version.txt"
                )
                File.WriteAllText(backup_version_file, version)
                if File.Exists(TEMP_ZIP):   File.Delete(TEMP_ZIP)
                if Directory.Exists(TEMP_DIR): Directory.Delete(TEMP_DIR, True)
                dispatch(self.window, lambda: self._on_s43_update_done(version))
            except System.Exception as ex:
                dispatch(self.window, lambda: self._on_error(str(ex)))

        t = Thread(ThreadStart(worker))
        t.IsBackground = True
        t.Start()

    def _on_s43_update_done(self, version):
        # Update the in-memory version so re-checking won't loop
        self._local_s43_version = version
        self.window.FindName("update_ribbon").Visibility = Visibility.Collapsed
        self.window.FindName("s43_version").Text = u"\u25CF  Installed  v{0}".format(version)
        self.window.FindName("s43_changelog").Text = u"Updated to v{0} \u2014 reload PyRevit to apply.".format(version)
        MessageBox.Show(
            "Seed43 updated to v{0}.\n\nReload PyRevit in Revit to apply the update.".format(version),
            "Seed43 Updated",
            MessageBoxButton.OK,
            MessageBoxImage.Information
        )

    def _update_pt_ui(self, local, remote):
        if local:
            self._installed = True
            btn         = self.window.FindName("pt_action_btn")
            btn.Content = "Uninstall"
            btn.Style   = self.window.FindResource("DangerBtn")
            self.window.FindName("pt_dot").Fill = \
                System.Windows.Media.BrushConverter().ConvertFrom("#208A3C")
            self.window.FindName("pt_status_lbl").Text = \
                u"Installed  v{0}  \u2192  {1}".format(local, self.tool["install_dir"])
        else:
            self.window.FindName("pt_status_lbl").Text = "Not installed"

        if remote:
            ver     = remote.get("version", "")
            changes = remote.get("changes", [])
            self.window.FindName("pt_changelog").Text = \
                u"Latest: v{0}\n{1}".format(ver, u"\n".join([u"\u203a " + c for c in changes]))
        else:
            self.window.FindName("pt_changelog").Text = "Could not reach GitHub"

    def show(self):
        self.window.ShowDialog()


# ── Entry point ───────────────────────────────────────────────────────────────
dialog = Seed43Dialog()
dialog.show()
