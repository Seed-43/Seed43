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
import shutil
import zipfile

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

CHANGELOG_URL = "https://raw.githubusercontent.com/{o}/{r}/{b}/Seed43.extension/Seed43.tab/About.panel/Stack01.stack/Seed43.pushbutton/bundle.yaml".format(
    o=GITHUB_ORG, r=MAIN_REPO, b=BRANCH)

APPDATA = os.environ.get("APPDATA", "")

TOOLS = [
    {
        "id":            "pytransmit",
        "name":          "PyTransmit",
        "repo":          "Seed43-PyTransmit",
        "install_dir":   os.path.join(APPDATA, "pyRevit", "Extensions", "Seed43.extension",
                             "Seed43.tab", "Document Studio.panel", "pyTransmit.pushbutton"),
        "yaml_file":     os.path.join(APPDATA, "pyRevit", "Extensions", "Seed43.extension",
                             "Seed43.tab", "Document Studio.panel", "pyTransmit.pushbutton", "bundle.yaml"),
        "changelog_url": "https://raw.githubusercontent.com/{o}/{r}/{b}/bundle.yaml".format(
                             o=GITHUB_ORG, r="Seed43-PyTransmit", b=BRANCH),
        "zip_url":       "https://github.com/{o}/{r}/archive/refs/heads/{b}.zip".format(
                             o=GITHUB_ORG, r="Seed43-PyTransmit", b=BRANCH),
    }
]

S43_YAML_FILE    = os.path.join(APPDATA, "pyRevit", "Extensions", "Seed43.extension",
                     "Seed43.tab", "About.panel", "Stack01.stack", "Seed43.pushbutton", "bundle.yaml")


def _parse_yaml_version(path):
    """Parse version from a simple yaml file."""
    try:
        if File.Exists(path):
            for line in File.ReadAllText(path).splitlines():
                stripped = line.strip()
                if stripped.startswith("version:"):
                    return stripped.split(":", 1)[1].strip()
    except Exception:
        pass
    return None


def _write_yaml_version(path, version):
    """Write version field into a yaml file."""
    try:
        if File.Exists(path):
            lines = list(File.ReadAllLines(path))
            new_lines = []
            replaced = False
            for line in lines:
                if line.strip().startswith("version:") and not replaced:
                    new_lines.append("version: " + version)
                    replaced = True
                else:
                    new_lines.append(line)
            if not replaced:
                new_lines.append("version: " + version)
            File.WriteAllLines(path, new_lines)
        else:
            Directory.CreateDirectory(Path.GetDirectoryName(path))
            File.WriteAllText(path, "version: " + version + "\n")
    except Exception:
        pass

# ── Load XAML from file ───────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(__file__)
XAML_PATH  = os.path.join(SCRIPT_DIR, "seed43.xaml")

# ── Helpers ───────────────────────────────────────────────────────────────────

def load_xaml(path):
    reader = StreamReader(path)
    window = XamlReader.Load(reader.BaseStream)
    reader.Close()
    return window

def fetch_bundle(url):
    """Fetch and parse a simple bundle.yaml from a URL. Returns dict or None."""
    try:
        wc = WebClient()
        wc.Headers.Add("Cache-Control", "no-cache, no-store")
        wc.Headers.Add("Pragma", "no-cache")
        raw = wc.DownloadString(url)
        result = {}
        changelog = []
        in_changelog = False
        for line in raw.splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if stripped.startswith("- ") and in_changelog:
                changelog.append(stripped[2:].strip())
                continue
            if ":" in stripped:
                in_changelog = False
                key, _, val = stripped.partition(":")
                key = key.strip()
                val = val.strip()
                if key == "changelog":
                    in_changelog = True
                elif val:
                    result[key] = val
        if changelog:
            result["changelog"] = changelog
        # Map to expected keys
        result["version"] = result.get("version", "")
        result["changes"] = result.get("changelog", [])
        return result
    except Exception:
        return None

def fetch_json(url):
    """Legacy alias — routes to fetch_bundle for yaml URLs."""
    return fetch_bundle(url)

def get_local_version(yaml_file):
    return _parse_yaml_version(yaml_file)

def write_local_version(yaml_file, version):
    _write_yaml_version(yaml_file, version)

def _write_yaml_version_py(path, version):
    """Plain Python version of yaml version writer — used in update worker."""
    try:
        if os.path.exists(path):
            with open(path, "r") as f:
                lines = f.readlines()
            new_lines = []
            replaced = False
            for line in lines:
                if line.strip().startswith("version:") and not replaced:
                    new_lines.append("version: {}\n".format(version))
                    replaced = True
                else:
                    new_lines.append(line)
            if not replaced:
                new_lines.append("version: {}\n".format(version))
            with open(path, "w") as f:
                f.writelines(new_lines)
        else:
            with open(path, "w") as f:
                f.write("version: {}\n".format(version))
    except Exception:
        pass


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
        self.window.FindName("footer_reload").Click             += self._on_reload
        self.window.FindName("update_ribbon").MouseLeftButtonUp += self._on_s43_update
        self.window.FindName("pt_header").MouseLeftButtonUp     += self._on_toggle
        self.window.FindName("pt_action_btn").Click             += self._on_action

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
                if os.path.exists(tool["install_dir"]):
                    shutil.rmtree(tool["install_dir"])
                os.makedirs(tool["install_dir"])
                # Copy all files from repo root directly into pushbutton folder
                for item in os.listdir(extracted_root):
                    s = os.path.join(extracted_root, item)
                    d = os.path.join(tool["install_dir"], item)
                    if os.path.isdir(s):
                        shutil.copytree(s, d)
                    else:
                        shutil.copy2(s, d)

                if File.Exists(tmp_zip):
                    File.Delete(tmp_zip)
                if Directory.Exists(tmp_dir):
                    Directory.Delete(tmp_dir, True)

                remote  = fetch_json(tool["changelog_url"])
                version = remote.get("version", "unknown") if remote else "unknown"
                write_local_version(tool["yaml_file"], version)

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
            local_s43  = get_local_version(S43_YAML_FILE)
            remote_s43 = fetch_bundle(CHANGELOG_URL)
            dispatch(self.window, lambda: self._update_s43_ui(local_s43, remote_s43))
            local_pt  = get_local_version(self.tool["yaml_file"])
            remote_pt = fetch_bundle(self.tool["changelog_url"])
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

        self.window.FindName("update_ribbon").Visibility = Visibility.Collapsed

        EXTENSIONS_DIR = os.path.join(os.environ.get("APPDATA", ""), "pyRevit", "Extensions")
        S43_INSTALL    = os.path.join(EXTENSIONS_DIR, "Seed43.extension")
        ZIP_URL        = "https://github.com/{o}/{r}/archive/refs/heads/{b}.zip".format(
                            o=GITHUB_ORG, r=MAIN_REPO, b=BRANCH)
        TEMP_ZIP       = os.path.join(os.environ.get("TEMP", ""), "seed43_update.zip")
        TEMP_DIR       = os.path.join(os.environ.get("TEMP", ""), "seed43_update_extracted")

        def log(msg):
            dispatch(self.window, lambda: setattr(
                self.window.FindName("s43_changelog"), "Text", msg))

        def worker():
            try:
                log("Downloading update...")
                wc = WebClient()
                wc.Headers.Add("Cache-Control", "no-cache, no-store")
                wc.DownloadFile(ZIP_URL, TEMP_ZIP)

                log("Extracting...")
                if os.path.exists(TEMP_DIR):
                    shutil.rmtree(TEMP_DIR)
                os.makedirs(TEMP_DIR)
                with zipfile.ZipFile(TEMP_ZIP, "r") as z:
                    z.extractall(TEMP_DIR)

                # Find extracted root folder
                extracted_root = None
                for item in os.listdir(TEMP_DIR):
                    full = os.path.join(TEMP_DIR, item)
                    if os.path.isdir(full):
                        extracted_root = full
                        break
                if not extracted_root:
                    raise Exception("Could not find extracted folder.")

                # Find Seed43.extension inside extracted root
                src = os.path.join(extracted_root, "Seed43.extension")
                if not os.path.exists(src):
                    raise Exception("Seed43.extension not found in ZIP. Found: " + str(os.listdir(extracted_root)))

                log("Installing update...")

                # Only update files inside About.panel — nothing else is touched
                src_about = os.path.join(src, "Seed43.tab", "About.panel")
                dst_about = os.path.join(S43_INSTALL, "Seed43.tab", "About.panel")

                if not os.path.exists(src_about):
                    raise Exception("About.panel not found in repo ZIP.")

                def sync_tree(src_dir, dst_dir):
                    if not os.path.exists(dst_dir):
                        os.makedirs(dst_dir)
                    src_items = set(os.listdir(src_dir))
                    dst_items = set(os.listdir(dst_dir)) if os.path.exists(dst_dir) else set()
                    for item in src_items:
                        s = os.path.join(src_dir, item)
                        d = os.path.join(dst_dir, item)
                        if os.path.isdir(s):
                            sync_tree(s, d)
                        else:
                            shutil.copy2(s, d)
                    for item in dst_items - src_items:
                        d = os.path.join(dst_dir, item)
                        if os.path.isdir(d):
                            shutil.rmtree(d)
                            log("Removed: {0}".format(item))
                        else:
                            os.remove(d)
                            log("Removed: {0}".format(item))

                sync_tree(src_about, dst_about)

                # Fetch and write version into bundle.yaml in script folder
                remote  = fetch_bundle(CHANGELOG_URL)
                version = remote.get("version", "unknown") if remote else "unknown"

                bundle_path = os.path.join(
                    S43_INSTALL, "Seed43.tab", "About.panel",
                    "Stack01.stack", "Seed43.pushbutton", "bundle.yaml")
                _write_yaml_version_py(bundle_path, version)

                log("Done — v{0}".format(version))

                # Cleanup
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
        self.window.FindName("s43_changelog").Text = u"Updated to v{0} \u2014 reloading PyRevit...".format(version)
        result = MessageBox.Show(
            "Seed43 updated to v{0}.\n\nPyRevit will now reload to apply the update.".format(version),
            "Seed43 Updated",
            MessageBoxButton.OK,
            MessageBoxImage.Information
        )
        # Close dialog then reload PyRevit
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
