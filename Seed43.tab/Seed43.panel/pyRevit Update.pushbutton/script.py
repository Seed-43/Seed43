# -*- coding: utf-8 -*-
# "pyRevit Update"
# "Seed43"
# """
# Update a pyRevit that was installed with the Windows installer. pyRevit's
# own Update button only updates git clones, and it gives up when a firewall
# blocks its 8.8.8.8:53 connection test. This asks GitHub directly for the
# latest release, and if it is newer, hands a background PowerShell helper
# the job of downloading the installer and starting it once Revit closes.
# """

import os
import json
import clr

clr.AddReference("System")
from System.Net import WebClient, ServicePointManager, SecurityProtocolType
from System.Diagnostics import Process, ProcessStartInfo, ProcessWindowStyle
from Microsoft.Win32 import Registry

import pyrevit
from pyrevit import forms

try:
    from Snippets import _dialogs as sdlg
except Exception:
    sdlg = None


# ── CONSTANTS ──────────────────────────────────────────────────────────────

TITLE = "pyRevit Update"
RELEASE_API = "https://api.github.com/repos/pyrevitlabs/pyRevit/releases/latest"
RELEASES_PAGE = "https://github.com/pyrevitlabs/pyRevit/releases/latest"
HELPER_PS1 = os.path.join(os.path.dirname(__file__), "install_when_closed.ps1")
WORK_DIR = os.path.join(os.environ.get("TEMP", os.path.expanduser("~")),
                        "Seed43", "pyRevitUpdate")
PENDING_FILE = os.path.join(WORK_DIR, "pending.json")
UNINSTALL_KEY = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"


# ── DIALOGS ────────────────────────────────────────────────────────────────

def _alert(text):
    """Themed OK popup, falling back to forms.alert without the shared lib."""
    if sdlg:
        sdlg.message(text, title=TITLE)
    else:
        forms.alert(text, title=TITLE)


def _choose(text, options):
    """Themed multi-button popup. options is a list of (key, label)."""
    if sdlg:
        return sdlg.choice(text, options, title=TITLE)
    labels = [label for _, label in options]
    picked = forms.alert(text, title=TITLE, options=labels)
    for key, label in options:
        if label == picked:
            return key
    return None


# ── HELPERS ────────────────────────────────────────────────────────────────

def _version_tuple(text):
    """'v6.5.5.26237+2044' -> (6, 5, 5, 26237). Unparseable -> (0,)."""
    core = text.strip().lstrip("vV").split("+")[0]
    try:
        return tuple(int(x) for x in core.split("."))
    except ValueError:
        return (0,)


def _installed_version():
    """Read pyRevit's own version file, e.g. '6.4.0.26100+0515'."""
    path = os.path.join(pyrevit.HOME_DIR, "pyrevitlib", "pyrevit", "version")
    try:
        with open(path) as f:
            return f.read().strip()
    except Exception:
        return None


def _is_admin_install():
    """An all-users install lives under Program Files and needs the admin installer."""
    home = pyrevit.HOME_DIR.lower()
    return "program files" in home


def _fetch_via_revit():
    """Ask GitHub from inside Revit's own process."""
    # NOTE: Revit's .NET may default to TLS 1.0/1.1, which GitHub refuses.
    ServicePointManager.SecurityProtocol = (ServicePointManager.SecurityProtocol
                                            | SecurityProtocolType.Tls12)
    client = WebClient()
    client.Headers.Add("User-Agent", "Seed43-pyRevitUpdate")  # GitHub API rejects requests without one
    client.Headers.Add("Accept", "application/vnd.github+json")
    return client.DownloadString(RELEASE_API)


def _fetch_via_powershell():
    """Ask GitHub through a hidden powershell.exe instead."""
    command = ("$ProgressPreference='SilentlyContinue';"
               "[Net.ServicePointManager]::SecurityProtocol="
               "[Net.ServicePointManager]::SecurityProtocol -bor "
               "[Net.SecurityProtocolType]::Tls12;"
               "(Invoke-WebRequest -UseBasicParsing -Uri '{}' "
               "-Headers @{{'User-Agent'='Seed43-pyRevitUpdate'}}).Content").format(RELEASE_API)
    info = ProcessStartInfo("powershell.exe",
                            '-NoProfile -NonInteractive -Command "{}"'.format(command))
    info.UseShellExecute = False
    info.CreateNoWindow = True
    info.RedirectStandardOutput = True
    info.RedirectStandardError = True
    proc = Process.Start(info)
    output = proc.StandardOutput.ReadToEnd()
    error = proc.StandardError.ReadToEnd()
    if not proc.WaitForExit(30000) or proc.ExitCode != 0 or not output.strip():
        raise Exception(error.strip() or "PowerShell returned nothing")
    return output


def _fetch_latest_release():
    """Return the GitHub 'latest release' JSON as a dict.

    Tries Revit first, then PowerShell. HACK: an app firewall (Portmaster on
    Fred's PC) can block Revit.exe from the internet while allowing
    powershell.exe, which also does the download, so the check follows it.
    """
    try:
        raw = _fetch_via_revit()
    except Exception as revit_error:
        try:
            raw = _fetch_via_powershell()
        except Exception as ps_error:
            raise Exception("Revit: {}\n\nPowerShell: {}".format(revit_error, ps_error))
    return json.loads(raw)


def _pick_installer(release, admin):
    """Find the main (not CLI) signed .exe installer for this install type."""
    for asset in release.get("assets", []):
        name = asset.get("name", "")
        low = name.lower()
        if not (low.startswith("pyrevit_") and low.endswith("_signed.exe")):
            continue
        if "_cli_" in low:
            continue
        if ("_admin_" in low) == admin:
            return asset
    return None


# --- Stale uninstall entries ---
def _norm_dir(path):
    return os.path.normcase(os.path.normpath((path or "").strip().strip('"')))


def _install_entries():
    """Every per-user 'pyRevit version X' uninstall entry for THIS pyRevit folder.

    Returns a list of dicts: key, name, version, uninstaller.
    """
    entries = []
    home = _norm_dir(pyrevit.HOME_DIR)
    root = Registry.CurrentUser.OpenSubKey(UNINSTALL_KEY)
    if root is None:
        return entries
    try:
        for sub in root.GetSubKeyNames():
            key = root.OpenSubKey(sub)
            if key is None:
                continue
            try:
                name = key.GetValue("DisplayName") or ""
                if not name.startswith("pyRevit version"):
                    continue
                if _norm_dir(key.GetValue("InstallLocation")) != home:
                    continue
                entries.append({
                    "key": sub,
                    "name": name,
                    "version": key.GetValue("DisplayVersion") or "",
                    "uninstaller": (key.GetValue("UninstallString") or "").strip().strip('"'),
                })
            finally:
                key.Close()
    finally:
        root.Close()
    return entries


def _clean_stale_entries(installed):
    """Remove apps-list entries left behind by older installs over this folder.

    Installing a new pyRevit over an old one leaves the old entry in Windows'
    apps list, pointing at the SAME folder with its own unins00N.exe. Running
    that old uninstaller would delete the current pyRevit's files, so the
    entry and its unins00N.exe/.dat are removed instead. Nothing is touched
    unless the entry for the running version is found, and only entries with
    a lower version and an uninstaller inside this pyRevit folder qualify.

    Returns a list of 'name' strings that were removed.
    """
    if not installed:
        return []
    current = _version_tuple(installed)[:4]
    home = _norm_dir(pyrevit.HOME_DIR)
    entries = _install_entries()
    keep = [e for e in entries if _version_tuple(e["version"])[:4] == current]
    if not keep:
        return []  # can't tell which uninstaller is live, so leave everything
    keep_files = set(_norm_dir(e["uninstaller"]) for e in keep)

    removed = []
    for e in entries:
        if _version_tuple(e["version"]) >= current:
            continue
        exe = _norm_dir(e["uninstaller"])
        if os.path.dirname(exe) != home or exe in keep_files:
            continue
        Registry.CurrentUser.DeleteSubKeyTree(UNINSTALL_KEY + "\\" + e["key"])
        for path in (e["uninstaller"], os.path.splitext(e["uninstaller"])[0] + ".dat"):
            try:
                if os.path.isfile(path):
                    os.remove(path)
            except Exception:
                pass  # the apps-list entry is the dangerous part and it is gone
        removed.append(e["name"])
    return removed


# --- Pending update ---
def _process_alive(pid):
    try:
        return not Process.GetProcessById(int(pid)).HasExited
    except Exception:
        return False


def _pending_update():
    """The update already queued by an earlier run, if its helper is still running."""
    try:
        with open(PENDING_FILE) as f:
            data = json.load(f)
        if _process_alive(data.get("pid", -1)):
            return data
    except Exception:
        pass
    return None


def _start_helper(asset, version):
    """Launch the hidden PowerShell helper and record its PID. Returns the PID."""
    if not os.path.isdir(WORK_DIR):
        os.makedirs(WORK_DIR)
    out_file = os.path.join(WORK_DIR, asset["name"])
    args = ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{}" '
            '-Url "{}" -Size {} -OutFile "{}" -Version "{}"').format(
                HELPER_PS1, asset["browser_download_url"], asset["size"],
                out_file, version)
    info = ProcessStartInfo("powershell.exe", args)
    info.UseShellExecute = True  # detached, so it outlives Revit
    info.WindowStyle = ProcessWindowStyle.Hidden
    proc = Process.Start(info)
    with open(PENDING_FILE, "w") as f:
        json.dump({"pid": proc.Id, "version": version}, f)
    return proc.Id


# ── UI / ENTRY POINT ───────────────────────────────────────────────────────

def main():
    if os.path.isdir(os.path.join(pyrevit.HOME_DIR, ".git")):
        _alert("This pyRevit is a git clone, not an installer install.\n\n"
               "Use pyRevit's own Update button, or the pyRevit CLI.")
        return

    pending = _pending_update()
    if pending:
        _alert("pyRevit {} is already queued.\n\nClose every Revit window and "
               "the installer will start.".format(pending.get("version", "")))
        return

    installed = _installed_version()
    try:
        removed = _clean_stale_entries(installed)
    except Exception as ex:
        removed = []
        _alert("Could not tidy old pyRevit entries from the apps list.\n\n{}".format(ex))
    cleanup_note = ""
    if removed:
        cleanup_note = ("\n\nRemoved from the Windows apps list (left over from "
                        "an earlier install, sharing this folder):\n- {}").format(
                            "\n- ".join(removed))

    try:
        release = _fetch_latest_release()
    except Exception as ex:
        _alert("Could not reach GitHub to check for updates.\n\n{}{}".format(
            ex, cleanup_note))
        return

    latest_tag = release.get("tag_name", "")
    latest = latest_tag.lstrip("vV").split("+")[0]
    current = (installed or "unknown").split("+")[0]

    if installed and _version_tuple(installed) >= _version_tuple(latest_tag):
        _alert("pyRevit is up to date.\n\nInstalled: {}\nLatest: {}{}".format(
            current, latest, cleanup_note))
        return

    if cleanup_note:
        _alert(cleanup_note.strip())

    admin = _is_admin_install()
    asset = _pick_installer(release, admin)
    if asset is None:
        _alert("pyRevit {} is out, but its {}installer was not found in the "
               "release.\n\n{}".format(latest, "admin " if admin else "",
                                        RELEASES_PAGE))
        return

    size_mb = int(round(asset["size"] / (1024.0 * 1024.0)))
    pick = _choose(
        "pyRevit {} is available. You have {}.\n\n"
        "Install downloads {} ({} MB) in the background. The installer "
        "opens by itself once every Revit window is closed, since pyRevit "
        "can't be replaced while Revit has it loaded. Your extensions and "
        "settings are kept.".format(latest, current, asset["name"], size_mb),
        [("cancel", "Cancel"), ("page", "Release notes"), ("install", "Install")])

    if pick == "page":
        os.startfile(RELEASES_PAGE)  # Process.Start(url) fails on Revit 2025+ (.NET 8)
    elif pick == "install":
        try:
            _start_helper(asset, latest)
        except Exception as ex:
            _alert("Could not start the update helper.\n\n{}".format(ex))
            return
        _alert("Downloading pyRevit {} now.\n\nKeep working. When you close "
               "Revit, the installer will open. If the download fails you'll "
               "get a message, and a log is kept in:\n{}".format(latest, WORK_DIR))


main()
