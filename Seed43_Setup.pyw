"""
Seed43 Setup
Installs the Seed43 PyRevit extension into:
  %APPDATA%/pyRevit/Extensions/seed43

Downloads directly from GitHub - no Git installation required.
Double-click Seed43_Setup.pyw to run.
"""

import os
import math
import threading
import zipfile
import shutil
import urllib.request
import json
import tkinter as tk
from tkinter import messagebox

# ── Constants ─────────────────────────────────────────────────────────────────
GITHUB_ORG    = "Seed-43"
MAIN_REPO     = "Seed43"
BRANCH        = "main"
ZIP_URL       = "https://github.com/{o}/{r}/archive/refs/heads/{b}.zip".format(
                    o=GITHUB_ORG, r=MAIN_REPO, b=BRANCH)
CHANGELOG_URL = "https://raw.githubusercontent.com/{o}/{r}/{b}/changelog.json".format(
                    o=GITHUB_ORG, r=MAIN_REPO, b=BRANCH)

EXTENSIONS_DIR = os.path.join(os.environ.get("APPDATA", ""), "pyRevit", "Extensions")
INSTALL_DIR    = os.path.join(EXTENSIONS_DIR, "Seed43.extension")
VERSION_FILE   = os.path.join(INSTALL_DIR, "version.txt")
PUSHBUTTON_DIR = os.path.join(
    INSTALL_DIR,
    "Seed43.tab", "About.panel",
    "Stack01.stack", "Seed43.pushbutton"
)
TEMP_DIR      = os.environ.get("TEMP", os.environ.get("TMP", ""))
TEMP_ZIP      = os.path.join(TEMP_DIR, "seed43_install.zip")
TEMP_EXTRACT  = os.path.join(TEMP_DIR, "seed43_extracted")

# ── Colours ───────────────────────────────────────────────────────────────────
C_WIN_BG          = "#3B4553"
C_CARD_BG         = "#2B3340"
C_HEADER_BG       = "#232933"
C_GREEN           = "#208A3C"
C_GREEN_HOVER     = "#2B933F"
C_GREEN_PRESSED   = "#1A6E2E"
C_TEXT            = "#F4FAFF"
C_TEXT_DIM        = "#A0AABB"
C_INPUT_BG        = "#F4FAFF"
C_INPUT_FG        = "#2B3340"
C_SECONDARY_BTN   = "#404553"
C_SECONDARY_HOVER = "#4E5566"
C_DANGER          = "#C53030"
FONT              = "Segoe UI"


# ── Helpers ───────────────────────────────────────────────────────────────────

def get_installed_version():
    if os.path.exists(VERSION_FILE):
        with open(VERSION_FILE, "r") as f:
            return f.read().strip()
    return None


def write_version(version):
    os.makedirs(os.path.dirname(VERSION_FILE), exist_ok=True)
    with open(VERSION_FILE, "w") as f:
        f.write(version)


def fetch_changelog():
    try:
        with urllib.request.urlopen(CHANGELOG_URL, timeout=8) as r:
            return json.loads(r.read().decode())
    except Exception:
        return None


def download_and_install(log_fn, done_fn, error_fn):
    """Runs in a background thread — downloads and installs the extension."""
    try:
        log_fn("Connecting to GitHub...")
        urllib.request.urlretrieve(ZIP_URL, TEMP_ZIP)

        log_fn("Extracting files...")
        if os.path.exists(TEMP_EXTRACT):
            shutil.rmtree(TEMP_EXTRACT)
        with zipfile.ZipFile(TEMP_ZIP, "r") as z:
            z.extractall(TEMP_EXTRACT)

        # GitHub ZIPs extract to reponame-branch/
        extracted_root = None
        for item in os.listdir(TEMP_EXTRACT):
            full = os.path.join(TEMP_EXTRACT, item)
            if os.path.isdir(full):
                extracted_root = full
                break

        if not extracted_root:
            raise Exception("Could not find extracted folder.")

        log_fn("Installing extension files...")
        # We only want the Seed43.extension subfolder from the repo ZIP
        src = os.path.join(extracted_root, "Seed43.extension")
        if not os.path.exists(src):
            raise Exception("Seed43.extension folder not found in repo ZIP.")
        if os.path.exists(INSTALL_DIR):
            shutil.rmtree(INSTALL_DIR)
        shutil.copytree(src, INSTALL_DIR)

        # Ensure pushbutton folder exists and script files are in place
        log_fn("Installing About button scripts...")
        os.makedirs(PUSHBUTTON_DIR, exist_ok=True)

        # script.py and seed43.xaml are already inside the repo ZIP
        # under Seed43.extension/.../Seed43.pushbutton/ — nothing extra needed.
        # This step confirms they landed correctly.
        script_ok = os.path.exists(os.path.join(PUSHBUTTON_DIR, "script.py"))
        xaml_ok   = os.path.exists(os.path.join(PUSHBUTTON_DIR, "seed43.xaml"))

        if not script_ok or not xaml_ok:
            log_fn("Warning: script.py or seed43.xaml missing from repo ZIP.")
            log_fn("Check the repo includes the Seed43.extension folder.")

        # Write version — stored outside the extension folder so it survives updates
        changelog = fetch_changelog()
        version = changelog.get("version", "unknown") if changelog else "unknown"
        write_version(version)
        # Also write backup version file at Extensions level
        backup_version = os.path.join(EXTENSIONS_DIR, "seed43_version.txt")
        with open(backup_version, "w") as f:
            f.write(version)

        # Cleanup
        if os.path.exists(TEMP_ZIP):
            os.remove(TEMP_ZIP)
        if os.path.exists(TEMP_EXTRACT):
            shutil.rmtree(TEMP_EXTRACT)

        done_fn(version)

    except Exception as e:
        error_fn(str(e))


def uninstall(log_fn, done_fn, error_fn):
    """Runs in a background thread — removes the extension."""
    try:
        log_fn("Removing files...")
        if os.path.exists(INSTALL_DIR):
            shutil.rmtree(INSTALL_DIR)
        log_fn("Cleaning up...")
        done_fn()
    except Exception as e:
        error_fn(str(e))


# ── Widgets ───────────────────────────────────────────────────────────────────

class HoverButton(tk.Button):
    def __init__(self, master, bg_normal, bg_hover, bg_press=None, **kwargs):
        self._n = bg_normal
        self._h = bg_hover
        self._p = bg_press or bg_hover
        super().__init__(master, bg=bg_normal, activebackground=self._p,
                         relief="flat", bd=0, cursor="hand2", **kwargs)
        self.bind("<Enter>",           lambda e: self.config(bg=self._h))
        self.bind("<Leave>",           lambda e: self.config(bg=self._n))
        self.bind("<Button-1>",        lambda e: self.config(bg=self._p))
        self.bind("<ButtonRelease-1>", lambda e: self.config(bg=self._h))

    def set_normal(self, bg, bg_hover, bg_press=None):
        self._n = bg
        self._h = bg_hover
        self._p = bg_press or bg_hover
        self.config(bg=bg, activebackground=self._p)


class SectionLabel(tk.Label):
    def __init__(self, master, text, **kw):
        super().__init__(master, text=text,
                         font=(FONT, 10, "bold"),
                         fg=C_GREEN, bg=kw.pop("bg", C_CARD_BG), **kw)


class FieldLabel(tk.Label):
    def __init__(self, master, text, **kw):
        super().__init__(master, text=text,
                         font=(FONT, 9),
                         fg=C_TEXT_DIM, bg=kw.pop("bg", C_CARD_BG), **kw)


class Card(tk.Frame):
    def __init__(self, master, **kw):
        super().__init__(master, bg=C_CARD_BG, padx=16, pady=14,
                         highlightbackground=C_GREEN, highlightthickness=1, **kw)


# ── Main window ───────────────────────────────────────────────────────────────

class Seed43Setup(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Seed43 Setup")
        self.resizable(False, False)
        self.configure(bg=C_WIN_BG)
        self._installing = False
        self._installed  = False
        self._center(500, 560)
        self._build()
        self._check_status()

    def _center(self, w, h):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry("{w}x{h}+{x}+{y}".format(
            w=w, h=h, x=(sw - w) // 2, y=(sh - h) // 2))

    # ── Build UI ──────────────────────────────────────────────────────────────

    def _build(self):

        # Header
        header = tk.Frame(self, bg=C_HEADER_BG, height=64)
        header.pack(fill="x")
        header.pack_propagate(False)
        wm = tk.Frame(header, bg=C_HEADER_BG)
        wm.place(relx=0.5, rely=0.5, anchor="center")
        tk.Label(wm, text="Seed", font=(FONT, 22, "bold"),
                 fg=C_GREEN,  bg=C_HEADER_BG).pack(side="left")
        tk.Label(wm, text="43",   font=(FONT, 22),
                 fg=C_TEXT,   bg=C_HEADER_BG).pack(side="left")
        tk.Label(wm, text="  |  Setup", font=(FONT, 14),
                 fg=C_TEXT,   bg=C_HEADER_BG).pack(side="left")

        # Body
        body = tk.Frame(self, bg=C_WIN_BG, padx=20, pady=16)
        body.pack(fill="both", expand=True)

        # Status card
        status_card = Card(body)
        status_card.pack(fill="x", pady=(0, 12))
        SectionLabel(status_card, "Installation Status").pack(anchor="w")
        row = tk.Frame(status_card, bg=C_CARD_BG)
        row.pack(fill="x", pady=(8, 0))
        self._dot = tk.Label(row, text="●", font=(FONT, 10),
                             fg=C_TEXT_DIM, bg=C_CARD_BG)
        self._dot.pack(side="left")
        self._status_lbl = tk.Label(row, text="Checking...",
                                    font=(FONT, 9), fg=C_TEXT_DIM, bg=C_CARD_BG)
        self._status_lbl.pack(side="left", padx=(6, 0))
        self._version_lbl = tk.Label(status_card, text="",
                                     font=(FONT, 9), fg=C_TEXT_DIM, bg=C_CARD_BG)
        self._version_lbl.pack(anchor="w", pady=(4, 0))

        # Install path card
        path_card = Card(body)
        path_card.pack(fill="x", pady=(0, 12))
        SectionLabel(path_card, "Install Location").pack(anchor="w")
        FieldLabel(path_card, "PyRevit Extensions folder", pady=4).pack(anchor="w")
        path_bg = tk.Frame(path_card, bg=C_INPUT_BG, padx=8, pady=5,
                           highlightbackground=C_GREEN, highlightthickness=1)
        path_bg.pack(fill="x", pady=(4, 0))
        tk.Label(path_bg, text=INSTALL_DIR, font=(FONT, 8),
                 fg=C_INPUT_FG, bg=C_INPUT_BG, anchor="w").pack(fill="x")

        # Log card
        log_card = Card(body)
        log_card.pack(fill="x", pady=(0, 12))
        SectionLabel(log_card, "Activity Log").pack(anchor="w", pady=(0, 6))
        log_bg = tk.Frame(log_card, bg=C_HEADER_BG,
                          highlightbackground=C_GREEN, highlightthickness=1)
        log_bg.pack(fill="x")
        self._log = tk.Text(log_bg, height=5, bg=C_HEADER_BG, fg=C_TEXT,
                            font=(FONT, 8), bd=0, relief="flat",
                            state="disabled", padx=8, pady=6,
                            selectbackground=C_GREEN, wrap="word")
        self._log.pack(fill="x")

        # Progress bar
        self._prog_frame = tk.Frame(body, bg=C_WIN_BG)
        prog_bg = tk.Frame(self._prog_frame, bg=C_CARD_BG, height=4,
                           highlightbackground=C_GREEN, highlightthickness=1)
        prog_bg.pack(fill="x")
        self._prog_fill = tk.Frame(prog_bg, bg=C_GREEN, height=4)
        self._prog_fill.place(relwidth=0, relheight=1)

        # Buttons
        btn_row = tk.Frame(body, bg=C_WIN_BG)
        btn_row.pack(fill="x", pady=(12, 4))

        self._action_btn = HoverButton(
            btn_row,
            bg_normal=C_GREEN, bg_hover=C_GREEN_HOVER, bg_press=C_GREEN_PRESSED,
            text="Install", font=(FONT, 10, "bold"), fg=C_TEXT,
            width=14, height=1, command=self._on_action)
        self._action_btn.pack(side="left")

        HoverButton(
            btn_row,
            bg_normal=C_SECONDARY_BTN, bg_hover=C_SECONDARY_HOVER,
            text="Close", font=(FONT, 9), fg=C_TEXT,
            width=10, height=1, command=self.destroy).pack(side="right")

    # ── Status check ──────────────────────────────────────────────────────────

    def _check_status(self):
        version = get_installed_version()
        if version:
            self._installed = True
            self._set_status("installed", "Installed — v{0}".format(version))
            self._version_lbl.config(text="Location: {0}".format(INSTALL_DIR))
            self._action_btn.config(text="Uninstall")
            self._action_btn.set_normal(C_DANGER, "#a82828")
            self._log_line("Existing installation found: v{0}".format(version))
            self._log_line("Click Uninstall to remove, or close to exit.")
        else:
            self._set_status("not_installed", "Not installed")
            self._log_line("No existing installation detected.")
            self._log_line("Will install to: {0}".format(INSTALL_DIR))

    def _set_status(self, state, text):
        colours = {
            "installed":     C_GREEN,
            "not_installed": "#888888",
            "busy":          "#F0A500",
            "error":         C_DANGER,
        }
        c = colours.get(state, C_TEXT_DIM)
        self._dot.config(fg=c)
        self._status_lbl.config(text=text, fg=c)

    # ── Log ───────────────────────────────────────────────────────────────────

    def _log_line(self, msg):
        self._log.config(state="normal")
        self._log.insert("end", "> {0}\n".format(msg))
        self._log.see("end")
        self._log.config(state="disabled")

    # ── Progress animation ────────────────────────────────────────────────────

    def _animate(self, val):
        if not self._installing:
            return
        t = (math.sin(val * 0.15) + 1) / 2
        self._prog_fill.place(relwidth=max(0.08, t), relheight=1)
        self.after(30, self._animate, val + 1)

    # ── Action button ─────────────────────────────────────────────────────────

    def _on_action(self):
        if self._installed:
            if messagebox.askyesno(
                "Uninstall Seed43",
                "This will delete the Seed43 extension from:\n\n{0}\n\nContinue?".format(INSTALL_DIR)
            ):
                self._run_uninstall()
        else:
            self._run_install()

    # ── Install ───────────────────────────────────────────────────────────────

    def _run_install(self):
        self._action_btn.config(state="disabled")
        self._set_status("busy", "Installing...")
        self._prog_frame.pack(fill="x", pady=(8, 0))
        self._installing = True
        self._animate(0)

        threading.Thread(
            target=download_and_install,
            args=(
                lambda m: self.after(0, self._log_line, m),
                lambda v: self.after(0, self._on_install_done, v),
                lambda e: self.after(0, self._on_error, e),
            ),
            daemon=True
        ).start()

    def _on_install_done(self, version):
        self._installing = False
        self._installed  = True
        self._prog_fill.place(relwidth=1, relheight=1)
        self._set_status("installed", "Installed — v{0}".format(version))
        self._version_lbl.config(text="Location: {0}".format(INSTALL_DIR))
        self._log_line("Installation complete — v{0}".format(version))
        self._log_line("Reload PyRevit in Revit to activate the Seed43 tab.")
        self._action_btn.config(state="normal", text="Uninstall")
        self._action_btn.set_normal(C_DANGER, "#a82828")
        messagebox.showinfo(
            "Seed43 Setup",
            "Installation complete — v{0}\n\nReload PyRevit inside Revit to see the Seed43 tab.".format(version)
        )

    # ── Uninstall ─────────────────────────────────────────────────────────────

    def _run_uninstall(self):
        self._action_btn.config(state="disabled")
        self._set_status("busy", "Uninstalling...")
        self._prog_frame.pack(fill="x", pady=(8, 0))
        self._installing = True
        self._animate(0)

        threading.Thread(
            target=uninstall,
            args=(
                lambda m: self.after(0, self._log_line, m),
                lambda:   self.after(0, self._on_uninstall_done),
                lambda e: self.after(0, self._on_error, e),
            ),
            daemon=True
        ).start()

    def _on_uninstall_done(self):
        self._installing = False
        self._installed  = False
        self._prog_fill.place(relwidth=0, relheight=1)
        self._prog_frame.pack_forget()
        self._set_status("not_installed", "Not installed")
        self._version_lbl.config(text="")
        self._log_line("Uninstall complete.")
        self._action_btn.config(state="normal", text="Install")
        self._action_btn.set_normal(C_GREEN, C_GREEN_HOVER, C_GREEN_PRESSED)

    # ── Error ─────────────────────────────────────────────────────────────────

    def _on_error(self, err):
        self._installing = False
        self._set_status("error", "Failed")
        self._log_line("Error: {0}".format(err))
        self._action_btn.config(state="normal")
        messagebox.showerror("Seed43 Setup", "Operation failed:\n\n{0}".format(err))


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = Seed43Setup()
    app.mainloop()
