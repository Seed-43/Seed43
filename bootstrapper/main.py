"""
Seed43 Bootstrapper
Installs the Seed43 PyRevit extension into %APPDATA%/pyRevit/Extensions/seed43
Downloads directly from GitHub - no Git installation required.
"""

import os
import sys
import threading
import zipfile
import shutil
import urllib.request
import json
import tkinter as tk
from tkinter import messagebox

# ── Constants ────────────────────────────────────────────────────────────────
GITHUB_ORG        = "Seed-43"
MAIN_REPO         = "Seed43"
BRANCH            = "main"
ZIP_URL           = f"https://github.com/{GITHUB_ORG}/{MAIN_REPO}/archive/refs/heads/{BRANCH}.zip"
CHANGELOG_URL     = f"https://raw.githubusercontent.com/{GITHUB_ORG}/{MAIN_REPO}/{BRANCH}/changelog.json"

INSTALL_DIR       = os.path.join(os.environ.get("APPDATA", ""), "pyRevit", "Extensions", "seed43")
VERSION_FILE      = os.path.join(INSTALL_DIR, "version.txt")
TEMP_ZIP          = os.path.join(os.environ.get("TEMP", ""), "seed43_install.zip")

# ── Colours (matching pyTransmit style guide) ────────────────────────────────
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

FONT_FAMILY       = "Segoe UI"


# ══════════════════════════════════════════════════════════════════════════════
# Helpers
# ══════════════════════════════════════════════════════════════════════════════

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
    """Runs in a background thread."""
    try:
        log_fn("Connecting to GitHub…")
        urllib.request.urlretrieve(ZIP_URL, TEMP_ZIP)

        log_fn("Extracting files…")
        with zipfile.ZipFile(TEMP_ZIP, "r") as z:
            z.extractall(os.environ.get("TEMP", ""))

        # The ZIP extracts to e.g. Seed43-main/
        extracted_root = os.path.join(
            os.environ.get("TEMP", ""),
            f"{MAIN_REPO}-{BRANCH}"
        )

        log_fn("Installing to PyRevit extensions folder…")
        if os.path.exists(INSTALL_DIR):
            shutil.rmtree(INSTALL_DIR)
        shutil.copytree(extracted_root, INSTALL_DIR)

        # Fetch changelog for version tag
        changelog = fetch_changelog()
        version = changelog.get("version", "unknown") if changelog else "unknown"
        write_version(version)

        # Cleanup
        if os.path.exists(TEMP_ZIP):
            os.remove(TEMP_ZIP)
        if os.path.exists(extracted_root):
            shutil.rmtree(extracted_root)

        done_fn(version)

    except Exception as e:
        error_fn(str(e))


# ══════════════════════════════════════════════════════════════════════════════
# Widgets
# ══════════════════════════════════════════════════════════════════════════════

class HoverButton(tk.Button):
    """Button that swaps background on hover — matches pyTransmit style."""

    def __init__(self, master, bg_normal, bg_hover, bg_press=None, **kwargs):
        self._bg_normal = bg_normal
        self._bg_hover  = bg_hover
        self._bg_press  = bg_press or bg_hover
        super().__init__(master, bg=bg_normal, activebackground=self._bg_press,
                         relief="flat", bd=0, cursor="hand2", **kwargs)
        self.bind("<Enter>",    lambda e: self.config(bg=self._bg_hover))
        self.bind("<Leave>",    lambda e: self.config(bg=self._bg_normal))
        self.bind("<Button-1>", lambda e: self.config(bg=self._bg_press))
        self.bind("<ButtonRelease-1>", lambda e: self.config(bg=self._bg_hover))


class SectionLabel(tk.Label):
    def __init__(self, master, text, **kwargs):
        super().__init__(
            master, text=text,
            font=(FONT_FAMILY, 10, "bold"),
            fg=C_GREEN, bg=kwargs.pop("bg", C_CARD_BG),
            **kwargs
        )


class FieldLabel(tk.Label):
    def __init__(self, master, text, **kwargs):
        super().__init__(
            master, text=text,
            font=(FONT_FAMILY, 9),
            fg=C_TEXT_DIM, bg=kwargs.pop("bg", C_CARD_BG),
            **kwargs
        )


class Card(tk.Frame):
    def __init__(self, master, **kwargs):
        super().__init__(
            master,
            bg=C_CARD_BG,
            padx=16, pady=14,
            highlightbackground=C_GREEN,
            highlightthickness=1,
            **kwargs
        )


# ══════════════════════════════════════════════════════════════════════════════
# Main Window
# ══════════════════════════════════════════════════════════════════════════════

class Seed43Installer(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Seed43 Setup")
        self.resizable(False, False)
        self.configure(bg=C_WIN_BG)
        self._center(500, 560)
        self._build()
        self._check_status()

    def _center(self, w, h):
        self.geometry(f"{w}x{h}")
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x  = (sw - w) // 2
        y  = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build(self):
        # Header
        header = tk.Frame(self, bg=C_HEADER_BG, height=64)
        header.pack(fill="x")
        header.pack_propagate(False)

        wordmark = tk.Frame(header, bg=C_HEADER_BG)
        wordmark.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(wordmark, text="Seed", font=(FONT_FAMILY, 22, "bold"),
                 fg=C_GREEN, bg=C_HEADER_BG).pack(side="left")
        tk.Label(wordmark, text="43", font=(FONT_FAMILY, 22),
                 fg=C_TEXT, bg=C_HEADER_BG).pack(side="left")
        tk.Label(wordmark, text="  |  Setup", font=(FONT_FAMILY, 14),
                 fg=C_TEXT, bg=C_HEADER_BG).pack(side="left")

        # Body
        body = tk.Frame(self, bg=C_WIN_BG, padx=20, pady=16)
        body.pack(fill="both", expand=True)

        # Status card
        self._status_card = Card(body)
        self._status_card.pack(fill="x", pady=(0, 12))

        SectionLabel(self._status_card, "Installation Status").pack(anchor="w")

        status_row = tk.Frame(self._status_card, bg=C_CARD_BG)
        status_row.pack(fill="x", pady=(8, 0))

        self._status_dot = tk.Label(status_row, text="●", font=(FONT_FAMILY, 10),
                                    fg=C_TEXT_DIM, bg=C_CARD_BG)
        self._status_dot.pack(side="left")

        self._status_lbl = tk.Label(status_row, text="Checking…",
                                    font=(FONT_FAMILY, 9), fg=C_TEXT_DIM,
                                    bg=C_CARD_BG)
        self._status_lbl.pack(side="left", padx=(6, 0))

        self._version_lbl = tk.Label(self._status_card, text="",
                                     font=(FONT_FAMILY, 9), fg=C_TEXT_DIM,
                                     bg=C_CARD_BG)
        self._version_lbl.pack(anchor="w", pady=(4, 0))

        # Install path card
        path_card = Card(body)
        path_card.pack(fill="x", pady=(0, 12))

        SectionLabel(path_card, "Install Location").pack(anchor="w")
        FieldLabel(path_card, "PyRevit Extensions folder", pady=4).pack(anchor="w")

        path_frame = tk.Frame(path_card, bg=C_INPUT_BG, padx=8, pady=5,
                              highlightbackground=C_GREEN, highlightthickness=1)
        path_frame.pack(fill="x", pady=(4, 0))

        tk.Label(path_frame, text=INSTALL_DIR, font=(FONT_FAMILY, 8),
                 fg=C_INPUT_FG, bg=C_INPUT_BG, anchor="w").pack(fill="x")

        # Log card
        log_card = Card(body)
        log_card.pack(fill="x", pady=(0, 12))

        SectionLabel(log_card, "Activity Log").pack(anchor="w", pady=(0, 6))

        log_frame = tk.Frame(log_card, bg=C_HEADER_BG,
                             highlightbackground=C_GREEN, highlightthickness=1)
        log_frame.pack(fill="x")

        self._log = tk.Text(log_frame, height=5, bg=C_HEADER_BG, fg=C_TEXT,
                            font=(FONT_FAMILY, 8), bd=0, relief="flat",
                            state="disabled", padx=8, pady=6,
                            selectbackground=C_GREEN, wrap="word")
        self._log.pack(fill="x")

        # Buttons
        btn_row = tk.Frame(body, bg=C_WIN_BG)
        btn_row.pack(fill="x", pady=(12, 4))

        self._install_btn = HoverButton(
            btn_row,
            bg_normal=C_GREEN,
            bg_hover=C_GREEN_HOVER,
            bg_press=C_GREEN_PRESSED,
            text="Install",
            font=(FONT_FAMILY, 10, "bold"),
            fg=C_TEXT,
            width=14, height=1,
            command=self._on_install
        )
        self._install_btn.pack(side="left")

        self._close_btn = HoverButton(
            btn_row,
            bg_normal=C_SECONDARY_BTN,
            bg_hover=C_SECONDARY_HOVER,
            text="Close",
            font=(FONT_FAMILY, 9),
            fg=C_TEXT,
            width=10, height=1,
            command=self.destroy
        )
        self._close_btn.pack(side="right")

        # Progress bar (hidden until install starts)
        self._progress_frame = tk.Frame(body, bg=C_WIN_BG)
        self._progress_bar_bg = tk.Frame(self._progress_frame, bg=C_CARD_BG,
                                         height=4,
                                         highlightbackground=C_GREEN,
                                         highlightthickness=1)
        self._progress_bar_bg.pack(fill="x")
        self._progress_fill = tk.Frame(self._progress_bar_bg, bg=C_GREEN, height=4)
        self._progress_fill.place(relwidth=0, relheight=1)

    # ── Status check ──────────────────────────────────────────────────────────

    def _check_status(self):
        version = get_installed_version()
        if version:
            self._set_status("installed", f"Installed — v{version}")
            self._version_lbl.config(text=f"Location: {INSTALL_DIR}")
            self._install_btn.config(text="Reinstall")
            self._log_line(f"Existing installation found: v{version}")
            self._log_line("Click Reinstall to download the latest version.")
        else:
            self._set_status("not_installed", "Not installed")
            self._install_btn.config(text="Install")
            self._log_line("No existing installation detected.")
            self._log_line(f"Will install to: {INSTALL_DIR}")

    def _set_status(self, state, text):
        colours = {
            "installed":     C_GREEN,
            "not_installed": "#888888",
            "busy":          "#F0A500",
            "error":         C_DANGER,
        }
        self._status_dot.config(fg=colours.get(state, C_TEXT_DIM))
        self._status_lbl.config(text=text, fg=colours.get(state, C_TEXT))

    # ── Log ───────────────────────────────────────────────────────────────────

    def _log_line(self, msg):
        self._log.config(state="normal")
        self._log.insert("end", f"› {msg}\n")
        self._log.see("end")
        self._log.config(state="disabled")

    # ── Install ───────────────────────────────────────────────────────────────

    def _on_install(self):
        self._install_btn.config(state="disabled")
        self._set_status("busy", "Installing…")
        self._progress_frame.pack(fill="x", pady=(8, 0))
        self._animate_progress(0)

        threading.Thread(
            target=download_and_install,
            args=(
                lambda msg: self.after(0, self._log_line, msg),
                lambda ver: self.after(0, self._on_done, ver),
                lambda err: self.after(0, self._on_error, err),
            ),
            daemon=True
        ).start()

    def _animate_progress(self, val):
        """Indeterminate bounce animation while installing."""
        if not hasattr(self, "_installing") or not self._installing:
            return
        # ping-pong 0→1→0
        import math
        t = (math.sin(val * 0.15) + 1) / 2
        self._progress_fill.place(relwidth=max(0.1, t), relheight=1)
        self.after(30, self._animate_progress, val + 1)

    def _on_install(self):  # noqa: F811 — overrides above stub
        self._install_btn.config(state="disabled")
        self._set_status("busy", "Installing…")
        self._progress_frame.pack(fill="x", pady=(8, 0))
        self._installing = True
        self._animate_progress(0)

        threading.Thread(
            target=download_and_install,
            args=(
                lambda msg: self.after(0, self._log_line, msg),
                lambda ver: self.after(0, self._on_done, ver),
                lambda err: self.after(0, self._on_error, err),
            ),
            daemon=True
        ).start()

    def _on_done(self, version):
        self._installing = False
        self._progress_fill.place(relwidth=1, relheight=1)
        self._set_status("installed", f"Installed — v{version}")
        self._version_lbl.config(text=f"Location: {INSTALL_DIR}")
        self._log_line(f"✓ Installation complete — v{version}")
        self._log_line("Reload PyRevit in Revit to activate the Seed43 tab.")
        self._install_btn.config(state="normal", text="Reinstall")
        messagebox.showinfo(
            "Seed43 Setup",
            f"Installation complete — v{version}\n\nReload PyRevit inside Revit to see the Seed43 tab."
        )

    def _on_error(self, err):
        self._installing = False
        self._set_status("error", "Installation failed")
        self._log_line(f"✗ Error: {err}")
        self._install_btn.config(state="normal")
        messagebox.showerror("Seed43 Setup", f"Installation failed:\n\n{err}")


# ══════════════════════════════════════════════════════════════════════════════
# Entry point
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import traceback
    try:
        app = Seed43Installer()
        app.mainloop()
    except Exception as e:
        traceback.print_exc()
        input("Press Enter to close...")
