# -*- coding: utf-8 -*-
"""Shared helpers behind every Seed43 tool's hamburger menu: the installed
version lookup, the pre-filled support email, and the pre-filled GitHub
issue link."""

import os
import urllib

# ── CONSTANTS ─────────────────────────────────────────────────────────────────
SUPPORT_EMAIL = "support@seed43.org"
DONATE_URL    = "https://buymeacoffee.com/seed43"
ISSUE_NEW_URL = "https://github.com/Seed-43/Seed43/issues/new"


# ── VERSION ───────────────────────────────────────────────────────────────────

def find_seed43_version(start_dir):
    """Walk up from start_dir to Seed43.extension/version.txt and return just
    the version string (its first line). Returns 'unknown' if the file can't
    be found or read."""
    folder = start_dir
    for _ in range(6):
        candidate = os.path.join(folder, 'version.txt')
        if os.path.isfile(candidate):
            try:
                with open(candidate, 'r') as f:
                    return f.readline().strip()
            except Exception:
                return 'unknown'
        parent = os.path.dirname(folder)
        if parent == folder:
            break
        folder = parent
    return 'unknown'


def revit_version():
    """Return the running Revit version, or 'unknown' outside Revit."""
    try:
        from pyrevit import HOST_APP
        return str(HOST_APP.version)
    except Exception:
        return 'unknown'


def _quote(text):
    """Percent-encode text for use as a URL query value.

    IronPython 2's urllib.quote works on bytes, and raises KeyError on any
    non-ASCII character in a unicode string, so encode to UTF-8 first."""
    try:
        return urllib.quote(text.encode('utf-8'))
    except Exception:
        return urllib.quote(str(text))


# ── OPENING LINKS ─────────────────────────────────────────────────────────────

_last_open_times = {}


def open_url(url, window=None, on_error=None):
    """Open a URL (or a mailto: link) in the default handler, without blocking
    the UI thread.

    subprocess.Popen('cmd /c start ...') spawns cmd.exe as a shell wrapper, and
    that first launch can hang for a long time from inside Revit's process
    (shell resolution, security scanning). Run synchronously on the UI thread
    that freezes the whole window for the entire wait, and every click made
    during the freeze then fires at once the moment it unblocks. os.startfile
    skips the shell wrapper entirely, and running it on a background thread
    means even a slow launch can never block the UI.

    Repeat calls for the same url within 2 seconds are ignored, so a
    double-clicked menu item can't open two browser windows.

    on_error, if given, is called with the message when both launch attempts
    fail. When window is also given it is marshalled onto that window's
    dispatcher first, so on_error is free to show a dialog.
    """
    import time
    now = time.time()
    if now - _last_open_times.get(url, 0.0) < 2.0:
        return
    _last_open_times[url] = now

    def _report(msg):
        if not on_error:
            return
        if window is not None:
            try:
                import System
                window.Dispatcher.Invoke(System.Action(lambda: on_error(msg)))
                return
            except Exception:
                pass
        try:
            on_error(msg)
        except Exception:
            pass

    def _launch():
        try:
            os.startfile(url)
        except Exception:
            try:
                import subprocess
                subprocess.Popen(['cmd', '/c', 'start', '', url])
            except Exception as ex:
                _report(u"Could not open browser:\n{0}".format(ex))

    import threading
    threading.Thread(target=_launch).start()


# ── SUPPORT LINKS ─────────────────────────────────────────────────────────────

def support_mailto(app_name, start_dir):
    """Return a mailto: URI for a support email, pre-filled with the app name
    and the installed Seed43 version."""
    body = (
        u"Hi Seed43 Team,\n\n"
        u"Support Request\n\n"
        u"App: {0}\n"
        u"Seed43 Version: {1}\n"
        u"Revit Version: {2}\n\n"
        u"Please describe your issue below:\n\n"
    ).format(app_name, find_seed43_version(start_dir), revit_version())
    return u"mailto:{0}?subject={1}&body={2}".format(
        SUPPORT_EMAIL,
        _quote(u"{0} Support Ticket".format(app_name)),
        _quote(body))


def github_issue_url(app_name, start_dir):
    """Return a GitHub 'new issue' URL for the Seed43 repo, pre-filled with
    the app name, the installed Seed43 version and the running Revit version.

    GitHub reads title and body straight off the query string, so the form
    opens already filled in. The user still has to be signed in, and still has
    to press Submit - this only saves them the typing."""
    body = (
        u"**App:** {0}\n"
        u"**Seed43 version:** {1}\n"
        u"**Revit version:** {2}\n\n"
        u"### What happened\n\n\n"
        u"### Steps to reproduce\n\n"
        u"1. \n"
        u"2. \n"
        u"3. \n\n"
        u"### What you expected instead\n\n"
    ).format(app_name, find_seed43_version(start_dir), revit_version())
    return u"{0}?title={1}&body={2}".format(
        ISSUE_NEW_URL,
        _quote(u"[{0}] ".format(app_name)),
        _quote(body))
