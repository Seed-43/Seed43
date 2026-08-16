# -*- coding: utf-8 -*-
"""The one place that knows where pyTransmit's data lives.

Ten different modules used to recompute os.path.join(script_dir, 'Settings')
for themselves, with layouts and Studio layouts resolved separately again. Any
one of them missed during a move would silently split settings across two
locations, so they all import from here instead.

Layout on disk:

    .user/pyTransmit/
        Settings/          vocabularies + branding + pytransmit_setup
        Layouts/           Layout Builder templates
        studio_layouts/    Layout Studio templates
        layout_config.json
        studio_config.json
        pytransmit_sync.json

Nothing here lives beside the tool any more, so an update cannot touch it.

The app folder keeps its own copies for good:

    Settings/            the six vocabularies, shipped
    Layout/Layouts/      Layout Builder templates, shipped
    Studio/studio_layouts/  Studio templates, shipped
    Logos/               the stock logo, shipped

Those are the app's content - read-only, identical for every user, and never
written to or deleted at runtime. They are what a new install starts from and
what restore_defaults() copies back from.

Everything the tool actually reads and writes lives in .user. A collection is
seeded from the shipped folder ONLY while the user's own is empty; the moment
it holds anything, the shipped folder is never consulted again. That is what
stops an update putting a file back on top of someone's work.

Importing this module performs the migration and seeding. That is deliberate:
it happens once, before any caller reads a path, and no module has to remember
to call an init function.
"""

import os

from Snippets import _userdata

__all__ = [
    "USER_DIR", "SETTINGS_DIR", "LAYOUTS_DIR", "STUDIO_LAYOUTS_DIR",
    "LOGOS_DIR",
    "LAYOUT_CONFIG", "STUDIO_CONFIG", "SYNC_FILE", "SETUP_FILE", "LOGO_FILE",
    "settings_file", "restore_defaults",
]

TOOL = "pyTransmit"

_HERE = os.path.dirname(os.path.abspath(__file__))

# ── SHIPPED DEFAULTS (read-only at runtime - nothing loads from here) ───────

_SHIPPED_SETTINGS = os.path.join(_HERE, "Settings")
_SHIPPED_LAYOUTS = os.path.join(_HERE, "Layout", "Layouts")
_SHIPPED_STUDIO_LAYOUTS = os.path.join(_HERE, "Studio", "studio_layouts")
# The stock logo lives in a folder of its own rather than loose beside the
# script, so it sits alongside the other shipped content and the user's Logos
# folder has an obvious counterpart.
_SHIPPED_LOGOS = os.path.join(_HERE, "Logos")

# Only these are templates. branding.json and pytransmit_setup.json live in
# the same folder but are personal - one person's branding is not a sensible
# starting point for everyone else, so they are never seeded.
TEMPLATE_SETTINGS = [
    "recipients.json", "reason.json", "method.json",
    "format.json", "printsize.json", "distribution.json",
]

# ── USER DATA ───────────────────────────────────────────────────────────────

USER_DIR = _userdata.user_dir(TOOL)
SETTINGS_DIR = _userdata.user_dir(TOOL, "Settings")
LAYOUTS_DIR = _userdata.user_dir(TOOL, "Layouts")
STUDIO_LAYOUTS_DIR = _userdata.user_dir(TOOL, "studio_layouts")

# The user's own logo library, added when Studio grew a logo picker. Nothing
# is shipped into it - a logo is by definition the user's, so there is no
# seeding step and no marker: the folder simply starts empty and fills up as
# logos are loaded. A layout stores the path of the logo it uses, which is why
# they are copied in here rather than referenced where they were found: a logo
# left on someone's Desktop breaks the moment the layout is opened elsewhere.
LOGOS_DIR = _userdata.user_dir(TOOL, "Logos")

LAYOUT_CONFIG = _userdata.user_path(TOOL, "layout_config.json")
STUDIO_CONFIG = _userdata.user_path(TOOL, "studio_config.json")
SYNC_FILE = _userdata.user_path(TOOL, "pytransmit_sync.json")

# Two different files really are both called pytransmit_setup.json, holding
# different keys for different owners:
#   SETUP_FILE           SetupSettingsController - show_method, out_pdf, ...
#   Settings/ version    file naming and paths - output_path_template, ...
# They sat in different folders before and still do, so nothing collides.
SETUP_FILE = _userdata.user_path(TOOL, "pytransmit_setup.json")

# The shipped fallback logo, once the user owns it. Branding's own
# logo_source sync and any Settings/logo.* the user drops in both take
# precedence - this is only the last resort.
LOGO_FILE = _userdata.user_path(TOOL, "logo.png")


def settings_file(name):
    """Path to a file inside the user's Settings folder.

    settings_file('branding.json') rather than rebuilding the join, so no
    caller needs to know where Settings actually is.
    """
    return os.path.join(SETTINGS_DIR, name)


# ── MIGRATION (runs once, on first import after the update) ─────────────────

def _has_content(folder, suffix=".json"):
    """Does the user already own files in this collection?"""
    try:
        return any(n.lower().endswith(suffix) for n in os.listdir(folder))
    except Exception:
        return False


def _seed_collection(shipped, target, suffix=".json", names=None):
    """Copy the shipped starting point in, but only into an empty collection.

    The presence of anything in the user's folder is the whole test. Once it
    holds files they are the ones in use, and the shipped folder must not be
    consulted again for any reason - not to add a missing name, not to update
    an edited one. Anything else risks putting a file back on top of the
    user's work.

    Nothing is ever removed from the shipped folder. It is the app's own
    content: read-only, the same for everyone, and the thing "restore
    defaults" copies from. That is the difference from the old behaviour,
    which MOVED these folders into .user on first run and left the app with
    nothing - so restore_defaults() had no source left and quietly restored
    nothing at all.
    """
    if not os.path.isdir(shipped) or _has_content(target, suffix):
        return 0
    return _userdata.seed_from_defaults(shipped, target, suffix=suffix,
                                        names=names)


def _init_collections():
    """Give a new install something to start from; never touch an old one.

    Templates and the stock logo are COPIED out of the app folder into .user
    the first time, and only while .user has nothing of its own. Settings that
    are genuinely the user's - the configs below - are still migrated out of
    the old beside-the-script locations, because there the shipped file and
    the user's file really were the same file and leaving a copy behind lets
    the two drift.
    """
    _seed_collection(_SHIPPED_SETTINGS, SETTINGS_DIR, names=TEMPLATE_SETTINGS)
    _seed_collection(_SHIPPED_LAYOUTS, LAYOUTS_DIR)
    _seed_collection(_SHIPPED_STUDIO_LAYOUTS, STUDIO_LAYOUTS_DIR)
    _seed_collection(_SHIPPED_LOGOS, LOGOS_DIR, suffix=".png")

    # Single configs, not templates: migrate() already skips when the user
    # has their own, and nobody deletes these deliberately. These are the only
    # remaining moves, and their sources are legacy beside-the-script paths,
    # not shipped content - on a current install they no longer exist.
    _userdata.migrate(os.path.join(_HERE, "Layout", "layout_config.json"),
                      LAYOUT_CONFIG)
    _userdata.migrate(os.path.join(_HERE, "Studio", "studio_config.json"),
                      STUDIO_CONFIG)
    _userdata.migrate(os.path.join(_HERE, "pytransmit_sync.json"), SYNC_FILE)
    _userdata.migrate(os.path.join(_HERE, "pytransmit_setup.json"), SETUP_FILE)

    # The fallback logo, copied rather than moved for the same reason as the
    # templates - and only when the user has no logo of their own.
    if not os.path.isfile(LOGO_FILE):
        _shipped_logo = os.path.join(_SHIPPED_LOGOS, "logo.png")
        if os.path.isfile(_shipped_logo):
            try:
                import shutil
                shutil.copy2(_shipped_logo, LOGO_FILE)
            except Exception:
                pass


_init_collections()


# ── RESTORE (the way back from a deletion) ──────────────────────────────────

def restore_defaults(which="all"):
    """Re-copy shipped templates the user has deleted, on request.

    Seeding only happens into an empty collection, so without this a deleted
    template is gone for good. Additive and never overwrites: a template the
    user still has is left exactly as they edited it, and only missing names
    come back.

    This is the reason the shipped folders must keep their contents. They used
    to be MOVED into .user on first run, which left this function copying from
    an empty folder and restoring nothing.

    which: 'settings', 'layouts', 'studio', 'logos', or 'all'. Returns files
    copied.
    """
    jobs = {
        "settings": (_SHIPPED_SETTINGS, SETTINGS_DIR, TEMPLATE_SETTINGS, ".json"),
        "layouts": (_SHIPPED_LAYOUTS, LAYOUTS_DIR, None, ".json"),
        "studio": (_SHIPPED_STUDIO_LAYOUTS, STUDIO_LAYOUTS_DIR, None, ".json"),
        "logos": (_SHIPPED_LOGOS, LOGOS_DIR, None, ".png"),
    }
    if which != "all":
        jobs = {which: jobs[which]} if which in jobs else {}
    return sum(_userdata.seed_from_defaults(src, dst, suffix=suffix, names=names)
               for src, dst, names, suffix in jobs.values())
