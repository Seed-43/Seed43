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

No files were relocated in the repo to make this work. The existing folders
(Settings/, Layout/Layouts/, Studio/studio_layouts/) still ship exactly as
before and now serve as the read-only defaults - nothing reads them at
runtime, so the shipped copy and the user's copy stopped being the same file
without anything moving. The move happens on the user's machine, when the tool
first opens.

Three collections are seeded from those folders on first run and never again -
they are user-editable templates, so one the user deletes must stay deleted
through every future update. Each has its OWN marker, so a template added to
one collection later cannot resurrect a deletion from another.

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

def _init_collections():
    """Bring each collection into .user, once ever.

    Because the shipped folder and the migration source are the SAME folder,
    the move doubles as the seeding: on a fresh install it copies the shipped
    templates in; on an existing install it carries the user's edited versions
    across. Either way the user ends up owning the files.

    The marker is what makes it safe to run every launch. Without it, an
    update restores a template the user deleted, migration finds it missing
    from .user and moves it straight back - deletions would never stick. With
    it, later runs ignore the shipped folder entirely.

    One marker per collection, so adding a sixth layout template later can
    seed it without also reviving a vocabulary deleted months ago.
    """
    _userdata.migrate_dir(
        _SHIPPED_SETTINGS, SETTINGS_DIR,
        once_marker=_userdata.user_path(TOOL, ".seeded_settings"))
    _userdata.migrate_dir(
        _SHIPPED_LAYOUTS, LAYOUTS_DIR,
        once_marker=_userdata.user_path(TOOL, ".seeded_layouts"))
    _userdata.migrate_dir(
        _SHIPPED_STUDIO_LAYOUTS, STUDIO_LAYOUTS_DIR,
        once_marker=_userdata.user_path(TOOL, ".seeded_studio_layouts"))

    # Single configs, not templates: migrate() already skips when the user
    # has their own, and nobody deletes these deliberately.
    _userdata.migrate(os.path.join(_HERE, "Layout", "layout_config.json"),
                      LAYOUT_CONFIG)
    _userdata.migrate(os.path.join(_HERE, "Studio", "studio_config.json"),
                      STUDIO_CONFIG)
    _userdata.migrate(os.path.join(_HERE, "pytransmit_sync.json"), SYNC_FILE)
    _userdata.migrate(os.path.join(_HERE, "pytransmit_setup.json"), SETUP_FILE)

    # once_marker, unlike the configs above: the logo is a shipped default the
    # user may legitimately replace or delete outright, so it must never be
    # reinstated or overwritten by a later update.
    _userdata.migrate(os.path.join(_HERE, "logo.png"), LOGO_FILE,
                      once_marker=_userdata.user_path(TOOL, ".seeded_logo"))


_init_collections()


# ── RESTORE (the way back from a deletion) ──────────────────────────────────

def restore_defaults(which="all"):
    """Re-copy shipped templates the user has deleted, on request.

    seed_once() only ever fires once, so without this a deleted template is
    gone for good. Additive and never overwrites, so it is safe to call twice
    or after editing a template.

    which: 'settings', 'layouts', 'studio', or 'all'. Returns files copied.
    """
    jobs = {
        "settings": (_SHIPPED_SETTINGS, SETTINGS_DIR, TEMPLATE_SETTINGS),
        "layouts": (_SHIPPED_LAYOUTS, LAYOUTS_DIR, None),
        "studio": (_SHIPPED_STUDIO_LAYOUTS, STUDIO_LAYOUTS_DIR, None),
    }
    if which != "all":
        jobs = {which: jobs[which]} if which in jobs else {}
    return sum(_userdata.seed_from_defaults(src, dst, names=names)
               for src, dst, names in jobs.values())
