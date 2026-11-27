# -*- coding: utf-8 -*-
# Opens View Organiser's setup window again, filled in with the saved setup, so
# the sheet-to-view / schedule / legend links and the casing can be changed.
from pyrevit import revit
from Snippets import _organiser_setup as setup

saved = setup.migrate_config(setup.load_config())

notice = ("Change the links or casing and save. Your current setup is filled in.")
missing = setup.missing_parts(saved)
if missing:
    notice = ("The saved setup is missing: {}. Fill them in and save.".format(
        ", ".join(missing)))

setup.show_setup(revit.doc, saved, notice)
