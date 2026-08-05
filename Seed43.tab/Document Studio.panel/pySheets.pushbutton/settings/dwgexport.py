# -*- coding: utf-8 -*-
# dwgexport.py
"""DWG Export Settings — read/write the DWG sub-tab controls."""
from collections import namedtuple

DWGSettings = namedtuple('DWGSettings', [
    'export_setup_name',   # name shown in dwg_setup_cb ('<Revit Default>' etc.)
    'export_xrefs',        # True = views/links as external references
])

DEFAULT_DWG_SETTINGS = DWGSettings(
    export_setup_name = '<Revit Default>',
    export_xrefs      = False,
)


def read_from_window(win):
    try:
        return DWGSettings(
            export_setup_name = win.dwg_setup_cb.SelectedItem or '<Revit Default>',
            export_xrefs      = bool(win.dwg_xrefs_cb.IsChecked),
        )
    except Exception:
        return DEFAULT_DWG_SETTINGS


def apply_to_window(win, settings=None):
    s = settings or DEFAULT_DWG_SETTINGS
    try:
        items = list(win.dwg_setup_cb.ItemsSource or [])
        if s.export_setup_name in items:
            win.dwg_setup_cb.SelectedItem = s.export_setup_name
        win.dwg_xrefs_cb.IsChecked = bool(s.export_xrefs)
    except Exception:
        pass
