# -*- coding: utf-8 -*-
# dgnexport.py
"""DGN Export Settings — read/write the DGN sub-tab controls."""
from collections import namedtuple

DGNSettings = namedtuple('DGNSettings', [
    'export_setup_name',
])

DEFAULT_DGN_SETTINGS = DGNSettings(
    export_setup_name = '<Revit Default>',
)


def read_from_window(win):
    try:
        return DGNSettings(
            export_setup_name = win.dgn_setup_cb.SelectedItem or '<Revit Default>',
        )
    except Exception:
        return DEFAULT_DGN_SETTINGS


def apply_to_window(win, settings=None):
    s = settings or DEFAULT_DGN_SETTINGS
    try:
        items = list(win.dgn_setup_cb.ItemsSource or [])
        if s.export_setup_name in items:
            win.dgn_setup_cb.SelectedItem = s.export_setup_name
    except Exception:
        pass
