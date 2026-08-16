# -*- coding: utf-8 -*-
# ifcexport.py
"""IFC Export Settings — read/write the IFC sub-tab controls."""
from collections import namedtuple

IFCSettings = namedtuple('IFCSettings', [
    'version',     # label shown in ifc_version_cb
])

DEFAULT_IFC_SETTINGS = IFCSettings(
    version = 'IFC 2x3 Coordination View 2.0',
)


def _combo_text(cb, fallback):
    try:
        item = cb.SelectedItem
        if item is None:
            return fallback
        return getattr(item, 'Content', item) or fallback
    except Exception:
        return fallback


def read_from_window(win):
    try:
        return IFCSettings(
            version = _combo_text(win.ifc_version_cb,
                                  'IFC 2x3 Coordination View 2.0'),
        )
    except Exception:
        return DEFAULT_IFC_SETTINGS


def apply_to_window(win, settings=None):
    s = settings or DEFAULT_IFC_SETTINGS
    try:
        for item in win.ifc_version_cb.Items:
            if getattr(item, 'Content', item) == s.version:
                win.ifc_version_cb.SelectedItem = item
                break
    except Exception:
        pass
