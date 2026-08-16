# -*- coding: utf-8 -*-
# nwcexport.py
"""NWC (Navisworks) Export Settings — read/write the NWC sub-tab controls."""
from collections import namedtuple

NWCSettings = namedtuple('NWCSettings', [
    'shared_coords',       # True = shared coordinates, False = internal
    'convert_properties',  # True = convert element properties
])

DEFAULT_NWC_SETTINGS = NWCSettings(
    shared_coords      = True,
    convert_properties = True,
)


def read_from_window(win):
    try:
        return NWCSettings(
            shared_coords      = bool(win.nwc_shared_coords_cb.IsChecked),
            convert_properties = bool(win.nwc_props_cb.IsChecked),
        )
    except Exception:
        return DEFAULT_NWC_SETTINGS


def apply_to_window(win, settings=None):
    s = settings or DEFAULT_NWC_SETTINGS
    try:
        win.nwc_shared_coords_cb.IsChecked = bool(s.shared_coords)
        win.nwc_props_cb.IsChecked         = bool(s.convert_properties)
    except Exception:
        pass
