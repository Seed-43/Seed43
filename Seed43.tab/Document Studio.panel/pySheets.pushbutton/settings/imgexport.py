# -*- coding: utf-8 -*-
# imgexport.py
"""Image Export Settings — read/write the IMG sub-tab controls."""
from collections import namedtuple

IMGSettings = namedtuple('IMGSettings', [
    'image_type',   # 'PNG' | 'JPEG' | 'TIFF'
    'resolution',   # '72 DPI' | '150 DPI' | '300 DPI' | '600 DPI'
    'pixel_size',   # longest-side pixels (int)
])

DEFAULT_IMG_SETTINGS = IMGSettings(
    image_type = 'PNG',
    resolution = '150 DPI',
    pixel_size = 2048,
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
        try:
            px = int(win.img_pixels_tb.Text or 2048)
        except Exception:
            px = 2048
        return IMGSettings(
            image_type = _combo_text(win.img_type_cb, 'PNG'),
            resolution = _combo_text(win.img_res_cb, '150 DPI'),
            pixel_size = px,
        )
    except Exception:
        return DEFAULT_IMG_SETTINGS


def apply_to_window(win, settings=None):
    s = settings or DEFAULT_IMG_SETTINGS
    try:
        for item in win.img_type_cb.Items:
            if getattr(item, 'Content', item) == s.image_type:
                win.img_type_cb.SelectedItem = item
                break
        for item in win.img_res_cb.Items:
            if getattr(item, 'Content', item) == s.resolution:
                win.img_res_cb.SelectedItem = item
                break
        win.img_pixels_tb.Text = str(s.pixel_size)
    except Exception:
        pass
