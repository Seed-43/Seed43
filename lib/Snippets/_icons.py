# -*- coding: utf-8 -*-
__title__     = "Icons"
__author__    = "Nagel Consultants"
__doc__       = """
Vector icons for Seed43 panels and scripts. Each icon is a scalable shape
rather than a bitmap, so it stays crisp at any size.

    from Snippets._icons import make_icon

    btn.Content = make_icon("pdf_output", size=16, color="#FFFFFF")

The name must match a key in _icons.json, which sits in this folder. Icons
are drawn on a 24x24 grid and scaled to fit, and the result drops straight
into any WPF element that accepts a child.
"""

import os
import json
import clr

import System.Windows as _SW
import System.Windows.Controls as _SWC
import System.Windows.Media as _SWM
import System.Windows.Shapes as _SWS

# ── LOAD ICON DATA ────────────────────────────────────────────────────
_HERE  = os.path.dirname(os.path.abspath(__file__))
_JSON  = os.path.join(_HERE, "_icons.json")

with open(_JSON, "r") as _f:
    _ICONS = json.load(_f)

# ── HELPERS ───────────────────────────────────────────────────────────
def _parse_color(hex_color):
    """Convert a hex color string to a WPF SolidColorBrush."""
    h = hex_color.lstrip("#")
    if len(h) == 6:
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return _SWM.SolidColorBrush(_SWM.Color.FromRgb(r, g, b))
    if len(h) == 8:
        a, r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16), int(h[6:8], 16)
        return _SWM.SolidColorBrush(_SWM.Color.FromArgb(a, r, g, b))
    return _SWM.Brushes.White

# ── PUBLIC API ────────────────────────────────────────────────────────
def make_icon(key, size=16, color="#FFFFFF"):
    """
    Return a WPF Viewbox containing the named icon, ready to use as
    the Content of a Button, or a child of any layout panel.

    key   : string matching a key in _icons.json
    size  : pixel width and height of the rendered icon
    color : fill color as a hex string, e.g. "#FFFFFF" or "#FF208A3C"
    """
    path_data = _ICONS.get(key)
    if not path_data:
        # Return an empty placeholder so missing keys don't crash the UI
        vb = _SWC.Viewbox()
        vb.Width = size
        vb.Height = size
        return vb

    shape = _SWS.Path()
    shape.Data    = _SWM.Geometry.Parse(path_data)
    shape.Fill    = _parse_color(color)
    shape.Stretch = _SWM.Stretch.Uniform

    vb = _SWC.Viewbox()
    vb.Width  = size
    vb.Height = size
    vb.Child  = shape
    return vb


def make_icon_with_label(key, label, icon_size=14, color="#FFFFFF", spacing=4):
    """
    Return a horizontal StackPanel with an icon and a text label beside it.
    Useful for buttons that need both a symbol and a word.

    key       : icon key
    label     : text to show beside the icon
    icon_size : pixel size of the icon
    color     : color applied to both the icon and the label text
    spacing   : gap in pixels between icon and label
    """
    brush = _parse_color(color)

    sp = _SWC.StackPanel()
    sp.Orientation = _SWC.Orientation.Horizontal

    icon = make_icon(key, size=icon_size, color=color)
    icon.VerticalAlignment = _SW.VerticalAlignment.Center
    icon.Margin = _SW.Thickness(0, 0, spacing, 0)
    sp.Children.Add(icon)

    tb = _SWC.TextBlock()
    tb.Text = label
    tb.Foreground = brush
    tb.VerticalAlignment = _SW.VerticalAlignment.Center
    sp.Children.Add(tb)

    return sp


def icon_keys():
    """Return a sorted list of all available icon key names."""
    return sorted(_ICONS.keys())
