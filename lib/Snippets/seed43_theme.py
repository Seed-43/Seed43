# -*- coding: utf-8 -*-
"""Seed43 shared colour palette loader/applier.

Reads UI/seed43_palette.json (found by walking up from the calling
script's own folder) and injects WPF SolidColorBrush resources into a
window's Resources dictionary AFTER wpf.LoadComponent / forms.WPFWindow
has already parsed the XAML.

IMPORTANT - call order:
    Every window in this extension is fully self-contained (its own XAML
    embeds its own hardcoded colour Setters, no external merged
    dictionary). Because of that, apply_seed43_palette() must be called
    AFTER the window's XAML has been loaded, not before - the XAML's own
    <Setter ... Value="{DynamicResource BrushXxx}"/> only shows whatever
    is in Resources at the moment WPF needs to paint, and calling this
    after load is what makes our injected brush win.

Usage in a tool's __init__:

    forms.WPFWindow.__init__(self, 'MyTool.xaml')   # or wpf.LoadComponent(...)
    apply_seed43_palette(self, _SCRIPT_DIR)

Colours only for now - sizing/dimensions are not applied by this module.
"""
import os
import json

import System
from System.Windows.Media import SolidColorBrush, Color


# Canonical semantic-key -> WPF brush-name mapping. Every tool's XAML
# should reference these same names via DynamicResource so one palette
# file drives all of them.
KEY_TO_BRUSH = {
    "accent":                 "BrushPrimaryGreen",
    "accent_hover":           "BrushPrimaryGreenHover",
    "accent_pressed":         "BrushPrimaryGreenPressed",
    "window_bg":               "BrushWindowBg",
    "card_bg":                 "BrushCardBg",
    "header_bg":                "BrushHeaderBg",
    "text_primary":            "BrushTextPrimary",
    "text_input":              "BrushTextInput",
    "input_bg":                "BrushInputBg",
    "input_bg_hover":          "BrushInputBgHover",
    "input_bg_pressed":        "BrushInputBgPressed",
    "border_default":          "BrushBorderDefault",
    "border_hover":            "BrushBorderHover",
    "border_pressed":          "BrushBorderPressed",
    "border_muted":            "BrushBorderMuted",
    "secondary_bg":             "BrushSecondaryBg",
    "secondary_hover":         "BrushSecondaryHover",
    "secondary_pressed":       "BrushSecondaryPressed",
    "secondary_disabled_bg":   "BrushSecondaryDisabledBg",
    "secondary_disabled_fg":   "BrushSecondaryDisabledFg",
    "dropdown_item_hover":     "BrushDropdownItemHover",
    "toggle_off_bg":            "BrushToggleOffBg",
    "toggle_off_hover":        "BrushToggleOffHover",
    "toggle_off_pressed":      "BrushToggleOffPressed",
    "toggle_knob":              "BrushToggleKnob",
    "danger":                   "BrushDanger",
    "danger_hover":             "BrushDangerHover",
    "danger_pressed":           "BrushDangerPressed",
}

_PALETTE_FILENAME = "seed43_palette.json"
_MAX_WALK_UP = 6


def _find_palette_path(start_dir):
    """Walk up from start_dir looking for a UI/seed43_palette.json."""
    folder = start_dir
    for _ in range(_MAX_WALK_UP):
        candidate = os.path.join(folder, "UI", _PALETTE_FILENAME)
        if os.path.isfile(candidate):
            return candidate
        parent = os.path.dirname(folder)
        if parent == folder:
            break
        folder = parent
    return None


def load_palette(start_dir):
    """Return the full parsed palette dict, or None if not found/unreadable."""
    path = _find_palette_path(start_dir)
    if not path:
        return None
    try:
        with open(path, "r") as f:
            return json.load(f)
    except Exception:
        return None


def get_active_profile(start_dir):
    """Return 'dark' or 'light' - defaults to 'dark' if anything's missing."""
    data = load_palette(start_dir)
    if not data:
        return "dark"
    return data.get("active_profile", "dark")


def _hex_to_color(hex_str):
    hex_str = hex_str.lstrip("#")
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    return Color.FromRgb(r, g, b)


def apply_seed43_palette(window, start_dir, profile=None):
    """Inject SolidColorBrush resources into window.Resources.

    Call this AFTER the window's XAML has already been loaded (see
    module docstring). Silently does nothing if the palette file can't
    be found or parsed, so a tool never hard-fails just because the
    palette is missing - it keeps whatever colours its own XAML shipped
    with as a fallback.
    """
    data = load_palette(start_dir)
    if not data:
        return False

    profile = profile or data.get("active_profile", "dark")
    colors = data.get("profiles", {}).get(profile)
    if not colors:
        return False

    for key, hex_value in colors.items():
        brush_name = KEY_TO_BRUSH.get(key)
        if not brush_name:
            continue
        try:
            window.Resources[brush_name] = SolidColorBrush(_hex_to_color(hex_value))
        except Exception:
            continue
    return True


def set_active_profile(start_dir, profile):
    """Write a new active_profile ('dark'/'light') back to the JSON on disk."""
    path = _find_palette_path(start_dir)
    if not path:
        return False
    try:
        with open(path, "r") as f:
            data = json.load(f)
        data["active_profile"] = profile
        with open(path, "w") as f:
            json.dump(data, f, indent=2)
        return True
    except Exception:
        return False


def _hex_shade(hex_str, percent):
    """percent > 0 lightens toward white, < 0 darkens toward black."""
    hex_str = hex_str.lstrip("#")
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    target = 0 if percent < 0 else 255
    p = abs(percent)
    r = int(round(r + (target - r) * p))
    g = int(round(g + (target - g) * p))
    b = int(round(b + (target - b) * p))
    return "#{0:02X}{1:02X}{2:02X}".format(r, g, b)


# Maps a dot-path in the JSON's "dimensions" block to a resource key +
# the WPF resource type it needs to be built as. Sizing needs several
# different .NET types (Thickness for padding, CornerRadius for corners,
# a plain double for font sizes/heights) unlike colours which are always
# a SolidColorBrush - so each entry says how to construct its value.
DIM_TO_RESOURCE = {
    # (json path)                          (resource key)                (kind)
    "button_primary.corner_radius":        ("CornerRadiusButton",        "corner"),
    "button_primary.font_size":            ("FontSizeButton",            "double"),
    "button_small.height":                 ("HeightButtonSmall",         "double"),
    "button_small.corner_radius":          ("CornerRadiusButtonSmall",   "corner"),
    "button_small.font_size":              ("FontSizeButtonSmall",       "double"),
    "button_secondary.corner_radius":      ("CornerRadiusButtonSecondary", "corner"),
    "button_secondary.font_size":          ("FontSizeButtonSecondary",   "double"),
    "button_close.width":                  ("WidthButtonClose",          "double"),
    "button_close.height":                 ("HeightButtonClose",         "double"),
    "button_close.corner_radius":          ("CornerRadiusButtonClose",   "corner"),
    "input_textbox.height":                ("HeightInput",               "double"),
    "input_textbox.corner_radius":         ("CornerRadiusInput",         "corner"),
    "input_textbox.border_thickness":      ("BorderThicknessInput",      "thickness_uniform"),
    "input_textbox.font_size":             ("FontSizeInput",             "double"),
    "input_combobox.height":               ("HeightCombobox",            "double"),
    "input_combobox.corner_radius":        ("CornerRadiusCombobox",      "corner"),
    "input_combobox.font_size":            ("FontSizeCombobox",          "double"),
    "dropdown_filter_combo.width":         ("WidthFilterCombo",          "double"),
    "dropdown_filter_combo.height":        ("HeightFilterCombo",         "double"),
    "dropdown_filter_combo.corner_radius": ("CornerRadiusFilterCombo",   "corner"),
    "dropdown_popup.corner_radius":        ("CornerRadiusDropdownPopup", "corner"),
    "dropdown_popup.min_width":            ("MinWidthDropdownPopup",     "double"),
    "dropdown_popup.item_corner_radius":   ("CornerRadiusDropdownItem",  "corner"),
    "dropdown_popup.arrow_size":           ("SizeDropdownArrow",         "double"),
    "top_bar.caption_height":              ("HeightTopBar",              "double"),
    "top_bar.corner_radius":               ("CornerRadiusWindow",        "corner"),
    "top_bar.control_btn_size":            ("SizeTopBarControlBtn",      "double"),
    "top_bar.control_btn_corner_radius":   ("CornerRadiusTopBarControlBtn", "corner"),
    "top_bar.logo_size":                   ("SizeTopBarLogo",            "double"),
    "toggle.width":                        ("WidthToggle",               "double"),
    "toggle.height":                       ("HeightToggle",              "double"),
    "toggle.corner_radius":                ("CornerRadiusToggle",        "corner"),
    "toggle.knob_size":                    ("SizeToggleKnob",            "double"),
    "toggle.knob_corner_radius":           ("CornerRadiusToggleKnob",    "corner"),
    "toggle.knob_margin":                  ("MarginToggleKnob",          "double"),
    "card.corner_radius":                  ("CornerRadiusCard",          "corner"),
    "card.padding":                        ("PaddingCard",               "thickness_uniform"),
    "icon.default_size":                   ("SizeIconDefault",           "double"),
    "text.section_label":                  ("FontSizeSectionLabel",      "double"),
    "text.field_label":                    ("FontSizeFieldLabel",        "double"),
    "text.heading":                        ("FontSizeHeading",           "double"),
    "text.body":                           ("FontSizeBody",              "double"),
}


def _get_dim(dimensions, dotted_path):
    a, b = dotted_path.split(".")
    group = dimensions.get(a)
    if not group:
        return None
    return group.get(b)


def apply_seed43_dimensions(window, start_dir):
    """Inject sizing resources (Thickness/CornerRadius/double) into
    window.Resources, same call-order rule as apply_seed43_palette:
    call this AFTER the window's XAML has loaded.
    """
    import System.Windows

    data = load_palette(start_dir)
    if not data:
        return False
    dimensions = data.get("dimensions")
    if not dimensions:
        return False

    for dotted_path, (res_key, kind) in DIM_TO_RESOURCE.items():
        value = _get_dim(dimensions, dotted_path)
        if value is None:
            continue
        try:
            if kind == "double":
                window.Resources[res_key] = float(value)
            elif kind == "corner":
                window.Resources[res_key] = System.Windows.CornerRadius(float(value))
            elif kind == "thickness_uniform":
                window.Resources[res_key] = System.Windows.Thickness(float(value))
        except Exception:
            continue

    # Padding pairs (padding_x/padding_y -> a single Thickness) need both
    # halves read together, so handle them separately from the 1:1 table.
    padding_pairs = {
        "button_primary":   "PaddingButtonPrimary",
        "button_small":     "PaddingButtonSmall",
        "button_secondary": "PaddingButtonSecondary",
        "input_textbox":    "PaddingInput",
        "input_combobox":   "PaddingCombobox",
        "dropdown_popup":   "PaddingDropdownItem",  # uses item_padding_x/y below
    }
    for group_name, res_key in padding_pairs.items():
        group = dimensions.get(group_name)
        if not group:
            continue
        if group_name == "dropdown_popup":
            px, py = group.get("item_padding_x"), group.get("item_padding_y")
        else:
            px, py = group.get("padding_x"), group.get("padding_y")
        if px is None or py is None:
            continue
        try:
            window.Resources[res_key] = System.Windows.Thickness(float(px), float(py), float(px), float(py))
        except Exception:
            continue

    return True


def set_accent(start_dir, hex_color, profiles=("dark", "light")):
    """Set a new accent colour, deriving hover/pressed/border/border_muted
    from it with the same 8%/27% lighten gradient used everywhere else
    in the palette. Writes both profiles by default.
    """
    path = _find_palette_path(start_dir)
    if not path:
        return False
    try:
        with open(path, "r") as f:
            data = json.load(f)
        hover = _hex_shade(hex_color, 0.08)
        pressed = _hex_shade(hex_color, 0.27)
        for profile in profiles:
            p = data.get("profiles", {}).get(profile)
            if not p:
                continue
            p["accent"] = hex_color
            p["accent_hover"] = hover
            p["accent_pressed"] = pressed
            p["border_default"] = hex_color
            p["border_hover"] = hover
            p["border_pressed"] = pressed
            p["border_muted"] = hex_color
        with open(path, "w") as f:
            json.dump(data, f, indent=2)
        return True
    except Exception:
        return False
