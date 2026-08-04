# -*- coding: utf-8 -*-
# pyfilter_pattern_preview.py
# Seed43 - Revit pattern preview renderer for WPF
#
# Parses Revit .pat files directly and renders accurate hatch previews.
# Falls back to Revit API FillGrid data if .pat file not found.
#
# Usage:
#   from pattern_preview import make_line_preview, make_fill_preview
#
#   canvas = make_line_preview(line_pattern_element, color_hex, width, height)
#   canvas = make_fill_preview(fg_pat_el, fg_color, bg_pat_el, bg_color, width, height)

import os
import math

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows.Controls import Canvas
from System.Windows.Shapes import Line, Rectangle
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Media import DoubleCollection

# ── COLOUR HELPERS ────────────────────────────────────────────────────────────

def _parse_hex(hex_str, fallback):
    if not hex_str or not hex_str.startswith("#") or len(hex_str) != 7:
        return fallback
    try:
        r = int(hex_str[1:3], 16)
        g = int(hex_str[3:5], 16)
        b = int(hex_str[5:7], 16)
        return Color.FromRgb(r, g, b)
    except Exception:
        return fallback

def _brush(color):
    return SolidColorBrush(color)

_WHITE   = Color.FromRgb(255, 255, 255)
_GREY    = Color.FromRgb(180, 180, 180)
_DIVIDER = Color.FromRgb(160, 160, 160)

# ── .PAT FILE PARSER ──────────────────────────────────────────────────────────

_PAT_CACHE = {}   # {revit_version: {pattern_name_lower: [grid_defs]}}

def _find_pat_files():
    """
    Return list of candidate .pat file paths, prioritising the currently
    running Revit version detected from the Revit API.
    """
    # Detect running Revit version
    running_year = None
    try:
        from pyrevit import revit
        ver = revit.doc.Application.VersionNumber
        running_year = str(ver).strip()
    except Exception:
        pass

    # Build search order -- running version first, then others as fallback
    years = []
    if running_year:
        years.append(running_year)
    for y in ("2026", "2025", "2024", "2023", "2022"):
        if y != running_year:
            years.append(y)

    candidates = []
    for year in years:
        for prog in (
            r"C:\Program Files\Autodesk\Revit {}".format(year),
            r"C:\Program Files (x86)\Autodesk\Revit {}".format(year),
        ):
            for fname in ("revit.pat", "fillpat.pat"):
                p = os.path.join(prog, "Data", fname)
                if os.path.isfile(p):
                    candidates.append(p)

    return candidates


def _parse_pat_file(path):
    """
    Parse a .pat file and return {name_lower: [grid_defs]}.
    Each grid_def is a dict:
      angle, x0, y0, shift, offset, dashes
    where dashes is a list of floats (positive=dash, negative=space).
    Units are kept as inches (as specified in the file).
    """
    patterns = {}
    current_name  = None
    current_grids = []
    units = "INCH"  # default

    try:
        with open(path, "r") as fh:
            lines = fh.readlines()
    except Exception:
        return patterns

    for raw in lines:
        line = raw.strip()
        if not line or line.startswith(";%VERSION"):
            continue
        if line.upper().startswith(";%UNITS="):
            units = line.split("=")[1].strip().upper()
            continue
        if line.startswith(";"):
            continue
        if line.startswith("*"):
            # Save previous
            if current_name and current_grids:
                patterns[current_name.lower()] = current_grids
            current_name  = line[1:].split(",")[0].strip()
            current_grids = []
            continue
        # Line family definition
        try:
            parts = [float(x.strip()) for x in line.split(",") if x.strip()]
            if len(parts) < 5:
                continue
            angle  = parts[0]
            x0     = parts[1]
            y0     = parts[2]
            shift  = parts[3]
            offset = parts[4]
            dashes = parts[5:] if len(parts) > 5 else []
            # Convert mm to inches if needed
            if units == "MM":
                x0     /= 25.4
                y0     /= 25.4
                shift  /= 25.4
                offset /= 25.4
                dashes  = [d / 25.4 for d in dashes]
            current_grids.append({
                "angle":  angle,
                "x0":     x0,
                "y0":     y0,
                "shift":  shift,
                "offset": offset,
                "dashes": dashes,
            })
        except Exception:
            continue

    if current_name and current_grids:
        patterns[current_name.lower()] = current_grids

    return patterns


def _load_pat_patterns():
    """Load and cache all available .pat patterns for the running Revit version."""
    global _PAT_CACHE
    # Use running version as cache key
    cache_key = "default"
    try:
        from pyrevit import revit
        cache_key = str(revit.doc.Application.VersionNumber).strip()
    except Exception:
        pass

    if cache_key in _PAT_CACHE:
        return _PAT_CACHE[cache_key]

    all_pats = {}
    for path in _find_pat_files():
        parsed = _parse_pat_file(path)
        all_pats.update(parsed)
        # Stop after first file found for the running version
        if parsed:
            break

    _PAT_CACHE[cache_key] = all_pats
    return all_pats


def _get_pat_grids(pattern_name):
    """Return grid defs for a named pattern, or None if not found."""
    if not pattern_name:
        return None
    pats = _load_pat_patterns()
    return pats.get(pattern_name.strip().lower())

# ── CANVAS HELPERS ────────────────────────────────────────────────────────────

def _make_canvas(width, height, bg=None):
    c = Canvas()
    c.Width        = width
    c.Height       = height
    c.ClipToBounds = True
    c.Background   = _brush(bg) if bg else _brush(Color.FromArgb(0, 0, 0, 0))
    return c


def _add_rect(canvas, x, y, w, h, color):
    r = Rectangle()
    r.Width  = w
    r.Height = h
    r.Fill   = _brush(color)
    Canvas.SetLeft(r, x)
    Canvas.SetTop(r,  y)
    canvas.Children.Add(r)

# ── HATCH RENDERER ────────────────────────────────────────────────────────────

def _draw_pat_grids(canvas, grids, color, width, height):
    """
    Render .pat grid definitions onto a canvas.

    The .pat format uses inches. We normalise so that the smallest
    positive offset renders at TARGET_PX pixels.
    """
    if not grids:
        return

    SCALE = 125.0  # pixels per inch, calibrated to Revit VG thumbnail density

    brush = _brush(color)
    cx = width  / 2.0
    cy = height / 2.0

    for gd in grids:
        angle_deg = gd["angle"]
        angle_rad = math.radians(angle_deg)
        cos_a = math.cos(angle_rad)
        sin_a = math.sin(angle_rad)

        # Perpendicular direction
        px = -sin_a
        py =  cos_a

        x0_px    = gd["x0"]    * SCALE
        y0_px    = gd["y0"]    * SCALE
        shift_px = gd["shift"] * SCALE
        offset_px = abs(gd["offset"]) * SCALE
        if offset_px < 1.5:
            offset_px = 1.5

        dashes_px = [d * SCALE for d in gd["dashes"]]

        diag    = math.sqrt(width * width + height * height)
        n_lines = int(diag / offset_px) + 4

        # WPF StrokeDashArray (relative to stroke thickness)
        thickness = 0.6
        dash_array = None
        if dashes_px:
            dc = DoubleCollection()
            for d in dashes_px:
                dc.Add(abs(d) / thickness)
            dash_array = dc

        for i in range(-n_lines, n_lines + 1):
            # Stagger: shift along line direction for each parallel line
            stagger = shift_px * i

            # Base point: canvas centre + pattern origin + perpendicular offset
            # Note: y0 is negated because WPF Y goes down, pat Y goes up
            base_u = cx + x0_px + px * (i * offset_px) + cos_a * stagger
            base_v = cy - y0_px + py * (i * offset_px) + sin_a * stagger

            ext = diag
            x1 = base_u - cos_a * ext
            y1 = base_v - sin_a * ext
            x2 = base_u + cos_a * ext
            y2 = base_v + sin_a * ext

            line = Line()
            line.X1 = x1
            line.Y1 = y1
            line.X2 = x2
            line.Y2 = y2
            line.Stroke          = brush
            line.StrokeThickness = thickness
            if dash_array:
                line.StrokeDashArray = dash_array
            Canvas.SetLeft(line, 0)
            Canvas.SetTop(line,  0)
            canvas.Children.Add(line)

# ── LINE PREVIEW ──────────────────────────────────────────────────────────────

def make_line_preview(line_pattern_element, color_hex,
                      width=100, height=16):
    """Render a horizontal line preview from a LinePatternElement."""
    color  = _parse_hex(color_hex, _GREY)
    canvas = _make_canvas(width, height)
    y      = height / 2.0

    line = Line()
    line.X1 = 4
    line.Y1 = y
    line.X2 = width - 4
    line.Y2 = y
    line.Stroke          = _brush(color)
    line.StrokeThickness = 1.5

    if line_pattern_element is not None:
        try:
            lp   = line_pattern_element.GetLinePattern()
            segs = list(lp.GetSegments())
            # Each segment has Type and Length (in feet)
            # Convert to pixels at a readable scale
            SCALE = 50.0  # pixels per foot
            dc = DoubleCollection()
            for seg in segs:
                length_px = max(abs(seg.Length) * SCALE, 0.5)
                dc.Add(length_px / line.StrokeThickness)
            if len(dc) > 0:
                line.StrokeDashArray = dc
        except Exception:
            pass

    Canvas.SetLeft(line, 0)
    Canvas.SetTop(line, 0)
    canvas.Children.Add(line)
    return canvas

# ── FILL PREVIEW ──────────────────────────────────────────────────────────────

def _is_solid(fill_pattern_element):
    try:
        return fill_pattern_element.GetFillPattern().IsSolidFill
    except Exception:
        return False


def _draw_half(canvas, pattern_element, color, x_offset, half_w, height):
    """
    Draw one half of the fill preview (FG or BG).
    Uses .pat file data if available, falls back to API FillGrid data.
    If the pattern element is missing (e.g. template opened in a different
    file where the pattern ID no longer resolves) but a colour is set, we
    treat it as a solid fill so the colour still previews.
    """
    is_solid = pattern_element is not None and _is_solid(pattern_element)

    # A colour with no resolvable pattern: assume solid fill (Revit's default
    # for filter colour overrides). Only treat as "no fill" when the colour is
    # the grey placeholder, meaning nothing was actually set.
    color_is_set = (color is not None and color != _GREY)
    if pattern_element is None and color_is_set:
        is_solid = True

    # Background rect
    bg_color = color if is_solid else _WHITE
    _add_rect(canvas, x_offset, 0, half_w, height, bg_color)

    if is_solid or pattern_element is None:
        return

    # Try .pat file first
    pat_name = None
    try:
        pat_name = pattern_element.Name
    except Exception:
        pass

    grids = _get_pat_grids(pat_name) if pat_name else None

    # Clip this half by drawing into a sub-canvas
    half_canvas = _make_canvas(half_w, height)

    if grids:
        _draw_pat_grids(half_canvas, grids, color, half_w, height)
    else:
        # Fallback: use API FillGrid data
        try:
            fp        = pattern_element.GetFillPattern()
            api_grids = list(fp.GetFillGrids())
            _draw_api_grids(half_canvas, api_grids, color, half_w, height)
        except Exception:
            pass

    Canvas.SetLeft(half_canvas, x_offset)
    Canvas.SetTop(half_canvas, 0)
    canvas.Children.Add(half_canvas)


def _draw_api_grids(canvas, fill_grids, color, width, height):
    """Fallback renderer using Revit API FillGrid objects."""
    if not fill_grids:
        return
    # API units are feet; 125 px/in * 12 in/ft = 1500 px/ft
    SCALE = 1500.0

    brush = _brush(color)
    cx = width  / 2.0
    cy = height / 2.0

    for grid in fill_grids:
        try:
            angle_rad = grid.Angle
            cos_a = math.cos(angle_rad)
            sin_a = math.sin(angle_rad)
            px    = -sin_a
            py    =  cos_a

            offset_px = grid.Offset * SCALE
            shift_px  = grid.Shift  * SCALE
            ox        = grid.Origin.U * SCALE
            oy        = grid.Origin.V * SCALE
            if offset_px < 1.5:
                offset_px = 1.5

            # Get dash/space segments from API
            thickness = 0.6
            dash_array = None
            try:
                segs = list(grid.GetSegments())
                if segs:
                    dc = DoubleCollection()
                    for s in segs:
                        dc.Add(max(abs(s) * SCALE, 0.3) / thickness)
                    dash_array = dc
            except Exception:
                pass

            diag    = math.sqrt(width * width + height * height)
            n_lines = int(diag / offset_px) + 4

            for i in range(-n_lines, n_lines + 1):
                stagger = shift_px * i
                base_u = cx + ox + px * (i * offset_px) + cos_a * stagger
                base_v = cy - oy + py * (i * offset_px) + sin_a * stagger

                ext  = diag
                line = Line()
                line.X1 = base_u - cos_a * ext
                line.Y1 = base_v - sin_a * ext
                line.X2 = base_u + cos_a * ext
                line.Y2 = base_v + sin_a * ext
                line.Stroke          = brush
                line.StrokeThickness = thickness
                if dash_array:
                    line.StrokeDashArray = dash_array
                Canvas.SetLeft(line, 0)
                Canvas.SetTop(line, 0)
                canvas.Children.Add(line)
        except Exception:
            continue


def make_fill_preview(fg_pattern_element, fg_color_hex,
                      bg_pattern_element, bg_color_hex,
                      width=100, height=16):
    """
    Render fill pattern preview: FG left half, BG right half.
    White background, pattern colour lines over it.
    Solid fill shows colour block.
    """
    fg_color = _parse_hex(fg_color_hex, _GREY)
    bg_color = _parse_hex(bg_color_hex, _GREY)

    canvas = _make_canvas(width, height)
    canvas.ClipToBounds = True

    half    = width / 2.0
    divider = 1.0
    hw      = half - divider / 2.0

    _draw_half(canvas, fg_pattern_element, fg_color, 0,                    hw, height)
    _add_rect(canvas, hw, 0, divider, height, _DIVIDER)
    _draw_half(canvas, bg_pattern_element, bg_color, half + divider / 2.0, hw, height)
    # Redraw divider on top
    _add_rect(canvas, hw, 0, divider, height, _DIVIDER)

    return canvas

# ── ELEMENT LOOKUPS ───────────────────────────────────────────────────────────

def _unwrap_pattern_val(pattern_val):
    """Accept either a plain id string or a {"id","name"} dict.
    Returns (id_str, name_or_None)."""
    if pattern_val is None:
        return None, None
    if isinstance(pattern_val, dict):
        return pattern_val.get("id"), pattern_val.get("name")
    return pattern_val, None


def find_fill_pattern_element(doc, pattern_val):
    id_str, name = _unwrap_pattern_val(pattern_val)
    if not id_str:
        return None
    try:
        from System import Int64
        from pyrevit import DB
        eid = DB.ElementId(Int64(int(id_str)))
        el  = doc.GetElement(eid)
        if isinstance(el, DB.FillPatternElement):
            return el
    except Exception:
        pass
    # ID didn't resolve in this doc — try by name
    if name:
        try:
            from pyrevit import DB
            for el in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
                try:
                    if el.Name == name:
                        return el
                except Exception:
                    pass
        except Exception:
            pass
    return None


def find_line_pattern_element(doc, pattern_val):
    id_str, name = _unwrap_pattern_val(pattern_val)
    if not id_str:
        return None
    try:
        from System import Int64
        from pyrevit import DB
        eid = DB.ElementId(Int64(int(id_str)))
        el  = doc.GetElement(eid)
        if isinstance(el, DB.LinePatternElement):
            return el
    except Exception:
        pass
    if name:
        try:
            from pyrevit import DB
            for el in DB.FilteredElementCollector(doc).OfClass(DB.LinePatternElement):
                try:
                    if el.Name == name:
                        return el
                except Exception:
                    pass
        except Exception:
            pass
    return None
