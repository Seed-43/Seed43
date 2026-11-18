# -*- coding: utf-8 -*-
"""Beam Rebar: reinforce concrete beams with native Revit rebar from a
dialogue. Pick a rebar layout (closed stirrup, stirrup with centre links,
overlapping stirrups, two U bars lapped, flat link round the bottom bars),
set the bars, covers and stirrup spacing, and Apply to new beams clicked
in the model or to beams already there. Update reads a beam's own size
back into the dialogue.
"""
import io
import os.path as op
import json
import math

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import FileStream, FileMode, FileAccess
from System.Windows import (Point, FontWeights, Visibility, Size, Thickness,
                            HorizontalAlignment, VerticalAlignment, Rect)
from System.Windows import LogicalTreeHelper
from System.Windows.Controls import (Canvas, TextBlock, TextBox, ComboBox, ComboBoxItem,
                                     StackPanel, Orientation)
from System.Windows.Controls.Primitives import TextBoxBase, ButtonBase, ScrollBar
from System.Windows.Input import Keyboard, FocusManager
from System.Windows.Markup import XamlReader
from System.Windows.Media import (Visual, VisualTreeHelper, PointCollection, RotateTransform,
                                  SolidColorBrush, Color, DoubleCollection, PenLineJoin,
                                  PenLineCap, DrawingBrush, DrawingGroup, GeometryDrawing,
                                  EllipseGeometry, TileMode, BrushMappingMode, Stretch)
from System.Windows.Shapes import Line, Polygon, Polyline, Ellipse, Rectangle

from pyrevit import revit, DB, forms
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.DB.Structure import RebarBarType, RebarHookType, RebarHostData

from Snippets import _dialogs as dlg
from Snippets import _userdata
from Snippets._icons import set_header_icon
from Snippets.seed43_theme import apply_seed43_palette, apply_seed43_dimensions

import beam_engine as engine

doc = revit.doc
uidoc = revit.uidoc

# ── CONSTANTS ───────────────────────────────────────────────────────────────

SCRIPT_DIR = op.dirname(__file__)
TITLE = u"Beam Rebar"
MM = 304.8
SETTINGS = _userdata.user_path(u"Beam Rebar", u"settings.json")
STYLES_XAML = op.join(op.dirname(op.abspath(dlg.__file__)), u"Seed43Styles.xaml")
NO_HOOK = u"(no hook)"
STEEL = 7850.0                                  # kg/m3

# the window's own settings; each layout's defaults (engine.LAYOUTS) go over these
DEFAULTS = {"layout": engine.LAYOUTS[0][0], "name": u"FB1", "width": u"600", "depth": u"600",
            "length": u"6000", "beams": u"1", "top_offset": u"0", "slab": u"100", "grade": u"500E",
            "beam_type": u"", "level": u"",
            "cover_side": u"75", "cover_top": u"50", "cover_bottom": u"75", "end_cover": u"50",
            "main_bar": u"HD16", "side_bar": u"HD12", "tie_bar": u"HR10", "link_bar": u"HR10",
            "tie_hook": u"Stirrup/Tie Seismic - 135 deg.", "link_hook": u"Standard - 180 deg.",
            "top_count": u"3", "bottom_count": u"3", "side_count": u"1", "tie_count": u"1",
            "link_count": u"0", "u_lap": u"200",
            "tie_first": u"50", "tie_sp": u"200", "zone_len": u"0", "zone_sp": u"100"}
DEFAULTS.update(engine.LAYOUTS[0][3])

# text settings and the box each one sits in (key + "_tb")
TEXTS = ("name", "width", "depth", "length", "beams", "top_offset", "slab", "cover_side",
         "cover_top", "cover_bottom", "end_cover", "top_count", "bottom_count", "side_count",
         "tie_count", "u_lap", "tie_first", "tie_sp", "zone_len", "zone_sp", "link_count")
# the bar and hook lists: (setting, combo)
BAR_LISTS = (("main_bar", "main_bar_cb"), ("side_bar", "side_bar_cb"), ("tie_bar", "tie_bar_cb"),
             ("link_bar", "link_bar_cb"))
HOOK_LISTS = (("tie_hook", "tie_hook_cb"), ("link_hook", "link_hook_cb"))
# kept when the layout changes: the beam itself, not its bars
KEEP_ON_LAYOUT = ("name", "width", "depth", "length", "beams", "top_offset", "slab", "grade",
                  "level", "beam_type", "cover_side", "cover_top", "cover_bottom", "end_cover",
                  "main_bar", "side_bar", "tie_bar", "tie_first", "tie_sp", "zone_len", "zone_sp")

BAR_PICK = u"Beam Rebar: select the beams to reinforce, then Finish. ESC goes back."
BAR_START = u"Beam Rebar: click the START of beam {0} ({1} placed so far). ESC when done."
BAR_END = u"Beam Rebar: click the END of beam {0}. ESC drops it."
BAR_UPDATE = u"Beam Rebar: click the beam to update. Its size and level fill the dialogue. ESC goes back."

# concrete in the drawings: the same greys on a dark or a light theme
CONCRETE_BASE = SolidColorBrush(Color.FromRgb(0xC4, 0xC6, 0xC8))
CONCRETE_EDGE = SolidColorBrush(Color.FromRgb(0x3A, 0x40, 0x48))


def _speckled_concrete():
    """The grain of concrete as on the section drawings (as Column Rebar's)."""
    import random
    rnd = random.Random(43)
    group = DrawingGroup()
    for shade, count in ((0x9C, 22), (0xAE, 26), (0xDA, 18)):
        brush = SolidColorBrush(Color.FromRgb(shade, shade + 2, shade + 4))
        for _ in range(count):
            size = rnd.uniform(0.5, 1.6)
            group.Children.Add(GeometryDrawing(brush, None, EllipseGeometry(
                Point(rnd.uniform(2, 58), rnd.uniform(2, 58)), size, size)))
    tile = DrawingBrush(group)
    tile.TileMode = TileMode.Tile
    tile.Stretch = Stretch.None
    tile.Viewbox = Rect(0, 0, 60, 60)
    tile.ViewboxUnits = BrushMappingMode.Absolute
    tile.Viewport = Rect(0, 0, 60, 60)
    tile.ViewportUnits = BrushMappingMode.Absolute
    tile.Freeze()
    return tile


CONCRETE_SPECKLED = _speckled_concrete()


# ── HELPERS (private to this file) ──────────────────────────────────────────

def _name(element):
    return DB.Element.Name.GetValue(element)


def _load_settings():
    data = {}
    try:
        with io.open(SETTINGS, "r", encoding="utf-8") as f:
            data = json.loads(f.read())
    except Exception:
        pass
    return data


def _save_settings(data):
    try:
        with io.open(SETTINGS, "w", encoding="utf-8") as f:
            f.write(unicode(json.dumps(data, indent=2)))
    except Exception:
        pass


def _full(s):
    """A settings dict with everything filled: the window's defaults, the
    layout's own defaults over them, then `s`."""
    data = dict(DEFAULTS)
    data.update(engine.layout_info(s.get("layout", DEFAULTS["layout"]))[2])
    data.update(s)
    return data


def _line(canvas, ink, x1, y1, x2, y2, thick=1.0, opacity=1.0, dash=None):
    ln = Line()
    ln.X1, ln.Y1, ln.X2, ln.Y2 = x1, y1, x2, y2
    ln.Stroke = ink
    ln.StrokeThickness = thick
    ln.Opacity = opacity
    if dash:
        ln.StrokeDashArray = DoubleCollection(dash)
    canvas.Children.Add(ln)
    return ln


def _arrow(canvas, ink, px, py, ux, uy, size=11.0, half=3.6):
    """A filled arrowhead with its point on (px, py), pointing along (ux, uy)."""
    bx, by = px - ux * size, py - uy * size
    head = Polygon()
    head.Points = PointCollection([Point(px, py), Point(bx - uy * half, by + ux * half),
                                   Point(bx + uy * half, by - ux * half)])
    head.Fill = ink
    canvas.Children.Add(head)


def _dim_line(canvas, ink, x1, y1, x2, y2):
    """A dimension line, CAD style: thin line, a filled arrowhead at each end."""
    _line(canvas, ink, x1, y1, x2, y2, 1.0, 0.95)
    length = math.hypot(x2 - x1, y2 - y1)
    if length < 2.0:
        return
    ux, uy = (x2 - x1) / length, (y2 - y1) / length
    size = min(12.0, length * 0.45)
    _arrow(canvas, ink, x1, y1, -ux, -uy, size, size * 0.3)
    _arrow(canvas, ink, x2, y2, ux, uy, size, size * 0.3)


def _dim_outside(canvas, ink, x1, y1, x2, y2):
    """A small dimension: arrows coming in from outside each end."""
    _line(canvas, ink, x1, y1, x2, y2, 1.0, 0.95)
    run = math.hypot(x2 - x1, y2 - y1) or 1.0
    for px, py, ux, uy in ((x1, y1, (x2 - x1) / run, (y2 - y1) / run),
                           (x2, y2, (x1 - x2) / run, (y1 - y2) / run)):
        _line(canvas, ink, px - ux * 14.0, py - uy * 14.0, px, py)
        _arrow(canvas, ink, px, py, ux, uy, 9.0, 3.0)


def _ext_line(canvas, ink, x1, y1, x2, y2):
    _line(canvas, ink, x1, y1, x2, y2, 1.0, 0.5)


def _concrete_shape(canvas, shape):
    shape.Fill = CONCRETE_BASE
    shape.Stroke = CONCRETE_EDGE
    shape.StrokeThickness = 1.5
    canvas.Children.Add(shape)


def _grain(canvas, shape):
    shape.Fill = CONCRETE_SPECKLED
    shape.IsHitTestVisible = False
    canvas.Children.Add(shape)


def _polygon(pts):
    pg = Polygon()
    pg.Points = PointCollection([Point(x, y) for x, y in pts])
    return pg


def _ellipse(cx, cy, rx, ry=None):
    e = Ellipse()
    e.Width, e.Height = 2 * rx, 2 * (ry if ry is not None else rx)
    Canvas.SetLeft(e, cx - rx)
    Canvas.SetTop(e, cy - (ry if ry is not None else rx))
    return e


def _rect(x0, y0, x1, y1, radius=0.0):
    r = Rectangle()
    r.Width, r.Height = abs(x1 - x0), abs(y1 - y0)
    r.RadiusX = r.RadiusY = radius
    Canvas.SetLeft(r, min(x0, x1))
    Canvas.SetTop(r, min(y0, y1))
    return r


def _label(canvas, text, x, y, size=11, bold=False, opacity=0.9, ink=None):
    tb = TextBlock()
    tb.Text = text
    tb.FontSize = size
    tb.Opacity = opacity
    if bold:
        tb.FontWeight = FontWeights.SemiBold
    if ink is not None:
        tb.Foreground = ink
    Canvas.SetLeft(tb, x)
    Canvas.SetTop(tb, y)
    canvas.Children.Add(tb)
    return tb


def _rounded(pts, closed, r):
    """A polyline with its corners rounded by about r px, as a bar is bent."""
    if len(pts) < 3 or r <= 0.5:
        return list(pts) + ([pts[0]] if closed else [])
    ring = list(pts)
    n = len(ring)
    out = []
    corners = range(n) if closed else range(1, n - 1)
    if not closed:
        out.append(ring[0])
    for i in corners:
        a, b, c = ring[i - 1], ring[i], ring[(i + 1) % n]
        la = math.hypot(a[0] - b[0], a[1] - b[1]) or 1.0
        lc = math.hypot(c[0] - b[0], c[1] - b[1]) or 1.0
        k = min(r, la / 2.0, lc / 2.0)
        p = (b[0] + (a[0] - b[0]) / la * k, b[1] + (a[1] - b[1]) / la * k)
        q = (b[0] + (c[0] - b[0]) / lc * k, b[1] + (c[1] - b[1]) / lc * k)
        for t in (0.0, 0.25, 0.5, 0.75, 1.0):
            # a quadratic curve through the corner
            x = (1 - t) ** 2 * p[0] + 2 * (1 - t) * t * b[0] + t ** 2 * q[0]
            y = (1 - t) ** 2 * p[1] + 2 * (1 - t) * t * b[1] + t ** 2 * q[1]
            out.append((x, y))
    if closed:
        out.append(out[0])
    else:
        out.append(ring[-1])
    return out


# ── UI / ENTRY POINT ────────────────────────────────────────────────────────

class BeamWindow(forms.WPFWindow):
    """The Beam Rebar dialogue. Picks hide it; it comes back after."""

    # Class level: the XAML's own Text values fire TextChanged while the
    # window is still loading, before __init__ runs on.
    _ready = False
    _updating = False
    _syncing = False
    update_target = None                # the beam loaded by Update (its id)

    # --- construction ---
    def __init__(self):
        forms.WPFWindow.__init__(self, op.join(SCRIPT_DIR, "BeamRebar.xaml"))
        stream = FileStream(STYLES_XAML, FileMode.Open, FileAccess.Read)
        try:
            self.Resources.MergedDictionaries.Add(XamlReader.Load(stream))
        finally:
            stream.Close()
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)
        set_header_icon(self, SCRIPT_DIR)

        self.action = None
        self._baseline = None
        self._bar_types = sorted(DB.FilteredElementCollector(doc).OfClass(RebarBarType),
                                 key=lambda b: (b.BarNominalDiameter, _name(b)))
        self._hook_types = sorted(DB.FilteredElementCollector(doc).OfClass(RebarHookType),
                                  key=_name)
        self._layouts = engine.layout_names()
        for _k, n in self._layouts:
            self.layout_cb.Items.Add(n)
        self._levels = sorted(DB.FilteredElementCollector(doc).OfClass(DB.Level),
                              key=lambda l: l.ProjectElevation)
        for level in self._levels:
            self.level_cb.Items.Add(_name(level))
        self._types = engine.beam_types(doc)
        for t in self._types:
            self.beam_type_cb.Items.Add(u"{0} : {1}".format(t.FamilyName, _name(t)))
        for _key, combo in BAR_LISTS:
            for b in self._bar_types:
                getattr(self, combo).Items.Add(_name(b))
        for _key, combo in HOOK_LISTS:
            getattr(self, combo).Items.Add(NO_HOOK)
            for h in self._hook_types:
                getattr(self, combo).Items.Add(_name(h))
        self._build_editors()
        self.PreviewMouseDown += self._click_off
        for key in TEXTS:
            box = getattr(self, key + "_tb")
            box.GotKeyboardFocus += self._editor_focused
            box.PreviewMouseLeftButtonDown += self._select_on_click
        self._apply_settings(_load_settings())

    # --- filling the form ---
    def _apply_settings(self, s):
        """Put a settings dict (as settings() makes) into the form. Anything
        missing takes the layout's default."""
        s = _full(s)
        self._ready = False
        try:
            keys = [k for k, _n in self._layouts]
            self.layout_cb.SelectedIndex = keys.index(s["layout"]) if s["layout"] in keys else 0
            for key in TEXTS:
                getattr(self, key + "_tb").Text = unicode(s.get(key, u""))
            for key, combo in BAR_LISTS + HOOK_LISTS:
                cb = getattr(self, combo)
                names = [unicode(i) for i in cb.Items]
                wanted = s.get(key) or NO_HOOK
                if wanted in names:
                    cb.SelectedIndex = names.index(wanted)
                elif combo.endswith("hook_cb"):
                    cb.SelectedIndex = 0
                else:
                    cb.SelectedIndex = 0 if names else -1
            self.grade_cb.Text = s.get("grade", u"500E")
            names = [_name(t) for t in self._types]
            if s.get("beam_type") in names:
                self.beam_type_cb.SelectedIndex = names.index(s["beam_type"])
            else:
                self._match_type()
            names = [_name(l) for l in self._levels]
            current = getattr(doc.ActiveView, "GenLevel", None)
            if s.get("level") in names:
                self.level_cb.SelectedIndex = names.index(s["level"])
            elif current is not None and _name(current) in names:
                self.level_cb.SelectedIndex = names.index(_name(current))
            elif names:
                self.level_cb.SelectedIndex = 0
            if not self._bar_types:
                self.apply_btn.IsEnabled = False
        finally:
            self._ready = True
        self._show_for_layout()
        self._update()

    def _match_type(self):
        """A type of the panel's size, else the first."""
        try:
            size = (float(self.width_tb.Text), float(self.depth_tb.Text))
        except ValueError:
            size = None
        for i, t in enumerate(self._types):
            have = engine.type_size(t)
            if size and have and abs(have[0] - size[0]) < 0.5 and abs(have[1] - size[1]) < 0.5:
                self.beam_type_cb.SelectedIndex = i
                return
        self.beam_type_cb.SelectedIndex = 0 if self._types else -1

    def type_changed(self, sender, args):
        """A type chosen by hand: its size goes into the size boxes."""
        if not self._ready:
            return
        t = self.beam_type()
        size = engine.type_size(t) if t is not None else None
        if size:
            self._ready = False
            try:
                self.width_tb.Text = u"{0:g}".format(round(size[0], 1))
                self.depth_tb.Text = u"{0:g}".format(round(size[1], 1))
            finally:
                self._ready = True
        self._update()

    def beam_type(self):
        i = self.beam_type_cb.SelectedIndex
        return self._types[i] if 0 <= i < len(self._types) else None

    def level(self):
        i = self.level_cb.SelectedIndex
        return self._levels[i] if 0 <= i < len(self._levels) else None

    def placing(self):
        """What Apply needs to place beams: (type, size, level, top offset,
        name). Raises ValueError."""
        t = self.beam_type()
        if t is None:
            raise ValueError(u"The model has no concrete beam type to place. Load a concrete "
                             u"beam family, or use Reinforce existing beams.")
        level = self.level()
        if level is None:
            raise ValueError(u"Choose the level.")
        size = (self._num(self.width_tb, u"Width"), self._num(self.depth_tb, u"Depth"))
        if min(size) <= 0:
            raise ValueError(u"The width and depth must be above 0.")
        return t, size, level, self._num(self.top_offset_tb, u"Top"), (self.name_tb.Text or u"").strip()

    def settings(self):
        """Everything in the form, as strings (what is saved)."""
        t = self.beam_type()
        s = {"layout": self.layout_key(), "grade": self.grade_cb.Text or u"",
             "beam_type": _name(t) if t is not None else u"",
             "level": unicode(self.level_cb.SelectedItem or u"")}
        for key in TEXTS:
            s[key] = getattr(self, key + "_tb").Text or u""
        for key, combo in BAR_LISTS + HOOK_LISTS:
            s[key] = unicode(getattr(self, combo).SelectedItem or u"")
        return s

    def layout_key(self):
        i = self.layout_cb.SelectedIndex
        return self._layouts[i][0] if 0 <= i < len(self._layouts) else self._layouts[0][0]

    def layout_changed(self, sender, args):
        """A new layout: its standard bars and stirrups, keeping the beam."""
        if not self._ready:
            return
        keep = dict((k, v) for k, v in self.settings().items() if k in KEEP_ON_LAYOUT)
        keep["layout"] = self.layout_key()
        self._apply_settings(keep)

    def reset_clicked(self, sender, args):
        if self._baseline is not None:
            self._apply_settings(self._baseline)
        else:
            self._apply_settings({"layout": self.layout_key()})

    def _show_for_layout(self):
        """Show the inputs the layout uses."""
        kind = engine.layout_info(self.layout_key())[1]
        on, off = Visibility.Visible, Visibility.Collapsed
        self.tie_count_row.Visibility = on if kind == u"multi" else off
        self.u_lap_row.Visibility = on if kind == u"two_u" else off
        self.links_panel.Visibility = on if kind == u"links" else off
        self.tie_hook_row.Visibility = off if kind == u"two_u" else on
        self.side_row.Visibility = off if kind == u"flat" else on
        self.side_bar_row.Visibility = off if kind == u"flat" else on
        self.top_row.Visibility = off if kind == u"flat" else on
        self.ties_title_tb.Text = {u"two_u": u"U Bars", u"flat": u"Flat Link"}.get(kind, u"Stirrups")

    # --- reading the form ---
    def _num(self, box, label):
        try:
            return float((box.Text or u"").strip())
        except ValueError:
            raise ValueError(u"{0} is not a number.".format(label))

    def _pick(self, combo, items, offset=0):
        i = combo.SelectedIndex - offset
        return items[i] if 0 <= i < len(items) else None

    def options(self):
        """The form as engine options (mm, Revit types). Raises ValueError."""
        kind = engine.layout_info(self.layout_key())[1]
        n = self._num
        o = {"layout": self.layout_key(),
             "cover_side": n(self.cover_side_tb, u"Side cover"),
             "cover_top": n(self.cover_top_tb, u"Top cover"),
             "cover_bottom": n(self.cover_bottom_tb, u"Bottom cover"),
             "end_cover": n(self.end_cover_tb, u"Bar ends"),
             "main_bar": self._pick(self.main_bar_cb, self._bar_types),
             "side_bar": self._pick(self.side_bar_cb, self._bar_types),
             "tie_bar": self._pick(self.tie_bar_cb, self._bar_types),
             "link_bar": self._pick(self.link_bar_cb, self._bar_types),
             "tie_hook": self._pick(self.tie_hook_cb, self._hook_types, 1),
             "link_hook": self._pick(self.link_hook_cb, self._hook_types, 1),
             "top_count": int(n(self.top_count_tb, u"Top bars")) if kind != u"flat" else 0,
             "bottom_count": int(n(self.bottom_count_tb, u"Bottom bars")),
             "side_count": int(n(self.side_count_tb, u"Side bars")) if kind != u"flat" else 0,
             "tie_count": int(n(self.tie_count_tb, u"Stirrups across")),
             "link_count": int(n(self.link_count_tb, u"Links across")),
             "u_lap": n(self.u_lap_tb, u"U bars' lap"),
             "tie_first": n(self.tie_first_tb, u"First stirrup"),
             "tie_sp": n(self.tie_sp_tb, u"Stirrup spacing"),
             "zone_len": n(self.zone_len_tb, u"End zone length"),
             "zone_sp": n(self.zone_sp_tb, u"End zone spacing")}
        if kind == u"two_u":
            o["tie_hook"] = None
        if o["main_bar"] is None or o["tie_bar"] is None:
            raise ValueError(u"Choose the bars and the stirrup bar (the model has no rebar bar "
                             u"types? Load them with Rebar Settings > Load NZ Rebar Sizes).")
        if o["bottom_count"] < 1:
            raise ValueError(u"There must be at least one bottom bar.")
        if kind in (u"closed", u"links", u"multi", u"two_u") and min(o["top_count"], o["bottom_count"]) < 2:
            raise ValueError(u"The stirrup needs a bar in each corner: at least 2 top and 2 bottom bars.")
        if min(o["tie_sp"], o["zone_sp"]) <= 0:
            raise ValueError(u"Spacings must be above 0.")
        if min(o["cover_side"], o["cover_top"], o["cover_bottom"]) < 0:
            raise ValueError(u"The covers cannot be below 0.")
        if kind == u"links" and o["link_count"] > 0 and o["link_hook"] is None:
            raise ValueError(u"Choose the links' hooks (a 180 degree standard hook).")
        return o

    def preview(self):
        """(options, plan) for the drawings and the calculations, with the
        size typed in the panel. Raises ValueError."""
        o = engine.numbers(self.options())
        n = self._num
        o["W"], o["H"], o["L"] = n(self.width_tb, u"Width"), n(self.depth_tb, u"Depth"), \
            n(self.length_tb, u"Length")
        if min(o["W"], o["H"], o["L"]) <= 0:
            raise ValueError(u"The size and length must be above 0.")
        if o["W"] - 2 * (o["cover_side"] + o["tie_dm"] + o["main_dm"]) <= 0 or \
                o["H"] - o["cover_top"] - o["cover_bottom"] - 2 * (o["tie_dm"] + o["main_dm"]) <= 0:
            raise ValueError(u"The bars do not fit inside the cover.")
        return o, engine.layout(o)

    # --- refreshing ---
    def inputs_changed(self, sender, args):
        self._update()

    def _update(self):
        if not self._ready or self._updating:
            return
        self._updating = True
        try:
            self._refresh()
        finally:
            self._updating = False

    def _refresh(self):
        try:
            o, plan = self.preview()
        except ValueError as err:
            self.status_tb.Text = unicode(err)
            self.note_tb.Text = unicode(err)
            self.apply_btn.IsEnabled = False
            for canvas in (self.section_canvas, self.elev_canvas):
                canvas.Children.Clear()
            for box in list(self._editors.values()) + list(self._callouts.values()):
                box.Visibility = Visibility.Collapsed
            for tb in (self.calc_bottom_tb, self.calc_top_tb, self.calc_ties_tb, self.calc_ratio_tb,
                       self.calc_steel_tb, self.calc_conc_tb, self.calc_rate_tb,
                       self.calc_total_steel_tb, self.calc_total_conc_tb, self.calc_note_tb):
                tb.Text = u""
            return
        self.status_tb.Text = u""
        self.note_tb.Text = u""
        self.apply_btn.IsEnabled = True
        rls = self._rls(o)
        self.rl_tb.Text = (u"Top RL {0:.3f} m, underside RL {1:.3f} m".format(rls[0], rls[1])
                           if rls else u"")
        t = self.beam_type()
        have = engine.type_size(t) if t is not None else None
        if t is not None and not (have and abs(have[0] - o["W"]) < 0.5 and abs(have[1] - o["H"]) < 0.5):
            self.rl_tb.Text += u"\nApply uses a type of this size: {0}".format(
                engine.type_name((o["W"], o["H"])))
        self._placed, self._shown = set(), set()
        self._draw_section(o, plan)
        self._draw_elevation(o, plan)
        for key, box in self._editors.items():
            if key not in self._placed:
                box.Visibility = Visibility.Collapsed
        for key, group in self._callouts.items():
            if key not in self._shown:
                group.Visibility = Visibility.Collapsed
        self._calc(o, plan)

    def _rls(self, o):
        """(top RL m, underside RL m) from the level and the top offset."""
        level = self.level()
        try:
            offset = float(self.top_offset_tb.Text)
        except ValueError:
            return None
        if level is None:
            return None
        elev = level.get_Parameter(DB.BuiltInParameter.LEVEL_ELEV)
        top = (elev.AsDouble() if elev is not None else level.Elevation) * MM + offset
        return top / 1000.0, (top - o["H"]) / 1000.0

    # --- calculations ---
    def _calc(self, o, plan):
        area = lambda n, d: n * math.pi * d ** 2 / 4.0
        nb, nt = len(plan["bottom"]["ys"]), len(plan["top"]["ys"])
        d_eff = o["H"] - o["cover_bottom"] - o["tie_d"] - o["main_d"] / 2.0
        ab, at = area(nb, o["main_d"]), area(nt, o["main_d"])
        bar = unicode(self.main_bar_cb.SelectedItem or u"")
        self.calc_bottom_tb.Text = u"{0} x {1}  {2:.2f}%".format(nb, bar, 100.0 * ab / (o["W"] * d_eff))
        self.calc_top_tb.Text = u"{0} x {1}  {2:,.0f} mm²".format(nt, bar, at) if nt else u"none"
        n_ties = len(plan["xs"]) * len(plan["ties"])
        self.calc_ties_tb.Text = u"{0} @ {1:g}".format(n_ties, o["tie_sp"])
        self.calc_ratio_tb.Text = u"{0:.1f}".format(o["L"] / o["H"])
        long_m, tie_m = engine.bar_length(plan, o)
        sd = o["side_bar"].BarNominalDiameter * MM if o.get("side_bar") else o["main_d"]
        n_side = 2 * len(plan["side"]["zs"]) if plan["side"] else 0
        n_long = nb + nt + n_side
        side_m = long_m * n_side / float(n_long) if n_long else 0.0
        kg = STEEL * ((long_m - side_m) * math.pi * (o["main_d"] / 1000.0) ** 2 / 4.0
                      + side_m * math.pi * (sd / 1000.0) ** 2 / 4.0
                      + tie_m * math.pi * (o["tie_d"] / 1000.0) ** 2 / 4.0)
        vol = o["W"] * o["H"] * o["L"] / 1e9
        try:
            beams = max(int(float(self.beams_tb.Text)), 1)
        except ValueError:
            beams = 1
        self.calc_steel_tb.Text = u"{0:,.1f} kg".format(kg)
        self.calc_conc_tb.Text = u"{0:.3f} m³".format(vol)
        self.calc_rate_tb.Text = u"{0:,.0f} kg/m³".format(kg / vol if vol > 0 else 0.0)
        self.calc_total_lbl.Text = u"Steel, {0} beams".format(beams) if beams > 1 else u"Steel, all"
        self.calc_total_steel_tb.Text = u"{0:,.1f} kg".format(kg * beams)
        self.calc_total_conc_tb.Text = u"{0:.3f} m³".format(vol * beams)
        self.calc_note_tb.Text = (u"{0} grade {1}. Bottom steel as a share of W x d. Bar lengths "
                                  u"approximate (bends square, hooks at 10 d); Revit's schedules give "
                                  u"the exact ones.".format(self.name_tb.Text or u"", self.grade_cb.Text or u""))

    # --- the drawings' input boxes and callouts ---
    EDITORS = (("s_w", "width_tb", "section_edit"), ("s_h", "depth_tb", "section_edit"),
               ("s_a", "slab_tb", "section_edit"),
               ("s_x1", "cover_top_tb", "section_edit"), ("s_x2", "cover_side_tb", "section_edit"),
               ("s_x3", "cover_bottom_tb", "section_edit"),
               ("e_len", "length_tb", "elev_edit"), ("e_first", "tie_first_tb", "elev_edit"),
               ("e_zone", "zone_len_tb", "elev_edit"))

    CALLOUTS = (
        ("c_top", "section_edit", (("num", "top_count_tb"), ("w", u"-"), ("bar", "main_bar_cb"),
                                   ("w", u"T1"))),
        ("c_tie", "section_edit", (("bar", "tie_bar_cb"), ("dyn", u"STIRRUPS"), ("br",), ("w", u"@"),
                                   ("num", "tie_sp_tb"), ("w", u"mm CRS"))),
        ("c_side", "section_edit", (("num", "side_count_tb"), ("w", u"-"), ("bar", "side_bar_cb"),
                                    ("w", u"E.F. E1"))),
        ("c_bot", "section_edit", (("num", "bottom_count_tb"), ("w", u"-"), ("bar", "main_bar_cb"),
                                   ("w", u"B1"))),
        ("c_link", "section_edit", (("num", "link_count_tb"), ("w", u"-"), ("bar", "link_bar_cb"),
                                    ("w", u"L1"))),
        ("z_mid", "elev_edit", (("bar", "tie_bar_cb"), ("dyn", u"STIRRUPS"), ("w", u"@"),
                                ("num", "tie_sp_tb"), ("w", u"CRS"))),
        ("z_end", "elev_edit", (("w", u"END ZONES @"), ("num", "zone_sp_tb"), ("w", u"CRS"))),
    )

    def _number_box(self, key, source, width=38):
        box = TextBox()
        box.Style = self.TryFindResource("TextBoxStyle")
        box.Tag = key
        box.Width = width
        box.Height = 20
        box.FontSize = 12
        box.Padding = Thickness(0)
        box.HorizontalContentAlignment = HorizontalAlignment.Center
        box.ToolTip = getattr(self, source).ToolTip
        box.TextChanged += self._editor_handler(key, source)
        box.GotKeyboardFocus += self._editor_focused
        box.LostKeyboardFocus += self._editor_left
        box.PreviewMouseLeftButtonDown += self._select_on_click
        self._sources[key] = source
        return box

    def _build_editors(self):
        """One small input box for each editable dimension and one group for
        each callout, made once and moved by every redraw."""
        self._editors, self._sources = {}, {}
        for key, source, canvas_name in self.EDITORS:
            box = self._number_box(key, source)
            box.Visibility = Visibility.Collapsed
            getattr(self, canvas_name).Children.Add(box)
            self._editors[key] = box
        self._callouts, self._callout_parts, self._callout_words = {}, {}, {}
        ink = self.TryFindResource("BrushTextPrimary")
        for key, canvas_name, parts in self.CALLOUTS:
            group = StackPanel()
            row = StackPanel()
            row.Orientation = Orientation.Horizontal
            group.Children.Add(row)
            synced, words = [], []
            for i, part in enumerate(parts):
                kind = part[0]
                if kind == "br":
                    row = StackPanel()
                    row.Orientation = Orientation.Horizontal
                    row.Margin = Thickness(0, 2, 0, 0)
                    group.Children.Add(row)
                elif kind in ("w", "dyn"):
                    tb = TextBlock()
                    tb.Text = part[1]
                    tb.FontSize = 11
                    tb.FontWeight = FontWeights.SemiBold
                    tb.Foreground = ink
                    tb.VerticalAlignment = VerticalAlignment.Center
                    tb.Margin = Thickness(3, 0, 3, 0)
                    row.Children.Add(tb)
                    if kind == "dyn":
                        words.append(tb)
                elif kind == "bar":
                    combo = ComboBox()
                    combo.Style = self.TryFindResource("ComboBoxStyle")
                    combo.Width = 66
                    combo.Height = 22
                    combo.FontSize = 11
                    combo.ToolTip = u"The bar (the same as in the panel)"
                    combo.SelectionChanged += self._callout_bar_handler(part[1])
                    row.Children.Add(combo)
                    synced.append(("bar", combo, part[1]))
                else:
                    box = self._number_box(key + "_%d" % i, part[1], 34)
                    row.Children.Add(box)
                    synced.append(("num", box, part[1]))
            group.Visibility = Visibility.Collapsed
            getattr(self, canvas_name).Children.Add(group)
            self._callouts[key] = group
            self._callout_parts[key] = synced
            self._callout_words[key] = words

    def _callout_bar_handler(self, source):
        def changed(sender, args):
            if not self._syncing and sender.SelectedIndex >= 0:
                getattr(self, source).SelectedIndex = sender.SelectedIndex
        return changed

    def _sync_callout(self, key, words=None):
        self._syncing = True
        try:
            for kind, control, src in self._callout_parts[key]:
                source = getattr(self, src)
                if kind == "bar":
                    if control.Items.Count != source.Items.Count:
                        control.Items.Clear()
                        for item in source.Items:
                            control.Items.Add(item)
                    if not control.IsDropDownOpen and control.SelectedIndex != source.SelectedIndex:
                        control.SelectedIndex = source.SelectedIndex
                elif not control.IsKeyboardFocused and control.Text != source.Text:
                    control.Text = source.Text
            for tb, text in zip(self._callout_words[key], words or ()):
                tb.Text = text
        finally:
            self._syncing = False

    def _editor_handler(self, key, source):
        def changed(sender, args):
            if not self._syncing:
                getattr(self, source).Text = sender.Text
        return changed

    def _editor_focused(self, sender, args):
        sender.SelectAll()

    def _select_on_click(self, sender, args):
        if not sender.IsKeyboardFocusWithin:
            sender.Focus()
            args.Handled = True

    def _editor_left(self, sender, args):
        self._update()

    def _click_off(self, sender, args):
        element = args.OriginalSource
        while element is not None:
            if isinstance(element, (TextBoxBase, ComboBox, ComboBoxItem, ButtonBase, ScrollBar)):
                return
            if isinstance(element, Visual):
                element = VisualTreeHelper.GetParent(element)
            else:
                element = LogicalTreeHelper.GetParent(element)
        Keyboard.ClearFocus()
        FocusManager.SetFocusedElement(self, self)
        Keyboard.Focus(self)

    def _place_editor(self, key, cx, cy, angle=0.0):
        box = self._editors[key]
        Canvas.SetLeft(box, cx - box.Width / 2.0)
        Canvas.SetTop(box, cy - box.Height / 2.0)
        box.RenderTransformOrigin = Point(0.5, 0.5)
        box.RenderTransform = RotateTransform(angle)
        box.Visibility = Visibility.Visible
        text = getattr(self, self._sources[key]).Text
        if not box.IsKeyboardFocused and box.Text != text:
            self._syncing = True
            try:
                box.Text = text
            finally:
                self._syncing = False
        self._placed.add(key)

    def _callout(self, key, canvas, ax, ay, ex, ey, stop=0.0, words=None, side=1):
        """A leader from the bar (arrowhead on it) to (ex, ey), a short level
        tail and the words, which hold the panel's lists and boxes."""
        group = self._callouts[key]
        self._sync_callout(key, words)
        group.Visibility = Visibility.Visible
        group.Measure(Size(1000.0, 200.0))
        width = group.DesiredSize.Width
        ink = self.TryFindResource("BrushTextPrimary")
        reach = math.hypot(ex - ax, ey - ay) or 1.0
        ux, uy = (ex - ax) / reach, (ey - ay) / reach
        ax, ay = ax + ux * stop, ay + uy * stop
        tail = ex + side * 12.0
        leader = Polyline()
        leader.Points = PointCollection([Point(ax, ay), Point(ex, ey), Point(tail, ey)])
        leader.Stroke = ink
        leader.StrokeThickness = 1
        canvas.Children.Add(leader)
        _arrow(canvas, ink, ax, ay, -ux, -uy, 10.0, 3.6)
        Canvas.SetLeft(group, tail + 3.0 if side > 0 else tail - 3.0 - width)
        first_row = group.Children[0]
        first_row.Measure(Size(1000.0, 200.0))
        Canvas.SetTop(group, ey - first_row.DesiredSize.Height / 2.0)
        self._shown.add(key)

    def _bar_brushes(self):
        accent = self.TryFindResource("BrushPrimaryGreen")
        c = getattr(accent, "Color", None) or Color.FromRgb(0x20, 0x8A, 0x3C)
        edge = Color.FromRgb(int(c.R * 0.45), int(c.G * 0.45), int(c.B * 0.45))
        return SolidColorBrush(c), SolidColorBrush(edge)

    def _bar_line(self, canvas, pts, width=4.0, opacity=1.0):
        fill, edge = self._bar_brushes()
        for brush, w in ((edge, width + 2.2), (fill, width)):
            pl = Polyline()
            pl.Points = PointCollection([Point(x, y) for x, y in pts])
            pl.Stroke = brush
            pl.StrokeThickness = w
            pl.StrokeLineJoin = PenLineJoin.Round
            pl.StrokeStartLineCap = pl.StrokeEndLineCap = PenLineCap.Round
            pl.Opacity = opacity
            canvas.Children.Add(pl)

    def _bar_dot(self, canvas, x, y, size):
        fill, edge = self._bar_brushes()
        e = _ellipse(x, y, size / 2.0)
        e.Fill = fill
        e.Stroke = edge
        e.StrokeThickness = 1.2
        canvas.Children.Add(e)

    # --- Section: to scale, the bars and stirrups as the panel sets them ---
    def _draw_section(self, o, plan):
        canvas = self.section_canvas
        canvas.Children.Clear()
        ink = self.TryFindResource("BrushTextPrimary")
        W, H = o["W"], o["H"]
        try:
            A = max(float(self.slab_tb.Text), 0.0)
        except ValueError:
            A = 0.0
        A = min(A, H)
        k = min(210.0 / W, 240.0 / H)
        cx, top = 190.0, 92.0

        def X(y):                        # the section's y (left +) to canvas x: left is left
            return cx - y * k

        def Y(z):
            return top + (H - z) * k
        x0, x1, yt, yb = X(W / 2.0), X(-W / 2.0), Y(H), Y(0.0)
        # the slab on it, either side, broken off at the canvas edges
        slab_l, slab_r = max(x0 - 80.0, 4.0), min(x1 + 30.0, 318.0)
        if A > 0:
            ys = Y(H - A)
            for xa, xb in ((slab_l, x0), (x1, slab_r)):
                if xb - xa > 1.0:
                    _concrete_shape(canvas, _rect(xa, yt, xb, ys))
                    _grain(canvas, _rect(xa, yt, xb, ys))
            # the mesh in the slab, and FFL
            xm = slab_l + 8.0
            while xm < slab_r - 4.0:
                if xm < x0 - 4 or xm > x1 + 4:
                    e = _ellipse(xm, yt + max(A * k * 0.35, 5.0), 1.8)
                    e.Stroke = ink
                    e.StrokeThickness = 0.8
                    canvas.Children.Add(e)
                xm += 18.0
            _label(canvas, u"FFL", slab_r - 22.0, yt - 16.0, 10, True, ink=ink)
        body = _polygon([(x0, yt), (x1, yt), (x1, yb), (x0, yb)])
        _concrete_shape(canvas, body)
        _grain(canvas, _polygon([(x0, yt), (x1, yt), (x1, yb), (x0, yb)]))
        if A > 0:
            # the slab and the beam are one pour: no line between them
            _line(canvas, CONCRETE_BASE, x0 + 1.0, yt, x1 - 1.0, yt, 3.0)
            _line(canvas, ink, slab_l, yt, slab_r, yt, 0.9, 0.8)
        # stirrups, links, U bars: their centrelines, bent round
        tw = max(o["tie_dm"] * k, 2.4)
        lw = max(o["link_dm"] * k, 2.4)
        bend = max(2.0 * o["tie_dm"] * k, 4.0)
        for t in plan["ties"]:
            pts = [(X(y), Y(z)) for y, z in t["pts"]]
            if t["role"] == u"link":
                # the straight leg and a 180 degree turn round each bar, centred on it
                zt_, zb_ = t["zs"]
                cxb = X(t["bar_y"])
                s = 1.0 if X(t["pts"][0][0]) > cxb else -1.0     # the leg's side of the bars
                rad = t["r"] * k
                angles = [math.pi * i / 8.0 for i in range(9)]
                top_turn = [(cxb + s * rad * math.cos(a), Y(zt_) - rad * math.sin(a)) for a in angles]
                bot_turn = [(cxb + s * rad * math.cos(a), Y(zb_) + rad * math.sin(a)) for a in angles]
                leg = 6 * o["link_dm"] * k
                self._bar_line(canvas, [(top_turn[-1][0], Y(zt_) + leg)] + list(reversed(top_turn))
                               + bot_turn + [(bot_turn[-1][0], Y(zb_) - leg)], lw)
                continue
            self._bar_line(canvas, _rounded(pts, t["closed"], bend), tw)
            if t["closed"] and t["role"] == u"tie":
                # the 135 degree hooks at the stirrup's first corner, turned in
                hx, hy = pts[0]
                sx = 1.0 if (pts[2][0] - hx) > 0 else -1.0
                hk = max(7 * o["tie_dm"] * k, 9.0)
                self._bar_line(canvas, [(hx + sx * 2.0, hy + 2.0),
                                        (hx + sx * hk * 0.75, hy + hk * 0.75)], tw - 0.6)
                self._bar_line(canvas, [(hx + sx * 4.0, hy),
                                        (hx + sx * (4.0 + hk * 0.7), hy + hk * 0.7)], tw - 0.6)
        size = max(o["main_dm"] * k, 9.0)
        for row in (plan["top"], plan["bottom"]):
            for y in row["ys"]:
                self._bar_dot(canvas, X(y), Y(row["z"]), size)
        if plan["side"]:
            ss = max((o["side_bar"].BarModelDiameter * MM if o.get("side_bar") else o["main_dm"]) * k, 6.0)
            for z in plan["side"]["zs"]:
                for y in (plan["side"]["y"], -plan["side"]["y"]):
                    self._bar_dot(canvas, X(y), Y(z), ss)

        # --- dimensions: W under it, H (and A) on the left ---
        yd = yb + 34.0
        for x in (x0, x1):
            _ext_line(canvas, ink, x, yb + 4.0, x, yd + 6.0)
        _dim_line(canvas, ink, x0, yd, x1, yd)
        self._place_editor("s_w", cx, yd)
        xd = x0 - 34.0
        for y in (yt, yb):
            _ext_line(canvas, ink, x0 - 4.0 if y == yb else xd + 6.0, y, xd - 6.0, y)
        _dim_line(canvas, ink, xd, yt, xd, yb)
        self._place_editor("s_h", xd, (yt + yb) / 2.0, -90.0)
        if A > 0:
            xa = xd - 28.0
            ys = Y(H - A)
            _ext_line(canvas, ink, 14.0, ys, xa - 6.0, ys)
            _dim_line(canvas, ink, xa, yt, xa, ys)
            self._place_editor("s_a", xa, (yt + ys) / 2.0 if (ys - yt) > 44 else ys + 16.0, -90.0)
        # covers: X1 at the top, X2 on the right side, X3 at the bottom
        tz1 = Y(H - o["cover_top"])
        tz0 = Y(o["cover_bottom"])
        ty1 = X(-(W / 2.0 - o["cover_side"]))
        xm = cx + W * k * 0.18
        _dim_outside(canvas, ink, xm, yt, xm, tz1)
        self._place_editor("s_x1", xm + 24.0, yt - 22.0)
        _dim_outside(canvas, ink, xm, yb, xm, tz0)
        self._place_editor("s_x3", xm + 24.0, yb + 14.0)
        ym = Y(H * 0.42)
        _dim_outside(canvas, ink, x1, ym, ty1, ym)
        self._place_editor("s_x2", x1 + 22.0, ym - 14.0)

        # --- callouts, on the right ---
        col = 330.0
        kind = plan["kind"]
        tie_word = {u"two_u": u"U BARS", u"flat": u"LINKS"}.get(kind, u"STIRRUPS")
        # each callout points at the bar nearest it: the right side is -y
        if plan["top"]["ys"]:
            self._callout("c_top", canvas, X(plan["top"]["ys"][0]), Y(plan["top"]["z"]),
                          col, Y(plan["top"]["z"]) - 30.0, size / 2.0)
        if plan["ties"]:
            t = [t for t in plan["ties"] if t["role"] != u"link"][-1]
            ry = min(y for y, _z in t["pts"])             # its right leg (right is -y)
            self._callout("c_tie", canvas, X(ry), Y(H * 0.62) if kind != u"flat" else Y(plan["bottom"]["z"]),
                          col, yt - 56.0 if A > 0 else yt - 20.0, tw / 2.0, [tie_word])
        if plan["side"]:
            self._callout("c_side", canvas, X(-plan["side"]["y"]), Y(plan["side"]["zs"][0]),
                          col, Y(plan["side"]["zs"][0]) + 6.0, size / 2.0)
        self._callout("c_bot", canvas, X(plan["bottom"]["ys"][0]), Y(plan["bottom"]["z"]),
                      col, Y(plan["bottom"]["z"]) + 34.0, size / 2.0)
        links = [t for t in plan["ties"] if t["role"] == u"link"]
        if links:
            t = links[0]
            self._callout("c_link", canvas, X(t["pts"][0][0]), Y(H * 0.45), col, yb + 62.0, lw / 2.0)

    # --- Elevation: along the beam, to scale along it ---
    def _draw_elevation(self, o, plan):
        canvas = self.elev_canvas
        canvas.Children.Clear()
        ink = self.TryFindResource("BrushTextPrimary")
        L, H = o["L"], o["H"]
        xa, xb = 20.0, canvas.Width - 20.0
        ya, yb = 46.0, 116.0                        # top and underside (depth not to scale)
        kx = (xb - xa) / L
        kz = (yb - ya) / H

        def X(x):
            return xa + x * kx

        def Y(z):
            return yb - z * kz
        _concrete_shape(canvas, _rect(xa, ya, xb, yb))
        _grain(canvas, _rect(xa, ya, xb, yb))
        # stirrups: seen side on, from their lowest to their highest point
        tz = [z for t in plan["ties"] for _y, z in t["pts"]] or [o["cover_bottom"], H - o["cover_top"]]
        z0, z1 = max(min(tz), 0.0), min(max(tz), H)
        tw = 1.6
        for x in plan["xs"]:
            self._bar_line(canvas, [(X(x), Y(z0)), (X(x), Y(z1))], tw)
        # top and bottom bars, and a side bar row
        for row in (plan["top"], plan["bottom"]):
            if row["ys"]:
                self._bar_line(canvas, [(X(plan["x0"]), Y(row["z"])), (X(plan["x1"]), Y(row["z"]))], 2.6)
        if plan["side"]:
            for z in plan["side"]["zs"]:
                self._bar_line(canvas, [(X(plan["x0"]), Y(z)), (X(plan["x1"]), Y(z))], 2.0, 0.8)
        # L under it
        yd = yb + 30.0
        for x in (xa, xb):
            _ext_line(canvas, ink, x, yb + 4.0, x, yd + 6.0)
        _dim_line(canvas, ink, xa, yd, xb, yd)
        self._place_editor("e_len", (xa + xb) / 2.0, yd)
        # the first stirrup in from the start, small, arrows from outside
        xf = X(o["tie_first"])
        yf = ya - 12.0
        _ext_line(canvas, ink, xa, ya - 3.0, xa, yf - 6.0)
        _ext_line(canvas, ink, xf, ya - 3.0, xf, yf - 6.0)
        _dim_outside(canvas, ink, xa, yf, xf, yf)
        self._place_editor("e_first", xa + 16.0, yf - 18.0)
        # the end zone, when there is one
        zones = [s for s in plan["sets"] if abs(s[2] - o["zone_sp"]) < 1e-6 and o["zone_len"] > 0]
        if zones:
            za, zb_ = zones[0][0], zones[0][1]
            _ext_line(canvas, ink, X(zb_), ya - 3.0, X(zb_), yf - 6.0)
            _dim_line(canvas, ink, X(za), yf, X(zb_), yf)
            self._place_editor("e_zone", max((X(za) + X(zb_)) / 2.0, xa + 60.0), yf - 18.0)
            self._callout("z_end", canvas, X(L - (za + zb_) / 2.0), ya + 6.0, xb - 150.0, 18.0, 1.0)
        mid = plan["xs"][len(plan["xs"]) // 2] if plan["xs"] else L / 2.0
        self._callout("z_mid", canvas, X(mid), Y(H * 0.3), X(mid) - 180.0, yb + 58.0, 1.0,
                      [{u"two_u": u"U BARS", u"flat": u"LINKS"}.get(plan["kind"], u"STIRRUPS")])

    # --- buttons: each pick hides the window and is done by main() ---
    def _hide_for(self, action):
        self.action = action
        self.Hide()

    def update_pick_clicked(self, sender, args):
        self._hide_for("pick_update")

    def apply_clicked(self, sender, args):
        self._hide_for("apply")

    def reinforce_existing_clicked(self, sender, args):
        self._hide_for("reinforce_existing")

    def close_clicked(self, sender, args):
        self.action = None
        self.Close()

    # --- a beam already in the model ---
    def load_beam(self, host, frame):
        """Fill the form from a beam: its size, type, level and Mark. Apply
        then replaces its bars with these."""
        s = self.settings()
        s["width"] = u"{0:g}".format(round(frame.width, 1))
        s["depth"] = u"{0:g}".format(round(frame.depth, 1))
        s["length"] = u"{0:g}".format(round(frame.length, 1))
        mark = host.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
        s["name"] = mark.AsString() if mark is not None and mark.AsString() else u""
        s["beam_type"] = _name(host.Symbol)
        level, offset = engine.beam_level(doc, host, frame)
        if level is not None:
            s["level"] = _name(level)
            s["top_offset"] = u"{0:g}".format(round(offset, 1))
        self._apply_settings(s)
        self.update_target = host.Id
        self.apply_btn.ToolTip = u"Replace the bars in beam {0} with these".format(engine.eid(host.Id))
        self.header_subtitle_tb.Text = u"| Editing {0} ({1})".format(
            s["name"] or _name(host.Symbol), engine.eid(host.Id))
        self.apply_btn.Content = (u"Apply to {0}".format(s["name"])
                                  if s["name"] and len(s["name"]) < 10 else u"Apply")
        self._baseline = self.settings()


# ── PICKING AND PLACING ─────────────────────────────────────────────────────

class _BeamFilter(ISelectionFilter):
    """Structural framing that can host rebar."""

    def AllowElement(self, element):
        if not isinstance(element, DB.FamilyInstance) or element.Category is None:
            return False
        if engine.eid(element.Category.Id) != int(DB.BuiltInCategory.OST_StructuralFraming):
            return False
        data = RebarHostData.GetRebarHostData(element)
        return data is not None and data.IsValidHost()

    def AllowReference(self, reference, position):
        return False


def _pick_update(window):
    with forms.WarningBar(title=BAR_UPDATE):
        ref = uidoc.Selection.PickObject(ObjectType.Element, _BeamFilter(), BAR_UPDATE)
    host = doc.GetElement(ref)
    window.load_beam(host, engine.from_host(doc, host))


def _has_bars(host):
    data = RebarHostData.GetRebarHostData(host)
    return data is not None and data.GetRebarsInHost().Count > 0


def _pick_xyz(prompt):
    """A point: on the work plane in a plan, or snapped to a model face
    anywhere else (a 3D view has no work plane to click on)."""
    if isinstance(doc.ActiveView, DB.ViewPlan):
        return uidoc.Selection.PickPoint(prompt)
    return uidoc.Selection.PickObject(ObjectType.PointOnElement, prompt).GlobalPoint


def _marks(name, n):
    if not name:
        return [None] * n
    return [name] if n == 1 else [u"{0}-{1}".format(name, i + 1) for i in range(n)]


def _place(window):
    """Click the start and end of each new beam; each is placed (its type
    made if this size is new) and reinforced. One undo."""
    opts = window.options()
    symbol, size, level, top_offset, name = window.placing()
    pairs = []
    with forms.WarningBar(title=BAR_START.format(1, 0)) as bar:
        while True:
            bar.message_tb.Text = BAR_START.format(len(pairs) + 1, len(pairs))
            try:
                a = _pick_xyz(BAR_START.format(len(pairs) + 1, len(pairs)))
            except OperationCanceledException:
                break
            bar.message_tb.Text = BAR_END.format(len(pairs) + 1)
            try:
                b = _pick_xyz(BAR_END.format(len(pairs) + 1))
            except OperationCanceledException:
                break
            if DB.XYZ(b.X - a.X, b.Y - a.Y, 0).GetLength() * MM < 100.0:
                continue
            pairs.append((a, b))
    if not pairs:
        return False
    t = DB.Transaction(doc, u"Beam Rebar - place")
    t.Start()
    try:
        made_type = engine.ensure_type(doc, symbol, size)
        hosts = [engine.place_beam(doc, made_type, a, b, level, top_offset, mark)
                 for (a, b), mark in zip(pairs, _marks(name, len(pairs)))]
        doc.Regenerate()
        for host in hosts:
            engine.reinforce(doc, host, engine.from_host(doc, host), opts)
        t.Commit()
    except Exception as err:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        dlg.message(u"The beams could not be made, nothing was changed.\n\n{0}".format(err), title=TITLE)
        return False
    return True


def _update(window):
    """Replace the bars of the beam loaded by Update. One undo."""
    host = doc.GetElement(window.update_target)
    if host is None:
        window.update_target = None
        dlg.message(u"That beam is no longer in the model.", title=TITLE)
        return False
    opts = window.options()
    t = DB.Transaction(doc, u"Beam Rebar - update")
    t.Start()
    try:
        engine.remove_bars(doc, host)
        engine.reinforce(doc, host, engine.from_host(doc, host), opts)
        t.Commit()
    except Exception as err:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        dlg.message(u"The beam could not be updated, nothing was changed.\n\n{0}".format(err),
                    title=TITLE)
        return False
    return True


def _reinforce_existing(window):
    """Pick beams already in the model and reinforce them, each at its own
    size. One undo."""
    opts = window.options()
    with forms.WarningBar(title=BAR_PICK):
        refs = uidoc.Selection.PickObjects(ObjectType.Element, _BeamFilter(), BAR_PICK)
    hosts = [doc.GetElement(r) for r in refs]
    if not hosts:
        return False
    frames = [(h, engine.from_host(doc, h)) for h in hosts]
    with_bars = [h for h, _f in frames if _has_bars(h)]
    if with_bars:
        if not dlg.confirm(u"{0} of the {1} already have bars in them. Replace them?".format(
                len(with_bars), len(hosts)), title=TITLE, yes=u"Replace", no=u"Cancel"):
            return False
    t = DB.Transaction(doc, u"Beam Rebar")
    t.Start()
    try:
        for host, frame in frames:
            engine.remove_bars(doc, host)
            engine.reinforce(doc, host, frame, opts)
        t.Commit()
    except Exception as err:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        dlg.message(u"The bars could not be made, nothing was changed.\n\n{0}".format(err), title=TITLE)
        return False
    return True


def main():
    window = BeamWindow()
    while True:
        window.action = None
        window.ShowDialog()
        action = window.action
        if action is None:
            break
        _save_settings(window.settings())
        try:
            if action == "pick_update":
                _pick_update(window)
            elif action == "apply" and window.update_target is not None:
                if _update(window):
                    break
            elif action == "apply":
                if _place(window):
                    break
            elif action == "reinforce_existing":
                if _reinforce_existing(window):
                    break
        except OperationCanceledException:
            pass                        # ESC: back to the window
        except ValueError as err:
            dlg.message(unicode(err), title=TITLE)
    _save_settings(window.settings())


if __name__ == "__main__":
    main()
