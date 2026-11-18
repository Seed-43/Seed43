# -*- coding: utf-8 -*-
# "Column Rebar"
# "Seed43"
# """
# Reinforce concrete columns and piles with native Revit rebar from a
# dialogue: pick a rebar layout (bored pile with a spiral, square column with
# ties and starters, round column with circular ties), set the bars, ties
# and distribution zones, and Apply to one or more picked columns. Update
# reads a column's own size and bars back into the dialogue.
# """
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
                            HorizontalAlignment, VerticalAlignment, Rect, CornerRadius)
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

import column_engine as engine

doc = revit.doc
uidoc = revit.uidoc

# ── CONSTANTS ───────────────────────────────────────────────────────────────

SCRIPT_DIR = op.dirname(__file__)
TITLE = u"Column Rebar"
MM = 304.8
SETTINGS = _userdata.user_path(u"Column Rebar", u"settings.json")
STYLES_XAML = op.join(op.dirname(op.abspath(dlg.__file__)), u"Seed43Styles.xaml")
NO_HOOK = u"(no hook)"
STEEL = 7850.0                                  # kg/m3

# the window's own settings; each layout's defaults (engine.LAYOUTS) go over these
DEFAULTS = {"layout": engine.LAYOUTS[0][0], "name": u"P1", "diameter": u"600", "b": u"500",
            "h": u"500", "length": u"7000", "piles": u"1", "top_offset": u"0", "grade": u"500E",
            "column_type": u"", "level": u"",
            "main_count": u"8", "bars_b": u"4", "bars_h": u"4", "arrangement": u"axes",
            "tie_hook": u"Stirrup/Tie Seismic - 135 deg.", "inner_ties": True, "turns": u"3",
            "starters": False, "starter_bar": u"HD20", "starter_embed": u"480",
            "starter_foot": u"320", "starter_lap": u"1300", "starter_dir": u"out"}
DEFAULTS.update(engine.LAYOUTS[0][3])

# text settings and the box each one sits in (key + "_tb")
TEXTS = ("name", "diameter", "b", "h", "length", "piles", "top_offset", "cover_side", "cover_top",
         "cover_bottom", "main_count", "bars_b", "bars_h", "hook_leg", "lap", "bottom_ext",
         "tie_first", "tie_last", "turns", "zone_bottom_len", "zone_bottom_sp", "zone_mid_sp",
         "zone_top_len", "zone_top_sp", "starter_embed", "starter_foot", "starter_lap")
# the bar and hook lists: (setting, combo)
BAR_LISTS = (("main_bar", "main_bar_cb"), ("tie_bar", "tie_bar_cb"), ("starter_bar", "starter_bar_cb"))
HOOK_LISTS = (("main_hook", "main_hook_cb"), ("tie_hook", "tie_hook_cb"))
# fixed choices: (setting, combo, the values in the combo's order)
CHOICES = (("arrangement", "arrangement_cb", (u"axes", u"between")),
           ("hook_dir", "hook_dir_cb", (u"in", u"out")),
           ("starter_dir", "starter_dir_cb", (u"out", u"in")))
SWITCHES = (("inner_ties", "inner_ties_cb"), ("starters", "starters_cb"))

TIE_WORDS = {u"spiral": (u"Spiral", u"SPIRAL @", u"mm PITCH"),
             u"circle": (u"Circular Ties", u"TIES @", u"mm CRS"),
             u"rect": (u"Stirrups", u"STIRRUPS @", u"mm CRS")}

BAR_PICK = u"Column Rebar: select the columns or piles to reinforce, then Finish. ESC goes back."
BAR_PLACE = (u"Column Rebar: click the centre of each new pile or column ({0} so far). "
             u"ESC when done.")
BAR_UPDATE = u"Column Rebar: click the column or pile to update. Its size and bars fill the dialogue. ESC goes back."

# concrete in the drawings: the same greys on a dark or a light theme
CONCRETE_BASE = SolidColorBrush(Color.FromRgb(0xC4, 0xC6, 0xC8))
CONCRETE_EDGE = SolidColorBrush(Color.FromRgb(0x3A, 0x40, 0x48))


def _speckled_concrete():
    """The grain of concrete as on the section drawings: a fine speckle of
    darker and lighter grains, tiled, drawn over CONCRETE_BASE (as the
    Concrete Stair tool's details)."""
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
    """A small dimension: the line between the two points, and an arrow at
    each end coming in from outside (its tail beyond the end, its point on it)."""
    _line(canvas, ink, x1, y1, x2, y2, 1.0, 0.95)
    run = math.hypot(x2 - x1, y2 - y1) or 1.0
    for px, py, ux, uy in ((x1, y1, (x2 - x1) / run, (y2 - y1) / run),
                           (x2, y2, (x1 - x2) / run, (y1 - y2) / run)):
        _line(canvas, ink, px - ux * 18.0, py - uy * 18.0, px, py)
        _arrow(canvas, ink, px, py, ux, uy)


def _ext_line(canvas, ink, x1, y1, x2, y2):
    """A thin extension line from a feature out to its dimension line."""
    _line(canvas, ink, x1, y1, x2, y2, 1.0, 0.5)


def _centre_line(canvas, ink, x1, y1, x2, y2):
    """A dash-dot centreline."""
    _line(canvas, ink, x1, y1, x2, y2, 0.8, 0.6, [14.0, 3.0, 2.0, 3.0])


def _dim_angle(x1, y1, x2, y2):
    """The angle (degrees, -90 up to 90) a dimension's box is turned by to run
    along its line, reading left to right or upwards."""
    ang = math.degrees(math.atan2(y2 - y1, x2 - x1))
    while ang >= 90.0:
        ang -= 180.0
    while ang < -90.0:
        ang += 180.0
    return ang


def _concrete_shape(canvas, shape):
    """Fill a Polygon / Ellipse / Rectangle as concrete: grey, then the
    speckle over it, with a dark edge."""
    shape.Fill = CONCRETE_BASE
    shape.Stroke = CONCRETE_EDGE
    shape.StrokeThickness = 1.5
    canvas.Children.Add(shape)


def _grain(canvas, shape):
    shape.Fill = CONCRETE_SPECKLED
    shape.IsHitTestVisible = False
    canvas.Children.Add(shape)


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


# ── UI / ENTRY POINT ────────────────────────────────────────────────────────

class ColumnWindow(forms.WPFWindow):
    """The Column Rebar dialogue. Picks hide it; it comes back after."""

    # Class level, not just set in __init__: the XAML's own Text values fire
    # TextChanged while the window is still loading, before __init__ runs on.
    _ready = False
    _updating = False
    _syncing = False                    # a drawing box is being filled from its panel input
    update_target = None                # the column loaded by Update (its id)

    # --- construction ---
    def __init__(self):
        forms.WPFWindow.__init__(self, op.join(SCRIPT_DIR, "ColumnRebar.xaml"))
        stream = FileStream(STYLES_XAML, FileMode.Open, FileAccess.Read)
        try:
            self.Resources.MergedDictionaries.Add(XamlReader.Load(stream))
        finally:
            stream.Close()
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)
        set_header_icon(self, SCRIPT_DIR)

        self.action = None              # what to do once the window is hidden
        self._baseline = None           # what Reset goes back to; None = the layout's defaults
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
        self._all_types = engine.column_types(doc)
        self._types = []                # the column types shown: the layout's shape
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
        missing takes the layout's default. Used at start, by Reset, by a
        layout change and by Update."""
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
            for key, combo, values in CHOICES:
                getattr(self, combo).SelectedIndex = (values.index(s[key]) if s.get(key) in values else 0)
            for key, box in SWITCHES:
                getattr(self, box).IsChecked = bool(s.get(key))
            self.grade_cb.Text = s.get("grade", u"500E")
            self._fill_types(s.get("column_type", u""))
            # the saved level if it is still there, else the active view's level
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

    def _fill_types(self, wanted):
        """The column types of the layout's shape (round or rectangular); the
        one named `wanted` chosen, else one of the panel's size, else the first."""
        shape = engine.layout_info(self.layout_key())[1]
        self._types = [t for t in self._all_types if engine.type_shape(t) in (shape, None)]
        self.column_type_cb.Items.Clear()
        for t in self._types:
            self.column_type_cb.Items.Add(u"{0} : {1}".format(t.FamilyName, _name(t)))
        names = [_name(t) for t in self._types]
        if wanted in names:
            self.column_type_cb.SelectedIndex = names.index(wanted)
            return
        try:
            if shape == u"round":
                size = (float(self.diameter_tb.Text),) * 2
            else:
                size = (float(self.b_tb.Text), float(self.h_tb.Text))
        except ValueError:
            size = None
        for i, t in enumerate(self._types):
            have = engine.type_size(t)
            if size and have and abs(have[0] - size[0]) < 0.5 and abs(have[1] - size[1]) < 0.5:
                self.column_type_cb.SelectedIndex = i
                return
        self.column_type_cb.SelectedIndex = 0 if self._types else -1

    def type_changed(self, sender, args):
        """A type chosen by hand: its size goes into the size boxes."""
        if not self._ready:
            return
        t = self.column_type()
        size = engine.type_size(t) if t is not None else None
        if size:
            self._ready = False
            try:
                self.diameter_tb.Text = u"{0:g}".format(round(size[0], 1))
                self.b_tb.Text = u"{0:g}".format(round(size[0], 1))
                self.h_tb.Text = u"{0:g}".format(round(size[1], 1))
            finally:
                self._ready = True
        self._update()

    def column_type(self):
        i = self.column_type_cb.SelectedIndex
        return self._types[i] if 0 <= i < len(self._types) else None

    def level(self):
        i = self.level_cb.SelectedIndex
        return self._levels[i] if 0 <= i < len(self._levels) else None

    def size(self):
        """(D, D) or (b, h) of the pile in the panel, mm. Raises ValueError."""
        shape = engine.layout_info(self.layout_key())[1]
        if shape == u"round":
            d = self._num(self.diameter_tb, u"Diameter")
            return shape, (d, d)
        return shape, (self._num(self.b_tb, u"Width b"), self._num(self.h_tb, u"Depth h"))

    def placing(self):
        """What Apply needs to place a pile: (type, shape, size, level, top
        offset, length, name). Raises ValueError."""
        t = self.column_type()
        if t is None:
            raise ValueError(u"The model has no concrete column type of this shape to place. "
                             u"Load a concrete column family, or use Reinforce existing columns.")
        level = self.level()
        if level is None:
            raise ValueError(u"Choose the level.")
        shape, size = self.size()
        length = self._num(self.length_tb, u"Length")
        if min(size) <= 0 or length <= 0:
            raise ValueError(u"The size and length must be above 0.")
        return (t, shape, size, level, self._num(self.top_offset_tb, u"Top"), length,
                (self.name_tb.Text or u"").strip())

    def settings(self):
        """Everything in the form, as strings and flags (what is saved)."""
        t = self.column_type()
        s = {"layout": self.layout_key(), "grade": self.grade_cb.Text or u"",
             "column_type": _name(t) if t is not None else u"",
             "level": unicode(self.level_cb.SelectedItem or u"")}
        for key in TEXTS:
            s[key] = getattr(self, key + "_tb").Text or u""
        for key, combo in BAR_LISTS + HOOK_LISTS:
            s[key] = unicode(getattr(self, combo).SelectedItem or u"")
        for key, combo, values in CHOICES:
            i = getattr(self, combo).SelectedIndex
            s[key] = values[i] if 0 <= i < len(values) else values[0]
        for key, box in SWITCHES:
            s[key] = bool(getattr(self, box).IsChecked)
        return s

    def layout_key(self):
        i = self.layout_cb.SelectedIndex
        return self._layouts[i][0] if 0 <= i < len(self._layouts) else self._layouts[0][0]

    def layout_changed(self, sender, args):
        """A new layout: its standard bars and ties, keeping the pile's own
        size, name, count and cap."""
        if not self._ready:
            return
        keep = dict((k, v) for k, v in self.settings().items()
                    if k in ("name", "diameter", "b", "h", "length", "piles", "top_offset", "grade",
                             "level", "column_type"))
        keep["layout"] = self.layout_key()
        self._apply_settings(keep)

    def reset_clicked(self, sender, args):
        """Back to what the selected column has now; with none selected, the
        layout's standard settings."""
        if self._baseline is not None:
            self._apply_settings(self._baseline)
        else:
            keep = {"layout": self.layout_key()}
            self._apply_settings(keep)

    def _show_for_layout(self):
        """Show the inputs the layout's shape and ties use."""
        _n, shape, d = engine.layout_info(self.layout_key())
        kind = d["tie_kind"]
        on, off = Visibility.Visible, Visibility.Collapsed
        rnd = shape == u"round"
        for name in ("diameter_row", "count_row", "arrangement_row"):
            getattr(self, name).Visibility = on if rnd else off
        for name in ("b_row", "h_row", "bars_b_row", "bars_h_row", "starters_panel"):
            getattr(self, name).Visibility = off if rnd else on
        self.tie_hook_row.Visibility = on if kind == u"rect" else off
        self.inner_ties_cb.Visibility = on if kind == u"rect" else off
        self.turns_row.Visibility = on if kind == u"spiral" else off
        self.zone_top_row.Visibility = off if kind == u"spiral" else on
        self.zone_bottom_row.Visibility = off if kind == u"spiral" else on
        self.zone_mid_lbl.Text = u"Pitch" if kind == u"spiral" else u"Middle: spacing"
        self.ties_title_tb.Text = TIE_WORDS[kind][0]
        self.calc_ties_lbl.Text = u"Spiral" if kind == u"spiral" else u"Ties"

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
        _n, shape, d = engine.layout_info(self.layout_key())
        n = self._num
        o = {"layout": self.layout_key(), "tie_kind": d["tie_kind"],
             "cover_side": n(self.cover_side_tb, u"Side cover"),
             "cover_top": n(self.cover_top_tb, u"Top cover"),
             "cover_bottom": n(self.cover_bottom_tb, u"Bottom cover"),
             "main_bar": self._pick(self.main_bar_cb, self._bar_types),
             "main_count": n(self.main_count_tb, u"Quantity"),
             "bars_b": n(self.bars_b_tb, u"Bars along b"), "bars_h": n(self.bars_h_tb, u"Bars along h"),
             "main_hook": self._pick(self.main_hook_cb, self._hook_types, 1),
             "hook_leg": n(self.hook_leg_tb, u"Hook leg"),
             "lap": n(self.lap_tb, u"Lap splice"), "bottom_ext": n(self.bottom_ext_tb, u"Into the cap"),
             "tie_bar": self._pick(self.tie_bar_cb, self._bar_types),
             "tie_hook": self._pick(self.tie_hook_cb, self._hook_types, 1),
             "tie_first": n(self.tie_first_tb, u"First tie"), "tie_last": n(self.tie_last_tb, u"Last tie"),
             "turns": n(self.turns_tb, u"Finishing turns"),
             "zone_bottom_len": n(self.zone_bottom_len_tb, u"Bottom zone length"),
             "zone_bottom_sp": n(self.zone_bottom_sp_tb, u"Bottom zone spacing"),
             "zone_mid_sp": n(self.zone_mid_sp_tb, u"Middle spacing"),
             "zone_top_len": n(self.zone_top_len_tb, u"Top zone length"),
             "zone_top_sp": n(self.zone_top_sp_tb, u"Top zone spacing"),
             "inner_ties": bool(self.inner_ties_cb.IsChecked),
             "starters": bool(self.starters_cb.IsChecked) and shape == u"rect",
             "starter_bar": self._pick(self.starter_bar_cb, self._bar_types),
             "starter_embed": n(self.starter_embed_tb, u"Starter foot depth"),
             "starter_foot": n(self.starter_foot_tb, u"Starter foot length"),
             "starter_lap": n(self.starter_lap_tb, u"Starter lap")}
        for key, combo, values in CHOICES:
            i = getattr(self, combo).SelectedIndex
            o[key] = values[i] if 0 <= i < len(values) else values[0]
        if o["main_bar"] is None or o["tie_bar"] is None:
            raise ValueError(u"Choose the longitudinal bar and the tie bar (the model has no rebar "
                             u"bar types? Load them with Rebar > Load NZ Rebar Sizes).")
        if o["main_count"] < 3 and shape == u"round":
            raise ValueError(u"A round pile needs at least 3 bars.")
        if min(o["bars_b"], o["bars_h"]) < 2 and shape == u"rect":
            raise ValueError(u"A face needs at least its 2 corner bars.")
        for key in ("zone_mid_sp", "zone_bottom_sp", "zone_top_sp"):
            if o[key] <= 0:
                raise ValueError(u"Spacings and the pitch must be above 0.")
        if o["cover_side"] < 0:
            raise ValueError(u"The cover cannot be below 0.")
        return o

    def preview(self):
        """(options, plan) for the drawings and the calculations, with the
        size typed in the panel. Raises ValueError."""
        opts = self.options()
        _n, shape, _d = engine.layout_info(self.layout_key())
        o = engine.numbers(opts)
        o["shape"] = shape
        n = self._num
        if shape == u"round":
            o["D"] = o["b"] = o["h"] = n(self.diameter_tb, u"Diameter")
        else:
            o["b"], o["h"] = n(self.b_tb, u"Width b"), n(self.h_tb, u"Depth h")
            o["D"] = o["b"]
        o["L"] = n(self.length_tb, u"Length")
        if min(o["D"], o["b"], o["h"], o["L"]) <= 0:
            raise ValueError(u"The size and length must be above 0.")
        if o["D"] / 2.0 - o["cover_side"] - o["tie_dm"] - o["main_dm"] <= 0:
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
            for canvas in (self.section_canvas, self.elev_canvas, self.view3d_canvas):
                canvas.Children.Clear()
            for box in list(self._editors.values()) + list(self._callouts.values()):
                box.Visibility = Visibility.Collapsed
            for tb in (self.calc_bars_tb, self.calc_ratio_tb, self.calc_ties_tb, self.calc_steel_tb,
                       self.calc_conc_tb, self.calc_rate_tb, self.calc_total_steel_tb,
                       self.calc_total_conc_tb, self.calc_note_tb):
                tb.Text = u""
            return
        self.status_tb.Text = u""
        self.note_tb.Text = u""
        self.apply_btn.IsEnabled = True
        rls = self._rls(o)
        self.rl_tb.Text = (u"Top RL {0:.3f} m, toe RL {1:.3f} m ({2} {3:+g})".format(
            rls[0], rls[1], _name(self.level()), rls[2]) if rls else u"")
        t = self.column_type()
        have = engine.type_size(t) if t is not None else None
        size = (o["D"], o["D"]) if o["shape"] == u"round" else (o["b"], o["h"])
        if t is not None and not (have and abs(have[0] - size[0]) < 0.5 and abs(have[1] - size[1]) < 0.5):
            self.rl_tb.Text += u"\nApply uses a type of this size: {0}".format(
                engine.type_name(o["shape"], size))
        self._placed, self._shown = set(), set()
        self._draw_section(o, plan)
        self._draw_elevation(o, plan)
        self._draw_3d(o, plan)
        for key, box in self._editors.items():
            if key not in self._placed:
                box.Visibility = Visibility.Collapsed
        for key, group in self._callouts.items():
            if key not in self._shown:
                group.Visibility = Visibility.Collapsed
        self._calc(o, plan)

    def _rls(self, o):
        """(top RL m, toe RL m, top offset mm) from the level and the top
        offset, or None when they do not read yet. RLs as the levels show them
        (their elevation above the project base point)."""
        level = self.level()
        try:
            offset = float(self.top_offset_tb.Text)
        except ValueError:
            return None
        if level is None:
            return None
        # the level's Elevation as it shows it (from the model's elevation base)
        elev = level.get_Parameter(DB.BuiltInParameter.LEVEL_ELEV)
        top = (elev.AsDouble() if elev is not None else level.Elevation) * MM + offset
        return top / 1000.0, (top - o["L"]) / 1000.0, offset

    # --- calculations ---
    def _calc(self, o, plan):
        """The Calculations card: bars, steel ratio, weights and volumes,
        each and for all the piles."""
        n_main = len(plan["main"]["pts"])
        area = n_main * math.pi * o["main_d"] ** 2 / 4.0
        if o["shape"] == u"round":
            gross = math.pi * o["D"] ** 2 / 4.0
        else:
            gross = o["b"] * o["h"]
        main_m, tie_m, st_m = engine.bar_length(plan, o)
        kg = STEEL * (main_m * math.pi * (o["main_d"] / 1000.0) ** 2 / 4.0
                      + tie_m * math.pi * (o["tie_d"] / 1000.0) ** 2 / 4.0
                      + st_m * math.pi * (o["starter_dm"] / 1000.0) ** 2 / 4.0)
        vol = gross / 1e6 * o["L"] / 1000.0
        try:
            piles = max(int(float(self.piles_tb.Text)), 1)
        except ValueError:
            piles = 1
        bar = unicode(self.main_bar_cb.SelectedItem or u"")
        self.calc_bars_tb.Text = u"{0} x {1}".format(n_main, bar)
        self.calc_ratio_tb.Text = u"{0:,.0f} mm²  {1:.2f}%".format(area, 100.0 * area / gross)
        t = plan["ties"]
        if t["kind"] == u"spiral":
            self.calc_ties_tb.Text = u"{0:.1f} m @ {1:g}".format(tie_m, t["pitch"])
        else:
            self.calc_ties_tb.Text = u"{0} sets, {1:.1f} m".format(
                len(t["zs"]) * (len(t.get("boxes", [])) or 1), tie_m)
        self.calc_steel_tb.Text = u"{0:,.1f} kg".format(kg)
        self.calc_conc_tb.Text = u"{0:.3f} m³".format(vol)
        self.calc_rate_tb.Text = u"{0:,.0f} kg/m³".format(kg / vol if vol > 0 else 0.0)
        self.calc_total_lbl.Text = u"Steel, {0} piles".format(piles) if piles > 1 else u"Steel, all"
        self.calc_total_steel_tb.Text = u"{0:,.1f} kg".format(kg * piles)
        self.calc_total_conc_tb.Text = u"{0:.3f} m³".format(vol * piles)
        self.calc_note_tb.Text = (u"{0} grade {1}. Bar lengths approximate (bends square, hooks "
                                  u"at 12 d); Revit's schedules give the exact ones.".format(
                                      self.name_tb.Text or u"", self.grade_cb.Text or u""))

    # --- the drawings' input boxes and callouts ---
    # (key, the panel input it writes into, canvas it sits on)
    EDITORS = (("s_d", "diameter_tb", "section_edit"), ("s_b", "b_tb", "section_edit"),
               ("s_h", "h_tb", "section_edit"),
               ("e_len", "length_tb", "elev_edit"), ("e_first", "tie_first_tb", "elev_edit"),
               ("e_last", "tie_last_tb", "elev_edit"), ("e_bot", "zone_bottom_len_tb", "elev_edit"),
               ("e_top", "zone_top_len_tb", "elev_edit"), ("e_lap", "lap_tb", "elev_edit"),
               ("e_ext", "bottom_ext_tb", "elev_edit"), ("e_cover_top", "cover_top_tb", "elev_edit"),
               ("e_cover_bottom", "cover_bottom_tb", "elev_edit"),
               ("e_starter", "starter_lap_tb", "elev_edit"))

    # the callouts: (key, canvas, parts). A part is ("bar", panel list),
    # ("num", panel box), ("w", words), ("dyn", words set when drawn) or
    # ("br",) for a new line. Each list and box writes into the panel's own.
    CALLOUTS = (
        ("c_tie", "section_edit", (("bar", "tie_bar_cb"), ("dyn", u"STIRRUPS"), ("br",), ("w", u"@"),
                                   ("num", "zone_mid_sp_tb"), ("dyn", u"mm CRS"))),
        ("c_main_n", "section_edit", (("num", "main_count_tb"), ("w", u"-"), ("bar", "main_bar_cb"),
                                      ("w", u"VERTICAL BARS"))),
        ("c_main", "section_edit", (("dyn", u"12 -"), ("bar", "main_bar_cb"), ("w", u"VERTICAL BARS"))),
        ("c_cover", "section_edit", (("num", "cover_side_tb"), ("w", u"COVER"))),
        ("c_start", "section_edit", (("dyn", u"12 -"), ("bar", "starter_bar_cb"), ("w", u"STARTER BARS"))),
        ("z_top", "elev_edit", (("bar", "tie_bar_cb"), ("dyn", u"STIRRUPS"), ("br",), ("w", u"@"),
                                ("num", "zone_top_sp_tb"), ("dyn", u"mm CRS"))),
        ("z_mid", "elev_edit", (("bar", "tie_bar_cb"), ("dyn", u"STIRRUPS"), ("br",), ("w", u"@"),
                                ("num", "zone_mid_sp_tb"), ("dyn", u"mm CRS"))),
        ("z_bot", "elev_edit", (("bar", "tie_bar_cb"), ("dyn", u"STIRRUPS"), ("br",), ("w", u"@"),
                                ("num", "zone_bottom_sp_tb"), ("dyn", u"mm CRS"))),
        ("z_main", "elev_edit", (("dyn", u"8 -"), ("bar", "main_bar_cb"), ("br",),
                                 ("w", u"VERTICAL BARS"))),
        ("z_start", "elev_edit", (("dyn", u"12 -"), ("bar", "starter_bar_cb"), ("br",),
                                  ("w", u"STARTER BARS"))),
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
        each callout, made once and moved by every redraw (a box redrawn while
        it is typed in would lose the cursor)."""
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
                    combo.Width = 62 if key.startswith("z_") else 70
                    combo.Height = 22
                    combo.FontSize = 11
                    combo.ToolTip = u"The bar (the same as in the panel)"
                    combo.SelectionChanged += self._callout_bar_handler(part[1])
                    row.Children.Add(combo)
                    synced.append(("bar", combo, part[1]))
                else:
                    box = self._number_box(key + "_%d" % i, part[1], 36)
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
        """Show the panel's bars and numbers in a callout (its lists filled
        from the panel's), unless that part is being changed right now."""
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
        """The first click in a box selects all its text, so typing replaces
        it; clicks after that place the cursor as usual."""
        if not sender.IsKeyboardFocusWithin:
            sender.Focus()
            args.Handled = True

    def _editor_left(self, sender, args):
        self._update()

    def _click_off(self, sender, args):
        """A click on empty space ends the typing: the box loses its cursor
        and selection. Clicks on an input, a list or a button are left alone."""
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
        """Put a dimension's box centred on (cx, cy), turned by `angle` degrees
        to run along its dimension line, showing its panel input's value."""
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
        """A callout as on the column details: a leader from the bar (an
        arrowhead on it, `stop` px short of its middle) to (ex, ey), then a
        short level tail and the words, which hold the panel's lists and boxes."""
        group = self._callouts[key]
        self._sync_callout(key, words)
        group.Visibility = Visibility.Visible
        group.Measure(Size(1000.0, 200.0))
        width, height = group.DesiredSize.Width, group.DesiredSize.Height
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
        _arrow(canvas, ink, ax, ay, -ux, -uy, 12.0, 4.2)
        Canvas.SetLeft(group, tail + 3.0 if side > 0 else tail - 3.0 - width)
        first_row = group.Children[0]
        first_row.Measure(Size(1000.0, 200.0))
        Canvas.SetTop(group, ey - first_row.DesiredSize.Height / 2.0)
        self._shown.add(key)

    def _bar_brushes(self):
        """(fill, edge) for the bars: the theme's accent, and the accent at
        45% for the edge, so the bars follow the accent picked in About."""
        accent = self.TryFindResource("BrushPrimaryGreen")
        c = getattr(accent, "Color", None) or Color.FromRgb(0x20, 0x8A, 0x3C)
        edge = Color.FromRgb(int(c.R * 0.45), int(c.G * 0.45), int(c.B * 0.45))
        return SolidColorBrush(c), SolidColorBrush(edge)

    def _bar_line(self, canvas, pts, width=4.0, opacity=1.0):
        """A bar seen along its length: the accent with a darker edge."""
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

    def _bar_dot(self, canvas, x, y, size, hollow=False):
        """A bar seen end on: a solid dot in the accent with a darker edge
        (a starter bar: a ring)."""
        fill, edge = self._bar_brushes()
        e = _ellipse(x, y, size / 2.0)
        e.Fill = None if hollow else fill
        e.Stroke = fill if hollow else edge
        e.StrokeThickness = 2.0 if hollow else 1.2
        canvas.Children.Add(e)

    # --- Section A-A: a fixed picture, the bars as the panel sets them ---
    def _draw_section(self, o, plan):
        canvas = self.section_canvas
        canvas.Children.Clear()
        ink = self.TryFindResource("BrushTextPrimary")
        kind = plan["ties"]["kind"]
        tie_words = TIE_WORDS[kind]
        if o["shape"] == u"round":
            self._section_round(canvas, ink, o, plan, tie_words)
        else:
            self._section_rect(canvas, ink, o, plan, tie_words)

    def _section_round(self, canvas, ink, o, plan, tie_words):
        cx, cy, R = 128.0, 200.0, 112.0
        cover_px, tie_w = 20.0, 4.0
        rt = R - cover_px - tie_w / 2.0                  # the tie / spiral ring, px
        rb = rt - tie_w / 2.0 - 8.0                       # the bars' centres
        _concrete_shape(canvas, _ellipse(cx, cy, R))
        _grain(canvas, _ellipse(cx, cy, R))
        _centre_line(canvas, ink, cx - R - 22, cy, cx + R + 22, cy)
        _centre_line(canvas, ink, cx, cy - R - 22, cx, cy + R + 22)
        fill, edge = self._bar_brushes()
        rings = [rt] if plan["ties"]["kind"] == u"circle" else [rt, rt - 1.5]
        for r in rings[:1]:
            for brush, w in ((edge, tie_w + 2.2), (fill, tie_w)):
                ring = _ellipse(cx, cy, r)
                ring.Stroke = brush
                ring.StrokeThickness = w
                canvas.Children.Add(ring)
        dots = []
        for a in plan["main"]["ang"]:
            x, y = cx + rb * math.cos(a), cy - rb * math.sin(a)
            self._bar_dot(canvas, x, y, 13.0)
            dots.append((x, y))
        # the diameter, across the top
        y_d = cy - R - 30.0
        for x in (cx - R, cx + R):
            _ext_line(canvas, ink, x, cy - 4.0, x, y_d - 6.0)
        _dim_line(canvas, ink, cx - R, y_d, cx + R, y_d)
        self._place_editor("s_d", cx, y_d - 12.0)
        # callouts, in one column to the right, as on the column details
        col = cx + R + 26.0
        a = math.radians(40.0)
        self._callout("c_tie", canvas, cx + rt * math.cos(a), cy - rt * math.sin(a), col, cy - R + 10.0,
                      tie_w / 2.0 + 1.0, [tie_words[1].replace(u" @", u""), tie_words[2]])
        right = max(dots, key=lambda p: p[0] - abs(p[1] - cy) * 0.3)
        self._callout("c_main_n", canvas, right[0], right[1], col, cy + 18.0, 6.5)
        a = math.radians(-45.0)
        rc = R - cover_px / 2.0
        self._callout("c_cover", canvas, cx + rc * math.cos(a), cy - rc * math.sin(a), col,
                      cy + R - 6.0, 0.0)

    def _section_rect(self, canvas, ink, o, plan, tie_words):
        ratio = max(0.5, min(1.6, o["h"] / o["b"]))
        W = 210.0 if ratio <= 1.0 else 210.0 / ratio
        H = W * ratio
        cx, cy = 132.0, 200.0
        x0, y0, x1, y1 = cx - W / 2.0, cy - H / 2.0, cx + W / 2.0, cy + H / 2.0
        cover_px, tie_w = 18.0, 3.4
        _concrete_shape(canvas, _rect(x0, y0, x1, y1))
        _grain(canvas, _rect(x0, y0, x1, y1))
        _centre_line(canvas, ink, cx, y0 - 46.0, cx, y1 + 20.0)
        _centre_line(canvas, ink, x0 - 46.0, cy, x1 + 20.0, cy)
        fill, edge = self._bar_brushes()
        ti = cover_px + tie_w / 2.0                       # tie centreline, in from the face
        bi = ti + tie_w / 2.0 + 7.5                       # bar centres, in from the face
        xs = plan["xs"]
        ys = plan["ys"]

        def bx(i, n):
            return x0 + bi + (W - 2 * bi) * i / float(n - 1)

        def by(i, n):
            return y1 - bi - (H - 2 * bi) * i / float(n - 1)
        nb, nh = len(xs), len(ys)
        boxes = [(x0 + ti, y0 + ti, x1 - ti, y1 - ti)]
        if o.get("inner_ties"):
            wrap = 7.5 + tie_w / 2.0
            if nb >= 4:
                boxes.append((bx(1, nb) - wrap, y0 + ti, bx(nb - 2, nb) + wrap, y1 - ti))
            if nh >= 4:
                boxes.append((x0 + ti, by(nh - 2, nh) - wrap, x1 - ti, by(1, nh) + wrap))
        for bx0, by0, bx1, by1 in boxes:
            for brush, w in ((edge, tie_w + 2.2), (fill, tie_w)):
                r = _rect(bx0, by0, bx1, by1, 7.0)
                r.Stroke = brush
                r.StrokeThickness = w
                canvas.Children.Add(r)
        # the perimeter tie's 135 degree hooks, at its top left corner, into the middle
        hx, hy = boxes[0][0] + 3.0, boxes[0][1] + 3.0
        for dx, dy in ((16.0, 9.0), (9.0, 16.0)):
            self._bar_line(canvas, [(hx + dx * 0.25, hy + dy * 0.25), (hx + dx, hy + dy)], tie_w - 0.8)
        ring = set()
        for i in range(nb):
            ring.add((bx(i, nb), y1 - bi))
            ring.add((bx(i, nb), y0 + bi))
        for i in range(nh):
            ring.add((x0 + bi, by(i, nh)))
            ring.add((x1 - bi, by(i, nh)))
        for x, y in ring:
            self._bar_dot(canvas, x, y, 13.0)
        starters = plan["starters"]
        if starters:
            for x, y in ring:
                dx, dy = cx - x, cy - y
                ln = math.hypot(dx, dy) or 1.0
                self._bar_dot(canvas, x + dx / ln * 12.5, y + dy / ln * 12.5, 10.0, True)
        # b across the top, h up the left
        yd = y0 - 30.0
        for x in (x0, x1):
            _ext_line(canvas, ink, x, y0 - 4.0, x, yd - 6.0)
        _dim_line(canvas, ink, x0, yd, x1, yd)
        self._place_editor("s_b", (x0 + cx) / 2.0, yd - 12.0)
        xd = x0 - 30.0
        for y in (y0, y1):
            _ext_line(canvas, ink, x0 - 4.0, y, xd - 6.0, y)
        _dim_line(canvas, ink, xd, y0, xd, y1)
        self._place_editor("s_h", xd - 12.0, (y1 + cy) / 2.0, -90.0)
        # the callouts, one column to the right
        col = x1 + 38.0
        n = len(plan["main"]["pts"])
        self._callout("c_tie", canvas, x1 - ti, y0 + H * 0.22, col, y0 + 8.0, tie_w / 2.0 + 1.0,
                      [tie_words[1].replace(u" @", u""), tie_words[2]])
        self._callout("c_main", canvas, x1 - bi, by(min(1, nh - 1), nh), col, cy + 26.0, 6.5,
                      [u"{0} -".format(n)])
        self._callout("c_cover", canvas, x1 - cover_px / 2.0, y1 - H * 0.12, col, y1 - 6.0, 0.0)
        if starters:
            sx, sy = x0 + bi + 12.5 * 0.7071, y1 - bi - 12.5 * 0.7071
            self._callout("c_start", canvas, sx + 5.0, sy, col, y1 + 26.0, 0.0,
                          [u"{0} -".format(len(starters["pts"]))])

    # --- Elevation: a fixed picture, the zones and ends as the panel sets them ---
    def _draw_elevation(self, o, plan):
        canvas = self.elev_canvas
        canvas.Children.Clear()
        ink = self.TryFindResource("BrushTextPrimary")
        kind = plan["ties"]["kind"]
        words = TIE_WORDS[kind]
        x0, x1 = 86.0, 136.0                        # the column's faces
        y_top, y_bot = 92.0, 560.0                   # its top and bottom
        gap0, gap1 = 318.0, 338.0                    # the break, between the two pieces
        cap = o["bottom_ext"] > 0 or bool(plan["starters"])
        # the cap or footing below: dashed, wider
        if cap:
            outline = _rect(x0 - 50.0, y_bot, x1 + 50.0, y_bot + 86.0)
            outline.Stroke = ink
            outline.StrokeThickness = 1
            outline.Opacity = 0.6
            outline.StrokeDashArray = DoubleCollection([4.0, 3.0])
            canvas.Children.Add(outline)
            _label(canvas, u"CAP / FOOTING", x0 - 46.0, y_bot + 70.0, 9, opacity=0.6, ink=ink)
        # the column in two pieces, a break line across each cut
        for ya, yb in ((y_top, gap0), (gap1, y_bot)):
            _concrete_shape(canvas, _rect(x0, ya, x1, yb))
            _grain(canvas, _rect(x0, ya, x1, yb))
        for y in (gap0, gap1):
            zig = Polyline()
            zig.Points = PointCollection([Point(x0 - 8, y), Point(x0 + 22, y), Point(x0 + 27, y - 6),
                                          Point(x0 + 33, y + 6), Point(x0 + 38, y), Point(x1 + 8, y)])
            zig.Stroke = ink
            zig.StrokeThickness = 1.4
            canvas.Children.Add(zig)
        _centre_line(canvas, ink, (x0 + x1) / 2.0, y_top - 30.0, (x0 + x1) / 2.0, y_bot + 30.0)
        _label(canvas, self.name_tb.Text or u"", x0 - 4.0, 4.0, 13, True, ink=ink)
        rls = self._rls(o)
        if rls:
            _label(canvas, u"TOP RL {0:.3f}".format(rls[0]), x0 - 4.0, 22.0, 10, True, ink=ink)
            _label(canvas, u"TOE RL {0:.3f}".format(rls[1]), x0 - 4.0, y_bot + 92.0 if cap else y_bot + 64.0,
                   10, True, ink=ink)

        cover_px, tie_w = 6.0, 3.0
        xt0, xt1 = x0 + cover_px, x1 - cover_px                   # the ties
        xb0, xb1 = xt0 + tie_w + 3.5, xt1 - tie_w - 3.5            # the bars
        top = y_top - 52.0 if o["lap"] > 0 else y_top + 8.0
        bottom = y_bot + 52.0 if o["bottom_ext"] > 0 else y_bot - 12.0
        hook = plan["main"]["hook"]

        def main_bars():
            """The longitudinal bars, with their hook (drawn over the ties)."""
            for xb, side in ((xb0, -1.0), (xb1, 1.0)):
                for ya, yb in ((top, gap0), (gap1, bottom)):
                    pts = [(xb, ya), (xb, yb)]
                    if yb == bottom and hook:
                        turn = side if hook == u"out" else -side
                        pts = [(xb, ya), (xb, yb - 6.0), (xb + turn * 6.0, yb), (xb + turn * 24.0, yb)]
                    self._bar_line(canvas, pts, 4.0)
        # --- the starters, beside them ---
        starters = plan["starters"]
        if starters:
            ys = y_bot - 150.0
            for xb, side in ((xb0 + 6.5, -1.0), (xb1 - 6.5, 1.0)):
                fx = side if o.get("starter_dir") != u"in" else -side
                self._bar_line(canvas, [(xb, ys), (xb, y_bot + 46.0), (xb + fx * 6.0, y_bot + 52.0),
                                        (xb + fx * 40.0, y_bot + 52.0)], 3.0, 0.7)
        # --- the ties: dense in the end zones, wider in the middle ---
        first_y = y_bot - 14.0 if o["tie_first"] >= 0 else y_bot + 40.0
        last_y = y_top + 12.0
        zb = o["zone_bottom_len"] > 0 and kind != u"spiral"
        zt = o["zone_top_len"] > 0 and kind != u"spiral"
        bot_end = first_y - 120.0 if zb else first_y
        top_end = last_y + 100.0 if zt else last_y
        tie_ys = []

        def run(ya, yb, step):
            y = ya
            while y >= yb - 0.1:
                if not (gap0 - 2 < y < gap1 + 2):
                    tie_ys.append(y)
                y -= step
        if zb:
            run(first_y, bot_end, 8.0)
        run(bot_end - (16.0 if zb else 0.0), top_end + (16.0 if zt else 0.0), 16.0)
        if zt:
            run(top_end, last_y, 8.0)
        tie_ys.append(last_y)
        if kind == u"spiral":
            # the spiral: a zigzag up the column, front runs solid, back runs faint
            for ya, yb in ((y_top + 12.0, gap0), (gap1, first_y)):
                y = yb
                while y - 16.0 >= ya:
                    self._bar_line(canvas, [(xt0, y), (xt1, y - 8.0)], tie_w - 0.8)
                    self._bar_line(canvas, [(xt1, y - 8.0), (xt0, y - 16.0)], tie_w - 1.4, 0.35)
                    y -= 16.0
        else:
            for y in tie_ys:
                self._bar_line(canvas, [(xt0, y), (xt1, y)], tie_w - 0.6)
        main_bars()
        # --- dimensions, on the left: each box sits on its line, breaking it ---
        xl = 14.0
        for y in (y_top, y_bot):
            _ext_line(canvas, ink, x0 - 4.0, y, xl - 6.0, y)
        _dim_line(canvas, ink, xl, y_top, xl, y_bot)
        self._place_editor("e_len", xl, (y_top + gap0) / 2.0, -90.0)
        xm = 64.0
        # the first tie above the bottom, and the last below the top: small, arrows from outside
        _ext_line(canvas, ink, x0 - 4.0, first_y, xm - 6.0, first_y)
        _ext_line(canvas, ink, x0 - 4.0, last_y, xm - 6.0, last_y)
        if o["tie_first"] >= 0:
            _dim_outside(canvas, ink, xm, y_bot, xm, first_y)
            self._place_editor("e_first", xm, y_bot + 40.0, -90.0)
        else:
            _dim_line(canvas, ink, xm, y_bot, xm, first_y)
            self._place_editor("e_first", xm, y_bot + 22.0, -90.0)
        _dim_outside(canvas, ink, xm, y_top, xm, last_y)
        self._place_editor("e_last", xm, y_top - 40.0, -90.0)
        if zb:
            _ext_line(canvas, ink, x0 - 4.0, bot_end, xm - 6.0, bot_end)
            _dim_line(canvas, ink, xm, first_y, xm, bot_end)
            self._place_editor("e_bot", xm, (first_y + bot_end) / 2.0, -90.0)
        if zt:
            _ext_line(canvas, ink, x0 - 4.0, top_end, xm - 6.0, top_end)
            _dim_line(canvas, ink, xm, last_y, xm, top_end)
            self._place_editor("e_top", xm, (last_y + top_end) / 2.0, -90.0)
        # the bars' ends, on the right: lap above the top or the top cover, into the cap or the bottom cover
        xr = x1 + 16.0
        if o["lap"] > 0:
            _ext_line(canvas, ink, xb1 + 4.0, top, xr + 6.0, top)
            _dim_line(canvas, ink, xr, top, xr, y_top)
            self._place_editor("e_lap", xr, (top + y_top) / 2.0, -90.0)
        else:
            _ext_line(canvas, ink, x1 + 4.0, y_top, xr + 6.0, y_top)
            _ext_line(canvas, ink, xb1 + 4.0, top, xr + 6.0, top)
            _dim_outside(canvas, ink, xr, y_top, xr, top)
            self._place_editor("e_cover_top", xr, y_top - 40.0, -90.0)
        if o["bottom_ext"] > 0:
            _ext_line(canvas, ink, xb1 + 26.0, bottom, xr + 6.0, bottom)
            _dim_line(canvas, ink, xr, y_bot, xr, bottom)
            self._place_editor("e_ext", xr, (y_bot + bottom) / 2.0, -90.0)
        else:
            _ext_line(canvas, ink, x1 + 4.0, y_bot, xr + 6.0, y_bot)
            _dim_outside(canvas, ink, xr, y_bot, xr, bottom + 4.0)
            self._place_editor("e_cover_bottom", xr, y_bot + 40.0, -90.0)
        if starters:
            ys = y_bot - 150.0
            xs_ = 38.0
            _ext_line(canvas, ink, xb0 + 2.0, ys, xs_ - 6.0, ys)
            _dim_line(canvas, ink, xs_, ys, xs_, y_bot)
        # --- callouts, one column to the right ---
        col = x1 + 34.0
        tie_word = words[1].replace(u" @", u"")
        n = len(plan["main"]["pts"])
        self._callout("z_main", canvas, xb1, gap0 - 60.0, col, gap0 - 24.0, 3.0,
                      [u"{0} -".format(n)])
        if kind == u"spiral":
            self._callout("z_mid", canvas, xt1, (y_top + gap0) / 2.0, col, y_top + 40.0, 1.5,
                          [tie_word, words[2]])
        else:
            if zt:
                self._callout("z_top", canvas, xt1, top_end - 30.0, col, y_top + 20.0, 1.5,
                              [tie_word, words[2]])
            mid_y = [y for y in tie_ys if top_end + 10 < y < bot_end - 10 and y > gap1]
            target = mid_y[len(mid_y) // 2] if mid_y else (gap1 + bot_end) / 2.0
            self._callout("z_mid", canvas, xt1, target, col, gap1 + 30.0, 1.5, [tie_word, words[2]])
            if zb:
                self._callout("z_bot", canvas, xt1, (first_y + bot_end) / 2.0, col, y_bot - 70.0, 1.5,
                              [tie_word, words[2]])
        if starters:
            self._callout("z_start", canvas, xb1 - 6.5, y_bot - 40.0, col, y_bot + 18.0, 2.0,
                          [u"{0} -".format(len(starters["pts"]))])
        if starters:
            self._place_editor("e_starter", 38.0, y_bot - 75.0, -90.0)

    # --- 3D: the pile as typed, see-through, its bars and ties inside ---
    def _draw_3d(self, o, plan):
        canvas = self.view3d_canvas
        canvas.Children.Clear()
        cw, ch = canvas.Width, canvas.Height
        L = o["L"]
        z_lo = min(0.0, plan["main"]["z0"], plan["ties"]["zs"][0] if plan["ties"]["zs"] else 0.0)
        z_hi = max(L, plan["main"]["z1"])
        # not to scale across: the pile is drawn at least a quarter as wide as it is tall
        span = z_hi - z_lo
        height = ch - 36.0
        k = height / span
        wide = max(o["b"] * k, 0.5 * height)
        sx = wide / o["b"]                                      # px per mm across
        cx, base_y = cw / 2.0, ch - 18.0 + z_lo * k

        def P(x, y, z):
            """Axonometric: x to the right and down, y to the left and down, z up."""
            return (cx + (x - y) * 0.8660 * sx * 0.75,
                    base_y + (x + y) * 0.5 * sx * 0.75 - z * k)

        ink = self.TryFindResource("BrushTextPrimary")
        fill, edge = self._bar_brushes()
        ghost = SolidColorBrush(Color.FromArgb(0x55, 0xC4, 0xC6, 0xC8))
        rnd = o["shape"] == u"round"
        hb, hh = o["b"] / 2.0, o["h"] / 2.0

        def outline(z):
            if rnd:
                return [P(hb * math.cos(t * math.pi / 18.0), hb * math.sin(t * math.pi / 18.0), z)
                        for t in range(37)]
            return [P(-hb, -hh, z), P(hb, -hh, z), P(hb, hh, z), P(-hb, hh, z), P(-hb, -hh, z)]

        def poly(pts, stroke, w, op=1.0, fill_=None):
            if fill_ is not None:
                pg = Polygon()
                pg.Points = PointCollection([Point(x, y) for x, y in pts])
                pg.Fill = fill_
                pg.Stroke = stroke
                pg.StrokeThickness = w
                pg.Opacity = op
                canvas.Children.Add(pg)
                return
            pl = Polyline()
            pl.Points = PointCollection([Point(x, y) for x, y in pts])
            pl.Stroke = stroke
            pl.StrokeThickness = w
            pl.Opacity = op
            canvas.Children.Add(pl)

        # the body: its silhouette, see-through
        bottom, top = outline(0.0), outline(L)
        xs = [p[0] for p in bottom]
        left, right = min(xs), max(xs)
        lo = [p for p in bottom if abs(p[0] - left) < 0.5][0]
        ro = [p for p in bottom if abs(p[0] - right) < 0.5][0]
        lt = [p for p in top if abs(p[0] - left) < 0.5][0]
        rt = [p for p in top if abs(p[0] - right) < 0.5][0]
        body = [lt, lo] + [p for p in bottom if p[1] >= min(lo[1], ro[1]) - 0.1] + [ro, rt]
        poly([lt, lo, ro, rt], CONCRETE_EDGE, 0.8, 1.0, ghost)
        poly(bottom, CONCRETE_EDGE, 0.8, 0.5)
        # the ties (a sample when there are many) and the bars, then the top face over them
        zs = plan["ties"]["zs"]
        step = max(1, len(zs) // 24)
        if plan["ties"]["kind"] == u"spiral":
            z_a, z_b, pitch, _f = plan["ties"]["sets"][0]
            r = plan["ties"]["r"]
            turns = (z_b - z_a) / pitch
            pts = [P(r * math.cos(t * math.pi / 8.0), r * math.sin(t * math.pi / 8.0),
                     z_a + (z_b - z_a) * t / (16.0 * turns))
                   for t in range(int(16 * turns) + 1)]
            poly(pts, fill, 1.0, 0.9)
        else:
            for z in zs[::step]:
                if plan["ties"]["kind"] == u"circle":
                    r = plan["ties"]["r"]
                    poly([P(r * math.cos(t * math.pi / 12.0), r * math.sin(t * math.pi / 12.0), z)
                          for t in range(25)], fill, 1.0, 0.9)
                else:
                    for bx0, by0, bx1, by1 in plan["ties"]["boxes"]:
                        poly([P(bx0, by0, z), P(bx1, by0, z), P(bx1, by1, z), P(bx0, by1, z), P(bx0, by0, z)],
                             fill, 1.0, 0.9)
        m = plan["main"]
        for (x, y), a in zip(m["pts"], m["ang"]):
            pts = [P(x, y, m["z1"]), P(x, y, m["z0"])]
            if m["hook"]:
                d = -1.0 if m["hook"] == u"in" else 1.0
                pts.append(P(x + d * math.cos(a) * 150.0, y + d * math.sin(a) * 150.0, m["z0"]))
            back = (x + y) < 0
            poly(pts, edge, 2.6, 0.35 if back else 1.0)
            poly(pts, fill, 1.6, 0.45 if back else 1.0)
        if plan["starters"]:
            s = plan["starters"]
            for x, y, fx, fy in s["pts"]:
                poly([P(x, y, s["z_top"]), P(x, y, s["z_foot"]),
                      P(x + fx * s["foot"], y + fy * s["foot"], s["z_foot"])], fill, 1.2, 0.5)
        poly(top, CONCRETE_EDGE, 0.8, 1.0, SolidColorBrush(Color.FromArgb(0x88, 0xDE, 0xE1, 0xE4)))

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

    # --- a column already in the model ---
    def load_column(self, host, frame):
        """Fill the form from a column: its size, and the bars in it read
        back; Apply then replaces them."""
        found = engine.read_existing(doc, host, frame)
        s = {"layout": found.get("layout") or (u"rect_ties" if frame.shape == u"rect"
                                                 else self.layout_key())}
        if engine.layout_info(s["layout"])[1] != frame.shape:
            s["layout"] = u"rect_ties" if frame.shape == u"rect" else u"pile_spiral"
        s.update(found)
        if frame.shape == u"round":
            s["diameter"] = u"{0:g}".format(round(frame.size[0], 1))
        else:
            s["b"] = u"{0:g}".format(round(frame.size[0], 1))
            s["h"] = u"{0:g}".format(round(frame.size[1], 1))
        s["length"] = u"{0:g}".format(round(frame.length, 1))
        mark = host.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
        # the Mark only: an empty one stays empty (Apply writes the name back as the Mark)
        s["name"] = mark.AsString() if mark is not None and mark.AsString() else u""
        s["column_type"] = _name(host.Symbol)
        level, offset = engine.column_levels(doc, host)
        if level is not None:
            s["level"] = _name(level)
            s["top_offset"] = u"{0:g}".format(round(offset, 1))
        keep = self.settings()
        for key in ("piles", "grade"):
            s[key] = keep[key]
        self._apply_settings(s)
        self.update_target = host.Id
        self.apply_btn.ToolTip = (u"Rebuild column {0} where it stands: its type, levels and Mark "
                                  u"from these settings, and its bars replaced".format(engine.eid(host.Id)))
        self.header_subtitle_tb.Text = u"| Editing {0} ({1})".format(
            s["name"] or _name(host.Symbol), engine.eid(host.Id))
        self.apply_btn.Content = (u"Apply to {0}".format(s["name"])
                                  if s["name"] and len(s["name"]) < 10 else u"Apply")
        self._baseline = self.settings()


# ── PICKING AND PLACING ─────────────────────────────────────────────────────

class _ColumnFilter(ISelectionFilter):
    """Columns and foundations that can host rebar."""

    def AllowElement(self, element):
        if not isinstance(element, DB.FamilyInstance) or element.Category is None:
            return False
        cat = engine.eid(element.Category.Id)
        if cat not in (int(DB.BuiltInCategory.OST_StructuralColumns),
                       int(DB.BuiltInCategory.OST_StructuralFoundation),
                       int(DB.BuiltInCategory.OST_Columns)):
            return False
        data = RebarHostData.GetRebarHostData(element)
        return data is not None and data.IsValidHost()

    def AllowReference(self, reference, position):
        return False


def _pick_update(window):
    with forms.WarningBar(title=BAR_UPDATE):
        ref = uidoc.Selection.PickObject(ObjectType.Element, _ColumnFilter(), BAR_UPDATE)
    host = doc.GetElement(ref)
    frame = engine.from_host(doc, host)
    window.load_column(host, frame)


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
    """The Mark for each of n new piles: the name, numbered when there are several."""
    if not name:
        return [None] * n
    return [name] if n == 1 else [u"{0}-{1}".format(name, i + 1) for i in range(n)]


def _place(window):
    """Click a point for each new pile or column; each is placed (its type
    made if this size is new) and reinforced. One undo. Returns True when
    done, False to go back to the window."""
    opts = window.options()
    symbol, shape, size, level, top_offset, length, name = window.placing()
    points = []
    with forms.WarningBar(title=BAR_PLACE.format(0)) as bar:
        while True:
            try:
                points.append(_pick_xyz(BAR_PLACE.format(len(points))))
            except OperationCanceledException:
                break
            bar.message_tb.Text = BAR_PLACE.format(len(points))
    if not points:
        return False
    t = DB.Transaction(doc, u"Column Rebar - place")
    t.Start()
    try:
        made_type = engine.ensure_type(doc, symbol, shape, size)
        hosts = [engine.place_column(doc, made_type, p, level, top_offset, length, mark)
                 for p, mark in zip(points, _marks(name, len(points)))]
        doc.Regenerate()
        for host in hosts:
            engine.reinforce(doc, host, engine.from_host(doc, host), opts)
        t.Commit()
    except Exception as err:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        dlg.message(u"The piles could not be made, nothing was changed.\n\n{0}".format(err), title=TITLE)
        return False
    return True


def _update(window):
    """Rebuild the column loaded by Update where it stands: its type (made
    if this size is new), levels and Mark from the window, and its bars
    replaced. One undo. Returns True when done."""
    host = doc.GetElement(window.update_target)
    if host is None:
        window.update_target = None
        dlg.message(u"That column is no longer in the model.", title=TITLE)
        return False
    opts = window.options()
    symbol, shape, size, level, top_offset, length, name = window.placing()
    t = DB.Transaction(doc, u"Column Rebar - update")
    t.Start()
    try:
        if symbol.Family.Id == host.Symbol.Family.Id or engine.type_shape(host.Symbol) is not None:
            engine.update_column(doc, host, engine.ensure_type(doc, symbol, shape, size), level,
                                 top_offset, length, name)
        doc.Regenerate()
        engine.remove_bars(doc, host)
        engine.reinforce(doc, host, engine.from_host(doc, host), opts)
        t.Commit()
    except Exception as err:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        dlg.message(u"The column could not be updated, nothing was changed.\n\n{0}".format(err),
                    title=TITLE)
        return False
    return True


def _reinforce_existing(window):
    """Pick columns already in the model and reinforce them, each at its own
    size. One undo. Returns True when done, False to go back to the window."""
    opts = window.options()
    shape = engine.layout_info(opts["layout"])[1]
    with forms.WarningBar(title=BAR_PICK):
        refs = uidoc.Selection.PickObjects(ObjectType.Element, _ColumnFilter(), BAR_PICK)
    hosts = [doc.GetElement(r) for r in refs]
    if not hosts:
        return False
    frames = []
    for host in hosts:
        frame = engine.from_host(doc, host)
        if frame.shape != shape:
            raise ValueError(u"Column {0} is {1}, but the layout is for {2} columns. Choose a {1} "
                             u"layout.".format(engine.eid(host.Id),
                                               u"round" if frame.shape == u"round" else u"rectangular",
                                               u"round" if shape == u"round" else u"rectangular"))
        frames.append((host, frame))
    with_bars = [h for h, _f in frames if _has_bars(h)]
    if with_bars:
        if not dlg.confirm(u"{0} of the {1} already have bars in them. Replace them?".format(
                len(with_bars), len(hosts)), title=TITLE, yes=u"Replace", no=u"Cancel"):
            return False
    t = DB.Transaction(doc, u"Column Rebar")
    t.Start()
    try:
        made = {"main": 0, "ties": 0, "starters": 0}
        for host, frame in frames:
            engine.remove_bars(doc, host)
            got = engine.reinforce(doc, host, frame, opts)
            for k in made:
                made[k] += got[k]
        t.Commit()
    except Exception as err:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        dlg.message(u"The bars could not be made, nothing was changed.\n\n{0}".format(err), title=TITLE)
        return False
    return True


def main():
    window = ColumnWindow()
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
