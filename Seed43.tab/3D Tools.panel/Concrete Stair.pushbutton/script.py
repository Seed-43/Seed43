# -*- coding: utf-8 -*-
# "Concrete Stair"
# "Seed43"
# """
# Build a native Revit concrete stair (Assembled Stair, Monolithic Run) from
# sizes typed into a dialogue: levels or picked surfaces for base and top,
# width, riser, going, waist and top landing thickness. Each size set gets
# its own stair type, e.g. Concrete Stair - SR180 SG300 SD100 LD100. The top
# landing length is worked out so the soffit meets the landing underside.
# """
import io
import os
import os.path as op
import json
import math

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import FileStream, FileMode, FileAccess
from System.Windows import (Point, FontWeights, TextAlignment, Visibility, Size,
                            Thickness, HorizontalAlignment, GridLength)
from System.Windows import LogicalTreeHelper
from System.Windows.Controls import Canvas, TextBlock, TextBox, ComboBox, ComboBoxItem, Button
from System.Windows.Controls.Primitives import TextBoxBase, ButtonBase, ScrollBar
from System.Windows.Input import Keyboard, FocusManager
from System.Windows.Media import Visual, VisualTreeHelper
from System.Windows.Markup import XamlReader
from System.Windows.Media import (PointCollection, RotateTransform, SolidColorBrush,
                                  Color, DoubleCollection, PenLineJoin, LinearGradientBrush)
from System.Windows.Media.Effects import BlurEffect
from System.Windows.Shapes import Line, Polygon, Polyline, Ellipse

from pyrevit import revit, DB, forms
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from Snippets import _dialogs as dlg
from Snippets import _userdata
from Snippets._icons import set_header_icon
from Snippets.seed43_theme import apply_seed43_palette, apply_seed43_dimensions

import stair_engine as engine
import rebar_engine
import fill_family
from Autodesk.Revit.DB.Structure import RebarBarType, RebarHookType

doc = revit.doc
uidoc = revit.uidoc

# ── CONSTANTS ───────────────────────────────────────────────────────────────

SCRIPT_DIR = op.dirname(__file__)
TITLE = u"Concrete Stair"
MM = 304.8
LANDING_SHOWN = 1000.0       # a top landing longer than this is drawn this long, with a break
FLOOR_SHOWN = 400.0          # the floor slab beside the top landing, as far as the top detail shows it
# Concrete in the drawings: the same greys on a dark or a light theme
CONCRETE = SolidColorBrush(Color.FromRgb(0xC9, 0xCD, 0xD1))
CONCRETE_DARK = SolidColorBrush(Color.FromRgb(0xA4, 0xA9, 0xAF))
CONCRETE_EDGE = SolidColorBrush(Color.FromRgb(0x3A, 0x40, 0x48))
CONCRETE_TOP = SolidColorBrush(Color.FromRgb(0xDE, 0xE1, 0xE4))
CONCRETE_RISER = SolidColorBrush(Color.FromRgb(0xBB, 0xC0, 0xC5))
CONCRETE_SIDE = SolidColorBrush(Color.FromRgb(0x9C, 0xA2, 0xA8))
SETTINGS =_userdata.user_path(u"Concrete Stair", u"settings.json")
STYLES_XAML = op.join(op.dirname(op.abspath(dlg.__file__)), u"Seed43Styles.xaml")

DEFAULTS = {"rl": u"0", "width": u"1200",
            "riser": u"180", "going": u"300", "waist": u"100", "landing": u"100",
            "railings": False, "slab": False, "floor_type": u"",
            "steps": u"4", "working": u"top", "level": u"",
            # Rebar tab, defaults from the typical stair detail
            "rebar": False, "main_bar": u"HD12", "main_spacing": u"450",
            "main_hook": u"Standard - 180 deg.", "lap": u"300",
            "dist": True, "dist_bar": u"HD12", "dist_spacing": u"300",
            "mesh": False, "mesh_bar": u"R6", "mesh_spacing": u"200",
            "cover_top": u"35", "soffit_min": u"70", "corner_clear": u"50", "cover_side": u"40",
            "cover_bottom": u"50", "cover_end": u"50", "top_landing": u"", "rebar_layout": u"hooked", "footing": True,
            "footing_w": u"250", "footing_d": u"250"}

NO_HOOK = u"(no hook)"
KEEP_MATERIAL = u"(keep the types' own)"

BAR_P1 = (u"Concrete Stair: click P1, the foot of the FIRST riser on the "
          u"stair centreline. ESC goes back.")
BAR_P2 = (u"Concrete Stair: click P2, any point further UP the stair, to set "
          u"its direction. ESC goes back.")
BAR_POINT = u"Concrete Stair: click point {0} of 2 for the {1}. ESC goes back."
BAR_TL = (u"Concrete Stair: click the far edge of the top slab; the top landing runs "
          u"from the top riser to it. ESC goes back.")
TL_PICK = u"pick"                       # TL typed as this: picked after P1 and P2
BAR_SURFACE = (u"Concrete Stair: click the surface or reference plane for the "
               u"{0}. ESC goes back.")
BAR_STAIR = (u"Concrete Stair: click the stair to reinforce. ESC goes back.")
BAR_UPDATE = (u"Concrete Stair: click the stair to update. Its sizes fill the "
              u"dialogue. ESC goes back.")
BAR_LINES = (u"Concrete Stair: pick the lines that outline the landing slab, "
             u"then Finish. ESC skips the slab.")


# ── HELPERS (private to this file) ──────────────────────────────────────────

def _name(element):
    return DB.Element.Name.GetValue(element)


def _badge(canvas, text, cx, cy):
    """A round marker (number or code) centred on (cx, cy): the same marker
    sits beside the input it stands for."""
    from System.Windows.Controls import Border
    from System.Windows import CornerRadius, Thickness, HorizontalAlignment
    tb = TextBlock()
    tb.Text = text
    tb.FontSize = 10
    tb.FontWeight = FontWeights.Bold
    tb.Foreground = SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF))
    tb.HorizontalAlignment = HorizontalAlignment.Center
    b = Border()
    # the theme's accent, as the badges beside the inputs
    b.Background = canvas.TryFindResource("BrushPrimaryGreen")
    b.CornerRadius = CornerRadius(8)
    b.Height = 16
    b.MinWidth = 16
    b.Padding = Thickness(3, 0, 3, 0)
    b.Child = tb
    b.Measure(Size(200, 200))
    Canvas.SetLeft(b, cx - b.DesiredSize.Width / 2.0)
    Canvas.SetTop(b, cy - 8)
    canvas.Children.Add(b)
    return b


def _load_settings():
    data = dict(DEFAULTS)
    try:
        with io.open(SETTINGS, "r", encoding="utf-8") as f:
            data.update(json.loads(f.read()))
    except Exception:
        pass
    return data


def _save_settings(data):
    try:
        with io.open(SETTINGS, "w", encoding="utf-8") as f:
            f.write(unicode(json.dumps(data, indent=2)))
    except Exception:
        pass


def _pick_xyz(prompt):
    """A point: on the work plane in a plan, or snapped to a model face
    anywhere else (a 3D view has no work plane to click on)."""
    view = doc.ActiveView
    if isinstance(view, DB.ViewPlan):
        return uidoc.Selection.PickPoint(prompt)
    ref = uidoc.Selection.PickObject(ObjectType.PointOnElement, prompt)
    return ref.GlobalPoint


class _CurveFilter(ISelectionFilter):
    """Model and detail lines only."""

    def AllowElement(self, element):
        return isinstance(element, DB.CurveElement)

    def AllowReference(self, reference, position):
        return False


def _dim_line(canvas, ink, x1, y1, x2, y2):
    """A dimension line, CAD style: thin line, a filled arrowhead at each end."""
    ln = Line()
    ln.X1, ln.Y1, ln.X2, ln.Y2 = x1, y1, x2, y2
    ln.Stroke = ink
    ln.StrokeThickness = 1
    ln.Opacity = 0.95
    canvas.Children.Add(ln)
    length = math.hypot(x2 - x1, y2 - y1)
    if length < 2.0:
        return
    ux, uy = (x2 - x1) / length, (y2 - y1) / length
    size = min(7.0, length * 0.4)
    for px, py, dx, dy in ((x1, y1, ux, uy), (x2, y2, -ux, -uy)):
        bx, by = px + dx * size, py + dy * size
        nx, ny = -dy * size * 0.3, dx * size * 0.3
        head = Polygon()
        head.Points = PointCollection([Point(px, py), Point(bx + nx, by + ny),
                                       Point(bx - nx, by - ny)])
        head.Fill = ink
        canvas.Children.Add(head)


def _ext_line(canvas, ink, x1, y1, x2, y2):
    """A thin extension line from a feature out to its dimension line."""
    ln = Line()
    ln.X1, ln.Y1, ln.X2, ln.Y2 = x1, y1, x2, y2
    ln.Stroke = ink
    ln.StrokeThickness = 1
    ln.Opacity = 0.5
    canvas.Children.Add(ln)


def _dim_angle(x1, y1, x2, y2):
    """The angle (degrees, -90 up to 90) a dimension's box is turned by to run
    along its line, reading left to right or upwards."""
    ang = math.degrees(math.atan2(y2 - y1, x2 - x1))
    while ang >= 90.0:
        ang -= 180.0
    while ang < -90.0:
        ang += 180.0
    return ang


# ── UI / ENTRY POINT ────────────────────────────────────────────────────────

class StairWindow(forms.WPFWindow):
    """The Concrete Stair dialogue. Picks hide it; it comes back after."""

    # Class level, not just set in __init__: the XAML's own Text values fire
    # TextChanged while the window is still loading, before __init__ runs on.
    _ready = False
    _updating = False
    _syncing = False                    # a drawing box is being filled from its panel input

    # --- construction ---
    def __init__(self):
        forms.WPFWindow.__init__(self, op.join(SCRIPT_DIR, "ConcreteStair.xaml"))
        stream = FileStream(STYLES_XAML, FileMode.Open, FileAccess.Read)
        try:
            self.Resources.MergedDictionaries.Add(XamlReader.Load(stream))
        finally:
            stream.Close()
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)
        set_header_icon(self, SCRIPT_DIR)

        self.action = None              # what to do once the window is hidden
        self._ready = False
        self._updating = False          # _refresh writes the steps box itself
        self._levels = sorted(DB.FilteredElementCollector(doc).OfClass(DB.Level),
                              key=lambda l: l.ProjectElevation)
        self._floor_types = sorted(DB.FilteredElementCollector(doc).OfClass(DB.FloorType),
                                   key=_name)
        for level in self._levels:
            self.level_cb.Items.Add(_name(level))
        for ft in self._floor_types:
            self.floor_type_cb.Items.Add(_name(ft))

        self._materials = sorted(DB.FilteredElementCollector(doc).OfClass(DB.Material),
                                 key=lambda m: m.Name.lower())
        self.material_cb.Items.Add(KEEP_MATERIAL)
        for m in self._materials:
            self.material_cb.Items.Add(m.Name)
        self._baseline = None           # what Reset goes back to; None = the defaults
        self._build_editors()
        self.PreviewMouseDown += self._click_off
        # the panel's number boxes behave the same way: click, type, done
        for name in ("rl_tb", "steps_tb", "riser_tb", "going_tb", "width_tb", "waist_tb",
                     "landing_tb", "top_landing_tb", "footing_w_tb", "footing_d_tb"):
            box = getattr(self, name)
            box.GotKeyboardFocus += self._editor_focused
            box.PreviewMouseLeftButtonDown += self._select_on_click
        self._apply_settings(_load_settings())

    # --- filling the form ---
    def _apply_settings(self, s):
        """Put a settings dict (as settings() makes) into the form. Anything
        missing takes its default. Used at start and by Reset."""
        data = dict(DEFAULTS)
        data.update(s)
        s = data
        self._forced.clear()
        self._ready = False
        try:
            self.rl_tb.Text = s["rl"]
            for box, key in ((self.width_tb, "width"), (self.riser_tb, "riser"),
                             (self.going_tb, "going"), (self.waist_tb, "waist"),
                             (self.landing_tb, "landing"),
                             (self.top_landing_tb, "top_landing")):
                box.Text = s[key]
            self.steps_tb.Text = s["steps"]
            self.work_top_rb.IsChecked = s["working"] != "base"
            self.work_base_rb.IsChecked = s["working"] == "base"
            self.railings_cb.IsChecked = bool(s["railings"])
            self.slab_cb.IsChecked = bool(s["slab"])
            # the saved level if it is still there, else the active view's level
            names = [_name(level) for level in self._levels]
            current = getattr(doc.ActiveView, "GenLevel", None)
            if s["level"] in names:
                index = names.index(s["level"])
            elif current is not None and _name(current) in names:
                index = names.index(_name(current))
            else:
                index = 0
            self.level_cb.SelectedIndex = index
            names = [_name(ft) for ft in self._floor_types]
            if s["floor_type"] in names:
                self.floor_type_cb.SelectedIndex = names.index(s["floor_type"])
            elif names:
                match = [i for i, n in enumerate(names) if n.startswith(s["landing"])]
                self.floor_type_cb.SelectedIndex = match[0] if match else 0
            self._select_material_name(s.get("material", u""))
            self._load_rebar(s)
        finally:
            self._ready = True
        self._update()

    def reset_clicked(self, sender, args):
        """Back to what the selected stair has now; with no stair selected,
        back to the defaults."""
        self._apply_settings(self._baseline if self._baseline is not None else DEFAULTS)

    # --- tabs ---
    def _show_tab(self, name):
        stair = name == "stair"
        self.page_stair.Visibility = Visibility.Visible if stair else Visibility.Collapsed
        self.page_rebar.Visibility = Visibility.Collapsed if stair else Visibility.Visible
        self.tab_stair_btn.Tag = u"Active" if stair else u""
        self.tab_rebar_btn.Tag = u"" if stair else u"Active"
        # the big card on the right follows the tab; plan, 3D and numbers stay
        self.section_card.Visibility = Visibility.Visible if stair else Visibility.Collapsed
        self.rebar_card.Visibility = Visibility.Collapsed if stair else Visibility.Visible
        # the row of three cards below: plan, 3D and numbers for the stair; the
        # footing and the top, zoomed and bigger, for the reinforcement
        for card in (self.plan_card, self.view3d_card, self.calc_card):
            card.Visibility = Visibility.Visible if stair else Visibility.Collapsed
        for card in (self.footing_card, self.top_card):
            card.Visibility = Visibility.Collapsed if stair else Visibility.Visible
        self.right_grid.RowDefinitions[2].Height = GridLength(244.0 if stair else 340.0)
        self.bottom_grid.ColumnDefinitions[3].Width = GridLength(10.0 if stair else 0.0)
        self.bottom_grid.ColumnDefinitions[4].Width = GridLength(248.0 if stair else 0.0)

    def tab_stair_clicked(self, sender, args):
        self._show_tab("stair")

    def tab_rebar_clicked(self, sender, args):
        self._show_tab("rebar")

    # --- rebar tab ---
    def _load_rebar(self, s):
        self._bar_types = sorted(DB.FilteredElementCollector(doc).OfClass(RebarBarType),
                                 key=lambda b: (b.BarNominalDiameter, _name(b)))
        self._hook_types = sorted(DB.FilteredElementCollector(doc).OfClass(RebarHookType),
                                  key=_name)
        framing = DB.FilteredElementCollector(doc).OfCategory(
            DB.BuiltInCategory.OST_StructuralFraming).WhereElementIsElementType()
        bar_names = [_name(b) for b in self._bar_types]

        def fill(combo, names, wanted):
            combo.Items.Clear()
            for n in names:
                combo.Items.Add(n)
            if wanted in names:
                combo.SelectedIndex = names.index(wanted)
            elif names:
                combo.SelectedIndex = 0

        self._layouts = rebar_engine.layout_names()
        names = [n for _, n in self._layouts]
        keys = [k for k, _ in self._layouts]
        fill(self.rebar_layout_cb, names,
             names[keys.index(s["rebar_layout"])] if s["rebar_layout"] in keys else names[0])
        fill(self.main_bar_cb, bar_names, s["main_bar"])
        fill(self.dist_bar_cb, bar_names, s["dist_bar"])
        fill(self.mesh_bar_cb, bar_names, s["mesh_bar"])
        fill(self.main_hook_cb, [NO_HOOK] + [_name(h) for h in self._hook_types], s["main_hook"])
        for box, key in ((self.main_spacing_tb, "main_spacing"), (self.lap_tb, "lap"),
                         (self.dist_spacing_tb, "dist_spacing"),
                         (self.mesh_spacing_tb, "mesh_spacing"),
                         (self.cover_top_tb, "cover_top"), (self.cover_soffit_tb, "soffit_min"),
                         (self.corner_clear_tb, "corner_clear"),
                         (self.cover_side_tb, "cover_side"), (self.cover_bottom_tb, "cover_bottom"),
                         (self.cover_end_tb, "cover_end")):
            box.Text = s[key]
        self.rebar_cb.IsChecked = bool(s["rebar"])
        self.dist_cb.IsChecked = bool(s["dist"])
        self.mesh_cb.IsChecked = bool(s["mesh"])
        self.footing_cb.IsChecked = bool(s["footing"])
        self.footing_w_tb.Text = s["footing_w"]
        self.footing_d_tb.Text = s["footing_d"]
        if not self._bar_types:
            self.rebar_cb.IsChecked = False
            self.rebar_cb.IsEnabled = False
            self.rebar_cb.Content = u"This model has no rebar bar types to use"

    def footing_options(self):
        """(width, depth) of the Stair Footing placed with the stair, or None
        when Footing is off. Raises ValueError."""
        if not self.footing_cb.IsChecked:
            return None
        size = (self._num(self.footing_w_tb, u"Footing width (FW)"),
                self._num(self.footing_d_tb, u"Footing depth (FD)"))
        if min(size) <= 0:
            raise ValueError(u"Footing width and depth must be above 0.")
        return size

    def rebar_options(self, always=False):
        """The Rebar tab as engine options; None when it is off, unless
        `always` (the drawing, and reinforcing an existing stair)."""
        if not always and not self.rebar_cb.IsChecked:
            return None

        def pick(combo, items):
            i = combo.SelectedIndex
            return items[i] if 0 <= i < len(items) else None

        hook_i = self.main_hook_cb.SelectedIndex
        i = self.rebar_layout_cb.SelectedIndex
        opts = {
            "layout": self._layouts[i][0] if 0 <= i < len(self._layouts) else None,
            "main_bar": pick(self.main_bar_cb, self._bar_types),
            "main_spacing": self._num(self.main_spacing_tb, u"Main bar spacing"),
            "main_hook": self._hook_types[hook_i - 1] if hook_i > 0 else None,
            "lap": self._num(self.lap_tb, u"Leg into the landing"),
            "cover_top": self._num(self.cover_top_tb, u"Top cover"),
            "cover_soffit": self._num(self.cover_soffit_tb, u"Soffit cover"),
            "corner_clear": self._num(self.corner_clear_tb, u"Clear below the step corners"),
            "cover_side": self._num(self.cover_side_tb, u"Side cover"),
            "cover_bottom": self._num(self.cover_bottom_tb, u"Footing bottom cover"),
            "cover_end": self._num(self.cover_end_tb, u"Footing end cover"),
            "dist_bar": pick(self.dist_bar_cb, self._bar_types) if self.dist_cb.IsChecked else None,
            "dist_spacing": self._num(self.dist_spacing_tb, u"Bars across spacing"),
            "mesh_bar": pick(self.mesh_bar_cb, self._bar_types) if self.mesh_cb.IsChecked else None,
            "mesh_spacing": self._num(self.mesh_spacing_tb, u"Mesh spacing"),
        }
        # the footing is set on the Stair tab; the bars hook into it
        footing = self.footing_options()
        opts["footing"], opts["footing_size"] = (True, footing) if footing else (None, None)
        if opts["main_bar"] is None:
            raise ValueError(u"Choose the main bar.")
        for key in ("main_spacing", "dist_spacing", "mesh_spacing"):
            if opts[key] <= 0:
                raise ValueError(u"Spacings must be above 0.")
        return opts

    def rebar_settings(self):
        return {
            "rebar": bool(self.rebar_cb.IsChecked),
            "main_bar": unicode(self.main_bar_cb.SelectedItem or u""),
            "main_spacing": self.main_spacing_tb.Text,
            "main_hook": unicode(self.main_hook_cb.SelectedItem or u""),
            "lap": self.lap_tb.Text,
            "dist": bool(self.dist_cb.IsChecked),
            "dist_bar": unicode(self.dist_bar_cb.SelectedItem or u""),
            "dist_spacing": self.dist_spacing_tb.Text,
            "mesh": bool(self.mesh_cb.IsChecked),
            "mesh_bar": unicode(self.mesh_bar_cb.SelectedItem or u""),
            "mesh_spacing": self.mesh_spacing_tb.Text,
            "cover_top": self.cover_top_tb.Text, "soffit_min": self.cover_soffit_tb.Text,
            "corner_clear": self.corner_clear_tb.Text,
            "cover_side": self.cover_side_tb.Text, "cover_bottom": self.cover_bottom_tb.Text,
            "cover_end": self.cover_end_tb.Text,
            "rebar_layout": (self._layouts[self.rebar_layout_cb.SelectedIndex][0]
                             if self.rebar_layout_cb.SelectedIndex >= 0 else u"hooked"),
            "footing": bool(self.footing_cb.IsChecked),
            "footing_w": self.footing_w_tb.Text, "footing_d": self.footing_d_tb.Text,
        }

    # --- reading the form ---
    def _num(self, box, label):
        try:
            return float((box.Text or u"").strip())
        except ValueError:
            raise ValueError(u"{0} is not a number.".format(label))

    def working(self):
        """'top' or 'base': the end that is set; the other is worked out."""
        return u"base" if self.work_base_rb.IsChecked else u"top"

    def values(self):
        """Everything in the form, checked. Raises ValueError with a reason.

        Only the working end is read from the form. The other end is
        steps x riser above or below it, on the same level.
        """
        if self.level_cb.SelectedIndex < 0:
            raise ValueError(u"Choose the working level.")
        level = self._levels[self.level_cb.SelectedIndex]
        v = {
            "width": self._num(self.width_tb, u"Width"),
            "riser": self._num(self.riser_tb, u"Riser"),
            "going": self._num(self.going_tb, u"Going"),
            "waist": self._num(self.waist_tb, u"Waist"),
            "landing": self._num(self.landing_tb, u"Landing thickness"),
            "top_level": level, "base_level": level,
        }
        if v["width"] <= 0:
            raise ValueError(u"Width must be above 0.")
        # bars need a minimum slope thickness: raise SD to suit
        need = self._min_waist()
        if need is not None and v["waist"] < need:
            v["waist_entered"] = v["waist"]
            v["waist"] = need
        try:
            steps = int(float((self.steps_tb.Text or u"").strip()))
        except ValueError:
            raise ValueError(u"Steps is not a whole number.")
        if steps < 1:
            raise ValueError(u"There must be at least one step.")
        rise = steps * v["riser"]
        if self.working() == u"top":
            v["top_offset"] = self._num(self.rl_tb, u"RL top")
            v["base_offset"] = v["top_offset"] - rise
        else:
            v["base_offset"] = self._num(self.rl_tb, u"RL base")
            v["top_offset"] = v["base_offset"] + rise
        text = (self.top_landing_tb.Text or u"").strip()
        v["pick_top_landing"] = text.lower() == TL_PICK
        try:
            top_landing = float(text) if text and not v["pick_top_landing"] else None
        except ValueError:
            raise ValueError(u"Top landing (TL) is not a number. Leave it empty for the "
                             u"least, or use its pick button.")
        v["sizes"] = engine.plan(rise, v["riser"], v["going"], v["waist"],
                                 v["landing"], steps, top_landing)
        v["top_landing_short"] = (top_landing is not None
                                  and top_landing < v["sizes"]["xyz_min"] - 0.05)
        v["type_riser"] = v["riser"]
        v["material"] = self.material()
        return v

    def _min_waist(self):
        """The slope (SD) the bars need, when Rebar is on and readable."""
        if not self.rebar_cb.IsChecked:
            return None
        try:
            return rebar_engine.min_waist(rebar_engine.numbers(self.rebar_options()))
        except ValueError:
            return None

    # --- material ---
    def material(self):
        """The chosen material's id, or None to keep the types' own."""
        i = self.material_cb.SelectedIndex
        return self._materials[i - 1].Id if i > 0 else None

    def _select_material_name(self, name):
        names = [m.Name for m in self._materials]
        self.material_cb.SelectedIndex = names.index(name) + 1 if name in names else 0

    def select_material(self, material_id):
        ids = [m.Id for m in self._materials]
        self.material_cb.SelectedIndex = ids.index(material_id) + 1 if material_id in ids else 0

    def _update(self):
        if not self._ready or self._updating:
            return
        self._updating = True
        try:
            self._refresh()
        finally:
            self._updating = False

    def _refresh(self):
        top = self.working() == u"top"
        self.rl_lbl.Text = u"RL top" if top else u"RL base"
        try:
            v = self.values()
            s = v["sizes"]
        except ValueError as err:
            for box in (self.calc_rise_tb, self.calc_run_tb, self.calc_slope_tb,
                        self.calc_landing_tb, self.calc_length_tb, self.calc_stair_vol_tb,
                        self.calc_foot_vol_tb, self.calc_total_vol_tb):
                box.Text = u""
            self.type_tb.Text = u""
            self.status_tb.Text = unicode(err)
            self.place_btn.IsEnabled = False
            self.drawing.Children.Clear()
            self.plan_canvas.Children.Clear()
            self.view3d_canvas.Children.Clear()
            return
        self._last = v
        self.rise_tb.Text = u"= {0} x {1:g} = {2:g} rise".format(
            s["risers"], v["riser"], round(s["rise"], 1))
        self.other_end_tb.Text = (
            u"{0} works out at RL {1:g} on level {2} ({3} steps x {4:g})".format(
                u"Base" if top else u"Top",
                round(v["base_offset"] if top else v["top_offset"], 1),
                _name(v["top_level"]), s["risers"], v["riser"]))
        self.status_tb.Text = u""
        self.place_btn.IsEnabled = True
        self._draw(v)
        self._draw_plan(v)
        self._draw_3d(v)
        try:
            self._draw_rebar(v["sizes"], self.rebar_options(True))
            self.rebar_note_tb.Text = u""
        except ValueError as err:
            self.rebar_drawing.Children.Clear()
            self.footing_canvas.Children.Clear()
            self.top_canvas.Children.Clear()
            for box in self._rb_editors.values():
                box.Visibility = Visibility.Collapsed
            self.rebar_note_tb.Text = unicode(err)
        self._calc(v)
        self.type_tb.Text = u"Type: " + engine.type_name(v["type_riser"], v["going"],
                                                          v["waist"], v["landing"])
        if "waist_entered" in v:
            self.type_tb.Text += (u"\nSlope SD raised from {0:g} to {1:g} to fit the bars "
                                  u"(see Reinforcement > Cover)".format(v["waist_entered"], v["waist"]))
        if v.get("pick_top_landing"):
            self.type_tb.Text += u"\nTL is picked after P1 and P2."
        elif v.get("top_landing_short"):
            self.type_tb.Text += (u"\nTL {0:g} was too short: the least, {1:g}, is used."
                                  .format(float(self.top_landing_tb.Text), s["xyz"]))

    def _calc(self, v):
        """The Calculations card: sizes in metres, concrete in cubic metres."""
        s = v["sizes"]
        fw, fd = self._footing_drawn()
        fill = None
        if fd:
            shape = fill_family.fill_shape(s, fw, fd)
            if shape:
                fill = (shape[0], shape[1])
        vol = engine.volume(s, v["width"], (fw, fd) if fd else None, fill)

        def m(mm):
            return u"{0:.3f} m".format(mm / 1000.0)

        self.calc_rise_tb.Text = m(s["rise"])
        self.calc_run_tb.Text = m(s["run_length"])
        self.calc_slope_tb.Text = u"{0:.2f}°".format(s["angle"])
        self.calc_landing_tb.Text = m(s["xyz"])
        self.calc_length_tb.Text = m(s["run_length"] + s["xyz"])
        self.calc_stair_vol_tb.Text = u"{0:.3f} m³".format(vol["stair"])
        self.calc_foot_vol_tb.Text = (u"{0:.3f} m³".format(vol["footing"])
                                      if fd else u"none")
        self.calc_total_vol_tb.Text = u"{0:.3f} m³".format(vol["total"])

    def inputs_changed(self, sender, args):
        self._update()

    # --- step counter ---
    def _bump_steps(self, delta):
        try:
            n = int(float(self.steps_tb.Text))
        except ValueError:
            n = 1
        self.steps_tb.Text = unicode(max(1, n + delta))
        self._update()

    def steps_up_clicked(self, sender, args):
        self._bump_steps(1)

    def steps_down_clicked(self, sender, args):
        self._bump_steps(-1)

    def working_changed(self, sender, args):
        """Switching ends keeps the stair where it is: the RL box takes the
        value the new working end already had."""
        last = getattr(self, "_last", None)
        if last is not None:
            value = last["top_offset"] if self.working() == u"top" else last["base_offset"]
            self.rl_tb.Text = u"{0:g}".format(round(value, 1))
        self._update()

    # --- the section drawing ---
    def _footing_drawn(self):
        """(width, depth) of the footing placed with the stair, or (0, 0)
        when there is none or the sizes do not read as numbers yet."""
        if not self.footing_cb.IsChecked:
            return 0.0, 0.0
        try:
            w, d = float(self.footing_w_tb.Text), float(self.footing_d_tb.Text)
        except ValueError:
            return 0.0, 0.0
        return (w, d) if w > 0 and d > 0 else (0.0, 0.0)

    def _solid_outline(self, s, fw, fd):
        """The stair's section as one outline, the footing (fw x fd, none when
        fd is 0) and the fill wedge beside it included: one cast shape, no seam
        where the footing meets the first step. (x, z) points in mm."""
        pts = engine.profile(s)
        if not fd:
            return pts
        shape = fill_family.fill_shape(s, fw, fd)
        if shape and not shape[2]:
            # the fill wedge between the soffit and the footing's top is
            # part of the shape: the soffit runs on down to the footing's face
            return [(0.0, -fd)] + pts[:-1] + [(fw, shape[1]), (fw, -fd)]
        if shape:
            # the soffit's toe is past the footing: the wedge under it fills in
            return [(0.0, -fd)] + pts + [(fw + shape[0], 0.0), (fw, -shape[1]), (fw, -fd)]
        return [(0.0, -fd)] + pts + [(fw, 0.0), (fw, -fd)]

    # --- the section's input boxes: the dimensions are typed where they are drawn ---
    # (key used when drawing, the input in the left panel it stands for)
    EDITORS = (("riser", "riser_tb"), ("going", "going_tb"), ("waist", "waist_tb"),
               ("landing", "landing_tb"), ("top_landing", "top_landing_tb"),
               ("footing_w", "footing_w_tb"), ("footing_d", "footing_d_tb"))
    # the same on the Reinforcement tab's two details: (key, panel input, canvas it sits on)
    REBAR_EDITORS = (("cover_end", "cover_end_tb", "footing_edit"),
                     ("cover_bottom", "cover_bottom_tb", "footing_edit"),
                     ("cover_soffit", "cover_soffit_tb", "footing_edit"),
                     ("corner_clear", "corner_clear_tb", "footing_edit"),
                     ("cover_top", "cover_top_tb", "top_edit"),
                     ("lap", "lap_tb", "top_edit"))

    def _build_editors(self):
        """One small input box for each editable dimension, made once and
        moved by every redraw (a box redrawn while it is typed in would lose
        the cursor). Each writes into its panel input, which does the rest."""
        self._editors = {}
        self._rb_editors = {}
        self._pickers = {}
        self._forced = set()            # boxes whose value was put back to the least, shown red
        self._sources = {}              # key -> the panel input a box writes into
        style = self.TryFindResource("TextBoxStyle")
        wanted = ([(key, source, "drawing_edit", self._editors) for key, source in self.EDITORS]
                  + [(key, source, canvas, self._rb_editors)
                     for key, source, canvas in self.REBAR_EDITORS])
        for key, source, canvas_name, store in wanted:
            box = TextBox()
            box.Style = style
            box.Tag = key
            box.Width = 38                  # short: a dimension is at most four digits
            box.Height = 20
            box.FontSize = 12
            box.Padding = Thickness(0)
            box.HorizontalContentAlignment = HorizontalAlignment.Center
            box.ToolTip = getattr(self, source).ToolTip
            box.Visibility = Visibility.Collapsed
            box.TextChanged += self._editor_handler(key, source)
            box.GotKeyboardFocus += self._editor_focused
            box.LostKeyboardFocus += self._editor_left
            box.PreviewMouseLeftButtonDown += self._select_on_click
            getattr(self, canvas_name).Children.Add(box)
            store[key] = box
            self._sources[key] = source
        # the pick target above the boxes that have one, the same picks as the panel's
        for key, handler, tip in (
                ("riser", self.pick_riser_clicked,
                 u"Pick 2 points: the height between them becomes the riser"),
                ("going", self.pick_going_clicked,
                 u"Pick 2 points: the distance between them in plan becomes the going"),
                ("top_landing", self.pick_top_landing_clicked,
                 u"Click the far edge of the top slab to set the landing length")):
            btn = Button()
            btn.Style = self.TryFindResource("SmallIconButtonStyle")
            btn.Content = u"⌖"
            btn.FontSize = 16
            btn.Width = btn.Height = 22
            btn.ToolTip = tip
            btn.Visibility = Visibility.Collapsed
            btn.Click += handler
            self.drawing_edit.Children.Add(btn)
            self._pickers[key] = btn

    def _editor_handler(self, key, source):
        def changed(sender, args):
            if not self._syncing:
                self._forced.discard(key)       # typed in again: no longer a forced value
                getattr(self, source).Text = sender.Text
        return changed

    def _least(self, key):
        """The smallest value a box may have, or None. The top landing is never
        shorter than where the soffit meets its underside; the slope thickness
        is never thinner than the bars need (Reinforcement on)."""
        if key == "top_landing":
            last = getattr(self, "_last", None)
            return last["sizes"]["xyz_min"] if last else None
        if key == "waist":
            return self._min_waist()
        return None

    def _enforce_least(self, key):
        """A value below its least goes back to the least and is shown red."""
        least = self._least(key)
        if least is None:
            return
        source = getattr(self, dict(self.EDITORS)[key])
        text = (source.Text or u"").strip()
        if not text or text.lower() == TL_PICK:         # empty top landing: the least, by itself
            return
        try:
            value = float(text)
        except ValueError:
            return
        if value < least - 0.05:
            self._forced.add(key)
            source.Text = u"{0:g}".format(round(least, 1))

    def _editor_focused(self, sender, args):
        sender.SelectAll()

    def _select_on_click(self, sender, args):
        """The first click in a box selects all its text, so typing replaces
        it; clicks after that place the cursor as usual."""
        if not sender.IsKeyboardFocusWithin:
            sender.Focus()                  # GotKeyboardFocus selects everything
            args.Handled = True             # the click must not move the cursor after

    def _editor_left(self, sender, args):
        """Leaving a box: a value below its least becomes the least, in red;
        the box then goes back to what the panel says (an empty landing
        length shows its least again)."""
        key = sender.Tag
        if key is not None:
            self._enforce_least(key)
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
        # Keyboard.ClearFocus alone, then Window.Focus, would hand the focus straight back
        # to the box it was in; the window itself becomes the focused element instead
        Keyboard.ClearFocus()
        FocusManager.SetFocusedElement(self, self)
        Keyboard.Focus(self)

    def _place_editor(self, key, cx, cy, angle=0.0, editors=None):
        """Put a dimension's box centred on (cx, cy), turned by `angle` degrees
        to run along its dimension line, and show the panel input's value in
        it, unless it is the box being typed in. `editors` is the set of boxes
        it belongs to (the section's, unless it is one of the details')."""
        box = (editors or self._editors)[key]
        Canvas.SetLeft(box, cx - box.Width / 2.0)
        Canvas.SetTop(box, cy - box.Height / 2.0)
        box.RenderTransformOrigin = Point(0.5, 0.5)
        box.RenderTransform = RotateTransform(angle)
        box.Visibility = Visibility.Visible
        box.Foreground = (self.TryFindResource("BrushDanger") if key in self._forced
                          else self.TryFindResource("BrushTextInput"))
        picker = self._pickers.get(key)
        if picker is not None:
            # the target sits just above the box (above its end when it is turned up)
            lift = (box.Width if abs(angle) > 45 else box.Height) / 2.0 + 14.0
            Canvas.SetLeft(picker, cx - picker.Width / 2.0)
            Canvas.SetTop(picker, cy - lift - picker.Height / 2.0)
            picker.Visibility = Visibility.Visible
        source = self._sources[key]
        text = getattr(self, source).Text
        if key == "top_landing" and not (text or u"").strip():
            # empty means the least: show what that works out to
            text = u"{0:g}".format(round(self._last["sizes"]["xyz"], 1))
        if not box.IsKeyboardFocused and box.Text != text:
            self._syncing = True
            try:
                box.Text = text
            finally:
                self._syncing = False

    def _draw(self, v):
        """Redraw the section. Four steps at most are drawn: with more, the
        first two and the last two, a break between them (the step numbers on
        them are the real ones). A top landing over 1 m is drawn 1 m long with
        a break in it. So the scale depends on the sizes, not on how many steps
        or how long a landing there is. Every dimension is a box to type in,
        turned to run along its dimension line. Canvas units are pixels, y down."""
        s = v["sizes"]
        top_working = self.working() == u"top"
        N, r, g = s["risers"], s["riser"], s["going"]
        xyz = s["xyz"]
        broken = N > 4
        Nd = 4 if broken else N                      # steps drawn
        xyz_d = min(xyz, max(LANDING_SHOWN, s["xyz_min"]))   # landing drawn
        land_broken = xyz > xyz_d + 0.5
        sd = dict(s)                                 # the stair as drawn
        sd["xyz"] = xyz_d
        if broken:
            sd.update(risers=4, treads=3, rise=4 * r, run_length=3 * g)
        run_len = sd["run_length"]
        x_end = run_len + xyz_d
        rise_d = sd["rise"]
        fw, fd = self._footing_drawn()
        canvas = self.drawing
        cw, ch = canvas.Width, canvas.Height
        ox = 170.0                                   # room for the base RL and the SR / FD boxes
        top_pad = 100.0                              # the going / TL boxes, their targets and the RL labels
        below = 44.0 if fd else 28.0                 # under the footing, for its FW box
        gap = 0.9                                    # the stair's break, in steps along the slope
        sx, sz = (gap * g, gap * r) if broken else (0.0, 0.0)   # the upper part moves by this
        gx = 0.4 * g if land_broken else 0.0         # the landing's break, along the run
        # the scale fits four steps, a landing up to 1 m and the breaks
        scale = min((cw - ox - 125.0) / max(3 * g + xyz_d + gap * g + gx, 1.0),
                    (ch - top_pad - below) / max(4 * r + gap * r + fd, 1.0))
        oy = ch - below - fd * scale        # the stair sits on the bottom of the drawing
        placed = set()                      # the input boxes this drawing uses

        def X(x):
            return ox + x * scale

        def Y(z):
            return oy - z * scale

        def XU(x):                          # the upper part, past the stair's break
            return X(x + sx)

        def XR(x):                          # past the landing's break as well
            return X(x + sx + gx)

        def YU(z):
            return Y(z + sz)

        canvas.Children.Clear()
        ink = self.TryFindResource("BrushTextPrimary")
        fill = CONCRETE
        accent = self.TryFindResource("BrushPrimaryGreen")

        def line(x1, y1, x2, y2, opacity=0.7, thick=1.0):
            ln = Line()
            ln.X1, ln.Y1, ln.X2, ln.Y2 = x1, y1, x2, y2
            ln.Stroke = ink
            ln.StrokeThickness = thick
            ln.Opacity = opacity
            canvas.Children.Add(ln)

        def dim(x1, y1, x2, y2):
            """A dimension line, CAD style: thin line, a filled arrowhead at each end."""
            line(x1, y1, x2, y2, 0.95)
            length = math.hypot(x2 - x1, y2 - y1)
            if length < 2.0:
                return
            ux, uy = (x2 - x1) / length, (y2 - y1) / length
            size = min(8.0, length * 0.4)
            for px, py, sx_, sy_ in ((x1, y1, ux, uy), (x2, y2, -ux, -uy)):
                bx, by = px + sx_ * size, py + sy_ * size
                nx, ny = -sy_ * size * 0.3, sx_ * size * 0.3
                head = Polygon()
                head.Points = PointCollection([Point(px, py), Point(bx + nx, by + ny),
                                               Point(bx - nx, by - ny)])
                head.Fill = ink
                canvas.Children.Add(head)

        def datum(x1, y1, x2, y2):
            """A level line: green and dashed, like an RL on a drawing."""
            ln = Line()
            ln.X1, ln.Y1, ln.X2, ln.Y2 = x1, y1, x2, y2
            ln.Stroke = accent
            ln.StrokeThickness = 1
            ln.StrokeDashArray = DoubleCollection([5.0, 3.0])
            canvas.Children.Add(ln)

        def label(text, left, top, colour=None, bold=False, opacity=1.0):
            tb = TextBlock()
            tb.Text = text
            tb.FontSize = 12
            tb.Foreground = colour or ink
            tb.Opacity = opacity
            if bold:
                tb.FontWeight = FontWeights.SemiBold
            Canvas.SetLeft(tb, left)
            Canvas.SetTop(tb, top)
            canvas.Children.Add(tb)
            return tb

        def edit(key, cx, cy, angle):
            self._place_editor(key, cx, cy, angle)
            placed.add(key)

        def dimbox(key, x1, y1, x2, y2, off, min_x=None, max_x=None):
            """The box for a dimension, turned to run along its line and just
            off it by `off`, so it never covers the line."""
            ang = math.degrees(math.atan2(y2 - y1, x2 - x1))
            while ang >= 90.0:
                ang -= 180.0
            while ang < -90.0:
                ang += 180.0
            cx, cy = (x1 + x2) / 2.0 + off[0], (y1 + y2) / 2.0 + off[1]
            if min_x is not None:
                cx = max(cx, min_x)
            if max_x is not None:
                cx = min(cx, max_x)
            edit(key, cx, cy, ang)

        def num(x):
            return u"{0:g}".format(round(x, 1))

        def rl_marker(x, z, text, working):
            tri = Polygon()
            tri.Points = PointCollection([Point(x - 7, Y(z) - 11), Point(x + 7, Y(z) - 11),
                                          Point(x, Y(z) - 1)])
            tri.Fill = accent
            canvas.Children.Add(tri)
            # white text, the working RL bold and the worked-out one lighter
            label(u"RL " + text, x + 10, Y(z) - 24, ink, working, 1.0 if working else 0.7)

        # outline: up the steps, along the landing, back down the soffit.
        # The stair and its footing are one cast shape: one outline, one fill,
        # no seam where the footing meets the first step.
        pts = self._solid_outline(sd, fw, fd)

        # The outline is cut where a break goes, and the part past each break
        # moves on. Past four steps: square to the slope, between the second
        # and third step. A landing over 1 m: slanted, through the landing.
        pieces = [(pts, 0.0, 0.0)]
        cuts = []                                    # (side function, its gradient) for the break lines
        if broken:
            length = math.hypot(g, r)
            tx, tz = g / length, r / length
            c = tx * 1.5 * g + tz * 2.5 * r

            def side(p):
                return tx * p[0] + tz * p[1] - c

            cuts.append((side, (tx, tz)))
            pieces = [(engine.clip_half_plane(pts, side, False), 0.0, 0.0),
                      (engine.clip_half_plane(pts, side, True), sx, sz)]
        if land_broken:
            xc, zm = run_len + 0.5 * xyz_d, rise_d - s["landing"] / 2.0

            def lside(p):
                return (p[0] - xc) + 0.5 * (p[1] - zm)

            cuts.append((lside, (1.0, 0.5)))
            cut_pieces = []
            for points, dx_, dz_ in pieces:
                if (any(lside(p) > 1e-6 for p in points)
                        and any(lside(p) < -1e-6 for p in points)):
                    cut_pieces.append((engine.clip_half_plane(points, lside, False), dx_, dz_))
                    cut_pieces.append((engine.clip_half_plane(points, lside, True), dx_ + gx, dz_))
                else:
                    cut_pieces.append((points, dx_, dz_))
            pieces = cut_pieces
        for points, dx_, dz_ in pieces:
            if len(points) < 3:
                continue
            poly = Polygon()
            poly.Points = PointCollection([Point(X(x + dx_), Y(z + dz_)) for x, z in points])
            poly.Fill = fill
            poly.Stroke = CONCRETE_EDGE
            poly.StrokeThickness = 1.5
            poly.StrokeLineJoin = PenLineJoin.Round
            canvas.Children.Add(poly)
            for side_f, (ga_, gb_) in cuts:
                # the break line, a little past both faces of the cut
                on = [p for p in points if abs(side_f(p)) < 1e-6]
                if len(on) >= 2:
                    on.sort(key=lambda p: -gb_ * p[0] + ga_ * p[1])
                    (ax, az), (bx, bz) = on[0], on[-1]
                    pax, pay = X(ax + dx_), Y(az + dz_)
                    pbx, pby = X(bx + dx_), Y(bz + dz_)
                    seg = math.hypot(pbx - pax, pby - pay) or 1.0
                    ex_, ey_ = (pbx - pax) / seg * 7.0, (pby - pay) / seg * 7.0
                    line(pax - ex_, pay - ey_, pbx + ex_, pby + ey_, 1.0, 1.6)

        # the footing's FW / FD boxes
        if fd:
            # FW under the footing
            wy = Y(-fd) + 14
            dim(X(0), wy, X(fw), wy)
            line(X(0), Y(-fd) + 4, X(0), wy + 5, 0.4)
            line(X(fw), Y(-fd) + 4, X(fw), wy + 5, 0.4)
            dimbox("footing_w", X(0), wy, X(fw), wy, (0.0, 16.0))
            # FD beside the footing's outer face
            dx = X(0) - 14
            dim(dx, Y(0), dx, Y(-fd))
            line(dx - 5, Y(-fd), X(0) - 4, Y(-fd), 0.4)
            dimbox("footing_d", dx, Y(0), dx, Y(-fd), (-17.0, 0.0))

        # step numbers, the real ones: 1 2 ... N-1 N
        for d in range(1, Nd + 1):
            k = d if (d <= 2 or not broken) else N - 4 + d
            up = broken and d > 2
            label(unicode(k), (XU if up else X)((d - 1) * g) + 4,
                  (YU if up else Y)(d * r) + 2, CONCRETE_EDGE)

        # riser: on the first riser, down the left side, in the column of the
        # footing's FD dimension and just above it (every riser is the same)
        rx = X(0) - 14
        dim(rx, Y(0), rx, Y(r))
        line(rx - 5, Y(0), X(0) - 4, Y(0), 0.4)
        line(rx - 5, Y(r), X(0) - 4, Y(r), 0.4)
        dimbox("riser", rx, Y(0), rx, Y(r), (-17.0, 0.0))

        # going: on the top row beside the landing's dimension, over the last
        # tread (one step has no tread, so no going to show)
        if Nd >= 2:
            ga, gb = XU((Nd - 2) * g), XU((Nd - 1) * g)
            gy = YU(rise_d) - 40
            dim(ga, gy, gb, gy)
            line(ga, gy - 5, ga, YU((Nd - 1) * r) - 4, 0.4)
            line(gb, gy - 5, gb, YU(rise_d) - 4, 0.4)
            # above its line, kept clear of the landing's box
            dimbox("going", ga, gy, gb, gy, (0.0, -16.0), max_x=gb - 22.0)

        # slope / waist: from the inside corner at the foot of the second
        # riser, square down to the soffit - where it is measured
        # (one step has no inside corner to measure from, so no SD to show)
        slope = math.atan2(r, g)
        if Nd >= 2:
            nx, nz = math.sin(slope), -math.cos(slope)
            cx, cz = g, r
            ex, ez = cx + nx * s["waist"], cz + nz * s["waist"]
            dim(X(cx), Y(cz), X(ex), Y(ez))
            # the box just beyond the soffit, in the clear under the stair
            beyond = math.hypot(X(ex) - X(cx), Y(ez) - Y(cz)) / 2.0 + 28.0
            dimbox("waist", X(cx), Y(cz), X(ex), Y(ez),
                   (math.sin(slope) * beyond, math.cos(slope) * beyond))

        # slab: at the far end
        lx = XR(x_end) + 10
        dim(lx, YU(rise_d), lx, YU(rise_d - s["landing"]))
        line(XR(x_end), YU(rise_d - s["landing"]), lx + 5, YU(rise_d - s["landing"]), 0.4)
        dimbox("landing", lx, YU(rise_d), lx, YU(rise_d - s["landing"]), (17.0, 14.0))

        # xyz: over the top landing (the real length, broken or not)
        xy = YU(rise_d) - 40
        dim(XU(run_len), xy, XR(x_end), xy)
        line(XU(run_len), xy - 5, XU(run_len), YU(rise_d) - 4, 0.4)
        line(XR(x_end), xy - 5, XR(x_end), YU(rise_d) - 4, 0.4)
        dimbox("top_landing", XU(run_len), xy, XR(x_end), xy, (0.0, -16.0),
               min_x=XU(run_len) + 22.0)

        # RLs: top just past the end of the landing, base on the ground
        datum(XR(x_end) + 4, YU(rise_d), XR(x_end) + 110, YU(rise_d))
        rl_marker(XR(x_end) + 24, rise_d + sz, num(v["top_offset"]), top_working)
        # the base RL on the left, out past the boxes
        datum(X(0) - 160, Y(0), X(0), Y(0))
        rl_marker(X(0) - 146, 0, num(v["base_offset"]), not top_working)

        # the boxes (and targets) for dimensions this drawing does not have
        for key in self._editors:
            if key not in placed:
                self._editors[key].Visibility = Visibility.Collapsed
                if key in self._pickers:
                    self._pickers[key].Visibility = Visibility.Collapsed

    def _draw_plan(self, v):
        """The stair from above, to scale, drawn like the section: four steps
        at most (the first two and the last two with a break between), and a
        landing over 1 m drawn 1 m long with a break, so adding steps or
        lengthening the landing does not change the scale."""
        s = v["sizes"]
        canvas = self.plan_canvas
        canvas.Children.Clear()
        cw, ch = canvas.Width, canvas.Height
        N, g = s["risers"], s["going"]
        xyz = s["xyz"]
        real_end = s["run_length"] + xyz
        broken = N > 4
        Nd = 4 if broken else N
        xyz_d = min(xyz, max(LANDING_SHOWN, s["xyz_min"]))
        run_len = (Nd - 1) * g
        x_end = run_len + xyz_d
        gap = 0.9
        sx = gap * g if broken else 0.0                   # the stair's break moves the upper part by this
        gx = 0.4 * g if xyz > xyz_d + 0.5 else 0.0        # the landing's break, along the run
        width = v["width"]
        left, right, top, bottom = 38.0, 12.0, 36.0, 32.0
        scale = min((cw - left - right) / max(3 * g + xyz_d + gap * g + gx, 1.0),
                    (ch - top - bottom) / max(width, 1.0))
        ink = self.TryFindResource("BrushTextPrimary")

        # the cuts (where, in the drawn run, and how far what is past it moves)
        cuts = []
        if broken:
            cuts.append((1.5 * g, sx))
        if gx:
            cuts.append((run_len + 0.5 * xyz_d, gx))
        edges = [0.0] + [c for c, _ in cuts] + [x_end]
        shifts = [0.0]
        for _, moved in cuts:
            shifts.append(shifts[-1] + moved)

        def XS(x):                                         # a drawn position, past any breaks before it
            return left + (x + sum(m for c, m in cuts if c < x)) * scale

        def X(x):
            return left + x * scale

        def Y(y):
            return top + y * scale

        def line(x1, y1, x2, y2, stroke=None, thick=1.0, opacity=1.0):
            ln = Line()
            ln.X1, ln.Y1, ln.X2, ln.Y2 = x1, y1, x2, y2
            ln.Stroke = stroke or ink
            ln.StrokeThickness = thick
            ln.Opacity = opacity
            canvas.Children.Add(ln)

        def label(text, cx, cy, angle=0.0):
            tb = TextBlock()
            tb.Text = text
            tb.FontSize = 11
            tb.Foreground = ink
            tb.Width = 70
            tb.Height = 15
            tb.TextAlignment = TextAlignment.Center
            tb.RenderTransformOrigin = Point(0.5, 0.5)
            tb.RenderTransform = RotateTransform(angle)
            Canvas.SetLeft(tb, cx - 35)
            Canvas.SetTop(tb, cy - 7.5)
            canvas.Children.Add(tb)

        def dim(x1, y1, x2, y2):
            """A dimension line, CAD style: thin line, a filled arrowhead at each end."""
            line(x1, y1, x2, y2, opacity=0.95)
            length = math.hypot(x2 - x1, y2 - y1)
            if length < 2.0:
                return
            ux, uy = (x2 - x1) / length, (y2 - y1) / length
            size = min(7.0, length * 0.4)
            for px, py, sx_, sy_ in ((x1, y1, ux, uy), (x2, y2, -ux, -uy)):
                bx, by = px + sx_ * size, py + sy_ * size
                nx, ny = -sy_ * size * 0.3, sx_ * size * 0.3
                head = Polygon()
                head.Points = PointCollection([Point(px, py), Point(bx + nx, by + ny),
                                               Point(bx - nx, by - ny)])
                head.Fill = ink
                canvas.Children.Add(head)

        def num(x):
            return u"{0:g}".format(round(x, 1))

        def zigzag(bx):
            """A break line down the cut at screen x `bx`: straight, a zigzag
            in the middle, straight again. Top to bottom."""
            y0, y1 = Y(0), Y(width)
            ym, h, a = (y0 + y1) / 2.0, min(26.0, (y1 - y0) / 3.0), 7.0
            # two teeth, like the usual break line on a drawing
            return [(bx, y0), (bx, ym - h), (bx + a, ym - h / 2.0), (bx - a, ym),
                    (bx + a, ym + h / 2.0), (bx, ym + h), (bx, y1)]

        # the stair in one piece, or pieces with a zigzag break between them
        last = len(edges) - 2
        for i in range(last + 1):
            a, b, moved = edges[i], edges[i + 1], shifts[i]
            left_edge = zigzag(X(a + moved)) if i > 0 else [(X(a), Y(0)), (X(a), Y(width))]
            right_edge = (zigzag(X(b + moved)) if i < last
                          else [(X(b + moved), Y(0)), (X(b + moved), Y(width))])
            body = Polygon()
            body.Points = PointCollection(
                [Point(px, py) for px, py in [left_edge[0]] + right_edge + left_edge[::-1]])
            body.Fill = CONCRETE
            body.Stroke = CONCRETE_EDGE
            body.StrokeThickness = 1.5
            canvas.Children.Add(body)
            # one line for each riser; the last is where the top landing starts
            for k in range(1, Nd):
                if a < k * g < b:
                    line(X(k * g + moved), Y(0), X(k * g + moved), Y(width), CONCRETE_EDGE,
                         1.0 if k < Nd - 1 else 1.6)
            dash = Line()
            dash.X1, dash.Y1 = X(a + moved), Y(width / 2.0)
            dash.X2, dash.Y2 = X(b + moved), Y(width / 2.0)
            dash.Stroke = CONCRETE_EDGE
            dash.StrokeThickness = 0.8
            dash.StrokeDashArray = DoubleCollection([6.0, 3.0])
            canvas.Children.Add(dash)
        for i, (c, _) in enumerate(cuts):                  # the break line on each cut face
            for moved in (shifts[i], shifts[i + 1]):
                zig = Polyline()
                zig.Points = PointCollection([Point(px, py) for px, py in zigzag(X(c + moved))])
                zig.Stroke = ink
                zig.StrokeThickness = 1.4
                zig.StrokeLineJoin = PenLineJoin.Miter
                canvas.Children.Add(zig)

        # width down the left, on extension lines from the stair's edges
        wx = left - 14
        dim(wx, Y(0), wx, Y(width))
        line(wx - 5, Y(0), X(0) - 3, Y(0), opacity=0.5)
        line(wx - 5, Y(width), X(0) - 3, Y(width), opacity=0.5)
        label(num(width), wx - 10, Y(width / 2.0), -90)

        # overall length along the bottom (the real one, breaks or not)
        end_x = X(x_end + shifts[-1])
        ly = Y(width) + 14
        dim(X(0), ly, end_x, ly)
        line(X(0), Y(width) + 3, X(0), ly + 5, opacity=0.5)
        line(end_x, Y(width) + 3, end_x, ly + 5, opacity=0.5)
        label(num(real_end), (X(0) + end_x) / 2.0, ly + 11)

        # going and top landing along the top (one step has no tread)
        ty = Y(0) - 14
        if N >= 2:
            dim(X(0), ty, XS(g), ty)
            line(X(0), ty - 5, X(0), Y(0) - 2, opacity=0.5)
            line(XS(g), ty - 5, XS(g), Y(0) - 2, opacity=0.5)
            label(num(g), (X(0) + XS(g)) / 2.0, ty - 9)
        start_x = XS(run_len)
        dim(start_x, ty, end_x, ty)
        line(start_x, ty - 5, start_x, Y(0) - 2, opacity=0.5)
        line(end_x, ty - 5, end_x, Y(0) - 2, opacity=0.5)
        label(num(xyz), (start_x + end_x) / 2.0, ty - 9)

    def _draw_3d(self, v):
        """The stair as a solid, seen from the low end and one side. Plain
        axonometric: faces drawn as polygons, flat greys for top, riser and side."""
        s = v["sizes"]
        canvas = self.view3d_canvas
        canvas.Children.Clear()
        cw, ch = canvas.Width, canvas.Height
        N, r, g = s["risers"], s["riser"], s["going"]
        x_end = s["run_length"] + s["xyz"]
        W = v["width"]
        fw, fd = self._footing_drawn()
        profile = engine.profile(s)

        # camera at (-1, -1, +1): x runs up and to the right, y up and to the left
        k1, k2, k3 = 1.0 / math.sqrt(2.0), 1.0 / math.sqrt(6.0), 2.0 / math.sqrt(6.0)

        def project(x, y, z):
            return (x - y) * k1, -((x + y) * k2 + z * k3)

        corners = [project(x, y, z) for x, z in profile for y in (0.0, W)]
        if fd:
            corners += [project(x, y, z) for x in (0.0, fw) for y in (0.0, W) for z in (0.0, -fd)]
        xs = [c[0] for c in corners]
        ys = [c[1] for c in corners]
        pad = 12.0
        scale = min((cw - 2 * pad) / max(max(xs) - min(xs), 1.0),
                    (ch - 2 * pad) / max(max(ys) - min(ys), 1.0))
        ox = (cw - (max(xs) - min(xs)) * scale) / 2.0 - min(xs) * scale
        oy = (ch - (max(ys) - min(ys)) * scale) / 2.0 - min(ys) * scale

        def shade(top_rgb, bottom_rgb, angle=90.0):
            """A concrete grey fading from one shade to another (angle 90: down)."""
            return LinearGradientBrush(Color.FromRgb(*top_rgb), Color.FromRgb(*bottom_rgb), angle)

        # light from above and the left: treads brightest, risers in between,
        # the side face darkening towards the ground
        tread = shade((0xEC, 0xEE, 0xF0), (0xCF, 0xD3, 0xD7), 0.0)
        riser = shade((0xC9, 0xCD, 0xD1), (0xA9, 0xAE, 0xB3))
        side = shade((0xB6, 0xBB, 0xC0), (0x7E, 0x84, 0x8B))
        edge = SolidColorBrush(Color.FromRgb(0x6C, 0x72, 0x79))

        def pt(p):
            px, py = project(*p)
            return Point(ox + px * scale, oy + py * scale)

        def face(points3d, brush):
            poly = Polygon()
            poly.Points = PointCollection([pt(p) for p in points3d])
            poly.Fill = brush
            poly.Stroke = edge
            poly.StrokeThickness = 0.6
            poly.Opacity = 1.0
            poly.StrokeLineJoin = PenLineJoin.Round
            canvas.Children.Add(poly)

        # a soft shadow on the ground, first so everything sits on it
        ground = -fd if fd else 0.0
        reach = 0.35 * s["rise"]
        shadow = Polygon()
        shadow.Points = PointCollection([pt(p) for p in (
            (0, 0, ground), (x_end, 0, ground), (x_end + reach, -reach, ground),
            (x_end + reach, W - reach, ground), (x_end, W, ground), (0, W, ground))])
        shadow.Fill = SolidColorBrush(Color.FromArgb(0x70, 0x00, 0x00, 0x00))
        shadow.Effect = BlurEffect()
        shadow.Effect.Radius = 8.0
        canvas.Children.Add(shadow)

        # the faces never overlap on screen from here, so the order is free
        if fd:
            face([(0, 0, 0), (0, W, 0), (0, W, -fd), (0, 0, -fd)], riser)
            face([(0, 0, 0), (fw, 0, 0), (fw, 0, -fd), (0, 0, -fd)], side)
        face([(x, 0.0, z) for x, z in profile], side)
        for k in range(1, N + 1):
            x0 = (k - 1) * g
            x1 = k * g if k < N else x_end
            z0, z1 = (k - 1) * r, k * r
            face([(x0, 0, z0), (x0, W, z0), (x0, W, z1), (x0, 0, z1)], riser)
            face([(x0, 0, z1), (x1, 0, z1), (x1, W, z1), (x0, W, z1)], tread)

    def settings(self):
        s = {"rl": self.rl_tb.Text,
             "steps": self.steps_tb.Text, "working": self.working(),
             "level": unicode(self.level_cb.SelectedItem or u""),
             "width": self.width_tb.Text, "riser": self.riser_tb.Text,
             "going": self.going_tb.Text, "waist": self.waist_tb.Text,
             "landing": self.landing_tb.Text,
             "top_landing": self.top_landing_tb.Text,
             "railings": bool(self.railings_cb.IsChecked),
             "slab": bool(self.slab_cb.IsChecked),
             "floor_type": unicode(self.floor_type_cb.SelectedItem or u""),
             "material": (unicode(self.material_cb.SelectedItem)
                          if self.material_cb.SelectedIndex > 0 else u"")}
        s.update(self.rebar_settings())
        return s

    def floor_type(self):
        i = self.floor_type_cb.SelectedIndex
        return self._floor_types[i] if i >= 0 else None

    # --- buttons: each pick hides the window and is done by main() ---
    def _hide_for(self, action):
        self.action = action
        self.Hide()

    def rebar_existing_clicked(self, sender, args):
        self._hide_for("rebar_existing")

    # --- updating a stair already in the model ---
    update_target = None                # (stairs id, foot XYZ, direction XYZ)

    def update_pick_clicked(self, sender, args):
        self._hide_for("pick_update")

    def load_stair(self, stairs, frame, level):
        """Fill the form from a stair, and switch Place to Update."""
        s = frame.s
        base_z = stairs.BaseElevation
        base_rl = (base_z - level.ProjectElevation) * MM
        top_rl = base_rl + s["rise"]
        names = [_name(l) for l in self._levels]
        self._ready = False
        try:
            if _name(level) in names:
                self.level_cb.SelectedIndex = names.index(_name(level))
            self.rl_tb.Text = u"{0:g}".format(round(top_rl if self.working() == u"top" else base_rl, 1))
            self.steps_tb.Text = unicode(s["risers"])
            for box, key in ((self.riser_tb, "riser"), (self.going_tb, "going"),
                             (self.waist_tb, "waist"), (self.landing_tb, "landing")):
                box.Text = u"{0:g}".format(round(s[key], 1))
            self.width_tb.Text = u"{0:g}".format(round(frame.width, 1))
            least = engine.plan(s["rise"], s["riser"], s["going"], s["waist"],
                                s["landing"], s["risers"])["xyz_min"]
            self.top_landing_tb.Text = (u"{0:g}".format(round(s["xyz"], 1))
                                        if s["xyz"] > least + 0.5 else u"")
            self.rebar_cb.IsChecked = rebar_engine.has_rebar(doc, stairs)
            self.select_material(engine.stair_material(doc, stairs))
        finally:
            self._ready = True
        self.update_target = (stairs.Id, frame.foot, frame.d)
        self.place_btn.ToolTip = (u"Rebuild stair {0} in the same place with these settings; "
                                  u"its bars and footing are replaced too".format(
                                      rebar_engine.eid(stairs.Id)))
        self.header_subtitle_tb.Text = u"| Editing stair {0}".format(rebar_engine.eid(stairs.Id))
        self._update()
        # Reset comes back to what this stair has now
        self._baseline = self.settings()

    # --- the rebar drawing ---
    def _draw_rebar(self, s, opts):
        """The stair in section with its bars, as reinforce() will place
        them: main bars as a line, bars across as dots, the footing and the
        landing slab the top leg runs into, each set labelled."""
        plan = rebar_engine.layout(s, rebar_engine.numbers(opts))
        N, r, g = s["risers"], s["riser"], s["going"]
        x_end = s["run_length"] + s["xyz"]
        foot = min(s["waist"] / math.cos(math.atan2(r, g)) * g / r, x_end)
        fw, fd = (plan["footing"][1], plan["footing"][3] - plan["footing"][2]) \
            if plan["footing"] else (0.0, 0.0)
        x_max = max(x_end, plan["main"][0][0]) + 60.0
        z_min = -fd if plan["footing"] else 0.0
        scale = min(620.0 / max(x_max, 1.0), 220.0 / max(s["rise"] - z_min, 1.0))
        ox = 40.0
        oy = 48.0 + s["rise"] * scale        # room above for the mesh callout

        def X(x):
            return ox + x * scale

        def Y(z):
            return oy - z * scale

        canvas = self.rebar_drawing
        canvas.Children.Clear()
        ink = self.TryFindResource("BrushTextPrimary")
        fill = CONCRETE
        accent = self.TryFindResource("BrushPrimaryGreen")
        bar = SolidColorBrush(Color.FromRgb(0x2E, 0x8B, 0xFF))   # Revit's rebar blue

        def poly(points, closed, stroke, thick, fill_brush=None, dash=False, opacity=1.0):
            shape = Polygon() if closed else Polyline()
            shape.Points = PointCollection([Point(X(x), Y(z)) for x, z in points])
            shape.Stroke = stroke
            shape.StrokeThickness = thick
            shape.Opacity = opacity
            if fill_brush is not None:
                shape.Fill = fill_brush
            if dash:
                shape.StrokeDashArray = DoubleCollection([4.0, 3.0])   # floats: ints fail the cast
            canvas.Children.Add(shape)

        def text(t, x, z, dx=0, dy=0):
            tb = TextBlock()
            tb.Text = t
            tb.FontSize = 12
            tb.FontWeight = FontWeights.SemiBold
            tb.Foreground = accent
            Canvas.SetLeft(tb, X(x) + dx)
            Canvas.SetTop(tb, Y(z) + dy)
            canvas.Children.Add(tb)

        # landing slab beyond the stair, which the top leg laps into
        poly([(x_end, s["rise"]), (x_max, s["rise"]), (x_max, s["rise"] - s["landing"]),
              (x_end, s["rise"] - s["landing"])], False, ink, 1, dash=True, opacity=0.5)
        # the stair and its footing, one solid shape like the main section's
        poly(self._solid_outline(s, fw, fd), True, CONCRETE_EDGE, 1.5, fill)

        # main bars, with the hook drawn as a U turning uphill
        poly(self._bar_path(plan), False, bar, 2.5)

        # bars across: dots on the main bars, and one in each bend and the hook
        dot = 6.0
        for x, z in plan["across"] + plan["bends"]:
            e = Ellipse()
            e.Width = e.Height = dot
            e.Fill = bar
            Canvas.SetLeft(e, X(x) - dot / 2.0)
            Canvas.SetTop(e, Y(z) - dot / 2.0)
            canvas.Children.Add(e)

        # top landing mesh: top bars, and the bottom bars the main bars sit beside
        if plan["mesh"]:
            xa, xb, z1 = plan["mesh"]
            poly([(xa, z1), (xb, z1)], False, bar, 1.5)
            md = opts["mesh_bar"].BarNominalDiameter * 304.8
            poly([(xa, z1 - md), (xb, z1 - md)], False, bar, 1.5)

        def name(bar_type):
            return DB.Element.Name.GetValue(bar_type)

        def callout(number, words, x, z, dx, dy):
            """A numbered label off to one side, with a leader to the bar."""
            ax, ay = X(x), Y(z)
            ln = Line()
            ln.X1, ln.Y1, ln.X2, ln.Y2 = ax, ay, ax + dx, ay + dy
            ln.Stroke = accent
            ln.StrokeThickness = 1
            canvas.Children.Add(ln)
            _badge(canvas, number, ax + dx + 9, ay + dy)
            tb = TextBlock()
            tb.Text = words
            tb.FontSize = 12
            tb.FontWeight = FontWeights.SemiBold
            tb.Foreground = accent
            Canvas.SetLeft(tb, ax + dx + 20)
            Canvas.SetTop(tb, ay + dy - 9)
            canvas.Children.Add(tb)

        # 6 main bars, 7 bars across: leaders out below the soffit
        (xa_, za_), (xb_, zb_) = plan["main"][1], plan["main"][2]
        f = 0.3
        callout(u"6", u"{0} @ {1:g}".format(name(opts["main_bar"]), opts["main_spacing"]),
                xb_ + (xa_ - xb_) * f, zb_ + (za_ - zb_) * f, 70, 70)
        if plan["across"]:
            x7, z7 = plan["across"][min(len(plan["across"]) - 1, (2 * len(plan["across"])) // 3)]
            callout(u"7", u"{0} @ {1:g}".format(name(opts["dist_bar"]), opts["dist_spacing"]),
                    x7, z7, 70, 70)
        # 8 mesh: above the landing
        if plan["mesh"]:
            callout(u"8", u"{0} @ {1:g} both ways".format(name(opts["mesh_bar"]),
                                                          opts["mesh_spacing"]),
                    (xa + xb) / 2.0, z1, -40, -34)

        self._draw_details(plan, s, opts)

    def _bar_path(self, plan):
        """The main bar as a polyline (section mm), the hook drawn as a U
        turning uphill: the leg runs down to the hook's centre height, then
        round the U."""
        main = list(plan["main"])
        if plan["hook"]:
            hx, hz = plan["hook_centre"]
            rad = plan["hook_radius"]
            main[-1] = (hx - rad, hz)
            main += [(hx - rad * math.cos(a), hz - rad * math.sin(a))
                     for a in [math.pi * i / 8.0 for i in range(1, 9)]]
            main.append((hx + rad, hz + plan.get("hook_length", rad * 1.5)))
        return main

    def _draw_details(self, plan, s, opts):
        """The two zoomed details on the Reinforcement tab: the bottom two
        steps with the footing, and the top two steps with the landing and a
        piece of the floor slab it meets (dashed, with the floor mesh running
        through it), each cut square to the slope with a break line, so they
        show the setouts large and never rescale with the number of steps.
        Their dimensions are input boxes, like the stair section's."""
        o = rebar_engine.numbers(opts)
        N, r, g = s["risers"], s["riser"], s["going"]
        length = math.hypot(g, r)
        tx, tz = g / length, r / length                  # up the slope
        sn, cs = tz, tx                                  # sin and cos of the slope's angle
        t = r / g
        dm, dd = o["main_dm"], o["dist_dm"]
        fw, fd = ((plan["footing"][1], plan["footing"][3] - plan["footing"][2])
                  if plan["footing"] else (0.0, 0.0))
        outline = self._solid_outline(s, fw, fd)
        x_end = s["run_length"] + s["xyz"]
        slab_end = x_end + FLOOR_SHOWN                   # the floor slab's piece, past the landing
        rise, thick = s["rise"], s["landing"]
        bars = [(self._bar_path(plan), 2.5, False)]
        if plan["mesh"]:
            xa, xb, z1 = plan["mesh"]
            md = opts["mesh_bar"].BarNominalDiameter * MM
            bars += [([(xa, z1), (xb, z1)], 1.5, False), ([(xa, z1 - md), (xb, z1 - md)], 1.5, False)]
            # the floor's mesh carries on through the slab piece (dashed: not ours)
            bars += [([(xb, z1), (slab_end - 40.0, z1)], 1.5, True),
                     ([(xb, z1 - md), (slab_end - 40.0, z1 - md)], 1.5, True)]
        dots = plan["across"] + plan["bends"]
        ink = self.TryFindResource("BrushTextPrimary")
        blue = SolidColorBrush(Color.FromRgb(0x2E, 0x8B, 0xFF))    # Revit's rebar blue

        # where each detail is cut: square to the slope, 3.35 steps up from the
        # base for the footing and 3.35 down from the top for the top
        def sloped(c):
            return lambda p: tx * p[0] + tz * p[1] - c

        many = N > 2
        cut_b = sloped(2.35 * length)
        bottom = [([(cut_b, False)] if many else [], 0.0)]
        cuts_b = [cut_b] if many else []
        cut_t = sloped((N - 2.35) * length)
        clips_t = [(cut_t, True)] if many else []
        cuts_t = [cut_t] if many else []
        top = [(clips_t, 0.0)]
        xyz_d = min(s["xyz"], max(LANDING_SHOWN, s["xyz_min"]))
        if s["xyz"] > xyz_d + 0.5:
            # a landing over 1 m: its middle is left out, with a break at each end
            xc, xc2 = s["run_length"] + 0.5 * xyz_d, x_end - 0.5 * xyz_d
            zm = s["rise"] - s["landing"] / 2.0

            def ls1(p):
                return (p[0] - xc) + 0.5 * (p[1] - zm)

            def ls2(p):
                return (p[0] - xc2) + 0.5 * (p[1] - zm)

            top = [(clips_t + [(ls1, False)], 0.0),
                   (clips_t + [(ls2, True)], 0.4 * g - (xc2 - xc))]
            cuts_t += [ls1, ls2]

        def render(canvas, regions, cuts, right_top, margins, slab=None):
            """Draw the concrete, the bars and the break lines of one detail
            at a scale that fits it, and return the world -> screen functions.
            `slab` is the outline of an existing slab beside it, drawn dashed."""
            canvas.Children.Clear()
            cw, ch = canvas.Width, canvas.Height
            polys = []
            for clips, dx in regions:
                pts = outline
                for f, keep in clips:
                    pts = engine.clip_half_plane(pts, f, keep)
                if len(pts) >= 3:
                    polys.append((pts, dx))
            if not polys:
                return None
            slab_dx = regions[-1][1]
            xs = [p[0] + dx for pts, dx in polys for p in pts]
            zs = [p[1] for pts, dx in polys for p in pts]
            if slab:
                xs += [p[0] + slab_dx for p in slab]
                zs += [p[1] for p in slab]
            mL, mR, mT, mB = margins
            scale = min((cw - mL - mR) / max(max(xs) - min(xs), 1.0),
                        (ch - mT - mB) / max(max(zs) - min(zs), 1.0), 0.9)
            if right_top:       # the top detail hangs from the top right corner
                ax, ay = cw - mR - max(xs) * scale, mT + max(zs) * scale
            else:               # the footing detail stands on the bottom left
                ax, ay = mL - min(xs) * scale, ch - mB + min(zs) * scale

            def X(x):
                return ax + x * scale

            def Y(z):
                return ay - z * scale

            for pts, dx in polys:
                poly = Polygon()
                poly.Points = PointCollection([Point(X(x + dx), Y(z)) for x, z in pts])
                poly.Fill = CONCRETE
                poly.Stroke = CONCRETE_EDGE
                poly.StrokeThickness = 1.5
                poly.StrokeLineJoin = PenLineJoin.Round
                canvas.Children.Add(poly)
                for f in cuts:
                    # the break line along each cut face, a little past both ends
                    on = [p for p in pts if abs(f(p)) < 1e-6]
                    if len(on) >= 2:
                        a, b = max(((p, q) for p in on for q in on),
                                   key=lambda pq: (pq[0][0] - pq[1][0]) ** 2
                                   + (pq[0][1] - pq[1][1]) ** 2)
                        pax, pay, pbx, pby = X(a[0] + dx), Y(a[1]), X(b[0] + dx), Y(b[1])
                        seg = math.hypot(pbx - pax, pby - pay) or 1.0
                        ex_, ey_ = (pbx - pax) / seg * 7.0, (pby - pay) / seg * 7.0
                        brk = Line()
                        brk.X1, brk.Y1, brk.X2, brk.Y2 = pax - ex_, pay - ey_, pbx + ex_, pby + ey_
                        brk.Stroke = ink
                        brk.StrokeThickness = 1.6
                        canvas.Children.Add(brk)
            if slab:
                # the existing slab beside the stair: a dashed outline, its far end broken
                outline_ = Polygon()
                outline_.Points = PointCollection([Point(X(x + slab_dx), Y(z)) for x, z in slab])
                outline_.Stroke = ink
                outline_.StrokeThickness = 1
                outline_.Opacity = 0.7
                outline_.StrokeDashArray = DoubleCollection([4.0, 3.0])
                canvas.Children.Add(outline_)
                (ax_, az_), (bx_, bz_) = slab[2], slab[1]          # the slanted end, bottom to top
                pax, pay, pbx, pby = X(ax_ + slab_dx), Y(az_), X(bx_ + slab_dx), Y(bz_)
                seg = math.hypot(pbx - pax, pby - pay) or 1.0
                ex_, ey_ = (pbx - pax) / seg * 7.0, (pby - pay) / seg * 7.0
                brk = Line()
                brk.X1, brk.Y1, brk.X2, brk.Y2 = pax - ex_, pay - ey_, pbx + ex_, pby + ey_
                brk.Stroke = ink
                brk.StrokeThickness = 1.6
                canvas.Children.Add(brk)
            for points, thick, dashed in bars:
                for clips, dx in regions:
                    chains = [points]
                    for f, keep in clips:
                        chains = [c for chain in chains
                                  for c in engine.clip_polyline(chain, f, keep)]
                    for chain in chains:
                        line = Polyline()
                        line.Points = PointCollection([Point(X(x + dx), Y(z)) for x, z in chain])
                        line.Stroke = blue
                        line.StrokeThickness = thick
                        if dashed:
                            line.StrokeDashArray = DoubleCollection([5.0, 3.0])
                        canvas.Children.Add(line)
            for clips, dx in regions:
                for x, z in dots:
                    if all((f((x, z)) >= 0) == keep for f, keep in clips):
                        e = Ellipse()
                        e.Width = e.Height = 6.0
                        e.Fill = blue
                        Canvas.SetLeft(e, X(x + dx) - 3.0)
                        Canvas.SetTop(e, Y(z) - 3.0)
                        canvas.Children.Add(e)
            return X, Y

        placed = set()

        def dimension(canvas, key, a, b, off, X, Y, extend=None):
            """A dimension from world point a to b: arrowheads, its box beside
            it by `off` (screen px), turned to run along it."""
            x1, y1, x2, y2 = X(a[0]), Y(a[1]), X(b[0]), Y(b[1])
            _dim_line(canvas, ink, x1, y1, x2, y2)
            ang = _dim_angle(x1, y1, x2, y2)
            self._place_editor(key, (x1 + x2) / 2.0 + off[0], (y1 + y2) / 2.0 + off[1], ang,
                               self._rb_editors)
            placed.add(key)

        # ---- the footing: the bottom three steps ----
        canvas = self.footing_canvas
        mapped = render(canvas, bottom, cuts_b, False, (58.0, 58.0, 40.0, 36.0))
        if mapped is not None:
            X, Y = mapped
            if plan["footing"]:
                leg_x = plan["main"][-1][0]
                hx, hz = plan["hook_centre"]
                # footing end (the first riser's face) to the hooked leg; its box
                # sits out to the left of the footing, clear of the bar
                mid_x = (X(0.0) + X(leg_x - dm / 2.0)) / 2.0
                dimension(canvas, "cover_end", (0.0, hz), (leg_x - dm / 2.0, hz),
                          (X(0.0) - 26.0 - mid_x, 0.0), X, Y)
                # footing bottom to the bottom of the hook: the line and its box
                # out to the right of the footing
                bottom_z = hz - plan["hook_radius"] - dm / 2.0
                xl = X(fw) + 16.0
                for zz, xz in ((bottom_z, X(hx)), (-fd, X(fw))):
                    _ext_line(canvas, ink, xz, Y(zz), xl + 4.0, Y(zz))
                _dim_line(canvas, ink, xl, Y(bottom_z), xl, Y(-fd))
                self._place_editor("cover_bottom", xl + 17.0, (Y(bottom_z) + Y(-fd)) / 2.0, -90.0,
                                   self._rb_editors)
                placed.add("cover_bottom")
            # soffit to the underside of the main bars, square to the slope
            k = plan["main_k"]
            xs_ = (plan["across"][0][0] if plan["across"]
                   else (plan["main"][1][0] + plan["main"][2][0]) / 2.0)
            under = (xs_ + sn * dm / 2.0, t * xs_ + k - cs * dm / 2.0)
            gap = s["waist"] - o["corner_clear"] - dd - dm
            soffit = (under[0] + sn * gap, under[1] - cs * gap)
            beyond = math.hypot(X(soffit[0]) - X(under[0]), Y(soffit[1]) - Y(under[1])) / 2.0 + 24.0
            dimension(canvas, "cover_soffit", soffit, under, (sn * beyond, cs * beyond), X, Y)
            # a step's inside corner to the top of the bars across, square to the slope
            n = 2 if N >= 3 else 1
            corner = (n * g, n * r)
            cc = o["corner_clear"]
            dimension(canvas, "corner_clear", corner,
                      (corner[0] + sn * cc, corner[1] - cs * cc), (-26.0, -12.0), X, Y)

        # ---- the top: the top three steps and the landing ----
        canvas = self.top_canvas
        slab = [(x_end, rise), (slab_end - 0.25 * thick, rise), (slab_end + 0.25 * thick, rise - thick),
                (x_end, rise - thick)]
        mapped = render(canvas, top, cuts_t, True, (24.0, 28.0, 64.0, 30.0), slab)
        if mapped is not None:
            X, Y = mapped
            x_top, z_top = plan["main"][1]
            lap_end = plan["main"][0][0]
            # the leg into the landing, dimensioned above the slab
            up = Y(s["rise"]) - 26.0
            for ex_ in (X(x_top), X(lap_end)):
                ext = Line()
                ext.X1, ext.Y1, ext.X2, ext.Y2 = ex_, Y(z_top) - 4.0, ex_, up - 4.0
                ext.Stroke = ink
                ext.StrokeThickness = 1
                ext.Opacity = 0.5
                canvas.Children.Add(ext)
            _dim_line(canvas, ink, X(x_top), up, X(lap_end), up)
            self._place_editor("lap", (X(x_top) + X(lap_end)) / 2.0, up - 14.0, 0.0, self._rb_editors)
            placed.add("lap")
            # the top of the landing to the top of the mesh (or of the top leg)
            if plan["mesh"]:
                xa, xb, z1 = plan["mesh"]
                top_z = z1 + opts["mesh_bar"].BarNominalDiameter * MM / 2.0
            else:
                top_z = z_top + dm / 2.0
            x5 = (x_top + lap_end) / 2.0                 # on the slab, over the top leg
            dimension(canvas, "cover_top", (x5, s["rise"]), (x5, top_z), (17.0, 0.0), X, Y)

        for key, box in self._rb_editors.items():
            if key not in placed:
                box.Visibility = Visibility.Collapsed

    def pick_level_clicked(self, sender, args):
        self._hide_for("level")

    def pick_width_clicked(self, sender, args):
        self._hide_for("width")

    def pick_riser_clicked(self, sender, args):
        self._hide_for("riser")

    def pick_going_clicked(self, sender, args):
        self._hide_for("going")

    def pick_top_landing_clicked(self, sender, args):
        self._hide_for("top_landing")

    def place_clicked(self, sender, args):
        self._hide_for("place")

    def close_clicked(self, sender, args):
        self.action = None
        self.Close()

    # --- results of picks ---
    def set_surface(self, z_mm):
        """Turn a picked height into the working RL, an offset from the level."""
        box = self.rl_tb
        level = self._levels[max(self.level_cb.SelectedIndex, 0)]
        box.Text = u"{0:g}".format(round(z_mm - level.ProjectElevation * MM, 1))


def _picked_height(ref):
    """Height (ft) of what was clicked: a level or a flat reference plane
    gives its own height; anything else, the point clicked on it."""
    element = doc.GetElement(ref)
    if isinstance(element, DB.Level):
        return element.ProjectElevation
    if isinstance(element, DB.ReferencePlane):
        normal = element.Normal
        if abs(normal.Z) > 0.999:
            return element.BubbleEnd.Z
        raise ValueError(u"That reference plane is not level. Pick a flat one, "
                         u"or a surface.")
    if ref.GlobalPoint is None:
        raise ValueError(u"Could not read a height from that. Click on a surface.")
    return ref.GlobalPoint.Z


def _pick_two(label):
    with forms.WarningBar(title=BAR_POINT.format(1, label)) as bar:
        a = _pick_xyz(BAR_POINT.format(1, label))
        bar.message_tb.Text = BAR_POINT.format(2, label)
        b = _pick_xyz(BAR_POINT.format(2, label))
    return a, b


def _do_pick(window, action):
    """Carry out one pick with the window hidden."""
    if action == "level":
        label = (u"top landing" if window.working() == u"top"
                 else u"base the first step stands on")
        with forms.WarningBar(title=BAR_SURFACE.format(label)):
            ref = uidoc.Selection.PickObject(ObjectType.Element,
                                             BAR_SURFACE.format(label))
        window.set_surface(_picked_height(ref) * MM)
    elif action == "width":
        a, b = _pick_two(u"stair width")
        window.width_tb.Text = u"{0:g}".format(round(
            DB.XYZ(a.X, a.Y, 0).DistanceTo(DB.XYZ(b.X, b.Y, 0)) * MM, 1))
    elif action == "top_landing":
        if window.update_target is not None:
            # the stair is there: one click, measured from its top riser
            _stairs_id, foot, direction = window.update_target
            _pick_top_landing(window, foot, direction)
        else:
            # not placed yet: the top riser is only known once P1 and P2 are
            window.top_landing_tb.Text = TL_PICK
    elif action == "going":
        a, b = _pick_two(u"going")
        window.going_tb.Text = u"{0:g}".format(round(
            DB.XYZ(a.X, a.Y, 0).DistanceTo(DB.XYZ(b.X, b.Y, 0)) * MM, 1))
    elif action == "riser":
        a, b = _pick_two(u"riser")
        window.riser_tb.Text = u"{0:g}".format(round(abs(b.Z - a.Z) * MM, 1))


def _ensure_fill(footing):
    """Load the footing and fill families before the transactions start
    (LoadFamily runs its own). A missing family template is reported."""
    if footing is None:
        return
    for name in (fill_family.FOOTING, fill_family.FAMILY):
        try:
            fill_family.ensure_family(doc, name)
        except Exception as err:
            dlg.message(u"The {0} family could not be made: {1}".format(name, err), title=TITLE)


def _footing_fill(stairs, frame, level, size, mat, old=(), old_id=None):
    """Put the Stair Footing and the fill under a stair (needs an open
    transaction). `old` are footings already there (a Stair Footing is kept
    and moved, resized and retyped to suit, anything else is replaced);
    `old_id` names the fill type of the stair this one replaces. With no
    size, the old footing and fill are removed."""
    old = list(old)
    for f in old:                   # their bars across are remade with the stair's
        rebar_engine.remove_bars_on(doc, f)
    keep = None
    if size is not None:
        mine = [f for f in old if fill_family.is_footing(f)]
        keep = mine[0] if mine else None
        if mat is None:
            mat = engine.stair_material(doc, stairs)
        fill_family.place_footing(doc, frame, level, size, mat, keep)
        fill_family.place(doc, stairs, frame, level, size[0], mat, old_id, footing_depth=size[1])
    elif old_id is not None:
        fill_family.remove(doc, old_id)
    for f in old:
        if f is not keep:
            doc.Delete(f.Id)


def _footing_and_bars(stairs, foot, direction, v, footing, rebar, old_footings=(), old_id=None):
    """Place the footing and fill (if any) and the bars (if any) for a stair
    just built from the window values, in one transaction. On an update,
    `old_footings` and `old_id` (the replaced stair's id, which names its
    fill type) are the previous stair's footing and fill: they are moved and
    resized to suit, or removed when the footing is now off."""
    if (footing is None and rebar is None and not old_footings and old_id is None
            and v.get("material") is None):
        return
    t = DB.Transaction(doc, u"Concrete Stair - footing, material and rebar")
    t.Start()
    frame = rebar_engine.Frame(foot, direction, v["sizes"], v["width"])
    mat = v.get("material")
    if mat is not None:
        engine.set_material(doc, stairs, mat)
    if footing is not None or old_footings or old_id is not None:
        _footing_fill(stairs, frame, v["base_level"], footing, mat, old_footings, old_id)
    if rebar is not None:
        rebar_engine.reinforce(doc, stairs, frame, v["base_level"], rebar)
    t.Commit()


def _pick_top_landing(window, foot, direction):
    """One click on the far edge of the top slab: TL is from the stair's top
    riser (`foot` is the bottom of the first riser, `direction` up the
    stair) to that point, along the stair in plan."""
    with forms.WarningBar(title=BAR_TL):
        pt = _pick_xyz(BAR_TL)
    d = DB.XYZ(direction.X, direction.Y, 0).Normalize()
    along = DB.XYZ(pt.X - foot.X, pt.Y - foot.Y, 0).DotProduct(d) * MM
    window.top_landing_tb.Text = u""                 # run length without TL
    run = window.values()["sizes"]["run_length"]
    window.top_landing_tb.Text = u"{0:g}".format(round(along - run, 1))


def _place(window):
    """Pick P1 and P2, build the stair and, if asked, the landing slab.
    Returns True when done, False to go back to the window."""
    v = window.values()
    footing = window.footing_options()  # read before the picks: a bad value stops here
    rebar = window.rebar_options()
    with forms.WarningBar(title=BAR_P1) as bar:
        p1 = _pick_xyz(BAR_P1)
        bar.message_tb.Text = BAR_P2
        p2 = _pick_xyz(BAR_P2)
    direction = DB.XYZ(p2.X - p1.X, p2.Y - p1.Y, 0)
    if direction.GetLength() < 1e-3:
        dlg.message(u"P2 must be away from P1 in plan.", title=TITLE)
        return False
    if v.get("pick_top_landing"):
        _pick_top_landing(window, p1, direction)
        v = window.values()

    _ensure_fill(footing)
    group = DB.TransactionGroup(doc, u"Concrete Stair")
    group.Start()
    try:
        t = DB.Transaction(doc, u"Concrete Stair - type")
        t.Start()
        stair_type = engine.ensure_type(doc, v["type_riser"], v["going"], v["waist"], v["landing"])
        t.Commit()
        stairs = engine.build(doc, stair_type, v["base_level"], v["base_offset"],
                              v["top_level"], v["top_offset"], p1, direction,
                              v["width"], v["sizes"])
        if not window.railings_cb.IsChecked:
            t = DB.Transaction(doc, u"Concrete Stair - railings")
            t.Start()
            engine.remove_railings(doc, stairs)
            t.Commit()
        foot = DB.XYZ(p1.X, p1.Y, v["base_level"].ProjectElevation + v["base_offset"] / MM)
        _footing_and_bars(stairs, foot, direction, v, footing, rebar)
        group.Assimilate()
    except Exception as err:
        if group.HasStarted() and not group.HasEnded():
            group.RollBack()
        dlg.message(u"The stair could not be made, nothing was changed.\n\n{0}".format(err),
                    title=TITLE)
        return False

    if window.slab_cb.IsChecked and window.floor_type() is not None:
        _make_slab(window, v)
    return True


class _StairsFilter(ISelectionFilter):
    """Stairs only."""

    def AllowElement(self, element):
        from Autodesk.Revit.DB.Architecture import Stairs
        return isinstance(element, Stairs)

    def AllowReference(self, reference, position):
        return False


def _pick_update(window):
    """Pick the stair to update and fill the window from it."""
    with forms.WarningBar(title=BAR_UPDATE):
        ref = uidoc.Selection.PickObject(ObjectType.Element, _StairsFilter(), BAR_UPDATE)
    stairs = doc.GetElement(ref)
    frame, level = rebar_engine.from_stairs(doc, stairs)
    window.load_stair(stairs, frame, level)


def _update_stair(window):
    """Rebuild the picked stair from the window, at the same start point and
    direction, with new bars. Its footing beam and fill are kept and moved
    and resized to suit (the fill type is renamed to the new stair's id).
    One undo. Returns True when done, False to go back."""
    v = window.values()
    footing = window.footing_options()
    rebar = window.rebar_options()
    stairs_id, foot, direction = window.update_target
    old = doc.GetElement(stairs_id)
    if old is None:
        dlg.message(u"That stair is no longer in the model.", title=TITLE)
        window.update_target = None
        return False

    _ensure_fill(footing)
    group = DB.TransactionGroup(doc, u"Concrete Stair - update")
    group.Start()
    try:
        old_footings = rebar_engine.find_footings(doc, old)
        t = DB.Transaction(doc, u"Concrete Stair - remove old")
        t.Start()
        doc.Delete(old.Id)               # its hosted bars go with it
        t.Commit()
        t = DB.Transaction(doc, u"Concrete Stair - type")
        t.Start()
        stair_type = engine.ensure_type(doc, v["type_riser"], v["going"], v["waist"], v["landing"])
        t.Commit()
        stairs = engine.build(doc, stair_type, v["base_level"], v["base_offset"],
                              v["top_level"], v["top_offset"], foot, direction,
                              v["width"], v["sizes"])
        if not window.railings_cb.IsChecked:
            t = DB.Transaction(doc, u"Concrete Stair - railings")
            t.Start()
            engine.remove_railings(doc, stairs)
            t.Commit()
        base = DB.XYZ(foot.X, foot.Y, v["base_level"].ProjectElevation + v["base_offset"] / MM)
        _footing_and_bars(stairs, base, direction, v, footing, rebar, old_footings, stairs_id)
        group.Assimilate()
    except Exception as err:
        if group.HasStarted() and not group.HasEnded():
            group.RollBack()
        dlg.message(u"The stair could not be updated, nothing was changed.\n\n{0}".format(err),
                    title=TITLE)
        return False
    return True


def _reinforce_existing(window):
    """Pick a stair already in the model and reinforce it with the Rebar
    tab's settings. Returns True when done, False to go back."""
    opts = window.rebar_options(True)
    with forms.WarningBar(title=BAR_STAIR):
        ref = uidoc.Selection.PickObject(ObjectType.Element, _StairsFilter(), BAR_STAIR)
    stairs = doc.GetElement(ref)
    frame, level = rebar_engine.from_stairs(doc, stairs)
    need = rebar_engine.min_waist(rebar_engine.numbers(opts))
    if frame.s["waist"] < need - 0.5:
        dlg.message(u"This stair's slope (SD) is {0:g}, but these bars need {1:g}.\n\n"
                    u"Use Update to select it, tick Reinforcement, then Apply: it raises SD to suit and "
                    u"adds the bars.".format(round(frame.s["waist"], 1), need), title=TITLE)
        return False
    _ensure_fill(opts["footing"])
    t = DB.Transaction(doc, u"Concrete Stair - rebar")
    t.Start()
    try:
        if opts["footing"] is not None:
            # resize a footing (and fill) already under the stair rather than doubling it
            _footing_fill(stairs, frame, level, opts["footing_size"], window.material(),
                          rebar_engine.find_footings(doc, stairs))
        rebar_engine.reinforce(doc, stairs, frame, level, opts)
        t.Commit()
    except Exception as err:
        t.RollBack()
        dlg.message(u"The reinforcing could not be added, nothing was changed.\n\n{0}"
                    .format(err), title=TITLE)
        return False
    return True


def _make_slab(window, v):
    try:
        with forms.WarningBar(title=BAR_LINES):
            refs = uidoc.Selection.PickObjects(ObjectType.Element, _CurveFilter(), BAR_LINES)
    except OperationCanceledException:
        return
    curves = [doc.GetElement(r).GeometryCurve for r in refs]
    t = DB.Transaction(doc, u"Concrete Stair - landing slab")
    t.Start()
    try:
        engine.landing_slab(doc, curves, window.floor_type(), v["top_level"], v["top_offset"])
        t.Commit()
    except Exception as err:
        t.RollBack()
        dlg.message(u"The landing slab could not be made:\n{0}".format(err), title=TITLE)


def main():
    window = StairWindow()
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
            elif action == "place" and window.update_target is not None:
                if _update_stair(window):
                    break
            elif action == "rebar_existing":
                if _reinforce_existing(window):
                    break
            elif action == "place":
                if _place(window):
                    break
            else:
                _do_pick(window, action)
        except OperationCanceledException:
            pass                        # ESC: back to the window
        except ValueError as err:
            dlg.message(unicode(err), title=TITLE)
    _save_settings(window.settings())


if __name__ == "__main__":
    main()
