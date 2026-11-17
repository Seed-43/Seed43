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
from System.Windows import Point, FontWeights, TextAlignment, Visibility, Size
from System.Windows.Controls import Canvas, TextBlock
from System.Windows.Markup import XamlReader
from System.Windows.Media import (PointCollection, RotateTransform, SolidColorBrush,
                                  Color, DoubleCollection)
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
SETTINGS = _userdata.user_path(u"Concrete Stair", u"settings.json")
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


# ── UI / ENTRY POINT ────────────────────────────────────────────────────────

class StairWindow(forms.WPFWindow):
    """The Concrete Stair dialogue. Picks hide it; it comes back after."""

    # Class level, not just set in __init__: the XAML's own Text values fire
    # TextChanged while the window is still loading, before __init__ runs on.
    _ready = False
    _updating = False

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

        s = _load_settings()
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
        self._materials = sorted(DB.FilteredElementCollector(doc).OfClass(DB.Material),
                                 key=lambda m: m.Name.lower())
        self.material_cb.Items.Add(KEEP_MATERIAL)
        for m in self._materials:
            self.material_cb.Items.Add(m.Name)
        self._select_material_name(s.get("material", u""))
        self._load_rebar(s)
        self._ready = True
        self._update()

    # --- tabs ---
    def _show_tab(self, name):
        stair = name == "stair"
        self.page_stair.Visibility = Visibility.Visible if stair else Visibility.Collapsed
        self.page_rebar.Visibility = Visibility.Collapsed if stair else Visibility.Visible
        self.tab_stair_btn.Tag = u"Active" if stair else u""
        self.tab_rebar_btn.Tag = u"" if stair else u"Active"
        self.header_subtitle_tb.Text = u"| Stair" if stair else u"| Rebar"

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
            self.result_tb.Text = u""
            self.type_tb.Text = u""
            self.status_tb.Text = unicode(err)
            self.place_btn.IsEnabled = False
            self.drawing.Children.Clear()
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
        try:
            self._draw_rebar(v["sizes"], self.rebar_options(True))
            self.rebar_note_tb.Text = u""
        except ValueError as err:
            self.rebar_drawing.Children.Clear()
            self.rebar_note_tb.Text = unicode(err)
        self.result_tb.Text = (
            u"Rise {0:.0f}   ·   {1} risers of {2:.1f}   ·   {3} treads of {4:.0f}\n"
            u"Run {5:.0f} long   ·   top landing {6:.0f} {10}\n"
            u"Slope {8:.1f}°   ·   total length {9:.0f}".format(
                s["rise"], s["risers"], s["riser"], s["treads"], s["going"],
                s["run_length"], s["xyz"], s["landing"], s["angle"],
                s["run_length"] + s["xyz"],
                (u"for now: TL is picked after P1 and P2" if v.get("pick_top_landing") else
                 u"(set by TL)" if s["xyz"] > s["xyz_min"] + 0.05 else
                 u"(the least: TL {0:g} was too short)".format(
                     float(self.top_landing_tb.Text)) if v.get("top_landing_short") else
                 u"(the least, to meet the {0:.0f} landing underside)".format(s["landing"]))))
        self.type_tb.Text = u"Type: " + engine.type_name(v["type_riser"], v["going"],
                                                          v["waist"], v["landing"])
        if "waist_entered" in v:
            self.type_tb.Text += (u"\nSlope SD raised from {0:g} to {1:g} to fit the bars "
                                  u"(see Rebar > Cover)".format(v["waist_entered"], v["waist"]))

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

    def _draw(self, v):
        """Redraw the section to scale for the step count, every dimension a
        label. Canvas units are pixels, y down."""
        s = v["sizes"]
        top_working = self.working() == u"top"
        N, r, g = s["risers"], s["riser"], s["going"]
        run_len, xyz = s["run_length"], s["xyz"]
        x_end = run_len + xyz
        foot = min(s["soffit_foot"], x_end)
        fw, fd = self._footing_drawn()
        scale = min(400.0 / max(x_end, 1.0), 240.0 / max(s["rise"] + fd, 1.0))
        height = s["rise"] * scale
        ox = 66.0 if fd else 30.0                    # room for the FD dimension
        below = fd * scale + 34.0 if fd else 0.0     # the footing and its FW dimension
        oy = min(395.0 - below, max(height + 80.0, 195.0 + height / 2.0))

        def X(x):
            return ox + x * scale

        def Y(z):
            return oy - z * scale

        canvas = self.drawing
        canvas.Children.Clear()
        ink = self.TryFindResource("BrushTextPrimary")
        fill = self.TryFindResource("BrushSecondaryBg")
        accent = self.TryFindResource("BrushPrimaryGreen")

        def line(x1, y1, x2, y2, opacity=0.7):
            ln = Line()
            ln.X1, ln.Y1, ln.X2, ln.Y2 = x1, y1, x2, y2
            ln.Stroke = ink
            ln.StrokeThickness = 1
            ln.Opacity = opacity
            canvas.Children.Add(ln)

        def label(text, left, top, colour=None, bold=True, opacity=1.0, centre=False):
            tb = TextBlock()
            tb.Text = text
            tb.FontSize = 12
            tb.Foreground = colour or accent
            tb.Opacity = opacity
            if bold:
                tb.FontWeight = FontWeights.SemiBold
            if centre:
                tb.Width = 80
                tb.TextAlignment = TextAlignment.Center
                left -= 40
            Canvas.SetLeft(tb, left)
            Canvas.SetTop(tb, top)
            canvas.Children.Add(tb)
            return tb

        def turned(text, cx, cy, angle):
            """A label centred on (cx, cy), turned to run along a dimension
            line, the way a drawing's dimension text is."""
            tb = label(text, 0, 0)
            tb.Width = 60
            tb.Height = 16
            tb.TextAlignment = TextAlignment.Center
            tb.RenderTransformOrigin = Point(0.5, 0.5)
            tb.RenderTransform = RotateTransform(angle)
            Canvas.SetLeft(tb, cx - 30)
            Canvas.SetTop(tb, cy - 8)

        def num(x):
            return u"{0:g}".format(round(x, 1))

        def rl_marker(x, z, text, working):
            tri = Polygon()
            tri.Points = PointCollection([Point(x - 7, Y(z) - 11), Point(x + 7, Y(z) - 11),
                                          Point(x, Y(z) - 1)])
            tri.Fill = ink
            canvas.Children.Add(tri)
            # the working RL in the accent, the worked-out one dimmed
            label(u"RL " + text, x + 10, Y(z) - 24,
                  accent if working else ink, working, 1.0 if working else 0.6)

        # outline: up the steps, along the landing, back down the soffit
        pts = [(0, 0)]
        for k in range(1, N + 1):
            pts.append(((k - 1) * g, k * r))
            if k < N:
                pts.append((k * g, k * r))
        pts += [(x_end, s["rise"]), (x_end, s["rise"] - s["landing"]), (foot, 0)]
        poly = Polygon()
        poly.Points = PointCollection([Point(X(x), Y(z)) for x, z in pts])
        poly.Fill = fill
        poly.Stroke = ink
        poly.StrokeThickness = 1.5
        canvas.Children.Add(poly)

        # footing under the first step, its fill, and its FW / FD dimensions
        if fd:
            beam = Polygon()
            beam.Points = PointCollection([Point(X(x), Y(z)) for x, z in
                                           [(0, 0), (fw, 0), (fw, -fd), (0, -fd)]])
            beam.Fill = fill
            beam.Stroke = ink
            beam.StrokeThickness = 1.5
            canvas.Children.Add(beam)
            shape = fill_family.fill_shape(s, fw, fd)
            if shape:
                run_f, rise_f, under = shape
                tri_pts = ([(fw, 0), (fw + run_f, 0), (fw, -rise_f)] if under
                           else [(fw - run_f, 0), (fw, 0), (fw, rise_f)])
                tri = Polygon()
                tri.Points = PointCollection([Point(X(x), Y(z)) for x, z in tri_pts])
                tri.Fill = fill
                tri.Stroke = ink
                tri.StrokeThickness = 1
                tri.Opacity = 0.8
                canvas.Children.Add(tri)
            # FW under the footing
            wy = Y(-fd) + 14
            line(X(0), wy, X(fw), wy)
            line(X(0), Y(-fd) + 4, X(0), wy + 5, 0.4)
            line(X(fw), Y(-fd) + 4, X(fw), wy + 5, 0.4)
            label(num(fw), X(fw / 2.0), wy, centre=True)
            _badge(canvas, u"FW", X(fw / 2.0) + 30, wy + 9)
            # FD beside the footing's outer face
            dx = X(0) - 14
            line(dx, Y(0), dx, Y(-fd))
            line(dx - 5, Y(-fd), X(0) - 4, Y(-fd), 0.4)
            turned(num(fd), dx - 11, Y(-fd / 2.0), -90)
            _badge(canvas, u"FD", dx - 11, Y(-fd / 2.0) + 30)

        if N <= 20:
            for k in range(1, N + 1):
                label(unicode(k), X((k - 1) * g) + 4, Y(k * r) + 2, opacity=0.8)

        # riser: out to the left of the top riser, on extension lines, so it
        # works for one step as well as many
        top_x = X((N - 1) * g)
        rx = top_x - 34
        line(rx, Y(s["rise"]), rx, Y(s["rise"] - r))
        line(rx - 5, Y(s["rise"]), top_x - 4, Y(s["rise"]), 0.4)
        line(rx - 5, Y(s["rise"] - r), top_x - 4, Y(s["rise"] - r), 0.4)
        turned(num(r), rx - 11, Y(s["rise"] - r / 2.0), -90)
        _badge(canvas, u"SR", rx - 11, Y(s["rise"]) - 14)

        # going: on the top row beside the landing's dimension, over the last
        # tread (one step has no tread, so no going to show)
        if N >= 2:
            ga, gb = X((N - 2) * g), X((N - 1) * g)
            gy = Y(s["rise"]) - 40
            line(ga, gy, gb, gy)
            line(ga, gy - 5, ga, Y((N - 1) * r) - 4, 0.4)
            line(gb, gy - 5, gb, Y(s["rise"]) - 4, 0.4)
            label(num(g), (ga + gb) / 2.0, gy - 18, centre=True)
            _badge(canvas, u"SG", (ga + gb) / 2.0 - 30, gy - 9)

        # slope / waist: from the inside corner at the foot of the second
        # riser, square down to the soffit - where it is measured
        # (one step has no inside corner to measure from, so no SD to show)
        slope = math.atan2(r, g)
        if N >= 2:
            nx, nz = math.sin(slope), -math.cos(slope)
            cx, cz = g, r
            ex, ez = cx + nx * s["waist"], cz + nz * s["waist"]
            line(X(cx), Y(cz), X(ex), Y(ez), 0.6)
            mid_x, mid_y = (X(cx) + X(ex)) / 2.0, (Y(cz) + Y(ez)) / 2.0
            angle = math.degrees(math.atan2(Y(ez) - Y(cz), X(ex) - X(cx)))
            # text on the downhill side of the line, reading along it
            turned(num(s["waist"]), mid_x + math.cos(slope) * 12,
                   mid_y - math.sin(slope) * 12, angle)
            _badge(canvas, u"SD", mid_x + math.cos(slope) * 36,
                   mid_y - math.sin(slope) * 36 + 10)

        # slab: at the far end, text reading up the dimension
        lx = X(x_end) + 10
        line(lx, Y(s["rise"]), lx, Y(s["rise"] - s["landing"]))
        line(X(x_end), Y(s["rise"] - s["landing"]), lx + 5, Y(s["rise"] - s["landing"]), 0.4)
        turned(num(s["landing"]), lx + 12, Y(s["rise"] - s["landing"] / 2.0), -90)
        _badge(canvas, u"LD", lx + 12, Y(s["rise"] - s["landing"]) + 16)

        # xyz: over the top landing
        xy = Y(s["rise"]) - 40
        line(X(run_len), xy, X(x_end), xy)
        line(X(run_len), xy - 5, X(run_len), Y(s["rise"]) - 4, 0.4)
        line(X(x_end), xy - 5, X(x_end), Y(s["rise"]) - 4, 0.4)
        label(num(xyz), (X(run_len) + X(x_end)) / 2.0, xy - 18, centre=True)
        _badge(canvas, u"TL", (X(run_len) + X(x_end)) / 2.0 + 30, xy - 9)

        # RLs: top just past the end of the landing, base on the ground
        line(X(x_end) + 4, Y(s["rise"]), X(x_end) + 70, Y(s["rise"]))
        rl_marker(X(x_end) + 24, s["rise"], num(v["top_offset"]), top_working)
        line(X(foot), Y(0), X(foot) + 150, Y(0))
        rl_marker(X(foot) + 40, 0, num(v["base_offset"]), not top_working)

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
        self.place_btn.Content = u"Update stair"
        self.place_btn.ToolTip = (u"Rebuild stair {0} in the same place with these settings; "
                                  u"its bars and footing are replaced too".format(
                                      rebar_engine.eid(stairs.Id)))
        self.header_subtitle_tb.Text = u"| Update"
        self._update()

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
        scale = min(560.0 / max(x_max, 1.0), 250.0 / max(s["rise"] - z_min, 1.0))
        ox = 40.0
        oy = 30.0 + s["rise"] * scale

        def X(x):
            return ox + x * scale

        def Y(z):
            return oy - z * scale

        canvas = self.rebar_drawing
        canvas.Children.Clear()
        ink = self.TryFindResource("BrushTextPrimary")
        fill = self.TryFindResource("BrushSecondaryBg")
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
        # the stair
        pts = [(0, 0)]
        for k in range(1, N + 1):
            pts.append(((k - 1) * g, k * r))
            if k < N:
                pts.append((k * g, k * r))
        pts += [(x_end, s["rise"]), (x_end, s["rise"] - s["landing"]), (foot, 0)]
        poly(pts, True, ink, 1.5, fill)
        # footing, and the fill beside it
        if plan["footing"]:
            poly([(0, 0), (fw, 0), (fw, -fd), (0, -fd)], True, ink, 1.5, fill)
            shape = fill_family.fill_shape(s, fw, fd)
            if shape:
                run_f, rise_f, under = shape
                poly([(fw, 0), (fw + run_f, 0), (fw, -rise_f)] if under
                     else [(fw - run_f, 0), (fw, 0), (fw, rise_f)], True, ink, 1, fill,
                     opacity=0.8)

        # main bars, with the hook drawn as a U turning uphill
        main = list(plan["main"])
        if plan["hook"]:
            # the leg runs down to the hook's centre height, then round the U
            hx, hz = plan["hook_centre"]
            rad = plan["hook_radius"]
            main[-1] = (hx - rad, hz)
            main += [(hx - rad * math.cos(a), hz - rad * math.sin(a))
                     for a in [math.pi * i / 8.0 for i in range(1, 9)]]
            main.append((hx + rad, hz + rad * 1.5))
        poly(main, False, bar, 2.5)

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

        self._cover_dims(plan, s, opts, X, Y, canvas, ink, accent)

    def _cover_dims(self, plan, s, opts, X, Y, canvas, ink, accent):
        """The covers as numbered dimensions on the rebar drawing, each
        measured off the layout (so it shows what the bars get, e.g. 72 where
        SD rounded up). The numbers match the Cover inputs: 1 footing ends,
        2 footing bottom, 3 soffit, 4 below the step corners, 5 top."""
        o = rebar_engine.numbers(opts)
        r, g = s["riser"], s["going"]
        slope = math.atan2(r, g)
        sn, cs = math.sin(slope), math.cos(slope)
        t = r / g
        dm, dd = o["main_dm"], o["dist_dm"]

        def dim(a, b, value, number, side=1):
            """A dimension from a to b (section mm): ticks at both ends, the
            value beside its middle and the input's number past that."""
            (x1, y1), (x2, y2) = (X(a[0]), Y(a[1])), (X(b[0]), Y(b[1]))
            length = math.hypot(x2 - x1, y2 - y1)
            if length < 1.0:
                return
            ux, uy = (x2 - x1) / length, (y2 - y1) / length
            for x0, y0, x3, y3 in ((x1, y1, x2, y2),
                                   (x1 - 3 * (ux + uy), y1 - 3 * (uy - ux),
                                    x1 + 3 * (ux + uy), y1 + 3 * (uy - ux)),
                                   (x2 - 3 * (ux + uy), y2 - 3 * (uy - ux),
                                    x2 + 3 * (ux + uy), y2 + 3 * (uy - ux))):
                ln = Line()
                ln.X1, ln.Y1, ln.X2, ln.Y2 = x0, y0, x3, y3
                ln.Stroke = ink
                ln.StrokeThickness = 1
                ln.Opacity = 0.75
                canvas.Children.Add(ln)
            tb = TextBlock()
            tb.Text = u"{0:g}".format(round(value))
            tb.FontSize = 11
            tb.Foreground = accent
            mx, my = (x1 + x2) / 2.0, (y1 + y2) / 2.0
            ox_, oy_ = -side * uy, side * ux            # off the line, to one side
            Canvas.SetLeft(tb, mx + ox_ * 13 - 8)
            Canvas.SetTop(tb, my + oy_ * 13 - 8)
            canvas.Children.Add(tb)
            _badge(canvas, number, mx + ox_ * 32, my + oy_ * 32)

        if plan["footing"]:
            fd = plan["footing"][3] - plan["footing"][2]
            leg_x = plan["main"][-1][0]
            hx, hz = plan["hook_centre"]
            # 1  footing end (the first riser face) to the hooked leg
            dim((0.0, hz), (leg_x - dm / 2.0, hz), leg_x - dm / 2.0, u"1", -1)
            # 2  footing bottom to the bottom of the hook
            bottom = hz - plan["hook_radius"] - dm / 2.0
            dim((hx, -fd), (hx, bottom), bottom + fd, u"2")
        # 3  soffit to the underside of the main bars, square to the slope
        k = plan["main_k"]
        xs = plan["across"][0][0] if plan["across"] else (plan["main"][1][0] + plan["main"][2][0]) / 2.0
        under = (xs + sn * dm / 2.0, t * xs + k - cs * dm / 2.0)
        gap = s["waist"] - o["corner_clear"] - dd - dm
        dim((under[0] + sn * gap, under[1] - cs * gap), under, gap, u"3", -1)
        # 4  step corner to the top of the bars across, square to the slope
        n = 2 if s["risers"] >= 3 else 1
        corner = (n * g, n * r)
        cc = o["corner_clear"]
        dim(corner, (corner[0] + sn * cc, corner[1] - cs * cc), cc, u"4")
        # 5  top of the landing to the top of the mesh (or of the top leg)
        if plan["mesh"]:
            xa, xb, z1 = plan["mesh"]
            top = z1 + opts["mesh_bar"].BarNominalDiameter * 304.8 / 2.0
            x5 = xb - 40.0
        else:
            top = plan["z_top"] + dm / 2.0
            x5 = plan["main"][0][0] - 40.0
        dim((x5, s["rise"]), (x5, top), s["rise"] - top, u"5", -1)

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
                    u"Use Update a stair... with Rebar ticked: it raises SD to suit and "
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
