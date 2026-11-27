# -*- coding: utf-8 -*-
"""Organiser Setup: the one window that maps sheet parameters onto the views,
schedules and legends placed on those sheets, plus the casing options.

    from Snippets import _organiser_setup as setup

    cfg = setup.migrate_config(setup.load_config())
    if not setup.is_complete(cfg):
        cfg = setup.show_setup(doc, cfg, notice="There has been an update...")

The setup is saved in .user/ViewOrganiser/view_organiser_config.json. It is
asked for once; it comes back only when that file is deleted or is missing a
section (a new one added by an update). View Organiser and Organiser Reset
both use this module, so they always agree on the file.
"""
import os
import json

import clr
import wpf
from System.Windows import Window, Thickness, Visibility, GridLength, GridUnitType
from System.Windows.Controls import (ComboBox, Grid, ColumnDefinition, TextBlock,
                                     Border, StackPanel)
from pyrevit import DB
from Snippets import _userdata

# ── CONSTANTS ──────────────────────────────────────────────────────────────

HERE           = os.path.dirname(os.path.abspath(__file__))
XAML           = os.path.join(HERE, "OrganiserSetup.xaml")
CONFIG_PATH    = _userdata.user_path("ViewOrganiser", "view_organiser_config.json")
CONFIG_VERSION = 2
MAX_LINKS      = 10
NONE_LABEL     = "(None)"
MISSING_TAG    = "  (missing)"
SECTIONS       = ("views", "schedules", "legends")

CASE_OPTIONS = [
    ("ALL CAPS  (e.g. GROUND FLOOR PLAN)", "upper"),
    ("Title Case  (e.g. Ground Floor Plan)", "title"),
    ("Ignore  (leave as-is)", "ignore"),
]

# Views of these kinds are not ordinary views on a sheet
_NOT_PLAIN_VIEWS = ("DrawingSheet", "Legend", "Schedule", "Internal",
                    "ProjectBrowser", "SystemBrowser", "Report", "PanelSchedule",
                    "ColumnSchedule", "CostReport", "LoadsReport", "Undefined")

# ── CONFIG FILE ────────────────────────────────────────────────────────────

def load_config():
    """Return the saved config dict, or None if there is no usable file."""
    if os.path.isfile(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r") as f:
                return json.load(f)
        except Exception:
            return None
    return None


def migrate_config(raw):
    """Bring a saved config (any version) to the current shape. A section
    that was never set up is simply absent from "maps", which is how the
    caller knows the setup is incomplete."""
    cfg = {"version": CONFIG_VERSION, "maps": {}}
    if not raw:
        return cfg
    if isinstance(raw.get("maps"), dict):
        for k, v in raw["maps"].items():
            if k in SECTIONS and isinstance(v, list):
                cfg["maps"][k] = v
    else:
        # first-generation file: one folder link for views, one collection
        # link for schedules
        if "sheet_folder_param" in raw or "view_folder_param" in raw:
            s, t = raw.get("sheet_folder_param"), raw.get("view_folder_param")
            cfg["maps"]["views"] = [{"sheet": s, "target": t}] if s and t else []
        if raw.get("collection_asked"):
            s, t = raw.get("sheet_collection_param"), raw.get("schedule_collection_param")
            cfg["maps"]["schedules"] = [{"sheet": s, "target": t}] if s and t else []
    for k in ("title_on_sheet_case", "sheet_name_case"):
        if raw.get(k):
            cfg[k] = raw[k]
    return cfg


def missing_parts(cfg):
    """Names of what the config lacks (empty when it is complete)."""
    out = [k for k in SECTIONS if k not in cfg.get("maps", {})]
    if not cfg.get("title_on_sheet_case") or not cfg.get("sheet_name_case"):
        out.append("casing")
    return out


def is_complete(cfg):
    return not missing_parts(cfg)


def save_config(cfg):
    """Write the config. Returns True on success."""
    try:
        with open(CONFIG_PATH, "w") as f:
            json.dump(cfg, f, indent=2)
        return True
    except Exception:
        return False

# ── PARAMETER LISTS ────────────────────────────────────────────────────────

def _names(elements, kinds, writable):
    out = set()
    for el in elements:
        try:
            for p in el.Parameters:
                try:
                    if (p.Definition and p.Definition.Name and p.StorageType in kinds
                            and not (writable and p.IsReadOnly)):
                        out.add(p.Definition.Name)
                except Exception:
                    pass
        except Exception:
            pass
    return sorted(out)


def _section_sources(doc, limit=20):
    """{section: (title, singular, elements)} for what the project has;
    limit caps how many of each are sampled (None = all)."""
    sheets = list(DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).ToElements())
    views = [v for v in DB.FilteredElementCollector(doc).OfClass(DB.View)
             .WhereElementIsNotElementType().ToElements()]
    plain, legends, scheds = [], [], []
    for v in views:
        try:
            if v.IsTemplate:
                continue
            vt = str(v.ViewType)
            if vt == "Legend":
                legends.append(v)
            elif isinstance(v, DB.ViewSchedule):
                scheds.append(v)
            elif vt not in _NOT_PLAIN_VIEWS:
                plain.append(v)
        except Exception:
            pass
    return sheets, {
        "views":     ("Views", "view", plain[:limit]),
        "schedules": ("Schedules", "schedule", scheds[:limit]),
        "legends":   ("Legends", "legend", legends[:limit]),
    }

def _has_param(elements, name, kinds, writable):
    """True if any element carries a parameter of that name and kind."""
    for el in elements:
        try:
            for p in el.GetParameters(name):
                if p.StorageType in kinds and not (writable and p.IsReadOnly):
                    return True
        except Exception:
            pass
    return False


def unavailable_links(doc, cfg):
    """Describe every saved link whose parameter is no longer there: gone from
    the sheets, or gone / read-only on the views, schedules or legends. A kind
    the project has none of is not checked (nothing to look in)."""
    sheets, sources = _section_sources(doc, None)
    both = (DB.StorageType.String, DB.StorageType.ElementId)
    out = []
    for key in SECTIONS:
        title, single, elements = sources[key]
        for l in cfg.get("maps", {}).get(key) or []:
            if sheets and not _has_param(sheets, l.get("sheet"), both, False):
                out.append("{}: sheet parameter '{}'".format(title, l.get("sheet")))
            if elements and not _has_param(elements, l.get("target"),
                                           (DB.StorageType.String,), True):
                out.append("{}: {} parameter '{}'".format(
                    title, single, l.get("target")))
    return out

# ── WINDOW ─────────────────────────────────────────────────────────────────

class _Section(object):
    """The link rows of one section."""
    def __init__(self, key, rows):
        self.key = key
        self.rows = rows            # [(grid, source_combo, target_combo)]


class SetupWindow(Window):
    def __init__(self, doc, cfg, notice):
        wpf.LoadComponent(self, XAML)
        try:
            from Snippets.seed43_theme import apply_seed43_palette, apply_seed43_dimensions
            apply_seed43_palette(self, HERE)
            apply_seed43_dimensions(self, HERE)
        except Exception:
            pass
        self.result = None
        self._sections = []
        self._missing = 0
        self._cfg = cfg
        self.FindName("close_btn").Click += self._on_cancel
        self.FindName("cancel_btn").Click += self._on_cancel
        self.FindName("save_btn").Click += self._on_save
        if notice:
            self.FindName("notice_text").Text = notice
            self.FindName("notice_bar").Visibility = Visibility.Visible
        self._build(doc, cfg)

    # --- building ---
    def _style(self, key):
        return self.FindResource(key)

    def _label(self, text, style_key, margin=None):
        tb = TextBlock()
        tb.Text = text
        tb.Style = self._style(style_key)
        if margin is not None:
            tb.Margin = margin
        return tb

    def _combo(self, items, selected):
        cb = ComboBox()
        cb.Style = self._style("ComboBoxStyle")
        cb.Margin = Thickness(0)
        for it in items:
            cb.Items.Add(it)
        cb.SelectedItem = selected if selected in items else items[0]
        return cb

    def _two_columns(self):
        g = Grid()
        for _ in range(2):
            c = ColumnDefinition()
            c.Width = GridLength(1, GridUnitType.Star)
            g.ColumnDefinitions.Add(c)
        return g

    def _build(self, doc, cfg):
        panel = self.FindName("sections_panel")
        sheets, sources = _section_sources(doc)
        sheet_names = _names(sheets[:10],
                             (DB.StorageType.String, DB.StorageType.ElementId), False)

        for key in SECTIONS:
            title, single, sample = sources[key]
            saved = cfg.get("maps", {}).get(key) or []
            target_names = _names(sample, (DB.StorageType.String,), True)
            # a saved link must stay selectable even if nothing in the project
            # shows its parameter right now
            # a saved parameter the project no longer has stays in the list,
            # tagged "(missing)", so it is visible and can be replaced
            def _tag(name, have, check):
                return name if (not check or name in have) else name + MISSING_TAG
            s_gone = [_tag(l["sheet"], sheet_names, bool(sheets))
                      for l in saved if l.get("sheet")]
            t_gone = [_tag(l["target"], target_names, bool(sample))
                      for l in saved if l.get("target")]
            s_items = [NONE_LABEL] + sorted(set(sheet_names + s_gone))
            t_items = [NONE_LABEL] + sorted(set(target_names + t_gone))

            panel.Children.Add(self._label(
                title, "SectionLabelStyle", Thickness(0, 0, 0, 8)))
            card = Border()
            card.Style = self._style("CardStyle")
            stack = StackPanel()
            card.Child = stack
            stack.Children.Add(self._label(
                "Copy a sheet parameter to the {}s placed on that sheet".format(single),
                "FieldLabelStyle", Thickness(0, 0, 0, 8)))
            if not sample:
                stack.Children.Add(self._label(
                    "There are no {} in this project yet, so their parameters "
                    "cannot be listed. Use Organiser Reset later to link them.".format(
                        title.lower()),
                    "FieldLabelStyle", Thickness(0, 0, 0, 8)))

            head = self._two_columns()
            head.Margin = Thickness(0, 0, 0, 4)
            h1 = self._label("SHEET PARAMETER", "FieldLabelStyle", Thickness(0, 0, 8, 0))
            h2 = self._label("{} PARAMETER".format(single.upper()), "FieldLabelStyle")
            Grid.SetColumn(h1, 0)
            Grid.SetColumn(h2, 1)
            head.Children.Add(h1)
            head.Children.Add(h2)
            stack.Children.Add(head)

            rows = []
            for n in range(MAX_LINKS):
                link = dict(saved[n]) if n < len(saved) else {}
                if link.get("sheet") and sheets and link["sheet"] not in sheet_names:
                    link["sheet"] += MISSING_TAG
                if link.get("target") and sample and link["target"] not in target_names:
                    link["target"] += MISSING_TAG
                if MISSING_TAG in (link.get("sheet", "") + link.get("target", "")):
                    self._missing += 1
                g = self._two_columns()
                g.Margin = Thickness(0, 0, 0, 6)
                sc = self._combo(s_items, link.get("sheet"))
                tc = self._combo(t_items, link.get("target"))
                sc.Margin = Thickness(0, 0, 8, 0)
                Grid.SetColumn(sc, 0)
                Grid.SetColumn(tc, 1)
                g.Children.Add(sc)
                g.Children.Add(tc)
                sc.SelectionChanged += self._on_change
                tc.SelectionChanged += self._on_change
                stack.Children.Add(g)
                rows.append((g, sc, tc))
            panel.Children.Add(card)
            self._sections.append(_Section(key, rows))

        # casing
        panel.Children.Add(self._label("Casing", "SectionLabelStyle", Thickness(0, 0, 0, 8)))
        card = Border()
        card.Style = self._style("CardStyle")
        stack = StackPanel()
        card.Child = stack
        labels = [c[0] for c in CASE_OPTIONS]
        by_code = dict((c[1], c[0]) for c in CASE_OPTIONS)
        stack.Children.Add(self._label("Title on Sheet", "FieldLabelStyle"))
        self._case_title = self._combo(labels, by_code.get(cfg.get("title_on_sheet_case"), labels[0]))
        self._case_title.Margin = Thickness(0, 0, 0, 10)
        stack.Children.Add(self._case_title)
        stack.Children.Add(self._label("Sheet Name", "FieldLabelStyle"))
        self._case_sheet = self._combo(labels, by_code.get(cfg.get("sheet_name_case"), labels[0]))
        stack.Children.Add(self._case_sheet)
        panel.Children.Add(card)

        self._reveal()
        if self._missing:
            self.FindName("status_tb").Text = (
                "{} parameter(s) marked (missing) are not in this project. "
                "Choose another, or (None), before saving.".format(self._missing))

    # --- behaviour ---
    def _filled(self, row):
        return row[1].SelectedItem != NONE_LABEL and row[2].SelectedItem != NONE_LABEL

    def _reveal(self):
        """Show the rows in use plus one empty row to add the next link."""
        for sec in self._sections:
            last = -1
            for i, row in enumerate(sec.rows):
                if self._filled(row):
                    last = i
            for i, row in enumerate(sec.rows):
                row[0].Visibility = (Visibility.Visible if i <= last + 1
                                     else Visibility.Collapsed)

    def _on_change(self, sender, e):
        self._reveal()
        self.FindName("status_tb").Text = ""

    def _on_cancel(self, sender, e):
        self.result = None
        self.Close()

    def _on_save(self, sender, e):
        maps = {}
        for sec in self._sections:
            links, used = [], set()
            for i, row in enumerate(sec.rows):
                s, t = row[1].SelectedItem, row[2].SelectedItem
                if (s == NONE_LABEL) != (t == NONE_LABEL):
                    self.FindName("status_tb").Text = (
                        "{} link {}: choose both the sheet and the receiving "
                        "parameter, or set both to (None).".format(sec.key.title(), i + 1))
                    return
                if s == NONE_LABEL:
                    continue
                for v in (s, t):
                    if MISSING_TAG in str(v):
                        self.FindName("status_tb").Text = (
                            "'{}' is missing from this project - choose another "
                            "or (None).".format(str(v).replace(MISSING_TAG, "")))
                        return
                if t in used:
                    self.FindName("status_tb").Text = (
                        "{}: '{}' receives more than one sheet parameter.".format(
                            sec.key.title(), t))
                    return
                used.add(t)
                links.append({"sheet": str(s), "target": str(t)})
            maps[sec.key] = links
        code = dict((c[0], c[1]) for c in CASE_OPTIONS)
        cfg = {
            "version": CONFIG_VERSION,
            "maps": maps,
            "title_on_sheet_case": code[self._case_title.SelectedItem],
            "sheet_name_case": code[self._case_sheet.SelectedItem],
        }
        self.result = cfg
        self.Close()

# ── ENTRY POINT ────────────────────────────────────────────────────────────

def show_setup(doc, cfg, notice=None):
    """Open the setup window pre-filled from cfg. Returns the new config dict
    after Save (also written to the JSON file), or None if cancelled."""
    win = SetupWindow(doc, cfg, notice)
    win.ShowDialog()
    if win.result is not None:
        save_config(win.result)
    return win.result
