# -*- coding: utf-8 -*-
"""Scrolling wheel time picker - the hour / minute / AM-PM columns used in
Seed43 schedule cards.

Extracted from pySheets, which had the only copy, when Batch Upgrade needed
the same control. pySheets still runs its own inlined version; switching it
over to this module is a separate job.

The widget owns three StackPanels (one per column) that you declare in XAML
inside a Popup, plus optionally the ToggleButton whose caption shows the
chosen time. Everything else - drag, mouse wheel, snapping, the fade on
rows either side of centre - is handled here.

Typical use in a tool:

    from Snippets._timepicker import WheelTimePicker, current_time_rounded

    self.picker = WheelTimePicker(self.tp_hours, self.tp_minutes,
                                  self.tp_ampm, button=self.sched_time_btn)
    self.picker.set_time(*current_time_rounded())
    ...
    def tp_ok_clicked(self, sender, args):
        h24, minute = self.picker.commit()
        self.sched_time_btn.IsChecked = False

snippets.yaml entry:
  _timepicker.py:
    description: >
      Scrolling wheel time picker (hour/minute/AM-PM) for schedule cards,
      shared by pySheets-style scheduling UI.
    functions:
      WheelTimePicker: Drive three wheel columns as one time picker control.
      time_label:      Format an hour/minute pair as the picker's caption.
      current_time_rounded: Now, rounded up to the next step-minute mark.
"""

# ── IMPORTS ────────────────────────────────────────────────────────────────

from datetime import datetime

import System
from pyrevit.framework import Windows
from System.Windows.Media import SolidColorBrush, Color

__all__ = ["WheelTimePicker", "time_label", "current_time_rounded"]


# ── CONSTANTS ──────────────────────────────────────────────────────────────

ROW_HEIGHT = 30
VISIBLE_ROWS = 5                    # rows shown in the wheel viewport
_PAD_ROWS = VISIBLE_ROWS // 2       # blank rows so row 0 can sit centred

_DEFAULT_FG = "#F4FAFF"


# ── HELPERS ────────────────────────────────────────────────────────────────

def _brush(hex_str):
    hex_str = (hex_str or _DEFAULT_FG).lstrip("#")
    return SolidColorBrush(Color.FromRgb(int(hex_str[0:2], 16),
                                         int(hex_str[2:4], 16),
                                         int(hex_str[4:6], 16)))


def time_label(h24, minute):
    """'08:05 PM' - the caption shown on the picker's button."""
    hour = ((h24 - 1) % 12) + 1
    return "{:02d}:{:02d} {}".format(hour, minute, "PM" if h24 >= 12 else "AM")


def current_time_rounded(step=5):
    """Now, rounded up to the next `step`-minute mark, as (hour24, minute)."""
    now = datetime.now()
    h24, minute = now.hour, now.minute
    minute = ((minute // step) + (1 if minute % step else 0)) * step
    if minute >= 60:
        minute = 0
        h24 = (h24 + 1) % 24
    return h24, minute


# ── CLASSES ────────────────────────────────────────────────────────────────

class WheelTimePicker(object):
    """Three scrolling wheels (hour, minute, AM/PM) driven as one control."""

    # --- construction ---
    def __init__(self, hours_panel, minutes_panel, ampm_panel,
                 button=None, minute_step=5, foreground=_DEFAULT_FG,
                 on_change=None):
        """
        Args:
            hours_panel/minutes_panel/ampm_panel: the three StackPanels from
                XAML that each wheel column is built into.
            button: optional ToggleButton whose Content shows the time.
            minute_step (int): minute granularity, 5 by default.
            foreground (str): hex colour for the wheel text. Baked in at
                build time, so rebuild the picker if the theme changes.
            on_change: optional callable invoked after commit().
        """
        self._sel = {}
        self._wheels = {}
        self._button = button
        self._step = minute_step
        self._fg = foreground
        self._on_change = on_change
        self._time = None

        self._build("h", hours_panel, [str(h) for h in range(1, 13)])
        self._build("m", minutes_panel,
                    ["{:02d}".format(m) for m in range(0, 60, minute_step)])
        self._build("ap", ampm_panel, ["AM", "PM"])

    # --- public methods ---
    def set_time(self, h24, minute):
        """Centre the wheels on a time and update the button caption."""
        self._time = (h24, minute)
        hour = ((h24 - 1) % 12) + 1
        self._center("h", hour - 1, animate=False)
        self._center("m", minute // self._step, animate=False)
        self._center("ap", 0 if h24 < 12 else 1, animate=False)
        self._label()

    def get_time(self):
        """The last committed time as (hour24, minute), or None."""
        return self._time

    def refresh(self):
        """Re-centre on the current time - call when the popup opens."""
        h24, minute = self._time or current_time_rounded(self._step)
        self.set_time(h24, minute)

    def commit(self):
        """Read the wheels, store and caption the result. Returns (h24, min)."""
        hour = self._sel.get("h", 12)
        minute = self._sel.get("m", 0)
        ampm = self._sel.get("ap", "AM")
        h24 = (hour % 12) + (12 if ampm == "PM" else 0)
        self._time = (h24, minute)
        self._label()
        if self._on_change:
            self._on_change(h24, minute)
        return h24, minute

    # --- private helpers: construction ---
    def _label(self):
        if self._button is not None and self._time:
            self._button.Content = time_label(*self._time)

    def _spacer(self):
        block = Windows.Controls.TextBlock()
        block.Height = ROW_HEIGHT
        return block

    def _build(self, group, container, items):
        """Build one scrollable, snapping wheel column inside `container`."""
        viewer = Windows.Controls.ScrollViewer()
        viewer.Height = ROW_HEIGHT * VISIBLE_ROWS
        viewer.Width = 46
        viewer.VerticalScrollBarVisibility = \
            Windows.Controls.ScrollBarVisibility.Hidden
        viewer.HorizontalScrollBarVisibility = \
            Windows.Controls.ScrollBarVisibility.Disabled
        viewer.PanningMode = Windows.Controls.PanningMode.VerticalOnly
        viewer.Focusable = False

        panel = Windows.Controls.StackPanel()
        blocks = []
        for _ in range(_PAD_ROWS):
            panel.Children.Add(self._spacer())
        for text in items:
            block = Windows.Controls.TextBlock()
            block.Text = text
            block.Height = ROW_HEIGHT
            block.FontSize = 15
            block.HorizontalAlignment = Windows.HorizontalAlignment.Center
            block.VerticalAlignment = Windows.VerticalAlignment.Center
            block.Foreground = _brush(self._fg)
            panel.Children.Add(block)
            blocks.append(block)
        for _ in range(_PAD_ROWS):
            panel.Children.Add(self._spacer())

        viewer.Content = panel
        container.Children.Clear()
        container.Children.Add(viewer)

        self._wheels[group] = {
            "items": items, "blocks": blocks, "viewer": viewer,
            "dragging": False, "start_y": 0.0, "start_off": 0.0,
        }

        # `g=group` binds the column at wiring time. Not strictly needed while
        # _build is called once per column, but it keeps the handlers correct
        # if these are ever wired from a loop over the three groups.
        viewer.PreviewMouseWheel += \
            lambda s, a, g=group: self._on_wheel(g, a)
        viewer.PreviewMouseLeftButtonDown += \
            lambda s, a, g=group: self._on_down(g, a)
        viewer.PreviewMouseMove += \
            lambda s, a, g=group: self._on_move(g, a)
        viewer.PreviewMouseLeftButtonUp += \
            lambda s, a, g=group: self._on_up(g, a)

    # --- private helpers: interaction ---
    def _on_wheel(self, group, args):
        idx = self._index(group) + (-1 if args.Delta > 0 else 1)
        self._center(group, idx)
        args.Handled = True

    def _on_down(self, group, args):
        state = self._wheels[group]
        state["dragging"] = True
        state["start_y"] = args.GetPosition(state["viewer"]).Y
        state["start_off"] = state["viewer"].VerticalOffset
        state["viewer"].CaptureMouse()

    def _on_move(self, group, args):
        state = self._wheels[group]
        if not state["dragging"]:
            return
        y = args.GetPosition(state["viewer"]).Y
        offset = state["start_off"] - (y - state["start_y"])
        offset = max(0.0, min((len(state["items"]) - 1) * ROW_HEIGHT, offset))
        state["viewer"].ScrollToVerticalOffset(offset)
        self._fade(group)

    def _on_up(self, group, args):
        state = self._wheels[group]
        if not state["dragging"]:
            return
        state["dragging"] = False
        state["viewer"].ReleaseMouseCapture()
        self._center(group, self._index(group))

    # --- private helpers: positioning ---
    def _index(self, group):
        state = self._wheels[group]
        idx = int(round(state["viewer"].VerticalOffset / float(ROW_HEIGHT)))
        return max(0, min(len(state["items"]) - 1, idx))

    def _center(self, group, idx, animate=True):
        """Snap so item `idx` sits centred, and record the selection."""
        state = self._wheels[group]
        idx = max(0, min(len(state["items"]) - 1, idx))
        value = state["items"][idx]
        self._sel[group] = value if group == "ap" else int(value)
        target = idx * ROW_HEIGHT
        if not animate:
            state["viewer"].ScrollToVerticalOffset(target)
            self._fade(group)
            return
        self._animate(group, target)

    def _animate(self, group, target, steps=6):
        """Ease the wheel toward `target` over a few timer ticks."""
        state = self._wheels[group]
        start = state["viewer"].VerticalOffset
        counter = {"n": 0}
        timer = Windows.Threading.DispatcherTimer()
        timer.Interval = System.TimeSpan.FromMilliseconds(15)

        def tick(sender, args):
            counter["n"] += 1
            t = counter["n"] / float(steps)
            if t >= 1.0:
                state["viewer"].ScrollToVerticalOffset(target)
                self._fade(group)
                timer.Stop()
                return
            eased = 1 - (1 - t) ** 2
            state["viewer"].ScrollToVerticalOffset(
                start + (target - start) * eased)
            self._fade(group)

        timer.Tick += tick
        timer.Start()

    def _fade(self, group):
        """Fade rows by distance from the centred (selected) row."""
        state = self._wheels[group]
        pos = state["viewer"].VerticalOffset / float(ROW_HEIGHT)
        for i, block in enumerate(state["blocks"]):
            block.Opacity = max(0.18, 1.0 - abs(i - pos) * 0.4)
