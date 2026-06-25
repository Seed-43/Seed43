# -*- coding: utf-8 -*-
# BreakTime - Seed43
# Randomly launches one of 10 classic DOS-style games for a 2 minute break.

import os
import random
import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xml")

from System.Windows import Application, Window, WindowStartupLocation, ResizeMode
from System.Windows.Markup import XamlReader
from System.Xml import XmlReader
from System.IO import StringReader, StreamReader
from System.Windows.Threading import DispatcherTimer
from System.Windows.Controls import (StackPanel, Button, TextBlock,
                                      ScrollViewer, Canvas as WpfCanvas)
from System.Windows.Media import SolidColorBrush, Color, FontFamily
from System.Windows import Thickness, HorizontalAlignment, VerticalAlignment, FontWeights
from System import TimeSpan

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(__file__)
GAMES_DIR  = os.path.join(SCRIPT_DIR, "games")
XAML_PATH  = os.path.join(SCRIPT_DIR, "ui.xaml")

# ---------------------------------------------------------------------------
# Game registry
# ---------------------------------------------------------------------------
GAMES = [
    ("SNAKE",          "snake.py"),
    ("BOWLING",        "bowling.py"),
    ("MINESWEEPER",    "minesweeper.py"),
    ("TETRIS",         "tetris.py"),
    ("PONG",           "pong.py"),
    ("SPACE INVADERS", "space_invaders.py"),
    ("BREAKOUT",       "breakout.py"),
    ("FROGGER",        "frogger.py"),
    ("PAC-MAN",        "pacman.py"),
    ("ASTEROIDS",      "asteroids.py"),
]

TOTAL_SECONDS = 120  # 2 minutes

# ---------------------------------------------------------------------------
# Colours
# ---------------------------------------------------------------------------
def _brush(r, g, b):
    return SolidColorBrush(Color.FromRgb(r, g, b))

def _font(size=13, bold=False):
    tb = TextBlock()
    tb.FontFamily = FontFamily("Courier New")
    tb.FontSize   = size
    if bold:
        tb.FontWeight = FontWeights.Bold
    return tb

# ---------------------------------------------------------------------------
# Picker window - shown first, returns chosen (name, file) tuple
# ---------------------------------------------------------------------------
def show_picker():
    win = Window()
    win.Title                  = "SEED43 BREAK TIME  |  SELECT GAME"
    win.Width                  = 420
    win.Height                 = 520
    win.WindowStartupLocation  = WindowStartupLocation.CenterScreen
    win.ResizeMode             = ResizeMode.NoResize
    win.Background             = _brush(0, 13, 0)

    root = StackPanel()
    root.Margin = Thickness(20, 16, 20, 16)
    win.Content = root

    # Header
    hdr = _font(15, True)
    hdr.Text       = "C:\\SEED43\\GAMES>_"
    hdr.Foreground = _brush(0, 80, 20)
    hdr.Margin     = Thickness(0, 0, 0, 4)
    root.Children.Add(hdr)

    sub = _font(13, True)
    sub.Text       = "SELECT YOUR GAME"
    sub.Foreground = _brush(0, 255, 65)
    sub.Margin     = Thickness(0, 0, 0, 14)
    root.Children.Add(sub)

    # Result holder - mutable so inner functions can write to it
    result = [None]

    def make_btn(name, filename):
        btn = Button()
        btn.Background          = _brush(0, 25, 5)
        btn.BorderBrush         = _brush(0, 120, 30)
        btn.BorderThickness     = Thickness(1)
        btn.Margin              = Thickness(0, 3, 0, 3)
        btn.Padding             = Thickness(10, 6, 10, 6)
        btn.HorizontalContentAlignment = HorizontalAlignment.Left
        tb = _font(13, True)
        tb.Text       = ">  " + name
        tb.Foreground = _brush(0, 255, 65)
        btn.Content   = tb

        def on_hover_enter(s, e):
            s.Background  = _brush(0, 60, 10)
            tb.Foreground = _brush(255, 220, 0)

        def on_hover_leave(s, e):
            s.Background  = _brush(0, 25, 5)
            tb.Foreground = _brush(0, 255, 65)

        def on_click(s, e, n=name, f=filename):
            result[0] = (n, f)
            win.Close()

        btn.MouseEnter  += on_hover_enter
        btn.MouseLeave  += on_hover_leave
        btn.Click       += on_click
        return btn

    for name, filename in GAMES:
        root.Children.Add(make_btn(name, filename))

    # Random button at bottom
    sep = _font(11)
    sep.Text       = "-" * 38
    sep.Foreground = _brush(0, 60, 20)
    sep.Margin     = Thickness(0, 10, 0, 4)
    root.Children.Add(sep)

    rnd_btn = Button()
    rnd_btn.Background          = _brush(0, 40, 8)
    rnd_btn.BorderBrush         = _brush(0, 180, 50)
    rnd_btn.BorderThickness     = Thickness(1)
    rnd_btn.Margin              = Thickness(0, 3, 0, 3)
    rnd_btn.Padding             = Thickness(10, 6, 10, 6)
    rnd_btn.HorizontalContentAlignment = HorizontalAlignment.Left
    rnd_tb = _font(13, True)
    rnd_tb.Text       = ">  ** RANDOM SURPRISE **"
    rnd_tb.Foreground = _brush(0, 255, 65)
    rnd_btn.Content   = rnd_tb

    def on_rnd_hover_enter(s, e):
        s.Background   = _brush(0, 80, 15)
        rnd_tb.Foreground = _brush(255, 220, 0)

    def on_rnd_hover_leave(s, e):
        s.Background   = _brush(0, 40, 8)
        rnd_tb.Foreground = _brush(0, 255, 65)

    def on_rnd_click(s, e):
        result[0] = random.choice(GAMES)
        win.Close()

    rnd_btn.MouseEnter  += on_rnd_hover_enter
    rnd_btn.MouseLeave  += on_rnd_hover_leave
    rnd_btn.Click       += on_rnd_click
    root.Children.Add(rnd_btn)

    win.ShowDialog()
    return result[0]  # None if window closed without picking

# ---------------------------------------------------------------------------
# Load XAML game window
# ---------------------------------------------------------------------------
def load_window():
    reader   = StreamReader(XAML_PATH)
    xaml_txt = reader.ReadToEnd()
    reader.Close()
    xml_reader = XmlReader.Create(StringReader(xaml_txt))
    return XamlReader.Load(xml_reader)

# ---------------------------------------------------------------------------
# Context object passed to each game module
# ---------------------------------------------------------------------------
class GameContext(object):
    def __init__(self, window, canvas, status_label, score_label, game_title):
        self.window      = window
        self.canvas      = canvas
        self.status      = status_label
        self.score_label = score_label
        self.title_label = game_title
        self.score       = 0
        self.games_dir   = GAMES_DIR

    def set_status(self, text):
        self.status.Text = text

    def set_score(self, value):
        self.score = value
        self.score_label.Text = "SCORE: {}".format(value)

    def set_title(self, text):
        self.title_label.Text = text

# ---------------------------------------------------------------------------
# Countdown timer
# ---------------------------------------------------------------------------
class BreakTimer(object):
    def __init__(self, window, label, on_expire):
        self.window    = window
        self.label     = label
        self.on_expire = on_expire
        self.remaining = TOTAL_SECONDS
        self._timer    = DispatcherTimer()
        self._timer.Interval = TimeSpan.FromSeconds(1)
        self._timer.Tick += self._tick

    def start(self):
        self._update_label()
        self._timer.Start()

    def stop(self):
        self._timer.Stop()

    def _tick(self, sender, e):
        self.remaining -= 1
        self._update_label()
        if self.remaining <= 0:
            self._timer.Stop()
            self.on_expire()

    def _update_label(self):
        mins = self.remaining // 60
        secs = self.remaining % 60
        self.label.Text = "{}:{:02d}".format(mins, secs)
        if self.remaining <= 30:
            self.label.Foreground = SolidColorBrush(Color.FromRgb(255, 51, 0))
        elif self.remaining <= 60:
            self.label.Foreground = SolidColorBrush(Color.FromRgb(255, 170, 0))

# ---------------------------------------------------------------------------
# Load game module via execfile
# ---------------------------------------------------------------------------
def load_game_module(filename):
    path = os.path.join(GAMES_DIR, filename)
    ns   = {}
    execfile(path, ns)
    return ns

# ---------------------------------------------------------------------------
# Main entry
# ---------------------------------------------------------------------------
def main():
    # Show picker first
    choice = show_picker()
    if choice is None:
        return  # user closed without picking

    game_name, game_file = choice

    window    = load_window()
    canvas    = window.FindName("GameCanvas")
    countdown = window.FindName("CountdownLabel")
    status    = window.FindName("StatusLabel")
    score_lbl = window.FindName("ScoreLabel")
    title_lbl = window.FindName("GameTitle")

    title_lbl.Text = ">> {} <<".format(game_name)
    window.Title   = "SEED43 BREAK TIME  |  {}".format(game_name)

    ctx = GameContext(window, canvas, status, score_lbl, title_lbl)

    def on_time_up():
        ctx.set_status("TIME'S UP!  GET BACK TO WORK!")
        grace = DispatcherTimer()
        grace.Interval = TimeSpan.FromSeconds(2)
        def _close(s, e):
            grace.Stop()
            window.Close()
        grace.Tick += _close
        grace.Start()

    timer = BreakTimer(window, countdown, on_time_up)

    try:
        game_ns = load_game_module(game_file)
        init_fn = game_ns.get("init_game")
        if init_fn:
            init_fn(ctx, timer)
    except Exception as ex:
        ctx.set_status("ERROR LOADING GAME: {}".format(str(ex)))

    timer.start()
    window.ShowDialog()

main()
