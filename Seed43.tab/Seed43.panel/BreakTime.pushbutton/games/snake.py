# -*- coding: utf-8 -*-
# Snake - BreakTime game module
# Arrow keys to steer. Eat the dots. Don't hit yourself or the wall.

from System.Windows.Threading import DispatcherTimer
from System.Windows.Input import Key
from System.Windows.Shapes import Rectangle, Ellipse
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Controls import TextBlock, Canvas as WpfCanvas
from System.Windows import Thickness, FontWeight, FontWeights
from System import TimeSpan
import random

CELL  = 20
COLS  = 37
ROWS  = 26
DOS_GREEN  = Color.FromRgb(0,   255, 65)
DOS_DIM    = Color.FromRgb(0,   100, 20)
DOS_FOOD   = Color.FromRgb(255, 255, 0)
DOS_HEAD   = Color.FromRgb(0,   255, 65)
DOS_RED    = Color.FromRgb(255, 51,  0)

def _brush(color):
    return SolidColorBrush(color)

def _rect(x, y, w, h, color):
    r = Rectangle()
    r.Width  = w
    r.Height = h
    r.Fill   = _brush(color)
    WpfCanvas.SetLeft(r, x)
    WpfCanvas.SetTop(r, y)
    return r

def init_game(ctx, timer):
    canvas = ctx.canvas
    canvas.Width  = COLS * CELL
    canvas.Height = ROWS * CELL

    state = {
        "snake":     [(COLS//2, ROWS//2), (COLS//2-1, ROWS//2), (COLS//2-2, ROWS//2)],
        "direction": (1, 0),
        "next_dir":  (1, 0),
        "food":      None,
        "rects":     {},
        "food_el":   None,
        "alive":     True,
        "started":   False,
        "score":     0,
    }

    # Draw border
    border = _rect(0, 0, COLS*CELL, ROWS*CELL, Color.FromRgb(0, 40, 0))
    canvas.Children.Add(border)
    for i in range(COLS):
        canvas.Children.Add(_rect(i*CELL, 0,           CELL, CELL, DOS_DIM))
        canvas.Children.Add(_rect(i*CELL, (ROWS-1)*CELL, CELL, CELL, DOS_DIM))
    for j in range(ROWS):
        canvas.Children.Add(_rect(0,           j*CELL, CELL, CELL, DOS_DIM))
        canvas.Children.Add(_rect((COLS-1)*CELL, j*CELL, CELL, CELL, DOS_DIM))

    def place_food():
        while True:
            fx = random.randint(1, COLS-2)
            fy = random.randint(1, ROWS-2)
            if (fx, fy) not in state["snake"]:
                break
        if state["food_el"] and state["food_el"] in canvas.Children:
            canvas.Children.Remove(state["food_el"])
        el = _rect(fx*CELL+2, fy*CELL+2, CELL-4, CELL-4, DOS_FOOD)
        canvas.Children.Add(el)
        state["food"]    = (fx, fy)
        state["food_el"] = el

    def draw_snake():
        for seg, rect in list(state["rects"].items()):
            if rect in canvas.Children:
                canvas.Children.Remove(rect)
        state["rects"] = {}
        for i, seg in enumerate(state["snake"]):
            c = DOS_HEAD if i == 0 else DOS_GREEN
            if i > 0:
                c = Color.FromRgb(0, max(80, 200 - i*3), max(10, 40 - i))
            r = _rect(seg[0]*CELL+1, seg[1]*CELL+1, CELL-2, CELL-2, c)
            canvas.Children.Add(r)
            state["rects"][seg] = r

    def step(sender, e):
        if not state["alive"] or not state["started"]:
            return
        dx, dy = state["next_dir"]
        state["direction"] = (dx, dy)
        hx, hy = state["snake"][0]
        nx, ny = hx+dx, hy+dy

        # Wall collision
        if nx <= 0 or nx >= COLS-1 or ny <= 0 or ny >= ROWS-1:
            game_over()
            return
        # Self collision
        if (nx, ny) in state["snake"]:
            game_over()
            return

        state["snake"].insert(0, (nx, ny))
        if (nx, ny) == state["food"]:
            state["score"] += 10
            ctx.set_score(state["score"])
            place_food()
        else:
            state["snake"].pop()

        draw_snake()

    def game_over():
        state["alive"] = False
        loop.Stop()
        ctx.set_status("GAME OVER!  SCORE: {}  |  PRESS R TO RESTART".format(state["score"]))
        # Flash red
        for rect in state["rects"].values():
            rect.Fill = _brush(DOS_RED)

    def on_key(sender, e):
        k = e.Key
        dx, dy = state["direction"]
        if k == Key.Up    and dy == 0: state["next_dir"] = (0, -1)
        if k == Key.Down  and dy == 0: state["next_dir"] = (0,  1)
        if k == Key.Left  and dx == 0: state["next_dir"] = (-1, 0)
        if k == Key.Right and dx == 0: state["next_dir"] = ( 1, 0)
        if k == Key.R and not state["alive"]:
            restart()
        if not state["started"]:
            state["started"] = True
            ctx.set_status("SCORE: {}  |  ARROW KEYS TO STEER".format(state["score"]))
            loop.Start()

    def restart():
        canvas.Children.Clear()
        state["snake"]     = [(COLS//2, ROWS//2), (COLS//2-1, ROWS//2), (COLS//2-2, ROWS//2)]
        state["direction"] = (1, 0)
        state["next_dir"]  = (1, 0)
        state["food"]      = None
        state["food_el"]   = None
        state["rects"]     = {}
        state["alive"]     = True
        state["started"]   = False
        state["score"]     = 0
        ctx.set_score(0)
        # Redraw border
        border2 = _rect(0, 0, COLS*CELL, ROWS*CELL, Color.FromRgb(0, 40, 0))
        canvas.Children.Add(border2)
        for i in range(COLS):
            canvas.Children.Add(_rect(i*CELL, 0,           CELL, CELL, DOS_DIM))
            canvas.Children.Add(_rect(i*CELL, (ROWS-1)*CELL, CELL, CELL, DOS_DIM))
        for j in range(ROWS):
            canvas.Children.Add(_rect(0,           j*CELL, CELL, CELL, DOS_DIM))
            canvas.Children.Add(_rect((COLS-1)*CELL, j*CELL, CELL, CELL, DOS_DIM))
        draw_snake()
        place_food()
        ctx.set_status("PRESS ANY KEY TO START")

    loop = DispatcherTimer()
    loop.Interval = TimeSpan.FromMilliseconds(130)
    loop.Tick += step

    canvas.KeyDown += on_key
    canvas.Focusable = True
    canvas.Focus()

    place_food()
    draw_snake()
    ctx.set_status("PRESS ANY KEY TO START  |  ARROW KEYS TO STEER")
    ctx.set_title(">> SNAKE <<")
