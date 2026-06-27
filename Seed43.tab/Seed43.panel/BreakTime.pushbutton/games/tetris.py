# -*- coding: utf-8 -*-
# Tetris - BreakTime game module

from System.Windows.Threading import DispatcherTimer
from System.Windows.Input import Key
from System.Windows.Shapes import Rectangle
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Controls import TextBlock, Canvas as WpfCanvas
from System.Windows import Thickness, FontWeights
from System import TimeSpan
import random

COLS = 12
ROWS = 22
CELL = 24

C_BG      = Color.FromRgb(0,  15,  0)
C_BORDER  = Color.FromRgb(0,  80,  20)
C_GHOST   = Color.FromRgb(0,  60,  0)

PIECES = {
    "I": (Color.FromRgb(0,220,200), [[(0,0),(1,0),(2,0),(3,0)],[(0,0),(0,1),(0,2),(0,3)]]),
    "O": (Color.FromRgb(255,220,0), [[(0,0),(1,0),(0,1),(1,1)]]),
    "T": (Color.FromRgb(180,0,200), [[(0,1),(1,0),(1,1),(2,1)],[(0,0),(0,1),(0,2),(1,1)],[(0,0),(1,0),(2,0),(1,1)],[(0,1),(1,0),(1,1),(1,2)]]),
    "S": (Color.FromRgb(0,220,80),  [[(0,1),(1,0),(1,1),(2,0)],[(0,0),(0,1),(1,1),(1,2)]]),
    "Z": (Color.FromRgb(220,50,0),  [[(0,0),(1,0),(1,1),(2,1)],[(0,1),(0,2),(1,0),(1,1)]]),
    "J": (Color.FromRgb(0,80,220),  [[(0,0),(0,1),(1,1),(2,1)],[(0,0),(1,0),(0,1),(0,2)],[(0,0),(1,0),(2,0),(2,1)],[(0,2),(1,0),(1,1),(1,2)]]),
    "L": (Color.FromRgb(255,140,0), [[(2,0),(0,1),(1,1),(2,1)],[(0,0),(0,1),(0,2),(1,0)],[(0,0),(1,0),(2,0),(0,1)],[(0,2),(1,0),(1,1),(1,2)]]),
}
PIECE_KEYS = list(PIECES.keys())

def _brush(c):
    return SolidColorBrush(c)

def _rect(canvas, x, y, w, h, color, stroke=None, z=0):
    r = Rectangle()
    r.Width = w; r.Height = h
    r.Fill  = _brush(color)
    if stroke:
        r.Stroke = _brush(stroke)
        r.StrokeThickness = 1
    WpfCanvas.SetLeft(r, x); WpfCanvas.SetTop(r, y)
    WpfCanvas.SetZIndex(r, z)
    canvas.Children.Add(r)
    return r

def init_game(ctx, timer):
    canvas = ctx.canvas
    canvas.Width  = COLS * CELL + 160
    canvas.Height = ROWS * CELL + 2

    grid = [[None]*COLS for _ in range(ROWS)]   # color or None
    grid_rects = [[None]*COLS for _ in range(ROWS)]

    state = {
        "piece":      None,
        "piece_x":    0,
        "piece_y":    0,
        "rot":        0,
        "bag":        [],
        "next_piece": None,
        "score":      0,
        "lines":      0,
        "level":      1,
        "alive":      True,
        "started":    False,
    }

    # Background
    _rect(canvas, 0, 0, COLS*CELL, ROWS*CELL, C_BG)
    for r in range(ROWS):
        for c in range(COLS):
            _rect(canvas, c*CELL, r*CELL, CELL-1, CELL-1, C_BG, C_BORDER)

    # Side panel
    px = COLS*CELL + 10
    from System.Windows.Media import FontFamily
    def _lbl(x, y, text, size=12, color=Color.FromRgb(0,180,40)):
        tb = TextBlock()
        tb.Text = text; tb.FontFamily = FontFamily("Courier New")
        tb.FontSize = size; tb.Foreground = _brush(color)
        WpfCanvas.SetLeft(tb, x); WpfCanvas.SetTop(tb, y)
        canvas.Children.Add(tb)
        return tb

    _lbl(px, 20,  "NEXT",  11)
    _lbl(px, 140, "SCORE", 11)
    state["score_lbl"] = _lbl(px, 158, "0", 14, Color.FromRgb(0,255,65))
    _lbl(px, 195, "LINES", 11)
    state["lines_lbl"] = _lbl(px, 213, "0", 14, Color.FromRgb(0,255,65))
    _lbl(px, 250, "LEVEL", 11)
    state["level_lbl"] = _lbl(px, 268, "1", 14, Color.FromRgb(0,255,65))
    state["next_rects"] = []

    def next_from_bag():
        if not state["bag"]:
            state["bag"] = random.sample(PIECE_KEYS, len(PIECE_KEYS))
        return state["bag"].pop()

    def spawn():
        key = state["next_piece"] or next_from_bag()
        state["next_piece"] = next_from_bag()
        color, rotations = PIECES[key]
        state["piece"] = (key, color, rotations)
        state["piece_x"] = COLS//2 - 2
        state["piece_y"] = 0
        state["rot"] = 0
        draw_next()
        if not can_place(state["piece_x"], state["piece_y"], state["rot"]):
            state["alive"] = False
            loop.Stop()
            ctx.set_status("GAME OVER!  SCORE: {}  LINES: {}".format(state["score"], state["lines"]))

    def cells_for(px2, py2, rot):
        _, _, rotations = state["piece"]
        rot = rot % len(rotations)
        return [(px2+dc, py2+dr) for (dc,dr) in rotations[rot]]

    def can_place(px2, py2, rot):
        for (cx2, cy2) in cells_for(px2, py2, rot):
            if cx2 < 0 or cx2 >= COLS or cy2 >= ROWS:
                return False
            if cy2 >= 0 and grid[cy2][cx2] is not None:
                return False
        return True

    def lock_piece():
        _, color, _ = state["piece"]
        for (cx2, cy2) in cells_for(state["piece_x"], state["piece_y"], state["rot"]):
            if cy2 >= 0:
                grid[cy2][cx2] = color
        clear_lines()
        spawn()

    def clear_lines():
        full = [r for r in range(ROWS) if all(grid[r][c] is not None for c in range(COLS))]
        for r in full:
            for rr in range(r, 0, -1):
                grid[rr] = list(grid[rr-1])
            grid[0] = [None]*COLS
        pts = [0, 100, 300, 500, 800][min(len(full), 4)]
        state["score"] += pts * state["level"]
        state["lines"] += len(full)
        state["level"]  = state["lines"] // 10 + 1
        ctx.set_score(state["score"])
        state["score_lbl"].Text = str(state["score"])
        state["lines_lbl"].Text = str(state["lines"])
        state["level_lbl"].Text = str(state["level"])
        loop.Interval = TimeSpan.FromMilliseconds(max(80, 500 - state["level"]*40))
        redraw_grid()

    def redraw_grid():
        for r in range(ROWS):
            for c in range(COLS):
                el = grid_rects[r][c]
                color = grid[r][c]
                if color:
                    if el is None:
                        el = _rect(canvas, c*CELL+1, r*CELL+1, CELL-2, CELL-2, color, None, 1)
                        grid_rects[r][c] = el
                    else:
                        el.Fill = _brush(color)
                        WpfCanvas.SetLeft(el, c*CELL+1); WpfCanvas.SetTop(el, r*CELL+1)
                else:
                    if el is not None:
                        canvas.Children.Remove(el)
                        grid_rects[r][c] = None

    piece_rects = []
    def draw_piece():
        for el in piece_rects:
            if el in canvas.Children:
                canvas.Children.Remove(el)
        piece_rects[:] = []
        if not state["piece"] or not state["alive"]:
            return
        _, color, _ = state["piece"]
        # Ghost
        gy = state["piece_y"]
        while can_place(state["piece_x"], gy+1, state["rot"]):
            gy += 1
        for (cx2, cy2) in cells_for(state["piece_x"], gy, state["rot"]):
            if cy2 >= 0:
                el = _rect(canvas, cx2*CELL+1, cy2*CELL+1, CELL-2, CELL-2, C_GHOST, None, 1)
                piece_rects.append(el)
        for (cx2, cy2) in cells_for(state["piece_x"], state["piece_y"], state["rot"]):
            if cy2 >= 0:
                el = _rect(canvas, cx2*CELL+1, cy2*CELL+1, CELL-2, CELL-2, color, None, 2)
                piece_rects.append(el)

    def draw_next():
        for el in state["next_rects"]:
            if el in canvas.Children:
                canvas.Children.Remove(el)
        state["next_rects"][:] = []
        if not state["next_piece"]:
            return
        ncolor, nrots = PIECES[state["next_piece"]]
        for (dc, dr) in nrots[0]:
            el = _rect(canvas, px+dc*CELL+5, 40+dr*CELL, CELL-2, CELL-2, ncolor, None, 1)
            state["next_rects"].append(el)

    def drop_tick(sender, e):
        if not state["alive"] or not state["started"]:
            return
        if can_place(state["piece_x"], state["piece_y"]+1, state["rot"]):
            state["piece_y"] += 1
        else:
            lock_piece()
        draw_piece()

    def on_key(sender, e):
        k = e.Key
        if not state["started"]:
            state["started"] = True
            ctx.set_status("LEFT/RIGHT: MOVE  |  DOWN: FAST  |  SPACE: ROTATE  |  UP: DROP")
            loop.Start()
            return
        if not state["alive"]:
            return
        if k == Key.Left:
            if can_place(state["piece_x"]-1, state["piece_y"], state["rot"]):
                state["piece_x"] -= 1
        elif k == Key.Right:
            if can_place(state["piece_x"]+1, state["piece_y"], state["rot"]):
                state["piece_x"] += 1
        elif k == Key.Down:
            if can_place(state["piece_x"], state["piece_y"]+1, state["rot"]):
                state["piece_y"] += 1
                state["score"] += 1
        elif k == Key.Space:
            nr = (state["rot"] + 1) % len(state["piece"][2])
            if can_place(state["piece_x"], state["piece_y"], nr):
                state["rot"] = nr
        elif k == Key.Up:
            while can_place(state["piece_x"], state["piece_y"]+1, state["rot"]):
                state["piece_y"] += 1
                state["score"] += 2
            lock_piece()
        draw_piece()

    loop = DispatcherTimer()
    loop.Interval = TimeSpan.FromMilliseconds(500)
    loop.Tick += drop_tick

    canvas.KeyDown += on_key
    canvas.Focusable = True
    canvas.Focus()

    spawn()
    draw_piece()
    ctx.set_status("PRESS ANY KEY TO START")
    ctx.set_title(">> TETRIS <<")
