# -*- coding: utf-8 -*-
# Pac-Man lite - BreakTime game module
# Simple maze, dots, power pellets, 2 ghosts.

from System.Windows.Threading import DispatcherTimer
from System.Windows.Input import Key
from System.Windows.Shapes import Rectangle, Ellipse
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Controls import Canvas as WpfCanvas
from System import TimeSpan
import random

CELL = 24
C_BG    = Color.FromRgb(0,  15,  0)
C_WALL  = Color.FromRgb(0,  100, 20)
C_DOT   = Color.FromRgb(0,  200, 60)
C_PPEL  = Color.FromRgb(255,220, 0)
C_PAC   = Color.FromRgb(255,220, 0)
C_G1    = Color.FromRgb(255,  51, 0)
C_G2    = Color.FromRgb(0,   180,255)
C_GFRIT = Color.FromRgb(100, 100,255)

# 0=path, 1=wall, 2=dot, 3=power, 4=ghost spawn, 5=pac spawn
MAZE_TEMPLATE = [
    "1111111111111111111111111111111",
    "1222222222222222222222222222221",
    "1211121112111111111211121112121",
    "1311121112111111111211121112131",
    "1211121112111111111211121112121",
    "1222222222222222222222222222221",
    "1211121122111211112211122112121",
    "1211121122111211112211122112121",
    "1222222002222202222200222222221",
    "1111121102111000111201111211111",
    "1111121100111444111001111211111",
    "1111121100140444041001111211111",
    "0000001000440000440001000000000",
    "1111121100140444041001111211111",
    "1111121100111000111001111211111",
    "1111121100111111111001111211111",
    "1222222222222222222222222222221",
    "1211121112111111111211121112121",
    "1311121112111111111211121112131",
    "1221122222222225222222222211221",
    "1211211112111111111211111121121",
    "1222222222222222222222222222221",
    "1111111111111111111111111111111",
]

ROWS = len(MAZE_TEMPLATE)
COLS = len(MAZE_TEMPLATE[0])

def _brush(c): return SolidColorBrush(c)
def _rect(canvas, x, y, w, h, color, z=0):
    r = Rectangle(); r.Width=w; r.Height=h; r.Fill=_brush(color)
    WpfCanvas.SetLeft(r,x); WpfCanvas.SetTop(r,y); WpfCanvas.SetZIndex(r,z)
    canvas.Children.Add(r); return r
def _ellipse(canvas, cx, cy, rx, ry, color, z=1):
    e = Ellipse(); e.Width=rx*2; e.Height=ry*2; e.Fill=_brush(color)
    WpfCanvas.SetLeft(e, cx-rx); WpfCanvas.SetTop(e, cy-ry)
    WpfCanvas.SetZIndex(e, z); canvas.Children.Add(e); return e

def init_game(ctx, timer):
    canvas = ctx.canvas
    canvas.Width  = COLS * CELL
    canvas.Height = ROWS * CELL

    maze = [list(row) for row in MAZE_TEMPLATE]
    dot_els  = {}   # (r,c) -> el
    pac_spawn = None; ghost_spawns = []

    for r in range(ROWS):
        for c in range(COLS):
            ch = maze[r][c]
            x = c*CELL; y = r*CELL
            if ch == "1":
                _rect(canvas, x, y, CELL, CELL, C_WALL)
            if ch == "5":
                pac_spawn = (r, c); maze[r][c] = "0"
            if ch == "4":
                ghost_spawns.append((r,c)); maze[r][c] = "0"

    def reset_dots():
        for el in list(dot_els.values()):
            if el in canvas.Children:
                canvas.Children.Remove(el)
        dot_els.clear()
        for r in range(ROWS):
            for c in range(COLS):
                ch = MAZE_TEMPLATE[r][c]
                if ch == "2":
                    el = _ellipse(canvas, c*CELL+CELL//2, r*CELL+CELL//2, 3, 3, C_DOT, 1)
                    dot_els[(r,c)] = el
                    maze[r][c] = "2"
                elif ch == "3":
                    el = _ellipse(canvas, c*CELL+CELL//2, r*CELL+CELL//2, 7, 7, C_PPEL, 1)
                    dot_els[(r,c)] = el
                    maze[r][c] = "3"

    reset_dots()

    pr, pc = pac_spawn or (19, 15)
    state = {
        "pr": float(pr), "pc": float(pc),
        "pd": (0,0), "pnd": (0,0),
        "score": 0, "lives": 3, "alive": True, "started": False,
        "fright": 0,
        "ghosts": [],
    }

    pac_el = _ellipse(canvas, pc*CELL+CELL//2, pr*CELL+CELL//2, CELL//2-1, CELL//2-1, C_PAC, 3)
    state["pac_el"] = pac_el

    # Spawn ghosts
    ghost_colors = [C_G1, C_G2]
    ghost_els = []
    for i, (gr, gc) in enumerate(ghost_spawns[:2]):
        el = _ellipse(canvas, gc*CELL+CELL//2, gr*CELL+CELL//2, CELL//2-1, CELL//2-1, ghost_colors[i%2], 2)
        state["ghosts"].append({"r": float(gr), "c": float(gc), "dr": 0, "dc": 0, "color": ghost_colors[i%2], "el": el})
        ghost_els.append(el)

    life_els = [_rect(canvas, 10+i*20, ROWS*CELL-18, 14, 8, C_PAC) for i in range(3)]

    def is_open(r, c):
        ri = int(round(r)); ci = int(round(c))
        ri = ri % ROWS; ci = ci % COLS
        return MAZE_TEMPLATE[ri][ci] != "1"

    def move_pac():
        nr = state["pr"] + state["pnd"][0] * 0.15
        nc = state["pc"] + state["pnd"][1] * 0.15
        if is_open(nr, nc):
            state["pd"] = state["pnd"]
        nr2 = state["pr"] + state["pd"][0] * 0.15
        nc2 = state["pc"] + state["pd"][1] * 0.15
        if is_open(nr2, nc2):
            state["pr"] = nr2 % ROWS
            state["pc"] = nc2 % COLS

        WpfCanvas.SetLeft(pac_el, state["pc"]*CELL+2)
        WpfCanvas.SetTop(pac_el,  state["pr"]*CELL+2)

        # Eat dot
        ri = int(round(state["pr"])); ci = int(round(state["pc"]))
        if (ri,ci) in dot_els:
            is_power = MAZE_TEMPLATE[ri][ci] == "3"
            canvas.Children.Remove(dot_els.pop((ri,ci)))
            maze[ri][ci] = "0"
            state["score"] += 50 if is_power else 10
            ctx.set_score(state["score"])
            if is_power:
                state["fright"] = 60
            if not dot_els:
                reset_dots()
                ctx.set_status("LEVEL CLEAR!  SCORE: {}".format(state["score"]))

    def move_ghost(g):
        SPEED = 0.12
        ri = int(round(g["r"])); ci = int(round(g["c"]))
        if g["dr"] == 0 and g["dc"] == 0:
            g["dr"] = 1; g["dc"] = 0
        dirs = [(0,1),(0,-1),(1,0),(-1,0)]
        random.shuffle(dirs)
        if state["fright"] > 0:
            best = random.choice(dirs)
        else:
            pr2 = state["pr"]; pc2 = state["pc"]
            best = min(dirs, key=lambda d: (ri+d[0]-pr2)**2+(ci+d[1]-pc2)**2)
        for dr,dc in [best]+dirs:
            if (dr,dc) == (-g["dr"],-g["dc"]): continue
            nr = g["r"]+dr*SPEED; nc = g["c"]+dc*SPEED
            if is_open(nr, nc):
                g["r"] = nr % ROWS; g["c"] = nc % COLS
                g["dr"] = dr; g["dc"] = dc; break
        c = C_GFRIT if state["fright"] > 0 else g["color"]
        g["el"].Fill = _brush(c)
        WpfCanvas.SetLeft(g["el"], g["c"]*CELL+2)
        WpfCanvas.SetTop(g["el"],  g["r"]*CELL+2)

    def check_ghost_collision():
        for g in state["ghosts"]:
            if abs(g["r"]-state["pr"]) < 0.8 and abs(g["c"]-state["pc"]) < 0.8:
                if state["fright"] > 0:
                    gr, gc = ghost_spawns[0] if ghost_spawns else (11,15)
                    g["r"] = float(gr); g["c"] = float(gc)
                    state["score"] += 200; ctx.set_score(state["score"])
                else:
                    state["lives"] -= 1
                    if state["lives"] >= 0 and state["lives"] < len(life_els):
                        life_els[state["lives"]].Fill = _brush(Color.FromRgb(0,40,0))
                    if state["lives"] <= 0:
                        state["alive"] = False
                        loop.Stop()
                        ctx.set_status("GAME OVER!  SCORE: {}".format(state["score"]))
                    else:
                        state["pr"] = float(pr); state["pc"] = float(pc)
                        ctx.set_status("OUCH!  LIVES: {}".format(state["lives"]))

    def tick(sender, e):
        if not state["alive"] or not state["started"]:
            return
        move_pac()
        for g in state["ghosts"]:
            move_ghost(g)
        check_ghost_collision()
        if state["fright"] > 0:
            state["fright"] -= 1

    def on_key_down(sender, e):
        k = e.Key
        if k == Key.Up:    state["pnd"] = (-1, 0)
        if k == Key.Down:  state["pnd"] = ( 1, 0)
        if k == Key.Left:  state["pnd"] = ( 0,-1)
        if k == Key.Right: state["pnd"] = ( 0, 1)
        if not state["started"]:
            state["started"] = True
            loop.Start()
            ctx.set_status("ARROWS: MOVE  |  EAT DOTS  |  BIG DOTS = EAT GHOSTS!")

    loop = DispatcherTimer()
    loop.Interval = TimeSpan.FromMilliseconds(33)
    loop.Tick += tick

    canvas.KeyDown += on_key_down
    canvas.Focusable = True; canvas.Focus()
    ctx.set_status("PRESS ANY KEY TO START  |  ARROWS TO MOVE")
    ctx.set_title(">> PAC-MAN <<")
