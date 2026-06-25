# -*- coding: utf-8 -*-
# Breakout - BreakTime game module

from System.Windows.Threading import DispatcherTimer
from System.Windows.Input import Key
from System.Windows.Shapes import Rectangle, Ellipse
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Controls import TextBlock, Canvas as WpfCanvas
from System.Windows import FontWeights
from System import TimeSpan
import random, math

W = 740; H = 520
PAD_W = 90; PAD_H = 12; PAD_Y = H - 50
BALL_R = 7
BRICK_COLS = 14; BRICK_ROWS = 8
BRICK_W = 46; BRICK_H = 18; BRICK_PAD = 4
BRICK_TOP = 50

C_BG  = Color.FromRgb(0, 15, 0)
C_PAD = Color.FromRgb(0, 255, 65)
C_BALL= Color.FromRgb(255,220,0)
C_DIM = Color.FromRgb(0, 60, 20)

ROW_COLORS = [
    Color.FromRgb(255,  80,  0),
    Color.FromRgb(255, 160,  0),
    Color.FromRgb(255, 220,  0),
    Color.FromRgb(0,   220, 80),
    Color.FromRgb(0,   200,200),
    Color.FromRgb(0,   120,255),
    Color.FromRgb(160,  0, 255),
    Color.FromRgb(200,  0, 200),
]

def _brush(c): return SolidColorBrush(c)

def _rect(canvas, x, y, w, h, color, z=0):
    r = Rectangle(); r.Width=w; r.Height=h; r.Fill=_brush(color)
    WpfCanvas.SetLeft(r,x); WpfCanvas.SetTop(r,y); WpfCanvas.SetZIndex(r,z)
    canvas.Children.Add(r); return r

def init_game(ctx, timer):
    canvas = ctx.canvas
    canvas.Width=W; canvas.Height=H
    _rect(canvas,0,0,W,H,C_BG)

    state = {
        "px": float(W//2 - PAD_W//2),
        "bx": float(W//2), "by": float(PAD_Y - BALL_R - 2),
        "vx": 4.0, "vy": -5.0,
        "bricks": {},  # (r,c) -> rect or None
        "alive": True, "started": False,
        "left": False, "right": False,
        "score": 0, "lives": 3,
        "attached": True,
    }

    pad_el = _rect(canvas, state["px"], PAD_Y, PAD_W, PAD_H, C_PAD, 2)
    ball_el= _rect(canvas, state["bx"]-BALL_R, state["by"]-BALL_R, BALL_R*2, BALL_R*2, C_BALL, 2)

    # Lives
    life_els = [_rect(canvas, 10+i*20, H-22, 14, 8, C_PAD) for i in range(3)]
    state["life_els"] = life_els

    def spawn_bricks():
        for (r,c), el in list(state["bricks"].items()):
            if el and el in canvas.Children:
                canvas.Children.Remove(el)
        state["bricks"] = {}
        for r in range(BRICK_ROWS):
            for c in range(BRICK_COLS):
                bx2 = 20 + c * (BRICK_W + BRICK_PAD)
                by2 = BRICK_TOP + r * (BRICK_H + BRICK_PAD)
                el  = _rect(canvas, bx2, by2, BRICK_W-2, BRICK_H-2, ROW_COLORS[r], 1)
                state["bricks"][(r,c)] = el

    spawn_bricks()

    def update_visuals():
        WpfCanvas.SetLeft(pad_el,  state["px"])
        WpfCanvas.SetLeft(ball_el, state["bx"]-BALL_R)
        WpfCanvas.SetTop(ball_el,  state["by"]-BALL_R)

    def tick(sender, e):
        if not state["alive"] or not state["started"]:
            return
        spd = 5
        if state["left"]:  state["px"] = max(0, state["px"] - spd)
        if state["right"]: state["px"] = min(W-PAD_W, state["px"] + spd)

        if state["attached"]:
            state["bx"] = state["px"] + PAD_W//2
            update_visuals()
            return

        state["bx"] += state["vx"]
        state["by"] += state["vy"]

        # Wall bounces
        if state["bx"] - BALL_R <= 0:
            state["bx"] = BALL_R; state["vx"] = abs(state["vx"])
        if state["bx"] + BALL_R >= W:
            state["bx"] = W-BALL_R; state["vx"] = -abs(state["vx"])
        if state["by"] - BALL_R <= 0:
            state["by"] = BALL_R; state["vy"] = abs(state["vy"])

        # Paddle
        if (state["by"]+BALL_R >= PAD_Y and
            state["by"]+BALL_R <= PAD_Y+PAD_H+8 and
            state["bx"] >= state["px"] and
            state["bx"] <= state["px"]+PAD_W):
            state["by"] = PAD_Y - BALL_R
            rel = (state["bx"] - (state["px"]+PAD_W/2)) / (PAD_W/2)
            speed = math.sqrt(state["vx"]**2 + state["vy"]**2)
            angle = rel * 1.2
            state["vx"] = math.sin(angle) * speed
            state["vy"] = -abs(math.cos(angle) * speed)

        # Miss
        if state["by"] > H + 20:
            state["lives"] -= 1
            if state["lives"] >= 0 and state["lives"] < len(life_els):
                life_els[state["lives"]].Fill = _brush(C_DIM)
            if state["lives"] <= 0:
                state["alive"] = False
                loop.Stop()
                ctx.set_status("GAME OVER!  SCORE: {}".format(state["score"]))
                return
            state["attached"] = True
            state["bx"] = state["px"] + PAD_W//2
            state["by"] = PAD_Y - BALL_R - 2
            ctx.set_status("PRESS SPACE TO LAUNCH  |  LIVES: {}".format(state["lives"]))

        # Brick hits
        for (r,c), el in list(state["bricks"].items()):
            if el is None: continue
            bx2 = WpfCanvas.GetLeft(el); by2 = WpfCanvas.GetTop(el)
            if (bx2-BALL_R <= state["bx"] <= bx2+BRICK_W+BALL_R and
                by2-BALL_R <= state["by"] <= by2+BRICK_H+BALL_R):
                canvas.Children.Remove(el)
                state["bricks"][(r,c)] = None
                pts = (BRICK_ROWS - r) * 10
                state["score"] += pts
                ctx.set_score(state["score"])
                # Bounce direction based on entry
                mid_x = bx2 + BRICK_W/2; mid_y = by2 + BRICK_H/2
                if abs(state["bx"]-mid_x)/BRICK_W > abs(state["by"]-mid_y)/BRICK_H:
                    state["vx"] *= -1
                else:
                    state["vy"] *= -1
                # Small speed increase
                spd2 = math.sqrt(state["vx"]**2 + state["vy"]**2)
                if spd2 < 9:
                    state["vx"] *= 1.01; state["vy"] *= 1.01
                break

        # All clear
        if all(el is None for el in state["bricks"].values()):
            spawn_bricks()
            state["attached"] = True
            ctx.set_status("CLEARED!  PRESS SPACE FOR NEXT WAVE!")

        update_visuals()

    def on_key_down(sender, e):
        if e.Key == Key.Left:  state["left"]  = True
        if e.Key == Key.Right: state["right"] = True
        if e.Key == Key.Space:
            if state["attached"]:
                state["attached"] = False
                ang = random.uniform(-0.3, 0.3)
                state["vx"] = math.sin(ang) * 6
                state["vy"] = -math.cos(ang) * 6
        if not state["started"]:
            state["started"] = True
            loop.Start()
            ctx.set_status("ARROWS: MOVE  |  SPACE: LAUNCH")

    def on_key_up(sender, e):
        if e.Key == Key.Left:  state["left"]  = False
        if e.Key == Key.Right: state["right"] = False

    loop = DispatcherTimer()
    loop.Interval = TimeSpan.FromMilliseconds(16)
    loop.Tick += tick

    canvas.KeyDown += on_key_down
    canvas.KeyUp   += on_key_up
    canvas.Focusable = True; canvas.Focus()
    update_visuals()
    ctx.set_status("PRESS ANY KEY THEN SPACE TO LAUNCH")
    ctx.set_title(">> BREAKOUT <<")
