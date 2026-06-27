# -*- coding: utf-8 -*-
# Pong - BreakTime game module  (1 player vs CPU)

from System.Windows.Threading import DispatcherTimer
from System.Windows.Input import Key
from System.Windows.Shapes import Rectangle, Ellipse, Line
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Controls import TextBlock, Canvas as WpfCanvas
from System.Windows import FontWeights
from System import TimeSpan
import random, math

W = 740; H = 500
PAD_W = 14; PAD_H = 80
BALL_R = 8
C_GREEN = Color.FromRgb(0,255,65)
C_DIM   = Color.FromRgb(0,80,20)
C_BG    = Color.FromRgb(0,15,0)

def _brush(c): return SolidColorBrush(c)
def _font():
    from System.Windows.Media import FontFamily
    return FontFamily("Courier New")

def init_game(ctx, timer):
    canvas = ctx.canvas
    canvas.Width = W; canvas.Height = H

    # BG
    bg = Rectangle(); bg.Width=W; bg.Height=H; bg.Fill=_brush(C_BG)
    WpfCanvas.SetLeft(bg,0); WpfCanvas.SetTop(bg,0); canvas.Children.Add(bg)
    # Centre line dashes
    for y in range(0, H, 20):
        l = Rectangle(); l.Width=2; l.Height=10; l.Fill=_brush(C_DIM)
        WpfCanvas.SetLeft(l, W//2-1); WpfCanvas.SetTop(l, y)
        canvas.Children.Add(l)

    def _score_lbl(x, y):
        tb = TextBlock(); tb.FontFamily=_font(); tb.FontSize=36
        tb.FontWeight=FontWeights.Bold; tb.Foreground=_brush(C_GREEN); tb.Text="0"
        WpfCanvas.SetLeft(tb,x); WpfCanvas.SetTop(tb,y); canvas.Children.Add(tb)
        return tb

    lbl_p  = _score_lbl(W//2 - 80, 20)
    lbl_cpu= _score_lbl(W//2 + 40, 20)

    def _pad(x, y):
        r = Rectangle(); r.Width=PAD_W; r.Height=PAD_H; r.Fill=_brush(C_GREEN)
        WpfCanvas.SetLeft(r,x); WpfCanvas.SetTop(r,y)
        WpfCanvas.SetZIndex(r,1); canvas.Children.Add(r); return r

    pad_p   = _pad(20,          H//2 - PAD_H//2)
    pad_cpu = _pad(W-20-PAD_W,  H//2 - PAD_H//2)

    ball_el = Ellipse(); ball_el.Width=BALL_R*2; ball_el.Height=BALL_R*2
    ball_el.Fill=_brush(Color.FromRgb(255,220,0))
    WpfCanvas.SetZIndex(ball_el,2); canvas.Children.Add(ball_el)

    state = {
        "py": float(H//2 - PAD_H//2),
        "cy": float(H//2 - PAD_H//2),
        "bx": float(W//2), "by": float(H//2),
        "vx": 5.0, "vy": 3.0,
        "score_p": 0, "score_c": 0,
        "up": False, "down": False,
        "alive": True, "started": False,
    }

    def reset_ball(dir_=1):
        state["bx"] = float(W//2); state["by"] = float(H//2)
        ang = random.uniform(-0.4, 0.4)
        state["vx"] = dir_ * 5.0
        state["vy"] = math.tan(ang) * 5.0

    def move_ball():
        state["bx"] += state["vx"]
        state["by"] += state["vy"]
        # Top/bottom bounce
        if state["by"] - BALL_R <= 0:
            state["by"] = BALL_R; state["vy"] = abs(state["vy"])
        if state["by"] + BALL_R >= H:
            state["by"] = H - BALL_R; state["vy"] = -abs(state["vy"])
        # Player paddle
        px2 = 20 + PAD_W
        if (state["bx"] - BALL_R <= px2 and
            state["by"] >= state["py"] and state["by"] <= state["py"]+PAD_H):
            state["bx"] = px2 + BALL_R
            rel = (state["by"] - (state["py"] + PAD_H/2)) / (PAD_H/2)
            state["vx"] = abs(state["vx"]) * 1.05
            state["vy"] = rel * 6
        # CPU paddle
        cx2 = W - 20 - PAD_W
        if (state["bx"] + BALL_R >= cx2 and
            state["by"] >= state["cy"] and state["by"] <= state["cy"]+PAD_H):
            state["bx"] = cx2 - BALL_R
            rel = (state["by"] - (state["cy"] + PAD_H/2)) / (PAD_H/2)
            state["vx"] = -abs(state["vx"]) * 1.05
            state["vy"] = rel * 6
            state["vx"] = max(-12.0, min(-4.0, state["vx"]))
        # Score
        if state["bx"] < 0:
            state["score_c"] += 1
            lbl_cpu.Text = str(state["score_c"])
            reset_ball(1)
        if state["bx"] > W:
            state["score_p"] += 1
            lbl_p.Text = str(state["score_p"])
            ctx.set_score(state["score_p"])
            reset_ball(-1)

    def move_cpu():
        mid = state["cy"] + PAD_H/2
        speed = 3.5
        if mid < state["by"] - 5:
            state["cy"] = min(H - PAD_H, state["cy"] + speed)
        elif mid > state["by"] + 5:
            state["cy"] = max(0, state["cy"] - speed)

    def update_visuals():
        WpfCanvas.SetTop(pad_p,   state["py"])
        WpfCanvas.SetTop(pad_cpu, state["cy"])
        WpfCanvas.SetLeft(ball_el, state["bx"] - BALL_R)
        WpfCanvas.SetTop(ball_el,  state["by"] - BALL_R)

    def tick(sender, e):
        if not state["alive"] or not state["started"]:
            return
        if state["up"]:
            state["py"] = max(0, state["py"] - 6)
        if state["down"]:
            state["py"] = min(H - PAD_H, state["py"] + 6)
        move_cpu()
        move_ball()
        update_visuals()

    def on_key_down(sender, e):
        if e.Key == Key.Up:    state["up"]   = True
        if e.Key == Key.Down:  state["down"] = True
        if not state["started"]:
            state["started"] = True
            loop.Start()
            ctx.set_status("UP/DOWN ARROWS  |  YOU vs CPU")

    def on_key_up(sender, e):
        if e.Key == Key.Up:   state["up"]   = False
        if e.Key == Key.Down: state["down"] = False

    loop = DispatcherTimer()
    loop.Interval = TimeSpan.FromMilliseconds(16)
    loop.Tick += tick

    canvas.KeyDown += on_key_down
    canvas.KeyUp   += on_key_up
    canvas.Focusable = True
    canvas.Focus()
    reset_ball()
    update_visuals()
    ctx.set_status("PRESS ANY KEY TO START  |  UP/DOWN ARROWS TO MOVE")
    ctx.set_title(">> PONG <<")
