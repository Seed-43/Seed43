# -*- coding: utf-8 -*-
# Space Invaders - BreakTime game module

from System.Windows.Threading import DispatcherTimer
from System.Windows.Input import Key
from System.Windows.Shapes import Rectangle, Ellipse, Polygon
from System.Windows.Media import SolidColorBrush, Color, PointCollection
from System.Windows.Controls import TextBlock, Canvas as WpfCanvas
from System.Windows import FontWeights, Point
from System import TimeSpan
import random

W = 740; H = 520
C_GREEN  = Color.FromRgb(0,255,65)
C_DIM    = Color.FromRgb(0,80,20)
C_BG     = Color.FromRgb(0,15,0)
C_BULLET = Color.FromRgb(255,255,0)
C_ALIEN1 = Color.FromRgb(0,255,65)
C_ALIEN2 = Color.FromRgb(0,200,200)
C_ALIEN3 = Color.FromRgb(255,100,0)
C_SHIELD = Color.FromRgb(0,150,40)
C_RED    = Color.FromRgb(255,51,0)

ALIEN_COLS = 11
ALIEN_ROWS = 5
ALIEN_W    = 36
ALIEN_H    = 26
ALIEN_PAD_X= 14
ALIEN_PAD_Y= 12

def _brush(c): return SolidColorBrush(c)
def _font():
    from System.Windows.Media import FontFamily
    return FontFamily("Courier New")

def _rect(canvas, x, y, w, h, color, z=0):
    r = Rectangle(); r.Width=w; r.Height=h; r.Fill=_brush(color)
    WpfCanvas.SetLeft(r,x); WpfCanvas.SetTop(r,y); WpfCanvas.SetZIndex(r,z)
    canvas.Children.Add(r); return r

def init_game(ctx, timer):
    canvas = ctx.canvas
    canvas.Width=W; canvas.Height=H

    _rect(canvas, 0,0,W,H,C_BG)
    # Ground line
    _rect(canvas, 0, H-30, W, 2, C_DIM)

    state = {
        "aliens":       {},    # (r,c) -> rect or None if dead
        "alien_dir":    1,
        "alien_step":   0,
        "alien_x":      0.0,
        "alien_y":      60.0,
        "alive":        True,
        "started":      False,
        "player_x":     float(W//2 - 15),
        "bullets":      [],    # list of [x, y, rect]
        "abombs":       [],    # alien bombs
        "score":        0,
        "lives":        3,
        "left":  False, "right": False, "fire_cd": 0,
        "wave":         1,
    }

    # Draw player ship
    player_el = _rect(canvas, state["player_x"], H-60, 30, 14, C_GREEN, 2)
    nose_el   = _rect(canvas, state["player_x"]+13, H-70, 4, 10, C_GREEN, 2)
    state["player_el"] = player_el
    state["nose_el"]   = nose_el

    # Life indicators
    life_els = []
    for i in range(3):
        el = _rect(canvas, 10+i*30, H-24, 20, 10, C_GREEN)
        life_els.append(el)
    state["life_els"] = life_els

    # Shields
    shield_cells = []
    for si in range(4):
        sx = 80 + si * 160
        sy = H - 110
        for row in range(4):
            for col in range(8):
                el = _rect(canvas, sx+col*8, sy+row*8, 7, 7, C_SHIELD)
                shield_cells.append([sx+col*8, sy+row*8, el, True])
    state["shields"] = shield_cells

    def alien_color(r):
        if r == 0: return C_ALIEN3
        if r <= 2: return C_ALIEN2
        return C_ALIEN1

    def spawn_aliens(wave):
        for (r,c), el in list(state["aliens"].items()):
            if el and el in canvas.Children:
                canvas.Children.Remove(el)
        state["aliens"] = {}
        state["alien_x"] = 0.0
        state["alien_y"] = 60.0
        state["alien_dir"] = 1
        for r in range(ALIEN_ROWS):
            for c in range(ALIEN_COLS):
                x = 60 + c * (ALIEN_W + ALIEN_PAD_X)
                y = 60 + r * (ALIEN_H + ALIEN_PAD_Y)
                el = _rect(canvas, x, y, ALIEN_W, ALIEN_H, alien_color(r), 1)
                state["aliens"][(r,c)] = el

    spawn_aliens(1)

    def update_visuals():
        WpfCanvas.SetLeft(player_el, state["player_x"])
        WpfCanvas.SetLeft(nose_el,   state["player_x"]+13)

    from System.Windows.Media import FontFamily
    score_tb = TextBlock(); score_tb.FontFamily=FontFamily("Courier New")
    score_tb.FontSize=13; score_tb.Foreground=_brush(C_GREEN); score_tb.Text="SCORE: 0"
    WpfCanvas.SetLeft(score_tb, W-120); WpfCanvas.SetTop(score_tb, 5); canvas.Children.Add(score_tb)

    def fire_bullet():
        bx = state["player_x"] + 13
        by = H - 72
        el = _rect(canvas, bx, by, 3, 10, C_BULLET, 3)
        state["bullets"].append([bx, by, el])

    def move_aliens(tick_count=[0]):
        tick_count[0] += 1
        alive_count = sum(1 for el in state["aliens"].values() if el is not None)
        speed = max(2, 8 - alive_count // 5)
        if tick_count[0] % speed != 0:
            return
        dx = state["alien_dir"] * 6
        hit_edge = False
        for (r,c), el in state["aliens"].items():
            if el is None: continue
            nx = WpfCanvas.GetLeft(el) + dx
            if nx < 10 or nx > W - ALIEN_W - 10:
                hit_edge = True; break
        if hit_edge:
            state["alien_dir"] *= -1
            for (r,c), el in state["aliens"].items():
                if el is None: continue
                WpfCanvas.SetTop(el, WpfCanvas.GetTop(el) + 12)
        else:
            for (r,c), el in state["aliens"].items():
                if el is None: continue
                WpfCanvas.SetLeft(el, WpfCanvas.GetLeft(el) + dx)

    def alien_shoot():
        live = [(k,el) for k,el in state["aliens"].items() if el is not None]
        if not live: return
        (r,c), el = random.choice(live)
        bx = WpfCanvas.GetLeft(el) + ALIEN_W//2
        by = WpfCanvas.GetTop(el) + ALIEN_H
        bel = _rect(canvas, bx, by, 3, 10, C_RED, 3)
        state["abombs"].append([bx, by, bel])

    def check_bullet_hits():
        dead_bullets = []
        for b in state["bullets"]:
            bx, by, bel = b
            # Aliens
            for (r,c), el in list(state["aliens"].items()):
                if el is None: continue
                ax = WpfCanvas.GetLeft(el); ay = WpfCanvas.GetTop(el)
                if ax <= bx <= ax+ALIEN_W and ay <= by <= ay+ALIEN_H:
                    canvas.Children.Remove(el)
                    state["aliens"][(r,c)] = None
                    dead_bullets.append(b)
                    pts = (ALIEN_ROWS - r) * 10
                    state["score"] += pts
                    ctx.set_score(state["score"])
                    score_tb.Text = "SCORE: {}".format(state["score"])
                    break
            # Shields
            for sh in state["shields"]:
                if not sh[3]: continue
                if sh[0] <= bx <= sh[0]+7 and sh[1] <= by <= sh[1]+7:
                    sh[3] = False
                    sh[2].Fill = _brush(C_BG)
                    dead_bullets.append(b)
                    break
        for b in dead_bullets:
            if b in state["bullets"]:
                state["bullets"].remove(b)
            if b[2] in canvas.Children:
                canvas.Children.Remove(b[2])

    def check_bomb_hits():
        dead = []
        for b in state["abombs"]:
            bx, by, bel = b
            px = state["player_x"]
            if px <= bx <= px+30 and H-60 <= by <= H-46:
                dead.append(b)
                state["lives"] -= 1
                if state["lives"] >= 0 and state["lives"] < len(life_els):
                    life_els[state["lives"]].Fill = _brush(C_DIM)
                if state["lives"] <= 0:
                    state["alive"] = False
                    loop.Stop()
                    ctx.set_status("GAME OVER!  SCORE: {}".format(state["score"]))
            for sh in state["shields"]:
                if not sh[3]: continue
                if sh[0] <= bx <= sh[0]+7 and sh[1] <= by <= sh[1]+7:
                    sh[3] = False; sh[2].Fill = _brush(C_BG)
                    dead.append(b); break
        for b in dead:
            if b in state["abombs"]:
                state["abombs"].remove(b)
            if b[2] in canvas.Children:
                canvas.Children.Remove(b[2])

    shoot_cd = [0]
    wave_clear_cd = [0]

    def tick(sender, e):
        if not state["alive"] or not state["started"]:
            return
        # Player movement
        if state["left"]:
            state["player_x"] = max(0, state["player_x"] - 5)
        if state["right"]:
            state["player_x"] = min(W-30, state["player_x"] + 5)
        update_visuals()

        # Bullets
        for b in list(state["bullets"]):
            b[1] -= 10
            WpfCanvas.SetTop(b[2], b[1])
            if b[1] < 0:
                state["bullets"].remove(b)
                if b[2] in canvas.Children:
                    canvas.Children.Remove(b[2])

        # Alien bombs
        for b in list(state["abombs"]):
            b[1] += 6
            WpfCanvas.SetTop(b[2], b[1])
            if b[1] > H:
                state["abombs"].remove(b)
                if b[2] in canvas.Children:
                    canvas.Children.Remove(b[2])

        move_aliens()
        check_bullet_hits()
        check_bomb_hits()

        if state["fire_cd"] > 0:
            state["fire_cd"] -= 1

        # Alien random shoot
        shoot_cd[0] += 1
        if shoot_cd[0] >= 40:
            shoot_cd[0] = 0
            if random.random() < 0.4:
                alien_shoot()

        # Check wave clear
        if all(el is None for el in state["aliens"].values()):
            wave_clear_cd[0] += 1
            if wave_clear_cd[0] > 30:
                wave_clear_cd[0] = 0
                state["wave"] += 1
                spawn_aliens(state["wave"])
                ctx.set_status("WAVE {}!".format(state["wave"]))

        # Aliens reach bottom
        for (r,c), el in state["aliens"].items():
            if el is None: continue
            if WpfCanvas.GetTop(el) + ALIEN_H >= H - 40:
                state["alive"] = False
                loop.Stop()
                ctx.set_status("THEY LANDED!  GAME OVER!  SCORE: {}".format(state["score"]))
                return

    def on_key_down(sender, e):
        if e.Key == Key.Left:  state["left"]  = True
        if e.Key == Key.Right: state["right"] = True
        if e.Key == Key.Space:
            if state["fire_cd"] == 0 and len(state["bullets"]) < 3:
                fire_bullet(); state["fire_cd"] = 15
        if not state["started"]:
            state["started"] = True
            loop.Start()
            ctx.set_status("ARROWS: MOVE  |  SPACE: FIRE  |  SURVIVE!")

    def on_key_up(sender, e):
        if e.Key == Key.Left:  state["left"]  = False
        if e.Key == Key.Right: state["right"] = False

    loop = DispatcherTimer()
    loop.Interval = TimeSpan.FromMilliseconds(33)
    loop.Tick += tick

    canvas.KeyDown += on_key_down
    canvas.KeyUp   += on_key_up
    canvas.Focusable = True; canvas.Focus()
    ctx.set_status("PRESS ANY KEY TO START  |  ARROWS + SPACE")
    ctx.set_title(">> SPACE INVADERS <<")
