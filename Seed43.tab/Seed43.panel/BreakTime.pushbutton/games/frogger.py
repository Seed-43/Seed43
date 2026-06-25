# -*- coding: utf-8 -*-
# Frogger - BreakTime game module

from System.Windows.Threading import DispatcherTimer
from System.Windows.Input import Key
from System.Windows.Shapes import Rectangle
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Controls import TextBlock, Canvas as WpfCanvas
from System.Windows import FontWeights
from System import TimeSpan
import random

W = 740; H = 520
CELL = 40
COLS = W // CELL   # 18
ROWS = H // CELL   # 13

C_BG        = Color.FromRgb(0,  15,  0)
C_SAFE      = Color.FromRgb(0,  30,  0)
C_ROAD      = Color.FromRgb(0,  20,  0)
C_WATER     = Color.FromRgb(0,  10,  30)
C_FROG      = Color.FromRgb(0,  255, 65)
C_CAR1      = Color.FromRgb(255,80,  0)
C_CAR2      = Color.FromRgb(200,200, 0)
C_CAR3      = Color.FromRgb(0,  150, 255)
C_LOG       = Color.FromRgb(80, 40,  0)
C_LOG_G     = Color.FromRgb(0,  120, 0)
C_GOAL      = Color.FromRgb(0,  200, 100)
C_DIM       = Color.FromRgb(0,  60,  0)
C_RED       = Color.FromRgb(255,51,  0)

def _brush(c): return SolidColorBrush(c)
def _rect(canvas, x, y, w, h, color, z=0):
    r = Rectangle(); r.Width=w; r.Height=h; r.Fill=_brush(color)
    WpfCanvas.SetLeft(r,x); WpfCanvas.SetTop(r,y); WpfCanvas.SetZIndex(r,z)
    canvas.Children.Add(r); return r

# Lane definitions: (row_from_bottom, type, obj_w, speed, count, color, direction)
LANES = [
    # row 1  = safe start
    # rows 2-6 = road
    (1, "road", 70,  2.5,  4, C_CAR1,  1),
    (2, "road", 100, 2.0,  3, C_CAR2, -1),
    (3, "road", 60,  3.5,  4, C_CAR3,  1),
    (4, "road", 80,  2.0,  3, C_CAR1, -1),
    (5, "road", 50,  4.0,  5, C_CAR2,  1),
    # row 7 = safe median
    # rows 8-12 = water/logs
    (7,  "log",  110, 2.0, 3, C_LOG,   1),
    (8,  "log",  80,  2.8, 4, C_LOG_G,-1),
    (9,  "log",  130, 1.8, 2, C_LOG,   1),
    (10, "log",  90,  2.5, 3, C_LOG_G,-1),
    (11, "log",  100, 2.0, 3, C_LOG,   1),
]

def init_game(ctx, timer):
    canvas = ctx.canvas
    canvas.Width=W; canvas.Height=H

    # Background strips
    for row in range(ROWS):
        y = row * CELL
        row_from_bottom = ROWS - 1 - row
        if row == 0:          # goal row
            c = C_GOAL
        elif row_from_bottom in (0, 6):  # safe zones
            c = C_SAFE
        elif row_from_bottom >= 7:       # water
            c = C_WATER
        else:                            # road
            c = C_ROAD
        _rect(canvas, 0, y, W, CELL, c)

    # Goal slots
    goal_slots = []
    for i in range(5):
        gx = 40 + i * 130
        el = _rect(canvas, gx, 2, 60, CELL-4, C_DIM, 1)
        goal_slots.append({"x": gx, "w": 60, "filled": False, "el": el})
    state = {
        "fx": W//2 - CELL//2,
        "fy": (ROWS-1) * CELL,
        "alive": True, "started": False,
        "score": 0, "lives": 3,
        "on_log": False, "log_vx": 0.0,
        "goals": goal_slots,
        "objs": [],   # [x, y, w, h, vx, el, type]
        "move_cd": 0,
    }

    frog_el = _rect(canvas, state["fx"], state["fy"], CELL-4, CELL-4, C_FROG, 3)
    state["frog_el"] = frog_el

    life_els = [_rect(canvas, W-10-i*22, H-22, 16, 10, C_FROG) for i in range(3)]
    state["life_els"] = life_els

    # Spawn lane objects
    for lane_def in LANES:
        row_from_bottom, ltype, obj_w, speed, count, color, direction = lane_def
        y = (ROWS - 1 - row_from_bottom) * CELL + 2
        spacing = W // count
        for i in range(count):
            x = i * spacing + random.randint(0, spacing//2)
            el = _rect(canvas, x, y, obj_w, CELL-4, color, 2)
            state["objs"].append([float(x), float(y), obj_w, CELL-4, speed*direction, el, ltype])

    def frog_row():
        return int(state["fy"] // CELL)

    def row_from_bottom():
        return ROWS - 1 - frog_row()

    def check_death():
        rfb = row_from_bottom()
        fx = state["fx"]; fy = state["fy"]
        # Road - hit by car
        if 1 <= rfb <= 5:
            for obj in state["objs"]:
                ox,oy,ow,oh,vx,el,otype = obj
                if otype != "road": continue
                if abs(oy - fy) < CELL and ox < fx+CELL-4 and ox+ow > fx:
                    return True
        # Water - not on log
        if 7 <= rfb <= 11:
            if not state["on_log"]:
                return True
        # Off sides
        if fx < -10 or fx > W:
            return True
        return False

    def check_log_riding():
        rfb = row_from_bottom()
        state["on_log"] = False
        state["log_vx"] = 0.0
        if 7 <= rfb <= 11:
            fx = state["fx"]; fy = state["fy"]
            for obj in state["objs"]:
                ox,oy,ow,oh,vx,el,otype = obj
                if otype != "log": continue
                if abs(oy - fy) < CELL-2 and ox < fx+CELL-4 and ox+ow > fx+4:
                    state["on_log"] = True
                    state["log_vx"] = vx
                    break

    def check_goal():
        if frog_row() != 0: return
        for g in state["goals"]:
            if not g["filled"] and g["x"] <= state["fx"] <= g["x"]+g["w"]-20:
                g["filled"] = True
                g["el"].Fill = _brush(C_FROG)
                state["score"] += 50
                ctx.set_score(state["score"])
                respawn()
                if all(g["filled"] for g in state["goals"]):
                    for g2 in state["goals"]:
                        g2["filled"] = False
                        g2["el"].Fill = _brush(C_DIM)
                    state["score"] += 200
                    ctx.set_score(state["score"])
                    ctx.set_status("ALL GOALS!  BONUS 200!  KEEP GOING!")
                return

    def respawn():
        state["fx"] = W//2 - CELL//2
        state["fy"] = (ROWS-1)*CELL
        state["on_log"] = False
        state["log_vx"] = 0.0
        WpfCanvas.SetLeft(frog_el, state["fx"])
        WpfCanvas.SetTop(frog_el,  state["fy"])

    def die():
        state["lives"] -= 1
        if state["lives"] >= 0 and state["lives"] < len(life_els):
            life_els[state["lives"]].Fill = _brush(Color.FromRgb(0,40,0))
        if state["lives"] <= 0:
            state["alive"] = False
            loop.Stop()
            ctx.set_status("GAME OVER!  SCORE: {}".format(state["score"]))
            return
        ctx.set_status("OUCH!  LIVES: {}  |  ARROWS TO MOVE".format(state["lives"]))
        respawn()

    def tick(sender, e):
        if not state["alive"] or not state["started"]:
            return
        # Move objects
        for obj in state["objs"]:
            obj[0] += obj[4]
            if obj[4] > 0 and obj[0] > W + 10:
                obj[0] = -obj[2]
            elif obj[4] < 0 and obj[0] + obj[2] < -10:
                obj[0] = W + 10
            WpfCanvas.SetLeft(obj[5], obj[0])

        check_log_riding()
        if state["on_log"]:
            state["fx"] += state["log_vx"]
            WpfCanvas.SetLeft(frog_el, state["fx"])

        if check_death():
            die()

        check_goal()

        if state["move_cd"] > 0:
            state["move_cd"] -= 1

    def on_key_down(sender, e):
        if not state["started"]:
            state["started"] = True
            loop.Start()
            ctx.set_status("ARROWS: HOP  |  GET TO THE TOP!")
            return
        if not state["alive"]: return
        if state["move_cd"] > 0: return
        k = e.Key
        if k == Key.Up:
            state["fy"] = max(0, state["fy"] - CELL)
            state["score"] += 10; ctx.set_score(state["score"])
        elif k == Key.Down:
            state["fy"] = min((ROWS-1)*CELL, state["fy"] + CELL)
        elif k == Key.Left:
            state["fx"] -= CELL
        elif k == Key.Right:
            state["fx"] += CELL
        WpfCanvas.SetLeft(frog_el, state["fx"])
        WpfCanvas.SetTop(frog_el,  state["fy"])
        state["move_cd"] = 4

    loop = DispatcherTimer()
    loop.Interval = TimeSpan.FromMilliseconds(33)
    loop.Tick += tick

    canvas.KeyDown += on_key_down
    canvas.Focusable = True; canvas.Focus()
    ctx.set_status("PRESS ANY KEY TO START  |  ARROWS TO HOP")
    ctx.set_title(">> FROGGER <<")
