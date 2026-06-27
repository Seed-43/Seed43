# -*- coding: utf-8 -*-
# Asteroids lite - BreakTime game module

from System.Windows.Threading import DispatcherTimer
from System.Windows.Input import Key
from System.Windows.Shapes import Polygon, Ellipse, Rectangle
from System.Windows.Media import SolidColorBrush, Color, PointCollection
from System.Windows.Controls import Canvas as WpfCanvas
from System.Windows import Point, FontWeights
from System import TimeSpan
import random, math

W = 740; H = 520
C_BG   = Color.FromRgb(0,  15,  0)
C_SHIP = Color.FromRgb(0,  255, 65)
C_ROCK = Color.FromRgb(0,  180, 40)
C_BULL = Color.FromRgb(255,220,  0)
C_THRUST=Color.FromRgb(255,100,  0)

def _brush(c): return SolidColorBrush(c)
def _font():
    from System.Windows.Media import FontFamily
    return FontFamily("Courier New")

def make_polygon(canvas, points_list, color, z=1):
    poly = Polygon()
    pts  = PointCollection()
    for (x,y) in points_list:
        pts.Add(Point(x, y))
    poly.Points = pts
    poly.Stroke = _brush(color)
    poly.StrokeThickness = 2
    poly.Fill   = _brush(Color.FromRgb(0,15,0))
    WpfCanvas.SetZIndex(poly, z)
    canvas.Children.Add(poly)
    return poly

def set_polygon_points(poly, points_list):
    pts = PointCollection()
    for (x,y) in points_list:
        pts.Add(Point(x,y))
    poly.Points = pts

def init_game(ctx, timer):
    canvas = ctx.canvas
    canvas.Width=W; canvas.Height=H

    bg = Rectangle(); bg.Width=W; bg.Height=H; bg.Fill=_brush(C_BG)
    WpfCanvas.SetLeft(bg,0); WpfCanvas.SetTop(bg,0); canvas.Children.Add(bg)

    state = {
        "x": float(W//2), "y": float(H//2),
        "angle": 0.0,
        "vx": 0.0, "vy": 0.0,
        "thrusting": False,
        "rotating_l": False, "rotating_r": False,
        "firing": False, "fire_cd": 0,
        "bullets": [],   # [x, y, vx, vy, life, el]
        "rocks":   [],   # [x, y, vx, vy, size, angle, spin, poly]
        "alive": True, "started": False,
        "score": 0, "lives": 3,
        "invincible": 60,
    }

    # Ship polygon (points around origin, rotated/translated at draw time)
    SHIP_PTS = [(0,-14),(9,10),(-9,10)]

    ship_poly   = make_polygon(canvas, [(0,0),(0,0),(0,0)], C_SHIP, 3)
    thrust_poly = make_polygon(canvas, [(0,0),(0,0),(0,0)], C_THRUST, 3)
    state["ship_poly"]   = ship_poly
    state["thrust_poly"] = thrust_poly

    from System.Windows.Controls import TextBlock
    score_tb = TextBlock(); score_tb.FontFamily=_font(); score_tb.FontSize=13
    score_tb.Foreground=_brush(C_SHIP); score_tb.Text="SCORE: 0"
    WpfCanvas.SetLeft(score_tb,10); WpfCanvas.SetTop(score_tb,5); canvas.Children.Add(score_tb)

    from System.Windows.Shapes import Rectangle as Rect
    life_els = []
    for i in range(3):
        el = Rect(); el.Width=14; el.Height=10; el.Fill=_brush(C_SHIP)
        WpfCanvas.SetLeft(el, W-10-i*20); WpfCanvas.SetTop(el, 8)
        canvas.Children.Add(el); life_els.append(el)
    state["life_els"] = life_els

    def rotate_pt(x, y, angle):
        c = math.cos(angle); s = math.sin(angle)
        return (x*c - y*s, x*s + y*c)

    def transform_ship(pts, ox, oy, angle):
        return [(ox + rotate_pt(x,y,angle)[0], oy + rotate_pt(x,y,angle)[1]) for (x,y) in pts]

    def draw_ship():
        ox = state["x"]; oy = state["y"]; a = state["angle"]
        SHIP = [(0,-14),(9,10),(-9,10)]
        set_polygon_points(ship_poly,   transform_ship(SHIP, ox, oy, a))
        if state["thrusting"]:
            THRUST = [(4,10),(-4,10),(0,22)]
            set_polygon_points(thrust_poly, transform_ship(THRUST, ox, oy, a))
            thrust_poly.Stroke = _brush(C_THRUST)
        else:
            set_polygon_points(thrust_poly, [(ox,oy),(ox,oy),(ox,oy)])

    def spawn_rock(size=3, x=None, y=None, vx=None, vy=None):
        if x is None:
            while True:
                x = random.uniform(0, W); y = random.uniform(0, H)
                if abs(x-state["x"]) > 80 or abs(y-state["y"]) > 80:
                    break
        if vx is None:
            speed = random.uniform(0.5, 1.5) * (4-size)
            angle2 = random.uniform(0, math.pi*2)
            vx = math.cos(angle2)*speed; vy = math.sin(angle2)*speed
        r  = size * 18
        pts_count = random.randint(7,11)
        pts = []
        for i in range(pts_count):
            ang = i * math.pi*2 / pts_count + random.uniform(-0.3,0.3)
            radius = r * random.uniform(0.7, 1.0)
            pts.append((math.cos(ang)*radius, math.sin(ang)*radius))
        poly2 = make_polygon(canvas, [(x+px2, y+py2) for (px2,py2) in pts], C_ROCK, 1)
        state["rocks"].append({"x":x,"y":y,"vx":vx,"vy":vy,"size":size,
                                "angle":random.uniform(0,math.pi*2),
                                "spin":random.uniform(-0.03,0.03),
                                "pts":pts,"poly":poly2})

    def spawn_wave():
        for _ in range(4 + len([r for r in state["rocks"]]) // 2):
            spawn_rock()

    spawn_wave()
    draw_ship()

    def update_rocks():
        for rock in state["rocks"]:
            rock["x"] = (rock["x"] + rock["vx"]) % W
            rock["y"] = (rock["y"] + rock["vy"]) % H
            rock["angle"] += rock["spin"]
            a = rock["angle"]; ox = rock["x"]; oy = rock["y"]
            transformed = [(ox + rotate_pt(px2,py2,a)[0], oy + rotate_pt(px2,py2,a)[1])
                           for (px2,py2) in rock["pts"]]
            set_polygon_points(rock["poly"], transformed)

    def update_bullets():
        dead = []
        for b in state["bullets"]:
            b[0] = (b[0]+b[2]) % W
            b[1] = (b[1]+b[3]) % H
            b[4] -= 1
            WpfCanvas.SetLeft(b[5], b[0]-2); WpfCanvas.SetTop(b[5], b[1]-2)
            if b[4] <= 0:
                dead.append(b)
        for b in dead:
            state["bullets"].remove(b)
            if b[5] in canvas.Children: canvas.Children.Remove(b[5])

    def check_bullet_rock():
        dead_b = []; dead_r = []
        for b in state["bullets"]:
            for rock in state["rocks"]:
                r_radius = rock["size"] * 18
                dx = b[0]-rock["x"]; dy = b[1]-rock["y"]
                if dx*dx+dy*dy < r_radius*r_radius:
                    dead_b.append(b); dead_r.append(rock)
                    pts_val = [0,100,50,20][min(rock["size"],3)]
                    state["score"] += pts_val
                    ctx.set_score(state["score"])
                    score_tb.Text = "SCORE: {}".format(state["score"])
                    if rock["size"] > 1:
                        for _ in range(2):
                            spawn_rock(rock["size"]-1, rock["x"], rock["y"])
                    break
        for b in set(id(x) for x in dead_b):
            pass
        for b in dead_b:
            if b in state["bullets"]: state["bullets"].remove(b)
            if b[5] in canvas.Children: canvas.Children.Remove(b[5])
        for rock in dead_r:
            if rock in state["rocks"]: state["rocks"].remove(rock)
            if rock["poly"] in canvas.Children: canvas.Children.Remove(rock["poly"])
        if not state["rocks"]:
            spawn_wave()

    def check_ship_rock():
        if state["invincible"] > 0:
            state["invincible"] -= 1
            return
        sx = state["x"]; sy = state["y"]
        for rock in state["rocks"]:
            r_radius = rock["size"] * 18
            dx = sx-rock["x"]; dy = sy-rock["y"]
            if dx*dx+dy*dy < (r_radius+10)**2:
                state["lives"] -= 1
                if state["lives"] >= 0 and state["lives"] < len(life_els):
                    life_els[state["lives"]].Fill = _brush(Color.FromRgb(0,40,0))
                if state["lives"] <= 0:
                    state["alive"] = False; loop.Stop()
                    ctx.set_status("GAME OVER!  SCORE: {}".format(state["score"]))
                else:
                    state["x"]=float(W//2); state["y"]=float(H//2)
                    state["vx"]=0; state["vy"]=0; state["invincible"]=90
                break

    def tick(sender, e):
        if not state["alive"] or not state["started"]: return
        ROT_SPD = 0.07; THRUST_PWR = 0.18; DRAG = 0.98; MAX_SPD = 8
        if state["rotating_l"]: state["angle"] -= ROT_SPD
        if state["rotating_r"]: state["angle"] += ROT_SPD
        if state["thrusting"]:
            state["vx"] += math.sin(state["angle"]) * THRUST_PWR
            state["vy"] -= math.cos(state["angle"]) * THRUST_PWR
        state["vx"] *= DRAG; state["vy"] *= DRAG
        spd = math.sqrt(state["vx"]**2+state["vy"]**2)
        if spd > MAX_SPD:
            state["vx"] *= MAX_SPD/spd; state["vy"] *= MAX_SPD/spd
        state["x"] = (state["x"]+state["vx"]) % W
        state["y"] = (state["y"]+state["vy"]) % H
        if state["fire_cd"] > 0: state["fire_cd"] -= 1
        if state["firing"] and state["fire_cd"] == 0:
            bx = state["x"]; by = state["y"]
            bvx = math.sin(state["angle"])*12+state["vx"]
            bvy = -math.cos(state["angle"])*12+state["vy"]
            el = Ellipse(); el.Width=5; el.Height=5; el.Fill=_brush(C_BULL)
            WpfCanvas.SetZIndex(el,2); canvas.Children.Add(el)
            state["bullets"].append([bx,by,bvx,bvy,40,el])
            state["fire_cd"] = 10
        draw_ship()
        update_bullets()
        update_rocks()
        check_bullet_rock()
        check_ship_rock()

    def on_key_down(sender, e):
        k = e.Key
        if k == Key.Left:  state["rotating_l"] = True
        if k == Key.Right: state["rotating_r"] = True
        if k == Key.Up:    state["thrusting"]   = True
        if k == Key.Space: state["firing"]       = True
        if not state["started"]:
            state["started"] = True; loop.Start()
            ctx.set_status("ARROWS: ROTATE/THRUST  |  SPACE: FIRE")

    def on_key_up(sender, e):
        k = e.Key
        if k == Key.Left:  state["rotating_l"] = False
        if k == Key.Right: state["rotating_r"] = False
        if k == Key.Up:    state["thrusting"]   = False
        if k == Key.Space: state["firing"]       = False

    loop = DispatcherTimer()
    loop.Interval = TimeSpan.FromMilliseconds(33)
    loop.Tick += tick

    canvas.KeyDown += on_key_down
    canvas.KeyUp   += on_key_up
    canvas.Focusable = True; canvas.Focus()
    ctx.set_status("PRESS ANY KEY TO START  |  ARROWS + SPACE")
    ctx.set_title(">> ASTEROIDS <<")
