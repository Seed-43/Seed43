# -*- coding: utf-8 -*-
# Bowling - BreakTime game module
# Layout faithful to TechStaff Corporation BOWL.COM (1989-92)
# Black scoreboard top half, green lane+pins bottom half.
# Ball travels from bottom-left toward pins in top-right.

from System.Windows.Threading import DispatcherTimer
from System.Windows.Input import Key
from System.Windows.Shapes import Rectangle, Ellipse, Line
from System.Windows.Media import SolidColorBrush, Color, FontFamily
from System.Windows.Controls import TextBlock, Canvas as WpfCanvas
from System.Windows import FontWeights, TextAlignment
from System import TimeSpan
import random

# ---------------------------------------------------------------------------
# Layout
# ---------------------------------------------------------------------------
CANVAS_W  = 740
CANVAS_H  = 520

# Scoreboard - top black area
SCORE_H   = 85
SCORE_Y   = 0

# Green play area - bottom half
PLAY_Y    = SCORE_H
PLAY_H    = CANVAS_H - SCORE_H   # ~380px tall
PLAY_W    = CANVAS_W

# Ball starts left-centre of play area
BALL_START_X = 60.0
BALL_START_Y = -1.0   # calculated after PLAY_H known

# Pins are on the right side of the play area
# Pin triangle: pin 1 (head) at LEFT, back row 7-10 at RIGHT
PIN_TIP_X  = 600   # head pin X (leftmost pin)
PIN_TIP_Y  = -1.0   # set after lane calc
PIN_GAP_X  = 34    # horizontal spacing between rows (going right)
PIN_GAP_Y  = 28    # vertical spacing within rows

# Aim bar - ball oscillates UP/DOWN on the left side
AIM_MIN_Y  = 0.0   # set after lane calc
AIM_MAX_Y  = 1.0   # set after lane calc

# ---------------------------------------------------------------------------
# Colours
# ---------------------------------------------------------------------------
def _b(r, g, b):
    return SolidColorBrush(Color.FromRgb(r, g, b))

C_BLACK    = _b(0,   0,   0)
C_GREEN    = _b(0,   180, 0)       # bright green play area
C_WHITE    = _b(220, 255, 220)     # scorecard text
C_BRIGHT   = _b(0,   255, 65)      # highlight green
C_YELLOW   = _b(255, 220, 0)
C_BALL     = _b(255, 220, 0)
C_PIN_UP   = _b(0,   255, 65)
C_PIN_DN   = _b(0,   80,  0)
C_DIM      = _b(0,   120, 0)
C_AIM      = _b(0,   100, 0)
C_AIM_HOT  = _b(0,   220, 0)
C_BORDER   = _b(0,   220, 0)
C_RED      = _b(255, 51,  0)

def _rect(cv, x, y, w, h, color, z=0):
    r = Rectangle(); r.Width=w; r.Height=h; r.Fill=color
    WpfCanvas.SetLeft(r,x); WpfCanvas.SetTop(r,y); WpfCanvas.SetZIndex(r,z)
    cv.Children.Add(r); return r

def _line(cv, x1, y1, x2, y2, color, thick=1, z=0):
    l = Line(); l.X1=x1; l.Y1=y1; l.X2=x2; l.Y2=y2
    l.Stroke=color; l.StrokeThickness=thick; WpfCanvas.SetZIndex(l,z)
    cv.Children.Add(l); return l

def _dot(cv, cx, cy, r, color, z=2):
    e = Ellipse(); e.Width=r*2; e.Height=r*2; e.Fill=color
    WpfCanvas.SetLeft(e,cx-r); WpfCanvas.SetTop(e,cy-r)
    WpfCanvas.SetZIndex(e,z); cv.Children.Add(e); return e

def _lbl(cv, x, y, text, color, size=13, bold=False, align=None):
    from System.Windows.Media import FontFamily as FF
    tb = TextBlock()
    tb.FontFamily = FF("Courier New")
    tb.FontSize   = size
    tb.Foreground = color
    tb.Text       = text
    if bold: tb.FontWeight = FontWeights.Bold
    if align == "center": tb.TextAlignment = TextAlignment.Center
    WpfCanvas.SetLeft(tb, x); WpfCanvas.SetTop(tb, y)
    cv.Children.Add(tb); return tb

# ---------------------------------------------------------------------------
# Pin positions - triangle with tip at bottom-left, spreads up and right
# Row 0 = pin 1 (head, bottom-left)
# Row 3 = pins 7-10 (back, top-right)
# ---------------------------------------------------------------------------
def build_pin_pos(tip_y):
    rows = [[0],[1,2],[3,4,5],[6,7,8,9]]
    pos  = [None]*10
    for ri, pins in enumerate(rows):
        n  = len(pins)
        px = PIN_TIP_X + ri * PIN_GAP_X
        for ci, pn in enumerate(pins):
            py = tip_y - (n-1)*PIN_GAP_Y//2 + ci*PIN_GAP_Y
            pos[pn] = (px, py)
    return pos

# ---------------------------------------------------------------------------
# Scoring
# ---------------------------------------------------------------------------
def score_frames(rolls):
    frames = []
    i = 0
    for f in range(10):
        if i >= len(rolls):
            frames.append((None, None, None, None)); continue
        b1 = rolls[i]   if i   < len(rolls) else None
        b2 = rolls[i+1] if i+1 < len(rolls) else None
        if f < 9:
            if b1 == 10:
                n1 = rolls[i+1] if i+1 < len(rolls) else None
                n2 = rolls[i+2] if i+2 < len(rolls) else None
                sc = (10+n1+n2) if (n1 is not None and n2 is not None) else None
                frames.append((sc, b1, None, None)); i += 1
            elif b1 is not None and b2 is not None and b1+b2 == 10:
                n1 = rolls[i+2] if i+2 < len(rolls) else None
                sc = (10+n1) if n1 is not None else None
                frames.append((sc, b1, b2, None)); i += 2
            else:
                sc = (b1+b2) if b2 is not None else None
                frames.append((sc, b1, b2, None)); i += 2
        else:
            b3 = rolls[i+2] if i+2 < len(rolls) else None
            if b1 == 10 or (b1 is not None and b2 is not None and b1+b2 == 10):
                needed = 3
            else:
                needed = 2
            balls = [b for b in [b1,b2,b3] if b is not None]
            sc = sum(balls) if len(balls) >= needed else None
            frames.append((sc, b1, b2, b3))
    return frames

def fmt_ball(val, prev=None, ball_idx=0):
    if val is None:   return " "
    if val == 10 and ball_idx == 0: return "X"
    if prev is not None and prev+val == 10: return "/"
    if val == 0:      return "-"
    return str(val)

# ---------------------------------------------------------------------------
# Best score persistence
# ---------------------------------------------------------------------------
import os as _os

def _load_best(games_dir):
    try:
        with open(_os.path.join(games_dir, "bowling_best.txt")) as f:
            return int(f.read().strip())
    except:
        return 0

def _save_best(games_dir, score):
    try:
        with open(_os.path.join(games_dir, "bowling_best.txt"), 'w') as f:
            f.write(str(score))
    except:
        pass

def _load_diff(games_dir):
    try:
        with open(_os.path.join(games_dir, "bowling_diff.txt")) as f:
            return int(f.read().strip())
    except:
        return 50

def _save_diff(games_dir, d):
    try:
        with open(_os.path.join(games_dir, "bowling_diff.txt"), 'w') as f:
            f.write(str(d))
    except:
        pass

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def init_game(ctx, timer):
    cv = ctx.canvas
    cv.Width = CANVAS_W; cv.Height = CANVAS_H

    # Load persistent values
    gd           = ctx.games_dir
    best_score   = _load_best(gd)
    start_diff   = _load_diff(gd)
    next_diff    = start_diff + 5
    if next_diff > 60: next_diff = 0
    _save_diff(gd, next_diff)

    # --- Black scoreboard area ---
    _rect(cv, 0, 0, CANVAS_W, SCORE_H, C_BLACK)

    # Divider line
    _line(cv, 0, SCORE_H, CANVAS_W, SCORE_H, C_BORDER, 2)

    # --- Black background, narrow bright green lane in the middle ---
    GUTTER_H = int(PLAY_H * 0.30)
    LANE_TOP  = PLAY_Y + GUTTER_H
    LANE_BOT  = PLAY_Y + PLAY_H - GUTTER_H
    LANE_H_PX = LANE_BOT - LANE_TOP
    _rect(cv, 0, PLAY_Y, CANVAS_W, PLAY_H, C_BLACK)         # black surround
    _rect(cv, 0, LANE_TOP, CANVAS_W, LANE_H_PX, C_GREEN)    # bright lane only
    _line(cv, 0, LANE_TOP, CANVAS_W, LANE_TOP, _b(0,140,0), 2)
    _line(cv, 0, LANE_BOT, CANVAS_W, LANE_BOT, _b(0,140,0), 2)

    # Guide lines - top, centre, bottom of lane
    for fy in [0.15, 0.5, 0.85]:
        ly = int(LANE_TOP + LANE_H_PX * fy)
        _line(cv, 20, ly, PIN_TIP_X - 50, ly, _b(0,130,0), 1)

    # Approach dots - 3 columns, aligned to guide lines
    for dx in [80, 140, 200]:
        for fy in [0.15, 0.5, 0.85]:
            _dot(cv, dx, int(LANE_TOP + LANE_H_PX * fy), 4, _b(0,140,0))

    # Arrow markers pointing RIGHT, aligned to guide lines
    arrow_x = 340
    for fy in [0.15, 0.5, 0.85]:
        ay = int(LANE_TOP + LANE_H_PX * fy)
        _line(cv, arrow_x,    ay, arrow_x-20, ay-8, _b(0,150,0), 2)
        _line(cv, arrow_x,    ay, arrow_x-20, ay+8, _b(0,150,0), 2)

    # --- Scoreboard layout ---
    # Frame number row
    frame_w = (CANVAS_W - 120) // 10
    card_x  = 120



    for f in range(10):
        fx = card_x + f * frame_w
        _lbl(cv, fx + frame_w//2 - 6, 4, str(f+1), C_WHITE, 11)
        _line(cv, fx, 0, fx, SCORE_H, C_DIM, 1)

    _line(cv, card_x + 10*frame_w, 0, card_x + 10*frame_w, SCORE_H, C_DIM, 1)

    # Ball boxes row (inside each frame)
    BOX_Y  = 20
    BOX_H  = 28
    box_lbls = []
    for f in range(10):
        fx = card_x + f * frame_w
        _line(cv, fx, BOX_Y, fx + frame_w, BOX_Y, C_DIM, 1)
        _line(cv, fx, BOX_Y+BOX_H, fx+frame_w, BOX_Y+BOX_H, C_DIM, 1)
        if f < 9:
            half = frame_w // 2
            _line(cv, fx+half, BOX_Y, fx+half, BOX_Y+BOX_H, C_DIM, 1)
            b1_tb = _lbl(cv, fx+3,      BOX_Y+6, "", C_BRIGHT, 12, True)
            b2_tb = _lbl(cv, fx+half+3, BOX_Y+6, "", C_BRIGHT, 12, True)
            b3_tb = None
        else:
            third = frame_w // 3
            _line(cv, fx+third,   BOX_Y, fx+third,   BOX_Y+BOX_H, C_DIM, 1)
            _line(cv, fx+third*2, BOX_Y, fx+third*2, BOX_Y+BOX_H, C_DIM, 1)
            b1_tb = _lbl(cv, fx+2,          BOX_Y+6, "", C_BRIGHT, 11, True)
            b2_tb = _lbl(cv, fx+third+2,    BOX_Y+6, "", C_BRIGHT, 11, True)
            b3_tb = _lbl(cv, fx+third*2+2,  BOX_Y+6, "", C_BRIGHT, 11, True)

        sc_tb  = _lbl(cv, fx+4, BOX_Y+BOX_H+4,  "", C_BRIGHT, 13, True)
        run_tb = _lbl(cv, fx+4, BOX_Y+BOX_H+22, "", C_WHITE,  11)
        box_lbls.append((b1_tb, b2_tb, b3_tb, sc_tb, run_tb))

    # Score, best score, difficulty in left panel
    total_lbl    = _lbl(cv, 10, 6,  "SCORE: 0",                     C_YELLOW, 12, True)
    best_lbl     = _lbl(cv, 10, 24, "BEST:  {}".format(best_score), C_YELLOW, 12, True)
    _lbl(cv,         10, 42, "DIFF: ",                               C_YELLOW, 12, True)
    diff_val_lbl = _lbl(cv, 62, 42, str(start_diff),                 C_YELLOW, 12, True)

    diff_lbl = [start_diff]  # mutable holder

    # --- Pins ---
    pin_tip_y = float(LANE_TOP + LANE_H_PX // 2)
    aim_min_y = float(LANE_TOP + 16)
    aim_max_y = float(LANE_BOT - 16)
    ball_start_y = float(LANE_TOP + LANE_H_PX // 2)

    pin_positions = build_pin_pos(pin_tip_y)
    pin_els = []
    for px2, py2 in pin_positions:
        el = _dot(cv, px2, py2, 10, C_PIN_UP, 3)
        pin_els.append(el)

    # --- Ball ---
    ball_el = _dot(cv, int(BALL_START_X), int(ball_start_y), 14, C_BALL, 4)

    # --- Aim line - horizontal line showing aim height ---
    AIM_LINE_END = PIN_TIP_X - 80   # stop before pin area
    aim_el = _line(cv,
                   int(BALL_START_X) + 30, int(ball_start_y),
                   AIM_LINE_END,            int(ball_start_y),
                   C_AIM, 1, 2)

    # ---------------------------------------------------------------------------
    # State
    # ---------------------------------------------------------------------------
    state = {
        "phase":          "aim",
        "frame":          0,
        "ball_in_frame":  0,
        "pins_up":        [True]*10,
        "rolls":          [],
        "aim_y":          ball_start_y,
        "aim_dir":        1,
        "ball_x":         BALL_START_X,
        "ball_y":         ball_start_y,
        "launched_aim":   0.5,
        "gutter":         False,
        "difficulty":     start_diff,
        "started":        False,
        "waiting":        False,
        "game_over":      False,
        "total":          0,
    }

    def redraw_pins():
        for i, el in enumerate(pin_els):
            el.Fill = C_PIN_UP if state["pins_up"][i] else C_PIN_DN

    def update_scorecard():
        frames = score_frames(state["rolls"])
        running = 0
        for f, (sc, b1, b2, b3) in enumerate(frames):
            b1_tb, b2_tb, b3_tb, sc_tb, run_tb = box_lbls[f]
            if f < 9 and b1 == 10:
                b1_tb.Text = ""; b2_tb.Text = "X"
                b2_tb.Foreground = C_YELLOW
            else:
                s1 = fmt_ball(b1, None, 0)
                s2 = fmt_ball(b2, b1,  1)
                b1_tb.Text = s1; b1_tb.Foreground = C_BRIGHT
                if b2_tb:
                    b2_tb.Text = s2
                    b2_tb.Foreground = C_YELLOW if s2 == "/" else C_BRIGHT
            if b3_tb:
                s3 = fmt_ball(b3, b2, 2)
                b3_tb.Text = s3
                b3_tb.Foreground = C_YELLOW if s3 in ("/","X") else C_BRIGHT
            if sc is not None:
                running += sc
                sc_tb.Text  = str(sc)
                run_tb.Text = str(running)
            else:
                sc_tb.Text = ""; run_tb.Text = ""
        state["total"] = running
        ctx.set_score(running)
        total_lbl.Text = "SCORE: {}".format(running)

    def place_ball(x, y):
        WpfCanvas.SetLeft(ball_el, x - 14)
        WpfCanvas.SetTop(ball_el,  y - 14)

    def reset_ball():
        state["ball_x"] = BALL_START_X
        # keep ball_y / aim_y at current position for next ball
        state["gutter"] = False
        place_ball(state["ball_x"], state["ball_y"])

    # ---------------------------------------------------------------------------
    # Collision
    # ---------------------------------------------------------------------------
    #
    # Pin numbering and positions (viewed from above, ball comes from left):
    #
    #   Row 0  Row 1  Row 2  Row 3
    #                         6(top)
    #              3(top)    5
    #   1(head) 2  4         8
    #              3(bot)    5... wait, standard layout:
    #
    # Standard 10-pin layout (ball arrives from left, pins spread right):
    #
    #   pin 1              (head, row 0, centre)
    #   pins 2,3           (row 1, top and bottom of centre)
    #   pins 4,5,6         (row 2)
    #   pins 7,8,9,10      (row 3, back row)
    #
    # pin_norm maps each pin to -1..+1 vertical space
    # (matching our build_pin_pos which spreads pins vertically within rows)
    #
    # Chain reaction rules:
    #   - A pin can only knock over its direct neighbours
    #   - Corner pins (6,7,10) are very hard to chain to
    #   - The ball must directly hit a pin for it to fall initially
    #
    def compute_knocked():
        aim = state["launched_aim"] * 2.0 - 1.0  # -1..+1, maps to vertical position

        # Each pin's vertical normalised position (-1=top, +1=bottom of lane)
        # Rows 0-3 left to right, within each row pins spread top to bottom
        pin_norm = [0.0]*10
        rows = [[0],[1,2],[3,4,5],[6,7,8,9]]
        for ri, pins in enumerate(rows):
            n = len(pins)
            for ci, pn in enumerate(pins):
                pin_norm[pn] = 0.0 if n==1 else -0.8 + ci*(1.6/(n-1))

        # How wide the ball path is - difficulty controls this
        # At diff=0 (pro): narrow, precise - 0.15 spread
        # At diff=60 (easy): wide, forgiving - 0.85 spread
        spread = 0.15 + state["difficulty"] * 0.012

        # Direct hits - ball physically contacts pins in its path
        # Head pin (0) is slightly easier to hit as it sticks out
        direct_hit = set()
        for i in range(10):
            if not state["pins_up"][i]: continue
            dist = abs(pin_norm[i] - aim)
            # Back row pins (6,7,8,9) are harder to hit directly -
            # the ball is deflected by front pins first
            row = next(ri for ri, pins in enumerate(rows) if i in pins)
            row_penalty = row * 0.06  # harder to hit back rows directly
            thresh = spread - row_penalty + (0.08 if i==0 else 0.0)
            thresh += random.uniform(0, 0.04)
            if thresh > 0 and dist < thresh:
                direct_hit.add(i)

        # Chain reaction - knocked pins deflect into neighbours
        # Uses directional physics: a pin deflects toward its neighbours
        # based on where it was hit
        # Adjacency with deflection probability weights
        # (lower = harder to chain to that pin)
        adj_weights = {
            0: {1: 0.85, 2: 0.85},
            1: {0: 0.5,  3: 0.80, 4: 0.75},
            2: {0: 0.5,  4: 0.75, 5: 0.80},
            3: {1: 0.5,  6: 0.65, 7: 0.60},
            4: {1: 0.5,  2: 0.5,  7: 0.55, 8: 0.55},
            5: {2: 0.5,  8: 0.55, 9: 0.65},
            6: {3: 0.3},   # corner - very hard to chain to
            7: {3: 0.35, 4: 0.30},
            8: {4: 0.30, 5: 0.35},
            9: {5: 0.3},   # corner - very hard to chain to
        }

        # Scale chain probabilities by difficulty
        diff_scale = 0.4 + state["difficulty"] * 0.010

        hit = set(direct_hit)
        changed = True
        while changed:
            changed = False
            for i in list(hit):
                for j, base_prob in adj_weights.get(i, {}).items():
                    if state["pins_up"][j] and j not in hit:
                        prob = base_prob * diff_scale
                        if random.random() < prob:
                            hit.add(j)
                            changed = True

        return list(hit)

    # ---------------------------------------------------------------------------
    # Process result
    # ---------------------------------------------------------------------------
    def process_result():
        if state["gutter"]:
            knocked = 0
            ctx.set_status("GUTTER BALL!  PRESS SPACE")
        else:
            hit = compute_knocked()
            knocked = len(hit)
            for i in hit:
                state["pins_up"][i] = False
            redraw_pins()

        state["rolls"].append(knocked)
        update_scorecard()

        f   = state["frame"]
        bal = state["ball_in_frame"]
        pins_down_total = sum(1 for p in state["pins_up"] if not p)

        if f < 9:
            if bal == 0:
                if knocked == 10:
                    ctx.set_status("STRIKE!!!  PRESS SPACE")
                    state["frame"] += 1; state["ball_in_frame"] = 0
                    state["pins_up"] = [True]*10
                else:
                    ctx.set_status("{} PIN{}  DOWN  |  PRESS SPACE FOR 2ND BALL".format(
                        knocked, "S" if knocked!=1 else ""))
                    state["ball_in_frame"] = 1
            else:
                if pins_down_total == 10:
                    ctx.set_status("SPARE!  PRESS SPACE")
                else:
                    ctx.set_status("{} DOWN THIS BALL  |  PRESS SPACE".format(knocked))
                state["frame"] += 1; state["ball_in_frame"] = 0
                state["pins_up"] = [True]*10
        else:
            if bal == 0:
                if knocked == 10:
                    ctx.set_status("STRIKE!  PRESS SPACE FOR BALL 2")
                    state["pins_up"] = [True]*10
                else:
                    ctx.set_status("{} DOWN  |  PRESS SPACE FOR BALL 2".format(knocked))
                state["ball_in_frame"] = 1
            elif bal == 1:
                r = state["rolls"]
                b1 = r[-2] if len(r)>=2 else 0
                if b1==10 or b1+knocked==10:
                    if b1==10: state["pins_up"] = [True]*10
                    ctx.set_status("PRESS SPACE FOR BONUS BALL")
                    state["ball_in_frame"] = 2
                else:
                    ctx.set_status("GAME OVER!  FINAL: {}  PRESS SPACE".format(state["total"]))
                    state["game_over"] = True
            else:
                ctx.set_status("GAME OVER!  FINAL: {}  PRESS SPACE".format(state["total"]))
                state["game_over"] = True

        state["phase"]   = "result"
        state["waiting"] = True

    # ---------------------------------------------------------------------------
    # Timers
    # ---------------------------------------------------------------------------
    BALL_SPEED = 10

    def get_aim_speed():
        return 2.0 + (60 - state["difficulty"]) * (5.0 / 60.0)

    # Aim oscillates UP/DOWN on left side of play area
    def aim_tick(sender, e):
        if state["phase"] != "aim": return
        state["aim_y"] += get_aim_speed() * state["aim_dir"]
        if state["aim_y"] >= aim_max_y: state["aim_y"] = aim_max_y; state["aim_dir"] = -1
        if state["aim_y"] <= aim_min_y: state["aim_y"] = aim_min_y; state["aim_dir"] =  1
        # Pulse aim line near centre
        centre = (aim_min_y + aim_max_y) / 2.0
        dist   = abs(state["aim_y"] - centre)
        aim_el.Stroke = C_AIM_HOT if dist < (aim_max_y-aim_min_y)*0.15 else C_AIM
        state["ball_y"] = state["aim_y"]
        aim_el.X1 = int(BALL_START_X) + 30
        aim_el.X2 = AIM_LINE_END
        aim_el.Y1 = state["ball_y"]
        aim_el.Y2 = state["ball_y"]
        place_ball(state["ball_x"], state["ball_y"])

    # Ball travels straight LEFT to RIGHT at constant speed
    def roll_tick(sender, e):
        if state["phase"] != "rolling": return
        state["ball_x"] += BALL_SPEED
        place_ball(state["ball_x"], state["ball_y"])
        if state["ball_x"] >= PIN_TIP_X:
            roll_timer.Stop()
            place_ball(-50, -50)  # hide ball
            process_result()
            reset_ball()

    aim_timer  = DispatcherTimer()
    aim_timer.Interval = TimeSpan.FromMilliseconds(18)
    aim_timer.Tick  += aim_tick

    roll_timer = DispatcherTimer()
    roll_timer.Interval = TimeSpan.FromMilliseconds(16)
    roll_timer.Tick += roll_tick

    # ---------------------------------------------------------------------------
    # Keyboard
    # ---------------------------------------------------------------------------
    def on_key(sender, e):
        k = e.Key
        if not state["started"]:
            if k == Key.Up:
                state["difficulty"] = min(60, state["difficulty"]+5)
                diff_lbl[0] = state["difficulty"]
                diff_val_lbl.Text = str(state["difficulty"])
                ctx.set_status("DIFF: {}  |  SPACE TO BOWL".format(state["difficulty"]))
                return
            if k == Key.Down:
                state["difficulty"] = max(0, state["difficulty"]-5)
                diff_lbl[0] = state["difficulty"]
                diff_val_lbl.Text = str(state["difficulty"])
                ctx.set_status("DIFF: {}  |  SPACE TO BOWL".format(state["difficulty"]))
                return

        if k != Key.Space: return

        if not state["started"]:
            state["started"] = True
            aim_timer.Start()
            ctx.set_status("FRAME 1  |  TIME YOUR AIM  |  SPACE TO RELEASE")
            return

        if state["waiting"]:
            state["waiting"] = False
            if state["game_over"]:
                if state["total"] > best_score:
                    _save_best(gd, state["total"])
                    best_lbl.Text = "BEST:  {}  NEW RECORD!".format(state["total"])
                    best_lbl.Foreground = C_YELLOW
                else:
                    best_lbl.Text = "BEST:  {}".format(best_score)
                ctx.set_status("GAME OVER!  SCORE: {}  BEST: {}  THANKS FOR PLAYING!".format(
                    state["total"], max(state["total"], best_score)))
                aim_timer.Stop()
                return
            state["phase"] = "aim"
            redraw_pins()
            # restore aim line
            aim_el.X1 = int(BALL_START_X) + 30
            aim_el.X2 = AIM_LINE_END
            aim_el.Y1 = state["ball_y"]
            aim_el.Y2 = state["ball_y"]
            aim_timer.Start()
            ctx.set_status("FRAME {}  BALL {}  |  SPACE TO BOWL".format(
                state["frame"]+1, state["ball_in_frame"]+1))
            return

        if state["phase"] == "aim":
            state["launched_aim"] = (state["ball_y"] - aim_min_y) / (aim_max_y - aim_min_y)
            state["launched_aim"] = max(0.0, min(1.0, state["launched_aim"]))
            state["phase"] = "rolling"
            aim_timer.Stop()
            # hide aim line
            aim_el.Y1 = -10; aim_el.Y2 = -10
            roll_timer.Start()
            ctx.set_status("ROLLING...")

    cv.KeyDown   += on_key
    cv.Focusable  = True
    cv.Focus()

    redraw_pins()
    update_scorecard()
    ctx.set_status("DIFF: {}  (0=PRO 60=EASY)  |  UP/DOWN TO CHANGE  |  SPACE TO START".format(start_diff))
    ctx.set_title(">> BOWLING <<")
