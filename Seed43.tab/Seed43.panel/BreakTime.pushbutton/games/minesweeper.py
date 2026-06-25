# -*- coding: utf-8 -*-
# Minesweeper - BreakTime game module

from System.Windows.Threading import DispatcherTimer
from System.Windows.Input import Key, MouseButton
from System.Windows.Shapes import Rectangle
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Controls import TextBlock, Canvas as WpfCanvas, Border
from System.Windows import Thickness, FontWeight, FontWeights, HorizontalAlignment, VerticalAlignment
from System import TimeSpan
import random

COLS  = 20
ROWS  = 16
MINES = 40
CELL  = 28

C_HIDDEN  = Color.FromRgb(0,   60,  0)
C_HOVER   = Color.FromRgb(0,   100, 0)
C_OPEN    = Color.FromRgb(0,   20,  0)
C_MINE    = Color.FromRgb(255, 51,  0)
C_FLAG    = Color.FromRgb(255, 200, 0)
C_BORDER  = Color.FromRgb(0,   180, 40)
C_TEXT    = [
    None,
    Color.FromRgb(0,   255, 65),   # 1
    Color.FromRgb(0,   200, 200),  # 2
    Color.FromRgb(255, 150, 0),    # 3
    Color.FromRgb(255, 80,  80),   # 4
    Color.FromRgb(200, 0,   0),    # 5
    Color.FromRgb(0,   200, 200),  # 6
    Color.FromRgb(200, 200, 200),  # 7
    Color.FromRgb(150, 150, 150),  # 8
]

def _brush(c):
    return SolidColorBrush(c)

def _font():
    from System.Windows.Media import FontFamily
    return FontFamily("Courier New")

def init_game(ctx, timer):
    canvas = ctx.canvas
    W = COLS * CELL + 2
    H = ROWS * CELL + 2
    canvas.Width  = W
    canvas.Height = H

    board = {
        "mines":    set(),
        "revealed": [[False]*COLS for _ in range(ROWS)],
        "flagged":  [[False]*COLS for _ in range(ROWS)],
        "counts":   [[0]*COLS for _ in range(ROWS)],
        "cells":    {},   # (r,c) -> (bg_rect, text_block)
        "started":  False,
        "alive":    True,
        "flags_left": MINES,
    }

    def place_mines(exclude_r, exclude_c):
        safe = set()
        for dr in range(-1, 2):
            for dc in range(-1, 2):
                safe.add((exclude_r+dr, exclude_c+dc))
        candidates = [(r,c) for r in range(ROWS) for c in range(COLS) if (r,c) not in safe]
        chosen = random.sample(candidates, MINES)
        board["mines"] = set(chosen)
        for r in range(ROWS):
            for c in range(COLS):
                if (r,c) not in board["mines"]:
                    cnt = sum(1 for dr in (-1,0,1) for dc in (-1,0,1)
                              if (r+dr, c+dc) in board["mines"])
                    board["counts"][r][c] = cnt

    def draw_cell(r, c):
        bg, tb = board["cells"][(r,c)]
        revealed = board["revealed"][r][c]
        flagged  = board["flagged"][r][c]
        is_mine  = (r,c) in board["mines"]
        cnt      = board["counts"][r][c]

        if revealed:
            if is_mine:
                bg.Fill = _brush(C_MINE)
                tb.Text = "X"
                tb.Foreground = _brush(Color.FromRgb(255,255,255))
            else:
                bg.Fill = _brush(C_OPEN)
                if cnt > 0:
                    tb.Text       = str(cnt)
                    tb.Foreground = _brush(C_TEXT[cnt])
                else:
                    tb.Text = ""
        elif flagged:
            bg.Fill = _brush(C_HIDDEN)
            tb.Text       = "F"
            tb.Foreground = _brush(C_FLAG)
        else:
            bg.Fill = _brush(C_HIDDEN)
            tb.Text = ""

    def reveal(r, c):
        if r < 0 or r >= ROWS or c < 0 or c >= COLS:
            return
        if board["revealed"][r][c] or board["flagged"][r][c]:
            return
        board["revealed"][r][c] = True
        draw_cell(r, c)
        if board["counts"][r][c] == 0 and (r,c) not in board["mines"]:
            for dr in (-1,0,1):
                for dc in (-1,0,1):
                    if dr != 0 or dc != 0:
                        reveal(r+dr, c+dc)

    def check_win():
        for r in range(ROWS):
            for c in range(COLS):
                if (r,c) not in board["mines"] and not board["revealed"][r][c]:
                    return False
        return True

    def reveal_all_mines():
        for (r,c) in board["mines"]:
            board["revealed"][r][c] = True
            draw_cell(r,c)

    # Build cell grid
    for r in range(ROWS):
        for c in range(COLS):
            x = c * CELL + 1
            y = r * CELL + 1
            bg = Rectangle()
            bg.Width  = CELL - 1
            bg.Height = CELL - 1
            bg.Fill   = _brush(C_HIDDEN)
            bg.Stroke = _brush(C_BORDER)
            bg.StrokeThickness = 0.5
            WpfCanvas.SetLeft(bg, x)
            WpfCanvas.SetTop(bg, y)
            WpfCanvas.SetZIndex(bg, 0)

            tb = TextBlock()
            tb.FontFamily  = _font()
            tb.FontSize    = 14
            tb.FontWeight  = FontWeights.Bold
            tb.Text        = ""
            tb.Foreground  = _brush(C_TEXT[1])
            tb.Width       = CELL - 1
            tb.Height      = CELL - 1
            tb.TextAlignment = System_TextAlignment()
            tb.VerticalAlignment = VerticalAlignment.Center
            WpfCanvas.SetLeft(tb, x)
            WpfCanvas.SetTop(tb, y + 4)
            WpfCanvas.SetZIndex(tb, 1)

            canvas.Children.Add(bg)
            canvas.Children.Add(tb)
            board["cells"][(r,c)] = (bg, tb)

    def on_mouse_left(sender, e):
        if not board["alive"]:
            restart()
            return
        pos = e.GetPosition(canvas)
        c = int(pos.X // CELL)
        r = int(pos.Y // CELL)
        if r < 0 or r >= ROWS or c < 0 or c >= COLS:
            return
        if board["flagged"][r][c] or board["revealed"][r][c]:
            return

        if not board["started"]:
            place_mines(r, c)
            board["started"] = True

        if (r,c) in board["mines"]:
            board["revealed"][r][c] = True
            draw_cell(r,c)
            reveal_all_mines()
            board["alive"] = False
            ctx.set_status("BOOM!  YOU HIT A MINE!  CLICK TO RESTART")
            return

        reveal(r, c)
        if check_win():
            board["alive"] = False
            ctx.set_status("YOU WIN!  ALL MINES FOUND!  CLICK TO RESTART")
            ctx.set_score(ctx.score + 100)

    def on_mouse_right(sender, e):
        if not board["alive"]:
            return
        pos = e.GetPosition(canvas)
        c = int(pos.X // CELL)
        r = int(pos.Y // CELL)
        if r < 0 or r >= ROWS or c < 0 or c >= COLS:
            return
        if board["revealed"][r][c]:
            return
        board["flagged"][r][c] = not board["flagged"][r][c]
        if board["flagged"][r][c]:
            board["flags_left"] -= 1
        else:
            board["flags_left"] += 1
        draw_cell(r,c)
        ctx.set_status("FLAGS REMAINING: {}".format(board["flags_left"]))

    def restart():
        # Reset all state first
        board["mines"]      = set()
        board["revealed"]   = [[False]*COLS for _ in range(ROWS)]
        board["flagged"]    = [[False]*COLS for _ in range(ROWS)]
        board["counts"]     = [[0]*COLS for _ in range(ROWS)]
        board["cells"]      = {}
        board["started"]    = False
        board["alive"]      = True
        board["flags_left"] = MINES
        # Clear and rebuild canvas
        canvas.Children.Clear()
        for r in range(ROWS):
            for c in range(COLS):
                x = c * CELL + 1
                y = r * CELL + 1
                bg = Rectangle()
                bg.Width  = CELL - 1; bg.Height = CELL - 1
                bg.Fill   = _brush(C_HIDDEN)
                bg.Stroke = _brush(C_BORDER); bg.StrokeThickness = 0.5
                WpfCanvas.SetLeft(bg, x); WpfCanvas.SetTop(bg, y); WpfCanvas.SetZIndex(bg, 0)
                tb = TextBlock()
                tb.FontFamily = _font(); tb.FontSize = 14; tb.FontWeight = FontWeights.Bold
                tb.Text = ""; tb.Foreground = _brush(C_TEXT[1])
                tb.Width = CELL-1; tb.Height = CELL-1
                tb.TextAlignment = System_TextAlignment()
                tb.VerticalAlignment = VerticalAlignment.Center
                WpfCanvas.SetLeft(tb, x); WpfCanvas.SetTop(tb, y+4); WpfCanvas.SetZIndex(tb, 1)
                canvas.Children.Add(bg); canvas.Children.Add(tb)
                board["cells"][(r,c)] = (bg, tb)
        ctx.set_status("LEFT CLICK: REVEAL  |  RIGHT CLICK: FLAG  |  {} MINES".format(MINES))
        canvas.Focus()

    canvas.MouseLeftButtonDown  += on_mouse_left
    canvas.MouseRightButtonDown += on_mouse_right
    canvas.Focusable = True
    canvas.Focus()
    ctx.set_status("LEFT CLICK: REVEAL  |  RIGHT CLICK: FLAG  |  {} MINES".format(MINES))
    ctx.set_title(">> MINESWEEPER <<")

def System_TextAlignment():
    from System.Windows import TextAlignment
    return TextAlignment.Center
