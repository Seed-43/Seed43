# -*- coding: utf-8 -*-
"""Seed43-styled dialogs: message, confirm, and text input."""
from pyrevit.framework import Windows

BG      = '#2B3340'
HEADER  = '#232933'
GREEN   = '#208A3C'
GREENH  = '#27AE60'
TEXT    = '#F4FAFF'
RED     = '#C53030'


def _brush(hexcolor):
    return Windows.Media.SolidColorBrush(
        Windows.Media.ColorConverter.ConvertFromString(hexcolor))


def _button(text, primary=True):
    b = Windows.Controls.Button()
    b.Content = text
    b.MinWidth = 80
    b.Height = 30
    b.Margin = Windows.Thickness(6, 0, 0, 0)
    b.Background = _brush(GREEN if primary else HEADER)
    b.Foreground = _brush(TEXT)
    b.BorderBrush = _brush(GREEN)
    b.BorderThickness = Windows.Thickness(0 if primary else 1)
    b.FontWeight = Windows.FontWeights.SemiBold
    b.Cursor = Windows.Input.Cursors.Hand
    return b


def _window(title, height):
    w = Windows.Window()
    w.Title = title
    w.Width = 420
    w.Height = height
    w.WindowStartupLocation = \
        Windows.WindowStartupLocation.CenterScreen
    w.ResizeMode = Windows.ResizeMode.NoResize
    w.Background = _brush(BG)
    w.WindowStyle = Windows.WindowStyle.ToolWindow
    return w


def _textblock(text, error=False):
    t = Windows.Controls.TextBlock()
    t.Text = text
    t.Foreground = _brush(RED if error else TEXT)
    t.FontSize = 12
    t.TextWrapping = Windows.TextWrapping.Wrap
    t.Margin = Windows.Thickness(0, 0, 0, 10)
    return t


def message(text, title='pySheets'):
    """Styled message box."""
    w = _window(title, 160)
    panel = Windows.Controls.StackPanel()
    panel.Margin = Windows.Thickness(16)
    panel.Children.Add(_textblock(text))
    row = Windows.Controls.StackPanel()
    row.Orientation = Windows.Controls.Orientation.Horizontal
    row.HorizontalAlignment = Windows.HorizontalAlignment.Right
    ok = _button('OK')
    ok.Click += lambda s, a: w.Close()
    row.Children.Add(ok)
    panel.Children.Add(row)
    w.Content = panel
    w.ShowDialog()


def confirm(text, title='pySheets', yes='Yes', no='Cancel'):
    """Styled yes/no. Returns True on yes."""
    w = _window(title, 170)
    result = {'ok': False}
    panel = Windows.Controls.StackPanel()
    panel.Margin = Windows.Thickness(16)
    panel.Children.Add(_textblock(text))
    row = Windows.Controls.StackPanel()
    row.Orientation = Windows.Controls.Orientation.Horizontal
    row.HorizontalAlignment = Windows.HorizontalAlignment.Right

    def _yes(s, a):
        result['ok'] = True
        w.Close()
    yb = _button(yes)
    yb.Click += _yes
    nb = _button(no, primary=False)
    nb.Click += lambda s, a: w.Close()
    row.Children.Add(nb)
    row.Children.Add(yb)
    panel.Children.Add(row)
    w.Content = panel
    w.ShowDialog()
    return result['ok']


def ask_string(prompt, title='pySheets', default='', error=''):
    """Styled text input. Returns the string or None on cancel."""
    w = _window(title, 210 if error else 190)
    result = {'text': None}
    panel = Windows.Controls.StackPanel()
    panel.Margin = Windows.Thickness(16)
    panel.Children.Add(_textblock(prompt))
    if error:
        panel.Children.Add(_textblock(error, error=True))
    tb = Windows.Controls.TextBox()
    tb.Text = default or ''
    tb.Height = 30
    tb.FontSize = 12
    tb.Padding = Windows.Thickness(6, 4, 6, 4)
    tb.Background = _brush(TEXT)
    tb.Foreground = _brush(HEADER)
    tb.BorderBrush = _brush(GREEN)
    tb.BorderThickness = Windows.Thickness(1.5)
    tb.Margin = Windows.Thickness(0, 0, 0, 12)
    panel.Children.Add(tb)
    row = Windows.Controls.StackPanel()
    row.Orientation = Windows.Controls.Orientation.Horizontal
    row.HorizontalAlignment = Windows.HorizontalAlignment.Right

    def _ok(s, a):
        result['text'] = tb.Text
        w.Close()
    ok = _button('OK')
    ok.Click += _ok
    ca = _button('Cancel', primary=False)
    ca.Click += lambda s, a: w.Close()
    row.Children.Add(ca)
    row.Children.Add(ok)
    panel.Children.Add(row)

    def _key(s, a):
        if a.Key == Windows.Input.Key.Enter:
            _ok(s, a)
        elif a.Key == Windows.Input.Key.Escape:
            w.Close()
    tb.KeyDown += _key
    w.Content = panel
    tb.Focus()
    tb.SelectAll()
    w.ShowDialog()
    return result['text']
