# -*- coding: utf-8 -*-
"""Seed43 shared dialogs: message, confirm, ask_string.

Borderless rounded dark card with a green accent bar and drop shadow, no
visible title bar, drag-movable by clicking anywhere on the card. This is
the one popup style every Seed43 tool should use instead of building its
own message box.

Install: drop this file in Seed43.extension/lib/Snippets/ as _dialogs.py,
then from any pushbutton:

    from _dialogs import message, confirm, ask_string

    message('Please select sheets for PDF.')
    if confirm('Delete profile "{}"?'.format(name), yes='Delete'):
        ...
    name = ask_string('Name for the new profile:', default=name, error=err)
"""
from pyrevit.framework import Windows

BG     = '#2B3340'
HEADER = '#232933'
GREEN  = '#208A3C'
GREENH = '#27AE60'
TEXT   = '#F4FAFF'
RED    = '#C53030'


def _brush(hexcolor):
    return Windows.Media.SolidColorBrush(
        Windows.Media.ColorConverter.ConvertFromString(hexcolor))


def _button(text, primary=True):
    b = Windows.Controls.Button()
    b.Content = text
    b.MinWidth = 84
    b.Height = 32
    b.Margin = Windows.Thickness(8, 0, 0, 0)
    b.Background = _brush(GREEN if primary else HEADER)
    b.Foreground = _brush(TEXT)
    b.BorderBrush = _brush(GREEN)
    b.BorderThickness = Windows.Thickness(0 if primary else 1)
    b.FontWeight = Windows.FontWeights.SemiBold
    b.FontSize = 12
    b.Cursor = Windows.Input.Cursors.Hand

    import System.Windows.Markup as _Markup
    hover = GREENH if primary else '#333B48'
    template_xaml = (
        '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
        'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" '
        'TargetType="Button">'
        '<Border x:Name="Bd" Background="{{TemplateBinding Background}}" '
        'BorderBrush="{{TemplateBinding BorderBrush}}" '
        'BorderThickness="{{TemplateBinding BorderThickness}}" '
        'CornerRadius="6" Padding="14,0">'
        '<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>'
        '</Border>'
        '<ControlTemplate.Triggers>'
        '<Trigger Property="IsMouseOver" Value="True">'
        '<Setter TargetName="Bd" Property="Background" Value="{hover}"/>'
        '</Trigger>'
        '</ControlTemplate.Triggers>'
        '</ControlTemplate>'
    ).format(hover=hover)
    b.Template = _Markup.XamlReader.Parse(template_xaml)
    return b


def _window(width=420):
    w = Windows.Window()
    w.Width = width
    w.SizeToContent = Windows.SizeToContent.Height
    w.WindowStartupLocation = Windows.WindowStartupLocation.CenterScreen
    w.ResizeMode = Windows.ResizeMode.NoResize
    w.WindowStyle = getattr(Windows.WindowStyle, 'None')  # 'None' is a Py keyword
    w.AllowsTransparency = True
    w.Background = _brush('Transparent')
    w.ShowInTaskbar = False
    return w


def _card(w):
    """Rounded dark card with a green accent bar, drop shadow, drag-movable."""
    outer = Windows.Controls.Border()
    outer.Background = _brush(BG)
    outer.CornerRadius = Windows.CornerRadius(10)
    outer.Margin = Windows.Thickness(12)
    outer.Padding = Windows.Thickness(24, 20, 24, 20)

    shadow = Windows.Media.Effects.DropShadowEffect()
    shadow.Color = Windows.Media.Colors.Black
    shadow.Opacity = 0.5
    shadow.ShadowDepth = 4
    shadow.BlurRadius = 16
    outer.Effect = shadow

    def _drag(s, a):
        if a.ButtonState == Windows.Input.MouseButtonState.Pressed:
            w.DragMove()
    outer.MouseLeftButtonDown += _drag

    root = Windows.Controls.StackPanel()
    accent = Windows.Controls.Border()
    accent.Background = _brush(GREEN)
    accent.Height = 3
    accent.CornerRadius = Windows.CornerRadius(2)
    accent.Margin = Windows.Thickness(0, 0, 0, 16)
    root.Children.Add(accent)
    outer.Child = root
    w.Content = outer
    return root


def _textblock(text, error=False, title=False):
    t = Windows.Controls.TextBlock()
    t.Text = text
    t.Foreground = _brush(RED if error else TEXT)
    t.FontSize = 15 if title else 12
    t.FontWeight = Windows.FontWeights.Bold if title else Windows.FontWeights.Normal
    t.Opacity = 1.0 if title else 0.85
    t.TextWrapping = Windows.TextWrapping.Wrap
    t.Margin = Windows.Thickness(0, 0, 0, 8 if title else 20)
    return t


def message(text, title=''):
    """Themed OK-only info popup. Escape closes it, same as OK."""
    w = _window()
    root = _card(w)
    if title:
        root.Children.Add(_textblock(title, title=True))
    root.Children.Add(_textblock(text))
    row = Windows.Controls.StackPanel()
    row.Orientation = Windows.Controls.Orientation.Horizontal
    row.HorizontalAlignment = Windows.HorizontalAlignment.Right
    ok = _button('OK')
    ok.Click += lambda s, a: w.Close()
    row.Children.Add(ok)
    root.Children.Add(row)

    def _key(s, a):
        if a.Key == Windows.Input.Key.Escape:
            w.Close()
    w.KeyDown += _key

    w.ShowDialog()


def confirm(text, title='', yes='Yes', no='Cancel'):
    """Themed yes/cancel popup, returns True on yes."""
    w = _window()
    root = _card(w)
    result = {'ok': False}
    if title:
        root.Children.Add(_textblock(title, title=True))
    root.Children.Add(_textblock(text))
    row = Windows.Controls.StackPanel()
    row.Orientation = Windows.Controls.Orientation.Horizontal
    row.HorizontalAlignment = Windows.HorizontalAlignment.Right

    def _yes(s, a):
        result['ok'] = True
        w.Close()
    nb = _button(no, primary=False)
    nb.Click += lambda s, a: w.Close()
    yb = _button(yes)
    yb.Click += _yes
    row.Children.Add(nb)
    row.Children.Add(yb)
    root.Children.Add(row)
    w.ShowDialog()
    return result['ok']


def ask_string(prompt, title='', default='', error=''):
    """Themed single-line text prompt, returns the string or None on cancel."""
    w = _window()
    root = _card(w)
    result = {'text': None}
    if title:
        root.Children.Add(_textblock(title, title=True))
    root.Children.Add(_textblock(prompt))
    if error:
        err = _textblock(error, error=True)
        err.Margin = Windows.Thickness(0, -12, 0, 12)
        root.Children.Add(err)
    tb = Windows.Controls.TextBox()
    tb.Text = default or ''
    tb.Height = 32
    tb.FontSize = 12
    tb.Padding = Windows.Thickness(8, 4, 8, 4)
    tb.Background = _brush(TEXT)
    tb.Foreground = _brush(HEADER)
    tb.BorderBrush = _brush(GREEN)
    tb.BorderThickness = Windows.Thickness(1.5)
    tb.Margin = Windows.Thickness(0, 0, 0, 16)
    root.Children.Add(tb)
    row = Windows.Controls.StackPanel()
    row.Orientation = Windows.Controls.Orientation.Horizontal
    row.HorizontalAlignment = Windows.HorizontalAlignment.Right

    def _ok(s, a):
        result['text'] = tb.Text
        w.Close()
    ca = _button('Cancel', primary=False)
    ca.Click += lambda s, a: w.Close()
    ok = _button('OK')
    ok.Click += _ok
    row.Children.Add(ca)
    row.Children.Add(ok)
    root.Children.Add(row)

    def _key(s, a):
        if a.Key == Windows.Input.Key.Enter:
            _ok(s, a)
        elif a.Key == Windows.Input.Key.Escape:
            w.Close()
    tb.KeyDown += _key
    w.ContentRendered += lambda s, a: (tb.Focus(), tb.SelectAll())
    w.ShowDialog()
    return result['text']
