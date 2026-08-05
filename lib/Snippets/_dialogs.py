# -*- coding: utf-8 -*-
"""Seed43 shared dialogs: message, confirm, ask_string.

Borderless rounded dark card with a green accent bar and drop shadow, no
visible title bar, drag-movable by clicking anywhere on the card. The one
popup style every Seed43 tool should use instead of rolling its own.

    from _dialogs import message, confirm, ask_string

    message('Please select sheets for PDF.')
    if confirm('Delete profile "{}"?'.format(name), yes='Delete'):
        ...
    name = ask_string('Name for the new profile:', default=name, error=err)
"""
from pyrevit.framework import Windows

try:
    from Snippets.seed43_theme import apply_seed43_palette, get_color
except Exception:
    apply_seed43_palette = None
    get_color = None

# Fallbacks only, so a missing/unparseable seed43_palette.json never hard-
# fails this module. Once apply_seed43_palette() has run, _res() prefers the
# live Brush* tokens every other Seed43 XAML references.
BG      = '#2B3340'
HEADER  = '#232933'
GREEN   = '#208A3C'
GREENH  = '#27AE60'
TEXT    = '#F4FAFF'
RED     = '#C53030'
INPUTBG = '#1D1D20'


def _brush(hexcolor):
    return Windows.Media.SolidColorBrush(
        Windows.Media.ColorConverter.ConvertFromString(hexcolor))


def _res(w, key, fallback_hex):
    """Snapshot (not live-updating) brush for `key` - fine for a modal built
    and closed in one shot. Falls back to the literal hex above."""
    try:
        b = w.TryFindResource(key)
        if b:
            return b
    except Exception:
        pass
    return _brush(fallback_hex)


# pyTransmit's Settings/Dialogs.py calls this name. Kept as an alias rather
# than renamed, so both spellings work.
_theme_brush = _res


def _theme_dim(w, key, fallback):
    """Live sizing lookup (double / CornerRadius / Thickness). fallback must
    already be the right type for the property it is assigned to."""
    try:
        v = w.TryFindResource(key)
        if v is not None:
            return v
    except Exception:
        pass
    return fallback


def _textblock_label(w, text):
    """Small dim field label, used above folder/filename inputs."""
    t = Windows.Controls.TextBlock()
    t.Text = text
    t.Foreground = _res(w, 'BrushTextPrimary', TEXT)
    t.FontSize = 11
    t.Opacity = 0.7
    t.Margin = Windows.Thickness(0, 0, 0, 4)
    return t


def _plain_textbox(w):
    """An editable box styled with direct properties only.

    Deliberately NOT a custom ControlTemplate: the templated version this
    replaces rendered with no visible text, which is why Dialogs.save_log
    hand-rolls Border+TextBlock rather than reusing it. Same styling as
    ask_string's box, which has always worked.
    """
    tb = Windows.Controls.TextBox()
    tb.Height = 32
    tb.FontSize = 12
    tb.Padding = Windows.Thickness(8, 4, 8, 4)
    tb.Background = _res(w, 'BrushInputBg', TEXT)
    tb.Foreground = _res(w, 'BrushTextInput', HEADER)
    tb.BorderBrush = _res(w, 'BrushBorderDefault', GREEN)
    tb.BorderThickness = Windows.Thickness(1.5)
    tb.Margin = Windows.Thickness(0, 0, 0, 16)
    return tb


def _button(w, text, primary=True):
    bg = _res(w, 'BrushPrimaryGreen' if primary else 'BrushSecondaryBg', GREEN if primary else HEADER)
    border = _res(w, 'BrushPrimaryGreen', GREEN)
    fg = _res(w, 'BrushTextPrimary', TEXT)

    # Hover must be a literal hex baked into the XAML template below, so read
    # the palette JSON directly rather than resolving a Brush and converting
    # back - no dependency on TryFindResource returning a usable .Color.
    hover_fallback = GREENH if primary else '#333B48'
    if get_color:
        hover_hex = get_color(None, 'accent_hover' if primary else 'secondary_hover', fallback=hover_fallback)
    else:
        hover_hex = hover_fallback

    b = Windows.Controls.Button()
    b.Content = text
    b.MinWidth = 84
    b.Height = 32
    b.Margin = Windows.Thickness(8, 0, 0, 0)
    b.Background = bg
    b.Foreground = fg
    b.BorderBrush = border
    b.BorderThickness = Windows.Thickness(0 if primary else 1)
    b.FontWeight = Windows.FontWeights.SemiBold
    b.FontSize = 12
    b.Cursor = Windows.Input.Cursors.Hand

    import System.Windows.Markup as _Markup
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
    ).format(hover=hover_hex)
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
    if apply_seed43_palette:
        try:
            apply_seed43_palette(w, None)
        except Exception:
            pass
    return w


def _card(w):
    """Rounded dark card with a green accent bar, drop shadow, drag-movable."""
    outer = Windows.Controls.Border()
    outer.Background = _res(w, 'BrushCardBg', BG)
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
    accent.Background = _res(w, 'BrushPrimaryGreen', GREEN)
    accent.Height = 3
    accent.CornerRadius = Windows.CornerRadius(2)
    accent.Margin = Windows.Thickness(0, 0, 0, 16)
    root.Children.Add(accent)
    outer.Child = root
    w.Content = outer
    return root


def _textblock(w, text, error=False, title=False):
    t = Windows.Controls.TextBlock()
    t.Text = text
    t.Foreground = _res(w, 'BrushDanger' if error else 'BrushTextPrimary', RED if error else TEXT)
    t.FontSize = 15 if title else 12
    t.FontWeight = Windows.FontWeights.Bold if title else Windows.FontWeights.Normal
    t.Opacity = 1.0 if title else 0.85
    t.TextWrapping = Windows.TextWrapping.Wrap
    t.Margin = Windows.Thickness(0, 0, 0, 8 if title else 20)
    return t


def message(text, title='', ok_label='OK'):
    """Themed OK-only info popup. Escape closes it, same as OK."""
    w = _window()
    root = _card(w)
    if title:
        root.Children.Add(_textblock(w, title, title=True))
    root.Children.Add(_textblock(w, text))
    row = Windows.Controls.StackPanel()
    row.Orientation = Windows.Controls.Orientation.Horizontal
    row.HorizontalAlignment = Windows.HorizontalAlignment.Right
    ok = _button(w, ok_label)
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
        root.Children.Add(_textblock(w, title, title=True))
    root.Children.Add(_textblock(w, text))
    row = Windows.Controls.StackPanel()
    row.Orientation = Windows.Controls.Orientation.Horizontal
    row.HorizontalAlignment = Windows.HorizontalAlignment.Right

    def _yes(s, a):
        result['ok'] = True
        w.Close()
    nb = _button(w, no, primary=False)
    nb.Click += lambda s, a: w.Close()
    yb = _button(w, yes)
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
        root.Children.Add(_textblock(w, title, title=True))
    root.Children.Add(_textblock(w, prompt))
    if error:
        err = _textblock(w, error, error=True)
        err.Margin = Windows.Thickness(0, -12, 0, 12)
        root.Children.Add(err)
    tb = Windows.Controls.TextBox()
    tb.Text = default or ''
    tb.Height = 32
    tb.FontSize = 12
    tb.Padding = Windows.Thickness(8, 4, 8, 4)
    tb.Background = _res(w, 'BrushInputBg', TEXT)
    tb.Foreground = _res(w, 'BrushTextInput', HEADER)
    tb.BorderBrush = _res(w, 'BrushBorderDefault', GREEN)
    tb.BorderThickness = Windows.Thickness(1.5)
    tb.Margin = Windows.Thickness(0, 0, 0, 16)
    root.Children.Add(tb)
    row = Windows.Controls.StackPanel()
    row.Orientation = Windows.Controls.Orientation.Horizontal
    row.HorizontalAlignment = Windows.HorizontalAlignment.Right

    def _ok(s, a):
        result['text'] = tb.Text
        w.Close()
    ca = _button(w, 'Cancel', primary=False)
    ca.Click += lambda s, a: w.Close()
    ok = _button(w, 'OK')
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


def choice(text, options, title='', detail_text=''):
    """Themed multi-button choice popup, for when confirm()'s two aren't
    enough. `options` is a list of (key, label) tuples; the last is rendered
    as the primary/green button, the rest secondary, left to right. Returns
    the chosen key, or None if dismissed without choosing.

    detail_text, if given, adds a read-only preview panel below the buttons -
    a diff or log excerpt. Omitted entirely when empty.
    """
    w = _window()
    root = _card(w)
    result = {'key': None}
    if title:
        root.Children.Add(_textblock(w, title, title=True))
    root.Children.Add(_textblock(w, text))

    row = Windows.Controls.StackPanel()
    row.Orientation = Windows.Controls.Orientation.Horizontal
    row.HorizontalAlignment = Windows.HorizontalAlignment.Right

    def _make_pick(key):
        def _pick(s, a):
            result['key'] = key
            w.Close()
        return _pick

    last_index = len(options) - 1
    for i, (key, label) in enumerate(options):
        btn = _button(w, label, primary=(i == last_index))
        btn.Click += _make_pick(key)
        row.Children.Add(btn)
    root.Children.Add(row)

    if detail_text:
        panel = Windows.Controls.Border()
        panel.Background = _res(w, 'BrushHeaderBg', HEADER)
        panel.CornerRadius = _theme_dim(w, 'CornerRadiusInput',
                                        Windows.CornerRadius(6))
        panel.Padding = Windows.Thickness(12)
        panel.Margin = Windows.Thickness(0, 16, 0, 0)
        detail = Windows.Controls.TextBlock()
        detail.Text = detail_text
        detail.Foreground = _res(w, 'BrushTextPrimary', TEXT)
        detail.FontSize = 11
        detail.Opacity = 0.85
        detail.TextWrapping = Windows.TextWrapping.Wrap
        detail.FontFamily = Windows.Media.FontFamily('Consolas')
        panel.Child = detail
        root.Children.Add(panel)

    w.ShowDialog()
    return result['key']


def save_file_as(title, filename, ext, initial_folder=None):
    """Themed 'save as': a folder with a native browse button, plus an
    editable filename. Returns the full path, or None if cancelled.

    The folder box is read-only on purpose - it is set via Browse, so there
    is nothing to gain from letting a path be typed in wrong.
    """
    import os as _os

    w = _window(width=480)
    root = _card(w)
    result = {'path': None}
    root.Children.Add(_textblock(w, title, title=True))

    root.Children.Add(_textblock_label(w, 'Folder'))
    folder_row = Windows.Controls.Grid()
    c0 = Windows.Controls.ColumnDefinition()
    c0.Width = Windows.GridLength(1, Windows.GridUnitType.Star)
    c1 = Windows.Controls.ColumnDefinition()
    c1.Width = Windows.GridLength.Auto
    folder_row.ColumnDefinitions.Add(c0)
    folder_row.ColumnDefinitions.Add(c1)

    folder_tb = _plain_textbox(w)
    folder_tb.IsReadOnly = True
    folder_tb.Margin = Windows.Thickness(0, 0, 6, 12)
    Windows.Controls.Grid.SetColumn(folder_tb, 0)

    browse_btn = _button(w, 'Browse')
    browse_btn.Margin = Windows.Thickness(0, 0, 0, 12)
    Windows.Controls.Grid.SetColumn(browse_btn, 1)

    folder_row.Children.Add(folder_tb)
    folder_row.Children.Add(browse_btn)
    root.Children.Add(folder_row)

    root.Children.Add(_textblock_label(w, 'File name'))
    filename_tb = _plain_textbox(w)
    root.Children.Add(filename_tb)

    base = filename or ''
    if base.lower().endswith('.' + ext.lower()):
        base = base[:-(len(ext) + 1)]
    filename_tb.Text = base

    desktop = _os.path.join(_os.path.expanduser('~'), 'Desktop')
    folder_tb.Text = (initial_folder
                      if (initial_folder and _os.path.isdir(initial_folder))
                      else desktop)

    def on_browse(s, a):
        try:
            from System.Windows.Forms import FolderBrowserDialog, DialogResult
            fb = FolderBrowserDialog()
            fb.SelectedPath = folder_tb.Text or desktop
            if fb.ShowDialog() == DialogResult.OK:
                folder_tb.Text = fb.SelectedPath
        except Exception:
            pass
    browse_btn.Click += on_browse

    row = Windows.Controls.StackPanel()
    row.Orientation = Windows.Controls.Orientation.Horizontal
    row.HorizontalAlignment = Windows.HorizontalAlignment.Right

    def on_save(s, a):
        folder = (folder_tb.Text or '').strip()
        name = (filename_tb.Text or '').strip()
        if not name:
            return
        if not name.lower().endswith('.' + ext.lower()):
            name = name + '.' + ext
        result['path'] = _os.path.join(folder, name)
        w.Close()

    cancel_btn = _button(w, 'Cancel', primary=False)
    cancel_btn.Click += lambda s, a: w.Close()
    save_btn = _button(w, 'Save')
    save_btn.Click += on_save
    row.Children.Add(cancel_btn)
    row.Children.Add(save_btn)
    root.Children.Add(row)

    def _key(s, a):
        if a.Key == Windows.Input.Key.Escape:
            w.Close()
    w.KeyDown += _key

    w.ShowDialog()
    return result['path']
