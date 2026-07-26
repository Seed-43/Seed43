# -*- coding: utf-8 -*-
"""Seed43 shared dialogs: message, confirm, ask_string.

Borderless rounded dark card with a green accent bar, a themed close (X)
button, and a drop shadow. No visible title bar; drag-movable by clicking
anywhere on the card background. This is the one popup style every Seed43
tool should use instead of building its own message box.

Colours and sizing are pulled live from UI/seed43_palette.json via
seed43_theme, the same as every tool's own window - so this dialog reskins
itself automatically whenever the shared palette changes, instead of
carrying its own frozen copy of the colours. The module-level constants
below are only a fallback for if the palette file can't be found.

Install: drop this file in Seed43.extension/lib/Snippets/ as _dialogs.py,
then from any pushbutton:

    from Snippets._dialogs import message, confirm, ask_string, choice, save_file_as

    message('Please select sheets for PDF.')
    if confirm('Delete profile "{}"?'.format(name), yes='Delete'):
        ...
    name = ask_string('Name for the new profile:', default=name, error=err)
    result = choice('Settings differ from the last issue.',
                     [('ignore', 'Ignore'), ('session', 'This Issue Only'),
                      ('update', 'Update Settings')], title='Settings Mismatch')
    path = save_file_as('Save Log File', 'log.zip', 'zip')
"""
import os

from pyrevit.framework import Windows

try:
    from Snippets.seed43_theme import apply_seed43_palette, apply_seed43_dimensions, get_color
except Exception:
    apply_seed43_palette = None
    apply_seed43_dimensions = None
    get_color = None

try:
    from Snippets._icons import make_icon as _make_icon
except Exception:
    _make_icon = None

# Fallback-only defaults (current dark profile) - used if the palette file
# can't be found, same principle as every tool's own XAML fallback Setters.
BG       = '#2B3340'
HEADER   = '#232933'
GREEN    = '#208A3C'
GREENH   = '#32934C'
GREENP   = '#5CAA71'
TEXT     = '#F4FAFF'
RED      = '#E01B24'
REDH     = '#E22D36'
SECOND   = '#48484B'
SECONDH  = '#515155'
INPUTBG  = '#1D1D20'

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def _brush(hexcolor):
    return Windows.Media.SolidColorBrush(
        Windows.Media.ColorConverter.ConvertFromString(hexcolor))


def _theme_brush(w, key, fallback_hex):
    """Live palette lookup with a hardcoded fallback, same pattern used by
    every tool's own code-built UI (see gotcha #3 in seed43-pyrevit-ui)."""
    try:
        b = w.TryFindResource(key)
        if b:
            return b
    except Exception:
        pass
    return _brush(fallback_hex)


def _brush_hex(brush):
    """A live theme brush as a hex string, for anywhere that needs one
    (make_icon's colour param, or building a hover colour into a parsed
    XAML template string)."""
    c = brush.Color
    return "#{0:02X}{1:02X}{2:02X}".format(c.R, c.G, c.B)


def _theme_dim(w, key, fallback):
    """Live sizing-resource lookup (double / CornerRadius / Thickness);
    fallback must already be the correct type for the property it's used on."""
    try:
        v = w.TryFindResource(key)
        if v is not None:
            return v
    except Exception:
        pass
    return fallback


def _button(w, text, primary=True):
    b = Windows.Controls.Button()
    b.Content = text
    b.MinWidth = 84
    b.Height = _theme_dim(w, 'HeightButtonSmall', 24.0)
    b.Margin = Windows.Thickness(8, 0, 0, 0)
    b.Background = _theme_brush(w, 'BrushPrimaryGreen' if primary else 'BrushSecondaryBg', GREEN if primary else SECOND)
    b.Foreground = _theme_brush(w, 'BrushTextPrimary', TEXT)
    b.BorderThickness = Windows.Thickness(0)
    b.FontWeight = Windows.FontWeights.SemiBold
    b.FontSize = _theme_dim(w, 'FontSizeButtonSmall', 11.0)
    b.Cursor = Windows.Input.Cursors.Hand

    # Colours/sizing below are static {DynamicResource ...} XAML, resolved
    # once this template is attached to the button and rendered in the
    # window's tree - not computed hex/numeric strings substituted into the
    # XAML from Python. That substitution approach used to crash on some
    # machines: a .NET double's ToString() is culture-aware, so under a
    # comma-decimal locale a CornerRadius like 13.0 could come out as the
    # string "13,0" - which WPF's parser reads as two values, not one,
    # corrupting that attribute and throwing off the parser for the rest of
    # the (single-line) template string. DynamicResource has no such risk.
    hover_key = 'BrushPrimaryGreenHover' if primary else 'BrushSecondaryHover'
    import System.Windows.Markup as _Markup
    template_xaml = (
        '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
        'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" '
        'TargetType="Button">'
        '<Border x:Name="Bd" Background="{TemplateBinding Background}" '
        'CornerRadius="{DynamicResource CornerRadiusButtonSmall}" Padding="14,0">'
        '<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>'
        '</Border>'
        '<ControlTemplate.Triggers>'
        '<Trigger Property="IsMouseOver" Value="True">'
        '<Setter TargetName="Bd" Property="Background" Value="{DynamicResource ' + hover_key + '}"/>'
        '</Trigger>'
        '</ControlTemplate.Triggers>'
        '</ControlTemplate>'
    )
    b.Template = _Markup.XamlReader.Parse(template_xaml)
    return b


def _close_button(w):
    """Small round close (X) button, top-right of the card - every other
    Seed43 window has one of these, this shared dialog previously didn't."""
    b = Windows.Controls.Button()
    size = _theme_dim(w, 'WidthButtonClose', 26.0)
    b.Width = size
    b.Height = size
    b.Background = Windows.Media.Brushes.Transparent
    b.BorderThickness = Windows.Thickness(0)
    b.Cursor = Windows.Input.Cursors.Hand
    b.HorizontalAlignment = Windows.HorizontalAlignment.Right
    b.VerticalAlignment = Windows.VerticalAlignment.Top

    text_brush = _theme_brush(w, 'BrushTextPrimary', TEXT)
    icon_hex = get_color(_SCRIPT_DIR, 'text_primary', fallback=TEXT) if get_color else _brush_hex(text_brush)
    if _make_icon:
        try:
            b.Content = _make_icon('close', size=12, color=icon_hex)
        except Exception:
            pass
    if not b.Content:
        t = Windows.Controls.TextBlock()
        t.Text = u"\u2715"
        t.Foreground = text_brush
        t.FontSize = 11
        t.HorizontalAlignment = Windows.HorizontalAlignment.Center
        t.VerticalAlignment = Windows.VerticalAlignment.Center
        b.Content = t

    import System.Windows.Markup as _Markup
    template_xaml = (
        '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
        'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" '
        'TargetType="Button">'
        '<Border x:Name="Bd" Background="{TemplateBinding Background}" '
        'CornerRadius="{DynamicResource CornerRadiusButtonClose}">'
        '<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>'
        '</Border>'
        '<ControlTemplate.Triggers>'
        '<Trigger Property="IsMouseOver" Value="True">'
        '<Setter TargetName="Bd" Property="Background" Value="{DynamicResource BrushDanger}"/>'
        '</Trigger>'
        '</ControlTemplate.Triggers>'
        '</ControlTemplate>'
    )
    b.Template = _Markup.XamlReader.Parse(template_xaml)
    b.Click += lambda s, a: w.Close()
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
            apply_seed43_palette(w, _SCRIPT_DIR)
        except Exception:
            pass
    if apply_seed43_dimensions:
        try:
            apply_seed43_dimensions(w, _SCRIPT_DIR)
        except Exception:
            pass
    return w


def _card(w):
    """Rounded dark card with an accent bar + close button up top, drop
    shadow, drag-movable by clicking the card background."""
    outer = Windows.Controls.Border()
    outer.Background = _theme_brush(w, 'BrushCardBg', BG)
    outer.CornerRadius = _theme_dim(w, 'CornerRadiusCard', Windows.CornerRadius(10))
    outer.Margin = Windows.Thickness(12)
    outer.Padding = Windows.Thickness(24, 16, 24, 20)

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

    grid = Windows.Controls.Grid()
    r0 = Windows.Controls.RowDefinition(); r0.Height = Windows.GridLength.Auto
    r1 = Windows.Controls.RowDefinition(); r1.Height = Windows.GridLength.Auto
    grid.RowDefinitions.Add(r0)
    grid.RowDefinitions.Add(r1)

    # Header: accent bar filling the row, close button docked over its right end
    header = Windows.Controls.DockPanel()
    header.Margin = Windows.Thickness(0, 0, 0, 16)
    close_btn = _close_button(w)
    Windows.Controls.DockPanel.SetDock(close_btn, Windows.Controls.Dock.Right)
    header.Children.Add(close_btn)
    accent = Windows.Controls.Border()
    accent.Background = _theme_brush(w, 'BrushPrimaryGreen', GREEN)
    accent.Height = 3
    accent.CornerRadius = Windows.CornerRadius(2)
    accent.VerticalAlignment = Windows.VerticalAlignment.Center
    accent.Margin = Windows.Thickness(0, 0, 8, 0)
    header.Children.Add(accent)
    Windows.Controls.Grid.SetRow(header, 0)

    root = Windows.Controls.StackPanel()
    Windows.Controls.Grid.SetRow(root, 1)

    grid.Children.Add(header)
    grid.Children.Add(root)
    outer.Child = grid
    w.Content = outer

    def _key(s, a):
        if a.Key == Windows.Input.Key.Escape:
            w.Close()
    w.KeyDown += _key

    return root


def _textblock(w, text, error=False, title=False):
    t = Windows.Controls.TextBlock()
    t.Text = text
    t.Foreground = _theme_brush(w, 'BrushDanger', RED) if error else _theme_brush(w, 'BrushTextPrimary', TEXT)
    t.FontSize = 15 if title else 12
    t.FontWeight = Windows.FontWeights.Bold if title else Windows.FontWeights.Normal
    t.Opacity = 1.0 if title else 0.85
    t.TextWrapping = Windows.TextWrapping.Wrap
    t.Margin = Windows.Thickness(0, 0, 0, 8 if title else 20)
    return t


def _themed_textbox(w):
    """Rounded, dark-bg/light-text TextBox matching input_textbox - a plain
    WPF TextBox has no CornerRadius of its own, so this builds the same
    rounded-border template every other themed input in this extension
    uses, via a parsed ControlTemplate string like _button()/_close_button().

    Sized a bit more generously (34px, not the compact 28px HeightInput
    token) since a modal dialog has more breathing room than a tight
    main-window field row, and the compact sizing was reported as clipping
    text here."""
    tb = Windows.Controls.TextBox()
    tb.Height = 34.0
    tb.FontSize = _theme_dim(w, 'FontSizeInput', 12.0)
    tb.Padding = Windows.Thickness(10, 8, 10, 8)
    tb.Background = _theme_brush(w, 'BrushInputBg', INPUTBG)
    tb.Foreground = _theme_brush(w, 'BrushTextInput', TEXT)
    tb.BorderBrush = _theme_brush(w, 'BrushBorderDefault', GREEN)
    tb.BorderThickness = Windows.Thickness(1)
    tb.Margin = Windows.Thickness(0, 0, 0, 16)
    tb.VerticalContentAlignment = Windows.VerticalAlignment.Center

    import System.Windows.Markup as _Markup
    template_xaml = (
        '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
        'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" '
        'TargetType="TextBox">'
        '<Border x:Name="Bd" Background="{TemplateBinding Background}" '
        'BorderBrush="{TemplateBinding BorderBrush}" '
        'BorderThickness="{TemplateBinding BorderThickness}" '
        'CornerRadius="{DynamicResource CornerRadiusInput}">'
        '<ScrollViewer x:Name="PART_ContentHost" '
        'Focusable="False" '
        'HorizontalScrollBarVisibility="Hidden" '
        'VerticalScrollBarVisibility="Hidden"/>'
        '</Border>'
        '<ControlTemplate.Triggers>'
        '<Trigger Property="IsMouseOver" Value="True">'
        '<Setter TargetName="Bd" Property="BorderBrush" Value="{DynamicResource BrushBorderHover}"/>'
        '</Trigger>'
        '<Trigger Property="IsFocused" Value="True">'
        '<Setter TargetName="Bd" Property="BorderBrush" Value="{DynamicResource BrushBorderHover}"/>'
        '<Setter TargetName="Bd" Property="BorderThickness" Value="2"/>'
        '</Trigger>'
        '</ControlTemplate.Triggers>'
        '</ControlTemplate>'
    )
    tb.Template = _Markup.XamlReader.Parse(template_xaml)
    return tb


def choice(text, options, title='', detail_text=''):
    """Themed multi-button choice popup (more than the two confirm() gives
    you). `options` is a list of (key, label) tuples - the last one is
    rendered as the primary/green button, all others secondary/grey, left
    to right in the order given. Returns the chosen key, or None if closed
    via the X or Escape without choosing.

    detail_text, if given, shows a read-only preview panel below the
    buttons (collapsed entirely if empty) - e.g. a diff or log excerpt.
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
        panel.Background = _theme_brush(w, 'BrushHeaderBg', HEADER)
        panel.CornerRadius = _theme_dim(w, 'CornerRadiusInput', Windows.CornerRadius(6))
        panel.Padding = Windows.Thickness(12)
        panel.Margin = Windows.Thickness(0, 16, 0, 0)
        detail = Windows.Controls.TextBlock()
        detail.Text = detail_text
        detail.Foreground = _theme_brush(w, 'LocalBrushNoticeText', '#7EC8A0')
        detail.FontSize = 11
        detail.Opacity = 0.85
        detail.TextWrapping = Windows.TextWrapping.Wrap
        detail.FontFamily = Windows.Media.FontFamily('Consolas')
        panel.Child = detail
        root.Children.Add(panel)

    w.ShowDialog()
    return result['key']


def save_file_as(title, filename, ext, initial_folder=None):
    """Themed 'save as' dialog: an editable folder path (with a native
    folder-browse button) and an editable filename. Returns the chosen
    full path, or None if cancelled.
    """
    import os as _os
    from System.IO import Path as _Path

    w = _window(width=480)
    root = _card(w)
    result = {'path': None}
    root.Children.Add(_textblock(w, title, title=True))

    root.Children.Add(_textblock_label(w, 'Folder'))
    folder_row = Windows.Controls.Grid()
    c0 = Windows.Controls.ColumnDefinition(); c0.Width = Windows.GridLength(1, Windows.GridUnitType.Star)
    c1 = Windows.Controls.ColumnDefinition(); c1.Width = Windows.GridLength.Auto
    folder_row.ColumnDefinitions.Add(c0)
    folder_row.ColumnDefinitions.Add(c1)

    folder_tb = _themed_textbox(w)
    folder_tb.Margin = Windows.Thickness(0, 0, 6, 12)
    folder_tb.IsReadOnly = True
    Windows.Controls.Grid.SetColumn(folder_tb, 0)

    browse_btn = _button(w, 'Browse')
    browse_btn.Margin = Windows.Thickness(0, 0, 0, 12)
    Windows.Controls.Grid.SetColumn(browse_btn, 1)

    folder_row.Children.Add(folder_tb)
    folder_row.Children.Add(browse_btn)
    root.Children.Add(folder_row)

    root.Children.Add(_textblock_label(w, 'File name'))
    filename_tb = _themed_textbox(w)
    root.Children.Add(filename_tb)

    filename_tb.Text = _Path.GetFileNameWithoutExtension(filename)
    desktop = _os.path.expanduser("~\\Desktop")
    folder_tb.Text = (
        initial_folder if (initial_folder and _os.path.isdir(initial_folder))
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
        folder = folder_tb.Text.strip()
        name = filename_tb.Text.strip()
        if not name.lower().endswith('.' + ext):
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

    w.ShowDialog()
    return result['path']


def _textblock_label(w, text):
    """Small dim field-label text, used above folder/filename in save_file_as."""
    t = Windows.Controls.TextBlock()
    t.Text = text
    t.Foreground = _theme_brush(w, 'BrushTextPrimary', TEXT)
    t.FontSize = 11
    t.Opacity = 0.7
    t.Margin = Windows.Thickness(0, 0, 0, 4)
    return t


def message(text, title='', ok_label='OK'):
    """Themed OK-only info popup. Escape or the X closes it, same as OK."""
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
    tb = _themed_textbox(w)
    tb.Text = default or ''
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
