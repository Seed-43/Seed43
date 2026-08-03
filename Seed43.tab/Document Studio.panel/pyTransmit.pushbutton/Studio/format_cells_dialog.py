# -*- coding: utf-8 -*-
# format_cells_dialog.py
#
# "Format Cells..." dialog for pyTransmit Studio, opened via right-click on a
# cell. Does the same things the Home ribbon does (font, fill, alignment,
# borders) in one place, matching Excel's Format Cells dialog - just not
# split into tabs, since the ribbon already covers the same ground and this
# is meant as a single-stop equivalent, not a second UI to maintain in
# parallel. Takes the already-open StudioSettingsWindow (`win`) so it can
# reuse its selection helpers instead of duplicating them.

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System.Windows as _SW
import System.Windows.Controls as _SWC
import System.Windows.Media as _SWM
import System.Windows.Media.Effects as _SWME
import System.Windows.Input as _SWI

import studio_blocks


def _brush(hexcolor):
    h = (hexcolor or '#000000').lstrip('#')
    if len(h) == 3:
        h = h[0] * 2 + h[1] * 2 + h[2] * 2
    return _SWM.SolidColorBrush(_SWM.Color.FromRgb(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)))


def _section_label(text):
    tb = _SWC.TextBlock()
    tb.Text = text
    tb.FontSize = 10
    tb.FontWeight = _SW.FontWeights.SemiBold
    tb.Foreground = _brush('#8A96A8')
    tb.Margin = _SW.Thickness(0, 12, 0, 4)
    return tb


def open_format_cells_dialog(win):
    block = win._anchor_block() or {}

    w = _SW.Window()
    w.Title = 'Format Cells'
    w.Width = 380
    w.SizeToContent = _SW.SizeToContent.Height
    w.WindowStartupLocation = _SW.WindowStartupLocation.CenterScreen
    w.ResizeMode = _SW.ResizeMode.NoResize
    w.WindowStyle = getattr(_SW.WindowStyle, 'None')
    w.AllowsTransparency = True
    w.Background = _SWM.Brushes.Transparent
    w.ShowInTaskbar = False

    outer = _SWC.Border()
    outer.Background = _brush('#2B3340')
    outer.CornerRadius = _SW.CornerRadius(10)
    outer.Margin = _SW.Thickness(12)
    outer.Padding = _SW.Thickness(20)
    shadow = _SWME.DropShadowEffect()
    shadow.Color = _SWM.Colors.Black
    shadow.Opacity = 0.5
    shadow.ShadowDepth = 4
    shadow.BlurRadius = 16
    outer.Effect = shadow

    def _drag(s, a):
        if a.ButtonState == _SWI.MouseButtonState.Pressed:
            w.DragMove()
    outer.MouseLeftButtonDown += _drag

    root = _SWC.StackPanel()
    outer.Child = root
    w.Content = outer

    title = _SWC.TextBlock()
    title.Text = 'Format Cells'
    title.FontSize = 15
    title.FontWeight = _SW.FontWeights.Bold
    title.Foreground = _SWM.Brushes.White
    root.Children.Add(title)

    # -- Font ------------------------------------------------------------------
    root.Children.Add(_section_label('FONT'))
    font_row = _SWC.StackPanel()
    font_row.Orientation = _SWC.Orientation.Horizontal
    family_combo = _SWC.ComboBox()
    family_combo.Width = 150
    family_combo.Margin = _SW.Thickness(0, 0, 6, 0)
    for fam in ['Arial', 'Calibri', 'Segoe UI', 'Times New Roman', 'Verdana']:
        family_combo.Items.Add(fam)
    family_combo.SelectedItem = block.get('font', 'Arial')
    if family_combo.SelectedItem is None:
        family_combo.SelectedIndex = 0
    size_combo = _SWC.ComboBox()
    size_combo.Width = 70
    # Points, matching the Home ribbon and Excel (model stores mm).
    for sz in studio_blocks.FONT_SIZES_PT:
        size_combo.Items.Add(str(sz))
    cur_size = block.get('size_mm')
    cur_pt = str(int(round(studio_blocks.mm_to_pt(cur_size)))) if cur_size else '9'
    size_combo.SelectedItem = cur_pt if cur_pt in list(size_combo.Items) else '9'
    font_row.Children.Add(family_combo)
    font_row.Children.Add(size_combo)
    root.Children.Add(font_row)

    style_row = _SWC.StackPanel()
    style_row.Orientation = _SWC.Orientation.Horizontal
    style_row.Margin = _SW.Thickness(0, 6, 0, 0)
    bold_cb = _SWC.CheckBox(); bold_cb.Content = 'Bold'; bold_cb.Foreground = _SWM.Brushes.White
    bold_cb.IsChecked = bool(block.get('bold')); bold_cb.Margin = _SW.Thickness(0, 0, 12, 0)
    italic_cb = _SWC.CheckBox(); italic_cb.Content = 'Italic'; italic_cb.Foreground = _SWM.Brushes.White
    italic_cb.IsChecked = bool(block.get('italic')); italic_cb.Margin = _SW.Thickness(0, 0, 12, 0)
    underline_cb = _SWC.CheckBox(); underline_cb.Content = 'Underline'; underline_cb.Foreground = _SWM.Brushes.White
    underline_cb.IsChecked = bool(block.get('underline'))
    style_row.Children.Add(bold_cb); style_row.Children.Add(italic_cb); style_row.Children.Add(underline_cb)
    root.Children.Add(style_row)

    color_state = {'font': block.get('color', '#000000'), 'fill': block.get('bg_color')}
    color_row = _SWC.StackPanel()
    color_row.Orientation = _SWC.Orientation.Horizontal
    color_row.Margin = _SW.Thickness(0, 8, 0, 0)
    font_color_btn = _SWC.Button(); font_color_btn.Content = 'Font Color…'; font_color_btn.Padding = _SW.Thickness(8, 3, 8, 3)
    font_color_btn.Margin = _SW.Thickness(0, 0, 6, 0)

    def _pick_font_color(s, a):
        picked = win._pick_color(color_state['font'])
        if picked:
            color_state['font'] = picked
    font_color_btn.Click += _pick_font_color
    fill_color_btn = _SWC.Button(); fill_color_btn.Content = 'Fill Color…'; fill_color_btn.Padding = _SW.Thickness(8, 3, 8, 3)

    def _pick_fill_color(s, a):
        picked = win._pick_color(color_state['fill'] or '#FFFFFF')
        if picked:
            color_state['fill'] = picked
    fill_color_btn.Click += _pick_fill_color
    color_row.Children.Add(font_color_btn)
    color_row.Children.Add(fill_color_btn)
    root.Children.Add(color_row)

    # -- Alignment ---------------------------------------------------------
    root.Children.Add(_section_label('ALIGNMENT'))
    align_state = {'just': block.get('just', 'left'), 'v_just': block.get('v_just', 'middle')}
    h_row = _SWC.StackPanel()
    h_row.Orientation = _SWC.Orientation.Horizontal
    h_buttons = []

    def _make_align_btn(text, key, state_key, row):
        # A plain Button with a manually-toggled highlight, not a real
        # ToggleButton - ToggleButton's actual CLR namespace is
        # System.Windows.Controls.Primitives, not System.Windows.Controls,
        # so this sidesteps needing that import just for a simple picker.
        btn = _SWC.Button()
        btn.Content = text
        btn.Width = 70
        btn.Margin = _SW.Thickness(0, 0, 6, 0)
        btn.Foreground = _SWM.Brushes.White
        btn.BorderThickness = _SW.Thickness(1)

        def _refresh():
            selected = (align_state[state_key] == key)
            btn.Background = _brush('#208A3C') if selected else _brush('#232933')
            btn.BorderBrush = _brush('#208A3C') if selected else _brush('#404E62')

        def _click(s, a):
            align_state[state_key] = key
            for b in row:
                b['refresh']()
        btn.Click += _click
        row.append({'key': key, 'btn': btn, 'refresh': _refresh})
        _refresh()
        return btn

    for key, label in (('left', 'Left'), ('center', 'Center'), ('right', 'Right')):
        h_row.Children.Add(_make_align_btn(label, key, 'just', h_buttons))
    root.Children.Add(h_row)

    v_row = _SWC.StackPanel()
    v_row.Orientation = _SWC.Orientation.Horizontal
    v_row.Margin = _SW.Thickness(0, 6, 0, 0)
    v_buttons = []
    for key, label in (('top', 'Top'), ('middle', 'Middle'), ('bottom', 'Bottom')):
        v_row.Children.Add(_make_align_btn(label, key, 'v_just', v_buttons))
    root.Children.Add(v_row)

    # -- Borders (reuses the same icon builder as the ribbon dropdown) --------
    root.Children.Add(_section_label('BORDERS'))
    border_state = {'preset': None}
    # WrapPanel, not UniformGrid - UniformGrid's real CLR namespace is
    # System.Windows.Controls.Primitives, which plain "import
    # System.Windows.Controls" doesn't reliably expose. WrapPanel is already
    # proven working (the Blocks panel uses it), so reuse that instead of a
    # 3rd different container type just for a 3-per-row icon layout.
    border_grid = _SWC.WrapPanel()
    border_grid.Width = 3 * 40
    border_buttons = []
    for key, label in win.BORDER_PRESETS:
        btn = _SWC.Button()
        btn.Tag = key
        btn.ToolTip = label
        btn.Margin = _SW.Thickness(3)
        btn.Padding = _SW.Thickness(3)
        btn.Background = _SWM.Brushes.Transparent
        btn.BorderBrush = _SWM.Brushes.Transparent
        btn.BorderThickness = _SW.Thickness(1)
        btn.Content = win._build_border_icon(key)

        def _click(s, a, k=key):
            border_state['preset'] = k
            for b in border_buttons:
                b.BorderBrush = _brush('#20A344') if b.Tag == k else _SWM.Brushes.Transparent
        btn.Click += _click
        border_buttons.append(btn)
        border_grid.Children.Add(btn)
    root.Children.Add(border_grid)

    # -- OK / Cancel -------------------------------------------------------
    btn_row = _SWC.StackPanel()
    btn_row.Orientation = _SWC.Orientation.Horizontal
    btn_row.HorizontalAlignment = _SW.HorizontalAlignment.Right
    btn_row.Margin = _SW.Thickness(0, 16, 0, 0)

    def _styled_btn(text, primary):
        b = _SWC.Button()
        b.Content = text
        b.MinWidth = 84
        b.Height = 30
        b.Margin = _SW.Thickness(8, 0, 0, 0)
        b.Background = _brush('#208A3C') if primary else _brush('#232933')
        b.Foreground = _SWM.Brushes.White
        b.BorderThickness = _SW.Thickness(0)
        return b

    cancel_btn = _styled_btn('Cancel', False)
    cancel_btn.Click += lambda s, a: w.Close()
    ok_btn = _styled_btn('OK', True)

    def _on_ok(s, a):
        fam = family_combo.SelectedItem or 'Arial'
        try:
            size_mm = studio_blocks.pt_to_mm(float(size_combo.SelectedItem or 9))
        except Exception:
            size_mm = studio_blocks.pt_to_mm(9)

        def _apply(b):
            b['font'] = fam
            b['size_mm'] = size_mm
            b['bold'] = bool(bold_cb.IsChecked)
            b['italic'] = bool(italic_cb.IsChecked)
            b['underline'] = bool(underline_cb.IsChecked)
            b['color'] = color_state['font']
            if color_state['fill']:
                b['bg_color'] = color_state['fill']
            b['just'] = align_state['just']
            b['v_just'] = align_state['v_just']
        win._apply_to_selection(_apply)
        if border_state['preset']:
            win._apply_border_preset(border_state['preset'])
        w.Close()
    ok_btn.Click += _on_ok

    btn_row.Children.Add(cancel_btn)
    btn_row.Children.Add(ok_btn)
    root.Children.Add(btn_row)

    w.ShowDialog()
