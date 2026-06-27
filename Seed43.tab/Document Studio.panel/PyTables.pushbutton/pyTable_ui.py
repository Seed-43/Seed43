# -*- coding: utf-8 -*-
"""
pyTable WPF window.

Architecture: one Card (Border/CardStyle) per loaded file.
Each card has a filename heading and a StackPanel of rows.
Adding a row appends to the active file's card.
"""
import os
import wpf
from pyrevit import forms, script
from System.Windows import (
    ResourceDictionary, Visibility, Thickness,
    VerticalAlignment, HorizontalAlignment,
    FontWeights, CornerRadius, TextTrimming
)
from System import Uri, DateTime
from System.Windows.Controls import (
    StackPanel, Border, CheckBox, TextBlock, TextBox,
    ComboBox, Button, Orientation, ScrollViewer
)
from System.Windows.Shapes import Ellipse
from System.Windows.Media import SolidColorBrush, Color

logger = script.get_logger()

VIEW_TYPES      = ['Schedule View', 'Legend View', 'Drafting View']
WORD_VIEW_TYPES = ['Legend View', 'Drafting View']
SHEET_SIZES     = [
    'A4 Landscape', 'A4 Portrait',
    'A3 Landscape', 'A3 Portrait',
    'A2 Landscape', 'A2 Portrait',
    'A1 Landscape', 'A1 Portrait',
    'A0 Landscape', 'A0 Portrait',
]

SRC_COLOURS    = {'xl': '#217346', 'word': '#2B579A'}
STATUS_COLOURS = {
    'pending': '#6B7280', 'success': '#16A34A',
    'error':   '#DC2626', 'skipped': '#CA8A04',
    'sync':    '#3B82F6',
}

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def _find_seed43_styles():
    folder = _SCRIPT_DIR
    for _ in range(6):
        for sub in ('', 'UI'):
            candidate = os.path.join(folder, sub, 'Seed43Styles.xaml')
            if os.path.isfile(candidate):
                return candidate
        folder = os.path.dirname(folder)
    return None


def hb(h):
    h = h.lstrip('#')
    return SolidColorBrush(Color.FromRgb(
        int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)))


class Row(object):
    _counter = [0]

    def __init__(self, file_path, source_type,
                 sheets=None, sheet_range_map=None):
        Row._counter[0] += 1
        self._id              = Row._counter[0]
        self.FilePath         = file_path
        self.SourceType       = source_type
        self.Enabled          = True
        self.ViewName         = ''
        self.Sheet            = ''
        self.NamedRange       = ''
        self.ViewType         = VIEW_TYPES[0]
        self.Status           = 'pending'
        self.LastModified     = self._mtime(file_path)
        self._sheets          = sheets or []
        self._sheet_range_map = sheet_range_map or {}
        self._dot             = None   # WPF Ellipse, set by _make_row_ui
        self._refresh_btn     = None   # per-row refresh button, shown on sync
        self._vn_textbox      = None   # TextBox ref for ViewName — update when auto-filling
        self._error_label     = None   # inline error pill Border
        self._error_text      = None   # TextBlock inside error pill
        self._applied_mtime   = None   # mtime when last applied
        self._applied_hash    = None   # MD5 of range content at last apply
        # Word-specific
        self.ColNo            = 1      # column assignment (1-based)
        self._col_textbox     = None   # TextBox for col number
        self._drag_origin     = None   # mouse Y when drag started
        self._drag_panel_ref  = None   # card_panel ref for reorder
        if self._sheets:
            self.Sheet = self._sheets[0]

    def _mtime(self, path):
        try:
            dt = DateTime.FromFileTime(
                int(os.path.getmtime(path) * 10000000) + 116444736000000000)
            return dt.ToString('dd/MM/yyyy HH:mm')
        except Exception:
            return ''

    def ranges_for(self, sheet=None):
        s = sheet if sheet is not None else self.Sheet
        return self._sheet_range_map.get(s, [])

    @property
    def SourceLabel(self):
        return {'xl': 'XL', 'word': 'W'}.get(self.SourceType, '?')

    @property
    def SourceColour(self):
        return hb(SRC_COLOURS.get(self.SourceType, '#555'))


class PyTableWindow(forms.WPFWindow):

    def __init__(self):
        styles_path = _find_seed43_styles()
        if styles_path:
            try:
                rd = ResourceDictionary()
                rd.Source = Uri(styles_path)
                self.Resources = rd
            except Exception as ex:
                logger.warning('Seed43Styles load failed: {}'.format(ex))
        forms.WPFWindow.__init__(self, 'pyTable_ui.xaml')
        self.Closing += self._on_window_closing
        # _file_data: {path: {sheets, sheet_range_map, source_type, rows, card_panel}}
        self._file_data   = {}
        self._active_file = None
        self._status_card  = None
        self._progress_bar = None
        self._lbl_pending  = None
        self._lbl_success  = None
        self._lbl_error    = None
        self._lbl_skipped  = None
        self._sync_panel   = None
        self._sync_label   = None
        self._update_file_combo()
        self._build_status_card()
        self._refresh_status_card()
        self._load_persisted_state()
        self._refresh_drop_zone()

    # ------------------------------------------------------------------
    # File combo
    # ------------------------------------------------------------------

    def _update_file_combo(self):
        combo = self.FileCombo
        combo.Items.Clear()
        for path in self._file_data:
            combo.Items.Add(os.path.basename(path))
        if self._active_file:
            combo.SelectedItem = os.path.basename(self._active_file)
        elif combo.Items.Count > 0:
            combo.SelectedIndex = 0
            self._active_file = list(self._file_data.keys())[0]

    def OnFileComboChanged(self, sender, e):
        idx = sender.SelectedIndex
        if idx < 0:
            return
        paths = list(self._file_data.keys())
        if idx < len(paths):
            self._active_file = paths[idx]

    # ------------------------------------------------------------------
    # Drop zone visibility
    # ------------------------------------------------------------------

    def _refresh_drop_zone(self):
        if self._file_data:
            self.DropZonePanel.Visibility = Visibility.Collapsed
        else:
            self.DropZonePanel.Visibility = Visibility.Visible
        self._update_footer()

    def _update_footer(self):
        total = sum(len(fd['rows']) for fd in self._file_data.values())
        xl    = sum(1 for fd in self._file_data.values()
                    if fd.get('source_type') == 'xl')
        word  = len(self._file_data) - xl
        self.FooterCounts.Text = (
            'Total {} | Excel {} | Word {}'.format(total, xl, word)
            if self._file_data else 'No files loaded')

    def _set_status(self, text):
        pass  # status shown via dots and progress bar only

    # ------------------------------------------------------------------
    # Card builder
    # ------------------------------------------------------------------

    def _make_card(self, path):
        """Create a card Border for a file and add it to CardsPanel."""
        fd = self._file_data[path]

        outer = Border()
        try:
            outer.Style = self.FindResource('CardStyle')
        except Exception:
            outer.Background     = hb('#2B3340')
            outer.CornerRadius   = CornerRadius(8)
            outer.Padding        = Thickness(16)
            outer.Margin         = Thickness(0, 0, 0, 12)

        inner = StackPanel()
        inner.Orientation = Orientation.Vertical

        # Card header row: delete button + file path heading
        header_row = StackPanel()
        header_row.Orientation = Orientation.Horizontal
        header_row.Margin      = Thickness(0, 0, 0, 8)

        del_card_btn = Button()
        del_card_btn.Content         = '✕'
        del_card_btn.FontSize        = 9
        del_card_btn.Foreground      = hb('#F4FAFF')
        del_card_btn.Opacity         = 0.3
        del_card_btn.Background      = hb('#2B3340')
        del_card_btn.BorderThickness = Thickness(0)
        del_card_btn.Width           = 20
        del_card_btn.Height          = 20
        del_card_btn.VerticalAlignment   = VerticalAlignment.Center
        del_card_btn.Margin          = Thickness(0, 0, 8, 0)
        del_card_btn.Tag             = path
        del_card_btn.Click          += self._del_card_click
        header_row.Children.Add(del_card_btn)

        # Per-card Add Row button
        add_row_btn = self._green_btn(u'+ Row')
        add_row_btn.VerticalAlignment = VerticalAlignment.Center
        add_row_btn.Margin = Thickness(0, 0, 8, 0)
        add_row_btn.Tag    = path
        add_row_btn.Click += self._add_row_for_card
        header_row.Children.Add(add_row_btn)

        heading = TextBlock()
        heading.Text         = path
        heading.TextTrimming = TextTrimming.CharacterEllipsis
        heading.ToolTip      = path
        heading.VerticalAlignment = VerticalAlignment.Center
        try:
            heading.Style = self.FindResource('SectionLabelStyle')
            heading.Margin = Thickness(0)
        except Exception:
            heading.Foreground = hb('#208A3C')
            heading.FontWeight  = FontWeights.SemiBold
            heading.FontSize    = 13
        header_row.Children.Add(heading)

        # Row container
        row_panel = StackPanel()
        row_panel.Orientation = Orientation.Vertical

        inner.Children.Add(header_row)

        # ── Word-specific card controls (sheet size + col count) ──────
        is_word = fd.get('source_type') == 'word'
        if is_word:
            word_ctrl_row = StackPanel()
            word_ctrl_row.Orientation = Orientation.Horizontal
            word_ctrl_row.Margin      = Thickness(0, 0, 0, 8)

            def _ctrl_label(text):
                tb = TextBlock()
                tb.Text              = text
                tb.FontSize          = 10
                tb.Foreground        = hb('#F4FAFF')
                tb.Opacity           = 0.55
                tb.VerticalAlignment = VerticalAlignment.Center
                tb.Margin            = Thickness(0, 0, 6, 0)
                return tb

            # View name — defaults to filename stem, user can edit
            default_vname = fd.get(
                'view_name',
                os.path.splitext(os.path.basename(path))[0])
            fd['view_name'] = fd.get('view_name', default_vname)
            word_ctrl_row.Children.Add(_ctrl_label('View Name:'))

            vn_box = TextBox()
            vn_box.Text          = fd['view_name']
            vn_box.Width         = 180
            vn_box.Height        = 26
            vn_box.FontSize      = 11
            try:
                vn_box.Style = self.FindResource('GridCellStyle')
            except Exception:
                vn_box.Background      = hb('#F4FAFF')
                vn_box.Foreground      = hb('#2B3340')
                vn_box.BorderBrush     = hb('#208A3C')
                vn_box.BorderThickness = Thickness(1)
            vn_box.VerticalContentAlignment = VerticalAlignment.Center
            vn_box.Margin        = Thickness(0, 0, 12, 0)
            vn_box.VerticalAlignment = VerticalAlignment.Center
            vn_box.Tag           = path
            vn_box.LostFocus    += self._word_view_name_changed
            word_ctrl_row.Children.Add(vn_box)

            word_ctrl_row.Children.Add(_ctrl_label('Sheet size:'))

            sz_combo = ComboBox()
            self._combo_style(sz_combo, 115)
            sz_combo.Margin = Thickness(0, 0, 12, 0)
            for sz in SHEET_SIZES:
                sz_combo.Items.Add(sz)
            sz_combo.SelectedItem = fd.get('sheet_size', 'A3')
            sz_combo.Tag          = path
            sz_combo.SelectionChanged += self._word_sheet_size_changed
            word_ctrl_row.Children.Add(sz_combo)

            word_ctrl_row.Children.Add(_ctrl_label('Columns:'))

            cc_box = TextBox()
            cc_box.Text             = str(fd.get('col_count', 2))
            cc_box.Width            = 36
            cc_box.Height           = 26
            cc_box.FontSize         = 11
            try:
                cc_box.Style = self.FindResource('GridCellStyle')
            except Exception:
                cc_box.Background  = hb('#F4FAFF')
                cc_box.Foreground  = hb('#2B3340')
                cc_box.BorderBrush = hb('#208A3C')
                cc_box.BorderThickness = Thickness(1)
            cc_box.VerticalContentAlignment = VerticalAlignment.Center
            cc_box.Margin           = Thickness(0, 0, 0, 0)
            cc_box.VerticalAlignment = VerticalAlignment.Center
            cc_box.Tag              = path
            cc_box.LostFocus       += self._word_col_count_changed
            word_ctrl_row.Children.Add(cc_box)

            inner.Children.Add(word_ctrl_row)

        # Column headers inside card
        col_hdr = StackPanel()
        col_hdr.Orientation = Orientation.Horizontal
        col_hdr.Margin      = Thickness(0, 0, 0, 4)

        def _ch(text, width, pad_left=4):
            tb = TextBlock()
            tb.Text             = text
            tb.Width            = width
            tb.FontSize         = 10
            tb.Foreground       = hb('#F4FAFF')
            tb.Opacity          = 0.45
            tb.VerticalAlignment = VerticalAlignment.Center
            tb.Padding          = Thickness(pad_left, 0, 0, 0)
            return tb

        if is_word:
            col_hdr.Children.Add(_ch('',        24, 0))   # drag handle
            col_hdr.Children.Add(_ch('',        14, 0))   # status dot
            col_hdr.Children.Add(_ch('',        28, 0))   # checkbox
            col_hdr.Children.Add(_ch('',        28, 0))   # badge
            col_hdr.Children.Add(_ch('Section', 200))
            col_hdr.Children.Add(_ch('Col',      44))
            col_hdr.Children.Add(_ch('View Type', 130))
        else:
            col_hdr.Children.Add(_ch('',           14, 0))
            col_hdr.Children.Add(_ch('',           28, 0))
            col_hdr.Children.Add(_ch('',           28, 0))
            col_hdr.Children.Add(_ch('View Name', 120))
            col_hdr.Children.Add(_ch('Sheet',     120))
            col_hdr.Children.Add(_ch('Range',     130))
            col_hdr.Children.Add(_ch('Modified',   96))
            col_hdr.Children.Add(_ch('View Type', 130))

        inner.Children.Add(col_hdr)
        inner.Children.Add(row_panel)
        outer.Child = inner

        fd['card_panel'] = row_panel
        fd['card_border'] = outer

        self.CardsPanel.Children.Add(outer)

    # ------------------------------------------------------------------
    # Row builder
    # ------------------------------------------------------------------

    def _short_error(self, message):
        """Map a full error message to a short inline label (max 5 words)."""
        m = message.lower()
        if 'no data'      in m or 'empty'     in m: return 'Empty range'
        if 'header'       in m:                      return 'No header row'
        if 'view type'    in m:                      return 'Invalid view type'
        if 'already exist' in m or 'name'     in m: return 'View name exists'
        if 'sheet'        in m:                      return 'Sheet not found'
        if 'ref'          in m:                      return 'Broken reference'
        if 'transaction'  in m or 'revit'     in m: return 'Revit error'
        return 'Failed'

    def _combo_style(self, combo, width):
        try:
            combo.Style = self.FindResource('ModernComboBoxStyle')
        except Exception:
            combo.Background      = hb('#F4FAFF')
            combo.Foreground      = hb('#2B3340')
            combo.BorderBrush     = hb('#208A3C')
            combo.BorderThickness = Thickness(1)
            combo.FontSize        = 11
        # Override style defaults: Height=28 and Margin="0,0,0,10"
        combo.Width             = width
        combo.Height            = 26
        combo.Margin            = Thickness(0, 0, 4, 0)
        combo.VerticalAlignment = VerticalAlignment.Center

    # ------------------------------------------------------------------
    # Word row builder
    # ------------------------------------------------------------------

    def _make_word_row_ui(self, row):
        """Build the row StackPanel for a Word section row."""
        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal
        sp.Height      = 32
        sp.Margin      = Thickness(0, 0, 0, 4)
        sp.Tag         = row
        # Store panel ref on row for drag reorder
        row._drag_panel_ref = None  # set by caller after insertion

        # Drag handle :::
        drag_lbl = TextBlock()
        drag_lbl.Text             = u'∷'   # ⠿ grid dots
        drag_lbl.FontSize         = 14
        drag_lbl.Foreground       = hb('#F4FAFF')
        drag_lbl.Opacity          = 0.35
        drag_lbl.Width            = 20
        drag_lbl.VerticalAlignment = VerticalAlignment.Center
        drag_lbl.Margin           = Thickness(0, 0, 4, 0)
        drag_lbl.Tag              = row
        drag_lbl.Cursor           = __import__(
            'System.Windows.Input', fromlist=['Cursors']).Cursors.SizeNS
        drag_lbl.PreviewMouseLeftButtonDown += self._drag_start
        sp.Children.Add(drag_lbl)

        # Status dot
        dot = Ellipse()
        dot.Width             = 8
        dot.Height            = 8
        dot.Fill              = hb(STATUS_COLOURS.get(row.Status, '#3A4A3A'))
        dot.VerticalAlignment = VerticalAlignment.Center
        dot.Margin            = Thickness(0, 0, 6, 0)
        dot.Tag               = row
        row._dot              = dot
        sp.Children.Add(dot)

        # Checkbox
        cb = CheckBox()
        cb.IsChecked         = row.Enabled
        cb.VerticalAlignment = VerticalAlignment.Center
        cb.Margin            = Thickness(0, 0, 6, 0)
        cb.Tag               = row
        cb.Click            += self._cb_click
        sp.Children.Add(cb)

        # W badge
        badge = Border()
        badge.Width             = 22
        badge.Height            = 22
        badge.CornerRadius      = CornerRadius(3)
        badge.Background        = row.SourceColour
        badge.Margin            = Thickness(0, 0, 6, 0)
        badge.VerticalAlignment = VerticalAlignment.Center
        blbl = TextBlock()
        blbl.Text               = row.SourceLabel
        blbl.FontSize           = 8
        blbl.FontWeight         = FontWeights.Bold
        blbl.Foreground         = hb('#FFFFFF')
        blbl.HorizontalAlignment = HorizontalAlignment.Center
        blbl.VerticalAlignment  = VerticalAlignment.Center
        badge.Child = blbl
        sp.Children.Add(badge)

        # Section picker combo (headings from the docx)
        sc = ComboBox()
        self._combo_style(sc, 200)
        for h in row.ranges_for():
            sc.Items.Add(h)
        if row.NamedRange and row.NamedRange in [
                sc.Items.GetItemAt(i) for i in range(sc.Items.Count)]:
            sc.SelectedItem = row.NamedRange
        elif sc.Items.Count > 0:
            sc.SelectedIndex = 0
            row.NamedRange   = sc.Items.GetItemAt(0)
        sc.Tag               = row
        sc.SelectionChanged += self._word_section_changed
        sp.Children.Add(sc)

        # Col number textbox
        col_tb = TextBox()
        col_tb.Text          = str(row.ColNo)
        col_tb.Width         = 36
        col_tb.Height        = 26
        col_tb.FontSize      = 11
        try:
            col_tb.Style = self.FindResource('GridCellStyle')
        except Exception:
            col_tb.Background  = hb('#F4FAFF')
            col_tb.Foreground  = hb('#2B3340')
            col_tb.BorderBrush = hb('#208A3C')
            col_tb.BorderThickness = Thickness(1)
        col_tb.VerticalContentAlignment = VerticalAlignment.Center
        col_tb.Margin        = Thickness(0, 0, 4, 0)
        col_tb.VerticalAlignment = VerticalAlignment.Center
        col_tb.Tag           = row
        col_tb.LostFocus    += self._word_col_no_changed
        row._col_textbox     = col_tb
        sp.Children.Add(col_tb)

        # View type combo — Legend/Drafting only
        vtc = ComboBox()
        self._combo_style(vtc, 130)
        for vt in WORD_VIEW_TYPES:
            vtc.Items.Add(vt)
        if row.ViewType in WORD_VIEW_TYPES:
            vtc.SelectedItem = row.ViewType
        else:
            vtc.SelectedIndex = 0
            row.ViewType      = WORD_VIEW_TYPES[0]
        vtc.Tag              = row
        vtc.SelectionChanged += self._vt_changed
        sp.Children.Add(vtc)

        rb = self._green_btn(u'\u21bb', height=24,
                             padding=(0, 0, 0, 0),
                             font_size=13, width=24)
        rb.HorizontalContentAlignment = HorizontalAlignment.Center
        rb.VerticalContentAlignment   = VerticalAlignment.Center
        rb.VerticalAlignment = VerticalAlignment.Center
        rb.Margin  = Thickness(4, 0, 0, 0)
        rb.Tag     = row
        rb.Visibility = (Visibility.Visible
                         if row.Status == 'sync'
                         else Visibility.Collapsed)
        rb.Click           += self._refresh_row_click
        row._refresh_btn    = rb
        sp.Children.Add(rb)

        # Delete button
        db = Button()
        db.Content          = u'✕'
        db.FontSize         = 10
        db.Foreground       = hb('#C53030')
        db.Background       = hb('#2B3340')
        db.BorderThickness  = Thickness(0)
        db.Width            = 24
        db.Height           = 24
        db.Cursor           = __import__(
            'System.Windows.Input',
            fromlist=['Cursors']).Cursors.Hand
        db.VerticalAlignment = VerticalAlignment.Center
        db.Margin           = Thickness(4, 0, 0, 0)
        db.Tag              = row
        db.Click           += self._del_click
        sp.Children.Add(db)

        # Error pill
        err_pill = Border()
        err_pill.Background      = hb('#3B1515')
        err_pill.BorderBrush     = hb('#DC2626')
        err_pill.BorderThickness = Thickness(1)
        err_pill.CornerRadius    = CornerRadius(4)
        err_pill.Padding         = Thickness(8, 2, 8, 2)
        err_pill.Margin          = Thickness(6, 0, 0, 0)
        err_pill.VerticalAlignment = VerticalAlignment.Center
        err_pill.Visibility      = Visibility.Collapsed
        err_txt = TextBlock()
        err_txt.FontSize         = 10
        err_txt.Foreground       = hb('#F87171')
        err_txt.VerticalAlignment = VerticalAlignment.Center
        err_txt.Text             = ''
        err_pill.Child           = err_txt
        row._error_label         = err_pill
        row._error_text          = err_txt
        sp.Children.Add(err_pill)

        return sp

    def _make_row_ui(self, row):
        if row.SourceType == 'word':
            return self._make_word_row_ui(row)

        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal
        sp.Height      = 32
        sp.Margin      = Thickness(0, 0, 0, 4)

        # Status dot — store ref on row so Apply can update it
        dot = Ellipse()
        dot.Width             = 8
        dot.Height            = 8
        dot.Fill              = hb(STATUS_COLOURS.get(row.Status, '#3A4A3A'))
        dot.VerticalAlignment = VerticalAlignment.Center
        dot.Margin            = Thickness(0, 0, 6, 0)
        dot.Tag               = row
        row._dot              = dot
        sp.Children.Add(dot)

        # Checkbox
        cb = CheckBox()
        cb.IsChecked         = row.Enabled
        cb.VerticalAlignment = VerticalAlignment.Center
        cb.Margin            = Thickness(0, 0, 6, 0)
        cb.Tag               = row
        cb.Click             += self._cb_click
        sp.Children.Add(cb)

        # Source badge
        badge = Border()
        badge.Width             = 22
        badge.Height            = 22
        badge.CornerRadius      = CornerRadius(3)
        badge.Background        = row.SourceColour
        badge.Margin            = Thickness(0, 0, 6, 0)
        badge.VerticalAlignment = VerticalAlignment.Center
        lbl = TextBlock()
        lbl.Text                = row.SourceLabel
        lbl.FontSize            = 8
        lbl.FontWeight          = FontWeights.Bold
        lbl.Foreground          = hb('#FFFFFF')
        lbl.HorizontalAlignment = HorizontalAlignment.Center
        lbl.VerticalAlignment   = VerticalAlignment.Center
        badge.Child = lbl
        sp.Children.Add(badge)

        # View Name
        vn = TextBox()
        vn.Text  = row.ViewName
        vn.Width = 120
        try:
            vn.Style = self.FindResource('GridCellStyle')
        except Exception:
            vn.Background             = hb('#F4FAFF')
            vn.Foreground             = hb('#2B3340')
            vn.BorderBrush            = hb('#208A3C')
            vn.BorderThickness        = Thickness(1)
            vn.FontSize               = 11
            vn.Height                 = 26
            vn.Padding                = Thickness(6, 3, 6, 3)
            vn.VerticalContentAlignment = VerticalAlignment.Center
        vn.Margin    = Thickness(0, 0, 4, 0)
        row._vn_textbox = vn
        vn.Tag       = row
        vn.LostFocus += self._vn_lost
        sp.Children.Add(vn)

        # Sheet combo
        sc = ComboBox()
        self._combo_style(sc, 120)
        sc.Margin = Thickness(0, 0, 4, 0)
        for s in row._sheets:
            sc.Items.Add(s)
        if row.Sheet:
            sc.SelectedItem = row.Sheet
        elif sc.Items.Count > 0:
            sc.SelectedIndex = 0
            row.Sheet = sc.Items[0]
        sp.Children.Add(sc)

        # Named Range combo
        rc = ComboBox()
        self._combo_style(rc, 130)
        rc.Margin = Thickness(0, 0, 4, 0)
        for r in row.ranges_for():
            rc.Items.Add(r)
        if row.NamedRange:
            rc.SelectedItem = row.NamedRange
        elif rc.Items.Count > 0:
            rc.SelectedIndex = 0
            row.NamedRange = rc.Items[0]
        # Auto-fill ViewName on first row creation if still blank
        if not row.ViewName and row.NamedRange:
            row.ViewName = row.NamedRange
            if row._vn_textbox is not None:
                row._vn_textbox.Text = row.ViewName
        rc.Tag              = row
        rc.SelectionChanged += self._rc_changed
        sp.Children.Add(rc)

        sc.Tag              = (row, rc)
        sc.SelectionChanged += self._sc_changed

        # Last modified
        lm = TextBlock()
        lm.Text             = row.LastModified
        lm.Width            = 96
        lm.FontSize         = 10
        lm.Foreground       = hb('#F4FAFF')
        lm.Opacity          = 0.55
        lm.VerticalAlignment = VerticalAlignment.Center
        lm.Padding          = Thickness(4, 0, 4, 0)
        sp.Children.Add(lm)

        # View Type combo
        vtc = ComboBox()
        self._combo_style(vtc, 130)
        vtc.Margin = Thickness(0, 0, 4, 0)
        for vt in VIEW_TYPES:
            vtc.Items.Add(vt)
        vtc.SelectedItem    = row.ViewType
        vtc.Tag             = row
        vtc.SelectionChanged += self._vt_changed
        sp.Children.Add(vtc)

        rb = self._green_btn(u'\u21bb', height=24,
                             padding=(0, 0, 0, 0),
                             font_size=13, width=24)
        rb.HorizontalContentAlignment = HorizontalAlignment.Center
        rb.VerticalContentAlignment   = VerticalAlignment.Center
        rb.VerticalAlignment = VerticalAlignment.Center
        rb.Margin  = Thickness(4, 0, 0, 0)
        rb.Tag     = row
        rb.Visibility = (Visibility.Visible
                         if row.Status == 'sync'
                         else Visibility.Collapsed)
        rb.Click           += self._refresh_row_click
        row._refresh_btn    = rb
        sp.Children.Add(rb)

        # Delete button
        db = Button()
        db.Content          = u'\u2715'
        db.FontSize         = 10
        db.Foreground       = hb('#C53030')
        db.Background       = hb('#2B3340')
        db.BorderThickness  = Thickness(0)
        db.Width            = 24
        db.Height           = 24
        db.Cursor           = __import__('System.Windows.Input',
                                  fromlist=['Cursors']).Cursors.Hand
        db.VerticalAlignment = VerticalAlignment.Center
        db.Margin           = Thickness(4, 0, 0, 0)
        db.Tag              = row
        db.Click            += self._del_click
        sp.Children.Add(db)

        # Error pill — appears after ✕ as a small card
        err_pill = Border()
        err_pill.Background     = hb('#3B1515')
        err_pill.BorderBrush    = hb('#DC2626')
        err_pill.BorderThickness = Thickness(1)
        err_pill.CornerRadius   = CornerRadius(4)
        err_pill.Padding        = Thickness(8, 2, 8, 2)
        err_pill.Margin         = Thickness(6, 0, 0, 0)
        err_pill.VerticalAlignment = VerticalAlignment.Center
        err_pill.Visibility     = Visibility.Collapsed
        err_txt = TextBlock()
        err_txt.FontSize        = 10
        err_txt.Foreground      = hb('#F87171')
        err_txt.VerticalAlignment = VerticalAlignment.Center
        err_txt.Text            = ''
        err_pill.Child          = err_txt
        row._error_label        = err_pill
        row._error_text         = err_txt
        sp.Children.Add(err_pill)

        return sp

    # ------------------------------------------------------------------
    # Row event handlers
    # ------------------------------------------------------------------

    def _cb_click(self, sender, e):
        row = sender.Tag
        if row:
            row.Enabled = (sender.IsChecked == True)

    def _vn_lost(self, sender, e):
        row = sender.Tag
        if not row:
            return
        row.ViewName = sender.Text.strip()
        sender.Text  = row.ViewName
        # Blank = neutral, just reset border — dot stays as-is
        if not row.ViewName:
            sender.BorderBrush = hb('#208A3C')
            if row._dot and row.Status not in ('success', 'error'):
                row._dot.Fill = hb('#6B7280')
                row.Status = 'pending'
            return
        # Red border + dot if duplicate, green + grey dot if unique
        all_names = [r.ViewName for fd in self._file_data.values()
                     for r in fd['rows'] if r is not row and r.ViewName]
        if row.ViewName in all_names:
            sender.BorderBrush = hb('#DC2626')
            if row._dot:
                row._dot.Fill = hb('#DC2626')
                row.Status = 'error'
        else:
            sender.BorderBrush = hb('#208A3C')
            if row._dot and row.Status not in ('success', 'skipped'):
                row._dot.Fill = hb('#6B7280')
                row.Status = 'pending'

    def _sc_changed(self, sender, e):
        if sender.SelectedItem is None:
            return
        tag = sender.Tag
        if not isinstance(tag, tuple):
            return
        row, rc = tag
        new_sheet = sender.SelectedItem
        row.Sheet = new_sheet
        rc.Items.Clear()
        for r in row.ranges_for(new_sheet):
            rc.Items.Add(r)
        if rc.Items.Count > 0:
            rc.SelectedIndex = 0
            row.NamedRange = rc.Items[0]
            if not row.ViewName:
                row.ViewName = row.NamedRange
                if row._vn_textbox is not None:
                    row._vn_textbox.Text = row.ViewName

    def _rc_changed(self, sender, e):
        if sender.SelectedItem is None:
            return
        row = sender.Tag
        if row:
            row.NamedRange = sender.SelectedItem
            if not row.ViewName:
                row.ViewName = row.NamedRange
                if row._vn_textbox is not None:
                    row._vn_textbox.Text = row.ViewName

    def _vt_changed(self, sender, e):
        if sender.SelectedItem is None:
            return
        row = sender.Tag
        if row:
            row.ViewType = sender.SelectedItem

    def _del_click(self, sender, e):
        row = sender.Tag
        name = row.ViewName or row.NamedRange or 'this row'
        if not forms.alert(
                'Remove row: {}?'.format(name),
                title='Confirm Remove',
                yes=True, no=True):
            return
        for path, fd in self._file_data.items():
            if row in fd['rows']:
                fd['rows'].remove(row)
                # Remove from card panel
                panel = fd['card_panel']
                to_remove = None
                for child in panel.Children:
                    if getattr(child, 'Tag', None) is row:
                        to_remove = child
                        break
                if to_remove is not None:
                    panel.Children.Remove(to_remove)
                break
        self._update_footer()
        self._refresh_status_card()
        self._save_persisted_state()

    # ------------------------------------------------------------------
    # File loading
    # ------------------------------------------------------------------

    def _load_files(self, paths):
        if not paths:
            return
        for path in paths:
            ext = os.path.splitext(path)[1].lower()
            if ext in ('.xlsx', '.xls', '.ods'):
                self._parse_excel(path)
            elif ext in ('.docx', '.doc', '.odt'):
                self._parse_word(path)
        self._update_file_combo()
        self._refresh_drop_zone()

    def _parse_excel(self, path):
        if path in self._file_data:
            self._active_file = path
            return
        try:
            from script import get_named_ranges_from_workbook
            wb = get_named_ranges_from_workbook(path)
        except Exception as ex:
            logger.error('Read failed: {}'.format(ex))
            wb = {}
        sheets = wb.get('sheets', [])
        srmap  = wb.get('sheet_ranges', {})
        if not srmap:
            all_r = wb.get('named_ranges', [])
            srmap = {s: all_r for s in sheets}
        self._file_data[path] = {
            'sheets':          sheets,
            'sheet_range_map': srmap,
            'source_type':     'xl',
            'rows':            [],
            'card_panel':      None,
            'card_border':     None,
        }
        self._active_file = path
        self._make_card(path)

    def _parse_word(self, path):
        if path in self._file_data:
            self._active_file = path
            return
        # Extract headings from docx so section picker is populated
        headings = []
        try:
            from script import get_word_headings
            headings = get_word_headings(path)
        except Exception as ex:
            logger.warning('pyTable word parse: {}'.format(ex))
        if not headings:
            headings = ['(no headings found)']
        srmap = {'Document': headings}
        self._file_data[path] = {
            'sheets':          ['Document'],
            'sheet_range_map': srmap,
            'source_type':     'word',
            'rows':            [],
            'card_panel':      None,
            'card_border':     None,
            'sheet_size':      'A3 Landscape',
            'col_count':       2,
        }
        self._active_file = path
        self._make_card(path)

    # ------------------------------------------------------------------
    # Toolbar handlers
    # ------------------------------------------------------------------

    def OnBrowse(self, sender, e):
        import clr
        clr.AddReference('System.Windows.Forms')
        import System.Windows.Forms as WF
        dlg            = WF.OpenFileDialog()
        dlg.Title      = 'Select Excel or Word files'
        dlg.Filter     = 'Supported|*.xlsx;*.xls;*.docx;*.doc;*.ods;*.odt|Excel|*.xlsx;*.xls|Word|*.docx;*.doc|LibreOffice|*.ods;*.odt'
        dlg.Multiselect = True
        if dlg.ShowDialog() == WF.DialogResult.OK:
            self._load_files(list(dlg.FileNames))

    def OnImportFolder(self, sender, e):
        import clr
        clr.AddReference('System.Windows.Forms')
        import System.Windows.Forms as WF
        dlg             = WF.FolderBrowserDialog()
        dlg.Description = 'Select folder'
        if dlg.ShowDialog() == WF.DialogResult.OK:
            supported = ('.xlsx', '.xls', '.docx', '.doc', '.ods', '.odt')
            files = [os.path.join(dlg.SelectedPath, f)
                     for f in os.listdir(dlg.SelectedPath)
                     if os.path.splitext(f)[1].lower() in supported]
            self._load_files(files)

    def OnDrop(self, sender, e):
        import System
        if e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop):
            self._load_files(list(
                e.Data.GetData(System.Windows.DataFormats.FileDrop)))

    def OnDragOver(self, sender, e):
        import System
        if e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop):
            e.Effects = System.Windows.DragDropEffects.Copy
        else:
            e.Effects = getattr(System.Windows.DragDropEffects, 'None')
        e.Handled = True

    def OnAddRow(self, sender, e):
        if not self._active_file:
            self.OnBrowse(sender, e)
            return
        fd  = self._file_data.get(self._active_file, {})
        row = Row(
            file_path      = self._active_file,
            source_type    = fd.get('source_type', 'xl'),
            sheets         = fd.get('sheets', []),
            sheet_range_map= fd.get('sheet_range_map', {}),
        )
        fd['rows'].append(row)
        row_ui = self._make_row_ui(row)
        row_ui.Tag = row
        fd['card_panel'].Children.Add(row_ui)
        self._update_footer()
        self._refresh_status_card()
        try:
            self.RowsScroll.ScrollToBottom()
        except Exception:
            pass

    def OnRefresh(self, sender, e):
        for fd in self._file_data.values():
            for row in fd['rows']:
                row.LastModified = row._mtime(row.FilePath)
        self._refresh_status_card()
        self._set_status('Refreshed')

    def OnBatchActions(self, sender, e):
        from System.Windows.Controls import ContextMenu, MenuItem
        menu = ContextMenu()
        def item(label, fn):
            mi = MenuItem()
            mi.Header = label
            mi.Click += fn
            return mi

        def all_rows():
            return [r for fd in self._file_data.values() for r in fd['rows']]

        menu.Items.Add(item('Select all',
            lambda s, ev: [setattr(r, 'Enabled', True) for r in all_rows()]))
        menu.Items.Add(item('Deselect all',
            lambda s, ev: [setattr(r, 'Enabled', False) for r in all_rows()]))
        menu.Items.Add(item('Set all \u2192 Schedule View',
            lambda s, ev: [setattr(r, 'ViewType', 'Schedule View')
                           for r in all_rows()]))
        menu.Items.Add(item('Set all \u2192 Legend View',
            lambda s, ev: [setattr(r, 'ViewType', 'Legend View')
                           for r in all_rows()]))
        menu.IsOpen = True

    def OnMenuOpen(self, sender, e):
        pass  # reserved for future options panel

    def OnApply(self, sender, e):
        all_rows = [r for fd in self._file_data.values()
                    for r in fd['rows'] if r.Enabled]
        if not all_rows:
            self._set_status('No rows enabled.')
            return

        # ── Reset all dots to pending before re-running ──────────────
        for r in all_rows:
            if r.Status not in ('pending',):
                r.Status = 'pending'
                if r._dot:
                    r._dot.Fill = hb('#6B7280')
            if r._error_label is not None:
                r._error_label.Visibility = Visibility.Collapsed
                if r._error_text is not None:
                    r._error_text.Text = ''

        # ── Validation — mark invalid rows, skip them but run valid ones ──
        # Word rows use NamedRange (section heading) as identifier —
        # they have no ViewName so bypass that check entirely.
        seen_names = {}
        skip_rows  = set()
        for r in all_rows:
            if r.SourceType == 'word':
                # Word row is valid as long as it has a section selected
                if not r.NamedRange.strip():
                    r.Status = 'skipped'
                    if r._dot:
                        r._dot.Fill = hb('#CA8A04')
                    skip_rows.add(id(r))
                continue
            name = r.ViewName.strip()
            if not name:
                r.Status = 'skipped'
                if r._dot:
                    r._dot.Fill = hb('#CA8A04')
                skip_rows.add(id(r))
            elif name in seen_names:
                r.Status = 'error'
                if r._dot:
                    r._dot.Fill = hb('#DC2626')
                skip_rows.add(id(r))
            else:
                seen_names[name] = r

        # Re-check at run time — Word rows pass if NamedRange set,
        # Excel rows pass if ViewName non-blank
        run_rows = [r for r in all_rows
                    if id(r) not in skip_rows
                    and (r.SourceType == 'word'
                         or r.ViewName.strip())]
        if skip_rows:
            self._refresh_status_card()
        if not run_rows:
            self._set_status('{} row{} need attention — nothing to apply.'.format(
                len(skip_rows), 's' if len(skip_rows) > 1 else ''))
            return
        self._set_status('Applying...')
        # Show progress bar
        if self._progress_bar is not None:
            self._progress_bar.Value      = 0
            self._progress_bar.Visibility = Visibility.Visible
        from script import apply_row, TableRow, apply_notes_row
        results = []
        total_run = len(run_rows)

        # Group word rows by file so they go into one combined view per doc
        import collections as _col
        _word_by_file = _col.OrderedDict()
        _xl_rows      = []
        for r in run_rows:
            if r.SourceType == 'word':
                _word_by_file.setdefault(r.FilePath, []).append(r)
            else:
                _xl_rows.append(r)

        # Apply word files first (one view per docx)
        for fpath, wrows in _word_by_file.items():
            fd = self._file_data.get(fpath, {})
            view_name  = (self._file_data.get(fpath, {})
                          .get('view_name') or
                          os.path.splitext(os.path.basename(fpath))[0])
            view_type  = wrows[0].ViewType
            sheet_size = fd.get('sheet_size', 'A3')
            col_count  = fd.get('col_count', 2)
            from script import read_word_sections, get_word_headings
            raw_sections = read_word_sections(fpath)
            labels       = get_word_headings(fpath)
            # Map display label -> section dict
            label_to_sec = {}
            for lbl, sec in zip(labels, [s for s in raw_sections
                                          if s.get('heading')]):
                label_to_sec[lbl] = sec
            sections_payload = []
            for wr in wrows:
                sec = label_to_sec.get(wr.NamedRange)
                if sec:
                    sections_payload.append({
                        'heading':    sec['heading'],
                        'paragraphs': sec['paragraphs'],
                        'col':        wr.ColNo,
                    })
            result = apply_notes_row(
                sections_payload, view_name, view_type,
                sheet_size, col_count, fpath, size_mm=2.3)
            results.append(result)
            for wr in wrows:
                wr.Status = result.get('status', 'error')
                wr.ViewName = view_name
                if wr.Status == 'success':
                    try:
                        wr._applied_mtime = os.path.getmtime(fpath)
                        from script import _hash_word_section
                        wr._applied_hash = _hash_word_section(
                            fpath, wr.NamedRange)
                    except Exception:
                        pass
                if wr._dot is not None:
                    wr._dot.Fill = hb(STATUS_COLOURS.get(
                        wr.Status, '#3A4A3A'))
                if wr._refresh_btn is not None:
                    wr._refresh_btn.Visibility = (
                        Visibility.Visible if wr.Status == 'sync'
                        else Visibility.Collapsed)
                if wr._error_label is not None:
                    if wr.Status == 'error':
                        msg = result.get('message', '')
                        if wr._error_text is not None:
                            wr._error_text.Text = self._short_error(msg)
                        wr._error_label.Visibility = Visibility.Visible
                    else:
                        if wr._error_text is not None:
                            wr._error_text.Text = ''
                        wr._error_label.Visibility = Visibility.Collapsed
            self._refresh_status_card()

        for i, row in enumerate(_xl_rows):
            self._set_status('Processing {} of {}...'.format(
                i + 1, len(_xl_rows)))
            if self._progress_bar is not None:
                self._progress_bar.Value = int(
                    (i / float(total_run)) * 100)
            tr             = TableRow()
            tr.view_name   = row.ViewName
            tr.named_range = row.NamedRange
            tr.sheet_name  = row.Sheet
            tr.view_type   = row.ViewType
            tr.view_scale  = 1
            tr.file_path   = row.FilePath
            tr.auto_sync   = False
            result = apply_row(tr)
            results.append(result)
            row.Status = result.get('status', 'error')
            if row.Status == 'success':
                try:
                    row._applied_mtime = os.path.getmtime(row.FilePath)
                    from script import _hash_range
                    row._applied_hash = _hash_range(
                        row.FilePath, row.NamedRange, row.Sheet)
                except Exception:
                    pass
            if row._dot is not None:
                row._dot.Fill = hb(STATUS_COLOURS.get(
                    row.Status, '#3A4A3A'))
            if row._refresh_btn is not None:
                row._refresh_btn.Visibility = (
                    Visibility.Visible
                    if row.Status == 'sync'
                    else Visibility.Collapsed)
            # Show/hide inline error label
            if row._error_label is not None:
                if row.Status == 'error':
                    msg = result.get('message', '')
                    if row._error_text is not None:
                        row._error_text.Text    = self._short_error(msg)
                    row._error_label.Visibility = Visibility.Visible
                else:
                    if row._error_text is not None:
                        row._error_text.Text    = ''
                    row._error_label.Visibility = Visibility.Collapsed
            self._refresh_status_card()
        ok  = sum(1 for r in results if r['status'] == 'success')
        sk  = sum(1 for r in results if r['status'] == 'skipped')
        err = sum(1 for r in results if r['status'] == 'error')
        self._set_status('{} ok  {} skip  {} err'.format(ok, sk, err))
        # Save state to Revit parameter
        self._save_persisted_state()
        # Complete and hide progress bar
        if self._progress_bar is not None:
            self._progress_bar.Value      = 100
            import threading, System
            def _hide():
                import time
                time.sleep(1.0)
                self._progress_bar.Dispatcher.Invoke(
                    System.Action(lambda: setattr(
                        self._progress_bar, 'Visibility', Visibility.Collapsed)))
            threading.Thread(target=_hide).start()
        self._refresh_status_card()

    # ------------------------------------------------------------------
    # Status card
    # ------------------------------------------------------------------

    def _all_rows(self):
        return [r for fd in self._file_data.values() for r in fd['rows']]

    def _build_status_card(self):
        """Create the status summary card — always visible above file cards."""
        card = Border()
        card.Background   = hb('#2B3340')
        card.CornerRadius = CornerRadius(8)
        card.Padding      = Thickness(14, 10, 14, 10)
        card.Margin       = Thickness(0, 0, 0, 0)
        try:
            from System.Windows.Media.Effects import DropShadowEffect
            fx = DropShadowEffect()
            fx.Color       = Color.FromRgb(0, 0, 0)
            fx.Opacity     = 0.2
            fx.ShadowDepth = 1
            fx.BlurRadius  = 4
            card.Effect    = fx
        except Exception:
            pass

        # Progress bar — hidden when idle, fills left-to-right during Apply
        from System.Windows.Controls import ProgressBar
        self._progress_bar = ProgressBar()
        self._progress_bar.Minimum   = 0
        self._progress_bar.Maximum   = 100
        self._progress_bar.Value     = 0
        self._progress_bar.Height    = 3
        self._progress_bar.Margin    = Thickness(0, 0, 0, 8)
        self._progress_bar.Visibility = Visibility.Collapsed
        try:
            self._progress_bar.Foreground = hb('#208A3C')
            self._progress_bar.Background = hb('#404553')
            self._progress_bar.BorderThickness = Thickness(0)
        except Exception:
            pass

        row = StackPanel()
        row.Orientation      = Orientation.Horizontal
        row.VerticalAlignment = VerticalAlignment.Center

        def _item(dot_colour, label_text):
            """One dot + label pair."""
            sp = StackPanel()
            sp.Orientation       = Orientation.Horizontal
            sp.VerticalAlignment = VerticalAlignment.Center
            sp.Margin            = Thickness(0, 0, 24, 0)
            d = Ellipse()
            d.Width              = 10
            d.Height             = 10
            d.Fill               = hb(dot_colour)
            d.VerticalAlignment  = VerticalAlignment.Center
            d.Margin             = Thickness(0, 0, 7, 0)
            lbl = TextBlock()
            lbl.FontSize         = 12
            lbl.Foreground       = hb('#F4FAFF')
            lbl.VerticalAlignment = VerticalAlignment.Center
            sp.Children.Add(d)
            sp.Children.Add(lbl)
            lbl.Text = label_text
            return sp, lbl

        sp_p, self._lbl_pending = _item('#6B7280', 'Pending')
        sp_s, self._lbl_success = _item('#16A34A', 'Success')
        sp_e, self._lbl_error   = _item('#DC2626', 'Error')
        sp_k, self._lbl_skipped = _item('#CA8A04', 'Skipped')

        for sp in (sp_p, sp_s, sp_e, sp_k):
            row.Children.Add(sp)

        # Divider
        div = Border()
        div.Width             = 1
        div.Height            = 16
        div.Background        = hb('#404553')
        div.VerticalAlignment = VerticalAlignment.Center
        div.Margin            = Thickness(0, 0, 24, 0)
        row.Children.Add(div)

        # Sync indicator (hidden until needed)
        self._sync_panel = StackPanel()
        self._sync_panel.Orientation      = Orientation.Horizontal
        self._sync_panel.VerticalAlignment = VerticalAlignment.Center
        self._sync_panel.Visibility       = Visibility.Visible

        sync_dot = Ellipse()
        sync_dot.Width             = 10
        sync_dot.Height            = 10
        sync_dot.Fill              = hb('#3B82F6')
        sync_dot.VerticalAlignment = VerticalAlignment.Center
        sync_dot.Margin            = Thickness(0, 0, 7, 0)

        self._sync_label = TextBlock()
        self._sync_label.FontSize          = 12
        self._sync_label.Foreground        = hb('#F4FAFF')
        self._sync_label.VerticalAlignment = VerticalAlignment.Center

        self._sync_label.Text    = '0 update required'
        self._sync_label.Opacity = 0.9
        self._sync_panel.Children.Add(sync_dot)
        self._sync_panel.Children.Add(self._sync_label)
        row.Children.Add(self._sync_panel)

        # Blurb
        blurb = TextBlock()
        blurb.Text = (
            'View Name auto-fills from the selected range. '
            'Clear the View Name to allow it to update when the range changes.'
        )
        blurb.FontSize    = 10
        blurb.Foreground  = hb('#F4FAFF')
        blurb.Opacity     = 0.35
        from System.Windows import TextWrapping
        blurb.TextWrapping = TextWrapping.Wrap
        blurb.Margin      = Thickness(0, 8, 0, 0)

        outer_stack = StackPanel()
        outer_stack.Orientation = Orientation.Vertical
        outer_stack.Children.Add(self._progress_bar)
        outer_stack.Children.Add(row)
        outer_stack.Children.Add(blurb)

        card.Child            = outer_stack
        self._status_card     = card
        self.StatusCardHost.Child = card

    def _refresh_status_card(self):
        """Update counts and sync indicator on the status card."""
        if self._status_card is None or self._lbl_pending is None:
            return

        rows = self._all_rows()
        counts = {'pending': 0, 'success': 0, 'error': 0, 'skipped': 0}
        for r in rows:
            s = r.Status
            if s == 'sync':
                counts['success'] += 1  # view exists, counts as success
            elif s in counts:
                counts[s] += 1
            else:
                counts['pending'] += 1

        # Format: "2 Pending" — number first, muted if zero
        def _fmt(n, label):
            return '{} {}'.format(n, label)

        self._lbl_pending.Text    = _fmt(counts['pending'], 'Pending')
        self._lbl_success.Text    = _fmt(counts['success'], 'Success')
        self._lbl_error.Text      = _fmt(counts['error'],   'Error')
        self._lbl_skipped.Text    = _fmt(counts['skipped'], 'Skipped')

        # Dim zero-count items
        self._lbl_pending.Opacity  = 0.4 if counts['pending']  == 0 else 0.9
        self._lbl_success.Opacity  = 0.4 if counts['success']  == 0 else 0.9
        self._lbl_error.Opacity    = 0.4 if counts['error']     == 0 else 0.9
        self._lbl_skipped.Opacity  = 0.4 if counts['skipped']   == 0 else 0.9

        # Out-of-sync check — use content hash if available, else mtime
        stale = []
        seen_rows = set()
        for r in rows:
            key = '{}|{}'.format(r.FilePath, r.NamedRange)
            if key in seen_rows:
                continue
            seen_rows.add(key)
            try:
                if r.SourceType == 'word':
                    # Per-section hash — same pattern as Excel
                    if r._applied_hash:
                        from script import _hash_word_section
                        cur = _hash_word_section(r.FilePath, r.NamedRange)
                        if cur and cur != r._applied_hash:
                            stale.append(os.path.basename(r.FilePath))
                    elif r._applied_mtime:
                        if os.path.getmtime(r.FilePath) > r._applied_mtime + 1:
                            stale.append(os.path.basename(r.FilePath))
                elif r._applied_hash:
                    from script import _hash_range
                    current_hash = _hash_range(r.FilePath, r.NamedRange, r.Sheet)
                    if current_hash and current_hash != r._applied_hash:
                        stale.append(os.path.basename(r.FilePath))
                elif r._applied_mtime:
                    if os.path.getmtime(r.FilePath) > r._applied_mtime + 1:
                        stale.append(os.path.basename(r.FilePath))
            except Exception:
                pass

        # Sync indicator is always blue — acts as legend + counter
        n_stale = len(stale)
        self._sync_label.Text       = '{} update required'.format(n_stale)
        self._sync_label.Foreground = hb('#F4FAFF')
        self._sync_label.Opacity    = 0.4 if n_stale == 0 else 0.9

    # ------------------------------------------------------------------
    # Persistence
    # ------------------------------------------------------------------

    def _save_persisted_state(self):
        """Write current UI state to the pyTable shared parameter."""
        try:
            from script import save_pytable_state
            save_pytable_state(self._file_data)
        except Exception as ex:
            logger.warning('pyTable save: {}'.format(ex))

    def _load_persisted_state(self):
        """
        Read pyTable shared parameter and rebuild cards + rows.
        After rebuilding, check if any source files have been modified
        since the parameter was last written (triggers sync indicator).
        """
        try:
            from script import load_pytable_state, get_named_ranges_from_workbook
        except Exception:
            return
        try:
            cards = load_pytable_state()
        except Exception as ex:
            logger.warning('pyTable load: {}'.format(ex))
            return
        if not cards:
            return

        for card in cards:
            path = card.get('path', '')
            if not path or not os.path.exists(path):
                logger.warning('pyTable restore: file not found: {}'.format(path))
                continue
            # Parse the file to get current sheets/ranges
            if path not in self._file_data:
                ext     = os.path.splitext(path)[1].lower()
                is_word = ext not in ('.xlsx', '.xls', '.ods')
                if is_word:
                    # Word file — use heading parser, not workbook parser
                    try:
                        from script import get_word_headings
                        headings = get_word_headings(path)
                    except Exception:
                        headings = []
                    if not headings:
                        headings = ['(no headings found)']
                    sheets = ['Document']
                    srmap  = {'Document': headings}
                else:
                    try:
                        wb = get_named_ranges_from_workbook(path)
                    except Exception:
                        wb = {}
                    sheets = wb.get('sheets', [])
                    srmap  = wb.get('sheet_ranges', {})
                    if not srmap:
                        all_r = wb.get('named_ranges', [])
                        srmap = {s: all_r for s in sheets}
                self._file_data[path] = {
                    'sheets':          sheets,
                    'sheet_range_map': srmap,
                    'source_type':     'word' if is_word else 'xl',
                    'rows':            [],
                    'card_panel':      None,
                    'card_border':     None,
                    'sheet_size':      card.get('sheet_size', 'A3 Landscape'),
                    'col_count':       card.get('col_count', 2),
                    'view_name':       card.get('view_name', ''),
                }
                self._make_card(path)

            # Restore rows
            fd = self._file_data[path]
            for rdata in card.get('rows', []):
                row = Row(
                    file_path      = path,
                    source_type    = fd['source_type'],
                    sheets         = fd['sheets'],
                    sheet_range_map= fd['sheet_range_map'],
                )
                row.ViewName        = rdata.get('view_name', '')
                row.Sheet           = rdata.get('sheet', '')
                row.ColNo           = int(rdata.get('col_no', 1))
                row.NamedRange      = rdata.get('named_range', '')
                row.ViewType        = rdata.get('view_type', 'Schedule View')
                row._applied_mtime  = rdata.get('applied_mtime', None)
                row._applied_hash   = rdata.get('applied_hash', None)

                # ── Check Revit view existence and sync status ────────
                row.Status = self._check_row_status(row, path)

                fd['rows'].append(row)
                row_ui     = self._make_row_ui(row)
                row_ui.Tag = row
                # Sync ViewName textbox with restored value
                if row._vn_textbox is not None:
                    row._vn_textbox.Text = row.ViewName
                # Apply dot colour from restored status
                if row._dot is not None:
                    row._dot.Fill = hb(STATUS_COLOURS.get(row.Status, '#6B7280'))
                if row._refresh_btn is not None:
                    row._refresh_btn.Visibility = (
                        Visibility.Visible
                        if row.Status == 'sync'
                        else Visibility.Collapsed)
                fd['card_panel'].Children.Add(row_ui)

        self._update_file_combo()
        self._update_footer()
        self._refresh_status_card()
        logger.debug('pyTable: restored {} card(s)'.format(len(cards)))

    def _check_row_status(self, row, file_path):
        """
        On load, determine row status by checking:
        1. Does the Revit view exist with that name?
        2. If yes, has the source content changed since last apply?

        Returns: 'success', 'pending', or 'sync'.
        """
        try:
            from pyrevit import revit, DB
            from pyrevit.revit import query

            # Word rows use the card-level view name (filename stem)
            if row.SourceType == 'word':
                fd = self._file_data.get(file_path, {})
                view_name = (fd.get('view_name') or
                             os.path.splitext(os.path.basename(file_path))[0])
            else:
                view_name = row.ViewName.strip()
            if not view_name:
                return 'pending'

            # Check if a view with this name exists
            found = False
            for v in query.get_elements_by_class(DB.View, doc=revit.doc):
                try:
                    if v.Name == view_name:
                        found = True
                        break
                except Exception:
                    continue

            if not found:
                return 'pending'

            # View exists — compare content hash to detect changes
            try:
                if row.SourceType == 'word':
                    # Per-section hash for Word rows
                    from script import _hash_word_section
                    current_hash = _hash_word_section(file_path, row.NamedRange)
                    if current_hash and row._applied_hash:
                        if current_hash != row._applied_hash:
                            return 'sync'
                        else:
                            row._applied_hash = current_hash
                            return 'success'
                    elif current_hash and not row._applied_hash:
                        row._applied_hash = current_hash
                        return 'success'
                    return 'success'
                from script import _hash_range
                current_hash = _hash_range(file_path, row.NamedRange, row.Sheet)

                if current_hash and row._applied_hash:
                    if current_hash != row._applied_hash:
                        # Content changed — blue dot
                        return 'sync'
                    else:
                        # Up to date
                        row._applied_mtime = os.path.getmtime(file_path)
                        row._applied_hash  = current_hash
                        return 'success'
                elif current_hash and not row._applied_hash:
                    # First open after hash tracking added — store now
                    row._applied_hash  = current_hash
                    row._applied_mtime = os.path.getmtime(file_path)
                    return 'success'
                else:
                    # Hash failed — mtime fallback
                    file_mtime = os.path.getmtime(file_path)
                    if row._applied_mtime and file_mtime > row._applied_mtime + 1:
                        return 'sync'
                    row._applied_mtime = file_mtime
                    return 'success'
            except Exception as ex:
                logger.warning('pyTable hash check failed: {}'.format(ex))
                return 'success'

        except Exception as ex:
            logger.debug('_check_row_status failed: {}'.format(ex))
            return 'pending'

    def _refresh_row_click(self, sender, e):
        """Re-apply a single stale (blue) row without touching others."""
        row = sender.Tag
        if row is None:
            return
        # Clear error pill
        if row._error_label is not None:
            row._error_label.Visibility = Visibility.Collapsed
        if row._error_text is not None:
            row._error_text.Text = ''
        # Set dot to pending while running
        if row._dot is not None:
            row._dot.Fill = hb('#6B7280')
        if row._refresh_btn is not None:
            row._refresh_btn.Visibility = Visibility.Collapsed
        self._set_status('Refreshing {}...'.format(
            row.NamedRange if row.SourceType == 'word' else row.ViewName))
        try:
            if row.SourceType == 'word':
                # Word row: re-apply just this section via apply_notes_row
                from script import (apply_notes_row, read_word_sections,
                                    get_word_headings, _hash_word_section)
                fpath      = row.FilePath
                fd         = self._file_data.get(fpath, {})
                view_name  = (fd.get('view_name') or
                              os.path.splitext(os.path.basename(fpath))[0])
                raw_secs   = read_word_sections(fpath)
                labels     = get_word_headings(fpath)
                label_map  = {lbl: s for lbl, s in zip(
                    labels, [s for s in raw_secs if s.get('heading')])}
                sec = label_map.get(row.NamedRange)
                sections_payload = []
                if sec:
                    sections_payload.append({
                        'heading':    sec['heading'],
                        'paragraphs': sec['paragraphs'],
                        'col':        row.ColNo,
                    })
                # Collect all rows for same file to rebuild view —
                # include success, sync, and the current row being refreshed
                all_word_rows = [r for r in fd.get('rows', [])
                                 if r.Status in ('success', 'sync')
                                 or r is row]
                full_payload  = []
                for wr in all_word_rows:
                    s = label_map.get(wr.NamedRange)
                    if s:
                        full_payload.append({
                            'heading':    s['heading'],
                            'paragraphs': s['paragraphs'],
                            'col':        wr.ColNo,
                        })
                result = apply_notes_row(
                    full_payload, view_name,
                    row.ViewType,
                    fd.get('sheet_size', 'A3 Landscape'),
                    fd.get('col_count', 2),
                    fpath, size_mm=2.3)
                row.Status = result.get('status', 'error')
                if row.Status == 'success':
                    try:
                        row._applied_mtime = os.path.getmtime(fpath)
                        row._applied_hash  = _hash_word_section(
                            fpath, row.NamedRange)
                    except Exception:
                        pass
            else:
                from script import apply_row, TableRow, _hash_range
                tr             = TableRow()
                tr.view_name   = row.ViewName
                tr.named_range = row.NamedRange
                tr.sheet_name  = row.Sheet
                tr.view_type   = row.ViewType
                tr.view_scale  = 1
                tr.file_path   = row.FilePath
                tr.auto_sync   = False
                result = apply_row(tr)
                row.Status = result.get('status', 'error')
                if row.Status == 'success':
                    try:
                        row._applied_mtime = os.path.getmtime(row.FilePath)
                        row._applied_hash  = _hash_range(
                            row.FilePath, row.NamedRange, row.Sheet)
                    except Exception:
                        pass
            if row._dot is not None:
                row._dot.Fill = hb(STATUS_COLOURS.get(
                    row.Status, '#3A4A3A'))
            if row._refresh_btn is not None:
                row._refresh_btn.Visibility = (
                    Visibility.Visible
                    if row.Status == 'sync'
                    else Visibility.Collapsed)
            if row._error_label is not None:
                if row.Status == 'error':
                    msg = result.get('message', '')
                    if row._error_text is not None:
                        row._error_text.Text = self._short_error(msg)
                    row._error_label.Visibility = Visibility.Visible
            self._set_status('{} refreshed.'.format(row.ViewName))
        except Exception as ex:
            row.Status = 'error'
            if row._dot is not None:
                row._dot.Fill = hb(STATUS_COLOURS.get('error', '#DC2626'))
            self._set_status('Refresh failed: {}'.format(ex))
        self._save_persisted_state()
        self._refresh_status_card()

    # ------------------------------------------------------------------
    # Word row event handlers
    # ------------------------------------------------------------------

    def _word_view_name_changed(self, sender, e):
        path = sender.Tag
        if path and path in self._file_data:
            self._file_data[path]['view_name'] = sender.Text.strip()
            self._save_persisted_state()

    def _word_section_changed(self, sender, e):
        row = sender.Tag
        if row is None:
            return
        row.NamedRange = sender.SelectedItem or ''
        self._save_persisted_state()

    def _word_col_no_changed(self, sender, e):
        row = sender.Tag
        if row is None:
            return
        try:
            row.ColNo = max(1, int(sender.Text.strip()))
        except Exception:
            row.ColNo = 1
        sender.Text = str(row.ColNo)
        self._save_persisted_state()

    def _word_sheet_size_changed(self, sender, e):
        path = sender.Tag
        if path and path in self._file_data:
            self._file_data[path]['sheet_size'] = (
                sender.SelectedItem or 'A3')
            self._save_persisted_state()

    def _word_col_count_changed(self, sender, e):
        path = sender.Tag
        if path and path in self._file_data:
            try:
                n = max(1, int(sender.Text.strip()))
            except Exception:
                n = 2
            self._file_data[path]['col_count'] = n
            sender.Text = str(n)
            self._save_persisted_state()

    # ------------------------------------------------------------------
    # Drag reorder (Word rows)
    # ------------------------------------------------------------------

    def _drag_indicator_show(self, panel, idx):
        """Position or create the blue drop-line indicator in *panel*."""
        # Create indicator Border on first call
        if getattr(self, '_drag_indicator', None) is None:
            ind = Border()
            ind.Height          = 2
            ind.Background      = hb('#3B82F6')
            ind.IsHitTestVisible = False   # don't interfere with mouse events
            ind.Margin          = Thickness(0)
            self._drag_indicator       = ind
            self._drag_indicator_panel = None

        ind = self._drag_indicator

        # Remove from previous panel if different
        if (getattr(self, '_drag_indicator_panel', None) is not None
                and self._drag_indicator_panel is not panel):
            try:
                self._drag_indicator_panel.Children.Remove(ind)
            except Exception:
                pass

        # Insert/move in current panel
        self._drag_indicator_panel = panel
        try:
            panel.Children.Remove(ind)
        except Exception:
            pass
        insert_at = min(idx, panel.Children.Count)
        try:
            panel.Children.Insert(insert_at, ind)
        except Exception:
            pass

    def _drag_indicator_remove(self):
        """Remove the drop-line indicator from whichever panel it's in."""
        ind = getattr(self, '_drag_indicator', None)
        if ind is None:
            return
        panel = getattr(self, '_drag_indicator_panel', None)
        if panel is not None:
            try:
                panel.Children.Remove(ind)
            except Exception:
                pass
        self._drag_indicator       = None
        self._drag_indicator_panel = None

    def _drag_start(self, sender, e):
        """Record drag origin and wire move/up events on the window."""
        row = sender.Tag
        if row is None:
            return
        self._dragging_row    = row
        self._drag_start_y    = e.GetPosition(self).Y
        self._drag_indicator  = None
        self._drag_target_idx = None
        sender.CaptureMouse()
        sender.PreviewMouseMove             += self._drag_move
        sender.PreviewMouseLeftButtonUp     += self._drag_drop
        e.Handled = True

    def _drag_move(self, sender, e):
        """Update drop-line position as the user drags."""
        if not getattr(self, '_dragging_row', None):
            return
        row = self._dragging_row
        for path, fd in self._file_data.items():
            if row not in fd['rows']:
                continue
            panel   = fd['card_panel']
            mouse_y = e.GetPosition(panel).Y
            idx     = 0
            for i in range(panel.Children.Count):
                child = panel.Children[i]
                # Skip the indicator itself when measuring
                if child is getattr(self, '_drag_indicator', None):
                    continue
                try:
                    pt = child.TransformToAncestor(panel).Transform(
                        __import__('System.Windows', fromlist=['Point'])
                        .Point(0, child.ActualHeight / 2.0))
                    if mouse_y > pt.Y:
                        idx = i + 1
                except Exception:
                    pass
            self._drag_target_idx = idx
            self._drag_indicator_show(panel, idx)
            break
        e.Handled = True

    def _drag_drop(self, sender, e):
        """Reorder the row to the computed insertion index."""
        row = getattr(self, '_dragging_row', None)
        if row is None:
            return
        sender.ReleaseMouseCapture()
        sender.PreviewMouseMove         -= self._drag_move
        sender.PreviewMouseLeftButtonUp -= self._drag_drop
        self._drag_indicator_remove()
        target_idx = getattr(self, '_drag_target_idx', None)
        self._dragging_row = None
        if target_idx is None:
            return

        for path, fd in self._file_data.items():
            if row not in fd['rows']:
                continue
            rows  = fd['rows']
            panel = fd['card_panel']
            cur   = rows.index(row)
            if cur == target_idx or cur + 1 == target_idx:
                return
            # Find the row's StackPanel child
            row_sp = None
            for child in list(panel.Children):
                if getattr(child, 'Tag', None) is row:
                    row_sp = child
                    break
            if row_sp is None:
                return
            # Reorder in data list
            rows.remove(row)
            ins = target_idx if target_idx <= cur else target_idx - 1
            rows.insert(ins, row)
            # Reorder in UI panel
            panel.Children.Remove(row_sp)
            panel.Children.Insert(ins, row_sp)
            self._save_persisted_state()
            break
        e.Handled = True

    def _on_window_closing(self, sender, e):
        """Save state when window is closed so deletions are persisted."""
        try:
            self._save_persisted_state()
        except Exception:
            pass

    def _green_btn(self, content, height=24, padding=(10,0,10,0),
                   font_size=11, width=None):
        """Rounded green button matching SmallButtonStyle."""
        btn = Button()
        btn.Content = content
        btn.Height  = height
        btn.Padding = Thickness(padding[0], padding[1], padding[2], padding[3])
        btn.FontSize = font_size
        if width is not None:
            btn.Width = width
        try:
            btn.Style = self.FindResource('SmallButtonStyle')
        except Exception:
            btn.Background      = hb('#208A3C')
            btn.Foreground      = hb('#F4FAFF')
            btn.BorderThickness = Thickness(0)
            btn.FontWeight      = FontWeights.SemiBold
            try:
                from System.Windows.Markup import XamlReader
                tmpl = (
                    '<ControlTemplate xmlns="http://schemas.microsoft.com'
                    '/winfx/2006/xaml/presentation" '
                    'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" '
                    'TargetType="Button">'
                    '<Border Background="{TemplateBinding Background}" '
                    'CornerRadius="4" '
                    'Padding="{TemplateBinding Padding}">'
                    '<ContentPresenter HorizontalAlignment="Center" '
                    'VerticalAlignment="Center"/>'
                    '</Border>'
                    '</ControlTemplate>'
                )
                btn.Template = XamlReader.Parse(tmpl)
            except Exception:
                pass
        return btn

    def _add_row_for_card(self, sender, e):
        """Add a new row to a specific card regardless of active file."""
        path = sender.Tag
        if path not in self._file_data:
            return
        prev_active      = self._active_file
        self._active_file = path
        self._update_file_combo()
        self._on_add_row(None, None)
        self._active_file = prev_active
        self._update_file_combo()

    def _del_card_click(self, sender, e):
        """Remove an entire file card and all its rows."""
        path = sender.Tag
        if path not in self._file_data:
            return
        name = os.path.basename(path)
        if not forms.alert(
                'Remove card for {}\nThis will delete all rows for this file.'.format(name),
                title='Confirm Remove',
                yes=True, no=True):
            return
        # Remove card Border from CardsPanel
        card_border = self._file_data[path].get('card_border')
        if card_border is not None:
            try:
                self.CardsPanel.Children.Remove(card_border)
            except Exception:
                pass
        # Remove from data
        del self._file_data[path]
        # Update active file
        if self._active_file == path:
            paths = list(self._file_data.keys())
            self._active_file = paths[0] if paths else None
        self._update_file_combo()
        self._update_footer()
        self._refresh_status_card()
        self._refresh_drop_zone()
        self._save_persisted_state()
