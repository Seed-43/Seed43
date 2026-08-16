# -*- coding: utf-8 -*-
# pdfexport.py
"""
PDF Export Settings
Loaded by the main pySheets script when user selects the PDF sub-tab
on the Settings tab.

This module provides get_settings() and apply_settings() so the main
window can read/write PDF-specific options without knowing the internals.
"""
from collections import namedtuple

PDFSettings = namedtuple('PDFSettings', [
    'printer',
    'print_setting_name',
    'paper_size',
    'paper_source',
    'orientation',      # 'portrait' | 'landscape' | 'from_sheet'
    'placement',        # 'center' | 'offset'
    'offset_x',
    'offset_y',
    'zoom_fit',         # True = Fit to Page, False = use zoom_pct
    'zoom_pct',
    'hidden_lines',     # 'vector' | 'raster'
    'raster_quality',   # 'High' | 'Medium' | 'Low'
    'colors',           # 'Color' | 'Black and White' | 'Grayscale'
    'combine',          # True = single PDF, False = separate files
    'opt_links_blue',
    'opt_hide_refplanes',
    'opt_hide_unreftags',
    'opt_hide_scopeboxes',
    'opt_hide_cropbounds',
    'opt_halftone_thin',
    'opt_region_edges',
])

# Default PDF settings
DEFAULT_PDF_SETTINGS = PDFSettings(
    printer            = '',
    print_setting_name = '',
    paper_size         = 'From Titleblock',
    paper_source       = '<default tray>',
    orientation        = 'from_sheet',
    placement          = 'center',
    offset_x           = '0.0000',
    offset_y           = '0.0000',
    zoom_fit           = True,
    zoom_pct           = 100,
    hidden_lines       = 'vector',
    raster_quality     = 'High',
    colors             = 'Color',
    combine            = True,           # default: combine into single PDF
    opt_links_blue     = False,
    opt_hide_refplanes = True,
    opt_hide_unreftags = False,
    opt_hide_scopeboxes= True,
    opt_hide_cropbounds= True,
    opt_halftone_thin  = False,
    opt_region_edges   = False,
)


def _combo_text(cb, fallback):
    try:
        item = cb.SelectedItem
        if item is None:
            return cb.Text or fallback
        return getattr(item, 'Content', item) or fallback
    except Exception:
        return fallback


def _select_combo(cb, text):
    try:
        for item in (cb.ItemsSource or cb.Items):
            label = getattr(item, 'Content', getattr(item, 'name', item))
            if label == text:
                cb.SelectedItem = item
                return True
    except Exception:
        pass
    return False


def read_from_window(win):
    """
    Read current PDF settings from the main window controls.

    Args:
        win: PrintSheetsWindow instance

    Returns:
        PDFSettings namedtuple
    """
    try:
        orient = 'from_sheet'
        if win.orient_portrait_btn.Tag  == 'Viewing': orient = 'portrait'
        if win.orient_landscape_btn.Tag == 'Viewing': orient = 'landscape'

        return PDFSettings(
            printer            = win.printer_cb.SelectedItem or '',
            print_setting_name = (win.printsetting_cb.SelectedItem.name
                                  if win.printsetting_cb.SelectedItem else ''),
            paper_size         = _combo_text(win.papersize_cb, 'From Titleblock'),
            paper_source       = _combo_text(win.papersource_cb, '<default tray>'),
            orientation        = orient,
            placement          = ('center' if win.placement_center_rb.IsChecked
                                  else 'offset'),
            offset_x           = win.offset_x_tb.Text or '0.0000',
            offset_y           = win.offset_y_tb.Text or '0.0000',
            zoom_fit           = bool(win.zoom_fit_rb.IsChecked),
            zoom_pct           = int(win.zoom_pct_tb.Text or 100),
            hidden_lines       = ('vector' if win.hlv_vector_rb.IsChecked
                                  else 'raster'),
            raster_quality     = _combo_text(win.raster_quality_cb, 'High'),
            colors             = _combo_text(win.colors_cb, 'Color'),
            combine            = bool(win.file_combine_rb.IsChecked),
            opt_links_blue     = bool(win.opt_links_blue_cb.IsChecked),
            opt_hide_refplanes = bool(win.opt_hide_refplanes_cb.IsChecked),
            opt_hide_unreftags = bool(win.opt_hide_unreftags_cb.IsChecked),
            opt_hide_scopeboxes= bool(win.opt_hide_scopeboxes_cb.IsChecked),
            opt_hide_cropbounds= bool(win.opt_hide_cropbounds_cb.IsChecked),
            opt_halftone_thin  = bool(win.opt_halftone_thin_cb.IsChecked),
            opt_region_edges   = bool(win.opt_region_edges_cb.IsChecked),
        )
    except Exception:
        return DEFAULT_PDF_SETTINGS


def apply_to_window(win, settings=None):
    """
    Push a PDFSettings into the main window controls.

    Args:
        win:      PrintSheetsWindow instance
        settings: PDFSettings namedtuple (uses DEFAULT if None)
    """
    s = settings or DEFAULT_PDF_SETTINGS
    try:
        # Printer + print setting + paper (each guarded on its own)
        try:
            _select_combo(win.printer_cb, s.printer)
        except Exception:
            pass
        try:
            _select_combo(win.printsetting_cb, s.print_setting_name)
        except Exception:
            pass
        try:
            _select_combo(win.papersize_cb, s.paper_size)
            _select_combo(win.papersource_cb, s.paper_source)
        except Exception:
            pass

        # Orientation buttons — neither on = from each sheet
        win.orient_portrait_btn.Tag  = 'Viewing' if s.orientation == 'portrait'  else ''
        win.orient_landscape_btn.Tag = 'Viewing' if s.orientation == 'landscape' else ''

        # Paper placement
        win.placement_center_rb.IsChecked = (s.placement == 'center')
        win.placement_offset_rb.IsChecked = (s.placement == 'offset')
        win.offset_x_tb.Text = s.offset_x
        win.offset_y_tb.Text = s.offset_y

        # Zoom
        win.zoom_fit_rb.IsChecked = s.zoom_fit
        win.zoom_pct_rb.IsChecked = not s.zoom_fit
        win.zoom_pct_tb.Text      = str(s.zoom_pct)

        # Hidden lines
        win.hlv_vector_rb.IsChecked = (s.hidden_lines == 'vector')
        win.hlv_raster_rb.IsChecked = (s.hidden_lines == 'raster')

        # Appearance
        _select_combo(win.raster_quality_cb, s.raster_quality)
        _select_combo(win.colors_cb, s.colors)

        # File output
        win.file_combine_rb.IsChecked  = s.combine
        win.file_separate_rb.IsChecked = not s.combine

        # Options checkboxes
        win.opt_links_blue_cb.IsChecked      = s.opt_links_blue
        win.opt_hide_refplanes_cb.IsChecked  = s.opt_hide_refplanes
        win.opt_hide_unreftags_cb.IsChecked  = s.opt_hide_unreftags
        win.opt_hide_scopeboxes_cb.IsChecked = s.opt_hide_scopeboxes
        win.opt_hide_cropbounds_cb.IsChecked = s.opt_hide_cropbounds
        win.opt_halftone_thin_cb.IsChecked   = s.opt_halftone_thin
        win.opt_region_edges_cb.IsChecked    = s.opt_region_edges

    except Exception as e:
        pass  # silently skip if a control is missing
