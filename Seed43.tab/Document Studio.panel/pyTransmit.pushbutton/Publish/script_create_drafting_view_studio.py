# -*- coding: utf-8 -*-
# script_create_drafting_view_studio.py
#
# Drafting View writer for pyTransmit STUDIO layouts.
#
# A separate script from script_create_drafting_view.py on purpose, not a
# patch of it. The two read different layout schemas:
#
#     script_create_drafting_view.py   Layout Builder: rows of exactly 4
#                                      slots, slot 3 a "spine" that fans out
#                                      into rev_count columns, with the
#                                      column geometry inferred from col_pct.
#     this file                        Studio: a real grid - every column has
#                                      a width in mm and every row a height
#                                      in mm, so the drawing is a direct
#                                      transcription of the grid.
#
# That difference is the whole reason this is short by comparison. There is
# no spine to expand and no percentage arithmetic: a cell's rectangle is the
# sum of the column widths to its left and the row heights above it, in
# millimetres, converted to feet once.
#
# The grid, the row plan and every cell's text come from
# Studio/studio_publish.py - the same reading the Studio schedule writer
# uses, built on the same studio_rows module the Studio canvas draws from.
#
# Payload keys used:
#   layout_json_path        the Studio template to draw
#   group_params            sheet parameters to group the documentation table by
#   group_label             False = group header rows present but blank
#   page_height_mode        'none' = one continuous column, never split
#   _legend_temp_view_name  draw into this view name instead (see
#                           script_create_legend_studio.py)

_p = globals().get('PYTRANSMIT_PAYLOAD', {})

import os
import math

from pyrevit import revit, script, DB, forms

from Autodesk.Revit.DB import (
    FilteredElementCollector, XYZ, Line, TextNote, TextNoteType,
    CurveElement, ViewFamilyType, ViewFamily, ViewDrafting,
    TextNoteOptions, HorizontalTextAlignment,
    ImageType, ImageTypeOptions, ImageInstance, ImageTypeSource,
    ImagePlacementOptions, BoxPlacement,
    FilledRegion, FilledRegionType, CurveLoop,
    GraphicsStyleType, FillPatternElement, Color,
)

# ── Studio modules (pure Python: no WPF pulled in here) ──────────────────────
_PUBLISH_DIR = os.path.dirname(os.path.abspath(__file__))
_PT_ROOT = os.path.dirname(_PUBLISH_DIR)
_STUDIO_DIR = os.path.join(_PT_ROOT, 'Studio')
import sys as _sys
for _d in (_PT_ROOT, _STUDIO_DIR):
    if _d not in _sys.path:
        _sys.path.insert(0, _d)

import studio_publish
import studio_rows
import studio_live_data
from pytransmit_paths import SETTINGS_DIR

try:
    from Snippets import _dialogs as sdlg
except Exception:
    sdlg = None


def _alert(message, title='', exitscript=False):
    """Themed popup via the shared Snippets dialog lib, falls back to
    pyRevit's default forms.alert if the shared lib isn't available."""
    if sdlg:
        sdlg.message(message, title=title)
    else:
        forms.alert(message, title=title)
    if exitscript:
        script.exit()


_log_lines = _p.get('_log_lines', [])


def _log(msg):
    try:
        _log_lines.append(str(msg))
    except Exception:
        pass


doc = revit.doc
try:
    script.get_output().hide()
except Exception:
    pass

# ── Geometry ─────────────────────────────────────────────────────────────────
MM = 1.0 / 304.8
SHORT_CURVE_TOL = 0.002083333    # Revit refuses to draw anything shorter
INDENT = 0.8 * MM                # text inset from its cell edge
PAGE_GAP_MM = 5.0                # gap between overflow columns
TITLE = 'pyTransmit Studio - Drafting View'

_DEFAULT_SIZE_MM = 9.0 * studio_publish.MM_PER_PT

# ── Load the Studio layout ───────────────────────────────────────────────────
_layout_path = _p.get('layout_json_path')
if not _layout_path or not os.path.isfile(_layout_path):
    _alert('No Studio layout was assigned for the Drafting View output.',
           title=TITLE, exitscript=True)

try:
    LAYOUT = studio_publish.load_layout(_layout_path)
except ValueError:
    # Guarded in pyTransmit too, but a script that can be run directly should
    # not depend on its caller having checked.
    _alert('"{}" is not a pyTransmit Studio layout.\n\nUse '
           'script_create_drafting_view.py for Layout Builder '
           'templates.'.format(os.path.basename(_layout_path)),
           title=TITLE, exitscript=True)
except Exception as _e:
    _alert('Could not read the Studio layout:\n{}\n\n{}'.format(_layout_path, _e),
           title=TITLE, exitscript=True)

_log('Studio layout: {} ({}x{})'.format(
    os.path.basename(_layout_path),
    LAYOUT.get('n_rows'), LAYOUT.get('n_cols')))

# The same reader the Studio canvas uses, so the preview and the drafting view
# are built from one description of the model rather than two.
DATA = studio_live_data.get_live_data(SETTINGS_DIR)
SL = studio_publish.StudioLayout(LAYOUT, DATA, _p, log=_log)

PLACEMENTS = SL.placements()

# ── Page splitting ───────────────────────────────────────────────────────────
# Overflow rows go into a fresh column to the RIGHT on the same view, which is
# what script_create_drafting_view.py does: a drafting view is one continuous
# sheet of paper, so "page 2" can only mean "further across". Which rows go on
# which page is studio_publish's decision, so the Studio schedule writer
# breaks in the same places.
#
# The height comes from the LAYOUT, not from the Setup panel, matching
# script_create_drafting_view.py: Setup's page height only decides whether to
# split at all, because a drafting view is drawn at the layout's paper size.
_SPLIT = (_p.get('page_height_mode') or 'a4') != 'none'

# One entry per printed column: {'x_mm': left edge, 'rows': {v: (y_top, y_bot)}}
PAGES = []
_table_w_mm = SL.table_w()
for _pi, _page_vs in enumerate(SL.page_rows(split=_SPLIT)):
    _rows = {}
    _y = 0.0
    for _v in _page_vs:
        _h = SL.row_h(_v)
        _rows[_v] = (_y, _y - _h)
        _y -= _h
    PAGES.append({'x_mm': _pi * (_table_w_mm + PAGE_GAP_MM), 'rows': _rows})

# ── Revit types: text styles and fills ───────────────────────────────────────
# One TextNoteType per distinct font/size/weight in the layout and one
# FilledRegionType per distinct fill colour, both created up front in a single
# transaction. Creating them inside the drawing loop would open a transaction
# per cell on a 2000-sheet transmittal.


def _hex_to_rgb(value, default=(255, 255, 255)):
    try:
        h = str(value or '').strip().lstrip('#')
        if len(h) == 3:
            h = h[0] * 2 + h[1] * 2 + h[2] * 2
        if len(h) == 6:
            return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except Exception:
        pass
    return default


def _font_key(block):
    """(font, size in mm, bold, italic) - what a TextNoteType has to match."""
    b = block or {}
    return (b.get('font') or 'Arial',
            round(float(b.get('size_mm') or _DEFAULT_SIZE_MM), 3),
            bool(b.get('bold')), bool(b.get('italic')))


def _fill_hex(placement):
    """The fill a placement wants, as a hex string, or '' for no fill."""
    block = placement['block'] or {}
    if placement['kind'] == 'space':
        # A deliberate gap: no fill, no rules, nothing. Drawn with the block's
        # colours it would read as an empty data row rather than a separator.
        return ''
    if placement['kind'] == 'group':
        return block.get('group_color') or block.get('bg_color') or '#E8E8E8'
    if placement['alt']:
        return block.get('alt_color') or '#F5F7FA'
    return block.get('bg_color') or ''


def _get_line_style(name):
    try:
        lines_cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
        for sub in lines_cat.SubCategories:
            if sub.Name == name:
                return sub.GetGraphicsStyle(GraphicsStyleType.Projection)
    except Exception:
        pass
    return None


def _solid_fill_pattern():
    for fp in revit.query.get_elements_by_class(FillPatternElement, doc=doc):
        try:
            if fp.GetFillPattern().IsSolidFill:
                return fp
        except Exception:
            pass
    return None


def _get_or_create_text_style(name, font, size_mm, bold=False, italic=False):
    """Return an existing TextNoteType by name, or duplicate one and set it up.

    Must run inside a transaction.
    """
    all_types = list(FilteredElementCollector(doc).OfClass(TextNoteType).ToElements())
    if not all_types:
        return None
    existing = None
    for tt in all_types:
        try:
            if tt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == name:
                existing = tt
                break
        except Exception:
            pass
    tt = existing if existing else all_types[0].Duplicate(name)
    for bip, val in [
        (DB.BuiltInParameter.TEXT_SIZE, size_mm * MM),
        (DB.BuiltInParameter.TEXT_FONT, font),
        (DB.BuiltInParameter.TEXT_STYLE_BOLD, 1 if bold else 0),
        (DB.BuiltInParameter.TEXT_STYLE_ITALIC, 1 if italic else 0),
        (DB.BuiltInParameter.TEXT_BACKGROUND, 1),
        (DB.BuiltInParameter.LEADER_OFFSET_SHEET, MM),
        (DB.BuiltInParameter.TEXT_TAB_SIZE, MM),
        (DB.BuiltInParameter.TEXT_WIDTH_SCALE, 0.8),
    ]:
        try:
            p = tt.get_Parameter(bip)
            if p and not p.IsReadOnly:
                p.Set(val)
        except Exception:
            pass
    try:
        ap = tt.get_Parameter(DB.BuiltInParameter.LEADER_ARROWHEAD)
        if ap and not ap.IsReadOnly:
            ap.Set(DB.ElementId.InvalidElementId)
    except Exception:
        pass
    return tt


def _get_or_create_fill_type(name, rgb, solid_pattern):
    """Return a solid FilledRegionType in the given colour, drawn with an
    invisible boundary so the fill never adds a line of its own.

    Must run inside a transaction.
    """
    if not solid_pattern:
        return None

    def _apply(frt):
        col = Color(rgb[0], rgb[1], rgb[2])
        try:
            frt.ForegroundPatternId = solid_pattern.Id
            frt.ForegroundPatternColor = col
            frt.BackgroundPatternColor = col
        except Exception:
            pass
        for ls_name in ('<Invisible lines>', '<Invisible Lines>'):
            gs = _get_line_style(ls_name)
            if gs:
                try:
                    frt.LineStyleId = gs.Id
                    return
                except Exception:
                    pass

    for frt in FilteredElementCollector(doc).OfClass(FilledRegionType).ToElements():
        try:
            if frt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == name:
                _apply(frt)
                return frt
        except Exception:
            pass
    all_frt = list(FilteredElementCollector(doc).OfClass(FilledRegionType).ToElements())
    if not all_frt:
        return None
    new_frt = all_frt[0].Duplicate(name)
    _apply(new_frt)
    return new_frt


_wanted_fonts = set()
_wanted_fills = set()
for _pl in PLACEMENTS:
    if str(_pl['text'] or '').strip():
        _wanted_fonts.add(_font_key(_pl['block']))
    _hexc = _fill_hex(_pl)
    if _hexc:
        _wanted_fills.add(_hexc.upper())

TEXT_TYPES = {}
FILL_TYPES = {}
with revit.Transaction('pyTransmit Studio DV - Types') as _t_types:
    for _fk in sorted(_wanted_fonts):
        # The name carries the same rounding the cache key uses, so two sizes
        # that round together share one type rather than colliding on a
        # Duplicate() that Revit would refuse.
        _name = 'pyT Studio {} {:.3f}{}{}'.format(
            _fk[0], _fk[1], 'B' if _fk[2] else '', 'I' if _fk[3] else '')
        try:
            _created = _get_or_create_text_style(
                _name, _fk[0], _fk[1], bold=_fk[2], italic=_fk[3])
            if _created is not None:
                TEXT_TYPES[_fk] = _created
        except Exception as _e:
            _log('Text style "{}" could not be created: {}'.format(_name, _e))
    _solid = _solid_fill_pattern()
    for _hexc in sorted(_wanted_fills):
        try:
            _created = _get_or_create_fill_type(
                'pyT Studio {}'.format(_hexc), _hex_to_rgb(_hexc), _solid)
            if _created is not None:
                FILL_TYPES[_hexc] = _created
        except Exception as _e:
            _log('Fill "{}" could not be created: {}'.format(_hexc, _e))

# Only a layout with text needs a text style, so an empty set here means the
# model has no TextNoteType to duplicate - which no amount of drawing fixes.
if _wanted_fonts and not TEXT_TYPES:
    _alert('No text style could be created - this model has no TextNoteType '
           'to copy from.', title=TITLE, exitscript=True)

_FALLBACK_TT = TEXT_TYPES[sorted(TEXT_TYPES.keys())[0]] if TEXT_TYPES else None


def _tt(block):
    return TEXT_TYPES.get(_font_key(block)) or _FALLBACK_TT


# ── The drafting view ────────────────────────────────────────────────────────
VIEW_NAME = _p.get('_legend_temp_view_name') or 'pyTransmit Document View'

drafting_view = None
for _v in revit.query.get_elements_by_class(ViewDrafting, doc=doc):
    try:
        if _v.IsValidObject and _v.Name == VIEW_NAME:
            drafting_view = _v
            break
    except Exception:
        pass

if not drafting_view:
    _vft = None
    for vft in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if vft.ViewFamily == ViewFamily.Drafting:
            _vft = vft
            break
    if not _vft:
        _alert('No Drafting ViewFamilyType found.', title=TITLE, exitscript=True)
    with revit.Transaction('pyTransmit Studio DV - Create view') as _t:
        drafting_view = ViewDrafting.Create(doc, _vft.Id)
        drafting_view.Name = VIEW_NAME
        try:
            drafting_view.Scale = 1
        except Exception:
            pass
else:
    with revit.Transaction('pyTransmit Studio DV - Set scale') as _t:
        try:
            drafting_view.Scale = 1
        except Exception:
            pass

with revit.Transaction('pyTransmit Studio DV - Clear view') as _t:
    for _cls in (CurveElement, TextNote, ImageInstance, FilledRegion):
        for _el in list(FilteredElementCollector(
                doc, drafting_view.Id).OfClass(_cls).ToElements()):
            try:
                doc.Delete(_el.Id)
            except Exception:
                pass

# ── Draw helpers ─────────────────────────────────────────────────────────────
_drawn_lines = set()


def _line(vw, x1, y1, x2, y2):
    """A detail line, drawn once.

    Neighbouring cells share edges, so the same segment comes up twice for
    every internal border in the grid - a 2000-sheet transmittal would draw
    tens of thousands of duplicate lines on top of each other.
    """
    key = (round(x1, 6), round(y1, 6), round(x2, 6), round(y2, 6))
    if key in _drawn_lines:
        return None
    _drawn_lines.add(key)
    try:
        start = XYZ(float(x1), float(y1), 0.0)
        end = XYZ(float(x2), float(y2), 0.0)
        if (end - start).GetLength() < SHORT_CURVE_TOL:
            return None
        return doc.Create.NewDetailCurve(vw, Line.CreateBound(start, end))
    except Exception:
        return None


def _fill(vw, frt, x1, y1, x2, y2):
    try:
        loop = CurveLoop()
        pts = [XYZ(float(x1), float(y1), 0.0), XYZ(float(x2), float(y1), 0.0),
               XYZ(float(x2), float(y2), 0.0), XYZ(float(x1), float(y2), 0.0)]
        for i in range(4):
            loop.Append(Line.CreateBound(pts[i], pts[(i + 1) % 4]))
        region = FilledRegion.Create(doc, frt.Id, vw.Id, [loop])
        if region:
            for ls_id in FilledRegion.GetValidLineStyleIdsForFilledRegion(doc):
                el = doc.GetElement(ls_id)
                name = ''
                if el is not None:
                    name = el.Name if hasattr(el, 'Name') else ''
                if name and 'invisible' in name.lower():
                    region.SetLineStyleId(ls_id)
                    break
        return region
    except Exception as e:
        _log('Fill could not be drawn: {}'.format(e))
        return None


_H_ALIGN = {'left': HorizontalTextAlignment.Left,
            'center': HorizontalTextAlignment.Center,
            'right': HorizontalTextAlignment.Right}


def _text(vw, block, text, x1, y1, x2, y2):
    """Place one cell's text inside the rectangle (x1, y1) - (x2, y2), in feet.

    Revit anchors a text note at the TOP of its box, so vertical placement is
    arithmetic on the text's own height rather than a property to set.
    """
    text = u'{}'.format(text or '')
    if not text.strip():
        return None
    b = block or {}
    size_mm = float(b.get('size_mm') or _DEFAULT_SIZE_MM)
    size_ft = size_mm * MM
    just = b.get('just', 'left')
    v_just = b.get('v_just', 'middle')
    rotation = int(b.get('rotation') or 0)
    tt = _tt(b)
    if tt is None:
        return None

    if rotation in (90, 270):
        # Sideways text. Revit rotates anticlockwise about the insertion
        # point, so both of Studio's rotations read bottom-to-top; the
        # difference between them is one Excel keeps and Revit does not.
        # Width is left unset - the text runs along the column, and its length
        # is the row height, which the layout already sized for it.
        try:
            opts = TextNoteOptions(tt.Id)
            opts.HorizontalAlignment = HorizontalTextAlignment.Left
            opts.Rotation = math.pi / 2.0
            x_ctr = (x1 + x2) / 2.0 - size_ft / 2.0
            y_ctr = (y1 + y2) / 2.0 - len(text) * size_ft * 0.65 / 2.0
            return TextNote.Create(doc, vw.Id, XYZ(float(x_ctr), float(y_ctr), 0.0),
                                   text, opts)
        except Exception as e:
            _log('Rotated text "{}" failed: {}'.format(text[:30], e))
            return None

    width = max(2.0 * MM, (x2 - x1) - INDENT * 2)
    n_lines = text.count('\n') + 1
    block_h = n_lines * size_ft * 1.2
    if v_just == 'top':
        y = y1 - INDENT
    elif v_just == 'bottom':
        y = y2 + block_h
    else:
        y = (y1 + y2) / 2.0 + block_h / 2.0
    try:
        # Always created Left-aligned so Revit anchors the box at x1, then
        # realigned: passing Center/Right to Create() moves the box instead.
        opts = TextNoteOptions(tt.Id)
        opts.HorizontalAlignment = HorizontalTextAlignment.Left
        note = TextNote.Create(doc, vw.Id, XYZ(float(x1 + INDENT), float(y), 0.0),
                               float(width), text, opts)
        if note is not None and just in ('center', 'right'):
            try:
                note.HorizontalAlignment = _H_ALIGN[just]
            except Exception:
                pass
        return note
    except Exception as e:
        _log('Text "{}" failed: {}'.format(text[:30], e))
        return None


def _get_or_create_logo_type():
    """The layout's logo as an ImageType, imported once and reused."""
    path = SL.logo_path
    if not path:
        return None
    if not os.path.isfile(path):
        _log('Logo file is missing, skipped: {}'.format(path))
        return None

    def _src(img):
        try:
            return (img.SourcePath or '').strip()
        except Exception:
            return ''

    norm = os.path.normcase(path)
    for img in FilteredElementCollector(doc).OfClass(ImageType).ToElements():
        try:
            if os.path.normcase(_src(img)) == norm:
                return img
        except Exception:
            pass
    try:
        opts = ImageTypeOptions(path, False, ImageTypeSource.Link)
        opts.Resolution = 300
        return ImageType.Create(doc, opts)
    except Exception as e:
        _log('Logo could not be imported: {}'.format(e))
        return None


def _place_logo(vw, logo_type, block, x1, y1, x2, y2):
    """Scale the logo to fit its cell and place it by the block's alignment."""
    if logo_type is None:
        return
    b = block or {}
    just = b.get('just', 'left')
    v_just = b.get('v_just', 'middle')
    box_w = max(MM, (x2 - x1) - INDENT * 2)
    box_h = max(MM, (y1 - y2) - INDENT * 2)
    try:
        img_w = float(logo_type.Width)
        img_h = float(logo_type.Height)
        scale = 1.0
        if img_w > 0 and img_h > 0:
            # Fit inside the cell either way round; aspect ratio preserved, so
            # a logo always looks like itself.
            scale = min(box_w / img_w, box_h / img_h)
        draw_w, draw_h = img_w * scale, img_h * scale

        x = {'center': (x1 + x2) / 2.0 - draw_w / 2.0,
             'right': x2 - INDENT - draw_w}.get(just, x1 + INDENT)
        y = {'middle': (y1 + y2) / 2.0 + draw_h / 2.0,
             'bottom': y2 + INDENT + draw_h}.get(v_just, y1 - INDENT)

        opts = ImagePlacementOptions()
        opts.Location = XYZ(float(x), float(y), 0.0)
        inst = ImageInstance.Create(doc, vw, logo_type.Id, opts)
        if inst is None:
            return
        wp = inst.LookupParameter('Width')
        hp = inst.LookupParameter('Height')
        if wp and not wp.IsReadOnly:
            wp.Set(draw_w)
        if hp and not hp.IsReadOnly:
            hp.Set(draw_h)
        inst.SetLocation(XYZ(float(x), float(y), 0.0), BoxPlacement.TopLeft)
    except Exception as e:
        _log('Logo could not be placed: {}'.format(e))


# ── Draw ─────────────────────────────────────────────────────────────────────
_logo_block_seen = False

with revit.Transaction('pyTransmit Studio DV - Draw transmittal') as _t_draw:

    LOGO_TYPE = _get_or_create_logo_type()

    for _page in PAGES:
        _rows = _page['rows']
        _x_off = _page['x_mm'] * MM

        for _pl in PLACEMENTS:
            # A placement can start above this column and end inside it - a
            # static cell beside the sheet list spans every row of the table -
            # so it is drawn against whichever of its rows this column holds.
            _vs = [v for v in range(_pl['v'], _pl['v'] + _pl['n_v']) if v in _rows]
            if not _vs:
                continue
            _block = _pl['block']
            _y_top = _rows[_vs[0]][0] * MM
            _y_bot = _rows[_vs[-1]][1] * MM
            _x1 = _x_off + SL.col_x(_pl['c']) * MM
            _x2 = _x1 + SL.span_w(_pl['c'], _pl['col_span']) * MM

            _hexc = _fill_hex(_pl)
            _frt = FILL_TYPES.get(_hexc.upper()) if _hexc else None
            if _frt is not None:
                _fill(drafting_view, _frt, _x1, _y_top, _x2, _y_bot)

            if (_block or {}).get('type') == 'logo':
                if SL.logo_path:
                    _place_logo(drafting_view, LOGO_TYPE, _block,
                                _x1, _y_top, _x2, _y_bot)
                    _logo_block_seen = True
                else:
                    _log('Logo block at row {} col {} has no logo assigned - '
                         'pick one in Studio (Layout tab > Logo).'.format(
                             _pl['r'] + 1, _pl['c'] + 1))
            else:
                _text(drafting_view, _block, _pl['text'], _x1, _y_top, _x2, _y_bot)

            # Group headers can rule themselves separately from the data rows,
            # and a deliberate gap carries no rules at all - both decided by
            # the row kind, in the one place all three writers ask.
            _borders = studio_rows.borders_for(_block, _pl['kind'])
            if _borders.get('t'):
                _line(drafting_view, _x1, _y_top, _x2, _y_top)
            if _borders.get('b'):
                _line(drafting_view, _x1, _y_bot, _x2, _y_bot)
            if _borders.get('l'):
                _line(drafting_view, _x1, _y_top, _x1, _y_bot)
            if _borders.get('r'):
                _line(drafting_view, _x2, _y_top, _x2, _y_bot)

if SL.logo_path and not _logo_block_seen:
    _log('A logo is assigned ({}) but the layout has no Logo block to put it '
         'in - drag one onto a cell in Studio.'.format(
             os.path.basename(SL.logo_path)))

_log('Drafting view "{}" updated: {} rows, {} column(s).'.format(
    VIEW_NAME, len(SL.vrows), len(PAGES)))
