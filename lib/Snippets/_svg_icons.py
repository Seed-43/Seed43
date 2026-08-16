# -*- coding: utf-8 -*-
"""Render the extension's icon.png files from SVG sources in lib/icons.

Why this exists: every button icon is a baked PNG, so changing the accent
colour in About recoloured every window but left the toolbar looking like
the old theme. This regenerates the PNGs from SVG, swapping the two brand
greens for whatever the palette's accent is now.

Layout:
    lib/icons/<Name>.svg  ->  the folder whose name (minus .pushbutton /
                              .pulldown / ...) is <Name>, written as icon.png.

The extension's own icon.png at the root deliberately has no SVG: it stays
the fixed green branding while the buttons follow the accent.

Only two exact hex values are treated as brand colours (see ACCENT_SOURCES).
Every other colour is left alone - pyLink's green document (#28b74a) is a
deliberate part of that icon and must not follow the accent.

The module is deliberately split:

  * Parsing (parse_svg, and everything it calls) is plain Python with no
    .NET imports, so it can be tested outside Revit.
  * Rendering (render_png) is WPF, and only runs under IronPython.

This is a targeted reader for Inkscape's output, not a general SVG engine.
It covers what the Seed43 icons actually use; anything else is skipped with
a note rather than guessed at.
"""

import os
import re
import sys
import os.path as op
import xml.etree.ElementTree as ET

__all__ = [
    'ACCENT_SOURCES', 'accent_mapping', 'substitute_colours',
    'parse_svg', 'render_png', 'icon_targets', 'rebuild_all',
]

SVG_NS = '{http://www.w3.org/2000/svg}'

# The two brand greens as drawn, mapped to the palette keys they represent.
# Base green is the body of each icon; the lighter one is the top highlight.
ACCENT_SOURCES = [
    ('#1f893c', 'accent'),
    ('#1faa3c', 'accent_hover'),
]

# Folder suffixes a bundle name can carry, stripped when matching an SVG.
BUNDLE_SUFFIXES = ('.pushbutton', '.pulldown', '.stack', '.panel', '.tab')

DEFAULT_LONG_EDGE = 96      # px on the longer side; the rest follows the aspect


# ── COLOUR SUBSTITUTION ──────────────────────────────────────────────────────

def accent_mapping(palette, profile='dark'):
    """{source_hex: replacement_hex} for the current palette.

    Returns an empty dict if the palette can't supply a colour, which leaves
    the SVG untouched rather than rendering it some default green."""
    out = {}
    colours = (palette or {}).get('profiles', {}).get(profile, {})
    for source_hex, key in ACCENT_SOURCES:
        value = colours.get(key)
        if value:
            out[source_hex.lower()] = value
    return out


def substitute_colours(svg_text, mapping):
    """Swap the brand greens for the accent, case-insensitively.

    Done on the raw text before parsing: the greens appear in `style=`,
    in `fill=`/`stroke=` attributes and inside gradient stops, and a text
    swap catches all of them without having to touch each one in the tree."""
    if not mapping:
        return svg_text
    for source_hex, replacement in mapping.items():
        svg_text = re.sub(re.escape(source_hex), replacement,
                          svg_text, flags=re.IGNORECASE)
    return svg_text


# ── STYLE ────────────────────────────────────────────────────────────────────

def _style_dict(element):
    """Presentation attributes merged with `style=`, style winning.

    Inkscape writes both, and they disagree: Seed43.svg has
    fill="#e8e020" alongside style="fill:#000000". CSS says the style
    property wins, and getting this backwards turns that icon yellow."""
    out = {}
    for key in ('fill', 'stroke', 'stroke-width', 'fill-opacity',
                'stroke-opacity', 'stroke-linecap', 'stroke-linejoin',
                'display', 'opacity'):
        value = element.get(key)
        if value is not None:
            out[key] = value.strip()
    for declaration in (element.get('style') or '').split(';'):
        if ':' not in declaration:
            continue
        key, _, value = declaration.partition(':')
        out[key.strip()] = value.strip()
    return out


def _inherit(parent_style, own_style):
    """Child style over inherited, the way SVG cascades paint properties."""
    merged = dict(parent_style or {})
    merged.update(own_style or {})
    return merged


def _number(value, default=0.0):
    """A float from an SVG attribute, falling back on anything unusable.

    None is checked before float() rather than caught after: an absent
    attribute is the normal case (a rect with no rx, a gradient with no x1),
    and under IronPython float(None) does not raise the TypeError that
    catching only TypeError/ValueError would expect. Hence the bare except -
    a malformed number must never take an icon down."""
    if value is None:
        return default
    try:
        return float(value)
    except Exception:
        return default


def _paint(style, key):
    """Resolve a fill/stroke to a hex string, 'none', or a url(#id) reference."""
    value = (style.get(key) or '').strip()
    if not value or value == 'none':
        return None
    if value.startswith('url(') and value.endswith(')'):
        return value[4:-1].strip().lstrip('#')
    return value


def _opacity(style, key):
    value = style.get(key)
    if value is None:
        return 1.0
    return max(0.0, min(1.0, _number(value, 1.0)))


# ── TRANSFORMS ───────────────────────────────────────────────────────────────
# Matrices are (a, b, c, d, e, f), the same order SVG and WPF both use.

IDENTITY = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)

_TRANSFORM_RE = re.compile(r'(matrix|translate|scale|rotate)\s*\(([^)]*)\)')


def _multiply(m1, m2):
    """m1 then m2, i.e. the child's own transform applied inside the parent's."""
    a1, b1, c1, d1, e1, f1 = m1
    a2, b2, c2, d2, e2, f2 = m2
    return (
        a1 * a2 + b1 * c2,
        a1 * b2 + b1 * d2,
        c1 * a2 + d1 * c2,
        c1 * b2 + d1 * d2,
        e1 * a2 + f1 * c2 + e2,
        e1 * b2 + f1 * d2 + f2,
    )


def parse_transform(text):
    """An SVG transform list as a single matrix tuple.

    Handles the forms Inkscape emits: matrix, translate, scale and rotate
    (pyTable rotates two rules by -90)."""
    matrix = IDENTITY
    for name, raw_args in _TRANSFORM_RE.findall(text or ''):
        args = [_number(a) for a in re.split(r'[\s,]+', raw_args.strip()) if a]
        if name == 'matrix' and len(args) == 6:
            step = tuple(args)
        elif name == 'translate' and args:
            step = (1.0, 0.0, 0.0, 1.0, args[0], args[1] if len(args) > 1 else 0.0)
        elif name == 'scale' and args:
            sx = args[0]
            sy = args[1] if len(args) > 1 else sx
            step = (sx, 0.0, 0.0, sy, 0.0, 0.0)
        elif name == 'rotate' and args:
            import math
            angle = math.radians(args[0])
            cos, sin = math.cos(angle), math.sin(angle)
            step = (cos, sin, -sin, cos, 0.0, 0.0)
            if len(args) >= 3:      # rotate about a point
                cx, cy = args[1], args[2]
                step = _multiply(_multiply((1, 0, 0, 1, -cx, -cy), step),
                                 (1, 0, 0, 1, cx, cy))
        else:
            continue
        matrix = _multiply(step, matrix)
    return matrix


# ── GRADIENTS ────────────────────────────────────────────────────────────────

def _parse_gradients(root):
    """{id: gradient dict} for every linearGradient, stops resolved.

    Inkscape splits a gradient in two: one element holding the stops, and
    another holding the geometry that points at it via xlink:href. Both are
    followed here so the caller gets one complete gradient."""
    raw = {}
    for element in root.iter(SVG_NS + 'linearGradient'):
        gradient_id = element.get('id')
        if not gradient_id:
            continue
        stops = []
        for stop in element.findall(SVG_NS + 'stop'):
            style = _style_dict(stop)
            stops.append({
                'offset': _number(stop.get('offset'), 0.0),
                'color': style.get('stop-color', '#000000'),
                'opacity': _opacity(style, 'stop-opacity'),
            })
        href = (element.get('{http://www.w3.org/1999/xlink}href')
                or element.get('href') or '').lstrip('#')
        raw[gradient_id] = {
            'stops': stops,
            'href': href or None,
            'x1': _number(element.get('x1'), 0.0),
            'y1': _number(element.get('y1'), 0.0),
            'x2': _number(element.get('x2'), 1.0),
            'y2': _number(element.get('y2'), 0.0),
            'transform': parse_transform(element.get('gradientTransform')),
            'units': element.get('gradientUnits', 'objectBoundingBox'),
        }

    for gradient in raw.values():          # pull stops down the href chain
        seen = set()
        source = gradient
        while not source['stops'] and source['href'] and source['href'] not in seen:
            seen.add(source['href'])
            source = raw.get(source['href'])
            if not source:
                break
            gradient['stops'] = source['stops']
    return raw


# ── SHAPES ───────────────────────────────────────────────────────────────────

def _rect_shape(element):
    """A <rect>, with Inkscape's rounded-corner quirk handled.

    Inkscape writes ry="0" next to a real rx on rounded rectangles and still
    draws them rounded. Read literally, SVG says square corners - which would
    flatten the rounded tile every one of these icons is built on."""
    rx = _number(element.get('rx'), 0.0)
    ry = _number(element.get('ry'), 0.0)
    if rx and not ry:
        ry = rx
    elif ry and not rx:
        rx = ry
    return {
        'kind': 'rect',
        'x': _number(element.get('x')), 'y': _number(element.get('y')),
        'width': _number(element.get('width')),
        'height': _number(element.get('height')),
        'rx': rx, 'ry': ry,
    }


def _points(raw):
    values = [_number(v) for v in re.split(r'[\s,]+', (raw or '').strip()) if v]
    return list(zip(values[0::2], values[1::2]))


def _shape_of(element):
    """Geometry for one element, or None if it isn't a shape this reads."""
    tag = element.tag.replace(SVG_NS, '')
    if tag == 'rect':
        return _rect_shape(element)
    if tag == 'circle':
        return {'kind': 'ellipse',
                'cx': _number(element.get('cx')), 'cy': _number(element.get('cy')),
                'rx': _number(element.get('r')), 'ry': _number(element.get('r'))}
    if tag == 'ellipse':
        return {'kind': 'ellipse',
                'cx': _number(element.get('cx')), 'cy': _number(element.get('cy')),
                'rx': _number(element.get('rx')), 'ry': _number(element.get('ry'))}
    if tag == 'path':
        data = element.get('d')
        return {'kind': 'path', 'data': data} if data else None
    if tag == 'line':
        return {'kind': 'line',
                'x1': _number(element.get('x1')), 'y1': _number(element.get('y1')),
                'x2': _number(element.get('x2')), 'y2': _number(element.get('y2'))}
    if tag in ('polyline', 'polygon'):
        pts = _points(element.get('points'))
        if len(pts) < 2:
            return None
        return {'kind': 'polyline', 'points': pts, 'close': tag == 'polygon'}
    return None


def parse_svg(svg_text):
    """Flatten an SVG into a drawing list, in paint order.

    Returns (viewport, shapes, gradients, skipped) where viewport is
    (width, height) in user units, and each shape carries its resolved
    paint and its fully accumulated transform."""
    root = ET.fromstring(svg_text.encode('utf-8')
                         if isinstance(svg_text, type(u'')) else svg_text)

    width = _number(root.get('width'), 0.0)
    height = _number(root.get('height'), 0.0)
    view_box = [_number(v) for v in
                re.split(r'[\s,]+', (root.get('viewBox') or '').strip()) if v]
    if len(view_box) == 4:
        width = width or view_box[2]
        height = height or view_box[3]

    gradients = _parse_gradients(root)
    shapes, skipped = [], []

    def walk(element, parent_matrix, parent_style):
        for child in element:
            tag = child.tag.replace(SVG_NS, '')
            if tag in ('defs', 'metadata', 'namedview', 'style', 'title', 'desc'):
                continue

            style = _inherit(parent_style, _style_dict(child))
            matrix = _multiply(parse_transform(child.get('transform')),
                               parent_matrix)

            if style.get('display') == 'none':
                continue
            if tag == 'g':
                walk(child, matrix, style)
                continue
            if tag == 'text':
                # Inkscape's flowed text (shape-inside) is SVG 2 and has no
                # WPF equivalent. Flatten it in Inkscape (Path > Object to
                # Path) and it renders like any other shape.
                skipped.append('text')
                continue

            shape = _shape_of(child)
            if shape is None:
                if tag not in ('tspan',):
                    skipped.append(tag)
                continue

            # SVG's initial fill is black, not none. Several glyphs here carry
            # only a stroke-width and rely on that default - read as "no fill"
            # they vanish, leaving a bare coloured tile.
            fill = ('#000000' if style.get('fill') is None
                    else _paint(style, 'fill'))

            shape.update({
                'matrix': matrix,
                'fill': fill,
                'fill_opacity': _opacity(style, 'fill-opacity'),
                'stroke': _paint(style, 'stroke'),
                'stroke_opacity': _opacity(style, 'stroke-opacity'),
                'stroke_width': _number(style.get('stroke-width'), 1.0),
                'linecap': style.get('stroke-linecap', 'butt'),
                'linejoin': style.get('stroke-linejoin', 'miter'),
            })
            shapes.append(shape)

    walk(root, IDENTITY, {})
    return (width, height), shapes, gradients, skipped


# ── TARGETS ──────────────────────────────────────────────────────────────────

def icon_targets(extension_dir):
    """{svg_path: [icon.png path, ...]} for every SVG matching a bundle folder.

    One rule, no exceptions: a name maps to the bundle folder that matches it
    with the suffix stripped. The extension's own icon.png at the root has no
    SVG on purpose - it stays the fixed green branding rather than following
    the accent."""
    icons_dir = op.join(extension_dir, 'lib', 'icons')
    if not op.isdir(icons_dir):
        return {}, []

    folders = {}
    for current, dirnames, _files in os.walk(extension_dir):
        # Prune dot-directories in place. .claude/worktrees holds whole copies
        # of this extension, and walking into them means a name can resolve to
        # a worktree instead of the live tool - icons written somewhere Revit
        # never reads, with nothing to show that anything went wrong.
        dirnames[:] = [d for d in dirnames if not d.startswith('.')]
        for dirname in dirnames:
            for suffix in BUNDLE_SUFFIXES:
                if dirname.lower().endswith(suffix):
                    key = dirname[:-len(suffix)].lower()
                    folders.setdefault(key, []).append(op.join(current, dirname))
                    break

    targets, unmatched = {}, []
    for filename in sorted(os.listdir(icons_dir)):
        if not filename.lower().endswith('.svg'):
            continue
        name = op.splitext(filename)[0]
        source = op.join(icons_dir, filename)
        matches = folders.get(name.lower()) or []
        if not matches:
            unmatched.append(filename)
            continue
        if len(matches) > 1:
            # Prefer one that already has an icon.png, then the shallowest -
            # but say so, because a genuine duplicate name is a mistake worth
            # seeing rather than resolving by luck.
            matches = sorted(
                matches,
                key=lambda p: (not op.isfile(op.join(p, 'icon.png')),
                               p.count(os.sep), p))
            unmatched.append('{} matches {} folders, using {}'.format(
                filename, len(matches), op.basename(op.dirname(matches[0]))))
        targets[source] = [op.join(matches[0], 'icon.png')]
    return targets, unmatched


# ── RENDERING  (WPF - IronPython only) ───────────────────────────────────────

def _wpf():
    """Import the WPF bits lazily so the parser stays importable anywhere."""
    import clr
    clr.AddReference('PresentationCore')
    clr.AddReference('PresentationFramework')
    clr.AddReference('WindowsBase')
    import System
    from System.Windows import Media, Point, Rect
    from System.Windows.Media import Imaging
    return System, Media, Imaging, Point, Rect


def _colour(Media, text, opacity):
    """A WPF Color from an SVG colour string, or None if unreadable.

    Alpha is rebuilt through Color.FromArgb rather than assigning .A, because
    ConvertFromString hands back a boxed struct and mutating that in place
    does not reliably stick."""
    try:
        base = Media.ColorConverter.ConvertFromString(text)
    except Exception:
        return None
    if base is None:
        return None
    alpha = int(max(0.0, min(1.0, opacity)) * base.A)
    return Media.Color.FromArgb(alpha, base.R, base.G, base.B)


def _brush(Media, paint, opacity, gradients):
    """A WPF brush for a resolved paint value, or None for no paint.

    Called inside the shape's pushed transform, so gradient geometry only
    needs the gradient's own transform on top of that."""
    if not paint:
        return None
    if paint in gradients:
        gradient = gradients[paint]
        stops = Media.GradientStopCollection()
        for stop in (gradient['stops']
                     or [{'color': '#000000', 'offset': 0.0, 'opacity': 1.0}]):
            colour = _colour(Media, stop['color'], stop['opacity'] * opacity)
            if colour is None:
                continue
            stops.Add(Media.GradientStop(colour, stop['offset']))
        if stops.Count == 0:
            return None
        brush = Media.LinearGradientBrush(stops)
        brush.StartPoint = _point(Media, gradient['x1'], gradient['y1'])
        brush.EndPoint = _point(Media, gradient['x2'], gradient['y2'])
        if gradient['units'] == 'userSpaceOnUse':
            brush.MappingMode = Media.BrushMappingMode.Absolute
            brush.Transform = _matrix_transform(Media, gradient['transform'])
        return brush

    colour = _colour(Media, paint, opacity)
    if colour is None:
        return None
    return Media.SolidColorBrush(colour)


def _point(Media, x, y):
    from System.Windows import Point
    return Point(x, y)


def _matrix_transform(Media, matrix):
    a, b, c, d, e, f = matrix
    return Media.MatrixTransform(Media.Matrix(a, b, c, d, e, f))


def _geometry(Media, Rect, shape):
    kind = shape['kind']
    if kind == 'rect':
        rect = Rect(shape['x'], shape['y'], shape['width'], shape['height'])
        return Media.RectangleGeometry(rect, shape['rx'], shape['ry'])
    if kind == 'ellipse':
        return Media.EllipseGeometry(_point(Media, shape['cx'], shape['cy']),
                                     shape['rx'], shape['ry'])
    if kind == 'path':
        return Media.Geometry.Parse(shape['data'])
    if kind == 'line':
        return Media.LineGeometry(_point(Media, shape['x1'], shape['y1']),
                                  _point(Media, shape['x2'], shape['y2']))
    if kind == 'polyline':
        figure = Media.PathFigure()
        figure.StartPoint = _point(Media, *shape['points'][0])
        for x, y in shape['points'][1:]:
            figure.Segments.Add(Media.LineSegment(_point(Media, x, y), True))
        figure.IsClosed = shape['close']
        figure.IsFilled = shape['close']
        geometry = Media.PathGeometry()
        geometry.Figures.Add(figure)
        return geometry
    return None


_CAPS = {'butt': 'Flat', 'round': 'Round', 'square': 'Square'}
_JOINS = {'miter': 'Miter', 'round': 'Round', 'bevel': 'Bevel'}


def render_png(viewport, shapes, gradients, out_path, long_edge=DEFAULT_LONG_EDGE):
    """Draw the parsed shapes and write a PNG. Returns (width, height)."""
    System, Media, Imaging, Point, Rect = _wpf()

    src_w, src_h = viewport
    if not src_w or not src_h:
        raise ValueError('SVG has no usable width/height')
    scale = float(long_edge) / max(src_w, src_h)
    out_w = max(1, int(round(src_w * scale)))
    out_h = max(1, int(round(src_h * scale)))

    visual = Media.DrawingVisual()
    context = visual.RenderOpen()
    try:
        for shape in shapes:
            geometry = _geometry(Media, Rect, shape)
            if geometry is None:
                continue

            fill = _brush(Media, shape['fill'], shape['fill_opacity'], gradients)
            pen = None
            stroke = _brush(Media, shape['stroke'], shape['stroke_opacity'],
                            gradients)
            if stroke is not None and shape['stroke_width'] > 0:
                pen = Media.Pen(stroke, shape['stroke_width'])
                cap = getattr(Media.PenLineCap,
                              _CAPS.get(shape['linecap'], 'Flat'))
                pen.StartLineCap = pen.EndLineCap = cap
                pen.LineJoin = getattr(Media.PenLineJoin,
                                       _JOINS.get(shape['linejoin'], 'Miter'))
            if fill is None and pen is None:
                continue

            # The transform goes on the context, not the geometry. Geometry
            # .Parse hands back a frozen object, so assigning .Transform to it
            # throws - and pushing it here is more correct anyway: strokes and
            # gradient coordinates end up scaled with the shape, as SVG says
            # they should.
            context.PushTransform(_matrix_transform(Media, shape['matrix']))
            try:
                context.DrawGeometry(fill, pen, geometry)
            finally:
                context.Pop()
    finally:
        context.Close()

    visual.Transform = Media.ScaleTransform(scale, scale)
    bitmap = Imaging.RenderTargetBitmap(out_w, out_h, 96, 96,
                                        Media.PixelFormats.Pbgra32)
    bitmap.Render(visual)

    encoder = Imaging.PngBitmapEncoder()
    encoder.Frames.Add(Imaging.BitmapFrame.Create(bitmap))
    folder = op.dirname(out_path)
    if folder and not op.isdir(folder):
        os.makedirs(folder)
    stream = System.IO.FileStream(out_path, System.IO.FileMode.Create)
    try:
        encoder.Save(stream)
    finally:
        stream.Close()
    return out_w, out_h


# ── ENTRY POINT ──────────────────────────────────────────────────────────────

def is_stale(svg_path, png_path, palette_path=None):
    """True if an icon needs redrawing: missing, or older than its source.

    The palette counts as a source - an accent change makes every PNG stale
    without any SVG being touched."""
    if not op.isfile(png_path):
        return True
    try:
        png_time = op.getmtime(png_path)
        if op.getmtime(svg_path) > png_time:
            return True
        if palette_path and op.isfile(palette_path):
            return op.getmtime(palette_path) > png_time
    except Exception:
        return True     # can't tell, so redraw rather than skip silently
    return False


def rebuild_all(extension_dir, palette=None, profile='dark',
                long_edge=DEFAULT_LONG_EDGE, only_stale=False):
    """Regenerate every icon.png from lib/icons. Returns (written, problems).

    only_stale skips icons already newer than their SVG and the palette,
    which is what makes this cheap enough to run on every Revit start: with
    nothing changed it costs a handful of stat calls and renders nothing.

    Never raises: a bad SVG costs that one icon, not the whole run, because
    this fires from a window closing or a startup hook, with nowhere useful
    to report an exception to."""
    import io
    palette_path = op.join(extension_dir, 'lib', 'Snippets',
                           'seed43_palette.json')
    if palette is None:
        try:
            from Snippets import seed43_theme
            palette = seed43_theme.load_palette(extension_dir)
            profile = seed43_theme.get_active_profile(extension_dir)
        except Exception:
            palette = None

    mapping = accent_mapping(palette, profile)
    targets, unmatched = icon_targets(extension_dir)

    written, problems = [], ['no folder matches ' + n for n in unmatched]
    for source, destinations in sorted(targets.items()):
        name = op.basename(source)
        due = [d for d in destinations
               if not only_stale or is_stale(source, d, palette_path)]
        if not due:
            continue
        try:
            with io.open(source, 'r', encoding='utf-8') as handle:
                svg_text = handle.read()
            viewport, shapes, gradients, skipped = parse_svg(
                substitute_colours(svg_text, mapping))
            if not shapes:
                problems.append(name + ': nothing to draw')
                continue
            for destination in due:
                render_png(viewport, shapes, gradients, destination, long_edge)
                written.append(destination)
            if skipped:
                problems.append('{}: skipped {}'.format(
                    name, ', '.join(sorted(set(skipped)))))
        except Exception as ex:
            # The message alone rarely says where a .NET error came from -
            # "Object reference not set" could be any line in the render.
            # The last frame is what makes the next fix a lookup, not a guess.
            import traceback
            frames = traceback.extract_tb(sys.exc_info()[2])
            where = ''
            if frames:
                filename, lineno, func, _text = frames[-1]
                where = ' [{}:{} in {}]'.format(
                    op.basename(filename), lineno, func)
            problems.append('{}: {}{}'.format(name, ex, where))
    return written, problems
