# -*- coding: utf-8 -*-
"""Wrap a U3D stream in a one-page PDF with a 3D annotation.

The page carries a title line and the 3D annotation filling the rest. The
annotation has preset views, opens with the model tree showing, and a
poster for viewers without 3D support saying how to see the model.
Runs under CPython 3 and IronPython 2.7.
"""
import math
import zlib

__all__ = ["View", "standard_views", "write_pdf"]

# ── CONSTANTS ───────────────────────────────────────────────────────────────

PAGE_W = 1190.55        # A3 landscape, points
PAGE_H = 841.89
MARGIN = 24.0
TITLE_H = 28.0
FOV_DEG = 30.0

POSTER_LINES = [
    u"3D model",
    u"Open this PDF in Adobe Acrobat or Acrobat Reader and click here.",
    u"If nothing happens, turn on 3D content: Edit > Preferences >",
    u"3D & Multimedia > Enable playing of 3D content.",
]


# ── VIEWS ───────────────────────────────────────────────────────────────────

def _norm(v):
    l = math.sqrt(v[0] * v[0] + v[1] * v[1] + v[2] * v[2])
    return (v[0] / l, v[1] / l, v[2] / l)


def _cross(a, b):
    return (a[1] * b[2] - a[2] * b[1],
            a[2] * b[0] - a[0] * b[2],
            a[0] * b[1] - a[1] * b[0])


class View(object):
    """A named camera looking along forward with up roughly up."""

    def __init__(self, name, forward, up, perspective=True):
        self.name = name
        self.forward = _norm(forward)
        self.up = _norm(up)
        self.perspective = perspective

    def c2w(self, center, radius):
        """Camera-to-world matrix and orbit distance framing the sphere.

        Columns are the camera's left, up and forward vectors, then its
        position: the layout media9 has long written for Acrobat.
        """
        f = self.forward
        right = _cross(f, self.up)
        if abs(right[0]) + abs(right[1]) + abs(right[2]) < 1e-9:
            right = _cross(f, (0.0, 1.0, 0.0))
        right = _norm(right)
        up = _norm(_cross(right, f))
        left = (-right[0], -right[1], -right[2])
        dist = radius / math.sin(math.radians(FOV_DEG / 2.0)) * 1.05
        pos = (center[0] - f[0] * dist,
               center[1] - f[1] * dist,
               center[2] - f[2] * dist)
        return list(left) + list(up) + list(f) + list(pos), dist


def standard_views():
    """Preset views, Revit world axes (Z up, Y north)."""
    z = (0.0, 0.0, 1.0)
    return [
        View(u"Isometric", (-1.0, 1.0, -1.0), z),
        View(u"Top", (0.0, 0.0, -1.0), (0.0, 1.0, 0.0)),
        View(u"Front", (0.0, 1.0, 0.0), z),
        View(u"Back", (0.0, -1.0, 0.0), z),
        View(u"Left", (1.0, 0.0, 0.0), z),
        View(u"Right", (-1.0, 0.0, 0.0), z),
    ]


# ── PDF ─────────────────────────────────────────────────────────────────────

def _text_hex(s):
    """PDF text string as UTF-16BE hex, safe for any characters."""
    data = bytearray(b"\xfe\xff") + bytearray(s.encode("utf-16-be"))
    return "<" + "".join("{:02X}".format(b) for b in data) + ">"


def _ascii(s):
    """Plain text for page content in a standard font."""
    out = []
    for ch in s:
        if ch in u"()\\":
            out.append(u"\\" + ch)
        elif 32 <= ord(ch) < 127:
            out.append(ch)
        else:
            out.append(u"?")
    return u"".join(out)


def _num(v):
    s = "{:.6f}".format(float(v)).rstrip("0").rstrip(".")
    return s if s not in ("", "-0") else "0"


class _Pdf(object):
    def __init__(self):
        self.objs = []

    def add(self, body=None):
        self.objs.append(body)
        return len(self.objs)

    def set(self, num, body):
        self.objs[num - 1] = body

    @staticmethod
    def stream(dict_body, data):
        if not isinstance(data, (bytes, bytearray)):
            data = data.encode("latin-1")
        head = u"<< {} /Length {} >>\nstream\n".format(dict_body, len(data))
        return head.encode("latin-1") + bytes(data) + b"\nendstream"

    def to_bytes(self, root):
        out = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
        offsets = []
        for i, body in enumerate(self.objs):
            offsets.append(len(out))
            if not isinstance(body, (bytes, bytearray)):
                body = body.encode("latin-1")
            out += u"{} 0 obj\n".format(i + 1).encode("latin-1")
            out += body
            out += b"\nendobj\n"
        xref = len(out)
        lines = [u"xref", u"0 {}".format(len(self.objs) + 1),
                 u"0000000000 65535 f "]
        for off in offsets:
            lines.append(u"{:010d} 00000 n ".format(off))
        lines.append(u"trailer")
        lines.append(u"<< /Size {} /Root {} 0 R >>".format(
            len(self.objs) + 1, root))
        lines.append(u"startxref")
        lines.append(u"{}".format(xref))
        lines.append(u"%%EOF\n")
        out += u"\n".join(lines).encode("latin-1")
        return bytes(out)


def write_pdf(path, u3d_bytes, views, center, radius, title=u"",
              subtitle=u""):
    """Write the PDF. views[0] is the view it opens on."""
    pdf = _Pdf()
    catalog = pdf.add()
    pages = pdf.add()
    page = pdf.add()
    annot = pdf.add()
    font = pdf.add(u"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica "
                   u"/Encoding /WinAnsiEncoding >>")

    # Views.
    view_refs = []
    for v in views:
        c2w, dist = v.c2w(center, radius)
        proj = (u"<< /Subtype /P /FOV {} /PS /Min >>".format(_num(FOV_DEG))
                if v.perspective else
                u"<< /Subtype /O /OB /Min >>")
        body = (u"<< /Type /3DView /XN {name} /IN {name} /MS /M "
                u"/C2W [{c2w}] /CO {co} /P {proj} "
                u"/BG << /Type /3DBG /Subtype /SC /CS /DeviceRGB "
                u"/C [1 1 1] >> "
                u"/LS << /Type /3DLightingScheme /Subtype /CAD >> "
                u"/RM << /Type /3DRenderMode /Subtype /SolidOutline >> "
                u">>").format(name=_text_hex(v.name),
                              c2w=u" ".join(_num(x) for x in c2w),
                              co=_num(dist), proj=proj)
        view_refs.append(pdf.add(body))

    model = pdf.add(pdf.stream(
        u"/Type /3D /Subtype /U3D /DV 0 /VA [{}] /Filter /FlateDecode".format(
            u" ".join(u"{} 0 R".format(r) for r in view_refs)),
        zlib.compress(bytes(u3d_bytes), 9)))

    # Annotation area and poster.
    x0, y0 = MARGIN, MARGIN
    x1, y1 = PAGE_W - MARGIN, PAGE_H - MARGIN - TITLE_H
    w, h = x1 - x0, y1 - y0
    poster = [u"q 0.94 0.94 0.94 rg 0 0 {} {} re f Q".format(_num(w), _num(h)),
              u"BT 0.25 0.25 0.25 rg"]
    size = 22
    y = h / 2.0 + 40
    for i, line in enumerate(POSTER_LINES):
        poster.append(u"/F1 {} Tf 1 0 0 1 {} {} Tm ({}) Tj".format(
            size, _num(40), _num(y), _ascii(line)))
        y -= size + 10
        size = 14
    poster.append(u"ET")
    ap = pdf.add(pdf.stream(
        u"/Type /XObject /Subtype /Form /BBox [0 0 {} {}] "
        u"/Resources << /Font << /F1 {} 0 R >> >>".format(
            _num(w), _num(h), font),
        u"\n".join(poster)))

    pdf.set(annot, (
        u"<< /Type /Annot /Subtype /3D /Rect [{} {} {} {}] /F 4 "
        u"/NM {} /Contents {} /3DD {} 0 R /3DV 0 "
        u"/3DA << /A /PO /D /PC /TB true /NP true >> /3DI true "
        u"/AP << /N {} 0 R >> /P {} 0 R >>").format(
            _num(x0), _num(y0), _num(x1), _num(y1),
            _text_hex(u"3D Model"), _text_hex(title or u"3D model"),
            model, ap, page))

    # Page content: title line above the model.
    ty = PAGE_H - MARGIN - 18
    content = [u"BT 0 0 0 rg /F1 14 Tf 1 0 0 1 {} {} Tm ({}) Tj ET".format(
        _num(MARGIN), _num(ty), _ascii(title))]
    if subtitle:
        content.append(
            u"BT 0.4 0.4 0.4 rg /F1 9 Tf 1 0 0 1 {} {} Tm ({}) Tj ET".format(
                _num(x1 - len(subtitle) * 9 * 0.52), _num(ty),
                _ascii(subtitle)))
    contents = pdf.add(pdf.stream(u"", u"\n".join(content)))

    pdf.set(page, (
        u"<< /Type /Page /Parent {} 0 R /MediaBox [0 0 {} {}] "
        u"/Resources << /Font << /F1 {} 0 R >> >> /Contents {} 0 R "
        u"/Annots [{} 0 R] >>").format(
            pages, _num(PAGE_W), _num(PAGE_H), font, contents, annot))
    pdf.set(pages, u"<< /Type /Pages /Kids [{} 0 R] /Count 1 >>".format(page))
    pdf.set(catalog, u"<< /Type /Catalog /Pages {} 0 R >>".format(pages))

    data = pdf.to_bytes(catalog)
    with open(path, "wb") as f:
        f.write(data)
    return len(data)
