# -*- coding: utf-8 -*-
"""A small IFC (STEP physical file) reader: entities, values, units, placements.

Reads only what the steel importer needs, and parses an entity's arguments
the first time they are asked for, so a 100 MB file is not decoded in full.
No Revit types in here, so it runs under CPython for testing as well as
IronPython 2 inside Revit.

    model = IfcFile.load(path)
    for eid in model.by_type("IFCBEAM"): args = model.args(eid)
"""
import re

__all__ = ["Ref", "IfcFile", "decode_string", "mat_mul", "mat_point",
           "mat_vector", "IDENTITY"]


class Ref(int):
    """A #123 reference, kept apart from plain integers."""
    __slots__ = ()


_TOKEN = re.compile(
    r"'(?:[^']|'')*'"                   # string
    r"|#\d+"                            # reference
    r"|\.[A-Z0-9_]+\."                  # enumeration / boolean
    r"|[A-Z][A-Z0-9_]*\("               # typed value, e.g. IFCLABEL(
    r"|[-+]?[0-9][0-9.]*(?:[Ee][-+]?[0-9]+)?|[-+]?\.[0-9]+(?:[Ee][-+]?[0-9]+)?"
    r"|[()$*]")

_ENTITY = re.compile(r"#(\d+)\s*=\s*([A-Za-z0-9_]+)\s*\(")

_X2 = re.compile(r"\\X2\\((?:[0-9A-Fa-f]{4})+)\\X0\\")
_X4 = re.compile(r"\\X4\\((?:[0-9A-Fa-f]{8})+)\\X0\\")
_X1 = re.compile(r"\\X\\([0-9A-Fa-f]{2})")
_S = re.compile(r"\\S\\(.)")
_P = re.compile(r"\\P[A-Z]\\")

try:
    _CHR = unichr       # IronPython 2
except NameError:
    _CHR = chr


def decode_string(raw):
    """STEP string body (without the quotes) to text."""
    s = raw.replace("''", "'")
    if "\\" not in s:
        return s
    s = _X2.sub(lambda m: u"".join(_CHR(int(m.group(1)[i:i + 4], 16))
                                   for i in range(0, len(m.group(1)), 4)), s)
    s = _X4.sub(lambda m: u"".join(_safe_chr(int(m.group(1)[i:i + 8], 16))
                                   for i in range(0, len(m.group(1)), 8)), s)
    s = _X1.sub(lambda m: _CHR(int(m.group(1), 16)), s)
    s = _S.sub(lambda m: _CHR(ord(m.group(1)) + 128), s)
    s = _P.sub(u"", s)
    return s.replace("\\\\", "\\")


def _safe_chr(code):
    try:
        return _CHR(code)
    except ValueError:
        return u"?"


def parse_args(text):
    """Argument list text (inside the outer brackets) to Python values:
    None for $ and *, Ref for references, str for enums (.T. -> 'T'),
    text for strings, float for numbers, lists, and (TYPE, value) for typed
    values such as IFCLABEL('x')."""
    stack = [[]]
    for m in _TOKEN.finditer(text):
        tok = m.group(0)
        c = tok[0]
        if c == "'":
            stack[-1].append(decode_string(tok[1:-1]))
        elif c == "#":
            stack[-1].append(Ref(int(tok[1:])))
        elif c == "." and not tok[1:2].isdigit():
            stack[-1].append(tok[1:-1])
        elif c == "(":
            stack.append([])
        elif c == ")":
            if len(stack) == 1:
                break
            inner = stack.pop()
            if isinstance(inner, _Typed):
                stack[-1].append((inner.type,
                                  inner[0] if len(inner) == 1 else list(inner)))
            else:
                stack[-1].append(inner)
        elif c in "$*":
            stack[-1].append(None)
        elif tok.endswith("("):
            stack.append(_Typed(tok[:-1]))
        else:
            stack[-1].append(float(tok))
    return stack[0]


class _Typed(list):
    """The argument list of a typed value while it is being read."""

    def __init__(self, type_name):
        list.__init__(self)
        self.type = type_name


# ── 3x4 MATRICES ────────────────────────────────────────────────────────────
# (x axis, y axis, z axis, origin), each a 3-tuple: point = o + x*X + y*Y + z*Z

IDENTITY = ((1.0, 0.0, 0.0), (0.0, 1.0, 0.0), (0.0, 0.0, 1.0), (0.0, 0.0, 0.0))


def mat_vector(m, v):
    x, y, z = v
    return (m[0][0] * x + m[1][0] * y + m[2][0] * z,
            m[0][1] * x + m[1][1] * y + m[2][1] * z,
            m[0][2] * x + m[1][2] * y + m[2][2] * z)


def mat_point(m, p):
    v = mat_vector(m, p)
    return (v[0] + m[3][0], v[1] + m[3][1], v[2] + m[3][2])


def mat_mul(a, b):
    """a after b: the matrix of applying b, then a."""
    return (mat_vector(a, b[0]), mat_vector(a, b[1]), mat_vector(a, b[2]),
            mat_point(a, b[3]))


def _norm(v):
    l = (v[0] * v[0] + v[1] * v[1] + v[2] * v[2]) ** 0.5
    if l < 1e-12:
        return None
    return (v[0] / l, v[1] / l, v[2] / l)


def _cross(a, b):
    return (a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2],
            a[0] * b[1] - a[1] * b[0])


def _dot(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def frame(z_axis, x_ref, origin):
    """Right-handed matrix from an axis and a reference direction (IFC rules)."""
    z = _norm(z_axis) or (0.0, 0.0, 1.0)
    x = x_ref
    if x is not None:
        d = _dot(x, z)
        x = _norm((x[0] - d * z[0], x[1] - d * z[1], x[2] - d * z[2]))
    if x is None:
        x = (1.0, 0.0, 0.0)
        if abs(z[0]) > 0.9:
            x = (0.0, 1.0, 0.0)
        d = _dot(x, z)
        x = _norm((x[0] - d * z[0], x[1] - d * z[1], x[2] - d * z[2]))
    y = _cross(z, x)
    return (x, y, z, origin)


# ── THE FILE ────────────────────────────────────────────────────────────────

class IfcFile(object):
    def __init__(self):
        self.raw = {}           # id -> (TYPE, argument text)
        self.types = {}         # TYPE -> [ids]
        self._args = {}
        self._place = {}
        self.schema = u""
        self.header = u""
        self.mm = 1.0           # file length unit in millimetres

    @classmethod
    def load(cls, path, progress=None):
        f = cls()
        with open(path, "rb") as fh:
            data = fh.read()
        f._read(data, progress)
        return f

    @classmethod
    def from_bytes(cls, data, progress=None):
        f = cls()
        f._read(data, progress)
        return f

    def _read(self, data, progress):
        if not isinstance(data, type(u"")):
            data = data.decode("latin-1")
        head_end = data.find("DATA;")
        self.header = data[:head_end if head_end > 0 else 2000]
        m = re.search(r"FILE_SCHEMA\s*\(\s*\(\s*'([^']*)'", self.header)
        self.schema = m.group(1).upper() if m else u""
        raw = self.raw
        types = self.types
        total = len(data)
        pos = head_end if head_end > 0 else 0
        count = 0
        while True:
            m = _ENTITY.search(data, pos)
            if m is None:
                break
            start = m.end()
            end = _statement_end(data, start)
            if end < 0:
                break
            eid = int(m.group(1))
            t = m.group(2).upper()
            raw[eid] = (t, data[start:end])
            lst = types.get(t)
            if lst is None:
                types[t] = [eid]
            else:
                lst.append(eid)
            pos = end + 1
            count += 1
            if progress is not None and count % 20000 == 0:
                progress(pos, total)
        self.mm = self._length_unit()

    # -- access ---------------------------------------------------------------

    def type_of(self, eid):
        r = self.raw.get(eid)
        return r[0] if r else None

    def args(self, eid):
        a = self._args.get(eid)
        if a is None:
            r = self.raw.get(eid)
            if r is None:
                return None
            a = parse_args(r[1])
            self._args[eid] = a
        return a

    def by_type(self, *names):
        out = []
        for n in names:
            out.extend(self.types.get(n, ()))
        return out

    # -- units ------------------------------------------------------------------

    def _length_unit(self):
        prefixes = {"EXA": 1e21, "PETA": 1e18, "TERA": 1e15, "GIGA": 1e12,
                    "MEGA": 1e9, "KILO": 1e6, "HECTO": 1e5, "DECA": 1e4,
                    "DECI": 100.0, "CENTI": 10.0, "MILLI": 1.0,
                    "MICRO": 1e-3, "NANO": 1e-6}
        for ua in self.by_type("IFCUNITASSIGNMENT"):
            for ref in self.args(ua)[0] or []:
                t = self.type_of(ref)
                a = self.args(ref)
                if t == "IFCSIUNIT" and a[1] == "LENGTHUNIT":
                    return prefixes.get(a[2], 1000.0) if a[2] else 1000.0
                if t == "IFCCONVERSIONBASEDUNIT" and a[1] == "LENGTHUNIT":
                    name = (a[2] or u"").upper()
                    if "FOOT" in name or "FEET" in name:
                        return 304.8
                    if "INCH" in name:
                        return 25.4
                    try:
                        mwu = self.args(a[3])           # IfcMeasureWithUnit
                        value = mwu[0][1] if isinstance(mwu[0], tuple) else mwu[0]
                        base = self.args(mwu[1])
                        scale = prefixes.get(base[2], 1000.0) if base[2] else 1000.0
                        return float(value) * scale
                    except Exception:
                        pass
        return 1.0

    # -- geometry primitives ----------------------------------------------------

    def point(self, ref):
        """IfcCartesianPoint in mm, always 3D."""
        c = self.args(ref)[0]
        s = self.mm
        if len(c) == 2:
            return (c[0] * s, c[1] * s, 0.0)
        return (c[0] * s, c[1] * s, c[2] * s)

    def direction(self, ref):
        if ref is None:
            return None
        c = self.args(ref)[0]
        if len(c) == 2:
            return _norm((c[0], c[1], 0.0))
        return _norm((c[0], c[1], c[2]))

    def axis_placement(self, ref):
        """IfcAxis2Placement2D/3D (or None) to a matrix in mm."""
        if ref is None:
            return IDENTITY
        t = self.type_of(ref)
        a = self.args(ref)
        loc = self.point(a[0])
        if t == "IFCAXIS2PLACEMENT2D":
            return frame((0.0, 0.0, 1.0), self.direction(a[1]), loc)
        if t == "IFCAXIS2PLACEMENT3D":
            return frame(self.direction(a[1]) or (0.0, 0.0, 1.0),
                         self.direction(a[2]), loc)
        return IDENTITY

    def transform_operator(self, ref):
        """IfcCartesianTransformationOperator3D (uniform scale) to a matrix."""
        a = self.args(ref)
        t = self.type_of(ref)
        x = self.direction(a[0]) or (1.0, 0.0, 0.0)
        y = self.direction(a[1]) or (0.0, 1.0, 0.0)
        o = self.point(a[2])
        scale = float(a[3]) if len(a) > 3 and a[3] is not None else 1.0
        z = (0.0, 0.0, 1.0)
        if t.startswith("IFCCARTESIANTRANSFORMATIONOPERATOR3D") and len(a) > 4 and a[4] is not None:
            z = self.direction(a[4])
        else:
            z = _norm(_cross(x, y)) or z
        return (tuple(c * scale for c in x), tuple(c * scale for c in y),
                tuple(c * scale for c in z), o)

    def placement(self, ref):
        """IfcLocalPlacement chain to a world matrix in mm."""
        if ref is None:
            return IDENTITY
        m = self._place.get(ref)
        if m is not None:
            return m
        t = self.type_of(ref)
        if t != "IFCLOCALPLACEMENT":
            return IDENTITY         # grid placements are not used by members
        a = self.args(ref)
        local = self.axis_placement(a[1])
        m = mat_mul(self.placement(a[0]), local) if a[0] is not None else local
        self._place[ref] = m
        return m


def _statement_end(data, start):
    """Index of the ';' ending a statement, skipping quoted strings."""
    i = start
    n = len(data)
    find = data.find
    while i < n:
        semi = find(";", i)
        # only quotes before the next ";" matter; a search to the end of
        # the file for every statement made quote-free files very slow
        quote = find("'", i, semi) if semi >= 0 else -1
        if semi < 0:
            return -1
        if quote < 0 or semi < quote:
            j = semi - 1
            while j > start and data[j].isspace():
                j -= 1
            return j if data[j] == ")" else semi
        # skip the string, '' is an escaped quote inside it
        j = quote + 1
        while True:
            k = find("'", j)
            if k < 0:
                return -1
            if k + 1 < n and data[k + 1] == "'":
                j = k + 2
                continue
            i = k + 1
            break
    return -1
