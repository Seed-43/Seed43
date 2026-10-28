# -*- coding: utf-8 -*-
"""Steel fabrication IFC into Revit: native framing, columns and braces,
DirectShapes for the rest.

    from Snippets import _steel_import
    _steel_import.run(doc, ifc_path, source_name)

A member (a profile swept along a straight line, see _ifc_steel) whose
section matches a Revit family type becomes a native structural framing,
brace or column, rolled and shifted so its section sits where the IFC
outline is. Everything else (plates, bent parts, sections chosen as shapes)
becomes a DirectShape. Bolts, nuts and washers are skipped.

Sections are matched against the loaded framing/column types and the type
catalogues in the Revit content library: a saved choice first, then the
section dimensions, then the section name. Sections that still have no
match are put to the user in one window, and the choices are saved in
.user/Steel IFC/profile_map.json for the next import.

The matching and orientation maths at the top of the file use no Revit
types, so they run under CPython for testing.
"""
import json
import math
import os
import re

try:
    from Snippets._ifc_steel import read_steel, name_key, loop_points, \
        signed_area, Extrusion, Brep
except ImportError:                     # run outside the package (tests)
    from _ifc_steel import read_steel, name_key, loop_points, \
        signed_area, Extrusion, Brep

try:
    from Snippets._shape_families import repeated_shapes, make_shape_families
except ImportError:                     # run outside the package (tests)
    from _shape_families import repeated_shapes, make_shape_families

try:
    from Snippets._bsw_steel import read_steel_bsw, is_bsw
except ImportError:                     # run outside the package (tests)
    from _bsw_steel import read_steel_bsw, is_bsw

try:
    from Snippets._userdata import user_path
except ImportError:
    try:
        from _userdata import user_path
    except ImportError:
        user_path = None

try:
    from Autodesk.Revit import DB
except ImportError:                     # CPython tests
    DB = None

__all__ = ["run", "Entry", "Matcher", "family_kind", "parse_catalogue",
           "decode_catalogue", "entry_values", "profile_values",
           "values_match", "fit_orientation", "section_dirs", "section_axes",
           "rotation_angle", "clip_line", "Mapper", "load_map", "save_map",
           "seed_format", "seed_ifc", "plate_spec"]

TITLE = "Steel IFC"

# ── CONSTANTS ───────────────────────────────────────────────────────────────

MM_PER_FT = 304.8
BATCH = 100                 # members per transaction
SHAPE = "*shape*"           # profile map value: import as a DirectShape
SHAPE_LABEL = u"Import as shape"
PLATE = "*plate*"           # profile map value: import as a Revit steel plate
PLATE_LABEL = u"Import as plate"
# the choices a section can have besides a family type, (label, map value)
SPECIAL_CHOICES = [(SHAPE_LABEL, SHAPE), (PLATE_LABEL, PLATE)]
MAP_TOOL = "Steel IFC"
MAP_FILE = "profile_map.json"
CACHE_FILE = "catalogue_cache.json"
MAX_CHOICES = 150           # candidates listed per section in the window
REPORT_SCORE_MM = 10.0      # orientation fits worse than this are listed

PLATES_NATIVE = "Revit steel plates (as the Steel tab makes)"
PLATES_SHAPE = "Shapes"
REPEATS_FAMILY = "Make them Generic Model families (copies share one)"
REPEATS_SHAPE = "Keep them as shapes"
BOLTS_NATIVE = "Make them Revit steel bolts (as the Steel tab makes)"
BOLTS_SKIP = "Leave the bolts out"
RECT_AREA_TOL = 0.01        # outline area within 1% of its box: a rectangle

PLACE_SHARED = "Shared coordinates (where it was modelled)"
PLACE_INTERNAL = "Internal coordinates"
PLACE_CENTRE = "Centred on the internal origin"

VERTICAL_COS = 0.999999     # axis this close to vertical: a vertical column
TILT_COS = math.cos(math.radians(45.0))     # columns flatter than this: beams
END_CUT_COS = 0.5           # clip planes within 60 deg of the axis trim it

# Family name prefix to section kind (see _ifc_steel.Profile).
FAMILY_PREFIX_KINDS = {
    "UB": "I", "UC": "I", "WB": "I", "WC": "I", "TFB": "I",
    "PFC": "U", "EA": "L", "UA": "L", "BT": "T", "CT": "T",
    "SHS": "RH", "RHS": "RH", "CHS": "CH",
}
# name_key shape word to section kind.
WORD_KINDS = {
    "UB": "I", "UC": "I", "WB": "I", "WC": "I", "TFB": "I", "PFC": "U",
    "EA": "L", "UA": "L", "BT": "T", "CT": "T", "SHS": "RH", "RHS": "RH",
    "CHS": "CH", "ROD": "C", "FL": "R", "PL": "R", "SB": "R",
}
KIND_NAMES = {"I": "I section", "U": "channel", "L": "angle", "T": "tee",
              "RH": "RHS/SHS", "CH": "CHS", "C": "round bar", "R": "flat/plate",
              "CS": "lipped C", "Z": "Z section", "ARB": "outline"}

# Type parameter names (and short aliases) to the dimension they hold.
PARAM_ALIASES = [
    ("height", "H"), ("d", "H"),
    ("width", "W"), ("bf", "W"), ("b", "W"),
    ("flange thickness", "TF"), ("tf", "TF"),
    ("web thickness", "TW"), ("tw", "TW"),
    ("wall nominal thickness", "T"), ("t", "T"), ("thickness", "T"),
    ("diameter", "D"), ("outside diameter", "D"), ("od", "D"),
]
PARAM_NAMES = ["Height", "d", "Width", "bf", "b", "Flange Thickness", "tf",
               "Web Thickness", "tw", "Wall Nominal Thickness", "t",
               "Thickness", "Diameter", "Outside Diameter", "OD"]
_ALIAS = dict(PARAM_ALIASES)

CATALOGUE_UNITS = {"MILLIMETERS": 1.0, "CENTIMETERS": 10.0,
                   "DECIMETERS": 100.0, "METERS": 1000.0,
                   "INCHES": 25.4, "FEET": MM_PER_FT}

_GRADE = re.compile(r"(?<![A-Z])(?:C\d{3}|G\d{3}|\d{3}PLUS)|L0(?![0-9])")


# ── TEXT ────────────────────────────────────────────────────────────────────

try:
    TEXT = unicode      # IronPython 2
except NameError:
    TEXT = str          # CPython 3 (tests)


def safe_text(value):
    try:
        return TEXT(value)
    except Exception:
        return u"?"


def fmt_mm(v):
    if v is None:
        return u"?"
    if abs(v - round(v)) < 0.05:
        return u"{}".format(int(round(v)))
    return u"{:.1f}".format(v)


# ── FAMILIES AND CATALOGUES ─────────────────────────────────────────────────

def family_kind(family):
    """Section kind from a family name: 'UB-Universal Beam' -> 'I'."""
    up = (family or u"").upper()
    if "ROUND BAR" in up:
        return "C"
    if "SQUARE BAR" in up or "FLAT BAR" in up:
        return "R"
    m = re.match(r"\s*([A-Z]+)", up)
    if m:
        return FAMILY_PREFIX_KINDS.get(m.group(1))
    return None


def is_column_library(path, family):
    p = (path or u"").replace("/", "\\").lower()
    return "34-11.1 columns" in p or "-column" in (family or u"").lower()


def decode_catalogue(data):
    """Type catalogue bytes to text (UTF-16 with BOM, UTF-8 or ANSI).
    No isinstance test: IronPython 2 reads bytes into str, which is unicode."""
    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return data.decode("utf-16")
    if data[:3] == b"\xef\xbb\xbf":
        return data[3:].decode("utf-8", "replace")
    try:
        return data.decode("utf-8")
    except UnicodeDecodeError:
        return data.decode("latin-1")


def split_csv(line):
    """One catalogue line to cells; quotes may hold commas."""
    cells = []
    cur = []
    quoted = False
    i = 0
    n = len(line)
    while i < n:
        c = line[i]
        if quoted:
            if c == '"':
                if i + 1 < n and line[i + 1] == '"':
                    cur.append('"')
                    i += 1
                else:
                    quoted = False
            else:
                cur.append(c)
        elif c == '"':
            quoted = True
        elif c == ",":
            cells.append(u"".join(cur).strip())
            cur = []
        else:
            cur.append(c)
        i += 1
    cells.append(u"".join(cur).strip())
    return cells


def parse_catalogue(text):
    """Type catalogue text to [(type name, {param: mm}, section name key)].
    Header cells are 'Name##TYPE##UNITS'; only length columns are kept."""
    lines = [l for l in text.splitlines() if l.strip()]
    if not lines:
        return []
    cols = []
    for h in split_csv(lines[0])[1:]:
        bits = h.split("##")
        name = bits[0].strip()
        unit = bits[2].strip().upper() if len(bits) > 2 else u""
        cols.append((name, CATALOGUE_UNITS.get(unit)))
    rows = []
    for line in lines[1:]:
        cells = split_csv(line)
        tname = cells[0].strip()
        if not tname:
            continue
        dims = {}
        key = u""
        for (name, factor), cell in zip(cols, cells[1:]):
            if name.lower() == "section name key":
                key = cell.strip()
                continue
            if factor is None or not cell:
                continue
            try:
                dims[name] = float(cell) * factor
            except ValueError:
                pass
        rows.append((tname, dims, key))
    return rows


def norm_dims(raw):
    """{parameter name: mm} to {H, W, TF, TW, T, D: mm}."""
    out = {}
    for name, value in raw.items():
        k = _ALIAS.get(name.strip().lower())
        if k and value and k not in out:
            out[k] = float(value)
    return out


class Entry(object):
    """One family type a section can become: loaded or in the library."""

    def __init__(self, family, type_name, dims, column, path=None,
                 symbol=None, section_key=u""):
        self.family = family
        self.type_name = type_name
        self.dims = norm_dims(dims)
        self.column = column
        self.path = path            # .rfa for a library type
        self.symbol = symbol        # FamilySymbol once loaded
        self.loaded = symbol is not None
        self.keys = []
        for text in (type_name, section_key):
            k = name_key(text)
            if k and k not in self.keys:
                self.keys.append(k)
        self.kind = family_kind(family)
        if self.kind is None:
            for k in self.keys:
                if k[0] in WORD_KINDS:
                    self.kind = WORD_KINDS[k[0]]
                    break

    @property
    def label(self):
        return u"{}: {}".format(self.family, self.type_name)

    @property
    def cat(self):
        return "column" if self.column else "framing"


# ── MATCHING ────────────────────────────────────────────────────────────────

def tol(a, b):
    return max(0.6, 0.015 * max(abs(a), abs(b)))


def _desc(a, b):
    return (a, b) if a >= b else (b, a)


def entry_values(kind, d):
    """Comparable dimensions of a catalogue type for a section kind."""
    H, W, TF, TW, T, D = [d.get(k) for k in ("H", "W", "TF", "TW", "T", "D")]
    if kind in ("I", "U", "T"):
        if H and W:
            return (H, W, TF, TW)
    elif kind == "L":
        if H and W:
            return _desc(H, W) + (TF or T or TW,)
    elif kind == "RH":
        if H and W:
            return _desc(H, W) + (T,)
    elif kind == "CH":
        dia = D or H
        if dia:
            return (dia, T)
    elif kind == "C":
        dia = D or H or W
        if dia:
            return (dia, dia)
    elif kind == "R":
        vals = sorted([v for v in (H, W, T) if v], reverse=True)
        if len(vals) == 1:
            return (vals[0], vals[0])
        if len(vals) >= 2:
            if H and W:
                return _desc(H, W)
            return (vals[0], vals[-1])
    return None


def profile_values(prof):
    """Comparable dimensions of an IFC profile, or None (outlines)."""
    d = prof.dims
    k = prof.kind
    g = d.get
    if k in ("I", "U", "T") and g("depth") and g("width"):
        return (g("depth"), g("width"), g("tf"), g("tw"))
    if k in ("L", "RH") and g("depth") and g("width"):
        return _desc(g("depth"), g("width")) + (g("t"),)
    if k == "CH" and g("diameter"):
        return (g("diameter"), g("t"))
    if k == "C" and g("diameter"):
        return (g("diameter"), g("diameter"))
    if k == "R" and g("depth") and g("width"):
        return _desc(g("depth"), g("width"))
    return None


def values_match(a, b):
    """(match, distance): the first two must be present and agree, the
    rest agree where both have them."""
    if a is None or b is None or len(a) != len(b):
        return False, 1e9
    dist = 0.0
    for i, (x, y) in enumerate(zip(a, b)):
        if x is None or y is None:
            if i < 2:
                return False, 1e9
            continue
        if abs(x - y) > tol(x, y):
            return False, 1e9
        dist += abs(x - y)
    return True, dist


def section_kind(prof, section):
    if prof is not None and prof.kind not in ("ARB", None):
        return prof.kind
    k = name_key(section) or (name_key(prof.name) if prof is not None else None)
    if k:
        return WORD_KINDS.get(k[0])
    return None


def bbox_ok(prof, e):
    """The IFC outline's extents agree with the type's Height/Width."""
    if prof is None:
        return False
    d = e.dims
    if e.kind in ("CH", "C"):
        dia = d.get("D") or d.get("H")
        if not dia:
            return None
        want = (dia, dia)
    else:
        if not (d.get("H") and d.get("W")):
            return None
        want = _desc(d["H"], d["W"])
    have = _desc(prof.extent[0], prof.extent[1])
    return all(abs(x - y) <= 2 * tol(x, y) for x, y in zip(have, want))


def name_match(key, e):
    """0 = same name, 1 = one number list starts the other, None = no."""
    for k in e.keys:
        if k[0] != key[0]:
            continue
        if k[1] == key[1]:
            return 0
        n = min(len(k[1]), len(key[1]))
        if k[1][:n] == key[1][:n]:
            return 1
    return None


def grade_miss(family, material):
    """0: the family's grade is the material's, 1: no grade in the family
    name, 2+: a different grade."""
    toks = _GRADE.findall((family or u"").upper())
    if not toks:
        return 1
    mat = (material or u"").upper()
    missing = len([t for t in toks if t not in mat])
    return 0 if missing == 0 else 2 + missing


class Matcher(object):
    """Section + category ('framing' / 'column') to an Entry or SHAPE."""

    def __init__(self, entries, saved=None):
        self.entries = {"framing": [], "column": []}
        self.by_label = {"framing": {}, "column": {}}
        for e in entries:
            if e.label in self.by_label[e.cat]:
                old = self.by_label[e.cat][e.label]
                if old.loaded or not e.loaded:
                    continue        # loaded type wins over its library copy
                self.entries[e.cat].remove(old)
            self.entries[e.cat].append(e)
            self.by_label[e.cat][e.label] = e
        self.saved = saved or {}

    def find(self, cat, label):
        return self.by_label.get(cat, {}).get(label)

    def match(self, section, cat, prof, material=u""):
        """(Entry or SHAPE or None, how)."""
        saved = (self.saved.get(cat) or {}).get(section)
        if saved in (SHAPE, PLATE):
            return saved, "saved choice"
        if saved:
            e = self.find(cat, saved)
            if e is not None:
                return e, "saved choice"
        pool = self.entries.get(cat, [])
        hits = []
        how = "dimensions"
        pv = profile_values(prof) if prof is not None else None
        if pv is not None:
            for e in pool:
                if e.kind != prof.kind:
                    continue
                ok, dist = values_match(pv, entry_values(e.kind, e.dims))
                if ok:
                    hits.append((e, dist))
        if not hits:
            how = "name"
            key = name_key(section) or (name_key(prof.name) if prof else None)
            if key:
                for e in pool:
                    nm = name_match(key, e)
                    if nm is None:
                        continue
                    box = bbox_ok(prof, e)
                    if box is False or (nm == 1 and box is not True):
                        continue
                    hits.append((e, float(nm)))
        if not hits:
            return None, ""
        hits.sort(key=lambda h: (not h[0].loaded,
                                 grade_miss(h[0].family, material),
                                 len(h[0].family), h[1], h[0].label))
        return hits[0][0], how

    def closeness(self, section, prof, e):
        kind = section_kind(prof, section)
        score = 0.0 if (kind is not None and e.kind == kind) else 1000.0
        pv = profile_values(prof) if prof is not None else None
        ev = entry_values(e.kind, e.dims) if e.kind else None
        if pv is not None and ev is not None and len(pv) == len(ev):
            score += sum(abs(x - y) for x, y in zip(pv, ev)
                         if x is not None and y is not None)
        elif prof is not None and e.dims.get("H") and e.dims.get("W"):
            have = _desc(prof.extent[0], prof.extent[1])
            want = _desc(e.dims["H"], e.dims["W"])
            score += abs(have[0] - want[0]) + abs(have[1] - want[1])
        else:
            score += 500.0
        key = name_key(section)
        if key is not None and name_match(key, e) is not None:
            score -= 50.0
        return score

    def candidates(self, section, cat, prof, limit=MAX_CHOICES):
        """[(score, Entry)] best first."""
        scored = [(self.closeness(section, prof, e), e)
                  for e in self.entries.get(cat, [])]
        scored.sort(key=lambda s: (s[0], not s[1].loaded, s[1].label))
        return scored[:limit]


# ── PROFILE MAP ─────────────────────────────────────────────────────────────

def map_path():
    if user_path is None:
        return None
    return user_path(MAP_TOOL, MAP_FILE)


def load_map(path=None):
    path = path or map_path()
    out = {"framing": {}, "column": {}}
    if not path or not os.path.isfile(path):
        return out
    try:
        with open(path, "rb") as f:
            data = json.loads(f.read().decode("utf-8"))
        for cat in out:
            out[cat].update(data.get(cat) or {})
    except Exception:
        pass
    return out


_NON_ASCII = re.compile(u"[^\x00-\x7f]")


def _escape(match):
    code = ord(match.group(0))
    if code > 0xFFFF:
        code -= 0x10000
        return u"\\u%04x\\u%04x" % (0xD800 + (code >> 10), 0xDC00 + (code & 0x3FF))
    return u"\\u%04x" % code


def json_text(data, **kw):
    """json.dumps that survives IronPython 2: its ensure_ascii encoder tries
    to decode text that is already unicode and fails on names such as
    '16Ø THREADED ROD', so non-ASCII characters are escaped here."""
    kw["ensure_ascii"] = False
    return _NON_ASCII.sub(_escape, json.dumps(data, **kw))


def replace_map(data, path=None):
    """Write data as the whole saved map (the Section Map editor)."""
    path = path or map_path()
    if not path:
        return False
    data = {"framing": dict(data.get("framing", {})),
            "column": dict(data.get("column", {}))}
    try:
        with open(path, "wb") as f:
            f.write(json_text(data, indent=2, sort_keys=True).encode("utf-8"))
        return True
    except Exception:
        return False


def save_map(choices, path=None):
    """Merge {cat: {section: label or SHAPE}} into the saved map."""
    path = path or map_path()
    if not path:
        return False
    data = load_map(path)
    for cat, rows in choices.items():
        data.setdefault(cat, {}).update(rows)
    text = json_text(data, indent=2, sort_keys=True)
    try:
        with open(path, "wb") as f:
            f.write(text.encode("utf-8"))
        return True
    except Exception:
        return False


# ── VECTORS ─────────────────────────────────────────────────────────────────

def _add(a, b):
    return (a[0] + b[0], a[1] + b[1], a[2] + b[2])


def _sub(a, b):
    return (a[0] - b[0], a[1] - b[1], a[2] - b[2])


def _mul(a, s):
    return (a[0] * s, a[1] * s, a[2] * s)


def _dot(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def _cross(a, b):
    return (a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2],
            a[0] * b[1] - a[1] * b[0])


def _len(a):
    return _dot(a, a) ** 0.5


def _unit(a):
    l = _len(a)
    return (a[0] / l, a[1] / l, a[2] / l) if l > 1e-12 else None


# ── ORIENTATION ─────────────────────────────────────────────────────────────
# A Revit section has axes (S1, S2) across the member, S1 x S2 = the axis.
# The IFC section has axes (Xp, Yp). The type's geometry, read once from a
# probe, has bounding-box centre m, extents e and centroid f (from the box
# centre) in (S1, S2). A signed permutation A maps IFC section coordinates
# to Revit ones: (u, v) = A (x, y), so S1 = a11 Xp + a12 Yp and
# S2 = a21 Xp + a22 Yp. det A must be h = sign((S1 x S2) . (Xp x Yp)).

SIGNED_PERMS = [((1, 0), (0, 1)), ((-1, 0), (0, -1)),
                ((1, 0), (0, -1)), ((-1, 0), (0, 1)),
                ((0, 1), (1, 0)), ((0, -1), (-1, 0)),
                ((0, 1), (-1, 0)), ((0, -1), (1, 0))]


def _det2(a):
    return a[0][0] * a[1][1] - a[0][1] * a[1][0]


def fit_orientation(probe, extent, centroid, h):
    """(A, score mm): the permutation that best lays the type's section
    (probe = (m, e, f)) over the IFC one (extent, centroid from bbox centre)."""
    m, e, f = probe
    ex, ey = extent
    cx, cy = centroid
    best = None
    for A in SIGNED_PERMS:
        if _det2(A) != h:
            continue
        ae = (abs(A[0][0]) * ex + abs(A[0][1]) * ey,
              abs(A[1][0]) * ex + abs(A[1][1]) * ey)
        ac = (A[0][0] * cx + A[0][1] * cy, A[1][0] * cx + A[1][1] * cy)
        score = math.hypot(e[0] - ae[0], e[1] - ae[1]) + \
            2.0 * math.hypot(f[0] - ac[0], f[1] - ac[1])
        if best is None or score < best[1] - 1e-9:
            best = (A, score)
    return best


def section_dirs(A, xp, yp):
    """The world directions the Revit section axes must take."""
    d1 = _add(_mul(xp, A[0][0]), _mul(yp, A[0][1]))
    d2 = _add(_mul(xp, A[1][0]), _mul(yp, A[1][1]))
    return d1, d2


def section_axes(basis, axis):
    """(S1, S2) of a transform (BasisX, BasisY, BasisZ): the two bases
    across the member axis, in cyclic order (X axis -> Y, Z; Z -> X, Y)."""
    dots = [abs(_dot(b, axis)) for b in basis]
    i = dots.index(max(dots))
    return basis[(i + 1) % 3], basis[(i + 2) % 3]


def handedness(s1, s2, xp, yp):
    return 1 if _dot(_cross(s1, s2), _cross(xp, yp)) >= 0 else -1


def rotation_angle(d1, s1, s2):
    """Angle about S1 x S2 that turns S1 onto d1."""
    return math.atan2(_dot(d1, s2), _dot(d1, s1))


def clip_line(start, end, clips, min_cos=END_CUT_COS):
    """The member line trimmed by its end cuts: plane clips (and bounded
    half-spaces the line passes through) facing along the axis.
    Returns (start, end, trimmed, other clips left)."""
    d = _sub(end, start)
    length = _len(d)
    if length < 1e-9:
        return start, end, False, len(clips)
    u = _mul(d, 1.0 / length)
    lo, hi = 0.0, 1.0
    others = 0
    for c in clips:
        if c.get("kind") not in ("plane", "bounded"):
            others += 1
            continue
        k = c["keep"]
        dk = _dot(d, k)
        if abs(_dot(u, k)) < min_cos:
            others += 1
            continue
        s = -_dot(_sub(start, c["origin"]), k) / dk
        if c["kind"] == "bounded" and not _inside_bound(
                _add(start, _mul(d, s)), c):
            others += 1
            continue
        if dk > 0:
            lo = max(lo, s)
        else:
            hi = min(hi, s)
    if hi - lo < 1e-6:
        return start, end, False, others
    return (_add(start, _mul(d, lo)), _add(start, _mul(d, hi)),
            lo > 1e-9 or hi < 1.0 - 1e-9, others)


def _inside_bound(p, clip):
    o, x, y = clip["plane"]
    q = _sub(p, o)
    return point_in_poly((_dot(q, x), _dot(q, y)), clip["poly"])


def point_in_poly(pt, poly):
    x, y = pt
    inside = False
    n = len(poly)
    for i in range(n):
        x0, y0 = poly[i]
        x1, y1 = poly[(i + 1) % n]
        if (y0 > y) != (y1 > y):
            xc = x0 + (y - y0) * (x1 - x0) / (y1 - y0)
            if xc > x:
                inside = not inside
    return inside


def loop_orientations(loops, xaxis, yaxis, direction):
    """[(loop, flip)]: flip the point order (or circle sense) so the outer
    loop runs anticlockwise looking down the extrusion direction and holes
    clockwise."""
    s = 1.0 if _dot(_cross(xaxis, yaxis), direction) >= 0 else -1.0
    out = []
    for i, loop in enumerate(loops):
        want = 1.0 if i == 0 else -1.0
        if isinstance(loop, tuple) and loop and loop[0] == "circle":
            area = 1.0              # circles are drawn anticlockwise
        else:
            area = signed_area(loop)
        out.append((loop, area * s * want < 0))
    return out


class Mapper(object):
    """File mm to Revit internal feet: p_int = R (p + shift) / 304.8 + o."""

    def __init__(self, basis=None, origin=(0.0, 0.0, 0.0),
                 shift=(0.0, 0.0, 0.0)):
        self.basis = basis or ((1.0, 0.0, 0.0), (0.0, 1.0, 0.0),
                               (0.0, 0.0, 1.0))
        self.origin = origin
        self.shift = shift

    def vector(self, v):
        bx, by, bz = self.basis
        return (bx[0] * v[0] + by[0] * v[1] + bz[0] * v[2],
                bx[1] * v[0] + by[1] * v[1] + bz[1] * v[2],
                bx[2] * v[0] + by[2] * v[1] + bz[2] * v[2])

    def point(self, p):
        q = _mul(_add(p, self.shift), 1.0 / MM_PER_FT)
        return _add(self.vector(q), self.origin)


def classify(member, role):
    """(structural use, category, mode) for a member:
    beam/brace -> framing; column -> column (vertical or slanted), or a
    beam when it leans more than 45 degrees."""
    d = member.direction
    if role == "column":
        c = abs(d[2])
        if c < TILT_COS:
            return "beam", "framing", "beam"
        return "column", "column", "column" if c >= VERTICAL_COS else "slanted"
    return role, "framing", "beam"


def section_text(prof, section):
    if prof is None:
        return u""
    kind = KIND_NAMES.get(prof.kind, prof.kind)
    d = prof.dims
    if prof.kind in ("I", "U", "T"):
        nums = [d.get("depth"), d.get("width"), d.get("tf"), d.get("tw")]
    elif prof.kind in ("L", "RH", "CS"):
        nums = [d.get("depth"), d.get("width"), d.get("t")]
    elif prof.kind == "CH":
        nums = [d.get("diameter"), d.get("t")]
    elif prof.kind == "C":
        nums = [d.get("diameter")]
    else:
        nums = [prof.extent[1], prof.extent[0]]
    return u"{} {}".format(kind, u" x ".join(fmt_mm(v) for v in nums
                                             if v is not None))


# ── PLATES ──────────────────────────────────────────────────────────────────

def _is_rect(pts, bbox):
    """True when an outline fills its bounding box (an axis-aligned rectangle)."""
    box = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
    return box > 0 and abs(abs(signed_area(pts)) - box) <= RECT_AREA_TOL * box


PLATE_FLAT_MM = 0.05        # face points this far off a plate face: not flat
PLATE_SQUARE_COS = 0.002    # side faces must be this square to the plate
HOLE_WIDTH_MM = 60.0        # a curved opening no wider than this is a bolt
HOLE_LENGTH_MM = 150.0      # hole (round, or slotted up to this long)
HOLE_MIN_POINTS = 6         # straight-sided openings are cut-outs, not holes


def is_hole(loop, normal):
    """True for a bolt hole: a small round or slotted opening (a cut loop on
    a plate face, world points), as opposed to a notch or cut-out."""
    if len(loop) < HOLE_MIN_POINTS:
        return False
    n = _unit(normal)
    u = _unit(_sub(loop[1], loop[0]))
    if n is None or u is None:
        return False
    v = _cross(n, u)
    best = None
    # the loop's narrowest box over a few turns, so slots on a skew count
    for k in range(12):
        a = math.pi * k / 12.0
        du = _add(_mul(u, math.cos(a)), _mul(v, math.sin(a)))
        dv = _cross(n, du)
        xs = [_dot(q, du) for q in loop]
        ys = [_dot(q, dv) for q in loop]
        w, h = sorted((max(xs) - min(xs), max(ys) - min(ys)))
        if best is None or w < best[0]:
            best = (w, h)
    return best[0] <= HOLE_WIDTH_MM and best[1] <= HOLE_LENGTH_MM


def plate_spec(part, force=False):
    """How to make a part as a Revit steel plate, or None.

    Returns {"outline": world points (mm) on one face, "normal": the way the
    thickness goes from that face, "thickness": mm}. A plate is one outline
    with no holes, made as:
      - an extrusion swept through its thickness (an IfcPlate),
      - a flat bar extrusion swept along its length (PL/FL sections, which
        Tekla also writes as IfcBeam),
      - a faceted solid that is two matching parallel faces with square
        sides (ProStructures writes some flats that way).
    Straight cuts through the thickness (chamfers, trimmed ends) trim the
    outline; any other cut leaves the part a shape. force: the user mapped
    the part's section to "import as plate", so any flat part qualifies.
    """
    if len(part.solids) != 1:
        return None
    word = (name_key(part.section) or (u"",))[0]
    platey = force or part.ifc_type == "IFCPLATE" or word in ("PL", "FL")
    body = part.solids[0]
    if isinstance(body, Brep):
        return _brep_plate(body) if platey else None
    if not isinstance(body, Extrusion):
        return None
    ex = body
    prof = ex.profile
    if prof is None or len(prof.loops) != 1:
        return None
    o, x, y = ex.plane
    z = _unit(_cross(x, y))
    if z is None or abs(_dot(z, ex.dir)) < 0.9999 or ex.depth <= 0:
        return None
    if not (platey or (prof.kind == "R" and part.role is not None)):
        return None
    pts = loop_points(prof.loops[0])
    if len(pts) < 3:
        return None
    d = ex.dir

    def world(px, py, pz=0.0):
        return (o[0] + px * x[0] + py * y[0] + pz * d[0],
                o[1] + px * x[1] + py * y[1] + pz * d[1],
                o[2] + px * x[2] + py * y[2] + pz * d[2])

    if ex.depth <= min(prof.extent):
        # swept through its thickness: the outline is the plate
        spec = {"outline": [world(px, py) for px, py in pts], "normal": d,
                "thickness": ex.depth}
    elif _is_rect(pts, prof.bbox):
        # a flat bar swept along its length: the plate lies on its wide side
        x0, y0, x1, y1 = prof.bbox
        length = ex.depth
        if prof.extent[0] <= prof.extent[1]:
            spec = {"outline": [world(x0, y0), world(x0, y1),
                                world(x0, y1, length), world(x0, y0, length)],
                    "normal": x, "thickness": x1 - x0}
        else:
            spec = {"outline": [world(x0, y0), world(x1, y0),
                                world(x1, y0, length), world(x0, y0, length)],
                    "normal": y, "thickness": y1 - y0}
    else:
        return None
    return _clip_plate(spec, ex.clips)


HOLLOW_KINDS = ("RH", "CH")  # their inner loop is the hollow, not a hole


def _profile_hole(loop):
    if isinstance(loop, tuple) and loop and loop[0] == "circle":
        return 2.0 * loop[3] <= HOLE_WIDTH_MM
    pts = [(x, y, 0.0) for x, y in loop]
    return is_hole(pts, (0.0, 0.0, 1.0))


def _key(p):
    return (round(p[0], 2), round(p[1], 2), round(p[2], 2))


def strip_holes(part):
    """Take the bolt holes out of a part's own geometry (Tekla and StruCad
    model them into the body), for "Include bolt holes" off: small inner
    loops of extruded outlines and of brep faces, and the bore faces
    between them. Returns the number of hole loops removed."""
    import copy
    removed = 0
    bodies = []
    for body in part.solids:
        if isinstance(body, Extrusion) and len(body.profile.loops) > 1 and                 body.profile.kind not in HOLLOW_KINDS:
            keep = [body.profile.loops[0]] + [
                l for l in body.profile.loops[1:] if not _profile_hole(l)]
            if len(keep) < len(body.profile.loops):
                removed += len(body.profile.loops) - len(keep)
                prof = copy.copy(body.profile)  # profiles are shared
                prof.loops = keep
                body = copy.copy(body)
                body.profile = prof
        elif isinstance(body, Brep):
            rims = set()
            faces = []
            for face in body.faces:
                if not face or len(face) < 2:
                    faces.append(face)
                    continue
                n = _face_normal(face[0])
                keep = [face[0]]
                for loop in face[1:]:
                    if is_hole(loop, n):
                        rims.update(_key(q) for q in loop)
                        removed += 1
                    else:
                        keep.append(loop)
                faces.append(keep)
            if rims:
                # the bore: faces made only of hole rim points
                faces = [f for f in faces if not f or len(f) > 1 or
                         not all(_key(q) in rims for q in f[0])]
                body = Brep(faces)
        bodies.append(body)
    part.solids = bodies
    return removed


def plate_cuts(spec, openings):
    """Contour cut outlines (world points on the plate's face plane) for the
    IFC openings of a plate, or None when an opening is not a straight
    through-cut (the plate is then better made as a shape). Openings that
    miss the plate are left out."""
    n = _unit(spec["normal"])
    o = spec["outline"][0]
    u = _unit(_sub(spec["outline"][1], o))
    if n is None or u is None:
        return None
    v = _cross(n, u)

    def flat(q):                    # onto the face plane
        return _sub(q, _mul(n, _dot(_sub(q, o), n)))

    def uv(q):
        r = _sub(q, o)
        return (_dot(r, u), _dot(r, v))

    plate2d = [uv(q) for q in spec["outline"]]
    lo = (min(p[0] for p in plate2d), min(p[1] for p in plate2d))
    hi = (max(p[0] for p in plate2d), max(p[1] for p in plate2d))
    cuts = []
    for body in openings:
        loop = None
        if isinstance(body, Extrusion) and body.profile is not None:
            if abs(_dot(body.dir, n)) < 0.999 or len(body.profile.loops) != 1:
                return None
            bo, bx, by = body.plane
            loop = [_add(bo, _add(_mul(bx, px), _mul(by, py)))
                    for px, py in loop_points(body.profile.loops[0])]
        elif isinstance(body, Brep):
            best = None
            for face in body.faces:
                if not face or len(face[0]) < 3:
                    continue
                fn = _face_normal(face[0])
                area = _len(fn)
                if area <= 0:
                    continue
                c = abs(_dot(_mul(fn, 1.0 / area), n))
                if c > 0.999:
                    if best is None or area > best[0]:
                        best = (area, face[0])
                elif c > PLATE_SQUARE_COS:
                    return None     # a bevelled or stepped opening
            if best is None:
                return None
            loop = best[1]
        else:
            return None
        loop = drop_inline([flat(q) for q in loop])
        if len(loop) < 3:
            continue
        pts = [uv(q) for q in loop]
        if max(p[0] for p in pts) < lo[0] or min(p[0] for p in pts) > hi[0] or \
                max(p[1] for p in pts) < lo[1] or min(p[1] for p in pts) > hi[1]:
            continue                # misses the plate
        cuts.append(loop)
    return cuts


def _clip_plate(spec, clips):
    """Trim a plate outline by straight cuts through its thickness."""
    outline = spec["outline"]
    n = spec["normal"]
    for clip in clips:
        if clip.get("kind") != "plane" or abs(_dot(clip["keep"], n)) > PLATE_SQUARE_COS:
            return None             # bevelled, bounded or solid cut
        outline = _clip_polygon(outline, clip["origin"], clip["keep"])
        if len(outline) < 3:
            return None
    spec["outline"] = outline
    return spec


def _clip_polygon(pts, origin, keep):
    """Sutherland-Hodgman: the part of a 3D polygon on the keep side."""
    out = []
    n = len(pts)
    for i in range(n):
        a = pts[i]
        b = pts[(i + 1) % n]
        da = _dot(_sub(a, origin), keep)
        db = _dot(_sub(b, origin), keep)
        if da >= 0:
            out.append(a)
        if (da >= 0) != (db >= 0):
            t = da / (da - db)
            out.append(_add(a, _mul(_sub(b, a), t)))
    clean = []
    for p in out:
        if not clean or _len(_sub(p, clean[-1])) > 1e-6:
            clean.append(p)
    if len(clean) > 1 and _len(_sub(clean[0], clean[-1])) <= 1e-6:
        clean.pop()
    return clean


def drop_inline(pts, tol=0.01):
    """An outline without repeated corners and corners that sit on a straight
    edge (within tol mm); the steel plate tool can refuse those."""
    out = []
    for p in pts:
        if not out or _len(_sub(p, out[-1])) > tol:
            out.append(p)
    if len(out) > 1 and _len(_sub(out[0], out[-1])) <= tol:
        out.pop()
    changed = True
    while changed and len(out) > 3:
        changed = False
        for i in range(len(out)):
            a = out[i - 1]
            b = out[i]
            c = out[(i + 1) % len(out)]
            ac = _sub(c, a)
            l = _len(ac)
            if l <= tol or _len(_cross(_sub(b, a), ac)) / l <= tol:
                del out[i]
                changed = True
                break
    return out


# ── BOLTS ───────────────────────────────────────────────────────────────────

BOLT_STANDARD = "AS 1252"   # the AU/NZ structural bolt in Revit's steel data
BOLT_GRADE = "8.8"
BOLT_REACH_MM = 150.0       # how far along a bolt line to look for material
BOLT_STEP_MM = 0.5
BOLT_GAP_MM = 2.0           # gaps up to this between plies still one grip

_M_SIZE = re.compile(r"M\s*(\d{1,2}(?:\.\d)?)", re.I)    # CHEMSET_M16260 is M16
_GRADE_TXT = re.compile(r"(?<![\d.])(\d{1,2}\.\d)(?![\d])")


def bolt_size(props):
    """(diameter mm, grade, on site) from a fastener's properties, as Tekla
    (Bolt Name BOLTM16*25, Bolt standard 8.8S, Location Site) and
    ProStructures (Description 'M20 x 45 8.8/S GALV') write them."""
    props = props or {}
    texts = [safe_text(props.get(k) or u"") for k in
             ("Bolt Name", "_Description", "Description", "Name",
              "Style Name", "Bolt standard")]
    diameter = None
    for txt in texts:
        m = _M_SIZE.search(txt)
        if m:
            diameter = float(m.group(1))
            break
    if diameter is None:
        for k in ("_NominalDiameter", "Nominal Diameter", "Bolt size"):
            try:
                v = float(props.get(k) or 0)
            except (TypeError, ValueError):
                v = 0
            if v > 0:
                diameter = v
                break
    grade = BOLT_GRADE
    for txt in (texts[5], texts[4], texts[1], texts[2], texts[0]):
        m = _GRADE_TXT.search(txt)
        if m:
            grade = m.group(1)
            break
    site = u"shop" not in safe_text(props.get("Location") or u"").lower()
    return diameter, grade, site


def _principal_axis(pts):
    n = float(len(pts))
    c = tuple(sum(p[i] for p in pts) / n for i in range(3))
    cov = [[0.0] * 3 for _ in range(3)]
    for p in pts:
        d = (p[0] - c[0], p[1] - c[1], p[2] - c[2])
        for i in range(3):
            for j in range(3):
                cov[i][j] += d[i] * d[j]
    v = (1.0, 0.7, 0.4)
    for _ in range(60):
        w = tuple(sum(cov[i][j] * v[j] for j in range(3)) for i in range(3))
        v = _unit(w) or v
    return c, v


def bolt_axes(part):
    """Centre lines [(point, direction)] of the bolts in one fastener:
    round shanks/washers grouped by line (Tekla), else the long axis of each
    faceted bolt (ProStructures)."""
    lines = []
    for s in part.solids:
        if isinstance(s, Extrusion) and s.profile is not None and \
                s.profile.kind in ("C", "CH"):
            o, x, y = s.plane
            cx, cy = s.profile.centre
            lines.append((_add(o, _add(_mul(x, cx), _mul(y, cy))), s.dir))
    if not lines:
        for s in part.solids:
            if isinstance(s, Brep):
                pts = [q for f in s.faces for loop in f for q in loop]
                if len(pts) < 6:
                    continue
                # the bolt line is square to its biggest flat face (head or
                # nut top); the long axis of the shape misleads on a nut,
                # which is wider than it is tall
                best = None
                for f in s.faces:
                    if f and len(f[0]) >= 3:
                        fn = _face_normal(f[0])
                        if best is None or _len(fn) > _len(best):
                            best = fn
                c, v = _principal_axis(pts)
                lines.append((c, _unit(best) if best is not None and
                              _len(best) > 0 else v))
    if not lines:
        for s in part.solids:
            if isinstance(s, Extrusion):
                o, x, y = s.plane
                cx, cy = s.profile.centre
                lines.append((_add(o, _add(_mul(x, cx), _mul(y, cy))), s.dir))
    axes = []
    for p, d in lines:
        for i, (q, e, n) in enumerate(axes):
            if abs(_dot(d, e)) > 0.999:
                off = _sub(p, q)
                off = _sub(off, _mul(e, _dot(off, e)))
                if _len(off) < 2.0:
                    axes[i] = (q, e, n + 1)
                    break
        else:
            axes.append((p, d, 1))
    return [(p, d) for p, d, n in axes]


class SolidIndex(object):
    """The plates and profile members of a model, for asking which of them a
    bolt line passes through (and for how long)."""

    def __init__(self, parts, plate_specs=None):
        self.entries = []
        specs = dict((id(p), s) for p, s in (plate_specs or []))
        for part in parts:
            if part.bolt:
                continue
            spec = specs.get(id(part))
            if spec is not None:
                self._add_plate(part, spec)
                continue
            for s in part.solids:
                if isinstance(s, Extrusion) and s.profile is not None:
                    self._add_extrusion(part, s)

    def _add_plate(self, part, spec):
        out = spec["outline"]
        n = _unit(spec["normal"])
        u = _unit(_sub(out[1], out[0]))
        if n is None or u is None:
            return
        v = _cross(n, u)
        o = out[0]
        poly = [(_dot(_sub(q, o), u), _dot(_sub(q, o), v)) for q in out]
        t = spec["thickness"]
        corners = list(out) + [_add(q, _mul(n, t)) for q in out]
        self.entries.append((part, ("plate", o, u, v, n, t, poly, []),
                             _box(corners)))

    def _add_extrusion(self, part, ex):
        o, x, y = ex.plane
        loops = [loop_points(l) for l in ex.profile.loops]
        if not loops or len(loops[0]) < 3:
            return
        corners = []
        for px, py in loops[0]:
            q = _add(o, _add(_mul(x, px), _mul(y, py)))
            corners.append(q)
            corners.append(_add(q, _mul(ex.dir, ex.depth)))
        self.entries.append((part, ("ext", o, x, y, ex.dir, ex.depth,
                                    loops[0], loops[1:]), _box(corners)))

    @staticmethod
    def _inside(geo, p):
        kind, o, a, b, d, depth, outer, inner = geo
        r = _sub(p, o)
        h = _dot(r, d)
        if h < -1e-6 or h > depth + 1e-6:
            return False
        q = (_dot(r, a), _dot(r, b))
        if not point_in_poly(q, outer):
            return False
        return not any(point_in_poly(q, l) for l in inner)

    def grip(self, point, d):
        """(start, end, grip mm, parts crossed) of the material a bolt line
        passes through around point, or None. start is on the far side from
        end along d."""
        reach = BOLT_REACH_MM
        lo = _sub(point, _mul(d, reach))
        hi = _add(point, _mul(d, reach))
        seg = _box([lo, hi])
        near = [(part, geo) for part, geo, box in self.entries
                if _boxes_meet(seg, box)]
        if not near:
            return None
        steps = int(2 * reach / BOLT_STEP_MM) + 1
        hits = []
        for i in range(steps):
            s = -reach + i * BOLT_STEP_MM
            p = _add(point, _mul(d, s))
            inside = [part for part, geo in near if self._inside(geo, p)]
            hits.append((s, inside))
        runs = []
        cur = None
        for s, inside in hits:
            if inside:
                if cur is not None and s - cur[1] <= BOLT_GAP_MM + 1e-9:
                    cur[1] = s
                    cur[2].update(id(p) for p in inside)
                    cur[3].extend(p for p in inside if id(p) not in cur[4])
                    cur[4].update(id(p) for p in inside)
                else:
                    cur = [s, s, set(id(p) for p in inside), list(inside),
                           set(id(p) for p in inside)]
                    runs.append(cur)
        if not runs:
            return None
        best = min(runs, key=lambda r: 0.0 if r[0] <= 0 <= r[1]
                   else min(abs(r[0]), abs(r[1])))
        s0 = best[0] - BOLT_STEP_MM / 2.0
        s1 = best[1] + BOLT_STEP_MM / 2.0
        crossed = []
        seen = set()
        for p in best[3]:
            if id(p) not in seen:
                seen.add(id(p))
                crossed.append(p)
        return (_add(point, _mul(d, s0)), _add(point, _mul(d, s1)), s1 - s0,
                crossed)


def _box(pts):
    return (tuple(min(p[i] for p in pts) for i in range(3)),
            tuple(max(p[i] for p in pts) for i in range(3)))


def _boxes_meet(a, b, pad=1.0):
    return all(a[0][i] <= b[1][i] + pad and b[0][i] <= a[1][i] + pad
               for i in range(3))


ANCHOR_NAME = re.compile(r"CHEM|ANCHOR|HOLD.?DOWN|(?<![A-Z])HD(?![A-Z])|"
                         r"MASONRY|EPOXY|EXPANSION", re.I)
_ANCHOR_LEN = re.compile(r"M\s*\d{1,2}\s*[X*_ -]?\s*(\d{2,4})(?!\d)", re.I)


def _axis_extent(part, point, d):
    """(low, high) of a fastener's geometry along its line, from point."""
    lo, hi = 1e30, -1e30
    for s in part.solids:
        pts = []
        if isinstance(s, Extrusion) and s.profile is not None:
            o, x, y = s.plane
            for px, py in loop_points(s.profile.loops[0]):
                q = _add(o, _add(_mul(x, px), _mul(y, py)))
                pts.append(q)
                pts.append(_add(q, _mul(s.dir, s.depth)))
        elif isinstance(s, Brep):
            pts = [q for f in s.faces for loop in f for q in loop]
        for q in pts:
            t = _dot(_sub(q, point), d)
            lo = min(lo, t)
            hi = max(hi, t)
    return (lo, hi) if lo <= hi else None


def anchor_length(props):
    """Anchor length from its name (CHEMSET_M16260 is M16 x 260), or None."""
    for k in ("Bolt Name", "_Description", "Description", "Name"):
        m = _ANCHOR_LEN.search(safe_text((props or {}).get(k) or u""))
        if m and 50 <= int(m.group(1)) <= 3000:
            return float(m.group(1))
    return None


_NOMINAL_LEN = re.compile(r"M\s*\d{1,2}(?:\.\d)?\s*[X*]\s*(\d{1,4})(?!\d)", re.I)


# Grating names (ProStructures A325MSG, C325MPG...): clipped down, not anchored.
GRATING_NAME = re.compile(r"GRAT|\d{3}M[SP][GU]\b", re.I)


def nominal_length(props):
    """The bolt length the IFC states (M20 x 45, BOLTM16*25, Nominal
    Length), 0 for a nut-only fastener (M16 x 0), or None when not given."""
    props = props or {}
    for k in ("_Description", "Description", "Bolt Name"):
        m = _NOMINAL_LEN.search(safe_text(props.get(k) or u""))
        if m:
            return float(m.group(1))
    for k in ("Nominal Length", "_NominalLength"):
        try:
            v = float(props.get(k))
        except (TypeError, ValueError):
            continue
        return v
    return None


def _base_plate_bolt(crossed, d, names):
    """A vertical bolt through a single plate (not a grating) with nothing
    modelled under it: it goes into concrete or steel outside the model."""
    return len(crossed) == 1 and crossed[0].ifc_type == "IFCPLATE" and \
        abs(d[2]) > 0.9 and not ANCHOR_NAME.search(names) and \
        not GRATING_NAME.search(crossed[0].section or u"")


def _rod_along(part, d):
    """True for a round bar or threaded rod running along d: a fastener on
    it is a nut, not a bolt."""
    m = part.member
    if m is None:
        return False
    word = (name_key(part.section) or (u"",))[0]
    if m.profile.kind != "C" and word != "ROD":
        return False
    md = m.direction
    return md is not None and abs(_dot(md, d)) > 0.98


def plan_bolts(parts, index, skip_base_plates=True):
    """Native bolt and anchor specs for the fasteners in parts, and the ones
    that could not be placed [(part, why)].

    A fastener is an anchor (base plate hold-down, chemical anchor) when its
    geometry runs well past the joint on one side, or its name says so; its
    top is the short side, where the nut is."""
    bolts, skipped = [], []
    for part in parts:
        if part.ifc_type not in ("IFCMECHANICALFASTENER", "IFCFASTENER"):
            continue
        diameter, grade, site = bolt_size(part.props)
        if not diameter:
            skipped.append((part, u"no bolt size"))
            continue
        nominal = nominal_length(part.props)
        if nominal is not None and nominal <= 0:
            skipped.append((part, u"a nut only (bolt length 0), e.g. on a "
                                  u"threaded rod"))
            continue
        axes = bolt_axes(part)
        if not axes:
            skipped.append((part, u"no bolt line in the geometry"))
            continue
        names = u" ".join(safe_text((part.props or {}).get(k) or u"") for k in
                          ("Bolt Name", "_Description", "Description")) +             u" " + safe_text(part.name)
        for point, d in axes:
            g = index.grip(point, d)
            if g is None:
                skipped.append((part, u"passes through no plate or member"))
                continue
            start, end, grip, crossed = g
            if any(_rod_along(c, d) for c in crossed):
                skipped.append((part, u"a nut on a threaded rod"))
                continue
            if skip_base_plates and _base_plate_bolt(crossed, d, names):
                skipped.append((part, u"a base plate bolt (left out; rods "
                                      u"and anchors are kept)"))
                continue
            spec = {"part": part, "head": end,
                    "normal": _unit(_sub(end, start)), "grip": grip,
                    "diameter": diameter, "grade": grade, "site": site,
                    "crossed": crossed, "anchor": False, "length": None,
                    "nominal": nominal}
            ext = _axis_extent(part, point, d)
            if ext is not None:
                t0 = _dot(_sub(start, point), d)
                t1 = _dot(_sub(end, point), d)
                below = t0 - ext[0]         # past the joint on the start side
                above = ext[1] - t1         # past it on the end side
                long_side = max(below, above)
                if long_side > max(60.0, 4.0 * diameter) or                         ANCHOR_NAME.search(names):
                    spec["anchor"] = True
                    if above > below:       # embedded towards the end side
                        spec["head"] = start
                        spec["normal"] = _mul(spec["normal"], -1.0)
                    spec["length"] = anchor_length(part.props) or nominal or                         (ext[1] - ext[0])
            bolts.append(spec)
    return bolts, skipped


def _face_normal(loop):
    """Newell normal (length = 2 x area) of a 3D polygon."""
    nx = ny = nz = 0.0
    m = len(loop)
    for i in range(m):
        a = loop[i]
        b = loop[(i + 1) % m]
        nx += (a[1] - b[1]) * (a[2] + b[2])
        ny += (a[2] - b[2]) * (a[0] + b[0])
        nz += (a[0] - b[0]) * (a[1] + b[1])
    return (nx, ny, nz)


def _brep_plate(br):
    """A faceted solid as a plate: its biggest face, a matching opposite face
    a thickness away, every corner on one of the two, square sides."""
    faces = []
    for face in br.faces:
        if not face or len(face[0]) < 3:
            continue
        nv = _face_normal(face[0])
        area = _len(nv) / 2.0
        if area <= 0:
            continue
        faces.append((area, _unit(nv), face))
    if len(faces) < 4:
        return None
    faces.sort(key=lambda f: -f[0])
    area, n, top = faces[0]
    if len(top) != 1:
        return None                 # holes are not made on plates yet
    p0 = top[0][0]
    bottom = None
    for a2, n2, f2 in faces[1:]:
        if _dot(n, n2) < -0.9999 and abs(a2 - area) <= 0.01 * area:
            bottom = (a2, n2, f2)
            break
    if bottom is None:
        return None
    offset = _dot(_sub(bottom[2][0][0], p0), n)
    thickness = abs(offset)
    if thickness <= 0 or thickness * thickness > area:
        return None                 # thicker than it is wide: not a plate
    for a2, n2, f2 in faces:
        if f2 is top or f2 is bottom[2]:
            continue
        if abs(_dot(n2, n)) > PLATE_SQUARE_COS:
            return None             # a bevelled or stepped side
    for a2, n2, f2 in faces:
        for loop in f2:
            for q in loop:
                h = abs(_dot(_sub(q, p0), n))
                if min(h, abs(h - thickness)) > PLATE_FLAT_MM:
                    return None
    # the thickness runs from the top face towards the opposite one (face
    # winding in IFC files is not always outward, so it is not trusted)
    return {"outline": list(top[0]),
            "normal": n if offset > 0 else _mul(n, -1.0),
            "thickness": thickness}


# ── SEED FILES ──────────────────────────────────────────────────────────────

SEED_MAGIC = b"SEED43-STUDIO-V1"
SEED_BEGIN = b"---BEGIN-SOURCE-DATA---\n"


def seed_fields(data):
    """(header fields, offset of the source data) of a .seed file's bytes."""
    if not data.startswith(SEED_MAGIC):
        raise ValueError("not a Seed43 Studio .seed file")
    start = data.find(SEED_BEGIN)
    if start < 0:
        raise ValueError("the file has no model data")
    fields = {}
    for line in data[:start].decode("utf-8", "replace").splitlines()[1:]:
        if u":" in line:
            key, value = line.split(u":", 1)
            fields[key.strip()] = value.strip()
    return fields, start + len(SEED_BEGIN)


def seed_format(path):
    """The sourceFormat of a .seed file ('navis', 'ifc', 'bsw'...), or ''."""
    try:
        with open(path, "rb") as f:
            head = f.read(4096)
        return seed_fields(head + SEED_BEGIN)[0].get(u"sourceFormat", u"")
    except Exception:
        return u""


def _gunzip(data):
    import struct
    import zlib
    flags = ord(data[3:4])
    pos = 10
    if flags & 4:
        pos += 2 + struct.unpack("<H", data[pos:pos + 2])[0]
    if flags & 8:
        pos = data.index(b"\x00", pos) + 1
    if flags & 16:
        pos = data.index(b"\x00", pos) + 1
    if flags & 2:
        pos += 2
    return zlib.decompressobj(-15).decompress(data[pos:])


SEED_MODELS = {u"ifc": u".ifc", u"bsw": u".bsw"}


def seed_ifc(path, out_dir=None):
    """Write the model a .seed carries (an IFC, or a StruCad .bsw) to a
    temporary file: (file path, source name). Raises ValueError for the
    other formats (meshes and the rest are 3D Import's own)."""
    import tempfile
    with open(path, "rb") as f:
        data = f.read()
    fields, start = seed_fields(data)
    fmt = fields.get(u"sourceFormat", u"")
    if fmt not in SEED_MODELS:
        raise ValueError(u"this .seed holds a {} model, not an IFC or "
                         u"StruCad model".format(fmt or u"unknown"))
    try:
        size = int(fields.get(u"sourceBytes", u""))
    except ValueError:
        raise ValueError("the header has no data size")
    raw = data[start:start + size]
    data = None
    if raw[:2] == b"\x1f\x8b":
        raw = _gunzip(raw)
    name = fields.get(u"sourceName", u"") or os.path.basename(path)
    base = re.sub(r"[^A-Za-z0-9_. -]", u"_", os.path.splitext(name)[0])
    out = os.path.join(out_dir or tempfile.gettempdir(),
                       u"Seed43 {}{}".format(base or u"model",
                                             SEED_MODELS[fmt]))
    with open(out, "wb") as f:
        f.write(raw)
    return out, name


# ════════════════════════════════════════════════════════════════════════════
# REVIT
# ════════════════════════════════════════════════════════════════════════════

# ── REVIT HELPERS ───────────────────────────────────────────────────────────

def element_name(el):
    try:
        return safe_text(DB.Element.Name.__get__(el))
    except Exception:
        try:
            return safe_text(el.Name)
        except Exception:
            return u""


def _xyz(t):
    return DB.XYZ(t[0], t[1], t[2])


def _tup(p):
    return (p.X, p.Y, p.Z)


def _swallow_warnings():
    class SwallowWarnings(DB.IFailuresPreprocessor):
        """Warnings would stop every batch with a dialog; they are dropped,
        errors still roll the batch back."""

        def PreprocessFailures(self, accessor):
            for f in accessor.GetFailureMessages():
                if f.GetSeverity() == DB.FailureSeverity.Warning:
                    accessor.DeleteWarning(f)
            return DB.FailureProcessingResult.Continue
    return SwallowWarnings()


def start_transaction(doc, name):
    t = DB.Transaction(doc, name)
    opts = t.GetFailureHandlingOptions()
    opts.SetFailuresPreprocessor(_swallow_warnings())
    t.SetFailureHandlingOptions(opts)
    t.Start()
    return t


def set_param(el, bip, value):
    try:
        p = el.get_Parameter(bip)
        if p is not None and not p.IsReadOnly:
            p.Set(value)
            return True
    except Exception:
        pass
    return False


# ── CATALOGUE ───────────────────────────────────────────────────────────────

def symbol_dims(sym):
    out = {}
    for name in PARAM_NAMES:
        try:
            p = sym.LookupParameter(name)
        except Exception:
            p = None
        if p is None or p.StorageType != DB.StorageType.Double:
            continue
        v = p.AsDouble() * MM_PER_FT
        if v > 0:
            out[name] = v
    return out


def symbol_section_key(sym):
    try:
        p = sym.LookupParameter("Section Name Key")
        if p is not None and p.StorageType == DB.StorageType.String:
            return safe_text(p.AsString() or u"")
    except Exception:
        pass
    return u""


def loaded_entries(doc):
    out = []
    for bic, column in ((DB.BuiltInCategory.OST_StructuralFraming, False),
                        (DB.BuiltInCategory.OST_StructuralColumns, True)):
        for sym in DB.FilteredElementCollector(doc) \
                .OfClass(DB.FamilySymbol).OfCategory(bic):
            try:
                fam = safe_text(sym.FamilyName)
            except Exception:
                fam = u""
            out.append(Entry(fam, element_name(sym), symbol_dims(sym), column,
                             symbol=sym, section_key=symbol_section_key(sym)))
    return out


def library_roots(app):
    roots = []
    try:
        for v in app.GetLibraryPaths().Values:
            roots.append(safe_text(v))
    except Exception:
        pass
    try:
        ver = safe_text(app.VersionNumber)
    except Exception:
        ver = u""
    pd = os.environ.get("PROGRAMDATA", r"C:\ProgramData")
    roots.append(os.path.join(pd, "Autodesk", "RVT " + ver, "Libraries"))
    out = []
    seen = set()
    for r in roots:
        k = os.path.normcase(os.path.normpath(r))
        if r and k not in seen and os.path.isdir(r):
            seen.add(k)
            out.append(r)
    return out


def find_catalogues(roots):
    """[(rfa path, txt path)] for steel families with a type catalogue."""
    found = []
    seen = set()
    for root in roots:
        for folder, dirs, files in os.walk(root):
            if "steel" not in folder.lower():
                continue
            lower = dict((f.lower(), f) for f in files)
            for f in files:
                if not f.lower().endswith(".rfa"):
                    continue
                txt = lower.get(f[:-4].lower() + ".txt")
                if txt is None:
                    continue
                rfa = os.path.join(folder, f)
                k = os.path.normcase(rfa)
                if k not in seen:
                    seen.add(k)
                    found.append((rfa, os.path.join(folder, txt)))
    return found


def library_entries(app):
    cache_path = user_path(MAP_TOOL, CACHE_FILE) if user_path else None
    cache = {}
    if cache_path and os.path.isfile(cache_path):
        try:
            with open(cache_path, "rb") as f:
                cache = json.loads(f.read().decode("utf-8"))
        except Exception:
            cache = {}
    new_cache = {}
    out = []
    for rfa, txt in find_catalogues(library_roots(app)):
        family = os.path.splitext(os.path.basename(rfa))[0]
        if family_kind(family) is None:
            continue
        try:
            mtime = os.path.getmtime(txt)
        except Exception:
            continue
        hit = cache.get(txt)
        if hit and hit.get("mtime") == mtime:
            rows = hit.get("rows") or []
        else:
            try:
                with open(txt, "rb") as f:
                    rows = parse_catalogue(decode_catalogue(f.read()))
            except Exception:
                rows = []
        new_cache[txt] = {"mtime": mtime, "rows": rows}
        column = is_column_library(rfa, family)
        for tname, dims, key in rows:
            out.append(Entry(family, tname, dims, column, path=rfa,
                             section_key=key))
    if cache_path:
        try:
            with open(cache_path, "wb") as f:
                f.write(json_text(new_cache).encode("utf-8"))
        except Exception:
            pass
    return out


def load_entry(doc, e):
    """Load a library type (in an open transaction): the FamilySymbol."""
    if e.symbol is not None:
        return e.symbol
    sym = None
    try:
        import clr
        ref = clr.Reference[DB.FamilySymbol]()
        if doc.LoadFamilySymbol(e.path, e.type_name, ref):
            sym = ref.Value
    except Exception:
        sym = None
    if sym is None:
        bic = DB.BuiltInCategory.OST_StructuralColumns if e.column \
            else DB.BuiltInCategory.OST_StructuralFraming
        for s in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol) \
                .OfCategory(bic):
            try:
                if safe_text(s.FamilyName) == e.family and \
                        element_name(s) == e.type_name:
                    sym = s
                    break
            except Exception:
                pass
    e.symbol = sym
    return sym


# ── PROBES ──────────────────────────────────────────────────────────────────

PROBE_BASE = (3000.0, 3000.0)       # ft, away from the model
PROBE_SPACING = 40.0
PROBE_LENGTH = 20.0
PROBE_SLANT = math.radians(30.0)


def _solids(geo, out):
    if geo is None:
        return out
    for g in geo:
        if isinstance(g, DB.Solid):
            if g.Volume > 1e-9:
                out.append(g)
        elif isinstance(g, DB.GeometryInstance):
            _solids(g.GetInstanceGeometry(), out)
    return out


def location_base(fi):
    loc = fi.Location
    if isinstance(loc, DB.LocationCurve):
        return _tup(loc.Curve.GetEndPoint(0))
    if isinstance(loc, DB.LocationPoint):
        return _tup(loc.Point)
    return _tup(fi.GetTransform().Origin)


def instance_axes(fi, axis):
    tr = fi.GetTransform()
    basis = (_tup(tr.BasisX), _tup(tr.BasisY), _tup(tr.BasisZ))
    return section_axes(basis, axis)


def measure(fi, axis):
    """(m, e, f) of an instance's section in its (S1, S2), mm, or None."""
    s1, s2 = instance_axes(fi, axis)
    base = location_base(fi)
    opts = DB.Options()
    opts.ComputeReferences = False
    opts.DetailLevel = DB.ViewDetailLevel.Fine
    solids = _solids(fi.get_Geometry(opts), [])
    if not solids:
        return None
    us, vs = [], []
    cu = cv = vol = 0.0
    for s in solids:
        for edge in s.Edges:
            for p in edge.Tessellate():
                q = _sub(_tup(p), base)
                us.append(_dot(q, s1))
                vs.append(_dot(q, s2))
        try:
            c = _sub(_tup(s.ComputeCentroid()), base)
            cu += _dot(c, s1) * s.Volume
            cv += _dot(c, s2) * s.Volume
            vol += s.Volume
        except Exception:
            pass
    if not us:
        return None
    m = ((min(us) + max(us)) / 2.0, (min(vs) + max(vs)) / 2.0)
    e = (max(us) - min(us), max(vs) - min(vs))
    f = (cu / vol - m[0], cv / vol - m[1]) if vol > 0 else (0.0, 0.0)
    k = MM_PER_FT
    return ((m[0] * k, m[1] * k), (e[0] * k, e[1] * k), (f[0] * k, f[1] * k))


def place_member(doc, sym, mode, p0, p1, level, use):
    """A native member at rotation 0; p0/p1 are internal-ft tuples (for
    columns p0 is the bottom)."""
    if mode == "beam":
        st = DB.Structure.StructuralType.Brace if use == "brace" \
            else DB.Structure.StructuralType.Beam
        fi = doc.Create.NewFamilyInstance(
            DB.Line.CreateBound(_xyz(p0), _xyz(p1)), sym, level, st)
        set_param(fi, DB.BuiltInParameter.Y_JUSTIFICATION, 1)
        set_param(fi, DB.BuiltInParameter.Z_JUSTIFICATION, 1)
        for end in (0, 1):
            try:
                DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(fi, end)
            except Exception:
                pass
        return fi
    st = DB.Structure.StructuralType.Column
    if mode == "slanted":
        return doc.Create.NewFamilyInstance(
            DB.Line.CreateBound(_xyz(p0), _xyz(p1)), sym, level, st)
    elev = level.ProjectElevation
    fi = doc.Create.NewFamilyInstance(DB.XYZ(p0[0], p0[1], elev), sym,
                                      level, st)
    set_param(fi, DB.BuiltInParameter.FAMILY_TOP_LEVEL_PARAM, level.Id)
    set_param(fi, DB.BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM,
              p1[2] - elev)
    set_param(fi, DB.BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM,
              p0[2] - elev)
    return fi


def probe_symbols(doc, wanted, level):
    """{(symbol id, mode): (m, e, f)} measured on throwaway instances."""
    out = {}
    if not wanted:
        return out
    t = start_transaction(doc, TITLE + " probe")
    placed = []
    try:
        z = level.ProjectElevation
        for i, (sym, mode) in enumerate(wanted):
            x = PROBE_BASE[0]
            y = PROBE_BASE[1] + i * PROBE_SPACING
            p0 = (x, y, z)
            if mode == "beam":
                p1 = (x + PROBE_LENGTH, y, z)
                axis = (1.0, 0.0, 0.0)
            elif mode == "column":
                p1 = (x, y, z + PROBE_LENGTH)
                axis = (0.0, 0.0, 1.0)
            else:
                axis = (math.sin(PROBE_SLANT), 0.0, math.cos(PROBE_SLANT))
                p1 = _add(p0, _mul(axis, PROBE_LENGTH))
            try:
                fi = place_member(doc, sym, mode, p0, p1, level, "beam")
                placed.append((sym, mode, fi, axis))
            except Exception:
                pass
        doc.Regenerate()
        for sym, mode, fi, axis in placed:
            try:
                r = measure(fi, axis)
            except Exception:
                r = None
            if r is not None:
                out[(sym.Id.ToString(), mode)] = r
    finally:
        t.RollBack()
    return out


# ── SHAPES ──────────────────────────────────────────────────────────────────

class ShapeMaker(object):
    def __init__(self, doc, mapper):
        self.doc = doc
        self.mapper = mapper
        self.tol = doc.Application.ShortCurveTolerance
        self.clip_failures = 0

    def _world(self, plane, x, y):
        o, ax, ay = plane
        return (o[0] + x * ax[0] + y * ay[0], o[1] + x * ax[1] + y * ay[1],
                o[2] + x * ax[2] + y * ay[2])

    def _curve_loop(self, plane, loop, flip):
        mp = self.mapper
        cl = DB.CurveLoop()
        if isinstance(loop, tuple) and loop and loop[0] == "circle":
            c = _xyz(mp.point(self._world(plane, loop[1], loop[2])))
            r = loop[3] / MM_PER_FT
            xv = _xyz(mp.vector(plane[1]))
            yv = _xyz(mp.vector(plane[2]))
            if flip:
                yv = yv.Negate()
            cl.Append(DB.Arc.Create(c, r, 0.0, math.pi, xv, yv))
            cl.Append(DB.Arc.Create(c, r, math.pi, 2 * math.pi, xv, yv))
            return cl
        pts = [_xyz(mp.point(self._world(plane, x, y))) for x, y in loop]
        if flip:
            pts.reverse()
        clean = []
        for p in pts:
            if not clean or p.DistanceTo(clean[-1]) > self.tol:
                clean.append(p)
        while len(clean) > 2 and clean[0].DistanceTo(clean[-1]) <= self.tol:
            clean.pop()
        if len(clean) < 3:
            return None
        for i in range(len(clean)):
            cl.Append(DB.Line.CreateBound(clean[i], clean[(i + 1) % len(clean)]))
        return cl

    def _prism(self, plane, loops, direction, depth_mm):
        from System.Collections.Generic import List
        curve_loops = List[DB.CurveLoop]()
        for loop, flip in loop_orientations(loops, plane[1], plane[2],
                                            direction):
            cl = self._curve_loop(plane, loop, flip)
            if cl is None:
                if not curve_loops.Count:
                    return None
                continue
            curve_loops.Add(cl)
        return DB.GeometryCreationUtilities.CreateExtrusionGeometry(
            curve_loops, _xyz(self.mapper.vector(direction)),
            depth_mm / MM_PER_FT)

    def extrusion(self, ex):
        solid = self._prism(ex.plane, ex.profile.loops, ex.dir, ex.depth)
        if solid is None:
            return None
        size = ex.depth + max(ex.profile.extent) + 1000.0
        for clip in ex.clips:
            try:
                cut = self._clip(solid, clip, size)
            except Exception:
                cut = None
            if cut is None:
                self.clip_failures += 1
            elif cut is False:
                return None         # everything cut away
            else:
                solid = cut
        return solid

    def _clip(self, solid, clip, size):
        mp = self.mapper
        B = DB.BooleanOperationsUtils
        if clip["kind"] == "plane":
            plane = DB.Plane.CreateByNormalAndOrigin(
                _xyz(mp.vector(clip["keep"])), _xyz(mp.point(clip["origin"])))
            out = B.CutWithHalfSpace(solid, plane)
            return out if out is not None and out.Volume > 1e-9 else False
        if clip["kind"] == "bounded":
            bo, bx, by = clip["plane"]
            axis = clip["axis"]
            start = _sub(bo, _mul(axis, size))
            prism = self._prism((start, bx, by), [clip["poly"]], axis,
                                2.0 * size)
            if prism is None:
                return None
            away = clip["keep"]
            plane = DB.Plane.CreateByNormalAndOrigin(
                _xyz(mp.vector((-away[0], -away[1], -away[2]))),
                _xyz(mp.point(clip["origin"])))
            removed = B.CutWithHalfSpace(prism, plane)
            if removed is None or removed.Volume < 1e-9:
                return solid
            return B.ExecuteBooleanOperation(
                solid, removed, DB.BooleanOperationsType.Difference)
        if clip["kind"] == "solid":
            other = self.solid_of(clip["solid"])
            if other is None:
                return None
            return B.ExecuteBooleanOperation(
                solid, other, DB.BooleanOperationsType.Difference)
        return None

    def solid_of(self, body):
        if isinstance(body, Extrusion):
            return self.extrusion(body)
        return None

    def brep(self, br):
        from System.Collections.Generic import List, IList
        mp = self.mapper
        b = DB.TessellatedShapeBuilder()
        b.Target = DB.TessellatedShapeBuilderTarget.AnyGeometry
        b.Fallback = DB.TessellatedShapeBuilderFallback.Mesh
        b.OpenConnectedFaceSet(True)
        faces = 0
        for face in br.faces:
            loops = List[IList[DB.XYZ]]()
            for loop in face:
                pts = List[DB.XYZ]()
                last = None
                for p in loop:
                    q = _xyz(mp.point(p))
                    if last is None or q.DistanceTo(last) > 1e-6:
                        pts.Add(q)
                        last = q
                if pts.Count >= 3:
                    loops.Add(pts)
            if loops.Count:
                b.AddFace(DB.TessellatedFace(loops, DB.ElementId.InvalidElementId))
                faces += 1
        b.CloseConnectedFaceSet()
        if not faces:
            return []
        b.Build()
        res = b.GetBuildResult()
        if res.Outcome == DB.TessellatedShapeBuilderOutcome.Nothing:
            return []
        return list(res.GetGeometricalObjects())

    def geometry(self, part):
        out = []
        for body in part.solids:
            if isinstance(body, Extrusion):
                s = self.extrusion(body)
                if s is not None:
                    out.append(s)
            elif isinstance(body, Brep):
                out.extend(self.brep(body))
        return out


class Categories(object):
    def __init__(self, doc):
        self.doc = doc
        self._cache = {}

    def _valid(self, bic):
        try:
            cid = DB.ElementId(bic)
            if DB.DirectShape.IsValidCategoryId(cid, self.doc):
                return cid
        except Exception:
            pass
        return None

    def get(self, ifc_type):
        hit = self._cache.get(ifc_type)
        if hit is None:
            B = DB.BuiltInCategory
            want = {"IFCPLATE": getattr(B, "OST_StructConnectionPlates", None),
                    "IFCBEAM": B.OST_StructuralFraming,
                    "IFCMEMBER": B.OST_StructuralFraming,
                    "IFCCOLUMN": B.OST_StructuralColumns}.get(ifc_type)
            hit = (self._valid(want) if want is not None else None) or \
                DB.ElementId(B.OST_GenericModel)
            self._cache[ifc_type] = hit
        return hit


STEEL_ASSEMBLIES = ("ASGeometryMgd", "ASObjectsMgd", "Autodesk.SteelConnectionsDB")


def steel_api():
    """Revit's steel fabrication types (the Steel tab's Plate), or None when
    this Revit has no steel connections add-in."""
    import clr
    import System
    loaded = dict((a.GetName().Name, a)
                  for a in System.AppDomain.CurrentDomain.GetAssemblies())
    folder = os.path.join(os.path.dirname(
        System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName),
        "AddIns", "SteelConnections")
    try:
        for name in STEEL_ASSEMBLIES:
            if name in loaded:
                clr.AddReference(loaded[name])
            else:
                clr.AddReferenceToFileAndPath(os.path.join(folder, name + ".dll"))
        from Autodesk.AdvanceSteel.Geometry import Point3d, Vector3d
        from Autodesk.AdvanceSteel.Geometry import Plane as SteelPlane
        from Autodesk.AdvanceSteel.Modelling import Plate, \
            InfinitMidScrewBoltPattern, AnchorPattern, PlateFeatContour
        from Autodesk.AdvanceSteel.Geometry import Point2d, Matrix2d
        from Autodesk.AdvanceSteel.CADAccess import FilerObject
        from Autodesk.AdvanceSteel.ConstructionTypes import AtomicElement
        from Autodesk.SteelConnectionsDB import FabricationTransaction
    except Exception:
        return None
    return {"point": Point3d, "vector": Vector3d, "plane": SteelPlane,
            "plate": Plate, "bolt": InfinitMidScrewBoltPattern, "anchor": AnchorPattern,
            "contour": PlateFeatContour, "point2d": Point2d, "matrix2d": Matrix2d,
            "filer": FilerObject, "location": AtomicElement.eAssemblyLocation,
            "transaction": FabricationTransaction}


class PlateMaker(object):
    """Native Revit steel plates (Structural Connections > Plates).

    Measured in Revit 2026: a plate made on a plane with normal n and
    Portioning 1 runs from the outline's plane a thickness along n, which is
    the IFC extrusion exactly. Coordinates are internal, in millimetres. The
    new plates come out in the order they were written, which is how each
    one gets its part's Mark and Comments.
    """

    def __init__(self, doc, mapper):
        self.doc = doc
        self.mapper = mapper
        self.api = steel_api()
        self.cuts = 0
        self.cuts_failed = 0

    def _point(self, p):
        q = self.mapper.point(p)
        return self.api["point"](q[0] * MM_PER_FT, q[1] * MM_PER_FT,
                                 q[2] * MM_PER_FT)

    def make(self, specs, bolts=(), progress=None):
        """Write plates for [(part, spec)] and bolts (see plan_bolts) inside
        the open Revit transaction. Returns (made plate parts, [(part, why)]
        plates that failed, bolts made, [(part, why)] bolts that failed).

        Each plate is found again by its fabrication GUID, so a plate Revit
        made although WriteToDb complained counts as made (it is not made a
        second time as a shape), and Mark/Comments land on the right plate.
        Bolts are written in the same steel transaction so they can connect
        to the plates they pass through (Revit then drills those plates)."""
        import System
        from Autodesk.Revit.DB.Steel import SteelElementProperties
        api = self.api
        written = []            # (part, guid or None, error text)
        plates = {}             # id(part) -> steel plate object
        bolts_written = []      # (bolt spec, guid or None, error text)
        fab = api["transaction"](self.doc, False, TITLE + " plates", True)
        try:
            for i, (part, spec) in enumerate(specs):
                if progress is not None and i % SHOW_EVERY == 0:
                    progress(u"Installing plates", u"{} {:g} mm".format(
                        part.section or part.name, spec["thickness"]), i)
                plate = None
                error = u""
                try:
                    outline = drop_inline(spec["outline"])
                    if len(outline) < 3:
                        raise ValueError("outline has no area")
                    pts = [self._point(p) for p in outline]
                    n = self.mapper.vector(spec["normal"])
                    plane = api["plane"](pts[0], api["vector"](n[0], n[1], n[2]))
                    arr = System.Array[api["point"]](pts)
                    plate = api["plate"](plane, arr, float(spec["thickness"]))
                    plate.Portioning = 1.0
                    plate.WriteToDb()
                except Exception as ex:
                    error = safe_text(ex)
                if plate is not None and not error:
                    for loop in spec.get("cuts") or ():
                        try:
                            self._cut(plate, loop)
                            self.cuts += 1
                        except Exception:
                            self.cuts_failed += 1
                guid = None
                if plate is not None:
                    try:
                        guid = plate.GetUniqueId()
                    except Exception:
                        guid = None
                    if not error:
                        plates[id(part)] = (plate, float(spec["thickness"]))
                written.append((part, guid, error))
            for j, b in enumerate(bolts):
                if progress is not None and j % SHOW_EVERY == 0:
                    progress(u"Connecting anchors" if b.get("anchor") else u"Connecting bolts",
                             u"M{:g} {}".format(b["diameter"], b["grade"]),
                             len(specs) + j)
                pattern = None
                error = u""
                try:
                    pattern = self._bolt(b, plates)
                except Exception as ex:
                    error = safe_text(ex)
                guid = None
                if pattern is not None:
                    try:
                        guid = pattern.GetUniqueId()
                    except Exception:
                        guid = None
                bolts_written.append((b, guid, error))
            fab.Commit()
        finally:
            fab.Dispose()

        def element(guid):
            if guid is None or guid == System.Guid.Empty:
                return None
            try:
                ref = SteelElementProperties.GetReference(self.doc, guid)
                return self.doc.GetElement(ref) if ref is not None else None
            except Exception:
                return None

        made, failed = [], []
        for part, guid, error in written:
            el = element(guid)
            if el is None:
                failed.append((part, error or u"Revit made no plate"))
                continue
            set_param(el, DB.BuiltInParameter.ALL_MODEL_MARK, part.mark or u"")
            set_param(el, DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
                      comments(part))
            made.append(part)
        bolts_made, bolts_failed = [], []
        for b, guid, error in bolts_written:
            el = element(guid)
            if el is None:
                bolts_failed.append((b["part"], error or u"Revit made no bolt"))
                continue
            part = b["part"]
            set_param(el, DB.BuiltInParameter.ALL_MODEL_MARK, part.mark or u"")
            set_param(el, DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
                      u"{} | {}M{:g} {} | {} | {}".format(
                          ifc_class(part.ifc_type),
                          u"anchor " if b.get("anchor") else u"",
                          b["diameter"], b["grade"],
                          u"length {:.0f}".format(b["length"]) if b.get("anchor")
                          else u"grip {:.0f}".format(b["grip"]),
                          part.assembly or part.name))
            bolts_made.append(b)
        return made, failed, bolts_made, bolts_failed

    def _cut(self, plate, loop):
        """A contour cut through a plate (the Steel tab's Contour Cut): a hole
        where the outline is inside the plate, a notch where it crosses the
        edge. Its points are in the plate's own system (measured in Revit
        2026: origin on the plate's first corner), so world points go through
        the inverse of the plate's CS."""
        import System
        api = self.api
        inv = plate.CS.GetInverse()
        pts = []
        for q in loop:
            p = self._point(q)
            moved = p.TransformBy(inv)
            if moved is not None:
                p = moved
            pts.append(api["point2d"](p.x, p.y))
        feat = api["contour"](api["matrix2d"](),
                              System.Array[api["point2d"]](pts))
        plate.AddFeature(feat)

    def _bolt(self, b, plates):
        """One bolt: a 1 x 1 bolt pattern centred on the head-side face of the
        joint, its normal pointing out of the joint (measured in Revit 2026;
        a circle pattern will not take radius 0 and puts the bolt 100 mm off),
        connected to the native plates it passes through. Revit works the
        grip out from those plates; the rest of the IFC grip (plies that are
        members or shapes) goes in as Grip Length Increase."""
        import System
        api = self.api
        n = _unit(self.mapper.vector(b["normal"]))
        ref = (1.0, 0.0, 0.0) if abs(n[0]) < 0.9 else (0.0, 1.0, 0.0)
        vx = _unit(_cross(n, ref))
        vy = _cross(n, vx)
        kind = "anchor" if b.get("anchor") else "bolt"
        pattern = api[kind](self._point(b["head"]),
                            api["vector"](vx[0], vx[1], vx[2]),
                            api["vector"](vy[0], vy[1], vy[2]))
        pattern.Nx = 1
        pattern.Ny = 1
        if kind == "bolt":
            pattern.Standard = BOLT_STANDARD
            pattern.Grade = b["grade"]
        pattern.ScrewDiameter = float(b["diameter"])
        if kind == "anchor" and b.get("length"):
            pattern.ScrewLength = float(b["length"])
        pattern.WriteToDb()
        hit = [plates[id(p)] for p in b["crossed"] if id(p) in plates]
        if hit:
            loc = api["location"].kOnSite if b["site"] else api["location"].kInShop
            pattern.Connect(System.Array[api["filer"]]([o for o, t in hit]), loc)
            if kind == "anchor":
                return pattern      # its length is its own, not the grip
            # Revit sets the grip from the connected plates only; plies that
            # are members or shapes are added as Grip Length Increase
            extra = float(b["grip"]) - sum(t for o, t in hit)
        elif kind == "bolt":
            pattern.BindingLength = float(b["grip"])
            extra = 0.0
        if extra > 0.5:
            pattern.BindingLengthAddition = extra
        return pattern


# ── IMPORT ──────────────────────────────────────────────────────────────────

def ifc_class(t):
    """IFCBEAM -> IfcBeam (for Comments)."""
    names = {"IFCBEAM": "IfcBeam", "IFCCOLUMN": "IfcColumn",
             "IFCMEMBER": "IfcMember", "IFCPLATE": "IfcPlate",
             "IFCDISCRETEACCESSORY": "IfcDiscreteAccessory",
             "IFCBUILDINGELEMENTPROXY": "IfcBuildingElementProxy",
             "IFCMECHANICALFASTENER": "IfcMechanicalFastener",
             "IFCELEMENTASSEMBLY": "IfcElementAssembly"}
    return names.get(t, t.title() if t else u"")


def comments(part):
    bits = [ifc_class(part.ifc_type), part.name, part.section, part.material,
            part.assembly]
    return u" | ".join(safe_text(b) for b in bits)


def make_mapper(doc, model, placement):
    if placement == PLACE_SHARED:
        t = doc.ActiveProjectLocation.GetTotalTransform()
        return Mapper((_tup(t.BasisX), _tup(t.BasisY), _tup(t.BasisZ)),
                      _tup(t.Origin))
    if placement == PLACE_CENTRE:
        b = model_bounds(model)
        if b is not None:
            lo, hi = b
            return Mapper(shift=tuple(-(lo[i] + hi[i]) / 2.0 for i in range(3)))
    return Mapper()


def model_bounds(model):
    b = model.bounds()
    if b is not None:
        return b
    lo = [1e30] * 3
    hi = [-1e30] * 3
    for p in model.parts:
        for s in p.solids:
            pts = []
            if isinstance(s, Extrusion):
                pts.append(s.plane[0])
            elif isinstance(s, Brep):
                pts.extend(q for f in s.faces[:4] for q in f[0][:1])
            for q in pts:
                for i in range(3):
                    lo[i] = min(lo[i], q[i])
                    hi[i] = max(hi[i], q[i])
    return (lo, hi) if lo[0] <= hi[0] else None


class Plan(object):
    """One native member to make."""

    def __init__(self, part, use, cat, mode, entry, how):
        self.part = part
        self.use = use
        self.cat = cat
        self.mode = mode
        self.entry = entry
        self.how = how
        self.trimmed = False
        self.notches = 0
        self.fi = None
        self.axis = None
        self.target = None      # (d1, d2) in internal coordinates
        self.phi = 0.0
        self.sign = -1.0
        self.flipped = False
        self.shift = (0.0, 0.0)
        self.score = None
        self.verified = True


def level_for(levels, z):
    """Highest level at or below z (internal ft), else the lowest."""
    best = None
    for lv in levels:
        if lv.ProjectElevation <= z + 1e-6:
            best = lv
    return best or levels[0]


class Importer(object):
    def __init__(self, doc, model, mapper, source_name):
        self.doc = doc
        self.model = model
        self.mapper = mapper
        self.source = source_name or u"IFC"
        self.levels = []
        self.probes = {}
        self.sign = {"beam": -1.0, "slanted": -1.0}
        self.made = {}
        self.shapes = 0
        self.failed = []
        self.fits = []
        self.unverified = 0
        self.fallback = []      # natives that became shapes
        self.cats = Categories(doc)
        self.maker = ShapeMaker(doc, mapper)
        self.plates = 0
        self.plate_failed = []  # (part, why) made as shapes instead
        self.bolts = 0
        self.anchors = 0
        self.shape_families = None  # _shape_families.ShapeFamilyResult
        self.bolt_failed = []   # (fastener part, why)
        self.grids_levels = None    # _grids_levels.GridsLevelsResult

    # -- levels -----------------------------------------------------------------

    def ensure_levels(self):
        self.levels = sorted(DB.FilteredElementCollector(self.doc)
                             .OfClass(DB.Level), key=lambda l: l.ProjectElevation)
        if not self.levels:
            self.levels = [DB.Level.Create(self.doc, 0.0)]

    # -- native members ---------------------------------------------------------

    def ends(self, plan):
        m = plan.part.member
        clips = plan.part.solids[0].clips
        s, e, trimmed, others = clip_line(m.start, m.end, clips)
        plan.trimmed = trimmed
        plan.notches = others
        p0 = self.mapper.point(s)
        p1 = self.mapper.point(e)
        if plan.cat == "column" and p1[2] < p0[2]:
            p0, p1 = p1, p0
        return p0, p1

    def create(self, plan):
        sym = plan.entry.symbol
        p0, p1 = self.ends(plan)
        level = level_for(self.levels, min(p0[2], p1[2]))
        fi = place_member(self.doc, sym, plan.mode, p0, p1, level, plan.use)
        plan.fi = fi
        plan.axis = _unit(_sub(p1, p0))
        part = plan.part
        set_param(fi, DB.BuiltInParameter.ALL_MODEL_MARK, part.mark or u"")
        set_param(fi, DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
                  comments(part))

    def orient(self, plan):
        """Work out (d1, d2) and turn the member to it (rotation only)."""
        fi = plan.fi
        m = plan.part.member
        prof = m.profile
        s1, s2 = instance_axes(fi, plan.axis)
        xp = self.mapper.vector(m.x)
        yp = self.mapper.vector(m.y)
        h = handedness(s1, s2, xp, yp)
        probe = self.probes.get((plan.entry.symbol.Id.ToString(), plan.mode))
        if probe is None:
            probe = ((0.0, 0.0), prof.extent, prof.centroid)
            plan.verified = False
        A, score = fit_orientation(probe, prof.extent, prof.centroid, h)
        plan.score = score
        plan.shift = probe[0]
        d1, d2 = section_dirs(A, xp, yp)
        plan.target = (d1, d2)
        phi = rotation_angle(d1, s1, s2)
        if abs(phi) < 1e-9:
            return
        plan.phi = phi
        plan.flipped = False
        plan.sign = self.sign.get(plan.mode, 1.0)
        self._turn(plan, phi if plan.mode == "column" else plan.sign * phi)

    def _turn(self, plan, angle):
        """Vertical columns turn about their axis; framing and slanted
        columns take a cross-section rotation (absolute, not added)."""
        fi = plan.fi
        if plan.mode == "column":
            base = location_base(fi)
            axis = DB.Line.CreateBound(_xyz(base), _xyz(_add(base, (0, 0, 1))))
            DB.ElementTransformUtils.RotateElement(self.doc, fi.Id, axis, angle)
        else:
            set_param(fi, DB.BuiltInParameter.STRUCTURAL_BEND_DIR_ANGLE, angle)

    def check(self, plan, final=False):
        """True when the section axes are where they should be. On the first
        check a member turned the wrong way is turned the other way; when
        that proves right, later members of that mode start with it."""
        if plan.target is None:
            return True
        s1, s2 = instance_axes(plan.fi, plan.axis)
        if _dot(s1, plan.target[0]) > 0.999:
            if final and plan.flipped and plan.mode != "column":
                self.sign[plan.mode] = -plan.sign
            return True
        if not final:
            phi = plan.phi
            plan.flipped = True
            if plan.mode == "column":
                self._turn(plan, -2.0 * phi)
            else:
                self._turn(plan, -plan.sign * phi)
        return False

    def shift(self, plan):
        s1, s2 = instance_axes(plan.fi, plan.axis)
        m1, m2 = plan.shift
        move = _mul(_add(_mul(s1, m1), _mul(s2, m2)), -1.0 / MM_PER_FT)
        if _len(move) > 1e-7:
            DB.ElementTransformUtils.MoveElement(self.doc, plan.fi.Id,
                                                 _xyz(move))

    def natives(self, plans, pb=None, done=0, total=1):
        for start in range(0, len(plans), BATCH):
            if pb is not None and pb.cancelled:
                return False
            batch = plans[start:start + BATCH]
            t = start_transaction(self.doc, TITLE)
            made = []
            for k, plan in enumerate(batch):
                if k % SHOW_EVERY == 0:
                    show_progress(pb, u"Placing members", u"{} as {}".format(
                        plan.part.section, plan.entry.type_name),
                        done + start + k, total)
                try:
                    self.create(plan)
                    made.append(plan)
                except Exception as ex:
                    self.fallback.append((plan, safe_text(ex)))
            self.doc.Regenerate()
            for plan in made:
                try:
                    self.orient(plan)
                except Exception as ex:
                    plan.verified = False
                    self.failed.append((plan.part.section,
                                        u"orientation: " + safe_text(ex)))
            self.doc.Regenerate()
            redo = [p for p in made if not self.check(p)]
            if redo:
                self.doc.Regenerate()
                for p in redo:
                    if not self.check(p, final=True):
                        p.verified = False
            for plan in made:
                try:
                    self.shift(plan)
                except Exception:
                    pass
            if t.Commit() == DB.TransactionStatus.Committed:
                for plan in made:
                    key = plan.use
                    self.made[key] = self.made.get(key, 0) + 1
                    if not plan.verified:
                        self.unverified += 1
                    if plan.score is not None and plan.score > REPORT_SCORE_MM:
                        self.fits.append(plan)
            else:
                self.failed.append((u"Members {} to {}".format(
                    start + 1, start + len(batch)), u"Revit rolled the batch back"))
            if pb is not None:
                pb.update_progress(done + start + len(batch), total)
        return True

    # -- plates -----------------------------------------------------------------

    def steel_of(self, specs, bolts, maker, pb=None, done=0, total=1):
        """Native plates for [(part, spec)] and native bolts, in one steel
        transaction so bolts can connect to the plates. Plate failures go to
        plate_failed (they become shapes), bolt failures to bolt_failed."""
        if pb is not None and pb.cancelled:
            return False
        t = start_transaction(self.doc, TITLE + " plates and bolts")

        def progress(stage, what, i):
            show_progress(pb, stage, what, done + i, total)

        try:
            made, failed, bmade, bfailed = maker.make(specs, bolts, progress)
        except Exception as ex:
            made, bmade = [], []
            failed = [(p, safe_text(ex)) for p, s in specs]
            bfailed = [(b["part"], safe_text(ex)) for b in bolts]
        if t.Commit() == DB.TransactionStatus.Committed:
            self.plates += len(made)
            self.plate_failed.extend(failed)
            self.bolts += len([b for b in bmade if not b.get("anchor")])
            self.anchors += len([b for b in bmade if b.get("anchor")])
            self.bolt_failed.extend(bfailed)
        else:
            self.plate_failed.extend((p, u"Revit rolled the plates back")
                                     for p, s in specs)
            self.bolt_failed.extend((b["part"], u"Revit rolled the bolts back")
                                    for b in bolts)
        if pb is not None:
            pb.update_progress(done + len(specs) + len(bolts), total)
        return True

    # -- shapes -----------------------------------------------------------------

    def shapes_of(self, parts, pb=None, done=0, total=1):
        from System.Collections.Generic import List
        for start in range(0, len(parts), BATCH):
            if pb is not None and pb.cancelled:
                return False
            batch = parts[start:start + BATCH]
            t = start_transaction(self.doc, TITLE)
            made = 0
            for k, part in enumerate(batch):
                ds = None
                label = part.section or part.name or ifc_class(part.ifc_type)
                if k % SHOW_EVERY == 0:
                    show_progress(pb, u"Locating shapes", label, done + start + k, total)
                try:
                    geo = self.maker.geometry(part)
                    if not geo:
                        self.failed.append((label, u"no usable geometry"))
                        continue
                    ds = DB.DirectShape.CreateElement(
                        self.doc, self.cats.get(part.ifc_type))
                    ds.ApplicationId = "Seed43 Steel IFC"
                    ds.ApplicationDataId = part.guid or self.source
                    ds.SetShape(List[DB.GeometryObject](geo))
                    try:
                        ds.SetName(_clean_name(label))
                    except Exception:
                        pass
                    set_param(ds, DB.BuiltInParameter.ALL_MODEL_MARK,
                              part.mark or u"")
                    set_param(ds, DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
                              comments(part))
                    made += 1
                except Exception as ex:
                    self.failed.append((label, safe_text(ex)))
                    if ds is not None:
                        try:
                            self.doc.Delete(ds.Id)
                        except Exception:
                            pass
            if t.Commit() == DB.TransactionStatus.Committed:
                self.shapes += made
            else:
                self.failed.append((u"Shapes {} to {}".format(
                    start + 1, start + len(batch)), u"Revit rolled the batch back"))
            if pb is not None:
                pb.update_progress(done + start + len(batch), total)
        return True


SHOW_EVERY = 5              # progress bar updates every this many items


def show_progress(pb, stage, what, value, total):
    """Progress bar title that says what is being made, e.g.
    'Native members: 180PFC20.9 (131 of 415)'."""
    if pb is None:
        return
    text = safe_text(what or u"").strip()
    # the title is a format string: braces in names ({WG} marks) would break it
    text = text.replace(u"{", u"(").replace(u"}", u")")
    pb.title = u"{}{} ({{value}} of {{max_value}})".format(
        stage, u": " + text if text else u"")
    pb.update_progress(value, max(total, 1))


def _part_marks(inst, part):
    set_param(inst, DB.BuiltInParameter.ALL_MODEL_MARK, part.mark or u"")
    set_param(inst, DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
              comments(part))


def _clean_name(name):
    for ch in u'\\:{}[]|;<>?`~':
        name = name.replace(ch, u"_")
    return name.strip() or u"Steel"


def plan_parts(model, matcher):
    """Sort parts into members to match (grouped by section and category),
    shapes and bolts."""
    groups = {}
    shapes = []
    bolts = 0
    for part in model.parts:
        if part.bolt:
            bolts += 1
            continue
        if part.member is None or part.role is None:
            shapes.append(part)
            continue
        use, cat, mode = classify(part.member, part.role)
        key = (cat, part.section)
        g = groups.get(key)
        if g is None:
            g = {"profile": part.member.profile, "parts": [], "materials": {}}
            groups[key] = g
        g["parts"].append((part, use, mode))
        mat = part.material or u""
        g["materials"][mat] = g["materials"].get(mat, 0) + 1
    return groups, shapes, bolts


def main_material(g):
    best = u""
    n = -1
    for k, v in g["materials"].items():
        if v > n:
            best, n = k, v
    return best


def mapping_rows(matcher, unmatched, groups):
    rows = []
    for key in unmatched:
        cat, section = key
        g = groups[key]
        prof = g["profile"]
        cands = matcher.candidates(section, cat, prof)
        options = [e.label for s, e in cands]
        kind = section_kind(prof, section)
        word = (name_key(section) or (u"",))[0]
        default = SHAPE_LABEL
        if kind == "R" or word in ("PL", "FL"):
            default = PLATE_LABEL
        elif cands:
            best_score, best = cands[0]
            if best.kind == kind and best_score < 5.0:
                default = best.label
        rows.append({"section": section, "cat": cat,
                     "text": section_text(prof, section),
                     "count": len(g["parts"]), "options": options,
                     "default": default,
                     "valid": set(matcher.by_label[cat])})
    rows.sort(key=lambda r: (r["cat"], r["section"]))
    return rows


DEFAULT_OPTIONS = {
    "placement": "shared",          # shared / internal / centre
    "natives": True,                # members as native framing and columns
    "ask_map": True,                # ask about sections with no match
    "shape_families": True,         # custom shapes as Generic Model families
    "plates": True,                 # Revit steel plates
    "plate_cuts": True,             # IFC openings as contour cuts
    "plate_holes": True,            # bolt holes in plates and shapes
    "grids_levels": True,           # the file's grids and levels
    "bolts": True,                  # Revit bolts and anchors
    "skip_base_plate_bolts": True,  # leave base plate bolts out
}
PLACEMENTS = {"shared": PLACE_SHARED, "internal": PLACE_INTERNAL,
              "centre": PLACE_CENTRE}


def read_model(path):
    """The steel model in a .ifc, .bsw or a .seed holding one."""
    if path.lower().endswith(".seed"):
        tmp, name = seed_ifc(path)
        try:
            return read_model(tmp)
        finally:
            try:
                os.remove(tmp)
            except Exception:
                pass
    return read_steel_bsw(path) if is_bsw(path) else read_steel(path)


def _choice_label(res, flat):
    """A match result as the label the Section Map shows."""
    if res == PLATE:
        return PLATE_LABEL
    if res == SHAPE:
        return SHAPE_LABEL
    if res is None:
        return PLATE_LABEL if flat else SHAPE_LABEL
    return res.label


def file_section_rows(model, entries, saved):
    """Rows for the Section Map of one file: every member section in it,
    with its type now ("default": the saved choice, else the automatic
    match) and the automatic match on its own ("auto"), so only real
    changes get saved."""
    matcher = Matcher(entries, saved)
    auto = Matcher(entries, {})
    groups, shapes, bolts = plan_parts(model, matcher)
    rows = []
    for key in sorted(groups):
        cat, section = key
        g = groups[key]
        prof = g["profile"]
        mat = main_material(g)
        kind = section_kind(prof, section)
        flat = kind == "R" or (name_key(section) or (u"",))[0] in ("PL", "FL")
        now, how = matcher.match(section, cat, prof, mat)
        guess, guess_how = auto.match(section, cat, prof, mat)
        options = [e.label for sc, e in matcher.candidates(section, cat, prof)]
        rows.append({"section": section, "cat": cat,
                     "text": section_text(prof, section),
                     "count": len(g["parts"]), "options": options,
                     "default": _choice_label(now, flat),
                     "auto": _choice_label(guess, flat),
                     "matched": guess is not None,
                     "how": u"saved" if how == "saved choice" else
                            (u"no match" if guess is None else guess_how),
                     "valid": set(matcher.by_label[cat])})
    return rows


def save_section_edits(rows, choices, path=None):
    """Write the Section Map's choices for one file into the saved map:
    a choice that differs from the automatic match is saved, one that is
    the automatic match drops any saved entry. choices: {(cat, section):
    type label or map value} as _section_map returns them."""
    special = dict(SPECIAL_CHOICES)
    data = load_map(path)
    for row in rows:
        key = (row["cat"], row["section"])
        if key not in choices:
            continue
        value = choices[key]
        auto = special.get(row["auto"], row["auto"])
        shown = special.get(row["default"], row["default"])
        cat_map = data.setdefault(row["cat"], {})
        if value is None or (value == auto and
                             (row["matched"] or row.get("reset"))):
            # back to the automatic match (with none, the import asks again)
            cat_map.pop(row["section"], None)
        elif value != shown or row["how"] == u"saved":
            cat_map[row["section"]] = value
    return replace_map(data, path)


def section_labels(doc):
    """{cat: [type labels]} of every framing/column type a section can be
    mapped to (loaded types and the library catalogues), for the map editor."""
    entries = loaded_entries(doc)
    entries.extend(library_entries(doc.Application))
    m = Matcher(entries, {})
    return dict((cat, sorted(m.by_label[cat])) for cat in m.by_label)


def run(doc, ifc_path, source_name=None, options=None):
    """Import a steel IFC or StruCad .bsw into doc with the options from
    3D Import's window (DEFAULT_OPTIONS for any left out). Returns False
    if cancelled."""
    from pyrevit import forms, script
    opts = dict(DEFAULT_OPTIONS)
    opts.update(options or {})

    source = source_name or os.path.basename(ifc_path)
    strucad = is_bsw(ifc_path)
    what = u"StruCad model" if strucad else u"IFC"
    with forms.ProgressBar(title=u"Reading {}...".format(what)) as pb:
        try:
            reader = read_steel_bsw if strucad else read_steel
            model = reader(ifc_path, progress=lambda i, n:
                           pb.update_progress(i, max(n, 1)))
        except Exception as ex:
            forms.alert(u"Could not read the {}: {}".format(what, ex),
                        title=TITLE)
            return False
    if not model.parts:
        forms.alert("The IFC has no steel parts with geometry.", title=TITLE)
        return False
    if not opts["plate_holes"]:
        for part in model.parts:
            if not part.bolt:
                strip_holes(part)

    placement = PLACEMENTS.get(opts["placement"], PLACE_SHARED)
    mapper = make_mapper(doc, model, placement)


    with forms.ProgressBar(title="Reading family types...") as pb:
        pb.update_progress(1, 3)
        entries = loaded_entries(doc)
        entries.extend(library_entries(doc.Application))
        pb.update_progress(3, 3)
    matcher = Matcher(entries, load_map())

    groups, shape_parts, bolts = plan_parts(model, matcher)
    matched = {}
    unmatched = []
    for key in sorted(groups):
        cat, section = key
        g = groups[key]
        if not opts["natives"]:
            matched[key] = (SHAPE, "members as shapes")
            continue
        res, how = matcher.match(section, cat, g["profile"], main_material(g))
        if res is None:
            unmatched.append(key)
        else:
            matched[key] = (res, how)

    if unmatched and not opts["ask_map"]:
        for key in unmatched:
            matched[key] = (SHAPE, "no match")
        unmatched = []
    if unmatched:
        from Snippets._section_map import ask_section_map
        rows = mapping_rows(matcher, unmatched, groups)
        answer = ask_section_map(rows, SPECIAL_CHOICES)
        if answer is None:
            return False
        choices, remember = answer
        to_save = {}
        for key, value in choices.items():
            cat, section = key
            if value in (SHAPE, PLATE):
                matched[key] = (value, "chosen")
            else:
                matched[key] = (matcher.find(cat, value), "chosen")
            to_save.setdefault(cat, {})[section] = value
        if remember:
            save_map(to_save)

    importer = Importer(doc, model, mapper, source)
    plans = []
    natives_by_key = {}
    as_plate = set()                    # members mapped to "import as plate"
    for key in sorted(groups):
        res, how = matched.get(key, (SHAPE, "shape"))
        cat = key[0]
        for part, use, mode in groups[key]["parts"]:
            if res == PLATE:
                as_plate.add(id(part))
                shape_parts.append(part)
            elif res == SHAPE or res is None:
                shape_parts.append(part)
            else:
                plans.append(Plan(part, use, cat, mode, res, how))
        natives_by_key[key] = (res, how, len(groups[key]["parts"]))

    # flat plates can be Revit's own steel plates instead of shapes
    plate_specs = []
    plate_maker = None
    candidates = [(p, plate_spec(p, id(p) in as_plate)) for p in shape_parts]
    candidates = [(p, s) for p, s in candidates if s is not None]
    kept = []
    for p, s in candidates:
        cuts = plate_cuts(s, p.openings) if p.openings else []
        if cuts is None:                # an opening a contour cut cannot make
            continue
        if cuts and not opts["plate_holes"]:
            cuts = [c for c in cuts if not is_hole(c, s["normal"])]
        if not opts["plate_cuts"]:
            if cuts:                    # uncut it would be wrong; a shape keeps
                continue                # the opening
            cuts = []
        s["cuts"] = cuts
        kept.append((p, s))
    candidates = kept
    if candidates:
        plate_maker = PlateMaker(doc, mapper)
        if plate_maker.api is not None and opts["plates"]:
            plate_specs = candidates
            taken = set(id(p) for p, s in candidates)
            shape_parts = [p for p in shape_parts if id(p) not in taken]

    # bolts can be Revit's own steel bolts
    bolt_specs = []
    bolt_skipped = []
    fasteners = [p for p in model.parts
                 if p.ifc_type in ("IFCMECHANICALFASTENER", "IFCFASTENER")]
    if fasteners:
        if plate_maker is None:
            plate_maker = PlateMaker(doc, mapper)
        if plate_maker.api is not None and opts["bolts"]:
            index = SolidIndex(model.parts, plate_specs)
            bolt_specs, bolt_skipped = plan_bolts(
                fasteners, index, opts["skip_base_plate_bolts"])

    # repeated custom shapes can be one family each, placed at every copy
    shape_groups = []
    if opts["shape_families"]:
        shape_groups = repeated_shapes(shape_parts)

    group = DB.TransactionGroup(doc, TITLE)
    group.Start()
    try:
        # grids and levels first, so members are hosted on the file's levels
        if opts["grids_levels"]:
            from Snippets._grids_levels import make_grids_levels
            importer.grids_levels = make_grids_levels(
                doc, getattr(model, "grids", None), getattr(model, "levels", None),
                lambda x, y, z: _xyz(mapper.point((x, y, z))))
        # load and activate the chosen types
        t = start_transaction(doc, TITLE + " families")
        importer.ensure_levels()
        bad = set()
        for plan in plans:
            e = plan.entry
            if id(e) in bad:
                continue
            sym = load_entry(doc, e)
            if sym is None:
                bad.add(id(e))
                continue
            if not sym.IsActive:
                sym.Activate()
        t.Commit()
        for plan in [p for p in plans if id(p.entry) in bad]:
            importer.fallback.append((plan, u"could not load " + plan.entry.label))
        plans = [p for p in plans if id(p.entry) not in bad]

        wanted = []
        seen = set()
        for plan in plans:
            k = (plan.entry.symbol.Id.ToString(), plan.mode)
            if k not in seen:
                seen.add(k)
                wanted.append((plan.entry.symbol, plan.mode))
        importer.probes = probe_symbols(doc, wanted, importer.levels[0])

        total = len(plans) + len(plate_specs) + len(bolt_specs) + len(shape_parts)
        with forms.ProgressBar(title="Starting... ({value} of {max_value})",
                               cancellable=True) as pb:
            # build order: members, shapes, then plates with their bolts
            # (bolts must share the plates' steel transaction to connect)
            ok = importer.natives(plans, pb, 0, total)
            shape_parts.extend(p.part for p, why in importer.fallback)
            total = len(plans) + len(plate_specs) + len(bolt_specs) + len(shape_parts)
            done = len(plans)
            if ok and shape_groups:
                importer.shape_families = make_shape_families(
                    doc, mapper, shape_groups, start_transaction, _part_marks,
                    lambda i, n, part: show_progress(
                        pb, u"Locating shapes", part.section or part.name,
                        done + i, total))
                placed = importer.shape_families.placed
                shape_parts = [q for q in shape_parts if id(q) not in placed]
            if ok:
                ok = importer.shapes_of(shape_parts, pb, done, total)
            done = total - len(plate_specs) - len(bolt_specs)
            if ok and (plate_specs or bolt_specs):
                importer.plate_maker = plate_maker
                ok = importer.steel_of(plate_specs, bolt_specs, plate_maker, pb,
                                       done, total)
                importer.bolt_failed[:0] = bolt_skipped
                # plates Revit would not make still come in, as shapes
                failed = [p for p, why in importer.plate_failed]
                if ok and failed:
                    ok = importer.shapes_of(failed, pb, total - len(failed), total)
            cancelled = pb.cancelled
        if cancelled:
            group.RollBack()
            return False
        group.Assimilate()
    except Exception:
        if group.HasStarted():
            group.RollBack()
        raise

    report(script.get_output(), ifc_path, model, placement, importer,
           natives_by_key, plans, bolts)
    return True


# ── REPORT ──────────────────────────────────────────────────────────────────

def report(out, path, model, placement, imp, by_key, plans, bolts):
    out.print_md(u"## Steel IFC\n`{}`".format(path))
    out.print_md(u"Schema {}. Placed in: {}.".format(
        safe_text(model.ifc.schema) or u"?", placement))
    natives = sum(imp.made.values())
    bits = [u"{} {}".format(v, k + ("es" if k == "brace" else "s"))
            for k, v in sorted(imp.made.items())]
    out.print_md(u"**{}** native members ({}), **{}** steel plates, **{}** "
                 u"steel bolts, **{}** anchors, **{}** shapes. {} bolt, nut "
                 u"and washer parts were not made as shapes.".format(
                     natives, u", ".join(bits) or u"none", imp.plates,
                     imp.bolts, imp.anchors, imp.shapes, bolts))
    if imp.grids_levels is not None:
        from Snippets._grids_levels import report_md
        for line in report_md(imp.grids_levels):
            out.print_md(line)
    if imp.anchors:
        out.print_md(u"Anchors use Revit's default anchor standard and snap "
                     u"to its stock lengths; check them against the drawings.")
    maker = getattr(imp, "plate_maker", None)
    if maker is not None and (maker.cuts or maker.cuts_failed):
        out.print_md(u"**{}** contour cuts made on plates from the IFC "
                     u"openings{}.".format(maker.cuts, u", {} failed".format(
                         maker.cuts_failed) if maker.cuts_failed else u""))
    sf = imp.shape_families
    if sf is not None:
        out.print_md(u"**{}** repeated custom shapes made as **{}** Generic "
                     u"Model families ({} as solids, {} as meshes), placed "
                     u"**{}** times. Open a family to remodel all its "
                     u"copies at once.".format(
                         sf.families, sf.families, sf.solid, sf.mesh,
                         sf.instances))
        for e in sf.errors[:10]:
            out.print_md(u"- {}".format(e))
    if imp.plates or imp.bolts:
        out.print_md(u"Revit draws steel plates and bolts only at the Fine "
                     u"detail level. Bolts are {} with the IFC grade; their "
                     u"grip comes from the IFC geometry.".format(BOLT_STANDARD))
    if imp.bolt_failed:
        out.print_md(u"### Bolts not made ({})".format(len(imp.bolt_failed)))
        for part, why in imp.bolt_failed[:100]:
            out.print_md(u"- {} {}: {}".format(
                part.name or ifc_class(part.ifc_type), part.assembly, why))
    if imp.plate_failed:
        out.print_md(u"### Plates made as shapes instead ({})".format(
            len(imp.plate_failed)))
        for part, why in imp.plate_failed[:100]:
            out.print_md(u"- {} {}: {}".format(part.section, part.mark, why))

    out.print_md(u"### Sections")
    lines = [u"| Section | Category | Count | Revit type | How |",
             u"|---|---|---|---|---|"]
    for key in sorted(by_key):
        res, how, n = by_key[key]
        cat, section = key
        target = {SHAPE: u"shape", PLATE: u"plate", None: u"shape"}.get(res) \
            if res in (SHAPE, PLATE, None) else res.label
        if res not in (SHAPE, PLATE, None) and not res.loaded:
            target += u" (loaded from library)"
        lines.append(u"| {} | {} | {} | {} | {} |".format(
            section or u"(no name)", cat, n, target, how or u"shape"))
    out.print_md(u"\n".join(lines))

    trimmed = len([p for p in plans if p.trimmed])
    notched = len([p for p in plans if p.notches])
    if trimmed or notched:
        out.print_md(u"End cuts trimmed {} native members to the cut. {} "
                     u"native members have copes or notches that are not "
                     u"modelled on them.".format(trimmed, notched))
    if imp.fallback:
        out.print_md(u"### Made as shapes instead of native ({})".format(
            len(imp.fallback)))
        for plan, why in imp.fallback[:100]:
            out.print_md(u"- {} {}: {}".format(plan.part.section,
                                               plan.part.mark, why))
    if imp.fits:
        out.print_md(u"### Orientation fits over {} mm ({})".format(
            int(REPORT_SCORE_MM), len(imp.fits)))
        out.print_md(u"The type's section does not lay well over the IFC "
                     u"outline; check these.")
        for plan in imp.fits[:100]:
            out.print_md(u"- {} {} -> {}: {:.0f} mm {}".format(
                plan.part.section, plan.part.mark, plan.entry.label,
                plan.score, out.linkify(plan.fi.Id) if plan.fi else u""))
    if imp.unverified:
        out.print_md(u"{} members could not have their rotation checked."
                     .format(imp.unverified))
    if model.unsupported or imp.maker.clip_failures:
        out.print_md(u"### Unsupported geometry")
        for k in sorted(model.unsupported):
            out.print_md(u"- {} ({})".format(k, model.unsupported[k]))
        if imp.maker.clip_failures:
            out.print_md(u"- cuts that failed and were left off ({})".format(
                imp.maker.clip_failures))
    if imp.failed:
        out.print_md(u"### Not imported ({})".format(len(imp.failed)))
        for name, why in imp.failed[:200]:
            out.print_md(u"- {}: {}".format(name, why))
