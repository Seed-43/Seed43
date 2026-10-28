# -*- coding: utf-8 -*-
"""CPython tests for the StruCad .bsw reader.

    python3 tests/test_bsw_steel.py

Writes a small synthetic .bsw (no client files): a hollow member, a member
with a cut end, a plate with a bolt hole, a 2 x 2 bolt group, a dummy
member with no geometry, a grid line and one unreadable record. Reads it
with _bsw_steel and checks the parts _steel_import gets. When the TACH
sample models are on this PC (Reference info/BSW in Proton Drive) they are
read too. Nothing here needs Revit.
"""
import glob
import math
import os
import struct
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "lib"))

from Snippets import _bsw_steel as bsw            # noqa: E402
from Snippets import _ifc_steel as steel          # noqa: E402
from Snippets import _steel_import as imp         # noqa: E402

SAMPLES = os.path.join(os.path.expanduser("~"), "Proton Drive", "juvenciodas",
                       "My files", ".Claude", "Reference info", "BSW")


# ── SYNTHETIC BSW ───────────────────────────────────────────────────────────

def _s(text):
    raw = text.encode("latin-1")
    return struct.pack("<H", len(raw)) + raw


class Geometry(object):
    """One record's geometry: vertices, edges and faces built from vertex
    cycles, written in the order StruCad writes them."""

    def __init__(self, origin, x, y):
        self.head = b"O" + struct.pack("<3f", *origin) + \
            b"X" + struct.pack("<3f", *x) + b"Y" + struct.pack("<3f", *y)
        self.verts = []
        self.edges = []         # (a, b, arc)
        self.faces = []         # (normal, [[signed edge numbers]])
        self.tail = b""

    def vert(self, p):
        self.verts.append(p)
        return len(self.verts) - 1

    def ring(self, cycle):
        out = []
        for i in range(len(cycle)):
            a, b = cycle[i], cycle[(i + 1) % len(cycle)]
            for k, e in enumerate(self.edges):
                if e[:2] == (a, b):
                    out.append(k + 1)
                    break
                if e[:2] == (b, a):
                    out.append(-(k + 1))
                    break
            else:
                self.edges.append((a, b, None))
                out.append(len(self.edges))
        return out

    def hole_ring(self, centre, rad):
        v = self.vert((centre[0] + rad, centre[1], centre[2]))
        self.edges.append((v, v, (centre, (0.0, 0.0, 1.0), rad)))
        return [len(self.edges)]

    def face(self, normal, *loops):
        self.faces.append((normal, list(loops)))

    def prism(self, outer, z0, z1, top=None, holes=()):
        """outer: anticlockwise 2D loop; top(x, y) gives the top z when the
        end is cut; holes: (cx, cy, r) round holes through it."""
        n = len(outer)
        bot = [self.vert((x, y, z0)) for x, y in outer]
        tip = [self.vert((x, y, top(x, y) if top else z1)) for x, y in outer]
        rings_b = [self.hole_ring((cx, cy, z0), r) for cx, cy, r in holes]
        rings_t = [self.hole_ring((cx, cy, z1), r) for cx, cy, r in holes]
        self.face((0.0, 0.0, -1.0), self.ring(list(reversed(bot))), *rings_b)
        tn = (0.0, 0.0, 1.0)
        if top:
            a = [self.verts[i] for i in tip[:3]]
            u = [a[1][k] - a[0][k] for k in range(3)]
            w = [a[2][k] - a[0][k] for k in range(3)]
            c = (u[1] * w[2] - u[2] * w[1], u[2] * w[0] - u[0] * w[2],
                 u[0] * w[1] - u[1] * w[0])
            l = math.sqrt(sum(q * q for q in c))
            tn = tuple(q / l for q in c)
        self.face(tn, self.ring(tip), *rings_t)
        for i in range(n):
            j = (i + 1) % n
            (x0, y0), (x1, y1) = outer[i], outer[j]
            d = math.hypot(x1 - x0, y1 - y0)
            self.face(((y1 - y0) / d, -(x1 - x0) / d, 0.0),
                      self.ring([bot[i], bot[j], tip[j], tip[i]]))
        for cx, cy, r in holes:
            self.tail += b"C" + struct.pack("<3d", cx, cy, (z0 + z1) / 2.0) + \
                struct.pack("<3f", 0.0, 0.0, 1.0) + struct.pack("<ff", r, 0.0)

    def tube(self, w, h, t, z0, z1):
        """A rectangular hollow section: outer and inner walls, two ends
        with an inner loop each."""
        def rect(a, b):
            return [(-a / 2, -b / 2), (a / 2, -b / 2), (a / 2, b / 2),
                    (-a / 2, b / 2)]
        o, i = rect(w, h), rect(w - 2 * t, h - 2 * t)
        ob = [self.vert((x, y, z0)) for x, y in o]
        ot = [self.vert((x, y, z1)) for x, y in o]
        ib = [self.vert((x, y, z0)) for x, y in i]
        it = [self.vert((x, y, z1)) for x, y in i]
        self.face((0.0, 0.0, -1.0), self.ring(list(reversed(ob))),
                  self.ring(ib))
        self.face((0.0, 0.0, 1.0), self.ring(ot), self.ring(list(reversed(it))))
        for k in range(4):
            m = (k + 1) % 4
            (x0, y0), (x1, y1) = o[k], o[m]
            d = math.hypot(x1 - x0, y1 - y0)
            self.face(((y1 - y0) / d, -(x1 - x0) / d, 0.0),
                      self.ring([ob[k], ob[m], ot[m], ot[k]]))
            self.face((-(y1 - y0) / d, (x1 - x0) / d, 0.0),
                      self.ring([ib[m], ib[k], it[k], it[m]]))

    def box(self, x0, y0, z0, x1, y1, z1):
        self.prism([(x0, y0), (x1, y0), (x1, y1), (x0, y1)], z0, z1)

    def bytes(self):
        out = [self.head]
        for p in self.verts:
            out.append(b"V" + struct.pack("<3d", *p))
        for a, b, arc in self.edges:
            if arc is None:
                out.append(b"S" + struct.pack("<hh", a + 1, b + 1))
            else:
                c, axis, r = arc
                out.append(b"E" + struct.pack("<3d", *c) +
                           struct.pack("<3f", *axis) +
                           struct.pack("<fhh", r, a + 1, b + 1))
        for n, loops in self.faces:
            out.append(b"P" + struct.pack("<3f", *n))
            for loop in loops:
                out.append(b"L" + struct.pack("<{}h".format(len(loop) + 1),
                                              *(loop + [0])))
        out.append(self.tail)
        return b"".join(out)


def member(section, mark, grade, attrs, geo):
    out = b"BM" + struct.pack("<i", 7) + _s(section) + _s(mark) + _s(grade)
    for k, v in attrs:
        out += b"A" + _s(k) + _s(v)
    return out + (geo or Geometry((0.0, 0.0, 0.0), (1.0, 0.0, 0.0),
                                  (0.0, 1.0, 0.0))).bytes()


def plate(section, mark, thickness, geo):
    return b"BP" + struct.pack("<ii", 7, 0) + _s(section) + _s(mark) + \
        _s(u"300PLUS") + struct.pack("<f", thickness) + geo.bytes()


def bolt(size, length, name, pattern, geo):
    return b"BB" + struct.pack("<ii", 7, 0) + \
        struct.pack("<fff", size, 0.0, length) + _s(name) + _s(u"") + \
        _s(u"") + struct.pack("<hhff", *pattern) + geo.bytes()


def write_sample(path):
    parts = [struct.pack("<hid", 4, 0, 16.0)]
    # 1. a 150 x 100 x 6 RHS column, 3 m tall, standing at (1000, 0, 0)
    g = Geometry((1000.0, 0.0, 0.0), (1.0, 0.0, 0.0), (0.0, 1.0, 0.0))
    g.tube(150.0, 100.0, 6.0, 0.0, 3000.0)
    parts.append(member(u"150X100X6.0RHS", u"C1", u"C350L0",
                        [(u"PHASE", u"1"), (u"DESC", u"COLUMN")], g))
    # 2. a 200 x 100 flat-ish member lying along world X, its top end cut
    g = Geometry((0.0, 0.0, 3000.0), (0.0, 1.0, 0.0), (0.0, 0.0, 1.0))
    g.prism([(-50.0, -100.0), (50.0, -100.0), (50.0, 100.0), (-50.0, 100.0)],
            0.0, 2000.0, top=lambda x, y: 2000.0 + 0.2 * x)
    parts.append(member(u"200X100RHS", u"B1", u"C350L0", [], g))
    # 3. a dummy member with no geometry
    parts.append(member(u"20X1.6SHS", u"D1", u"", [(u"LAYER", u"DUMMY")], None))
    # 4. junk StruCad cannot have written: skipped with a warning
    parts.append(b"BQ" + b"\x00\x01junk")
    # 5. a 300 x 200 x 12 plate with one 18 mm hole, lying flat
    g = Geometry((0.0, 0.0, 500.0), (1.0, 0.0, 0.0), (0.0, 1.0, 0.0))
    g.prism([(0.0, 0.0), (300.0, 0.0), (300.0, 200.0), (0.0, 200.0)],
            0.0, 12.0, holes=[(150.0, 100.0, 9.0)])
    parts.append(plate(u"PL12", u"P1", 12.0, g))
    # 6. a 2 x 2 group of M16 x 40 bolts, local Z down through the plate
    g = Geometry((150.0, 100.0, 540.0), (1.0, 0.0, 0.0), (0.0, -1.0, 0.0))
    g.box(-12.0, -12.0, -17.0, 12.0, 12.0, 57.0)
    parts.append(bolt(16.0, 40.0, u"8.8/S", (2, 2, 70.0, 80.0), g))
    # 7. a grid line, then the viewer trailer
    parts.append(b"g" + _s(u"A") + struct.pack("<6f", 0, -5000, 0,
                                               0, 5000, 0))
    parts.append(b"v")
    with open(path, "wb") as f:
        f.write(b"".join(parts))


# ── TESTS ───────────────────────────────────────────────────────────────────

class SyntheticBsw(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        fd, cls.path = tempfile.mkstemp(suffix=".bsw")
        os.close(fd)
        write_sample(cls.path)
        cls.raw = bsw.read_bsw(cls.path)
        cls.model = bsw.to_steel_model(cls.raw)
        cls.by_mark = {}
        for p in cls.model.parts:
            cls.by_mark.setdefault(p.mark, []).append(p)

    @classmethod
    def tearDownClass(cls):
        os.remove(cls.path)

    def test_is_bsw(self):
        self.assertTrue(bsw.is_bsw(self.path))
        self.assertFalse(bsw.is_bsw(__file__))

    def test_records(self):
        kinds = [e.kind for e in self.raw.entities]
        self.assertEqual(kinds.count("member"), 3)
        self.assertEqual(kinds.count("plate"), 1)
        self.assertEqual(kinds.count("bolt"), 4)       # 2 x 2 from one record
        self.assertEqual(len(self.raw.grids), 1)
        self.assertEqual(self.raw.grids[0].label, u"A")
        self.assertEqual(len(self.raw.warnings), 1)
        self.assertEqual(self.raw.warnings[0][1], "record")
        c1 = [e for e in self.raw.entities if e.mark == u"C1"][0]
        self.assertEqual(c1.attrs, {u"phase": u"1", u"desc": u"COLUMN"})
        self.assertAlmostEqual(c1.length, 3000.0)

    def test_dummy_and_unsupported(self):
        self.assertNotIn(u"D1", self.by_mark)
        self.assertEqual(self.model.unsupported,
                         {"member with no geometry": 1})
        self.assertEqual(self.model.ifc.schema, u"StruCad BSW")

    def test_hollow_column(self):
        p = self.by_mark[u"C1"][0]
        self.assertEqual(p.ifc_type, "IFCCOLUMN")
        self.assertEqual(p.material, u"C350L0")
        self.assertIsInstance(p.solids[0], steel.Extrusion)
        m = p.member
        self.assertEqual(len(m.profile.loops), 2)
        self.assertEqual(tuple(round(v, 3) for v in m.profile.extent),
                         (150.0, 100.0))
        self.assertEqual(tuple(round(v, 3) for v in m.start),
                         (1000.0, 0.0, 0.0))
        self.assertEqual(tuple(round(v, 3) for v in m.end),
                         (1000.0, 0.0, 3000.0))
        self.assertEqual(imp.classify(m, p.role)[1], "column")

    def test_cut_member(self):
        p = self.by_mark[u"B1"][0]
        self.assertEqual(p.ifc_type, "IFCBEAM")      # lies level
        self.assertIsInstance(p.solids[0], steel.Brep)
        self.assertEqual(p.solids[0].clips, [])
        m = p.member
        self.assertEqual(tuple(round(v, 3) for v in m.profile.extent),
                         (100.0, 200.0))
        self.assertEqual(tuple(round(v, 3) for v in m.start),
                         (0.0, 0.0, 3000.0))
        self.assertAlmostEqual(m.length, 2010.0, places=3)
        for k in range(3):
            self.assertAlmostEqual(m.direction[k], (1.0, 0.0, 0.0)[k])

    def test_plate_and_hole(self):
        p = self.by_mark[u"P1"][0]
        self.assertEqual(p.ifc_type, "IFCPLATE")
        spec = imp.plate_spec(p)
        self.assertIsNotNone(spec)
        self.assertAlmostEqual(spec["thickness"], 12.0)
        self.assertEqual(len(p.openings), 1)         # the C hole, once
        hole = p.openings[0]
        self.assertEqual(hole.profile.kind, "C")
        self.assertAlmostEqual(hole.profile.dims["diameter"], 18.0)
        cuts = imp.plate_cuts(spec, p.openings)
        self.assertEqual(len(cuts), 1)

    def test_bolts(self):
        bolts = [p for p in self.model.parts if p.bolt]
        self.assertEqual(len(bolts), 4)
        size, grade, site = imp.bolt_size(bolts[0].props)
        self.assertEqual(size, 16.0)
        self.assertEqual(imp.nominal_length(bolts[0].props), 40.0)
        heads = sorted((round(b.solids[0].plane[0][0], 3),
                        round(b.solids[0].plane[0][1], 3)) for b in bolts)
        self.assertEqual(heads, [(150.0, 20.0), (150.0, 100.0),
                                 (220.0, 20.0), (220.0, 100.0)])
        point, d = imp.bolt_axes(bolts[0])[0]
        self.assertAlmostEqual(abs(d[2]), 1.0)


@unittest.skipUnless(glob.glob(os.path.join(SAMPLES, "*.bsw")),
                     "no sample .bsw files on this PC")
class SampleModels(unittest.TestCase):
    def test_samples_read(self):
        for path in sorted(glob.glob(os.path.join(SAMPLES, "*.bsw"))):
            raw = bsw.read_bsw(path)
            self.assertEqual(raw.warnings, [], path)
            model = bsw.to_steel_model(raw)
            members = [p for p in model.parts if p.role is not None]
            lined = [p for p in members if p.member is not None]
            self.assertGreater(len(lined), 0.99 * len(members), path)
            # every vertex of a member lies between its line's ends
            for p in lined[:300]:
                ent = raw.entities[p.eid - 1]
                d = p.member.direction
                for v in ent.verts:
                    w = ent.world(v)
                    t = sum((w[k] - p.member.start[k]) * d[k] for k in range(3))
                    self.assertGreater(t, -0.5)
                    self.assertLess(t, p.member.length + 0.5)


if __name__ == "__main__":
    unittest.main(verbosity=2)
