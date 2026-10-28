# -*- coding: utf-8 -*-
"""CPython tests for the Steel IFC reader and the importer's maths.

    python3 tests/test_steel_ifc.py

Writes a small synthetic IFC2x3 (no client files), reads it with
_ifc_steel.read_steel, and checks the matching and orientation maths of
_steel_import against a simulated Revit section frame. Nothing here needs
Revit.
"""
import math
import os
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "lib"))

from Snippets import _ifc_steel as steel          # noqa: E402
from Snippets import _steel_import as imp         # noqa: E402


# ── SYNTHETIC IFC ───────────────────────────────────────────────────────────

class IfcWriter(object):
    def __init__(self):
        self.lines = []
        self.n = 0

    def add(self, text):
        self.n += 1
        self.lines.append("#{}={};".format(self.n, text))
        return "#{}".format(self.n)

    def pt(self, *c):
        return self.add("IFCCARTESIANPOINT(({}))".format(
            ",".join("{:.6f}".format(float(v)) for v in c)))

    def dir(self, *c):
        return self.add("IFCDIRECTION(({}))".format(
            ",".join("{:.6f}".format(float(v)) for v in c)))

    def a3(self, loc, z=None, x=None):
        return self.add("IFCAXIS2PLACEMENT3D({},{},{})".format(
            self.pt(*loc), self.dir(*z) if z else "$",
            self.dir(*x) if x else "$"))

    def a2(self):
        return self.add("IFCAXIS2PLACEMENT2D({},$)".format(self.pt(0, 0)))

    def text(self, header):
        return ("ISO-10303-21;\nHEADER;\nFILE_DESCRIPTION((''),'2;1');\n"
                "FILE_NAME('t.ifc','',(''),(''),'','','');\n"
                "FILE_SCHEMA(('IFC2X3'));\nENDSEC;\nDATA;\n" +
                "\n".join(self.lines) + "\nENDSEC;\nEND-ISO-10303-21;\n")


def build_ifc():
    w = IfcWriter()
    ids = {}
    unit = w.add("IFCSIUNIT(*,.LENGTHUNIT.,.MILLI.,.METRE.)")
    ang = w.add("IFCSIUNIT(*,.PLANEANGLEUNIT.,$,.RADIAN.)")
    ua = w.add("IFCUNITASSIGNMENT(({},{}))".format(unit, ang))
    ctx = w.add("IFCGEOMETRICREPRESENTATIONCONTEXT($,'Model',3,1.E-05,{},$)"
                .format(w.a3((0, 0, 0))))
    w.add("IFCPROJECT('p',$,'P',$,$,$,$,({}),{})".format(ctx, ua))
    world = w.add("IFCLOCALPLACEMENT($,{})".format(w.a3((0, 0, 0))))
    storey = w.add("IFCBUILDINGSTOREY('s',$,'L1',$,$,{},$,$,.ELEMENT.,0.)"
                   .format(world))
    contained = []
    steel_mat = w.add("IFCMATERIAL('C350L0')")
    mats = []

    def product(cls, name, objtype, items, place=world, contain=True,
                guid=None):
        rep = w.add("IFCSHAPEREPRESENTATION({},'Body','SweptSolid',({}))"
                    .format(ctx, ",".join(items)))
        pds = w.add("IFCPRODUCTDEFINITIONSHAPE($,$,({}))".format(rep))
        tail = ",$" if cls in ("IFCBEAM", "IFCCOLUMN", "IFCMEMBER",
                               "IFCPLATE") else ""
        if cls == "IFCMECHANICALFASTENER":
            tail = ",$,$"
        e = w.add("{}('{}',$,'{}',$,'{}',{},{},$ {})".format(
            cls, guid or name, name, objtype, place, pds, tail))
        if contain:
            contained.append(e)
        mats.append(e)
        return e

    def extrude(profile, loc, z, x, depth):
        return w.add("IFCEXTRUDEDAREASOLID({},{},{},{:.6f})".format(
            profile, w.a3(loc, z, x), w.dir(0, 0, 1), depth))

    # 1. UB along +X at z=3000, depth up (profile y = world Z)
    ub = w.add("IFCISHAPEPROFILEDEF(.AREA.,'UB180*18',{},90.,179.,6.,10.,8.9)"
               .format(w.a2()))
    ids["ub"] = product("IFCBEAM", "BEAM", "UB180*18",
                        [extrude(ub, (0, 0, 3000), (1, 0, 0), (0, 1, 0), 6000)])

    # 2. PFC column, vertical, from (5000, 0, 0) up 3000
    pfc = w.add("IFCUSHAPEPROFILEDEF(.AREA.,'PFC180*75',{},180.,75.,6.,9.5,"
                "$,$,$,$)".format(w.a2()))
    ids["pfc"] = product("IFCCOLUMN", "COLUMN", "PFC180*75",
                         [extrude(pfc, (5000, 0, 0), (0, 0, 1), (1, 0, 0), 3000)])

    # 3. EA brace, diagonal in the XZ plane
    ea = w.add("IFCLSHAPEPROFILEDEF(.AREA.,'EA75*75*5',{},75.,75.,5.,$,$,$,$,$)"
               .format(w.a2()))
    d = (1 / math.sqrt(2), 0, 1 / math.sqrt(2))
    ids["ea"] = product("IFCMEMBER", "BRACE", "EA75*75*5",
                        [extrude(ea, (0, 0, 0), d, (0, 1, 0), 4242.64)])

    # 4. SHS column leaning 30 degrees
    shs = w.add("IFCRECTANGLEHOLLOWPROFILEDEF(.AREA.,'SHS100*100*6.0',{},"
                "100.,100.,6.,$,$)".format(w.a2()))
    t30 = (math.sin(math.radians(30)), 0, math.cos(math.radians(30)))
    ids["shs"] = product("IFCCOLUMN", "COLUMN", "SHS100*100*6.0",
                         [extrude(shs, (8000, 0, 0), t30, (0, 1, 0), 3000)])

    # 5. UB with an end cut: keep x < 5800 (the half-space is x > 5800)
    plane = w.add("IFCPLANE({})".format(w.a3((5800, 2000, 0), (1, 0, 0))))
    hs = w.add("IFCHALFSPACESOLID({},.F.)".format(plane))
    body = extrude(ub, (0, 2000, 3000), (1, 0, 0), (0, 1, 0), 6000)
    clip = w.add("IFCBOOLEANCLIPPINGRESULT(.DIFFERENCE.,{},{})".format(body, hs))
    ids["clipped"] = product("IFCBEAM", "BEAM", "UB180*18", [clip])

    # 6. Plate with a hole
    outer = w.add("IFCPOLYLINE(({},{},{},{},{}))".format(
        w.pt(0, 0), w.pt(200, 0), w.pt(200, 300), w.pt(0, 300), w.pt(0, 0)))
    hole = w.add("IFCPOLYLINE(({},{},{},{},{}))".format(
        w.pt(50, 50), w.pt(50, 100), w.pt(100, 100), w.pt(100, 50),
        w.pt(50, 50)))
    plp = w.add("IFCARBITRARYPROFILEDEFWITHVOIDS(.AREA.,'PL10',{},({}))"
                .format(outer, hole))
    ids["plate"] = product("IFCPLATE", "PLATE", "PL10",
                           [extrude(plp, (0, 5000, 0), None, None, 10)])

    # 7. Bent part as a faceted brep (a unit box, 100 mm)
    P = [w.pt(x, y, z) for z in (0, 100) for y in (0, 100) for x in (0, 100)]

    def face(*k):
        loop = w.add("IFCPOLYLOOP(({}))".format(",".join(P[i] for i in k)))
        b = w.add("IFCFACEOUTERBOUND({},.T.)".format(loop))
        return w.add("IFCFACE(({}))".format(b))
    faces = [face(0, 2, 3, 1), face(4, 5, 7, 6), face(0, 1, 5, 4),
             face(2, 6, 7, 3), face(0, 4, 6, 2), face(1, 3, 7, 5)]
    shell = w.add("IFCCLOSEDSHELL(({}))".format(",".join(faces)))
    brep = w.add("IFCFACETEDBREP({})".format(shell))
    ids["brep"] = product("IFCBEAM", "RAIL", "D20", [brep])

    # 8. Bolt: a rod under a mechanical fastener
    rod = w.add("IFCCIRCLEPROFILEDEF(.AREA.,'M16',{},8.)".format(w.a2()))
    fast = product("IFCMECHANICALFASTENER", "BOLT", "M16",
                   [extrude(rod, (0, 0, 0), None, None, 60)])
    shank = product("IFCBEAM", "SHANK", "TR(8.8)16",
                    [extrude(rod, (0, 0, 0), None, None, 50)], contain=False)
    w.add("IFCRELAGGREGATES('a1',$,$,$,{},({}))".format(fast, shank))
    ids["bolt"] = fast

    # 9. ProStructures-style beam: section in Name, outline profile
    x, y, tf, tw = 45.0, 89.5, 10.0, 3.0
    opts = [(-x, -y), (x, -y), (x, -y + tf), (tw, -y + tf), (tw, y - tf),
            (x, y - tf), (x, y), (-x, y), (-x, y - tf), (-tw, y - tf),
            (-tw, -y + tf), (-x, -y + tf), (-x, -y)]
    oline = w.add("IFCPOLYLINE(({}))".format(",".join(w.pt(a, b)
                                                      for a, b in opts)))
    arb = w.add("IFCARBITRARYCLOSEDPROFILEDEF(.AREA.,$,{})".format(oline))
    ids["pro"] = product("IFCBEAM", "180UB18.1", "Beam",
                         [extrude(arb, (0, 8000, 3000), (0, 1, 0), (-1, 0, 0),
                                  4000)])

    # 10. Tekla nut exported as a beam under an assembly
    hexp = w.add("IFCCIRCLEPROFILEDEF(.AREA.,'M16_NUT',{},12.)".format(w.a2()))
    nut = product("IFCBEAM", "NUT", "M16_NUT",
                  [extrude(hexp, (0, 0, 0), None, None, 14)], contain=False)
    asm = w.add("IFCELEMENTASSEMBLY('asm',$,'BOLT ASSEMBLY',$,$,{},$,$,"
                ".NOTDEFINED.,.NOTDEFINED.)".format(world))
    contained.append(asm)
    w.add("IFCRELAGGREGATES('a2',$,$,$,{},({}))".format(asm, nut))
    ids["nut"] = nut

    # marks and material
    mark = w.add("IFCPROPERTYSINGLEVALUE('Assembly/Cast unit Mark',$,"
                 "IFCLABEL('B1'),$)")
    pset = w.add("IFCPROPERTYSET('ps',$,'Tekla Assembly',$,({}))".format(mark))
    w.add("IFCRELDEFINESBYPROPERTIES('r',$,$,$,({}),{})".format(ids["ub"], pset))
    w.add("IFCRELASSOCIATESMATERIAL('m',$,$,$,({}),{})".format(
        ",".join(mats), steel_mat))
    w.add("IFCRELCONTAINEDINSPATIALSTRUCTURE('c',$,$,$,({}),{})".format(
        ",".join(contained), storey))
    return w.text(None), ids


def close(a, b, eps=1e-3):
    return all(abs(x - y) < eps for x, y in zip(a, b))


# ── READER ──────────────────────────────────────────────────────────────────

class ReaderTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        text, cls.ids = build_ifc()
        fd, cls.path = tempfile.mkstemp(suffix=".ifc")
        with os.fdopen(fd, "w") as f:
            f.write(text)
        cls.model = steel.read_steel(cls.path)
        cls.parts = dict((p.eid, p) for p in cls.model.parts)

    @classmethod
    def tearDownClass(cls):
        os.remove(cls.path)

    def part(self, key):
        return self.parts[int(self.ids[key][1:])]

    def test_counts_and_roles(self):
        self.assertEqual(self.model.ifc.schema, "IFC2X3")
        self.assertEqual(self.part("ub").role, "beam")
        self.assertEqual(self.part("pfc").role, "column")
        self.assertEqual(self.part("ea").role, "brace")
        self.assertIsNone(self.part("plate").role)

    def test_bolts(self):
        self.assertTrue(self.part("bolt").bolt)
        self.assertTrue(self.part("nut").bolt)
        shank = [p for p in self.model.parts if p.name == "SHANK"]
        self.assertTrue(shank and shank[0].bolt)
        self.assertFalse(self.part("ub").bolt)

    def test_ub_member(self):
        p = self.part("ub")
        m = p.member
        self.assertEqual(m.profile.kind, "I")
        self.assertAlmostEqual(m.profile.dims["depth"], 179.0)
        self.assertTrue(close(m.start, (0, 0, 3000)))
        self.assertTrue(close(m.end, (6000, 0, 3000)))
        self.assertTrue(close(m.x, (0, 1, 0)))
        self.assertTrue(close(m.y, (0, 0, 1)))
        self.assertEqual(p.mark, "B1")
        self.assertEqual(p.material, "C350L0")
        self.assertEqual(p.section, "UB180*18")

    def test_channel_centroid(self):
        m = self.part("pfc").member
        self.assertEqual(m.profile.kind, "U")
        self.assertTrue(close(m.profile.extent, (75, 180)))
        self.assertLess(m.profile.centroid[0], -10)    # web side is -x
        self.assertAlmostEqual(m.profile.centroid[1], 0.0, 6)
        self.assertTrue(close(m.direction, (0, 0, 1)))

    def test_clip(self):
        p = self.part("clipped")
        self.assertEqual(len(p.solids[0].clips), 1)
        c = p.solids[0].clips[0]
        self.assertEqual(c["kind"], "plane")
        self.assertTrue(close(c["keep"], (-1, 0, 0)))
        s, e, trimmed, others = imp.clip_line(p.member.start, p.member.end,
                                              p.solids[0].clips)
        self.assertTrue(trimmed)
        self.assertEqual(others, 0)
        self.assertTrue(close(e, (5800, 2000, 3000)))

    def test_plate_and_brep(self):
        pl = self.part("plate")
        self.assertEqual(len(pl.solids[0].profile.loops), 2)
        br = self.part("brep")
        self.assertIsInstance(br.solids[0], steel.Brep)
        self.assertEqual(len(br.solids[0].faces), 6)
        self.assertIsNone(br.member)

    def test_prostructures_name(self):
        p = self.part("pro")
        self.assertEqual(p.member.profile.kind, "ARB")
        self.assertEqual(p.section, "180UB18.1")
        self.assertEqual(steel.name_key(p.section), steel.name_key("UB180*18"))


# ── CATALOGUE AND MATCHING ──────────────────────────────────────────────────

CATALOGUE = (u",Width##SECTION_PROPERTY##MILLIMETERS,Height##SECTION_PROPERTY"
             u"##MILLIMETERS,Flange Thickness##SECTION_PROPERTY##MILLIMETERS,"
             u"Web Thickness##SECTION_PROPERTY##MILLIMETERS,Section Area##"
             u"SECTION_PROPERTY##SQUARE_MILLIMETERS,Section Name Key##OTHER##\n"
             u"180UB16.1,90,173,7,4.5,2040,180UB16.1\n"
             u"180UB18.1,90,179,10,6,2300,180UB18.1\n"
             u"\"200UB, 22.3\",133,202,7,5,2860,200UB22.3\n")


def entries():
    out = []
    for fam in (u"UB-Universal Beam", u"UB-Universal Beam-C350L0",
                u"UB-Universal Beam-C450"):
        for tname, dims, key in imp.parse_catalogue(CATALOGUE):
            out.append(imp.Entry(fam, tname, dims, False, path=u"x.rfa",
                                 section_key=key))
    out.append(imp.Entry(u"PFC-Parallel Flange Channel", u"180PFC",
                         {"Height": 180, "Width": 75, "Flange Thickness": 9.5,
                          "Web Thickness": 6}, False, path=u"y.rfa"))
    out.append(imp.Entry(u"PFC-Parallel Flange Channel-Column", u"180PFC",
                         {"d": 180, "bf": 75, "tf": 9.5, "tw": 6}, True,
                         path=u"y.rfa"))
    out.append(imp.Entry(u"SHS-Square Hollow Section", u"100x100x6.0SHS",
                         {"Height": 100, "Width": 100,
                          "Wall Nominal Thickness": 6}, False, path=u"z.rfa"))
    return out


class MatchTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        text, ids = build_ifc()
        fd, path = tempfile.mkstemp(suffix=".ifc")
        with os.fdopen(fd, "w") as f:
            f.write(text)
        cls.model = steel.read_steel(path)
        os.remove(path)
        cls.by_name = {}
        for p in cls.model.parts:
            if p.member is not None:
                cls.by_name.setdefault(p.section, p)

    def test_parse_catalogue(self):
        rows = imp.parse_catalogue(CATALOGUE)
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[2][0], u"200UB, 22.3")
        self.assertEqual(rows[1][1]["Height"], 179.0)
        self.assertNotIn("Section Area", rows[1][1])
        self.assertEqual(rows[1][2], u"180UB18.1")
        utf16 = CATALOGUE.encode("utf-16")
        self.assertEqual(imp.parse_catalogue(imp.decode_catalogue(utf16)), rows)

    def test_family_kind(self):
        self.assertEqual(imp.family_kind(u"UB-Universal Beam"), "I")
        self.assertEqual(imp.family_kind(u"PFC-Parallel Flange Channel"), "U")
        self.assertEqual(imp.family_kind(u"EA-Equal Angle"), "L")
        self.assertEqual(imp.family_kind(u"CHS-Circular Hollow Section"), "CH")
        self.assertEqual(imp.family_kind(u"Round Bars"), "C")
        self.assertEqual(imp.family_kind(u"Flat Bars"), "R")
        self.assertTrue(imp.is_column_library(
            u"C:\\Lib\\34-11.1 Columns\\Steel\\UC.rfa", u"UC"))
        self.assertTrue(imp.is_column_library(u"x", u"UC-Universal Column-Column"))

    def test_dims_match_prefers_grade(self):
        m = imp.Matcher(entries())
        p = self.by_name["UB180*18"]
        e, how = m.match(p.section, "framing", p.member.profile, u"C350L0")
        self.assertEqual(how, "dimensions")
        self.assertEqual(e.label, u"UB-Universal Beam-C350L0: 180UB18.1")
        e, how = m.match(p.section, "framing", p.member.profile, u"")
        self.assertEqual(e.label, u"UB-Universal Beam: 180UB18.1")

    def test_loaded_preferred(self):
        es = entries()
        es.append(imp.Entry(u"UB-Universal Beam-C450", u"180UB18.1",
                            {"Height": 179, "Width": 90, "tf": 10, "tw": 6},
                            False, symbol=object()))
        m = imp.Matcher(es)
        p = self.by_name["UB180*18"]
        e, how = m.match(p.section, "framing", p.member.profile, u"C350L0")
        self.assertTrue(e.loaded)
        self.assertEqual(len([x for x in m.entries["framing"]
                              if x.label == e.label]), 1)

    def test_name_match_outline(self):
        m = imp.Matcher(entries())
        p = self.by_name["180UB18.1"]
        e, how = m.match(p.section, "framing", p.member.profile, u"")
        self.assertEqual(how, "name")
        self.assertEqual(e.type_name, u"180UB18.1")

    def test_column_category(self):
        m = imp.Matcher(entries())
        p = self.by_name["PFC180*75"]
        e, how = m.match(p.section, "column", p.member.profile)
        self.assertTrue(e.column)
        e, how = m.match(p.section, "framing", p.member.profile)
        self.assertFalse(e.column)

    def test_unmatched_and_saved(self):
        m = imp.Matcher(entries())
        p = self.by_name["EA75*75*5"]
        e, how = m.match(p.section, "framing", p.member.profile)
        self.assertIsNone(e)
        cands = m.candidates(p.section, "framing", p.member.profile)
        self.assertTrue(cands)
        m2 = imp.Matcher(entries(), {"framing": {u"EA75*75*5": imp.SHAPE}})
        self.assertEqual(m2.match(p.section, "framing", p.member.profile)[0],
                         imp.SHAPE)
        label = u"SHS-Square Hollow Section: 100x100x6.0SHS"
        m3 = imp.Matcher(entries(), {"framing": {u"EA75*75*5": label}})
        self.assertEqual(m3.match(p.section, "framing",
                                  p.member.profile)[0].label, label)

    def test_shs_dims(self):
        m = imp.Matcher(entries())
        p = self.by_name["SHS100*100*6.0"]
        e, how = m.match(p.section, "framing", p.member.profile)
        self.assertEqual(e.type_name, u"100x100x6.0SHS")

    def test_map_roundtrip(self):
        fd, path = tempfile.mkstemp(suffix=".json")
        os.close(fd)
        os.remove(path)
        try:
            self.assertTrue(imp.save_map({"framing": {u"EA75*75*5": imp.SHAPE}},
                                         path))
            imp.save_map({"column": {u"X": u"F: T"}}, path)
            data = imp.load_map(path)
            self.assertEqual(data["framing"][u"EA75*75*5"], imp.SHAPE)
            self.assertEqual(data["column"][u"X"], u"F: T")
        finally:
            if os.path.exists(path):
                os.remove(path)

    def test_plan_and_mapping_rows(self):
        m = imp.Matcher(entries())
        groups, shapes, bolts = imp.plan_parts(self.model, m)
        self.assertEqual(bolts, 3)                  # fastener, shank, nut
        self.assertEqual(sorted(p.name for p in shapes), ["PLATE", "RAIL"])
        self.assertIn(("column", u"PFC180*75"), groups)
        self.assertIn(("column", u"SHS100*100*6.0"), groups)
        self.assertEqual(len(groups[("framing", u"UB180*18")]["parts"]), 2)
        unmatched = [k for k in groups
                     if m.match(k[1], k[0], groups[k]["profile"])[0] is None]
        self.assertEqual(sorted(unmatched), [("column", u"SHS100*100*6.0"),
                                             ("framing", u"EA75*75*5")])
        rows = imp.mapping_rows(m, unmatched, groups)
        self.assertEqual([r["section"] for r in rows],
                         [u"SHS100*100*6.0", u"EA75*75*5"])
        for r in rows:
            self.assertEqual(r["default"], imp.SHAPE_LABEL)
            self.assertTrue(set(r["options"]) <= r["valid"])
        self.assertEqual(rows[0]["options"],
                         [u"PFC-Parallel Flange Channel-Column: 180PFC"])

    def test_file_section_map(self):
        fd, path = tempfile.mkstemp(suffix=".json")
        os.close(fd)
        os.remove(path)
        try:
            other = {u"Z-1": u"F: T"}              # another file's choice
            imp.save_map({"framing": dict(other, **{u"EA75*75*5": imp.PLATE})},
                         path)
            rows = imp.file_section_rows(self.model, entries(), imp.load_map(path))
            by = dict((r["section"], r) for r in rows)
            self.assertNotIn(u"Z-1", by)             # only this file's sections
            self.assertEqual(by[u"EA75*75*5"]["default"], imp.PLATE_LABEL)
            self.assertEqual(by[u"EA75*75*5"]["how"], u"saved")
            ub = by[u"UB180*18"]
            self.assertTrue(ub["matched"])
            self.assertEqual(ub["default"], ub["auto"])
            pick = [o for o in ub["options"] if o != ub["auto"]][0]
            special = dict(imp.SPECIAL_CHOICES)
            choices = dict(((r["cat"], r["section"]),
                            special.get(r["default"], r["default"]))
                           for r in rows)       # Save with nothing changed

            choices[("framing", u"EA75*75*5")] = None        # reset
            choices[("framing", u"UB180*18")] = pick         # changed
            imp.save_section_edits(rows, choices, path)
            data = imp.load_map(path)
            self.assertEqual(data["framing"], {u"Z-1": u"F: T",
                                               u"UB180*18": pick})
            self.assertEqual(data["column"], {})   # unchanged rows not saved
        finally:
            if os.path.exists(path):
                os.remove(path)

    def test_strip_holes(self):
        import math as m
        Brep, Part = steel.Brep, steel.Part
        hole = [(50 + 9 * m.cos(a * m.pi / 8), 50 + 9 * m.sin(a * m.pi / 8), 0.0)
                for a in range(16)]
        cutout = [(100, 100, 0), (180, 100, 0), (180, 180, 0), (100, 180, 0)]
        top = [[(0, 0, 0), (300, 0, 0), (300, 300, 0), (0, 300, 0)], hole, cutout]
        bore = [hole[0], hole[1], (hole[1][0], hole[1][1], -10.0),
                (hole[0][0], hole[0][1], -10.0)]
        rims = [hole[0], hole[1], hole[2]]          # a face of rim points only
        part = Part(1, "IFCPLATE")
        part.solids = [Brep([top, [bore], [rims]])]
        self.assertEqual(imp.strip_holes(part), 1)
        faces = part.solids[0].faces
        self.assertEqual(faces[0], [top[0], cutout])    # the cut-out stays
        self.assertEqual(len(faces), 2)                 # the rim face went
        self.assertEqual(faces[1], [bore])              # bore points are off-rim

    def test_classify(self):
        self.assertEqual(imp.classify(self.by_name["PFC180*75"].member,
                                      "column")[2], "column")
        self.assertEqual(imp.classify(self.by_name["SHS100*100*6.0"].member,
                                      "column")[2], "slanted")
        flat = steel.Member((0, 0, 0), (1000, 0, 500), (0, 1, 0), (0, 0, 1),
                            None)
        self.assertEqual(imp.classify(flat, "column")[:2], ("beam", "framing"))


# ── ORIENTATION (simulated Revit) ───────────────────────────────────────────

def rot(v, axis, a):
    """Rodrigues rotation of v about unit axis by a."""
    c, s = math.cos(a), math.sin(a)
    k = axis
    kv = imp._cross(k, v)
    kd = imp._dot(k, v)
    return tuple(v[i] * c + kv[i] * s + k[i] * kd * (1 - c) for i in range(3))


def revit_frame_beam(axis):
    """Framing at rotation 0: X along the axis, Z0 up-ish, Y0 = Z0 x X."""
    up = (0.0, 0.0, 1.0)
    if abs(imp._dot(up, axis)) > 0.999:
        up = (1.0, 0.0, 0.0)
    z0 = imp._unit(imp._sub(up, imp._mul(axis, imp._dot(up, axis))))
    y0 = imp._cross(z0, axis)
    return axis, y0, z0


def family_outline(kind):
    """A type's section in its own (S1, S2) frame, not bbox-centred, as a
    Revit family has it (PFC180: S1 range -51..24, measured in 2026)."""
    if kind == "U":
        # channel: web at S1 = -51, flanges toward +S1 up to 24, depth on S2
        x0, x1, h, tw, tf = -51.0, 24.0, 90.0, 6.0, 9.5
        return [(x0, -h), (x1, -h), (x1, -h + tf), (x0 + tw, -h + tf),
                (x0 + tw, h - tf), (x1, h - tf), (x1, h), (x0, h)]
    if kind == "L":
        # angle: legs along +S1 and -S2 from the heel at (-55, 20)
        hx, hy, t = -55.0, 20.0, 5.0
        return [(hx, hy), (hx + 75, hy), (hx + 75, hy - t), (hx + t, hy - t),
                (hx + t, hy - 75), (hx, hy - 75)]
    raise ValueError(kind)


def probe_of(outline):
    us = [p[0] for p in outline]
    vs = [p[1] for p in outline]
    m = ((min(us) + max(us)) / 2, (min(vs) + max(vs)) / 2)
    e = (max(us) - min(us), max(vs) - min(vs))
    (gx, gy), area = steel._centroid([outline])
    return m, e, (gx - m[0], gy - m[1])


def ifc_world(member):
    prof = member.profile
    cx, cy = prof.centre
    pts = []
    for x, y in steel.loop_points(prof.loops[0]):
        pts.append(imp._add(member.start, imp._add(
            imp._mul(member.x, x - cx), imp._mul(member.y, y - cy))))
    return pts


def place(outline, line_pt, s1, s2, m):
    return [imp._add(line_pt, imp._add(imp._mul(s1, u - m[0]),
                                       imp._mul(s2, v - m[1])))
            for u, v in outline]


def same_points(a, b, eps=0.01):
    return all(min(imp._len(imp._sub(p, q)) for q in b) < eps for p in a) and \
        all(min(imp._len(imp._sub(p, q)) for q in a) < eps for p in b)


class OrientationTests(unittest.TestCase):
    def simulate_beam(self, member, kind, reverse=False):
        start, end = member.start, member.end
        if reverse:
            start, end = end, start
        axis = imp._unit(imp._sub(end, start))
        bx, by, bz = revit_frame_beam(axis)
        s1, s2 = imp.section_axes((bx, by, bz), axis)
        self.assertEqual((s1, s2), (by, bz))
        outline = family_outline(kind)
        probe = probe_of(outline)
        h = imp.handedness(s1, s2, member.x, member.y)
        A, score = imp.fit_orientation(probe, member.profile.extent,
                                       member.profile.centroid, h)
        self.assertLess(score, 0.01)
        d1, d2 = imp.section_dirs(A, member.x, member.y)
        self.assertTrue(close(imp._cross(d1, d2), axis, 1e-9))
        phi = imp.rotation_angle(d1, s1, s2)
        # the measured Revit rule: Z(t) = cos t Z0 + sin t Y0, t = -phi
        t = -phi
        z_t = imp._add(imp._mul(bz, math.cos(t)), imp._mul(by, math.sin(t)))
        y_t = imp._cross(z_t, bx)
        self.assertTrue(close(y_t, d1, 1e-9))
        self.assertTrue(close(z_t, d2, 1e-9))
        spec = math.atan2(imp._dot(d2, by), imp._dot(d2, bz))
        diff = (t - spec + math.pi) % (2 * math.pi) - math.pi
        self.assertAlmostEqual(diff, 0.0, 9)
        # the line passes through the IFC bbox centre; the type is moved
        # by -(m1 d1 + m2 d2) so its bbox centre lands there too
        revit = place(outline, member.start, y_t, z_t, probe[0])
        self.assertTrue(same_points(revit, ifc_world(member)))

    def test_channel_beam(self):
        # an IFC channel with its web on +x, the type's web on -S1
        u = steel.Profile("U", "PFC", [[(-x, y) for x, y in
                                        reversed(family_outline("U"))]])
        m = steel.Member((0, 0, 0), (6000, 0, 0), (0, 1, 0), (0, 0, 1), u)
        self.simulate_beam(m, "U")

    def test_channel_beam_reversed_line(self):
        u = steel.Profile("U", "PFC", [family_outline("U")])
        m = steel.Member((0, 0, 0), (3000, 4000, 1000), None, None, u)
        d = m.direction
        m.x = imp._unit(imp._cross((0, 0, 1), d))
        m.y = imp._cross(d, m.x)
        self.simulate_beam(m, "U", reverse=True)    # h = -1

    def test_angle_brace(self):
        ea = steel.Profile("L", "EA", [[(-37.5, -37.5), (37.5, -37.5),
                                        (37.5, -32.5), (-32.5, -32.5),
                                        (-32.5, 37.5), (-37.5, 37.5)]])
        d = imp._unit((1, 0, 1))
        for roll in (0.0, 0.7, 2.0, -2.5):
            x = rot((0, 1, 0), d, roll)
            y = imp._cross(d, x)
            m = steel.Member((0, 0, 0), imp._mul(d, 4000), x, y, ea)
            self.simulate_beam(m, "L")

    def test_vertical_column(self):
        u = steel.Profile("U", "PFC", [family_outline("U")])
        for up in (True, False):
            for roll in (0.3, 1.9):
                x = (math.cos(roll), math.sin(roll), 0.0)
                y = (-math.sin(roll), math.cos(roll), 0.0)
                if not up:
                    y = imp._mul(y, -1)     # IFC extruded downwards
                start = (0, 0, 0) if up else (0, 0, 3000)
                end = (0, 0, 3000) if up else (0, 0, 0)
                m = steel.Member(start, end, x, y, u)
                bottom = start if up else end
                axis = (0.0, 0.0, 1.0)
                x0, y0, z0 = (1.0, 0.0, 0.0), (0.0, 1.0, 0.0), axis
                s1, s2 = imp.section_axes((x0, y0, z0), axis)
                outline = family_outline("U")
                probe = probe_of(outline)
                h = imp.handedness(s1, s2, x, y)
                self.assertEqual(h, 1 if up else -1)
                A, score = imp.fit_orientation(probe, u.extent, u.centroid, h)
                self.assertLess(score, 0.01)
                d1, d2 = imp.section_dirs(A, x, y)
                phi = imp.rotation_angle(d1, s1, s2)
                # RotateElement about +Z: X(t) = cos t X0 + sin t Y0
                xt = rot(x0, axis, phi)
                yt = rot(y0, axis, phi)
                self.assertTrue(close(xt, d1, 1e-9))
                self.assertTrue(close(yt, d2, 1e-9))
                revit = place(outline, bottom, xt, yt, probe[0])
                ifc = [imp._add(p, imp._sub(bottom, start))
                       for p in ifc_world(m)]
                self.assertTrue(same_points(revit, ifc))

    def test_symmetric_section_takes_identity(self):
        A, score = imp.fit_orientation(((0, 0), (90, 179), (0, 0)),
                                       (90, 179), (0, 0), 1)
        self.assertEqual(A, ((1, 0), (0, 1)))
        A, score = imp.fit_orientation(((0, 0), (179, 90), (0, 0)),
                                       (90, 179), (0, 0), 1)
        self.assertEqual(abs(A[0][1]), 1)
        self.assertLess(score, 1e-9)


# ── IMPORTER WITH A FAKE REVIT ──────────────────────────────────────────────

class _V(object):
    def __init__(self, t):
        self.X, self.Y, self.Z = t


class FakeDB(object):
    """Just enough of Autodesk.Revit.DB for Importer.orient/check/shift."""

    class XYZ(object):
        def __init__(self, x, y, z):
            self.X, self.Y, self.Z = x, y, z

    class Line(object):
        def __init__(self, a, b):
            self.a, self.b = a, b

        @staticmethod
        def CreateBound(a, b):
            return FakeDB.Line(a, b)

        def GetEndPoint(self, i):
            return self.a if i == 0 else self.b

    class LocationCurve(object):
        def __init__(self, curve):
            self.Curve = curve

    class LocationPoint(object):
        def __init__(self, p):
            self.Point = p

    class BuiltInParameter(object):
        STRUCTURAL_BEND_DIR_ANGLE = "bend"

    class ElementTransformUtils(object):
        @staticmethod
        def RotateElement(doc, eid, line, angle):
            eid.owner.spin += angle

        @staticmethod
        def MoveElement(doc, eid, v):
            o = eid.owner
            o.offset = imp._add(o.offset, (v.X, v.Y, v.Z))


class FakeParam(object):
    IsReadOnly = False

    def __init__(self, owner):
        self.owner = owner

    def Set(self, v):
        self.owner.bend = v


class FakeId(object):
    def __init__(self, owner):
        self.owner = owner


class FakeMember(object):
    """A placed member: framing (cross-section rotation, Revit's rule
    Z(t) = cos t Z0 + sin t Y0 times rule) or a vertical column (turned)."""

    def __init__(self, p0, p1, column, rule=1.0):
        self.p0, self.p1 = p0, p1
        self.column = column
        self.rule = rule
        self.bend = 0.0
        self.spin = 0.0
        self.offset = (0.0, 0.0, 0.0)
        self.Id = FakeId(self)
        axis = imp._unit(imp._sub(p1, p0))
        if column:
            self.frame0 = ((1.0, 0.0, 0.0), (0.0, 1.0, 0.0), (0.0, 0.0, 1.0))
            self.Location = FakeDB.LocationPoint(_V(p0))
        else:
            self.frame0 = revit_frame_beam(axis)
            self.Location = FakeDB.LocationCurve(
                FakeDB.Line(_V(p0), _V(p1)))

    def get_Parameter(self, bip):
        return None if self.column else FakeParam(self)

    def frame(self):
        x0, y0, z0 = self.frame0
        if self.column:
            return (rot(x0, z0, self.spin), rot(y0, z0, self.spin), z0)
        t = self.rule * self.bend
        z = imp._add(imp._mul(z0, math.cos(t)), imp._mul(y0, math.sin(t)))
        return (x0, imp._cross(z, x0), z)

    def GetTransform(self):
        class T(object):
            pass
        t = T()
        t.BasisX, t.BasisY, t.BasisZ = [_V(b) for b in self.frame()]
        return t

    def world(self, outline):
        """The type's section outline where Revit would draw it (mm)."""
        b = self.frame()
        s1, s2 = imp.section_axes(b, imp._unit(imp._sub(self.p1, self.p0)))
        base = imp._add(self.p0, self.offset)
        return [imp._mul(imp._add(base, imp._add(
            imp._mul(s1, u / imp.MM_PER_FT), imp._mul(s2, v / imp.MM_PER_FT))),
            imp.MM_PER_FT) for u, v in outline]


class ImporterTests(unittest.TestCase):
    def setUp(self):
        self.saved = imp.DB
        imp.DB = FakeDB

    def tearDown(self):
        imp.DB = self.saved

    def run_plan(self, member, mode, outline, rule=1.0, importer=None):
        mp = imp.Mapper()
        if importer is None:
            importer = imp.Importer.__new__(imp.Importer)
            importer.mapper = mp
            importer.doc = None
            importer.sign = {"beam": -1.0, "slanted": -1.0}
            importer.probes = {}

        class Sym(object):
            class Id(object):
                @staticmethod
                def ToString():
                    return "1"

        class E(object):
            symbol = Sym()
        importer.probes[("1", mode)] = probe_of(outline)
        part = steel.Part(1, "IFCBEAM")
        part.member = member
        plan = imp.Plan(part, "beam", "framing", mode, E(), "test")
        p0, p1 = mp.point(member.start), mp.point(member.end)
        if mode == "column" and p1[2] < p0[2]:
            p0, p1 = p1, p0
        plan.fi = FakeMember(p0, p1, mode == "column", rule)
        plan.axis = imp._unit(imp._sub(p1, p0))
        importer.orient(plan)
        redo = not importer.check(plan)
        if redo:
            self.assertTrue(importer.check(plan, final=True))
        importer.shift(plan)
        self.assertLess(plan.score, 0.01)
        ifc = ifc_world(member)
        if mode == "column" and member.end[2] < member.start[2]:
            # the column is drawn from its bottom end
            ifc = [imp._add(p, imp._sub(member.end, member.start))
                   for p in ifc]
        self.assertTrue(same_points(plan.fi.world(outline), ifc, 0.05))
        return importer, redo

    def test_beam_and_learned_sign(self):
        u = steel.Profile("U", "PFC", [family_outline("U")])
        d = imp._unit((3, 4, 1))
        x = imp._unit(imp._cross((0, 0, 1), d))
        x = rot(x, d, 0.8)
        m = steel.Member((100, 200, 3000), imp._add((100, 200, 3000),
                         imp._mul(d, 5000)), x, imp._cross(d, x), u)
        importer, redo = self.run_plan(m, "beam", family_outline("U"))
        self.assertFalse(redo)                  # the measured rule holds
        # a Revit that turns the other way: flipped, then remembered
        importer, redo = self.run_plan(m, "beam", family_outline("U"), -1.0)
        self.assertTrue(redo)
        self.assertEqual(importer.sign["beam"], 1.0)
        importer, redo = self.run_plan(m, "beam", family_outline("U"), -1.0,
                                       importer)
        self.assertFalse(redo)

    def test_vertical_column_downwards(self):
        ea = steel.Profile("L", "EA", [[(-37.5, -37.5), (37.5, -37.5),
                                        (37.5, -32.5), (-32.5, -32.5),
                                        (-32.5, 37.5), (-37.5, 37.5)]])
        x = (math.cos(2.2), math.sin(2.2), 0.0)
        y = imp._cross((0, 0, -1), x)           # extruded downwards
        m = steel.Member((500, 500, 6000), (500, 500, 3000), x, y, ea)
        importer, redo = self.run_plan(m, "column", family_outline("L"))
        self.assertFalse(redo)


# ── GEOMETRY HELPERS ────────────────────────────────────────────────────────

class HelperTests(unittest.TestCase):
    def test_loop_orientations(self):
        outer = [(0, 0), (0, 10), (10, 10), (10, 0)]      # clockwise
        hole = [(2, 2), (4, 2), (4, 4), (2, 4)]           # anticlockwise
        r = imp.loop_orientations([outer, hole], (1, 0, 0), (0, 1, 0),
                                  (0, 0, 1))
        self.assertEqual([f for l, f in r], [True, True])
        r = imp.loop_orientations([outer, hole], (1, 0, 0), (0, 1, 0),
                                  (0, 0, -1))
        self.assertEqual([f for l, f in r], [False, False])
        r = imp.loop_orientations([("circle", 0, 0, 5), ("circle", 0, 0, 3)],
                                  (1, 0, 0), (0, 1, 0), (0, 0, 1))
        self.assertEqual([f for l, f in r], [False, True])

    def test_mapper(self):
        m = imp.Mapper(((0, 1, 0), (-1, 0, 0), (0, 0, 1)), (10, 0, 0))
        self.assertTrue(close(m.point((304.8, 0, 0)), (10, 1, 0)))
        self.assertTrue(close(m.vector((0, 1, 0)), (-1, 0, 0)))
        c = imp.Mapper(shift=(-304.8, 0, 0))
        self.assertTrue(close(c.point((304.8, 0, 0)), (0, 0, 0)))

    def test_clip_line_bounded(self):
        clip = {"kind": "bounded", "origin": (900, 0, 0), "keep": (-1, 0, 0),
                "plane": ((900, 0, 0), (0, 1, 0), (0, 0, 1)), "axis": (1, 0, 0),
                "poly": [(-10, -10), (10, -10), (10, 10), (-10, 10)]}
        s, e, t, o = imp.clip_line((0, 0, 0), (1000, 0, 0), [clip])
        self.assertTrue(t)
        self.assertTrue(close(e, (900, 0, 0)))
        clip["poly"] = [(20, 20), (40, 20), (40, 40), (20, 40)]    # a cope
        s, e, t, o = imp.clip_line((0, 0, 0), (1000, 0, 0), [clip])
        self.assertFalse(t)
        self.assertEqual(o, 1)

    def test_seed_ifc(self):
        import gzip
        ifc = build_ifc()[0].encode("latin-1")
        tmp = tempfile.mkdtemp()
        try:
            for payload in (ifc, gzip.compress(ifc)):
                seed = os.path.join(tmp, "m.seed")
                with open(seed, "wb") as f:
                    f.write(b"SEED43-STUDIO-V1\nsourceFormat: ifc\n"
                            b"sourceName: frame.ifc\nsourceBytes: " +
                            str(len(payload)).encode() + b"\n"
                            b"---BEGIN-SOURCE-DATA---\n" + payload + b"tail")
                self.assertEqual(imp.seed_format(seed), "ifc")
                out, name = imp.seed_ifc(seed, tmp)
                self.assertEqual(name, "frame.ifc")
                with open(out, "rb") as f:
                    self.assertEqual(f.read(), ifc)
                self.assertTrue(steel.read_steel(out).parts)
            with open(seed, "wb") as f:
                f.write(b"SEED43-STUDIO-V1\nsourceFormat: bsw\nsourceBytes: 1"
                        b"\n---BEGIN-SOURCE-DATA---\nx")
            self.assertEqual(imp.seed_format(seed), "bsw")
            out, name = imp.seed_ifc(seed, tmp)     # a StruCad model comes out as .bsw
            self.assertTrue(out.endswith(".bsw"))
            with open(seed, "wb") as f:
                f.write(b"SEED43-STUDIO-V1\nsourceFormat: step\nsourceBytes: 1"
                        b"\n---BEGIN-SOURCE-DATA---\nx")
            self.assertRaises(ValueError, imp.seed_ifc, seed, tmp)
        finally:
            import shutil
            shutil.rmtree(tmp)

    def test_grade(self):
        self.assertEqual(imp.grade_miss(u"SHS-C350L0", u"C350L0"), 0)
        self.assertEqual(imp.grade_miss(u"SHS", u"C350L0"), 1)
        self.assertGreater(imp.grade_miss(u"SHS-C450L0", u"C350L0"), 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
