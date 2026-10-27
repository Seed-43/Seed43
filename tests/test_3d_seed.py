# -*- coding: utf-8 -*-
"""CPython tests for 3D Seed's IFC-style layout.

    python3 tests/test_3d_seed.py

Checks ifc_style (IFC class mapping, GlobalId compression, storeys) and
runs the button's build_scene against a fake Revit, then writes the .seed
and reads it back with Seed Import's reader. Nothing here needs Revit.
"""
import os
import sys
import tempfile
import types
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
PANEL = os.path.join(HERE, "..", "Seed43.tab", "Document Studio.panel")
SEED_DIR = os.path.join(PANEL, "3D Seed.pushbutton")
sys.path.insert(0, SEED_DIR)
sys.path.insert(0, os.path.join(PANEL, "Seed Import.pushbutton"))

import ifc_style                                  # noqa: E402


class IfcStyleTests(unittest.TestCase):
    def test_classes(self):
        c = ifc_style.ifc_class
        self.assertEqual(c("OST_StructuralFraming", "Beam"), "IfcBeam")
        self.assertEqual(c("OST_StructuralFraming", "Brace"), "IfcMember")
        self.assertEqual(c("OST_StructuralColumns", "Column"), "IfcColumn")
        self.assertEqual(c("OST_StructConnectionPlates"), "IfcPlate")
        self.assertEqual(c("OST_Walls"), "IfcWallStandardCase")
        self.assertEqual(c("OST_Something"), "IfcBuildingElementProxy")
        self.assertEqual(c("OST_GenericModel", "", "IfcPlate"), "IfcPlate")
        self.assertEqual(c("OST_GenericModel", "", "IfcMember.USERDEFINED"),
                         "IfcMember")
        self.assertEqual(c("OST_GenericModel", "", "IfcBeamType"), "IfcBeam")
        self.assertEqual(c("OST_Walls", "", "DontExport"),
                         "IfcWallStandardCase")

    def test_kind(self):
        self.assertEqual(ifc_style.kind_of("IfcBeam"), "Beam")
        self.assertEqual(ifc_style.kind_of("IfcWallStandardCase"), "Wall")
        self.assertEqual(ifc_style.kind_of("IfcBuildingElementProxy"),
                         "Proxy")

    def test_guid(self):
        # Reference values from ifcopenshell.guid.compress
        g = ifc_style.ifc_guid
        self.assertEqual(g("00000000-0000-0000-0000-000000000000"),
                         "0000000000000000000000")
        self.assertEqual(g("ffffffff-ffff-ffff-ffff-ffffffffffff"),
                         "3$$$$$$$$$$$$$$$$$$$$$")
        self.assertEqual(g("3f2504e0-4f89-11d3-9a0c-0305e82c3301"),
                         REF_GUID)
        self.assertEqual(g("nonsense"), "")

    def test_storey(self):
        lv = [("L2", 10.0), ("L1", 0.0), ("Roof", 20.0)]
        self.assertEqual(ifc_style.storey_of(lv, 12.0), ("L2", 10.0))
        self.assertEqual(ifc_style.storey_of(lv, 10.0), ("L2", 10.0))
        self.assertEqual(ifc_style.storey_of(lv, -5.0), ("L1", 0.0))
        self.assertIsNone(ifc_style.storey_of([], 1.0))


REF_GUID = "0$9GJWJuaHqveC0mNeB3C1"     # ifcopenshell.guid.compress


# ── FAKE REVIT ──────────────────────────────────────────────────────────────

class Id(object):
    def __init__(self, v):
        self.Value = v

    def __eq__(self, other):
        return isinstance(other, Id) and other.Value == self.Value

    def __ne__(self, other):
        return not self.__eq__(other)

    def __hash__(self):
        return hash(self.Value)


INVALID = Id(-1)


class Param(object):
    def __init__(self, value, double=None, eid=None):
        self.value = value
        self.double = double
        self.eid = eid
        self.HasValue = True
        self.StorageType = "Double" if double is not None else "String"

    def AsValueString(self):
        return self.value

    def AsString(self):
        return self.value

    def AsDouble(self):
        return self.double

    def AsElementId(self):
        return self.eid


class Cat(object):
    def __init__(self, name, bic):
        self.Name = name
        self.BuiltInCategory = bic
        self.CategoryType = "Model"


class El(object):
    def __init__(self, eid, name="", cat=None, type_id=None, params=None,
                 level=None, struct="", assembly=None):
        self.Id = Id(eid)
        self.Name = name
        self.Category = cat
        self._type = type_id
        self.params = params or {}
        self.LevelId = level or INVALID
        self.StructuralType = struct
        self.AssemblyInstanceId = assembly or INVALID
        self.FamilyName = ""
        self.Parameters = []
        self.UniqueId = "u%d" % eid

    def GetTypeId(self):
        return self._type or INVALID

    def get_Parameter(self, bip):
        return self.params.get(bip)

    def LookupParameter(self, name):
        return self.params.get(name)

    def get_BoundingBox(self, view):
        return None


class Doc(object):
    def __init__(self, title, elements):
        self.Title = title
        self.els = dict((e.Id.Value, e) for e in elements)
        self.ActiveProjectLocation = El(0, "Survey Point Site")
        self.ProjectInformation = types.SimpleNamespace(BuildingName="Shed")
        self.Application = types.SimpleNamespace(VersionNumber="2026")

    def GetElement(self, eid):
        return self.els.get(eid.Value) if eid is not None else None


def fake_db():
    class BIP(object):
        pass
    for name in ("ALL_MODEL_MARK", "ALL_MODEL_DESCRIPTION", "PHASE_CREATED",
                 "IFC_GUID", "IFC_EXPORT_ELEMENT_AS",
                 "INSTANCE_REFERENCE_LEVEL_PARAM", "FAMILY_BASE_LEVEL_PARAM",
                 "SCHEDULE_LEVEL_PARAM", "STRUCTURAL_FRAME_CUT_LENGTH",
                 "INSTANCE_LENGTH_PARAM", "CURVE_ELEM_LENGTH",
                 "STRUCTURAL_MATERIAL_PARAM"):
        setattr(BIP, name, name)

    class Element(object):
        Name = property(lambda self: self.Name)

    class FEC(object):
        def __init__(self, d, *a):
            self.d = d

        def OfClass(self, cls):
            return [e for e in self.d.els.values() if getattr(e, "is_level",
                                                              False)]

    db = types.SimpleNamespace(
        BuiltInParameter=BIP,
        ElementId=types.SimpleNamespace(InvalidElementId=INVALID),
        StorageType=types.SimpleNamespace(Double="Double", String="String"),
        CategoryType=types.SimpleNamespace(Model="Model"),
        FilteredElementCollector=FEC, Level=object,
        ExportUtils=types.SimpleNamespace(
            GetExportId=lambda d, eid: "3f2504e0-4f89-11d3-9a0c-%012x"
            % eid.Value),
        Element=types.SimpleNamespace(Name=property(lambda e: e.Name)),
        IExportContext=object, LabelUtils=None)
    return db


def load_script(document):
    """Import 3D Seed's script.py with pyrevit faked out."""
    db = fake_db()
    pyrevit = types.ModuleType("pyrevit")
    pyrevit.revit = types.SimpleNamespace(doc=document)
    pyrevit.DB = db
    pyrevit.forms = None
    sys.modules["pyrevit"] = pyrevit
    path = os.path.join(SEED_DIR, "script.py")
    mod = types.ModuleType("seed_script")
    mod.__file__ = path
    with open(path, encoding="utf-8") as f:
        code = compile(f.read(), path, "exec")
    exec(code, mod.__dict__)
    # Element.Name.__get__ in the script: the fake elements hold Name
    mod.element_name = lambda el: getattr(el, "Name", "") or ""
    return mod


class BuildSceneTests(unittest.TestCase):
    def setUp(self):
        from seed_writer import SeedPart
        steel = Cat("Structural Framing", "OST_StructuralFraming")
        cols = Cat("Structural Columns", "OST_StructuralColumns")
        lv1 = El(10, "Level 1")
        lv1.is_level = True
        lv1.ProjectElevation = 0.0
        lv2 = El(11, "Level 2")
        lv2.is_level = True
        lv2.ProjectElevation = 10.0
        ub = El(20, "180UB18.1", params={"ALL_MODEL_DESCRIPTION":
                                         Param("Universal beam")})
        ub.FamilyName = "UB-Universal Beam"
        uc = El(21, "150UC23.4")
        uc.FamilyName = "UC-Universal Column-Column"
        mat = El(30, "C350L0")
        asm = El(40, "Frame A")
        beam = El(100, cat=steel, type_id=Id(20), struct="Beam", assembly=Id(40),
                  params={"ALL_MODEL_MARK": Param("B1"),
                          "INSTANCE_REFERENCE_LEVEL_PARAM": Param("", eid=Id(11)),
                          "STRUCTURAL_FRAME_CUT_LENGTH": Param("6000", 19.685),
                          "STRUCTURAL_MATERIAL_PARAM": Param("C350L0", eid=Id(30)),
                          "PHASE_CREATED": Param("New Construction")})
        brace = El(101, cat=steel, type_id=Id(20), struct="Brace",
                   level=Id(10))
        col = El(102, cat=cols, type_id=Id(21), struct="Column",
                 params={"FAMILY_BASE_LEVEL_PARAM": Param("", eid=Id(10)),
                         "IFC_GUID": Param("2KjfQ5sB9Ag8Y0Ng2yQ4lH")})
        plate = El(103, "Plate", cat=Cat("Generic Models", "OST_GenericModel"),
                   params={"IfcExportAs": Param("IfcPlate")})
        self.doc = Doc("Shed", [lv1, lv2, ub, uc, mat, asm, beam, brace, col,
                                plate])
        self.mod = load_script(self.doc)

        class Collected(object):
            pass
        collected = {}
        for el in (plate, col, brace, beam):
            c = Collected()
            c.el, c.doc, c.model = el, self.doc, 0
            p = SeedPart((0.5, 0.5, 0.5), 1.0, "Steel")
            p.add([0, 0, 0, 1, 0, 0, 0, 1, 0], [0, 1, 2])
            c.parts = [p]
            collected[(0, el.Id.Value)] = c
        self.ctx = types.SimpleNamespace(collected=collected, models=["Shed"])

    def tearDown(self):
        sys.modules.pop("pyrevit", None)

    def test_tree_and_fields(self):
        scene = self.mod.build_scene(self.ctx, False)
        self.assertIsNone(scene.check_tree_order())
        tree = [(n.cls, n.name, n.parent) for n in scene.nodes]
        self.assertEqual(tree[:3], [("IfcProject", "Shed", -1),
                                    ("IfcSite", "Survey Point Site", 0),
                                    ("IfcBuilding", "Shed", 1)])
        storeys = [n.name for n in scene.nodes if n.cls == "IfcBuildingStorey"]
        self.assertEqual(storeys, ["Level 1", "Level 2", "(no level)"])
        by_name = dict((scene.nodes[it.node].name, it) for it in scene.items)
        beam = by_name["UB-Universal Beam:180UB18.1:100"]
        self.assertEqual(scene.nodes[beam.node].cls, "IfcBeam")
        self.assertEqual(scene.nodes[scene.nodes[beam.node].parent].name,
                         "Level 2")
        self.assertEqual(beam.cat, "Structural Framing")
        self.assertEqual(beam.props["StruCad"], {
            "Kind": "Beam", "Mark": "B1", "Section": "180UB18.1",
            "Grade": "C350L0", "Assembly": "Frame A", "Length": "6000",
            "Desc": "Universal beam", "Phase": "New Construction"})
        self.assertEqual(beam.props["IFC"]["GlobalId"],
                         ifc_style.ifc_guid("3f2504e0-4f89-11d3-9a0c-%012x" % 100))
        brace = by_name["UB-Universal Beam:180UB18.1:101"]
        self.assertEqual(scene.nodes[brace.node].cls, "IfcMember")
        self.assertEqual(brace.props["StruCad"]["Kind"], "Member")
        col = by_name["UC-Universal Column-Column:150UC23.4:102"]
        self.assertEqual(col.props["IFC"]["GlobalId"], "2KjfQ5sB9Ag8Y0Ng2yQ4lH")
        self.assertEqual(col.props["IFC"]["Storey"], "Level 1")
        plate = by_name["Plate:103"]
        self.assertEqual(scene.nodes[plate.node].cls, "IfcPlate")
        self.assertNotIn("Section", plate.props["StruCad"])

    def test_writes_and_reads_back(self):
        from seed_reader import read_seed
        scene = self.mod.build_scene(self.ctx, False)
        fd, path = tempfile.mkstemp(suffix=".seed")
        os.close(fd)
        try:
            scene.write(path, u"Shed.s43nw")
            model = read_seed(path)
            self.assertEqual(len(model.items), 4)
            names = [model.nodes[it["node"]]["n"] for it in model.items]
            self.assertIn("Plate:103", names)
            beam = [it for it in model.items
                    if model.nodes[it["node"]]["c"] == "IfcBeam"][0]
            self.assertEqual(model.path(beam["node"]),
                             ["Shed", "Survey Point Site", "Shed", "Level 2"])
            self.assertEqual(beam["props"]["StruCad"]["Mark"], "B1")
        finally:
            os.remove(path)


if __name__ == "__main__":
    unittest.main(verbosity=2)
