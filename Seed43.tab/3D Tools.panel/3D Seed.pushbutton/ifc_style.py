# -*- coding: utf-8 -*-
"""IFC-style naming for 3D Seed: the IFC class Revit's own IFC exporter
would give an element, its IFC GlobalId, and the spatial tree and fixed
fields Studio shows for an IFC model.

No Revit types in here, so it runs under CPython for testing as well as
IronPython 2 inside Revit.
"""
import re

__all__ = ["ifc_class", "ifc_guid", "kind_of", "storey_of", "STRUCAD_PSET"]

# The property set Studio reads its fixed card fields from on an IFC.
STRUCAD_PSET = u"StruCad"

# BuiltInCategory name to IFC2x3 class, as Revit's IFC exporter maps them by
# default. Anything not listed is an IfcBuildingElementProxy.
CATEGORY_CLASSES = {
    "OST_StructuralFraming": "IfcBeam",
    "OST_StructuralColumns": "IfcColumn",
    "OST_Columns": "IfcColumn",
    "OST_StructuralFoundation": "IfcFooting",
    "OST_StructConnectionPlates": "IfcPlate",
    "OST_StructConnectionBolts": "IfcMechanicalFastener",
    "OST_StructConnectionAnchors": "IfcMechanicalFastener",
    "OST_StructConnectionWelds": "IfcFastener",
    "OST_StructConnectionShearStuds": "IfcMechanicalFastener",
    "OST_StructConnectionProfiles": "IfcMember",
    "OST_StructConnectionHoles": "IfcOpeningElement",
    "OST_Rebar": "IfcReinforcingBar",
    "OST_FabricReinforcement": "IfcReinforcingMesh",
    "OST_Walls": "IfcWallStandardCase",
    "OST_Floors": "IfcSlab",
    "OST_StructuralFramingSystem": "IfcBeam",
    "OST_Roofs": "IfcRoof",
    "OST_Ceilings": "IfcCovering",
    "OST_Doors": "IfcDoor",
    "OST_Windows": "IfcWindow",
    "OST_Stairs": "IfcStair",
    "OST_StairsRuns": "IfcStairFlight",
    "OST_StairsLandings": "IfcSlab",
    "OST_StairsRailing": "IfcRailing",
    "OST_Railings": "IfcRailing",
    "OST_Ramps": "IfcRamp",
    "OST_CurtainWallPanels": "IfcPlate",
    "OST_CurtainWallMullions": "IfcMember",
    "OST_Furniture": "IfcFurnishingElement",
    "OST_FurnitureSystems": "IfcFurnishingElement",
    "OST_Casework": "IfcFurnishingElement",
    "OST_PipeCurves": "IfcPipeSegment",
    "OST_FlexPipeCurves": "IfcPipeSegment",
    "OST_PipeFitting": "IfcPipeFitting",
    "OST_PipeAccessory": "IfcValve",
    "OST_DuctCurves": "IfcDuctSegment",
    "OST_FlexDuctCurves": "IfcDuctSegment",
    "OST_DuctFitting": "IfcDuctFitting",
    "OST_DuctTerminal": "IfcAirTerminal",
    "OST_CableTray": "IfcCableCarrierSegment",
    "OST_CableTrayFitting": "IfcCableCarrierFitting",
    "OST_Conduit": "IfcCableCarrierSegment",
    "OST_ConduitFitting": "IfcCableCarrierFitting",
    "OST_LightingFixtures": "IfcLightFixture",
    "OST_PlumbingFixtures": "IfcSanitaryTerminal",
    "OST_MechanicalEquipment": "IfcFlowTerminal",
    "OST_ElectricalEquipment": "IfcElectricDistributionPoint",
    "OST_Topography": "IfcSite",
    "OST_Toposolid": "IfcSite",
    "OST_Site": "IfcSite",
    "OST_Parts": "IfcBuildingElementPart",
    "OST_GenericModel": "IfcBuildingElementProxy",
}

# FamilyInstance.StructuralType name to the class it overrides with.
STRUCTURAL_CLASSES = {"Brace": "IfcMember", "Column": "IfcColumn",
                      "Footing": "IfcFooting"}

_CLASS = re.compile(r"^\s*(Ifc[A-Za-z]+)")
_GUID_CHARS = ("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
               "abcdefghijklmnopqrstuvwxyz_$")


def ifc_class(category, structural_type=u"", export_as=u""):
    """The IFC class for an element: its "Export to IFC As" value when it
    names one (e.g. 'IfcPlate' or 'IfcMember.USERDEFINED'), else the
    structural type (braces are members), else its category's class."""
    m = _CLASS.match(export_as or u"")
    if m:
        name = m.group(1)
        if name.lower().endswith("type"):
            name = name[:-4]
        if name.lower() not in ("ifcnone", "ifcdontexport"):
            return name
    cls = CATEGORY_CLASSES.get(category or "", "IfcBuildingElementProxy")
    if category in ("OST_StructuralFraming", "OST_StructuralColumns"):
        cls = STRUCTURAL_CLASSES.get(structural_type or u"", cls)
    return cls


def kind_of(cls):
    """'IfcBeam' -> 'Beam', 'IfcWallStandardCase' -> 'Wall',
    'IfcBuildingElementProxy' -> 'Proxy'."""
    if cls == "IfcBuildingElementProxy":
        return "Proxy"
    name = cls[3:] if cls.startswith("Ifc") else cls
    if name.endswith("StandardCase") and len(name) > 12:
        name = name[:-12]
    return name


def ifc_guid(guid):
    """A GUID ('3f2504e0-4f89-11d3-9a0c-0305e82c3301' or 32 hex digits) as
    a 22-character IFC GlobalId."""
    hexs = re.sub(r"[^0-9A-Fa-f]", u"", guid or u"")
    if len(hexs) != 32:
        return u""
    n = int(hexs, 16)
    out = []
    for i in range(22):
        bits = 2 if i == 0 else 6
        shift = 6 * (21 - i)
        out.append(_GUID_CHARS[(n >> shift) & ((1 << bits) - 1)])
    return u"".join(out)


def storey_of(levels, z, tol=1e-6):
    """The (name, elevation) of the highest level at or below z, else the
    lowest one, else None. levels is [(name, elevation)], same units as z."""
    if not levels:
        return None
    best = None
    for lv in sorted(levels, key=lambda l: l[1]):
        if lv[1] <= z + tol:
            best = lv
    return best or min(levels, key=lambda l: l[1])
