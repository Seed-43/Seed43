# -*- coding: utf-8 -*-
# inspector.py
import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from pyrevit import revit, output, forms

doc   = revit.doc
uidoc = revit.uidoc
out   = output.get_output()

# ── SELECTION FILTER ──────────────────────────────────────────────────────────

class AnyFilter(ISelectionFilter):
    def AllowElement(self, element):
        return not isinstance(element, RevitLinkInstance)
    def AllowReference(self, reference, xyz):
        return False

# ── HELPERS ───────────────────────────────────────────────────────────────────

def bip_name(param_id):
    """Return the built-in parameter name for a given parameter ID."""
    try:
        return str(BuiltInParameter(int(str(param_id))))
    except Exception:
        return "-"

def param_value(param):
    """Return a readable string value for a parameter."""
    if not param or not param.HasValue:
        return "<no value>"
    return param.AsString() or param.AsValueString() or str(param.AsDouble())

def element_name(eid):
    """Resolve an ElementId to a display name string."""
    if not eid or eid == ElementId.InvalidElementId:
        return "-"
    el = doc.GetElement(eid)
    return getattr(el, "Name", str(eid)) if el else str(eid)

def xyz_str(xyz):
    """Format an XYZ point as a readable string in feet."""
    if xyz is None:
        return "-"
    return "({:.3f}, {:.3f}, {:.3f}) ft".format(xyz.X, xyz.Y, xyz.Z)

def dump_params(element):
    """Return sorted rows of name, ID, and value for all parameters."""
    rows = []
    for p in element.Parameters:
        try:
            name = p.Definition.Name if p.Definition else "?"
            rows.append([name, str(p.Id), param_value(p)])
        except Exception:
            pass
    return sorted(rows, key=lambda r: r[0].lower())

# ── SECTIONS ──────────────────────────────────────────────────────────────────

def section_identity(element, cat):
    out.print_md("## Identity")
    rows = [
        ["Element ID",  str(element.Id)],
        ["Unique ID",   element.UniqueId],
        ["Class",       type(element).__name__],
        ["Name",        getattr(element, "Name", "-") or "-"],
        ["Category",    cat.Name if cat else "-"],
    ]
    if cat:
        try:
            rows.append(["BuiltInCategory", str(cat.BuiltInCategory)])
        except Exception:
            pass
    try:
        lvl_id = getattr(element, "LevelId", None)
        if lvl_id and lvl_id != ElementId.InvalidElementId:
            rows.append(["Level", element_name(lvl_id)])
    except Exception:
        pass
    try:
        rows.append(["Phase Created",    element_name(element.CreatedPhaseId)])
        rows.append(["Phase Demolished", element_name(element.DemolishedPhaseId)])
    except Exception:
        pass
    try:
        ws_id = element.WorksetId
        if ws_id:
            ws_table = doc.GetWorksetTable()
            ws       = ws_table.GetWorkset(ws_id)
            rows.append(["Workset", ws.Name if ws else str(ws_id)])
    except Exception:
        pass
    out.print_table(rows, columns=["Property", "Value"])

def section_type_info(element):
    try:
        type_id = element.GetTypeId()
        if not type_id or type_id == ElementId.InvalidElementId:
            return
        etype = doc.GetElement(type_id)
        if not etype:
            return
    except Exception:
        return

    out.print_md("## Element Type")
    rows = [
        ["Type class", type(etype).__name__],
        ["Type name",  getattr(etype, "Name", "-") or "-"],
        ["Type ID",    str(etype.Id)],
    ]
    out.print_table(rows, columns=["Property", "Value"])
    out.print_md("### Type Parameters")
    trows = dump_params(etype)
    out.print_table(trows or [["-", "", ""]], columns=["Name", "Param ID", "Value"])

def section_family_info(element):
    if not isinstance(element, FamilyInstance):
        return
    out.print_md("## Family Info")
    rows = []
    try:
        rows.append(["Family", element.Symbol.Family.Name])
        rows.append(["Symbol", element.Symbol.Name])
    except Exception:
        pass
    try:
        host = element.Host
        rows.append(["Host class", type(host).__name__ if host else "-"])
        rows.append(["Host ID",    str(host.Id) if host else "-"])
    except Exception:
        pass
    try:
        rows.append(["Hand Flipped",   str(element.HandFlipped)])
        rows.append(["Facing Flipped", str(element.FacingFlipped)])
    except Exception:
        pass
    try:
        rows.append(["Mirrored", str(element.Mirrored)])
    except Exception:
        pass
    out.print_table(rows or [["-", ""]], columns=["Property", "Value"])

def section_location(element):
    loc = getattr(element, "Location", None)
    if not loc:
        return
    out.print_md("## Location and Geometry")
    rows = []
    if isinstance(loc, LocationPoint):
        rows.append(["Location type", "Point"])
        rows.append(["Position",      xyz_str(loc.Point)])
        try:
            rows.append(["Rotation (rad)", "{:.4f}".format(loc.Rotation)])
        except Exception:
            pass
    elif isinstance(loc, LocationCurve):
        curve = loc.Curve
        rows.append(["Location type", "Curve"])
        rows.append(["Start point",   xyz_str(curve.GetEndPoint(0))])
        rows.append(["End point",     xyz_str(curve.GetEndPoint(1))])
        try:
            rows.append(["Length", "{:.3f} ft".format(curve.Length)])
        except Exception:
            pass
    try:
        bb = element.get_BoundingBox(None)
        if bb:
            rows.append(["BBox Min", xyz_str(bb.Min)])
            rows.append(["BBox Max", xyz_str(bb.Max)])
    except Exception:
        pass
    out.print_table(rows or [["-", ""]], columns=["Property", "Value"])

def section_mep(element):
    try:
        mep = getattr(element, "MEPModel", None)
        if not mep:
            return
        cm = mep.ConnectorManager
        if not cm:
            return
    except Exception:
        return

    out.print_md("## MEP Connectors")
    rows = []
    try:
        for conn in cm.Connectors:
            try:
                sys_type = str(conn.MEPSystemType) if hasattr(conn, "MEPSystemType") else "-"
                domain   = str(conn.Domain)        if hasattr(conn, "Domain")        else "-"
                rows.append([str(conn.Id), domain, sys_type, xyz_str(conn.Origin)])
            except Exception:
                pass
    except Exception:
        pass
    out.print_table(
        rows or [["-", "", "", ""]],
        columns=["Connector ID", "Domain", "System Type", "Origin"]
    )

def section_instance_params(element):
    out.print_md("## Instance Parameters")
    rows = dump_params(element)
    out.print_table(rows or [["-", "", ""]], columns=["Name", "Param ID", "Value"])

def section_filterable_params(element, cat):
    out.print_md("## Filterable Parameters for `{}`".format(cat.Name))
    try:
        filterable = ParameterFilterElement.GetAllFilterableParams(
            doc, List[ElementId]([cat.Id]))
        frows = []
        for fid in filterable:
            p = element.get_Parameter(fid)
            frows.append([bip_name(fid), str(fid), param_value(p) if p else "<no value>"])
        out.print_table(
            frows or [["-", "", ""]],
            columns=["BIP Name", "Param ID", "Value"]
        )
    except Exception as ex:
        out.print_md("**GetAllFilterableParams failed:** `{}`".format(ex))

# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    try:
        ref     = uidoc.Selection.PickObject(
            ObjectType.Element, AnyFilter(), "Pick an element")
        element = doc.GetElement(ref.ElementId)
    except Exception as ex:
        if "cancelled" not in str(ex).lower():
            forms.alert(str(ex), title="Inspector")
        return

    cat = getattr(element, "Category", None)

    out.print_md("# Inspector")
    section_identity(element, cat)
    section_type_info(element)
    section_family_info(element)
    section_location(element)
    section_mep(element)
    section_instance_params(element)
    if cat:
        section_filterable_params(element, cat)

if __name__ == "__main__":
    main()
