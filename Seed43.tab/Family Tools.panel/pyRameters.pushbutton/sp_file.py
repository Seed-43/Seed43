# -*- coding: utf-8 -*-
# """
# Revit shared parameter file: read, edit, write.
#
# No WPF in here on purpose. Everything that can corrupt a parameter file
# lives in this module, so it can be tested under desktop CPython without
# launching Revit. Written to run unchanged on IronPython 2 and CPython 3.
#
# FORMAT, CONFIRMED AGAINST A REVIT WRITTEN FILE
#     encoding    UTF-16 LE with BOM. UTF-8 and cp1252 are read as well.
#     newline     CRLF
#     delimiter   TAB
#     comments    lines starting with '#', tab padded to the table width
#     sections    '*META', '*GROUP', '*PARAM' name the columns; data lines
#                 start with 'META', 'GROUP', 'PARAM'
#
# THE COLUMN ORDER IS READ, NEVER ASSUMED
#     A MINVERSION 1 file has no USERMODIFIABLE and no HIDEWHENNOVALUE
#     column. Parsing by position instead of by the '*PARAM' header puts
#     the description into the visibility field on those files.
#
# VALUES STAY AS RAW STRINGS
#     A field this tool does not understand survives a round trip untouched
#     unless the user edits that cell. The Nagel file carries VISIBLE=5 on
#     two rows, which is neither 0 nor 1 and is not a value to flatten.
#     Anything past the declared columns is kept as well.
#
# Target: IronPython 2 and CPython 3.
# """

import os
import re
import shutil
from datetime import datetime

__all__ = [
    "SharedParameterFile", "ParseError",
    "PARAM_COLUMNS", "GROUP_COLUMNS", "META_COLUMNS",
    "new_guid", "is_guid", "DATA_TYPES", "DATA_TYPE_GROUPS",
]


# -- CONSTANTS --------------------------------------------------------------

META_COLUMNS = ["VERSION", "MINVERSION"]
GROUP_COLUMNS = ["ID", "NAME"]
PARAM_COLUMNS = ["GUID", "NAME", "DATATYPE", "DATACATEGORY", "GROUP",
                 "VISIBLE", "DESCRIPTION", "USERMODIFIABLE", "HIDEWHENNOVALUE"]

DEFAULT_COMMENTS = ["# This is a Revit shared parameter file.",
                    "# Do not edit manually."]

EXTRA_KEY = "_EXTRA"

_GUID_RE = re.compile(
    r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}"
    r"-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")

# Curated DATATYPE tokens for the picker. The combo is editable and whatever
# the loaded file already uses is merged in, so a token missing from here is
# a missing convenience, never a lost value.
DATA_TYPE_GROUPS = [
    ("Common", [
        "TEXT", "MULTILINETEXT", "INTEGER", "NUMBER", "LENGTH", "AREA",
        "VOLUME", "ANGLE", "SLOPE", "YESNO", "MATERIAL", "IMAGE",
        "FAMILYTYPE", "URL", "CURRENCY"]),
    ("Structural", [
        "FORCE", "LINEAR_FORCE", "AREA_FORCE", "MOMENT", "LINEAR_MOMENT",
        "STRESS", "UNIT_WEIGHT", "WEIGHT", "WEIGHT_PER_UNIT_LENGTH",
        "MASS_DENSITY", "MASS_PER_UNIT_LENGTH", "THERMAL_EXPANSION",
        "POINT_SPRING_COEFFICIENT", "ROTATIONAL_POINT_SPRING_COEFFICIENT",
        "LINE_SPRING_COEFFICIENT", "ROTATIONAL_LINE_SPRING_COEFFICIENT",
        "AREA_SPRING_COEFFICIENT", "SECTION_AREA", "SECTION_DIMENSION",
        "SECTION_MODULUS", "SECTION_PROPERTY", "MOMENT_OF_INERTIA",
        "WARPING_CONSTANT", "SURFACE_AREA", "PERIOD", "PULSATION",
        "REINFORCEMENT_AREA", "REINFORCEMENT_AREA_PER_UNIT_LENGTH",
        "REINFORCEMENT_COVER", "REINFORCEMENT_LENGTH",
        "REINFORCEMENT_SPACING", "REINFORCEMENT_VOLUME", "BAR_DIAMETER",
        "CRACK_WIDTH", "DISPLACEMENT_DEFLECTION", "ROTATION"]),
    ("HVAC", [
        "HVAC_DENSITY", "HVAC_ENERGY", "HVAC_FRICTION", "HVAC_POWER",
        "HVAC_POWER_DENSITY", "HVAC_PRESSURE", "HVAC_TEMPERATURE",
        "HVAC_VELOCITY", "HVAC_AIRFLOW", "HVAC_DUCTSIZE",
        "HVAC_CROSSSECTION", "HVAC_HEATGAIN", "HVAC_THERMAL_RESISTANCE",
        "HVAC_COEFFICIENT_OF_HEAT_TRANSFER", "HVAC_THERMAL_MASS",
        "HVAC_THERMAL_CONDUCTIVITY", "HVAC_SPECIFIC_HEAT",
        "HVAC_SPECIFIC_HEAT_OF_VAPORIZATION", "HVAC_PERMEABILITY",
        "HVAC_VISCOSITY", "HVAC_AIRFLOW_DENSITY", "HVAC_SLOPE",
        "HVAC_COOLING_LOAD", "HVAC_HEATING_LOAD", "HVAC_ROUGHNESS",
        "HVAC_DUCT_INSULATION_THICKNESS", "HVAC_DUCT_LINING_THICKNESS",
        "HVAC_FACTOR"]),
    ("Electrical", [
        "ELECTRICAL_CURRENT", "ELECTRICAL_POTENTIAL",
        "ELECTRICAL_FREQUENCY", "ELECTRICAL_ILLUMINANCE",
        "ELECTRICAL_LUMINOUS_FLUX", "ELECTRICAL_LUMINOUS_INTENSITY",
        "ELECTRICAL_LUMINANCE", "ELECTRICAL_EFFICACY",
        "ELECTRICAL_WATTAGE", "ELECTRICAL_POWER",
        "ELECTRICAL_POWER_DENSITY", "ELECTRICAL_APPARENT_POWER",
        "ELECTRICAL_RESISTIVITY", "ELECTRICAL_TEMPERATURE",
        "ELECTRICAL_CABLETRAYSIZE", "ELECTRICAL_CONDUITSIZE",
        "ELECTRICAL_DEMAND_FACTOR", "ELECTRICAL_POWER_PER_LENGTH",
        "COLOR_TEMPERATURE", "WIRE_DIAMETER"]),
    ("Piping", [
        "PIPING_DENSITY", "PIPING_FLOW", "PIPING_FRICTION",
        "PIPING_PRESSURE", "PIPING_TEMPERATURE", "PIPING_VELOCITY",
        "PIPING_VISCOSITY", "PIPING_VOLUME", "PIPING_SLOPE",
        "PIPING_ROUGHNESS", "PIPING_MASS_PER_TIME", "PIPE_DIMENSION",
        "PIPE_INSULATION_THICKNESS", "PIPE_MASS",
        "PIPE_MASS_PER_UNIT_LENGTH", "PIPE_SIZE", "CROSSSECTIONALAREA"]),
    ("Other", [
        "ENERGY", "SPEED", "TIME", "DISTANCE", "ACCELERATION",
        "STATIONING", "COST_PER_AREA", "COST_RATE_ENERGY",
        "COST_RATE_POWER", "APPARENT_POWER_DENSITY", "DIFFUSIVITY",
        "MASS", "MASSPERUNITAREA", "MOISTURE_CONTENT", "NUMBEROFPOLES",
        "PERCENTAGE", "SITE_ANGLE", "THERMAL_MASS",
        "ISOTHERMAL_MOISTURE_CAPACITY", "VOLUMETRIC_FLOW_DENSITY"]),
]

DATA_TYPES = []
for _label, _items in DATA_TYPE_GROUPS:
    DATA_TYPES.extend(_items)


class ParseError(Exception):
    """Not a shared parameter file, or unreadable."""


# -- SMALL HELPERS ----------------------------------------------------------

def new_guid():
    """A fresh lowercase GUID in the form Revit writes."""
    try:
        import uuid
        return str(uuid.uuid4())
    except Exception:
        # IronPython leans on ctypes for uuid, which is not always happy
        # inside Revit. System.Guid is always there.
        from System import Guid
        return Guid.NewGuid().ToString()


def is_guid(text):
    return bool(_GUID_RE.match((text or "").strip()))


def _decode(raw):
    """Return (text, encoding_name). A BOM wins, then UTF-8, then cp1252."""
    if raw[:2] == b"\xff\xfe":
        return raw.decode("utf-16"), "utf-16"
    if raw[:2] == b"\xfe\xff":
        return raw.decode("utf-16"), "utf-16"
    if raw[:3] == b"\xef\xbb\xbf":
        return raw.decode("utf-8-sig"), "utf-8-sig"
    try:
        return raw.decode("utf-8"), "utf-8"
    except UnicodeDecodeError:
        return raw.decode("cp1252", "replace"), "cp1252"


def _atomic_replace(tmp, target):
    """Move tmp onto target without ever leaving a half written file.

    os.replace is CPython 3 only, so IronPython 2 goes through .NET, which
    carries the same guarantee. The remove-then-rename branch is a last
    resort and the only one with a window where the file is missing.
    """
    try:
        os.replace(tmp, target)
        return
    except AttributeError:
        pass
    try:
        from System.IO import File
        if os.path.exists(target):
            File.Replace(tmp, target, None)
        else:
            File.Move(tmp, target)
        return
    except ImportError:
        pass
    if os.path.exists(target):
        os.remove(target)
    os.rename(tmp, target)


def _row_to_dict(fields, columns):
    row = {}
    for i, col in enumerate(columns):
        row[col] = fields[i].strip() if i < len(fields) else ""
    extra = [f.strip() for f in fields[len(columns):] if f.strip()]
    if extra:
        row[EXTRA_KEY] = "\x1f".join(extra)
    return row


# -- THE FILE ---------------------------------------------------------------

class SharedParameterFile(object):
    """An in-memory shared parameter file.

    ``params`` is a list of dicts keyed by the PARAM column names, ``groups``
    a list of dicts keyed by the GROUP column names. Both hold raw strings.
    """

    def __init__(self):
        self.path = None
        self.encoding = "utf-16"
        self.newline = "\r\n"
        self.comments = list(DEFAULT_COMMENTS)
        self.meta_columns = list(META_COLUMNS)
        self.meta = {"VERSION": "2", "MINVERSION": "1"}
        self.group_columns = list(GROUP_COLUMNS)
        self.groups = []
        self.param_columns = list(PARAM_COLUMNS)
        self.params = []

    # --- loading ---

    @classmethod
    def load(cls, path):
        fh = open(path, "rb")
        try:
            raw = fh.read()
        finally:
            fh.close()
        self = cls.loads(raw)
        self.path = path
        return self

    @classmethod
    def loads(cls, raw):
        text, encoding = _decode(raw)
        self = cls()
        self.encoding = encoding
        self.newline = "\r\n" if "\r\n" in text else "\n"
        self.comments = []
        self.groups = []
        self.params = []

        meta_cols = None
        group_cols = None
        param_cols = None
        seen_section = False

        for line in text.splitlines():
            if not line.replace("\t", "").strip():
                continue

            stripped = line.strip()
            if stripped.startswith("#"):
                if not seen_section:
                    self.comments.append(stripped.rstrip("\t").rstrip())
                continue

            fields = line.split("\t")
            token = fields[0].strip()

            if token.startswith("*"):
                seen_section = True
                cols = [f.strip() for f in fields[1:]]
                while cols and not cols[-1]:
                    cols.pop()
                name = token[1:].upper()
                if name == "META":
                    meta_cols = cols or list(META_COLUMNS)
                elif name == "GROUP":
                    group_cols = cols or list(GROUP_COLUMNS)
                elif name == "PARAM":
                    param_cols = cols or list(PARAM_COLUMNS)
                continue

            upper = token.upper()
            if upper == "META":
                self.meta = _row_to_dict(fields[1:], meta_cols or META_COLUMNS)
            elif upper == "GROUP":
                self.groups.append(
                    _row_to_dict(fields[1:], group_cols or GROUP_COLUMNS))
            elif upper == "PARAM":
                self.params.append(
                    _row_to_dict(fields[1:], param_cols or PARAM_COLUMNS))

        if param_cols is None and not self.params:
            raise ParseError(
                "No *PARAM section found. This does not look like a Revit "
                "shared parameter file.")

        self.meta_columns = meta_cols or list(META_COLUMNS)
        self.group_columns = group_cols or list(GROUP_COLUMNS)
        self.param_columns = param_cols or list(PARAM_COLUMNS)

        # Guarantee the canonical columns exist so the editor can rely on
        # them. A file written by Revit 2014 has no HIDEWHENNOVALUE.
        for col in PARAM_COLUMNS:
            if col not in self.param_columns:
                self.param_columns.append(col)
        for row in self.params:
            for col in self.param_columns:
                if col not in row:
                    row[col] = ""
        for col in GROUP_COLUMNS:
            if col not in self.group_columns:
                self.group_columns.append(col)
        for row in self.groups:
            for col in self.group_columns:
                if col not in row:
                    row[col] = ""

        if not self.comments:
            self.comments = list(DEFAULT_COMMENTS)
        return self

    # --- writing ---

    def dumps(self):
        width = 1 + len(self.param_columns)

        def line(fields):
            padded = list(fields) + [""] * (width - len(fields))
            return "\t".join(padded)

        out = []
        for comment in self.comments:
            out.append(line([comment]))
        out.append(line(["*META"] + self.meta_columns))
        out.append(line(["META"] +
                        [self.meta.get(c, "") for c in self.meta_columns]))
        out.append(line(["*GROUP"] + self.group_columns))
        for g in self.groups:
            out.append(line(["GROUP"] + self._values(g, self.group_columns)))
        out.append(line(["*PARAM"] + self.param_columns))
        for p in self.params:
            out.append(line(["PARAM"] + self._values(p, self.param_columns)))
        return self.newline.join(out) + self.newline

    @staticmethod
    def _values(row, columns):
        values = [row.get(c, "") for c in columns]
        extra = row.get(EXTRA_KEY, "")
        if extra:
            values.extend(extra.split("\x1f"))
        return values

    def save(self, path=None, backup=True):
        """Write atomically. Returns the path written.

        Writing a temp file and moving it into place is what stops a crash
        part way through leaving a truncated parameter file: either the old
        file is intact or the new one is complete, never something between.
        """
        target = path or self.path
        if not target:
            raise ValueError("No path to save to.")
        target = os.path.abspath(target)
        data = self.dumps().encode(self.encoding)

        folder = os.path.dirname(target) or "."
        if not os.path.isdir(folder):
            os.makedirs(folder)

        if backup and os.path.exists(target):
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            bak_dir = os.path.join(folder, "_backups")
            if not os.path.isdir(bak_dir):
                os.makedirs(bak_dir)
            base = os.path.basename(target)
            shutil.copy2(
                target, os.path.join(bak_dir, "{0}.{1}.bak".format(base, stamp)))

        tmp = target + ".tmp"
        fh = open(tmp, "wb")
        try:
            fh.write(data)
            fh.flush()
            try:
                os.fsync(fh.fileno())
            except Exception:
                pass
        finally:
            fh.close()
        _atomic_replace(tmp, target)
        self.path = target
        return target

    # --- groups ---

    def group_names(self):
        names = {}
        for g in self.groups:
            names[g.get("ID", "")] = g.get("NAME", "")
        return names

    def name_for_group(self, group_id):
        return self.group_names().get(group_id, "")

    def next_group_id(self):
        used = set()
        for g in self.groups:
            try:
                used.add(int(g.get("ID", "0")))
            except ValueError:
                pass
        n = 1
        while n in used:
            n += 1
        return str(n)

    def add_group(self, name):
        row = {}
        for c in self.group_columns:
            row[c] = ""
        row["ID"] = self.next_group_id()
        row["NAME"] = name
        self.groups.append(row)
        return row

    def group_usage(self, group_id):
        return len([p for p in self.params if p.get("GROUP", "") == group_id])

    # --- parameters ---

    def blank_param(self, name="", group_id=""):
        row = {}
        for c in self.param_columns:
            row[c] = ""
        row["GUID"] = new_guid()
        row["NAME"] = name
        row["DATATYPE"] = "TEXT"
        row["GROUP"] = group_id or (self.groups[0]["ID"] if self.groups else "1")
        row["VISIBLE"] = "1"
        row["USERMODIFIABLE"] = "1"
        row["HIDEWHENNOVALUE"] = "0"
        return row

    def unique_name(self, stem="Parameter"):
        taken = set([p.get("NAME", "") for p in self.params])
        n = 1
        while True:
            candidate = "{0} {1:03d}".format(stem, n)
            if candidate not in taken:
                return candidate
            n += 1

    # --- validation ---

    def validate(self):
        """Return (errors, warnings). An error must block the save."""
        errors, warnings, _rows = self.validate_rows()
        return errors, warnings

    def validate_rows(self):
        """Return (errors, warnings, row_errors).

        row_errors maps a position in ``params`` to the messages against
        that row, so the caller can mark the offending rows without having
        to parse the message text back apart.
        """
        errors = []
        warnings = []
        row_errors = {}
        seen_guid = {}
        seen_name = {}
        valid_groups = set(self.group_names().keys())

        def fault(index, text):
            errors.append(text)
            row_errors.setdefault(index, []).append(text)

        for i, p in enumerate(self.params):
            row_no = i + 1
            name = p.get("NAME", "")
            guid = p.get("GUID", "")
            label = name or "row {0}".format(row_no)

            if not name.strip():
                fault(i, "Row {0}: the name is empty.".format(row_no))

            for field in ("NAME", "DESCRIPTION", "DATATYPE"):
                value = p.get(field, "")
                if "\t" in value or "\n" in value or "\r" in value:
                    fault(i,
                          "{0}: {1} holds a tab or a line break, which would "
                          "corrupt the file.".format(label, field))

            if not is_guid(guid):
                fault(i, "{0}: '{1}' is not a valid GUID.".format(label, guid))
            else:
                key = guid.lower()
                if key in seen_guid:
                    fault(i,
                          "{0}: the GUID repeats row {1}. Revit reads two "
                          "parameters sharing a GUID as one parameter."
                          .format(label, seen_guid[key]))
                else:
                    seen_guid[key] = row_no

            if not p.get("DATATYPE", "").strip():
                fault(i, "{0}: no data type.".format(label))

            group_id = p.get("GROUP", "")
            if group_id not in valid_groups:
                fault(i,
                      "{0}: group id '{1}' is not in the group table."
                      .format(label, group_id))

            if name.strip():
                key = name.strip().lower()
                if key in seen_name:
                    warnings.append(
                        "'{0}' is on rows {1} and {2}."
                        .format(name, seen_name[key], row_no))
                else:
                    seen_name[key] = row_no

            if (p.get("DATATYPE", "") == "FAMILYTYPE"
                    and not p.get("DATACATEGORY", "").strip()):
                warnings.append(
                    "'{0}' is a FAMILYTYPE with no category id, so Revit "
                    "will not know which family to offer.".format(name))

        group_seen = {}
        for g in self.groups:
            gid = g.get("ID", "")
            gname = g.get("NAME", "")
            if not gid.strip():
                errors.append("Group '{0}' has no id.".format(gname))
            elif gid in group_seen:
                errors.append("Group id {0} is used twice.".format(gid))
            group_seen[gid] = gname
            if not gname.strip():
                errors.append("Group {0} has no name.".format(gid))
            if "\t" in gname or "\n" in gname:
                errors.append(
                    "Group '{0}' holds a tab or a line break.".format(gname))

        return errors, warnings, row_errors
