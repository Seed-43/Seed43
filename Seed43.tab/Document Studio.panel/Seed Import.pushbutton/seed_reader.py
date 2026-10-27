# -*- coding: utf-8 -*-
"""Seed43 Studio .seed reader for mesh files (sourceFormat navis).

Reads what 3D Seed and Navisworks' Export to Studio write (the contract is
STUDIO-NAVISWORKS-IMPORT.md in the Seed43 Site project):

    SEED43-STUDIO-V1 text header, then
    gzip( "S43NAVIS1\\n" + uint32 json length + json + pad to 4
          + float32 positions (mm, Z up, minus origin) + uint32 indices )

A .seed made from another source (bsw, ifc, step...) carries that original
file instead of a mesh, which only Studio can read; read_seed says so.

No Revit types in here, so it runs under CPython for testing as well as
IronPython 2 inside Revit.
"""
import array
import json
import struct
import zlib

__all__ = ["SeedError", "SeedModel", "read_seed"]

BEGIN = b"---BEGIN-SOURCE-DATA---\n"
MAGIC = b"S43NAVIS1\n"


class SeedError(Exception):
    pass


def _typed(code, data):
    """array of 4-byte items from raw little-endian bytes."""
    arr = array.array(code)
    if arr.itemsize != 4:
        raise SeedError("array '{}' is not 4 bytes here".format(code))
    if hasattr(arr, "frombytes"):
        arr.frombytes(data)
    else:
        arr.fromstring(data)    # IronPython 2
    return arr


def _gunzip(data):
    """One gzip member, read by hand so it does not lean on zlib's gzip
    wbits mode (IronPython 2's zlib is a port and not every mode exists)."""
    if len(data) < 18 or data[0:2] != b"\x1f\x8b":
        raise SeedError("the model data is not gzip")
    flags = ord(data[3:4])
    pos = 10
    if flags & 4:                                   # FEXTRA
        pos += 2 + struct.unpack("<H", data[pos:pos + 2])[0]
    if flags & 8:                                   # FNAME
        pos = data.index(b"\x00", pos) + 1
    if flags & 16:                                  # FCOMMENT
        pos = data.index(b"\x00", pos) + 1
    if flags & 2:                                   # FHCRC
        pos += 2
    return zlib.decompressobj(-15).decompress(data[pos:])


class SeedModel(object):
    """The JSON (meta), plus positions and triangle indices as flat arrays."""

    def __init__(self, meta, positions, indices, source_format, source_name):
        self.meta = meta
        self.positions = positions      # x, y, z per vertex, mm minus origin
        self.indices = indices          # 3 per triangle, into all vertices
        self.source_format = source_format
        self.source_name = source_name
        self.nodes = meta.get("nodes") or []
        self.items = meta.get("items") or []
        self.origin = meta.get("origin") or [0.0, 0.0, 0.0]

    def path(self, node):
        """Names from the model root down to (not including) node."""
        names = []
        k = self.nodes[node].get("p", -1) if 0 <= node < len(self.nodes) else -1
        while k >= 0:
            names.append(self.nodes[k].get("n", u""))
            k = self.nodes[k].get("p", -1)
        names.reverse()
        return names

    def hidden(self, item):
        """Hidden on the item or on any node above it."""
        if item.get("h"):
            return True
        k = item.get("node", -1)
        while 0 <= k < len(self.nodes):
            if self.nodes[k].get("h"):
                return True
            k = self.nodes[k].get("p", -1)
        return False


def read_seed(path):
    with open(path, "rb") as f:
        data = f.read()
    if not data.startswith(b"SEED43-STUDIO-V1"):
        raise SeedError("not a Seed43 Studio .seed file")
    start = data.find(BEGIN)
    if start < 0:
        raise SeedError("the file has no model data")
    fields = {}
    for line in data[:start].decode("utf-8", "replace").splitlines()[1:]:
        if u":" in line:
            key, value = line.split(u":", 1)
            fields[key.strip()] = value.strip()
    fmt = fields.get(u"sourceFormat", u"")
    name = fields.get(u"sourceName", u"")
    if fmt != u"navis":
        raise SeedError(
            u"this .seed holds a {} model ({}), which only Studio can read. "
            u"Seed Import reads the mesh files made by 3D Seed in Revit and "
            u"Export to Studio in Navisworks.".format(fmt or u"unknown", name))
    try:
        size = int(fields.get(u"sourceBytes", u""))
    except ValueError:
        raise SeedError("the header has no data size")
    start += len(BEGIN)
    raw = _gunzip(data[start:start + size])
    data = None
    if raw[:10] != MAGIC:
        raise SeedError("the model data is not a Seed43 mesh")
    jlen = struct.unpack("<I", raw[10:14])[0]
    meta = json.loads(raw[14:14 + jlen].decode("utf-8"))
    pos = 14 + jlen
    pos += (-pos) % 4
    vcount = int(meta.get("vertexCount", 0))
    tcount = int(meta.get("triangleCount", 0))
    end_v = pos + vcount * 12
    end_t = end_v + tcount * 12
    if end_t > len(raw):
        raise SeedError("the model data is shorter than its header says")
    positions = _typed("f", raw[pos:end_v])
    indices = _typed("I", raw[end_v:end_t])
    if struct.pack("<I", 1) != struct.pack("=I", 1):
        positions.byteswap()
        indices.byteswap()
    return SeedModel(meta, positions, indices, fmt, name)
