# -*- coding: utf-8 -*-
"""Seed43 Studio .seed writer, the same file Navisworks' Export to Studio makes.

The contract is STUDIO-NAVISWORKS-IMPORT.md in the Seed43 Site project:

    SEED43-STUDIO-V1 text header, sourceFormat navis, then
    gzip( "S43NAVIS1\\n" + uint32 json length + json + pad to 4
          + float32 positions (mm, Z up, minus origin) + uint32 indices )

No Revit types in here, so it runs under CPython for testing as well as
IronPython 2 inside Revit.
"""
import json
import re
import struct
import time
import zlib

__all__ = ["SeedNode", "SeedItem", "SeedPart", "SeedScene"]

CHUNK = 60000       # values per struct.pack call, keeps the argument list small


class SeedNode(object):
    def __init__(self, name, cls, parent=-1, hidden=False):
        self.name = name
        self.cls = cls
        self.parent = parent
        self.hidden = hidden


class SeedPart(object):
    """One piece of an item's geometry: absolute mm positions, local indices."""

    def __init__(self, rgb, opacity=1.0, mat=u""):
        self.rgb = rgb
        self.opacity = opacity
        self.mat = mat
        self.positions = []
        self.triangles = []

    def add(self, positions, triangles):
        base = len(self.positions) // 3
        self.positions.extend(positions)
        if base:
            self.triangles.extend([i + base for i in triangles])
        else:
            self.triangles.extend(triangles)


class SeedItem(object):
    def __init__(self, node):
        self.node = node
        self.hidden = False
        self.parts = []
        self.cat = u""
        self.type = u""
        self.mat = u""
        self.props = None


def _hex(rgb):
    return "#" + "".join("%02x" % int(round(max(0.0, min(1.0, c)) * 255))
                         for c in rgb)


_NON_ASCII = re.compile(u"[^\x00-\x7f]")


def _escape(match):
    """A JSON \\u escape; characters past U+FFFF become a surrogate pair
    (IronPython strings already hold them as two UTF-16 units)."""
    code = ord(match.group(0))
    if code > 0xFFFF:
        code -= 0x10000
        return u"\\u%04x\\u%04x" % (0xD800 + (code >> 10), 0xDC00 + (code & 0x3FF))
    return u"\\u%04x" % code


def _r4(v):
    return [round(v[0], 4), round(v[1], 4), round(v[2], 4)]


class SeedScene(object):
    def __init__(self, document=u"Revit", product=u""):
        self.document = document
        self.product = product
        self.scope = "all"
        self.origin = (0.0, 0.0, 0.0)
        self.view = None    # (position, direction, up, perspective), position absolute mm
        self.nodes = []
        self.items = []

    def add_node(self, name, cls, parent=-1):
        self.nodes.append(SeedNode(name, cls, parent))
        return len(self.nodes) - 1

    def counts(self):
        v = t = 0
        for it in self.items:
            for p in it.parts:
                v += len(p.positions) // 3
                t += len(p.triangles) // 3
        return v, t

    def check_tree_order(self):
        """What breaks Studio's tree-order rule, or None (see the contract)."""
        n = len(self.nodes)
        for i, node in enumerate(self.nodes):
            if node.parent >= i:
                return "node %d comes before its parent" % i
        size = [1] * n
        end = list(range(n))
        for i in range(n - 1, -1, -1):
            p = self.nodes[i].parent
            if p >= 0:
                size[p] += size[i]
                if end[i] > end[p]:
                    end[p] = end[i]
        for i in range(n):
            if end[i] - i + 1 != size[i]:
                return "branch of node %d is not one run" % i
        for j in range(1, len(self.items)):
            if self.items[j].node <= self.items[j - 1].node:
                return "item %d is out of tree order" % j
        return None

    # -- writing ------------------------------------------------------------

    def _json(self, vcount, tcount):
        nodes = [{"n": n.name, "c": n.cls, "p": n.parent, "h": n.hidden}
                 for n in self.nodes]
        items = []
        first = 0
        for it in self.items:
            parts = []
            for p in it.parts:
                count = len(p.triangles) // 3
                d = {"t": [first, count], "rgb": _hex(p.rgb)}
                first += count
                if p.opacity < 0.999:
                    d["op"] = round(p.opacity, 3)
                if p.mat and p.mat != it.mat:
                    d["mat"] = p.mat
                parts.append(d)
            o = {"node": it.node, "parts": parts, "h": it.hidden}
            if it.cat:
                o["cat"] = it.cat
            if it.type:
                o["type"] = it.type
            if it.mat:
                o["mat"] = it.mat
            if it.props:
                o["props"] = it.props
            items.append(o)
        root = {
            "format": "seed43-navisworks", "version": 1,
            "document": self.document,
            "exported": time.strftime("%Y-%m-%dT%H:%M:%S"),
            "product": self.product, "scope": self.scope,
            "units": "mm", "upAxis": "Z", "origin": list(self.origin),
            "vertexCount": vcount, "triangleCount": tcount,
            "nodes": nodes, "items": items,
        }
        if self.view is not None:
            pos, direction, up, persp = self.view
            o = self.origin
            root["view"] = {
                "position": _r4((pos[0] - o[0], pos[1] - o[1], pos[2] - o[2])),
                "direction": _r4(direction), "up": _r4(up),
                "perspective": bool(persp)}
        # ensure_ascii=True breaks on IronPython 2 (its encoder tries to
        # decode text that is already unicode, e.g. a "m³" value), so the
        # non-ASCII characters are escaped here instead. ASCII JSON is valid
        # UTF-8 for Studio.
        text = json.dumps(root, separators=(",", ":"), ensure_ascii=False)
        return _NON_ASCII.sub(_escape, text).encode("ascii")

    def _chunks(self, vcount, tcount, progress=None):
        """The uncompressed S43NAVIS1 data, a piece at a time."""
        js = self._json(vcount, tcount)
        head = b"S43NAVIS1\n" + struct.pack("<I", len(js)) + js
        pad = (-len(head)) % 4
        yield head + b"\x00" * pad
        ox, oy, oz = self.origin
        total = len(self.items) * 2
        step = 0
        for it in self.items:
            for p in it.parts:
                src = p.positions
                vals = []
                for i in range(0, len(src), 3):
                    vals.append(src[i] - ox)
                    vals.append(src[i + 1] - oy)
                    vals.append(src[i + 2] - oz)
                for k in range(0, len(vals), CHUNK):
                    piece = vals[k:k + CHUNK]
                    yield struct.pack("<%df" % len(piece), *piece)
            step += 1
            if progress is not None and step % 200 == 0:
                progress(step, total)
        base = 0
        for it in self.items:
            for p in it.parts:
                tri = p.triangles
                if base:
                    tri = [i + base for i in tri]
                for k in range(0, len(tri), CHUNK):
                    piece = tri[k:k + CHUNK]
                    yield struct.pack("<%dI" % len(piece), *piece)
                base += len(p.positions) // 3
            step += 1
            if progress is not None and step % 200 == 0:
                progress(step, total)

    def source_bytes(self, progress=None):
        """gzip of the S43NAVIS1 data, built by hand from raw deflate so it
        does not depend on the gzip module or zlib's gzip wbits mode."""
        vcount, tcount = self.counts()
        comp = zlib.compressobj(6, zlib.DEFLATED, -15)
        out = [b"\x1f\x8b\x08\x00\x00\x00\x00\x00\x00\xff"]
        crc = 0
        size = 0
        for chunk in self._chunks(vcount, tcount, progress):
            crc = zlib.crc32(chunk, crc)
            size += len(chunk)
            out.append(comp.compress(chunk))
        out.append(comp.flush())
        out.append(struct.pack("<II", crc & 0xffffffff, size & 0xffffffff))
        return b"".join(out)

    def write(self, path, source_name, progress=None):
        src = self.source_bytes(progress)
        name = source_name.replace(u"\r", u"").replace(u"\n", u"")
        head = (u"SEED43-STUDIO-V1\nsourceFormat: navis\nsourceName: %s\n"
                u"sourceBytes: %d\n---BEGIN-SOURCE-DATA---\n"
                % (name, len(src))).encode("utf-8")
        tail = (b"\n---END-SOURCE-DATA---\n"
                b'{"paramDefs":[],"itemParams":[],"dimensions":[]}')
        with open(path, "wb") as f:
            f.write(head)
            f.write(src)
            f.write(tail)
        return len(head) + len(src) + len(tail)
