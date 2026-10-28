# -*- coding: utf-8 -*-
"""Minimal U3D (ECMA-363) writer: triangle meshes, colours, a node tree.

Writes what a 3D PDF needs and nothing more. Each model is one CLOD mesh
with every position in the base mesh (no progressive refinement), one
shader per colour, and group nodes that become the model tree in Acrobat.

The bit stream follows the Intel U3D reference library exactly: plain
values are raw little-endian bytes until the first compressed symbol, then
everything goes through the arithmetic coder. Runs under CPython 3 for
tests and IronPython 2.7 inside Revit, so no Python 3 only syntax.
"""
import struct

__all__ = ["BitStream", "U3DScene"]

# ── CONSTANTS ───────────────────────────────────────────────────────────────

STATIC_FULL = 0x400
MAX_RANGE = STATIC_FULL + 0x3FFF
ELEPHANT = 0x1FFF
HIST_MAX_SYMBOL = 0xFFFF
CONTEXT_BASE_SHADING_ID = 1

BLOCK_FILE_HEADER = 0x00443355
BLOCK_MODIFIER_CHAIN = 0xFFFFFF14
BLOCK_GROUP_NODE = 0xFFFFFF21
BLOCK_MODEL_NODE = 0xFFFFFF22
BLOCK_MESH_DECLARATION = 0xFFFFFF31
BLOCK_MESH_CONTINUATION = 0xFFFFFF3B
BLOCK_SHADING_MODIFIER = 0xFFFFFF45
BLOCK_LIT_TEXTURE_SHADER = 0xFFFFFF53
BLOCK_MATERIAL = 0xFFFFFF54

PROFILE_DEFINED_UNITS = 0x00000008
CHAR_ENCODING_UTF8 = 106

IDENTITY = (1.0, 0.0, 0.0, 0.0,
            0.0, 1.0, 0.0, 0.0,
            0.0, 0.0, 1.0, 0.0,
            0.0, 0.0, 0.0, 1.0)

_SWAP8 = [int("{:08b}".format(i)[::-1], 2) for i in range(256)]

# Bit length of every 16-bit value, for counting shared leading bits.
_BITLEN = [0] * 0x10000
for _i in range(1, 0x10000):
    _BITLEN[_i] = _BITLEN[_i >> 1] + 1
del _i


# ── BIT STREAM ──────────────────────────────────────────────────────────────

class BitStream(object):
    """Write side of the U3D bit stream, arithmetic coder included.

    Bits fill each byte from the least significant end, which is what the
    reference library's little-endian U32 words amount to.
    """

    def __init__(self):
        self.out = bytearray()
        self._acc = 0       # partial byte, bits below _nbits are valid
        self._nbits = 0     # always 0..7
        self._high = 0xFFFF
        self._low = 0
        self._under = 0
        self._compressed = False
        self._hist = {}

    # -- raw bits -----------------------------------------------------------

    def _bit(self, b):
        self._acc |= b << self._nbits
        self._nbits += 1
        if self._nbits == 8:
            self.out.append(self._acc)
            self._acc = 0
            self._nbits = 0

    def _raw_byte(self, v):
        if self._nbits == 0:
            self.out.append(v)
        else:
            acc = self._acc | (v << self._nbits)
            self.out.append(acc & 0xFF)
            self._acc = acc >> 8

    # -- arithmetic coder ---------------------------------------------------

    def _encode(self, cum, freq, total):
        low = self._low
        rng = self._high + 1 - low
        high = low - 1 + rng * (cum + freq) // total
        low = low + rng * cum // total
        while (high & 0x8000) == (low & 0x8000):
            bit = low >> 15
            high = ((high & 0x7FFF) << 1) | 1
            self._bit(bit)
            while self._under > 0:
                self._under -= 1
                self._bit(bit ^ 1)
            low = (low & 0x7FFF) << 1
        while (high & 0x4000) == 0 and (low & 0x4000):
            high = ((high & 0x7FFF) << 1) | 0x8001
            low = (low << 1) & 0x7FFF
            self._under += 1
        self._high = high
        self._low = low

    def _hist_for(self, ctx):
        h = self._hist.get(ctx)
        if h is None:
            h = [1]         # symbol 0 is the escape, count 1
            self._hist[ctx] = h
        return h

    @staticmethod
    def _hist_add(h, sym):
        if sym > HIST_MAX_SYMBOL:
            return
        if sum(h) >= ELEPHANT:
            for i in range(len(h)):
                h[i] >>= 1
            h[0] += 1
        if sym >= len(h):
            h.extend([0] * (sym + 1 - len(h)))
        h[sym] += 1

    # -- plain values -------------------------------------------------------

    def u8(self, v):
        v &= 0xFF
        if self._high == 0xFFFF and self._low == 0 and self._under == 0:
            self._raw_byte(v)
        else:
            self._encode(_SWAP8[v], 1, 256)

    def u16(self, v):
        self.u8(v & 0xFF)
        self.u8((v >> 8) & 0xFF)

    def u32(self, v):
        v &= 0xFFFFFFFF
        if self._high == 0xFFFF and self._low == 0 and self._under == 0 \
                and self._nbits == 0:
            self.out += struct.pack("<I", v)
            return
        self.u16(v & 0xFFFF)
        self.u16((v >> 16) & 0xFFFF)

    def u64(self, v):
        self.u32(v & 0xFFFFFFFF)
        self.u32((v >> 32) & 0xFFFFFFFF)

    def f32(self, v):
        self.u32(struct.unpack("<I", struct.pack("<f", v))[0])

    def f64(self, v):
        self.u64(struct.unpack("<Q", struct.pack("<d", v))[0])

    def string(self, s):
        if not isinstance(s, bytearray):
            s = bytearray(s.encode("utf-8"))
        self.u16(len(s))
        for b in s:
            self.u8(b)

    def raw(self, data):
        """Append bytes as plain U8s (for nested blocks)."""
        if self._high == 0xFFFF and self._low == 0 and self._under == 0 \
                and self._nbits == 0:
            self.out += data
        else:
            for b in bytearray(data):
                self.u8(b)

    def align4(self):
        """Pad to a 4-byte boundary, relative to the start of this stream."""
        if self._nbits:
            self.out.append(self._acc)
            self._acc = 0
            self._nbits = 0
        while len(self.out) % 4:
            self.out.append(0)

    # -- compressed values --------------------------------------------------

    def cu32(self, ctx, v):
        self._compressed = True
        if not (ctx and ctx < MAX_RANGE):
            self.u32(v)
            return
        if ctx > STATIC_FULL:
            n = ctx - STATIC_FULL
            if v >= n:
                raise ValueError("value {} out of static range {}".format(v, n))
            self._encode(v, 1, n)
            return
        h = self._hist_for(ctx)
        sym = v + 1
        freq = h[sym] if sym < len(h) else 0
        if freq == 0:
            sym = 0
            freq = h[0]
        self._encode(sum(h[:sym]), freq, sum(h))
        self._hist_add(h, sym)
        if sym == 0:
            self.u32(v)
            self._hist_add(h, v + 1)

    def faces(self, faces, npos):
        """Base mesh faces: flat a, b, c, group per face, all coded inline.

        Same output as cu32 per value (shading id in the dynamic context,
        then three position indices), with the coder unrolled because this
        loop is most of the export time under IronPython.
        """
        self._compressed = True
        h = self._hist_for(CONTEXT_BASE_SHADING_ID)
        big = not (npos and STATIC_FULL + npos < MAX_RANGE)
        swap = _SWAP8
        bitlen = _BITLEN
        out = self.out
        acc = self._acc
        nb = self._nbits
        low = self._low
        high = self._high
        under = self._under
        for i in range(0, len(faces), 4):
            # Symbols for this face as (cum, freq, total); total 0 marks
            # a plain U8 whose coding depends on the coder state.
            group = faces[i + 3]
            sym = group + 1
            freq = h[sym] if sym < len(h) else 0
            if freq:
                syms = [(sum(h[:sym]), freq, sum(h))]
                self._hist_add(h, sym)
            else:
                syms = [(0, h[0], sum(h))]
                self._hist_add(h, 0)
                syms.extend([(group & 0xFF, 0, 0),
                             ((group >> 8) & 0xFF, 0, 0),
                             ((group >> 16) & 0xFF, 0, 0),
                             ((group >> 24) & 0xFF, 0, 0)])
                self._hist_add(h, group + 1)
            for v in (faces[i], faces[i + 1], faces[i + 2]):
                if big:
                    syms.extend([(v & 0xFF, 0, 0),
                                 ((v >> 8) & 0xFF, 0, 0),
                                 ((v >> 16) & 0xFF, 0, 0),
                                 ((v >> 24) & 0xFF, 0, 0)])
                else:
                    syms.append((v, 1, npos))

            for cum, freq, total in syms:
                if total == 0:
                    if high == 0xFFFF and low == 0 and under == 0:
                        if nb == 0:
                            out.append(cum)
                        else:
                            acc |= cum << nb
                            out.append(acc & 0xFF)
                            acc >>= 8
                        continue
                    cum = swap[cum]
                    freq = 1
                    total = 256
                rng = high + 1 - low
                high = low - 1 + rng * (cum + freq) // total
                low = low + rng * cum // total
                # Shift out every leading bit low and high share in one go,
                # the first followed by any pending underflow bits.
                k = 16 - bitlen[(low ^ high) & 0xFFFF]
                if k:
                    top = low >> (16 - k)
                    m = k
                    if under:
                        b = top >> (k - 1)
                        acc |= b << nb
                        nb += 1
                        if not b:
                            acc |= ((1 << under) - 1) << nb
                        nb += under
                        under = 0
                        m = k - 1
                        top &= (1 << m) - 1
                    if m:
                        acc |= (((swap[top & 0xFF] << 8) | swap[top >> 8])
                                >> (16 - m)) << nb
                        nb += m
                    while nb >= 8:
                        out.append(acc & 0xFF)
                        acc >>= 8
                        nb -= 8
                    high = ((high << k) & 0xFFFF) | ((1 << k) - 1)
                    low = (low << k) & 0xFFFF
                # Underflow: low 01..., high 10... for j places.
                z = ((low & ~high) << 1) & 0xFFFF
                j = 16 - bitlen[~z & 0xFFFF]
                if j:
                    high = ((high << j) & 0x7FFF) | 0x8000 | ((1 << j) - 1)
                    low = (low << j) & 0x7FFF
                    under += j
        self._acc = acc
        self._nbits = nb
        self._low = low
        self._high = high
        self._under = under

    def data(self):
        """Finish the stream and return its bytes, as GetDataBlock does."""
        if self._compressed:
            self.u32(0)
            self._compressed = False
        if self._nbits:
            self.out.append(self._acc)
            self._acc = 0
            self._nbits = 0
        return bytes(self.out)


# ── BLOCKS ──────────────────────────────────────────────────────────────────

def block(block_type, data, meta=b""):
    """Frame one block: type, sizes, data and metadata padded to 4 bytes."""
    out = bytearray(struct.pack("<III", block_type, len(data), len(meta)))
    out += data
    out += b"\0" * ((4 - len(data) % 4) % 4)
    out += meta
    out += b"\0" * ((4 - len(meta) % 4) % 4)
    return bytes(out)


def _matrix(w, m):
    for v in m:
        w.f32(v)


def _modifier_chain(name, chain_type, blocks):
    """Wrap blocks in a modifier chain. Type 0 is a node, 1 a resource."""
    w = BitStream()
    w.string(name)
    w.u32(chain_type)
    w.u32(0)            # no bounding sphere or box
    w.align4()
    w.u32(len(blocks))
    for b in blocks:
        w.raw(b)
    return block(BLOCK_MODIFIER_CHAIN, w.data())


def _node_parents(w, parents):
    w.u32(len(parents))
    for parent_name, m in parents:
        w.string(parent_name)
        _matrix(w, m)


# ── SCENE ───────────────────────────────────────────────────────────────────

class _Model(object):
    __slots__ = ("name", "parent", "positions", "faces", "shaders")

    def __init__(self, name, parent):
        self.name = name
        self.parent = parent
        self.positions = []     # flat x, y, z floats
        self.faces = []         # flat a, b, c, shading group per face
        self.shaders = []       # shader name per shading group


class U3DScene(object):
    """Collects groups, colours and meshes, then writes a .u3d file.

    Names are unique per kind: nodes share one namespace, resources and
    shaders their own. Positions are in the file's units (millimetres by
    default, see units_m).
    """

    def __init__(self, units_m=0.001):
        self.units_m = units_m
        self._groups = []           # (name, parent, matrix)
        self._node_names = set()
        self._shaders = {}          # key -> shader name
        self._shader_list = []      # (shader name, rgb, opacity)
        self._models = []

    def unique_node_name(self, name):
        base = name or "Unnamed"
        n = base
        i = 2
        while n in self._node_names:
            n = u"{} ({})".format(base, i)
            i += 1
        self._node_names.add(n)
        return n

    def add_group(self, name, parent=u"", matrix=IDENTITY):
        """Add a group node. matrix is 16 floats, column major."""
        name = self.unique_node_name(name)
        self._groups.append((name, parent, tuple(matrix)))
        return name

    def shader(self, rgb, opacity=1.0):
        """Shader name for a colour, rgb as 0..1 floats."""
        key = (round(rgb[0], 3), round(rgb[1], 3), round(rgb[2], 3),
               round(opacity, 3))
        name = self._shaders.get(key)
        if name is None:
            name = u"S{}".format(len(self._shader_list))
            self._shaders[key] = name
            self._shader_list.append((name, key[:3], key[3]))
        return name

    def add_model(self, name, parent=u""):
        """Start a model node; fill it with add_triangles."""
        m = _Model(self.unique_node_name(name), parent)
        self._models.append(m)
        return m

    def translate(self, dx, dy, dz):
        """Shift every position, e.g. to centre a model far from origin."""
        for m in self._models:
            p = m.positions
            for i in range(0, len(p), 3):
                p[i] += dx
                p[i + 1] += dy
                p[i + 2] += dz

    @staticmethod
    def add_triangles(model, shader, positions, triangles):
        """Add a patch: flat xyz positions and flat index triples."""
        try:
            group = model.shaders.index(shader)
        except ValueError:
            group = len(model.shaders)
            model.shaders.append(shader)
        base = len(model.positions) // 3
        model.positions.extend(positions)
        faces = model.faces
        for i in range(0, len(triangles), 3):
            faces.append(triangles[i] + base)
            faces.append(triangles[i + 1] + base)
            faces.append(triangles[i + 2] + base)
            faces.append(group)

    # -- encoding -----------------------------------------------------------

    def _group_block(self, name, parent, matrix):
        w = BitStream()
        w.string(name)
        _node_parents(w, [(parent, matrix)])
        return _modifier_chain(name, 0, [block(BLOCK_GROUP_NODE, w.data())])

    def _model_blocks(self, m, res_name):
        # Node chain: model node, then its shading modifier.
        w = BitStream()
        w.string(m.name)
        _node_parents(w, [(m.parent, IDENTITY)])
        w.string(res_name)
        w.u32(3)            # visible front and back
        node = block(BLOCK_MODEL_NODE, w.data())

        w = BitStream()
        w.string(m.name)
        w.u32(1)            # chain index, after the node
        w.u32(15)           # meshes, lines, points, glyphs
        w.u32(len(m.shaders))
        for s in m.shaders:
            w.u32(1)
            w.string(s)
        shading = block(BLOCK_SHADING_MODIFIER, w.data())
        node_chain = _modifier_chain(m.name, 0, [node, shading])

        npos = len(m.positions) // 3
        nface = len(m.faces) // 4
        ngroup = len(m.shaders)

        # Mesh declaration, wrapped in a resource chain.
        w = BitStream()
        w.string(res_name)
        w.u32(0)            # chain index
        w.u32(1)            # no normals, viewer derives them
        w.u32(nface)
        w.u32(npos)
        w.u32(0)            # normals
        w.u32(0)            # diffuse colours
        w.u32(0)            # specular colours
        w.u32(0)            # texture coords
        w.u32(ngroup)
        for i in range(ngroup):
            w.u32(0)        # no per-vertex colours
            w.u32(0)        # no texture layers
            w.u32(i)        # original shading id
        w.u32(npos)         # min resolution: all in the base mesh
        w.u32(npos)         # max resolution
        w.u32(1000)         # position quality
        w.u32(1000)         # normal quality
        w.u32(1000)         # texture quality
        w.f32(0.001)        # inverse quant: position
        w.f32(0.001)        # normal
        w.f32(0.001)        # texture coord
        w.f32(0.001)        # diffuse colour
        w.f32(0.001)        # specular colour
        w.f32(0.9)          # normal crease
        w.f32(0.5)          # normal update
        w.f32(0.985)        # normal tolerance
        w.u32(0)            # bones
        decl = block(BLOCK_MESH_DECLARATION, w.data())
        res_chain = _modifier_chain(res_name, 1, [decl])

        # Base mesh: every position and face, indices arithmetic coded.
        w = BitStream()
        w.string(res_name)
        w.u32(0)            # chain index
        w.u32(nface)
        w.u32(npos)
        w.u32(0)
        w.u32(0)
        w.u32(0)
        w.u32(0)
        if npos and w._nbits == 0:
            w.out += struct.pack("<{}f".format(len(m.positions)),
                                 *m.positions)
        else:
            for v in m.positions:
                w.f32(v)
        w.faces(m.faces, npos)
        cont = block(BLOCK_MESH_CONTINUATION, w.data())
        return node_chain, res_chain, cont

    def _shader_blocks(self, name, rgb, opacity):
        mat_name = u"M" + name[1:]
        w = BitStream()
        w.string(name)
        w.u32(1)            # lighting on
        w.f32(0.0)          # alpha test reference
        w.u32(0x617)        # alpha test: always
        w.u32(0x606)        # colour blend: alpha blend
        w.u32(1)            # render pass 0
        w.u32(0)            # no texture channels
        w.u32(0)            # no alpha texture channels
        w.string(mat_name)
        shader = block(BLOCK_LIT_TEXTURE_SHADER, w.data())

        r, g, b = rgb
        w = BitStream()
        w.string(mat_name)
        w.u32(0x3F)         # ambient, diffuse, specular, emissive, refl, opacity
        for v in (r * 0.25, g * 0.25, b * 0.25):    # ambient
            w.f32(v)
        for v in (r, g, b):                         # diffuse
            w.f32(v)
        for v in (0.15, 0.15, 0.15):                # specular
            w.f32(v)
        for v in (r * 0.1, g * 0.1, b * 0.1):       # emissive
            w.f32(v)
        w.f32(0.1)          # reflectivity
        w.f32(opacity)
        material = block(BLOCK_MATERIAL, w.data())
        return shader, material

    def to_bytes(self, progress=None):
        decl = []
        cont = []
        for name, parent, matrix in self._groups:
            decl.append(self._group_block(name, parent, matrix))
        for name, rgb, opacity in self._shader_list:
            decl.extend(self._shader_blocks(name, rgb, opacity))
        total = len(self._models)
        for i, m in enumerate(self._models):
            if not m.faces:
                continue
            node_chain, res_chain, c = self._model_blocks(m, u"R{}".format(i))
            decl.append(node_chain)
            decl.append(res_chain)
            cont.append(c)
            if progress:
                progress(i + 1, total)

        decl_bytes = b"".join(decl)
        cont_bytes = b"".join(cont)

        w = BitStream()
        w.u32(0)            # version 0.0, first edition
        w.u32(PROFILE_DEFINED_UNITS)
        w.u32(0)            # declaration size, patched below
        w.u64(0)            # file size, patched below
        w.u32(CHAR_ENCODING_UTF8)
        w.f64(self.units_m)
        header = bytearray(block(BLOCK_FILE_HEADER, w.data()))
        decl_size = len(header) + len(decl_bytes)
        file_size = decl_size + len(cont_bytes)
        struct.pack_into("<I", header, 20, decl_size)
        struct.pack_into("<Q", header, 24, file_size)
        return bytes(header) + decl_bytes + cont_bytes
