#!/usr/bin/env python3
"""
make_xlsm.py — Inject VBA modules into a real Excel-made skeleton .xlsm.

Strategy:
  We start from build/template_skeleton.xlsm — a file the user created in
  Excel for Mac by hand. It already contains a fully valid vbaProject.bin
  with the exact REFERENCES, _VBA_PROJECT performance-cache header,
  PROJECT/PROJECTwm structure and dir-stream record layout that Excel
  itself produces. We do NOT touch any of that.

  All we change is:
    - PROJECT stream:  rewrite the Document=/Module= lines and [Workspace]
                       to list our modules. Keep ID / CMG / DPB / GC /
                       [Host Extender Info] from the skeleton verbatim.
    - PROJECTwm stream: rebuild the name map for our module list.
    - VBA/dir stream:  keep bytes 0..PROJMODULES verbatim (= SYSKIND,
                       LCID, codepage 932, MSForms+Office references,
                       PROJECTVERSION blob), then write our own
                       PROJMODULES section.
    - VBA/_VBA_PROJECT: keep the skeleton's bytes verbatim. The cache it
                       contains is stale relative to our modules; Excel
                       discards stale cache entries by MCOOKIE mismatch
                       and recompiles from source.
    - VBA/Module1: dropped.
    - VBA/ThisWorkbook, VBA/Sheet1: replaced with our compressed source.
    - VBA/<NewModule>: added per module.

Creates:
  dist/Chatbot.xlsm
  dist/Admin_KnowledgeBuilder.xlsm

No external dependencies (stdlib only).
"""

import io
import os
import shutil
import struct
import zipfile

# ---------------------------------------------------------------------------
# Paths — script lives at <repo>/build/, source at <repo>/src/.
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
WORKTREE_ROOT = os.path.dirname(SCRIPT_DIR)

# Main repo is three levels up from worktree root (.claude/worktrees/<id>).
MAIN_REPO = os.path.normpath(os.path.join(WORKTREE_ROOT, '..', '..', '..'))
if not os.path.isdir(os.path.join(MAIN_REPO, 'src')):
    MAIN_REPO = WORKTREE_ROOT

SRC_DIR    = os.path.join(MAIN_REPO, "src")
BUILD_DIR  = os.path.join(MAIN_REPO, "build")
VENDOR_DIR = os.path.join(BUILD_DIR, "vendor")
TEMPLATE   = os.path.join(BUILD_DIR, "template_skeleton.xlsm")
DIST_DIR   = os.path.join(WORKTREE_ROOT, "dist")

# ---------------------------------------------------------------------------
# MS-OVBA compression (section 2.4) — symmetric with MS-OVBA decompress used
# inside Excel. Round-trips with oletools.olevba.decompress_stream.
# ---------------------------------------------------------------------------

def _ovba_compress_chunk(data: bytes) -> bytes:
    assert 1 <= len(data) <= 4096
    pos = 0
    out = bytearray()
    while pos < len(data):
        flag_pos = len(out)
        out.append(0)
        flag = 0
        for bit in range(8):
            if pos >= len(data):
                break
            # Bit widths depend on current decompressed position.
            if pos <= 16:    lbits, obits = 12, 4
            elif pos <= 32:  lbits, obits = 11, 5
            elif pos <= 64:  lbits, obits = 10, 6
            elif pos <= 128: lbits, obits = 9,  7
            elif pos <= 256: lbits, obits = 8,  8
            elif pos <= 512: lbits, obits = 7,  9
            elif pos <= 1024:lbits, obits = 6, 10
            elif pos <= 2048:lbits, obits = 5, 11
            else:            lbits, obits = 4, 12

            max_len = (1 << lbits) - 1 + 3
            best_len = 0
            best_off = 0
            window_start = max(0, pos - (1 << obits))
            if pos > 0 and (len(data) - pos) >= 3:
                for j in range(window_start, pos):
                    m = 0
                    while (m < max_len and pos + m < len(data)
                           and data[j + m] == data[pos + m]):
                        m += 1
                    if m > best_len:
                        best_len = m
                        best_off = j

            if best_len >= 3:
                flag |= (1 << bit)
                off_val = pos - best_off - 1
                len_val = best_len - 3
                out += struct.pack('<H', (off_val << lbits) | len_val)
                pos += best_len
            else:
                out.append(data[pos])
                pos += 1
        out[flag_pos] = flag
    return bytes(out)


def ovba_compress(data: bytes) -> bytes:
    """Compress a byte string into an OVBA CompressedContainer."""
    result = bytearray()
    result.append(0x01)  # SignatureByte
    offset = 0
    while offset < len(data):
        chunk = data[offset:offset + 4096]
        compressed = _ovba_compress_chunk(chunk)
        if len(chunk) == 4096 and len(compressed) >= 4096:
            # Use raw chunk: signature=0b011, flag=0, size field = 4095.
            result += struct.pack('<H', 0x3FFF)
            result += chunk
        else:
            header = 0xB000 | (len(compressed) - 1)
            result += struct.pack('<H', header)
            result += compressed
        offset += 4096
    return bytes(result)


# ---------------------------------------------------------------------------
# CFB (OLE Compound File Binary) writer — MS-CFB.
# Only what we need: one root, one sub-storage (VBA), small mini streams
# go through the Mini FAT.
# ---------------------------------------------------------------------------

FREESECT   = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT    = 0xFFFFFFFD
NOSTREAM   = 0xFFFFFFFF

SECTOR_SIZE = 512
MINI_CUTOFF = 4096
MINI_SECTOR = 64

# Storage CLSIDs Excel writes (little-endian on-disk form).
ROOT_CLSID = bytes([0x10, 0x42, 0x3B, 0x1C, 0x41, 0xF4, 0xCE, 0x11,
                    0xB9, 0xEA, 0x00, 0xAA, 0x00, 0x6B, 0x1A, 0x69])
VBA_CLSID  = bytes([0x70, 0xAE, 0x7B, 0xEA, 0x3B, 0xFB, 0xCD, 0x11,
                    0xA9, 0x03, 0x00, 0xAA, 0x00, 0x51, 0x0E, 0xA3])


def build_cfb(root_streams: dict, vba_streams: dict) -> bytes:
    """Build a CFB binary holding the given streams.

    root_streams: name -> bytes (children of Root Entry)
    vba_streams:  name -> bytes (children of VBA storage)
    """
    entries = []

    entries.append({'name': 'Root Entry', 'type': 5, 'clsid': ROOT_CLSID, 'data': None})
    entries.append({'name': 'VBA',        'type': 1, 'clsid': VBA_CLSID,  'data': None})

    root_slots = []
    for sn, sd in root_streams.items():
        root_slots.append(len(entries))
        entries.append({'name': sn, 'type': 2, 'clsid': b'\x00' * 16, 'data': sd})

    # Order VBA children: _VBA_PROJECT first, then dir, then modules. This
    # matches what tools like olevba expect, although the BST below sorts
    # them per spec regardless.
    vba_order = []
    if '_VBA_PROJECT' in vba_streams: vba_order.append('_VBA_PROJECT')
    if 'dir' in vba_streams:          vba_order.append('dir')
    for n in vba_streams:
        if n not in ('_VBA_PROJECT', 'dir'):
            vba_order.append(n)
    vba_slots = []
    for sn in vba_order:
        vba_slots.append(len(entries))
        entries.append({'name': sn, 'type': 2, 'clsid': b'\x00' * 16, 'data': vba_streams[sn]})

    # Pad entries to a multiple of 4 (one directory sector).
    while len(entries) % 4 != 0:
        entries.append({'name': '', 'type': 0, 'clsid': b'\x00' * 16, 'data': None})

    for e in entries:
        e.update(left=NOSTREAM, right=NOSTREAM, child=NOSTREAM,
                 start=ENDOFCHAIN, size=0)

    # Build a balanced BST for each parent's children, sorted by
    # (len(name), name.upper()) as MS-CFB requires.
    def key(idx):
        n = entries[idx]['name']
        return (len(n), n.upper())

    def bst(slots):
        if not slots:
            return NOSTREAM
        mid = len(slots) // 2
        r = slots[mid]
        entries[r]['left']  = bst(slots[:mid])
        entries[r]['right'] = bst(slots[mid + 1:])
        return r

    entries[0]['child'] = bst(sorted([1] + root_slots, key=key))
    entries[1]['child'] = bst(sorted(vba_slots, key=key))

    # Allocate stream data: small streams go in the Mini FAT.
    mini_streams = [e for e in entries
                    if e['type'] == 2 and e['data'] and len(e['data']) < MINI_CUTOFF]
    big_streams  = [e for e in entries
                    if e['type'] == 2 and e['data'] and len(e['data']) >= MINI_CUTOFF]

    # Mini stream container (the root entry's stream itself).
    mini_data = bytearray()
    minifat = []
    for e in mini_streams:
        d = e['data']
        start_mini = len(mini_data) // MINI_SECTOR
        padded = d + b'\x00' * ((-len(d)) % MINI_SECTOR)
        n = len(padded) // MINI_SECTOR
        for i in range(n):
            mini_data += padded[i * MINI_SECTOR:(i + 1) * MINI_SECTOR]
            minifat.append(start_mini + i + 1 if i < n - 1 else ENDOFCHAIN)
        e['start'] = start_mini
        e['size']  = len(d)
    mini_padded = bytes(mini_data) + b'\x00' * ((-len(mini_data)) % SECTOR_SIZE)

    # Regular sectors: big streams, then mini-container, then directory,
    # then mini-FAT, then FAT.
    sectors = bytearray()
    fat = []

    def append_chain(blob: bytes):
        start = len(sectors) // SECTOR_SIZE
        n = len(blob) // SECTOR_SIZE
        for i in range(n):
            sectors.extend(blob[i * SECTOR_SIZE:(i + 1) * SECTOR_SIZE])
            fat.append(start + i + 1 if i < n - 1 else ENDOFCHAIN)
        return start, n

    for e in big_streams:
        padded = e['data'] + b'\x00' * ((-len(e['data'])) % SECTOR_SIZE)
        start, _ = append_chain(padded)
        e['start'] = start
        e['size']  = len(e['data'])

    if mini_padded:
        start, _ = append_chain(mini_padded)
        entries[0]['start'] = start
        entries[0]['size']  = len(mini_data)
    else:
        entries[0]['start'] = ENDOFCHAIN
        entries[0]['size']  = 0

    # Directory sectors.
    dir_bytes = bytearray()
    for e in entries:
        name_utf16 = e['name'].encode('utf-16-le') if e['name'] else b''
        name_len = len(name_utf16) + 2 if name_utf16 else 0
        name_field = (name_utf16 + b'\x00' * (64 - len(name_utf16)))[:64]
        buf = bytearray(128)
        buf[0:64] = name_field
        struct.pack_into('<H', buf, 64, name_len)
        buf[66] = e['type']
        buf[67] = 0x01  # ColorFlag: black
        struct.pack_into('<I', buf, 68, e['left'])
        struct.pack_into('<I', buf, 72, e['right'])
        struct.pack_into('<I', buf, 76, e['child'])
        buf[80:96] = e['clsid']
        struct.pack_into('<I', buf,  96, 0)
        struct.pack_into('<Q', buf, 100, 0)
        struct.pack_into('<Q', buf, 108, 0)
        struct.pack_into('<I', buf, 116, e['start'])
        struct.pack_into('<I', buf, 120, e['size'])
        struct.pack_into('<I', buf, 124, 0)
        dir_bytes += buf
    dir_start, _ = append_chain(bytes(dir_bytes))

    # Mini-FAT sectors.
    ENTRIES_PER_FAT = SECTOR_SIZE // 4
    if minifat:
        mf = list(minifat)
        while len(mf) % ENTRIES_PER_FAT != 0:
            mf.append(FREESECT)
        minifat_start, num_minifat = append_chain(
            b''.join(struct.pack('<I', x) for x in mf))
    else:
        minifat_start, num_minifat = ENDOFCHAIN, 0

    # FAT sectors (self-describing — iterate until the count stabilizes).
    data_sectors = len(sectors) // SECTOR_SIZE
    num_fat = 1
    while True:
        total = data_sectors + num_fat
        need = (total + ENTRIES_PER_FAT - 1) // ENTRIES_PER_FAT
        if need <= num_fat:
            break
        num_fat = need
    fat_start = data_sectors
    fat_full = list(fat) + [FATSECT] * num_fat
    while len(fat_full) % ENTRIES_PER_FAT != 0:
        fat_full.append(FREESECT)
    fat_blob = b''.join(struct.pack('<I', x) for x in fat_full)

    # CFB Header.
    header = bytearray(512)
    header[0:8] = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'
    struct.pack_into('<H', header, 24, 0x003E)   # MinorVersion
    struct.pack_into('<H', header, 26, 0x0003)   # MajorVersion
    struct.pack_into('<H', header, 28, 0xFFFE)   # ByteOrder
    struct.pack_into('<H', header, 30, 0x0009)   # SectorSizePower (=512)
    struct.pack_into('<H', header, 32, 0x0006)   # MiniSectorSizePower (=64)
    struct.pack_into('<I', header, 40, 0)
    struct.pack_into('<I', header, 44, num_fat)
    struct.pack_into('<I', header, 48, dir_start)
    struct.pack_into('<I', header, 52, 0)
    struct.pack_into('<I', header, 56, MINI_CUTOFF)
    struct.pack_into('<I', header, 60, minifat_start)
    struct.pack_into('<I', header, 64, num_minifat)
    struct.pack_into('<I', header, 68, ENDOFCHAIN)
    struct.pack_into('<I', header, 72, 0)
    for i in range(109):
        val = fat_start + i if i < num_fat else FREESECT
        struct.pack_into('<I', header, 76 + i * 4, val)

    return bytes(header) + bytes(sectors) + fat_blob


# ---------------------------------------------------------------------------
# OVBA decompression (for parsing the skeleton's dir stream — we only need
# read access).
# ---------------------------------------------------------------------------

def ovba_decompress(data: bytes) -> bytes:
    """Decompress an OVBA CompressedContainer (MS-OVBA 2.4)."""
    if not data or data[0] != 0x01:
        raise ValueError("Bad OVBA signature byte")
    out = bytearray()
    i = 1
    while i < len(data):
        header = struct.unpack('<H', data[i:i + 2])[0]
        i += 2
        chunk_size = (header & 0x0FFF) + 3
        chunk_signature = (header >> 12) & 0x07
        chunk_flag = (header >> 15) & 0x01
        body = data[i:i + chunk_size - 2]
        i += chunk_size - 2
        if chunk_flag == 0:
            # Raw, exactly 4096 bytes
            out += body
            continue
        # Compressed
        chunk_start = len(out)
        j = 0
        while j < len(body):
            flag = body[j]; j += 1
            for bit in range(8):
                if j >= len(body):
                    break
                if not (flag & (1 << bit)):
                    out.append(body[j]); j += 1
                else:
                    pos = len(out) - chunk_start
                    if pos <= 16:    lbits = 12
                    elif pos <= 32:  lbits = 11
                    elif pos <= 64:  lbits = 10
                    elif pos <= 128: lbits = 9
                    elif pos <= 256: lbits = 8
                    elif pos <= 512: lbits = 7
                    elif pos <= 1024:lbits = 6
                    elif pos <= 2048:lbits = 5
                    else:            lbits = 4
                    token = struct.unpack('<H', body[j:j + 2])[0]; j += 2
                    length = (token & ((1 << lbits) - 1)) + 3
                    offset = (token >> lbits) + 1
                    src = len(out) - offset
                    for k in range(length):
                        out.append(out[src + k])
    return bytes(out)


# ---------------------------------------------------------------------------
# Skeleton CFB reader — minimal, just enough to pull the streams we need.
# ---------------------------------------------------------------------------

class CFBReader:
    def __init__(self, data: bytes):
        self.data = data
        assert data[:8] == b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'
        self.minifat_start = struct.unpack('<I', data[60:64])[0]
        self.minifat_count = struct.unpack('<I', data[64:68])[0]
        self.fat_first     = struct.unpack('<I', data[76:80])[0]
        self.dir_start     = struct.unpack('<I', data[48:52])[0]

        # FAT (assume single sector for skeleton-sized files).
        fat_off = (self.fat_first + 1) * SECTOR_SIZE
        self.fat = list(struct.unpack('<128I', data[fat_off:fat_off + SECTOR_SIZE]))

        # Mini FAT.
        self.minifat = []
        sec = self.minifat_start
        while sec != ENDOFCHAIN and sec < len(self.fat):
            off = (sec + 1) * SECTOR_SIZE
            self.minifat += list(struct.unpack('<128I', data[off:off + SECTOR_SIZE]))
            sec = self.fat[sec]

        # Directory entries.
        self.entries = {}  # name -> dict
        sec = self.dir_start
        chain = []
        while sec != ENDOFCHAIN and sec < len(self.fat):
            chain.append(sec)
            sec = self.fat[sec]
        dir_bytes = b''.join(data[(s + 1) * SECTOR_SIZE:(s + 2) * SECTOR_SIZE]
                             for s in chain)
        for i in range(0, len(dir_bytes), 128):
            e = dir_bytes[i:i + 128]
            name_len = struct.unpack('<H', e[64:66])[0]
            if name_len == 0: continue
            name = e[:name_len - 2].decode('utf-16-le')
            etype = e[66]
            start = struct.unpack('<I', e[116:120])[0]
            size  = struct.unpack('<I', e[120:124])[0]
            self.entries[name] = {'type': etype, 'start': start, 'size': size}

        # Root entry stores the mini-stream container.
        root = self.entries.get('Root Entry')
        if root and root['size']:
            sec = root['start']
            blob = bytearray()
            while sec != ENDOFCHAIN and sec < len(self.fat):
                off = (sec + 1) * SECTOR_SIZE
                blob += data[off:off + SECTOR_SIZE]
                sec = self.fat[sec]
            self.mini_data = bytes(blob[:root['size']])
        else:
            self.mini_data = b''

    def read(self, name: str) -> bytes:
        e = self.entries[name]
        if e['size'] < MINI_CUTOFF:
            sec = e['start']
            blob = bytearray()
            while sec != ENDOFCHAIN and sec < len(self.minifat):
                off = sec * MINI_SECTOR
                blob += self.mini_data[off:off + MINI_SECTOR]
                sec = self.minifat[sec]
            return bytes(blob[:e['size']])
        else:
            sec = e['start']
            blob = bytearray()
            while sec != ENDOFCHAIN and sec < len(self.fat):
                off = (sec + 1) * SECTOR_SIZE
                blob += self.data[off:off + SECTOR_SIZE]
                sec = self.fat[sec]
            return bytes(blob[:e['size']])


# ---------------------------------------------------------------------------
# dir-stream MODULES section builder.
# Layout (matches what real Excel writes — see skeleton verification):
#   PROJMODULES (0x000F, sz=2): module count
#   PROJCOOKIE  (0x0013, sz=2): 0xFFFF
#   For each module:
#     MNAME      (0x0019) ANSI module name
#     MNAME_U    (0x0047) UTF-16LE module name
#     MSTREAM    (0x001A) ANSI stream name (same as module name)
#     MSTREAM_U  (0x0032) UTF-16LE stream name
#     MDOCSTR    (0x001C) empty
#     MDOCSTR_U  (0x0048) empty
#     MOFFSET    (0x0031) uint32 = 0 (source starts at byte 0 of the stream)
#     MHELPCTX   (0x001E) uint32 = 0
#     MCOOKIE    (0x002C) uint16 = 0xFFFF
#     MTYPE      (0x0021 standard | 0x0022 document) empty
#     MTERM      (0x002B) empty
#   PROJTERM (0x0010): record terminator
# ---------------------------------------------------------------------------

def _rec(rid: int, data: bytes) -> bytes:
    return struct.pack('<HI', rid, len(data)) + data


def _ansi(s: str) -> bytes:
    # Project codepage is 932 (Shift-JIS) — set in the dir prefix we copy
    # from the skeleton, so encode names the same way.
    return s.encode('cp932')


def _u16(s: str) -> bytes:
    return s.encode('utf-16-le')


def build_modules_section(modules: list) -> bytes:
    """Modules + per-module records + project terminator."""
    buf = bytearray()
    buf += _rec(0x000F, struct.pack('<H', len(modules)))   # PROJMODULES
    buf += _rec(0x0013, struct.pack('<H', 0xFFFF))         # PROJCOOKIE

    for mod in modules:
        name = mod['name']
        is_class = mod.get('class', False)
        buf += _rec(0x0019, _ansi(name))                   # MNAME
        buf += _rec(0x0047, _u16(name))                    # MNAME_U
        buf += _rec(0x001A, _ansi(name))                   # MSTREAM
        buf += _rec(0x0032, _u16(name))                    # MSTREAM_U
        buf += _rec(0x001C, b'')                           # MDOCSTR
        buf += _rec(0x0048, b'')                           # MDOCSTR_U
        buf += _rec(0x0031, struct.pack('<I', 0))          # MOFFSET=0
        buf += _rec(0x001E, struct.pack('<I', 0))          # MHELPCTX
        buf += _rec(0x002C, struct.pack('<H', 0xFFFF))     # MCOOKIE
        buf += _rec(0x0022 if is_class else 0x0021, b'')   # MTYPE
        buf += _rec(0x002B, b'')                           # MTERM

    buf += _rec(0x0010, b'')                                # PROJTERM
    return bytes(buf)


def find_modules_section_offset(dir_uncompressed: bytes) -> int:
    """Byte offset of PROJMODULES record header in the dir stream."""
    marker = b'\x0f\x00\x02\x00\x00\x00'
    idx = dir_uncompressed.find(marker)
    if idx < 0:
        raise RuntimeError("PROJMODULES marker not found in skeleton dir stream")
    return idx


# ---------------------------------------------------------------------------
# PROJECT stream builder — preserve as much as possible of the skeleton's
# PROJECT, only swapping out the document/module list and [Workspace].
# ---------------------------------------------------------------------------

def parse_skeleton_project(text: str) -> dict:
    """Parse skeleton PROJECT stream into a dict of preserved fields."""
    info = {}
    for line in text.splitlines():
        if line.startswith('ID='):                  info['ID'] = line
        elif line.startswith('Name='):              info['Name'] = line
        elif line.startswith('HelpContextID='):     info['HelpContextID'] = line
        elif line.startswith('VersionCompatible'):  info['VersionCompatible'] = line
        elif line.startswith('CMG='):               info['CMG'] = line
        elif line.startswith('DPB='):               info['DPB'] = line
        elif line.startswith('GC='):                info['GC'] = line
        elif line.startswith('&H00000001=') and 'CF90' in line:
            info['HostExtenderVBE'] = line
    return info


def build_project_stream(skel_info: dict, modules: list) -> bytes:
    """Assemble a PROJECT stream for our module set."""
    lines = [skel_info['ID']]

    for mod in modules:
        if mod.get('class'):
            lines.append(f'Document={mod["name"]}/&H00000000')
        else:
            lines.append(f'Module={mod["name"]}')

    lines.append(skel_info.get('Name', 'Name="VBAProject"'))
    lines.append(skel_info.get('HelpContextID', 'HelpContextID="0"'))
    lines.append(skel_info.get('VersionCompatible', 'VersionCompatible32="393222000"'))
    lines.append(skel_info.get('CMG', 'CMG=""'))
    lines.append(skel_info.get('DPB', 'DPB=""'))
    lines.append(skel_info.get('GC',  'GC=""'))
    lines.append('')
    lines.append('[Host Extender Info]')
    lines.append(skel_info.get('HostExtenderVBE',
                 '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000'))
    lines.append('')
    lines.append('[Workspace]')
    for mod in modules:
        # 0, 0, 0, 0, C marks "code window closed".
        lines.append(f'{mod["name"]}=0, 0, 0, 0, C')
    lines.append('')

    return '\r\n'.join(lines).encode('cp932')


def build_projectwm(modules: list) -> bytes:
    """Module name map: ANSI name + 0 + UTF-16 name + 00 00 (repeat), 00 00 end."""
    buf = bytearray()
    for mod in modules:
        n = mod['name']
        buf += n.encode('cp932') + b'\x00'
        buf += n.encode('utf-16-le') + b'\x00\x00'
    buf += b'\x00\x00'
    return bytes(buf)


# ---------------------------------------------------------------------------
# Source file loading.
# ---------------------------------------------------------------------------

def strip_bas_header(content: str) -> str:
    out = []
    found = False
    for line in content.splitlines(keepends=True):
        stripped = line.lstrip('﻿')
        if not found and stripped.lstrip().startswith('Attribute VB_Name'):
            found = True
            continue
        out.append(stripped if not out else line)
    return ''.join(out)


def strip_cls_header(content: str) -> str:
    out = []
    past = False
    for line in content.splitlines(keepends=True):
        if not past:
            if line.lstrip('﻿').strip().startswith('Attribute VB_Exposed'):
                past = True
            continue
        out.append(line)
    return ''.join(out)


def load_module_source(path: str, is_cls: bool = False) -> bytes:
    with open(path, 'r', encoding='utf-8-sig', errors='replace') as f:
        content = f.read()
    src = strip_cls_header(content) if is_cls else strip_bas_header(content)
    src = src.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    return src.encode('cp932')


EMPTY_DOC_SOURCE = b''  # Excel is happy with truly empty document streams.


# ---------------------------------------------------------------------------
# Module set definitions.
# ---------------------------------------------------------------------------

def _doc(name: str, source: bytes) -> dict:
    return {'name': name, 'source': source, 'class': True}


def _std(name: str, source: bytes) -> dict:
    return {'name': name, 'source': source, 'class': False}


def load_modules_chatbot() -> list:
    tw_src = load_module_source(
        os.path.join(SRC_DIR, 'chatbot', 'ThisWorkbook.cls'), is_cls=True)

    modules = [
        _doc('ThisWorkbook', tw_src),
        _doc('Sheet1', EMPTY_DOC_SOURCE),
    ]

    for n in ['modTypes', 'modConfig', 'modPaths', 'modHttpClient',
              'modKeyVault', 'modApiGateway']:
        modules.append(_std(n, load_module_source(
            os.path.join(SRC_DIR, 'shared', n + '.bas'))))

    modules.append(_std('JsonConverter', load_module_source(
        os.path.join(VENDOR_DIR, 'JsonConverter.bas'))))

    for n in ['modIndexReader', 'modSimilarity', 'modRagEngine',
              'modUserProfile', 'modPiiGuard', 'modRateLimiter',
              'modUsageLogger', 'modChatUI', 'modBoot']:
        modules.append(_std(n, load_module_source(
            os.path.join(SRC_DIR, 'chatbot', n + '.bas'))))

    return modules


def load_modules_admin() -> list:
    modules = [
        _doc('ThisWorkbook', EMPTY_DOC_SOURCE),
        _doc('Sheet1', EMPTY_DOC_SOURCE),
    ]

    for n in ['modTypes', 'modConfig', 'modPaths', 'modHttpClient',
              'modKeyVault', 'modApiGateway']:
        modules.append(_std(n, load_module_source(
            os.path.join(SRC_DIR, 'shared', n + '.bas'))))

    modules.append(_std('JsonConverter', load_module_source(
        os.path.join(VENDOR_DIR, 'JsonConverter.bas'))))

    for n in ['modChunker', 'modExtractor', 'modExtractorWord',
              'modExtractorAcrobat', 'modExtractorExcel',
              'modIndexWriter', 'modKnowledgeBuilder',
              'modKeyEnroller', 'modUsageAggregator']:
        modules.append(_std(n, load_module_source(
            os.path.join(SRC_DIR, 'admin', n + '.bas'))))

    return modules


# ---------------------------------------------------------------------------
# OOXML patching: bind workbook/sheet to their VBA code modules via codeName.
# ---------------------------------------------------------------------------

def patch_workbook_codename(xml_bytes: bytes, code_name: str) -> bytes:
    xml = xml_bytes.decode('utf-8')
    if 'codeName' in xml:
        return xml.encode('utf-8')
    if '<workbookPr' in xml:
        xml = xml.replace('<workbookPr', f'<workbookPr codeName="{code_name}"', 1)
    else:
        end = xml.index('>', xml.index('<workbook')) + 1
        xml = xml[:end] + f'<workbookPr codeName="{code_name}"/>' + xml[end:]
    return xml.encode('utf-8')


def patch_sheet_codename(xml_bytes: bytes, code_name: str) -> bytes:
    xml = xml_bytes.decode('utf-8')
    if 'codeName' in xml:
        return xml.encode('utf-8')
    if '<sheetPr' in xml:
        xml = xml.replace('<sheetPr', f'<sheetPr codeName="{code_name}"', 1)
    else:
        end = xml.index('>', xml.index('<worksheet')) + 1
        xml = xml[:end] + f'<sheetPr codeName="{code_name}"/>' + xml[end:]
    return xml.encode('utf-8')


# ---------------------------------------------------------------------------
# Top-level builder: take skeleton, swap in our modules, produce .xlsm.
# ---------------------------------------------------------------------------

def _patch_projsyskind(dir_bytes: bytes, syskind: int) -> bytes:
    """Patch PROJSYSKIND in a decompressed VBA dir stream.

    The skeleton was created on a 64-bit machine (SYSKIND=3). On 32-bit
    corporate Excel (common even on 64-bit Windows) a Win64 SYSKIND causes
    the entire VBA project to be silently rejected.  We force Win32 (1)
    which is accepted by both 32-bit and 64-bit Excel. Since all module
    streams have MOFFSET=0, Excel always recompiles from source regardless
    of SYSKIND.
    """
    marker = struct.pack('<HI', 0x0001, 4)   # RecordID=PROJSYSKIND, Size=4
    idx = dir_bytes.find(marker)
    if idx < 0:
        return dir_bytes
    ba = bytearray(dir_bytes)
    struct.pack_into('<I', ba, idx + 6, syskind)
    return bytes(ba)


def build_vba_project(modules: list, skeleton_vba_bin: bytes) -> bytes:
    """Build a new vbaProject.bin by surgically modifying the skeleton."""
    skel = CFBReader(skeleton_vba_bin)

    # Skeleton streams we keep verbatim.
    skel_dir_compressed = skel.read('dir')
    skel_dir = ovba_decompress(skel_dir_compressed)

    # Force PROJSYSKIND = Win32 (1) so both 32-bit and 64-bit Excel accept
    # the project.  The skeleton was compiled on a 64-bit machine (SYSKIND=3)
    # and 32-bit Excel silently drops the entire VBA project when it sees 3.
    skel_dir = _patch_projsyskind(skel_dir, 0x00000001)
    vba_project_blob   = skel.read('_VBA_PROJECT')
    skel_project_text  = skel.read('PROJECT').decode('cp932', errors='replace')
    skel_info          = parse_skeleton_project(skel_project_text)

    # Build new dir: prefix (verbatim) + our MODULES section.
    modules_off = find_modules_section_offset(skel_dir)
    new_dir_uncompressed = skel_dir[:modules_off] + build_modules_section(modules)
    new_dir_compressed = ovba_compress(new_dir_uncompressed)

    # New PROJECT and PROJECTwm.
    new_project   = build_project_stream(skel_info, modules)
    new_projectwm = build_projectwm(modules)

    # Module streams: every module's content is just its OVBA-compressed
    # source — MOFFSET=0 means "no cached p-code, source starts at byte 0".
    root_streams = {'PROJECT': new_project, 'PROJECTwm': new_projectwm}
    vba_streams  = {'_VBA_PROJECT': vba_project_blob, 'dir': new_dir_compressed}
    for mod in modules:
        body = mod['source'] if mod['source'] else b'\r\n'
        vba_streams[mod['name']] = ovba_compress(body)

    return build_cfb(root_streams, vba_streams)


def patch_content_types(xml_bytes: bytes) -> bytes:
    xml = xml_bytes.decode('utf-8')
    # Swap the workbook.xml content type from xlsx (no macros) to xlsm (macro-enabled).
    # Without this, Excel rejects the file with "ファイル形式またはファイル拡張子が正しくありません".
    xml = xml.replace(
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
        'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
    )
    if 'vbaProject' not in xml:
        xml = xml.replace('</Types>',
            '<Override PartName="/xl/vbaProject.bin" '
            'ContentType="application/vnd.ms-office.vbaProject"/></Types>')
    return xml.encode('utf-8')


def patch_workbook_rels(xml_bytes: bytes) -> bytes:
    xml = xml_bytes.decode('utf-8')
    if 'vbaProject' not in xml:
        xml = xml.replace('</Relationships>',
            '<Relationship Id="rId99" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vbaProject" '
            'Target="vbaProject.bin"/></Relationships>')
    return xml.encode('utf-8')


def create_xlsm(template_path: str, output_path: str, modules: list):
    print(f"  Building vbaProject.bin from skeleton ({len(modules)} modules)...")
    with open(template_path, 'rb') as f:
        template_bytes = f.read()
    with zipfile.ZipFile(io.BytesIO(template_bytes)) as z:
        skel_vba = z.read('xl/vbaProject.bin')

    vba_bin = build_vba_project(modules, skel_vba)
    print(f"  vbaProject.bin size: {len(vba_bin):,} bytes")

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    with zipfile.ZipFile(io.BytesIO(template_bytes)) as zin, \
         zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'xl/vbaProject.bin':
                data = vba_bin
            elif item.filename == '[Content_Types].xml':
                data = patch_content_types(data)
            elif item.filename == 'xl/_rels/workbook.xml.rels':
                data = patch_workbook_rels(data)
            elif item.filename == 'xl/workbook.xml':
                data = patch_workbook_codename(data, 'ThisWorkbook')
            elif item.filename == 'xl/worksheets/sheet1.xml':
                data = patch_sheet_codename(data, 'Sheet1')
            zout.writestr(item, data)

    print(f"  Created: {output_path} ({os.path.getsize(output_path):,} bytes)")


# ---------------------------------------------------------------------------
# Main.
# ---------------------------------------------------------------------------

def main():
    print("=== make_xlsm.py ===")
    print(f"Main repo:   {MAIN_REPO}")
    print(f"Template:    {TEMPLATE}")
    print(f"Source dir:  {SRC_DIR}")
    print(f"Output dir:  {DIST_DIR}")
    print()

    if not os.path.exists(TEMPLATE):
        raise FileNotFoundError(f"Template not found: {TEMPLATE}")
    if not os.path.isdir(SRC_DIR):
        raise FileNotFoundError(f"src/ directory not found: {SRC_DIR}")

    os.makedirs(DIST_DIR, exist_ok=True)

    print("Building Chatbot.xlsm...")
    chatbot = load_modules_chatbot()
    print(f"  Modules ({len(chatbot)}): {[m['name'] for m in chatbot]}")
    create_xlsm(TEMPLATE, os.path.join(DIST_DIR, 'Chatbot.xlsm'), chatbot)

    print()
    print("Building Admin_KnowledgeBuilder.xlsm...")
    admin = load_modules_admin()
    print(f"  Modules ({len(admin)}): {[m['name'] for m in admin]}")
    create_xlsm(TEMPLATE, os.path.join(DIST_DIR, 'Admin_KnowledgeBuilder.xlsm'), admin)

    print()
    print("Done.")


if __name__ == '__main__':
    main()
