#!/usr/bin/env python3
"""
make_xlsm.py — Pure-Python script to inject VBA modules into a blank .xlsm template.

Creates:
  dist/Chatbot.xlsm
  dist/Admin_KnowledgeBuilder.xlsm

No external dependencies beyond stdlib.
"""

import io
import os
import shutil
import struct
import zipfile

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# The worktree build/ dir mirrors the main repo's build/; source files live
# in the main notebook repo which is the parent of .claude/worktrees/<id>
# Resolve: worktree root -> .claude/worktrees/<id>  <- SCRIPT_DIR/../
WORKTREE_ROOT = os.path.dirname(SCRIPT_DIR)   # e.g. .../agent-a55b728f520139c02

# Main repo is three levels up from the worktree root:
#   notebook/.claude/worktrees/agent-xxx  -> notebook
MAIN_REPO = os.path.normpath(os.path.join(WORKTREE_ROOT, '..', '..', '..'))
# Fallback: if running from the real build/ directory
if not os.path.isdir(os.path.join(MAIN_REPO, 'src')):
    MAIN_REPO = WORKTREE_ROOT

SRC_DIR    = os.path.join(MAIN_REPO, "src")
BUILD_DIR  = os.path.join(MAIN_REPO, "build")
VENDOR_DIR = os.path.join(BUILD_DIR, "vendor")
TEMPLATE   = os.path.join(BUILD_DIR, "template_blank.xlsm")
DIST_DIR   = os.path.join(WORKTREE_ROOT, "dist")

# ---------------------------------------------------------------------------
# MS-OVBA Compression (Section 2.4)
# ---------------------------------------------------------------------------

def _ovba_compress_chunk(data: bytes) -> bytes:
    """Compress up to 4096 bytes using OVBA LZ77."""
    assert 1 <= len(data) <= 4096

    decompressed_size = len(data)
    pos = 0  # current position in decompressed data
    output = bytearray()

    while pos < decompressed_size:
        flag_byte_pos = len(output)
        output.append(0)  # placeholder for flag byte
        flag_byte = 0

        for bit_index in range(8):
            if pos >= decompressed_size:
                break

            # Determine bit widths based on current decompressed position
            copy_token_pos = pos
            if copy_token_pos <= 16:
                length_mask_bits = 12
                offset_mask_bits = 4
            elif copy_token_pos <= 32:
                length_mask_bits = 11
                offset_mask_bits = 5
            elif copy_token_pos <= 64:
                length_mask_bits = 10
                offset_mask_bits = 6
            elif copy_token_pos <= 128:
                length_mask_bits = 9
                offset_mask_bits = 7
            elif copy_token_pos <= 256:
                length_mask_bits = 8
                offset_mask_bits = 8
            elif copy_token_pos <= 512:
                length_mask_bits = 7
                offset_mask_bits = 9
            elif copy_token_pos <= 1024:
                length_mask_bits = 6
                offset_mask_bits = 10
            elif copy_token_pos <= 2048:
                length_mask_bits = 5
                offset_mask_bits = 11
            else:
                length_mask_bits = 4
                offset_mask_bits = 12

            max_length = (1 << length_mask_bits) - 1 + 3  # +3 because length stored as len-3

            # Search for best match in window (up to pos bytes back)
            best_len = 0
            best_offset = 0
            window_start = max(0, pos - (1 << offset_mask_bits))

            if pos > 0 and (decompressed_size - pos) >= 3:
                for j in range(window_start, pos):
                    match_len = 0
                    while (match_len < max_length and
                           pos + match_len < decompressed_size and
                           data[j + match_len] == data[pos + match_len]):
                        match_len += 1
                    if match_len > best_len:
                        best_len = match_len
                        best_offset = j

            if best_len >= 3:
                # Emit copy token
                flag_byte |= (1 << bit_index)
                offset_val = pos - best_offset - 1
                length_val = best_len - 3
                token = (offset_val << length_mask_bits) | length_val
                output += struct.pack('<H', token)
                pos += best_len
            else:
                # Emit literal byte
                output.append(data[pos])
                pos += 1

        output[flag_byte_pos] = flag_byte

    return bytes(output)


def ovba_compress(data: bytes) -> bytes:
    """Compress data using OVBA compression (MS-OVBA 2.4.1)."""
    result = bytearray()
    result.append(0x01)  # SignatureByte

    offset = 0
    while offset < len(data):
        chunk_data = data[offset:offset + 4096]
        chunk_size = len(chunk_data)

        if chunk_size == 4096:
            # Try compressing
            compressed = _ovba_compress_chunk(chunk_data)
            if len(compressed) < 4096:
                # Use compressed
                header = 0xB000 | (len(compressed) - 3)
                result += struct.pack('<H', header)
                result += compressed
            else:
                # Use raw (2-byte header 0x3B11 + 4096 bytes data)
                result += struct.pack('<H', 0x3B11)
                result += chunk_data
        else:
            # Partial chunk: MUST use compressed format
            compressed = _ovba_compress_chunk(chunk_data)
            header = 0xB000 | (len(compressed) - 3)
            result += struct.pack('<H', header)
            result += compressed

        offset += 4096

    return bytes(result)


# ---------------------------------------------------------------------------
# dir stream builder (MS-OVBA 2.3.4)
# ---------------------------------------------------------------------------

def _record(rid: int, data: bytes) -> bytes:
    return struct.pack('<HI', rid, len(data)) + data


def _u16(s: str) -> bytes:
    return s.encode('utf-16-le')


def _ansi(s: str) -> bytes:
    return s.encode('cp1252')


def build_dir_stream(modules: list) -> bytes:
    """Build the uncompressed dir stream."""
    buf = bytearray()

    # PROJECTSYSKIND
    buf += _record(0x0001, struct.pack('<I', 0x00000001))
    # PROJECTLCID
    buf += _record(0x0002, struct.pack('<I', 0x00000409))
    # PROJECTLCIDINVOKE
    buf += _record(0x0014, struct.pack('<I', 0x00000409))
    # PROJECTCODEPAGE
    buf += _record(0x0003, struct.pack('<H', 0x04E4))
    # PROJECTNAME
    buf += _record(0x0004, _ansi("VBAProject"))
    # PROJECTDOCSTRING
    buf += _record(0x0005, b'')
    # PROJECTDOCSTRINGUNICODE
    buf += _record(0x0040, b'')
    # PROJECTHELPFILEPATH (first part)
    buf += _record(0x0006, b'')
    # PROJECTHELPFILEPATH (second part)
    buf += _record(0x003D, b'')
    # PROJECTHELPCONTEXT
    buf += _record(0x0007, struct.pack('<I', 0x00000000))
    # PROJECTLIBFLAGS
    buf += _record(0x0008, struct.pack('<I', 0x00000000))

    # PROJECTVERSION: special - RecordID(2) + Size(4)=4 + MajorVersion(4) + MinorVersion(2)
    buf += struct.pack('<HI', 0x0009, 0x00000004)
    buf += struct.pack('<IH', 0x00000061, 0x000D)

    # PROJECTCONSTANTS
    buf += _record(0x000C, b'')
    # PROJECTCONSTANTSUNICODE
    buf += _record(0x003C, b'')

    # No REFERENCES section (late binding via CreateObject)

    # PROJECTMODULES
    buf += _record(0x000F, struct.pack('<I', len(modules)))
    # PROJECTCOOKIE
    buf += _record(0x0013, struct.pack('<H', 0xFFFF))

    # Module records
    for mod in modules:
        name     = mod['name']
        is_class = mod.get('class', False)

        # MODULENAME
        buf += _record(0x0019, _ansi(name))
        # MODULENAMEUNICODE (0x0031 per MS-OVBA 2.3.4.2.3.2.2)
        buf += _record(0x0031, _u16(name))
        # MODULESTREAMNAME
        buf += _record(0x001A, _ansi(name))
        # MODULESTREAMNAMERECORDUNICODE
        buf += _record(0x0032, _u16(name))
        # MODULEDOCSTRING
        buf += _record(0x001C, b'')
        # MODULEDOCSTRINGUNICODE
        buf += _record(0x0048, b'')
        # MODULEOFFSET (TextOffset=0)
        buf += _record(0x0031, struct.pack('<I', 0x00000000))
        # MODULEHELPCONTEXT
        buf += _record(0x000E, struct.pack('<I', 0x00000000))
        # MODULECOOKIE
        buf += _record(0x0013, struct.pack('<H', 0xFFFF))
        # MODULETYPE: 0x0021 for standard, 0x0022 for class
        buf += _record(0x0022 if is_class else 0x0021, b'')
        # MODULETERM
        buf += _record(0x002B, b'')

    # PROJECTMODULES terminator
    buf += _record(0x0010, b'')

    return bytes(buf)


# ---------------------------------------------------------------------------
# PROJECT stream builder
# ---------------------------------------------------------------------------

def build_project_stream(modules: list) -> bytes:
    """Build the PROJECT stream (ASCII text)."""
    lines = []
    lines.append('ID="{00000000-0000-0000-0000-000000000000}"')

    for mod in modules:
        if mod.get('class'):
            lines.append(f'Document={mod["name"]}/&H00000000')
        else:
            lines.append(f'Module={mod["name"]}')

    lines.append('HelpContextID="0"')
    lines.append('VersionCompatible32="393222000"')
    lines.append('CMG="AAAAAAAAAA=="')
    lines.append('DPB="AAAAAAAAAAAAAAAAAAAAAAAAAA=="')
    lines.append('GC="AAAAAAAAAAAAAAAAAAAA"')
    lines.append('')

    return '\r\n'.join(lines).encode('cp1252')


# ---------------------------------------------------------------------------
# CFB (OLE Compound File Binary) Builder — MS-CFB
# ---------------------------------------------------------------------------

FREESECT   = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT    = 0xFFFFFFFD
NOSTREAM   = 0xFFFFFFFF

SECTOR_SIZE           = 512
DIR_ENTRIES_PER_SECTOR = SECTOR_SIZE // 128  # = 4


class CFBWriter:
    """Minimal CFB writer for creating vbaProject.bin."""

    def __init__(self):
        self._root_streams = {}  # name -> bytes  (direct children of Root)
        self._vba_streams  = {}  # name -> bytes  (children of VBA storage)

    def add_root_stream(self, name: str, data: bytes):
        self._root_streams[name] = data

    def add_vba_stream(self, name: str, data: bytes):
        self._vba_streams[name] = data

    def build(self) -> bytes:
        """Build the complete CFB binary."""
        # ---- Build ordered entry list ----
        # Slot 0: Root Entry  (always first)
        # Slot 1: VBA storage
        # Slot 2: PROJECT stream
        # Slot 3: PROJECTwm stream
        # Slots 4+: VBA sub-streams (_VBA_PROJECT, dir, module streams...)

        entries = []

        # Root Entry (slot 0)
        entries.append({
            'name':  'Root Entry',
            'type':  5,   # root
            # CLSID {1C3B4210-F441-11CE-B9EA-00AA006B1A69} in little-endian binary
            'clsid': bytes([0x10, 0x42, 0x3B, 0x1C,
                            0x41, 0xF4,
                            0xCE, 0x11,
                            0xB9, 0xEA,
                            0x00, 0xAA, 0x00, 0x6B, 0x1A, 0x69]),
            'data':  None,
        })

        # VBA storage (slot 1)
        entries.append({
            'name':  'VBA',
            'type':  1,   # storage
            # CLSID {EA7BAE70-FB3B-11CD-A903-00AA00510EA3} in little-endian binary
            'clsid': bytes([0x70, 0xAE, 0x7B, 0xEA,
                            0x3B, 0xFB,
                            0xCD, 0x11,
                            0xA9, 0x03,
                            0x00, 0xAA, 0x00, 0x51, 0x0E, 0xA3]),
            'data':  None,
        })

        # Root-level streams (PROJECT, PROJECTwm)
        root_stream_slots = []
        for sname in ('PROJECT', 'PROJECTwm'):
            idx = len(entries)
            entries.append({
                'name':  sname,
                'type':  2,
                'clsid': b'\x00' * 16,
                'data':  self._root_streams.get(sname, b''),
            })
            root_stream_slots.append(idx)

        # VBA sub-streams
        vba_order = []
        if '_VBA_PROJECT' in self._vba_streams:
            vba_order.append('_VBA_PROJECT')
        if 'dir' in self._vba_streams:
            vba_order.append('dir')
        for n in self._vba_streams:
            if n not in ('_VBA_PROJECT', 'dir'):
                vba_order.append(n)

        vba_stream_slots = []
        for sname in vba_order:
            idx = len(entries)
            entries.append({
                'name':  sname,
                'type':  2,
                'clsid': b'\x00' * 16,
                'data':  self._vba_streams[sname],
            })
            vba_stream_slots.append(idx)

        # Pad entries to a multiple of 4 (fills complete directory sectors)
        while len(entries) % DIR_ENTRIES_PER_SECTOR != 0:
            entries.append({
                'name':  '',
                'type':  0,   # unused/empty
                'clsid': b'\x00' * 16,
                'data':  None,
            })

        # ---- Initialize tree pointers ----
        for e in entries:
            e['left']  = NOSTREAM
            e['right'] = NOSTREAM
            e['child'] = NOSTREAM
            e['start'] = ENDOFCHAIN
            e['size']  = 0

        # CFB directory entries must form a valid BST sorted by
        # (len(name), name.upper()). Build a balanced BST recursively.
        def cfb_sort_key(idx):
            n = entries[idx]['name']
            return (len(n), n.upper())

        def build_bst(slots):
            """Return root DirID of a balanced BST from a sorted list of DirIDs."""
            if not slots:
                return NOSTREAM
            mid = len(slots) // 2
            root_idx = slots[mid]
            entries[root_idx]['left']  = build_bst(slots[:mid])
            entries[root_idx]['right'] = build_bst(slots[mid + 1:])
            return root_idx

        # Root Entry's children: VBA(3) < PROJECT(7) < PROJECTwm(9) — already sorted
        root_children = sorted([1] + root_stream_slots, key=cfb_sort_key)
        entries[0]['child'] = build_bst(root_children)

        # VBA storage's children: sort by (len, upper) before building BST
        vba_sorted = sorted(vba_stream_slots, key=cfb_sort_key)
        entries[1]['child'] = build_bst(vba_sorted)

        # ---- Allocate sectors for stream data ----
        sector_data = bytearray()  # raw sector bytes
        fat = []                    # one FAT entry per sector

        def alloc_stream(data: bytes):
            if not data:
                return ENDOFCHAIN, 0
            start_sector = len(sector_data) // SECTOR_SIZE
            padded = data + b'\x00' * ((-len(data)) % SECTOR_SIZE)
            n_sectors = len(padded) // SECTOR_SIZE
            for i in range(n_sectors):
                sector_data.extend(padded[i * SECTOR_SIZE:(i + 1) * SECTOR_SIZE])
                if i < n_sectors - 1:
                    fat.append(start_sector + i + 1)
                else:
                    fat.append(ENDOFCHAIN)
            return start_sector, len(data)

        for e in entries:
            if e['type'] == 2 and e['data'] is not None and len(e['data']) > 0:
                s, sz = alloc_stream(e['data'])
                e['start'] = s
                e['size']  = sz

        # ---- Directory sectors ----
        num_data_sectors = len(sector_data) // SECTOR_SIZE
        dir_sector_start = num_data_sectors
        num_dir_sectors  = len(entries) // DIR_ENTRIES_PER_SECTOR

        dir_bytes = bytearray()
        for e in entries:
            name_utf16 = e['name'].encode('utf-16-le') if e['name'] else b''
            name_len   = len(name_utf16) + 2 if name_utf16 else 0
            name_field = (name_utf16 + b'\x00' * (64 - len(name_utf16)))[:64]

            entry_buf = bytearray(128)
            entry_buf[0:64] = name_field
            struct.pack_into('<H', entry_buf, 64, name_len)
            entry_buf[66]   = e['type']
            entry_buf[67]   = 0x01  # ColorFlag: black
            struct.pack_into('<I', entry_buf, 68, e['left'])
            struct.pack_into('<I', entry_buf, 72, e['right'])
            struct.pack_into('<I', entry_buf, 76, e['child'])
            entry_buf[80:96] = e['clsid']
            struct.pack_into('<I', entry_buf,  96, 0)   # StateBits
            struct.pack_into('<Q', entry_buf, 100, 0)   # CreatedTime
            struct.pack_into('<Q', entry_buf, 108, 0)   # ModifiedTime
            struct.pack_into('<I', entry_buf, 116, e['start'])
            struct.pack_into('<I', entry_buf, 120, e['size'])
            struct.pack_into('<I', entry_buf, 124, 0)   # SizeHigh

            dir_bytes += entry_buf

        # Add FAT entries for directory sectors
        for i in range(num_dir_sectors):
            if i < num_dir_sectors - 1:
                fat.append(dir_sector_start + i + 1)
            else:
                fat.append(ENDOFCHAIN)

        # ---- FAT sector ----
        fat_sector_start = dir_sector_start + num_dir_sectors
        num_fat_sectors  = 1  # should be enough for small files
        entries_per_fat  = SECTOR_SIZE // 4  # 128

        fat_entries = list(fat)
        for _ in range(num_fat_sectors):
            fat_entries.append(FATSECT)

        # Pad to full sector
        while len(fat_entries) % entries_per_fat != 0:
            fat_entries.append(FREESECT)

        fat_sector_bytes = b''.join(struct.pack('<I', x) for x in fat_entries)

        # ---- CFB Header ----
        header = bytearray(512)
        header[0:8] = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'
        # CLSID: 16 zeros (already zeroed)
        struct.pack_into('<H', header, 24, 0x003E)  # MinorVersion
        struct.pack_into('<H', header, 26, 0x0003)  # MajorVersion
        struct.pack_into('<H', header, 28, 0xFFFE)  # ByteOrder
        struct.pack_into('<H', header, 30, 0x0009)  # SectorSizePower (2^9=512)
        struct.pack_into('<H', header, 32, 0x0006)  # MiniSectorSizePower (2^6=64)
        # Reserved 6 bytes at 34: already zero
        struct.pack_into('<I', header, 40, 0)                  # NumberOfDirectorySectors (v3: 0)
        struct.pack_into('<I', header, 44, num_fat_sectors)    # NumberOfFATSectors
        struct.pack_into('<I', header, 48, dir_sector_start)   # FirstDirectorySectorLocation
        struct.pack_into('<I', header, 52, 0)                  # TransactionSignatureNumber
        struct.pack_into('<I', header, 56, 0x1000)             # MiniStreamCutoffSize
        struct.pack_into('<I', header, 60, ENDOFCHAIN)         # FirstMiniFATSectorLocation
        struct.pack_into('<I', header, 64, 0)                  # NumberOfMiniFATSectors
        struct.pack_into('<I', header, 68, ENDOFCHAIN)         # FirstDIFATSectorLocation
        struct.pack_into('<I', header, 72, 0)                  # NumberOfDIFATSectors
        # DIFAT[109] at offset 76: first = FAT sector, rest = FREESECT
        struct.pack_into('<I', header, 76, fat_sector_start)
        for i in range(1, 109):
            struct.pack_into('<I', header, 76 + i * 4, FREESECT)

        # ---- Assemble ----
        return bytes(header) + bytes(sector_data) + bytes(dir_bytes) + fat_sector_bytes


# ---------------------------------------------------------------------------
# Source file loading
# ---------------------------------------------------------------------------

def strip_bas_header(content: str) -> str:
    """Remove the 'Attribute VB_Name = ...' line from .bas files."""
    lines = content.splitlines(keepends=True)
    out = []
    found_name = False
    for line in lines:
        # Strip BOM from any line for comparison
        stripped = line.lstrip('﻿')
        if not found_name and stripped.lstrip().startswith('Attribute VB_Name'):
            found_name = True
            continue
        # Preserve remaining lines exactly (except strip BOM from first line)
        if not out:
            out.append(stripped)
        else:
            out.append(line)
    return ''.join(out)


def strip_cls_header(content: str) -> str:
    """Remove everything up to and including 'Attribute VB_Exposed = True'."""
    lines = content.splitlines(keepends=True)
    out = []
    past_header = False
    for line in lines:
        stripped = line.lstrip('﻿').strip()
        if not past_header:
            if stripped.startswith('Attribute VB_Exposed'):
                past_header = True
            continue
        out.append(line)
    return ''.join(out)


def load_module_source(path: str, is_cls: bool = False) -> bytes:
    """Load VBA source, strip header, return as cp1252 bytes."""
    with open(path, 'r', encoding='utf-8-sig', errors='replace') as f:
        content = f.read()

    if is_cls:
        source = strip_cls_header(content)
    else:
        source = strip_bas_header(content)

    # Normalize line endings to CRLF
    source = source.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    return source.encode('cp1252', errors='replace')


# ---------------------------------------------------------------------------
# vbaProject.bin builder
# ---------------------------------------------------------------------------

def build_vba_project(modules: list) -> bytes:
    """
    Build a complete vbaProject.bin CFB file.

    modules: list of dicts with keys:
        name   : str   — module name
        source : bytes — VBA source (cp1252 encoded)
        class  : bool  — True for class modules (e.g. ThisWorkbook)
    """
    cfb = CFBWriter()

    # PROJECT stream (root)
    project_text = build_project_stream(modules)
    cfb.add_root_stream('PROJECT', project_text)

    # PROJECTwm stream (root, empty)
    cfb.add_root_stream('PROJECTwm', b'')

    # _VBA_PROJECT stream (performance cache placeholder)
    cfb.add_vba_stream('_VBA_PROJECT', b'\xCC\x61' + b'\x00' * 6)

    # dir stream (OVBA-compressed)
    dir_uncompressed = build_dir_stream(modules)
    dir_compressed   = ovba_compress(dir_uncompressed)
    cfb.add_vba_stream('dir', dir_compressed)

    # Module streams (each: OVBA-compressed source)
    for mod in modules:
        compressed_src = ovba_compress(mod['source'])
        cfb.add_vba_stream(mod['name'], compressed_src)

    return cfb.build()


# ---------------------------------------------------------------------------
# XLSM packaging
# ---------------------------------------------------------------------------

def patch_content_types(xml_bytes: bytes) -> bytes:
    """Add vbaProject.bin Override to [Content_Types].xml."""
    vba_override = (
        '<Override PartName="/xl/vbaProject.bin" '
        'ContentType="application/vnd.ms-office.vbaProject"/>'
    )
    xml = xml_bytes.decode('utf-8')
    if 'vbaProject' not in xml:
        xml = xml.replace('</Types>', vba_override + '</Types>')
    return xml.encode('utf-8')


def patch_workbook_rels(xml_bytes: bytes) -> bytes:
    """Add vbaProject relationship to xl/_rels/workbook.xml.rels."""
    vba_rel = (
        '<Relationship Id="rId99" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vbaProject" '
        'Target="vbaProject.bin"/>'
    )
    xml = xml_bytes.decode('utf-8')
    if 'vbaProject' not in xml:
        xml = xml.replace('</Relationships>', vba_rel + '</Relationships>')
    return xml.encode('utf-8')


def create_xlsm(template_path: str, output_path: str, modules: list):
    """Create an .xlsm file by injecting VBA into the template."""
    print(f"  Building VBA project ({len(modules)} modules)...")
    vba_bin = build_vba_project(modules)
    print(f"  vbaProject.bin size: {len(vba_bin):,} bytes")

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    with zipfile.ZipFile(template_path, 'r') as zin, \
         zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:

        for item in zin.infolist():
            data = zin.read(item.filename)

            if item.filename == '[Content_Types].xml':
                data = patch_content_types(data)
            elif item.filename == 'xl/_rels/workbook.xml.rels':
                data = patch_workbook_rels(data)

            zout.writestr(item, data)

        # Add vbaProject.bin
        zout.writestr('xl/vbaProject.bin', vba_bin)

    print(f"  Created: {output_path} ({os.path.getsize(output_path):,} bytes)")


# ---------------------------------------------------------------------------
# Module definitions
# ---------------------------------------------------------------------------

def load_modules_chatbot() -> list:
    shared_names  = ['modTypes', 'modConfig', 'modPaths', 'modHttpClient',
                     'modKeyVault', 'modApiGateway']
    chatbot_names = ['modIndexReader', 'modSimilarity', 'modRagEngine',
                     'modUserProfile', 'modPiiGuard', 'modRateLimiter',
                     'modUsageLogger', 'modChatUI', 'modBoot']

    modules = []

    for name in shared_names:
        path = os.path.join(SRC_DIR, 'shared', name + '.bas')
        modules.append({'name': name, 'source': load_module_source(path), 'class': False})

    # JsonConverter from build/vendor
    jc_path = os.path.join(VENDOR_DIR, 'JsonConverter.bas')
    modules.append({'name': 'JsonConverter', 'source': load_module_source(jc_path), 'class': False})

    for name in chatbot_names:
        path = os.path.join(SRC_DIR, 'chatbot', name + '.bas')
        modules.append({'name': name, 'source': load_module_source(path), 'class': False})

    # ThisWorkbook class module
    tw_path = os.path.join(SRC_DIR, 'chatbot', 'ThisWorkbook.cls')
    modules.append({'name': 'ThisWorkbook', 'source': load_module_source(tw_path, is_cls=True), 'class': True})

    return modules


def load_modules_admin() -> list:
    shared_names = ['modTypes', 'modConfig', 'modPaths', 'modHttpClient',
                    'modKeyVault', 'modApiGateway']
    admin_names  = ['modChunker', 'modExtractor', 'modExtractorWord',
                    'modExtractorAcrobat', 'modExtractorExcel',
                    'modIndexWriter', 'modKnowledgeBuilder',
                    'modKeyEnroller', 'modUsageAggregator']

    modules = []

    for name in shared_names:
        path = os.path.join(SRC_DIR, 'shared', name + '.bas')
        modules.append({'name': name, 'source': load_module_source(path), 'class': False})

    # JsonConverter from build/vendor
    jc_path = os.path.join(VENDOR_DIR, 'JsonConverter.bas')
    modules.append({'name': 'JsonConverter', 'source': load_module_source(jc_path), 'class': False})

    for name in admin_names:
        path = os.path.join(SRC_DIR, 'admin', name + '.bas')
        modules.append({'name': name, 'source': load_module_source(path), 'class': False})

    return modules


# ---------------------------------------------------------------------------
# Main
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

    # --- Chatbot ---
    print("Building Chatbot.xlsm...")
    chatbot_modules = load_modules_chatbot()
    print(f"  Modules ({len(chatbot_modules)}): {[m['name'] for m in chatbot_modules]}")
    create_xlsm(
        TEMPLATE,
        os.path.join(DIST_DIR, 'Chatbot.xlsm'),
        chatbot_modules,
    )

    print()

    # --- Admin ---
    print("Building Admin_KnowledgeBuilder.xlsm...")
    admin_modules = load_modules_admin()
    print(f"  Modules ({len(admin_modules)}): {[m['name'] for m in admin_modules]}")
    create_xlsm(
        TEMPLATE,
        os.path.join(DIST_DIR, 'Admin_KnowledgeBuilder.xlsm'),
        admin_modules,
    )

    print()
    print("Done.")


if __name__ == '__main__':
    main()
