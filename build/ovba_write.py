#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ovba_write.py -- MS-OVBA 準拠の vbaProject.bin **ライター**(裁定書27 W9-A)。

出典(移植元・R35 波1 で取り込み):
    姉妹PJ riskconsulting のリポジトリ
      ブランチ: claude/ai-risk-consulting-spec-review-91o5y5
      コミット: cf628c1
    build/ovba_write.py(595行)をそのまま取り込んだもの。**本体には手を
    入れない**(向こうが vbaexport.cxx / MS-OVBA ABNF 突合で修正版を出す
    可能性があり、来たら差し替えるだけで済むようにするため)。MyBookshelf
    固有の呼び出し(ThisWorkbook 本文・CP932 replace 等)は呼び出し側
    (build/build_mybookshelf.py)で吸収する。

背景(なぜ新設したか):
    従来の配布方式は template_skeleton.xlsm 由来の vbaProject.bin
    (ThisWorkbook 1本だけ)に自己インストーラを外科パッチし、隠しシート
    vba_src に置いたソースを起動時に VBComponents.Add + AddFromString で
    VBE へ注入する方式だった。この形は 2026-09-02 に社内AVの AMSI で検知され
    (VDI強制停止)、採用禁止になった(裁定書27)。
    本モジュールは「モジュールが最初から入った正規の vbaProject.bin を
    Linux 上でゼロから書き出す」ためのライターである。これにより配布物から
    隠しシート・VBProject 参照・AddFromString が消える。

参照仕様:
    [MS-OVBA] 2.3.1 PROJECT ストリーム / 2.3.3 PROJECTwm ストリーム /
              2.3.4.2 dir ストリーム / 2.3.4.3 モジュールストリーム /
              2.4.1 圧縮(build/ovba.py の ovba_compress を使う)
    [MS-CFB]  Compound File Binary(512B セクタ・FAT/DIFAT/MiniFAT/
              ディレクトリ)。本ファイルに最小限のライターを実装した。

設計方針(「写す」と「作る」の境界):
    ・PROJECTINFORMATION(SysKind/LCID/CodePage/Name/HelpFile/Constants/
      Version)と PROJECTREFERENCES(stdole/Office/Excel/VBA 等)は
      **template_skeleton.xlsm の dir からバイト列のまま写す**。新規参照は
      足さない(裁定書27)。
    ・PROJECTMODULES(モジュール台帳)だけを本モジュールが生成する。
    ・PROJECT ストリームの ID/CMG/DPB/GC/Name/HelpContextID/
      VersionCompatible32 行は template から写し、Document/Module/Class 行と
      [Workspace] だけを生成する。
    ・_VBA_PROJECT は template のものをそのまま写す。各モジュールの
      MODULEOFFSET を 0 にし、モジュールストリームには PerformanceCache を
      置かない(=p-code 無効)ので、Excel/LO はソースから再コンパイルする。
      これは従来の外科パッチと同じ状態である。

このモジュールは openpyxl を import しない(純粋にバイト列操作のみ)。
依存: 標準ライブラリ + 同ディレクトリの ovba.py。
"""

from __future__ import annotations

import struct

import ovba

# ---------------------------------------------------------------------------
# CFB 定数(ovba.py と同じ値。読み書きで食い違わないよう再輸入する)
# ---------------------------------------------------------------------------
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC
FREESECT = 0xFFFFFFFF
NOSTREAM = 0xFFFFFFFF

SECTOR_SIZE = 512
MINI_SECTOR = 64
MINI_CUTOFF = 4096

CFB_MAGIC = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'

# dir ストリームのレコードID(MS-OVBA 2.3.4.2)
REC_PROJECTMODULES = 0x000F
REC_PROJECTCOOKIE = 0x0013
REC_MODULENAME = 0x0019
REC_MODULENAMEUNICODE = 0x0047
REC_MODULESTREAMNAME = 0x001A
REC_MODULESTREAMNAME_UNI = 0x0032
REC_MODULEDOCSTRING = 0x001C
REC_MODULEDOCSTRING_UNI = 0x0048
REC_MODULEOFFSET = 0x0031
REC_MODULEHELPCONTEXT = 0x001E
REC_MODULECOOKIE = 0x002C
REC_MODULETYPE_PROCEDURAL = 0x0021
REC_MODULETYPE_DOCUMENT = 0x0022
REC_MODULE_TERMINATOR = 0x002B
REC_DIR_TERMINATOR = 0x0010
REC_PROJECTVERSION = 0x0009


class OvbaWriteError(Exception):
    """vbaProject.bin の組み立てに失敗したことを表す例外。"""


# ===========================================================================
# 1. dir ストリーム
# ===========================================================================
def iter_dir_records(dir_dec: bytes):
    """解凍済み dir ストリームを (offset, id, size, body) で順に返す。

    PROJECTVERSION(0x0009)だけは Size フィールドが 4 と書かれているのに
    実データが 6 バイトある(MS-OVBA 2.3.4.2.1.10 の既知の例外)。
    """
    i = 0
    n = len(dir_dec)
    while i + 6 <= n:
        rid = struct.unpack('<H', dir_dec[i:i + 2])[0]
        size = struct.unpack('<I', dir_dec[i + 2:i + 6])[0]
        real = 6 if rid == REC_PROJECTVERSION else size
        body = dir_dec[i + 6:i + 6 + real]
        yield (i, rid, size, body)
        i += 6 + real


def split_template_dir(dir_dec: bytes) -> bytes:
    """template の dir から PROJECTINFORMATION+PROJECTREFERENCES 部分
    (= PROJECTMODULES レコードの直前まで)をバイト列のまま切り出す。"""
    for off, rid, _size, _body in iter_dir_records(dir_dec):
        if rid == REC_PROJECTMODULES:
            return dir_dec[:off]
    raise OvbaWriteError("template の dir に PROJECTMODULES(0x000F)がありません")


def template_project_cookie(dir_dec: bytes) -> bytes:
    """template の PROJECTCOOKIE(0x0013)の2バイトを返す。"""
    for _off, rid, _size, body in iter_dir_records(dir_dec):
        if rid == REC_PROJECTCOOKIE:
            return body
    raise OvbaWriteError("template の dir に PROJECTCOOKIE(0x0013)がありません")


def _rec(rid: int, body: bytes) -> bytes:
    return struct.pack('<HI', rid, len(body)) + body


def _mbcs(name: str, codepage: str) -> bytes:
    try:
        return name.encode(codepage)
    except UnicodeEncodeError as exc:
        raise OvbaWriteError(
            "モジュール名 %r をコードページ %s で表現できません" % (name, codepage)) from exc


def build_modules_section(modules, cookie: bytes, codepage: str = "cp932") -> bytes:
    """PROJECTMODULES セクション(MS-OVBA 2.3.4.2.3)を組み立てる。

    modules: VbaModule のリスト。
    MODULEOFFSET は常に 0(=モジュールストリームに PerformanceCache を
    置かず、先頭から圧縮ソースが始まる)。
    """
    out = bytearray()
    out += _rec(REC_PROJECTMODULES, struct.pack('<H', len(modules)))
    out += _rec(REC_PROJECTCOOKIE, cookie)
    for m in modules:
        nm = _mbcs(m.name, codepage)
        nu = m.name.encode('utf-16-le')
        sn = _mbcs(m.stream_name, codepage)
        su = m.stream_name.encode('utf-16-le')
        out += _rec(REC_MODULENAME, nm)
        out += _rec(REC_MODULENAMEUNICODE, nu)
        out += _rec(REC_MODULESTREAMNAME, sn)
        out += _rec(REC_MODULESTREAMNAME_UNI, su)
        out += _rec(REC_MODULEDOCSTRING, b'')
        out += _rec(REC_MODULEDOCSTRING_UNI, b'')
        out += _rec(REC_MODULEOFFSET, struct.pack('<I', 0))
        out += _rec(REC_MODULEHELPCONTEXT, struct.pack('<I', 0))
        out += _rec(REC_MODULECOOKIE, struct.pack('<H', m.cookie))
        # [MS-OVBA] 2.3.4.2.3.2.8 MODULETYPE:
        #   0x0021 = 標準(procedural)モジュール
        #   0x0022 = document / class / designer モジュール
        # クラスモジュールは 0x0022 側である(標準と同じ 0x0021 にすると、
        # PROJECT の Class= 行と食い違ってVBEが読み違える)。
        # document と class の区別は PROJECT ストリームの Document= / Class= 行が持つ。
        out += _rec(REC_MODULETYPE_PROCEDURAL if m.module_type == "std"
                    else REC_MODULETYPE_DOCUMENT, b'')
        out += _rec(REC_MODULE_TERMINATOR, b'')
    out += _rec(REC_DIR_TERMINATOR, b'')
    return bytes(out)


# ===========================================================================
# 2. PROJECT / PROJECTwm ストリーム
# ===========================================================================
# PROJECT ストリームから「写す」行(MS-OVBA 2.3.1)。ここに無い行
# (Document=/Module=/Class=/[Workspace])はモジュール台帳から生成する。
PROJECT_COPY_KEYS = (
    "ID", "Name", "HelpContextID", "Description", "VersionCompatible32",
    "CMG", "DPB", "GC",
)


def parse_template_project(text: str) -> dict:
    """template の PROJECT ストリームから写す行を key->行全体 で拾う。"""
    copied = {}
    host_ext = []
    in_host = False
    for line in text.splitlines():
        s = line.strip()
        if s.startswith("["):
            in_host = (s.lower() == "[host extender info]")
            continue
        if in_host:
            if s:
                host_ext.append(s)
            continue
        if "=" in s:
            key = s.split("=", 1)[0]
            if key in PROJECT_COPY_KEYS:
                copied[key] = s
    return {"copied": copied, "host_extender": host_ext}


def build_project_stream(modules, tmpl: dict) -> bytes:
    """PROJECT ストリーム(CP932/CRLF)を組み立てる。"""
    copied = tmpl["copied"]
    if "ID" not in copied:
        raise OvbaWriteError("template の PROJECT に ID 行がありません")
    lines = [copied["ID"]]
    for m in modules:
        if m.is_document:
            lines.append("Document=%s/&H00000000" % m.name)
        elif m.module_type == "class":
            lines.append("Class=%s" % m.name)
        else:
            lines.append("Module=%s" % m.name)
    for key in ("Name", "HelpContextID", "Description", "VersionCompatible32",
                "CMG", "DPB", "GC"):
        if key in copied:
            lines.append(copied[key])
    lines.append("")
    lines.append("[Host Extender Info]")
    lines.extend(tmpl["host_extender"])
    lines.append("")
    lines.append("[Workspace]")
    for m in modules:
        if m.is_document:
            lines.append("%s=0, 0, 0, 0, C" % m.name)
        else:
            lines.append("%s=60, 60, 1354, 680, " % m.name)
    lines.append("")
    return ("\r\n".join(lines)).encode("cp932")


def build_projectwm(modules, codepage: str = "cp932") -> bytes:
    """PROJECTwm(MS-OVBA 2.3.3): MBCS名\\0 + UTF-16名\\0\\0 の並び + 終端 \\0\\0。"""
    out = bytearray()
    for m in modules:
        out += _mbcs(m.name, codepage) + b'\x00'
        out += m.name.encode('utf-16-le') + b'\x00\x00'
    out += b'\x00\x00'
    return bytes(out)


# ===========================================================================
# 3. CFB ライター
# ===========================================================================
def _cfb_sort_key(name: str):
    """MS-CFB のディレクトリ兄弟順(名前長 → 大文字化して比較)。"""
    return (len(name), name.upper())


class _DirEntry:
    __slots__ = ("name", "etype", "data", "start", "size",
                 "left", "right", "child")

    def __init__(self, name, etype, data=b''):
        if len(name) > 31:
            raise OvbaWriteError("CFB のエントリ名が31文字を超えています: %r" % name)
        self.name = name
        self.etype = etype          # 1=storage, 2=stream, 5=root
        self.data = data
        self.start = ENDOFCHAIN
        self.size = len(data)
        self.left = NOSTREAM
        self.right = NOSTREAM
        self.child = NOSTREAM


def _build_tree(entries, ids):
    """ソート済みの子エントリ列から平衡二分木を作り、根のIDを返す。"""
    if not ids:
        return NOSTREAM
    mid = len(ids) // 2
    root = ids[mid]
    entries[root].left = _build_tree(entries, ids[:mid])
    entries[root].right = _build_tree(entries, ids[mid + 1:])
    return root


def write_cfb(root_children) -> bytes:
    """CFB(v3・512Bセクタ)を書き出す。

    root_children: [(name, etype, data_or_children)] の入れ子。
      etype=1(storage) のとき第3要素は同じ形式の子リスト。
      etype=2(stream)  のとき第3要素は bytes。
    """
    entries = [_DirEntry("Root Entry", 5)]

    def add(children, parent_index):
        ids = []
        for name, etype, payload in children:
            e = _DirEntry(name, etype, payload if etype == 2 else b'')
            entries.append(e)
            idx = len(entries) - 1
            ids.append((name, idx))
            if etype == 1:
                add(payload, idx)
        ids.sort(key=lambda t: _cfb_sort_key(t[0]))
        entries[parent_index].child = _build_tree(entries, [i for _n, i in ids])

    add(root_children, 0)

    # --- ミニストリーム(4096バイト未満のストリーム)を先に配置する ---
    minifat = []
    mini_blob = bytearray()
    for e in entries:
        if e.etype != 2:
            continue
        if e.size == 0:
            e.start = ENDOFCHAIN
            continue
        if e.size >= MINI_CUTOFF:
            continue
        start = len(mini_blob) // MINI_SECTOR
        nsec = (e.size + MINI_SECTOR - 1) // MINI_SECTOR
        for k in range(nsec):
            minifat.append(start + k + 1 if k < nsec - 1 else ENDOFCHAIN)
        mini_blob += e.data.ljust(nsec * MINI_SECTOR, b'\x00')
        e.start = start

    sectors = []   # 512B ペイロード
    fat = []       # sectors と同じ長さ

    def alloc(data: bytes) -> int:
        if not data:
            return ENDOFCHAIN
        start = len(sectors)
        nsec = (len(data) + SECTOR_SIZE - 1) // SECTOR_SIZE
        for k in range(nsec):
            sectors.append(data[k * SECTOR_SIZE:(k + 1) * SECTOR_SIZE]
                           .ljust(SECTOR_SIZE, b'\x00'))
            fat.append(start + k + 1 if k < nsec - 1 else ENDOFCHAIN)
        return start

    # --- 大きいストリーム ---
    for e in entries:
        if e.etype == 2 and e.size >= MINI_CUTOFF:
            e.start = alloc(e.data)

    # --- MiniFAT ---
    minifat_start = ENDOFCHAIN
    minifat_count = 0
    if minifat:
        per = SECTOR_SIZE // 4
        padded = minifat + [FREESECT] * ((-len(minifat)) % per)
        blob = b''.join(struct.pack('<I', v) for v in padded)
        minifat_start = alloc(blob)
        minifat_count = len(padded) // per

    # --- ミニストリームコンテナ(Root Entry のストリーム) ---
    entries[0].size = len(mini_blob)
    entries[0].start = alloc(bytes(mini_blob)) if mini_blob else ENDOFCHAIN

    # --- ディレクトリ ---
    dir_blob = bytearray()
    for e in entries:
        nm = e.name.encode('utf-16-le') + b'\x00\x00'
        rec = bytearray(128)
        rec[0:len(nm)] = nm
        struct.pack_into('<H', rec, 64, len(nm))
        rec[66] = e.etype
        rec[67] = 1                      # colorFlag: black(検証しない実装が大半)
        struct.pack_into('<I', rec, 68, e.left)
        struct.pack_into('<I', rec, 72, e.right)
        struct.pack_into('<I', rec, 76, e.child)
        struct.pack_into('<I', rec, 116, 0 if e.etype == 1 else e.start)
        struct.pack_into('<Q', rec, 120, 0 if e.etype == 1 else e.size)
        dir_blob += rec
    # ディレクトリセクタは 4 エントリ単位で埋める(空きは FREE エントリ)
    while len(dir_blob) % SECTOR_SIZE:
        rec = bytearray(128)
        struct.pack_into('<I', rec, 68, NOSTREAM)
        struct.pack_into('<I', rec, 72, NOSTREAM)
        struct.pack_into('<I', rec, 76, NOSTREAM)
        dir_blob += rec
    dir_start = alloc(bytes(dir_blob))

    # --- FAT セクタ(自分自身も FAT に載るので反復して本数を決める) ---
    per = SECTOR_SIZE // 4
    nfat = 1
    while True:
        need = -(-(len(sectors) + nfat) // per)
        if need == nfat:
            break
        nfat = need
    if nfat > 109:
        raise OvbaWriteError(
            "FAT セクタが %d 本必要で DIFAT がヘッダに収まりません(未対応)" % nfat)
    fat_sector_ids = list(range(len(sectors), len(sectors) + nfat))
    for _ in range(nfat):
        sectors.append(b'\x00' * SECTOR_SIZE)
        fat.append(FATSECT)

    fat_padded = fat + [FREESECT] * (nfat * per - len(fat))
    fat_blob = b''.join(struct.pack('<I', v) for v in fat_padded)
    for k, sid in enumerate(fat_sector_ids):
        sectors[sid] = fat_blob[k * SECTOR_SIZE:(k + 1) * SECTOR_SIZE]

    # --- ヘッダ ---
    hdr = bytearray(SECTOR_SIZE)
    hdr[0:8] = CFB_MAGIC
    struct.pack_into('<H', hdr, 24, 0x003E)     # minor version
    struct.pack_into('<H', hdr, 26, 0x0003)     # major version (v3)
    struct.pack_into('<H', hdr, 28, 0xFFFE)     # byte order (little endian)
    struct.pack_into('<H', hdr, 30, 9)          # sector shift (512)
    struct.pack_into('<H', hdr, 32, 6)          # mini sector shift (64)
    struct.pack_into('<I', hdr, 44, nfat)
    struct.pack_into('<I', hdr, 48, dir_start)
    struct.pack_into('<I', hdr, 56, MINI_CUTOFF)
    struct.pack_into('<I', hdr, 60, minifat_start)
    struct.pack_into('<I', hdr, 64, minifat_count)
    struct.pack_into('<I', hdr, 68, ENDOFCHAIN)  # first DIFAT sector
    struct.pack_into('<I', hdr, 72, 0)           # number of DIFAT sectors
    for k in range(109):
        v = fat_sector_ids[k] if k < nfat else FREESECT
        struct.pack_into('<I', hdr, 76 + 4 * k, v)

    return bytes(hdr) + b''.join(sectors)


# ===========================================================================
# 4. モジュール記述と組み立て本体
# ===========================================================================
class VbaModule:
    """vbaProject.bin へ焼き込む1モジュール。

    name        : VBE 上のモジュール名(=dir の MODULENAME)
    source      : ソース本文(CP932・CRLF・末尾改行あり。Attribute 行は含めない)
    module_type : "std" | "class" | "document"
    """

    def __init__(self, name, source: bytes, module_type: str = "std",
                 stream_name: str | None = None, cookie: int = 0xFFFF):
        if module_type not in ("std", "class", "document"):
            raise OvbaWriteError("未知の module_type: %r" % module_type)
        self.name = name
        self.source = source
        self.module_type = module_type
        self.stream_name = stream_name or name
        self.cookie = cookie

    @property
    def is_document(self) -> bool:
        return self.module_type == "document"

    def stream_bytes(self) -> bytes:
        """モジュールストリーム本体(MODULEOFFSET=0 なので圧縮ソースのみ)。

        全モジュールで合計1MB超を圧縮するため高速版を使う(ovba.py 参照)。
        解凍結果は素朴版と同一で、読み戻し検証がそれを毎回確認する。
        """
        return ovba.ovba_compress(self.source, fast=True)


# VBE が .bas/.cls をエクスポートするときに先頭へ付ける属性行。
# vbaProject.bin のモジュールストリームには「モジュール属性 + ソース」が
# そのまま入る(Attribute 行はソースの一部である)。
_ATTR_STD = 'Attribute VB_Name = "%s"\r\n'
_ATTR_CLASS = (
    'Attribute VB_Name = "%s"\r\n'
    'Attribute VB_GlobalNameSpace = False\r\n'
    'Attribute VB_Creatable = False\r\n'
    'Attribute VB_PredeclaredId = False\r\n'
    'Attribute VB_Exposed = False\r\n'
)
_ATTR_DOCUMENT = (
    'Attribute VB_Name = "%s"\r\n'
    'Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"\r\n'
    'Attribute VB_GlobalNameSpace = False\r\n'
    'Attribute VB_Creatable = False\r\n'
    'Attribute VB_PredeclaredId = True\r\n'
    'Attribute VB_Exposed = True\r\n'
    'Attribute VB_TemplateDerived = False\r\n'
    'Attribute VB_Customizable = True\r\n'
)


def module_stream_source(name: str, body: str, module_type: str,
                         attributes: str | None = None) -> bytes:
    """モジュールストリームへ入れるソース(属性行 + 本文)を CP932/CRLF で作る。

    attributes を与えた場合はそれを使う(.cls の元ヘッダを保つとき用)。
    """
    if attributes is None:
        if module_type == "class":
            attributes = _ATTR_CLASS % name
        elif module_type == "document":
            attributes = _ATTR_DOCUMENT % name
        else:
            attributes = _ATTR_STD % name
    text = body.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    if text and not text.endswith("\r\n"):
        text += "\r\n"
    return (attributes + text).encode("cp932")


def build_vba_project(template_bin: bytes, modules,
                      vba_project_stream: bytes | None = None) -> bytes:
    """template の vbaProject.bin から情報を写しつつ、modules を焼き込んだ
    完成品 vbaProject.bin を返す。

    modules は VbaModule のリスト。並び順がそのまま dir/PROJECT の並びになる
    (慣例として ThisWorkbook などの document module を先頭に置く)。
    vba_project_stream を渡すと _VBA_PROJECT ストリームをその内容で置き換える。
    """
    if not modules:
        raise OvbaWriteError("modules が空です")
    names = [m.name for m in modules]
    if len(set(n.upper() for n in names)) != len(names):
        raise OvbaWriteError("モジュール名が重複しています: %r" % (names,))

    skel = ovba.CFBReader(template_bin)
    tmpl_dir = ovba.ovba_decompress(skel.read("dir"))
    prefix = split_template_dir(tmpl_dir)
    cookie = template_project_cookie(tmpl_dir)

    dir_dec = prefix + build_modules_section(modules, cookie)
    dir_stream = ovba.ovba_compress(dir_dec, fast=True)

    project_tmpl = parse_template_project(
        skel.read("PROJECT").decode("cp932", errors="replace"))
    project_stream = build_project_stream(modules, project_tmpl)
    projectwm = build_projectwm(modules)
    # _VBA_PROJECT は template のものを写す。呼び出し側が「PerformanceCache 無害化
    # 済み(Version=0xFFFF)」の版を渡してきたらそれを優先する。
    vba_project = (vba_project_stream if vba_project_stream is not None
                   else skel.read("_VBA_PROJECT"))

    vba_children = [(m.stream_name, 2, m.stream_bytes()) for m in modules]
    vba_children.append(("_VBA_PROJECT", 2, vba_project))
    vba_children.append(("dir", 2, dir_stream))

    return write_cfb([
        ("VBA", 1, vba_children),
        ("PROJECT", 2, project_stream),
        ("PROJECTwm", 2, projectwm),
    ])


# ===========================================================================
# 5. 読み戻し(検証用)
# ===========================================================================
def read_modules(vba_bin: bytes) -> dict:
    """完成品 vbaProject.bin を読み戻し、モジュール名 -> ソース(bytes)を返す。

    tools/bin_roundtrip.py と build_rpn.py の自己検証がこれを使う。
    dir の MODULEOFFSET を尊重するので、PerformanceCache 付きの bin も読める。
    """
    cfb = ovba.CFBReader(vba_bin)
    dir_dec = ovba.ovba_decompress(cfb.read("dir"))
    out = {}
    cur = None
    for _off, rid, _size, body in iter_dir_records(dir_dec):
        if rid == REC_MODULENAME:
            cur = {"name": body.decode("cp932"), "offset": 0, "type": None}
        elif cur is None:
            continue
        elif rid == REC_MODULESTREAMNAME:
            cur["stream"] = body.decode("cp932")
        elif rid == REC_MODULEOFFSET:
            cur["offset"] = struct.unpack('<I', body)[0]
        elif rid in (REC_MODULETYPE_PROCEDURAL, REC_MODULETYPE_DOCUMENT):
            # 0x0022 は document と class の両方に使われるので、dir だけでは
            # 区別できない。区別は PROJECT ストリームの Document= 行が持つ
            # (下で上書きする)。
            cur["type"] = ("document" if rid == REC_MODULETYPE_DOCUMENT
                           else "procedural")
        elif rid == REC_MODULE_TERMINATOR:
            raw = cfb.read(cur.get("stream", cur["name"]))
            src = ovba.ovba_decompress(raw)[cur["offset"]:]
            out[cur["name"]] = {"source": src, "type": cur["type"]}
            cur = None
    # PROJECT ストリームの Document= 行だけが「本当の document module」である。
    # 0x0022 だが Document= に無いものは class module。
    try:
        proj = cfb.read("PROJECT").decode("cp932", errors="replace")
    except KeyError:
        proj = ""
    doc_names = set()
    for line in proj.splitlines():
        if line.startswith("Document="):
            doc_names.add(line[len("Document="):].split("/")[0].strip())
    for nm, info in out.items():
        if info["type"] == "document" and nm not in doc_names:
            info["type"] = "class"
    return out


def strip_attribute_lines(src: bytes) -> bytes:
    """モジュールソース先頭の `Attribute VB_...` 行を落とした残りを返す。"""
    lines = src.split(b"\r\n")
    i = 0
    while i < len(lines) and lines[i].startswith(b"Attribute "):
        i += 1
    return b"\r\n".join(lines[i:])
