#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ovba.py -- MS-OVBA (VBAプロジェクト格納バイナリ) の圧縮/解凍と、
          CFB (OLE Compound File Binary = vbaProject.bin の入れ物形式) の
          最小限リーダー。

役割:
    build_mybookshelf.py が template_skeleton.xlsm 内の vbaProject.bin を
    「外科手術的パッチ」(ThisWorkbook ストリームの中身だけ差し替え、
    dir ストリームの ThisWorkbook.MOFFSET を 0 に書き換え)するために
    必要な低レベル関数一式を提供する。

設計判断(出自):
    /home/user/notebook/build/make_xlsm.py の OVBA 圧縮/解凍関数・CFB
    リーダー(CFBReader)をそのまま抽出した。同一リポジトリ内の既存実証
    済みコードであり、著作はコピーで良いという判断(担当1-Hの指示)による。
    OVBA 空チャンク pad ヘルパー(pad_to_exact)は
    /home/user/notebook/build/build_chatbot_v2.py 由来だが、性質としては
    「OVBA バイト列の圧縮補助」でしかなく ThisWorkbook 固有のロジックを
    含まないため、本モジュールに同居させる方が自己完結性が高いと判断した。
    自己インストーラのソース文字列やシート生成などプロダクト固有のロジック
    は一切ここに置かない(すべて build_mybookshelf.py 側)。

    mybookshelf/ は V2 (chatbot_v2) とは別プロダクトであり、
    /home/user/notebook/build/ 配下のファイルは一切 import/変更しない
    (このファイル自体が「自己完結コピー」であることの理由)。

このモジュールは openpyxl 等を import しない(純粋にバイト列操作のみ)。
依存: 標準ライブラリ(struct)のみ。
"""

from __future__ import annotations

import struct

# ---------------------------------------------------------------------------
# CFB (OLE Compound File Binary) 構造定数
# ---------------------------------------------------------------------------
FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF

SECTOR_SIZE = 512
MINI_CUTOFF = 4096
MINI_SECTOR = 64


# ---------------------------------------------------------------------------
# MS-OVBA 圧縮 (2.4節) -- Excel が内部で使う解凍アルゴリズムと対称。
# oletools.olevba.decompress_stream と往復可能なことを確認済み(V2実証済み)。
# ---------------------------------------------------------------------------
def _ovba_compress_chunk_fast(data: bytes) -> bytes:
    """_ovba_compress_chunk と同じ形式を出す、3バイトハッシュ表つきの高速版。

    なぜ要るか(裁定書27 W9-A):
        自己インストーラ方式では圧縮対象が ThisWorkbook(約1KB)と dir だけ
        だったので素朴な全窓走査で足りていた。配布方式Bは**全モジュール
        (約1.5MB)**を圧縮するため、位置ごとに窓(最大4,096B)を舐める
        O(n^2) では実用時間に収まらない。
        そこで「3バイト一致の候補位置」だけをハッシュ表から引く。
        候補が無い位置(=VBAソースでは大多数)が O(1) で片付く。

    素朴版と出力バイト列が一致するとは限らない(同じ長さの一致が複数ある
    とき、どの位置を選ぶかが違いうる)。**解凍結果は必ず一致する**ので
    データとしては等価だが、`pad_to_exact` は「圧縮後サイズがちょうど
    目標値に届くか」に依存するため、外科パッチ経路(installer)は素朴版の
    ままにしてある(ovba_compress(fast=False) が既定)。
    """
    assert 1 <= len(data) <= 4096
    pos = 0
    out = bytearray()
    table: dict[bytes, list[int]] = {}
    n = len(data)
    MAX_CANDIDATES = 96          # 直近から見る候補数の上限(速度と圧縮率の折衷)
    while pos < n:
        flag_pos = len(out)
        out.append(0)
        flag = 0
        for bit in range(8):
            if pos >= n:
                break
            if pos <= 16:
                lbits, obits = 12, 4
            elif pos <= 32:
                lbits, obits = 11, 5
            elif pos <= 64:
                lbits, obits = 10, 6
            elif pos <= 128:
                lbits, obits = 9, 7
            elif pos <= 256:
                lbits, obits = 8, 8
            elif pos <= 512:
                lbits, obits = 7, 9
            elif pos <= 1024:
                lbits, obits = 6, 10
            elif pos <= 2048:
                lbits, obits = 5, 11
            else:
                lbits, obits = 4, 12

            max_len = (1 << lbits) - 1 + 3
            best_len = 0
            best_off = 0
            window_start = max(0, pos - (1 << obits))
            if pos > 0 and (n - pos) >= 3:
                cands = table.get(data[pos:pos + 3])
                if cands:
                    for k in range(len(cands) - 1, -1, -1):
                        j = cands[k]
                        if j < window_start:
                            break
                        if len(cands) - k > MAX_CANDIDATES:
                            break
                        m = 0
                        while (m < max_len and pos + m < n
                               and data[j + m] == data[pos + m]):
                            m += 1
                        if m > best_len:
                            best_len = m
                            best_off = j
                            if m >= max_len:
                                break

            if best_len >= 3:
                flag |= (1 << bit)
                off_val = pos - best_off - 1
                len_val = best_len - 3
                out += struct.pack('<H', (off_val << lbits) | len_val)
                for q in range(pos, pos + best_len):
                    if q + 3 <= n:
                        table.setdefault(data[q:q + 3], []).append(q)
                pos += best_len
            else:
                if pos + 3 <= n:
                    table.setdefault(data[pos:pos + 3], []).append(pos)
                out.append(data[pos])
                pos += 1
        out[flag_pos] = flag
    return bytes(out)


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
            # ビット幅は現在の解凍済み位置(pos)に応じて変わる(MS-OVBA仕様)。
            if pos <= 16:
                lbits, obits = 12, 4
            elif pos <= 32:
                lbits, obits = 11, 5
            elif pos <= 64:
                lbits, obits = 10, 6
            elif pos <= 128:
                lbits, obits = 9, 7
            elif pos <= 256:
                lbits, obits = 8, 8
            elif pos <= 512:
                lbits, obits = 7, 9
            elif pos <= 1024:
                lbits, obits = 6, 10
            elif pos <= 2048:
                lbits, obits = 5, 11
            else:
                lbits, obits = 4, 12

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


def ovba_compress(data: bytes, fast: bool = False) -> bytes:
    """バイト列を OVBA CompressedContainer に圧縮する。

    fast=True で3バイトハッシュ表つきの高速版を使う(配布方式Bの全モジュール
    圧縮用。解凍結果は同一だが圧縮後のバイト列は素朴版と一致しない)。
    外科パッチ経路(pad_to_exact で目標長へ寄せる)は既定の素朴版を使う。
    """
    chunk_fn = _ovba_compress_chunk_fast if fast else _ovba_compress_chunk
    result = bytearray()
    result.append(0x01)  # SignatureByte
    offset = 0
    while offset < len(data):
        chunk = data[offset:offset + 4096]
        compressed = chunk_fn(chunk)
        if len(chunk) == 4096 and len(compressed) >= 4096:
            # 圧縮しても縮まらない場合は非圧縮チャンクを使う
            # (signature=0b011, flag=0, size field = 4095)。
            result += struct.pack('<H', 0x3FFF)
            result += chunk
        else:
            header = 0xB000 | (len(compressed) - 1)
            result += struct.pack('<H', header)
            result += compressed
        offset += 4096
    return bytes(result)


def ovba_decompress(data: bytes) -> bytes:
    """OVBA CompressedContainer (MS-OVBA 2.4) を解凍する。"""
    if not data or data[0] != 0x01:
        raise ValueError("Bad OVBA signature byte")
    out = bytearray()
    i = 1
    while i < len(data):
        header = struct.unpack('<H', data[i:i + 2])[0]
        i += 2
        chunk_size = (header & 0x0FFF) + 3
        chunk_flag = (header >> 15) & 0x01
        body = data[i:i + chunk_size - 2]
        i += chunk_size - 2
        if chunk_flag == 0:
            # 非圧縮チャンク(必ず4096バイト)
            out += body
            continue
        # 圧縮チャンク
        chunk_start = len(out)
        j = 0
        while j < len(body):
            flag = body[j]
            j += 1
            for bit in range(8):
                if j >= len(body):
                    break
                if not (flag & (1 << bit)):
                    out.append(body[j])
                    j += 1
                else:
                    pos = len(out) - chunk_start
                    if pos <= 16:
                        lbits = 12
                    elif pos <= 32:
                        lbits = 11
                    elif pos <= 64:
                        lbits = 10
                    elif pos <= 128:
                        lbits = 9
                    elif pos <= 256:
                        lbits = 8
                    elif pos <= 512:
                        lbits = 7
                    elif pos <= 1024:
                        lbits = 6
                    elif pos <= 2048:
                        lbits = 5
                    else:
                        lbits = 4
                    token = struct.unpack('<H', body[j:j + 2])[0]
                    j += 2
                    length = (token & ((1 << lbits) - 1)) + 3
                    offset = (token >> lbits) + 1
                    src = len(out) - offset
                    for k in range(length):
                        out.append(out[src + k])
    return bytes(out)


# ---------------------------------------------------------------------------
# 空チャンク pad ヘルパー -- In-place ストリーム置換(olefile.write_stream)は
# バイト長が完全一致しないと使えないため、3バイト/5バイトの「解凍すると
# ゼロバイトになる」空チャンクで圧縮後サイズをちょうど目標長に合わせる。
# ---------------------------------------------------------------------------
_EMPTY3 = struct.pack('<H', 0xB000) + b'\x00'
_PAD5 = struct.pack('<H', 0xB002) + b'\x00\x00\x00'


def pad_to_exact(compressed: bytes, target: int) -> bytes:
    """compressed の末尾に空チャンクを継ぎ足し、ちょうど target バイトにする。"""
    diff = target - len(compressed)
    if diff < 0:
        raise ValueError(f"compressed ({len(compressed)}) > target ({target})")
    if diff == 0:
        return compressed
    for n5 in range(diff // 5 + 1):
        rem = diff - n5 * 5
        if rem >= 0 and rem % 3 == 0:
            return compressed + (_PAD5 * n5) + (_EMPTY3 * (rem // 3))
    raise ValueError(f"Cannot reach exact target byte size {target} from {len(compressed)}")


# ---------------------------------------------------------------------------
# スケルトン CFB リーダー -- vbaProject.bin から必要なストリームだけを
# 読み出す最小限の実装(書き込みは olefile に任せる。読み出し専用)。
# ---------------------------------------------------------------------------
class CFBReader:
    """CFB (OLE Compound File Binary) の読み取り専用パーサー。

    テンプレートスケルトンサイズ(FATが単一セクタに収まる程度)を前提とした
    簡易実装。vbaProject.bin の各ストリーム(dir / ThisWorkbook 等)の
    生バイト列とサイズ(entries)を取得するために使う。
    """

    def __init__(self, data: bytes):
        self.data = data
        assert data[:8] == b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1', "not a CFB file (bad magic)"
        self.minifat_start = struct.unpack('<I', data[60:64])[0]
        self.minifat_count = struct.unpack('<I', data[64:68])[0]
        self.fat_first = struct.unpack('<I', data[76:80])[0]
        self.dir_start = struct.unpack('<I', data[48:52])[0]
        self.num_fat_sectors = struct.unpack('<I', data[44:48])[0]
        self.difat_first = struct.unpack('<I', data[68:72])[0]

        # FAT。DIFAT(ヘッダ内109本 + DIFATセクタ連鎖)を辿って全FATセクタを繋ぐ。
        # 旧実装は「FATは単一セクタ」を前提にしていたが、配布方式B(裁定書27 W9-A)
        # の完成品binは1MBを超えFATが数十セクタになるため、ここを一般化した。
        difat = [struct.unpack('<I', data[76 + 4 * k:80 + 4 * k])[0]
                 for k in range(109)]
        sec = self.difat_first
        per = SECTOR_SIZE // 4
        while sec not in (ENDOFCHAIN, FREESECT) and (sec + 2) * SECTOR_SIZE <= len(data):
            off = (sec + 1) * SECTOR_SIZE
            block = list(struct.unpack('<%dI' % per, data[off:off + SECTOR_SIZE]))
            difat.extend(block[:-1])
            sec = block[-1]
        self.fat = []
        for fs in difat[:max(self.num_fat_sectors, 1)]:
            if fs in (ENDOFCHAIN, FREESECT) or (fs + 2) * SECTOR_SIZE > len(data):
                continue
            off = (fs + 1) * SECTOR_SIZE
            self.fat += list(struct.unpack('<%dI' % per,
                                           data[off:off + SECTOR_SIZE]))

        # Mini FAT。
        self.minifat = []
        sec = self.minifat_start
        while sec != ENDOFCHAIN and sec < len(self.fat):
            off = (sec + 1) * SECTOR_SIZE
            self.minifat += list(struct.unpack('<128I', data[off:off + SECTOR_SIZE]))
            sec = self.fat[sec]

        # ディレクトリエントリ。
        self.entries = {}  # name -> {'type','start','size'}
        sec = self.dir_start
        chain = []
        while sec != ENDOFCHAIN and sec < len(self.fat):
            chain.append(sec)
            sec = self.fat[sec]
        dir_bytes = b''.join(
            data[(s + 1) * SECTOR_SIZE:(s + 2) * SECTOR_SIZE] for s in chain)
        for i in range(0, len(dir_bytes), 128):
            e = dir_bytes[i:i + 128]
            name_len = struct.unpack('<H', e[64:66])[0]
            if name_len == 0:
                continue
            name = e[:name_len - 2].decode('utf-16-le')
            etype = e[66]
            start = struct.unpack('<I', e[116:120])[0]
            size = struct.unpack('<I', e[120:124])[0]
            self.entries[name] = {'type': etype, 'start': start, 'size': size}

        # ルートエントリはミニストリームのコンテナを保持する。
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
