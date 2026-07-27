#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""audit_ingest.py - 実行時の取込経路(modChunker)を実物PDFで検査する。

なぜ必要か
----------
シードは専用ツールで丁寧に切っているが、利用者が自分の資料を
「📁 資料を入れる」で取り込むときに走るのは modChunker(VBA)である。
そちらに同じ欠陥が残っていれば、利用者の資料だけ精度が落ちる。
しかも誰も気づけない。入力がゴミなら出力もゴミになる。

このスクリプトは modChunker.bas のロジック(ClassifyLine / MatchesDaiN /
IsNumberHeading / ChunkOnePageStructured / NormalizeWhitespace)を
Pythonへ忠実に移植し、実物のPDFに対して走らせて、
「条見出しをいくつ取りこぼすか」「条文が何本ページ境界で割れるか」を数える。

使い方:
    python3 tools/audit_ingest.py <PDF> [<PDF> ...]
"""
from __future__ import annotations

import re
import sys

try:
    import pypdf
except ImportError:
    sys.exit("pypdf が必要です: pip install pypdf")


# ------------------------------------------------------------------
# modChunker.bas の忠実な移植(2026-07-27時点)
# ------------------------------------------------------------------

def is_digit_char(c: str) -> bool:
    return c.isdigit() or ("０" <= c <= "９")


FW_ZERO, FW_NINE = 65296, 65305


def half_digit(c: str) -> str:
    o = ord(c)
    return chr(48 + (o - FW_ZERO)) if FW_ZERO <= o <= FW_NINE else c


def join_split_numbers(line: str) -> str:
    """VBA: JoinSplitNumbers。第〜条/章/節/項/号/編 の間の空白だけを除去。"""
    out, i, ln = [], 0, len(line)
    while i < ln:
        c = line[i]
        if c != "第":
            out.append(c); i += 1; continue
        j, digits = i + 1, ""
        while j < ln:
            d = line[j]
            if is_digit_char(d):
                digits += half_digit(d); j += 1
            elif d in " \u3000\t":
                j += 1
            else:
                break
        unit = line[j] if j < ln else ""
        if digits and unit in "条章節項号編":
            out.append("第" + digits + unit); i = j + 1
        else:
            out.append(c); i += 1
    return "".join(out)


def is_page_number_line(line: str) -> bool:
    t = line.strip().replace("\u3000", "").replace(" ", "")
    if not (3 <= len(t) <= 8):
        return False
    dashes = "-\u2010\u2011\u2012\u2013\u2014\u2015\u30fc\uff0d\u2212"
    if t[0] not in dashes or t[-1] not in dashes:
        return False
    mid = t[1:-1]
    return bool(mid) and all(is_digit_char(ch) for ch in mid)


def normalize_for_ingest(s: str) -> str:
    rows = s.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    return "\n".join(join_split_numbers(r) for r in rows if not is_page_number_line(r))


def matches_dai_n(t: str, kanji: str) -> bool:
    """VBA: MatchesDaiN。第の直後に数字が"連続"し、その直後が kanji であること。"""
    if not t or t[0] != "第":
        return False
    i = 1
    digit_n = 0
    while i < len(t) and is_digit_char(t[i]):
        digit_n += 1
        i += 1
    if digit_n == 0:
        return False
    return t[i:i + len(kanji)] == kanji


def is_number_heading(t: str) -> bool:
    if len(t) >= 50:
        return False
    if not t or not is_digit_char(t[0]):
        return False
    i = 0
    while i < len(t) and is_digit_char(t[i]):
        i += 1
    if i >= len(t):
        return False
    return t[i] in ".．"


def is_item_marker(t: str) -> bool:
    return bool(t) and t[0] in "・-−ー*●○▲△◇◆"


def classify_line(line: str) -> int:
    """0=本文 1=文書見出し 2=節見出し 3=箇条書き 4=表行"""
    t = line.strip()
    if not t:
        return 0
    if "\t" in line or "   " in line or "|" in line:
        return 4
    if line[:2] == "# ":
        return 1
    if t[0] == "【" and t[-1] == "】" and len(t) <= 60:
        return 1
    if matches_dai_n(t, "編") or matches_dai_n(t, "章"):
        return 1
    if matches_dai_n(t, "条") or matches_dai_n(t, "節"):
        return 2
    if is_number_heading(t):
        return 2
    if t[0] in "■●◆" and len(t) < 40 and "。" not in t:
        return 2
    if matches_dai_n(t, "項") or is_item_marker(t):
        return 3
    return 0


def build_blocks_cross_page(pages: list[str]) -> list[tuple[int, str]]:
    """VBA: ChunkAllPagesStructured。全ページを1つの行列として見出しで割る。

    ページ境界では切らない。ブロックのページ番号は先頭行が載っていたページ。
    """
    blocks: list[tuple[int, str]] = []
    block: list[str] = []
    block_page = 1
    for pno, text in enumerate(pages, start=1):
        norm = normalize_for_ingest(text)
        for row in norm.split("\n"):
            lbl = classify_line(row)
            if lbl in (1, 2):
                if block:
                    blocks.append((block_page, "\n".join(block)))
                block = [row.strip()]
                block_page = pno
            elif row.strip():
                if not block:
                    block_page = pno
                block.append(row.strip())
    if block:
        blocks.append((block_page, "\n".join(block)))
    return [b for b in blocks if b[1].strip()]


def count_headless_blocks(blocks: list[tuple[int, str]]) -> int:
    """見出しを持たないブロック = 何の話か分からないチャンクの数。

    これが多いほど「第12条は?」に対して見出しの無い断片が返る危険が高い。
    """
    n = 0
    for _, body in blocks:
        first = body.split("\n", 1)[0].strip()
        if classify_line(first) not in (1, 2):
            n += 1
    return n


# ------------------------------------------------------------------
# 検査
# ------------------------------------------------------------------

ARTICLE_ANY = re.compile(r"第\s*\d[\d\s]*\s*条")


def audit(path: str) -> dict:
    reader = pypdf.PdfReader(path)
    pages = [(p.extract_text() or "") for p in reader.pages]

    total_article_lines = 0     # 「第N条」で始まる行(空白入りも含む=人間が見て条見出し)
    recognized = 0              # modChunker が節見出し(2)として認識できた数
    missed_examples = []

    for text in pages:
        for line in text.replace("\r", "\n").split("\n"):
            t = line.strip()
            if not t:
                continue
            # 人間の目で「条見出し」と分かる行: 行頭が 第…条
            m = re.match(r"^第\s*[\d０-９][\d０-９\s]*\s*条", t)
            if not m:
                continue
            total_article_lines += 1
            if classify_line(join_split_numbers(t)) == 2:
                recognized += 1
            elif len(missed_examples) < 5:
                missed_examples.append(t[:44])

    # 修正後の実装(ページ横断)でブロックを組み、見出しの無いブロックを数える。
    # ページ境界で切らないので、条文が途中で割れることは構造的に起きない。
    blocks = build_blocks_cross_page(pages)
    headless = count_headless_blocks(blocks)

    return {
        "path": path,
        "pages": len(pages),
        "article_lines": total_article_lines,
        "recognized": recognized,
        "missed": total_article_lines - recognized,
        "missed_examples": missed_examples,
        "blocks": len(blocks),
        "headless": headless,
    }


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 2
    print("=" * 78)
    print("実行時の取込経路(modChunker structure モード)の実測")
    print("=" * 78)
    tot_a = tot_r = tot_s = 0
    for p in argv[1:]:
        r = audit(p)
        tot_a += r["article_lines"]
        tot_r += r["recognized"]
        tot_s += r["headless"]
        name = p.rsplit("/", 1)[-1][:44]
        print(f"\n{name}  ({r['pages']}頁)")
        if r["article_lines"]:
            rate = r["recognized"] / r["article_lines"] * 100
            print(f"  条見出し {r['article_lines']:4d} 行中 {r['recognized']:4d} 行を認識 "
                  f"({rate:5.1f}%)  取りこぼし {r['missed']}")
            for ex in r["missed_examples"]:
                print(f"      見落とし例: 「{ex}」")
        else:
            print("  条見出しなし(ガイドライン系)")
        print(f"  ブロック {r['blocks']} 件 / うち見出し無し {r['headless']} 件"
              f" ({r['headless']/max(r['blocks'],1)*100:.1f}%)")

    print("\n" + "-" * 78)
    if tot_a:
        print(f"合計: 条見出し {tot_a} 行中 {tot_r} 行認識 ({tot_r/tot_a*100:.1f}%) / "
              f"取りこぼし {tot_a - tot_r}")
    print(f"合計: 見出しを持たないブロック {tot_s} 件"
          "  ← 何の話か分からないチャンク。少ないほど回答が的を外さない")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
