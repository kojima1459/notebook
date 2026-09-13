#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""pdf_layout.py - 座標を使って PDF の段組・表を壊さずに読む。

なぜ必要か
----------
PDFのテキスト抽出は、紙の上の「見た目の位置」を捨てて、内部の描画順で
文字を並べ直す。2列の表があると、左列(項目)と右列(説明)が行ごとに
交互に連結され、次のような読めない文字列になる:

    2. 制度運営体制 補償提供者（被保険者）が主体となって制度運営
    をすることを理解しているか。 この保険はあくまでも…

人間には「項目」と「説明」の対応が見えているのに、AIには1本の壊れた文に
見える。引受ガイドラインは大半がこの形式の表なので、ここが崩れると
その資料は丸ごと使い物にならない。

やり方
------
pypdf の visitor_text で1断片ごとに (x, y, テキスト) を取り、

  1. y でグルーピングして「行」を作る(印刷のわずかな上下ブレを吸収)
  2. 各行の断片の x 開始位置を全ページぶん集め、
     「どの行にも文字が来ない x の帯」= 列の谷 を探す
  3. 谷が見つかれば段組とみなし、列ごとに縦に読んでから
     「左 ｜ 右」の形で1行に組み直す
  4. 谷が無ければ通常の1段組として素直に連結する

谷の判定を「全行に共通して空いている」に限るのが肝。1行だけ偶然空いた
場所を列境界と誤認すると、普通の文章を勝手に切ってしまう。
"""
from __future__ import annotations

import sys
from collections import defaultdict

try:
    import pypdf
except ImportError:
    sys.exit("pypdf が必要です: pip install pypdf")


# 行のまとめ幅(pt)。日本語の本文は18pt行送りが多いので、その半分以下にする。
ROW_TOL = 4.0
# 列の谷とみなす最小の幅(pt)。全角1文字が約10ptなので、それより広いこと。
MIN_GAP = 24.0
# 谷が「本物の列境界」と言えるために必要な行数の割合。
GAP_ROW_RATIO = 0.55
# 段組として扱うのに必要な最小行数(数行しかない箇所を表と誤認しない)。
MIN_TABLE_ROWS = 4


def extract_fragments(page) -> list[tuple[float, float, float, str]]:
    """(y, x, 幅の目安, テキスト) の一覧を取り出す。"""
    out: list[tuple[float, float, float, str]] = []

    def visitor(text, cm, tm, font_dict, font_size):
        t = text.strip()
        if not t:
            return
        try:
            x = float(tm[4])
            y = float(tm[5])
        except (TypeError, ValueError, IndexError):
            return
        # 幅は概算(文字数×フォントサイズ)。全角は約1.0em、半角は約0.5em。
        try:
            size = float(font_size) or 10.0
        except (TypeError, ValueError):
            size = 10.0
        w = sum(1.0 if ord(c) > 0x2000 else 0.5 for c in t) * size
        out.append((y, x, w, t))

    page.extract_text(visitor_text=visitor)
    return out


def group_rows(frags) -> list[tuple[float, list[tuple[float, float, str]]]]:
    """y が近いものを1行にまとめる。戻り値は y の降順(紙の上から下)。"""
    buckets: dict[float, list[tuple[float, float, str]]] = defaultdict(list)
    for y, x, w, t in frags:
        key = round(y / ROW_TOL) * ROW_TOL
        buckets[key].append((x, w, t))
    rows = []
    for key in sorted(buckets, reverse=True):
        rows.append((key, sorted(buckets[key])))
    return rows


def find_column_split(rows) -> float | None:
    """全行に共通して文字が来ない x の帯(=列境界)を1つ探す。

    1行だけ偶然空いた場所を境界と誤認しないよう、
    「一定割合以上の行で空いている」ことを条件にする。
    """
    if len(rows) < MIN_TABLE_ROWS:
        return None

    xs = [x for _, cells in rows for x, _, _ in cells]
    if not xs:
        return None
    lo, hi = min(xs), max(x + w for _, cells in rows for x, w, _ in cells)
    if hi - lo < MIN_GAP * 3:
        return None

    # 5pt刻みで「その位置を文字が跨いでいる行の数」を数える
    step = 5.0
    positions = [lo + i * step for i in range(int((hi - lo) / step) + 1)]
    covered = [0] * len(positions)
    for _, cells in rows:
        for x, w, _ in cells:
            for i, px in enumerate(positions):
                if x <= px <= x + w:
                    covered[i] += 1

    n_rows = len(rows)
    free_needed = n_rows * (1.0 - GAP_ROW_RATIO)

    # 連続して「ほぼ誰も跨いでいない」帯を探し、最も広いものを採る
    best = None
    i = 0
    while i < len(positions):
        if covered[i] <= free_needed:
            j = i
            while j < len(positions) and covered[j] <= free_needed:
                j += 1
            width = (j - i) * step
            # 紙の端の余白は列境界ではない。中央寄りだけを見る
            center = positions[i] + width / 2
            if width >= MIN_GAP and lo + (hi - lo) * 0.15 < center < lo + (hi - lo) * 0.85:
                if best is None or width > best[1]:
                    best = (center, width)
            i = j
        else:
            i += 1
    return best[0] if best else None


def render_page(page) -> str:
    """1ページを、段組を保ったテキストへ変換する。"""
    frags = extract_fragments(page)
    if not frags:
        return ""
    rows = group_rows(frags)
    split = find_column_split(rows)

    lines: list[str] = []
    if split is None:
        for _, cells in rows:
            lines.append(" ".join(t for _, _, t in cells))
        return "\n".join(lines)

    # 段組あり: 行ごとに左右へ振り分け、「左 ｜ 右」で1行に組み直す。
    # こうすると「項目」と「説明」の対応がテキスト上でも保たれる。
    for _, cells in rows:
        left = [t for x, _, t in cells if x < split]
        right = [t for x, _, t in cells if x >= split]
        l = " ".join(left).strip()
        r = " ".join(right).strip()
        if l and r:
            lines.append(f"{l} ｜ {r}")
        elif l:
            lines.append(l)
        elif r:
            # 右列だけの行 = 前の項目の説明の続き。ぶら下げて対応を保つ
            lines.append(f"　｜ {r}")
    return "\n".join(lines)


def extract_document(path: str) -> list[tuple[int, str]]:
    """(ページ番号, 段組を保ったテキスト) の一覧。"""
    reader = pypdf.PdfReader(path)
    out = []
    for i, page in enumerate(reader.pages, start=1):
        try:
            out.append((i, render_page(page)))
        except Exception:
            # 1ページの失敗で資料全体を落とさない。素の抽出へ退避する。
            try:
                out.append((i, page.extract_text() or ""))
            except Exception:
                out.append((i, ""))
    return out


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 2
    for path in argv[1:]:
        print("=" * 74)
        print(path.rsplit("/", 1)[-1])
        tables = 0
        for pno, text in extract_document(path):
            if "｜" in text:
                tables += 1
        print(f"  段組(表)として復元したページ: {tables}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
