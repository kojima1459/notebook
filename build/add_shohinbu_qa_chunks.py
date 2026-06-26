#!/usr/bin/env python3
"""
add_shohinbu_qa_chunks.py - Ingest the product department's official Q&A
(営業からの照会に対する商品部の公式回答まとめ) into the knowledge base.

Source: build/source_data/shohinbu_qa.xlsx
  Sheet columns:
    A: 詳細カテゴリ1-1   (大分類: 新種共通 / 瑕疵・費用 ...)
    B: 詳細カテゴリ3-1   (小分類: 補償内容(補償可否) / 保険料計算 / その他 ...)
    C: ナレッジ名（質問） (営業からの照会)
    D: 回答（新種）       (商品部の回答; 改行は <br>)

  Record boundary: a NEW record begins at every row whose column A is
  non-empty (a category). All following rows whose column A is empty are
  CONTINUATIONS of that record — their C cells extend the question and their
  D cells extend the answer (multi-line Q&A spread over several rows).

Output: appends chunks with doc_type="商品部公式Q&A" to dist/index/chunks_enriched.json.

Why a distinct doc_type:
  These are AUTHORITATIVE precedents (商品部の公式回答). The pipeline is tuned
  to (1) retrieve a topically-matching official Q&A FIRST, then (2) corroborate
  it against the 約款 / 引受ガイドライン before answering. Keeping them as a
  separate doc_type lets the router prioritise them and lets the drafter cite
  them with their own [#N] marker — while still grounding the final 引受可否 /
  補償可否 / 支払可否 conclusion in the primary 約款 / ガイドライン evidence.

Re-running is idempotent: prior shohinbu_qa chunks are removed first.
"""
import json
import os
import re
import sys

import openpyxl

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT       = os.path.dirname(SCRIPT_DIR)
SRC_XLSX   = os.path.join(SCRIPT_DIR, "source_data", "shohinbu_qa.xlsx")
ENRICHED   = os.path.join(ROOT, "dist", "index", "chunks_enriched.json")

SOURCE_NAME = "shohinbu_qa.xlsx"
DOC_TYPE    = "商品部公式Q&A"

_ILLEGAL = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')
# Product-name detector: longest "…保険" run (kanji / katakana / middle dot).
_PRODUCT = re.compile(r'[一-龥ァ-ヶー・]{2,}?保険')
# Secondary domain keywords when no "…保険" appears in the question.
_TOPIC   = re.compile(r'[一-龥ァ-ヶー]{3,}(?:費用|特約|割引|プラン|補償|申請|区分)')


def _clean(s):
    if s is None:
        return ""
    s = str(s).replace("<br>", "\n").replace("\r\n", "\n").replace("\r", "\n")
    s = _ILLEGAL.sub("", s)
    s = re.sub(r"[ \t　]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def parse_records(path):
    """Return list of dicts: {cat, sub, q, a}."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    records, cur = [], None
    for row in ws.iter_rows(min_row=2, values_only=True):
        A = _clean(row[0]) if len(row) > 0 else ""
        B = _clean(row[1]) if len(row) > 1 else ""
        C = _clean(row[2]) if len(row) > 2 else ""
        D = _clean(row[3]) if len(row) > 3 else ""
        if A:                      # new record starts here
            if cur:
                records.append(cur)
            cur = {"cat": A, "sub": B if B and B != "(未設定)" else "", "q": [], "a": []}
            if C:
                cur["q"].append(C)
            if D:
                cur["a"].append(D)
        else:
            if cur is None:        # stray continuation before first category
                continue
            if not cur["sub"] and B and B != "(未設定)":
                cur["sub"] = B
            if C:
                cur["q"].append(C)
            if D:
                cur["a"].append(D)
    if cur:
        records.append(cur)

    out = []
    for r in records:
        q = "\n".join(r["q"]).strip()
        a = "\n".join(r["a"]).strip()
        if not q and not a:
            continue
        out.append({"cat": r["cat"], "sub": r["sub"], "q": q, "a": a})
    return out


def _derive_domain(cat, q):
    m = _PRODUCT.findall(q)
    if m:
        # pick the longest distinct product name (most specific)
        return max(m, key=len)
    return f"新種保険（{cat}）"


def _derive_keywords(cat, sub, q, a):
    kws = []
    for term in _PRODUCT.findall(q) + _PRODUCT.findall(a):
        if term not in kws:
            kws.append(term)
    for term in _TOPIC.findall(q):
        if term not in kws and len(kws) < 8:
            kws.append(term)
    if sub:
        kws.append(sub)
    kws.append(cat)
    kws.append("商品部回答")
    # de-dup preserving order, cap length
    seen, final = set(), []
    for k in kws:
        if k and k not in seen:
            seen.add(k)
            final.append(k)
    return ",".join(final[:12])


def _one_line(q):
    line = q.replace("\n", " ").strip()
    return line[:78] + ("…" if len(line) > 78 else "")


def build_chunks(records):
    chunks = []
    for i, r in enumerate(records, start=1):
        cat, sub, q, a = r["cat"], r["sub"], r["q"], r["a"]
        cid = f"{SOURCE_NAME}::qa{i:03d}"
        label = sub or cat
        header = f"[商品部公式Q&A] {cat}" + (f" > {sub}" if sub else "")
        text = (
            "【商品部公式Q&A（回答事例）】\n"
            f"分類: {cat}" + (f" / {sub}" if sub else "") + "\n\n"
            "■ 営業からの照会\n" + q + "\n\n"
            "■ 商品部の回答\n" + a
        )
        chunks.append({
            "id": cid,
            "source": SOURCE_NAME,
            "page": i,
            "start": 0,
            "header": header,
            "text": text,
            "display": f"商品部公式Q&A：{label}",
            "domain": _derive_domain(cat, q),
            "doc_type": DOC_TYPE,
            "summary": f"[商品部Q&A] {_one_line(q)}",
            "keywords": _derive_keywords(cat, sub, q, a),
        })
    return chunks


def main():
    if not os.path.exists(SRC_XLSX):
        sys.exit(f"Missing source: {SRC_XLSX}")
    if not os.path.exists(ENRICHED):
        sys.exit(f"Missing {ENRICHED} — run the base build first")

    records = parse_records(SRC_XLSX)
    new_chunks = build_chunks(records)

    existing = json.load(open(ENRICHED, encoding="utf-8"))
    before = len(existing)
    existing = [c for c in existing if c.get("source") != SOURCE_NAME]
    removed = before - len(existing)
    existing.extend(new_chunks)

    with open(ENRICHED, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)

    print(f"Parsed {len(records)} official Q&A records.")
    print(f"Removed {removed} prior {SOURCE_NAME} chunks, added {len(new_chunks)}.")
    print(f"Total chunks now: {len(existing)}")
    from collections import Counter
    print("Categories:", dict(Counter(r["cat"] for r in records)))


if __name__ == "__main__":
    main()
