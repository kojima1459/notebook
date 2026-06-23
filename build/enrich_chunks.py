#!/usr/bin/env python3
"""
enrich_chunks.py - Annotate each chunk with:
    - 1-line summary (<= 60 chars)
    - keywords (5-10 comma-separated terms)
    - 業務領域 tag (per-PDF, from manifest below)
    - 書類種別 tag (per-PDF, from manifest below)

Skips corrupted/duplicate PDFs (sample_02, sample_05, sample_09).

Input:  dist/index/chunks.json   (785 chunks)
Output: dist/index/chunks_enriched.json (filtered + enriched)
"""
import os, sys, json, re, time, math, pathlib, requests
from concurrent.futures import ThreadPoolExecutor, as_completed

_ILLEGAL_CHARS = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')

API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
if not API_KEY:
    sys.exit("ERROR: set GEMINI_API_KEY env var")

MODEL = "gemini-2.5-flash"  # cheap fast model for batch enrichment
URL = f"https://generativelanguage.googleapis.com/v1beta/models/{MODEL}:generateContent?key={API_KEY}"

# Per-PDF metadata (manual, after sampling each PDF)
# domain: 業務領域  /  doc_type: 書類種別 (普通保険約款 / 引受ガイドライン / FAQ / 解説 / 研修資料 / ハンドブック)
PDF_MANIFEST = {
    # 約款 (sample_02 / sample_05 extracted via pdfplumber — custom font encoding)
    "sample_02.pdf":  {"display": "費用・利益保険 普通保険約款 (令3.10.1)",        "domain": "費用利益保険全般",       "doc_type": "普通保険約款"},
    "sample_03.pdf":  {"display": "瑕疵保証責任保険 普通保険約款 (令3.10.1)",   "domain": "瑕疵保証責任保険",       "doc_type": "普通保険約款"},
    "sample_05.pdf":  {"display": "約定履行費用保険 普通保険約款 (令3.10.1)",    "domain": "費用利益保険全般",       "doc_type": "普通保険約款"},
    "sample_11.pdf":  {"display": "瑕疵保証責任保険 普通保険約款 (関連版)",     "domain": "瑕疵保証責任保険",       "doc_type": "普通保険約款"},
    "sample_14.pdf":  {"display": "生産物回収費用保険 普通保険約款 (令3.10.1)", "domain": "生産物回収費用保険",     "doc_type": "普通保険約款"},
    "sample_15.pdf":  {"display": "家主費用・利益保険 (保種コード104)",        "domain": "家主費用・利益保険",     "doc_type": "普通保険約款"},
    # 引受ガイドライン
    "sample_06.pdf":  {"display": "顧客サービス費用保険(債務履行型) 引受ガイドライン",
                                                                         "domain": "顧客サービス費用保険",   "doc_type": "引受ガイドライン"},
    "sample_16.pdf":  {"display": "瑕疵保証責任保険 引受ガイドライン",         "domain": "瑕疵保証責任保険",       "doc_type": "引受ガイドライン"},
    # ハンドブック/解説
    "sample_12.pdf":  {"display": "リコールプロテクション(生産物回収費用保険) ハンドブック",
                                                                         "domain": "生産物回収費用保険",     "doc_type": "ハンドブック"},
    "sample_10.pdf":  {"display": "保証(補償)制度と保険について(いろはの「い」)",
                                                                         "domain": "費用利益保険全般",       "doc_type": "解説"},
    # FAQ
    "sample_01.pdf":  {"display": "普通約款 FAQ集 (身体の障害ほか)",        "domain": "費用利益保険全般",       "doc_type": "FAQ"},
    # 研修資料
    "sample_04.pdf":  {"display": "再保険基礎 (商品カレッジ基礎講座)",         "domain": "費用利益保険全般",       "doc_type": "研修資料"},
    "sample_07.pdf":  {"display": "収益改善の基礎 (商品カレッジ基礎講座)",     "domain": "費用利益保険全般",       "doc_type": "研修資料"},
    "sample_08.pdf":  {"display": "損害保険の基礎知識② (商品カレッジ)",      "domain": "費用利益保険全般",       "doc_type": "研修資料"},
    "sample_13.pdf":  {"display": "損害保険の基礎知識① (商品カレッジ)",      "domain": "費用利益保険全般",       "doc_type": "研修資料"},
}
# Skip list: sample_09 is an exact duplicate of sample_01
SKIP = {"sample_09.pdf"}

PROMPT_TMPL = """以下は損害保険会社の社内マニュアル抜粋です。LLMルーター用のメタデータを作成してください。

【抜粋】
{text}

【出力形式 (JSON厳守、他の文章は書かない)】
{{
  "summary": "1行60文字以内の要約。条文番号や項目があれば明記。例『第8条第1項：告知義務違反時の解除権、1ヶ月の期限あり』",
  "keywords": "カンマ区切り5-10語。検索用キーワード。例『告知義務,違反,解除,1ヶ月,失効,責任,通知,保険会社,保険契約者』"
}}"""

def enrich_one(text: str, retries: int = 3):
    body = {
        "contents": [{"role": "user", "parts": [{"text": PROMPT_TMPL.format(text=text[:3000])}]}],
        "generationConfig": {
            "temperature": 0.1,
            "responseMimeType": "application/json",
            "maxOutputTokens": 512,
            "thinkingConfig": {"thinkingBudget": 0},  # disable thinking for speed
        }
    }
    delay = 2.0
    for attempt in range(retries):
        try:
            r = requests.post(URL, json=body, timeout=60)
            if r.status_code == 429 or r.status_code >= 500:
                raise RuntimeError(f"HTTP {r.status_code}: {r.text[:200]}")
            r.raise_for_status()
            j = r.json()
            cand = j.get("candidates", [{}])[0]
            parts = cand.get("content", {}).get("parts", [])
            text_out = "".join(p.get("text", "") for p in parts)
            data = json.loads(text_out)
            return data.get("summary", "")[:80], data.get("keywords", "")[:300]
        except Exception as e:
            if attempt == retries - 1:
                print(f"  FAILED: {e}", file=sys.stderr)
                return "(自動要約失敗)", ""
            time.sleep(delay)
            delay *= 2
    return "(自動要約失敗)", ""

def main():
    in_path = pathlib.Path("dist/index/chunks.json")
    out_path = pathlib.Path("dist/index/chunks_enriched.json")
    chunks = json.load(open(in_path))
    print(f"Loaded {len(chunks)} chunks from {in_path}")

    # Filter: skip corrupted/duplicate PDFs
    keep = [c for c in chunks if c["source"] not in SKIP]
    print(f"After skip ({SKIP}): {len(keep)} chunks remain")

    # Annotate each with domain/doc_type
    for c in keep:
        meta = PDF_MANIFEST.get(c["source"])
        if meta:
            c["display"] = meta["display"]
            c["domain"] = meta["domain"]
            c["doc_type"] = meta["doc_type"]
        else:
            c["display"] = c["source"]
            c["domain"] = "未分類"
            c["doc_type"] = "未分類"

    # Strip illegal characters (e.g. \x08 from pdfplumber on some PDFs)
    for c in keep:
        for field in ("text", "summary", "keywords", "header"):
            if field in c and isinstance(c[field], str):
                c[field] = _ILLEGAL_CHARS.sub("", c[field])

    # Enrich with summary + keywords (parallel)
    print(f"\nEnriching {len(keep)} chunks with summary + keywords...")
    t0 = time.time()
    done = [0]

    def task(i_ch):
        i, ch = i_ch
        summary, keywords = enrich_one(ch["text"])
        ch["summary"] = summary
        ch["keywords"] = keywords
        done[0] += 1
        if done[0] % 25 == 0 or done[0] == len(keep):
            elapsed = time.time() - t0
            rate = done[0] / max(elapsed, 0.1)
            eta = (len(keep) - done[0]) / max(rate, 0.1)
            print(f"  {done[0]}/{len(keep)} ({rate:.1f}/s, ETA {eta:.0f}s)")

    with ThreadPoolExecutor(max_workers=10) as ex:
        list(ex.map(task, enumerate(keep)))

    # Save
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(keep, f, ensure_ascii=False, indent=1)
    print(f"\nWrote {out_path} ({out_path.stat().st_size:,} bytes)")
    print(f"DONE in {time.time()-t0:.0f}s")

    # Stats
    by_src = {}
    for c in keep:
        by_src.setdefault(c["source"], []).append(c)
    print("\n--- Per-PDF stats ---")
    for src in sorted(by_src):
        n = len(by_src[src])
        avg_summary_len = sum(len(c.get("summary","")) for c in by_src[src]) / n
        print(f"  {src}: {n} chunks, avg summary={avg_summary_len:.0f} chars  [{by_src[src][0].get('domain','-')}/{by_src[src][0].get('doc_type','-')}]")

if __name__ == "__main__":
    main()
