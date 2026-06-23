#!/usr/bin/env python3
"""
extract_and_enrich_missing.py
Re-extract sample_02.pdf and sample_05.pdf using pdfplumber (pypdf fails on
their custom font encoding), enrich with Gemini summaries/keywords, and merge
the resulting chunks into chunks_enriched.json.

Run with:
  GEMINI_API_KEY=xxx python3 build/extract_and_enrich_missing.py
"""
import os, sys, json, time, pathlib, re, requests
from concurrent.futures import ThreadPoolExecutor
import pdfplumber
sys.path.insert(0, str(pathlib.Path(__file__).parent))
import structure_chunker as sc

API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
if not API_KEY:
    sys.exit("ERROR: set GEMINI_API_KEY env var before running")

MODEL = "gemini-2.5-flash"
URL = f"https://generativelanguage.googleapis.com/v1beta/models/{MODEL}:generateContent?key={API_KEY}"

PDF_DIR = pathlib.Path("demo_data")
ENRICHED_PATH = pathlib.Path("dist/index/chunks_enriched.json")

PDF_MANIFEST = {
    "sample_02.pdf": {
        "display": "費用・利益保険 普通保険約款 (令3.10.1)",
        "domain": "費用利益保険全般",
        "doc_type": "普通保険約款",
    },
    "sample_05.pdf": {
        "display": "約定履行費用保険 普通保険約款 (令3.10.1)",
        "domain": "費用利益保険全般",
        "doc_type": "普通保険約款",
    },
}

PROMPT_TMPL = """以下は損害保険会社の社内マニュアル抜粋です。LLMルーター用のメタデータを作成してください。

【抜粋】
{text}

【出力形式 (JSON厳守、他の文章は書かない)】
{{
  "summary": "1行60文字以内の要約。条文番号や項目があれば明記。例『第8条第1項：告知義務違反時の解除権、1ヶ月の期限あり』",
  "keywords": "カンマ区切り5-10語。検索用キーワード。例『告知義務,違反,解除,1ヶ月,失効,責任,通知,保険会社,保険契約者』"
}}"""


def clean_text(s: str) -> str:
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def extract_pdf(pdf_path: pathlib.Path):
    chunks = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            text = clean_text(text)
            if not text:
                continue
            page_chunks = sc.chunk_page(text, pdf_path.name, i, doc_title=pdf_path.name)
            chunks.extend(page_chunks)
    return chunks


def enrich_one(text: str, retries: int = 3):
    body = {
        "contents": [{"role": "user", "parts": [{"text": PROMPT_TMPL.format(text=text[:3000])}]}],
        "generationConfig": {
            "temperature": 0.1,
            "responseMimeType": "application/json",
            "maxOutputTokens": 512,
            "thinkingConfig": {"thinkingBudget": 0},
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
            parts = j.get("candidates", [{}])[0].get("content", {}).get("parts", [])
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
    existing = json.load(open(ENRICHED_PATH)) if ENRICHED_PATH.exists() else []
    existing_ids = {c["id"] for c in existing}
    print(f"Existing enriched chunks: {len(existing)}")

    new_chunks = []
    for pdf_name, meta in PDF_MANIFEST.items():
        pdf_path = PDF_DIR / pdf_name
        if not pdf_path.exists():
            print(f"WARNING: {pdf_path} not found, skipping")
            continue
        print(f"\n[{pdf_name}] extracting with pdfplumber...")
        chunks = extract_pdf(pdf_path)
        print(f"  {len(chunks)} chunks extracted")
        for c in chunks:
            c["display"] = meta["display"]
            c["domain"] = meta["domain"]
            c["doc_type"] = meta["doc_type"]
            c["dept_scope"] = "common"
        new_chunks.extend(chunks)

    print(f"\nEnriching {len(new_chunks)} new chunks...")
    t0 = time.time()
    done = [0]

    def task(i_ch):
        i, ch = i_ch
        summary, keywords = enrich_one(ch["text"])
        ch["summary"] = summary
        ch["keywords"] = keywords
        done[0] += 1
        if done[0] % 20 == 0 or done[0] == len(new_chunks):
            elapsed = time.time() - t0
            rate = done[0] / max(elapsed, 0.1)
            eta = (len(new_chunks) - done[0]) / max(rate, 0.1)
            print(f"  {done[0]}/{len(new_chunks)} ({rate:.1f}/s, ETA {eta:.0f}s)")

    with ThreadPoolExecutor(max_workers=10) as ex:
        list(ex.map(task, enumerate(new_chunks)))

    combined = existing + new_chunks
    with open(ENRICHED_PATH, "w", encoding="utf-8") as f:
        json.dump(combined, f, ensure_ascii=False, indent=1)
    print(f"\nWrote {ENRICHED_PATH}: {len(combined)} total chunks ({len(new_chunks)} added)")
    print(f"Done in {time.time()-t0:.0f}s")

    by_src = {}
    for c in new_chunks:
        by_src.setdefault(c["source"], 0)
        by_src[c["source"]] += 1
    for src, n in sorted(by_src.items()):
        print(f"  {src}: {n} chunks")


if __name__ == "__main__":
    main()
