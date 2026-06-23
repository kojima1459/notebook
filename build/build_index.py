#!/usr/bin/env python3
"""
build_index.py - Build manifest.json + embeddings.bin + chunks.json
                 from PDFs in demo_data/ using Gemini text-embedding-004.

Output format must match what src/chatbot/modIndexReader.bas expects:
  manifest.json: {dim, count, embed_model, version}
  embeddings.bin: contiguous little-endian float64 (count * dim values, L2-normalized)
  chunks.json: [{id, source, page, start, text}, ...]
"""
import os
import sys
import json
import time
import struct
import math
import re
import pathlib
import requests
from concurrent.futures import ThreadPoolExecutor
import re as _re
from pypdf import PdfReader
try:
    import pdfplumber as _pdfplumber
    _HAS_PDFPLUMBER = True
except ImportError:
    _HAS_PDFPLUMBER = False
sys.path.insert(0, str(pathlib.Path(__file__).parent))
import structure_chunker as sc

# PDFs with custom font encoding that pypdf cannot decode; use pdfplumber instead
_PDFPLUMBER_SET = {"sample_02.pdf", "sample_05.pdf"}
_ILLEGAL_CHARS = _re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')

# ---- Settings -------------------------------------------------------------

API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
if not API_KEY:
    sys.exit("ERROR: set GEMINI_API_KEY env var before running")

EMBED_MODEL = "gemini-embedding-001"
EMBED_URL = (
    f"https://generativelanguage.googleapis.com/v1beta/"
    f"models/{EMBED_MODEL}:embedContent?key={API_KEY}"
)
EMBED_DIM = 3072  # max precision tier (gemini-embedding-001 returns L2-normalized at 3072)

# Chunking
CHUNK_CHARS = 1200
CHUNK_OVERLAP = 200

PDF_DIR = pathlib.Path("demo_data")
OUT_DIR = pathlib.Path("dist/index")
OUT_DIR.mkdir(parents=True, exist_ok=True)

# ---- PDF extraction -------------------------------------------------------

def clean_text(s: str) -> str:
    s = _ILLEGAL_CHARS.sub("", s)
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()

def extract_pdf_pages(path: pathlib.Path):
    if path.name in _PDFPLUMBER_SET and _HAS_PDFPLUMBER:
        return _extract_pdf_pages_pdfplumber(path)
    return _extract_pdf_pages_pypdf(path)

def _extract_pdf_pages_pypdf(path: pathlib.Path):
    reader = PdfReader(str(path))
    out = []
    for i, page in enumerate(reader.pages, start=1):
        try:
            t = page.extract_text() or ""
        except Exception as e:
            print(f"  page {i} extract failed: {e}", file=sys.stderr)
            t = ""
        t = clean_text(t)
        if t:
            out.append((i, t))
    return out

def _extract_pdf_pages_pdfplumber(path: pathlib.Path):
    out = []
    with _pdfplumber.open(str(path)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            try:
                t = page.extract_text() or ""
            except Exception as e:
                print(f"  page {i} extract failed: {e}", file=sys.stderr)
                t = ""
            t = clean_text(t)
            if t:
                out.append((i, t))
    return out

def chunk_text(text: str, source: str, page: int, start_id: int):
    """Structure-aware chunker (article-based for 約款, sliding for FAQ).
    Returns list of {id, source, page, start, text, header}."""
    return sc.chunk_page(text, source, page, doc_title=source)

# ---- Embedding ------------------------------------------------------------

def embed_one(text: str, task: str = "RETRIEVAL_DOCUMENT", retries: int = 4):
    body = {
        "model": f"models/{EMBED_MODEL}",
        "content": {"parts": [{"text": text}]},
        "taskType": task,
        "outputDimensionality": EMBED_DIM,
    }
    delay = 2.0
    for attempt in range(retries):
        try:
            r = requests.post(EMBED_URL, json=body, timeout=60)
            if r.status_code == 429 or r.status_code >= 500:
                raise RuntimeError(f"HTTP {r.status_code}: {r.text[:200]}")
            r.raise_for_status()
            data = r.json()
            return data["embedding"]["values"]
        except Exception as e:
            if attempt == retries - 1:
                raise
            print(f"  embed retry {attempt+1}: {e}", file=sys.stderr)
            time.sleep(delay)
            delay *= 2
    raise RuntimeError("unreachable")

def l2_normalize(v):
    s = math.sqrt(sum(x * x for x in v))
    if s == 0:
        return v
    return [x / s for x in v]

# ---- Main -----------------------------------------------------------------

def main():
    pdf_files = sorted(PDF_DIR.glob("*.pdf"))
    print(f"Found {len(pdf_files)} PDFs in {PDF_DIR}/")

    all_chunks = []
    for pdf in pdf_files:
        print(f"\n[{pdf.name}]")
        pages = extract_pdf_pages(pdf)
        print(f"  {len(pages)} pages with text")
        cid = 0
        for page_num, page_text in pages:
            ch = chunk_text(page_text, pdf.name, page_num, cid)
            cid += len(ch)
            all_chunks.extend(ch)
        print(f"  produced {sum(1 for c in all_chunks if c['source'] == pdf.name)} chunks")

    print(f"\nTotal chunks: {len(all_chunks)}")
    if not all_chunks:
        sys.exit("No chunks extracted; aborting")

    # ---- Embed all chunks ---- (parallel, paid tier)
    vectors = [None] * len(all_chunks)
    t0 = time.time()
    done = [0]

    def task(i_ch):
        i, ch = i_ch
        v = embed_one(ch["text"])
        if len(v) != EMBED_DIM:
            raise RuntimeError(f"dim {len(v)} != {EMBED_DIM} at chunk {i}")
        vectors[i] = l2_normalize(v)
        done[0] += 1
        d = done[0]
        if d % 25 == 0 or d == len(all_chunks):
            elapsed = time.time() - t0
            rate = d / max(elapsed, 0.1)
            eta = (len(all_chunks) - d) / max(rate, 0.1)
            print(f"  embedded {d}/{len(all_chunks)}  "
                  f"({rate:.1f}/s, ETA {eta:.0f}s)")

    with ThreadPoolExecutor(max_workers=12) as ex:
        for _ in ex.map(task, enumerate(all_chunks)):
            pass

    # ---- Write embeddings.bin (little-endian float64) ----
    bin_path = OUT_DIR / "embeddings.bin"
    with open(bin_path, "wb") as f:
        for v in vectors:
            f.write(struct.pack(f"<{EMBED_DIM}d", *v))
    print(f"\nWrote {bin_path}  ({bin_path.stat().st_size:,} bytes)")

    # ---- Write chunks.json ----
    chunks_path = OUT_DIR / "chunks.json"
    with open(chunks_path, "w", encoding="utf-8") as f:
        json.dump(all_chunks, f, ensure_ascii=False, indent=1)
    print(f"Wrote {chunks_path}  ({chunks_path.stat().st_size:,} bytes)")

    # ---- Write manifest.json ----
    manifest = {
        "dim": EMBED_DIM,
        "count": len(all_chunks),
        "embed_model": EMBED_MODEL,
        "version": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
    }
    mf_path = OUT_DIR / "manifest.json"
    with open(mf_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    print(f"Wrote {mf_path}")
    print(f"\nDONE. count={manifest['count']} dim={manifest['dim']}")

if __name__ == "__main__":
    main()
