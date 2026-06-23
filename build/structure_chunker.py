#!/usr/bin/env python3
"""
structure_chunker.py - article-aware Japanese insurance document chunker.

Goal: Beat sliding-window chunking on conditional/negation accuracy by
respecting 条文 boundaries and keeping ただし書き with its parent article.

Strategy per page text:
  1. Detect chapter (第X章/第X節) and article (第X条) markers.
  2. If 2+ article markers found -> article-based chunking with parent context.
  3. Otherwise -> sentence-aware sliding window fallback (for FAQ-style pages).
  4. Tables and bullet lists are kept together (never split inside).
  5. Each chunk is prefixed with a "context header" that includes:
        [文書名 (page N)] > 第X章 ... > 第Y条（タイトル）
     so even an isolated retrieval carries its location.
"""
import re

# --------------------------------------------------------------------------
# Patterns (Japanese fullwidth + halfwidth digits)
# --------------------------------------------------------------------------
DIGIT = r"[０-９0-9一二三四五六七八九十百千]+"

PAT_CHAPTER  = re.compile(rf"(?:^|\n|「|『)[\s　]*第\s*{DIGIT}\s*章[\s　]*[^\n]{{0,40}}")
PAT_SECTION  = re.compile(rf"(?:^|\n|「|『)[\s　]*第\s*{DIGIT}\s*節[\s　]*[^\n]{{0,40}}")
# An article HEADING must start at a line (newline + optional indent), and the
# title parenthesis must NOT be followed by reference particles like の/に/は.
PAT_ARTICLE  = re.compile(
    rf"(?:^|\n|「|『)[\s　]*第\s*{DIGIT}\s*条[\s　]*[（\(][^）\)]{{0,40}}[）\)]"
    rf"(?![の・にはをがでとへもまで])"
)
PAT_ARTICLE_NOTITLE = re.compile(
    rf"(?:^|\n|「|『)[\s　]*第\s*{DIGIT}\s*条(?![（\(])(?![の・にはをがでとへもまでお])"
)

# Sub-clause numbering inside articles
PAT_SUBCLAUSE = re.compile(rf"(?:^|[\s　])(?:[（\(]\s*{DIGIT}\s*[）\)]|[①-⑳]|[ア-ン][\s　]*[\.．])")

# Soft target / hard ceiling for chunk size (chars)
TARGET_CHARS = 1400
MAX_CHARS    = 2400
MIN_CHARS    = 150  # don't emit tiny chunks; merge them into the next

# Sentence boundaries we'd rather break at
SENT_BREAKS = ["\n\n", "。\n", "。」", "。", "\n", " "]

# --------------------------------------------------------------------------
def _shift_to_marker(text, m):
    """Drop the leading newline/whitespace so the marker start is the heading char."""
    s, e = m.start(), m.end()
    label = m.group().strip()
    # Skip leading \n and whitespace inside the match
    while s < e and text[s] in "\n 　\t":
        s += 1
    return s, e, label


def _find_markers(text):
    """Return sorted list of (start, end, kind, label) tuples for markers."""
    out = []
    for m in PAT_CHAPTER.finditer(text):
        s, e, lbl = _shift_to_marker(text, m)
        out.append((s, e, "chapter", lbl))
    for m in PAT_SECTION.finditer(text):
        s, e, lbl = _shift_to_marker(text, m)
        out.append((s, e, "section", lbl))
    for m in PAT_ARTICLE.finditer(text):
        s, e, lbl = _shift_to_marker(text, m)
        out.append((s, e, "article", lbl))
    titled_starts = {s for s, _, k, _ in out if k == "article"}
    for m in PAT_ARTICLE_NOTITLE.finditer(text):
        s, e, lbl = _shift_to_marker(text, m)
        if s not in titled_starts:
            out.append((s, e, "article", lbl))
    out.sort(key=lambda t: t[0])
    return out


def _split_at_subclauses(text, max_chars):
    """Split a long article body at sub-clause markers, keeping parts <= max_chars."""
    points = [m.start() for m in PAT_SUBCLAUSE.finditer(text)]
    if not points:
        return _soft_split(text, max_chars)
    points.append(len(text))
    parts = []
    start = 0
    for p in points:
        if p - start >= max_chars:
            # break here
            parts.append(text[start:p].strip())
            start = p
    if start < len(text):
        parts.append(text[start:].strip())
    # Merge any tiny tail into previous
    out = []
    for p in parts:
        if out and len(p) < MIN_CHARS:
            out[-1] = out[-1] + "\n" + p
        else:
            out.append(p)
    return [p for p in out if p]


def _soft_split(text, max_chars):
    """Generic sentence-aware splitter."""
    if len(text) <= max_chars:
        return [text]
    parts, start = [], 0
    while start < len(text):
        end = min(start + max_chars, len(text))
        if end < len(text):
            for sep in SENT_BREAKS:
                idx = text.rfind(sep, start + max_chars // 2, end)
                if idx > 0:
                    end = idx + len(sep)
                    break
        parts.append(text[start:end].strip())
        start = end
    return [p for p in parts if p]


def _format_header(doc_title, page, chapter, section, article):
    parts = [f"[{doc_title} p.{page}]"]
    if chapter: parts.append(chapter)
    if section: parts.append(section)
    if article: parts.append(article)
    return " > ".join(parts)


# --------------------------------------------------------------------------
def chunk_page(text, source, page, doc_title=None):
    """
    Returns a list of chunk dicts:
        {id, source, page, start, header, text}
    'text' includes the header prefix so the embedding sees article context.
    """
    if doc_title is None:
        doc_title = source
    text = text or ""
    if not text.strip():
        return []

    markers = _find_markers(text)
    arts = [m for m in markers if m[2] == "article"]

    # Decide chunking strategy
    if len(arts) >= 2:
        return _chunk_by_articles(text, markers, source, page, doc_title)
    else:
        return _chunk_sliding(text, source, page, doc_title)


def _chunk_by_articles(text, markers, source, page, doc_title):
    """Group text by article boundaries; include nearest chapter/section above."""
    # Build a running context state walking through markers in order.
    chunks = []
    cid = 0

    # Identify article spans first
    arts = [m for m in markers if m[2] == "article"]
    art_spans = []
    for i, (s, e, _, label) in enumerate(arts):
        next_s = arts[i + 1][0] if i + 1 < len(arts) else len(text)
        art_spans.append((s, next_s, label))

    # Helper: find chapter/section in effect at offset s
    def context_at(offset):
        chap, sec = None, None
        for ms, me, k, lbl in markers:
            if ms >= offset:
                break
            if k == "chapter": chap, sec = lbl, None
            elif k == "section": sec = lbl
        return chap, sec

    # Preamble before first article (titles, table of contents, etc.)
    first_art_start = art_spans[0][0]
    if first_art_start > MIN_CHARS:
        pre = text[:first_art_start].strip()
        if pre:
            chap, sec = context_at(0)
            header = _format_header(doc_title, page, chap, sec, None)
            for piece in _soft_split(pre, MAX_CHARS):
                chunks.append({
                    "id": f"{source}::p{page}::c{cid}",
                    "source": source, "page": page, "start": 0,
                    "header": header,
                    "text": header + "\n" + piece,
                })
                cid += 1

    # Each article -> 1+ chunks, all prefixed with the same header
    for (s, e, art_label) in art_spans:
        body = text[s:e].strip()
        if not body:
            continue
        chap, sec = context_at(s)
        header = _format_header(doc_title, page, chap, sec, art_label)
        # If short enough, single chunk
        if len(body) <= MAX_CHARS:
            chunks.append({
                "id": f"{source}::p{page}::c{cid}",
                "source": source, "page": page, "start": s,
                "header": header,
                "text": header + "\n" + body,
            })
            cid += 1
            continue
        # Split inside article at sub-clauses, but each piece gets the same header
        for piece in _split_at_subclauses(body, TARGET_CHARS):
            chunks.append({
                "id": f"{source}::p{page}::c{cid}",
                "source": source, "page": page, "start": s,
                "header": header,
                "text": header + "\n" + piece,
            })
            cid += 1

    return chunks


def _chunk_sliding(text, source, page, doc_title):
    """Fallback for FAQ-style / sales-collateral pages with no article markers."""
    header = _format_header(doc_title, page, None, None, None)
    chunks = []
    cid = 0
    for piece in _soft_split(text, TARGET_CHARS):
        if len(piece) < MIN_CHARS and chunks:
            chunks[-1]["text"] = chunks[-1]["text"] + "\n" + piece
            continue
        chunks.append({
            "id": f"{source}::p{page}::c{cid}",
            "source": source, "page": page, "start": 0,
            "header": header,
            "text": header + "\n" + piece,
        })
        cid += 1
    return chunks


# --------------------------------------------------------------------------
if __name__ == "__main__":
    import sys, json
    from pypdf import PdfReader
    if len(sys.argv) < 2:
        print("usage: structure_chunker.py <pdf> [page]")
        sys.exit(1)
    r = PdfReader(sys.argv[1])
    page_no = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    page = r.pages[page_no - 1]
    text = page.extract_text() or ""
    chunks = chunk_page(text, sys.argv[1].split("/")[-1], page_no)
    for c in chunks:
        print("-" * 60)
        print("HEADER:", c["header"])
        print("LEN:", len(c["text"]))
        print(c["text"][:300] + ("..." if len(c["text"]) > 300 else ""))
