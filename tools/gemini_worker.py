#!/usr/bin/env python3
"""gemini_worker.py - Gemini をワーカー(ジュニア)として呼ぶ司令塔専用の道具。

使い方:
  python3 tools/gemini_worker.py <prompt_file> <out_file> [--model gemini-3.1-pro-preview]
  (既定は 3.1 Pro=難易度1〜4の担当。速さ優先の雑務は --model gemini-3.8-flash)
                                 [--system system_file] [--temp 0.2] [--max-out 65536]

キーの探し方(順に。見つかった最初のものを使う。値は絶対に出力しない):
  1. 環境変数 GEMINI_API_KEY
  2. ファイル ~/.gemini_key
  3. 環境変数 CLAUDE_SCRATCHPAD/.gemini_key(セッションの作業用フォルダ)
  4. カレントの scratchpad/.gemini_key
リポジトリの中にキーを置かない(.gemini_key は .gitignore 済み)。

出力: out_file に本文をそのまま書く。標準出力には model/秒数/トークン数/finishReason だけ。
"""
from __future__ import annotations  # Python 3.9(会社Mac)でも `str | None` 注釈で落ちないため(姉妹PJ指摘 2026-09-06)

import json
import os
import sys
import time
import urllib.error
import urllib.request

API = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"


def find_key() -> str:
    k = os.environ.get("GEMINI_API_KEY", "").strip()
    if k:
        return k
    cands = [
        os.path.expanduser("~/.gemini_key"),
        os.path.join(os.environ.get("CLAUDE_SCRATCHPAD", ""), ".gemini_key"),
        os.path.join(os.getcwd(), "scratchpad", ".gemini_key"),
    ]
    for p in cands:
        if p and os.path.isfile(p):
            k = open(p, encoding="utf-8").read().strip()
            if k:
                return k
    raise SystemExit("GEMINI_API_KEY が無い(環境変数か ~/.gemini_key に置く。値は出力しない)")


def call(prompt: str, system: str | None, model: str, temperature: float, max_out: int):
    body = {
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": temperature, "maxOutputTokens": max_out},
    }
    if system:
        body["systemInstruction"] = {"parts": [{"text": system}]}
    req = urllib.request.Request(
        API.format(model=model),
        data=json.dumps(body).encode("utf-8"),
        headers={"Content-Type": "application/json", "x-goog-api-key": find_key()},
        method="POST",
    )
    t0 = time.time()
    try:
        with urllib.request.urlopen(req, timeout=900) as r:
            d = json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        msg = e.read().decode("utf-8", "replace")
        raise SystemExit(f"HTTP {e.code}: {msg[:400]}")
    sec = time.time() - t0
    try:
        text = "".join(p.get("text", "") for p in d["candidates"][0]["content"]["parts"])
    except Exception:
        raise SystemExit("unexpected response: " + json.dumps(d, ensure_ascii=False)[:400])
    usage = d.get("usageMetadata", {})
    finish = d["candidates"][0].get("finishReason", "")
    return text, sec, usage, finish


def main() -> None:
    a = sys.argv[1:]
    if len(a) < 2:
        print(__doc__)
        sys.exit(2)
    prompt = open(a[0], encoding="utf-8").read()
    out = a[1]
    model = a[a.index("--model") + 1] if "--model" in a else "gemini-3.1-pro-preview"
    system = open(a[a.index("--system") + 1], encoding="utf-8").read() if "--system" in a else None
    temperature = float(a[a.index("--temp") + 1]) if "--temp" in a else 0.2
    max_out = int(a[a.index("--max-out") + 1]) if "--max-out" in a else 65536
    text, sec, usage, finish = call(prompt, system, model, temperature, max_out)
    with open(out, "w", encoding="utf-8") as f:
        f.write(text)
    print(
        f"model={model} sec={sec:.1f} in={usage.get('promptTokenCount')} "
        f"out={usage.get('candidatesTokenCount')} finish={finish} chars={len(text)} -> {out}"
    )


if __name__ == "__main__":
    main()
