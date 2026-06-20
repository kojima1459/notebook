#!/usr/bin/env python3
"""Test the generated index: embed a question, find top-K, see what we'd send to Gemini."""
import os, sys, json, struct, math, requests, pathlib

API_KEY = os.environ["GEMINI_API_KEY"]
EMBED_MODEL = "gemini-embedding-001"
CHAT_MODEL = "gemini-2.5-flash"
DIM = 768
INDEX_DIR = pathlib.Path("dist/index")

def load_index():
    mf = json.load(open(INDEX_DIR / "manifest.json"))
    chunks = json.load(open(INDEX_DIR / "chunks.json"))
    raw = open(INDEX_DIR / "embeddings.bin", "rb").read()
    n, d = mf["count"], mf["dim"]
    vecs = []
    for i in range(n):
        v = struct.unpack(f"<{d}d", raw[i*d*8:(i+1)*d*8])
        vecs.append(v)
    return chunks, vecs, d

def embed_query(text):
    r = requests.post(
        f"https://generativelanguage.googleapis.com/v1beta/models/{EMBED_MODEL}:embedContent?key={API_KEY}",
        json={
            "model": f"models/{EMBED_MODEL}",
            "content": {"parts": [{"text": text}]},
            "taskType": "RETRIEVAL_QUERY",
            "outputDimensionality": DIM,
        }, timeout=60
    )
    r.raise_for_status()
    v = r.json()["embedding"]["values"]
    s = math.sqrt(sum(x*x for x in v))
    return [x/s for x in v]

def top_k(qv, vecs, k=8):
    scored = []
    for i, v in enumerate(vecs):
        s = sum(a*b for a, b in zip(qv, v))
        scored.append((s, i))
    scored.sort(reverse=True)
    return scored[:k]

def gemini_chat(system_prompt, user_prompt):
    r = requests.post(
        f"https://generativelanguage.googleapis.com/v1beta/models/{CHAT_MODEL}:generateContent?key={API_KEY}",
        json={
            "system_instruction": {"parts": [{"text": system_prompt}]},
            "contents": [{"role": "user", "parts": [{"text": user_prompt}]}],
            "generationConfig": {"temperature": 0.15, "maxOutputTokens": 3072, "topP": 0.95}
        }, timeout=120
    )
    r.raise_for_status()
    j = r.json()
    cand = j.get("candidates", [{}])[0]
    parts = cand.get("content", {}).get("parts", [])
    return "".join(p.get("text", "") for p in parts)

SYSTEM_PROMPT = """あなたは社内アンダーライターの業務支援AIです。提示された社内ナレッジ抜粋のみを根拠に回答してください。

■ 回答の組み立てかた
1. まず質問の要点を1行で言い換える(『ご質問は◯◯ですね』形式)。
2. 結論を最初に箇条書き(3項目以内)で述べる。各項目末に根拠の出典マーカーを付ける(例: [#1, #3])。
3. 続けて『詳細』として、ナレッジ本文の語彙をできる限り維持して説明する。条文番号や金額がある場合は必ず引用する。
4. ナレッジ間に矛盾がある場合は『記述差分』節を設けて両方の出典を併記する。
5. 最後に『次のアクション』として、本社アンダーライティング部に確認すべき事項があれば1〜2点示す。

■ 厳守ルール
・ ナレッジに無い情報は推測しない。該当が無ければ『社内ナレッジに該当する記載がありません。アンダーライターへ確認してください』とだけ返す。
・ 出典マーカー([#1]形式)を必ず本文中の該当箇所に挿入する。末尾にまとめて貼るだけは禁止。
・ 個人情報(契約番号・氏名・電話番号・マイナンバー等)が質問に含まれる場合は『個人情報を含めないでください』とだけ返答し、回答しない。
・ ナレッジに無い数値や条文番号を創作しない。数値はナレッジから抜粋し、無ければ『金額/期間の記載なし』と書く。
・ 営業担当者の視点で平易に説明する。専門用語には括弧で短い説明を添える。"""

def answer(question, chunks, vecs):
    print(f"\n{'='*70}\n質問: {question}\n{'='*70}")
    qv = embed_query(question)
    top = top_k(qv, vecs, k=8)

    print(f"\n[Top-8 検索結果]")
    for rank, (s, i) in enumerate(top, 1):
        c = chunks[i]
        head = c["text"][:60].replace("\n", " ")
        print(f"  #{rank} score={s:.4f}  {c['source']} p.{c['page']}  : {head}...")

    context = ""
    used = 0
    cite_blocks = []
    for rank, (s, i) in enumerate(top, 1):
        c = chunks[i]
        block = f"[#{rank} {c['source']} p.{c['page']}]\n{c['text']}\n\n"
        if used + len(block) > 9000:
            break
        context += block
        used += len(block)
        cite_blocks.append((rank, c['source'], c['page']))

    user_msg = f"## ナレッジ\n{context}\n## 質問\n{question}"
    print(f"\n[回答 ({CHAT_MODEL})]")
    print(gemini_chat(SYSTEM_PROMPT, user_msg))

if __name__ == "__main__":
    chunks, vecs, d = load_index()
    print(f"Loaded {len(chunks)} chunks, dim={d}")

    questions = [
        "傷害保険における身体の障害の範囲を教えてください。単なるショックは対象になりますか？",
        "地震保険の保険金支払い条件と損害認定基準を教えて",
        "契約者が告知義務違反をした場合、どのような対応となりますか？",
    ]
    for q in questions:
        answer(q, chunks, vecs)
