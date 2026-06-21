#!/usr/bin/env python3
"""End-to-end test of the upgraded pipeline:
   - 3072-dim embeddings
   - gemini-2.5-pro for draft AND for self-verification
   - Side-by-side comparison: flash-only vs pro+verify"""
import os, sys, json, struct, math, requests, pathlib, time

API_KEY = os.environ["GEMINI_API_KEY"]
EMBED_MODEL = "gemini-embedding-001"
DRAFT_MODEL = "gemini-2.5-pro"
VERIFIER_MODEL = "gemini-2.5-pro"
DIM = 3072
INDEX_DIR = pathlib.Path("dist/index")

def load_index():
    mf = json.load(open(INDEX_DIR / "manifest.json"))
    chunks = json.load(open(INDEX_DIR / "chunks.json"))
    raw = open(INDEX_DIR / "embeddings.bin", "rb").read()
    n, d = mf["count"], mf["dim"]
    print(f"manifest: dim={d} count={n}  (bin size={len(raw):,})")
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

def chat(model, system_prompt, user_prompt, max_tokens=16384, thinking_budget=4096):
    t0 = time.time()
    r = requests.post(
        f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={API_KEY}",
        json={
            "system_instruction": {"parts": [{"text": system_prompt}]},
            "contents": [{"role": "user", "parts": [{"text": user_prompt}]}],
            "generationConfig": {
                "temperature": 0.15, "maxOutputTokens": max_tokens, "topP": 0.95,
                "thinkingConfig": {"thinkingBudget": thinking_budget},
            }
        }, timeout=240
    )
    r.raise_for_status()
    j = r.json()
    cand = j.get("candidates", [{}])[0]
    parts = cand.get("content", {}).get("parts", [])
    fin = cand.get("finishReason")
    return "".join(p.get("text", "") for p in parts), time.time() - t0, fin

DRAFT_SYS = """あなたは社内アンダーライターの業務支援AIです。提示された社内ナレッジ抜粋のみを根拠に回答してください。

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

VERIFY_SYS = """あなたは損害保険会社引受部門の品質管理者です。AIが書いたドラフト回答が、引用元ナレッジに本当に書かれている内容のみで構成されているか厳格に検証します。

■ 検証ルール
1. ドラフト回答の各主張(箇条書きの各項目、本文の各文)を1つずつ、引用元ナレッジの該当チャンクと照合する。
2. 引用元ナレッジに明示的に書かれていない主張(言外の推測、要約しすぎ、ナレッジ外の数値・条文番号・金額・割合)を発見したら、その箇所を全文削除する。
3. 条件分岐(『ただし』『〜の場合を除く』『〜に限り』)の取り違いを検出したら修正する。否定/限定(『支払わない』『対象外』)を反対の意味に取り違えていたら必ず修正する。
4. 引用元ナレッジに無い条文番号・金額・期間・割合は『記載なし』に置き換える。
5. 各主張の末尾に出典マーカー([#1]形式)を維持する。マーカーが本当にその主張を支える出典を指しているかも確認し、誤っていれば修正する。

■ 出力形式
・ 検証後の最終回答のみを出力する(検証プロセスの説明や前置きは書かない)。
・ ドラフトの構造(要点要約→結論→詳細→次のアクション)は維持する。
・ 削除や修正があった場合、回答末尾に『■ 検証で除外/修正した内容』節を設けて簡潔に列挙する。除外/修正が無ければこの節は省略。
・ ドラフトが全面的にナレッジに無い内容だった場合は『社内ナレッジに該当する記載がありません。アンダーライターへ確認してください』のみ返す。"""

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
    for rank, (s, i) in enumerate(top, 1):
        c = chunks[i]
        block = f"[#{rank} {c['source']} p.{c['page']}]\n{c['text']}\n\n"
        if used + len(block) > 9000:
            break
        context += block
        used += len(block)

    user_msg = f"## ナレッジ\n{context}\n## 質問\n{question}"
    print(f"\n[ドラフト回答 ({DRAFT_MODEL}, ~{used}chars context)]")
    draft, t1, fin1 = chat(DRAFT_MODEL, DRAFT_SYS, user_msg)
    print(draft)
    print(f"\n  draft latency: {t1:.1f}s  finishReason: {fin1}  draft len: {len(draft)}")

    verify_msg = f"## 引用元ナレッジ\n{context}\n## 元の質問\n{question}\n\n## ドラフト回答\n{draft}"
    print(f"\n[検証後の最終回答 ({VERIFIER_MODEL})]")
    final, t2, fin2 = chat(VERIFIER_MODEL, VERIFY_SYS, verify_msg)
    print(final)
    print(f"\n  verify latency: {t2:.1f}s  total: {t1+t2:.1f}s  finishReason: {fin2}  final len: {len(final)}")

if __name__ == "__main__":
    chunks, vecs, d = load_index()
    questions = [
        "傷害保険における身体の障害の範囲を教えてください。単なるショックは対象になりますか？",
        "地震保険の保険金支払い条件と損害認定基準を教えて",
        "契約者が告知義務違反をした場合、どのような対応となりますか？",
    ]
    for q in questions:
        answer(q, chunks, vecs)
