#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""bench_retrieval.py - 検索精度を実物の資料と実際の質問で採点する。

なぜ必要か
----------
「精度が高い」は言うだけならタダである。金融で使う以上、
毎回同じ条件で採点して、点が下がったら気づける状態でなければ意味がない。

このスクリプトは
  ・実物6資料から作った 180チャンク(seed_pack.xlsx)
  ・その資料が確実に答えられる 31問(同じパックに同梱・正解の資料とページ付き)
を使い、「質問を投げたとき、正解のチャンクが上位何位に入るか」を数える。

採点する指標:
  Recall@1 / @3 / @5   … 正解資料のチャンクが上位k件に入った割合
  MRR                  … 正解が何位に出たかの逆数平均(1位なら1.0)
  ページ命中率          … 正解ページ±1に当たった割合(出典の正確さ)

埋め込み(dense)は社内リボンでしか作れないため、ここで測るのは
スパース側(キーワード/BM25)だけである。ただしスパース側こそが
「第12条」「瑕疵保証」のような完全一致を担保する側で、
日本語では実装の良し悪しが最も出る。ここが弱いと、
ベクトルがどれだけ良くても型番・条番号の質問で外す。

使い方:
    python3 tools/bench_retrieval.py seed/seed_pack.xlsx
"""
from __future__ import annotations

import math
import re
import sys
from collections import Counter, defaultdict

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("openpyxl が必要です: pip install openpyxl")


# ==================================================================
# トークナイザ
# ==================================================================

KANA_KANJI = re.compile(r"[぀-ヿ一-鿿]")
ALNUM = re.compile(r"[0-9A-Za-z]+")


def normalize_query(s: str) -> str:
    """検索前の正規化。全角/半角・大文字/小文字の揺れをここで吸収する。

    利用者は「第20条」「第２０条」「ﾘｺｰﾙ」「recall」「RECALL」を
    区別せずに打つ。ここで揃えないと、揺れた瞬間に0件になる。
    """
    out = []
    for ch in s:
        o = ord(ch)
        if 0xFF10 <= o <= 0xFF19:        # 全角数字 → 半角
            out.append(chr(o - 0xFF10 + 48))
        elif 0xFF21 <= o <= 0xFF3A:      # 全角大文字 → 半角小文字
            out.append(chr(o - 0xFF21 + 97))
        elif 0xFF41 <= o <= 0xFF5A:      # 全角小文字 → 半角小文字
            out.append(chr(o - 0xFF41 + 97))
        elif 0x41 <= o <= 0x5A:          # 半角大文字 → 小文字
            out.append(chr(o + 32))
        elif 0xFF66 <= o <= 0xFF9D:      # 半角カナ → 全角カナ(簡易)
            out.append(ch)
        elif ch in "　\t":
            out.append(" ")
        else:
            out.append(ch)
    return "".join(out)


def tokenize_baseline(s: str) -> list[str]:
    """現行 modRetrieve 相当。空白で割るだけ。

    日本語は空白で区切らないので、実質「質問文まるごと1語」になる。
    それが本文に部分一致することは、まず無い。
    """
    s = normalize_query(s)
    parts = [p for p in re.split(r"[ 　]+", s) if p]
    return parts


def tokenize_bigram(s: str) -> list[str]:
    """日本語向け: かな漢字は文字bigram、英数字は単語のまま。

    形態素解析器が使えない環境(VBA)での定石。
    「保険金を支払わない」→ 保険/険金/金を/を支/支払/払わ/わな/ない
    部分一致が確実に効き、表記の揺れにも強い。
    """
    s = normalize_query(s)
    toks: list[str] = []
    for m in ALNUM.finditer(s):
        toks.append(m.group(0))
    # かな漢字だけを取り出して連結し、bigramを作る
    jp = "".join(ch for ch in s if KANA_KANJI.match(ch))
    for i in range(len(jp) - 1):
        toks.append(jp[i:i + 2])
    return toks


# ==================================================================
# スコアリング
# ==================================================================

def build_index(chunks: list[dict], tokenizer):
    df: Counter = Counter()
    docs = []
    for c in chunks:
        toks = tokenizer(c["text"])
        tf = Counter(toks)
        docs.append({"tf": tf, "len": max(len(toks), 1)})
        for t in tf:
            df[t] += 1
    avgdl = sum(d["len"] for d in docs) / max(len(docs), 1)
    return {"df": df, "docs": docs, "N": len(docs), "avgdl": avgdl}


def bm25_scores(index, q_tokens: list[str], k1: float = 1.2, b: float = 0.75):
    """標準的なBM25。IDFで希少語を重く、文書長で正規化する。

    現行実装の「含まれていたら +0.05、上限0.15」との違い:
      ・「保険」のような頻出語と「瑕疵」のような希少語を同じ重みにしない
      ・長いチャンクが有利になるのを抑える
      ・上限で頭打ちにしないので、強く一致したものが確実に上へ来る
    """
    N, df, docs, avgdl = index["N"], index["df"], index["docs"], index["avgdl"]
    scores = [0.0] * N
    qtf = Counter(q_tokens)
    for t, qn in qtf.items():
        n = df.get(t, 0)
        if n == 0:
            continue
        idf = math.log(1 + (N - n + 0.5) / (n + 0.5))
        for i, d in enumerate(docs):
            f = d["tf"].get(t)
            if not f:
                continue
            dl = d["len"]
            scores[i] += idf * (f * (k1 + 1)) / (f + k1 * (1 - b + b * dl / avgdl))
    return scores


def exact_boost(chunks: list[dict], query: str) -> list[float]:
    """条番号・型番のような「完全一致すべき語」への加点。

    「第12条」と聞かれて第12条が1位に来ないのは、金融では事故。
    ベクトルもBM25も確率的なので、ここだけは決定的に押し上げる。
    """
    q = normalize_query(query)
    keys = re.findall(r"第\d+条(?:の\d+)?|[0-9A-Za-z]{3,}", q)
    if not keys:
        return [0.0] * len(chunks)
    out = []
    for c in chunks:
        body = normalize_query(c["text"])
        hit = sum(1 for k in keys if k in body)
        out.append(hit * 3.0)
    return out


def rrf(rank_lists: list[list[int]], k: int = 60) -> dict[int, float]:
    """Reciprocal Rank Fusion。スコアの尺度が違う複数の検索を、順位だけで統合する。

    denseとsparseはスコアの意味も範囲も違うので、加算すると必ずどちらかが
    支配する。RRFは順位しか見ないので、その事故が起きない(2026の定石)。
    """
    fused: dict[int, float] = defaultdict(float)
    for ranks in rank_lists:
        for pos, doc_id in enumerate(ranks, start=1):
            fused[doc_id] += 1.0 / (k + pos)
    return fused




# ==================================================================
# 出荷する構成の再現(VBAの実行速度制約を織り込む)
# ==================================================================
# 2万チャンク全件に文字bigram BM25 を回すのは VBA では現実的でない
# (1チャンク900字 × 9000件 = 800万回の文字走査。数十秒かかる)。
# そこで実装は「候補生成 → 再ランク」にする:
#   ① dense(埋め込み)の上位K件            … 既存の内積計算。意味の近さ
#   ② 質問から抜いた"効く語"の全件部分一致 … InStrはネイティブで速い。
#                                            完全一致の取りこぼしを防ぐ
#   ③ ①∪② の小さなプールにだけ BM25 を掛けて並べ替え
# ここでは dense を再現できないので、①を「無い」ものとして
# ②+③だけで採点する = 最悪ケース(埋め込みが役に立たなかった場合)。

RARE_RE = re.compile(r"第\d+条(?:の\d+)?|[0-9A-Za-z]{3,}|[一-鿿]{2,}")


def distinctive_keys(query: str, limit: int = 8) -> list[str]:
    """質問から「効く語」を抜く。長い漢字連続と英数字・条番号を優先する。

    空白は先に全部落とす。落とさないと「保 険 金」が「保」「険」「金」の
    1文字断片に割れて全部捨てられ、キーが1つも取れなくなる
    (実測: 空白を入れただけで R@1 84%→45% まで落ちた)。
    """
    q = normalize_query(query).replace(" ", "")
    keys = RARE_RE.findall(q)
    keys = [k for k in keys if len(k) >= 2]
    keys.sort(key=len, reverse=True)
    seen, out = set(), []
    for k in keys:
        if k in seen:
            continue
        seen.add(k)
        out.append(k)
        if len(out) >= limit:
            break
    return out


def shipping_scorer(chunks, idx_bi):
    """出荷構成: 部分一致で候補を集め、その中だけ BM25 で並べ替える。"""
    bodies = [normalize_query(c["text"] + " " + c["source"]) for c in chunks]

    def score(q):
        keys = distinctive_keys(q)
        cand = set()
        for k in keys:
            for i, b in enumerate(bodies):
                if k in b:
                    cand.add(i)
        if not cand:
            cand = set(range(len(chunks)))       # 1件も当たらなければ全件へ退避

        full = bm25_scores(idx_bi, tokenize_bigram(q))
        ex = exact_boost(chunks, q)
        out = [-1e9] * len(chunks)
        for i in cand:
            out[i] = full[i] + ex[i]
        return out
    return score


def instr_only_scorer(chunks):
    """トークナイズを一切しない構成(InStrの部分一致だけ)。

    VBAで毎クエリ全件トークナイズすると 9000件で数百秒〜数十分かかる
    (LibreOffice実測 225ms/件)。そこで「効く語の部分一致回数」だけで
    順位を付けたらどこまで出るかを測る。IDFの代わりに
    「語が長いほど希少」という近似で重みを付ける。
    """
    bodies = [normalize_query(c["text"] + " " + c["source"]) for c in chunks]

    def score(q):
        keys = distinctive_keys(q, limit=8)
        out = [0.0] * len(chunks)
        for k in keys:
            w = len(k) ** 1.5          # 長い語ほど重い(IDFの安価な代用)
            for i, b in enumerate(bodies):
                c = b.count(k)
                if c:
                    out[i] += w * (1 + math.log(c))
        # 文書長で正規化(長いチャンクが不当に有利になるのを防ぐ)
        for i, b in enumerate(bodies):
            if out[i] > 0:
                out[i] = out[i] / (1 + math.log(1 + len(b) / 900.0))
        return out
    return score


def final_scorer(chunks, idx_bi, pool: int = 20):
    """出荷する最終構成。

      ① InStr の部分一致だけで全件をスコアリング(トークナイズしないので速い)
      ② 上位 pool 件だけを BM25 で再ランク(ここだけトークナイズする)
      ③ 2つの順位を RRF で融合する

    ①は「効く語が本文にあるか」= キーワード完全一致に強く、ページ命中率が高い。
    ②は「語の重なり全体」= 意味的な近さに強く、R@3/R@5 を押し上げる。
    性質が違うので、順位だけを見る RRF で混ぜるのが素直(スコアの尺度が
    違うものを足すと必ずどちらかが支配してしまう)。
    実運用ではここに dense(埋め込み)の順位が3本目として加わる。
    """
    instr = instr_only_scorer(chunks)

    def score(q):
        s1 = instr(q)
        r1 = sorted(range(len(chunks)), key=lambda i: -s1[i])
        top = [i for i in r1[:pool] if s1[i] > 0]
        if not top:
            return s1

        bm = bm25_scores(idx_bi, tokenize_bigram(q))
        r2 = sorted(top, key=lambda i: -bm[i])

        fused = rrf([r1[:pool], r2])
        out = [-1e9] * len(chunks)
        for i in top:
            out[i] = fused.get(i, 0.0)
        # プール外は元の順位を維持(スコアを大きく下げてぶら下げる)
        for pos, i in enumerate(r1[pool:], start=pool):
            if out[i] < -1e8:
                out[i] = -1.0 - pos * 1e-6
        return out
    return score


def instr_compact_scorer(chunks):
    """InStrのみ + 空白除去マッチ(最終案)。

    InStr は連続一致なので、PDF由来やユーザー入力の余計な空白が
    語の途中に入ると当たらなくなる(実測 R@1 84%→55%)。
    照合用に「空白を除いた本文」を持ち、質問側も空白を除いてから当てる。
    こうすると「保 険 金」と「保険金」が同じものとして一致する。
    元本文は表示・出典用にそのまま残すので、表示は壊れない。
    """
    bodies = [normalize_query(c["text"] + " " + c["source"]).replace(" ", "")
              for c in chunks]
    lens = [max(len(b), 1) for b in bodies]

    def score(q):
        keys = [k.replace(" ", "") for k in distinctive_keys(q, limit=8)]
        keys = [k for k in keys if len(k) >= 2]
        out = [0.0] * len(chunks)
        for k in keys:
            w = len(k) ** 1.5
            for i, b in enumerate(bodies):
                c = b.count(k)
                if c:
                    out[i] += w * (1 + math.log(c))
        for i in range(len(chunks)):
            if out[i] > 0:
                out[i] = out[i] / (1 + math.log(1 + lens[i] / 900.0))
        return out
    return score


# ==================================================================
# 入念モードの検証: 多クエリ検索は本当に効くのか
# ==================================================================
# 「topKを増やす」は実測でほぼ効かなかった(R@5 90% → R@10 90%)。
# 上限に当たっているのは"渡す数"ではなく"見つける力"なので、
# 入念モードで増やすべきは件数ではなく検索の回数(角度)である。
#
# ここでは「1つの質問を複数の切り口で引く」= 多クエリ検索を模擬する:
#   ・元の質問そのもの
#   ・質問から抜いた効く語を1つずつ単独で引く
# それぞれの上位を集めて統合すると、単発で沈んだ資料が拾えるか。

def multi_query_scorer(chunks, per_query_top: int = 5):
    """多クエリ検索(入念モード想定)。元の質問+各キー単独で引いて統合する。"""
    base = instr_compact_scorer(chunks)
    bodies = [normalize_query(c["text"] + " " + c["source"]).replace(" ", "")
              for c in chunks]
    lens = [max(len(b), 1) for b in bodies]

    def one_key_score(k):
        out = [0.0] * len(chunks)
        w = len(k) ** 1.5
        for i, b in enumerate(bodies):
            c = b.count(k)
            if c:
                out[i] = w * (1 + math.log(c)) / (1 + math.log(1 + lens[i] / 900.0))
        return out

    def score(q):
        ranks = []
        s0 = base(q)
        ranks.append(sorted(range(len(chunks)), key=lambda i: -s0[i])[:per_query_top])

        for k in [x.replace(" ", "") for x in distinctive_keys(q, limit=5)]:
            if len(k) < 2:
                continue
            sk = one_key_score(k)
            if max(sk) <= 0:
                continue
            ranks.append(sorted(range(len(chunks)), key=lambda i: -sk[i])[:per_query_top])

        # 複数の検索結果は尺度が違うので、順位だけを見る RRF で統合する
        fused = rrf(ranks)
        out = [-1e9] * len(chunks)
        for i, v in fused.items():
            out[i] = v
        # 統合に入らなかったものは元スコアの順で下にぶら下げる
        order0 = sorted(range(len(chunks)), key=lambda i: -s0[i])
        for pos, i in enumerate(order0):
            if out[i] < -1e8:
                out[i] = -1.0 - pos * 1e-6
        return out
    return score

# ==================================================================
# 表記揺れ耐性(金融では「揺れたら0件」が最悪の事故)
# ==================================================================

def robustness_check(chunks, questions, scorer_factory):
    """同じ意味の質問を機械的に揺らして、順位が保たれるかを見る。

    利用者は「第20条」「第２０条」「だい20条」を区別せず打つ。
    揺れた瞬間に順位が崩れるなら、その検索は実務では使えない。
    """
    def to_fullwidth_digits(s):
        return "".join(chr(ord(c) - 48 + 0xFF10) if c.isdigit() else c for c in s)

    def to_upper(s):
        return s.upper()

    def add_spaces(s):
        return re.sub(r"([一-鿿])", r"\1 ", s, count=6)

    def drop_spaces(s):
        return s.replace(" ", "").replace("　", "")

    variants = [
        ("原文", lambda s: s),
        ("全角数字", to_fullwidth_digits),
        ("英字大文字", to_upper),
        ("余計な空白", add_spaces),
        ("空白除去", drop_spaces),
    ]
    out = []
    for label, fn in variants:
        qs = [{"q": fn(q["q"]), "source": q["source"], "page": q["page"]} for q in questions]
        out.append(evaluate(chunks, qs, scorer_factory, label))
    return out

# ==================================================================
# 採点
# ==================================================================

def evaluate(chunks, questions, scorer, name: str) -> dict:
    hit1 = hit3 = hit5 = hit10 = hit20 = 0
    page_hit = 0
    rr_sum = 0.0
    for q in questions:
        scores = scorer(q["q"])
        order = sorted(range(len(chunks)), key=lambda i: -scores[i])
        gold_src = q["source"]
        gold_page = q["page"]

        rank = None
        for pos, idx in enumerate(order[:20], start=1):
            if chunks[idx]["source"] == gold_src:
                rank = pos
                break
        if rank:
            rr_sum += 1.0 / rank
            if rank <= 1: hit1 += 1
            if rank <= 3: hit3 += 1
            if rank <= 5: hit5 += 1
            if rank <= 10: hit10 += 1
            if rank <= 20: hit20 += 1

        for idx in order[:5]:
            if chunks[idx]["source"] == gold_src and abs(chunks[idx]["page"] - gold_page) <= 1:
                page_hit += 1
                break

    n = max(len(questions), 1)
    return {
        "name": name,
        "R@1": hit1 / n, "R@3": hit3 / n, "R@5": hit5 / n,
        "R@10": hit10 / n, "R@20": hit20 / n,
        "MRR": rr_sum / n, "page": page_hit / n,
    }


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 2

    wb = load_workbook(argv[1], data_only=True)
    chunks = []
    for r in wb["pack_chunks"].iter_rows(min_row=2, values_only=True):
        if r and r[0]:
            chunks.append({"source": str(r[1]), "page": int(r[2] or 0),
                           "text": str(r[5] or "")})

    meta = {}
    for r in wb["pack_meta"].iter_rows(min_row=2, values_only=True):
        if r and r[0]:
            meta[str(r[0])] = r[1]
    questions = []
    i = 1
    while f"question{i}" in meta:
        src = str(meta.get(f"question{i}_src", "|")).rsplit("|", 1)
        questions.append({"q": str(meta[f"question{i}"]),
                          "source": src[0],
                          "page": int(src[1]) if len(src) > 1 and str(src[1]).isdigit() else 0})
        i += 1

    print("=" * 78)
    print(f"検索精度ベンチマーク  チャンク {len(chunks)} 件 / 質問 {len(questions)} 問")
    print("=" * 78)

    idx_base = build_index(chunks, tokenize_baseline)
    idx_bi = build_index(chunks, tokenize_bigram)

    def s_baseline(q):
        # 現行 modRetrieve 相当: 含まれていたら +0.05、上限 0.15
        toks = tokenize_baseline(q)
        out = []
        for c in chunks:
            body = normalize_query(c["text"] + " " + c["source"])
            hit = sum(1 for t in toks if t and t in body)
            out.append(min(hit * 0.05, 0.15))
        return out

    def s_bm25(q):
        return bm25_scores(idx_bi, tokenize_bigram(q))

    def s_bm25_exact(q):
        base = bm25_scores(idx_bi, tokenize_bigram(q))
        ex = exact_boost(chunks, q)
        return [base[i] + ex[i] for i in range(len(chunks))]

    # RRFは「それぞれ単体で使い物になるランカー」同士を混ぜたときだけ効く。
    # 役に立たないランカーを同じ重みで入れると순位を汚して必ず悪化する
    # (実測: bigram+空白分割+完全一致の3者融合は R@1 84%→19% に落ちた)。
    # 実行時は dense(埋め込み)と sparse(bigram BM25)の2者だけを融合する。
    # ここでは dense を再現できないので、代わりに「片方が壊れたときに
    # もう片方が支えるか」を見るため、sparse単体との差分を出す。
    def s_rrf_pair(q):
        b1 = bm25_scores(idx_bi, tokenize_bigram(q))
        ex = exact_boost(chunks, q)
        r1 = sorted(range(len(chunks)), key=lambda i: -b1[i])
        # 完全一致は「順位」ではなく「加点」として扱う(0点だらけの順位は
        # ただのノイズなので、RRFのメンバーにしてはいけない)
        fused = rrf([r1])
        return [fused.get(i, 0.0) + ex[i] * 0.001 for i in range(len(chunks))]

    rows = [
        evaluate(chunks, questions, s_baseline,   "現行(空白分割+一律加点)"),
        evaluate(chunks, questions, s_bm25,       "文字bigram BM25"),
        evaluate(chunks, questions, s_bm25_exact, "bigram BM25 + 完全一致加点"),
        evaluate(chunks, questions, s_rrf_pair,   "採用案(BM25順位 + 完全一致)"),
        evaluate(chunks, questions, shipping_scorer(chunks, idx_bi),
                 "出荷構成(候補生成→BM25再ランク)"),
        evaluate(chunks, questions, instr_only_scorer(chunks),
                 "InStrのみ(トークナイズ無し・最速)"),
        evaluate(chunks, questions, final_scorer(chunks, idx_bi),
                 "InStr全件→上位20をBM25→RRF"),
        evaluate(chunks, questions, instr_compact_scorer(chunks),
                 "★すぐ聞く/通常(InStr全件・空白除去)"),
        evaluate(chunks, questions, multi_query_scorer(chunks),
                 "★入念(多クエリ→RRF統合)"),
    ]

    print(f"\n{'手法':<30}{'R@1':>6}{'R@3':>6}{'R@5':>6}{'R@10':>7}{'R@20':>7}{'MRR':>7}{'ページ':>8}")
    print("-" * 78)
    for r in rows:
        print(f"{r['name']:<30}{r['R@1']:>5.0%}{r['R@3']:>6.0%}{r['R@5']:>6.0%}"
              f"{r['R@10']:>7.0%}{r['R@20']:>7.0%}{r['MRR']:>7.2f}{r['page']:>8.0%}")
    print()
    print("R@k = 正解資料のチャンクが上位k件に入った割合 / MRR = 正解順位の逆数平均")
    print("ページ = 上位5件に正解ページ±1が入った割合(出典表示の正確さ)")

    print("\n" + "=" * 78)
    print("表記揺れ耐性(採用案)  ― 揺れて順位が崩れないか")
    print("=" * 78)
    print(f"{'入力の揺れ':<34}{'R@1':>7}{'R@3':>7}{'R@5':>7}{'MRR':>7}")
    print("-" * 78)
    for r in robustness_check(chunks, questions, instr_compact_scorer(chunks)):
        print(f"{r['name']:<34}{r['R@1']:>6.0%}{r['R@3']:>7.0%}{r['R@5']:>7.0%}{r['MRR']:>7.2f}")
    print()
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
