#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""doc_gate.py - 文書検問(2026-09-12 R47 新設)

なぜ要るのか
------------------------------------------------------------------------------
このプロジェクトの検問(vba_lint / run_lo_tests / build の自己検証 /
check_launcher / bin_roundtrip)は **コードだけを守っていて、文書を1文字も
検査していなかった**。だから文書は書いた瞬間から劣化し、誰も気づかず、
外へ出る直前に人間が手で止めるしかなかった。

CLAUDE.md §3-3 にその事故が記録されている ――
「発表資料に設計書を典拠として『個人情報検査で自動中止』『30日で自動消去』と
書いたが、出荷既定は pii_scan_enabled=FALSE / knowledge_expire_days=0 だった。
部長への説明の直前に止めた。」
止めたのは偶然で、次は止まらない。実際、同じ嘘は一次文書側に残り続けていた。

そして、これが「実機テストのループが終わらない」の上流でもある。
**実機テスターが読む手順書が古いと、テスターの報告に「本当のバグ」と
「文書が古いだけ」が混ざる。** 司令塔はそれを毎回人力で切り分け、後者を
「修正」として実装ラウンドへ投入してしまう。バグが多いのではなく、
バグでないものがバグとして入ってくるからループが終わらない。

検査する4つ
------------------------------------------------------------------------------
1. config の既定値   … 文書が「既定N」と書いた値が build_config_rows と一致するか
2. 容量台帳          … CLAUDE.md §12 の「残N」が src の実測と一致するか
3. エラーコード      … src が吐く E**** が FriendlyMessage と運用ガイドにあるか
4. 配布文書のボタン名 … zip へ同梱する文書が、現行UIに無いボタン名を説明していないか

いずれも「10行のスクリプトで機械的に取れる事実」で、人が覚えておく類のもので
はない。典拠は必ず下流(src / build_config_rows / dist)から取る。

使い方
------------------------------------------------------------------------------
    python3 tools/doc_gate.py            # 全検査。ERROR があれば exit 1
    python3 tools/doc_gate.py --only cap # 検査を絞る(cfg / cap / err / ui)
"""
from __future__ import annotations

import argparse
import importlib.util
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"
DOCS = ROOT / "docs"

ERRORS: list[str] = []
WARNS: list[str] = []


def err(msg: str) -> None:
    ERRORS.append(msg)


def warn(msg: str) -> None:
    WARNS.append(msg)


def doc_files() -> list[Path]:
    """既定値の検査対象。

    **過去の裁定書(docs/dev/spec_*.md)は除く。** 裁定書は「そのとき何を何へ
    変えたか」の記録で、旧値が書いてあるのが正しい(例: R44 の裁定書が
    「sparse_keyscore_cap を 10 から 3 へ」と書くのは事実の記録)。
    そこを赤くすると、書き手は歴史を書き換えるか検査を無視するかのどちらかを
    選ぶことになる ―― どちらも失う。見るのは【いま読まれて行動の根拠になる
    文書】だけにする。
    """
    out = [ROOT / "CLAUDE.md", ROOT / "README.md"]
    if DOCS.is_dir():
        for f in sorted(DOCS.rglob("*.md")) + sorted(DOCS.rglob("*.html")):
            name = f.name
            # spec_ / audit_ は裁定と監査の記録。ファイル名に日付を持つ
            # 「その時点の記録」も同じ扱いにする(過去の事実を現在の嘘として
            # 叱ると、歴史を書き換える圧力になる)。
            if name.startswith("spec_") or name.startswith("audit_"):
                continue
            if re.search(r"_20\d{6}\.(md|html)$", name):
                continue
            out.append(f)
    return [p for p in out if p.is_file()]


# ==============================================================================
# 1. config の既定値 vs 文書の「既定N」
# ==============================================================================
def load_config_rows() -> dict:
    """build_mybookshelf.build_config_rows() を import して既定値を取る。
    設計書ではなく【ビルドが実際に焼くもの】が典拠(CLAUDE.md §3-3)。"""
    spec = importlib.util.spec_from_file_location(
        "_bm", ROOT / "build" / "build_mybookshelf.py")
    mod = importlib.util.module_from_spec(spec)
    sys.modules["_bm"] = mod
    spec.loader.exec_module(mod)
    rows = mod.build_config_rows(False, "", "1970-01-01 00:00")
    return {k: v for k, v, _ in rows}


# 文書が既定値を書く言い方。キー名と同じ行に現れたものだけ見る
#   例: `knowledge_expire_days`(既定30) / 既定 450000 / 既定値は 300
DEFAULT_PAT = re.compile(r"既定(?:値)?(?:は)?\s*[（(]?\s*([0-9][0-9,]*)")


# キーと既定値の距離。これより離れていたら「その key の既定」とは見なさない。
# 1行に複数キーが並ぶ表(FEATURES.md の1行にキー3つ)で、全キー×全数値を
# 総当たりすると誤検知しか出ない ―― 検査が嘘を言い始めると、書き手は
# 検査を無視するようになる。キーの【直後】に現れた最初の1つだけを見る。
NEAR_CHARS = 30


# 文書が bool の既定を書く言い方。数値と同じくキー名と同じ行だけを見る。
#   例: `pii_scan_enabled`(既定FALSE) / 既定 TRUE / 既定値は 有効
BOOL_PAT = re.compile(r"既定(?:値)?(?:は)?\s*[（(]?\s*(TRUE|FALSE|True|False|true|false|有効|無効|オン|オフ|ON|OFF)")
_BOOL_TRUE = {"TRUE", "True", "true", "有効", "オン", "ON"}
_BOOL_FALSE = {"FALSE", "False", "false", "無効", "オフ", "OFF"}

# 文書に現れるモデル名リテラル。出荷 config のどれとも一致しなければ ERROR。
#   R48 で gpt-5.5 → gpt-5.6-{luna,terra,sol} へ入れ替えたとき、
#   doc_gate は **str を1文字も検査していなかった**ため7ファイル12行が
#   古いモデル名のまま ERROR 0 で通った(監査 A-A-5 / H-6)。
# 「既定」の語を要求する。数値(DEFAULT_PAT)・bool(BOOL_PAT)と同じ作法。
# 要求しないと「値を gpt-5.6-terra に書き換える」のような
# **既定の記述ではない文**まで拾ってしまう(docs/09 の切り分け手順で実証)。
MODEL_PAT = re.compile(r"既定(?:値)?(?:は)?\s*[（(]?\s*`?(gpt-[0-9]+(?:\.[0-9]+)?(?:-[A-Za-z0-9]+)*)")


def check_config_defaults(cfg: dict) -> None:
    numeric = {k: v for k, v in cfg.items() if isinstance(v, int) and not isinstance(v, bool)}
    booleans = {k: v for k, v in cfg.items() if isinstance(v, bool)}
    for f in doc_files():
        text = f.read_text(encoding="utf-8", errors="ignore")
        for lineno, line in enumerate(text.splitlines(), 1):
            for key, real in numeric.items():
                start = 0
                while True:
                    k = line.find(key, start)
                    if k < 0:
                        break
                    start = k + len(key)
                    # 別のキーの一部(接頭辞が一致するだけ)なら見ない
                    tail = line[start:start + 1]
                    if tail and (tail.isalnum() or tail == "_"):
                        continue
                    m = DEFAULT_PAT.search(line, start, start + NEAR_CHARS)
                    if not m:
                        continue
                    got = int(m.group(1).replace(",", ""))
                    if got != real:
                        err(f"{f.relative_to(ROOT)}:{lineno} "
                            f"{key} の既定を {got} と書いていますが、"
                            f"出荷ビルドは {real} です"
                            "(典拠は build_config_rows。文書を直してください)")
            # R49(監査 A-A-5 / H-6): bool も同じ作法で照合する。
            # 初版は int だけを見ていたため、安全スイッチ(pii_scan_enabled 等)の
            # 嘘が構造的に検出できなかった ―― CLAUDE.md §3-3 の事故そのもの。
            for key, real_b in booleans.items():
                start = 0
                while True:
                    k = line.find(key, start)
                    if k < 0:
                        break
                    start = k + len(key)
                    tail = line[start:start + 1]
                    if tail and (tail.isalnum() or tail == "_"):
                        continue
                    mb = BOOL_PAT.search(line, start, start + NEAR_CHARS)
                    if not mb:
                        continue
                    got_b = mb.group(1) in _BOOL_TRUE
                    if got_b != real_b:
                        err(f"{f.relative_to(ROOT)}:{lineno} "
                            f"{key} の既定を {mb.group(1)} と書いていますが、"
                            f"出荷ビルドは {real_b} です"
                            "(典拠は build_config_rows。文書を直してください)")


def check_model_names(cfg: dict) -> None:
    """モデル名の既定を、**キー名の近傍に限って**照合する。

    R48 の実例: gpt-5.5 → gpt-5.6-{luna,terra,sol} へ入れ替えたのに、
    check_config_defaults は int だけを見ており str を1文字も検査していなかった。
    7ファイル12行が古いモデル名のまま「ERROR 0」で通った(監査 A-A-5 / H-6)。

    **文書全体から gpt-* を拾う形にはしない。** 初版でそれを試したところ、
    アドイン側のモデル一覧の説明(gpt-6-astra が luna を指している等)や
    過去の記録まで6件が誤検知になった。この検査自身のコメントが書いている
    とおり「検査が嘘を言い始めると、書き手は検査を無視するようになる」。
    数値・bool と同じく **キー名と同じ行・NEAR_CHARS 以内**だけを見る。
    """
    models = {k: v for k, v in cfg.items()
              if isinstance(v, str) and v.startswith("gpt-")}
    if not models:
        return
    for f in doc_files():
        text = f.read_text(encoding="utf-8", errors="ignore")
        for lineno, line in enumerate(text.splitlines(), 1):
            for key, real in models.items():
                start = 0
                while True:
                    k = line.find(key, start)
                    if k < 0:
                        break
                    start = k + len(key)
                    tail = line[start:start + 1]
                    if tail and (tail.isalnum() or tail == "_"):
                        continue
                    # 窓の右端で名前が切れないよう、探索は行末まで行い
                    # 「一致の開始位置が NEAR_CHARS 以内か」で近さを判定する
                    # (窓で切ると gpt-5.6-terra が gpt-5.6 に化けた)。
                    m = MODEL_PAT.search(line, start)
                    if not m or m.start() >= start + NEAR_CHARS:
                        continue
                    got = m.group(1)
                    if got != real:
                        err(f"{f.relative_to(ROOT)}:{lineno} "
                            f"{key} の既定を {got} と書いていますが、"
                            f"出荷ビルドは {real} です"
                            "(典拠は build_config_rows。文書を直してください)")


# 表形式の config 台帳。`| key | 既定値 | 意味 |` の**見出し行にだけ**
# 「既定値」があり、各行には無い。上の3検査は「キーの近くに『既定』の語」を
# 要求するので、**この表を1行も見ていなかった** ―― つまり MASTER_SPEC §5 の
# config キー台帳(約60行)は、R49 まで完全に無検査だった。
# R48 で古いモデル名が残った 12 行のうち、MASTER_SPEC のこの表の3行が
# まさにこれに当たる。近傍検査を足しただけでは**空振りしたまま**なので、
# 表の行そのものを読む。
TABLE_ROW_PAT = re.compile(r"^\s{0,3}\|(.+)\|\s*$")
_CELL_STRIP = " \t`*　"


def _cell(s: str) -> str:
    return s.strip(_CELL_STRIP).strip()


def check_config_table_rows(cfg: dict) -> None:
    """`| key | value | 説明 |` 形式の台帳行を照合する。

    **2列目が一意に読めるときだけ**見る。台帳には
      `| mock_llm | TRUE(開発ビルド)/FALSE(本番) |`
      `| topk_quick / topk_deep | 6 / 12 |`
      `| shelf_folder | (空) |`
    のように「1行に2つの値」「注釈つき」「空」が混ざっている。
    これらは機械には一意に読めないので**黙って飛ばす**(誤検知を出すくらいなら
    見ないほうがよい ―― 嘘を言う検査は無視されるようになる)。
    """
    if not cfg:
        return
    num_pat = re.compile(r"^-?[0-9][0-9,]*$")
    model_pat = re.compile(r"^gpt-[0-9A-Za-z.\-]+$")
    for f in doc_files():
        text = f.read_text(encoding="utf-8", errors="ignore")
        for lineno, line in enumerate(text.splitlines(), 1):
            m = TABLE_ROW_PAT.match(line)
            if not m:
                continue
            cells = [_cell(c) for c in m.group(1).split("|")]
            if len(cells) < 2:
                continue
            key, raw = cells[0], cells[1]
            if key not in cfg or not raw:
                continue
            real = cfg[key]
            if isinstance(real, bool):
                if raw not in _BOOL_TRUE and raw not in _BOOL_FALSE:
                    continue
                got_s, ok = raw, (raw in _BOOL_TRUE) == real
            elif isinstance(real, int):
                if not num_pat.match(raw):
                    continue
                got_s, ok = raw, int(raw.replace(",", "")) == real
            elif isinstance(real, str) and real.startswith("gpt-"):
                if not model_pat.match(raw):
                    continue
                got_s, ok = raw, raw == real
            else:
                continue
            if not ok:
                err(f"{f.relative_to(ROOT)}:{lineno} "
                    f"[台帳の表] {key} を {got_s} と書いていますが、"
                    f"出荷ビルドは {real} です"
                    "(典拠は build_config_rows。文書を直してください)")


# ==============================================================================
# 2. 容量台帳(CLAUDE.md §12)vs 実測
# ==============================================================================
CAP_PAT = re.compile(r"`?(mod[A-Za-z0-9]+|opt[A-Za-z0-9]+)`?\s*(?:は\s*)?残\s*([0-9][0-9,]*)")
# R49(監査 B-1): 台帳には「残」を書かない書式も混ざっている
#   例: **`modUIShelf`10** / `modViewport`118 / **`modShare`186**
# 初版はこの形を1行も見ておらず、**守るべき台帳の一部が無検査のまま**だった。
# 「`名前`直後の数字」という形はかなり限定的なので、誤検知は出にくい。
CAP_PAT_BARE = re.compile(r"`(mod[A-Za-z0-9]+|opt[A-Za-z0-9]+)`\s*([0-9][0-9,]{1,6})(?![0-9])")
MAX_CHARS = 30000


def check_capacity_ledger() -> None:
    actual = {}
    for p in SRC.rglob("*.bas"):
        actual[p.stem] = len(p.read_text(encoding="utf-8", errors="replace"))
    f = ROOT / "CLAUDE.md"
    if not f.is_file():
        return
    for lineno, line in enumerate(f.read_text(encoding="utf-8").splitlines(), 1):
        seen_spans = []
        for m in CAP_PAT.finditer(line):
            seen_spans.append(m.span())
            name, got = m.group(1), int(m.group(2).replace(",", ""))
            if name not in actual:
                continue
            real = MAX_CHARS - actual[name]
            if got != real:
                err(f"CLAUDE.md:{lineno} {name} の残りを {got} と書いていますが、"
                    f"実測は {real}({actual[name]}字)です"
                    "(台帳の数字は触るたびに python の len で測り直すこと)")
        # 「残」を書かない書式(`modX`NNN)。上で拾った範囲とは重ねない。
        for m in CAP_PAT_BARE.finditer(line):
            if any(a <= m.start() < b for a, b in seen_spans):
                continue
            name, got = m.group(1), int(m.group(2).replace(",", ""))
            if name not in actual:
                continue
            real = MAX_CHARS - actual[name]
            if got != real:
                err(f"CLAUDE.md:{lineno} {name} の残りを {got} と書いていますが、"
                    f"実測は {real}({actual[name]}字)です"
                    "(『残』を書かない書式も検査対象。R49 監査 B-1)")


# ------------------------------------------------------------------------------
# 台帳そのものを機械が書く(R49)
# ------------------------------------------------------------------------------
# なぜ: 検査を足した初回、台帳は **35 件ずれていた**。原因は数字の不注意では
# なく台帳の**構造**で、§12 は R43 / R46 / R48 / 受け皿 の【日付つきスナップ
# ショットが4枚重なった状態】だった ―― 同じモジュールが3か所に3つの違う数字で
# 載る。ラウンドごとに「前の行を直す」のではなく「新しい行を足す」ので、
# ずれは必ず増える。CLAUDE.md §12 自身が「伝聞のまま持ち回ると実装班が
# 『入る』と判断して入らない」と書いているとおりの事故が、台帳の中で起きていた。
#
# なので手で直さない。**1枚の実測表を機械が書き、機械が検査する。**
CAP_BEGIN = "<!-- CAP:BEGIN"
CAP_END = "<!-- CAP:END -->"
FROZEN = {"modAsk", "modBoot", "modShelf", "modRetrieve", "modPrompts"}
CAP_BANDS = [
    (0, 100, "実質凍結（残100字未満）", "**1行も入らない。** 実体を受け皿へ置き、ここからは1行呼び出しに留めること"),
    (100, 300, "逼迫（残300字未満）", "**分割裁定必須。** 次に触るなら追加の分割が先"),
    (300, 1000, "準逼迫（残1,000字未満）", "次に触るときは要注意"),
    (1000, 2000, "28,000〜29,000字帯（残1,000〜2,000）", "まだ入るが、まとまった追加は分割を考える"),
    (2000, 2100, "WARN直下（残2,000〜2,100）", "**lint は警告を出さないので台帳が唯一の記録。** コメント1行でWARN帯へ落ちる"),
    (2100, 10000, "余裕（残2,100〜10,000）", ""),
    (10000, 10 ** 9, "受け皿（残10,000以上）", "新しい実体の置き場所はここから選ぶ"),
]


def _measure_all() -> list[tuple[int, str]]:
    out = []
    for p in sorted(SRC.rglob("*.bas")):
        n = len(p.read_text(encoding="utf-8", errors="replace"))
        out.append((MAX_CHARS - n, p.stem))
    out.sort()
    return out


def render_capacity_ledger() -> str:
    rows = _measure_all()
    lines = [f"{CAP_BEGIN} 自動生成。手で書き換えない。"
             "`python3 tools/doc_gate.py --write-cap` で書き直し、"
             "`--only cap` が実測と照合する（R49 新設） -->"]
    over = [(r, n) for r, n in rows if r < 0]
    if over:
        body = " / ".join(f"**`{n}`は{-r}字超過**" for r, n in over)
        lines.append(f"- **🔴 上限30,000字を超えている**: {body}")
    for lo, hi, title, note in CAP_BANDS:
        sel = [(r, n) for r, n in rows if lo <= r < hi]
        if not sel:
            continue
        body = " / ".join(
            f"`{n}`{r:,}" + ("（凍結）" if n in FROZEN else "") for r, n in sel)
        tail = f" — {note}" if note else ""
        lines.append(f"- **{title}・{len(sel)}本**{tail}: {body}")
    lines.append(CAP_END)
    return "\n".join(lines)


def write_capacity_ledger() -> int:
    f = ROOT / "CLAUDE.md"
    text = f.read_text(encoding="utf-8")
    block = render_capacity_ledger()
    i = text.find(CAP_BEGIN)
    j = text.find(CAP_END)
    if i < 0 or j < 0:
        print("CLAUDE.md に CAP:BEGIN / CAP:END の目印がありません。"
              "§12 の台帳をこの2行で囲んでから実行してください。", file=sys.stderr)
        return 1
    new = text[:i] + block + text[j + len(CAP_END):]
    if new == text:
        print("容量台帳: 変更なし（実測と一致）")
        return 0
    f.write_text(new, encoding="utf-8")
    print("容量台帳を実測で書き直しました（CLAUDE.md §12）")
    return 0


# ==============================================================================
# 3. エラーコード: src が吐くもの ⊆ FriendlyMessage ⊆ 運用ガイド
# ==============================================================================
CODE_PAT = re.compile(r'"(E0\d{3})"')


def check_error_codes() -> None:
    emitted: dict[str, str] = {}
    for p in SRC.rglob("*.bas"):
        if "test" in p.parts:
            continue
        txt = p.read_text(encoding="utf-8", errors="replace")
        for lineno, line in enumerate(txt.splitlines(), 1):
            if "LogError" not in line:
                continue
            for m in CODE_PAT.finditer(line):
                emitted.setdefault(m.group(1), f"{p.relative_to(ROOT)}:{lineno}")

    friendly = SRC / "core" / "modLog.bas"
    known = set()
    if friendly.is_file():
        known = set(CODE_PAT.findall(friendly.read_text(encoding="utf-8", errors="replace")))

    guide = DOCS / "30_運用保守ガイド.md"
    documented = set()
    if guide.is_file():
        documented = set(re.findall(r"E0\d{3}", guide.read_text(encoding="utf-8", errors="ignore")))

    for code, where in sorted(emitted.items()):
        if code not in known:
            err(f"{where} が {code} を出しますが、modLog.FriendlyMessage に分岐が"
                "ありません(汎用文だけが出て、利用者に次の一手が伝わりません)")
        elif code not in documented:
            warn(f"{code} が docs/30_運用保守ガイド.md の表にありません"
                 f"(現場で出たとき保守担当が引けません。出所 {where})")


# ==============================================================================
# 4. 配布文書が、現行UIに無いボタン名を説明していないか
# ==============================================================================
def shipped_docs() -> list[Path]:
    """zip へ同梱する文書(build 側の一覧が正)。"""
    bm = (ROOT / "build" / "build_mybookshelf.py").read_text(encoding="utf-8")
    names = re.findall(r'"(\d\d_[^"]+\.(?:md|html))"', bm)
    return [DOCS / n for n in names if (DOCS / n).is_file()]


def ui_captions() -> set[str]:
    """現行UI(nexus_ui=TRUE 側)に実在するボタン名の語。"""
    words: set[str] = set()
    for p in SRC.rglob("*.bas"):
        txt = p.read_text(encoding="utf-8", errors="replace")
        for lit in re.findall(r'"([^"]{2,40})"', txt):
            if re.search(r"[ぁ-んァ-ヶ一-龥]", lit):
                words.add(lit.strip())
    return words


# 旧UI(nexus_ui=FALSE)にしか無い語。配布文書に出ていたら赤。
# 値は (正規表現, 説明)。正規表現にしているのは、正しい呼び名を部分一致で
# 拾ってしまう誤検知を避けるため ―― 実装の本物は「🩺 診断を開く」(modHelp の
# AddHelpAction)なので、「🩺 診断」だけを禁止すると正解まで赤くなる。
# 検査が正しい書き方を叱ると、書き手は検査を信用しなくなる。
LEGACY_TERMS = {
    r"マイ本棚」タブ": "旧シートUIのタブ名。Nexus 画面には存在しません",
    r"「ホーム」タブ": "旧シートUIのタブ名。Nexus 画面には存在しません",
    r"🩺\s*診断(?!を開く)": "この名前のボタンはありません(正しくは「🩺 診断を開く」)",
    r"🟢\s*解決した[!！]": "現行は「✅ 解決した」です",
    r"🟡\s*ヒントになった": "現行の評価ボタンにこの選択肢はありません",
    r"🔴\s*だめだった": "現行は「❌ 違う」です",
}


def check_shipped_docs_ui() -> None:
    ship = shipped_docs()
    if not ship:
        warn("配布zipへ同梱する文書の一覧を build から取れませんでした(検査未実施)")
        return
    for f in ship:
        text = f.read_text(encoding="utf-8", errors="ignore")
        for lineno, line in enumerate(text.splitlines(), 1):
            for pat, why in LEGACY_TERMS.items():
                m = re.search(pat, line)
                if m:
                    err(f"{f.relative_to(ROOT)}:{lineno} 「{m.group(0)}」— {why}"
                        "(配布物に入る文書なので、読んだ人がその画面を探して詰まります)")


# ==============================================================================
def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="文書検問(出荷ビルドを典拠に文書を検査する)")
    ap.add_argument("--only", choices=["cfg", "cap", "err", "ui"], help="検査を絞る")
    ap.add_argument("--write-cap", action="store_true",
                    help="CLAUDE.md §12 の容量台帳を実測で書き直す（検査はしない）")
    args = ap.parse_args(argv[1:])

    if args.write_cap:
        return write_capacity_ledger()

    cfg = load_config_rows()
    run = args.only
    if run in (None, "cfg"):
        check_config_defaults(cfg)
        check_model_names(cfg)
        check_config_table_rows(cfg)
    if run in (None, "cap"):
        check_capacity_ledger()
    if run in (None, "err"):
        check_error_codes()
    if run in (None, "ui"):
        check_shipped_docs_ui()

    print("=" * 78)
    print("文書検問(doc_gate) — 典拠は出荷ビルドと src の実測")
    print("=" * 78)
    for m in WARNS:
        print(f"  WARN  {m}")
    for m in ERRORS:
        print(f"  ERROR {m}")
    print("-" * 78)
    print(f"ERROR: {len(ERRORS)} 件 / WARN: {len(WARNS)} 件")
    if ERRORS:
        print("結果: NG(exit code 1) — 文書が出荷ビルドと食い違っています")
        return 1
    print("結果: OK(exit code 0)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
