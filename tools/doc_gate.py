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


def check_config_defaults(cfg: dict) -> None:
    numeric = {k: v for k, v in cfg.items() if isinstance(v, int) and not isinstance(v, bool)}
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


# ==============================================================================
# 2. 容量台帳(CLAUDE.md §12)vs 実測
# ==============================================================================
CAP_PAT = re.compile(r"`?(mod[A-Za-z0-9]+|opt[A-Za-z0-9]+)`?\s*(?:は\s*)?残\s*([0-9][0-9,]*)")
MAX_CHARS = 30000


def check_capacity_ledger() -> None:
    actual = {}
    for p in SRC.rglob("*.bas"):
        actual[p.stem] = len(p.read_text(encoding="utf-8", errors="replace"))
    f = ROOT / "CLAUDE.md"
    if not f.is_file():
        return
    for lineno, line in enumerate(f.read_text(encoding="utf-8").splitlines(), 1):
        for m in CAP_PAT.finditer(line):
            name, got = m.group(1), int(m.group(2).replace(",", ""))
            if name not in actual:
                continue
            real = MAX_CHARS - actual[name]
            if got != real:
                err(f"CLAUDE.md:{lineno} {name} の残りを {got} と書いていますが、"
                    f"実測は {real}({actual[name]}字)です"
                    "(台帳の数字は触るたびに python の len で測り直すこと)")


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
    args = ap.parse_args(argv[1:])

    cfg = load_config_rows()
    run = args.only
    if run in (None, "cfg"):
        check_config_defaults(cfg)
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
