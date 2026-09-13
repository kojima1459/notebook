#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_launcher.py — 起動ランチャー bat の静的検査(spec_20260909_R42 §2-5)
================================================================================
役割:
    build/build_mybookshelf.py の _launcher_bat_text() は純関数(引数の
    xlsm_name から bat の中身の文字列を組み立てるだけ)なので、実機でも
    LibreOffice でも検証できない bat 特有の落とし穴(ラベルの重複・
    goto先の未定義・CP932エンコード不能・.closed印とSIDの対応・xcopy /D
    の重複)を、生成された文字列に対して機械的に検査する。

    R42 B1〜B4(版の印・起動番号・成否検知・Ghostscript両ファイル確認)の
    実装後に生まれうる事故を検問で止めるための小さな専用チェッカー。
    final-gates の一環として実行する。

検査項目(spec §2-5):
    (1) CP932 へエンコードできる(配布zipへ書き出せる)
    (2) ラベル :RETIRE / :COPY_NEW / :COPY_LEGACY / :COPY_DONE /
        :COPY_FAIL / :BACKUP_FAIL / :GS_COPY / :GS_DONE / :FINAL_COPY /
        :WAIT_AND_MARK が、それぞれちょうど1回だけ定義されている
    (3) "goto :X" の X が、:EOF(cmd.exeの組み込みラベル)を除き、
        すべて bat 内に定義済みのラベルである
    (4) ".closed.%SID%"(印のファイル名)と "%~2"(__wait__側の受け取り)
        が対応している。具体的には:
          - ".closed.%SID%" を参照する行が存在する(自分の印を見る)
          - "%~2" を参照する行が存在する(__wait__がSIDを受け取る)
          - __wait__ を起動する行が %SID% を第2引数として渡している
    (5) 本体(%XLSM%)を対象にした "xcopy ... /D" が :COPY_LEGACY の
        区間(ラベル:COPY_LEGACY 〜 次のラベル)にちょうど1回だけ現れる
        (:COPY_NEW 側は強制上書きの copy /Y であり、/D 判定は
        「同じ版」の経路にだけ残す設計=LP-L04/DTA-16)
    (6) copy/move/xcopy の直後に成否検査(if errorlevel / if exist)がある
        (レビュー R42 m5: B3 の安全弁の欠落を検出する)
    (7) 版の印の比較に fc(外部コマンド)を使わず set /p で読む
        (レビュー R42 M3)

使い方:
    python3 tools/check_launcher.py
    終了コード 0 = 全項目OK、1 = いずれか NG。
"""
import os
import re
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, ROOT)

from build.build_mybookshelf import _launcher_bat_text  # noqa: E402

REQUIRED_LABELS = [
    "RETIRE", "COPY_NEW", "COPY_LEGACY", "COPY_DONE", "COPY_FAIL",
    "BACKUP_FAIL", "BUILD_FAIL", "RETIRE_WARN", "GS_COPY", "GS_WARN", "GS_DONE",
    "FINAL_COPY", "WAIT_AND_MARK",
]

# (6) レビュー R42 m5: copy/move/xcopy の直後(rem と空行を挟んでよい)には必ず
#     成否を見る行が来ること。B3 の安全弁が1本でも欠けると見つかるようにする。
#     2周目 MAJOR-1: 「成否を見る行」は `if [not] errorlevel …` か、
#     `if not exist "…" goto :X`(複製先の事後確認)に限る。無関係な
#     `if exist "…" copy …` を検査と誤認しない。
#     行頭が `if … copy` の1行 if 形(任意の複製: .build の同版時複製・手置き
#     _旧版 の複製)は「失敗してよい任意複製」として対象外(2周目 MINOR-3 記録)。
COPY_CMD_RE = re.compile(r'^(copy|move|xcopy)\b', re.IGNORECASE)
CHECK_LINE_RE = re.compile(
    r'^if\s+(not\s+)?errorlevel\b|^if\s+not\s+exist\s+"[^"]+"\s+goto\s+:', re.IGNORECASE)


def check_copy_followed_by_check(lines, errors: list) -> None:
    for i, ln in enumerate(lines):
        s = ln.strip()
        if not COPY_CMD_RE.match(s):
            continue
        j = i + 1
        while j < len(lines):
            t = lines[j].strip()
            if t == "" or t.lower().startswith("rem"):
                j += 1
                continue
            break
        nxt = lines[j].strip() if j < len(lines) else ""
        if not CHECK_LINE_RE.match(nxt):
            errors.append(f"(6) 行{i} の {s.split()[0]} の直後に成否検査(if errorlevel/if exist)がありません: 次行=[{nxt}]")


def check_build_compare(text: str, errors: list) -> None:
    # (7) レビュー R42 M3: 版の印の比較は外部コマンド fc ではなく set /p で行う。
    if re.search(r'^\s*fc\b', text, re.IGNORECASE | re.MULTILINE):
        errors.append("(7) fc(外部コマンド)が使われています。版の印の比較は set /p で行うこと")
    if "set /p SB=<" not in text or "set /p DB=<" not in text:
        errors.append("(7) 版の印を set /p で読む行(SB/DB)がありません")

LABEL_DEF_RE = re.compile(r'^:([A-Za-z0-9_]+)\s*$')
GOTO_RE = re.compile(r'goto\s+:([A-Za-z0-9_]+)', re.IGNORECASE)


def _all_label_defs(lines):
    """行番号→ラベル名 の一覧(定義側 ":LABEL" のみ。goto先ではない)。"""
    defs = []
    for i, ln in enumerate(lines):
        m = LABEL_DEF_RE.match(ln.strip())
        if m:
            defs.append((i, m.group(1)))
    return defs


def check_cp932(text: str, errors: list) -> None:
    try:
        text.encode("cp932")
    except UnicodeEncodeError as e:
        errors.append(f"(1) CP932エンコード不能: {e}")


def check_label_counts(lines, errors: list) -> None:
    defs = _all_label_defs(lines)
    counts = {}
    for _i, name in defs:
        counts[name] = counts.get(name, 0) + 1
    for name in REQUIRED_LABELS:
        n = counts.get(name, 0)
        if n != 1:
            errors.append(f"(2) ラベル :{name} が{n}回定義されています(期待1回)")


def check_goto_targets(lines, errors: list) -> None:
    defs = _all_label_defs(lines)
    defined = {name for _i, name in defs}
    missing = set()
    for ln in lines:
        for m in GOTO_RE.finditer(ln):
            target = m.group(1)
            if target.upper() == "EOF":
                continue
            if target not in defined:
                missing.add(target)
    for target in sorted(missing):
        errors.append(f"(3) goto :{target} の飛び先ラベルが未定義です")


def check_sid_mark(text: str, lines, errors: list) -> None:
    if ".closed.%SID%" not in text:
        errors.append('(4) ".closed.%SID%"(自分の印)を参照する行がありません')
    if "%~2" not in text:
        errors.append('(4) "%~2"(__wait__側のSID受け取り)を参照する行がありません')
    wait_start_re = re.compile(r'"%~f0"\s+\S*__wait__\S*\s+%SID%')
    if not any(wait_start_re.search(ln) for ln in lines):
        errors.append('(4) __wait__ を %SID% を第2引数に付けて起動する行が見つかりません')


def check_xcopy_d_scope(lines, errors: list) -> None:
    xcopy_d_xlsm_re = re.compile(r'xcopy\b.*%XLSM%.*\/D\b', re.IGNORECASE)
    hits = [i for i, ln in enumerate(lines) if xcopy_d_xlsm_re.search(ln)]
    if len(hits) != 1:
        errors.append(
            f"(5) 本体(%XLSM%)への xcopy .../D が{len(hits)}箇所にあります(期待1箇所): "
            + ", ".join(str(i) for i in hits)
        )
        return
    # :COPY_LEGACY 〜 次のラベル定義行、の区間に収まっているか。
    legacy_idx = None
    next_label_idx = None
    for i, ln in enumerate(lines):
        stripped = ln.strip()
        if stripped == ":COPY_LEGACY":
            legacy_idx = i
        elif legacy_idx is not None and next_label_idx is None and i > legacy_idx:
            if LABEL_DEF_RE.match(stripped):
                next_label_idx = i
    if legacy_idx is None:
        errors.append("(5) :COPY_LEGACY ラベルが見つかりません")
        return
    hit = hits[0]
    upper = next_label_idx if next_label_idx is not None else len(lines)
    if not (legacy_idx < hit < upper):
        errors.append(
            f"(5) xcopy .../D (行{hit}) が :COPY_LEGACY の区間"
            f"(行{legacy_idx}〜{upper})の外にあります"
        )


def run_checks(xlsm_name: str) -> list:
    text = _launcher_bat_text(xlsm_name)
    lines = text.split("\r\n")
    errors = []
    check_cp932(text, errors)
    check_label_counts(lines, errors)
    check_goto_targets(lines, errors)
    check_sid_mark(text, lines, errors)
    check_xcopy_d_scope(lines, errors)
    check_copy_followed_by_check(lines, errors)
    check_build_compare(text, errors)
    return errors


def main() -> int:
    targets = ["MyBookshelf.xlsm", "MyBookshelf_発行者用.xlsm", "MyBookshelf_dev.xlsm"]
    overall_ok = True
    for name in targets:
        errors = run_checks(name)
        if errors:
            overall_ok = False
            print(f"NG: {name}")
            for e in errors:
                print(f"  - {e}")
        else:
            print(f"OK: {name}")
    if overall_ok:
        print("check_launcher: OK(全項目)")
        return 0
    print("check_launcher: NG")
    return 1


if __name__ == "__main__":
    sys.exit(main())
