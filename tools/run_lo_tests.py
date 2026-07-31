#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_lo_tests.py — LibreOffice headlessによるVBA実行テスト(MASTER_SPEC.md §11.2)
================================================================================
役割:
    vba_lint.py が「読むだけ」で見つけられる違反を検出するのに対し、本スクリプトは
    実際に LibreOffice(soffice --headless)にVBAソースを読み込ませ、
      モード1: 純ロジックモジュール一式を実行して modTestRunner.RunAllPureTests
               →ReportText の結果(PASS/FAIL)を回収する
      モード2: 全モジュール(Excel依存を含む)を1本ずつ隔離したライブラリに
               読み込み、実行はせず「コンパイルが通るか」だけを確認する
    の2段構えでテストする。

技術メモ(実験して確定させた挙動。ここが今回の技術リスクだった):
    1. soffice --headless で「vnd.sun.star.script:Lib.Mod.Sub?language=Basic&
       location=application」形式のURIを叩いてマクロを実行させるには、
       -env:UserInstallation=file://<profile> で指すプロファイルが
       「一度でも正常に起動を終えたことがある」状態でなければならない。
       手組みしただけの真っさらなプロファイル(script.xlc等だけ用意した状態)
       ではスクリプトが黙って実行されない(exit codeは0のまま何も起きない)。
       このため本スクリプトはまず --terminate_after_init で「一次起動」して
       プロファイルを初期化してから、そこへ Basic モジュールを注入する。
       (このテンプレートプロファイルは使い回して初期化コストを毎回払わない
       ようにしている。詳細は _ensure_template_profile を参照)
    2. モジュール間で Public Type(modTypes.ExtractedPage 等)を跨いで
       参照すると、既定(VBA非互換)のStarBasicコンパイラでは
       コンパイルがスタックし、呼び出しがハングする(タイムアウトでしか
       検知できない)。各モジュール先頭に "Option VBASupport 1" を追加すると
       この問題が解消することを実験で確認した。逆に、この1行が無いと
       modTypes を使う全モジュールのテストが原因不明のタイムアウトになる。
       そのため本スクリプトが注入する全モジュールの先頭に機械的に
       "Option VBASupport 1" を付与する(元の.bas/.clsファイルは変更しない。
       あくまでLO実行用に生成する一時コピーだけに付与する)。
    3. 1つのライブラリ内に構文エラーを含むモジュールが1つでもあると、
       そのライブラリ内の「どのマクロを呼んでも」呼び出しがハングする
       (ライブラリ単位でまとめてコンパイルされるため、壊れていない
       モジュールの実行も巻き添えを食う)。したがってモード2では
       モジュール1本ずつを専用ライブラリに隔離し、他モジュールの構文エラーに
       巻き込まれないようにしている。
    4. Excel固有オブジェクト(Worksheets/Range/Application/ThisWorkbook/
       MsgBox)は、それらに実際に「実行が到達」しない限りコンパイルは通る
       (未定義のグローバル識別子の解決は実行時に遅延される)。これを
       実験で確認済みなので、モード2は「対象モジュールのコンパイルが
       通るか」を、対象モジュール中の何かのPublicプロシージャを実際に
       呼ぶことはせず、同じライブラリに同居させたダミーの
       Chk_Driver.Probe() だけを呼ぶことで検査する(=ライブラリ全体の
       コンパイルを強制するが、対象モジュールの中身は実行しない)。
    5. タイムアウトの検知とプロセス後始末は、Pythonで自前のkill処理を
       書くより信頼できたため、coreutilsの `timeout --kill-after=N` に
       委譲している(実験でsoffice.binの完全終了を確認済み)。
    6. 【重要・既知のVBA/LO差異】 `Public Function Foo(...) As String()` の
       ように「配列を返す関数」の宣言は、Option VBASupport 1を付けても
       LibreOffice Basicではコンパイルが通らない(=ハングする)ことを実験で
       確認した(配列を「引数」として受け取るのは問題なく、配列を
       Variantに包んで返すのも問題ない。関数の戻り値型として配列型
       "T()" をそのまま書いた場合だけが壊れる)。これは実際に
       MASTER_SPEC §7.1 の modUtil.SplitKeepNonEmpty の契約シグネチャ
       ( `As String()` )がそのまま該当し、対処しないとLOテストが
       全滅する。§11.2の指示(「VBA固有でLOが解釈できない構文が出た場合は
       lint側で当該構文の代替を規約化する(勝手にテスト対象から外さない)」)
       に従い、本スクリプトは .xba へ変換する際にだけ
       `Function Foo(...) As T()` を `Function Foo(...) As Variant` へ
       機械的に書き換える(_fix_array_return_types)。元の.bas/.clsファイルは
       一切変更しない。関数本体が配列をそのままReturnValueに代入する分には
       Variant宣言でも実行時の挙動(呼び出し側で `Dim r() As String: r = ...`
       と受けてUBound/LBound/添字アクセスする)は同一であることを実証済み。
       実Excel(VBA)側は元の `As String()` のままビルドされるため、
       この書き換えはLO実行テストの内部実装だけの話であり、§7の公開契約
       (シグネチャ)そのものを変更するものではない。

使い方:
    python3 tools/run_lo_tests.py                  # モード1+モード2 両方
    python3 tools/run_lo_tests.py --mode pure       # モード1のみ
    python3 tools/run_lo_tests.py --mode compile    # モード2のみ
    python3 tools/run_lo_tests.py --keep-profile    # 一時プロファイルを残す(デバッグ用)
    exit code: 0 = 全テストPASS+全モジュールコンパイル成功 / 1 = いずれか失敗
================================================================================
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape

TOOLS_DIR = Path(__file__).resolve().parent
MYBOOKSHELF_ROOT = TOOLS_DIR.parent
DEFAULT_SRC_ROOT = MYBOOKSHELF_ROOT / "src"

SOFFICE_CANDIDATES = ["/usr/bin/soffice", "soffice"]

# モード1(純ロジック実行)に含めるモジュール(存在するものだけを注入する)
# 2026-07-11 Wave3(テスト完成担当)で追加: modAppDef/modShelfSync/modPack。
# MASTER_SPEC §7.8はmodShelfSync.DiffDecisionとmodPack.ValidatePackMetaを
# 「純関数として切り出してmodTestsPureから直接検証する」ことを明示指示している
# (これらのモジュール全体がR4準拠というわけではなく、Excelに触れる他のSub/
# Functionと同居しているが、この2関数自体はExcelオブジェクトに触れない)。
# 追加前はこの2モジュールが未注入のため、modTestsPure.RunAll内の
# modShelfSync.DiffDecision呼び出しが実行時エラー12(Variable not defined)に
# なり、テストが「実際のロジックを検証しないまま失敗扱い」になっていた。
# modPack.ValidatePackMetaはmodAppDef.PACK_FORMAT_VERSIONも参照するため
# modAppDefも合わせて追加する。3モジュールとも「対象モジュールを1本だけ
# 隔離してコンパイル」するモード2(run_compile_mode)で既にコンパイル成功が
# 確認済み(Excel専用トークンはtechメモ4のとおり未実行なら未解決のままで
# 良い)。追加後、実際にモード1を実行してPASS/FAIL件数の悪化がないことを
# 確認済み(tools/README.mdまたはWave3完了報告のverification参照)。
# 2026-07-12 Wave3-Tで追加: modTestsPure2。modTestsPureが§7.1の
# 「1モジュール30,000字以内」を超過したため、modPrompts/modShelfSync/
# modPack関連のテスト(TestModPrompts/TestModShelfSync/TestModPack)を
# modTestsPure2.RunAll2へ分割した(src/test/modTestsPure2.bas冒頭コメント
# 参照)。modTestsPure.RunAllの末尾がmodTestsPure2.RunAll2を呼ぶため、
# この一時ライブラリに注入しないと実行時エラー(Variable not defined)に
# なり分割先のテストが「実行されないまま」になる。
PURE_ALLOWLIST = [
    "modTypes", "modUtil", "modChunker", "modPii", "modPrompts",
    "modTestRunner", "modTestsPure", "modTestsPure2",
    "modAppDef", "modShelfSync", "modPack",
    # modSparse: 日本語キーワード検索(文字bigram+BM25+完全一致)。
    # 純ロジックなので実行テストで検証できる。検索精度の要なので必ず載せる。
    "modSparse", "modMode",
    # modChannel: 部門チャンネル。Excel依存が多いが、origin タグの組み立て
    # (ChannelOriginTag)だけは純関数で、ここがズレると切替・更新配信が
    # まるごと空振りする(2026-07-28 レビュー C-1 の実バグ)。コンパイルごと
    # 載せてタグ規約を実行テストで固定する。
    "modChannel",
    # modClarify: 聞き返し。IsNumberChoiceOnly は純関数で、ここがゆるいと
    # 利用者が打った質問が黙って捨てられる(2026-07-28 レビュー H-12 の実バグ)。
    "modClarify",
    # modStats: バッジ表の単一情報源(BadgeCatalog)。判定と表示で表が
    # 二重化して「獲得しても見えないバッジ」が4種あった(解説書 §11-11)。
    # 表の整合(4配列の長さ一致・id重複なし)はここで固定する。
    "modStats",
    # 2026-07-30 R2要件B/C対応で追加。modShelfSync/modPackと同じ考え方
    # (モジュール全体がR4準拠というわけではないが、テストで実際に呼ぶ関数
    # 自体はExcel/COMオブジェクトに触れない)。
    #   modExtractor: SharedCopyNextChunkLen(共有読みコピーの分割サイズ計算)
    #     だけが純ロジックだが、テストから modExtractor.SharedCopyNextChunkLen
    #     を呼ぶには本モジュール自体をこの一時ライブラリへ注入する必要がある。
    #     未注入のまま呼ぶと実行時エラー12(Variable not defined)になる
    #     (「対象モジュールを1本だけ隔離してコンパイル」するモード2では
    #     既にコンパイル成功を確認済み=Excel専用トークンは未実行なら
    #     未解決のままで良い、というtechメモ4のとおり)。
    #   modTestsPure3: modTestsPure/modTestsPure2とも30,000字上限まで残りが
    #     少なく、要件Bの新規テストを追加する場所が無かったための分割先
    #     (src/test/modTestsPure3.bas冒頭コメント参照)。modTestsPure2.RunAll2
    #     の末尾がmodTestsPure3.RunAll3を呼ぶため、この一時ライブラリに
    #     注入しないと同じく実行時エラー12になり分割先のテストが
    #     「実行されないまま」になる(modTestsPure2追加時と同型の理由)。
    "modExtractor", "modTestsPure3",
    # modChrome(2026-07-30 R4要件C/D): ツールバーとヘッダーピルの配置計算。
    # 「Wがいくつでも枠内に収まる/タイトルに重ならない」という保証は、
    # 実機で描いて目視するのではなくここで実行テストとして固定する。
    "modChrome",
    # 2026-07-31 R6(画像PDFのOCR取込)で追加。
    #   optOcrCore: Ghostscriptコマンド文字列の組み立てとページ上限の算数だけを
    #     持つ純ロジック(src/opt配下だが副作用ゼロ)。会社公式ツールが実際に
    #     踏んでいた「gsPathを引用符で囲み忘れる」バグを二度と出さないため、
    #     組み立て結果を1文字単位のゴールデンテストで固定する。
    #   modTestsPure4: modTestsPure3(24,111字)に要件R6のテストを足すと
    #     30,000字上限を超えるための分割先。modTestsPure3.RunAll3の末尾が
    #     modTestsPure4.RunAll4を呼ぶため、未注入だと実行時エラー12になり
    #     分割先のテストが「実行されないまま」になる(modTestsPure3と同型の理由)。
    "optOcrCore", "modTestsPure4",
]

TEMPLATE_PROFILE_DIR = Path(tempfile.gettempdir()) / "mybookshelf_lo_template_profile"

ATTR_LINE_PATTERN = re.compile(r"^\s*Attribute\s+")
# "Function Foo(...) As T()" のみ(引数側の "name() As T" は対象外)を検出する。
# 理由・実証結果は本ファイル冒頭コメントの技術メモ6を参照。
ARRAY_RETURN_TYPE_PATTERN = re.compile(
    r"(Function\s+\w+\s*\((?:[^()]|\([^()]*\))*\)\s*)As\s+([A-Za-z_]\w*)\s*\(\s*\)",
    re.IGNORECASE | re.DOTALL,
)

XBA_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
    '<script:module xmlns:script="http://openoffice.org/2000/script" '
    'script:name="{name}" script:language="StarBasic">{body}\n'
    "</script:module>\n"
)

XLB_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">\n'
    '<library:library xmlns:library="http://openoffice.org/2000/library" '
    'library:name="{libname}" library:readonly="false" library:passwordprotected="false">\n'
    "{elements}\n"
    "</library:library>\n"
)

XLC_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<!DOCTYPE library:libraries PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "libraries.dtd">\n'
    '<library:libraries xmlns:library="http://openoffice.org/2000/library" '
    'xmlns:xlink="http://www.w3.org/1999/xlink">\n'
    "{libs}\n"
    "</library:libraries>\n"
)

CHK_DRIVER_SRC = (
    "Option Explicit\n\n"
    "Public Function Probe() As Boolean\n"
    "    Probe = True\n"
    "End Function\n"
)


def find_soffice() -> str:
    for cand in SOFFICE_CANDIDATES:
        p = shutil.which(cand) or (cand if Path(cand).exists() else None)
        if p:
            return p
    print("[run_lo_tests] soffice が見つかりません。LibreOfficeをインストールしてください。")
    sys.exit(2)


CLS_HEADER_PATTERN = re.compile(
    r"^\s*(VERSION\s+[\d.]+\s+CLASS|BEGIN|MultiUse\s*=.*|END)\s*$", re.IGNORECASE)


def strip_attributes(text: str) -> str:
    """Attribute行に加え、.clsファイル先頭のクラスヘッダブロック
    (VERSION 1.0 CLASS / BEGIN / MultiUse=... / END)も除去する。
    これらはVBEのエクスポート形式であってBasicソースではないため、
    残すとLO Basicが構文エラー(ハング)になる。"""
    lines = text.splitlines()
    out = []
    in_header = True
    for l in lines:
        if in_header:
            if CLS_HEADER_PATTERN.match(l) or ATTR_LINE_PATTERN.match(l) or l.strip() == "":
                continue
            in_header = False
        if not ATTR_LINE_PATTERN.match(l):
            out.append(l)
    return "\n".join(out)


def fix_array_return_types(text: str) -> str:
    """"Function Foo(...) As T()" を "Function Foo(...) As Variant" に書き換える
    (LO実行専用の一時変換。理由は本ファイル冒頭コメントの技術メモ6を参照)。"""
    return ARRAY_RETURN_TYPE_PATTERN.sub(lambda m: m.group(1) + "As Variant", text)


def to_module_body(source_text: str) -> str:
    """Attribute行を除去し、配列返り値の宣言をVariantへ書き換え、先頭に
    Option VBASupport 1 を付与する(理由は本ファイル冒頭コメント参照)。"""
    body = strip_attributes(source_text)
    body = fix_array_return_types(body)
    return "Option VBASupport 1\n" + body


def write_module_xba(lib_dir: Path, name: str, source_text: str) -> None:
    body = xml_escape(to_module_body(source_text))
    xba = XBA_TEMPLATE.format(name=name, body=body)
    (lib_dir / f"{name}.xba").write_text(xba, encoding="utf-8")


def write_library(profile_dir: Path, lib_name: str, modules: dict[str, str]) -> None:
    """modules: {モジュール名: 元の.bas/.clsソーステキスト}"""
    lib_dir = profile_dir / "user" / "basic" / lib_name
    lib_dir.mkdir(parents=True, exist_ok=True)
    elements = []
    for name, src in modules.items():
        write_module_xba(lib_dir, name, src)
        elements.append(f' <library:element library:name="{name}"/>')
    xlb = XLB_TEMPLATE.format(libname=lib_name, elements="\n".join(elements))
    (lib_dir / "script.xlb").write_text(xlb, encoding="utf-8")


def register_libraries(profile_dir: Path, lib_names: list[str]) -> None:
    libs = [' <library:library library:name="Standard" library:link="false"/>']
    for n in lib_names:
        libs.append(f' <library:library library:name="{n}" library:link="false"/>')
    xlc = XLC_TEMPLATE.format(libs="\n".join(libs))
    (profile_dir / "user" / "basic" / "script.xlc").write_text(xlc, encoding="utf-8")


def ensure_template_profile(soffice: str, verbose: bool) -> Path:
    """一度だけ soffice を「一次起動」させて雛形プロファイルを作る(§技術メモ1)。
    既にあれば使い回す(--fresh-template で強制作り直し可能)。"""
    marker = TEMPLATE_PROFILE_DIR / "user" / "basic" / "Standard" / "script.xlb"
    if marker.exists():
        return TEMPLATE_PROFILE_DIR
    if verbose:
        print(f"[run_lo_tests] 雛形プロファイルを初期化中: {TEMPLATE_PROFILE_DIR}")
    if TEMPLATE_PROFILE_DIR.exists():
        shutil.rmtree(TEMPLATE_PROFILE_DIR)
    cmd = [
        "timeout", "--kill-after=5", "60",
        soffice, "--headless", "--invisible", "--nologo", "--norestore",
        f"-env:UserInstallation=file://{TEMPLATE_PROFILE_DIR}",
        "--terminate_after_init",
    ]
    subprocess.run(cmd, capture_output=True, text=True)
    if not marker.exists():
        print("[run_lo_tests] 雛形プロファイルの初期化に失敗しました(soffice起動不可の可能性)。")
        sys.exit(2)
    return TEMPLATE_PROFILE_DIR


def fresh_profile_copy(template: Path, dest: Path) -> None:
    shutil.copytree(template, dest)


def run_uri(soffice: str, profile_dir: Path, uri: str, timeout_sec: int) -> tuple[int, str, str]:
    cmd = [
        "timeout", "--kill-after=5", str(timeout_sec),
        soffice, "--headless", "--invisible", "--nologo", "--norestore",
        f"-env:UserInstallation=file://{profile_dir}",
        uri,
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    return proc.returncode, proc.stdout, proc.stderr


def discover_modules(src_root: Path) -> list[tuple[str, Path]]:
    """(モジュール名, パス) のリスト。モジュール名はAttribute VB_Nameがあればそれ、
    無ければファイル名(拡張子抜き)。"""
    out = []
    for path in sorted(list(src_root.rglob("*.bas")) + list(src_root.rglob("*.cls"))):
        text = path.read_text(encoding="utf-8", errors="replace")
        name = path.stem
        m = re.search(r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]*)"', text, re.MULTILINE)
        if m:
            name = m.group(1)
        out.append((name, path))
    return out


# ==============================================================================
# モード1: 純ロジック実行(RunAllPureTests → ReportText)
# ==============================================================================
def run_pure_mode(soffice: str, template: Path, all_modules: dict[str, Path],
                   work_dir: Path, timeout_sec: int, verbose: bool) -> tuple[bool, str]:
    print("=" * 78)
    print("モード1: LibreOffice上で純ロジックテスト(modTestRunner.RunAllPureTests)を実行")
    print("=" * 78)

    pure_srcs: dict[str, str] = {}
    missing = []
    for name in PURE_ALLOWLIST:
        path = all_modules.get(name)
        if path is None:
            missing.append(name)
            continue
        pure_srcs[name] = path.read_text(encoding="utf-8", errors="replace")

    if missing:
        print(f"  (未実装のため注入をスキップ: {', '.join(missing)})")

    if "modTestRunner" not in pure_srcs:
        print("  modTestRunner.bas が見つからないためモード1を実行できません。")
        return False, "modTestRunner.bas not found"

    out_path = work_dir / "pure_result.txt"
    if out_path.exists():
        out_path.unlink()

    test_main_src = (
        "Option Explicit\n\n"
        "Sub Main\n"
        "    On Error Resume Next\n"
        "    Err.Clear\n"
        "    modTestRunner.RunAllPureTests\n"
        "    Dim runErr As String\n"
        "    If Err.Number <> 0 Then\n"
        '        runErr = "RUNNER_ERROR " & Err.Number & ": " & Err.Description\n'
        "        Err.Clear\n"
        "    End If\n"
        "    On Error GoTo 0\n\n"
        "    Dim iFile As Integer\n"
        "    iFile = FreeFile\n"
        f'    Open "{out_path.as_posix()}" For Output As #iFile\n'
        "    Print #iFile, modTestRunner.ReportText()\n"
        "    If Len(runErr) > 0 Then Print #iFile, runErr\n"
        "    Close #iFile\n"
        "End Sub\n"
    )

    profile_dir = work_dir / "profile_pure"
    fresh_profile_copy(template, profile_dir)
    modules = dict(pure_srcs)
    modules["TestMain"] = test_main_src
    write_library(profile_dir, "MbPureRun", modules)
    register_libraries(profile_dir, ["MbPureRun"])

    uri = "vnd.sun.star.script:MbPureRun.TestMain.Main?language=Basic&location=application"
    rc, out, err = run_uri(soffice, profile_dir, uri, timeout_sec)

    if not out_path.exists():
        msg = f"実行結果ファイルが生成されませんでした(soffice exit={rc}, timeout={timeout_sec}s)"
        print(f"  FAIL: {msg}")
        if verbose:
            print(f"    stdout: {out.strip()}")
            print(f"    stderr: {err.strip()}")
        return False, msg

    report = out_path.read_text(encoding="utf-8", errors="replace").strip()
    print(report if report else "(空の結果)")

    m = re.search(r"FAIL\s+(\d+)", report)
    fail_count = int(m.group(1)) if m else None
    has_runner_error = "RUNNER_ERROR" in report

    ok = (fail_count == 0) and not has_runner_error
    return ok, report


# ==============================================================================
# モード2: 全モジュールのコンパイルチェック(実行はしない)
# ==============================================================================
def safe_lib_name(module_name: str) -> str:
    return "Chk_" + re.sub(r"[^A-Za-z0-9_]", "_", module_name)


def run_compile_mode(soffice: str, template: Path, all_modules: dict[str, Path],
                      work_dir: Path, timeout_sec: int, verbose: bool) -> tuple[bool, list[tuple[str, bool, str]]]:
    print("\n" + "=" * 78)
    print("モード2: 全モジュールの構文コンパイルチェック(実行はしない)")
    print("=" * 78)

    modtypes_path = all_modules.get("modTypes")
    modtypes_src = modtypes_path.read_text(encoding="utf-8", errors="replace") if modtypes_path else None

    results: list[tuple[str, bool, str]] = []
    all_ok = True

    for name, path in sorted(all_modules.items()):
        target_src = path.read_text(encoding="utf-8", errors="replace")
        lib_name = safe_lib_name(name)

        modules_for_lib: dict[str, str] = {}
        if name != "modTypes" and modtypes_src is not None:
            modules_for_lib["modTypes"] = modtypes_src
        modules_for_lib[name] = target_src
        modules_for_lib["Chk_Driver"] = CHK_DRIVER_SRC

        profile_dir = work_dir / f"profile_{lib_name}"
        fresh_profile_copy(template, profile_dir)
        write_library(profile_dir, lib_name, modules_for_lib)
        register_libraries(profile_dir, [lib_name])

        uri = f"vnd.sun.star.script:{lib_name}.Chk_Driver.Probe?language=Basic&location=application"
        t0 = time.time()
        rc, out, err = run_uri(soffice, profile_dir, uri, timeout_sec)
        elapsed = time.time() - t0

        ok = (rc == 0)
        detail = f"exit={rc} ({elapsed:.1f}s)"
        if rc == 124:
            detail = f"タイムアウト({timeout_sec}s) — 構文エラーの疑い"
        results.append((name, ok, detail))
        all_ok = all_ok and ok

        status = "PASS" if ok else "FAIL"
        print(f"  {status:<4} {name:<24} {detail}")

        # 使い終わったプロファイルは都度削除してディスクを節約
        shutil.rmtree(profile_dir, ignore_errors=True)

    return all_ok, results


def main() -> int:
    parser = argparse.ArgumentParser(description="マイ本棚AI LibreOffice実行テスト")
    parser.add_argument("--path", type=str, default=str(DEFAULT_SRC_ROOT), help="対象src(既定: mybookshelf/src)")
    parser.add_argument("--mode", choices=["pure", "compile", "all"], default="all")
    parser.add_argument("--pure-timeout", type=int, default=120, help="モード1のタイムアウト秒(既定120)")
    parser.add_argument("--compile-timeout", type=int, default=15, help="モード2の1モジュールあたりタイムアウト秒(既定15)")
    parser.add_argument("--keep-profile", action="store_true", help="一時プロファイルを削除せず残す(デバッグ用)")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    src_root = Path(args.path).resolve()
    if not src_root.exists():
        print(f"[run_lo_tests] 対象ディレクトリが存在しません: {src_root}")
        return 2

    soffice = find_soffice()
    template = ensure_template_profile(soffice, args.verbose)

    modules_list = discover_modules(src_root)
    all_modules: dict[str, Path] = {name: path for name, path in modules_list}
    print(f"[run_lo_tests] soffice={soffice}")
    print(f"[run_lo_tests] 対象モジュール数: {len(all_modules)} (in {src_root})")

    work_dir = Path(tempfile.mkdtemp(prefix="mybookshelf_lo_run_"))
    overall_ok = True
    try:
        if args.mode in ("pure", "all"):
            ok, _report = run_pure_mode(soffice, template, all_modules, work_dir, args.pure_timeout, args.verbose)
            overall_ok = overall_ok and ok

        if args.mode in ("compile", "all"):
            ok, _results = run_compile_mode(soffice, template, all_modules, work_dir, args.compile_timeout, args.verbose)
            overall_ok = overall_ok and ok
    finally:
        if args.keep_profile:
            print(f"\n[run_lo_tests] --keep-profile 指定のため一時ディレクトリを残します: {work_dir}")
        else:
            shutil.rmtree(work_dir, ignore_errors=True)

    print("\n" + "-" * 78)
    print("結果: OK(exit code 0)" if overall_ok else "結果: NG(exit code 1)")
    print("-" * 78)
    return 0 if overall_ok else 1


if __name__ == "__main__":
    sys.exit(main())
