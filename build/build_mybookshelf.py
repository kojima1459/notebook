#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_mybookshelf.py — 「マイ本棚AI」ビルドスクリプト。
mybookshelf/dist/MyBookshelf.xlsm (または MyBookshelf_dev.xlsm) を生成する。

--------------------------------------------------------------------------
アーキテクチャ(V2 = /home/user/notebook/build/build_chatbot_v2.py で実証済みの機構を
mybookshelf/ 配下へ自己完結コピーしたもの。V2側のファイルは一切 import/変更しない):

  1. build/template_skeleton.xlsm — 本物のExcelで作られた.xlsmのコピー。
     中の vbaProject.bin は「本物」であり、どのExcelでも文句なく開ける。
  2. openpyxl(keep_vba=True)でこのスケルトンを読み込み、MASTER_SPEC §4の
     全13シート(使い方/ホーム/マイ本棚/ダッシュボード/config/my_knowledge/
     my_vectors/my_manifest/my_stats/usage_log/err_log/ui_state/vba_src)を
     生成する。この時点では vbaProject.bin はスケルトンのバイト列のまま
     一切触れていない。
  3. vba_src シートに、modules.json に列挙された標準モジュール(*.bas)の
     ソースを1行1モジュールで格納する(role=core/opt/testを問わず、
     modules.json に載っていて実在するものは全部載せる。撤去したいときは
     modules.json から行を消せばよい)。
  4. openpyxlで保存(vbaProject.binはスケルトン由来のまま)。
  5. 保存後の .xlsm から vbaProject.bin だけを取り出し、外科的パッチを当てる:
       - VBA/ThisWorkbook ストリームを「自己インストーラ」のソースに差し替え。
         このインストーラは Workbook_Open で vba_src シートを読み、
         VBProject.VBComponents.Add で標準モジュールを全部インストールし、
         最後に Application.Run "modBoot.Boot" を呼ぶ。
       - VBA/dir ストリームの ThisWorkbook.MOFFSET を 0 に書き換える
         (パフォーマンスキャッシュのオフセットが無効になるため、
         Excelに「バイト0からソースを解凍しろ」と教える)。
     このパッチは stream 単位のバイト長を変えられない(olefile.write_streamの
     制約)ため、OVBA空チャンクでpaddingしてぴったりの長さに揃える。
     CFBヘッダ・FAT・_VBA_PROJECT・PROJECT/PROJECTwm・他の全ストリームは
     一切触れない。
  6. ビルド後、生成物を再オープンして自己検証する(§10)。失敗したら exit 1。

このスクリプトが書き込むのは mybookshelf/build/ 配下と mybookshelf/dist/ 配下
のみ。/home/user/notebook/build/, /home/user/notebook/src/, /home/user/notebook/docs/
(V2チャットボットの資産)には一切触れない。
--------------------------------------------------------------------------
"""

from __future__ import annotations

import argparse
import datetime
import io
import json
import os
import re
import struct
import subprocess
import sys
import tempfile
import zipfile

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import olefile

# ovba.py は同ディレクトリの自己完結モジュール(OVBA圧縮/解凍・CFBリーダー)。
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ovba

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_ROOT = os.path.dirname(SCRIPT_DIR)                      # mybookshelf/
DEFAULT_TEMPLATE = os.path.join(SCRIPT_DIR, "template_skeleton.xlsm")
DEFAULT_MODULES_JSON = os.path.join(SCRIPT_DIR, "modules.json")

APP_TITLE = "マイ本棚AI"
EXCEL_CELL_LIMIT = 32000        # Excelの技術上限(セル1個あたりの文字数)
MODULE_CONTRACT_LIMIT = 30000   # MASTER_SPEC §7 の契約上限(1モジュールあたり)

# 軽量マクロ無効ガード: マクロが無効なまま開かれた場合(Boot未実行)に最初に
# 見える案内シート。ビルド側の責務は「先頭シート・アクティブ・visible」で
# 保存するところまで。Boot成功後にこれを隠すのはランタイム側(modBoot等)の
# 責務であり、このビルドスクリプトは一切hideしない。
GUARD_SHEET_NAME = "はじめにお読みください"

# MASTER_SPEC §4 のシート定義(名前 -> 可視性)。ビルド完了判定・自己検証の両方で使う。
EXPECTED_SHEETS = {
    GUARD_SHEET_NAME: "visible",
    "使い方": "visible",
    "ホーム": "visible",
    "マイ本棚": "visible",
    "ダッシュボード": "visible",
    "config": "hidden",
    "my_knowledge": "veryHidden",
    "my_vectors": "veryHidden",
    # 初期ナレッジ(同梱シード)。ビルド時に焼き込み、初回起動で modSeed が
    # my_knowledge / my_vectors へ写す。利用者には一切見せない。
    "seed_meta": "veryHidden",
    "seed_chunks": "veryHidden",
    "seed_vectors": "veryHidden",
    "my_manifest": "hidden",
    "my_stats": "hidden",
    "usage_log": "hidden",
    "err_log": "hidden",
    "ui_state": "veryHidden",
    "insight_inbox": "hidden",
    "vba_src": "veryHidden",
}

# 可視4シートのタブ色 (MASTER_SPEC §14 手順6): 使い方=緑, ホーム=青, マイ本棚=オレンジ, ダッシュボード=紫
TAB_COLORS = {
    "使い方": "00B050",
    "ホーム": "0070C0",
    "マイ本棚": "ED7D31",
    "ダッシュボード": "7030A0",
}

_ILLEGAL_XML = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')


def _clean(s):
    """XMLに書けない制御文字を除去する(V2実証済みの防御処理を踏襲)。"""
    return _ILLEGAL_XML.sub("", s) if isinstance(s, str) else s


class BuildError(Exception):
    """ビルド契約違反(モジュール欠落・文字数超過など)。呼び出し側でexit 1にする。"""


# ---------------------------------------------------------------------------
# direct埋め込み(Azure)のURL/キーは、ソースにハードコードせず環境変数から
# 読む(社内リポジトリに平文の認証情報を追加で増やさないため)。
# 未設定なら空文字列のまま(config側は空欄扱いになり、modGateway.DirectEmbedSlice
# が「azure_embed_url/keyが未設定」を検知して自動的にribbon経路へフォール
# バックする。空文字列でもビルド・実行時エラーにはならない)。
# ---------------------------------------------------------------------------
AZURE_EMBED_URL_ENV = "AZURE_EMBED_URL"
AZURE_EMBED_KEY_ENV = "AZURE_EMBED_KEY"

# 正典を発行できる端末に入れる合言葉(config publish_key)。
# 2026-07-28(解説書 §12.3 B2): これが入ったブックを一般配布すると、
# 受け取った全員が部門の正典を上書き発行できてしまう。
# 発行者用ビルドは --publisher を明示したときだけ作られ、出力ファイル名も
# 変わる(見分けがつかない2つのブックを作らないため)。
PUBLISH_KEY_ENV = "MYBOOKSHELF_PUBLISH_KEY"

# configシートに平文で置かないための軽い難読化(暗号的な秘匿ではない。
# VBAプロジェクト自体に触れる人には無意味 = このアプリの配布モデル上、
# それ以上の防御は不可能。configシートを開いただけの人の目に平文キーが
# 直接触れないようにする程度の対策)。VBA側 modUtil.DeobfuscateSecret と
# 対になる実装(XOR + 16進エンコード)。鍵・アルゴリズムを変える場合は
# 両方を同時に直すこと。
_OBF_PREFIX = "OBF1:"
_OBF_KEY = "NexusAgentBuildObfuscationKey2026"


def obfuscate_secret(plain: str) -> str:
    if not plain:
        return ""
    # VBA側 modUtil.XorWithObfKey はバイト志向(Chr$/Asc)の実装で、XOR結果が
    # 0x7F以下(ASCII平文 xor ASCII鍵は最上位ビットが立たない)である前提の
    # 上でしか正しく動かない。平文(環境変数由来。運用ミスで非ASCIIが混入
    # しうる)・鍵(このファイル内の定数)のどちらかがASCII外だと、その前提が
    # 崩れて復元側で文字化けや例外につながるため、ここで機械的に弾く。
    for ch in plain:
        if ord(ch) >= 0x80:
            raise BuildError(
                "obfuscate_secret: 平文がASCII外の文字を含んでいます"
                f"(環境変数 {AZURE_EMBED_KEY_ENV} を確認してください): {ch!r}"
            )
    for ch in _OBF_KEY:
        if ord(ch) >= 0x80:
            raise BuildError(
                f"obfuscate_secret: _OBF_KEY がASCII外の文字を含んでいます: {ch!r}"
            )
    xored = bytes(
        (ord(c) ^ ord(_OBF_KEY[i % len(_OBF_KEY)])) & 0xFF
        for i, c in enumerate(plain)
    )
    return _OBF_PREFIX + xored.hex()


# ---------------------------------------------------------------------------
# --zip (R9): 大規模配布用のワンコマンド梱包。完成した .xlsm と dist/Ghostscript
# (雑務担当が取得・常置済みの gswin32c.exe + gsdll32.dll + LICENSE + README)を
# 1本のzipへまとめる。この梱包物はgitignore対象(dist/*_配布.zip): リポジトリ
# 常置はGhostscriptの生ファイルそのもので足りており、zipまで履歴に積むと
# 肥大するだけのため。既定ビルド(--prod/--dev、--zip無し)の挙動には一切影響しない。
#
# 2026-08-01(R12-9-5・憲章§5-3): 配布物には実機スモークテスト手順書を必ず
# 添える。従来は「zip+docs/45を別送」という運用依存の穴があり(大規模配布で
# 別送を忘れると手順書なしのファイルだけが現場に届く)、ここでzip自体に
# 同梱して運用ミスの発生余地を無くす。簡易README.txtも同梱し、zipを渡された
# だけの人でも「展開して・全部同じ場所に置いて・マクロを有効化する」の3点が
# 分かるようにする(CP932でエンコードできる文字のみ使用)。
# ---------------------------------------------------------------------------
_README_TEXT = (
    "マイ本棚AI 配布物 README\n"
    "==========================================\n"
    "\n"
    "【1. 開き方】\n"
    "  1. この zip ファイルの中身を、すべて同じフォルダへ展開(解凍)してください。\n"
    "  2. 展開してできた MyBookshelf.xlsm (または MyBookshelf_dev.xlsm) を\n"
    "     Excel で開いてください。\n"
    "\n"
    "【2. zip の中身はすべて展開してください】\n"
    "  同梱の Ghostscript フォルダは、画像だけのPDF(スキャンした資料など)を\n"
    "  読み取るための部品です。.xlsm と同じ場所に置かれていないと、\n"
    "  画像PDFの取り込みができません。1つも欠かさず、すべて同じフォルダへ\n"
    "  展開してください。\n"
    "\n"
    "【3. マクロを有効にしてください】\n"
    "  開いた直後、画面の上に黄色い帯で「セキュリティの警告」と出たら、\n"
    "  その中の「コンテンツの有効化」ボタンを押してください。\n"
    "  有効化しないと、案内画面が表示されるだけで実際の機能が使えません。\n"
    "\n"
    "詳しい手順は、同梱の docs\\45_実機スモークテスト手順.md を参照してください。\n"
)


def build_dist_zip(xlsm_path: str, dist_dir: str, is_publisher: bool, is_dev: bool,
                    root: str) -> str:
    gs_dir = os.path.join(dist_dir, "Ghostscript")
    if not os.path.isdir(gs_dir):
        raise BuildError(f"--zip: dist/Ghostscript が見つかりません: {gs_dir}")

    # 2026-08-01(R12-9-4): ディレクトリの存在だけでなくOCRに必須の実体ファイル
    # を検査する。ビルド機のAV/EDRがexeを隔離した場合や部分チェックアウトで
    # 欠落していると、OCR不能なzipが無警告で完成してしまう(受領者側では
    # 「画像PDFが読めない」としてしか現れない)。
    for fn in ("gswin32c.exe", "gsdll32.dll"):
        fp = os.path.join(gs_dir, fn)
        if not os.path.isfile(fp) or os.path.getsize(fp) <= 0:
            raise BuildError(
                f"--zip: dist/Ghostscript/{fn} が見つからないか空です({fp})。"
                "OCR(画像PDF読み取り)が使えないzipになるため、配布前に"
                "Ghostscriptの再取得が必要です。"
            )

    # 2026-08-01(R12-9-1): --dev --zip が利用者向けと同名の配布zipを作り、
    # 既存の本番配布zipを無警告で上書きしていた(mock_llmビルドが全受領者へ
    # 渡る事故)。発行者/開発の2軸それぞれで出力名を分離する
    # (--out の既定ファイル名の4分岐と同じ考え方)。
    if is_publisher and is_dev:
        zip_name = "MyBookshelf_発行者用_dev_配布.zip"
    elif is_publisher:
        zip_name = "MyBookshelf_発行者用_配布.zip"
    elif is_dev:
        zip_name = "MyBookshelf_dev_配布.zip"
    else:
        zip_name = "MyBookshelf_配布.zip"
    zip_path = os.path.join(dist_dir, zip_name)
    xlsm_name = os.path.basename(xlsm_path)

    docs45_path = os.path.join(root, "docs", "45_実機スモークテスト手順.md")
    if not os.path.exists(docs45_path):
        raise BuildError(f"--zip: 同梱すべき手順書が見つかりません: {docs45_path}")
    try:
        readme_bytes = _README_TEXT.encode("cp932")
    except UnicodeEncodeError as e:
        raise BuildError(f"--zip: README.txt がCP932でエンコードできません: {e}")

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(xlsm_path, xlsm_name)
        zf.writestr("README.txt", readme_bytes)
        zf.write(docs45_path, "docs/45_実機スモークテスト手順.md")
        for walk_root, _dirs, files in os.walk(gs_dir):
            for fn in files:
                fp = os.path.join(walk_root, fn)
                arcname = os.path.join("Ghostscript", os.path.relpath(fp, gs_dir))
                zf.write(fp, arcname)

    return zip_path


# ---------------------------------------------------------------------------
# build_stamp: 「実機で今テストしているファイルが、本当に最新のソースから
# 作られたものか」を後から確定できるようにするための識別子(2026-07-21:
# 何度も往復した実機デバッグで、テスト対象が最新ビルドかどうか自体を疑わざるを
# 得ない場面があったことへの恒久対策)。ビルド日時+可能ならgitコミットの短縮
# ハッシュを config シートへ焼き込み、modLog.LogError が全err_log行に自動で
# 付記する(src/core/modLog.bas参照)。gitが無い/リポジトリ外でもビルドは
# 止めない(取得失敗時はコミット部分を省略するだけ)。
# ---------------------------------------------------------------------------
def compute_build_stamp() -> str:
    ts = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%d-%H%M%S")
    commit = ""
    try:
        commit = subprocess.check_output(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=SCRIPT_DIR, stderr=subprocess.DEVNULL, timeout=5,
        ).decode("ascii", errors="ignore").strip()
    except Exception:
        commit = ""
    return f"{ts}Z" + (f"+{commit}" if commit else "")


# ---------------------------------------------------------------------------
# config 既定値 (MASTER_SPEC §5 config キー台帳を完全反映。値・説明とも準拠)
# ---------------------------------------------------------------------------
def build_config_rows(mock_llm: bool, publish_key: str = ""):
    return [
        ("build_stamp", compute_build_stamp(),
         "このビルドの識別子(日時+gitコミット短縮ハッシュ)。err_logの全行に自動付記される。"
         "実機テストの結果が実際にどのビルドのものか後から特定するための識別子(書き換え不要)"),
        ("mock_llm", mock_llm,
         "TRUE=社内AIリボンを呼ばずダミー応答で動作確認(リボン無しでも取込→検索→回答が一通り動く)。本番はFALSE"),
        ("ribbon_addin_name", "リボンちゃん", "社内AIリボンのアドイン名(Application.AddInsからの検出に使用。名称が変わったときだけ書き換える)"),
        ("limit_check", True, "TRUE=起動時にリボンのLimitCheck(期限・利用同意チェック)を行う。FALSEで無効化(エスケープハッチ)"),
        ("recommended_model", "gpt-5.5", "精査モード(🔍しっかり調べる)で使うモデル名"),
        ("quick_model", "gpt-5.5", "即答モード(⚡すぐ聞く)で使うモデル名(将来 gpt-5.4-nano 等に差し替え可)"),
        ("quick_effort", "low", "即答モードの reasoning_effort"),
        ("quick_verbosity", "low", "即答モードの verbosity"),
        ("deep_draft_effort", "medium", "精査モード・ドラフト生成の reasoning_effort"),
        ("deep_draft_verbosity", "high", "精査モード・ドラフト生成の verbosity"),
        ("deep_verify_effort", "high", "精査モード・検証の reasoning_effort"),
        ("deep_verify_verbosity", "medium", "精査モード・検証の verbosity"),
        ("reasoning_tuning", True, "TRUEでeffort/verbosityを指定。FALSEにすると空送信(古いモデル互換用)"),
        ("llm_wait_sec", 1200, "ChatGPT() 呼び出しのWait秒数"),
        ("topk_quick", 6, "即答モードでLLMに渡す上位ヒット件数"),
        ("topk_deep", 12, "精査モードでLLMに渡す上位ヒット件数"),
        ("max_context_chars", 40000, "プロンプトに載せる本文合計の文字数上限"),
        ("answer_language", "日本語", "回答言語(プロンプトに指定を挿入)"),
        ("embed_dim", 768, "埋め込みベクトルの保存次元数(Plan B: 1536取得→先頭768切詰め+再正規化。パック互換検査にも使用)"),
        ("embed_sleep_ms", 0, "埋め込みAPI呼び出し間のスロットリング(ミリ秒)。0=待たない(取込を速くする)。レート制限が出る場合のみ50〜150へ"),
        ("vector_precision", "d6", "ベクトル保存精度: full=フル精度 / d6=小数6桁丸め(Plan B推奨。サイズ約-42%)"),
        ("embed_transport", "ribbon",
         "埋め込みの通信経路: ribbon=AIリボン経由 / direct=Azure APIへバッチ直接送信。"
         "2026-07-28: 既定を ribbon にした。本番ビルドは azure_embed_key を焼き込めない"
         "(配布=キー配布になるため)ので、direct のままだと毎回 ribbon へフォールバックし、"
         "err_log に E0203 が積み上がって本当の障害が埋もれる。"
         "direct は、キーを自分で入れる管理端末で明示的に切り替えて使う"),
        ("embed_batch_size", 128, "direct時に1リクエストへまとめるチャンク数"),
        ("azure_http_timeout_ms", 60000, "direct埋め込みのHTTPタイムアウト(ms)。NW瞬断時の無限フリーズ防止。resolve/connectは内部で短めに固定"),
        ("azure_embed_url", os.environ.get(AZURE_EMBED_URL_ENV, ""),
         "direct時の埋め込みエンドポイント(URL全体)。ビルド時の環境変数 " + AZURE_EMBED_URL_ENV + " から注入(未設定なら空欄=ribbonへ自動フォールバック)"),
        ("azure_embed_key", obfuscate_secret(os.environ.get(AZURE_EMBED_KEY_ENV, "")),
         "direct時のAPIキー(注意: ブック配布=キー配布になる)。ビルド時の環境変数 " + AZURE_EMBED_KEY_ENV + " から注入し、configシートには平文で置かず軽い難読化(OBF1:接頭辞)を施す(modUtil.DeobfuscateSecretで復元。VBAプロジェクトへアクセスできる人には無意味な軽量対策)"),
        ("chunk_mode", "structure", "チャンク化方式: legacy=700字機械分割 / structure=見出し・条文の構造認識(推奨)"),
        ("chunk_target_chars", 700, "チャンクの目安文字数"),
        ("chunk_overlap_chars", 150, "チャンクのオーバーラップ文字数"),
        ("chunk_max_chars", 1800, "条文を分割せず1チャンクに収める上限(構造認識時)"),
        ("embed_prefix_breadcrumb", True, "TRUE=チャンク先頭に【資料名>章>条】を前置して文脈付きで保存・検索する"),
        ("retrieve_mode", "multi", "検索方式: single=単段 / multi=多段RAG(拡張→マルチクエリ→再ランク)"),
        ("expand_enabled", True, "TRUE=質問をAIで検索用に拡張(独立質問化+サブクエリ+仮回答)"),
        ("expand_subqueries", 3, "クエリ拡張で生成するサブクエリ本数"),
        ("expand_model", "", "拡張段のモデル(空=quick_modelを使用)"),
        ("expand_effort", "low", "拡張段のreasoning_effort"),
        ("expand_verbosity", "low", "拡張段のverbosity"),
        ("multi_candidates", 40, "マルチクエリ検索の候補プール上限"),
        ("rerank_enabled", True, "TRUE=候補チャンクをAIで再ランクしてから回答生成する"),
        ("quick_expand", False, "TRUE=「すぐ聞く」でも質問拡張を行う(AI呼び出しが1回増え数秒遅くなる)"),
        ("quick_rerank", False, "TRUE=「すぐ聞く」でも再ランクを行う(AI呼び出しが1回増え数秒遅くなる)"),
        ("topk_thorough", 16, "「入念に調べる」でLLMに渡す上位ヒット件数。実測で件数増は効果が薄いため控えめ"),
        ("thorough_subqueries", 6, "「入念に調べる」の質問拡張で作るサブクエリ数。角度の数が精度に効く(実測 R@10 90%→97%)"),
        ("rerank_model", "", "再ランク段のモデル(空=quick_modelを使用)"),
        ("rerank_effort", "low", "再ランク段のreasoning_effort"),
        ("rerank_verbosity", "low", "再ランク段のverbosity"),
        ("answer_tags", True, "TRUE=回答を<thinking>/<answer>構造で生成し<answer>のみ表示"),
        ("strict_grounding", True, "TRUE=資料のみ・出典必須・『資料からは判断できません』を強制"),
        ("quick_expand_light", True, "TRUE=すぐ聞くモードでは拡張を軽量化(速度優先)"),
        ("nexus_ui", True, "TRUE=起動時にNexus Agent(SPA風チャットUI)を表示する"),
        ("nexus_share_path", "\\\\pgiofs01\\Nexus_Share\\", "P2Pナレッジ共有フォルダ(Phase 4。ダミーパス・要書き換え)"),
        ("exp_question", 5, "ゲーミフィケーション: 質問1回で得るEXP"),
        ("exp_register", 20, "ゲーミフィケーション: ナレッジ登録1件で得るEXP"),
        ("exp_thumbup", 10, "ゲーミフィケーション: 🟢自己解決(役立った)1回で得るEXP"),
        ("exp_pack_share", 30, "ゲーミフィケーション: パック共有(出力)1回で得るEXP"),
        ("exp_feedback", 5, "ゲーミフィケーション: ご意見箱(フィードバック/バグ報告)1回で得るEXP(1日1回まで)"),
        ("exp_level_divisor", 100, "ゲーミフィケーション: レベル計算の除数。Lv=Int(√(EXP/除数))+1"),
        ("user_department", "", "分析用: あなたの部署名(分析CSVの部署比較フラグ列に入る。空なら未設定)"),
        ("noise_global_threshold", 2, "ナレッジ自浄: 異なるN人からの⚠️ノイズ報告(P2P集計)でその資料を全ユーザーの検索から組織的除外する閾値"),
        ("admin_users", "", "組織的除外を解除できる管理者ADユーザー名(カンマ区切り)。空なら誰も解除不可"),
        ("shelf_max_chunks", 20500, "本棚のチャンク数上限(Plan B)。大きくするほど資料が入るがサイズと検索時間が増える"),
        ("binary_rag", False, "狂気案Lv.1: TRUE=大規模時にバイナリ量子化(XORハミング)で候補を高速粗選別してからFloatコサインで再ランク。小規模(binary_rag_min未満)では自動的に厳密Floatのまま=挙動不変。FALSEでも binary_rag_auto=TRUE なら大規模時は自動で有効になる"),
        ("binary_rag_auto", True, "TRUE=binary_rag が FALSE でも、チャンク数が binary_rag_min 以上なら粗選別を自動で有効にする(大規模な本棚ほど検索が重くなるため既定TRUE)。粗選別を完全に止めたいときだけ FALSE にする"),
        ("binary_rag_min", 5000, "バイナリ量子化ハイブリッドが作動する最小チャンク数(これ未満は従来どおり全件Floatスキャン)"),
        ("binary_rag_prefilter", 200, "バイナリ粗選別で残す候補数(topKより十分大きくFloat再ランクの精度を担保)"),
        ("binary_rag_debug", False, "狂気案Lv.1のデモ用: TRUE=ハイブリッド検索の所要ms(バイナリ選別/Float再ランク)をToastで画面表示(爆速証明)。Debug.Printには常時出力"),
        ("shelf_folder", "", "自動同期する本棚フォルダのパス(空なら未設定)"),
        ("sync_interval_min", 0, "自動同期の間隔(分)。0でOFF"),
        ("sync_on_open", True, "TRUE=起動時に本棚フォルダと差分同期する"),
        ("enrich_mode", "off", "off/light/full: バッチ富化(要約・キーワード付与)の強さ"),
        ("max_pages_per_file", 300, "1ファイルあたりの抽出ページ数上限(超過分は打ち切りpartial扱い)"),
        ("followup_max_pairs", 3, "『続けて質問』で引き継ぐ会話履歴の最大ペア数。0以下で機能無効"),
        ("word_export_effort", "medium", "『Wordで開く』の文書整形に使う reasoning_effort"),
        ("word_export_verbosity", "medium", "『Wordで開く』の文書整形に使う verbosity"),
        ("feature_tts", False, "opt機能フラグ: 読み上げ(AIリボン非公開機能のため提供不可。既定FALSEのまま変更しない)"),
        ("feature_vision", True, "opt機能フラグ: 画像読み取り・スクショ取込(公式仕様確定済み。問題があればFALSEで無効化)"),
        ("feature_markdown", True, "opt機能フラグ: Markdown表示・Wordで開く(公式仕様確定済み。問題があればFALSEで無効化)"),
        ("feature_diffdoc", True, "opt機能フラグ: 約款差分比較(確認済み関数のみ使用のため既定TRUE)"),
        ("ghostscript_path", "",
         "画像PDFの読み取り(OCR)に使う gswin32c.exe のフルパス。"
         "空欄のときは、このファイルと同じ場所にある Ghostscript フォルダを自動で探す。"
         "詳しい置き方は docs/43_画像PDFのOCR取込設定.md を参照"),
        ("ghostscript_search_dirs", "",
         "ghostscript_path・同梱(Ghostscriptフォルダ)のどちらでも見つからないときに"
         "追加で探すフォルダ(セミコロン区切り、複数可)。各フォルダの直下と"
         "bin\\直下の両方に gswin32c.exe が無いか探す。IT部門が社内の標準配置先を"
         "焼き込んでおく用途を想定した項目で、通常の利用者は空欄のままでよい"
         "(同梱のGhostscriptで動く)。詳細は docs/30_運用保守ガイド.md を参照"),
        ("vision_pdf_max_pages", 20,
         "画像PDFを何ページ目まで読み取るか。1ページごとにAIを1回呼ぶので、"
         "大きくすると時間もコストも比例して増える(超過分は打ち切り=partial表示)"),
        ("vision_pdf_dpi", 150,
         "画像PDFをページ画像にするときの解像度。150で十分読める。"
         "細かい文字が読めないときだけ300へ(処理時間は約2倍になる)"),
        ("vision_pdf_timeout_sec", 120,
         "Ghostscriptの処理が【まったく進まなくなってから】何秒で見切るか。"
         "ページが進んでいる限りは待ち続けるので、分厚い資料でも途中で"
         "打ち切られない。ページ数の多い資料で頻発するときだけ大きくする"),
        ("gs_abs_timeout_sec", 1200,
         "Ghostscriptの処理を待つ絶対上限(秒)。進んでいても必ずここで打ち切り、"
         "壊れたPDF1つでExcelが何十分も戻ってこない事態を防ぐ。"
         "数百ページの資料を入れるときだけ大きくする"),
        ("pack_author", "", "パック作成者名(空の場合は初回起動時に入力を促す)"),
        ("debug_mode", False, "TRUE=ゲートウェイのプロンプト/応答を診断用にログへ残す"),
        ("chat_log_enabled", True, "TRUE=チャット履歴シートに質問と回答を記録する(最新100件・古い順に自動削除)"),
        ("low_hit_warn_score", 0.3, "検索ヒットの最高スコアがこの値未満のとき回答に⚠️関連薄い警告を付ける(0で無効)"),
        ("active_channel", "",
         "いま接続している部門の公式ナレッジ(1つだけ)。ナレッジ画面の「部門チャンネル」で"
         "切り替える。切り替えると前の部門の内容は本棚から外れる(マイ本棚の資料は残る)"),
        ("unsubscribed_channels", "",
         "購読【しない】部門チャンネル(カンマ区切り)。既定は空=全チャンネルを自動購読する。"
         "『今日は商品、明日はシステム』という実際の使われ方に、購読操作を挟ませないため"),
        ("publish_key", publish_key,
         "正典を発行できる端末に入れる合言葉。空欄だと発行ボタン自体が出ない。"
         "各部門の発行担当者にだけ伝える(パスワードではなく誤操作防止の関所)。"
         "ビルド時に --publisher を付けた発行者用ブックにだけ焼き込まれる"),
        ("startup_jitter_ms", 3000,
         "起動時に共有フォルダを見に行くまでのランダム待機の上限(ミリ秒)。"
         "始業時に全員が同時アクセスしてファイルサーバが詰まるのを避ける。0で無効"),
        ("allowed_domain", "",
         "利用を許可するWindowsドメイン(カンマ区切り)。空欄=チェックしない。"
         "設定すると、許可外の端末では知識を消去して内容を表示しない(PC紛失対策)"),
        ("knowledge_expire_days", 0,
         "共有フォルダに最後に到達できた日から何日で知識を自動消去するか。"
         "0で無効。出荷既定は0(PoC中は失効させない)。本格展開時に30以上を設定する。"
         "7日前から警告あり。持ち出された端末が永久に中身を保持しないための保険。"
         "一度も共有へ到達したことが無い端末は、この設定に関わらず決して消さない"),
        ("telemetry_enabled", True,
         "TRUE=利用状況(質問回数・画面別表示回数・解決率)を終了時に共有フォルダへ送る。"
         "質問文は先頭40字までしか送らない。FALSEで送信停止"),
        ("insight_share_enabled", True,
         "回答が見つからなかった質問を部門の共有フォルダへ送り、詳しい人が答える仕組み。"
         "FALSEで送信を停止(氏名・部署・質問文が共有されなくなります)"),
        ("chunk_limit", 20000,
         "本棚に置けるチャンクの上限。Hubの「本棚の使用量」はこの値に対する割合。"
         "8割を超えると警告し、使っていない部門チャンネルの購読解除を促す"),
        ("log_max_rows", 2000,
         "err_log / usage_log に残す行数の上限。超えた分は古い方から自動的に消える。"
         "このブックは開くたび自己保存するため、ログが無限に伸びると毎起動の保存が重くなる。0以下でローテーション無効"),
        ("domain_wipe_after_n_boots", 3,
         "allowed_domain と一致しない状態が何回続いたら本棚を消すか。1回目は表示ブロックのみ。"
         "VPN未接続や一時的なプロファイル不整合でも不一致になり得るため、即消さない。0以下なら決して消さない(表示ブロックのみ)"),
        ("feedback_mail_to", "",
         "ご意見・不具合報告のメール送信先。共有フォルダへ書けない環境の逃がし先として使う。"
         "空欄のときはメール経路そのものを出さない(本文はクリップボードへ入る)。"
         "個人アドレスをソースへ直書きすると担当交代・異動のたびに再ビルドが要るためconfig化した"),
        ("thanks_gc_days", 60,
         "共有フォルダに置いた感謝状(thanks\\thx_*.txt)と共有知(insight\\)を何日で片付けるか。"
         "放置すると共有フォルダに無限に溜まり、毎起動の列挙が重くなる。0以下でGC無効。"
         "30日だと育休・長期出張・長期休職の人宛の感謝状が届く前に消えるため60日"),
        ("exp_correction", 20,
         "「違う」「微妙」を押した上で正しい内容を書いてくれた人に付与するEXP。"
         "資料登録(20)と同格。手間に見合う対価が無いと誰も書かない"),
        ("confidence_score_x100", 55,
         "回答の信頼度バッジのスコア閾値×100。この値以上のヒットが2件以上で「強く一致」、"
         "1件で「部分的に一致」、それ未満は「根拠なし」と表示する"),
        ("holidays", (
            "2026-01-01,2026-01-12,2026-02-11,2026-02-23,2026-03-20,2026-04-29,"
            "2026-05-03,2026-05-04,2026-05-05,2026-05-06,2026-07-20,2026-08-11,"
            "2026-09-21,2026-09-22,2026-09-23,2026-10-12,2026-11-03,2026-11-23,"
            "2026-12-29,2026-12-30,2026-12-31,"
            "2027-01-01,2027-01-11,2027-02-11,2027-02-23,2027-03-21,2027-03-22,"
            "2027-04-29,2027-05-03,2027-05-04,2027-05-05,2027-07-19,2027-08-11,"
            "2027-09-20,2027-09-23,2027-10-11,2027-11-03,2027-11-23"
         ), "連続ログイン日数で休みとして数える日(yyyy-mm-dd をカンマ区切り)。土日は自動で除外されるので祝日・年末年始だけ書けばよい"),
        ("ambiguous_max_chars", 10, "この文字数以下の質問だけを『曖昧かも』の判定対象にする(長い質問は常にそのままAIへ)"),
        ("ambiguous_score_x100", 60, "曖昧判定のスコア閾値×100。全ヒットの最高スコアがこの値未満なら聞き返す。0で機能OFF"),
    ]


# ---------------------------------------------------------------------------
# シート生成
# ---------------------------------------------------------------------------
def _clear_sheet(ws, max_row=60, max_col=8):
    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for c in row:
            c.value = None


def _banner(ws, text):
    ws["A1"] = text
    ws.merge_cells("A1:B1")
    ws["A1"].fill = PatternFill("solid", fgColor="3C5AA0")
    ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32


def _make_macro_guard(wb):
    """はじめにお読みください シート: マクロ無効ガード(軽量版)。
    マクロが無効なまま開かれると自己インストーラ(Workbook_Open)が走らず
    Bootも実行されないため、他の画面は素のプレースホルダーのまま(=中途半端な
    画面)になる。それを避けるため、このシートを常に先頭・アクティブ・visible
    にしてビルドする。マクロが有効化されればBoot側がこのシートを隠す
    (そちらの実装はこのビルドスクリプトの管轄外・ここでは一切hideしない)。
    index=0 で作成することで、後続の create_sheet(ホーム/マイ本棚/…)が
    末尾に追加されても本シートは常に最左タブのまま残る。"""
    ws = wb.create_sheet(GUARD_SHEET_NAME, 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 100
    for col in ("B", "C", "D", "E", "F"):
        ws.column_dimensions[col].width = 14

    bg = PatternFill("solid", fgColor="FFF9E6")   # 明るい中立背景
    ink = "1F2933"                                 # 濃色文字(テーマ非依存)

    entries = [
        (2, "⚡ Nexus Agent", Font(bold=True, size=22, color=ink), 40),
        (4,
         "このファイルを使うには、上の黄色いバーの［コンテンツの有効化］ボタンを押して、"
         "マクロ(コンテンツ)を有効にしてください。",
         Font(size=13, color=ink), 60),
        (6, "有効化すると、この案内は自動的に消え、AIチャット画面が表示されます。",
         Font(size=12, color="3E4C59"), 34),
        (8, "※ 有効化しても画面が変わらない場合は、ファイルを一度閉じて開き直してください。",
         Font(size=11, italic=True, color="52606D"), 34),
    ]
    for row, text, font, height in entries:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        cell = ws.cell(row=row, column=1, value=text)
        cell.font = font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
        ws.row_dimensions[row].height = height

    # 背景を軽く塗って独立した案内ページに見えるようにする(装飾のみ)
    for r in range(1, 10):
        for c in range(1, 7):
            ws.cell(row=r, column=c).fill = bg

    ws.sheet_state = "visible"
    return ws


def _make_howto(wb):
    """使い方シート: マクロ無効でも読める唯一の救済ページ。

    実機要望(2026-07-26)「文字だらけで表もなく、改行も適切でなく、認知負荷が
    高くて読む気が失せる」への全面作り直し。Shapeは使わずセルだけで、
    (1) 番号つきの手順、(2) 1行1項目の早見表、(3) 用語のミニ辞書 という
    3つの構造に分解した。1セルに長文を詰め込まず、手順は1ステップ1行にする。
    画面名・ボタン名は 2026-07-26 の3画面リデザイン(Hub / チャット / ナレッジ)に
    合わせてある(この関数を編集するときは実装の文字列と必ず突き合わせること)。
    """
    ws = wb["Sheet1"]
    ws.title = "使い方"
    _clear_sheet(ws)
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 78

    _banner(ws, f"{APP_TITLE} — 使い方")

    head_font = Font(bold=True, size=12, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="1F4E78")
    key_font = Font(bold=True, size=10, color="1F4E78")
    body_font = Font(size=10)
    step_font = Font(bold=True, size=10, color="FFFFFF")
    step_fill = PatternFill("solid", fgColor="7F9DB9")
    warn_fill = PatternFill("solid", fgColor="FFF4D6")
    zebra_fill = PatternFill("solid", fgColor="F4F6FB")
    thin = Side(style="thin", color="D9E1EC")

    row = [3]

    def section(title, fill="1F4E78"):
        if row[0] > 3:
            row[0] += 1          # 直前ブロックとの間に必ず1行空ける
        r = row[0]
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        c = ws.cell(row=r, column=1, value=title)
        c.font = head_font
        c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[r].height = 24
        row[0] = r + 2

    def step(n, text, warn=False):
        r = row[0]
        c0 = ws.cell(row=r, column=1, value=n)
        c0.font = step_font
        c0.fill = step_fill
        c0.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=2, value=text)
        c.font = body_font
        c.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        if warn:
            c.fill = warn_fill
        ws.row_dimensions[r].height = max(20, 15 * (text.count("\n") + 1) + 6)
        row[0] = r + 1

    def kv(k, v, i=0):
        r = row[0]
        ck = ws.cell(row=r, column=2, value=k)
        ck.font = key_font
        ck.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        cv = ws.cell(row=r, column=3, value=v)
        cv.font = body_font
        cv.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        if i % 2 == 1:
            ck.fill = zebra_fill
            cv.fill = zebra_fill
        for c in (ck, cv):
            c.border = Border(bottom=thin)
        ws.row_dimensions[r].height = max(20, 15 * (v.count("\n") + 1) + 5)
        row[0] = r + 1

    def note(text):
        r = row[0]
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=2, value=text)
        c.font = Font(size=9, italic=True, color="6B7280")
        c.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        ws.row_dimensions[r].height = max(18, 14 * (text.count("\n") + 1) + 5)
        row[0] = r + 2

    # ---- これは何か -------------------------------------------------------
    section("これは何ですか")
    kv("ひとことで言うと", "自分で入れた資料にAIが答えてくれる、社内版のNotebookLMです。", 0)
    kv("普通のAIとの違い", "答えの根拠になった資料名とページが必ず一緒に出ます。原文もその場で開けます。", 1)
    kv("入れられる資料", "PDF / Word / Excel / テキスト / 画面のスクリーンショット", 0)
    note("※ 契約者名・電話番号などの個人情報を含む資料は入れないでください。")

    # ---- 最初の設定 -------------------------------------------------------
    section("はじめに 1回だけやること(ここで9割の人がつまずきます)", "B45F06")
    step("1", "このファイルをExcelで開く。")
    step("2", "画面の上に黄色い帯で「セキュリティの警告」と出たら、その中の\n「コンテンツの有効化」ボタンを押す。", warn=True)
    step("3", "[ファイル] → [オプション] → [トラストセンター] →\n[トラストセンターの設定] → [マクロの設定] と進む。")
    step("4", "「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」に\nチェックを入れて [OK]。", warn=True)
    step("5", "Excelをいったん全部閉じて、もう一度このファイルを開く。")
    step("6", "数秒待つと画面が自動で組み上がります。これで準備完了です。")
    note("英語で「VBA Project trust required」と出たときは、手順3〜4がまだ終わっていない合図です。\nこのファイルは初回に自分で画面を組み立てる作りなので、この許可が必要です。")

    # ---- 画面の見取り図 ---------------------------------------------------
    section("画面は5つです")
    kv("🏠 Hub(拠点)", "最初に出る画面。自分の記録と、他の画面への入口が並んでいます。", 0)
    kv("💬 チャット", "AIに質問する画面。ここだけで会話が完結します。", 1)
    kv("📚 マイ本棚", "資料を入れる・探す・共有する画面。上のピルで\n「ナレッジ倉庫(カード表示)」「一覧表」「みんなの解決事例」の3モードを切り替えます。", 0)
    kv("📊 ダッシュボード", "節約時間・レベル・獲得バッジなど、自分の利用状況を見る画面。", 1)
    kv("🩺 診断", "動作がおかしいときに状態を確認する画面。ヘルプ「❓」からも開けます。", 0)
    note("画面の行き来はすべてボタンで行います。左上の「← Hub」でいつでも拠点に戻れます。")

    # ---- 質問する ---------------------------------------------------------
    section("質問してみる")
    step("1", "Hubの「💬 チャットで質問する」を押す。")
    step("2", "上の白い入力らんをクリックして、知りたいことを文章で書く。\n  例) 契約者が亡くなったときの手続きを教えて")
    step("3", "「💬 質問する」を押す(Ctrlキー+Enterでも送れます)。")
    step("4", "答えの下に出る資料名のボタンを押して、元の文章を必ず確認する。")
    step("5", "役に立ったら「✅ 解決した」を押す。記録が貯まり、資料を作った人にも届きます。")

    # ---- ボタン早見表 -----------------------------------------------------
    section("ボタン早見表(チャット画面)")
    kv("⚡ すぐ聞く", "10〜20秒。ふだんはこちら。", 0)
    kv("🔍 しっかり調べる", "40〜60秒。念入りに調べて答えを検証してから返します。", 1)
    kv("🏢 社内ナレッジ検索", "自分の本棚の資料だけを見て答えます(出典つき)。", 0)
    kv("🌐 一般アシスタント", "本棚を見ずに一般知識で答えます。文章の下書きなどに。", 1)
    kv("📎 (クリップ)", "画面のスクリーンショットを貼って、その中身について質問できます。", 0)
    kv("✅ 解決した", "自己解決として記録。節約時間と経験値が増え、資料を書いた人にも感謝が届きます。", 1)
    kv("🤔 微妙", "参考にはなったが不十分だったときに記録します。", 0)
    kv("❌ 違う", "正しい内容を入力すると、それを覚えて次から反映します。", 1)
    kv("🔍 深掘り", "直前の会話をふまえて、続けて質問します。", 0)
    kv("📋 コピー / 📄 Word", "答えをコピー、またはWord文書として書き出します。", 1)
    kv("🗑 クリア", "会話を消して最初からやり直します(資料は消えません)。", 0)
    kv("🚪 (ドア)", "保存してこのファイルを閉じます。", 1)

    # ---- 資料を入れる -----------------------------------------------------
    section("資料を入れる")
    step("1", "Hubの「📚 ナレッジ倉庫」を押す。")
    step("2", "上のツールバーの「📁 追加」でファイルを選ぶ。")
    step("3", "状態が ⏳(変換中) から ✅(完了) に変わったら質問できます。")
    note("「📂 フォルダ」で本棚フォルダを決めておくと、そこに置いたファイルは自動で取り込まれます\n(フォルダから消せば本棚からも消えます)。")

    # ---- 状態の記号 -------------------------------------------------------
    section("状態の記号")
    kv("✅", "取り込み完了。質問に使えます。", 0)
    kv("⏳", "変換中。しばらく待ってください。", 1)
    kv("⚠️", "一部うまく読めませんでした。ページ番号がメモ欄に出ます。", 0)
    kv("🖼", "画像だけのPDFです。文字が入っていないため読み取れません。", 1)
    kv("🕒", "元のファイルが見つかりません(移動・削除された可能性)。", 0)

    # ---- 困ったとき -------------------------------------------------------
    section("困ったときは", "9C1F1F")
    kv("画面が崩れた・ボタンが消えた", "Hub右上の「🔄」を押すと画面を描き直します。", 0)
    kv("答えが途中で止まる", "いったん保存して閉じ、開き直してからもう一度お試しください。", 1)
    kv("それでも直らない", "ヘルプ「❓」→「診断」の画面をスクリーンショットで撮って管理者へ送ってください。", 0)
    kv("操作を思い出したい", "ヘルプ「❓」→「ツアーをもう一度見る」で、最初の案内を再表示できます。", 1)
    note("自己判断で設定を変える必要はありません。まずスクリーンショットを送ってください。")

    # ---- 部門チャンネル ---------------------------------------------------
    section("部門の知識が自動で届きます")
    # 2026-07-28(レビュー I-4): 実装に無い動作を書かない。
    # 自動購読(開くだけで届く)は未実装で、実際は Hub の「更新があります」を
    # 押したときにまとめて取り込む。期待値がずれたまま PoC に入ると
    # 「届いていない」という問い合わせになる。
    kv("部門チャンネルとは", "商品部・システム部・人事部などが公開している『正典』です。\nHubに出る案内を1回押すだけで、全部門ぶんがまとめて入ります。", 0)
    kv("何が良いのか", "自分で資料を集めなくても、開いた初日から答えが返ります。\nポータルを探し回る必要がなくなります。", 1)
    kv("更新されたら", "Hubに「更新があります」と出ます。押すだけで最新版に入れ替わります。\n古い内容は入れ替えのときに消えるので、古い条文で回答されることはありません。", 0)
    kv("自分の資料は?", "消えません。部門の正典と、あなたが入れた資料の両方から答えます。", 1)
    # 購読解除のUIは未実装(config unsubscribed_channels の手編集のみ)。
    # 「Hubから外せます」は嘘になるので、実際にできることを書く。
    note("本棚の使用量が8割を超えたらHubが知らせます。減らしたいときは管理者にご相談ください。")

    # ---- データの取り扱い -------------------------------------------------
    section("このツールが記録していること", "5B6B7B")
    kv("この端末に残るもの", "あなたの質問と回答の履歴(最新100件)。この端末の中だけです。", 0)
    kv("部内で共有されるもの", "「解決した」を押した質問と答え / 答えが見つからなかった質問 /\n利用回数・解決率などの集計値。改善のために使います。", 1)
    kv("共有されないもの", "質問の全文は共有されません(先頭40文字までの要約のみ)。", 0)
    kv("匿名で意見を送る", "Hub右上の 📮 から、名前を残さずに感想・不満を送れます。", 1)
    note("記録を止めたい場合は config シートの telemetry_enabled を FALSE にしてください。")

    # ---- このシートについて -----------------------------------------------
    section("このシートについて")
    note("このシートだけは、マクロが無効な状態でも読めるようにしてあります\n(マクロ有効化の案内は、ここでしか出せないためです)。\nマクロを有効にして開き直すと、実際の操作画面が使えるようになります。")

    ws.sheet_properties.tabColor = TAB_COLORS["使い方"]
    ws.sheet_view.showGridLines = False
    return ws


def _make_placeholder(wb, name, note):
    """ホーム/マイ本棚/ダッシュボードの初期プレースホルダー。
    実際の画面は modUIMain/modUIShelf/modUIDashboard の EnsureLayout が
    起動時(Boot)に Shapes で冪等再構築する。ビルド時点ではこの案内文のみ。"""
    ws = wb.create_sheet(name)
    ws.column_dimensions["A"].width = 90
    ws["A1"] = f"{APP_TITLE} — {name}"
    ws["A1"].fill = PatternFill("solid", fgColor="3C5AA0")
    ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32
    ws["A3"] = note
    ws["A3"].font = Font(size=11)
    ws["A3"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.sheet_properties.tabColor = TAB_COLORS[name]
    return ws



def _make_seed_sheets(wb, seed_path: str | None):
    """初期ナレッジ(シードパック)を seed_* シートとして焼き込む。

    tools/make_seed_pack.py が作った pack形式(.xlsx) の3シートを、そのまま
    veryHidden で持ち込む。実行時に modSeed が my_knowledge / my_vectors へ
    写すだけなので、初回起動でチャンク分割もAPI呼び出しも起きない。

    pack_vectors が空のパックでも壊れない(embedded=0 で入り、次の同期で
    埋め込まれる)。ただし初回体験のためには、リボンのあるPCで一度
    ベクトル化したパックを指定するのが本来の運用。
    """
    from openpyxl import load_workbook as _lw

    specs = [
        ("seed_meta",    ["key", "value"],                 [28, 110]),
        ("seed_chunks",  ["chunk_id", "source", "page", "summary", "keywords", "full_text"],
                                                           [32, 28, 6, 44, 30, 90]),
        ("seed_vectors", ["chunk_id", "vector_csv"],        [32, 100]),
    ]
    if not seed_path or not os.path.exists(seed_path):
        for name, hdr, w in specs:
            _make_headers_only(wb, name, hdr, "veryHidden", widths=w)
        return 0, 0, 0

    src = _lw(seed_path, read_only=True, data_only=True)
    n_chunks = n_vecs = n_docs = 0
    docs = set()

    for name, hdr, widths in specs:
        ws = wb.create_sheet(name)
        ws.append(hdr)
        for c, wd in enumerate(widths, start=1):
            ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = wd

        pack_name = name.replace("seed_", "pack_")
        if pack_name not in src.sheetnames:
            ws.sheet_state = "veryHidden"
            continue

        srcws = src[pack_name]
        for r, row in enumerate(srcws.iter_rows(min_row=2, values_only=True)):
            if row is None or all(v is None for v in row):
                continue
            ws.append(list(row[:len(hdr)]))
            if name == "seed_chunks":
                n_chunks += 1
                if row[1]:
                    docs.add(str(row[1]))
            elif name == "seed_vectors":
                n_vecs += 1
        # full_text 等は必ず文字列として持たせる(数字だけの本文が数値化するのを防ぐ)
        if name == "seed_chunks":
            for row in ws.iter_rows(min_row=2):
                for idx in (2, 4, 5, 6):
                    row[idx - 1].number_format = "@"
        ws.sheet_state = "veryHidden"

    src.close()
    n_docs = len(docs)
    return n_docs, n_chunks, n_vecs


def _make_config(wb, mock_llm: bool, publish_key: str = ""):
    ws = wb.create_sheet("config")
    for c, h in enumerate(["key", "value", "description"], 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
    for i, (k, v, d) in enumerate(build_config_rows(mock_llm, publish_key), 2):
        ws.cell(row=i, column=1, value=k)
        ws.cell(row=i, column=2, value=v)
        ws.cell(row=i, column=3, value=d)
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 70
    ws.sheet_state = "hidden"
    return ws


def _make_headers_only(wb, name, headers, state, widths=None, text_cols=None):
    ws = wb.create_sheet(name)
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    # 数式インジェクション防御: 非信頼テキスト列(取込文書/ファイル名/入力由来)を
    # Excelのテキスト書式("@")へ固定し、先頭 =/+/-/@ が格納型数式として評価される
    # のを配布テンプレート段階で封じる(実行時のEnsureKnowledgeSheetガードと二重化)。
    if text_cols:
        for i in text_cols:
            ws.column_dimensions[get_column_letter(i)].number_format = "@"
    ws.sheet_state = state
    return ws


def _make_vba_src(wb, present_modules, root):
    """vba_src シート: 標準モジュール(*.bas)のソースを1行1モジュールで格納する。
    自己インストーラ(ThisWorkbookストリーム)がこのシートを読んで
    VBComponents.Add(1)でモジュールを注入する。
    注意: クラスモジュール(type=class, 例 ThisWorkbook.cls)は
    VBComponents.Add(1) では追加できない(標準モジュール専用API)ため対象外。"""
    ws = wb.create_sheet("vba_src")
    for c, h in enumerate(["module_name", "type", "source"], 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
    # E1(row=1, col=5)は実行時の一時領域として予約する。自己インストーラが
    # Application.OnTime で予約した modBoot.Boot の時刻(シリアル値)を置き、
    # modBoot.CancelPendingInstallerBoot がそれを読んで予約を取り消す。
    # 保存後に書かれるのでファイルには残らない(セッション内だけの状態)。
    # ここを別用途に使わないこと。
    ws.cell(row=1, column=5, value="").font = Font(bold=True)

    injected = []
    row = 2
    for m in present_modules:
        if m.get("type") == "class" or m.get("vba_src") is False:
            continue
        path = os.path.join(root, m["path"])
        with open(path, encoding="utf-8-sig") as fp:
            txt = fp.read()
        out_lines = []
        for line in txt.split("\n"):
            stripped = line.lstrip("﻿")
            if stripped.lstrip().startswith("Attribute "):
                continue
            out_lines.append(stripped)
        cleaned = _clean("\n".join(out_lines))

        if len(cleaned) > MODULE_CONTRACT_LIMIT:
            raise BuildError(
                f"{m['name']}.bas は{len(cleaned)}字でMASTER_SPEC §7の"
                f"モジュール契約上限({MODULE_CONTRACT_LIMIT}字)を超過しています")
        if len(cleaned) >= EXCEL_CELL_LIMIT:
            raise BuildError(
                f"{m['name']}.bas は{len(cleaned)}字でExcelセルの技術上限"
                f"({EXCEL_CELL_LIMIT}字)を超過しています")

        ws.cell(row=row, column=1, value=m["name"])
        ws.cell(row=row, column=2, value="std")
        ws.cell(row=row, column=3, value=cleaned)
        injected.append(m["name"])
        row += 1

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 80
    ws.sheet_state = "veryHidden"
    return injected


# ---------------------------------------------------------------------------
# 自己インストーラ (VBA/ThisWorkbook ストリームの中身)
# cp932でエンコードするが、文字は必ずASCIIのみに限定する(MASTER_SPEC §14手順6)。
# ロジックはV2 build_chatbot_v2.py の INSTALLER_SRC を踏襲し、起動先だけ
# modBoot.Boot に変更(V2実証済み機構をそのまま流用。値自体は元から
# modBoot.Boot だったので、mybookshelfでも変更なしで成立する)。
# ---------------------------------------------------------------------------
# 注: このVBAソースは vbaProject.bin の ThisWorkbook ストリームへ
# 「元のバイト長ぴったり」に圧縮して差し込む。長いコメントを入れると
# pad_to_exact が溢れてビルドが落ちるため、意図の説明はここに書く。
#
#   ・モジュール追加の失敗は1本ずつローカルに握る(2026-07-26)。
#     以前は Add の失敗で Done へ飛び、残り全モジュールの導入を黙って
#     打ち切ったうえ Boot も呼ばれなかった。実機では「なぜか特定の
#     モジュールだけ未定義」という再現しにくいコンパイルエラーになる。
#   ・CodeModule の既存行を消してから AddFromString するのは、VBEの
#     「変数の宣言を強制する」がONだと Option Explicit が自動挿入され、
#     ソース側の Option Explicit と重複してコンパイルエラーになるため。
_INSTALLER_SRC_TEXT = '''Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Workbook_Open()
  Install
End Sub
Public Sub Install()
  Dim p As Object, w As Worksheet, c As Object, e As Object
  Dim r As Long, n As String, s As String, l As Long, f As Long
  On Error GoTo Trust
  Set p = ThisWorkbook.VBProject
  On Error GoTo Done
  Set w = ThisWorkbook.Worksheets("vba_src")
  l = w.Cells(w.Rows.Count, 1).End(-4162).Row
  For r = 2 To l
    n = CStr(w.Cells(r, 1).Value)
    s = CStr(w.Cells(r, 3).Value)
    If LenB(n) > 0 Then
      On Error Resume Next
      Set e = Nothing: Set e = p.VBComponents(n)
      If Not e Is Nothing Then p.VBComponents.Remove e
      Set c = Nothing
      Set c = p.VBComponents.Add(1)
      If c Is Nothing Then
        f = f + 1
      Else
        c.Name = n
        If c.CodeModule.CountOfLines > 0 Then c.CodeModule.DeleteLines 1, c.CodeModule.CountOfLines
        If LenB(s) > 0 Then c.CodeModule.AddFromString s
      End If
      Err.Clear
      On Error GoTo Done
    End If
  Next r
  On Error Resume Next
  Application.Run "modBoot.RunFirstRunPromptEarly"
  Err.Clear
  ' f>0: half-injected. Do not Save.
  If f > 0 Then
    MsgBox "Setup incomplete. Please get a fresh copy of this file.", vbCritical
    Exit Sub
  End If
  ThisWorkbook.Save
  If Err.Number<>0 Then Application.Run "modLog.LogUsage","save_fail","",Err.Description
  Err.Clear
  ' Detach Boot(1004); E1=time. R12-3-8.
  Dim bt As Date
  bt = Now + TimeSerial(0, 0, 1)
  Application.OnTime bt, "'" & ThisWorkbook.Name & "'!modBoot.Boot"
  If Err.Number <> 0 Then
    Err.Clear
    Application.Run "modBoot.Boot"
  Else
    w.Cells(1, 5).Value = CDbl(bt)
    Err.Clear
  End If
  Exit Sub
Trust:
  MsgBox "VBA Project trust required. See howto sheet.", vbCritical
  Exit Sub
Done:
End Sub
'''


def build_installer_src() -> bytes:
    try:
        _INSTALLER_SRC_TEXT.encode("ascii")
    except UnicodeEncodeError as e:
        raise BuildError(f"INSTALLER_SRC must be ASCII-only (MASTER_SPEC §14): {e}")
    return _INSTALLER_SRC_TEXT.replace("\n", "\r\n").encode("cp932")


# ---------------------------------------------------------------------------
# 2026-08-01(R12-9-6): ovba.pad_to_exact の到達不能差分をBuildError化する。
# 空チャンクは5バイト/3バイトの組合せでしか埋められないため、
# diff ∈ {1,2,4,7} は原理的に到達不能で ovba.py 側は生の ValueError を出す
# (ovba.py はプロダクト固有ロジックを持たない自己完結モジュールという設計
# 方針のため、BuildError化・誘導文はここ=呼び出し側で行う)。
# ---------------------------------------------------------------------------
def _pad_to_exact_or_die(compressed: bytes, target: int, stream_name: str) -> bytes:
    try:
        return ovba.pad_to_exact(compressed, target)
    except ValueError as e:
        raise BuildError(
            f"{stream_name}ストリームのpaddingが目標バイト数に到達できません({e})。"
            "OVBA空チャンクは3バイト単位/5バイト単位の組合せでしか長さを埋められず、"
            "元サイズとの差分が1/2/4/7バイトのときは原理的に到達不能です。"
            "_INSTALLER_SRC_TEXT のコメント・変数名を1〜2バイト増減してから再実行してください。"
        )


# ---------------------------------------------------------------------------
# vbaProject.bin 外科パッチ (ThisWorkbookストリーム差し替え + dir MOFFSET=0)
# ovba.py の低レベル関数(圧縮/解凍/CFBReader/pad)だけを使い、
# インストーラ文字列などプロダクト固有の中身はここに閉じ込める。
# ---------------------------------------------------------------------------
def patch_installer(vba_bin: bytes, installer_src: bytes) -> bytes:
    skel = ovba.CFBReader(vba_bin)

    dir_dec = ovba.ovba_decompress(skel.read("dir"))
    needle = struct.pack("<HI", 0x0019, len("ThisWorkbook")) + b"ThisWorkbook"
    idx = dir_dec.find(needle)
    if idx < 0:
        raise BuildError("dir stream: ThisWorkbook MNAME record not found "
                          "(template_skeleton.xlsm may be incompatible)")
    i = idx
    found = False
    while i < len(dir_dec):
        rid = struct.unpack("<H", dir_dec[i:i + 2])[0]
        sz = struct.unpack("<I", dir_dec[i + 2:i + 6])[0]
        if rid == 0x0031:  # MOFFSET
            patched = bytearray(dir_dec)
            struct.pack_into("<I", patched, i + 6, 0)
            dir_dec = bytes(patched)
            found = True
            break
        i += 6 + sz
    if not found:
        raise BuildError("dir stream: ThisWorkbook MOFFSET record not found")

    orig_dir_size = skel.entries["dir"]["size"]
    orig_tw_size = skel.entries["ThisWorkbook"]["size"]
    new_dir = _pad_to_exact_or_die(ovba.ovba_compress(dir_dec), orig_dir_size, "dir")

    # ThisWorkbookストリームは「元と同じバイト数」でしか差し替えられない
    # (in-place外科パッチのため。ストリームを伸ばすとCFBのFATを組み直す
    # 必要があり、Excelが読めなくなる)。自己インストーラのVBAソースには
    # 圧縮後のサイズ上限があり、うっかり数行足すと超える。
    # 2026-07-28: 実際にOnTime予約時刻の記録を足した際に超過した。
    # ovba.pad_to_exact の生の ValueError では原因が読めないので、
    # ここで「何をすればよいか」まで含めて落とす。
    tw_compressed = ovba.ovba_compress(installer_src)
    if len(tw_compressed) > orig_tw_size:
        raise BuildError(
            f"自己インストーラ(ThisWorkbookストリーム)が圧縮後{len(tw_compressed)}バイトで、"
            f"差し替え可能な上限{orig_tw_size}バイトを{len(tw_compressed) - orig_tw_size}バイト超過しました。"
            "_INSTALLER_SRC_TEXT のコメント/変数名を削るか、処理そのものを "
            "modBoot 側(vba_srcから注入される標準モジュール。サイズ上限が緩い)へ移してください。"
        )
    new_tw = _pad_to_exact_or_die(tw_compressed, orig_tw_size, "ThisWorkbook")

    buf = io.BytesIO(vba_bin)
    ole = olefile.OleFileIO(buf, write_mode=True)
    ole.write_stream("VBA/dir", new_dir)
    ole.write_stream("VBA/ThisWorkbook", new_tw)
    ole.close()
    buf.seek(0)
    return buf.read()


# ---------------------------------------------------------------------------
# モジュール台帳 (modules.json) の検証
# ---------------------------------------------------------------------------
def load_manifest(path):
    with open(path, encoding="utf-8") as fp:
        data = json.load(fp)
    modules = data["modules"]
    # 2026-08-01(R12-9-3): name/path の重複検出。インストーラは行順に
    # Remove→Add するため、重複があると「後の行のソースが無言で勝つ」
    # (Stage6の集合照合も重複が両側に同数入るため通過してしまう)。
    seen_names = {}
    seen_paths = {}
    for m in modules:
        for key in ("name", "path", "role"):
            if key not in m:
                raise BuildError(f"modules.json: エントリに必須キー'{key}'がありません: {m}")
        if m["role"] not in ("core", "opt", "test"):
            raise BuildError(f"modules.json: {m['name']} の role が不正です: {m['role']}")
        nm, pth = m["name"], m["path"]
        if nm in seen_names:
            raise BuildError(
                f"modules.json: name '{nm}' が重複しています"
                f"(既存: {seen_names[nm]['path']} / 重複: {pth})。"
                "重複を解消してください(後の行が無言で勝つ事故を防ぐため)。"
            )
        seen_names[nm] = m
        if pth in seen_paths:
            raise BuildError(
                f"modules.json: path '{pth}' が重複しています"
                f"(name={seen_paths[pth]['name']!r} と name={nm!r} の2エントリ)。"
            )
        seen_paths[pth] = m
    return modules


# 2026-07-26 恒久対策: src/ 配下に実在するのに modules.json へ登録されていない
# .bas を検出してビルドを止める。
#   実際に起きた事故: modGuard/modTelemetry/modPublish を作ったのに台帳への
#   登録が漏れ、ビルドは「台帳と件数一致」で通り、実機で初めて
#   「変数が定義されていません(modGuard)」のコンパイルエラーになった。
#   台帳との突き合わせだけでは、台帳に無いものは永久に検出できない。
#   ファイルシステムを正として突き合わせる検査をここに置く。
# 除外したいファイル(意図的にビルドへ含めないもの)は EXCLUDE に明記する。
UNREGISTERED_EXCLUDE = {
    "src/opt/optTts.bas",   # 音声合成はリボン非公開で確定(裁定D4)。ソース保管のみ
}


def check_unregistered(modules, root):
    import glob as _glob
    registered = {m["path"].replace("\\", "/") for m in modules}
    found = set()
    for p in _glob.glob(os.path.join(root, "src", "**", "*.bas"), recursive=True):
        rel = os.path.relpath(p, root).replace("\\", "/")
        found.add(rel)
    orphans = sorted(found - registered - UNREGISTERED_EXCLUDE)
    if orphans:
        print("modules.json に未登録の .bas があります(実機でコンパイルエラーになります):",
              file=sys.stderr)
        for o in orphans:
            print(f"  {o}", file=sys.stderr)
        print("  → build/modules.json に追加するか、"
              "意図的に除外するなら UNREGISTERED_EXCLUDE へ明記してください。",
              file=sys.stderr)
        sys.exit(1)


def validate_modules(modules, root, allow_missing):
    check_unregistered(modules, root)
    present, missing = [], []
    for m in modules:
        p = os.path.join(root, m["path"])
        (present if os.path.exists(p) else missing).append(m)

    if missing:
        lines = [f"  [{m['role']:<4}] {m['name']:<20} -> {m['path']}" for m in missing]
        header = f"モジュールファイルが見つかりません({len(missing)}件):"
        print(header, file=sys.stderr)
        print("\n".join(lines), file=sys.stderr)
        if not allow_missing:
            print("  (--allow-missing を付けると警告に緩和して続行できます)", file=sys.stderr)
            sys.exit(1)
        else:
            print("  --allow-missing 指定のため警告として続行します。", file=sys.stderr)
    return present, missing


def detect_app_version(root):
    path = os.path.join(root, "src", "core", "modAppDef.bas")
    if not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8-sig") as fp:
            txt = fp.read()
    except OSError:
        return None
    m = re.search(r'APP_VERSION\s+As\s+String\s*=\s*"([^"]+)"', txt)
    return m.group(1) if m else None


# ---------------------------------------------------------------------------
# ビルド後自己検証 (MASTER_SPEC §10)
# 再オープンして: 全シート存在 / vba_srcモジュール数一致 /
# 各ソースセル<=32,000字 / olefileでThisWorkbookストリーム復元確認
# ---------------------------------------------------------------------------
def verify_build(out_path, expected_vba_src_names, installer_src, mock_llm_expected):
    errors = []

    try:
        wb2 = openpyxl.load_workbook(out_path, keep_vba=True)
    except Exception as e:
        return [f"再オープン失敗: {e}"]

    got_sheets = set(wb2.sheetnames)
    want_sheets = set(EXPECTED_SHEETS.keys())
    if got_sheets != want_sheets:
        errors.append(f"シート集合が不一致: 期待={sorted(want_sheets)} 実際={sorted(got_sheets)}")

    # 軽量マクロ無効ガード: 先頭シート・アクティブシートであることを検証
    # (マクロ無効時に開いた瞬間、他の何より先にこの案内が見える必要がある)。
    if wb2.sheetnames and wb2.sheetnames[0] != GUARD_SHEET_NAME:
        errors.append(
            f"'{GUARD_SHEET_NAME}' が先頭シートになっていません: 実際の先頭={wb2.sheetnames[0]!r}")
    active_title = wb2.active.title if wb2.active is not None else None
    if active_title != GUARD_SHEET_NAME:
        errors.append(
            f"アクティブシートが'{GUARD_SHEET_NAME}'ではありません: 実際={active_title!r}")

    for name, state in EXPECTED_SHEETS.items():
        if name in wb2.sheetnames:
            actual = wb2[name].sheet_state
            if actual != state:
                errors.append(f"シート'{name}'の可視性不一致: 期待={state} 実際={actual}")

    if "vba_src" in wb2.sheetnames:
        ws = wb2["vba_src"]
        got_names = []
        r = 2
        while ws.cell(row=r, column=1).value:
            nm = ws.cell(row=r, column=1).value
            src = ws.cell(row=r, column=3).value or ""
            got_names.append(nm)
            if len(src) > EXCEL_CELL_LIMIT:
                errors.append(
                    f"vba_src '{nm}' のソースが{len(src)}字でExcelセル上限"
                    f"({EXCEL_CELL_LIMIT})を超過")
            r += 1
        if sorted(got_names) != sorted(expected_vba_src_names):
            errors.append(
                f"vba_srcモジュール集合が不一致: 期待={sorted(expected_vba_src_names)} "
                f"実際={sorted(got_names)}")
    else:
        errors.append("vba_src シートが存在しない")

    if "config" in wb2.sheetnames:
        ws = wb2["config"]
        r, got_mock, n_keys = 2, None, 0
        while ws.cell(row=r, column=1).value:
            if ws.cell(row=r, column=1).value == "mock_llm":
                got_mock = ws.cell(row=r, column=2).value
            n_keys += 1
            r += 1
        if bool(got_mock) != mock_llm_expected:
            errors.append(f"config!mock_llm が期待値と不一致: 期待={mock_llm_expected} 実際={got_mock}")
        if n_keys != len(build_config_rows(mock_llm_expected)):
            errors.append(f"config のキー数が期待({len(build_config_rows(mock_llm_expected))})と不一致: {n_keys}")

    try:
        with zipfile.ZipFile(out_path) as z:
            vba_bin = z.read("xl/vbaProject.bin")
        cfb = ovba.CFBReader(vba_bin)
        tw_dec = ovba.ovba_decompress(cfb.read("ThisWorkbook"))
        # 空チャンクpaddingは圧縮後バイト長をぴったり揃えるためのものだが、
        # 解凍すると末尾にNULバイトが数バイト付加される(OVBA圧縮チャンクの
        # 性質上、5バイト変形チャンクは実際には2バイトのゼロ値に解凍される。
        # 3バイト変形は0バイト)。VBAコンパイラは末尾のNULを無視するため実害は
        # ないが、ここでは「先頭が完全一致し、余剰があるならNULバイトのみ」を
        # もって復元確認とする(V2実証済み挙動)。
        tail = tw_dec[len(installer_src):]
        if not tw_dec.startswith(installer_src) or any(b != 0 for b in tail):
            errors.append("ThisWorkbookストリームの復元結果が自己インストーラソースと不一致"
                           "(先頭一致+末尾NULパディングという想定パターンから外れている)")

        dir_dec = ovba.ovba_decompress(cfb.read("dir"))
        needle = struct.pack("<HI", 0x0019, len("ThisWorkbook")) + b"ThisWorkbook"
        idx = dir_dec.find(needle)
        moffset_ok = False
        if idx >= 0:
            i = idx
            while i < len(dir_dec):
                rid = struct.unpack("<H", dir_dec[i:i + 2])[0]
                sz = struct.unpack("<I", dir_dec[i + 2:i + 6])[0]
                if rid == 0x0031:
                    val = struct.unpack("<I", dir_dec[i + 6:i + 10])[0]
                    moffset_ok = (val == 0)
                    break
                i += 6 + sz
        if not moffset_ok:
            errors.append("dirストリームのThisWorkbook.MOFFSETが0になっていない")

        # olefile側でも同一バイナリを開けることを確認(異なる実装での復元確認)。
        ole = olefile.OleFileIO(io.BytesIO(vba_bin))
        if not ole.exists("VBA/ThisWorkbook") or not ole.exists("VBA/dir"):
            errors.append("olefileでVBA/ThisWorkbookまたはVBA/dirストリームが検出できない")
        ole.close()
    except Exception as e:
        errors.append(f"vbaProject.bin検証中に例外: {e}")

    return errors


# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------
def _sweep_build_leftovers(dist_dir: str) -> None:
    """dist/ に残った *.building.xlsm / *.failed.xlsm を消す(R12-H-7)。
    消せなくてもビルドは続ける(掃除の失敗で配布物を作れなくしない)。"""
    if not os.path.isdir(dist_dir):
        return
    for name in os.listdir(dist_dir):
        if name.endswith(".building.xlsm") or name.endswith(".failed.xlsm"):
            path = os.path.join(dist_dir, name)
            try:
                os.remove(path)
                print(f"  前回の中間生成物を削除: {name}")
            except OSError as e:
                print(f"  注意: 中間生成物を削除できませんでした({name}: {e})")


def main():
    ap = argparse.ArgumentParser(description=f"{APP_TITLE} ビルドスクリプト")
    mode = ap.add_mutually_exclusive_group(required=True)
    mode.add_argument("--dev", action="store_true",
                       help="開発ビルド(config mock_llm=TRUE)。既定出力: MyBookshelf_dev.xlsm")
    mode.add_argument("--prod", action="store_true",
                       help="本番ビルド(config mock_llm=FALSE)。既定出力: MyBookshelf.xlsm")
    ap.add_argument("--root", default=DEFAULT_ROOT,
                     help="modules.json内パスの解決基準ディレクトリ(既定: mybookshelf/)")
    ap.add_argument("--modules", default=DEFAULT_MODULES_JSON, help="modules.json のパス")
    ap.add_argument("--template", default=DEFAULT_TEMPLATE, help="template_skeleton.xlsm のパス")
    ap.add_argument("--out", default=None,
                     help="出力先.xlsmパス(既定: <root>/dist/MyBookshelf[_dev].xlsm)")
    ap.add_argument("--seed", default=None,
                    help="初期ナレッジのパック(.xlsx)。tools/make_seed_pack.py の出力")
    ap.add_argument("--allow-missing", action="store_true",
                     help="modules.jsonに列挙されたファイルの欠落をエラーでなく警告にして続行する")
    ap.add_argument("--allow-embedded-key", action="store_true",
                     help="本番ビルドに azure_embed_key を焼き込むことを明示的に許可する"
                          "(既定は禁止。ブック配布=キー配布になるため)")
    ap.add_argument("--publisher", action="store_true",
                     help="発行者用ビルドを作る(config publish_key を環境変数 "
                          + PUBLISH_KEY_ENV + " から焼き込む)。"
                          "既定の出力名も MyBookshelf_発行者用.xlsm に変わる")
    ap.add_argument("--zip", action="store_true",
                     help="ビルド後、完成した.xlsm + dist/Ghostscript を "
                          "dist/MyBookshelf[_発行者用]_配布.zip へ梱包する"
                          "(大規模配布用。既定ビルドの挙動は変えない)")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    is_dev = bool(args.dev)
    mock_llm = is_dev

    # 2026-07-28(レビュー H-17): 配布ビルドにAzureキーを焼き込ませない。
    # このブックは全社員へ配る前提で、難読化はXOR+16進の可逆変換、鍵も
    # modUtil.bas に平文で同梱されている。つまり「配った時点で全受領者に
    # キーが渡る」。一般配布は embed_transport=ribbon に倒すのが正で、
    # direct が要る特殊ケースだけ --allow-embedded-key で明示的に外す。
    if not is_dev and not args.allow_embedded_key:
        if os.environ.get(AZURE_EMBED_KEY_ENV, "").strip():
            sys.exit(
                f"ERROR: 本番ビルドに {AZURE_EMBED_KEY_ENV} が設定されています。\n"
                "  ブックの配布はそのままキーの配布になります(難読化は可逆・鍵も同梱)。\n"
                "  一般配布は embed_transport=ribbon で行ってください。\n"
                f"  どうしても焼き込む場合は --allow-embedded-key を付けてください。"
            )

    # 2026-07-28(解説書 §12.3 B2): 発行者用と利用者用を、ビルドの時点で
    # 物理的に別物にする。
    # publish_key が入ったブックを配ると、受け取った全員が部門の正典を
    # 上書き発行できてしまう(誤操作防止の関所が外れた状態で配ることになる)。
    # 「発行者用のコピーを作って手で config を消す」という運用は必ず忘れるので、
    # 焼き込みは --publisher を明示したときだけにし、出力名も変える。
    publish_key = ""
    env_pub = os.environ.get(PUBLISH_KEY_ENV, "").strip()
    if args.publisher:
        if not env_pub:
            sys.exit(
                f"ERROR: --publisher を指定しましたが環境変数 {PUBLISH_KEY_ENV} が空です。\n"
                "  発行者用ブックには合言葉が必要です。設定してから再実行してください。"
            )
        publish_key = env_pub
    elif env_pub:
        print(f"注意: {PUBLISH_KEY_ENV} が設定されていますが、--publisher が無いため"
              "焼き込みません(利用者用ビルドとして作ります)。")

    if args.out:
        out_path = os.path.abspath(args.out)
    else:
        if args.publisher:
            fname = "MyBookshelf_発行者用_dev.xlsm" if is_dev else "MyBookshelf_発行者用.xlsm"
        else:
            fname = "MyBookshelf_dev.xlsm" if is_dev else "MyBookshelf.xlsm"
        out_path = os.path.join(root, "dist", fname)

    # 2026-08-01(R12-H-7): 前回の中断・失敗で残った中間生成物を先に片付ける。
    # *.building は Stage5→6 の途中経過、*.failed は自己検証に落ちた不良品。
    # 残しておくと「dist に .xlsm が3つある」状態になり、利用者がどれを開けば
    # よいか分からなくなる(配布物の一意性は R12-9 で決めた原則)。
    _sweep_build_leftovers(os.path.join(root, "dist"))

    print(f"=== build_mybookshelf.py ({'dev' if is_dev else 'prod'}) ===")
    print(f"root:     {root}")
    print(f"modules:  {args.modules}")
    print(f"template: {args.template}")
    print(f"out:      {out_path}")
    ver = detect_app_version(root)
    print(f"APP_VERSION (modAppDef.bas 検出): {ver or '未検出(1-A未実装、またはroot不一致)'}")
    print()

    if not os.path.exists(args.template):
        sys.exit(f"ERROR: テンプレートが見つかりません: {args.template}")
    if not os.path.exists(args.modules):
        sys.exit(f"ERROR: modules.json が見つかりません: {args.modules}")

    try:
        modules = load_manifest(args.modules)
    except BuildError as e:
        sys.exit(f"ERROR: {e}")

    present, missing = validate_modules(modules, root, args.allow_missing)
    print(f"モジュール台帳: 全{len(modules)}件 / 実在{len(present)}件 / 欠落{len(missing)}件")
    for role in ("core", "opt", "test"):
        n = sum(1 for m in present if m["role"] == role)
        print(f"  role={role}: {n}件")

    print("\nStage 1: テンプレート読込 (keep_vba=True)...")
    wb = openpyxl.load_workbook(args.template, keep_vba=True)
    print(f"  初期シート: {wb.sheetnames}")

    print("Stage 2: シート生成 (MASTER_SPEC §4 全13シート + 軽量マクロ無効ガード)...")
    _make_macro_guard(wb)
    _make_howto(wb)
    _make_placeholder(wb, "ホーム", "この画面はマクロ実行時に自動的に構築されます。\n「使い方」タブをご覧ください。")
    _make_placeholder(wb, "マイ本棚", "この画面はマクロ実行時に自動的に構築されます。\n「使い方」タブをご覧ください。")
    _make_placeholder(wb, "ダッシュボード", "この画面はマクロ実行時に自動的に構築されます。\n「使い方」タブをご覧ください。")
    _make_config(wb, mock_llm, publish_key)
    # norm_text(10列目・R12-4): 照合用の正規化済みテキスト。取込時に前計算し、
    # 空欄の行は検索時に遅延バックフィルする(modShelfStore)。数式注入防御の
    # ため text_cols にも入れる(先頭"="の本文が格納型数式にならないように)。
    _make_headers_only(wb, "my_knowledge",
                        ["chunk_id", "source", "origin", "page", "summary",
                         "keywords", "full_text", "added_at", "embedded", "norm_text"],
                        "veryHidden", widths=[32, 24, 14, 6, 50, 40, 80, 20, 10, 80],
                        text_cols=[2, 5, 6, 7, 10])   # source/summary/keywords/full_text/norm_text
    _make_headers_only(wb, "my_vectors", ["chunk_id", "vector_csv"], "veryHidden",
                        widths=[32, 100])
    _sd, _sc, _sv = _make_seed_sheets(wb, args.seed)
    if _sc:
        print(f"  初期ナレッジ: {_sd}資料 / {_sc}チャンク / ベクトル{_sv}件"
              + ("" if _sv else "  ← 未ベクトル化(初回起動で埋め込みが走ります)"))
    else:
        print("  初期ナレッジ: なし(--seed 未指定)")
    _make_headers_only(wb, "my_manifest",
                        ["file_path", "file_name", "modified_at", "size", "chunk_count",
                         "status", "error_note", "ingested_at", "origin"],
                        "hidden", widths=[60, 30, 20, 12, 12, 12, 40, 20, 16],
                        text_cols=[1, 2, 7])   # file_path/file_name/error_note
    _make_headers_only(wb, "my_stats", ["key", "value", "updated_at"], "hidden",
                        widths=[24, 14, 20])
    _make_headers_only(wb, "usage_log",
                        ["timestamp", "event", "mode", "detail", "latency_ms", "hit_count"],
                        "hidden", widths=[20, 16, 10, 50, 12, 10],
                        text_cols=[4])   # detail(質問文等の非信頼テキスト)
    _make_headers_only(wb, "err_log",
                        ["timestamp", "code", "context", "detail", "version", "err_number", "http_status"],
                        "hidden", widths=[20, 10, 24, 60, 12, 12, 12],
                        text_cols=[3, 4])   # context/detail
    _make_headers_only(wb, "ui_state", ["key", "value"], "veryHidden", widths=[24, 40])
    # 共有知フライホイールの受信箱(modInsight)。共有フォルダから届いた
    # 「解決済みQ&A」と「答えられなかった質問」をここに溜め、本棚への
    # 取り込みは利用者が押したときだけ行う。
    _make_headers_only(wb, "insight_inbox",
                        ["nonce", "kind", "user_id", "author", "created_at",
                         "question", "answer_or_reason", "source_or_dept", "consumed",
                         "selected"],
                        "hidden", widths=[34, 8, 20, 18, 18, 60, 80, 30, 10, 10],
                        text_cols=[4, 6, 7, 8])   # author/question/answer/source

    try:
        injected = _make_vba_src(wb, present, root)
    except BuildError as e:
        sys.exit(f"ERROR: {e}")
    print(f"  vba_src: {len(injected)}モジュールを格納 ({injected})")
    print(f"  シート最終構成({len(wb.sheetnames)}件): {wb.sheetnames}")

    if set(wb.sheetnames) != set(EXPECTED_SHEETS):
        sys.exit(f"ERROR: シート構成がMASTER_SPEC §4と不一致: {wb.sheetnames}")

    print("Stage 3: openpyxl保存 (vbaProject.binはスケルトンのまま保持)...")
    # マクロ無効ガードを常に先頭・アクティブにする(マクロ無効時に最初に見える
    # ようにするため)。_make_macro_guard がindex=0で作成しているので通常は
    # 既に先頭だが、保存直前にここで再度保証する。
    guard_idx = wb.sheetnames.index(GUARD_SHEET_NAME)
    if guard_idx != 0:
        sheets = wb._sheets
        sheets.insert(0, sheets.pop(guard_idx))
    wb.active = 0
    with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as tmp:
        tmp_path = tmp.name
    wb.save(tmp_path)
    print(f"  一時保存: {tmp_path} ({os.path.getsize(tmp_path):,} bytes)")

    print("Stage 4: vbaProject.bin 外科パッチ (自己インストーラ注入)...")
    try:
        installer_src = build_installer_src()
    except BuildError as e:
        sys.exit(f"ERROR: {e}")
    with zipfile.ZipFile(tmp_path) as zin:
        parts = {n: zin.read(n) for n in zin.namelist()}
    skel_bin = parts["xl/vbaProject.bin"]
    try:
        patched_bin = patch_installer(skel_bin, installer_src)
    except BuildError as e:
        sys.exit(f"ERROR: {e}")
    assert len(skel_bin) == len(patched_bin), "vbaProject.binのバイト長が変化した(バイナリ整合性エラー)"
    parts["xl/vbaProject.bin"] = patched_bin
    print(f"  vbaProject.bin: {len(skel_bin):,} bytes (不変)")

    # 2026-08-01(R12-9-2): Stage5/6を原子的に確定する。従来はStage5が
    # dist/MyBookshelf.xlsm(正規配布パス。HANDOFF §1の「Code→Download ZIP→
    # dist/…を開く」経路そのもの)を直接上書きしていたため、Stage6の検証に
    # 失敗しても不良ファイルがそこに残置され、前回の良品は既に破壊済みだった。
    # 一時パスへ書いてStage6合格後にos.replaceで確定し、失敗時は不良品を
    # *.failed へ退避する(前回の良品には一切触れない)。
    print("Stage 5: 最終.xlsm書き出し(一時パスへ)...")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    # 拡張子を .xlsm のまま保つ(verify_build内のopenpyxl.load_workbookが
    # 拡張子でフォーマットを判定するため、".building"のような接尾辞を単純に
    # 足すと"サポートしていない形式"として自己検証自体が失敗してしまう)。
    _out_base, _out_ext = os.path.splitext(out_path)
    staging_path = f"{_out_base}.building{_out_ext}"
    failed_path = f"{_out_base}.failed{_out_ext}"
    with zipfile.ZipFile(staging_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for n, data in parts.items():
            zout.writestr(n, data)
    os.unlink(tmp_path)
    print(f"  一時出力: {staging_path} ({os.path.getsize(staging_path):,} bytes)")

    print("\nStage 6: ビルド後自己検証...")
    errors = verify_build(staging_path, injected, installer_src, mock_llm)
    if errors:
        print("自己検証 失敗:")
        for e in errors:
            print(f"  - {e}")
        try:
            if os.path.exists(failed_path):
                os.remove(failed_path)
            os.replace(staging_path, failed_path)
            print(f"  不良な出力を退避しました(前回の良品は無傷): {failed_path}")
        except OSError as e2:
            print(f"  不良な出力の退避にも失敗しました: {e2}")
        sys.exit(1)

    # 検証OKになって初めて正規パスへ確定する(os.replaceは同一ファイル
    # システム内であれば原子的)。ここまで前回の良品は無傷のまま。
    os.replace(staging_path, out_path)
    if os.path.exists(failed_path):
        try:
            os.remove(failed_path)   # 過去の失敗退避物が残っていれば掃除する
        except OSError:
            pass
    print(f"  出力: {out_path} ({os.path.getsize(out_path):,} bytes)")
    print("自己検証 OK: 全シート存在 / vba_srcモジュール数一致 / 各ソース<=32000字 / "
          "ThisWorkbookストリーム復元確認 / dir MOFFSET=0確認")

    if args.zip:
        print("\nStage 7: --zip 配布梱包...")
        zip_path = build_dist_zip(out_path, os.path.dirname(out_path), args.publisher,
                                   is_dev, root)
        print(f"  出力: {zip_path} ({os.path.getsize(zip_path):,} bytes)")

    print("\nDone.")


if __name__ == "__main__":
    main()
