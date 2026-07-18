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
import io
import json
import os
import re
import struct
import sys
import tempfile
import zipfile

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
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
    "my_manifest": "hidden",
    "my_stats": "hidden",
    "usage_log": "hidden",
    "err_log": "hidden",
    "ui_state": "veryHidden",
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
# config 既定値 (MASTER_SPEC §5 config キー台帳を完全反映。値・説明とも準拠)
# ---------------------------------------------------------------------------
def build_config_rows(mock_llm: bool):
    return [
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
        ("embed_transport", "direct", "埋め込みの通信経路: ribbon=AIリボン単発 / direct=Azure APIへバッチ直接送信(裁定②)"),
        ("embed_batch_size", 128, "direct時に1リクエストへまとめるチャンク数"),
        ("azure_http_timeout_ms", 60000, "direct埋め込みのHTTPタイムアウト(ms)。NW瞬断時の無限フリーズ防止。resolve/connectは内部で短めに固定"),
        ("azure_embed_url", "https://hd-us-e2-openai.openai.azure.com/openai/deployments/text-embedding-3-small-g/embeddings?api-version=2024-10-21", "direct時の埋め込みエンドポイント(URL全体)"),
        ("azure_embed_key", "1d545a26153a4f2c990a543031d77b96", "direct時のAPIキー(注意: ブック配布=キー配布になる)"),
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
        ("exp_level_divisor", 100, "ゲーミフィケーション: レベル計算の除数。Lv=Int(√(EXP/除数))+1"),
        ("user_department", "", "分析用: あなたの部署名(分析CSVの部署比較フラグ列に入る。空なら未設定)"),
        ("noise_global_threshold", 2, "ナレッジ自浄: 異なるN人からの⚠️ノイズ報告(P2P集計)でその資料を全ユーザーの検索から組織的除外する閾値"),
        ("admin_users", "", "組織的除外を解除できる管理者ADユーザー名(カンマ区切り)。空なら誰も解除不可"),
        ("shelf_max_chunks", 20000, "本棚のチャンク数上限(Plan B)。大きくするほど資料が入るがサイズと検索時間が増える"),
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
        ("pack_author", "", "パック作成者名(空の場合は初回起動時に入力を促す)"),
        ("debug_mode", False, "TRUE=ゲートウェイのプロンプト/応答を診断用にログへ残す"),
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
    マクロ有効化4ステップ+最初の1冊+質問+困ったら診断、をdocs/00_はじめての方へ.md と
    同じトーン・同じ手順で構成する(D1担当の清書済みユーザー向け文書と歩調を合わせる)。
    ボタン名・タブ名は src/ui/modUIMain.bas・modUIShelf.bas の実際の文字列と一致させてある
    (この関数を編集するときは実装の文字列が変わっていないか必ず突き合わせること)。"""
    ws = wb["Sheet1"]
    ws.title = "使い方"
    _clear_sheet(ws)
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 90

    _banner(ws, f"{APP_TITLE} — 使い方")

    sections = [
        ("📚 マイ本棚AIとは",
         "自分で入れた資料に、承認なしですぐAIに質問できる社内版NotebookLMです。\n"
         "「マイ本棚」タブで資料を追加すると、AIが読める形(ベクトル化)に自動で変換され、\n"
         "「ホーム」タブから質問できるようになります。"),
        ("⚠️ ステップ1: マクロを有効にする(初回のみ・ここで9割の人がつまずきます)",
         "次の(A)(B)は【両方とも必須】です。どちらか片方だけでは動きません。\n"
         "(A) マクロを許可する\n"
         "  1. このファイルをExcelで開く\n"
         "  2. 画面上部に黄色い「セキュリティの警告」バーが出たら「コンテンツの有効化」を押す\n"
         "(B) 自動インストールを許可する(初回のみ)\n"
         "  3. [ファイル]→[オプション]→[トラストセンター]→[トラストセンターの設定]→\n"
         "     [マクロの設定] で「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」に\n"
         "     チェックを入れる→OKで閉じてExcelを再起動→もう一度ファイルを開き直す\n"
         "  4. 数秒待つと、画面が自動的に組み立てられます(タブが増えます)\n"
         "英語の「VBA Project trust required」という表示が出た場合は、(B)がまだ済んでいない合図です。\n"
         "なぜこの手順が必要か: このファイルは初回起動時に自分で画面と機能を組み立てる方式のため、\n"
         "マクロの許可(A)と、組み立てに使う仕組みの許可(B)の両方が必要だからです。"),
        ("📖 ステップ2: 「マイ本棚」タブで資料を1つ追加する",
         "1. 画面下のタブから「マイ本棚」をクリックして開く\n"
         "2. 左上の「＋ 資料を追加」ボタンを押す\n"
         "3. PDF・Word・Excel・テキストなどのファイルを1つ選んで開く\n"
         "しばらくすると、資料カードの状態が ⏳(変換中)→ ✅(完了)に変わります。\n"
         "これで資料が「AIが読める形(ベクトル化)」に変換され、質問できる状態になりました。"),
        ("💬 ステップ3: 「ホーム」タブで質問する",
         "1. 画面下のタブから「ホーム」をクリックして開く\n"
         "2. 質問入力欄(「質問をここに入力してください(例: 〇〇の手続きに必要な書類は?)」)に、\n"
         "   知りたいことを書く\n"
         "3. 「⚡ すぐ聞く (10〜20秒)」が選ばれていることを確認する(初期状態で選択済みです)\n"
         "4. 中央の「💬 質 問 す る」ボタンを押す\n"
         "状態表示が 🔍検索中… → ✍️回答作成中… と進み、回答と出典\n"
         "(「📖 この回答のもと: ○○.pdf p.3」など)が表示されます。"),
        ("🩺 困ったときは",
         "1. 画面右上の「🩺 診断」ボタンを押す\n"
         "2. 表示された画面をスクリーンショットで撮る\n"
         "3. 管理者に送る\n"
         "自己判断で設定をいじる必要はありません。まずスクリーンショットを送ってください。"),
        ("✅ 使えるタブ",
         "「使い方」「ホーム」「マイ本棚」「ダッシュボード」の4つだけです。\n"
         "それ以外のタブはシステムが自動管理しており、通常は表示されません。"),
        ("⚠️ 個人情報について",
         "契約者名・電話番号などの個人情報を含む資料の取込・質問は避けてください。\n"
         "本ツールは社内AIリボン経由でクラウドLLMを呼び出します。"),
        ("このシートについて",
         "このシートはマクロが無効でも読めるようにしてあります(マクロ有効化の案内はここでしか出せないため)。\n"
         "マクロを有効にして開き直すと、実際の操作画面(ホーム/マイ本棚/ダッシュボード)が使えるようになります。\n"
         "全機能を詳しく知りたい方は docs/10_使い方ガイド.md もあわせてご覧ください。"),
    ]

    row = 3
    label_font = Font(bold=True, size=11, color="3C5AA0")
    body_font = Font(size=10)
    warning_fill = PatternFill("solid", fgColor="FFF4D6")
    section_fill = PatternFill("solid", fgColor="F4F6FB")
    for label, body in sections:
        is_warning = "⚠️" in label
        ws.cell(row=row, column=1, value=label).font = label_font
        ws.cell(row=row, column=1).alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(row=row, column=2, value=body).font = body_font
        ws.cell(row=row, column=2).alignment = Alignment(vertical="top", wrap_text=True)
        fill = warning_fill if is_warning else section_fill
        ws.cell(row=row, column=1).fill = fill
        ws.cell(row=row, column=2).fill = fill
        n_lines = body.count("\n") + 1
        ws.row_dimensions[row].height = max(28, 16 * n_lines + 6)
        row += 1

    ws.sheet_properties.tabColor = TAB_COLORS["使い方"]
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


def _make_config(wb, mock_llm: bool):
    ws = wb.create_sheet("config")
    for c, h in enumerate(["key", "value", "description"], 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
    for i, (k, v, d) in enumerate(build_config_rows(mock_llm), 2):
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
  Dim r As Long, n As String, s As String, l As Long
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
      On Error GoTo Done
      Set c = p.VBComponents.Add(1)
      c.Name = n
      ' Strip any auto-inserted lines (e.g. Option Explicit when VBE's
      ' "Require Variable Declaration" is ON). Without this, the source's
      ' own Option Explicit becomes a duplicate -> compile error.
      If c.CodeModule.CountOfLines > 0 Then c.CodeModule.DeleteLines 1, c.CodeModule.CountOfLines
      If LenB(s) > 0 Then c.CodeModule.AddFromString s
    End If
  Next r
  ' Save failures (read-only file, locked share, etc.) must not skip Boot.
  On Error Resume Next
  ThisWorkbook.Save
  Err.Clear
  Application.Run "modBoot.Boot"
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
    new_dir = ovba.pad_to_exact(ovba.ovba_compress(dir_dec), orig_dir_size)
    new_tw = ovba.pad_to_exact(ovba.ovba_compress(installer_src), orig_tw_size)

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
    for m in modules:
        for key in ("name", "path", "role"):
            if key not in m:
                raise BuildError(f"modules.json: エントリに必須キー'{key}'がありません: {m}")
        if m["role"] not in ("core", "opt", "test"):
            raise BuildError(f"modules.json: {m['name']} の role が不正です: {m['role']}")
    return modules


def validate_modules(modules, root, allow_missing):
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
    ap.add_argument("--allow-missing", action="store_true",
                     help="modules.jsonに列挙されたファイルの欠落をエラーでなく警告にして続行する")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    is_dev = bool(args.dev)
    mock_llm = is_dev

    if args.out:
        out_path = os.path.abspath(args.out)
    else:
        fname = "MyBookshelf_dev.xlsm" if is_dev else "MyBookshelf.xlsm"
        out_path = os.path.join(root, "dist", fname)

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
    _make_config(wb, mock_llm)
    _make_headers_only(wb, "my_knowledge",
                        ["chunk_id", "source", "origin", "page", "summary",
                         "keywords", "full_text", "added_at", "embedded"],
                        "veryHidden", widths=[32, 24, 14, 6, 50, 40, 80, 20, 10],
                        text_cols=[2, 5, 6, 7])   # source/summary/keywords/full_text
    _make_headers_only(wb, "my_vectors", ["chunk_id", "vector_csv"], "veryHidden",
                        widths=[32, 100])
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
    _make_headers_only(wb, "err_log", ["timestamp", "code", "context", "detail", "version"],
                        "hidden", widths=[20, 10, 24, 60, 12],
                        text_cols=[3, 4])   # context/detail
    _make_headers_only(wb, "ui_state", ["key", "value"], "veryHidden", widths=[24, 40])

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

    print("Stage 5: 最終.xlsm書き出し...")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for n, data in parts.items():
            zout.writestr(n, data)
    os.unlink(tmp_path)
    print(f"  出力: {out_path} ({os.path.getsize(out_path):,} bytes)")

    print("\nStage 6: ビルド後自己検証...")
    errors = verify_build(out_path, injected, installer_src, mock_llm)
    if errors:
        print("自己検証 失敗:")
        for e in errors:
            print(f"  - {e}")
        sys.exit(1)
    print("自己検証 OK: 全シート存在 / vba_srcモジュール数一致 / 各ソース<=32000字 / "
          "ThisWorkbookストリーム復元確認 / dir MOFFSET=0確認")

    print("\nDone.")


if __name__ == "__main__":
    main()
