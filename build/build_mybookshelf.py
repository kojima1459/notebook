#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_mybookshelf.py — 「マイ本棚AI」ビルドスクリプト。
mybookshelf/dist/MyBookshelf.xlsm (または MyBookshelf_dev.xlsm) を生成する。

--------------------------------------------------------------------------
R35(spec_20260903_R35_配布方式転換.md): 既定の --vba-mode は baked(方式B・
完成品 vbaProject.bin を build_baked_vba_project が書く)。以下1〜6の
「自己インストーラ外科パッチ」の説明は --vba-mode installer(1リリース
限りの開発用フォールバック)にのみ当てはまる。baked モードの説明は
build_baked_thisworkbook / build_baked_vba_project / build/ovba_write.py
の docstring を参照。

アーキテクチャ(V2 = /home/user/notebook/build/build_chatbot_v2.py で実証済みの機構を
mybookshelf/ 配下へ自己完結コピーしたもの。V2側のファイルは一切 import/変更しない):

  1. build/template_skeleton.xlsm — 本物のExcelで作られた.xlsmのコピー。
     中の vbaProject.bin は「本物」であり、どのExcelでも文句なく開ける。
  2. openpyxl(keep_vba=True)でこのスケルトンを読み込み、MASTER_SPEC §4の
     全シート(使い方/ホーム/マイ本棚/ダッシュボード/config/my_knowledge/
     my_vectors/my_manifest/my_stats/usage_log/err_log/ui_state/ocr_cache/
     insight_inbox/seed_*/vba_src。実体は EXPECTED_SHEETS が唯一の台帳)を
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
from openpyxl.worksheet.hyperlink import Hyperlink
import olefile

# ovba.py は同ディレクトリの自己完結モジュール(OVBA圧縮/解凍・CFBリーダー)。
# ovba_write.py は配布方式B(R35)の vbaProject.bin ライター(riskconsulting 移植)。
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ovba
import ovba_write

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
# R35 波1(spec §2-3・§3): vba_src は installer モード専用のシートなので、
# ここ(グローバル定数)には含めない。モードによる分岐は expected_sheets() が
# 関数として持つ(「グローバル定数の書換で分岐しない」という繋ぎ目の指示)。
EXPECTED_SHEETS = {
    GUARD_SHEET_NAME: "visible",
    "使い方": "visible",
    # 管理者向け(R31波3・実機第16報F-C④): 正典発行担当だけが読む説明。
    # 一般利用者の画面からは導線が張られないが(発行ボタン自体が
    # 発行キーを持つ端末にしか出ない)、シートは常時visibleで置く
    # (使い方ページ末尾の「🔑発行を担当する方はこちら→」から誰でも
    # たどり着けるようにするため。隠すと導線が死ぬ)。
    "管理者向け": "visible",
    "ホーム": "visible",
    "マイ本棚": "visible",
    "ダッシュボード": "visible",
    "config": "hidden",
    "my_knowledge": "veryHidden",
    "my_vectors": "veryHidden",
    # R17波0: 構造メタデータ(chunk_id/section_path/refs_out)の受け皿。
    # Phase1本体(パース・取込フック)は未着手で、波0は器のみを焼き込む
    # (R15-FixB FB-2 ocr_cache と同じ理由=実行時Addだと画面が飛ぶため)。
    "chunk_meta": "veryHidden",
    # R17 Phase2: 章単位要約(疑似グローバル検索)の受け皿。chunk_meta と同じく
    # ビルドで器だけ焼き込む(実行時 Add は壊れたブックの自己修復専用)。
    "doc_outline": "veryHidden",
    # R17 Phase3: 用語の表記ゆれ辞書(term, canonical)。取込末尾の名寄せバッチ
    # (modSynonymStore.BuildSynonymsFor)が書き、質問時のクエリ展開が読む。
    # chunk_meta/doc_outline と同じくビルドで器だけ焼き込む。
    "synonyms": "veryHidden",
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
    # 画像PDF OCRの頁チェックポイント(R15-7d。R15-FixB FB-2 でビルド生成へ)。
    # 実行時に optOcrCache が Worksheets.Add で作ることもできるが、Add は
    # 追加したシートをアクティブにするため、取込の途中で画面が知らないシートへ
    # 飛ぶ。最初から在れば、その経路は「壊れたブックの自己修復」だけになる。
    "ocr_cache": "veryHidden",
    "insight_inbox": "hidden",
}


def expected_sheets(vba_mode: str) -> dict:
    """モード別の期待シート集合を返す(名前 -> 可視性)。
    R35 波1: baked(既定・配布方式B)は vba_src シートを持たない。
    installer(開発用フォールバック・spec §2-1)は従来どおり vba_src を足す。"""
    sheets = dict(EXPECTED_SHEETS)
    if vba_mode == "installer":
        sheets["vba_src"] = "veryHidden"
    return sheets

# 可視シートのタブ色 (MASTER_SPEC §10 ビルド仕様): 使い方=緑, ホーム=青, マイ本棚=オレンジ,
# ダッシュボード=紫, 管理者向け=灰(R31波3: 一般利用者の4色系とは別系統にして
# 「発行担当だけが用がある特別なタブ」と一目で分かるようにする)。
TAB_COLORS = {
    "使い方": "00B050",
    "ホーム": "0070C0",
    "マイ本棚": "ED7D31",
    "ダッシュボード": "7030A0",
    "管理者向け": "6B7280",
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
    "  2. 展開してできた「MyBookshelfを起動.bat」をダブルクリックしてください。\n"
    "\n"
    "  ★ 同梱の「MyBookshelfを起動.bat」があればそれから、無ければ他のExcelを\n"
    "     全て閉じてから開き直してください。\n"
    "     他のExcelで仕事中でも安全に開けます。\n"
    "     xlsm を直接ダブルクリックすると、開いたままの他のExcelの中へ\n"
    "     取り込まれてしまい、取込中にその Excel も一緒に固まります。\n"
    "     ランチャーは Excel を必ず別プロセスで起動するので、この巻き添えが\n"
    "     起きません。\n"
    "  ※ bat は zip と同じフォルダに置いたまま使ってください。デスクトップから\n"
    "     起動したい場合は、bat を右クリック →「ショートカットの作成」で作った\n"
    "     ショートカットのほうをデスクトップへ置いてください\n"
    "     (bat 本体だけを移すと、隣にあるはずの xlsm を見つけられません)。\n"
    "  ※ 初回だけ「WindowsによってPCが保護されました」等の警告が出ることが\n"
    "     あります。その場合は bat ファイルを右クリック →「プロパティ」→\n"
    "     下の方にある「許可する」にチェック →「OK」を押してから開いてください。\n"
    "\n"
    "【2. zip の中身はすべて展開してください】\n"
    "  ★ zip の中から直接開くと、取り込んだ資料が次回に残りません。\n"
    "     zip をダブルクリックして中の xlsm をそのまま開くと、Windows は\n"
    "     一時フォルダへ展開したコピーを開きます。保存は成功しますが、\n"
    "     次に開くのは別のコピーで、その日の取り込みは1件も残りません。\n"
    "     必ず zip を右クリック →「すべて展開」してから、展開されたフォルダの\n"
    "     中のファイルを開いてください。\n"
    "  同梱の Ghostscript フォルダは、画像だけのPDF(スキャンした資料など)を\n"
    "  読み取るための部品です。.xlsm と同じ場所に置かれていないと、\n"
    "  画像PDFの取り込みができません。1つも欠かさず、すべて同じフォルダへ\n"
    "  展開してください。\n"
    "  ★ 前のバージョンの MyBookshelf.xlsm が同じ場所にある場合は、\n"
    "     展開する前に必ず削除(またはフォルダごと別の場所へ移動)してください。\n"
    "     上書きを聞かれて「スキップ」を押すと古いファイルが残り、新しい方は\n"
    "     一切展開されません。見た目は更新できたようでも中身は前のままです。\n"
    "\n"
    "【3. マクロを有効にしてください】\n"
    "  開いた直後、画面の上に黄色い帯で「セキュリティの警告」と出たら、\n"
    "  その中の「コンテンツの有効化」ボタンを押してください。\n"
    "  有効化しないと、案内画面が表示されるだけで実際の機能が使えません。\n"
    "\n"
    "{doc_guide}"
    "\n"
    "共有機能(部内でみんなの節約時間を合算する機能)を使う場合は、"
    "ヘルプ→共有フォルダ設定で部の共有パスを入力してください"
    "(未設定なら「みんな」の統計は動きません)。\n"
)

# 同梱した手順書への案内。役割で中身が変わる(発行者用zipにだけ 46 が入る)
# ため、zip を作るときに差し込む。README を2本に分けると片方だけ直す事故が
# 起きるので、変わるのはこの数行だけに閉じ込める。
_README_DOCS_GENERAL = (
    "【同梱の手順書】\n"
    "  docs\\00_はじめての方へ.md            … 開いて最初の質問をするまで(5分)\n"
    "  docs\\41_実機テスト依頼手順_同僚向け.md … 実機テストで見ていただきたい点\n"
    "  docs\\45_実機スモークテスト手順.md     … 短時間の動作確認\n"
)
_README_DOCS_PUBLISHER = _README_DOCS_GENERAL + (
    "  docs\\46_正典発行ガイド_発行担当者向け.md\n"
    "                                        … ★このファイル(発行者用)を\n"
    "                                          受け取った方は、まずこれを\n"
    "                                          お読みください\n"
    "\n"
    "★ このファイルは【発行者用】です。部門の正典を上書きできます。\n"
    "   そのまま部門の方へ配らないでください(配る用は MyBookshelf.xlsm)。\n"
)


def _readme_text(is_publisher: bool) -> str:
    return _README_TEXT.format(
        doc_guide=_README_DOCS_PUBLISHER if is_publisher else _README_DOCS_GENERAL
    )


# ---------------------------------------------------------------------------
# 2026-08-06 R19-5c(実機第6報⑤): 起動ランチャー「MyBookshelfを起動.bat」。
#
# なぜ要るのか: Excelは既定で複数ブックを1プロセスへ結合(マージ)する。本体
# xlsm をエクスプローラーからダブルクリックすると、生き残っていた別のExcel
# (作業用Excelなど)のプロセスへ本体が吸い込まれ、以後は本体のOCR同期呼び出しが
# プロセス全体のメッセージポンプを止める=相手のブックも「■中断」も丸ごと
# 無反応になる(実機第6報⑤)。DisableMergeInstance はダブルクリックには効かない
# ことがMicrosoft公式に明記されており、HKCR の関連付け書き換えは管理者権限が
# 要るうえOffice更新で戻る。VBAだけで結合を防ぐ実装解は存在しない。
#
# 唯一確実なのは「/x で開く入口を配って、そこから起動してもらう」こと。
# excel.exe /x は「新しいインスタンス(別プロセス)を起動する」と公式に明記された
# 正式スイッチで、この経路なら既存プロセスへのマージは起きない。
#   ・excel.exe をフルパスで書かない: Windows の App Paths が excel.exe を
#     解決するので、Office のインストール先(32/64bit・Click-to-Run・年度違い)を
#     ビルド時に決め打ちせずに済む。
#   ・start "" の空タイトルは必須: start の第1引数は【ウィンドウタイトル】で、
#     省略して "..." のパスを書くとそれがタイトルとして食われ、Excelが起動しない。
#   ・%~dp0 でbat自身の場所を基準にする(展開先がどこでも動く。末尾は \ 付き)。
#   ・chcp は打たない: CP932 で書き出すので日本語コメントがそのまま読める。
#
# 2026-09-04 R35 F3b(spec_20260903_R35_配布方式転換.md §10-2): 往復
# (OneDrive原本 → D:複製 → 書き戻し)を追加。R36(a)の前倒し。
# 前提(09-03/04実機確定): D:は信頼できる場所だがシャットダウンで消える/
# OneDriveの同期フォルダは消えないが信頼できる場所ではない/組織ポリシーで
# 信頼できる場所の追加はできない。→「動かす場所はD:、消えない置き場は
# OneDrive」を、bat が利用者の操作なしに両立させる。VBA側は1行も変えない
# (ThisWorkbook.Path が何であっても動くのは方式B移行時点で確認済み)。
#   ・bat 自身の場所(%~dp0)が原本フォルダ。SRC/DSTが同じ(=D:に全部置く
#     旧来運用)なら複製も書き戻しもせず従来どおり起動だけにする。
#   ・D: が無い端末では複製先を作れないので、メッセージを出さず従来どおり
#     SRCから起動するだけにする(D:前提でない端末を壊さない)。
#   ・Excelの起動と終了待ちを別プロセス(自分自身をhiddenな__wait__引数で
#     再起動)に切り出しているのは、10分ごとの書き戻しループと同時に
#     走らせるため。start "" /wait をそのままここで書くとバッチ本体が
#     Excel終了までブロックされ、ループが1回も回らなくなる。
#   ・ループは for /l ではなく goto によるラベル巻き戻しにする。for /l は
#     カウンタ次第で無限ループの温床になりやすく、ブロック内で読む変数は
#     遅延展開(setlocal enabledelayedexpansion)が要る。goto方式なら行ごとに
#     解釈されるため %ERRORLEVEL% や %RETRY% を素の書き方のまま安全に読める。
#   ・コピーの成否判定は %ERRORLEVEL% を変数展開せず `if errorlevel N` の
#     コマンド形で見る(遅延展開の落とし穴を踏まない)。
#   ・ディレクトリの存在確認は末尾に \ を付けて書く(`if exist "D:\"` 等)。
#     付けないと同名ファイルとの誤マッチが起きうる。
# ---------------------------------------------------------------------------
_LAUNCHER_BAT_NAME = "MyBookshelfを起動.bat"

# R35 F3b: D:側の複製先フォルダ(固定)。dev/発行者用など複数の派生ファイルを
# 同じ端末で同時にテストすると、この1フォルダを取り合って複製・書き戻しが
# 混線しうる(通常運用では利用者は1種類のzipしか受け取らないため実害は無い
# 想定。実機テストで複数variantを同時に置く場合は要注意=司令塔へ報告)。
_LAUNCHER_DST_DIR = "D:\\MyBookshelf\\"
_LAUNCHER_GS_DIRNAME = "Ghostscript"
_LAUNCHER_CLOSED_MARK = ".closed"
_LAUNCHER_WAIT_ARG = "__wait__"


def _launcher_backup_name(xlsm_name: str) -> str:
    """書き戻し失敗に備えた退避名(拡張子の直前に _前回 を挿む)。"""
    stem, ext = os.path.splitext(xlsm_name)
    return f"{stem}_前回{ext}"


def _launcher_bat_text(xlsm_name: str) -> str:
    """ランチャーbatの中身(CRLF・CP932で書き出す)。

    R35 F3b: OneDrive(SRC=bat自身の場所)とD:(DST=固定の複製先)を往復する。
    SRC/DSTが同じ・D:が無い、の2ケースは複製せず従来どおり起動だけにする。
    """
    dst = _LAUNCHER_DST_DIR
    gs_dir = _LAUNCHER_GS_DIRNAME
    mark = _LAUNCHER_CLOSED_MARK
    wait_arg = _LAUNCHER_WAIT_ARG
    backup_name = _launcher_backup_name(xlsm_name)
    lines = [
        "@echo off",
        "rem ===== MyBookshelf 起動ランチャー(R35 F3b 往復版) =====",
        "rem 他のExcelで仕事中でも、必ず別プロセスで開くための入口です。",
        "rem xlsm を直接ダブルクリックすると、開いたままの他のExcelに",
        "rem 取り込まれてしまい、取込中にそのExcelも一緒に固まります。",
        "",
        "rem このbatは2つの顔を持つ。通常の起動時と、自分自身をExcel終了",
        "rem 待ちの別プロセスとして再び呼び出したとき(__wait__引数)の2つ。",
        "rem 10分ごとの書き戻しループとExcel終了待ちを同じプロセスで行うと",
        "rem 待ちの間ループが止まってしまうため、待ちだけ別プロセスに分ける。",
        f'if "%~1"=="{wait_arg}" goto :WAIT_AND_MARK',
        "",
        "setlocal",
        "rem SRC = このbatがある場所(OneDriveの消えないフォルダを想定)。",
        "rem DST = D:上の実行用複製先(固定)。どちらも末尾は \\ で揃っている。",
        "set \"SRC=%~dp0\"",
        f'set "DST={dst}"',
        f'set "XLSM={xlsm_name}"',
        "",
        "rem SRCとDSTが同じ場所(=D:に全部置く運用)なら複製も書き戻しも行わず",
        "rem 従来どおり起動するだけにする(/iで大文字小文字を無視して比較)。",
        'if /i "%SRC%"=="%DST%" (',
        '    start "" excel.exe /x "%SRC%%XLSM%"',
        "    goto :EOF",
        ")",
        "",
        "rem D: が無い端末は複製先を作れないので、メッセージを出さずSRCから",
        "rem 起動するだけにする(D:前提ではない端末の動作を壊さない)。",
        'if not exist "D:\\" (',
        '    start "" excel.exe /x "%SRC%%XLSM%"',
        "    goto :EOF",
        ")",
        "",
        "rem 複製先フォルダを用意する。",
        'if not exist "%DST%" mkdir "%DST%"',
        "",
        "rem 本体を新しい方で複製する。/D は「元が複製先より新しいときだけ",
        "rem 複製」なので、前回の書き戻しに失敗してD:側の方が新しい場合は",
        "rem 上書きしない(取込済みデータを消さない)。",
        'xcopy "%SRC%%XLSM%" "%DST%" /D /Y /Q',
        "",
        "rem Ghostscript一式が複製先に無ければフォルダごと複製する",
        "rem (初回だけ・約14MB)。",
        f'if not exist "%DST%{gs_dir}\\gswin32c.exe" (',
        f'    xcopy "%SRC%{gs_dir}" "%DST%{gs_dir}\\" /E /I /Y /Q',
        ")",
        "",
        "rem 前回の「閉じた」印が残っていると、開いた直後に誤って書き戻し",
        "rem 済みと判定してしまうので、起動前に消しておく。",
        f'if exist "%DST%{mark}" del "%DST%{mark}"',
        "",
        "rem Excelの起動と終了待ちを別プロセスに切り出す(自分自身を",
        "rem __wait__引数付きで再度startする)。このプロセスは待たずに",
        "rem 次の10分ループへ進む。",
        f'start "MyBookshelf起動" "%~f0" {wait_arg} "%DST%%XLSM%" "%DST%{mark}"',
        "",
        "rem ===== 10分ごとにD:からSRC(OneDrive)へ書き戻すループ =====",
        "rem for /l ではなくgoto巻き戻しにしているのは、%ERRORLEVEL%等の",
        "rem 遅延展開が不要な素直な形にするため。",
        ":COPY_LOOP",
        "timeout /t 600 /nobreak >nul",
        f'if exist "%DST%{mark}" goto :FINAL_COPY',
        'copy /Y "%DST%%XLSM%" "%SRC%%XLSM%" >nul',
        "goto :COPY_LOOP",
        "",
        ":FINAL_COPY",
        "rem 最終書き戻し。失敗の巻き添えを防ぐため、先にSRC側の現行を",
        "rem 退避してから上書きする。",
        f'if exist "%SRC%%XLSM%" copy /Y "%SRC%%XLSM%" "%SRC%{backup_name}" >nul',
        "",
        "set \"RETRY=0\"",
        ":RETRY_COPY",
        'copy /Y "%DST%%XLSM%" "%SRC%%XLSM%"',
        "if not errorlevel 1 goto :DONE",
        "set /a RETRY=%RETRY%+1",
        "if %RETRY% GEQ 3 goto :FAIL",
        "timeout /t 5 /nobreak >nul",
        "goto :RETRY_COPY",
        "",
        ":FAIL",
        "echo OneDriveへ書き戻せませんでした。",
        'echo %DST%%XLSM% を手でOneDriveへコピーしてください。',
        "pause",
        "goto :EOF",
        "",
        ":DONE",
        "goto :EOF",
        "",
        ":WAIT_AND_MARK",
        "rem %2 %3 は呼び出し側で引用符付きで渡しているので、そのまま",
        "rem 展開すれば引用符ごと引き継がれる。",
        "start \"\" /wait excel.exe /x %2",
        "type nul > %3",
        "goto :EOF",
    ]
    return "\r\n".join(lines) + "\r\n"


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

    # 同梱する手順書。受け取った人が「zipを開いた時点で自分のやることが分かる」
    # 状態にするため、役割ごとに必要な文書だけを入れる。
    #   全員   : 45(スモークテスト) / 00(はじめての方へ) / 41(実機テスト手順)
    #   発行者 : + 46(正典発行ガイド) … 発行者用ブックにしか出ないボタンの説明書。
    #            一般配布zipに入れると「押せないボタンの手順書」を配ることになる。
    # 欠落は BuildError にする(手順書の無いzipが無警告で完成すると、受領者側では
    # 「何をすればいいか分からない」としてしか現れず、原因に辿り着けない)。
    doc_names = [
        "45_実機スモークテスト手順.md",
        "00_はじめての方へ.md",
        "41_実機テスト依頼手順_同僚向け.md",
    ]
    if is_publisher:
        doc_names.append("46_正典発行ガイド_発行担当者向け.md")
    doc_paths = []
    for dn in doc_names:
        dp = os.path.join(root, "docs", dn)
        if not os.path.exists(dp):
            raise BuildError(f"--zip: 同梱すべき手順書が見つかりません: {dp}")
        doc_paths.append((dp, "docs/" + dn))
    try:
        readme_bytes = _readme_text(is_publisher).encode("cp932")
    except UnicodeEncodeError as e:
        raise BuildError(f"--zip: README.txt がCP932でエンコードできません: {e}")

    # R19-5c: ランチャーbatはCP932(cmd.exeの既定コードページ)で書き出す。
    # UTF-8で書くと日本語コメント行が文字化けし、環境によっては行そのものが
    # 壊れて起動コマンドまで巻き添えになる。
    try:
        launcher_bytes = _launcher_bat_text(xlsm_name).encode("cp932")
    except UnicodeEncodeError as e:
        raise BuildError(f"--zip: {_LAUNCHER_BAT_NAME} がCP932でエンコードできません: {e}")

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(xlsm_path, xlsm_name)
        zf.writestr(_LAUNCHER_BAT_NAME, launcher_bytes)
        zf.writestr("README.txt", readme_bytes)
        for dp, arc in doc_paths:
            zf.write(dp, arc)
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
        # R26-1(2026-08-11): 一般アシスタント(本棚を使わない回答)の3段化。
        # 上の quick_*/deep_* は社内ナレッジ検索側の設定で、一般アシスタントは
        # これまで常に quick_* 固定だった。gen_ 接頭辞はその一般側の3段専用。
        ("gen_deep_effort", "medium", "一般アシスタント・しっかり聞くの reasoning_effort(high と medium の差は一般タスクでは薄く、構造化指示のほうが効くため medium)"),
        ("gen_deep_verbosity", "high", "一般アシスタント・しっかり聞くの verbosity(入念の起草・改稿でも同じ値を使う)"),
        ("gen_thorough_effort", "high", "一般アシスタント・入念に聞くの起草/改稿の reasoning_effort(多段推論なので high が効く)"),
        ("gen_thorough_verify_effort", "medium", "一般アシスタント・入念に聞くの検証(査読)の reasoning_effort"),
        ("gen_thorough_verify_verbosity", "medium", "一般アシスタント・入念に聞くの検証(査読)の verbosity(R28 W3-1: 従来のハードコードをconfig化)"),
        ("gen_thorough_loops_max", 2, "一般アシスタント・入念に聞くの「検証→改稿」の最大周回数(verdict:PASSで早期終了。0で検証しない=しっかり相当。4を超える値を書いても4で頭打ち)"),
        # R26-2(2026-08-11): モード切替時の会話メモリ橋渡し(modConvBridge)。
        ("conv_bridge", True, "TRUE=モード切替時に直前1往復の会話をもう片方のモードへ橋渡しする(既定)。FALSEで従来どおり(モードごとの記憶は独立のまま)"),
        ("reasoning_tuning", True, "TRUEでeffort/verbosityを指定。FALSEにすると空送信(古いモデル互換用)"),
        ("llm_wait_sec", 1200, "ChatGPT() 呼び出しのWait秒数"),
        ("topk_quick", 6, "即答モードでLLMに渡す上位ヒット件数"),
        ("topk_deep", 12, "精査モードでLLMに渡す上位ヒット件数"),
        # R34 C(2026-08-20 裁定): 既定を 40,000 → 60,000 へ。段階的にコンテキストを
        # 厚くする方向(B2 の再ランク抜粋 300→700字、B3 の deep 近傍結合)と対で、
        # 増えた材料が打ち切りで捨てられないようにするための引き上げ。
        # VBA側 modPrompts.SafeMaxContextChars / modAskRetrieve の GetLong 第2引数は
        # 40000 のまま据え置く(configシートが欠損・破損した端末で保守的な側へ倒す
        # ための既定であって、正常な配布物ではこの行の 60000 が必ず勝つ)。
        ("max_context_chars", 60000, "プロンプトに載せる本文合計の文字数上限(既定60,000。遅い・固まるときは40000へ戻す。管理者向けシート⑦参照)"),
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
        ("chunk_target_chars", 700, "チャンクの目安文字数(グローバル既定)"),
        ("chunk_overlap_chars", 150, "チャンクのオーバーラップ文字数(グローバル既定)"),
        ("chunk_max_chars", 1800, "条文を分割せず1チャンクに収める上限(構造認識時・グローバル既定)"),
        # R14-5b(実機第3報 RC7・チャンク粒度改善): PDF/Wordはchunk_max_charsの
        # 支配で粒度が粗くなりやすい(Excelの細かさは行長との偶然の一致)。
        # modShelf.ChunkParamForの拡張子別キー(baseKey_ext)がこの3種だけを
        # PDF/Word専用値へ倒す。グローバル既定(上の3行)とxlsx等の他拡張子は
        # 従来どおり不変(拡張子別キーが無ければグローバル既定へ2段フォール
        # バックする)。
        ("chunk_target_chars_pdf", 450, "PDFのチャンク目安文字数(グローバル既定700より細かく。RC7対応)"),
        ("chunk_overlap_chars_pdf", 100, "PDFのチャンクオーバーラップ文字数(RC7対応)"),
        ("chunk_max_chars_pdf", 900, "PDFの1チャンク上限文字数(グローバル既定1800より細かく。RC7対応)"),
        ("chunk_target_chars_docx", 450, "Word(docx)のチャンク目安文字数(RC7対応。PDFと同値)"),
        ("chunk_overlap_chars_docx", 100, "Word(docx)のチャンクオーバーラップ文字数(RC7対応)"),
        ("chunk_max_chars_docx", 900, "Word(docx)の1チャンク上限文字数(RC7対応)"),
        ("chunk_target_chars_doc", 450, "Word(旧doc)のチャンク目安文字数(RC7対応。docxと同値)"),
        ("chunk_overlap_chars_doc", 100, "Word(旧doc)のチャンクオーバーラップ文字数(RC7対応)"),
        ("chunk_max_chars_doc", 900, "Word(旧doc)の1チャンク上限文字数(RC7対応)"),
        ("embed_prefix_breadcrumb", True, "TRUE=チャンク先頭に【資料名>章>条】を前置して文脈付きで保存・検索する"),
        ("retrieve_mode", "multi", "検索方式: single=単段 / multi=多段RAG(拡張→マルチクエリ→再ランク)"),
        ("expand_enabled", True, "TRUE=質問をAIで検索用に拡張(独立質問化+サブクエリ+仮回答)"),
        ("expand_subqueries", 3, "クエリ拡張で生成するサブクエリ本数"),
        ("expand_model", "", "拡張段のモデル(空=quick_modelを使用)"),
        ("expand_effort", "low", "拡張段のreasoning_effort"),
        ("expand_verbosity", "low", "拡張段のverbosity"),
        ("multi_candidates", 40, "マルチクエリ検索の候補プール上限"),
        # R27 F1-1(実機第12報②): キーワード加点(modSparse.KeyScore)が無上限で、
        # 長い資料名やHyDE由来の長語キーでスコアが330まで伸びていた。
        # 0.06(SPARSE_WEIGHT)×330=20の加点はベクトル類似度(-1〜1)を完全に
        # 押し流し、実質キーワード検索になっていた。既定10で加点上限0.6=cosと同格。
        ("sparse_keyscore_cap", 10, "キーワード加点(KeyScore)の上限。0.06倍してベクトル類似度へ足すので、10なら加点は最大0.6=cosと同じ土俵。大きくするほどキーワード一致が順位を支配する。0で上限なし(R27以前の挙動)"),
        # R29 W2-5(実機第14報): DiversitySwapPick(最終hitsが1資料へ収束した
        # ときの最小介入)に相対スコア下限を追加。pool側の"別資料"代表が
        # 現hitsの最下位よりも著しく弱いスコアだと、弱い他資料で薄めるだけの
        # 差し替えになっていたため。
        ("diversify_min_ratio_x100", 70,
         "分散差し替え(DiversitySwapPick)の相対スコア下限(%×100)。pool側の"
         "差し替え候補のスコアが現hits最下位スコアのこの%未満なら差し替えない。"
         "0で下限なし(旧挙動=無条件差し替え)"),
        ("rerank_enabled", True, "TRUE=候補チャンクをAIで再ランクしてから回答生成する"),
        ("quick_expand", False, "TRUE=「すぐ聞く」でも質問拡張を行う(AI呼び出しが1回増え数秒遅くなる)"),
        ("quick_rerank", False, "TRUE=「すぐ聞く」でも再ランクを行う(AI呼び出しが1回増え数秒遅くなる)"),
        ("topk_thorough", 16, "「入念に調べる」でLLMに渡す上位ヒット件数。実測で件数増は効果が薄いため控えめ"),
        ("thorough_subqueries", 6, "「入念に調べる」の質問拡張で作るサブクエリ数。角度の数が精度に効く(実測 R@10 90%→97%)"),
        # R14-8a(入念モードの本格強化): 入念だけ生成側のパラメータを分ける。
        # ここが deep と共有だったため「同じことを少し多い資料でやり直すだけの
        # 遅いモード」になっていた(実機第3報 RC8)。
        ("rerank_effort_thorough", "medium",
         "「入念に調べる」の再ランク段のreasoning_effort(他モードは rerank_effort)。"
         "どの資料を根拠にするかを決める段で、ここが雑だと後段の検証では直らない"),
        ("thorough_digest_effort", "low",
         "「入念に調べる」(1)資料の要点整理のreasoning_effort。抜粋から質問に効く部分を"
         "抜き出すだけの作業なので既定は軽い"),
        ("thorough_draft_effort", "high", "「入念に調べる」(2)下書き生成のreasoning_effort"),
        ("thorough_draft_verbosity", "high", "「入念に調べる」(2)下書き生成のverbosity"),
        ("thorough_critique_effort", "medium",
         "「入念に調べる」(3)自己批判のreasoning_effort。下書きの未検証の断定・出典不備・"
         "論点漏れ・憶測を指摘だけさせる段(書き直しはさせない)"),
        ("thorough_verify_effort", "high",
         "「入念に調べる」(4)検証のreasoning_effort。verbosityは deep_verify_verbosity を共用"),
        # R16-3A(複合質問の分解): 「入念に調べる」だけ、質問を論点へ割ってから
        # 論点ごとに調べる。off にすると段0の判定ごと呼ばない=従来動作に戻る。
        ("decompose_mode", "auto",
         "「入念に調べる」で複合質問を論点ごとに分けて調べるか。auto=既定"
         "(decompose_min_chars以上の長さの質問だけ判定する)/always=長さを見ずに毎回判定/"
         "off=分解しない(従来どおり1本の質問として調べる)"),
        ("decompose_max_parts", 3,
         "複合質問を分ける論点数の上限(2〜5)。1論点ごとにAI呼び出しが1回増えるため、"
         "増やすほど正確になる代わりに待ち時間が伸びる"),
        ("decompose_min_chars", 25,
         "decompose_mode=auto のとき、この文字数以上の質問だけ論点分けの判定を行う。"
         "短い質問は割る論点が無く、判定の1回ぶんだけ遅くなるため"),
        # R16-3B(逆質問=番号選択肢): 読み方が複数ある質問に、薄い回答ではなく
        # 「どれを調べますか?」を番号で返す。off にすると従来どおり回答を作る。
        ("clarify_mode", "auto",
         "「入念に調べる」で読み方が定まらない質問に、番号の選択肢で聞き返すか。"
         "auto=既定(選択肢が2件以上作れたときだけ聞き返す。「1」「1と3」で返信でき、"
         "質問の書き直しも受け付ける)/off=聞き返さずそのまま回答を作る"),
        # R16-3C(精読=近傍チャンク束ね): 根拠に選んだチャンクの前後も一緒に読む。
        # R16H FA-4: 適用先は「入念に調べる」だけだった(深掘りからは外していた。
        # 深掘りの戻り件数は「N件ヒット」の件数バッジと直結していて、近傍を
        # 混ぜると嘘になるという理由)。
        # R34 B3(2026-08-20 裁定): 適用先へ「しっかり調べる」を加えた
        # (modAskRetrieve.ExpandNeighborsIfDeep。ゲートは modMode.UseNeighborExpand
        # で deep 限定=入念の自前呼び出しとの二重結合は起きない)。この1つの数字が
        # 2モードの精読半径を兼ねるようになったため、説明文の適用先を実装へ合わせる。
        ("deep_neighbor", 2,
         "「しっかり調べる」「入念に調べる」の精読半径。根拠チャンクの前後何個ぶんを"
         "一緒に読むか(0=off。両モード共通の設定です)。表や条文が途中で切れて"
         "『資料に記載なし』になるのを防ぎます。"
         "増やすほど読む量が増える(1回あたり max_context_chars で頭打ち)"),
        # R17 Phase1(構造グラフ): 取込時に拾った section_path / refs_out を検索へ
        # 合流させるか。期待効果は設計書 docs/dev/design_20260805_R17 §3 Phase1 —
        # 規程QAの最頻質問(「第5条の免責は?」「8条との関係は?」)への直撃で、
        # LLM呼び出しは1回も増えない(増えるのはワークシート読みだけ)。
        # off にすると chunk_meta を一切読まず R16 までと同じ動きに戻る。
        ("graph_refs", "on",
         "条文の参照関係(「第8条による」「別表2のとおり」)を回答の材料に足すか。"
         "on=既定(根拠にした条文が参照している条文・別表を同じ資料の中から一緒に読み、"
         "質問が名指しした条番号のチャンクは必ず材料に入れる)/off=検索ヒットだけで答える。"
         "この設定を on にしても、資料を取り込み直すまでは従来どおりの動きになる"),
        # R17 Phase2(章単位要約=疑似グローバル検索): 取込時に章ごと1回だけ
        # 要約を作って doc_outline へ貯め、「〜を全部教えて」型の質問(入念モード
        # の段0が verdict=global と判定したもの)で章をまたいで答える。
        # off にすると取込時のLLM呼び出しが増えず、質問も従来どおり single へ
        # 落ちる(doc_outline が0行の本棚でも同じ=フェイルセーフ)。
        ("graph_outline", "on",
         "資料の章ごとの要約を取込時に作るか(俯瞰質問「〜を全部教えて」「全体像は?」への対応)。"
         "on=既定(章の数だけAIを呼ぶので取込は長くなる。254頁の規程で+3〜8分。中断ボタンで途中まで保存)/"
         "off=作らない(取込は従来どおりの速さ。俯瞰質問は従来の検索で答える)。"
         "この設定を on にしても、資料を取り込み直すまでは従来どおりの動きになる"),
        # R17 Phase3(名寄せ辞書): 章要約のあとに用語一覧をAIへ1回渡し、表記ゆれ
        # (「回収」⇔「リコール」等)のグループを synonyms シートへためる。
        # 質問時はその表を読んで質問文へ同義語を足してから検索する(語彙が違う
        # だけでヒットしない事故を減らす)。LLM呼び出しは資料1本の取込につき
        # 最大1回だけ増える(質問側は呼び出しを増やさない=文字列処理のみ)。
        # R17H FB-4(A-M6・記録+説明): 名寄せは【章要約の配線に相乗り】している。
        # BuildSynonymsFor を呼ぶのは modOutlineBuild.BuildOutlineFor の末尾で、
        # そこへ到達するには chunk_meta があり章キーが1つ以上取れる必要がある。
        # つまり graph_outline=off の本棚・章立ての無い資料では graph_synonyms=on
        # でも辞書は1行も増えない。配線の分離(IngestFile から直接呼ぶ)は
        # modShelf 凍結解除後の次期。それまでは事実を説明文へ書いておく。
        ("graph_synonyms", "on",
         "取込時に用語の表記ゆれ辞書(synonyms)を作り、質問時に質問文へ同義語を足すか。"
         "on=既定(章要約のあとAIへ用語一覧を1回だけ渡しsynonymsシートを更新する。"
         "質問は一致した語の同義語を最大3語まで質問文に足してから検索する)/"
         "off=辞書を作らず質問文もそのまま検索する(既に作った辞書は残るがoffの間は読まれない)。"
         "【前提】辞書づくりは章の要約の直後に走るため、graph_outline が off の本棚と、"
         "章・条の見出しが取れない資料では on にしても辞書は増えない(質問側は従来どおり動く)。"
         "作れた件数は usage_log の synonyms_built(groups=N)で確認できる。"
         "synonymsシートは取込・同期のたびには読み直さない(1セッション1回のみ。"
         "更新は次にブックを開いたときから反映)"),
        ("deep_scope_subqueries", 6, "「続けて質問」を『しっかり調べる』で行うときに、会話で引用済みの資料の中だけを掘るために作るサブクエリ数。狭い範囲を多角度から見るための本数(0以下は6扱い)"),
        ("rerank_model", "", "再ランク段のモデル(空=quick_modelを使用)"),
        ("rerank_effort", "low", "再ランク段のreasoning_effort"),
        ("rerank_verbosity", "low", "再ランク段のverbosity"),
        ("answer_tags", True, "TRUE=回答を<thinking>/<answer>構造で生成し<answer>のみ表示"),
        ("strict_grounding", True, "TRUE=資料のみ・出典必須・『資料からは判断できません』を強制"),
        ("quick_expand_light", True, "TRUE=すぐ聞くモードでは拡張を軽量化(速度優先)"),
        ("nexus_ui", True, "TRUE=起動時にNexus Agent(SPA風チャットUI)を表示する"),
        ("freeze_keep_banner", True,
         "TRUE=取込・入念な質問応答などの長時間ブロック中、DWMの「応答なし」白画面化"
         "(ゴースト化)をuser32.DisableProcessWindowsGhostingで抑止し、進捗バナーの"
         "文字を最後まで見えるようにする(R16-2b)。抑止中はウィンドウの移動・"
         "最小化・×閉じが効かなくなる(公式の既知の制約。強制終了はタスクマネー"
         "ジャーから)。FALSEにすると従来どおり約5秒で白画面化する挙動へ戻る"),
        ("nexus_share_path", "",
         "P2Pナレッジ共有フォルダ(Phase 4)。未設定なら共有機能は休止。"
         "ヘルプ→共有フォルダ設定で入力(R13-7f: 実在しないダミーパスを既定に"
         "していたため「未設定」警告が働かず、届かないまま『みんな』統計が"
         "0のまま凍結して見えた。空文字を既定にして警告経路を正しく働かせる)"),
        ("exp_question", 5, "ゲーミフィケーション: 質問1回で得るEXP"),
        ("exp_register", 20, "ゲーミフィケーション: ナレッジ登録1件で得るEXP"),
        ("exp_thumbup", 10, "ゲーミフィケーション: 🟢自己解決(役立った)1回で得るEXP"),
        ("exp_pack_share", 30, "ゲーミフィケーション: パック共有(出力)1回で得るEXP"),
        ("exp_feedback", 5, "ゲーミフィケーション: ご意見箱(フィードバック/バグ報告)1回で得るEXP(1日1回まで)"),
        ("exp_level_divisor", 100, "ゲーミフィケーション: レベル計算の除数。Lv=Int(√(EXP/除数))+1"),
        ("user_department", "", "分析用: あなたの部署名(分析CSVの部署比較フラグ列に入る。空なら未設定)"),
        ("minutes_per_selfsolve", 15,
         "Hub「自分の節約時間/みんなの節約」の換算係数(自己解決1件=何分の節約とみなすか)。"
         "R13-7d: modStats/modBoard/modDashStatに3重複していた同じ意味の定数を、"
         "この1キーへ統合した(modP2PIo.MinutesPerSelfsolveが読む唯一の窓口)"),
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
        # R17 Phase3: 常時ON化(off→light)。取込末尾で一括処理すると254頁規模で
        # 1時間級になるため、EnrichPending側の既定上限(30チャンク/回)で小口化し、
        # 取込直後に少し・残りは同期のたびに少しずつ追いつく「後追い」方式にした。
        ("enrich_mode", "light",
         "off/light/full: 取込後にAIが少しずつ資料を要約・キーワード付けする"
         "「バッチ富化」の強さ。light=既定(取込直後に少量、残りは同期のたびに"
         "少しずつ追いつく。1回の処理はEnrichPending既定30チャンクぶんだけで、"
         "取込や同期を長時間ブロックしない)/full=現状はlightと処理内容の差は無い/"
         "off=作らない(要約・キーワード列は空のまま。検索は本文だけで行う)"),
        ("max_pages_per_file", 300, "1ファイルあたりの抽出ページ数上限(超過分は打ち切りpartial扱い)"),
        ("followup_max_pairs", 5, "『続けて質問』で引き継ぐ会話履歴の最大ペア数。0以下で機能無効"),
        ("word_export_effort", "medium", "『Wordで開く』の文書整形に使う reasoning_effort"),
        ("word_export_verbosity", "medium", "『Wordで開く』の文書整形に使う verbosity"),
        ("feature_tts", False, "opt機能フラグ: 読み上げ(AIリボン非公開機能のため提供不可。既定FALSEのまま変更しない)"),
        ("feature_vision", True, "opt機能フラグ: 画像読み取り・スクショ取込(公式仕様確定済み。問題があればFALSEで無効化)"),
        ("feature_markdown", True, "opt機能フラグ: Markdown表示・Wordで開く(公式仕様確定済み。問題があればFALSEで無効化)"),
        # 2026-08-16(R33 W4-4): 既定TRUE→FALSE。optDiffDoc.bas は配布に含まれ
        # 注入もされるが、InvokeFeature("diffdoc",…)/FeatureEnabled("diffdoc") の
        # 呼び出しが src 全体で0件＝利用者から起動する導線がどの画面にも無い。
        # 既定TRUEのままだと診断画面に「モジュールあり/設定有効」と出て、
        # 管理者が「有効なのに使えない」と読む。機能は殺さず(モジュールは同梱の
        # まま・キーもTRUEにすれば効く)、既定だけを実態に合わせて落とす。
        # 導線の新設は機能追加なので docs/dev/TODO.md §F の宿題として残す。
        ("feature_diffdoc", False, "opt機能フラグ: 約款差分比較(現在は起動する導線が無いため既定FALSE)"),
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
        ("vision_pdf_max_pages", 300,
         "画像PDFを何ページ目まで読み取るか(2026-08-04 R15-7a: 100→300)。"
         "1ページごとにAIを1回呼ぶので、大きくすると時間もコストも比例して増える。"
         "20ページ単位で画像化→OCR→画像削除を繰り返すため、上限を上げても"
         "一時領域は20ページぶんで頭打ち。超過分は打ち切り(本棚カードに"
         "「上限Nページのため先頭Nページのみ」と表示)。上限の上限は300"),
        ("ocr_confirm_min_minutes", 5,
         "画像PDFの読み取りに何分以上かかる見込みのとき、取り込む前に確認"
         "ダイアログを出すか(2026-08-04 R15-7b。2026-08-05 R18-1f: 15→5分)。"
         "総ページ数×1ページあたりの実測時間(実測が無い端末は25秒/ページ)で"
         "見積もる。前回の途中まで読めているページは待ち時間に数えない。"
         "15分だと24ページ(8〜10分)のような【十分に長い】資料でも何も聞かず"
         "始まってしまい、確認と作業用Excelへの導線がどちらも出なかった"
         "(実機第5報①)。0にすると確認せずに始める"),
        ("vision_pdf_dpi", 150,
         "画像PDFをページ画像にするときの解像度。150で十分読める。"
         "細かい文字が読めないときだけ300へ(処理時間は約2倍になる)"),
        ("vision_pdf_timeout_sec", 120,
         "本文抽出でGhostscriptの処理が【まったく進まなくなってから】何秒で見切るか。"
         "進捗が観測できる場合はページが進む限り待ち続ける。"
         "ページ数の多い資料で頻発するときだけ大きくする"),
        ("gs_abs_timeout_sec", 1200,
         "Ghostscriptの処理を待つ絶対上限(秒)。進んでいても必ずここで打ち切り、"
         "壊れたPDF1つでExcelが何十分も戻ってこない事態を防ぐ。"
         "画像PDFのOCR経路(進捗を見ない待ち)はこの絶対上限だけを使う。"
         "数百ページの資料を入れるときだけ大きくする"),
        ("pack_author", "", "パック作成者名(空の場合は初回起動時に入力を促す)"),
        ("debug_mode", False, "TRUE=ゲートウェイのプロンプト/応答を診断用にログへ残す"),
        ("chat_log_enabled", True, "TRUE=チャット履歴シートに質問と回答を記録する(最新100件・古い順に自動削除)"),
        ("low_hit_warn_score", 0.3, "検索ヒットの最高スコアがこの値未満のとき回答に⚠️関連薄い警告を付ける(0で無効)"),
        ("active_channel", "",
         "いま接続している部門の公式ナレッジ(1つだけ)。ナレッジ画面の「部門チャンネル」で"
         "切り替える。切り替えると前の部門の内容は本棚から外れる(マイ本棚の資料は残る)"),
        ("unsubscribed_channels", "",
         "購読【しない】部門チャンネル(縦棒 | 区切り。先頭にも | が付きます)。既定は空=全チャンネルを自動購読する。"
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
        # R32波2(2026-08-14 W2-1): 波1(みんなの困りごと)が使う2キー。
        # 衝突回避のため波1では見送り、波2でまとめて追加する。既定値は
        # modInsightIo.bas(gap_keep_days・GcOldNonces手前)/modInsight.bas
        # (gap_dup_hours・GapDupBlocked手前)のGetLong呼び出しの既定値と一致させる。
        ("gap_keep_days", 30,
         "「みんなの困りごと」の受信箱に残す保持期間(日)。created_atがこれより古い"
         "gap行は受信箱から落とす(件数が永久に増え続けるのを防ぐ)。手動の「解決済みに"
         "する」操作は無いため、この保持期間だけが件数を減らす唯一の経路"),
        ("gap_dup_hours", 24,
         "同じ趣旨の質問(NormKeyで正規化した先頭40字が一致)を短時間に連投したとき、"
         "この時間(時間単位)は再送しない。0以下で連投抑止を無効化"),
        # R32波2(2026-08-14 W2-1・ユーザー裁定=config切替+既定オフ): PII検知が
        # 厳しすぎるとの実機報告(誤検知で発行が止まる)を受け、走査そのものを
        # 一旦オフに倒せるスイッチを新設。既定FALSE=走査しない(誤検知で発行が
        # 止まる実害を止める)。関所は modPackExport.ExportPackToFile の
        # ScanChunksForPii呼び出し手前1箇所のみ(src/pack/modPackExport.bas)。
        # 【注意】共有フォルダへ自動発信される「みんなの困りごと」側のPII走査
        # (modInsightGate.PiiBlocked)はこのキーの影響を受けず常時走る(R32 W1-8)。
        ("pii_scan_enabled", False,
         "TRUE=パック書き出し(📦パック出力・📤正典発行)の前に個人情報っぽい文字列が"
         "無いか走査し、見つかったら書き出しを中止する。FALSEで走査自体を行わない"
         "(既定。誤検知で正常な発行まで止まっていた実機報告を受けての既定変更)。"
         "この設定は「みんなの困りごと」を共有フォルダへ送る前の走査には影響しない"
         "(あちらは自動発火かつ他人の目に触れる経路のため常時走る)"),
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
        # 2026-08-06 R19-4c(実機第6報④): 「免責は?」のように、スコアは高いのに
        # 当たり先が複数の資料へ割れる質問への聞き返し。1位資料と2位資料の
        # 最高スコアの差がこの値÷100より小さいときだけ発動する。小さめ(0.10)に
        # 取るのは、僅差のときだけに限って誤発動を抑えるため(調査④班6.1)。
        # 2026-08-07 R21-2 D1(実機第8報⑧計器修理): 以下2キーは非推奨。
        # SparseBoost(modRetrieve、上限なし)が資料ごとに桁違いに乗ると絶対差が
        # 実測不能な値(実機deepモード gap=390)まで膨らみ、どこに閾値を置いても
        # 機能しなかった。modAskRetrieve.IsTooVague からの参照は廃止し、
        # 下の dispersion_rel_gap_x100 系(相対gap)へ完全移行した。値そのものは
        # 校正データの参考として残す(キーは削除しない)。
        ("ambiguous_dispersion_gap_x100", 10, "[非推奨/R21-2で参照廃止] 資料分散の絶対gap閾値×100(旧実装)。相対gap版(dispersion_rel_gap_x100)に置き換わった"),
        ("thorough_dispersion_gap_x100", 25, "[非推奨/R21-2で参照廃止] 「入念に調べる」専用の絶対gap閾値×100(旧実装)。相対gap版(thorough_dispersion_rel_gap_x100)に置き換わった"),
        # 2026-08-07 R21-2 D1: 相対gap = (1位資料の最高スコア - 2位資料の最高
        # スコア) ÷ 1位資料の最高スコア × 100。全スコアを定数倍しても値が
        # 変わらない(スケール不変)ため、SparseBoostの乗り方が資料で違っても
        # 閾値の意味が保たれる。判定はrerank/絞り込み前のpool(候補集合)に対して
        # 行う(modAskRetrieve.RunMultiRetrieve)。usage_log "dispersion" の
        # detail(src/b1/b2/rel)で校正する。
        ("dispersion_rel_gap_x100", 15, "資料分散で聞き返す閾値×100(相対gap)。1位資料と2位資料の最高スコアの差が1位に対してこの割合未満なら「どの資料か」を聞き返す。0で機能OFF"),
        ("thorough_dispersion_rel_gap_x100", 20, "「入念に調べる」専用の資料分散閾値×100(相対gap。既定20。quick/deepはdispersion_rel_gap_x100=15のまま)"),
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
        # 2026-08-05 R18-2g(実機第5報⑧): zip をダブルクリックして中の xlsm を
        # 直接開くと、Windows は %TEMP%\\Temp1_....zip\\ へ展開してそこを開く。
        # 保存は成功し読み取り専用にもならないため、アプリからは完全に正常に
        # 見えるのに、次に開くのは別の展開コピーで取り込んだ資料は1件も残らない
        # (検知ゼロだった=調査agent1 §4)。マクロ無効でも必ず見える面はこの
        # シートだけなので、配布 zip の README.txt と同じ注意をここにも1行置く。
        (8, "※ zip の中から直接開くと、取り込んだ資料が次回に残りません。"
            "必ず zip を右クリック →「すべて展開」してから、"
            "展開されたフォルダの中のファイルを開いてください。",
         Font(size=11, bold=True, color="8A3B00"), 46),
        # 2026-08-06 R19-5c(実機第6報⑤): Excelは既定で複数ブックを1プロセスへ
        # 結合する。他のExcelを開いたまま本体をダブルクリックすると、そちらの
        # プロセスへ吸い込まれ、取込中に相手のブックごと固まる(中断ボタンも
        # 効かなくなる)。VBAでは結合を防げないので「正しい入口」を配るしかない。
        # マクロ無効でも必ず見える面はこのシートだけなので、README と同じ注意を
        # ここにも1行置く(利用者が実際に見るのは、ほぼこの1面だけ)。
        # 2026-08-06 R19H FA-7(A-M⑦・B-M⑤): 「必ずランチャーから」だけだと、
        # bat を持っていない人(旧版のzip・xlsm だけ転送された人・zipから1個
        # だけ取り出した人)には存在しないものへの誘導になる。言われたとおりに
        # できない案内は、警告そのものを読まれなくする。有る場合と無い場合の
        # 両方に効く1文へ揃える(README・modIntegrity.CohabitWarnMsg と同文)。
        (10, "※ 同梱の「MyBookshelfを起動.bat」があればそれから、無ければ"
             "他のExcelを全て閉じてから開き直してください"
             "(他のExcelで仕事中でも安全に開けます)。"
             "そのまま直接ダブルクリックすると、開いたままの他のExcelに"
             "取り込まれ、取込中にそちらも一緒に固まることがあります。",
         Font(size=11, bold=True, color="8A3B00"), 46),
        (12, "※ 有効化しても画面が変わらない場合は、ファイルを一度閉じて開き直してください。",
         Font(size=11, italic=True, color="52606D"), 34),
    ]
    for row, text, font, height in entries:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        cell = ws.cell(row=row, column=1, value=text)
        cell.font = font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
        ws.row_dimensions[row].height = height

    # 背景を軽く塗って独立した案内ページに見えるようにする(装飾のみ)
    # R18-2g で行を2行ぶん増やし、R19-5c でさらに2行増やしたので、塗る範囲も
    # 末尾(13行目)まで広げる。
    for r in range(1, 14):
        for c in range(1, 7):
            ws.cell(row=r, column=c).fill = bg

    ws.sheet_state = "visible"
    return ws


def _make_howto(wb):
    """使い方シート: マクロ無効でも読める唯一の救済ページ。

    実機要望(2026-07-26)「文字だらけで表もなく、改行も適切でなく、認知負荷が
    高くて読む気が失せる」への全面作り直し(第1版)。Shapeは使わずセルだけで、
    (1) 番号つきの手順、(2) 1行1項目の早見表、(3) 用語のミニ辞書 という
    3つの構造に分解した。1セルに長文を詰め込まず、手順は1ステップ1行にする。

    R30波3(実機第15報・調査班E確定): 11セクション上から下へ読むだけの構成が
    「マイ本棚のボタン説明が皆無」「削除の仕方が書かれていない」「部門設定の
    手順が書かれていない」「文脈引き継ぎ等の目玉機能の紹介がない」「ページ内で
    迷子になる」という不満に直結していたため、目次つき7章構成へ全面刷新した。
    冒頭の目次は openpyxl の Hyperlink(location=シート内アンカー)でクリック
    ジャンプでき、各章末の「▲ 目次へ戻る」で戻れる。
    画面名・ボタン名・確認ダイアログの文言は実装(src/ui/*.bas)の文字列と
    必ず突き合わせること(この関数を編集するときも同様)。
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
    # R30波3: 目次・章末リンクの見た目(クリックできることが一目で分かる配色)。
    link_font = Font(bold=True, size=10.5, color="1155CC", underline="single")
    back_font = Font(size=9, bold=True, color="1155CC", underline="single")

    row = [3]

    def _anchor(cell, target_row, tip):
        # シート内アンカー: location はシート名をクォートする
        # ("'使い方'!A<行>")。ref は Hyperlink 代入時に自セル座標へ
        # 自動上書きされるため空文字でよい(openpyxl 3.1系の挙動)。
        cell.hyperlink = Hyperlink(ref="", location=f"'使い方'!A{target_row}", tooltip=tip)

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
        return r      # R30波3: 目次・章末リンクのアンカー行として使う

    def back_to_toc(target_row):
        # 各章末の「▲ 目次へ戻る」。note() と同じ余白規約(r+2)に合わせて
        # 次章の section() が付け足す1行と合わせ、既存の note→section 間隔
        # (2行空け)と統一する。
        r = row[0]
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=2, value="▲ 目次へ戻る")
        c.font = back_font
        c.alignment = Alignment(vertical="center", indent=1)
        _anchor(c, target_row, "クリックで目次に戻ります")
        ws.row_dimensions[r].height = 18
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

    # ---- 目次(予約のみ・書き戻しは全章を書き終えたあと) --------------------
    # section()等の行カーソルで先に本文を全部書き、確定した行番号を使って
    # 目次セルへ後からハイパーリンクを書き戻す(2パス)。ここでは目次が
    # 占める行数だけ、section()と同じ間隔規約で先に「空けて」おく。
    TOC_N = 8
    toc_header_row = row[0]                    # 目次見出し行(=3)
    row[0] = toc_header_row + 2                 # section()のheader→本文間隔を模す
    toc_first_row = row[0]                      # 目次1件目の行
    row[0] += TOC_N                             # 章の数だけ1行1件で予約(kv()と同型)
    row[0] += 1                                  # note()相当の空行を1つ確保してから次章へ

    # ==== ① これは何? + 最初に1回だけやること =============================
    ch1_row = section("① これは何?+最初に1回だけやること")
    section("これは何ですか", "355C86")
    kv("ひとことで言うと", "自分で入れた資料にAIが答えてくれる、社内版のNotebookLMです。", 0)
    kv("普通のAIとの違い", "答えの根拠になった資料名とページが必ず一緒に出ます。原文もその場で開けます。", 1)
    kv("入れられる資料", "PDF / Word / Excel / テキスト / 画面のスクリーンショット", 0)
    note("※ 契約者名・電話番号などの個人情報を含む資料は入れないでください。")

    section("はじめに 1回だけやること", "B45F06")
    step("1", "このファイルをExcelで開く。")
    step("2", "画面の上に黄色い帯で「セキュリティの警告」と出たら、その中の\n「コンテンツの有効化」ボタンを押す。", warn=True)
    step("3", "数秒待つと画面が自動で組み上がります。これで準備完了です。")
    back_to_toc(toc_header_row)

    # ==== ② 画面の説明 ======================================================
    ch2_row = section("② 画面の説明")
    section("画面は5つです", "355C86")
    kv("🏠 Hub(拠点)", "最初に出る画面。自分の記録と、他の画面への入口が並んでいます。", 0)
    kv("💬 チャット", "AIに質問する画面。ここだけで会話が完結します。", 1)
    kv("📚 マイ本棚", "資料を入れる・探す・共有する画面。上のピルで\n「🃏 ギャラリー」「📋 マイ本棚」「🎁 みんなの解決事例」の3モードを切り替えます。\n※ピル「📋 マイ本棚」は画面名と同じ文字列ですが、画面の中で一覧表示に切り替えるボタンです。", 0)
    kv("📊 ダッシュボード", "節約時間・レベル・獲得バッジなど、自分の利用状況を見る画面。", 1)
    kv("🩺 診断", "動作がおかしいときに状態を確認する画面。ヘルプ「❓」からも開けます。", 0)
    note("画面の行き来はすべてボタンで行います。左上の「← Hub」でいつでも拠点に戻れます。")

    section("質問してみる", "355C86")
    step("1", "Hubの「💬 チャットで質問する」を押す。")
    step("2", "上の白い入力らんをクリックして、知りたいことを文章で書く。\n  例) 契約者が亡くなったときの手続きを教えて")
    step("3", "「💬 質問する」を押す(Ctrlキー+Enterでも送れます)。")
    step("4", "答えの下に出る資料名のボタンを押して、元の文章を必ず確認する。")
    step("5", "役に立ったら「✅ 解決した」を押す。記録が貯まり、資料を作った人にも届きます。")
    back_to_toc(toc_header_row)

    # ==== ③ ボタン早見表 ====================================================
    ch3_row = section("③ ボタン早見表")
    # R30F2-3(敵対的レビュー2周目MINOR裁定): 実装(modKnowledgeBar.ToolButton)を
    # 確認すると、マウスホバーの説明(ScreenTip)は「📚マイ本棚」画面のツール
    # バーだけに配線されており、しかも🗑削除ボタンはR30F3で意図的に除外
    # されている(OnDeleteSourceのActiveCell.Row依存とHyperlinks.Addの
    # 相性リスクを避けるため)。チャット画面のボタンにはこの仕組み自体が
    # 無いため、「どのボタンも」という全称断言を実装の範囲に限定する。
    note("📚 マイ本棚のツールバーの「❓説明」を押すと、全ボタンの説明が一覧で出ます(もう一度カードを押すと閉じます)。")

    section("チャット画面のボタン", "355C86")
    kv("🏢 社内ナレッジ検索 ⇔ 🌐 一般アシスタント",
       "押すたびに切り替わります。社内ナレッジ検索=本棚の資料だけを見て、出典つきで答えます。\n一般アシスタント=本棚を見ずに一般知識で答えます(計算・要約・下書きに便利)。", 0)
    # R32 Fix波 F15: 「40〜60秒」は実装と食い違っていた。画面のボタン
    # (modUIMain.ModeCaption)とモード説明(modMode.Description)はどちらも
    # 「1〜2分」と出す。同ページ下部の目安(1212行目)とも揃える。
    kv("⚡ すぐ聞く / 🔍 しっかり調べる", "すぐ聞く=10〜20秒のふだん使い。しっかり調べる=1〜2分かけて、念入りに調べて答えを検証してから返します。", 1)
    kv("📎 (クリップ)", "画面のスクリーンショットを貼って、その中身について質問できます。", 0)
    kv("🔍 深掘り", "直前の会話をふまえて、続けて質問します。話の続きをもう一度書かなくても大丈夫です。", 1)
    kv("✅ 解決した", "自己解決として記録。節約時間と経験値が増え、資料を書いた人にも感謝が届きます。", 0)
    kv("🤔 微妙", "参考にはなったが不十分だったときに記録します。", 1)
    kv("❌ 違う", "正しい内容を入力すると、それを覚えて次から反映します。", 0)
    kv("📋 コピー / 📄 Word", "答えをコピー、またはWord文書として書き出します。", 1)
    kv("🗑 クリア", "会話を消して最初からやり直します(資料は消えません)。", 0)
    kv("🚪 (ドア)", "保存してこのファイルを閉じます。", 1)

    section("マイ本棚のボタン(新設)", "355C86")
    note("資料の追加・削除・共有は、すべて「📚マイ本棚」画面の上のツールバーから行います。")
    kv("➕ 登録", "文章を直接ここに書いて、ナレッジとして登録します。", 0)
    kv("📁 追加", "PDFやWordなどのファイルを資料として追加します。", 1)
    kv("⚡ 仕上げ", "昔の形式で取り込んだ資料に、章の目次と要約を後付けします(④で詳しく説明します)。", 0)
    kv("🗑 1件(赤いボタン)", "いま選んでいる1件だけを本棚から削除します(④で詳しく説明します)。", 1)
    # 2026-08-16(R33 W5-23): ツールバーへ実際に生えたボタンなので一覧へ載せる。
    kv("🗑 部門ぶん一括(赤いボタン)", "選んだ部門から取り込んだ資料を、まとめて本棚から削除します。\n"
                                   "自分で登録・追加した資料と、パックで受け取った資料は消えません。", 0)
    kv("📦 パック出力", "この本棚の中身をファイルにまとめて書き出します。", 0)
    kv("📥 パック取込", "書き出しておいた本棚ファイルを読み込みます。", 1)
    kv("🔄 同期", "共有フォルダの最新版と資料を同期します。", 0)
    kv("📂 フォルダ", "資料を自動で取り込む共有フォルダを選びます。", 1)
    kv("📡 部門チャンネル", "部門ごとの共有先を設定します(⑤で詳しく説明します)。", 0)
    kv("💡 みんなの困りごと", "回答が見つからなかった質問の一覧を見られます。", 1)
    kv("❓説明", "ツールバーの全ボタンの説明が一覧で出ます(開いた説明カードをクリックすると閉じます)。", 0)
    note("「検索」は🃏 ギャラリーのときだけ、「🗑 1件」は📋 マイ本棚(一覧表示)のときだけ出ます。\n"
         "「📤正典を発行」は発行の権限がある端末だけに、「📊利用状況」は運営の権限がある端末だけに、\n"
         "「📸スクショ取込」は画像解析が使える環境だけに出ます。")
    back_to_toc(toc_header_row)

    # ==== ④ 資料の入れ方・削除の仕方・「仕上げ」とは =========================
    ch4_row = section("④ 資料の入れ方・削除の仕方・「仕上げ」とは")
    section("資料を入れる", "355C86")
    step("1", "Hubの「📚 ナレッジと本棚」を押す。")
    step("2", "上のツールバーの「📁 追加」でファイルを選ぶ。")
    step("3", "状態が ⏳(変換中) から ✅(完了) に変わったら質問できます。")
    note("「📂 フォルダ」で共有フォルダを決めておくと、そこに置いたファイルが自動で取り込まれます\n(フォルダから消せば本棚からも消えます)。")

    section("対応形式・置き場所のルール(新設)", "355C86")
    kv("対応している形式", "txt / md / csv / pdf / docx / doc / xlsx / xls / xlsm\nこれ以外の形式は追加できません(エラーコード E0301)。", 0)
    kv("フォルダは第1階層だけ", "「📂 フォルダ」で決めたフォルダの、第1階層(直下)に置いたファイルだけを\n読みます。サブフォルダの中までは読みに行きません。取り込みたいファイルは\nフォルダの直下(サブフォルダの外)に置いてください。", 1)
    kv("同じ名前のファイルは禁止(E0504)", "資料はファイル名で見分けています。別の場所にある同じ名前のファイルを\n追加しようとすると、エラーで止まります(黙って上書きはしません)。\nファイル名に年度や版を入れて区別してください(例: 料率表_2025年度.xlsx)。", 0)
    note("同じ名前・同じ場所のファイルを更新した場合は「置き換え」になります(これは正しい動作です)。")

    section("量と速さの目安(新設)", "355C86")
    kv("画像PDF(スキャン)は特に時間がかかる", "文字の入っていないPDFは、1ページごとにAIが読み取ります。\n目安として20ページで約9分かかります(実機実測)。\nページ数が多い資料は時間に余裕をもって取り込んでください。", 0)
    kv("本棚が大きくなるほど質問も少し遅くなる", "検索は本棚の資料全体を見るためです。実装からの見積もりでは、\n〜1,000件はほぼ待ちなし/5,000件前後で数秒/20,000件前後で\n十数秒(いずれも1回目だけで、2回目以降は数秒で済みます)。", 1)
    note("上限はチャンク数20,500件(資料150〜200冊相当)です。近づくとHubに使用量の警告が出ます。")

    section("資料を消すには(新設)", "B45F06")
    step("1", "「📚 マイ本棚」画面の上のピルで「📋 マイ本棚」(一覧表示)に切り替える。")
    step("2", "削除したい資料の行をクリックする。")
    step("3", "ツールバー右端の赤い「🗑 1件」ボタンを押す。")
    step("4", "確認画面が出たら「はい」を押す。", warn=True)
    note("この操作はあとから取り消せません。行をクリックせずに削除を押すと\n「行をクリックしてから、もう一度押してください」と案内が出ます。焦らずもう一度お試しください。")

    section("「仕上げ」とは(新設)", "355C86")
    kv("仕上げって何?", "昔の形式で取り込んだ資料に、章の目次と要約を後付けする処理です。", 0)
    kv("何ができるようになる?", "『◯◯について全部教えて』のような俯瞰質問や、『第3条について』のような条文参照ができるようになります。", 1)
    kv("すでに仕上げ済みなら?", "ツールバーの「⚡仕上げ」を押しても『仕上げ済みです』と出るだけで、何も壊れません。安心して押してください。", 0)
    note("仕上げに元のファイルは要りません(再取込は不要です)。")

    section("資料の状態を表す記号", "355C86")
    kv("✅", "取り込み完了。質問に使えます。", 0)
    kv("⏳", "変換中。しばらく待ってください。", 1)
    kv("⚠️", "一部うまく読めませんでした。ページ番号がメモ欄に出ます。", 0)
    kv("🖼", "画像だけのPDFです。文字が入っていないため読み取れません。", 1)
    kv("🕒", "元のファイルが見つかりません(移動・削除された可能性)。", 0)
    back_to_toc(toc_header_row)

    # ==== ⑤ 部門とみんなの共有(4ステップ) ===================================
    ch5_row = section("⑤ 部門とみんなの共有(4ステップ)")
    section("① なぜ設定するのか", "355C86")
    note("部門=あなたが所属するグループ、共有フォルダ=資料の受け渡し場所です。\n"
         "この2つが両方そろって初めて、感謝状・部門チャンネル・みんなの節約などの機能が動き始めます。\n"
         "どちらか一方だけでは動きません。")

    section("② 部門を選ぶ", "355C86")
    step("1", "ホーム画面の、自分の名前の下にある「(部門を設定:クリック)」を押す。")
    step("2", "出てきた一覧から、自分の部門を番号で選ぶ。")

    section("③ 共有フォルダを選ぶ", "355C86")
    step("1", "画面右上の「❓」ヘルプボタンを押す。")
    step("2", "出てきたカードの「⚙ 共有フォルダ設定」を押す。")
    step("3", "共有フォルダ(部門の資料が置いてある場所)を選ぶ。")

    section("④ 両方終わると起きること", "355C86")
    # 2026-07-28(レビュー I-4)を継承: 実装に無い動作は書かない。自動購読
    # (開くだけで届く)は未実装で、実際は Hub の「更新があります」を押した
    # ときにまとめて取り込む。
    kv("部門チャンネルとは", "商品部・システム部・人事部などが公開している『正典』です。\nHubに出る案内を1回押すだけで、全部門ぶんがまとめて入ります。", 0)
    kv("何が良いのか", "自分で資料を集めなくても、開いた初日から答えが返ります。ポータルを探し回る必要がなくなります。", 1)
    kv("更新されたら", "Hubに「更新があります」と出ます。押すだけで最新版に入れ替わります。\n古い内容は入れ替えのときに消えるので、古い条文で回答されることはありません。", 0)
    kv("自分の資料は?", "消えません。部門の正典と、あなたが入れた資料の両方から答えます。", 1)
    kv("感謝状", "あなたが「✅ 解決した」を押した回答は、資料を作った人へ感謝状として届きます。", 0)
    note("部門・共有フォルダのどちらか一方だけを設定した状態では、これらは動きません。②③の両方を確認してください。")
    back_to_toc(toc_header_row)

    # ==== ⑥ 便利な使い方 ====================================================
    ch6_row = section("⑥ 便利な使い方")
    section("(a) 文脈引き継ぎ: 調べた答えをそのまま計算に使う", "355C86")
    note("社内ナレッジ検索で調べた答えの数字や内容を、そのまま一般アシスタントへ持ち越して\n計算や言い換えをさせることができます。要約・言い換え・練習問題づくりにも使えます。")
    step("1", "「🏢 社内ナレッジ検索」のまま「TMPの支払限度額は?」のように質問する。")
    step("2", "答えが出たら、ボタンを押して「🌐 一般アシスタント」に切り替える。")
    step("3", "続けて「請求が8万円なら、支払いはいくらになる?」のように質問する。")
    step("4", "直前の答えの数字を使って計算してくれます。")

    section("(b) 深掘りボタン", "355C86")
    kv("🔍 深掘り", "直前の会話をふまえて、続けて質問できます。話の続きをもう一度書かなくても大丈夫です。", 0)

    section("(c) 逆質問には番号で答える", "355C86")
    note("AIから『どちらについてですか? 1)… 2)…』のように番号つきで聞き返されることがあります。")
    step("1", "チャット入力欄に、選びたい番号だけを書く(例:「2」)。")
    step("2", "「💬 質問する」を押す。")
    note("番号でなく、質問を書き直して送っても大丈夫です。")
    back_to_toc(toc_header_row)

    # ==== ⑦ 困ったとき + 検索のコツ ==========================================
    ch7_row = section("⑦ 困ったとき+検索のコツ", "9C1F1F")
    section("困ったときは", "9C1F1F")
    kv("画面が崩れた・ボタンが消えた", "Hub右上の「🔄」を押すと画面を描き直します。", 0)
    kv("答えが途中で止まる", "いったん保存して閉じ、開き直してからもう一度お試しください。", 1)
    kv("それでも直らない", "ヘルプ「❓」→「診断」の画面をスクリーンショットで撮って管理者へ送ってください。", 0)
    kv("操作を思い出したい", "ヘルプ「❓」→「ツアーをもう一度見る」で、最初の案内を再表示できます。", 1)
    note("自己判断で設定を変える必要はありません。まずスクリーンショットを送ってください。")

    section("検索のコツ(新設)", "9C1F1F")
    kv("いちばん確実な聞き方", "資料に書かれているのと同じ言葉・同じ表記で聞くことです。", 0)
    kv("言い換えはどこまで通じる?", "資料に『クマ』とあるとき「熊」「Bear」のような言い換えもAIがある程度は補いますが、必ず見つかる保証ではありません。", 1)
    kv("0件だったら", "資料の中で実際に使われていそうな言葉に言い換えて、もう一度聞いてみてください。", 0)
    back_to_toc(toc_header_row)

    # ==== ⑧ よくある質問(新設) =============================================
    # R32波3(実機第17報④・班D調査): 実機からの「もっと易しくより細かく」
    # 要望を受け新設。文言はすべて modLog.FriendlyMessage / modUIShelf /
    # modKnowledgeBar 等の実装から引用し、推測でFAQを作らない(CLAUDE.md
    # 「書く内容は必ずコードで裏取りしてから書く」)。
    ch8_row = section("⑧ よくある質問", "9C1F1F")
    section("よくある操作の疑問", "9C1F1F")
    kv("削除ボタンが効かない",
       "「🗑 1件」は📋マイ本棚(一覧表示)で、削除したい資料の行をクリックしてから\n"
       "押すボタンです。🃏ギャラリー表示には削除ボタン自体が出ません(押し間違えを防ぐため)。\n"
       "行を選ばずに押すと「削除したい資料の行をクリックしてから、もう一度押してください。」\n"
       "と出ます。この案内が出たら、まず表示を📋マイ本棚に切り替えてください。", 0)
    kv("取込が遅い(特に画像PDF)",
       "文字の入っていない画像PDF(スキャン)は、1ページごとにAIが読み取ります。\n"
       "目安は20ページで約9分(実機実測)です。ページ数が多い資料ほど時間がかかります。", 1)
    kv("回答に3〜5分かかる",
       "「入念に調べる」は6段階(要点整理→下書き→自己批判→検証→出典突合まで)を\n"
       "必ず全部通すため、実機では3〜5分ほどかかることがあります(表示上の目安は2〜4分)。\n"
       "急ぐときは「⚡すぐ聞く」(10〜20秒)や「🔍しっかり調べる」(1〜2分)をお使いください。", 0)
    kv("資料が見つからない(E0601)",
       "「本棚の中に手がかりが見つかりませんでした」と出ます。資料が本棚に入っているか\n"
       "確認してください。言い換え(表記ゆれ)は辞書である程度補いますが、必ず見つかる\n"
       "保証はありません。資料で実際に使われていそうな言葉に言い換えて再検索してください。", 1)
    # 2026-08-16(R33H F28): 利用者向けページから設定キー名 admin_users を外す。
    # 📊利用状況は「匿名を約束して集めた投書の全文」が出る画面で、この
    # ページは全利用者が読む。誰の端末に出るかは「運営の権限がある端末」で
    # 用が足り、鍵の名前まで降ろす必要はない(管理者ページ側には残す)。
    kv("ボタンが出ない",
       "「📸スクショ取込」は画像解析が使える環境だけに、\n"
       "「📤正典を発行」は発行の権限がある端末だけに出ます。「📊利用状況」は\n"
       "運営の権限がある端末に出ます。運営の担当者がまだ1人も決まっていない\n"
       "間は、発行担当の端末にも出ます。いずれも「押しても断られるボタン」を\n"
       "見せない設計です。", 0)
    # 2026-08-16(R33 W5-26): 波3(W3-1)で「Excelの行内の空セルを詰めるせいで
    # 列がずれ、表の値が別の見出しの値として取り込まれる」を直したが、
    # 直ったのは【これから取り込む資料】だけで、既に取込済みのデータは
    # ずれたまま本棚に残る。コードでは救済できないので、ここで案内する。
    kv("Excel資料の検索結果がおかしい",
       "Excelの表を取り込むしくみを改良し、空欄のあるセルで列がずれて\n"
       "別の見出しの値として覚えてしまう問題を直しました。ただし直ったのは\n"
       "これから取り込む資料だけで、すでに本棚に入っている資料は古いまま残ります。\n"
       "Excelの資料で答えがかみ合わないときは、その資料を一度「🗑 1件」で消してから\n"
       "「📁 追加」で取り込み直してください。", 1)
    note("上記以外の疑問は、まずヘルプ「❓」→「診断」の画面をスクリーンショットで撮って管理者へ送ってください。")

    section("エラーコード表(主なもの)", "9C1F1F")
    note("表示された案内文の末尾に「(コード: E0xxx)」がある場合、その番号をこの表で調べられます。\n症状/原因/対処の順で書いています。文言は実装(modLog.FriendlyMessage)からの引用です。")
    err_rows = [
        ("E0101 起動直後に使えない",
         "原因: 設定ファイル(configシート)に必要な項目が見つからない。\n対処: 配布元にこのファイルの再入手を依頼してください。"),
        ("E0102 検索・同期・集計が動かない",
         "原因: この端末では検索機能に必要な部品(Windows Script Runtime)が利用できない。\n対処: 管理者へご連絡ください。"),
        # 2026-08-16(R33 W5-22): E0201/E0204/E0701/E0801/E0808 の5件を追加。
        # いずれも「(コード: E0xxx)」付きで利用者の画面に出るのに、この表に
        # 一度も載っていなかった(= 案内どおり番号を引いた人が空振りする)。
        # ・E0201/E0204 … modGateway が "#ERR:E0201/E0204" を返し、
        #   modRagParse.BuildErrorAnswer がコードを付けて回答欄へ出す
        # ・E0701 … modPack のパック取込失敗(利用者操作は非silent側が本線)
        # ・E0801 … 起動失敗ダイアログ(modBoot)と modLog.ShowError。src全体で
        #   最頻出のコードで、いちばん詰む場面(起動不能)に出る
        # ・E0808 … 正典発行の中断・ロールバック(発行担当の端末)
        # 表に載せていない E0702/E0705/E0806/E0807/E0901 は err_log 専用で、
        # 利用者の画面に「(コード: …)」付きでは出ない(= 引かれることが無い)。
        ("E0201 AIリボンが見つからない",
         "原因: AIリボン(ChatGPT関数)が使えるExcelで開かれていない。\n対処: AIリボン入りのExcelで開き直すか、ヘルプ「❓」→「診断」で状態を確認してください。"),
        ("E0202 質問がエラーで返る",
         "原因: AIとの通信が混み合っている、または実行中の処理と操作が重なった(操作は誤りではない)。\n対処: 少し待ってからもう一度。続くなら診断結果を管理者へ。"),
        ("E0203 取込中にエラー",
         "原因: 文章をAIが読める形に変換する処理(埋め込み取得)に失敗。\n対処: 時間を置いてもう一度。続くなら診断結果を管理者へ。"),
        ("E0204 今日はAIに問い合わせできない",
         "原因: 本日のAI利用上限に達した可能性がある。\n対処: 今日はここまでにして、明日「🔄同期」を押せば続きから再開できます。"),
        ("E0301 ファイルを追加できない",
         "原因: 対応していない種類のファイル。\n対処: 対応形式(txt/md/csv/pdf/docx/doc/xlsx等)のファイルをお使いください。"),
        ("E0302 ファイルを開けない",
         "原因: 他のアプリで開いている、または権限が無い可能性。\n対処: ファイルを閉じてもう一度。PDFはこのファイルと同じ場所のGhostscriptフォルダも確認。"),
        ("E0303 画像PDFの文字が読めない",
         "原因: 画像として保存されたPDF(文字が埋め込まれていない)。\n対処: 画像解析(config feature_vision)を有効にしGhostscriptを配置(43_画像PDFのOCR取込設定を参照)。"),
        ("E0304 ファイルの中身が空",
         "原因: ファイルは開けたが中身の文章が空(保護されている可能性)。\n対処: 別の資料でお試しください。"),
        ("E0401 登録できた内容が0件",
         "原因: 文字数が少なすぎる資料の可能性。\n対処: 別の資料でお試しください。"),
        ("E0501 本棚の上限を超えている",
         "原因: 本棚に入れられる資料(チャンク数)の上限に到達。\n対処: 使っていない資料を削除。または管理者にshelf_max_chunksの拡大を相談。"),
        ("E0502 同期フォルダが見つからない",
         "原因: 同期するフォルダが見つからない。\n対処: 「📂フォルダ」からフォルダを選び直してください。"),
        ("E0503 今は操作できない",
         "原因: 今、別の取込処理が実行中。\n対処: 処理が完了するまで少々お待ちください。"),
        ("E0504 同じ名前の資料がある",
         "原因: 同じ名前の資料が、別の場所からすでに登録されている。\n対処: ファイル名を変えるか、先に元の資料を削除してから追加し直してください。"),
        ("E0601 資料が見つからない",
         "原因: 本棚の中に手がかりが見つからない。\n対処: 資料が本棚に入っているか確認、または資料を追加してから試してください。"),
        ("E0602 回答をまとめられない",
         "原因: こちら側の処理の都合(あなたの操作の誤りではない)。\n対処: もう一度送信。続く場合は質問の言い回しを少し変えてみてください。"),
        ("E0701 パックを取り込めない",
         "原因: パックの形式が正しくない、またはバージョンが合っていない。\n対処: 渡してくれた相手に、パックの再書き出しを依頼してください。"),
        ("E0703 発行が中止された",
         "原因: 個人情報チェック(config pii_scan_enabled。既定オフ)で該当の文字列を検知し、書き出しを無条件で中止した。\n対処: 該当の資料を本棚から外すか内容を修正してから、もう一度発行してください。"),
        ("E0801 画面の組み立てに失敗した",
         "原因: 画面を描くところで予期しない問題が起きた(データは失われていません)。\n対処: ブックを一度閉じて開き直せば元どおり使えます。繰り返すときは画面を撮影して管理者へ。"),
        ("E0805 結果が保存されない",
         "原因: このファイルが読み取り専用で開かれている。\n対処: 既に別のExcelで同じファイルを開いていないか確認してください。"),
        ("E0808 正典の発行が中断した",
         "原因: 発行の処理中に予期しない問題が起きた(発行担当の端末だけに出ます)。\n対処: もう一度お試しください。繰り返すときは共有フォルダの接続と書き込み権限をご確認ください。"),
        ("E0904 作業用Excelが開かない",
         "原因: 端末ポリシーでWScript.Shellが使えない。\n対処: タスクバーのExcelをAltキーを押しながらクリックしても開けます。"),
        ("E0905 社内ポータルが開かない",
         "原因: この端末からブラウザを開く手段が両方とも使えない。\n対処: アドレスはクリップボードにコピー済みです。ブラウザのアドレス欄に貼り付け(Ctrl+V)てください。"),
    ]
    for i, (k, v) in enumerate(err_rows):
        kv(k, v, i)
    back_to_toc(toc_header_row)

    # ==== 付録(目次には含めない・コンプライアンス上の既存記載を維持) ========
    section("付録: このツールが記録していること", "5B6B7B")
    kv("この端末に残るもの", "あなたの質問と回答の履歴(最新100件)。この端末の中だけです。", 0)
    kv("部内で共有されるもの", "「解決した」を押した質問と答え / 答えが見つからなかった質問 /\n利用回数・解決率などの集計値。改善のために使います。", 1)
    kv("共有されないもの", "質問の全文は共有されません(先頭40文字までの要約のみ)。", 0)
    kv("匿名で意見を送る", "Hub右上の 📮 から、名前を残さずに感想・不満を送れます。", 1)
    note("記録を止めたい場合は config シートの telemetry_enabled を FALSE にしてください。")

    section("このシートについて")
    note("このシートだけは、マクロが無効な状態でも読めるようにしてあります\n(マクロ有効化の案内は、ここでしか出せないためです)。\nマクロを有効にして開き直すと、実際の操作画面が使えるようになります。")

    # R31波3(実機第16報F-C④・B-3): 正典発行の担当者向けページへの導線。
    # 目次には足さない(TOC_N不変で安全)。末尾章のさらに後ろへ1行だけ足す。
    r = row[0]
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    c = ws.cell(row=r, column=2, value="\U0001F511 発行を担当する方はこちら →")
    c.font = link_font
    c.alignment = Alignment(vertical="center", indent=1)
    c.hyperlink = Hyperlink(ref="", location="'管理者向け'!A1", tooltip="部門の正典を発行する担当者向けの説明を開きます")
    ws.row_dimensions[r].height = 20
    row[0] = r + 2

    # ---- 目次の書き戻し(全章の行番号が確定したのでここでハイパーリンク化) ----
    ws.merge_cells(start_row=toc_header_row, start_column=1, end_row=toc_header_row, end_column=3)
    hc = ws.cell(row=toc_header_row, column=1, value="目次(クリックすると各章へ移動します)")
    hc.font = head_font
    hc.fill = PatternFill("solid", fgColor="1F4E78")
    hc.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[toc_header_row].height = 24

    toc_entries = [
        (ch1_row, "① これは何?+最初に1回だけやること"),
        (ch2_row, "② 画面の説明"),
        (ch3_row, "③ ボタン早見表(チャット/マイ本棚)"),
        (ch4_row, "④ 資料の入れ方・削除の仕方・「仕上げ」とは"),
        (ch5_row, "⑤ 部門とみんなの共有(4ステップ)"),
        (ch6_row, "⑥ 便利な使い方(文脈引き継ぎ・深掘り・逆質問)"),
        (ch7_row, "⑦ 困ったとき+検索のコツ"),
        (ch8_row, "⑧ よくある質問(エラーコード表つき)"),
    ]
    if len(toc_entries) != TOC_N:
        # 目次の予約行数(TOC_N)と実際の章数が食い違うと、以降の章の見出しが
        # 目次の下に食い込む(気づきにくい表示崩れ)。ビルド時点で確実に止める。
        raise BuildError(
            f"_make_howto: 目次の予約行数({TOC_N})と実際の章数({len(toc_entries)})が不一致です")
    for i, (target_row, title) in enumerate(toc_entries):
        r = toc_first_row + i
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=2, value=title)
        c.font = link_font
        c.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        if i % 2 == 1:
            c.fill = zebra_fill
        _anchor(c, target_row, "クリックでこの章へ移動します")
        ws.row_dimensions[r].height = 20

    ws.sheet_properties.tabColor = TAB_COLORS["使い方"]
    ws.sheet_view.showGridLines = False
    return ws


def _make_admin_guide(wb):
    """管理者向けシート: 正典(部門の公式ナレッジ)を発行する担当者むけの説明。

    R31波3(実機第16報F-C④): 発行担当は非エンジニアの現場社員であり、
    「発行」「巻き戻し」の実際の挙動(何が起きるか・何を確認すればよいか)を
    知らないまま使うと事故になる。_make_howto と同型の2パス目次
    (TOC予約→章書き→Hyperlink書き戻し)を踏襲し、体裁(見出し色・ステップ・
    早見表・注記)も使い方ページのヘルパーをそのまま流用する。

    内容の一次資料は docs/70_データ管理者向け_置き場所と運用.md §B と、
    src/ui/modPublishUI.OnPublish / src/pack/modPublish.bas / modPackExport.bas /
    modChannel.bas の実装コメント・ダイアログ文言。合言葉(publish_key)の
    実際の値・環境変数名(MYBOOKSHELF_PUBLISH_KEY)はこのページに一切書かない
    (このシートは一般配布物にも同梱されるため)。
    """
    ws = wb.create_sheet("管理者向け")
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 78

    _banner(ws, f"{APP_TITLE} — 管理者向け(正典発行の手引き)")

    head_font = Font(bold=True, size=12, color="FFFFFF")
    key_font = Font(bold=True, size=10, color="1F4E78")
    body_font = Font(size=10)
    step_font = Font(bold=True, size=10, color="FFFFFF")
    step_fill = PatternFill("solid", fgColor="7F9DB9")
    warn_fill = PatternFill("solid", fgColor="FFF4D6")
    zebra_fill = PatternFill("solid", fgColor="F4F6FB")
    thin = Side(style="thin", color="D9E1EC")
    link_font = Font(bold=True, size=10.5, color="1155CC", underline="single")
    back_font = Font(size=9, bold=True, color="1155CC", underline="single")

    row = [3]

    def _anchor(cell, target_row, tip, sheet="管理者向け"):
        cell.hyperlink = Hyperlink(ref="", location=f"'{sheet}'!A{target_row}", tooltip=tip)

    def section(title, fill="1F4E78"):
        if row[0] > 3:
            row[0] += 1
        r = row[0]
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        c = ws.cell(row=r, column=1, value=title)
        c.font = head_font
        c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[r].height = 24
        row[0] = r + 2
        return r

    def back_to_toc(target_row):
        r = row[0]
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=2, value="▲ 目次へ戻る")
        c.font = back_font
        c.alignment = Alignment(vertical="center", indent=1)
        _anchor(c, target_row, "クリックで目次に戻ります")
        ws.row_dimensions[r].height = 18
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

    def diagram(text):
        # R32波3(⑥ファイルの置き場所・新設): フォルダ構造図をセル内で
        # そのまま読める形にする。wrap_text はONのままだが、行の折り返しに
        # 頼らず自分で改行しているので体裁が崩れない(等幅フォントで揃える)。
        r = row[0]
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=2, value=text)
        c.font = Font(size=9.5, name="Consolas", color="1F4E78")
        c.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        c.fill = zebra_fill
        for side in (Border(top=thin, bottom=thin, left=thin, right=thin),):
            c.border = side
        ws.row_dimensions[r].height = max(18, 14 * (text.count("\n") + 1) + 8)
        row[0] = r + 2

    # ---- 目次(予約のみ・書き戻しは全章を書き終えたあと) --------------------
    TOC_N = 9
    toc_header_row = row[0]
    row[0] = toc_header_row + 2
    toc_first_row = row[0]
    row[0] += TOC_N
    row[0] += 1

    # ==== ① 正典ナレッジとは ================================================
    ch1_row = section("① 正典ナレッジとは")
    kv("正典って何?", "部門で確認済みの知識を1つだけ発行し、部門の全員がそれを受け取る仕組みです。", 0)
    kv("なぜ必要?", "全員が自分の本棚にバラバラに資料を集めると、同じことを何度も調べたり、\n"
                    "古い情報のまま止まっている人が出たりします。正典を1つ発行すれば、\n"
                    "開いた初日から全員が同じ答えにたどり着けます。", 1)
    kv("誰が発行する?", "部門で決めた発行担当だけです(このページを開いているあなたです)。", 0)
    note("このページは誰でも読めます(発行そのものには合言葉と発行者用ファイルが必要なので、読めても発行はできません)。\n"
         "ただし「📤正典を発行」ボタンは一般の利用者には表示されません。")
    back_to_toc(toc_header_row)

    # ==== ② 発行者になるには(新設) ==========================================
    # R32波3(実機第17報④・ユーザー裁定=発行者用ファイルは個別に渡す/手順は
    # 別途連絡): configシート直接編集の手順・合言葉の値・環境変数名は
    # 一切書かない(誰でも発行者になれてしまうため)。「発行ボタンが出るかは
    # 合言葉の有無だけで決まる」という仕組みの説明は、手順を書かない限り
    # 悪用できないため書いてよい(司令塔裁定)。
    ch2_row = section("② 発行者になるには", "B45F06")
    kv("発行担当用のファイルが必要", "「📤正典を発行」ボタンは、発行担当用のファイル\n"
       "(MyBookshelf_発行者用.xlsm)にだけ表示されます。通常配布のファイルには\n"
       "最初から出ません(ボタンそのものが無いので、探しても見つかりません)。", 0)
    kv("入手方法", "配布元(このツールの管理担当)から、発行担当用ファイルを個別に受け取って\n"
       "ください。受け取り方(渡し方)は別途ご連絡します。", 1)
    kv("仕組み(参考)", "発行ボタンが出るかどうかは、そのファイルに「合言葉(発行キー)」が\n"
       "設定されているかだけで決まります。発行担当用ファイルには、あらかじめ\n"
       "合言葉が設定済みです。設定のしかた自体はこのページには書きません。", 0)
    kv("ファイルを人に渡すと?", "発行担当用ファイルを渡した相手も、その場で発行できる状態になります。\n"
       "取り扱いは慎重にお願いします(④注意点も参照)。", 1)
    # 2026-08-16(R33 W5-25/W5-27): admin_users の役割が「組織的な除外の解除」
    # だけではなくなった(📊利用状況の閲覧が加わった)ので、実態に合わせる。
    kv("「発行者」と「組織管理者」は別物",
       "publish_key(発行担当・正典を発行できる人)と admin_users(組織管理者・\n"
       "資料の組織的な除外を解除できる人、📊利用状況を見られる人)は、名前は\n"
       "似ていますが別の権限です。発行担当だからといって組織管理者の操作は\n"
       "できません。ただし組織管理者が1人も登録されていない間に限り、\n"
       "📊利用状況だけは発行担当の端末でも開けます(誰も運営状況を確認\n"
       "できない状態を避けるためです)。", 0)
    # 2026-08-16(R33 W5-24): ここは「configシートを直接編集して自分で発行者に
    # なる方法はありません。」と断定していたが、事実に反する(設定シートは
    # Excelの標準機能で再表示でき、ブック構造の保護もかかっていない)。
    # 配布物の可視ドキュメントが存在しない防御を約束している状態を止めるのが
    # 目的で、手順を教えることではないため、具体的な操作は一切書かない。
    # (R33波2の裁定: 設定シートを隠す対策は①正規の設定変更手順を壊す
    #  ②起動のたびに元へ戻るので実効性が無い ③ブック構造保護は画面遷移を
    #  無言で失敗させる、の3点で不成立と確定している)
    note("この仕組みは善意の運用を前提にしており、技術的に完全に防いではいません。"
         "アプリ側の関門は誤操作を防ぐためのもので、部門の正典を本当に守るのは共有フォルダ側の"
         "アクセス権限(channels\\ フォルダを書き込める人を絞ること)です。"
         "発行担当を限定したい組織では、必ず共有フォルダの権限設定とあわせて運用してください。")
    back_to_toc(toc_header_row)

    # ==== ③ 発行のしかた ====================================================
    ch3_row = section("③ 発行のしかた")
    step("1", "「📚 マイ本棚」に取り込んだ資料の内容を確認し、発行できる状態(確認できたQ&Aや要点)に整えておく。")
    step("2", "ツールバーの「📤 正典を発行」ボタンを押す。")
    step("3", "発行する部門名を入力する(例: 商品部)。\n更新のときも必ず前と同じ名前を入れる。")
    step("4", "発行キー(合言葉)を入力する。分からないときは管理担当に確認する。", warn=True)
    step("5", "画面に出る「発行する件数」を確認してから「はい」を押す。")
    kv("発行されるのは?", "自分の本棚に取り込んだ資料です。他部門から受け取った正典は自動で除外されます\n"
                          "(機械的な「確認済み」フィルタはありません)。発行前に内容を確認し、\n"
                          "確認できたQ&Aと要点だけに整えておくのは発行担当の大切な仕事です。", 0)
    # R32波2(2026-08-14 W2-1)のPIIオフスイッチ既定化を受け、W3-3で訂正。
    # 「自動で走る」と言い切ると、既定FALSEの現状(何も走らない)と矛盾する。
    kv("個人情報のチェック", "config pii_scan_enabled で有効にしたときだけ、発行の直前に走ります\n"
                            "(既定はオフ=何もチェックしません)。有効なときに個人情報らしき\n"
                            "内容が見つかると、発行そのものが自動で中止されます。", 1)
    kv("同時発行だったら?", "先に発行している人がいると発行は始まらず、その旨のメッセージが出ます。\n"
                            "数分おいてからもう一度お試しください。", 0)
    note("発行は数十秒かかります。進捗はステータスバーに表示され、終わると「発行しました」の画面が出ます。")
    back_to_toc(toc_header_row)

    # ==== ④ 間違えたときの巻き戻し ==========================================
    ch4_row = section("④ 間違えたときの巻き戻し")
    step("1", "もう一度「📤 正典を発行」を押し、部門名を入力して確認画面まで進む。")
    step("2", "確認画面で「いいえ」を選ぶ(「このまま発行しますか?」の問いに対して)。")
    step("3", "戻す版が表示されるので、間違いなければ「OK」を押す。")
    note("巻き戻しは数十秒で終わります。戻したことも、発行と同じように部内の全員へ次回起動時に自動で配信されます\n"
         "(間違った内容は各PCから自動で消えます)。")
    back_to_toc(toc_header_row)

    # ==== ⑤ 注意点 ===========================================================
    ch5_row = section("⑤ 注意点", "B45F06")
    kv("原文を丸ごと正典にしない", "内容を確認したQ&Aや要点だけを発行してください。分厚い資料をまるごと\n"
                                  "正典にすると、1部門でチャンク数を大量に消費して全体が破綻します。\n"
                                  "原文がそのまま必要な人は、各自の本棚に個別で追加してもらってください。", 0)
    kv("共有フォルダへ手で置かない", "正典の実体(pack.xlsx・version.txt)は「📤正典を発行」ボタンだけが作ります。\n"
                                    "共有フォルダへ手作業でファイルを置いたり書き換えたりしないでください。", 1)
    kv("合言葉とファイルの管理", "発行キー(合言葉)と、発行に使うこのファイルは発行担当だけが持ってください。\n"
                                "他の人に渡ると、その人も部門の正典を書き換えられる状態になります。", 0)
    kv("発行者名の設定", "初回起動でお名前(pack_author)を入力しておいてください。空欄のままだと\n"
                        "発行物に「不明」という名前で記録されます。", 1)
    back_to_toc(toc_header_row)

    # ==== ⑥ ファイルの置き場所(新設) ========================================
    # R32波3(実機第17報④・班D調査): 配布物のどこにも共有フォルダの構造が
    # 書かれていなかった実機指摘への対応。共有フォルダ設定画面(❓ヘルプ→
    # ⚙共有フォルダ設定)の実装は src/ui/modHelp.bas OnShareSetup / PickSharePath、
    # フォルダ構造は modShareRule.StandardSubDirs と modChannel.bas の定数
    # (PACK_NAME/VER_NAME/ARCHIVE_DIR)・modPublish.bas LOG_NAME で確認済み。
    ch6_row = section("⑥ ファイルの置き場所", "355C86")
    kv("共有フォルダの設定はどこから?", "❓ヘルプ →「⚙ 共有フォルダ設定」から選びます。フォルダ選択の画面が\n"
       "開き、キャンセルすると共有パス(\\\\サーバー名\\共有\\…)を直接入力する\n"
       "こともできます。存在しないパスは保存できません(確認してから保存します)。", 0)
    diagram(
        "共有フォルダの中身(初回に自動で作られます)\n"
        "<共有フォルダ>\\\n"
        "  ├ channels\\<部門名>\\  … pack.xlsx / version.txt / publish_log.txt / _archive\\\n"
        "  └ thanks\\ noise\\ insight\\ telemetry\\ board\\ questions\\"
    )
    kv("権限のかけ方", "共有フォルダ全体は、全員が読み書きできる必要があります。\n"
       "channels\\<部門名>\\ への書き込みだけは発行担当に絞るのが望ましいです\n"
       "(誤って正典を書き換えられないため)。この絞り込みはWindowsのフォルダ\n"
       "権限で設定します。IT管理者への依頼が必要です(アプリ側に権限を制御する\n"
       "機能はありません)。", 1)
    kv("共有フォルダの自動整理", "発行担当のファイルを持つ人が、このアプリを定期的に起動すると、共有\n"
       "フォルダの古いファイルが自動で整理されます。発行担当が長期不在だと、\n"
       "共有フォルダのファイルが増え続けることがあります。", 0)
    back_to_toc(toc_header_row)

    # ==== ⑦ 運用のコツ =======================================================
    ch7_row = section("⑦ 運用のコツ")
    kv("💡 みんなの困りごと", "答えが見つからなかった質問の一覧です。ここに並ぶ内容に答える資料を\n"
                             "本棚に登録し、内容を確認してから発行すると、部門の正典がどんどん育ちます。", 0)
    # 2026-08-16(R33 W5-25/W5-27): 関門を発行キーから組織管理者(config
    # admin_users)へ寄せた。ただし admin_users が未設定の組織では誰も開けなく
    # なるため、未設定のときに限り発行担当の端末にも出す。
    kv("📊 利用状況", "運営向けの画面です。使われ方の集計と、名前の記録されない匿名の声を\n"
                     "確認できます。組織管理者(config admin_users)に登録された端末で開けます。\n"
                     "組織管理者がまだ1人も登録されていない間は、発行担当の端末でも開けます。", 1)
    note("💡みんなの困りごとは誰でも見られます。📊利用状況は運営の権限がある端末にだけ表示されます。")
    # R34 C(2026-08-20 裁定): max_context_chars の既定を 40,000 → 60,000 へ上げた。
    # 上げれば厚くなるが遅くもなる設定なので、「戻し方」を先に書いておく。
    kv("1回に読ませる資料の量(max_context_chars)",
       "configシートの max_context_chars は、1つの質問でAIに読ませる資料本文の\n"
       "合計文字数の上限です。今回の版から既定は 60,000 文字(以前は 40,000)。\n"
       "多く読ませるほど答えは厚く・取りこぼしが減りますが、待ち時間は伸びます。\n"
       "この設定は質問のときだけでなく、資料を取り込むときの章ごとの要約でも\n"
       "使われる共通の設定です。上げると取り込みも重くなる点にご注意ください。", 0)
    kv("回答が遅い・固まるときの戻し方",
       "configシートを開き、A列から max_context_chars の行を探して、値の欄を\n"
       "40000 に書き換えて保存し、ブックを開き直してください(それだけで元の\n"
       "動きに戻ります)。逆に、もっと精度を上げたい場合は 80000 までを上限に\n"
       "1段ずつ(60000 → 70000 → 80000)上げ、そのつど実際の質問で待ち時間を\n"
       "確かめてから次の段へ進んでください。いきなり大きくしないことが大切です。", 1)
    back_to_toc(toc_header_row)

    # ==== ⑧ よくある質問(新設) =============================================
    # R32波3(実機第17報④・班D調査): 文言は src/pack/modPublish.bas
    # (VerifyKey=vbBinaryCompare完全一致・AcquireLock/E0807)、
    # src/ui/modPublishUI.bas(OnPublish/DoRollback)、src/pack/modChannel.bas
    # (ChunkLimit既定20000・IsBudgetTight 80%・shelf_max_chunks既定20500)、
    # src/ui/modHubStat.bas(更新案内)、src/core/modLog.bas(E0807)から確認。
    ch8_row = section("⑧ よくある質問", "9C1F1F")
    kv("発行ボタンが出ない", "発行担当用のファイル(MyBookshelf_発行者用.xlsm)を使っているか確認して\n"
       "ください。通常配布のファイルにはボタン自体が出ません。", 0)
    # R32 Fix波 F15: 「余分な空白にも注意して」は事実に反する。照合する
    # modPublish.VerifyKey は入力・設定値の両方に Trim$ を掛けるため、
    # 前後の空白は無視される(区別されるのは大文字・小文字と半角/全角)。
    kv("合言葉が違うと言われる", "発行キーは大文字・小文字を区別した完全一致で照合されます。前後の半角空白は\n"
       "自動で取り除かれますが、半角/全角の違いは別の文字として扱われます。\n"
       "落ち着いてもう一度入力してください。", 1)
    kv("発行したのに他の人に届かない", "自動では配られません。購読側がHubに出る「(部門)に更新があります」の\n"
       "案内を押してはじめて取り込まれます(自動プッシュではありません)。届いて\n"
       "いない人には、その案内を押してもらってください。", 0)
    kv("間違えて発行してしまった", "もう一度「📤正典を発行」を押し、確認画面で「いいえ」を選ぶと、直前の\n"
       "版に戻す操作に進みます。数十秒で完了し、全員へ次回起動時に自動で配信\n"
       "されます(④間違えたときの巻き戻しを参照)。", 1)
    kv("部門名を変えたい", "部門名(チャンネル名)をあとから変更する機能はありません。新しい名前で\n"
       "発行すると別の部門として扱われ、古い名前の正典と共存します。", 0)
    kv("容量が足りないと言われる", "本棚のチャンク数が chunk_limit(既定20,000)の8割を超えると警告が出ます。\n"
       "shelf_max_chunks(既定20,500)が実際のハード上限です。どちらもconfig\n"
       "シートの値で、大きくすることもできます。", 1)
    kv("E0807が出て発行できない", "発行の見張り(publish.lock)を扱えませんでした、というエラーです。共有\n"
       "フォルダへの書き込み権限や端末の時計のずれを確認してください。繰り返す\n"
       "場合は共有フォルダの channels\\<部門名>\\publish.lock を手動で削除すると\n"
       "再開できます。", 0)
    back_to_toc(toc_header_row)

    # ==== ⑨ 困ったとき =======================================================
    ch9_row = section("⑨ 困ったとき", "9C1F1F")
    kv("「📤正典を発行」ボタンが出ない", "発行キーが設定されていない端末です。発行担当用に配られたファイルを\n"
                                        "使っているか確認してください。", 0)
    kv("合言葉が違うと言われる", "発行キーの入力を間違えている可能性があります。管理担当に確認のうえ、\n"
                                "もう一度落ち着いて入力してください。", 1)
    kv("発行が終わらない・進まない", "共有フォルダへの書き込み権限、社内ネットワーク(VPN)への接続を\n"
                                    "確認してください。同時に別の人が同じ部門を発行中の場合は、\n"
                                    "数分待ってからもう一度お試しください。", 0)
    note("解決しない場合は、ヘルプ「❓」→「診断」の画面をスクリーンショットで撮って、開発担当へ送ってください。")
    back_to_toc(toc_header_row)

    # ---- 目次の書き戻し ----
    ws.merge_cells(start_row=toc_header_row, start_column=1, end_row=toc_header_row, end_column=3)
    hc = ws.cell(row=toc_header_row, column=1, value="目次(クリックすると各章へ移動します)")
    hc.font = head_font
    hc.fill = PatternFill("solid", fgColor="1F4E78")
    hc.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[toc_header_row].height = 24

    toc_entries = [
        (ch1_row, "① 正典ナレッジとは"),
        (ch2_row, "② 発行者になるには"),
        (ch3_row, "③ 発行のしかた"),
        (ch4_row, "④ 間違えたときの巻き戻し"),
        (ch5_row, "⑤ 注意点"),
        (ch6_row, "⑥ ファイルの置き場所"),
        (ch7_row, "⑦ 運用のコツ"),
        (ch8_row, "⑧ よくある質問"),
        (ch9_row, "⑨ 困ったとき"),
    ]
    if len(toc_entries) != TOC_N:
        raise BuildError(
            f"_make_admin_guide: 目次の予約行数({TOC_N})と実際の章数({len(toc_entries)})が不一致です")
    for i, (target_row, title) in enumerate(toc_entries):
        r = toc_first_row + i
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=2, value=title)
        c.font = link_font
        c.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        if i % 2 == 1:
            c.fill = zebra_fill
        _anchor(c, target_row, "クリックでこの章へ移動します")
        ws.row_dimensions[r].height = 20

    # ---- 使い方ページへ戻る導線(先頭付近) ----
    home_r = toc_header_row - 1 if toc_header_row > 3 else 2
    hb = ws.cell(row=home_r, column=1, value="◀ 使い方ページへ戻る")
    hb.font = back_font
    hb.alignment = Alignment(vertical="center", indent=1)
    _anchor(hb, 1, "クリックで使い方ページへ移動します", sheet="使い方")
    ws.row_dimensions[home_r].height = 18

    ws.sheet_properties.tabColor = TAB_COLORS["管理者向け"]
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


def _vba_src_modules(present_modules):
    """vba_src シートへ載せる対象モジュールだけを返す。
    クラスモジュール(VBComponents.Add(1)で追加できない)と
    台帳で vba_src=False にされたものは対象外。
    _make_vba_src と verify_build(FA-R23-1c)で同じ判定を使うための共通化。"""
    return [m for m in present_modules
            if m.get("type") != "class" and m.get("vba_src") is not False]


def _vba_src_text(root, m):
    """src/ の .bas から vba_src セルへ格納する文字列を作る(ビルド規則の唯一の実装)。
    Attribute 行を落とし、_clean() で XML に書けない制御文字を除去する。
    ※ 整形ロジックはここ1箇所に閉じる。verify_build(FA-R23-1c)はこの関数の
      戻り値と成果物のC列を突き合わせるので、二重実装があると検査が無意味になる。
    2026-08-10(R23H-MI-2): 改行をLFへ正規化する。core.autocrlf=trueで
    checkoutするとsrc/*.basがCRLFになり得るが、この関数はペイロード生成
    (_make_vba_src)と検査(_verify_vba_src_bodies/1c)の両方から呼ばれる唯一の
    実装なので、ここで正規化すれば両方に自動で効く(改行方式の違いだけで
    1c全135本が不一致になりビルド不能になるのを防ぐ)。"""
    path = os.path.join(root, m["path"])
    with open(path, encoding="utf-8-sig") as fp:
        txt = fp.read()
    txt = txt.replace("\r\n", "\n").replace("\r", "\n")
    out_lines = []
    # R35 波1: .cls のクラスヘッダ(VERSION 1.0 CLASS / BEGIN ... END)は VBE の
    # エクスポート形式であってソースではない(riskconsulting の _vba_src_text と
    # 同じ規則:**先頭の連続ヘッダだけ**を対象にし、本文の "End Sub"/"END" 等を
    # 巻き込まない)。09-03 敵対的レビュー班A m-4: 現状 build_baked_vba_project は
    # _vba_src_modules(type=="class"を除外)しか回さないため、この分岐自体は
    # 到達しない。.cls を焼く経路ができたときのために残す。
    # BEGIN があって END が無い入力は、下の body_started が一度も真にならず
    # 全行が飲み込まれ「先頭行が空白のみ」という原因の読めないBuildErrorに
    # なるため、専用のBuildErrorをここで先に出す。
    in_cls_header = False
    body_started = False
    for line in txt.split("\n"):
        stripped = line.lstrip("﻿")
        if not body_started:
            bare = stripped.strip()
            if in_cls_header:
                if bare == "END":
                    in_cls_header = False
                continue
            if bare.startswith("VERSION ") and bare.endswith("CLASS"):
                continue
            if bare == "BEGIN":
                in_cls_header = True
                continue
            if bare != "" and not bare.startswith("Attribute "):
                body_started = True
        if stripped.lstrip().startswith("Attribute "):
            continue
        out_lines.append(stripped)
    if in_cls_header:
        raise BuildError(
            f"{m['name']}: .clsヘッダの END が見つかりません"
            "(BEGIN はあるが END が無い入力。ヘッダ除去がファイル全体を"
            "飲み込んでしまうため中断します)")
    cleaned = _clean("\n".join(out_lines))

    # 2026-08-10(R23bH-F4): 整形後ソースの先頭行が空白のみ(空行/タブのみ)
    # ならビルドを止める。理由: modInstallCheck.ExpectedLineCount(および
    # ここの _expected_line_count)が落とすのは【末尾】の空行だけで、
    # 【先頭】の空行はそのまま1行として数える(=非対称)。CodeModule側の
    # AddFromStringも先頭の空行を1行として素直に受け入れるので、行数自体は
    # 現状ズレない。しかし将来、先頭が空行のモジュールが増えたときに
    # 「先頭空行トリムだけが期待/実測のどちらかにだけ紛れ込む」実装変更が
    # 入ると、D列突合(F1)が恒久的な偽陽性になりかねない。現在該当0本なので、
    # ここで先に塞いで将来の踏み抜きを防ぐ。
    first_line = cleaned.split("\n", 1)[0]
    if first_line.strip() == "":
        raise BuildError(
            f"{m['name']}.bas: 整形後ソースの先頭行が空白のみです"
            "(先頭空行は期待/実測の行数計算が非対称になりうるため禁止)")
    return cleaned


def _expected_line_count(src):
    """ソース文字列の「末尾空行を落とした行数」を返す(純関数)。
    src/core/modInstallCheck.bas の Public Function ExpectedLineCount と
    完全同一の規則(CRLF/CR→LF正規化→split→末尾の空白・タブのみ行を落とす)
    をPython側に持つ、唯一の実装(2026-08-10 R23bH-F1)。
    _make_vba_src(D列の焼き込み)と verify_build/_verify_vba_src_bodies
    (D列突合検査)の両方がここを呼ぶ。空文字列は0行。"""
    s = src.replace("\r\n", "\n").replace("\r", "\n")
    if s == "":
        return 0
    parts = s.split("\n")
    n = len(parts)
    while n > 0 and parts[n - 1].replace("\t", " ").strip() == "":
        n -= 1
    return n


def _make_vba_src(wb, present_modules, root):
    """vba_src シート: 標準モジュール(*.bas)のソースを1行1モジュールで格納する。
    自己インストーラ(ThisWorkbookストリーム)がこのシートを読んで
    VBComponents.Add(1)でモジュールを注入する。
    注意: クラスモジュール(type=class, 例 ThisWorkbook.cls)は
    VBComponents.Add(1) では追加できない(標準モジュール専用API)ため対象外。
    2026-08-10(R23bH-F1): D列(row2以降)へ各モジュールの「期待行数」
    (_expected_line_count、modInstallCheck.ExpectedLineCountと同一規則)を
    焼き込む。modInstallCheck.VI はこれを読むだけで検算でき、実行時に
    C列(全135本・約260万字)を再読込してReplace/Splitする必要がなくなる
    (32bit Excelのメモリ逼迫時、検証器自身が新たな失敗点になるのを防ぐ)。
    E1(row1,col5)はOnTime予約時刻用で不可侵のため、D列もrow2以降にしか
    書かない(row1のD1は空のまま)。"""
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
    for m in _vba_src_modules(present_modules):
        cleaned = _vba_src_text(root, m)

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
        ws.cell(row=row, column=4, value=_expected_line_count(cleaned))
        injected.append(m["name"])
        row += 1

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 80
    ws.column_dimensions["D"].width = 10
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
#   ・ソース中にあった2行のコメント("f>0: half-injected. Do not Save." と
#     "Detach Boot(1004); E1=time. R12-3-8.")は 2026-08-06(R20-1f)に
#     ここへ退避した。意味は変わっていない:
#       - f>0 は「一部のモジュールしか注入できなかった」状態。半端な状態を
#         保存すると次回以降ずっと壊れたブックになるので Save しない。
#       - Boot の切り離しは、VBE注入直後に同期で Boot を呼ぶと 1004 になる
#         ことがあるための逃げ。予約時刻は vba_src!E1 に置き、
#         modBoot.CancelPendingInstallerBoot / Auto_Close が取り消す(R12-3-8)。
#   ・Workbook_WindowResize(2026-08-06 R20-1f): 窓の大きさが変わったら
#     modViewport.OnWindowResized へ転送する。このアプリの列幅・カード幅・
#     塗り範囲は全て「今の可視幅」から決まるのに、再計算の機会が「画面を
#     開いたとき」だけで、開いた後に窓を広げると右と下に空白が残っていた
#     (実機第7報⑦)。デバウンス・再入抑止・OnTimeの後始末は modViewport 側。
#     Application.Run 経由なのは、この時点で modViewport が未注入でも
#     コンパイルを通す必要があるため(既存の modBoot.Boot 呼びと同じ作法)。
#     ※ Workbook_BeforeClose はここに【絶対に足さない】。インストーラは
#       Workbook_Open で自分自身を書き換えて Save する設計で、閉じる側に
#       手を入れると保存済みブックの整合が崩れる(§14手順6)。
#   ・失敗検出の強化(2026-08-10 R23-1a): 以前は f が増えるのは
#     VBComponents.Add が Nothing を返したときだけで、c.Name / DeleteLines /
#     AddFromString の失敗は On Error Resume Next 下で握り潰されていた。
#     「名前だけ付いた空モジュール」でも f=0 のまま Save まで到達し、
#     壊れた状態がファイルに焼き付く(実機第9報①)。今は
#       (a) 上記3操作のいずれかで Err.Number<>0 なら f を加算、
#       (b) 本文があるはず(LenB(s)>0)なのに注入後 CountOfLines<1 なら f を加算
#           (Err が立たない無言破損への防御。実機第9報①がこれ)
#     とし、f>0 のときは Save しない既存ガードで確実に止める。
#     ペイロード(vba_src)はファイル上で無傷なので、保存せずに閉じて開き直せば
#     全量が再試行される。MsgBox はその手順をそのまま伝える文言にした。
#     ※ VBAの And は短絡評価しないため (b) は入れ子の If で書いている。
#   ・BL-1(2026-08-10 R23H): f>0でExit Subした直後にThisWorkbook.Savedを
#     Trueへ立てる。MsgBoxで「保存しないで閉じて開き直せ」と伝えても、実際に
#     ExcelがWorkbook_Open完了後Workbook_Close時に「変更を保存しますか?」と
#     聞いてしまうと、既定ボタン[保存]の反射押しで半端な状態がそのまま
#     ファイルへ焼き付く(MsgBoxの案内と矛盾する)。Savedを立てて未保存扱いを
#     解除することでこのプロンプト自体を出させない。
#   ・MI-1(2026-08-10 R23H): (b)のCountOfLines読取自体が例外を出すケースへの
#     防御を追加。On Error Resume Next下でCountOfLines<1判定式そのものが
#     例外を出すと、Ifステートメントが中断されてf加算されず無言通過する
#     穴があった。判定の直後にErr.Number<>0を見てfを加算する。判定が
#     例外なく完了した通常経路ではこの時点でErr.Numberは0のままなので
#     二重加算はしない(例外時は判定文自体が中断してf未加算のまま次行へ
#     進むため、続くErr.Numberチェックが単独でfを1回だけ加算する)。
#   ・MA-3(2026-08-10 R23b): 【部分注入】の検出。R23の(b)は「1行も入って
#     いない」ことしか見ておらず、旧文字列注入APIが途中で切れて1行でも
#     入っていればCountOfLines>=1で素通りする。実機で繰り返し出ている
#     「modViewport2.BadgeRowsFor が見つかりません」は、まさに
#     「モジュールは在るが中身が足りない」形のコンパイルエラーである。
#     そこで注入ループ完了後に modInstallCheck.VI を呼んで全モジュールの
#     行数を vba_src の期待値と突合させていた(実装は
#     src/core/modInstallCheck.bas)。
#   ・R35で VI 呼び出しを撤去(spec_20260903_R35_配布方式転換.md §2-1追記):
#     方式Bでは modInstallCheck.VI 自体が撤去され(波1・9e165e3)、
#     installerモードでもこの2行(f = f + Application.Run("modInstallCheck.VI")
#     とその直後のErr.Number検査)は実機で常に失敗し「Setup NG」に落ちる
#     退行だった。呼び出し2行を削り、f>0のSave抑止とSaved=Trueの既存ガードは
#     Add失敗検出(上記(a)(b))だけで成立させる。
#   ・MsgBox文言の短縮(2026-08-10 R23b): ThisWorkbookストリームの
#     圧縮後サイズ上限に収めるため、英字の2文言を短くした
#     ("Setup incomplete (...)... then reopen to retry."→"Setup NG(...)...
#     reopen."、"VBA Project trust required. See howto sheet."→
#     "Trust VBA project. See howto sheet.")。R35でVI呼び出しが消えたため
#     この短縮は現在は不要な余裕だが、文言は変えていない。
#     実測(R35波2): 圧縮後1,115B / ストリーム上限1,148B(残り33B)。
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
Private Sub Workbook_WindowResize(ByVal Wn As Window)
  On Error Resume Next
  Application.Run "modViewport.OnWindowResized"
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
        Err.Clear
        c.Name = n
        If c.CodeModule.CountOfLines > 0 Then c.CodeModule.DeleteLines 1, c.CodeModule.CountOfLines
        If LenB(s) > 0 Then c.CodeModule.AddFromString s
        If Err.Number <> 0 Then
          f = f + 1
        ElseIf LenB(s) > 0 Then
          If c.CodeModule.CountOfLines < 1 Then f = f + 1
          If Err.Number <> 0 Then f = f + 1
        End If
      End If
      Err.Clear
      On Error GoTo Done
    End If
  Next r
  On Error Resume Next
  Application.Run "modBoot.RunFirstRunPromptEarly"
  Err.Clear
  If f > 0 Then
    MsgBox "Setup NG(" & f & "). Close WITHOUT saving, reopen.", vbCritical
    ThisWorkbook.Saved = True
    Exit Sub
  End If
  ThisWorkbook.Save
  If Err.Number<>0 Then Application.Run "modLog.LogUsage","save_fail","",Err.Description
  Err.Clear
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
  MsgBox "Trust VBA project. See howto sheet.", vbCritical
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
# R23c-F1: _VBA_PROJECTストリームの無害化(幽霊コンパイルエラー根治)
#
# 症状: 実機(Excel 365 2502/32bit)で、注入されたモジュールが477行完全・
# テキスト完全一致にもかかわらず modViewport2.BadgeRowsFor だけがコンパイル
# エラーになる現象が、クリーンインストール3回で100%再現した。
#
# 原因: 配布xlsmの vbaProject.bin は build/template_skeleton.xlsm 由来で、
# その _VBA_PROJECT ストリーム(3,061B)が
#   [0:2] 0x61CC (Reserved1)
#   [2:4] Version = 0x00DF   ← テンプレートを作ったOfficeビルドのスタンプ
#   [4]   0x00 / [5:7] Reserved3
#   [7:]  PerformanceCache 3,054B  ← 当時のコンパイル済み状態
# を保持したままだった。MS-OVBA仕様は「Versionが開き手のOfficeと一致する
# 場合、PerformanceCache がソースの代わりに使われる」「相互運用のため書き手は
# Version=0xFFFF とし PerformanceCache を含めてはならない(MUST)」と定めて
# いる。ユーザーのExcelビルドのスタンプが 0x00DF と一致した結果、
# 我々がVBIDE経由で注入した新しいソースではなく古いキャッシュ側の名前解決が
# 信用され、実在するはずのプロシージャが見つからなくなったと推定される。
#
# 対策(二重防御): Version を 0xFFFF へ書き換え、かつ PerformanceCache 全体を
# ゼロ埋めする。ストリーム長は 3,061B のまま一切変えない(olefile.write_stream
# は同サイズ書き込みしかできず、サイズを変えるとCFBのFATを組み直す必要がある。
# ThisWorkbook差し替えと同じ制約)。Version=0xFFFF により Module1/Sheet1 等の
# 各モジュールストリームのキャッシュも一括で無視されるため、それらは触らない。
# ---------------------------------------------------------------------------
_VBA_PROJECT_STREAM_SIZE = 3061      # template_skeleton.xlsm 由来の固定長
_VBA_PROJECT_RESERVED1 = 0x61CC      # MS-OVBA 2.3.4.1 Reserved1(固定値)
_VBA_PROJECT_VERSION_IGNORE = 0xFFFF  # 「キャッシュを使うな」の相互運用値
_VBA_PROJECT_CACHE_OFFSET = 7        # PerformanceCache の開始オフセット


def _neutralize_vba_project(stream: bytes) -> bytes:
    """_VBA_PROJECTストリームを同サイズのままキャッシュ無効化する。"""
    if len(stream) != _VBA_PROJECT_STREAM_SIZE:
        raise BuildError(
            f"_VBA_PROJECTストリームのサイズが想定外です: "
            f"期待={_VBA_PROJECT_STREAM_SIZE}バイト 実際={len(stream)}バイト "
            "(template_skeleton.xlsm が差し替わった可能性。"
            "サイズを変えずに書き換える前提が崩れるため中断します)")
    reserved1 = struct.unpack("<H", stream[0:2])[0]
    if reserved1 != _VBA_PROJECT_RESERVED1:
        raise BuildError(
            f"_VBA_PROJECTストリームの先頭2バイトが 0x{_VBA_PROJECT_RESERVED1:04X} "
            f"ではありません(実際=0x{reserved1:04X})。"
            "MS-OVBA の _VBA_PROJECT ヘッダとして解釈できないため中断します。")

    out = bytearray(stream)
    struct.pack_into("<H", out, 2, _VBA_PROJECT_VERSION_IGNORE)
    # [4] と [5:7] (Reserved2/Reserved3) は現状維持。
    for i in range(_VBA_PROJECT_CACHE_OFFSET, len(out)):
        out[i] = 0
    return bytes(out)


# ===========================================================================
# R35 波1: 配布方式B(裁定書27 W9-A・spec_20260903_R35_配布方式転換.md)
# 完成品 vbaProject.bin をビルド時に書く(build/ovba_write.py 移植を使う)。
# --vba-mode baked(既定)側の実装。installer(1リリース限りの開発用
# フォールバック)は下の patch_installer 以下に従来どおり残る。
# ===========================================================================

# spec §2-2(09-03 レビュー班B MAJOR-2 裁定変更): baked モードの ThisWorkbook
# (document module)本文そのもの。Workbook_Open は持たない。起動は旧方式と
# 同じく Excel が自動実行する modBoot.Auto_Open(標準モジュール)が担う。
# ThisWorkbook に Workbook_Open を置くと Auto_Open と合わせて Boot が2回走る
# (2周目は gBootDone 分岐で3画面を再描画し共有フォルダへ実I/Oを打つ。R8 F3
# の「起動中に共有I/Oを走らせない」設計を破る)。旧方式(installer)でも実際の
# 起動経路は Auto_Open で、インストーラの OnTime 予約は Auto_Open→Boot の
# CancelPendingInstallerBoot が取り消していた=新しい前提ではない。
# Workbook_BeforeClose も入れない(modBoot.Auto_Close との二重実行を避ける)。
_BAKED_THISWORKBOOK_BODY = (
    "Option Explicit\n"
    "Private Sub Workbook_WindowResize(ByVal Wn As Window)\n"
    "    On Error Resume Next\n"
    "    modViewport.OnWindowResized\n"
    "End Sub\n"
)


def build_baked_thisworkbook() -> str:
    """spec §2-2 の固定文字列を返す(ASCII限定を自己検証する)。
    document module の属性行(Attribute VB_Base 等)はここでは付けない
    (ovba_write.module_stream_source(..., "document") が付ける)。
    Workbook_Open を含まないことも明示検査する(班B MAJOR-2。将来誰かが
    足したらここで赤にする=起動時Boot二重実行の再発を機械的に止める)。"""
    try:
        _BAKED_THISWORKBOOK_BODY.encode("ascii")
    except UnicodeEncodeError as e:
        raise BuildError(f"build_baked_thisworkbook はASCII限定です(spec §2-2): {e}")
    if "Workbook_Open" in _BAKED_THISWORKBOOK_BODY:
        raise BuildError(
            "build_baked_thisworkbook に Workbook_Open が含まれています"
            "(spec §2-2: Auto_Open との二重Boot実行を避けるため禁止)")
    return _BAKED_THISWORKBOOK_BODY


def _cp932_encode_replace(text: str):
    """text を CP932 へエンコードする。1文字ずつ試み、表現できない文字は
    '?' へ置換して数える(§2-5: 置換した文字数を自己検証で報告するため、
    Python標準の errors="replace" に頼らず自前でカウントする)。
    戻り値: (encoded: bytes, replaced_count: int)"""
    out = bytearray()
    replaced = 0
    for ch in text:
        try:
            out += ch.encode("cp932")
        except UnicodeEncodeError:
            out += b"?"
            replaced += 1
    return bytes(out), replaced


def _baked_std_module_bytes(name: str, body: str):
    """標準モジュール1本のモジュールストリーム本文(Attribute行+CRLF本文)を
    CP932(置換あり)で作る。ovba_write.module_stream_source と同じ正規化を
    行うが、置換ができないと本体側は例外を投げる(strict encode)ため、
    ここでは呼ばずに同じ変換を自前で行う(§2-5の裁定はコメントの?置換を
    許すが、それは呼び出し側の裁量であって ovba_write.py 本体には無い)。"""
    text = body.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    if text and not text.endswith("\r\n"):
        text += "\r\n"
    full = (ovba_write._ATTR_STD % name) + text
    return _cp932_encode_replace(full)


def _template_document_sources(template_bin: bytes) -> dict:
    """template_skeleton.xlsm の vbaProject.bin から document module のソースを
    MODULEOFFSET を尊重して取り出す(build_baked_vba_project 専用)。

    **`ovba_write.read_modules` は MODULEOFFSET≠0 のストリームを正しく
    解凍しない(2026-09-03 実Excelで発覚したBLOCKER)**。read_modules は
        src = ovba_write.ovba.ovba_decompress(raw)[cur["offset"]:]
    と書いているが、MODULEOFFSET は「生ストリーム内で圧縮ソースが始まる
    位置」であって解凍後の位置ではない。正しくは
        src = ovba.ovba_decompress(raw[cur["offset"]:])
    である。Excel が書いたテンプレートの Sheet1(document module)は
    MODULEOFFSET=969(先頭969バイトがPerformanceCache)なので、offset 0 から
    解凍すると p-code キャッシュをゴミとして解凍し、それを Sheet1 の
    「ソース」として焼いてしまう(実機で 0xFF の羅列になっていた。
    実Excelで開くとThisWorkbookのイベント宣言行でコンパイルエラーになった)。
    ovba_write.py 本体は移植元へバグ報告済みでここでは直さない(触ってはいけない
    凍結対象ではないが、修正版が来たら差し替える前提)。この関数だけが
    正しい取り出し方をする。baked方式が焼く document module はビルド後は
    全て MODULEOFFSET=0 になる(build_vba_project が焼く側)ので、この
    バグは「Excel製テンプレートから読む」経路(=ここ)にしか無い。

    戻り値: {モジュール名: {"source": bytes, "type": "document"|"procedural"}}
    (PROJECT ストリームの Document= 行にある名前だけを返す。read_modules と
    同じ判定: 0x0022 だが Document= に無いものは class module なので除く)。
    """
    cfb = ovba.CFBReader(template_bin)
    dir_dec = ovba.ovba_decompress(cfb.read("dir"))
    out = {}
    cur = None
    for _off, rid, _size, body in ovba_write.iter_dir_records(dir_dec):
        if rid == ovba_write.REC_MODULENAME:
            cur = {"name": body.decode("cp932"), "offset": 0, "type": None,
                   "stream": None}
        elif cur is None:
            continue
        elif rid == ovba_write.REC_MODULESTREAMNAME:
            cur["stream"] = body.decode("cp932")
        elif rid == ovba_write.REC_MODULEOFFSET:
            cur["offset"] = struct.unpack('<I', body)[0]
        elif rid in (ovba_write.REC_MODULETYPE_PROCEDURAL,
                     ovba_write.REC_MODULETYPE_DOCUMENT):
            cur["type"] = ("document" if rid == ovba_write.REC_MODULETYPE_DOCUMENT
                           else "procedural")
        elif rid == ovba_write.REC_MODULE_TERMINATOR:
            raw = cfb.read(cur.get("stream") or cur["name"])
            # 09-03 敵対的レビュー班C 記録のみ(2周目): MODULEOFFSET がストリーム長を
            # 超えていたり圧縮データ自体が壊れていると ovba_decompress が
            # 素の ValueError を投げ、原因不明のトレースバックでビルドが落ちる。
            # ここで捕まえて「どのモジュールで何が起きたか」を BuildError として
            # 明示する(壊れたテンプレートを早期に fail-closed で弾く)。
            try:
                src = ovba.ovba_decompress(raw[cur["offset"]:])   # ← ここが正しい順序
            except ValueError as e:
                raise BuildError(
                    f"{cur['name']}: MODULEOFFSET がストリーム長を超えているか"
                    f"圧縮データが壊れています: {e}")
            out[cur["name"]] = {"source": src, "type": cur["type"]}
            cur = None
    try:
        proj = cfb.read("PROJECT").decode("cp932", errors="replace")
    except KeyError:
        proj = ""
    doc_names = set()
    for line in proj.splitlines():
        if line.startswith("Document="):
            doc_names.add(line[len("Document="):].split("/")[0].strip())
    return {nm: info for nm, info in out.items()
            if info["type"] == "document" and nm in doc_names}


# document module の VB_Base に埋め込まれるホストの型ライブラリGUID。
# baked方式が焼く document module は ThisWorkbook(Excel.Workbook)と
# Sheet1等のワークシート(Excel.Worksheet)の2種類だけ(spec §2-11)。
def _document_module_expected_base_guid(name: str) -> str:
    """document moduleの名前からVB_Baseに入るべきGUIDを返す。"""
    return (ovba_write._VB_BASE_WORKBOOK if name == "ThisWorkbook"
            else ovba_write._VB_BASE_WORKSHEET)


def document_module_sanity_errors(name: str, source: bytes,
                                   expected_base_guid: str) -> list:
    """document module のソースが本当にテキストか(p-codeのゴミではないか)を
    確かめる fail-closed 検査。build_baked_vba_project・_verify_baked_build・
    tools/bin_roundtrip.py の3箇所で共有する(二重実装禁止)。
    spec_20260903_R35_配布方式転換.md §6 リスク台帳 #6 の再発防止策。

    2026-09-03 実Excelで発覚したBLOCKER(_template_document_sources の
    docstring参照)の再現形そのもの(MODULEOFFSETを無視して解凍した
    p-codeキャッシュを「ソース」として焼いた)を機械的に検出する:
      ① `Attribute VB_Name = "<name>"` で始まる
      ② `Attribute VB_Base = "0{<expected_base_guid>-` を含む
      ③ 0x00〜0x08 や 0xFF のバイトを含まない(=テキストであることの
         直接証拠。実際に壊れたときの中身がまさにこれだった=0xFFの羅列)

    注記(司令塔への報告事項): 指示書の(a)(b)は文字通り読むと
    `Attribute VB_Base = "0{00020820-` (Excel.Worksheet)を全document
    moduleに一律要求しているが、ThisWorkbookは実際には
    Excel.Workbook(00020819)であり、一律適用すると常にThisWorkbookが
    赤くなる。ここではモジュールごとの実際のGUID(expected_base_guid、
    _document_module_expected_base_guidが決める)を検査するよう解釈した。
    """
    errors = []
    want_name = ('Attribute VB_Name = "%s"' % name).encode("cp932")
    if not source.startswith(want_name):
        errors.append(
            f"document module '{name}': ソース先頭が {want_name!r} では"
            f"ありません(実際の先頭60バイト: {source[:60]!r})")
    want_base = ('Attribute VB_Base = "0{%s-' % expected_base_guid).encode("cp932")
    if want_base not in source:
        errors.append(
            f"document module '{name}': {want_base!r} が見当たりません")
    bad_bytes = sorted({b for b in source if b <= 0x08 or b == 0xFF})
    if bad_bytes:
        errors.append(
            f"document module '{name}': テキストとして不正なバイトを含みます"
            f"(0x00〜0x08または0xFF): {[hex(b) for b in bad_bytes]}"
            "(p-codeキャッシュ混入の疑い)")
    return errors


# 09-03 敵対的レビュー班A m-1: vbaProjectストレージ直下の固定ストリーム名
# (大小文字を区別しない。空文字列はストレージ自身を指すため同様に禁止)。
_RESERVED_STREAM_NAMES = frozenset({"dir", "_vba_project", "project", "projectwm", ""})


def build_baked_vba_project(template_bin: bytes, shipped_modules, root: str):
    """spec §2-11・§4波1-4: 完成品 vbaProject.bin を組み立てる。

    - ThisWorkbook: build_baked_thisworkbook() の固定文字列(document module。
      VB_Base はブックのGUID=既定)。
    - Sheet1: template_bin の document module のうち ThisWorkbook 以外
      (=Sheet1。09-03追記の裁定: MyBookshelfはriskconsultingと違いSheet1を
      残す)を read_modules() で取り出し、ソースをそのまま(Attribute行込み・
      Excelが書いたVB_Base込み)写す。
    - 標準モジュール(_vba_src_modules と同じ判定=155本。class型のThisWorkbook
      とvba_src=Falseは除く): _vba_src_text → CRLF化 → Attribute VB_Name行 →
      CP932(§2-5: コード行はlint ERRORで塞がれているため置換はコメントにしか
      起きない。置換文字数を合算して返す)。
    - _VBA_PROJECT は _neutralize_vba_project 済みの版を渡す(installerパッチと
      同じPerformanceCache無害化。R23c-F1)。

    戻り値: (vba_bin: bytes, replaced_chars: int)
    """
    tw_body = build_baked_thisworkbook()
    tw_src = ovba_write.module_stream_source("ThisWorkbook", tw_body, "document")
    vmods = [ovba_write.VbaModule("ThisWorkbook", tw_src, "document")]

    tmpl_docs = _template_document_sources(template_bin)
    other_docs = [(nm, info) for nm, info in tmpl_docs.items()
                  if nm != "ThisWorkbook"]
    if len(other_docs) != 1 or other_docs[0][0] != "Sheet1":
        raise BuildError(
            "template_skeleton.xlsm の document module 構成が想定外です"
            "(baked モードは ThisWorkbook + Sheet1 の2本を前提にしている。"
            f"実際に見つかった document module: "
            f"{['ThisWorkbook'] + [n for n, _ in other_docs]})")
    sheet1_name, sheet1_info = other_docs[0]
    # fail-closed 検査(a): spec §6 リスク台帳 #6(2026-09-03実Excelで発覚した
    # BLOCKERの再発防止)。写す前にソースが本物のテキストであることを確かめる。
    sanity = document_module_sanity_errors(
        sheet1_name, sheet1_info["source"],
        _document_module_expected_base_guid(sheet1_name))
    if sanity:
        raise BuildError(
            f"template_skeleton.xlsm から写した document module "
            f"'{sheet1_name}' のソースが壊れています"
            f"(spec §6 リスク台帳 #6): " + " / ".join(sanity))
    vmods.append(ovba_write.VbaModule(sheet1_name, sheet1_info["source"], "document"))

    replaced_total = 0
    for m in _vba_src_modules(shipped_modules):
        body = _vba_src_text(root, m)
        encoded, replaced = _baked_std_module_bytes(m["name"], body)
        replaced_total += replaced
        vmods.append(ovba_write.VbaModule(m["name"], encoded, "std"))

    # 09-03 敵対的レビュー班A m-1 / 班C MINOR(2周目): モジュール名がCFBの
    # 固定ストリーム名(dir/_VBA_PROJECT/PROJECT/PROJECTwm)と衝突すると、
    # ovba_write.py本体を通さずとも書き出したbinが壊れる(該当ストリームを
    # 上書きしてしまう)。実際にCFBへ書かれるのは vm.stream_name(未指定なら
    # name にフォールバック)なので、衝突検査は name と stream_name の両方を
    # 見る。ovba_write.py本体には手を入れず、こちら側で先に fail-closed で止める。
    for vm in vmods:
        if (vm.name.lower() in _RESERVED_STREAM_NAMES
                or vm.stream_name.lower() in _RESERVED_STREAM_NAMES):
            raise BuildError(
                f"モジュール名'{vm.name}'(stream_name='{vm.stream_name}')が"
                f"CFBの固定ストリーム名と衝突します"
                f"(予約名: {sorted(_RESERVED_STREAM_NAMES)})。"
                "vbaProject.binの該当ストリームを上書きしてしまうため中断します。")

    vba_project_stream = _neutralize_vba_project(
        ovba.CFBReader(template_bin).read("_VBA_PROJECT"))
    vba_bin = ovba_write.build_vba_project(
        template_bin, vmods, vba_project_stream=vba_project_stream)
    return vba_bin, replaced_total


# spec §2-7・§2-8: 配布binに現れてはいけない文字列(AVが重く見る「隠しシートの
# コードを自分に書き込むドロッパー」の形そのもの)。FORBIDDEN_BIN_STRINGS は
# FAIL(0件でなければ出荷を止める)。REPORT_BIN_STRINGS は件数報告のみ
# (WScript.Shell/new:{ の撤去は機能設計を伴うためR36送り。spec §2-8末尾)。
# **この表の唯一の実装**。tools/bin_roundtrip.py がこれをimportして使う
# (二重実装禁止)。
FORBIDDEN_BIN_STRINGS = ("VBProject", "AddFromString", "ExecuteExcel4Macro",
                          "VBComponents", "CodeModule")
REPORT_BIN_STRINGS = ("WScript.Shell", "new:{")


def _decompressed_bin_text(vba_bin):
    """完成品binの全モジュールソースを解凍して1本のテキストにする。
    禁止文字列の検査は**解凍した本文**に対して行う(圧縮バイト列をgrepしても
    中身は見えない)。コメントも対象(静的スキャナは文字列を区別しない)。"""
    parts = []
    for nm, info in ovba_write.read_modules(vba_bin).items():
        parts.append(nm)
        parts.append(info["source"].decode("cp932", errors="replace"))
    return "\n".join(parts)


def forbidden_strings_in_bin(vba_bin):
    """完成品binに現れたFAIL側の禁止文字列を返す(spec §2-7・§2-8)。
    tools/bin_roundtrip.py もこれをimportして使う(二重実装禁止)。"""
    low = _decompressed_bin_text(vba_bin).lower()
    return [w for w in FORBIDDEN_BIN_STRINGS if w.lower() in low]


def report_strings_in_bin(vba_bin):
    """件数報告のみ(FAILにしない)の文字列の出現回数を返す(spec §2-8)。
    {文字列: 出現回数} を返す。0件でもキー自体は残す(表示側で0件と分かるように)。"""
    low = _decompressed_bin_text(vba_bin).lower()
    return {w: low.count(w.lower()) for w in REPORT_BIN_STRINGS}


def _verify_baked_build(out_path, present_modules, root):
    """spec §2-7: baked モードの読み戻し検査。
    ①モジュール集合=台帳 ②各本文が_vba_src_textの結果とバイト一致
    ③ThisWorkbookが§2-2と一致 ④全MODULEOFFSET=0
    ⑤_VBA_PROJECT Version=0xFFFF・PerformanceCacheゼロ埋め
    + 成果物のいずれかのワークシートXMLが codeName="Sheet1" を持つこと
    (Sheet1 document module が孤児でない証跡)。"""
    errors = []

    try:
        with zipfile.ZipFile(out_path) as z:
            vba_bin = z.read("xl/vbaProject.bin")
            sheet1_codename_found = any(
                b'codeName="Sheet1"' in z.read(n)
                for n in z.namelist()
                if n.startswith("xl/worksheets/") and n.endswith(".xml")
            )
    except Exception as e:
        return [f"baked bin: 成果物zipの読み出しに失敗: {e}"]

    if not sheet1_codename_found:
        errors.append(
            "成果物のどのワークシートXMLにも codeName=\"Sheet1\" がありません"
            "(Sheet1 document module が孤児化している可能性。spec §2-11 09-03追記)")

    try:
        got = ovba_write.read_modules(vba_bin)
    except Exception as e:
        return errors + [f"baked bin: ovba_write.read_modulesでの読み戻しに失敗: {e}"]

    # ① モジュール集合=台帳(155本の標準モジュール + ThisWorkbook + Sheet1)
    expected_names = {m["name"] for m in present_modules} | {"Sheet1"}
    got_names = set(got.keys())
    if got_names != expected_names:
        errors.append(
            f"baked: モジュール集合が台帳と不一致: "
            f"台帳のみ={sorted(expected_names - got_names)} "
            f"成果物のみ={sorted(got_names - expected_names)}")

    # ② 各標準モジュールの本文が _vba_src_text の結果とバイト一致
    for m in _vba_src_modules(present_modules):
        nm = m["name"]
        if nm not in got:
            errors.append(f"baked: 標準モジュール'{nm}'が成果物にありません")
            continue
        want, _ = _baked_std_module_bytes(nm, _vba_src_text(root, m))
        if got[nm]["source"] != want:
            pos = _first_diff_pos(want, got[nm]["source"])
            errors.append(
                f"baked: '{nm}'の本文が_vba_src_textの結果とバイト不一致 "
                f"(先頭差分位置={pos}バイト目, 期待{len(want)}B/実際{len(got[nm]['source'])}B)")

    # ③ ThisWorkbook が spec §2-2 と一致
    if "ThisWorkbook" not in got:
        errors.append("baked: ThisWorkbook モジュールが成果物にありません")
    else:
        want_tw = ovba_write.module_stream_source(
            "ThisWorkbook", build_baked_thisworkbook(), "document")
        if got["ThisWorkbook"]["source"] != want_tw:
            errors.append("baked: ThisWorkbookの本文がspec §2-2の固定文字列と不一致")

    # ③-2 fail-closed検査(b): 全document module(ThisWorkbook含む)のソースが
    # 本当にテキストか(spec §6 リスク台帳 #6・2026-09-03実Excelで発覚した
    # BLOCKERの再発防止)。document_module_sanity_errorsを(a)(c)と共有する。
    for nm, info in got.items():
        if info["type"] != "document":
            continue
        sanity = document_module_sanity_errors(
            nm, info["source"], _document_module_expected_base_guid(nm))
        if sanity:
            errors.append(
                f"baked: document module '{nm}' の健全性検査に失敗"
                f"(spec §6 リスク台帳 #6): " + " / ".join(sanity))

    # ④ 全 MODULEOFFSET=0(dirストリームを直接なめる。ovba_write.iter_dir_records
    #    を使い、本体には手を入れずに検査だけこちらで行う)。
    try:
        dir_dec = ovba.ovba_decompress(ovba.CFBReader(vba_bin).read("dir"))
        bad_offsets = []
        cur_name = None
        for _off, rid, _size, body in ovba_write.iter_dir_records(dir_dec):
            if rid == ovba_write.REC_MODULENAME:
                cur_name = body.decode("cp932", errors="replace")
            elif rid == ovba_write.REC_MODULEOFFSET:
                val = struct.unpack("<I", body)[0]
                if val != 0:
                    bad_offsets.append((cur_name, val))
        if bad_offsets:
            errors.append(f"baked: MODULEOFFSETが0でないモジュールがあります: {bad_offsets}")
    except Exception as e:
        errors.append(f"baked: dirストリームのMODULEOFFSET検査中に例外: {e}")

    # ⑤ _VBA_PROJECT Version=0xFFFF・PerformanceCacheゼロ埋め(R23c-F2と同じ検査)
    try:
        vp = ovba.CFBReader(vba_bin).read("_VBA_PROJECT")
        if len(vp) != _VBA_PROJECT_STREAM_SIZE:
            errors.append(
                f"baked: _VBA_PROJECTストリームのサイズが変化しています: "
                f"期待={_VBA_PROJECT_STREAM_SIZE}バイト 実際={len(vp)}バイト")
        else:
            ver = struct.unpack("<H", vp[2:4])[0]
            if ver != _VBA_PROJECT_VERSION_IGNORE:
                errors.append(
                    f"baked: _VBA_PROJECTのVersionが0x{_VBA_PROJECT_VERSION_IGNORE:04X}"
                    f"ではありません(実際=0x{ver:04X})")
            nonzero = sum(1 for b in vp[_VBA_PROJECT_CACHE_OFFSET:] if b != 0)
            if nonzero:
                errors.append(
                    f"baked: _VBA_PROJECTのPerformanceCacheがゼロ埋めされていません"
                    f"(非ゼロ={nonzero}バイト)")
    except Exception as e:
        errors.append(f"baked: _VBA_PROJECT検査中に例外: {e}")

    # ⑥ 禁止文字列(FAIL側。spec §2-7・§2-8・波2 タスク6)。解凍した本文で検査
    # (コメント含む)。件数報告側(WScript.Shell/new:{)はここでは検査しない
    # (Stage 4 の出力で表示のみ。R36で機能設計とあわせて撤去)。
    try:
        hits = forbidden_strings_in_bin(vba_bin)
        if hits:
            errors.append(
                f"baked: vbaProject.binに配布禁止の文字列があります"
                f"(spec §2-7・§2-8): {', '.join(hits)}")
    except Exception as e:
        errors.append(f"baked: 禁止文字列検査中に例外: {e}")

    # ⑦ ThisWorkbook本文が参照する modX.Proc の綴りが実在すること(班A M-3)。
    # ThisWorkbookはdocument moduleでありLOの隔離コンパイル(run_lo_tests)に
    # 載らないため、これが唯一の綴り検査になる。
    errors.extend(_thisworkbook_reference_errors(root))

    return errors


_THISWORKBOOK_REF_RE = re.compile(r"\b(mod[A-Za-z0-9_]*)\.([A-Za-z_][A-Za-z0-9_]*)")


def _thisworkbook_reference_errors(root: str) -> list:
    """spec波2追加(班A M-3): build_baked_thisworkbook()の本文に出てくる
    modX.Proc形の参照(現状はmodViewport.OnWindowResizedのみ)について、
    src/**/modX.bas に 'Public Sub Proc' または 'Public Function Proc' が
    実在することを確認する。無ければリストへ理由を追加して返す(空なら合格)。
    ThisWorkbookはLOの隔離コンパイルに載らないので、これが唯一の綴り検査。"""
    errors = []
    body = build_baked_thisworkbook()
    seen = set()
    for mod_name, proc_name in _THISWORKBOOK_REF_RE.findall(body):
        key = (mod_name, proc_name)
        if key in seen:
            continue
        seen.add(key)
        import glob as _glob
        hits = _glob.glob(os.path.join(root, "src", "**", mod_name + ".bas"), recursive=True)
        if not hits:
            errors.append(
                f"baked: ThisWorkbookが参照する'{mod_name}.{proc_name}'の"
                f"モジュールファイル'{mod_name}.bas'がsrc配下に見つかりません")
            continue
        try:
            with open(hits[0], encoding="utf-8") as fp:
                src_text = fp.read()
        except Exception as e:
            errors.append(f"baked: '{hits[0]}' の読み込みに失敗: {e}")
            continue
        proc_re = re.compile(
            r"(?im)^\s*Public\s+(Sub|Function)\s+" + re.escape(proc_name) + r"\b")
        if not proc_re.search(src_text):
            errors.append(
                f"baked: ThisWorkbookが参照する'{mod_name}.{proc_name}'が"
                f"'{os.path.relpath(hits[0], root)}'にPublic Sub/Functionとして"
                "見つかりません(綴り誤りの可能性)")
    return errors


# ---------------------------------------------------------------------------
# vbaProject.bin 外科パッチ (ThisWorkbookストリーム差し替え + dir MOFFSET=0
#                            + _VBA_PROJECT のPerformanceCache無害化)
# --vba-mode installer(1リリース限りの開発用フォールバック。spec §2-1)専用。
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

    # R23c-F1: PerformanceCache無害化。同サイズ書き込みのみ(上のコメント参照)。
    new_vbaproj = _neutralize_vba_project(skel.read("_VBA_PROJECT"))

    buf = io.BytesIO(vba_bin)
    ole = olefile.OleFileIO(buf, write_mode=True)
    ole.write_stream("VBA/dir", new_dir)
    ole.write_stream("VBA/ThisWorkbook", new_tw)
    ole.write_stream("VBA/_VBA_PROJECT", new_vbaproj)
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
# 各ソースセル<=32,000字 / olefileでThisWorkbookストリーム復元確認 /
# vba_srcのC列本文が src/ の対応ファイルと完全一致(FA-R23-1c)
# ---------------------------------------------------------------------------
def _first_diff_pos(a, b):
    """2つの文字列の先頭からの最初の相違位置(0始まり)を返す。"""
    n = min(len(a), len(b))
    for i in range(n):
        if a[i] != b[i]:
            return i
    return n


def _verify_vba_src_bodies(ws, got_names, present_modules, root):
    """成果物 vba_src のC列本文が src/ の対応ファイル(ビルド規則適用後)と
    完全一致することを全モジュールで検査する(FA-R23-1c)。
    2026-08-10(R23bH-F1): あわせてD列(期待行数)が src由来の
    _expected_line_count(C列本文)と一致することも全数突合する
    (modInstallCheck.VI が実行時に信頼するD列の値そのものを、出荷前に
    ビルド側で検算しておくため)。"""
    errors = []
    expected = {}
    for m in _vba_src_modules(present_modules):
        try:
            expected[m["name"]] = _vba_src_text(root, m)
        except OSError as e:
            errors.append(f"vba_src本文検査: '{m['name']}' のソースを読めません: {e}")

    seen = set()
    for i, nm in enumerate(got_names):
        actual = ws.cell(row=i + 2, column=3).value or ""
        seen.add(nm)
        if nm not in expected:
            errors.append(f"vba_src本文検査: '{nm}' に対応する src/ のモジュールがありません")
            continue
        want = expected[nm]
        if actual != want:
            pos = _first_diff_pos(want, actual)
            errors.append(
                f"vba_src本文が src/ と不一致: '{nm}' "
                f"(先頭差分位置={pos}, 期待{len(want)}字/実際{len(actual)}字, "
                f"期待={want[pos:pos + 40]!r} 実際={actual[pos:pos + 40]!r})")

        want_lines = _expected_line_count(want)
        got_d = ws.cell(row=i + 2, column=4).value
        # openpyxl は整数値のセルでも int/float いずれで返すかが環境依存なため、
        # 数値型であれば int() へ正規化してから比較する(非数値/Noneはそのまま
        # 不一致として拾う)。
        got_d_num = int(got_d) if isinstance(got_d, (int, float)) and not isinstance(got_d, bool) else None
        if got_d_num != want_lines:
            errors.append(
                f"vba_src D列(期待行数)が src由来の計算値と不一致: '{nm}' "
                f"D列={got_d!r} 期待={want_lines}")

    for nm in expected:
        if nm not in seen:
            errors.append(f"vba_src本文検査: '{nm}' の行が成果物にありません")
    return errors


# ---------------------------------------------------------------------------
# R32波2で足した3キーの【値そのもの】を出荷物で検算する(2026-08-14 R32 Fix波 F6)。
#
# もともとこの3件は modTestsPure31 のVBAテストが見ていたが、あれは【恒真】だった:
# LibreOffice の純ロジック実行環境には config シートが存在せず、
# modConfig.GetBool/GetLong はどのキーでも必ず Fallback(=呼び出し側が渡した
# 既定値)を返す。つまり `GetBool("pii_scan_enabled", False) = False` は
# 「False = False」を比べているだけで、config 行を消しても値を書き換えても
# 落ちない。守っているように見えて何も守っていない検査だった。
#
# 本当に守りたいのは「出荷する xlsm の config シートに、コード側の既定値と
# 同じ値の行が実在すること」。それは Python 側でしか確かめられない
# (再オープンして B列を読む)。ここが唯一の検査点になる。
#   pii_scan_enabled : FALSE  … 誤検知で発行が止まる実害を止めるための既定オフ
#                               (関所は modPackExport.ExportPackToFile)
#   gap_keep_days    : 30     … modInsight.TrimInboxRows の GetLong 既定値と一致
#   gap_dup_hours    : 24     … modInsightGate.GapDupBlocked の GetLong 既定値と一致
# 片方だけ変えるとズレる。VBA側の既定値を変えたら必ずここも同時に変えること。
R32_CONFIG_EXPECTED = {
    "pii_scan_enabled": False,
    "gap_keep_days": 30,
    "gap_dup_hours": 24,
}


def _verify_r32_config_defaults(got_values):
    errors = []
    for key, want in R32_CONFIG_EXPECTED.items():
        if key not in got_values:
            errors.append(f"config に '{key}' の行がありません(R32の既定値検算)")
            continue
        got = got_values[key]
        if isinstance(want, bool):
            ok = isinstance(got, bool) and got == want
        else:
            ok = isinstance(got, int) and not isinstance(got, bool) and got == want
        if not ok:
            errors.append(
                f"config!{key} が期待値と不一致: 期待={want!r} 実際={got!r}"
                f"(VBA側のGetBool/GetLong既定値と揃っている必要がある)")
    return errors


# ---------------------------------------------------------------------------
def verify_build(out_path, vba_mode, expected_vba_src_names, installer_src, mock_llm_expected,
                 present_modules=None, root=None):
    """ビルド後自己検証(MASTER_SPEC §10)。vba_mode で installer/baked を分岐する
    (R35 波1・spec §3「verify_build の引数」)。installer 側の検査項目は完全に
    従来どおり(引数の意味も不変)。baked 側は §2-7 の読み戻し5項目 +
    Sheet1 codeName 証跡を _verify_baked_build() が担う。"""
    errors = []

    try:
        wb2 = openpyxl.load_workbook(out_path, keep_vba=True)
    except Exception as e:
        return [f"再オープン失敗: {e}"]

    sheets_for_mode = expected_sheets(vba_mode)
    got_sheets = set(wb2.sheetnames)
    want_sheets = set(sheets_for_mode.keys())
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

    for name, state in sheets_for_mode.items():
        if name in wb2.sheetnames:
            actual = wb2[name].sheet_state
            if actual != state:
                errors.append(f"シート'{name}'の可視性不一致: 期待={state} 実際={actual}")

    if vba_mode == "installer":
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
            # FA-R23-1c: モジュール名の集合だけでなく本文まで突き合わせる。
            # 実機第9報①(注入されたモジュールに手続きが無い)のように、名前は
            # 揃っているのに中身が欠けている配布物を出荷段階で止めるため。
            # 期待値は _vba_src_text() ＝ ビルド時にC列を作ったのと同じ関数から
            # 作り直す(整形ロジックを二重実装しない)。
            if present_modules is not None and root is not None:
                errors.extend(
                    _verify_vba_src_bodies(ws, got_names, present_modules, root))
            else:
                # 2026-08-10(R23H-MI-3): 未指定を黙ってスキップすると検査が
                # 実施されたかのように見えてしまう(exit 0で通過)。本文検査が
                # 実質未実施だったことをerrorsへ明示的に積み、失格扱いにする。
                errors.append("vba_src本文検査: 未実施(present_modules/rootが未指定)")
        else:
            errors.append("vba_src シートが存在しない")
    else:
        # R35 波1(spec §2-3・波4班B「残骸」): baked モードに vba_src が
        # 残っていたら方式Aの痕跡が消えていない証拠なので不合格にする。
        if "vba_src" in wb2.sheetnames:
            errors.append("vba_src シートが baked モードの成果物に残っています"
                           "(方式Aの痕跡。expected_sheets(baked)には含まれないはず)")

    if "config" in wb2.sheetnames:
        ws = wb2["config"]
        r, got_mock, n_keys = 2, None, 0
        got_values = {}
        while ws.cell(row=r, column=1).value:
            key = ws.cell(row=r, column=1).value
            if key == "mock_llm":
                got_mock = ws.cell(row=r, column=2).value
            got_values[key] = ws.cell(row=r, column=2).value
            n_keys += 1
            r += 1
        if bool(got_mock) != mock_llm_expected:
            errors.append(f"config!mock_llm が期待値と不一致: 期待={mock_llm_expected} 実際={got_mock}")
        if n_keys != len(build_config_rows(mock_llm_expected)):
            errors.append(f"config のキー数が期待({len(build_config_rows(mock_llm_expected))})と不一致: {n_keys}")
        errors.extend(_verify_r32_config_defaults(got_values))

    if vba_mode == "installer":
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

            # R23c-F2: _VBA_PROJECT が無害化されていることの読み戻し検査。
            # (a)サイズ不変 (b)Version==0xFFFF (c)PerformanceCache全ゼロ
            vp = cfb.read("_VBA_PROJECT")
            if len(vp) != _VBA_PROJECT_STREAM_SIZE:
                errors.append(
                    f"_VBA_PROJECTストリームのサイズが変化しています: "
                    f"期待={_VBA_PROJECT_STREAM_SIZE}バイト 実際={len(vp)}バイト")
            else:
                ver = struct.unpack("<H", vp[2:4])[0]
                if ver != _VBA_PROJECT_VERSION_IGNORE:
                    errors.append(
                        f"_VBA_PROJECTのVersionが0x{_VBA_PROJECT_VERSION_IGNORE:04X}では"
                        f"ありません(実際=0x{ver:04X})。開き手のOfficeビルドと一致すると"
                        "PerformanceCacheがソースより優先され、幽霊コンパイルエラーの原因になります")
                nonzero = sum(1 for b in vp[_VBA_PROJECT_CACHE_OFFSET:] if b != 0)
                if nonzero:
                    errors.append(
                        f"_VBA_PROJECTのPerformanceCacheがゼロ埋めされていません"
                        f"(非ゼロ={nonzero}バイト)")

            # olefile側でも同一バイナリを開けることを確認(異なる実装での復元確認)。
            ole = olefile.OleFileIO(io.BytesIO(vba_bin))
            if not ole.exists("VBA/ThisWorkbook") or not ole.exists("VBA/dir"):
                errors.append("olefileでVBA/ThisWorkbookまたはVBA/dirストリームが検出できない")
            ole.close()
        except Exception as e:
            errors.append(f"vbaProject.bin検証中に例外: {e}")
    else:
        # R35 波1(spec §2-7): baked モードの読み戻し5項目 + Sheet1 codeName証跡。
        if present_modules is None or root is None:
            errors.append("baked bin検査: 未実施(present_modules/rootが未指定)")
        else:
            errors.extend(_verify_baked_build(out_path, present_modules, root))

    return errors


# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------
def _default_out_fname(is_dev: bool, publisher: bool) -> str:
    """既定の出力ファイル名(正規配布名)を1箇所から返す。installerモードの
    正規名ガード(main内)と既定出力名の生成が別々の定義に分かれてズレる
    (R35波2周目レビュー班C BLOCKER)のを防ぐため、双方がこの関数を使う。"""
    if publisher:
        return "MyBookshelf_発行者用_dev.xlsm" if is_dev else "MyBookshelf_発行者用.xlsm"
    return "MyBookshelf_dev.xlsm" if is_dev else "MyBookshelf.xlsm"


def _canonical_out_paths(root: str) -> set:
    """正規4名(dev/prod × 利用者用/発行者用)の絶対パス集合を返す。
    _default_out_fname と同じ組み合わせを列挙するだけの薄いラッパ。"""
    names = {
        _default_out_fname(is_dev, publisher)
        for is_dev in (False, True)
        for publisher in (False, True)
    }
    return {os.path.abspath(os.path.join(root, "dist", n)) for n in names}


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
    ap.add_argument("--vba-mode", choices=("baked", "installer"), default="baked",
                     dest="vba_mode",
                     help="vbaProject.bin の作り方(R35 spec §2-1)。"
                          "baked(既定・配布方式B)=ビルドが完成品binを書く"
                          "(vba_srcシート・VBComponents.Add・AddFromStringが"
                          "配布物から消える)。installer=1リリース限りの開発用"
                          "フォールバック(従来の自己インストーラ外科パッチ。"
                          "--zipとの併用不可)")
    args = ap.parse_args()

    # R35 spec §2-1: installer(AMSI検知済みの自己インストーラ形)は開発用の
    # フォールバックであり、配布経路(--zip)に乗せない。BuildError相当で止める。
    if args.vba_mode == "installer" and args.zip:
        sys.exit(
            "ERROR(BuildError): --vba-mode installer は --zip と併用できません。\n"
            "  installer モードは自己インストーラ外科パッチ(2026-09-02 に社内AVの"
            "AMSIで検知され採用禁止になった形)を再生成する開発用フォールバックで、\n"
            "  配布経路に乗せないことが裁定(spec_20260903_R35_配布方式転換.md §2-1)"
            "です。配布物が要るときは既定の --vba-mode baked を使ってください。"
        )

    root = os.path.abspath(args.root)
    is_dev = bool(args.dev)
    mock_llm = is_dev

    if args.out:
        out_path = os.path.abspath(args.out)
    else:
        out_path = os.path.join(root, "dist", _default_out_fname(is_dev, args.publisher))

    # 09-03 敵対的レビュー班A M-2 / 班C BLOCKER(2周目): installer は開発用
    # フォールバックであり、正規配布名(dist/MyBookshelf*.xlsm)を書けてしまうと
    # 「dist/ に baked と installer のどちらが出荷物か分からないファイルが
    # 並ぶ」事故になる。--out の「有無」ではなく出力先の「値」で判定する
    # (--out で正規名そのものを指定しても素通りしていた抜け穴を塞ぐ)。
    if args.vba_mode == "installer" and out_path in _canonical_out_paths(root):
        sys.exit(
            "ERROR(BuildError): installer モードの成果物は正規配布名で書けません。"
            "--out で別名を指定してください。"
        )

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

    # out_path は上(installerモードの正規名ガード)で既に確定済み。

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

    sheets_for_mode = expected_sheets(args.vba_mode)
    print(f"Stage 2: シート生成 (MASTER_SPEC §4 全{len(sheets_for_mode)}シート"
          f" ※マクロ無効ガードを含む・--vba-mode {args.vba_mode})...")
    _make_macro_guard(wb)
    _make_howto(wb)
    _make_admin_guide(wb)
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
    # chunk_meta(R17波0): section_path/refs_outは自由記述の参照ラベル文字列
    # なので、full_text等と同じく数式インジェクション対策でtext_colsへ入れる。
    _make_headers_only(wb, "chunk_meta", ["chunk_id", "section_path", "refs_out"],
                        "veryHidden", widths=[32, 40, 60], text_cols=[2, 3])
    # doc_outline(R17 Phase2): 章ごとの要約とキーワード。source/section_key/
    # summary/keywords はどれもLLM出力または資料由来の自由文なので、数式
    # インジェクション対策で text_cols へ入れる(chunk_n だけが数値列)。
    _make_headers_only(wb, "doc_outline",
                        ["source", "section_key", "summary", "keywords", "chunk_n"],
                        "veryHidden", widths=[24, 40, 90, 40, 10],
                        text_cols=[1, 2, 3, 4])
    # synonyms(R17 Phase3): term/canonical は自由記述の用語文字列なので、
    # chunk_meta/doc_outline と同じく数式インジェクション対策で text_cols へ。
    _make_headers_only(wb, "synonyms", ["term", "canonical"], "veryHidden",
                        widths=[30, 30], text_cols=[1, 2])
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
    # value 列(B)は利用者の生の質問文・保留中の元質問が入る(modState.SaveState)。
    # 先頭 "=" の質問が数式として解釈されると、不正なら1004で保存が無言で失敗し、
    # 有効なら計算結果や #NAME? として読み出される(R33 W4-8)。他シートと同じく
    # テキスト書式へ固定する(実行時は modState.SaveState 側でも張る=二重化)。
    _make_headers_only(wb, "ui_state", ["key", "value"], "veryHidden", widths=[24, 40],
                        text_cols=[2])
    # 頁チェックポイント(R15-FixB FB-2)。text 列は取り込んだ本文そのものが
    # 入るので、他の非信頼テキスト列と同じくテキスト書式へ固定する
    # (opt側は番兵1字を前置して書くが、書式の防御も二重に掛けておく)。
    _make_headers_only(wb, "ocr_cache", ["key", "text", "saved_at"], "veryHidden",
                        widths=[40, 100, 20], text_cols=[2])
    # 共有知フライホイールの受信箱(modInsight)。共有フォルダから届いた
    # 「解決済みQ&A」と「答えられなかった質問」をここに溜め、本棚への
    # 取り込みは利用者が押したときだけ行う。
    _make_headers_only(wb, "insight_inbox",
                        ["nonce", "kind", "user_id", "author", "created_at",
                         "question", "answer_or_reason", "source_or_dept", "consumed",
                         "selected"],
                        "hidden", widths=[34, 8, 20, 18, 18, 60, 80, 30, 10, 10],
                        text_cols=[4, 6, 7, 8])   # author/question/answer/source

    if args.vba_mode == "installer":
        try:
            injected = _make_vba_src(wb, present, root)
        except BuildError as e:
            sys.exit(f"ERROR: {e}")
        print(f"  vba_src: {len(injected)}モジュールを格納 ({injected})")
    else:
        # R35 baked(既定): vba_src シートは作らない(方式Bは vbaProject.bin へ
        # 直接焼き込む。§2-3)。Stage 4 が present から同じ155本を焼く。
        injected = []
        print("  vba_src: 生成しません(baked モード。Stage 4 で vbaProject.bin へ直接焼き込みます)")
    print(f"  シート最終構成({len(wb.sheetnames)}件): {wb.sheetnames}")

    if set(wb.sheetnames) != set(sheets_for_mode):
        sys.exit(f"ERROR: シート構成がMASTER_SPEC §4と不一致(--vba-mode {args.vba_mode}): {wb.sheetnames}")

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

    with zipfile.ZipFile(tmp_path) as zin:
        parts = {n: zin.read(n) for n in zin.namelist()}
    skel_bin = parts["xl/vbaProject.bin"]
    replaced_chars = None
    if args.vba_mode == "installer":
        print("Stage 4: vbaProject.bin 外科パッチ (自己インストーラ注入)...")
        try:
            installer_src = build_installer_src()
        except BuildError as e:
            sys.exit(f"ERROR: {e}")
        try:
            patched_bin = patch_installer(skel_bin, installer_src)
        except BuildError as e:
            sys.exit(f"ERROR: {e}")
        assert len(skel_bin) == len(patched_bin), "vbaProject.binのバイト長が変化した(バイナリ整合性エラー)"
        parts["xl/vbaProject.bin"] = patched_bin
        print(f"  vbaProject.bin: {len(skel_bin):,} bytes (不変)")
    else:
        # R35 baked(既定): 完成品 vbaProject.bin を書く(spec §2-11・§4波1-4)。
        print("Stage 4: vbaProject.bin 生成 (完成品・baked)...")
        installer_src = None
        try:
            baked_bin, replaced_chars = build_baked_vba_project(skel_bin, present, root)
        except BuildError as e:
            sys.exit(f"ERROR: {e}")
        parts["xl/vbaProject.bin"] = baked_bin
        print(f"  vbaProject.bin: スケルトン{len(skel_bin):,} bytes → "
              f"baked完成品{len(baked_bin):,} bytes (CP932置換{replaced_chars}字)")
        # spec §2-7・§2-8: 件数報告側(WScript.Shell/new:{)は表示のみ(FAILにしない)。
        report_counts = report_strings_in_bin(baked_bin)
        print("  件数報告(FAILにしない。R36で機能設計とあわせて撤去): "
              + ", ".join(f"{w}={n}" for w, n in report_counts.items()))

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
    # R35 波1(spec §2-6): openpyxl の保存結果は [Content_Types].xml が先頭に
    # 来ない。OPCの慣例([Content_Types].xml → _rels/.rels → 残り)に並べ替える
    # (LibreOffieはこの順でないと開けない・本日実測)。並べ替えはモードに
    # 依存しないので installer/baked どちらでも適用する。
    _zip_priority = ("[Content_Types].xml", "_rels/.rels")
    ordered_names = sorted(
        parts.keys(),
        key=lambda n: _zip_priority.index(n) if n in _zip_priority else len(_zip_priority))
    with zipfile.ZipFile(staging_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for n in ordered_names:
            zout.writestr(n, parts[n])
    os.unlink(tmp_path)
    print(f"  一時出力: {staging_path} ({os.path.getsize(staging_path):,} bytes)")

    print("\nStage 6: ビルド後自己検証...")
    errors = verify_build(staging_path, args.vba_mode, injected, installer_src, mock_llm,
                          present_modules=present, root=root)
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
    if args.vba_mode == "installer":
        print("自己検証 OK: 全シート存在 / vba_srcモジュール数一致 / 各ソース<=32000字 / "
              "vba_src本文がsrc/と完全一致 / vba_src D列(期待行数)がsrc由来の計算値と一致 / "
              "ThisWorkbookストリーム復元確認 / dir MOFFSET=0確認 / "
              "_VBA_PROJECT無害化確認(Version=0xFFFF・PerformanceCacheゼロ埋め・3,061B不変)")
    else:
        n_baked_modules = len(present) + 1   # present(155本+ThisWorkbook) + Sheet1
        print("自己検証 OK: 全シート存在(vba_srcシート無し) / "
              f"baked vbaProject.binモジュール集合=台帳一致(計{n_baked_modules}本) / "
              "各標準モジュール本文が_vba_src_textの結果とバイト一致 / "
              "ThisWorkbookがspec §2-2の固定文字列と一致 / 全MODULEOFFSET=0確認 / "
              "_VBA_PROJECT無害化確認(Version=0xFFFF・PerformanceCacheゼロ埋め) / "
              "Sheet1 codeName証跡あり / "
              f"CP932置換 {replaced_chars} 字")

    if args.zip:
        print("\nStage 7: --zip 配布梱包...")
        zip_path = build_dist_zip(out_path, os.path.dirname(out_path), args.publisher,
                                   is_dev, root)
        print(f"  出力: {zip_path} ({os.path.getsize(zip_path):,} bytes)")

    print("\nDone.")


if __name__ == "__main__":
    main()
