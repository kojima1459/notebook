#!/usr/bin/env python3
"""
build_chatbot_v2.py - Build dist/Chatbot_v2.xlsm using in-place VBA patching.

Architecture (after several failed attempts at hand-built vbaProject.bin):

  1. Start from build/template_skeleton.xlsm — a REAL Excel-made .xlsm.
     Its vbaProject.bin is genuine and is accepted by every Excel build.
  2. openpyxl loads it with keep_vba=True (binary preserved byte-for-byte).
  3. We add all data sheets (main, config, system_prompt, department,
     manifest, knowledge_base, feedback, usage_log, vba_src).
  4. We save with openpyxl — vbaProject.bin is still 100% the skeleton's.
  5. We then SURGICALLY patch only TWO streams inside vbaProject.bin:
       - VBA/ThisWorkbook: replace empty class with a self-installer that
         reads vba_src on Workbook_Open and installs all standard modules
         via the VBProject API.
       - VBA/dir: change ThisWorkbook's MOFFSET to 0 so the new source is
         decompressed from byte 0 of the stream (the perf-cache prefix is
         no longer valid).
     CFB header, FAT, directory entries, _VBA_PROJECT, all other streams,
     PROJECT, PROJECTwm — ALL untouched.
  6. We pad both rewritten streams with empty OVBA chunks (3-byte chunks
     that decompress to zero bytes) so the stream sizes match exactly —
     olefile.write_stream requires byte-exact replacement.

Result: ~99.9% of the binary is byte-identical to Excel's own output. Excel
accepts it. On Workbook_Open, the installer reads vba_src and instantiates
all 13 standard modules.

This works around the corporate-Excel rejection of fully hand-built
vbaProject.bin binaries (CFB writer / OVBA compressor / cache mismatch).
"""

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

# Reuse OVBA compress/decompress from make_xlsm
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import make_xlsm

_ILLEGAL = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')
def _clean(s):
    return _ILLEGAL.sub("", s) if isinstance(s, str) else s

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT       = os.path.dirname(SCRIPT_DIR)
SRC_V2     = os.path.join(ROOT, "src", "chatbot_v2")
TEMPLATE   = os.path.join(SCRIPT_DIR, "template_skeleton.xlsm")
ENRICHED   = os.path.join(ROOT, "dist", "index", "chunks_enriched.json")
OUT_XLSM   = os.path.join(ROOT, "dist", "Chatbot_v2.xlsm")

# ---------------------------------------------------------------------------
# Installer source — lives in ThisWorkbook stream of the patched vbaProject.bin
# ---------------------------------------------------------------------------
INSTALLER_SRC = '''Attribute VB_Name = "ThisWorkbook"
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
      If LenB(s) > 0 Then c.CodeModule.AddFromString s
    End If
  Next r
  ThisWorkbook.Save
  On Error Resume Next
  Application.Run "modBoot.Boot"
  Exit Sub
Trust:
  MsgBox "VBA Project trust required", vbCritical
  Exit Sub
Done:
End Sub
'''.replace('\n', '\r\n').encode('cp932')


# ---------------------------------------------------------------------------
# OVBA empty-chunk padding helpers (needed for in-place stream replacement)
# ---------------------------------------------------------------------------
# A 3-byte compressed chunk with header 0xB000 + flag byte 0 decompresses to
# zero bytes. We use these (and 5-byte variants) to pad compressed streams
# to the exact byte size required by olefile.write_stream.
_EMPTY3 = struct.pack('<H', 0xB000) + b'\x00'
_PAD5   = struct.pack('<H', 0xB002) + b'\x00\x00\x00'

def _pad_to_exact(compressed: bytes, target: int) -> bytes:
    diff = target - len(compressed)
    if diff < 0:
        raise ValueError(f"compressed ({len(compressed)}) > target ({target})")
    if diff == 0:
        return compressed
    for n5 in range(diff // 5 + 1):
        rem = diff - n5 * 5
        if rem >= 0 and rem % 3 == 0:
            return compressed + (_PAD5 * n5) + (_EMPTY3 * (rem // 3))
    raise ValueError(f"Cannot reach exact target byte size {target} from {len(compressed)}")


def patch_vba_in_place(vba_bin: bytes) -> bytes:
    """Return a patched vbaProject.bin where ThisWorkbook = installer source
       and dir's ThisWorkbook MOFFSET = 0. All other bytes preserved."""
    skel = make_xlsm.CFBReader(vba_bin)

    # 1. Patch dir: set ThisWorkbook MOFFSET to 0
    dir_dec = make_xlsm.ovba_decompress(skel.read('dir'))
    needle = struct.pack('<HI', 0x0019, len('ThisWorkbook')) + b'ThisWorkbook'
    idx = dir_dec.find(needle)
    if idx < 0:
        raise RuntimeError("ThisWorkbook MNAME not found in dir stream")
    i = idx
    found = False
    while i < len(dir_dec):
        rid = struct.unpack('<H', dir_dec[i:i+2])[0]
        sz  = struct.unpack('<I', dir_dec[i+2:i+6])[0]
        if rid == 0x0031:  # MOFFSET
            patched = bytearray(dir_dec)
            struct.pack_into('<I', patched, i+6, 0)
            dir_dec = bytes(patched)
            found = True
            break
        i += 6 + sz
    if not found:
        raise RuntimeError("ThisWorkbook MOFFSET record not found")

    orig_dir_size = skel.entries['dir']['size']
    orig_tw_size  = skel.entries['ThisWorkbook']['size']
    new_dir = _pad_to_exact(make_xlsm.ovba_compress(dir_dec), orig_dir_size)
    new_tw  = _pad_to_exact(make_xlsm.ovba_compress(INSTALLER_SRC), orig_tw_size)

    # 2. Write streams in place via olefile
    buf = io.BytesIO(vba_bin)
    ole = olefile.OleFileIO(buf, write_mode=True)
    ole.write_stream('VBA/dir', new_dir)
    ole.write_stream('VBA/ThisWorkbook', new_tw)
    ole.close()
    buf.seek(0)
    return buf.read()


# ---------------------------------------------------------------------------
# Data: config defaults / departments / system prompts
# ---------------------------------------------------------------------------
CONFIG_ROWS = [
    ("router_max_chunks",   8,    "ルーターが選ぶ関連チャンク数の上限"),
    ("verifier_enabled",    True, "自己検証パスを実行するか (True=精度優先 / False=コスト優先)"),
    ("mock_llm",            False, "True=社内AIリボンを呼ばずダミー応答（動作確認用 / Mac可）。本番はFalse"),
    ("max_context_chars",   60000, "ドラフトに渡すコンテキストの最大文字数"),
    ("debug_mode",          False, "True にすると全LLMプロンプトをdebug_logシートに残す"),
    ("feedback_top_n",      3,    "ドラフトに参照させる過去Q&Aの最大件数"),
    ("feedback_min_score",  0.18, "類似Q&A採用の閾値 (0..1)"),
    ("router_model",        "gpt-5.5", "(備考) ルーター用モデル"),
    ("drafter_model",       "gpt-5.5", "(備考) ドラフト用モデル"),
    ("verifier_model",      "gpt-5.5", "(備考) 検証用モデル"),
    ("draft_temperature",   0.15, "(備考) ドラフト温度"),
    ("department_id",       "",   "初回起動時に設定される"),
    ("department_name",     "",   ""),
    ("role",                "",   ""),
]
DEPARTMENTS = [("expense_profit", "費用利益保険チーム", "common,expense_profit")]

ROUTER_PROMPT = '''あなたは社内ナレッジ検索のルーターです。
ユーザーの質問に対し、提供されたナレッジ一覧から最も関連性の高いチャンクのIDを最大{max_n}件、JSON形式で返してください。

【最重要原則】
ナレッジには「書類種別」が複数あります(普通保険約款 / 引受ガイドライン / 研修資料 / 解説 / FAQ / ハンドブック)。
質問の意図に応じて適切な書類種別を選んでください。書類種別の偏りに注意：

- 「○○とは」「教えて」「どんな保険」「どんなケース」「どんな逸失利益」など
  説明・解説を求める質問
    → 研修資料 / 解説 / FAQ / ハンドブック / 引受ガイドライン を優先
    → 普通保険約款"だけ"で答えない（約款は契約上の規定で、解説や事例は含まれない）

- 「○○の引受基準」「リスク」「引受可否」「引受要件」「対象外」など
  引受判断に関する質問
    → 引受ガイドライン / 研修資料 を優先

- 「○条」「条文」「規定」「定義」「支払対象額の計算式」など
  約款上の規定を尋ねる質問
    → 普通保険約款 を優先

- 上記いずれでもなければ、複数の書類種別を組み合わせて多角的に選ぶ

【選び方】
- 質問のキーワード(保険種類名、用語、トピック)に直接触れているチャンクを最優先
- 同じ業務領域(domain)内で複数の書類種別から選ぶ。約款だけで埋めない
- summary と keywords を見て、質問と意味的に近いものを選ぶ

【ユーザー質問】
{question}

【ナレッジ一覧】
形式: chunk_id | 業務領域 | 書類種別 | 出典 | 1行要約 | キーワード

{knowledge_table}

【出力 (JSON、他のテキストは絶対に書かない)】
{
  "selected_ids": ["sample_03.pdf::p5::c2", "..."],
  "reasoning": "なぜこれらを選んだか1〜2文。書類種別の組み合わせ理由も。"
}'''

DRAFTER_PROMPT = '''あなたは社内アンダーライターの業務支援AIです。提示された社内ナレッジ抜粋および認定済みQ&Aのみを根拠に回答してください。

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
・ 営業担当者の視点で平易に説明する。専門用語には括弧で短い説明を添える。
・ 過去の認定済みQ&Aがある場合は、それを最優先で参照しつつ、新しい質問に合わせて再構成する。'''

VERIFIER_PROMPT = '''あなたは損害保険会社引受部門の品質管理者です。AIが書いたドラフト回答が、引用元ナレッジに本当に書かれている内容のみで構成されているか厳格に検証します。

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
・ ドラフトが全面的にナレッジに無い内容だった場合は『社内ナレッジに該当する記載がありません。アンダーライターへ確認してください』のみ返す。'''

SYSTEM_PROMPTS = [
    ("router",   ROUTER_PROMPT,   "Step1: 関連チャンクを選ぶ"),
    ("drafter",  DRAFTER_PROMPT,  "Step3: 構造化された回答を生成"),
    ("verifier", VERIFIER_PROMPT, "Step4: 各主張をナレッジと照合し誤りを削除/修正"),
]


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------
def _make_main_sheet(ws):
    """Repurpose skeleton's existing Sheet1 as the main sheet."""
    ws.title = "main"
    # Clear any existing content
    for row in ws['A1:Z100']:
        for c in row:
            c.value = None

    ws['A1'] = "社内ナレッジ QA ボット (v2)"
    ws['A1'].fill = PatternFill("solid", fgColor="3C5AA0")
    ws['A1'].font = Font(bold=True, size=16, color="FFFFFF")
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    ws['A3'] = "▼ 起動状況"
    ws['A3'].font = Font(bold=True, size=11, color="CC0000")

    notes = [
        "",
        "・このファイルを開くと、自動でモジュールがインストールされ、チャットUIが表示されます。",
        "・「セットアップ完了」のメッセージが出ない場合は、トラストセンターで以下を有効にしてください:",
        "  [ファイル]→[オプション]→[トラストセンター]→[トラストセンターの設定]→[マクロの設定]",
        "  → 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェック → OK → Excel再起動",
        "",
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
        "■ このシステムの仕組み",
        "  ・費用利益保険チームのPDFマニュアル15冊(813チャンク)を搭載",
        "  ・社内AIリボン(ChatGPT関数)で回答を4段階生成 (ルート→検索→ドラフト→検証)",
        "  ・○良かった/×修正ボタンでQ&Aを蓄積し精度向上",
        "",
        "■ 重要",
        "  ・回答は必ずアンダーライターの最終確認を取ること",
        "  ・契約番号・氏名・電話番号など個人情報を質問に含めないこと",
    ]
    for i, line in enumerate(notes, start=4):
        c = ws.cell(row=i, column=1, value=line)
        if line.startswith("■"):
            c.font = Font(bold=True, size=11)
        elif line.startswith("━"):
            c.font = Font(color="888888")
        else:
            c.font = Font(size=11)
    ws.column_dimensions['A'].width = 90


def _make_config(wb):
    ws = wb.create_sheet("config")
    for c, h in enumerate(["key", "value", "description"], 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
    for i, (k, v, d) in enumerate(CONFIG_ROWS, 2):
        ws.cell(row=i, column=1, value=k)
        ws.cell(row=i, column=2, value=v)
        ws.cell(row=i, column=3, value=d)
    ws.column_dimensions['A'].width = 24
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 70


def _make_system_prompt(wb):
    ws = wb.create_sheet("system_prompt")
    for c, h in enumerate(["key", "prompt", "description"], 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
    for i, (k, p, d) in enumerate(SYSTEM_PROMPTS, 2):
        ws.cell(row=i, column=1, value=k)
        ws.cell(row=i, column=2, value=p).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=i, column=3, value=d)
        ws.row_dimensions[i].height = 240
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 80
    ws.column_dimensions['C'].width = 40


def _make_department(wb):
    ws = wb.create_sheet("department")
    for c, h in enumerate(["dept_id", "dept_name", "knowledge_scope_csv"], 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
    for i, row in enumerate(DEPARTMENTS, 2):
        for c, v in enumerate(row, 1):
            ws.cell(row=i, column=c, value=v)
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 28
    ws.column_dimensions['C'].width = 40


def _make_manifest(wb, chunks):
    ws = wb.create_sheet("manifest")
    headers = ["source", "display", "domain", "doc_type",
               "dept_scope", "chunk_count", "notes"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)

    by_src = {}
    for c in chunks:
        by_src.setdefault(c["source"], []).append(c)

    row = 2
    for src in sorted(by_src):
        cks = by_src[src]
        first = cks[0]
        ws.cell(row=row, column=1, value=src)
        ws.cell(row=row, column=2, value=first.get("display", src))
        ws.cell(row=row, column=3, value=first.get("domain", "未分類"))
        ws.cell(row=row, column=4, value=first.get("doc_type", "未分類"))
        ws.cell(row=row, column=5, value="common,expense_profit")
        ws.cell(row=row, column=6, value=len(cks))
        row += 1
    widths = [18, 48, 24, 18, 28, 14, 30]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _make_knowledge_base(wb, chunks):
    ws = wb.create_sheet("knowledge_base")
    headers = ["chunk_id", "source", "display", "domain", "doc_type",
               "section_header", "summary", "keywords", "full_text", "dept_scope"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
    for i, c in enumerate(chunks, 2):
        ws.cell(row=i, column=1, value=c["id"])
        ws.cell(row=i, column=2, value=c["source"])
        ws.cell(row=i, column=3, value=c.get("display", c["source"]))
        ws.cell(row=i, column=4, value=c.get("domain", "未分類"))
        ws.cell(row=i, column=5, value=c.get("doc_type", "未分類"))
        ws.cell(row=i, column=6, value=c.get("header", ""))
        ws.cell(row=i, column=7, value=c.get("summary", ""))
        ws.cell(row=i, column=8, value=c.get("keywords", ""))
        ws.cell(row=i, column=9, value=_clean(c.get("text", ""))[:32000])
        ws.cell(row=i, column=10, value="common")
    widths = [28, 18, 40, 22, 18, 50, 60, 40, 80, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    print(f"  knowledge_base: {len(chunks)} rows")


def _make_feedback(wb):
    ws = wb.create_sheet("feedback")
    headers = ["fb_id", "timestamp", "dept_id", "submitter",
               "question", "answer", "status", "approver",
               "approved_at", "correction", "tags", "source_chunk_ids"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
    widths = [22, 20, 18, 14, 50, 80, 12, 14, 20, 80, 24, 60]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _make_usage_log(wb):
    ws = wb.create_sheet("usage_log")
    headers = ["timestamp", "dept_id", "user",
               "q_chars", "a_chars", "router_ms", "draft_ms",
               "verify_ms", "total_ms", "chunks_used", "feedback_status"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
    widths = [20, 18, 14, 10, 10, 10, 10, 10, 10, 60, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _make_vba_src(wb):
    """Hidden sheet carrying every standard module's source code."""
    ws = wb.create_sheet("vba_src")
    for c, h in enumerate(["module_name", "type", "source"], 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)

    mods = ["modBoot", "modConfig", "modUserProfile",
            "modRibbonGateway", "modKnowledgeBase",
            "modPrompts", "modFeedbackLookup", "modPipeline",
            "modFeedback", "modChatUI", "modPii", "modUsageLogger",
            "modDiag"]

    EXCEL_CELL_LIMIT = 32000
    for i, name in enumerate(mods, 2):
        bas_path = os.path.join(SRC_V2, name + ".bas")
        with open(bas_path, encoding="utf-8") as fp:
            txt = fp.read()
        out_lines = []
        for line in txt.split("\n"):
            stripped = line.lstrip("﻿")
            if stripped.lstrip().startswith("Attribute "):
                continue
            out_lines.append(stripped)
        cleaned = "\n".join(out_lines)
        if len(cleaned) >= EXCEL_CELL_LIMIT:
            raise RuntimeError(
                f"{name}.bas ({len(cleaned)} chars) exceeds {EXCEL_CELL_LIMIT}-char cell limit")
        ws.cell(row=i, column=1, value=name)
        ws.cell(row=i, column=2, value="std")
        ws.cell(row=i, column=3, value=cleaned)

    ws.column_dimensions['A'].width = 24
    ws.column_dimensions['B'].width = 8
    ws.column_dimensions['C'].width = 80
    ws.sheet_state = "hidden"
    print(f"  vba_src: embedded {len(mods)} module sources")


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------
def main():
    if not os.path.exists(ENRICHED):
        sys.exit(f"Missing {ENRICHED} — run build/enrich_chunks.py first")

    print("=== build_chatbot_v2.py (in-place VBA patching) ===")
    print(f"Template:    {TEMPLATE}")
    print(f"Source v2:   {SRC_V2}")
    print(f"Output:      {OUT_XLSM}")
    print()

    os.makedirs(os.path.dirname(OUT_XLSM), exist_ok=True)
    chunks = json.load(open(ENRICHED, encoding="utf-8"))

    # Stage 1: load skeleton, preserve vbaProject.bin byte-for-byte
    print("Stage 1: load skeleton with keep_vba=True...")
    wb = openpyxl.load_workbook(TEMPLATE, keep_vba=True)
    print(f"  Initial sheets: {wb.sheetnames}")

    # Stage 2: build all data sheets
    print("Stage 2: build data sheets...")
    _make_main_sheet(wb[wb.sheetnames[0]])  # repurpose Sheet1 → main
    _make_config(wb)
    _make_system_prompt(wb)
    _make_department(wb)
    _make_manifest(wb, chunks)
    _make_knowledge_base(wb, chunks)
    _make_feedback(wb)
    _make_usage_log(wb)
    _make_vba_src(wb)
    print(f"  Sheets after build: {wb.sheetnames}")

    # Stage 3: save with openpyxl (preserves binary)
    print("Stage 3: openpyxl save (keep_vba)...")
    with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as tmp:
        tmp_path = tmp.name
    wb.save(tmp_path)
    print(f"  Saved {tmp_path} ({os.path.getsize(tmp_path):,} bytes)")

    # Stage 4: in-place patch vbaProject.bin (installer in ThisWorkbook)
    print("Stage 4: in-place patch vbaProject.bin...")
    with zipfile.ZipFile(tmp_path, 'r') as zin:
        parts = {n: zin.read(n) for n in zin.namelist()}
    skel_bin = parts['xl/vbaProject.bin']
    patched_bin = patch_vba_in_place(skel_bin)
    parts['xl/vbaProject.bin'] = patched_bin
    print(f"  Original vbaProject.bin: {len(skel_bin):,} bytes")
    print(f"  Patched  vbaProject.bin: {len(patched_bin):,} bytes")
    assert len(skel_bin) == len(patched_bin), "size mismatch — binary integrity broken"

    # Stage 5: write final .xlsm
    print("Stage 5: write final xlsm...")
    with zipfile.ZipFile(OUT_XLSM, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for n, data in parts.items():
            zout.writestr(n, data)
    os.unlink(tmp_path)
    print(f"  Final: {OUT_XLSM} ({os.path.getsize(OUT_XLSM):,} bytes)")
    print("\nDone.")


if __name__ == '__main__':
    main()
