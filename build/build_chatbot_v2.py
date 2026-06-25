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
    ("router_max_chunks",   12,   "ルーターが選ぶ関連チャンク数の上限（精度重視で 12 推奨）"),
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
ユーザーの質問に対し、最も関連するチャンクIDを最大{max_n}件、JSON形式で返してください。

【選定の基本方針】
回答に必要な「定義」「除外」「算定方法」「具体例」を**横断的に**選びます。1つの条文だけでは答えにならない質問が多いため、関連条文を**セット**で選ぶことを優先します。

例：「グッズ代は補償対象か？」という質問なら、
  - 損害の定義条文（何が対象か）
  - 除外条文（対象外は何か）
  - 売上不足などの不払い条文
  - 損害額の算定条文
  を一緒に選ぶ。

【書類種別ごとの使い分け】
ナレッジには「普通保険約款 / 引受ガイドライン / 研修資料 / 解説 / FAQ / ハンドブック」があります：

- 「○○とは」「教えて」「どんな〜」など説明系の質問
    → 研修資料 / 解説 / FAQ / ハンドブック を最優先で組み入れる
    → 普通保険約款"だけ"で埋めない（約款は法律的な定義のみで、事例や解説は含まれない）
- 「引受基準」「リスク」「引受可否」「対象外」「引受要件」
    → 引受ガイドライン / 研修資料 を優先
- 「○条」「条文」「規定」「定義」「算定方法」「支払額」
    → 普通保険約款 を優先（ただし、補足的に研修資料・解説・ハンドブックも1〜2件含める）
- 上記いずれでもなければ、複数の書類種別を組み合わせる

【多様性ルール（必須）】
- 同じ source（PDFファイル）ばかりで埋めない。最低2つの source から選ぶ。
- 普通保険約款だけで全件を埋めない。可能なら研修資料・解説・ガイドライン・FAQ・ハンドブックから少なくとも1〜2件は含める。
- 同じ業務領域(domain)内なら複数のページから選ぶ。

【ユーザー質問】
{question}

【ナレッジ一覧】
形式: chunk_id | 業務領域 | 書類種別 | 出典 | 1行要約 | キーワード

{knowledge_table}

【出力 (JSON、他のテキストは絶対に書かない)】
{
  "selected_ids": ["sample_03.pdf::p5::c2", "..."],
  "reasoning": "選定した書類種別の組み合わせと、なぜそれらを組み合わせたか1〜2文"
}'''

DRAFTER_PROMPT = '''あなたは費用利益保険のベテランアンダーライターです。NotebookLM が出すような、「質問の本質を捉え、論理を通し、読みやすく実務で使える」回答を書きます。

【目指す品質】
- 質問の **本当に知りたいこと** を読み解き、結論からズバッと答える
- 「結論から言うと」「ポイントは」など人が話す自然な言い回し
- 見出しは ## と絵文字（⭕❌💡🛍️🎯💰など）で視覚的に整理
- 重要な結論は **太字** で強調
- 必要なら「💡 深掘り」「💡 補足」のような節で背景や含意を加える
- 構成は質問の性質に合わせて柔軟に。テンプレートを機械的に当てはめない
- 引用元にない内容は推測しないが、引用元の用語に業界常識で具体例を補うのは積極的に行う

【絶対のルール】
- 引用元ナレッジ（## 社内ナレッジ抜粋）が根拠の唯一の素材
- 主張の末尾に出典マーカー [#1] を必ず挿入
- 条文を **組み合わせて** 結論を出す（例：「第2条で〜と定義され、第3条で〜と除外されているので、〜は対象外」）
- 引用元に **無い** 具体的な数字・割合・期間・金額・条文番号は記載しない
- 「ご質問は〜ですね」「最後に〜について確認してください」「本社アンダーライティング部に確認」などの **定型句は書かない**
- 個人情報（契約番号・氏名・電話番号・マイナンバー等）が含まれる質問は「個人情報を含めないでください」とだけ返す

【参考スタイル（あくまで例。質問に応じて柔軟に変えてよい）】

質問: 「○○保険で △△ は補償対象ですか？」

```
## 結論
**結論から言うと**、△△は **補償の対象外** です [#2]。

⭕ **補償対象** : 〜（例: 〜） [#1]
❌ **対象外**  : △△、〜（理由: 〜） [#2]

## なぜそうなるのか
第◯条（損害の定義）では、補償される損害は「〜」と限定されています [#1]。一方、第△条では「〜は含まない」と明記されています [#2]。
△△は性質上「〜」に該当するため、定義の網に入りません。

## 💡 深掘り：似たケースとの違い
（あれば 1〜2 段落で背景や論理を補足）

## 🎯 実務で確認すべきこと
- 保険証券の「〜」欄を確認 [#3]
```

質問が単純な定義系（「○○とは？」）であれば、見出し1〜2個で簡潔に答えればよい。
質問が複合的（「○○と△△の違いは？」）なら、比較表で並べる。
質問が手順系（「○○はどう進める？」）なら、番号付きステップで。

要は **質問のかたちに合わせて、最も読みやすい構成を選ぶ**。'''

VERIFIER_PROMPT = '''あなたはシニアアンダーライターです。ドラフト回答が引用元ナレッジを正しく解釈しているかを「軽く」確認します。完璧主義で削りすぎない。NotebookLM のような豊かさを保ちつつ、明らかな誤りだけを直します。

【修正対象（これだけ）】
1. 引用元ナレッジに **無い** 具体的な数字・割合・期間・金額・条文番号 → 削除または「（記載なし）」
2. 条件分岐の取り違え（「ただし」「〜の場合を除く」「〜に限り」を逆に解釈）
3. 否定/限定の逆転（「支払う」と「支払わない」を取り違え）
4. 出典マーカー [#N] が指す chunk と本文の主張が一致しない場合 → マーカーを正しい番号に
5. 保険種類の取り違えなど、明らかな事実誤認

【絶対に削除・修正しない（NotebookLM の質感を保つため）】
- 複数の条文を組み合わせた論理的推論
- 条文の用語に対する業界一般の具体例補足（例: 条文に「収益」とあるところに「物販収入など」と例示）
- ⭕❌🎯💡 などの絵文字、太字、見出し構成
- 「結論から言うと」「ポイントは」のような自然な語り口
- 「深掘り」「補足」「背景」など、ドラフトが付け加えた洞察セクション
- 業界実務上の解釈（条文の趣旨から自然に導けるもの）

【出力】
- 検証後の最終回答 **のみ** を出力（プロセス説明や前置きを書かない）
- ドラフトの見出し構成・絵文字・スタイルはそのまま維持
- 修正があった場合のみ末尾に `※検証で修正: <要点>` と1行追記。修正なしなら何も追記しない
- 引用元に該当が全くない場合のみ「社内ナレッジに該当する記載がありません。アンダーライターへ確認してください」と返す'''

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
