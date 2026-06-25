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
  ' Save failures (read-only file, locked share, etc.) must not skip Boot.
  ' Swallow any save error so Application.Run still fires.
  On Error Resume Next
  ThisWorkbook.Save
  Err.Clear
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

【最優先：トピック整合性】
質問のトピックと無関係なチャンクを「多様性のため」だけで入れない。
- 「興行中止保険」の質問にリコールプロテクション（生産物回収費用保険）のチャンクは入れない。
- 「瑕疵保証責任保険」の質問に家主費用保険のチャンクは入れない。
- 質問のキーワードと、チャンクの domain（業務領域）や summary・keywords が全くかみ合わなければ、その doc_type を入れたくても **除外** する。
- 関連 domain は「費用利益保険全般」など、質問対象を包含する一般カテゴリも含む。

【選定の基本方針】
回答に必要な「定義」「除外」「算定方法」「具体例」を **横断的に** 選びます。1つの条文だけでは答えにならない質問が多いため、関連条文を **セット** で選ぶことを優先します。

例：「興行中止保険でグッズ代は補償対象か？」という質問なら、
  - 損害の定義条文（何が対象か）
  - 除外条文（対象外は何か）
  - 売上不足など不払い条文
  - 損害額の算定条文
  を費用利益保険ドメインから一緒に選ぶ。
  （他の保険種類のハンドブックなどはトピック違いで入れない）

【書類種別ごとの使い分け】
ナレッジには「普通保険約款 / 引受ガイドライン / 研修資料 / 解説 / FAQ / ハンドブック」があります：

- 「○○とは」「教えて」「どんな〜」など説明系の質問
    → 研修資料 / 解説 / FAQ / ハンドブック を最優先で組み入れる（**ただしトピックが合うものに限る**）
    → 普通保険約款"だけ"で埋めない
- 「引受基準」「リスク」「引受可否」「対象外」「引受要件」
    → 引受ガイドライン / 研修資料 を優先
- 「○条」「条文」「規定」「定義」「算定方法」「支払額」
    → 普通保険約款 を優先（補足的に研修資料・解説・ハンドブックも1〜2件含めるが、**トピックが合うものに限る**）

【多様性ルール】
- トピック整合性を満たした上で、可能なら2つ以上の source PDF から選ぶ。
- ただしトピック合致が1つの PDF にしかない場合は、無理に別の PDF を入れない。

【ユーザー質問】
{question}

【ナレッジ一覧】
形式: chunk_id | 業務領域 | 書類種別 | 出典 | 1行要約 | キーワード

{knowledge_table}

【出力 (JSON、他のテキストは絶対に書かない)】
{
  "selected_ids": ["sample_03.pdf::p5::c2", "..."],
  "reasoning": "選定したチャンクのトピック関連性と、書類種別の組み合わせ理由を1〜2文"
}'''

DRAFTER_PROMPT = '''あなたは費用利益保険を熟知したベテランアンダーライターです。営業担当者の「これは塡補されるのか？」という問いに、約款の枠組みと損保実務の専門知識を総動員して、**具体的で実務的な結論**を返します。NotebookLM のように、質問の本質を捉え、推論を効かせ、読みやすく答えてください。

【最重要：抽象的な逃げを禁止する】
「保険証券記載の費用・収益に含まれるかで決まります」だけで終わるのは**禁止**です。それは誰でも知っている当たり前の話で、質問者が本当に知りたいのは「**その項目（例: グッズ原価、チケット払戻手数料）が、実務上 塡補対象になりやすいのか、なりにくいのか、その理由**」です。
必ず **具体的な見解（原則対象 / 原則対象外 / 条件次第だが主たる傾向は◯◯）** を示してください。

【回答の作り方】
1. **結論**：質問された個別項目それぞれについて、塡補可否の見解を一言で言い切る（「原則として対象外と考えられます」等）。
2. **約款の枠組み**：根拠となる条文を引用し、出典マーカー [#N] を付ける。ここは約款に書いてある事実のみ。
3. **当てはめ（ここが核心）**：その枠組みに、質問の個別項目を当てはめて推論する。損保実務の一般的な考え方を使ってよい。例えば：
   - 「逸失利益」として塡補されやすいのは、**契約時点で金額が客観的に確定している収益**（例: 放映権料、前売券の確定済み売上）。
   - 逆に、**当日の天候・客足で変動する不確実な見込み収益**（例: 当日券、グッズ物販、募金）は、確定性を欠くため塡補対象になりにくい。
   - 原価・仕入れ代金は「収益の喪失」ではなく「費用損害」の議論。費用として塡補されるかは、それが興行実施のための費用か、無駄になったサンクコストか等で考える。
   この当てはめ部分は「**約款の条文そのもの**」ではなく「**実務上の一般的な解釈**」であることが分かるように、【実務上の考え方】のように見出しや前置きで区別する。
4. **留意点**：実務で確認・交渉すべき点があれば簡潔に。

【約款にある事実 と 実務上の推論 を区別する】
- 約款に明記された定義・除外・算定方法 → 出典マーカー [#N] を付けて断定。
- あなたの専門知識による推論・一般的な実務慣行・具体例 → 「実務上は」「一般的に」「と考えられます」と明示し、断定しすぎない。出典マーカーは付けない（約款の引用ではないため）。
- この2つを混ぜず、読み手が「どこまでが約款で、どこからが解釈か」を区別できるようにする。

【スタイル】
- 見出しに絵文字（🎯💰🛍️🎫💡⭕❌など）を使い、視覚的に整理。
- 重要な結論は **太字**。
- 「結論から言うと」など自然な語り口。
- 表が有効なら使う（| 項目 | 塡補可否 | 理由 | の3列など）。
- 質問のかたちに応じて構成は柔軟に。テンプレを機械的に当てはめない。

【守るべき一線（ハルシネーション防止）】
- 約款に**無い**具体的な数字・割合・期間・金額・条文番号を、約款の事実であるかのように書かない。
- ただし一般的な実務知識に基づく**定性的な推論・具体例**（放映権は確定収益、グッズは変動収益 等）は積極的に活用してよい。これは禁止される「捏造」ではなく「専門的解釈」。
- 個人情報（契約番号・氏名・電話番号等）が質問に含まれる場合は「個人情報を含めないでください」とだけ返す。

【出力例の雰囲気（あくまで例。質問に応じて変える）】
```
## 結論 🎯
**グッズの仕入れ原価**：原則として塡補対象になりにくいと考えられます。
**前払いチケットの払戻手数料**：中止対応に伴う費用として、条件次第で対象になり得ます。

## 約款の枠組み
第2条で損害は「費用の負担」または「収益の喪失」と定義され…（引用）[#3]

## 💡 実務上の考え方：なぜそうなるのか
逸失利益として塡補されやすいのは契約時点で金額が確定した収益です。グッズ物販は当日の客足で変動するため…（推論）
```'''

VERIFIER_PROMPT = '''あなたはシニアアンダーライターです。ドラフト回答を「軽く」レビューします。完璧主義で削りすぎないこと。NotebookLM のような具体性・推論の豊かさを保ったまま、明確な誤りだけを正します。

【この回答に期待される性質】
- 約款の枠組みを土台にしつつ、損保実務の専門知識で「塡補対象になりやすい/なりにくい」という**具体的な見解**を示している。
- 「約款に書いてある事実」と「実務上の推論・一般的解釈」が区別されている。
この2つの性質は **守るべき長所** であり、削ってはいけません。

【修正する（これだけ）】
1. 約款の**事実として**提示された数字・割合・期間・条文番号が、引用元ナレッジに無い → 削除または「（約款上は記載なし）」に。
2. 条件分岐の取り違え（「ただし」「〜を除く」「〜に限り」を逆に解釈）。
3. 否定/限定の逆転（「支払う」と「支払わない」を取り違え）。
4. 出典マーカー [#N] が指す chunk と本文の主張が一致しない → マーカーを正す。
5. 明らかな事実誤認（保険種類の取り違え 等）。

【絶対に削除・改変しない】
- 実務上の推論・一般的な解釈（「放映権は確定収益なので対象、グッズ物販は変動収益なので対象外になりやすい」等）。これらは「実務上は」「一般的に」と明示されていれば**正しい付加価値**であり、約款外でも残す。
- 具体的な結論・見解（「原則対象外と考えられます」等）。「証券次第」へ後退させない。
- ⭕❌🎯💡 等の絵文字、太字、表、見出し構成、自然な語り口、深掘りセクション。

【判断基準】
- 「実務上は」「一般的に」「と考えられます」と明示された推論 → **残す**（出典マーカーが無くても誤りではない）。
- 約款の事実として断定されているのに約款に無い → 直す。
要は「専門家としての解釈」は活かし、「約款の捏造」だけを潰す。

【出力】
- 検証後の最終回答**のみ**を出力（プロセス説明・前置きを書かない）。
- ドラフトの見出し・絵文字・表・スタイルを維持。
- 修正した場合のみ末尾に `※検証で修正: <要点>` を1行追記。修正なしなら何も追記しない。
- 引用元に該当が全く無い場合のみ「社内ナレッジに該当する記載がありません。アンダーライターへ確認してください」と返す。'''

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


def _make_howto(wb):
    """Quick-start guide. Lives at sheet index 1 (after main) so users see
    it first when they tab over from the chat sheet."""
    ws = wb.create_sheet("使い方")
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 90

    # Banner
    ws['A1'] = "使い方ガイド"
    ws.merge_cells('A1:B1')
    ws['A1'].fill = PatternFill("solid", fgColor="3C5AA0")
    ws['A1'].font = Font(bold=True, size=16, color="FFFFFF")
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # (label, body) pairs rendered as 2 columns
    sections = [
        ("基本の流れ",
         "1. main タブの質問欄に知りたいことを入力\n"
         "2. 「送信」ボタンを押す → 60〜90秒待つ\n"
         "3. 答えが表示される（出典・引用箇所も同時に表示）\n"
         "4. 答えに対して「○良かった」または「×修正」で評価する\n"
         "5. 「続けて質問」で深掘りができる（NotebookLM 風の追加対話）"),

        ("○×評価が重要な理由",
         "ボタンを押すたびに、ボットは「どんな質問にどう答えるべきか」を学習します。\n"
         "  ⭕「良かった」: その質問と回答が『模範回答』として feedback シートに保存される\n"
         "  ❌「修正」  : その質問と回答が『要修正例』として保存され、次回類似質問時に避ける\n"
         "評価を積み重ねるほど、似た質問に対して過去の正解を参照し精度が上がります。\n"
         "使い始めの数十回が肝心。気になった答えはすかさず ○ か × を押してください。"),

        ("続けて質問（深掘り）",
         "答えを読んで「もっと知りたい」「ここはどう？」と思ったら『続けて質問』ボタン。\n"
         "前回の質問と回答を踏まえて、より深い答えが返ります。\n"
         "例:\n"
         "  最初の質問: 興行中止保険でグッズ代は補償対象？\n"
         "  → 続けて: じゃあ前売券の払い戻し手数料は？\n"
         "  → 続けて: その手数料は誰が負担する慣行？"),

        ("良い質問の例",
         "⭕ 「興行中止保険でグッズ代は補償対象？」          ← 具体的・はい/いいえ系\n"
         "⭕ 「瑕疵保証責任保険の引受基準を教えて」         ← 特定の保険・特定の論点\n"
         "⭕ 「費用利益保険と利益保険の違いを表で比較して」  ← 比較系\n"
         "❌ 「保険のこと教えて」                          ← 漠然すぎる\n"
         "❌ 「お客様 山田太郎さんの保険」                  ← 個人情報を含む"),

        ("⚠️ 触ってはいけないタブ",
         "下のタブはシステムが自動で管理しています。**直接編集すると壊れます**：\n"
         "  ・config        : 設定値 (管理者のみ変更可)\n"
         "  ・system_prompt : AI への指示文 (管理者のみ)\n"
         "  ・manifest      : 知識ベースの一覧 (自動生成)\n"
         "  ・knowledge_base: 813件の知識データ本体 (自動生成)\n"
         "  ・feedback      : ○×評価ログ (ボタンが自動で書く)\n"
         "  ・usage_log     : 利用ログ (自動で書く)\n"
         "  ・vba_src       : 非表示。プログラムソース（再起動時に使う）"),

        ("✅ 触ってよいタブ",
         "  ・使い方     : このシート（参照のみ）\n"
         "  ・main      : 質問と回答（普段の作業場所）\n"
         "  ・department: 部署設定（初回入力済み。変更したい時のみ）"),

        ("⚠️ 個人情報・機密情報を入れないでください",
         "契約番号、氏名、電話番号、マイナンバー、E メールアドレスを質問に書かないでください。\n"
         "個人情報を検知すると警告が出ます。警告が出たら必ず書き換えてから送信してください。\n"
         "ボットは社内 AI リボン経由でクラウドの LLM を呼びます。送信した本文は社外には残りませんが、\n"
         "**社内ログには記録される** 前提で書いてください。"),

        ("⚠️ AI の答えは最終決定ではありません",
         "回答は社内ナレッジ（15PDF / 813チャンク）を根拠に AI が生成しています。\n"
         "条文の解釈や引受可否の最終判断は、必ずアンダーライターの確認を取ってください。\n"
         "特に金額・期間・割合・条文番号は、必ず原本（保険証券・約款）と突き合わせて確認してください。"),

        ("困ったときは",
         "  ・回答が出ない/エラー    → main の『自己診断』ボタンを押す\n"
         "  ・AI リボンが繋がらない  → main の『リボン接続テスト』ボタン\n"
         "  ・ボットが反応しない    → main の『再起動』ボタン\n"
         "  ・全部ダメ              → Excel 終了 → ファイル開き直す"),
    ]

    row = 3
    label_font = Font(bold=True, size=11, color="3C5AA0")
    body_font = Font(size=10)
    warning_fill = PatternFill("solid", fgColor="FFF4D6")
    section_fill = PatternFill("solid", fgColor="F4F6FB")

    for label, body in sections:
        is_warning = "⚠️" in label or "触ってはいけない" in label

        ws.cell(row=row, column=1, value=label).font = label_font
        ws.cell(row=row, column=1).alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(row=row, column=2, value=body).font = body_font
        ws.cell(row=row, column=2).alignment = Alignment(vertical="top", wrap_text=True)
        if is_warning:
            ws.cell(row=row, column=1).fill = warning_fill
            ws.cell(row=row, column=2).fill = warning_fill
        else:
            ws.cell(row=row, column=1).fill = section_fill
            ws.cell(row=row, column=2).fill = section_fill

        # Approximate height
        n_lines = body.count("\n") + 1
        ws.row_dimensions[row].height = max(28, 16 * n_lines + 6)
        row += 1

    # Move 使い方 to position 1 (right after main)
    sheet_order = wb.sheetnames
    if "使い方" in sheet_order and sheet_order[1] != "使い方":
        idx = sheet_order.index("使い方")
        sheets = wb._sheets
        sheets.insert(1, sheets.pop(idx))


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
    _make_howto(wb)
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
