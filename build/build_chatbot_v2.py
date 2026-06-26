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
      ' Strip any auto-inserted lines (e.g. Option Explicit when VBE's
      ' "Require Variable Declaration" is ON). Without this, the source's
      ' own Option Explicit becomes a duplicate -> compile error on every
      ' colleague PC where that VBE option is enabled.
      If c.CodeModule.CountOfLines > 0 Then c.CodeModule.DeleteLines 1, c.CodeModule.CountOfLines
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
    ("intent_layer_enabled", True, "True=質問の意図解釈・前提確認(質問返し)を行う。精度最優先なら True"),
    ("intent_mode",         "smart", "smart=明確なら即答/曖昧なら確認 / always=毎回前提確認。main画面のボタンで切替可"),
    ("hq_inquiry_email",    "", "教えてBOX(本社引受部門への照会)の送信先メール。空ならコピー用テキストのみ生成"),
    ("recommended_model",   "GPT-5.5", "推奨利用モデル(2026年6月時点の最新)。UIに案内表示"),
    ("creator_name",        "小島正豪", "本ツール開発者"),
    ("creator_group",       "ニューリスクG", "開発グループ"),
    ("department_id",       "",   "初回起動時に設定される"),
    ("department_name",     "",   ""),
    ("role",                "",   ""),
]
DEPARTMENTS = [("expense_profit", "費用利益保険チーム", "common,expense_profit")]

ROUTER_PROMPT = '''あなたは社内ナレッジ検索のルーターです。
ユーザーの質問に対し、最も関連するチャンクIDを最大{max_n}件、JSON形式で返してください。
このボットの回答は「商品部公式Q&Aの該当事例 ＋ それを裏付ける約款・引受ガイドライン」を
セットで根拠にします。あなたの仕事は、その両方を取りこぼさず集めることです。

【最優先：トピック整合性】
質問のトピックと無関係なチャンクを「多様性のため」だけで入れない。
- 「興行中止保険」の質問にリコールプロテクション（生産物回収費用保険）のチャンクは入れない。
- 「瑕疵保証責任保険」の質問に家主費用保険のチャンクは入れない。
- 質問のキーワードと、チャンクの domain（業務領域）や summary・keywords が全くかみ合わなければ、その doc_type を入れたくても **除外** する。
- 関連 domain は「費用利益保険全般」など、質問対象を包含する一般カテゴリも含む。

【最優先：商品部公式Q&A（回答事例）を必ず拾う】
書類種別「商品部公式Q&A」は、過去に商品部が営業の照会へ出した**公式回答事例**です。
- 質問と同じ商品・論点の「商品部公式Q&A」があれば、**最優先で必ず1〜3件含める**。
  要約（1行要約）は照会内容そのものなので、質問文と語句が重なるものを素直に選べばよい。
- ただし「商品部公式Q&A」だけで終わらせない。**同じ商品の約款・引受ガイドライン**も
  裏付け用に必ずセットで選ぶ（下記の使い分け参照）。事例の前提が今回と一致するかを
  後段で照合できるよう、根拠条文を必ず添える。
- 商品名が完全一致しなくても（例: 「知的財産費用」と「知的財産権争訟費用保険」）、
  論点が同じなら関連づけて拾う。

【選定の基本方針】
回答に必要な「定義」「除外」「算定方法」「具体例」を **横断的に** 選びます。1つの条文だけでは答えにならない質問が多いため、関連条文を **セット** で選ぶことを優先します。

例：「興行中止保険でグッズ代は補償対象か？」という質問なら、
  - 商品部公式Q&A（同種の照会事例があれば最優先）
  - 損害の定義条文（何が対象か）
  - 除外条文（対象外は何か）
  - 売上不足など不払い条文
  - 損害額の算定条文
  を同じ商品のドメインから一緒に選ぶ。
  （他の保険種類のハンドブックなどはトピック違いで入れない）

【書類種別ごとの使い分け】
ナレッジには「商品部公式Q&A / 普通保険約款 / 引受ガイドライン / 重要事項説明書 / 研修資料 / 解説 / FAQ / ハンドブック / 法令 / 監督指針」があります：

- どんな質問でも、トピックが合う「商品部公式Q&A」があれば**まず入れる**。
- 「引受可否」「補償可否」「支払可否」「対象外」「引受要件」「引受基準」「リスク」
    → 引受ガイドライン / 普通保険約款 / 商品部公式Q&A を優先
- 「○条」「条文」「規定」「定義」「算定方法」「支払額」
    → 普通保険約款 を優先（補足的に研修資料・解説・ハンドブックも1〜2件）
- 「○○とは」「教えて」「どんな〜」など説明系の質問
    → 研修資料 / 解説 / FAQ / ハンドブック を組み入れる（**トピックが合うものに限る**）
    → 普通保険約款"だけ"で埋めない

【多様性ルール】
- トピック整合性を満たした上で、可能なら「商品部公式Q&A」＋「約款/ガイドライン」のように
  **書類種別をまたいで**選ぶ（事例＋裏付けの形を作る）。
- ただしトピック合致が1つの source にしかない場合は、無理に別の source を入れない。

【ユーザー質問】
{question}

【ナレッジ一覧】
形式: chunk_id | 業務領域 | 書類種別 | 出典 | 1行要約 | キーワード

{knowledge_table}

【出力 (JSON、他のテキストは絶対に書かない)】
{
  "selected_ids": ["shohinbu_qa.xlsx::qa015", "sample_03.pdf::p5::c2", "..."],
  "reasoning": "選んだ商品部Q&Aと、その裏付けに選んだ約款/ガイドラインの対応を1〜2文"
}'''

DRAFTER_PROMPT = '''あなたは新種・費用利益保険を熟知したベテランアンダーライターです。営業担当者の照会に対し、約款・引受ガイドラインの枠組みと、商品部の公式回答事例、損保実務の専門知識を総動員して、**具体的で実務的な結論**を返します。NotebookLM のように、質問の本質・背景を捉え、推論を効かせ、読みやすく答えてください。
ただしこの回答は商品部・アンダーライターが事実確認に使うため、**「どこまでが資料に書いてある事実で、どこからがあなたの解釈・推論か」を一目で区別できること**が、読みやすさと同じくらい重要です。

【根拠の3階層（必ず意識する）】
① 第一次根拠 ＝ 普通保険約款 / 引受ガイドライン / 商品規定 / 重要事項説明書。
   引受可否・補償可否・保険金支払可否・契約条件の**最終的な根拠は必ずこれ**。引用には出典マーカー [#N] を付ける。
② 商品部公式Q&A（回答事例）＝ 過去に商品部が出した公式回答。**強力な参考**であり、結論の方向性はこれに合わせる。
   ただし「過去事例」なので、(a) その事例の前提条件（商品・始期・契約形態）が今回の照会と一致するか、
   (b) ①の約款・ガイドラインと矛盾しないか を必ず確認してから使う。
   **商品部Q&Aだけを根拠に断定しない**。引用するときは [#N] を付け、「商品部の回答事例では…」と明示する。
   事例が今回の論点を完全には捉えていない可能性にも留意し、ズレがあれば指摘する。
③ 補足知識 ＝ あなたの一般的な損保実務知識。**保険判断（可否）の根拠にはしない**。説明の補助としてのみ使い、必ず「実務上は」「一般的に」と明示し、出典マーカーは付けない。

【回答の作り方（この順で考える）】
1. **照会の真意を読む**：営業が本当に知りたいこと（背景・前提）を一言で押さえる。曖昧なら前提を明示して場合分けする。
2. **商品部公式Q&Aを探す**：同じ商品・論点の回答事例があれば、それを軸に置く。
3. **約款・ガイドラインで裏付ける**：その結論を①の条文・規定で裏付ける。該当条文を引用し [#N] を付ける。
4. **整合と差分を確認**：商品部Q&Aと約款が一致していればその旨を、前提が違う/事例が見当たらない場合はその限界を正直に書く。
5. **結論を言い切る**：「原則対象」「原則対象外」「全件引受申請が必要」「条件次第（主たる傾向は◯◯）」など、具体的な見解を先頭に出す。

【最重要：抽象的な逃げを禁止する】
「証券・契約内容によります」だけで終わるのは**禁止**です。質問者が知りたいのは「**実務上どちらに転びやすいか、その理由**」です。資料の範囲で必ず具体的な見解を示す。
ただし、約款・ガイドライン・商品部Q&Aのいずれにも該当する記載が無い場合は、無理に断定せず
「社内ナレッジ上、明確な規定・回答事例は確認できませんでした。アンダーライターへご確認ください」と正直に返す。

【事実と推論を厳格に分ける（商品部の要望）】
- 約款・ガイドライン・商品部Q&Aに**明記された内容** → 出典マーカー [#N] を付けて断定してよい。
- あなたの専門知識による**推論・一般的な実務慣行・具体例** → 「実務上は」「一般的に」「〜と考えられます」と明示し、出典マーカーは付けない。
- この2つを**同じ文・同じ箇条書きに混ぜない**。読み手が線引きできるよう、【根拠】と【解釈・当てはめ】の見出しで物理的に分離する。
- 可否の結論（補償される/されない、引受可/不可）は、**必ず①第一次根拠か②商品部Q&Aの [#N] に紐づける**。推論だけを根拠に可否を断定しない。

【出典は「これでもか」と網羅する】
- 結論を支える条文・規定・回答事例は、面倒がらず**該当するものすべて**に [#N] を付ける。
- 引用は「どの資料の何条・何の事例か」が分かる粒度で示す（例: 「所定特約第4条（保険金を支払わない場合）⑤ [#2]」）。
- 1つの結論に複数の根拠があるなら [#2][#5] のように併記する。

【スタイル（読みやすさは妥協しない）】
- 見出しに絵文字（🎯📌💡⭕❌📖など）を使い、視覚的に整理。
- 重要な結論は **太字**。「結論から言うと」など自然な語り口。
- 項目ごとの可否は表が有効（| 項目 | 可否 | 根拠[#N] | の3列など）。
- 質問のかたちに応じて構成は柔軟に。テンプレを機械的に当てはめない。

【守るべき一線（ハルシネーション防止）】
- 資料に**無い**具体的な数字・割合・期間・金額・条文番号を、資料の事実であるかのように書かない。
- ただし定性的な実務推論・具体例は、②③の区別を明示した上で積極的に活用してよい（捏造ではなく専門的解釈）。
- 個人情報（契約番号・氏名・電話番号等）が質問に含まれる場合は「個人情報を含めないでください」とだけ返す。

【出力フォーマット（この骨格を守る。中身は質問に応じて柔軟に）】
```
## 結論 🎯
（照会への答えを先に言い切る。項目が複数なら箇条書き/表で各々の可否）

## 📖 根拠（資料に書いてある事実）
- 約款／引受ガイドライン：第○条「…」… [#N]
- 商品部公式Q&A（回答事例）：「（照会）…」→「（回答）…」 [#N]
  （※今回の照会との前提の一致／差分があればここで触れる）

## 💡 解釈・当てはめ（ここからは実務上の推論）
（根拠を今回の事案に当てはめてどう考えるか。「実務上は」「一般的に」と明示。出典マーカーは付けない）

## 📌 留意点・確認事項
（前提の確認、アンダーライターに上げるべき点、事例と異なる条件があれば明記）
```
（注：根拠が商品部Q&Aしか無い／約款しか無い場合は、無い側を「該当条文は今回の抜粋には含まれていません」等と正直に書く）'''

VERIFIER_PROMPT = '''あなたはシニアアンダーライターであり、商品部に提出する回答の事実確認担当です。ドラフト回答を「軽く」レビューします。完璧主義で削りすぎないこと。NotebookLM のような具体性・推論の豊かさを保ったまま、明確な誤りと「事実と推論の混線」だけを正します。

【この回答に期待される性質（＝守るべき長所、削らない）】
- 約款・引受ガイドライン・商品部公式Q&Aを土台にしつつ、実務知識で**具体的な見解**（可否）を示している。
- 「資料に書いてある事実（[#N]付き）」と「実務上の推論・一般的解釈（出典なし）」が見出しレベルで分離されている。
- 商品部公式Q&Aの回答事例を軸にしつつ、約款・ガイドラインで裏付けている。

【修正する（これだけ）】
1. 資料の**事実として** [#N] 付きで提示された数字・割合・期間・条文番号・可否が、引用元ナレッジに実際には無い → 削除、または推論欄へ移して「実務上は」と明示。
2. 条件分岐の取り違え（「ただし」「〜を除く」「〜に限り」を逆に解釈）。
3. 否定／限定の逆転（「補償する」と「補償しない」「引受可」と「引受不可」の取り違え）。
4. 出典マーカー [#N] が指す chunk と本文の主張が一致しない → マーカーを正す。
5. **可否の結論が推論だけで断定されている**（第一次根拠＝約款/ガイドライン、または商品部Q&Aの [#N] に紐づいていない）
   → 紐づく根拠があればマーカーを補う。無ければ結論を「実務上は〜と考えられます（要確認）」と推論トーンに直す。
6. **商品部Q&Aだけを根拠に可否を断定し、約款・ガイドラインの裏付けが無い** → その旨を末尾の確認事項に「※約款・ガイドライン上の裏付けは今回の抜粋では未確認」と1行補う（結論自体は消さない）。
7. 明らかな事実誤認（保険種類・商品の取り違え 等）。

【絶対に削除・改変しない】
- 「実務上は」「一般的に」「〜と考えられます」と明示された推論・具体例。出典マーカーが無くても誤りではない。**正しい付加価値**として残す。
- 具体的な結論・見解（「原則対象外と考えられます」「全件引受申請が必要です」等）。「契約内容次第」へ安易に後退させない。
- ⭕❌🎯💡📖📌 等の絵文字、太字、表、見出し構成（結論／根拠／解釈・当てはめ／留意点）、自然な語り口。

【判断基準】
- 「実務上は」と明示された推論 → **残す**（出典マーカー不要）。
- 資料の事実として [#N] で断定されているのに引用元に無い → 直す（推論欄へ移すか削る）。
- 可否の結論 → 必ず約款/ガイドライン or 商品部Q&A の [#N] に紐づいているか確認。
要は「専門家としての解釈」は活かし、「資料の捏造」と「事実/推論の混線」だけを潰す。

【出力】
- 検証後の最終回答**のみ**を出力（プロセス説明・前置きを書かない）。
- ドラフトの見出し・絵文字・表・スタイル・4セクション構成を維持。
- 修正した場合のみ末尾に `※検証で修正: <要点>` を1行追記。修正なしなら何も追記しない。
- 約款・ガイドライン・商品部Q&Aのいずれにも該当が全く無い場合のみ「社内ナレッジに該当する記載・回答事例がありません。アンダーライターへ確認してください」と返す。

【最後に必ず付ける機械処理用の2行（本文の一番最後。ユーザーには表示されない）】
回答本文の後に、改行して次の2行を**必ず**この書式で出力する（他の場所には書かない）:
[[CONFIDENCE:高]]   ← 高/中/低 のいずれか。判定基準:
   高 = 補償可否/引受可否/支払可否が、約款または引受ガイドラインの明確な根拠[#N]で断定できる
   中 = 商品部Q&A事例や実務推論が主たる根拠で、約款の直接の条文根拠が弱い／前提次第で変わる
   低 = 該当する規定・回答事例が乏しく、規定を特定できない／要確認事項が多い
[[FOLLOWUP: 質問1 | 質問2 | 質問3]]   ← この回答の次に営業が深掘りすべき質問を最大3つ。無ければ「[[FOLLOWUP: なし]]」'''

INTENT_PROMPT = '''あなたは損害保険の照会対応を支援する「意図解釈アシスタント」です。
営業担当者の照会文から、本当に聞きたいこと(intent)と、回答に影響する前提(assumptions)を読み取り、
必要なら確認質問(clarifying_questions)を返します。さらに、社内ナレッジ検索に使うサブクエリ(search_queries)へ分解します。
あなたは回答そのものは作りません。「何を聞かれているのか」を正確に捉えることだけに集中してください。

【モード】現在のモード: {mode}
- smart : 照会が十分明確なら確認質問は出さず needs_clarification=false で即座に回答へ進ませる。
  回答が分岐するほど重要な前提が欠けている場合のみ needs_clarification=true。
- always: 必ず前提を1つ以上確認する(needs_clarification=true)。ただし的外れな確認はせず、回答が変わる前提に絞る。

【確認質問の作り方（最重要）】
- 「答えが実際に変わる」前提だけを聞く（契約形態=派遣/請負、被保険者の範囲、事故の発生主体、対象商品名、補償か引受かの別 等）。
- 営業が答えやすいよう、選択肢形式や具体例を添える（例:「契約形態は『労働者派遣』ですか『業務請負』ですか？」）。
- 最大3問。ゼロでもよい。挨拶・雑談・保険と無関係な照会には確認質問を出さない。

【サブクエリ(search_queries)】
- 照会を、約款・引受ガイドラインを引くための2〜4個の具体的な論点に分解する。
- 例:「派遣SEが派遣先でサイバー事故、派遣会社の賠償は？」
  → ["労働者派遣契約における被保険者の範囲","派遣先で生じた損害に対する派遣元の賠償責任","サイバー保険の賠償責任補償の対象範囲","他人の業務遂行中の事故の免責有無"]

【会話履歴（あれば踏まえる。続きの照会なら前提を引き継ぐ）】
{history}

【照会文】
{question}

【出力（JSONのみ。前後に文章を書かない。コードフェンスも付けない）】
{
  "intent": "営業が本当に知りたいことを1文で",
  "assumptions": ["回答にあたり置いた前提1","前提2"],
  "needs_clarification": true,
  "clarifying_questions": ["確認質問1","確認質問2"],
  "search_queries": ["サブクエリ1","サブクエリ2","サブクエリ3"]
}'''

SYSTEM_PROMPTS = [
    ("intent",   INTENT_PROMPT,   "Step0: 照会の意図解釈・前提確認・サブクエリ分解"),
    ("router",   ROUTER_PROMPT,   "Step1: 関連チャンクを選ぶ"),
    ("drafter",  DRAFTER_PROMPT,  "Step3: 構造化された回答を生成"),
    ("verifier", VERIFIER_PROMPT, "Step4: 各主張をナレッジと照合し誤りを削除/修正"),
]


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------
def _make_main_sheet(ws, chunks=None):
    """Repurpose skeleton's existing Sheet1 as the main sheet."""
    ws.title = "main"
    chunks = chunks or []
    total = len(chunks)
    qa_n = sum(1 for c in chunks if c.get("doc_type") == "商品部公式Q&A")
    src_n = len({c.get("source") for c in chunks if c.get("doc_type") != "商品部公式Q&A"})
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
        "・このファイルを開くと、自動でチャットUIが表示されます。",
        "・チャット画面が表示されない場合は、マクロが有効になっていません。",
        "  下記の手順でVBAを有効化してから、Excelを再起動してください。",
        "",
        "  [ファイル]→[オプション]→[トラストセンター]→[トラストセンターの設定]→[マクロの設定]",
        "  → 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェック → OK",
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


def _make_howto(wb, chunks=None):
    """Quick-start guide. Lives at sheet index 1 (after main) so users see
    it first when they tab over from the chat sheet."""
    chunks = chunks or []
    total = len(chunks)
    qa_n = sum(1 for c in chunks if c.get("doc_type") == "商品部公式Q&A")
    src_n = len({c.get("source") for c in chunks if c.get("doc_type") != "商品部公式Q&A"})
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
        ("基本の流れ（質問→回答→自己理解→評価で1セット）",
         "1. main タブの質問欄に知りたいことを入力\n"
         "2. 「送信」→ AI がまず照会の意図・前提を確認（曖昧なら質問が返ります）\n"
         "3. 60〜90秒で答えが表示される（出典・引用箇所も同時に表示）\n"
         "4. ★まず出典元（約款・ガイドライン）を自分で確認する\n"
         "5. ★理解できたか確認する。曖昧なら「続けて質問」で深掘り\n"
         "6. 「○良かった」「×修正」で評価（ここまでで1セット）\n"
         "7. それでも判断に迷う時だけ「教えてBOX」で本社へ照会"),

        ("⚡ 必ず最新モデル(GPT-5.5)に切り替えて使う",
         "社内 AI リボンのモデル選択を、2026年6月時点の最新『GPT-5.5』に切り替えてから使ってください。\n"
         "古いモデルのままだと、推論の精度・出典の正確さが大きく落ちます。\n"
         "リボン上のモデル選択 → GPT-5.5 を選ぶ → このボットを使う、の順で。"),

        ("🧭 確認モード（賢く / 毎回）の切り替え",
         "main 画面の『確認モード』ボタンで、AI の前提確認の強さを切り替えられます。\n"
         "  ・賢く出し分け（既定）: 質問が明確なら即回答、曖昧な時だけ前提を聞き返す\n"
         "  ・毎回前提確認        : どんな質問でもまず前提を確認してから回答（誤解を最小化）\n"
         "「認識をきっちり合わせたい」「新人で不安」な時は『毎回前提確認』がおすすめ。"),

        ("📮 教えてBOX（本社への照会）",
         "AI の自信度が『中／低』の時、または自分で判断に迷う時だけ使ってください。\n"
         "ボタンを押すと、あなたの質問・AIが調べた範囲・該当しそうな条文を整理した\n"
         "『照会パケット』を自動生成します（Outlook があればメール下書きが開きます）。\n"
         "本社は一から聞き返す手間が減り、あなたも整理された状態で照会できます。\n"
         "※安易な照会は避け、まず出典確認・続けて質問で自己解決を試みましょう。"),

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

        ("✅ 使えるタブはこの2つだけ",
         "  ・使い方     : このシート（参照のみ）\n"
         "  ・main      : 質問と回答（普段の作業場所）\n\n"
         "それ以外のタブは、システムが自動管理しており、通常は表示されません。\n"
         "部署の変更が必要な場合は、管理者に連絡してください。"),

        ("⚠️ 個人情報・機密情報を入れないでください",
         "契約番号、氏名、電話番号、マイナンバー、E メールアドレスを質問に書かないでください。\n"
         "個人情報を検知すると警告が出ます。警告が出たら必ず書き換えてから送信してください。\n"
         "ボットは社内 AI リボン経由でクラウドの LLM を呼びます。送信した本文は社外には残りませんが、\n"
         "**社内ログには記録される** 前提で書いてください。"),

        ("⚠️ AI の答えは最終決定ではありません",
         "回答は社内ナレッジ（約款・引受ガイドライン・商品部公式Q&A）を根拠に AI が生成しています。\n"
         "商品部公式Q&Aは過去の回答事例です。今回の照会と前提条件が同じか必ず確認してください。\n"
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

    # Copyright footer
    row += 1
    cr = ws.cell(row=row, column=1,
                 value="© ニューリスクG 小島正豪 — 本チャットボットの設計・思想・プロンプトは開発者に帰属します。無断複製・転用を禁じます。")
    cr.font = Font(size=9, italic=True, color="888888")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)

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
    ws.sheet_state = "hidden"


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
    ws.sheet_state = "veryHidden"


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
    ws.sheet_state = "veryHidden"


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
    ws.sheet_state = "veryHidden"
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
    ws.sheet_state = "veryHidden"


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
    ws.sheet_state = "hidden"


def _make_vba_src(wb):
    """Hidden sheet carrying every standard module's source code."""
    ws = wb.create_sheet("vba_src")
    for c, h in enumerate(["module_name", "type", "source"], 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)

    mods = ["modBoot", "modConfig", "modUserProfile",
            "modRibbonGateway", "modKnowledgeBase",
            "modPrompts", "modIntent", "modFeedbackLookup", "modPipeline",
            "modFeedback", "modInquiryBox", "modChatUI", "modPii",
            "modUsageLogger", "modDiag"]

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
    ws.sheet_state = "veryHidden"
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
    _make_main_sheet(wb[wb.sheetnames[0]], chunks)  # repurpose Sheet1 → main
    _make_howto(wb, chunks)
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

    # Stage 6: VBA project password protection (Windows + Excel COM only)
    print("Stage 6: VBA project password protection...")
    try:
        import win32com.client
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = False
        try:
            wb = excel.Workbooks.Open(os.path.abspath(OUT_XLSM))
            wb.VBProject.Protection = 1
            wb.VBProject.ProtectionsPassword = "AD1459"
            wb.Save()
            wb.Close()
            print("  VBA project password protection applied (AD1459)")
        finally:
            excel.Quit()
    except ImportError:
        print("  (Skipped: requires Windows + pywin32. Install via: pip install pywin32)")
    except Exception as e:
        print(f"  Warning: could not protect VBA project: {e}")

    print("\nDone.")


if __name__ == '__main__':
    main()
