#!/usr/bin/env python3
"""
build_chatbot_v2.py - Assemble Chatbot_v2.xlsm:
  1. Use openpyxl to construct base xlsm with all data sheets pre-populated:
       main, config, system_prompt, manifest, knowledge_base,
       feedback, department, usage_log
  2. Then inject the v2 VBA modules using the existing make_xlsm.py
     CFB/OVBA machinery (template_skeleton.xlsm provides the verified
     vbaProject.bin scaffold).

This produces dist/Chatbot_v2.xlsm — single file, no external dependencies.
The corporate user just opens it via Excel and the macros call:
    Application.Run("ChatGPT", prompt)
through the 社内 AI ribbon.
"""

import io
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile

_ILLEGAL_CHARS = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')


def _clean(s: str) -> str:
    return _ILLEGAL_CHARS.sub("", s) if isinstance(s, str) else s

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Reuse helpers from make_xlsm.py
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import make_xlsm

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(SCRIPT_DIR)
SRC_V2 = os.path.join(ROOT, "src", "chatbot_v2")
TEMPLATE = os.path.join(SCRIPT_DIR, "template_skeleton.xlsm")
ENRICHED_JSON = os.path.join(ROOT, "dist", "index", "chunks_enriched.json")
OUT_XLSM = os.path.join(ROOT, "dist", "Chatbot_v2.xlsm")

# ---------------------------------------------------------------------------
# Data we'll bake into the workbook
# ---------------------------------------------------------------------------

# Config defaults — admins edit these in-Excel
CONFIG_ROWS = [
    # key, value, description
    ("router_max_chunks",   8,    "ルーターが選ぶ関連チャンク数の上限"),
    ("verifier_enabled",    True, "自己検証パスを実行するか (True=精度優先 / False=コスト優先)"),
    ("mock_llm",            False, "True=社内AIリボンを呼ばずダミー応答（動作確認用 / Mac可）。本番はFalse"),
    ("max_context_chars",   60000, "ドラフトに渡すコンテキストの最大文字数"),
    ("debug_mode",          False, "True にすると全LLMプロンプトをdebug_logシートに残す"),
    ("feedback_top_n",      3,    "ドラフトに参照させる過去Q&Aの最大件数"),
    ("feedback_min_score",  0.18, "類似Q&A採用の閾値 (0..1)"),
    ("router_model",        "gpt-5.5", "(備考) ルーター用モデル — リボン設定と整合させる"),
    ("drafter_model",       "gpt-5.5", "(備考) ドラフト用モデル"),
    ("verifier_model",      "gpt-5.5", "(備考) 検証用モデル"),
    ("draft_temperature",   0.15, "(備考) ドラフト温度 — リボン設定があれば併記"),
    ("department_id",       "",   "初回起動時に設定される"),
    ("department_name",     "",   ""),
    ("role",                "",   ""),
]

# Department master — start with one dept for MVP. Easy to add rows later.
DEPARTMENTS = [
    ("expense_profit", "費用利益保険チーム", "common,expense_profit"),
]

# ---------- System prompts ----------------------------------------------------

ROUTER_SYSTEM_PROMPT = """あなたは社内ナレッジ検索のルーターです。
ユーザーの質問に対し、提供されたナレッジ一覧から最も関連性の高いチャンクのIDを最大{max_n}件、JSON形式で返してください。

【選び方】
- 質問の主題に直接関わるチャンクを優先する
- 約款の条文を尋ねる質問なら『普通保険約款』タイプを優先
- 引受可否の質問なら『引受ガイドライン』を優先
- 用語の意味を尋ねる質問なら『研修資料』『解説』『FAQ』も含めてよい
- 同じ業務領域の中でも、条文間で関連が深いものは複数選んでよい

【ユーザー質問】
{question}

【ナレッジ一覧】
形式: chunk_id | 業務領域 | 書類種別 | 出典 | 1行要約 | キーワード

{knowledge_table}

【出力 (JSON、他のテキストは絶対に書かない)】
{
  "selected_ids": ["sample_03.pdf::p5::c2", "..."],
  "reasoning": "なぜこれらを選んだか1〜2文"
}"""

DRAFTER_SYSTEM_PROMPT = """あなたは社内アンダーライターの業務支援AIです。提示された社内ナレッジ抜粋および認定済みQ&Aのみを根拠に回答してください。

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
・ 過去の認定済みQ&Aがある場合は、それを最優先で参照しつつ、新しい質問に合わせて再構成する。"""

VERIFIER_SYSTEM_PROMPT = """あなたは損害保険会社引受部門の品質管理者です。AIが書いたドラフト回答が、引用元ナレッジに本当に書かれている内容のみで構成されているか厳格に検証します。

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
・ ドラフトが全面的にナレッジに無い内容だった場合は『社内ナレッジに該当する記載がありません。アンダーライターへ確認してください』のみ返す。"""

SYSTEM_PROMPTS = [
    ("router",   ROUTER_SYSTEM_PROMPT,   "Step1: 関連チャンクを選ぶ"),
    ("drafter",  DRAFTER_SYSTEM_PROMPT,  "Step3: 構造化された回答を生成"),
    ("verifier", VERIFIER_SYSTEM_PROMPT, "Step4: 各主張をナレッジと照合し誤りを削除/修正"),
]


# ---------------------------------------------------------------------------
# Build the openpyxl workbook with all data sheets
# ---------------------------------------------------------------------------

THIN_BORDER = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)


def write_data_workbook(out_path: str):
    """Create the base xlsm with all data sheets pre-populated.

    We start with openpyxl's empty workbook (keep_vba unsupported on a fresh
    workbook). The output xlsm is technically an xlsx until vbaProject.bin
    is injected later — that's fine.
    """
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    _make_main_sheet(wb)
    _make_config_sheet(wb)
    _make_system_prompt_sheet(wb)
    _make_department_sheet(wb)
    _make_manifest_sheet(wb)
    _make_knowledge_base_sheet(wb)
    _make_feedback_sheet(wb)
    _make_usage_log_sheet(wb)
    _make_vba_src_sheet(wb)

    wb.save(out_path)


def _bold(ws, cell, value):
    ws[cell] = value
    ws[cell].font = Font(bold=True)


def _make_main_sheet(wb):
    ws = wb.create_sheet("main")
    # Macros normally rebuild this sheet on open. If they don't fire, this
    # static content tells the user exactly what to do.
    ws["A1"] = "社内ナレッジ QA ボット (v2)"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A1"].fill = PatternFill("solid", fgColor="3C5AA0")
    ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    ws["A3"] = "▼ もしこの画面のままで操作ボタンが出ない場合 ▼"
    ws["A3"].font = Font(bold=True, size=12, color="CC0000")

    instructions = [
        "",
        "① 上の黄色帯「セキュリティの警告」が出ていたら『コンテンツの有効化』を押す",
        "",
        "② それでもボタンが出ない場合は次の手順:",
        "   1. キーボードで Alt + F11 を押す (VBAエディタが開く)",
        "   2. 左の『プロジェクト』ツリーから『ThisWorkbook』をダブルクリック",
        "   3. 開いたコードの中の『Public Sub Setup()』の行をクリック",
        "   4. F5 キーを押す",
        "   5. 『セットアップ完了』のメッセージが出れば成功",
        "",
        "③ それでも動かない場合は管理者に画面写真を送ってください",
        "",
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━",
        "■ このシステムの仕組み",
        "  ・PDFマニュアル15冊(813チャンク)から関連箇所を検索",
        "  ・社内AIリボン(リボンちゃん/MSAD-Addin)で回答生成",
        "  ・○良かった / ×修正 で学習(同じ質問が来たら蓄積を参照)",
        "",
        "■ 重要",
        "  ・回答は必ずアンダーライターの最終確認を取ること",
        "  ・契約番号・氏名・電話番号は質問に含めないこと",
    ]
    for i, line in enumerate(instructions, start=4):
        cell = ws.cell(row=i, column=1, value=line)
        if line.startswith("■"):
            cell.font = Font(bold=True, size=11)
        elif line.startswith("━"):
            cell.font = Font(color="888888")
        else:
            cell.font = Font(size=11)

    ws.column_dimensions["A"].width = 80


def _make_vba_src_sheet(wb):
    """Hidden sheet carrying every standard module's source code as fallback.

    If Excel drops the binary's standard modules (some versions do this with
    programmatically built vbaProject.bin), ThisWorkbook.Setup reads this
    sheet and recreates the modules via the VBProject API.
    """
    import re
    ws = wb.create_sheet("vba_src")
    headers = ["module_name", "type", "source"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)

    mods = ["modBoot", "modConfig", "modUserProfile",
            "modRibbonGateway", "modKnowledgeBase",
            "modPrompts", "modFeedbackLookup", "modPipeline",
            "modFeedback", "modChatUI", "modPii", "modUsageLogger",
            "modDiag"]

    EXCEL_CELL_LIMIT = 32000  # actual limit 32767; leave a margin

    for i, name in enumerate(mods, start=2):
        bas_path = os.path.join(SRC_V2, name + ".bas")
        with open(bas_path, encoding="utf-8") as fp:
            txt = fp.read()
        # Strip Attribute lines and the BOM; AddFromString does not need them.
        out_lines = []
        for line in txt.split("\n"):
            stripped = line.lstrip("﻿")
            if stripped.lstrip().startswith("Attribute "):
                continue
            out_lines.append(stripped)
        cleaned = "\n".join(out_lines)
        if len(cleaned) >= EXCEL_CELL_LIMIT:
            raise RuntimeError(
                f"{name}.bas ({len(cleaned)} chars) exceeds the {EXCEL_CELL_LIMIT}-char "
                f"cell limit; needs split across multiple cells.")
        ws.cell(row=i, column=1, value=name)
        ws.cell(row=i, column=2, value="std")
        ws.cell(row=i, column=3, value=cleaned)

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 80
    ws.sheet_state = "hidden"
    print(f"  vba_src: embedded {len(mods)} module sources")


def _make_config_sheet(wb):
    ws = wb.create_sheet("config")
    headers = ["key", "value", "description"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
    for i, (k, v, d) in enumerate(CONFIG_ROWS, start=2):
        ws.cell(row=i, column=1, value=k)
        ws.cell(row=i, column=2, value=v)
        ws.cell(row=i, column=3, value=d)
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 70


def _make_system_prompt_sheet(wb):
    ws = wb.create_sheet("system_prompt")
    headers = ["key", "prompt", "description"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
    for i, (k, p, d) in enumerate(SYSTEM_PROMPTS, start=2):
        ws.cell(row=i, column=1, value=k)
        ws.cell(row=i, column=2, value=p).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=i, column=3, value=d)
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 80
    ws.column_dimensions["C"].width = 40
    for i in range(2, 2 + len(SYSTEM_PROMPTS)):
        ws.row_dimensions[i].height = 240


def _make_department_sheet(wb):
    ws = wb.create_sheet("department")
    headers = ["dept_id", "dept_name", "knowledge_scope_csv"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
    for i, row in enumerate(DEPARTMENTS, start=2):
        for c, v in enumerate(row, start=1):
            ws.cell(row=i, column=c, value=v)
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 40


def _make_manifest_sheet(wb):
    ws = wb.create_sheet("manifest")
    headers = ["source", "display", "domain", "doc_type",
               "dept_scope", "chunk_count", "notes"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)

    # Build manifest from enriched chunks
    chunks = json.load(open(ENRICHED_JSON, encoding="utf-8"))
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
        ws.cell(row=row, column=7, value="")
        row += 1
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 48
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 28
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 30


def _make_knowledge_base_sheet(wb):
    ws = wb.create_sheet("knowledge_base")
    headers = ["chunk_id", "source", "display", "domain", "doc_type",
               "section_header", "summary", "keywords", "full_text", "dept_scope"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)

    chunks = json.load(open(ENRICHED_JSON, encoding="utf-8"))
    for i, c in enumerate(chunks, start=2):
        ws.cell(row=i, column=1, value=c["id"])
        ws.cell(row=i, column=2, value=c["source"])
        ws.cell(row=i, column=3, value=c.get("display", c["source"]))
        ws.cell(row=i, column=4, value=c.get("domain", "未分類"))
        ws.cell(row=i, column=5, value=c.get("doc_type", "未分類"))
        ws.cell(row=i, column=6, value=c.get("header", ""))
        ws.cell(row=i, column=7, value=c.get("summary", ""))
        ws.cell(row=i, column=8, value=c.get("keywords", ""))
        # Excel cell limit is 32767 chars; chunks are well under
        ws.cell(row=i, column=9, value=_clean(c.get("text", ""))[:32000])
        ws.cell(row=i, column=10, value="common")

    # Column widths
    widths = [28, 18, 40, 22, 18, 50, 60, 40, 80, 16]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    print(f"  Wrote {len(chunks)} knowledge_base rows")


def _make_feedback_sheet(wb):
    ws = wb.create_sheet("feedback")
    headers = ["fb_id", "timestamp", "dept_id", "submitter",
               "question", "answer", "status", "approver",
               "approved_at", "correction", "tags", "source_chunk_ids"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
    widths = [22, 20, 18, 14, 50, 80, 12, 14, 20, 80, 24, 60]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _make_usage_log_sheet(wb):
    ws = wb.create_sheet("usage_log")
    headers = ["timestamp", "dept_id", "user",
               "q_chars", "a_chars", "router_ms", "draft_ms",
               "verify_ms", "total_ms", "chunks_used", "feedback_status"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
    widths = [20, 18, 14, 10, 10, 10, 10, 10, 10, 60, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ---------------------------------------------------------------------------
# Load v2 VBA modules
# ---------------------------------------------------------------------------

def load_v2_modules() -> list:
    """Build module list for v2 chatbot."""
    tw_src = make_xlsm.load_module_source(
        os.path.join(SRC_V2, "ThisWorkbook.cls"), is_cls=True)

    modules = [
        make_xlsm._doc("ThisWorkbook", tw_src),
    ]
    # One Document object per worksheet — match the sheet order in xlsm
    # (openpyxl wrote 8 sheets). VBA does not need full code in these,
    # just placeholder so the project compiles.
    # 9 sheets now: main, config, system_prompt, department, manifest,
    # knowledge_base, feedback, usage_log, vba_src
    for sheet_codename in ["Sheet1", "Sheet2", "Sheet3", "Sheet4",
                            "Sheet5", "Sheet6", "Sheet7", "Sheet8", "Sheet9"]:
        modules.append(make_xlsm._doc(sheet_codename, make_xlsm.EMPTY_DOC_SOURCE))

    for n in ["modBoot", "modConfig", "modUserProfile",
              "modRibbonGateway", "modKnowledgeBase",
              "modPrompts", "modFeedbackLookup", "modPipeline",
              "modFeedback", "modChatUI", "modPii", "modUsageLogger",
              "modDiag"]:
        modules.append(make_xlsm._std(n, make_xlsm.load_module_source(
            os.path.join(SRC_V2, n + ".bas"))))

    return modules


# ---------------------------------------------------------------------------
# Re-zip xlsm with vbaProject.bin + patched [Content_Types].xml + rels
# ---------------------------------------------------------------------------

def inject_vba_and_patch_xml(in_xlsx: str, out_xlsm: str, modules: list):
    """Take an openpyxl-built xlsx, inject vbaProject.bin, patch XML for macros
       and assign codeName attributes per sheet so the modules bind."""
    with open(TEMPLATE, "rb") as f:
        skel_bytes = f.read()
    with zipfile.ZipFile(io.BytesIO(skel_bytes)) as z:
        skel_vba = z.read("xl/vbaProject.bin")

    vba_bin = make_xlsm.build_vba_project(modules, skel_vba)
    print(f"  vbaProject.bin: {len(vba_bin):,} bytes")

    # Identify our sheets to know how many codeNames to assign
    with zipfile.ZipFile(in_xlsx, "r") as z:
        sheet_files = sorted([
            n for n in z.namelist()
            if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")
        ], key=lambda s: int(s.replace("xl/worksheets/sheet", "").replace(".xml", "")))
        print(f"  Found {len(sheet_files)} sheets: {[s.split('/')[-1] for s in sheet_files]}")

    with zipfile.ZipFile(in_xlsx, "r") as zin, \
         zipfile.ZipFile(out_xlsm, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        names = zin.namelist()
        for name in names:
            data = zin.read(name)
            if name == "[Content_Types].xml":
                data = make_xlsm.patch_content_types(data)
            elif name == "xl/_rels/workbook.xml.rels":
                data = make_xlsm.patch_workbook_rels(data)
            elif name == "xl/workbook.xml":
                data = make_xlsm.patch_workbook_codename(data, "ThisWorkbook")
            elif name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
                # Assign Sheet1, Sheet2... codeNames in order
                idx = int(name.replace("xl/worksheets/sheet", "").replace(".xml", ""))
                data = make_xlsm.patch_sheet_codename(data, f"Sheet{idx}")
            zout.writestr(name, data)
        # Inject vbaProject.bin
        zout.writestr("xl/vbaProject.bin", vba_bin)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if not os.path.exists(ENRICHED_JSON):
        sys.exit(f"Missing {ENRICHED_JSON} — run build/enrich_chunks.py first")

    print("=== build_chatbot_v2.py ===")
    print(f"Source v2 dir : {SRC_V2}")
    print(f"Enriched data : {ENRICHED_JSON}")
    print(f"Output xlsm   : {OUT_XLSM}")
    print()

    os.makedirs(os.path.dirname(OUT_XLSM), exist_ok=True)

    # Stage 1: write data workbook (xlsx)
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name
    print("Stage 1: Building data workbook with openpyxl...")
    write_data_workbook(tmp_path)
    print(f"  Wrote {tmp_path} ({os.path.getsize(tmp_path):,} bytes)")

    # Stage 2: load VBA modules
    print("\nStage 2: Loading v2 VBA modules...")
    modules = load_v2_modules()
    print(f"  Modules ({len(modules)}): {[m['name'] for m in modules]}")

    # Stage 3: inject VBA + patch XML
    print("\nStage 3: Injecting vbaProject.bin + patching XML...")
    inject_vba_and_patch_xml(tmp_path, OUT_XLSM, modules)
    print(f"  Final: {OUT_XLSM} ({os.path.getsize(OUT_XLSM):,} bytes)")

    os.unlink(tmp_path)
    print("\nDone.")


if __name__ == "__main__":
    main()
