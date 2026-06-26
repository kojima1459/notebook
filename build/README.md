# build/

This directory contains the build pipeline for the two distribution xlsm files.

## Required artifacts

The chat UI is rendered onto a worksheet at runtime by `modChatUI`, so no
UserForm and no template workbooks are required. `build.ps1` creates
blank macro-enabled workbooks from scratch via Excel COM, then injects
the modules.

VBA-JSON v2.3.1 (`vendor/JsonConverter.bas`, MIT licensed) is bundled in
the repo, so no manual download is needed. Source:
<https://github.com/VBA-tools/VBA-JSON/releases/tag/v2.3.1>

## Build

On a Windows machine with Excel installed:
```
cd build
powershell -ExecutionPolicy Bypass -File build.ps1
```

Output appears in `../dist/`.

## Chatbot v2 知識ベースの再ビルド (Python)

v2 (`dist/Chatbot_v2.xlsm`) は Python で組み立てる。知識データは
`dist/index/chunks_enriched.json` に集約され、各取り込みスクリプトが追記する。

商品部公式Q&A (`source_data/shohinbu_qa.xlsx`) を更新したら:
```
python build/add_shohinbu_qa_chunks.py   # Q&A を chunks_enriched.json に取り込み (冪等)
python build/build_chatbot_v2.py          # dist/Chatbot_v2.xlsm を再生成
```
`add_shohinbu_qa_chunks.py` は同一 source の旧チャンクを除去してから入れ直すので、
何度実行しても重複しない。system_prompt（ルーター/ドラフター/検証）は
`build_chatbot_v2.py` 内に定義され、ビルドのたびにシートへ書き込まれる。

## VBA Project Password Protection

After the build, the VBA project is optionally locked with password `AD1459`
(Windows environment only):

```
pip install pywin32
python build/build_chatbot_v2.py   # applies password during final stage
```

Without this, users can unhide internal sheets and read/modify VBA code via
Tools > [Project Name] Properties > Protection > Edit Existing Password.

**Effect of password protection:**
- VBA Editor メニューが表示されなくなる （Tools, View など）
- シートのunhide機能は効果なし（VBA内で `xlSheetVeryHidden` を使っているため）
- パスワードを知らないユーザーは、VBAコードの編集・読込ができない

If the build environment is Linux / non-Windows, the password protection
step is skipped safely (no error) — the file still works, but VBA is
unprotected.

## Trust Center prerequisite

`build.ps1` uses the `VBProject` object to inject modules. This is gated by
**File > Options > Trust Center > Trust Center Settings > Macro Settings >
Trust access to the VBA project object model**. Enable it on the build
machine only (not user PCs).
