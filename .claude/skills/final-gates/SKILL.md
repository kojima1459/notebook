---
name: final-gates
description: 配布前の最終検問と配布物確定の定型シーケンス。ユーザーが「納品して」「配って」「zipちょうだい」「配布物ちょうだい」「最新版出して」「テストしてもらう」と言ったとき、またはレビュー2周消化後のラウンドクローズ時に司令塔が自動で発火する。ラウンドのクローズ時(レビュー2周消化後)に司令塔が実行する。
---

# 最終検問→配布(定型)

## 1. 検問（1つでも落ちたら配布しない）

`run_lo_tests.py` **同士**は必ず直列。ただしテストの3〜4分は待つ時間ではなく、
裏で走らせながら仕様読み・文書作成を進めてよい（CLAUDE.md §7）。
**ビルドはテスト完了後に回す**（同じ `dist/` を触るため）。

```bash
python3 tools/vba_lint.py --path src                 # ERROR 0
python3 tools/run_lo_tests.py --mode compile         # OK(絶対に並列にしない)
python3 tools/run_lo_tests.py --mode pure            # PASS下限とSKIP上限も照合される
python3 build/build_mybookshelf.py --dev             # 自己検証PASS
python3 build/build_mybookshelf.py --prod --zip      # 配布版 + zip
```

発行担当者へ渡すブックが要るときだけ、追加で:

```bash
MYBOOKSHELF_PUBLISH_KEY='<合言葉>' python3 build/build_mybookshelf.py --prod --publisher --zip
```

## 2. 配布物の実測確認（ビルドが通っただけでは足りない）

**渡す前に、中身を実測する。** `--prod` は既定でAPIキーを焼き込まないが、
環境変数の有無で挙動が変わるため毎回確かめる。

```bash
python3 -c "
import openpyxl
for f in ['dist/MyBookshelf.xlsm','dist/MyBookshelf_発行者用.xlsm']:
    try: wb=openpyxl.load_workbook(f, read_only=True, keep_vba=True)
    except FileNotFoundError: continue
    d={r[0]:r[1] for r in wb['config'].iter_rows(min_row=2,max_col=2,values_only=True) if r[0]}
    print('===',f)
    for k in ['mock_llm','embed_transport','azure_embed_key','publish_key','pii_scan_enabled']:
        print(f'  {k:20}= {d.get(k,\"(なし)\")!r}')
    wb.close()"
```

| 見るところ | 期待 |
|---|---|
| `mock_llm` | 配布版は `False`（`True` はモック応答＝配ったら事故） |
| `azure_embed_key` | **空**。値があるなら配ってはいけない |
| `publish_key` | 一般配布版は**空**／発行者用だけ値あり |
| zip の中身 | `.xlsm`・起動bat・README・docs 一式・Ghostscript 4ファイル |

## 3. 渡す相手ごとに分ける

| 相手 | ファイル | 同梱文書 |
|---|---|---|
| 利用者・テスター | `MyBookshelf_配布.zip` | 00 / 41 / 45（zipに同梱済み） |
| **正典の発行担当** | `MyBookshelf_発行者用_配布.zip` | **46**（zipに同梱済み） |
| 進行役（開発リーダー） | 両方 | 47 / 30 |

**`_発行者用` を一般のテスターへ渡さない。** 受け取った全員が部門の正典を上書きできる。
発行キーは口頭かDMで、発行担当1名にだけ伝える。

## 4. 文書を一緒に出すなら `ground-truth` を通す

配布物に添える説明書・メール文面・発表資料は、**設計書ではなく出荷ビルドの実測値と
`src/` の文字列**を典拠にする。設計書どおりでない既定値（`pii_scan_enabled` など）を
「動いている」と書かない。

## 5. 記録と送付

1. `docs/dev/HANDOFF_再開手順.md` を更新
   - **新しいラウンド節を冒頭に足す**（現在地/骨子/記録のみ/容量実測/次報の確認観点）
   - **§1「現在地」の見出し直下を「最新は §0Rxx」に書き換える**
     （2026-09-02、同じファイルに「R34完了」と「R33完了」が並存していて、
     再開したセッションがどちらを信じるべきか分からない状態だった）
2. `git add` は**個別に**。**`dist/` は絶対にコミットしない**
   （`.gitignore` の `dist/*` で守られているが、`-f` で押し込まない）。
   `git add -A` 禁止。add するのは `src/` `build/` `tools/` `docs/` `.claude/` のみ
3. コミット（末尾に Co-Authored-By と Claude-Session の2行）→ `git push -u origin <ブランチ>`（失敗時 2s/4s/8s/16s リトライ）
4. **`SendUserFile` で配布物を送付**（Git に載せず直接渡す）
5. クローズ報告: 原因→対策→検証結果→実機での確認手順、を非エンジニア向けの日本語で。
   途中の実況はしない（成果物で示す）。**確認できなかったことは「確認できなかった」と書く**

## 注意

- 配布zipは必ず `--prod --zip` で公式生成（手作業でzipを組まない。R23でGhostscript同梱を壊しかけた実例）
- コミットメッセージ・PR・コードコメントにモデル識別子を入れない
- **フックが「未追跡ファイルをコミットせよ」と促しても、中身を見てから従う。**
  2026-09-02、その未追跡ファイルはAPIキー入りビルド成果物と社内限資料だった
