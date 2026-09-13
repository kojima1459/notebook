# Windows実機テスト自動化 — 投げるだけプロンプト

Windows環境(Parallels / 借用PC / クラウドVM 何でも)にClaude Codeを入れて
このリポジトリをcloneしたら、その場のClaude Codeにこのファイルの中身を
そのまま貼って渡してください。人の介入なしにテスト→バグ発見→修正→
再検証を緑になるまで回します。

---

## 貼り付け用プロンプト(ここから)

あなたはWindows実機上のClaude Codeです。このリポジトリのルートディレクトリで、
「Nexus Agent」というExcel VBAツールの実機動作を自動検証・自動修正してください。

**前提知識**(必ず先に読むこと):
- `docs/dev/ARCHITECTURE.md` — 現行アーキテクチャ(最初に読む)
- `docs/dev/MASTER_SPEC.md` — 初期実装契約書(§3依存ルールR1-R5、§7モジュール契約、§12コーディング規約は今も有効)
- `docs/dev/RIBBON_API_CONFIRMED.md` — 社内AIリボンの確定API仕様+PM裁定
- `wintest/README.md` — このテスト環境自体の説明

**やること(ループ)**:

1. `python build/build_mybookshelf.py --dev` でビルド
2. `wintest/mock_ribbon/modMockRibbon.bas` がまだアドイン化されていなければ、
   READMEの手順に従って `リボンちゃん(検証用).xlam` として保存・有効化する
3. `powershell -ExecutionPolicy Bypass -File wintest\run_excel_tests.ps1` を実行
4. `wintest/result_*.log` の最新ファイルを読む
5. **FAILがあれば**: ログのエラー内容(VBAコンパイルエラー・実行時エラー・
   アサート失敗のいずれか)から原因のモジュール・行を特定し、`src/` 配下の
   該当ファイルを修正する。修正は以下を厳守:
   - `MASTER_SPEC.md` §12(コーディング規約)・§3(依存ルールR1-R5)に違反しない
   - 修正後、`python tools/vba_lint.py` と `python tools/run_lo_tests.py` を
     Windows側でも実行してみて(Pythonがあれば動く)ERROR 0を確認
   - 1〜3に戻ってやり直す
6. **全部PASSになったら**: `python build/build_mybookshelf.py --prod` で本番ビルド
   →もう一度 `run_excel_tests.ps1 -Target prod -MockAddinPath <ニセリボンのフルパス>`
   を実行し、リボン依存の配管(続けて質問/Wordで開く/スクショ取込)も
   ボタンのマクロを直接実行して動作確認する
   (例: `$excel.Run("'ブック名'!modUIShelf.OnIngestScreenshot")` のように
   COM経由でSubを直接呼べる。事前に何か画像をクリップボードにコピーしておくこと)
7. 見つけたバグと直した内容を日本語で一覧にして報告する
   (ファイル名・症状・原因・修正内容の4点セットで、1件1行程度に簡潔に)
8. 修正をコミットする(pushはしない。ユーザーに確認してから)

**やってはいけないこと**:
- このリポジトリの外のファイルには触れない(旧`notebook`リポジトリの別
  プロダクト`chatbot_v2`とは既に分離済みで無関係)
- ブランチを変えない・mainへの統合はしない
- テストを緑にするために `tools/vba_lint.py` や `tools/run_lo_tests.py` の
  検査ルール自体を緩めない(検査対象のコード側を直すこと)
- 見つけたバグを「なかったこと」にして握りつぶさない。直せなかったものは
  正直に「未解決」として報告する

**行き詰まったら**: 3回試して同じ箇所で失敗する場合は、無理に自己判断で
仕様変更せず、`wintest/result_*.log` の該当エラーメッセージそのままと、
自分の仮説を添えてユーザーに報告し、判断を仰いでください。

---

## 貼り付け用プロンプト(ここまで)
