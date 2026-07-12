# 再開メモ(セッション上限対策・随時更新)

> 目的: セッション5時間上限で中断しても、リセット後に即座に現在地から再開するための状態記録。
> 最終更新: 2026-07-12 Wave4実行中(レビュー6体完了/fixer実行中)

## 現在地

- **Wave0 完了**: MASTER_SPEC.md(実装契約書)= `mybookshelf/docs/dev/MASTER_SPEC.md`
- **Wave1 完了**(commit 89f59f4): 基盤8モジュール(src/core)+独立ビルド機構(build/)+テストハーネス(tools/)
  - 検証済み: vba_lint.py ERROR 0 / run_lo_tests.py --mode compile 全PASS
- **Wave2 完了**: 6並列エージェント全員 ok。検品でR1矛盾4件→PM裁定でR1例外を4件に拡張(lint+MASTER_SPEC修正済み)。
  ハーネスの.clsヘッダ除去バグも修正済み。lint ERROR0 / LOコンパイルは modTestsPure(2-X書きかけ)以外全PASS。
  - 残: modTestsPure.bas完成+modTestsExcel.bas新規(Wave3のテスト担当が実施)
  - Wave3持ち越し判断事項: (a) Hit型にfull_text追加(現状プロンプトに120字プレビューしか渡らない品質問題) →
    modTypes/modRetrieve/modPrompts改修で対応する(PM裁定済み)。(b) modAskが仮定した質問セルNamed Range
    "mb_question" をmodUIMain.EnsureLayoutが実際に定義しているか照合。(c) ask_*_totalのBump元がmodAsk.Answer
    唯一であることの確認。(d) Wave1申し送りのHex$/非BMP文字挙動はpureテスト実行で確認
- **Wave3 実行中**: 3-R(整合修正: Hit.full_text追加/modRetrieve・modPrompts改修/mb_question照合/カウンタ確認)= **完了ok**。
  3-T(modTestsPure完成+modTestsExcel新規+LO純テスト全緑)= セッション上限で死亡・未完了(リセット3am UTC後にresumeFromRunIdで3-Tのみ再実行せよ。このスナップショット時点で
  modTestsPure/modUtil/run_lo_tests.pyに作業中変更あり。PURE_ALLOWLISTにmodAppDef/modShelfSync/modPack追加済み)。
- **Wave3 完了**: 検品ゲート全緑(lint ERROR0 / LO compile 36本PASS / pureテスト PASS77 FAIL0)。
  3-T成果: modTestsPure分割(Pure2)+modTestsExcel(実機E2E)+テスト側Hit配列添字バグ修正。
  既知の残確認事項(実機のみ): LOはUDT配列テストをスキップ(CanUseTypeArrays=False)→ChunkPages/プロンプト系は
  実機の🩺診断(RunAllPureTests)で確認する(受入チェックリストに載せること)。
- **Wave4 実行中**: Opusレビュー6レンズ = **全員完了(指摘15件: 3+2+3+3+1+3)、キャッシュ済み**。
  fixer(sonnet: 指摘を検証→修正→ゲート全緑)が実行中に中断の可能性。
  - Wave4再開: スクリプト /root/.claude/projects/-home-user-notebook/defe7e58-c73e-5740-854b-09056e616d9b/workflows/scripts/mybookshelf-wave4-wf_04c6c4fb-942.js
    を Workflow({scriptPath:<上記>, resumeFromRunId:"wf_04c6c4fb-942"}) で再開(レビュー6体はキャッシュから即返り、fixerのみ再実行。
    必要なら先に TaskStop wtyicql4v)。指摘一覧はjournal.jsonl(同ディレクトリのsubagents/workflows/wf_04c6c4fb-942/)からも読める。
  - fixer完了後の検品ゲート: python3 tools/vba_lint.py ERROR0 / python3 tools/run_lo_tests.py 全緑
    → fixerのrejected(棄却根拠)とdeferred(実機送り)をPMが妥当性確認 → コミット&プッシュ → Wave5
- **Wave5 未着手**: build --dev/--prod(--allow-missing禁止)→バイナリ検証→ドキュメント(使い方/運用保守+Eコード表/
  撤去手順50/受入チェックリスト40/開発者ガイド)→Opus可読性レビュー→最終コミット→最終報告

## Wave3 ワークフロー再開情報
- スクリプト: /root/.claude/projects/-home-user-notebook/defe7e58-c73e-5740-854b-09056e616d9b/workflows/scripts/mybookshelf-wave3-wf_efd8c3ba-e36.js
- Run ID: wf_efd8c3ba-e36(3-R完了=キャッシュ済み。resumeFromRunIdで3-Tのみ再実行)
- 再開: Workflow({scriptPath:<上記>, resumeFromRunId:"wf_efd8c3ba-e36"})(必要なら先にTaskStop wfhmaetrn)
- 3-T再実行時の注意: 中断時点の作業中ファイル(modTestsPure等)が中途半端な可能性 → プロンプトに従い下書き扱いで検証・完成
- 検品ゲート(3-T完了後に必ず): python3 tools/vba_lint.py ERROR0 / run_lo_tests.py --mode compile 全PASS /
  --mode pure 全緑 → コミット→Wave4(Opus敵対的レビュー6レンズ)へ

## ワークフロー再開情報(このセッション内でのみ有効)

- Wave2 スクリプト: `/root/.claude/projects/-home-user-notebook/defe7e58-c73e-5740-854b-09056e616d9b/workflows/scripts/mybookshelf-wave2-wf_1e1131d8-783.js`
- Wave2 実行中 Run ID: `wf_34ca95c7-9b6`(6エージェント完了=キャッシュ済み。resumeFromRunId 指定で 2-X のみ再実行される)
- 再開コマンド: `Workflow({scriptPath: <上記>, resumeFromRunId: "wf_34ca95c7-9b6"})`
  (先に TaskStop wlc7q5ov0 が必要な場合あり)
- セッションが完全に失われた場合(resumeFromRunId が効かない場合): 2-X 相当のエージェントを単発で
  スクリプト内 Wave2X プロンプト(スクリプトファイル末尾に全文あり)のまま起動すればよい。
  6グループの成果はすべて git にコミット済みなので出戻りなし。

## Wave2 後の段取り(ユーザー指示: Wave2完了→一旦コミット→Wave3)

1. 検品: `cd mybookshelf && python3 tools/vba_lint.py && python3 tools/run_lo_tests.py`(モード1+2)
2. コミット&プッシュ(正式コミット: "feat(mybookshelf): Wave2 機能モジュール全実装+テスト")
3. **Wave3**: 統合検証ループ — lint+LO全緑まで修正(エージェントに全体を回させる。
   未解決の相互参照・契約ズレは6エージェントのconcernsをjournal.jsonlから読むこと:
   `/root/.claude/projects/-home-user-notebook/defe7e58-c73e-5740-854b-09056e616d9b/subagents/workflows/wf_34ca95c7-9b6/journal.jsonl`)
4. **Wave4**: Opus敵対的レビュー6レンズ(①VBA落とし穴 ②契約適合/統合 ③非エンジニアUX ④保守性 ⑤機能分離/撤去 ⑥エッジケース§13網羅)→検証→修正
5. **Wave5**: `python3 build/build_mybookshelf.py --dev` と `--prod`(--allow-missing 禁止)→バイナリ自己検証→
   ドキュメント執筆(使い方/運用保守+エラーコード表/機能撤去手順=docs/50/受入チェックリスト15分=docs/40/開発者ガイド)
   →Opus可読性レビュー→最終コミット&プッシュ→ユーザー向け最終報告
6. タスク管理: TaskList #11(Wave2)→#12(Wave3)→#13(Wave4)→#14(Wave5) を順に更新

## 不変の前提(忘れそうなこと)

- ブランチ: `claude/internal-notebook-lm-chatbot-B6BE7` のみ。main統合禁止。プッシュは `git push -u origin <branch>`
- mybookshelf/ 外(V2資産)への変更禁止
- ビルドは modules.json 駆動。最終ビルドで --allow-missing を使わない
- LO既知差分: 配列返しFunctionはハーネスが自動書換(tools/README.md 技術メモ)。LOのpureテスト失敗時はまずそこを疑う
- Wave1エージェントの申し送り(要Wave3確認): Hex$負数挙動 / 非BMP文字のLen挙動(Excel vs LO) /
  modDiagがmy_knowledge列順(9列目=embedded)に依存 / ui_stateはkey-value2列
