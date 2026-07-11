# 再開メモ(セッション上限対策・随時更新)

> 目的: セッション5時間上限で中断しても、リセット後に即座に現在地から再開するための状態記録。
> 最終更新: 2026-07-11 06:5x UTC(Wave2終盤)

## 現在地

- **Wave0 完了**: MASTER_SPEC.md(実装契約書)= `mybookshelf/docs/dev/MASTER_SPEC.md`
- **Wave1 完了**(commit 89f59f4): 基盤8モジュール(src/core)+独立ビルド機構(build/)+テストハーネス(tools/)
  - 検証済み: vba_lint.py ERROR 0 / run_lo_tests.py --mode compile 全PASS
- **Wave2 ほぼ完了**: 6並列エージェント(extract/shelf/qa/pack/ui/opt)は **全員 ok=true で完了**。
  - 残: 2-X(テスト執筆エージェント)が実行中だった(modTestsPure.bas 書きかけ、modTestsExcel.bas 未着手)
- **Wave3〜5 未着手**

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
