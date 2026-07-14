# 再開メモ(セッション上限対策・随時更新)

> 目的: セッション5時間上限で中断しても、リセット後に即座に現在地から再開するための状態記録。
> 最終更新: 2026-07-12 【全Wave完了】プロジェクト完成

## 現在地(2026-07-14 第2ウェーブ: 機能改修D10-D13)

- 第1ウェーブ(公開API確定対応)は検品ゲート全緑でコミット済み(a72e353)。Opusレビュー+fixerは
  組織の月間利用上限で死亡→PM(メインループ)が代行検品(lint契約更新/AddInsフォールスルー確認/
  modBoot配線確認)。以後の中流レビューはSonnet代行。
- **実行中ワークフロー**: mybookshelf-features-d10-d13 / Run ID `wf_b468f419-881` / Task `weah0bku9`
  - スクリプト: /root/.claude/projects/-home-user-notebook-mybookshelf/defe7e58-c73e-5740-854b-09056e616d9b/workflows/scripts/mybookshelf-features-d10-d13-wf_b468f419-881.js
  - 再開: Workflow({scriptPath:<上記>, resumeFromRunId:"wf_b468f419-881"})(必要なら先にTaskStop weah0bku9)
  - N1 modUIMain(TTSボタン削除/対話型Word/続けて質問ボタン) N2 modGateway+modAsk+modPrompts+optMarkdown
    (prevU/prevA・深掘り移植・ExportAnswerAsDoc) N3 optVision+modShelf+modUIShelf(スクショ取込)
    N5 docs(Haiku) → N4 MASTER_SPEC+lint契約同期 → Sonnetレビュー
- **完了後のPM作業(未実施)**: ①レビュー指摘の処置 ②build --dev/--prod再ビルド ③最終検証
  ④コミット&プッシュ ⑤タスク#16完了化 ⑥ユーザーへ最終報告
- 保留: 高橋くんのモジュール採取待ち(effort/verbosity・LimitCheck真偽の裏取り→台帳§3)

## 現在地(2026-07-14 追加ラウンド: リボン公開API確定対応)

- 担当部署メール+公開ページ入手 → 確定台帳 `docs/dev/RIBBON_API_CONFIRMED.md` 作成済み(PM裁定D1-D9)。
- **実行中ワークフロー**: mybookshelf-ribbon-confirmed / Run ID `wf_bd8a9cc5-a21` / Task `wg0gjswbh`
  - スクリプト: /root/.claude/projects/-home-user-notebook/defe7e58-c73e-5740-854b-09056e616d9b/workflows/scripts/mybookshelf-ribbon-confirmed-wf_bd8a9cc5-a21.js
  - 再開: Workflow({scriptPath:<上記>, resumeFromRunId:"wf_bd8a9cc5-a21"})(必要なら先に TaskStop wg0gjswbh)
  - 内容: W1 modGateway(toolNブランド/AddIns検出/LimitCheck) W2 optVision(ChatGPTV確定規約)
    W3 optMarkdown(CellMarkDown確定規約+OpenAnswerInWord) W4 optTts撤去+UI+configシード
    W5 docs一斉更新(Haiku) → Opusレビュー2レンズ → fixer(lint CONTRACT更新込み)検品ゲート
- **ワークフロー完了後のPM作業(未実施)**: ①fixer結果検分 ②MASTER_SPEC §5/§6/§7同期改訂
  (spec_deltasをjournalから回収) ③build --dev/--prod再ビルド ④最終検証 ⑤コミット&プッシュ
- 保留(ユーザー判断待ち): D8 会話継続機能(prevU/prevA) / 高橋くんのモジュール抽出待ち
  (effort・verbosity正式仕様とLimitCheck真偽の裏取り)

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
- **Wave4 完了**: fixer全処置(修正7/棄却13=前回分修正済みを実コード確認/実機送り3)。E0504新設。ゲート全緑。
- **Wave5 完了**: ドキュメント8本(README/00/10/20/30/40/50/60)+使い方シート清書。Opusレビュー5件全反映+
  残件(使い方シートのTrust文言)もPMが直接修正・再ビルド済み。最終状態:
  lint ERROR0 / LO pure PASS84 FAIL0 / compile36本PASS / dev+prodビルド自己検証OK。
- **残タスク(ユーザー側・実機のみ)**: docs/40_受入チェックリスト15分.md を配布先Windows Excel(リボンあり)で実施。

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
