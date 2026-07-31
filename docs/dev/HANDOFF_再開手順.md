# 再開手順（セッション中断対策・最終更新: R7実装中に作成）

5時間リミットで中断した場合、次のセッションはこのファイルから読むこと。
リポジトリ: `kojima1459/notebook`、ブランチ: `claude/internal-notebook-lm-chatbot-B6BE7`。

## 1. まず最初にやること

```bash
git log --oneline -10          # 直近の状態を確認
git status --porcelain         # 未コミットが無いか(あるはずだが念のため)
```

直近コミット `6e7d71e`(wip: R7実装の途中経過スナップショット)が最新。
これは**検証未了のスナップショット**。実機ビルドに使わないこと。

## 2. 現在地(ラウンド表)

| ラウンド | 内容 | 状態 |
|---|---|---|
| R1〜R5 | コードレビュー対応・取込コア・Hub統計/バッジ・UI統合・敵対的レビュー | **完了・検収済み・push済み** |
| R6 | 画像PDFのOCR取込(Ghostscript連携)+全ページ化けPDF振替 | **完了・検収済み・push済み**(`5e2a578`, 検収dist `a3f3757`) |
| R7 | UI/UX修復5件+応答性2件(要件定義書: `docs/dev/spec_20260731_R7_UIUX修復と応答性.md`) | **実装中に中断**。WIPスナップショット `6e7d71e` あり。**下記3を実施** |
| R8 | P2P/共有系修正12件(裁定書: `docs/dev/spec_20260731_R8_P2P修正裁定.md`) | **裁定済み・未着手**。R7完了後に着手 |

## 3. R7の再開方法

R7はOpusサブエージェントに実装を委任中だった。バックグラウンドエージェントは
セッション終了とともに失われるため、**再開後は完了通知が来ない**。
以下のいずれかで判断すること。

### 3-1. まず現状を検証する
```bash
python3 tools/vba_lint.py
timeout 1500 python3 tools/run_lo_tests.py
```
- lint ERROR 0 かつ 261件以上PASSなら、**R7はほぼ完了していた可能性が高い**。
  `docs/dev/spec_20260731_R7_UIUX修復と応答性.md` の受入条件
  (症状#2〜#9対応表・hub_tile_drawnログ・DoEvents挿入位置一覧)が
  報告されていないので、コードを読んで自分で確認・報告を作成してから
  検収コミットを打つ。
- ERRORが出る、またはテストが失敗するなら、**未完成の途中**。
  `git diff a3f3757 HEAD --stat` で変更ファイル一覧を見て、要件定義書の
  A-1〜A-5・B-1〜B-2のどこまで実装されているかをファイル単位で確認し、
  残りを新しいOpusエージェントへ引き継ぐ(要件定義書はそのまま使える。
  「前任エージェントの途中経過が `git diff a3f3757 HEAD` にある。
  未実装の項目だけ続きをやれ」と指示に追記すること)。

### 3-2. 中断時点で分かっていたこと(実装エージェントの最終報告は届いていない)
`git diff a3f3757 HEAD --stat` に出るはずのファイル(要件定義書の対象):
- `src/ui/modDash.bas`(A-1: 共通クローム化)
- `src/ui/modUI.bas`(A-2: EnsureAppView新設)
- `src/ui/modKnowledge.bas` / `modHub.bas` / `modHubStat.bas`(A-2呼び出し・A-3余白・A-4タイル2Shape化)
- `src/ui/modApp.bas`(A-5ヘッダー共通化)
- `src/ingest/modEmbed.bas` / `modEnrich.bas`(B-1 ETA表示。`modUtil.ProgressText`/`EtaText`が既に見えている→実装済みの可能性大)
- `src/ingest/modExtractorWord.bas` / `modExtractorExcel.bas` / `modShelfSync.bas` / `modShelf.bas`(B-2 DoEvents+busy制御)
- `src/ui/modUiLock.bas`(新設。B-2の共有ロックと推測される)
- `src/core/modUtil.bas`(B-1の純ロジック関数)
- `src/test/modTestsPure4.bas`(追加テスト)
- `tools/vba_lint.py`(新モジュール登録)

`modUiLock.bas` が新設されている場合、**要件定義書に無い設計判断**なので
中身を読んで役割を確認すること(IsBusy集約用の可能性が高い)。

## 4. R7完了後にやること

1. 検収(lint/テスト/両ビルド)→正式コミット→push
2. `docs/dev/spec_20260731_R8_P2P修正裁定.md` を読み、Opusへ実装委任
   (プロンプトの型は本セッションのR2〜R7委任と同じ。修正12件・受容2件、
   modHub等の容量制約に注意、要件外発見は自己修正せず報告、という縛りを
   必ず含めること)
3. R8完了後、検収→敵対的レビュー(推奨)→push→ユーザーへ総括報告

## 5. ユーザーへの未伝達・要確認事項

- **緊急運用アクション未確認**: 実機configに `knowledge_expire_days=0`
  を追加するようお願いしたが、ユーザーが実施したか未確認
  (R8のF4修正が入れば恒久対策されるが、それまでは出荷時ダミー共有パスの
  端末で失効タイマーが進行中の可能性がある)。再開後、最初のユーザー
  応答で確認すること。
- R6完了時に案内した「Ghostscriptフォルダを本ブックの隣に配置」の
  実機確認結果は未報告。

## 6. タスクリスト(TaskList)の対応

session内タスクIDは再開後のセッションには引き継がれない可能性がある。
`TaskList` を呼んで無ければ、本ファイルの§2の表を元に再作成すること。
主要な未完了: 「R7実装の検収」「R8実装の投入」「R8検収」「総括報告」。

## 7. 開いている調査・裁定書の一覧(全てpush済み・参照用)

```
docs/dev/spec_20260730_R2_取込コア.md
docs/dev/spec_20260730_R3_統計バッジ.md(+追記)
docs/dev/spec_20260730_R4_UI統合.md(+追記)
docs/dev/spec_20260730_R5_レビュー対応.md
docs/dev/spec_20260731_R6_画像PDF_OCR取込.md
docs/dev/spec_20260731_R7_UIUX修復と応答性.md   ← 今ここ
docs/dev/spec_20260731_R8_P2P修正裁定.md        ← 次
docs/dev/reference_公式OCRツール調査_20260731.md
docs/82_実機不具合対応記録_20260729.md
```
