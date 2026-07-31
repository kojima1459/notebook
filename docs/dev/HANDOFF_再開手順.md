# 再開手順（セッション中断対策・最終更新: R11完了時点）

中断したら、次のセッションはこのファイルから読むこと。
**docs/dev/00_プロダクト憲章.md が全裁定の判定基準(必読)。**
リポジトリ: `kojima1459/notebook`、ブランチ: `claude/internal-notebook-lm-chatbot-B6BE7`。

## 1. 現在地

**最新の完成コミット = HEAD(c7587b7)。R11全6波+G完了。敵対的レビュー→配布zip待ち。**
R1〜R11全体品質総点検まで全ラウンド完了・検収済み・push済み。テスト491件・lint WARN 3(全てテスト系)。
次: 敵対的レビュー(読取専用Opus・R11全差分)→司令塔裁定→修正→配布zip(--prod --zip)+docs/45実機スモークテスト手順書。

| R | 実装者 | 内容 | コミット | 状態 |
|---|---|---|---|---|
| R1〜R7 | Opus/Sonnet | レビュー対応/取込/統計/UI統合/OCR/UIUX修復 | 〜 | 完了 |
| R8/R8b/R8c | Opus/Sonnet | P2P修正14件+敵対的レビュー16件+再レビュー6件。テスト272→417 | 〜 | 完了 |
| R9 | Opus/Sonnet | Ghostscript同梱(dist常置)+検出4段階+案内カード+--zip+docs | 〜 | 完了 |
| R10-1 | Sonnet | Activate堅牢化とギャラリー救済(err91×4・ギャラリー無反応) | f2239ba | 完了 |
| R10-2 | Sonnet | GS解決の観測性と堅牢化(E0303に候補パス・カード状態記録) | 94c27d7 | 完了 |
| R10-3/3b | Opus | GhostscriptによるテキストPDF抽出第1選択+移設 | 3872a0a/f01b3d2 | 完了 |
| R10-5/5b | Sonnet | 取込進捗の可視化(進捗バナー・完了トースト) | b833692/91e0ac4 | 完了 |
| R10-4 | Haiku | docs追随 | 7cd686e | 完了 |
| R10c | Opus | 敵対的レビュー裁定の修正8件(H1/M1-M5/L1/L5/L6) | e117750 | 完了 |
| R10d | Opus | modShelf分割(modShelfBatch新設)+H2/M4/M5残 | 69a3364/16eeab4 | 完了 |
| R11-A | Opus | データ保全Critical(C1取込中終了禁止/C2Word所有判定/C3失敗検知/C4silent伝播/C5統計防御) | eded8b2 | 完了 |
| R11-B | Sonnet | #30恒久対策+Activate全数(C5ビューポート/C6素のActivate) | 0d9fa66 | 完了 |
| R11-C | Sonnet | 関所とログの全数配線(BlockIfIngesting/再入ガード/ツールバーログ/委譲先ログ) | b3d446e | 完了 |
| R11-D | Opus | GS/COM/実機層堅牢化(観測性/タイムアウト/Dictionary/ComError) | ed977de | 完了 |
| R11-E | Sonnet | 進捗・文言・ヘルプ・UI細部(OnChannels/UIUX修復/ヘルプ刷新/Caller不一致) | 6fb3707 | 完了 |
| R11-F1/F2 | Opus | 保守性・分割・重複統合(modVaultGallery/modSkin/重複統合/lint) | 130e314/3087a33 | 完了 |
| R11-F2 | Opus | 重複統合最終(UTF-8/移動平均/ページ分割/Timer)+保留回収 | c7587b7 | 完了 |
| R11-G | Haiku | docs同期(エラー表/構成/HANDOFF)+実機スモークテスト手順書(docs/45)新規 | TBD | 実施中 |

配布方法: GitHubの「Code → Download ZIP」→解凍→ dist/MyBookshelf.xlsm を開く
(dist/Ghostscript が隣にあるのでOCRも追加作業なし)。
大規模配布は `python3 build/build_mybookshelf.py --prod --zip`。

## 2. 残タスク(優先順)

1. **実機再テスト待ち**: docs/44_P2P実機テスト手順.md の10項目(a〜j。R10で(i)(j)追加)を
   利用者が実機2台で確認する段階。R10実装の進捗バナー・テキストPDF・Ghostscriptエラー観測性が正常か確認。
   不具合報告が来たら次ラウンド(R11)として裁定。
2. タスク#30: ダッシュボード等の右寄せクロムをウィンドウ幅基準に(恒久対策)。
3. 記録済みの次期課題(裁定で受容・未対応):
   - GetStatText がエラー値セルで err13(B1で消去は封鎖済み。読み側の保護は次期)
   - ロック取得のTOCTOU(PoCでは受容。本格展開時に Open For Output 排他)
   - Dir(vbDirectory) イディオム残存6箇所の ProbeIsReachable への統一
   - SyncSubscribed 未配線 / SwitchTo デッドコード
   - ダミーローカルパスが Reachable=True になる(既定expire=0で無害化済み)
   - modInsight/modApp の容量逼迫(WARN 11件)
   - B9副作用: ホーム表示→Hub描画の間に最大3秒の間
   - modShare の lint 契約テーブル未登録(ProbePath追加時の検討事項)
   - 発行者不在部門の共有肥大(docs/30 §9-1 に注意記載済み)
   - R10で発見した懸念(裁定で受容):
     - modUI・modVault の容量枯渇(残20数字)
     - nx_progress が RemoveNexusShapes の nx_ 一括削除対象(受容・再作成されるので無害)
     - PathExist のopt層内相互参照(受容)
     - 画像PDF取込でGhostscript起動2回(受容)
     - E0801の OnToChat/OnGoChat側 2件は環境要因の疑い(activate_recovered ログで経過観察)

## 3. 運用メモ

- 失効タイマー: 出荷既定 knowledge_expire_days=0(消えない)。本格展開時に30以上。
  利用者のconfig手動変更は不要になった。
- 体制: Fable=司令塔(要件・裁定・検収のみ)/Opus・Sonnet=実装/Haiku=雑務。
  敵対的レビュー(読取専用)→裁定→修正→再レビューのループが有効に機能した。
- LibreOfficeテスト(tools/run_lo_tests.py)は多重起動禁止。
- API月次上限で実装エージェントが落ちた場合: WIPを即コミットして保全し、
  軽作業のみで待つ(今回それで乗り切った)。
- モジュール数: 98本(実装)/テスト: 491件/WARN: 3本(全てテスト系)。
- R11での事実確認・修正メモ:
  - LibreOffice Private Const の参照不可: Public Const へ揃えて回避(modDashStatで実測)。
  - LogError context ラベル: Public エントリ名を指すこと(lintの参照チェックが文字列リテラル内も見る)。
  - modFeatures.InvokeFeature 引数: 最大6個。opt側の引数追加は末尾Optional固定。
  - Dir(vbDirectory): 実測24箇所(記録済み・次期統一)。
