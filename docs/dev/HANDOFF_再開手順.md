# 再開手順（セッション中断対策・最終更新: R11完了時点）

中断したら、次のセッションはこのファイルから読むこと。
**docs/dev/00_プロダクト憲章.md が全裁定の判定基準(必読)。**
リポジトリ: `kojima1459/notebook`、ブランチ: `claude/internal-notebook-lm-chatbot-B6BE7`。

## 1. 現在地

**最新コミット = R12裁定記録まで。コードのHEADはR11-I(56a9014)+dist再生成(a541474)。**
R1〜R11全体品質総点検まで全ラウンド完了・検収済み・push済み。テスト491件・lint WARN 3(全てテスト系)。
次: **API月次上限が回復したら docs/dev/spec_20260801_R12_未踏観点対応.md の
R12-1(最優先小粒: 失効フォールバック0/終了ボタンのAuto_Close/streak正規化ほか)から
実装を再開する**。R12-6(未実行監査4本)も同時に。生報告=audit_20260801_R12_*.md。
利用者の実機テスト(docs/45/44)は継続中(配布zipの注意事項はR12 spec §4)。

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
| R11-A2 | Opus | Word設定復元の無条件化(外部レビューHigh)+発行後始末統一 | b3d446e | 完了 |
| R11-C | Sonnet | 関所とログの全数配線(BlockIfIngesting/再入ガード/ツールバーログ/委譲先ログ) | ed977de | 完了 |
| R11-D | Opus | GS/COM/実機層堅牢化(観測性/タイムアウト/Dictionary/ComError/Word所有事実判定) | 6fb3707 | 完了 |
| R11-E | Sonnet | 進捗・文言・ヘルプ・UI細部(進捗配線/ヘルプ刷新/死にコード削除/自己取込ガード) | 130e314 | 完了 |
| R11-F1 | Opus | 容量救済の分割6本(純移設。実装側WARNゼロ化) | 3087a33 | 完了 |
| R11-F2 | Opus | 重複統合(UTF-8/移動平均/ページ分割/Timer)+保留回収 | c7587b7 | 完了 |
| R11-G | Haiku | docs同期+実機スモークテスト手順書(docs/45)新規 | 7f2216a | 完了 |
| R11-H | Opus | 敵対的レビュー裁定の最終修正8件(Word保全対称化/バッジ遅延通知ほか) | 1855097 | 完了 |

配布方法: GitHubの「Code → Download ZIP」→解凍→ dist/MyBookshelf.xlsm を開く
(dist/Ghostscript が隣にあるのでOCRも追加作業なし)。
大規模配布は `python3 build/build_mybookshelf.py --prod --zip`。

## 2. 残タスク(優先順)

2. **実機テスト待ち(利用者)**: docs/45(15分スモーク・1台)と docs/44(P2P・2台)。
   不具合報告が来たら次ラウンドとして裁定。
3. 記録済みの次期課題(裁定で受容・未対応):
   - ロック取得のTOCTOU(PoCでは受容。本格展開時に Open For Output 排他)
   - Dir(vbDirectory) イディオム実測24箇所の統一(次期クリーンアップ)
   - SwitchTo / Subscribe/Unsubscribe(意図的未使用・契約表で可視化済み)
   - ダミーローカルパスが Reachable=True(既定expire=0で無害化済み)
   - B9副作用: ホーム表示→Hub描画の間に最大3秒の間(経過観察)
   - 発行者不在部門の共有肥大(docs/30 §9-1 に注意記載済み)
   - nx_progress が nx_ 一括削除対象(受容・再作成されるので無害)
   - PathExists のopt層内相互参照 / 画像PDFでGS起動2回(受容)
   - E0801の OnToChat/OnGoChat側2件(activate_recovered ログで経過観察)
   - State Loss時の mFeedbackDone 消失(二重FB可能になるのみ・受容)
   - チャンク重複排除の先着帰属(設計判断・受容)
   - Excel強制終了時のWordゾンビ(プラットフォーム限界・受容)
   - NextEmbeddingArray のAPI形式変更耐性(現行形式では正・将来課題)
   - PACK_SAVE_FAILED の独自コード(E0704化は次期。docs/30に注記済み)
   - check_raw_activate の回避余地強化(マーカー位置限定等・次期)
   - ArchiveCurrent 初回_archive未作成時の過剰警告エッジ(次期)
   - AddFilesViaDialog の busy 早期戻りが無言(通常経路は関所が先に停止・次期)
   - modBoot 残2,244字(次の機能追加時に分割裁定)

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
