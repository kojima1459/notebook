# Nexus Agent — 再開ハンドオフ (RESTART)

最終更新コミット: `135e100`（Step3 完了時点）。作業ツリーはクリーン・全push済み。

## 0. 30秒で状況把握
- ブランチ: `claude/internal-notebook-lm-chatbot-B6BE7`
- 対象: `mybookshelf/`（Excel VBAの社内RAGチャットボット「Nexus Agent」/自己インストーラ配布 .xlsm）
- 実行環境は **Windows版Excelのみ**。LibreOfficeは **構文コンパイル検査だけ** に使用（実行・UNO不可）。
- 私(opus)はExcelランタイム非保有 → 検証は静的（lint/LO構文/ビルド自己検証/コードレビュー/幾何オーディット）。**実クリック・実描画・実共有フォルダ・P2Pは実機テストが必須**。

## 1. 3ゲート（変更後は必ず全部通す。cwd=`mybookshelf/`）
```
python3 tools/vba_lint.py                 # 契約・恒久ルール(下記5の罠)を静的検査
python3 tools/run_lo_tests.py             # 各モジュールをLibreOfficeで構文コンパイル
python3 build/build_mybookshelf.py --dev  # dev(.xlsm)ビルド+自己検証
python3 build/build_mybookshelf.py --prod # prod(.xlsm)ビルド+自己検証
```
新規 .bas は必ず `build/modules.json` に登録。closed契約モジュール（modStats/modAsk/modAppDef/modConfig/modUtil等）へPublic追加時は `tools/vba_lint.py` の CONTRACT も更新。

## 2. オーケストレーション方針（ユーザー指定）
- **opus(私)** = 上流設計/中流/ホットパス・closed契約・配線・**全成果物の検証検品**。sonnet/haikuがスタックしたらヘルプ。
- **sonnet** = 実装（隔離ファイルへ・並列可）。**disjoint file ownership** で並列させると衝突しない。
- **haiku** = 軽作業。※commitは overhead 回避のため実際はopusが直接実行している。
- Fableはレート制限で不可 → opusが代役。

## 3. 完成済み機能（コミット順）
- Vaultギャラリー / EXP・レベル(活動+感謝合算) / Nexusダッシュボード(KPI/バッジ/EXPバー) / K-Meansナレッジ地図(modCluster・Shape円) / 分析CSV出力(modAnalytics)
- **Phase4 P2P感謝状**(modP2P): ✅→作者へ感謝状、受領で感謝EXP(自己申告不可)。AD連携ID。I/Oリトライ+GC。
- **UI鉄壁化**: FreezeShapePlacement(絶対配置), BringFixedToFront(Z-Order), CapBubbles(40件上限), DoEvents, ScreenUpdating復帰, DPI(WordWrap+Margin)
- **記憶喪失対策**: modState(ui_state退避/復元) — 会話履歴prevU/prevA
- **ナレッジ自浄 P2Pノイズ集計**(Step1-3): 下記4

## 4. P2Pノイズ集計アーキテクチャ（直近の本命・実機未検証）
- **投票**: `modP2P.EmitNoiseVote(source)` → 共有 `…\noise\noise_<Fnv1a64Hex(source)>_<reporterAD>.txt`。1報告者1ファイル=**票の水増し不可**。
- **集計**: `modP2P.CollectNoiseVotes(silent)`（SyncNow/Boot から呼ぶ）→ ソース毎の**異なる報告者数**（=ファイル数）が `noise_global_threshold`(config既定2)以上かつ未解除なら `modStats.MarkGlobalExcluded`。台帳は削除しない（毎回再計算）。**fail-safe**: 共有ルート到達不能時は `ResetGlobalExcluded` せず既存除外を保持。
- **除外の適用**: `modStats.ExcludedSources()` = 個人ミュート(`noise:`本人1票即時) ∪ 組織的除外(`gexcl:`集計結果)。`modRetrieve.Search`/`SearchExpanded` 両ループが検索前に取得して論理除外（物理削除しない）。
- **管理者**: config `admin_users`(カンマ区切りAD)。`modP2P.IsAdmin/ClearNoise`。ダッシュボードに管理者だけ見える「組織的除外の管理」+「↩復帰」(`modDash.OnDashRestore` → ClearNoise→再集計→再描画)。
- **透明性**: Vaultで除外中カードは非表示にせず「⚠️組織的除外(調査中)」グレーアウト(`modVault.DrawOneCard`)。
- 関連config: `noise_global_threshold`, `admin_users`, `nexus_share_path`。

## 5. LO/実機の恒久的な罠（lintで自動検出。ハマると時間を溶かす）
1. **`base` を識別子にしない** — StarBasic予約語(Option Base)でLO構文チェックが**沈黙ハング**(実機Excelでは有効だが不可)。lint検出済み。`rem`も不可(コメント化)。
2. **文字列リテラルの末尾に `\` を置かない**（`"foo\"`）— LOトークナイザがハング。区切りは `& "\"` で連結。
3. **モジュールレベル宣言は必ず先頭**（プロシージャ後の宣言は実機コンパイルエラー）。lint検出。
4. **変数/引数名 == 同一モジュールのプロシージャ名 禁止**（実機コンパイルエラー）。lint検出。
5. **`ReDim x(0 To -1)` 禁止**（実機実行時エラー。LOは通る）。件数ガード必須。lint検出。
6. **1モジュール ~30000字上限**（lint検出）。modAskは分割済みだが依然上限近め→追記は別モジュールへ。
7. UI層(src/ui)を下位層(ingest/qa/pack/stats)から参照禁止(R1)。共有ヘルパーは core へ(modStateはcore)。

## 6. 主要モジュール地図
- core: modAppDef(定数/SH_*), modConfig, modUtil(Fnv1a64Hex/CsvToVector/DotProduct等), modGateway(LLM/embed), modLog, modState(状態退避), modTypes
- ingest: modShelf(取込/SourceList/TotalChunks), modShelfSync(フォルダ同期/SyncNow), modEmbed, modChunker, modExtractor*
- qa: modRetrieve(ベクトル検索/除外適用), modAsk(QAオーケストレーション), modFollowup(続けて質問の純関数), modPrompts, modRagParse
- pack: modPack(パック出力/取込・origin="pack:<作者>"), modP2P(感謝状+ノイズ集計+AD), modPii
- stats: modStats(EXP/バッジ/ノイズ除外)
- ui: modUI(Nexusチャット), modApp(Controller), modVault(倉庫), modDash(ダッシュボード), modCluster(地図), modAnalytics(CSV), modUIMain/modUIShelf/modUIDashboard(旧UI), modBoot(起動)

## 7. 実機テスト待ち（再開後にユーザーへ依頼）
1. P2Pノイズ: 3台で2人が同一資料を⚠️報告 → 全台で検索除外に収束するか
2. 管理者UI: `admin_users`該当者のダッシュボードにのみ「管理」節が出る/復帰で全員反映
3. fail-safe: 共有を一時切断 → 除外が消えず保持されるか
4. ダッシュボードのShape座標(実DPI)、ナレッジ地図の描画、感謝状の着弾

## 8. 既知の申し送り/次の候補
- ノイズ投票台帳(noise\)は無GCで増える → 将来、解除済み/古い票のクリーンアップ検討。
- 悪意ある大量⚠️報告への耐性は閾値(config)頼み → 監査ログ(誰が報告したか)は各noiseファイルに残るので調査可能。
- modAsk は上限近め → 次に触るなら feedback系 or history系を別モジュールへ更に分離。
- Vaultギャラリーで personal-mute(noise:)と gexcl の見分け表示は未実装(現状gexclのみバッジ)。

## 2026-07-18 セッション末尾スナップショット(usage上限前の保全)
- ブランチ: claude/internal-notebook-lm-chatbot-B6BE7 / HEAD=bc1658b(push済・作業ツリーclean)
- 検証: lint 0 ERROR(55モジュール) / LO 85 PASS 0 FAIL / dev+prod自己検証OK
- 完了済み: SPA UI+ポリッシュ(MS&ADグリーン/Toast/EmptyState) / グローバルロック /
  暗転・誤爆・GC競合封鎖 / 狂気Lv.1(modBitwiseOpt, binary_rag=FALSE既定) /
  Peek View(modPeek) / Mentor送受信+返信(modMentor) / チーム連帯ボード(modBoard:
  ビーコン=board\stats_*.txt、称号💡5+/🌟20+、sv:d|m|y:日付キー節約時間) /
  ツアー(modTour)+ヘルプ(modHelp) / 会話復元(modApp nexus_hist_u/a)
- フック集約点: modApp.LaunchNexus(Board/Questions/Help/Tour)+OnSend/OnActDrill
  (Peek/Mentor/SaveTurn)。新機能は全て「新規モジュール+1行フック+内部安全弁」方式。
- 残タスク: 実機Excel受入テスト(月曜PoC)のみ。コード側の未完・配線漏れなし。
- 実機で見る点: ツアー表示→スキップ/完了、?ヘルプ、サイドバー「みんなの節約」
  (共有フォルダ必要)、✅解決→節約時間加算→ビーコン反映(翌起動)、Mentor往復。
