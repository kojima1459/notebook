# 再開手順（セッション中断対策・最終更新: R16完了時点）

中断したら、次のセッションはこのファイルから読むこと。
**docs/dev/00_プロダクト憲章.md が全裁定の判定基準(必読)。**
リポジトリ: `kojima1459/notebook`、ブランチ: `claude/internal-notebook-lm-chatbot-B6BE7`。

## 1. 現在地

**R16完了(要望4件→調査5班→R16仕様→波1〜3実装→敵対的レビュー2面→R16H裁定FA8+FB12全消化)。実機配布可。**
R1〜R16まで全ラウンド完了・検収済み・push済み。テスト1,391件・lint ERROR 0/WARN 4(全てテスト系)・
モジュール117本(実装側WARNゼロ)。仕様は docs/dev/spec_20260805_R16_要望4件.md と
docs/dev/spec_20260805_R16H_レビュー裁定.md(§3C/3Dの読み替え改訂を含む)。
R16の骨子: ①文言統一(節約した時間)/②取込中のPC作業両立(🗔作業用Excel=WScript.Shellで
excel.exe /x・白画面抑止=リポジトリ初のDeclare(user32.DisableProcessWindowsGhosting、
config freeze_keep_banner 既定on)・案内文とdocs)/③入念モードの複合質問分解(modAskMulti、
段0判定→論点別軽量検索→統合→critique/verify→union出典突合。auto/最大3論点。質問全体で
LLM論点数+6回)・逆質問=番号選択肢(clarify、ok=False非回答契約・「1と3」複数選択・TTL30分)・
精読=近傍チャンク束ね(modAskFocus、source基準・thoroughのみ・deep_neighbor既定2)・
deep深掘りの既出チャンク降格(modFollowup、followup全検索に適用)。
④は docs/dev/design_20260805_R17_構造グラフ設計.md(GO待ち)。
残: 利用者の実機テスト(docs/45スモーク全27項目、特に23〜27+19〜22)+R17のGO判断。
容量の分割必須ライン(次に触る波は先に分割裁定。1行でも足すとWARN帯):
**optOcrPage(残3) / modShelfBatch(残6) / modUtil(残8) / modTestsPure6(残9) / modSkin(残10) /
modTestsPure11(残19) / optVision(残26) / modBoot(残32) / modUI(残35) / modChunker(残38) /
modAsk(残43) / optGsTxt(残45) / modUIMain(残207) / modRetrieve(残259) / modPrompts(残386) /
modTestsPure12(WARN帯28,280字・追記禁止)**。

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

| R12-4 | Opus | 検索スケール恒久対策(ベクトルキャッシュ/norm_text前計算/binary_rag自動化/進捗+DoEvents) | 13fe5e3ほか | 完了 |
| R13-A1 | Opus | modExtractor分割(modExtractorPdf新設)+一時コピー名ASCII化(mbtmp_fnv16.ext) | 0789a87 | 完了 |
| R13-A2 | Opus | GS根治: フラグ内容検証/空出力の画像PDF分類/生存監視型待機/WMI起動+PID/kill/診断tail化/StageBanner | 3a89aa7 | 完了 |
| R13-A3 | Opus | 薄い抽出の品質ゲート/FriendlyFailMsg/OCR経路の安全コピー/全段階バナー/送信・終了保護/Word試行短縮 | c5313d3 | 完了 |
| R13-B | Opus | 深掘りのスコープ限定検索+会話出典メモリ+armed followup(チップ)+LLM段階計測・実況 | 8515d61 | 完了 |
| R13-C | Sonnet | Hub「みんな」意味化+未接続/未設定表示+部別集計(ビーコンteam列)+係数config化+チャット導線前方化 | edb91ce | 完了 |
| R13-Fix1 | Opus | レビュー裁定: GS判定窓4000字/kill前PID本人確認/OCR上限是正/生存監視の実効化/optGsProc分割 | 16c5d50 | 完了 |
| R13-Fix2 | Opus | レビュー裁定: armed漏れ止め/終了承諾取消/未設定表示/段ログ集約(ask_steps)/スコープ時prefilter回避 | bdae2b2 | 完了 |
| R14-A | Opus | コピー根治(Stream+サイズ突合)+OCR20頁バッチ全量化(optOcrPage新設・上限100頁・正直メモ) | d2a8ba8 | 完了 |
| R14-B | Opus | 質問カウンタ3箇所+一般モードfeedback解禁(誤共有ゲート)+入念モード6段(modAskThorough新設)+可読性 | df5f535 | 完了 |
| R14-C | Sonnet | ヘッダー整列+ギャラリー可視化+着せ替え5件+質問例オンデマンド生成+チャンク拡張子別設定 | 83387a4 | 完了 |
| R14-FixA | Opus | レビュー裁定: GS起動失敗の切り分け/バッチ中断の可視化/共有読みコピー復活+理由4分岐/PID kill/上限メモ正直化 | 6e6a4db | 完了 |
| R14-FixB | Opus | レビュー裁定: 発信の関所(修正/gap/感謝状)/出典タグ]対応/質問例後始末+メモリ/着せ替えガード/docs追随 | 282f183 | 完了 |
| R15波1〜4+Fix | Opus/Sonnet | OCR254頁対応(ETA/中断/再開/確認)+E0202/save_fail根治+レビュー裁定FA9/FB14 | 0ce3f7a〜f65251a | 完了 |
| R16波1 | Sonnet | ①文言統一+②作業用Excelボタン/白画面抑止(初Declare)/案内+docs | 635cfef/24e6f18/74712a7 | 完了 |
| R16波2 | Opus | ③-A複合質問の分解→統合(modAskMulti+段0判定+union出典突合+テスト57件) | d528286/37352c2/461e1d2 | 完了 |
| R16波3 | Opus | ③-B逆質問番号選択肢+③-C精読(modAskFocus)+③-D既出降格+裁定1〜3+docs+テスト82件 | 7251a25〜f72f12a | 完了 |
| R16-Fix | Opus | R16H裁定FA8(精読source基準化/逆質問非回答化/降格全followup化ほか)+FB12 | e71b940/e5e8507/bbef504 | 完了 |

### R12-4 で増えたもの(次に触る人が最初に知るべき3点)

1. **新モジュール `modVecCache`(src/qa)** = セッション内ベクトルキャッシュ+
   **埋め込み世代カウンタ**。my_vectors の vector_csv は1セッション1回しか
   パースしない。世代を進める場所は3つだけ: `modEmbed.EmbedPending`(書込あり時)・
   `modEmbed.MarkAllForReembed`・`modShelfStore.RemoveVectorsByIds`。
   **ベクトルを書く/消すコードを足したらここを呼ぶこと**(呼び忘れると
   「古いベクトルで検索し続ける」という無言の誤り)。取込・同期の開始時
   (modShelfBatch/modShelfSync)は `ResetVecCache` でメモリを解放している。
   構築失敗(err7)は捕捉して従来経路へ落ち、usage_log に `veccache_fallback`。
2. **my_knowledge に第10列 `norm_text`** = 照合用の正規化済みテキスト
   (`modSparse.MatchDocText` が唯一の作り手)。取込時に前計算し、空欄の行は
   検索時に遅延バックフィルする。**行の詰め直しは必ず10列を運ぶこと**
   (modShelfStore の KNOWLEDGE_COLS)。要約・キーワードを書き換えたら
   norm_text を空へ戻す(modEnrich がそうしている)。
3. **config `binary_rag_auto`(既定TRUE)** = binary_rag=FALSE でも
   チャンク数 >= binary_rag_min なら粗選別を自動有効化。止めるには
   binary_rag_auto=FALSE。判定は `modBitwiseOpt.ShouldPrefilter`(純ロジック)に
   1本化してあり、真理値表は modTestsPure9 が固定している。

### R13 で増えたもの(次に触る人が最初に知るべき6点)

1. **新モジュール `modExtractorPdf`(src/ingest)** = PDFフォールバック連鎖+一時コピー
   (旧modExtractorから純移設)。一時コピー名は `mbtmp_<fnv1a64hex16>.<拡張子>`
   (TempBaseNameFor)。元ファイル名は表示/ログ専用。**NFD分解文字・CP932非対応文字
   対策の要**なので、抽出経路を足すときは必ず CopyToLocalTemp 経由にすること
   (OCR経路は modShelfVision が自前でコピー取得+全出口Kill)。
2. **新モジュール `optGsProc`(src/opt)** = GSプロセスの起動(WMI優先→wsh.Run退避)と
   停止(taskkill前に Win32_Process で本人確認)。優待機は optGsTxt.WaitGsTextDone
   = アイドル(vision_pdf_timeout_sec)+絶対上限(gs_abs_timeout_sec 既定1200)の
   二段で、gs_out.log のページ進行/出力サイズが進む限り待つ。どのシグナルが
   効いたかは usage_log `gs_progress_signal`(page=..;size=..;none=..)で観測できる
   (実機初回はこれを必ず確認。両シグナル0なら固定タイムアウト相当に退化している)。
3. **空txtwrite出力の分類** = optOcrCore.ClassifyGsTextResult(純・真理表固定)。
   rc=0+ページ進行あり→「画像PDF」としてERR_303→OCRへ。**「txtwrite出力を
   読めず」でWordへ流れてゴミ登録される事故(RC1)の再発防止の要**。判定材料の
   gs_out読み窓は4000字(300字に戻すとバナー+xref警告で再発する)。
4. **armed followup** = 「続けて質問/深掘り」はInputBoxではなく、チップ
   (Shape `nx_fchip`)を出して nx_input へ誘導し、OnSend が消費する方式。
   消費はモード分岐より先(一般モードなら黙って解除)。RAG↔一般トグルで解除、
   エフォートトグルでチップ文言更新。deepのfollowupは会話出典スコープ内
   多段検索(modAskRetrieve.RunDeepScoped、<2件で無スコープへ自動退避)。
   **スコープ指定時は binary_rag 粗選別をバイパスする**(EnabledScoped)。
5. **ask_steps 1行集約** = LLM/埋め込みの段階レイテンシは usage_log に
   毎段書かず、modGateway のバッファへ積んで modAsk が質問1回につき
   `ask_steps` 1行で書く(usage_log 2000行ローテを圧迫しない)。
6. **ビーコン第10列 team** = 部別集計用(modP2PIo.TeamCodeOf/IsTeamCode/
   SanitizeId経由)。user_department は規約(英大数4-6字)に合う時だけ採用。
   共有パス既定は空文字になった(未設定=「未設定」表示、設定済み到達不能=
   「未接続」表示。ダミーパス出荷は廃止)。

### R14 で増えたもの(次に触る人が最初に知るべき7点)

1. **新モジュール `optOcrPage`** = OCRのバッチ制御(GSを20頁ずつ複数回起動、バッチごとに
   JPEG逐次削除、頁間DoEvents、PID捕捉+タイムアウト時KillGsTree)。上限は
   vision_pdf_max_pages(既定100)。RenderBatch は3値(完了/起動失敗/タイムアウト)で、
   起動失敗は ghostscript_path 案内、途中中断は truncated=True+中断メモ+partial
   (「頁が黙って欠ける」経路は全て塞いだ)。gs_abs_timeout_sec は資料単位の絶対予算
   (バッチ間で残額を配分)。
2. **コピーは2段構え** = Dir$でANSI可視なら旧来の1MB分割・共有読み(ロック中でも読める・
   メモリ一定)、不可視(NFD等)や空読み時のみ ADODB.Stream。サイズ突合必須。失敗理由は
   src_empty/locked/too_big/name_busy/特殊文字 の5系統で文言が分かれる
   (modLog.CopyFailMsgOf が一元管理)。0バイトコピーがGSへ届く経路は消滅。
3. **新モジュール `modAskThorough`** = 入念モード専用6段(再ランク強化→資料要約→下書き→
   自己批判→批判反映検証→機械的出典突合)。deep/quickは不変。出典タグ照合は
   modPrompts.SourceTag が唯一の書式源。検証段エラー注記は Decorate 後に付く
   (履歴/共有に混入しない)。
4. **発信の関所** = 部内への発信(EmitVerifiedQA/EmitThanks/EmitGap/EmitCorrection)は
   全て modAsk.CanShareInsight()(=ShouldEmitInsight(mLastMode, mLastNHits))を通る。
   一般モード・生成失敗ターン・検索0件残留からの誤共有は構造的に不可能になった。
   一般モードの「解決した」はカウント/節約時間のみ加算(共有なし)。
5. **回答の可読性** = modLive.NormalizeAnswerText(markdown混入の保険変換)+vbCr段落化+
   ■見出し段落の太字化(StyleAnswerParas)。バブル本文はvbCr、外へ出す時は
   TargetText がvbLfへ戻す(コピーはCRLF)。復元バブルも同経路。
6. **質問例のオンデマンド生成** = modStarter がシード無しでも my_knowledge から
   1回のLLM呼び出しで生成しmodStateへキャッシュ(指紋=manifest件数+日付シリアル)。
   プレビュー読みは対象8行のみ(全列一括読み禁止)。
7. **チャンクは拡張子別** = chunk_*_chars_{pdf,docx,doc}(target450/overlap100/max900)。
   グローバル既定とExcelは不変。既存資料は再取込しない限り不変(入れ直す場合は先に削除)。

R14で受容した次期課題: ADODB.Stream共有モードの実機検証(開いたままのxlsx/pdf取込=
docs/45項目15で確認)/全角コロン等の異形出典タグは検査対象外(偽陽性なし)/
OCR中断時のJPEG削除は実行される(OCR済みのため実害なし)/一般モードのno_hit gapは
RAG限定のまま/config実キー数は約120(MASTER_SPECは固定値を書かない)。

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
   - modBoot.Boot(:165でApplication.EnableEvents=False)→SyncNow(:464でTrueへ
     復帰)のEnableEvents入れ子崩れ: Boot内5)でSyncNowを呼ぶと、SyncNow自身が
     戻り際にEnableEventsをTrueへ戻してしまうため、Boot手順6)〜8.6)が
     Trueのまま走る(Bootが意図した「起動シーケンス全体を抑止」が5)以降
     効かなくなる)。R16-2e(2026-08-05)で発見・記録のみ。R16では触らない。
   - Q&A長時間経路への🗔/中断の到達手段(R16H FB-4/B-H1): 質問の実行中は❓も
     バナー内ボタンもクリックが届かないため、「🗔作業用Excel」は取込中にしか
     押せない。docs/10は「質問を送る前に開いておく」運用で暫定回避したが、
     Q&A側にも到達可能な導線(送信前の常設ボタン・ESC以外の中断手段)が要る。
   - **R13で受容した次期課題**: 深掘りfollowupは retrieve_mode=single でも
     スコープ内多段を実行(機能仕様として明記・escape hatchの例外)/
     ReadTextTail のDBCS境界で先頭1文字化けの可能性(診断限定・受容)/
     ナレッジ画面ツールバー総幅は未解決(チャットへは前方化で救済済み。
     折返し or リサイズ再描画は次期)/embed_transport=direct 時の質問側
     embed_step 行は未集約(既定ribbonでは無影響)/OCR経路のWMI起動はPID破棄
     (kill不能はOCR側は従来どおり)/ThinExtract のページ数は maxPages
     打ち切り後の値/OnSend弱音キーワード経路はarmed維持/本番ビルドの
     ×ボタン終了は無防備のまま(BeforeCloseはdev構成のみ。終了ボタン側のみ保護)
   - **R18-1gで事実を確定・次期課題化(2026-08-05)**: 上記の「BeforeCloseはdev
     構成のみ」を調査agent0がfile:lineで裏取りした(build_mybookshelf.py:1002-1010
     が Workbook_Open だけを書き込む/Auto_Close に Cancel 引数が無い)。
     **「取込中終了禁止ガードは開発構成のみ」**を MASTER_SPEC §7.6 に明記した。
     実機で×・最小化が効かないのはガードではなく DisableProcessWindowsGhosting の
     副作用。本番でも取込中の終了を守る手段(インストーラ側で BeforeClose を
     注入する/Application.OnKey で退避する等)の設計は**次期**に回す(R18では
     事実の記録のみ。修正はしない=仕様どおり)。
   - **FindKeyRow裁定前提の更新(2026-08-01 R12-5-11)**: spec_20260731_R11
     §9「FindKeyRow線形探索(103キーでms級)」は modConfig.FindKeyRow
     (config = 103キーで固定・安定)の前提であり、modStats.FindKeyRowには
     そのまま適用できない。my_statsは "thx:"/"ins:" のnonce行が記録先
     (modP2P.MarkNonce/modInsight系)で、GC(thanks_gc_days・GcOldNonces)は
     入っているものの周期実行のため、GCが効くまでの間はS(my_statsの行数)が
     数百〜千行に成長し得る事実がある。線形探索の即時最適化は見送り継続
     (体感数百ms。まず前提行数の事実更新のみ)だが、「103キーでms級だから
     全FindKeyRowが軽い」という読み方はしないこと。

## 3. 運用メモ

- 失効タイマー: 出荷既定 knowledge_expire_days=0(消えない)。本格展開時に30以上。
  利用者のconfig手動変更は不要になった。
- 体制: Fable=司令塔(要件・裁定・検収のみ)/Opus・Sonnet=実装/Haiku=雑務。
  敵対的レビュー(読取専用)→裁定→修正→再レビューのループが有効に機能した。
- LibreOfficeテスト(tools/run_lo_tests.py)は多重起動禁止。
- API月次上限で実装エージェントが落ちた場合: WIPを即コミットして保全し、
  軽作業のみで待つ(今回それで乗り切った)。
- モジュール数: 98本(実装)/テスト: 491件/WARN: 3本(全てテスト系)。
  ※R12-4時点: 実装90本+テスト11本=101本(modVecCache/modTestsPure9を追加)、
    LOモード1のPASSは669件、WARNは3本(modTestsPure/2/5)のまま。
  ※R13完了時点: 105本(modExtractorPdf/optGsProc/modTestsPure10を追加)、
    LOモード1のPASSは818件、lint ERROR 0/WARN 3(modTestsPure/2/5)のまま。
  ※R14完了時点: 109本(optOcrPage/modAskThorough/modTestsPure11/modTestsPure12を追加)、
    LOモード1のPASSは1,036件、lint ERROR 0/WARN 3(modTestsPure/2/5)のまま。
  ※R15完了時点: 114本(optOcrEta/optOcrCache/modTestsPure13/modTestsPure14を追加)、
    LOモード1のPASSは1,228件、lint ERROR 0/WARN 4(modTestsPure/2/5/12。全てテスト系)。
  ※R16完了時点: 117本(modWorkExcel/modAskMulti/modAskFocus/modTestsPure15を追加)、
    LOモード1のPASSは1,391件、lint ERROR 0/WARN 4(同上)のまま。
- **R15で記録した次期課題・確定事項**:
  - ChatGPTV は待ち秒数引数なし・同期永久待ち(実機確認済み)。VBA側の根治は不可能。
    リボン側に引数が追加されたら optOcrPage の呼び出し1行で反映可能。
  - 複数の大型PDF一括取込では事前確認が1件ごとに出る(Nが抽出後にしか判明しないため
    構造上不可避・受容)。確認ダイアログ表示中は後続ファイルも停止する(モーダルの性質)。
  - 手動/自動同期は事前確認なしで長時間OCRが走り得る(silent設計・R11-A C4準拠。
    自動同期は既定OFF。チェックポイントで被害限定)。
  - modEmbed 単独実行(同期の埋め込み再開等)では中断ボタンは出ない(ESC中断は既存)。
    OCR起点で出たボタンはベクトル化中も残り、押せば埋め込みループも止まる。
  - modGateway 全捕捉による ESC(err18)握りつぶし(modEmbed/modEnrich のESC中断が
    CallLLM内で失敗文字列化され得る): 共有コアのため未着手。フラグ方式が代替。
  - SourceList の manifest/pack 同名非合算(pack利用時のみの表示問題・受容)。
  - vba_lint が手続き内 Const を module-level と誤認する(KEEP_DAYS衝突で実測。
    回避は改名。lint側の修正は次期)。
  - SleepMs が modEmbed/modExtractorPdf に private 複製2箇所(共通化するなら基盤層へ)。
  - modShelfSync の分割(現27,819字。次に本格的に触る波は先に分割裁定)。
  - EndClock の「翌 」は2日以上先でも「翌」(現行最大ETAでは到達不能・受容)。
  - キャッシュ復元は「そのバッチをGSが描けた」ことが前提(描画0枚バッチの頁は
    キャッシュがあっても打ち切り。頁番号整合を優先した設計判断)。
  - 32,000字超で切り詰めた頁はキャッシュ保存しない(次回読み直し・FB-10)。
- **R16で記録した次期課題・確定事項**(詳細は spec_20260805_R16H_レビュー裁定.md の「記録のみ」):
  - usage_log のリング(log_max_rows=2000)回転が新タグ(decomposed/multi_*/neighbor_*等)で
    速まる。観測価値優先で容認。次期でタグ整理を検討(節約時間の前月比が早く欠ける)。
  - StatusBar 通知(E0904等)はQ&A中は数秒で上書きされ実質見えない(err_logで追跡可。
    次期でトースト化検討)。取込中は modUiLock の文言を上書きしたまま残る。
  - source型逆質問(「2-②」教育)と topic型(「1と3」)の文言完全統一は次期
    (入力互換は R16H FA-7 の区切り拡張で確保済み)。
  - 逆質問TTL失効直後の番号入力への能動ガード(現状は普通の質問として流れる。
    案内文の「30分で無効」明記のみ実施)。
  - single-thorough/deep の経過時間表示(分解経路のみ実装済み。B-M7残)。
  - 既存の文字列リテラル末尾区切り文字パターン2件(modHelp.bas OnShareSetup付近/
    modTelemetry)。LO実測は通過・実害未確認(EDGE_CASES §1.3 のトリガ条件が実際は
    狭い可能性)。新規コードは Chr$(92) 経由を徹底。
  - mock_llm では decompose は常に single(FB-2で固定)。R16-3系のデモは実LLMのみ。
  - 精読(NeighborExpand)の資料キーは source(表示・スコープと同じ一意基準)。
    chunk_id のハッシュ部は「チャンク本文ハッシュ」であり資料キーではない(FA-1の教訓。
    テストは実データ形状=行ごと異ハッシュで書くこと)。
  - deep深掘りの精読適用は取り下げ(nHits表示契約と衝突・R16H §3C改訂)。
  - 非#ERR文字列は BuildErrorAnswer を素通しする契約(FA-2。非回答ターン=逆質問の表示経路)。
- R11での事実確認・修正メモ:
  - LibreOffice Private Const の参照不可: Public Const へ揃えて回避(modDashStatで実測)。
  - LogError context ラベル: Public エントリ名を指すこと(lintの参照チェックが文字列リテラル内も見る)。
  - modFeatures.InvokeFeature 引数: 最大6個。opt側の引数追加は末尾Optional固定。
  - Dir(vbDirectory): 実測24箇所(記録済み・次期統一)。
