# 再開手順（セッション中断対策・最終更新: R20完了時点）

中断したら、次のセッションはこのファイルから読むこと。
**docs/dev/00_プロダクト憲章.md が全裁定の判定基準(必読)。**
リポジトリ: `kojima1459/notebook`、ブランチ: `claude/internal-notebook-lm-chatbot-B6BE7`。

## 1. 現在地

**R20完了(実機第7報①〜⑨=調査5班→波A〜D順次→敵対的レビュー2面+おもてなし監査→R20H裁定FA16+Haiku掃除波 全消化)。実機配布可。**
R1〜R20まで全ラウンド完了・検収済み・push済み。
テスト1,843件・lint ERROR 0/WARN 12(容量WARNのみ)・モジュール134本。
R20の仕様: spec_20260806_R20_実機第7報.md + spec_20260806_R20H_レビュー裁定.md。
**実機検証は未実施**(docs/45スモーク+R20新機能: 余白の窓幅追随とリサイズ再フィット/⚡仕上げ/初回ウィザード/MS&AD配色)。
R20の骨子: ⑦余白3層根治=帯だけでなく(層1)中身の窓幅追随(Dashカード幅130-220pt可変+
帯内センタリング/本棚gallery3-6列可変+table J列吸収/ピルContentRight統一)・(層2)縦の
行高明示範囲を実寸+αへ縮小(Hub60→40行/Dash120→40行/本棚412→実測。UseStandardHeightで
既存ブックも毎回救済)・(層3)測定頑健化(ViewportWidth他ブックガード=FB-5同型/クランプ
1600→3000pt/Dash全画面化順序)+Workbook_WindowResizeデバウンス0.7s自動再フィット
(dev/prodインストーラ両対応・OnTimeゾンビ防止・busy3情報源=modUiLock+modShelf+modShelfSync)/
①深掘りの一般アシスタント対応(ゲートをmGenPrevU系へ分岐・先消し順序逆転・ShowToast化・
OnClearChatの一般履歴クリア新設)/⑧モード可観測性(回答フッターにモード名+段数+所要を永続
表示・dispersionログのmode列正常化・自己エコー防御(全角対応)・thorough専用分散閾値
thorough_dispersion_gap_x100=25・thorough文体差別化はmodAskThorough側追記連結)/
②再取込ゼロの⚡仕上げ(modBackfill新設: 保存済みfull_text+breadcrumbからPhase1(AI0回)→
BuildOutlineFor委譲でPhase2/3。本棚ツールバー⚡+初回トースト)/③④設定導線(modSetupWizard
新設: 部門=商品部/リスコン部/自由入力・共有フォルダFolderPicker化・Hub/タイル/カードの
到達不能config案内全廃・初回2ステップウィザード(nexus_wizard_doneで1回性))/⑤引継ぎ
ファイル説明ダイアログ+パック表記統一/⑥MS&AD配色(msadテーマ既定: PRIMARY #01675B/
DARK #014D44/LIGHT #0B7D6E・自分バブル緑グラデ白字(6.8:1/5.0:1 AA)・☀🌙はmsad⇄dark)/
APP_NAME="MyBookshelf"統一(表示リテラル18箇所もHaiku波で置換)。
R19の骨子: ①余白根治=ScrollArea前提(R18)を廃し寸法で解く(modViewport:
FitBandToViewportのアフィン2点補正で吸収列を可視幅ぴったりへ・ContentRight一本化・
BoundAddr=内容下端+24pt・mChatBottom連動塗り。modSkinのA1:P2000全面塗り廃止)/
②取込完了時「型が一致しません」=Split直渡し3箇所(modSynonymStore)を型付き配列受けへ
+lint検査15(配列型引数へのSplit/Filter/Array直渡し検知。LO Basicは検査しない穴を塞ぐ)/
③「入念に調べる」バナー残留=nx_toastをClearChat+ClearProgressで掃除/
④短文×複数資料の分散シグナル聞き返し(DispersionGapX100+発動非発動両方でusage_log
"dispersion"・gap既定10・資料名明示なら発動せず)/⑤作業用Excel同居フリーズ根治=
Excelの既定インスタンスマージが原因。可視他ブック検知(CohabitOtherCount:
personal.xlsb/アドイン/不可視は除外)+取込前関所(スクショ・手動同期へも拡大)+
ランチャー『MyBookshelfを起動.bat』(excel.exe /x=別プロセス起動)同梱+docs00/10是正+
起動後案内はShowToast。UsedRange膨張判定は倍率→絶対値(横2,400/縦12,000pt)。
R19H中にAPI上限で実装エージェントが死亡→WIP保全コミット(ff1f5b8)→上限回復後に
レジュームで完走(§8の手順が実際に機能した)。
R17Hの骨子: 名寄せ辞書の永久0行バグ(添字)根治+MergeSynPairs純関数化/俯瞰回答の
二重警告排除(WasGlobalTurn+🔭バッジ・全モード入口リセット)/ArticleEnsureのスコープ遵守・
score=0化・0件時seed/同名章キー×複数資料の箱詰め修正/無人同期の章要約を12章超で先送り
(outline_deferred)/HasGlobalSignal(俯瞰語彙14語)で短文俯瞰を発火/chunk_metaセッション
キャッシュ/取込確認文言に章要約時間を明記/既存本棚への再取込案内トースト。
仕様: docs/dev/spec_20260805_R18_実機第5報.md + spec_20260805_R18H_レビュー裁定.md。
R18の骨子: ①バナー可変幅/2行化+ボタン常時前面(modProgressBar新設)+砂時計廃止+2段目
「作業用Excelを開く?」(セッション1回)/⑧manifest偽0根治(chunk_count -1保持+SourceList
実数突合自動修復+保存一本化+置換反転+modIntegrity起動時突合+Temp検知)/⑤非BMP絵文字の
ダイアログ排除+lint検査14+手動同期のinteractive配線/⑦HasCompoundSignal(？×2・。×2)で
短い複合質問も段0へ/E0303分類のusage_log格下げ/②全域書式の範囲限定+modViewport
(ScrollArea)+チャットへ上段移設/③ナレッジ地図の可視化撤去(分析CSV系は温存)/
④hubチップのカード化/⑨フッター+社内ポータル導線(E0905)。
R16の骨子(参考): ①文言統一/②作業用Excel+白画面抑止(初Declare)/③複合質問分解(modAskMulti)
・逆質問番号選択肢・精読(modAskFocus)・深掘り既出降格。
R16の骨子: ①文言統一(節約した時間)/②取込中のPC作業両立(🗔作業用Excel=WScript.Shellで
excel.exe /x・白画面抑止=リポジトリ初のDeclare(user32.DisableProcessWindowsGhosting、
config freeze_keep_banner 既定on)・案内文とdocs)/③入念モードの複合質問分解(modAskMulti、
段0判定→論点別軽量検索→統合→critique/verify→union出典突合。auto/最大3論点。質問全体で
LLM論点数+6回)・逆質問=番号選択肢(clarify、ok=False非回答契約・「1と3」複数選択・TTL30分)・
精読=近傍チャンク束ね(modAskFocus、source基準・thoroughのみ・deep_neighbor既定2)・
deep深掘りの既出チャンク降格(modFollowup、followup全検索に適用)。
次: 利用者の実機検証(下記)→実機第8報の受領。R20H Fix波+Haiku掃除波までクローズ済み。
R20H 記録のみ(次期): 部門ラベルのクリック配線(modUINexusDrawピル化・分割裁定とセット)/
Hub左カラム(統計タイル/バッジ帯)の幅追随/ACCENT #07A963の白字未達箇所(既存継承)/
darkテーマのヘッダー帯ブランド色固定/ToolbarContentRight孤児化の契約掃除/
breadcrumb診断の複数チャンク多数決化/MaybeShowBackfillToastの全走査の重さ(2万行級)/
OnTime予約のui_state二重化+mRefitRunning失効ガード/ブック名「'」のOnTime修飾(既存作法全体)/
「ナレッジ」⇔「資料」全面用語統一/docs10目次の全面再構成/modBackfill残りDictionaryの
比較方式統一/modAnalytics等のNexus残存(表示外)/診断画面の「チャンク数(my_knowledge)」は
技術者向けとして意図的に温存。
R19H 記録のみ(次期): 分散聞き返しの意図5区分の冗長(実機フィードバック後に省略検討)/
検査15の名前付き引数・単一行If(現リポ0件)/ViewportWidthの他ブックガード(modUIMain
凍結解除後)/ambiguous_dispersion_gap_x100=10の校正(usage_log "dispersion"の観測待ち)/
**cohabit_detectedの値の意味が「可視な他ブック数」へ変わった(次報のログ読み取り注意)**。
R17H 記録のみ(次期): 同義語追記が DistinctiveKeys 8枠を圧迫し得る(発火後に実測評価)/
名寄せの outline 非依存化(modShelf 凍結解除後)/既存本棚カードへの「再取込で新機能有効」
バッジ表示。Phase1(構造メタ+参照エッジ)は R17波0/波1、
Phase2(章単位要約=疑似グローバル検索)は R17波2、Phase3(enrich常時ON+用語名寄せ辞書)は
R17波3 で実装完了(いずれも下記「R17 で増えたもの」)。**modPrompts の分割は
Phase2/Phase3 とも見送った**(章要約・章選択・俯瞰回答・名寄せの4本とも新モジュール内の
Private プロンプトにしたため。調査agent7 §5-4 が前提にしていた分割は今回も不要だった。
modPromptsは残321字のまま=Phase3でも1文字も触っていない)。
残: 利用者の実機テスト(docs/45スモーク+R20分: 28=余白(全画面+窓リサイズ0.7s再フィット+
本棚gallery列数追随)・⚡仕上げ(旧形式資料の検出→実行→32-34機能有効化)・初回ウィザード
(2回目起動で出ないこと)・MS&AD配色・深掘りボタンの一般モード動作・回答フッターの
モード名表示(3モードの体感差)+usage_log "dispersion"(mode列が実モードに変更済み)/
"viewport"/"cohabit_detected"とask_stepsの回収→thorough_dispersion_gap_x100=25と
ambiguous_score_x100の校正)。
容量の分割必須ライン(次に触る波は先に分割裁定。R20H後の実測。上限30,000/WARN28,000):
**WARN帯12本(追記は分割裁定後のみ): modTestsPure2(29,542)/modTestsPure(28,906)/
modSkin(28,690・R20-7で急伸)/modTestsPure5(28,622)/modHub(28,620)/modKnowledge(28,566)/
modUIShelf(28,549)/modChannel(28,476)/modHubStat(28,431)/modUIMain(28,358)/
modUI(28,323)/modTestsPure12(28,280)**。
WARN線まで100字未満: optOcrPage(残3)/modBoot(残8)/modUtil(残8)/modTestsPure6(残9)/
modUINexusDraw(残13)/modTestsPure11(残19)/optVision(残26)/modChunker(残38)/
modAsk(残43)/modShelfStore(残44)/modShelfBatch(残45)/optGsTxt(残45)/
modShelfSync(残55・**分割裁定必須**)/modShelf(残65・**同**)/modRagParse(残90)。
新設・余裕: modBackfill(15,841)/modSetupWizard(4,924)/modTestsPure20(7,147)/
modTestsPure21(新)/modViewport(21,490)/modDashStat(21,409)/modClarify(24,655)。

| R | 実装者 | 内容 | コミット | 状態 |
|---|---|---|---|---|
| R20H2 | Haiku | "Nexus Agent"表示リテラル18箇所→APP_NAME参照+分析CSV名 | 62c10af | 完了 |
| R20-Fix | Sonnet | R20H裁定FA16(ビルド台帳3件/busy3情報源/仕上げロック/行高高水位分離/テーマトグル/管理者行RowX0/段数リーク/ウィザード1回性/全角エコー/部門ピッカー/APP_NAME/文言・docs是正) | 86a65cc〜2ec9a5f | 完了 |
| R20波D | Sonnet | ③④設定ウィザード(modSetupWizard/FolderPicker/導線全付替)+⑥MS&AD配色(msadテーマ)+Dashセンタリング | 74ba6a4〜f4f3cbb | 完了 |
| R20波C | Sonnet | ②⚡仕上げバックフィル(modBackfill)+⑤引継ぎ/パック説明 | 637b28c〜0fb0754 | 完了 |
| R20波B | Sonnet | ①深掘り一般対応+⑧モード可観測性・聞き返し強化 | 8ff2571〜2c43493 | 完了 |
| R20波A | Opus | ⑦余白3層根治(測定頑健化/縦掃除/幅追随/Hub縦詰め/リサイズ再フィット) | f8ec173〜13232dd | 完了 |
| R19-Fix | Opus | R19H裁定FA9(アフィン2点補正/shared下端/同居誤検知除外/HUB_BAND統一/塗り巻き戻し+絶対値判定/関所拡大/文言両対応/ShowToast/docs00)+FB10 | ff1f5b8/6daffd9 | 完了 |
| R19波C | Sonnet | ⑤同居検知+取込前関所+ランチャーbat生成/④分散シグナル聞き返し(HasScoreDispersion/config gap=10) | 9d2736e〜f457f85 | 完了 |
| R19波B | Opus | ①modViewport新設+全画面の余白根治(吸収列/ContentRight/BoundAddr/mChatBottom連動塗り) | 177c79b〜d4b37b8 | 完了 |
| R19波A | Opus | ②Split直渡し3箇所の型不一致根治+lint検査15(検知実証→修正→0件)/③nx_toast掃除(ClearChat+ClearProgress) | 7fdd2ba〜3d23d56 | 完了 |
| R17-Fix | Opus | R17H裁定FA10(名寄せ0行根治/俯瞰二重警告/スコープ遵守ほか)+FB9+FA-2補 | 28c7838/0bf73cd/0a1f7aa | 完了 |
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
| R18波A | Opus | ①⑪バナー根治(modProgressBar新設)+⑧永続化(modIntegrity新設・保存一本化・置換反転) | ea6b8ed〜cceb2d6 | 完了 |
| R17波0 | Sonnet | chunk_metaシート基盤(SH_CHUNK_META/EXPECTED_SHEETS/modChunkMetaStore)+modShelf圧縮 | 729a75f〜4746a61 | 完了 |
| R17波1 | Opus | Phase1: 構造メタ+参照エッジ(modChunkMeta新設/取込フック/RefsExpand/ArticleEnsure/別表・様式キー/graph_refs/テスト43件) | (本波) | 完了 |
| R17波2 | Opus | Phase2: 章単位要約=疑似グローバル検索(doc_outline/modOutlineStore/modOutlineBuild/modAskGlobal/verdict=global/graph_outline/テスト43件) | (本波) | 完了 |
| R17波3 | Sonnet | Phase3: enrich常時ON(EnrichPending既定maxCount30+modShelfSync小口呼び出し)+用語名寄せ辞書(synonyms/modSynonymStore/BuildSynonymsFor/name_dedup)+クエリ展開(ExpandQueryBySyn/ParseSynResp・modAskRetrieve入口)+graph_synonyms/テスト17件 | (本波) | 完了 |
| R18波B | Sonnet | ⑤非BMP排除+lint検査14+同期interactive+⑦HasCompoundSignal+E0303格下げ | 7d99933〜5e5b2cf | 完了 |
| R18波C | Opus | ②範囲限定+modViewport+チャットへ上段+③地図撤去+④カード化+⑨フッター | 5b1d786〜a41bcfa | 完了 |
| R18-Fix | Opus | R18H裁定FA8(2段目1回化/バナー2行/虚偽警告根治/全重複保護/フッター当たり判定ほか)+FB8 | 294afa4/f9664ef | 完了 |

### R17 で増えたもの(Phase1 分。次に触る人が最初に知るべき5点)

0. **新シート `chunk_meta`(chunk_id / section_path / refs_out)** = 規程の「章>条」と
   本文中の明示参照(「第8条による」「別表2のとおり」)の保存先。**列を足すのではなく
   新シートにした**のは、my_knowledge に列を増やすと配布済み本棚の移行が要るため
   (R17設計書§3)。ビルドが headers-only で焼き込み(`EXPECTED_SHEETS` + `_make_headers_only`)、
   実行時の `Worksheets.Add` は壊れたブックの自己修復専用。
1. **フェイルセーフが全ての前提**: chunk_meta が0行(=まだ取り込み直していない
   既存本棚)なら、参照展開も条番号の直接ヒット保証も**1行も足さずに戻る**。
   その条件は `modChunkMeta.GraphActive(metaCount, nHits)` の**1本だけ**にある。
   ここを2箇所へ書き写してはならない(片方だけ直すと古い本棚で実行時エラーになる)。
   移行処理は書かない=既存資料は「もう一度取り込む」ことでだけ構造が付く。
   これは設計書の明文の判断で、**版移行(modMigrate)のたびに効果を失う**という
   運用上の摩擦もそのまま受け入れている(調査agent7 §5-1)。
2. **新モジュール `modChunkMeta`(src/ingest・PURE_LOGIC_MODULES)** = 抽出の純ロジック。
   `ExtractSectionPath` は breadcrumb 行から、`ExtractRefs` は本文から InStr 走査だけで
   拾う。**VBScript.RegExp / ScriptControl は使わない**(政策ブロックのリスクと
   LO実行テスト不能=回帰を機械で固定できない)。全角/半角の吸収は
   `modSparse.NormalizeForSearch` に一本化してある(取込側と質問側で式が割れると
   「第１２条」が「第12条」で外れる)。
3. **section_path は ApplyCrumb を通す【前】の生チャンクから採る**
   (`modShelf.IngestFile` のループ内 `modChunkMeta.MetaOf` 1行)。保存後の
   my_knowledge から読んではいけない: config `embed_prefix_breadcrumb` が FALSE だと
   ApplyCrumb が breadcrumb 行を丸ごと消すため、admin 設定ひとつで構造が全滅する。
4. **旧メタ行の掃除は「消す前」に呼ぶ**。`modChunkMetaStore.RemoveMetaForSource` は
   my_knowledge を引いて chunk_id を集めるので、`RemoveKnowledgeAndVectorsForSource`
   の**後**に置くと何も引けない。呼び出しは `IngestFile` 手順7.5 と `DeleteSource` の
   各1行。chunk_meta 側の失敗は全て握って `usage_log("chunk_meta_fail")` 1行に留め、
   取込は止めない(検索精度の上積みであってデータ保全ではない)。
5. **検索合流は modAskFocus の2本だけ**: `RefsExpand`(精読束へ参照先を1ホップ・
   **同じ資料の中だけ**・既定8件・score=0)と `ArticleEnsure`(質問が名指しした条番号の
   チャンクが1件も無いときだけ先頭へ最大2件・score は先頭ヒット同値)。
   config `graph_refs`(既定on)のゲートもこの層に閉じてあるので、呼び出し元
   (modAskThorough / modAskMulti / modAskRetrieve)はどれも1行。**LLM呼び出しは
   1回も増えていない**。RefsExpand が資料を跨がないのは、跨いだ瞬間に無関係な規程の
   第8条が正しい出典タグ付きで根拠に混ざるため(利用者が気付けない外し方)。
   - modShelf は圧縮後 27,773字(残227)。次に触る波は先に分割を裁定すること。

### R17 Phase2 で増えたもの(俯瞰質問。Phase1 の上に積んである)

6. **新シート `doc_outline`(source / section_key / summary / keywords / chunk_n)**
   = 章ごとの要約。取込時に章の数だけ AI を呼んで作り(254頁で+3〜8分)、
   質問時は**要約だけ**を読んで「どの章を読むか」を選ぶ。GraphRAG のコミュニティ
   要約を「マニュアルが既に持つ章」で代替する、という R17設計書§3 Phase2 の骨子。
   ビルドが headers-only で焼き込む(chunk_meta と同じ)。
   - **`modBoot.HideInternalSheets` へは足していない**(modBoot は残8字で1行も
     入らない。憲章§4-6)。veryHidden はビルドの焼き込みと
     `modOutlineStore.EnsureOutlineSheet` の自己設定の**二重**で守っている。
     次に modBoot を触る波は、分割と同時にここへ `doc_outline` を足すこと。
7. **章キーは `section_path` の第1要素をそのまま**(`modOutlineBuild.ChapterKeyOf`)。
   追加の正規化を掛けてはいけない(section_path は取込時に
   `modSparse.NormalizeForSearch` を通っており、別の式を足すと**保存側と照合側で
   章キーが割れて、章を選べたのに本文が1件も引けない**=俯瞰が無音で死ぬ)。
   章見出しの無い資料では第1要素が条になる=条単位の要約になるが、それが正しい
   保守的動作(無い章立てを推測するより外れ方が小さい)。
8. **フェイルセーフは `modAskGlobal.OutlineActive` の1本**。doc_outline が0行
   (=取り込み直していない本棚)なら俯瞰は一切動かず、回答は R16 までと完全に同じ。
   章が選ばれない・章のチャンクが引けない・回答生成が失敗も**全て False** で
   従来の入念フローへ落ちる(2回呼んだ後の失敗でも同じ。確実に答えが出る側へ倒す)。
   不発の理由は `usage_log("global_zero" why=…)` に必ず1行残る。
9. **取込側の中断・保存は R15 の枠組みそのまま**: 章の境界で
   `modShelfBatch.CancelRequested`(と Err18)を見て、**そこまでの章を保存して
   正常終了**する。失敗章は `(要約失敗)` の行として保存して続行(行ごと落とすと
   「その章だけ要約が無い」ことが誰にも見えない)。`SaveCheckpoint(1, 120)`。
10. **プロンプト3本(章要約 / 章選択 / 俯瞰回答)は新モジュールの Private**。
   modPrompts(残321字)には置けないという容量裁定で、Phase2 では
   **modPrompts の分割はしていない**。出典タグの書式だけは
   `modPrompts.SourceTag`(唯一の持ち主)を必ず通すこと=ここを自前で書くと
   出典突合(`modAskThorough.CiteIndexFrom`)が全件不一致になる。
   段0の verdict に `global` を足したのは modPrompts / modRagParse / modAskMulti の
   各1〜3行だけで、知らない語は従来どおり single へ落ちる。

### R17 Phase3 で増えたもの(enrich常時ON+用語名寄せ辞書。Phase1/2 の上に積んである)

11. **enrich常時ON化(config `enrich_mode` 既定 off→light)**。取込・同期のたびに
    チャンクの要約・キーワードを少しずつ作る。254頁≒500チャンク規模を一括処理すると
    1時間級になる(設計書§3 Phase3)ため、**取込直後の同期実行では一気にやらない**。
    `modEnrich.EnrichPending` の `Optional maxCount` 既定値を `-1`(無制限)から
    `30` へ変えただけで小口化を実現した(呼び出し元 `modShelf.IngestFile` は
    既存どおり引数無しで呼ぶ=無改修)。`modShelfSync.SyncNow` にも
    `EnrichPending` の1行を新規追加し(既存 `EmbedPending` 直後)、同期のたびに
    少しずつ追いつく後追い巡回にした。**modShelf/modEnrich本体は無改修**
    (司令塔裁定どおり、凍結モジュールに1文字も足していない)。
12. **新シート `synonyms`(term, canonical)** = 用語の表記ゆれ辞書。
    「回収」⇔「リコール」のように意味が同じでも書き方が違う語を吸収する
    (docs/45 項目34)。chunk_meta/doc_outline と同じくビルドが headers-only で
    焼き込む。**`modBoot.HideInternalSheets` へは足していない**(残8字で
    1行も入らないため。次に modBoot を触る波が分割と同時に足すこと)。
13. **新モジュール `modSynonymStore`(src/ingest)** = synonymsのシートI/O
    (Ensure/Write/Read/RemoveAll)**と**名寄せバッチ(`BuildSynonymsFor`)の
    両方を1本に置いた。modChunkMetaStore/modOutlineStore(シートI/Oのみ)と
    modOutlineBuild(業務ロジック)を分けた前例とは違う構成だが、司令塔裁定で
    Phase3の新設モジュールをこの1本だけに絞った指示どおり。当該資料の
    chunk_meta(section_path)・doc_outline(keywords)・my_knowledge(keywords列)
    から用語候補を集め(重複排除・最大200語)、`CallLLM(step="name_dedup")`を
    資料1本の取込につき最大1回だけ呼ぶ。出力契約
    `<syn>表記>正規形|表記>正規形</syn>` は `modRagParse.ParseSynResp`。
    **既存termは上書き**: `ReadMapCsv`で読んだ既存分から新規termと重なる行を
    除き、`RemoveAll`してから全件を書き直す(1つの表に追記と上書きの2つの
    書き方を混在させない設計)。呼び出しは `modOutlineBuild.BuildOutlineFor`
    の末尾に1行(config `graph_synonyms`・失敗握り・`CancelRequested`確認も
    その1行の内側=呼び出し元は増えない)。
14. **クエリ展開は「質問文への同義語追記」方式**。`modRetrieve`/`modSparse`の
    スコアリング本体は無改修(司令塔裁定どおり)で、`modAskRetrieve.
    RunMultiRetrieve`の入口が質問文 `q` を書き換えてから既存の多段RAGへ渡す
    だけ。展開そのものは純関数 `modRagParse.ExpandQueryBySyn(q, mapCsv, maxAdd)`
    で、`modSparse.NormalizeForSearch`を両辺に通して全角/半角の表記ゆれも
    吸収し、双方向(表記→正規形・正規形→表記)・自己一致除外・最大3語・
    質問文に既にある語や追記済みの語の二重追記防止を1本の関数に閉じてある。
    `modSynonymStore.ReadMapCsv`は**1セッション1回だけ**呼び、モジュール変数
    (`mSynMapCsv`/`mSynLoaded`)へ控える。**取込・同期でsynonymsが更新されても、
    開いたままのセッションには次にブックを開き直すまで反映されない**
    (docs/10・docs/30に明記した既知のトレードオフ)。
15. **mock対応**: `modGateway.MockLLMResponse`に`Case "name_dedup"`を追加
    (`<syn></syn>`=0グループ。R16H FB-2の decompose/chapter_summary と同じ
    「Case Elseの汎用ダミーはタグを含まないためパーサが読めない」教訓の踏襲)。
    `enrich`のCaseは既存(R15以前)にあったため今回の追加は不要だった。

### R18 で増えたもの(次に触る人が最初に知るべき5点)

0. **新モジュール `modViewport`(src/ui)** = 各画面の「行ける範囲」の宣言
   (`Worksheet.ScrollArea`)。R18-3b で新設。呼び出しは Hub / マイ本棚(共通クロム
   `modKnowledge.DrawChrome` 1箇所で3モード分)/ チャット(`modUINexusDraw.
   DrawInputArea` 末尾。modUI に1行も入らないため)/ ダッシュボード(実測下端)/
   ナレッジ登録フォームの5箇所だけ。範囲文字列は画面ごとに1つの定数
   (`modHub.HUB_BOUND` / `modKnowledge.SHELF_BOUND` / `modUINexusDraw.NEXUS_BOUND`
   / `modVault.VAULT_BOUND`)で、**書式を当てる範囲と同じもの**を指す。
   ScrollArea を締めるときは必ず「実際に描いた最大到達点+余白」で決めること。
   きつく締めるとボタンが境界の外に取り残され、憲章§3-1違反(見えない=押せない)
   になる。**チャットだけは行方向を締めない**(バブルは行と無関係にpt座標で
   下へ伸び続けるため。A1:P2000)。

1. **新モジュール `modProgressBar`(src/ui)** = 進捗バナー(nx_progress)と
   ■中断 / 作業用Excel の描画・撤去。R18-1a で modSkin(残り10字)から忠実移設。
   呼び出し口は `modUIMain.ShowProgress/HideProgress`、
   `modShelfBatch.ShowIngestBanner`、`modHub.EnsureHubLayout`(SweepOrphans)の4本だけ。
   **バナーを最前面化したら必ず両ボタンを前面へ戻す**(R18-1c)。ここを外すと
   cancellable=False の実況が走るたびにボタンが不透明バナーの下へ埋まり、
   クリックがバナーに吸われて完全に無反応になる(実機第5報①の主犯)。
   幅は `BarWidthFor(modUIMain.ViewportWidth())`= min(760, viewport-16)。
   固定幅に戻してはならない(#30恒久対策の適用漏れが本文を殺した)。
2. **新モジュール `modIntegrity`(src/core)** = データ整合性の観測点。
   ・`ReconcileChunkCount`: カードの件数と my_knowledge の実行数の突合+自動修復
   (`modShelf.SourceList` から1行)。**カードは manifest の chunk_count を出す**
   という事実がここで初めて安全になった。
   ・`RecordSaveMark` / `WarnAtStartup`: 保存成功時に (行数, FullName) を ui_state へ
   控え、起動時に突合する。**zip直開き・一時展開コピーは保存も成功しReadOnlyでも
   ないため、これが唯一の検知手段**(調査agent1 §4「検知ゼロ」)。
3. **manifest の chunk_count は -1 で「前値保持」**(R18-2a)。実データを消して
   いない失敗経路が 0 を書くと、カードだけが「0件」に化ける。新しい失敗経路を
   足すときは必ず -1 を渡すこと。
4. **中間保存(SaveCheckpoint)の呼び出し点は modShelf.IngestFile の Finish 1箇所**
   (R18-2c)。取込の入口をいくつ増やしても保存が漏れない構造にしてある。
   AddFilesResult のループへ戻してはならない(二重保存になる)。

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
  ※R18完了時点: 121本(modProgressBar/modIntegrity/modViewport/modTestsPure16を追加)、
    LOモード1のPASSは1,486件、lint ERROR 0/WARN 4(同上)のまま。
    R18で記録した確定事項: 全重複再取込はfail_countを進め3回でfailed_permanent(復帰口
    あり・仕様)。ナレッジ地図の可視化コードは a41bcfa の親(75ba385^)から復活可能。
  ※R17完了時点: 129本(modChunkMeta/modChunkMetaStore/modOutlineStore/modOutlineBuild/
    modAskGlobal/modSynonymStore/modTestsPure17/modTestsPure18を追加)、
    LOモード1のPASSは1,626件、lint ERROR 0/WARN 4(同上)のまま。
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
- **R18で記録した次期課題(2026-08-05)**:
  - **ナレッジ地図の再導入案**(R18-4で可視化は撤去。調査agent3 §1-6):
    もし戻すなら「見るだけの絵」ではなく行動につながる形にすること。最小差分は
    (a) 各クラスタ円に `OnAction` を付け、代表ラベルで穴埋めテンプレを作って
    Nexus入力欄へ流し込む(`modHub.OnQuickAsk` と同型。新規ロジックほぼ不要)、
    (b) `ingest_files_total` や総チャンク数を鍵にした再計算キャッシュを入れて
    「開くたびにK-Means+最大約800回のCOM単発読み」をやめる、の2点。
    (b)を入れずに戻すと撤去理由の半分がそのまま戻る。削除したコード
    (DrawClusterMap/DrawBubbles/ClusterColor/MdsCoords/JacobiEigen/TopTwo/
    BuildLabels/TokenizeKw/NearestSource)は 81aa2d9 の親コミットに残っている。
  - `modCluster.LoadVectors` の my_vectors 側は今も1セルずつ読み(M-5の片割れ・
    最大約800回)。分析CSV経路だけになったので体感は消えたが、CSV出力は遅いまま。
    なお `kw()`(keywords列)の読みは R18H FB-3 で削除済み(消費者ゼロだった)。
- **R18H(敵対的レビュー裁定)で記録した次期課題(2026-08-05・FB-4 / B-M3)**:
  1. **本番ビルドに `Workbook_BeforeClose` が無い**(=「取込中は終了できない」
     ガードは**開発構成でしか効かない**)。事実は R18-1g で file:line まで裏取り
     済み(`build_mybookshelf.py:1002-1010` が `Workbook_Open` だけを書き込む /
     `Auto_Close` に `Cancel` 引数が無い)。実機で×・最小化が効かないのはこの
     ガードではなく `DisableProcessWindowsGhosting` の副作用。本番でも取込中の
     終了を守るには、インストーラ側で `BeforeClose` を注入するか
     `Application.OnKey` で退避する等の設計が要る。R18では**事実の記録のみ**。
     MASTER_SPEC §7.6 に「取込中終了禁止ガードは開発構成のみ」と明記済み。
  2. **`my_stats` のキーGC**(調査agent5)。`thx:` / `ins:` の nonce 行が単調
     増加する。最大25KB程度で実害は無いと裁定したが、`modStats.FindKeyRow` は
     線形探索なので行数が増えるほど遅くなる(HANDOFF「FindKeyRow裁定前提の
     更新」も参照)。GCの周期・保持日数は次期に決める。
  3. **deep / quick 向けの軽量 clarify**。論点分解と番号での聞き返しは入念
     モードだけの機能で、deep/quick では複合質問が1本のクエリのまま流れる。
     R18H FB-5 で「回答後に入念モードを1回案内するトースト」を入れたのは
     導線の応急処置であって、聞き返しそのものではない。LLM追加呼び出し1回に
     収まる軽量版(選択肢を出さず「どちらの論点ですか?」だけ聞く等)を次期に
     設計する。判定材料は `modRagParse.HasCompoundSignal` が既にある。
- R11での事実確認・修正メモ:
  - LibreOffice Private Const の参照不可: Public Const へ揃えて回避(modDashStatで実測)。
  - LogError context ラベル: Public エントリ名を指すこと(lintの参照チェックが文字列リテラル内も見る)。
  - modFeatures.InvokeFeature 引数: 最大6個。opt側の引数追加は末尾Optional固定。
  - Dir(vbDirectory): 実測24箇所(記録済み・次期統一)。
