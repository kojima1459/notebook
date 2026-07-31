# マイ本棚AI — マスター仕様書(実装契約書) v1.0

> **この文書が唯一の正**。実装エージェントは本書の契約(シグネチャ・シート列・エラーコード)から逸脱してはならない。
> 逸脱が必要だと判断した場合は、実装せずに最終レポートで逸脱提案として報告すること。
> 製品構想の背景は `/home/user/notebook/docs/v3/00_マイ本棚_構想と設計書.md` を参照(ただし本書と矛盾する場合は本書が勝つ)。

---

## 1. ゴールと完成の定義(DoD)

**プロダクト**: 「マイ本棚AI」— 承認なしで自分の資料をぶちこんで、すぐAIに質問できる社内版NotebookLM。
Excel単体(.xlsm)で動作し、社内AIリボン「リボンちゃん」の `ChatGPT()` / `GetEmbeddings()` に相乗りする。
**V2チャットボットとは別プロダクト**(コードは流用するがファイル参照はしない。`mybookshelf/` 配下で完全自己完結)。

DoD:
1. `mybookshelf/build/build_mybookshelf.py` で `mybookshelf/dist/MyBookshelf.xlsm` が本環境でビルドできる
2. `mybookshelf/tools/vba_lint.py` が全モジュールで違反0
3. LibreOffice headless で全モジュールがコンパイル成功+純ロジックテスト全緑
4. コア機能(取込・同期・2速回答・パック・ダッシュボード)は確認済み関数のみで完結
5. opt機能はビルドリスト1行削除+configフラグOFFで撤去でき、コアは無傷
6. 非エンジニア向け使い方/運用保守/撤去手順ドキュメント完備
7. 指定ブランチにコミット済み・mainと未統合・`git subtree split` で切り出し可能

## 2. スコープ

**入る**: ファイル取込(ダイアログ+本棚フォルダ差分同期)、抽出(txt/md/csv/pdf/docx/doc/xlsx)、チャンク分割、
埋め込み(再開可能バッチ)、任意のバッチ富化、2速QA(⚡すぐ聞く/🔍しっかり調べる)+出典表示、
ナレッジパック書き出し/取込(PIIスキャン+重複排除)、個人ダッシュボード+バッジ、環境診断、
opt: 画像読み取り(Vision・スクショ取込含む)/Markdown表示・Wordで開く/約款差分。
(読み上げ(TTS)は公式回答で「音声合成は非公開」と確定したためスコープ外。optTts.basはソース保管のみで
ビルド対象外=modules.jsonから撤去。確定台帳 RIBBON_API_CONFIRMED.md 裁定D4/D10)

**入らない**: V2公式ナレッジベースとの結合(公式ナレッジは将来「公式パック」として取り込む設計で代替)、
Graph API・外部HTTP(リボン以外の外部依存ゼロ)、リアルタイムファイル監視(ポーリング差分で代替)、UFフォーム(シートUIのみ)。

## 3. アーキテクチャと依存ルール

```
[UI層]      modUIMain / modUIShelf / modUIDashboard / modBoot
[機能層]    modShelf / modShelfSync / modAsk / modPack / modStats / modDiag / modEnrich
[部品層]    modExtractor(+Word/Excel/Acrobat) / modChunker / modEmbed / modRetrieve / modPrompts / modPii
[基盤層]    modAppDef / modConfig / modLog / modUtil / modGateway / modTypes / modFeatures
[opt層]     optTts / optVision / optMarkdown / optDiffDoc   ←— コアからの参照は modFeatures 経由の遅延バインドのみ
```

**依存ルール(lintで機械検査する)**:
- R1: 下層は上層を呼ばない(基盤→機能はNG)。例外(UI通知コールバック): 機能層→`modUIMain.SetStage`(進捗実況)、
  `modUIMain.RenderSourcesPreview`(出典先出し)、`modUIMain.RenderAnswer`(回答描画)、`modUIShelf.RenderShelf`(カード再描画)のみ許可。
  (Wave2 PM裁定: §7.2/§7.3が命じる呼び出しと整合させるため例外を4件に拡張。これ以外のUI参照は引き続き禁止)
- R2: `opt*` モジュール名の直接参照(`optTts.` 等のトークン)はコアのどのモジュールにも書かない。
  呼び出しは必ず `modFeatures.InvokeFeature` / `Application.Run` の文字列経由。
- R3: リボン関数(`ChatGPT`/`GetEmbeddings`/その他)の `Application.Run` は modGateway 内のみ。opt層も自前でRunせず `modGateway.TryRibbonRun` を使う。
- R4: **純ロジックモジュール**(modUtil / modChunker / modPii / modTypes と各Pureテスト)は
  Excelオブジェクト(`Worksheets`/`Range`/`Application`/`ThisWorkbook`/`MsgBox`)に触れてはならない(LibreOffice実行互換のため)。
- R5: エラーを握りつぶさない。`On Error Resume Next` は1行スコープ(直後に `On Error GoTo 0`)のみ。失敗は modLog に必ず残す。

## 4. データモデル(シート定義)

シート名は全て英小文字。タブ表示ユーザー向けシートのみ日本語。行1=ヘッダ固定。

| シート | 可視性 | 役割 |
|---|---|---|
| `使い方` | 可視(先頭) | マクロ有効化手順+クイックスタート。マクロ無効でも読める救済ページ |
| `ホーム` | 可視 | 質問UI |
| `マイ本棚` | 可視 | 本棚UI |
| `ダッシュボード` | 可視 | 個人統計+バッジ |
| `config` | hidden | 設定(§5) |
| `my_knowledge` | veryHidden | チャンク本体 |
| `my_vectors` | veryHidden | ベクトル |
| `my_manifest` | hidden | 同期台帳 |
| `my_stats` | hidden | 統計カウンタ+バッジ取得日 |
| `usage_log` | hidden | 利用ログ(1行=1質問/1操作) |
| `err_log` | hidden | エラーログ |
| `ui_state` | veryHidden | UI内部状態(モード選択等) |
| `vba_src` | veryHidden | 自己インストーラ用ソース格納(ビルド時生成) |

**my_knowledge** 列: `chunk_id, source, origin, page, summary, keywords, full_text, added_at, embedded`
- chunk_id 形式: `bs::<fnv64hex(full_text正規化後)>::p<page>::c<連番>`。fnvハッシュ部が重複排除キー。
- origin: `self` | `pack:<作成者名>`
- embedded: 0/1(1=my_vectorsに行がある)。再開可能バッチの走査キー。
- full_textは1セル32,000字を超えない(チャンカー保証)。

**my_vectors** 列: `chunk_id, vector_csv`(L2正規化済みDoubleのカンマ結合。次元はconfig `embed_dim`=1536)

**my_manifest** 列: `file_path, file_name, modified_at, size, chunk_count, status, error_note, ingested_at, origin`
- status: `done` | `pending` | `partial`(埋め込み未了あり) | `failed` | `image_pdf` | `missing`(同期でファイル消失検知→削除待ち)
- ダイアログ取込のfile_pathは実パス。パック由来はmanifestに載せない(my_knowledge.originで管理)。

**my_stats** 列: `key, value, updated_at`(key例: `ask_quick_total`, `ask_deep_total`, `selfsolve_total`, `hint_total`, `fail_total`, `ingest_files_total`, `pack_export_total`, `pack_import_total`, `streak_days`, `last_used_date`, `badge:<id>`=取得日)

**usage_log** 列: `timestamp, event, mode, detail, latency_ms, hit_count`
(event: `ask` / `feedback_green` / `feedback_yellow` / `feedback_red` / `ingest` / `sync` / `pack_export` / `pack_import` / `diag`)

**err_log** 列: `timestamp, code, context, detail, version`

## 5. config キー台帳(config シート A=key, B=value, C=説明(日本語))

| key | 既定値 | 意味 |
|---|---|---|
| mock_llm | TRUE(開発ビルド)/FALSE(本番) | リボン無しモック動作 |
| recommended_model | gpt-5.5 | 精査モードモデル |
| quick_model | gpt-5.5 | 即答モードモデル(将来 gpt-5.4-nano 等に差し替え可) |
| quick_effort / quick_verbosity | low / low | 即答モードの reasoning_effort / verbosity |
| deep_draft_effort / deep_draft_verbosity | medium / high | 精査ドラフト |
| deep_verify_effort / deep_verify_verbosity | high / medium | 精査検証 |
| reasoning_tuning | TRUE | FALSEでeffort/verbosity引数を空送信 |
| llm_wait_sec | 1200 | ChatGPT() Wait |
| topk_quick / topk_deep | 6 / 12 | LLMに渡す上位件数 |
| max_context_chars | 40000 | プロンプトに載せる本文合計上限 |
| answer_language | 日本語 | 回答言語(プロンプトに指定を挿入) |
| embed_dim | 1536 | ベクトル次元(パック互換性検査に使用) |
| embed_sleep_ms | 150 | 埋め込み呼び出し間スロットリング |
| shelf_max_chunks | 5000 | 本棚チャンク上限(超過時は取込拒否+整理案内) |
| shelf_folder | (空) | 本棚フォルダパス |
| sync_interval_min | 0 | OnTime自動同期間隔(0=off) |
| sync_on_open | TRUE | 起動時に差分同期 |
| enrich_mode | off | off/light/full: バッチ富化(§7.7) |
| max_pages_per_file | 300 | 抽出ページ上限(超過は打ち切り+partial) |
| ribbon_addin_name | リボンちゃん | AIリボンのアドイン検出名(RibbonAvailable用。裁定D2) |
| limit_check | TRUE | 起動時LimitCheck(期限・利用同意)。FALSEで無効化(裁定D3) |
| followup_max_pairs | 3 | 『続けて質問』で引き継ぐ履歴の最大ペア数。0以下で無効(裁定D11) |
| word_export_effort / word_export_verbosity | medium / medium | 『Wordで開く』の文書整形パラメータ(裁定D12) |
| feature_tts | FALSE | 読み上げ: 非公開確定のため提供不可・FALSE固定(裁定D4) |
| feature_vision / feature_markdown | TRUE | opt機能フラグ(公式仕様確定によりTRUE昇格。裁定D14) |
| feature_diffdoc | TRUE | 約款差分(確認済み関数のみ使用) |
| ghostscript_path | (空) | 画像PDFのOCRに使うgswin32c.exeのフルパス。空ならブックの隣の`Ghostscript\`を探す(R6) |
| vision_pdf_max_pages | 20 | 画像PDFを読み取る最大ページ数(1ページ=AI1回。超過は打ち切りpartial) |
| vision_pdf_dpi | 150 | 画像PDFのページ画像化の解像度(公式帳票OCR版と同値。300は約2倍重い) |
| vision_pdf_timeout_sec | 120 | ページ画像化の待ち時間上限(秒)。超過はその資料を失敗にして固まらせない |
| pack_author | (空:初回起動で入力) | パック作成者名 |
| debug_mode | FALSE | ゲートウェイのプロンプト/応答ログ |

## 6. エラーコード台帳(modLog)

形式 `E<2桁カテゴリ><2桁連番>`。**ユーザー向けメッセージは必ず「何が起きたか+どうすればよいか」の2文構成**。
コードは小さく末尾表示(保守者がerr_logと突合するため)。

| コード | 意味 | ユーザー向け対処文の要点 |
|---|---|---|
| E0101 | configシート欠損/必須キー欠落 | 配布元にこのファイルの再入手を依頼 |
| E0201 | リボン未検出(ChatGPT呼び出し失敗) | AIリボン入りExcelで開き直す。診断ボタンで確認 |
| E0202 | LLM応答がエラー/空 | 時間を置いて再実行。続くなら診断結果を管理者へ |
| E0203 | 埋め込み取得失敗 | 同上 |
| E0204 | 利用上限(LimitCheck)らしき拒否 | 今日はここまで。明日「同期」を押せば続きから再開 |
| E0301 | 対応外の拡張子 | 対応形式一覧を提示 |
| E0302 | ファイルが開けない(ロック/権限) | ファイルを閉じてから再実行 |
| E0303 | 画像PDF(文字が取れない) | このPDFは画像のため取込不可(Vision対応は将来) |
| E0304 | 抽出結果が空 | 中身が空/保護されている可能性 |
| E0401 | チャンク0件 | 文字数が少なすぎる資料 |
| E0501 | 本棚上限超過 | 使っていない資料を削除してから追加 |
| E0502 | 同期フォルダが見つからない | フォルダ選び直しを案内 |
| E0503 | 取込中の再実行(再入禁止) | 処理完了を待つ |
| E0504 | 同名別ファイルの衝突(別フォルダの同名資料が既に登録済み) | ファイル名を変えるか元の資料を削除してから追加(Wave4追加。§7.2 modShelf.IngestFile参照) |
| E0601 | 検索0件 | 資料が本棚にあるか確認を案内 |
| E0602 | 回答生成失敗 | 再実行案内 |
| E0701 | パック形式不正/バージョン不一致 | 相手に再書き出しを依頼 |
| E0702 | ベクトル次元不一致 | リボンの版が違う可能性。管理者へ |
| E0703 | PII検知(書き出し中断可) | 内容確認を促す(警告であり停止ではない) |
| E0801 | UI再構築失敗 | ブックを開き直す |
| E0901 | 診断で異常検知 | 診断レポートの指示に従う |

## 7. モジュール契約(公開API)

> 記載のない Public は作らない(Privateは自由)。引数・戻り値・失敗時挙動は厳守。
> 全モジュール: `Option Explicit` 必須 / `Attribute VB_Name` = ファイル名 / 1モジュール30,000字以内。

### 7.1 基盤層

**modAppDef.bas** — 定数のみ(ロジック禁止)
```vba
Public Const APP_NAME As String = "マイ本棚AI"
Public Const APP_VERSION As String = "0.1.0"   ' ビルド時にbuildスクリプトが検証表示
Public Const PACK_FORMAT_VERSION As Long = 1
Public Const SH_HOWTO = "使い方": Public Const SH_HOME = "ホーム"
Public Const SH_SHELF = "マイ本棚": Public Const SH_DASH = "ダッシュボード"
Public Const SH_CONFIG = "config": Public Const SH_KNOWLEDGE = "my_knowledge"
Public Const SH_VECTORS = "my_vectors": Public Const SH_MANIFEST = "my_manifest"
Public Const SH_STATS = "my_stats": Public Const SH_USAGE = "usage_log"
Public Const SH_ERRLOG = "err_log": Public Const SH_UISTATE = "ui_state"
' (実際は行を分けて Public Const … As String = … と正しいVBA構文で書くこと)
```

**modTypes.bas** — 型のみ
```vba
Public Type ExtractedPage
    page As Long
    Text As String
End Type
Public Type ShelfChunk
    chunk_id As String: source As String: origin As String
    page As Long: summary As String: keywords As String
    full_text As String
End Type
Public Type Hit
    chunk_id As String: score As Double: source As String
    page As Long: preview As String        ' 先頭120字(出典先出し表示用)
    origin As String
    full_text As String                    ' チャンク本文全体(回答生成の根拠。Wave3 PM裁定1)
End Type
```

**modConfig.bas** — V2 `src/chatbot_v2/modConfig.bas` を流用(SHEET_NAME="config")。
公開API: `EnsureLoaded` / `GetString` / `GetLong` / `GetDouble` / `GetBool` / `SetValue`(V2と同一シグネチャ)。

**modLog.bas**
```vba
Public Sub LogError(ByVal code As String, ByVal context As String, ByVal detail As String)
    ' err_logへ追記(失敗しても例外は出さない)。debug時はDebug.Printも。
Public Sub LogUsage(ByVal event_name As String, ByVal mode As String, ByVal detail As String, _
                    Optional ByVal latency_ms As Long = 0, Optional ByVal hit_count As Long = 0)
Public Function FriendlyMessage(ByVal code As String) As String  ' §6の対処文(2文構成)を返す
Public Sub ShowError(ByVal code As String, ByVal context As String, ByVal detail As String)
    ' LogError + MsgBox(FriendlyMessage & vbLf & "(コード: " & code & ")")
```

**modUtil.bas** — 純ロジック(R4遵守: Excelオブジェクト禁止)
```vba
Public Function Fnv1a64Hex(ByVal s As String) As String        ' 16桁hex。Currency/2xLong実装可。決定的であること
Public Function NormalizeForHash(ByVal s As String) As String  ' 空白圧縮+改行統一+Trim
Public Function VectorToCsv(vec() As Double) As String
Public Function CsvToVector(ByVal s As String, ByRef vec() As Double) As Boolean ' 不正時False
Public Function DotProduct(a() As Double, b() As Double) As Double ' 次元不一致は0を返しはしない→-2を返す規約は複雑なので: 次元不一致時は0返し+呼び出し側が次元検査
Public Function L2Normalize(ByRef vec() As Double) As Boolean  ' ゼロベクトルはFalse
Public Function HasVector(vec() As Double) As Boolean          ' 未初期化/空配列判定(On Error使用可)
Public Function SplitKeepNonEmpty(ByVal s As String, ByVal sep As String) As String()
Public Function HumanBytes(ByVal n As Double) As String        ' "1.2 MB"
Public Function HumanSeconds(ByVal sec As Double) As String    ' "約1分20秒"
Public Function SafeLeft(ByVal s As String, ByVal n As Long) As String
Public Function NowStamp() As String                            ' "yyyy-mm-dd hh:nn:ss"
Public Function FileNameOf(ByVal path As String) As String      ' 区切りは \ と / 両対応
Public Function ExtOf(ByVal path As String) As String           ' 小文字拡張子(ドットなし)
Public Function IsSameTimestamp(ByVal a As Date, ByVal b As Date) As Boolean ' 2秒丸め(FAT/OneDrive誤差)
Public Function EtaText(ByVal remainCount As Long, ByVal msPerItem As Double) As String
    ' R7 B-1: 残り件数×1件あたりミリ秒 → "残り約3分" / 60秒未満は "まもなく完了" /
    ' 見積り不能(msPerItem<=0)・残り0件は "" を返す(呼び出し側は件数だけ出す)
Public Function ProgressText(ByVal doneN As Long, ByVal totalN As Long, ByVal etaPart As String) As String
    ' R7 B-1: "12/88件・残り約3分"(etaPartが空なら "12/88件")
```
注: `NowStamp`はNow()を使う(Excel/LO両対応)。`FileNameOf`/`ExtOf`は文字列処理のみなのでR4適合。

**modGateway.bas** — リボン唯一の窓口。V2 `modRibbonGateway.bas` を基に拡張。
```vba
Public Function CallLLM(ByVal prompt As String, ByVal step_name As String, _
                        ByVal effort As String, ByVal verbosity As String, _
                        Optional ByVal model_override As String = "", _
                        Optional ByRef latency_ms As Long = 0, _
                        Optional ByVal prevU As String = "", _
                        Optional ByVal prevA As String = "") As String
    ' prevU/prevA: 会話継続用の履歴(確定台帳§1 #1 第7・8引数。新しい順;;;区切り。裁定D11)
    ' 成功: 応答文字列 / 失敗: "#ERR:E0202:..." で始まる文字列(例外は出さない)
    ' mock_llm=TRUE時: step_nameに応じた整形済みダミー(quick_draft/deep_draft/deep_verify/enrich/diff)
Public Function GetEmbedding(ByVal Text As String, Optional ByRef latency_ms As Long = 0) As Double()
    ' 成功: L2正規化済み配列 / 失敗: 空配列(modUtil.HasVector=False)。E0203をLogError
    ' mock時: テキストのFnvハッシュをシードにした決定的擬似ベクトル(L2正規化済み, embed_dim次元)
    '         → リボン無しでも取込→検索→回答の全画面フローが動く
Public Function RibbonAvailable() As Boolean   ' AddInsループでアドイン検出(裁定D2。結果をセッションキャッシュ)
Public Function RunLimitCheck() As Boolean     ' リボンLimitCheck()の唯一の呼び出し口(裁定D3。True=続行不可、エラー時False)
Public Function TryRibbonRun(ByVal funcName As String, ByVal args As Variant) As Variant
    ' Variant配列argsを展開してApplication.Run(要素数0〜6対応のSelect Case)。
    ' 失敗時: 文字列 "#ERR:E0202:<説明>" を返す。opt層はこれだけを使う
Public Function LooksLikeLimitError(ByVal response As String) As Boolean ' 上限/limit/回数の語を検知→E0204系判定
```
ChatGPT位置引数規約は確定済み(V2実証+確定台帳§0「互換後方追加」。裁定D1):
`Application.Run("ChatGPT", prompt, "", 0.4, 0, waitSec, model, prevU, prevA, "マイ本棚AI:" & step_name, effort, verbosity)`。
第9引数toolNは「マイ本棚AI:」ブランドを付けて管理側ログでツールを識別できるようにする。
GetEmbeddings の実引数規約は V2 `src/chatbot_v2/modEmbeddings.bas` の実装をそのまま踏襲すること。

**modFeatures.bas** — opt機能の分離実行
```vba
Public Function FeatureEnabled(ByVal featureId As String) As Boolean
    ' config "feature_<id>"=TRUE かつ ModulePresent
Public Function ModulePresent(ByVal featureId As String) As Boolean
    ' Application.Run "opt<Pascal>.Ping" をOn Errorで試す(結果キャッシュ)。opt側はPublic Function Ping() As Boolean を必ず持つ
Public Function InvokeFeature(ByVal featureId As String, ByVal procName As String, ByVal args As Variant) As Variant
    ' Application.Run("opt<Pascal>." & procName, ...) 遅延バインド。
    ' 無効/不在/失敗 → "#ERR:FEATURE_UNAVAILABLE" を返し、呼び出しUI側が
    ' 「この機能は現在利用できません(管理者が有効化すると使えます)」を表示
```
featureId→モジュール名対応表: tts→optTts, vision→optVision, markdown→optMarkdown, diffdoc→optDiffDoc(modAppDefではなくmodFeatures内Privateに保持)。

**modDiag.bas**
```vba
Public Sub RunDiagnostics()   ' 「🩺診断」ボタン。diag_reportシートを(再)生成して表示:
    ' バージョン/ビルド日、リボン検出、mock_llm状態、各シート存在+行数、config必須キー、
    ' 本棚統計(資料数/チャンク数/未埋め込み数)、opt機能の在否、直近エラー5件、
    ' 各行に ✅/⚠️ と対処文。最後に「この画面をスクリーンショットして管理者に送ってください」
Public Function QuickHealthCheck() As String  ' Boot時の軽量版: 問題なければ "" / あれば警告文
```

### 7.2 取込層

**modExtractor.bas / modExtractorWord.bas / modExtractorExcel.bas / modExtractorAcrobat.bas**
— `src/admin/` の同名モジュールをコピーし適応。公開契約:
```vba
' modExtractor
Public Function ExtractFile(ByVal path As String, ByRef pages() As ExtractedPage, _
                            ByRef errCode As String, ByRef errDetail As String) As Boolean
    ' 拡張子ディスパッチ。失敗時False+errCode(E0301/E0302/E0303/E0304)。例外は内部で捕捉
Public Function SupportedExts() As String   ' "txt,md,csv,pdf,docx,doc,xlsx,xls,xlsm"
```
適応要件: V1の`Err.Raise`スタイルはこの関数内で捕捉してBoolean+errCodeに変換。
画像PDF検知: 抽出結果の総文字数がページ数×10未満 → E0303。
`max_pages_per_file`超過は打ち切り、`errCode="PARTIAL_PAGES"`(Boolean=True扱い・呼び出し側がpartial記録)。

**modChunker.bas** — `src/admin/modChunker.bas` をコピー適応。純ロジック(R4)。
```vba
Public Function ChunkPages(pages() As ExtractedPage, ByVal targetChars As Long, _
                           ByVal overlapChars As Long, ByRef chunks() As ShelfChunk) As Long
    ' 戻り値=チャンク数(0可)。chunk_id/source/originは空のまま(modShelfが付番)。
    ' full_textは32000字を絶対に超えない。既定 targetChars=700 / overlap=150
```

**modEmbed.bas** — 再開可能バッチ埋め込み(V2 modVectorize のパターン移植)
```vba
Public Function EmbedPending(Optional ByVal maxCount As Long = -1) As Long
    ' my_knowledge の embedded=0 行を順に GetEmbedding → my_vectors 追記 → embedded=1
    ' 進捗を modUIMain.SetStage("📥 ベクトル化中 12/250 …") で実況。DoEvents。
    ' ESCで中断(Application.EnableCancelKey=xlErrorHandler, Err=18捕捉)→ そこまで保存し件数返す
    ' 3連続失敗 or LooksLikeLimitError → 中断し E0204 案内(「明日🔄同期で続きから」)
    ' 戻り値: 今回埋め込んだ件数。呼び出し間に embed_sleep_ms スリープ
Public Function PendingCount() As Long
```

**modShelf.bas** — 本棚中核
```vba
Public Sub AddFilesViaDialog()      ' 複数選択FileDialog(フィルタ=SupportedExts)→ 各IngestFile
Public Function IngestFile(ByVal path As String, ByVal origin As String) As String
    ' 戻り値=manifest status("done"/"partial"/"failed"/"image_pdf")
    ' 手順: 再入guard(E0503) → 同名衝突検査(E0504。Wave4追加: source識別がファイル名の
    '       みであるため、manifestに「同じfile_nameだが別のfile_path」の行が既にあれば、
    '       削除せず明示エラーで止める。誠実な失敗の方が黙ったデータ消失より安全という
    '       R5の判断。同一パスの再取込=置換はこの検査を通過する)
    '       → 上限検査(E0501) → 既存同名sourceは置換(古いchunk/vector削除)
    '       → ExtractFile → ChunkPages → chunk_id付番(bs::hash::pN::cN, ハッシュ重複はスキップ)
    '       → my_knowledge追記(embedded=0) → manifest upsert(status=pending)
    '       → EmbedPending → manifest status確定 → 成功時(done/partial)のみ
    '       my_stats.ingest_files_totalをBump+usage_logに"ingest"イベントを記録
    '       → カード再描画(modUIShelf.RenderShelf)
Public Sub DeleteSource(ByVal sourceName As String)  ' knowledge/vectors/manifest から一括削除+再描画
Public Function SourceList(ByRef names() As String, ByRef stats() As String) As Long ' カード描画用
Public Function TotalChunks() As Long
Public Function IsBusy() As Boolean
    ' R7 B-2: 再入ガード mIngesting の公開。抽出ループのDoEventsで発火した
    ' クリックを画面遷移側(modUiLock.BlockIfIngesting)が受け流すために使う。
    ' 焼き付き対策としてGUARD_EXPIRY_MIN超過のガードはFalseを返す。
```
注(R7 B-2): 再入guard(E0503)は ShowError(モーダル)ではなく
`modLog.LogError` + `modUIMain.SetStage` のみで通知する(連打のたびに
ダイアログが積み上がると、それ自体が「応答なし」に見えるため)。
E0503 のログ記録は従来どおり残す。

**modShelfSync.bas**
```vba
Public Sub PickShelfFolder()      ' フォルダ選択→config shelf_folder 保存→即SyncNow
Public Sub SyncNow(Optional ByVal silent As Boolean = False)
    ' Dir()走査(サブフォルダ不要・第1階層のみ) vs my_manifest:
    '   新規→IngestFile / modified_at or size 変化(IsSameTimestamp)→置換取込 /
    '   消失→DeleteSource / pending・partial→EmbedPending再開
    ' フォルダ不存在→E0502。実行結果サマリをステータス表示+LogUsage("sync")
    ' silent(既定False。Wave4追加): TrueのときはE0502警告・完了サマリをMsgBoxで出さず
    ' SetStage(状態表示行+StatusBar)のみに留める。modBoot.Boot(sync_on_open)と
    ' AutoSyncTick(定期自動同期)はユーザー操作なしで走るバックグラウンド処理のため
    ' silent:=Trueで呼ぶ(対話ダイアログでフォーカスを奪わない。§8 UXレビュー対応:
    ' 配布直後にshelf_folder未設定のままE0502警告が必ず出る問題/定期MsgBoxが
    ' 作業を中断する問題への対応)。手動🔄ボタン・PickShelfFolder直後はsilent=False。
Public Function IsBusy() As Boolean
    ' R7 B-2: 再入ガード mSyncRunning の公開(modShelf.IsBusyと対)。
    ' 同期のファイルループも1ファイルごとにDoEventsを回すため必要。
Public Sub ScheduleAutoSync()     ' sync_interval_min>0ならApplication.OnTimeで次回予約(自己再帰)
Public Sub CancelAutoSync()       ' 予約解除(必ずOn Error握り: 予約なしでも安全)
Public Sub AutoSyncTick()         ' OnTimeコールバック本体(Wave4修正: 契約上Public必須。下記注参照)
Public Function ResolveDecision(ByVal decision As String, ByVal existsInManifest As Boolean, _
                                ByVal currentStatus As String) As String
    ' 純関数(§7.8): DiffDecisionの結果にmanifestの現在statusによる上書きルールを適用する。
    ' decision="keep"かつstatus="failed"/"missing" → "replace"(前回失敗の再試行/
    ' フォルダ復活時の復帰。Wave4修正: missing行が復活後もkeep固定のまま戻らない不具合の修正)
' 注意: 予約解除は modBoot.Auto_Close から必ず呼ぶ(OnTime残存→勝手にExcelが再起動する事故防止)
' 注意(Wave4修正・契約自体の欠陥だったための修正): Application.OnTimeのProcedure引数は
'   Application.Runと同じ遅延バインド(文字列からの実行時解決)であり、Private Subは
'   解決できない。予約自体は成立してしまうが、発火時に「マクロを実行できません」という
'   未処理エラーになり自動同期が機能しなくなる。そのためAutoSyncTickは
'   「OnTimeコールバックとして公開が必須」の契約とし、Public Subとして実装すること。
```

**modEnrich.bas** — バッチ富化(summary/keywords付与)
```vba
Public Function EnrichPending(Optional ByVal maxCount As Long = -1) As Long
    ' enrich_mode=off なら即0。summary空のチャンクを10件/1回のCallLLM(JSON配列返し)で富化。
    ' JSONパースは軽量自前(期待形: [{"i":1,"summary":"…","keywords":"a,b"},…])。
    ' パース失敗はそのバッチをスキップして続行(富化は無くても検索は動く=非致命)
```

### 7.3 QA層

**modRetrieve.bas**
```vba
Public Function Search(ByVal query As String, ByVal topK As Long, ByRef hits() As Hit) As Long
    ' GetEmbedding(query) → my_vectors全件Dot積(全てL2済みなのでコサイン同値)
    ' + キーワードボーナス: queryの語がsummary/keywords/sourceに含まれれば +0.05/語(上限0.15)
    ' 戻り値=件数(0可)。埋め込み失敗時は-1(呼び出し側がE0203表示)
    ' hits().full_text にはmy_knowledgeのfull_text列(チャンク本文全体)をそのまま格納する
    ' (Wave3 PM裁定1: previewだけでは回答生成の根拠として不十分なため)。
    ' hits().preview は引き続き先頭120字(出典先出し表示用)。
```

**modPrompts.bas** — プロンプト組み立て(純文字列。answer_language挿入)
```vba
Public Function BuildQuickPrompt(ByVal q As String, hits() As Hit, ByVal nHits As Long) As String
Public Function BuildDeepDraftPrompt(ByVal q As String, hits() As Hit, ByVal nHits As Long, ByVal history As String) As String
Public Function BuildDeepVerifyPrompt(ByVal q As String, ByVal draft As String, hits() As Hit, ByVal nHits As Long) As String
Public Function BuildEnrichPrompt(ByVal batchText As String) As String
```
出典指示: 回答内の引用は `[本棚:ファイル名 p.页]` 形式で、と明示。パック由来は `[パック(作成者):ファイル名]`。
「資料に無いことは『資料には見当たらない』と述べ、推測は推測と明示」を全モードで指示。
本棚抜粋ブロックの本文には `hits().full_text`(チャンク本文全体)を使う(Wave3 PM裁定1。
previewは出典先出し表示専用でプロンプト本文には使わない)。
本文組み込みは `max_context_chars` で打ち切り(切ったら「(一部省略)」を挿入)。

**modAsk.bas**
```vba
Public Sub AskFromUI()   ' ホームの質問セル+モード(ui_state)を読み、Answer実行→RenderAnswer
Public Function Answer(ByVal question As String, ByVal mode As String) As String
    ' mode "quick": SetStage(🔍検索中)→Search(topk_quick)→出典先出し表示→SetStage(✍️回答作成中)
    '               →CallLLM(quick)→回答
    ' mode "deep":  Search(topk_deep)→出典先出し→draft→SetStage(✅検証中)→verify→回答
    ' 検索0件: LLMを呼ばず「本棚に手がかりが見つからない」定型文+資料追加の案内(E0601はログのみ)
    ' ESC対応・所要秒を回答末尾に小さく表示。LogUsage("ask")
Public Sub FeedbackGreen() / FeedbackYellow() / FeedbackRed()  ' modStatsへ+お礼表示
Public Function CanFollowup() As Boolean       ' 続けて質問できる直近回答があるか(裁定D11)
Public Sub AskFollowup(ByVal followupText As String) ' 履歴付き追質問(検索も再実行→出典付き回答)
```
会話履歴(裁定D11): 直近 followup_max_pairs(既定3)往復を module 変数に保持し、CallLLM の
prevU/prevA(新しい順;;;区切り)へ渡す。履歴に積むのは深掘り候補ブロック除去後の本文のみ。
深掘り候補: modPrompts が最終応答系プロンプト(quick/deep_verify)に [[FOLLOWUP: 候補1 | 候補2]] の
出力指示を加え、modAsk がパース・除去して「🔎 深掘り候補(『続けて質問』でそのまま聞けます)」
ブロックを本文末尾に整形追記する(RenderAnswer契約は不変。マーカー無し応答でも壊れない)。

### 7.4 パック層

**modPii.bas** — `src/chatbot_v2/modPii.bas` コピー(純ロジックR4)。
`Public Function ScanText(ByVal s As String) As String`(検知種別のカンマ列挙 or 空)。

**modPack.bas**
```vba
Public Sub ExportPackDialog()
    ' 本棚全体 or アクティブ行の資料1件を選択(vbYesNoCancel)→保存先ダイアログ
    ' → 新規Workbook(.xlsx マクロ無し)に pack_meta/pack_chunks/pack_vectors を書いて保存
    ' → 書き出し前に modPii.ScanText を全チャンクに実行、検知時は件数+例を出して続行確認(E0703)
Public Sub ImportPackDialog()
    ' ファイル選択→ValidatePack→重複(fnvハッシュ)スキップしつつmy_knowledge/my_vectorsへ
    ' origin="pack:"&作成者。取込サマリ表示(「32件取込 / 8件は既にありました」)
Public Function ValidatePack(ByVal wb As Workbook, ByRef reason As String) As Boolean
    ' 3シート存在・pack_format_version一致・embed_dim一致(E0701/E0702)
```
**pack_meta** 列(key,value): pack_format_version, pack_name, author, created_at, description, embed_dim, app_version, chunk_count。
**pack_chunks** 列: chunk_id, source, page, summary, keywords, full_text。**pack_vectors** 列: chunk_id, vector_csv。

### 7.5 統計層

**modStats.bas**
```vba
Public Sub Bump(ByVal key As String, Optional ByVal delta As Long = 1)  ' my_stats upsert
Public Function GetStat(ByVal key As String) As Long
Public Sub TouchToday()            ' streak_days/last_used_date 更新(連続利用日数)
Public Sub EvaluateBadges()        ' §9のバッジ条件を検査、新規獲得は badge:<id> に日付+お祝いMsgBox(1回だけ)
Public Function SavedMinutesEstimate() As Long ' 自己解決1件=15分換算(config化不要・定数コメントで明示)
```

### 7.6 UI層(§8のレイアウト仕様に従う)

**modUIMain.bas**
```vba
Public Sub EnsureLayout()          ' ホームを冪等再構築(既存Shapes全削除→再生成)
Public Sub SetStage(ByVal msg As String)  ' ステータス行+Application.StatusBar 両方
Public Sub RenderAnswer(ByVal answerText As String, hits() As Hit, ByVal nHits As Long, ByVal mode As String, ByVal seconds As Long)
    ' mode=""(Wave4追加の目印): 空質問等で検索・回答生成を一切行わなかったことを示す。
    ' このときは所要秒・出典欄(前回質問の情報に見えてしまう)を付けず案内文だけを表示する。
Public Sub RenderSourcesPreview(hits() As Hit, ByVal nHits As Long)  ' 出典先出し(ドラフト生成前に呼ぶ)
Public Sub OnAskButton() / OnModeQuick() / OnModeDeep() / OnOpenHowto() / OnRunDiag()
Public Sub ShowTip()               ' 待ち時間豆知識(定型10本からRnd選択)
```

**modUIShelf.bas**
```vba
Public Sub EnsureLayout()
Public Sub RenderShelf()           ' manifest+SourceListから資料カード行を再描画(§8.2)
Public Sub OnAddFiles() / OnSyncNow() / OnPickFolder() / OnExportPack() / OnImportPack() / OnDeleteSource()
' OnDeleteSource: アクティブ行のsourceを特定(カード行範囲外なら案内)→「『x』を本棚から削除しますか?」確認
```

**modUIDashboard.bas**
```vba
Public Sub EnsureLayout()
Public Sub RenderDashboard()       ' 統計タイル+バッジ棚+REPT("■")棒グラフ。EvaluateBadges後に呼ばれる
```

**modBoot.bas + ThisWorkbook.cls**
```vba
Public Sub Boot()      ' V2パターン踏襲: config確認→first-run(pack_author入力)→3画面EnsureLayout
                       ' →QuickHealthCheck表示→AIリボン利用期限確認(modGateway.RunLimitCheck。
                       '   True=制限中でもvbInformation案内のみで起動は止めない。裁定D3)
                       ' →sync_on_openならSyncNow→ScheduleAutoSync→内部シート隠蔽
Public Sub Auto_Open() ' Boot呼び(ガード付き)
Public Sub Auto_Close()' CancelAutoSync(必須!)+Application.StatusBar=False
```
ThisWorkbook.cls は Workbook_Open→Boot / Workbook_BeforeClose→Auto_Close の薄い転送のみ
(ビルド版ではインストーラがThisWorkbookを占有するため、実行時はAuto_Open/Auto_Closeが本命。両方書く=V2実証済み二重化)。

### 7.7 opt層(全モジュール共通契約)

- 必ず `Public Function Ping() As Boolean`(True返すだけ)を持つ
- リボン呼び出しは `modGateway.TryRibbonRun` のみ。**リボン関数のシグネチャは確定台帳
  RIBBON_API_CONFIRMED.md §1 が唯一の根拠**(公開仕様の入手により、当初のSIGNATURE ASSUMPTION
  ブロックは「確定済み」の出典記載に置き換え済み)
- コアモジュールへの参照は基盤層+modUIMain.SetStageのみ可
- 失敗しても例外を外に出さない(文字列 "#ERR:..." 返し)

```vba
' optTts.bas    : ビルド対象外(音声合成は非公開と確定。裁定D4/D10。ソースのみ保管)
' optVision.bas : Public Function ExtractImagePdf(ByVal path As String, ByRef pages() As ExtractedPage) As Boolean
'                 ' 確定経路: Base64FromFile(path)→ChatGPTV(prompt, b64, "", "high", "マイ本棚AI:vision")(裁定D5)
'                 ' 対応形式はpng/jpg/jpegのみ(PDF直渡しは公式仕様上不可能と確定し撤去)
'                 Public Function ExtractImagePdfText(ByVal path As String) As String ' "#ERR:.." or 全文
'                 ' modShelfは E0303(画像PDF) または E0301+画像拡張子 のとき
'                 ' feature_vision有効なら InvokeFeature("vision","ExtractImagePdfText",path) を試す(裁定D13)
'                 Public Function HasClipboardImage() As Boolean   ' IsImageInCB確定関数(裁定D13)
'                 Public Function SaveClipboardImage() As String   ' Base64FromCB(Ptn=1)→Temp jpgパス/"#ERR:..."
' optMarkdown.bas: Public Function RenderMarkdownAt(ByVal sheetName As String, ByVal cellAddr As String, ByVal md As String) As String
'                 ' 確定方式: セルにSafeLeft(md,32000)を書き込んでから CellMarkDown(rng, False)(裁定D6)
'                 Public Function OpenAnswerInWord(ByVal md As String) As String  ' OpenWordMark確定関数(裁定D6)
'                 Public Function ExportAnswerAsDoc(ByVal answerText As String, ByVal instruction As String) As String
'                 ' 対話型文書生成(裁定D12): 指示文ありならCallLLM(step="word_export")で整形→OpenWordMark。
'                 ' 指示文空ならOpenAnswerInWordへ直行(そのまま転記)
' optDiffDoc.bas : Public Sub CompareTwoDocsDialog()  ' 新旧2ファイル→抽出→ChatGPT差分分析→diff_reportシート
'                  (確認済み関数のみ使用だがサブ機能なのでopt隔離)
```
UI側: opt機能のボタンは `FeatureEnabled` がTrueのときだけ生成(EnsureLayout内で分岐)。
ボタンOnActionは modUIMain/modUIShelf 内の `OnOpenWordButton`/`OnIngestScreenshot` 等の
ラッパー(そこから InvokeFeature)。ホーム31:32行は btn_followup「💬 続けて質問」(A31:D32・常時)+
btn_word「📝 Wordで開く」(E31:H32・feature_markdown時のみ)。本棚K1:L2は btn_screenshot
「📸 スクショ取込」(feature_vision時のみ)。読み上げボタンは存在しない(裁定D10)。

### 7.8 テストモジュール

**modTestRunner.bas**(R4準拠・ExcelでもLOでも動く中核)
```vba
Public Sub ResetTests()
Public Sub Check(ByVal name As String, ByVal cond As Boolean, Optional ByVal detail As String = "")
Public Function Failures() As Long
Public Function ReportText() As String     ' "PASS 42 / FAIL 0" + 失敗一覧
Public Sub RunAllPureTests()               ' 各modTests*のRun~を列挙呼び出し
```
**modTestsPure.bas**: modUtil(ハッシュ決定性・CSV往復・L2/Dot・タイムスタンプ丸め)、
modChunker(空/1文/長文/32000字上限/オーバーラップ検証/段落境界)、modPii、
modPrompts(出典形式・打ち切り)、manifest差分ロジック(modShelfSyncの差分判定をPrivateではなく
`Public Function DiffDecision(ByVal existsInManifest As Boolean, ByVal sizeChanged As Boolean, ByVal timeChanged As Boolean) As String`
のような純関数に切り出してテスト可能にする)、pack検証(ValidatePackの列検査部を純関数に切り出し)。
**Excel側のみ**: modTestsExcel.bas(シートI/O・ゲートウェイmock経由のE2E: mock取込→検索→回答)。診断から実行可能に。

## 8. UI仕様(非エンジニア向け・iPhone基準)

共通原則: ボタンは角丸四角Shape+OnAction、高さ2行以上・ラベルは動詞(「質問する」)。
選択状態は色で明示(選択=濃紺白字 / 非選択=薄灰黒字)。専門用語禁止(「ベクトル化」→「AIが読める形に変換」)。
全ての失敗メッセージは§6の2文構成。フォント: 游ゴシック11pt基準・見出し14pt太字。
列幅・行高はEnsureLayoutが設定(ユーザーが壊しても開き直せば直る=冪等)。

### 8.1 ホーム
```
[📚 マイ本棚AI  v0.1.0]                      [❓使い方] [🩺診断]
「自分で入れた資料に、すぐ聞ける」
┌───────────────┐ ┌───────────────┐
│ ⚡ すぐ聞く (10〜20秒) │ │ 🔍 しっかり調べる (1〜2分) │   ← トグル
└───────────────┘ └───────────────┘
質問をここに入力してください(例: 〇〇の手続きに必要な書類は?)
[                    大きな入力セル(結合・折返し)                    ]
              [ 💬  質 問 す る ]
状態: 🔍 検索中… / 📄 3件の資料がヒット / ✍️ 回答作成中…
─ 回答 ─────────────────────────
(回答本文: 結合セル・折返し・上詰め)
📖 この回答のもと: ファイルA p.3 / ファイルB p.12
この回答は役に立ちましたか?  [🟢 解決した!] [🟡 ヒントになった] [🔴 だめだった]
💡 豆知識: (待ち時間に表示)
```
- 本棚が空のとき: 回答エリアに「まず『マイ本棚』タブで資料を1つ追加してみましょう →」を常設表示。

### 8.2 マイ本棚
```
[＋ 資料を追加] [🔄 フォルダと同期] [📦 パックにして渡す] [📥 パックを取り込む] [🗑 選んだ資料を削除]
本棚フォルダ: C:\Users\...\OneDrive\マイ本棚  [📁 フォルダを選ぶ]  自動同期: 30分ごと/オフ
ここにファイルを入れておくと、自動で本棚に追加されます(消せば本棚からも消えます)
──資料カード(1行=1資料)──────────────────
 状態   資料名                    追加日     ページ数→チャンク数   メモ
 ✅     医療保険_引受ガイド.pdf   7/11      98p → 245件
 ⏳     新商品QA.docx            7/11      変換中… (34/120)
 ⚠️     約款_旧版.pdf            7/10      取込失敗: ファイルが開けません(E0302)
 🖼     スキャン申込書.pdf        7/10      画像PDFのため取込不可
```
- 削除はアクティブセルの行の資料を対象(ボタン押下時にファイル名を確認ダイアログで明示)。

### 8.3 ダッシュボード
統計タイル(今月の質問数/🟢自己解決数/取り戻した時間≒解決×15分/本棚の冊数)+
バッジ棚(獲得=カラー絵文字+日付、未獲得=灰色+条件文)+蔵書数の推移棒(REPT)。

## 9. バッジ定義
`first_ingest`初めての取込 / `shelf10`本棚10冊 / `shelf30`本棚30冊 / `first_pack_out`初パック共有 /
`first_pack_in`初パック取込 / `solve10`自己解決10件 / `solve50`自己解決50件 / `streak7`7日連続利用。

## 10. ビルド仕様(build_mybookshelf.py)

- `build/build_chatbot_v2.py` の実証機構(template_skeleton.xlsm + openpyxlシート生成 + vbaProject.bin の
  ThisWorkbook/dirストリーム外科パッチ + vba_src自己インストーラ)を **mybookshelf/build/ にコピーして独立させる**
  (make_xlsm.pyの圧縮関数含む。元ファイルは変更しない)。template_skeleton.xlsm もコピー。
- インストーラは `modBoot.Boot` を起動(cp932エンコード制約により installer 内文字列はASCIIのみ)。
- モジュールリストはマニフェスト(`build/modules.json`)で管理: `{"name","path","role":"core|opt|test"}`。
  **opt機能の撤去=このJSONから1行削除+configフラグFALSEでビルド**、を保証。
- `--dev`(mock_llm=TRUE, テストモジュール同梱)/`--prod`(mock_llm=FALSE, テスト同梱は維持=診断用)。
- シート生成: §4の全シート+ヘッダ+config初期値(C列に日本語説明)+使い方シート本文。
- ビルド後自己検証(同スクリプト内): 再オープンして全シート存在・vba_srcモジュール数一致・
  各ソースセル≤32,000字・olefileでThisWorkbookストリーム復元確認。検証失敗はexit 1。

## 11. テスト計画(3層)

1. **静的Lint(`tools/vba_lint.py`)** — 全.bas/.clsに対し:
   Option Explicit / VB_Name=ファイル名 / モジュール≤30,000字 /
   Public宣言が§7契約と一致(過不足検出: 契約表はlint内に埋め込む) /
   クロスモジュール参照 `modX.Proc` の実在 / R1〜R4依存ルール / `Application.Run`第1引数のホワイトリスト /
   `Dim a, b As Long` 型落ち警告 / `Integer`警告 / セル書込値に32,767字超リスクのSafeLeft無し警告(full_text系) /
   opt直接参照検出。**Exit code非0=違反あり**。
2. **LibreOffice実行(`tools/run_lo_tests.py`)** — ユーザープロファイルのStandardライブラリに
   R4準拠モジュール+modTestRunner+modTestsPureを.xba化して注入し、
   `soffice --headless "vnd.sun.star.script:...location=application"` でRunAllPureTests実行→
   結果を/tmpファイル経由で回収しexit code化。**さらに全モジュール(Excel依存含む)を構文コンパイルのみ**
   別ライブラリでロードして構文エラー検出(実行はしない)。Attribute行はローダが除去。
   VBA固有でLOが解釈できない構文が出た場合は、lint側で当該構文の代替を規約化する(勝手にテスト対象から外さない)。
3. **実機受入(`docs/40_受入チェックリスト15分.md`)** — 配布先PCで: 診断→mockでE2E→mock解除→
   リボン実呼び出し1回→txt1件取込→質問→パック書き出し/取込→終了時OnTime残存なし確認。

## 12. コーディング規約(VBA落とし穴・全員厳守)

- 文字列連結の長大化は `&` 連鎖でなく配列+Join(パフォーマンス)。ループ内Range直接アクセス禁止→配列一括読み書き。
- `Integer`禁止(Long)。`Dim a, b As Long`禁止(1行1変数か全てに型)。
- 日付比較は `IsSameTimestamp`(OneDrive/FATの2秒丸め)。
- セル書込は `SafeLeft(s, 32000)` を経由(full_text/回答/ログ)。
- 絵文字はセル値・Shape文字列では可、**MsgBox/InputBox文字列では使わない**(豆腐化)。VBA文字列リテラルには
  BMP外絵文字(🟢等)を書いてよい(V2実証済み)が、**ビルドスクリプトがUTF-8→セル格納**の経路のみ(cp932直埋め禁止)。
- `On Error Resume Next` の乱用禁止(R5)。エラーハンドラでは必ず `Err.Clear`/`On Error GoTo 0` を明示。
- `ActiveWorkbook`禁止→`ThisWorkbook`。`Select/Activate`はUI表示目的以外禁止。
- モジュール間の循環参照禁止(R1)。
- Shapes: 名前は `btn_`/`lbl_` プレフィクス+機械名。EnsureLayoutは既存の自前Shapesを全削除してから再生成(冪等)。
- OnTime予約は必ず変数に予約時刻を保持し、Cancelは同時刻指定(Excel仕様)。
- 純ロジックモジュール(R4)では `vbCrLf`ではなく`vbLf`基準(LO互換)。Excel層で表示時にvbLf使用可(セル内改行はvbLf)。

## 13. エッジケースカタログ(実装とレビューの検査表)

取込: 空ファイル/0字抽出/1文字/巨大(300p打切→partial表示)/画像PDF/開けない(ロック)/対応外拡張子/
日本語・スペース・長いパス/同名ファイル再取込(置換)/別フォルダの同名別ファイル(E0504で明示エラー・削除しない。Wave4追加)/
同一内容別名(ハッシュでスキップ)/上限5000超/
取込中にESC/取込中にもう一度ボタン(再入guard)/embedded=0が残った状態でExcel強制終了→次回同期で再開。
同期: フォルダ未設定で同期ボタン/フォルダ削除・リネーム(E0502+missing)/ファイル差替え(更新日時2秒差)/
OneDriveオフライン(プレースホルダ: Dir()で見えるがOpenで失敗→E0302でfailed記録・次回再試行)/
自動同期中にユーザーが質問実行(guardで同期側をスキップ)。
QA: 本棚空で質問/検索0件/LLMがエラー文字列/LLMが出典形式を無視(そのまま表示・落ちない)/
質問が空/超長文質問(3000字で打切り確認)/連打(実行中フラグ)。
パック: 壊れたxlsx/シート欠落/次元不一致/自分のパックを自分で取込(全スキップ)/PII検知で中断選択/
書き出し先が読み取り専用。
環境: リボン無しPC(mock案内)/マクロ無効(使い方シートが説明)/VBAプロジェクトアクセス不可(インストーラがTrustメッセージ)/
Excel 32/64bit(Declare不使用で回避)/Mac(COM抽出不可→E0302系の丁寧な案内があること。コア動作はWindows前提と明記)。
終了: OnTime残存なし/StatusBar復元。

## 14. 実装ウェーブとファイル分担(エージェント間でファイル重複禁止)

| Wave | 担当 | 書くファイル |
|---|---|---|
| 1-A | foundation | src/core/{modAppDef,modTypes,modConfig,modLog,modUtil,modGateway,modFeatures,modDiag}.bas |
| 1-H | build | build/{build_mybookshelf.py, ovba.py, template_skeleton.xlsm(コピー), modules.json, README.md} |
| 1-I | testharness | tools/{vba_lint.py, run_lo_tests.py, README.md}, src/test/modTestRunner.bas |
| 2-B | extract | src/ingest/{modExtractor,modExtractorWord,modExtractorExcel,modExtractorAcrobat,modChunker}.bas |
| 2-C | shelf | src/ingest/{modEmbed,modShelf,modShelfSync,modEnrich}.bas |
| 2-D | qa | src/qa/{modRetrieve,modPrompts,modAsk}.bas |
| 2-E | pack | src/pack/{modPii,modPack}.bas, src/stats/modStats.bas |
| 2-F | ui | src/ui/{modUIMain,modUIShelf,modUIDashboard,modBoot}.bas, src/ui/ThisWorkbook.cls |
| 2-G | opt | src/opt/{optTts,optVision,optMarkdown,optDiffDoc}.bas |
| 2-X | tests | src/test/{modTestsPure,modTestsExcel}.bas(2-B〜2-Eの成果を読んで書く) |

Wave3: lint+LO全緑までの修正ループ / Wave4: Opus6レンズレビュー / Wave5: ビルド+ドキュメント+コミット。
