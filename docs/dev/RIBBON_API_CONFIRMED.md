# AIリボン公開API 確定台帳(一次情報)

> 出典: 担当部署からの返信メール+社内公開ページ「AIリボンAPI拡張ツール」
> (原本ヘッダ日付 2024/10/08〜2026/01/05、By N.Chikada)。2026-07-14 入手。
> 本書は MASTER_SPEC §7.1/§7.7 の「仮定」を上書きする**確定情報**である。
> 矛盾がある場合は本書が優先し、MASTER_SPEC側を本書に合わせて改訂する。

## 0. 運用上の確定事実(メール本文より)

- ラッパー関数は AzureOpenAI の APIレスポンスをそのまま返す(ラッパー側の機能制限なし。
  トークン上限エラー等もそのまま文字列で返却される)。
- **関数は互換性を保つよう維持される**(引数は後方追加)。→ V2本番実績の
  ChatGPT 12引数呼び(9=toolN, 10=effort, 11=verbosity)は公開ページ(9引数版)より
  新しい互換拡張であり、現行 modGateway.CallLLM の呼び出し規約は**変更不要**。
- 使用可能モデルはAIリボン側に依存。リボンは3/6/9/12月に更新され最新モデルが使える。
  → モデル名はconfig(`recommended_model`)駆動を維持(ハードコード禁止)。
- **非公開(公開予定なし)**: 画像生成・音声合成(TTS)・音声認識。理由: OSSモデルの
  自前サーバー負荷、image2の利用上限。これらは「AIリボンUIからの利用」のみ。
  → **optTts が仮定した `ttsSpeak` は存在しない。optTtsはビルドから撤去(D4)。**

## 1. 確定関数シグネチャ一覧

| # | 関数名 | シグネチャ(確定) | 戻り値 | 用途 |
|---|--------|------------------|--------|------|
| 1 | `ChatGPT` | `(Text, [roleSystem], [Temperature=0.4], [MaxTokens=4096], [Wait=120], [optModel], [prevU], [prevA], [toolN])` ※公開ページ版。互換後方追加で effort/verbosity(10・11番目)が存在(V2本番実績) | String | テキスト生成 |
| 2 | `GetEmbeddings` | `(Text As String)` | String(カンマ区切りベクトル。失敗時 "error") | 埋め込み |
| 3 | `CosineSimilarityN2` | `(str1, str2)`(カンマ区切りベクトル文字列×2) | 数値(0〜1) | 類似度 |
| 4 | `CosineSimilarityN` | `(vec1 As Variant, vec2 As Variant)`(一次元配列×2) | Double | 類似度(配列版) |
| 5 | `ChatGPTV` | `(Text, imageInputs, [roleSystem], [resolution], [toolN])`。imageInputs=**Base64文字列**(複数はカンマ区切り、最大10枚)。resolution="high"で高精細 | String | 画像解析 |
| 6 | `IsImageInCB` | `()` | Boolean | クリップボードに画像有無 |
| 7 | `Base64FromCB` | `([Ptn=0])`。Ptn=0→Base64文字列 / Ptn=1→Tempに保存したjpgパス | String | CB画像→Base64/パス |
| 8 | `Base64FromFile` | `(filePath As String)`。PNG/JPEG等をそのままBase64化 | String | 画像ファイル→Base64 |
| 9 | `ConvToJpeg` | `(imagePath As String)`。Temp\yyyymmdd_hhmmss.jpg を作成 | String(jpgパス) | 画像→JPEG変換 |
| 10 | `OpenWordMark` | `(Text As String)`(Markdown文字列) | (未使用) | MarkdownをWordで開く |
| 11 | `OpenMemo` | `(Text As String)` | (未使用) | テキストをメモ帳で開く |
| 12 | `CellMarkDown` | `(rng As Range, [isComment As Boolean=False])`。**セルに既に入っているMarkdown**を書式付き表示に変換する(文字列を渡す方式ではない) | (未使用) | セルMarkdown装飾 |
| 13 | `LimitCheck` | `()` | Boolean(**True=続行不可**: 期限切れ等。内部で日初の利用同意表示も行う) | 起動時チェック |

アドイン検出の公式作法: `Application.AddIns` をループし `InStr(addIn.Name, "リボンちゃん")>0`
かつ `addIn.Installed` で判定(API呼び出し不要・即時)。

## 2. PM裁定(D1〜D9)

- **D1 (modGateway.CallLLM)**: 12引数規約を維持。根拠コメントを本書参照に差し替える。
  toolN(第9引数)は `"マイ本棚AI:" & step_name` を渡す(管理側ログでツール識別できるように)。
- **D2 (RibbonAvailable)**: 「実呼び出しプローブ」を廃止し、公式のAddInsループ判定へ変更。
  アドイン名は config `ribbon_addin_name`(既定 "リボンちゃん")で可変にする。
  判定結果のセッションキャッシュは維持。
- **D3 (LimitCheck)**: modGatewayに `RunLimitCheck() As Boolean`(True=続行不可)を新設し、
  modBoot起動時に mock_llm=FALSE かつリボン検出時のみ呼ぶ。True時は案内メッセージを
  出して機能を停止せず「回答系の実行時に再度案内」する穏当運用(config `limit_check`
  既定TRUEで無効化可能なエスケープハッチ付き)。
- **D4 (optTts撤去)**: `ttsSpeak`は存在しないため modules.json から optTts を削除
  (ソースは src/opt に保管)。UI読み上げボタンは「社内AIリボンでは読み上げ機能は
  提供されていません(AIリボンUIからのみ利用可)」の親切メッセージに差し替え。
  config `feature_tts` 既定 FALSE。docs該当箇所を全て更新。
- **D5 (optVision改修)**: 確定規約へ全面改修。経路は
  `Base64FromFile(path)` → `ChatGPTV(VISION_PROMPT, b64, "", "high", toolN)`。
  PDF直渡し路線(`vision_pdf_direct`)は公式仕様上不可能と確定したため撤去し、
  対応形式は png/jpg/jpeg のみ(PDFは従来どおり丁寧な案内)。呼び出しは全て
  modGateway.TryRibbonRun 経由(R3)。公開API(ExtractImagePdf/ExtractImagePdfText)の
  シグネチャは不変。
- **D6 (optMarkdown改修)**: 確定規約へ改修。`RenderMarkdownAt` は「対象セルに
  Markdown文字列を書き込んでから `CellMarkDown(rng, False)` を呼ぶ」方式に変更。
  加えて確定関数 `OpenWordMark` を使う `OpenAnswerInWord(md As String) As String`
  ("" =成功 / "#ERR:...") を新設(UI配線は別途)。
- **D7 (CosineSimilarityN2は使わない)**: 類似検索は自前のmodUtil純ロジックを維持
  (mock動作・LOテスト可能性・全件走査の一括最適化のため)。台帳には記録のみ。
- **D8 (会話継続 prevU/prevA)**: 追加機能候補として保留(ユーザー判断待ち)。
  実装する場合は modAsk/ui_state に直近Q&A履歴を持たせ CallLLM に引数追加。
- **D9 (モデル名)**: 公開ページ例(GPT-3.5_Turbo/GPT-4o)は旧版。V2本番実績の
  config駆動(既定 gpt-5.5)を維持。リボン四半期更新に追随できるようconfigのみで変更可。

## 2b. 追加PM裁定(D10〜D13、2026-07-14 ユーザー指示に基づく)

- **D10 (TTSボタン完全削除)**: D4の「ボタンを残して案内」を改め、ボタン自体を削除する
  (ユーザー裁定: 引き算の美学・UI簡素化)。modUIMainのOnTtsButtonと btn_tts 生成を撤去。
  optTts.basソースの保管とmodules.json撤去(D4)は維持。docsの読み上げ記述は
  「AIリボン本体でのみ利用可」の1行注記に縮約。
- **D11 (続けて質問=深掘り機能の移植)**: V2実証済みのUX(回答末尾に深掘り候補を提示+
  「続けて質問」ボタン)をマイ本棚AIへ移植する。V2との違い: 会話履歴の受け渡しに
  確定引数 prevU/prevA(;;;区切り・新しい順)を使う。
  インターフェース契約(実装者はこれに従う):
  - modGateway.CallLLM に Optional prevU As String = "" / Optional prevA As String = "" を
    後方互換で追加し、Application.Run("ChatGPT",...)の第7・8引数に渡す。
  - modAsk 新公開API: CanFollowup() As Boolean(直近回答が存在するか)/
    AskFollowup(ByVal followupText As String)(直近の会話履歴を添えて再質問。
    履歴はmodAskモジュール内変数でセッション保持、最大 config followup_max_pairs
    既定3 ペア)。深掘り候補はV2同様、検証応答の [[FOLLOWUP: a | b]] 形式を
    パースし、回答本文末尾に「深掘り候補(『続けて質問』でそのまま聞けます)」
    ブロックとして追記してから RenderAnswer に渡す(RenderAnswer契約は不変)。
  - modPrompts: 深掘り候補を要求する指示行を deep_verify/quick 系プロンプトに追加。
  - modUIMain 新ボタン OnFollowupButton()(InputBoxで追質問を取り、
    modAsk.AskFollowup を呼ぶ。CanFollowup=Falseなら丁寧な案内)。
- **D12 (Wordで開く=対話型文書生成へ強化)**: 単純転記(弱)ではなく「指示文→LLM整形→
  Word起動」(強)にする(ユーザー裁定)。UIボタン押下→InputBox
  「どんな文書に仕上げますか?(例: お客様向けの回答文書風に/社内回覧用の要約に)」→
  optMarkdown 新公開API ExportAnswerAsDoc(回答本文+指示文を受け、modGateway.CallLLM
  (step_name="word_export")で整形→TryRibbonRun("OpenWordMark", ...)でWordを開く)。
  指示文が空なら整形をスキップして直接OpenWordMark(=旧OpenAnswerInWord相当)。
  OpenAnswerInWordは残す(ExportAnswerAsDocの空指示経路が内部で使う)。
  引数の渡し方は modFeatures.InvokeFeature の既存規約に従うこと。
- **D13 (スクリーンショット/画像ファイル取込)**: 確定関数 IsImageInCB/Base64FromCB(Ptn=1)を
  使い「クリップボードの画像を本棚に取り込む」導線を追加する(ユーザー裁定: 社内は
  スクショ文化。マニュアル抜粋・Web検索結果等の取込ニーズ大)。
  - optVision 新公開API: HasClipboardImage() As Boolean / SaveClipboardImage() As String
    (Base64FromCB(Ptn=1)でTempへjpg保存しパスを返す。失敗は"#ERR:...")。
  - modUIShelf 新ボタン OnIngestScreenshot(): HasClipboardImage確認→タイトルをInputBoxで
    取得→SaveClipboardImageのjpgを「<タイトル>_yyyymmdd_hhmmss.jpg」として
    本棚フォルダ(未設定ならTempのまま)へFileCopy→modShelf.IngestFileへ渡す。
  - modShelf.IngestFile のvisionフォールバック条件を拡張: 現行「E0303のみ」→
    「E0303、または E0301 かつ 拡張子がpng/jpg/jpeg」でも
    InvokeFeature("vision","ExtractImagePdfText",path) を試す(これにより画像ファイルの
    ドラッグ追加も自然に取込可能になる)。feature_vision無効時の案内文も画像拡張子に
    対応させる。

- **D14 (opt機能フラグの既定TRUE昇格)**: feature_vision / feature_markdown の既定値を
  FALSE→TRUEへ変更(公式公開仕様でシグネチャが確定したため。「仕様確認後にTRUEへ」の
  条件が成就)。これによりスクショ取込ボタン・Wordで開くボタンが初期状態で表示される。
  実機で問題が出た場合はconfigシートでFALSEに戻すだけで撤去できる(従来どおり)。
  feature_tts は非公開確定のためFALSE固定。
- **D15 (config新キー)**: followup_max_pairs(既定5)/ word_export_effort(medium)/
  word_export_verbosity(medium) をシードに追加(D11/D12の実装が参照)。

## 3. 未確定のまま残るもの

- effort/verbosity(第10・11引数)の正式仕様(V2本番実績はあるが公開ページ未掲載)。
  高橋くんのモジュール抽出で確認予定。
- `LimitCheck` の True/False の正確な意味(公式サンプルの解釈: True=中断)。
  同上で確認予定。挙動が逆と判明した場合は modGateway.RunLimitCheck 内の1行のみ修正。
- optDiffDoc の対応形式拡張(modExtractor層制限の扱い)は本件と無関係に継続保留。
