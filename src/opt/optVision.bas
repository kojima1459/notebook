Attribute VB_Name = "optVision"
Option Explicit

' ============================================================================
' optVision - Vision API(文字起こし)+画像PDF OCRのオーケストレーション
'             (opt機能・MASTER_SPEC §7.7)
' ----------------------------------------------------------------------------
' 役割:
'   通常抽出(modExtractor)がE0303(画像PDF・文字が取れない)を返した資料や、
'   画像ファイル(png/jpg/jpeg)そのものについて、AIリボンのVision機能
'   (ChatGPTV・確定済み)へ渡して文字起こしする。この層が無くても製品は
'   成立し、modules.jsonから1行削除するだけで撤去できる(§7.7)。
'
' 設計判断:
'   ・対応形式は画像ファイル(png/jpg/jpeg)のみ。PDFの直渡し
'     (旧config vision_pdf_direct 路線)は、ChatGPTVのimageInputsが
'     「Base64エンコードした画像」に限定されると確定した
'     (出典: RIBBON_API_CONFIRMED.md §1 #5)ため完全撤去した。
'     PDFが渡された場合は丁寧な案内文を含む失敗として扱う。
'   ・抽出経路は2段階の確定規約:
'       (1) Base64FromFile(filePath) で画像ファイルをBase64文字列化
'       (2) ChatGPTV(Text, imageInputs, roleSystem, resolution, toolN) で文字起こし
'     どちらも modGateway.TryRibbonRun のみ経由で呼ぶ(R3)。
'   ・失敗しても例外を外に出さない。ExtractImagePdfはBoolean(§7.7契約の
'     とおり)、ExtractImagePdfTextは "#ERR:..." 文字列 または 全文。
'   ・スクリーンショット取込(裁定D13): クリップボードの画像を本棚へ取り込む
'     導線のために HasClipboardImage / SaveClipboardImage を公開する。
'     確定関数 IsImageInCB / Base64FromCB(Ptn=1) の唯一の利用箇所
'     (呼び出しはTryRibbonRun経由・R3)。UI導線はmodUIShelf.OnIngestScreenshot。
'   ・mock_llm=TRUE のときはリボンを呼ばない: HasClipboardImageはFalse、
'     SaveClipboardImageは親切な "#ERR:mockモード…" 案内を返す(optMarkdownの
'     mock時の流儀に合わせる)。
'   ・ExtractImagePdfの失敗理由をExtractImagePdfTextへ伝えるため、
'     モジュール内Private変数(mLastErrorMsg)に直近の失敗メッセージを
'     保持する(VBAは単一スレッドで実行されるため競合の心配がない)。
'   ・modUIMain.SetStage 以外のコア参照は行わない(基盤層のみ参照可・§7.7)。
'   ・画像PDF(E0303)のOCR取込(2026-07-31 R6): 会社公式ツール同梱の
'     Ghostscript でページ毎にJPEG化してから、上と同じ経路(Base64FromFile→
'     ChatGPTV)で1ページずつ読む。「PDF→画像変換の手段がVBA単体に無い」という
'     唯一の穴を、公式配布物の再利用で塞ぐ。入口は ExtractPdfOcrPagedText。
'     コマンド文字列の組み立ては optOcrCore(純ロジック)に切り出してある。
'     GSを動かす道具(一時フォルダ・非同期起動・完了待ち・後始末)は
'     optGsTxt へ移設済み(R10-3b)で、ここは「何ページ描いて何をOCRするか」
'     という段取りだけを持つ。
'   ・Ghostscriptの配布・自動検出(2026-07-31 R9): dist/Ghostscript を常置
'     配布物に同梱する方針(要件書R9)に伴い、ResolveGsExeの解決順を
'     「config ghostscript_path → 同梱(ThisWorkbook.Path\Ghostscript\
'     gswin32c.exe)→ config ghostscript_search_dirs(IT焼き込み用)→
'     案内カード(1回きり・赤エラーでなく)」の4段へ拡張した。
'     候補パスの組み立ては optOcrCore.GsCandidatePaths / GsCandidatesForFolder
'     (純ロジック)に切り出し、実在確認(Dir$)とApplication.FileDialogは
'     副作用なのでここに残す(§7.7契約: opt層の唯一のコア参照は
'     modUIMain.SetStageのみ・opt層内は自由参照)。
'     解決結果はセッション内(モジュール変数)にキャッシュし、案内カードは
'     このブックを開いている間は1回しか出さない(見つからなくても毎回の
'     OCR実行のたびにダイアログで割り込まない)。
' ============================================================================

' === 確定済みシグネチャ(出典: RIBBON_API_CONFIRMED.md §1 #5, #8 / 裁定D5) ===
'   Base64FromFile(filePath As String) -> String
'     PNG/JPEG等の画像ファイルをそのままBase64文字列にして返す(#8)。
'   ChatGPTV(Text, imageInputs, [roleSystem], [resolution], [toolN]) -> String
'     imageInputs = Base64FromFileで得たBase64文字列(複数はカンマ区切り、
'     最大10枚)。resolution="high" で高精細解析(#5)。
'     本モジュールは (VISION_PROMPT, b64, "", "high", "マイ本棚AI:vision") で呼ぶ。
'   戻り値の扱い(§0: ラッパーはAPIレスポンスをそのまま文字列で返す):
'     失敗時は "#ERR:" 始まり(TryRibbonRun側)/空文字列/上限系メッセージが
'     返り得るため、IsVisionError() で失敗と判定する。
'   PDF直渡しは公式仕様上不可能と確定(imageInputsは画像のBase64限定)。
'   PDF→画像変換の手段はVBA単体に無いため、PDFは案内文つきの失敗とする。
'   IsImageInCB() -> Boolean(台帳§1 #6)
'     クリップボードに画像があるかどうかを即時判定する。
'   Base64FromCB([Ptn=0]) -> String(台帳§1 #7)
'     Ptn=0でBase64文字列、Ptn=1で「Tempに保存したjpgのパス」を返す。
'     本モジュール(SaveClipboardImage)は Ptn=1 のjpgパス取得のみ使う。
' ============================================================================

Private Const RIBBON_FUNC_NAME As String = "ChatGPTV"
Private Const BASE64_FUNC_NAME As String = "Base64FromFile"
Private Const VISION_RESOLUTION As String = "high"
Private Const VISION_TOOL_NAME As String = "マイ本棚AI:vision"
Private Const IMAGE_EXTS As String = "png,jpg,jpeg"
' 2026-07-16 実機フィードバック(「どんなスクショでも1チャンクにしかならず
' 精度が悪い」)を受けて逐語書き起こしの指示を強化: 表の構造・数値・単位の
' 厳密な転記と、見出しによる構造化を明示的に要求する(書き起こしが厚くなる
' ほどチャンク分割・検索の材料が増える)。
Private Const VISION_PROMPT As String = _
    "この画像に写っている文字情報を、一字一句省略せずにすべて書き起こしてください。" & vbLf & _
    "・見出しやセクション名は行頭に「# 」を付けて構造を保つこと" & vbLf & _
    "・表は1行=1レコードの形で、列名と値の対応が分かるように書き出すこと" & vbLf & _
    "・数値・金額・日付・単位・記号は画像のとおり正確に転記すること(要約や丸めは禁止)" & vbLf & _
    "・図やグラフは、読み取れる軸・凡例・数値を含めて内容を文章で説明すること" & vbLf & _
    "・画像に無い情報を補って書かないこと"

Private mLastErrorMsg As String

' R9: Ghostscript解決のセッションキャッシュと案内カードの1回きりフラグ。
Private mGsExeCache As String        ' 解決済みパス(mGsResolved=Trueのときのみ有効)
Private mGsResolved As Boolean       ' Trueなら以後mGsExeCacheをそのまま使う
Private mGsGuidanceShown As Boolean  ' 案内カード(FileDialog)はセッション中1回だけ

' R10-2: GS解決の観測性(実機初報A)。E0303のerr_log detailへ載せるための
' 直近状態を保持する(いずれも副作用ゼロの記録用途のみ)。
Private mGsLastCands As String  ' FindGsExeByCandidatesが最後に試した候補(|区切り)
Private mGsCardState As String  ' 案内カードの結果: ""/"cancel"/"nofind"/"saved"

Public Function Ping() As Boolean
    Ping = True
End Function

' ----------------------------------------------------------------------------
' HasClipboardImage - クリップボードに画像があるか(裁定D13)。
'   確定関数 IsImageInCB()(台帳§1 #6)のラッパー。mock_llm=TRUE・リボン
'   不在・失敗時はFalse(失敗の記録はTryRibbonRun側がE0202で行うため、
'   ここでの追加ログは不要=握りつぶしではない)。
' ----------------------------------------------------------------------------
Public Function HasClipboardImage() As Boolean
    HasClipboardImage = False
    If modConfig.GetBool("mock_llm", True) Then Exit Function
    Dim res As Variant
    res = modGateway.TryRibbonRun("IsImageInCB", Array())
    If VarType(res) = vbBoolean Then HasClipboardImage = CBool(res)
End Function

' ----------------------------------------------------------------------------
' SaveClipboardImage - クリップボードの画像をTempへjpg保存しパスを返す
'   (裁定D13)。確定関数 Base64FromCB(Ptn=1)(台帳§1 #7: Ptn=1で
'   「Tempに保存したjpgのパス」が返る)のラッパー。
'   成功: jpgのフルパス / 失敗: "#ERR:..."(例外は出さない)。
' ----------------------------------------------------------------------------
Public Function SaveClipboardImage() As String
    If modConfig.GetBool("mock_llm", True) Then
        SaveClipboardImage = "#ERR:mockモード(mock_llm=TRUE)ではクリップボード取込を利用できません。" & _
            "config の mock_llm を FALSE にすると実際に取り込めるようになります。"
        Exit Function
    End If

    Dim res As Variant
    res = modGateway.TryRibbonRun("Base64FromCB", Array(1&))
    Dim s As String: s = CStr(res)

    If LenB(s) = 0 Then
        SaveClipboardImage = "#ERR:クリップボード画像の保存に失敗しました(応答が空)"
    Else
        ' TryRibbonRun失敗時の "#ERR:E0202:..." はそのまま呼び出し元へ渡す
        SaveClipboardImage = s
    End If
End Function

' ----------------------------------------------------------------------------
' ExtractImagePdf - path(画像ファイル)からVisionで文字を読み取り、
'   ExtractedPage 1件(page=1、全文)としてpagesへ格納する(MASTER_SPEC §7.7)。
'   成功: True(pagesに1件) / 失敗: False(pagesは空配列のまま)。
'   ※関数名の「Pdf」は§7.7契約の歴史的経緯によるもの。実際の対応形式は
'     png/jpg/jpeg のみで、PDFは案内文つきの失敗になる(裁定D5)。
' ----------------------------------------------------------------------------
Public Function ExtractImagePdf(ByVal path As String, ByRef pages() As ExtractedPage) As Boolean
    mLastErrorMsg = ""
    ReDim pages(0 To 0)

    On Error GoTo Fail

    Dim ext As String
    ext = modUtil.ExtOf(path)

    If Not IsImageExt(ext) Then
        If ext = "pdf" Then
            mLastErrorMsg = "#ERR:E0303:" & modLog.FriendlyMessage("E0303") & _
                "(社内AIリボンの画像解析はPDFを直接読み取れない仕様です。お手数ですが、" & _
                "該当ページをスクリーンショット等で画像(png/jpg)にしてからお試しください。)"
        Else
            mLastErrorMsg = "#ERR:E0301:このファイル形式(" & ext & ")はVisionで直接読み取れません。" & _
                "対応形式は画像(png/jpg/jpeg)のみです。"
        End If
        modLog.LogError "E0303", "optVision.ExtractImagePdf", "未対応ext=" & ext & " path=" & modUtil.SafeLeft(path, 300)
        ExtractImagePdf = False
        Exit Function
    End If

    modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDDBC) & " 画像から文字を読み取っています…"

    ' (1) 画像ファイル → Base64文字列(確定: Base64FromFile / 台帳§1 #8)
    Dim b64 As String
    b64 = SafeResultToString(modGateway.TryRibbonRun(BASE64_FUNC_NAME, Array(path)))

    If LenB(Trim$(b64)) = 0 Or Left$(b64, 5) = "#ERR:" Then
        modUIMain.SetStage ""
        mLastErrorMsg = "#ERR:E0303:画像ファイルの読み込み(Base64変換)に失敗しました。" & _
            "ファイルが開けるか、壊れていないかをご確認ください。" & _
            IIf(LenB(b64) > 0, "(詳細: " & modUtil.SafeLeft(b64, 200) & ")", "")
        modLog.LogError "E0303", "optVision.ExtractImagePdf", "Base64FromFile失敗 path=" & _
            modUtil.SafeLeft(path, 300) & " : " & modUtil.SafeLeft(b64, 200)
        ExtractImagePdf = False
        Exit Function
    End If

    ' (2) Base64 → 文字起こし(確定: ChatGPTV / 台帳§1 #5・裁定D5)
    Dim result As Variant
    result = modGateway.TryRibbonRun(RIBBON_FUNC_NAME, _
        Array(VISION_PROMPT, b64, "", VISION_RESOLUTION, VISION_TOOL_NAME))

    modUIMain.SetStage ""

    Dim s As String
    s = SafeResultToString(result)

    If IsVisionError(s) Then
        mLastErrorMsg = "#ERR:E0303:文字の読み取りに失敗しました。" & _
            "画像が不鮮明か、対応外の内容の可能性があります。" & _
            IIf(LenB(s) > 0, "(詳細: " & modUtil.SafeLeft(s, 200) & ")", "")
        modLog.LogError "E0303", "optVision.ExtractImagePdf", modUtil.SafeLeft(path & " : " & s, 300)
        ExtractImagePdf = False
        Exit Function
    End If

    Dim tmp(0 To 0) As ExtractedPage
    tmp(0).page = 1
    tmp(0).Text = s
    pages = tmp

    ExtractImagePdf = True
    Exit Function

Fail:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    On Error GoTo 0
    modUIMain.SetStage ""
    modLog.LogError "E0303", "optVision.ExtractImagePdf", errDesc
    mLastErrorMsg = "#ERR:E0303:文字の読み取り処理でエラーが発生しました: " & errDesc
    ExtractImagePdf = False
End Function

' ----------------------------------------------------------------------------
' ExtractImagePdfText - ExtractImagePdfのText版ラッパ(MASTER_SPEC §7.7)。
'   modFeatures.InvokeFeature("vision","ExtractImagePdfText", path) から
'   文字列一枚で呼べるようにするための契約(戻り値: 全文 または "#ERR:..." )。
' ----------------------------------------------------------------------------
Public Function ExtractImagePdfText(ByVal path As String) As String
    Dim pages() As ExtractedPage
    If ExtractImagePdf(path, pages) Then
        ExtractImagePdfText = JoinPagesText(pages)
    ElseIf LenB(mLastErrorMsg) > 0 Then
        ExtractImagePdfText = mLastErrorMsg
    Else
        ExtractImagePdfText = "#ERR:E0303:画像からの文字抽出に失敗しました"
    End If
End Function

' ============================================================================
' ExtractPdfOcrPagedText - 画像PDFをGhostscriptでページ毎にJPEG化し、
'   1ページ=1回のChatGPTV(既存のVISION_PROMPT・resolution="high")で
'   文字起こしして、modUtil.JoinPagedText の1本の文字列にして返す(R6)。
' ----------------------------------------------------------------------------
'   成功: ページ付きテキスト(先頭行が@@NEXUS_TRUNCATED@@なら上限で打ち切り)
'   失敗: "#ERR:E0303:<何が起きたか+どうすればよいか>"
'
'   ・GSの起動はブロッキング待ちにしない(壊れたPDF1つでExcelが永久に固まる
'     ことを避けるため)。完了フラグファイルの出現をDoEventsつきで監視し、
'     config vision_pdf_timeout_sec(既定120秒)を超えたら失敗として返す。
'     GSプロセスのkillはしない(こちらが壊す方が危ない)。
'   ・一時フォルダは成功・失敗どちらの経路でも必ず片付ける。後始末は
'     R6規約に従って別Sub(optGsTxt.CleanupOcrFolder)へ切り出してある
'     (稼働中のエラーハンドラの中では On Error Resume Next が効かないため)。
'   ・将来課題: ChatGPTVは最大10枚のバッチ入力に対応しているが、複数枚を
'     まとめて渡すとページ境界が崩れて出典ページ番号が信用できなくなるため
'     v1では使わない。ページ番号を保ったまま束ねる方法が確立できたら再検討する。
' ============================================================================
Public Function ExtractPdfOcrPagedText(ByVal path As String) As String
    Dim folderPath As String: folderPath = ""

    On Error GoTo Fail

    If modUtil.ExtOf(path) <> "pdf" Then
        ExtractPdfOcrPagedText = "#ERR:E0303:この処理はPDF専用です(画像ファイルは" & _
            "そのまま取り込めます)。"
        Exit Function
    End If

    If modConfig.GetBool("mock_llm", True) Then
        ExtractPdfOcrPagedText = "#ERR:E0303:mockモード(mock_llm=TRUE)では画像PDFの" & _
            "OCR取込を利用できません。config の mock_llm を FALSE にすると実際に" & _
            "読み取れるようになります。"
        Exit Function
    End If

    Dim gsExe As String: gsExe = ResolveGsExe()
    If LenB(gsExe) = 0 Then
        ' R10-2: 「GS未検出」は実機初報Aの本体だが従来は完全無言だった。
        ' 候補パス・案内カードの状態を必ずerr_logへ残す(観測性ゼロの解消)。
        modLog.LogError "E0303", "optVision.ExtractPdfOcrPagedText", modUtil.SafeLeft( _
            "GS未解決 wbPath=" & ThisWorkbook.path & " cand=" & mGsLastCands & _
            " card=" & mGsCardState, 2000)
        ExtractPdfOcrPagedText = "#ERR:E0303:画像PDFを読み取るための Ghostscript が" & _
            "見つかりませんでした。配布zip(リポジトリのZIPをDL→解凍したもの)は " & _
            "Ghostscript フォルダを同梱済みのはずです。このファイルと同じ場所に " & _
            "Ghostscript フォルダがあるかご確認いただくか、config シートの " & _
            "ghostscript_path に gswin32c.exe のフルパスを設定してください。" & vbLf & _
            "このファイルの場所: " & ThisWorkbook.path & vbLf & _
            "zipの中から直接開くと動きません。zipを右クリック→[すべて展開]で展開し、" & _
            "展開先の MyBookshelf.xlsm を開いてください。"
        Exit Function
    End If

    folderPath = optGsTxt.MakeOcrFolder()
    If LenB(folderPath) = 0 Then
        ExtractPdfOcrPagedText = "#ERR:E0303:作業用の一時フォルダを作成できませんでした。" & _
            "一時フォルダ(TEMP)の空き容量と書き込み権限をご確認ください。"
        Exit Function
    End If

    Dim maxPages As Long: maxPages = modConfig.GetLong("vision_pdf_max_pages", 20)
    Dim dpi As Long: dpi = modConfig.GetLong("vision_pdf_dpi", 150)
    Dim waitSec As Long: waitSec = modConfig.GetLong("vision_pdf_timeout_sec", 120)
    If waitSec < 10 Then waitSec = 10

    Dim renderCap As Long: renderCap = optOcrCore.RenderCapFor(maxPages)
    Dim flagPath As String: flagPath = optOcrCore.DoneFlagFor(folderPath)
    Dim runCmd As String
    runCmd = optOcrCore.BuildRunCommand( _
        optOcrCore.BuildGsCommand(gsExe, path, optOcrCore.OutPatternFor(folderPath), dpi, renderCap), _
        flagPath)

    modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDDBC) & " PDFを画像に変換しています…"

    Dim gsErrNum As Long: gsErrNum = 0
    Dim gsErrDesc As String: gsErrDesc = ""
    If Not optGsTxt.RunGsAsync(runCmd, gsErrNum, gsErrDesc) Then
        modUIMain.SetStage ""
        optGsTxt.CleanupOcrFolder folderPath
        ' R10-2: 従来はパスのみで原因が残らなかった。WScript.Shellがポリシーで
        ' ブロックされる端末の切り分けに、CreateObject/Run失敗時のErr情報が必須。
        modLog.LogError "E0303", "optVision.ExtractPdfOcrPagedText", modUtil.SafeLeft( _
            "GS起動失敗 err#" & gsErrNum & ": " & gsErrDesc & " " & path, 2000)
        ExtractPdfOcrPagedText = "#ERR:E0303:PDFを画像に変換する処理を開始できませんでした。" & _
            "config の ghostscript_path が正しいかご確認ください。"
        Exit Function
    End If

    Dim finished As Boolean: finished = optGsTxt.WaitForDoneFlag(flagPath, waitSec)
    Dim foundN As Long: foundN = CountRenderedPages(folderPath, renderCap)

    ' タイムアウトした場合、最後の1枚はGSが書いている途中の可能性がある。
    ' 途中のJPEGを読ませても意味が無いので使わない(2枚以上あるときだけ)。
    If Not finished And foundN > 1 Then foundN = foundN - 1

    If foundN = 0 Then
        modUIMain.SetStage ""
        optGsTxt.CleanupOcrFolder folderPath
        modLog.LogError "E0303", "optVision.ExtractPdfOcrPagedText", _
            "描画0枚 finished=" & finished & " " & modUtil.SafeLeft(path, 200)
        If finished Then
            ExtractPdfOcrPagedText = "#ERR:E0303:このPDFからページ画像を作れませんでした。" & _
                "ファイルが壊れているか、パスワードで保護されている可能性があります。"
        Else
            ExtractPdfOcrPagedText = "#ERR:E0303:PDFの画像変換が" & waitSec & "秒以内に" & _
                "終わりませんでした。ページ数の少ないPDFに分けるか、config の " & _
                "vision_pdf_timeout_sec を大きくしてからお試しください。"
        End If
        Exit Function
    End If

    Dim truncated As Boolean: truncated = optOcrCore.IsTruncatedCount(foundN, maxPages)
    Dim keepN As Long: keepN = optOcrCore.KeepPageCount(foundN, maxPages)

    Dim ocrText As String
    ocrText = OcrRenderedPages(folderPath, keepN, truncated)

    modUIMain.SetStage ""
    optGsTxt.CleanupOcrFolder folderPath

    ' 全ページ失敗のときだけ "#ERR:..." が返る(OcrRenderedPages内で判定済み)。
    ExtractPdfOcrPagedText = ocrText
    Exit Function

Fail:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    On Error GoTo 0
    modUIMain.SetStage ""
    optGsTxt.CleanupOcrFolder folderPath
    modLog.LogError "E0303", "optVision.ExtractPdfOcrPagedText", errDesc
    ExtractPdfOcrPagedText = "#ERR:E0303:画像PDFの読み取り処理でエラーが発生しました: " & errDesc
End Function

' ----------------------------------------------------------------------------
' ResetGsGuidance - GS未検出の案内カード(セッション1回きり)を再提示可能に
'   戻す(R10-2)。mGsGuidanceShownのみFalseへ戻し、解決済みキャッシュ
'   (mGsResolved/mGsExeCache)には触れない(設置済みのGSは再解決不要)。
'   利用者が能動的に「資料を追加」を実行する入口(modShelf.AddFilesResult)
'   から modFeatures.InvokeFeature("vision","ResetGsGuidance",…) 経由で
'   呼ばれる想定。自動同期(modShelfSync)からは呼ばれない=起動時に
'   モーダルが出ない現行性質を維持する。戻り値は契約合わせの空文字列。
' ----------------------------------------------------------------------------
Public Function ResetGsGuidance() As String
    mGsGuidanceShown = False
    ResetGsGuidance = ""
End Function

' R10-3: テキストPDFのCOM無し抽出(実体はoptGsTxt・容量都合の分割)。
' vision機能の行き先はoptVision固定なので受け口だけ置く。
Public Function ExtractPdfTextNoOcr(ByVal path As String) As String
    ExtractPdfTextNoOcr = optGsTxt.ExtractPdfTextNoOcr(path)
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー(GS実行まわりの5本のみoptGsTxt用にPublic・R10-3)
' ----------------------------------------------------------------------------

' ページ画像を1枚ずつOCRして1本のページ付きテキストにする。
' 1ページの失敗で資料全体を捨てない(読めたページだけを返す)。
Private Function OcrRenderedPages(ByVal folderPath As String, ByVal keepN As Long, _
                                  ByVal truncated As Boolean) As String
    Dim pages() As ExtractedPage
    ReDim pages(0 To keepN - 1)
    Dim okN As Long: okN = 0

    Dim i As Long
    For i = 1 To keepN
        modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDDBC) & " OCR中… " & i & "/" & keepN & " ページ"

        Dim jpgPath As String
        jpgPath = folderPath & "\" & optOcrCore.PageJpgName(i)

        Dim b64 As String
        b64 = SafeResultToString(modGateway.TryRibbonRun(BASE64_FUNC_NAME, Array(jpgPath)))

        If LenB(Trim$(b64)) > 0 And Left$(b64, 5) <> "#ERR:" Then
            Dim s As String
            s = SafeResultToString(modGateway.TryRibbonRun(RIBBON_FUNC_NAME, _
                Array(VISION_PROMPT, b64, "", VISION_RESOLUTION, VISION_TOOL_NAME)))
            If Not IsVisionError(s) Then
                pages(okN).page = i
                pages(okN).Text = s
                okN = okN + 1
            Else
                modLog.LogError "E0303", "optVision.ExtractPdfOcrPagedText", _
                    "p" & i & " 読取失敗: " & modUtil.SafeLeft(s, 200)
            End If
        Else
            modLog.LogError "E0303", "optVision.ExtractPdfOcrPagedText", _
                "p" & i & " Base64失敗: " & modUtil.SafeLeft(b64, 200)
        End If
    Next i

    If okN = 0 Then
        OcrRenderedPages = "#ERR:E0303:ページ画像は作れましたが、文字を読み取れませんでした。" & _
            "しばらく時間を置いてからもう一度お試しください。"
        Exit Function
    End If

    If okN < keepN Then ReDim Preserve pages(0 To okN - 1)
    OcrRenderedPages = modUtil.JoinPagedText(pages, truncated)
End Function

' ----------------------------------------------------------------------------
' ResolveGsExe - Ghostscript実行ファイルを解決する(R9: 4段階)。
'   (1) config ghostscript_path(明示フルパス)
'   (2) 同梱(ThisWorkbook.Path\Ghostscript\gswin32c.exe)
'   (3) config ghostscript_search_dirs(セミコロン区切り。IT焼き込み用)
'   (4) 見つからない場合: 1回きりの案内カード→フォルダ選択→即続行
'   候補の組み立ては optOcrCore.GsCandidatePaths(純ロジック)、実在確認は
'   ここ(Dir$)。解決済みパスはセッション内(モジュール変数)にキャッシュする。
' ----------------------------------------------------------------------------
Private Function ResolveGsExe() As String
    If mGsResolved Then
        ResolveGsExe = mGsExeCache
        Exit Function
    End If

    Dim found As String
    found = FindGsExeByCandidates()
    If LenB(found) > 0 Then
        CacheGsExe found
        ResolveGsExe = found
        Exit Function
    End If

    ' (4) 案内カード。見つからなかった/カード自体を出したことが既にある
    '     場合は、今回はダイアログを出さず未解決のまま返す(1回きり)。
    '     mGsResolvedはTrueにしない: config書き換え・フォルダ設置が後から
    '     行われた場合に、次回呼び出しで(3)までを毎回再チェックできるように
    '     しておくため(ダイアログさえ出さなければ再探索自体は軽い)。
    If Not mGsGuidanceShown Then
        mGsGuidanceShown = True
        Dim viaCard As String
        viaCard = OfferGsFolderPicker()
        If LenB(viaCard) > 0 Then
            CacheGsExe viaCard
            ResolveGsExe = viaCard
            Exit Function
        End If
    End If

    ResolveGsExe = ""
End Function

' config ghostscript_path / 同梱 / ghostscript_search_dirs の順で候補を
' 実在確認する(optOcrCore.GsCandidatePathsが組み立てた候補文字列を使う)。
' R10-3でPublic化: 案内カードを出せないoptGsTxt用の「静かな解決」がこの関数。
Public Function FindGsExeByCandidates() As String
    Dim cfgPath As String
    cfgPath = modConfig.GetString("ghostscript_path", "")
    Dim searchDirs As String
    searchDirs = modConfig.GetString("ghostscript_search_dirs", "")

    Dim candStr As String
    candStr = optOcrCore.GsCandidatePaths(cfgPath, ThisWorkbook.path, searchDirs)
    mGsLastCands = candStr   ' R10-2: E0303のerr_log detailで参照するため毎回更新
    FindGsExeByCandidates = FirstExistingCandidate(candStr)
End Function

' "|" 区切りの候補文字列から、実在する最初の1件を返す(無ければ空文字列)。
Private Function FirstExistingCandidate(ByVal candStr As String) As String
    If LenB(candStr) = 0 Then Exit Function
    Dim cands() As String: cands = Split(candStr, "|")
    Dim i As Long
    For i = LBound(cands) To UBound(cands)
        If PathExists(cands(i)) Then
            FirstExistingCandidate = cands(i)
            Exit Function
        End If
    Next i
End Function

Private Sub CacheGsExe(ByVal resolvedPath As String)
    mGsExeCache = resolvedPath
    mGsResolved = True
End Sub

' ----------------------------------------------------------------------------
' OfferGsFolderPicker - 1回きりの丁寧な案内(赤エラーでなく)+フォルダ選択。
'   選択フォルダ直下または bin\ 直下に gswin32c.exe があれば
'   config ghostscript_path へ保存して即そのパスを返す。無ければ
'   modUIMain.SetStage で1行案内して ""(未解決)を返す(§7.7契約により
'   opt層からのUI参照はSetStageのみ許可)。
'   ※LibreOfficeテスト環境にはApplication.FileDialogが無いため、この関数と
'   PickGsFolder(実際のダイアログ呼び出し)はLOの純ロジックテスト(モード1)
'   からは一切呼ばれない(候補の組み立てはoptOcrCore側の純ロジックとして
'   分離済み・そちらだけが境界値テストの対象)。
' ----------------------------------------------------------------------------
Private Function OfferGsFolderPicker() As String
    Dim resp As VbMsgBoxResult
    resp = MsgBox( _
        "画像PDFの読み取りには Ghostscript が必要です。" & vbLf & _
        "配布zipの Ghostscript フォルダをこのファイルの隣に置くか、" & vbLf & _
        "場所を指定してください。" & vbLf & vbLf & _
        "[OK] を押すとフォルダを選ぶ画面が開きます。", _
        vbInformation + vbOKCancel, modAppDef.APP_NAME)
    If resp <> vbOK Then
        mGsCardState = "cancel"
        modLog.LogUsage "gs_guidance_card", mGsCardState, ""
        Exit Function
    End If

    Dim picked As String
    picked = PickGsFolder()
    If LenB(picked) = 0 Then
        mGsCardState = "cancel"
        modLog.LogUsage "gs_guidance_card", mGsCardState, ""
        Exit Function
    End If

    Dim candStr As String: candStr = optOcrCore.GsCandidatesForFolder(picked)
    Dim found As String: found = FirstExistingCandidate(candStr)
    If LenB(found) = 0 Then
        modUIMain.SetStage "選んだフォルダに gswin32c.exe が見つかりませんでした。"
        mGsCardState = "nofind"
        modLog.LogUsage "gs_guidance_card", mGsCardState, modUtil.SafeLeft(picked, 200)
        Exit Function
    End If

    modConfig.SetValue "ghostscript_path", found
    mGsCardState = "saved"
    modLog.LogUsage "gs_guidance_card", mGsCardState, modUtil.SafeLeft(found, 200)
    OfferGsFolderPicker = found
End Function

' フォルダ選択ダイアログ(msoFileDialogFolderPicker=4)。名前付き定数は使わない
' (modShelfSync.PickShelfFolderと同じ慣習・LibreOffice互換のため)。
' キャンセル・実行時エラー(FileDialog非対応環境)はどちらも空文字列で返す。
Private Function PickGsFolder() As String
    On Error GoTo NoDialog
    Dim fd As Object
    Set fd = Application.FileDialog(4)   ' msoFileDialogFolderPicker
    fd.Title = "Ghostscript フォルダを選択してください"
    If fd.Show <> -1 Then Exit Function   ' キャンセル
    If fd.SelectedItems.count < 1 Then Exit Function
    PickGsFolder = CStr(fd.SelectedItems(1))
    Exit Function
NoDialog:
    PickGsFolder = ""
End Function

' page_001.jpgから連番で何枚できているかを数える(Dirの列挙状態に依存しない)。
Private Function CountRenderedPages(ByVal folderPath As String, ByVal renderCap As Long) As Long
    Dim i As Long
    For i = 1 To renderCap
        If Not PathExists(folderPath & "\" & optOcrCore.PageJpgName(i)) Then Exit For
    Next i
    CountRenderedPages = i - 1
End Function

' R10-3bで optGsTxt へ移設したGS実行の道具(MakeOcrFolder/RunGsAsync/
' WaitForDoneFlag/CleanupOcrFolder)からも使うためPublic。
Public Function PathExists(ByVal p As String) As Boolean
    On Error Resume Next
    PathExists = (LenB(Dir$(p)) > 0)
    On Error GoTo 0
End Function

Private Function IsImageExt(ByVal ext As String) As Boolean
    IsImageExt = (InStr(1, "," & IMAGE_EXTS & ",", "," & ext & ",", vbTextCompare) > 0)
End Function

' 失敗判定(確定仕様準拠: ラッパーはエラーも文字列で返す。台帳§0)。
'   "#ERR:" 始まり = TryRibbonRun側の失敗 / 空・"error" = リボン側の失敗 /
'   上限系メッセージ = E0204の可能性(LooksLikeLimitError)。
Private Function IsVisionError(ByVal s As String) As Boolean
    If LenB(Trim$(s)) = 0 Then
        IsVisionError = True
    ElseIf Left$(s, 5) = "#ERR:" Then
        IsVisionError = True
    ElseIf StrComp(Trim$(s), "error", vbTextCompare) = 0 Then
        IsVisionError = True
    Else
        IsVisionError = modGateway.LooksLikeLimitError(s)
    End If
End Function

Private Function SafeResultToString(ByVal result As Variant) As String
    On Error GoTo Empty0
    If IsError(result) Then Exit Function
    SafeResultToString = CStr(result)
    Exit Function
Empty0:
    SafeResultToString = ""
End Function

Private Function JoinPagesText(pages() As ExtractedPage) As String
    On Error GoTo Empty0
    Dim n As Long
    n = UBound(pages) - LBound(pages) + 1
    If n <= 0 Then Exit Function

    Dim parts() As String
    ReDim parts(0 To n - 1)
    Dim lo As Long
    lo = LBound(pages)
    Dim i As Long
    For i = 0 To n - 1
        parts(i) = pages(lo + i).Text
    Next i
    JoinPagesText = Join(parts, vbLf)
    Exit Function
Empty0:
    JoinPagesText = ""
End Function
