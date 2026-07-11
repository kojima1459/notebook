Attribute VB_Name = "optVision"
Option Explicit

' ============================================================================
' optVision - 画像PDF/画像ファイルの文字抽出(opt機能・MASTER_SPEC §7.7)
' ----------------------------------------------------------------------------
' 役割:
'   通常抽出(modExtractor)がE0303(画像PDF・文字が取れない)を返した資料や、
'   画像ファイル(png/jpg)そのものについて、AIリボンのVision機能
'   (ChatGPTV仮定)へ渡して文字起こしする。この層が無くても製品は成立し、
'   modules.jsonから1行削除するだけで撤去できる(§7.7)。
'
' 設計判断:
'   ・PDFをページ画像へ変換する手段がVBA単体に無いため、v1実装は
'     2つの経路を両方実装し、config vision_pdf_direct(既定FALSE)で切替える:
'       (a) 楽観路線(vision_pdf_direct=TRUE): PDFパスをそのままChatGPTVに
'           渡せると仮定して直接呼ぶ。
'       (b) 現実路線(既定・vision_pdf_direct=FALSE): PDFの直接処理は行わず、
'           画像ファイル(png/jpg/jpeg)のみ対応する。PDFが渡された場合は
'           丁寧な案内文を含む失敗として扱う。
'     どちらの仮定が崩れても、呼び出し面(ExtractImagePdf/ExtractImagePdfText
'     の引数・戻り値の型)は一切変わらない設計にしてある。
'   ・リボン呼び出しは modGateway.TryRibbonRun のみを使う(R3)。
'   ・失敗しても例外を外に出さない。ExtractImagePdfはBoolean(§7.7契約の
'     とおり)、ExtractImagePdfTextは "#ERR:..." 文字列 または 全文。
'   ・ExtractImagePdfの失敗理由をExtractImagePdfTextへ伝えるため、
'     モジュール内Private変数(mLastErrorMsg)に直近の失敗メッセージを
'     保持する(VBAは単一スレッドで実行されるため競合の心配がない)。
'   ・modUIMain.SetStage 以外のコア参照は行わない(基盤層のみ参照可・§7.7)。
' ============================================================================

' === SIGNATURE ASSUMPTION ===
'   仮定シグネチャ: ChatGPTV(prompt As String, imagePath As String) -> String
'     成功時: 画像/PDFから読み取った全文を1つの文字列で返す。
'     失敗時: 空文字列 または "error" を含む文字列を返す
'     (ChatGPT/GetEmbeddingsと同じ「エラーも例外を投げず文字列で返す」設計を
'     踏襲していると仮定)。
'   根拠: MASTER_SPEC §7.7 optVision.bas の記載「ChatGPTV仮定: (prompt, imagePath)」
'     のみが根拠。引数の順序(prompt先/画像パス先)、画像パスが絶対パス文字列で
'     良いか(Base64エンコード等が要るか)、複数ページ画像の一括指定が
'     可能かは未確認。
'   仕様回答が来たら直す行:
'     ・引数の数/順序が違う場合 -> BuildVisionArgs() の Array(...) 部分のみ修正。
'     ・戻り値の成功/失敗判定基準が違う場合 -> IsVisionError() の判定式のみ修正。
'     ・関数名自体が違う場合(ChatGPTV以外) -> Private Const RIBBON_FUNC_NAME
'       の値のみ修正。
'     ・PDFをページ画像化する手段が別途判明した場合(例: 別のリボン関数で
'       PDF→PNG変換ができる) -> ConvertPdfToImages() 相当の関数を新設し、
'       ExtractImagePdf内のPDF分岐を丸ごと差し替える(現状はプレースホルダの
'       まま、vision_pdf_direct設定による2路線の切替のみ)。
' ================================================================================

Private Const RIBBON_FUNC_NAME As String = "ChatGPTV"
Private Const IMAGE_EXTS As String = "png,jpg,jpeg"
Private Const VISION_PROMPT As String = _
    "この画像(またはPDF)に写っている文字を、レイアウトの意味を保ったまま、" & _
    "省略せずにすべて書き出してください。図表の説明文やキャプションも含めてください。"

Private mLastErrorMsg As String

Public Function Ping() As Boolean
    Ping = True
End Function

' ----------------------------------------------------------------------------
' ExtractImagePdf - path(画像 or PDF)からVisionで文字を読み取り、
'   ExtractedPage 1件(page=1、全文)としてpagesへ格納する(MASTER_SPEC §7.7)。
'   成功: True(pagesに1件) / 失敗: False(pagesは空配列のまま)。
' ----------------------------------------------------------------------------
Public Function ExtractImagePdf(ByVal path As String, ByRef pages() As ExtractedPage) As Boolean
    mLastErrorMsg = ""
    ReDim pages(0 To -1)

    On Error GoTo Fail

    Dim ext As String
    ext = modUtil.ExtOf(path)

    Dim allowDirectPdf As Boolean
    allowDirectPdf = modConfig.GetBool("vision_pdf_direct", False)

    If Not IsImageExt(ext) And Not (ext = "pdf" And allowDirectPdf) Then
        If ext = "pdf" Then
            mLastErrorMsg = "#ERR:E0303:" & modLog.FriendlyMessage("E0303") & _
                "(画像PDFの直接処理は現在この設定では無効です。管理者がconfigの" & _
                "vision_pdf_directをTRUEにすると有効になります。)"
        Else
            mLastErrorMsg = "#ERR:E0301:このファイル形式(" & ext & ")はVisionで直接読み取れません。" & _
                "対応形式は画像(png/jpg/jpeg)" & IIf(allowDirectPdf, "とpdf", "") & "です。"
        End If
        modLog.LogError "E0303", "optVision.ExtractImagePdf", "未対応ext=" & ext & " path=" & modUtil.SafeLeft(path, 300)
        ExtractImagePdf = False
        Exit Function
    End If

    modUIMain.SetStage "🖼 画像から文字を読み取っています…"

    Dim result As Variant
    result = modGateway.TryRibbonRun(RIBBON_FUNC_NAME, BuildVisionArgs(path))

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
        ExtractImagePdfText = "#ERR:E0303:画像/PDFからの文字抽出に失敗しました"
    End If
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー(すべてPrivate: optVisionの公開契約はPing/ExtractImagePdf/
' ExtractImagePdfTextのみ)
' ----------------------------------------------------------------------------

Private Function IsImageExt(ByVal ext As String) As Boolean
    IsImageExt = (InStr(1, "," & IMAGE_EXTS & ",", "," & ext & ",", vbTextCompare) > 0)
End Function

' === SIGNATURE ASSUMPTION: 引数配列の組み立て。仕様が判明したらここだけ直す ===
Private Function BuildVisionArgs(ByVal path As String) As Variant
    BuildVisionArgs = Array(VISION_PROMPT, path)
End Function

' === SIGNATURE ASSUMPTION: 失敗判定。仕様が判明したらここだけ直す ===
Private Function IsVisionError(ByVal s As String) As Boolean
    If LenB(Trim$(s)) = 0 Then
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
