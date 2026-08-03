Attribute VB_Name = "optOcrPage"
Option Explicit

' ============================================================================
' optOcrPage - 画像PDFのページ描画とOCRの実行ループ(opt機能・R14-4a)
' ----------------------------------------------------------------------------
' 役割:
'   optVision から移設した「何ページ描いて、どれをOCRして、いつ画像を捨てるか」
'   の段取りだけを持つ。optVision は Ghostscript の解決・一時フォルダの用意・
'   後始末という「前後」だけを残し、ここへ1回呼び出すだけになった
'   (optVision が28,000字のWARN帯に達したための分割でもある。憲章§4-6)。
'
' なぜバッチにするのか(実機第3報 RC4):
'   旧実装は「上限ページ分を1回のGS起動で全部描く → 全部できてから1枚ずつ
'   OCR」だった。上限を100ページへ上げると、
'     ・画像が全部できるまでOCRが1文字も始まらない(体感が全部止まる)
'     ・100枚ぶんのJPEGが %TEMP% に同時に載る
'     ・ループ中に DoEvents が1つも無く、Excelが完全に無反応になる
'   の3つが同時に悪化する。20ページずつ描き、そのバッチをOCRし終えたら
'   即座に画像を消す形にすれば、最初の20ページは数十秒で読み始められ、
'   一時領域は20枚ぶんで頭打ちになる。
'
' 正直な制約(記録):
'   Application.Run(=AIリボンのChatGPTV)は【同期】呼び出しで、VBAからは
'   中断できない。1ページの応答が返らない限りExcelは固まる。ここでできるのは
'   「1ページ分より長くは固めない」ことだけで、そのために各ページの
'   TryRibbonRun の前後で必ず DoEvents を回し、ページ間で進捗を描き直す。
'   ハード中断(利用者がEscで止める)は構造上できない。
'
' 設計判断:
'   ・GSの出力番号(page_%03d.jpg)は「その実行で出力した順」に1から振られる
'     のが Ghostscript の挙動だが、実機で確かめられない層(憲章§4-3)なので
'     絶対ページ番号で振られる版にも耐えるよう、両方の起点を実在確認してから
'     読む(CountBatchFiles)。
'   ・進捗の文面と分割の算数は optOcrCore(純ロジック)が持ち、ここは
'     副作用(Shell起動・待ち・ファイル削除・ログ)だけを持つ。
'   ・失敗しても例外を外へ出さない。戻り値は本文か "#ERR:E0303:…"。
' ============================================================================

' 1回のGS起動で描かせるページ数。20は「最初のOCRが始まるまでの待ち」と
' 「GS起動の回数」の釣り合いで決めた値(実機第3報の裁定 R14-4a)。
Private Const BATCH_PAGES As Long = 20

' 確定済みリボン関数名(出典: RIBBON_API_CONFIRMED.md §1 #5, #8)。
' 呼び出しは必ず modGateway.TryRibbonRun 経由(R3)。
Private Const RIBBON_FUNC_NAME As String = "ChatGPTV"
Private Const BASE64_FUNC_NAME As String = "Base64FromFile"
Private Const VISION_RESOLUTION As String = "high"
Private Const VISION_TOOL_NAME As String = "マイ本棚AI:vision"

Public Function Ping() As Boolean
    Ping = True
End Function

' ----------------------------------------------------------------------------
' OcrPdfByBatch - PDFを20ページずつ画像化し、そのつどOCRして1本のページ付き
'   テキストにして返す(失敗時は "#ERR:E0303:<何が起きたか+次の一手>")。
'   gsExe / pdfPath / folderPath : optVision が解決済みのもの
'   dpi / maxPages / waitSec     : config 由来(丸めは optOcrCore が持つ)
'   visionPrompt                 : optVision の VISION_PROMPT(文言の一次情報は
'                                  あちら側のまま。ここでは持たない)
'   outKeepWork : True で返したら【作業フォルダを消さないこと】。書きかけの
'                 gs_out.log が唯一の手がかりになるタイムアウト時だけ立つ
'                 (R11-D 監査3 H-1の性質をそのまま引き継ぐ)。
' ----------------------------------------------------------------------------
Public Function OcrPdfByBatch(ByVal gsExe As String, ByVal pdfPath As String, _
                              ByVal folderPath As String, ByVal dpi As Long, _
                              ByVal maxPages As Long, ByVal waitSec As Long, _
                              ByVal visionPrompt As String, _
                              ByRef outKeepWork As Boolean) As String
    Dim pages() As ExtractedPage
    Dim okN As Long: okN = 0
    Dim doneN As Long: doneN = 0          ' OCRを試した実ページ数
    Dim foundTotal As Long: foundTotal = 0 ' 描画できた総枚数(上限+1まで)
    Dim sumMs As Double: sumMs = 0#
    Dim lastFinished As Boolean: lastFinished = True

    outKeepWork = False

    Dim safeMax As Long: safeMax = optOcrCore.SafeMaxPages(maxPages)
    Dim renderCap As Long: renderCap = optOcrCore.RenderCapFor(maxPages)
    Dim batchN As Long: batchN = optOcrCore.BatchCountFor(renderCap, BATCH_PAGES)
    If batchN < 1 Then batchN = 1
    ReDim pages(0 To safeMax - 1)

    ' 総頁の見込み。短いPDFだと最後のバッチで実際の枚数が分かるので、
    ' 分かった時点で分母を本当の値へ寄せる(嘘の分母を出し続けない)。
    Dim estTotal As Long: estTotal = safeMax

    Dim b As Long
    For b = 1 To batchN
        Dim bounds As String
        bounds = optOcrCore.BatchBoundsFor(renderCap, BATCH_PAGES, b)
        If LenB(bounds) = 0 Then Exit For

        Dim firstP As Long: firstP = BoundPart(bounds, 0)
        Dim lastP As Long: lastP = BoundPart(bounds, 1)
        If firstP < 1 Or lastP < firstP Then Exit For

        modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDDBC) & " PDFを画像に変換しています…"
        modShelfBatch.StageBanner "画像化中… バッチ" & b & "/" & batchN

        lastFinished = RenderBatch(gsExe, pdfPath, folderPath, dpi, firstP, lastP, waitSec)

        Dim baseIdx As Long: baseIdx = 1
        Dim gotN As Long
        gotN = CountBatchFiles(folderPath, firstP, lastP, baseIdx)

        ' タイムアウト時の最後の1枚はGSが書いている途中の可能性がある
        ' (2枚以上あるときだけ捨てる。旧実装と同じ考え方)。
        If Not lastFinished And gotN > 1 Then gotN = gotN - 1
        If gotN = 0 Then Exit For

        foundTotal = foundTotal + gotN
        If gotN < (lastP - firstP + 1) Then estTotal = firstP - 1 + gotN
        If estTotal > safeMax Then estTotal = safeMax

        ' このバッチのうち上限内に入るページだけをOCRへ回す
        ' (上限+1枚目は「まだ先がある」ことの証拠にだけ使い、読まない)。
        Dim ocrN As Long: ocrN = gotN
        If firstP - 1 + ocrN > safeMax Then ocrN = safeMax - (firstP - 1)
        If ocrN < 0 Then ocrN = 0

        Dim batchMs As Double: batchMs = 0#
        Dim maxMs As Double: maxMs = 0#

        Dim i As Long
        For i = 0 To ocrN - 1
            Dim pageNo As Long: pageNo = firstP + i

            Dim avgMs As Double: avgMs = 0#
            If doneN >= 2 Then avgMs = sumMs / CDbl(doneN)
            modShelfBatch.StageBanner optOcrCore.OcrPageBanner(pageNo, estTotal, b, batchN, avgMs)
            modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDDBC) & " OCR中… " & _
                pageNo & "/" & estTotal & " ページ"

            ' 読めたページだけ okN が進む(1ページの失敗で資料全体を捨てない)。
            Dim t0 As Double: t0 = Timer
            OcrOnePage folderPath & "\" & optOcrCore.PageJpgName(baseIdx + i), _
                       pageNo, visionPrompt, pages, okN
            Dim ms As Double: ms = modUtilText.ElapsedMsSince(t0)

            doneN = doneN + 1
            sumMs = sumMs + ms
            batchMs = batchMs + ms
            If ms > maxMs Then maxMs = ms
        Next i

        ' このバッチのJPEGは用が済んだら即座に消す(%TEMP%を溜めない)。
        KillBatchJpgs folderPath, baseIdx, gotN

        If ocrN > 0 Then
            On Error Resume Next
            modLog.LogUsage "vision_page_ms", "", "b=" & b & ";pages=" & ocrN & _
                ";avg=" & CLng(batchMs / CDbl(ocrN)) & ";max=" & CLng(maxMs)
            On Error GoTo 0
        End If

        DoEvents
        If gotN < (lastP - firstP + 1) Then Exit For   ' 資料の終端に達した
        If firstP - 1 + gotN >= renderCap Then Exit For
    Next b

    If foundTotal = 0 Then
        OcrPdfByBatch = NoRenderResult(folderPath, pdfPath, lastFinished, waitSec, outKeepWork)
        Exit Function
    End If

    If okN = 0 Then
        OcrPdfByBatch = "#ERR:E0303:ページ画像は作れましたが、文字を読み取れませんでした。" & _
            "しばらく時間を置いてからもう一度お試しください。"
        Exit Function
    End If

    Dim truncated As Boolean
    truncated = optOcrCore.IsTruncatedCount(foundTotal, maxPages)
    If okN < safeMax Then ReDim Preserve pages(0 To okN - 1)
    OcrPdfByBatch = modUtil.JoinPagedText(pages, truncated)
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' 1ページ分のOCR。読めたら pages へ積んで okN を1つ進める。
' TryRibbonRun の前後で必ず DoEvents を回す(§3-1: 押せるものは反応する。
' 同期呼び出しそのものは中断できないので、せめて1ページ境界では画面を返す)。
Private Sub OcrOnePage(ByVal jpgPath As String, ByVal pageNo As Long, _
                       ByVal visionPrompt As String, _
                       ByRef pages() As ExtractedPage, ByRef okN As Long)
    DoEvents
    Dim b64 As String
    b64 = optVision.SafeResultToString(modGateway.TryRibbonRun(BASE64_FUNC_NAME, Array(jpgPath)))
    DoEvents

    If LenB(Trim$(b64)) = 0 Or Left$(b64, 5) = "#ERR:" Then
        On Error Resume Next
        modLog.LogError "E0303", "optOcrPage.OcrPdfByBatch", _
            "p" & pageNo & " Base64失敗: " & modUtil.SafeLeft(b64, 200)
        On Error GoTo 0
        Exit Sub
    End If

    Dim s As String
    s = optVision.SafeResultToString(modGateway.TryRibbonRun(RIBBON_FUNC_NAME, _
        Array(visionPrompt, b64, "", VISION_RESOLUTION, VISION_TOOL_NAME)))
    DoEvents

    If optVision.IsVisionError(s) Then
        On Error Resume Next
        modLog.LogError "E0303", "optOcrPage.OcrPdfByBatch", _
            "p" & pageNo & " 読取失敗: " & modUtil.SafeLeft(s, 200)
        On Error GoTo 0
        Exit Sub
    End If

    pages(okN).page = pageNo
    pages(okN).Text = s
    okN = okN + 1
End Sub

' 1バッチ分だけGSを起動して完了を待つ(Trueで完了・Falseでタイムアウト)。
' 完了フラグは毎回消してから起動する(前のバッチのフラグが残っていると、
' 待ちループが「もう終わっている」と即座に誤判定する)。
Private Function RenderBatch(ByVal gsExe As String, ByVal pdfPath As String, _
                             ByVal folderPath As String, ByVal dpi As Long, _
                             ByVal firstP As Long, ByVal lastP As Long, _
                             ByVal waitSec As Long) As Boolean
    Dim flagPath As String: flagPath = optOcrCore.DoneFlagFor(folderPath)
    KillIfExists flagPath

    Dim runCmd As String
    runCmd = optOcrCore.BuildRunCommand( _
        optOcrCore.BuildGsCommand(gsExe, pdfPath, optOcrCore.OutPatternFor(folderPath), _
                                  dpi, lastP, firstP), _
        flagPath, optOcrCore.GsLogFor(folderPath))

    Dim gsErrNum As Long: gsErrNum = 0
    Dim gsErrDesc As String: gsErrDesc = ""
    If Not optGsProc.RunGsAsync(runCmd, gsErrNum, gsErrDesc) Then
        On Error Resume Next
        modLog.LogError "E0303", "optOcrPage.OcrPdfByBatch", modUtil.SafeLeft( _
            "GS起動失敗 p" & firstP & "-" & lastP & " err#" & gsErrNum & ": " & _
            gsErrDesc & " " & pdfPath, 2000)
        On Error GoTo 0
        Exit Function
    End If

    Dim gsRc As Long: gsRc = -1
    RenderBatch = optGsTxt.WaitForDoneFlag(flagPath, waitSec, gsRc)

    ' 画像は出来ていても終了コードが非0のことがある(GSは軽微な警告でも非0)。
    ' 取込は続けるが、事実は次の調査の手がかりとして必ず残す(R11-D)。
    If gsRc > 0 Then
        On Error Resume Next
        modLog.LogUsage "gs_ocr_rc_nonzero", "", modUtil.SafeLeft( _
            "b(" & firstP & "-" & lastP & ") " & optGsTxt.GsFailureDetail(folderPath), 500)
        On Error GoTo 0
    End If
End Function

' このバッチで実際に何枚できたかを数える。GSの連番の起点(1始まり=その実行の
' 出力順 / firstP始まり=絶対ページ番号)を実在確認で見分けて baseIdx へ返す。
Private Function CountBatchFiles(ByVal folderPath As String, ByVal firstP As Long, _
                                 ByVal lastP As Long, ByRef baseIdx As Long) As Long
    baseIdx = 1
    If Not optVision.PathExists(folderPath & "\" & optOcrCore.PageJpgName(1)) Then
        If Not optVision.PathExists(folderPath & "\" & optOcrCore.PageJpgName(firstP)) Then Exit Function
        baseIdx = firstP
    End If

    Dim want As Long: want = lastP - firstP + 1
    Dim i As Long
    For i = 0 To want - 1
        If Not optVision.PathExists(folderPath & "\" & optOcrCore.PageJpgName(baseIdx + i)) Then Exit For
    Next i
    CountBatchFiles = i
End Function

' 読み終えたバッチのJPEGを消す(失敗しても無視。TEMPなのでOSが後で片付ける)。
Private Sub KillBatchJpgs(ByVal folderPath As String, ByVal baseIdx As Long, ByVal n As Long)
    On Error Resume Next
    Dim i As Long
    For i = 0 To n - 1
        Kill folderPath & "\" & optOcrCore.PageJpgName(baseIdx + i)
    Next i
    On Error GoTo 0
End Sub

Private Sub KillIfExists(ByVal p As String)
    On Error Resume Next
    If optVision.PathExists(p) Then Kill p
    On Error GoTo 0
End Sub

' "21|40" の n 番目(0始まり)を数値で取り出す。
Private Function BoundPart(ByVal bounds As String, ByVal idx As Long) As Long
    Dim parts() As String: parts = Split(bounds, "|")
    If idx < 0 Or idx > UBound(parts) - LBound(parts) Then Exit Function
    BoundPart = CLng(Val(parts(LBound(parts) + idx)))
End Function

' 1枚も描けなかったときの戻り値と記録(旧 optVision の「描画0枚」経路)。
' タイムアウトのときだけ作業フォルダを残す(書きかけの gs_out.log と出力が
' 原因究明の唯一の材料で、消すと二度と調べられない。R11-D 監査3 H-1)。
Private Function NoRenderResult(ByVal folderPath As String, ByVal pdfPath As String, _
                                ByVal finished As Boolean, ByVal waitSec As Long, _
                                ByRef outKeepWork As Boolean) As String
    Dim workNote As String: workNote = ""
    If Not finished Then
        outKeepWork = True
        workNote = " work=" & folderPath
    End If

    On Error Resume Next
    modLog.LogError "E0303", "optOcrPage.OcrPdfByBatch", modUtil.SafeLeft( _
        "描画0枚 finished=" & finished & " " & optGsTxt.GsFailureDetail(folderPath) & _
        workNote & " " & modUtil.SafeLeft(pdfPath, 200), 2000)
    On Error GoTo 0

    If finished Then
        NoRenderResult = "#ERR:E0303:このPDFからページ画像を作れませんでした。" & _
            "ファイルが壊れているか、パスワードで保護されている可能性があります。"
    Else
        NoRenderResult = "#ERR:E0303:PDFの画像変換が" & waitSec & "秒以内に" & _
            "終わりませんでした。ページ数の少ないPDFに分けるか、config の " & _
            "gs_abs_timeout_sec を大きくしてからお試しください。"
    End If
End Function
