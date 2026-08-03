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

' RenderBatch の結果(2026-08-03 R14-F1/F2)。「起動できなかった」と
' 「起動はしたが終わらなかった」を Boolean 1つで混ぜていたため、GSが1度も
' 起動していない資料に対して「1200秒以内に終わりませんでした」という
' 事実と違う案内を出していた。3値にして最後まで別の事実として扱う。
Private Const RB_DONE As Long = 0          ' 完了フラグが出た
Private Const RB_LAUNCH_FAIL As Long = 1   ' GSを起動できなかった
Private Const RB_TIMEOUT As Long = 2       ' 起動したが待ち時間内に終わらなかった

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
'   dpi / maxPages / absSec      : config 由来(丸めは optOcrCore が持つ)。
'                                  absSec は gs_abs_timeout_sec で、【1資料
'                                  あたり】の画像化待ちの絶対上限(R14-F6)。
'                                  バッチごとに残り時間だけを次の待ちへ渡す。
'   visionPrompt                 : optVision の VISION_PROMPT(文言の一次情報は
'                                  あちら側のまま。ここでは持たない)
'   outKeepWork : True で返したら【作業フォルダを消さないこと】。
'                 (a) 1枚も描けずタイムアウトした(書きかけの gs_out.log が
'                     唯一の手がかり。R11-D 監査3 H-1)
'                 (b) 時間切れなのにGSを止められなかった(PID不明)。生きている
'                     GSが書き込んでいるフォルダを消してはならない(R14-F5)
'   outAborted  : True で返したら【まだ先があるのに途中で止めた】。上限まで
'                 読み切ったのか、変換エラー/時間切れで欠けたのかを呼び出し元
'                 (optVision)がカードのメモで言い分けるための唯一の材料
'                 (R14-F2。従来はどちらも同じ形で返り、欠落が無言だった)。
' ----------------------------------------------------------------------------
Public Function OcrPdfByBatch(ByVal gsExe As String, ByVal pdfPath As String, _
                              ByVal folderPath As String, ByVal dpi As Long, _
                              ByVal maxPages As Long, ByVal absSec As Long, _
                              ByVal visionPrompt As String, _
                              ByRef outKeepWork As Boolean, _
                              ByRef outAborted As Boolean) As String
    Dim pages() As ExtractedPage
    Dim okN As Long: okN = 0
    Dim doneN As Long: doneN = 0          ' OCRを試した実ページ数
    Dim foundTotal As Long: foundTotal = 0 ' 描画できた総枚数(上限+1まで)
    Dim sumMs As Double: sumMs = 0#
    Dim usedSec As Long: usedSec = 0      ' 画像化待ちに使った秒数の累計(F6)
    Dim aborted As Boolean: aborted = False
    Dim launchFailed As Boolean: launchFailed = False

    outKeepWork = False
    outAborted = False

    ' R14-F12: 予期しない実行時エラーでも契約(本文 or "#ERR:E0303:…")を守る。
    On Error GoTo Failed

    Dim safeMax As Long: safeMax = optOcrCore.SafeMaxPages(maxPages)
    Dim renderCap As Long: renderCap = optOcrCore.RenderCapFor(maxPages)
    Dim batchN As Long: batchN = optOcrCore.BatchCountFor(renderCap, BATCH_PAGES)
    If batchN < 1 Then batchN = 1
    ReDim pages(0 To safeMax - 1)

    ' 総頁は「終端に達した」か「上限+1枚目が出た」瞬間にしか確定しない。
    ' 確定するまでは0のままにして、分母もバッチ総数も表示しない(R14-F9)。
    Dim knownTotal As Long: knownTotal = 0

    Dim b As Long
    For b = 1 To batchN
        Dim bounds As String
        bounds = optOcrCore.BatchBoundsFor(renderCap, BATCH_PAGES, b)
        If LenB(bounds) = 0 Then Exit For

        Dim firstP As Long: firstP = BoundPart(bounds, 0)
        Dim lastP As Long: lastP = BoundPart(bounds, 1)
        If firstP < 1 Or lastP < firstP Then Exit For

        ' 1資料あたりの絶対上限の【残り】だけを次の待ちへ渡す(R14-F6)。
        ' 使い切っていたら、そこで正直に打ち切る(=中断扱い)。
        Dim waitSec As Long: waitSec = optOcrCore.RemainingWaitSec(absSec, usedSec)
        If waitSec <= 0 Then
            aborted = True
            Exit For
        End If

        modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDDBC) & " PDFを画像に変換しています…"
        modShelfBatch.StageBanner "画像化中… " & BatchLabel(b, knownTotal)

        Dim usedThis As Long: usedThis = 0
        Dim killed As Boolean: killed = False
        Dim rb As Long
        rb = RenderBatch(gsExe, pdfPath, folderPath, dpi, firstP, lastP, waitSec, _
                         usedThis, killed)
        usedSec = usedSec + usedThis

        If rb = RB_LAUNCH_FAIL Then
            launchFailed = True
            Exit For
        End If

        Dim baseIdx As Long: baseIdx = 1
        Dim gotN As Long
        gotN = CountBatchFiles(folderPath, firstP, lastP, baseIdx)

        If rb = RB_TIMEOUT Then
            aborted = True
            ' 止められなかった(PID不明)なら、GSがまだ書いているフォルダを
            ' 消してはならない(R14-F5。optGsTxtのタイムアウトと同じ扱い)。
            If Not killed Then outKeepWork = True
            ' 最後の1枚はGSが書いている途中の可能性がある
            ' (2枚以上あるときだけ捨てる。旧実装と同じ考え方)。
            If gotN > 1 Then gotN = gotN - 1
        End If
        If gotN = 0 Then Exit For

        foundTotal = foundTotal + gotN

        ' ここで初めて総頁が確定する2つの場面(R14-F9)。
        Dim isLastBatch As Boolean: isLastBatch = (gotN < (lastP - firstP + 1))
        If isLastBatch Then
            knownTotal = firstP - 1 + gotN
            If knownTotal > safeMax Then knownTotal = safeMax
        ElseIf firstP - 1 + gotN >= renderCap Then
            knownTotal = safeMax          ' 上限で切る=総頁は上限ぶん取り込む
        End If

        ' このバッチのうち上限内に入るページだけをOCRへ回す
        ' (上限+1枚目は「まだ先がある」ことの証拠にだけ使い、読まない)。
        Dim ocrN As Long: ocrN = gotN
        If firstP - 1 + ocrN > safeMax Then ocrN = safeMax - (firstP - 1)
        If ocrN < 0 Then ocrN = 0

        Dim shownBatches As Long: shownBatches = 0
        If knownTotal > 0 Then shownBatches = optOcrCore.BatchCountFor(knownTotal, BATCH_PAGES)

        Dim batchMs As Double: batchMs = 0#
        Dim maxMs As Double: maxMs = 0#
        Dim failN As Long: failN = 0
        Dim firstErr As String: firstErr = ""

        Dim i As Long
        For i = 0 To ocrN - 1
            Dim pageNo As Long: pageNo = firstP + i

            Dim avgMs As Double: avgMs = 0#
            If doneN >= 2 Then avgMs = sumMs / CDbl(doneN)
            Dim banner As String
            banner = optOcrCore.OcrPageBanner(pageNo, knownTotal, b, shownBatches, avgMs)
            modShelfBatch.StageBanner banner
            modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDDBC) & " " & banner

            ' 読めたページだけ okN が進む(1ページの失敗で資料全体を捨てない)。
            Dim t0 As Double: t0 = Timer
            OcrOnePage folderPath & "\" & optOcrCore.PageJpgName(baseIdx + i), _
                       pageNo, visionPrompt, pages, okN, failN, firstErr
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
            On Error GoTo Failed
        End If

        ' R14-F10: 頁ごとに1行ずつE0303を書くと、20頁失敗しただけで err_log が
        ' 埋まり、直前の本当の原因が上限行数のローテで流れる。vision_page_ms と
        ' 同じくバッチ単位で1行(件数+最初の1件)に畳む。
        If failN > 0 Then
            On Error Resume Next
            modLog.LogError "E0303", "optOcrPage.OcrPdfByBatch", modUtil.SafeLeft( _
                "b=" & b & " p" & firstP & "-" & (firstP + ocrN - 1) & " 読取失敗 " & _
                failN & "/" & ocrN & "頁 first=" & firstErr, 2000)
            On Error GoTo Failed
        End If

        DoEvents
        If aborted Then Exit For
        If isLastBatch Then Exit For                   ' 資料の終端に達した
        If firstP - 1 + gotN >= renderCap Then Exit For
    Next b

    If foundTotal = 0 Then
        OcrPdfByBatch = NoRenderResult(folderPath, pdfPath, launchFailed, aborted, _
                                       absSec, outKeepWork)
        Exit Function
    End If

    If okN = 0 Then
        OcrPdfByBatch = "#ERR:E0303:ページ画像は作れましたが、文字を読み取れませんでした。" & _
            "しばらく時間を置いてからもう一度お試しください。"
        Exit Function
    End If

    ' R14-F2: 途中で止めた場合は【必ず】打ち切り扱いにする。ここをFalseで
    ' 返していたため、20頁で中断した100頁のPDFが status=done として
    ' 「全部入った」顔で本棚に並んでいた(欠落が無言=憲章§4-1違反)。
    Dim truncated As Boolean
    truncated = optOcrCore.IsTruncatedCount(foundTotal, maxPages)
    If aborted Or launchFailed Then
        truncated = True
        outAborted = True
    End If
    If okN < safeMax Then ReDim Preserve pages(0 To okN - 1)
    OcrPdfByBatch = modUtil.JoinPagedText(pages, truncated)
    Exit Function

Failed:
    ' ハンドラ稼働中は On Error Resume Next が効かない。Resumeで抜けてから
    ' 記録する(opt層GS系モジュール共通の作法)。outKeepWork は途中で立てた
    ' 値をそのまま残す(生きているGSのフォルダを消させないため)。
    Dim failNum As Long: failNum = Err.Number
    Dim failDesc As String: failDesc = Err.Description
    Resume BatchCleanup
BatchCleanup:
    On Error Resume Next
    modLog.LogError "E0303", "optOcrPage.OcrPdfByBatch", modUtil.SafeLeft( _
        "err#" & failNum & ": " & failDesc & " " & pdfPath, 2000)
    On Error GoTo 0
    OcrPdfByBatch = "#ERR:E0303:画像PDFの読み取り中に問題が発生しました" & _
        "(err#" & failNum & ")。もう一度お試しください。"
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' 1ページ分のOCR。読めたら pages へ積んで okN を1つ進める。
' TryRibbonRun の前後で必ず DoEvents を回す(§3-1: 押せるものは反応する。
' 同期呼び出しそのものは中断できないので、せめて1ページ境界では画面を返す)。
' R14-F10: 失敗はここでは【数えるだけ】。err_log への記録は呼び出し元が
' バッチ単位で1行にまとめる(failN / firstErr がそのための受け皿)。
Private Sub OcrOnePage(ByVal jpgPath As String, ByVal pageNo As Long, _
                       ByVal visionPrompt As String, _
                       ByRef pages() As ExtractedPage, ByRef okN As Long, _
                       ByRef failN As Long, ByRef firstErr As String)
    DoEvents
    Dim b64 As String
    b64 = optVision.SafeResultToString(modGateway.TryRibbonRun(BASE64_FUNC_NAME, Array(jpgPath)))
    DoEvents

    If LenB(Trim$(b64)) = 0 Or Left$(b64, 5) = "#ERR:" Then
        NotePageFail failN, firstErr, "p" & pageNo & " Base64失敗: " & modUtil.SafeLeft(b64, 200)
        Exit Sub
    End If

    Dim s As String
    s = optVision.SafeResultToString(modGateway.TryRibbonRun(RIBBON_FUNC_NAME, _
        Array(visionPrompt, b64, "", VISION_RESOLUTION, VISION_TOOL_NAME)))
    DoEvents

    If optVision.IsVisionError(s) Then
        NotePageFail failN, firstErr, "p" & pageNo & " 読取失敗: " & modUtil.SafeLeft(s, 200)
        Exit Sub
    End If

    pages(okN).page = pageNo
    pages(okN).Text = s
    okN = okN + 1
End Sub

' 頁OCRの失敗を数え、最初の1件だけ理由を控える(R14-F10)。
Private Sub NotePageFail(ByRef failN As Long, ByRef firstErr As String, ByVal note As String)
    failN = failN + 1
    If LenB(firstErr) = 0 Then firstErr = note
End Sub

' 進捗バナーのバッチ表示。総頁が確定するまで「/B」を出さない(R14-F9)。
Private Function BatchLabel(ByVal batchIdx As Long, ByVal knownTotal As Long) As String
    BatchLabel = "バッチ" & batchIdx
    If knownTotal > 0 Then
        BatchLabel = BatchLabel & "/" & optOcrCore.BatchCountFor(knownTotal, BATCH_PAGES)
    End If
End Function

' 1バッチ分だけGSを起動して完了を待つ。戻り値は RB_DONE / RB_LAUNCH_FAIL /
' RB_TIMEOUT(R14-F1: 起動できなかったことを時間切れと混ぜない)。
' outUsedSec : この待ちに実際に使った秒数(絶対上限の残り計算用。R14-F6)
' outKilled  : タイムアウト時にGSを止められたか(R14-F5)。止められなかった
'              ときだけ、呼び出し元が作業フォルダを残す判断をする。
' 完了フラグは毎回消してから起動する(前のバッチのフラグが残っていると、
' 待ちループが「もう終わっている」と即座に誤判定する)。
Private Function RenderBatch(ByVal gsExe As String, ByVal pdfPath As String, _
                             ByVal folderPath As String, ByVal dpi As Long, _
                             ByVal firstP As Long, ByVal lastP As Long, _
                             ByVal waitSec As Long, ByRef outUsedSec As Long, _
                             ByRef outKilled As Boolean) As Long
    outUsedSec = 0
    outKilled = False

    Dim flagPath As String: flagPath = optOcrCore.DoneFlagFor(folderPath)
    KillIfExists flagPath

    Dim runCmd As String
    runCmd = optOcrCore.BuildRunCommand( _
        optOcrCore.BuildGsCommand(gsExe, pdfPath, optOcrCore.OutPatternFor(folderPath), _
                                  dpi, lastP, firstP), _
        flagPath, optOcrCore.GsLogFor(folderPath))

    Dim gsErrNum As Long: gsErrNum = 0
    Dim gsErrDesc As String: gsErrDesc = ""
    ' R14-F5: PIDを控えて起動する(optGsTxtと同じ4引数の呼び方)。PIDが無いと
    ' 真のハングでGS/cmd.exeがログオン中ずっと残り、書きかけのフォルダも
    ' 消せない。WMIが塞がれている端末では0のまま返る(そのときは止めない)。
    Dim gsPid As Long: gsPid = 0
    If Not optGsProc.RunGsAsync(runCmd, gsErrNum, gsErrDesc, gsPid) Then
        On Error Resume Next
        modLog.LogError "E0303", "optOcrPage.OcrPdfByBatch", modUtil.SafeLeft( _
            "GS起動失敗 p" & firstP & "-" & lastP & " err#" & gsErrNum & ": " & _
            gsErrDesc & " " & pdfPath, 2000)
        On Error GoTo 0
        RenderBatch = RB_LAUNCH_FAIL
        Exit Function
    End If

    Dim t0 As Double: t0 = Timer
    Dim gsRc As Long: gsRc = -1
    Dim finished As Boolean
    finished = optGsTxt.WaitForDoneFlag(flagPath, waitSec, gsRc)
    outUsedSec = CLng(modUtilText.ElapsedMsSince(t0) / 1000#)

    ' 画像は出来ていても終了コードが非0のことがある(GSは軽微な警告でも非0)。
    ' 取込は続けるが、事実は次の調査の手がかりとして必ず残す(R11-D)。
    If gsRc > 0 Then
        On Error Resume Next
        modLog.LogUsage "gs_ocr_rc_nonzero", "", modUtil.SafeLeft( _
            "b(" & firstP & "-" & lastP & ") " & optGsTxt.GsFailureDetail(folderPath), 500)
        On Error GoTo 0
    End If

    If finished Then Exit Function        ' RB_DONE(=0)

    ' R14-F5: 掃除を決める【前】に止める。順序が逆だと、生きているGSが
    ' 書き込んでいるフォルダを消しにいくことになる(optGsTxtと同じ順序)。
    If gsPid > 0 Then
        optGsProc.KillGsTree gsPid
        outKilled = True
    End If

    On Error Resume Next
    modLog.LogError "E0303", "optOcrPage.OcrPdfByBatch", modUtil.SafeLeft( _
        "GS_TIMEOUT p" & firstP & "-" & lastP & " wait=" & waitSec & "s pid=" & gsPid & _
        " killed=" & outKilled & " " & optGsTxt.GsFailureDetail(folderPath) & _
        " " & modUtil.SafeLeft(pdfPath, 200), 2000)
    On Error GoTo 0
    RenderBatch = RB_TIMEOUT
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
' 3つの事実を最後まで別のものとして扱う(R14-F1):
'   launchFailed : GSを【起動できなかった】。待ってすらいないので秒数の話を
'                  してはならない(従来はここも「1200秒以内に終わりません
'                  でした」と言い、config の gs_abs_timeout_sec を大きくする
'                  よう案内していた=何度やっても直らない案内)。調べる材料も
'                  何も出ていないので作業フォルダは残さない。
'   aborted      : 起動はしたが時間切れ。書きかけの gs_out.log が原因究明の
'                  唯一の材料なので作業フォルダを残す(R11-D 監査3 H-1)。
'   それ以外     : GSは正常終了したのに1枚も出なかった(壊れ・パスワード)。
Private Function NoRenderResult(ByVal folderPath As String, ByVal pdfPath As String, _
                                ByVal launchFailed As Boolean, ByVal aborted As Boolean, _
                                ByVal absSec As Long, ByRef outKeepWork As Boolean) As String
    If launchFailed Then
        outKeepWork = False
        On Error Resume Next
        modLog.LogError "E0303", "optOcrPage.OcrPdfByBatch", modUtil.SafeLeft( _
            "描画0枚 GS起動失敗 " & modUtil.SafeLeft(pdfPath, 200), 2000)
        On Error GoTo 0
        NoRenderResult = "#ERR:E0303:PDFを画像に変換する処理を開始できませんでした。" & _
            "config の ghostscript_path が正しいかご確認ください。"
        Exit Function
    End If

    Dim workNote As String: workNote = ""
    If aborted Then
        outKeepWork = True
        workNote = " work=" & folderPath
    End If

    On Error Resume Next
    modLog.LogError "E0303", "optOcrPage.OcrPdfByBatch", modUtil.SafeLeft( _
        "描画0枚 aborted=" & aborted & " " & optGsTxt.GsFailureDetail(folderPath) & _
        workNote & " " & modUtil.SafeLeft(pdfPath, 200), 2000)
    On Error GoTo 0

    If aborted Then
        NoRenderResult = "#ERR:E0303:PDFの画像変換が" & absSec & "秒以内に" & _
            "終わりませんでした。ページ数の少ないPDFに分けるか、config の " & _
            "gs_abs_timeout_sec を大きくしてからお試しください。"
    Else
        NoRenderResult = "#ERR:E0303:このPDFからページ画像を作れませんでした。" & _
            "ファイルが壊れているか、パスワードで保護されている可能性があります。"
    End If
End Function
