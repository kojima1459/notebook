Attribute VB_Name = "optOcrPage"
Option Explicit

' ============================================================================
' optOcrPage - 画像PDFのページ描画とOCRの実行ループ(opt機能・R14-4a)
' ----------------------------------------------------------------------------
' 役割:
'   optVision から移設した「何ページ描いて、どれをOCRして、いつ画像を捨てるか」
'   の段取りだけを持つ。optVision は Ghostscript の解決・一時フォルダ・後始末
'   という「前後」だけを残し、ここへ1回呼び出すだけになった(optVisionが
'   28,000字のWARN帯に達したための分割でもある。憲章§4-6)。
'
' なぜバッチにするのか(実機第3報 RC4): 旧実装は「上限ページ分を1回で全部描く
'   → 全部できてから1枚ずつOCR」で、上限を上げるほど (a)OCRが始まらない待ち
'   (b)%TEMP%へ同時に載るJPEG枚数 (c)無反応、が同時に悪化した。
'
' 正直な制約(記録): Application.Run(=ChatGPTV)は【同期】でVBAから中断でき
'   ない。1ページの応答が返るまでExcelは固まる。できるのは「1ページ分より長くは
'   固めない」ことだけ(TryRibbonRunの前後でDoEvents)。2026-08-04(R15-6b): 進捗
'   バナーの中断ボタンが modShelfBatch のフラグを立て、ここは頁・バッチ境界で
'   それを読んで止まる。止めたことは黙らず、カードのメモに残す。
'
' 設計判断:
'   ・GSの出力番号(page_%03d.jpg)は「その実行の出力順」に1から振られるのが
'     Ghostscript の挙動だが、実機で確かめられない層(憲章§4-3)なので、絶対
'     ページ番号で振る版にも耐えるよう両方の起点を実在確認する(CountBatchFiles)。
'   ・進捗の文面と分割の算数は optOcrCore/optOcrEta(純ロジック)が持ち、
'     ここは副作用(Shell起動・待ち・ファイル削除・ログ)だけを持つ。
'   ・失敗しても例外を外へ出さない。戻り値は本文か "#ERR:E0303:…"。
' ============================================================================

' 1回のGS起動で描かせるページ数。20は「最初のOCRが始まるまでの待ち」と
' 「GS起動の回数」の釣り合いで決めた(実機第3報の裁定 R14-4a)。
Private Const BATCH_PAGES As Long = 20

' R15-4d(RC10): 描画前に消す残骸のパターン。命名規約は optOcrCore.PageJpgName
' が持つ(あちらは容量が無い)。必ず対で直すこと。
Private Const PAGE_JPG_GLOB As String = "page_*.jpg"

' R15-4b(RC6): 「AI利用の上限」らしい頁失敗がこれだけ連続したら打ち切る
' (modEmbed の consecutiveFail と同型)。上限に当たった状態で残りを叩いても
' 全部失敗するだけで、利用者の時間とAPIの枠を捨てる。
Private Const LIMIT_STREAK_MAX As Long = 3

' R15-5c: 1頁あたり所要ミリ秒の過去実績(ui_state)。当該資料の実測が2頁
' 貯まるまではこの値でETAを出す(初回端末は従来どおり無表示)。
Private Const RATE_KEY As String = "ocr_avg_page_ms"

' RenderBatch の結果(R14-F1/F2)。「起動できなかった」と「起動はしたが終わら
' なかった」を Boolean 1つで混ぜ、GSが1度も起動していない資料へ時間切れの
' 案内を出していた。3値で別の事実として扱う。
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
'                                  absSec = gs_abs_timeout_sec は【1資料あたり】
'                                  の画像化待ちの絶対上限(R14-F6)。バッチには
'                                  残り時間だけを渡す。
'   visionPrompt                 : optVision の VISION_PROMPT(文言の一次情報は
'                                  あちら側のまま。ここでは持たない)
'   outKeepWork : True で返したら【作業フォルダを消さないこと】。(a)1枚も描けず
'                 タイムアウト(書きかけの gs_out.log が唯一の手がかり。R11-D
'                 監査3 H-1) (b)時間切れなのにGSを止められなかった(PID不明。
'                 生きているGSが書いているフォルダは消せない。R14-F5)
'   outAborted  : True で返したら【まだ先があるのに途中で止めた】(R14-F2。
'                 従来はどちらも同じ形で返り、欠落が無言だった)。
'   gsTotalPages: 取込前に判明している総頁数(0=不明。R15-5a)。txtwriteの
'                 分類パスがGSに出させた "Processing pages 1 through N" から
'                 取れる。従来 knownTotal は最終バッチまで0のままで、24頁中
'                 20頁が終わるまでETAが出なかった(RC2)。0なら従来どおり
'                 最終バッチでの確定に自然縮退する。
'   outMemo     : 本棚カードへ出す正直なメモ。再開の冒頭文・中断/上限/頁欠けの
'                 理由・設定上限の説明を、ここで【1回だけ】組み立てて返す
'                 (R15-FixA FA-4。理由を知っているのはここだけなので、
'                 optVision.OcrCapMemo は受け取った文をそのまま渡すだけ)。
' ----------------------------------------------------------------------------
Public Function OcrPdfByBatch(ByVal gsExe As String, ByVal pdfPath As String, _
                              ByVal folderPath As String, ByVal dpi As Long, _
                              ByVal maxPages As Long, ByVal absSec As Long, _
                              ByVal visionPrompt As String, _
                              ByVal gsTotalPages As Long, _
                              ByRef outKeepWork As Boolean, _
                              ByRef outAborted As Boolean, _
                              ByRef outMemo As String) As String
    Dim pages() As ExtractedPage
    Dim okN As Long: okN = 0
    Dim doneN As Long: doneN = 0          ' OCRを試した実ページ数
    Dim foundTotal As Long: foundTotal = 0 ' 描画できた総枚数(上限+1まで)
    Dim sumMs As Double: sumMs = 0#
    Dim usedSec As Long: usedSec = 0      ' 画像化待ちに使った秒数の累計(F6)
    Dim aborted As Boolean: aborted = False
    Dim launchFailed As Boolean: launchFailed = False
    ' R15-4a/4b/6b: 打ち切りの【理由】。"" = 打ち切っていない。
    Dim abortReason As String: abortReason = ""
    Dim limitStreak As Long: limitStreak = 0   ' 上限らしき失敗の連続回数
    Dim allFailN As Long: allFailN = 0         ' 読み取れなかった頁の総数
    ' R15-5c: 過去実績のレート(1頁あたりms)。当該資料の実測が2頁貯まるまで
    ' これでETAを出し、バッチ完了ごとにブレンドして書き戻す。
    Dim rateMs As Double: rateMs = PriorPageMs()

    outKeepWork = False
    outAborted = False
    outMemo = ""

    ' R14-F12: 予期しない実行時エラーでも契約(本文 or "#ERR:E0303:…")を守る。
    On Error GoTo Failed

    Dim safeMax As Long: safeMax = optOcrCore.SafeMaxPages(maxPages)
    Dim renderCap As Long: renderCap = optOcrCore.RenderCapFor(maxPages)
    Dim batchN As Long: batchN = optOcrCore.BatchCountFor(renderCap, BATCH_PAGES)
    If batchN < 1 Then batchN = 1
    ReDim pages(0 To safeMax - 1)

    ' 総頁は「終端に達した」か「上限+1枚目が出た」瞬間にしか確定しない。確定
    ' するまでは0のままにして分母もバッチ総数も出さない(R14-F9)。R15-5a: ただし
    ' 取込前に総頁が分かっている(txtwriteの分類パスのログから読めた)なら最初の
    ' バッチからそれを分母にする。上限より多い資料は上限で頭打ちにする。
    Dim knownTotal As Long: knownTotal = 0
    If gsTotalPages > 0 Then
        knownTotal = gsTotalPages
        If knownTotal > safeMax Then knownTotal = safeMax
    End If

    Dim b As Long
    For b = 1 To batchN
        ' R15-6b: バッチ境界での中断確認(読めた頁は手元に残り partial になる)。
        If CancelWanted() Then
            aborted = True
            abortReason = "cancel"
            Exit For
        End If

        Dim bounds As String
        bounds = optOcrCore.BatchBoundsFor(renderCap, BATCH_PAGES, b)
        If LenB(bounds) = 0 Then Exit For

        Dim firstP As Long: firstP = BoundPart(bounds, 0)
        Dim lastP As Long: lastP = BoundPart(bounds, 1)
        If firstP < 1 Or lastP < firstP Then Exit For

        ' 1資料あたりの絶対上限の【残り】だけを次の待ちへ渡す(R14-F6)。
        ' R15-FixA(FA-5iii): 残り予算をそのまま1バッチへ渡すと1バッチ目のハングが
        ' 予算を全部食い潰すので、1バッチの上限も併せて掛ける。使い切っていたら
        ' そこで正直に打ち切る(=中断扱い)。
        Dim waitSec As Long
        waitSec = optOcrEta.BatchWaitSec(absSec, usedSec, BATCH_PAGES)
        If waitSec <= 0 Then
            aborted = True
            abortReason = "error"
            Exit For
        End If

        Dim batchLbl As String
        batchLbl = optOcrEta.BatchLabel(b, knownTotal, BATCH_PAGES)
        ' R15-FixA(FA-8): 中断の実況(「中断しています…」)を上書きしない。
        If Not CancelWanted() Then
            modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDDBC) & " PDFを画像に変換しています…"
            modShelfBatch.StageBanner "画像化中… " & batchLbl
        End If

        ' R15-FixB(FB-4): レートの起点は【描画の開始】。従来はOCRループ直前から
        ' 測っており、1頁あたりの実績にGSの描画時間が入らず、次の資料のETAが
        ' 必ず短めに出ていた(残り時間が減らない=最も不安な壊れ方)。
        Dim batchT0 As Double: batchT0 = Timer
        Dim usedThis As Long: usedThis = 0
        Dim killed As Boolean: killed = False
        Dim rb As Long
        rb = RenderBatch(gsExe, pdfPath, folderPath, dpi, firstP, lastP, waitSec, _
                         usedThis, killed, batchLbl)
        usedSec = usedSec + usedThis

        ' R15-FixA(FA-5i): 画像化の待ちから戻った直後に、時間切れ判定より先に
        ' 中断を見る。順序が逆だと、利用者が押して止めた待ちが「1200秒以内に
        ' 終わりませんでした」という無関係な時間切れとして記録・表示される。
        If CancelWanted() Then
            aborted = True
            abortReason = "cancel"
            Exit For
        End If

        If rb = RB_LAUNCH_FAIL Then
            launchFailed = True
            Exit For
        End If

        Dim baseIdx As Long: baseIdx = 1
        Dim gotN As Long
        gotN = CountBatchFiles(folderPath, firstP, lastP, baseIdx)

        If rb = RB_TIMEOUT Then
            aborted = True
            If LenB(abortReason) = 0 Then abortReason = "error"
            ' 止められなかった(PID不明)なら、GSがまだ書いているフォルダを
            ' 消してはならない(R14-F5。optGsTxtのタイムアウトと同じ扱い)。
            If Not killed Then outKeepWork = True
            ' 最後の1枚は書きかけの可能性がある(2枚以上あるときだけ捨てる)。
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

        ' 上限内に入るページだけをOCRへ回す(上限+1枚目は「まだ先がある」
        ' ことの証拠にだけ使い、読まない)。
        Dim ocrN As Long: ocrN = gotN
        If firstP - 1 + ocrN > safeMax Then ocrN = safeMax - (firstP - 1)
        If ocrN < 0 Then ocrN = 0

        Dim shownBatches As Long: shownBatches = 0
        If knownTotal > 0 Then shownBatches = optOcrCore.BatchCountFor(knownTotal, BATCH_PAGES)

        ' R15-7d: このバッチで読めた頁の pages() 上の範囲(控え書込み用)。
        Dim okAtBatch As Long: okAtBatch = okN
        Dim batchMs As Double: batchMs = 0#
        Dim maxMs As Double: maxMs = 0#
        Dim failN As Long: failN = 0
        Dim firstErr As String: firstErr = ""
        Dim doneBefore As Long: doneBefore = doneN

        Dim i As Long
        For i = 0 To ocrN - 1
            ' R15-6b: 頁境界での中断確認(止まるのは今の頁が終わってから)。
            If CancelWanted() Then
                aborted = True
                abortReason = "cancel"
                Exit For
            End If

            Dim pageNo As Long: pageNo = firstP + i

            ' R15-7d: 前回読めている頁は読み直さない(JPEGは描いてあるが捨てる
            ' =描画はバッチ単位のままで CountBatchFiles の想定を崩さない)。
            ' 復元頁は実測レートにも didN にも混ぜない(待っていない頁のぶん
            ' ETAが速くなるのは嘘)。R15-FixB(FB-3): 復元できたかは outHit だけが
            ' 答える。戻り値の長さで見ると「控えてある空の頁」を毎回読み直す。
            Dim cacheHit As Boolean
            Dim cachedText As String: cachedText = optOcrCache.CachedText(pageNo, cacheHit)
            If cacheHit Then
                pages(okN).page = pageNo
                pages(okN).Text = cachedText
                okN = okN + 1
                GoTo NextPage
            End If

            ' R15-5c: 当該資料の実測が2頁貯まるまでは過去実績でETAを出す。
            ' どちらも無ければ0=残り時間そのものを出さない(従来どおり)。
            Dim avgMs As Double: avgMs = rateMs
            If doneN >= 2 Then avgMs = sumMs / CDbl(doneN)
            Dim banner As String
            banner = optOcrEta.OcrPageBanner(pageNo, knownTotal, b, shownBatches, avgMs, Now)
            ' R15-FixA(FA-8): 中断中は頁バナーで上書きしない(押した直後に出た
            ' 「中断しています…」が次の頁の表示で消えると、押した実感が無い)。
            If Not CancelWanted() Then
                modShelfBatch.StageBanner banner
                modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDDBC) & " " & banner
            End If

            ' 読めたページだけ okN が進む(1ページの失敗で資料全体を捨てない)。
            Dim okBefore As Long: okBefore = okN
            Dim wasLimit As Boolean: wasLimit = False
            Dim t0 As Double: t0 = Timer
            OcrOnePage folderPath & "\" & optOcrCore.PageJpgName(baseIdx + i), _
                       pageNo, visionPrompt, pages, okN, failN, firstErr, wasLimit
            Dim ms As Double: ms = modUtilText.ElapsedMsSince(t0)

            doneN = doneN + 1
            sumMs = sumMs + ms
            batchMs = batchMs + ms
            If ms > maxMs Then maxMs = ms

            ' R15-4b: 上限らしき失敗が3連続で打ち切る(連続を切るのは成功だけ)。
            If okN > okBefore Then
                limitStreak = 0
            ElseIf wasLimit Then
                limitStreak = limitStreak + 1
                If limitStreak >= LIMIT_STREAK_MAX Then
                    aborted = True
                    abortReason = "limit"
                    Exit For
                End If
            End If
NextPage:
        Next i

        ' R15-7d: 画像を消す前にまとめて控える(以後ハングしても失われない)。
        optOcrCache.SaveRange pages, okAtBatch, okN - 1

        allFailN = allFailN + failN
        ' 途中まで回ったバッチもあるので、実際に試した頁数(didN)で割る
        ' (ocrN で割ると1頁あたりが不当に短くなる)。
        Dim didN As Long: didN = doneN - doneBefore

        ' 用が済んだJPEGは即座に消す(%TEMP%を溜めない)。
        KillBatchJpgs folderPath, baseIdx, gotN

        ' R15-5c: レートを過去実績へブレンド(次の資料の1頁目からETAが出る)。
        ' R15-FixB(FB-4): 失敗頁を含むバッチは混ぜない。失敗は成功よりずっと速く
        ' 返る(Base64で即エラー等)ので、実績が不当に短くなり、次の資料で
        ' 「残り約2分」と言って40分かかる。復元だけのバッチ(didN=0)も従来どおり。
        If didN > 0 And failN = 0 Then
            rateMs = modUtilText.BlendPerItemMs(rateMs, batchT0, didN)
            SavePageMs rateMs
        End If

        If didN > 0 Then
            On Error Resume Next
            modLog.LogUsage "vision_page_ms", "", "b=" & b & ";pages=" & didN & _
                ";avg=" & CLng(batchMs / CDbl(didN)) & ";max=" & CLng(maxMs)
            On Error GoTo Failed
        End If

        ' R14-F10: 頁ごとに1行書くと err_log が埋まり本当の原因が流れる。
        ' vision_page_ms と同じくバッチ単位で1行に畳む。
        If failN > 0 Then
            On Error Resume Next
            modLog.LogError "E0303", "optOcrPage.OcrPdfByBatch", modUtil.SafeLeft( _
                "b=" & b & " p" & firstP & "-" & (firstP + didN - 1) & " 読取失敗 " & _
                failN & "/" & didN & "頁 first=" & firstErr, 2000)
            On Error GoTo Failed
        End If

        DoEvents
        If aborted Then Exit For
        If isLastBatch Then Exit For                   ' 資料の終端に達した
        If firstP - 1 + gotN >= renderCap Then Exit For
    Next b

    ' R15-6b: 1頁も読めないうちに止めた場合。NoRenderResult に任せると
    ' 「1200秒以内に終わりませんでした」という、押したボタンと無関係な時間切れ
    ' の案内が出る(しかも作業フォルダを残す)。R15-FixA(FA-4): 文言は他の中断と
    ' 同じ純関数から作る(従来ここだけが「最初から」と直書きで、控えが残って
    ' いても最初からだと断言していた)。
    If okN = 0 And abortReason = "cancel" Then
        outAborted = True
        OcrPdfByBatch = "#ERR:E0303:" & _
            optOcrEta.OcrAbortMemoFor(0, "cancel", optOcrCache.HasSaved())
        Exit Function
    End If
    If okN = 0 And abortReason = "limit" Then
        outAborted = True
        OcrPdfByBatch = "#ERR:E0303:AI利用の上限に達した可能性があります。" & _
            "時間をおいて再度取り込んでください。"
        Exit Function
    End If

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

    ' R14-F2: 途中で止めた場合は【必ず】打ち切り扱いにする。ここをFalseで返して
    ' いたため、20頁で中断した100頁のPDFが status=done として「全部入った」顔で
    ' 本棚に並んでいた(欠落が無言=憲章§4-1違反)。
    Dim truncated As Boolean
    truncated = optOcrCore.IsTruncatedCount(foundTotal, maxPages)
    ' R15-FixA(FA-4): 「設定上限で切った」ことは、このあと中断や頁欠けが
    ' 重なっても消えない別の事実。上書きされる前に控えておく。
    Dim capTrunc As Boolean: capTrunc = truncated
    If aborted Or launchFailed Then
        truncated = True
        outAborted = True
        If LenB(abortReason) = 0 Then abortReason = "error"
        ' R15-7d: 「続きから再開します」と書いてよいのは、頁キャッシュが
        ' 【実際に1頁でも残っている】ときだけ。書込みに失敗していれば False の
        ' まま=「最初から再試行します」と言う(嘘をつかない。憲章§4-1)。
        outMemo = optOcrEta.OcrAbortMemoFor(okN, abortReason, optOcrCache.HasSaved())
    ElseIf allFailN > 0 Then
        ' R15-4a(RC6): 全ページ描けて一部の頁だけ読めなかった第3の出口。
        ' 従来はここが status="done" のまま通り、頁が無言で欠けていた。
        truncated = True
        outAborted = True
        outMemo = optOcrEta.OcrPartialMemoFor(okN, allFailN, knownTotal)
    End If

    ' R15-FixA(FA-4): カードのメモはここで【1回だけ】組み立てる。冒頭=前回の
    ' 続きから再開した事実(FinishDoc に空を渡すと usage_log を済ませて冒頭文
    ' だけが返る。控えの削除は R15-FixB FB-1 で modShelf の done 確定後へ移した)
    ' / 本体=中断・上限エラー・頁欠けの理由 / 末尾=設定上限で切った説明。
    ' 従来はこの3つが optOcrPage と optVision.OcrCapMemo に分かれ、あとから
    ' 来たものが前のものを【置換】していた(A-H4)。
    outMemo = optOcrEta.ComposeOcrMemo( _
        optOcrCache.FinishDoc("", doneN), outMemo, _
        optOcrEta.OcrCapMemoFor(capTrunc, okN, safeMax))

    If okN < safeMax Then ReDim Preserve pages(0 To okN - 1)
    OcrPdfByBatch = modUtil.JoinPagedText(pages, truncated)
    Exit Function

Failed:
    ' ハンドラ稼働中は On Error Resume Next が効かない。Resumeで抜けてから記録
    ' する(opt層GS系共通の作法)。outKeepWork は途中で立てた値をそのまま残す
    ' (生きているGSのフォルダを消させないため)。
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

' 1ページ分のOCR。読めたら pages へ積んで okN を1つ進める。TryRibbonRun の
' 前後で必ず DoEvents を回す(§3-1: せめて1ページ境界では画面を返す)。
' R14-F10: 失敗はここでは【数えるだけ】で、err_log は呼び出し元がバッチ単位で
' 1行にまとめる。R15-4b: 「AI利用の上限らしいか」だけは outLimit で返す
' (判定は modGateway.LooksLikeLimitError の1本だけ。§4-5)。
Private Sub OcrOnePage(ByVal jpgPath As String, ByVal pageNo As Long, _
                       ByVal visionPrompt As String, _
                       ByRef pages() As ExtractedPage, ByRef okN As Long, _
                       ByRef failN As Long, ByRef firstErr As String, _
                       ByRef outLimit As Boolean)
    outLimit = False
    DoEvents
    Dim b64 As String
    b64 = optVision.SafeResultToString(modGateway.TryRibbonRun(BASE64_FUNC_NAME, Array(jpgPath)))
    DoEvents

    If LenB(Trim$(b64)) = 0 Or Left$(b64, 5) = "#ERR:" Then
        outLimit = LooksLikeLimit(b64)
        NotePageFail failN, firstErr, "p" & pageNo & " Base64失敗: " & modUtil.SafeLeft(b64, 200)
        Exit Sub
    End If

    Dim s As String
    s = optVision.SafeResultToString(modGateway.TryRibbonRun(RIBBON_FUNC_NAME, _
        Array(visionPrompt, b64, "", VISION_RESOLUTION, VISION_TOOL_NAME)))
    DoEvents

    If optVision.IsVisionError(s) Then
        outLimit = LooksLikeLimit(s)
        NotePageFail failN, firstErr, "p" & pageNo & " 読取失敗: " & modUtil.SafeLeft(s, 200)
        Exit Sub
    End If

    pages(okN).page = pageNo
    pages(okN).Text = s
    okN = okN + 1
End Sub

' 失敗応答が「AI利用の上限」らしく見えるか(R15-4b)。判定そのものは
' modGateway が持つ(基盤層なのでopt層から呼んでよい)。例外が出ても取込は
' 止めない=分からないときは「上限ではない」に倒す(安全側)。
Private Function LooksLikeLimit(ByVal s As String) As Boolean
    On Error Resume Next
    LooksLikeLimit = modGateway.LooksLikeLimitError(s)
    On Error GoTo 0
End Function

' 頁OCRの失敗を数え、最初の1件だけ理由を控える(R14-F10)。
Private Sub NotePageFail(ByRef failN As Long, ByRef firstErr As String, ByVal note As String)
    failN = failN + 1
    If LenB(firstErr) = 0 Then firstErr = note
End Sub

' CancelWanted - 利用者が進捗バナーの中断ボタンを押したか(R15-6b)。
'   ChatGPTVは同期でVBAから止められないので、止まるのは「今の頁が終わって
'   から」(押した直後の案内で先に伝えてある)。印の実体は modShelfBatch にあり
'   (取込の入口でリセットされる)ここは読むだけ。OERNで包む(§4-4)。
Private Function CancelWanted() As Boolean
    On Error Resume Next
    CancelWanted = modShelfBatch.CancelRequested()
    On Error GoTo 0
End Function

' PriorPageMs / SavePageMs - 1頁あたり所要ミリ秒の過去実績(R15-5c)。
'   RC2: ETAは当該資料で2頁読み終えるまで出せず、24頁中20頁が終わるまで何も
'   出なかった。同じ端末・同じリボンならレートはほぼ同じなので、実績を
'   ui_state に残して次の取込の1頁目からETAを出す。読み書きの失敗は黙って
'   諦める(憲章§4-4)。R15-FixB(FB-4): 読み出しは常識の幅(3〜120秒/頁)へ
'   丸める。丸めの表は optOcrEta.ClampPageMs(純関数)が1本だけ持つ。
Private Function PriorPageMs() As Double
    On Error Resume Next
    PriorPageMs = optOcrEta.ClampPageMs(Val(modState.LoadState(RATE_KEY, "0")))
    On Error GoTo 0
End Function

Private Sub SavePageMs(ByVal ms As Double)
    If ms <= 0# Then Exit Sub
    On Error Resume Next
    modState.SaveState RATE_KEY, CStr(CLng(ms))
    On Error GoTo 0
End Sub

' 1バッチ分だけGSを起動して完了を待つ。戻り値は RB_DONE / RB_LAUNCH_FAIL /
' RB_TIMEOUT(R14-F1: 起動できなかったことを時間切れと混ぜない)。
' outUsedSec : この待ちに実際に使った秒数(絶対上限の残り計算用。R14-F6)
' outKilled  : タイムアウト時にGSを止められたか(R14-F5)。止められなかった
'              ときだけ、呼び出し元が作業フォルダを残す判断をする。
' 完了フラグは毎回消してから起動する(前のフラグが残っていると、待ちループが
' 「もう終わっている」と即座に誤判定する)。
' bannerLabel: 待っているあいだ1秒ごとに出す実況のラベル(R15-FixA FA-5i)。
Private Function RenderBatch(ByVal gsExe As String, ByVal pdfPath As String, _
                             ByVal folderPath As String, ByVal dpi As Long, _
                             ByVal firstP As Long, ByVal lastP As Long, _
                             ByVal waitSec As Long, ByRef outUsedSec As Long, _
                             ByRef outKilled As Boolean, _
                             ByVal bannerLabel As String) As Long
    outUsedSec = 0
    outKilled = False

    Dim flagPath As String: flagPath = optOcrCore.DoneFlagFor(folderPath)
    KillIfExists flagPath

    ' R15-4d(RC10): 描き始める前に前のバッチのJPEGを1枚残らず消す。従来は完了
    ' フラグしか消さず、KillBatchJpgs がEDRロック等で失敗して残った1枚を次の
    ' CountBatchFiles が自分の出力として数え、その枚数のズレがそのまま頁番号の
    ' ズレ(出典ページが静かに狂う=§4-1違反)になっていた。
    If Not PurgePageJpgs(folderPath) Then
        On Error Resume Next
        modLog.LogError "E0303", "optOcrPage.OcrPdfByBatch", modUtil.SafeLeft( _
            "JPEG残骸を消せずバッチ中止 p" & firstP & "-" & lastP & " " & _
            folderPath & " " & modUtil.SafeLeft(pdfPath, 200), 2000)
        On Error GoTo 0
        RenderBatch = RB_LAUNCH_FAIL
        Exit Function
    End If

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
    finished = optGsTxt.WaitForDoneFlag(flagPath, waitSec, gsRc, bannerLabel)
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

    RenderBatch = RB_TIMEOUT
    ' R15-FixA(FA-5i): 利用者が止めた待ちは【時間切れではない】。GS_TIMEOUT を
    ' err_logへ残すと、押した本人の操作と無関係な障害として調査を誤らせる。
    If CancelWanted() Then Exit Function

    On Error Resume Next
    modLog.LogError "E0303", "optOcrPage.OcrPdfByBatch", modUtil.SafeLeft( _
        "GS_TIMEOUT p" & firstP & "-" & lastP & " wait=" & waitSec & "s pid=" & gsPid & _
        " killed=" & outKilled & " " & optGsTxt.GsFailureDetail(folderPath) & _
        " " & modUtil.SafeLeft(pdfPath, 200), 2000)
    On Error GoTo 0
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

' ----------------------------------------------------------------------------
' PurgePageJpgs - 作業フォルダの page_*.jpg を全部消す(2026-08-04 R15-4d)。
'   戻り値 True=1枚も残っていない / False=残骸が残った(頁番号がズレるので
'   呼び出し元は当該バッチを起動失敗扱いにする)。
'   Dir$ は非再入(列挙の途中で別のDir$を呼ぶと最初の列挙が壊れる)なので、
'   消したあとの確認は「最初の1件が返るか」の1回だけで列挙を続けない。
' ----------------------------------------------------------------------------
Private Function PurgePageJpgs(ByVal folderPath As String) As Boolean
    On Error Resume Next
    Kill folderPath & "\" & PAGE_JPG_GLOB
    Err.Clear
    Dim leftover As String: leftover = ""
    leftover = Dir$(folderPath & "\" & PAGE_JPG_GLOB)
    Err.Clear
    On Error GoTo 0
    PurgePageJpgs = (LenB(leftover) = 0)
End Function

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
' 3つの事実を最後まで別物として扱う(R14-F1):
'   launchFailed : GSを【起動できなかった】。待ってすらいないので秒数の話を
'                  してはならない(従来は時間切れと同じ=直らない案内だった)。
'                  材料も出ないので作業フォルダは残さない。
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
