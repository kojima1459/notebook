Attribute VB_Name = "modEmbed"
Option Explicit

' ============================================================================
' modEmbed - 埋め込み再開可能バッチ(MASTER_SPEC §7.2)
' ----------------------------------------------------------------------------
' 役割:
'   my_knowledge の embedded=0 の行を順に modGateway.GetEmbedding で埋め込み、
'   my_vectors へ追記して embedded=1 に確定させる。V2 src/chatbot_v2/
'   modVectorize.bas の「進捗+ESC中断+再開可能バッチ」パターンを踏襲する。
'
' 設計判断:
'   ・再開可能性は「embedded列がその場で1になる」ことそのもので実現する。
'     メモリ上でまとめて処理して最後に一括書込みすると、途中でExcelが
'     強制終了した場合に進捗が失われてしまう(MASTER_SPEC §13
'     「embedded=0が残った状態でExcel強制終了→次回同期で再開」)。
'     成功した行は都度シートへ書き込む(V2 modVectorize.bas も同じ設計)。
'   ・スキャン対象の特定(embedded=0の行探し)自体は my_knowledge 全体を
'     1回のRange読み取りでVariant配列に落としてメモリ内で判定する
'     (ループ内でCellsの直接アクセスを繰り返さない。MASTER_SPEC §12)。
'     成功行の書込み(embeddedフラグ・vectors追記)自体は再開可能性のため
'     都度Cellsへ書く必要があるが、これは「毎回シート全体を読み直す」
'     パターンではなく1～2セルへの確定書込みのみなので同じ問題ではない。
'   ・3連続失敗の判定に加え、modGateway.GetEmbedding は失敗時に応答文字列を
'     外へ返さない契約(空配列のみ)のため、LooksLikeLimitError へ渡す
'     材料がこのモジュールには直接無い。そこで直近の modLog.LogError 書込み
'     (err_log の最終行)をこのモジュールが読み返し、その detail 文字列に
'     対して modRibbonFail.LooksLikeLimitError を適用することで、利用上限
'     らしき応答を検知する(modGatewayの契約シグネチャは変更していない。
'     内部実装だけの工夫)。
'   ・進捗実況は modUIMain.SetStage を1行スコープの On Error Resume Next で
'     呼ぶ(modUIMain未実装/実行時エラーでも埋め込み処理自体は止めない。
'     恒久実装としてはWave2-Fが提供するので通常呼びでよい)。
'   ・ESC対応: Application.EnableCancelKey = 2 (xlErrorHandler)。名前付き
'     定数を使わないのは modLog.bas の xlSheetHidden 回避と同じ理由
'     (LibreOffice実行環境での互換性を優先するV2以来の慣習)。ESCはErr 18
'     として捕捉し、ここまでの成功分を保持したまま関数を終える。
' ============================================================================

Private Const COL_ID As Long = 1
Private Const COL_SOURCE As Long = 2
Private Const COL_ORIGIN As Long = 3
Private Const COL_PAGE As Long = 4
Private Const COL_SUMMARY As Long = 5
Private Const COL_KEYWORDS As Long = 6
Private Const COL_FULLTEXT As Long = 7
Private Const COL_ADDED As Long = 8
Private Const COL_EMBEDDED As Long = 9

Private mRunning As Boolean   ' 再入防止(ESC中のDoEventsによる多重起動対策)
' ガードの自己回復(2026-07-16): modShelf.mIngestingと同じ理由のタイムスタンプ。
Private mRunningSince As Date
Private Const GUARD_EXPIRY_MIN As Long = 30

' ----------------------------------------------------------------------------
' EmbedPending - embedded=0 の行を順に埋め込む。戻り値=今回埋め込んだ件数。
' ----------------------------------------------------------------------------
Public Function EmbedPending(Optional ByVal maxCount As Long = -1) As Long
    Dim doneCount As Long: doneCount = 0

    ' R15-1a: 失効は「開始から」ではなく「最後のビートから」で数える
    ' (埋め込みが長引いている最中にガードが解けるのを止める。実機第4報 RC5)。
    If mRunning Then
        If modShelfBatch.GuardExpiredNow(mRunningSince, GUARD_EXPIRY_MIN) Then
            On Error Resume Next
            modLog.LogUsage "guard_recover", "embed", _
                "埋め込みガードが無音のまま" & GUARD_EXPIRY_MIN & "分以上残留していたため自動解除"
            On Error GoTo 0
            mRunning = False
        End If
    End If
    If mRunning Then
        EmbedPending = 0
        Exit Function
    End If
    mRunning = True
    mRunningSince = Now

    Dim wsK As Worksheet: Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    Dim wsV As Worksheet: Set wsV = EnsureVectorSheet()
    If wsK Is Nothing Or wsV Is Nothing Then
        mRunning = False
        EmbedPending = 0
        Exit Function
    End If

    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, COL_ID).End(xlUp).row
    If lastK < 2 Then
        mRunning = False
        EmbedPending = 0
        Exit Function
    End If

    ' 一括読み取り(ループ内Range直接アクセス禁止・MASTER_SPEC §12)
    Dim arr As Variant
    arr = wsK.Range(wsK.Cells(2, 1), wsK.Cells(lastK, COL_EMBEDDED)).Value

    Dim pendingRows() As Long
    ReDim pendingRows(0 To UBound(arr, 1) - 1)
    Dim pendingN As Long: pendingN = 0
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If CStr(arr(i, COL_EMBEDDED)) <> "1" Then
            pendingRows(pendingN) = i    ' arr内の行インデックス(1始まり)
            pendingN = pendingN + 1
        End If
    Next i

    If pendingN = 0 Then
        mRunning = False
        EmbedPending = 0
        Exit Function
    End If

    Dim limit As Long: limit = pendingN
    If maxCount >= 0 And maxCount < limit Then limit = maxCount

    ' R41 B2: 取込/同期のファイルループがバナーを所有しているか(明示の印。
    ' AfterLoopの所有権判定に使う)。「見えているか」で判定しない理由は
    ' modShelfBatch.SetLoopBanner のコメント参照(レビュー1周目 MAJOR-2)。
    Dim hadBanner As Boolean
    On Error Resume Next
    hadBanner = modShelfBatch.LoopBannerOwned()
    On Error GoTo 0

    ' スロットリング待ち(ミリ秒)。2026-07-16: 実機で「取込が遅すぎる」との
    ' 報告。埋め込みは1チャンクごとに実行され、既定150msのスリープが
    ' 数百チャンクの資料では数十秒の純粋な待ち時間になっていた。
    ' ①mock_llm=TRUE(ローカル生成でレート制限が無い)ときはスリープ0、
    ' ②実機(リボン)でも既定を0へ引き下げる(必要ならconfigの
    ' embed_sleep_msで各自調整。レート制限は3連続失敗検知・LimitCheckで別途保護)。
    Dim throttleMs As Long
    If modConfig.GetBool("mock_llm", False) Then
        throttleMs = 0
    Else
        throttleMs = modConfig.GetLong("embed_sleep_ms", 0)
    End If

    Dim consecutiveFail As Long: consecutiveFail = 0
    Dim abortReason As String: abortReason = ""

    Application.EnableCancelKey = 2   ' xlErrorHandler: ESCをErr 18として捕捉する

    ' バッチ収集→GetEmbeddingsBatch(設計書§E-1)。direct時は1リクエストで
    ' まとめて取得し取込を桁違いに高速化。ribbon時は内部で単発ループ=従来同等。
    ' 各成功要素はその場でembedded=1確定(再開可能性は従来どおり)。書込み先は
    ' chunk_id起点で都度再解決(Wave4の再入対策を維持)。例外(ESC=Err18含む)は
    ' バッチ全体を1つのOn Error GoTo EscOrErrで覆う(未捕捉でガードが残る事故防止)。
    Dim batchSize As Long: batchSize = modConfig.GetLong("embed_batch_size", 128)
    If batchSize < 1 Then batchSize = 1

    Dim n As Long
    Dim chunkId As String

    ' 2026-07-31(R7 B-1): 進捗に目安時間を出す。1件あたりの所要ミリ秒は
    ' 直近2バッチの移動平均で追う(先頭バッチは接続の立ち上がりで遅く、
    ' 累積平均だと最後まで悲観的な数字を出し続けるため)。
    ' 最初のバッチが終わるまでは0のまま=件数だけの表示になる。
    Dim msPerItem As Double: msPerItem = 0
    Dim batchT0 As Double

    Dim bStart As Long
    For bStart = 0 To limit - 1 Step batchSize
        ' R15-FixA(FA-5ii・レビューB-H3): 中断の確認。取込の最後はこのベクトル化
        ' で、数百チャンクなら何分もかかる。ここに確認が無かったため、利用者が
        ' ⏹中断を押しても「今のファイルのベクトル化が全部終わるまで」止まらず、
        ' 押しても効かないように見えていた。既存のESC中断とまったく同じ
        ' 後始末(abortReason → AfterLoop)に乗せる: 書けたぶんは embedded=1 で
        ' 確定済み、残りは embedded=0 のまま=次の同期が続きから再開する。
        If CancelWanted() Then
            abortReason = "中断(利用者操作)"
            Exit For
        End If

        Dim bEnd As Long: bEnd = bStart + batchSize - 1
        If bEnd > limit - 1 Then bEnd = limit - 1
        Dim bN As Long: bN = bEnd - bStart + 1

        ' R10-5/R10-5b: ShowProgress化(文言は現行のまま)。関数を抜けるとき
        ' (AfterLoop)に必ずHideProgressするので、ここでは表示だけを都度更新する。
        On Error Resume Next
        modUIMain.ShowProgress "" & ChrW(&HD83D) & ChrW(&HDCE5) & " ベクトル化中 " & _
            modUtil.ProgressText(bEnd + 1, limit, modUtil.EtaText(limit - bStart, msPerItem)) & " …"
        On Error GoTo 0

        batchT0 = Timer
        On Error GoTo EscOrErr
        DoEvents

        Dim texts() As String
        ReDim texts(0 To bN - 1)
        For n = 0 To bN - 1
            texts(n) = CStr(arr(pendingRows(bStart + n), COL_FULLTEXT))
        Next n

        Dim outCsv() As String
        modGateway.GetEmbeddingsBatch texts, outCsv

        For n = 0 To bN - 1
            DoEvents
            chunkId = CStr(arr(pendingRows(bStart + n), COL_ID))

            If LenB(outCsv(n)) > 0 Then
                Dim targetCell As Range
                ' 2026-08-16(R33波3 W3-10): 下の my_vectors 側と同じ理由で明示する。
                ' こちらの方が実害は重い ―― Nothing になると GetEmbeddingsBatch で
                ' 取得済みのベクトルを捨てて【何も書かずに黙って飛ばし】、embedded も
                ' 立たず consecutiveFail も増えないので3連続失敗の安全装置が鳴らない
                ' (=APIを消費するのに1件も進まず、err_log にも痕跡が残らない)。
                Set targetCell = wsK.Columns(COL_ID).Find(What:=chunkId, LookAt:=1, _
                    MatchCase:=True, LookIn:=xlValues, SearchOrder:=xlByRows, _
                    MatchByte:=False)
                If targetCell Is Nothing Then
                    ' 直前に削除されたチャンク(再入DeleteSource等)はスキップ。
                Else
                    ' 2026-07-28(レビュー M-3): 追記ではなく upsert にする。
                    ' 「ベクトル追記 → embedded=1」の間に ESC やエラーで
                    ' 中断すると、次回の再埋め込みで同じ chunk_id の行が
                    ' もう1本できていた。単段検索は my_vectors を素直に
                    ' 走査するので、同じチャンクが topK の枠を2つ占める。
                    ' 既存行があれば上書きする(無ければ末尾へ追記)。
                    ' 2026-08-16(R33波3 W3-5): LookIn/SearchOrder/MatchByte を
                    ' 明示する。Range.Find はこの3つを省略すると【そのExcel
                    ' セッションで最後に使われた値】を引き継ぐ仕様で、利用者が
                    ' 別ブックで Ctrl+F の「検索対象=コメント」を一度使うだけで、
                    ' セル値と一致していても Nothing が返るようになる。ここが
                    ' Nothing になると upsert が常に追記へ倒れ、同じ chunk_id の
                    ' ベクトル行が二重化して topK の枠を2つ占める(2026-07-28 M-3
                    ' で潰したはずの症状)。LibreOffice の Find にはこの持ち越しが
                    ' 無いため、LO実行テストでは再現しない型の事故。
                    Dim vRow As Long
                    Dim vFound As Range
                    Set vFound = wsV.Columns(1).Find(What:=chunkId, LookAt:=1, _
                        MatchCase:=True, LookIn:=xlValues, SearchOrder:=xlByRows, _
                        MatchByte:=False)
                    If vFound Is Nothing Then
                        vRow = wsV.Cells(wsV.Rows.count, 1).End(xlUp).row + 1
                        If vRow < 2 Then vRow = 2
                    Else
                        vRow = vFound.row
                    End If
                    wsV.Cells(vRow, 1).Value = chunkId
                    wsV.Cells(vRow, 2).Value = modUtil.SafeLeft(outCsv(n), 32000)
                    wsK.Cells(targetCell.row, COL_EMBEDDED).Value = 1
                    doneCount = doneCount + 1
                    consecutiveFail = 0
                End If
            Else
                ' R48: 本文が空のかけらは「失敗」ではない(問い合わせる相手が無い)。
                ' modGateway.RibbonEmbedRange に Erase を入れるまでは、空のかけらへ
                ' 直前のかけらのベクトルが流用されて「成功」に化けていた。
                ' Erase を入れた結果ここへ落ちてくるようになったので、3連続失敗の
                ' 安全装置には数えない(空行の続く資料で取込が止まってしまう)。
                ' 印(embedded)だけ立てて先へ進める。ベクトル行は書かないので
                ' 検索には出てこない ―― それが空のかけらの正しい状態。
                If LenB(Trim$(texts(n))) = 0 Then
                    Dim emptyCell As Range
                    Set emptyCell = wsK.Columns(COL_ID).Find(What:=chunkId, LookAt:=1, _
                        MatchCase:=True, LookIn:=xlValues, SearchOrder:=xlByRows, _
                        MatchByte:=False)
                    If Not emptyCell Is Nothing Then
                        wsK.Cells(emptyCell.row, COL_EMBEDDED).Value = 1
                    End If
                    On Error Resume Next
                    modLog.LogUsage "embed_empty", "ingest", "chunk=" & chunkId
                    On Error GoTo EscOrErr
                Else
                    consecutiveFail = consecutiveFail + 1
                    If consecutiveFail >= 3 Then
                        abortReason = "3連続失敗"
                    ElseIf LastFailureLooksLikeLimit() Then
                        abortReason = "利用上限の疑い"
                    End If
                End If
            End If
            If LenB(abortReason) > 0 Then Exit For
        Next n

        If throttleMs > 0 Then SleepMs throttleMs
        msPerItem = modUtilText.BlendPerItemMs(msPerItem, batchT0, bN)
        On Error GoTo 0
        If LenB(abortReason) > 0 Then Exit For
    Next bStart
    GoTo AfterLoop

EscOrErr:
    abortReason = "中断(ESCまたは例外)"
    ' Resume で抜けてハンドラ実行中の状態を解除する(On Error GoTo 0 では
    ' 解除されず、AfterLoop: の後始末で起きたエラーが呼び出し元へ素通りする)。
    Resume AfterLoop

AfterLoop:
    On Error Resume Next
    Application.EnableCancelKey = 1   ' xlInterrupt(既定へ戻す)
    ' R10-5b: ShowProgressを呼んだ者が自分で消す原則に統一する。従来は
    ' 「呼び出し元(AddFilesResult/SyncNow)が必ずHideする」前提だったが、
    ' modShelf.IngestFileを直接呼ぶ経路(modUIShelf.OnAddScreenshot/
    ' modVault.RegisterKnowledgeText)はHideを呼ばないため、進捗バナーが
    ' 消え残っていた(発見事項1の裁定)。ここは正常完了・ESC中断・バッチ内
    ' 例外のいずれもAfterLoopへ合流する唯一の出口なので、ここで1回Hideすれば
    ' 全経路をカバーできる(呼び出し元側の既存Hide呼び出しは冗長になるが
    ' 二重Hideは無害なのでそのまま残す)。
    ' R41 B2: ただし「同期中N/M」のようにフォルダ同期がファイルごとに出す
    ' バナーに、この関数(1ファイルの末尾)が相乗りしているだけの場合は消さない
    ' (消すのはバナーを出したループの最後)。hadBannerで開始時の所有権を控え、
    ' このループ自身が出していないバナーには手を触れない。
    If Not hadBanner Then modUIMain.HideProgress
    On Error GoTo 0

    If InStr(abortReason, "上限") > 0 Or InStr(abortReason, "連続失敗") > 0 Then
        modLog.LogError "E0204", "modEmbed.EmbedPending", abortReason & " done=" & doneCount & "/" & limit
    End If

    ' R12-4: my_vectors に1本でも書いたらセッション内キャッシュは古い。
    ' 途中中断(ESC/上限)でも書いた分はあるので doneCount で判断する。
    If doneCount > 0 Then
        On Error Resume Next
        modVecCache.BumpGeneration
        On Error GoTo 0
    End If

    mRunning = False
    EmbedPending = doneCount
End Function

' ----------------------------------------------------------------------------
' PendingCount - my_knowledge の embedded<>1 件数(重い処理はしない前提の
'   軽量カウント。1回のRange読み取りで済ませる)。
' ----------------------------------------------------------------------------
Public Function PendingCount() As Long
    Dim wsK As Worksheet: Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If wsK Is Nothing Then Exit Function

    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, COL_ID).End(xlUp).row
    If lastK < 2 Then Exit Function

    Dim cnt As Long: cnt = 0
    If lastK = 2 Then
        ' 単一セル範囲は.Valueがスカラーを返すため配列読みできない特殊ケース
        If CStr(wsK.Cells(2, COL_EMBEDDED).Value) <> "1" Then cnt = 1
        PendingCount = cnt
        Exit Function
    End If

    Dim arr As Variant
    arr = wsK.Range(wsK.Cells(2, COL_EMBEDDED), wsK.Cells(lastK, COL_EMBEDDED)).Value
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If CStr(arr(i, 1)) <> "1" Then cnt = cnt + 1
    Next i
    PendingCount = cnt
End Function

' ----------------------------------------------------------------------------
' MarkAllForReembed - 全チャンクを再ベクトル化対象へ戻す(設計書§G-T10)。
'   embed_dim/vector_precision変更後の移行導線: my_knowledgeのembedded列を
'   全行0へ(200行ずつバッチ書込み)、my_vectorsのデータ行を全消去。
'   この後EmbedPending(または🔄同期)で新しい設定のベクトルが再構築される。
' ----------------------------------------------------------------------------
Public Sub MarkAllForReembed()
    Dim wsK As Worksheet: Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If wsK Is Nothing Then Exit Sub

    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, COL_ID).End(xlUp).row
    If lastK >= 2 Then
        Const RESET_BATCH As Long = 200
        Dim r As Long
        For r = 2 To lastK Step RESET_BATCH
            Dim rEnd As Long: rEnd = r + RESET_BATCH - 1
            If rEnd > lastK Then rEnd = lastK
            Dim zeros() As Variant
            ReDim zeros(1 To rEnd - r + 1, 1 To 1)
            Dim z As Long
            For z = 1 To rEnd - r + 1
                zeros(z, 1) = 0
            Next z
            wsK.Range(wsK.Cells(r, COL_EMBEDDED), wsK.Cells(rEnd, COL_EMBEDDED)).Value = zeros
        Next r
    End If

    Dim wsV As Worksheet: Set wsV = GetSheet(modAppDef.SH_VECTORS)
    If Not wsV Is Nothing Then
        Dim lastV As Long: lastV = wsV.Cells(wsV.Rows.count, 1).End(xlUp).row
        If lastV >= 2 Then
            wsV.Range(wsV.Cells(2, 1), wsV.Cells(lastV, 2)).ClearContents
        End If
    End If

    On Error Resume Next
    ' R12-4: 再埋め込みは「件数も先頭/末尾idも同じまま中身だけ変わる」典型例。
    ' 世代を進めないと、セッション内キャッシュも粗選別の量子化コードも
    ' 古いベクトルのまま使われ続ける(RAG監査3の「再埋め込みを検知しない」)。
    modVecCache.ResetVecCache
    modVecCache.BumpGeneration
    modLog.LogUsage "reembed_reset", "embed", "全チャンクを再ベクトル化対象に設定"
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function EnsureVectorSheet() As Worksheet
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_VECTORS)
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = modAppDef.SH_VECTORS
        ws.Cells(1, 1).Value = "chunk_id"
        ws.Cells(1, 2).Value = "vector_csv"
        On Error Resume Next
        ws.Visible = 2   ' xlSheetVeryHidden
        On Error GoTo 0
    End If
    Set EnsureVectorSheet = ws
    Exit Function
Fail:
    Set EnsureVectorSheet = Nothing
End Function

' 利用者が進捗バナーの中断ボタンを押したか(R15-FixA FA-5ii)。印の実体は
' modShelfBatch(取込の入口でリセットされる)にあり、ここは読むだけ。
' 問い合わせが埋め込みを壊してはならないのでOERNで包む(=分からないときは
' 「押されていない」に倒す。安全側)。
Private Function CancelWanted() As Boolean
    On Error Resume Next
    CancelWanted = modShelfBatch.CancelRequested()
    On Error GoTo 0
End Function

' err_log最終行のdetail文字列にmodRibbonFail.LooksLikeLimitErrorを適用する
' (GetEmbedding失敗時は応答文字列そのものが得られないための代替手段。
' モジュール冒頭の設計判断コメント参照)。
Private Function LastFailureLooksLikeLimit() As Boolean
    On Error GoTo NoLog
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_ERRLOG)
    If ws Is Nothing Then GoTo NoLog
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then GoTo NoLog
    Dim detail As String: detail = CStr(ws.Cells(lastR, 4).Value)
    LastFailureLooksLikeLimit = modRibbonFail.LooksLikeLimitError(detail)
    Exit Function
NoLog:
    LastFailureLooksLikeLimit = False
End Function


' Declareを使わないスリープ(§13: 32/64bit互換のためDeclare不使用で回避)。
' DoEventsで応答性を保ちながらTimer基準で待つ(Timerの日跨ぎは軽微な誤差
' として許容: スロットリング用途であり厳密なタイマーではないため)。
Private Sub SleepMs(ByVal ms As Long)
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000# < ms
        DoEvents
        If Timer < t0 Then Exit Do   ' 深夜0時のTimerロールオーバーガード
    Loop
End Sub
