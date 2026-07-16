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
'     パターンではなく1〜2セルへの確定書込みのみなので同じ問題ではない。
'   ・3連続失敗の判定に加え、modGateway.GetEmbedding は失敗時に応答文字列を
'     外へ返さない契約(空配列のみ)のため、LooksLikeLimitError へ渡す
'     材料がこのモジュールには直接無い。そこで直近の modLog.LogError 書込み
'     (err_log の最終行)をこのモジュールが読み返し、その detail 文字列に
'     対して modGateway.LooksLikeLimitError を適用することで、利用上限
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

    If mRunning Then
        If DateDiff("n", mRunningSince, Now) >= GUARD_EXPIRY_MIN Then
            On Error Resume Next
            modLog.LogUsage "guard_recover", "embed", _
                "前回の埋め込みガードが" & GUARD_EXPIRY_MIN & "分以上残留していたため自動解除"
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

    ' スロットリング待ち(ミリ秒)。2026-07-16: 実機で「取込が遅すぎる」との
    ' 報告。埋め込みは1チャンクごとに実行され、既定150msのスリープが
    ' 数百チャンクの資料では数十秒の純粋な待ち時間になっていた。
    ' ①mock_llm=TRUE(ローカル生成でレート制限が無い)ときはスリープ0、
    ' ②実機(リボン)でも既定を0へ引き下げる(必要ならconfigの
    ' embed_sleep_msで各自調整。レート制限は3連続失敗検知・LimitCheckで別途保護)。
    Dim throttleMs As Long
    If modConfig.GetBool("mock_llm", True) Then
        throttleMs = 0
    Else
        throttleMs = modConfig.GetLong("embed_sleep_ms", 0)
    End If

    Dim consecutiveFail As Long: consecutiveFail = 0
    Dim abortReason As String: abortReason = ""

    Application.EnableCancelKey = 2   ' xlErrorHandler: ESCをErr 18として捕捉する

    Dim n As Long
    Dim ai As Long
    Dim chunkId As String, fullText As String
    Dim latency As Long
    Dim vec() As Double

    For n = 0 To limit - 1
        ai = pendingRows(n)

        On Error Resume Next
        modUIMain.SetStage "📥 ベクトル化中 " & (n + 1) & "/" & limit & " …"
        On Error GoTo 0

        ' Wave4修正(実コード検証で判明): 以前は書込み区間(Find〜2セル書込み)の
        ' 直前で On Error GoTo 0 していたため、この極短区間でESC(Err18)や
        ' 書込み例外が起きると EscOrErr に飛ばず未捕捉のまま EmbedPending から
        ' 抜けてしまい、AfterLoopのmRunning=False・呼び出し元IngestFileの
        ' mIngesting=False(Finishラベル)のどちらにも到達しない事故があった
        ' (以後IngestFileが常にE0503を返す=Excel再起動まで取込不能)。
        ' 書込み区間・SleepMsまで含めて1回のOn Error GoTo EscOrErrで覆う。
        On Error GoTo EscOrErr
        DoEvents   ' ESC割込みとUI応答性の確保
        chunkId = CStr(arr(ai, COL_ID))
        fullText = CStr(arr(ai, COL_FULLTEXT))
        vec = modGateway.GetEmbedding(fullText, latency)

        If modUtil.HasVector(vec) Then
            ' Wave4修正(実コード検証で判明): モジュール冒頭コメント・
            ' modShelfSync.bas冒頭の設計判断コメントは「chunk_id起点に書込み先を
            ' 都度再解決する自己防衛を入れている」と述べていたが、実装が伴って
            ' いなかった(sheetRow/writeRowをDoEvents前の位置のままキャッシュして
            ' 書き込んでいた)。DoEvents/GetEmbedding呼び出し中に別のマクロ
            ' (SyncNowの削除ループ→DeleteSource等)が再入してmy_knowledge/
            ' my_vectorsの行を削除・圧縮すると、キャッシュした行番号がずれて
            ' 無関係な行にembedded=1を立てたりベクトルを誤った行へ書いたりする
            ' 事故になる。書込み直前にchunk_id起点で現在の行を都度再解決する
            ' ことで、この危険な窓を閉じる。
            Dim targetCell As Range
            Set targetCell = wsK.Columns(COL_ID).Find(What:=chunkId, LookAt:=1, MatchCase:=True)
            If targetCell Is Nothing Then
                ' 書込み直前にこのチャンクが削除された(再入によるDeleteSource等)。
                ' 孤児ベクトルを残さないよう、このチャンク分の書込みはスキップする。
            Else
                Dim vRow As Long
                vRow = wsV.Cells(wsV.Rows.count, 1).End(xlUp).row + 1
                If vRow < 2 Then vRow = 2
                wsV.Cells(vRow, 1).Value = chunkId
                wsV.Cells(vRow, 2).Value = modUtil.SafeLeft(modUtil.VectorToCsv(vec), 32000)
                wsK.Cells(targetCell.row, COL_EMBEDDED).Value = 1
                doneCount = doneCount + 1
                consecutiveFail = 0
            End If
        Else
            consecutiveFail = consecutiveFail + 1
            If consecutiveFail >= 3 Then
                abortReason = "3連続失敗"
            ElseIf LastFailureLooksLikeLimit() Then
                abortReason = "利用上限の疑い"
            End If
        End If

        If throttleMs > 0 Then SleepMs throttleMs
        On Error GoTo 0
        If LenB(abortReason) > 0 Then Exit For
    Next n
    GoTo AfterLoop

EscOrErr:
    Err.Clear
    On Error GoTo 0
    abortReason = "中断(ESCまたは例外)"

AfterLoop:
    On Error Resume Next
    Application.EnableCancelKey = 1   ' xlInterrupt(既定へ戻す)
    On Error GoTo 0

    If InStr(abortReason, "上限") > 0 Or InStr(abortReason, "連続失敗") > 0 Then
        modLog.LogError "E0204", "modEmbed.EmbedPending", abortReason & " done=" & doneCount & "/" & limit
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

' err_log最終行のdetail文字列にmodGateway.LooksLikeLimitErrorを適用する
' (GetEmbedding失敗時は応答文字列そのものが得られないための代替手段。
' モジュール冒頭の設計判断コメント参照)。
Private Function LastFailureLooksLikeLimit() As Boolean
    On Error GoTo NoLog
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_ERRLOG)
    If ws Is Nothing Then GoTo NoLog
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then GoTo NoLog
    Dim detail As String: detail = CStr(ws.Cells(lastR, 4).Value)
    LastFailureLooksLikeLimit = modGateway.LooksLikeLimitError(detail)
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
    Loop
End Sub
