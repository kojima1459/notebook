Attribute VB_Name = "modEnrich"
Option Explicit

' ============================================================================
' modEnrich - バッチ富化(summary/keywords付与。MASTER_SPEC §7.2)
' ----------------------------------------------------------------------------
' 役割:
'   my_knowledge の summary が空のチャンクを10件ずつまとめ、modGateway.CallLLM
'   (step_name="enrich")へJSON配列を期待するプロンプトを投げて、summary/
'   keywordsを埋める。富化は検索・回答の必須要件ではない(非致命)ため、
'   失敗しても本棚の他の機能を止めない。
'
' 設計判断:
'   ・enrich_mode=off(config既定値)なら即0を返して何もしない。
'     light/fullの違いはMASTER_SPECが具体的な差分定義を与えていないため、
'     現段階ではどちらもEnrichPendingの通常実行として扱う(offだけを
'     明確に区別する)。将来light/fullで対象範囲やプロンプト濃度を分ける
'     場合はここに条件分岐を追加すればよい。
'   ・JSONパースは軽量自前(JsonConverter等の外部依存禁止・§7.2本文)。
'     期待形 [{"i":1,"summary":"…","keywords":"a,b"},…] に対して、
'     "i"キーの出現位置を走査の区切りとして使い、各オブジェクトの中から
'     "summary"/"keywords"の文字列値をInStr+手動エスケープ解除で抜き出す
'     (フル文法パーサーではなく、期待される形からの逸脱には寛容な
'     ベストエフォート抽出。厳密な文法チェックはしない)。
'   ・パース失敗(iが1件も見つからない)バッチはそのままスキップして
'     続行する(§7.2「パース失敗はそのバッチをスキップして続行(非致命)」)。
'     一部の"i"だけ見つかった場合は見つかった分だけ反映し、対応する
'     summary/keywordsが無いキーはそのチャンクをスキップする。
'   ・再開可能性: summaryが空かどうかで対象を判定するため、成功した行は
'     summary列が埋まり、次回EnrichPending呼び出し時には自動的に対象から
'     外れる(modEmbedのembedded列と同じ考え方)。
'   ・ESC対応・3失敗ではなく「利用上限の疑い」検知のみで中断する
'     (CallLLMは失敗時に応答文字列そのものを返す契約のため、
'     modGateway.LooksLikeLimitErrorをそのまま適用できる。modEmbedの
'     GetEmbeddingのように応答文字列が取れない制約が無いため、err_log
'     を読み返す回り道は不要)。
'   ・effort/verbosityはMASTER_SPEC §5のconfigキー台帳にenrich専用の
'     項目が無いため、バッチ処理という性質が近い quick_effort/
'     quick_verbosity を流用する(設計判断。専用キーが必要になれば
'     configキー追加で対応可能)。
'   ・進捗実況は modUIMain.SetStage を1行スコープの On Error Resume Next
'     で呼ぶ(modEmbedと同じ作法。Wave2-F完了後は通常呼びとして機能する)。
' ============================================================================

' 2026-07-28: chunk_id での書き戻し先の引き直し(レビュー M-15)を足したときに
' COL_ID の宣言を入れ忘れ、コンパイルエラーになっていた。
Private Const COL_ID As Long = 1
Private Const COL_SUMMARY As Long = 5
Private Const COL_KEYWORDS As Long = 6
Private Const COL_FULLTEXT As Long = 7

Private Const BATCH_SIZE As Long = 10
Private Const SUMMARY_MAX_CHARS As Long = 500
Private Const KEYWORDS_MAX_CHARS As Long = 300
Private Const SOURCE_TEXT_MAX_CHARS As Long = 1500

Private mRunning As Boolean   ' 再入防止

' ----------------------------------------------------------------------------
' EnrichPending - summary空のチャンクを10件/1回のCallLLMで富化する。
'   戻り値=今回富化できた件数。enrich_mode=offなら即0。
' ----------------------------------------------------------------------------
Public Function EnrichPending(Optional ByVal maxCount As Long = -1) As Long
    Dim mode As String: mode = LCase$(Trim$(modConfig.GetString("enrich_mode", "off")))
    If mode = "off" Or LenB(mode) = 0 Then
        EnrichPending = 0
        Exit Function
    End If

    If mRunning Then
        EnrichPending = 0
        Exit Function
    End If
    mRunning = True

    Dim wsK As Worksheet: Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If wsK Is Nothing Then
        mRunning = False
        EnrichPending = 0
        Exit Function
    End If

    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then
        mRunning = False
        EnrichPending = 0
        Exit Function
    End If

    ' 一括読み取り(ループ内Range直接アクセス禁止・§12)
    Dim arr As Variant
    arr = wsK.Range(wsK.Cells(2, 1), wsK.Cells(lastK, 9)).Value

    Dim pendingRows() As Long: ReDim pendingRows(0 To UBound(arr, 1) - 1)
    Dim pendingCount As Long: pendingCount = 0
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If LenB(Trim$(CStr(arr(i, COL_SUMMARY)))) = 0 Then
            pendingRows(pendingCount) = i   ' arr内の行インデックス(1始まり)
            pendingCount = pendingCount + 1
        End If
    Next i

    If pendingCount = 0 Then
        mRunning = False
        EnrichPending = 0
        Exit Function
    End If

    Dim limit As Long: limit = pendingCount
    If maxCount >= 0 And maxCount < limit Then limit = maxCount

    Dim eff As String: eff = modConfig.GetString("quick_effort", "low")
    Dim vrb As String: vrb = modConfig.GetString("quick_verbosity", "low")

    Application.EnableCancelKey = 2   ' xlErrorHandler: ESCをErr 18として捕捉する

    Dim doneCount As Long: doneCount = 0
    Dim batchStart As Long: batchStart = 0
    Dim abortReason As String: abortReason = ""

    ' 2026-07-31(R7 B-1): 進捗に目安時間を出す(modEmbedと同じ作法)。
    ' 1件あたり所要ミリ秒は直近2バッチの移動平均。最初のバッチが終わるまでは
    ' 0のまま=件数だけの表示になる。
    Dim msPerItem As Double: msPerItem = 0
    Dim batchT0 As Double

    Do While batchStart < limit
        Dim batchSize As Long: batchSize = BATCH_SIZE
        If batchStart + batchSize > limit Then batchSize = limit - batchStart

        On Error Resume Next
        modUIMain.SetStage ChrW(&HD83C) & ChrW(&HDFF7) & ChrW(&HFE0F) & " 富化中 " & _
            modUtil.ProgressText(batchStart + batchSize, limit, _
                                 modUtil.EtaText(limit - batchStart, msPerItem)) & " …"
        On Error GoTo 0

        batchT0 = Timer
        On Error GoTo EscOrErr
        DoEvents   ' ESC割込みとUI応答性の確保
        On Error GoTo 0

        ' 2026-07-28(レビュー M-15): 書き戻し先を行番号ではなく chunk_id で
        ' 解決する。LLM呼び出しには数秒かかり、その間の DoEvents で利用者が
        ' 資料を削除すると行が詰まって【別のチャンクの要約を上書きする】。
        ' modEmbed は既に chunk_id を Find で引き直しており、こちらだけ
        ' 素の行番号のままだった。
        Dim idxs() As Long: ReDim idxs(0 To batchSize - 1)
        Dim ids() As String: ReDim ids(0 To batchSize - 1)
        Dim parts() As String: ReDim parts(0 To batchSize - 1)
        Dim j As Long
        For j = 0 To batchSize - 1
            Dim ai As Long: ai = pendingRows(batchStart + j)
            idxs(j) = ai
            ids(j) = CStr(arr(ai, COL_ID))
            parts(j) = "[" & (j + 1) & "] " & modUtil.SafeLeft(CStr(arr(ai, COL_FULLTEXT)), SOURCE_TEXT_MAX_CHARS)
        Next j
        Dim batchText As String: batchText = Join(parts, vbLf & vbLf)

        Dim prompt As String: prompt = modPrompts.BuildEnrichPrompt(batchText)

        Dim latency As Long
        Dim resp As String
        On Error GoTo EscOrErr
        resp = modGateway.CallLLM(prompt, "enrich", eff, vrb, "", latency)
        On Error GoTo 0

        If Left$(resp, 5) <> "#ERR:" Then
            Dim pIdx() As Long, pSummary() As String, pKeywords() As String, pCount As Long
            If ParseEnrichJson(resp, pIdx, pSummary, pKeywords, pCount) Then
                Dim p As Long
                For p = 0 To pCount - 1
                    Dim localI As Long: localI = pIdx(p) - 1   ' "i"は1始まり
                    If localI >= 0 And localI < batchSize Then
                        Dim sheetRow As Long
                        sheetRow = ResolveRowByChunkId(wsK, ids(localI), idxs(localI) + 1)
                        If sheetRow > 0 Then
                            wsK.Cells(sheetRow, COL_SUMMARY).Value = modUtil.SafeLeft(pSummary(p), SUMMARY_MAX_CHARS)
                            wsK.Cells(sheetRow, COL_KEYWORDS).Value = modUtil.SafeLeft(pKeywords(p), KEYWORDS_MAX_CHARS)
                            ' R12-4: 照合用テキスト(norm_text)は要約・キーワードを含む。
                            ' 富化で中身が変わったのに古い値が残ると、その行だけ
                            ' 【富化前の本文で採点され続ける】。空に戻して検索側の
                            ' 遅延バックフィルに作り直させる。
                            wsK.Cells(sheetRow, modShelfStore.COL_K_NORM).Value = ""
                            doneCount = doneCount + 1
                        End If
                    End If
                Next p
            End If
            ' パース失敗(pCount=0扱い)はこのバッチをスキップして続行(非致命・§7.2)
        Else
            If modGateway.LooksLikeLimitError(resp) Then
                abortReason = "利用上限の疑い"
            End If
        End If

        msPerItem = modUtilText.BlendPerItemMs(msPerItem, batchT0, batchSize)
        batchStart = batchStart + batchSize
        If LenB(abortReason) > 0 Then Exit Do
    Loop
    GoTo AfterLoop

EscOrErr:
    abortReason = "中断(ESCまたは例外)"
    ' Resume で抜けてハンドラ実行中の状態を解除する(On Error GoTo 0 では
    ' 解除されず、AfterLoop: の後始末で起きたエラーが呼び出し元へ素通りする)。
    Resume AfterLoop

AfterLoop:
    On Error Resume Next
    Application.EnableCancelKey = 1   ' xlInterrupt(既定へ戻す)
    On Error GoTo 0

    On Error Resume Next
    modUIMain.SetStage ""
    On Error GoTo 0

    If InStr(abortReason, "上限") > 0 Then
        modLog.LogError "E0204", "modEnrich.EnrichPending", abortReason & " done=" & doneCount & "/" & limit
    End If

    mRunning = False
    EnrichPending = doneCount
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function


' 期待形 [{"i":1,"summary":"…","keywords":"a,b"},…] を軽量自前パースする。
' "i"キーの出現を各オブジェクトの区切りとみなし、その範囲内から"summary"/
' "keywords"の文字列値を抜き出す。1件も"i"が見つからなければFalse。
Private Function ParseEnrichJson(ByVal json As String, ByRef outIdx() As Long, _
        ByRef outSummary() As String, ByRef outKeywords() As String, ByRef outCount As Long) As Boolean
    outCount = 0
    Dim capacity As Long: capacity = 16
    ReDim outIdx(0 To capacity - 1)
    ReDim outSummary(0 To capacity - 1)
    ReDim outKeywords(0 To capacity - 1)

    Dim n As Long: n = Len(json)
    Dim searchFrom As Long: searchFrom = 1
    Dim foundAny As Boolean: foundAny = False

    Do While searchFrom <= n
        Dim iPos As Long: iPos = InStr(searchFrom, json, """i""", vbBinaryCompare)
        If iPos = 0 Then Exit Do

        Dim colonPos As Long: colonPos = InStr(iPos, json, ":")
        If colonPos = 0 Then Exit Do

        Dim numStr As String: numStr = ExtractLeadingNumber(json, colonPos + 1)
        If LenB(numStr) = 0 Then
            searchFrom = iPos + 3
        Else
            Dim iVal As Long: iVal = CLng(Val(numStr))

            Dim nextIPos As Long: nextIPos = InStr(iPos + 3, json, """i""", vbBinaryCompare)
            Dim boundary As Long: boundary = n
            If nextIPos > 0 Then boundary = nextIPos - 1

            Dim summaryPos As Long: summaryPos = InStr(colonPos, json, """summary""", vbTextCompare)
            Dim keywordsPos As Long: keywordsPos = InStr(colonPos, json, """keywords""", vbTextCompare)

            Dim summaryVal As String: summaryVal = ""
            Dim keywordsVal As String: keywordsVal = ""
            If summaryPos > 0 And summaryPos <= boundary Then
                summaryVal = ExtractJsonStringValue(json, summaryPos + Len("""summary"""))
            End If
            If keywordsPos > 0 And keywordsPos <= boundary Then
                keywordsVal = ExtractJsonStringValue(json, keywordsPos + Len("""keywords"""))
            End If

            If outCount > UBound(outIdx) Then
                Dim newCap As Long: newCap = (UBound(outIdx) + 1) * 2
                ReDim Preserve outIdx(0 To newCap - 1)
                ReDim Preserve outSummary(0 To newCap - 1)
                ReDim Preserve outKeywords(0 To newCap - 1)
            End If
            outIdx(outCount) = iVal
            outSummary(outCount) = summaryVal
            outKeywords(outCount) = keywordsVal
            outCount = outCount + 1
            foundAny = True

            searchFrom = boundary + 1
        End If
    Loop

    If outCount > 0 Then
        ReDim Preserve outIdx(0 To outCount - 1)
        ReDim Preserve outSummary(0 To outCount - 1)
        ReDim Preserve outKeywords(0 To outCount - 1)
    Else
        ReDim outIdx(0 To 0)
        ReDim outSummary(0 To 0)
        ReDim outKeywords(0 To 0)
    End If

    ParseEnrichJson = foundAny
End Function

' startPos以降の空白をスキップし、後続の数字列(0-9)を返す(無ければ空文字列)。
Private Function ExtractLeadingNumber(ByVal s As String, ByVal startPos As Long) As String
    Dim n As Long: n = Len(s)
    Dim p As Long: p = startPos
    Do While p <= n And Mid$(s, p, 1) = " "
        p = p + 1
    Loop
    Dim st As Long: st = p
    Do While p <= n
        Dim ch As String: ch = Mid$(s, p, 1)
        If ch < "0" Or ch > "9" Then Exit Do
        p = p + 1
    Loop
    If p > st Then ExtractLeadingNumber = Mid$(s, st, p - st)
End Function

' fromPos(キー文字列の直後)から ":" → 開始の """ → 終了の """ まで走査し、
' \" \\ \n \t の簡易エスケープ解除をしながら文字列値を返す。
Private Function ExtractJsonStringValue(ByVal s As String, ByVal fromPos As Long) As String
    Dim n As Long: n = Len(s)
    Dim p As Long: p = fromPos

    Do While p <= n And Mid$(s, p, 1) <> ":"
        p = p + 1
    Loop
    p = p + 1

    Do While p <= n And Mid$(s, p, 1) <> """"
        p = p + 1
    Loop
    p = p + 1

    Dim buf() As String: ReDim buf(1 To n - p + 2)
    Dim bufLen As Long: bufLen = 0

    Do While p <= n
        Dim c As String: c = Mid$(s, p, 1)
        If c = "\" And p < n Then
            Dim nc As String: nc = Mid$(s, p + 1, 1)
            Select Case nc
                Case """": bufLen = bufLen + 1: buf(bufLen) = """": p = p + 2
                Case "\": bufLen = bufLen + 1: buf(bufLen) = "\": p = p + 2
                Case "n": bufLen = bufLen + 1: buf(bufLen) = vbLf: p = p + 2
                Case "t": bufLen = bufLen + 1: buf(bufLen) = vbTab: p = p + 2
                Case Else: bufLen = bufLen + 1: buf(bufLen) = nc: p = p + 2
            End Select
        ElseIf c = """" Then
            Exit Do
        Else
            bufLen = bufLen + 1
            buf(bufLen) = c
            p = p + 1
        End If
    Loop

    If bufLen = 0 Then
        ExtractJsonStringValue = ""
    Else
        ReDim Preserve buf(1 To bufLen)
        ExtractJsonStringValue = Join(buf, "")
    End If
End Function

' chunk_id から現在の行番号を引き直す。見つからなければ0(=そのチャンクは
' LLM 呼び出しの最中に削除された。書かずに捨てるのが正しい)。
' hintRow は読み取り時の行番号で、まだ同じ chunk_id ならそのまま使う
' (毎回 Find すると数千行で遅いため、当たりを先に確かめる)。
Private Function ResolveRowByChunkId(ByVal wsK As Worksheet, ByVal chunkId As String, _
                                     ByVal hintRow As Long) As Long
    On Error Resume Next
    If LenB(chunkId) = 0 Then Exit Function
    If hintRow >= 2 Then
        If StrComp(CStr(wsK.Cells(hintRow, COL_ID).Value), chunkId, vbBinaryCompare) = 0 Then
            ResolveRowByChunkId = hintRow
            Exit Function
        End If
    End If
    Dim found As Range
    Set found = wsK.Columns(COL_ID).Find(What:=chunkId, LookAt:=1, MatchCase:=True)
    If Not found Is Nothing Then ResolveRowByChunkId = found.row
    On Error GoTo 0
End Function
