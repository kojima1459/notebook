Attribute VB_Name = "modAskRetrieve"
Option Explicit

' ========================================
' modAskRetrieve - modAsk から切り出した「検索して材料を揃える」層
'
' 多段RAG(拡張→マルチクエリ→再ランク)と、ヒット結果の評価(低関連度の警告・
' 曖昧すぎる質問の判定・資料名の列挙)。いずれも modAsk のモジュール変数
' (mLastHits 等)を一切触らず、引数の hits() だけで完結する。
' そのため「答えが変」の原因が検索側か生成側かを、ここだけ読めば切り分けられる。
'
' 切り出しの理由(2026-07-28): modAsk が契約上限30,000字に対し残り88字で、
' 出典アクセサの添字修正(レビュー C-2)すら入らなかった(レビュー I-2)。
' ========================================

' 多段RAG(§C): 拡張→マルチクエリ→再ランク。失敗時は単段Searchへ退化。
Public Function RunMultiRetrieve(ByVal q As String, ByVal mdMode As String, _
                                  ByVal topK As Long, ByRef hits() As Hit) As Long
    On Error GoTo FallbackSingle

    ' 1) クエリ拡張
    Dim queries() As String
    Dim queryN As Long: queryN = 0
    ReDim queries(0 To 8)

    Dim standalone As String, hyde As String
    Dim subs() As String
    standalone = "": hyde = ""

    ' 各段を通すかはモードで決まる(modMode)。
    Dim useExpand As Boolean
    useExpand = modMode.UseExpand(mdMode, modConfig.GetBool("expand_enabled", False), _
                                  modConfig.GetBool("quick_expand", False))

    If useExpand Then
        modUIMain.SetStage "" & ChrW(&HD83E) & ChrW(&HDDED) & " 質問を分析中…"
        Dim lightMode As Boolean
        lightMode = modMode.UseLightExpand(mdMode, modConfig.GetBool("quick_expand_light", True))

        Dim exPrompt As String
        exPrompt = modPrompts.BuildExpandPrompt(q, modAsk.HistoryBlock(), _
            modMode.SubQueryCount(mdMode, modConfig.GetLong("expand_subqueries", 3), _
                                  modConfig.GetLong("thorough_subqueries", 6)), lightMode)

        Dim exModel As String: exModel = modConfig.GetString("expand_model", "")
        If LenB(exModel) = 0 Then exModel = modConfig.GetString("quick_model", "gpt-5.5")
        Dim exLat As Long
        Dim exResp As String
        exResp = modGateway.CallLLM(exPrompt, "expand", _
            modConfig.GetString("expand_effort", "low"), _
            modConfig.GetString("expand_verbosity", "low"), exModel, exLat)

        If Not modAsk.IsErrorResponse(exResp) Then
            modRagParse.ParseExpand exResp, standalone, subs, hyde
        End If
    End If

    If LenB(Trim$(standalone)) = 0 Then standalone = q
    queries(queryN) = standalone: queryN = queryN + 1

    Dim si As Long
    On Error Resume Next
    For si = LBound(subs) To UBound(subs)
        If queryN <= 7 And LenB(Trim$(subs(si))) > 0 Then
            queries(queryN) = subs(si): queryN = queryN + 1
        End If
    Next si
    On Error GoTo FallbackSingle
    If LenB(Trim$(hyde)) > 0 And queryN <= 8 Then
        queries(queryN) = hyde: queryN = queryN + 1
    End If
    ReDim Preserve queries(0 To queryN - 1)

    ' 2) マルチクエリ検索(候補プール)
    Dim poolK As Long: poolK = modConfig.GetLong("multi_candidates", 40)
    If poolK < topK Then poolK = topK
    Dim poolHits() As Hit
    Dim poolN As Long
    poolN = modRetrieve.SearchExpanded(queries, poolK, poolHits)
    If poolN <= 0 Then GoTo FallbackSingle

    ' 3) 再ランク(候補がtopKより多いときだけ意味がある)
    Dim orderN As Long: orderN = 0
    Dim rankOrder() As Long
    Dim useRerank As Boolean
    useRerank = modMode.UseRerank(mdMode, modConfig.GetBool("rerank_enabled", False), _
                                  modConfig.GetBool("quick_rerank", False))
    If useRerank And poolN > topK Then
        modUIMain.SetStage "" & ChrW(&HD83E) & ChrW(&HDDEE) & " 関連度を精査中…"
        Dim rkPrompt As String
        rkPrompt = modPrompts.BuildRerankPrompt(q, poolHits, poolN, _
            modConfig.GetLong("max_context_chars", 40000))
        Dim rkModel As String: rkModel = modConfig.GetString("rerank_model", "")
        If LenB(rkModel) = 0 Then rkModel = modConfig.GetString("quick_model", "gpt-5.5")
        Dim rkLat As Long
        Dim rkResp As String
        rkResp = modGateway.CallLLM(rkPrompt, "rerank", _
            modConfig.GetString("rerank_effort", "low"), _
            modConfig.GetString("rerank_verbosity", "low"), rkModel, rkLat)
        If Not modAsk.IsErrorResponse(rkResp) Then
            orderN = modRagParse.ParseRankOrder(rkResp, poolN, rankOrder)
        End If
    End If

    ' 4) 最終topKへ絞り込み(再ランク順 or スコア順)
    Dim outN As Long
    If orderN > 0 Then
        outN = orderN
    Else
        outN = poolN
    End If
    If outN > topK Then outN = topK

    ReDim hits(1 To outN)
    Dim oi As Long
    For oi = 1 To outN
        If orderN > 0 Then
            hits(oi) = poolHits(rankOrder(oi - 1))
        Else
            hits(oi) = poolHits(oi)
        End If
    Next oi
    RunMultiRetrieve = outN
    Exit Function

FallbackSingle:
    Err.Clear
    On Error GoTo 0
    RunMultiRetrieve = modRetrieve.Search(q, topK, hits)
End Function

' 低関連度警告(表示専用): 最高スコアが config low_hit_warn_score(既定0.3)
' 未満なら注意書きを先頭に付ける。0以下は無効化扱い。
Public Function ApplyLowHitWarning(ByVal result As String, hits() As Hit, ByVal nHits As Long) As String
    ApplyLowHitWarning = result
    Dim threshold As Double
    threshold = modConfig.GetDouble("low_hit_warn_score", 0.3)
    If threshold <= 0 Then Exit Function

    Dim maxScore As Double
    maxScore = -1E+30
    Dim i As Long
    For i = 1 To nHits
        If hits(i).score > maxScore Then maxScore = hits(i).score
    Next i
    If maxScore >= threshold Then Exit Function

    ApplyLowHitWarning = ChrW(&H26A0) & ChrW(&HFE0F) & " 手元の資料との関連が薄い可能性があります。回答は参考程度にご覧ください。" & _
        vbLf & vbLf & result
End Function

' IsTooVague - 短すぎ かつ どの資料とも関連が薄い質問だけTrue。LLMを呼ばず
'   聞き方の例を返す。判定は全ヒットの最高スコア(1位だけ見ると誤発動する)。
Public Function IsTooVague(ByVal q As String, hits() As Hit, ByVal nHits As Long) As Boolean
    If nHits < 1 Then Exit Function

    Dim maxChars As Long, thr As Double
    On Error Resume Next
    maxChars = modConfig.GetLong("ambiguous_max_chars", 10)
    thr = CDbl(modConfig.GetLong("ambiguous_score_x100", 60)) / 100#
    On Error GoTo 0
    If thr <= 0# Then Exit Function              ' 0=無効化
    If maxChars < 1 Then maxChars = 10
    If Len(q) > maxChars Then Exit Function

    Dim best As Double, i As Long
    On Error Resume Next
    For i = LBound(hits) To UBound(hits)
        If i > nHits Then Exit For
        If hits(i).score > best Then best = hits(i).score
    Next i
    On Error GoTo 0

    IsTooVague = (best < thr)
End Function

' 逆質問の材料: ヒットした資料名を重複除去して最大4件、| 区切りで返す。
Public Function HitSourceList(hits() As Hit, ByVal nHits As Long) As String
    Dim sb As String, cnt As Long
    Dim i As Long
    On Error Resume Next
    For i = 1 To nHits
        Dim nm As String: nm = Trim$(hits(i).source)
        If LenB(nm) > 0 Then
            If InStr(1, "|" & sb & "|", "|" & nm & "|", vbTextCompare) = 0 Then
                If LenB(sb) > 0 Then sb = sb & "|"
                sb = sb & nm
                cnt = cnt + 1
                If cnt >= 4 Then Exit For
            End If
        End If
    Next i
    On Error GoTo 0
    HitSourceList = sb
End Function
