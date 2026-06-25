Attribute VB_Name = "modPipeline"
Option Explicit

' ============================================================================
' modPipeline - 4-step query pipeline
' ----------------------------------------------------------------------------
'   1. Router       : LLM picks top-N relevant chunk IDs from summaries
'   2. FeedbackLookup: find similar approved/own pending Q&A
'   3. Drafter      : LLM writes structured answer with citations
'   4. Verifier     : LLM checks each claim against source, strips fabrications
'
' Returns a PipelineResult with .answer (final), .citations (display), and
' .timings (per-step latency).
'
' All LLM calls go through modRibbonGateway.CallLLM().
' ============================================================================

Public Type PipelineResult
    OK As Boolean
    Question As String
    SelectedIds As String       ' comma sep
    Draft As String
    Answer As String            ' final answer (after verifier if enabled)
    Citations As String
    RouterMs As Long
    DraftMs As Long
    VerifyMs As Long
    TotalMs As Long
    ErrorMsg As String
End Type

Public Function RunQuery(ByVal question As String) As PipelineResult
    Dim res As PipelineResult
    res.Question = question
    Dim tAll As Double: tAll = Timer
    On Error GoTo Trap

    Dim deptId As String: deptId = modUserProfile.CurrentDept()
    Dim user As String: user = modUserProfile.CurrentUser()

    ' ---- Step 1: Router ---------------------------------------------------
    Dim routerTable As String
    routerTable = modKnowledgeBase.BuildRouterTable(deptId)
    If LenB(routerTable) = 0 Then
        res.OK = False
        res.ErrorMsg = "[Step1 ルーター] あなたの部署(" & deptId & ")で参照できるナレッジが0件です。" & vbLf & _
                       "knowledge_base の dept_scope と、選択中の部署IDを確認してください。"
        GoTo Finish
    End If

    Dim routerN As Long: routerN = modConfig.GetLong("router_max_chunks", 8)
    Dim routerPrompt As String
    routerPrompt = modPrompts.GetRouter()
    ' For follow-up questions, prepend the conversation history to the router
    ' question so elliptical phrases ("その手数料は？") still retrieve the
    ' chunks that match the implicit topic. The drafter sees the history too,
    ' but if the router only sees the short follow-up text it can pick wrong
    ' or empty chunks and starve the drafter.
    Dim routerQuestion As String: routerQuestion = question
    Dim rHist As String: rHist = modBoot.HistoryBlock()
    If modBoot.gFollowupMode And LenB(rHist) > 0 Then
        routerQuestion = "【これまでの会話（参考）】" & vbLf & rHist & vbLf & _
                         "【続きの質問】" & vbLf & question
    End If
    routerPrompt = Replace(routerPrompt, "{question}", routerQuestion)
    routerPrompt = Replace(routerPrompt, "{max_n}", CStr(routerN))
    routerPrompt = Replace(routerPrompt, "{knowledge_table}", routerTable)

    Dim routerOut As String
    routerOut = modRibbonGateway.CallLLM(routerPrompt, "router", res.RouterMs)
    modRibbonGateway.LogDebugCall "router", routerPrompt, routerOut, res.RouterMs

    If Left$(routerOut, 12) = "#LLM_ERROR: " Then
        res.OK = False
        res.ErrorMsg = "[Step1 ルーター] 社内AIリボン呼び出しに失敗。" & vbLf & routerOut
        GoTo Finish
    End If

    res.SelectedIds = ParseSelectedIds(routerOut)
    If LenB(res.SelectedIds) = 0 Then
        res.OK = False
        res.ErrorMsg = "[Step1 ルーター] 応答からチャンクIDを抽出できませんでした。" & vbLf & _
                       "リボン生応答(先頭400字):" & vbLf & Left$(routerOut, 400)
        GoTo Finish
    End If

    ' ---- Step 2: Feedback lookup -----------------------------------------
    Dim feedbackBlock As String
    feedbackBlock = modFeedbackLookup.FindSimilarApproved(question, deptId, user)

    ' ---- Step 3: Drafter --------------------------------------------------
    Dim maxCtx As Long: maxCtx = modConfig.GetLong("max_context_chars", 60000)
    Dim citations As String
    Dim context As String
    context = modKnowledgeBase.BuildContextFromIds(res.SelectedIds, maxCtx, citations)
    res.Citations = citations

    Dim drafterPrompt As String
    drafterPrompt = modPrompts.GetDrafter()
    drafterPrompt = drafterPrompt & vbLf & vbLf & "## 社内ナレッジ抜粋" & vbLf & context
    If LenB(feedbackBlock) > 0 Then
        drafterPrompt = drafterPrompt & vbLf & vbLf & feedbackBlock
    End If
    ' If follow-up: include the running conversation so the drafter continues
    ' the thread instead of starting over. History is capped to the last N turns.
    Dim hist As String: hist = modBoot.HistoryBlock()
    If modBoot.gFollowupMode And LenB(hist) > 0 Then
        drafterPrompt = drafterPrompt & vbLf & vbLf & "## これまでの会話（流れを踏まえて続けて回答）" & vbLf & hist
        drafterPrompt = drafterPrompt & vbLf & "## 続きの質問・深掘り" & vbLf & question & vbLf & vbLf & _
            "（指示）上の会話の流れを踏まえ、この続きの質問に答えてください。すでに説明済みの内容は繰り返さず、新しい論点に集中してください。"
    Else
        drafterPrompt = drafterPrompt & vbLf & vbLf & "## ユーザーの質問" & vbLf & question
    End If

    res.Draft = modRibbonGateway.CallLLM(drafterPrompt, "drafter", res.DraftMs)
    modRibbonGateway.LogDebugCall "drafter", drafterPrompt, res.Draft, res.DraftMs

    If Left$(res.Draft, 12) = "#LLM_ERROR: " Then
        res.OK = False
        res.ErrorMsg = "[Step3 ドラフト生成] 社内AIリボン呼び出しに失敗。" & vbLf & res.Draft
        GoTo Finish
    End If

    ' ---- Step 4: Verifier (if enabled) -----------------------------------
    Dim verifierOn As Boolean: verifierOn = modConfig.GetBool("verifier_enabled", True)
    If verifierOn Then
        Dim verifierPrompt As String
        verifierPrompt = modPrompts.GetVerifier()
        verifierPrompt = verifierPrompt & vbLf & vbLf & "## 引用元ナレッジ" & vbLf & context
        verifierPrompt = verifierPrompt & vbLf & vbLf & "## 元の質問" & vbLf & question
        verifierPrompt = verifierPrompt & vbLf & vbLf & "## ドラフト回答" & vbLf & res.Draft

        res.Answer = modRibbonGateway.CallLLM(verifierPrompt, "verifier", res.VerifyMs)
        modRibbonGateway.LogDebugCall "verifier", verifierPrompt, res.Answer, res.VerifyMs
        If Left$(res.Answer, 12) = "#LLM_ERROR: " Then
            ' Verifier failed: fall back to draft, but note it
            res.Answer = res.Draft & vbLf & vbLf & "(注: 検証パスでエラー発生。ドラフトを表示しています)"
        End If
    Else
        res.Answer = res.Draft
    End If

    res.TotalMs = CLng((Timer - tAll) * 1000)
    res.OK = True
    ' Trim citations to only the [#N] markers actually used in the final answer.
    ' Prevents leaking unused chunks (e.g., off-topic ones the router included).
    res.Citations = FilterUsedCitations(res.Citations, res.Answer)

Finish:
    ' CRITICAL: a UDT return value must be assigned explicitly, at every exit.
    RunQuery = res
    Exit Function

Trap:
    res.OK = False
    res.ErrorMsg = "[パイプライン内部エラー] Err " & Err.Number & ": " & Err.Description & vbLf & _
                   "選択ID: " & res.SelectedIds
    res.TotalMs = CLng((Timer - tAll) * 1000)
    RunQuery = res
End Function

' ----------------------------------------------------------------------------
' Scan the answer for [#N] markers actually used, then keep only matching
' lines from the citations block. Each citations line starts with "[#N] ".
' ----------------------------------------------------------------------------
Public Function FilterUsedCitations(ByVal cit As String, ByVal answer As String) As String
    If LenB(cit) = 0 Or LenB(answer) = 0 Then
        FilterUsedCitations = cit
        Exit Function
    End If

    ' Build a set of used N values by scanning the answer text.
    ' Markers may appear as [#1], [#1, #3], [#1,#3,#7], or [#1 #3]. An
    ' earlier "[#N" prefix scan missed every number after the first inside
    ' a combined marker (e.g. [#1, #3] dropped #3), which silently filtered
    ' legitimate citations out of the display. Scan all "#N" tokens INSIDE
    ' [...] brackets instead.
    Dim usedFlag As String      ' delimited "/1/3/5/" form
    Dim aLen As Long: aLen = Len(answer)
    Dim scanI As Long: scanI = 1
    Dim scanJ As Long
    Dim scanCh As String
    Dim scanDc As String
    Dim scanInBr As Boolean
    Dim scanNum As String
    Do While scanI <= aLen
        scanCh = Mid$(answer, scanI, 1)
        If scanCh = "[" Then
            scanInBr = True
        ElseIf scanCh = "]" Then
            scanInBr = False
        ElseIf scanInBr And scanCh = "#" Then
            scanNum = ""
            scanJ = scanI + 1
            Do While scanJ <= aLen
                scanDc = Mid$(answer, scanJ, 1)
                If scanDc >= "0" And scanDc <= "9" Then
                    scanNum = scanNum & scanDc
                    scanJ = scanJ + 1
                Else
                    Exit Do
                End If
            Loop
            If LenB(scanNum) > 0 Then
                usedFlag = usedFlag & "/" & scanNum & "/"
            End If
            scanI = scanJ - 1
        End If
        scanI = scanI + 1
    Loop

    If LenB(usedFlag) = 0 Then
        ' No markers found in answer -- show original citations as fallback
        FilterUsedCitations = cit
        Exit Function
    End If

    Dim lines() As String: lines = Split(cit, vbLf)
    Dim out As String, i As Long
    For i = LBound(lines) To UBound(lines)
        Dim line As String: line = lines(i)
        If LenB(line) = 0 Then GoTo NextLine
        ' Extract leading [#N]
        Dim p As Long: p = InStr(line, "[#")
        If p = 0 Then GoTo NextLine
        Dim q As Long: q = InStr(p, line, "]")
        If q = 0 Then GoTo NextLine
        Dim numStr As String: numStr = Mid$(line, p + 2, q - p - 2)
        If Not IsNumeric(numStr) Then GoTo NextLine
        If InStr(usedFlag, "/" & numStr & "/") > 0 Then
            out = out & line & vbLf
        End If
NextLine:
    Next i
    FilterUsedCitations = out
End Function

' ----------------------------------------------------------------------------
' Parse the router's JSON output into a comma-separated list of chunk IDs.
' Accepts either:
'   {"selected_ids": ["a", "b"], ...}
'   ["a", "b"]
' Returns "" if cannot parse.
' ----------------------------------------------------------------------------
Public Function ParseSelectedIds(ByVal raw As String) As String
    Dim s As String: s = raw
    ' Strip code fence if present
    s = Replace(s, "```json", "")
    s = Replace(s, "```", "")
    s = Trim$(s)
    ' Find selected_ids array
    Dim startIdx As Long, endIdx As Long
    startIdx = InStr(1, s, "selected_ids", vbTextCompare)
    If startIdx > 0 Then
        startIdx = InStr(startIdx, s, "[")
    Else
        startIdx = InStr(1, s, "[")
    End If
    If startIdx = 0 Then Exit Function
    endIdx = InStr(startIdx, s, "]")
    If endIdx = 0 Then Exit Function
    Dim arr As String: arr = Mid$(s, startIdx + 1, endIdx - startIdx - 1)

    ' Extract quoted ids
    Dim sb As String, i As Long, depth As Long
    Dim inQuote As Boolean
    Dim qStart As Long
    For i = 1 To Len(arr)
        Dim ch As String: ch = Mid$(arr, i, 1)
        If ch = """" Then
            If Not inQuote Then
                inQuote = True
                qStart = i + 1
            Else
                Dim tok As String: tok = Mid$(arr, qStart, i - qStart)
                If LenB(sb) > 0 Then sb = sb & ","
                sb = sb & Trim$(tok)
                inQuote = False
            End If
        End If
    Next i
    ParseSelectedIds = sb
End Function
