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
    Intent As String            ' from intent layer
    Assumptions As String       ' from intent layer
    Confidence As String        ' 高 / 中 / 低 (parsed from verifier)
    Followups As String         ' 深掘り候補 (vbLf sep, parsed from verifier)
    RouterMs As Long
    DraftMs As Long
    VerifyMs As Long
    TotalMs As Long
    ErrorMsg As String
End Type

' ir carries the intent-layer analysis (may be a blank IntentResult if the layer
' is disabled); extraContext is the user's answer to any clarifying questions.
Public Function RunQuery(ByVal question As String, _
                         ByRef ir As modIntent.IntentResult, _
                         ByVal extraContext As String) As PipelineResult
    Dim res As PipelineResult
    res.Question = question
    res.Intent = ir.Intent
    res.Assumptions = ir.Assumptions
    Dim tAll As Double: tAll = Timer
    On Error GoTo Trap

    ' ---- Easter egg: "誰が作った？" は約款と無関係なので最初に横取り ----
    If IsCreatorQuestion(question) Then
        res.Answer = CreatorAnswer()
        res.Confidence = "高"
        res.OK = True
        res.TotalMs = CLng((Timer - tAll) * 1000)
        RunQuery = res
        Exit Function
    End If

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
    ' Intent layer: append decomposed sub-queries so the router retrieves chunks
    ' for every facet of the question, not just its surface wording.
    If LenB(ir.SearchQueries) > 0 Then
        routerQuestion = routerQuestion & vbLf & _
            "【関連する論点（検索補助）】" & vbLf & ir.SearchQueries
    End If
    routerPrompt = Replace(routerPrompt, "{question}", routerQuestion)
    routerPrompt = Replace(routerPrompt, "{max_n}", CStr(routerN))
    routerPrompt = Replace(routerPrompt, "{knowledge_table}", routerTable)

    SetBar "🔎 関連する社内ナレッジを検索中... (2/4)"
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
    ' Intent layer output: tell the drafter the analysed intent/assumptions so it
    ' answers the REAL question and leads with the premises it assumed.
    If LenB(ir.Intent) > 0 Or LenB(ir.Assumptions) > 0 Then
        drafterPrompt = drafterPrompt & vbLf & vbLf & _
            "## 照会の意図・前提（意図解釈レイヤーの解析。これを踏まえて回答する）"
        If LenB(ir.Intent) > 0 Then drafterPrompt = drafterPrompt & vbLf & "意図: " & ir.Intent
        If LenB(ir.Assumptions) > 0 Then drafterPrompt = drafterPrompt & vbLf & "前提: " & ir.Assumptions
        drafterPrompt = drafterPrompt & vbLf & _
            "（指示）回答の冒頭に「【前提】…」を1行で明示し、不確実な前提は留意点に書く。"
    End If
    ' Clarification answers the user typed back (if any).
    If LenB(extraContext) > 0 Then
        drafterPrompt = drafterPrompt & vbLf & vbLf & _
            "## 照会者からの補足（前提確認への回答。これを最優先で反映）" & vbLf & extraContext
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

    SetBar "✍ 約款・ガイドラインを踏まえて回答を作成中... (3/4)"
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

        SetBar "🔍 事実と出典を照合・検証中... (4/4)"
        res.Answer = modRibbonGateway.CallLLM(verifierPrompt, "verifier", res.VerifyMs)
        modRibbonGateway.LogDebugCall "verifier", verifierPrompt, res.Answer, res.VerifyMs
        If Left$(res.Answer, 12) = "#LLM_ERROR: " Then
            ' Verifier failed: fall back to draft, but note it
            res.Answer = res.Draft & vbLf & vbLf & "(注: 検証パスでエラー発生。ドラフトを表示しています)"
        End If
    Else
        res.Answer = res.Draft
    End If

    ' Extract the verifier's machine-readable trailing lines (confidence + follow-ups)
    ' and strip them from the displayed answer.
    ParseTrailers res
    If LenB(res.Confidence) = 0 Then res.Confidence = "中"   ' default when verifier off

    res.TotalMs = CLng((Timer - tAll) * 1000)
    res.OK = True
    ' Trim citations to only the [#N] markers actually used in the final answer.
    ' Prevents leaking unused chunks (e.g., off-topic ones the router included).
    res.Citations = FilterUsedCitations(res.Citations, res.Answer)

Finish:
    SetBar False
    ' CRITICAL: a UDT return value must be assigned explicitly, at every exit.
    RunQuery = res
    Exit Function

Trap:
    SetBar False
    res.OK = False
    res.ErrorMsg = "[パイプライン内部エラー] Err " & Err.Number & ": " & Err.Description & vbLf & _
                   "選択ID: " & res.SelectedIds
    res.TotalMs = CLng((Timer - tAll) * 1000)
    RunQuery = res
End Function

' Update Excel's status bar so the user sees live progress during the
' (otherwise UI-blocking) LLM calls. Never raises.
Private Sub SetBar(ByVal msg As Variant)
    On Error Resume Next
    Application.StatusBar = msg
    On Error GoTo 0
End Sub

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
' Pull [[CONFIDENCE:..]] and [[FOLLOWUP: a | b | c]] off the end of the answer,
' set res.Confidence / res.Followups, and remove those lines from res.Answer.
' Tolerant of the verifier omitting them (older prompt / verifier off).
' ----------------------------------------------------------------------------
Public Sub ParseTrailers(ByRef res As PipelineResult)
    Dim a As String: a = res.Answer
    If LenB(a) = 0 Then Exit Sub

    ' CONFIDENCE
    Dim p As Long: p = InStr(1, a, "[[CONFIDENCE:", vbTextCompare)
    If p > 0 Then
        Dim q As Long: q = InStr(p, a, "]]")
        If q > 0 Then
            Dim cv As String: cv = Mid$(a, p + Len("[[CONFIDENCE:"), q - p - Len("[[CONFIDENCE:"))
            cv = Trim$(cv)
            If InStr(cv, "高") > 0 Then
                res.Confidence = "高"
            ElseIf InStr(cv, "低") > 0 Then
                res.Confidence = "低"
            ElseIf InStr(cv, "中") > 0 Then
                res.Confidence = "中"
            End If
        End If
    End If

    ' FOLLOWUP
    Dim fp As Long: fp = InStr(1, a, "[[FOLLOWUP:", vbTextCompare)
    If fp > 0 Then
        Dim fq As Long: fq = InStr(fp, a, "]]")
        If fq > 0 Then
            Dim fv As String: fv = Mid$(a, fp + Len("[[FOLLOWUP:"), fq - fp - Len("[[FOLLOWUP:"))
            fv = Trim$(fv)
            If StrComp(fv, "なし", vbTextCompare) <> 0 And LenB(fv) > 0 Then
                Dim parts() As String: parts = Split(fv, "|")
                Dim out As String, i As Long
                For i = LBound(parts) To UBound(parts)
                    Dim t As String: t = Trim$(parts(i))
                    If LenB(t) > 0 Then
                        If LenB(out) > 0 Then out = out & vbLf
                        out = out & t
                    End If
                Next i
                res.Followups = out
            End If
        End If
    End If

    ' Strip every line that contains a [[...]] machine marker from the display text.
    Dim lines() As String: lines = Split(a, vbLf)
    Dim keep As String, j As Long
    For j = LBound(lines) To UBound(lines)
        If InStr(lines(j), "[[CONFIDENCE:") = 0 And InStr(lines(j), "[[FOLLOWUP:") = 0 Then
            If LenB(keep) > 0 Then keep = keep & vbLf
            keep = keep & lines(j)
        End If
    Next j
    ' Trim trailing blank lines
    Do While Len(keep) > 0 And (Right$(keep, 1) = vbLf Or Right$(keep, 1) = vbCr Or Right$(keep, 1) = " ")
        keep = Left$(keep, Len(keep) - 1)
    Loop
    res.Answer = keep
End Sub

' ----------------------------------------------------------------------------
' Hidden credit: questions about who built this bot are answered directly,
' independent of the knowledge base. Kept tight (short question + co-occurring
' "this bot/tool/AI" + "made/developed/who") to avoid false positives on
' genuine insurance questions.
' ----------------------------------------------------------------------------
Public Function IsCreatorQuestion(ByVal q As String) As Boolean
    Dim s As String: s = LCase$(q)
    If Len(s) > 60 Then Exit Function                 ' creator questions are short
    Dim subjectHit As Boolean, verbHit As Boolean
    If InStr(s, "このチャット") > 0 Or InStr(s, "このbot") > 0 Or InStr(s, "このボット") > 0 _
       Or InStr(s, "このツール") > 0 Or InStr(s, "このアプリ") > 0 Or InStr(s, "このシステム") > 0 _
       Or InStr(s, "君") > 0 Or InStr(s, "あなた") > 0 Or InStr(s, "notebook") > 0 _
       Or InStr(s, "このai") > 0 Or InStr(s, "このチャットボット") > 0 Then subjectHit = True
    If InStr(s, "作っ") > 0 Or InStr(s, "作成") > 0 Or InStr(s, "開発") > 0 Or InStr(s, "制作") > 0 _
       Or InStr(s, "誰が") > 0 Or InStr(s, "who made") > 0 Or InStr(s, "who built") > 0 Then verbHit = True
    If InStr(s, "開発者") > 0 Or InStr(s, "作者") > 0 Then
        IsCreatorQuestion = True
        Exit Function
    End If
    IsCreatorQuestion = (subjectHit And verbHit)
End Function

Public Function CreatorAnswer() As String
    Dim nm As String: nm = modConfig.GetString("creator_name", "小島正豪")
    Dim grp As String: grp = modConfig.GetString("creator_group", "ニューリスクG")
    CreatorAnswer = "## このチャットボットについて 🤖" & vbLf & vbLf & _
        "このチャットボットは、" & grp & " のAIプロフェッショナル **" & nm & "** さんが開発しました。" & vbLf & vbLf & _
        "営業の照会対応を効率化し、本社引受部門の負荷を下げ、ナレッジを全社で資産化することを" & _
        "目的に、約款・引受ガイドライン・商品部公式Q&Aを根拠とした回答と、本社への照会支援" & _
        "（教えてBOX）を一体で設計しています。" & vbLf & vbLf & _
        "© " & grp & " " & nm
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
