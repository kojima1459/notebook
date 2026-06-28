Attribute VB_Name = "modRibbonGateway"
Option Explicit

' ============================================================================
' modRibbonGateway - Single entry point for calling the corporate AI ribbon
' ----------------------------------------------------------------------------
' Confirmed ribbon ChatGPT signature (excel-addin リボンちゃん ver202606, GPT.bas):
'   Function ChatGPT(text, [roleSystem], [Temperature As Double=0.4],
'                    [MaxTokens As Long], [Wait As Long=120], [optModel],
'                    [prevU], [prevA], [toolN], [reasoning_effort], [verbosity])
' Key facts that shape how we call it:
'   - arg6 optModel selects the model (we read it from config recommended_model).
'   - For GPT-5 models (gpt-5.x) the ribbon IGNORES Temperature/top_p and instead
'     honours reasoning_effort (low/medium/high) and verbosity (low/medium/high).
'     Those two are our real quality levers, so we set them PER STEP.
'   - args 3-5 are typed Double/Long; passing "" raises a type-mismatch. We pass
'     explicit safe values (0.4, 0, 1200) instead. Wait=1200 because gpt-5.5 is a
'     slow reasoning model (the ribbon itself bumps its default to 1200s).
'   - toolN (arg9) is a free log label; we pass the step name for the ribbon log.
' All LLM calls in this workbook go through this gateway so model/effort tuning,
' failure handling and mock/debug interception live in one place.
' ============================================================================

Public Function CallLLM(ByVal prompt As String, _
                        Optional ByVal step_name As String = "(unspecified)", _
                        Optional ByRef latency_ms As Long = 0) As String
    Dim t0 As Double: t0 = Timer
    On Error GoTo ErrHandler

    ' Mock mode: exercise the whole pipeline/UI without the corporate AI ribbon.
    ' Turn on by setting config key `mock_llm` = TRUE. Use this to verify the
    ' app on a Mac, or to smoke-test on a corporate PC before the ribbon's
    ' argument convention has been confirmed.
    If modConfig.GetBool("mock_llm", False) Then
        CallLLM = MockResponse(prompt, step_name)
        latency_ms = CLng((Timer - t0) * 1000)
        Exit Function
    End If

    ' Resolve model: config stores "GPT-5.5" (display), ribbon expects "gpt-5.5".
    Dim mdl As String: mdl = LCase$(Trim$(modConfig.GetString("recommended_model", "gpt-5.5")))
    If LenB(mdl) = 0 Then mdl = "gpt-5.5"

    ' Per-step reasoning depth / verbosity (only honoured by the ribbon for GPT-5).
    ' Cheap selection steps run shallow & terse; the fact-check verifier runs deep.
    Dim eff As String, vrb As String
    StepParams step_name, eff, vrb
    ' Escape hatch: if a non-GPT-5 model is configured, or to force ribbon defaults.
    If Not modConfig.GetBool("reasoning_tuning", True) Then
        eff = "": vrb = ""
    End If

    ' Full positional call. Explicit numeric args avoid type-mismatch on the
    ' Double/Long parameters; "" is only used for the genuine String slots.
    '            text  roleSys Temp MaxTok Wait  model prevU prevA toolN       effort verbosity
    Dim result As Variant
    result = Application.Run("ChatGPT", prompt, "", 0.4, 0, 1200, mdl, "", "", step_name, eff, vrb)
    CallLLM = CStr(result)

    latency_ms = CLng((Timer - t0) * 1000)
    Exit Function

ErrHandler:
    CallLLM = "#LLM_ERROR: step=" & step_name & " err=" & Err.Description
    latency_ms = CLng((Timer - t0) * 1000)
End Function

' Map a pipeline step to GPT-5 reasoning_effort + verbosity.
'   intent / router : shallow & terse  -> fast, this is classification/selection
'   drafter         : medium reasoning, high verbosity -> a thorough written answer
'   verifier        : deep reasoning, medium verbosity -> precision-critical fact check
Private Sub StepParams(ByVal step_name As String, ByRef effort As String, ByRef verbosity As String)
    Select Case LCase$(step_name)
        Case "intent":   effort = "low":    verbosity = "low"
        Case "router":   effort = "low":    verbosity = "low"
        Case "drafter":  effort = "medium": verbosity = "high"
        Case "verifier": effort = "high":   verbosity = "medium"
        Case Else:       effort = "":       verbosity = ""
    End Select
End Sub

' ----------------------------------------------------------------------------
' MockResponse - canned, format-correct replies so the pipeline runs offline.
'   router   -> JSON {"selected_ids":[ first N ids found in the prompt ]}
'   drafter  -> a structured Japanese answer that cites [#1]
'   verifier -> returns the draft unchanged (passthrough)
' ----------------------------------------------------------------------------
Private Function MockResponse(ByVal prompt As String, ByVal step_name As String) As String
    Select Case LCase$(step_name)
        Case "intent"
            MockResponse = "{""intent"": ""(mock) 照会の意図"", ""assumptions"": [""(mock)前提""], " & _
                           """needs_clarification"": false, ""clarifying_questions"": [], " & _
                           """search_queries"": [""(mock)サブクエリ1"", ""(mock)サブクエリ2""]}"
        Case "router"
            MockResponse = "{""selected_ids"": [" & MockPickIds(prompt, 4) & "], " & _
                           """reasoning"": ""(mock) picked leading chunks""}"
        Case "verifier"
            ' Passthrough: the draft is the text after the last "## ドラフト回答",
            ' plus the machine-readable trailers the real verifier emits.
            Dim p As Long: p = InStrRev(prompt, "## ドラフト回答")
            Dim drafted As String
            If p > 0 Then
                drafted = Trim$(Mid$(prompt, p + Len("## ドラフト回答")))
            Else
                drafted = prompt
            End If
            MockResponse = drafted & vbLf & "[[CONFIDENCE:中]]" & vbLf & _
                           "[[FOLLOWUP: (mock)深掘り例1 | (mock)深掘り例2]]"
        Case Else   ' drafter
            MockResponse = "【モック回答】これはリボン未接続の動作確認用ダミー回答です。" & vbLf & vbLf & _
                           "■ 結論" & vbLf & _
                           "・ 実際のAI回答はここに表示されます [#1]" & vbLf & _
                           "・ 送信→ルーター→ドラフト→検証 のパイプラインは正常に動作しています [#1]" & vbLf & vbLf & _
                           "■ 次のアクション" & vbLf & _
                           "・ config シートの mock_llm を FALSE に戻すと、社内AIリボンを実呼び出しします。"
    End Select
End Function

' Extract up to n chunk-id tokens (pattern "..::pN::cN") from the router prompt.
Private Function MockPickIds(ByVal prompt As String, ByVal n As Long) As String
    Dim lines() As String: lines = Split(prompt, vbLf)
    Dim out As String, taken As Long
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim ln As String: ln = lines(i)
        Dim bar As Long: bar = InStr(ln, " | ")
        If bar > 0 Then
            Dim idTok As String: idTok = Trim$(Left$(ln, bar - 1))
            If InStr(idTok, "::p") > 0 And InStr(idTok, "::c") > 0 Then
                If taken > 0 Then out = out & ", "
                out = out & """" & idTok & """"
                taken = taken + 1
                If taken >= n Then Exit For
            End If
        End If
    Next i
    MockPickIds = out
End Function

' ----------------------------------------------------------------------------
' For debug: append every call's prompt/response to a log sheet (if enabled)
' Use sparingly - LLM prompts are huge.
' ----------------------------------------------------------------------------
Public Sub LogDebugCall(ByVal step_name As String, ByVal prompt As String, ByVal response As String, ByVal latency_ms As Long)
    If Not modConfig.GetBool("debug_mode", False) Then Exit Sub
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("debug_log")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = "debug_log"
        ws.Range("A1:E1").value = Array("timestamp", "step", "latency_ms", "prompt", "response")
    End If
    On Error GoTo 0
    Dim r As Long: r = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1).value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(r, 2).value = step_name
    ws.Cells(r, 3).value = latency_ms
    ws.Cells(r, 4).value = Left$(prompt, 30000)
    ws.Cells(r, 5).value = Left$(response, 30000)
End Sub
