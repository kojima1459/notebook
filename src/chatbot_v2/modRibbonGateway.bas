Attribute VB_Name = "modRibbonGateway"
Option Explicit

' ============================================================================
' modRibbonGateway - Single entry point for calling the corporate AI ribbon
' ----------------------------------------------------------------------------
' The AI ribbon (リボンちゃん.xlam / MSAD-Addin.xlam) exposes:
'   Application.Run("ChatGPT", prompt) -> response string
' All LLM calls in this workbook go through this gateway so:
'   - Failure handling is centralized
'   - Future model-switching can be added in one place
'   - Mock/debug mode can intercept without touching pipeline logic
' ============================================================================

Public Function CallLLM(ByVal prompt As String, _
                        Optional ByVal step_name As String = "(unspecified)", _
                        Optional ByRef latency_ms As Long = 0) As String
    Dim t0 As Double: t0 = Timer
    On Error GoTo ErrHandler

    Dim result As Variant
    result = Application.Run("ChatGPT", prompt)
    CallLLM = CStr(result)

    latency_ms = CLng((Timer - t0) * 1000)
    Exit Function

ErrHandler:
    CallLLM = "#LLM_ERROR: step=" & step_name & " err=" & Err.Description
    latency_ms = CLng((Timer - t0) * 1000)
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
