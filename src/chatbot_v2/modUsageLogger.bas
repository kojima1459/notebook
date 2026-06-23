Attribute VB_Name = "modUsageLogger"
Option Explicit

' ============================================================================
' modUsageLogger - Append one row per query to `usage_log` sheet
' ----------------------------------------------------------------------------
' Cols:
'   A: timestamp
'   B: dept_id
'   C: user
'   D: question_chars
'   E: answer_chars
'   F: router_ms
'   G: draft_ms
'   H: verify_ms
'   I: total_ms
'   J: chunks_used    (comma sep ids)
'   K: feedback_status   (set later by feedback handlers if good/bad clicked)
' ============================================================================

Public Sub LogQuery(ByVal q As String, ByRef res As modPipeline.PipelineResult)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("usage_log")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Dim r As Long: r = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1).value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(r, 2).value = modUserProfile.CurrentDept()
    ws.Cells(r, 3).value = modUserProfile.CurrentUser()
    ws.Cells(r, 4).value = Len(q)
    ws.Cells(r, 5).value = Len(res.Answer)
    ws.Cells(r, 6).value = res.RouterMs
    ws.Cells(r, 7).value = res.DraftMs
    ws.Cells(r, 8).value = res.VerifyMs
    ws.Cells(r, 9).value = res.TotalMs
    ws.Cells(r, 10).value = res.SelectedIds
    ws.Cells(r, 11).value = ""
End Sub
