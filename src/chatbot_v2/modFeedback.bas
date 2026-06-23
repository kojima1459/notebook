Attribute VB_Name = "modFeedback"
Option Explicit

' ============================================================================
' modFeedback - ○/× button handlers
' ----------------------------------------------------------------------------
' On ○: append row to feedback with status=pending. Pending entries are
'        immediately valid for the submitter (per FindSimilarApproved logic),
'        so they shape the user's own future queries.
' On ×: ask "正しい回答は？", append row with status=rejected + correction.
'        Correction is what gets used as the future trusted answer.
'
' Sheet `feedback` columns (must match modFeedbackLookup):
'   A: fb_id          (timestamp+user hash)
'   B: timestamp
'   C: dept_id
'   D: submitter
'   E: question
'   F: answer
'   G: status         (pending / approved / rejected)
'   H: approver
'   I: approved_at
'   J: correction
'   K: tags
'   L: source_chunk_ids   (the comma-sep IDs used for this answer)
' ============================================================================

Private Const SHEET_NAME As String = "feedback"

Public Sub Submit_Feedback_Good()
    If LenB(modBoot.gLastQuestion) = 0 Then
        MsgBox "最新の質問・回答が見つかりません。", vbExclamation: Exit Sub
    End If
    Dim tag As String
    tag = InputBox("この回答が良かった理由 (任意・1行)" & vbCrLf & "後で管理者承認時の手掛かりになります。", _
                   "○ フィードバック")
    AppendFeedback "pending", modBoot.gLastQuestion, modBoot.gLastAnswer, _
                   modBoot.gLastSelectedIds, "", tag
    MsgBox "フィードバックを記録しました。" & vbCrLf & _
           "同様の質問が次回来たら、この回答を優先的に参照します。", vbInformation
End Sub

Public Sub Submit_Feedback_Bad()
    If LenB(modBoot.gLastQuestion) = 0 Then
        MsgBox "最新の質問・回答が見つかりません。", vbExclamation: Exit Sub
    End If
    Dim correction As String
    correction = InputBox("正しい回答を入力してください。" & vbCrLf & _
                          "（マニュアル参照部分や条文番号を明記すると、後のナレッジ蓄積に役立ちます）", _
                          "× フィードバック - 矯正")
    If LenB(correction) = 0 Then
        MsgBox "矯正版が空のため、フィードバックを記録しませんでした。", vbInformation
        Exit Sub
    End If
    AppendFeedback "rejected", modBoot.gLastQuestion, modBoot.gLastAnswer, _
                   modBoot.gLastSelectedIds, correction, ""
    MsgBox "矯正版を記録しました。" & vbCrLf & _
           "同様の質問が次回来たら、矯正版を優先的に参照します。", vbInformation
End Sub

Private Sub AppendFeedback(ByVal status As String, ByVal q As String, ByVal a As String, _
                           ByVal sourceIds As String, ByVal correction As String, ByVal tag As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "feedback シートが見つかりません。", vbCritical: Exit Sub
    End If
    Dim r As Long: r = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    Dim fbId As String: fbId = Format$(Now, "yyyymmdd_hhnnss") & "_" & modUserProfile.CurrentUser()
    ws.Cells(r, 1).value = fbId
    ws.Cells(r, 2).value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(r, 3).value = modUserProfile.CurrentDept()
    ws.Cells(r, 4).value = modUserProfile.CurrentUser()
    ws.Cells(r, 5).value = q
    ws.Cells(r, 6).value = a
    ws.Cells(r, 7).value = status
    ws.Cells(r, 8).value = ""    ' approver
    ws.Cells(r, 9).value = ""    ' approved_at
    ws.Cells(r, 10).value = correction
    ws.Cells(r, 11).value = tag
    ws.Cells(r, 12).value = sourceIds

    ' Save workbook so feedback persists immediately
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0
End Sub
