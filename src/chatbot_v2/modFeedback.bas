Attribute VB_Name = "modFeedback"
Option Explicit

' ============================================================================
' modFeedback - post-answer reaction handlers ("この回答どうでした？")
' ----------------------------------------------------------------------------
' Reframed from a 正誤判定 (○正しい / ×間違い) model, which felt like grading the
' AI and went unused, to a "how did it land for you, what's your next step?"
' model. Three reactions, each routing to the natural next action:
'
'   🟢 スッキリ解決   React_Cleared  -> status="pending"  (= reusable model answer,
'                                       same learning value as the old ○)
'   🟡 もう少し知りたい React_More     -> status="more"     + 『続けて質問』へ
'   🔴 まだ不安・要確認 React_Unsure   -> status="unsure"   + 『教えてBOX』へ誘導
'
' Plus an explicit, gentle correction path (not a scary primary button):
'   ✏ 正しい答えを知っている Submit_Correction -> status="rejected" + correction
'     (= the old × correction path; keeps the high-value correction learning)
'
' LEARNING-LOOP SAFETY: modFeedbackLookup.FindSimilarApproved only ever reuses
' rows with status "approved" or (own) "pending", and applies col-J correction
' on "rejected" rows. The new statuses "more" / "unsure" are therefore INERT to
' retrieval — they record the signal for analytics/triage without ever polluting
' the answer pool. This was verified against modFeedbackLookup before shipping.
'
' Sheet `feedback` columns (must match modFeedbackLookup):
'   A: fb_id   B: timestamp   C: dept_id   D: submitter   E: question
'   F: answer  G: status      H: approver  I: approved_at J: correction
'   K: tags    L: source_chunk_ids
' ============================================================================

Private Const SHEET_NAME As String = "feedback"

' --- 🟢 スッキリ解決 : the answer helped. Record as a reusable model answer. ---
Public Sub React_Cleared()
    If LenB(modBoot.gLastQuestion) = 0 Then
        MsgBox "最新の質問・回答が見つかりません。", vbExclamation: Exit Sub
    End If
    Dim tag As String
    tag = InputBox("スッキリ解決ですね！ どの点が役に立ちましたか？（任意・1行）" & vbCrLf & _
                   "※ 空欄でもOK。書いておくと、似た質問への回答精度が上がります。", _
                   "🟢 スッキリ解決")
    AppendFeedback "pending", modBoot.gLastQuestion, modBoot.gLastAnswer, _
                   modBoot.gLastSelectedIds, "", tag
    MsgBox "ありがとうございます。この回答を“役に立った例”として記録しました。" & vbCrLf & _
           "次に似た質問が来たら、この回答を優先的に参照します。", vbInformation, "🟢 記録しました"
End Sub

' --- 🟡 もう少し知りたい : not wrong, just wants to go deeper. Route to follow-up. ---
Public Sub React_More()
    If LenB(modBoot.gLastQuestion) = 0 Then
        MsgBox "最新の質問・回答が見つかりません。", vbExclamation: Exit Sub
    End If
    Dim fu As String
    fu = InputBox("もう少し知りたいんですね。どこを深掘りしたいですか？" & vbCrLf & _
                  "（入力するとそのまま続けて質問します。空欄なら深掘りのヒントを表示します）", _
                  "🟡 もう少し知りたい")
    ' Record the signal (inert to the learning loop) for triage/analytics.
    AppendFeedback "more", modBoot.gLastQuestion, modBoot.gLastAnswer, _
                   modBoot.gLastSelectedIds, "", fu
    If LenB(Trim$(fu)) > 0 Then
        modChatUI.RunFollowup fu          ' deepen on exactly what they asked
    Else
        modChatUI.OnFollowupClick         ' open the normal follow-up helper (with examples)
    End If
End Sub

' --- 🔴 まだ不安・要確認 : didn't fully land. Capture what's unclear, offer escalation. ---
Public Sub React_Unsure()
    If LenB(modBoot.gLastQuestion) = 0 Then
        MsgBox "最新の質問・回答が見つかりません。", vbExclamation: Exit Sub
    End If
    Dim note As String
    note = InputBox("どの点が不安・引っかかっていますか？（任意・1行）" & vbCrLf & _
                    "※ ここで整理しておくと、本社へ照会するときもスムーズです。", _
                    "🔴 まだ不安・要確認")
    AppendFeedback "unsure", modBoot.gLastQuestion, modBoot.gLastAnswer, _
                   modBoot.gLastSelectedIds, "", note
    If MsgBox("記録しました。" & vbCrLf & vbCrLf & _
              "本社の『教えてBOX』で照会しますか？" & vbCrLf & _
              "（あなたの質問・AIが調べた範囲・該当しそうな条文を自動で整理します）", _
              vbYesNo + vbQuestion, "🔴 本社へ照会しますか？") = vbYes Then
        modInquiryBox.OpenInquiryBox
    End If
End Sub

' --- ✏ 正しい答えを知っている : gentle correction path (the old ×, de-emphasised). ---
Public Sub Submit_Correction()
    If LenB(modBoot.gLastQuestion) = 0 Then
        MsgBox "最新の質問・回答が見つかりません。", vbExclamation: Exit Sub
    End If
    Dim correction As String
    correction = InputBox("正しい内容・あるべき回答を入力してください。" & vbCrLf & _
                          "（条文番号・参照箇所を書いておくと、今後のナレッジ蓄積に役立ちます）", _
                          "✏ 正しい答えを登録")
    If LenB(correction) = 0 Then
        MsgBox "入力が空のため、記録しませんでした。", vbInformation
        Exit Sub
    End If
    AppendFeedback "rejected", modBoot.gLastQuestion, modBoot.gLastAnswer, _
                   modBoot.gLastSelectedIds, correction, ""
    MsgBox "正しい回答として記録しました。" & vbCrLf & _
           "次に似た質問が来たら、こちらを優先的に参照します。", vbInformation, "✏ 記録しました"
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
