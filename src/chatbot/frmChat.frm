VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChat
   Caption         =   "社内ナレッジチャット"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10800
   OleObjectBlob   =   "frmChat.frx":0000
   StartUpPosition =   1  'Owner
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' frmChat - main chat UI
' ----------------------------------------------------------------------------
' Controls (created by build.ps1 if the .frx scaffold is missing):
'   txtQuestion       (TextBox, multiline)
'   btnSend           (CommandButton)
'   txtConversation   (TextBox, multiline, locked)
'   chkConsent        (CheckBox: ナレッジ改善に質問本文を提供する)
'   lblStatus         (Label)
' ============================================================================

Private Sub UserForm_Initialize()
    txtConversation.Text = LoadDisclaimer() & vbCrLf & vbCrLf & _
                           "（質問を入力して『送信』を押してください）" & vbCrLf
    chkConsent.value = modUserProfile.ShareQuestionConsent()
    lblStatus.Caption = "Ready"
End Sub

Private Sub chkConsent_Click()
    modUserProfile.SetShareQuestionConsent CBool(chkConsent.value)
End Sub

Private Sub btnSend_Click()
    If Not modBoot.gReady Then
        MsgBox "indexが読み込まれていません。", vbExclamation
        Exit Sub
    End If

    Dim q As String: q = Trim$(txtQuestion.Text)
    If Len(q) = 0 Then Exit Sub

    ' Rate-limit check
    Dim retry As Long
    If Not modRateLimiter.CanSend(retry) Then
        MsgBox "本日の質問上限に達しました。約 " & retry & " 秒後に再度お試しください。", vbExclamation
        Exit Sub
    End If

    ' PII check
    Dim piiWarn As Boolean
    piiWarn = modPiiGuard.DetectPii(q)
    If piiWarn Then
        If MsgBox("質問に個人情報が含まれている可能性があります。送信を続けますか？", _
                  vbYesNo + vbExclamation, "確認") <> vbYes Then
            Exit Sub
        End If
    End If

    lblStatus.Caption = "問い合わせ中…"
    btnSend.Enabled = False
    DoEvents

    Dim res As AnswerResult
    res = modRagEngine.Answer(modBoot.gIndex, q)

    If res.OK Then
        Append "あなた: " & q
        Append "AI: " & res.Answer
        Append FormatCitations(res.Citations)
        Append ""
        modRateLimiter.RecordSend
        modUsageLogger.LogUsage _
            res.PromptTokens, res.CompletionTokens, _
            res.EmbedLatencyMs + res.ChatLatencyMs, _
            piiWarn, _
            CBool(chkConsent.value), q
        lblStatus.Caption = "Ready (" & res.PromptTokens & " / " & res.CompletionTokens & " tokens, " _
            & (res.EmbedLatencyMs + res.ChatLatencyMs) & " ms)"
    Else
        Append "[エラー] " & res.ErrorMessage
        lblStatus.Caption = "Error"
    End If

    btnSend.Enabled = True
    txtQuestion.Text = ""
End Sub

Private Sub Append(ByVal s As String)
    txtConversation.Text = txtConversation.Text & s & vbCrLf
    txtConversation.SelStart = Len(txtConversation.Text)
End Sub

Private Function FormatCitations(ByRef c() As Citation) As String
    If LBound(c) > UBound(c) Then Exit Function
    Dim sb As String, i As Long
    sb = "  出典:"
    For i = LBound(c) To UBound(c)
        sb = sb & vbCrLf & "   [#" & (i + 1) & "] " & c(i).Source
        If c(i).page > 0 Then sb = sb & " p." & c(i).page
        sb = sb & "  (score " & Format$(c(i).Score, "0.000") & ")"
    Next i
    FormatCitations = sb
End Function

Private Function LoadDisclaimer() As String
    Dim path As String
    path = ThisWorkbook.Path & "\" & modConfig.GetString("ui", "disclaimer_path", "disclaimer.txt")
    If Len(Dir$(path)) = 0 Then Exit Function
    Dim fnum As Integer: fnum = FreeFile
    Open path For Input As #fnum
    Dim buf As String, line As String
    Do While Not EOF(fnum)
        Line Input #fnum, line
        buf = buf & line & vbCrLf
    Loop
    Close #fnum
    LoadDisclaimer = buf
End Function
