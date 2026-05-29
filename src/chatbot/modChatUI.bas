Attribute VB_Name = "modChatUI"
Option Explicit

' ============================================================================
' modChatUI - in-worksheet chat interface (no UserForm required)
' ----------------------------------------------------------------------------
' Builds the entire chat UI on a worksheet named "ChatUI" so the build
' pipeline does not need to ship a binary UserForm. Users get something
' that looks like normal Excel: type in a cell, click a button, see the
' conversation grow downward.
'
' Layout (column widths set in EnsureLayout):
'   A1:F1   Title banner (merged)
'   A3      "質問:"
'   B3:E4   Question input cell (merged, multi-line)
'   F3:F4   "▶ 送信" button (Form Control, runs SendButton_Click)
'   A6      "会話履歴"
'   A7 ...  Rolling conversation log
'   I1      "同意" checkbox label
'   I2      "ナレッジ改善に質問本文を提供する" (TRUE/FALSE cell)
'   I4      Status label ("Ready" / "問い合わせ中…" / etc.)
' ============================================================================

Public Const SHEET_NAME As String = "ChatUI"
Public Const CELL_QUESTION As String = "B3"
Public Const CELL_CONSENT As String = "I2"
Public Const CELL_STATUS As String = "I4"
Public Const ROW_HISTORY_START As Long = 7

Private Const BTN_SEND As String = "btnSendQuestion"

Public Sub EnsureLayout()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_NAME)
    Application.ScreenUpdating = False

    ws.Cells.Clear
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    On Error GoTo 0

    ' Column widths
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B:E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G:H").ColumnWidth = 3
    ws.Columns("I").ColumnWidth = 36

    ' Title banner
    With ws.Range("A1:F1")
        .Merge
        .value = "社内ナレッジQAボット (デモ版)"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = RGB(60, 90, 160)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .RowHeight = 28
    End With

    ' Question label + input
    ws.Range("A3").value = "質問:"
    ws.Range("A3").Font.Bold = True
    With ws.Range("B3:E4")
        .Merge
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.Color = RGB(255, 252, 230)
        .Borders.LineStyle = xlContinuous
    End With
    ws.Rows("3:4").RowHeight = 36

    ' Send button
    Dim btn As Button
    Dim btnRange As Range
    Set btnRange = ws.Range("F3:F4")
    Set btn = ws.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.Width, btnRange.Height)
    btn.OnAction = "modChatUI.SendButton_Click"
    btn.Caption = "▶ 送信"
    btn.Name = BTN_SEND
    btn.Font.Size = 11
    btn.Font.Bold = True

    ' History header
    ws.Range("A6").value = "会話履歴"
    ws.Range("A6").Font.Bold = True
    ws.Range("A6:F6").Borders(xlEdgeBottom).LineStyle = xlContinuous

    ' Right column - consent + status
    ws.Range("I1").value = "オプション"
    ws.Range("I1").Font.Bold = True
    ws.Range("H2").value = "本文提供:"
    ws.Range(CELL_CONSENT).value = False
    With ws.Range(CELL_CONSENT)
        .Interior.Color = RGB(245, 245, 245)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("I3").value = "(TRUE にすると質問本文も管理者ログに記録)"
    ws.Range("I3").Font.Size = 8
    ws.Range("H4").value = "状態:"
    ws.Range(CELL_STATUS).value = "Ready"
    ws.Range(CELL_STATUS).Interior.Color = RGB(220, 240, 220)

    ' Disclaimer
    AppendHistory ws, "システム", LoadDisclaimer()
    AppendHistory ws, "", "（質問を入力して『送信』を押してください）"

    ws.Activate
    ws.Range(CELL_QUESTION).Select
    Application.ScreenUpdating = True
End Sub

' ----------------------------------------------------------------------------
' Send button entry point (invoked via OnAction)
' ----------------------------------------------------------------------------
Public Sub SendButton_Click()
    If Not modBoot.gReady Then
        MsgBox "ナレッジindexがまだ読み込まれていません。", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Dim q As String: q = Trim$(CStr(ws.Range(CELL_QUESTION).value))
    If LenB(q) = 0 Then
        MsgBox "質問を入力してください。", vbInformation
        Exit Sub
    End If

    ' Rate limit
    Dim retry As Long
    If Not modRateLimiter.CanSend(retry) Then
        MsgBox "本日の質問上限に達しました。約 " & retry & " 秒お待ちください。", vbExclamation
        Exit Sub
    End If

    ' PII guard
    Dim piiWarn As Boolean
    piiWarn = modPiiGuard.DetectPii(q)
    If piiWarn Then
        If MsgBox("個人情報が含まれている可能性があります。送信しますか？", _
                  vbYesNo + vbExclamation, "確認") <> vbYes Then
            Exit Sub
        End If
    End If

    SetStatus ws, "問い合わせ中…", RGB(255, 235, 180)
    DisableSend ws, True
    DoEvents

    Dim res As AnswerResult
    res = modRagEngine.Answer(modBoot.gIndex, q)

    If res.OK Then
        AppendHistory ws, "あなた", q
        AppendHistory ws, "AI", res.Answer
        AppendHistory ws, "出典", FormatCitations(res.Citations)
        modRateLimiter.RecordSend
        Dim consent As Boolean
        consent = CBool(ws.Range(CELL_CONSENT).value)
        modUsageLogger.LogUsage _
            res.PromptTokens, res.CompletionTokens, _
            res.EmbedLatencyMs + res.ChatLatencyMs, _
            piiWarn, consent, q
        SetStatus ws, "Ready (" & res.PromptTokens & "/" & res.CompletionTokens _
            & " tok, " & (res.EmbedLatencyMs + res.ChatLatencyMs) & " ms)", _
            RGB(220, 240, 220)
        ws.Range(CELL_QUESTION).value = ""
    Else
        AppendHistory ws, "エラー", res.ErrorMessage
        SetStatus ws, "Error", RGB(250, 200, 200)
    End If

    DisableSend ws, False
    ws.Range(CELL_QUESTION).Select
End Sub

' ----------------------------------------------------------------------------
' Append a labelled block to the conversation history
' ----------------------------------------------------------------------------
Public Sub AppendHistory(ByVal ws As Worksheet, ByVal label As String, ByVal text As String)
    Dim row As Long
    row = NextHistoryRow(ws)

    Dim labelCell As Range, bodyCell As Range
    Set labelCell = ws.Cells(row, 1)
    Set bodyCell = ws.Range(ws.Cells(row, 2), ws.Cells(row, 6))

    labelCell.value = label
    labelCell.Font.Bold = True
    labelCell.VerticalAlignment = xlTop

    bodyCell.Merge
    bodyCell.value = text
    bodyCell.WrapText = True
    bodyCell.VerticalAlignment = xlTop

    Select Case label
        Case "あなた":  labelCell.Font.Color = RGB(20, 90, 50): bodyCell.Interior.Color = RGB(232, 246, 234)
        Case "AI":      labelCell.Font.Color = RGB(60, 90, 160): bodyCell.Interior.Color = RGB(232, 238, 252)
        Case "出典":    labelCell.Font.Size = 9:                bodyCell.Font.Size = 9: bodyCell.Interior.Color = RGB(248, 248, 240)
        Case "エラー":  labelCell.Font.Color = RGB(170, 30, 30): bodyCell.Interior.Color = RGB(252, 232, 232)
        Case "システム": labelCell.Font.Color = RGB(120, 120, 120): bodyCell.Font.Size = 9: bodyCell.Font.Italic = True
    End Select

    ws.Rows(row).RowHeight = EstimateRowHeight(text)
    ws.Cells(row, 1).EntireRow.Select
    ActiveWindow.ScrollRow = row
    ws.Range(CELL_QUESTION).Select
End Sub

Private Function NextHistoryRow(ByVal ws As Worksheet) As Long
    Dim r As Long: r = ROW_HISTORY_START
    Do While LenB(CStr(ws.Cells(r, 1).value)) > 0 Or LenB(CStr(ws.Cells(r, 2).value)) > 0
        r = r + 1
        If r > 5000 Then Exit Do ' safety cap
    Loop
    NextHistoryRow = r
End Function

Private Function EstimateRowHeight(ByVal text As String) As Double
    Dim lines As Long
    lines = 1 + (Len(text) \ 80)
    Dim lf As Long: lf = Len(text) - Len(Replace(text, vbLf, ""))
    lines = lines + lf
    If lines > 30 Then lines = 30
    EstimateRowHeight = 14 + lines * 14
End Function

Private Sub SetStatus(ByVal ws As Worksheet, ByVal text As String, ByVal color As Long)
    ws.Range(CELL_STATUS).value = text
    ws.Range(CELL_STATUS).Interior.color = color
End Sub

Private Sub DisableSend(ByVal ws As Worksheet, ByVal disable As Boolean)
    On Error Resume Next
    Dim btn As Object
    Set btn = ws.Buttons(BTN_SEND)
    btn.Enabled = Not disable
    btn.Caption = IIf(disable, "・・・", "▶ 送信")
    On Error GoTo 0
End Sub

Private Function FormatCitations(ByRef c() As Citation) As String
    If LBound(c) > UBound(c) Then
        FormatCitations = "(参照ナレッジなし)"
        Exit Function
    End If
    Dim sb As String, i As Long
    For i = LBound(c) To UBound(c)
        If i > LBound(c) Then sb = sb & vbLf
        sb = sb & "[#" & (i + 1) & "] " & c(i).Source
        If c(i).page > 0 Then sb = sb & " p." & c(i).page
        sb = sb & "  (score " & Format$(c(i).Score, "0.000") & ")"
    Next i
    FormatCitations = sb
End Function

Private Function LoadDisclaimer() As String
    Dim path As String
    path = ThisWorkbook.Path & "\" & modConfig.GetString("ui", "disclaimer_path", "disclaimer.txt")
    If Len(Dir$(path)) = 0 Then
        path = Environ$("LOCALAPPDATA") & "\InternalNotebookLM\disclaimer.txt"
    End If
    If Len(Dir$(path)) = 0 Then Exit Function
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    On Error Resume Next
    st.LoadFromFile path
    LoadDisclaimer = st.ReadText
    On Error GoTo 0
    st.Close
End Function

Private Function GetOrCreateSheet(ByVal name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets( _
            ThisWorkbook.Worksheets.count))
        ws.name = name
    End If
    Set GetOrCreateSheet = ws
End Function
