Attribute VB_Name = "modChatUI"
Option Explicit

' ============================================================================
' modChatUI - v2 in-worksheet UI
' ----------------------------------------------------------------------------
' Cell-based UI (no ActiveX needed). Generated at Workbook_Open.
' Layout:
'   A1:H1    Title banner
'   A2:H2    Status bar (dept, user, model)
'   A4       "質問:"
'   B4:H7    Question input cell (merged)
'   B8:H8    Send + clear buttons
'   A10      "回答"
'   B10:H40  Answer area (merged, scrollable via row height)
'   B41:C41  good button / B41:D41 bad button
'   A42      "出典"
'   B42:H50  Citations
'   A52      Diagnostics
' ============================================================================

Public Const SHEET_NAME As String = "main"
Public Const CELL_QUESTION As String = "B4"
Public Const CELL_ANSWER As String = "B10"
Public Const CELL_CITATIONS As String = "B42"
Public Const CELL_STATUS As String = "B2"
Public Const CELL_DIAG As String = "B52"

Private Const BTN_SEND As String = "btnSend"
Private Const BTN_CLEAR As String = "btnClear"
Private Const BTN_GOOD As String = "btnGood"
Private Const BTN_BAD As String = "btnBad"
Private Const BTN_TXT As String = "btnSaveTxt"
Private Const BTN_HTML As String = "btnSaveHtml"

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

    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B:G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 14

    ' Title banner
    With ws.Range("A1:H1")
        .Merge
        .value = "社内ナレッジ QA ボット (v2 - 社内AIリボン)"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.color = RGB(60, 90, 160)
        .Font.color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .RowHeight = 28
    End With

    ' Status bar
    With ws.Range("A2:H2")
        .Merge
        .value = "起動中..."
        .HorizontalAlignment = xlCenter
        .Interior.color = RGB(255, 245, 200)
        .Font.Size = 10
    End With

    ' Question section
    ws.Range("A4").value = "質問:"
    ws.Range("A4").Font.Bold = True
    With ws.Range("B4:H7")
        .Merge
        .value = "ここに質問を書いてください"
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.color = RGB(255, 252, 230)
        .Borders.LineStyle = xlContinuous
    End With
    ws.Rows("4:7").RowHeight = 22

    ' Send + Clear buttons row
    AddButton ws, ws.Range("B8:C8"), BTN_SEND, "> 送信", "modChatUI.OnSendClick"
    AddButton ws, ws.Range("D8:E8"), BTN_CLEAR, "クリア", "modChatUI.OnClearClick"
    ws.Range("F8:H8").Merge
    ws.Range("F8").value = "(回答に60〜90秒かかります)"
    ws.Range("F8").HorizontalAlignment = xlRight
    ws.Range("F8").Font.Size = 9
    ws.Range("F8").Font.color = RGB(120, 120, 120)
    ws.Rows("8").RowHeight = 26

    ' Answer header
    ws.Range("A10").value = "回答:"
    ws.Range("A10").Font.Bold = True
    With ws.Range("B10:H40")
        .Merge
        .value = "（質問を入力して『> 送信』を押してください）"
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.color = RGB(232, 238, 252)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 10
    End With
    ws.Rows("10:40").RowHeight = 14

    ' Feedback buttons + save buttons
    AddButton ws, ws.Range("B41:C41"), BTN_GOOD, "○ 良かった", "modFeedback.Submit_Feedback_Good"
    AddButton ws, ws.Range("D41:E41"), BTN_BAD, "× 修正", "modFeedback.Submit_Feedback_Bad"
    AddButton ws, ws.Range("F41:G41"), BTN_TXT, "テキスト出力", "modChatUI.OnSaveTxtClick"
    AddButton ws, ws.Range("H41:H41"), BTN_HTML, "HTML出力", "modChatUI.OnSaveHtmlClick"
    ws.Rows("41").RowHeight = 26

    ' Diagnostics row (always available - helps locate problems fast)
    AddButton ws, ws.Range("B53:C53"), "btnDiag", "自己診断", "modDiag.RunDiagnostics"
    AddButton ws, ws.Range("D53:E53"), "btnRibbon", "リボン接続テスト", "modDiag.TestRibbon"
    AddButton ws, ws.Range("F53:G53"), "btnReboot", "再起動", "modBoot.Boot"
    ws.Rows("53").RowHeight = 24

    ' Citations
    ws.Range("A42").value = "出典:"
    ws.Range("A42").Font.Bold = True
    With ws.Range("B42:H50")
        .Merge
        .value = ""
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.color = RGB(248, 248, 240)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 9
    End With

    ' Diagnostics
    ws.Range("A52").value = "状態:"
    ws.Range("A52").Font.Bold = True
    ws.Range("B52:H52").Merge
    ws.Range("B52").Font.Size = 9
    ws.Range("B52").Font.color = RGB(120, 120, 120)

    ' Disclaimer
    With ws.Range("A54:H56")
        .Merge
        .value = "■ 重要：この回答はAI生成です。最終判断はアンダーライターへ確認してください。" & vbLf & _
                 "■ 質問に契約番号・氏名・電話番号などの個人情報を含めないでください。"
        .WrapText = True
        .Font.Size = 9
        .Font.color = RGB(150, 60, 60)
    End With

    ' Refresh status bar with dept/user info
    UpdateStatus ws, "Ready", RGB(220, 240, 220)

    ws.Activate
    ws.Range(CELL_QUESTION).Select
    Application.ScreenUpdating = True
End Sub

' ----------------------------------------------------------------------------
Public Sub OnSendClick()
    On Error GoTo Trap
    If Not modBoot.gReady Then
        If MsgBox("ボットが起動していません。今すぐ起動しますか？", vbYesNo + vbExclamation) = vbYes Then
            modBoot.Boot
        End If
        If Not modBoot.gReady Then Exit Sub
    End If
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Dim q As String: q = Trim$(CStr(ws.Range(CELL_QUESTION).value))
    If LenB(q) = 0 Or q = "ここに質問を書いてください" Then
        MsgBox "質問を入力してください。", vbInformation: Exit Sub
    End If

    ' PII guard
    Dim piiWarn As Boolean: piiWarn = modPii.DetectPii(q)
    If piiWarn Then
        If MsgBox("個人情報が含まれている可能性があります。送信しますか？", _
                  vbYesNo + vbExclamation, "確認") <> vbYes Then Exit Sub
    End If

    UpdateStatus ws, "問い合わせ中...しばらくお待ちください (60〜90秒)", RGB(255, 235, 180)
    SetButtonsEnabled ws, False
    ws.Range(CELL_ANSWER).value = "(生成中...)"
    ws.Range(CELL_CITATIONS).value = ""
    ws.Range(CELL_DIAG).value = "ルーター起動中..."
    DoEvents

    Dim res As modPipeline.PipelineResult
    res = modPipeline.RunQuery(q)

    If res.OK Then
        ws.Range(CELL_ANSWER).value = res.Answer
        ws.Range(CELL_CITATIONS).value = res.Citations
        ws.Range(CELL_DIAG).value = "Ready (合計 " & res.TotalMs & " ms : router " & res.RouterMs & _
                                    " + draft " & res.DraftMs & " + verify " & res.VerifyMs & ")"
        ' Save for feedback buttons
        modBoot.gLastQuestion = q
        modBoot.gLastAnswer = res.Answer
        modBoot.gLastSelectedIds = res.SelectedIds
        UpdateStatus ws, "Ready (○/× で評価できます)", RGB(220, 240, 220)
        ' Usage log
        modUsageLogger.LogQuery q, res
    Else
        ws.Range(CELL_ANSWER).value = "■ 処理を完了できませんでした" & vbLf & vbLf & _
            "失敗ステップ・原因:" & vbLf & res.ErrorMsg & vbLf & vbLf & _
            "対処: 下の『自己診断』『リボン接続テスト』ボタンで原因を特定できます。"
        ws.Range(CELL_DIAG).value = "Error: " & Left$(Replace(res.ErrorMsg, vbLf, " "), 200)
        UpdateStatus ws, "Error (診断ボタンで詳細確認)", RGB(250, 200, 200)
    End If

    SetButtonsEnabled ws, True
    AutoSizeAnswer ws
    ws.Range(CELL_QUESTION).Select
    Exit Sub

Trap:
    On Error Resume Next
    SetButtonsEnabled ThisWorkbook.Worksheets(SHEET_NAME), True
    On Error GoTo 0
    modDiag.ReportError "modChatUI.OnSendClick", Err.Number, Err.Description, _
        "質問送信処理で予期せぬエラー。『自己診断』で各サブシステムを確認してください。"
End Sub

Public Sub OnClearClick()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    ws.Range(CELL_QUESTION).value = "ここに質問を書いてください"
    ws.Range(CELL_ANSWER).value = "（質問を入力して『> 送信』を押してください）"
    ws.Range(CELL_CITATIONS).value = ""
    ws.Range(CELL_DIAG).value = ""
    modBoot.gLastQuestion = "": modBoot.gLastAnswer = "": modBoot.gLastSelectedIds = ""
    ws.Range(CELL_QUESTION).Select
End Sub

Public Sub OnSaveTxtClick()
    If LenB(modBoot.gLastAnswer) = 0 Then MsgBox "回答がありません。": Exit Sub
    Dim default_name As String: default_name = "ai_answer_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
    Dim savePath As Variant
    savePath = Application.GetSaveAsFilename(InitialFileName:=default_name, _
        FileFilter:="テキスト (*.txt), *.txt", Title:="テキスト保存")
    If savePath = False Then Exit Sub
    SaveUtf8 CStr(savePath), modBoot.gLastQuestion & vbLf & vbLf & "----" & vbLf & vbLf & modBoot.gLastAnswer
    MsgBox "保存しました: " & savePath, vbInformation
End Sub

Public Sub OnSaveHtmlClick()
    If LenB(modBoot.gLastAnswer) = 0 Then MsgBox "回答がありません。": Exit Sub
    Dim default_name As String: default_name = "ai_answer_" & Format$(Now, "yyyymmdd_hhnnss") & ".html"
    Dim savePath As Variant
    savePath = Application.GetSaveAsFilename(InitialFileName:=default_name, _
        FileFilter:="HTML (*.html), *.html", Title:="HTML保存")
    If savePath = False Then Exit Sub
    Dim html As String
    html = "<!doctype html><html><head><meta charset='utf-8'><title>AI回答</title>" & _
           "<style>body{font-family:Meiryo,sans-serif;max-width:780px;margin:24px auto;padding:0 12px;color:#222}" & _
           "h1{font-size:18px;border-bottom:2px solid #3c5aa0;padding-bottom:6px}" & _
           "pre{background:#f5f5f0;padding:12px;white-space:pre-wrap;font-family:Meiryo;font-size:13px}</style></head><body>"
    html = html & "<h1>質問</h1><pre>" & HtmlEscape(modBoot.gLastQuestion) & "</pre>"
    html = html & "<h1>AI回答</h1><pre>" & HtmlEscape(modBoot.gLastAnswer) & "</pre>"
    html = html & "</body></html>"
    SaveUtf8 CStr(savePath), html
    Shell "rundll32.exe url.dll,FileProtocolHandler " & Chr(34) & CStr(savePath) & Chr(34), vbNormalFocus
End Sub

' ----------------------------------------------------------------------------
Private Sub AddButton(ByVal ws As Worksheet, ByVal rng As Range, _
                      ByVal name As String, ByVal caption As String, ByVal action As String)
    Dim btn As Button
    Set btn = ws.Buttons.Add(rng.Left, rng.Top, rng.Width, rng.Height)
    btn.OnAction = action
    btn.Caption = caption
    btn.name = name
    btn.Font.Size = 10
End Sub

Public Sub UpdateStatus(ByVal ws As Worksheet, ByVal text As String, ByVal color As Long)
    Dim msg As String
    msg = "部署: " & modUserProfile.CurrentDeptName() & _
          "  |  役割: " & modUserProfile.CurrentRole() & _
          "  |  状態: " & text
    ws.Range(CELL_STATUS).value = msg
    ws.Range(CELL_STATUS).Interior.color = color
End Sub

Private Sub SetButtonsEnabled(ByVal ws As Worksheet, ByVal isEnabled As Boolean)
    On Error Resume Next
    ws.Buttons(BTN_SEND).Enabled = isEnabled
    ws.Buttons(BTN_CLEAR).Enabled = isEnabled
    ws.Buttons(BTN_GOOD).Enabled = isEnabled
    ws.Buttons(BTN_BAD).Enabled = isEnabled
    ws.Buttons(BTN_SEND).caption = IIf(isEnabled, "> 送信", "...生成中...")
    On Error GoTo 0
End Sub

Private Sub AutoSizeAnswer(ByVal ws As Worksheet)
    ' Approximate row height based on answer length
    Dim ans As String: ans = CStr(ws.Range(CELL_ANSWER).value)
    Dim lines As Long
    lines = 1 + (Len(ans) \ 70) + (Len(ans) - Len(Replace(ans, vbLf, "")))
    If lines > 200 Then lines = 200
    ' Distribute across rows 10-40
    Dim each_h As Double: each_h = Application.Max(14, (lines * 14) / 31)
    Dim r As Long
    For r = 10 To 40
        ws.Rows(r).RowHeight = each_h
    Next r
End Sub

' Save text as UTF-8. Primary path uses ADODB.Stream (Windows). If that object
' is unavailable (Mac Excel, or a hardened PC where msado is blocked), fall back
' to a native VBA byte write that encodes UTF-8 by hand. Never raises to caller.
Private Sub SaveUtf8(ByVal path As String, ByVal content As String)
    On Error GoTo Fallback
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open
    stm.WriteText content
    stm.Position = 0
    stm.Type = 1
    stm.Position = 3                      ' skip UTF-8 BOM ADODB prepends
    Dim bin() As Byte: bin = stm.Read
    stm.Close
    Dim stm2 As Object: Set stm2 = CreateObject("ADODB.Stream")
    stm2.Type = 1: stm2.Open
    stm2.Write bin
    stm2.SaveToFile path, 2
    stm2.Close
    Exit Sub
Fallback:
    SaveUtf8Native path, content
End Sub

' Pure-VBA UTF-8 writer (no external objects). Encodes the BMP correctly,
' which covers Japanese, kanji, kana and ASCII.
Private Sub SaveUtf8Native(ByVal path As String, ByVal content As String)
    On Error Resume Next
    Dim fnum As Integer: fnum = FreeFile
    Open path For Binary Access Write As #fnum
    Dim i As Long, cp As Long
    For i = 1 To Len(content)
        cp = AscW(Mid$(content, i, 1)) And &HFFFF&
        If cp < &H80 Then
            Put #fnum, , CByte(cp)
        ElseIf cp < &H800 Then
            Put #fnum, , CByte(&HC0 Or (cp \ &H40))
            Put #fnum, , CByte(&H80 Or (cp And &H3F))
        Else
            Put #fnum, , CByte(&HE0 Or (cp \ &H1000))
            Put #fnum, , CByte(&H80 Or ((cp \ &H40) And &H3F))
            Put #fnum, , CByte(&H80 Or (cp And &H3F))
        End If
    Next i
    Close #fnum
End Sub

Private Function HtmlEscape(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    HtmlEscape = s
End Function

Private Function GetOrCreateSheet(ByVal name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.name = name
    End If
    Set GetOrCreateSheet = ws
End Function
