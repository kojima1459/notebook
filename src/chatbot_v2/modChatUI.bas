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
Private Const BTN_FOLLOWUP As String = "btnFollowup"
Private Const BTN_INQUIRY As String = "btnInquiry"
Private Const BTN_MODE As String = "btnMode"
Public Const CELL_MEMO As String = "B58"

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

    ' Send + Follow-up + Clear buttons row
    AddButton ws, ws.Range("B8:C8"), BTN_SEND, "> 送信", "modChatUI.OnSendClick"
    AddButton ws, ws.Range("D8:E8"), BTN_FOLLOWUP, "続けて質問", "modChatUI.OnFollowupClick"
    AddButton ws, ws.Range("F8:G8"), BTN_CLEAR, "クリア", "modChatUI.OnClearClick"
    ws.Range("H8").value = "(回答に60〜90秒)"
    ws.Range("H8").HorizontalAlignment = xlRight
    ws.Range("H8").Font.Size = 9
    ws.Range("H8").Font.color = RGB(120, 120, 120)
    ws.Rows("8").RowHeight = 26

    ' Intent-mode toggle (smart <-> always). Lets each user choose how strongly
    ' the AI confirms premises before answering.
    AddButton ws, ws.Range("B9:D9"), BTN_MODE, ModeButtonCaption(), "modChatUI.OnToggleMode"
    With ws.Range("E9:H9")
        .Merge
        .value = "← 質問の前提をAIに確認させる強さ（迷う/新人なら『毎回前提確認』推奨）"
        .Font.Size = 9
        .Font.color = RGB(120, 120, 120)
        .VerticalAlignment = xlCenter
    End With
    ws.Rows("9").RowHeight = 22

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

    ' Post-answer action row: evaluate (○/×) then escalate (教えてBOX) if still unsure.
    AddButton ws, ws.Range("B41:C41"), BTN_GOOD, "○ 良かった", "modFeedback.Submit_Feedback_Good"
    AddButton ws, ws.Range("D41:E41"), BTN_BAD, "× 修正", "modFeedback.Submit_Feedback_Bad"
    AddButton ws, ws.Range("F41:H41"), BTN_INQUIRY, "📮 教えてBOX（本社へ照会）", "modChatUI.OnInquiryClick"
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

    ' Save buttons (less prominent, below citations)
    AddButton ws, ws.Range("B51:C51"), BTN_TXT, "テキスト出力", "modChatUI.OnSaveTxtClick"
    AddButton ws, ws.Range("D51:E51"), BTN_HTML, "HTML出力", "modChatUI.OnSaveHtmlClick"
    ws.Range("F51:H51").value = "← 回答を保存"
    ws.Range("F51").Font.Size = 9
    ws.Range("F51").Font.color = RGB(120, 120, 120)
    ws.Rows("51").RowHeight = 24

    ' Diagnostics
    ws.Range("A52").value = "状態:"
    ws.Range("A52").Font.Bold = True
    ws.Range("B52:H52").Merge
    ws.Range("B52").Font.Size = 9
    ws.Range("B52").Font.color = RGB(120, 120, 120)

    ' Self-understanding flow + guard text (always visible under the answer)
    With ws.Range("A54:H57")
        .Merge
        .value = _
            "【回答を受け取ったら（この順で自己解決 → 学習が深まります）】" & vbLf & _
            "  ① まず出典元（約款・引受ガイドライン）を自分の目で確認する" & vbLf & _
            "  ② 理解できたか確認する。曖昧なら『続けて質問』で深掘り（何度でもOK）" & vbLf & _
            "  ③ 「○良かった／×修正」で評価する（ここまでで1セット。次の人の役に立ちます）" & vbLf & _
            "  ④ それでも判断に迷う時だけ『📮 教えてBOX』で本社へ照会" & vbLf & _
            "■ 本回答はAIが生成しており、必ずしも正しいとは限りません。最終判断はアンダーライター／原本確認を。" & vbLf & _
            "■ 質問に契約番号・氏名・電話番号などの個人情報を含めないでください。"
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Size = 9
        .Font.color = RGB(120, 70, 70)
        .Interior.color = RGB(255, 250, 240)
        .Borders.LineStyle = xlContinuous
    End With

    ' Next-question memo (jot the next question while you read / wait)
    ws.Range("A58").value = "💭メモ:"
    ws.Range("A58").Font.Size = 9
    ws.Range("A58").Font.Bold = True
    With ws.Range("B58:H59")
        .Merge
        .value = "（思いついた次の質問をここに書きためておけます。送信欄にコピーして使ってください）"
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.color = RGB(252, 252, 245)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 9
        .Font.color = RGB(150, 150, 150)
    End With

    ' Footer: model recommendation + copyright
    With ws.Range("A60:H60")
        .Merge
        .value = "⚡ 必ず最新モデル「" & modConfig.GetString("recommended_model", "GPT-5.5") & _
                 "」(2026年6月時点)に切り替えてご利用ください。" & _
                 "　　© " & modConfig.GetString("creator_group", "ニューリスクG") & " " & _
                 modConfig.GetString("creator_name", "小島正豪")
        .Font.Size = 9
        .Font.color = RGB(110, 110, 130)
        .HorizontalAlignment = xlLeft
    End With

    ' Refresh status bar with dept/user info
    UpdateStatus ws, "Ready", RGB(220, 240, 220)

    ' Activate/Select can fail during Workbook_Open before the window is
    ' fully ready. Never let cosmetics propagate an error.
    On Error Resume Next
    ws.Activate
    ws.Range(CELL_QUESTION).Select
    On Error GoTo 0
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

    ' A direct "送信" (not via 続けて質問) starts a fresh thread -> clear history.
    If Not modBoot.gFollowupMode Then modBoot.ResetHistory

    SetButtonsEnabled ws, False

    ' ---- Step 0: Intent layer (意図解釈 + 前提確認) -----------------------
    Dim ir As modIntent.IntentResult
    Dim extra As String
    If modConfig.GetBool("intent_layer_enabled", True) Then
        UpdateStatus ws, "照会の意図を解釈中... (1/4)", RGB(255, 235, 180)
        SetCellSafe ws.Range(CELL_ANSWER), "(照会の意図を解釈しています...)"
        DoEvents
        Dim mode As String: mode = modConfig.GetString("intent_mode", "smart")
        ir = modIntent.AnalyzeIntent(q, mode, modBoot.HistoryBlock())
        If ir.OK And ir.NeedsClarification And LenB(ir.ClarifyingText) > 0 Then
            SetCellSafe ws.Range(CELL_ANSWER), _
                "【確認させてください】より正確にお答えするため、以下を教えてください:" & vbLf & vbLf & _
                ir.ClarifyingText & vbLf & vbLf & _
                "→ ポップアップに分かる範囲で補足してください（空欄のままでも、一般的な前提で回答します）"
            AutoSizeAnswer ws
            DoEvents
            extra = InputBox( _
                ir.ClarifyingText & vbCrLf & vbCrLf & _
                "↑ について分かる範囲で補足してください。" & vbCrLf & _
                "（空欄でOK。その場合は一般的な前提で回答します）", _
                "前提の確認 - より正確な回答のために")
        End If
    End If

    UpdateStatus ws, "問い合わせ中...しばらくお待ちください (60〜90秒)", RGB(255, 235, 180)
    SetCellSafe ws.Range(CELL_ANSWER), "(回答を生成中... 検索→ドラフト→検証。Excelが応答なしに見えても正常です)"
    SetCellSafe ws.Range(CELL_CITATIONS), ""
    SetCellSafe ws.Range(CELL_DIAG), "パイプライン実行中..."
    DoEvents

    Dim res As modPipeline.PipelineResult
    res = modPipeline.RunQuery(q, ir, extra)

    If res.OK Then
        ' Compose the display: confidence banner + answer + follow-up suggestions.
        Dim disp As String
        disp = ConfidenceBanner(res.Confidence) & vbLf & vbLf & res.Answer
        If LenB(res.Followups) > 0 Then
            disp = disp & vbLf & vbLf & "🔎 深掘り候補（『続けて質問』でそのまま聞けます）:"
            Dim fl() As String: fl = Split(res.Followups, vbLf)
            Dim fi As Long
            For fi = LBound(fl) To UBound(fl)
                If LenB(Trim$(fl(fi))) > 0 Then disp = disp & vbLf & "  ・" & fl(fi)
            Next fi
        End If
        SetCellSafe ws.Range(CELL_ANSWER), disp
        SetCellSafe ws.Range(CELL_CITATIONS), res.Citations
        SetCellSafe ws.Range(CELL_DIAG), "Ready (合計 " & res.TotalMs & " ms : router " & res.RouterMs & _
                                    " + draft " & res.DraftMs & " + verify " & res.VerifyMs & " / 自信度 " & res.Confidence & ")"
        ' Save state for feedback / 教えてBOX / follow-up
        modBoot.gLastQuestion = q
        modBoot.gLastAnswer = res.Answer
        modBoot.gLastSelectedIds = res.SelectedIds
        modBoot.gLastIntent = res.Intent
        modBoot.gLastAssumptions = res.Assumptions
        modBoot.gLastCitations = res.Citations
        modBoot.gLastConfidence = res.Confidence
        modBoot.gLastFollowups = res.Followups
        ' Nudge 教えてBOX based on confidence (caption changes; Form buttons can't recolor)
        SetInquiryProminence ws, res.Confidence
        ' Record this exchange so a later 続けて質問 has the running context
        modBoot.AppendHistory q, res.Answer
        UpdateStatus ws, "完了 (自信度 " & res.Confidence & ")。出典を確認→理解できたか確認→○/×評価を。迷えば教えてBOX", RGB(220, 240, 220)
        ' Usage log
        modUsageLogger.LogQuery q, res
    Else
        SetCellSafe ws.Range(CELL_ANSWER), "■ 処理を完了できませんでした" & vbLf & vbLf & _
            "失敗ステップ・原因:" & vbLf & res.ErrorMsg & vbLf & vbLf & _
            "対処: 下の『自己診断』『リボン接続テスト』ボタンで原因を特定できます。"
        SetCellSafe ws.Range(CELL_DIAG), "Error: " & Left$(Replace(res.ErrorMsg, vbLf, " "), 200)
        UpdateStatus ws, "Error (診断ボタンで詳細確認)", RGB(250, 200, 200)
    End If

    SetButtonsEnabled ws, True
    AutoSizeAnswer ws
    AutoSizeCitations ws
    On Error Resume Next
    ws.Range(CELL_QUESTION).Select
    On Error GoTo 0
    Exit Sub

Trap:
    ' Capture original err details BEFORE any cleanup can clobber them.
    Dim origErrN As Long: origErrN = Err.Number
    Dim origErrD As String: origErrD = Err.Description
    On Error Resume Next
    SetButtonsEnabled ThisWorkbook.Worksheets(SHEET_NAME), True
    On Error GoTo 0
    modDiag.ReportError "modChatUI.OnSendClick", origErrN, origErrD, _
        "質問送信処理で予期せぬエラー。『自己診断』で各サブシステムを確認してください。"
End Sub

Public Sub OnFollowupClick()
    On Error GoTo Trap
    If LenB(modBoot.gLastQuestion) = 0 Or LenB(modBoot.gLastAnswer) = 0 Then
        MsgBox "まず最初の質問を送信して、回答を受け取ってから『続けて質問』を使ってください。", _
               vbInformation, "続けて質問"
        Exit Sub
    End If

    Dim hint As String
    hint = "前回までの会話を踏まえて、追加の質問・深掘りを入力してください。" & vbCrLf & _
           "（直近 " & modBoot.HISTORY_MAX_TURNS & " 往復ぶんの会話を記憶しています。何度でも続けられます）" & vbCrLf & vbCrLf & _
           "例:" & vbCrLf & _
           "  ・前売券の払い戻し手数料は対象になる？" & vbCrLf & _
           "  ・グッズ代以外で対象外になりやすい収益は？" & vbCrLf & _
           "  ・中止と部分中止で扱いは変わる？"
    Dim followup As String
    followup = InputBox(hint, "続けて質問（深掘り）", "")
    If LenB(followup) = 0 Then Exit Sub

    ' Follow-up mode: OnSendClick will keep (not reset) the running history.
    modBoot.gFollowupMode = True

    ' Reuse the main send flow: set the question into the cell and call OnSendClick
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    ws.Range(CELL_QUESTION).value = followup
    OnSendClick

    ' OnSendClick sets gLastQuestion/gLastAnswer; reset the flag so the next
    ' top-level "送信" doesn't accidentally carry follow-up context.
    modBoot.gFollowupMode = False
    Exit Sub

Trap:
    modBoot.gFollowupMode = False
    modDiag.ReportError "modChatUI.OnFollowupClick", Err.Number, Err.Description, _
        "続けて質問の処理でエラーが発生しました。"
End Sub

Public Sub OnClearClick()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    ws.Range(CELL_QUESTION).value = "ここに質問を書いてください"
    ws.Range(CELL_ANSWER).value = "（質問を入力して『> 送信』を押してください）"
    ws.Range(CELL_CITATIONS).value = ""
    ws.Range(CELL_DIAG).value = ""
    ' Reset follow-up state so a new top-level question starts clean
    modBoot.gFollowupMode = False
    modBoot.ResetHistory
    modBoot.gLastQuestion = "": modBoot.gLastAnswer = "": modBoot.gLastSelectedIds = ""
    modBoot.gLastIntent = "": modBoot.gLastAssumptions = "": modBoot.gLastCitations = ""
    modBoot.gLastConfidence = "": modBoot.gLastFollowups = ""
    SetInquiryProminence ws, ""
    ws.Range(CELL_QUESTION).Select
End Sub

' ----------------------------------------------------------------------------
' Intent-mode toggle: smart <-> always. Persisted in config so each distributed
' file remembers the user's choice.
' ----------------------------------------------------------------------------
Public Function ModeButtonCaption() As String
    If LCase$(modConfig.GetString("intent_mode", "smart")) = "always" Then
        ModeButtonCaption = "確認モード: 毎回前提確認"
    Else
        ModeButtonCaption = "確認モード: 賢く出し分け"
    End If
End Function

Public Sub OnToggleMode()
    Dim m As String: m = LCase$(modConfig.GetString("intent_mode", "smart"))
    If m = "always" Then m = "smart" Else m = "always"
    modConfig.SetValue "intent_mode", m
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_NAME).Buttons(BTN_MODE).caption = ModeButtonCaption()
    On Error GoTo 0
    Dim desc As String
    If m = "always" Then
        desc = "どんな質問でも、まず前提を確認してから回答します（誤解を最小化）。"
    Else
        desc = "質問が明確なら即回答、曖昧な時だけ前提を聞き返します。"
    End If
    MsgBox "確認モードを切り替えました。" & vbCrLf & vbCrLf & ModeButtonCaption() & vbCrLf & desc, _
           vbInformation, "確認モード"
End Sub

Public Sub OnInquiryClick()
    modInquiryBox.OpenInquiryBox
End Sub

' Confidence banner prepended to the answer.
Private Function ConfidenceBanner(ByVal c As String) As String
    Select Case c
        Case "高"
            ConfidenceBanner = "🟢 AIの自信度: 高 — 約款・ガイドラインに明確な根拠があります（出典は必ずご確認ください）"
        Case "低"
            ConfidenceBanner = "🔴 AIの自信度: 低 — 該当規定が乏しく不確実です。『📮 教えてBOX』で本社への照会を推奨します"
        Case Else
            ConfidenceBanner = "🟡 AIの自信度: 中 — 事例・推論を含みます。出典を確認し、迷えば『📮 教えてBOX』へ"
    End Select
End Function

' Form-control buttons can't be recolored, so we nudge via the caption.
Private Sub SetInquiryProminence(ByVal ws As Worksheet, ByVal c As String)
    On Error Resume Next
    Dim cap As String
    Select Case c
        Case "中", "低": cap = "📮 教えてBOX ← 本社へ照会を推奨"
        Case "高": cap = "📮 教えてBOX（迷う時のみ）"
        Case Else: cap = "📮 教えてBOX（本社へ照会）"
    End Select
    ws.Buttons(BTN_INQUIRY).caption = cap
    On Error GoTo 0
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

' Defensive cell setter. Excel raises Err 1004 when:
'   - the string starts with '=' '+' '-' '@' (parsed as a formula)
'   - the string contains an embedded NULL byte (Chr(0))
'   - the string exceeds 32767 characters (cell limit)
' Each cause is handled before the assignment, with the final write
' wrapped in On Error Resume Next so a bad payload never crashes the UI.
Public Sub SetCellSafe(ByVal cell As Range, ByVal text As String)
    Dim t As String: t = CStr(text)
    ' Strip NUL bytes that some LLM responses carry
    If InStr(t, Chr(0)) > 0 Then t = Replace(t, Chr(0), "")
    ' Truncate to Excel's per-cell limit (32767), leaving headroom
    If Len(t) > 32000 Then
        t = Left$(t, 31900) & vbLf & "(...以降省略 / 全文はテキスト保存ボタンから取得してください)"
    End If
    ' Prevent Excel from parsing as a formula
    If LenB(t) > 0 Then
        Dim ch As String: ch = Left$(t, 1)
        If ch = "=" Or ch = "+" Or ch = "-" Or ch = "@" Then
            t = " " & t
        End If
    End If
    On Error Resume Next
    cell.value = t
    If Err.Number <> 0 Then
        ' Last-resort fallback: tell the user but never crash
        Err.Clear
        cell.value = "(セル書き込みに失敗。回答全文はテキスト保存ボタンから取得してください)"
    End If
    On Error GoTo 0
End Sub

Private Sub AutoSizeAnswer(ByVal ws As Worksheet)
    ' Size the merged answer block (B10:H40) so the WHOLE answer is visible.
    ' B:H = 7 cols × 14 width ≈ 98 units. Japanese full-width chars take 2 units each,
    ' so Len() of ~44 full-width chars fills one visual line at 10pt.
    Dim ans As String: ans = CStr(ws.Range(CELL_ANSWER).value)
    Const CHARS_PER_LINE As Long = 44
    Const ROWS As Long = 31             ' rows 10..40
    Dim segs() As String: segs = Split(ans, vbLf)
    Dim totalLines As Long, i As Long
    For i = LBound(segs) To UBound(segs)
        Dim segLen As Long: segLen = Len(segs(i))
        If segLen = 0 Then
            totalLines = totalLines + 1
        Else
            totalLines = totalLines + ((segLen + CHARS_PER_LINE - 1) \ CHARS_PER_LINE)
        End If
    Next i
    totalLines = totalLines + 3          ' safety buffer
    Dim perRow As Double
    perRow = (totalLines * 15#) / ROWS
    If perRow < 15 Then perRow = 15
    If perRow > 409 Then perRow = 409   ' Excel's per-row maximum
    For i = 10 To 40
        ws.Rows(i).RowHeight = perRow
    Next i
End Sub

Private Sub AutoSizeCitations(ByVal ws As Worksheet)
    ' Size the merged citations block (B42:H50) so all citation lines are visible.
    ' Citation font is 9pt; ~13 pt per display line. Same column width as answer.
    Dim cit As String: cit = CStr(ws.Range(CELL_CITATIONS).value)
    If LenB(cit) = 0 Then Exit Sub
    Const CHARS_PER_LINE As Long = 48
    Const ROWS As Long = 9              ' rows 42..50
    Dim segs() As String: segs = Split(cit, vbLf)
    Dim totalLines As Long, i As Long
    For i = LBound(segs) To UBound(segs)
        Dim segLen As Long: segLen = Len(segs(i))
        If segLen = 0 Then
            totalLines = totalLines + 1
        Else
            totalLines = totalLines + ((segLen + CHARS_PER_LINE - 1) \ CHARS_PER_LINE)
        End If
    Next i
    totalLines = totalLines + 2
    Dim perRow As Double
    perRow = (totalLines * 13#) / ROWS
    If perRow < 13 Then perRow = 13
    If perRow > 409 Then perRow = 409
    For i = 42 To 50
        ws.Rows(i).RowHeight = perRow
    Next i
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
