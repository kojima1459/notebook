Attribute VB_Name = "modChat"
Option Explicit

' ============================================================================
' modChat - Chat専用画面(サイドバーなし・フルワイド)
' ----------------------------------------------------------------------------
' 役割:
'   チャット専用タブのレイアウトを構築する。旧Nexusサイドバーを廃止し、
'   画面全幅をチャットバブルに使う。入力はfrmChatInput(モードレスUserForm)。
'
' 設計判断:
'   ・サイドバー完全削除 → バブル表示領域が約30%拡大
'   ・ヘッダー: 簡素な1行(タイトル+Hub戻る+クリア+モード切替)
'   ・既存のAddChatBubble/RenderSourcesPreviewはそのまま利用(modUIMain)
'   ・nx_chat_プレフィックスでShape管理(テーマ再彩色対応)
'   ・frmChatInputはmodChatInput.ShowChatInputで表示
' ============================================================================

Private Const CHAT_SHEET As String = "Nexus"

' ----------------------------------------------------------------------------
' EnsureChatLayout - Chat画面を構築(冪等)
' ----------------------------------------------------------------------------
Public Sub EnsureChatLayout()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CHAT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False

    ' 旧サイドバーShapeを削除(nx_sb_プレフィックス)
    RemoveOldSidebar ws

    ' Chat用ヘッダーを再構築
    RemoveChatShapes ws
    DrawChatHeader ws

    ' 入力欄表示
    modChatInput.ShowChatInput

    Application.ScreenUpdating = True
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    modLog.LogError "E0802", "modChat.EnsureChatLayout", Err.Description
End Sub

' ----------------------------------------------------------------------------
' Chatヘッダー(簡素: 1行)
' ----------------------------------------------------------------------------
Private Sub DrawChatHeader(ByVal ws As Worksheet)
    ' ヘッダー背景
    Dim hdr As Shape
    Set hdr = ws.Shapes.AddShape(1, 0, 0, 720, 36)  ' Rectangle
    hdr.Name = "nx_chat_hdr"
    hdr.Line.Visible = 0
    hdr.Fill.ForeColor.RGB = RGB(255, 255, 255)
    With hdr.TextFrame2
        .TextRange.Text = "  " & ChrW(&HD83D) & ChrW(&HDCAC) & " チャット"
        .TextRange.Font.Size = 12
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(30, 41, 59)
        .VerticalAnchor = 3
        .MarginLeft = 8
    End With

    ' Hub戻るボタン
    Dim btnBack As Shape
    Set btnBack = ws.Shapes.AddShape(5, 620, 6, 44, 24)
    btnBack.Name = "nx_chat_back"
    btnBack.Adjustments(1) = 0.3
    btnBack.Line.Visible = 0
    btnBack.Fill.ForeColor.RGB = RGB(241, 245, 249)
    With btnBack.TextFrame2
        .TextRange.Text = ChrW(&H2190) & " Hub"
        .TextRange.Font.Size = 9
        .TextRange.Font.Fill.ForeColor.RGB = RGB(71, 85, 105)
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    btnBack.OnAction = "modChat.OnBackToHub"

    ' クリアボタン
    Dim btnClr As Shape
    Set btnClr = ws.Shapes.AddShape(5, 668, 6, 44, 24)
    btnClr.Name = "nx_chat_clr"
    btnClr.Adjustments(1) = 0.3
    btnClr.Line.Visible = 0
    btnClr.Fill.ForeColor.RGB = RGB(254, 242, 242)  ' 薄い赤
    With btnClr.TextFrame2
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDDE1) & " 消去"
        .TextRange.Font.Size = 9
        .TextRange.Font.Fill.ForeColor.RGB = RGB(220, 38, 38)
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    btnClr.OnAction = "modChat.OnClearChat"

    ' モード切替(すぐ聞く/しっかり調べる)
    Dim btnMode As Shape
    Set btnMode = ws.Shapes.AddShape(5, 540, 6, 74, 24)
    btnMode.Name = "nx_chat_mode"
    btnMode.Adjustments(1) = 0.3
    btnMode.Line.ForeColor.RGB = RGB(232, 236, 240)
    btnMode.Line.Weight = 0.75
    btnMode.Fill.ForeColor.RGB = RGB(255, 255, 255)
    With btnMode.TextFrame2
        .TextRange.Text = GetModeLabel()
        .TextRange.Font.Size = 8.5
        .TextRange.Font.Fill.ForeColor.RGB = RGB(31, 78, 120)
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    btnMode.OnAction = "modChat.OnToggleMode"
End Sub

' ============================================================================
' ボタンハンドラ
' ============================================================================

Public Sub OnBackToHub()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modChatInput.HideChatInput
    modHub.EnsureHubLayout
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnClearChat()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    ' 既存のクリア処理を流用
    modApp.OnClearChat
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnToggleMode()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    Dim cur As String
    cur = modConfig.GetString("qa_mode", "quick")
    If cur = "quick" Then
        modConfig.SetValue "qa_mode", "deep"
    Else
        modConfig.SetValue "qa_mode", "quick"
    End If
    ' ボタンラベル更新
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CHAT_SHEET)
    ws.Shapes("nx_chat_mode").TextFrame2.TextRange.Text = GetModeLabel()
    modSkin.ShowToast "モード: " & GetModeLabel(), "info"
    On Error GoTo 0
    modUiLock.Leave
End Sub

' ============================================================================
' 内部ヘルパー
' ============================================================================

Private Function GetModeLabel() As String
    On Error Resume Next
    Dim m As String: m = modConfig.GetString("qa_mode", "quick")
    If m = "deep" Then
        GetModeLabel = ChrW(&HD83D) & ChrW(&HDD0D) & " しっかり調べる"
    Else
        GetModeLabel = ChrW(&HD83D) & ChrW(&HDCA8) & " すぐ聞く"
    End If
    On Error GoTo 0
End Function

Private Sub RemoveOldSidebar(ByVal ws As Worksheet)
    ' 旧Nexusサイドバー(nx_sb_プレフィックス)を削除
    Dim shp As Shape
    Dim names() As String
    Dim n As Long: n = 0
    For Each shp In ws.Shapes
        If Left$(shp.Name, 6) = "nx_sb_" Or Left$(shp.Name, 7) = "nx_side" Then
            ReDim Preserve names(n)
            names(n) = shp.Name
            n = n + 1
        End If
    Next shp
    Dim i As Long
    For i = 0 To n - 1
        On Error Resume Next
        ws.Shapes(names(i)).Delete
        On Error GoTo 0
    Next i
End Sub

Private Sub RemoveChatShapes(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim names() As String
    Dim n As Long: n = 0
    For Each shp In ws.Shapes
        If Left$(shp.Name, 8) = "nx_chat_" Then
            ReDim Preserve names(n)
            names(n) = shp.Name
            n = n + 1
        End If
    Next shp
    Dim i As Long
    For i = 0 To n - 1
        On Error Resume Next
        ws.Shapes(names(i)).Delete
        On Error GoTo 0
    Next i
End Sub
