Attribute VB_Name = "modKnowledge"
Option Explicit

' ============================================================================
' modKnowledge - Knowledge統合画面(ギャラリー/テーブル モード切替)
' ----------------------------------------------------------------------------
' 役割:
'   旧modVault(ナレッジ倉庫ギャラリー)と旧modUIShelf(マイ本棚テーブル)を
'   1シート2モードで統合する。ヘッダーのトグルボタンで切替。
'
' 設計判断:
'   ・モードはmodConfig("knowledge_view")に永続化("gallery"/"table")
'   ・ギャラリーモード: modVault.ShowVaultGalleryをそのまま呼ぶ
'   ・テーブルモード: modUIShelf.EnsureLayoutをそのまま呼ぶ
'   ・ヘッダー(トグル+Hub戻る)のみmodKnowledgeが描画(nx_kn_プレフィックス)
'   ・既存モジュールの内部ロジックには一切手を入れない(委譲パターン)
' ============================================================================

Private Const KN_SHEET As String = "Shelf"

' ----------------------------------------------------------------------------
' EnsureKnowledgeLayout - Knowledge画面を構築(冪等)
' ----------------------------------------------------------------------------
Public Sub EnsureKnowledgeLayout()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(KN_SHEET)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("Vault")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False

    ' ヘッダー再構築
    RemoveKnShapes ws
    DrawKnHeader ws

    ' モードに応じて委譲
    Dim viewMode As String
    On Error Resume Next
    viewMode = modConfig.GetString("knowledge_view", "gallery")
    On Error GoTo 0

    If viewMode = "table" Then
        modUIShelf.EnsureLayout
    Else
        modVault.ShowVaultGallery
    End If

    Application.ScreenUpdating = True
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    modLog.LogError "E0803", "modKnowledge.EnsureKnowledgeLayout", Err.Description
End Sub

' ----------------------------------------------------------------------------
' ヘッダー(トグル+戻る)
' ----------------------------------------------------------------------------
Private Sub DrawKnHeader(ByVal ws As Worksheet)
    ' 背景バー
    Dim hdr As Shape
    Set hdr = ws.Shapes.AddShape(1, 0, 0, 720, 36)
    hdr.Name = "nx_kn_hdr"
    hdr.Line.Visible = 0
    hdr.Fill.ForeColor.RGB = RGB(255, 255, 255)
    With hdr.TextFrame2
        .TextRange.Text = "  " & ChrW(&HD83D) & ChrW(&HDCDA) & " ナレッジ"
        .TextRange.Font.Size = 12
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(30, 41, 59)
        .VerticalAnchor = 3
        .MarginLeft = 8
    End With

    ' ギャラリー/テーブル トグル
    Dim viewMode As String
    On Error Resume Next
    viewMode = modConfig.GetString("knowledge_view", "gallery")
    On Error GoTo 0

    Dim togG As Shape
    Set togG = ws.Shapes.AddShape(5, 480, 6, 70, 24)
    togG.Name = "nx_kn_tog_g"
    togG.Adjustments(1) = 0.3
    togG.Line.Visible = 0
    If viewMode = "gallery" Then
        togG.Fill.ForeColor.RGB = RGB(31, 78, 120)
        togG.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Else
        togG.Fill.ForeColor.RGB = RGB(241, 245, 249)
        togG.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(71, 85, 105)
    End If
    With togG.TextFrame2
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDDC2) & " ギャラリー"
        .TextRange.Font.Size = 8.5
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    togG.OnAction = "modKnowledge.OnToggleGallery"

    Dim togT As Shape
    Set togT = ws.Shapes.AddShape(5, 554, 6, 60, 24)
    togT.Name = "nx_kn_tog_t"
    togT.Adjustments(1) = 0.3
    togT.Line.Visible = 0
    If viewMode = "table" Then
        togT.Fill.ForeColor.RGB = RGB(31, 78, 120)
        togT.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Else
        togT.Fill.ForeColor.RGB = RGB(241, 245, 249)
        togT.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(71, 85, 105)
    End If
    With togT.TextFrame2
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCB) & " テーブル"
        .TextRange.Font.Size = 8.5
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    togT.OnAction = "modKnowledge.OnToggleTable"

    ' Hub戻る
    Dim btnBack As Shape
    Set btnBack = ws.Shapes.AddShape(5, 660, 6, 50, 24)
    btnBack.Name = "nx_kn_back"
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
    btnBack.OnAction = "modKnowledge.OnBackToHub"
End Sub

' ============================================================================
' ボタンハンドラ
' ============================================================================

Public Sub OnToggleGallery()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modConfig.SetValue "knowledge_view", "gallery"
    EnsureKnowledgeLayout
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnToggleTable()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modConfig.SetValue "knowledge_view", "table"
    EnsureKnowledgeLayout
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnBackToHub()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modHub.EnsureHubLayout
    On Error GoTo 0
    modUiLock.Leave
End Sub

' ============================================================================
' 内部ヘルパー
' ============================================================================

Private Sub RemoveKnShapes(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim names() As String
    Dim n As Long: n = 0
    For Each shp In ws.Shapes
        If Left$(shp.Name, 6) = "nx_kn_" Then
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
