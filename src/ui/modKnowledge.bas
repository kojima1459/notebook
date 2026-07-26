Attribute VB_Name = "modKnowledge"
Option Explicit

' modKnowledge - ナレッジ画面の共通クロム(ヘッダー+モード切替+ツールバー)。
'
' 2026-07-26 再設計(nexus-spec-v1 §2.3 / nexus-ui-final 画面3):
'   実機報告「ナレッジ倉庫とマイ本棚の違いが分からない」への対処。
'   中身は今までどおり2枚のシート(Vault=カードギャラリー / マイ本棚=一覧表)
'   だが、タブは隠してあるので利用者にシートの区別は見えない。両方に
'   まったく同じヘッダーとツールバーを描き、上部のピルで
'   「🃏 ギャラリー / 📋 マイ本棚」を切り替える1画面2モードとして見せる。
'
'   ツールバーの各ボタンは既存のmodVault/modUIShelf/modPack/modShelfSyncへ
'   そのまま配線するだけで、取込・同期・パックのロジックには一切触っていない
'   (仕様書の「既存コード変更最小化」原則)。これまで2画面に重複していた
'   パック出力/取込の入口も、この1本のツールバーに集約した。
'
' 設計の鉄則:
'   ・Shape座標は実セル幾何(Range.Left/.Width/Rows().Top)から導く。
'     旧modVaultのツールバーはx=390,460,528...と決め打ちで、列幅を変えると
'     すぐ画面外へはみ出していた。
'   ・配色は modUI.UiColor()。絵文字はChrW()で組み立てる。
'   ・Shape名は nxk_ 接頭辞(冪等な全削除に使う)。

Private Const HDR_H As Double = 40
Private Const BAR_H As Double = 24
Private Const PILL_W As Double = 88

' 上部クロムが占める行(1..6)。本文はDrawChrome後の ContentTop から下に描く。
Public Const CHROME_ROWS As Long = 6

' DrawChrome - ヘッダー+モードピル+ツールバーを描く(冪等)。
'   mode: "gallery"(ナレッジ倉庫のカード) / "table"(マイ本棚の一覧)
Public Sub DrawChrome(ByVal ws As Worksheet, ByVal mode As String)
    If ws Is Nothing Then Exit Sub
    On Error GoTo Fail

    RemoveChrome ws

    ' 幾何を先に確定させる(順序が逆だとShape座標がズレる)。
    ws.Rows(1).RowHeight = HDR_H
    ws.Rows(2).RowHeight = 6
    ws.Rows(3).RowHeight = BAR_H
    ws.Rows(4).RowHeight = 6
    ws.Rows(5).RowHeight = 22
    ws.Rows(6).RowHeight = 8

    Dim L As Double, W As Double
    L = ws.Range("A1").Left
    W = ws.Range("A1:N1").Width

    ' --- ヘッダーバー ---
    Dim hdr As Shape
    Set hdr = ws.Shapes.AddShape(5, L, 0, W, HDR_H)
    hdr.Name = "nxk_hdr"
    hdr.Adjustments(1) = 0.02
    hdr.Line.Visible = 0
    hdr.Fill.ForeColor.RGB = modUI.UiColor("sidebar")
    With hdr.TextFrame2
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCDA) & " ナレッジ"
        .TextRange.Font.Size = 12
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .MarginLeft = 82
        .VerticalAnchor = 3
    End With

    Pill ws, "nxk_back", ChrW(&H2190) & " Hub", L + 8, 62, _
         "modKnowledge.OnBackHub", False

    ' --- モード切替ピル(右肩) ---
    Dim isTable As Boolean: isTable = (LCase$(mode) = "table")
    Dim px As Double: px = L + W - 8 - PILL_W
    Pill ws, "nxk_m_table", ChrW(&HD83D) & ChrW(&HDCCB) & " マイ本棚", px, PILL_W, _
         "modKnowledge.OnGoTable", isTable
    px = px - 6 - PILL_W
    Pill ws, "nxk_m_gallery", ChrW(&HD83C) & ChrW(&HDCCF) & " ギャラリー", px, PILL_W, _
         "modKnowledge.OnGoGallery", Not isTable

    ' --- ツールバー(行3の帯) ---
    DrawToolbar ws, isTable, L, W

    On Error Resume Next
    modUI.FreezeShapePlacement ws
    On Error GoTo 0
    Exit Sub

Fail:
    modLog.LogError "E0801", "modKnowledge.DrawChrome", Err.Description, Err.Number
End Sub

' 本文(カード/表)を描き始めてよいY座標。
Public Function ContentTop(ByVal ws As Worksheet) As Double
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    ContentTop = ws.Rows(CHROME_ROWS + 1).Top
    On Error GoTo 0
End Function

' 検索キーワードのセル(ギャラリーだけが使う)。
Public Function SearchCellAddress() As String
    SearchCellAddress = "B5:E5"
End Function

Private Sub DrawToolbar(ByVal ws As Worksheet, ByVal isTable As Boolean, _
                        ByVal L As Double, ByVal W As Double)
    Dim caps As Variant, acts As Variant, widths As Variant
    caps = Array(ChrW(&HD83D) & ChrW(&HDD0D) & " 検索", _
                 ChrW(&H2795) & " 登録", _
                 ChrW(&HD83D) & ChrW(&HDCC1) & " 追加", _
                 ChrW(&HD83D) & ChrW(&HDCE6) & " パック出力", _
                 ChrW(&HD83D) & ChrW(&HDCE5) & " パック取込", _
                 ChrW(&HD83D) & ChrW(&HDD04) & " 同期", _
                 ChrW(&HD83D) & ChrW(&HDCC2) & " フォルダ", _
                 ChrW(&HD83D) & ChrW(&HDDD1) & " 削除", _
                 ChrW(&HD83D) & ChrW(&HDCAC) & " チャットへ")
    acts = Array("OnSearch", "OnRegister", "OnAddFiles", "OnPackOut", "OnPackIn", _
                 "OnSync", "OnPickFolder", "OnDelete", "OnToChat")
    widths = Array(62, 58, 58, 80, 80, 54, 68, 58, 76)

    Dim barTop As Double: barTop = ws.Rows(3).Top
    Dim x As Double: x = L + 8
    Dim i As Long
    For i = 0 To 8
        ' 検索はギャラリー専用、削除は一覧表専用(押しても何も起きないボタンを
        ' 見せない=実機報告「どっちで押せばいいか分からない」への対処)。
        Dim skip As Boolean
        skip = (i = 0 And isTable) Or (i = 7 And Not isTable)
        If Not skip Then
            ' 1個の1004で残りを道連れにしない。
            On Error Resume Next
            Dim btn As Shape
            Set btn = ws.Shapes.AddShape(5, x, barTop, CDbl(widths(i)), BAR_H)
            If Err.Number = 0 And Not btn Is Nothing Then
                btn.Name = "nxk_tb" & i
                btn.Adjustments(1) = 0.35
                btn.Line.Visible = -1
                btn.Line.Weight = 0.75
                btn.Line.ForeColor.RGB = modUI.UiColor("border")
                btn.Fill.ForeColor.RGB = modUI.UiColor("surface")
                With btn.TextFrame2
                    .WordWrap = -1
                    .TextRange.Text = CStr(caps(i))
                    .TextRange.Font.Size = 8.5
                    .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
                    .TextRange.ParagraphFormat.Alignment = 2
                    .VerticalAnchor = 3
                    .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
                End With
                btn.OnAction = "modKnowledge." & CStr(acts(i))
            End If
            Set btn = Nothing
            Err.Clear
            On Error GoTo 0
            x = x + CDbl(widths(i)) + 5
        End If
    Next i

    ' 画像解析が使える環境でだけスクショ取込を出す(無効環境で「押したら
    ' 断られるボタン」を見せない。既存modUIShelfの方針をそのまま踏襲)。
    On Error Resume Next
    If modFeatures.FeatureEnabled("vision") Then
        Dim sc As Shape
        Set sc = ws.Shapes.AddShape(5, x, barTop, 84, BAR_H)
        If Not sc Is Nothing Then
            sc.Name = "nxk_tbshot"
            sc.Adjustments(1) = 0.35
            sc.Line.Visible = -1
            sc.Line.Weight = 0.75
            sc.Line.ForeColor.RGB = modUI.UiColor("border")
            sc.Fill.ForeColor.RGB = modUI.UiColor("surface")
            With sc.TextFrame2
                .WordWrap = -1
                .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCF8) & " スクショ取込"
                .TextRange.Font.Size = 8.5
                .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
                .TextRange.ParagraphFormat.Alignment = 2
                .VerticalAnchor = 3
                .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
            End With
            sc.OnAction = "modUIShelf.OnIngestScreenshot"
        End If
    End If
    On Error GoTo 0
End Sub

' ヘッダー上のピル。active:=Trueで「今いるモード」を塗りつぶして示す。
Private Sub Pill(ByVal ws As Worksheet, ByVal shapeName As String, _
                 ByVal caption As String, ByVal x As Double, ByVal w As Double, _
                 ByVal action As String, ByVal active As Boolean)
    On Error Resume Next
    Dim p As Shape
    Set p = ws.Shapes.AddShape(5, x, (HDR_H - 26) / 2, w, 26)
    If p Is Nothing Then Exit Sub
    p.Name = shapeName
    p.Adjustments(1) = 0.35
    p.Line.Visible = 0
    If active Then
        p.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Else
        p.Fill.ForeColor.RGB = modUI.UiColor("sidebarActive")
    End If
    With p.TextFrame2
        .TextRange.Text = caption
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = -1
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
    End With
    If active Then
        p.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("sidebar")
    Else
        p.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End If
    p.OnAction = action
    On Error GoTo 0
End Sub

Private Sub RemoveChrome(ByVal ws As Worksheet)
    Dim names() As String
    ReDim names(0 To ws.Shapes.Count)
    Dim n As Long
    Dim shp As Shape
    For Each shp In ws.Shapes
        ' 旧マイ本棚のボタン(btn_)も一緒に消す。残すと新ツールバーの上に
        ' 浮いたまま二重表示になる(Hub移植時に踏んだのと同じ罠)。
        If Left$(shp.Name, 4) = "nxk_" Or Left$(shp.Name, 4) = "btn_" Then
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

' ---- ツールバーのハンドラ(既存エンジンへの配線に徹する) ----

Public Sub OnGoGallery()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modVault.ShowVaultGallery
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnGoTable()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUIShelf.EnsureLayout
    modUI.GoToNativeSheet modAppDef.SH_SHELF, "modKnowledge.OnGoTable"
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnBackHub()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modHub.EnsureHubLayout activate:=True
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnToChat()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUI.GoToNexus "modKnowledge.OnToChat"
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnSearch()
    modVault.OnVaultSearch
End Sub

Public Sub OnRegister()
    modVault.ShowVaultInput
End Sub

Public Sub OnAddFiles()
    On Error Resume Next
    modShelf.AddFilesViaDialog
    On Error GoTo 0
    RefreshCurrent
End Sub

Public Sub OnPackOut()
    modPack.ExportPackDialog
End Sub

Public Sub OnPackIn()
    modPack.ImportPackDialog
    RefreshCurrent
End Sub

Public Sub OnSync()
    modShelfSync.SyncNow
    RefreshCurrent
End Sub

Public Sub OnPickFolder()
    modShelfSync.PickShelfFolder
    RefreshCurrent
End Sub

Public Sub OnDelete()
    modUIShelf.OnDeleteSource
    RefreshCurrent
End Sub

' 今表示しているモードだけを描き直す(モード切替をまたいで表示がズレないよう、
' 資料を足した/消した直後は必ずここを通す)。
Private Sub RefreshCurrent()
    On Error Resume Next
    If ThisWorkbook.ActiveSheet.Name = modAppDef.SH_SHELF Then
        modUIShelf.RenderShelf
    Else
        modVault.ShowVaultGallery
    End If
    On Error GoTo 0
End Sub
