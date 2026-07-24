Attribute VB_Name = "modHub"
Option Explicit

' ============================================================================
' modHub - Hub画面(ダッシュボード+ナビゲーション+プロフィール統合)
' ----------------------------------------------------------------------------
' 役割:
'   起動時の既定画面。旧Dashboard(modUIDashboard)と旧サイドバー(modUINexusDraw)
'   の全要素を1画面に統合し、「ごちゃつきを分散」させる3タブ設計の入口。
'
' 設計判断:
'   ・ヘッダー/プロフィール/ナビボタン/ガチャ = Shape(丸角+影+グラデーション)
'   ・統計タイル/バッジ/クイックチップ = セル(Merge+Interior.Color)
'   ・EXPゲージ = Shape(Rectangle×2: 背景+Fill)
'   ・全Shapeは nx_hub_ プレフィックス(既存のテーマ再彩色/ZOrderループが自動対応)
'   ・ボタンハンドラは全て modUiLock.Enter/Leave で保護(連打防止)
'   ・色は modUI.UiColor() から取得(ハードコード禁止)
' ============================================================================

Private Const HUB_SHEET As String = "Home"

' レイアウト定数
Private Const HDR_TOP As Double = 0
Private Const HDR_H As Double = 42
Private Const BODY_TOP As Double = 52
Private Const CARD_W As Double = 340
Private Const NAV_W As Double = 300
Private Const NAV_H As Double = 48
Private Const NAV_GAP As Double = 10
Private Const NAV_LEFT As Double = 380

' ----------------------------------------------------------------------------
' EnsureHubLayout - Hub画面を構築(冪等: 何度呼んでも同じ結果)
' ----------------------------------------------------------------------------
Public Sub EnsureHubLayout()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HUB_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False

    ' 既存のnx_hub_ Shapeを全削除(冪等性確保)
    RemoveHubShapes ws

    ' セル初期化
    ws.Cells.Clear
    ws.Cells.Font.Name = "Yu Gothic UI"
    ws.Cells.Font.Size = 10
    ws.Cells.Interior.Color = modUI.UiColor("bg")

    ' 列幅設定
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B:E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 3
    ws.Columns("G:J").ColumnWidth = 12
    ws.Columns("K").ColumnWidth = 2

    ' 枠線・見出し非表示
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    On Error GoTo 0

    ' ---- ヘッダーバー(Shape: グラデーション丸角) ----
    DrawHeader ws

    ' ---- プロフィールカード(Shape: 影付き丸角) ----
    DrawProfileCard ws

    ' ---- 統計タイル(セル) ----
    DrawStatTiles ws

    ' ---- ナビボタン(Shape×4) ----
    DrawNavButtons ws

    ' ---- クイックチップ + ガチャ(セル+Shape) ----
    DrawExtras ws

    ' ---- バッジ(セル) ----
    DrawBadges ws

    ' ---- EXPゲージ(Shape×2) ----
    DrawExpGauge ws

    Application.ScreenUpdating = True
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    modLog.LogError "E0801", "modHub.EnsureHubLayout", Err.Description
End Sub

' ----------------------------------------------------------------------------
' ヘッダーバー
' ----------------------------------------------------------------------------
Private Sub DrawHeader(ByVal ws As Worksheet)
    Dim hdr As Shape
    Set hdr = ws.Shapes.AddShape(5, 0, HDR_TOP, 720, HDR_H)
    hdr.Name = "nx_hub_hdr"
    hdr.Line.Visible = 0
    hdr.Adjustments(1) = 0.02
    With hdr.Fill
        .TwoColorGradient 1, 1   ' msoGradientHorizontal
        .ForeColor.RGB = RGB(26, 54, 93)    ' #1a365d
        .BackColor.RGB = RGB(31, 78, 120)   ' #1f4e78
    End With
    With hdr.TextFrame2
        .TextRange.Text = ChrW(&H26A1) & " Nexus Agent"
        .TextRange.Font.Size = 14
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .MarginLeft = 16
        .VerticalAnchor = 3
    End With

    ' 右側アイコン群(言語/テーマ/ヘルプ/終了)
    Dim icons As Variant, acts As Variant, xs As Variant
    icons = Array(ChrW(&HD83C) & ChrW(&HDF10), ChrW(&HD83C) & ChrW(&HDF19), "?", ChrW(&HD83D) & ChrW(&HDEAA))
    acts = Array("modHub.OnLangCycle", "modHub.OnThemeToggle", "modHub.OnHelp", "modHub.OnSaveAndExit")
    xs = Array(580#, 614#, 648#, 682#)
    Dim i As Long
    For i = 0 To 3
        Dim btn As Shape
        Set btn = ws.Shapes.AddShape(9, xs(i), 8, 26, 26)  ' Oval
        btn.Name = "nx_hub_ic" & i
        btn.Line.Visible = 0
        btn.Fill.ForeColor.RGB = RGB(255, 255, 255)
        btn.Fill.Transparency = 0.85
        With btn.TextFrame2
            .TextRange.Text = CStr(icons(i))
            .TextRange.Font.Size = 11
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = 2
            .VerticalAnchor = 3
        End With
        btn.OnAction = CStr(acts(i))
    Next i
End Sub

' ----------------------------------------------------------------------------
' プロフィールカード
' ----------------------------------------------------------------------------
Private Sub DrawProfileCard(ByVal ws As Worksheet)
    Dim card As Shape
    Set card = ws.Shapes.AddShape(5, 16, BODY_TOP, CARD_W, 72)
    card.Name = "nx_hub_profile"
    card.Adjustments(1) = 0.06
    card.Line.Visible = 0
    card.Fill.ForeColor.RGB = RGB(255, 255, 255)
    SafeShadow card

    Dim uName As String, dept As String
    On Error Resume Next
    uName = modConfig.GetString("user_name", "ゲスト")
    dept = modConfig.GetString("user_dept", "")
    On Error GoTo 0
    If LenB(uName) = 0 Then uName = "ゲスト"

    Dim lvl As Long, exp As Long
    On Error Resume Next
    lvl = modStats.Level()
    exp = modStats.ExpTotal()
    On Error GoTo 0

    With card.TextFrame2
        .WordWrap = -1
        .MarginLeft = 14: .MarginTop = 10
        .TextRange.Text = uName & IIf(LenB(dept) > 0, "  (" & dept & ")", "") & vbLf & _
            "Lv." & lvl & "  /  EXP " & exp
        .TextRange.Font.Size = 11
        .TextRange.Font.Fill.ForeColor.RGB = RGB(30, 41, 59)
        .TextRange.Paragraphs(1).Font.Bold = -1
        .TextRange.Paragraphs(1).Font.Size = 12
        .TextRange.Paragraphs(2).Font.Size = 9
        .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = RGB(148, 163, 184)
    End With
End Sub

' ----------------------------------------------------------------------------
' EXPゲージ(プロフィールカード内)
' ----------------------------------------------------------------------------
Private Sub DrawExpGauge(ByVal ws As Worksheet)
    Dim gaugeTop As Double: gaugeTop = BODY_TOP + 54
    Dim gaugeLeft As Double: gaugeLeft = 30
    Dim gaugeW As Double: gaugeW = CARD_W - 28

    ' 背景バー
    Dim bg As Shape
    Set bg = ws.Shapes.AddShape(5, gaugeLeft, gaugeTop, gaugeW, 8)
    bg.Name = "nx_hub_expbg"
    bg.Adjustments(1) = 0.5
    bg.Line.Visible = 0
    bg.Fill.ForeColor.RGB = RGB(226, 232, 240)  ' #E2E8F0

    ' Fillバー(進捗)
    Dim lvl As Long, exp As Long, divisor As Long
    On Error Resume Next
    lvl = modStats.Level()
    exp = modStats.ExpTotal()
    divisor = modConfig.GetLong("level_divisor", 100)
    If divisor < 1 Then divisor = 100
    On Error GoTo 0

    Dim curLvlExp As Long: curLvlExp = (lvl - 1) ^ 2 * divisor
    Dim nextLvlExp As Long: nextLvlExp = lvl ^ 2 * divisor
    Dim progress As Double
    If nextLvlExp > curLvlExp Then
        progress = CDbl(exp - curLvlExp) / CDbl(nextLvlExp - curLvlExp)
    Else
        progress = 0#
    End If
    If progress < 0 Then progress = 0
    If progress > 1 Then progress = 1

    Dim fillW As Double: fillW = gaugeW * progress
    If fillW < 4 Then fillW = 4
    Dim fl As Shape
    Set fl = ws.Shapes.AddShape(5, gaugeLeft, gaugeTop, fillW, 8)
    fl.Name = "nx_hub_expfill"
    fl.Adjustments(1) = 0.5
    fl.Line.Visible = 0
    With fl.Fill
        .TwoColorGradient 1, 1
        .ForeColor.RGB = RGB(245, 158, 11)   ' #F59E0B
        .BackColor.RGB = RGB(251, 191, 36)   ' #FBBF24
    End With
End Sub

' ----------------------------------------------------------------------------
' 統計タイル(セルベース: modUIDashboard.FormatTile流用)
' ----------------------------------------------------------------------------
Private Sub DrawStatTiles(ByVal ws As Worksheet)
    Dim startRow As Long: startRow = 10
    Dim tiles As Variant
    ' 8タイル: 2列×4行
    tiles = Array( _
        Array("B", "C", "今月の質問", CLng(modStats.GetStat("question_total"))), _
        Array("D", "E", ChrW(&HD83C) & ChrW(&HDFE2) & " 自己解決", CLng(modStats.GetStat("selfsolve_total"))), _
        Array("B", "C", ChrW(&H23F1) & " 節約時間", FmtMin(modStats.SavedMinutesEstimate())), _
        Array("D", "E", ChrW(&HD83D) & ChrW(&HDCD6) & " チャンク", modShelf.TotalChunks()), _
        Array("B", "C", ChrW(&HD83C) & ChrW(&HDFCD) & " みんな(今日)", GetOrgMin("d")), _
        Array("D", "E", ChrW(&HD83C) & ChrW(&HDFCD) & " みんな(今月)", GetOrgMin("m")), _
        Array("B", "C", ChrW(&HD83D) & ChrW(&HDD25) & " 連続ログイン", CLng(modStats.GetStat("streak_days")) & "日"), _
        Array("D", "E", ChrW(&HD83D) & ChrW(&HDCE6) & " パック共有", CLng(modStats.GetStat("pack_export_total"))) _
    )

    Dim i As Long
    For i = 0 To 7
        Dim r As Long: r = startRow + (i \ 2) * 3
        Dim c1 As String: c1 = CStr(tiles(i)(0))
        Dim c2 As String: c2 = CStr(tiles(i)(1))
        Dim label As String: label = CStr(tiles(i)(2))
        Dim val As String: val = CStr(tiles(i)(3))

        ' タイトル行
        With ws.Range(c1 & r & ":" & c2 & r)
            .Merge
            .Value = label
            .Font.Size = 8
            .Font.Color = RGB(148, 163, 184)
            .HorizontalAlignment = -4108  ' xlCenter
        End With
        ' 値行
        With ws.Range(c1 & (r + 1) & ":" & c2 & (r + 1))
            .Merge
            .Value = val
            .Font.Size = 16
            .Font.Bold = True
            .Font.Color = RGB(31, 78, 120)
            .HorizontalAlignment = -4108
            .Interior.Color = RGB(255, 255, 255)
            .Borders.LineStyle = 1
            .Borders.Color = RGB(232, 236, 240)
        End With
        ws.Rows(r + 1).RowHeight = 28
    Next i
End Sub

' ----------------------------------------------------------------------------
' ナビボタン(Shape×4: 丸角+影+OnAction)
' ----------------------------------------------------------------------------
Private Sub DrawNavButtons(ByVal ws As Worksheet)
    Dim defs As Variant, acts As Variant, descs As Variant
    defs = Array(ChrW(&HD83D) & ChrW(&HDCAC) & " チャットで質問する", _
                 ChrW(&HD83D) & ChrW(&HDCDA) & " ナレッジ倉庫", _
                 ChrW(&HD83D) & ChrW(&HDCD6) & " マイ本棚", _
                 ChrW(&HD83D) & ChrW(&HDCE6) & " パック共有(P2P)")
    acts = Array("modHub.OnGoChat", "modHub.OnGoVault", "modHub.OnGoShelf", "modHub.OnGoPack")
    descs = Array("本棚の資料からAIが回答", "資料の検索・登録・同期", "登録済み資料の一覧・管理", "部内でナレッジ共有")

    Dim i As Long
    For i = 0 To 3
        Dim top As Double: top = BODY_TOP + i * (NAV_H + NAV_GAP)
        Dim btn As Shape
        Set btn = ws.Shapes.AddShape(5, NAV_LEFT, top, NAV_W, NAV_H)
        btn.Name = "nx_hub_nav" & i
        btn.Adjustments(1) = 0.08
        btn.Line.ForeColor.RGB = RGB(232, 236, 240)
        btn.Line.Weight = 0.75
        btn.Fill.ForeColor.RGB = RGB(255, 255, 255)
        SafeShadow btn
        With btn.TextFrame2
            .WordWrap = -1
            .MarginLeft = 14: .MarginTop = 6
            .TextRange.Text = CStr(defs(i)) & vbLf & CStr(descs(i))
            .TextRange.Font.Size = 11
            .TextRange.Font.Fill.ForeColor.RGB = RGB(30, 41, 59)
            .TextRange.Paragraphs(1).Font.Bold = -1
            .TextRange.Paragraphs(2).Font.Size = 8
            .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = RGB(148, 163, 184)
            .VerticalAnchor = 3
        End With
        btn.OnAction = CStr(acts(i))
    Next i
End Sub

' ----------------------------------------------------------------------------
' クイックチップ + ガチャ
' ----------------------------------------------------------------------------
Private Sub DrawExtras(ByVal ws As Worksheet)
    Dim chipRow As Long: chipRow = 23
    With ws.Range("G" & chipRow & ":J" & chipRow)
        .Merge
        .Value = ChrW(&HD83D) & ChrW(&HDCA1) & " こんなふうに聞いてみよう"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(71, 85, 105)
    End With

    Dim chips As Variant
    chips = Array("改定ポイントを教えて", "用語をやさしく解説", "手続きの流れを知りたい")
    Dim i As Long
    For i = 0 To 2
        Dim chip As Shape
        Set chip = ws.Shapes.AddShape(5, NAV_LEFT + i * 105, BODY_TOP + 220, 100, 22)
        chip.Name = "nx_hub_qa" & i
        chip.Adjustments(1) = 0.4
        chip.Line.Visible = 0
        chip.Fill.ForeColor.RGB = RGB(241, 245, 249)
        With chip.TextFrame2
            .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCAC) & " " & CStr(chips(i))
            .TextRange.Font.Size = 7.5
            .TextRange.Font.Fill.ForeColor.RGB = RGB(71, 85, 105)
            .TextRange.ParagraphFormat.Alignment = 2
            .VerticalAnchor = 3
            .MarginLeft = 4: .MarginRight = 4
        End With
        chip.OnAction = "modHub.OnQuickAsk"
    Next i

    ' ガチャボタン(点線枠)
    Dim gacha As Shape
    Set gacha = ws.Shapes.AddShape(5, NAV_LEFT, BODY_TOP + 252, NAV_W, 28)
    gacha.Name = "nx_hub_gacha"
    gacha.Adjustments(1) = 0.1
    gacha.Fill.Visible = 0
    gacha.Line.ForeColor.RGB = RGB(167, 139, 250)  ' #A78BFA
    gacha.Line.Weight = 1.5
    gacha.Line.DashStyle = 3  ' msoLineDash
    With gacha.TextFrame2
        .TextRange.Text = ChrW(&HD83C) & ChrW(&HDFB2) & " 今日のワンポイントを引く"
        .TextRange.Font.Size = 9.5
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(124, 58, 237)  ' #7C3AED
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    gacha.OnAction = "modHub.OnGacha"
End Sub

' ----------------------------------------------------------------------------
' バッジ(セル)
' ----------------------------------------------------------------------------
Private Sub DrawBadges(ByVal ws As Worksheet)
    Dim r As Long: r = 26
    With ws.Range("B" & r & ":E" & r)
        .Merge
        .Value = ChrW(&HD83C) & ChrW(&HDFC5) & " バッジ"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(71, 85, 105)
    End With

    Dim ids As Variant, titles As Variant
    ids = Split("first_ingest,shelf10,shelf30,first_pack_out,first_pack_in,solve10,solve50,streak7", ",")
    titles = Split("初取込,本棚10,本棚30,初共有,初取込(P),解決10,解決50,7日連続", ",")

    Dim sb As String: sb = ""
    Dim i As Long
    For i = 0 To UBound(ids)
        Dim earned As Boolean
        On Error Resume Next
        earned = (modStats.GetStat("badge:" & CStr(ids(i))) > 0)
        On Error GoTo 0
        If earned Then
            sb = sb & ChrW(&HD83C) & ChrW(&HDFC5) & " " & CStr(titles(i)) & "  "
        Else
            sb = sb & ChrW(&HD83D) & ChrW(&HDD12) & " " & CStr(titles(i)) & "  "
        End If
    Next i

    With ws.Range("B" & (r + 1) & ":E" & (r + 2))
        .Merge
        .WrapText = True
        .Value = sb
        .Font.Size = 8.5
        .Font.Color = RGB(100, 116, 139)
        .VerticalAlignment = -4160
    End With
End Sub

' ============================================================================
' ボタンハンドラ
' ============================================================================

Public Sub OnGoChat()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    modUI.GoToNexus "modHub.OnGoChat"
    modChat.EnsureChatLayout
Done:
    modUiLock.Leave
End Sub

Public Sub OnGoVault()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modConfig.SetValue "knowledge_view", "gallery"
    modKnowledge.EnsureKnowledgeLayout
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnGoShelf()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modConfig.SetValue "knowledge_view", "table"
    modKnowledge.EnsureKnowledgeLayout
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnGoPack()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modConfig.SetValue "knowledge_view", "gallery"
    modKnowledge.EnsureKnowledgeLayout
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnQuickAsk()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    Dim tpl As String
    Select Case CStr(Application.Caller)
        Case "nx_hub_qa0": tpl = "【知りたい改定】: (資料名や年度)" & vbLf & "【気になる点】: (例: 保険料への影響)"
        Case "nx_hub_qa1": tpl = "【わからない用語】: (ここに記入)" & vbLf & "【どこまで理解したいか】: (例: お客様に説明できるレベル)"
        Case "nx_hub_qa2": tpl = "【手続き名】: (ここに記入)" & vbLf & "【知りたい結論】: (例: 必要書類と所要日数)"
        Case Else: GoTo Done
    End Select
    ' Chat遷移 + 入力欄にテンプレセット
    modUI.GoToNexus "modHub.OnQuickAsk"
    On Error Resume Next
    ThisWorkbook.Names("nx_input").RefersToRange.Value = tpl
    On Error GoTo 0
    modSkin.ShowToast "(ここに記入)を埋めて Ctrl+Enter で送信してください。", "info"
Done:
    modUiLock.Leave
End Sub

Public Sub OnGacha()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modGacha.ShowGacha
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnLangCycle()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    Dim cur As String: cur = modConfig.GetString("answer_language", "日本語")
    Dim nxt As String
    Select Case cur
        Case "日本語": nxt = "English"
        Case "English": nxt = "中文"
        Case Else: nxt = "日本語"
    End Select
    modConfig.SetValue "answer_language", nxt
    modSkin.ShowToast "回答言語: " & nxt, "success"
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnThemeToggle()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modSkin.CycleSkin
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnHelp()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modHelp.OnHelpClick
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnSaveAndExit()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modApp.OnSaveAndExit
    On Error GoTo 0
    modUiLock.Leave
End Sub

' ============================================================================
' 内部ヘルパー
' ============================================================================

Private Sub RemoveHubShapes(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim names() As String
    Dim n As Long: n = 0
    For Each shp In ws.Shapes
        If Left$(shp.Name, 7) = "nx_hub_" Then
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

Private Sub SafeShadow(ByVal shp As Shape)
    On Error Resume Next
    With shp.Shadow
        .Visible = -1
        .Blur = 6
        .OffsetY = 2
        .Transparency = 0.85
        .ForeColor.RGB = RGB(0, 0, 0)
    End With
    On Error GoTo 0
End Sub

Private Function FmtMin(ByVal minutes As Long) As String
    If minutes < 60 Then
        FmtMin = minutes & "分"
    Else
        FmtMin = (minutes \ 60) & "時間" & IIf(minutes Mod 60 > 0, (minutes Mod 60) & "分", "")
    End If
End Function

Private Function GetOrgMin(ByVal period As String) As String
    On Error Resume Next
    Dim v As Long
    If period = "d" Then
        v = modBoard.OrgMinutesDay()
    Else
        v = modBoard.OrgMinutesMon()
    End If
    GetOrgMin = FmtMin(v)
    On Error GoTo 0
    If LenB(GetOrgMin) = 0 Then GetOrgMin = "—"
End Function
