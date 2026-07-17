Attribute VB_Name = "modUI"
Option Explicit

' ============================================================================
' modUI - Nexus Agent UIコア(DOCS_NEXUS_SPEC Phase 1 / HTMLモック完全再現)
' ----------------------------------------------------------------------------
' 対象実行環境はWindows版Microsoft Excelのみ。SPA風UI:
'   InitUI        - ネイティブUI(リボン/枠線/数式バー等)の完全隠蔽+骨格描画
'   AddChatBubble - チャットバブル(角丸Shape)+アクションボタン列の動的生成
'   ToggleTheme   - ライト/ダークモードの即時反転(全Shape再彩色)
'   RestoreExcelUI- ネイティブUIの復元(Auto_Close時に必ず呼ぶこと)
' 設計メモ:
'   ・アクションボタンはShape+OnAction方式。仕様書のWorksheet_FollowHyperlink
'     方式は自己インストーラ配布(標準モジュールのみ注入可能・シートモジュール
'     不可)と両立しないため、同一UXをOnActionで実現する(配布性を優先)。
'   ・UserFormモーダルも同じ理由でShapeオーバーレイ方式(Phase 2)。
'   ・テーマ状態はui_stateシート(key="nexus_theme")に永続化。
'   ・Shape命名規約: nx_sb_*(サイドバー) nx_top_*(トップバー)
'     nx_msg_*(バブル) nx_act_*(アクションボタン) nx_thk_*(思考プロセス)
' ============================================================================

Private Const NEXUS_SHEET As String = "Nexus"
Private Const THEME_KEY As String = "nexus_theme"

' レイアウト(ポイント単位)
Private Const SIDEBAR_W As Double = 195      ' 260px相当
Private Const TOPBAR_H As Double = 45
Private Const CHAT_LEFT_PAD As Double = 15
Private Const BUBBLE_RATIO As Double = 0.62  ' チャット幅に対するバブル最大幅
Private Const BUBBLE_GAP As Double = 14
Private Const ACT_H As Double = 24
Private Const CHAT_TOP As Double = 130   ' 固定領域(トップバー+入力+アクションバー)の直下

Private mChatBottom As Double   ' 最後のバブルの下端(モジュール状態リセット時はRecalc)

' ----------------------------------------------------------------------------
' InitUI - ネイティブUI隠蔽+Nexus骨格(サイドバー/トップバー/入力欄)描画
' ----------------------------------------------------------------------------
Public Sub InitUI()
    Dim ws As Worksheet
    Set ws = GetOrCreateNexusSheet()
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' --- ネイティブUIの完全隠蔽(失敗しても続行: 環境差に耐える) ---
    On Error Resume Next
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    On Error GoTo 0
    On Error Resume Next
    Application.DisplayFormulaBar = False
    On Error GoTo 0
    On Error Resume Next
    Application.DisplayStatusBar = False
    On Error GoTo 0

    ws.Activate
    On Error Resume Next
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
        .DisplayWorkbookTabs = False
        .DisplayHorizontalScrollBar = False
    End With
    On Error GoTo 0

    ' --- キャンバス骨格 ---
    RemoveNexusShapes ws
    ws.Cells.Clear
    ws.Cells.Font.Name = "Yu Gothic UI"

    ' サイドバー列(A:C)を固定幅に、チャット面はD以降
    ws.Columns("A:C").ColumnWidth = 12
    ws.Columns("D:P").ColumnWidth = 14
    ws.Rows("1:400").RowHeight = 18

    DrawSidebar ws
    DrawTopbar ws
    DrawInputArea ws
    DrawFloatingActionBar ws

    ' スクロール制御: 上部(トップバー+入力+アクションバー)とサイドバー列を固定
    On Error Resume Next
    ws.Range("D8").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    ApplyTheme ws
    mChatBottom = CHAT_TOP

    Application.ScreenUpdating = True
End Sub

' ----------------------------------------------------------------------------
' AddChatBubble - チャットバブル1件を追加する。
'   role: "user" / "ai"。thinkingはAIの思考プロセス(空なら省略)。
'   アクションはバブル毎には生成しない(オーナー裁定: Shape増殖による32bit
'   メモリクラッシュ防止)。画面上部固定のフローティング・アクションバーが、
'   選択中のAIバブル(クリックでmodApp.OnSelectBubbleが記録)に対して発火する。
'   withActionsは後方互換のため残置(無視)。戻り値: バブルShape名。
' ----------------------------------------------------------------------------
Public Function AddChatBubble(ByVal role As String, ByVal bodyText As String, _
                              Optional ByVal withActions As Boolean = False, _
                              Optional ByVal thinking As String = "") As String
    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then Exit Function

    If mChatBottom < CHAT_TOP Then RecalcChatBottom ws

    Dim isUser As Boolean
    isUser = (LCase$(role) = "user")

    Dim chatL As Double, chatW As Double
    chatL = SIDEBAR_W + CHAT_LEFT_PAD
    chatW = 640
    Dim bubbleW As Double: bubbleW = chatW * BUBBLE_RATIO

    Dim topY As Double: topY = mChatBottom + BUBBLE_GAP
    Dim seq As String: seq = NextSeq(ws)

    ' --- 思考プロセス(AIのみ・小さな灰色イタリック) ---
    If (Not isUser) And LenB(Trim$(thinking)) > 0 Then
        Dim thk As Shape
        Set thk = ws.Shapes.AddShape(1, chatL + 8, topY, bubbleW - 16, 20)   ' 1=四角
        thk.Name = "nx_thk_" & seq
        thk.Line.Visible = 0
        thk.Fill.Visible = 0
        With thk.TextFrame2
            .WordWrap = -1
            .TextRange.Text = "思考プロセス: " & thinking
            .TextRange.Font.Size = 8.5
            .TextRange.Font.Italic = -1
            .AutoSize = 1   ' msoAutoSizeShapeToFitText
        End With
        topY = thk.Top + thk.Height + 4
    End If

    ' --- バブル本体(角丸四角) ---
    Dim bx As Double
    If isUser Then
        bx = chatL + chatW - bubbleW
    Else
        bx = chatL
    End If

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, bx, topY, bubbleW, 30)   ' 5=角丸四角
    shp.Name = "nx_msg_" & IIf(isUser, "u", "a") & "_" & seq
    shp.Adjustments(1) = 0.08   ' 角丸を小さめに(HTMLの12px相当)
    With shp.TextFrame2
        .WordWrap = -1
        .MarginLeft = 10: .MarginRight = 10: .MarginTop = 8: .MarginBottom = 8
        .TextRange.Text = bodyText
        .TextRange.Font.Size = 10.5
        .TextRange.ParagraphFormat.Alignment = 0   ' 左揃え
        .AutoSize = 1
    End With
    If shp.Height < 28 Then shp.Height = 28
    shp.Shadow.Visible = 0

    ' AIバブルはクリックで「選択」できる(コンテキスト・アクションの対象指定)
    If Not isUser Then shp.OnAction = "modApp.OnSelectBubble"

    PaintBubble shp, isUser
    mChatBottom = shp.Top + shp.Height

    ScrollToBottom ws
    AddChatBubble = shp.Name
End Function

' ----------------------------------------------------------------------------
' ToggleTheme - ライト/ダーク反転(状態を永続化して全体を再彩色)
' ----------------------------------------------------------------------------
Public Sub ToggleTheme()
    If CurrentTheme() = "dark" Then
        SaveTheme "light"
    Else
        SaveTheme "dark"
    End If

    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    ApplyTheme ws
    Application.ScreenUpdating = True
End Sub

' ----------------------------------------------------------------------------
' RestoreExcelUI - ネイティブUI復元(Auto_Close/緊急脱出用)
' ----------------------------------------------------------------------------
Public Sub RestoreExcelUI()
    On Error Resume Next
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    On Error GoTo 0
    On Error Resume Next
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    With ActiveWindow
        .DisplayGridlines = True
        .DisplayHeadings = True
        .DisplayWorkbookTabs = True
        .DisplayHorizontalScrollBar = True
    End With
    On Error GoTo 0
End Sub

' Phase 1の暫定アクション受け(Phase 2でmodApp=Controllerへ移管)。
Public Sub NexusActionStub()
    MsgBox "このボタンは準備中です(次のフェーズで有効になります)。", vbInformation, "Nexus Agent"
End Sub

' ----------------------------------------------------------------------------
' 公開ゲッター: 他のNexus画面(modDash等)がテーマ一貫の配色/現在テーマを
' 得るための唯一の窓口。配色定義(ThemeColor)を各画面へ複製せず一元管理する。
'   key: bg/surface/text/muted/border/primary/accent/userBubble/aiBubble/
'        sidebar/sidebarText/sidebarActive
' ----------------------------------------------------------------------------
Public Function UiColor(ByVal key As String) As Long
    UiColor = ThemeColor(key)
End Function

Public Function UiTheme() As String
    UiTheme = CurrentTheme()
End Function

' ----------------------------------------------------------------------------
' 内部: 骨格描画
' ----------------------------------------------------------------------------

Private Sub DrawSidebar(ByVal ws As Worksheet)
    Dim sb As Shape
    Set sb = ws.Shapes.AddShape(1, 0, 0, SIDEBAR_W, 760)
    sb.Name = "nx_sb_bg"
    sb.Line.Visible = 0

    Dim brand As Shape
    Set brand = ws.Shapes.AddShape(1, 0, 0, SIDEBAR_W, 44)
    brand.Name = "nx_sb_brand"
    brand.Line.Visible = 0
    brand.Fill.Visible = 0
    With brand.TextFrame2
        .TextRange.Text = ChrW(&H26A1) & " Nexus Agent"
        .TextRange.Font.Size = 15
        .TextRange.Font.Bold = -1
        .MarginLeft = 14
        .VerticalAnchor = 3
    End With

    Dim prof As Shape
    Set prof = ws.Shapes.AddShape(5, 10, 52, SIDEBAR_W - 20, 52)
    prof.Name = "nx_sb_profile"
    prof.Line.Visible = 0
    With prof.TextFrame2
        .TextRange.Text = ProfileCaption()
        .TextRange.Font.Size = 9.5
        .MarginLeft = 10: .MarginTop = 6
        .WordWrap = -1
    End With

    Dim items As Variant
    items = Array(ChrW(&H1F4AC) & " チャット", ChrW(&H1F4DA) & " ナレッジ倉庫", ChrW(&H1F4CA) & " ダッシュボード")
    Dim i As Long
    For i = 0 To 2
        Dim nav As Shape
        Set nav = ws.Shapes.AddShape(1, 0, 120 + i * 40, SIDEBAR_W, 38)
        nav.Name = "nx_sb_nav" & (i + 1)
        nav.Line.Visible = 0
        With nav.TextFrame2
            .TextRange.Text = CStr(items(i))
            .TextRange.Font.Size = 10.5
            .MarginLeft = 16
            .VerticalAnchor = 3
        End With
        nav.OnAction = Array("modApp.OnNavChat", "modApp.OnNavVault", "modApp.OnNavDash")(i)
    Next i
End Sub

Private Sub DrawTopbar(ByVal ws As Worksheet)
    Dim tb As Shape
    Set tb = ws.Shapes.AddShape(1, SIDEBAR_W, 0, 700, TOPBAR_H)
    tb.Name = "nx_top_bg"
    tb.Line.Visible = 0

    Dim lang As Shape
    Set lang = ws.Shapes.AddShape(5, SIDEBAR_W + 480, 8, 130, 28)
    lang.Name = "nx_top_lang"
    With lang.TextFrame2
        .TextRange.Text = ChrW(&H1F1EF) & ChrW(&H1F1F5) & " 日本語で回答"
        .TextRange.Font.Size = 9.5
        .TextRange.Font.Bold = -1
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    lang.OnAction = "modApp.OnLangCycle"

    Dim modeBtn As Shape
    Set modeBtn = ws.Shapes.AddShape(5, SIDEBAR_W + 320, 8, 150, 28)
    modeBtn.Name = "nx_top_mode"
    With modeBtn.TextFrame2
        .TextRange.Text = ChrW(&H1F3E2) & " 社内ナレッジ検索"
        .TextRange.Font.Size = 9.5
        .TextRange.Font.Bold = -1
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    modeBtn.OnAction = "modApp.OnToggleMode"

    Dim theme As Shape
    Set theme = ws.Shapes.AddShape(9, SIDEBAR_W + 625, 8, 28, 28)   ' 9=楕円
    theme.Name = "nx_top_theme"
    theme.Line.Visible = 0
    With theme.TextFrame2
        .TextRange.Text = ThemeIcon()
        .TextRange.Font.Size = 12
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
    End With
    theme.OnAction = "modUI.ToggleTheme"
End Sub

Private Sub DrawInputArea(ByVal ws As Worksheet)
    ' 入力はセル(名前付き範囲 nx_input)+送信ボタン。SPA下部固定は
    ' FreezePanesと相性が悪いため、Phase 1では上部(トップバー直下)に置く。
    On Error Resume Next
    ThisWorkbook.Names("nx_input").Delete
    On Error GoTo 0
    On Error Resume Next
    ThisWorkbook.Names.Add "nx_input", "='" & ws.Name & "'!$D$2"
    On Error GoTo 0

    Dim send As Shape
    Set send = ws.Shapes.AddShape(5, SIDEBAR_W + 560, TOPBAR_H + 6, 90, 30)
    send.Name = "nx_top_send"
    send.Line.Visible = 0
    With send.TextFrame2
        .TextRange.Text = "送信"
        .TextRange.Font.Size = 11
        .TextRange.Font.Bold = -1
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    send.OnAction = "modApp.OnSend"

    Dim clip As Shape
    Set clip = ws.Shapes.AddShape(9, SIDEBAR_W + 15, TOPBAR_H + 6, 30, 30)
    clip.Name = "nx_top_clip"
    clip.Line.Visible = 0
    With clip.TextFrame2
        .TextRange.Text = ChrW(&H1F4CE)
        .TextRange.Font.Size = 12
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    clip.OnAction = "modApp.OnAttachImage"
End Sub

' フローティング・アクションバー(オーナー裁定②): 画面上部固定領域に
' 6ボタンを1セットだけ常設し、選択中のAIバブルに対して発火させる。
' バブル毎のボタン生成を全廃してShape増殖(32bitメモリクラッシュ)を防ぐ。
Private Sub DrawFloatingActionBar(ByVal ws As Worksheet)
    Dim labels As Variant, widths As Variant, kinds As Variant
    labels = Array(ChrW(&H1F44D) & " グッド", ChrW(&H1F44E) & " バッド", _
                   ChrW(&H1F50D) & " 深掘り", ChrW(&H2705) & " 解決した", _
                   ChrW(&H1F198) & " 本社へ照会", ChrW(&H1F4C4) & " Word出力")
    widths = Array(70, 70, 70, 80, 92, 88)
    kinds = Array("good", "bad", "drill", "resolve", "hq", "word")
    Dim handlers As Variant
    handlers = Array("OnActGood", "OnActBad", "OnActDrill", "OnActResolve", "OnActHq", "OnActWord")

    Dim x As Double: x = SIDEBAR_W + CHAT_LEFT_PAD
    Dim topY As Double: topY = TOPBAR_H + 44
    Dim i As Long
    For i = 0 To 5
        Dim btn As Shape
        Set btn = ws.Shapes.AddShape(5, x, topY, CDbl(widths(i)), ACT_H)
        btn.Name = "nx_fab_" & CStr(kinds(i))
        btn.Adjustments(1) = 0.5
        With btn.TextFrame2
            .TextRange.Text = CStr(labels(i))
            .TextRange.Font.Size = 8.5
            .TextRange.ParagraphFormat.Alignment = 2
            .VerticalAnchor = 3
            .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
        End With
        btn.OnAction = "modApp." & CStr(handlers(i))
        PaintActionButton btn, CStr(kinds(i))
        x = x + CDbl(widths(i)) + 6
    Next i
End Sub

' ----------------------------------------------------------------------------
' 内部: テーマ(HTMLモックのCSS変数を移植)
' ----------------------------------------------------------------------------

Private Function CurrentTheme() As String
    CurrentTheme = "light"
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_UISTATE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(CStr(ws.Cells(i, 1).Value), THEME_KEY, vbTextCompare) = 0 Then
            If LCase$(Trim$(CStr(ws.Cells(i, 2).Value))) = "dark" Then CurrentTheme = "dark"
            Exit Function
        End If
    Next i
End Function

Private Sub SaveTheme(ByVal themeName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_UISTATE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim r As Long: r = 0
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(CStr(ws.Cells(i, 1).Value), THEME_KEY, vbTextCompare) = 0 Then
            r = i
            Exit For
        End If
    Next i
    If r = 0 Then
        r = lastRow + 1
        If r < 1 Then r = 1
        ws.Cells(r, 1).Value = THEME_KEY
    End If
    ws.Cells(r, 2).Value = themeName
End Sub

' CSS変数に対応する色(RGB)。key: bg/surface/text/muted/border/primary/accent/
' userBubble/aiBubble/sidebar/sidebarText/sidebarActive
Private Function ThemeColor(ByVal key As String) As Long
    Dim dark As Boolean: dark = (CurrentTheme() = "dark")
    Select Case key
        Case "bg":            ThemeColor = IIf(dark, RGB(15, 23, 42), RGB(243, 244, 246))
        Case "surface":       ThemeColor = IIf(dark, RGB(30, 41, 59), RGB(255, 255, 255))
        Case "text":          ThemeColor = IIf(dark, RGB(248, 250, 252), RGB(17, 24, 39))
        Case "muted":         ThemeColor = IIf(dark, RGB(148, 163, 184), RGB(107, 114, 128))
        Case "border":        ThemeColor = IIf(dark, RGB(51, 65, 85), RGB(229, 231, 235))
        Case "primary":       ThemeColor = IIf(dark, RGB(59, 130, 246), RGB(37, 99, 235))
        Case "accent":        ThemeColor = RGB(16, 185, 129)
        Case "userBubble":    ThemeColor = IIf(dark, RGB(51, 65, 85), RGB(239, 246, 255))
        Case "aiBubble":      ThemeColor = IIf(dark, RGB(30, 41, 59), RGB(255, 255, 255))
        Case "sidebar":       ThemeColor = IIf(dark, RGB(11, 15, 25), RGB(17, 24, 39))
        Case "sidebarText":   ThemeColor = RGB(209, 213, 219)
        Case "sidebarActive": ThemeColor = IIf(dark, RGB(30, 41, 59), RGB(31, 41, 55))
        Case Else:            ThemeColor = RGB(0, 0, 0)
    End Select
End Function

' テーマを画面全体へ適用(背景セル+全nx_Shape再彩色)。
Private Sub ApplyTheme(ByVal ws As Worksheet)
    ws.Cells.Interior.Color = ThemeColor("bg")

    Dim shp As Shape
    For Each shp In ws.Shapes
        Dim nm As String: nm = shp.Name
        If Left$(nm, 3) <> "nx_" Then GoTo NextShp

        If Left$(nm, 6) = "nx_sb_" Then
            If nm = "nx_sb_bg" Then
                shp.Fill.ForeColor.RGB = ThemeColor("sidebar")
            ElseIf nm = "nx_sb_profile" Then
                shp.Fill.ForeColor.RGB = ThemeColor("sidebarActive")
                SetShapeTextColor shp, ThemeColor("sidebarText")
            ElseIf nm = "nx_sb_nav1" Then
                shp.Fill.ForeColor.RGB = ThemeColor("sidebarActive")
                SetShapeTextColor shp, RGB(255, 255, 255)
            ElseIf Left$(nm, 9) = "nx_sb_nav" Then
                shp.Fill.Visible = 0
                SetShapeTextColor shp, ThemeColor("sidebarText")
            Else
                SetShapeTextColor shp, RGB(255, 255, 255)
            End If
        ElseIf Left$(nm, 7) = "nx_top_" Then
            If nm = "nx_top_bg" Then
                shp.Fill.ForeColor.RGB = ThemeColor("surface")
            ElseIf nm = "nx_top_send" Then
                shp.Fill.ForeColor.RGB = ThemeColor("primary")
                SetShapeTextColor shp, RGB(255, 255, 255)
            ElseIf nm = "nx_top_theme" Then
                shp.Fill.Visible = 0
                SetShapeTextColor shp, ThemeColor("text")
                shp.TextFrame2.TextRange.Text = ThemeIcon()
            Else
                shp.Fill.ForeColor.RGB = ThemeColor("bg")
                shp.Line.ForeColor.RGB = ThemeColor("border")
                SetShapeTextColor shp, ThemeColor("text")
            End If
        ElseIf Left$(nm, 9) = "nx_msg_u_" Then
            PaintBubble shp, True
        ElseIf Left$(nm, 9) = "nx_msg_a_" Then
            PaintBubble shp, False
        ElseIf Left$(nm, 7) = "nx_fab_" Then
            PaintActionButton shp, Mid$(nm, 8)
        ElseIf Left$(nm, 7) = "nx_thk_" Then
            SetShapeTextColor shp, ThemeColor("muted")
        End If
NextShp:
    Next shp
End Sub

Private Sub PaintBubble(ByVal shp As Shape, ByVal isUser As Boolean)
    If isUser Then
        shp.Fill.ForeColor.RGB = ThemeColor("userBubble")
        shp.Line.Visible = 0
    Else
        shp.Fill.ForeColor.RGB = ThemeColor("aiBubble")
        shp.Line.Visible = -1
        shp.Line.ForeColor.RGB = ThemeColor("border")
        shp.Line.Weight = 0.75
    End If
    SetShapeTextColor shp, ThemeColor("text")
End Sub

Private Sub PaintActionButton(ByVal shp As Shape, ByVal kind As String)
    shp.Fill.ForeColor.RGB = ThemeColor("surface")
    shp.Line.Visible = -1
    shp.Line.Weight = 0.75
    Select Case kind
        Case "resolve"
            shp.Line.ForeColor.RGB = RGB(16, 185, 129)
            SetShapeTextColor shp, RGB(16, 185, 129)
        Case "hq"
            shp.Line.ForeColor.RGB = RGB(239, 68, 68)
            SetShapeTextColor shp, RGB(239, 68, 68)
        Case "word"
            shp.Line.ForeColor.RGB = ThemeColor("primary")
            SetShapeTextColor shp, ThemeColor("primary")
        Case Else
            shp.Line.ForeColor.RGB = ThemeColor("border")
            SetShapeTextColor shp, ThemeColor("text")
    End Select
End Sub

Private Sub SetShapeTextColor(ByVal shp As Shape, ByVal rgbVal As Long)
    On Error Resume Next
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgbVal
    On Error GoTo 0
End Sub

Private Function ThemeIcon() As String
    If CurrentTheme() = "dark" Then
        ThemeIcon = ChrW(&H2600)    ' 太陽
    Else
        ThemeIcon = ChrW(&H1F319)   ' 月
    End If
End Function

Private Function ProfileCaption() As String
    ' Phase 4でimportAD(ADSystemInfo)連携に置換。失敗時はGuest扱い。
    Dim lv As Long, ex As Long
    On Error Resume Next
    lv = modStats.Level()
    ex = modStats.ExpTotal()
    On Error GoTo 0
    If lv < 1 Then lv = 1
    ProfileCaption = "ゲスト ユーザー" & vbLf & "Lv." & lv & " ・ EXP " & ex
End Function

' ----------------------------------------------------------------------------
' 内部: ユーティリティ
' ----------------------------------------------------------------------------

Private Function GetNexusSheet() As Worksheet
    On Error Resume Next
    Set GetNexusSheet = ThisWorkbook.Worksheets(NEXUS_SHEET)
    On Error GoTo 0
End Function

Private Function GetOrCreateNexusSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = NEXUS_SHEET
        On Error GoTo 0
    End If
    Set GetOrCreateNexusSheet = ws
    Exit Function
Fail:
    Set GetOrCreateNexusSheet = Nothing
End Function

Private Sub RemoveNexusShapes(ByVal ws As Worksheet)
    Dim names() As String
    ReDim names(0 To ws.Shapes.count)
    Dim n As Long: n = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 3) = "nx_" Then
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

' バブル連番(既存Shape数から決定的に採番)。
Private Function NextSeq(ByVal ws As Worksheet) As String
    Dim maxN As Long: maxN = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 7) = "nx_msg_" Then maxN = maxN + 1
    Next shp
    NextSeq = Format$(maxN + 1, "0000")
End Function

' モジュール状態リセット後の再計算: 既存バブルの最下端を探す。
Private Sub RecalcChatBottom(ByVal ws As Worksheet)
    mChatBottom = CHAT_TOP
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 3) = "nx_" And Left$(shp.Name, 6) <> "nx_sb_" _
           And Left$(shp.Name, 7) <> "nx_top_" And Left$(shp.Name, 7) <> "nx_fab_" Then
            If shp.Top + shp.Height > mChatBottom Then mChatBottom = shp.Top + shp.Height
        End If
    Next shp
End Sub

' 選択中バブルの強調表示(primary色の太枠)。他のAIバブルは通常枠へ戻す。
Public Sub MarkActiveBubble(ByVal shapeName As String)
    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then Exit Sub

    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 9) = "nx_msg_a_" Then
            If shp.Name = shapeName Then
                shp.Line.Visible = -1
                shp.Line.ForeColor.RGB = ThemeColor("primary")
                shp.Line.Weight = 1.75
            Else
                PaintBubble shp, False
            End If
        End If
    Next shp
End Sub

' 指定バブルの本文テキストを返す(無ければ"")。
Public Function BubbleTextOf(ByVal shapeName As String) As String
    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    BubbleTextOf = ws.Shapes(shapeName).TextFrame2.TextRange.Text
    On Error GoTo 0
End Function

' 最新(最下端)のAIバブル名を返す(無ければ"")。
Public Function LatestAiBubbleName() As String
    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then Exit Function

    Dim bestName As String: bestName = ""
    Dim bestBottom As Double: bestBottom = -1
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 9) = "nx_msg_a_" Then
            If shp.Top + shp.Height > bestBottom Then
                bestBottom = shp.Top + shp.Height
                bestName = shp.Name
            End If
        End If
    Next shp
    LatestAiBubbleName = bestName
End Function

Private Sub ScrollToBottom(ByVal ws As Worksheet)
    On Error Resume Next
    Dim targetRow As Long
    targetRow = CLng(mChatBottom / 18) + 3
    If targetRow < 4 Then targetRow = 4
    Application.GoTo ws.Cells(targetRow, 4), True
    On Error GoTo 0
End Sub
