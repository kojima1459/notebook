Attribute VB_Name = "modUI"
Option Explicit

' modUI - Nexus Agent UIコア(Windows版Excel専用・SPA風UI)。ボタンはShape+
' OnAction。テーマはui_state(key=nexus_theme)。Shape命名はnx_接頭辞で分類。

Private Const NEXUS_SHEET As String = "Nexus"
Private Const THEME_KEY As String = "nexus_theme"
Private Const MSO_BRING_TO_FRONT As Long = 0   ' msoBringToFront(数値でLO互換)
Private Const MAX_BUBBLES As Long = 40         ' 32bitメモリ保護: 吹き出し保持上限

' レイアウト(pt)
Private Const SIDEBAR_W As Double = 195      ' 260px相当
Private Const TOPBAR_H As Double = 45
Private Const CHAT_LEFT_PAD As Double = 15
Private Const BUBBLE_RATIO As Double = 0.62  ' チャット幅に対するバブル最大幅
Private Const BUBBLE_GAP As Double = 14
Private Const ACT_H As Double = 24
Private Const CHAT_TOP As Double = 130   ' 固定領域(トップバー+入力+アクションバー)の直下

Private mChatBottom As Double   ' 最後のバブルの下端(モジュール状態リセット時はRecalc)

' InitUI - ネイティブUI隠蔽+Nexus骨格描画
Public Sub InitUI()
    Dim ws As Worksheet
    Set ws = GetOrCreateNexusSheet()
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    On Error Resume Next
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    On Error GoTo 0
    On Error Resume Next
    Application.DisplayFormulaBar = False
    On Error GoTo 0
    On Error Resume Next
    Application.DisplayStatusBar = False
    On Error GoTo 0

    On Error Resume Next
    ws.Activate
    If Err.Number <> 0 Then
        modLog.LogError "E0801", "modApp.LaunchNexus", _
            "InitUI [ws.Activate] ws.Visible=" & ws.Visible & _
            " ActiveSheet=" & ThisWorkbook.ActiveSheet.Name & _
            " AppWin=" & Application.Windows.Count & " WbWin=" & ThisWorkbook.Windows.Count, Err.Number
        Err.Clear
    End If
    On Error GoTo 0
    On Error Resume Next
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
        .DisplayWorkbookTabs = False
        .DisplayHorizontalScrollBar = False
    End With
    On Error GoTo 0

    On Error Resume Next
    ActiveWorkbook.AutoSaveOn = False   ' D2: 自動保存がVBAへ割り込みクラッシュ/遅延するのを止める
    ActiveWindow.Zoom = 100             ' D4: Ctrl+ホイール等のズームでShape配置が崩れる基準を100%へ固定
    ws.EnableSelection = 1              ' C1: xlUnlockedCells(完全抑止はProtect併用時のみ。park運用と併せ誤選択を抑える)
    On Error GoTo 0

    ' Undoを封じる(Shapeと隠しDBの整合が崩れるため)。
    DisableUndoRedo

    ' --- キャンバス骨格 ---
    RemoveNexusShapes ws
    ws.Cells.Clear
    ws.Cells.Font.Name = "Yu Gothic UI"

    ws.Columns("A:C").ColumnWidth = 12
    ws.Columns("D:P").ColumnWidth = 14
    ws.Rows("1:400").RowHeight = 18

    On Error Resume Next
    DrawSidebar ws
    If Err.Number <> 0 Then LogDrawStageError "DrawSidebar", ws: Err.Clear
    On Error GoTo 0

    On Error Resume Next
    DrawTopbar ws
    If Err.Number <> 0 Then LogDrawStageError "DrawTopbar", ws: Err.Clear
    On Error GoTo 0

    On Error Resume Next
    DrawInputArea ws
    If Err.Number <> 0 Then LogDrawStageError "DrawInputArea", ws: Err.Clear
    On Error GoTo 0

    On Error Resume Next
    DrawFloatingActionBar ws
    If Err.Number <> 0 Then LogDrawStageError "DrawFloatingActionBar", ws: Err.Clear
    On Error GoTo 0

    On Error Resume Next
    ws.Range("D8").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    ApplyTheme ws
    FreezeShapePlacement ws   ' 全Shapeを絶対配置に固定(ズレ防止)
    BringFixedToFront ws      ' 固定UIを最前面へ(Z-Order維持)
    modSkin.BeautifyAll ws    ' フォント統一(Yu Gothic UI)+固定クロムに柔らかい影
    mChatBottom = CHAT_TOP

    Application.ScreenUpdating = True
    ParkFocus                 ' A2/C4: Shape選択解除+アクティブセルpark(白ハンドルを出さない)
End Sub

' AddChatBubble - チャットバブル1件を追加(role:user/ai。戻り値=バブルShape名)。
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

    ' --- 思考プロセス(AI限定) ---
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
        .MarginLeft = 12: .MarginRight = 12: .MarginTop = 9: .MarginBottom = 9
        .TextRange.Text = bodyText
        .TextRange.Font.Size = 10.5
        .TextRange.ParagraphFormat.Alignment = 1   ' msoAlignLeft(左揃え。0は無効値)
        .AutoSize = 1
    End With
    If shp.Height < 28 Then shp.Height = 28
    shp.Shadow.Visible = 0

    ' AIバブルはクリックで選択可(アクション対象指定)
    If Not isUser Then shp.OnAction = "modApp.OnSelectBubble"

    PaintBubble shp, isUser
    modSkin.StyleBubble shp  ' Yu Gothic UI(バブルはフラット=影は選択時のみ)
    mChatBottom = shp.Top + shp.Height

    CapBubbles ws            ' 古い吹き出しを間引いてShape増殖(32bitクラッシュ)を防ぐ
    ScrollToBottom ws
    BringFixedToFront ws     ' 固定UI(サイドバー/トップバー/アクションバー)を最前面へ
    AddChatBubble = shp.Name
End Function

' 固定UIを最前面に維持(描画末に必ず呼ぶ)。
Public Sub BringFixedToFront(ByVal ws As Worksheet)
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        Dim nm As String: nm = shp.Name
        If Left$(nm, 6) = "nx_sb_" Or Left$(nm, 7) = "nx_top_" Or Left$(nm, 7) = "nx_fab_" Then
            shp.ZOrder MSO_BRING_TO_FRONT
        End If
    Next shp
    On Error GoTo 0
End Sub

' nx_msg_が上限超過なら古い順に削除(nx_thk_も一緒に)。
Private Sub CapBubbles(ByVal ws As Worksheet)
    Dim names() As String, seqs() As Long
    ReDim names(0 To 255)
    ReDim seqs(0 To 255)
    Dim n As Long: n = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 7) = "nx_msg_" Then
            If n > UBound(names) Then
                ReDim Preserve names(0 To UBound(names) + 256)
                ReDim Preserve seqs(0 To UBound(seqs) + 256)
            End If
            names(n) = shp.Name
            seqs(n) = CLng(Val(Right$(shp.Name, 4)))
            n = n + 1
        End If
    Next shp
    If n <= MAX_BUBBLES Then Exit Sub

    Dim toRemove As Long: toRemove = n - MAX_BUBBLES

    ' seq昇順バブルソート(nは高々数百)
    Dim a As Long, b As Long
    For a = 0 To n - 2
        For b = 0 To n - 2 - a
            If seqs(b) > seqs(b + 1) Then
                Dim ts As Long: ts = seqs(b): seqs(b) = seqs(b + 1): seqs(b + 1) = ts
                Dim tn As String: tn = names(b): names(b) = names(b + 1): names(b + 1) = tn
            End If
        Next b
    Next a

    Dim ri As Long
    For ri = 0 To toRemove - 1
        On Error Resume Next
        ws.Shapes(names(ri)).Delete
        ws.Shapes("nx_thk_" & Format$(seqs(ri), "0000")).Delete
        On Error GoTo 0
    Next ri
End Sub

' ToggleTheme - ライト/ダーク反転+全体再彩色
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
    On Error Resume Next   ' 再彩色が中断しても必ず暗転解除へ到達させる
    ApplyTheme ws
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

' RestoreExcelUI - ネイティブUI復元
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
    ' Ctrl+Z/Ctrl+Yの無効化を解除(引数省略=既定へ戻す)。
    Application.OnKey "^z"
    Application.OnKey "^y"
    ' 砂時計/ステータスバーも念のため既定へ。
    Application.Cursor = -4143   ' xlDefault
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' GoToNexus - 「戻る」共通口。Activate失敗時はRestoreExcelUIで脱出路を残す。
Public Sub GoToNexus(ByVal source As String)
    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Activate
    If Err.Number <> 0 Then
        modLog.LogError "E0801", source, _
            "GoToNexus [ws.Activate失敗→ネイティブタブ復元で脱出路確保] ws.Visible=" & ws.Visible & _
            " ActiveSheet=" & ThisWorkbook.ActiveSheet.Name & _
            " AppWin=" & Application.Windows.Count & " WbWin=" & ThisWorkbook.Windows.Count, Err.Number
        Err.Clear
        RestoreExcelUI
    End If
    On Error GoTo 0
End Sub

' GoToNativeSheet - GoToNexus同様の脱出路付き遷移をホーム/マイ本棚等にも提供。
Public Sub GoToNativeSheet(ByVal sheetName As String, ByVal source As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Activate
    If Err.Number <> 0 Then
        modLog.LogError "E0801", source, _
            "GoToNativeSheet(" & sheetName & ") [ws.Activate失敗→ネイティブタブ復元で脱出路確保] ws.Visible=" & ws.Visible & _
            " ActiveSheet=" & ThisWorkbook.ActiveSheet.Name & _
            " AppWin=" & Application.Windows.Count & " WbWin=" & ThisWorkbook.Windows.Count, Err.Number
        Err.Clear
        RestoreExcelUI
    End If
    On Error GoTo 0
End Sub

' InitUIの各Draw*段の失敗記録役。
Private Sub LogDrawStageError(ByVal stageName As String, ByVal ws As Worksheet)
    Dim n As Long: n = Err.Number
    On Error Resume Next
    modLog.LogError "E0801", "modApp.LaunchNexus", _
        stageName & " [段階失敗・以降を継続] ActiveSheet=" & ThisWorkbook.ActiveSheet.Name, n
    On Error GoTo 0
End Sub

' Ctrl+Z/Ctrl+Y無効化(Shapeと隠しDBの整合をUndoが壊すため。復元はRestoreExcelUI)。
Private Sub DisableUndoRedo()
    On Error Resume Next
    Application.OnKey "^z", ""
    Application.OnKey "^y", ""
    On Error GoTo 0
End Sub

' Shape選択解除+アクティブセルpark(スクロール崩壊防止)。
Public Sub ParkFocus()
    On Error Resume Next
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub   ' 別ブックの選択状態を汚さない
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    If ws.Name = NEXUS_SHEET Then
        ws.Range("F4").Select   ' 入力セル(nx_input)へpark=Shape解除+次の入力に即備える
    Else
        ws.Range("A1").Select   ' Vault/Dashboard等は左上(固定領域)へpark
    End If
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

' 会話履歴を消さずにNexus画面を再描画(手動リフレッシュ用)。
Public Sub Repaint()
    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    On Error Resume Next
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    ActiveWorkbook.AutoSaveOn = False
    If ActiveSheet Is ws Then
        With ActiveWindow
            .DisplayGridlines = False
            .DisplayHeadings = False
            .DisplayWorkbookTabs = False
            .DisplayHorizontalScrollBar = False
            .Zoom = 100
        End With
    End If
    On Error GoTo 0

    ApplyTheme ws            ' 全nx_Shapeを再彩色(ゴースト=前画面の残像を塗り直す)
    FreezeShapePlacement ws  ' 絶対配置に再固定
    BringFixedToFront ws     ' 固定UIを最前面へ
    modSkin.BeautifyAll ws   ' フォント統一+固定クロムに柔らかい影
    Application.ScreenUpdating = True
    ParkFocus
End Sub

' Phase1暫定(Phase2でController移管)。
Public Sub NexusActionStub()
    MsgBox "このボタンは準備中です(次のフェーズで有効になります)。", vbInformation, "Nexus Agent"
End Sub

' 公開ゲッター: 他のNexus画面がテーマ一貫の配色/現在テーマを得る窓口。
Public Function UiColor(ByVal key As String) As Long
    UiColor = ThemeColor(key)
End Function

Public Function UiTheme() As String
    UiTheme = CurrentTheme()
End Function

' FreezeShapePlacement - 全Shapeを絶対配置固定しズレを防ぐ。
Public Sub FreezeShapePlacement(ByVal ws As Worksheet)
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Placement = 3   ' xlFreeFloating
    Next shp
    On Error GoTo 0
End Sub

'描画

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

    ' タブ非表示中でもホーム/マイ本棚へ行けるようサイドバーに直接導線を追加。
    Dim items As Variant
    items = Array(ChrW(&HD83D) & ChrW(&HDCAC) & " チャット", ChrW(&HD83C) & ChrW(&HDFE0) & " ホーム", _
                  ChrW(&HD83D) & ChrW(&HDCD6) & " マイ本棚", ChrW(&HD83D) & ChrW(&HDCDA) & " ナレッジ倉庫", _
                  ChrW(&HD83D) & ChrW(&HDCCA) & " ダッシュボード", ChrW(&HD83D) & ChrW(&HDD04) & " 画面を再描画")
    Dim navActions As Variant
    navActions = Array("modApp.OnNavChat", "modApp.OnNavHome", "modApp.OnNavShelf", _
                       "modApp.OnNavVault", "modApp.OnNavDash", "modApp.OnRefreshUI")
    Dim i As Long
    For i = 0 To 5
        On Error Resume Next
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
        nav.OnAction = CStr(navActions(i))
        If Err.Number <> 0 Then
            modLog.LogError "E0801", "modApp.LaunchNexus", _
                "DrawSidebar[nav" & (i + 1) & "]", Err.Number
            Err.Clear
        End If
        On Error GoTo 0
    Next i

    ' 診断用: 実際の生成数をusage_logへ残す。
    On Error Resume Next
    Dim navCount As Long, shp2 As Shape
    For Each shp2 In ws.Shapes
        If Left$(shp2.Name, 10) = "nx_sb_nav" Then navCount = navCount + 1
    Next shp2
    modLog.LogUsage "diag", "nexus_sidebar", "nav=" & navCount & "/6"
    On Error GoTo 0
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
        .TextRange.Text = ChrW(&HD83C) & ChrW(&HDDEF) & ChrW(&HD83C) & ChrW(&HDDF5) & " 日本語で回答"
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
        .TextRange.Text = ChrW(&HD83C) & ChrW(&HDFE2) & " 社内ナレッジ検索"
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
    ' 編集中はExcelの編集オーバーレイがShapeより前面に出て隠すため、編集可能
    ' セルは中央F4:J4のみに絞り、両端D4:E4/K4:N4は非編集の白背景土台にする。
    On Error Resume Next
    ThisWorkbook.Names("nx_input").Delete
    On Error GoTo 0

    With ws.Range("D4:N4")
        .Interior.Color = RGB(255, 255, 255)
        .VerticalAlignment = -4108   ' xlCenter
        .BorderAround LineStyle:=1, Weight:=2, Color:=modUI.UiColor("border")
    End With
    ws.Rows("4").RowHeight = 30

    With ws.Range("F4:J4")
        .Merge
        .VerticalAlignment = -4108   ' xlCenter
    End With

    On Error Resume Next
    ThisWorkbook.Names.Add "nx_input", "='" & ws.Name & "'!$F$4"
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
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCCE)
        .TextRange.Font.Size = 12
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    clip.OnAction = "modApp.OnAttachImage"
End Sub

' フローティングアクションバー: 6ボタン常設し選択中バブルへ発火。
Private Sub DrawFloatingActionBar(ByVal ws As Worksheet)
    Dim labels As Variant, widths As Variant, kinds As Variant
    labels = Array(ChrW(&HD83D) & ChrW(&HDC4D) & " グッド", ChrW(&HD83D) & ChrW(&HDC4E) & " バッド", _
                   ChrW(&HD83D) & ChrW(&HDD0D) & " 深掘り", ChrW(&H2705) & " 解決した", _
                   ChrW(&HD83C) & ChrW(&HDD98) & " 本社へ照会", ChrW(&HD83D) & ChrW(&HDCC4) & " Word出力", _
                   ChrW(&HD83D) & ChrW(&HDCCB) & " コピー")
    widths = Array(70, 70, 70, 80, 92, 88, 76)
    kinds = Array("good", "bad", "drill", "resolve", "hq", "word", "copy")
    Dim handlers As Variant
    handlers = Array("OnActGood", "OnActBad", "OnActDrill", "OnActResolve", "OnActHq", "OnActWord", "OnActCopy")

    Dim x As Double: x = SIDEBAR_W + CHAT_LEFT_PAD
    Dim topY As Double: topY = TOPBAR_H + 44
    Dim i As Long
    For i = 0 To 6
        Dim btn As Shape
        Set btn = ws.Shapes.AddShape(5, x, topY, CDbl(widths(i)), ACT_H)
        btn.Name = "nx_fab_" & CStr(kinds(i))
        btn.Adjustments(1) = 0.28   ' 浅めの角丸(Webアプリ風のシャープでモダンな角)
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

'テーマ

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
            Dim v As String
            v = LCase$(Trim$(CStr(ws.Cells(i, 2).Value)))
            If LenB(v) > 0 Then CurrentTheme = v   ' スキン名も可(検証はResolveColor側)
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

' 配色解決はmodSkin.ResolveColorへ委譲。
Private Function ThemeColor(ByVal key As String) As Long
    ThemeColor = modSkin.ResolveColor(key, CurrentTheme())
End Function

' テーマ適用(背景+全nx_Shape再彩色)。
Private Sub ApplyTheme(ByVal ws As Worksheet)
    ws.Cells.Interior.Color = ThemeColor("bg")

    ' 入力欄(F4:J4/土台D4:N4)は上の一括塗りで消えるため塗り直す。
    On Error Resume Next
    With ws.Range("D4:N4")
        .Interior.Color = RGB(255, 255, 255)
        .BorderAround LineStyle:=1, Weight:=2, Color:=ThemeColor("border")
    End With
    On Error GoTo 0

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
        ThemeIcon = ChrW(&HD83C) & ChrW(&HDF19)   ' 月
    End If
End Function

Private Function ProfileCaption() As String
    ' Phase 4でAD連携に置換予定。失敗時はGuest扱い。
    Dim lv As Long, ex As Long
    On Error Resume Next
    lv = modStats.Level()
    ex = modStats.ExpTotal()
    On Error GoTo 0
    If lv < 1 Then lv = 1
    ProfileCaption = "ゲスト ユーザー" & vbLf & "Lv." & lv & " ・ EXP " & ex
End Function

'util

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
    ' Name代入失敗時、既定名のまま孤児化するのを防ぐ。
    If Not ws Is Nothing Then
        On Error Resume Next
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        On Error GoTo 0
    End If
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

' バブル連番(既存Shape数から採番)。
Private Function NextSeq(ByVal ws As Worksheet) As String
    ' 個数ではなく既存の最大seqを見る(CapBubbles後は個数が減り、個数+1だと
    ' 既存Shapeと同名衝突する実バグだった)。
    Dim maxN As Long: maxN = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 7) = "nx_msg_" Then
            Dim n As Long: n = CLng(Val(Right$(shp.Name, 4)))
            If n > maxN Then maxN = n
        End If
    Next shp
    NextSeq = Format$(maxN + 1, "0000")
End Function

' 状態リセット後の再計算: 既存バブルの最下端を探す。
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

' 選択中バブルを強調表示(primary色の太枠)。他は通常枠へ戻す。
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
                modSkin.ApplySoftShadow shp   ' Active State: 選択中バブルだけ浮遊させる
            Else
                PaintBubble shp, False
                On Error Resume Next
                shp.Shadow.Visible = 0        ' 非選択はフラットへ戻す
                On Error GoTo 0
            End If
        End If
    Next shp
End Sub

' 指定バブルの本文を返す(無ければ"")。
Public Function BubbleTextOf(ByVal shapeName As String) As String
    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    BubbleTextOf = ws.Shapes(shapeName).TextFrame2.TextRange.Text
    On Error GoTo 0
End Function

' 最新AIバブル名を返す(無ければ"")。
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
    ' GoToでアクティブセルがD列へ動くため、生成中の入力ずれ防止にF4へ戻す
    ' (F4は固定領域内なのでスクロール位置は崩れない)。
    ws.Range("F4").Select
    On Error GoTo 0
End Sub
