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
Private Const MSO_BRING_TO_FRONT As Long = 0   ' msoBringToFront(数値でLO互換)
Private Const MAX_BUBBLES As Long = 40         ' 32bitメモリ保護: 吹き出し保持上限

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

    ' 2026-07-21: modUIMain/modUIShelfと同じ理由でActivate失敗を致命的に
    ' しない(失敗しても以降の描画を試みる。診断はerr_logへ直接記録する)。
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

    ' --- 実機耐性(盲点D2/D4/C1): OneDrive自動保存の割り込み停止・ズーム基準固定・
    '     選択制限。いずれも環境差で例外になり得るのでOn Error Resume Next配下。 ---
    On Error Resume Next
    ActiveWorkbook.AutoSaveOn = False   ' D2: 自動保存がVBAへ割り込みクラッシュ/遅延するのを止める
    ActiveWindow.Zoom = 100             ' D4: Ctrl+ホイール等のズームでShape配置が崩れる基準を100%へ固定
    ws.EnableSelection = 1              ' C1: xlUnlockedCells(完全抑止はProtect併用時のみ。park運用と併せ誤選択を抑える)
    On Error GoTo 0

    ' C3: Ctrl+Z/Ctrl+Yを無効化。Shapeと隠しDBの整合が崩れるUndoを封じる
    '     (RestoreExcelUIで既定へ復元。Auto_Close経由で必ず復元される)。
    DisableUndoRedo

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
    FreezeShapePlacement ws   ' 全Shapeを絶対配置に固定(ズレ防止)
    BringFixedToFront ws      ' 固定UIを最前面へ(Z-Order維持)
    modSkin.BeautifyAll ws    ' フォント統一(Yu Gothic UI)+固定クロムに柔らかい影
    mChatBottom = CHAT_TOP

    Application.ScreenUpdating = True
    ParkFocus                 ' A2/C4: Shape選択解除+アクティブセルpark(白ハンドルを出さない)
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
        .MarginLeft = 12: .MarginRight = 12: .MarginTop = 9: .MarginBottom = 9
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
    modSkin.StyleBubble shp  ' Yu Gothic UI(バブルはフラット=影は選択時のみ)
    mChatBottom = shp.Top + shp.Height

    CapBubbles ws            ' 古い吹き出しを間引いてShape増殖(32bitクラッシュ)を防ぐ
    ScrollToBottom ws
    BringFixedToFront ws     ' 固定UI(サイドバー/トップバー/アクションバー)を最前面へ
    AddChatBubble = shp.Name
End Function

' 固定UI(サイドバー/トップバー/フローティングアクションバー)を最前面に維持する。
' 新規バブルは常に最前面へ追加されるため、描画サイクル末に必ず呼び、操作用の
' 固定要素がバブルの背面に隠れないようにする(裁定: Z-Order維持)。
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

' nx_msg_ の数が上限を超えたら、連番が小さい(古い)順に超過分を削除する。
' 対応する思考プロセス(nx_thk_)も一緒に消す。※ローカル変数に予約語 Rem を
' 使わないこと(コメント扱いされる)。
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

    ' seq昇順にバブルソート(nは高々数百)
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
    On Error Resume Next   ' 再彩色が中断しても必ず暗転解除へ到達させる
    ApplyTheme ws
    On Error GoTo 0
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
    ' C3: Ctrl+Z/Ctrl+Yの無効化を解除(引数省略=Excel既定の動作へ戻す)。
    Application.OnKey "^z"
    Application.OnKey "^y"
    ' D2/D4: 砂時計/ステータスバーも念のため既定へ(緊急脱出時の後始末)。
    Application.Cursor = -4143   ' xlDefault
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' C3: Ctrl+Z(Undo)/Ctrl+Y(Redo)を無効化する。空文字""を渡すと「そのキーを無視」。
' Nexus表示中はShapeと隠しシート(DB)の整合をUndoが壊すため封じる。復元は
' RestoreExcelUI(Auto_Close時に必ず呼ばれる)が担う。
Private Sub DisableUndoRedo()
    On Error Resume Next
    Application.OnKey "^z", ""
    Application.OnKey "^y", ""
    On Error GoTo 0
End Sub

' A2/C4: Shape選択解除+アクティブセルpark(スクロール崩壊防止)。Leaveから必ず呼ばれる。
Public Sub ParkFocus()
    On Error Resume Next
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub   ' 別ブックの選択状態を汚さない
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    If ws.Name = NEXUS_SHEET Then
        ws.Range("D2").Select   ' 入力セル(nx_input)へpark=Shape解除+次の入力に即備える
    Else
        ws.Range("A1").Select   ' Vault/Dashboard等は左上(固定領域)へpark
    End If
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

' 盲点B2/C5: Nexusチャット画面を「壊さずに」再描画する。Shape(会話履歴)は削除せず、
' ネイティブUI隠蔽・ズーム100%・絶対配置固定・Z-Order・テーマ・フォーカスparkだけを
' 再適用し、リサイズ/Alt+Tab復帰/マルチモニタ移動で生じたゴーストやズレを解消する。
' 手動リフレッシュボタン(modApp.OnRefreshUI)から呼ばれる。
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

' FreezeShapePlacement - 全Shapeをセル非依存の絶対配置(xlFreeFloating=3)へ固定し、
'   列幅差・スクロール・処理負荷でのズレを防ぐ。各画面の描画終端から呼ぶ。
Public Sub FreezeShapePlacement(ByVal ws As Worksheet)
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Placement = 3   ' xlFreeFloating
    Next shp
    On Error GoTo 0
End Sub

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
    items = Array(ChrW(&HD83D) & ChrW(&HDCAC) & " チャット", ChrW(&HD83D) & ChrW(&HDCDA) & " ナレッジ倉庫", _
                  ChrW(&HD83D) & ChrW(&HDCCA) & " ダッシュボード", ChrW(&HD83D) & ChrW(&HDD04) & " 画面を再描画")
    Dim navActions As Variant
    navActions = Array("modApp.OnNavChat", "modApp.OnNavVault", "modApp.OnNavDash", "modApp.OnRefreshUI")
    Dim i As Long
    For i = 0 To 3
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
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCCE)
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

' 配色解決はmodSkin.ResolveColorへ委譲(スキン=きせかえ対応の単一実装。
' 未解放スキンはResolveColor側でlightへ強制されるためチート不能)。
Private Function ThemeColor(ByVal key As String) As Long
    ThemeColor = modSkin.ResolveColor(key, CurrentTheme())
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
        ThemeIcon = ChrW(&HD83C) & ChrW(&HDF19)   ' 月
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
