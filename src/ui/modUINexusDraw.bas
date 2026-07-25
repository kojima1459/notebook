Attribute VB_Name = "modUINexusDraw"
Option Explicit

' modUINexusDraw - Nexus画面の骨格描画(サイドバー/トップバー/入力欄/
' フローティングアクションバー)。modUI.InitUIから呼ばれる。定数・共有ヘルパー
' (SIDEBAR_W等/PaintActionButton/ThemeIcon/UiColor)はmodUI側のPublicを使う。

Public Sub DrawSidebar(ByVal ws As Worksheet)
    Dim sb As Shape
    Set sb = ws.Shapes.AddShape(1, 0, 0, modUI.SIDEBAR_W, 760)
    sb.Name = "nx_sb_bg"
    sb.Line.Visible = 0

    Dim brand As Shape
    Set brand = ws.Shapes.AddShape(1, 0, 0, modUI.SIDEBAR_W, 34)
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

    ' 会話クリア/保存して終了ボタン。実機報告(2026-07-22)「ロゴと被る」対策:
    ' ブランド文字と同じ行に詰め込まず、専用の行(Y=38)へ分離する。
    Dim clearBtn As Shape
    Set clearBtn = ws.Shapes.AddShape(5, 10, 38, 83, 24)
    clearBtn.Name = "nx_sb_clear"
    clearBtn.Adjustments(1) = 0.3
    clearBtn.Line.Visible = 0
    clearBtn.Fill.ForeColor.RGB = modUI.UiColor("sidebarActive")
    With clearBtn.TextFrame2
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDDD1) & " クリア"
        .TextRange.Font.Size = 7.5
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    clearBtn.OnAction = "modApp.OnClearChat"

    Dim exitBtn As Shape
    Set exitBtn = ws.Shapes.AddShape(5, 101, 38, 83, 24)
    exitBtn.Name = "nx_sb_exit"
    exitBtn.Adjustments(1) = 0.3
    exitBtn.Line.Visible = 0
    exitBtn.Fill.ForeColor.RGB = modUI.UiColor("sidebarActive")
    With exitBtn.TextFrame2
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDEAA) & " 終了"
        .TextRange.Font.Size = 7.5
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    exitBtn.OnAction = "modApp.OnSaveAndExit"

    Dim prof As Shape
    Set prof = ws.Shapes.AddShape(5, 10, 68, modUI.SIDEBAR_W - 20, 52)
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
        ' 項目ごとにResume Next(1つの1004が後続項目を道連れにしないため)。
        On Error Resume Next
        Dim nav As Shape
        Set nav = ws.Shapes.AddShape(1, 0, 136 + i * 40, modUI.SIDEBAR_W, 38)
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
        If Left$(shp2.Name, 9) = "nx_sb_nav" Then navCount = navCount + 1
    Next shp2
    modLog.LogUsage "diag", "nexus_sidebar", "nav=" & navCount & "/6"
    On Error GoTo 0
End Sub

Public Sub DrawTopbar(ByVal ws As Worksheet)
    Dim tb As Shape
    Set tb = ws.Shapes.AddShape(1, modUI.SIDEBAR_W, 0, 700, modUI.TOPBAR_H)
    tb.Name = "nx_top_bg"
    tb.Line.Visible = 0

    ' すぐ聞く/しっかり調べる切替(ホームと同じui_state "mode"キーを共有)。
    Dim speedBtn As Shape
    Set speedBtn = ws.Shapes.AddShape(5, modUI.SIDEBAR_W + 15, 8, 150, 28)
    speedBtn.Name = "nx_top_speed"
    With speedBtn.TextFrame2
        .TextRange.Text = modApp.SpeedCaption()
        .TextRange.Font.Size = 9.5
        .TextRange.Font.Bold = -1
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    speedBtn.OnAction = "modApp.OnToggleSpeed"

    Dim lang As Shape
    Set lang = ws.Shapes.AddShape(5, modUI.SIDEBAR_W + 480, 8, 130, 28)
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
    Set modeBtn = ws.Shapes.AddShape(5, modUI.SIDEBAR_W + 320, 8, 150, 28)
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
    Set theme = ws.Shapes.AddShape(9, modUI.SIDEBAR_W + 625, 8, 28, 28)   ' 9=楕円
    theme.Name = "nx_top_theme"
    theme.Line.Visible = 0
    With theme.TextFrame2
        .TextRange.Text = modUI.ThemeIcon()
        .TextRange.Font.Size = 12
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
    End With
    theme.OnAction = "modUI.ToggleTheme"
End Sub

Public Sub DrawInputArea(ByVal ws As Worksheet)
    ' 編集中はExcelの編集オーバーレイがShapeより前面に出て隠すため、編集可能
    ' セルは中央F4:J4のみに絞る。白背景/枠線は入力セルF4:J4だけに限定する
    ' (両端D4:E4/K4:N4は無地のままクリップ/送信ボタンの土台に徹させる)。
    ' WrapText+行を高くして長文が右へあふれず折り返すようにする。
    On Error Resume Next
    ThisWorkbook.Names("nx_input").Delete
    On Error GoTo 0

    ws.Rows("4").RowHeight = 48

    With ws.Range("F4:J4")
        .Merge
        .Interior.Color = RGB(255, 255, 255)
        .VerticalAlignment = -4108   ' xlCenter
        .WrapText = True
        .BorderAround LineStyle:=1, Weight:=2, Color:=modUI.UiColor("border")
    End With

    On Error Resume Next
    ThisWorkbook.Names.Add "nx_input", "='" & ws.Name & "'!$F$4"
    On Error GoTo 0

    ' 実機報告(2026-07-22)「送信ボタンが入力欄に食い込む」対策: 列幅から
    ' ピクセル位置を推測していたのがズレの原因だった。実際のセル座標
    ' (D4:E4/K4:N4=土台セルの実測Left/Width)から動的に算出し、F4:J4の
    ' 外側に確実に収まるようにする。
    Dim baseD As Range, baseK As Range
    Set baseD = ws.Range("D4:E4")
    Set baseK = ws.Range("K4:N4")
    Dim rowTop As Double, rowH As Double
    rowTop = ws.Rows("4").Top
    rowH = ws.Rows("4").Height

    Dim send As Shape
    Dim sendW As Double, sendH As Double
    sendW = 80: sendH = 30
    Set send = ws.Shapes.AddShape(5, _
        baseK.Left + (baseK.Width - sendW) / 2, rowTop + (rowH - sendH) / 2, sendW, sendH)
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

    ' 実機報告(2026-07-22)「お化けみたいなボタン」対策: 背景色・枠線を明示せず
    ' 素の絵文字だけが浮いて見えていた。可視な円形ボタンとして描画する。
    Dim clip As Shape
    Dim clipD As Double
    clipD = 30
    Set clip = ws.Shapes.AddShape(9, _
        baseD.Left + (baseD.Width - clipD) / 2, rowTop + (rowH - clipD) / 2, clipD, clipD)
    clip.Name = "nx_top_clip"
    clip.Line.Visible = 0
    ' 実機報告(2026-07-22)「まだお化け」対策: surface色が背景と近く目立たな
    ' かった。彩度のあるaccent色の塗りつぶし+白アイコンで明確なボタンにする。
    clip.Fill.ForeColor.RGB = modUI.UiColor("accent")
    With clip.TextFrame2
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCCE)
        .TextRange.Font.Size = 12
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    clip.OnAction = "modApp.OnAttachImage"
End Sub

' フローティングアクションバー: 6ボタン常設し選択中バブルへ発火。
Public Sub DrawFloatingActionBar(ByVal ws As Worksheet)
    Dim labels As Variant, widths As Variant, kinds As Variant
    labels = Array(ChrW(&HD83D) & ChrW(&HDC4D) & " グッド", ChrW(&HD83D) & ChrW(&HDC4E) & " バッド", _
                   ChrW(&HD83D) & ChrW(&HDD0D) & " 深掘り", ChrW(&H2705) & " 解決した", _
                   ChrW(&HD83C) & ChrW(&HDD98) & " 本社へ照会", ChrW(&HD83D) & ChrW(&HDCC4) & " Word出力", _
                   ChrW(&HD83D) & ChrW(&HDCCB) & " コピー")
    widths = Array(70, 70, 70, 80, 92, 88, 76)
    kinds = Array("good", "bad", "drill", "resolve", "hq", "word", "copy")
    Dim handlers As Variant
    handlers = Array("OnActGood", "OnActBad", "OnActDrill", "OnActResolve", "OnActHq", "OnActWord", "OnActCopy")

    Dim x As Double: x = modUI.SIDEBAR_W + modUI.CHAT_LEFT_PAD
    Dim topY As Double: topY = modUI.TOPBAR_H + 62   ' 入力行(行4=48pt)拡張分を反映
    Dim i As Long
    For i = 0 To 6
        Dim btn As Shape
        Set btn = ws.Shapes.AddShape(5, x, topY, CDbl(widths(i)), modUI.ACT_H)
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
        modUI.PaintActionButton btn, CStr(kinds(i))
        x = x + CDbl(widths(i)) + 6
    Next i
End Sub

Private Function ProfileCaption() As String
    ' 実機報告(2026-07-22)「名前を入力してもゲストのまま」対策: 表示名を
    ' "ゲスト ユーザー"でハードコードしており、初回起動で入力した名前
    ' (modBoot.EnsureFirstRunがconfig pack_authorへ保存)を読んでいなかった。
    Dim lv As Long, ex As Long, nm As String
    On Error Resume Next
    lv = modStats.Level()
    ex = modStats.ExpTotal()
    nm = Trim$(modConfig.GetString("pack_author", ""))
    On Error GoTo 0
    If lv < 1 Then lv = 1
    If LenB(nm) = 0 Or nm = "名称未設定" Then nm = "ゲスト ユーザー"
    ProfileCaption = nm & vbLf & "Lv." & lv & " ・ EXP " & ex
End Function
