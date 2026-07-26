Attribute VB_Name = "modUINexusDraw"
Option Explicit

' modUINexusDraw - チャット画面(Nexusシート)の骨格描画。
'
' 2026-07-26 再設計(nexus-spec-v1 §2.2 / nexus-ui-final 画面5):
'   ・サイドバーを全廃した。移動導線はHub画面(modHub)に集約し、チャット画面は
'     「会話」だけを担う。結果、バブルの表示幅が約30%広がる。
'   ・常時表示のフローティング・アクションバー(7個)を廃止し、最新のAI回答
'     バブルの直下にだけ6個のアクションpillを出す「文脈表示」に変えた。
'     質問前はボタンが0個になるので入力に集中できる。
'   ・ヘッダーは1行(← Hub / タイトル / 速度 / モード / 言語 / テーマ / クリア)。
'
' 設計の鉄則(実機で繰り返し事故った点):
'   ・Shape座標は必ず実セル幾何(Range.Left/.Width/Rows().Top)から導く。
'     pt決め打ちは列幅・行高・DPIの差で必ずズレる。
'   ・配色は modUI.UiColor() を単一情報源にする。
'   ・絵文字はChrW()で組み立てる(ソースへの直書きは自己インストーラの
'     文字列注入で化ける)。BMP外はサロゲートペアで2つ繋ぐ。

Public Const HDR_H As Double = 42          ' ヘッダー行(行1)の高さ
Public Const INPUT_ROW As Long = 3         ' 入力欄の行
Public Const ACT_H As Double = 22          ' 文脈アクションpillの高さ
Private Const HDR_BTN_H As Double = 26

' チャット領域(バブル/アクションの基準)。列幅の実測から求める。
Public Function ChatLeft(ByVal ws As Worksheet) As Double
    ChatLeft = ws.Range("B1").Left
End Function

Public Function ChatWidth(ByVal ws As Worksheet) As Double
    ChatWidth = ws.Range("B1:L1").Width
End Function

' 固定領域(行1〜4)の直下=会話の開始Y。FreezePanesの境界と必ず一致する。
Public Function ChatTop(ByVal ws As Worksheet) As Double
    ChatTop = ws.Rows(5).Top
End Function

' ヘッダー1行。左に「← Hub」、右に操作pillを右詰めで並べる。
Public Sub DrawChatHeader(ByVal ws As Worksheet)
    Dim L As Double, W As Double
    L = ws.Range("A1").Left
    W = ws.Range("A1:M1").Width

    Dim bg As Shape
    Set bg = ws.Shapes.AddShape(5, L, 0, W, HDR_H)
    bg.Name = "nx_top_bg"
    bg.Adjustments(1) = 0.02
    bg.Line.Visible = 0
    bg.Fill.ForeColor.RGB = modUI.UiColor("sidebar")
    modSkin.ApplyHeaderDepth bg           ' §9: 濃紺の2色グラデーション
    ' 「今どの部門の公式ナレッジにつないでいるか」を必ず見せる。
    ' 切替式である以上、これが見えないと「なぜ答えられないのか」が
    ' 分からなくなる(迷子の最大要因)。
    Dim chLabel As String
    On Error Resume Next
    chLabel = modChannel.ActiveLabel()
    On Error GoTo 0

    With bg.TextFrame2
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCAC) & " チャット   " & chLabel
        .TextRange.Font.Size = 12
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .MarginLeft = 82
        .VerticalAnchor = 3
    End With

    HeaderButton ws, "nx_top_back", ChrW(&H2190) & " Hub", _
                 L + 8, 62, "modApp.OnNavHome"

    ' 右端から左へ順に積む(文字数が変わっても右揃えが崩れない)。
    Dim x As Double: x = L + W - 8
    ' 🚪はタブもExcelの×ボタンも隠している構成での唯一の脱出路。Hubまで
    ' 戻らないと終われない状態にしないため、チャット側にも必ず置く。
    x = x - 30:  HeaderButton ws, "nx_top_exit", ChrW(&HD83D) & ChrW(&HDEAA), _
                              x, 30, "modApp.OnSaveAndExit"
    x = x - 6 - 62: HeaderButton ws, "nx_top_clear", ChrW(&HD83D) & ChrW(&HDDD1) & " クリア", _
                              x, 62, "modApp.OnClearChat"
    x = x - 6 - 30:  HeaderButton ws, "nx_top_help", ChrW(&H2753), x, 30, "modHelp.OnHelpClick"
    x = x - 6 - HDR_BTN_H
    HeaderTheme ws, x
    x = x - 6 - 92:  HeaderButton ws, "nx_top_lang", LangCaption(), x, 92, "modApp.OnLangCycle"
    x = x - 6 - 118: HeaderButton ws, "nx_top_mode", modApp.ModeCaption(), x, 118, "modApp.OnToggleMode"
    x = x - 6 - 124: HeaderButton ws, "nx_top_speed", modApp.SpeedCaption(), x, 124, "modApp.OnToggleSpeed"
End Sub

Private Sub HeaderButton(ByVal ws As Worksheet, ByVal shapeName As String, _
                         ByVal caption As String, ByVal x As Double, _
                         ByVal w As Double, ByVal action As String)
    On Error Resume Next
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(5, x, (HDR_H - HDR_BTN_H) / 2, w, HDR_BTN_H)
    If btn Is Nothing Then Exit Sub
    btn.Name = shapeName
    btn.Adjustments(1) = 0.35
    btn.Line.Visible = 0
    btn.Fill.ForeColor.RGB = modUI.UiColor("sidebarActive")
    With btn.TextFrame2
        .TextRange.Text = caption
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
    End With
    btn.OnAction = action
    On Error GoTo 0
End Sub

Private Sub HeaderTheme(ByVal ws As Worksheet, ByVal x As Double)
    On Error Resume Next
    Dim th As Shape
    Set th = ws.Shapes.AddShape(9, x, (HDR_H - HDR_BTN_H) / 2, HDR_BTN_H, HDR_BTN_H)
    If th Is Nothing Then Exit Sub
    th.Name = "nx_top_theme"
    th.Line.Visible = 0
    th.Fill.ForeColor.RGB = modUI.UiColor("sidebarActive")
    With th.TextFrame2
        .TextRange.Text = modUI.ThemeIcon()
        .TextRange.Font.Size = 11
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
    End With
    th.OnAction = "modUI.ToggleTheme"
    On Error GoTo 0
End Sub

Private Function LangCaption() As String
    Dim v As String
    On Error Resume Next
    v = modConfig.GetString("answer_language", "日本語")
    On Error GoTo 0
    If LenB(v) = 0 Then v = "日本語"
    LangCaption = ChrW(&HD83C) & ChrW(&HDF10) & " " & v
End Function

' 入力欄(行3)。編集できるセルはC3:K3だけに絞り、📎と送信はその外側の
' 土台セル(B列/L列)の実測幾何の中央へ置く。構造的に重ならない。
Public Sub DrawInputArea(ByVal ws As Worksheet)
    On Error Resume Next
    ThisWorkbook.Names("nx_input").Delete
    On Error GoTo 0

    Dim r As String: r = CStr(INPUT_ROW)
    With ws.Range("C" & r & ":K" & r)
        .Merge
        .Interior.Color = RGB(255, 255, 255)
        .VerticalAlignment = -4108        ' xlCenter
        .WrapText = True
        .Font.Size = 11
        .IndentLevel = 1
        .BorderAround LineStyle:=1, Weight:=2, Color:=modUI.UiColor("border")
    End With

    On Error Resume Next
    ThisWorkbook.Names.Add "nx_input", "='" & ws.Name & "'!$C$" & r
    On Error GoTo 0

    Dim rowTop As Double, rowH As Double
    rowTop = ws.Rows(INPUT_ROW).Top
    rowH = ws.Rows(INPUT_ROW).Height

    Dim cellB As Range: Set cellB = ws.Range("B" & r)
    Dim clipD As Double: clipD = 28
    Dim clip As Shape
    Set clip = ws.Shapes.AddShape(9, _
        cellB.Left + (cellB.Width - clipD) / 2, rowTop + (rowH - clipD) / 2, clipD, clipD)
    clip.Name = "nx_top_clip"
    clip.Line.Visible = 0
    clip.Fill.ForeColor.RGB = modUI.UiColor("accent")
    With clip.TextFrame2
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCCE)
        .TextRange.Font.Size = 12
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
    End With
    clip.OnAction = "modApp.OnAttachImage"

    Dim cellL As Range: Set cellL = ws.Range("L" & r)
    Dim sendH As Double: sendH = 32
    Dim send As Shape
    Set send = ws.Shapes.AddShape(5, _
        cellL.Left + 4, rowTop + (rowH - sendH) / 2, cellL.Width - 8, sendH)
    send.Name = "nx_top_send"
    send.Adjustments(1) = 0.3
    send.Line.Visible = 0
    send.Fill.ForeColor.RGB = modUI.UiColor("accent")
    With send.TextFrame2
        .TextRange.Text = ChrW(&H27A4) & " 送信"
        .TextRange.Font.Size = 11
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    send.OnAction = "modApp.OnSend"

    ' 入力欄の下に小さなヒント(セル。Shapeを増やさない)。
    With ws.Range("C" & (INPUT_ROW + 1))
        .Value = "Ctrl+Enter で送信 ・ Ctrl+Shift+Q でどこからでも呼び出し"
        .Font.Size = 8
        .Font.Color = modUI.UiColor("muted")
        .VerticalAlignment = -4160        ' xlTop
    End With
End Sub

' 文脈アクション: 最新のAI回答バブルの直下にだけ6個のpillを出す。
' 新しい質問を送るたびにClearContextActionsで消えるので、古い回答の下には残らない。
Public Sub DrawContextActions(ByVal ws As Worksheet, ByVal bubbleName As String)
    ClearContextActions ws
    If ws Is Nothing Then Exit Sub
    If LenB(bubbleName) = 0 Then Exit Sub

    Dim anchor As Shape
    On Error Resume Next
    Set anchor = ws.Shapes(bubbleName)
    On Error GoTo 0
    If anchor Is Nothing Then Exit Sub

    Dim y As Double: y = anchor.Top + anchor.Height + 6

    ' --- 信頼度バッジ ---
    ' 「この答えを信じてよいか」を利用者が自分で判断できないことが、
    ' フィードバックが集まらない根本原因だった。検索スコアという機械側の
    ' 情報を、確認すべきときだけ確認を促す一文に翻訳して先に見せる。
    Dim confText As String
    On Error Resume Next
    confText = modAsk.LastConfidenceText()
    On Error GoTo 0
    If LenB(confText) > 0 Then
        On Error Resume Next
        Dim badge As Shape
        Set badge = ws.Shapes.AddShape(5, anchor.Left, y, 330, 18)
        If Err.Number = 0 And Not badge Is Nothing Then
            badge.Name = "nx_act_conf"
            badge.Adjustments(1) = 0.4
            badge.Line.Visible = 0
            badge.Fill.Visible = 0
            With badge.TextFrame2
                .WordWrap = -1
                .TextRange.Text = confText
                .TextRange.Font.Size = 8
                .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
                .VerticalAnchor = 3
                .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
            End With
            badge.Placement = 3
            y = y + 20
        End If
        Set badge = Nothing
        Err.Clear
        On Error GoTo 0
    End If

    Dim caps As Variant, kinds As Variant, widths As Variant, acts As Variant
    ' 評価は「解決した/微妙/違う」の3択。どれも1クリックで完結し、
    ' 入力を強制しない(入力を求めた途端に誰も押さなくなる)。
    caps = Array(ChrW(&H2705) & " 解決した", _
                 ChrW(&HD83E) & ChrW(&HDD14) & " 微妙", _
                 ChrW(&H274C) & " 違う", _
                 ChrW(&HD83D) & ChrW(&HDD0D) & " 深掘り", _
                 ChrW(&HD83D) & ChrW(&HDCCB) & " コピー", _
                 ChrW(&HD83D) & ChrW(&HDCC4) & " Word", _
                 ChrW(&HD83C) & ChrW(&HDD98) & " 本社照会")
    kinds = Array("resolve", "unsure", "bad", "drill", "copy", "word", "hq")
    widths = Array(84, 64, 64, 74, 72, 68, 86)
    acts = Array("OnActResolve", "OnActUnsure", "OnActBad", "OnActDrill", _
                 "OnActCopy", "OnActWord", "OnActHq")

    Dim x As Double: x = anchor.Left
    Dim i As Long
    For i = 0 To 6
        ' 1個の1004で残りを道連れにしない(実機で繰り返した描画中断の教訓)。
        On Error Resume Next
        Dim btn As Shape
        Set btn = ws.Shapes.AddShape(5, x, y, CDbl(widths(i)), ACT_H)
        If Err.Number = 0 And Not btn Is Nothing Then
            btn.Name = "nx_act_" & CStr(kinds(i))
            btn.Adjustments(1) = 0.4
            With btn.TextFrame2
                .TextRange.Text = CStr(caps(i))
                .TextRange.Font.Size = 8.5
                .TextRange.ParagraphFormat.Alignment = 2
                .VerticalAnchor = 3
                .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
            End With
            btn.OnAction = "modApp." & CStr(acts(i))
            btn.Placement = 3
            modUI.PaintActionButton btn, CStr(kinds(i))
        End If
        Set btn = Nothing
        Err.Clear
        On Error GoTo 0
        x = x + CDbl(widths(i)) + 5
    Next i
End Sub

Public Sub ClearContextActions(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Dim names() As String
    ReDim names(0 To ws.Shapes.Count)
    Dim n As Long
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 7) = "nx_act_" Then
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

' 文脈アクションの最下端(出典チップ等をその下へ積むために使う)。
' アクションが無ければ0を返す。
Public Function ContextActionsBottom(ByVal ws As Worksheet) As Double
    If ws Is Nothing Then Exit Function
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 7) = "nx_act_" Then
            If shp.Top + shp.Height > ContextActionsBottom Then
                ContextActionsBottom = shp.Top + shp.Height
            End If
        End If
    Next shp
End Function
