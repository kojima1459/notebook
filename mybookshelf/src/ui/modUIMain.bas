Attribute VB_Name = "modUIMain"
Option Explicit

' ============================================================================
' modUIMain - 「ホーム」画面(質問UI)の構築と描画(MASTER_SPEC §7.6/§8.1)
' ----------------------------------------------------------------------------
' 役割:
'   非エンジニアが説明書なしで使える「質問する画面」。EnsureLayoutが冪等に
'   Shape(ボタン)とセルレイアウトを再構築し、SetStage/RenderAnswer/
'   RenderSourcesPreviewが実行中の進捗と結果を表示する。modAskはこのモジュール
'   の名前付き範囲 "mb_question"(質問入力セル)を読み、SetStage/
'   RenderSourcesPreview/RenderAnswerを呼び返す(§7.3 modAskの設計コメント
'   のとおり)。
'
' 設計判断:
'   ・ラベルは原則セル値(V2 modChatUI.bas踏襲)とし、実際にクリックされる
'     ボタンだけをShape(msoShapeRoundedRectangle)にする。Shape名は
'     btn_/lbl_ プレフィクス(§12)。EnsureLayoutは既存の自前Shapesを
'     全削除してから再生成する(冪等。ユーザーが壊しても開き直せば直る)。
'   ・質問入力セルは "mb_question" という定義済み名前(Excel Name)を
'     A6セル(結合セルB6:H9ではなく左上のA6セルそのもの。結合セルの値は
'     左上セルに格納されるVBAの仕様を利用)に張る。modAsk.bas側の
'     アサンプションコメントと一致させること(Wave3で突き合わせ済み)。
'     プレースホルダ文字列は使わない(空のまま)。理由: プレースホルダ文字列
'     を入れると、ユーザーが消し忘れた場合にmodAsk側がそれをそのまま質問文
'     として送ってしまう(modAskはこのモジュールのプレースホルダ規約を
'     知らないため、二重管理を避けて「空セル=未入力」に統一した)。
'   ・SetStageは「セルへの書込み」「Application.StatusBarへの反映」を
'     必ず両方行う(§7.6契約)。処理中(msgが空でない)ときは待ち時間豆知識
'     ShowTipも合わせて更新し、ユーザーが手持ち無沙汰にならないようにする。
'   ・opt機能ボタン(読み上げ)はmodFeatures.FeatureEnabledがTrueのときだけ
'     EnsureLayout内で生成する(§7.7)。ボタンのOnActionはこのモジュール内の
'     OnTtsButtonラッパー経由でmodFeatures.InvokeFeatureを呼ぶ(opt直接
'     参照はR2違反になるため、このモジュールにoptTts等のトークンは一切
'     書かない)。
'   ・本棚が空のときの案内(§8.1「まず『マイ本棚』タブで資料を1つ追加して
'     みましょう →」)はShowEmptyShelfHintとして公開し、呼び出し判断(本棚が
'     空かどうか)はmodBoot側が行う(modUIMainはmodShelfの状態を勝手に
'     判定しない。UIモジュールは「表示のしかた」の責務に留める)。
'   ・OnAskButtonは質問実行後、EvaluateBadges/RenderDashboardの更新を
'     On Error Resume Next(1行スコープ)で試みる。ダッシュボードの更新に
'     失敗しても質問応答自体は成功として扱う(非致命)。
'   ・On Error Resume Next は §12の規約どおり必ず1行スコープ(直後に
'     On Error GoTo 0)で使う。複数行にまたがる保護が必要な場面では
'     個別のヘルパーSubに切り出して1行スコープを保つ。
' ============================================================================

Private Const RNG_STATUS As String = "A12:H12"
Private Const RNG_ANSWER As String = "A14:H25"
Private Const RNG_SOURCES As String = "A26:H27"
Private Const RNG_TIP As String = "A33:H34"
Private Const QUESTION_TOP_LEFT As String = "A6"
Private Const QUESTION_RANGE_NAME As String = "mb_question"

Private Const COLOR_SELECTED_BG As Long = 2039071    ' RGB(31,78,120) 濃紺
Private Const COLOR_SELECTED_FG As Long = 16777215   ' RGB(255,255,255) 白
Private Const COLOR_UNSELECTED_BG As Long = 15921906 ' RGB(242,242,242) 薄灰
Private Const COLOR_UNSELECTED_FG As Long = 0        ' RGB(0,0,0) 黒

Private mLastAnswerText As String

' ----------------------------------------------------------------------------
' EnsureLayout - ホームを冪等再構築(既存Shapes全削除→再生成)
' ----------------------------------------------------------------------------
Public Sub EnsureLayout()
    Dim ws As Worksheet
    Set ws = GetOrCreateHomeSheet()
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    RemoveManagedShapes ws
    ws.Cells.Clear

    ws.Cells.Font.Name = "游ゴシック"
    ws.Cells.Font.Size = 11

    ws.Columns("A:H").ColumnWidth = 12

    ' ---- タイトル行 ----------------------------------------------------
    With ws.Range("A1:E1")
        .Merge
        .Value = "📚 " & modAppDef.APP_NAME & "  v" & modAppDef.APP_VERSION
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = COLOR_SELECTED_BG
        .Font.Color = COLOR_SELECTED_FG
        .HorizontalAlignment = -4108   ' xlCenter
        .VerticalAlignment = -4108     ' xlCenter
    End With
    ws.Rows("1:2").RowHeight = 22

    With ws.Range("A2:E2")
        .Merge
        .Value = "「自分で入れた資料に、すぐ聞ける」"
        .HorizontalAlignment = -4108
        .Font.Size = 10
        .Font.Italic = True
    End With

    AddButton ws, ws.Range("F1:G2"), "btn_howto", "❓ 使い方", "modUIMain.OnOpenHowto"
    AddButton ws, ws.Range("H1:H2"), "btn_diag", "🩺 診断", "modUIMain.OnRunDiag"

    ' ---- モードトグル ----------------------------------------------------
    ws.Rows("3:4").RowHeight = 20
    AddButton ws, ws.Range("A3:D4"), "btn_mode_quick", ModeCaption("quick"), "modUIMain.OnModeQuick"
    AddButton ws, ws.Range("E3:H4"), "btn_mode_deep", ModeCaption("deep"), "modUIMain.OnModeDeep"

    ' ---- 質問入力 ----------------------------------------------------
    With ws.Range("A5:H5")
        .Merge
        .Value = "質問をここに入力してください(例: 〇〇の手続きに必要な書類は?)"
        .Font.Size = 10
        .Font.Bold = True
    End With
    ws.Rows("5").RowHeight = 18

    With ws.Range("A6:H9")
        .Merge
        .WrapText = True
        .VerticalAlignment = -4160   ' xlTop
        .Interior.Color = 16513526   ' 薄い黄色
        .Borders.LineStyle = 1
    End With
    ws.Rows("6:9").RowHeight = 20
    EnsureQuestionName ws

    ' ---- 質問するボタン ----------------------------------------------------
    ws.Rows("10:11").RowHeight = 20
    AddButton ws, ws.Range("C10:F11"), "btn_ask", "💬  質 問 す る", "modUIMain.OnAskButton"

    ' ---- 状態表示 ----------------------------------------------------
    With ws.Range(RNG_STATUS)
        .Merge
        .Value = "状態: 準備できています"
        .Font.Size = 9
        .Font.Color = 8421504   ' RGB(128,128,128)
    End With
    ws.Rows("12").RowHeight = 16

    With ws.Range("A13:H13")
        .Merge
        .Value = "─ 回答 ───────────────────────────"
        .Font.Bold = True
        .Font.Size = 10
    End With
    ws.Rows("13").RowHeight = 16

    ' ---- 回答本文 ----------------------------------------------------
    With ws.Range(RNG_ANSWER)
        .Merge
        .Value = "（質問を入力して「💬 質問する」を押してください）"
        .WrapText = True
        .VerticalAlignment = -4160
        .Interior.Color = 15987946   ' 薄い水色
        .Borders.LineStyle = 1
        .Font.Size = 10
    End With
    ws.Rows("14:25").RowHeight = 15

    With ws.Range(RNG_SOURCES)
        .Merge
        .Value = ""
        .WrapText = True
        .VerticalAlignment = -4160
        .Font.Size = 9
        .Font.Color = 4210752   ' 濃い灰
    End With
    ws.Rows("26:27").RowHeight = 15

    ' ---- フィードバック ----------------------------------------------------
    With ws.Range("A28:H28")
        .Merge
        .Value = "この回答は役に立ちましたか?"
        .Font.Size = 10
    End With
    ws.Rows("28").RowHeight = 16

    ws.Rows("29:30").RowHeight = 18
    AddButton ws, ws.Range("A29:C30"), "btn_fb_green", "🟢 解決した!", "modAsk.FeedbackGreen"
    AddButton ws, ws.Range("D29:E30"), "btn_fb_yellow", "🟡 ヒントになった", "modAsk.FeedbackYellow"
    AddButton ws, ws.Range("F29:H30"), "btn_fb_red", "🔴 だめだった", "modAsk.FeedbackRed"

    ' ---- opt機能: 読み上げ(有効な時だけ生成。§7.7) ----------------------------
    If modFeatures.FeatureEnabled("tts") Then
        ws.Rows("31:32").RowHeight = 18
        AddButton ws, ws.Range("A31:D32"), "btn_tts", "🔊 読み上げる", "modUIMain.OnTtsButton"
    End If

    ' ---- 待ち時間豆知識 ----------------------------------------------------
    With ws.Range(RNG_TIP)
        .Merge
        .Value = "💡 豆知識: "
        .WrapText = True
        .VerticalAlignment = -4160
        .Font.Size = 9
        .Font.Color = 8421504
    End With
    ws.Rows("33:34").RowHeight = 15

    ApplyModeColors CurrentMode()
    ShowTip

    Application.ScreenUpdating = True
End Sub

' ----------------------------------------------------------------------------
' SetStage - ステータス行+Application.StatusBar 両方
' ----------------------------------------------------------------------------
Public Sub SetStage(ByVal msg As String)
    Dim displayMsg As String
    displayMsg = msg
    If LenB(displayMsg) = 0 Then displayMsg = "準備できています"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetHomeSheet()
    On Error GoTo 0

    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Range(RNG_STATUS).Value = "状態: " & displayMsg
        On Error GoTo 0
    End If

    On Error Resume Next
    Application.StatusBar = displayMsg
    On Error GoTo 0

    If LenB(msg) > 0 Then ShowTip
End Sub

' ----------------------------------------------------------------------------
' RenderAnswer - 回答本文セル(SafeLeft)+出典ブロック+所要秒
' ----------------------------------------------------------------------------
Public Sub RenderAnswer(ByVal answerText As String, hits() As Hit, ByVal nHits As Long, _
                        ByVal mode As String, ByVal seconds As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetHomeSheet()
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    mLastAnswerText = answerText

    Dim modeLabel As String
    If LCase$(mode) = "deep" Then
        modeLabel = "🔍 しっかり調べる"
    Else
        modeLabel = "⚡ すぐ聞く"
    End If

    Dim footer As String
    footer = vbLf & vbLf & "（" & modeLabel & " ・ 所要 " & modUtil.HumanSeconds(CDbl(seconds)) & "）"

    WriteSafe ws.Range(RNG_ANSWER), answerText & footer
    WriteSafe ws.Range(RNG_SOURCES), "📖 この回答のもと: " & JoinSourceLabels(hits, nHits)

    On Error Resume Next
    ws.Range(RNG_STATUS).Value = "状態: 回答ができました。出典もあわせてご確認ください。"
    On Error GoTo 0

    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' RenderSourcesPreview - 出典先出し(ドラフト生成前に呼ぶ)
' ----------------------------------------------------------------------------
Public Sub RenderSourcesPreview(hits() As Hit, ByVal nHits As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetHomeSheet()
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim preview As String
    preview = "📄 " & nHits & "件の資料がヒットしました: " & JoinSourceLabels(hits, nHits)
    WriteSafe ws.Range(RNG_SOURCES), preview
    WriteSafe ws.Range(RNG_ANSWER), "（資料を確認しました。ここから回答を作成します。もう少しお待ちください…）"
End Sub

' ----------------------------------------------------------------------------
' OnAskButton / OnModeQuick / OnModeDeep / OnOpenHowto / OnRunDiag
' ----------------------------------------------------------------------------
Public Sub OnAskButton()
    On Error GoTo Fail
    modAsk.AskFromUI
    RefreshBadgesAndDashboard
    Exit Sub

Fail:
    modLog.LogError "E0602", "modUIMain.OnAskButton", Err.Description
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub OnModeQuick()
    SetMode "quick"
End Sub

Public Sub OnModeDeep()
    SetMode "deep"
End Sub

Public Sub OnOpenHowto()
    On Error Resume Next
    ThisWorkbook.Worksheets(modAppDef.SH_HOWTO).Activate
    On Error GoTo 0
End Sub

Public Sub OnRunDiag()
    On Error GoTo Fail
    modDiag.RunDiagnostics
    Exit Sub
Fail:
    Err.Clear
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' ShowTip - 待ち時間豆知識(定型10本からRnd選択)
' ----------------------------------------------------------------------------
Public Sub ShowTip()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetHomeSheet()
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim tips() As String
    tips = TipList()

    Randomize
    Dim idx As Long
    idx = Int(Rnd() * (UBound(tips) - LBound(tips) + 1)) + LBound(tips)

    WriteSafe ws.Range(RNG_TIP), "💡 豆知識: " & tips(idx)
End Sub

' ----------------------------------------------------------------------------
' OnTtsButton - opt機能(読み上げ)のUIラッパー(§7.7)。opt直接参照はしない。
' ----------------------------------------------------------------------------
Public Sub OnTtsButton()
    If LenB(mLastAnswerText) = 0 Then
        MsgBox "読み上げる回答がありません。まず質問して、回答を受け取ってください。", _
               vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    Dim result As Variant
    result = modFeatures.InvokeFeature("tts", "SpeakAnswer", mLastAnswerText)

    Dim isErr As Boolean
    isErr = False
    If VarType(result) = vbString Then
        If Left$(CStr(result), 5) = "#ERR:" Then isErr = True
    End If

    If isErr Then
        MsgBox "この機能は現在利用できません(管理者が有効化すると使えます)。", _
               vbInformation, modAppDef.APP_NAME
    End If
End Sub

' ----------------------------------------------------------------------------
' ShowEmptyShelfHint - 本棚が空のとき、回答エリアに常設案内を表示する。
'   呼び出し判断はmodBoot側の責務(本棚が空かどうかの判定はここではしない)。
' ----------------------------------------------------------------------------
Public Sub ShowEmptyShelfHint()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetHomeSheet()
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    WriteSafe ws.Range(RNG_ANSWER), _
        "まだ本棚に資料がありません。" & vbLf & _
        "まず『マイ本棚』タブで資料を1つ追加してみましょう →"
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

Private Sub RefreshBadgesAndDashboard()
    On Error Resume Next
    modStats.EvaluateBadges
    On Error GoTo 0

    On Error Resume Next
    modUIDashboard.RenderDashboard
    On Error GoTo 0
End Sub

Private Sub SetMode(ByVal mode As String)
    WriteUiStateMode mode
    ApplyModeColors mode
End Sub

Private Function CurrentMode() As String
    CurrentMode = "quick"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_UISTATE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row   ' xlUp
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(CStr(ws.Cells(i, 1).Value), "mode", vbTextCompare) = 0 Then
            Dim v As String
            v = LCase$(Trim$(CStr(ws.Cells(i, 2).Value)))
            If v = "deep" Then CurrentMode = "deep"
            Exit Function
        End If
    Next i
End Function

Private Sub WriteUiStateMode(ByVal mode As String)
    Dim ws As Worksheet
    Set ws = EnsureUiStateSheet()
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    Dim r As Long
    r = 0
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(CStr(ws.Cells(i, 1).Value), "mode", vbTextCompare) = 0 Then
            r = i
            Exit For
        End If
    Next i

    If r = 0 Then
        r = lastRow + 1
        If r < 1 Then r = 1
        ws.Cells(r, 1).Value = "mode"
    End If
    ws.Cells(r, 2).Value = mode
End Sub

Private Function EnsureUiStateSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_UISTATE)
    On Error GoTo 0

    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = modAppDef.SH_UISTATE
        ws.Cells(1, 1).Value = "key"
        ws.Cells(1, 2).Value = "value"
        On Error Resume Next
        ws.Visible = 2   ' xlSheetVeryHidden
        On Error GoTo 0
    End If
    Set EnsureUiStateSheet = ws
    Exit Function
Fail:
    Set EnsureUiStateSheet = Nothing
End Function

Private Sub ApplyModeColors(ByVal mode As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetHomeSheet()
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim quickSelected As Boolean
    quickSelected = (LCase$(mode) <> "deep")

    On Error Resume Next
    PaintModeButton ws, "btn_mode_quick", quickSelected
    On Error GoTo 0
    On Error Resume Next
    PaintModeButton ws, "btn_mode_deep", Not quickSelected
    On Error GoTo 0
End Sub

Private Sub PaintModeButton(ByVal ws As Worksheet, ByVal shapeName As String, ByVal selected As Boolean)
    Dim shp As Shape
    Set shp = ws.Shapes(shapeName)
    If selected Then
        shp.Fill.ForeColor.RGB = COLOR_SELECTED_BG
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_SELECTED_FG
    Else
        shp.Fill.ForeColor.RGB = COLOR_UNSELECTED_BG
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_UNSELECTED_FG
    End If
End Sub

Private Function ModeCaption(ByVal mode As String) As String
    If mode = "deep" Then
        ModeCaption = "🔍 しっかり調べる (1〜2分)"
    Else
        ModeCaption = "⚡ すぐ聞く (10〜20秒)"
    End If
End Function

Private Sub EnsureQuestionName(ByVal ws As Worksheet)
    On Error Resume Next
    ThisWorkbook.Names(QUESTION_RANGE_NAME).Delete
    On Error GoTo 0

    On Error Resume Next
    ThisWorkbook.Names.Add Name:=QUESTION_RANGE_NAME, _
        RefersTo:="='" & ws.Name & "'!$" & Mid$(QUESTION_TOP_LEFT, 1, 1) & "$" & Mid$(QUESTION_TOP_LEFT, 2)
    On Error GoTo 0
End Sub

Private Function JoinSourceLabels(hits() As Hit, ByVal nHits As Long) As String
    If nHits <= 0 Then
        JoinSourceLabels = "(該当資料なし)"
        Exit Function
    End If

    Const MAX_SHOW As Long = 8
    Dim showCount As Long
    showCount = nHits
    If showCount > MAX_SHOW Then showCount = MAX_SHOW

    Dim parts() As String
    ReDim parts(0 To showCount - 1)
    Dim i As Long
    For i = 0 To showCount - 1
        parts(i) = OneSourceLabel(hits(i))
    Next i

    Dim joined As String
    joined = Join(parts, " / ")
    If nHits > MAX_SHOW Then
        joined = joined & " 他" & (nHits - MAX_SHOW) & "件"
    End If
    JoinSourceLabels = joined
End Function

Private Function OneSourceLabel(ByRef h As Hit) As String
    Dim label As String
    label = h.source & " p." & h.page
    If Left$(h.origin, 5) = "pack:" Then
        label = label & "（パック: " & Mid$(h.origin, 6) & "）"
    End If
    OneSourceLabel = label
End Function

Private Function TipList() As Variant
    TipList = Array( _
        "「🔍しっかり調べる」は下書き→検証の2段階なので少し時間がかかりますが、より丁寧な回答になります。", _
        "資料を追加すると本棚が育ち、答えられる質問がどんどん増えていきます。", _
        "資料は「マイ本棚」タブの「フォルダと同期」でまとめて自動追加できます。", _
        "回答の下にある出典を見れば、元の資料のどこに書いてあるかすぐ確認できます。", _
        "同じ資料を入れ直すと、自動的に新しい内容に置き換わります(重複しません)。", _
        "他の人が作った「パック」を取り込むと、自分で資料を集めなくても本棚が増やせます。", _
        "🟢🟡🔴のボタンで感想を送ると、ダッシュボードの記録に残ります。", _
        "困ったときは🩺診断ボタンを押すと、今の状態が一目でわかります。", _
        "質問はできるだけ具体的に書くと、より的確な回答が返ってきます。", _
        "ダッシュボードでは、これまで取り戻した時間やバッジの獲得状況が見られます。" _
    )
End Function

Private Sub WriteSafe(ByVal cell As Range, ByVal text As String)
    Dim t As String
    t = modUtil.SafeLeft(text, 32000)
    On Error Resume Next
    cell.Value = t
    On Error GoTo 0
End Sub

Private Function GetHomeSheet() As Worksheet
    On Error Resume Next
    Set GetHomeSheet = ThisWorkbook.Worksheets(modAppDef.SH_HOME)
    On Error GoTo 0
End Function

Private Function GetOrCreateHomeSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_HOME)
    On Error GoTo 0
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = modAppDef.SH_HOME
        On Error GoTo 0
    End If
    Set GetOrCreateHomeSheet = ws
    Exit Function
Fail:
    Set GetOrCreateHomeSheet = Nothing
End Function

Private Sub AddButton(ByVal ws As Worksheet, ByVal rng As Range, ByVal shapeName As String, _
                      ByVal caption As String, ByVal action As String)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, rng.Left, rng.Top, rng.Width, rng.Height)   ' 5 = msoShapeRoundedRectangle
    shp.Name = shapeName
    shp.TextFrame2.TextRange.Text = caption
    shp.TextFrame2.WordWrap = -1   ' msoTrue
    shp.TextFrame2.TextRange.Font.Size = 11
    shp.TextFrame2.TextRange.Font.Bold = -1   ' msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = 2   ' msoAlignCenter
    shp.TextFrame2.VerticalAnchor = 3   ' msoAnchorMiddle
    shp.Fill.ForeColor.RGB = COLOR_UNSELECTED_BG
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_UNSELECTED_FG
    shp.Line.Visible = 0   ' msoFalse
    shp.OnAction = action
End Sub

Private Sub RemoveManagedShapes(ByVal ws As Worksheet)
    Dim names() As String
    ReDim names(0 To ws.Shapes.Count)
    Dim n As Long
    n = 0

    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 4) = "btn_" Or Left$(shp.Name, 4) = "lbl_" Then
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
