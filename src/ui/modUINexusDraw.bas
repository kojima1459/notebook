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

Private Const NEXUS_SHEET As String = "Nexus"
Public Const HDR_H As Double = 42          ' ヘッダー1段ぶんの高さ(行1の既定)
Public Const INPUT_ROW As Long = 3         ' 入力欄の行
Public Const ACT_H As Double = 22          ' 文脈アクションpillの高さ
Private Const HDR_BTN_H As Double = 26
Private Const HDR_GAP As Double = 6        ' 右端ピルの間隔
Private Const HDR_RIGHT_PAD As Double = 8  ' ヘッダー右端の余白
Private Const HDR_TITLE_MIN As Double = 280 ' タイトルのために必ず空けておく幅
Private Const HDR_TITLE_LEFT As Double = 82 ' タイトル文字の左余白(bgのMarginLeft)
Private Const HDR_TITLE_HEAD As Double = 96 ' "💬 チャット   " の実測目安(pt)
Private Const HDR_TITLE_PITCH As Double = 12 ' タイトル12pt太字の全角1字ぶん
Private Const HDR_PILL_PITCH As Double = 9   ' ピル9ptの全角1字ぶん
Private Const HDR_PILL_PAD As Double = 14    ' ピル内の左右余白の合計
Private Const HDR_PILL_MIN As Double = 30    ' アイコン1個ぶんの最小幅
Private Const HDR_PILLS As Long = 8          ' 右端ピルの個数

' 直近に描いたヘッダーの実使用高さ(段数×HDR_H)。行1の高さに反映する。
Private mHeaderH As Double

' チャット領域(バブル/アクションの基準)。列幅の実測から求める。
Public Function ChatLeft(ByVal ws As Worksheet) As Double
    ChatLeft = ws.Range("B1").Left
End Function

Public Function ChatWidth(ByVal ws As Worksheet) As Double
    ChatWidth = ws.Range("B1:L1").Width
End Function

' 固定領域(行1～4)の直下=会話の開始Y。FreezePanesの境界と必ず一致する。
Public Function ChatTop(ByVal ws As Worksheet) As Double
    ChatTop = ws.Rows(5).Top
End Function

' 直近のDrawChatHeaderが実際に使った高さ。呼び出し元(modUI.InitUI)が
' 行1の高さに入れることで、2段目へ流れた場合でも入力欄と重ならない。
Public Function HeaderHeight() As Double
    If mHeaderH < HDR_H Then mHeaderH = HDR_H
    HeaderHeight = mHeaderH
End Function

' ----------------------------------------------------------------------------
' DrawChatHeader - ヘッダー。左に「← Hub」、右に操作pillを右詰めで並べる。
'   戻り値 = 実際に使った高さ(pt)。
' ----------------------------------------------------------------------------
' 2026-07-30(R4要件D)の作り直し:
'   実機で ⚡すぐ聞く がタイトル「💬 チャット」に重なっていた。原因は幅の
'   決め打ちで、右端ピル8個の固定予約幅の合計は 606pt(exit30/clear62/sq74/
'   help30/theme26/lang92/mode118/speed124 + gap6×7 + 右余白8)、対して
'   Nexusシートのヘッダー実幅 W(A1:M1)は約597pt。**予約が実幅を超えている
'   のに、一度も W と突き合わせずに右から積んでいた**ので、最内側に来る
'   speed が必ずタイトルへ食い込む。「たまたま重なった」のではなく
'   「重ならない置き方が存在しない」状態だった。
'
'   直し方は3段構え。どれも「固定幅の合計 vs 予算」を必ずコードで比較する。
'     1. タイトル用に HDR_TITLE_MIN(280pt)を必ず確保し、残りを予算とする。
'     2. 通常表記→短縮表記→アイコンのみ、と収まる中で最も情報量の多い
'        表記を選ぶ。幅はキャプションの実文字から出す(予約と実物が
'        食い違わないようにする)。
'     3. それでも収まらない端末では modChrome.FlowRight が段を増やして
'        2段目以降へ右詰めで流す(同じ濃色帯を1行ぶん足す)。
'   3のおかげで、Wがどんな値でも「タイトル領域へ入り込む配置」は
'   構造的に作れない(FlowRightのコメント参照)。
Public Function DrawChatHeader(ByVal ws As Worksheet) As Double
    Dim L As Double, cellW As Double, W As Double
    L = ws.Range("A1").Left
    cellW = ws.Range("A1:M1").Width
    ' R11-B(#30恒久対策): セル範囲幅(cellW)だけでなく実可視幅も見る。
    ' 帯の背景(下のbg)は従来どおりcellWいっぱいのまま、操作系(ピル)の
    ' 右端だけをmodChrome.BarWidthでウィンドウ実幅へクランプする。
    W = modChrome.BarWidth(cellW, modUIMain.ViewportWidth(), HDR_RIGHT_PAD)

    ' --- 幅予算の決定 ---------------------------------------------------
    ' タイトル用に HDR_TITLE_MIN を確保し、残りを右端ピルの予算とする。
    ' 帯そのものが極端に狭い端末では確保幅のほうを縮める(確保しすぎると
    ' ピルの置き場所が1ptも残らず、かえって壊れる)。
    Dim titleMin As Double
    titleMin = modChrome.TitleReserve(W, HDR_TITLE_MIN)
    Dim budget As Double
    budget = W - HDR_RIGHT_PAD - titleMin

    ' 通常表記のキャプションは、短縮したときのツールチップに使うので控えておく。
    Dim capsFull() As String, nmFull() As String, actsFull() As String
    Dim wFull() As Double
    PillSpec 0, capsFull, nmFull, actsFull, wFull

    ' 収まる中でいちばん情報量の多い表記を選ぶ。ここが「固定幅の合計と
    ' 帯幅Wを突き合わせる」比較そのもの(旧実装にはこの比較が無かった)。
    ' 判断は modChrome.PickTier(純ロジック)に置き、実行テストが同じ
    ' 判断を通れるようにする(レビュー3-A(3))。
    Dim caps1() As String, nm1() As String, acts1() As String, w1() As Double
    Dim caps2() As String, nm2() As String, acts2() As String, w2() As Double
    PillSpec 1, caps1, nm1, acts1, w1
    PillSpec 2, caps2, nm2, acts2, w2

    Dim tier As Long
    Dim caps() As String, nm() As String, acts() As String
    Dim widths() As Double
    tier = modChrome.PickTier(wFull, w1, w2, HDR_PILLS, HDR_GAP, budget)
    PillSpec tier, caps, nm, acts, widths

    ' --- 配置(右詰め・段数無制限) ---------------------------------------
    Dim xs() As Double, rws() As Long, useW() As Double
    Dim rowN As Long
    rowN = modChrome.FlowRight(widths, HDR_PILLS, L + W - HDR_RIGHT_PAD, _
                               L + titleMin, L + HDR_RIGHT_PAD, HDR_GAP, _
                               xs, rws, useW)
    If rowN < 1 Then rowN = 1
    mHeaderH = rowN * HDR_H

    ' --- 濃色帯(段数ぶんの高さ) -----------------------------------------
    ' 背景の帯は従来どおりセル幅いっぱい(cellW)。狭めるのは操作系の右端だけ。
    Dim bg As Shape
    Set bg = ws.Shapes.AddShape(5, L, 0, cellW, mHeaderH)
    bg.Name = "nx_top_bg"
    bg.Placement = 3
    bg.Adjustments(1) = 0.02
    bg.Line.Visible = 0
    bg.Fill.ForeColor.RGB = modUI.UiColor("sidebar")
    modSkin.ApplyHeaderDepth bg           ' §9: 濃紺の2色グラデーション

    ' --- タイトル(部門ラベルは残り幅に合わせて切り詰める) ---------------
    ' 「今どの部門の公式ナレッジにつないでいるか」を必ず見せる。
    ' 切替式である以上、これが見えないと「なぜ答えられないのか」が
    ' 分からなくなる(迷子の最大要因)。ただし幅を超えてまで出さない。
    Dim chLabel As String
    On Error Resume Next
    chLabel = modChannel.ActiveLabel()
    On Error GoTo 0

    Dim titleAvail As Double
    titleAvail = LeftmostRow0(xs, rws, HDR_PILLS, L + W - HDR_RIGHT_PAD) _
                 - (L + HDR_TITLE_LEFT) - HDR_GAP
    chLabel = modChrome.ClipToWidth(chLabel, titleAvail - HDR_TITLE_HEAD, HDR_TITLE_PITCH)

    With bg.TextFrame2
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCAC) & " チャット   " & chLabel
        .TextRange.Font.Size = 12
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .MarginLeft = HDR_TITLE_LEFT
        .VerticalAnchor = 3
    End With

    ' 戻り導線はヘッダーと同色だと埋没する。白地+濃紺文字で最も目立たせる。
    HeaderButton ws, "nx_top_back", ChrW(&H2190) & " Hub", _
                 L + HDR_RIGHT_PAD, PillTop(0), 72, "modApp.OnNavHome"
    On Error Resume Next
    With ws.Shapes("nx_top_back")
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("sidebar")
        .TextFrame2.TextRange.Font.Size = 10
    End With
    On Error GoTo 0

    ' --- 右端ピル -------------------------------------------------------
    Dim i As Long
    For i = 0 To HDR_PILLS - 1
        If nm(i) = "nx_top_theme" Then
            HeaderTheme ws, xs(i), PillTop(rws(i))
        Else
            HeaderButton ws, nm(i), caps(i), xs(i), PillTop(rws(i)), useW(i), acts(i)
        End If
        ' 短縮表記のときは「何のボタンか」をツールチップで補う。
        ' 絵文字だけにした瞬間に意味が消えるボタン(モード/速度)があるため。
        If tier > 0 Then SetPillTip ws, nm(i), capsFull(i)
    Next i

    DrawChatHeader = mHeaderH
End Function

' ----------------------------------------------------------------------------
' RedrawChatHeader - ヘッダーだけを作り直す(モード/速度/言語のトグル後)。
' ----------------------------------------------------------------------------
' 2026-07-30(レビュー3-A(1)): トグルの3箇所は、ピルのTextRangeへ新しい
' キャプションを直接書き込んでいた。書き込みは幅計算(ティア選択+FlowRight)
' を一切通らないので、
'   ・「⚡ すぐ聞く」→「🔬 入念に調べる」のように長い語に変わった瞬間、
'     ピルの幅は元のままで文字だけがはみ出す
'   ・幅が変わらない以上、隣のピルもタイトルも位置を譲らないので重なる
' という壊れ方をする。実機でヘッダーが崩れていた直接の原因がこれ。
' 状態が変わったらヘッダー全体を描き直す(= 必ず幅から作り直す)。
Public Sub RedrawChatHeader()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NEXUS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ClearChatHeader ws
    On Error Resume Next
    DrawChatHeader ws
    ' 2段目へ流れた場合に備えて行1の高さも合わせ直す(InitUIと同じ手順)。
    ws.Rows(1).RowHeight = HeaderHeight()
    modUI.BringFixedToFront ws
    On Error GoTo 0
End Sub

' ヘッダーが持つShapeだけを消す。入力欄の nx_top_add / nx_top_send は
' 同じ接頭辞だがヘッダーの持ち物ではないので消さない(消すと📎と送信が
' 二度と戻らない)。
Private Sub ClearChatHeader(ByVal ws As Worksheet)
    Dim names() As String
    ReDim names(0 To ws.Shapes.Count)
    Dim n As Long
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 7) = "nx_top_" Then
            If shp.Name <> "nx_top_add" And shp.Name <> "nx_top_send" Then
                names(n) = shp.Name
                n = n + 1
            End If
        End If
    Next shp
    Dim i As Long
    For i = 0 To n - 1
        On Error Resume Next
        ws.Shapes(names(i)).Delete
        On Error GoTo 0
    Next i
End Sub

' 段番号からピルのY座標(その段の中央)を出す。
Private Function PillTop(ByVal rowIdx As Long) As Double
    PillTop = rowIdx * HDR_H + (HDR_H - HDR_BTN_H) / 2
End Function

' 1段目に置かれたピルのうち最も左のX(=タイトルが使える右端)。
Private Function LeftmostRow0(ByRef xs() As Double, ByRef rws() As Long, _
                              ByVal n As Long, ByVal fallbackX As Double) As Double
    Dim edge As Double: edge = fallbackX
    Dim i As Long
    For i = 0 To n - 1
        If rws(i) = 0 Then
            If xs(i) < edge Then edge = xs(i)
        End If
    Next i
    LeftmostRow0 = edge
End Function

' ----------------------------------------------------------------------------
' PillSpec - 右端ピル8個の表記と幅を組み立てる。
'   tier 0=通常 / 1=無状態ボタンをアイコンのみへ / 2=状態ピルも短い語へ
'   幅はキャプションの実文字から出す(固定の予約幅が実物と食い違って
'   いたことが重なりの原因だったため、予約という概念自体をやめる)。
' ----------------------------------------------------------------------------
' ティア方針(2026-07-30 レビュー3-A(2)):
'   状態を持つ3つ(モード/速度/言語)は、どのティアでも短い文字ラベルを保つ。
'   アイコンだけにすると「今どちらなのか」が画面から消え、押す前に押した
'   結果が分からないトグルになる。実機幅(約597pt)では常に最短ティアが
'   選ばれるため、ここを落とすと【常に状態が読めない】ことになっていた。
'   無状態の5つ(終了/クリア/質問例/ヘルプ/テーマ)は意味がアイコンに載って
'   いるので、先にこちらを落とす。それでも1段に入らない分は
'   modChrome.FlowRight が2段目へ流す(重なりゼロが最優先)。
Private Sub PillSpec(ByVal tier As Long, ByRef caps() As String, ByRef nm() As String, _
                     ByRef acts() As String, ByRef widths() As Double)
    ReDim caps(0 To HDR_PILLS - 1)
    ReDim nm(0 To HDR_PILLS - 1)
    ReDim acts(0 To HDR_PILLS - 1)
    ReDim widths(0 To HDR_PILLS - 1)

    Dim clearCap As String, sqCap As String
    Dim langCap As String, modeCap As String, speedCap As String
    clearCap = ChrW(&HD83D) & ChrW(&HDDD1) & " クリア"
    sqCap = ChrW(&HD83D) & ChrW(&HDCA1) & " 質問例"
    langCap = LangCaption()
    modeCap = modApp.ModeCaption()
    speedCap = modApp.SpeedCaption()

    If tier >= 1 Then
        ' 無状態のボタンはアイコンのみ(意味が絵柄に載っている)。
        clearCap = modChrome.LeadIcon(clearCap)
        sqCap = modChrome.LeadIcon(sqCap)
    End If
    If tier >= 2 Then
        ' 状態ピルは短い語へ。アイコンだけにはしない(状態が読めなくなる)。
        modeCap = modChrome.LeadIcon(modeCap) & " " & ModeShortWord()
        speedCap = modChrome.LeadIcon(speedCap) & " " & SpeedShortWord()
        langCap = modChrome.LeadIcon(langCap) & " " & LangShortWord()
    End If

    ' 右から左へ積む順(配列の先頭=最も右)。
    SetPill caps, nm, acts, widths, 0, ChrW(&HD83D) & ChrW(&HDEAA), _
            "nx_top_exit", "modApp.OnSaveAndExit"
    SetPill caps, nm, acts, widths, 1, clearCap, "nx_top_clear", "modApp.OnClearChat"
    SetPill caps, nm, acts, widths, 2, sqCap, "nx_top_sq", "modStarter.OnShowList"
    SetPill caps, nm, acts, widths, 3, ChrW(&H2753), "nx_top_help", "modHelp.OnHelpClick"
    SetPill caps, nm, acts, widths, 4, "", "nx_top_theme", "modUI.ToggleTheme"
    widths(4) = HDR_BTN_H                 ' テーマは正円(別Shape)なので固定
    SetPill caps, nm, acts, widths, 5, langCap, "nx_top_lang", "modApp.OnLangCycle"
    SetPill caps, nm, acts, widths, 6, modeCap, "nx_top_mode", "modApp.OnToggleMode"
    SetPill caps, nm, acts, widths, 7, speedCap, "nx_top_speed", "modApp.OnToggleSpeed"
End Sub

Private Sub SetPill(ByRef caps() As String, ByRef nm() As String, ByRef acts() As String, _
                    ByRef widths() As Double, ByVal idx As Long, ByVal capText As String, _
                    ByVal shapeName As String, ByVal action As String)
    caps(idx) = capText
    nm(idx) = shapeName
    acts(idx) = action
    widths(idx) = modChrome.PillWidth(capText, HDR_PILL_PITCH, HDR_PILL_PAD, HDR_PILL_MIN)
End Sub

' 短縮表記にしたピルへツールチップ(代替テキスト)で元の意味を添える。
Private Sub SetPillTip(ByVal ws As Worksheet, ByVal shapeName As String, ByVal tipText As String)
    On Error Resume Next
    ws.Shapes(shapeName).AlternativeText = tipText
    On Error GoTo 0
End Sub

Private Sub HeaderButton(ByVal ws As Worksheet, ByVal shapeName As String, _
                         ByVal caption As String, ByVal x As Double, ByVal y As Double, _
                         ByVal w As Double, ByVal action As String)
    On Error Resume Next
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(5, x, y, w, HDR_BTN_H)
    If btn Is Nothing Then Exit Sub
    btn.Name = shapeName
    btn.Placement = 3          ' 行1の高さを後から変えてもピルは動かさない
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

Private Sub HeaderTheme(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double)
    On Error Resume Next
    Dim th As Shape
    Set th = ws.Shapes.AddShape(9, x, y, HDR_BTN_H, HDR_BTN_H)
    If th Is Nothing Then Exit Sub
    th.Name = "nx_top_theme"
    th.Placement = 3           ' 行1の高さを後から変えても動かさない
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

' ----------------------------------------------------------------------------
' 状態ピルの短い語(レビュー3-A(2))。アイコン+この語で「今どちらか」を保つ。
' 語は状態の単一情報源(modAppState/modMode/config)から引く。
' ----------------------------------------------------------------------------
Private Function ModeShortWord() As String
    Dim m As String
    On Error Resume Next
    m = modAppState.CurrentMode()
    On Error GoTo 0
    If m = "normal" Then
        ModeShortWord = "一般"
    Else
        ModeShortWord = "社内"
    End If
End Function

Private Function SpeedShortWord() As String
    Dim m As String: m = "quick"
    On Error Resume Next
    m = modMode.Normalize(modAppState.ReadUiState("mode", "quick"))
    On Error GoTo 0
    Select Case m
        Case "thorough": SpeedShortWord = "徹底"
        Case "deep": SpeedShortWord = "調べ"
        Case Else: SpeedShortWord = "すぐ"
    End Select
End Function

Private Function LangShortWord() As String
    Dim v As String
    v = modChrome.TailWords(LangCaption())
    Select Case v
        Case "日本語": LangShortWord = "日"
        Case "English": LangShortWord = "EN"
        Case "中文": LangShortWord = "中"
        Case "関西弁": LangShortWord = "関西"
        Case Else
            ' "Tiếng Việt" 等(CP932外の文字を含むのでリテラルで比較しない)。
            If InStr(v, "Vi") > 0 Then
                LangShortWord = "VN"
            Else
                LangShortWord = Left$(v, 2)
            End If
    End Select
End Function

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

    ' 入力欄の左は「資料を入れる」。
    '
    ' READMEが売っているのは「ぶちこんで、すぐ聞ける」という2つの動詞なのに、
    ' 起動して最初に立つこの画面には、後半の「聞く」しか無かった。資料を
    ' 入れるには Hub → ナレッジ倉庫 → 11個並んだ同じ見た目のボタンの3番目、
    ' と3画面移動する必要がある。初めて開いた人が辿り着けるはずがない。
    ' 送信ボタンの真向かいに置いて、2つの動詞を同じ場所に揃える。
    '
    ' ここにあった📎(クリップボード画像)はナレッジ画面の「📸 スクショ取込」
    ' と同じ機能で、そちらに残っている。1等地は主役の動詞に譲る。
    Dim cellB As Range: Set cellB = ws.Range("B" & r)
    Dim addD As Double: addD = 28
    Dim addBtn As Shape
    Set addBtn = ws.Shapes.AddShape(9, _
        cellB.Left + (cellB.Width - addD) / 2, rowTop + (rowH - addD) / 2, addD, addD)
    addBtn.Name = "nx_top_add"
    addBtn.Line.Visible = 0
    addBtn.Fill.ForeColor.RGB = modUI.UiColor("accent")
    With addBtn.TextFrame2
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCC1)
        .TextRange.Font.Size = 12
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
    End With
    addBtn.OnAction = "modApp.OnAddDocs"
    On Error Resume Next
    addBtn.AlternativeText = "資料を入れる"
    On Error GoTo 0

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
    ' 資料が1件も無いあいだは、ショートカットの案内より先に
    ' 「何をすれば使えるようになるか」を出す。空のときにキーボード
    ' ショートカットを教えても、押す先が無い。
    Dim hint As String
    Dim n As Long
    On Error Resume Next
    n = modShelf.TotalChunks()
    On Error GoTo 0
    If n = 0 Then
        hint = ChrW(&HD83D) & ChrW(&HDCC1) & " 左の緑のボタンから約款やマニュアルを入れると、" & _
               "出典付きで答えられるようになります(そのまま質問もできます)"
    Else
        ' 2026-07-28(レビュー L-19): セル編集中は OnKey が効かず1回目の
        ' Ctrl+Enter は「確定」になる(Excelの仕様)。実挙動に合わせて書く。
        hint = "入力後に Ctrl+Enter で送信 ・ Ctrl+Shift+Q でどこからでも呼び出し ・ " & _
               ChrW(&HD83D) & ChrW(&HDCC1) & " で資料を追加"
    End If
    With ws.Range("C" & (INPUT_ROW + 1))
        .Value = hint
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

    ' 評価ボタンは「信頼度バッジ」と「出典チップ」より必ず下に置く。
    ' 画面の縦順は、何を先に読ませたいかの主張そのもの。根拠を見る前に
    ' 「解決した / 違う」を押させる並びは、評価してから確かめろと言っている
    ' に等しい。損保でAIの回答を信じてよい理由は原文確認だけなので、
    ' 根拠のほうを先に、上に置く。
    Dim y As Double: y = anchor.Top + anchor.Height + 6
    Dim cb As Double
    On Error Resume Next
    cb = ConfidenceBottom(ws)
    If cb + 6 > y Then y = cb + 6
    cb = modPeek.CitationsBottom(ws)
    If cb + 8 > y Then y = cb + 8
    On Error GoTo 0

    Dim caps As Variant, kinds As Variant, acts As Variant
    ' 評価は「解決した/微妙/違う」の3択。どれも1クリックで完結し、
    ' 入力を強制しない(入力を求めた途端に誰も押さなくなる)。
    '
    ' 「本社照会」は削除した。押しても常に「準備中です」としか返らない
    ' ボタンが、全ての回答の下に永久に並んでいた。動かないものが1つでも
    ' 混じっていると、利用者は「他も見せかけかもしれない」と学習する。
    ' 出せないなら出さないほうが、信用は減らない。
    caps = Array(ChrW(&H2705) & " 解決した", _
                 ChrW(&HD83E) & ChrW(&HDD14) & " 微妙", _
                 ChrW(&H274C) & " 違う", _
                 ChrW(&HD83D) & ChrW(&HDD0D) & " 深掘り", _
                 ChrW(&HD83D) & ChrW(&HDCCB) & " コピー", _
                 ChrW(&HD83D) & ChrW(&HDCC4) & " Word")
    kinds = Array("resolve", "unsure", "bad", "drill", "copy", "word")
    acts = Array("OnActResolve", "OnActUnsure", "OnActBad", "OnActDrill", _
                 "OnActCopy", "OnActWord")

    ' R11-B(#30): 固定幅6個の決め打ちをやめ、modChrome.PillWidth+FlowLeftで
    ' 実可視幅(modUIMain.ViewportWidth)基準に流し込む。入り切らない分は
    ' modChrome.FlowLeftが段を増やして下へ流す(画面外へ見切れない)。
    Dim widths(0 To 5) As Double
    Dim i As Long
    For i = 0 To 5
        widths(i) = modChrome.PillWidth(CStr(caps(i)), 13, 16, 40)
    Next i
    Dim maxX As Double
    maxX = ChatLeft(ws) + modChrome.BarWidth(ChatWidth(ws), modUIMain.ViewportWidth(), 8)

    Dim xs() As Double, rws() As Long, useW() As Double
    Dim rowN As Long
    rowN = modChrome.FlowLeft(widths, 6, anchor.Left, maxX, 5, xs, rws, useW)
    If rowN < 1 Then rowN = 1

    For i = 0 To 5
        Dim rowY As Double: rowY = y + rws(i) * (ACT_H + 6)
        ' 1個の1004で残りを道連れにしない(実機で繰り返した描画中断の教訓)。
        On Error Resume Next
        Dim btn As Shape
        Set btn = ws.Shapes.AddShape(5, xs(i), rowY, useW(i), ACT_H)
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
    Next i
End Sub

' ----------------------------------------------------------------------------
' DrawConfidence - 信頼度バッジ。回答バブルの直下、いちばん最初に読ませる1行。
' ----------------------------------------------------------------------------
' 「この答えを信じてよいか」を利用者が自分で判断できないことが、フィードバックが
' 集まらない根本原因だった。検索スコアという機械側の情報を、確認すべきときだけ
' 確認を促す一文に翻訳して先に見せる。
'
' Shape名を nx_conf_ にしたのは、評価ボタン(nx_act_)の掃除に巻き込まれないため。
' 描画順が バッジ → 出典 → 評価 になったので、最後に走る ClearContextActions で
' 消されてしまうのを構造的に防ぐ。
Public Function DrawConfidence(ByVal ws As Worksheet, ByVal bubbleName As String) As Double
    ClearConfidence ws
    If ws Is Nothing Then Exit Function
    If LenB(bubbleName) = 0 Then Exit Function

    Dim anchor As Shape
    On Error Resume Next
    Set anchor = ws.Shapes(bubbleName)
    On Error GoTo 0
    If anchor Is Nothing Then Exit Function

    DrawConfidence = anchor.Top + anchor.Height

    Dim confText As String
    On Error Resume Next
    confText = modAsk.LastConfidenceText()
    On Error GoTo 0
    If LenB(confText) = 0 Then Exit Function

    On Error Resume Next
    Dim badge As Shape
    Set badge = ws.Shapes.AddShape(5, anchor.Left, anchor.Top + anchor.Height + 6, 330, 22)
    If Err.Number = 0 And Not badge Is Nothing Then
        badge.Name = "nx_conf_badge"
        badge.Adjustments(1) = 0.45
        badge.Line.Visible = 0
        ' 8pt・塗りなし・muted では、いちばん大事な一文がいちばん目立たない
        ' 字になっていた。色の付いたピルにして、先に目に入るようにする。
        badge.Fill.Visible = -1
        badge.Fill.ForeColor.RGB = ConfidenceTint(confText)
        With badge.TextFrame2
            .WordWrap = -1
            .TextRange.Text = confText
            .TextRange.Font.Size = 9
            .TextRange.Font.Bold = -1
            .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
            .VerticalAnchor = 3
            .MarginLeft = 8: .MarginRight = 8: .MarginTop = 0: .MarginBottom = 0
        End With
        badge.Placement = 3
        DrawConfidence = badge.Top + badge.Height
    End If
    Set badge = Nothing
    Err.Clear
    On Error GoTo 0
End Function

Public Sub ClearConfidence(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Shapes("nx_conf_badge").Delete
    On Error GoTo 0
End Sub

' 信頼度バッジの下端(無ければ0)。出典・評価の積み上げ基準に使う。
Public Function ConfidenceBottom(ByVal ws As Worksheet) As Double
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Dim shp As Shape
    Set shp = ws.Shapes("nx_conf_badge")
    If Not shp Is Nothing Then ConfidenceBottom = shp.Top + shp.Height
    On Error GoTo 0
End Function

' 信頼度バッジの背景色。🟢/🟡/🔴 のどれで始まるかだけで決める
' (文言そのものは modAsk が持つ単一情報源。ここでは色だけを足す)。
Private Function ConfidenceTint(ByVal confText As String) As Long
    Dim head As String
    head = Left$(confText, 2)
    If head = ChrW(&HD83D) & ChrW(&HDFE2) Then          ' U+1F7E2 緑 強く一致
        ConfidenceTint = RGB(220, 245, 225)
    ElseIf head = ChrW(&HD83D) & ChrW(&HDFE1) Then      ' U+1F7E1 黄 部分的
        ConfidenceTint = RGB(253, 246, 214)
    ElseIf head = ChrW(&HD83D) & ChrW(&HDD34) Then      ' U+1F534 赤 根拠が乏しい
        ConfidenceTint = RGB(253, 226, 226)
    Else
        ConfidenceTint = modUI.UiColor("bg")
    End If
End Function

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
