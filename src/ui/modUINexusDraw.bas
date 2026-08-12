Attribute VB_Name = "modUINexusDraw"
Option Explicit

' modUINexusDraw - チャット画面(Nexusシート)の骨格描画。
'
' 2026-07-26 再設計(nexus-spec-v1 §2.2 / nexus-ui-final 画面5):
'   ・サイドバーを全廃。移動導線はHub画面(modHub)に集約し、チャット画面は
'     「会話」だけを担う(バブルの表示幅が約30%広がる)。
'   ・常時表示のアクションバー(7個)を廃止し、最新のAI回答バブルの直下にだけ
'     6個のpillを出す「文脈表示」へ(質問前はボタン0個で入力に集中できる)。
'   ・ヘッダーは1行(← Hub / タイトル / 速度 / モード / 言語 / テーマ / クリア)。
'
' 設計の鉄則(実機で繰り返し事故った点):
'   ・Shape座標は必ず実セル幾何(Range.Left/.Width/Rows().Top)から導く
'     (pt決め打ちは列幅・行高・DPIの差で必ずズレる)。
'   ・配色は modUI.UiColor() を単一情報源にする。
'   ・絵文字はChrW()で組み立てる(直書きは自己インストーラの文字列注入で
'     化ける)。BMP外はサロゲートペアで2つ繋ぐ。

Private Const NEXUS_SHEET As String = "Nexus"

' R18-3a/3b → R19-1b(実機第6報①): チャット画面が使うセル範囲。
' 列は A=左余白 / B=資料を入れる / C:K=入力欄(結合。Kが余りを吸う) /
' L=送信 / M=右余白。旧実装は M:P の4列を右余白にして帯を671ptにし、
' さらに書式を A1:P2000(約31,000pt=40画面ぶん)へ当てていた ―― これは
' 「制限を入れた」のではなく「無限スクロールを明文で許可した」状態で、
' 起動のたびに32,000セルぶんのスタイルをブックへ焼き付けてもいた。
' 帯は A:M をKで可視幅ぴったりに詰め、下端は会話の実下端(mChatBottom)から
' 決める(NexusBound)。行の上限だけは大きく残す: バブルは行と無関係にpt座標で
' 下へ伸び続けるので、上限が低いと後半のバブルが到達不能になる(調査agent2 §2.3)。
Public Const NEXUS_BAND As String = "A1:M1"
Public Const NEXUS_PAD_COL As String = "M"
Public Const NEXUS_MAX_ROW As Long = 2000

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

' NexusBound - チャットの実使用範囲 "A1:M<行>"。塗り・ScrollArea・Locked が
'   共有する単一情報源。会話が伸びれば下端も伸びる(modUI.AddChatBubble が
'   modSkin.ExtendChatBand 経由で追随させ、ClearChat で戻る)。
Public Function NexusBound(ByVal ws As Worksheet) As String
    If ws Is Nothing Then Exit Function
    NexusBound = modViewport.BoundAddr(ws, NEXUS_PAD_COL, _
                                       modUI.ChatBottomFor(ws), NEXUS_MAX_ROW)
End Function

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
'   実機で ⚡すぐ聞く がタイトルに重なっていた。右端ピル8個の固定予約幅の
'   合計606ptが帯の実幅(約597pt)を超えているのに、一度も突き合わせずに
'   右から積んでいたため「重ならない置き方が存在しない」状態だった。
'   直し方は3段構えで、どれも「固定幅の合計 vs 予算」を必ずコードで比較する。
'     1. タイトル用に HDR_TITLE_MIN(280pt)を確保し、残りを予算とする。
'     2. 通常→短縮→アイコンのみ、と収まる中で最も情報量の多い表記を選ぶ
'        (幅はキャプションの実文字から出す=予約と実物が食い違わない)。
'     3. それでも収まらない端末では modChrome.FlowRight が段を増やす。
Public Function DrawChatHeader(ByVal ws As Worksheet) As Double
    Dim L As Double, cellW As Double, W As Double
    L = ws.Range("A1").Left
    ' R19-1b: 帯の右端も操作系の右端も modViewport.ContentRight 1本から取る
    ' (帯637.5 / ピル629.5 / 送信622.25 の3段ズレは、同じ右端を3通りに
    ' 計算していたのが原因)。帯は余白0=可視幅いっぱい、ピルと送信ボタンは
    ' 余白 HDR_RIGHT_PAD ぶん内側で、必ず互いに一致する。
    cellW = modViewport.ContentRight(ws, NEXUS_BAND, 0) - L
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
' 2026-07-30(レビュー3-A(1)): トグルの3箇所はピルのTextRangeへ新しい
' キャプションを直接書き込んでいた。書き込みは幅計算(ティア選択+FlowRight)を
' 通らないので、長い語に変わった瞬間に文字がはみ出し、幅が変わらない以上
' 隣のピルもタイトルも位置を譲らず重なる(実機のヘッダー崩れの直接原因)。
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

' ----------------------------------------------------------------------------
' ReapplyChatBound - チャット境界(ScrollArea)を今の幾何で貼り直す(R27 F2-1c)。
' ----------------------------------------------------------------------------
' RedrawChatHeader は行1の高さを HeaderHeight() で作り直す。境界はptから
' 行番号へ換算した結果(BoundAddr)なので、行1が1段ぶん伸縮すると同じptが
' 別の行を指し、境界が実下端とずれる。トグル3経路(速度/言語/モード)は
' ヘッダーだけを描き直して境界を更新しないため、そのぶん会話が境界の外へ
' はみ出していた。呼び出し側にwsを持たせないよう引数なしの窓口にする。
Public Sub ReapplyChatBound()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(NEXUS_SHEET)
    If Not ws Is Nothing Then modViewport.ApplyScrollBound ws, NexusBound(ws)
    On Error GoTo 0
End Sub

' ヘッダーが持つShapeだけを消す。入力欄の nx_top_add / nx_top_send は同じ
' 接頭辞だがヘッダーの持ち物ではない(消すと📎と送信が二度と戻らない)。
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

' 段番号→ピルのY座標(その段の中央)。
Private Function PillTop(ByVal rowIdx As Long) As Double
    PillTop = rowIdx * HDR_H + (HDR_H - HDR_BTN_H) / 2
End Function

' 1段目のピルのうち最も左のX(=タイトルが使える右端)。
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
'   幅はキャプションの実文字から出す(固定の予約幅が実物と食い違ったことが
'   重なりの原因だったため、予約という概念自体をやめる)。
' ----------------------------------------------------------------------------
' ティア方針(2026-07-30 レビュー3-A(2)):
'   状態を持つ3つ(モード/速度/言語)はどのティアでも短い文字ラベルを保つ。
'   アイコンだけにすると「今どちらか」が画面から消える。実機幅(約597pt)では
'   常に最短ティアが選ばれるため、落とすと【常に状態が読めない】ことになる。
'   無状態の5つ(終了/クリア/質問例/ヘルプ/テーマ)は意味が絵柄に載っている
'   ので先にこちらを落とす。入らない分は FlowRight が2段目へ流す。
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

' 短縮表記のピルへツールチップ(代替テキスト)で元の意味を添える。
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
    ' R24-2b: 白枠0.75ptは視認不可と実機判定。塗り白+濃緑文字+濃緑枠へ反転。
    btn.Line.Visible = -1
    btn.Line.ForeColor.RGB = RGB(1, 77, 68)
    btn.Line.Weight = 0.75
    btn.Fill.ForeColor.RGB = RGB(255, 255, 255)
    With btn.TextFrame2
        .TextRange.Text = caption
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(1, 77, 68)
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
    ' R24-2b: 白枠0.75ptは視認不可と実機判定。塗り白+濃緑文字+濃緑枠へ反転。
    th.Line.Visible = -1
    th.Line.ForeColor.RGB = RGB(1, 77, 68)
    th.Line.Weight = 0.75
    th.Fill.ForeColor.RGB = RGB(255, 255, 255)
    With th.TextFrame2
        .TextRange.Text = modSkin.ThemeIcon()
        .TextRange.Font.Size = 11
        .TextRange.Font.Fill.ForeColor.RGB = RGB(1, 77, 68)
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
        .NumberFormat = "@"               ' R28 W2-5: 「1-3」を日付に変えない
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

    ' 入力欄の左は「資料を入れる」。READMEが売っているのは「ぶちこんで、
    ' すぐ聞ける」の2つの動詞なのに、この画面には「聞く」しか無く、資料を
    ' 入れるには3画面移動が要った(初めて開いた人が辿り着けない)。送信の
    ' 真向かいに置いて2つの動詞を揃える。旧📎(クリップボード画像)は
    ' ナレッジ画面の「📸 スクショ取込」と同じ機能なのでそちらへ譲った。
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

    ' R19-1b: 送信ボタンの右端をヘッダー帯・ピルと揃える(実機要望)。
    ' 左端はL列の実セル幾何から、右端は ContentRight から取る。M列(右余白)が
    ' 受け皿になるので、数pt はみ出しても隣の列を割らない。
    Dim cellL As Range: Set cellL = ws.Range("L" & r)
    Dim sendH As Double: sendH = 32
    Dim sendL As Double: sendL = cellL.Left + 4
    Dim sendW As Double
    sendW = modViewport.ContentRight(ws, NEXUS_BAND, HDR_RIGHT_PAD) - sendL
    If sendW < 40 Then sendW = 40
    Dim send As Shape
    Set send = ws.Shapes.AddShape(5, sendL, rowTop + (rowH - sendH) / 2, sendW, sendH)
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

    ' 入力欄の下に小さなヒント(セル。Shapeを増やさない)。資料が1件も無い
    ' あいだは、ショートカットより先に「何をすれば使えるようになるか」を出す
    ' (空のときにショートカットを教えても押す先が無い)。
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

    ' R18-3b: 行ける範囲の宣言はここで行う(modUI は上限30,000字に対し残りが
    ' 無く1行も足せないため、Nexusの幾何を持つ本モジュール側に置く)。
    ' InitUI からは DrawInputArea → FreezePanes(A5選択)→ Protect の順に進む。
    ' A5 も入力欄 C3 も境界の内側なので、既存の Select は影響を受けない。
    modViewport.ApplyScrollBound ws, NexusBound(ws)
    ' R21-S7: フィット直後の5値観測点(実体は modViewport2.LogChat)。
    modViewport2.LogChat ws
End Sub

' 文脈アクション: 最新のAI回答バブルの直下にだけ6個のpillを出す。新しい質問の
' たびにClearContextActionsで消えるので、古い回答の下には残らない。
Public Sub DrawContextActions(ByVal ws As Worksheet, ByVal bubbleName As String)
    ClearContextActions ws
    If ws Is Nothing Then Exit Sub
    If LenB(bubbleName) = 0 Then Exit Sub

    Dim anchor As Shape
    On Error Resume Next
    Set anchor = ws.Shapes(bubbleName)
    On Error GoTo 0
    If anchor Is Nothing Then Exit Sub

    ' 評価ボタンは「信頼度バッジ」と「出典チップ」より必ず下に置く。根拠を
    ' 見る前に「解決した/違う」を押させる並びは、評価してから確かめろと
    ' 言うに等しい。損保でAIの回答を信じてよい理由は原文確認だけ。
    Dim y As Double: y = anchor.Top + anchor.Height + 6
    Dim cb As Double
    On Error Resume Next
    cb = ConfidenceBottom(ws)
    If cb + 6 > y Then y = cb + 6
    cb = modPeek.CitationsBottom(ws)
    If cb + 8 > y Then y = cb + 8
    On Error GoTo 0

    Dim caps As Variant, kinds As Variant, acts As Variant
    ' 評価は「解決した/微妙/違う」の3択。1クリックで完結し入力を強制しない
    ' (入力を求めた途端に誰も押さなくなる)。「本社照会」は削除した ――
    ' 常に「準備中です」としか返らないボタンが全回答の下に並んでいると、
    ' 利用者は「他も見せかけかもしれない」と学習する。
    ' R26-3: 7個目に💾保存(この会話を考察メモとして本棚へ)。実体は
    ' modInsightCard で、ここは配列に1要素足すだけ=描画・境界(SettleChatの
    ' 内側)・掃除(nx_act_接頭辞)の全てを既存の文法にそのまま乗せる。
    caps = Array(ChrW(&H2705) & " 解決した", _
                 ChrW(&HD83E) & ChrW(&HDD14) & " 微妙", _
                 ChrW(&H274C) & " 違う", _
                 ChrW(&HD83D) & ChrW(&HDD0D) & " 深掘り", _
                 ChrW(&HD83D) & ChrW(&HDCCB) & " コピー", _
                 ChrW(&HD83D) & ChrW(&HDCC4) & " Word", _
                 ChrW(&HD83D) & ChrW(&HDCBE) & " 保存")
    kinds = Array("resolve", "unsure", "bad", "drill", "copy", "word", "insight")
    acts = Array("OnActResolve", "OnActUnsure", "OnActBad", "OnActDrill", _
                 "OnActCopy", "OnActWord", "OnActSaveInsight")

    ' R11-B(#30): 固定幅6個の決め打ちをやめ、modChrome.PillWidth+FlowLeftで
    ' 実可視幅(modUIMain.ViewportWidth)基準に流し込む。入り切らない分は
    ' modChrome.FlowLeftが段を増やして下へ流す(画面外へ見切れない)。
    Dim widths(0 To 6) As Double
    Dim i As Long
    For i = 0 To 6
        widths(i) = modChrome.PillWidth(CStr(caps(i)), 13, 16, 40)
    Next i
    Dim maxX As Double
    maxX = ChatLeft(ws) + modChrome.BarWidth(ChatWidth(ws), modUIMain.ViewportWidth(), 8)

    Dim xs() As Double, rws() As Long, useW() As Double
    Dim rowN As Long
    rowN = modChrome.FlowLeft(widths, 7, anchor.Left, maxX, 5, xs, rws, useW)
    If rowN < 1 Then rowN = 1

    For i = 0 To 6
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
                .TextRange.Font.Size = 9.5   ' R12-7-5: 8.5pt→9.5pt(a11y監査Med)
                .TextRange.ParagraphFormat.Alignment = 2
                .VerticalAnchor = 3
                .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
            End With
            btn.OnAction = "modAppAct." & CStr(acts(i))
            btn.Placement = 3
            modSkin.PaintActionButton btn, CStr(kinds(i))
        End If
        Set btn = Nothing
        Err.Clear
        On Error GoTo 0
    Next i
End Sub

' ----------------------------------------------------------------------------
' DrawConfidence - 信頼度バッジ。回答バブルの直下、いちばん最初に読ませる1行。
' ----------------------------------------------------------------------------
' 「この答えを信じてよいか」を自分で判断できないことがフィードバックの
' 集まらない根本原因だった。検索スコアを、確認すべきときだけ確認を促す一文へ
' 翻訳して先に見せる。Shape名 nx_conf_ は評価ボタン(nx_act_)の掃除に
' 巻き込まれないため(描画順は バッジ → 出典 → 評価)。
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
    ' 俯瞰(章の要約から答えたターン)は出自を言う(R17H FA-2)。score=0 は維持。
    If modAskGlobal.WasGlobalTurn() Then confText = ChrW(&HD83D) & ChrW(&HDD2D) & " 章の要約に基づく回答(俯瞰)"

    On Error Resume Next
    Dim badge As Shape
    Set badge = ws.Shapes.AddShape(5, anchor.Left, anchor.Top + anchor.Height + 6, 330, 22)
    If Err.Number = 0 And Not badge Is Nothing Then
        badge.Name = "nx_conf_badge"
        badge.Adjustments(1) = 0.45
        badge.Line.Visible = 0
        ' 8pt・塗りなし・muted では一番大事な一文が一番目立たない字だった。
        ' 色の付いたピルにして先に目に入るようにする。
        badge.Fill.Visible = -1
        badge.Fill.ForeColor.RGB = ConfidenceTint(confText)
        With badge.TextFrame2
            .WordWrap = -1
            .TextRange.Text = confText
            .TextRange.Font.Size = 9
            .TextRange.Font.Bold = -1
            ' 2026-08-01(R12-7-1): テーマ追従のtext色だと dark/gold の明るい
            ' textが淡色tintに乗って比1.0台まで沈み、一番大事な一文が読めなく
            ' なる(a11y監査High)。tintは常に淡色固定なので文字側もテーマ非依存の
            ' 固定濃色にする(RGB(17,24,39)。3色とも比7以上を機械確認済み)。
            .TextRange.Font.Fill.ForeColor.RGB = RGB(17, 24, 39)
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

' 信頼度バッジの背景色。先頭の🟢/🟡/🔴 だけで決める(文言の持ち主は modAsk)。
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
