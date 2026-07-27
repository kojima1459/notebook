Attribute VB_Name = "modStarter"
Option Explicit

' ============================================================================
' modStarter - すぐ押せる質問(初期ナレッジの案内 兼 実用ランチャー)
' ----------------------------------------------------------------------------
' なぜ「資料名を並べる」だけでは足りないか:
'   「6つの資料が入っています。聞いてみてください」と案内しても、利用者は
'   まだ "何を聞くか" を自分で考えないといけない。しかも課長の隣で的外れな
'   質問をする気まずさも負う。初手で一瞬でも迷いが生まれたら、その人は
'   二度と開かない。
'   だから質問そのものを用意して、クリックだけで送信できるようにする。
'   考える量をゼロにするのが目的。
'
' なぜ嘘をつかないか:
'   質問はビルド時に、実際に同梱した資料の見出し(第N条(…)・章見出し)から
'   生成している(tools/make_seed_pack.py)。つまり全ての質問について、
'   答えが載っているチャンクが本棚に必ず存在する。
'   人が手で書くと「資料に無いこと」を聞いてしまい、初手で外す。
'
' デモで終わらせない:
'   同じ一覧をヘッダーの「質問例」からいつでも呼び出せる。初日は案内として、
'   翌日からは実際の調べもののショートカットとして働く。
'   1画面に出すのは6件。多すぎる選択肢はそれ自体が認知負荷なので、
'   「別の質問」で入れ替える方式にした(全件は保持している)。
' ============================================================================

Private Const PREFIX As String = "nx_sq_"
Private Const PAGE_SIZE As Long = 6
Private Const BTN_H As Double = 24

Private mOffset As Long        ' 何件目から表示しているか(「別の質問」で進む)

' ----------------------------------------------------------------------------
' Draw - チャットの現在位置に質問ボタンを並べる。
'   質問が無いビルド(シード未同梱)では何も描かずに戻る。
' ----------------------------------------------------------------------------
Public Sub Draw()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub

    Clear

    Dim all As String
    all = modSeed.SeedQuestions()
    If LenB(all) = 0 Then Exit Sub

    Dim qs() As String: qs = Split(all, "|")
    Dim total As Long: total = UBound(qs) - LBound(qs) + 1
    If total <= 0 Then Exit Sub
    If mOffset >= total Then mOffset = 0

    Dim L As Double: L = modUINexusDraw.ChatLeft(ws)
    Dim W As Double: W = modUINexusDraw.ChatWidth(ws) * 0.86
    Dim y As Double: y = modUI.ChatBottomFor(ws) + 10

    Dim shown As Long
    Dim i As Long
    For i = 0 To PAGE_SIZE - 1
        Dim idx As Long: idx = (mOffset + i) Mod total
        Dim q As String: q = Trim$(qs(idx))
        If LenB(q) > 0 Then
            Dim b As Shape
            Set b = ws.Shapes.AddShape(5, L, y, W, BTN_H)
            If Err.Number = 0 And Not b Is Nothing Then
                b.Name = PREFIX & CStr(idx)
                b.Adjustments(1) = 0.4
                b.Line.Visible = -1
                b.Line.Weight = 0.75
                b.Line.ForeColor.RGB = modUI.UiColor("primary")
                b.Fill.ForeColor.RGB = modUI.UiColor("surface")
                With b.TextFrame2
                    .WordWrap = -1
                    .TextRange.Text = ChrW(&H25B8) & " " & q
                    .TextRange.Font.Name = "Yu Gothic UI"
                    .TextRange.Font.Size = 9
                    .TextRange.ParagraphFormat.Alignment = 1
                    .VerticalAnchor = 3
                    .MarginLeft = 10: .MarginRight = 6: .MarginTop = 0: .MarginBottom = 0
                End With
                b.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("primary")
                b.OnAction = "modStarter.OnPick"
                b.Placement = 3
                y = y + BTN_H + 4
                shown = shown + 1
            End If
            Set b = Nothing
            Err.Clear
        End If
    Next i

    If shown = 0 Then Exit Sub

    ' 「別の質問」。全件を一度に出すと選択肢が多すぎて逆に止まるので、
    ' 入れ替え式にして、常に画面には6件だけ置く。
    If total > PAGE_SIZE Then
        Dim more As Shape
        Set more = ws.Shapes.AddShape(5, L, y + 2, 150, BTN_H - 2)
        If Err.Number = 0 And Not more Is Nothing Then
            more.Name = PREFIX & "more"
            more.Adjustments(1) = 0.45
            more.Line.Visible = 0
            more.Fill.Visible = 0
            With more.TextFrame2
                .TextRange.Text = ChrW(&HD83D) & ChrW(&HDD04) & " 別の質問を見る（全" & total & "件）"
                .TextRange.Font.Name = "Yu Gothic UI"
                .TextRange.Font.Size = 8.5
                .VerticalAnchor = 3
                .MarginLeft = 4: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
            End With
            more.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
            more.OnAction = "modStarter.OnMore"
            more.Placement = 3
        End If
        Set more = Nothing
        Err.Clear
    End If

    modUI.SettleChat
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' OnPick - 質問ボタンのクリック。入力欄へ入れて、そのまま送信する。
'   「入れるだけ」で止めると、利用者はもう一度送信を押す必要があり、
'   そこで手が止まる。押した=聞きたい、なので最後まで通す。
' ----------------------------------------------------------------------------
Public Sub OnPick()
    Dim caller As String
    On Error Resume Next
    caller = CStr(Application.Caller)
    On Error GoTo 0
    If Left$(caller, Len(PREFIX)) <> PREFIX Then Exit Sub

    Dim tail As String: tail = Mid$(caller, Len(PREFIX) + 1)
    If tail = "more" Then Exit Sub

    Dim all As String
    On Error Resume Next
    all = modSeed.SeedQuestions()
    On Error GoTo 0
    If LenB(all) = 0 Then Exit Sub

    Dim qs() As String: qs = Split(all, "|")
    Dim idx As Long: idx = CLng(Val(tail))
    If idx < LBound(qs) Or idx > UBound(qs) Then Exit Sub

    Clear
    On Error Resume Next
    ThisWorkbook.Names("nx_input").RefersToRange.Value = qs(idx)
    On Error GoTo 0

    modApp.OnSend
End Sub

' 「別の質問を見る」。6件ずつ送って一巡する。
Public Sub OnMore()
    mOffset = mOffset + PAGE_SIZE
    Draw
End Sub

' ヘッダーの「質問例」。初日は案内、翌日からは調べもののショートカット。
Public Sub OnShowList()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUI.GoToNexus "modStarter.OnShowList"
    Dim n As Long: n = modSeed.SeedDocCount()
    If n > 0 Then
        modUI.AddChatBubble "ai", _
            ChrW(&HD83D) & ChrW(&HDCA1) & " この端末に入っている " & n & "つの資料から、" & _
            "そのまま聞ける質問です。押すだけで答えます。"
    Else
        modUI.AddChatBubble "ai", _
            ChrW(&HD83D) & ChrW(&HDCA1) & " すぐ押せる質問はまだありません。" & vbLf & _
            "資料を取り込むと、ここに質問例が並びます。"
    End If
    Draw
    On Error GoTo 0
    modUiLock.Leave
End Sub

' ----------------------------------------------------------------------------
' Clear - 質問ボタンを全部消す。送信時・画面リセット時に呼ぶ。
' ----------------------------------------------------------------------------
Public Sub Clear()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub

    Dim names() As String
    ReDim names(0 To ws.Shapes.count)
    Dim n As Long
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, Len(PREFIX)) = PREFIX Then
            names(n) = shp.Name
            n = n + 1
        End If
    Next shp
    Dim i As Long
    For i = 0 To n - 1
        ws.Shapes(names(i)).Delete
    Next i
    On Error GoTo 0
End Sub
