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

' R14-7a: シード0でも my_knowledge に実データがあれば、質問例をLLMで
' オンデマンド生成する(1回だけ・キャッシュ)。取込内容が変わらない限り
' 再生成しない(srcsigが一致する間はキャッシュを表示するだけ)。
Private Const QCACHE_KEY As String = "my_questions_cache"    ' "|"区切りの質問(SeedQuestionsと同形式)
Private Const QSRCSIG_KEY As String = "my_questions_srcsig"  ' 生成時点の資料の状態の指紋
Private Const QGEN_MAXSRC As Long = 8      ' プロンプトへ渡す資料数の上限
Private Const QGEN_PREVIEW_CHARS As Long = 200  ' 資料1件あたりの抜粋の長さ
Private Const QGEN_MAX_Q As Long = 5       ' 生成させる質問数

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
    all = QuestionPool()
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
    If Left$(caller, Len(PREFIX)) <> PREFIX Then
        modLog.LogUsage "caller_mismatch", "modStarter.OnPick", caller
        Exit Sub
    End If

    Dim tail As String: tail = Mid$(caller, Len(PREFIX) + 1)
    If tail = "more" Then Exit Sub

    Dim all As String
    On Error Resume Next
    all = QuestionPool()
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
' R14-7a: シード0でも my_knowledge に実データがあれば、案内で終わらせず
' LLMでオンデマンド生成する(キャッシュが有効ならそれを表示するだけ)。
Public Sub OnShowList()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUI.GoToNexus "modStarter.OnShowList"
    Dim n As Long: n = modSeed.SeedDocCount()
    If n > 0 Then
        modUI.AddChatBubble "ai", _
            ChrW(&HD83D) & ChrW(&HDCA1) & " この端末に入っている " & n & "つの資料から、" & _
            "そのまま聞ける質問です。押すだけで答えます。"
        Draw
    ElseIf modShelf.TotalChunks() = 0 Then
        modUI.AddChatBubble "ai", _
            ChrW(&HD83D) & ChrW(&HDCA1) & " すぐ押せる質問はまだありません。" & vbLf & _
            "資料を取り込むと、ここに質問例が並びます。"
    Else
        ShowQuestionsFromKnowledge
    End If
    On Error GoTo 0
    modUiLock.Leave
End Sub

' ----------------------------------------------------------------------------
' ShowQuestionsFromKnowledge - シード0・資料ありのときの質問例(R14-7a)。
'   キャッシュが有効ならそのまま表示。無効ならLLMを1回だけ呼んで生成し、
'   キャッシュへ保存してから表示する。失敗時は従来の空メッセージへ倒す
'   (キャッシュは書かない=次回また生成を試す)。
' ----------------------------------------------------------------------------
Private Sub ShowQuestionsFromKnowledge()
    Dim sig As String: sig = KnowledgeSignature()
    Dim cached As String: cached = modState.LoadState(QCACHE_KEY, "")

    If LenB(cached) > 0 And LenB(sig) > 0 And sig = modState.LoadState(QSRCSIG_KEY, "") Then
        modUI.AddChatBubble "ai", _
            ChrW(&HD83D) & ChrW(&HDCA1) & " お使いの資料から、そのまま聞ける質問です。押すだけで答えます。"
        Draw
        Exit Sub
    End If

    modUIMain.SetStage ChrW(&HD83D) & ChrW(&HDCA1) & " 質問例を作成中…"

    Dim digest As String: digest = RecentSourceDigest(QGEN_MAXSRC, QGEN_PREVIEW_CHARS)
    Dim qs As String
    If LenB(digest) > 0 Then
        Dim prompt As String: prompt = modPrompts.BuildQuestionsPrompt(digest)
        Dim lat As Long
        Dim resp As String
        resp = modGateway.CallLLM(prompt, "seed_questions", _
            modConfig.GetString("quick_effort", "low"), modConfig.GetString("quick_verbosity", "low"), _
            "", lat)
        If Left$(resp, 5) <> "#ERR:" Then
            qs = modRagParse.ParseQuestionLines(resp, QGEN_MAX_Q)
        End If
        If LenB(qs) = 0 Then
            modLog.LogError "E0202", "modStarter.OnShowList", _
                "質問例の生成に失敗: " & modUtil.SafeLeft(resp, 300)
        End If
    End If

    If LenB(qs) = 0 Then
        modUI.AddChatBubble "ai", _
            ChrW(&HD83D) & ChrW(&HDCA1) & " すぐ押せる質問はまだありません。" & vbLf & _
            "資料を取り込むと、ここに質問例が並びます。"
        Exit Sub
    End If

    modState.SaveState QCACHE_KEY, qs
    modState.SaveState QSRCSIG_KEY, sig
    modUI.AddChatBubble "ai", _
        ChrW(&HD83D) & ChrW(&HDCA1) & " お使いの資料から、そのまま聞ける質問です。押すだけで答えます。"
    Draw
End Sub

' ----------------------------------------------------------------------------
' QuestionPool - Drawが実際に描く質問の唯一の取得口(R14-7a)。
'   同梱シードの質問があればそれを優先し、無ければ(シード0ビルド)
'   オンデマンド生成でキャッシュした質問へフォールバックする。Draw自体は
'   「どちらの由来か」を意識しない(pipe区切りの形式が同じであるだけ)。
' ----------------------------------------------------------------------------
Private Function QuestionPool() As String
    Dim s As String: s = modSeed.SeedQuestions()
    If LenB(s) > 0 Then
        QuestionPool = s
    Else
        QuestionPool = modState.LoadState(QCACHE_KEY, "")
    End If
End Function

' 資料の状態の指紋(my_manifestの件数+直近処理時刻の最大値)。取込・削除・
' 再取込のいずれでも変わるので、キャッシュの有効/無効判定に使える
' (7b: 取込側の変更は不要。ここが自然に変わるだけで失効する)。
Private Function KnowledgeSignature() As String
    On Error Resume Next
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Worksheets(modAppDef.SH_MANIFEST)
    If wsM Is Nothing Then Exit Function
    Dim lastM As Long: lastM = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
    If lastM < 2 Then Exit Function

    ' 列2(fileName)～列8(処理時刻)を一括読み(複数列なので1行でも2次元配列。
    ' 単一セル範囲のスカラー化を避けるための定石)。
    Dim arr As Variant
    arr = wsM.Range(wsM.Cells(2, 2), wsM.Cells(lastM, 8)).Value
    Dim maxStamp As String
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim s As String: s = CStr(arr(i, 7))   ' 列8 = 配列オフセット7
        If s > maxStamp Then maxStamp = s
    Next i
    KnowledgeSignature = (lastM - 1) & "_" & maxStamp
    On Error GoTo 0
End Function

' 直近maxN件の資料名+その先頭チャンクの抜粋を「資料名: 抜粋」の行へまとめる。
' modShelf.SourceList(既存・manifest由来)は末尾ほど新規追加に近いので、
' 末尾から拾う(厳密な最終更新日時ソートまではしない=質問例の材料として
' 十分な精度)。
Private Function RecentSourceDigest(ByVal maxN As Long, ByVal maxCharsEach As Long) As String
    On Error Resume Next
    Dim names() As String, stats() As String
    Dim total As Long: total = modShelf.SourceList(names, stats)
    If total = 0 Then Exit Function

    Dim pickCount As Long: pickCount = total
    If pickCount > maxN Then pickCount = maxN
    Dim pick() As String: ReDim pick(0 To pickCount - 1)
    Dim i As Long
    For i = 0 To pickCount - 1
        pick(i) = names(total - 1 - i)
    Next i

    Dim previews() As String: ReDim previews(0 To pickCount - 1)
    LoadFirstChunkPreviews pick, pickCount, maxCharsEach, previews

    Dim sb As String
    For i = 0 To pickCount - 1
        If LenB(previews(i)) > 0 Then
            sb = sb & ChrW(&H30FB) & pick(i) & ": " & previews(i) & vbLf
        End If
    Next i
    RecentSourceDigest = sb
    On Error GoTo 0
End Function

' my_knowledgeを一括読みし、指定した資料名(最大QGEN_MAXSRC件)それぞれの
' 最初に見つかったチャンクの本文冒頭を拾う(EnsurePreviewIndexと同型の
' 「1回の一括読みで済ませる」作法。呼び出しは生成が要るときだけなので
' 頻度は低い)。
Private Sub LoadFirstChunkPreviews(ByRef srcNames() As String, ByVal n As Long, _
                                   ByVal maxCharsEach As Long, ByRef outPreviews() As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then Exit Sub
    Dim lastK As Long: lastK = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then Exit Sub

    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, 2), ws.Cells(lastK, 7)).Value  ' 2=source .. 7=full_text
    Dim found() As Boolean: ReDim found(0 To n - 1)
    Dim foundCount As Long: foundCount = 0

    Dim i As Long, j As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If foundCount >= n Then Exit For
        Dim nm As String: nm = CStr(arr(i, 1))
        For j = 0 To n - 1
            If Not found(j) Then
                If StrComp(nm, srcNames(j), vbTextCompare) = 0 Then
                    outPreviews(j) = modUtil.SafeLeft(CStr(arr(i, 6)), maxCharsEach)
                    found(j) = True
                    foundCount = foundCount + 1
                    Exit For
                End If
            End If
        Next j
    Next i
    On Error GoTo 0
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
