Attribute VB_Name = "modInsight"
Option Explicit

' ============================================================================
' modInsight - 共有知フライホイール(組織で自己成長させる仕組み)
' ----------------------------------------------------------------------------
' 何のためにあるか:
'   NotebookLMは「個人に閉じた」ナレッジツールで、誰かが調べて分かったことは
'   その人の中で終わる。このプロダクトの存在価値は、部の縦割りを越えて
'   「営業が困ったこと」と「商品部が持っている答え」を自動で行き来させること。
'   そのための最小で確実な2本の流れをこのモジュールが持つ。
'
'   [流れ1] 解決したQ&A → みんなの知識
'     ✅解決した が押された社内ナレッジ回答は「人間が正しいと確認済み」という
'     一次情報になる。これを共有フォルダへ発信し、他の人の本棚へ取り込めるように
'     する。同じ質問を次にした人は、いきなり検証済みの答えに当たる。
'
'   [流れ2] 答えられなかった質問 → 知識の穴リスト
'     根拠が薄い/回答が違うと言われた質問は「組織に文書が存在しない領域」の
'     一次情報になる。これを共有フォルダへ発信し、資料を書ける人(商品部)の
'     画面に一覧で出す。そこから直接ナレッジを書けば穴が塞がる。
'
'   この2本が回ると、使えば使うほど answers が増え gaps が減る。パック共有が
'   「まとめて手渡す」のに対し、こちらは日々の利用そのものが供給源になる。
'
' 設計判断:
'   ・共有フォルダI/Oは modP2P/modBoard と同じ「1人1ファイル・上書き禁止・
'     新規ファイル追記型」。書込競合が原理的に起きない実証済みの形を踏襲する。
'   ・受信はBoot時に自動(ファイルコピーだけ=API呼び出しなし=速い)。
'     本棚への取り込み(埋め込みAPIが要る重い処理)は利用者が押したときだけ。
'     起動が遅くなること自体が最大のストレスなので、そこは絶対に自動化しない。
'   ・受信済み判定は nonce(ファイル名)を my_stats に記録。二重取り込み不可。
'   ・自分が出した情報は自分では取り込まない(自作自演のループ防止)。
'   ・全Publicエントリの先頭で On Error Resume Next。共有フォルダが無い・
'     権限が無い環境でも、本体機能には一切影響させない。
' ============================================================================

Private Const INSIGHT_SUBDIR As String = "insight"
Private Const QA_SUBDIR As String = "qa"
Private Const GAP_SUBDIR As String = "gap"
Private Const SHEET_NAME As String = "insight_inbox"
Private Const MAX_COLLECT As Long = 60      ' 1回の起動で読むファイル数の上限
Private Const FIELD_SEP As String = vbTab

' ----------------------------------------------------------------------------
' 発信: 解決済みQ&A(✅解決したの発火点から呼ばれる)
' ----------------------------------------------------------------------------
Public Sub EmitVerifiedQA(ByVal q As String, ByVal ans As String, ByVal src As String)
    On Error Resume Next
    If LenB(Trim$(q)) = 0 Or LenB(Trim$(ans)) = 0 Then Exit Sub

    Dim dirPath As String: dirPath = SubDir(QA_SUBDIR)
    If LenB(dirPath) = 0 Then Exit Sub
    EnsureDir dirPath

    Dim myId As String: myId = SafeUserId()
    If LenB(myId) = 0 Then Exit Sub

    Dim body As String
    body = "v1" & FIELD_SEP & myId & FIELD_SEP & AuthorName() & FIELD_SEP & _
           Format$(Now, "yyyy-mm-dd hh:nn") & FIELD_SEP & _
           Clean1(q) & FIELD_SEP & Clean1(ans) & FIELD_SEP & Clean1(src)

    WriteShared dirPath & MakeNonce(myId) & ".txt", body
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 発信: 答えられなかった質問(根拠が薄い/違うと言われた)
'   reason: "no_hit"(本棚に無い) / "low_conf"(根拠が薄い) / "wrong"(回答が違う)
' ----------------------------------------------------------------------------
Public Sub EmitGap(ByVal q As String, ByVal reason As String)
    On Error Resume Next
    If LenB(Trim$(q)) = 0 Then Exit Sub

    Dim dirPath As String: dirPath = SubDir(GAP_SUBDIR)
    If LenB(dirPath) = 0 Then Exit Sub
    EnsureDir dirPath

    Dim myId As String: myId = SafeUserId()
    If LenB(myId) = 0 Then Exit Sub

    Dim body As String
    body = "v1" & FIELD_SEP & myId & FIELD_SEP & AuthorName() & FIELD_SEP & _
           Format$(Now, "yyyy-mm-dd hh:nn") & FIELD_SEP & _
           Clean1(q) & FIELD_SEP & reason & FIELD_SEP & DeptName()

    WriteShared dirPath & MakeNonce(myId) & ".txt", body
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 発信: 修正内容(誰かが「違う」を押して正しい内容を書いてくれた)
'   1人の訂正を全員の訂正にする。ただし個人の申告が即座に正になるのは危険なので、
'   受け取り側では「参考情報」として扱い、同じ趣旨の訂正が複数集まったものだけを
'   商品部が正式なナレッジに昇格させる運用を前提にしている(CorrectionAgreeCount)。
' ----------------------------------------------------------------------------
Public Sub EmitCorrection(ByVal answerText As String, ByVal fixText As String)
    On Error Resume Next
    If LenB(Trim$(fixText)) = 0 Then Exit Sub

    Dim dirPath As String: dirPath = SubDir(GAP_SUBDIR)
    If LenB(dirPath) = 0 Then Exit Sub
    EnsureDir dirPath

    Dim myId As String: myId = SafeUserId()
    If LenB(myId) = 0 Then Exit Sub

    Dim body As String
    body = "v1" & FIELD_SEP & myId & FIELD_SEP & AuthorName() & FIELD_SEP & _
           Format$(Now, "yyyy-mm-dd hh:nn") & FIELD_SEP & _
           Clean1("【訂正】" & modUtil.SafeLeft(answerText, 200)) & FIELD_SEP & _
           "correction" & FIELD_SEP & Clean1(fixText)

    WriteShared dirPath & MakeNonce(myId) & ".txt", body
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 受信: 共有フォルダの新着をinsight_inboxへ取り込む(APIは呼ばない=速い)。
'   戻り値 = 新しく受け取った件数。
' ----------------------------------------------------------------------------
Public Function CollectInsights() As Long
    On Error Resume Next
    Dim ws As Worksheet: Set ws = EnsureSheet()
    If ws Is Nothing Then Exit Function
    CollectInsights = CollectFrom(ws, SubDir(QA_SUBDIR), "qa") + _
                      CollectFrom(ws, SubDir(GAP_SUBDIR), "gap")
    On Error GoTo 0
End Function

' 未取り込みの解決済みQ&A件数(Hubのお知らせに出す)。
Public Function PendingQACount() As Long
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet()
    If ws Is Nothing Then Exit Function
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastR
        If CStr(ws.Cells(r, 2).Value) = "qa" Then
            If CStr(ws.Cells(r, 9).Value) <> "1" Then PendingQACount = PendingQACount + 1
        End If
    Next r
    On Error GoTo 0
End Function

' 未解決質問(gap)の件数。
Public Function GapCount() As Long
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet()
    If ws Is Nothing Then Exit Function
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastR
        If CStr(ws.Cells(r, 2).Value) = "gap" Then
            If CStr(ws.Cells(r, 9).Value) <> "1" Then GapCount = GapCount + 1
        End If
    Next r
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' 未取り込みQ&Aの取り出し(UI層が1件ずつ引いて本棚へ登録する)。
'   idx は 1..PendingQACount()。取得できたら True。
'   本棚登録そのものは埋め込みAPIを使う重い処理なので、いつ実行するかは
'   UI層(利用者のクリック)が決める。ここは箱の役に徹する。
' ----------------------------------------------------------------------------
Public Function PendingQAAt(ByVal idx As Long, ByRef outAuthor As String, _
                            ByRef outQ As String, ByRef outA As String, _
                            ByRef outSrc As String, ByRef outRow As Long) As Boolean
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet()
    If ws Is Nothing Then Exit Function
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim seen As Long, r As Long
    For r = 2 To lastR
        If CStr(ws.Cells(r, 2).Value) = "qa" And CStr(ws.Cells(r, 9).Value) <> "1" Then
            seen = seen + 1
            If seen = idx Then
                outAuthor = CStr(ws.Cells(r, 4).Value)
                outQ = CStr(ws.Cells(r, 6).Value)
                outA = CStr(ws.Cells(r, 7).Value)
                outSrc = CStr(ws.Cells(r, 8).Value)
                outRow = r
                PendingQAAt = True
                Exit Function
            End If
        End If
    Next r
    On Error GoTo 0
End Function

' 選択状態(J列)。選択式取り込みのために行単位で持つ。
Public Function IsSelected(ByVal rowIdx As Long) As Boolean
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet()
    If ws Is Nothing Then Exit Function
    IsSelected = (CStr(ws.Cells(rowIdx, 10).Value) = "1")
    On Error GoTo 0
End Function

Public Sub ToggleSelected(ByVal rowIdx As Long)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub
    If rowIdx < 2 Then Exit Sub
    If CStr(ws.Cells(rowIdx, 10).Value) = "1" Then
        ws.Cells(rowIdx, 10).Value = ""
    Else
        ws.Cells(rowIdx, 10).Value = "1"
    End If
    On Error GoTo 0
End Sub

Public Sub SelectAllPending(ByVal onOff As Boolean)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastR
        If CStr(ws.Cells(r, 2).Value) = "qa" And CStr(ws.Cells(r, 9).Value) <> "1" Then
            ws.Cells(r, 10).Value = IIf(onOff, "1", "")
        End If
    Next r
    On Error GoTo 0
End Sub

Public Function SelectedCount() As Long
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet()
    If ws Is Nothing Then Exit Function
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastR
        If CStr(ws.Cells(r, 2).Value) = "qa" And CStr(ws.Cells(r, 9).Value) <> "1" _
           And CStr(ws.Cells(r, 10).Value) = "1" Then SelectedCount = SelectedCount + 1
    Next r
    On Error GoTo 0
End Function

' 未取り込みQ&Aの行番号を「同じ質問が多い順」に並べて返す。
'   12000人規模では同じ質問が重なる。件数の多い=多くの人が困っているものから
'   上に出すことで、全部を読まなくても価値の高いものだけ拾える。
Public Function PendingRowsRanked(ByRef outRows() As Long) As Long
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet()
    If ws Is Nothing Then Exit Function
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Function

    ReDim outRows(0 To lastR)
    Dim scores() As Long: ReDim scores(0 To lastR)
    Dim n As Long, r As Long
    For r = 2 To lastR
        If CStr(ws.Cells(r, 2).Value) = "qa" And CStr(ws.Cells(r, 9).Value) <> "1" Then
            outRows(n) = r
            scores(n) = SameQuestionCount(CStr(ws.Cells(r, 6).Value))
            n = n + 1
        End If
    Next r
    If n = 0 Then Exit Function

    ' 件数の降順(nは高々数百なので単純なバブルで足りる)
    Dim a As Long, b As Long
    For a = 0 To n - 2
        For b = 0 To n - 2 - a
            If scores(b) < scores(b + 1) Then
                Dim ts As Long: ts = scores(b): scores(b) = scores(b + 1): scores(b + 1) = ts
                Dim tr As Long: tr = outRows(b): outRows(b) = outRows(b + 1): outRows(b + 1) = tr
            End If
        Next b
    Next a
    PendingRowsRanked = n
    On Error GoTo 0
End Function

' 行番号から表示用の値を取り出す。
Public Function RowField(ByVal rowIdx As Long, ByVal col As Long) As String
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet()
    If ws Is Nothing Then Exit Function
    RowField = CStr(ws.Cells(rowIdx, col).Value)
    On Error GoTo 0
End Function

' 取り込み済みの印(二重登録を防ぐ)。
Public Sub MarkQAConsumed(ByVal rowIdx As Long)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub
    If rowIdx >= 2 Then ws.Cells(rowIdx, 9).Value = "1"
    On Error GoTo 0
End Sub

' 同じ趣旨の質問が何件届いているか(重複統合の要)。
'   12000人規模では「同じ質問」が大量に重なる。1件ずつ見せるとお知らせ欄が
'   ただの流れるログになり、既存の教えてBOXと同じ末路をたどる。件数の多い
'   ものから見せることで、読む価値のある順に並ぶ。
'   照合は「記号と空白を落とした先頭40字の一致」。形態素解析が使えないVBAで
'   実務上いちばん誤爆が少なかった近似。
Public Function SameQuestionCount(ByVal qText As String) As Long
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet()
    If ws Is Nothing Then Exit Function
    Dim key As String: key = NormKey(qText)
    If LenB(key) = 0 Then Exit Function

    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastR
        If NormKey(CStr(ws.Cells(r, 6).Value)) = key Then SameQuestionCount = SameQuestionCount + 1
    Next r
    On Error GoTo 0
End Function

Private Function NormKey(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace(t, " ", ""): t = Replace(t, ChrW(&H3000), "")
    t = Replace(t, ChrW(&H3001), ""): t = Replace(t, ChrW(&H3002), "")
    t = Replace(t, "?", ""): t = Replace(t, ChrW(&HFF1F), "")
    t = Replace(t, ChrW(&HFF08), ""): t = Replace(t, ChrW(&HFF09), "")
    NormKey = LCase$(modUtil.SafeLeft(t, 40))
End Function

' 取り込む本文の組み立て(表現をここに集約し、UI層は登録するだけにする)。
Public Function QABodyText(ByVal author As String, ByVal qText As String, _
                           ByVal aText As String, ByVal srcText As String) As String
    Dim body As String
    body = "【この内容は " & author & " さんが実務で確認し、解決済みとした回答です】" & vbLf & vbLf & _
           "■ 質問" & vbLf & qText & vbLf & vbLf & _
           "■ 確認済みの回答" & vbLf & aText
    If LenB(srcText) > 0 Then body = body & vbLf & vbLf & "■ もとになった資料" & vbLf & srcText
    QABodyText = body
End Function

' ----------------------------------------------------------------------------
' 内部: 共有フォルダ→シート
' ----------------------------------------------------------------------------
Private Function CollectFrom(ByVal ws As Worksheet, ByVal dirPath As String, _
                             ByVal kind As String) As Long
    If LenB(dirPath) = 0 Then Exit Function
    On Error Resume Next
    If Len(Dir(dirPath, vbDirectory)) = 0 Then Exit Function

    Dim myId As String: myId = SafeUserId()
    Dim fname As String
    fname = Dir(dirPath & "*.txt")

    Dim guard As Long
    Do While LenB(fname) > 0 And guard < MAX_COLLECT
        guard = guard + 1
        Dim nc As String: nc = Left$(fname, Len(fname) - 4)
        ' 自分が出したものは取り込まない(自作自演の循環を作らない)。
        If InStr(1, nc, myId, vbTextCompare) <> 1 Then
            If modStats.GetStat("ins:" & nc) = 0 Then
                Dim raw As String
                If ReadShared(dirPath & fname, raw) Then
                    If AppendRow(ws, kind, nc, raw) Then CollectFrom = CollectFrom + 1
                End If
                modStats.Bump "ins:" & nc
            End If
        End If
        fname = Dir()
    Loop
    On Error GoTo 0
End Function

' insight_inbox の列: nonce / kind / user_id / author / created_at /
'                     question / answer / source / consumed
Private Function AppendRow(ByVal ws As Worksheet, ByVal kind As String, _
                           ByVal nc As String, ByVal raw As String) As Boolean
    Dim f() As String
    f = Split(raw, FIELD_SEP)
    If UBound(f) < 5 Then Exit Function

    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = nc
    ws.Cells(r, 2).Value = kind
    ws.Cells(r, 3).Value = f(1)
    ws.Cells(r, 4).Value = f(2)
    ws.Cells(r, 5).Value = f(3)
    ws.Cells(r, 6).Value = f(4)
    If kind = "qa" Then
        ws.Cells(r, 7).Value = f(5)
        If UBound(f) >= 6 Then ws.Cells(r, 8).Value = f(6)
    Else
        ws.Cells(r, 7).Value = f(5)                       ' reason
        If UBound(f) >= 6 Then ws.Cells(r, 8).Value = f(6)   ' 部署
    End If
    ws.Cells(r, 9).Value = ""
    AppendRow = True
End Function

' ----------------------------------------------------------------------------
' 内部: 共有フォルダI/O(modBoard.WriteBeacon と同仕様の意図的な複製。
'   モジュール完全隔離のため、他モジュールのPrivateには依存しない)
' ----------------------------------------------------------------------------
Private Function WriteShared(ByVal filePath As String, ByVal content As String) As Boolean
    Dim attempt As Long
    For attempt = 1 To 3
        If TryWrite(filePath, content) Then
            WriteShared = True
            Exit Function
        End If
        ShortWait 200 * attempt
    Next attempt
End Function

Private Function TryWrite(ByVal filePath As String, ByVal content As String) As Boolean
    Dim st As Object
    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText content
    st.SaveToFile filePath, 2
    st.Close
    Set st = Nothing
    TryWrite = True
    Exit Function
Fail:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Set st = Nothing
    On Error GoTo 0
End Function

Private Function ReadShared(ByVal filePath As String, ByRef outText As String) As Boolean
    Dim st As Object
    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile filePath
    outText = CStr(st.ReadText(-1))
    st.Close
    Set st = Nothing
    ReadShared = True
    Exit Function
Fail:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Set st = Nothing
    On Error GoTo 0
End Function

Private Sub ShortWait(ByVal ms As Long)
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000# < ms
        DoEvents
        If Timer < t0 Then Exit Do
    Loop
End Sub

' ----------------------------------------------------------------------------
' 内部: 小物
' ----------------------------------------------------------------------------
Private Function SubDir(ByVal leaf As String) As String
    Dim basePath As String
    On Error Resume Next
    basePath = modConfig.GetString("nexus_share_path", "")
    On Error GoTo 0
    If LenB(basePath) = 0 Then Exit Function
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
    EnsureDir basePath & INSIGHT_SUBDIR & "\"
    SubDir = basePath & INSIGHT_SUBDIR & "\" & leaf & "\"
End Function

Private Sub EnsureDir(ByVal folderPath As String)
    On Error Resume Next
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    On Error GoTo 0
End Sub

Private Function MakeNonce(ByVal myId As String) As String
    MakeNonce = myId & "-" & Format$(Now, "yyyymmddhhnnss") & "-" & _
            Format$(Int(Timer * 1000) Mod 100000, "00000")
End Function

Private Function SafeUserId() As String
    On Error Resume Next
    SafeUserId = modP2P.CurrentUserId()
    On Error GoTo 0
End Function

Private Function AuthorName() As String
    On Error Resume Next
    AuthorName = Trim$(modConfig.GetString("pack_author", ""))
    On Error GoTo 0
    If LenB(AuthorName) = 0 Or AuthorName = "名称未設定" Then AuthorName = "どなたか"
End Function

Private Function DeptName() As String
    On Error Resume Next
    DeptName = Trim$(modConfig.GetString("user_department", ""))
    On Error GoTo 0
End Function

' タブ・改行はフィールド区切りと衝突するため潰す。
Private Function Clean1(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace(t, vbTab, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    Clean1 = modUtil.SafeLeft(Trim$(t), 2000)
End Function

Private Function GetSheet() As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
End Function

Private Function EnsureSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = GetSheet()
    If Not ws Is Nothing Then
        Set EnsureSheet = ws
        Exit Function
    End If
    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = SHEET_NAME
    ws.Range("A1:I1").Value = Array("nonce", "kind", "user_id", "author", "created_at", _
                                    "question", "answer_or_reason", "source_or_dept", "consumed")
    ws.Visible = 0    ' xlSheetHidden
    Set EnsureSheet = ws
    Exit Function
Fail:
End Function

' ----------------------------------------------------------------------------
' 「みんなの困りごと」ボード: 組織で答えが見つからなかった質問の一覧。
'   資料を書ける人(商品部)がここから直接ナレッジを登録できる導線を持つ。
'   これが部の縦割りを越える一番短い経路。
' ----------------------------------------------------------------------------
Public Function GapListText() As String
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet()
    If ws Is Nothing Then
        GapListText = "(まだ届いていません)"
        Exit Function
    End If

    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim sb As String, n As Long
    Dim r As Long
    For r = lastR To 2 Step -1
        If CStr(ws.Cells(r, 2).Value) = "gap" Then
            n = n + 1
            If n > 20 Then Exit For
            Dim who As String: who = CStr(ws.Cells(r, 4).Value)
            Dim dept As String: dept = CStr(ws.Cells(r, 8).Value)
            If LenB(dept) > 0 Then who = who & "/" & dept
            sb = sb & n & ". " & CStr(ws.Cells(r, 6).Value) & vbLf & _
                 "     (" & who & " ・ " & CStr(ws.Cells(r, 5).Value) & _
                 " ・ " & ReasonText(CStr(ws.Cells(r, 7).Value)) & ")" & vbLf
        End If
    Next r

    If n = 0 Then
        GapListText = "まだ届いていません。" & vbLf & _
            "誰かが質問して本棚に答えが無かったとき、その質問がここに自動で並びます。"
    Else
        GapListText = sb
    End If
    On Error GoTo 0
End Function

Private Function ReasonText(ByVal code As String) As String
    Select Case code
        Case "no_hit":   ReasonText = "本棚に該当資料なし"
        Case "low_conf": ReasonText = "根拠が薄い"
        Case "wrong":    ReasonText = "回答が違うと報告"
        Case Else:       ReasonText = code
    End Select
End Function
