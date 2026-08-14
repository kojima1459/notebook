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


Private Const SHEET_NAME As String = "insight_inbox"

' 受信箱の列: A=nonce B=kind C=user_id D=author E=created_at F=question
'             G=answer_or_reason H=source_or_dept I=consumed J=selected
Private Const INBOX_COLS As Long = 10

' ----------------------------------------------------------------------------
' InboxArray - 受信箱を【一括Range読み1回】で配列にする(2026-08-14 R32 W1-5)。
'   件数カード(GapCount/PendingQACount)は Hub を描くたびに走るのに、1セルずつ
'   Cells(r,c).Value でCOM往復していた。500行なら1,000回の往復が描画のたびに
'   起きる(PendingRowsRanked は R12-3-5 で既に一括読みへ直してあり、同型)。
'   戻り値は Range.Value の2次元配列(1始まり・行=データ行・列=A..J)。
'   データが無いときは outN=0 で Empty を返す(呼び出し側は outN だけ見る)。
' ----------------------------------------------------------------------------
Private Function InboxArray(ByRef outN As Long) As Variant
    On Error Resume Next
    outN = 0
    Dim ws As Worksheet: Set ws = GetSheet()
    If ws Is Nothing Then Exit Function
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Function
    ' 複数列のRangeは1行でも必ず2次元配列になる(単一セルのときだけスカラーに
    ' なるVBAの仕様に当たらない)。
    InboxArray = ws.Range(ws.Cells(2, 1), ws.Cells(lastR, INBOX_COLS)).Value
    outN = lastR - 1
    On Error GoTo 0
End Function

' 未取り込みの解決済みQ&A件数(Hubのお知らせに出す)。
Public Function PendingQACount() As Long
    On Error Resume Next
    Dim n As Long
    Dim arr As Variant: arr = InboxArray(n)
    Dim i As Long
    For i = 1 To n
        If CStr(arr(i, 2)) = "qa" Then
            If CStr(arr(i, 9)) <> "1" Then PendingQACount = PendingQACount + 1
        End If
    Next i
    On Error GoTo 0
End Function

' 未解決質問(gap)の件数。
'   2026-08-14(R32 r3): 判定を板の一覧(GapListText)と同じ IsGapRow に揃える。
'   従来は件数だけが consumed を見ており、一覧は見ていなかったので、
'   「(3件)」と題に出ているのに一覧には5件並ぶ状態が起こりえた。
Public Function GapCount() As Long
    On Error Resume Next
    Dim n As Long
    Dim arr As Variant: arr = InboxArray(n)
    Dim i As Long
    For i = 1 To n
        If IsGapRow(CStr(arr(i, 2)), CStr(arr(i, 7)), CStr(arr(i, 9))) Then
            GapCount = GapCount + 1
        End If
    Next i
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' IsGapRow - 「みんなの困りごと」板に出す行かどうか(純関数・2026-08-14 R32)。
'   ・kind が gap であること
'   ・取り込み済み(consumed=1)でないこと(r3: 件数と一覧の判定を揃える)
'   ・reason が correction でないこと(W1-1 B2 の旧データ救済)
'     R32以前の訂正投稿は gap フォルダへ書かれていたため、受信側では
'     kind="gap" として受信箱に入っている。フォルダを分けても【既に届いて
'     いる行】は消えないので、板に出す側でも reason 列を見て弾く。
' ----------------------------------------------------------------------------
Public Function IsGapRow(ByVal kind As String, ByVal reason As String, _
                         ByVal consumed As String) As Boolean
    If kind <> "gap" Then Exit Function
    If consumed = "1" Then Exit Function
    IsGapRow = (LCase$(Trim$(reason)) <> "correction")
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

    ' 2026-08-01(R12-3-5): 行ごとに SameQuestionCount を呼ぶと、その中で
    ' 受信箱を毎回全走査するのでO(n²)になっていた(未取込1,000件で100万回の
    ' セル読み=お知らせ欄を開くだけで数十秒固まる)。集計表を【1回】作って
    ' 引くだけにする。行の読みも kind/question/consumed を一括Rangeで取る。
    Dim counts As Object: Set counts = BuildQuestionCounts()
    Dim arr As Variant: arr = ws.Range(ws.Cells(2, 2), ws.Cells(lastR, 9)).Value

    ReDim outRows(0 To lastR)
    Dim scores() As Long: ReDim scores(0 To lastR)
    Dim n As Long, r As Long
    For r = 2 To lastR
        ' 複数列のRangeは1行でも必ず2次元配列になる(単一セルのときだけ
        ' スカラーになるVBAの仕様に当たらない)。
        Dim kind As String, qTxt As String, cons As String
        kind = CStr(arr(r - 1, 1))       ' 2列目=kind
        qTxt = CStr(arr(r - 1, 5))       ' 6列目=question
        cons = CStr(arr(r - 1, 8))       ' 9列目=consumed
        If kind = "qa" And cons <> "1" Then
            outRows(n) = r
            Dim ky As String: ky = NormKey(qTxt)
            If counts.Exists(ky) Then scores(n) = counts.Item(ky)
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
    Dim key As String: key = NormKey(qText)
    If LenB(key) = 0 Then Exit Function

    Dim counts As Object: Set counts = BuildQuestionCounts()
    If counts Is Nothing Then Exit Function
    If counts.Exists(key) Then SameQuestionCount = counts.Item(key)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' BuildQuestionCounts - 質問キー→件数の集計表を【一括Range読み1回】で作る。
'   2026-08-01(R12-3-5)。キャッシュは持たない。受信箱は起動時の収集で増え、
'   取込・トリムで減る「動く表」なので、古い集計を持ち回るくらいなら
'   毎回作り直す方が安全で、それでも全体はO(n)で収まる(作るのは1回、
'   引くのはDictionaryのO(1))。
' ----------------------------------------------------------------------------
Private Function BuildQuestionCounts() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set BuildQuestionCounts = d

    On Error Resume Next
    Dim n As Long
    Dim arr As Variant: arr = InboxArray(n)
    ' 2026-08-14(R32 r2): kind を見ずに全行を数えていたため、困りごと(gap)や
    ' 訂正(correction)の行まで「同じ質問が何件届いたか」に混ざり、解決済み
    ' Q&Aの並び(PendingRowsRanked)と「同じ質問◯件」の表示が歪んでいた。
    ' 数えるのは qa だけ ―― この集計表を読むのは qa の画面しかない。
    Dim r As Long
    For r = 1 To n
        If CStr(arr(r, 2)) = "qa" Then BumpKey d, NormKey(CStr(arr(r, 6)))
    Next r
    On Error GoTo 0
End Function

Private Sub BumpKey(ByVal d As Object, ByVal key As String)
    If LenB(key) = 0 Then Exit Sub
    If d.Exists(key) Then
        d.Item(key) = CLng(d.Item(key)) + 1
    Else
        d.Add key, 1
    End If
End Sub

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

Private Function GetSheet() As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
End Function

' EnsureSheet: 受信箱シートの作成/取得。2026-07-31(R11-F1)で収集側を
' modInsightIo へ分けたため、そこから呼べるよう Public にした。
Public Function EnsureSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = GetSheet()
    If Not ws Is Nothing Then
        Set EnsureSheet = ws
        Exit Function
    End If
    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = SHEET_NAME
    ' 2026-08-14(R32 m5): ヘッダーがA1:I1の9列で、J列(selected=選択状態)だけ
    ' 名前が無かった。実データは10列で、引き継ぎ(modMigrate)もビルド側の
    ' 検証も10列を前提にしている。名前が無い列は「無い列」と読まれる。
    ws.Range("A1:J1").Value = Array("nonce", "kind", "user_id", "author", "created_at", _
                                    "question", "answer_or_reason", "source_or_dept", _
                                    "consumed", "selected")
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
