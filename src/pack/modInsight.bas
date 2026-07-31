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

' 既読印(my_stats)の接頭辞。modP2P の "thx:" と同じ作法(R8 F10)。
Private Const INS_PREFIX As String = "ins:"

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

    ' 2026-07-31(レビュー R8 F10): 既読印(my_stats の "ins:" 行)を
    ' 【1回の一括読み】で集合(Dictionary)にしてから照合する。
    ' 従来は1ファイルにつき modStats.GetStat("ins:"&nc) を呼んでいた。
    ' GetStat は my_stats を上から探す実装なので、共有フォルダのファイル数
    ' × my_stats の行数の計算量になり、しかも my_stats はこの機能自体が
    ' 際限なく行を足す(下の GcOldNonces まで掃除する仕組みが無かった)。
    ' 使うほど起動が重くなる形だった。
    Dim seen As Object: Set seen = LoadSeenSet()

    Dim qaDir As String: qaDir = SubDir(QA_SUBDIR)
    Dim gapDir As String: gapDir = SubDir(GAP_SUBDIR)
    CollectInsights = CollectFrom(ws, qaDir, "qa", seen) + _
                      CollectFrom(ws, gapDir, "gap", seen)

    ' GC(R8 F10)。my_stats 側の既読印は【全端末】が自分のブックを掃除する。
    ' 自分のシートが太るのは自分の問題なので、誰がやっても構わない。
    GcOldNonces

    ' 2026-07-31(R8b B12): 共有フォルダ側(qa/gap の実ファイル)を消すのは
    ' 【発行者端末だけ】に限定する。
    ' 全端末が消しに行くと、次の2つが同時に起きる:
    '   ・朝の一斉起動で数十台が同じフォルダに対して Kill を撃ち合う
    '     (消えた直後のファイルへの Kill でエラー、列挙と削除の競合)
    '   ・「まだ誰も受け取っていない投稿」を、たまたま最初に起動した1台の
    '     時計や設定(thanks_gc_days)の都合で全員から奪える
    ' 共有フォルダの寿命管理は、本来1箇所が責任を持つべき運用作業に近い。
    ' 判定は既存の発行者判定(modPublish.CanPublish = config publish_key が
    ' 入っている = --publisher 付きでビルドした発行者用ブック)を再利用する。
    ' 新しい設定キーを増やさない(増やすほど運用が壊れやすくなる)。
    Dim isPublisher As Boolean
    isPublisher = False
    On Error Resume Next
    isPublisher = modPublish.CanPublish()
    On Error GoTo 0
    If Not isPublisher Then Exit Function

    Dim killedN As Long
    killedN = GcOldInsights(qaDir) + GcOldInsights(gapDir)
    If killedN > 0 Then
        On Error Resume Next
        modLog.LogUsage "insight_gc", "", _
            "共有フォルダの古い共有知を " & killedN & "件片付けました" & _
            "(発行者端末のみが実行。保持日数は config thanks_gc_days)。"
        On Error GoTo 0
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' LoadSeenSet - my_stats の "ins:" 行を1回で読んで既読集合にする(R8 F10)。
'   キーは nonce(ファイル名から拡張子を除いたもの)。値は使わない。
' ----------------------------------------------------------------------------
Private Function LoadSeenSet() As Object
    Dim d As Object
    On Error Resume Next
    Set d = CreateObject("Scripting.Dictionary")
    Set LoadSeenSet = d
    If d Is Nothing Then Exit Function

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_STATS)
    If ws Is Nothing Then Exit Function
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Function

    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastR, 1)).Value

    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim k As String: k = CStr(arr(i, 1))
        If Left$(k, Len(INS_PREFIX)) = INS_PREFIX Then
            d(LCase$(Mid$(k, Len(INS_PREFIX) + 1))) = 1
        End If
    Next i
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' GcOldInsights - N日以上前の qa/gap ファイルを共有フォルダから消す(R8 F10)。
'   戻り値 = 実際に消した件数(R8b B12: 呼び出し側が usage_log に残す)。
'   modP2P.GcOldThanks と同じ作法(集めてから消す。列挙中に Kill しない)。
'   保持日数は感謝状と揃えて config thanks_gc_days(既定60)。0以下でGC無効。
'   ここに掃除が無かったため、共有の insight\qa\ は永久に増え続け、
'   起動のたびの Dir 列挙が全員ぶん重くなっていた。
'   【呼ぶのは発行者端末だけ】(R8b B12。理由は CollectInsights 側のコメント)。
' ----------------------------------------------------------------------------
Private Function GcOldInsights(ByVal dirPath As String) As Long
    If LenB(dirPath) = 0 Then Exit Function
    On Error Resume Next
    Dim days As Long: days = modConfig.GetLong("thanks_gc_days", 60)
    If days <= 0 Then Exit Function
    If Len(Dir(dirPath, vbDirectory)) = 0 Then Exit Function

    Dim old() As String: ReDim old(0 To 63)
    Dim n As Long
    Dim limit As Date: limit = DateAdd("d", -days, Now)

    Dim fn As String: fn = Dir(dirPath & "*.txt")
    Do While LenB(fn) > 0
        If n > UBound(old) Then ReDim Preserve old(0 To UBound(old) + 64)
        old(n) = fn
        n = n + 1
        fn = Dir()
    Loop

    Dim i As Long
    For i = 0 To n - 1
        Dim full As String: full = dirPath & old(i)
        Dim stamp As Date
        Err.Clear
        stamp = FileDateTime(full)
        If Err.Number = 0 Then
            If stamp < limit Then
                If modP2PIo.KillRetry(full) Then GcOldInsights = GcOldInsights + 1
            End If
        End If
        Err.Clear
    Next i
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' GcOldNonces - 期限を過ぎた "ins:" 行を my_stats から取り除く(R8 F10)。
'   modP2P.GcOldNonces と同じ作法。共有側のファイルを N日で消す以上、
'   既読印だけ永久に残すと my_stats が無限に伸びる(=起動が重くなる)。
'   保持は共有側+7日。
'
'   旧い端末の既読印は Bump で書かれた数値("1")で、日付として読めない。
'   modP2P は読めない値を掃除対象にしているが、ここで同じことをすると
'   【まだ共有に残っている投稿を二度取り込む】ことになり、受信箱に同じ
'   Q&Aが並ぶ。読めない値は今日の日付へ書き直すだけにして、次の周回から
'   正しく期限管理する(移行は1回で終わる)。
' ----------------------------------------------------------------------------
Private Sub GcOldNonces()
    On Error Resume Next
    Dim keepDays As Long
    keepDays = modConfig.GetLong("thanks_gc_days", 60) + 7
    If keepDays < 1 Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_STATS)
    If ws Is Nothing Then Exit Sub
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Sub

    Dim limit As Date: limit = DateAdd("d", -keepDays, Date)
    Dim today As String: today = Format$(Date, "yyyy-mm-dd")
    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastR, 2)).Value

    ' 下から消す(上から消すと行番号がずれる)。
    Dim i As Long
    For i = UBound(arr, 1) To LBound(arr, 1) Step -1
        Dim k As String: k = CStr(arr(i, 1))
        If Left$(k, Len(INS_PREFIX)) = INS_PREFIX Then
            Dim v As String: v = Trim$(CStr(arr(i, 2)))
            ' 旧形式は Bump が書いた数値("1")。CDate("1") は 1899-12-31 という
            ' 「有効な日付」になってしまうので、数値かどうかを先に見る
            ' (ここを間違えると、旧端末の既読印が全部期限切れ扱いで消える)。
            Dim isOldForm As Boolean
            isOldForm = (LenB(v) = 0)
            If Not isOldForm Then isOldForm = IsNumeric(v)
            If isOldForm Then
                ws.Cells(i + 1, 2).Value = today     ' 旧形式("1")の移行
            Else
                Dim d As Date
                Err.Clear
                d = CDate(v)
                If Err.Number = 0 Then
                    If d < limit Then ws.Rows(i + 1).Delete
                End If
                Err.Clear
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

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
                             ByVal kind As String, ByVal seen As Object) As Long
    If LenB(dirPath) = 0 Then Exit Function
    On Error Resume Next
    If Len(Dir(dirPath, vbDirectory)) = 0 Then Exit Function

    Dim myId As String: myId = SafeUserId()

    ' 1) まずファイル名を全部集める。Dir()は列挙状態を1つしか持たないので、
    '    ループの中で読み書きをすると列挙が壊れる(他モジュールと同じ2段方式)。
    Dim names() As String: ReDim names(0 To 255)
    Dim nFiles As Long
    Dim fname As String
    fname = Dir(dirPath & "*.txt")
    Do While LenB(fname) > 0
        If nFiles > UBound(names) Then ReDim Preserve names(0 To UBound(names) + 256)
        names(nFiles) = fname
        nFiles = nFiles + 1
        fname = Dir()
    Loop

    ' 2) 未読だけを処理する。
    '
    ' 2026-07-28(レビュー H-15)で直した3点:
    '   (a) 読み取りに失敗しても既読にしていた
    '       → その投稿はその端末に二度と入らない。既読化は成功パスへ移す。
    '   (b) 上限60を「列挙した件数」で数えていた
    '       → 既読・自分発のファイルも枠を食うため、部内の累計投稿が60件を
    '         超えた日から新着が誰にも届かなくなる(エラー表示も無し)。
    '         共有知フライホイールが静かに止まる。数えるのは「実際に
    '         処理した未読の件数」にする。
    '   (c) 自分発の判定が前方一致だった(レビュー M-22)
    '       → ID "鈴木" が "鈴木一" さんの投稿まで自分発と誤認して受信しない。
    '         ファイル名は "<ID>-<nonce>" なので、区切りまで含めて比較する。
    Dim processed As Long
    Dim i As Long
    For i = 0 To nFiles - 1
        If processed >= MAX_COLLECT Then Exit For
        Dim nc As String: nc = Left$(names(i), Len(names(i)) - 4)
        If Not IsMine(nc, myId) Then
            ' 既読判定は一括読みした集合で行う(R8 F10)。my_stats を
            ' ファイル1件ごとに走査しない。
            If Not IsSeen(seen, nc) Then
                processed = processed + 1
                Dim raw As String
                If ReadShared(dirPath & names(i), raw) Then
                    If AppendRow(ws, kind, nc, raw) Then CollectFrom = CollectFrom + 1
                    ' 読めたときだけ既読にする。値は日付にしておく
                    ' (GcOldNonces が期限を判定できるようにするため。R8 F10)。
                    modStats.SetStatText INS_PREFIX & nc, Format$(Date, "yyyy-mm-dd")
                    If Not seen Is Nothing Then seen(LCase$(nc)) = 1
                End If
            End If
        End If
    Next i
    On Error GoTo 0
End Function

' 既読集合の照合(R8 F10)。集合が作れなかった環境(Scripting.Dictionary が
' 使えない等)では、従来どおり my_stats を1件ずつ引く経路へ落とす。
' 「速くするための仕組みが無いと動かない」は、機能そのものを壊す作りなので避ける。
Private Function IsSeen(ByVal seen As Object, ByVal nc As String) As Boolean
    If seen Is Nothing Then
        IsSeen = (LenB(modStats.GetStatText(INS_PREFIX & nc)) > 0)
        Exit Function
    End If
    IsSeen = seen.Exists(LCase$(nc))
End Function

' ファイル名の先頭が自分のIDか。前方一致だと "鈴木" が "鈴木一" を巻き込むため、
' 「ID そのもの」か「ID + 区切り」でだけ一致とみなす(レビュー M-22)。
Private Function IsMine(ByVal nameCore As String, ByVal myId As String) As Boolean
    If LenB(myId) = 0 Then Exit Function
    If StrComp(nameCore, myId, vbTextCompare) = 0 Then
        IsMine = True
        Exit Function
    End If
    IsMine = (StrComp(Left$(nameCore, Len(myId) + 1), myId & "-", vbTextCompare) = 0)
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
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup21
FailCleanup21:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Set st = Nothing
    On Error GoTo 0
End Function

' 共有フォルダからの読み取り。AVスキャン中・書込み直後のロックで一時的に
' 失敗することがあるため、modP2P と同じく3回まで待って試す(レビュー H-15)。
' 1回で諦めると、その投稿は既読化されないまま毎回失敗し続けるか、
' 修正前のように既読化されて永久に落ちる。
Private Function ReadShared(ByVal filePath As String, ByRef outText As String) As Boolean
    Dim attempt As Long
    For attempt = 1 To 3
        If ReadSharedOnce(filePath, outText) Then
            ReadShared = True
            Exit Function
        End If
        WaitMs 150 * attempt
    Next attempt
End Function

' 指定ミリ秒だけ待つ(Sleep宣言を増やさないためTimerで回す。共有I/Oの
' リトライ間隔なので精度は要らない)。
Private Sub WaitMs(ByVal ms As Long)
    On Error Resume Next
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000 < ms
        DoEvents
        If Timer < t0 Then Exit Do      ' 日付をまたいだ(Timerが0へ戻る)
    Loop
    On Error GoTo 0
End Sub

Private Function ReadSharedOnce(ByVal filePath As String, ByRef outText As String) As Boolean
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
    ReadSharedOnce = True
    Exit Function
Fail:
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup24
FailCleanup24:
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
' 到達性の判定は modShare が1セッション1回だけ行う(レビュー M-20)。
Private Function SubDir(ByVal leaf As String) As String
    Dim root As String: root = modShare.SubDir(INSIGHT_SUBDIR)
    If LenB(root) = 0 Then Exit Function
    EnsureDir root
    SubDir = root & leaf & "\"
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
