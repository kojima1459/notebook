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

' 受信箱(insight_inbox)のトリム条件(2026-08-01 R12-3-5)。
' 取り込み済み(consumed=1)の行は、本棚に入った時点で役目を終えている。
' 消さないと、共有知が回るほど受信箱が一方的に伸び、起動時の収集も
' 「みんなの困りごと」の描画も毎回そのぶん重くなる(chat_logの100件
' ローテと同じ考え方。ただしこちらは【未取込の行は絶対に消さない】)。
Private Const INBOX_MAX_ROWS As Long = 500
' 保持日数は固定値をやめ、既読印GCと同じ config thanks_gc_days から導く
' (2026-08-14 R32 F1)。「既読印GC日数 < 受信箱保持日数」の不等式そのものは
' modInsightGate.NonceKeepDays / InboxKeepDays が持つ(理由と検算はそちら)。

' ----------------------------------------------------------------------------
' TrimInboxRows - 受信箱の古い行を片付ける(2026-08-14 R32 W1-4で改称・拡張)。
'   2026-08-14(R32 Fix波): modInsightIo から本モジュールへ移設した。
'   modInsightIo が30,000字上限を超えたための移設だが、置き場所としても
'   こちらが正しい ―― 本モジュールの持ち分は「受信箱シートの参照・選択・
'   取り込み済み管理」(modInsightIo 冒頭の分割方針)で、受信箱の行を
'   落とす掃除はまさにそれに当たる(呼び出しは CollectInsights の1箇所)。
'   ・取り込み済み(consumed=1): 「保持日数より古い」または「500行を超えた超過分」。
'     未取込のQ&Aは何行あっても消さない。届いた知恵を、読む前にこちらの
'     都合で捨てないため。
'   ・困りごと(gap)と訂正(correction): consumed が立つ経路が存在しないため、
'     上の条件では永久に残る。保持期間(config gap_keep_days・既定30日)で
'     落とす(判定は下の GapAged。理由はそちらのコメント)。
'   行の詰め直しは modShelfStore と同じ「一括読み→配列でフィルタ→一括書戻し」
'   (1行ずつ Rows().Delete すると数千行で実機が固まる。MASTER_SPEC §12)。
' ----------------------------------------------------------------------------
Public Sub TrimInboxRows(ByVal ws As Worksheet)
    On Error Resume Next
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 3 Then Exit Sub                      ' データ1行以下なら触らない

    Dim nRows As Long: nRows = lastR - 1
    Dim excessN As Long: excessN = nRows - INBOX_MAX_ROWS
    If excessN < 0 Then excessN = 0

    ' 期限の境界は文字列比較で判定する(ISO日付は辞書順=時系列順。
    ' 和暦カレンダー端末でも壊れない。R12-1-4と同じ理由でCDateを通さない)。
    Dim keepDays As Long
    keepDays = modInsightGate.InboxKeepDays(modConfig.GetLong("thanks_gc_days", 60))
    Dim cutoff As String: cutoff = modUtilText.IsoDate(Date - keepDays)

    ' R32 W1-4: 困りごと/訂正の保持期間。0以下で無効(=従来どおり残す)。
    Dim gapCut As String
    Dim gapDays As Long: gapDays = modConfig.GetLong("gap_keep_days", 30)
    If gapDays > 0 Then gapCut = modUtilText.IsoDate(Date - gapDays)

    Dim arr As Variant: arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastR, INBOX_COLS)).Value
    Dim keep() As Variant: ReDim keep(1 To nRows, 1 To INBOX_COLS)
    Dim keepN As Long, dropped As Long
    ' 2026-08-01(R12-H-8): 上限超過ぶんと期限切れぶんのカウンタを分ける。
    ' 従来は1つの dropped を「超過何件目か」の判定にも使っていたため、
    ' 期限切れの行を1件消すたびに超過枠が1つ埋まったことになり、
    ' 【上限500行を超えたぶんを掃除しきれない】(次回も超過が残る)状態だった。
    Dim overN As Long, agedN As Long, gapN As Long
    Dim i As Long, c As Long
    For i = 1 To nRows
        Dim drop As Boolean: drop = False
        If CStr(arr(i, 9)) = "1" Then                      ' consumed
            If overN < excessN Then
                drop = True                                ' 上限超過分(古い順)
                overN = overN + 1
            ElseIf LenB(Trim$(CStr(arr(i, 5)))) > 0 Then   ' created_at
                drop = (Left$(modUtilText.NormalizeIsoDate(CStr(arr(i, 5))), 10) < cutoff)
                If drop Then agedN = agedN + 1
            End If
        Else
            ' R32 W1-4: 未取込でも gap/correction は保持期間で落とす。
            drop = GapAged(CStr(arr(i, 2)), CStr(arr(i, 5)), gapCut)
            If drop Then gapN = gapN + 1
        End If
        If drop Then
            dropped = dropped + 1
        Else
            keepN = keepN + 1
            For c = 1 To INBOX_COLS
                keep(keepN, c) = arr(i, c)
            Next c
        End If
    Next i
    If dropped = 0 Then Exit Sub

    If keepN > 0 Then
        Dim outArr() As Variant: ReDim outArr(1 To keepN, 1 To INBOX_COLS)
        Dim k As Long
        For k = 1 To keepN
            For c = 1 To INBOX_COLS
                outArr(k, c) = keep(k, c)
            Next c
        Next k
        ws.Range(ws.Cells(2, 1), ws.Cells(1 + keepN, INBOX_COLS)).Value = outArr
    End If
    ws.Range(ws.Cells(2 + keepN, 1), ws.Cells(1 + nRows, INBOX_COLS)).ClearContents

    modLog.LogUsage "insight_inbox_trim", "", _
        "受信箱の古い行を" & dropped & "件片付けました(上限超過" & overN & _
        "件/取込済みの期限切れ" & agedN & "件/困りごと・訂正の期限切れ" & gapN & _
        "件。残り" & keepN & "行。上限" & INBOX_MAX_ROWS & "行/" & _
        keepDays & "日/困りごと" & gapDays & "日)"
    On Error GoTo 0
End Sub

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
' @unused:**受信箱の1件を取り出す口だが、呼び出す画面がまだ無い**。
'   「箱の役に徹する」設計のまま、UI 側の導線が作られていない(R49 監査)。
'   R50 で導線を足すか、この口ごと畳むかを裁定する。
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

' 質問の照合キー(記号と空白を落とした先頭40字)。2026-08-14(R32 W1-6)で
' Public化 ―― 同じ質問の連投抑止(modInsightIo.EmitGap)が発信側でも同じ
' 正規化を使う必要があるため。2箇所に書き写せば必ずズレる。
Public Function NormKey(ByVal s As String) As String
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
' InboxNonceMemo - 受信箱に既に在る nonce の集合(2026-08-14 R32 W1-2 B1)。
'   共有フォルダの実ファイルを消すのは【発行者端末だけ】(R8b B12)なのに、
'   既読印 "ins:" は【全端末】が自分で67日で消す(GcOldNonces)。この非対称の
'   ため、発行担当が居ない・久しく起動していない組織では、67日目に全端末が
'   過去の投稿を丸ごと再受信する。AppendRow に重複ガードが無いと受信箱が
'   二重化し、取り込み済みだったQ&Aが「未取込」に戻り、
'   RegisterKnowledgeText の一時ファイル名に時刻が入る(=別ファイル扱い)
'   ため【本棚に同じQ&Aが二重登録される】。
'   既読印だけでは止まらないので、受信箱そのものを突き合わせの真実にする。
'
'   2026-08-14(R32 F11): 集合の器を Scripting.Dictionary から
'   【1本の文字列 "|nonce|nonce|…"】へ変えた。理由は2つ:
'     (1) Dictionary版は LibreOffice の純ロジックテストから作れないため、
'         「集合に在るとき True を返す」という B1 の【核心】を一度も検証
'         できていなかった(known=Nothing の陰性経路しか撃てず、実装を壊しても
'         全部PASSのまま出荷できる状態だった)。
'     (2) Dictionary が使えない端末では集合が Nothing になり、二重登録ガードが
'         【丸ごと効かない】まま静かに素通りしていた。文字列なら必ず作れる。
'   照合コストは InStr の線形走査だが、受信箱は最大500行=メモは約2万字で、
'   1回の起動で引くのは高々 MAX_COLLECT(60)回。my_stats を1件ずつ引く形
'   (R8 F10で捨てた形。共有ファイル数×my_stats行数)へ戻すわけではない。
'   メモの形は modBackdrop.MemoKey と同じ「両端を "|" で挟む」= 前方一致の
'   取り違え("abc" が "abcd" に当たる)が起きない。
' ----------------------------------------------------------------------------
Public Function InboxNonceMemo() As String
    On Error Resume Next
    Dim n As Long
    Dim arr As Variant: arr = InboxArray(n)
    Dim memo As String, i As Long
    For i = 1 To n
        memo = NonceMemoAdd(memo, CStr(arr(i, 1)))
    Next i
    InboxNonceMemo = memo
    On Error GoTo 0
End Function

' メモへ1件足す(純関数・R32 F11)。空メモには先頭の "|" から作る。
' 既に在るものは足さない(同じ nonce の行が2つある受信箱でもメモは太らない)。
Public Function NonceMemoAdd(ByVal memo As String, ByVal nc As String) As String
    NonceMemoAdd = memo
    Dim k As String: k = NonceKey(nc)
    If LenB(k) = 0 Then Exit Function
    If NonceIsKnownIn(memo, nc) Then Exit Function
    If LenB(memo) = 0 Then NonceMemoAdd = "|"
    NonceMemoAdd = NonceMemoAdd & k & "|"
End Function

' ----------------------------------------------------------------------------
' GapAged - 保持期間を過ぎた困りごと/訂正の行か(純関数・2026-08-14 R32 W1-4)。
'   gap 行に consumed=1 が立つ経路は【1つも無い】。TrimInboxRows は
'   consumed=1 しか落とさず、INBOX_MAX_ROWS=500 の超過掃除も同様なので、
'   「みんなの困りごと」の件数は永久に減らない = Hubのバッジが単調増加する
'   (R20H H-13 と同型の壊れ方)。手で「解決済みにする」操作UIは容量が
'   許さないため、保持期間(config gap_keep_days・既定30日)で落とす。
'   訂正(correction)も同じ扱いにする ―― W1-1でフォルダを分けた結果、訂正行は
'   kind="correction" になり、gap の保持期間から外れて【二度と減らない行】に
'   なるため(consumed が立つ経路はこちらにも無い)。
'   日付はISO文字列の辞書順で比較する(CDateを通さない。R12-1-4と同じ理由)。
'   created_at が空の行は落とさない(判定材料が無いものを消さない)。
' ----------------------------------------------------------------------------
Public Function GapAged(ByVal kind As String, ByVal createdAt As String, _
                        ByVal cutoff As String) As Boolean
    If kind <> "gap" And kind <> "correction" Then Exit Function
    If LenB(cutoff) = 0 Then Exit Function
    Dim t As String: t = Trim$(createdAt)
    If LenB(t) = 0 Then Exit Function
    GapAged = (Left$(modUtilText.NormalizeIsoDate(t), 10) < cutoff)
End Function

' ----------------------------------------------------------------------------
' WithinWindow - 記録した時刻が期間の内側か(純関数・R32 W1-6)。
'   ISO文字列("yyyy-mm-dd hh:nn:ss")は辞書順=時系列順なので、CDate を通さず
'   文字列比較だけで判定する(和暦カレンダー端末でも壊れない。R12-1-4と同じ
'   理由。GcOldNonces が CDate を使っているのはそれ以前の作法)。
'   読めない値(旧形式の "1"・空)は False=「期間の外」へ倒す。抑止側では
'   「送ってよい」に、GC側では「掃除する」に倒れる ―― どちらも安全側。
' ----------------------------------------------------------------------------
Public Function WithinWindow(ByVal stampText As String, ByVal limitText As String) As Boolean
    Dim t As String: t = Trim$(stampText)
    If Len(t) < 10 Then Exit Function
    If LenB(limitText) = 0 Then Exit Function
    WithinWindow = (t > limitText)
End Function

' ----------------------------------------------------------------------------
' NonceIsKnownIn - 重複ガードの判定を【文字列の集合】で行う版(純関数・R32 F11)。
' ----------------------------------------------------------------------------
'   なぜ切り出したのか: 元の判定は Scripting.Dictionary を受け取る形で、
'   Dictionary が使えない LibreOffice の純ロジックテストからは
'   「known が Nothing のとき False」しか撃てなかった。W1-2(B1)の核心は
'   【集合に在るときに True を返して二重登録を止める】ほうであり、そこが
'   一度もテストされていなかった(壊しても全部PASSのまま出荷できた)。
'   そこで判定そのものを「メモ文字列に在るか」の形へ切り出し、陽性・陰性の
'   両方を固定できるようにする。メモの形は modBackdrop.MemoKey と同じ
'   "|キー|キー|" ―― 両端を "|" で挟むので、前方一致("abc" が "abcd" に
'   当たる)の取り違えが起きない。
'   nonce の正規化は NonceKey 1本に通す(作る側・引く側で表記がズレると
'   ガードが素通りする)。
' ----------------------------------------------------------------------------
Public Function NonceIsKnownIn(ByVal memo As String, ByVal nc As String) As Boolean
    Dim k As String: k = NonceKey(nc)
    If LenB(k) = 0 Then Exit Function
    If LenB(memo) = 0 Then Exit Function
    NonceIsKnownIn = (InStr(1, memo, "|" & k & "|", vbBinaryCompare) > 0)
End Function

' 集合に入れる/引くときの nonce の正規化(純関数)。Windowsのファイル名は
' 大文字小文字を区別しないので、集合を作った側と引く側で表記が違うだけで
' 重複ガードが素通りする。作る側(InboxNonceMemo/NonceMemoAdd)・引く側(NonceIsKnownIn)・
' 足した直後に自分で登録する側(modInsightIo.AppendRow)の3箇所が必ず
' この1本を通る。
Public Function NonceKey(ByVal nc As String) As String
    NonceKey = LCase$(Trim$(nc))
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
    Dim n As Long
    Dim arr As Variant: arr = InboxArray(n)
    If n = 0 Then
        GapListText = GapEmptyText()
        Exit Function
    End If

    ' 板に出す行だけを4列(0=質問 1=部署 2=created_at 3=reason)へ写してから
    ' 純関数へ渡す。氏名(D列)は読まない ―― W1-8で発信側が氏名を送らなく
    ' なったのに合わせ、旧データが持っている氏名もここで表示しない。
    Dim g() As String: ReDim g(0 To 3, 0 To n - 1)
    Dim cnt As Long, i As Long
    For i = 1 To n
        If IsGapRow(CStr(arr(i, 2)), CStr(arr(i, 7)), CStr(arr(i, 9))) Then
            g(0, cnt) = CStr(arr(i, 6))
            g(1, cnt) = CStr(arr(i, 8))
            g(2, cnt) = CStr(arr(i, 5))
            g(3, cnt) = CStr(arr(i, 7))
            cnt = cnt + 1
        End If
    Next i

    If cnt = 0 Then
        GapListText = GapEmptyText()
    Else
        GapListText = GapListBuild(g, cnt)
    End If
    On Error GoTo 0
End Function

Private Function GapEmptyText() As String
    GapEmptyText = "まだ届いていません。" & vbLf & _
        "誰かが質問して本棚に答えが無かったとき、その質問がここに自動で並びます。"
End Function

' ----------------------------------------------------------------------------
' GapListBuild - 板の本文を組み立てる(純関数・2026-08-14 R32 W1-3/W1-7)。
'   g(0..3, 0..cnt-1): 0=質問 1=部署 2=created_at 3=reason
'
'   MsgBox の本文は実機で約1,024字を超えると【無言で切り落とされる】。従来は
'   最大20件×1件あたり最長2,000字を素で流し込んでいたため、超過分だけでなく
'   本文の後ろに置いた「今すぐ登録しますか?」の一文まで消え、何を聞かれて
'   いるのか分からないYes/Noダイアログになっていた(M1)。
'
'   【上限の逆算】(2026-08-14 R32 F3。W1-3の見込み124字は実測と38字ずれていた)
'   呼び出し側 modShared.ShowGapBoard が本文の前後に足す案内文は【実測162字】:
'       前 36 + 2 + 37 + 2 + 2 = 79字
'       後  2 + 22 + 2 + 35 + 2 + 20 = 83字   (2 は vbCrLf)
'   1,024 − 162 = 862 が本文に使える上限。ここを 900 にしていたので、
'   最悪ケースで 1,062字 = 上限超過(=末尾の「今すぐ登録しますか?」が
'   また消える)だった。余裕を42字とって【820】へ下げる。
'   案内文を書き換えるときは、この逆算をやり直すこと。
'     ・1件の質問文は先頭100字まで(SafeLeft=サロゲートペアを割らない)
'     ・部署も40字で切る(共有フォルダの他人のファイル由来=長さは信用しない)
'     ・created_at も20字で切る(F10。他人のファイル由来なので長さを信用しない。
'       ここだけクランプが無く、壊れた1行で本文全体を押し流せる状態だった)
'     ・入りきらない件数は捨てずに「…ほか N 件」として残す
'     ・並びは created_at の降順(M5: 従来は受信順=Dir()の列挙順=ユーザーID順
'       なのに「新しい順」と名乗っていた)
'   1件目だけは字数に関わらず必ず入れる(全部消えて空の板になるのを避ける。
'   上の切り詰めにより1件は高々 3+100+1+6+40+3+20+3+25+1+1 = 203字で、
'   820字を割ることはない)。
' ----------------------------------------------------------------------------
Public Function GapListBuild(ByRef g() As String, ByVal cnt As Long) As String
    Const MAX_ITEMS As Long = 20
    Const MAX_CHARS As Long = 820       ' 1024 − 案内文162 − 余裕42
    Const TAIL_ROOM As Long = 24        ' 「…ほか N 件」の予約枠
    Const Q_CHARS As Long = 100
    Const DEPT_CHARS As Long = 40
    Const DATE_CHARS As Long = 20       ' F10: created_at のクランプ
    If cnt <= 0 Then Exit Function

    Dim lim As Long: lim = cnt
    If lim > MAX_ITEMS Then lim = MAX_ITEMS
    SortGapDesc g, cnt, lim

    Dim sb As String, shown As Long, i As Long
    For i = 0 To lim - 1
        Dim who As String: who = modUtil.SafeLeft(Trim$(g(1, i)), DEPT_CHARS)
        If LenB(who) = 0 Then who = "部署の記録なし"
        Dim entry As String
        entry = (shown + 1) & ". " & modUtil.SafeLeft(g(0, i), Q_CHARS) & vbLf & _
                "     (" & who & " ・ " & modUtil.SafeLeft(g(2, i), DATE_CHARS) & _
                " ・ " & ReasonText(g(3, i)) & ")" & vbLf
        If shown > 0 Then
            If Len(sb) + Len(entry) > MAX_CHARS - TAIL_ROOM Then Exit For
        End If
        sb = sb & entry
        shown = shown + 1
    Next i

    If cnt > shown Then sb = sb & ChrW(&H2026) & "ほか " & (cnt - shown) & " 件" & vbLf
    GapListBuild = sb
End Function

' ----------------------------------------------------------------------------
' SortGapDesc - created_at の降順(新しい順)に並べ替える(純関数・W1-7)。
'   上位 topK 件だけを確定させる部分選択ソート。受信箱は最大500行あり、
'   表示は最大20件なので、全体を泡立てると12万回の比較を毎回捨てることになる。
'   ISO文字列は辞書順=時系列順なので CDate を通さない(和暦カレンダー端末でも
'   壊れない。R12-1-4 と同じ理由)。
' ----------------------------------------------------------------------------
Public Sub SortGapDesc(ByRef g() As String, ByVal cnt As Long, ByVal topK As Long)
    If cnt < 2 Then Exit Sub
    Dim k As Long: k = topK
    If k > cnt - 1 Then k = cnt - 1
    Dim i As Long, j As Long, best As Long, c As Long
    For i = 0 To k - 1
        best = i
        For j = i + 1 To cnt - 1
            If g(2, j) > g(2, best) Then best = j
        Next j
        If best <> i Then
            For c = 0 To 3
                Dim t As String
                t = g(c, i): g(c, i) = g(c, best): g(c, best) = t
            Next c
        End If
    Next i
End Sub

' 理由コードの日本語訳(2026-08-14 R32 m4)。未知コードをそのまま出すと、
' 板に生の英語("correction" 等)が並ぶ ―― 実際に訂正投稿の混入(B2)で
' 起きていた。既知の4種以外も必ず日本語の器に入れて出す。
Public Function ReasonText(ByVal code As String) As String
    Select Case LCase$(Trim$(code))
        Case "no_hit":     ReasonText = "本棚に該当資料なし"
        Case "low_conf":   ReasonText = "根拠が薄い"
        Case "wrong":      ReasonText = "回答が違うと報告"
        Case "correction": ReasonText = "回答への訂正"
        Case "":           ReasonText = "理由の記録なし"
        Case Else:         ReasonText = "その他(" & modUtil.SafeLeft(Trim$(code), 20) & ")"
    End Select
End Function
