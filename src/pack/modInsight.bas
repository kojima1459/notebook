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

' 連投抑止キー(my_stats)の接頭辞。modP2P の "thx:" / 本機能の "ins:" と同作法。
' 2026-08-14(R32 W1-6)。実体の説明は下の「発信側の関所」節。
Private Const GAPQ_PREFIX As String = "gapq:"
' 抑止キーそのものの保持日数(抑止期間を過ぎたキーを my_stats に残す理由は無い)。
Private Const GAPQ_KEEP_DAYS As Long = 7

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
' InboxNonceSet - 受信箱に既に在る nonce の集合(2026-08-14 R32 W1-2 B1)。
'   共有フォルダの実ファイルを消すのは【発行者端末だけ】(R8b B12)なのに、
'   既読印 "ins:" は【全端末】が自分で67日で消す(GcOldNonces)。この非対称の
'   ため、発行担当が居ない・久しく起動していない組織では、67日目に全端末が
'   過去の投稿を丸ごと再受信する。AppendRow に重複ガードが無いと受信箱が
'   二重化し、取り込み済みだったQ&Aが「未取込」に戻り、
'   RegisterKnowledgeText の一時ファイル名に時刻が入る(=別ファイル扱い)
'   ため【本棚に同じQ&Aが二重登録される】。
'   既読印だけでは止まらないので、受信箱そのものを突き合わせの真実にする。
'   照合はDictionaryのO(1)。my_stats を1件ずつ引く形(R8 F10で捨てた形)へは
'   戻さない。
' ----------------------------------------------------------------------------
Public Function InboxNonceSet() As Object
    Dim d As Object
    On Error Resume Next
    Set d = CreateObject("Scripting.Dictionary")
    Set InboxNonceSet = d
    If d Is Nothing Then Exit Function
    Dim n As Long
    Dim arr As Variant: arr = InboxArray(n)
    Dim i As Long
    For i = 1 To n
        Dim k As String: k = LCase$(Trim$(CStr(arr(i, 1))))
        If LenB(k) > 0 Then d(k) = 1
    Next i
    On Error GoTo 0
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

' ============================================================================
' 発信側の関所(2026-08-14 R32 W1-6/W1-8)
' ----------------------------------------------------------------------------
' 置き場所の理由: 判定の実体は modInsightIo(発信の持ち主)に置きたいが、
'   あちらは30,000字上限まで残り1,800字を切っており1件も入らない。憲章§4-6に
'   従い、余裕のある本モジュールへ実体を置き、modInsightIo からは1行で呼ぶ。
'   連投抑止キーの接頭辞・読み書き・GCを1モジュールに集めておくことで、
'   キー名が2箇所に書き写されてズレる事故も同時に防いでいる。
' ============================================================================

' ----------------------------------------------------------------------------
' PiiBlocked - 共有フォルダへ出す前の個人情報走査(R32 W1-8(b) M7)。
'   True を返したら【送らない】。
'   困りごとの投稿は質問の全文がそのまま部内へ出ていく経路で、しかも
'   no_hit は利用者の操作を介さず自動で発火する(押した覚えのないまま
'   「取引先の田中様の携帯 090-xxxx-xxxx は…」が部内に配られる)。
'   共有フォルダに一度書いたものは取り消せないため、ここは検知したら黙って
'   見送る(利用者の作業は止めない。痕跡は usage_log に1行残す)。
'
'   【重要】この走査は、パック発行側の個人情報チェックのオフスイッチ
'   (config pii_scan_enabled・R32 波2)の影響を受けない。あちらは自分で
'   選んで自分の資料を出す操作で、利用者が中身を分かって押している。
'   こちらは自動発火かつ他人の目に触れる経路なので、常に走らせる。
' ----------------------------------------------------------------------------
Public Function PiiBlocked(ByVal s As String, ByVal what As String) As Boolean
    On Error Resume Next
    Dim hit As String
    hit = modPii.ScanText(s)
    If LenB(hit) = 0 Then Exit Function
    PiiBlocked = True
    modLog.LogUsage "insight_pii_skip", what, _
        "個人情報を含む可能性があるため部内共有を見送りました(" & hit & ")"
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' GapDupBlocked - 同じ趣旨の質問を短時間に連投していないか(R32 W1-6 M4)。
'   EmitGap には重複抑止が無く、しかも no_hit 経路は感想ボタンを通らないので
'   「1ターン1回」のガードすら掛からない。言い換えて3回聞けば3件投稿され、
'   板が1人の質問で埋まる(見る側からは同じ質問が3行並ぶ)。
'   照合は解決済みQ&A側の SameQuestionCount と同じ NormKey(記号と空白を
'   落とした先頭40字)。思想を揃えておかないと、片方で「同じ」もう片方で
'   「別」という食い違いが必ず出る。
'   抑止時間は config gap_dup_hours(既定24時間)。0以下で無効。
' ----------------------------------------------------------------------------
Public Function GapDupBlocked(ByVal qText As String) As Boolean
    On Error Resume Next
    Dim hours As Long: hours = modConfig.GetLong("gap_dup_hours", 24)
    If hours <= 0 Then Exit Function
    Dim k As String: k = NormKey(qText)
    If LenB(k) = 0 Then Exit Function
    GapDupBlocked = WithinWindow(modStats.GetStatText(GAPQ_PREFIX & k), _
                                 modUtilText.IsoDateTime(DateAdd("h", -hours, Now)))
    On Error GoTo 0
End Function

' 発信できたことを記録する(次の GapDupBlocked がこれを見る)。
Public Sub MarkGapEmitted(ByVal qText As String)
    On Error Resume Next
    Dim k As String: k = NormKey(qText)
    If LenB(k) = 0 Then Exit Sub
    modStats.SetStatText GAPQ_PREFIX & k, modUtilText.IsoDateTime(Now)
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' GcGapDupKeys - 抑止期間を過ぎた連投抑止キーを my_stats から取り除く。
'   modInsightIo.GcOldNonces と同じ「下から消す」作法。掃除しないと、質問の
'   種類ぶんだけ my_stats が伸び、FindKeyRow の線形探索が重くなる
'   (HANDOFF §2 の既知課題・R32 r5 と同型)。
' ----------------------------------------------------------------------------
Public Sub GcGapDupKeys()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_STATS)
    If ws Is Nothing Then Exit Sub
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Sub

    Dim limit As String
    limit = modUtilText.IsoDateTime(DateAdd("d", -GAPQ_KEEP_DAYS, Now))
    ' A:B の2列で読む(1行しか無いときA列単独だと配列ではなくスカラーが返る
    ' VBAの仕様を避けるため。GcOldNonces と同じ回避策)。
    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastR, 2)).Value

    Dim i As Long
    For i = UBound(arr, 1) To LBound(arr, 1) Step -1
        Dim k As String: k = CStr(arr(i, 1))
        If Left$(k, Len(GAPQ_PREFIX)) = GAPQ_PREFIX Then
            If Not WithinWindow(CStr(arr(i, 2)), limit) Then ws.Rows(i + 1).Delete
        End If
    Next i
    On Error GoTo 0
End Sub

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

' nonce重複ガードの判定(純関数・W1-2)。集合そのものを作れない環境
' (Scripting.Dictionary が使えない等)では False=従来どおり足す へ倒す。
' 「速くするための仕組みが無いと機能そのものが止まる」を作らない(IsSeen と
' 同じ考え方)。二重取り込みは受信箱側の実害だが、取りこぼしは共有知の
' 断絶で、後者のほうが重い。
Public Function NonceIsKnown(ByVal known As Object, ByVal nc As String) As Boolean
    If known Is Nothing Then Exit Function
    Dim k As String: k = LCase$(Trim$(nc))
    If LenB(k) = 0 Then Exit Function
    NonceIsKnown = known.Exists(k)
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
'   いるのか分からないYes/Noダイアログになっていた(M1)。ここで合計900字に
'   収める(1,024との差は、呼び出し側が前後に足す案内文のぶんの余裕)。
'     ・1件の質問文は先頭100字まで(SafeLeft=サロゲートペアを割らない)
'     ・部署も40字で切る(共有フォルダの他人のファイル由来=長さは信用しない)
'     ・入りきらない件数は捨てずに「…ほか N 件」として残す
'     ・並びは created_at の降順(M5: 従来は受信順=Dir()の列挙順=ユーザーID順
'       なのに「新しい順」と名乗っていた)
'   1件目だけは字数に関わらず必ず入れる(全部消えて空の板になるのを避ける。
'   上の切り詰めにより1件あたりは高々200字程度で、900字を割ることはない)。
' ----------------------------------------------------------------------------
Public Function GapListBuild(ByRef g() As String, ByVal cnt As Long) As String
    Const MAX_ITEMS As Long = 20
    Const MAX_CHARS As Long = 900
    Const TAIL_ROOM As Long = 24        ' 「…ほか N 件」の予約枠
    Const Q_CHARS As Long = 100
    Const DEPT_CHARS As Long = 40
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
                "     (" & who & " ・ " & g(2, i) & " ・ " & ReasonText(g(3, i)) & ")" & vbLf
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
