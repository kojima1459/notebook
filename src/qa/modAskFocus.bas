Attribute VB_Name = "modAskFocus"
Option Explicit

' ============================================================================
' modAskFocus - 精読(根拠チャンクの前後を一緒に読む) 2026-08-05 R16-3C
' ----------------------------------------------------------------------------
' なぜ要るのか(R16 要望③・仕様 spec_20260805_R16 §3):
'   検索はチャンク単位で当たる。ところが実務の資料は「表の見出しが1つ前の
'   チャンク、値が次のチャンク」「条文の本体と但し書きがページ跨ぎ」のように、
'   意味の単位がチャンクの境界と一致しない。当たったチャンクだけを渡すと、
'   AIは【文の途中から読み始めて途中で読み終わる】ことになり、そこに書いて
'   あるのに「資料からは確認できませんでした」と答える。利用者から見ると
'   「入っている資料の話なのに読めていない」という一番信用を失う外し方になる。
'
'   そこで、根拠に選んだチャンクの前後 radius 個ぶんを文書順で拾って足す。
'   検索の当て方は1文字も変えない(順位も件数もそのまま)。当たった場所の
'   【周辺を読む】だけを足す=精読。
'
' 設計判断:
'   ・足した近傍は元ヒットの【後ろへ付ける】。前後へ差し込むと、上位ヒットが
'     max_context_chars(既定40,000字・LLM呼び出しごと)の打ち切りで押し出され、
'     一番関連の高い資料が黙って落ちる。後ろに付ければ、溢れたときに落ちるのは
'     必ず「おまけの近傍」の側になる(既存の打ち切りがそのまま安全弁になる)。
'   ・近傍の score は 0 にする。親ヒットのスコアを継がせると
'     modAsk.LastConfidence の「閾値以上のヒットが2件以上=強く一致」が
'     1件の強ヒット+その近傍4個で成立してしまい、信頼度バッジが嘘をつく。
'     出典タグに使う source / page は【近傍チャンク自身の値】を使う
'     (親のページを継ぐと、突合は通るのに原文を開くと別ページになる)。
'   ・文書順は chunk_id "bs::ハッシュ::pN::cM" の (page, seq) 昇順で復元する。
'     seq はページごとに振り直されるので、page を先に見ないと順序が壊れる。
'     文書順を返す公開APIは本棚側に無く、列の追加は既存本棚の移行が要るため、
'     ここで chunk_id を自前パースする(仕様 §3 前提事実)。
'   ・my_knowledge の読み方は modRetrieve と同じ直接 Range 読み。ただし読むのは
'     【chunk_id 列だけ】で、本文(full_text=最大32,000字)は採用した数十行だけ
'     セル単位で読む。全列を読み直すと、1回の質問で本棚全体の本文を2度
'     メモリへ載せることになる(20,000件規模で数百MB)。
'   ・シートが読めない・形が違うときは何もせずに戻る(無操作=従来動作)。
'     精読は上積みであって、失敗して回答自体を止めてよい種類の処理ではない。
' ============================================================================

' my_knowledge の列(modShelfStore の定義と同じ並び。ここは読むだけ)。
Private Const COL_ID As Long = 1
Private Const COL_SOURCE As Long = 2
Private Const COL_ORIGIN As Long = 3
Private Const COL_PAGE As Long = 4
Private Const COL_FULLTEXT As Long = 7

Private Const PREVIEW_LEN As Long = 120

' 1回の質問で足す近傍チャンクの上限。max_context_chars の打ち切りが本来の
' 安全弁だが、配列の再確保とセル読みの回数はその手前で効かせておく
' (topK=16・radius=2 なら理屈上の最大は64個)。
Private Const MAX_ADD As Long = 80

' ----------------------------------------------------------------------------
' NeighborExpand - hits() の各ヒットの前後 radius チャンクを末尾へ足す。
'   radius <= 0 / ヒット0件 / シートが読めない ときは何もしない(無操作)。
'   nHits は足したぶんだけ増える(呼び出し元は増えた件数で本棚抜粋を組む)。
' ----------------------------------------------------------------------------
Public Sub NeighborExpand(ByRef hits() As Hit, ByRef nHits As Long, ByVal radius As Long)
    If radius < 1 Then Exit Sub
    If nHits < 1 Then Exit Sub

    On Error GoTo Quiet

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then Exit Sub

    ' 3行目(=データ2件)より少ない本棚には「隣のチャンク」が存在しない。
    ' ついでに、1セルだけの Range.Value が配列にならないVBAの罠も避けられる。
    Dim lastK As Long
    lastK = ws.Cells(ws.Rows.count, COL_ID).End(xlUp).row
    If lastK < 3 Then Exit Sub

    ' 1) 今回のヒットから「どの資料の・どのページ周辺を見ればよいか」を作る。
    Dim hitIds As String, needKeys As String, winLine As String
    If Not BuildWindows(hits, nHits, radius, hitIds, needKeys, winLine) Then Exit Sub

    ' 2) chunk_id 列だけを1回で読み、窓に入る行だけを候補に残す。
    Dim idData As Variant
    idData = ws.Range(ws.Cells(2, COL_ID), ws.Cells(lastK, COL_ID)).Value

    Dim candIds As String
    Dim candRow() As Long
    Dim candN As Long
    candN = CollectCandidates(idData, needKeys, winLine, candIds, candRow)
    If candN < 1 Then Exit Sub

    ' 3) 文書順に並べ、各ヒットの前後 radius を拾う(純ロジック)。
    Dim addLine As String
    addLine = NeighborIdList(candIds, hitIds, radius)
    If LenB(addLine) = 0 Then Exit Sub

    ' 4) 採用した行だけを読み出して末尾へ足す。
    Dim adds() As String
    adds = Split(addLine, vbLf)
    Dim added As Long
    Dim i As Long
    For i = LBound(adds) To UBound(adds)
        If added >= MAX_ADD Then Exit For
        Dim r As Long
        r = RowOfId(candIds, candRow, candN, adds(i))
        If r > 0 Then
            ReDim Preserve hits(1 To nHits + 1)
            nHits = nHits + 1
            added = added + 1
            FillHitFromRow ws, r, hits(nHits)
        End If
    Next i

    On Error Resume Next
    modLog.LogUsage "neighbor_expand", "focus", _
        "radius=" & radius & " added=" & added, 0, nHits
    On Error GoTo 0
    Exit Sub

Quiet:
    ' 本棚シートが無い/形が違う/読めない。精読を諦めるだけで回答は続ける。
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きたエラーが
    ' 呼び出し元へ飛んで本来の原因を上書きするので、必ず Resume で抜けてから
    ' 記録する(2026-07-30 実機err#462 と同型の作法)。
    Resume QuietDone
QuietDone:
    Err.Clear
    On Error Resume Next
    modLog.LogUsage "neighbor_skip", "focus", "radius=" & radius
    On Error GoTo 0
End Sub

' ============================================================================
' 純ロジック(LibreOfficeの実行テストで固定する)
' ============================================================================

' ----------------------------------------------------------------------------
' ParseChunkKey - chunk_id "bs::ハッシュ::pN::cM" を分解する。
'   docKey = 末尾2要素を除いた全部(=資料を一意に決める部分)。
'   pageNo / seqNo = pN / cM の数値。読めなければ False(呼び出し側は無視する)。
'
'   末尾から2つを取るのは、ハッシュ自体に "::" が入っても壊れないようにするため
'   (先頭から4要素と決め打つと、その1件だけ文書順から落ちて近傍が片側に寄る)。
' ----------------------------------------------------------------------------
Public Function ParseChunkKey(ByVal chunkId As String, ByRef docKey As String, _
                              ByRef pageNo As Long, ByRef seqNo As Long) As Boolean
    docKey = ""
    pageNo = 0
    seqNo = 0

    Dim t As String: t = Trim$(chunkId)
    If LenB(t) = 0 Then Exit Function

    Dim parts() As String: parts = Split(t, "::")
    Dim lo As Long: lo = LBound(parts)
    Dim n As Long: n = UBound(parts) - lo + 1
    If n < 4 Then Exit Function

    Dim sp As String: sp = Trim$(parts(lo + n - 2))
    Dim sc As String: sc = Trim$(parts(lo + n - 1))
    If LCase$(Left$(sp, 1)) <> "p" Then Exit Function
    If LCase$(Left$(sc, 1)) <> "c" Then Exit Function
    If Not IsDigitRun(Mid$(sp, 2)) Then Exit Function
    If Not IsDigitRun(Mid$(sc, 2)) Then Exit Function

    Dim i As Long
    For i = lo To lo + n - 3
        If LenB(docKey) > 0 Then docKey = docKey & "::"
        docKey = docKey & parts(i)
    Next i
    If LenB(docKey) = 0 Then Exit Function

    pageNo = CLng(Val(Mid$(sp, 2)))
    seqNo = CLng(Val(Mid$(sc, 2)))
    ParseChunkKey = True
End Function

' ----------------------------------------------------------------------------
' NeighborIdList - 候補チャンクを文書順に並べ、ヒットの前後 radius を選ぶ。
'   orderedIds: 候補の chunk_id(vbLf区切り。並び順は問わない)
'   hitIds    : ヒットの chunk_id(vbLf区切り)
'   戻り値    : 足すべき chunk_id(vbLf区切り・文書順・重複なし・ヒット自身は除く)
'
'   ・(docKey, page, seq) の昇順に並べてから前後を取る(ページ跨ぎでも
'     seq のリセットに引きずられない)。
'   ・資料の境界は越えない(別の資料の先頭が「前のチャンク」にならない)。
'   ・窓が重なっても結果は重複しない(印を付けてから1回だけ書き出す)。
'   ・読めない chunk_id は並びから外す(位置が決められないものを混ぜると、
'     その分だけ隣が1つずれる)。
' ----------------------------------------------------------------------------
Public Function NeighborIdList(ByVal orderedIds As String, ByVal hitIds As String, _
                               ByVal radius As Long) As String
    If radius < 1 Then Exit Function
    If LenB(orderedIds) = 0 Or LenB(hitIds) = 0 Then Exit Function

    Dim raw() As String: raw = Split(orderedIds, vbLf)
    Dim cap As Long: cap = UBound(raw) - LBound(raw) + 1
    If cap < 1 Then Exit Function

    Dim ids() As String, dk() As String
    Dim pg() As Long, sq() As Long
    ReDim ids(1 To cap)
    ReDim dk(1 To cap)
    ReDim pg(1 To cap)
    ReDim sq(1 To cap)

    Dim n As Long: n = 0
    Dim i As Long, j As Long
    For i = LBound(raw) To UBound(raw)
        Dim d As String, p As Long, s As Long
        If ParseChunkKey(raw(i), d, p, s) Then
            n = n + 1
            ids(n) = Trim$(raw(i))
            dk(n) = d
            pg(n) = p
            sq(n) = s
        End If
    Next i
    If n < 1 Then Exit Function

    ' 挿入ソート((docKey, page, seq) 昇順)。候補はヒット周辺だけに絞って
    ' 渡される前提なので数百件が上限で、単純な安定ソートで十分。
    For i = 2 To n
        Dim ti As String, td As String
        Dim tp As Long, ts As Long
        ti = ids(i): td = dk(i): tp = pg(i): ts = sq(i)
        j = i - 1
        Do While j >= 1
            If KeyGreater(dk(j), pg(j), sq(j), td, tp, ts) Then
                ids(j + 1) = ids(j): dk(j + 1) = dk(j)
                pg(j + 1) = pg(j): sq(j + 1) = sq(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        ids(j + 1) = ti: dk(j + 1) = td: pg(j + 1) = tp: sq(j + 1) = ts
    Next i

    Dim mark() As Boolean
    ReDim mark(1 To n)
    Dim hitBox As String: hitBox = vbLf & hitIds & vbLf
    For i = 1 To n
        If InStr(1, hitBox, vbLf & ids(i) & vbLf, vbBinaryCompare) > 0 Then
            For j = i - radius To i + radius
                If j >= 1 And j <= n Then
                    If dk(j) = dk(i) Then mark(j) = True
                End If
            Next j
        End If
    Next i

    Dim outS As String
    For i = 1 To n
        If mark(i) Then
            If InStr(1, hitBox, vbLf & ids(i) & vbLf, vbBinaryCompare) = 0 Then
                If InStr(1, vbLf & outS & vbLf, vbLf & ids(i) & vbLf, vbBinaryCompare) = 0 Then
                    If LenB(outS) > 0 Then outS = outS & vbLf
                    outS = outS & ids(i)
                End If
            End If
        End If
    Next i
    NeighborIdList = outS
End Function

' (docKey, page, seq) の大小比較。a > b なら True。
Private Function KeyGreater(ByVal ad As String, ByVal ap As Long, ByVal asq As Long, _
                            ByVal bd As String, ByVal bp As Long, ByVal bsq As Long) As Boolean
    Dim c As Long
    c = StrComp(ad, bd, vbBinaryCompare)
    If c <> 0 Then
        KeyGreater = (c > 0)
        Exit Function
    End If
    If ap <> bp Then
        KeyGreater = (ap > bp)
        Exit Function
    End If
    KeyGreater = (asq > bsq)
End Function

' 半角数字だけで1文字以上あるか(chunk_id の pN / cM の数値部の検査)。
Private Function IsDigitRun(ByVal s As String) As Boolean
    If LenB(s) = 0 Then Exit Function
    Dim i As Long
    For i = 1 To Len(s)
        Dim ch As String: ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    IsDigitRun = True
End Function

' ============================================================================
' 内部ヘルパー(シート読み)
' ============================================================================

' ヒットから「探す資料(docKey)」と「見るページ窓(docKey#page)」を作る。
' 窓を作るのは、本棚全体の chunk_id を並べ替えないため。radius が何チャンク
' でも、離れるページ数は radius を超えない(1ページに最低1チャンクあるため)。
Private Function BuildWindows(hits() As Hit, ByVal nHits As Long, ByVal radius As Long, _
                              ByRef hitIds As String, ByRef needKeys As String, _
                              ByRef winLine As String) As Boolean
    Dim i As Long, p As Long
    For i = 1 To nHits
        Dim d As String, pg As Long, sq As Long
        If ParseChunkKey(hits(i).chunk_id, d, pg, sq) Then
            If LenB(hitIds) > 0 Then hitIds = hitIds & vbLf
            hitIds = hitIds & Trim$(hits(i).chunk_id)
            If InStr(1, needKeys, "|" & d & "|", vbBinaryCompare) = 0 Then
                needKeys = needKeys & "|" & d & "|"
            End If
            For p = pg - radius To pg + radius
                If p >= 0 Then
                    Dim w As String: w = "|" & d & "#" & p & "|"
                    If InStr(1, winLine, w, vbBinaryCompare) = 0 Then winLine = winLine & w
                End If
            Next p
        End If
    Next i
    BuildWindows = (LenB(hitIds) > 0)
End Function

' chunk_id 列を1回だけ走査し、窓に入る行の id と行番号を集める。
Private Function CollectCandidates(ByVal idData As Variant, ByVal needKeys As String, _
                                   ByVal winLine As String, ByRef candIds As String, _
                                   ByRef candRow() As Long) As Long
    Dim lo As Long: lo = LBound(idData, 1)
    Dim hi As Long: hi = UBound(idData, 1)
    ReDim candRow(1 To hi - lo + 1)

    Dim n As Long: n = 0
    Dim i As Long
    For i = lo To hi
        Dim id As String: id = CStr(idData(i, 1))
        If LenB(id) > 0 Then
            Dim d As String, pg As Long, sq As Long
            If ParseChunkKey(id, d, pg, sq) Then
                If InStr(1, needKeys, "|" & d & "|", vbBinaryCompare) > 0 Then
                    If InStr(1, winLine, "|" & d & "#" & pg & "|", vbBinaryCompare) > 0 Then
                        n = n + 1
                        candRow(n) = i + 1          ' idData は2行目起点
                        If LenB(candIds) > 0 Then candIds = candIds & vbLf
                        candIds = candIds & id
                    End If
                End If
            End If
        End If
    Next i
    CollectCandidates = n
End Function

' 候補リストの中から chunk_id に対応するシート行番号を引く(見つからなければ0)。
Private Function RowOfId(ByVal candIds As String, candRow() As Long, ByVal candN As Long, _
                         ByVal id As String) As Long
    If LenB(id) = 0 Then Exit Function
    Dim arr() As String: arr = Split(candIds, vbLf)
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If i - LBound(arr) + 1 > candN Then Exit For
        If StrComp(arr(i), id, vbBinaryCompare) = 0 Then
            RowOfId = candRow(i - LBound(arr) + 1)
            Exit Function
        End If
    Next i
End Function

' 近傍チャンク1件を Hit へ組む。score=0(モジュール冒頭の設計判断)。
' source / page / origin は近傍チャンク自身の値を使う(出典タグの整合)。
Private Sub FillHitFromRow(ByVal ws As Worksheet, ByVal r As Long, ByRef h As Hit)
    Dim fullText As String
    fullText = CStr(ws.Cells(r, COL_FULLTEXT).Value)

    h.chunk_id = CStr(ws.Cells(r, COL_ID).Value)
    h.source = CStr(ws.Cells(r, COL_SOURCE).Value)
    h.origin = CStr(ws.Cells(r, COL_ORIGIN).Value)
    h.page = CLng(Val(CStr(ws.Cells(r, COL_PAGE).Value)))
    h.full_text = fullText
    h.preview = modUtil.SafeLeft(fullText, PREVIEW_LEN)
    h.score = 0
End Sub
