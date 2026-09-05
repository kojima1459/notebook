Attribute VB_Name = "modXDocStore"
Option Explicit

' ============================================================================
' modXDocStore - 資料間リンク(R37 §3)の受け皿(シートI/Oと純ロジック)。
' ----------------------------------------------------------------------------
' 持っている地図は2枚:
'   ・doc_centroids(source, chap, vector_csv, chunk_n)
'       その章に属する全チャンクのベクトルの平均をL2正規化した「章の重心」。
'   ・doc_links(src_a, chap_a, src_b, chap_b, sim, built_at)
'       (src_a,chap_a) から見て近い【他資料の章】を sim 降順で最大3件。
'       双方向は別行として持つ(a→b と b→a を別に書く)。
'
' なぜフルGraphRAGにしないか(R37 §0): エンティティ抽出・コミュニティ要約は
'   取込1資料あたり数百回のLLM呼び出しになり、Excel 32bit VBA では割に合わない。
'   章は既に人の手で資料に入っていて、ベクトルも既にある。だから
'   「章の重心どうしのコサイン」だけで地図を作る = LLM 0回で済む。
'
' 失敗の扱い(modOutlineStore と同じ線引き): この2枚は上積みで、無くても
'   取込も検索も回答も従来どおり動く。失敗は握って usage_log に1行だけ残す。
'   config xdoc_links=off で計算そのものが起きない。
'
' 引き継ぎ・パックには載せない(R37 §3-2 記録): 取込から再計算できる派生
'   データなので modMigrate/modPack は触らない。版上げ直後は空になるが、
'   次の取込・同期でその資料ぶんが作り直される(全量再構築は R38 送り)。
'
' 【容量の裁定・2026-09-05】BuildLinksFor の実体(章重心の計算・総当たり・
'   doc_links の作り直し)はこのモジュールには入らなかった(I/Oと純ロジック
'   だけで既に上限の半分を超える)。実体は modXDocBuild へ置き、ここからは
'   1行で呼ぶ(憲章§12 の分割の型)。公開契約 BuildLinksFor の名前と場所は
'   R37 §3-2 のとおり据え置く。
' ============================================================================

' doc_centroids の列
Private Const COL_C_SRC As Long = 1
Private Const COL_C_CHAP As Long = 2
Private Const COL_C_VEC As Long = 3
Private Const COL_C_N As Long = 4
Public Const CENT_COLS As Long = 4

' doc_links の列
Private Const COL_L_SA As Long = 1
Private Const COL_L_CA As Long = 2
Private Const COL_L_SB As Long = 3
Private Const COL_L_CB As Long = 4
Private Const COL_L_SIM As Long = 5
Private Const COL_L_AT As Long = 6
Public Const LINK_COLS As Long = 6

' 1つの(資料,章)から保持する相手の数。増やすと回答時の候補が広がり、
' 「関係ない資料が混ざる」危険を閾値だけでは抑えきれなくなる(R37 §7)。
Public Const KEEP_N As Long = 3

' 重心CSVの小数桁。768次元で約7KB=1セル上限(32,767字)に収まり、
' コサインへの影響は1e-5未満(modUtil.VectorToCsvPrec の設計注記)。
Private Const CENT_DEC As Long = 6

' ----------------------------------------------------------------------------
' GateOn - config xdoc_links(既定on)。off のときだけ切る(graph_outline と
'   同じ作法)。取込時の計算と回答時の末尾補充(modXDoc)の両方がこの1本を
'   見るので、「片方だけ止まっている」状態が構造的に作れない。
' ----------------------------------------------------------------------------
Public Function GateOn() As Boolean
    GateOn = (LCase$(Trim$(modConfig.GetString("xdoc_links", "on"))) <> "off")
End Function

' 総当たりの上限章数(config xdoc_max_chapters・既定4,000=資料200冊×20章)。
' 超えたら既存側の上位3件の作り直しはやらず新資料側だけ書く(R37 §7)。
Public Function MaxChapters() As Long
    Dim v As Long: v = 4000
    On Error Resume Next
    v = modConfig.GetLong("xdoc_max_chapters", 4000)
    On Error GoTo 0
    If v < 1 Then v = 4000
    MaxChapters = v
End Function

' ----------------------------------------------------------------------------
' BuildLinksFor - 取込時の唯一の入口(実体は modXDocBuild。上の容量裁定)。
'   呼び出しは modOutlineBuild.BuildOutlineFor の末尾から1行だけ。
' ----------------------------------------------------------------------------
Public Sub BuildLinksFor(ByVal sourceName As String)
    If Not GateOn() Then Exit Sub
    modXDocBuild.BuildFor sourceName
End Sub

' ============================================================================
' 純ロジック(LibreOffice の実行テストで固定する = modTestsPure42)
' ============================================================================

' ----------------------------------------------------------------------------
' MeanNormalizedCsv - ベクトルCSVを n 本受け取り、平均→L2正規化して返す。
'   ・次元は「最初に読めた行」で確定(modVecCache.BuildFrom と同じ規約)。
'     違う次元の行は数に入れない(混ざると平均が意味を失う)。
'   ・1本も読めない/ゼロベクトルになるなら空文字(=「重心が作れない」という
'     正しい答え。呼び出し元はその章を書かない)。
'   ・配列が0始まりでも1始まりでも動くよう LBound から相対で読む。
' ----------------------------------------------------------------------------
Public Function MeanNormalizedCsv(ByRef csvRows() As String, ByVal n As Long) As String
    If n < 1 Then Exit Function

    Dim lo As Long
    On Error Resume Next
    lo = LBound(csvRows)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Dim acc() As Double
    Dim dimN As Long
    Dim used As Long
    Dim v() As Double
    Dim i As Long, j As Long
    For i = 0 To n - 1
        If modUtil.CsvToVector(csvRows(lo + i), v) Then
            Dim d As Long: d = UBound(v) - LBound(v) + 1
            If dimN = 0 And d > 0 Then
                dimN = d
                ReDim acc(0 To dimN - 1)
            End If
            If d = dimN Then
                Dim vb0 As Long: vb0 = LBound(v)
                For j = 0 To dimN - 1
                    acc(j) = acc(j) + v(vb0 + j)
                Next j
                used = used + 1
            End If
        End If
    Next i
    If used < 1 Then Exit Function

    For j = 0 To dimN - 1
        acc(j) = acc(j) / used
    Next j
    If Not modUtil.L2Normalize(acc) Then Exit Function
    MeanNormalizedCsv = modUtil.VectorToCsvPrec(acc, CENT_DEC)
End Function

' ----------------------------------------------------------------------------
' TopNLinks - (相手キー, 類似度)の並びから上位 keep 件を選び、
'   「相手キー|類似度」の行を vbLf 区切りで返す。
'   同点は【先に並んでいたほう】が勝つ(比較に > を使い >= を使わない)。
'   同じ本棚・同じ順序なら毎回同じ相手が選ばれる=見た目が揺れない。
'   pairs/sims は同じ LBound 起点の並行配列。keep<=0 / n<=0 は空文字。
' ----------------------------------------------------------------------------
Public Function TopNLinks(ByRef pairs() As String, ByRef sims() As Double, _
                          ByVal n As Long, ByVal keep As Long) As String
    If n < 1 Then Exit Function
    If keep < 1 Then Exit Function

    Dim lp As Long, ls As Long
    On Error Resume Next
    lp = LBound(pairs): ls = LBound(sims)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Dim takeN As Long: takeN = keep
    If takeN > n Then takeN = n

    Dim takenFlag() As Boolean: ReDim takenFlag(0 To n - 1)
    Dim outArr() As String: ReDim outArr(0 To takeN - 1)
    Dim got As Long
    Dim r As Long, i As Long
    For r = 0 To takeN - 1
        Dim best As Long: best = -1
        Dim bestV As Double: bestV = 0
        For i = 0 To n - 1
            If Not takenFlag(i) Then
                If best < 0 Then
                    best = i
                    bestV = sims(ls + i)
                ElseIf sims(ls + i) > bestV Then
                    best = i
                    bestV = sims(ls + i)
                End If
            End If
        Next i
        If best < 0 Then Exit For
        takenFlag(best) = True
        If LenB(pairs(lp + best)) > 0 Then
            outArr(got) = pairs(lp + best) & "|" & Trim$(Str$(bestV))
            got = got + 1
        End If
    Next r
    If got < 1 Then Exit Function
    ReDim Preserve outArr(0 To got - 1)
    TopNLinks = Join(outArr, vbLf)
End Function

' ============================================================================
' シートI/O(自己修復・一括読み・一括書き)
' ============================================================================

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

' EnsureSheetOf - 無ければ作る+見出し+veryHidden、既存ならそのまま(冪等)。
'   通常は build_mybookshelf.py が焼き込むので必ず見つかる。ここでの Add は
'   「壊れたブックの自己修復」専用(modOutlineStore.EnsureOutlineSheet と同型)。
'   textCols 列までを "@" に固定して数式インジェクションを防ぐ
'   (資料名も章キーも "=" や "-" で始まりうる)。
Private Function EnsureSheetOf(ByVal nm As String, ByVal hdr As Variant, _
                               ByVal textCols As Long) As Worksheet
    Dim ws As Worksheet: Set ws = GetSheet(nm)
    Dim i As Long
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = nm
        For i = LBound(hdr) To UBound(hdr)
            ws.Cells(1, i + 1).Value = hdr(i)
        Next i
    End If
    On Error Resume Next
    ws.Visible = 2                      ' xlSheetVeryHidden
    For i = 1 To textCols
        ws.Columns(i).NumberFormat = "@"
    Next i
    On Error GoTo 0
    Set EnsureSheetOf = ws
    Exit Function
Fail:
    Set EnsureSheetOf = Nothing
End Function

Public Function EnsureCentroidSheet() As Worksheet
    Set EnsureCentroidSheet = EnsureSheetOf(modAppDef.SH_DOC_CENTROIDS, _
        Array("source", "chap", "vector_csv", "chunk_n"), 3)
End Function

Public Function EnsureLinksSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = EnsureSheetOf(modAppDef.SH_DOC_LINKS, _
        Array("src_a", "chap_a", "src_b", "chap_b", "sim", "built_at"), 4)
    ' sim(5列目)だけは数値。built_at(6列目)は日付に化けさせない=文字列固定
    ' (EnsureSheetOf は先頭からの連続列しか見ないので1列だけここで足す)。
    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Columns(COL_L_AT).NumberFormat = "@"
        On Error GoTo 0
    End If
    Set EnsureLinksSheet = ws
End Function

' ----------------------------------------------------------------------------
' ReadCentroids - doc_centroids の全行を並行配列(0 To n-1)へ。件数を返す。
'   シートが無い/0行でも呼び出し側は必ず ReDim 済みの配列を受け取る規約
'   (modOutlineStore.ReadOutline と同じ)。
' ----------------------------------------------------------------------------
Public Function ReadCentroids(ByRef outSrcs() As String, ByRef outChaps() As String, _
                              ByRef outVecs() As String) As Long
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_DOC_CENTROIDS)
    Dim lastR As Long
    If Not ws Is Nothing Then lastR = ws.Cells(ws.Rows.count, COL_C_SRC).End(xlUp).row

    If ws Is Nothing Or lastR < 2 Then
        ReDim outSrcs(0 To 0)
        ReDim outChaps(0 To 0)
        ReDim outVecs(0 To 0)
        Exit Function
    End If

    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, COL_C_SRC), ws.Cells(lastR, COL_C_VEC)).Value
    Dim n As Long: n = lastR - 1
    ReDim outSrcs(0 To n - 1)
    ReDim outChaps(0 To n - 1)
    ReDim outVecs(0 To n - 1)

    Dim i As Long
    For i = 1 To n
        outSrcs(i - 1) = CStr(arr(i, COL_C_SRC))
        outChaps(i - 1) = CStr(arr(i, COL_C_CHAP))
        outVecs(i - 1) = CStr(arr(i, COL_C_VEC))
    Next i
    ReadCentroids = n
End Function

' ReadLinks - doc_links の全行を並行配列(0 To n-1)へ。件数を返す。
Public Function ReadLinks(ByRef outA() As String, ByRef outCA() As String, _
                          ByRef outB() As String, ByRef outCB() As String, _
                          ByRef outSim() As Double) As Long
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_DOC_LINKS)
    Dim lastR As Long
    If Not ws Is Nothing Then lastR = ws.Cells(ws.Rows.count, COL_L_SA).End(xlUp).row

    If ws Is Nothing Or lastR < 2 Then
        ReDim outA(0 To 0)
        ReDim outCA(0 To 0)
        ReDim outB(0 To 0)
        ReDim outCB(0 To 0)
        ReDim outSim(0 To 0)
        Exit Function
    End If

    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, COL_L_SA), ws.Cells(lastR, COL_L_SIM)).Value
    Dim n As Long: n = lastR - 1
    ReDim outA(0 To n - 1)
    ReDim outCA(0 To n - 1)
    ReDim outB(0 To n - 1)
    ReDim outCB(0 To n - 1)
    ReDim outSim(0 To n - 1)

    Dim i As Long
    For i = 1 To n
        outA(i - 1) = CStr(arr(i, COL_L_SA))
        outCA(i - 1) = CStr(arr(i, COL_L_CA))
        outB(i - 1) = CStr(arr(i, COL_L_SB))
        outCB(i - 1) = CStr(arr(i, COL_L_CB))
        outSim(i - 1) = CDbl(Val(CStr(arr(i, COL_L_SIM))))
    Next i
    ReadLinks = n
End Function

' ----------------------------------------------------------------------------
' LinksCsvFor - 起点キー(vbLf区切りの "資料|章" の箱)に当たる doc_links の
'   行を "src_a|chap_a|src_b|chap_b|sim" で vbLf 連結して返す(modXDoc.Expand)。
'   ここは読むだけ。閾値も件数も見ない(判断は純関数 modXDoc.PickLinked)。
' ----------------------------------------------------------------------------
Public Function LinksCsvFor(ByVal keyBox As String) As String
    If LenB(keyBox) = 0 Then Exit Function
    On Error GoTo Quiet

    Dim la() As String, lca() As String, lb() As String, lcb() As String
    Dim lsim() As Double
    Dim n As Long: n = ReadLinks(la, lca, lb, lcb, lsim)
    If n < 1 Then Exit Function

    Dim outArr() As String: ReDim outArr(0 To n - 1)
    Dim got As Long
    Dim i As Long
    For i = 0 To n - 1
        Dim k As String: k = Trim$(la(i)) & "|" & Trim$(lca(i))
        If InStr(1, keyBox, vbLf & k & vbLf, vbTextCompare) > 0 Then
            outArr(got) = k & "|" & Trim$(lb(i)) & "|" & Trim$(lcb(i)) & _
                          "|" & Trim$(Str$(lsim(i)))
            got = got + 1
        End If
    Next i
    If got < 1 Then Exit Function
    ReDim Preserve outArr(0 To got - 1)
    LinksCsvFor = Join(outArr, vbLf)
    Exit Function
Quiet:
    Resume LinksDone
LinksDone:
    Err.Clear
End Function

' ----------------------------------------------------------------------------
' LinkedLabel - 📖本文のメタ行に出す「関連する資料: B(第3章)、C(第1章)」。
'   その資料を起点にした doc_links を sim 降順に並べ、同じ相手資料は最初の
'   1件だけ採って最大 KEEP_N 件。1件も無ければ空文字=呼び出し元は行を出さない。
' ----------------------------------------------------------------------------
Public Function LinkedLabel(ByVal sourceName As String) As String
    If LenB(Trim$(sourceName)) = 0 Then Exit Function
    If Not GateOn() Then Exit Function
    On Error GoTo Quiet

    Dim la() As String, lca() As String, lb() As String, lcb() As String
    Dim lsim() As Double
    Dim n As Long: n = ReadLinks(la, lca, lb, lcb, lsim)
    If n < 1 Then Exit Function

    Dim pairs() As String: ReDim pairs(0 To n - 1)
    Dim sims() As Double: ReDim sims(0 To n - 1)
    Dim m As Long
    Dim i As Long
    For i = 0 To n - 1
        If StrComp(Trim$(la(i)), Trim$(sourceName), vbTextCompare) = 0 Then
            pairs(m) = Trim$(lb(i)) & "|" & Trim$(lcb(i))
            sims(m) = lsim(i)
            m = m + 1
        End If
    Next i
    If m < 1 Then Exit Function

    ' 並べ替えの規則は1本(TopNLinks)に閉じる。相手資料の重複を落とすため、
    ' 全件を降順に並べてから先着で KEEP_N 件だけ採る。
    Dim lines0 As String: lines0 = TopNLinks(pairs, sims, m, m)
    If LenB(lines0) = 0 Then Exit Function

    Dim rowsArr() As String: rowsArr = Split(lines0, vbLf)
    Dim seen As String, outS As String
    Dim shown As Long
    For i = LBound(rowsArr) To UBound(rowsArr)
        If shown >= KEEP_N Then Exit For
        Dim f() As String: f = Split(rowsArr(i), "|")
        If UBound(f) >= 1 Then
            Dim sb As String: sb = Trim$(f(0))
            If LenB(sb) > 0 Then
                If InStr(1, seen, vbLf & sb & vbLf, vbTextCompare) = 0 Then
                    seen = seen & vbLf & sb & vbLf
                    If shown > 0 Then outS = outS & ChrW(&H3001)
                    outS = outS & sb & "(" & Trim$(f(1)) & ")"
                    shown = shown + 1
                End If
            End If
        End If
    Next i
    LinkedLabel = outS
    Exit Function
Quiet:
    Resume LabelDone
LabelDone:
    Err.Clear
End Function

' ----------------------------------------------------------------------------
' RemoveFor - その資料の doc_centroids/doc_links を落とす(両側)。
'   doc_links は src_a 側だけでなく src_b 側の行も落とす。片側だけ消すと
'   「もう本棚に無い資料へのリンク」が残り、回答時に候補と1件も当たらない
'   死んだ行を毎回引くことになる。
'   ※相手側の上位3件は補充しない(2位以下へ繰り上げない)。削除1回のために
'     本棚全体の重心を読み直すのは重すぎ、次にその資料を取り込み直した時点で
'     正しい上位3件に戻るため(記録・R37 §3-2)。
'   呼ばれる場所: modOutlineStore.RemoveOutlineForSource(doc_outline の掃除と
'   同じ1本。再取込・資料削除・章要約の作り直しの3経路が必ずそこを通る)。
' ----------------------------------------------------------------------------
Public Sub RemoveFor(ByVal sourceName As String)
    If LenB(Trim$(sourceName)) = 0 Then Exit Sub
    On Error GoTo Quiet
    RemoveCentroidsFor sourceName
    RemoveLinksFor sourceName
    Exit Sub
Quiet:
    Resume RemoveDone
RemoveDone:
    Err.Clear
    XDocFail "remove"
End Sub

' 1行ずつの Delete は数千行で固まるので、残す行を詰め直して1回で書き戻す
' (modOutlineStore.RemoveOutlineForSource と同じ作法)。
' matchB=True のとき src_b 側の一致でも落とす(doc_links 用)。
Public Sub RemoveCentroidsFor(ByVal sourceName As String)
    RemoveRowsFor modAppDef.SH_DOC_CENTROIDS, CENT_COLS, sourceName, 0
End Sub

Public Sub RemoveLinksFor(ByVal sourceName As String)
    RemoveRowsFor modAppDef.SH_DOC_LINKS, LINK_COLS, sourceName, COL_L_SB
End Sub

Private Sub RemoveRowsFor(ByVal nm As String, ByVal nCols As Long, _
                          ByVal sourceName As String, ByVal alsoCol As Long)
    Dim ws As Worksheet: Set ws = GetSheet(nm)
    If ws Is Nothing Then Exit Sub
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then Exit Sub

    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastR, nCols)).Value
    Dim nRows As Long: nRows = lastR - 1

    Dim keepArr() As Variant: ReDim keepArr(1 To nRows, 1 To nCols)
    Dim k As Long
    Dim i As Long, c As Long
    For i = 1 To nRows
        Dim drop As Boolean
        drop = (StrComp(Trim$(CStr(arr(i, 1))), sourceName, vbTextCompare) = 0)
        If Not drop And alsoCol > 0 Then
            drop = (StrComp(Trim$(CStr(arr(i, alsoCol))), sourceName, vbTextCompare) = 0)
        End If
        If Not drop Then
            k = k + 1
            For c = 1 To nCols
                keepArr(k, c) = arr(i, c)
            Next c
        End If
    Next i
    If k = nRows Then Exit Sub

    ws.Range(ws.Cells(2, 1), ws.Cells(lastR, nCols)).ClearContents
    If k > 0 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(1 + k, nCols)).Value = Compact2D(keepArr, k, nCols)
    End If
End Sub

' 先頭k行だけの2次元配列(書き戻す範囲と配列の形をぴったり合わせる)。
Private Function Compact2D(ByRef src() As Variant, ByVal k As Long, ByVal nCols As Long) As Variant
    Dim outArr() As Variant: ReDim outArr(1 To k, 1 To nCols)
    Dim r As Long, c As Long
    For r = 1 To k
        For c = 1 To nCols
            outArr(r, c) = src(r, c)
        Next c
    Next r
    Compact2D = outArr
End Function

' ----------------------------------------------------------------------------
' WriteCentroidRows - その資料の章重心を末尾へ1回の Range 書込みで追記する。
' ----------------------------------------------------------------------------
Public Sub WriteCentroidRows(ByVal sourceName As String, ByRef chaps() As String, _
                             ByRef csvs() As String, ByRef ns() As Long, ByVal n As Long)
    If n < 1 Then Exit Sub
    Dim ws As Worksheet: Set ws = EnsureCentroidSheet()
    If ws Is Nothing Then Exit Sub

    Dim arr() As Variant: ReDim arr(1 To n, 1 To CENT_COLS)
    Dim i As Long
    For i = 1 To n
        arr(i, COL_C_SRC) = sourceName
        arr(i, COL_C_CHAP) = chaps(i)
        arr(i, COL_C_VEC) = csvs(i)
        arr(i, COL_C_N) = ns(i)
    Next i

    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, COL_C_SRC).End(xlUp).row
    If lastR < 1 Then lastR = 1
    Dim firstRow As Long: firstRow = lastR + 1
    If firstRow < 2 Then firstRow = 2
    ws.Range(ws.Cells(firstRow, COL_C_SRC), ws.Cells(firstRow + n - 1, COL_C_N)).Value = arr
End Sub

' ----------------------------------------------------------------------------
' WriteAllLinks - doc_links を作り直す(既存を消してから1回の Range 書込み)。
'   行の並びと中身を決めるのは modXDocBuild。ここは「シートへ落とす」だけ。
' ----------------------------------------------------------------------------
Public Sub WriteAllLinks(ByRef oA() As String, ByRef oCA() As String, ByRef oB() As String, _
                         ByRef oCB() As String, ByRef oS() As Double, ByVal nOut As Long)
    Dim ws As Worksheet: Set ws = EnsureLinksSheet()
    If ws Is Nothing Then Exit Sub

    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, COL_L_SA).End(xlUp).row
    If lastR >= 2 Then
        ws.Range(ws.Cells(2, COL_L_SA), ws.Cells(lastR, COL_L_AT)).ClearContents
    End If
    If nOut < 1 Then Exit Sub

    Dim stamp As String: stamp = modUtil.NowStamp()
    Dim arr() As Variant: ReDim arr(1 To nOut, 1 To LINK_COLS)
    Dim i As Long
    For i = 1 To nOut
        arr(i, COL_L_SA) = oA(i)
        arr(i, COL_L_CA) = oCA(i)
        arr(i, COL_L_SB) = oB(i)
        arr(i, COL_L_CB) = oCB(i)
        arr(i, COL_L_SIM) = oS(i)
        arr(i, COL_L_AT) = stamp
    Next i
    ws.Range(ws.Cells(2, COL_L_SA), ws.Cells(1 + nOut, COL_L_AT)).Value = arr
End Sub

' 失敗の記録。Err はここへ来た時点の値をまず控える
' (On Error Resume Next 自体が Err を消す。憲章§11)。
Public Sub XDocFail(ByVal whereAt As String)
    Dim d As String: d = "err#" & Err.Number & " " & Err.Description
    On Error Resume Next
    modLog.LogUsage "xdoc_fail", "ingest", whereAt & " " & d
    On Error GoTo 0
End Sub
