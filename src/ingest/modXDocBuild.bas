Attribute VB_Name = "modXDocBuild"
Option Explicit

' ============================================================================
' modXDocBuild - 資料間リンク(R37 §3)の【作る側】の実体。
' ----------------------------------------------------------------------------
' 入口は BuildFor の1本だけで、呼ぶのは modXDocStore.BuildLinksFor
' (そこからさらに modOutlineBuild.BuildOutlineFor の末尾1行へ繋がる)。
' なぜ modXDocStore と分かれているか: シートI/Oと純ロジックだけで上限
' 30,000字の半分を超え、章重心の計算・総当たり・doc_links の作り直しが
' 同居できなかった(憲章§12 の分割。契約名 BuildLinksFor の場所は据え置き)。
'
' 手順は4つだけ。LLM は1回も呼ばない:
'   1) この資料の章ごとの重心を作る(章キーは modOutlineBuild.ChapterKeyOf)。
'   2) 自分の古い行(重心・リンク両側)を落として今回ぶんの重心を書く。
'   3) 既存資料の章重心を平坦なDouble配列へ1回だけ展開し、内積で総当たり。
'   4) 新資料側の上位3件を書き、相手側の上位3件も作り直す。
'
' 失敗は全部この層で握る(usage_log "xdoc_fail" 1行)。取込は絶対に止めない
' ―― この地図は上積みで、無くても取込も検索も回答も従来どおり動く。
' ============================================================================

' my_knowledge / my_vectors の列(ここは読むだけ)。
Private Const COL_K_ID As Long = 1
Private Const COL_K_SRC As Long = 2
Private Const COL_V_ID As Long = 1
Private Const COL_V_CSV As Long = 2

' ----------------------------------------------------------------------------
' BuildFor - 1資料ぶんのリンクを作り直す。
'   完全に無操作になる条件: chunk_meta が0行 / その資料のチャンクが無い /
'   章キーが1つも取れない / ベクトルが1本も読めない / 既存資料が無い。
' ----------------------------------------------------------------------------
Public Sub BuildFor(ByVal sourceName As String)
    If LenB(Trim$(sourceName)) = 0 Then Exit Sub
    On Error GoTo Quiet

    Dim chaps() As String, chN() As Long, chCsv() As String
    Dim nCh As Long
    nCh = BuildCentroidsFor(sourceName, chaps, chN, chCsv)
    If nCh < 1 Then Exit Sub

    ' 【R37 Fix A-M2】いちばん落ちやすい処理(4,000章×768次元 Double ≒24.6MB の
    '   ReDim = 32bit で実行時エラー7)を【何かを消す前】に済ませる。旧版は
    '   Remove の後に展開していたため、エラー7で抜けると「相手から自分への
    '   リンクを消しただけ」の状態が残り、次にその相手を取り込み直すまで
    '   復活しなかった。読めてから消す。
    Dim exSrcs() As String, exChaps() As String
    Dim exFlat() As Double
    Dim exDim As Long, exTotal As Long
    Dim capN As Long: capN = modXDocStore.MaxChapters()
    Dim loadFailed As Boolean
    Dim nEx As Long
    nEx = LoadWithRetry(sourceName, exSrcs, exChaps, exFlat, exDim, capN, exTotal, loadFailed)

    ' 先に消すのは、同じ資料の行が2組並ぶと相手の上位3件が自分だけで
    ' 埋まってしまうため(doc_outline の作り直しと同じ判断)。
    ' ただし展開に失敗した回(loadFailed)は【相手側の行を消さない】。作り直す
    ' 材料が無いのに消せば、相手の📖から「関連する資料」が永久に消えるだけ
    ' になる。自分から出ている行(src_a=自分)だけ落として重心を書き、
    ' リンクは次に取り込み直したときに作り直す(R37 §10 A-M2)。
    modXDocStore.RemoveCentroidsFor sourceName
    If loadFailed Then
        modXDocStore.RemoveLinksForwardFor sourceName
    Else
        modXDocStore.RemoveLinksFor sourceName
    End If
    modXDocStore.WriteCentroidRows sourceName, chaps, chCsv, chN, nCh

    If nEx < 1 Then GoTo LogDone

    Dim fwd As String
    fwd = MatchChapters(chaps, chCsv, nCh, exSrcs, exChaps, exFlat, exDim, nEx)
    If LenB(fwd) = 0 Then GoTo LogDone

    ' 章数が上限を超えている本棚では相手側の作り直しをやらず新資料側だけ書く
    ' (R37 §7 リスク台帳)。相手側の行はそのまま残り、次にその資料を
    ' 取り込み直したときに正しくなる。
    ComposeLinks sourceName, fwd, (exTotal > capN)

LogDone:
    On Error Resume Next
    modLog.LogUsage "xdoc_built", "ingest", _
        "source=" & sourceName & " chapters=" & nCh & " peers=" & nEx & _
        " cap=" & capN & " loadfail=" & Abs(CLng(loadFailed))
    On Error GoTo 0
    Exit Sub

Quiet:
    ' ハンドラ稼働中は On Error Resume Next が効かないので必ず Resume で
    ' 抜けてから記録する(2026-07-30 実機err#462 と同型の作法)。
    ' 【R37 Fix B-m3】Err の値はここでしか読めないので控えてから渡す。
    Dim eN0 As Long: eN0 = Err.Number
    Dim eD0 As String: eD0 = Err.Description
    Resume BuildDone
BuildDone:
    Err.Clear
    modXDocStore.XDocFail "build", eN0, eD0
End Sub

' ----------------------------------------------------------------------------
' LoadWithRetry - LoadExistingCentroids を1回だけやり直す包み(R37 Fix A-M2)。
'   実行時エラー7(メモリ不足)だけは capN を半分にして1回再挑戦する
'   ―― 4,000章×768次元の Double 配列は約24.6MB で、32bit Excel では他の
'   アドインが載っているだけで確保できないことがある。半分(約12MB)なら
'   通ることが多く、上位3件の相手が少し減るだけで済む。
'   それ以外のエラーはやり直しても同じなので即あきらめる。
'   あきらめたときは outFailed=True(呼び出し元は「相手側の行を消さない」)。
'   capN は ByRef: 呼び出し元が「相手側の作り直しをやるか」を実際に使った
'   上限で判定する(半分にしたのに元の上限で比べると判定が狂う)。
' ----------------------------------------------------------------------------
Private Function LoadWithRetry(ByVal sourceName As String, ByRef outSrcs() As String, _
                               ByRef outChaps() As String, ByRef outFlat() As Double, _
                               ByRef outDim As Long, ByRef capN As Long, _
                               ByRef outTotal As Long, ByRef outFailed As Boolean) As Long
    Dim tries As Long
    Dim n As Long
    For tries = 1 To 2
        Err.Clear
        On Error Resume Next
        n = LoadExistingCentroids(sourceName, outSrcs, outChaps, outFlat, outDim, capN, outTotal)
        Dim eN As Long: eN = Err.Number
        Err.Clear
        On Error GoTo 0
        If eN = 0 Then
            LoadWithRetry = n
            Exit Function
        End If
        If eN <> 7 Then Exit For
        capN = capN \ 2
        If capN < 1 Then Exit For
    Next tries
    outFailed = True
End Function

' ----------------------------------------------------------------------------
' BuildCentroidsFor - その資料の章ごとの重心を作る。戻り値=章数。
'   章キーは modOutlineBuild.ChapterKeyOf を必ず通す。ここへ「同じような式」を
'   書き直すと、章要約の章と重心の章が静かに食い違って doc_links が毎回
'   1件も当たらない=誰にも見えない無音の失敗になる(憲章§4-5)。
' ----------------------------------------------------------------------------
Private Function BuildCentroidsFor(ByVal sourceName As String, ByRef outChaps() As String, _
                                   ByRef outN() As Long, ByRef outCsv() As String) As Long
    Dim mIds() As String, mPaths() As String, mRefs() As String
    Dim metaN As Long: metaN = modChunkMetaStore.ReadAllMeta(mIds, mPaths, mRefs)
    If metaN < 1 Then Exit Function

    Dim ids() As String
    Dim nSrc As Long: nSrc = CollectIds(sourceName, ids)
    If nSrc < 1 Then Exit Function

    ' chunk_id の箱(vbLf区切り)。meta を1周して位置を割り出す
    ' (modOutlineBuild.MapPaths と同じ流儀)。
    Dim box As String
    Dim i As Long
    For i = 1 To nSrc
        box = box & vbLf & ids(i)
    Next i
    box = box & vbLf

    Dim chapOf() As String: ReDim chapOf(1 To nSrc)
    For i = 0 To metaN - 1
        If LenB(mPaths(i)) > 0 Then
            Dim k As Long: k = IndexInBox(box, mIds(i))
            If k >= 1 And k <= nSrc Then chapOf(k) = modOutlineBuild.ChapterKeyOf(mPaths(i))
        End If
    Next i

    Dim vecOf() As String: ReDim vecOf(1 To nSrc)
    LoadVectorsFor box, nSrc, vecOf

    ' 章キーのユニーク化(出現順=文書順)。
    Dim uniq() As String: ReDim uniq(1 To nSrc)
    Dim nU As Long
    Dim j As Long
    For i = 1 To nSrc
        If LenB(chapOf(i)) > 0 Then
            Dim at As Long: at = 0
            For j = 1 To nU
                If StrComp(uniq(j), chapOf(i), vbTextCompare) = 0 Then
                    at = j
                    Exit For
                End If
            Next j
            If at = 0 Then
                nU = nU + 1
                uniq(nU) = chapOf(i)
            End If
        End If
    Next i
    If nU < 1 Then Exit Function

    ReDim outChaps(1 To nU)
    ReDim outN(1 To nU)
    ReDim outCsv(1 To nU)
    Dim bufCsv() As String: ReDim bufCsv(1 To nSrc)
    Dim outIdx As Long
    For j = 1 To nU
        Dim m As Long: m = 0
        For i = 1 To nSrc
            If StrComp(chapOf(i), uniq(j), vbTextCompare) = 0 Then
                If LenB(vecOf(i)) > 0 Then
                    m = m + 1
                    bufCsv(m) = vecOf(i)
                End If
            End If
        Next i
        If m > 0 Then
            Dim cs As String: cs = modXDocStore.MeanNormalizedCsv(bufCsv, m)
            If LenB(cs) > 0 Then
                outIdx = outIdx + 1
                outChaps(outIdx) = uniq(j)
                outN(outIdx) = m
                outCsv(outIdx) = cs
            End If
        End If
    Next j
    BuildCentroidsFor = outIdx
End Function

' その資料の chunk_id を文書順に集める(id列とsource列の2列だけ読む。本文まで
' 読むと2万行規模で数百MBになる。modOutlineBuild.CollectSourceRows と同型)。
Private Function CollectIds(ByVal sourceName As String, ByRef outIds() As String) As Long
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Dim lastK As Long: lastK = ws.Cells(ws.Rows.count, COL_K_ID).End(xlUp).row
    If lastK < 2 Then Exit Function

    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, COL_K_ID), ws.Cells(lastK, COL_K_SRC)).Value
    ReDim outIds(1 To lastK - 1)
    Dim n As Long
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(Trim$(CStr(arr(i, COL_K_SRC))), sourceName, vbTextCompare) = 0 Then
            Dim id As String: id = Trim$(CStr(arr(i, COL_K_ID)))
            If LenB(id) > 0 Then
                n = n + 1
                outIds(n) = id
            End If
        End If
    Next i
    CollectIds = n
End Function

' my_vectors から【この資料の行だけ】vector_csv を読む。chunk_id 列を1回だけ
' 一括で読み、当たった行の2列目をセル単位で読む2段読み(2万行×768次元のCSVを
' 丸ごと配列にすると32bit Excelはメモリ不足で無言で死ぬ。modTextView.
' ShowForSource / modCorrect.InjectHits と同じ作法)。
Private Sub LoadVectorsFor(ByVal box As String, ByVal nSrc As Long, ByRef outVec() As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_VECTORS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Dim lastV As Long: lastV = ws.Cells(ws.Rows.count, COL_V_ID).End(xlUp).row
    If lastV < 2 Then Exit Sub

    Dim idArr As Variant
    idArr = modVecCache.AsColumnArray( _
        ws.Range(ws.Cells(2, COL_V_ID), ws.Cells(lastV, COL_V_ID)).Value)

    Dim r As Long
    For r = LBound(idArr, 1) To UBound(idArr, 1)
        Dim vid As String: vid = Trim$(CStr(idArr(r, 1)))
        If LenB(vid) > 0 Then
            Dim k As Long: k = IndexInBox(box, vid)
            If k >= 1 And k <= nSrc Then outVec(k) = CStr(ws.Cells(r + 1, COL_V_CSV).Value)
        End If
    Next r
End Sub

' vbLf区切りの箱の中で id が何番目か(1始まり。無ければ0)。位置は「一致した
' 所より前にある vbLf の数」で決まる(箱は必ず vbLf で始まる)。
Private Function IndexInBox(ByVal box As String, ByVal id As String) As Long
    If LenB(id) = 0 Then Exit Function
    Dim p As Long
    p = InStr(1, box, vbLf & id & vbLf, vbBinaryCompare)
    If p = 0 Then Exit Function
    Dim head As String: head = Left$(box, p)
    IndexInBox = Len(head) - Len(Replace(head, vbLf, ""))
End Function

' ----------------------------------------------------------------------------
' LoadExistingCentroids - 自分以外の章重心を平坦なDouble配列へ1回だけ展開する。
'   次元は「最初に読めた行」で確定(modVecCache.BuildFrom と同じ規約)。
'   cap 件で打ち切り、打ち切る前の総数を outTotal へ返す(呼び出し元が
'   「相手側の作り直しをやるか」を決める材料。R37 §7)。
'   ReDim は 4,000×768 の Double = 約24.6MB で32bitでは実行時エラー7に
'   なり得る。BuildFor の On Error GoTo Quiet が受けて「リンク無し」で静かに
'   終わる(取込は止めない)。
' ----------------------------------------------------------------------------
Private Function LoadExistingCentroids(ByVal sourceName As String, ByRef outSrcs() As String, _
                                       ByRef outChaps() As String, ByRef outFlat() As Double, _
                                       ByRef outDim As Long, ByVal cap As Long, _
                                       ByRef outTotal As Long) As Long
    Dim cs() As String, cc() As String, cv() As String
    Dim n As Long: n = modXDocStore.ReadCentroids(cs, cc, cv)
    outTotal = 0
    If n < 1 Then Exit Function

    Dim dimN As Long
    Dim probe() As Double
    Dim i As Long
    For i = 0 To n - 1
        If StrComp(Trim$(cs(i)), sourceName, vbTextCompare) <> 0 Then
            outTotal = outTotal + 1
            If dimN < 1 Then
                If modUtil.CsvToVector(cv(i), probe) Then
                    dimN = UBound(probe) - LBound(probe) + 1
                End If
            End If
        End If
    Next i
    If dimN < 1 Then Exit Function
    If outTotal < 1 Then Exit Function

    Dim room As Long: room = cap
    If room < 1 Then room = 1
    If room > outTotal Then room = outTotal

    ReDim outSrcs(1 To room)
    ReDim outChaps(1 To room)
    ReDim outFlat(0 To room * dimN - 1)

    Dim k As Long
    Dim v() As Double
    Dim j As Long
    For i = 0 To n - 1
        If k >= room Then Exit For
        If StrComp(Trim$(cs(i)), sourceName, vbTextCompare) <> 0 Then
            If modUtil.CsvToVector(cv(i), v) Then
                If (UBound(v) - LBound(v) + 1) = dimN Then
                    Dim vb0 As Long: vb0 = LBound(v)
                    Dim baseI As Long: baseI = k * dimN
                    For j = 0 To dimN - 1
                        outFlat(baseI + j) = v(vb0 + j)
                    Next j
                    k = k + 1
                    outSrcs(k) = Trim$(cs(i))
                    outChaps(k) = cc(i)
                End If
            End If
        End If
    Next i
    outDim = dimN
    LoadExistingCentroids = k
End Function

' ----------------------------------------------------------------------------
' MatchChapters - 新資料の各章 × 既存の全章の内積を取り、章ごとに上位3件。
'   戻り値は "chap_a|src_b|chap_b|sim" の行を vbLf で連結したもの。
'   重心はどちらも L2 正規化済みなので、内積がそのままコサイン
'   (=modXDoc.CosineCsv がその場で正規化して出す値)と一致する。ここで
'   CosineCsv を章ごとに呼ばないのは、CSVの再パースが支配項になるため
'   (20章×4,000章で数百万回の Split/Val = 取込が分単位で止まる)。
'   等価であることは modTestsPure42 の C1 が固定する。
' ----------------------------------------------------------------------------
Private Function MatchChapters(ByRef chaps() As String, ByRef chCsv() As String, _
                               ByVal nCh As Long, ByRef exSrcs() As String, _
                               ByRef exChaps() As String, ByRef exFlat() As Double, _
                               ByVal exDim As Long, ByVal nEx As Long) As String
    Dim pairs() As String: ReDim pairs(1 To nEx)
    Dim sims() As Double: ReDim sims(1 To nEx)
    Dim acc() As String: ReDim acc(1 To nCh)
    Dim got As Long

    ' 【R37 Fix A-M3=B-M2】上位3件でも「近くない」ものは doc_links へ書かない。
    '   閾値を読むのは1回だけ(章ごとに config を引くと総当たりが遅くなる)。
    Dim minS As Long: minS = modXDocStore.SimFloor()

    Dim nv() As Double
    Dim i As Long, j As Long, d As Long
    For i = 1 To nCh
        If modUtil.CsvToVector(chCsv(i), nv) Then
            If (UBound(nv) - LBound(nv) + 1) = exDim Then
                Dim nb As Long: nb = LBound(nv)
                For j = 1 To nEx
                    Dim s As Double: s = 0
                    Dim baseI As Long: baseI = (j - 1) * exDim
                    For d = 0 To exDim - 1
                        s = s + nv(nb + d) * exFlat(baseI + d)
                    Next d
                    pairs(j) = exSrcs(j) & "|" & exChaps(j)
                    sims(j) = s
                Next j
                Dim topLn As String
                topLn = modXDocStore.TopNLinks(pairs, sims, nEx, modXDocStore.KEEP_N)
                topLn = modXDocStore.FilterBySim(topLn, minS)
                If LenB(topLn) > 0 Then
                    got = got + 1
                    acc(got) = PrefixLines(topLn, chaps(i) & "|")
                End If
            End If
        End If
    Next i
    If got < 1 Then Exit Function
    ReDim Preserve acc(1 To got)
    MatchChapters = Join(acc, vbLf)
End Function

' vbLf区切りの各行の先頭へ同じ文字列を足す。
Private Function PrefixLines(ByVal src As String, ByVal head As String) As String
    Dim arr() As String: arr = Split(src, vbLf)
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        arr(i) = head & arr(i)
    Next i
    PrefixLines = Join(arr, vbLf)
End Function

' ----------------------------------------------------------------------------
' ComposeLinks - doc_links の全行を組み立てて書き直す。
'   fwd は "chap_a|src_b|chap_b|sim" の行(chap_a は sourceName の章)。
'   ① 新資料側(src_a=sourceName)の行をそのまま採る。
'   ② 相手側(src_b,chap_b)から見た上位3件も作り直す。既存の相手行と今回の
'      逆向き(相手→新資料)を混ぜて TopNLinks に掛ける = 「片側だけ新しい」
'      状態を作らない。
'   ③ onlyForward=True(章数が上限超え)なら②はやらない。
'   ※この資料が絡む行は呼び出し元が RemoveLinksFor 済みなので、ここで読む
'     既存行に src_a/src_b=sourceName は入っていない。
' ----------------------------------------------------------------------------
Private Sub ComposeLinks(ByVal sourceName As String, ByVal fwd As String, _
                         ByVal onlyForward As Boolean)
    Dim fLines() As String: fLines = Split(fwd, vbLf)
    Dim nF As Long: nF = UBound(fLines) - LBound(fLines) + 1
    If nF < 1 Then Exit Sub

    Dim la() As String, lca() As String, lb() As String, lcb() As String
    Dim lsim() As Double
    Dim nL As Long: nL = modXDocStore.ReadLinks(la, lca, lb, lcb, lsim)

    ' 相手側の作り直し対象キー("src|chap")の箱。
    Dim affBox As String
    Dim i As Long
    If Not onlyForward Then
        For i = LBound(fLines) To UBound(fLines)
            Dim f0() As String: f0 = Split(fLines(i), "|")
            If UBound(f0) >= 3 Then
                Dim kb As String: kb = f0(1) & "|" & f0(2)
                If InStr(1, affBox, vbLf & kb & vbLf, vbTextCompare) = 0 Then
                    affBox = affBox & vbLf & kb & vbLf
                End If
            End If
        Next i
    End If

    ' 出力の器: 既存の残り nL + 新資料側 nF + 相手側の作り直し nF×KEEP_N。
    ' ここを小さく見積もると添字の範囲外で【リンクだけが静かに消える】ので、
    ' 最悪(相手が全部別キー・全員が上位3件を作り直す)で数える。
    Dim capRows As Long: capRows = nL + nF * (modXDocStore.KEEP_N + 1) + 4
    Dim oA() As String: ReDim oA(1 To capRows)
    Dim oCA() As String: ReDim oCA(1 To capRows)
    Dim oB() As String: ReDim oB(1 To capRows)
    Dim oCB() As String: ReDim oCB(1 To capRows)
    Dim oS() As Double: ReDim oS(1 To capRows)
    Dim nOut As Long

    ' 相手側の作り直しに使う「保留」行。
    Dim pKey() As String: ReDim pKey(1 To nL + 1)
    Dim pPair() As String: ReDim pPair(1 To nL + 1)
    Dim pSim() As Double: ReDim pSim(1 To nL + 1)
    Dim nP As Long

    For i = 0 To nL - 1
        Dim ka As String: ka = Trim$(la(i)) & "|" & Trim$(lca(i))
        Dim held As Boolean: held = False
        If LenB(affBox) > 0 Then
            held = (InStr(1, affBox, vbLf & ka & vbLf, vbTextCompare) > 0)
        End If
        If held Then
            nP = nP + 1
            pKey(nP) = ka
            pPair(nP) = Trim$(lb(i)) & "|" & Trim$(lcb(i))
            pSim(nP) = lsim(i)
        Else
            nOut = nOut + 1
            oA(nOut) = la(i): oCA(nOut) = lca(i)
            oB(nOut) = lb(i): oCB(nOut) = lcb(i)
            oS(nOut) = lsim(i)
        End If
    Next i

    ' ① 新資料側。
    For i = LBound(fLines) To UBound(fLines)
        Dim f1() As String: f1 = Split(fLines(i), "|")
        If UBound(f1) >= 3 Then
            nOut = nOut + 1
            oA(nOut) = sourceName: oCA(nOut) = f1(0)
            oB(nOut) = f1(1): oCB(nOut) = f1(2)
            oS(nOut) = CDbl(Val(f1(3)))
        End If
    Next i

    ' ② 相手側。
    If Not onlyForward Then
        Dim doneBox As String
        For i = LBound(fLines) To UBound(fLines)
            Dim f2() As String: f2 = Split(fLines(i), "|")
            If UBound(f2) >= 3 Then
                Dim key2 As String: key2 = f2(1) & "|" & f2(2)
                If InStr(1, doneBox, vbLf & key2 & vbLf, vbTextCompare) = 0 Then
                    doneBox = doneBox & vbLf & key2 & vbLf
                    nOut = RebuildPeer(key2, sourceName, fLines, pKey, pPair, pSim, nP, _
                                       oA, oCA, oB, oCB, oS, nOut)
                End If
            End If
        Next i
    End If

    modXDocStore.WriteAllLinks oA, oCA, oB, oCB, oS, nOut
End Sub

' 相手側(key2 = "src|chap")の上位3件を作り直して出力へ足す。材料は
' 「保留にした既存行」+「今回の逆向き(相手→新資料)」。戻り値=新しい nOut。
Private Function RebuildPeer(ByVal key2 As String, ByVal sourceName As String, _
                             ByRef fLines() As String, ByRef pKey() As String, _
                             ByRef pPair() As String, ByRef pSim() As Double, ByVal nP As Long, _
                             ByRef oA() As String, ByRef oCA() As String, _
                             ByRef oB() As String, ByRef oCB() As String, _
                             ByRef oS() As Double, ByVal nOut As Long) As Long
    RebuildPeer = nOut

    Dim capN As Long: capN = nP + (UBound(fLines) - LBound(fLines) + 1) + 1
    Dim pairs() As String: ReDim pairs(1 To capN)
    Dim sims() As Double: ReDim sims(1 To capN)
    Dim m As Long
    Dim i As Long

    For i = 1 To nP
        If StrComp(pKey(i), key2, vbTextCompare) = 0 Then
            m = m + 1
            pairs(m) = pPair(i)
            sims(m) = pSim(i)
        End If
    Next i
    For i = LBound(fLines) To UBound(fLines)
        Dim f3() As String: f3 = Split(fLines(i), "|")
        If UBound(f3) >= 3 Then
            If StrComp(f3(1) & "|" & f3(2), key2, vbTextCompare) = 0 Then
                m = m + 1
                pairs(m) = sourceName & "|" & f3(0)
                sims(m) = CDbl(Val(f3(3)))
            End If
        End If
    Next i
    If m < 1 Then Exit Function

    Dim topLn As String
    topLn = modXDocStore.TopNLinks(pairs, sims, m, modXDocStore.KEEP_N)
    ' 相手側の行も同じ閾値で切る(R37 Fix A-M3。書く側は1つの規則で揃える)。
    topLn = modXDocStore.FilterBySim(topLn, modXDocStore.SimFloor())
    If LenB(topLn) = 0 Then Exit Function

    Dim kf() As String: kf = Split(key2, "|")
    If UBound(kf) < 1 Then Exit Function

    Dim tl() As String: tl = Split(topLn, vbLf)
    Dim outN As Long: outN = nOut
    For i = LBound(tl) To UBound(tl)
        Dim tf() As String: tf = Split(tl(i), "|")
        If UBound(tf) >= 2 Then
            outN = outN + 1
            oA(outN) = kf(0): oCA(outN) = kf(1)
            oB(outN) = tf(0): oCB(outN) = tf(1)
            oS(outN) = CDbl(Val(tf(2)))
        End If
    Next i
    RebuildPeer = outN
End Function
