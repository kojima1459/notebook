Attribute VB_Name = "modXDoc"
Option Explicit

' ============================================================================
' modXDoc - 資料間リンクの【回答側】(R37 §3)。
' ----------------------------------------------------------------------------
' やること(これだけ):
'   1) 検索の候補プール(topKで切られる前の40件)を控える(RememberPool)。
'   2) 最終hitsが確定した直後、上位2件が属する章と強く繋がる【他資料の章】を
'      doc_links から引き、その章に属する候補が【プールの中に居るなら】
'      score 0 で末尾へ最大2件だけ足す(Expand)。
'
' やらないこと(意図的に):
'   ・新しい検索は走らせない。プールに居ない資料は絶対に足さない。
'     「別の資料も見に行く」ではなく「もう取ってきた候補の中から、章の
'     繋がりで拾い直す」だけ。取込・回答のどちらのLLM呼び出しも増えない。
'   ・順位は動かさない。足すのは末尾で score=0(modAskFocus.NeighborExpand /
'     RefsExpand と同じ扱い。信頼度バッジ・分散判定を汚さない)。
'
' なぜ「章の重心の近さ」という条件を付けるのか:
'   modAskFocus.RefsExpand が資料を跨がない理由(bas:204-207)と同じ危険が
'   ここにはある ―― 「まったく無関係な規程の第8条が根拠として混ざる。
'   出典タグは正しいので利用者は気付けない」。だから跨いでよいのは
'   「章の重心どうしが config xdoc_min_sim(既定0.80)以上に近い」ときだけ、
'   末尾に、最大 xdoc_add(既定2)件に絞る。config xdoc_links=off で丸ごと止まる。
'
' 効果の約束はしない(R37 §3-1):
'   単一資料で完結する質問には何も変わらない。効くのは「A の言葉で聞いたが
'   答えは B にある」型だけで、そのときも出るとは限らない。
'
' 【LO の死角】RememberPool/Expand は Hit 型を跨ぐので LibreOffice の実行
'   テストからは呼べない(run_lo_tests.py の既知の制約)。Pure テストで
'   固定するのは純関数(CosineCsv/ChapterOf/PickLinked)だけで、配線は
'   実機受入(R37 §3-2)に委ねる。
' ============================================================================

' 起点にする hits の件数。1件だと「たまたま1位が外れた」回に引きずられ、
' 3件以上にすると候補の章が広がって閾値だけでは絞りきれない。
Private Const SEED_N As Long = 2

' 控えたプール(modAskRetrieve.StashDispersionPool の直後に渡される)。
' Expand は【1回使ったら必ず捨てる】: 検索が単段へ退化した回
' (FallbackSingle)に前回の質問のプールを掴んだままだと、まったく別の
' 質問の候補が末尾へ紛れ込む。
Private mPoolHits() As Hit
Private mPoolN As Long

' ----------------------------------------------------------------------------
' RememberPool - 候補プール(rerank・topK絞り込みより前)を控える。
'   呼び出しは modAskRetrieve.RunMultiRetrieve の StashDispersionPool の直後
'   1行だけ(あちらは残り147字なので、判断はすべてこちら側に置く)。
'   【R37 Fix B-M1(2)】続けて質問・深掘り・逆質問スコープの回は控えない。
'     それらのターンは「前の回答の続き」を出す約束の上に立っていて、件数を
'     増やす側の細工(deep_scope_fallback の n>=2 ゲート・DemoteUsed・
'     ClarifyScopeKept・件数バッジ)がその件数を見ている(R34 F1 が禁じた型)。
'     控えなければ mPoolN=0 のまま = Expand は何もしない no-op になる。
'     スコープ付き検索そのものを外す判定は呼び出し側(modAskRetrieve の
'     If scopeSources Is Nothing)。ここは「深掘りターンか」だけを見る。
'   【R37 Fix A-m2】配列コピーが落ちた回に件数だけ更新すると、前の質問の
'     プールを掴んだまま「n件ある」と言い張る状態になる。成功したときだけ。
'   【R37 Fix2 C-m1】mPoolN=0 のリセットは scopeSources の有無に関係なく
'     必ず先に通す。旧版は呼び出し側の If scopeSources Is Nothing で
'     RememberPool 自体を呼ばないようにしていたため、スコープ付き検索が
'     FallbackSingle(単段検索)に落ちた回だけ、前のスコープ無しターンの
'     プールが mPoolN>0 のまま生き残り、次のスコープ付きターンで Expand に
'     使われてしまう(見えない誤根拠)。scopeSources 判定はここへ引き取る。
' ----------------------------------------------------------------------------
Public Sub RememberPool(ByRef hits() As Hit, ByVal n As Long, _
                         Optional ByVal scopeSources As Object)
    On Error Resume Next
    mPoolN = 0
    If Not scopeSources Is Nothing Then Exit Sub
    If n < 1 Then Exit Sub
    If Not modXDocStore.GateOn() Then Exit Sub
    If modFollowup.IsFollowupTurn() Then Exit Sub
    Err.Clear
    mPoolHits = hits
    If Err.Number = 0 Then mPoolN = n
    Err.Clear
End Sub

' ----------------------------------------------------------------------------
' Expand - 最終hitsの末尾へ、章の繋がりで拾い直した候補を足す。戻り値=新しい件数。
'   呼び出しは modAskRetrieve.RunMultiRetrieve の「最終hits確定点」
'   (DiversifyFinalHits の直後・ArticleEnsure より前)1行だけ。
'   ArticleEnsure は【先頭】へ入れるので、順序の取り合いは起きない。
'   何もできないときは n をそのまま返す(呼び出し元は分岐を1つも持たない)。
' ----------------------------------------------------------------------------
Public Function Expand(ByRef hits() As Hit, ByVal n As Long, ByVal topK As Long) As Long
    Expand = n
    If n < 1 Then GoTo Done
    If mPoolN < 1 Then GoTo Done
    If Not modXDocStore.GateOn() Then GoTo Done

    On Error GoTo Quiet

    Dim maxAdd As Long: maxAdd = modConfig.GetLong("xdoc_add", 2)
    If maxAdd < 1 Then GoTo Done

    ' 【R37 Fix A-m4=B-m6】リンクが1行も無いなら、ここで終わる。旧版は先に
    '   chunk_meta を丸ごと配列へコピーしていたので、doc_links が空の本棚
    '   (版上げ直後・パック取込だけの利用者・xdoc を使っていない人)でも
    '   毎問その代金を払っていた。行数を見るのはセル1つぶん。
    If modXDocStore.LinkRowCount() < 1 Then GoTo Done

    ' 1) 起点は「上位 SEED_N 件」。topK より後ろ(近傍・参照展開で足された
    '    score 0 のチャンク)は起点にしない。
    Dim seedN As Long: seedN = n
    If topK > 0 And seedN > topK Then seedN = topK
    If seedN > SEED_N Then seedN = SEED_N
    If seedN < 1 Then GoTo Done

    Dim mIds() As String, mPaths() As String, mRefs() As String
    Dim metaN As Long: metaN = modChunkMetaStore.ReadAllMeta(mIds, mPaths, mRefs)
    If metaN < 1 Then GoTo Done

    Dim seedKeys() As String: ReDim seedKeys(1 To seedN)
    Dim keyBox As String
    Dim kN As Long
    Dim i As Long
    For i = 1 To seedN
        Dim ch As String: ch = ChapterOfChunk(hits(i).chunk_id, mIds, mPaths, metaN)
        If LenB(ch) > 0 Then
            kN = kN + 1
            seedKeys(kN) = Trim$(hits(i).source) & "|" & ch
            keyBox = keyBox & vbLf & seedKeys(kN) & vbLf
        End If
    Next i
    If kN < 1 Then GoTo Done
    ReDim Preserve seedKeys(1 To kN)

    ' 2) その章と繋がっている他資料の章を doc_links から引く(閾値も件数も
    '    見ない生の行。判断は純関数 PickLinked に閉じる)。
    Dim linksCsv As String: linksCsv = modXDocStore.LinksCsvFor(keyBox)
    If LenB(linksCsv) = 0 Then GoTo Done

    Dim cand As String
    cand = PickLinked(seedKeys, linksCsv, modXDocStore.SimFloor(), maxAdd)
    If LenB(cand) = 0 Then GoTo Done

    ' 3) プールの中から、その章に属していてまだ hits に無いものを末尾へ。
    Dim outN As Long: outN = n
    Dim added As Long
    added = AppendFromPool(cand, hits, outN, maxAdd, mIds, mPaths, metaN)
    If added > 0 Then
        Expand = outN
        On Error Resume Next
        modLog.LogUsage "xdoc_expand", "", _
            "added=" & added & " via=" & Replace(cand, vbLf, " / "), 0, outN
    End If

Done:
    mPoolN = 0
    Exit Function

Quiet:
    ' ハンドラ稼働中は On Error Resume Next が効かないので Resume で抜ける
    ' (2026-07-30 実機err#462 と同型の作法)。足せなくても質問には必ず答える。
    Resume QuietDone
QuietDone:
    Err.Clear
    mPoolN = 0
End Function

' ============================================================================
' 純ロジック(LibreOffice の実行テストで固定する = modTestsPure42)
' ============================================================================

' ----------------------------------------------------------------------------
' CosineCsv - ベクトルCSV 2本のコサイン類似度。
'   ・次元が違う/どちらかが読めない/ゼロベクトル は 0(「分からない」ではなく
'     「近くない」を返す。呼び出し元に分岐を作らせないため)。
'   ・doc_centroids へ書く重心は既に L2 正規化済みなので、保存済みの2本に
'     ついてはこの関数の値と単純な内積が一致する。取込の総当たり
'     (modXDocStore.MatchChapters)はその性質を使って内積だけを回す
'     ―― 章ごとに CSV を再パースすると 20章×4,000章 で取込が分単位で
'     止まるため。等価であることは modTestsPure42 が固定する。
' ----------------------------------------------------------------------------
Public Function CosineCsv(ByVal a As String, ByVal b As String) As Double
    Dim v1() As Double, v2() As Double
    If Not modUtil.CsvToVector(a, v1) Then Exit Function
    If Not modUtil.CsvToVector(b, v2) Then Exit Function

    Dim n1 As Long: n1 = UBound(v1) - LBound(v1) + 1
    Dim n2 As Long: n2 = UBound(v2) - LBound(v2) + 1
    If n1 <> n2 Then Exit Function
    If n1 < 1 Then Exit Function

    Dim la As Double: la = Sqr(modUtil.DotProduct(v1, v1))
    If la <= 0 Then Exit Function
    Dim lb As Double: lb = Sqr(modUtil.DotProduct(v2, v2))
    If lb <= 0 Then Exit Function

    CosineCsv = modUtil.DotProduct(v1, v2) / (la * lb)
End Function

' ----------------------------------------------------------------------------
' ChapterOf - section_path から章キーを取る。
'   式は modOutlineBuild.ChapterKeyOf ただ1本(doc_outline・doc_centroids・
'   doc_links・回答時の照合が全部これを通る)。ここで「同じような式」を
'   書き直すと、章要約の章とリンクの章が静かに食い違い、doc_links が
'   毎回1件も当たらない=誰にも見えない無音の失敗になる(憲章§4-5)。
' ----------------------------------------------------------------------------
Public Function ChapterOf(ByVal sectionPath As String) As String
    ChapterOf = modOutlineBuild.ChapterKeyOf(sectionPath)
End Function

' ----------------------------------------------------------------------------
' PickLinked - 「上位hitsの (資料|章)」と doc_links の行から、末尾へ足す
'   候補の (資料|章) を最大 maxN 件返す(vbLf区切り)。
'   linksCsv の1行 = "src_a|chap_a|src_b|chap_b|sim"(sim は 0〜1 の実数表記)。
'   採る条件は3つ全部:
'     ① src_a|chap_a が起点キーのどれかと一致する
'     ② sim×100 が minSim 以上(既定80=0.80)
'     ③ src_b が起点hitsの資料【ではない】(自資料の章は既に検索が見ている。
'        同じ資料の別の章を足しても「A で聞いて B に届く」にはならない)
'   同じ (資料|章) は1回だけ。行の並び順で先着を採る(doc_links は
'   modXDocStore.TopNLinks が sim 降順で書いているので、先着=いちばん近い)。
' ----------------------------------------------------------------------------
Public Function PickLinked(ByRef hitsSrcChap() As String, ByVal linksCsv As String, _
                           ByVal minSim As Long, ByVal maxN As Long) As String
    If maxN < 1 Then Exit Function
    If LenB(linksCsv) = 0 Then Exit Function

    Dim lo As Long, hi As Long
    On Error Resume Next
    lo = LBound(hitsSrcChap): hi = UBound(hitsSrcChap)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    If hi < lo Then Exit Function

    ' 起点キーの箱と、起点hitsの資料名の箱。
    Dim keyBox As String, srcBox As String
    Dim i As Long
    For i = lo To hi
        Dim k As String: k = Trim$(hitsSrcChap(i))
        If LenB(k) > 0 Then
            keyBox = keyBox & vbLf & k & vbLf
            Dim sp As Long: sp = InStr(1, k, "|", vbBinaryCompare)
            If sp > 1 Then
                Dim sName As String: sName = Trim$(Left$(k, sp - 1))
                If InStr(1, srcBox, vbLf & sName & vbLf, vbTextCompare) = 0 Then
                    srcBox = srcBox & vbLf & sName & vbLf
                End If
            End If
        End If
    Next i
    If LenB(keyBox) = 0 Then Exit Function

    Dim picks() As String: ReDim picks(1 To maxN)
    Dim got As Long
    Dim seen As String
    Dim lines0() As String: lines0 = Split(linksCsv, vbLf)
    For i = LBound(lines0) To UBound(lines0)
        If got >= maxN Then Exit For
        Dim ln As String: ln = Trim$(lines0(i))
        If LenB(ln) > 0 Then
            Dim f() As String: f = Split(ln, "|")
            If UBound(f) >= 4 Then
                Dim ka As String: ka = Trim$(f(0)) & "|" & Trim$(f(1))
                Dim sb As String: sb = Trim$(f(2))
                Dim kb As String: kb = sb & "|" & Trim$(f(3))
                If InStr(1, keyBox, vbLf & ka & vbLf, vbTextCompare) > 0 Then
                    If SimPasses(f(4), minSim) Then
                        If InStr(1, srcBox, vbLf & sb & vbLf, vbTextCompare) = 0 Then
                            If InStr(1, seen, vbLf & kb & vbLf, vbTextCompare) = 0 Then
                                seen = seen & vbLf & kb & vbLf
                                got = got + 1
                                picks(got) = kb
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
    If got < 1 Then Exit Function
    ReDim Preserve picks(1 To got)
    PickLinked = Join(picks, vbLf)
End Function

' sim(0〜1の実数表記)×100 が minSim 以上か。
' 【R37 Fix2 C-m4】式の実体は modXDocStore.SimOk の1本だけに置く(丸め落ち
' 対策の 1e-7 もそちら)。ここは Val() で文字列→実数に変換して渡すだけの薄皮。
Private Function SimPasses(ByVal simText As String, ByVal minSim As Long) As Boolean
    SimPasses = modXDocStore.SimOk(Val(simText), minSim)
End Function

' ============================================================================
' 内部ヘルパー(Excel/Hit 型に触れる部分)
' ============================================================================

' chunk_id から章キーを引く(chunk_meta を1周する)。無ければ空文字。
Private Function ChapterOfChunk(ByVal cid As String, ByRef mIds() As String, _
                                ByRef mPaths() As String, ByVal metaN As Long) As String
    Dim t As String: t = Trim$(cid)
    If LenB(t) = 0 Then Exit Function
    Dim i As Long
    For i = 0 To metaN - 1
        If StrComp(Trim$(mIds(i)), t, vbBinaryCompare) = 0 Then
            ChapterOfChunk = ChapterOf(mPaths(i))
            Exit Function
        End If
    Next i
End Function

' 候補の章に属するプール内のチャンクを hits の末尾へ足す。戻り値=足した件数。
' ・既に hits に居る chunk_id は足さない(二重の出典になる)。
' ・score は 0(modAskFocus.NeighborExpand/RefsExpand と同じ。低関連度の
'   警告・資料分散の判定はスコアを見るので、0 のまま末尾に置くのが
'   「材料は増やすが計器は汚さない」唯一の形)。
' 【R37 Fix A-m1】外側のループは【候補(cands)の順】。PickLinked は sim の
'   高い章から並べて返しているのに、旧版はプールの順に走って「候補のどれかに
'   入っていればいい」という集合判定をしていた。最大2件しか足さないので、
'   これは「いちばん近い章」ではなく「たまたまプールで上に居た章」を採ることに
'   なる。順位を作った側の意図をここで捨てない。
' プール各件の "資料|章" は先に1回だけ引く(章の解決は chunk_meta の総なめ
'   なので、候補×プールの二重ループの内側に置くと桁が変わる)。
Private Function AppendFromPool(ByVal candBox As String, ByRef hits() As Hit, _
                                ByRef nHits As Long, ByVal cap As Long, _
                                ByRef mIds() As String, ByRef mPaths() As String, _
                                ByVal metaN As Long) As Long
    If mPoolN < 1 Then Exit Function
    Dim cands() As String: cands = Split(candBox, vbLf)

    Dim have As String
    Dim i As Long
    For i = 1 To nHits
        have = have & vbLf & Trim$(hits(i).chunk_id)
    Next i
    have = have & vbLf

    Dim poolKey() As String: ReDim poolKey(1 To mPoolN)
    For i = 1 To mPoolN
        Dim cid0 As String: cid0 = Trim$(mPoolHits(i).chunk_id)
        If LenB(cid0) > 0 Then
            If InStr(1, have, vbLf & cid0 & vbLf, vbBinaryCompare) = 0 Then
                Dim ch0 As String: ch0 = ChapterOfChunk(cid0, mIds, mPaths, metaN)
                If LenB(ch0) > 0 Then
                    poolKey(i) = Trim$(mPoolHits(i).source) & "|" & ch0
                End If
            End If
        End If
    Next i

    Dim added As Long
    Dim ci As Long
    For ci = LBound(cands) To UBound(cands)
        If added >= cap Then Exit For
        Dim want As String: want = Trim$(cands(ci))
        If LenB(want) > 0 Then
            For i = 1 To mPoolN
                If added >= cap Then Exit For
                If LenB(poolKey(i)) > 0 Then
                    If StrComp(poolKey(i), want, vbTextCompare) = 0 Then
                        Dim cid As String: cid = Trim$(mPoolHits(i).chunk_id)
                        If InStr(1, have, vbLf & cid & vbLf, vbBinaryCompare) = 0 Then
                            nHits = nHits + 1
                            ReDim Preserve hits(1 To nHits)
                            hits(nHits) = mPoolHits(i)
                            hits(nHits).score = 0
                            have = have & cid & vbLf
                            added = added + 1
                        End If
                    End If
                End If
            Next i
        End If
    Next ci
    AppendFromPool = added
End Function
