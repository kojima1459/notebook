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
'   ・資料の同一性は【source 列(=Hit.source)】で見る。文書順は chunk_id
'     "bs::ハッシュ::pN::cM" の (page, seq) 昇順で復元する。seq はページごとに
'     振り直されるので、page を先に見ないと順序が壊れる。文書順を返す公開APIは
'     本棚側に無く、列の追加は既存本棚の移行が要るため、ここで chunk_id を
'     自前パースする(仕様 §3 前提事実)。
'     2026-08-05(R16H FA-1 / A-H1): ハッシュ部を資料キーにしていたのを source へ
'     直した。あのハッシュは【チャンク本文の】ハッシュで行ごとに違う
'     (modShelf: "bs::" & Fnv1a64Hex(本文) & "::pN::cM")。資料キーに使うと
'     どの行も「自分ひとりの資料」になり、隣が1件も見つからない=精読が実データで
'     恒久的に0件になっていた。テストだけが同じハッシュを共有する形で書かれて
'     いたため、真理表は通るのに実機では一度も効かない状態だった。
'   ・my_knowledge の読み方は modRetrieve と同じ直接 Range 読み。ただし読むのは
'     【chunk_id 列と source 列の2列だけ】で、本文(full_text=最大32,000字)は
'     採用した数十行だけセル単位で読む。全列を読み直すと、1回の質問で本棚全体の
'     本文を2度メモリへ載せることになる(20,000件規模で数百MB)。
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

' 参照展開(R17 Phase1)で足す既定の件数と、chunk_meta 側から候補として拾う
' 上限。候補箱は my_knowledge の全行に対して InStr を掛ける材料になるので、
' 大きくすると1質問あたりの走査量がそのまま伸びる(20,000行×候補箱の長さ)。
Private Const REFS_ADD_DEFAULT As Long = 8
Private Const REFS_MAX_CAND As Long = 120

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
    If Not BuildWindows(hits, nHits, radius, hitIds, needKeys, winLine) Then
        LogZero radius, "no_key"
        Exit Sub
    End If

    ' 2) chunk_id 列と source 列を1回で読み、窓に入る行だけを候補に残す。
    Dim idData As Variant
    idData = ws.Range(ws.Cells(2, COL_ID), ws.Cells(lastK, COL_SOURCE)).Value

    Dim candIds As String, candSrcs As String
    Dim candRow() As Long
    Dim candN As Long
    candN = CollectCandidates(idData, needKeys, winLine, candIds, candSrcs, candRow)
    If candN < 1 Then
        LogZero radius, "no_cand"
        Exit Sub
    End If

    ' 3) 文書順に並べ、各ヒットの前後 radius を拾う(純ロジック)。
    Dim addLine As String
    addLine = NeighborIdList(candIds, candSrcs, hitIds, radius)
    If LenB(addLine) = 0 Then
        LogZero radius, "no_neighbor"
        Exit Sub
    End If

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

    If added < 1 Then
        LogZero radius, "no_row"
        Exit Sub
    End If

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

' 近傍を1件も足せなかったことの記録(2026-08-05 R16H FA-1)。
' 精読は失敗しても回答が返るため、効いていない状態が【無音】で続く。実際、
' 資料キーの取り違えで実データでは恒久0件だったのに、ログにも画面にも
' 何ひとつ出ていなかった(A-H1)。理由の語を1つ添えて必ず1行残す。
'   no_key      = ヒットの chunk_id が1つも読めない(形式が違う)
'   no_cand     = 窓に入る行が本棚に無い(その資料が1チャンクだけ 等)
'   no_neighbor = 候補はあるが前後が全てヒット自身だった
'   no_row      = 拾った chunk_id に対応する行を引けなかった
Private Sub LogZero(ByVal radius As Long, ByVal reason As String)
    On Error Resume Next
    modLog.LogUsage "neighbor_zero", "focus", "radius=" & radius & " why=" & reason
    On Error GoTo 0
End Sub

' ============================================================================
' R17 Phase1: 構造グラフ(chunk_meta)への合流
' ----------------------------------------------------------------------------
' NeighborExpand が「物理的に隣のチャンク」を足すのに対し、ここは
' 「意味の上で繋がっているチャンク」を足す。規程は第6条の中に答えが無く
' 「第8条の定めによる」「別表2のとおり」と書いてあることが普通で、
' 物理近傍だけでは何ページ離れた参照先に永久に届かない(R17設計書§1)。
'
' 共通の設計判断:
'   ・chunk_meta が無い/空(=まだ再取込していない本棚)なら【完全に無操作】。
'     判定は modChunkMeta.GraphActive 1本に集約する(条件が2箇所に割れると、
'     片方だけ直したときに古い本棚で実行時エラーが出る)。
'   ・config graph_refs=off でも無操作。ゲートの読みはこの層に閉じるので、
'     呼び出し元(modAskThorough/modAskMulti/modAskRetrieve)は1行のまま。
'   ・LLMは1回も呼ばない。増えるのはワークシート読み1〜2回だけ。
'   ・足したチャンクの score は 0(NeighborExpand と同じ理由=信頼度バッジを
'     水増ししない)。ArticleEnsure も同じ(2026-08-05 R17H FA-4 / A-M4)。
'     以前は先頭ヒットと同値にしていたが、それは「検索で当たっていない
'     チャンクに、当たったチャンクの点数を貸す」ことで、1件の強ヒットと
'     その条番号一致で modAsk.LastConfidence の「強く一致」が成立してしまう。
'     位置は先頭のまま(順位で本命だと示し、点数では騙さない)=
'     max_context_chars の打ち切りで真っ先に落ちることも無い。
' ============================================================================

' config graph_refs(既定on)。off のときだけ切る。
Private Function GraphGateOn() As Boolean
    GraphGateOn = (LCase$(Trim$(modConfig.GetString("graph_refs", "on"))) <> "off")
End Function

' ----------------------------------------------------------------------------
' RefsExpand - ヒット群が参照している条文・別表のチャンクを末尾へ足す。
'   maxAdd <= 0 なら既定 REFS_ADD_DEFAULT(8)。nHits は足したぶんだけ増える。
'   足す条件は「同じ資料(source)の中で、section_path に参照ラベルを含む」。
'   資料をまたがない理由: 「第8条」は資料ごとに別の条文で、跨いだ瞬間に
'   まったく無関係な規程の第8条が根拠として混ざる(出典タグは正しいので
'   利用者は気付けない=一番危険な外し方)。
' ----------------------------------------------------------------------------
Public Sub RefsExpand(ByRef hits() As Hit, ByRef nHits As Long, ByVal maxAdd As Long)
    If nHits < 1 Then Exit Sub
    If Not GraphGateOn() Then Exit Sub

    Dim cap As Long: cap = maxAdd
    If cap < 1 Then cap = REFS_ADD_DEFAULT
    If cap > MAX_ADD Then cap = MAX_ADD

    On Error GoTo Quiet

    Dim metaIds() As String, metaPaths() As String, metaRefs() As String
    Dim metaN As Long: metaN = modChunkMetaStore.ReadAllMeta(metaIds, metaPaths, metaRefs)
    If Not modChunkMeta.GraphActive(metaN, nHits) Then Exit Sub

    ' 1) ヒットの chunk_id(箱)と source(箱)を作る。
    Dim hitBox As String, srcBox As String
    Dim i As Long
    For i = 1 To nHits
        hitBox = hitBox & vbLf & Trim$(hits(i).chunk_id)
        Dim sName As String: sName = Trim$(hits(i).source)
        If LenB(sName) > 0 Then
            If InStr(1, srcBox, "|" & sName & "|", vbBinaryCompare) = 0 Then
                srcBox = srcBox & "|" & sName & "|"
            End If
        End If
    Next i
    hitBox = hitBox & vbLf
    If LenB(srcBox) = 0 Then Exit Sub

    ' 2) ヒット行の refs_out から参照ラベルを集める(自分自身の条番号は除く)。
    Dim labLine As String
    For i = 0 To metaN - 1
        If InStr(1, hitBox, vbLf & metaIds(i) & vbLf, vbBinaryCompare) > 0 Then
            labLine = MergeLabels(labLine, modChunkMeta.RefLabelsFor(metaRefs(i), metaPaths(i)))
        End If
    Next i
    If LenB(labLine) = 0 Then
        LogRefsZero "no_ref"
        Exit Sub
    End If

    ' 3) ラベルを含む section_path のチャンクを候補にする(ヒット自身は除く)。
    Dim labs() As String: labs = Split(labLine, "|")
    Dim candBox As String
    Dim candN As Long
    candN = CollectByLabel(metaIds, metaPaths, metaN, labs, hitBox, candBox)
    If candN < 1 Then
        LogRefsZero "no_cand"
        Exit Sub
    End If

    ' 4) my_knowledge から候補行を引き、同じ資料のものだけ末尾へ足す。
    Dim added As Long
    added = AppendByIds(candBox, srcBox, cap, hits, nHits)
    If added < 1 Then
        LogRefsZero "no_row"
        Exit Sub
    End If

    On Error Resume Next
    modLog.LogUsage "refs_expand", "focus", "labels=" & labLine & " added=" & added, 0, nHits
    On Error GoTo 0
    Exit Sub

Quiet:
    Resume RefsDone
RefsDone:
    Err.Clear
    LogRefsZero "err"
End Sub

' ----------------------------------------------------------------------------
' ArticleEnsure - 質問が名指しした条番号のチャンクを、必ず材料に入れる。
'   「第5条の免責は?」に対して dense+sparse が第5条を1件も返さないことは
'   実際に起きる(条番号は短くて特徴が薄いため)。section_path という
'   決定的なキーが手元にあるのだから、確率的な検索の後ろで1回だけ確かめる。
'   ・既に条番号一致のチャンクがヒットに居れば【何もしない】(順位を触らない)。
'   ・入れる場合は先頭へ最大 maxIns 件。score は 0(R17H FA-4。理由は冒頭)。
'   ・scopeSrcBox(R17H FA-3): 空でなければ、その資料(source)の行だけを注入
'     する。深掘りのようにスコープを絞って検索した経路で、スコープ外の資料を
'     注ぎ足すと「会話の資料の中で何件見つかったか」の判定(RunDeepScoped の
'     n>=2)が注入分だけで成立し、スコープの外で答えたのに内で答えたことになる。
'     形式は AppendByIds の srcBox と同じ "|資料名|" の連結。
'   ・allowZeroHits(R17H FA-8): True なら nHits=0 でも動く(検索が1件も
'     当たらなかったときに、直接キーで最大 maxIns 件だけ材料を用意する経路)。
'     既定 False = 従来どおり「ヒットが1件以上あるとき」だけ働く。
'   ・空でない scopeSrcBox を渡さない限り資料は跨いでよい(質問が資料を名指し
'     していないため。逆に RefsExpand は跨がない=既に根拠になっている資料の
'     中の話だから)。
' ----------------------------------------------------------------------------
Public Sub ArticleEnsure(ByVal query As String, ByRef hits() As Hit, _
                         ByRef nHits As Long, ByVal maxIns As Long, _
                         Optional ByVal scopeSrcBox As String = "", _
                         Optional ByVal allowZeroHits As Boolean = False)
    If nHits < 1 And Not allowZeroHits Then Exit Sub
    If Not GraphGateOn() Then Exit Sub

    Dim cap As Long: cap = maxIns
    If cap < 1 Then cap = 2

    ' 条番号・別表・様式が1つも無い質問はここで終わり(大多数の質問)。
    Dim labLine As String: labLine = modChunkMeta.ExtractRefs(query)
    If LenB(labLine) = 0 Then Exit Sub

    On Error GoTo Quiet2

    Dim metaIds() As String, metaPaths() As String, metaRefs() As String
    Dim metaN As Long: metaN = modChunkMetaStore.ReadAllMeta(metaIds, metaPaths, metaRefs)
    If Not modChunkMeta.GraphActive(metaN, nHits, allowZeroHits) Then Exit Sub

    Dim hitBox As String
    Dim i As Long
    For i = 1 To nHits
        hitBox = hitBox & vbLf & Trim$(hits(i).chunk_id)
    Next i
    hitBox = hitBox & vbLf

    ' 一致チャンクが既にヒットに居るか / 居なければ候補は何か、を1周で見る。
    Dim labs() As String: labs = Split(labLine, "|")
    Dim candBox As String
    Dim candN As Long
    Dim alreadyHit As Boolean
    For i = 0 To metaN - 1
        If LenB(metaPaths(i)) > 0 Then
            If AnyLabelIn(metaPaths(i), labs) Then
                If InStr(1, hitBox, vbLf & metaIds(i) & vbLf, vbBinaryCompare) > 0 Then
                    alreadyHit = True
                    Exit For
                End If
                If candN < REFS_MAX_CAND Then
                    candBox = candBox & vbLf & metaIds(i)
                    candN = candN + 1
                End If
            End If
        End If
    Next i
    If alreadyHit Or candN < 1 Then Exit Sub
    candBox = candBox & vbLf

    Dim baseN As Long: baseN = nHits
    Dim added As Long
    added = AppendByIds(candBox, scopeSrcBox, cap, hits, nHits)
    If added < 1 Then Exit Sub

    ' 末尾へ付いた added 件を先頭へ回す(順位は「本命が先」。score は
    ' AppendByIds が入れた 0 のまま=点数は騙さない。R17H FA-4)。
    Dim tmp() As Hit: ReDim tmp(1 To added)
    For i = 1 To added
        tmp(i) = hits(baseN + i)
    Next i
    For i = baseN To 1 Step -1
        hits(i + added) = hits(i)
    Next i
    For i = 1 To added
        hits(i) = tmp(i)
    Next i

    On Error Resume Next
    modLog.LogUsage "article_ensure", "focus", _
        "labels=" & labLine & " added=" & added & " base=" & baseN, 0, nHits
    On Error GoTo 0
    Exit Sub

Quiet2:
    Resume ArtDone
ArtDone:
    Err.Clear
    LogRefsZero "err_article"
End Sub

' 参照展開が1件も足せなかった理由を1行残す(NeighborExpand の LogZero と同じ
' 考え方。効いていない状態が無音で続くのを防ぐ=R16H FA-1 の教訓)。
Private Sub LogRefsZero(ByVal reason As String)
    On Error Resume Next
    modLog.LogUsage "refs_zero", "focus", "why=" & reason
    On Error GoTo 0
End Sub

' "|" 区切りのラベル列を重複なしで足し合わせる。
Private Function MergeLabels(ByVal acc As String, ByVal addLine As String) As String
    Dim outS As String: outS = acc
    If LenB(addLine) = 0 Then
        MergeLabels = outS
        Exit Function
    End If
    Dim arr() As String: arr = Split(addLine, "|")
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If LenB(arr(i)) > 0 Then
            If InStr(1, "|" & outS & "|", "|" & arr(i) & "|", vbBinaryCompare) = 0 Then
                If LenB(outS) > 0 Then outS = outS & "|"
                outS = outS & arr(i)
            End If
        End If
    Next i
    MergeLabels = outS
End Function

' section_path がラベル列のどれかを含むか。
Private Function AnyLabelIn(ByVal sectionPath As String, ByRef labs() As String) As Boolean
    Dim i As Long
    For i = LBound(labs) To UBound(labs)
        If modChunkMeta.PathHasLabel(sectionPath, labs(i)) Then
            AnyLabelIn = True
            Exit Function
        End If
    Next i
End Function

' chunk_meta を1周し、ラベルに当たる chunk_id を候補箱(vbLf区切り)へ集める。
' ヒット自身は除く(既に材料になっている)。戻り値=候補件数。
Private Function CollectByLabel(ByRef metaIds() As String, ByRef metaPaths() As String, _
                                ByVal metaN As Long, ByRef labs() As String, _
                                ByVal hitBox As String, ByRef outBox As String) As Long
    Dim n As Long
    Dim i As Long
    For i = 0 To metaN - 1
        If LenB(metaPaths(i)) > 0 Then
            If InStr(1, hitBox, vbLf & metaIds(i) & vbLf, vbBinaryCompare) = 0 Then
                If AnyLabelIn(metaPaths(i), labs) Then
                    outBox = outBox & vbLf & metaIds(i)
                    n = n + 1
                    If n >= REFS_MAX_CAND Then Exit For
                End If
            End If
        End If
    Next i
    If n > 0 Then outBox = outBox & vbLf
    CollectByLabel = n
End Function

' 候補箱の chunk_id を my_knowledge から引いて hits の末尾へ足す。
' srcBox が空でなければ、その資料(source)の行だけを採る。戻り値=足した件数。
' 読み方は NeighborExpand と同じ「chunk_id列とsource列を1回だけRange読み」。
Private Function AppendByIds(ByVal candBox As String, ByVal srcBox As String, _
                             ByVal cap As Long, ByRef hits() As Hit, _
                             ByRef nHits As Long) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then Exit Function

    Dim lastK As Long
    lastK = ws.Cells(ws.Rows.count, COL_ID).End(xlUp).row
    If lastK < 3 Then Exit Function

    Dim idData As Variant
    idData = ws.Range(ws.Cells(2, COL_ID), ws.Cells(lastK, COL_SOURCE)).Value

    Dim added As Long
    Dim i As Long
    For i = LBound(idData, 1) To UBound(idData, 1)
        If added >= cap Then Exit For
        Dim id As String: id = Trim$(CStr(idData(i, COL_ID)))
        Dim src As String: src = Trim$(CStr(idData(i, COL_SOURCE)))
        If LenB(id) > 0 And LenB(src) > 0 Then
            If InStr(1, candBox, vbLf & id & vbLf, vbBinaryCompare) > 0 Then
                If LenB(srcBox) = 0 Or InStr(1, srcBox, "|" & src & "|", vbBinaryCompare) > 0 Then
                    ' 0件から足す経路(R17H FA-8)では hits() がまだ確保されて
                    ' いない。ReDim Preserve を掛けずに1件目を作る
                    ' (modAskMulti.MergeHits と同じ作法)。
                    If nHits = 0 Then
                        ReDim hits(1 To 1)
                    Else
                        ReDim Preserve hits(1 To nHits + 1)
                    End If
                    nHits = nHits + 1
                    added = added + 1
                    FillHitFromRow ws, i + 1, hits(nHits)   ' idData は2行目起点
                End If
            End If
        End If
    Next i
    AppendByIds = added
End Function

' ============================================================================
' 純ロジック(LibreOfficeの実行テストで固定する)
' ============================================================================

' ----------------------------------------------------------------------------
' ParseChunkKey - chunk_id "bs::ハッシュ::pN::cM" を分解する。
'   docKey = 末尾2要素を除いた全部(="bs::チャンク本文のハッシュ")。
'     【資料キーではない】。このハッシュは行ごとに一意なので、これで資料を
'     まとめると全ての行が別資料になる(R16H FA-1)。資料の同一性は source 列で
'     見ること。ここが返すのは「形式が正しいか」と (page, seq) を得るためのもの。
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
'   orderedIds : 候補の chunk_id(vbLf区切り。並び順は問わない)
'   orderedSrcs: 同じ並びの source(vbLf区切り。orderedIds と1対1で対応させる)
'   hitIds     : ヒットの chunk_id(vbLf区切り)
'   戻り値     : 足すべき chunk_id(vbLf区切り・文書順・重複なし・ヒット自身は除く)
'
'   ・(source, page, seq) の昇順に並べてから前後を取る(ページ跨ぎでも
'     seq のリセットに引きずられない)。並びの主キーが chunk_id のハッシュ部
'     ではなく source なのが R16H FA-1 の要点で、ハッシュは行ごとに違うため
'     資料をまとめる役には立たない。
'   ・資料(source)の境界は越えない(別の資料の先頭が「前のチャンク」にならない)。
'   ・窓が重なっても結果は重複しない(印を付けてから1回だけ書き出す)。
'   ・読めない chunk_id / source が無い行は並びから外す(位置が決められない
'     ものを混ぜると、その分だけ隣が1つずれる)。
' ----------------------------------------------------------------------------
Public Function NeighborIdList(ByVal orderedIds As String, ByVal orderedSrcs As String, _
                               ByVal hitIds As String, ByVal radius As Long) As String
    If radius < 1 Then Exit Function
    If LenB(orderedIds) = 0 Or LenB(hitIds) = 0 Then Exit Function

    Dim raw() As String: raw = Split(orderedIds, vbLf)
    Dim srcRaw() As String: srcRaw = Split(orderedSrcs, vbLf)
    Dim cap As Long: cap = UBound(raw) - LBound(raw) + 1
    If cap < 1 Then Exit Function
    Dim srcN As Long: srcN = UBound(srcRaw) - LBound(srcRaw) + 1

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
        Dim src As String: src = ""
        If i - LBound(raw) < srcN Then src = Trim$(srcRaw(LBound(srcRaw) + i - LBound(raw)))
        If LenB(src) > 0 Then
            If ParseChunkKey(raw(i), d, p, s) Then
                n = n + 1
                ids(n) = Trim$(raw(i))
                dk(n) = src
                pg(n) = p
                sq(n) = s
            End If
        End If
    Next i
    If n < 1 Then Exit Function

    ' 挿入ソート((source, page, seq) 昇順)。候補はヒット周辺だけに絞って
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

' (source, page, seq) の大小比較。a > b なら True。
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

' ヒットから「探す資料(source)」と「見るページ窓(source#page)」を作る。
' 窓を作るのは、本棚全体の chunk_id を並べ替えないため。radius が何チャンク
' でも、離れるページ数は radius を超えない(1ページに最低1チャンクあるため)。
' 資料の一致は source で見る(chunk_id のハッシュ部は行ごとに違う。R16H FA-1)。
Private Function BuildWindows(hits() As Hit, ByVal nHits As Long, ByVal radius As Long, _
                              ByRef hitIds As String, ByRef needKeys As String, _
                              ByRef winLine As String) As Boolean
    Dim i As Long, p As Long
    For i = 1 To nHits
        Dim d As String, pg As Long, sq As Long
        Dim src As String: src = Trim$(hits(i).source)
        If LenB(src) > 0 Then
            If ParseChunkKey(hits(i).chunk_id, d, pg, sq) Then
                If LenB(hitIds) > 0 Then hitIds = hitIds & vbLf
                hitIds = hitIds & Trim$(hits(i).chunk_id)
                If InStr(1, needKeys, "|" & src & "|", vbBinaryCompare) = 0 Then
                    needKeys = needKeys & "|" & src & "|"
                End If
                For p = pg - radius To pg + radius
                    If p >= 0 Then
                        Dim w As String: w = "|" & src & "#" & p & "|"
                        If InStr(1, winLine, w, vbBinaryCompare) = 0 Then winLine = winLine & w
                    End If
                Next p
            End If
        End If
    Next i
    BuildWindows = (LenB(hitIds) > 0)
End Function

' chunk_id 列と source 列を1回だけ走査し、窓に入る行の id・source・行番号を集める。
Private Function CollectCandidates(ByVal idData As Variant, ByVal needKeys As String, _
                                   ByVal winLine As String, ByRef candIds As String, _
                                   ByRef candSrcs As String, ByRef candRow() As Long) As Long
    Dim lo As Long: lo = LBound(idData, 1)
    Dim hi As Long: hi = UBound(idData, 1)
    ReDim candRow(1 To hi - lo + 1)

    Dim n As Long: n = 0
    Dim i As Long
    For i = lo To hi
        Dim id As String: id = CStr(idData(i, COL_ID))
        Dim src As String: src = Trim$(CStr(idData(i, COL_SOURCE)))
        If LenB(id) > 0 And LenB(src) > 0 Then
            Dim d As String, pg As Long, sq As Long
            If ParseChunkKey(id, d, pg, sq) Then
                If InStr(1, needKeys, "|" & src & "|", vbBinaryCompare) > 0 Then
                    If InStr(1, winLine, "|" & src & "#" & pg & "|", vbBinaryCompare) > 0 Then
                        n = n + 1
                        candRow(n) = i + 1          ' idData は2行目起点
                        If LenB(candIds) > 0 Then candIds = candIds & vbLf
                        candIds = candIds & id
                        If LenB(candSrcs) > 0 Then candSrcs = candSrcs & vbLf
                        candSrcs = candSrcs & src
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
