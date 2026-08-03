Attribute VB_Name = "modRetrieve"
Option Explicit

' ============================================================================
' modRetrieve - my_vectors全件をコサイン類似度(内積)でスコアリングする検索
' ----------------------------------------------------------------------------
' 役割:
'   質問文をベクトル化し、my_vectors(埋め込み済みチャンク全件)との内積
'   (全てL2正規化済みなのでコサイン類似度と同値)でスコアリングし、
'   上位topK件をmy_knowledgeのメタ情報と合わせてHit配列で返す(MASTER_SPEC
'   §7.3)。
'
' 設計判断:
'   ・V2 src/chatbot_v2/modRetrieve.bas の「配列一括読込み→ストリーミング
'     top-k(最小値を追跡して置換)→選択ソート」パターンをそのまま踏襲する。
'     топK(quick=6/deep=12)は小さいため、全件ソートよりストリーミング
'     top-kの方が効率的、というV2の設計判断を引き継ぐ。
'   ・my_vectors(chunk_id, vector_csv)とmy_knowledge(chunk_id, source,
'     origin, page, summary, keywords, full_text, ...)をそれぞれ1回の
'     Range.Value一括読込みで取得し(§12: ループ内Range直接アクセス禁止)、
'     chunk_id→my_knowledge行番号の対応はScripting.Dictionaryで引く
'     (src/ingest/modShelf.bas が既に同じ手法を使っており本アプリ内で
'     実証済み)。modRetrieveはR4対象外(§11.1のPURE_LOGIC_MODULESに
'     含まれない)なのでExcelオブジェクト・CreateObjectの利用は自由。
'   ・キーワード側は modSparse(文字bigram + BM25 + 完全一致)へ委譲する。
'     旧実装は「質問文を空白で分割して語ごとに+0.05(上限+0.15)」
'     「本文への出現で+0.06(上限+0.24)」「質問文まるごと一致で+0.15」
'     という加点の積み上げだったが、日本語は空白で区切らないため
'     実測 R@1 は 32% しか出なかった(本実装は 84%)。
'     2026-07-28(解説書 §11-8/§11-9): 単段 Search だけが新実装へ移行し、
'     既定である多段 SearchExpanded は旧実装のまま取り残されていた。
'     つまり測って上げた精度が、既定構成では一度も効いていなかった。
'     両方を modSparse へ寄せ、スコア式はこのモジュールに1つだけにする。
'     式が1つなら「多段が失敗して単段へ落ちたら結果が変わる」も起きない。
'   ・GetEmbedding失敗(空配列)時は-1を返し、実際のE0203ログは
'     modGateway.GetEmbedding内で既に行われているため、ここで重複して
'     modLogを呼ばない(呼び出し側=modAskがFriendlyMessage("E0203")を
'     組み立てて表示する)。
'   ・my_vectors/my_knowledgeが無い/空の場合は例外にせず0件を返す
'     (検索0件の扱いは呼び出し側=modAskの責務: MASTER_SPEC §7.3)。
'   ・Hit.full_text: my_knowledge.full_text列(チャンク本文全体)をそのまま
'     格納する(Wave3 PM裁定1)。preview(先頭120字)は出典先出し表示専用として
'     引き続き別に保持し、full_textは回答生成の根拠としてmodPromptsが使う。
'     既にfull_textを1回のRange一括読込み(kData)で読んでいるため、追加の
'     シートアクセスは発生しない。
' ============================================================================

Private Const COL_V_ID As Long = 1
Private Const COL_V_VEC As Long = 2

Private Const COL_K_ID As Long = 1
Private Const COL_K_SOURCE As Long = 2
Private Const COL_K_ORIGIN As Long = 3
Private Const COL_K_PAGE As Long = 4
Private Const COL_K_SUMMARY As Long = 5
Private Const COL_K_KEYWORDS As Long = 6
Private Const COL_K_FULLTEXT As Long = 7
' 第10列 norm_text は 2026-08-01(R12-4)で追加した「照合用の正規化済みテキスト」。
' 取込時に1回だけ作り、検索側は読むだけにする。列の中身の作り方・空欄の
' 遅延バックフィル・書き戻しは my_knowledge の行操作層(modShelfStore)が持つ。
Private Const COL_K_NORM As Long = 10

Private Const PREVIEW_LEN As Long = 120

' 重ループの進捗更新+DoEvents間隔(行)。憲章§3-2「無反応を作らない」。
Private Const PROGRESS_STEP As Long = 2000


' modSparse のスコアをベクトル(-1～1)と同じ土俵へ乗せるための係数。
' SPARSE_WEIGHT: BM25(質問長で正規化済み)にかける倍率
' EXACT_WEIGHT : 条番号・型番の完全一致1件あたりの加点。決定的に効かせる
' KeyScore は 0～20程度のスケール。ベクトル(-1～1)と同じ土俵に乗せる係数。
' 大きすぎるとキーワードだけで順位が決まり、小さすぎるとベクトルに埋もれる。
' 実測(31問)で R@1 が最大になる範囲の中央を採った。
Private Const SPARSE_WEIGHT As Double = 0.06

' ----------------------------------------------------------------------------
' Search - MASTER_SPEC §7.3 唯一の公開関数。
'   戻り値 = 件数(0可)。埋め込み失敗時は-1(呼び出し側がE0203表示)。
' ----------------------------------------------------------------------------
' scopeSources(R13-5a): 許可資料名のDictionary。Nothing=従来動作(既存呼び出しは
'   無指定で完全互換)。判定は excl.Exists の真隣に対称に置く。行の採用点は
'   Search/SearchExpanded とも1箇所で、粗選別も再スキャンも必ずそこを通る。
' ----------------------------------------------------------------------------
Public Function Search(ByVal query As String, ByVal topK As Long, ByRef hits() As Hit, _
                       Optional ByVal scopeSources As Object) As Long
    Dim emptyHits() As Hit
    hits = emptyHits

    Dim k As Long
    k = topK
    If k < 1 Then k = 1

    Dim qv() As Double
    qv = modGateway.GetEmbedding(query)
    If Not modUtil.HasVector(qv) Then
        Search = -1
        Exit Function
    End If
    Dim qDim As Long: qDim = UBound(qv) - LBound(qv) + 1

    Dim wsV As Worksheet
    Dim wsK As Worksheet
    Set wsV = GetSheet(modAppDef.SH_VECTORS)
    Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If wsV Is Nothing Or wsK Is Nothing Then
        Search = 0
        Exit Function
    End If

    Dim lastV As Long
    lastV = wsV.Cells(wsV.Rows.count, COL_V_ID).End(xlUp).row
    If lastV < 2 Then
        Search = 0
        Exit Function
    End If

    Dim lastK As Long
    lastK = wsK.Cells(wsK.Rows.count, COL_K_ID).End(xlUp).row
    If lastK < 2 Then
        Search = 0
        Exit Function
    End If

    ' R12-4: ベクトルはセッション内キャッシュ(modVecCache)から引く。
    ' キャッシュが生きていれば vector_csv 列は読み込まない(idDataだけ)。
    ' 作れなかった/次元が食い違うときは False が返り、vData に従来どおり
    ' 2列が読み込まれて元の経路(毎回パース)で動く=機能は落とさない。
    Dim idData As Variant, vData As Variant
    Dim useCache As Boolean
    useCache = modVecCache.PrepareVectors(wsV, lastV, qDim, idData, vData)
    If IsEmpty(idData) Then          ' シートを読めなかった(0件と同じ扱いで静かに終える)
        Search = 0
        Exit Function
    End If

    ' my_knowledgeは1～10列目を読む(10列目=norm_text)。
    Dim kData As Variant
    kData = wsK.Range(wsK.Cells(2, COL_K_ID), wsK.Cells(lastK, COL_K_NORM)).Value
    modShelfStore.ResetNormBackfill

    Dim idx As Object
    Set idx = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(kData, 1) To UBound(kData, 1)
        Dim kid As String
        kid = CStr(kData(i, COL_K_ID))
        If LenB(kid) > 0 Then
            If Not idx.Exists(kid) Then idx.Add kid, i
        End If
    Next i


    ' キーワード側の準備は「1クエリにつき1回」だけ。
    ' 以前はチャンクごとに modSparse.Tokenize を呼んでおり、実測 225ms/件、
    ' 9000件で34分かかる状態だった(LibreOffice実測・閾値テストで固定済み)。
    ' いまはチャンクごとの処理を InStr の部分一致だけに抑えている。
    Dim sparseKeys As String
    On Error Resume Next
    sparseKeys = modSparse.DistinctiveKeys(query)
    On Error GoTo 0

    ' ナレッジ自浄: ノイズ報告が閾値以上の資料を検索対象から論理除外する
    Dim excl As Object: Set excl = modStats.ExcludedSources()

    ' [狂気案Lv.1] バイナリ量子化ハイブリッド: 大規模KB(binary_rag_min以上)かつ
    ' binary_rag=TRUEのときだけ、ハミング距離で候補行を粗選別する。以降のFloat
    ' コサイン/キーワードボーナス/ストリーミングtop-kは一切不変(候補以外をスキップ
    ' する1行のガードを足すだけ)。Prefilterが辞退(次元不一致等)したら全件Floatへ。
    Dim nV As Long: nV = UBound(idData, 1) - LBound(idData, 1) + 1
    Dim useCand As Boolean: useCand = False
    Dim candRows As Object
    Dim binMs As Double: binMs = 0
    If modBitwiseOpt.Enabled(nV) Then
        Dim tB0 As Double: tB0 = modBitwiseOpt.MicroTimerMs()
        useCand = modBitwiseOpt.Prefilter(qv, vData, modBitwiseOpt.PrefilterN(), candRows)
        binMs = modBitwiseOpt.MicroTimerMs() - tB0   ' 爆速証明: バイナリ粗選別のms
        If useCand Then UnionKeyMatchRows idData, kData, idx, sparseKeys, candRows
    End If

    Dim bestId() As String: ReDim bestId(1 To k)
    Dim bestScore() As Double: ReDim bestScore(1 To k)
    Dim bestSource() As String: ReDim bestSource(1 To k)
    Dim bestPage() As Long: ReDim bestPage(1 To k)
    Dim bestPreview() As String: ReDim bestPreview(1 To k)
    Dim bestOrigin() As String: ReDim bestOrigin(1 To k)
    Dim bestFullText() As String: ReDim bestFullText(1 To k)
    Dim filled As Long: filled = 0
    Dim minIdx As Long: minIdx = 0
    Dim minScore As Double: minScore = 0
    Dim e0702Logged As Boolean: e0702Logged = False

    Dim tF0 As Double: tF0 = modBitwiseOpt.MicroTimerMs()   ' 爆速証明: Float再ランク計測開始
    Dim r As Long
    Dim rLo As Long: rLo = LBound(idData, 1)
RescanAll:
    For r = rLo To UBound(idData, 1)
        ' R12-4: 20,500行の全走査は数秒〜十数秒かかる。無反応にしない。
        If ((r - rLo) Mod PROGRESS_STEP) = 0 And r > rLo Then
            TickScan r - rLo, nV
            ' R12-H-2: DoEvents の隙に取込/同期がキャッシュを解放し得る。
            ' 気付かずに続けると SlotOfRow が全て -1 を返し、残りの行が
            ' 【無言で検索結果から消える】。検知したら記録して、その質問は
            ' 従来経路で最初からやり直す(欠けた結果を返すより時間を使う)。
            If useCache Then
                If Not modVecCache.Ready() Then
                    LostMidScan "modRetrieve.Search"
                    useCache = False
                    modVecCache.LoadDirectVectors wsV, lastV, idData, vData
                    filled = 0
                    minIdx = 0
                    minScore = 0
                    e0702Logged = False
                    GoTo RescanAll
                End If
            End If
        End If
        If useCand Then If Not candRows.Exists(r) Then GoTo NextR   ' [Lv.1] 粗選別候補以外は除外
        Dim vid As String
        vid = CStr(idData(r, 1))
        If LenB(vid) = 0 Then GoTo NextR
        If Not idx.Exists(vid) Then GoTo NextR

        ' Wave4修正: modUtil.DotProductは次元不一致を「呼び出し側が別途次元検査を
        ' 行う規約」(modUtil.bas冒頭コメント)としているが、以前はここで一切
        ' 検査せず0点のままランキングに混ぜていた。mock_llm切替等でembed_dimが
        ' 変わった保存済みベクトルが残っていると、全チャンクが無言で0点になり
        ' err_log/診断のどちらにも痕跡が残らない不具合があった(E0702は元々
        ' パック検証専用だったが、検索時の次元不一致もE0702として記録する)。
        ' R12-4: キャッシュ経路では次元判定も構築時に済んでいる(slot=-2)。
        Dim slot As Long: slot = -1
        Dim vv() As Double
        If useCache Then
            slot = modVecCache.SlotOfRow(r)
            If slot < 0 Then
                If slot = -2 And Not e0702Logged Then
                    modLog.LogError "E0702", "modRetrieve.Search", _
                        "chunk_id=" & vid & " queryDim=" & qDim & " vectorDim=" & modVecCache.DimOfRow(r)
                    e0702Logged = True
                End If
                GoTo NextR
            End If
        Else
            Dim vcsv As String
            vcsv = CStr(vData(r, COL_V_VEC))
            If LenB(vcsv) = 0 Then GoTo NextR
            If Not modUtil.CsvToVector(vcsv, vv) Then GoTo NextR
            If UBound(vv) - LBound(vv) + 1 <> qDim Then
                If Not e0702Logged Then
                    modLog.LogError "E0702", "modRetrieve.Search", _
                        "chunk_id=" & vid & " queryDim=" & qDim & " vectorDim=" & (UBound(vv) - LBound(vv) + 1)
                    e0702Logged = True
                End If
                GoTo NextR
            End If
        End If

        Dim kRow As Long
        kRow = idx.Item(vid)
        Dim srcName As String: srcName = CStr(kData(kRow, COL_K_SOURCE))
        If excl.Exists(srcName) Then GoTo NextR   ' ノイズ報告で論理除外された資料
        If Not scopeSources Is Nothing Then If Not scopeSources.Exists(srcName) Then GoTo NextR
        Dim origin As String: origin = CStr(kData(kRow, COL_K_ORIGIN))
        Dim pageNum As Long: pageNum = CLng(Val(kData(kRow, COL_K_PAGE)))
        Dim fullText As String: fullText = CStr(kData(kRow, COL_K_FULLTEXT))

        Dim sc As Double
        If useCache Then
            sc = modVecCache.DotAt(qv, slot)      ' 加算順序はDotProductと同一(等価テストで固定)
        Else
            sc = modUtil.DotProduct(qv, vv)
        End If
        ' 日本語のキーワード側は modSparse(文字bigram+BM25+完全一致)へ委譲する。
        ' 旧実装(空白分割+一律加点)は実測 R@1 32%、本実装は 84%。
        ' 差の大半は「日本語は空白で区切らない」という一点から来ていた。
        ' ベクトル(意味)とキーワード(完全一致)は別々に効かせる。
        ' キーワード側は照合用の空白除去テキストに対して当てるので、
        ' PDF字詰めや利用者の余計な空白があっても一致する。
        sc = sc + SparseBoost(sparseKeys, modShelfStore.NormTextAt(kData, kRow))

        If filled < k Then
            filled = filled + 1
            bestId(filled) = vid
            bestScore(filled) = sc
            bestSource(filled) = srcName
            bestPage(filled) = pageNum
            bestPreview(filled) = modUtil.SafeLeft(fullText, PREVIEW_LEN)
            bestOrigin(filled) = origin
            bestFullText(filled) = fullText
            If filled = k Then RecomputeMin bestScore, k, minIdx, minScore
        ElseIf sc > minScore Then
            bestId(minIdx) = vid
            bestScore(minIdx) = sc
            bestSource(minIdx) = srcName
            bestPage(minIdx) = pageNum
            bestPreview(minIdx) = modUtil.SafeLeft(fullText, PREVIEW_LEN)
            bestOrigin(minIdx) = origin
            bestFullText(minIdx) = fullText
            RecomputeMin bestScore, k, minIdx, minScore
        End If
NextR:
    Next r

    ' 爆速証明ログ(バイナリ選別が作動したときだけ。Debug.Print + debug時Toast)
    If useCand Then
        modBitwiseOpt.LogPerf binMs, modBitwiseOpt.MicroTimerMs() - tF0, candRows.count, nV
    End If
    modShelfStore.FlushNormBackfill wsK, lastK, kData

    If filled = 0 Then
        Search = 0
        Exit Function
    End If

    ' filled <= topK(既定6/12)と小さいため選択ソートで十分(V2実証パターン踏襲)
    Dim a As Long, b As Long, mx As Long
    For a = 1 To filled - 1
        mx = a
        For b = a + 1 To filled
            If bestScore(b) > bestScore(mx) Then mx = b
        Next b
        If mx <> a Then
            SwapResult bestId, bestScore, bestSource, bestPage, bestPreview, bestOrigin, bestFullText, a, mx
        End If
    Next a

    Dim outHits() As Hit
    ReDim outHits(1 To filled)
    For a = 1 To filled
        outHits(a).chunk_id = bestId(a)
        outHits(a).score = bestScore(a)
        outHits(a).source = bestSource(a)
        outHits(a).page = bestPage(a)
        outHits(a).preview = bestPreview(a)
        outHits(a).origin = bestOrigin(a)
        outHits(a).full_text = bestFullText(a)
    Next a

    hits = outHits
    Search = filled
End Function

' ----------------------------------------------------------------------------
' SearchExpanded - マルチクエリ検索(設計書§C-2)。複数クエリを同一ストアで
'   スコアリングし、chunk_id単位のunion(最大スコア採用)から上位poolK件を返す。
'   シート読込はクエリ数に関わらず1回。有効クエリ0本・空ストアは0を返す。
'   既存Search(単段)の挙動には一切影響しない。
'   scopeSources(R13-5a)の意味と置き場所は Search と同じ。
' ----------------------------------------------------------------------------
Public Function SearchExpanded(queries() As String, ByVal poolK As Long, ByRef hits() As Hit, _
                               Optional ByVal scopeSources As Object) As Long
    Dim emptyHits() As Hit
    hits = emptyHits
    SearchExpanded = 0

    Dim pk As Long: pk = poolK
    If pk < 1 Then pk = 1

    Dim qLo As Long, qHi As Long
    Dim badArr As Boolean: badArr = False
    On Error Resume Next
    qLo = LBound(queries)
    qHi = UBound(queries)
    badArr = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    If badArr Or qHi < qLo Then Exit Function

    Dim wsV As Worksheet, wsK As Worksheet
    Set wsV = GetSheet(modAppDef.SH_VECTORS)
    Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If wsV Is Nothing Or wsK Is Nothing Then Exit Function

    Dim lastV As Long
    lastV = wsV.Cells(wsV.Rows.count, COL_V_ID).End(xlUp).row
    If lastV < 2 Then Exit Function
    Dim lastK As Long
    lastK = wsK.Cells(wsK.Rows.count, COL_K_ID).End(xlUp).row
    If lastK < 2 Then Exit Function

    ' R12-4: ベクトルはクエリ本数に関わらず1回だけパースする(modVecCache)。
    ' 多段は「クエリ数 × 全チャンク」を回るので、毎回のCSVパースが最も重かった。
    Dim idData As Variant, vData As Variant
    Dim useCache As Boolean: useCache = False
    Dim kData As Variant
    kData = wsK.Range(wsK.Cells(2, COL_K_ID), wsK.Cells(lastK, COL_K_NORM)).Value
    modShelfStore.ResetNormBackfill

    Dim idx As Object
    Set idx = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(kData, 1) To UBound(kData, 1)
        Dim kid As String
        kid = CStr(kData(i, COL_K_ID))
        If LenB(kid) > 0 Then
            If Not idx.Exists(kid) Then idx.Add kid, i
        End If
    Next i

    ' ナレッジ自浄: ノイズ報告が閾値以上の資料を検索対象から論理除外する
    Dim excl As Object: Set excl = modStats.ExcludedSources()

    Dim unionScore As Object
    Set unionScore = CreateObject("Scripting.Dictionary")
    Dim e0702Logged As Boolean: e0702Logged = False
    Dim validQ As Long: validQ = 0

    Dim qi As Long
    For qi = qLo To qHi
        Dim qText As String: qText = Trim$(queries(qi))
        If LenB(qText) = 0 Then GoTo NextQ

        Dim qv() As Double
        qv = modGateway.GetEmbedding(qText)
        If Not modUtil.HasVector(qv) Then GoTo NextQ
        validQ = validQ + 1
        Dim qDim As Long: qDim = UBound(qv) - LBound(qv) + 1

        Dim sparseKeys2 As String
        On Error Resume Next
        sparseKeys2 = modSparse.DistinctiveKeys(qText)
        On Error GoTo 0

        ' キャッシュはクエリごとに問い直す(2回目以降は判定だけで即返る)。
        ' 質問側の次元がキャッシュと違うクエリだけ従来経路へ落ちる。
        useCache = modVecCache.PrepareVectors(wsV, lastV, qDim, idData, vData)
        If IsEmpty(idData) Then GoTo NextQ    ' シートを読めなかった
        Dim nV As Long: nV = UBound(idData, 1) - LBound(idData, 1) + 1

        ' 2026-07-28(解説書 §11-7): 多段側にも粗選別を配線した。
        '
        ' バイナリ量子化による候補絞り込み(modBitwiseOpt)は単段 Search に
        ' しか組み込まれておらず、config の既定は retrieve_mode=multi なので、
        ' 【binary_rag=TRUE にしても既定構成では一切作動しない】状態だった。
        ' しかも多段は「クエリ数 × 全チャンク」の総当たりで、deep なら
        ' 最大5周する。速くしたい場所にこそ効いていなかったことになる。
        '
        ' 粗選別はクエリごとに行う(クエリが違えば近い候補も違う)。
        ' 候補が本棚を代表していないとき(次元混在など)は modBitwiseOpt 側が
        ' 辞退するので、その場合は従来どおり全件を見る。
        ' binary_rag の既定は FALSE のままなので、既定の挙動は変わらない。
        Dim useCandQ As Boolean: useCandQ = False
        Dim candRowsQ As Object
        If modBitwiseOpt.Enabled(nV) Then
            useCandQ = modBitwiseOpt.Prefilter(qv, vData, modBitwiseOpt.PrefilterN(), candRowsQ)
            If useCandQ Then UnionKeyMatchRows idData, kData, idx, sparseKeys2, candRowsQ
        End If

        Dim r As Long
        Dim rLo As Long: rLo = LBound(idData, 1)
RescanQuery:
        For r = rLo To UBound(idData, 1)
            If ((r - rLo) Mod PROGRESS_STEP) = 0 And r > rLo Then
                TickScan r - rLo, nV
                ' R12-H-2: 単段と同じ理由でキャッシュ消失を検知する。
                ' union は「同じchunk_idなら最大スコアを採る」ので、同じ行を
                ' もう一度採点しても結果は変わらない(スコアは両経路で同一)。
                If useCache Then
                    If Not modVecCache.Ready() Then
                        LostMidScan "modRetrieve.SearchExpanded"
                        useCache = False
                        modVecCache.LoadDirectVectors wsV, lastV, idData, vData
                        e0702Logged = False
                        GoTo RescanQuery
                    End If
                End If
            End If
            If useCandQ Then If Not candRowsQ.Exists(r) Then GoTo NextRow
            Dim vid As String
            vid = CStr(idData(r, 1))
            If LenB(vid) = 0 Then GoTo NextRow
            If Not idx.Exists(vid) Then GoTo NextRow
            Dim slot As Long: slot = -1
            Dim vv() As Double
            If useCache Then
                slot = modVecCache.SlotOfRow(r)
                If slot < 0 Then
                    If slot = -2 And Not e0702Logged Then
                        modLog.LogError "E0702", "modRetrieve.SearchExpanded", _
                            "chunk_id=" & vid & " queryDim=" & qDim & " vectorDim=" & modVecCache.DimOfRow(r)
                        e0702Logged = True
                    End If
                    GoTo NextRow
                End If
            Else
                Dim vcsv As String
                vcsv = CStr(vData(r, COL_V_VEC))
                If LenB(vcsv) = 0 Then GoTo NextRow
                If Not modUtil.CsvToVector(vcsv, vv) Then GoTo NextRow
                If UBound(vv) - LBound(vv) + 1 <> qDim Then
                    If Not e0702Logged Then
                        modLog.LogError "E0702", "modRetrieve.SearchExpanded", _
                            "chunk_id=" & vid & " queryDim=" & qDim & " vectorDim=" & (UBound(vv) - LBound(vv) + 1)
                        e0702Logged = True
                    End If
                    GoTo NextRow
                End If
            End If

            Dim kRow As Long: kRow = idx.Item(vid)
            Dim sName As String: sName = CStr(kData(kRow, COL_K_SOURCE))
            If excl.Exists(sName) Then GoTo NextRow   ' ノイズ論理除外
            If Not scopeSources Is Nothing Then If Not scopeSources.Exists(sName) Then GoTo NextRow

            ' 2026-07-28(解説書 §11-8/§11-9): スコア式を単段と同じにした。
            '
            ' それまで、単段 Search は modSparse(文字bigram+BM25+完全一致)へ
            ' 移行済みだったのに、多段 SearchExpanded だけが旧方式
            ' (KeywordBonus + GramBonus + FullTextBoost、加点上限の合計+0.64)
            ' のまま取り残されていた。config の既定は retrieve_mode=multi
            ' なので、【実測で R@1 32% -> 84% まで上げたスパース検索が、
            ' 既定の構成では一度も使われていなかった】ことになる。
            ' 直した側が利用者に届いていない、という最悪の形の作り込み。
            '
            ' これは解説書が指摘した2点の共通の根でもある:
            '   §11-8 加点上限+0.64 … 旧方式の合計。長い本文ほど有利になる
            '                          FullTextBoost もここに含まれる
            '   §11-9 単段と多段でスコアが違う … 式が2つあることそのもの
            ' 多段が例外で単段へフォールバックすると結果が変わる問題も、
            ' 式が1つになれば消える。
            '
            ' 数値を勘で調整するのではなく、既に実測した方の実装へ寄せる。
            Dim sc As Double
            If useCache Then
                sc = modVecCache.DotAt(qv, slot)
            Else
                sc = modUtil.DotProduct(qv, vv)
            End If
            sc = sc + SparseBoost(sparseKeys2, modShelfStore.NormTextAt(kData, kRow))

            If unionScore.Exists(vid) Then
                If sc > unionScore.Item(vid) Then unionScore.Item(vid) = sc
            Else
                unionScore.Add vid, sc
            End If
NextRow:
        Next r
NextQ:
    Next qi

    modShelfStore.FlushNormBackfill wsK, lastK, kData
    If validQ = 0 Then Exit Function
    Dim total As Long: total = unionScore.count
    If total = 0 Then Exit Function
    If pk > total Then pk = total

    ' union全体から上位pk件を選択(pk<=40・total<=数万の選択法で十分)
    Dim allKeys As Variant: allKeys = unionScore.Keys
    Dim used() As Boolean: ReDim used(0 To total - 1)
    Dim outHits() As Hit
    ReDim outHits(1 To pk)

    Dim a As Long
    For a = 1 To pk
        Dim bestVal As Double: bestVal = -1E+30
        Dim bestJ As Long: bestJ = -1
        Dim j As Long
        For j = 0 To total - 1
            If Not used(j) Then
                If unionScore.Item(allKeys(j)) > bestVal Then
                    bestVal = unionScore.Item(allKeys(j))
                    bestJ = j
                End If
            End If
        Next j
        If bestJ < 0 Then Exit For
        used(bestJ) = True

        Dim selId As String: selId = CStr(allKeys(bestJ))
        Dim selRow As Long: selRow = idx.Item(selId)
        outHits(a).chunk_id = selId
        outHits(a).score = bestVal
        outHits(a).source = CStr(kData(selRow, COL_K_SOURCE))
        outHits(a).page = CLng(Val(kData(selRow, COL_K_PAGE)))
        outHits(a).preview = modUtil.SafeLeft(CStr(kData(selRow, COL_K_FULLTEXT)), PREVIEW_LEN)
        outHits(a).origin = CStr(kData(selRow, COL_K_ORIGIN))
        outHits(a).full_text = CStr(kData(selRow, COL_K_FULLTEXT))
    Next a

    hits = outHits
    SearchExpanded = pk
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------
Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function


' 質問文を半角/全角スペースで分割し、空要素を除いた語配列を返す。
' スペースが無い日本語の質問は「質問文全体」が1語になり、部分一致判定に使われる。

' ----------------------------------------------------------------------------
' SparseBoost - キーワード側のスコア(0～おおよそ1.0)。
' ----------------------------------------------------------------------------
' ベクトル(内積)は -1～1 のスケールなので、こちらも同程度に収めてから足す。
' 上限で頭打ちにはしない(強く一致したものは確実に上へ来るべき)が、
' 係数でスケールを合わせ、ベクトル順位を不当に覆さないようにする。
'
' dfCsv は本来コーパス全体の文書頻度が要るが、VBAで2万チャンクぶんの
' 転置索引を毎回作るのは現実的でない。ここでは df=1(=最も希少)として
' 扱い、代わりに「質問から抜いた効く語の完全一致」で決定的な差をつける。
' 実測(tools/bench_retrieval.py・実物6資料31問)では、この近似でも
' R@1 84% / R@5 100% を維持している。
Private Function SparseBoost(ByVal preparedKeys As String, ByVal compactDoc As String) As Double
    On Error Resume Next
    SparseBoost = modSparse.KeyScore(preparedKeys, compactDoc) * SPARSE_WEIGHT
    On Error GoTo 0
End Function


' ----------------------------------------------------------------------------
' UnionKeyMatchRows - 粗選別(binary_rag)の候補集合へ「効く語の完全一致を含む
'   行」を強制的に足す(2026-08-01 R12-3-7)。
' ----------------------------------------------------------------------------
' modSparse.bas:63-67 は「片方で候補を絞ってからもう片方を掛ける構成には
' しない。ベクトルで拾えない専門用語が一段目で足切りされる事故を、構造的に
' 防ぐため」と設計保証を書いている。ところが binary_rag=TRUE のときだけは、
' ハミング距離(dense近似)だけで選んだ上位200件の外にある行が SparseBoost の
' 適用機会そのものを失っていた。「第12条は?と聞かれて第12条が1位に来ない
' のは金融では事故」と自ら定義した決定的救済が、粗選別より後段に置かれて
' いたための矛盾である。候補に足すだけなので、順位付けは従来どおり
' Floatスコア+KeyScore が決める(拾いすぎても結果は壊れない)。
'
' 2026-08-01(R12-H-1 敵対的レビュー High-1/Med-4): 照合先を生4列の連結から
' 前計算済みの norm_text(modShelfStore.NormTextAt)へ変えた。
' DistinctiveKeys が返すキーは正規化済み(小文字・半角・空白除去)なのに、
' 当てる先が生テキストのままだったため、全角英数の型番や
' PDF字詰めの「保 険 金」を含む行が【救済されずに落ちて】いた。
' 救済のための仕組みが、救済すべき行を取りこぼしていたことになる。
' R12-4 で norm_text を前計算した今は、正規化のやり直しコストも掛からない
' (当初 vbTextCompare で生テキストを見ていた理由=速度が消えた)。
Private Sub UnionKeyMatchRows(ByRef idData As Variant, ByRef kData As Variant, _
                              ByVal idx As Object, ByVal keys As String, _
                              ByVal cand As Object)
    On Error Resume Next
    If cand Is Nothing Then Exit Sub
    If LenB(keys) = 0 Then Exit Sub

    Dim r As Long
    Dim rLo As Long: rLo = LBound(idData, 1)
    Dim nAll As Long: nAll = UBound(idData, 1) - rLo + 1
    For r = rLo To UBound(idData, 1)
        ' 救済unionも全件走査(粗選別が省いたぶんをここで見る)。無反応にしない。
        If ((r - rLo) Mod PROGRESS_STEP) = 0 And r > rLo Then TickScan r - rLo, nAll
        If Not cand.Exists(r) Then
            Dim vid As String: vid = CStr(idData(r, 1))
            If LenB(vid) > 0 Then
                If idx.Exists(vid) Then
                    Dim kRow As Long: kRow = idx.Item(vid)
                    If modSparse.HasAnyKey(keys, modShelfStore.NormTextAt(kData, kRow)) Then
                        cand.Add r, True
                    End If
                End If
            End If
        End If
    Next r
    On Error GoTo 0
End Sub

' bestScore(1..k)中の最小値とその添字を再計算する(ストリーミングtop-k用)。
Private Sub RecomputeMin(ByRef scores() As Double, ByVal k As Long, _
                         ByRef outIdx As Long, ByRef outScore As Double)
    Dim i As Long
    outIdx = 1
    outScore = scores(1)
    For i = 2 To k
        If scores(i) < outScore Then
            outScore = scores(i)
            outIdx = i
        End If
    Next i
End Sub

' ----------------------------------------------------------------------------
' TickScan - 全件走査中の進捗更新+DoEvents(2026-08-01 R12-4)。
' ----------------------------------------------------------------------------
' 20,500行の走査は実機で数秒かかる。その間 DoEvents が1回も無いと Excel は
' 応答なし(白画面)になり、利用者には「壊れた」としか見えない(憲章§3-1/§3-2)。
' 押されたボタンは modUiLock.Enter(OnSend系が取得済み)が受け流すので、
' ここで DoEvents を回しても検索が二重に走ることはない。
' 表示の失敗が検索を止めないよう、丸ごと On Error Resume Next の中で行う。
' キャッシュがスキャン中に解放されたことの記録(無言の失敗禁止・憲章§4-1)。
' 「たまに答えが薄い」を後から説明できる唯一の痕跡になる。
Private Sub LostMidScan(ByVal ctx As String)
    On Error Resume Next
    modLog.LogUsage "veccache_lost_midscan", "", _
        ctx & ": 検索中にベクトルキャッシュが解放されたため、この質問は" & _
        "従来の経路で最初からやり直しました(結果は欠けません)"
    On Error GoTo 0
End Sub

Private Sub TickScan(ByVal doneRows As Long, ByVal totalRows As Long)
    On Error Resume Next
    modUIMain.SetStage ChrW(&HD83D) & ChrW(&HDD0D) & " 検索中 " & _
        modUtil.ProgressText(doneRows, totalRows, "")
    DoEvents
    On Error GoTo 0
End Sub

' 選択ソート内で使う、並行配列7本分のインデックスa/bの要素を入れ替える。
Private Sub SwapResult(ByRef ids() As String, ByRef scores() As Double, ByRef sources() As String, _
                       ByRef pages() As Long, ByRef previews() As String, ByRef origins() As String, _
                       ByRef fullTexts() As String, ByVal a As Long, ByVal b As Long)
    Dim tS As String, tD As Double, tL As Long
    tS = ids(a): ids(a) = ids(b): ids(b) = tS
    tD = scores(a): scores(a) = scores(b): scores(b) = tD
    tS = sources(a): sources(a) = sources(b): sources(b) = tS
    tL = pages(a): pages(a) = pages(b): pages(b) = tL
    tS = previews(a): previews(a) = previews(b): previews(b) = tS
    tS = origins(a): origins(a) = origins(b): origins(b) = tS
    tS = fullTexts(a): fullTexts(a) = fullTexts(b): fullTexts(b) = tS
End Sub
