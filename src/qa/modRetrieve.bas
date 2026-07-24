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
'   ・キーワードボーナス: 質問文を空白(半角・全角)で分割した「語」ごとに、
'     summary/keywords/source のいずれかに含まれていれば+0.05、
'     合計上限+0.15(MASTER_SPEC §7.3の記述どおり)。日本語の質問は
'     スペース区切りが無いことが多く、その場合は質問文全体を1語として
'     部分一致判定する(素朴だが単純で決定的な実装を優先)。
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

Private Const PREVIEW_LEN As Long = 120
Private Const KEYWORD_BONUS_PER_WORD As Double = 0.05
Private Const KEYWORD_BONUS_MAX As Double = 0.15
Private Const FULLTEXT_BONUS_PER_WORD As Double = 0.08
Private Const FULLTEXT_BONUS_MAX As Double = 0.25
Private Const EXACT_PHRASE_BONUS As Double = 0.15

' ----------------------------------------------------------------------------
' Search - MASTER_SPEC §7.3 唯一の公開関数。
'   戻り値 = 件数(0可)。埋め込み失敗時は-1(呼び出し側がE0203表示)。
' ----------------------------------------------------------------------------
Public Function Search(ByVal query As String, ByVal topK As Long, ByRef hits() As Hit) As Long
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

    ' 2列以上(>=2セル)の読込みは常に2次元配列になるため単一行の特別扱いは不要
    ' (§4データモデル: 列数はmy_vectors=2, my_knowledgeは1〜7列目を読む=7列)。
    Dim vData As Variant
    vData = wsV.Range(wsV.Cells(2, COL_V_ID), wsV.Cells(lastV, COL_V_VEC)).Value

    Dim kData As Variant
    kData = wsK.Range(wsK.Cells(2, COL_K_ID), wsK.Cells(lastK, COL_K_FULLTEXT)).Value

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

    Dim words() As String
    words = TokenizeQuery(query)

    ' ナレッジ自浄: ノイズ報告が閾値以上の資料を検索対象から論理除外する
    Dim excl As Object: Set excl = modStats.ExcludedSources()

    ' [狂気案Lv.1] バイナリ量子化ハイブリッド: 大規模KB(binary_rag_min以上)かつ
    ' binary_rag=TRUEのときだけ、ハミング距離で候補行を粗選別する。以降のFloat
    ' コサイン/キーワードボーナス/ストリーミングtop-kは一切不変(候補以外をスキップ
    ' する1行のガードを足すだけ)。Prefilterが辞退(次元不一致等)したら全件Floatへ。
    Dim useCand As Boolean: useCand = False
    Dim candRows As Object
    Dim binMs As Double: binMs = 0
    If modBitwiseOpt.Enabled(UBound(vData, 1) - LBound(vData, 1) + 1) Then
        Dim tB0 As Double: tB0 = modBitwiseOpt.MicroTimerMs()
        useCand = modBitwiseOpt.Prefilter(qv, vData, modBitwiseOpt.PrefilterN(), candRows)
        binMs = modBitwiseOpt.MicroTimerMs() - tB0   ' 爆速証明: バイナリ粗選別のms
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
    For r = LBound(vData, 1) To UBound(vData, 1)
        If useCand Then If Not candRows.Exists(r) Then GoTo NextR   ' [Lv.1] 粗選別候補以外は除外
        Dim vid As String
        vid = CStr(vData(r, COL_V_ID))
        If LenB(vid) = 0 Then GoTo NextR
        If Not idx.Exists(vid) Then GoTo NextR

        Dim vcsv As String
        vcsv = CStr(vData(r, COL_V_VEC))
        If LenB(vcsv) = 0 Then GoTo NextR

        Dim vv() As Double
        If Not modUtil.CsvToVector(vcsv, vv) Then GoTo NextR

        ' Wave4修正: modUtil.DotProductは次元不一致を「呼び出し側が別途次元検査を
        ' 行う規約」(modUtil.bas冒頭コメント)としているが、以前はここで一切
        ' 検査せず0点のままランキングに混ぜていた。mock_llm切替等でembed_dimが
        ' 変わった保存済みベクトルが残っていると、全チャンクが無言で0点になり
        ' err_log/診断のどちらにも痕跡が残らない不具合があった(E0702は元々
        ' パック検証専用だったが、検索時の次元不一致もE0702として記録する)。
        If UBound(vv) - LBound(vv) + 1 <> qDim Then
            If Not e0702Logged Then
                modLog.LogError "E0702", "modRetrieve.Search", _
                    "chunk_id=" & vid & " queryDim=" & qDim & " vectorDim=" & (UBound(vv) - LBound(vv) + 1)
                e0702Logged = True
            End If
            GoTo NextR
        End If

        Dim kRow As Long
        kRow = idx.Item(vid)
        Dim srcName As String: srcName = CStr(kData(kRow, COL_K_SOURCE))
        If excl.Exists(srcName) Then GoTo NextR   ' ノイズ報告で論理除外された資料
        Dim origin As String: origin = CStr(kData(kRow, COL_K_ORIGIN))
        Dim pageNum As Long: pageNum = CLng(Val(kData(kRow, COL_K_PAGE)))
        Dim summary As String: summary = CStr(kData(kRow, COL_K_SUMMARY))
        Dim keywords As String: keywords = CStr(kData(kRow, COL_K_KEYWORDS))
        Dim fullText As String: fullText = CStr(kData(kRow, COL_K_FULLTEXT))

        Dim sc As Double
        sc = modUtil.DotProduct(qv, vv)
        sc = sc + KeywordBonus(words, summary, keywords, srcName)
        sc = sc + FullTextBoost(words, fullText, query)

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
        modBitwiseOpt.LogPerf binMs, modBitwiseOpt.MicroTimerMs() - tF0, _
            candRows.count, UBound(vData, 1) - LBound(vData, 1) + 1
    End If

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
' ----------------------------------------------------------------------------
Public Function SearchExpanded(queries() As String, ByVal poolK As Long, ByRef hits() As Hit) As Long
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

    Dim vData As Variant
    vData = wsV.Range(wsV.Cells(2, COL_V_ID), wsV.Cells(lastV, COL_V_VEC)).Value
    Dim kData As Variant
    kData = wsK.Range(wsK.Cells(2, COL_K_ID), wsK.Cells(lastK, COL_K_FULLTEXT)).Value

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

        Dim words() As String
        words = TokenizeQuery(qText)
        Dim grams() As String
        grams = BigramsIfSingleToken(qText, words)

        Dim r As Long
        For r = LBound(vData, 1) To UBound(vData, 1)
            Dim vid As String
            vid = CStr(vData(r, COL_V_ID))
            If LenB(vid) = 0 Then GoTo NextRow
            If Not idx.Exists(vid) Then GoTo NextRow
            Dim vcsv As String
            vcsv = CStr(vData(r, COL_V_VEC))
            If LenB(vcsv) = 0 Then GoTo NextRow
            Dim vv() As Double
            If Not modUtil.CsvToVector(vcsv, vv) Then GoTo NextRow
            If UBound(vv) - LBound(vv) + 1 <> qDim Then
                If Not e0702Logged Then
                    modLog.LogError "E0702", "modRetrieve.SearchExpanded", _
                        "chunk_id=" & vid & " queryDim=" & qDim & " vectorDim=" & (UBound(vv) - LBound(vv) + 1)
                    e0702Logged = True
                End If
                GoTo NextRow
            End If

            Dim kRow As Long: kRow = idx.Item(vid)
            If excl.Exists(CStr(kData(kRow, COL_K_SOURCE))) Then GoTo NextRow   ' ノイズ論理除外
            Dim sc As Double
            sc = modUtil.DotProduct(qv, vv)
            sc = sc + KeywordBonus(words, CStr(kData(kRow, COL_K_SUMMARY)), _
                                   CStr(kData(kRow, COL_K_KEYWORDS)), CStr(kData(kRow, COL_K_SOURCE)))
            sc = sc + GramBonus(grams, CStr(kData(kRow, COL_K_SUMMARY)) & " " & _
                                CStr(kData(kRow, COL_K_KEYWORDS)) & " " & _
                                CStr(kData(kRow, COL_K_SOURCE)) & " " & _
                                modUtil.SafeLeft(CStr(kData(kRow, COL_K_FULLTEXT)), 200))
            sc = sc + FullTextBoost(words, CStr(kData(kRow, COL_K_FULLTEXT)), queries(LBound(queries)))

            If unionScore.Exists(vid) Then
                If sc > unionScore.Item(vid) Then unionScore.Item(vid) = sc
            Else
                unionScore.Add vid, sc
            End If
NextRow:
        Next r
NextQ:
    Next qi

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

' 日本語向け補助トークン: 分割語が1語だけ(=スペース無しの日本語質問)で
' 6文字以上のとき、文字2-gram(最大20個)を作る(設計書§C-2の改良)。
' 対象外のときは0要素配列(Split(vbNullString))。
Private Function BigramsIfSingleToken(ByVal qText As String, ByRef words() As String) As String()
    Dim wc As Long: wc = 0
    On Error Resume Next
    wc = UBound(words) - LBound(words) + 1
    Err.Clear
    On Error GoTo 0

    Dim src As String: src = Replace(Replace(Trim$(qText), " ", ""), "　", "")
    If wc > 1 Or Len(src) < 6 Then
        BigramsIfSingleToken = Split(vbNullString)
        Exit Function
    End If

    Dim maxG As Long: maxG = Len(src) - 1
    If maxG > 20 Then maxG = 20
    Dim outArr() As String: ReDim outArr(0 To maxG - 1)
    Dim i As Long
    For i = 1 To maxG
        outArr(i - 1) = Mid$(src, i, 2)
    Next i
    BigramsIfSingleToken = outArr
End Function

' 2-gram補助ボーナス: 1個一致につき+0.02、上限+0.10(SearchExpanded専用)。
Private Function GramBonus(ByRef grams() As String, ByVal haystack As String) As Double
    Dim n As Long: n = 0
    On Error Resume Next
    n = UBound(grams) - LBound(grams) + 1
    Err.Clear
    On Error GoTo 0
    If n <= 0 Then Exit Function

    Dim matched As Long: matched = 0
    Dim i As Long
    For i = LBound(grams) To UBound(grams)
        If LenB(grams(i)) > 0 Then
            If InStr(1, haystack, grams(i), vbTextCompare) > 0 Then matched = matched + 1
        End If
    Next i

    Dim bonus As Double
    bonus = CDbl(matched) * 0.02
    If bonus > 0.1 Then bonus = 0.1
    GramBonus = bonus
End Function

' 質問文を半角/全角スペースで分割し、空要素を除いた語配列を返す。
' スペースが無い日本語の質問は「質問文全体」が1語になり、部分一致判定に使われる。
Private Function TokenizeQuery(ByVal query As String) As String()
    Dim q As String
    q = Replace(query, "　", " ")
    TokenizeQuery = modUtil.SplitKeepNonEmpty(q, " ")
End Function

' summary/keywords/source のいずれかに語が含まれていれば+0.05、上限+0.15。
Private Function KeywordBonus(ByRef words() As String, ByVal summary As String, _
                              ByVal keywords As String, ByVal srcName As String) As Double
    Dim haystack As String
    haystack = summary & " " & keywords & " " & srcName

    Dim matched As Long: matched = 0
    Dim i As Long
    For i = LBound(words) To UBound(words)
        If LenB(words(i)) > 0 Then
            If InStr(1, haystack, words(i), vbTextCompare) > 0 Then
                matched = matched + 1
            End If
        End If
    Next i

    Dim bonus As Double
    bonus = CDbl(matched) * KEYWORD_BONUS_PER_WORD
    If bonus > KEYWORD_BONUS_MAX Then bonus = KEYWORD_BONUS_MAX
    KeywordBonus = bonus
End Function

' FullTextBoost - チャンク本文(full_text)内の語彙一致による追加ブースト。
' 既存のKeywordBonus(summary/keywords/source対象)を補完し、本文に直接
' 質問語が含まれるチャンクを優遇する。完全一致フレーズボーナス付き。
Private Function FullTextBoost(ByRef words() As String, ByVal fullText As String, _
                               ByVal rawQuery As String) As Double
    If LenB(fullText) = 0 Then
        FullTextBoost = 0#
        Exit Function
    End If

    ' 完全一致フレーズボーナス: 質問文全体が本文に含まれていれば大幅加算
    Dim bonus As Double: bonus = 0#
    If LenB(rawQuery) >= 4 Then
        If InStr(1, fullText, rawQuery, vbTextCompare) > 0 Then
            bonus = bonus + EXACT_PHRASE_BONUS
        End If
    End If

    ' 語彙単位ボーナス: 各語が本文に出現すれば加算(TF考慮: 2回以上で+50%)
    Dim matched As Long: matched = 0
    Dim i As Long
    For i = LBound(words) To UBound(words)
        If LenB(words(i)) >= 2 Then
            Dim pos As Long: pos = InStr(1, fullText, words(i), vbTextCompare)
            If pos > 0 Then
                matched = matched + 1
                ' 2回目出現チェック(TF: 頻出語はより関連度が高い)
                If InStr(pos + Len(words(i)), fullText, words(i), vbTextCompare) > 0 Then
                    matched = matched + 1  ' 0.5回分追加(整数で近似)
                End If
            End If
        End If
    Next i

    Dim wordBonus As Double
    wordBonus = CDbl(matched) * FULLTEXT_BONUS_PER_WORD * 0.667  ' 1.5倍を2/3で近似
    If wordBonus > FULLTEXT_BONUS_MAX Then wordBonus = FULLTEXT_BONUS_MAX
    bonus = bonus + wordBonus

    FullTextBoost = bonus
End Function

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
