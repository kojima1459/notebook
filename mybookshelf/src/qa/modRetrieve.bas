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

    Dim r As Long
    For r = LBound(vData, 1) To UBound(vData, 1)
        Dim vid As String
        vid = CStr(vData(r, COL_V_ID))
        If LenB(vid) = 0 Then GoTo NextR
        If Not idx.Exists(vid) Then GoTo NextR

        Dim vcsv As String
        vcsv = CStr(vData(r, COL_V_VEC))
        If LenB(vcsv) = 0 Then GoTo NextR

        Dim vv() As Double
        If Not modUtil.CsvToVector(vcsv, vv) Then GoTo NextR

        Dim kRow As Long
        kRow = idx.Item(vid)
        Dim srcName As String: srcName = CStr(kData(kRow, COL_K_SOURCE))
        Dim origin As String: origin = CStr(kData(kRow, COL_K_ORIGIN))
        Dim pageNum As Long: pageNum = CLng(Val(kData(kRow, COL_K_PAGE)))
        Dim summary As String: summary = CStr(kData(kRow, COL_K_SUMMARY))
        Dim keywords As String: keywords = CStr(kData(kRow, COL_K_KEYWORDS))
        Dim fullText As String: fullText = CStr(kData(kRow, COL_K_FULLTEXT))

        Dim sc As Double
        sc = modUtil.DotProduct(qv, vv)
        sc = sc + KeywordBonus(words, summary, keywords, srcName)

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
' 内部ヘルパー
' ----------------------------------------------------------------------------
Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
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
