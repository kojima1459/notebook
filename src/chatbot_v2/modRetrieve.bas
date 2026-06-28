Attribute VB_Name = "modRetrieve"
Option Explicit

' ============================================================================
' modRetrieve - Phase 2: embedding-based hybrid retrieval (GATED, fail-safe)
' ----------------------------------------------------------------------------
' What this changes in the pipeline, and NOTHING ELSE:
'   Today  : the router LLM is shown ALL visible chunks (e.g. 990 lines) and
'            picks the best N.
'   Phase 2: when rag_enabled = TRUE *and* kb_vectors holds embeddings, we first
'            use cosine similarity to pick the top-K candidates (default 40),
'            and show the router LLM only those K lines. The router then picks
'            the best N out of 40 instead of 990.
'
' Everything downstream of the router (parsing, drafter, verifier, citations)
' is byte-for-byte unchanged, because the router's output format is identical.
'
' SAFETY (the owner dreads breakage — this can never error a query):
'   - rag_enabled defaults FALSE  -> BuildRouterTableHybrid returns the full
'     table; behaviour is 100% identical to before Phase 2.
'   - If embedding the query fails ("" from modEmbeddings) ............ full table
'   - If kb_vectors is missing / empty (never vectorised) ............. full table
'   - If anything raises at all (On Error) ........................... full table
'   So the worst case of Phase 2 is "no speed-up", never "a broken query".
'   The ONLY observable change when everything works is a smaller router table.
'
' Toggle off instantly: set config rag_enabled = FALSE -> reverts to full table.
' ============================================================================

Private Const KB As String = "knowledge_base"
Private Const VEC As String = "kb_vectors"
Private Const KB_COL_ID As Long = 1
Private Const KB_COL_DEPT_SCOPE As Long = 10

' ----------------------------------------------------------------------------
' BuildRouterTableHybrid - the ONLY entry point the pipeline calls.
' Returns the router table string: narrowed (Phase 2) when possible, otherwise
' the full table (identical to modKnowledgeBase.BuildRouterTable).
' Never raises. Returns "" only when the dept genuinely has zero visible chunks
' (same contract as BuildRouterTable, so the existing empty->error path holds).
' ----------------------------------------------------------------------------
Public Function BuildRouterTableHybrid(ByVal qText As String, ByVal deptId As String) As String
    On Error GoTo FullFallback

    ' Gate: feature off -> exactly today's behaviour.
    If Not modConfig.GetBool("rag_enabled", False) Then
        BuildRouterTableHybrid = modKnowledgeBase.BuildRouterTable(deptId)
        Exit Function
    End If

    Dim k As Long: k = modConfig.GetLong("rag_candidates", 40)
    If k < 1 Then k = 40

    Dim ids As String
    ids = TopCandidateIds(qText, deptId, k)     ' "" => embedding/vectors unavailable
    If LenB(ids) = 0 Then GoTo FullFallback

    Dim tbl As String
    tbl = modKnowledgeBase.BuildRouterTableFromIds(ids, deptId)
    If LenB(tbl) = 0 Then GoTo FullFallback     ' nothing resolved -> safety net

    BuildRouterTableHybrid = tbl
    Exit Function

FullFallback:
    ' Any failure path: degrade gracefully to the full LLM-router table.
    On Error Resume Next
    BuildRouterTableHybrid = modKnowledgeBase.BuildRouterTable(deptId)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' TopCandidateIds - embed the query, score every dept-visible vectorised chunk
' by cosine similarity, return the top-k chunk_ids (comma separated, strongest
' first). Returns "" on ANY problem so the caller falls back to the full table.
'
' We embed the already-enriched router question (it includes the intent layer's
' search sub-queries), so retrieval sees every facet the router would have.
' ----------------------------------------------------------------------------
Public Function TopCandidateIds(ByVal qText As String, ByVal deptId As String, ByVal k As Long) As String
    On Error GoTo Fail
    If k < 1 Then k = 40

    ' 1) Embed the query. modEmbeddings returns "" on any failure (incl. ribbon
    '    "error"); mock_llm=TRUE yields a deterministic vector for offline tests.
    Dim qcsv As String: qcsv = modEmbeddings.EmbedText(qText)
    If LenB(qcsv) = 0 Then Exit Function
    Dim qv() As Double
    If Not modEmbeddings.ParseVector(qcsv, qv) Then Exit Function

    ' 2) Build the dept-visible id set (vbLf-delimited for InStr membership).
    '    This both restricts candidates to the user's scope (no cross-dept leak)
    '    and lets us skip scoring chunks the user can't see.
    Dim visible As String: visible = vbLf & VisibleIdSet(deptId) & vbLf
    If LenB(visible) <= 2 Then Exit Function     ' only the two vbLf -> nothing visible

    ' 3) Load kb_vectors (id + vector_csv) in one bulk read.
    Dim wsV As Worksheet
    On Error Resume Next
    Set wsV = ThisWorkbook.Worksheets(VEC)
    On Error GoTo Fail
    If wsV Is Nothing Then Exit Function
    Dim lastV As Long: lastV = wsV.Cells(wsV.Rows.count, 1).End(xlUp).row
    If lastV < 2 Then Exit Function              ' header only -> not vectorised yet

    ' Reading two columns (>=2 cells) always yields a 2-D array, so there is no
    ' single-cell scalar edge case to special-case here.
    Dim data As Variant
    data = wsV.Range(wsV.Cells(2, 1), wsV.Cells(lastV, 2)).value

    ' 4) Streaming top-k by cosine. K is small (40), N up to ~990, so a simple
    '    "track the current minimum, replace if better" beats sorting everything.
    Dim bestId() As String: ReDim bestId(1 To k)
    Dim bestScore() As Double: ReDim bestScore(1 To k)
    Dim filled As Long: filled = 0
    Dim minIdx As Long: minIdx = 0
    Dim minScore As Double: minScore = 0

    Dim r As Long
    For r = LBound(data, 1) To UBound(data, 1)
        Dim id As String: id = CStr(data(r, 1))
        If LenB(id) = 0 Then GoTo NextR
        If InStr(1, visible, vbLf & id & vbLf, vbBinaryCompare) = 0 Then GoTo NextR

        Dim vcsv As String: vcsv = CStr(data(r, 2))
        If LenB(vcsv) = 0 Then GoTo NextR
        Dim vv() As Double
        If Not modEmbeddings.ParseVector(vcsv, vv) Then GoTo NextR

        Dim sc As Double: sc = modEmbeddings.CosineSim(qv, vv)

        If filled < k Then
            filled = filled + 1
            bestId(filled) = id
            bestScore(filled) = sc
            If filled = k Then RecomputeMin bestScore, k, minIdx, minScore
        ElseIf sc > minScore Then
            bestId(minIdx) = id
            bestScore(minIdx) = sc
            RecomputeMin bestScore, k, minIdx, minScore
        End If
NextR:
    Next r

    If filled = 0 Then Exit Function

    ' 5) Sort the (<=k) winners by score descending so the strongest candidate is
    '    listed first for the router. Selection sort is fine at this size.
    Dim a As Long, b As Long, mx As Long
    For a = 1 To filled - 1
        mx = a
        For b = a + 1 To filled
            If bestScore(b) > bestScore(mx) Then mx = b
        Next b
        If mx <> a Then
            Dim ts As Double: ts = bestScore(a): bestScore(a) = bestScore(mx): bestScore(mx) = ts
            Dim ti As String: ti = bestId(a): bestId(a) = bestId(mx): bestId(mx) = ti
        End If
    Next a

    Dim out As String
    For a = 1 To filled
        If LenB(bestId(a)) = 0 Then GoTo NextOut
        If LenB(out) > 0 Then out = out & ","
        out = out & bestId(a)
NextOut:
    Next a
    TopCandidateIds = out
    Exit Function

Fail:
    TopCandidateIds = ""
End Function

' ----------------------------------------------------------------------------
' RecomputeMin - find the smallest score in bestScore(1..k) and its index.
' Called whenever the heap's minimum may have changed.
' ----------------------------------------------------------------------------
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
' VisibleIdSet - vbLf-joined chunk_ids visible to deptId (scope=common or =dept).
' Mirrors modKnowledgeBase.BuildRouterTable's visibility rule exactly.
' Bulk-reads col A (id) and col J (dept_scope); handles the single-data-row case.
' ----------------------------------------------------------------------------
Private Function VisibleIdSet(ByVal deptId As String) As String
    On Error GoTo Done
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(KB)
    On Error GoTo Done
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, KB_COL_ID).End(xlUp).row
    If lastRow < 2 Then Exit Function
    Dim n As Long: n = lastRow - 1

    Dim sb As String

    ' Single data row: a one-cell range returns a scalar, so read directly.
    If n = 1 Then
        Dim id1 As String: id1 = CStr(ws.Cells(2, KB_COL_ID).value)
        Dim sc1 As String: sc1 = CStr(ws.Cells(2, KB_COL_DEPT_SCOPE).value)
        If LenB(sc1) = 0 Then sc1 = "common"
        If LenB(id1) > 0 And (sc1 = "common" Or sc1 = deptId) Then sb = id1 & vbLf
        VisibleIdSet = sb
        Exit Function
    End If

    Dim idCol As Variant: idCol = ws.Range(ws.Cells(2, KB_COL_ID), ws.Cells(lastRow, KB_COL_ID)).value
    Dim scCol As Variant: scCol = ws.Range(ws.Cells(2, KB_COL_DEPT_SCOPE), ws.Cells(lastRow, KB_COL_DEPT_SCOPE)).value

    Dim i As Long
    For i = 1 To n
        Dim id As String: id = CStr(idCol(i, 1))
        If LenB(id) > 0 Then
            Dim scope As String: scope = CStr(scCol(i, 1))
            If LenB(scope) = 0 Then scope = "common"
            If scope = "common" Or scope = deptId Then sb = sb & id & vbLf
        End If
    Next i
    VisibleIdSet = sb
Done:
End Function
