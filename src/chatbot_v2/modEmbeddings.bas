Attribute VB_Name = "modEmbeddings"
Option Explicit

' ============================================================================
' modEmbeddings - safe gateway to the corporate ribbon's GetEmbeddings()
' ----------------------------------------------------------------------------
' Confirmed ribbon behaviour (excel-addin リボンちゃん ver202606, Embeddings.bas):
'   Function GetEmbeddings(text As String) As String
'     success -> comma-joined floats "0.013,-0.02,..." (1536 dims, unit-normalised)
'     failure -> the literal string "error" (HTTP status <> 200). It never raises;
'                it returns "error" as a normal String.
'   Model: text-embedding-3-small (Azure deployment). Vectors are L2-normalised,
'   so cosine similarity == dot product (used in Phase 2).
'
' This module is the ONLY place that calls GetEmbeddings, and it is defensive by
' design so a failure can never crash a query:
'   - empty input is never sent to the API
'   - input is capped (token-limit safety; all 990 chunks are < 3623 chars anyway)
'   - the "error" sentinel and blank/short replies are detected and turned into ""
'   - the whole call is wrapped in On Error -> returns "" on anything unexpected
'   Callers treat "" as "embedding unavailable" and fall back to the LLM router.
'
' Mock mode (config mock_llm = TRUE) returns a deterministic pseudo-vector so the
' entire vectorize/RAG flow can be exercised offline (Mac / pre-ribbon smoke test).
' ============================================================================

Public Const EMB_DIM As Long = 1536            ' text-embedding-3-small dimensions
Private Const EMB_MAX_CHARS As Long = 4000     ' headroom under the ~8k-token limit

' ----------------------------------------------------------------------------
' EmbedText - return a comma-joined vector string, or "" on ANY failure.
' Never raises. Callers MUST handle "" (fall back to the LLM router).
' ----------------------------------------------------------------------------
Public Function EmbedText(ByVal text As String) As String
    On Error GoTo Fail

    Dim t As String: t = Trim$(text)
    If LenB(t) = 0 Then Exit Function                 ' never call the API empty
    If Len(t) > EMB_MAX_CHARS Then t = Left$(t, EMB_MAX_CHARS)

    ' Offline / Mac / pre-ribbon: deterministic fake vector so the flow still runs.
    If modConfig.GetBool("mock_llm", False) Then
        EmbedText = MockVector(t)
        Exit Function
    End If

    Dim raw As Variant
    raw = Application.Run("GetEmbeddings", t)          ' resolves to the ribbon addin
    Dim s As String: s = CStr(raw)

    If LenB(s) = 0 Then Exit Function                  ' blank reply -> unavailable
    If StrComp(Trim$(s), "error", vbTextCompare) = 0 Then Exit Function  ' ribbon sentinel
    If InStr(s, ",") = 0 Then Exit Function            ' a real vector has many commas

    EmbedText = s
    Exit Function

Fail:
    EmbedText = ""        ' swallow everything: caller will fall back to LLM router
End Function

' ----------------------------------------------------------------------------
' EmbedSelfTest - used by the admin "接続テスト" button. Embeds one short sample
' and reports, in plain Japanese, whether the ribbon embedding works HERE.
' Returns True on success and fills `detail` with a human-readable summary.
' ----------------------------------------------------------------------------
Public Function EmbedSelfTest(ByVal sampleText As String, ByRef detail As String) As Boolean
    On Error GoTo Fail
    Dim t0 As Double: t0 = Timer
    Dim s As String: s = EmbedText(sampleText)
    Dim ms As Long: ms = CLng((Timer - t0) * 1000)

    If LenB(s) = 0 Then
        detail = "失敗 ❌  GetEmbeddings が空 または error を返しました。" & vbLf & vbLf & _
                 "確認してください:" & vbLf & _
                 "  ・社内AIリボン(リボンちゃん)が読み込まれているか" & vbLf & _
                 "  ・社内ネットワークに接続されているか" & vbLf & _
                 "  ・config の mock_llm が FALSE になっているか(本番) / TRUE(オフライン確認)"
        EmbedSelfTest = False
        Exit Function
    End If

    Dim n As Long: n = CountValues(s)
    detail = "成功 ✅  (" & ms & " ミリ秒)" & vbLf & _
             "  次元数: " & n & "  (期待値 " & EMB_DIM & ")" & vbLf & _
             "  先頭: " & Left$(s, 70) & " ..." & vbLf & vbLf & _
             IIf(n >= 1000, "問題ありません。②全件ベクトル化に進めます。", _
                            "※次元数が想定より少ないです。リボンのモデル設定を確認してください。")
    EmbedSelfTest = (n >= 1000)
    Exit Function

Fail:
    detail = "例外が発生しました: " & Err.Description
    EmbedSelfTest = False
End Function

' ----------------------------------------------------------------------------
' CountValues - number of comma-separated elements in a vector string.
' ----------------------------------------------------------------------------
Public Function CountValues(ByVal csv As String) As Long
    If LenB(csv) = 0 Then Exit Function
    CountValues = Len(csv) - Len(Replace(csv, ",", "")) + 1
End Function

' ----------------------------------------------------------------------------
' ParseVector - turn a comma-joined vector string into a Double() (0-based).
' Used by Phase-2 retrieval. Returns False if the string is unusable.
' ----------------------------------------------------------------------------
Public Function ParseVector(ByVal csv As String, ByRef outVec() As Double) As Boolean
    On Error GoTo Fail
    If LenB(csv) = 0 Then Exit Function
    Dim parts() As String: parts = Split(csv, ",")
    Dim n As Long: n = UBound(parts) - LBound(parts) + 1
    If n < 1 Then Exit Function
    ReDim outVec(0 To n - 1)
    Dim i As Long, j As Long: j = 0
    For i = LBound(parts) To UBound(parts)
        outVec(j) = Val(parts(i))      ' Val is locale-safe ("." decimal) & ignores spaces
        j = j + 1
    Next i
    ParseVector = True
    Exit Function
Fail:
    ParseVector = False
End Function

' ----------------------------------------------------------------------------
' CosineSim - cosine similarity of two pre-parsed Double() vectors.
' text-embedding-3 vectors are unit-normalised, so this equals their dot product;
' we still divide by norms for safety (handles mock / non-normalised input).
' Used by Phase-2 retrieval. Returns 0 on length mismatch or zero vector.
' ----------------------------------------------------------------------------
Public Function CosineSim(ByRef a() As Double, ByRef b() As Double) As Double
    On Error GoTo Fail
    Dim la As Long, lb As Long
    la = UBound(a) - LBound(a) + 1
    lb = UBound(b) - LBound(b) + 1
    If la <> lb Or la = 0 Then Exit Function

    Dim dot As Double, na As Double, nb As Double
    Dim i As Long, oa As Long, ob As Long
    oa = LBound(a): ob = LBound(b)
    For i = 0 To la - 1
        Dim x As Double, y As Double
        x = a(oa + i): y = b(ob + i)
        dot = dot + x * y
        na = na + x * x
        nb = nb + y * y
    Next i
    If na = 0 Or nb = 0 Then Exit Function
    CosineSim = dot / (Sqr(na) * Sqr(nb))
    Exit Function
Fail:
    CosineSim = 0
End Function

' ----------------------------------------------------------------------------
' MockVector - deterministic pseudo-vector for offline testing (mock_llm=TRUE).
' Same text -> same vector. Arithmetic kept small to avoid Long overflow.
' Quality is irrelevant; this only proves the plumbing end-to-end.
' ----------------------------------------------------------------------------
Private Function MockVector(ByVal t As String) As String
    Dim seed As Long, i As Long
    For i = 1 To Len(t)
        seed = (seed * 31 + AscW(Mid$(t, i, 1))) Mod 1000003   ' < 2^31, no overflow
    Next i
    Dim parts() As String: ReDim parts(1 To EMB_DIM)
    For i = 1 To EMB_DIM
        seed = (seed * 31 + i) Mod 1000003
        Dim v As Double: v = ((seed Mod 2000) - 1000) / 1000#   ' in [-1.000, 0.999]
        parts(i) = Format$(v, "0.###")
    Next i
    MockVector = Join(parts, ",")
End Function
