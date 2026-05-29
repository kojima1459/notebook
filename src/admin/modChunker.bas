Attribute VB_Name = "modChunker"
Option Explicit

' ============================================================================
' modChunker - Split text into overlapping chunks for embedding
' ----------------------------------------------------------------------------
' Strategy: paragraph-aware sliding window targeting ~700 chars per chunk,
' with ~150 char overlap, never splitting mid-sentence when avoidable.
'
' The Chunk type lives in shared\modTypes.bas so chatbot and admin share it.
' ============================================================================

Public Function ChunkText(ByVal source As String, _
                          ByVal page As Long, _
                          ByVal text As String, _
                          Optional ByVal targetLen As Long = 700, _
                          Optional ByVal overlap As Long = 150) As Chunk()
    Dim result() As Chunk
    Dim outCount As Long
    ReDim result(0 To 0)

    Dim normalized As String
    normalized = NormalizeWhitespace(text)
    If Len(normalized) = 0 Then
        ReDim result(-1 To -1)
        ChunkText = result
        Exit Function
    End If

    Dim start As Long: start = 1
    Dim total As Long: total = Len(normalized)
    Dim idx As Long: idx = 0

    Do While start <= total
        Dim windowEnd As Long
        windowEnd = start + targetLen - 1
        If windowEnd > total Then windowEnd = total

        ' Try to extend to next sentence boundary up to +200 chars
        If windowEnd < total Then
            Dim boundary As Long
            boundary = FindSentenceBoundary(normalized, windowEnd, targetLen \ 4)
            If boundary > 0 Then windowEnd = boundary
        End If

        Dim ck As Chunk
        ck.Id = source & "#" & CStr(idx)
        ck.Source = source
        ck.page = page
        ck.StartCharOffset = start - 1
        ck.Text = Mid$(normalized, start, windowEnd - start + 1)

        If outCount > UBound(result) Then ReDim Preserve result(0 To outCount * 2 + 1)
        result(outCount) = ck
        outCount = outCount + 1
        idx = idx + 1

        If windowEnd >= total Then Exit Do
        start = windowEnd - overlap + 1
        If start < 1 Then start = 1
    Loop

    ReDim Preserve result(0 To outCount - 1)
    ChunkText = result
End Function

Private Function NormalizeWhitespace(ByVal s As String) As String
    Dim t As String
    t = Replace(s, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    ' Collapse runs of spaces / tabs but preserve paragraph breaks
    Dim out As String, i As Long, ch As String, prevWs As Boolean
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch = " " Or ch = vbTab Then
            If Not prevWs Then out = out & " "
            prevWs = True
        Else
            out = out & ch
            prevWs = False
        End If
    Next i
    ' Collapse 3+ newlines to 2
    Do While InStr(out, vbLf & vbLf & vbLf) > 0
        out = Replace(out, vbLf & vbLf & vbLf, vbLf & vbLf)
    Loop
    NormalizeWhitespace = Trim$(out)
End Function

Private Function FindSentenceBoundary(ByVal s As String, ByVal nearPos As Long, _
                                      ByVal slack As Long) As Long
    Dim limit As Long: limit = nearPos + slack
    If limit > Len(s) Then limit = Len(s)
    Dim i As Long
    For i = nearPos To limit
        Dim ch As String: ch = Mid$(s, i, 1)
        If ch = "." Or ch = "!" Or ch = "?" Or ch = ChrW(&H3002) Or ch = ChrW(&HFF01) Or ch = ChrW(&HFF1F) Or ch = vbLf Then
            FindSentenceBoundary = i
            Exit Function
        End If
    Next i
    FindSentenceBoundary = 0
End Function

' ----------------------------------------------------------------------------
' L2 normalize a flat vector array in-place. Vectors are stored back-to-back.
' Pre-normalizing lets the chatbot use plain dot-product for cosine.
' ----------------------------------------------------------------------------
Public Sub L2NormalizeInPlace(ByRef flat() As Double, ByVal dim_ As Long, ByVal count As Long)
    Dim i As Long, j As Long, base As Long
    Dim sumsq As Double, inv As Double
    For i = 0 To count - 1
        base = i * dim_
        sumsq = 0#
        For j = 0 To dim_ - 1
            sumsq = sumsq + flat(base + j) * flat(base + j)
        Next j
        If sumsq > 0 Then
            inv = 1# / Sqr(sumsq)
            For j = 0 To dim_ - 1
                flat(base + j) = flat(base + j) * inv
            Next j
        End If
    Next i
End Sub
