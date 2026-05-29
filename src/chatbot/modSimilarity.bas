Attribute VB_Name = "modSimilarity"
Option Explicit

' ============================================================================
' modSimilarity - Top-K by dot product over L2-normalized vectors
' ----------------------------------------------------------------------------
' Tight loops only: typed Doubles, no Variants, no function calls inside.
' ============================================================================

Public Sub TopK(ByRef flat() As Double, _
                ByVal dim_ As Long, _
                ByVal count As Long, _
                ByRef query() As Double, _
                ByVal k As Long, _
                ByRef outIdx() As Long, _
                ByRef outScore() As Double)
    ' Empty index guard - VBA ReDim(0 To -1) errors, so produce an empty
    ' (-1 To -1) array which the chatbot's downstream code treats as "no hits".
    If count <= 0 Or k <= 0 Then
        ReDim outIdx(-1 To -1)
        ReDim outScore(-1 To -1)
        Exit Sub
    End If
    If k > count Then k = count
    ReDim outIdx(0 To k - 1)
    ReDim outScore(0 To k - 1)
    Dim i As Long
    For i = 0 To k - 1
        outIdx(i) = -1
        outScore(i) = -1E+30
    Next i

    Dim base As Long, j As Long
    Dim s As Double
    For i = 0 To count - 1
        base = i * dim_
        s = 0#
        For j = 0 To dim_ - 1
            s = s + flat(base + j) * query(j)
        Next j
        If s > outScore(k - 1) Then
            Dim pos As Long
            pos = k - 1
            Do While pos > 0
                If outScore(pos - 1) < s Then
                    outScore(pos) = outScore(pos - 1)
                    outIdx(pos) = outIdx(pos - 1)
                    pos = pos - 1
                Else
                    Exit Do
                End If
            Loop
            outScore(pos) = s
            outIdx(pos) = i
        End If
    Next i
End Sub
