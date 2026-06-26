Attribute VB_Name = "modFeedbackLookup"
Option Explicit

' ============================================================================
' modFeedbackLookup - Step 2 of the pipeline: find similar past Q&A
' ----------------------------------------------------------------------------
' Sheet `feedback` columns:
'   A: fb_id
'   B: timestamp
'   C: dept_id
'   D: submitter
'   E: question
'   F: answer            (the final answer that was approved)
'   G: status            (pending / approved / rejected)
'   H: approver
'   I: approved_at
'   J: correction        (if rejected, the corrected version)
'   K: tags
'
' Lookup logic for MVP:
'   - Only consider rows where:
'       (status == "approved")  OR
'       (status == "pending" AND submitter == current_user)
'     because pending entries are valid for the submitter immediately.
'   - Score similarity by keyword/character overlap (no embedding) +
'     dept match boost.
'   - Return top-N (config: feedback_top_n, default 3).
'
' Effectively this is "in-context learning via a curated FAQ". The system
' gets smarter as users approve/reject answers.
' ============================================================================

Private Const SHEET_NAME As String = "feedback"

' Returns a single text block to append into the drafter prompt, or "" if none
Public Function FindSimilarApproved(ByVal question As String, _
                                    ByVal deptId As String, _
                                    ByVal currentUser As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Function

    Dim topN As Long: topN = modConfig.GetLong("feedback_top_n", 3)
    Dim minScore As Double: minScore = modConfig.GetDouble("feedback_min_score", 0.2)

    Dim qKeywords As String: qKeywords = ExtractTokens(question)

    ' Score each candidate row
    Dim scores() As Double, rows() As Long, n As Long
    ReDim scores(1 To 1000)
    ReDim rows(1 To 1000)

    Dim i As Long, c As Long
    For i = 2 To lastRow
        Dim status As String: status = LCase$(CStr(ws.Cells(i, 7).value))
        Dim submitter As String: submitter = CStr(ws.Cells(i, 4).value)
        Dim isUsable As Boolean
        isUsable = (status = "approved") Or _
                   (status = "pending" And StrComp(submitter, currentUser, vbTextCompare) = 0)
        If Not isUsable Then GoTo NextRow

        Dim pastQ As String: pastQ = CStr(ws.Cells(i, 5).value)
        Dim pastDept As String: pastDept = CStr(ws.Cells(i, 3).value)

        Dim s As Double
        s = SimpleScore(qKeywords, ExtractTokens(pastQ))
        ' Boost for same dept
        If StrComp(pastDept, deptId, vbTextCompare) = 0 Then s = s + 0.1
        ' Boost for approved over pending
        If status = "approved" Then s = s + 0.05

        If s >= minScore Then
            c = c + 1
            scores(c) = s
            rows(c) = i
        End If
NextRow:
    Next i

    If c = 0 Then Exit Function

    ' Sort top N by score (simple insertion sort, c is usually small)
    Dim selRows() As Long: ReDim selRows(1 To topN)
    Dim selScores() As Double: ReDim selScores(1 To topN)
    Dim sel As Long
    For i = 1 To c
        If sel < topN Then
            sel = sel + 1
            selRows(sel) = rows(i)
            selScores(sel) = scores(i)
        Else
            ' find lowest among selected and replace if score higher
            Dim worstIdx As Long, worstScore As Double
            worstIdx = 1: worstScore = selScores(1)
            Dim k As Long
            For k = 2 To topN
                If selScores(k) < worstScore Then worstIdx = k: worstScore = selScores(k)
            Next k
            If scores(i) > worstScore Then
                selRows(worstIdx) = rows(i)
                selScores(worstIdx) = scores(i)
            End If
        End If
    Next i

    ' Build output block
    Dim sb As String
    sb = "## 過去の認定済みQ&A (参考)"
    For i = 1 To sel
        Dim rr As Long: rr = selRows(i)
        Dim pastQ2 As String: pastQ2 = CStr(ws.Cells(rr, 5).value)
        Dim pastA As String: pastA = CStr(ws.Cells(rr, 6).value)
        Dim correction As String: correction = CStr(ws.Cells(rr, 10).value)
        ' If a correction exists, use it instead
        Dim useA As String: useA = pastA
        If LenB(correction) > 0 Then useA = correction
        sb = sb & vbLf & vbLf & _
             "### 質問: " & pastQ2 & vbLf & _
             "### 回答: " & useA
    Next i
    FindSimilarApproved = sb
End Function

' ----------------------------------------------------------------------------
' Crude tokenizer: extracts character n-grams + alphanumeric words. Works for
' Japanese without morphological analysis. Returns space-separated tokens.
' ----------------------------------------------------------------------------
Private Function ExtractTokens(ByVal s As String) As String
    Dim t As String: t = LCase$(s)
    ' Replace punctuation/whitespace with space
    Dim ch As String, out As String
    Dim i As Long
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch Like "[A-Za-z0-9]" Or ch >= ChrW(&H3040) Then
            out = out & ch
        Else
            out = out & " "
        End If
    Next i
    ExtractTokens = out
End Function

' ----------------------------------------------------------------------------
' Score similarity by character-pair overlap, normalized.
' Range 0.0 (none) .. 1.0 (identical).
'
' Pure VBA implementation (no Scripting.Dictionary) so it runs identically on
' Windows (including hardened corporate PCs where scrrun.dll may be blocked)
' and on Mac Excel. Bigrams are stored as a vbLf-delimited string of unique
' pairs; membership is tested with InStr.
' ----------------------------------------------------------------------------
Private Function SimpleScore(ByVal a As String, ByVal b As String) As Double
    If LenB(a) = 0 Or LenB(b) = 0 Then Exit Function
    Dim setA As String: setA = UniqueBigrams(a)
    Dim setB As String: setB = UniqueBigrams(b)
    Dim cntA As Long: cntA = BigramCount(setA)
    Dim cntB As Long: cntB = BigramCount(setB)
    If cntA = 0 Or cntB = 0 Then Exit Function

    ' Count bigrams of A that also appear in B
    Dim sharedCnt As Long
    Dim parts() As String: parts = Split(setA, vbLf)
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If LenB(parts(i)) > 0 Then
            If InStr(1, setB, vbLf & parts(i) & vbLf, vbBinaryCompare) > 0 Then
                sharedCnt = sharedCnt + 1
            End If
        End If
    Next i

    Dim denom As Double: denom = (cntA + cntB) / 2
    SimpleScore = sharedCnt / denom
End Function

' Returns a string of unique bigrams, each wrapped so it reads as
' vbLf & bigram & vbLf ... vbLf (sentinels at both ends for exact InStr match).
Private Function UniqueBigrams(ByVal s As String) As String
    Dim out As String: out = vbLf
    Dim i As Long
    For i = 1 To Len(s) - 1
        Dim bg As String: bg = Mid$(s, i, 2)
        If LenB(Trim$(bg)) >= 2 Then
            If InStr(1, out, vbLf & bg & vbLf, vbBinaryCompare) = 0 Then
                out = out & bg & vbLf
            End If
        End If
    Next i
    UniqueBigrams = out
End Function

Private Function BigramCount(ByVal setStr As String) As Long
    ' setStr is "\n bg1 \n bg2 \n ... \n" -> count of separators minus 1
    Dim n As Long: n = Len(setStr) - Len(Replace(setStr, vbLf, ""))
    If n > 0 Then BigramCount = n - 1
End Function
