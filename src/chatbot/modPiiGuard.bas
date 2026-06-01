Attribute VB_Name = "modPiiGuard"
Option Explicit

' ============================================================================
' modPiiGuard - regex-based detection of likely PII in the user's question
' ----------------------------------------------------------------------------
' Returns True if the text matches any sensitive pattern. The chat UI uses
' this before sending to display a confirmation dialog.
' ============================================================================

Private mRegexes() As Object
Private mLoaded As Boolean

Public Function DetectPii(ByVal text As String) As Boolean
    EnsureRegexes
    Dim i As Long
    For i = LBound(mRegexes) To UBound(mRegexes)
        If mRegexes(i).Test(text) Then
            DetectPii = True
            Exit Function
        End If
    Next i
    DetectPii = False
End Function

Private Sub EnsureRegexes()
    If mLoaded Then Exit Sub
    Dim patterns As Variant
    patterns = Array( _
        "\b0\d{1,4}-\d{1,4}-\d{3,4}\b", _
        "\b0[789]0-\d{4}-\d{4}\b", _
        "[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", _
        "\b\d{4}[- ]?\d{4}[- ]?\d{4}\b", _
        "\b\d{12}\b", _
        "\b契約番号[:：]?\s*\d{6,}\b", _
        "\b証券番号[:：]?\s*\d{6,}\b" _
    )
    ReDim mRegexes(LBound(patterns) To UBound(patterns))
    Dim i As Long
    For i = LBound(patterns) To UBound(patterns)
        Dim re As Object
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = patterns(i)
        re.Global = False
        re.IgnoreCase = True
        Set mRegexes(i) = re
    Next i
    mLoaded = True
End Sub
