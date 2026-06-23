Attribute VB_Name = "modPii"
Option Explicit

' ============================================================================
' modPii - quick local PII detection for warning users before send
' ----------------------------------------------------------------------------
' Catches obvious patterns in the question text:
'   - Phone (XXX-XXXX-XXXX, 03-XXXX-XXXX, +81-XXX-...)
'   - Email
'   - マイナンバー (12 digits, sometimes with spaces)
'   - Contract numbers (heuristic: long digit runs)
' Returns True if any match. UI prompts user; this is advisory, not blocking.
' ============================================================================

Public Function DetectPii(ByVal text As String) As Boolean
    Dim t As String: t = text
    If LenB(t) = 0 Then Exit Function

    ' Email
    If InStr(t, "@") > 0 And t Like "*[A-Za-z0-9.]@[A-Za-z0-9.]*" Then
        DetectPii = True: Exit Function
    End If

    ' Long digit runs (10+ consecutive digits == phone / mynumber / contract)
    Dim run As Long, ch As String
    Dim i As Long
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch >= "0" And ch <= "9" Then
            run = run + 1
            If run >= 10 Then DetectPii = True: Exit Function
        ElseIf ch = "-" Or ch = " " Or ch = "‐" Then
            ' continue run through separators
        Else
            run = 0
        End If
    Next i
End Function
