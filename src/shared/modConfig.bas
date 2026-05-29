Attribute VB_Name = "modConfig"
Option Explicit

' ============================================================================
' modConfig - INI loader and config accessor
' ----------------------------------------------------------------------------
' Resolves config.ini from %LOCALAPPDATA%\InternalNotebookLM\config.ini,
' falling back to the workbook folder. Expands %ENV% style variables.
' Caches values in a Dictionary keyed by "section.key" (lower-case).
' ============================================================================

Private mLoaded As Boolean
Private mValues As Object ' Scripting.Dictionary

Public Sub EnsureLoaded()
    If mLoaded Then Exit Sub
    Set mValues = CreateObject("Scripting.Dictionary")
    mValues.CompareMode = 1 ' TextCompare
    LoadIniFile ResolveConfigPath()
    mLoaded = True
End Sub

Public Function GetString(ByVal section As String, ByVal key As String, _
                          Optional ByVal defaultValue As String = "") As String
    EnsureLoaded
    Dim k As String: k = LCase$(section) & "." & LCase$(key)
    If mValues.Exists(k) Then
        GetString = ExpandEnv(CStr(mValues(k)))
    Else
        GetString = defaultValue
    End If
End Function

Public Function GetLong(ByVal section As String, ByVal key As String, _
                        Optional ByVal defaultValue As Long = 0) As Long
    Dim s As String: s = GetString(section, key, "")
    If Len(s) = 0 Then
        GetLong = defaultValue
    Else
        GetLong = CLng(s)
    End If
End Function

Public Function ResolveConfigPath() As String
    Dim primary As String
    primary = Environ$("LOCALAPPDATA") & "\InternalNotebookLM\config.ini"
    If FileExists(primary) Then
        ResolveConfigPath = primary
        Exit Function
    End If
    ' Fallback: alongside the workbook (admin/dev convenience)
    Dim sibling As String
    sibling = ThisWorkbook.Path & "\config.ini"
    If FileExists(sibling) Then
        ResolveConfigPath = sibling
        Exit Function
    End If
    ResolveConfigPath = primary ' return primary even if missing
End Function

Private Sub LoadIniFile(ByVal path As String)
    If Not FileExists(path) Then Exit Sub
    Dim fnum As Integer: fnum = FreeFile
    Open path For Input As #fnum
    Dim line As String, section As String, eqPos As Long
    section = ""
    Do While Not EOF(fnum)
        Line Input #fnum, line
        line = Trim$(line)
        If Len(line) = 0 Then GoTo Continue
        Dim first As String: first = Left$(line, 1)
        If first = ";" Or first = "#" Then GoTo Continue
        If first = "[" And Right$(line, 1) = "]" Then
            section = Mid$(line, 2, Len(line) - 2)
            GoTo Continue
        End If
        eqPos = InStr(line, "=")
        If eqPos > 0 Then
            Dim k As String, v As String
            k = Trim$(Left$(line, eqPos - 1))
            v = Trim$(Mid$(line, eqPos + 1))
            mValues(LCase$(section) & "." & LCase$(k)) = v
        End If
Continue:
    Loop
    Close #fnum
End Sub

Private Function ExpandEnv(ByVal s As String) As String
    Dim out As String: out = s
    Dim i As Long, openPos As Long, closePos As Long
    i = 1
    Do
        openPos = InStr(i, out, "%")
        If openPos = 0 Then Exit Do
        closePos = InStr(openPos + 1, out, "%")
        If closePos = 0 Then Exit Do
        Dim name As String, val As String
        name = Mid$(out, openPos + 1, closePos - openPos - 1)
        val = Environ$(name)
        out = Left$(out, openPos - 1) & val & Mid$(out, closePos + 1)
        i = openPos + Len(val)
    Loop
    ExpandEnv = out
End Function

Private Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(path)) > 0)
    On Error GoTo 0
End Function
