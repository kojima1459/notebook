Attribute VB_Name = "modUsageLogger"
Option Explicit

' ============================================================================
' modUsageLogger - per-user CSV append, locally buffered, flushed to SharePoint
' ----------------------------------------------------------------------------
' Why per-user CSV: avoids cross-user write conflicts on SharePoint. The
' admin workbook stitches all CSVs together via Power Query.
'
' Columns (UTF-8 with BOM, ISO timestamps, doubled-quote escape):
'   timestamp, username, department, role, prompt_tokens, completion_tokens,
'   latency_ms, pii_warn, has_question_text, question_text
' ============================================================================

Private Const HEADER As String = _
    "timestamp,username,department,role,prompt_tokens,completion_tokens,latency_ms,pii_warn,has_question_text,question_text"

Public Sub LogUsage(ByVal promptTokens As Long, _
                    ByVal completionTokens As Long, _
                    ByVal latencyMs As Long, _
                    ByVal piiWarn As Boolean, _
                    ByVal hasQuestionText As Boolean, _
                    ByVal questionText As String)
    Dim row As String
    row = Iso8601(Now) & "," & _
          CsvCell(Environ$("USERNAME")) & "," & _
          CsvCell(modUserProfile.Department()) & "," & _
          CsvCell(modUserProfile.Role()) & "," & _
          promptTokens & "," & _
          completionTokens & "," & _
          latencyMs & "," & _
          IIf(piiWarn, "1", "0") & "," & _
          IIf(hasQuestionText, "1", "0") & "," & _
          CsvCell(IIf(hasQuestionText, questionText, ""))

    AppendLocalBuffer row
    TryFlushToRemote
End Sub

Private Sub AppendLocalBuffer(ByVal row As String)
    Dim path As String: path = modPaths.UsageBufferPath()
    modPaths.EnsureDir Left$(path, InStrRev(path, "\") - 1)
    Dim needsHeader As Boolean
    needsHeader = (Len(Dir$(path)) = 0)
    Dim fnum As Integer: fnum = FreeFile
    Open path For Append As #fnum
    If needsHeader Then
        ' UTF-8 BOM lead so Excel opens it correctly
        Print #fnum, Chr$(239) & Chr$(187) & Chr$(191) & HEADER
    End If
    Print #fnum, row
    Close #fnum
End Sub

Private Sub TryFlushToRemote()
    Dim remoteFolder As String
    remoteFolder = modConfig.GetString("logging", "remote_url", "")
    If LenB(remoteFolder) = 0 Then Exit Sub
    On Error Resume Next
    Dim src As String: src = modPaths.UsageBufferPath()
    Dim dst As String: dst = remoteFolder & "\" & Environ$("USERNAME") & ".csv"
    ' Append-on-remote: read remote tail, write combined back. For SharePoint
    ' mapped drives this is best-effort; failures stay in local buffer.
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(src) Then Exit Sub
    If fso.FileExists(dst) Then
        AppendFile dst, src
    Else
        fso.CopyFile src, dst, True
    End If
    ' On success, clear buffer to avoid duplicate rows
    If Err.Number = 0 Then
        fso.DeleteFile src, True
    End If
    On Error GoTo 0
End Sub

Private Sub AppendFile(ByVal dst As String, ByVal src As String)
    Dim fnumIn As Integer, fnumOut As Integer
    fnumIn = FreeFile
    Open src For Input As #fnumIn
    fnumOut = FreeFile
    Open dst For Append As #fnumOut
    Dim line As String
    Do While Not EOF(fnumIn)
        Line Input #fnumIn, line
        ' Skip header rows (they'd duplicate)
        If Left$(line, Len(HEADER) + 3) <> Chr$(239) & Chr$(187) & Chr$(191) & HEADER _
           And line <> HEADER Then
            Print #fnumOut, line
        End If
    Loop
    Close #fnumIn
    Close #fnumOut
End Sub

Private Function Iso8601(ByVal d As Date) As String
    Iso8601 = Format$(d, "yyyy-mm-ddTHH:nn:ss")
End Function

Private Function CsvCell(ByVal s As Variant) As String
    Dim t As String: t = CStr(s)
    If InStr(t, ",") > 0 Or InStr(t, """") > 0 Or InStr(t, vbLf) > 0 Or InStr(t, vbCr) > 0 Then
        CsvCell = """" & Replace(t, """", """""") & """"
    Else
        CsvCell = t
    End If
End Function
