Attribute VB_Name = "modUsageLogger"
Option Explicit

' ============================================================================
' modUsageLogger - per-user CSV append, locally buffered, flushed to SharePoint
' ----------------------------------------------------------------------------
' Why per-user CSV: avoids cross-user write conflicts on SharePoint. The
' admin workbook stitches all CSVs together via Power Query.
'
' All I/O goes through ADODB.Stream so non-ASCII department names and
' question bodies stay in UTF-8 regardless of the system codepage.
'
' Columns:
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
    On Error Resume Next
    TryFlushToRemote
    On Error GoTo 0
End Sub

Private Sub AppendLocalBuffer(ByVal row As String)
    Dim path As String: path = modPaths.UsageBufferPath()
    modPaths.EnsureDir Left$(path, InStrRev(path, "\") - 1)
    Dim isNew As Boolean: isNew = (Len(Dir$(path)) = 0)
    If isNew Then
        WriteTextUtf8 path, HEADER & vbLf & row & vbLf, True  ' include BOM
    Else
        AppendTextUtf8 path, row & vbLf
    End If
End Sub

Private Sub TryFlushToRemote()
    Dim remoteFolder As String
    remoteFolder = modConfig.GetString("logging", "remote_url", "")
    If LenB(remoteFolder) = 0 Then Exit Sub

    Dim src As String: src = modPaths.UsageBufferPath()
    If Len(Dir$(src)) = 0 Then Exit Sub

    Dim dst As String: dst = remoteFolder & "\" & Environ$("USERNAME") & ".csv"

    Dim rows As String
    rows = ReadDataRows(src) ' returns just the data rows (no header)
    If LenB(rows) = 0 Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(dst) Then
        WriteTextUtf8 dst, HEADER & vbLf & rows, True
    Else
        AppendTextUtf8 dst, rows
    End If

    ' Truncate the local buffer back to just the header so we never resend
    ' the same rows. If the remote write above raised, we won't reach here
    ' (caller wraps in On Error Resume Next), so rows stay buffered.
    WriteTextUtf8 src, HEADER & vbLf, True
End Sub

Private Function ReadDataRows(ByVal path As String) As String
    Dim full As String: full = ReadTextUtf8(path)
    Dim idx As Long: idx = InStr(full, vbLf)
    If idx = 0 Then Exit Function
    ReadDataRows = Mid$(full, idx + 1)
End Function

Private Sub WriteTextUtf8(ByVal path As String, ByVal text As String, ByVal includeBom As Boolean)
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    If Not includeBom Then
        st.Position = 0
        st.Type = 1
        st.Read 3
        Dim tail() As Byte: tail = st.Read
        st.Close
        Dim f As Integer: f = FreeFile
        Open path For Binary Access Write As #f
        Put #f, 1, tail
        Close #f
        Exit Sub
    End If
    st.SaveToFile path, 2
    st.Close
End Sub

Private Sub AppendTextUtf8(ByVal path As String, ByVal text As String)
    ' Convert text to UTF-8 bytes (no BOM) and append at EOF.
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    st.Position = 0
    st.Type = 1
    st.Read 3
    Dim payload() As Byte: payload = st.Read
    st.Close

    Dim f As Integer: f = FreeFile
    Open path For Binary Access Write As #f
    Dim flen As Long: flen = LOF(f)
    Put #f, flen + 1, payload
    Close #f
End Sub

Private Function ReadTextUtf8(ByVal path As String) As String
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadTextUtf8 = st.ReadText
    st.Close
End Function

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
