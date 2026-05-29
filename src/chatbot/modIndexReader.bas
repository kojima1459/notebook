Attribute VB_Name = "modIndexReader"
Option Explicit

' ============================================================================
' modIndexReader - Load the index into memory at Workbook_Open
' ----------------------------------------------------------------------------
' Reads manifest.json, embeddings.bin, chunks.json. Holds all chunk metadata
' + a single contiguous Double() array for fast scoring.
' ============================================================================

Public Type LoadedIndex
    OK As Boolean
    Dim_ As Long
    Count_ As Long
    EmbedModel As String
    Version As String
    Flat() As Double             ' [count_ * dim_]
    Chunks() As Chunk            ' parallel to vectors
    ErrorMessage As String
End Type

Public Function LoadIndex(ByVal dir As String) As LoadedIndex
    Dim r As LoadedIndex

    Dim manifestPath As String, binPath As String, chunksPath As String
    manifestPath = dir & "\manifest.json"
    binPath = dir & "\embeddings.bin"
    chunksPath = dir & "\chunks.json"

    If Not FileExists(manifestPath) Then
        r.OK = False
        r.ErrorMessage = "manifest.json not found in " & dir
        LoadIndex = r
        Exit Function
    End If

    On Error GoTo Failed

    Dim manifestJson As String: manifestJson = ReadAllText(manifestPath)
    Dim mf As Object: Set mf = JsonConverter.ParseJson(manifestJson)
    r.Dim_ = CLng(mf("dim"))
    r.Count_ = CLng(mf("count"))
    r.EmbedModel = CStr(mf("embed_model"))
    r.Version = CStr(mf("version"))

    If r.Count_ > 0 And r.Dim_ > 0 Then
        ReDim r.Flat(0 To r.Count_ * r.Dim_ - 1)
        Dim fnum As Integer: fnum = FreeFile
        Open binPath For Binary Access Read As #fnum
        Get #fnum, 1, r.Flat
        Close #fnum
    End If

    Dim chunksJson As String: chunksJson = ReadAllText(chunksPath)
    Dim arr As Object: Set arr = JsonConverter.ParseJson(chunksJson)
    Dim n As Long: n = arr.Count
    If n > 0 Then
        ReDim r.Chunks(0 To n - 1)
        Dim i As Long
        For i = 1 To n
            r.Chunks(i - 1).Id = CStr(arr(i)("id"))
            r.Chunks(i - 1).Source = CStr(arr(i)("source"))
            r.Chunks(i - 1).page = CLng(arr(i)("page"))
            r.Chunks(i - 1).StartCharOffset = CLng(arr(i)("start"))
            r.Chunks(i - 1).Text = CStr(arr(i)("text"))
        Next i
    Else
        ReDim r.Chunks(-1 To -1)
    End If

    r.OK = True
    LoadIndex = r
    Exit Function

Failed:
    r.OK = False
    r.ErrorMessage = "LoadIndex failed: " & Err.Description
    LoadIndex = r
End Function

Private Function ReadAllText(ByVal path As String) As String
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2 ' text
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadAllText = st.ReadText
    st.Close
End Function

Private Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(path)) > 0)
    On Error GoTo 0
End Function
