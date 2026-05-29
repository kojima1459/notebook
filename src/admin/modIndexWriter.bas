Attribute VB_Name = "modIndexWriter"
Option Explicit

' ============================================================================
' modIndexWriter - Persist the embedding index to disk
' ----------------------------------------------------------------------------
' Layout (atomic publish):
'   manifest.json     {"version", "dim", "count", "embed_model", "built_at"}
'   embeddings.bin    raw little-endian Double, count*dim values, no header
'   chunks.json       array of {"id","source","page","start","text"}
'
' Write order: chunks.json.new, embeddings.bin.new, then rename
' to *.new, then atomically replace manifest.json LAST. Readers gate on the
' manifest's version so half-written files are never observed.
' ============================================================================

Public Sub WriteIndex(ByVal outDir As String, _
                      ByRef chunks() As Chunk, _
                      ByRef flatVectors() As Double, _
                      ByVal dim_ As Long, _
                      ByVal count As Long, _
                      ByVal embedModel As String, _
                      ByVal version As String)
    modPaths.EnsureDir outDir

    Dim binPath As String, chunksPath As String, manifestPath As String
    binPath = outDir & "\embeddings.bin.new"
    chunksPath = outDir & "\chunks.json.new"
    manifestPath = outDir & "\manifest.json.new"

    WriteEmbeddingsBin binPath, flatVectors, dim_ * count
    WriteChunksJson chunksPath, chunks
    WriteManifestJson manifestPath, version, dim_, count, embedModel

    ' Atomic publish: rename .new -> final, manifest last.
    SafeReplace binPath, outDir & "\embeddings.bin"
    SafeReplace chunksPath, outDir & "\chunks.json"
    SafeReplace manifestPath, outDir & "\manifest.json"
End Sub

Private Sub WriteEmbeddingsBin(ByVal path As String, ByRef flat() As Double, ByVal n As Long)
    Dim fnum As Integer: fnum = FreeFile
    Open path For Binary Access Write As #fnum
    If n > 0 Then
        ' VBA Put on Double() writes the contiguous IEEE-754 bytes
        Put #fnum, 1, flat
    End If
    Close #fnum
End Sub

Private Sub WriteChunksJson(ByVal path As String, ByRef chunks() As Chunk)
    Dim sb As String, i As Long
    sb = "[" & vbLf
    For i = LBound(chunks) To UBound(chunks)
        Dim sep As String: sep = ","
        If i = UBound(chunks) Then sep = ""
        sb = sb & "  {" & _
            """id"":" & JsStr(chunks(i).Id) & "," & _
            """source"":" & JsStr(chunks(i).Source) & "," & _
            """page"":" & chunks(i).page & "," & _
            """start"":" & chunks(i).StartCharOffset & "," & _
            """text"":" & JsStr(chunks(i).Text) & _
            "}" & sep & vbLf
    Next i
    sb = sb & "]" & vbLf
    WriteTextUtf8 path, sb
End Sub

Private Sub WriteManifestJson(ByVal path As String, ByVal version As String, _
                              ByVal dim_ As Long, ByVal count As Long, _
                              ByVal embedModel As String)
    Dim sb As String
    sb = "{" & vbLf
    sb = sb & "  ""version"": " & JsStr(version) & "," & vbLf
    sb = sb & "  ""dim"": " & dim_ & "," & vbLf
    sb = sb & "  ""count"": " & count & "," & vbLf
    sb = sb & "  ""embed_model"": " & JsStr(embedModel) & "," & vbLf
    sb = sb & "  ""built_at"": " & JsStr(Format$(Now, "yyyy-mm-ddTHH:nn:ss")) & vbLf
    sb = sb & "}" & vbLf
    WriteTextUtf8 path, sb
End Sub

Private Sub WriteTextUtf8(ByVal path As String, ByVal text As String)
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2 ' text
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    ' Strip the BOM that ADODB.Stream prepends. JSON spec disallows BOM.
    st.Position = 0
    st.Type = 1
    st.Read 3
    Dim tail() As Byte
    tail = st.Read
    st.Close
    Dim fnum As Integer: fnum = FreeFile
    Open path For Binary Access Write As #fnum
    Put #fnum, 1, tail
    Close #fnum
End Sub

Private Sub SafeReplace(ByVal newPath As String, ByVal targetPath As String)
    On Error Resume Next
    Kill targetPath
    On Error GoTo 0
    Name newPath As targetPath
End Sub

Private Function JsStr(ByVal s As String) As String
    Dim out As String, i As Long, ch As String, code As Long
    out = """"
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        Select Case code
            Case 34:  out = out & "\"""
            Case 92:  out = out & "\\"
            Case 8:   out = out & "\b"
            Case 9:   out = out & "\t"
            Case 10:  out = out & "\n"
            Case 12:  out = out & "\f"
            Case 13:  out = out & "\r"
            Case Else
                If code < 32 Then
                    out = out & "\u" & Right$("0000" & Hex(code), 4)
                Else
                    out = out & ch
                End If
        End Select
    Next i
    JsStr = out & """"
End Function
