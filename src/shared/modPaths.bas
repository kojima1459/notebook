Attribute VB_Name = "modPaths"
Option Explicit

' ============================================================================
' modPaths - centralized path resolution
' ============================================================================

Public Function AppDataRoot() As String
    AppDataRoot = Environ$("LOCALAPPDATA") & "\InternalNotebookLM"
End Function

Public Function IndexCacheDir() As String
    IndexCacheDir = modConfig.GetString("index", "local_cache", AppDataRoot() & "\index")
End Function

Public Function UsageBufferPath() As String
    UsageBufferPath = modConfig.GetString("logging", "local_buffer", _
        AppDataRoot() & "\usage_buffer.csv")
End Function

Public Function ManifestPath() As String
    ManifestPath = IndexCacheDir() & "\manifest.json"
End Function

Public Function ChunksPath() As String
    ChunksPath = IndexCacheDir() & "\chunks.json"
End Function

Public Function EmbeddingsPath() As String
    EmbeddingsPath = IndexCacheDir() & "\embeddings.bin"
End Function

Public Sub EnsureDir(ByVal path As String)
    If LenB(path) = 0 Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim parts() As String, cur As String, i As Long
    parts = Split(path, "\")
    cur = parts(0)
    For i = 1 To UBound(parts)
        cur = cur & "\" & parts(i)
        If Not fso.FolderExists(cur) Then
            On Error Resume Next
            fso.CreateFolder cur
            On Error GoTo 0
        End If
    Next i
End Sub
