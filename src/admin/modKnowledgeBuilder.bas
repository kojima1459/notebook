Attribute VB_Name = "modKnowledgeBuilder"
Option Explicit

' ============================================================================
' modKnowledgeBuilder - admin entry point for index construction
' ----------------------------------------------------------------------------
' Flow:
'   1. Pick source files (FileDialog)
'   2. Extract -> normalize -> chunk
'   3. Batch-embed each ~16 chunks
'   4. L2-normalize all vectors
'   5. Write atomic index to local cache; optionally copy to remote
' ============================================================================

Private Const EMBED_BATCH As Long = 16

Public Sub BuildIndex()
    modConfig.EnsureLoaded

    Dim files() As String
    files = PickFiles()
    If LBound(files) > UBound(files) Then Exit Sub

    Dim allChunks() As Chunk
    Dim chunkCount As Long
    ReDim allChunks(0 To 0)

    Dim i As Long
    For i = LBound(files) To UBound(files)
        Application.StatusBar = "抽出中: " & files(i)
        DoEvents
        Dim source As String
        Dim pages() As ExtractedPage
        On Error Resume Next
        pages = modExtractor.ExtractFile(files(i), source)
        If Err.Number <> 0 Then
            Debug.Print "Skipped " & files(i) & ": " & Err.Description
            Err.Clear
            On Error GoTo 0
            GoTo NextFile
        End If
        On Error GoTo 0

        Dim p As Long
        For p = LBound(pages) To UBound(pages)
            Dim chunks() As Chunk
            chunks = modChunker.ChunkText(source, pages(p).page, pages(p).Text)
            Dim k As Long
            For k = LBound(chunks) To UBound(chunks)
                If chunkCount > UBound(allChunks) Then ReDim Preserve allChunks(0 To chunkCount * 2 + 1)
                allChunks(chunkCount) = chunks(k)
                chunkCount = chunkCount + 1
            Next k
        Next p
NextFile:
    Next i

    If chunkCount = 0 Then
        Application.StatusBar = False
        MsgBox "抽出可能なテキストが見つかりませんでした。", vbExclamation
        Exit Sub
    End If
    ReDim Preserve allChunks(0 To chunkCount - 1)

    ' Batch-embed
    Dim flat() As Double
    Dim dim_ As Long: dim_ = 0
    Dim filled As Long: filled = 0
    Dim batchStart As Long: batchStart = 0
    Do While batchStart < chunkCount
        Dim batchEnd As Long
        batchEnd = batchStart + EMBED_BATCH - 1
        If batchEnd > chunkCount - 1 Then batchEnd = chunkCount - 1
        Application.StatusBar = "embed中: " & (batchStart + 1) & " / " & chunkCount
        DoEvents

        Dim batch() As String
        ReDim batch(0 To batchEnd - batchStart)
        Dim j As Long
        For j = 0 To batchEnd - batchStart
            batch(j) = allChunks(batchStart + j).Text
        Next j

        Dim emb As EmbedResult
        emb = modApiGateway.Embed(batch)
        If Not emb.OK Then
            Application.StatusBar = False
            MsgBox "embed失敗: " & emb.ErrorMessage, vbCritical
            Exit Sub
        End If
        If dim_ = 0 Then
            dim_ = emb.Dim_
            ReDim flat(0 To chunkCount * dim_ - 1)
        ElseIf emb.Dim_ <> dim_ Then
            Application.StatusBar = False
            MsgBox "embedding 次元の不一致 (" & emb.Dim_ & " vs " & dim_ & ")", vbCritical
            Exit Sub
        End If

        Dim destBase As Long: destBase = batchStart * dim_
        Dim m As Long
        For m = 0 To (emb.Count_ * dim_) - 1
            flat(destBase + m) = emb.Vectors(m)
        Next m
        filled = filled + emb.Count_

        batchStart = batchEnd + 1
    Loop

    ' Normalize
    modChunker.L2NormalizeInPlace flat, dim_, chunkCount

    ' Write
    Dim outDir As String: outDir = modPaths.IndexCacheDir()
    Dim model As String: model = modConfig.GetString("api", "embed_deployment", "")
    modIndexWriter.WriteIndex outDir, allChunks, flat, dim_, chunkCount, model, _
        Format$(Now, "yyyymmdd-hhnnss")

    ' Optional: copy to remote folder for distribution
    PublishToRemote outDir

    Application.StatusBar = False
    MsgBox chunkCount & " chunks をindex化しました。", vbInformation
End Sub

Private Function PickFiles() As String()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.AllowMultiSelect = True
    fd.Title = "ナレッジファイルを選択"
    fd.Filters.Clear
    fd.Filters.Add "対応ファイル", "*.txt;*.md;*.csv;*.pdf;*.docx;*.xlsx"
    Dim out() As String
    If fd.Show <> -1 Then
        ReDim out(-1 To -1)
        PickFiles = out
        Exit Function
    End If
    ReDim out(0 To fd.SelectedItems.Count - 1)
    Dim i As Long
    For i = 1 To fd.SelectedItems.Count
        out(i - 1) = fd.SelectedItems(i)
    Next i
    PickFiles = out
End Function

Private Sub PublishToRemote(ByVal srcDir As String)
    Dim remote As String: remote = modConfig.GetString("index", "remote_url", "")
    If LenB(remote) = 0 Then Exit Sub
    If LCase$(Left$(remote, 4)) = "http" Then
        ' HTTP upload not implemented in MVP (need PUT/webdav semantics).
        Exit Sub
    End If
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If Not fso.FolderExists(remote) Then fso.CreateFolder remote
    fso.CopyFile srcDir & "\manifest.json", remote & "\manifest.json", True
    fso.CopyFile srcDir & "\embeddings.bin", remote & "\embeddings.bin", True
    fso.CopyFile srcDir & "\chunks.json", remote & "\chunks.json", True
    On Error GoTo 0
End Sub
