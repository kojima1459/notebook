Attribute VB_Name = "modBoot"
Option Explicit

' ============================================================================
' modBoot - Workbook_Open entry point
' ----------------------------------------------------------------------------
' Boot sequence (one-time per session):
'   1. Load config.ini
'   2. Ensure first-run profile prompts have been satisfied
'   3. Sync index from SharePoint to local cache (if newer)
'   4. Load index into memory (held in gIndex)
'   5. Show chat form
' Errors at any step are surfaced via MsgBox and abort the boot. Workbook
' should remain readable as a degraded view.
' ============================================================================

Public gIndex As LoadedIndex
Public gReady As Boolean

Public Sub Boot()
    On Error GoTo Failed
    modConfig.EnsureLoaded
    modUserProfile.EnsureFirstRun

    SyncIndexFromRemote

    Dim dir As String: dir = modPaths.IndexCacheDir()
    gIndex = modIndexReader.LoadIndex(dir)
    If Not gIndex.OK Then
        MsgBox "ナレッジindexの読み込みに失敗しました:" & vbCrLf & gIndex.ErrorMessage, _
               vbExclamation, "InternalNotebookLM"
        gReady = False
        Exit Sub
    End If
    gReady = True
    modChatUI.EnsureLayout
    Exit Sub

Failed:
    MsgBox "起動エラー: " & Err.Description, vbCritical, "InternalNotebookLM"
    gReady = False
End Sub

Private Sub SyncIndexFromRemote()
    Dim remote As String: remote = modConfig.GetString("index", "remote_url", "")
    If LenB(remote) = 0 Then Exit Sub

    modPaths.EnsureDir modPaths.IndexCacheDir()

    ' For HTTP-style remotes only. For SharePoint mapped drives, a simple
    ' FileSystemObject copy in admin-guide.md may be preferable.
    If LCase$(Left$(remote, 4)) = "http" Then
        modHttpClient.DownloadBinary remote & "/manifest.json", modPaths.ManifestPath()
        modHttpClient.DownloadBinary remote & "/embeddings.bin", modPaths.EmbeddingsPath()
        modHttpClient.DownloadBinary remote & "/chunks.json", modPaths.ChunksPath()
    End If
End Sub

