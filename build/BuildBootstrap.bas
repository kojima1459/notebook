Attribute VB_Name = "BuildBootstrap"
Option Explicit

' ============================================================================
' BuildBootstrap - in-Excel builder (no PowerShell needed)
' ----------------------------------------------------------------------------
' For locked-down corporate PCs where PowerShell is unavailable. Run this
' inside a freshly saved blank xlsm to populate it with all the modules
' needed for Chatbot.xlsm or Admin_KnowledgeBuilder.xlsm.
'
' Usage:
'   1. Open Excel, create a new workbook, Save As Chatbot.xlsm (macro-enabled)
'   2. Press Alt+F11 (VBA editor)
'   3. File > Import File... > select this file (BuildBootstrap.bas)
'   4. Press F5 with cursor inside BuildChatbot, or run from menu
'   5. Pick the unzipped notebook folder when prompted
'   6. Ctrl+S to save
'   7. Right-click BuildBootstrap in the tree > Remove BuildBootstrap > No (don't export)
'
' Repeat with BuildAdmin for Admin_KnowledgeBuilder.xlsm.
'
' Requires: Trust Center > Macro Settings > "Trust access to the VBA project
' object model" enabled.
' ============================================================================

Public Sub BuildChatbot()
    BuildOne "chatbot"
End Sub

Public Sub BuildAdmin()
    BuildOne "admin"
End Sub

Private Sub BuildOne(ByVal target As String)
    On Error GoTo Failed

    Dim folder As String
    folder = PickFolder()
    If LenB(folder) = 0 Then Exit Sub

    Dim sharedDir As String, roleDir As String, vendorDir As String
    sharedDir = folder & "\src\shared\"
    vendorDir = folder & "\build\vendor\"

    If Len(Dir$(sharedDir & "modConfig.bas")) = 0 Then
        MsgBox "選んだフォルダの中に src\shared\ が見つかりません。" & vbCrLf & _
               "解凍した notebook フォルダ自体 (中に build/, src/, docs/ がある階層) を選んでください。", _
               vbExclamation
        Exit Sub
    End If

    Dim sharedFiles As Variant, roleFiles As Variant
    sharedFiles = Array("modTypes.bas", "modConfig.bas", "modPaths.bas", _
                        "modHttpClient.bas", "modKeyVault.bas", "modApiGateway.bas")

    If target = "chatbot" Then
        roleDir = folder & "\src\chatbot\"
        roleFiles = Array("modIndexReader.bas", "modSimilarity.bas", "modRagEngine.bas", _
                          "modUserProfile.bas", "modPiiGuard.bas", "modRateLimiter.bas", _
                          "modUsageLogger.bas", "modChatUI.bas", "modBoot.bas")
    Else
        roleDir = folder & "\src\admin\"
        roleFiles = Array("modChunker.bas", "modExtractor.bas", "modExtractorWord.bas", _
                          "modExtractorAcrobat.bas", "modExtractorExcel.bas", _
                          "modIndexWriter.bas", "modKnowledgeBuilder.bas", _
                          "modKeyEnroller.bas", "modUsageAggregator.bas")
    End If

    Dim vbp As Object
    Set vbp = ThisWorkbook.VBProject

    Application.StatusBar = "Importing shared modules..."
    Dim i As Long
    For i = LBound(sharedFiles) To UBound(sharedFiles)
        ImportIfMissing vbp, sharedDir & sharedFiles(i)
    Next i

    Application.StatusBar = "Importing JsonConverter..."
    ImportIfMissing vbp, vendorDir & "JsonConverter.bas"

    Application.StatusBar = "Importing " & target & " modules..."
    For i = LBound(roleFiles) To UBound(roleFiles)
        ImportIfMissing vbp, roleDir & roleFiles(i)
    Next i

    If target = "chatbot" Then
        Application.StatusBar = "Injecting ThisWorkbook..."
        InjectThisWorkbook vbp, roleDir & "ThisWorkbook.cls"
    End If

    Application.StatusBar = False

    MsgBox "完了しました！" & vbCrLf & vbCrLf & _
           "次にやること:" & vbCrLf & _
           "1. Ctrl+S で保存" & vbCrLf & _
           "2. 左側のツリーから BuildBootstrap を右クリック → 解放 → 「いいえ」(エクスポートしない)" & vbCrLf & _
           "3. もう一度 Ctrl+S で保存して閉じる", _
           vbInformation, "Build " & target
    Exit Sub

Failed:
    Application.StatusBar = False
    MsgBox "エラー: " & Err.Description & vbCrLf & vbCrLf & _
           "「Trust access to the VBA project object model」がONになっているか確認してください。", _
           vbCritical
End Sub

Private Sub ImportIfMissing(ByVal vbp As Object, ByVal path As String)
    If Len(Dir$(path)) = 0 Then
        Debug.Print "Skip (missing): " & path
        Exit Sub
    End If

    Dim fileName As String, modName As String
    fileName = Mid$(path, InStrRev(path, "\") + 1)
    modName = Left$(fileName, InStr(fileName, ".") - 1)

    On Error Resume Next
    Dim existing As Object
    Set existing = vbp.VBComponents.Item(modName)
    On Error GoTo 0
    If Not existing Is Nothing Then
        Debug.Print "Skip (already present): " & modName
        Exit Sub
    End If

    vbp.VBComponents.Import path
    Debug.Print "Imported: " & modName
End Sub

Private Sub InjectThisWorkbook(ByVal vbp As Object, ByVal path As String)
    If Len(Dir$(path)) = 0 Then Exit Sub

    Dim text As String
    Dim fn As Integer: fn = FreeFile
    Open path For Input As #fn
    Dim line As String
    Do While Not EOF(fn)
        Line Input #fn, line
        text = text & line & vbCrLf
    Loop
    Close #fn

    Dim marker As String: marker = "Attribute VB_Exposed = True"
    Dim idx As Long: idx = InStr(text, marker)
    If idx > 0 Then
        Dim eol As Long
        eol = InStr(idx, text, vbLf)
        If eol > 0 Then text = Mid$(text, eol + 1)
    End If

    Dim cm As Object
    Set cm = vbp.VBComponents.Item("ThisWorkbook").CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString text
End Sub

Private Function PickFolder() As String
    Dim fd As Object
    Set fd = Application.FileDialog(4) ' msoFileDialogFolderPicker
    fd.Title = "解凍した notebook フォルダ (中に build/ や src/ がある階層) を選んでください"
    If fd.Show = -1 Then
        PickFolder = fd.SelectedItems(1)
    End If
End Function
