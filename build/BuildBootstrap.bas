Attribute VB_Name = "BuildBootstrap"
Option Explicit

' ============================================================================
' BuildBootstrap - in-Excel builder, ASCII-only for cross-platform safety
' ----------------------------------------------------------------------------
' Works on both Excel for Mac and Excel for Windows. All messages kept in
' English to avoid encoding issues when this .bas file is imported into
' Mac Excel's VBA editor (which assumes system codepage for .bas content).
'
' Usage:
'   1. Open Excel, new blank workbook, Save As Chatbot.xlsm (macro-enabled)
'   2. Tools > Macro > Visual Basic Editor (or option+F11 on Mac, Alt+F11 on Win)
'   3. File > Import File > pick this BuildBootstrap.bas
'   4. Click inside Sub BuildChatbot, press F5
'   5. Pick the unzipped notebook folder when prompted
'   6. Cmd+S (Mac) / Ctrl+S (Win) to save
'   7. Right-click BuildBootstrap in the tree > Remove > No (don't export)
'   8. Save again
'
' Repeat with BuildAdmin for Admin_KnowledgeBuilder.xlsm.
'
' Requires: Trust Center > "Trust access to the VBA project object model" ON.
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

    Dim sep As String
    sep = Application.PathSeparator

    Dim sharedDir As String, roleDir As String, vendorDir As String
    sharedDir = folder & sep & "src" & sep & "shared" & sep
    vendorDir = folder & sep & "build" & sep & "vendor" & sep

    If Len(Dir(sharedDir & "modConfig.bas")) = 0 Then
        MsgBox "src/shared not found inside: " & vbCrLf & folder & vbCrLf & vbCrLf & _
               "Please pick the root notebook folder (the one that directly contains " & _
               "the build/, src/, docs/ subfolders).", _
               vbExclamation, "BuildBootstrap"
        Exit Sub
    End If

    Dim sharedFiles As Variant, roleFiles As Variant
    sharedFiles = Array("modTypes.bas", "modConfig.bas", "modPaths.bas", _
                        "modHttpClient.bas", "modKeyVault.bas", "modApiGateway.bas")

    If target = "chatbot" Then
        roleDir = folder & sep & "src" & sep & "chatbot" & sep
        roleFiles = Array("modIndexReader.bas", "modSimilarity.bas", "modRagEngine.bas", _
                          "modUserProfile.bas", "modPiiGuard.bas", "modRateLimiter.bas", _
                          "modUsageLogger.bas", "modChatUI.bas", "modBoot.bas")
    Else
        roleDir = folder & sep & "src" & sep & "admin" & sep
        roleFiles = Array("modChunker.bas", "modExtractor.bas", "modExtractorWord.bas", _
                          "modExtractorAcrobat.bas", "modExtractorExcel.bas", _
                          "modIndexWriter.bas", "modKnowledgeBuilder.bas", _
                          "modKeyEnroller.bas", "modUsageAggregator.bas")
    End If

    Dim vbp As Object
    Set vbp = ThisWorkbook.VBProject

    Dim i As Long
    Application.StatusBar = "Importing shared modules..."
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
        Application.StatusBar = "Injecting ThisWorkbook code..."
        InjectThisWorkbook vbp, roleDir & "ThisWorkbook.cls"
    End If

    Application.StatusBar = False

    MsgBox "Done! Next steps:" & vbCrLf & vbCrLf & _
           "1. Press Cmd+S (Mac) or Ctrl+S (Windows) to save." & vbCrLf & _
           "2. Right-click 'BuildBootstrap' in the left tree, choose 'Remove BuildBootstrap', " & _
           "click No when asked to export." & vbCrLf & _
           "3. Save the workbook again, then close it.", _
           vbInformation, "Build " & target & " - complete"
    Exit Sub

Failed:
    Application.StatusBar = False
    MsgBox "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Most common cause: 'Trust access to the VBA project object model' is not enabled. " & _
           "On Mac: Excel > Preferences > Security & Privacy. " & _
           "On Win: File > Options > Trust Center > Trust Center Settings > Macro Settings.", _
           vbCritical, "BuildBootstrap error"
End Sub

Private Sub ImportIfMissing(ByVal vbp As Object, ByVal path As String)
    If Len(Dir(path)) = 0 Then
        Debug.Print "Skip (file missing): " & path
        Exit Sub
    End If

    Dim sep As String
    sep = Application.PathSeparator

    Dim fileName As String, modName As String
    fileName = Mid(path, InStrRev(path, sep) + 1)
    modName = Left(fileName, InStr(fileName, ".") - 1)

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
    If Len(Dir(path)) = 0 Then Exit Sub

    Dim text As String
    Dim fn As Integer
    fn = FreeFile
    Open path For Input As #fn
    Dim line As String
    Do While Not EOF(fn)
        Line Input #fn, line
        text = text & line & vbCrLf
    Loop
    Close #fn

    Dim marker As String
    marker = "Attribute VB_Exposed = True"
    Dim idx As Long
    idx = InStr(text, marker)
    If idx > 0 Then
        Dim eol As Long
        eol = InStr(idx, text, vbLf)
        If eol > 0 Then text = Mid(text, eol + 1)
    End If

    Dim cm As Object
    Set cm = vbp.VBComponents.Item("ThisWorkbook").CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString text
End Sub

Private Function PickFolder() As String
#If Mac Then
    Dim s As String
    On Error Resume Next
    s = MacScript("return POSIX path of (choose folder with prompt ""Pick the notebook folder (containing build/ and src/)"")")
    On Error GoTo 0
    If Right(s, 1) = "/" Then s = Left(s, Len(s) - 1)
    PickFolder = s
#Else
    Dim fd As Object
    Set fd = Application.FileDialog(4)
    fd.Title = "Pick the notebook folder (containing build/ and src/)"
    If fd.Show = -1 Then PickFolder = fd.SelectedItems(1)
#End If
End Function
