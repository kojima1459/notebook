Attribute VB_Name = "modKeyEnroller"
Option Explicit

' ============================================================================
' modKeyEnroller - admin-only key obfuscation tool
' ----------------------------------------------------------------------------
' Workflow:
'   1. Admin runs EnrollKeyInteractive() in the Admin workbook.
'   2. The plaintext key is entered via InputBox (it never touches a sheet).
'   3. The key is XOR+Base64 encoded with the per-distribution pad and
'      written as a hidden defined name into the open chatbot workbook
'      pointed to by the operator (FileDialog).
'   4. Admin saves the chatbot xlsm. The plaintext is discarded.
'
' SECURITY NOTE: The pad is constant per-build. Anyone with the same xlsm
' template + .bas could recompute it. This is "out of sight" only.
' ============================================================================

Public Sub EnrollKeyInteractive()
    Dim plaintext As String
    plaintext = InputBox("APIキーを貼り付けてください（画面表示のみ、保存されません）", "Enroll API Key")
    If LenB(plaintext) = 0 Then Exit Sub

    Dim fd As FileDialog: Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "対象のChatbot.xlsmを選択"
    fd.Filters.Clear
    fd.Filters.Add "Excel Macro Workbook", "*.xlsm"
    If fd.Show <> -1 Then Exit Sub

    Dim wbPath As String: wbPath = fd.SelectedItems(1)
    Dim wb As Workbook
    Set wb = Workbooks.Open(wbPath)

    Dim blob As String
    blob = modKeyVault.EncodeObfuscated(plaintext)

    ' Inject into target workbook's modKeyVault store.
    ' This relies on the target xlsm having modKeyVault.StoreObfuscated and
    ' the same Pad() salt as this enroller.
    Application.Run "'" & wb.Name & "'!modKeyVault.StoreObfuscated", blob

    wb.Save
    wb.Close False

    plaintext = String$(Len(plaintext), "*")
    MsgBox "APIキーを難読化して " & wbPath & " に埋め込みました。", vbInformation
End Sub
