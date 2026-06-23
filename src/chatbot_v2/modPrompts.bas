Attribute VB_Name = "modPrompts"
Option Explicit

' ============================================================================
' modPrompts - Read system prompts from the `system_prompt` sheet
' ----------------------------------------------------------------------------
' Sheet structure (column A = key, column B = prompt text):
'   row 2: router
'   row 3: drafter
'   row 4: verifier
' Prompts can be edited by admins without touching VBA.
' ============================================================================

Private Const SHEET_NAME As String = "system_prompt"

Public Function GetRouter() As String
    GetRouter = ReadKey("router")
End Function

Public Function GetDrafter() As String
    GetDrafter = ReadKey("drafter")
End Function

Public Function GetVerifier() As String
    GetVerifier = ReadKey("verifier")
End Function

Private Function ReadKey(ByVal key As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim i As Long
    For i = 2 To lastRow
        If StrComp(CStr(ws.Cells(i, 1).value), key, vbTextCompare) = 0 Then
            ReadKey = CStr(ws.Cells(i, 2).value)
            Exit Function
        End If
    Next i
End Function
