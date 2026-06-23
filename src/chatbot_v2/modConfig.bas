Attribute VB_Name = "modConfig"
Option Explicit

' ============================================================================
' modConfig - Read config from `config` sheet (named ranges)
' ----------------------------------------------------------------------------
' All settings live in the `config` sheet so non-engineers can tune behavior
' without opening VBA. Settings are looked up by key string against col A,
' returning col B value.
'
' Standard keys (see config sheet for full list):
'   router_max_chunks       (Long, default 8)
'   verifier_enabled        (Boolean, default TRUE)
'   max_context_chars       (Long, default 60000)
'   debug_mode              (Boolean, default FALSE)
'   department_id           (String, current user's dept)
'   draft_temperature       (Double, default 0.15) -- documentation only
' ============================================================================

Private Const SHEET_NAME As String = "config"
Private mLoaded As Boolean

Public Sub EnsureLoaded()
    ' Verify config sheet exists; nothing else to load
    On Error GoTo NoSheet
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    mLoaded = True
    Exit Sub
NoSheet:
    MsgBox "config シートが見つかりません。", vbCritical
End Sub

Public Function GetString(ByVal key As String, ByVal defaultValue As String) As String
    Dim v As Variant
    v = LookupValue(key)
    If IsEmpty(v) Then
        GetString = defaultValue
    Else
        GetString = CStr(v)
    End If
End Function

Public Function GetLong(ByVal key As String, ByVal defaultValue As Long) As Long
    Dim v As Variant
    v = LookupValue(key)
    If IsEmpty(v) Or Not IsNumeric(v) Then
        GetLong = defaultValue
    Else
        GetLong = CLng(v)
    End If
End Function

Public Function GetDouble(ByVal key As String, ByVal defaultValue As Double) As Double
    Dim v As Variant
    v = LookupValue(key)
    If IsEmpty(v) Or Not IsNumeric(v) Then
        GetDouble = defaultValue
    Else
        GetDouble = CDbl(v)
    End If
End Function

Public Function GetBool(ByVal key As String, ByVal defaultValue As Boolean) As Boolean
    Dim v As Variant
    v = LookupValue(key)
    If IsEmpty(v) Then
        GetBool = defaultValue
    ElseIf VarType(v) = vbBoolean Then
        GetBool = CBool(v)
    ElseIf IsNumeric(v) Then
        GetBool = (CLng(v) <> 0)
    Else
        Dim s As String: s = LCase$(Trim$(CStr(v)))
        GetBool = (s = "true" Or s = "1" Or s = "yes" Or s = "on")
    End If
End Function

Public Sub SetValue(ByVal key As String, ByVal value As Variant)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Dim r As Long: r = FindKeyRow(ws, key)
    If r = 0 Then
        ' append new row
        r = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
        ws.Cells(r, 1).value = key
    End If
    ws.Cells(r, 2).value = value
End Sub

Private Function LookupValue(ByVal key As String) As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Dim r As Long: r = FindKeyRow(ws, key)
    If r = 0 Then Exit Function
    LookupValue = ws.Cells(r, 2).value
End Function

Private Function FindKeyRow(ByVal ws As Worksheet, ByVal key As String) As Long
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim i As Long
    For i = 2 To lastRow      ' row 1 is header
        If StrComp(CStr(ws.Cells(i, 1).value), key, vbTextCompare) = 0 Then
            FindKeyRow = i
            Exit Function
        End If
    Next i
End Function
