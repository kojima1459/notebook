Attribute VB_Name = "modExtractorExcel"
Option Explicit

' ============================================================================
' modExtractorExcel - extract text from .xlsx / .xls / .xlsm workbooks
' ----------------------------------------------------------------------------
' Opens the file ReadOnly in the current Excel instance and emits one
' "page" per worksheet. Each cell value becomes a row in the page text.
' Formulas are skipped; only the computed display value is captured.
'
' Hidden sheets are included by default - if the workbook author hid them
' intentionally, the admin should pre-filter the file selection.
' ============================================================================

Public Function Extract(ByVal path As String) As ExtractedPage()
    Dim wb As Workbook
    Dim restoreScreen As Boolean: restoreScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False

    On Error GoTo Failed
    Set wb = Workbooks.Open( _
        FileName:=path, _
        ReadOnly:=True, _
        UpdateLinks:=0, _
        IgnoreReadOnlyRecommended:=True, _
        AddToMru:=False)

    Dim pages() As ExtractedPage
    ReDim pages(0 To wb.Worksheets.count - 1)

    Dim s As Long
    For s = 1 To wb.Worksheets.count
        Dim ws As Worksheet
        Set ws = wb.Worksheets(s)
        pages(s - 1).page = s
        pages(s - 1).Text = ExtractSheetText(ws)
    Next s

    wb.Close False
    Application.ScreenUpdating = restoreScreen
    Extract = pages
    Exit Function

Failed:
    Dim msg As String: msg = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    Application.ScreenUpdating = restoreScreen
    On Error GoTo 0
    Err.Raise vbObjectError + &H7301, "modExtractorExcel", _
        "Excel extraction failed for " & path & ": " & msg
End Function

Private Function ExtractSheetText(ByVal ws As Worksheet) As String
    Dim used As Range
    Set used = ws.UsedRange
    If used Is Nothing Then Exit Function
    If used.Cells.count = 0 Then Exit Function

    Dim sb As String
    sb = "[シート: " & ws.Name & "]" & vbCrLf

    Dim r As Long, c As Long, lastR As Long, lastC As Long
    lastR = used.Rows.count
    lastC = used.Columns.count
    Dim startR As Long, startC As Long
    startR = used.Row
    startC = used.Column

    Dim arr As Variant
    arr = used.value ' read everything in one COM round-trip

    If IsArray(arr) Then
        For r = 1 To lastR
            Dim rowBuf As String: rowBuf = ""
            For c = 1 To lastC
                Dim v As Variant: v = arr(r, c)
                If Not IsEmpty(v) Then
                    If Len(rowBuf) > 0 Then rowBuf = rowBuf & vbTab
                    rowBuf = rowBuf & CStr(v)
                End If
            Next c
            If Len(rowBuf) > 0 Then sb = sb & rowBuf & vbCrLf
        Next r
    Else
        ' Single cell case - UsedRange returned a scalar
        sb = sb & CStr(arr) & vbCrLf
    End If
    ExtractSheetText = sb
End Function
