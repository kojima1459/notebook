Attribute VB_Name = "modKnowledgeBase"
Option Explicit

' ============================================================================
' modKnowledgeBase - Read knowledge_base sheet, filter by department
' ----------------------------------------------------------------------------
' Sheet `knowledge_base` columns:
'   A: chunk_id
'   B: source              (sample_03.pdf etc)
'   C: display             (human-readable name)
'   D: domain              (瑕疵保証責任保険 / 生産物回収費用保険 / 費用利益保険全般 / ...)
'   E: doc_type            (普通保険約款 / 引受ガイドライン / FAQ / 研修資料 / 解説 / ハンドブック)
'   F: section_header      ([source p.N] > 第N章 > 第N条)
'   G: summary             (1-line, <=80 chars)
'   H: keywords            (comma separated, 5-10)
'   I: full_text           (the chunk body)
'   J: dept_scope          (公開対象部署: 'common' or specific dept_id)
'
' Department filter: a chunk is visible to user dept D if:
'   dept_scope == "common"  OR  dept_scope == D
' ============================================================================

Private Const SHEET_NAME As String = "knowledge_base"
Private Const COL_ID As Long = 1
Private Const COL_SOURCE As Long = 2
Private Const COL_DISPLAY As Long = 3
Private Const COL_DOMAIN As Long = 4
Private Const COL_DOC_TYPE As Long = 5
Private Const COL_HEADER As Long = 6
Private Const COL_SUMMARY As Long = 7
Private Const COL_KEYWORDS As Long = 8
Private Const COL_TEXT As Long = 9
Private Const COL_DEPT_SCOPE As Long = 10

Public Function RowCount() As Long
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, COL_ID).End(xlUp).row
    If lastRow < 2 Then Exit Function
    RowCount = lastRow - 1
End Function

' ----------------------------------------------------------------------------
' BuildRouterTable: returns the compact list for the router's prompt.
' One line per visible chunk: "id | domain | doc_type | display | summary | keywords"
' Filtered by department. For the MVP this is sent to the LLM in one shot.
' ----------------------------------------------------------------------------
Public Function BuildRouterTable(ByVal deptId As String) As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, COL_ID).End(xlUp).row
    If lastRow < 2 Then Exit Function

    Dim sb As String
    Dim r As Long
    For r = 2 To lastRow
        Dim scope As String: scope = CStr(ws.Cells(r, COL_DEPT_SCOPE).value)
        If LenB(scope) = 0 Then scope = "common"
        If scope = "common" Or scope = deptId Then
            sb = sb & CStr(ws.Cells(r, COL_ID).value) _
                   & " | " & CStr(ws.Cells(r, COL_DOMAIN).value) _
                   & " | " & CStr(ws.Cells(r, COL_DOC_TYPE).value) _
                   & " | " & CStr(ws.Cells(r, COL_DISPLAY).value) _
                   & " | " & CStr(ws.Cells(r, COL_SUMMARY).value) _
                   & " | " & CStr(ws.Cells(r, COL_KEYWORDS).value) _
                   & vbLf
        End If
    Next r
    BuildRouterTable = sb
End Function

' ----------------------------------------------------------------------------
' GetChunkById: returns (display, header, full_text) for one chunk.
' Returns "" entries if not found.
' ----------------------------------------------------------------------------
Public Sub GetChunkById(ByVal chunkId As String, _
                       ByRef display As String, ByRef header As String, _
                       ByRef fullText As String, ByRef source As String, _
                       ByRef domain As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, COL_ID).End(xlUp).row
    Dim r As Long
    For r = 2 To lastRow
        If CStr(ws.Cells(r, COL_ID).value) = chunkId Then
            source = CStr(ws.Cells(r, COL_SOURCE).value)
            display = CStr(ws.Cells(r, COL_DISPLAY).value)
            header = CStr(ws.Cells(r, COL_HEADER).value)
            fullText = CStr(ws.Cells(r, COL_TEXT).value)
            domain = CStr(ws.Cells(r, COL_DOMAIN).value)
            Exit Sub
        End If
    Next r
End Sub

' ----------------------------------------------------------------------------
' BuildContextFromIds: takes a comma-separated list of chunk IDs, returns
' a single context string with [#N] markers, plus a "citations" string
' for display purposes.
' Output context capped at max_chars.
' ----------------------------------------------------------------------------
Public Function BuildContextFromIds(ByVal idsCsv As String, _
                                    ByVal max_chars As Long, _
                                    ByRef citationsOut As String) As String
    Dim ids() As String
    ids = Split(idsCsv, ",")
    Dim sb As String, citSb As String
    Dim used As Long, n As Long
    Dim display As String, header As String, fullText As String, source As String, domain As String

    For n = LBound(ids) To UBound(ids)
        Dim chunkId As String: chunkId = Trim$(ids(n))
        If LenB(chunkId) = 0 Then GoTo NextId
        display = "": header = "": fullText = "": source = "": domain = ""
        GetChunkById chunkId, display, header, fullText, source, domain
        If LenB(fullText) = 0 Then GoTo NextId

        Dim block As String
        block = "[#" & (n + 1) & " " & display & "]" & vbLf & fullText & vbLf & vbLf
        If used + Len(block) > max_chars And used > 0 Then Exit For
        sb = sb & block
        used = used + Len(block)
        citSb = citSb & "[#" & (n + 1) & "] " & display
        If LenB(header) > 0 Then citSb = citSb & "  " & header
        citSb = citSb & vbLf
NextId:
    Next n
    BuildContextFromIds = sb
    citationsOut = citSb
End Function
