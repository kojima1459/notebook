Attribute VB_Name = "modUsageAggregator"
Option Explicit

' ============================================================================
' modUsageAggregator - merge per-user CSVs into a single rollup sheet
' ----------------------------------------------------------------------------
' Reads every *.csv under [logging] remote_url, parses UTF-8 with BOM, and
' writes the union into the "Usage" sheet. From there the admin can pivot
' with native Excel features (no Power Query required).
'
' We intentionally do not delete or rewrite the source CSVs - users keep
' appending to them between aggregations. Re-running this Sub fully rebuilds
' the Usage sheet from scratch.
' ============================================================================

Private Const SHEET_NAME As String = "Usage"

Public Sub RefreshUsageSheet()
    modConfig.EnsureLoaded
    Dim folder As String: folder = modConfig.GetString("logging", "remote_url", "")
    If LenB(folder) = 0 Then
        MsgBox "config.ini の [logging] remote_url が設定されていません。", vbExclamation
        Exit Sub
    End If

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then
        MsgBox "ログフォルダが見つかりません: " & folder, vbExclamation
        Exit Sub
    End If

    Dim sh As Worksheet
    Set sh = GetOrCreateSheet(SHEET_NAME)
    sh.Cells.Clear
    Application.ScreenUpdating = False

    Dim headerRow As Variant
    headerRow = Array("timestamp", "username", "department", "role", _
                      "prompt_tokens", "completion_tokens", "latency_ms", _
                      "pii_warn", "has_question_text", "question_text", "source_file")
    Dim col As Long
    For col = 0 To UBound(headerRow)
        sh.Cells(1, col + 1).value = headerRow(col)
    Next col
    sh.Range(sh.Cells(1, 1), sh.Cells(1, UBound(headerRow) + 1)).Font.Bold = True

    Dim destRow As Long: destRow = 2
    Dim f As Object
    For Each f In fso.GetFolder(folder).Files
        If LCase$(fso.GetExtensionName(f.Name)) = "csv" Then
            Application.StatusBar = "読込中: " & f.Name
            AppendCsvRows sh, destRow, f.path, f.Name
            DoEvents
        End If
    Next f

    Application.StatusBar = "集計シート: " & (destRow - 2) & " 行"
    sh.Columns.AutoFit
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox (destRow - 2) & " 件の利用ログを取り込みました。", vbInformation
End Sub

Private Sub AppendCsvRows(ByVal sh As Worksheet, _
                          ByRef destRow As Long, _
                          ByVal csvPath As String, _
                          ByVal fileName As String)
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    On Error GoTo Failed
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile csvPath
    Dim isFirst As Boolean: isFirst = True
    Do While Not st.EOS
        Dim line As String
        line = st.ReadText(-2) ' adReadLine
        If LenB(line) = 0 Then GoTo NextLine
        If isFirst Then
            isFirst = False
            ' Skip header (with or without BOM)
            GoTo NextLine
        End If
        Dim fields() As String
        fields = ParseCsvLine(line)
        Dim k As Long
        For k = 0 To UBound(fields)
            sh.Cells(destRow, k + 1).value = fields(k)
        Next k
        sh.Cells(destRow, 11).value = fileName
        destRow = destRow + 1
NextLine:
    Loop
    st.Close
    Exit Sub
Failed:
    On Error Resume Next
    st.Close
    On Error GoTo 0
    Debug.Print "Skipped " & csvPath & ": " & Err.Description
End Sub

' RFC-4180-ish CSV parsing: handles quoted fields with embedded commas and
' doubled-quote escaping. Newlines inside quoted fields not supported (the
' logger never produces them - vbLf in question text is escaped at write time).
Private Function ParseCsvLine(ByVal line As String) As String()
    Dim out() As String
    ReDim out(0 To 16)
    Dim count As Long: count = 0
    Dim i As Long, n As Long: n = Len(line)
    Dim cur As String
    Dim inQ As Boolean
    For i = 1 To n
        Dim ch As String: ch = Mid$(line, i, 1)
        If inQ Then
            If ch = """" Then
                If i < n And Mid$(line, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & ch
            End If
        Else
            If ch = """" Then
                inQ = True
            ElseIf ch = "," Then
                If count > UBound(out) Then ReDim Preserve out(0 To count * 2)
                out(count) = cur
                count = count + 1
                cur = ""
            Else
                cur = cur & ch
            End If
        End If
    Next i
    If count > UBound(out) Then ReDim Preserve out(0 To count)
    out(count) = cur
    count = count + 1
    ReDim Preserve out(0 To count - 1)
    ParseCsvLine = out
End Function

Private Function GetOrCreateSheet(ByVal name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets( _
            ThisWorkbook.Worksheets.count))
        ws.Name = name
    End If
    Set GetOrCreateSheet = ws
End Function
