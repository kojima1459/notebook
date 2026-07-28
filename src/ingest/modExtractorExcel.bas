Attribute VB_Name = "modExtractorExcel"
Option Explicit

' ============================================================================
' modExtractorExcel - xlsx/xls/xlsm からのテキスト抽出
' ----------------------------------------------------------------------------
' 役割:
'   ファイルを現在のExcelプロセス内で読取専用に開き、ワークシート1枚を
'   ExtractedPageの1「ページ」として発行する。数式は評価済みの表示値のみを
'   拾い、非表示シートも対象に含める(意図的に隠しているファイルの場合は
'   利用者側でファイル選択を絞り込む想定)。
'
' 流用元: /home/user/notebook/src/admin/modExtractorExcel.bas(V2資産・
'   コピー元。変更禁止)。
'
' 適応要件(MASTER_SPEC §7.2):
'   ・V1は Err.Raise で失敗を通知していたが、本モジュールは
'     Boolean 戻り値 + errDetail(ByRef)へ変換する。
'   ・maxPages を超えるシート数のブックは、超過分のシートを開かず先頭
'     maxPages 枚のみ処理し、truncated(ByRef)を True にする。
'   ・Workbooks.Open は必ず On Error で包み、成功・失敗いずれの経路でも
'     wb.Close を試みる(V1のFailedハンドラの作法を踏襲。ブックが開いた
'     ままだとユーザーが元ファイルを別プロセスで触れなくなる事故になる)。
'   ・§12「ループ内Range直接アクセス禁止」に従い、シート内容は
'     `used.Value` で1回のCOM往復にまとめて配列取得してからループする
'     (V1と同じ実装方針)。
' ============================================================================

Public Function Extract(ByVal path As String, ByVal maxPages As Long, _
                        ByRef pages() As ExtractedPage, ByRef truncated As Boolean, _
                        ByRef errDetail As String) As Boolean
    truncated = False

    Dim wb As Workbook
    Dim restoreScreen As Boolean: restoreScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False

    ' 2026-07-28(レビュー H-11): 同じExcelで既に開かれているブックは触らない。
    ' 開いていると Workbooks.Open は既存のブックオブジェクトを返すので、
    ' 抽出後の wb.Close False が【利用者が編集中のブックを破棄して閉じる】
    ' ことになる。%TEMP% コピー経由なら普通は起きないが、コピーに失敗して
    ' 原本を掴んだ場合に直撃する。取り込めないと伝える方が桁違いにましである。
    If IsWorkbookOpen(path) Then
        errDetail = "このファイルは今このExcelで開かれています。閉じてから取り込んでください"
        Application.ScreenUpdating = restoreScreen
        Extract = False
        Exit Function
    End If

    ' 取り込むブックのマクロ(Auto_Open/Workbook_Open)を走らせない。
    ' 3 = msoAutomationSecurityForceDisable。イベントも止める。
    ' どちらも必ず元の値へ戻す(Failed 経路を含む)。
    Dim prevSec As Long: prevSec = -1
    Dim prevEvents As Boolean: prevEvents = Application.EnableEvents
    On Error Resume Next
    prevSec = Application.AutomationSecurity
    Application.AutomationSecurity = 3
    Application.EnableEvents = False
    Err.Clear
    On Error GoTo Failed

    Set wb = Application.Workbooks.Open( _
        FileName:=path, _
        ReadOnly:=True, _
        UpdateLinks:=0, _
        IgnoreReadOnlyRecommended:=True, _
        AddToMru:=False, _
        Password:="__mybookshelf_no_password__", _
        WriteResPassword:="__mybookshelf_no_password__")

    Dim sheetCount As Long: sheetCount = wb.Worksheets.count
    If sheetCount < 1 Then sheetCount = 1

    Dim loopCount As Long: loopCount = sheetCount
    If loopCount > maxPages Then
        loopCount = maxPages
        truncated = True
    End If

    Dim tmp() As ExtractedPage: ReDim tmp(0 To loopCount - 1)
    Dim s As Long
    For s = 1 To loopCount
        Dim ws As Worksheet
        Set ws = wb.Worksheets(s)
        tmp(s - 1).page = s
        tmp(s - 1).Text = ExtractSheetText(ws)
    Next s

    wb.Close False
    RestoreAppState prevSec, prevEvents
    Application.ScreenUpdating = restoreScreen
    Set wb = Nothing

    pages = tmp
    Extract = True
    Exit Function

Failed:
    errDetail = DescribeComError(Err.Number, Err.Description)
    If Not wb Is Nothing Then
        On Error Resume Next
        wb.Close False
        On Error GoTo 0
    End If
    RestoreAppState prevSec, prevEvents
    Application.ScreenUpdating = restoreScreen
    Extract = False
End Function

' AutomationSecurity / EnableEvents を元へ戻す。戻し忘れると、
' このセッションのExcel全体でイベントが死ぬ(他ブックまで巻き添えになる)。
Private Sub RestoreAppState(ByVal prevSec As Long, ByVal prevEvents As Boolean)
    On Error Resume Next
    If prevSec >= 0 Then Application.AutomationSecurity = prevSec
    Application.EnableEvents = prevEvents
    Err.Clear
    On Error GoTo 0
End Sub

' 同名のブックが既にこのExcelで開かれているか。Workbooks(名前) は
' フルパスではなくファイル名で引くため、名前で照合する。
Private Function IsWorkbookOpen(ByVal path As String) As Boolean
    On Error Resume Next
    Dim nm As String: nm = modUtil.FileNameOf(path)
    If LenB(nm) = 0 Then Exit Function
    Dim wbx As Object
    Set wbx = Application.Workbooks(nm)
    IsWorkbookOpen = Not (wbx Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function

' エラー値セルを、画面に出ているのと同じ表記の文字列にする。
' CVErr の値から逆引きする(CStr は型不一致で落ちるため使えない)。
Private Function ErrorCellText(ByVal v As Variant) As String
    On Error Resume Next
    Dim code As Long
    code = CLng(v)                     ' vbError の中身は xlErr* の数値
    Select Case code
        Case 2000: ErrorCellText = "#NULL!"
        Case 2007: ErrorCellText = "#DIV/0!"
        Case 2015: ErrorCellText = "#VALUE!"
        Case 2023: ErrorCellText = "#REF!"
        Case 2029: ErrorCellText = "#NAME?"
        Case 2036: ErrorCellText = "#NUM!"
        Case 2042: ErrorCellText = "#N/A"
        Case 2043: ErrorCellText = "#GETTING_DATA"
        Case Else: ErrorCellText = "#ERR"
    End Select
    Err.Clear
    On Error GoTo 0
End Function

' シート全体を1回のCOM往復(used.Value)で配列取得し、タブ区切り行のテキストにする。
Private Function ExtractSheetText(ByVal ws As Worksheet) As String
    Dim used As Range
    Set used = ws.UsedRange
    If used Is Nothing Then Exit Function
    If used.Cells.count = 0 Then Exit Function

    Dim lastR As Long, lastC As Long
    lastR = used.Rows.count
    lastC = used.Columns.count

    Dim arr As Variant
    arr = used.Value   ' 一括読み取り(§12: ループ内Range直接アクセス禁止)

    Dim lineParts() As String: ReDim lineParts(0 To lastR)
    Dim lineCount As Long: lineCount = 0
    lineParts(0) = "[シート: " & ws.Name & "]"
    lineCount = 1

    Dim r As Long, c As Long
    If IsArray(arr) Then
        For r = 1 To lastR
            Dim cellParts() As String: ReDim cellParts(0 To lastC - 1)
            Dim cellCount As Long: cellCount = 0
            For c = 1 To lastC
                Dim v As Variant: v = arr(r, c)
                ' 2026-07-28(レビュー H-10): エラー値(#N/A/#REF!/#DIV/0! 等)は
                ' Variant の型が vbError で、CStr() が型の不一致(13)を投げる。
                ' 従来はそれで抽出済みシートまで破棄して E0302 で全体失敗して
                ' いた。VLOOKUP の #N/A を含む一覧表は実務で頻出なので、
                ' Excel 取込の実用性を大きく下げていた。しかもエラー詳細から
                ' 原因が読めない。エラー値はセルの見た目どおりの文字列にする
                ' (空にすると列がずれて表の意味が変わるため、残す)。
                If IsError(v) Then
                    cellParts(cellCount) = ErrorCellText(v)
                    cellCount = cellCount + 1
                ElseIf Not IsEmpty(v) Then
                    cellParts(cellCount) = CStr(v)
                    cellCount = cellCount + 1
                End If
            Next c
            If cellCount > 0 Then
                ReDim Preserve cellParts(0 To cellCount - 1)
                lineParts(lineCount) = Join(cellParts, vbTab)
                lineCount = lineCount + 1
            End If
        Next r
    Else
        ' UsedRangeが単一セルの場合、arrはスカラー値になる
        lineParts(lineCount) = CStr(arr)
        lineCount = lineCount + 1
    End If

    ReDim Preserve lineParts(0 To lineCount - 1)
    ExtractSheetText = Join(lineParts, vbLf) & vbLf
End Function

' Mac等COM不可環境向けの丁寧な案内文を生成する(§13)。
Private Function DescribeComError(ByVal errNum As Long, ByVal desc As String) As String
    If errNum = 429 Then
        DescribeComError = "この環境ではExcelブックの読込(COM)が利用できません。" & _
            "Mac版ExcelやCOM未対応環境の可能性があります。Windows版Excelでお試しください。" & _
            "(詳細: " & desc & ")"
    Else
        DescribeComError = desc
    End If
End Function
