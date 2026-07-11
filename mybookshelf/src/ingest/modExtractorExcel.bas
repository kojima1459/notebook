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

    On Error GoTo Failed
    Set wb = Application.Workbooks.Open( _
        FileName:=path, _
        ReadOnly:=True, _
        UpdateLinks:=0, _
        IgnoreReadOnlyRecommended:=True, _
        AddToMru:=False)

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
    Application.ScreenUpdating = restoreScreen
    Extract = False
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
                If Not IsEmpty(v) Then
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
