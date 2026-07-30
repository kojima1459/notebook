Attribute VB_Name = "modExtractorExcel"
Option Explicit

' 暗号化ブックでパスワード入力ダイアログを出さないためのダミー。
' 正しいはずがないので、保護ブックなら即エラーになり E0302 として扱える。
Private Const DUMMY_PASSWORD As String = "__mybookshelf_no_password__"

' パスワード引数つきで開けるか。0=未確定 / 1=引数つきでOK / 2=引数なしでOK。
' 実機で通った方をおぼえて、毎ファイル無駄打ちしないようにする。
Private mPwArgWorks As Long

' 2026-07-29(実機事故): 大きな .xlsx で E0302「メモリが不足しています」。
' 原因は used.Value による【シート全体の一括読み】。UsedRange が
' 100万行×数十列まで伸びた業務ブック(書式や空白セルの残骸で簡単にそうなる)
' では、Variant 配列だけで数百MBからGB級になり、32bit Excel の 2GB 空間を
' 使い切って落ちる。しかも落ちた時点でそのファイルの取込は全滅する。
'
' ここで「1回のCOM往復にまとめる」(§12)を、そのまま守り切ることはできない。
' §12 の狙いはセル1個ずつのCOM呼び出しを避けることであって、
' サイズ無制限の一括読みを保証することではない。
' そこで【行ブロック単位の一括読み】に変える。COM往復は行数/ブロック行数回
' で済み(1万行でも数十回)、1回あたりの確保量には上限がかかる。
'
' 併せて、取り込む総セル数と総文字数にも上限を置く。埋め込みに使うのは
' 要約可能な平文であって、100万行の明細そのものではない。上限で切ったことは
' 本文の末尾に明記して、利用者が「全部入っている」と誤解しないようにする。
Private Const MAX_CELLS_PER_READ As Long = 100000   ' 1回の一括読みの上限セル数
Private Const MAX_CELLS_PER_SHEET As Long = 400000  ' 1シートから取り込む上限セル数
Private Const MAX_COLS_PER_SHEET As Long = 256      ' 横方向の上限(これ以上は表ではない)
' 1シートの本文の上限文字数。セル数だけを見ていると、1セルに数千文字ある
' 備考欄のシートで結局メモリを使い切る。ここを超えたら読むのをやめる
' (この文字数でも既にチャンクは数百個になる。埋め込み費用の歯止めも兼ねる)。
Private Const MAX_CHARS_PER_SHEET As Long = 300000

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

    ' 2026-07-29(実機事故): .xlsx / .xlsm の取込が
    ' 「'Open' メソッドは失敗しました: 'Workbooks' オブジェクト」で全滅した。
    ' 2026-07-28 に足した Password / WriteResPassword 引数が環境によっては
    ' Open 自体を失敗させる。パスワード入力ダイアログでの停止を防ぐための
    ' 引数(レビュー M-11)が、保護されていない普通のブックまで開けなくして
    ' いては本末転倒である。
    '
    ' 引数付きで開いてみて、駄目なら引数なしでもう一度開く。
    ' 暗号化ブックは引数なしだとダイアログで止まってしまうので、
    ' 「引数付きで失敗 → 引数なしでも失敗」なら素直に取込失敗として返す
    ' (M-11 の意図はここで保たれる)。
    ' どちらで通ったかはセッション中おぼえて、次から無駄打ちしない。
    Set wb = OpenForExtract(path)
    If wb Is Nothing Then Err.Raise 1004, "modExtractorExcel", _
        "ブックを開けませんでした(パスワード保護、または破損の可能性)"

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
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailedCleanup0
FailedCleanup0:
    If Not wb Is Nothing Then
        On Error Resume Next
        wb.Close False
        On Error GoTo 0
    End If
    RestoreAppState prevSec, prevEvents
    Application.ScreenUpdating = restoreScreen
    Extract = False
End Function

' ----------------------------------------------------------------------------
' OpenForExtract - 抽出用にブックを開く。開けなければ Nothing。
'   1回目: パスワード引数つき(暗号化ブックでダイアログを出さないため)
'   2回目: パスワード引数なし(1回目の引数が原因で開けない環境への保険)
'   一度通った方をおぼえて、次からそちらを先に試す。
' ----------------------------------------------------------------------------
Private Function OpenForExtract(ByVal path As String) As Workbook
    Dim order(0 To 1) As Boolean
    If mPwArgWorks = 2 Then
        order(0) = False: order(1) = True      ' 引数なしが通る環境
    Else
        order(0) = True: order(1) = False      ' 既定は引数つきから
    End If

    Dim i As Long
    For i = 0 To 1
        Dim wb As Workbook
        Set wb = Nothing
        On Error Resume Next
        Err.Clear
        If order(i) Then
            Set wb = Application.Workbooks.Open( _
                FileName:=path, ReadOnly:=True, UpdateLinks:=0, _
                IgnoreReadOnlyRecommended:=True, AddToMru:=False, _
                Password:=DUMMY_PASSWORD, WriteResPassword:=DUMMY_PASSWORD)
        Else
            Set wb = Application.Workbooks.Open( _
                FileName:=path, ReadOnly:=True, UpdateLinks:=0, _
                IgnoreReadOnlyRecommended:=True, AddToMru:=False)
        End If
        Dim openErr As Long: openErr = Err.Number
        Err.Clear
        On Error GoTo 0

        If openErr = 0 And Not wb Is Nothing Then
            Dim mark As Long: mark = IIf(order(i), 1, 2)
            If mPwArgWorks <> mark Then
                mPwArgWorks = mark
                On Error Resume Next
                modLog.LogUsage "excel_open_mode", "", _
                    "この端末では" & IIf(order(i), "パスワード引数つき", "パスワード引数なし") & _
                    "でブックを開けました"
                On Error GoTo 0
            End If
            Set OpenForExtract = wb
            Exit Function
        End If
    Next i
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

' シートを行ブロック単位で一括読みし、タブ区切り行のテキストにする。
' 上限の考え方は本モジュール冒頭の宣言部を参照。
Private Function ExtractSheetText(ByVal ws As Worksheet) As String
    Dim used As Range
    Set used = ws.UsedRange
    If used Is Nothing Then Exit Function
    If used.Cells.count = 0 Then Exit Function

    Dim lastR As Long, lastC As Long
    lastR = used.Rows.count
    lastC = used.Columns.count
    If lastR < 1 Or lastC < 1 Then Exit Function

    Dim clipped As Boolean
    If lastC > MAX_COLS_PER_SHEET Then
        lastC = MAX_COLS_PER_SHEET
        clipped = True
    End If

    Dim maxRows As Long: maxRows = MAX_CELLS_PER_SHEET \ lastC
    If maxRows < 1 Then maxRows = 1
    Dim readRows As Long: readRows = lastR
    If readRows > maxRows Then
        readRows = maxRows
        clipped = True
    End If

    Dim blockRows As Long: blockRows = MAX_CELLS_PER_READ \ lastC
    If blockRows < 1 Then blockRows = 1

    Dim lineParts() As String: ReDim lineParts(0 To readRows + 1)
    Dim lineCount As Long
    lineParts(0) = "[シート: " & ws.Name & "]"
    lineCount = 1

    Dim r0 As Long, failCount As Long, charCount As Long
    r0 = 1
    Do While r0 <= readRows
        Dim blkRows As Long: blkRows = blockRows
        If r0 + blkRows - 1 > readRows Then blkRows = readRows - r0 + 1

        Dim arr As Variant
        If ReadBlock(used, r0, blkRows, lastC, arr) Then
            charCount = charCount + AppendBlockLines(arr, blkRows, lastC, lineParts, lineCount)
            If charCount >= MAX_CHARS_PER_SHEET Then
                clipped = True
                r0 = r0 + blkRows
                Exit Do
            End If
        Else
            ' このブロックだけ諦める。ここで全体を失敗させると、
            ' 壊れた1箇所のせいでシート全部が捨てられる。
            clipped = True
            failCount = failCount + 1
            ' 何度も失敗するなら、そのシートはもう読めない。1行ずつ
            ' 失敗し続けて何十万回も回るのを防ぐ。
            If failCount >= 5 Then Exit Do
        End If
        r0 = r0 + blkRows
    Loop

    If clipped Then
        lineParts(lineCount) = "[※このシートは大きいため一部のみ取り込みました]"
        lineCount = lineCount + 1
    End If

    ReDim Preserve lineParts(0 To lineCount - 1)
    ExtractSheetText = Join(lineParts, vbLf) & vbLf
End Function

' 行ブロックを1回のCOM往復で読む。確保に失敗したらブロックを半分にして
' もう一度だけ試す(空きメモリが足りないだけなら、これで通ることが多い)。
Private Function ReadBlock(ByVal used As Range, ByVal r0 As Long, ByRef blkRows As Long, _
                           ByVal cols As Long, ByRef arr As Variant) As Boolean
    Dim attempt As Long
    For attempt = 1 To 2
        On Error Resume Next
        Err.Clear
        arr = used.Cells(r0, 1).Resize(blkRows, cols).Value
        Dim e As Long: e = Err.Number
        Err.Clear
        On Error GoTo 0
        If e = 0 Then
            ReadBlock = True
            Exit Function
        End If
        If blkRows <= 1 Then Exit For
        blkRows = blkRows \ 2
        If blkRows < 1 Then blkRows = 1
    Next attempt
End Function

' 読み込んだブロックを行テキストへ変換して lineParts に積む。
' 戻り値は積んだ文字数(呼び出し元の文字数上限の判定に使う)。
Private Function AppendBlockLines(ByRef arr As Variant, ByVal blkRows As Long, ByVal cols As Long, _
                                  ByRef lineParts() As String, ByRef lineCount As Long) As Long
    Dim added As Long
    If Not IsArray(arr) Then
        ' 1セルだけの範囲は Value がスカラーになる。
        If Not IsEmpty(arr) Then
            If IsError(arr) Then
                lineParts(lineCount) = ErrorCellText(arr)
            Else
                lineParts(lineCount) = CStr(arr)
            End If
            added = Len(lineParts(lineCount))
            lineCount = lineCount + 1
        End If
        AppendBlockLines = added
        Exit Function
    End If

    Dim r As Long, c As Long
    For r = 1 To blkRows
        Dim cellParts() As String: ReDim cellParts(0 To cols - 1)
        Dim cellCount As Long: cellCount = 0
        For c = 1 To cols
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
            added = added + Len(lineParts(lineCount)) + 1
            lineCount = lineCount + 1
        End If
    Next r
    AppendBlockLines = added
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
