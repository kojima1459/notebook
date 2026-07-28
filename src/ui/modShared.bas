Attribute VB_Name = "modShared"
Option Explicit

' ============================================================================
' modShared - 「みんなのQ&A」画面(ナレッジ画面の3つ目のモード)
' ----------------------------------------------------------------------------
' なぜ一覧+選択式なのか:
'   届いたQ&Aを全部自動で本棚に入れる設計は、12,000人規模では必ず破綻する。
'     ・チャンク上限(2万)にすぐ当たる
'     ・自分がとっくに知っている内容まで混ざり、検索の邪魔になる
'     ・毎日「新着◯件」を全部読む運用は、既存の教えてBOXと同じ末路をたどる
'   だから既定は「入れない」。並びは【同じ質問が届いた件数の多い順】にし、
'   多くの人が困っていることほど上に来るようにする。全部読む必要はなく、
'   上から数件見て要るものだけチェックを付ければ終わる。
'
' 操作:
'   行の「□」をクリック → 選択/解除(トグル)
'   ツールバー「選択を取り込む」→ チェックした分だけ本棚へ登録(EXP対象)
'
' 実装:
'   セル主体。1行=1件で、Shapeは行頭のチェックボックスだけ(nxs_接頭辞)。
'   Shape数を一定に保つため、1ページ12件までとしページ送りで見る。
' ============================================================================

Private Const PAGE_SIZE As Long = 12
Private Const SHEET_NAME As String = "Vault"     ' ナレッジ画面と同じシートを使う

Private mPage As Long

' ----------------------------------------------------------------------------
' Show - みんなのQ&A画面を描画して表示する。
' ----------------------------------------------------------------------------
Public Sub Show()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    On Error GoTo Finish
    Application.ScreenUpdating = False

    RemoveRowShapes ws
    ws.Cells.Clear
    ws.Cells.Interior.Color = modUI.UiColor("bg")
    ws.Cells.Font.Name = "Yu Gothic UI"
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 5      ' チェックボックス
    ws.Columns("C").ColumnWidth = 6      ' 件数
    ws.Columns("D:J").ColumnWidth = 12   ' 質問
    ws.Columns("K:N").ColumnWidth = 10   ' 提供者/日付

    modKnowledge.DrawChrome ws, "shared"

    ws.Visible = -1
    ws.Activate

    Dim rows_() As Long
    Dim n As Long
    On Error Resume Next
    n = modInsight.PendingRowsRanked(rows_)
    On Error GoTo Finish

    Dim baseRow As Long: baseRow = modKnowledge.CHROME_ROWS + 1
    DrawHeaderRow ws, baseRow

    If n = 0 Then
        With ws.Range("B" & (baseRow + 2) & ":N" & (baseRow + 12))
            .Merge
            .WrapText = True
            ' 空のときこそ丁寧に。「押したけど何も起きない」が一番の離脱要因。
            .Value = ChrW(&HD83C) & ChrW(&HDF81) & " この画面は「部内のみんなが解決した質問と答え」が集まる場所です。" & vbLf & vbLf & _
                "■ まだ1件も届いていません" & vbLf & _
                "  誰かがチャットで質問し、良い答えが出たときに「" & ChrW(&H2705) & " 解決した」を" & vbLf & _
                "  押すと、その質問と答えがここへ自動で届きます。" & vbLf & vbLf & _
                "■ 届いたら何をするか" & vbLf & _
                "  左の□をクリックして、自分に必要なものだけ選び、" & vbLf & _
                "  上の「" & ChrW(&H2713) & " 選択を取り込む」を押します。" & vbLf & _
                "  取り込んだ内容は、次から自分の質問の答えに使われます。" & vbLf & vbLf & _
                "  ※ 全部取り込む必要はありません。既に知っていることは選ばなくて構いません。"
            .Font.Size = 10
            .Font.Color = modUI.UiColor("muted")
            .VerticalAlignment = -4160
        End With
        GoTo Finish
    End If

    If mPage < 0 Then mPage = 0
    If mPage * PAGE_SIZE >= n Then mPage = 0

    Dim shown As Long
    Dim i As Long
    For i = mPage * PAGE_SIZE To n - 1
        If shown >= PAGE_SIZE Then Exit For
        DrawItemRow ws, baseRow + 2 + shown, rows_(i)
        shown = shown + 1
    Next i

    ' 件数とページ表示
    With ws.Range("B" & (baseRow + 2 + shown + 1) & ":N" & (baseRow + 2 + shown + 1))
        .Merge
        .Value = "全 " & n & " 件中 " & (mPage * PAGE_SIZE + 1) & ChrW(&H301C) & _
                 (mPage * PAGE_SIZE + shown) & " 件を表示 ・ " & _
                 "同じ質問が多い順に並んでいます(上ほど多くの人が困っています)"
        .Font.Size = 8.5
        .Font.Color = modUI.UiColor("muted")
    End With

Finish:
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayWorkbookTabs = False
    modUI.FreezeShapePlacement ws
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Private Sub DrawHeaderRow(ByVal ws As Worksheet, ByVal r As Long)
    ws.Range("B" & r).Value = "選択"
    ws.Range("C" & r).Value = "件数"
    ws.Range("D" & r).Value = "質問(みんなが実際に聞いた内容)"
    ws.Range("K" & r).Value = "解決した人"
    ws.Range("M" & r).Value = "日時"
    With ws.Range("B" & r & ":N" & r)
        .Font.Bold = True
        .Font.Size = 9
        .Font.Color = modUI.UiColor("muted")
    End With
    ws.Rows(r).RowHeight = 18
End Sub

' 1件分の行。行頭のチェックボックスだけShapeで、あとはセル。
Private Sub DrawItemRow(ByVal ws As Worksheet, ByVal r As Long, ByVal srcRow As Long)
    On Error Resume Next
    ws.Rows(r).RowHeight = 30

    Dim cnt As Long
    cnt = modInsight.SameQuestionCount(modInsight.RowField(srcRow, 6))
    If cnt < 1 Then cnt = 1

    ws.Range("C" & r).Value = cnt & "人"
    ws.Range("C" & r).Font.Size = 9
    ws.Range("C" & r).HorizontalAlignment = -4108
    If cnt >= 3 Then
        ws.Range("C" & r).Font.Bold = True
        ws.Range("C" & r).Font.Color = modUI.UiColor("accent")
    Else
        ws.Range("C" & r).Font.Color = modUI.UiColor("muted")
    End If

    With ws.Range("D" & r & ":J" & r)
        .Merge
        .WrapText = True
        .Value = modUtil.SafeLeft(modInsight.RowField(srcRow, 6), 120)
        .Font.Size = 9.5
        .Font.Color = modUI.UiColor("text")
        .VerticalAlignment = -4108
    End With
    With ws.Range("K" & r & ":L" & r)
        .Merge
        .Value = modInsight.RowField(srcRow, 4)
        .Font.Size = 8.5
        .Font.Color = modUI.UiColor("muted")
        .VerticalAlignment = -4108
    End With
    With ws.Range("M" & r & ":N" & r)
        .Merge
        .Value = modInsight.RowField(srcRow, 5)
        .Font.Size = 8.5
        .Font.Color = modUI.UiColor("muted")
        .VerticalAlignment = -4108
    End With

    ' チェックボックス(Shape)。名前に元データの行番号を埋め込み、
    ' クリック時に Application.Caller から復元する。
    Dim cell As Range: Set cell = ws.Range("B" & r)
    Dim d As Double: d = 16
    Dim box As Shape
    Set box = ws.Shapes.AddShape(5, cell.Left + (cell.Width - d) / 2, _
                                 cell.Top + (cell.Height - d) / 2, d, d)
    If box Is Nothing Then Exit Sub
    box.Name = "nxs_ck_" & srcRow
    box.Adjustments(1) = 0.15
    box.Line.Visible = -1
    box.Line.Weight = 1#
    box.Placement = 3

    If modInsight.IsSelected(srcRow) Then
        box.Fill.ForeColor.RGB = modUI.UiColor("accent")
        box.Line.ForeColor.RGB = modUI.UiColor("accent")
        With box.TextFrame2
            .TextRange.Text = ChrW(&H2713)          ' チェックマーク
            .TextRange.Font.Size = 10
            .TextRange.Font.Bold = -1
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = 2
            .VerticalAnchor = 3
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        End With
    Else
        box.Fill.ForeColor.RGB = RGB(255, 255, 255)
        box.Line.ForeColor.RGB = modUI.UiColor("border")
        box.TextFrame2.TextRange.Text = ""
    End If
    box.OnAction = "modShared.OnToggle"
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' ハンドラ
' ----------------------------------------------------------------------------
Public Sub OnToggle()
    Dim caller As String
    On Error Resume Next
    caller = CStr(Application.Caller)
    On Error GoTo 0
    If Left$(caller, 7) <> "nxs_ck_" Then Exit Sub

    On Error Resume Next
    modInsight.ToggleSelected CLng(Val(Mid$(caller, 8)))
    On Error GoTo 0
    Show
End Sub

Public Sub OnSelectAll()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modInsight.SelectAllPending True
    On Error GoTo 0
    modUiLock.Leave
    Show
End Sub

Public Sub OnSelectNone()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modInsight.SelectAllPending False
    On Error GoTo 0
    modUiLock.Leave
    Show
End Sub

Public Sub OnNextPage()
    mPage = mPage + 1
    Show
End Sub

Public Sub OnPrevPage()
    If mPage > 0 Then mPage = mPage - 1
    Show
End Sub

' ----------------------------------------------------------------------------
' 選択したものだけを本棚へ取り込む。
'   既定が「入れない」なので、ここを押した分しかチャンクは増えない。
'   2万チャンクの上限に対して、利用者が意図的に選んだものだけが積まれる。
' ----------------------------------------------------------------------------
Public Sub OnImportSelected()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done

    Dim total As Long
    On Error Resume Next
    total = modInsight.SelectedCount()
    On Error GoTo Done

    If total < 1 Then
        modUiLock.Leave
        MsgBox "取り込むものが選ばれていません。" & vbCrLf & _
               "左の□をクリックして、本棚に入れたいQ&Aを選んでください。", _
               vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    If MsgBox("選択した " & total & " 件を本棚に取り込みます。" & vbCrLf & _
              "(件数によっては1～2分かかります)", _
              vbOKCancel + vbQuestion, modAppDef.APP_NAME) <> vbOK Then GoTo Done

    Dim rows_() As Long
    Dim n As Long
    On Error Resume Next
    n = modInsight.PendingRowsRanked(rows_)
    On Error GoTo Done

    Dim okN As Long, i As Long
    For i = 0 To n - 1
        On Error Resume Next
        Dim rr As Long: rr = rows_(i)
        If modInsight.IsSelected(rr) Then
            If modVault.RegisterKnowledgeText( _
                   "解決済みQ&A: " & modUtil.SafeLeft(modInsight.RowField(rr, 6), 40), _
                   modInsight.QABodyText(modInsight.RowField(rr, 4), _
                                         modInsight.RowField(rr, 6), _
                                         modInsight.RowField(rr, 7), _
                                         modInsight.RowField(rr, 8)), _
                   "解決済みQ&A," & modInsight.RowField(rr, 4)) Then
                modInsight.MarkQAConsumed rr
                okN = okN + 1
            End If
        End If
        On Error GoTo Done
    Next i

    modUiLock.Leave
    MsgBox okN & " 件を本棚に取り込みました。" & vbCrLf & _
           "次から同じ内容を聞かれたとき、確認済みの答えで応えられます。", _
           vbInformation, modAppDef.APP_NAME
    Show
    Exit Sub
Done:
    modUiLock.Leave
End Sub

Private Sub RemoveRowShapes(ByVal ws As Worksheet)
    Dim names() As String
    ReDim names(0 To ws.Shapes.Count)
    Dim n As Long
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 4) = "nxs_" Or Left$(shp.Name, 4) = "nxg_" Then
            names(n) = shp.Name
            n = n + 1
        End If
    Next shp
    Dim i As Long
    For i = 0 To n - 1
        On Error Resume Next
        ws.Shapes(names(i)).Delete
        On Error GoTo 0
    Next i
End Sub
