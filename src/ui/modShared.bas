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

Private mPage As Long

' ----------------------------------------------------------------------------
' Show - みんなのQ&A画面を描画して表示する。
' ----------------------------------------------------------------------------
Public Sub Show()
    ' 2026-07-30(R4要件A): 描画先を実行時生成の "Vault" シートから
    ' 「マイ本棚」へ移した。ギャラリー/一覧表/解決事例はもともと同じ
    ' modShelf.SourceList を見ており、違いは描画先シートだけだった。
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_SHELF)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    On Error GoTo Finish
    Application.ScreenUpdating = False

    ' R21H F7(実機第8報⑦と同型の適用漏れ): 従来は「描く→活性化→表示状態」の
    ' 順で、罫線・見出し・タブ・水平スクロールバーの確定が描画の【後】にあり、
    ' FitBandToViewport/DrawChromeが「タブと横バーが出たままの窓」で幾何を
    ' 決めていた。modVaultGallery.ShowVaultGallery(波A)と同型で
    ' 【活性化→EnsureViewState→表示の共通儀式】を描画の前へ動かす。
    ws.Visible = -1
    If modUI.ActivateSheetRobust(ws, "modShared.Show") Then
        On Error Resume Next
        modViewport2.EnsureViewState ws
        On Error GoTo Finish
        modKnowledge.PrepareScreenView ws
    Else
        modUI.RestoreExcelUI
    End If

    RemoveRowShapes ws
    ws.Cells.Clear
    ' A:N を全列ぶん明示する(モードごとに前提の列幅が違うので、
    ' 前のモードの列幅が残っていると DrawChrome の W がぶれる)。
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 5      ' チェックボックス
    ws.Columns("C").ColumnWidth = 6      ' 件数
    ws.Columns("D:I").ColumnWidth = 12   ' 質問(J列が余りを吸う)
    ws.Columns("J").ColumnWidth = 12
    ws.Columns("K:L").ColumnWidth = 10   ' 提供者
    ws.Columns("M").ColumnWidth = 12     ' 日付
    ws.Columns("N").ColumnWidth = 1      ' 吸収列(最小のまま)
    ' R20-1d: 毎回406行(7:412)を書き直すのをやめ、前回使った行までに絞る。
    modKnowledge.NormalizeShelfRows ws, 7
    ' R18-3a: 全域(ws.Cells)への書式はUsedRangeをシート最大へ膨らませる
    ' (無限スクロールの主因・調査agent2 §1.3)。実使用範囲だけに当てる。
    ' R19-1b: 一覧は1ページぶん(PAGE_SIZE行)なので1画面ぶんで足りる。範囲は
    ' 行高を戻した【後】に決める(前モードの行高のまま数えると足りなくなる)。
    Dim bandAddr As String: bandAddr = modKnowledge.ShelfBound(ws, 0)
    ws.Range(bandAddr).Interior.Color = modUI.UiColor("bg")
    ws.Range(bandAddr).Font.Name = "Yu Gothic UI"

    ' R21H F7: 質問列(D:J)に余りを吸わせる外部Fit(J)を先に打ち、直後の
    ' DrawChromeが吸収列Nで【二重に】Fitし直していた(2回目の丸めで帯が
    ' 可視幅を数pt超える経路)。波AのtableのFit(J)一発と同型にするため、
    ' 外部Fitを廃止しmodViewport2.ShelfPadColを"shared"もJへ倒す
    ' (DrawChrome内の1回のFitだけで質問列へ余りが吸われる)。
    modKnowledge.DrawChrome ws, "shared"

    Dim rows_() As Long
    Dim n As Long
    On Error Resume Next
    n = modInsight.PendingRowsRanked(rows_)
    On Error GoTo Finish

    Dim baseRow As Long: baseRow = modKnowledge.CHROME_ROWS + 1
    ' R19H FA-2(A-H②): 描き終えた実下端。塗り・境界をここから敷き直す
    ' (末尾の FinishCleanup0 で ApplyExtent へ渡す)。
    Dim botRow As Long: botRow = baseRow
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
        botRow = baseRow + 12
        ' 2026-07-30(レビュー5-A): ここは正常系(0件の案内を出しただけ)。
        ' Finish: はハンドラ本体で、先頭が Resume なのでエラーが起きていない
        ' 状態で踏むと実行時エラー20「Resume にエラーがありません」になる。
        ' 正常系は後始末ラベルへ直行し、Resume を跨ぐ(R6の標準形)。
        GoTo FinishCleanup0
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
    botRow = baseRow + 2 + shown + 1
    With ws.Range("B" & botRow & ":N" & botRow)
        .Merge
        .Value = "全 " & n & " 件中 " & (mPage * PAGE_SIZE + 1) & ChrW(&H301C) & _
                 (mPage * PAGE_SIZE + shown) & " 件を表示 ・ " & _
                 "同じ質問が多い順に並んでいます(上ほど多くの人が困っています)"
        .Font.Size = 8.5
        .Font.Color = modUI.UiColor("muted")
    End With

    ' 正常系はハンドラ本体(Resume)を跨いで後始末へ入る
    ' (Resume はエラーが起きていないと実行時エラー20になる)。
    GoTo FinishCleanup0
Finish:
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FinishCleanup0
FinishCleanup0:
    On Error Resume Next
    ApplyExtent ws, botRow
    ' 2026-07-31(R11-B H-1): ActiveWindow系は、そのシートが実際に前面の
    ' ときだけ触る(別シートの表示状態を巻き添えで変えないため)。
    If ThisWorkbook.ActiveSheet Is ws Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.DisplayWorkbookTabs = False
    End If
    modUI.FreezeShapePlacement ws
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' ApplyExtent - 描き終えた実下端(件数行)から塗りと境界を敷き直す(R19H FA-2)。
' ----------------------------------------------------------------------------
' modUIShelf.ApplyShelfExtent と同型。Show の冒頭で敷く帯は「まだ何も描いて
' いない時点の見積り(1画面ぶん)」で、1ページ12件の一覧+件数行がそれより
' 下へ伸びる狭い窓では、帯の外に文字が残る ―― そこは塗られておらず
' ScrollArea の外でもあるので、【見えているのに到達できない白い断崖】になる
' (A-H②)。行高が確定した後に実測(Top+Height)から敷き直せば必ず届く。
' 範囲の式は modKnowledge.ShelfBound が単一情報源(3モードで同じ算数を書かない)。
Private Sub ApplyExtent(ByVal ws As Worksheet, ByVal lastRow As Long)
    If ws Is Nothing Then Exit Sub
    If lastRow < 1 Then Exit Sub
    On Error Resume Next
    Dim bottomY As Double
    bottomY = ws.Rows(lastRow).Top + ws.Rows(lastRow).Height
    ' R20-1d: 塗り・ScrollArea・境界より下の行高リセットは modKnowledge へ集約。
    modKnowledge.ApplyShelfBound ws, bottomY
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
    If Left$(caller, 7) <> "nxs_ck_" Then
        modLog.LogUsage "caller_mismatch", "modShared.OnToggle", caller
        Exit Sub
    End If

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
    ' 2026-07-31(R11-B M-6): 旧実装は最終ページで押しても無言で先頭(1ページ目)
    ' へワープしていた(Show内のmPage*PAGE_SIZE>=n防御による副作用)。押した
    ' のに反応が見えない=無反応と同義(憲章§3-1)なので、ここで先読みして
    ' 「最後のページです」を明示する。
    Dim rows_() As Long
    Dim n As Long
    On Error Resume Next
    n = modInsight.PendingRowsRanked(rows_)
    On Error GoTo 0
    If (mPage + 1) * PAGE_SIZE >= n Then
        ' waitless: 端で押しただけの案内に1.1秒待たせない(R11-H Med4)。
        modSkin.ShowToast "最後のページです。", "info", True
        Exit Sub
    End If
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

    ' 2026-07-31(R11-E H-3): 件数によっては1～2分かかると案内しているのに
    ' 進捗が一切見えていなかった(監査1)。取込済み/選択総数のETA付き
    ' バナーを1件ごとに更新する(表示失敗が取込を止めないようOERNで包む)。
    ' 2026-07-31(R11-H Med2): 分母を「保留Q&Aの全行数 n」ではなく
    ' 「選ばれた件数 total」にした。従来は20件中3件だけ選んでも
    ' 「1/20」から始まり、3件目で終わって「3/20」で消える=数字が嘘になり、
    ' 残り時間も約6倍に見積もられていた。分子は処理済みの選択数、
    ' ETAも同じ基準(1件あたり実測×残りの選択数)で出す。
    Dim okN As Long, i As Long
    Dim doneN As Long: doneN = 0
    Dim tStart As Double: tStart = Timer
    For i = 0 To n - 1
        On Error Resume Next
        Dim rr As Long: rr = rows_(i)
        If modInsight.IsSelected(rr) Then
            Dim etaPart As String: etaPart = ""
            If doneN > 0 Then
                Dim elapsedMs As Double: elapsedMs = modUtilText.ElapsedMsSince(tStart)
                etaPart = modUtil.EtaText(total - doneN, elapsedMs / doneN)
            End If
            modUIMain.ShowProgress modUtil.ProgressText(doneN + 1, total, etaPart) & " Q&Aを取り込み中…"
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
            doneN = doneN + 1
        End If
        On Error GoTo Done
    Next i
    On Error Resume Next
    modUIMain.HideProgress
    On Error GoTo Done

    modUiLock.Leave
    MsgBox okN & " 件を本棚に取り込みました。" & vbCrLf & _
           "次から同じ内容を聞かれたとき、確認済みの答えで応えられます。", _
           vbInformation, modAppDef.APP_NAME
    Show
    Exit Sub
Done:
    ' modUIMain.HideProgress自体が内部でOERN込みの薄いラッパーなので、
    ' ここで新たにOn Errorを書く必要が無い(書くとハンドラ稼働中の
    ' On Error Resume Nextが効かない問題を踏む)。
    modUIMain.HideProgress
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
