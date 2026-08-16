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
    modUiLock.AlertsOff   ' R25-1a-1(対はFinishCleanup0)

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
    ' R27 F2-2: 帯(A:N)の右外も幅1へ(table/galleryと同じ手当て)。
    modChrome.SetupShelfColumns ws
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
                "  ※ 全部取り込む必要はありません。既に知っていることは選ばなくて構いません。" & vbLf & _
                "  ※ あなたが解決した事例は、ここ(自分の画面)には表示されません。仲間のPCの" & vbLf & _
                "    この一覧にだけ届きます。一般アシスタント(RAGオフ)の回答は共有対象外です。"
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
    modUiLock.AlertsOn
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
' RedrawCurrentIfFront - R33 W5-5: 「描き直し」が「画面遷移」に化けるのを止める。
' ----------------------------------------------------------------------------
' modUIShelf.RenderShelf は一覧表モード以外だと modKnowledge.RefreshCurrent へ
' 委譲する。ところが3モードの再描画実体は対称ではない ―― table(RenderShelf 本体)
' はセルを書くだけなのに、gallery(modVaultGallery.ShowVaultGallery:102-104)と
' shared(modShared.Show:51-52)は無条件に ws.Visible=-1 + ActivateSheetRobust を
' 実行する「必ず前面へ出す」経路である。そのため、チャットで作業している最中の
' 取込(👎修正の登録・洞察カードの保存)や自動同期タイマーが、利用者が何も押して
' いないのに画面を本棚へ引きずり出していた(mMode は Hub の📚から入ると gallery
' のまま残るので、既定の入り方をした端末では常に起きる)。
' 委譲するのは【マイ本棚シートが既に前面のとき】だけにする。前面でないときは
' 描かない ―― 描画は冪等なので、次にその画面を開いた時点で必ず追いつく。
' 実体をここに置くのは modUIShelf が残84字で分岐を書けないため(R33容量裁定)。
Public Sub RedrawCurrentIfFront()
    On Error Resume Next
    Dim activeName As String
    activeName = Application.ActiveSheet.Name
    If activeName = modAppDef.SH_SHELF Then modKnowledge.RefreshCurrent
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
    If mPage <= 0 Then
        ' R30 W2-4(おもてなし): modVaultGallery.OnVaultPrevと同型。端で無反応に
        ' しない(waitless: 端の案内に1.1秒待たせない。R11-H Med4)。
        modSkin.ShowToast "最初のページです。", "info", True
        Exit Sub
    End If
    mPage = mPage - 1
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
            ' R33 W5-19: silent:=True。ここは共有モードの画面を表示したまま
            ' 回すループなので、1件ごとの自動再描画を許すと【この画面自身】が
            ' Shape全削除→Cells.Clear→クロム再生成→ScreenUpdating復帰まで
            ' 丸ごと作り直され、件数ぶん画面が明滅して行が抜けたり増えたりする。
            ' ループを抜けた後の Show(:454)が1回だけ描き直す。
            ' modShelfSync/modShelfBatch が既に採っている形へ揃える。
            If modVault.RegisterKnowledgeText( _
                   "解決済みQ&A: " & modUtil.SafeLeft(modInsight.RowField(rr, 6), 40), _
                   modInsight.QABodyText(modInsight.RowField(rr, 4), _
                                         modInsight.RowField(rr, 6), _
                                         modInsight.RowField(rr, 7), _
                                         modInsight.RowField(rr, 8)), _
                   "解決済みQ&A," & modInsight.RowField(rr, 4), "", True) Then
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

' ============================================================================
' 💡 みんなの困りごと: 組織で答えが見つからなかった質問の一覧。
'   資料を書ける人(商品部)がここを見て、その場でナレッジを書けるようにする。
'   営業の「分からない」が、商品部の「書くべきこと」に直結する一番短い経路。
' ----------------------------------------------------------------------------
' 2026-08-14(R32 r1): 実体を modKnowledge.OnGapBoard からここへ移した。
'   移設理由は容量(modKnowledge は残り146字で1行も入らない)だが、置き場所
'   として本モジュールを選んだのは、「みんなのQ&A」と困りごと板が共有知
'   フライホイールの表と裏(届いた答え/足りない答え)で、読む先も
'   modInsight の同じ受信箱シートだから。
'   移設のついでに、旧実装が持っていた「Yesを選んだ経路で modUiLock.Leave を
'   2回通りうる」構造(Leave の直後に Exit Sub を置き、他方は Done: ラベルへ
'   落ちる書き方)を、Leave が1箇所しかない形へ組み直した。
' ============================================================================
'   2026-08-14(R32 F9): Enter のあとは【必ず】Leave を通る形へ組み直した。
'   移設時、GapListText/GapCount だけを On Error Resume Next で包み、その後の
'   MsgBox 2本と Leave を素で並べていたため、そこで例外が出るとUIロックを
'   掴んだまま抜け、以後この画面もHubのボタンも一切反応しなくなる
'   (旧 OnGapBoard は全体を包んでいたので、移設で後退していた)。
'   Leave が1箇所という R32 r1 の形はそのまま保つ(Done: ラベルへ一本化)。
Public Sub ShowGapBoard()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub

    Dim body As String, n As Long
    Dim goRegister As Boolean
    Dim errNum As Long, errDesc As String
    On Error GoTo Failed
    body = modInsight.GapListText()
    n = modInsight.GapCount()

    If n = 0 Then
        ' R32 m1: 0件のときに「今すぐ登録しますか?」のYes/Noを出さない。
        ' 自分の投稿は自分の板に出ない(IsMine)ので、1人で試している間
        ' 【必ずこの画面になる】。ここで答えようのないYes/Noを出すことが、
        ' この機能の第一印象そのものを壊していた。
        MsgBox body, vbInformation, _
               modAppDef.APP_NAME & " - みんなの困りごと (0件)"
    Else
        ' 【注意】この前後の案内文は【実測162字】(前79字+後83字)で、
        ' modInsight.GapListBuild の上限820字はここから逆算している
        ' (1,024 − 162 = 862 が本文の上限)。文言を増やすときは、あちらの
        ' MAX_CHARS も同じだけ下げること。増やしたまま放置すると、末尾の
        ' 「今すぐ登録しますか?」がMsgBoxから無言で消える(R32 F3)。
        goRegister = (MsgBox( _
            "みんなが質問して、本棚に答えが無かった質問です(新しい順・最大20件)。" & vbCrLf & _
            "ここに並ぶ質問に答える資料を用意すると、部内の全員がすぐ答えを得られます。" & vbCrLf & vbCrLf & _
            body & vbCrLf & _
            "この内容に答える資料を、今すぐ登録しますか?" & vbCrLf & _
            "(「はい」で登録画面が開きます。上の一覧から答えられるものを1つ選び、" & vbCrLf & _
            " その質問への答えとして書いてください)", _
            vbYesNo + vbInformation, _
            modAppDef.APP_NAME & " - みんなの困りごと (" & n & "件)") = vbYes)
    End If

Done:
    On Error GoTo 0
    ' R32マイクロ修正波 F19[MINOR]: Failedからここへ落ちたときだけ、
    ' 押しても無反応だった穴を埋める(旧実装は On Error Resume Next で
    ' 空のMsgBoxだけは出ていたが、Resume Done一本化(R32 F9)でそれも消えた)。
    ' errNum はFailed:で退避済みの値(On Error GoTo 0はErrをリセットするため、
    ' ここで読み直すと0になる。CLAUDE.mdのErr退避作法どおり先に読む)。
    If errNum <> 0 Then
        On Error Resume Next
        modLog.LogError "E0801", "modShared.ShowGapBoard", _
            "GapListText/GapCountの取得に失敗しました err=" & errDesc, errNum
        modSkin.ShowToast "みんなの困りごとの表示に失敗しました。もう一度お試しください。", "error"
        On Error GoTo 0
    End If
    modUiLock.Leave
    If goRegister Then modVault.ShowVaultInput
    Exit Sub

Failed:
    errNum = Err.Number
    errDesc = Err.Description
    ' Resume でハンドラを抜けてから後始末へ落ちる(ハンドラ稼働中は
    ' 同一プロシージャで次のエラーを捕まえられないため。vba_lint が検査)。
    Resume Done
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

' ============================================================================
' OnPurgeChannel - 🗑部門の資料(R33 W5-23。ナレッジ画面のツールバー)
' ----------------------------------------------------------------------------
' なぜ必要か:
'   部門の購読をやめても、その部門の資料は本棚に残り続けていた。消す関数
'   (modChannel.PurgeChannelChunks)は前から在ったのに呼び出し元が0件で、
'   実際の解除手段も config の手編集しか無かった(R33波2の差し戻し根拠)。
'   「解除経路から消す」を結線しても死にコードに死にコードを繋ぐだけなので、
'   利用者が自分で消せる入口を画面側に作る、というのが今回の裁定。
'
' 何を消して、何を消さないか(この線引きが本機能の全て):
'   消す   … origin が "channel:<選んだ部門名>" と【完全一致】する行だけ
'   消さない … 他部門(channel:<別名>) / 自作(self) / 手渡しパック(pack:…)
'   一致判定の実体は modShelfStore.RemoveRowsByOrigin(前置き一致ではなく
'   完全一致)で、前置き一括削除の RemoveRowsByOriginPrefix は使わない。
'
' 作法:
'   ・押しただけでは消さない。件数を見せ、「取り消せません」を明示した
'     Yes/No を挟む(既定はNo側=vbDefaultButton2)。
'   ・削除後の再描画は最後に1回だけ(波5b W5-19 と同じ。1件ごとに描かない)。
'   ・消した/消さなかったに関わらず usage_log へ1行残す(無言で消さない)。
'   ・my_stats の取込済み版(ch:<部門名>)も空へ戻す。中身を消したのに
'     「この部門の版Xを持っている」と記録が残ると、部門チャンネルを押しても
'     「最新です」と判定されて二度と戻せなくなる(SubscribeAllAvailable は
'     RemoteVersion<>LocalVersion のときだけ取り込む)。
' ============================================================================
Public Sub OnPurgeChannel()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub

    Dim counts As String, pick As String, ans As String
    Dim chName As String, shown As Long, removed As Long
    Dim errNum As Long, errDesc As String
    Dim redraw As Boolean
    Dim title As String
    On Error GoTo Failed
    title = modAppDef.APP_NAME & " - 部門の資料を削除"

    counts = modShelfStore.OriginCountsByPrefix("channel:")
    If LenB(counts) = 0 Then
        MsgBox "部門から取り込んだ資料は、いま本棚にありません。" & vbCrLf & vbCrLf & _
               "(ここで消せるのは部門の公式ナレッジだけです。ご自分で登録・追加した" & vbCrLf & _
               " 資料と、パックで受け取った資料は対象になりません)", _
               vbInformation, title
        GoTo Done
    End If

    ans = Trim$(InputBox( _
        "本棚に入っている部門の資料です。削除したい部門の番号を入力してください。" & vbCrLf & vbCrLf & _
        modShareRule.PurgeMenuText(counts) & vbCrLf & vbCrLf & _
        "消えるのは選んだ部門の資料だけです。" & vbCrLf & _
        "ご自分で登録・追加した資料、パックで受け取った資料、ほかの部門の資料は残ります。", _
        title))
    If LenB(ans) = 0 Then GoTo Done

    pick = ""
    If IsNumeric(ans) Then pick = modShareRule.PurgeMenuPick(counts, CLng(Val(ans)))
    If LenB(pick) = 0 Then
        MsgBox "一覧に無い番号です。もう一度、一覧の番号を入力してください。", _
               vbInformation, title
        GoTo Done
    End If

    Dim kv() As String: kv = Split(pick, vbTab)
    chName = kv(0)
    If UBound(kv) >= 1 Then shown = CLng(Val(kv(1)))

    If MsgBox("「" & chName & "」の資料 " & shown & "件を本棚から削除します。" & vbCrLf & vbCrLf & _
              "この操作は取り消せません。" & vbCrLf & _
              "(もう一度必要になったときは、上のツールバーの「部門チャンネル」" & vbCrLf & _
              " から読み込み直せます)", _
              vbYesNo + vbExclamation + vbDefaultButton2, title) <> vbYes Then
        modLog.LogUsage "channel_purge_cancel", "knowledge", _
                        "channel=" & chName & " shown=" & shown
        GoTo Done
    End If

    ' R33H F2(c): 再描画の予約は削除の【前】に立てる。旧実装は削除の後ろに
    ' あったため、途中で例外が飛ぶと Failed: へ抜けて再描画すら走らず、
    ' 消したはずの資料が画面に残り続けた(「消えていない」ようにしか見えない)。
    redraw = True
    Dim purgeOk As Boolean: purgeOk = True
    removed = modChannel.PurgeChannelChunks(chName, purgeOk)
    modStats.SetStatText modShareRule.ChannelStatKey(chName), ""

    ' R33H F2(d): 書き戻しに失敗したまま「削除しました」と言わない。
    ' 購読解除の確認も出さない(失敗したのに「今後届きません」まで言うと、
    ' 資料は残ったまま供給だけ止まる=いちばん悪い状態になる)。
    If Not purgeOk Then
        modLog.LogUsage "channel_purge_ui", "knowledge", _
                        "channel=" & chName & " shown=" & shown & _
                        " removed=" & removed & " purge_ok=no"
        MsgBox "「" & chName & "」の資料の削除に失敗した可能性があります。" & vbCrLf & vbCrLf & _
               "マイ本棚を開いて、資料が減っているかご確認ください。" & vbCrLf & _
               "残っているときは、もう一度この操作をやり直してください。", _
               vbExclamation, title
        GoTo Done
    End If

    ' R33 W5-28: 削除だけでは購読は続くので、次に「部門チャンネル」を押すと
    ' 同じ部門がまた入ってくる。「今回だけ片付けたい」と「もう要らない」を
    ' 使い分けられるよう、成功した直後にここで1回だけ聞く(既定はいいえ)。
    ' 「はい」でだけ modChannel.Unsubscribe を呼ぶ ―― この関数は今まで
    ' src 全体で呼び出し元0件で、ここが初の実呼び出しになる。
    ' 【再購読の導線は無い】: SubscribeAllAvailable は IsSubscribed が偽の
    ' 部門を飛ばし、対になる modChannel.Subscribe も呼び出し元0件のため、
    ' 📡部門チャンネルを押しても戻らない。戻す手段は設定の変更だけなので、
    ' そのことを聞く前に必ず書く(「元に戻せます」と嘘をつかない)。
    Dim stopSub As Boolean
    stopSub = (MsgBox("「" & chName & "」の資料を " & removed & "件 削除しました。" & vbCrLf & _
              "本棚の使用量: " & modChannel.ChunkUsagePercent() & "%" & vbCrLf & vbCrLf & _
              "今後この部門の資料を受け取らないようにしますか?" & vbCrLf & vbCrLf & _
              "「いいえ」… 購読は続きます。次に「部門チャンネル」を押したときに、" & vbCrLf & _
              "     この部門の最新版がまた入ります(片付けだけしたいときはこちら)。" & vbCrLf & _
              "「はい」… 以後この部門の資料は届かなくなります。元に戻すには、" & vbCrLf & _
              "     このツールの管理担当に設定の変更をご依頼ください" & vbCrLf & _
              "     (この画面からは戻せません)。", _
              vbYesNo + vbQuestion + vbDefaultButton2, title) = vbYes)

    If stopSub Then
        modChannel.Unsubscribe chName
        modSkin.ShowToast "今後「" & chName & "」の資料は届きません。", "info"
    End If

    modLog.LogUsage "channel_purge_ui", "knowledge", _
                    "channel=" & chName & " shown=" & shown & " removed=" & removed & _
                    " unsubscribed=" & IIf(stopSub, "yes", "no")

Done:
    On Error GoTo 0
    If errNum <> 0 Then
        On Error Resume Next
        modLog.LogError "E0801", "modShared.OnPurgeChannel", errDesc, errNum
        modSkin.ShowToast "部門の資料の削除に失敗しました。もう一度お試しください。", "error"
        On Error GoTo 0
    End If
    modUiLock.Leave
    ' 再描画はロックを離してから1回だけ(RenderShelf も Enter を取るため)。
    If redraw Then
        On Error Resume Next
        modKnowledge.RefreshCurrent
        On Error GoTo 0
    End If
    Exit Sub

Failed:
    errNum = Err.Number
    errDesc = Err.Description
    Resume Done
End Sub
