Attribute VB_Name = "modTextView"
Option Explicit

' ============================================================================
' modTextView - スクショ本文表示(読み取り結果をその場で確かめる・R36波1 §3)
' ----------------------------------------------------------------------------
' 役割:
'   1資料ぶんの全チャンクのfull_textを、ページ順に1枚の作業シート(text_view)
'   へ流し込んで見せる。modDiag.RunDiagnostics/RecreateDiagSheet と同じ
'   「毎回Delete→Add」方式(modDiag.bas:299-334)。編集→保存し直しは無い
'   (裁定: 内容が違うときは消して➕登録し直す導線を案内するだけ)。
' 入口:
'   ・modKnowledge.OnShowText(📖本文ボタン。一覧表モード限定) → ShowActiveRow
'   ・modUIShelf.OnIngestScreenshot(📸取込直後)→ AfterShot
'     (config screenshot_show_text=TRUE のときだけ自動で開く)
' 設計判断:
'   ・戻り先(マイ本棚/チャット)はmodule変数mBackActionへ保持する。ブック
'     再起動で消えても既定"shelf"へ倒す(§3-2契約どおり)。
'   ・SHELF_FIRST_CARD_ROW/SHELF_COL_NAMEはmodUIShelfのPrivate定数
'     (FIRST_CARD_ROW=13/COL_NAME=2。modUIShelf.bas:46,50)と同値を自前で
'     持つ(modUIShelfは残96字で関数を足せないため。片方を変えたら両方)。
'   ・my_knowledgeの列レイアウト(1=ID/2=SOURCE/3=ORIGIN/4=PAGE/5=SUMMARY/
'     6=KEYWORDS/7=FULLTEXT/8=ADDED)は凍結modShelf.bas:10-18の定数と同値。
' ============================================================================

Private Const TEXT_VIEW_SHEET As String = "text_view"
' modUIShelf.FIRST_CARD_ROW / COL_NAME と同値(modUIShelf.bas:46,50)。
' 片方を変えたら必ず両方変えること。
Private Const SHELF_FIRST_CARD_ROW As Long = 13
Private Const SHELF_COL_NAME As Long = 2

Private mBackAction As String
' R43 2-7: NoteShelfTruncated(モジュール末尾)が使う「案内済み」印。
Private mShelfTruncNoted As Boolean

' ----------------------------------------------------------------------------
' ShowForSource - 資料名 srcName の全チャンクをページ順に表示する。
'   backAction: "shelf"=戻る先マイ本棚 / "chat"=戻る先チャット。
' ----------------------------------------------------------------------------
Public Sub ShowForSource(ByVal srcName As String, ByVal backAction As String)
    mBackAction = backAction
    On Error GoTo Fail

    Dim wsK As Worksheet
    Set wsK = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)

    Dim lastK As Long
    lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row

    Dim hitRows() As Long
    Dim pages() As Long, texts() As String
    Dim n As Long: n = 0
    Dim addedAt As String: addedAt = ""
    Dim i As Long

    If lastK >= 2 Then
        ' 【2段で読む理由・R36 Fix A-M1】my_knowledge は実機で2万行あり、
        ' full_text は1行が最大32,000字。8列を丸ごと配列にすると1回の表示で
        ' 数百MBを読み、32bit Excel ではメモリ不足で【無言で死ぬ】。
        ' そこで modCorrect.InjectHits と同じ2段読みにする:
        '   (1) id/source/origin/page の4列だけを一括で読んで該当行を集める
        '   (2) ReDim は該当件数ぶん (3) 該当行だけ Cells(row,7) で full_text
        ' 4列読むのは、1列だけの Range.Value が1行のとき配列ではなくスカラーを
        ' 返して LBound で落ちるため(modCorrect.bas:305-308 と同じ理由)。
        Dim arr As Variant
        arr = wsK.Range(wsK.Cells(2, 1), wsK.Cells(lastK, 4)).Value
        Dim target As String: target = LCase$(Trim$(srcName))
        Dim cap As Long: cap = UBound(arr, 1) - LBound(arr, 1)
        ReDim hitRows(0 To cap)
        ReDim pages(0 To cap)
        For i = LBound(arr, 1) To UBound(arr, 1)
            If LCase$(Trim$(CStr(arr(i, 2)))) = target Then
                hitRows(n) = i + 1          ' arr の i 行目 = シートの i+1 行目(1行目は見出し)
                ' R36 Fix N9: page 列が空/非数値でも CLng の型不一致で
                ' 資料まるごと落とさない(Val は数字以外を 0 と読む)。
                pages(n) = CLng(Val(CStr(arr(i, 4))))
                n = n + 1
            End If
        Next i

        ' (3) 該当行だけ full_text(7列目)を読む。ReDim は【無条件に1回】通す
        '     (CLAUDE.md §10 の R33波3: 条件付き ReDim Preserve の穴を作らない)。
        ReDim texts(0 To n)
        For i = 0 To n - 1
            texts(i) = CStr(wsK.Cells(hitRows(i), 7).Value)
        Next i
        If n > 0 Then addedAt = CStr(wsK.Cells(hitRows(0), 8).Value)
        If n > 1 Then SortPagesStable pages, texts, n
    End If

    Dim ws As Worksheet
    Set ws = RecreateTextViewSheet()
    If ws Is Nothing Then Exit Sub

    ' R36 Fix B3: 資料の生テキストは「=」「+」「-」「@」で始まりうる。素の
    ' Value 代入だと Excel が数式として解釈し 1004 で落ち、戻るボタンの無い
    ' 書きかけシートだけが残る(modState.bas:79-95 の既知の型)。A列を文字列
    ' 書式に固定してから書く。
    ws.Columns("A").NumberFormat = "@"

    ' R38 M1(外部レビュー裏どり済): Worksheets.Add は新シートを【表示中】にする
    ' ので、以降の1セルずつの書き込み(大きい資料で数百回)が毎回再描画される。
    ' 書き終えるまで描画を止め、正常・失敗のどちらの出口でも必ず戻す
    ' (modUIShelf:109/130 と同じ Finally 型)。
    Application.ScreenUpdating = False

    Dim r As Long: r = 1
    PutCell ws, r, ChrW(&HD83D) & ChrW(&HDCD6) & " 本文: " & srcName
    r = r + 1
    PutCell ws, r, _
        "取込日時: " & IIf(LenB(addedAt) > 0, addedAt, "不明") & _
        " ／ " & n & "個のまとまりに分けて保存 ／ 内容が違うときは、この資料を " & _
        ChrW(&HD83D) & ChrW(&HDDD1) & " で消してから " & ChrW(&H2795) & _
        " 登録で正しい文章を登録し直してください" & _
        "(読み取った文章を直接直す機能はありません)。"
    ' R37 §3(資料間リンク): 章の重心が近い他資料があれば1行だけ添える。
    ' 1件も無ければ行そのものを出さない(=空の見出しを見せない)。判定と
    ' 並べ替えは modXDocStore.LinkedLabel に閉じるので、ここは受け取るだけ。
    Dim xdLink As String: xdLink = modXDocStore.LinkedLabel(srcName)
    If LenB(xdLink) > 0 Then
        r = r + 1
        PutCell ws, r, "関連する資料: " & xdLink
    End If
    r = r + 2

    If n = 0 Then
        PutCell ws, r, "(この資料のチャンクが見つかりませんでした)"
        r = r + 1
    Else
        For i = 0 To n - 1
            ' R41 §1 A: Excel由来は「シートN」(modMode.PageTagPart。単一情報源)。
            PutCell ws, r, "【" & Trim$(modMode.PageTagPart(srcName, pages(i))) & "】"
            r = r + 1
            Dim body As String: body = StripBreadcrumb(texts(i))
            Dim pos As Long: pos = 1
            Do While pos <= Len(body)
                PutCell ws, r, Mid$(body, pos, 500)
                r = r + 1
                pos = pos + 500
            Loop
            r = r + 1
        Next i
    End If

    ws.Columns("A").ColumnWidth = 110
    DrawBackButton ws
    Application.ScreenUpdating = True

    ' R36 Fix M4: ws.Activate の素呼びは R10-1 実機 err91 の経路(モーダル/
    ' 外部COM直後は例外を返すのにシートは切り替わっている)。成否判定と記録は
    ' modUI.ActivateSheetRobust へ一本化する。失敗しても内容は書き終えている
    ' ので続行し、脱出路だけ出す(modUI.GoToNexus:512-515 と同じ組み合わせ)。
    If Not modUI.ActivateSheetRobust(ws, "modTextView.ShowForSource") Then
        modUI.RestoreExcelUI
    End If
    Exit Sub

Fail:
    ' On Error 文と Exit は Err をリセットするので、ログの前に退避する(§11)。
    Dim eN As Long: eN = Err.Number
    Dim eD As String: eD = Err.Description
    Application.ScreenUpdating = True   ' 描画停止のまま抜けない(R38 M1)
    ' R36 Fix A-m1: 途中で落ちた text_view は「戻るボタンの無い書きかけ
    ' シート」として残る。掃除してから記録する(残骸を作らない)。
    CleanupSheet "modTextView.ShowForSource.Fail"
    modLog.LogError "E0801", "modTextView.ShowForSource", eD, eN
    Err.Clear
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' NoteShelfTruncated - 本棚一覧が MAX_CARD_ROWS(400件)で打ち切られたときの
'   案内(R43 2-7)。modUIShelf.RenderShelfが残79字で実体を持てないため
'   こちらへ置く(呼び出しは1行)。セッション中1回だけ知らせる。
' ----------------------------------------------------------------------------
Public Sub NoteShelfTruncated(ByVal hiddenN As Long)
    If mShelfTruncNoted Then Exit Sub
    If hiddenN < 1 Then Exit Sub
    ' レビュー R43 2周目 m8: 印は「実際に描ける状況」でだけ消費する。
    ' ShowToast は別ブックを見ている間は1枚も描かずに戻るので、先に
    ' 印を立てると一度も見ないまま権利を失う。
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    mShelfTruncNoted = True
    On Error Resume Next
    ' レビュー R43 M3: waitless=True。呼び元(modUIShelf.RenderShelf)は
    ' ScreenUpdating=False・AlertsOff の描画途中なので、待つ型のトースト
    ' (既定 False)だと 3 秒ぶん DoEvents で固まり、しかも画面が更新されない
    ' まま消えて一度も見えない。予約消去(modToast)に任せる。
    modSkin.ShowToast "ほか" & hiddenN & "件は表示していません(不要な資料を削除すると表示されます)。", "info", True
    On Error GoTo 0
End Sub

' 1セル単位の書き込み(R36 Fix B3)。1セルの失敗で画面全体を失わないよう、
' 保護はセル単位に閉じる(NumberFormat="@" を掛けてもなお、極端に長い文字列や
' 保護の掛かった環境で失敗しうる)。On Error はプロシージャ単位なので、
' 呼び出し元の On Error GoTo Fail はこの Sub の中の指定に影響されない。
Private Sub PutCell(ByVal ws As Worksheet, ByVal r As Long, ByVal s As String)
    On Error Resume Next
    ws.Cells(r, 1).Value = s
    Err.Clear
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' OnBack - 「← 戻る」ボタン。text_viewを削除し、開いた元の画面へ戻る。
' ----------------------------------------------------------------------------
' R36 Fix A-m2: 連打で二重遷移しないよう modKnowledge.OnDelete(:770-779)と
'   同型の Enter/Leave を掛ける。ただし遷移先の modApp.OnNavShelf は自分でも
'   modUiLock.Enter する(modApp.bas:551)ので、【掴んだまま呼ぶと必ず断られ
'   「押しても何も起きない」】。掃除までを錠の中で行い、Leave してから遷移する。
Public Sub OnBack()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    ' R36 Fix2 C-m1: 失敗ログはCleanupSheet自身の中(Err.Clearの前)へ移した。
    ' ここで改めてErr.Numberを見る判定は、CleanupSheetが戻る時点で必ず
    ' Err.Clear済み=常に0という「到達不能な死んだ判定」だったため削除。
    CleanupSheet "modTextView.OnBack"
    modUiLock.Leave

    Select Case mBackAction
        Case "chat"
            modUI.GoToNexus "modTextView.OnBack"
        Case Else
            modApp.OnNavShelf   ' 既定"shelf"(ブック再起動でmBackActionが消えた場合も含む)
    End Select
End Sub

' ----------------------------------------------------------------------------
' CleanupSheet - text_view を消す(R36 Fix A-r4)。
'   text_view は【可視】の作業シートなので、開いたまま保存すると全文が
'   ブックに焼き付き、誰でも見られる状態で OneDrive へ同期される。
'   呼ばれるのは (1)OnBack (2)ShowForSource の作り直し/失敗時
'   (3)modApp.OnSaveAndExit の直前(保存物に残さない)の3箇所。
'   凍結 modBoot.Auto_Close には触らない(§12)。
'   DisplayAlerts の退避と復元は On Error Resume Next の中で必ず対にする。
'   R36 Fix2 C-m1: 削除に失敗したら Err.Clear の【前】に記録する
'   (以前は OnBack 側で CleanupSheet 呼び出し後に Err.Number を見ていたが、
'   CleanupSheet はここで Err.Clear してから戻るため常に 0 = 到達不能な
'   死んだ判定だった。呼び出し元(OnBack/ShowForSource の Fail:/
'   modApp.OnSaveAndExit)には caller 名を渡させ、ログにどこからの掃除かを
'   残す)。「まだ無い」(該当シート無し=Err9)は失敗ではないので記録しない。
' ----------------------------------------------------------------------------
Public Sub CleanupSheet(Optional ByVal caller As String = "")
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(TEXT_VIEW_SHEET).Delete
    If Err.Number <> 0 And Err.Number <> 9 Then
        modLog.LogError "E0801", "modTextView.CleanupSheet(" & caller & ")", Err.Description, Err.Number
    End If
    Application.DisplayAlerts = True
    Err.Clear
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' ShowActiveRow - 📖本文ボタン(マイ本棚・一覧表モード限定)。
'   modUIShelf.OnDeleteSourceの前半(行→資料名の判定)と同じ4段の判定を
'   自前で持つ(modUIShelfは残96字で関数を足せないための複製。
'   modUIShelf.bas:489-522のOnDeleteSourceと同型)。
' ----------------------------------------------------------------------------
' R36 Fix N7: 例外の記録は【この関数の中】で持つ。modKnowledge は残324字の
'   逼迫モジュールで、OnShowText は Enter/Leave と1行呼び出しだけに縮めた
'   (§12「実体は余裕モジュールへ、逼迫側は1行呼び出し」)。
Public Sub ShowActiveRow()
    On Error GoTo Fail
    Dim activeName As String
    activeName = ""
    On Error Resume Next
    activeName = Application.ActiveSheet.Name
    On Error GoTo Fail        ' 0 に戻すと下の Fail 網が外れる(§11)

    If activeName <> modAppDef.SH_SHELF Then
        MsgBox "本文を見たい資料の行をクリックしてから、もう一度押してください。", _
               vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    Dim r As Long
    r = Application.ActiveCell.Row
    If r < SHELF_FIRST_CARD_ROW Then
        MsgBox "本文を見たい資料の行をクリックしてから、もう一度押してください。", _
               vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_SHELF)
    On Error GoTo Fail
    If ws Is Nothing Then Exit Sub

    Dim sourceName As String
    sourceName = Trim$(CStr(ws.Cells(r, SHELF_COL_NAME).Value))
    If LenB(sourceName) = 0 Then
        MsgBox "この行には資料がありません。", vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    ShowForSource sourceName, "shelf"
    Exit Sub

Fail:
    ' ログの前に Err を退避する(On Error 文と Exit は Err をリセットする・§11)。
    Dim eN2 As Long: eN2 = Err.Number
    Dim eD2 As String: eD2 = Err.Description
    modLog.LogError "E0801", "modTextView.ShowActiveRow", eD2, eN2
    Err.Clear
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' AfterShot - 📸スクショ取込の直後に呼ばれる(modUIShelf.OnIngestScreenshot
'   末尾。旧: modKnowledge.RefreshCurrent の呼び出しをこれに置換)。
'   本棚一覧の再描画は従来どおり必ず行い、config screenshot_show_text が
'   真のときだけ続けて読み取り結果を自動表示する(髙橋フィードバック対応)。
' ----------------------------------------------------------------------------
' R36 Fix M5 / A-M3: 取込が失敗したときまで本文画面を開いていた。中身は
'   空か、直前に見ていた別資料の残骸で、利用者は「読めている」と誤解する。
'   ingestStatus は modShelf.IngestFile の戻り値(modShelf.bas:363-365,
'   87/99/146/184/304/412 で "done"/"partial"/"failed"/"image_pdf")。
'   表示するのは done / partial のときだけで、それ以外は RefreshCurrent のみ
'   ―― 失敗の理由は IngestFile 側が既にトースト/ログで伝えている。
'   判定を modTextView 側に置くのは、modUIShelf が残93字で分岐を書けないため
'   (あちらの変更は引数 ", ingestStatus" の13字だけ)。
Public Sub AfterShot(ByVal destPath As String, ByVal ingestStatus As String)
    On Error Resume Next
    modKnowledge.RefreshCurrent
    On Error GoTo 0

    Dim st As String: st = LCase$(Trim$(ingestStatus))
    If st <> "done" Then
        If st <> "partial" Then Exit Sub
    End If

    If modConfig.GetBool("screenshot_show_text", True) Then
        On Error Resume Next
        ShowForSource modUtil.FileNameOf(destPath), "shelf"
        On Error GoTo 0
    End If
End Sub

' ----------------------------------------------------------------------------
' StripBreadcrumb - 取込時に焼かれた先頭行【資料名>章>条】を剥がす(純関数)。
'   modVaultGallery.MakePreviewText(:724-731)と同じ剥がし方: 先頭が「【」で
'   かつ改行があれば最初のvbLfまでを落とす。改行が無ければ変更しない
'   (「【だけ」のような不完全な行を誤って全部消さないため=既存踏襲)。
' ----------------------------------------------------------------------------
Public Function StripBreadcrumb(ByVal s As String) As String
    Dim txt As String: txt = s
    If Left$(txt, 1) = "【" Then
        Dim lfPos As Long: lfPos = InStr(txt, vbLf)
        If lfPos > 0 Then txt = Mid$(txt, lfPos + 1)
    End If
    StripBreadcrumb = txt
End Function

' ----------------------------------------------------------------------------
' SortPagesStable - pages/texts(並列配列・0始まりでもn始まりでも可)の
'   先頭n件を、pages昇順の安定ソート(挿入ソート)で並べ替える(純関数)。
'   UDTを跨がないための素朴な配列2本渡し(Hit型はここに来ない)。
' ----------------------------------------------------------------------------
Public Sub SortPagesStable(ByRef pages() As Long, ByRef texts() As String, ByVal n As Long)
    If n <= 1 Then Exit Sub
    Dim lo As Long: lo = LBound(pages)
    Dim i As Long, j As Long
    Dim keyP As Long, keyT As String
    For i = lo + 1 To lo + n - 1
        keyP = pages(i)
        keyT = texts(i)
        j = i - 1
        Do While j >= lo
            If pages(j) <= keyP Then Exit Do   ' <=なので等しい要素は追い越さない=安定
            pages(j + 1) = pages(j)
            texts(j + 1) = texts(j)
            j = j - 1
        Loop
        pages(j + 1) = keyP
        texts(j + 1) = keyT
    Next i
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' modDiag.RecreateDiagSheet(:299-334)と同じ「Delete→Add、改名失敗時は
' 既存を再利用してClearContents」方式。戻るボタンは無い診断シートと違い、
' 本シートは「← 戻る」を自前で描く(DrawBackButton)。
Private Function RecreateTextViewSheet() As Worksheet
    CleanupSheet "modTextView.ShowForSource"   ' 古い残骸を消す(DisplayAlertsの退避/復元も向こう側)

    Dim ws As Worksheet
    On Error GoTo Fallback
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    ws.Name = TEXT_VIEW_SHEET
    On Error GoTo 0
    Set RecreateTextViewSheet = ws
    Exit Function

Fallback:
    Resume FallbackCleanup
FallbackCleanup:
    On Error Resume Next
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    Set ws = ThisWorkbook.Worksheets(TEXT_VIEW_SHEET)
    If Not ws Is Nothing Then ws.Cells.ClearContents
    Err.Clear
    On Error GoTo 0
    Set RecreateTextViewSheet = ws
End Function

' 左上の「← 戻る」フローティングボタン(modHelp.DrawManualBackButtonと同型)。
Private Sub DrawBackButton(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Shapes("tv_back").Delete
    Err.Clear                 ' 「まだ無い」の91は失敗ではない(下の判定へ持ち越さない)

    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(5, 10, 8, 110, 26)   ' 5=角丸四角
    ' R36 Fix A-m8: AddShape が失敗すると以降が全部素通りし、【戻れない画面】
    ' が黙って出来上がる(シートタブでしか脱出できない)。1行でも残す。
    If btn Is Nothing Then
        Dim eN As Long: eN = Err.Number
        modLog.LogError "E0801", "modTextView(DrawBackButton)", _
            "戻るボタンのAddShapeに失敗(シートタブから戻ってください)", eN
        Err.Clear
        Exit Sub
    End If
    btn.Name = "tv_back"
    btn.Adjustments(1) = 0.3
    btn.Fill.ForeColor.RGB = modUI.UiColor("primary")
    btn.Line.Visible = 0
    With btn.TextFrame2
        .WordWrap = -1
        .TextRange.Text = ChrW(&H2190) & " 戻る"
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = -1
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
    End With
    btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    btn.OnAction = "modTextView.OnBack"
    btn.Placement = 3
    btn.ZOrder 0
    On Error GoTo 0
End Sub
