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

    Dim pages() As Long, texts() As String
    Dim n As Long: n = 0
    Dim addedAt As String: addedAt = ""

    If lastK >= 2 Then
        ReDim pages(0 To lastK - 2)
        ReDim texts(0 To lastK - 2)
        ' source(2列目)〜added(8列目)を一括読み(modVaultGallery.EnsurePreviewIndex
        ' と同じ配列一括読みの作法。modShelf.bas:10-18のCOL_*と同値の列位置)。
        Dim arr As Variant
        arr = wsK.Range(wsK.Cells(2, 1), wsK.Cells(lastK, 8)).Value
        Dim target As String: target = LCase$(Trim$(srcName))
        Dim i As Long
        For i = LBound(arr, 1) To UBound(arr, 1)
            If LCase$(Trim$(CStr(arr(i, 2)))) = target Then
                pages(n) = CLng(arr(i, 4))
                texts(n) = CStr(arr(i, 7))
                If LenB(addedAt) = 0 Then addedAt = CStr(arr(i, 8))
                n = n + 1
            End If
        Next i
        If n > 0 Then
            ReDim Preserve pages(0 To n - 1)
            ReDim Preserve texts(0 To n - 1)
            SortPagesStable pages, texts, n
        End If
    End If

    Dim ws As Worksheet
    Set ws = RecreateTextViewSheet()

    Dim r As Long: r = 1
    ws.Cells(r, 1).Value = ChrW(&HD83D) & ChrW(&HDCD6) & " 本文: " & srcName
    r = r + 1
    ws.Cells(r, 1).Value = _
        "取込日時: " & IIf(LenB(addedAt) > 0, addedAt, "不明") & _
        " ／ チャンク数: " & n & "件 ／ 内容が違うときは、この資料を " & _
        ChrW(&HD83D) & ChrW(&HDDD1) & " で消してから " & ChrW(&H2795) & _
        " 登録で正しい文章を登録し直してください" & _
        "(読み取った文章を直接直す機能はありません)。"
    r = r + 2

    If n = 0 Then
        ws.Cells(r, 1).Value = "(この資料のチャンクが見つかりませんでした)"
        r = r + 1
    Else
        For i = 0 To n - 1
            ws.Cells(r, 1).Value = "【p." & pages(i) & "】"
            r = r + 1
            Dim body As String: body = StripBreadcrumb(texts(i))
            Dim pos As Long: pos = 1
            Do While pos <= Len(body)
                ws.Cells(r, 1).Value = Mid$(body, pos, 500)
                r = r + 1
                pos = pos + 500
            Loop
            r = r + 1
        Next i
    End If

    ws.Columns("A").ColumnWidth = 110
    DrawBackButton ws

    ' R11-C(lint許容登録)と同じ許容続行: 前面化に失敗しても内容は書き
    ' 終えているため続行する(modDiag.RunDiagnostics:232-240と同型)。
    On Error Resume Next
    Err.Clear
    ws.Activate   ' lint:allow-raw-activate(許容続行・modDiagと同じ裁定)
    If Err.Number <> 0 Then
        modLog.LogError "E0801", "modTextView.ShowForSource", _
            "[許容続行] Activate失敗", Err.Number
    End If
    On Error GoTo 0
    Exit Sub

Fail:
    modLog.LogError "E0801", "modTextView.ShowForSource", Err.Description
    Err.Clear
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' OnBack - 「← 戻る」ボタン。text_viewを削除し、開いた元の画面へ戻る。
' ----------------------------------------------------------------------------
Public Sub OnBack()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(TEXT_VIEW_SHEET).Delete
    Application.DisplayAlerts = True
    Err.Clear
    On Error GoTo 0

    Select Case mBackAction
        Case "chat"
            modUI.GoToNexus "modTextView.OnBack"
        Case Else
            modApp.OnNavShelf   ' 既定"shelf"(ブック再起動でmBackActionが消えた場合も含む)
    End Select
End Sub

' ----------------------------------------------------------------------------
' ShowActiveRow - 📖本文ボタン(マイ本棚・一覧表モード限定)。
'   modUIShelf.OnDeleteSourceの前半(行→資料名の判定)と同じ4段の判定を
'   自前で持つ(modUIShelfは残96字で関数を足せないための複製。
'   modUIShelf.bas:489-522のOnDeleteSourceと同型)。
' ----------------------------------------------------------------------------
Public Sub ShowActiveRow()
    Dim activeName As String
    activeName = ""
    On Error Resume Next
    activeName = Application.ActiveSheet.Name
    On Error GoTo 0

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
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim sourceName As String
    sourceName = Trim$(CStr(ws.Cells(r, SHELF_COL_NAME).Value))
    If LenB(sourceName) = 0 Then
        MsgBox "この行には資料がありません。", vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    ShowForSource sourceName, "shelf"
End Sub

' ----------------------------------------------------------------------------
' AfterShot - 📸スクショ取込の直後に呼ばれる(modUIShelf.OnIngestScreenshot
'   末尾。旧: modKnowledge.RefreshCurrent の呼び出しをこれに置換)。
'   本棚一覧の再描画は従来どおり必ず行い、config screenshot_show_text が
'   真のときだけ続けて読み取り結果を自動表示する(髙橋フィードバック対応)。
' ----------------------------------------------------------------------------
Public Sub AfterShot(ByVal destPath As String)
    On Error Resume Next
    modKnowledge.RefreshCurrent
    On Error GoTo 0

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
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(TEXT_VIEW_SHEET).Delete
    Err.Clear
    On Error GoTo 0
    Application.DisplayAlerts = True

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

    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(5, 10, 8, 110, 26)   ' 5=角丸四角
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
