Attribute VB_Name = "modUIShelf"
Option Explicit

' ============================================================================
' modUIShelf - 「マイ本棚」画面(本棚UI)の構築と描画(MASTER_SPEC §7.6/§8.2)
' ----------------------------------------------------------------------------
' 役割:
'   資料の追加・同期・パック共有・削除をこの画面の5つのボタンから行う。
'   RenderShelfはmodShelf.SourceListから資料カード(1行=1資料)を再描画する。
'   modShelf.IngestFile/DeleteSource、modShelfSync.SyncNowは処理完了後に
'   自身でmodUIShelf.RenderShelfを呼ぶ設計(§7.2実装コメント)なので、
'   このモジュールのボタンハンドラは基本的に「呼ぶだけ」でよい。
'
' 設計判断:
'   ・資料カードは1行=1資料の「表(セル値+書式)」として描画する。V2
'     modChatUI.bas同様、ラベル類はセル値、実際にクリックされる操作だけを
'     Shape(btn_プレフィクス)にする。
'   ・カード領域は固定の最大行数(MAX_CARD_ROWS)をあらかじめ確保し、
'     RenderShelfの度に領域全体をクリア(ClearContents+アンマージ)してから
'     書き直す(件数が減っても古い行が残らないようにするため)。
'   ・SourceList が返す stats(i) は "status|ingested_at|chunk_count|
'     error_note|origin" のパイプ区切り(modShelf.bas §7.2実装コメント参照)。
'     error_note自体に "|" が含まれる可能性を考慮し、単純Splitの限界数指定
'     ではなく「先頭3要素+末尾1要素(origin)+その間を再結合(error_note)」
'     という頑健なパース(ParseStats)にした。
'   ・状態アイコンはMASTER_SPEC §8.2の凡例 ✅⏳⚠️🖼🕒 に対応させる
'     (done=✅ / pending・partial=⏳ / failed=⚠️ / image_pdf=🖼 / missing=🕒)。
'   ・削除(OnDeleteSource)は「今アクティブなセルがある行」の資料名を対象と
'     する(§8.2「削除はアクティブセルの行の資料を対象」)。マイ本棚シートが
'     アクティブでない、またはカード領域外の行が選択されている場合は
'     「削除したい資料の行をクリックしてから、もう一度押してください」と
'     案内するだけで、Excel組み込みのMsgBoxで確認してから実削除する
'     (ファイル名を明示: 「『x』を本棚から削除しますか?」)。
'   ・本棚フォルダのパス表示・自動同期間隔表示はRenderShelfが毎回
'     configから読み直して更新する(PickShelfFolder実行直後にも
'     RenderShelfがmodShelfSync経由で呼ばれるため、常に最新表示になる)。
'   ・opt機能(約款差分)はfeature_diffdocが常時TRUE(§5)だが「確認済み
'     関数のみ使用」の注記のとおりコアと同じ確実な呼び出しのみ許可されて
'     いるため、このマイ本棚画面には差分ボタンを置かない(仕様に明記の
'     無いUI追加はしない)。opt機能ボタンが必要になった場合はmodFeatures.
'     FeatureEnabled判定つきで追加できるよう、EnsureLayoutの末尾に
'     余白行(TTS等と同様の型)を確保できる設計にしてある。
' ============================================================================

Private Const HEADER_ROW As Long = 8
Private Const FIRST_CARD_ROW As Long = 9
Private Const MAX_CARD_ROWS As Long = 400

Private Const COL_ICON As Long = 1     ' A
Private Const COL_NAME As Long = 2     ' B(B:D 結合)
Private Const COL_DATE As Long = 5     ' E
Private Const COL_CHUNKS As Long = 6   ' F(F:G 結合)
Private Const COL_MEMO As Long = 8     ' H(H:J 結合)

Private Const RNG_FOLDER As String = "A3:J3"
Private Const RNG_SYNCINFO As String = "D4:J5"

' ----------------------------------------------------------------------------
' EnsureLayout
' ----------------------------------------------------------------------------
Public Sub EnsureLayout()
    Dim ws As Worksheet
    Set ws = GetOrCreateShelfSheet()
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    RemoveManagedShapes ws
    ws.Cells.Clear

    ws.Cells.Font.Name = "游ゴシック"
    ws.Cells.Font.Size = 11

    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B:D").ColumnWidth = 10
    ws.Columns("E").ColumnWidth = 8
    ws.Columns("F:G").ColumnWidth = 8
    ws.Columns("H:J").ColumnWidth = 15

    ' ---- 操作ボタン行 ----------------------------------------------------
    ws.Rows("1:2").RowHeight = 22
    AddButton ws, ws.Range("A1:B2"), "btn_add", "＋ 資料を追加", "modUIShelf.OnAddFiles"
    AddButton ws, ws.Range("C1:D2"), "btn_sync", "🔄 フォルダと同期", "modUIShelf.OnSyncNow"
    AddButton ws, ws.Range("E1:F2"), "btn_pack_out", "📦 パックにして渡す", "modUIShelf.OnExportPack"
    AddButton ws, ws.Range("G1:H2"), "btn_pack_in", "📥 パックを取り込む", "modUIShelf.OnImportPack"
    AddButton ws, ws.Range("I1:J2"), "btn_delete", "🗑 選んだ資料を削除", "modUIShelf.OnDeleteSource"

    ' ---- 本棚フォルダ情報 ----------------------------------------------------
    With ws.Range(RNG_FOLDER)
        .Merge
        .Value = "本棚フォルダ: "
        .Font.Size = 10
    End With
    ws.Rows("3").RowHeight = 16

    ws.Rows("4:5").RowHeight = 18
    AddButton ws, ws.Range("A4:C5"), "btn_pick_folder", "📁 フォルダを選ぶ", "modUIShelf.OnPickFolder"
    With ws.Range(RNG_SYNCINFO)
        .Merge
        .Value = "自動同期: "
        .Font.Size = 9
        .VerticalAlignment = -4160   ' xlTop
    End With

    With ws.Range("A6:J6")
        .Merge
        .Value = "ここにファイルを入れておくと、自動で本棚に追加されます(消せば本棚からも消えます)"
        .Font.Size = 9
        .Font.Italic = True
    End With
    ws.Rows("6").RowHeight = 16

    With ws.Range("A7:J7")
        .Merge
        .Value = "── 資料カード(1行=1資料) ───────────────────"
        .Font.Bold = True
        .Font.Size = 10
    End With
    ws.Rows("7").RowHeight = 16

    ' ---- カード見出し行 ----------------------------------------------------
    ws.Cells(HEADER_ROW, COL_ICON).Value = "状態"
    ws.Range(ws.Cells(HEADER_ROW, COL_NAME), ws.Cells(HEADER_ROW, COL_NAME + 2)).Merge
    ws.Cells(HEADER_ROW, COL_NAME).Value = "資料名"
    ws.Cells(HEADER_ROW, COL_DATE).Value = "追加日"
    ws.Range(ws.Cells(HEADER_ROW, COL_CHUNKS), ws.Cells(HEADER_ROW, COL_CHUNKS + 1)).Merge
    ws.Cells(HEADER_ROW, COL_CHUNKS).Value = "チャンク数"
    ws.Range(ws.Cells(HEADER_ROW, COL_MEMO), ws.Cells(HEADER_ROW, COL_MEMO + 2)).Merge
    ws.Cells(HEADER_ROW, COL_MEMO).Value = "メモ"
    With ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(HEADER_ROW, 10))
        .Font.Bold = True
        .Font.Size = 9
        .Interior.Color = 15921906   ' RGB(242,242,242)
    End With
    ws.Rows(HEADER_ROW).RowHeight = 16

    Application.ScreenUpdating = True

    RenderShelf
End Sub

' ----------------------------------------------------------------------------
' RenderShelf - manifest+modShelf.SourceListから1行=1資料で再描画(§8.2)
' ----------------------------------------------------------------------------
Public Sub RenderShelf()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetShelfSheet()
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    RefreshFolderInfo ws
    ClearCardArea ws

    Dim names() As String
    Dim stats() As String
    Dim n As Long
    n = modShelf.SourceList(names, stats)

    If n = 0 Then
        With ws.Range(ws.Cells(FIRST_CARD_ROW, COL_NAME), ws.Cells(FIRST_CARD_ROW, 10))
            .Merge
            .Value = "まだ資料がありません。上の「＋資料を追加」から始めましょう。"
            .Font.Italic = True
            .Font.Size = 10
        End With
        ws.Rows(FIRST_CARD_ROW).RowHeight = 18
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Dim shown As Long
    shown = n
    If shown > MAX_CARD_ROWS Then shown = MAX_CARD_ROWS

    Dim i As Long
    For i = 0 To shown - 1
        RenderOneCard ws, FIRST_CARD_ROW + i, names(i), stats(i)
    Next i

    Application.ScreenUpdating = True
End Sub

' ----------------------------------------------------------------------------
' OnAddFiles / OnSyncNow / OnPickFolder / OnExportPack / OnImportPack / OnDeleteSource
' ----------------------------------------------------------------------------
Public Sub OnAddFiles()
    On Error GoTo Fail
    modShelf.AddFilesViaDialog
    RefreshBadgesAndDashboard
    Exit Sub
Fail:
    modLog.LogError "E0801", "modUIShelf.OnAddFiles", Err.Description
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub OnSyncNow()
    On Error GoTo Fail
    modShelfSync.SyncNow
    RefreshBadgesAndDashboard
    Exit Sub
Fail:
    modLog.LogError "E0801", "modUIShelf.OnSyncNow", Err.Description
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub OnPickFolder()
    On Error GoTo Fail
    modShelfSync.PickShelfFolder
    RefreshBadgesAndDashboard
    Exit Sub
Fail:
    modLog.LogError "E0801", "modUIShelf.OnPickFolder", Err.Description
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub OnExportPack()
    On Error GoTo Fail
    modPack.ExportPackDialog
    RefreshBadgesAndDashboard
    Exit Sub
Fail:
    modLog.LogError "E0801", "modUIShelf.OnExportPack", Err.Description
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub OnImportPack()
    On Error GoTo Fail
    modPack.ImportPackDialog
    RefreshBadgesAndDashboard
    Exit Sub
Fail:
    modLog.LogError "E0801", "modUIShelf.OnImportPack", Err.Description
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub OnDeleteSource()
    On Error GoTo Fail

    Dim activeName As String
    activeName = ""
    On Error Resume Next
    activeName = Application.ActiveSheet.Name
    On Error GoTo 0

    If activeName <> modAppDef.SH_SHELF Then
        MsgBox "削除したい資料の行をクリックしてから、もう一度押してください。", _
               vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    Dim r As Long
    r = Application.ActiveCell.Row
    If r < FIRST_CARD_ROW Then
        MsgBox "削除したい資料の行をクリックしてから、もう一度押してください。", _
               vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = GetShelfSheet()
    If ws Is Nothing Then Exit Sub

    Dim sourceName As String
    sourceName = Trim$(CStr(ws.Cells(r, COL_NAME).Value))
    If LenB(sourceName) = 0 Then
        MsgBox "この行には資料がありません。削除したい資料の行をクリックしてから、もう一度押してください。", _
               vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    Dim answer As Long
    answer = MsgBox("『" & sourceName & "』を本棚から削除しますか?" & vbLf & _
                     "この操作は取り消せません。", vbYesNo + vbQuestion, modAppDef.APP_NAME)
    If answer <> vbYes Then Exit Sub

    modShelf.DeleteSource sourceName
    RefreshBadgesAndDashboard
    Exit Sub

Fail:
    modLog.LogError "E0801", "modUIShelf.OnDeleteSource", Err.Description
    Err.Clear
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

Private Sub RefreshBadgesAndDashboard()
    On Error Resume Next
    modStats.EvaluateBadges
    On Error GoTo 0

    On Error Resume Next
    modUIDashboard.RenderDashboard
    On Error GoTo 0
End Sub

Private Sub RefreshFolderInfo(ByVal ws As Worksheet)
    Dim folder As String
    folder = Trim$(modConfig.GetString("shelf_folder", ""))

    Dim folderText As String
    If LenB(folder) = 0 Then
        folderText = "本棚フォルダ: (未設定。「📁 フォルダを選ぶ」から選んでください)"
    Else
        folderText = "本棚フォルダ: " & folder
    End If
    WriteSafe ws.Range(RNG_FOLDER), folderText

    Dim minutes As Long
    minutes = modConfig.GetLong("sync_interval_min", 0)

    Dim syncText As String
    If minutes < 1 Then
        syncText = "自動同期: オフ(「🔄 フォルダと同期」を押すと今すぐ同期します)"
    Else
        syncText = "自動同期: " & minutes & "分ごと"
    End If
    WriteSafe ws.Range(RNG_SYNCINFO), syncText
End Sub

Private Sub ClearCardArea(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = FIRST_CARD_ROW + MAX_CARD_ROWS - 1

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(FIRST_CARD_ROW, 1), ws.Cells(lastRow, 10))

    On Error Resume Next
    rng.UnMerge
    On Error GoTo 0

    rng.ClearContents
    rng.WrapText = False
    ws.Rows(FIRST_CARD_ROW & ":" & lastRow).RowHeight = 15
End Sub

Private Sub RenderOneCard(ByVal ws As Worksheet, ByVal r As Long, ByVal sourceName As String, ByVal statsLine As String)
    Dim status As String
    Dim ingestedAt As String
    Dim chunkCount As String
    Dim errorNote As String
    Dim origin As String
    ParseStats statsLine, status, ingestedAt, chunkCount, errorNote, origin

    ws.Cells(r, COL_ICON).Value = StatusIcon(status)
    ws.Cells(r, COL_ICON).HorizontalAlignment = -4108   ' xlCenter

    With ws.Range(ws.Cells(r, COL_NAME), ws.Cells(r, COL_NAME + 2))
        .Merge
        .Value = sourceName
        .Font.Size = 9
    End With

    ws.Cells(r, COL_DATE).Value = ShortDate(ingestedAt)
    ws.Cells(r, COL_DATE).Font.Size = 9

    With ws.Range(ws.Cells(r, COL_CHUNKS), ws.Cells(r, COL_CHUNKS + 1))
        .Merge
        .Value = chunkCount & "件"
        .Font.Size = 9
        .HorizontalAlignment = -4152   ' xlRight
    End With

    Dim memo As String
    memo = BuildMemo(status, chunkCount, errorNote)
    With ws.Range(ws.Cells(r, COL_MEMO), ws.Cells(r, COL_MEMO + 2))
        .Merge
        .Value = memo
        .WrapText = True
        .VerticalAlignment = -4160   ' xlTop
        .Font.Size = 9
    End With

    ws.Rows(r).RowHeight = MemoRowHeight(memo)
End Sub

Private Function MemoRowHeight(ByVal memo As String) As Double
    Const CHARS_PER_LINE As Long = 34
    Dim lines As Long
    lines = (Len(memo) + CHARS_PER_LINE - 1) \ CHARS_PER_LINE
    If lines < 1 Then lines = 1

    Dim h As Double
    h = lines * 14 + 6
    If h < 26 Then h = 26
    If h > 90 Then h = 90
    MemoRowHeight = h
End Function

Private Function StatusIcon(ByVal status As String) As String
    Select Case LCase$(status)
        Case "done"
            StatusIcon = "✅"
        Case "pending", "partial"
            StatusIcon = "⏳"
        Case "failed"
            StatusIcon = "⚠️"
        Case "image_pdf"
            StatusIcon = "🖼"
        Case "missing"
            StatusIcon = "🕒"
        Case Else
            StatusIcon = "・"
    End Select
End Function

Private Function BuildMemo(ByVal status As String, ByVal chunkCount As String, ByVal errorNote As String) As String
    Select Case LCase$(status)
        Case "done"
            BuildMemo = chunkCount & "件のまとまりに分けて保存されました。"
        Case "pending"
            BuildMemo = "AIが読める形に変換中です。しばらくしてから確認してください。"
        Case "partial"
            BuildMemo = "一部だけ変換が完了していません。「🔄フォルダと同期」を押すと続きから再開します。"
        Case "image_pdf"
            BuildMemo = "画像として保存されたPDFのため、文字を取り込めませんでした。"
        Case "missing"
            BuildMemo = "元のファイルが見つかりません。フォルダを確認するか、次回の同期をお待ちください。"
        Case "failed"
            If LenB(errorNote) > 0 Then
                BuildMemo = "取込に失敗しました: " & errorNote
            Else
                BuildMemo = "取込に失敗しました。"
            End If
        Case Else
            BuildMemo = errorNote
    End Select
End Function

' stats(i) = "status|ingested_at|chunk_count|error_note|origin"(modShelf.SourceList契約)。
' error_note自体に "|" が含まれても壊れないよう、先頭3要素+末尾のorigin+中間を
' 再結合したerror_noteという頑健な形でパースする。
Private Sub ParseStats(ByVal s As String, ByRef status As String, ByRef ingestedAt As String, _
                       ByRef chunkCount As String, ByRef errorNote As String, ByRef origin As String)
    status = "": ingestedAt = "": chunkCount = "": errorNote = "": origin = ""

    Dim parts() As String
    parts = Split(s, "|")
    Dim n As Long
    n = UBound(parts) - LBound(parts) + 1

    If n >= 1 Then status = parts(0)
    If n >= 2 Then ingestedAt = parts(1)
    If n >= 3 Then chunkCount = parts(2)

    If n >= 5 Then
        origin = parts(n - 1)
        Dim midCount As Long
        midCount = n - 4
        Dim midParts() As String
        ReDim midParts(0 To midCount - 1)
        Dim i As Long
        For i = 0 To midCount - 1
            midParts(i) = parts(3 + i)
        Next i
        errorNote = Join(midParts, "|")
    ElseIf n = 4 Then
        errorNote = parts(3)
    End If
End Sub

Private Function ShortDate(ByVal stamp As String) As String
    Dim datePart As String
    datePart = Left$(stamp, 10)

    Dim parts() As String
    parts = Split(datePart, "-")
    If (UBound(parts) - LBound(parts) + 1) = 3 Then
        If IsNumeric(parts(1)) And IsNumeric(parts(2)) Then
            ShortDate = CStr(CLng(parts(1))) & "/" & CStr(CLng(parts(2)))
            Exit Function
        End If
    End If
    ShortDate = stamp
End Function

Private Sub WriteSafe(ByVal cell As Range, ByVal text As String)
    Dim t As String
    t = modUtil.SafeLeft(text, 32000)
    On Error Resume Next
    cell.Value = t
    On Error GoTo 0
End Sub

Private Function GetShelfSheet() As Worksheet
    On Error Resume Next
    Set GetShelfSheet = ThisWorkbook.Worksheets(modAppDef.SH_SHELF)
    On Error GoTo 0
End Function

Private Function GetOrCreateShelfSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_SHELF)
    On Error GoTo 0
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = modAppDef.SH_SHELF
        On Error GoTo 0
    End If
    Set GetOrCreateShelfSheet = ws
    Exit Function
Fail:
    Set GetOrCreateShelfSheet = Nothing
End Function

Private Sub AddButton(ByVal ws As Worksheet, ByVal rng As Range, ByVal shapeName As String, _
                      ByVal caption As String, ByVal action As String)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, rng.Left, rng.Top, rng.Width, rng.Height)   ' 5 = msoShapeRoundedRectangle
    shp.Name = shapeName
    shp.TextFrame2.TextRange.Text = caption
    shp.TextFrame2.WordWrap = -1   ' msoTrue
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Bold = -1   ' msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = 2   ' msoAlignCenter
    shp.TextFrame2.VerticalAnchor = 3   ' msoAnchorMiddle
    shp.Fill.ForeColor.RGB = 15921906   ' RGB(242,242,242)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
    shp.Line.Visible = 0   ' msoFalse
    shp.OnAction = action
End Sub

Private Sub RemoveManagedShapes(ByVal ws As Worksheet)
    Dim names() As String
    ReDim names(0 To ws.Shapes.Count)
    Dim n As Long
    n = 0

    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 4) = "btn_" Or Left$(shp.Name, 4) = "lbl_" Then
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
