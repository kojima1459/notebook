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

    ' uiStep: modUIMain.EnsureLayoutと同じ考え方(2026-07-15 実機E0801対策)。
    Dim uiStep As String
    On Error GoTo Fail

    Application.ScreenUpdating = False
    If ws.Visible <> -1 Then ws.Visible = -1   ' xlSheetVisible(非表示なら表示化)

    Dim prevActive As Object
    On Error Resume Next
    Set prevActive = ThisWorkbook.ActiveSheet
    On Error GoTo Fail

    uiStep = "既存ボタンの削除"
    RemoveManagedShapes ws
    DoEvents   ' 削除と追加の間でCOM/メモリを一拍解放する(先回り防衛#4)
    uiStep = "セルのクリア"
    ws.Cells.Clear

    uiStep = "既定フォント設定"
    ws.Cells.Font.Name = "游ゴシック"
    ws.Cells.Font.Size = 11

    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B:D").ColumnWidth = 10
    ws.Columns("E").ColumnWidth = 8
    ws.Columns("F:G").ColumnWidth = 8
    ws.Columns("H:J").ColumnWidth = 15
    ws.Columns("M:N").ColumnWidth = 10

    ' 実機防衛(2026-07-21再訂正): modUIMain.EnsureLayoutと同じ理由・同じ実装
    ' (Activateあり/なし双方で同一の1004が再現したため、Activate成否を致命的
    ' 前提にしない。失敗を許容して描画へ進む。詳細はmodUIMain.bas側参照)。
    uiStep = "描画前アクティブ化"
    Application.ScreenUpdating = True
    DoEvents
    On Error Resume Next
    ws.Activate
    If Err.Number <> 0 Then
        Dim actNum As Long: actNum = Err.Number
        modLog.LogError "E0801", "modUIShelf.EnsureLayout", _
            "[描画前アクティブ化(許容続行)] ws.Visible=" & ws.Visible & _
            " ActiveSheet=" & ThisWorkbook.ActiveSheet.Name & _
            " AppWin=" & Application.Windows.Count & " WbWin=" & ThisWorkbook.Windows.Count, actNum
        Err.Clear
    End If
    On Error GoTo Fail
    Application.ScreenUpdating = False

    ' ---- 操作ボタン行 ----------------------------------------------------
    uiStep = "操作ボタン行"
    ws.Rows("1:2").RowHeight = 22
    AddButton ws, ws.Range("A1:B2"), "btn_add", "＋ 資料を追加", "modUIShelf.OnAddFiles"
    AddButton ws, ws.Range("C1:D2"), "btn_sync", "" & ChrW(&HD83D) & ChrW(&HDD04) & " フォルダと同期", "modUIShelf.OnSyncNow"
    AddButton ws, ws.Range("E1:F2"), "btn_pack_out", "" & ChrW(&HD83D) & ChrW(&HDCE6) & " パックにして渡す", "modUIShelf.OnExportPack"
    AddButton ws, ws.Range("G1:H2"), "btn_pack_in", "" & ChrW(&HD83D) & ChrW(&HDCE5) & " パックを取り込む", "modUIShelf.OnImportPack"
    AddButton ws, ws.Range("I1:J2"), "btn_delete", "" & ChrW(&HD83D) & ChrW(&HDDD1) & " 選んだ資料を削除", "modUIShelf.OnDeleteSource"
    ' スクショ取込(裁定D13)は画像解析機能が有効なときだけボタンを出す
    ' (無効環境で「押したら断られるボタン」を見せないため)。
    uiStep = "スクショ取込ボタン(有効判定)"
    If modFeatures.FeatureEnabled("vision") Then
        uiStep = "スクショ取込ボタン(生成)"
        AddButton ws, ws.Range("K1:L2"), "btn_screenshot", "" & ChrW(&HD83D) & ChrW(&HDCF8) & " スクショ取込", "modUIShelf.OnIngestScreenshot"
    End If

    ' 2026-07-22実機報告対策: Nexus(チャット)へタブなしで戻れる導線
    uiStep = "チャットへ戻るボタン"
    AddButton ws, ws.Range("M1:N2"), "btn_back_chat", "" & ChrW(&HD83D) & ChrW(&HDCAC) & " チャットへ", "modUIShelf.OnBackToChat"

    ' ---- 本棚フォルダ情報 ----------------------------------------------------
    uiStep = "本棚フォルダ情報見出し"
    With ws.Range(RNG_FOLDER)
        .Merge
        .Value = "本棚フォルダ: "
        .Font.Size = 10
    End With
    ws.Rows("3").RowHeight = 16

    uiStep = "フォルダを選ぶボタン"
    ws.Rows("4:5").RowHeight = 18
    AddButton ws, ws.Range("A4:C5"), "btn_pick_folder", "" & ChrW(&HD83D) & ChrW(&HDCC1) & " フォルダを選ぶ", "modUIShelf.OnPickFolder"
    With ws.Range(RNG_SYNCINFO)
        .Merge
        .Value = "自動同期: "
        .Font.Size = 9
        .VerticalAlignment = -4160   ' xlTop
    End With

    uiStep = "案内文(自動追加)"
    With ws.Range("A6:J6")
        .Merge
        .Value = "ここにファイルを入れておくと、自動で本棚に追加されます(消せば本棚からも消えます)"
        .Font.Size = 9
        .Font.Italic = True
    End With
    ws.Rows("6").RowHeight = 16

    uiStep = "資料カード見出し帯"
    With ws.Range("A7:J7")
        .Merge
        .Value = "── 資料カード(1行=1資料) ───────────────────"
        .Font.Bold = True
        .Font.Size = 10
    End With
    ws.Rows("7").RowHeight = 16

    ' ---- カード見出し行 ----------------------------------------------------
    uiStep = "カード見出し行"
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

    On Error Resume Next
    If Not prevActive Is Nothing Then prevActive.Activate   ' 元のアクティブシートへ復帰
    On Error GoTo 0
    Application.ScreenUpdating = True

    uiStep = "資料一覧の再描画(RenderShelf)"
    RenderShelf
    Exit Sub

Fail:
    Dim origNum As Long, origDesc As String
    origNum = Err.Number
    origDesc = Err.Description
    ' 2026-07-21: Err.Raiseで包んだDescriptionが呼び出し元まで生き残らない事例
    ' が実機で確認されたため、伝播に依存せずここで直接err_logへ書く。
    Dim diag As String: diag = ""
    On Error Resume Next
    diag = " ws.Visible=" & ws.Visible & " ActiveSheet=" & ThisWorkbook.ActiveSheet.Name
    modLog.LogError "E0801", "modUIShelf.EnsureLayout", "[" & uiStep & "]" & diag, origNum
    If Not prevActive Is Nothing Then prevActive.Activate
    Application.ScreenUpdating = True
    On Error GoTo 0
    Err.Raise origNum, "modUIShelf.EnsureLayout", "[" & uiStep & "] " & origDesc & diag
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

    ' uiStep: modUIMain.EnsureLayoutと同じ考え方(2026-07-16 実機E0801
    ' 「資料一覧の再描画(RenderShelf)」報告への対策。EnsureLayout側の
    ' 粒度では原因ブロックまで特定できなかったため、ここでも追う)。
    Dim uiStep As String
    On Error GoTo Fail

    Application.ScreenUpdating = False

    uiStep = "本棚フォルダ情報の更新"
    RefreshFolderInfo ws
    uiStep = "カード領域のクリア"
    ClearCardArea ws

    uiStep = "資料一覧の取得(SourceList)"
    Dim names() As String
    Dim stats() As String
    Dim n As Long
    n = modShelf.SourceList(names, stats)

    If n = 0 Then
        uiStep = "空の本棚の案内表示"
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

    uiStep = "資料カードの描画(1件目)"
    Dim i As Long
    For i = 0 To shown - 1
        uiStep = "資料カードの描画(" & (i + 1) & "件目/" & shown & "件)"
        RenderOneCard ws, FIRST_CARD_ROW + i, names(i), stats(i)
    Next i

    Application.ScreenUpdating = True
    Exit Sub

Fail:
    Dim origNum As Long, origDesc As String
    origNum = Err.Number
    origDesc = Err.Description
    On Error Resume Next
    Application.ScreenUpdating = True
    On Error GoTo 0
    Err.Raise origNum, "modUIShelf.RenderShelf", "[" & uiStep & "] " & origDesc
End Sub

' ----------------------------------------------------------------------------
' OnAddFiles / OnSyncNow / OnPickFolder / OnExportPack / OnImportPack / OnDeleteSource
' ----------------------------------------------------------------------------

' 2026-07-22実機報告対策: タブが隠れていてもNexus(チャット)へ戻れるように
' する(modUI.GoToNexusと同じ脱出路付き遷移をここから呼ぶだけ)。
Public Sub OnBackToChat()
    modUI.GoToNexus "modUIShelf.OnBackToChat"
End Sub

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

' ----------------------------------------------------------------------------
' OnIngestScreenshot - クリップボードの画像(スクリーンショット)を本棚へ
'   取り込む(裁定D13)。流れ: 画像有無確認→タイトル入力→jpg保存(Temp)→
'   「<タイトル>_日時.jpg」で本棚フォルダ(未設定ならTempのまま)へコピー→
'   既存のmodShelf.IngestFileへ合流(画像はvision委譲フォールバックが処理)。
'   optVisionへの参照はR2に従いmodFeatures.InvokeFeature経由のみ。
' ----------------------------------------------------------------------------
Public Sub OnIngestScreenshot()
    On Error GoTo Fail

    ' 1) クリップボードに画像があるか(機能無効/mock時は#ERR文字列が返る)
    Dim hasRes As Variant
    hasRes = modFeatures.InvokeFeature("vision", "HasClipboardImage", Array())
    If VarType(hasRes) = vbString Then
        MsgBox "この機能は現在利用できません(管理者が有効化すると使えます)。", _
               vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If
    If Not CBool(hasRes) Then
        MsgBox "クリップボードに画像がありません。" & vbCrLf & _
               "取り込みたい画面をコピー(PrintScreen や Win+Shift+S)してから、もう一度押してください。", _
               vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    ' 2) 資料タイトルを尋ねる(空欄・キャンセルは中止)
    Dim docTitle As String
    docTitle = Trim$(InputBox("この画像の資料名を入力してください。" & vbCrLf & _
                              "(例: 経費精算マニュアル 12ページ)", "スクショ取込"))
    If LenB(docTitle) = 0 Then Exit Sub

    ' 3) クリップボード画像をjpg保存(Tempのパスが返る)
    Dim savedRes As Variant
    savedRes = modFeatures.InvokeFeature("vision", "SaveClipboardImage", Array())
    Dim savedPath As String
    savedPath = CStr(savedRes)
    If LenB(savedPath) = 0 Or Left$(savedPath, 5) = "#ERR:" Then
        MsgBox "画像の保存に失敗しました。もう一度画面をコピーしてからお試しください。", _
               vbExclamation, modAppDef.APP_NAME
        Exit Sub
    End If

    ' 4) 「<タイトル>_yyyymmddhhnnss.jpg」へ改名コピー。保存先は本棚フォルダ
    '    (shelf_folder。未設定ならTempと同じ場所)。同名衝突は日時秒で実質回避。
    Dim destDir As String
    destDir = modConfig.GetString("shelf_folder", "")
    If LenB(destDir) = 0 Then destDir = Left$(savedPath, InStrRev(savedPath, "\") - 1)
    If Right$(destDir, 1) = "\" Then destDir = Left$(destDir, Len(destDir) - 1)
    Dim destPath As String
    destPath = destDir & "\" & SanitizeFileName(docTitle) & "_" & Format$(Now, "yyyymmddhhnnss") & ".jpg"
    FileCopy savedPath, destPath

    ' 5) 既存の取込パイプラインへ(進捗実況・本棚一覧更新はIngestFile側の責務)
    Dim ingestStatus As String
    ingestStatus = modShelf.IngestFile(destPath, "self")

    RefreshBadgesAndDashboard
    Exit Sub
Fail:
    modLog.LogError "E0801", "modUIShelf.OnIngestScreenshot", Err.Description
    Err.Clear
    On Error GoTo 0
End Sub

' ファイル名に使えない文字を "_" に置換し、長すぎるタイトルは80字で切る。
Private Function SanitizeFileName(ByVal s As String) As String
    Dim bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)
    Dim t As String
    t = s
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        t = Replace(t, CStr(bad(i)), "_")
    Next i
    SanitizeFileName = modUtil.SafeLeft(Trim$(t), 80)
End Function

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
        folderText = "本棚フォルダ: (未設定。「" & ChrW(&HD83D) & ChrW(&HDCC1) & " フォルダを選ぶ」から選んでください)"
    Else
        folderText = "本棚フォルダ: " & folder
    End If
    WriteSafe ws.Range(RNG_FOLDER), folderText

    Dim minutes As Long
    minutes = modConfig.GetLong("sync_interval_min", 0)

    Dim syncText As String
    If minutes < 1 Then
        syncText = "自動同期: オフ(「" & ChrW(&HD83D) & ChrW(&HDD04) & " フォルダと同期」を押すと今すぐ同期します)"
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
    ' 実機報告(2026-07-22)「状態列が全部？になる」対策: ここだけVBEソースに
    ' 絵文字を直書きしたままで、他所で確定・修正済みの文字化けバグ(自己
    ' インストーラのAddFromString往復でリテラル絵文字が壊れる)を踏んでいた。
    ' ChrWのコードポイント指定に統一する。
    Select Case LCase$(status)
        Case "done"
            StatusIcon = ChrW(&H2705)                          ' ✅
        Case "pending", "partial"
            StatusIcon = ChrW(&H23F3)                          ' ⏳
        Case "failed"
            StatusIcon = ChrW(&H26A0) & ChrW(&HFE0F)            ' ⚠️
        Case "image_pdf"
            StatusIcon = ChrW(&HD83D) & ChrW(&HDDBC)
        Case "missing"
            StatusIcon = ChrW(&HD83D) & ChrW(&HDD52)
        Case Else
            StatusIcon = ChrW(&H30FB)                          ' ・
    End Select
End Function

Private Function BuildMemo(ByVal status As String, ByVal chunkCount As String, ByVal errorNote As String) As String
    Select Case LCase$(status)
        Case "done"
            BuildMemo = chunkCount & "件のまとまりに分けて保存されました。"
        Case "pending"
            BuildMemo = "AIが読める形に変換中です。しばらくしてから確認してください。"
        Case "partial"
            BuildMemo = "一部だけ変換が完了していません。「" & ChrW(&HD83D) & ChrW(&HDD04) & "フォルダと同期」を押すと続きから再開します。"
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
    ' 【2026-07-16 恒久修正】セルの日付型自動変換により、stampは
    ' "2026-07-16 15:02:33"(書込み時の文字列)と"2026/07/16 15:02:33"
    ' (日付型セルをCStrしたロケール表記)の両方があり得る。IsDateなら
    ' 型に依存せずCDateで月/日に整形する(従来は"-"区切り前提で、
    ' 実機では整形されない生の日時が資料カードに出ていた)。
    If IsDate(stamp) Then
        Dim d As Date: d = CDate(stamp)
        ShortDate = Month(d) & "/" & Day(d)
        Exit Function
    End If

    Dim dateOnly As String
    dateOnly = Left$(stamp, 10)

    Dim parts() As String
    parts = Split(dateOnly, "-")
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
    ' 未Mergeのまま.Valueへ代入すると全セルに同じ値が複製される(実機報告の
    ' バグ)。EnsureLayoutが途中で中断しMerge未実行のまま呼ばれても壊れない
    ' よう、呼び出し側の前提に頼らずここで都度Mergeする(冪等・低コスト)。
    If cell.Cells.Count > 1 Then cell.Merge
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

' マイクロログ(2026-07-21): modUIMain.AddButtonと同じ理由・同じ実装
' (失敗した「正確な1行」をerr_logへ残す。詳細はmodUIMain.bas側参照)。
Private Sub AddButton(ByVal ws As Worksheet, ByVal rng As Range, ByVal shapeName As String, _
                      ByVal caption As String, ByVal action As String)
    Dim uiStep As String
    On Error GoTo Fail
    Dim shp As Shape
    ' ログと実呼出しに同じサニタイズ済み値を使う(生値だとログと実態が食い違う)。
    Dim sL As Double, sT As Double, sW As Double, sH As Double
    sL = SafeCoord(rng.Left): sT = SafeCoord(rng.Top)
    sW = SafeCoord(rng.Width): sH = SafeCoord(rng.Height)
    uiStep = "AddShape実行(Type=5 L=" & sL & " T=" & sT & " W=" & sW & " H=" & sH & ")"
    Set shp = SafeRoundedRect(ws, sL, sT, sW, sH)
    uiStep = "図形名設定"
    shp.Name = shapeName
    uiStep = "テキスト代入"
    shp.TextFrame2.TextRange.Text = caption
    uiStep = "フォント設定"
    shp.TextFrame2.WordWrap = -1   ' msoTrue
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Bold = -1   ' msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = 2   ' msoAlignCenter
    shp.TextFrame2.VerticalAnchor = 3   ' msoAnchorMiddle
    uiStep = "色/塗りつぶし設定"
    shp.Fill.ForeColor.RGB = 15921906   ' RGB(242,242,242)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
    shp.Line.Visible = 0   ' msoFalse
    uiStep = "OnAction割当て"
    shp.OnAction = action
    Exit Sub
Fail:
    Dim btnErrNum As Long, btnErrDesc As String
    btnErrNum = Err.Number: btnErrDesc = Err.Description
    Dim diag As String: diag = ""
    On Error Resume Next
    diag = " ws.Visible=" & ws.Visible & " ws.ProtectContents=" & ws.ProtectContents & _
           " ActiveSheet=" & ThisWorkbook.ActiveSheet.Name & _
           " Interactive=" & Application.Interactive & _
           " AppWin=" & Application.Windows.Count & " WbWin=" & ThisWorkbook.Windows.Count
    modLog.LogError "E0801", "modUIShelf.EnsureLayout", "AddButton [" & uiStep & "]" & diag, btnErrNum
    On Error GoTo 0
    ' 2026-07-22実機再発: modUIMain.AddButtonと同じ理由でErr.Raise再伝播をやめる
    ' (1個のボタン失敗でEnsureLayout全体が中断→画面がほぼ空になるのを防ぐ)。
    Err.Clear
End Sub

' 2026-07-21訂正で撤去: 旧SafeBeginDraw/SafeEndDraw(ws.Activate・DisplayObjects
' 強制・Protect解除一式)。真因はコンパイルエラーであり、これらは的外れな
' 対策だった。座標サニタイズ(下記SafeCoord)のみ実効性があるため維持する。
Private Function SafeCoord(ByVal v As Double) As Double
    If v < 1 Then v = 1
    SafeCoord = v
End Function

' 呼び出し元がSafeCoordで既に安全化した値を渡す前提だが、直接呼ばれても
' 壊れないよう二重にクランプする(コストはほぼ無い)。
Private Function SafeRoundedRect(ByVal ws As Worksheet, ByVal L As Double, ByVal T As Double, _
                                 ByVal W As Double, ByVal H As Double) As Shape
    L = SafeCoord(L): T = SafeCoord(T): W = SafeCoord(W): H = SafeCoord(H)
    On Error GoTo Retry
    Set SafeRoundedRect = ws.Shapes.AddShape(5, L, T, W, H)   ' 5=msoShapeRoundedRectangle(リテラル)
    Exit Function
Retry:
    DoEvents
    Set SafeRoundedRect = ws.Shapes.AddShape(5, L, T, W, H)
End Function

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
