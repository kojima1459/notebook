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

Private Const HEADER_ROW As Long = 12
Private Const FIRST_CARD_ROW As Long = 13
Private Const MAX_CARD_ROWS As Long = 400

Private Const COL_ICON As Long = 1     ' A
Private Const COL_NAME As Long = 2     ' B(B:D 結合)
Private Const COL_DATE As Long = 5     ' E
Private Const COL_CHUNKS As Long = 6   ' F(F:G 結合)
Private Const COL_MEMO As Long = 8     ' H(H:J 結合)

Private Const RNG_FOLDER As String = "A7:J7"
Private Const RNG_SYNCINFO As String = "D8:J9"

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
    ' R18-3a: 全域(ws.Cells)への書式はUsedRangeをシート最大へ膨らませる
    ' (無限スクロールの主因・調査agent2 §1.3)。3モードが共有するこのシートの
    ' 実使用範囲(A:N / カード最終行412)だけに当てる。
    ws.Range(modKnowledge.SHELF_BOUND).Font.Name = "游ゴシック"
    ws.Range(modKnowledge.SHELF_BOUND).Font.Size = 11

    ' A:N を全列ぶん明示する(2026-07-30 R4要件A)。
    ' 旧実装はK列・L列だけ設定しておらず、DrawChromeが使う W=A1:N1 の幅が
    ' 機種と操作履歴に依存してぶれていた。ツールバーの折り返し判定はこのWを
    ' 基準にしているので、Wがぶれると「端末によってだけボタンが見切れる」
    ' という再現しにくい不具合になる。3モードとも同じシートを共有する
    ' ようになった今は、前のモードの列幅が残る経路も増えている。
    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B:D").ColumnWidth = 10
    ws.Columns("E").ColumnWidth = 8
    ws.Columns("F:G").ColumnWidth = 8
    ws.Columns("H:J").ColumnWidth = 15
    ws.Columns("K:N").ColumnWidth = 10

    ' 実機防衛(2026-07-21再訂正): modUIMain.EnsureLayoutと同じ理由・同じ実装
    ' (Activateあり/なし双方で同一の1004が再現したため、Activate成否を致命的
    ' 前提にしない。失敗を許容して描画へ進む。詳細はmodUIMain.bas側参照)。
    uiStep = "描画前アクティブ化"
    Application.ScreenUpdating = True
    DoEvents
    On Error Resume Next
    ws.Activate   ' lint:allow-raw-activate(許容続行・R11-C裁定)
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

    ' 表示の共通儀式(左端へ戻す/等倍/旧Vaultシートの掃除)。R4要件B。
    ' 一度右へスクロールした状態が持ち越されると、この画面には戻す手段が
    ' 無い(水平スクロールバーはNexus起動時に消してある)。実機報告の
    ' 「🗑削除が『除』しか見えない」の主因はこれ。
    modKnowledge.PrepareScreenView ws

    ' ---- 共通クロム(ヘッダー+モード切替+ツールバー) ----------------------
    ' 2026-07-26 再設計: ナレッジ倉庫(カード)とマイ本棚(一覧)で
    ' まったく同じ上部UIを出し、ピルで切り替える1画面2モードに見せる。
    ' 横一列に7個並べていた旧ボタン行(A1:N2)は画面幅からはみ出していた。
    uiStep = "共通クロム(modKnowledge)"
    modKnowledge.DrawChrome ws, "table"

    ' ---- 本棚フォルダ情報 ----------------------------------------------------
    uiStep = "本棚フォルダ情報見出し"
    With ws.Range(RNG_FOLDER)
        .Merge
        .Value = "本棚フォルダ: "
        .Font.Size = 10
    End With
    ws.Rows("7").RowHeight = 16

    ' 「フォルダを選ぶ」ボタンは共通ツールバーの📂フォルダに集約したので
    ' ここには置かない(同じ機能のボタンが2箇所にある状態を解消)。
    uiStep = "同期情報"
    ws.Rows("8:9").RowHeight = 18
    With ws.Range(RNG_SYNCINFO)
        .Merge
        .Value = "自動同期: "
        .Font.Size = 9
        .VerticalAlignment = -4160   ' xlTop
    End With

    uiStep = "案内文(自動追加)"
    With ws.Range("A10:J10")
        .Merge
        .Value = "ここにファイルを入れておくと、自動で本棚に追加されます(消せば本棚からも消えます)"
        .Font.Size = 9
        .Font.Italic = True
    End With
    ws.Rows("10").RowHeight = 16

    uiStep = "資料カード見出し帯"
    With ws.Range("A11:J11")
        .Merge
        ' 2026-07-28(レビュー H-16): 🗑 は CP932 に無く、VBE 注入時に "??" へ
    ' 化けていた("??削除" と表示される)。サロゲートペアを ChrW で組む。
    .Value = "── 資料カード(1行=1資料。行をクリックしてから" & _
             ChrW(&HD83D) & ChrW(&HDDD1) & "削除) ─────────"
        .Font.Bold = True
        .Font.Size = 10
    End With
    ws.Rows("11").RowHeight = 16

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
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup0
FailCleanup0:
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
    ' 2026-07-30(R4要件A): ギャラリー/一覧表/解決事例が同じ「マイ本棚」
    ' シートを共有するようになった。取込・同期の完了時に走る自動再描画
    ' (modShelf/modShelfSync → ここ)は、利用者がギャラリーを見ている
    ' 最中でも呼ばれるので、そのまま書くとカードの裏に一覧表が現れる。
    ' 一覧表モード以外では空振りさせ、表示中のモードの描き直しは
    ' modKnowledge.RefreshCurrent に任せる。
    Dim isTable As Boolean
    isTable = True
    On Error Resume Next
    isTable = modKnowledge.IsTableMode()
    On Error GoTo 0
    If Not isTable Then
        ' 2026-07-30(レビュー1-B): 空振り(no-op)だと、ギャラリー/解決事例を
        ' 見ている最中に取り込んだ資料が【画面のどこにも出ない】。利用者から
        ' 見れば「押したのに何も起きない」で、取り込めたかどうかも分からない。
        ' 表示中のモードを描き直す方へ委譲する。
        ' 再帰しない: RefreshCurrent の table 分岐はここを呼ぶが、その分岐へ
        ' 入るのは IsTableMode()=True のときだけで、そのときこの委譲は起きない。
        On Error Resume Next
        modKnowledge.RefreshCurrent
        On Error GoTo 0
        Exit Sub
    End If

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
            .Value = "まだ資料がありません。上のツールバーの「" & ChrW(&HD83D) & ChrW(&HDCC1) & " 追加」から始めましょう。"
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
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup1
FailCleanup1:
    On Error Resume Next
    Application.ScreenUpdating = True
    On Error GoTo 0
    Err.Raise origNum, "modUIShelf.RenderShelf", "[" & uiStep & "] " & origDesc
End Sub

' ----------------------------------------------------------------------------
' OnIngestScreenshot - クリップボードの画像(スクリーンショット)を本棚へ
'   取り込む(裁定D13)。流れ: 画像有無確認→タイトル入力→jpg保存(Temp)→
'   「<タイトル>_日時.jpg」で本棚フォルダ(未設定ならTempのまま)へコピー→
'   既存のmodShelf.IngestFileへ合流(画像はvision委譲フォールバックが処理)。
'   optVisionへの参照はR2に従いmodFeatures.InvokeFeature経由のみ。
' ----------------------------------------------------------------------------
Public Sub OnIngestScreenshot()
    If modUiLock.BlockIfIngesting() Then Exit Sub
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
    ' 2026-07-30(レビュー1-B): 取込結果を今見ている画面へ必ず反映する。
    ' スクショ取込はギャラリー/解決事例のツールバーからも押せるが、
    ' IngestFile 側の再描画は一覧表モードのときしか効かなかった。
    On Error Resume Next
    modKnowledge.RefreshCurrent
    On Error GoTo 0
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
End Sub

Private Sub RefreshFolderInfo(ByVal ws As Worksheet)
    Dim folder As String
    folder = Trim$(modConfig.GetString("shelf_folder", ""))

    Dim folderText As String
    If LenB(folder) = 0 Then
        folderText = "本棚フォルダ: (未設定。上のツールバーの「" & ChrW(&HD83D) & ChrW(&HDCC2) & " フォルダ」から選んでください)"
    Else
        folderText = "本棚フォルダ: " & folder
    End If
    WriteSafe ws.Range(RNG_FOLDER), folderText

    Dim minutes As Long
    minutes = modConfig.GetLong("sync_interval_min", 0)

    Dim syncText As String
    If minutes < 1 Then
        syncText = "自動同期: オフ(上のツールバーの「" & ChrW(&HD83D) & ChrW(&HDD04) & " 同期」を押すと今すぐ同期します)"
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
            StatusIcon = ChrW(&H2705)                          ' U+2705 チェック
        Case "pending", "partial"
            StatusIcon = ChrW(&H23F3)                          ' U+23F3 砂時計
        Case "failed", "failed_permanent"
            StatusIcon = ChrW(&H26A0) & ChrW(&HFE0F)            ' U+26A0+FE0F 警告
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
            ' R14-4c: メモがあるときは、それが「何が起きたか」の唯一の正確な
            ' 説明(OCRの上限打ち切り・薄い抽出)。どちらも続きから再開する
            ' 仕組みは無く、同期を押しても何も起きない(RC4の嘘の案内)。
            ' 本当に再開で進むのはベクトル化待ち(メモ無し)だけなので、
            ' 案内はそちらにだけ残す。
            If LenB(errorNote) > 0 Then
                BuildMemo = errorNote
            Else
                BuildMemo = "一部だけ変換が完了していません。上のツールバーの「" & ChrW(&HD83D) & ChrW(&HDD04) & " 同期」を押すと続きから再開します。"
            End If
        Case "image_pdf"
            ' R14-F8: メモがあるなら、それが「何が起きたか」の唯一の正確な説明
            ' (GS未検出・起動失敗・時間切れ・共有読みの失敗など)。固定文の
            ' 「Ghostscriptを置くと読み取れます」は、既にGhostscriptがある端末で
            ' 別の理由で失敗した資料にも出ていて、次の一手を誤らせていた。
            If LenB(errorNote) > 0 Then
                BuildMemo = errorNote
            Else
                BuildMemo = "画像として保存されたPDFです。Ghostscriptを置くとAIが1ページずつ読み取れます" & _
                    "(手順は「43_画像PDFのOCR取込設定」)。急ぐときは画面をコピーしてスクショ取込へ。"
            End If
        Case "missing"
            BuildMemo = "元のファイルが見つかりません。フォルダを確認するか、次回の同期をお待ちください。"
        Case "failed"
            If LenB(errorNote) > 0 Then
                BuildMemo = "取込に失敗しました: " & errorNote
            Else
                BuildMemo = "取込に失敗しました。"
            End If
        Case "failed_permanent"
            ' R12-3-3: 3回続けて失敗したので自動同期の対象から外した状態。
            ' 「なぜ同期しても直らないのか」と「次の一手」を必ず書く。
            BuildMemo = "3回続けて取込に失敗したため、自動同期の対象から外しました。" & _
                "上のツールバーの「資料を追加」でこのファイルを選び直すと、もう一度試します。"
            If LenB(errorNote) > 0 Then BuildMemo = BuildMemo & "(前回: " & errorNote & ")"
        Case Else
            BuildMemo = errorNote
    End Select
End Function

' stats(i) = "status|ingested_at|chunk_count|error_note|origin"(modShelf.SourceList契約)。
' error_note自体に "|" が含まれても壊れないよう、先頭3要素+末尾のorigin+中間を
' 再結合したerror_noteという頑健な形でパースする。
' 2026-08-04(R15-8b・実機第4報 RC1): modVaultGallery.DrawOneCardが単純な
' Split(statLine,"|")で同じ契約をパースする重複実装を持っていた(error_note
' に"|"が混じると値がズレる同じ危険を2箇所に埋める形)。共用できるようPublic化
' する(呼び出し元が増えても中身はこのまま=挙動不変)。
Public Sub ParseStats(ByVal s As String, ByRef status As String, ByRef ingestedAt As String, _
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
