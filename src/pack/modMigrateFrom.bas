Attribute VB_Name = "modMigrateFrom"
Option Explicit

' ============================================================================
' modMigrateFrom - 版上げ時の自動引き継ぎ(R36 spec §1-A・波4)
' ----------------------------------------------------------------------------
' 解決する問題:
'   利用者は新しい zip を旧版と同じフォルダへ展開する。README の従来手順
'   「旧版は展開前に削除」だと、削除した瞬間に本棚(旧xlsmの中)が消える。
'   R36裁定: 手順を「旧版の MyBookshelf.xlsm を MyBookshelf_前回.xlsm へ
'   改名して残す」へ変え、新版の初回起動でその旧xlsmを自動で見つけて
'   本棚・実績・設定を直接読み込む(modMigrate の .xlsx 引き継ぎファイル方式
'   とは別経路。あちらはユーザーが手で書き出し/読み込む「持ち運び」用、
'   こちらは「同じフォルダにある一つ前のバージョン」を自動で拾う版上げ専用)。
'
' 起点(架け元): modTour.StartTourIfFirstRun の先頭(凍結 modBoot には触らない。
'   ウィザードが走らない2回目以降の起動でも「本棚が空・未質問・前版あり」の
'   3条件が揃えば聞く必要があるため、tour完了フラグの早期Exitより前に置く。
'   条件そのものは OfferImportIfFirstRun 内部の3チェックが担保するので、
'   毎回呼ばれてもコストは ui_state 1回読みで済む=通常運用中は一瞬で戻る)。
'
' シートを写す実体(CopyUserData)は modMigrate.ImportUserData から移設した
' ものそのもの(挙動不変)。.xlsx 引き継ぎファイル方式(MIG_META 検査あり)と
' 旧xlsm直読み方式(MIG_META無し)の両方が「開いたブックから3シート+統計+
' config+受信箱を写す」という同じ仕事を必要とするための共用化であって、
' MIG_META の有無に依存する検査(ValidateMigFile)は modMigrate 側に残す。
'
' 防衛設計:
'   ・全Publicエントリの入口(OfferImportIfFirstRun)はサーキットブレーカー
'     (On Error)を持ち、失敗しても呼び出し元(起動シーケンス)を止めない。
'   ・旧版(R35以前)は Workbook_Open に自己インストーラを持つ。読み取り専用
'     で開くだけでは足りず、EnableEvents=False で開いている間だけそれを
'     止める(退避/復元はエラー経路を含む全経路で対にする。Finally型)。
'   ・config は modMigrate.CFG_KEEP_KEYS(利用者が決めた値のホワイトリスト)を
'     そのまま使う。ここで別表を持つと2つの台帳が食い違う事故(R12-3-2で
'     一度踏んだ「新版の既定値が旧値で巻き戻る」型)を再発させるため、
'     判定そのもの(IsUserOwnedKey/IsBuildOwnedKey)は modMigrate 側を
'     Public化して呼ぶ(値の複製ではなく判定関数の共用)。
' ============================================================================

Private Const ST_ASKED As String = "migfrom_asked"

' ----------------------------------------------------------------------------
' OfferImportIfFirstRun - 起点。3条件(本棚が空・未質問・前版あり)が揃った
'   ときだけ Yes/No を聞き、Yes なら ImportFromBook まで完了させる。
'   戻り値は「実際に引き継ぎ処理まで進んだか」(テスト・記録用。呼び出し元の
'   modTour は戻り値を見ない)。
' ----------------------------------------------------------------------------
Public Function OfferImportIfFirstRun() As Boolean
    On Error GoTo Fail
    OfferImportIfFirstRun = False

    ' 一番安いチェックから順に(通常運用中はここで即戻る)。
    If StrComp(modState.LoadState(ST_ASKED, ""), "1", vbBinaryCompare) = 0 Then Exit Function
    If ManifestRowCount() > 0 Then Exit Function

    Dim prevPath As String
    prevPath = FindPreviousBook(ThisWorkbook.Path, ThisWorkbook.Name)
    If LenB(prevPath) = 0 Then Exit Function

    ' §1-3契約: どちらを選んでも二度と聞かない。表示中に閉じられても
    ' 再質問し続けないよう、MsgBoxの【前】に書く(安全側)。
    modState.SaveState ST_ASKED, "1"

    Dim stamp As Double: stamp = CDbl(FileDateTime(prevPath))
    Dim msg As String: msg = OfferText(modUtil.FileNameOf(prevPath), stamp)
    If MsgBox(msg, vbYesNo + vbQuestion, modAppDef.APP_NAME) <> vbYes Then Exit Function

    Dim resultText As String: resultText = ImportFromBook(prevPath)
    MsgBox resultText, vbInformation, modAppDef.APP_NAME
    OfferImportIfFirstRun = True
    Exit Function

Fail:
    ' サーキットブレーカー: 単独で呼ばれても安全なように自前でも握る
    ' (modTour側は既に On Error Resume Next 配下だが、二重に防護する)。
End Function

' ----------------------------------------------------------------------------
' FindPreviousBook - folderPath 直下(第1階層)の "MyBookshelf*.xlsm" のうち
'   selfName 以外で、更新日時が最新のもののフルパス。無ければ空。
'   Dir() の再入禁止(modShelfScan.EnumFolderFiles と同じ作法)のため、
'   列挙結果を配列へ集めてから純関数 PickNewest へ渡す。
' ----------------------------------------------------------------------------
Public Function FindPreviousBook(ByVal folderPath As String, ByVal selfName As String) As String
    On Error Resume Next
    Dim folder As String: folder = modShelfScan.EnsureTrailingSlash(folderPath)
    If LenB(folder) = 0 Then Exit Function

    Dim names() As String: ReDim names(0 To 15)
    Dim stamps() As Double: ReDim stamps(0 To 15)
    Dim cnt As Long: cnt = 0

    Dim nm As String: nm = Dir$(folder & "MyBookshelf*.xlsm")
    Do While LenB(nm) > 0
        If cnt > UBound(names) Then
            ReDim Preserve names(0 To (UBound(names) + 1) * 2 - 1)
            ReDim Preserve stamps(0 To (UBound(stamps) + 1) * 2 - 1)
        End If
        names(cnt) = nm
        stamps(cnt) = CDbl(FileDateTime(folder & nm))
        cnt = cnt + 1
        nm = Dir$()
    Loop

    ' ReDim Preserveは条件に関わらず1度は必ず通す(条件付きだと0件経路だけ
    ' 丸ごと縮まらない罠がある。§10罠一覧)。
    If cnt = 0 Then
        ReDim Preserve names(0 To 0)
        ReDim Preserve stamps(0 To 0)
        Exit Function
    End If
    ReDim Preserve names(0 To cnt - 1)
    ReDim Preserve stamps(0 To cnt - 1)

    Dim idx As Long: idx = PickNewest(names, stamps, cnt, selfName)
    If idx < 0 Then Exit Function
    FindPreviousBook = folder & names(idx)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' PickNewest - 候補から「自分以外・最新・同時刻は先勝ち」の1件を選ぶ純関数
'   (Excel/COMに一切触れない)。該当なしは -1。
' ----------------------------------------------------------------------------
Public Function PickNewest(names() As String, stamps() As Double, ByVal n As Long, _
                            ByVal selfName As String) As Long
    Dim bestIdx As Long: bestIdx = -1
    Dim bestStamp As Double
    Dim i As Long
    For i = 0 To n - 1
        If StrComp(names(i), selfName, vbTextCompare) <> 0 Then
            If bestIdx < 0 Then
                bestIdx = i
                bestStamp = stamps(i)
            ElseIf stamps(i) > bestStamp Then   ' 同時刻は先勝ち(">"のみ。">="にしない)
                bestIdx = i
                bestStamp = stamps(i)
            End If
        End If
    Next i
    PickNewest = bestIdx
End Function

' ----------------------------------------------------------------------------
' OfferText - 引き継ぎ確認ダイアログの文面を組む純関数。
' ----------------------------------------------------------------------------
Public Function OfferText(ByVal fileName As String, ByVal stamp As Double) As String
    OfferText = "前の版の本棚が見つかりました。" & vbCrLf & vbCrLf & _
        fileName & "(更新: " & Format$(CDate(stamp), "yyyy-mm-dd hh:nn") & ")" & vbCrLf & vbCrLf & _
        "この版へ本棚・実績・設定を引き継ぎますか?" & vbCrLf & _
        "(「いいえ」を選ぶと、次回からは聞きません。あとから❓ヘルプの「" & _
        ChrW(&HD83D) & ChrW(&HDCE5) & " 引き継ぎファイルを読む」でも同じことができます)"
End Function

' ----------------------------------------------------------------------------
' ImportFromBook - path の旧版xlsmを読み取り専用で開き、CopyUserData で
'   本棚・実績・設定を写して閉じる。戻り値はトースト用の結果文(成功/失敗の
'   いずれも文言を返す。画面の再描画は呼び出し元(OfferImportIfFirstRun)に
'   任せる=ここでは行わない)。
' ----------------------------------------------------------------------------
Public Function ImportFromBook(ByVal path As String) As String
    Dim wb As Workbook
    Dim prevEvents As Boolean: prevEvents = Application.EnableEvents
    Dim prevAlerts As Boolean: prevAlerts = Application.DisplayAlerts

    ' F3a(R35実機第20報)と同じ作法: 開く前に先頭4バイトで暗号化かどうかを
    ' 判別する。暗号化/判別不能は Workbooks.Open を1回も呼ばずに倒す。
    Dim kind As String: kind = modShelfScan.EncryptedFileKind(path)
    Dim gate As String: gate = modPack.OpenGateReason(kind)
    If gate <> "open" Then
        modLog.LogError "E0801", "modMigrateFrom.ImportFromBook", _
            "開く前に判別して中止(kind=[" & kind & "]): " & path
        ImportFromBook = "前の版のファイルを読めませんでした(保護されている/読めない)。"
        Exit Function
    End If

    ' 2026-09-05(R36波4): 旧版(R35以前)はWorkbook_Openに自己インストーラを
    ' 持つ。EnableEvents=False で開いている間は走らない。退避/復元はエラー
    ' 経路を含む全経路で対にする(Finally型。modMigrate.ImportUserData と
    ' 同じ作法)。Password引数は付けない(この端末は普通のファイルでも
    ' Password引数付きだとOpen自体が失敗する。F3a実測)。
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    On Error GoTo OpenFail
    Set wb = Application.Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0, _
                                         IgnoreReadOnlyRecommended:=True, AddToMru:=False)
    On Error GoTo LoadFailed

    Dim srcDim As Long: srcDim = CLng(Val(ReadForeignConfigValue(wb, "embed_dim")))
    Dim myDim As Long: myDim = modConfig.GetLong("embed_dim", 1536)
    If srcDim > 0 And srcDim <> myDim Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
        Application.EnableEvents = prevEvents
        Application.DisplayAlerts = prevAlerts
        ImportFromBook = "AIの設定(次元数)が違うため引き継げません。"
        Exit Function
    End If

    Dim copied As Long: copied = CopyUserData(wb)

    wb.Close SaveChanges:=False
    Set wb = Nothing
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevAlerts

    If copied < 0 Then
        ImportFromBook = "前の版から一部を引き継げませんでした。もう一度お試しください。"
        Exit Function
    End If

    Dim got As Long
    On Error Resume Next
    got = modShelf.TotalChunks()
    modLog.LogUsage "migfrom_import", "", "chunks=" & got & " src=" & modUtil.FileNameOf(path)
    On Error GoTo 0

    ImportFromBook = "前の版から本棚を引き継ぎました(本棚: " & got & " 件)。" & vbCrLf & _
        "画面を描き直すため、一度閉じて開き直してください。"
    Exit Function

OpenFail:
    Dim openDesc As String: openDesc = Err.Description
    On Error GoTo 0
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevAlerts
    modLog.LogError "E0801", "modMigrateFrom.ImportFromBook", "ファイルを開けません: " & openDesc
    ImportFromBook = "前の版のファイルを開けませんでした。"
    Exit Function

LoadFailed:
    Dim loadDesc As String: loadDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かない。後始末の前に Resume で
    ' ハンドラを抜ける(modMigrate/modPack と同じ作法)。
    Resume LoadFailedCleanup
LoadFailedCleanup:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevAlerts
    modLog.LogError "E0801", "modMigrateFrom.ImportFromBook", "引き継ぎ中にエラー: " & loadDesc
    ImportFromBook = "前の版からの引き継ぎ中にエラーが発生しました。"
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' CopyUserData - 開いたブック(src)からこのブックへ資産を写す。
'   modMigrate.ImportUserData(.xlsx引き継ぎファイル経路)から移設した実体
'   そのもの(挙動不変)。旧xlsm直読み(ImportFromBook)と共用するための
'   受け皿。置換/併合方針は modMigrate 冒頭コメントのとおり(本棚3シートは
'   置換・my_statsはキー併合・configはCFG_KEEP_KEYSのみ・insight_inboxは
'   未取込行を追記)。戻り値は my_knowledge へ復元できた行数(「写した行数」)。
'   1枚でも RestoreSheet が失敗したら -1 を返す。
' ----------------------------------------------------------------------------
Public Function CopyUserData(ByVal src As Workbook) As Long
    Dim rowsK As Long, rowsV As Long, rowsM As Long
    Dim allOk As Boolean: allOk = True
    If Not RestoreSheet(src, modAppDef.SH_KNOWLEDGE, rowsK) Then allOk = False
    If Not RestoreSheet(src, modAppDef.SH_VECTORS, rowsV) Then allOk = False
    If Not RestoreSheet(src, modAppDef.SH_MANIFEST, rowsM) Then allOk = False

    ' 2026-08-10(R27波3-5)由来: 台帳はシートごと差し替わるので、書式も
    ' 引き継ぎ元のものになる。旧い版で作られた台帳には "@" が無く、移行した
    ' 直後から "-"始まりの資料名で取込が落ちる。ここで張り直す(冪等)。
    On Error Resume Next
    modShelfScan.EnsureManifestTextFormat
    On Error GoTo 0

    MergeStats src
    MergeConfig src
    MergeInsightInbox src

    If allOk Then
        CopyUserData = rowsK
    Else
        CopyUserData = -1
    End If
End Function

' ----------------------------------------------------------------------------
' 内部
' ----------------------------------------------------------------------------
Private Function ManifestRowCount() As Long
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_MANIFEST)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then Exit Function
    ManifestRowCount = lastR - 1
End Function

' 引き継ぎ元(src)の config シートから直接1キーを読む(modMigrate/modPack の
' ReadMetaValue と同じ走査。src は mig_meta を持たない生の製品ブックなので
' 対象シートは modAppDef.SH_CONFIG そのもの)。
Private Function ReadForeignConfigValue(ByVal wb As Workbook, ByVal key As String) As String
    On Error Resume Next
    Dim ws As Worksheet: Set ws = wb.Worksheets(modAppDef.SH_CONFIG)
    If ws Is Nothing Then Exit Function
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim i As Long
    For i = 2 To lastR
        If StrComp(CStr(ws.Cells(i, 1).Value), key, vbTextCompare) = 0 Then
            ReadForeignConfigValue = CStr(ws.Cells(i, 2).Value)
            Exit Function
        End If
    Next i
    On Error GoTo 0
End Function

Private Function SheetExistsIn(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExistsIn = Not (ws Is Nothing)
End Function

' 引き継ぎブックのシートで、このブックのシートを置き換える(modMigrateから
' 移設・挙動不変)。
'   outRows : 復元後に dst 側で【実際に】数えたデータ行数(見出し行を除く)。
'             書けたつもりの件数ではなく、書けた結果を数える。
Private Function RestoreSheet(ByVal srcWb As Workbook, ByVal sheetName As String, _
                              ByRef outRows As Long) As Boolean
    outRows = 0
    On Error GoTo Fail

    ' 引き継ぎ元側にそのシートが無いのは失敗ではない(旧版で存在しない
    ' シートがあり得る)。何も置き換えない。
    If Not SheetExistsIn(srcWb, sheetName) Then
        RestoreSheet = True
        Exit Function
    End If

    Dim src As Worksheet: Set src = srcWb.Worksheets(sheetName)
    ' 置き換え先が無い(=このブックが壊れている)場合は Worksheets(...) が
    ' 実行時エラー9を出し、下の Fail が記録する。黙って諦めない。
    Dim dst As Worksheet: Set dst = ThisWorkbook.Worksheets(sheetName)

    Dim lastR As Long: lastR = src.Cells(src.Rows.count, 1).End(xlUp).row
    Dim lastC As Long: lastC = src.Cells(1, src.Columns.count).End(xlToLeft).Column
    If lastC < 1 Then
        RestoreSheet = True
        Exit Function
    End If

    ' 既存を消してから入れる(混ぜない)。見出し行は入れ直すので全消しでよい。
    dst.Cells.ClearContents
    If lastR < 1 Then
        RestoreSheet = True
        Exit Function
    End If

    Dim rowStart As Long
    For rowStart = 1 To lastR Step 200
        Dim rowsN As Long
        rowsN = lastR - rowStart + 1
        If rowsN > 200 Then rowsN = 200
        dst.Range(dst.Cells(rowStart, 1), dst.Cells(rowStart + rowsN - 1, lastC)).Value = _
            src.Range(src.Cells(rowStart, 1), src.Cells(rowStart + rowsN - 1, lastC)).Value
    Next rowStart

    outRows = dst.Cells(dst.Rows.count, 1).End(xlUp).row - 1
    If outRows < 0 Then outRows = 0
    ' R12-H-6: 引き継ぎはシートを丸ごと入れ替える。my_vectors を入れ替えたら
    ' セッション内キャッシュは中身ごと別物なので、解放して世代を進める
    ' (件数も先頭/末尾idも偶然一致し得るため、印だけでは足りない)。
    If StrComp(sheetName, modAppDef.SH_VECTORS, vbTextCompare) = 0 Then
        On Error Resume Next
        modVecCache.ResetVecCache
        modVecCache.BumpGeneration
        On Error GoTo Fail
    End If
    RestoreSheet = True
    Exit Function

Fail:
    Dim failNum As Long: failNum = Err.Number
    Dim failDesc As String: failDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かない。記録の前に Resume で
    ' ハンドラを抜ける(modShelf.IngestFile と同じ作法。2026-07-30 実機err#462)。
    Resume FailCleanup
FailCleanup:
    On Error Resume Next
    modLog.LogError "E0801", "modMigrateFrom.CopyUserData", _
        "シート復元に失敗: " & sheetName & " err#" & failNum & ": " & failDesc
    On Error GoTo 0
End Function

' my_stats はキー単位で併合する(新版が先に書いた行を消さない)。
' (modMigrateから移設・挙動不変)
Private Sub MergeStats(ByVal srcWb As Workbook)
    On Error Resume Next
    If Not SheetExistsIn(srcWb, modAppDef.SH_STATS) Then Exit Sub
    Dim src As Worksheet: Set src = srcWb.Worksheets(modAppDef.SH_STATS)
    Dim lastR As Long: lastR = src.Cells(src.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then Exit Sub

    Dim arr As Variant
    arr = src.Range(src.Cells(2, 1), src.Cells(lastR, 2)).Value
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim k As String: k = Trim$(CStr(arr(i, 1)))
        If LenB(k) > 0 Then modStats.SetStatText k, CStr(arr(i, 2))
    Next i
    On Error GoTo 0
End Sub

' config は「この版に存在するキーだけ」値を引き継ぐ(modMigrateから移設・
' 挙動不変)。新版で増えたキーは新版の既定のまま(古い設定で新機能を殺さない)。
' 引き継ぎ元は .xlsx の config_user ではなく生の config シートなので、
' 判定(利用者が決めた値か/ビルドが決める値か)は modMigrate 側の
' IsUserOwnedKey/IsBuildOwnedKey(いずれもPublic化済み)をそのまま使う――
' 判定表を複製すると2つの台帳が食い違う事故(R12-3-2)を再発させるため。
Private Sub MergeConfig(ByVal srcWb As Workbook)
    On Error Resume Next
    If Not SheetExistsIn(srcWb, modAppDef.SH_CONFIG) Then Exit Sub
    Dim src As Worksheet: Set src = srcWb.Worksheets(modAppDef.SH_CONFIG)
    Dim lastR As Long: lastR = src.Cells(src.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then Exit Sub

    Dim me_ As Worksheet: Set me_ = ThisWorkbook.Worksheets(modAppDef.SH_CONFIG)
    If me_ Is Nothing Then Exit Sub
    Dim myLast As Long: myLast = me_.Cells(me_.Rows.count, 1).End(xlUp).row
    If myLast < 2 Then Exit Sub

    ' この版のキー集合(存在確認用)。
    Dim known As Object: Set known = CreateObject("Scripting.Dictionary")
    Dim myArr As Variant: myArr = me_.Range(me_.Cells(2, 1), me_.Cells(myLast, 1)).Value
    Dim j As Long
    For j = LBound(myArr, 1) To UBound(myArr, 1)
        Dim mk As String: mk = LCase$(Trim$(CStr(myArr(j, 1))))
        If LenB(mk) > 0 Then
            If Not known.Exists(mk) Then known.Add mk, True
        End If
    Next j

    Dim arr As Variant: arr = src.Range(src.Cells(2, 1), src.Cells(lastR, 2)).Value
    Dim i As Long, applied As Long, skipped As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim k As String: k = Trim$(CStr(arr(i, 1)))
        If LenB(k) > 0 Then
            If modMigrate.IsUserOwnedKey(k) And Not modMigrate.IsBuildOwnedKey(k) Then
                If known.Exists(LCase$(k)) Then
                    modConfig.SetValue k, CStr(arr(i, 2))
                    applied = applied + 1
                End If
            Else
                skipped = skipped + 1
            End If
        End If
    Next i
    modLog.LogUsage "migrate_config", "", "applied=" & applied & " kept_new_default=" & skipped
    On Error GoTo 0
End Sub

' insight_inbox は「未取込の行を追記」する(modMigrateから移設・挙動不変)。
' 置換ではない: 移行先ブックが起動後に自分で受信した共有知を消さないため、
' 既存の受信箱に合流させる。nonce(A列)が既に存在する行は二重取り込みしない。
Private Sub MergeInsightInbox(ByVal srcWb As Workbook)
    On Error Resume Next
    If Not SheetExistsIn(srcWb, "insight_inbox") Then Exit Sub
    Dim src As Worksheet: Set src = srcWb.Worksheets("insight_inbox")
    Dim srcLastR As Long: srcLastR = src.Cells(src.Rows.count, 1).End(xlUp).row
    If srcLastR < 2 Then Exit Sub

    Dim dst As Worksheet: Set dst = ThisWorkbook.Worksheets("insight_inbox")
    If dst Is Nothing Then Exit Sub
    Dim dstLastR As Long: dstLastR = dst.Cells(dst.Rows.count, 1).End(xlUp).row

    Dim lastC As Long: lastC = src.Cells(1, src.Columns.count).End(xlToLeft).Column
    If lastC < 1 Then lastC = 10

    ' 既存nonce(A列)の集合(重複取り込み防止)。
    Dim known As Object: Set known = CreateObject("Scripting.Dictionary")
    If dstLastR >= 2 Then
        Dim dstArr As Variant: dstArr = dst.Range(dst.Cells(2, 1), dst.Cells(dstLastR, 1)).Value
        Dim j As Long
        For j = LBound(dstArr, 1) To UBound(dstArr, 1)
            Dim dn As String: dn = CStr(dstArr(j, 1))
            If LenB(dn) > 0 Then
                If Not known.Exists(dn) Then known.Add dn, True
            End If
        Next j
    End If

    Dim srcArr As Variant: srcArr = src.Range(src.Cells(2, 1), src.Cells(srcLastR, lastC)).Value
    Dim outR As Long: outR = dstLastR + 1
    Dim added As Long, dup As Long, skipped As Long
    Dim i As Long, c As Long
    For i = LBound(srcArr, 1) To UBound(srcArr, 1)
        Dim nc As String: nc = CStr(srcArr(i, 1))
        If LenB(nc) > 0 And CStr(srcArr(i, 9)) <> "1" Then   ' consumed(念のため防御的に再確認)
            If known.Exists(nc) Then
                dup = dup + 1
            Else
                For c = 1 To lastC
                    dst.Cells(outR, c).Value = modMigrate.InboxCellValue(srcArr(i, c), c)   ' R12-H-4
                Next c
                known.Add nc, True
                outR = outR + 1
                added = added + 1
            End If
        Else
            skipped = skipped + 1
        End If
    Next i
    modLog.LogUsage "migrate_insight", "", "added=" & added & " dup_skip=" & dup & _
        " consumed_skip=" & skipped
    On Error GoTo 0
End Sub
