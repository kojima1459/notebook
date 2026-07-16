Attribute VB_Name = "modPack"
Option Explicit

' ============================================================================
' modPack - ナレッジパックの書き出し/取込(MASTER_SPEC §7.4)
' ----------------------------------------------------------------------------
' 役割:
'   本棚全体、またはアクティブ行の資料1件を、新規の.xlsx(マクロ無し)へ
'   pack_meta/pack_chunks/pack_vectors の3シートとして書き出す。取込側は
'   逆に他者のパックを読み込み、chunk_idのfnvハッシュ部で重複排除しながら
'   my_knowledge/my_vectorsへ追記する(origin="pack:"&作成者)。
'
' 設計判断(要点のみ。詳細は各関数直前のコメント参照):
'   ・検証は2つに分離(§7.4指示どおり): ValidatePack(wb,reason)がWorkbook
'     実物(シート存在+embed_dim一致)を検査し、内部で純関数ValidatePackMeta
'     (Long/Long/Stringのみ=テスト可能な「バージョン/次元」判定部)を呼ぶ。
'     自分のconfig embed_dimとの一致はValidatePack側で行う(ValidatePackMeta
'     の契約引数に「自分の次元」が無いため)。
'   ・「アクティブ行の資料1件」特定は、UI層(modUIShelf)のカード行レイア
'     ウトに依存せず、Application.ActiveCellの行にmy_manifestのfile_name
'     と一致する文字列が無いか探す疎結合な方式にした(R1・担当外不可侵)。
'   ・ImportPackDialog完了後の一覧再描画は、あえてmodPack側から呼ばない
'     (R1違反回避)。ボタンラッパー(modUIShelf.OnImportPack、担当外)の責務。
'   ・PIIスキャン(E0703)は書き出し前に全チャンクへmodPii.ScanTextを適用し、
'     検知時は件数+例示を出してYes/Noで続行確認する(警告であり停止ではない)。
'   ・重複排除はchunk_idの"bs::HASH::pN::cN"からHASH部を比較する
'     (modShelf.bas BuildExistingHashSetと同じ設計。非Publicのため複製)。
'   ・保存失敗(読み取り専用フォルダ等・§13)はOn Error捕捉。§6にこのケース
'     専用コードが無いため新規コードは追加せず(§6は契約として変更禁止)、
'     err_logには識別用タグ"PACK_SAVE_FAILED"で記録し、ユーザーには2文
'     構成の案内を直接表示する。
'   ・my_knowledge/my_vectorsへの読み書きは配列一括(§12)。
' ============================================================================

Private Const COL_ID As Long = 1
Private Const COL_SOURCE As Long = 2
Private Const COL_ORIGIN As Long = 3
Private Const COL_PAGE As Long = 4
Private Const COL_SUMMARY As Long = 5
Private Const COL_KEYWORDS As Long = 6
Private Const COL_FULLTEXT As Long = 7
Private Const COL_ADDED As Long = 8
Private Const COL_EMBEDDED As Long = 9

' ExportPackDialog - 本棚全体 or アクティブ行の資料1件を選び、新規.xlsxへ
'   pack_meta/pack_chunks/pack_vectors を書き出す。
Public Sub ExportPackDialog()
    Dim scopeAns As Long
    scopeAns = MsgBox( _
        "本棚全体をパックにして書き出しますか?" & vbLf & _
        "「いいえ」を選ぶと、選択中の行(アクティブセル)の資料1件だけを書き出します。", _
        vbYesNoCancel + vbQuestion, modAppDef.APP_NAME)
    If scopeAns = vbCancel Then Exit Sub

    Dim sourceFilter As String   ' 空="本棚全体"
    If scopeAns = vbNo Then
        sourceFilter = FindActiveRowSourceName()
        If LenB(sourceFilter) = 0 Then
            MsgBox "書き出す資料を特定できませんでした。" & vbLf & _
                "マイ本棚の資料カードの行をクリックしてから、もう一度お試しください。", _
                vbExclamation, modAppDef.APP_NAME
            Exit Sub
        End If
    End If

    Dim ids() As String, sources() As String, pages() As Long
    Dim summaries() As String, keywords() As String, fullTexts() As String
    Dim n As Long
    n = LoadChunksForExport(sourceFilter, ids, sources, pages, summaries, keywords, fullTexts)
    If n = 0 Then
        MsgBox "書き出せる資料がありません。" & vbLf & _
            "マイ本棚に資料を追加してから、もう一度お試しください。", vbExclamation, modAppDef.APP_NAME
        Exit Sub
    End If

    ' 書き出し前PIIスキャン(全チャンク・E0703)
    Dim piiCount As Long, piiExamples As String
    ScanChunksForPii sources, fullTexts, n, piiCount, piiExamples

    If piiCount > 0 Then
        Dim promptMsg As String
        promptMsg = modLog.FriendlyMessage("E0703") & vbLf & vbLf & _
            "件数: " & piiCount & "件" & vbLf & _
            "例: " & piiExamples & vbLf & vbLf & _
            "このまま書き出しを続けますか?" & vbLf & "(コード: E0703)"
        Dim contAns As Long
        contAns = MsgBox(promptMsg, vbYesNo + vbExclamation, modAppDef.APP_NAME)
        modLog.LogError "E0703", "modPack.ExportPackDialog", "件数=" & piiCount & " 例=" & piiExamples
        If contAns = vbNo Then Exit Sub
    End If

    ' 保存先ダイアログ
    Dim fd As Object
    Set fd = Application.FileDialog(2)   ' msoFileDialogSaveAs(名前付き定数は使わない)
    fd.Title = "パックの保存先を選んでください"
    Dim defaultName As String
    If LenB(sourceFilter) > 0 Then
        defaultName = "パック_" & sourceFilter & "_" & Format$(Now, "yyyymmdd")
    Else
        defaultName = "パック_本棚全体_" & Format$(Now, "yyyymmdd")
    End If
    fd.InitialFileName = defaultName & ".xlsx"
    If fd.Show <> -1 Then Exit Sub   ' キャンセル
    If fd.SelectedItems.count < 1 Then Exit Sub

    Dim savePath As String: savePath = CStr(fd.SelectedItems(1))
    If LCase$(modUtil.ExtOf(savePath)) <> "xlsx" Then savePath = savePath & ".xlsx"

    Dim authorName As String: authorName = Trim$(modConfig.GetString("pack_author", ""))
    If LenB(authorName) = 0 Then authorName = "不明"

    Dim newWb As Workbook
    On Error GoTo SaveFail
    Set newWb = Application.Workbooks.Add
    RemoveExtraSheets newWb

    Dim wsMeta As Worksheet: Set wsMeta = newWb.Worksheets(1)
    wsMeta.Name = "pack_meta"
    WritePackMeta wsMeta, authorName, n

    Dim wsChunks As Worksheet: Set wsChunks = newWb.Worksheets.Add(After:=wsMeta)
    wsChunks.Name = "pack_chunks"
    WritePackChunks wsChunks, ids, sources, pages, summaries, keywords, fullTexts, n

    Dim wsVec As Worksheet: Set wsVec = newWb.Worksheets.Add(After:=wsChunks)
    wsVec.Name = "pack_vectors"
    WritePackVectors wsVec, ids, n

    Application.DisplayAlerts = False
    newWb.SaveAs Filename:=savePath, FileFormat:=51   ' xlOpenXMLWorkbook(.xlsx・マクロ無し)
    Application.DisplayAlerts = True
    newWb.Close SaveChanges:=False
    On Error GoTo 0

    modStats.Bump "pack_export_total"
    modLog.LogUsage "pack_export", "", "n=" & n & " scope=" & IIf(LenB(sourceFilter) = 0, "all", sourceFilter)

    MsgBox "パックを書き出しました。" & vbLf & _
        n & "件の資料を「" & modUtil.FileNameOf(savePath) & "」として保存しました。", _
        vbInformation, modAppDef.APP_NAME
    Exit Sub

SaveFail:
    Dim saveErrDetail As String: saveErrDetail = Err.Description
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not newWb Is Nothing Then newWb.Close SaveChanges:=False
    On Error GoTo 0
    modLog.LogError "PACK_SAVE_FAILED", "modPack.ExportPackDialog", saveErrDetail
    MsgBox "書き出し先にファイルを保存できませんでした(読み取り専用のフォルダや権限不足の可能性があります)。" & vbLf & _
        "別の保存先を選ぶか、書き込み権限があるか確認してから、もう一度お試しください。", _
        vbExclamation, modAppDef.APP_NAME
End Sub

' ImportPackDialog - ファイル選択→ValidatePack→重複(fnvハッシュ)スキップ
'   しつつmy_knowledge/my_vectorsへ取込。origin="pack:"&作成者。
Public Sub ImportPackDialog()
    Dim fd As Object
    Set fd = Application.FileDialog(3)   ' msoFileDialogFilePicker(名前付き定数は使わない)
    fd.AllowMultiSelect = False
    fd.Title = "取り込むパックファイルを選んでください"
    fd.Filters.Clear
    fd.Filters.Add "Excelブック", "*.xlsx"

    If fd.Show <> -1 Then Exit Sub   ' キャンセル
    If fd.SelectedItems.count < 1 Then Exit Sub
    Dim packPath As String: packPath = CStr(fd.SelectedItems(1))

    Dim wb As Workbook
    On Error GoTo OpenFail
    Set wb = Application.Workbooks.Open(Filename:=packPath, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo 0

    Dim reason As String
    If Not ValidatePack(wb, reason) Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0

        Dim badCode As String
        If InStr(reason, "次元") > 0 Then
            badCode = "E0702"
        Else
            badCode = "E0701"
        End If
        modLog.LogError badCode, "modPack.ImportPackDialog", reason
        MsgBox modLog.FriendlyMessage(badCode) & vbLf & "(コード: " & badCode & ")" & vbLf & _
            "詳細: " & reason, vbExclamation, modAppDef.APP_NAME
        Exit Sub
    End If

    Dim authorName As String
    authorName = Trim$(ReadMetaValue(wb.Worksheets("pack_meta"), "author"))
    If LenB(authorName) = 0 Then authorName = "不明"

    Dim ids() As String, sources() As String, pages() As Long
    Dim summaries() As String, keywords() As String, fullTexts() As String, vectors() As String
    Dim n As Long
    n = LoadPackChunksAndVectors(wb, ids, sources, pages, summaries, keywords, fullTexts, vectors)

    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0

    If n = 0 Then
        MsgBox "パックの中に取込める資料がありませんでした。" & vbLf & _
            "別のパックファイルでお試しください。", vbExclamation, modAppDef.APP_NAME
        Exit Sub
    End If

    Dim importedCount As Long, skippedCount As Long
    ImportChunksDedup ids, sources, pages, summaries, keywords, fullTexts, vectors, n, authorName, _
        importedCount, skippedCount

    modStats.Bump "pack_import_total"
    modLog.LogUsage "pack_import", "", "imported=" & importedCount & " skipped=" & skippedCount & _
        " author=" & authorName

    MsgBox importedCount & "件取込 / " & skippedCount & "件は既にありました。", _
        vbInformation, modAppDef.APP_NAME
    Exit Sub

OpenFail:
    On Error GoTo 0
    modLog.LogError "E0701", "modPack.ImportPackDialog", "ファイルを開けません: " & Err.Description
    MsgBox "パックファイルを開けませんでした。" & vbLf & _
        "壊れたファイルでないか、パスに間違いがないか確認してください。" & vbLf & "(コード: E0701)", _
        vbExclamation, modAppDef.APP_NAME
End Sub

' ValidatePack - Workbook実物を検査(3シート存在・pack_format_version一致・
'   embed_dim一致)。E0701/E0702の判定材料をreasonに返す。
Public Function ValidatePack(ByVal wb As Workbook, ByRef reason As String) As Boolean
    reason = ""
    If wb Is Nothing Then
        reason = "ブックを読み込めませんでした"
        Exit Function
    End If

    If Not SheetExistsIn(wb, "pack_meta") Or Not SheetExistsIn(wb, "pack_chunks") _
            Or Not SheetExistsIn(wb, "pack_vectors") Then
        reason = "必要なシート(pack_meta/pack_chunks/pack_vectors)が見つかりません"
        Exit Function
    End If

    Dim wsMeta As Worksheet: Set wsMeta = wb.Worksheets("pack_meta")
    Dim fv As Long: fv = CLngSafe(ReadMetaValue(wsMeta, "pack_format_version"))
    Dim ed As Long: ed = CLngSafe(ReadMetaValue(wsMeta, "embed_dim"))

    Dim metaReason As String
    If Not ValidatePackMeta(fv, ed, metaReason) Then
        reason = metaReason
        Exit Function
    End If

    Dim myDim As Long: myDim = modConfig.GetLong("embed_dim", 1536)
    If ed <> myDim Then
        reason = "ベクトルの次元数が一致しません(パック=" & ed & " / このブック=" & myDim & ")"
        Exit Function
    End If

    ValidatePack = True
End Function

' ValidatePackMeta - バージョン/次元の判定ロジックを切り出した純関数(§7.8)。
'   Excelオブジェクトに触れない(Long/Long/Stringのみ)。形式バージョンが
'   対応版か・次元数が妥当か(1以上)だけを見る。「自分のconfig embed_dimと
'   の一致」はValidatePackが別途判定する(上記コメント参照)。
Public Function ValidatePackMeta(ByVal formatVersion As Long, ByVal embedDim As Long, ByRef reason As String) As Boolean
    reason = ""
    If formatVersion <> modAppDef.PACK_FORMAT_VERSION Then
        reason = "パックの形式バージョン(" & formatVersion & ")が対応バージョン(" & _
            modAppDef.PACK_FORMAT_VERSION & ")と一致しません"
        Exit Function
    End If
    If embedDim < 1 Then
        reason = "パックのベクトル次元数(" & embedDim & ")が不正です"
        Exit Function
    End If
    ValidatePackMeta = True
End Function

' 内部ヘルパー: 共通
Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function CLngSafe(ByVal s As String) As Long
    If IsNumeric(s) Then CLngSafe = CLng(Val(s))
End Function

' アクティブセルの行の中からmy_manifestのfile_nameと一致する文字列を探す
' (カード行の列レイアウトに依存しない特定方法。上部コメント参照)。
Private Function FindActiveRowSourceName() As String
    Dim wsM As Worksheet: Set wsM = GetSheet(modAppDef.SH_MANIFEST)
    If wsM Is Nothing Then Exit Function
    Dim lastM As Long: lastM = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
    If lastM < 2 Then Exit Function

    Dim names As Object: Set names = CreateObject("Scripting.Dictionary")
    If lastM = 2 Then
        AddNameIfNotEmpty names, CStr(wsM.Cells(2, 2).Value)
    Else
        Dim arrM As Variant: arrM = wsM.Range(wsM.Cells(2, 2), wsM.Cells(lastM, 2)).Value
        Dim i As Long
        For i = LBound(arrM, 1) To UBound(arrM, 1)
            AddNameIfNotEmpty names, CStr(arrM(i, 1))
        Next i
    End If
    If names.count = 0 Then Exit Function

    Dim wsShelf As Worksheet: Set wsShelf = GetSheet(modAppDef.SH_SHELF)
    If wsShelf Is Nothing Then Exit Function

    Dim r As Long
    On Error Resume Next
    r = Application.ActiveCell.row
    On Error GoTo 0
    If r < 1 Then Exit Function

    Dim rowVals As Variant
    rowVals = wsShelf.Range(wsShelf.Cells(r, 1), wsShelf.Cells(r, 30)).Value

    Dim c As Long
    For c = LBound(rowVals, 2) To UBound(rowVals, 2)
        Dim v As String: v = CStr(rowVals(1, c))
        If LenB(v) > 0 Then
            Dim key As Variant
            For Each key In names.Keys
                If StrComp(CStr(key), v, vbTextCompare) = 0 Then
                    FindActiveRowSourceName = CStr(key)
                    Exit Function
                End If
            Next key
        End If
    Next c
End Function

Private Sub AddNameIfNotEmpty(ByVal dict As Object, ByVal nm As String)
    If LenB(nm) = 0 Then Exit Sub
    If Not dict.Exists(nm) Then dict.Add nm, True
End Sub

Private Function SheetExistsIn(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExistsIn = Not (ws Is Nothing)
End Function

Private Function ReadMetaValue(ByVal wsMeta As Worksheet, ByVal key As String) As String
    Dim lastR As Long: lastR = wsMeta.Cells(wsMeta.Rows.count, 1).End(xlUp).row
    If lastR < 1 Then Exit Function
    If lastR = 1 Then
        If StrComp(CStr(wsMeta.Cells(1, 1).Value), key, vbTextCompare) = 0 Then
            ReadMetaValue = CStr(wsMeta.Cells(1, 2).Value)
        End If
        Exit Function
    End If

    Dim arr As Variant: arr = wsMeta.Range(wsMeta.Cells(1, 1), wsMeta.Cells(lastR, 2)).Value
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(CStr(arr(i, 1)), key, vbTextCompare) = 0 Then
            ReadMetaValue = CStr(arr(i, 2))
            Exit Function
        End If
    Next i
End Function

' 内部ヘルパー: 書き出し(ExportPackDialog)
Private Function LoadChunksForExport(ByVal sourceFilter As String, ByRef ids() As String, ByRef sources() As String, _
        ByRef pages() As Long, ByRef summaries() As String, ByRef keywords() As String, _
        ByRef fullTexts() As String) As Long
    ReDim ids(0 To 0): ReDim sources(0 To 0): ReDim pages(0 To 0)
    ReDim summaries(0 To 0): ReDim keywords(0 To 0): ReDim fullTexts(0 To 0)

    Dim wsK As Worksheet: Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If wsK Is Nothing Then Exit Function
    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then Exit Function

    Dim arr As Variant: arr = wsK.Range(wsK.Cells(2, 1), wsK.Cells(lastK, 9)).Value
    Dim nRows As Long: nRows = UBound(arr, 1) - LBound(arr, 1) + 1

    Dim tmpIds() As String: ReDim tmpIds(0 To nRows - 1)
    Dim tmpSources() As String: ReDim tmpSources(0 To nRows - 1)
    Dim tmpPages() As Long: ReDim tmpPages(0 To nRows - 1)
    Dim tmpSummaries() As String: ReDim tmpSummaries(0 To nRows - 1)
    Dim tmpKeywords() As String: ReDim tmpKeywords(0 To nRows - 1)
    Dim tmpFullTexts() As String: ReDim tmpFullTexts(0 To nRows - 1)
    Dim cnt As Long: cnt = 0

    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim srcName As String: srcName = CStr(arr(i, COL_SOURCE))
        If LenB(sourceFilter) = 0 Or StrComp(srcName, sourceFilter, vbTextCompare) = 0 Then
            tmpIds(cnt) = CStr(arr(i, COL_ID))
            tmpSources(cnt) = srcName
            If IsNumeric(arr(i, COL_PAGE)) Then tmpPages(cnt) = CLng(arr(i, COL_PAGE))
            tmpSummaries(cnt) = CStr(arr(i, COL_SUMMARY))
            tmpKeywords(cnt) = CStr(arr(i, COL_KEYWORDS))
            tmpFullTexts(cnt) = CStr(arr(i, COL_FULLTEXT))
            cnt = cnt + 1
        End If
    Next i

    If cnt > 0 Then
        ReDim Preserve tmpIds(0 To cnt - 1)
        ReDim Preserve tmpSources(0 To cnt - 1)
        ReDim Preserve tmpPages(0 To cnt - 1)
        ReDim Preserve tmpSummaries(0 To cnt - 1)
        ReDim Preserve tmpKeywords(0 To cnt - 1)
        ReDim Preserve tmpFullTexts(0 To cnt - 1)
        ids = tmpIds: sources = tmpSources: pages = tmpPages
        summaries = tmpSummaries: keywords = tmpKeywords: fullTexts = tmpFullTexts
    End If
    LoadChunksForExport = cnt
End Function

Private Sub ScanChunksForPii(sources() As String, fullTexts() As String, ByVal n As Long, _
        ByRef outCount As Long, ByRef outExamples As String)
    outCount = 0
    Dim exArr() As String: ReDim exArr(0 To 2)
    Dim exCnt As Long: exCnt = 0

    Dim i As Long
    For i = 0 To n - 1
        Dim cats As String: cats = modPii.ScanText(fullTexts(i))
        If LenB(cats) > 0 Then
            outCount = outCount + 1
            If exCnt < 3 Then
                exArr(exCnt) = sources(i) & "(" & cats & ")"
                exCnt = exCnt + 1
            End If
        End If
    Next i

    If exCnt > 0 Then
        ReDim Preserve exArr(0 To exCnt - 1)
        outExamples = Join(exArr, " / ")
    Else
        outExamples = ""
    End If
End Sub

Private Sub RemoveExtraSheets(ByVal wb As Workbook)
    Application.DisplayAlerts = False
    Do While wb.Worksheets.count > 1
        wb.Worksheets(wb.Worksheets.count).Delete
    Loop
    Application.DisplayAlerts = True
End Sub

Private Sub WritePackMeta(ByVal ws As Worksheet, ByVal authorName As String, ByVal chunkCount As Long)
    Dim rows(1 To 9, 1 To 2) As Variant
    rows(1, 1) = "key": rows(1, 2) = "value"
    rows(2, 1) = "pack_format_version": rows(2, 2) = modAppDef.PACK_FORMAT_VERSION
    rows(3, 1) = "pack_name": rows(3, 2) = "マイ本棚パック_" & Format$(Now, "yyyymmdd")
    rows(4, 1) = "author": rows(4, 2) = authorName
    rows(5, 1) = "created_at": rows(5, 2) = modUtil.NowStamp()
    rows(6, 1) = "description": rows(6, 2) = ""
    rows(7, 1) = "embed_dim": rows(7, 2) = modConfig.GetLong("embed_dim", 1536)
    rows(8, 1) = "app_version": rows(8, 2) = modAppDef.APP_VERSION
    rows(9, 1) = "chunk_count": rows(9, 2) = chunkCount
    ws.Range(ws.Cells(1, 1), ws.Cells(9, 2)).Value = rows
End Sub

Private Sub WritePackChunks(ByVal ws As Worksheet, ids() As String, sources() As String, pages() As Long, _
        summaries() As String, keywords() As String, fullTexts() As String, ByVal n As Long)
    Dim hdr As Variant: hdr = ChunkHeaderNames()
    Dim h As Long
    For h = LBound(hdr) To UBound(hdr)
        ws.Cells(1, h + 1).Value = hdr(h)
    Next h

    Dim arr() As Variant: ReDim arr(1 To n, 1 To 6)
    Dim i As Long
    For i = 1 To n
        arr(i, 1) = ids(i - 1)
        arr(i, 2) = sources(i - 1)
        arr(i, 3) = pages(i - 1)
        arr(i, 4) = summaries(i - 1)
        arr(i, 5) = keywords(i - 1)
        arr(i, 6) = modUtil.SafeLeft(fullTexts(i - 1), 32000)
    Next i
    ws.Range(ws.Cells(2, 1), ws.Cells(1 + n, 6)).Value = arr
End Sub

Private Function ChunkHeaderNames() As Variant
    ChunkHeaderNames = Array("chunk_id", "source", "page", "summary", "keywords", "full_text")
End Function

Private Sub WritePackVectors(ByVal ws As Worksheet, ids() As String, ByVal n As Long)
    ws.Cells(1, 1).Value = "chunk_id"
    ws.Cells(1, 2).Value = "vector_csv"

    Dim vecMap As Object: Set vecMap = BuildVectorMap()

    Dim arr() As Variant: ReDim arr(1 To n, 1 To 2)
    Dim i As Long
    For i = 1 To n
        arr(i, 1) = ids(i - 1)
        Dim csvVal As String: csvVal = ""
        If vecMap.Exists(ids(i - 1)) Then csvVal = CStr(vecMap(ids(i - 1)))
        arr(i, 2) = modUtil.SafeLeft(csvVal, 32000)
    Next i
    ws.Range(ws.Cells(2, 1), ws.Cells(1 + n, 2)).Value = arr
End Sub

Private Function BuildVectorMap() As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim wsV As Worksheet: Set wsV = GetSheet(modAppDef.SH_VECTORS)
    If wsV Is Nothing Then
        Set BuildVectorMap = dict
        Exit Function
    End If
    Dim lastV As Long: lastV = wsV.Cells(wsV.Rows.count, 1).End(xlUp).row
    If lastV < 2 Then
        Set BuildVectorMap = dict
        Exit Function
    End If

    Dim arr As Variant: arr = wsV.Range(wsV.Cells(2, 1), wsV.Cells(lastV, 2)).Value
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim cid As String: cid = CStr(arr(i, 1))
        If LenB(cid) > 0 Then
            If Not dict.Exists(cid) Then dict.Add cid, CStr(arr(i, 2))
        End If
    Next i
    Set BuildVectorMap = dict
End Function

' 内部ヘルパー: 取込(ImportPackDialog)
Private Function LoadPackChunksAndVectors(ByVal wb As Workbook, ByRef ids() As String, ByRef sources() As String, _
        ByRef pages() As Long, ByRef summaries() As String, ByRef keywords() As String, _
        ByRef fullTexts() As String, ByRef vectors() As String) As Long
    ReDim ids(0 To 0): ReDim sources(0 To 0): ReDim pages(0 To 0)
    ReDim summaries(0 To 0): ReDim keywords(0 To 0): ReDim fullTexts(0 To 0): ReDim vectors(0 To 0)

    Dim wsC As Worksheet: Set wsC = wb.Worksheets("pack_chunks")
    Dim lastC As Long: lastC = wsC.Cells(wsC.Rows.count, 1).End(xlUp).row
    If lastC < 2 Then Exit Function

    Dim arrC As Variant: arrC = wsC.Range(wsC.Cells(2, 1), wsC.Cells(lastC, 6)).Value
    Dim n As Long: n = UBound(arrC, 1) - LBound(arrC, 1) + 1

    Dim wsV As Worksheet: Set wsV = wb.Worksheets("pack_vectors")
    Dim vecMap As Object: Set vecMap = CreateObject("Scripting.Dictionary")
    Dim lastV As Long: lastV = wsV.Cells(wsV.Rows.count, 1).End(xlUp).row
    If lastV >= 2 Then
        Dim arrV As Variant: arrV = wsV.Range(wsV.Cells(2, 1), wsV.Cells(lastV, 2)).Value
        Dim j As Long
        For j = LBound(arrV, 1) To UBound(arrV, 1)
            Dim vid As String: vid = CStr(arrV(j, 1))
            If LenB(vid) > 0 Then
                If Not vecMap.Exists(vid) Then vecMap.Add vid, CStr(arrV(j, 2))
            End If
        Next j
    End If

    Dim tmpIds() As String: ReDim tmpIds(0 To n - 1)
    Dim tmpSources() As String: ReDim tmpSources(0 To n - 1)
    Dim tmpPages() As Long: ReDim tmpPages(0 To n - 1)
    Dim tmpSummaries() As String: ReDim tmpSummaries(0 To n - 1)
    Dim tmpKeywords() As String: ReDim tmpKeywords(0 To n - 1)
    Dim tmpFullTexts() As String: ReDim tmpFullTexts(0 To n - 1)
    Dim tmpVectors() As String: ReDim tmpVectors(0 To n - 1)
    Dim cnt As Long: cnt = 0

    Dim i As Long
    For i = LBound(arrC, 1) To UBound(arrC, 1)
        Dim cid As String: cid = CStr(arrC(i, 1))
        If LenB(cid) > 0 Then
            tmpIds(cnt) = cid
            tmpSources(cnt) = CStr(arrC(i, 2))
            If IsNumeric(arrC(i, 3)) Then tmpPages(cnt) = CLng(arrC(i, 3))
            tmpSummaries(cnt) = CStr(arrC(i, 4))
            tmpKeywords(cnt) = CStr(arrC(i, 5))
            tmpFullTexts(cnt) = CStr(arrC(i, 6))
            If vecMap.Exists(cid) Then
                tmpVectors(cnt) = CStr(vecMap(cid))
            Else
                tmpVectors(cnt) = ""
            End If
            cnt = cnt + 1
        End If
    Next i

    If cnt > 0 Then
        ReDim Preserve tmpIds(0 To cnt - 1)
        ReDim Preserve tmpSources(0 To cnt - 1)
        ReDim Preserve tmpPages(0 To cnt - 1)
        ReDim Preserve tmpSummaries(0 To cnt - 1)
        ReDim Preserve tmpKeywords(0 To cnt - 1)
        ReDim Preserve tmpFullTexts(0 To cnt - 1)
        ReDim Preserve tmpVectors(0 To cnt - 1)
        ids = tmpIds: sources = tmpSources: pages = tmpPages
        summaries = tmpSummaries: keywords = tmpKeywords: fullTexts = tmpFullTexts
        vectors = tmpVectors
    End If
    LoadPackChunksAndVectors = cnt
End Function

Private Sub ImportChunksDedup(ids() As String, sources() As String, pages() As Long, summaries() As String, _
        keywords() As String, fullTexts() As String, vectors() As String, ByVal n As Long, ByVal authorName As String, _
        ByRef importedCount As Long, ByRef skippedCount As Long)
    importedCount = 0
    skippedCount = 0

    Dim wsK As Worksheet: Set wsK = EnsureKnowledgeSheet()
    Dim wsV As Worksheet: Set wsV = EnsureVectorSheet()
    If wsK Is Nothing Or wsV Is Nothing Then Exit Sub

    Dim existingHashes As Object: Set existingHashes = BuildExistingHashSet(wsK)

    Dim outK() As Variant: ReDim outK(1 To n, 1 To 9)
    Dim outV() As Variant: ReDim outV(1 To n, 1 To 2)
    Dim addedStamp As String: addedStamp = modUtil.NowStamp()
    Dim originStr As String: originStr = "pack:" & authorName

    Dim i As Long
    For i = 0 To n - 1
        Dim hashHex As String: hashHex = ExtractHashFromChunkId(ids(i))
        If LenB(hashHex) = 0 Then hashHex = modUtil.Fnv1a64Hex(modUtil.NormalizeForHash(fullTexts(i)))

        If existingHashes.Exists(hashHex) Then
            skippedCount = skippedCount + 1
        Else
            existingHashes.Add hashHex, True
            importedCount = importedCount + 1
            outK(importedCount, COL_ID) = ids(i)
            outK(importedCount, COL_SOURCE) = sources(i)
            outK(importedCount, COL_ORIGIN) = originStr
            outK(importedCount, COL_PAGE) = pages(i)
            outK(importedCount, COL_SUMMARY) = summaries(i)
            outK(importedCount, COL_KEYWORDS) = keywords(i)
            outK(importedCount, COL_FULLTEXT) = modUtil.SafeLeft(fullTexts(i), 32000)
            outK(importedCount, COL_ADDED) = addedStamp
            outK(importedCount, COL_EMBEDDED) = IIf(LenB(vectors(i)) > 0, 1, 0)

            outV(importedCount, 1) = ids(i)
            outV(importedCount, 2) = modUtil.SafeLeft(vectors(i), 32000)
        End If
    Next i

    If importedCount > 0 Then
        Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
        If lastK < 1 Then lastK = 1
        Dim firstKRow As Long: firstKRow = lastK + 1
        If firstKRow < 2 Then firstKRow = 2

        Dim wArrK As Variant: wArrK = CompactRows(outK, importedCount, 9)
        wsK.Range(wsK.Cells(firstKRow, 1), wsK.Cells(firstKRow + importedCount - 1, 9)).Value = wArrK

        Dim lastV As Long: lastV = wsV.Cells(wsV.Rows.count, 1).End(xlUp).row
        If lastV < 1 Then lastV = 1
        Dim firstVRow As Long: firstVRow = lastV + 1
        If firstVRow < 2 Then firstVRow = 2

        Dim wArrV As Variant: wArrV = CompactRows(outV, importedCount, 2)
        wsV.Range(wsV.Cells(firstVRow, 1), wsV.Cells(firstVRow + importedCount - 1, 2)).Value = wArrV
    End If
End Sub

' my_knowledge/my_vectors共通: 無ければveryHiddenで新規作成しヘッダを書く
' (modShelf.bas/modEmbed.bas の同名Ensure*Sheetと同じ設計。複製理由は上部コメント参照)。
Private Function EnsureVeryHiddenSheet(ByVal sheetName As String, ByVal hdr As Variant) As Worksheet
    Dim ws As Worksheet: Set ws = GetSheet(sheetName)
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = sheetName
        Dim i As Long
        For i = LBound(hdr) To UBound(hdr)
            ws.Cells(1, i + 1).Value = hdr(i)
        Next i
        On Error Resume Next
        ws.Visible = 2   ' xlSheetVeryHidden
        On Error GoTo 0
    End If
    Set EnsureVeryHiddenSheet = ws
    Exit Function
Fail:
    Set EnsureVeryHiddenSheet = Nothing
End Function

Private Function EnsureKnowledgeSheet() As Worksheet
    Set EnsureKnowledgeSheet = EnsureVeryHiddenSheet(modAppDef.SH_KNOWLEDGE, _
        Array("chunk_id", "source", "origin", "page", "summary", "keywords", "full_text", "added_at", "embedded"))
End Function

Private Function EnsureVectorSheet() As Worksheet
    Set EnsureVectorSheet = EnsureVeryHiddenSheet(modAppDef.SH_VECTORS, Array("chunk_id", "vector_csv"))
End Function

' chunk_id「bs::HASH::pN::cN」のHASH部分を集めたDictionaryを返す
' (modShelf.bas BuildExistingHashSetと同じ設計。複製理由は上部コメント参照)。
Private Function BuildExistingHashSet(ByVal wsK As Worksheet) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    If wsK Is Nothing Then
        Set BuildExistingHashSet = dict
        Exit Function
    End If

    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then
        Set BuildExistingHashSet = dict
        Exit Function
    End If

    If lastK = 2 Then
        AddHashIfNew dict, ExtractHashFromChunkId(CStr(wsK.Cells(2, 1).Value))
    Else
        Dim ids As Variant: ids = wsK.Range(wsK.Cells(2, 1), wsK.Cells(lastK, 1)).Value
        Dim i As Long
        For i = LBound(ids, 1) To UBound(ids, 1)
            AddHashIfNew dict, ExtractHashFromChunkId(CStr(ids(i, 1)))
        Next i
    End If
    Set BuildExistingHashSet = dict
End Function

Private Sub AddHashIfNew(ByVal dict As Object, ByVal h As String)
    If LenB(h) = 0 Then Exit Sub
    If Not dict.Exists(h) Then dict.Add h, True
End Sub

Private Function ExtractHashFromChunkId(ByVal chunkId As String) As String
    If LenB(chunkId) = 0 Then Exit Function
    Dim parts() As String: parts = Split(chunkId, "::")
    If UBound(parts) - LBound(parts) + 1 < 2 Then Exit Function
    ExtractHashFromChunkId = parts(LBound(parts) + 1)
End Function

' src(1 To n以上, 1 To cols)の先頭n行だけを(1 To n, 1 To cols)へ詰め直す
' (Range書込みは配列サイズが対象範囲と一致していないといけないため)。
Private Function CompactRows(ByRef src As Variant, ByVal n As Long, ByVal cols As Long) As Variant
    Dim outArr() As Variant: ReDim outArr(1 To n, 1 To cols)
    Dim r As Long, c As Long
    For r = 1 To n
        For c = 1 To cols
            outArr(r, c) = src(r, c)
        Next c
    Next r
    CompactRows = outArr
End Function
