Attribute VB_Name = "modPackExport"
Option Explicit

' ========================================
' modPackExport - パック(pack.xlsx)の【書き出し】側
'
' modPack から切り出した発行系。取込(ImportPackFile)と書き出しが共有するのは
' pack.xlsx のシート名だけで、処理も失敗時の振る舞いも別物のため分ける。
' 発行の不具合を追うときに取込のコードを読まされない。
'
' 切り出しの理由(2026-07-28): modPack が契約上限30,000字に対し残り149字で、
' 発行の origin フィルタ(レビュー H-4)を足す余地が無かった(レビュー I-2)。
'
' silent:=True でもPII走査だけは必ず通す方針は、切り出し後も変えていない。
' ========================================

Private Const COL_ID As Long = 1
Private Const COL_SOURCE As Long = 2
Private Const COL_ORIGIN As Long = 3
Private Const COL_PAGE As Long = 4
Private Const COL_SUMMARY As Long = 5
Private Const COL_KEYWORDS As Long = 6
Private Const COL_FULLTEXT As Long = 7
Private Const COL_ADDED As Long = 8
Private Const COL_EMBEDDED As Long = 9

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

' ExportPackDialog - 対象を選び、保存先を選び、ExportPackToFileへ渡す。
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

    Dim wroteN As Long
    If ExportPackToFile(savePath, sourceFilter, False, wroteN) Then
        MsgBox "パックを書き出しました。" & vbLf & _
            wroteN & "件の資料を「" & modUtil.FileNameOf(savePath) & "」として保存しました。", _
            vbInformation, modAppDef.APP_NAME
    End If
End Sub

' ----------------------------------------------------------------------------
' ExportPackToFile - パスを直接指定してパックを書き出す(部門正典の発行から使う)。
'   ダイアログを一切出さずに完了できるので、発行を1操作にまとめられる。
'   silent:=True でも【PII走査だけは必ず通す】。個人情報入りの資料を無言で
'   全社配布する経路は作らない(検出したら中止してFalseを返す)。
'   戻り値 True=保存成功。outCount に書き出した件数を返す。
'   originFilter: 書き出す行の origin を絞る(レビュー H-4)。部門正典の
'   発行は "self" を渡す。詳細は LoadChunksForExport のコメント参照。
' ----------------------------------------------------------------------------
Public Function ExportPackToFile(ByVal savePath As String, ByVal sourceFilter As String, _
                                 ByVal silent As Boolean, ByRef outCount As Long, _
                                 Optional ByVal originFilter As String = "") As Boolean
    Dim ids() As String, sources() As String, pages() As Long
    Dim summaries() As String, keywords() As String, fullTexts() As String
    Dim n As Long
    n = LoadChunksForExport(sourceFilter, ids, sources, pages, summaries, keywords, fullTexts, originFilter)
    If n = 0 Then
        If Not silent Then
            MsgBox "書き出せる資料がありません。" & vbLf & _
                "マイ本棚に資料を追加してから、もう一度お試しください。", _
                vbExclamation, modAppDef.APP_NAME
        End If
        Exit Function
    End If

    Dim piiCount As Long, piiExamples As String
    ScanChunksForPii sources, fullTexts, n, piiCount, piiExamples
    If piiCount > 0 Then
        modLog.LogError "E0703", "modPackExport.ExportPackToFile", _
            "件数=" & piiCount & " 例=" & piiExamples
        MsgBox modLog.FriendlyMessage("E0703") & vbLf & vbLf & _
            "件数: " & piiCount & "件" & vbLf & "例: " & piiExamples & vbLf & vbLf & _
            "個人情報が含まれる可能性があるため、書き出しを中止しました。" & vbLf & _
            "該当の資料を本棚から外してから、もう一度お試しください。" & vbLf & _
            "(コード: E0703)", vbExclamation, modAppDef.APP_NAME
        Exit Function
    End If

    Dim authorName As String: authorName = Trim$(modConfig.GetString("pack_author", ""))
    If LenB(authorName) = 0 Then
        On Error Resume Next
        authorName = modP2P.CurrentUserId()
        On Error GoTo 0
    End If
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
    modStats.AddExp "pack_share"
    modLog.LogUsage "pack_export", "", "n=" & n & " scope=" & IIf(LenB(sourceFilter) = 0, "all", sourceFilter)

    outCount = n
    ExportPackToFile = True
    Exit Function

SaveFail:
    Dim saveErrDetail As String: saveErrDetail = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume SaveFailCleanup2
SaveFailCleanup2:
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not newWb Is Nothing Then newWb.Close SaveChanges:=False
    On Error GoTo 0
    modLog.LogError "PACK_SAVE_FAILED", "modPackExport.ExportPackToFile", saveErrDetail
    MsgBox "書き出し先にファイルを保存できませんでした。" & vbLf & _
        "読み取り専用のフォルダ、権限不足、パスの打ち間違いが考えられます。" & vbLf & _
        "保存先: " & savePath, vbExclamation, modAppDef.APP_NAME
End Function

' アクティブセルの行の中からmy_manifestのfile_nameと一致する文字列を探す
' (カード行の列レイアウトに依存しない特定方法。上部コメント参照)。
Public Function FindActiveRowSourceName() As String
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

Public Sub AddNameIfNotEmpty(ByVal dict As Object, ByVal nm As String)
    If LenB(nm) = 0 Then Exit Sub
    If Not dict.Exists(nm) Then dict.Add nm, True
End Sub

' 内部ヘルパー: 書き出し(ExportPackDialog)
' originFilter (2026-07-28 レビュー H-4):
'   書き出す行の origin を絞る。空なら全部(従来動作=手渡しパック)。
'   部門正典の発行は "self" を渡し、自分で取り込んだ資料だけを出す。
'   これが無かったため、他部門を購読している端末で発行すると
'   【他部門の正典の全文が自部門のパックとして再配布】されていた。
'   元部門が改定してもコピーは古いまま残るので、改定が効かない資料が
'   社内に増え続ける。1台のPCで発行者と利用者を兼ねてテストする
'   手順(=PoC初日にやること)で確実に踏む。
Public Function LoadChunksForExport(ByVal sourceFilter As String, ByRef ids() As String, ByRef sources() As String, _
        ByRef pages() As Long, ByRef summaries() As String, ByRef keywords() As String, _
        ByRef fullTexts() As String, Optional ByVal originFilter As String = "") As Long
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
        Dim okSrc As Boolean, okOrigin As Boolean
        okSrc = (LenB(sourceFilter) = 0) Or (StrComp(srcName, sourceFilter, vbTextCompare) = 0)
        okOrigin = (LenB(originFilter) = 0) Or _
                   (StrComp(Trim$(CStr(arr(i, COL_ORIGIN))), originFilter, vbTextCompare) = 0)
        If okSrc And okOrigin Then
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

Public Sub ScanChunksForPii(sources() As String, fullTexts() As String, ByVal n As Long, _
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

Public Sub RemoveExtraSheets(ByVal wb As Workbook)
    Application.DisplayAlerts = False
    Do While wb.Worksheets.count > 1
        wb.Worksheets(wb.Worksheets.count).Delete
    Loop
    Application.DisplayAlerts = True
End Sub

' 2026-07-28(レビュー H-5): author_id を追加した。
' 感謝状の宛先は「パックの作者」だが、これまで宛先キーに使っていたのは
' author(=初回起動で本人が打つ表示名)で、受け取る側の照合キーは
' AD の CN または %USERNAME% だった。両者が偶然一致しない限り、
' 感謝EXP も thanks_received_total も称号も永久に付かず、宛先不明の
' thx_ ファイルが共有フォルダに溜まり続ける。
' 「使うと作った人に感謝が届く」というこのプロダクトの中核ループが、
' 構造的に無効だった。表示は従来どおり author、宛先は author_id を使う。
Public Sub WritePackMeta(ByVal ws As Worksheet, ByVal authorName As String, ByVal chunkCount As Long)
    Dim authorId As String
    On Error Resume Next
    authorId = modP2P.CurrentUserId()
    On Error GoTo 0

    Dim rows(1 To 10, 1 To 2) As Variant
    rows(1, 1) = "key": rows(1, 2) = "value"
    rows(2, 1) = "pack_format_version": rows(2, 2) = modAppDef.PACK_FORMAT_VERSION
    rows(3, 1) = "pack_name": rows(3, 2) = "マイ本棚パック_" & Format$(Now, "yyyymmdd")
    rows(4, 1) = "author": rows(4, 2) = authorName
    rows(5, 1) = "created_at": rows(5, 2) = modUtil.NowStamp()
    rows(6, 1) = "description": rows(6, 2) = ""
    rows(7, 1) = "embed_dim": rows(7, 2) = modConfig.GetLong("embed_dim", 1536)
    rows(8, 1) = "app_version": rows(8, 2) = modAppDef.APP_VERSION
    rows(9, 1) = "chunk_count": rows(9, 2) = chunkCount
    rows(10, 1) = "author_id": rows(10, 2) = authorId
    ws.Range(ws.Cells(1, 1), ws.Cells(10, 2)).Value = rows
End Sub

Public Sub WritePackChunks(ByVal ws As Worksheet, ids() As String, sources() As String, pages() As Long, _
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

Public Function ChunkHeaderNames() As Variant
    ChunkHeaderNames = Array("chunk_id", "source", "page", "summary", "keywords", "full_text")
End Function

Public Sub WritePackVectors(ByVal ws As Worksheet, ids() As String, ByVal n As Long)
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

Public Function BuildVectorMap() As Object
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
