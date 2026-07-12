Attribute VB_Name = "modShelf"
Option Explicit

' ============================================================================
' modShelf - 本棚中核(ファイル取込・チャンク付番・重複排除・削除・一覧)
' ----------------------------------------------------------------------------
' 役割:
'   ダイアログ選択されたファイルを1件ずつ IngestFile へ渡し、抽出→チャンク化
'   →chunk_id付番(ハッシュ重複排除)→my_knowledge追記→manifest upsert→
'   埋め込み、までの一連の流れをMASTER_SPEC §7.2の手順どおりに実行する。
'
' 設計判断:
'   ・再入guard(E0503)は「別のIngestFile呼び出しが処理中に、DoEvents経由の
'     UI応答でもう一度ボタンが押される」ケースを防ぐためのモジュール変数
'     mIngesting。IngestFile内部は単一の出口(Finishラベル)へ必ず合流させ、
'     ガード解除・RenderShelf呼び出しを1箇所にまとめる(多重解除漏れの防止)。
'   ・「既存同名sourceは置換」はsource列(ファイル名)の完全一致(大文字小文字
'     を区別しない。Windowsファイルシステムの慣習に合わせる)で判定する。
'   ・chunk_idのハッシュ重複排除は my_knowledge 全体(他ファイル・他origin
'     含む)を対象にする。Scripting.Dictionaryで既存ハッシュ集合を1回だけ
'     構築し、O(1)判定する(5000件規模でも軽量)。
'   ・seq(chunk_idの::cN部分)はページが変わるたびに1へリセットする
'     「そのページ内での通し番号」として付番する(重複スキップされた
'     チャンクの分も含めて数える。元のチャンク化結果内での位置を表すため)。
'   ・my_knowledgeへの新規追記は「配列に集めてから1回のRange書込み」で行う
'     (MASTER_SPEC §12: ループ内Range直接アクセス禁止)。既存同名sourceの
'     削除やDeleteSourceでの一括削除も同様に「全体を配列で読み→フィルタ→
'     生存行だけを書き戻し→余った末尾行をクリア」というパターンに統一した。
'   ・manifestへの記録は origin="self" のときのみ行う(MASTER_SPEC §4:
'     「パック由来はmanifestに載せない(my_knowledge.originで管理)」)。
'   ・FileDialog / EnableCancelKey等のOffice/Excel名前付き定数は使わず
'     リテラル値+コメントで表す(modLog.bas の xlSheetHidden 回避と同じ、
'     LibreOffice実行互換のためのV2以来の慣習)。
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

Private Const CHUNK_TARGET_CHARS As Long = 700
Private Const CHUNK_OVERLAP_CHARS As Long = 150

Private mIngesting As Boolean   ' 再入防止(E0503)

' ----------------------------------------------------------------------------
' AddFilesViaDialog - 複数選択FileDialog(フィルタ=SupportedExts)→各IngestFile
' ----------------------------------------------------------------------------
Public Sub AddFilesViaDialog()
    Dim fd As Object
    Set fd = Application.FileDialog(3)   ' msoFileDialogFilePicker(名前付き定数は使わない)
    fd.AllowMultiSelect = True
    fd.Title = "本棚に追加する資料を選んでください"
    fd.Filters.Clear
    fd.Filters.Add "対応ファイル", BuildFilterPattern()

    If fd.Show <> -1 Then Exit Sub   ' キャンセル

    Dim okCount As Long: okCount = 0
    Dim ngCount As Long: ngCount = 0
    Dim i As Long
    For i = 1 To fd.SelectedItems.count
        Dim st As String
        st = IngestFile(CStr(fd.SelectedItems(i)), "self")
        If st = "done" Or st = "partial" Then
            okCount = okCount + 1
        Else
            ngCount = ngCount + 1
        End If
    Next i

    Dim msg As String
    msg = okCount & "件を本棚に追加しました。"
    If ngCount > 0 Then
        msg = msg & vbLf & ngCount & "件は取込めませんでした。マイ本棚の一覧で状態をご確認ください。"
    End If
    MsgBox msg, vbInformation, modAppDef.APP_NAME
End Sub

' ----------------------------------------------------------------------------
' IngestFile - MASTER_SPEC §7.2 の手順どおりに1ファイルを取込む。
'   戻り値=manifest status("done"/"partial"/"failed"/"image_pdf")
' ----------------------------------------------------------------------------
Public Function IngestFile(ByVal path As String, ByVal origin As String) As String
    Dim resultStatus As String: resultStatus = "failed"
    Dim isSelf As Boolean: isSelf = (StrComp(origin, "self", vbTextCompare) = 0)

    ' 1) 再入guard(E0503)
    If mIngesting Then
        modLog.ShowError "E0503", "modShelf.IngestFile", "path=" & modUtil.SafeLeft(path, 300)
        IngestFile = "failed"
        Exit Function
    End If
    mIngesting = True

    Dim sourceName As String: sourceName = modUtil.FileNameOf(path)

    ' 1.5) 同名衝突検査(E0504・Wave4追加): 「既存同名sourceは置換」(§7.2)は
    '   同じ資料を上書き更新する運用を前提にしている。source識別が
    '   ファイル名のみのため、別フォルダにある同名だが別内容のファイルを
    '   取込むと、名前が一致するだけで既存資料のチャンク/ベクトルが無言で
    '   全削除されてしまう(Wave4レビュー指摘: 顧客別フォルダ運用で実際に
    '   起こりうる silent data loss)。manifest(self由来のみ・パスで一意)に
    '   「同じファイル名だが別のfile_path」の行が既にあれば、削除せず明示
    '   エラーで止める(誠実な失敗の方が黙った破壊より安全・R5)。
    '   同一パスの再取込(置換)はこの検査を通過する(既存契約どおり)。
    If isSelf Then
        Dim conflictPath As String
        conflictPath = FindConflictingManifestPath(sourceName, path)
        If LenB(conflictPath) > 0 Then
            modLog.ShowError "E0504", "modShelf.IngestFile", _
                "source=" & sourceName & " newPath=" & path & " existingPath=" & conflictPath
            resultStatus = "failed"
            GoTo Finish
        End If
    End If

    ' 2) 上限検査(E0501)
    Dim maxChunks As Long: maxChunks = modConfig.GetLong("shelf_max_chunks", 5000)
    If maxChunks < 1 Then maxChunks = 5000
    If TotalChunks() >= maxChunks Then
        modLog.ShowError "E0501", "modShelf.IngestFile", "source=" & sourceName
        resultStatus = "failed"
        GoTo Finish
    End If

    ' 3) 既存同名sourceは置換(古いchunk/vector削除。manifestはこの後で上書きするので触らない)
    RemoveKnowledgeAndVectorsForSource sourceName

    ' 4) 抽出
    Dim pages() As ExtractedPage
    Dim errCode As String, errDetail As String
    Dim extractOk As Boolean
    extractOk = modExtractor.ExtractFile(path, pages, errCode, errDetail)

    Dim pagesTruncated As Boolean: pagesTruncated = (extractOk And errCode = "PARTIAL_PAGES")

    If Not extractOk Then
        ' Wave4修正(§7.7): E0303(画像PDF)のとき、feature_visionが有効なら
        ' modFeatures.InvokeFeature("vision","ExtractImagePdfText",path)経由の
        ' 委譲を試みる。以前はこの分岐が無く、optVisionが常にコアから未結線の
        ' 死コードになっていた(フラグをONにしても画像PDFは常にimage_pdf
        ' 確定していた)。成功時はOCR結果を1ページの通常抽出結果として
        ' 後続(チャンク分割以降)へ合流させる。
        If errCode = "E0303" And modFeatures.FeatureEnabled("vision") Then
            Dim visionResult As Variant
            visionResult = modFeatures.InvokeFeature("vision", "ExtractImagePdfText", path)
            Dim visionText As String: visionText = ""
            If VarType(visionResult) = vbString Then
                If Left$(CStr(visionResult), 5) <> "#ERR:" Then visionText = CStr(visionResult)
            End If
            If LenB(Trim$(visionText)) > 0 Then
                ReDim pages(0 To 0)
                pages(0).page = 1
                pages(0).Text = visionText
                extractOk = True
            End If
        End If

        If Not extractOk Then
            Dim failStatus As String
            If errCode = "E0303" Then
                failStatus = "image_pdf"
            Else
                failStatus = "failed"
            End If
            If isSelf Then
                UpsertManifestRow path, sourceName, SafeFileDateTime(path), SafeFileLen(path), 0, _
                    failStatus, modLog.FriendlyMessage(errCode) & "(コード: " & errCode & ")", origin
            End If
            resultStatus = failStatus
            GoTo Finish
        End If
    End If

    ' 5) チャンク分割
    Dim chunks() As ShelfChunk
    Dim chunkN As Long
    chunkN = modChunker.ChunkPages(pages, CHUNK_TARGET_CHARS, CHUNK_OVERLAP_CHARS, chunks)

    If chunkN = 0 Then
        modLog.LogError "E0401", "modShelf.IngestFile", "source=" & sourceName
        If isSelf Then
            UpsertManifestRow path, sourceName, SafeFileDateTime(path), SafeFileLen(path), 0, _
                "failed", modLog.FriendlyMessage("E0401") & "(コード: E0401)", origin
        End If
        resultStatus = "failed"
        GoTo Finish
    End If

    ' 6) chunk_id付番(bs::hash::pN::cN)+ハッシュ重複スキップ
    Dim wsK As Worksheet: Set wsK = EnsureKnowledgeSheet()
    Dim existingHashes As Object: Set existingHashes = BuildExistingHashSet(wsK)

    Dim outRows() As Variant: ReDim outRows(1 To chunkN, 1 To 9)
    Dim acceptedCount As Long: acceptedCount = 0
    Dim lastPage As Long: lastPage = -1
    Dim seq As Long: seq = 0
    Dim addedStamp As String: addedStamp = modUtil.NowStamp()

    Dim ci As Long
    For ci = 0 To chunkN - 1
        If chunks(ci).page <> lastPage Then
            lastPage = chunks(ci).page
            seq = 0
        End If
        seq = seq + 1

        Dim hashHex As String
        hashHex = modUtil.Fnv1a64Hex(modUtil.NormalizeForHash(chunks(ci).full_text))

        If Not existingHashes.Exists(hashHex) Then
            existingHashes.Add hashHex, True
            acceptedCount = acceptedCount + 1
            outRows(acceptedCount, COL_ID) = "bs::" & hashHex & "::p" & chunks(ci).page & "::c" & seq
            outRows(acceptedCount, COL_SOURCE) = sourceName
            outRows(acceptedCount, COL_ORIGIN) = origin
            outRows(acceptedCount, COL_PAGE) = chunks(ci).page
            outRows(acceptedCount, COL_SUMMARY) = ""     ' modEnrichが後で埋める
            outRows(acceptedCount, COL_KEYWORDS) = ""
            outRows(acceptedCount, COL_FULLTEXT) = modUtil.SafeLeft(chunks(ci).full_text, 32000)
            outRows(acceptedCount, COL_ADDED) = addedStamp
            outRows(acceptedCount, COL_EMBEDDED) = 0
        End If
    Next ci

    ' 7) my_knowledge追記(embedded=0)。配列一括書込み(§12)
    Dim firstNewRow As Long: firstNewRow = 0
    Dim lastNewRow As Long: lastNewRow = 0
    If acceptedCount > 0 Then
        Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
        If lastK < 1 Then lastK = 1
        firstNewRow = lastK + 1
        If firstNewRow < 2 Then firstNewRow = 2
        lastNewRow = firstNewRow + acceptedCount - 1

        Dim writeArr As Variant: writeArr = CompactRows(outRows, acceptedCount)
        wsK.Range(wsK.Cells(firstNewRow, 1), wsK.Cells(lastNewRow, 9)).Value = writeArr
    End If

    ' 8) manifest upsert(status=pending)
    If isSelf Then
        UpsertManifestRow path, sourceName, SafeFileDateTime(path), SafeFileLen(path), acceptedCount, _
            "pending", "", origin
    End If

    ' 9) EmbedPending(この資料分だけでなく、本棚全体の未埋め込み分をまとめて処理)
    modEmbed.EmbedPending

    ' 10) manifest status確定(この資料の行にまだembedded=0が残っていればpartial)
    Dim stillPending As Boolean: stillPending = False
    If acceptedCount > 0 Then
        If acceptedCount = 1 Then
            stillPending = (CStr(wsK.Cells(firstNewRow, COL_EMBEDDED).Value) <> "1")
        Else
            Dim chkArr As Variant
            chkArr = wsK.Range(wsK.Cells(firstNewRow, COL_EMBEDDED), wsK.Cells(lastNewRow, COL_EMBEDDED)).Value
            Dim k As Long
            For k = LBound(chkArr, 1) To UBound(chkArr, 1)
                If CStr(chkArr(k, 1)) <> "1" Then
                    stillPending = True
                    Exit For
                End If
            Next k
        End If
    End If

    If pagesTruncated Or stillPending Then
        resultStatus = "partial"
    Else
        resultStatus = "done"
    End If

    If isSelf Then
        UpsertManifestRow path, sourceName, SafeFileDateTime(path), SafeFileLen(path), acceptedCount, _
            resultStatus, "", origin
    End If

Finish:
    ' 11) カード再描画(modUIShelf未実装/実行時エラーでも取込処理自体は止めない)
    On Error Resume Next
    modUIShelf.RenderShelf
    On Error GoTo 0

    mIngesting = False
    IngestFile = resultStatus
End Function

' ----------------------------------------------------------------------------
' DeleteSource - knowledge/vectors/manifest から一括削除+再描画
' ----------------------------------------------------------------------------
Public Sub DeleteSource(ByVal sourceName As String)
    RemoveKnowledgeAndVectorsForSource sourceName
    RemoveManifestRowForSource sourceName

    On Error Resume Next
    modUIShelf.RenderShelf
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' SourceList - カード描画用。my_manifestの各行を names()/stats() に展開する。
'   names(i)  = 資料名(file_name)
'   stats(i)  = "status|ingested_at|chunk_count|error_note|origin" の
'               パイプ区切り文字列(呼び出し側=modUIShelfがSplit("|")して使う)
'   戻り値=件数(0可)
' ----------------------------------------------------------------------------
Public Function SourceList(ByRef names() As String, ByRef stats() As String) As Long
    Dim wsM As Worksheet: Set wsM = GetSheet(modAppDef.SH_MANIFEST)
    If wsM Is Nothing Then
        ReDim names(0 To -1)
        ReDim stats(0 To -1)
        SourceList = 0
        Exit Function
    End If

    Dim lastM As Long: lastM = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
    If lastM < 2 Then
        ReDim names(0 To -1)
        ReDim stats(0 To -1)
        SourceList = 0
        Exit Function
    End If

    Dim arr As Variant: arr = wsM.Range(wsM.Cells(2, 1), wsM.Cells(lastM, 9)).Value
    Dim n As Long: n = UBound(arr, 1) - LBound(arr, 1) + 1
    ReDim names(0 To n - 1)
    ReDim stats(0 To n - 1)

    Dim i As Long, lo As Long: lo = LBound(arr, 1)
    For i = lo To UBound(arr, 1)
        names(i - lo) = CStr(arr(i, 2))
        stats(i - lo) = CStr(arr(i, 6)) & "|" & CStr(arr(i, 8)) & "|" & _
                        CStr(arr(i, 5)) & "|" & CStr(arr(i, 7)) & "|" & CStr(arr(i, 9))
    Next i
    SourceList = n
End Function

' ----------------------------------------------------------------------------
' TotalChunks - my_knowledgeの総行数(ヘッダ除く)
' ----------------------------------------------------------------------------
Public Function TotalChunks() As Long
    Dim wsK As Worksheet: Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If wsK Is Nothing Then Exit Function
    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then Exit Function
    TotalChunks = lastK - 1
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function EnsureKnowledgeSheet() As Worksheet
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = modAppDef.SH_KNOWLEDGE
        Dim hdr As Variant
        hdr = Array("chunk_id", "source", "origin", "page", "summary", "keywords", "full_text", "added_at", "embedded")
        Dim i As Long
        For i = LBound(hdr) To UBound(hdr)
            ws.Cells(1, i + 1).Value = hdr(i)
        Next i
        On Error Resume Next
        ws.Visible = 2   ' xlSheetVeryHidden
        On Error GoTo 0
    End If
    Set EnsureKnowledgeSheet = ws
    Exit Function
Fail:
    Set EnsureKnowledgeSheet = Nothing
End Function

Private Function EnsureManifestSheet() As Worksheet
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_MANIFEST)
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = modAppDef.SH_MANIFEST
        Dim hdr As Variant
        hdr = Array("file_path", "file_name", "modified_at", "size", "chunk_count", "status", "error_note", "ingested_at", "origin")
        Dim i As Long
        For i = LBound(hdr) To UBound(hdr)
            ws.Cells(1, i + 1).Value = hdr(i)
        Next i
        On Error Resume Next
        ws.Visible = 0   ' xlSheetHidden
        On Error GoTo 0
    End If
    Set EnsureManifestSheet = ws
    Exit Function
Fail:
    Set EnsureManifestSheet = Nothing
End Function

' my_knowledge の chunk_id 列全体から「bs::HASH::pN::cN」のHASH部分を集めた
' Scripting.Dictionary を返す(§4のハッシュ重複排除キー)。
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
        AddHashFromId dict, CStr(wsK.Cells(2, 1).Value)
    Else
        Dim ids As Variant: ids = wsK.Range(wsK.Cells(2, 1), wsK.Cells(lastK, 1)).Value
        Dim i As Long
        For i = LBound(ids, 1) To UBound(ids, 1)
            AddHashFromId dict, CStr(ids(i, 1))
        Next i
    End If
    Set BuildExistingHashSet = dict
End Function

Private Sub AddHashFromId(ByVal dict As Object, ByVal chunkId As String)
    If LenB(chunkId) = 0 Then Exit Sub
    Dim parts() As String: parts = Split(chunkId, "::")
    If UBound(parts) - LBound(parts) + 1 < 2 Then Exit Sub
    Dim h As String: h = parts(LBound(parts) + 1)
    If LenB(h) > 0 Then
        If Not dict.Exists(h) Then dict.Add h, True
    End If
End Sub

' 指定sourceのknowledge行・対応するvector行を全て取り除く(見つからなければ何もしない)。
' 「全体を配列で読む→フィルタ→生存行を書き戻す→余った末尾行をクリア」で
' ループ内Range直接アクセスを避ける(§12)。
Private Sub RemoveKnowledgeAndVectorsForSource(ByVal sourceName As String)
    Dim wsK As Worksheet: Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If wsK Is Nothing Then Exit Sub
    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then Exit Sub

    Dim arr As Variant: arr = wsK.Range(wsK.Cells(2, 1), wsK.Cells(lastK, 9)).Value
    Dim nRows As Long: nRows = UBound(arr, 1) - LBound(arr, 1) + 1

    Dim removedIds As Object: Set removedIds = CreateObject("Scripting.Dictionary")
    Dim survivors() As Variant: ReDim survivors(1 To nRows, 1 To 9)
    Dim survivorCount As Long: survivorCount = 0

    Dim i As Long, c As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(CStr(arr(i, COL_SOURCE)), sourceName, vbTextCompare) = 0 Then
            Dim cid As String: cid = CStr(arr(i, COL_ID))
            If LenB(cid) > 0 Then
                If Not removedIds.Exists(cid) Then removedIds.Add cid, True
            End If
        Else
            survivorCount = survivorCount + 1
            For c = 1 To 9
                survivors(survivorCount, c) = arr(i, c)
            Next c
        End If
    Next i

    If removedIds.count = 0 Then Exit Sub   ' 一致無し(何もしなくてよい)

    If survivorCount > 0 Then
        Dim writeArr As Variant: writeArr = CompactRows(survivors, survivorCount)
        wsK.Range(wsK.Cells(2, 1), wsK.Cells(1 + survivorCount, 9)).Value = writeArr
    End If
    If survivorCount < nRows Then
        wsK.Range(wsK.Cells(2 + survivorCount, 1), wsK.Cells(1 + nRows, 9)).ClearContents
    End If

    RemoveVectorsByIds removedIds
End Sub

Private Sub RemoveVectorsByIds(ByVal removedIds As Object)
    Dim wsV As Worksheet: Set wsV = GetSheet(modAppDef.SH_VECTORS)
    If wsV Is Nothing Then Exit Sub
    Dim lastV As Long: lastV = wsV.Cells(wsV.Rows.count, 1).End(xlUp).row
    If lastV < 2 Then Exit Sub

    Dim arr As Variant: arr = wsV.Range(wsV.Cells(2, 1), wsV.Cells(lastV, 2)).Value
    Dim nRows As Long: nRows = UBound(arr, 1) - LBound(arr, 1) + 1

    Dim survivors() As Variant: ReDim survivors(1 To nRows, 1 To 2)
    Dim survivorCount As Long: survivorCount = 0
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If Not removedIds.Exists(CStr(arr(i, 1))) Then
            survivorCount = survivorCount + 1
            survivors(survivorCount, 1) = arr(i, 1)
            survivors(survivorCount, 2) = arr(i, 2)
        End If
    Next i

    If survivorCount = nRows Then Exit Sub   ' 一致無し

    If survivorCount > 0 Then
        Dim writeArr() As Variant: ReDim writeArr(1 To survivorCount, 1 To 2)
        Dim r As Long, c As Long
        For r = 1 To survivorCount
            For c = 1 To 2
                writeArr(r, c) = survivors(r, c)
            Next c
        Next r
        wsV.Range(wsV.Cells(2, 1), wsV.Cells(1 + survivorCount, 2)).Value = writeArr
    End If
    wsV.Range(wsV.Cells(2 + survivorCount, 1), wsV.Cells(1 + nRows, 2)).ClearContents
End Sub

Private Sub RemoveManifestRowForSource(ByVal sourceName As String)
    Dim wsM As Worksheet: Set wsM = GetSheet(modAppDef.SH_MANIFEST)
    If wsM Is Nothing Then Exit Sub
    Dim lastM As Long: lastM = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
    If lastM < 2 Then Exit Sub

    Dim arr As Variant: arr = wsM.Range(wsM.Cells(2, 1), wsM.Cells(lastM, 9)).Value
    Dim nRows As Long: nRows = UBound(arr, 1) - LBound(arr, 1) + 1

    Dim survivors() As Variant: ReDim survivors(1 To nRows, 1 To 9)
    Dim survivorCount As Long: survivorCount = 0
    Dim i As Long, c As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(CStr(arr(i, 2)), sourceName, vbTextCompare) <> 0 Then
            survivorCount = survivorCount + 1
            For c = 1 To 9
                survivors(survivorCount, c) = arr(i, c)
            Next c
        End If
    Next i

    If survivorCount = nRows Then Exit Sub   ' 一致無し

    If survivorCount > 0 Then
        Dim writeArr As Variant: writeArr = CompactRows(survivors, survivorCount)
        wsM.Range(wsM.Cells(2, 1), wsM.Cells(1 + survivorCount, 9)).Value = writeArr
    End If
    wsM.Range(wsM.Cells(2 + survivorCount, 1), wsM.Cells(1 + nRows, 9)).ClearContents
End Sub

Private Sub UpsertManifestRow(ByVal filePath As String, ByVal fileName As String, ByVal modifiedAt As Date, _
                              ByVal sizeBytes As Double, ByVal chunkCount As Long, ByVal status As String, _
                              ByVal errorNote As String, ByVal origin As String)
    Dim wsM As Worksheet: Set wsM = EnsureManifestSheet()
    If wsM Is Nothing Then Exit Sub

    Dim r As Long: r = FindManifestRowByPath(wsM, filePath)
    If r = 0 Then
        Dim lastM As Long: lastM = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
        r = lastM + 1
        If r < 2 Then r = 2
    End If

    wsM.Cells(r, 1).Value = filePath
    wsM.Cells(r, 2).Value = fileName
    wsM.Cells(r, 3).Value = modifiedAt
    wsM.Cells(r, 4).Value = sizeBytes
    wsM.Cells(r, 5).Value = chunkCount
    wsM.Cells(r, 6).Value = status
    wsM.Cells(r, 7).Value = modUtil.SafeLeft(errorNote, 2000)
    wsM.Cells(r, 8).Value = modUtil.NowStamp()
    wsM.Cells(r, 9).Value = origin
End Sub

Private Function FindManifestRowByPath(ByVal wsM As Worksheet, ByVal filePath As String) As Long
    Dim lastM As Long: lastM = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
    If lastM < 2 Then Exit Function
    If lastM = 2 Then
        If StrComp(CStr(wsM.Cells(2, 1).Value), filePath, vbTextCompare) = 0 Then FindManifestRowByPath = 2
        Exit Function
    End If

    Dim arr As Variant: arr = wsM.Range(wsM.Cells(2, 1), wsM.Cells(lastM, 1)).Value
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(CStr(arr(i, 1)), filePath, vbTextCompare) = 0 Then
            FindManifestRowByPath = i + 1
            Exit Function
        End If
    Next i
End Function

' src(1 To n以上, 1 To 9)の先頭n行だけを(1 To n, 1 To 9)へ詰め直す
' (Range書込みは配列の次元とサイズが対象範囲と一致していないといけないため)。
Private Function CompactRows(ByRef src As Variant, ByVal n As Long) As Variant
    Dim outArr() As Variant: ReDim outArr(1 To n, 1 To 9)
    Dim r As Long, c As Long
    For r = 1 To n
        For c = 1 To 9
            outArr(r, c) = src(r, c)
        Next c
    Next r
    CompactRows = outArr
End Function

Private Function SafeFileDateTime(ByVal path As String) As Date
    On Error GoTo Fail
    SafeFileDateTime = FileDateTime(path)
    Exit Function
Fail:
    SafeFileDateTime = Now
End Function

Private Function SafeFileLen(ByVal path As String) As Double
    On Error GoTo Fail
    SafeFileLen = CDbl(FileLen(path))
    Exit Function
Fail:
    SafeFileLen = 0
End Function

Private Function BuildFilterPattern() As String
    Dim exts() As String: exts = Split(modExtractor.SupportedExts(), ",")
    Dim parts() As String: ReDim parts(LBound(exts) To UBound(exts))
    Dim i As Long
    For i = LBound(exts) To UBound(exts)
        parts(i) = "*." & Trim$(exts(i))
    Next i
    BuildFilterPattern = Join(parts, ";")
End Function
