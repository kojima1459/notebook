Attribute VB_Name = "modShelf"
Option Explicit

' ========================================
' modShelf - 本棚中核(ファイル取込・チャンク付番・重複排除・削除・一覧)
'
' 抽出→チャンク化→chunk_id付番(ハッシュ重複排除)→my_knowledge追記→
' manifest upsert→埋め込み(MASTER_SPEC §7.2)。再入guard(E0503)=mIngesting、
' 出口はFinish一本化。同名source置換・シート書込は配列一括(§12)・manifestはself(§4)。
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

Private Const CHUNK_TARGET_CHARS As Long = 700
Private Const CHUNK_OVERLAP_CHARS As Long = 150

Private mIngesting As Boolean   ' 再入防止(E0503)
' 強制停止でガードが焼き付くため、時刻を記録し期限超過分は自動解除。
Private mIngestingSince As Date
Private Const GUARD_EXPIRY_MIN As Long = 30

'
' AddFilesViaDialog - 複数選択FileDialog(フィルタ=SupportedExts)→各IngestFile
'
Public Sub AddFilesViaDialog()
    Dim fd As Object
    Set fd = Application.FileDialog(3)   ' msoFileDialogFilePicker(名前付き定数は使わない)
    fd.AllowMultiSelect = True
    fd.Title = "本棚に追加する資料を選んでください"
    fd.Filters.Clear
    fd.Filters.Add "対応ファイル", BuildFilterPattern()

    If fd.Show <> -1 Then Exit Sub   ' キャンセル

    Dim capMax As Long: capMax = modConfig.GetLong("shelf_max_chunks", 10000)
    If capMax < 1 Then capMax = 10000

    Dim okCount As Long: okCount = 0
    Dim ngCount As Long: ngCount = 0
    Dim cappedN As Long: cappedN = 0
    Dim i As Long
    For i = 1 To fd.SelectedItems.count
        If TotalChunks() >= capMax Then
            cappedN = cappedN + 1
        Else
            Dim st As String: st = "failed"
            On Error Resume Next
            st = IngestFile(CStr(fd.SelectedItems(i)), "self")
            On Error GoTo 0
            If st = "done" Or st = "partial" Then
                okCount = okCount + 1
            Else
                ngCount = ngCount + 1
            End If
        End If
    Next i

    Dim msg As String
    msg = okCount & "件を本棚に追加しました。"
    If ngCount > 0 Then
        msg = msg & vbLf & ngCount & "件は取込めませんでした。マイ本棚の一覧で状態をご確認ください。"
    End If
    If cappedN > 0 Then
        msg = msg & vbLf & vbLf & _
            "※本棚の上限に達したため" & cappedN & "件は見送りました。" & vbLf & _
            "configシートの shelf_max_chunks を大きくすると上限を増やせます。"
        On Error Resume Next
        modLog.LogError "E0501", "modShelf.AddFilesViaDialog", _
            "上限" & capMax & "到達で" & cappedN & "件見送り"
        On Error GoTo 0
    End If
    MsgBox msg, vbInformation, modAppDef.APP_NAME
End Sub

'
' IngestFile - MASTER_SPEC §7.2 の手順どおりに1ファイルを取込む。
'   戻り値=manifest status("done"/"partial"/"failed"/"image_pdf")
'
Public Function IngestFile(ByVal path As String, ByVal origin As String) As String
    Dim resultStatus As String: resultStatus = "failed"
    Dim isSelf As Boolean: isSelf = (StrComp(origin, "self", vbTextCompare) = 0)

    ' 1) 再入guard(E0503)。焼き付きガードは自動解除。
    If mIngesting Then
        If DateDiff("n", mIngestingSince, Now) >= GUARD_EXPIRY_MIN Then
            On Error Resume Next
            modLog.LogUsage "guard_recover", "ingest", _
                "取込ガード残留を自動解除(" & GUARD_EXPIRY_MIN & "分超)"
            On Error GoTo 0
            mIngesting = False
        End If
    End If
    If mIngesting Then
        modLog.ShowError "E0503", "modShelf.IngestFile", "path=" & modUtil.SafeLeft(path, 300)
        IngestFile = "failed"
        Exit Function
    End If
    mIngesting = True
    mIngestingSince = Now

    Dim uiStep As String
    On Error GoTo Failed

    uiStep = "ファイル名の解決"
    Dim sourceName As String: sourceName = modUtil.FileNameOf(path)

    ' 1.5) 同名衝突検査(E0504): 別パス同名は明示停止(無言破壊防止)。
    uiStep = "同名衝突の検査"
    If isSelf Then
        Dim conflictPath As String
        conflictPath = modShelfStore.FindConflictingManifestPath(sourceName, path)
        If LenB(conflictPath) > 0 Then
            modLog.ShowError "E0504", "modShelf.IngestFile", _
                "source=" & sourceName & " newPath=" & path & " existingPath=" & conflictPath
            resultStatus = "failed"
            GoTo Finish
        End If
    End If

    ' 2) 上限検査(E0501)
    uiStep = "本棚上限の確認"
    Dim maxChunks As Long: maxChunks = modConfig.GetLong("shelf_max_chunks", 10000)
    If maxChunks < 1 Then maxChunks = 10000
    If TotalChunks() >= maxChunks Then
        modLog.ShowError "E0501", "modShelf.IngestFile", "source=" & sourceName
        resultStatus = "failed"
        GoTo Finish
    End If

    ' 3) 既存同名sourceの置換は「抽出に成功してから」消す(手順は下の
    '    7)の直前)。2026-07-28(レビュー H-6): ここで先に消していたため、
    '    そのPDFを本人が開いていてWordが起動できない・一時的なネットワーク断
    '    などで抽出に失敗すると、旧データは戻らずその資料の検索が即座に
    '    全滅していた。置き換えに失敗したら、前のままである方が正しい。

    ' 4) 抽出
    uiStep = "ファイルからの本文抽出"
    Dim pages() As ExtractedPage
    Dim errCode As String, errDetail As String
    Dim extractOk As Boolean
    extractOk = modExtractor.ExtractFile(path, pages, errCode, errDetail)

    Dim pagesTruncated As Boolean: pagesTruncated = (extractOk And errCode = "PARTIAL_PAGES")

    If Not extractOk Then
        Dim visionEligible As Boolean
        visionEligible = (errCode = "E0303")
        If (Not visionEligible) And errCode = "E0301" Then
            visionEligible = IsImageExtension(path)
        End If
        If visionEligible And modFeatures.FeatureEnabled("vision") Then
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
            Dim failMsg As String
            failMsg = modLog.FriendlyMessage(errCode) & "(コード: " & errCode & ")"
            If errCode = "E0301" And IsImageExtension(path) Then
                failMsg = "画像の取込には画像解析機能の有効化が必要です。" & _
                    "configシートの feature_vision を TRUE にしてから、もう一度お試しください。(コード: E0301)"
            End If
            If isSelf Then
                modShelfStore.UpsertManifestRow path, sourceName, SafeFileDateTime(path), SafeFileLen(path), 0, _
                    failStatus, failMsg, origin
            End If
            resultStatus = failStatus
            GoTo Finish
        End If
    End If

    ' 5) チャンク分割(§B)
    uiStep = "本文のチャンク分割"
    Dim chunks() As ShelfChunk
    Dim chunkN As Long
    chunkN = modChunker.ChunkPagesEx(pages, _
        modConfig.GetLong("chunk_target_chars", CHUNK_TARGET_CHARS), _
        modConfig.GetLong("chunk_overlap_chars", CHUNK_OVERLAP_CHARS), _
        modConfig.GetLong("chunk_max_chars", 1800), _
        modConfig.GetString("chunk_mode", "structure"), chunks)

    If chunkN = 0 Then
        modLog.LogError "E0401", "modShelf.IngestFile", "source=" & sourceName
        If isSelf Then
            modShelfStore.UpsertManifestRow path, sourceName, SafeFileDateTime(path), SafeFileLen(path), 0, _
                "failed", modLog.FriendlyMessage("E0401") & "(コード: E0401)", origin
        End If
        resultStatus = "failed"
        GoTo Finish
    End If

    ' 6) chunk_id付番(bs::hash::pN::cN)+ハッシュ重複スキップ
    uiStep = "チャンクの採番と重複排除"
    Dim wsK As Worksheet: Set wsK = modShelfStore.EnsureKnowledgeSheet()
    ' 置き換え対象(同名source)の行はまだ消していないので、ハッシュ集合から
    ' 除外する。除外しないと入れ直すチャンクが全部「自分自身との重複」と
    ' 判定されて1件も入らない(レビュー H-6)。
    Dim existingHashes As Object
    Set existingHashes = modShelfStore.BuildExistingHashSet(wsK, sourceName)

    Dim outRows() As Variant: ReDim outRows(1 To chunkN, 1 To 9)
    Dim acceptedCount As Long: acceptedCount = 0
    Dim lastPage As Long: lastPage = -1
    Dim seq As Long: seq = 0
    Dim addedStamp As String: addedStamp = modUtil.NowStamp()

    Dim crumbOn As Boolean
    crumbOn = modConfig.GetBool("embed_prefix_breadcrumb", False)

    Dim ci As Long
    For ci = 0 To chunkN - 1
        If chunks(ci).page <> lastPage Then
            lastPage = chunks(ci).page
            seq = 0
        End If
        seq = seq + 1

        Dim bodyText As String
        bodyText = ApplyCrumb(chunks(ci).full_text, sourceName, crumbOn)

        Dim hashHex As String
        hashHex = modUtil.Fnv1a64Hex(modUtil.NormalizeForHash(bodyText))

        If Not existingHashes.Exists(hashHex) Then
            existingHashes.Add hashHex, True
            acceptedCount = acceptedCount + 1
            outRows(acceptedCount, COL_ID) = "bs::" & hashHex & "::p" & chunks(ci).page & "::c" & seq
            outRows(acceptedCount, COL_SOURCE) = sourceName
            outRows(acceptedCount, COL_ORIGIN) = origin
            outRows(acceptedCount, COL_PAGE) = chunks(ci).page
            outRows(acceptedCount, COL_SUMMARY) = ""     ' modEnrichが後で埋める
            outRows(acceptedCount, COL_KEYWORDS) = ""
            outRows(acceptedCount, COL_FULLTEXT) = modUtil.SafeLeft(bodyText, 32000)
            outRows(acceptedCount, COL_ADDED) = addedStamp
            outRows(acceptedCount, COL_EMBEDDED) = 0
        End If
    Next ci

    ' 6.5) ここまで来て初めて旧データを消す(レビュー H-6)。
    '      抽出もチャンク分割も採番も終わっていて、あとは書くだけ。
    '      消してから書くまでの間に失敗する余地がほぼ無い位置。
    uiStep = "既存同名資料の置き換え"
    modShelfStore.RemoveKnowledgeAndVectorsForSource sourceName

    ' 7) my_knowledge追記(err#7対策で200行バッチ書込み)
    uiStep = "本棚への保存(my_knowledge書込み)"
    Dim firstNewRow As Long: firstNewRow = 0
    Dim lastNewRow As Long: lastNewRow = 0
    If acceptedCount > 0 Then
        Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
        If lastK < 1 Then lastK = 1
        firstNewRow = lastK + 1
        If firstNewRow < 2 Then firstNewRow = 2
        lastNewRow = firstNewRow + acceptedCount - 1

        Const WRITE_BATCH_ROWS As Long = 200
        Dim batchStart As Long
        For batchStart = 1 To acceptedCount Step WRITE_BATCH_ROWS
            Dim batchN As Long
            batchN = acceptedCount - batchStart + 1
            If batchN > WRITE_BATCH_ROWS Then batchN = WRITE_BATCH_ROWS

            uiStep = "本棚への保存(" & batchStart & "/" & acceptedCount & ")"
            Dim batchArr As Variant
            batchArr = modShelfStore.SliceRows(outRows, batchStart, batchN)

            Dim wTop As Long: wTop = firstNewRow + batchStart - 1
            wsK.Range(wsK.Cells(wTop, 1), wsK.Cells(wTop + batchN - 1, 9)).Value = batchArr
        Next batchStart
    End If

    ' 8) manifest upsert(status=pending)
    uiStep = "資料台帳の更新(pending)"
    If isSelf Then
        modShelfStore.UpsertManifestRow path, sourceName, SafeFileDateTime(path), SafeFileLen(path), acceptedCount, _
            "pending", "", origin
    End If

    ' 9) EmbedPending(本棚全体の未埋め込み分をまとめて処理)
    uiStep = "ベクトル化(埋め込み)"
    modEmbed.EmbedPending

    ' 10) status確定(embedded=0が残ればpartial)
    uiStep = "取込状態の確定"
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
        modShelfStore.UpsertManifestRow path, sourceName, SafeFileDateTime(path), SafeFileLen(path), acceptedCount, _
            resultStatus, "", origin
    End If

    If isSelf And (resultStatus = "done" Or resultStatus = "partial") Then
        On Error Resume Next
        modStats.Bump "ingest_files_total"
        modStats.AddExp "register"   ' 登録EXP(自己取込のみ=フォルダ自動同期の大量取込では加算されない)
        On Error GoTo 0
        On Error Resume Next
        modLog.LogUsage "ingest", origin, "source=" & sourceName & " chunks=" & acceptedCount & _
            " status=" & resultStatus
        On Error GoTo 0
    End If

    GoTo Finish

Failed:
    Dim failNum As Long: failNum = Err.Number
    Dim failDesc As String: failDesc = Err.Description
    On Error Resume Next
    modLog.LogError "E0801", "modShelf.IngestFile", _
        "[" & uiStep & "] path=" & modUtil.SafeLeft(path, 200) & " err#" & failNum & ": " & failDesc
    If isSelf Then
        modShelfStore.UpsertManifestRow path, sourceName, SafeFileDateTime(path), SafeFileLen(path), 0, _
            "failed", "取込に失敗[" & uiStep & "](#" & failNum & ")", origin
    End If
    On Error GoTo 0
    resultStatus = "failed"

Finish:
    On Error Resume Next
    modUIShelf.RenderShelf
    On Error GoTo 0

    mIngesting = False
    IngestFile = resultStatus
End Function

'
' DeleteSource - knowledge/vectors/manifest から一括削除+再描画
'
Public Sub DeleteSource(ByVal sourceName As String)
    modShelfStore.RemoveKnowledgeAndVectorsForSource sourceName
    modShelfStore.RemoveManifestRowForSource sourceName

    On Error Resume Next
    modUIShelf.RenderShelf
    On Error GoTo 0
End Sub

'
' SourceList - stats(i)="status|ingested_at|chunk_count|error_note|origin"
'
Public Function SourceList(ByRef names() As String, ByRef stats() As String) As Long
    Dim uiStep As String
    On Error GoTo Fail

    Dim outN As Long: outN = 0
    Dim cap As Long: cap = 16
    Dim tmpNames() As String: ReDim tmpNames(0 To cap - 1)
    Dim tmpStats() As String: ReDim tmpStats(0 To cap - 1)

    uiStep = "manifestシートの取得"
    Dim wsM As Worksheet: Set wsM = GetSheet(modAppDef.SH_MANIFEST)
    If Not wsM Is Nothing Then
        uiStep = "manifest最終行の特定"
        Dim lastM As Long: lastM = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
        If lastM >= 2 Then
            uiStep = "manifest範囲の一括読込(" & (lastM - 1) & "行)"
            Dim arr As Variant: arr = wsM.Range(wsM.Cells(2, 1), wsM.Cells(lastM, 9)).Value

            Dim i As Long, lo As Long: lo = LBound(arr, 1)
            For i = lo To UBound(arr, 1)
                uiStep = "行の変換"
                AppendSourceEntry tmpNames, tmpStats, outN, cap, CStr(arr(i, 2)), _
                    CStr(arr(i, 6)) & "|" & CStr(arr(i, 8)) & "|" & _
                    CStr(arr(i, 5)) & "|" & CStr(arr(i, 7)) & "|" & CStr(arr(i, 9))
            Next i
        End If
    End If

    uiStep = "パック由来資料の集計"
    Dim wsK As Worksheet: Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If Not wsK Is Nothing Then
        Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
        If lastK >= 2 Then
            Dim kArr As Variant: kArr = wsK.Range(wsK.Cells(2, 2), wsK.Cells(lastK, 8)).Value
            ' kArrの列: 1=source, 2=origin, 7=added_at(2列目起点のため)
            Dim packCount As Object: Set packCount = CreateObject("Scripting.Dictionary")
            Dim packOrigin As Object: Set packOrigin = CreateObject("Scripting.Dictionary")
            Dim packAdded As Object: Set packAdded = CreateObject("Scripting.Dictionary")

            Dim r As Long
            For r = LBound(kArr, 1) To UBound(kArr, 1)
                Dim org As String: org = CStr(kArr(r, 2))
                If LCase$(Left$(org, 5)) = "pack:" Then
                    Dim src As String: src = CStr(kArr(r, 1))
                    If LenB(src) > 0 Then
                        If packCount.Exists(src) Then
                            packCount(src) = packCount(src) + 1
                        Else
                            packCount.Add src, 1
                            packOrigin.Add src, org
                            packAdded.Add src, CStr(kArr(r, 7))
                        End If
                    End If
                End If
            Next r

            uiStep = "パック由来資料のカード化"
            Dim k As Variant
            For Each k In packCount.Keys
                AppendSourceEntry tmpNames, tmpStats, outN, cap, CStr(k), _
                    "done|" & packAdded(k) & "|" & packCount(k) & "||" & packOrigin(k)
            Next k
        End If
    End If

    uiStep = "一覧の出力"
    If outN = 0 Then
        ReDim names(0 To 0)
        ReDim stats(0 To 0)
        SourceList = 0
        Exit Function
    End If
    ReDim names(0 To outN - 1)
    ReDim stats(0 To outN - 1)
    For i = 0 To outN - 1
        names(i) = tmpNames(i)
        stats(i) = tmpStats(i)
    Next i
    SourceList = outN
    Exit Function

Fail:
    Dim origNum As Long, origDesc As String
    origNum = Err.Number
    origDesc = Err.Description
    On Error Resume Next
    modLog.LogError "E0801", "modShelf.SourceList", "[" & uiStep & "] err#" & origNum & ": " & origDesc
    On Error GoTo 0
    ReDim names(0 To 0)
    ReDim stats(0 To 0)
    SourceList = 0
End Function

'
' TotalChunks - my_knowledgeの総行数(ヘッダ除く)
'
Public Function TotalChunks() As Long
    Dim wsK As Worksheet: Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If wsK Is Nothing Then Exit Function
    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then Exit Function
    TotalChunks = lastK - 1
End Function

' ---- 内部ヘルパー ----

' breadcrumbの〔資料〕を実名置換(OFF=行除去/無し=そのまま)。
Private Function ApplyCrumb(ByVal s As String, ByVal sourceName As String, ByVal enabled As Boolean) As String
    If Left$(s, Len("【〔資料〕")) <> "【〔資料〕" Then
        ApplyCrumb = s
        Exit Function
    End If
    If enabled Then
        ApplyCrumb = "【" & sourceName & Mid$(s, Len("【〔資料〕") + 1)
    Else
        Dim lfPos As Long: lfPos = InStr(s, vbLf)
        If lfPos > 0 Then
            ApplyCrumb = Mid$(s, lfPos + 1)
        Else
            ApplyCrumb = s
        End If
    End If
End Function

' 画像拡張子か(vision委譲判定・D13)。optVisionのIMAGE_EXTSと揃える。
Private Function IsImageExtension(ByVal path As String) As Boolean
    Dim e As String
    e = LCase$(modUtil.ExtOf(path))
    IsImageExtension = (e = "png" Or e = "jpg" Or e = "jpeg")
End Function

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function


' names/statsへ1件追記(容量不足時は倍々拡張)。
Private Sub AppendSourceEntry(ByRef tmpNames() As String, ByRef tmpStats() As String, _
                              ByRef outN As Long, ByRef cap As Long, _
                              ByVal nameVal As String, ByVal statVal As String)
    If outN >= cap Then
        cap = cap * 2
        ReDim Preserve tmpNames(0 To cap - 1)
        ReDim Preserve tmpStats(0 To cap - 1)
    End If
    tmpNames(outN) = nameVal
    tmpStats(outN) = statVal
    outN = outN + 1
End Sub


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
