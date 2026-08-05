Attribute VB_Name = "modShelf"
Option Explicit

' ========================================
' modShelf - 本棚中核(ファイル取込・チャンク付番・重複排除・削除・一覧)
' 抽出→チャンク化→chunk_id付番(ハッシュ重複排除)→my_knowledge追記→
' manifest upsert→埋め込み(MASTER_SPEC §7.2)。再入guard(E0503)=mIngesting、
' 出口はFinish一本化(中間保存もそこ=R18-2c)。同名source置換は「書いてから
' 消す」(R18-2e。全重複なら消さない=R18H FA-5)。書込は配列一括(§12)。
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
' 強制停止でガードが焼き付くため時刻を記録し期限超過分は自動解除。
Private mIngestingSince As Date
Private Const GUARD_EXPIRY_MIN As Long = 30


' IngestFile - MASTER_SPEC§7.2の手順どおり1ファイルを取込む。
'   戻り値=manifest status("done"/"partial"/"failed"/"image_pdf")
' silent (2026-07-28 レビュー L-11): 一括取込で1件ごとにモーダルが出ると
'   同期が朝まで止まる。True なら記録だけ残す。
' outErrCode (2026-07-30・要件E): AddFilesResult 向け失敗理由内訳(E0504/
'   image_pdf/E0302等)の補助出口。Optional末尾追加で既存呼び出し元
'   (modVault/modUIShelf/modShelfSync等)は無改修。成功時は""。
' interactive (2026-08-04 R15-FixA FA-1・A-H1/B-H1):「利用者がその場にいる
'   取込か」。silent(結果モーダル抑止)と別問いを1本で表していたため
'   R15-7bの事前確認が【全経路で死んでいた】。TrueはAddFilesResultと
'   手動同期だけ(R18-6d)。無人同期は従来どおり無確認。
Public Function IngestFile(ByVal path As String, ByVal origin As String, _
                           Optional ByVal silent As Boolean = False, _
                           Optional ByRef outErrCode As String = "", _
                           Optional ByVal interactive As Boolean = False) As String
    outErrCode = ""   ' 呼び出し側の使い回し変数でも必ずここで初期化する
    Dim resultStatus As String: resultStatus = "failed"
    Dim isSelf As Boolean: isSelf = (StrComp(origin, "self", vbTextCompare) = 0)

    ' 1) 再入guard(E0503)。焼き付きガードは自動解除。R15-1a: 失効は「開始から」
    ' ではなく「最後のビートから30分」(取込途中の二重取込を防止。実機第4報RC5)。
    If mIngesting Then
        If modShelfBatch.GuardExpiredNow(mIngestingSince, GUARD_EXPIRY_MIN) Then
            On Error Resume Next
            modLog.LogUsage "guard_recover", "ingest", _
                "取込ガード残留を自動解除(無音" & GUARD_EXPIRY_MIN & "分超)"
            On Error GoTo 0
            mIngesting = False
        End If
    End If
    If mIngesting Then
        ' 2026-07-31(R7 B-2): ここは ShowError(モーダル)だった。再入のたび
        ' ダイアログが積み上がり実機で「応答なし」に見えた。ログは残し実況行で伝える。
        modLog.LogError "E0503", "modShelf.IngestFile", "path=" & modUtil.SafeLeft(path, 300)
        On Error Resume Next
        modUIMain.SetStage "処理中です。完了までお待ちください…"
        On Error GoTo 0
        outErrCode = "E0503"
        IngestFile = "failed"
        Exit Function
    End If
    mIngesting = True
    mIngestingSince = Now
    ' R15-FixA(FA-6): 中断の印はここでも必ず下ろす(modVault/modUIShelfは直接
    ' 呼ぶため前回の印が残ると幽霊中断になる)。一括取込・同期は判定済みで来る。
    modShelfBatch.ResetCancel

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
            If silent Then
                modLog.LogError "E0504", "modShelf.IngestFile", _
                    "source=" & sourceName & " newPath=" & path & " existingPath=" & conflictPath
            Else
                modLog.ShowError "E0504", "modShelf.IngestFile", _
                    "source=" & sourceName & " newPath=" & path & " existingPath=" & conflictPath
            End If
            resultStatus = "failed"
            outErrCode = "E0504"
            GoTo Finish
        End If
    End If

    ' 2) 上限検査(E0501)
    uiStep = "本棚上限の確認"
    Dim maxChunks As Long: maxChunks = modConfig.GetLong("shelf_max_chunks", modAppDef.DEFAULT_SHELF_MAX_CHUNKS)
    If maxChunks < 1 Then maxChunks = modAppDef.DEFAULT_SHELF_MAX_CHUNKS
    If TotalChunks() >= maxChunks Then
        modLog.ShowError "E0501", "modShelf.IngestFile", "source=" & sourceName
        resultStatus = "failed"
        outErrCode = "E0501"
        GoTo Finish
    End If

    ' 3) 既存同名sourceの置換は「新しい行を書き終えてから」消す(下の7.5)。

    ' 4) 抽出
    uiStep = "ファイルからの本文抽出"
    Dim pages() As ExtractedPage
    Dim errCode As String, errDetail As String
    Dim extractOk As Boolean
    extractOk = modExtractor.ExtractFile(path, pages, errCode, errDetail)

    Dim pagesTruncated As Boolean: pagesTruncated = (extractOk And errCode = "PARTIAL_PAGES")

    If Not extractOk Then
        ' vision フォールバック(画像の文字起こし・画像PDFのOCR)は判断ごと
        ' modShelfVision が引き受ける(2026-07-31 R6)。上限ページ打ち切りは
        ' pagesTruncated が立ちpartial判定へ合流。silentはそのまま渡す
        ' (R11-A C4: 無人同期からGSの案内が出ると同期が止まる)。
        Dim visionNote As String
        If modShelfVision.TryVisionFallback(path, errCode, pages, pagesTruncated, _
                                            visionNote, silent, interactive) Then
            extractOk = True
        Else
            Dim failStatus As String
            If errCode = "E0303" Then
                failStatus = "image_pdf"
            Else
                failStatus = "failed"
            End If
            Dim failMsg As String: failMsg = visionNote
            If LenB(failMsg) = 0 Then
                ' 2026-08-01(同梱-9): 50MBガード(txt/md/csv入力上限)だけは
                ' errDetailの理由をそのまま使う(E0302の一般文言で上書きすると
                ' 「大きすぎる」が出ない)。他は生ログなので不採用。
                If InStr(errDetail, "ファイルが大きすぎます") > 0 Then
                    failMsg = Split(errDetail, " [")(0)
                Else
                    ' 2026-08-03(R13-3b): 文言決定はmodLog.FriendlyFailMsgへ一本化
                    ' (docx/docにGhostscriptの話をしない。第2報RC6)。
                    failMsg = modLog.FriendlyFailMsg(errCode, errDetail, modUtil.ExtOf(path))
                End If
            End If
            If isSelf Then
                ' R18-2a: chunk_count は -1(=前値保持)。ここは実データを1行も
                ' 消していない失敗なので、0で上書きするとカードだけが「0件」に
                ' 化ける(実機第5報⑧)。status/メモの更新は従来どおり。
                modShelfStore.UpsertManifestRow path, sourceName, SafeFileDateTime(path), SafeFileLen(path), -1, _
                    failStatus, failMsg, origin
            End If
            resultStatus = failStatus
            outErrCode = errCode
            ' 2026-08-04(R15-FixB FB-5): 事前確認「いいえ」の資料は【失敗していない】。
            ' statusはimage_pdfのまま理由コードだけ差し替え、一括取込の集計で
            ' 見送りを失敗と分けて数える(先頭一致)。
            If modUtilText.IsDeclineNote(failMsg) Then outErrCode = "declined"
            GoTo Finish
        End If
    End If

    ' 5) チャンク分割(§B)
    uiStep = "本文のチャンク分割"
    Dim chunks() As ShelfChunk
    Dim chunkN As Long
    ' R14-5a(実機第3報 RC7): 拡張子別の粒度設定(PDF/Word)がここまで届いて
    ' いなかった。ChunkParamForが「拡張子別→グローバル→既定」の2段
    ' フォールバックを1箇所で行う(chunk_modeはグローバルのまま)。
    Dim chunkExt As String: chunkExt = modUtil.ExtOf(path)
    chunkN = modChunker.ChunkPagesEx(pages, _
        ChunkParamFor("chunk_target_chars", chunkExt, CHUNK_TARGET_CHARS), _
        ChunkParamFor("chunk_overlap_chars", chunkExt, CHUNK_OVERLAP_CHARS), _
        ChunkParamFor("chunk_max_chars", chunkExt, 1800), _
        modConfig.GetString("chunk_mode", "structure"), chunks)

    ' 2026-07-29: 1ページの失敗でファイル全体を捨てない。ただし黙って減らすと
    ' 「88チャンクが70」に気づけないので必ず1行残す。
    Dim skippedPages As Long: skippedPages = modChunker.SkippedPageCount()
    If skippedPages > 0 Then
        On Error Resume Next
        modLog.LogUsage "chunk_skipped_pages", "", _
            sourceName & ": " & skippedPages & "ページを処理できず飛ばしました"
        ' On Error GoTo 0ではなくFailedへ戻す(0にするとFailedハンドラごと
        ' 解除され以降の実行時エラーが素通りする=レビューH-3と同型)。
        On Error GoTo Failed
    End If

    If chunkN = 0 Then
        modLog.LogError "E0401", "modShelf.IngestFile", "source=" & sourceName
        If isSelf Then
            ' R18-2a: chunk_countは-1(前値保持)。旧データはまだ消していない。
            modShelfStore.UpsertManifestRow path, sourceName, SafeFileDateTime(path), SafeFileLen(path), -1, _
                "failed", modLog.FriendlyMessage("E0401") & "(コード: E0401)", origin
        End If
        resultStatus = "failed"
        outErrCode = "E0401"
        GoTo Finish
    End If

    ' 6) chunk_id付番(bs::hash::pN::cN)+ハッシュ重複スキップ
    uiStep = "チャンクの採番と重複排除"
    Dim wsK As Worksheet: Set wsK = modShelfStore.EnsureKnowledgeSheet()
    ' 置き換え対象(同名source)の行は未削除なのでハッシュ集合から除外する。
    ' 除外しないと入れ直すチャンクが全部「自分自身との重複」になる(H-6)。
    Dim existingHashes As Object
    Set existingHashes = modShelfStore.BuildExistingHashSet(wsK, sourceName)

    ' 10列目norm_text(R12-4): 照合用の正規化済みテキストをここで1回だけ作る。
    Dim outRows() As Variant: ReDim outRows(1 To chunkN, 1 To 10)
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
            ' 保存本文(SafeLeft後)から作る。1セル超は空にし検索側の遅延
            ' バックフィルへ委ねる(切り詰めるとスコアが揺れるため)。
            Dim normDoc As String
            normDoc = modSparse.MatchDocText("", "", sourceName, CStr(outRows(acceptedCount, COL_FULLTEXT)))
            If Len(normDoc) <= 32000 Then outRows(acceptedCount, modShelfStore.COL_K_NORM) = normDoc
        End If
    Next ci

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
            wsK.Range(wsK.Cells(wTop, 1), wsK.Cells(wTop + batchN - 1, modShelfStore.COL_K_NORM)).Value = batchArr
        Next batchStart
    End If

    ' 7.5) 旧データを消すのは【新しい行を全部書き終えてから】(R18-2e)。落ちても
    '      旧データが無傷になる順序=§3-5・調査agent1 §3(a)の真の消失経路対策。
    '      firstNewRow は今書いた行を道連れにしないための境界(同名source)。
    '      削除失敗時も取込は続けるが、err_log に1行(§4-1)+partial化(FA-6)。
    '      R18H FA-5: 全チャンクが既存と重複(acceptedCount=0)なら削除しない
    '      (新規0件で旧行を消すと資料が丸ごと0件になるため)。
    uiStep = "既存同名資料の置き換え"
    Dim rmNum As Long: rmNum = 0
    If acceptedCount > 0 Then
        On Error Resume Next
        Err.Clear
        modShelfStore.RemoveKnowledgeAndVectorsForSource sourceName, firstNewRow
        rmNum = Err.Number
        Dim rmDesc As String: rmDesc = Err.Description
        On Error GoTo Failed
        If rmNum <> 0 Then
            On Error Resume Next
            modLog.LogError "E0801", "modShelf.IngestFile", _
                "[" & uiStep & "] source=" & sourceName & " err#" & rmNum & ": " & rmDesc
            On Error GoTo Failed
        End If
        ' 削除で行が詰まった分だけ新しい行は上へ動く。読む位置をシートの事実
        ' から取り直す(新しい行は必ず末尾に連続している)。
        lastNewRow = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
        firstNewRow = lastNewRow - acceptedCount + 1
        If firstNewRow < 2 Then firstNewRow = 2
    Else
        ' R18H FA-5(A-M5): 旧行は温存したまま失敗として畳む。従来は素通りして
        ' done/0件で確定し、旧行だけが消えていた。
        On Error Resume Next
        modLog.LogUsage "ingest_all_dup", origin, "source=" & sourceName & _
            " 全" & chunkN & "件が既存と重複。旧データを保持しました"
        On Error GoTo Failed
        If isSelf Then
            modShelfStore.UpsertManifestRow path, sourceName, SafeFileDateTime(path), _
                SafeFileLen(path), -1, "failed", modShelfStore.MEMO_ALL_DUP, origin
        End If
        resultStatus = "failed"
        GoTo Finish
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

    ' 9.5) バッチ富化。config enrich_modeの既定"off"ならEnrichPendingは即0
    '      (2026-07-28 M-15: 呼び出しが1つも無く【死に設定】だった)。
    uiStep = "要約・キーワードの付与"
    On Error Resume Next
    modEnrich.EnrichPending
    On Error GoTo Failed

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

    ' 2026-08-03(R13-3a): 薄い抽出の検出(二段目の防衛)。判定と文言は
    ' modExtractorPdf.ThinExtractMemoFor。44ページの約款がchunks=1で
    ' 「登録成功」になる形(第2報RC1)をdoneと言わせない最後の関所。
    ' R13-F7: 拡張子も渡す(docx/xlsxを同じ関門にかけない)。
    Dim thinMemo As String: thinMemo = ""
    On Error Resume Next
    thinMemo = modExtractorPdf.ThinExtractMemoFor(pages, chunkN, modUtil.ExtOf(path))
    On Error GoTo Failed

    ' 2026-08-03(R14-4c): OCR打ち切りの正直なメモ(visionNote)。R15-FixA FA-4:
    ' カードのメモとpartial判定は別。「前回の続きから再開しました。」だけの
    ' ことがあり、材料にすると成功資料が「一部だけ」と名乗る。欠けの事実は
    ' pagesTruncatedが運ぶ。
    Dim cardMemo As String: cardMemo = thinMemo
    If LenB(cardMemo) = 0 Then cardMemo = visionNote
    ' R18H FA-6(A-M4): 旧行の削除に失敗した回は新旧2組の行が残ったまま
    ' 「done」と名乗っていた(検索結果重複に利用者が気付けない)。partialへ
    ' 倒し、自分で直せる唯一の手(もう一度取り込む)をメモに添える。
    If rmNum <> 0 Then cardMemo = Trim$(modShelfStore.MEMO_REPLACE_NG & " " & cardMemo)

    If pagesTruncated Or stillPending Or LenB(thinMemo) > 0 Or rmNum <> 0 Then
        resultStatus = "partial"
    Else
        resultStatus = "done"
        ' 2026-08-04(R15-FixB FB-1・A-M1/B-M): OCRの頁控えを消すのは【ここ】。
        ' 読み切った瞬間に消すと、後で落ちた資料がpartialのまま控えだけ失い
        ' 300頁を読み直す。doneと確定したことを知るのはこの位置だけ。
        On Error Resume Next
        modFeatures.InvokeFeature "vision", "OcrCachePurge", path
        On Error GoTo Failed
    End If

    If isSelf Then
        modShelfStore.UpsertManifestRow path, sourceName, SafeFileDateTime(path), SafeFileLen(path), acceptedCount, _
            resultStatus, cardMemo, origin
    End If

    If isSelf And (resultStatus = "done" Or resultStatus = "partial") Then
        On Error Resume Next
        modStats.Bump "ingest_files_total"
        modStats.AddExp "register"   ' 登録EXP(自己取込のみ。フォルダ自動同期では加算されない)
        On Error GoTo 0
        On Error Resume Next
        ' R15-8a(実機第4報 RC1): 画面の「127」とusage_logの「125」の食い違い
        ' (生成chunkN→重複排除後acceptedCountの乖離)がログに出ていなかった。
        ' dup=0のときは従来と同じ"chunks=N"(既存ログ・grep互換優先)。
        modLog.LogUsage "ingest", origin, "source=" & sourceName & " " & _
            modUtilText.IngestChunksDetail(acceptedCount, chunkN) & _
            " status=" & resultStatus
        On Error GoTo 0

        ' 要件B(2026-07-30 R3): バッジ評価の入口をIngestFileへ一本化。取込の
        ' 中核がどこからも呼んでおらず、スクショでは出てナレッジ登録では出ない
        ' 体験差があった。取得済みなら即Exitなので挙動は不変。
        On Error Resume Next
        modStats.EvaluateBadges
        On Error GoTo 0
    End If

    GoTo Finish

Failed:
    Dim failNum As Long: failNum = Err.Number
    Dim failDesc As String: failDesc = Err.Description
    ' ハンドラ稼働中はOERNが効かず、ここで起きたエラーが呼び出し元へ飛んで
    ' 本来の原因を上書きする。後始末の前にResumeで抜ける(実機err#462)。
    Resume FailedCleanup1
FailedCleanup1:
    On Error Resume Next
    modLog.LogError "E0801", "modShelf.IngestFile", _
        "[" & uiStep & "] path=" & modUtil.SafeLeft(path, 200) & " err#" & failNum & ": " & failDesc
    If isSelf Then
        ' R18-2a: chunk_countは-1(前値保持)。どの段で落ちたかに関わらず、
        ' 数え直していないのに0と名乗ってはいけない。実データと台帳のずれは
        ' SourceListの突合(R18-2b)が次の描画で必ず直す。
        modShelfStore.UpsertManifestRow path, sourceName, SafeFileDateTime(path), SafeFileLen(path), -1, _
            "failed", "取込に失敗[" & uiStep & "](#" & failNum & ")", origin
    End If
    resultStatus = "failed"
    outErrCode = "E0801"
    ' Resumeで抜けてハンドラ実行中の状態を解除する(On Error GoTo 0はトラップ
    ' 登録を消すだけでこの状態は消えず、そのままFinish:の後始末に落ちると
    ' そこで起きたエラーが呼び出し元へ素通りする)。
    Resume Finish

Finish:
    ' R18-2c(実機第5報⑧): 中間保存はここ1箇所に集約。従来はAddFilesResultの
    ' ループだけで、スクショ・ナレッジ登録・修正・共有登録・パック・部門
    ' チャンネルの6経路は一度も保存されなかった(調査agent1 §1-3)。取込の
    ' 出口はFinish1本なので全経路が一度に救われる(成功時のみ保存。
    ' 読み取り専用・共有ロックの後始末はSaveCheckpoint側が持つ)。
    ' R18H FA-7(A-M6): silent(無人同期)は120秒のスロットルを通す。100件の
    ' 同期で100回書き戻すとEDRスキャンが取込より重くなる端末がある。失うのは
    ' 最大2分ぶんで、同期末尾の1回(modShelfSync)は従来どおり必ず走る。
    If resultStatus = "done" Or resultStatus = "partial" Then
        Dim thrSec As Long: thrSec = 0
        If silent Then thrSec = 120
        On Error Resume Next
        modShelfBatch.SaveCheckpoint 1, thrSec
        On Error GoTo 0
    End If

    ' 2026-07-28(レビュー I-11): 同期中(silent)は1件ごとに再描画しない。
    ' 全読みして最大400枚のカードを描き直すため100件の同期では取込より重く
    ' なる。完了時に呼び出し側(SyncNow)が1回だけ描き直す。
    If Not silent Then
        On Error Resume Next
        modUIShelf.RenderShelf
        On Error GoTo 0
    End If

    mIngesting = False
    IngestFile = resultStatus
End Function

' DeleteSource - knowledge/vectors/manifestを一括削除+再描画
Public Sub DeleteSource(ByVal sourceName As String)
    modShelfStore.RemoveKnowledgeAndVectorsForSource sourceName
    modShelfStore.RemoveManifestRowForSource sourceName

    On Error Resume Next
    ' R18H FA-3: 意図して減らした事実を整合性マークへ即反映(保存はしない。
    ' ui_stateはブックと一緒に保存/破棄されるので辻褄は必ず合う)。無いと
    ' 次の起動で「資料が消えました」の虚偽警告が出る(A-M2)。
    modIntegrity.RecordSaveMark
    modUIShelf.RenderShelf
    On Error GoTo 0
End Sub

' SourceList - stats(i)="status|ingested_at|chunk_count|error_note|origin"
' R18-2b(実機第5報⑧): manifestのchunk_countとmy_knowledgeの実行数を突合し、
'   食い違えば【実行数を正】としてmanifestを直す(取込失敗がその列を壊すと
'   実データが生きていても「0件」に見えた)。全読み済みなので追加コスト無し。
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

    ' R18-2b: ここまでが manifest 由来。以降のパック由来は突合しない。
    Dim mfN As Long: mfN = outN
    Dim mfCount() As Long
    ReDim mfCount(0 To cap - 1)

    uiStep = "パック由来資料の集計"
    Dim wsK As Worksheet: Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If Not wsK Is Nothing Then
        Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
        If lastK >= 2 Then
            Dim kArr As Variant: kArr = wsK.Range(wsK.Cells(2, 2), wsK.Cells(lastK, 8)).Value
            ' kArrの列: 1=source,2=origin,7=added_at(2列目起点のため)
            Dim packCount As Object: Set packCount = CreateObject("Scripting.Dictionary")
            Dim packOrigin As Object: Set packOrigin = CreateObject("Scripting.Dictionary")
            Dim packAdded As Object: Set packAdded = CreateObject("Scripting.Dictionary")

            Dim r As Long
            Dim hint As Long: hint = -1
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
                Else
                    ' R18-2b: source別の実行数。パック由来は台帳に載らないので
                    ' 数えない。hintは【当たったときだけ】更新(-1に戻すと
                    ' 台帳に無い行が続く間ずっと全走査になる)。
                    Dim mi As Long
                    mi = modIntegrity.IndexOfName(tmpNames, mfN, CStr(kArr(r, 1)), hint)
                    If mi >= 0 Then
                        mfCount(mi) = mfCount(mi) + 1
                        hint = mi
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

    ' R18-2b: 突合(食い違う資料だけmanifestを直しusage_logに1行残す)。
    uiStep = "台帳と実データの突合"
    If Not wsK Is Nothing Then
        For i = 0 To mfN - 1
            tmpStats(i) = modIntegrity.ReconcileChunkCount(tmpNames(i), tmpStats(i), mfCount(i))
        Next i
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
    ' ハンドラ稼働中はOERNが効かず、ここで起きたエラーが呼び出し元へ飛んで
    ' 本来の原因を上書きする。後始末の前にResumeで抜ける(実機err#462)。
    Resume FailCleanup3
FailCleanup3:
    On Error Resume Next
    modLog.LogError "E0801", "modShelf.SourceList", "[" & uiStep & "] err#" & origNum & ": " & origDesc
    On Error GoTo 0
    ReDim names(0 To 0)
    ReDim stats(0 To 0)
    SourceList = 0
End Function

' TotalChunks - my_knowledgeの総行数(ヘッダ除く)
Public Function TotalChunks() As Long
    Dim wsK As Worksheet: Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If wsK Is Nothing Then Exit Function
    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then Exit Function
    TotalChunks = lastK - 1
End Function

' IsBusy - 取込が走っているか(2026-07-31 R7 B-2)。DoEventsで発火したクリックを
'   画面遷移側の入口で受け流す判定。焼き付きガードで全ボタン無反応にならぬよう、
'   IngestFileと同じ期限を過ぎたらbusy扱いにしない。
Public Function IsBusy() As Boolean
    ' R10c(H2): mIngestingはIngestFile1件ぶんしか覆わない。ファイル間の隙間も
    ' busyにするため一括取込のバッチガードをORで見る。
    If modShelfBatch.IsBatchBusy() Then
        IsBusy = True
        Exit Function
    End If
    If Not mIngesting Then Exit Function
    On Error Resume Next
    IsBusy = Not modShelfBatch.GuardExpiredNow(mIngestingSince, GUARD_EXPIRY_MIN)
    On Error GoTo 0
End Function

' ChunkKeyOrder - 拡張子別チャンク設定キーの名前を組み立てる(純ロジック)。
' 2026-08-03(R14-5a/5c): ChunkParamForはmodConfig依存で検証不能だが、
' 「拡張子別キーがbaseKeyより先」の順序はキー名だけ切り出せば固定できる
' (modTestsPure11)。extは念のため小文字化する。
Public Function ChunkKeyOrder(ByVal baseKey As String, ByVal ext As String) As String
    ChunkKeyOrder = baseKey & "_" & LCase$(Trim$(ext))
End Function

' ChunkParamFor - 拡張子別2段フォールバック(R14-5a・実機第3報 RC7)。
'   1) baseKey_ext(例 chunk_target_chars_pdf)→ 2) baseKey(グローバル既定)
'   → 3) fallbackDefault。chunk_modeは対象外(仕様どおりグローバルのまま)。
Private Function ChunkParamFor(ByVal baseKey As String, ByVal ext As String, _
                               ByVal fallbackDefault As Long) As Long
    ChunkParamFor = modConfig.GetLong(ChunkKeyOrder(baseKey, ext), _
                                      modConfig.GetLong(baseKey, fallbackDefault))
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

