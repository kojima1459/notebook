Attribute VB_Name = "modShelf"
Option Explicit

' ========================================
' modShelf - 本棚中核(ファイル取込・チャンク付番・重複排除・削除・一覧)
'
' 抽出→チャンク化→chunk_id付番(ハッシュ重複排除)→my_knowledge追記→
' manifest upsert→埋め込み(MASTER_SPEC §7.2)。再入guard(E0503)=mIngesting、
' 出口はFinish一本化(中間保存もそこ=R18-2c)。同名source置換は「書いてから
' 消す」(R18-2e)。シート書込は配列一括(§12)・manifestはself(§4)。
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


' IngestFile - MASTER_SPEC §7.2 の手順どおりに1ファイルを取込む。
'   戻り値=manifest status("done"/"partial"/"failed"/"image_pdf")
' silent (2026-07-28 レビュー L-11): 何十件もまとめて取り込む経路で1件ごとに
'   モーダルが出ると同期が止まる(同名衝突E0504は従来 silent を見ずに必ず
'   ダイアログを出し、人がいないと朝まで止まっていた)。True なら記録だけ残し、
'   結果は戻り値で伝える。
' outErrCode (2026-07-30・要件E): AddFilesResult が失敗理由の内訳(E0504/
'   image_pdf/E0302等)を集計する補助出口。Optional末尾追加なので既存の
'   呼び出し元(modVault/modUIShelf/modShelfSync等)は無改修。成功時は""。
' interactive (2026-08-04 R15-FixA FA-1・レビューA-H1/B-H1):「利用者がその場に
'   いる取込か」。silent(結果モーダルの抑止)とは別の問いなのに両方を silent
'   1本で表していたため、R15-7bの事前確認が【全経路で死んでいた】(手動の
'   一括取込も silent:=True で呼ぶため)。True は modShelfBatch.AddFilesResult
'   (利用者がFileDialogで選んだ直後)だけ。同期・起動時取込は従来どおり無確認。
Public Function IngestFile(ByVal path As String, ByVal origin As String, _
                           Optional ByVal silent As Boolean = False, _
                           Optional ByRef outErrCode As String = "", _
                           Optional ByVal interactive As Boolean = False) As String
    outErrCode = ""   ' 呼び出し側の使い回し変数でも必ずここで初期化する
    Dim resultStatus As String: resultStatus = "failed"
    Dim isSelf As Boolean: isSelf = (StrComp(origin, "self", vbTextCompare) = 0)

    ' 1) 再入guard(E0503)。焼き付きガードは自動解除。R15-1a: 失効は「開始から
    ' 30分」ではなく「最後のビートから30分」(取込の途中でガードが解けて二重
    ' 取込になるのを止める。実機第4報 RC5)。
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
        ' 2026-07-31(R7 B-2): ここは ShowError(モーダル)だった。取込中の
        ' DoEvents で発火したクリックが再入するたびダイアログが積み上がり実機で
        ' 「応答なし」に見えた。E0503のログは残し、利用者へは実況行で伝える。
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
    ' R15-FixA(FA-6): 中断の印はここでも必ず下ろす(modVault/modUIShelf は
    ' この関数を直接呼ぶため、前回の印が残ると押した覚えの無い取込が最初の頁
    ' 境界で止まる=幽霊中断)。一括取込・同期は呼ぶ前に自分で判定してから来る。
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
    '    H-6でここから抽出成功後へ、R18-2eで更に書込みの【後】へ移した。

    ' 4) 抽出
    uiStep = "ファイルからの本文抽出"
    Dim pages() As ExtractedPage
    Dim errCode As String, errDetail As String
    Dim extractOk As Boolean
    extractOk = modExtractor.ExtractFile(path, pages, errCode, errDetail)

    Dim pagesTruncated As Boolean: pagesTruncated = (extractOk And errCode = "PARTIAL_PAGES")

    If Not extractOk Then
        ' vision フォールバック(画像の文字起こし・画像PDFのOCR)は判断ごと
        ' modShelfVision が引き受ける(2026-07-31 R6)。上限ページで打ち切られたら
        ' pagesTruncated が立ち下の partial 判定へ合流。silent はそのまま渡す
        ' (R11-A C4: 無人同期からGSの案内ダイアログが出ると同期が止まる)。
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
                ' 2026-08-01(同梱-9): 50MBガード(txt/md/csv入力上限。E0501の
                ' 本棚上限とは別物)だけは errDetail の具体的な理由をそのまま
                ' 使う(従来は E0302 の一般文言で上書きし、カードのメモに
                ' 「大きすぎる」が出ていなかった)。他は生ログのため一律採用しない。
                If InStr(errDetail, "ファイルが大きすぎます") > 0 Then
                    failMsg = Split(errDetail, " [")(0)
                Else
                    ' 2026-08-03(R13-3b): 文言の決定は modLog.FriendlyFailMsg に
                    ' 一本化。errDetail に行動可能な案内があればそれを優先し、
                    ' docx/doc に Ghostscript の話をしない(実機第2報 RC6)。
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
            ' 2026-08-04(R15-FixB FB-5): 事前確認で「いいえ」を選んだ資料は
            ' 【失敗していない】。status は image_pdf のまま(カードの見え方・
            ' manifest の語彙は変えない)で理由コードだけ差し替え、一括取込の
            ' 集計が失敗と見送りを分けて数えられるようにする。判定は
            ' modUtilText.IsDeclineNote(先頭一致)の1本。
            If modUtilText.IsDeclineNote(failMsg) Then outErrCode = "declined"
            GoTo Finish
        End If
    End If

    ' 5) チャンク分割(§B)
    uiStep = "本文のチャンク分割"
    Dim chunks() As ShelfChunk
    Dim chunkN As Long
    ' R14-5a(実機第3報 RC7): 拡張子別の粒度設定(PDF/Word)がここまで届かず、
    ' Excelの行長と境界の偶然の一致でしか細かく割れていなかった。ChunkParamFor
    ' が「拡張子別→グローバル→既定」の2段フォールバックを1箇所で行う
    ' (chunk_modeはグローバルのまま)。
    Dim chunkExt As String: chunkExt = modUtil.ExtOf(path)
    chunkN = modChunker.ChunkPagesEx(pages, _
        ChunkParamFor("chunk_target_chars", chunkExt, CHUNK_TARGET_CHARS), _
        ChunkParamFor("chunk_overlap_chars", chunkExt, CHUNK_OVERLAP_CHARS), _
        ChunkParamFor("chunk_max_chars", chunkExt, 1800), _
        modConfig.GetString("chunk_mode", "structure"), chunks)

    ' 2026-07-29: 分割は1ページの失敗でファイル全体を捨てない。ただし黙って
    ' 減らすと「88チャンク入るはずが70」に気づけないので必ず1行残す。
    Dim skippedPages As Long: skippedPages = modChunker.SkippedPageCount()
    If skippedPages > 0 Then
        On Error Resume Next
        modLog.LogUsage "chunk_skipped_pages", "", _
            sourceName & ": " & skippedPages & "ページを処理できず飛ばしました"
        ' On Error GoTo 0 ではなく Failed へ戻す(0 にすると Failed ハンドラごと
        ' 解除され、以降の実行時エラーが素通りする=レビュー H-3 と同型)。
        On Error GoTo Failed
    End If

    If chunkN = 0 Then
        modLog.LogError "E0401", "modShelf.IngestFile", "source=" & sourceName
        If isSelf Then
            ' R18-2a: chunk_count は -1(前値保持)。旧データはまだ消していない。
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
    ' 置き換え対象(同名source)の行はまだ消していないのでハッシュ集合から除外。
    ' 除外しないと入れ直すチャンクが全部「自分自身との重複」になる(H-6)。
    Dim existingHashes As Object
    Set existingHashes = modShelfStore.BuildExistingHashSet(wsK, sourceName)

    ' 10列目 norm_text(R12-4): 照合用の正規化済みテキストをここで1回だけ作る。
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
            ' 保存する本文(SafeLeft後)から作る。1セルに収まらない場合だけ空に
            ' して検索側の遅延バックフィルへ委ねる(切り詰めて保存するとスコアが
            ' 保存済み行と未保存行で変わる)。
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

    ' 7.5) 旧データを消すのは【新しい行を全部書き終えてから】(R18-2e)。
    '      R11-H6 で「削除→書込み」を「抽出成功後に削除→書込み」へ直したが、
    '      その間はまだ空いており、32bit Excel の Range 一括代入は err#7 を
    '      起こし得る。落ちると Failed: が status=failed を書くだけで【消した
    '      旧データは戻らない】(調査agent1 §3(a) の真の消失経路)。反転すれば
    '      書込みが落ちても旧データは無傷(憲章§3-5)。firstNewRow を渡すのは
    '      いま書いた行を消させないため(同名sourceなので素直に消すと道連れ)。
    '      新旧が同時に存在する窓は取込中だけで、取込中は検索がUIロックで
    '      走らない(仕様の判断済み)。削除に失敗しても取込は続けるが、
    '      無言では済ませず既存E系の流儀で err_log に1行残す(§4-1)。
    uiStep = "既存同名資料の置き換え"
    On Error Resume Next
    Err.Clear
    modShelfStore.RemoveKnowledgeAndVectorsForSource sourceName, firstNewRow
    Dim rmNum As Long: rmNum = Err.Number
    Dim rmDesc As String: rmDesc = Err.Description
    On Error GoTo Failed
    If rmNum <> 0 Then
        On Error Resume Next
        modLog.LogError "E0801", "modShelf.IngestFile", _
            "[" & uiStep & "] source=" & sourceName & " err#" & rmNum & ": " & rmDesc
        On Error GoTo Failed
    End If
    ' 削除で行が詰まった分だけ新しい行は上へ動く。以降(10 status確定)が読む
    ' 位置をシートの事実から取り直す(新しい行は必ず末尾に連続している)。
    If acceptedCount > 0 Then
        lastNewRow = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
        firstNewRow = lastNewRow - acceptedCount + 1
        If firstNewRow < 2 Then firstNewRow = 2
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

    ' 9.5) バッチ富化(要約・キーワード付与)。config enrich_mode の既定 "off"
    '      なら EnrichPending は即0を返す。2026-07-28 レビュー M-15: 呼び出しが
    '      1つも無く enrich_mode は【どこからも読まれない死に設定】だった。
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
    ' modExtractorPdf.ThinExtractMemoFor。44ページの約款が chunks=1 で「登録
    ' 成功」になる形(実機第2報 RC1)を上流の分類が漏れても done と言わせない
    ' 最後の関所。R13-F7: 拡張子を渡す(docx/xlsx を同じ関門にかけると正当に
    ' 薄い資料を partial にしてしまう)。
    Dim thinMemo As String: thinMemo = ""
    On Error Resume Next
    thinMemo = modExtractorPdf.ThinExtractMemoFor(pages, chunkN, modUtil.ExtOf(path))
    On Error GoTo Failed

    ' 2026-08-03(R14-4c): OCRが上限ページで打ち切られたときの正直なメモ
    ' (modShelfVision が vision 側から受け取る visionNote)。R15-FixA FA-4:
    ' カードのメモと partial 判定は別。vision のメモは「前回の続きから再開
    ' しました。」だけのことがあり、partial の材料にすると【完全に成功した
    ' 資料】が「一部だけ」と名乗る。欠けの事実は pagesTruncated が運ぶ。
    Dim cardMemo As String: cardMemo = thinMemo
    If LenB(cardMemo) = 0 Then cardMemo = visionNote

    If pagesTruncated Or stillPending Or LenB(thinMemo) > 0 Then
        resultStatus = "partial"
    Else
        resultStatus = "done"
        ' 2026-08-04(R15-FixB FB-1・A-M1/B-M): OCRの頁控えを消すのは【ここ】。
        ' 従来は optOcrCache.FinishDoc(読み切った瞬間)で消しており、その後に
        ' 落ちると資料は partial のまま控えだけ消え、次の取込は300頁を読み直して
        ' いた。done として並んだことを知るのはこの位置だけ。鍵は【元のパス】。
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
        modStats.AddExp "register"   ' 登録EXP(自己取込のみ=フォルダ自動同期の大量取込では加算されない)
        On Error GoTo 0
        On Error Resume Next
        ' R15-8a(実機第4報 RC1): 画面の「127」とusage_logの「125」の食い違い
        ' (生成chunkN→重複排除後acceptedCountの乖離)がログに出ていなかった。
        ' dup=0 のときは従来と同じ "chunks=N"(既存ログ・grepとの互換優先)。
        modLog.LogUsage "ingest", origin, "source=" & sourceName & " " & _
            modUtilText.IngestChunksDetail(acceptedCount, chunkN) & _
            " status=" & resultStatus
        On Error GoTo 0

        ' 要件B(2026-07-30 R3): バッジ評価の入口をIngestFileへ一本化。従来の
        ' 呼び出しは modBoot/modUIMain/modUIShelf/modApp の4箇所だけで、取込の
        ' 中核であるIngestFile自体はどこからも呼んでおらず、スクショではバッジが
        ' 出てナレッジ登録では出ない体験差があった。CheckBadgeは取得済みなら即
        ' Exitするため大量取込中でもMsgBoxは増えない(挙動不変)。
        On Error Resume Next
        modStats.EvaluateBadges
        On Error GoTo 0
    End If

    GoTo Finish

Failed:
    Dim failNum As Long: failNum = Err.Number
    Dim failDesc As String: failDesc = Err.Description
    ' ハンドラ稼働中は OERN が効かず、ここで起きたエラーは呼び出し元へ飛んで
    ' 本来の原因を上書きする。後始末の前に Resume で抜ける(実機err#462)。
    Resume FailedCleanup1
FailedCleanup1:
    On Error Resume Next
    modLog.LogError "E0801", "modShelf.IngestFile", _
        "[" & uiStep & "] path=" & modUtil.SafeLeft(path, 200) & " err#" & failNum & ": " & failDesc
    If isSelf Then
        ' R18-2a: chunk_count は -1(前値保持)。どの段で落ちたかに関わらず、
        ' 件数を数え直していないのに0と名乗ってはいけない。実データと台帳の
        ' ずれは SourceList の突合(R18-2b)が次の描画で必ず直す。
        modShelfStore.UpsertManifestRow path, sourceName, SafeFileDateTime(path), SafeFileLen(path), -1, _
            "failed", "取込に失敗[" & uiStep & "](#" & failNum & ")", origin
    End If
    resultStatus = "failed"
    outErrCode = "E0801"
    ' Resume で抜けてハンドラ実行中の状態を解除する(On Error GoTo 0 は
    ' トラップの登録を消すだけでこの状態は消えず、そのまま Finish: の後始末へ
    ' 落ちると、そこで起きたエラーが呼び出し元へ素通りする)。
    Resume Finish

Finish:
    ' R18-2c(実機第5報⑧): 中間保存はここ1箇所に集約。従来 SaveCheckpoint を
    ' 呼ぶのは modShelfBatch.AddFilesResult のループ(手動FileDialog経路)だけで、
    ' スクショ・ナレッジ登録・修正・共有登録・パック・部門チャンネルの6経路は
    ' 【セッション中に一度も保存されなかった】(調査agent1 §1-3。R15-3aは
    ' 「1件ごとに保存」と謳ったが実装が1経路だけ=憲章§4-5)。取込の出口は
    ' この Finish 1本なので全経路が一度に救われる。成功(done/partial)時のみ
    ' 保存(silent の同期でも成功していれば保存してよい。読み取り専用・共有
    ' ロックの後始末は SaveCheckpoint 側が持つ)。
    If resultStatus = "done" Or resultStatus = "partial" Then
        On Error Resume Next
        modShelfBatch.SaveCheckpoint
        On Error GoTo 0
    End If

    ' 2026-07-28(レビュー I-11): 同期中(silent)は1件ごとに再描画しない。
    ' RenderShelf は manifest と my_knowledge を全読みして最大400枚のカードを
    ' 描き直すため、100件の同期では O(件数×総チャンク)で取込より重くなる。
    ' 同期の完了時に呼び出し側(SyncNow)が1回だけ描き直す。
    If Not silent Then
        On Error Resume Next
        modUIShelf.RenderShelf
        On Error GoTo 0
    End If

    mIngesting = False
    IngestFile = resultStatus
End Function

' DeleteSource - knowledge/vectors/manifest から一括削除+再描画
Public Sub DeleteSource(ByVal sourceName As String)
    modShelfStore.RemoveKnowledgeAndVectorsForSource sourceName
    modShelfStore.RemoveManifestRowForSource sourceName

    On Error Resume Next
    modUIShelf.RenderShelf
    On Error GoTo 0
End Sub

' SourceList - stats(i)="status|ingested_at|chunk_count|error_note|origin"
' R18-2b(実機第5報⑧): manifest の chunk_count と my_knowledge の実行数を
'   突合し、食い違えば【実行数を正】として manifest を直す(カードの件数は
'   manifest の5列目をそのまま出す作りで、取込の失敗がその列を壊すと実データが
'   生きていても「0件」に見えた)。既に my_knowledge を全読みしているので追加
'   コストはほぼ無い。直すのは manifest 由来(先頭 mfN 件)だけ。
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
            ' kArrの列: 1=source, 2=origin, 7=added_at(2列目起点のため)
            Dim packCount As Object: Set packCount = CreateObject("Scripting.Dictionary")
            Dim packOrigin As Object: Set packOrigin = CreateObject("Scripting.Dictionary")
            Dim packAdded As Object: Set packAdded = CreateObject("Scripting.Dictionary")

            Dim r As Long
            Dim hint As Long: hint = -1
            For r = LBound(kArr, 1) To UBound(kArr, 1)
                ' R18-2b: source別の実行数を数える(manifest由来のカードのみ)。
                hint = modIntegrity.IndexOfName(tmpNames, mfN, CStr(kArr(r, 1)), hint)
                If hint >= 0 Then mfCount(hint) = mfCount(hint) + 1
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

    ' R18-2b: 突合。食い違う資料だけ manifest を直し usage_log に1行残す
    ' (直せなくてもカードは実行数を表示する)。
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
    ' ハンドラ稼働中は OERN が効かず、ここで起きたエラーは呼び出し元へ飛んで
    ' 本来の原因を上書きする。後始末の前に Resume で抜ける(実機err#462)。
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

' IsBusy - 取込が走っているか(2026-07-31 R7 B-2)。抽出ループの DoEvents で
'   発火したクリックを画面遷移側の入口で受け流すための判定。焼き付きガードで
'   全ボタンが無反応にならないよう、IngestFile と同じ期限を過ぎたら busy 扱い
'   にしない。
Public Function IsBusy() As Boolean
    ' R10c(H2): mIngesting は IngestFile 1件ぶんしか覆わない。ファイル間の
    ' 隙間(バナー更新やトーストのDoEvents中)も busy にするため、一括取込の
    ' バッチガード(modShelfBatch)を OR で見る。
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
' 2026-08-03(R14-5a/5c): ChunkParamFor自体はmodConfig.GetLong依存で検証不能だが、
' 「拡張子別キーが baseKey より先に見られる」順序はキー名の組み立てだけを
' 切り出せば固定できる(modTestsPure11)。extは念のため小文字化する。
Public Function ChunkKeyOrder(ByVal baseKey As String, ByVal ext As String) As String
    ChunkKeyOrder = baseKey & "_" & LCase$(Trim$(ext))
End Function

' ChunkParamFor - 拡張子別2段フォールバック(R14-5a・実機第3報 RC7)。
'   1) baseKey_ext(例 chunk_target_chars_pdf)→ 2) baseKey(グローバル既定)
'   → 3) fallbackDefault。chunk_mode は対象外(仕様どおりグローバルのまま)。
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

