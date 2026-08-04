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
' IngestFile - MASTER_SPEC §7.2 の手順どおりに1ファイルを取込む。
'   戻り値=manifest status("done"/"partial"/"failed"/"image_pdf")
'
' silent (2026-07-28 レビュー L-11):
'   フォルダ同期のように「まとめて何十件も取り込む」経路では、1件ごとに
'   モーダルを出されると同期が止まる。同名衝突(E0504)は従来 silent を見ずに
'   必ずダイアログを出していたため、同期中に人がいないと朝まで止まっていた。
'   silent:=True のときは記録だけ残し、結果は戻り値で呼び出し側へ伝える。
'
' outErrCode (2026-07-30・要件E):
'   AddFilesResultが失敗理由の内訳(E0504/image_pdf/E0302等)を集計するための
'   補助出口。Optionalの末尾追加なので既存の呼び出し元(modVault/modUIShelf/
'   modShelfSync等)は無改修のまま動く。成功時は""のまま。
Public Function IngestFile(ByVal path As String, ByVal origin As String, _
                           Optional ByVal silent As Boolean = False, _
                           Optional ByRef outErrCode As String = "") As String
    outErrCode = ""   ' 呼び出し側が使い回した変数でも必ずここで初期化する
    Dim resultStatus As String: resultStatus = "failed"
    Dim isSelf As Boolean: isSelf = (StrComp(origin, "self", vbTextCompare) = 0)

    ' 1) 再入guard(E0503)。焼き付きガードは自動解除。
    ' R15-1a: 失効は「開始から30分」ではなく「最後のビートから30分」で判定する
    ' (長い取込の途中でガードが解けて二重取込になるのを止める。実機第4報 RC5)。
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
        ' DoEvents で発火したクリックが再入するたびにダイアログが積み上がり、
        ' 実機では「Excelが応答なし」に見えていた。E0503のログは従来どおり
        ' 残し、利用者へは実況行で「待てばよい」ことだけを伝える。
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
        ' vision フォールバック(画像ファイルの文字起こし・画像PDFのOCR)は
        ' 判断ごと modShelfVision が引き受ける(2026-07-31 R6)。上限ページで
        ' 打ち切られた場合は pagesTruncated が立ち、下の partial 判定へ合流する。
        Dim visionNote As String
        ' silent をそのまま渡す(R11-A C4)。無人のフォルダ同期から
        ' Ghostscript の案内ダイアログが出ると、そこで同期が止まる。
        If modShelfVision.TryVisionFallback(path, errCode, pages, pagesTruncated, _
                                            visionNote, silent) Then
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
                ' 本棚上限とは別物)は errDetail に具体的な理由(上限・実サイズ)
                ' が入っているのに、従来は errCode="E0302" の一般文言
                ' (「他のアプリで開いている…Ghostscript…」)で上書きしていて、
                ' 本棚カードのメモに「大きすぎる」ことが一切出ていなかった。
                ' このケースだけは具体的な理由文をそのまま使う(他のerrDetailは
                ' 技術的な生ログのままの箇所があるため一律採用はしない)。
                If InStr(errDetail, "ファイルが大きすぎます") > 0 Then
                    failMsg = Split(errDetail, " [")(0)
                Else
                    ' 2026-08-03(R13-3b): 文言の決定は modLog.FriendlyFailMsg に
                    ' 一本化した。errDetail に行動可能な案内(Wordを開いたままに
                    ' …)があればそれを優先し、docx/doc に Ghostscript の話を
                    ' しない(実機第2報 RC6)。
                    failMsg = modLog.FriendlyFailMsg(errCode, errDetail, modUtil.ExtOf(path))
                End If
            End If
            If isSelf Then
                modShelfStore.UpsertManifestRow path, sourceName, SafeFileDateTime(path), SafeFileLen(path), 0, _
                    failStatus, failMsg, origin
            End If
            resultStatus = failStatus
            outErrCode = errCode
            GoTo Finish
        End If
    End If

    ' 5) チャンク分割(§B)
    uiStep = "本文のチャンク分割"
    Dim chunks() As ShelfChunk
    Dim chunkN As Long
    ' R14-5a(実機第3報 RC7): 拡張子別の粒度設定(PDF/Word)がここまで
    ' 届いていなかったため、Excelの行長とスライディング境界の偶然の一致で
    ' しか細かく割れていなかった。ChunkParamForが「拡張子別 → グローバル →
    ' 既定」の2段フォールバックを1箇所で行う(chunk_modeはグローバルのまま)。
    Dim chunkExt As String: chunkExt = modUtil.ExtOf(path)
    chunkN = modChunker.ChunkPagesEx(pages, _
        ChunkParamFor("chunk_target_chars", chunkExt, CHUNK_TARGET_CHARS), _
        ChunkParamFor("chunk_overlap_chars", chunkExt, CHUNK_OVERLAP_CHARS), _
        ChunkParamFor("chunk_max_chars", chunkExt, 1800), _
        modConfig.GetString("chunk_mode", "structure"), chunks)

    ' 2026-07-29: 分割は1ページの失敗でファイル全体を捨てないようにした。
    ' ただし黙って減らすと「88チャンク入るはずが70だった」に誰も気づけない。
    ' 飛ばしたページ数は必ず1行残す。
    Dim skippedPages As Long: skippedPages = modChunker.SkippedPageCount()
    If skippedPages > 0 Then
        On Error Resume Next
        modLog.LogUsage "chunk_skipped_pages", "", _
            sourceName & ": " & skippedPages & "ページを処理できず飛ばしました"
        ' ここは On Error GoTo 0 ではなく Failed へ戻す。0 にすると
        ' この関数の Failed ハンドラごと解除され、以降の実行時エラーが
        ' 利用者の画面へ素通りする(レビュー H-3 と同じ型の事故)。
        On Error GoTo Failed
    End If

    If chunkN = 0 Then
        modLog.LogError "E0401", "modShelf.IngestFile", "source=" & sourceName
        If isSelf Then
            modShelfStore.UpsertManifestRow path, sourceName, SafeFileDateTime(path), SafeFileLen(path), 0, _
                "failed", modLog.FriendlyMessage("E0401") & "(コード: E0401)", origin
        End If
        resultStatus = "failed"
        outErrCode = "E0401"
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

    ' 10列目 norm_text(R12-4): 照合用の正規化済みテキストをここで1回だけ作る。
    ' 検索のたびに全チャンクを正規化し直していたぶんが丸ごと消える(総量は不変)。
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
            ' しておき、検索側の遅延バックフィルへ委ねる(切り詰めた値を保存すると
            ' 保存済み行と未保存行でスコアが変わる)。
            Dim normDoc As String
            normDoc = modSparse.MatchDocText("", "", sourceName, CStr(outRows(acceptedCount, COL_FULLTEXT)))
            If Len(normDoc) <= 32000 Then outRows(acceptedCount, modShelfStore.COL_K_NORM) = normDoc
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
            wsK.Range(wsK.Cells(wTop, 1), wsK.Cells(wTop + batchN - 1, modShelfStore.COL_K_NORM)).Value = batchArr
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

    ' 9.5) バッチ富化(要約・キーワード付与)。config enrich_mode の既定は "off"
    '      で、その場合 EnrichPending は即0を返して何もしない。
    '      2026-07-28(レビュー M-15): この呼び出しがソース全体に1つも無く、
    '      config の enrich_mode は【どこからも読まれない死に設定】だった
    '      (設定台帳には載っているので、管理者は効くと思って設定する)。
    '      既定offなので、配線しても既定の挙動は1ミリも変わらない。
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
    ' modExtractorPdf.ThinExtractMemoFor が持つ(ここは呼び出し1行)。
    ' 44ページの約款が chunks=1 で「登録成功」になる形(実機第2報 RC1)を、
    ' 上流の分類が漏れたときでも done と言わせないための最後の関所。
    ' R13-F7: 拡張子を渡す。docxの「ページ」やxlsxの「シート」まで同じ関門に
    ' かけると、PDF/OCR前提の文言で正当に薄い資料を partial にしてしまう。
    Dim thinMemo As String: thinMemo = ""
    On Error Resume Next
    thinMemo = modExtractorPdf.ThinExtractMemoFor(pages, chunkN, modUtil.ExtOf(path))
    On Error GoTo Failed

    ' 2026-08-03(R14-4c): OCRが上限ページで打ち切られたときの正直なメモ。
    ' modShelfVision が vision 側から受け取っている(成功時のvisionNote)。
    ' 本棚カードの partial は、メモがあればそれを出し、無いとき(=ベクトル化
    ' 待ち。同期で本当に続きから進む)だけ再開の案内を出す。
    If LenB(thinMemo) = 0 Then thinMemo = visionNote

    If pagesTruncated Or stillPending Or LenB(thinMemo) > 0 Then
        resultStatus = "partial"
    Else
        resultStatus = "done"
    End If

    If isSelf Then
        modShelfStore.UpsertManifestRow path, sourceName, SafeFileDateTime(path), SafeFileLen(path), acceptedCount, _
            resultStatus, thinMemo, origin
    End If

    If isSelf And (resultStatus = "done" Or resultStatus = "partial") Then
        On Error Resume Next
        modStats.Bump "ingest_files_total"
        modStats.AddExp "register"   ' 登録EXP(自己取込のみ=フォルダ自動同期の大量取込では加算されない)
        On Error GoTo 0
        On Error Resume Next
        ' R15-8a(実機第4報 RC1): 画面の「127」とusage_logの「125」が食い違って
        ' 見える原因(生成chunkN→重複排除後acceptedCountの乖離)がログにも
        ' 出ておらず診断できなかった。dup=0(重複なし)のときは従来と同じ
        ' "chunks=N" のまま(既存ログ・grepとの互換を優先=IngestChunksDetail参照)。
        modLog.LogUsage "ingest", origin, "source=" & sourceName & " " & _
            modUtilText.IngestChunksDetail(acceptedCount, chunkN) & _
            " status=" & resultStatus
        On Error GoTo 0

        ' 要件B(2026-07-30 R3): バッジ評価の入口をIngestFileへ一本化する。
        ' 従来EvaluateBadgesを呼んでいたのはmodBoot/modUIMain/modUIShelf/
        ' modApp の4箇所のみで、取込の中核であるIngestFile自体はどこからも
        ' 呼んでいなかった(R3要件定義書 背景2)。そのため、直後に呼び出し側が
        ' 自分でEvaluateBadgesを呼ぶ経路(スクショ取込等)ではバッジが出るのに、
        ' 呼ばない経路(ナレッジ登録等)では一切出ない、という体験差があった。
        ' ここが唯一の評価点になれば、以降どの入口(ダイアログ/フォルダ同期/
        ' ナレッジ登録/スクショ/チャット)を新設しても取りこぼさない。
        ' CheckBadgeは取得済みなら即Exitするため、フォルダ同期の大量取込中に
        ' 呼んでも新規獲得の瞬間以外はMsgBoxを出さない(既存の挙動を変えない。
        ' 詳細はR3要件定義書 要件B注記)。既存のRefreshBadgesAndDashboard等の
        ' 呼び出しは冪等なので削除しない。
        On Error Resume Next
        modStats.EvaluateBadges
        On Error GoTo 0
    End If

    GoTo Finish

Failed:
    Dim failNum As Long: failNum = Err.Number
    Dim failDesc As String: failDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailedCleanup1
FailedCleanup1:
    On Error Resume Next
    modLog.LogError "E0801", "modShelf.IngestFile", _
        "[" & uiStep & "] path=" & modUtil.SafeLeft(path, 200) & " err#" & failNum & ": " & failDesc
    If isSelf Then
        modShelfStore.UpsertManifestRow path, sourceName, SafeFileDateTime(path), SafeFileLen(path), 0, _
            "failed", "取込に失敗[" & uiStep & "](#" & failNum & ")", origin
    End If
    resultStatus = "failed"
    outErrCode = "E0801"
    ' Resume で抜けることでハンドラ実行中の状態を解除する。
    ' On Error GoTo 0 はトラップの登録を消すだけで、この状態は消えない。
    ' 消えないまま Finish: の後始末へ落ちると、そこで起きたエラーが
    ' 呼び出し元へ素通りする(フォルダ同期なら残りのファイルが中断する)。
    Resume Finish

Finish:
    ' 2026-07-28(レビュー I-11): 同期中(silent)は1件ごとに再描画しない。
    ' RenderShelf は manifest と my_knowledge を全読みして最大400枚の
    ' カードを描き直すので、フォルダ同期で100件取り込むと
    ' O(件数×総チャンク)になり、取込より描画の方が時間を食う。
    ' 同期の完了時に呼び出し側(SyncNow)が1回だけ描き直す。
    If Not silent Then
        On Error Resume Next
        modUIShelf.RenderShelf
        On Error GoTo 0
    End If

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
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup3
FailCleanup3:
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

'
' IsBusy - 取込が走っているか(2026-07-31 R7 B-2)。
'   抽出ループの DoEvents で発火したクリックを、画面遷移側の入口で
'   受け流すための判定。再入ガード mIngesting をそのまま公開する。
'   焼き付いたガードで全ボタンが無反応になるのを避けるため、
'   IngestFile と同じ期限(GUARD_EXPIRY_MIN)を過ぎたものは busy とみなさない。
'
Public Function IsBusy() As Boolean
    ' R10c(H2): mIngesting は IngestFile 1件ぶんしか覆わない。ファイルと
    ' ファイルの隙間(進捗バナー更新やトーストのDoEvents中)も busy として
    ' 返すため、一括取込のバッチガード(modShelfBatch)を OR で見る。
    If modShelfBatch.IsBatchBusy() Then
        IsBusy = True
        Exit Function
    End If
    If Not mIngesting Then Exit Function
    On Error Resume Next
    IsBusy = Not modShelfBatch.GuardExpiredNow(mIngestingSince, GUARD_EXPIRY_MIN)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' ChunkKeyOrder - 拡張子別チャンク設定キーの名前を組み立てる(純ロジック)。
' ----------------------------------------------------------------------------
' 2026-08-03(R14-5a/5c): ChunkParamFor自体はmodConfig.GetLong(configシート)へ
' 依存するためExcelなしでは検証できない。「拡張子別キーが baseKey より先に
' 見られる」というフォールバック順序は、キー名の組み立てだけを切り出せば
' Excel無しで固定できる(modTestsPure11)。extは modUtil.ExtOf の戻り値
' (常に小文字)を渡す想定だが、ここでも念のため小文字化する(二重の安全)。
Public Function ChunkKeyOrder(ByVal baseKey As String, ByVal ext As String) As String
    ChunkKeyOrder = baseKey & "_" & LCase$(Trim$(ext))
End Function

' ----------------------------------------------------------------------------
' ChunkParamFor - 拡張子別2段フォールバック(R14-5a・実機第3報 RC7)。
'   1) baseKey_ext (例: chunk_target_chars_pdf)
'   2) baseKey     (グローバル既定。例: chunk_target_chars)
'   3) fallbackDefault(configすら読めないときの最終値)
'   chunk_mode は拡張子別分岐の対象外(仕様どおりグローバルのまま)。
' ----------------------------------------------------------------------------
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

