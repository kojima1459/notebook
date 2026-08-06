Attribute VB_Name = "modBackfill"
Option Explicit

' ============================================================================
' modBackfill - 再取込ゼロの「資料の仕上げ」バックフィル
'   2026-08-06(R20-3・実機第7報②)
' ----------------------------------------------------------------------------
' なぜ要るのか:
'   R17(Phase1構造メタ/Phase2章要約/Phase3名寄せ)より前に取り込んだ資料は、
'   my_knowledge に breadcrumb付きの full_text(【資料名>章>条】+本文)を
'   既に持っているのに、chunk_meta(section_path/refs_out)とdoc_outline
'   (章要約)が1行も無いままになっている。ところが breadcrumb がある以上、
'   3機能を作るのに必要な材料は【もう保存済み】で、足りないのはそこから
'   後追いで作る処理だけ。ファイルの再抽出・OCR・埋め込みの再計算は
'   一切要らない(D班調査で確定済み)。
'
' 設計判断:
'   ・凍結(modShelf)は1文字も触らない。ここは modShelf.IngestFile の
'     Finish手順(Phase1メタ書込+Phase2章要約)と【同じ材料・同じ関数】を、
'     取込を経由しない独立入口から再実行する。既存のIngestFileの1箇所という
'     前提を崩すが、呼び先(modChunkMeta.MetaOf/modChunkMetaStore.WriteMetaRows/
'     modOutlineBuild.BuildOutlineFor)はどれも元から Public で、二重化していない。
'   ・Phase1(構造メタ)は保存済みの my_knowledge.full_text から直接
'     modChunkMeta.MetaOf を再実行する。embed_prefix_breadcrumb=True(既定)なら
'     保存本文は取込時にMetaOfへ渡していたテキストと資料名の埋め込み位置が
'     違うだけ(ExtractSectionPathは先頭要素=資料名を常に捨てるので無関係)で、
'     結果は取込直後と同一になる。False(breadcrumb行そのものを保存前に除去)の
'     資料は【保存本文に材料が残っていない】ため、この経路では直せない
'     (DetectLegacyDocsが「仕上げ不可」に分類し、再取込を案内する)。
'   ・Phase2(章要約)・Phase3(名寄せ)は modOutlineBuild.BuildOutlineFor を
'     1回呼ぶだけ。ゲート(graph_outline)・中断(■中断で【そこまでの章を保存】)・
'     1章の失敗握り・Phase3への連鎖(SynStep)は全てその関数の中に閉じており、
'     ここから重複して判断を持たない(単一情報源。憲章§4-5)。
'   ・1冊の失敗が全体を止めない: BackfillAllは各冊をOn Error配下で処理し、
'     失敗しても次の資料へ進む。AI呼び出しを伴うのはPhase2/3(BuildOutlineFor
'     が内部で握る)だけで、Phase1はLLMを1回も呼ばない純文字列処理。
' ============================================================================

Public Const STATUS_NEEDS As String = "needs_backfill"
Public Const STATUS_CANNOT As String = "cannot_backfill"

Public Const OUTCOME_NONE As String = "NONE"
Public Const OUTCOME_CANCELLED As String = "CANCELLED"
Public Const OUTCOME_DONE As String = "DONE"

' my_knowledge の列(modShelfStore/modOutlineBuild の定義と同じ並び。読むだけ)。
Private Const COL_ID As Long = 1
Private Const COL_SOURCE As Long = 2
Private Const COL_FULLTEXT As Long = 7

' ----------------------------------------------------------------------------
' DetectLegacyDocs - my_knowledge をsource単位に走査し、未仕上げの資料を
'   列挙する。戻り値は Collection<String>("資料名|状態")。状態は
'   STATUS_NEEDS(仕上げで直せる)/STATUS_CANNOT(breadcrumb形式でなく
'   仕上げ不可・再取込のみ)のいずれか。仕上げ済み(chunk_meta+doc_outline
'   とも揃っている)資料は含めない。AIは1回も呼ばない(判定のみ)。
' ----------------------------------------------------------------------------
Public Function DetectLegacyDocs() As Collection
    Dim outCol As New Collection
    Set DetectLegacyDocs = outCol

    On Error Resume Next
    Dim metaSet As Object: Set metaSet = BuildMetaIdSet()
    Dim outlineSet As Object: Set outlineSet = BuildOutlineSourceSet()
    On Error GoTo 0
    If metaSet Is Nothing Or outlineSet Is Nothing Then Exit Function

    Dim srcNames() As String, chunkCount() As Long, metaCount() As Long, sample() As String
    Dim n As Long
    n = ScanKnowledge(metaSet, srcNames, chunkCount, metaCount, sample)
    If n < 1 Then Exit Function

    ' 2026-08-06 R20H FA-14(波C裁定済み): graph_outline=offのときは
    ' doc_outlineが1行も無いのが正常(Phase2そのものが切られているため)で、
    ' 「未仕上げ」の根拠にはならない。offならhasOutlineを常に満たした
    ' 扱いにする(ClassifyDoc自体は無変更・純関数のまま)。
    Dim outlineGateOn As Boolean: outlineGateOn = GraphOutlineGateOn()

    Dim i As Long
    For i = 0 To n - 1
        If chunkCount(i) > 0 Then
            Dim hasMeta As Boolean: hasMeta = (metaCount(i) >= chunkCount(i))
            Dim hasOutline As Boolean: hasOutline = outlineSet.Exists(srcNames(i))
            If Not outlineGateOn Then hasOutline = True
            Dim crumbOk As Boolean: crumbOk = LooksLikeBreadcrumbLine(sample(i))
            Dim status As String: status = ClassifyDoc(hasMeta, hasOutline, crumbOk)
            If LenB(status) > 0 Then outCol.Add srcNames(i) & "|" & status
        End If
    Next i
End Function

' config graph_outline(既定on)。offのときだけdoc_outline欠落を「未仕上げ」
' の根拠から外す(modOutlineBuild.GateOnと同じ作法。各モジュールが自分の
' 呼び出し文脈でこの1行を持つ既存の設計を踏襲)。
Private Function GraphOutlineGateOn() As Boolean
    GraphOutlineGateOn = (LCase$(Trim$(modConfig.GetString("graph_outline", "on"))) <> "off")
End Function

' CountByStatus - DetectLegacyDocsの結果を状態ごとに数える(呼び出し側は
'   区切り文字の形を知らなくてよい)。
Public Function CountByStatus(ByVal cands As Collection, ByVal status As String) As Long
    If cands Is Nothing Then Exit Function
    Dim v As Variant
    For Each v In cands
        Dim item As String: item = CStr(v)
        Dim p As Long: p = InStr(item, "|")
        If p > 0 Then
            If Mid$(item, p + 1) = status Then CountByStatus = CountByStatus + 1
        End If
    Next v
End Function

' ----------------------------------------------------------------------------
' BackfillOne - 1資料ぶんを仕上げる。Phase1(構造メタ再計算)→Phase2/3
'   (modOutlineBuild.BuildOutlineFor に委譲)。戻り値はPhase1が書けたか
'   (=対象チャンクが見つかり、chunk_meta へ書き込めたか)。中断・失敗の
'   握りはBuildOutlineFor側の作法をそのまま踏襲するため、ここでの成否は
'   Phase1(構造メタ)だけの結果を表す(Phase2が中断・失敗しても、そこまでの
'   章は保存され、資料自体はTrueのまま返る=中断は「失敗」ではない)。
' ----------------------------------------------------------------------------
Public Function BackfillOne(ByVal sourceName As String) As Boolean
    If LenB(Trim$(sourceName)) = 0 Then Exit Function
    On Error GoTo Failed

    Dim ids() As String, texts() As String
    Dim nSrc As Long
    nSrc = CollectSourceChunks(sourceName, ids, texts)
    If nSrc < 1 Then Exit Function

    Dim mPath() As String: ReDim mPath(1 To nSrc)
    Dim mRefs() As String: ReDim mRefs(1 To nSrc)
    Dim i As Long
    For i = 1 To nSrc
        modChunkMeta.MetaOf texts(i), mPath(i), mRefs(i)
    Next i

    ' 冪等化: 既存分があれば道連れに作り直してから全件を書く(再実行しても
    ' 行が積み増しにならない。modShelf.DeleteSourceと同じ keepFromRow=0)。
    modChunkMetaStore.RemoveMetaForSource sourceName, 0
    modChunkMetaStore.WriteMetaRows ids, mPath, mRefs, nSrc

    On Error Resume Next
    modLog.LogUsage "backfill_meta", "ingest", "source=" & sourceName & " chunks=" & nSrc
    On Error GoTo Failed

    ' Phase2(章要約)+Phase3(名寄せ連鎖)は既存の唯一の入口へ委譲。ゲート・
    ' 中断・失敗握りはBuildOutlineFor側に閉じているため、ここは1行(冒頭参照)。
    On Error Resume Next
    modOutlineBuild.BuildOutlineFor sourceName
    On Error GoTo 0

    BackfillOne = True
    Exit Function

Failed:
    Dim d As String: d = "err#" & Err.Number & " " & Err.Description
    Resume FailedLog
FailedLog:
    Err.Clear
    On Error Resume Next
    modLog.LogError "E0801", "modBackfill.BackfillOne", "source=" & sourceName & " " & d
    On Error GoTo 0
    BackfillOne = False
End Function

' ----------------------------------------------------------------------------
' BackfillAll - 未仕上げの資料をまとめて仕上げる唯一の入口。開始前に確認
'   ダイアログを出し(0件なら出さない)、「いいえ」ならキャンセル扱い。
'   戻り値は "OUTCOME|メッセージ" の形(OUTCOME=OUTCOME_*)。呼び出し側は
'   OUTCOME_CANCELLEDのときだけ何も表示しない。
' ----------------------------------------------------------------------------
Public Function BackfillAll() As String
    Dim cands As Collection: Set cands = DetectLegacyDocs()
    Dim needN As Long: needN = CountByStatus(cands, STATUS_NEEDS)
    Dim cannotN As Long: cannotN = CountByStatus(cands, STATUS_CANNOT)

    If needN < 1 Then
        BackfillAll = OUTCOME_NONE & "|" & "すべての資料は仕上げ済みです。"
        Exit Function
    End If

    If MsgBox(ConfirmText(needN, cannotN), vbYesNo + vbQuestion, "資料の仕上げ") <> vbYes Then
        BackfillAll = OUTCOME_CANCELLED & "|"
        Exit Function
    End If

    On Error Resume Next
    modShelfBatch.ResetCancel
    modShelfBatch.ShowIngestBanner "資料の仕上げを開始します…"
    On Error GoTo 0

    Dim total As Long, okN As Long, ngN As Long
    Dim v As Variant
    For Each v In cands
        Dim item As String: item = CStr(v)
        Dim p As Long: p = InStr(item, "|")
        If p > 0 Then
            If Mid$(item, p + 1) = STATUS_NEEDS Then
                Dim src As String: src = Left$(item, p - 1)
                total = total + 1
                On Error Resume Next
                modShelfBatch.StageBanner "資料の仕上げ中… " & total & "/" & needN & _
                    "冊目(" & modUtil.SafeLeft(src, 40) & ")"
                On Error GoTo 0
                If BackfillOne(src) Then
                    okN = okN + 1
                Else
                    ngN = ngN + 1
                End If
                If modShelfBatch.CancelRequested() Then Exit For
            End If
        End If
    Next v

    On Error Resume Next
    modUIMain.SetStage ""
    modLog.LogUsage "backfill_all", "ingest", _
        "ok=" & okN & " ng=" & ngN & " cannot=" & cannotN
    On Error GoTo 0

    BackfillAll = OUTCOME_DONE & "|" & ResultText(okN, ngN, cannotN)
End Function

' ============================================================================
' 純ロジック(LibreOfficeの実行テストで固定する)
' ============================================================================

' ----------------------------------------------------------------------------
' ClassifyDoc - 1資料の判定表。hasMeta/hasOutlineが両方揃っていれば仕上げ済み
'   (空文字=対象外)。breadcrumbOkがFalseなら、他がどうであれ「仕上げ不可」
'   (Phase1の材料そのものが保存本文に無いため、この経路では直せない)。
' ----------------------------------------------------------------------------
Public Function ClassifyDoc(ByVal hasMeta As Boolean, ByVal hasOutline As Boolean, _
                            ByVal breadcrumbOk As Boolean) As String
    If hasMeta And hasOutline Then Exit Function
    If Not breadcrumbOk Then
        ClassifyDoc = STATUS_CANNOT
    Else
        ClassifyDoc = STATUS_NEEDS
    End If
End Function

' ----------------------------------------------------------------------------
' LooksLikeBreadcrumbLine - full_text の先頭行が「【…】」breadcrumb形式か。
'   modChunkMeta.ExtractSectionPathの冒頭ガード(先頭が「【」・「】」が3文字目
'   以降にある)と同じ判定をサンプル診断用に持つ(あちらはPrivate相当の内部
'   ガードで公開されていないため、ここは1章ぶんの薄い再実装)。
' ----------------------------------------------------------------------------
Public Function LooksLikeBreadcrumbLine(ByVal fullText As String) As Boolean
    If LenB(fullText) = 0 Then Exit Function
    Dim line1 As String
    Dim lf As Long: lf = InStr(fullText, vbLf)
    If lf > 0 Then
        line1 = Left$(fullText, lf - 1)
    Else
        line1 = fullText
    End If
    LooksLikeBreadcrumbLine = (Left$(line1, 1) = "【" And InStr(line1, "】") >= 3)
End Function

' ----------------------------------------------------------------------------
' ConfirmText/ResultText - 確認ダイアログと完了報告の文面(純文字列組立)。
' ----------------------------------------------------------------------------
Public Function ConfirmText(ByVal needN As Long, ByVal cannotN As Long) As String
    Dim s As String
    s = "旧形式の資料が" & needN & "冊あります。仕上げると『全部教えて』の俯瞰・" & _
        "条文参照・言い換え検索が使えるようになります。ファイルの再取込は不要です" & _
        "(AIが章の数だけ動きます。目安: 1冊あたり数十秒～数分)。実行しますか?"
    If cannotN > 0 Then
        s = s & vbLf & vbLf & cannotN & "冊は形式が古いため仕上げできません" & _
            "(再取込が必要です)。"
    End If
    ConfirmText = s
End Function

Public Function ResultText(ByVal okN As Long, ByVal ngN As Long, ByVal cannotN As Long) As String
    Dim s As String
    If ngN = 0 Then
        s = okN & "冊の仕上げが完了しました。"
    Else
        s = okN & "冊の仕上げが完了しました(" & ngN & "冊は失敗。ログをご確認ください)。"
    End If
    If cannotN > 0 Then s = s & "(" & cannotN & "冊は再取込が必要です)"
    ResultText = s
End Function

' ============================================================================
' 内部ヘルパー(シートI/O)
' ============================================================================

' chunk_meta の chunk_id 集合(O(1)突合のためDictionary化)。
Private Function BuildMetaIdSet() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim ids() As String, paths() As String, refs() As String
    Dim n As Long: n = modChunkMetaStore.ReadAllMeta(ids, paths, refs)
    Dim i As Long
    For i = 0 To n - 1
        If LenB(ids(i)) > 0 Then
            If Not d.Exists(ids(i)) Then d.Add ids(i), True
        End If
    Next i
    Set BuildMetaIdSet = d
End Function

' doc_outline に行がある source 集合。
Private Function BuildOutlineSourceSet() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim srcs() As String, keys() As String, sums() As String, kws() As String
    Dim n As Long: n = modOutlineStore.ReadOutline(srcs, keys, sums, kws)
    Dim i As Long
    For i = 0 To n - 1
        If LenB(srcs(i)) > 0 Then
            If Not d.Exists(srcs(i)) Then d.Add srcs(i), True
        End If
    Next i
    Set BuildOutlineSourceSet = d
End Function

' my_knowledge を1回だけ全走査し、source単位に(チャンク数/chunk_meta突合
' 件数/breadcrumbサンプル)を集計する(source数×全走査を避けるための1パス集計)。
Private Function ScanKnowledge(ByRef metaSet As Object, ByRef outNames() As String, _
                               ByRef outChunkN() As Long, ByRef outMetaN() As Long, _
                               ByRef outSample() As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then Exit Function
    Dim lastK As Long: lastK = ws.Cells(ws.Rows.count, COL_ID).End(xlUp).row
    If lastK < 2 Then Exit Function

    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, COL_ID), ws.Cells(lastK, COL_FULLTEXT)).Value

    Dim idxOf As Object: Set idxOf = CreateObject("Scripting.Dictionary")
    ' 2026-08-06 R20H FA-13: Dictionaryは既定でbinary(大小文字を区別)比較の
    ' ため、CollectSourceChunks側のStrComp(...,vbTextCompare)と食い違って
    ' いた(同じ資料名が大小違いだけで2件に割れ得る)。片方に統一する
    ' (推奨どおりvbTextCompareへ揃える。設定はキー追加より前で行うこと)。
    idxOf.CompareMode = vbTextCompare
    Dim cap As Long: cap = 32
    ReDim outNames(0 To cap - 1)
    ReDim outChunkN(0 To cap - 1)
    ReDim outMetaN(0 To cap - 1)
    ReDim outSample(0 To cap - 1)

    Dim outN As Long: outN = 0
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim src As String: src = Trim$(CStr(arr(i, COL_SOURCE)))
        If LenB(src) > 0 Then
            Dim k As Long
            If idxOf.Exists(src) Then
                k = idxOf(src)
            Else
                If outN >= cap Then
                    cap = cap * 2
                    ReDim Preserve outNames(0 To cap - 1)
                    ReDim Preserve outChunkN(0 To cap - 1)
                    ReDim Preserve outMetaN(0 To cap - 1)
                    ReDim Preserve outSample(0 To cap - 1)
                End If
                k = outN
                idxOf.Add src, k
                outNames(k) = src
                outSample(k) = CStr(arr(i, COL_FULLTEXT))
                outN = outN + 1
            End If
            outChunkN(k) = outChunkN(k) + 1
            Dim cid As String: cid = Trim$(CStr(arr(i, COL_ID)))
            If LenB(cid) > 0 Then
                If metaSet.Exists(cid) Then outMetaN(k) = outMetaN(k) + 1
            End If
        End If
    Next i
    ScanKnowledge = outN
End Function

' 1資料ぶんの (chunk_id, full_text) を文書順(シート順)で集める
' (modOutlineBuild.CollectSourceRowsと同じ「id/source列だけ先に絞る」流儀に
' full_text列を足したもの。全列を読み直さない=本棚全体の本文を二重に
' メモリへ載せない)。
Private Function CollectSourceChunks(ByVal sourceName As String, ByRef outIds() As String, _
                                     ByRef outTexts() As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then Exit Function
    Dim lastK As Long: lastK = ws.Cells(ws.Rows.count, COL_ID).End(xlUp).row
    If lastK < 2 Then Exit Function

    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, COL_ID), ws.Cells(lastK, COL_FULLTEXT)).Value

    ReDim outIds(1 To lastK - 1)
    ReDim outTexts(1 To lastK - 1)
    Dim n As Long
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(Trim$(CStr(arr(i, COL_SOURCE))), sourceName, vbTextCompare) = 0 Then
            Dim id As String: id = Trim$(CStr(arr(i, COL_ID)))
            If LenB(id) > 0 Then
                n = n + 1
                outIds(n) = id
                outTexts(n) = CStr(arr(i, COL_FULLTEXT))
            End If
        End If
    Next i
    CollectSourceChunks = n
End Function
