Attribute VB_Name = "modChunkMetaStore"
Option Explicit

' ========================================
' modChunkMetaStore - chunk_meta シート(chunk_id/section_path/refs_out)の
' Ensure(自己修復)+バッチ書込み+全行読み。
'
' 2026-08-05(R17波0): R17設計書§3 Phase1(構造メタデータ+参照エッジ)が使う
' 受け皿シートの器を先行整備した回。
' 2026-08-05(R17波1 Phase1): 取込フロー(modShelf.IngestFile)から実際に
' 呼ばれるようになり、書込みの入口 WriteMetaFromRows と、再取込で消える
' 旧行の掃除 RemoveMetaForSource を足した。パース本体は modChunkMeta(純ロジック)。
'
' 失敗の扱い: このシートは検索精度の【上積み】であって、無くても本棚も回答も
' 従来どおり動く。したがってここでの失敗は取込本体を止めず、usage_log に
' "chunk_meta_fail" を1行だけ残して黙って諦める(憲章§4-1「無言の失敗禁止」を
' 満たしつつ、§3-5「データ保全が最優先」を壊さないための線引き)。
'
' EnsureChunkMetaSheet は modShelfStore.EnsureKnowledgeSheet と同型:
'   ・build_mybookshelf.py が焼き込むので通常は無条件で見つかる。
'   ・シートが無い(古いブックの移行後・焼き込み失敗時の保険)場合だけ
'     Worksheets.Add で作る。既存なら何もしない(冪等)。
' WriteMetaRows/ReadAllMeta は modShelfStore.SliceRows/BuildExistingHashSet
' 系と同じ「Range一括読み・一括書き」流儀(1行ずつのセル操作は数千行で
' 実機が固まる。MASTER_SPEC §12)。
' ========================================

Private Const COL_ID As Long = 1
Private Const COL_SECTION As Long = 2
Private Const COL_REFS As Long = 3
Private Const META_COLS As Long = 3

' ----------------------------------------------------------------------------
' セッションキャッシュ(2026-08-05 R17H FA-9 / A-M10)
' ----------------------------------------------------------------------------
' ReadAllMeta は chunk_meta の全行を Range 一括で読む。R17 Phase1/2 で
' 呼び出し口が増え、入念1ターンで modAskRetrieve(検索ごと)・modAskFocus
' (RefsExpand/ArticleEnsure)・modAskGlobal(章の突合)から何度も同じ全行を
' 読み直すようになった(2万行の本棚で1ターン数回×全行)。読む内容は取込・
' 削除が走らない限り1文字も変わらないので、最初の1回だけ読んで控える。
' 契約(並行配列を 0 To n-1 で返す・0行なら (0 To 0) の空)は【不変】で、
' 呼び出し側は1行も変えない。キャッシュはこの層の内側に閉じる。
' 無効化(=次の ReadAllMeta で読み直し)は書込み・削除の入口すべて:
'   WriteMetaRows(追記の実体。WriteMetaFromRows もここを通る)
'   RemoveMetaForSource(再取込で消える行の掃除)
' これ以外に chunk_meta を書き換える経路が増えたら、必ずここへ足すこと
' (足し忘れると「取り込み直したのに古い構造で検索される」が無音で起きる)。
Private mCacheOk As Boolean
Private mCacheN As Long
Private mCacheIds() As String
Private mCachePaths() As String
Private mCacheRefs() As String

' キャッシュ世代を捨てる(次の ReadAllMeta がシートから読み直す)。
Private Sub InvalidateMetaCache()
    mCacheOk = False
    mCacheN = 0
End Sub

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

' EnsureChunkMetaSheet - シートを1枚用意して返す(無ければ作成+ヘッダ+
'   veryHidden、既存ならそのまま)。数式インジェクション防御(text列を"@"へ
'   固定)は毎回・冪等に適用する(EnsureKnowledgeSheetと同じ作法)。
Public Function EnsureChunkMetaSheet() As Worksheet
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_CHUNK_META)
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = modAppDef.SH_CHUNK_META
        Dim hdr As Variant
        hdr = Array("chunk_id", "section_path", "refs_out")
        Dim i As Long
        For i = LBound(hdr) To UBound(hdr)
            ws.Cells(1, i + 1).Value = hdr(i)
        Next i
        On Error Resume Next
        ws.Visible = 2   ' xlSheetVeryHidden
        On Error GoTo 0
    End If
    On Error Resume Next
    ws.Columns(COL_SECTION).NumberFormat = "@"
    ws.Columns(COL_REFS).NumberFormat = "@"
    On Error GoTo 0
    Set EnsureChunkMetaSheet = ws
    Exit Function
Fail:
    Set EnsureChunkMetaSheet = Nothing
End Function

' WriteMetaRows - ids/paths/refs(並行配列・同じLBound起点・n件ぶん)を
'   末尾へ1回のRange書込みで追記する。呼び出し側の配列が0始まりでも
'   1始まりでも動くよう LBound から相対で読む。書けなくても取込は止めない
'   (このシートは検索精度の補助であり、無くても既存機能は従来どおり動く
'   フェイルセーフ設計=R17設計書§3)。
Public Sub WriteMetaRows(ByRef ids() As String, ByRef paths() As String, _
                         ByRef refs() As String, ByVal n As Long)
    If n < 1 Then Exit Sub
    InvalidateMetaCache            ' R17H FA-9: 書いたら世代を捨てる
    Dim ws As Worksheet: Set ws = EnsureChunkMetaSheet()
    If ws Is Nothing Then Exit Sub

    Dim baseId As Long: baseId = LBound(ids)
    Dim baseP As Long: baseP = LBound(paths)
    Dim baseR As Long: baseR = LBound(refs)

    Dim arr() As Variant: ReDim arr(1 To n, 1 To META_COLS)
    Dim i As Long
    For i = 1 To n
        arr(i, COL_ID) = ids(baseId + i - 1)
        arr(i, COL_SECTION) = paths(baseP + i - 1)
        arr(i, COL_REFS) = refs(baseR + i - 1)
    Next i

    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, COL_ID).End(xlUp).row
    If lastR < 1 Then lastR = 1
    Dim firstRow As Long: firstRow = lastR + 1
    If firstRow < 2 Then firstRow = 2

    On Error Resume Next
    ws.Range(ws.Cells(firstRow, COL_ID), ws.Cells(firstRow + n - 1, COL_REFS)).Value = arr
    On Error GoTo 0
End Sub

' WriteMetaFromRows - 取込ループが組み立てた my_knowledge の行列(srcRows)から
'   chunk_id を取り出し、並行配列 paths/refs と合わせて1回で追記する。
'   modShelf.IngestFile は残り字数が数百字しかない(憲章§4-6)ため、id配列を
'   向こうで組み立てさせず、ここで受ける。実書込みは WriteMetaRows 1回。
'   失敗しても取込は止めない(モジュール冒頭の設計判断)。
Public Sub WriteMetaFromRows(ByRef srcRows() As Variant, ByVal idCol As Long, _
                             ByRef paths() As String, ByRef refs() As String, _
                             ByVal n As Long)
    If n < 1 Then Exit Sub
    On Error GoTo Failed

    Dim ids() As String: ReDim ids(1 To n)
    Dim i As Long
    For i = 1 To n
        ids(i) = CStr(srcRows(i, idCol))
    Next i
    WriteMetaRows ids, paths, refs, n
    Exit Sub
Failed:
    MetaFail "write"
End Sub

' RemoveMetaForSource - 再取込で消える my_knowledge 行(同名source かつ
'   keepFromRow より前の行)の chunk_id を集め、chunk_meta の同じ行を落とす。
'
'   【呼ぶ順序】modShelfStore.RemoveKnowledgeAndVectorsForSource の直前。
'   後に呼ぶと、消えた行の chunk_id はもうどこにも残っていない(chunk_meta には
'   source 列が無い。列を足すと配布済み本棚の移行が要るため、my_knowledge 側の
'   事実から引く設計にした)。引数は RemoveKnowledgeAndVectorsForSource と
'   同じ (sourceName, keepFromRow) で、選ぶ行がずれないようにしてある。
'
'   掃除に失敗した場合に残るのは「my_knowledge に対応する行が無い chunk_meta の
'   行」だけで、検索側は chunk_id を引けずにその行を無視する(害は無く、
'   次の再取込でまた掃除を試みる)。
Public Sub RemoveMetaForSource(ByVal sourceName As String, ByVal keepFromRow As Long)
    InvalidateMetaCache            ' R17H FA-9: 消したら世代を捨てる
    On Error GoTo Failed

    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_CHUNK_META)
    If ws Is Nothing Then Exit Sub
    Dim lastM As Long: lastM = ws.Cells(ws.Rows.count, COL_ID).End(xlUp).row
    If lastM < 2 Then Exit Sub

    Dim wsK As Worksheet: Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If wsK Is Nothing Then Exit Sub
    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then Exit Sub

    ' 1) 消える行の chunk_id を vbLf 区切りの箱にする(1列目=chunk_id・2列目=source)。
    Dim kArr As Variant: kArr = wsK.Range(wsK.Cells(2, 1), wsK.Cells(lastK, 2)).Value
    Dim box As String
    Dim i As Long
    For i = LBound(kArr, 1) To UBound(kArr, 1)
        If StrComp(CStr(kArr(i, 2)), sourceName, vbTextCompare) = 0 Then
            If keepFromRow < 2 Or (i + 1) < keepFromRow Then
                box = box & vbLf & CStr(kArr(i, 1))
            End If
        End If
    Next i
    If LenB(box) = 0 Then Exit Sub
    box = box & vbLf

    ' 2) 残す行だけを詰め直して1回で書き戻す(1行ずつのDeleteは数千行で固まる)。
    Dim mArr As Variant
    mArr = ws.Range(ws.Cells(2, COL_ID), ws.Cells(lastM, COL_REFS)).Value
    Dim nRows As Long: nRows = lastM - 1
    Dim keepArr() As Variant: ReDim keepArr(1 To nRows, 1 To META_COLS)
    Dim k As Long: k = 0
    Dim c As Long
    For i = 1 To nRows
        If InStr(1, box, vbLf & CStr(mArr(i, COL_ID)) & vbLf, vbBinaryCompare) = 0 Then
            k = k + 1
            For c = 1 To META_COLS
                keepArr(k, c) = mArr(i, c)
            Next c
        End If
    Next i
    If k = nRows Then Exit Sub          ' 落ちる行が1つも無い

    ws.Range(ws.Cells(2, COL_ID), ws.Cells(lastM, COL_REFS)).ClearContents
    If k > 0 Then
        ws.Range(ws.Cells(2, COL_ID), ws.Cells(1 + k, COL_REFS)).Value = CompactMeta(keepArr, k)
    End If
    Exit Sub
Failed:
    MetaFail "remove"
End Sub

' 先頭k行だけの配列を作る(書き戻す範囲と配列の形をぴったり合わせる)。
Private Function CompactMeta(ByRef src() As Variant, ByVal k As Long) As Variant
    Dim outArr() As Variant: ReDim outArr(1 To k, 1 To META_COLS)
    Dim r As Long, c As Long
    For r = 1 To k
        For c = 1 To META_COLS
            outArr(r, c) = src(r, c)
        Next c
    Next r
    CompactMeta = outArr
End Function

' 失敗の記録。ハンドラ稼働中は On Error Resume Next が効かない(2026-07-30
' 実機err#462 と同型)ため、別Subへ切り出して新しいエラー文脈で記録する。
' Err はここへ来た時点の値をまず控える(On Error Resume Next 自体がErrを消す)。
Private Sub MetaFail(ByVal whereAt As String)
    Dim d As String: d = "err#" & Err.Number & " " & Err.Description
    On Error Resume Next
    modLog.LogUsage "chunk_meta_fail", "ingest", whereAt & " " & d
    On Error GoTo 0
End Sub

' ReadAllMeta - 全行をids/paths/refsの並行配列(0 To n-1)へ出力し、件数を返す。
'   シートが無い/1行も無いときは0を返し、出力配列は(0 To 0)の空を保証する
'   (SourceListと同じ「呼び出し側は毎回ReDimされた配列を受け取れる」規約)。
'   2026-08-05(R17H FA-9): 1セッション中は最初の1回だけシートを読み、以降は
'   控えた配列を複製して返す。契約も戻り値も従来と1文字も変わらない
'   (呼び出し側が配列を書き換えても、次の呼び出しへ伝わらない=毎回コピー)。
Public Function ReadAllMeta(ByRef outIds() As String, ByRef outPaths() As String, _
                            ByRef outRefs() As String) As Long
    If mCacheOk Then
        outIds = mCacheIds
        outPaths = mCachePaths
        outRefs = mCacheRefs
        ReadAllMeta = mCacheN
        Exit Function
    End If

    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_CHUNK_META)
    If ws Is Nothing Then
        ReadAllMeta = KeepEmpty(outIds, outPaths, outRefs)
        Exit Function
    End If

    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, COL_ID).End(xlUp).row
    If lastR < 2 Then
        ReadAllMeta = KeepEmpty(outIds, outPaths, outRefs)
        Exit Function
    End If

    Dim arr As Variant: arr = ws.Range(ws.Cells(2, COL_ID), ws.Cells(lastR, COL_REFS)).Value
    Dim n As Long: n = lastR - 1
    ReDim outIds(0 To n - 1)
    ReDim outPaths(0 To n - 1)
    ReDim outRefs(0 To n - 1)

    Dim i As Long
    For i = 1 To n
        outIds(i - 1) = CStr(arr(i, COL_ID))
        outPaths(i - 1) = CStr(arr(i, COL_SECTION))
        outRefs(i - 1) = CStr(arr(i, COL_REFS))
    Next i

    mCacheIds = outIds
    mCachePaths = outPaths
    mCacheRefs = outRefs
    mCacheN = n
    mCacheOk = True
    ReadAllMeta = n
End Function

' 0行の答え(空配列3本+件数0)。この形も控えるので、chunk_meta が無い本棚では
' 質問のたびにシートを探し直すことも無くなる。
Private Function KeepEmpty(ByRef outIds() As String, ByRef outPaths() As String, _
                           ByRef outRefs() As String) As Long
    ReDim outIds(0 To 0)
    ReDim outPaths(0 To 0)
    ReDim outRefs(0 To 0)
    mCacheIds = outIds
    mCachePaths = outPaths
    mCacheRefs = outRefs
    mCacheN = 0
    mCacheOk = True
    KeepEmpty = 0
End Function
