Attribute VB_Name = "modChunkMetaStore"
Option Explicit

' ========================================
' modChunkMetaStore - chunk_meta シート(chunk_id/section_path/refs_out)の
' Ensure(自己修復)+バッチ書込み+全行読み。
'
' 2026-08-05(R17波0): R17設計書§3 Phase1(構造メタデータ+参照エッジ)が使う
' 受け皿シートの器のみを先行整備する回。パース処理(section_path/refs_out
' の抽出ロジック本体)と、取込フロー(modShelf.IngestFile)からの実呼び出しは
' まだ無い(次波が配線する)。
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

' ReadAllMeta - 全行をids/paths/refsの並行配列(0 To n-1)へ出力し、件数を返す。
'   シートが無い/1行も無いときは0を返し、出力配列は(0 To 0)の空を保証する
'   (SourceListと同じ「呼び出し側は毎回ReDimされた配列を受け取れる」規約)。
Public Function ReadAllMeta(ByRef outIds() As String, ByRef outPaths() As String, _
                            ByRef outRefs() As String) As Long
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_CHUNK_META)
    If ws Is Nothing Then
        ReDim outIds(0 To 0)
        ReDim outPaths(0 To 0)
        ReDim outRefs(0 To 0)
        ReadAllMeta = 0
        Exit Function
    End If

    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, COL_ID).End(xlUp).row
    If lastR < 2 Then
        ReDim outIds(0 To 0)
        ReDim outPaths(0 To 0)
        ReDim outRefs(0 To 0)
        ReadAllMeta = 0
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
    ReadAllMeta = n
End Function
