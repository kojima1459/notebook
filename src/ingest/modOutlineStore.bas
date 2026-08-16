Attribute VB_Name = "modOutlineStore"
Option Explicit

' ========================================
' modOutlineStore - doc_outline シート(source/section_key/summary/keywords/
' chunk_n)の Ensure(自己修復)+バッチ書込み+全行読み+資料単位の掃除。
'
' 2026-08-05(R17 Phase2): 章単位要約(疑似グローバル検索)の受け皿。
' 作りは modChunkMetaStore と同型で、違いは「1資料に数十行」しか入らない
' 小さな表であること(chunk_meta はチャンク数ぶんの行が入る)。
'
' 失敗の扱い: このシートは俯瞰質問(「〜を全部教えて」)への上積みであって、
' 無くても本棚も回答も従来どおり動く。したがってここでの失敗は取込本体を
' 止めず、usage_log に "outline_fail" を1行だけ残して黙って諦める
' (憲章§4-1「無言の失敗禁止」を満たしつつ、§3-5「データ保全が最優先」を
' 壊さないための線引き。chunk_meta と同じ判断)。
'
' EnsureOutlineSheet は modChunkMetaStore.EnsureChunkMetaSheet と同型:
'   ・build_mybookshelf.py が焼き込むので通常は無条件で見つかる。
'   ・シートが無い(古いブックの移行後・焼き込み失敗時の保険)場合だけ
'     Worksheets.Add で作る。既存なら何もしない(冪等)。
'   ・veryHidden はここでも毎回・冪等に自己設定する。modBoot.HideInternalSheets
'     へは足していない(残8字。憲章§4-6)ので、隠すことの責任はビルドの
'     焼き込みとこの1行が持つ。
' WriteOutlineRows/ReadOutline は「Range一括読み・一括書き」流儀
' (1行ずつのセル操作は数千行で実機が固まる。MASTER_SPEC §12)。
' ========================================

Private Const COL_SRC As Long = 1
Private Const COL_KEY As Long = 2
Private Const COL_SUM As Long = 3
Private Const COL_KW As Long = 4
Private Const COL_N As Long = 5
' logic_ver(2026-08-07 R21-3 E2): 章検出ロジックの世代キー
' (modOutlineBuild.OUTLINE_LOGIC_VER)。既存ブック(5列時代)への見出し追加は
' EnsureOutlineSheetが毎回・冪等に行う(modShelfStore.EnsureKnowledgeSheetの
' norm_text追加と同じ移行手法)。旧行(この列が空)は Val("")=0 として読める
' ので、読み手(modBackfill)は移行処理なしでそのまま「旧世代」と判定できる。
Private Const COL_VER As Long = 6
Private Const OUTLINE_COLS As Long = 6

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

' EnsureOutlineSheet - シートを1枚用意して返す(無ければ作成+ヘッダ+
'   veryHidden、既存ならそのまま)。数式インジェクション防御(text列を"@"へ
'   固定)は毎回・冪等に適用する(EnsureChunkMetaSheet と同じ作法)。
Public Function EnsureOutlineSheet() As Worksheet
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_DOC_OUTLINE)
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = modAppDef.SH_DOC_OUTLINE
        Dim hdr As Variant
        hdr = Array("source", "section_key", "summary", "keywords", "chunk_n", "logic_ver")
        Dim i As Long
        For i = LBound(hdr) To UBound(hdr)
            ws.Cells(1, i + 1).Value = hdr(i)
        Next i
    End If
    ' 既存ブック(5列時代)への見出し追加も毎回・冪等に(R21-3 E2)。空の行は
    ' 「世代0(未記録)」として読むので移行処理は要らない(modShelfStoreと同型)。
    On Error Resume Next
    If LenB(Trim$(CStr(ws.Cells(1, COL_VER).Value))) = 0 Then
        ws.Cells(1, COL_VER).Value = "logic_ver"
    End If
    On Error GoTo 0
    On Error Resume Next
    ws.Visible = 2   ' xlSheetVeryHidden(modBoot へ足さないぶん、ここで毎回守る)
    ws.Columns(COL_SRC).NumberFormat = "@"
    ws.Columns(COL_KEY).NumberFormat = "@"
    ws.Columns(COL_SUM).NumberFormat = "@"
    ws.Columns(COL_KW).NumberFormat = "@"
    On Error GoTo 0
    Set EnsureOutlineSheet = ws
    Exit Function
Fail:
    Set EnsureOutlineSheet = Nothing
End Function

' ----------------------------------------------------------------------------
' WriteOutlineRows - 章の要約 n 件を末尾へ1回のRange書込みで追記する。
'   srcs/keys/sums/kws/ns は並行配列(同じLBound起点・n件ぶん)。呼び出し側の
'   配列が0始まりでも1始まりでも動くよう LBound から相対で読む。
'   verNum: この一括書込み全行に記録する世代キー(呼び出し元=
'   modOutlineBuild.BuildOutlineFor が OutlineVerFor の結果を渡す。1資料の
'   章要約は必ず同じ呼び出しで一括作成されるため、行ごとに変える必要は無い)。
'   2026-08-16(R33波3 W3-4): 完走しなかった呼び出しは verNum=0 で来る。
'   0 は読み手(modBackfill)の後方互換読み Val("")=0 と同じ「未仕上げ」で、
'   その資料は⚡資料の仕上げの候補として再提案される。ここは渡された値を
'   そのまま刻むだけ=完走かどうかの判断は呼び出し元の責務(この層は
'   「どのロジックか」しか知らない)。
'   書けなくても取込は止めない(このシートは上積み=モジュール冒頭の判断)。
' ----------------------------------------------------------------------------
Public Sub WriteOutlineRows(ByRef srcs() As String, ByRef keys() As String, _
                            ByRef sums() As String, ByRef kws() As String, _
                            ByRef ns() As Long, ByVal n As Long, _
                            ByVal verNum As Long)
    If n < 1 Then Exit Sub
    On Error GoTo Failed

    Dim ws As Worksheet: Set ws = EnsureOutlineSheet()
    If ws Is Nothing Then Exit Sub

    Dim bS As Long: bS = LBound(srcs)
    Dim bK As Long: bK = LBound(keys)
    Dim bM As Long: bM = LBound(sums)
    Dim bW As Long: bW = LBound(kws)
    Dim bN As Long: bN = LBound(ns)

    Dim arr() As Variant: ReDim arr(1 To n, 1 To OUTLINE_COLS)
    Dim i As Long
    For i = 1 To n
        arr(i, COL_SRC) = srcs(bS + i - 1)
        arr(i, COL_KEY) = keys(bK + i - 1)
        arr(i, COL_SUM) = sums(bM + i - 1)
        arr(i, COL_KW) = kws(bW + i - 1)
        arr(i, COL_N) = ns(bN + i - 1)
        arr(i, COL_VER) = verNum
    Next i

    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, COL_SRC).End(xlUp).row
    If lastR < 1 Then lastR = 1
    Dim firstRow As Long: firstRow = lastR + 1
    If firstRow < 2 Then firstRow = 2

    ws.Range(ws.Cells(firstRow, COL_SRC), ws.Cells(firstRow + n - 1, COL_VER)).Value = arr
    Exit Sub
Failed:
    OutlineFail "write"
End Sub

' ----------------------------------------------------------------------------
' ReadOutline - 全行を並行配列(0 To n-1)へ出力し、件数を返す。
'   シートが無い/1行も無いときは0を返し、出力配列は(0 To 0)の空を保証する
'   (ReadAllMeta と同じ「呼び出し側は毎回ReDimされた配列を受け取れる」規約)。
'   chunk_n は読まない: 俯瞰の呼び出し側(modAskGlobal)が使うのは
'   source/section_key/summary/keywords の4つだけで、chunk_n は
'   「その章が何チャンクだったか」を人が後から見るための記録に留める。
' ----------------------------------------------------------------------------
Public Function ReadOutline(ByRef outSrcs() As String, ByRef outKeys() As String, _
                            ByRef outSums() As String, ByRef outKws() As String) As Long
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_DOC_OUTLINE)
    Dim lastR As Long
    If Not ws Is Nothing Then lastR = ws.Cells(ws.Rows.count, COL_SRC).End(xlUp).row

    If ws Is Nothing Or lastR < 2 Then
        ReDim outSrcs(0 To 0)
        ReDim outKeys(0 To 0)
        ReDim outSums(0 To 0)
        ReDim outKws(0 To 0)
        ReadOutline = 0
        Exit Function
    End If

    Dim arr As Variant: arr = ws.Range(ws.Cells(2, COL_SRC), ws.Cells(lastR, COL_KW)).Value
    Dim n As Long: n = lastR - 1
    ReDim outSrcs(0 To n - 1)
    ReDim outKeys(0 To n - 1)
    ReDim outSums(0 To n - 1)
    ReDim outKws(0 To n - 1)

    Dim i As Long
    For i = 1 To n
        outSrcs(i - 1) = CStr(arr(i, COL_SRC))
        outKeys(i - 1) = CStr(arr(i, COL_KEY))
        outSums(i - 1) = CStr(arr(i, COL_SUM))
        outKws(i - 1) = CStr(arr(i, COL_KW))
    Next i
    ReadOutline = n
End Function

' ----------------------------------------------------------------------------
' ReadOutlineVersions - 資料名とlogic_verだけの軽量読み(R21-3 E2)。
'   modBackfill.DetectLegacyDocsが「この資料のdoc_outlineは現行世代で
'   仕上げ済みか」を判定するための専用窓口(ReadOutlineの4列読みを流用せず
'   2列だけ読むのは modOutlineBuild.CollectSourceRows と同じ「本棚全体の
'   本文までは読まない」流儀)。旧世代の行(logic_ver列が空)はVal("")=0で
'   返る(後方互換: 旧doc_outlineを読んでもエラーにならない)。
' ----------------------------------------------------------------------------
Public Function ReadOutlineVersions(ByRef outSrcs() As String, ByRef outVers() As Long) As Long
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_DOC_OUTLINE)
    Dim lastR As Long
    If Not ws Is Nothing Then lastR = ws.Cells(ws.Rows.count, COL_SRC).End(xlUp).row

    If ws Is Nothing Or lastR < 2 Then
        ReDim outSrcs(0 To 0)
        ReDim outVers(0 To 0)
        ReadOutlineVersions = 0
        Exit Function
    End If

    Dim arr As Variant: arr = ws.Range(ws.Cells(2, COL_SRC), ws.Cells(lastR, COL_VER)).Value
    Dim n As Long: n = lastR - 1
    ReDim outSrcs(0 To n - 1)
    ReDim outVers(0 To n - 1)

    Dim i As Long
    For i = 1 To n
        outSrcs(i - 1) = CStr(arr(i, COL_SRC))
        outVers(i - 1) = CLng(Val(CStr(arr(i, COL_VER))))
    Next i
    ReadOutlineVersions = n
End Function

' ----------------------------------------------------------------------------
' RemoveOutlineForSource - その資料の章要約を全部落とす。
'   chunk_meta の掃除(RemoveMetaForSource)と違い、doc_outline は source 列を
'   自分で持っているので my_knowledge を引く必要が無い。したがって
'   RemoveKnowledgeAndVectorsForSource との前後関係の制約も無い(どちらでもよい)。
'   呼ばれる場所: 再取込(modShelf.IngestFile 手順7.5)/資料削除(DeleteSource)/
'   章要約の作り直し(modOutlineBuild.BuildOutlineFor の書込み直前)。
'   1行ずつのDeleteは数千行で固まるので、残す行を詰め直して1回で書き戻す。
' ----------------------------------------------------------------------------
Public Sub RemoveOutlineForSource(ByVal sourceName As String)
    On Error GoTo Failed
    If LenB(Trim$(sourceName)) = 0 Then Exit Sub

    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_DOC_OUTLINE)
    If ws Is Nothing Then Exit Sub
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, COL_SRC).End(xlUp).row
    If lastR < 2 Then Exit Sub

    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, COL_SRC), ws.Cells(lastR, COL_VER)).Value
    Dim nRows As Long: nRows = lastR - 1

    Dim keepArr() As Variant: ReDim keepArr(1 To nRows, 1 To OUTLINE_COLS)
    Dim k As Long: k = 0
    Dim i As Long, c As Long
    For i = 1 To nRows
        If StrComp(Trim$(CStr(arr(i, COL_SRC))), sourceName, vbTextCompare) <> 0 Then
            k = k + 1
            For c = 1 To OUTLINE_COLS
                keepArr(k, c) = arr(i, c)
            Next c
        End If
    Next i
    If k = nRows Then Exit Sub          ' 落ちる行が1つも無い

    ws.Range(ws.Cells(2, COL_SRC), ws.Cells(lastR, COL_VER)).ClearContents
    If k > 0 Then
        ws.Range(ws.Cells(2, COL_SRC), ws.Cells(1 + k, COL_VER)).Value = CompactOutline(keepArr, k)
    End If
    Exit Sub
Failed:
    OutlineFail "remove"
End Sub

' 先頭k行だけの配列を作る(書き戻す範囲と配列の形をぴったり合わせる)。
Private Function CompactOutline(ByRef src() As Variant, ByVal k As Long) As Variant
    Dim outArr() As Variant: ReDim outArr(1 To k, 1 To OUTLINE_COLS)
    Dim r As Long, c As Long
    For r = 1 To k
        For c = 1 To OUTLINE_COLS
            outArr(r, c) = src(r, c)
        Next c
    Next r
    CompactOutline = outArr
End Function

' 失敗の記録。ハンドラ稼働中は On Error Resume Next が効かない(2026-07-30
' 実機err#462 と同型)ため、別Subへ切り出して新しいエラー文脈で記録する。
' Err はここへ来た時点の値をまず控える(On Error Resume Next 自体がErrを消す)。
Private Sub OutlineFail(ByVal whereAt As String)
    Dim d As String: d = "err#" & Err.Number & " " & Err.Description
    On Error Resume Next
    modLog.LogUsage "outline_fail", "ingest", whereAt & " " & d
    On Error GoTo 0
End Sub
