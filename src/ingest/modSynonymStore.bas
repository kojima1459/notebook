Attribute VB_Name = "modSynonymStore"
Option Explicit

' ============================================================================
' modSynonymStore - 用語の表記ゆれ辞書(R17 Phase3・設計書§3 Phase3)
' ----------------------------------------------------------------------------
' なぜ要るのか:
'   規程・マニュアルの用語は「回収」「リコール」のように、同じ意味でも
'   資料と質問で書き方が割れる。dense/sparse のどちらも表記が違えば当たり
'   にくく、「回収の保険は?」のような質問がリコールの資料を拾えない事故に
'   なる(docs/45 項目34)。取込の最後に用語一覧をAIへ1回渡して表記ゆれの
'   グループを作り、synonyms シート(term, canonical)へためる。質問時は
'   このシートを読み、一致した語の同義語を質問文へ追記してから検索する。
'
' 設計判断:
'   ・synonymsシートは列を持つ表(term/canonical)だが、質問側(modAskRetrieve)
'     へ渡す形は "term>canonical|term>canonical" の1文字列(ReadMapCsv)。
'     modSparse/modRagParse はPURE_LOGIC_MODULES(Excelオブジェクト禁止)なので
'     シートを直接読めない。文字列化した地図を渡す構成で純度を保つ
'     (調査agent7 §3.3 の結論どおり)。
'   ・「既存termは上書き」: WriteSynonymRowsは末尾へ追記するだけの単純な
'     バッチ書込みで、上書きの判断は持たない。BuildSynonymsForがReadMapCsvで
'     既存分を読み、新規termと重なる行を除いてからRemoveAll+全件書き直しで
'     実現する(1つの表に「上書き」と「追記」の2種類の書き方を混在させない)。
'   ・用語候補は当該資料の chunk_meta(section_path)・doc_outline(keywords)・
'     my_knowledge(keywords列)から集める(重複排除・最大200語)。本文全体を
'     渡すと1回のプロンプトに載らない・要約されていない生の語のほうが
'     ノイズが多いため、既に足切りされた「短い語の集合」だけを渡す。
'   ・失敗の扱いはmodOutlineStore/modChunkMetaStoreと同じ線引き: この辞書は
'     検索精度の上積みであって、無くても本棚も回答も従来どおり動く。
'     取込は止めず usage_log に "synonyms_fail" を1行残す(憲章§4-1)。
'   ・ゲート(config graph_synonyms・既定on)はこの層に閉じる。呼び出し元
'     (modOutlineBuild.BuildOutlineFor)は1行のまま(modAskFocusの
'     graph_refs/GraphActiveと同じ作法)。
' ============================================================================

Private Const COL_TERM As Long = 1
Private Const COL_CANON As Long = 2
Private Const SYN_COLS As Long = 2

' 1回の名寄せバッチへ渡す用語候補の上限(設計書§3 Phase3)。
Private Const TERM_CAP As Long = 200

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' EnsureSynonymSheet - シートを1枚用意して返す(無ければ作成+ヘッダ+
'   veryHidden、既存ならそのまま)。数式インジェクション防御(text列を"@"へ
'   固定)は毎回・冪等に適用する(EnsureOutlineSheetと同じ作法)。
' ----------------------------------------------------------------------------
Public Function EnsureSynonymSheet() As Worksheet
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_SYNONYMS)
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = modAppDef.SH_SYNONYMS
        ws.Cells(1, COL_TERM).Value = "term"
        ws.Cells(1, COL_CANON).Value = "canonical"
    End If
    On Error Resume Next
    ws.Visible = 2   ' xlSheetVeryHidden(modBoot へ足さないぶん、ここで毎回守る)
    ws.Columns(COL_TERM).NumberFormat = "@"
    ws.Columns(COL_CANON).NumberFormat = "@"
    On Error GoTo 0
    Set EnsureSynonymSheet = ws
    Exit Function
Fail:
    Set EnsureSynonymSheet = Nothing
End Function

' ----------------------------------------------------------------------------
' WriteSynonymRows - term/canonical(並行配列・同じLBound起点・n件ぶん)を
'   末尾へ1回のRange書込みで追記する。上書き判断は持たない単純な追記のみ
'   (「既存termは上書き」は呼び出し元 BuildSynonymsFor が RemoveAll と
'   組み合わせて実現する。モジュール冒頭の設計判断)。
' ----------------------------------------------------------------------------
Public Sub WriteSynonymRows(ByRef terms() As String, ByRef canons() As String, ByVal n As Long)
    If n < 1 Then Exit Sub
    On Error GoTo Failed

    Dim ws As Worksheet: Set ws = EnsureSynonymSheet()
    If ws Is Nothing Then Exit Sub

    Dim bT As Long: bT = LBound(terms)
    Dim bC As Long: bC = LBound(canons)

    Dim arr() As Variant: ReDim arr(1 To n, 1 To SYN_COLS)
    Dim i As Long
    For i = 1 To n
        arr(i, COL_TERM) = terms(bT + i - 1)
        arr(i, COL_CANON) = canons(bC + i - 1)
    Next i

    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, COL_TERM).End(xlUp).row
    If lastR < 1 Then lastR = 1
    Dim firstRow As Long: firstRow = lastR + 1
    If firstRow < 2 Then firstRow = 2

    ws.Range(ws.Cells(firstRow, COL_TERM), ws.Cells(firstRow + n - 1, COL_CANON)).Value = arr
    Exit Sub
Failed:
    SynFail "write"
End Sub

' ----------------------------------------------------------------------------
' ReadMapCsv - 全行を "term>canonical|term>canonical" の1文字列で返す。
'   0行(シートが無い/ヘッダのみ)は空文字列。質問側(modAskRetrieve)が
'   1セッション1回だけ呼び、モジュール変数へ控えて使い回す契約
'   (取込・同期での更新は次セッションから反映=docs/10に明記)。
' ----------------------------------------------------------------------------
Public Function ReadMapCsv() As String
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_SYNONYMS)
    If ws Is Nothing Then Exit Function
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, COL_TERM).End(xlUp).row
    If lastR < 2 Then Exit Function

    Dim arr As Variant: arr = ws.Range(ws.Cells(2, COL_TERM), ws.Cells(lastR, COL_CANON)).Value
    Dim sb As String
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim t As String: t = Trim$(CStr(arr(i, 1)))
        Dim c As String: c = Trim$(CStr(arr(i, 2)))
        If LenB(t) > 0 And LenB(c) > 0 Then
            If LenB(sb) > 0 Then sb = sb & "|"
            sb = sb & t & ">" & c
        End If
    Next i
    ReadMapCsv = sb
End Function

' ----------------------------------------------------------------------------
' RemoveAll - 全行消去(再構築用)。ヘッダは残す。BuildSynonymsForが
'   「既存termは上書き」を実現するための土台(モジュール冒頭の設計判断)。
' ----------------------------------------------------------------------------
Public Sub RemoveAll()
    On Error GoTo Failed
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_SYNONYMS)
    If ws Is Nothing Then Exit Sub
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, COL_TERM).End(xlUp).row
    If lastR < 2 Then Exit Sub
    ws.Range(ws.Cells(2, COL_TERM), ws.Cells(lastR, COL_CANON)).ClearContents
    Exit Sub
Failed:
    SynFail "removeall"
End Sub

' 失敗の記録。ハンドラ稼働中は On Error Resume Next が効かない(2026-07-30
' 実機err#462 と同型)ため、別Subへ切り出して新しいエラー文脈で記録する。
Private Sub SynFail(ByVal whereAt As String)
    Dim d As String: d = "err#" & Err.Number & " " & Err.Description
    On Error Resume Next
    modLog.LogUsage "synonyms_fail", "ingest", whereAt & " " & d
    On Error GoTo 0
End Sub

' ============================================================================
' 名寄せバッチ(取込末尾から1回だけ呼ばれる)
' ============================================================================

' ----------------------------------------------------------------------------
' BuildSynonymsFor - 1資料ぶんの用語候補を集めてAIへ渡し、表記ゆれグループを
'   synonyms へ追記する(既存termは上書き)。config graph_synonyms=off /
'   用語候補0件 / LLM応答が空 のいずれかで完全に無操作。
'   呼び出しは modOutlineBuild.BuildOutlineFor の末尾から1行だけ。
' ----------------------------------------------------------------------------
Public Sub BuildSynonymsFor(ByVal sourceName As String)
    If Not GateOn() Then Exit Sub
    If LenB(Trim$(sourceName)) = 0 Then Exit Sub
    If modShelfBatch.CancelRequested() Then Exit Sub

    On Error GoTo Quiet

    Dim terms() As String
    Dim termN As Long
    termN = CollectTermCandidates(sourceName, terms)
    If termN < 1 Then Exit Sub

    Dim eff As String: eff = modConfig.GetString("quick_effort", "low")
    Dim vrb As String: vrb = modConfig.GetString("quick_verbosity", "low")
    Dim mdl As String: mdl = modConfig.GetString("quick_model", "gpt-5.5")

    Dim lat As Long
    Dim resp As String
    resp = modGateway.CallLLM(BuildSynPrompt(terms, termN), "name_dedup", eff, vrb, mdl, lat)

    Dim newTerms() As String, newCanons() As String, newN As Long
    newN = modRagParse.ParseSynResp(resp, newTerms, newCanons)
    If newN < 1 Then Exit Sub

    Dim wroteN As Long
    wroteN = MergeAndSave(newTerms, newCanons, newN)

    ' 書けた件数をそのまま残す(R17H FA-1): 「AIが何グループ返したか」ではなく
    ' 「シートに何行入ったか」を出す。旧実装は前者を出しており、保存が0行でも
    ' groups=N が並んで【無音の失敗】に見えなかった(A-H1)。
    On Error Resume Next
    modLog.LogUsage "synonyms_built", "ingest", _
        "source=" & sourceName & " groups=" & wroteN & " new=" & newN
    On Error GoTo 0
    Exit Sub

Quiet:
    ' ハンドラ稼働中は On Error Resume Next が効かないので、必ず Resume で
    ' 抜けてから記録する(modOutlineBuildと同型の作法)。
    Resume QuietDone
QuietDone:
    Err.Clear
    On Error Resume Next
    modLog.LogUsage "synonyms_fail", "ingest", "build source=" & sourceName
    On Error GoTo 0
End Sub

' config graph_synonyms(既定on)。off のときだけ切る(graph_outlineと同じ作法)。
Private Function GateOn() As Boolean
    GateOn = (LCase$(Trim$(modConfig.GetString("graph_synonyms", "on"))) <> "off")
End Function

' 新規グループを既存の地図(ReadMapCsv)へ合流させ、term が重複する古い行を
' 落としてから全件を書き直す(=既存termは上書き)。RemoveAllを一度だけ通す
' ことで、追記のたびにシートが増殖するのを避ける。戻り値=書いた行数。
'
' 2026-08-05(R17H FA-1 / A-H1): マージ規則そのものは純関数
' modRagParse.MergeSynPairs が唯一の持ち主で、ここは「読む→渡す→書く」だけの
' 薄い層になった。旧実装はここに規則を持ち、ParseSynResp が返す【0始まり】の
' 配列を 1〜n で読んでいたため、実データでは毎回「添字が範囲外」で
' BuildSynonymsFor のハンドラに握り潰され、synonyms は永久に0行だった。
' 配列の起点に依存しない形へ移し、LO実行テストのゴールデンで固定する。
Private Function MergeAndSave(ByRef newTerms() As String, ByRef newCanons() As String, _
                              ByVal newN As Long) As Long
    Dim mergedCsv As String
    mergedCsv = modRagParse.MergeSynPairs(ReadMapCsv(), PairsToCsv(newTerms, newCanons, newN))
    If LenB(mergedCsv) = 0 Then Exit Function

    Dim terms() As String, canons() As String, n As Long
    n = modRagParse.ParseSynResp("<syn>" & mergedCsv & "</syn>", terms, canons)
    If n < 1 Then Exit Function

    RemoveAll
    WriteSynonymRows terms, canons, n      ' LBound起点で読む(0始まりでも動く)
    MergeAndSave = n
End Function

' 並行配列(LBound起点・n件)を "term>canonical|…" の1文字列へ。
Private Function PairsToCsv(ByRef terms() As String, ByRef canons() As String, _
                            ByVal n As Long) As String
    If n < 1 Then Exit Function
    Dim bT As Long: bT = LBound(terms)
    Dim bC As Long: bC = LBound(canons)
    Dim sb As String
    Dim i As Long
    For i = 0 To n - 1
        If LenB(sb) > 0 Then sb = sb & "|"
        sb = sb & terms(bT + i) & ">" & canons(bC + i)
    Next i
    PairsToCsv = sb
End Function

' 名寄せプロンプト(Private。modPromptsは残321字で1本も入らない=Phase2と
' 同じ容量裁定)。出力契約はmodRagParse.ParseSynRespが読むタグ形式。
Private Function BuildSynPrompt(ByRef terms() As String, ByVal n As Long) As String
    Dim sb As String
    sb = "次は同じ資料から集めた用語の一覧です。表記のゆれ(同じ意味なのに書き方" & _
         "だけが違うもの。例: 「回収」と「リコール」)をグループにまとめ、" & _
         "それぞれのグループの正規形(検索で使うべき代表表記)を1つ決めてください。" & vbLf
    sb = sb & "・意味が違う語は絶対にまとめないこと(まとめすぎるより、まとめなくてよい)。" & vbLf
    sb = sb & "・表記ゆれの無い語は出力しない。" & vbLf & vbLf
    sb = sb & "## 用語一覧" & vbLf & Join(terms, "|") & vbLf & vbLf
    sb = sb & "## 出力形式(この形式のみで出力。説明文・前置きは一切禁止)" & vbLf
    sb = sb & "<syn>表記>正規形|表記>正規形</syn>" & vbLf
    BuildSynPrompt = sb
End Function

' ----------------------------------------------------------------------------
' CollectTermCandidates - 当該資料の chunk_meta(section_path)・
'   doc_outline(keywords)・my_knowledge(keywords列)から用語候補を集める
'   (重複排除・最大200語・モジュール冒頭の設計判断)。
' ----------------------------------------------------------------------------
Private Function CollectTermCandidates(ByVal sourceName As String, ByRef outTerms() As String) As Long
    Dim box As String: box = vbLf
    Dim cnt As Long

    ' 1) my_knowledge.keywords(カンマ区切り。modEnrichが書く形式)
    Dim ids() As String, kwCsv() As String, idN As Long
    idN = CollectFromKnowledge(sourceName, ids, kwCsv)
    Dim i As Long
    For i = 1 To idN
        Dim kwParts() As String: kwParts = Split(kwCsv(i), ",")
        AddTerms kwParts, box, cnt
    Next i

    ' 2) chunk_meta.section_path(この資料のチャンクだけ。">"区切り)
    If idN > 0 Then
        Dim idBox As String: idBox = vbLf
        For i = 1 To idN
            idBox = idBox & ids(i) & vbLf
        Next i
        Dim mIds() As String, mPaths() As String, mRefs() As String, metaN As Long
        metaN = modChunkMetaStore.ReadAllMeta(mIds, mPaths, mRefs)
        For i = 0 To metaN - 1
            If LenB(mPaths(i)) > 0 Then
                If InStr(1, idBox, vbLf & mIds(i) & vbLf, vbBinaryCompare) > 0 Then
                    Dim pathParts() As String: pathParts = Split(mPaths(i), ">")
                    AddTerms pathParts, box, cnt
                End If
            End If
        Next i
    End If

    ' 3) doc_outline.keywords(この資料の章だけ。"|"区切り)
    Dim oSrcs() As String, oKeys() As String, oSums() As String, oKws() As String, outN As Long
    outN = modOutlineStore.ReadOutline(oSrcs, oKeys, oSums, oKws)
    For i = 0 To outN - 1
        If StrComp(Trim$(oSrcs(i)), sourceName, vbTextCompare) = 0 Then
            Dim kwParts2() As String: kwParts2 = Split(oKws(i), "|")
            AddTerms kwParts2, box, cnt
        End If
    Next i

    If cnt = 0 Then Exit Function
    ReDim outTerms(1 To cnt)
    Dim parts() As String: parts = Split(Mid$(box, 2), vbLf)   ' 先頭のvbLfを外して分割
    For i = 1 To cnt
        outTerms(i) = parts(i - 1)
    Next i
    CollectTermCandidates = cnt
End Function

' その資料の chunk_id と keywords 列(2列だけ)を集める。本文まで全部読むと
' 資料が大きいときに重くなる(modOutlineBuild.CollectSourceRowsと同じ作法)。
Private Function CollectFromKnowledge(ByVal sourceName As String, ByRef outIds() As String, _
                                      ByRef outKw() As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then Exit Function

    Dim lastK As Long: lastK = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then Exit Function

    Dim arr As Variant: arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastK, 6)).Value
    ReDim outIds(1 To lastK - 1)
    ReDim outKw(1 To lastK - 1)
    Dim n As Long, i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(Trim$(CStr(arr(i, 2))), sourceName, vbTextCompare) = 0 Then
            n = n + 1
            outIds(n) = Trim$(CStr(arr(i, 1)))
            outKw(n) = Trim$(CStr(arr(i, 6)))
        End If
    Next i
    CollectFromKnowledge = n
End Function

' box(vbLf区切り・先頭vbLf)へ重複なく足す。TERM_CAP件で打ち切る。
' 1文字だけの断片はノイズが多いので拾わない(modSparseのkeys採用基準と同じ)。
Private Sub AddTerms(ByRef arr() As String, ByRef box As String, ByRef cnt As Long)
    Dim i As Long
    On Error Resume Next
    For i = LBound(arr) To UBound(arr)
        If cnt >= TERM_CAP Then Exit For
        Dim t As String: t = Trim$(arr(i))
        If Len(t) >= 2 Then
            If InStr(1, box, vbLf & t & vbLf, vbBinaryCompare) = 0 Then
                box = box & t & vbLf
                cnt = cnt + 1
            End If
        End If
    Next i
    On Error GoTo 0
End Sub
