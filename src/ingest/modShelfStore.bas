Attribute VB_Name = "modShelfStore"
Option Explicit

' ========================================
' modShelfStore - 本棚シート(my_knowledge / my_vectors / my_manifest)の行操作
'
' modShelf から切り出した「シートを作る・行を探す・行を消す・配列で書き戻す」層
' (2026-07-28 レビュー I-2: modShelf が上限30,000字に対し残り13字まで逼迫)。
' 行操作は取込フロー以外(同期・失効ワイプ)からも呼ぶ共通処理。
' 全行の読み書きは「Range一括読み→配列でフィルタ→一括書き戻し」で行う
' (1行ずつ Rows().Delete すると数千行で実機が数分固まる。MASTER_SPEC §12)。
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
' my_knowledge の列(1..10)。10列目 norm_text は 2026-08-01(R12-4)で追加。
' 照合用の正規化済みテキストを取込時に1回だけ作って持つ(以前は検索のたびに
' 全チャンクを正規化し直し、20,500件では1質問あたり1,540万文字のループ)。
'   ・見出しは EnsureKnowledgeSheet が毎回・冪等に付ける(移行不要)
'   ・空欄は「未計算」とみなし、検索側がその行だけ計算して書き戻す(遅延)
'   ・行の詰め直しは必ず全列(KNOWLEDGE_COLS)を運ぶ。9列のまま詰めると
'     norm_text だけが別の行に残り【別チャンクの照合テキストで採点する】事故。
Public Const COL_K_NORM As Long = 10
Private Const KNOWLEDGE_COLS As Long = 10

' my_manifest の列(1..10)。10列目 fail_count は 2026-08-01(R12-3-3)で追加。
' 恒久失敗ファイルのバックオフ: 同期は status="failed" を毎回 replace 判定で
' 拾い直すため、壊れたPDFが毎回の同期時間と err_log を食い続ける。連続失敗を
' 数え MAX_FAIL_STREAK 回で "failed_permanent" へ倒す。復帰の道は必ず残す:
' (a) 利用者が「資料を追加」で明示的に選び直したとき(ResetFailCountForPath)、
' (b) ファイル更新時(サイズ/更新日時の変化で DiffDecision が replace を返す)。
Private Const COL_M_STATUS As Long = 6
Private Const COL_M_FAILN As Long = 10
Private Const MANIFEST_COLS As Long = 10
Private Const MAX_FAIL_STREAK As Long = 3
' status 語彙は "done"/"pending"/"partial"/"failed"/"image_pdf"/"missing" と
' 同じく素の文字列で持つ(読み手=modShelfSync.ResolveDecision・modUIShelf は
' LOテストの都合で本モジュールを参照できないため、既存の語彙と作法を揃える)。
Private Const STATUS_FAILED_PERMANENT As String = "failed_permanent"

' カードのメモ文言(2026-08-05 R18H FA-5/FA-6)。modShelf.IngestFile から参照
' するので Public Const(Private Const は跨いで参照できずLOも通らない)。
' 台帳へ書く文言は台帳を書くこの層の持ち物(§4-5)。
'   MEMO_ALL_DUP   : 全チャンクが既存と同一ハッシュ(FA-5)。旧行を消すと資料が
'     丸ごと0件になるので消さずに温存し、消していないことを言葉で伝える。
'   MEMO_REPLACE_NG: 新行は書けたが旧行の削除に失敗(FA-6)。新旧2組が残り検索が
'     重複するので、自分で直せる唯一の手(もう一度取り込む)を必ず添える。
Public Const MEMO_ALL_DUP As String = _
    "内容が既存資料と重複しているため、既存データを保持しました"
Public Const MEMO_REPLACE_NG As String = _
    "置き換えが完了していません(もう一度取り込むと解消します)"

' norm_text(第10列)の遅延バックフィル用バッファ(検索1回ぶん・R12-4)。
Private Const BACKFILL_CELL_MAX As Long = 200   ' これ以下なら1セルずつ書く
Private mBfRow() As Long
Private mBfVal() As String
Private mBfN As Long

' シートを1枚取る。無ければ Nothing(呼び出し側が黙って諦められるように)。
' 2026-07-28: 切り出し時にこの関数だけ持ってくるのを忘れ、実機で「Sub または
' Function が定義されていません」となり資料の取込が全滅した。
Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Public Function EnsureKnowledgeSheet() As Worksheet
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = modAppDef.SH_KNOWLEDGE
        Dim hdr As Variant
        hdr = Array("chunk_id", "source", "origin", "page", "summary", "keywords", _
                    "full_text", "added_at", "embedded", "norm_text")
        Dim i As Long
        For i = LBound(hdr) To UBound(hdr)
            ws.Cells(1, i + 1).Value = hdr(i)
        Next i
        On Error Resume Next
        ws.Visible = 2   ' xlSheetVeryHidden
        On Error GoTo 0
    End If
    ' 既存ブック(9列時代)への見出し追加も毎回・冪等に。空の行は「未計算」と
    ' して読むので移行処理は要らない(R12-4)。
    On Error Resume Next
    If LenB(Trim$(CStr(ws.Cells(1, COL_K_NORM).Value))) = 0 Then
        ws.Cells(1, COL_K_NORM).Value = "norm_text"
    End If
    On Error GoTo 0
    ' 数式インジェクション防御(毎回・冪等。配布テンプレートの既存シートにも
    ' 効くよう If の外で適用): 非信頼テキスト列を"@"書式へ固定する。
    On Error Resume Next
    ws.Columns(COL_SOURCE).NumberFormat = "@"
    ws.Columns(COL_SUMMARY).NumberFormat = "@"
    ws.Columns(COL_KEYWORDS).NumberFormat = "@"
    ws.Columns(COL_FULLTEXT).NumberFormat = "@"
    ws.Columns(COL_K_NORM).NumberFormat = "@"
    On Error GoTo 0
    Set EnsureKnowledgeSheet = ws
    Exit Function
Fail:
    Set EnsureKnowledgeSheet = Nothing
End Function

Public Function EnsureManifestSheet() As Worksheet
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_MANIFEST)
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = modAppDef.SH_MANIFEST
        Dim hdr As Variant
        hdr = Array("file_path", "file_name", "modified_at", "size", "chunk_count", _
                    "status", "error_note", "ingested_at", "origin", "fail_count")
        Dim i As Long
        For i = LBound(hdr) To UBound(hdr)
            ws.Cells(1, i + 1).Value = hdr(i)
        Next i
        On Error Resume Next
        ws.Visible = 0   ' xlSheetHidden
        On Error GoTo 0
    End If
    ' 既存ブック(9列時代)への見出し追加も毎回・冪等に。空は fail_count=0。
    On Error Resume Next
    If LenB(Trim$(CStr(ws.Cells(1, COL_M_FAILN).Value))) = 0 Then
        ws.Cells(1, COL_M_FAILN).Value = "fail_count"
    End If
    On Error GoTo 0
    Set EnsureManifestSheet = ws
    Exit Function
Fail:
    Set EnsureManifestSheet = Nothing
End Function

' norm_text(第10列)の供給と遅延バックフィル(2026-08-01 R12-4)。
' 検索側(modRetrieve)は「この行の照合テキストをくれ」と言うだけ。空なら
' その行だけ計算して返し、走査の最後にまとめて書き戻す(列の都合を知るのは
' この層だけ=検索側にセル書込みを置かない)。
Public Sub ResetNormBackfill()
    mBfN = 0
    ReDim mBfRow(0 To 63)
    ReDim mBfVal(0 To 63)
End Sub

' kData は my_knowledge の1〜10列を読んだ配列、kRow はその配列内の行番号。
Public Function NormTextAt(ByRef kData As Variant, ByVal kRow As Long) As String
    Dim s As String
    s = CStr(kData(kRow, COL_K_NORM))
    If LenB(s) > 0 Then
        NormTextAt = s
        Exit Function
    End If
    s = modSparse.MatchDocText(CStr(kData(kRow, COL_SUMMARY)), CStr(kData(kRow, COL_KEYWORDS)), _
                               CStr(kData(kRow, COL_SOURCE)), CStr(kData(kRow, COL_FULLTEXT)))
    NormTextAt = s
    QueueNormBackfill kRow, s
End Function

Private Sub QueueNormBackfill(ByVal kRow As Long, ByVal s As String)
    On Error Resume Next
    ' 1セルに収まらない長さは保存しない(切り詰めると保存済みの行と未保存の行で
    ' スコアが変わる=挙動等価が壊れる)。
    If Len(s) > 32000 Then Exit Sub
    If mBfN > UBound(mBfRow) Then
        ReDim Preserve mBfRow(0 To UBound(mBfRow) * 2 + 1)
        ReDim Preserve mBfVal(0 To UBound(mBfVal) * 2 + 1)
    End If
    mBfRow(mBfN) = kRow
    mBfVal(mBfN) = s
    mBfN = mBfN + 1
    On Error GoTo 0
End Sub

' 走査で作った norm_text をまとめて書き戻す。書けなくても検索結果は不変
' (次回また計算するだけ)なので全体を OERN で包む。
Public Sub FlushNormBackfill(ByVal wsK As Worksheet, ByVal lastK As Long, ByRef kData As Variant)
    On Error Resume Next
    If mBfN < 1 Then Exit Sub
    If wsK Is Nothing Then Exit Sub
    ' 走査中のDoEventsで行が動いていたら書かない(別の行へ書く事故を避ける)。
    If wsK.Cells(wsK.Rows.count, COL_ID).End(xlUp).row <> lastK Then Exit Sub

    ' 数式インジェクション防御(R12-2と同じ作法): 書く直前に列を"@"へ固定。
    wsK.Columns(COL_K_NORM).NumberFormat = "@"

    ' 2026-08-01(R12-H-5): 書けた件数を数える。1件も書けないと「毎回計算し
    ' 直しているのに誰も気付けない」ので必ず痕跡を残す(憲章§4-1)。
    Dim wrote As Long: wrote = 0
    Dim i As Long
    If mBfN <= BACKFILL_CELL_MAX Then
        ' 数行だけなら直接書く(列まるごとの書戻しは1回でも数万セルを触る)。
        For i = 0 To mBfN - 1
            Err.Clear
            wsK.Cells(1 + mBfRow(i), COL_K_NORM).Value = mBfVal(i)
            If Err.Number = 0 Then
                kData(mBfRow(i), COL_K_NORM) = mBfVal(i)
                wrote = wrote + 1
            End If
        Next i
    Else
        ' 初回のような大量バックフィルは200行バッチで書く(modPack.
        ' WriteRowsBatched と同じ作法)。数万行×長文の1回代入は実行時エラー7に
        ' なり得るうえ、途中で失敗すると【1件も書けない】状態が固定される。
        Dim nRows As Long: nRows = UBound(kData, 1) - LBound(kData, 1) + 1
        Dim colArr() As Variant: ReDim colArr(1 To nRows, 1 To 1)
        For i = 1 To nRows
            If LenB(CStr(kData(i, COL_K_NORM))) > 0 Then colArr(i, 1) = CStr(kData(i, COL_K_NORM))
        Next i
        For i = 0 To mBfN - 1
            colArr(mBfRow(i), 1) = mBfVal(i)
        Next i

        Dim startR As Long
        For startR = 1 To nRows Step BACKFILL_CELL_MAX
            Dim batchN As Long: batchN = nRows - startR + 1
            If batchN > BACKFILL_CELL_MAX Then batchN = BACKFILL_CELL_MAX
            Dim batchArr() As Variant: ReDim batchArr(1 To batchN, 1 To 1)
            Dim b As Long
            For b = 1 To batchN
                batchArr(b, 1) = colArr(startR + b - 1, 1)
            Next b
            Err.Clear
            wsK.Range(wsK.Cells(1 + startR, COL_K_NORM), _
                      wsK.Cells(startR + batchN, COL_K_NORM)).Value = batchArr
            If Err.Number = 0 Then wrote = wrote + batchN
        Next startR
        If wrote > 0 Then
            For i = 0 To mBfN - 1
                kData(mBfRow(i), COL_K_NORM) = mBfVal(i)
            Next i
        End If
    End If

    If wrote = 0 Then
        modLog.LogUsage "normtext_backfill_failed", "", _
            "照合テキスト(norm_text)を1件も保存できませんでした。検索は動きますが、" & _
            "毎回の質問で計算し直すため遅いままです(対象" & mBfN & "件)"
    End If
    mBfN = 0
    On Error GoTo 0
End Sub

' 既存チャンクのハッシュ集合(重複排除用)。
' excludeSource (2026-07-28 レビュー H-6): このsource名の行を集合に入れない。
'   同名資料の置き換えを「抽出が成功してから消す」順(R18-2eで更に「書いてから
'   消す」)に変えたため、消す前に作る集合から自分自身を外さないと、入れ直す
'   チャンクが全部「自分との重複」と判定されて1件も入らない。
Public Function BuildExistingHashSet(ByVal wsK As Worksheet, _
                                     Optional ByVal excludeSource As String = "") As Object
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

    ' chunk_id(1列目)と source(2列目)を一括で読む。
    Dim arr As Variant
    arr = wsK.Range(wsK.Cells(2, COL_ID), wsK.Cells(lastK, COL_SOURCE)).Value
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If LenB(excludeSource) = 0 Then
            AddHashFromId dict, CStr(arr(i, 1))
        ElseIf StrComp(CStr(arr(i, 2)), excludeSource, vbTextCompare) <> 0 Then
            AddHashFromId dict, CStr(arr(i, 1))
        End If
    Next i
    Set BuildExistingHashSet = dict
End Function

Public Sub AddHashFromId(ByVal dict As Object, ByVal chunkId As String)
    If LenB(chunkId) = 0 Then Exit Sub
    Dim parts() As String: parts = Split(chunkId, "::")
    If UBound(parts) - LBound(parts) + 1 < 2 Then Exit Sub
    Dim h As String: h = parts(LBound(parts) + 1)
    If LenB(h) > 0 Then
        If Not dict.Exists(h) Then dict.Add h, True
    End If
End Sub

' 指定sourceのknowledge/vector行を除去(配列読み→フィルタ→書戻し)。
' keepFromRow (2026-08-05 R18-2e): このシート行番号【以降】の行は source が
'   一致していても消さない。再取込を「新行を書いてから旧行を消す」へ反転した
'   ため、いま書いた新行(同じ source 名)を巻き添えにしないための境界。
'   0(既定)なら従来どおり全ての一致行が対象=既存の呼び出し元は無改修。
'   arr は2行目起点なので、配列の i 行目のシート行番号は i+1。
Public Sub RemoveKnowledgeAndVectorsForSource(ByVal sourceName As String, _
                                              Optional ByVal keepFromRow As Long = 0)
    Dim wsK As Worksheet: Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If wsK Is Nothing Then Exit Sub
    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then Exit Sub

    Dim arr As Variant: arr = wsK.Range(wsK.Cells(2, 1), wsK.Cells(lastK, KNOWLEDGE_COLS)).Value
    Dim nRows As Long: nRows = UBound(arr, 1) - LBound(arr, 1) + 1

    Dim removedIds As Object: Set removedIds = CreateObject("Scripting.Dictionary")
    Dim survivors() As Variant: ReDim survivors(1 To nRows, 1 To KNOWLEDGE_COLS)
    Dim survivorCount As Long: survivorCount = 0

    Dim i As Long, c As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(CStr(arr(i, COL_SOURCE)), sourceName, vbTextCompare) = 0 _
           And (keepFromRow < 2 Or (i + 1) < keepFromRow) Then
            Dim cid As String: cid = CStr(arr(i, COL_ID))
            If LenB(cid) > 0 Then
                If Not removedIds.Exists(cid) Then removedIds.Add cid, True
            End If
        Else
            survivorCount = survivorCount + 1
            For c = 1 To KNOWLEDGE_COLS
                survivors(survivorCount, c) = arr(i, c)
            Next c
        End If
    Next i

    If removedIds.count = 0 Then Exit Sub   ' 一致無し(何もしなくてよい)

    If survivorCount > 0 Then
        Dim writeArr As Variant: writeArr = CompactRows(survivors, survivorCount)
        wsK.Range(wsK.Cells(2, 1), wsK.Cells(1 + survivorCount, KNOWLEDGE_COLS)).Value = writeArr
    End If
    If survivorCount < nRows Then
        wsK.Range(wsK.Cells(2 + survivorCount, 1), wsK.Cells(1 + nRows, KNOWLEDGE_COLS)).ClearContents
    End If

    RemoveVectorsByIds removedIds
End Sub

' RemoveRowsByOrigin - origin列が originTag と一致する行を knowledge/vectors
'   から取り除く。戻り値=消した件数。
' 2026-07-28(レビュー C-1): 同じ処理を modChannel が自前に持ち、消す側と書く側の
' タグが食い違ったまま誰も気付かなかった。書き手(modPack)と消し手(modChannel)
' の両方がここを呼び、タグの取り扱いを1箇所に集める。比較は大小無視。
Public Function RemoveRowsByOrigin(ByVal originTag As String) As Long
    If LenB(Trim$(originTag)) = 0 Then Exit Function
    Dim wsK As Worksheet: Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If wsK Is Nothing Then Exit Function
    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then Exit Function

    Dim arr As Variant: arr = wsK.Range(wsK.Cells(2, 1), wsK.Cells(lastK, KNOWLEDGE_COLS)).Value
    Dim nRows As Long: nRows = UBound(arr, 1) - LBound(arr, 1) + 1

    Dim removedIds As Object: Set removedIds = CreateObject("Scripting.Dictionary")
    Dim survivors() As Variant: ReDim survivors(1 To nRows, 1 To KNOWLEDGE_COLS)
    Dim survivorCount As Long: survivorCount = 0

    Dim i As Long, c As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(Trim$(CStr(arr(i, COL_ORIGIN))), originTag, vbTextCompare) = 0 Then
            Dim cid As String: cid = CStr(arr(i, COL_ID))
            If LenB(cid) > 0 Then
                If Not removedIds.Exists(cid) Then removedIds.Add cid, True
            End If
        Else
            survivorCount = survivorCount + 1
            For c = 1 To KNOWLEDGE_COLS
                survivors(survivorCount, c) = arr(i, c)
            Next c
        End If
    Next i

    If survivorCount = nRows Then Exit Function   ' 一致無し

    If survivorCount > 0 Then
        Dim writeArr As Variant: writeArr = CompactRows(survivors, survivorCount)
        wsK.Range(wsK.Cells(2, 1), wsK.Cells(1 + survivorCount, KNOWLEDGE_COLS)).Value = writeArr
    End If
    wsK.Range(wsK.Cells(2 + survivorCount, 1), wsK.Cells(1 + nRows, KNOWLEDGE_COLS)).ClearContents

    RemoveVectorsByIds removedIds
    RemoveRowsByOrigin = nRows - survivorCount
    MarkRowsChanged
End Function

' origin列が originTag と一致する行数を数える(消さない)。
' 発行前の確認ダイアログで「実際に何件出るか」を出すために使う。
' 本棚の総数を見せると、他部門を購読している端末で数字が合わない。
Public Function CountRowsByOrigin(ByVal originTag As String) As Long
    On Error Resume Next
    If LenB(Trim$(originTag)) = 0 Then Exit Function
    Dim wsK As Worksheet: Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If wsK Is Nothing Then Exit Function
    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then Exit Function

    Dim arr As Variant: arr = wsK.Range(wsK.Cells(2, COL_ORIGIN), wsK.Cells(lastK, COL_ORIGIN)).Value
    Dim i As Long
    If Not IsArray(arr) Then
        If StrComp(Trim$(CStr(arr)), originTag, vbTextCompare) = 0 Then CountRowsByOrigin = 1
        Exit Function
    End If
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(Trim$(CStr(arr(i, 1))), originTag, vbTextCompare) = 0 Then
            CountRowsByOrigin = CountRowsByOrigin + 1
        End If
    Next i
    On Error GoTo 0
End Function

' origin列が "<prefix>" で始まる行を全部消す(移行用。戻り値=消した件数)。
' 例: PrefixTag="pack:" で、旧仕様のチャンネル残骸をまとめて掃除する。
Public Function RemoveRowsByOriginPrefix(ByVal prefixTag As String) As Long
    If LenB(prefixTag) = 0 Then Exit Function
    Dim wsK As Worksheet: Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If wsK Is Nothing Then Exit Function
    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then Exit Function

    Dim arr As Variant: arr = wsK.Range(wsK.Cells(2, 1), wsK.Cells(lastK, KNOWLEDGE_COLS)).Value
    Dim nRows As Long: nRows = UBound(arr, 1) - LBound(arr, 1) + 1
    Dim pfxLen As Long: pfxLen = Len(prefixTag)

    Dim removedIds As Object: Set removedIds = CreateObject("Scripting.Dictionary")
    Dim survivors() As Variant: ReDim survivors(1 To nRows, 1 To KNOWLEDGE_COLS)
    Dim survivorCount As Long: survivorCount = 0

    Dim i As Long, c As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(Left$(Trim$(CStr(arr(i, COL_ORIGIN))), pfxLen), prefixTag, vbTextCompare) = 0 Then
            Dim cid As String: cid = CStr(arr(i, COL_ID))
            If LenB(cid) > 0 Then
                If Not removedIds.Exists(cid) Then removedIds.Add cid, True
            End If
        Else
            survivorCount = survivorCount + 1
            For c = 1 To KNOWLEDGE_COLS
                survivors(survivorCount, c) = arr(i, c)
            Next c
        End If
    Next i

    If survivorCount = nRows Then Exit Function

    If survivorCount > 0 Then
        Dim writeArr2 As Variant: writeArr2 = CompactRows(survivors, survivorCount)
        wsK.Range(wsK.Cells(2, 1), wsK.Cells(1 + survivorCount, KNOWLEDGE_COLS)).Value = writeArr2
    End If
    wsK.Range(wsK.Cells(2 + survivorCount, 1), wsK.Cells(1 + nRows, KNOWLEDGE_COLS)).ClearContents

    RemoveVectorsByIds removedIds
    RemoveRowsByOriginPrefix = nRows - survivorCount
    MarkRowsChanged
End Function

' MarkRowsChanged - 「意図して減らした」ことを整合性マークへ反映(R18H FA-3)。
' 起動時の減少警告(modIntegrity.WarnAtStartup)は ui_state のマークと今の行数を
' 比べる。利用者が資料やパックを自分で削除するとマークだけが古い大きな値のまま
' 残り、「資料が消えました」の虚偽警告になっていた(A-M2)。削除の直後に最新の
' 行数へ更新すれば、ui_state はブックと一緒に保存/破棄されるので保存してもし
' なくても事実と一致する(ここではブックを保存しない=削除は保存の合図ではない)。
' modIntegrity は基盤層(core)なので ingest→core で層順は正しい。
Private Sub MarkRowsChanged()
    On Error Resume Next
    modIntegrity.RecordSaveMark
    On Error GoTo 0
End Sub

Public Sub RemoveVectorsByIds(ByVal removedIds As Object)
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

    ' R12-4: ベクトルが1本でも消えたらセッション内キャッシュは無効
    ' (「消したはずのチャンクが検索に出続ける」の構造的防止)。my_vectors から
    ' 行を落とす唯一の場所であるここで世代を進める。
    On Error Resume Next
    modVecCache.BumpGeneration
    On Error GoTo 0

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

Public Sub RemoveManifestRowForSource(ByVal sourceName As String)
    Dim wsM As Worksheet: Set wsM = GetSheet(modAppDef.SH_MANIFEST)
    If wsM Is Nothing Then Exit Sub
    Dim lastM As Long: lastM = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
    If lastM < 2 Then Exit Sub

    ' manifest は10列(R12-3-3 で fail_count 追加)。行を詰める処理は全列を
    ' 運ばないと fail_count だけが別の行に残る。
    Dim arr As Variant: arr = wsM.Range(wsM.Cells(2, 1), wsM.Cells(lastM, MANIFEST_COLS)).Value
    Dim nRows As Long: nRows = UBound(arr, 1) - LBound(arr, 1) + 1

    Dim survivors() As Variant: ReDim survivors(1 To nRows, 1 To MANIFEST_COLS)
    Dim survivorCount As Long: survivorCount = 0
    Dim i As Long, c As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(CStr(arr(i, 2)), sourceName, vbTextCompare) <> 0 Then
            survivorCount = survivorCount + 1
            For c = 1 To MANIFEST_COLS
                survivors(survivorCount, c) = arr(i, c)
            Next c
        End If
    Next i

    If survivorCount = nRows Then Exit Sub   ' 一致無し

    If survivorCount > 0 Then
        Dim writeArr() As Variant: ReDim writeArr(1 To survivorCount, 1 To MANIFEST_COLS)
        Dim r2 As Long
        For r2 = 1 To survivorCount
            For c = 1 To MANIFEST_COLS
                writeArr(r2, c) = survivors(r2, c)
            Next c
        Next r2
        wsM.Range(wsM.Cells(2, 1), wsM.Cells(1 + survivorCount, MANIFEST_COLS)).Value = writeArr
    End If
    wsM.Range(wsM.Cells(2 + survivorCount, 1), wsM.Cells(1 + nRows, MANIFEST_COLS)).ClearContents
End Sub

' ----------------------------------------------------------------------------
' chunkCount が負(-1)なら「今回は数え直していない」の意味で、既存行の
' chunk_count をそのまま残す(2026-08-05 R18-2a・実機第5報⑧)。
' ----------------------------------------------------------------------------
' 背景: modShelf.IngestFile の失敗3経路(抽出失敗/chunkN=0/Failedハンドラ)は
' 実データ(my_knowledge の行)を1行も消さないのに chunk_count だけを 0 で
' 上書きしていた。マイ本棚のカードはこの列を表示するので、中断した再取込の
' あと「127件」が「0件」に化け、利用者には資料が消えたようにしか見えない
' (実データは生存。実機第5報⑧の確定原因)。status とエラーメモの更新は
' 従来どおり行う=カードは「失敗した」と正しく言いつつ、件数だけは嘘をつかない。
' 既存行が無いとき(新規)は前値も無いので 0 になる。
Public Sub UpsertManifestRow(ByVal filePath As String, ByVal fileName As String, ByVal modifiedAt As Date, _
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

    Dim newCount As Long: newCount = chunkCount
    If chunkCount < 0 Then newCount = ChunkCountAt(wsM, r)

    ' R12-3-3: 連続失敗のバックオフ。ここが manifest への status 書込みの唯一の
    ' 入口なので、数える場所もここ1箇所にする(§4-5)。
    Dim prevFail As Long
    prevFail = FailCountAt(wsM, r)
    Dim newFail As Long: newFail = prevFail
    Dim newStatus As String: newStatus = status
    Select Case LCase$(status)
        Case "failed"
            newFail = prevFail + 1
            If newFail >= MAX_FAIL_STREAK Then
                newStatus = STATUS_FAILED_PERMANENT
                On Error Resume Next
                modLog.LogUsage "ingest_backoff", "", _
                    "連続" & newFail & "回失敗のため自動同期の対象から外しました: " & _
                    modUtil.SafeLeft(fileName, 200)
                On Error GoTo 0
            End If
        Case "done", "partial"
            newFail = 0            ' 成功で連続記録はリセット
    End Select

    wsM.Cells(r, 1).Value = filePath
    wsM.Cells(r, 2).Value = fileName
    wsM.Cells(r, 3).Value = modifiedAt
    wsM.Cells(r, 4).Value = sizeBytes
    wsM.Cells(r, 5).Value = newCount
    wsM.Cells(r, COL_M_STATUS).Value = newStatus
    wsM.Cells(r, 7).Value = modUtil.SafeLeft(errorNote, 2000)
    wsM.Cells(r, 8).Value = modUtil.NowStamp()
    wsM.Cells(r, 9).Value = origin
    wsM.Cells(r, COL_M_FAILN).Value = newFail
End Sub

' 既存行の chunk_count 読み出し(行が無い・空欄・非数値・負は0)。R18-2a。
Private Function ChunkCountAt(ByVal wsM As Worksheet, ByVal r As Long) As Long
    On Error Resume Next
    If r < 2 Then Exit Function
    Dim v As Variant: v = wsM.Cells(r, 5).Value
    If IsError(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    ChunkCountAt = CLng(v)
    If ChunkCountAt < 0 Then ChunkCountAt = 0
    On Error GoTo 0
End Function

' 連続失敗回数の読み出し(空欄・非数値は0)。
Private Function FailCountAt(ByVal wsM As Worksheet, ByVal r As Long) As Long
    On Error Resume Next
    If r < 2 Then Exit Function
    Dim v As Variant: v = wsM.Cells(r, COL_M_FAILN).Value
    If IsError(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    FailCountAt = CLng(v)
    If FailCountAt < 0 Then FailCountAt = 0
    On Error GoTo 0
End Function

' ResetFailCountForPath - 恒久失敗(failed_permanent)からの復帰口(R12-3-3)。
'   利用者が「資料を追加」で明示的に選び直したときに呼ぶ。連続失敗回数を0へ、
'   status も "failed" へ戻して再取込の対象に復帰させる。自動処理からは呼ばない
'   =「もう一度やってみる」という人の意思だけがバックオフを解除できる。
Public Sub ResetFailCountForPath(ByVal filePath As String)
    On Error Resume Next
    If LenB(Trim$(filePath)) = 0 Then Exit Sub
    Dim wsM As Worksheet: Set wsM = GetSheet(modAppDef.SH_MANIFEST)
    If wsM Is Nothing Then Exit Sub
    Dim r As Long: r = FindManifestRowByPath(wsM, filePath)
    If r < 2 Then Exit Sub
    If FailCountAt(wsM, r) = 0 Then Exit Sub
    wsM.Cells(r, COL_M_FAILN).Value = 0
    If LCase$(CStr(wsM.Cells(r, COL_M_STATUS).Value)) = STATUS_FAILED_PERMANENT Then
        wsM.Cells(r, COL_M_STATUS).Value = "failed"
        modLog.LogUsage "ingest_backoff_reset", "", _
            "明示選択により再試行対象へ戻しました: " & modUtil.SafeLeft(filePath, 200)
    End If
    On Error GoTo 0
End Sub

' 同名別パスのmanifest行(self)を探す。同一パス(置換)は衝突扱いしない。
Public Function FindConflictingManifestPath(ByVal sourceName As String, ByVal newPath As String) As String
    Dim wsM As Worksheet: Set wsM = GetSheet(modAppDef.SH_MANIFEST)
    If wsM Is Nothing Then Exit Function
    Dim lastM As Long: lastM = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
    If lastM < 2 Then Exit Function

    Dim arr As Variant: arr = wsM.Range(wsM.Cells(2, 1), wsM.Cells(lastM, 9)).Value
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(CStr(arr(i, 9)), "self", vbTextCompare) = 0 Then
            If StrComp(CStr(arr(i, 2)), sourceName, vbTextCompare) = 0 Then
                If StrComp(CStr(arr(i, 1)), newPath, vbTextCompare) <> 0 Then
                    FindConflictingManifestPath = CStr(arr(i, 1))
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Public Function FindManifestRowByPath(ByVal wsM As Worksheet, ByVal filePath As String) As Long
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

' src の先頭n行だけを(1 To n, 1 To 10)へ詰め直す(Range書込みは配列サイズ一致が必須)。
Public Function CompactRows(ByRef src As Variant, ByVal n As Long) As Variant
    CompactRows = SliceRows(src, 1, n)
End Function

' src の startIdx 行目から count 行を(1 To count, 1 To 10)へ切り出す
' (バッチ書込み用。CompactRowsの一般化)。
Public Function SliceRows(ByRef src As Variant, ByVal startIdx As Long, ByVal count As Long) As Variant
    Dim outArr() As Variant: ReDim outArr(1 To count, 1 To KNOWLEDGE_COLS)
    Dim r As Long, c As Long
    For r = 1 To count
        For c = 1 To KNOWLEDGE_COLS
            outArr(r, c) = src(startIdx + r - 1, c)
        Next c
    Next r
    SliceRows = outArr
End Function
