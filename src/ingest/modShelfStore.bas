Attribute VB_Name = "modShelfStore"
Option Explicit

' ========================================
' modShelfStore - 本棚シート(my_knowledge / my_vectors / my_manifest)の行操作
'
' modShelf から切り出した「シートを作る・行を探す・行を消す・配列で書き戻す」層。
' 取込の判断ロジック(modShelf)と、シートという保存先の都合(ここ)を分けることで、
' 取込側を触るたびに行削除のコードまで読まされる状態を解消する。
'
' 切り出しの理由(2026-07-28): modShelf が契約上限30,000字に対し残り13字まで
' 逼迫しており、バグ修正で1行足すこともできなくなっていた(レビュー I-2)。
' 行操作は取込フロー以外(同期・失効ワイプ)からも呼びたい共通処理のため、
' ここを共有の置き場にする。
'
' 全行の読み書きは「Range一括読み→配列でフィルタ→一括書き戻し」で行う。
' 1行ずつ Rows().Delete すると数千行で実機が数分固まるため(MASTER_SPEC §12)。
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

' ----------------------------------------------------------------------------
' my_manifest の列(1..10)。10列目 fail_count は 2026-08-01(R12-3-3)で追加。
' ----------------------------------------------------------------------------
' 恒久失敗ファイルのバックオフ:
'   フォルダ同期は status="failed" の行を毎回 replace 判定で拾い直す
'   (modShelfSync.ResolveDecision)。壊れたPDF・権限の無いファイルのように
'   「何度やっても失敗するもの」は、そのぶん毎回の同期時間と err_log を
'   食い続ける。連続失敗を数え、MAX_FAIL_STREAK 回で "failed_permanent" へ
'   倒して同期スコープから外す。
'   復帰の道は必ず残す: (a) 利用者が「資料を追加」で明示的に選び直したとき
'   (ResetFailCountForPath)、(b) ファイル自体が更新されたとき(サイズ/更新
'   日時の変化は DiffDecision が replace を返すので再試行される)。
Private Const COL_M_STATUS As Long = 6
Private Const COL_M_FAILN As Long = 10
Private Const MANIFEST_COLS As Long = 10
Private Const MAX_FAIL_STREAK As Long = 3
' status 語彙は "done"/"pending"/"partial"/"failed"/"image_pdf"/"missing" と
' 同じく素の文字列で持つ(読み手=modShelfSync.ResolveDecision・modUIShelf は
' LOテストの都合で本モジュールを参照できないため、既存の語彙と作法を揃える)。
Private Const STATUS_FAILED_PERMANENT As String = "failed_permanent"

' シートを1枚取る。無ければ Nothing(呼び出し側が黙って諦められるように)。
' 2026-07-28: modShelf からここへ切り出したとき、この関数だけ切り出し範囲の
' 外にあり、持ってくるのを忘れていた。実機で
' 「Sub または Function が定義されていません」となり、資料の取込が全滅した。
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
        hdr = Array("chunk_id", "source", "origin", "page", "summary", "keywords", "full_text", "added_at", "embedded")
        Dim i As Long
        For i = LBound(hdr) To UBound(hdr)
            ws.Cells(1, i + 1).Value = hdr(i)
        Next i
        On Error Resume Next
        ws.Visible = 2   ' xlSheetVeryHidden
        On Error GoTo 0
    End If
    ' 数式インジェクション防御(毎回・冪等): 配布テンプレートの既存my_knowledgeでも
    ' 効くよう If の外で適用。非信頼テキスト列を"@"書式へ固定し先頭=等の格納型数式化を防ぐ。
    On Error Resume Next
    ws.Columns(COL_SOURCE).NumberFormat = "@"
    ws.Columns(COL_SUMMARY).NumberFormat = "@"
    ws.Columns(COL_KEYWORDS).NumberFormat = "@"
    ws.Columns(COL_FULLTEXT).NumberFormat = "@"
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
    ' 既存ブック(9列時代のmy_manifest)への見出し追加も毎回・冪等に行う。
    ' 値が空の行は fail_count=0 として読むので、移行処理は要らない。
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

' 既存チャンクのハッシュ集合(重複排除用)。
'
' excludeSource (2026-07-28 レビュー H-6):
'   このsource名の行をハッシュ集合に入れない。同名資料の置き換えで
'   「先に消してから抽出する」のをやめ、「抽出が成功してから消す」順に
'   変えるために要る。消す前に集合を作ると、入れ直すチャンクが全部
'   自分自身との重複と判定されて1件も入らなくなる。
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
Public Sub RemoveKnowledgeAndVectorsForSource(ByVal sourceName As String)
    Dim wsK As Worksheet: Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If wsK Is Nothing Then Exit Sub
    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then Exit Sub

    Dim arr As Variant: arr = wsK.Range(wsK.Cells(2, 1), wsK.Cells(lastK, 9)).Value
    Dim nRows As Long: nRows = UBound(arr, 1) - LBound(arr, 1) + 1

    Dim removedIds As Object: Set removedIds = CreateObject("Scripting.Dictionary")
    Dim survivors() As Variant: ReDim survivors(1 To nRows, 1 To 9)
    Dim survivorCount As Long: survivorCount = 0

    Dim i As Long, c As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(CStr(arr(i, COL_SOURCE)), sourceName, vbTextCompare) = 0 Then
            Dim cid As String: cid = CStr(arr(i, COL_ID))
            If LenB(cid) > 0 Then
                If Not removedIds.Exists(cid) Then removedIds.Add cid, True
            End If
        Else
            survivorCount = survivorCount + 1
            For c = 1 To 9
                survivors(survivorCount, c) = arr(i, c)
            Next c
        End If
    Next i

    If removedIds.count = 0 Then Exit Sub   ' 一致無し(何もしなくてよい)

    If survivorCount > 0 Then
        Dim writeArr As Variant: writeArr = CompactRows(survivors, survivorCount)
        wsK.Range(wsK.Cells(2, 1), wsK.Cells(1 + survivorCount, 9)).Value = writeArr
    End If
    If survivorCount < nRows Then
        wsK.Range(wsK.Cells(2 + survivorCount, 1), wsK.Cells(1 + nRows, 9)).ClearContents
    End If

    RemoveVectorsByIds removedIds
End Sub

' ----------------------------------------------------------------------------
' RemoveRowsByOrigin - origin列が originTag と一致する行を knowledge/vectors
'   から取り除く。戻り値=消した件数。
'
' 2026-07-28(レビュー C-1): この「origin で消す」処理は modChannel が
' 自前に持っており、消す側のタグ("pack:"&部門名)と書く側のタグ
' ("pack:"&作者名)が食い違ったまま誰も気付かなかった。同じ概念の実装が
' 2つあると片方だけ直る。書き手(modPack)と消し手(modChannel)の両方が
' ここを呼ぶようにして、タグの取り扱いを1箇所に集める。
'
' 比較は StrComp(vbTextCompare)=大文字小文字を無視。部門名の表記ゆれで
' 消し漏らすより、寄せて消せる方が事故が小さい。
' ----------------------------------------------------------------------------
Public Function RemoveRowsByOrigin(ByVal originTag As String) As Long
    If LenB(Trim$(originTag)) = 0 Then Exit Function
    Dim wsK As Worksheet: Set wsK = GetSheet(modAppDef.SH_KNOWLEDGE)
    If wsK Is Nothing Then Exit Function
    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then Exit Function

    Dim arr As Variant: arr = wsK.Range(wsK.Cells(2, 1), wsK.Cells(lastK, 9)).Value
    Dim nRows As Long: nRows = UBound(arr, 1) - LBound(arr, 1) + 1

    Dim removedIds As Object: Set removedIds = CreateObject("Scripting.Dictionary")
    Dim survivors() As Variant: ReDim survivors(1 To nRows, 1 To 9)
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
            For c = 1 To 9
                survivors(survivorCount, c) = arr(i, c)
            Next c
        End If
    Next i

    If survivorCount = nRows Then Exit Function   ' 一致無し

    If survivorCount > 0 Then
        Dim writeArr As Variant: writeArr = CompactRows(survivors, survivorCount)
        wsK.Range(wsK.Cells(2, 1), wsK.Cells(1 + survivorCount, 9)).Value = writeArr
    End If
    wsK.Range(wsK.Cells(2 + survivorCount, 1), wsK.Cells(1 + nRows, 9)).ClearContents

    RemoveVectorsByIds removedIds
    RemoveRowsByOrigin = nRows - survivorCount
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

    Dim arr As Variant: arr = wsK.Range(wsK.Cells(2, 1), wsK.Cells(lastK, 9)).Value
    Dim nRows As Long: nRows = UBound(arr, 1) - LBound(arr, 1) + 1
    Dim pfxLen As Long: pfxLen = Len(prefixTag)

    Dim removedIds As Object: Set removedIds = CreateObject("Scripting.Dictionary")
    Dim survivors() As Variant: ReDim survivors(1 To nRows, 1 To 9)
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
            For c = 1 To 9
                survivors(survivorCount, c) = arr(i, c)
            Next c
        End If
    Next i

    If survivorCount = nRows Then Exit Function

    If survivorCount > 0 Then
        Dim writeArr2 As Variant: writeArr2 = CompactRows(survivors, survivorCount)
        wsK.Range(wsK.Cells(2, 1), wsK.Cells(1 + survivorCount, 9)).Value = writeArr2
    End If
    wsK.Range(wsK.Cells(2 + survivorCount, 1), wsK.Cells(1 + nRows, 9)).ClearContents

    RemoveVectorsByIds removedIds
    RemoveRowsByOriginPrefix = nRows - survivorCount
End Function

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

    ' manifest は10列(R12-3-3 で fail_count を追加)。行を詰める処理だけは
    ' 全列を運ばないと、削除の前後で fail_count だけが別の行に残る。
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

    ' R12-3-3: 連続失敗のバックオフ。ここが manifest への status 書込みの
    ' 唯一の入口なので、数える場所もここ1箇所にする(§4-5 同型の問題は
    ' 共通部品で一度だけ)。
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
    wsM.Cells(r, 5).Value = chunkCount
    wsM.Cells(r, COL_M_STATUS).Value = newStatus
    wsM.Cells(r, 7).Value = modUtil.SafeLeft(errorNote, 2000)
    wsM.Cells(r, 8).Value = modUtil.NowStamp()
    wsM.Cells(r, 9).Value = origin
    wsM.Cells(r, COL_M_FAILN).Value = newFail
End Sub

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

' ----------------------------------------------------------------------------
' ResetFailCountForPath - 恒久失敗(failed_permanent)からの復帰口(R12-3-3)。
'   利用者が「資料を追加」で同じファイルを明示的に選び直したときに呼ぶ。
'   連続失敗回数を0に戻し、status も "failed" へ戻すことで、通常の再取込
'   (と、次回以降の同期)の対象へ復帰させる。
'   自動処理からは呼ばない。「もう一度やってみる」という人の意思だけが
'   バックオフを解除できる、というのがこの機能の約束である。
' ----------------------------------------------------------------------------
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

' src(1 To n以上, 1 To 9)の先頭n行だけを(1 To n, 1 To 9)へ詰め直す(Range書込みは配列サイズ一致が必須)。
Public Function CompactRows(ByRef src As Variant, ByVal n As Long) As Variant
    CompactRows = SliceRows(src, 1, n)
End Function

' src(1 To n以上, 1 To 9)の startIdx 行目から count 行を(1 To count, 1 To 9)へ
' 切り出す(バッチ書込み用。CompactRowsの一般化)。
Public Function SliceRows(ByRef src As Variant, ByVal startIdx As Long, ByVal count As Long) As Variant
    Dim outArr() As Variant: ReDim outArr(1 To count, 1 To 9)
    Dim r As Long, c As Long
    For r = 1 To count
        For c = 1 To 9
            outArr(r, c) = src(startIdx + r - 1, c)
        Next c
    Next r
    SliceRows = outArr
End Function
