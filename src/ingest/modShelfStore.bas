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
        hdr = Array("file_path", "file_name", "modified_at", "size", "chunk_count", "status", "error_note", "ingested_at", "origin")
        Dim i As Long
        For i = LBound(hdr) To UBound(hdr)
            ws.Cells(1, i + 1).Value = hdr(i)
        Next i
        On Error Resume Next
        ws.Visible = 0   ' xlSheetHidden
        On Error GoTo 0
    End If
    Set EnsureManifestSheet = ws
    Exit Function
Fail:
    Set EnsureManifestSheet = Nothing
End Function

Public Function BuildExistingHashSet(ByVal wsK As Worksheet) As Object
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

    If lastK = 2 Then
        AddHashFromId dict, CStr(wsK.Cells(2, 1).Value)
    Else
        Dim ids As Variant: ids = wsK.Range(wsK.Cells(2, 1), wsK.Cells(lastK, 1)).Value
        Dim i As Long
        For i = LBound(ids, 1) To UBound(ids, 1)
            AddHashFromId dict, CStr(ids(i, 1))
        Next i
    End If
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

    Dim arr As Variant: arr = wsM.Range(wsM.Cells(2, 1), wsM.Cells(lastM, 9)).Value
    Dim nRows As Long: nRows = UBound(arr, 1) - LBound(arr, 1) + 1

    Dim survivors() As Variant: ReDim survivors(1 To nRows, 1 To 9)
    Dim survivorCount As Long: survivorCount = 0
    Dim i As Long, c As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(CStr(arr(i, 2)), sourceName, vbTextCompare) <> 0 Then
            survivorCount = survivorCount + 1
            For c = 1 To 9
                survivors(survivorCount, c) = arr(i, c)
            Next c
        End If
    Next i

    If survivorCount = nRows Then Exit Sub   ' 一致無し

    If survivorCount > 0 Then
        Dim writeArr As Variant: writeArr = CompactRows(survivors, survivorCount)
        wsM.Range(wsM.Cells(2, 1), wsM.Cells(1 + survivorCount, 9)).Value = writeArr
    End If
    wsM.Range(wsM.Cells(2 + survivorCount, 1), wsM.Cells(1 + nRows, 9)).ClearContents
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

    wsM.Cells(r, 1).Value = filePath
    wsM.Cells(r, 2).Value = fileName
    wsM.Cells(r, 3).Value = modifiedAt
    wsM.Cells(r, 4).Value = sizeBytes
    wsM.Cells(r, 5).Value = chunkCount
    wsM.Cells(r, 6).Value = status
    wsM.Cells(r, 7).Value = modUtil.SafeLeft(errorNote, 2000)
    wsM.Cells(r, 8).Value = modUtil.NowStamp()
    wsM.Cells(r, 9).Value = origin
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
