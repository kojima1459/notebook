Attribute VB_Name = "modVectorize"
Option Explicit

' ============================================================================
' modVectorize - ADMIN-ONLY one-time job: embed every knowledge_base chunk and
' store the vector on the veryHidden `kb_vectors` sheet.
' ----------------------------------------------------------------------------
' Only the file owner (小島さん) runs this, ONCE, on a corporate PC where the AI
' ribbon works. End users never touch it; they just use the baked file.
'
' Design goals (the owner is non-technical and dreads breakage):
'   - Resumable : re-running only processes chunks not already vectorised, so a
'                 rate-limit / interruption is harmless — press the button again.
'   - Fault-tolerant : one failed chunk is recorded and skipped, never aborts the
'                 whole run. A final report lists how many failed.
'   - Visible : Application.StatusBar shows live progress; DoEvents keeps Excel
'                 responsive and lets ESC interrupt cleanly (progress is kept).
'   - Self-contained : everything goes through modEmbeddings (the safe gateway).
'
' kb_vectors layout:  A=chunk_id  B=vector_csv  C=dim  D=updated_at
' ============================================================================

Private Const KB As String = "knowledge_base"
Private Const VEC As String = "kb_vectors"
Private Const COL_ID As Long = 1
Private Const COL_SUMMARY As Long = 7
Private Const COL_KEYWORDS As Long = 8
Private Const COL_TEXT As Long = 9

' ----------------------------------------------------------------------------
' Phase-0 button: prove the ribbon embedding works in THIS environment, on 1 item.
' ----------------------------------------------------------------------------
Public Sub Admin_TestEmbeddingConnection()
    Dim detail As String, ok As Boolean
    ok = modEmbeddings.EmbedSelfTest( _
            "テスト：費用利益保険における被保険者の範囲について教えてください。", detail)
    MsgBox detail, IIf(ok, vbInformation, vbCritical), "① 埋め込み接続テスト"
End Sub

' ----------------------------------------------------------------------------
' Phase-1 button: vectorise all knowledge_base chunks (resumable).
' ----------------------------------------------------------------------------
Public Sub Admin_VectorizeAll()
    Dim wsK As Worksheet, wsV As Worksheet
    On Error GoTo Fatal

    Set wsK = ThisWorkbook.Worksheets(KB)
    Set wsV = EnsureVecSheet()

    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, COL_ID).End(xlUp).row
    If lastK < 2 Then
        MsgBox "knowledge_base が空です。先にナレッジを取り込んでください。", vbExclamation, "全件ベクトル化"
        Exit Sub
    End If

    ' Confirm before a long run.
    Dim total As Long: total = lastK - 1
    If MsgBox(total & " 件のチャンクをベクトル化します。" & vbLf & _
              "社内ネットワーク内・AIリボンが有効な状態で実行してください。" & vbLf & vbLf & _
              "数分かかります。途中でESCキーを押すと中断でき、もう一度押すと続きから再開します。" & vbLf & vbLf & _
              "実行しますか？", vbYesNo + vbQuestion, "全件ベクトル化") <> vbYes Then Exit Sub

    ' Resume support: collect ids already present in kb_vectors.
    Dim doneIds As String: doneIds = vbLf & DoneIdList(wsV) & vbLf
    Dim writeRow As Long: writeRow = wsV.Cells(wsV.Rows.count, 1).End(xlUp).row + 1
    If writeRow < 2 Then writeRow = 2

    Dim ok As Long, fail As Long, skip As Long, i As Long
    Dim failIds As String

    Application.ScreenUpdating = False
    For i = 2 To lastK
        Dim id As String: id = CStr(wsK.Cells(i, COL_ID).value)
        If LenB(id) = 0 Then GoTo NextI

        ' Already done? skip (resume).
        If InStr(1, doneIds, vbLf & id & vbLf, vbBinaryCompare) > 0 Then
            skip = skip + 1
            GoTo NextI
        End If

        Application.StatusBar = "ベクトル化中 " & (ok + fail + skip + 1) & " / " & total & _
            "   [成功 " & ok & " / 失敗 " & fail & " / 既存 " & skip & "]   中断:ESC"
        DoEvents

        Dim v As String
        v = modEmbeddings.EmbedText(ComposeText(wsK, i))
        If LenB(v) = 0 Then
            DoEvents                       ' brief yield, then one retry (rate-limit friendly)
            v = modEmbeddings.EmbedText(ComposeText(wsK, i))
        End If

        If LenB(v) = 0 Then
            fail = fail + 1
            If Len(failIds) < 600 Then failIds = failIds & id & "  "
        Else
            wsV.Cells(writeRow, 1).value = id
            wsV.Cells(writeRow, 2).value = Left$(v, 32000)
            wsV.Cells(writeRow, 3).value = modEmbeddings.CountValues(v)
            wsV.Cells(writeRow, 4).value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
            writeRow = writeRow + 1
            ok = ok + 1
            doneIds = doneIds & id & vbLf
        End If
NextI:
    Next i

    Application.StatusBar = False
    Application.ScreenUpdating = True
    SaveQuietly

    MsgBox "全件ベクトル化 完了" & vbLf & vbLf & _
           "  成功: " & ok & vbLf & _
           "  既存スキップ: " & skip & vbLf & _
           "  失敗: " & fail & vbLf & _
           IIf(fail > 0, "  失敗ID(一部): " & failIds & vbLf & _
                         "  → もう一度このボタンを押すと、失敗分だけ再試行します。", _
                         "  → 全件成功 ✅") & vbLf & vbLf & _
           "ファイルを保存しました。この状態で配布してください。", _
           IIf(fail > 0, vbExclamation, vbInformation), "全件ベクトル化"
    Exit Sub

Fatal:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    ' ESC (Err 18) or any error: keep whatever was written, allow resume.
    On Error Resume Next
    SaveQuietly
    On Error GoTo 0
    MsgBox "中断しました (Err " & Err.Number & ": " & Err.Description & ")" & vbLf & vbLf & _
           "ここまでの結果は保存しました。もう一度ボタンを押すと続きから再開します。", _
           vbExclamation, "全件ベクトル化"
End Sub

' ----------------------------------------------------------------------------
' Compose the text that represents a chunk for retrieval: summary + keywords +
' body. modEmbeddings caps the length, so no truncation needed here.
' ----------------------------------------------------------------------------
Private Function ComposeText(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim s As String
    s = CStr(ws.Cells(r, COL_SUMMARY).value) & " " & _
        CStr(ws.Cells(r, COL_KEYWORDS).value) & " " & _
        CStr(ws.Cells(r, COL_TEXT).value)
    ComposeText = Trim$(s)
End Function

' ----------------------------------------------------------------------------
' Return a vbLf-joined list of chunk_ids already in kb_vectors (for resume).
' ----------------------------------------------------------------------------
Private Function DoneIdList(ByVal ws As Worksheet) As String
    Dim last As Long: last = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If last < 2 Then Exit Function
    If last = 2 Then
        DoneIdList = CStr(ws.Cells(2, 1).value)
        Exit Function
    End If
    Dim arr As Variant: arr = ws.Range(ws.Cells(2, 1), ws.Cells(last, 1)).value
    Dim s As String, i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If LenB(arr(i, 1)) > 0 Then s = s & CStr(arr(i, 1)) & vbLf
    Next i
    DoneIdList = s
End Function

' ----------------------------------------------------------------------------
' Get (or create) the veryHidden kb_vectors sheet with its header row.
' ----------------------------------------------------------------------------
Private Function EnsureVecSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(VEC)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = VEC
        ws.Range("A1").value = "chunk_id"
        ws.Range("B1").value = "vector_csv"
        ws.Range("C1").value = "dim"
        ws.Range("D1").value = "updated_at"
    End If
    On Error Resume Next
    ws.Visible = 2          ' xlSheetVeryHidden
    On Error GoTo 0
    Set EnsureVecSheet = ws
End Function

Private Sub SaveQuietly()
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0
End Sub

' ============================================================================
' EnsureAdminPanel - build the hidden "admin" sheet with the two buttons.
' Called from modBoot.Boot. The sheet is xlSheetHidden (not veryHidden) so the
' owner can reach it via タブ右クリック → 再表示. End users never see it in use.
' ============================================================================
Public Sub EnsureAdminPanel()
    On Error GoTo Done
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("admin")
    On Error GoTo Done
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = "admin"
    End If

    ws.Cells.Clear
    ' Backward count-based delete: For Each + Delete can skip items and leave
    ' duplicate buttons stacking up on every Boot.
    Dim k As Long
    For k = ws.Shapes.count To 1 Step -1
        ws.Shapes(k).Delete
    Next k

    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B:E").ColumnWidth = 16

    With ws.Range("A1:F1")
        .Merge
        .value = "【管理者用】RAG ベクトル化パネル （配布前に小島さんが一度だけ実行）"
        .Font.Bold = True
        .Font.Size = 13
        .Interior.color = RGB(70, 70, 110)
        .Font.color = RGB(255, 255, 255)
        .RowHeight = 26
    End With

    With ws.Range("A3:F4")
        .Merge
        .value = "手順:  ① まず「埋め込み接続テスト」を押し、成功 ✅ を確認" & vbLf & _
                 "        ② 「全件ベクトル化」を押す（数分／中断ESC・再開可）→ 終わったら保存して配布"
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Size = 10
    End With

    AddBtn ws, ws.Range("B6:E6"), "btnEmbTest", "① 埋め込み接続テスト（まず押す）", _
           "modVectorize.Admin_TestEmbeddingConnection"
    AddBtn ws, ws.Range("B8:E8"), "btnVecAll", "② 全件ベクトル化を実行", _
           "modVectorize.Admin_VectorizeAll"

    With ws.Range("A10:F12")
        .Merge
        .value = "※ オフラインで動作だけ確認したい時は config の mock_llm を TRUE に。" & vbLf & _
                 "※ ナレッジを更新したら、再度「全件ベクトル化」を押してください（変更分のみ追加）。" & vbLf & _
                 "※ この admin シートは配布ファイルでは非表示です（タブ右クリック→再表示 で再度開けます）。"
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Size = 9
        .Font.color = RGB(110, 110, 110)
    End With

    On Error Resume Next
    ws.Visible = 0          ' xlSheetHidden (owner can 再表示; end users won't bother)
    On Error GoTo 0
Done:
End Sub

Private Sub AddBtn(ByVal ws As Worksheet, ByVal rng As Range, _
                   ByVal nm As String, ByVal cap As String, ByVal action As String)
    On Error Resume Next
    Dim b As Button
    Set b = ws.Buttons.Add(rng.Left, rng.Top, rng.Width, rng.Height)
    b.OnAction = action
    b.caption = cap
    b.name = nm
    b.Font.Size = 11
    On Error GoTo 0
End Sub
