Attribute VB_Name = "modState"
Option Explicit

' ============================================================================
' modState - ui_stateシート上の汎用キー値ストア(VBAリセット対策)
' ----------------------------------------------------------------------------
' End文/Ctrl+Break/コード編集によるVBAリセットはモジュールレベル変数を
' 全て消去する。会話履歴等の状態を失わないよう、ui_stateシート
' (modAppDef.SH_UISTATE、A列=key/B列=value)へ退避・復元するための
' 汎用ヘルパーをここに集約する(modApp/modAskの両方から使う)。
' 依存は modAppDef.SH_UISTATE のみ(modApp/modAskには依存しない)。
' ============================================================================

' LoadState - keyNameの値をui_stateシートから読む。シート/キーが無い、または
' 値が空の場合はdefaultValを返す(キー一致は大文字小文字を区別しない)。
Public Function LoadState(ByVal keyName As String, ByVal defaultVal As String) As String
    LoadState = defaultVal

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_UISTATE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(CStr(ws.Cells(i, 1).Value), keyName, vbTextCompare) = 0 Then
            ' 既存ブックの救済(2026-08-16 R33 W4-8)。B列がテキスト書式で
            ' なかった頃に保存された「= で始まる質問」は、数式として
            ' 格納されている(数式として有効なら計算結果や #NAME? が、
            ' 不正なら1004で書き込み自体が失われている)。数式セルから
            ' 読んだ値は保存した文字列ではないので、既定値のまま返す。
            ' ここで拾ってしまうと、質問文の代わりに計算結果や #NAME? が
            ' 会話復元・逆質問の照合に混ざる。
            If ws.Cells(i, 2).HasFormula Then Exit For
            Dim v As String
            v = CStr(ws.Cells(i, 2).Value)
            If LenB(v) > 0 Then LoadState = v
            Exit For
        End If
    Next i
    On Error GoTo 0
End Function

' SaveState - keyNameの行を検索し(無ければ末尾に追加)、valTextを書き込む
' (upsert)。値は大小文字を保持したまま保存する(履歴は大小文字区別あり)。
' シートが存在しない場合は何もしない。
Public Sub SaveState(ByVal keyName As String, ByVal valText As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_UISTATE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ' 【なぜ書く前に書式を張るのか】(2026-08-16 R33 W4-8)
    ' Range.Value への代入は、先頭が "=" の文字列を【数式として】解釈する。
    ' 数式として不正なら実行時エラー1004になり、それは下の
    ' On Error Resume Next に飲まれて【そのキーだけが無言で書かれない】。
    ' 実例: 「=IFERRORの使い方を教えて」と聞くと nexus_hist_u だけが落ち、
    ' nexus_hist_a は成功するので、次回起動の会話復元で質問と回答が
    ' 1つズレる(同じズレた組が深掘り時の文脈として LLM にも渡る)。
    ' modClarify の clarify_q が落ちれば HasPending()=False になり、
    ' 利用者が指示どおり返した「(2)」が新しい質問として実行される。
    ' 数式として【有効】な文字列は保存自体は通るが、読み出しでは
    ' 計算結果や #NAME? に化ける。
    ' 他シートは同じ危険をビルド時(_make_headers_only の text_cols)と
    ' 実行時(modShelfStore/modChunkMetaStore 等の Columns(...).NumberFormat)
    ' の二重で潰しており、ui_state だけが両方とも抜けていた。ここは
    ' 実行時側の一枚で、ビルド側は build_mybookshelf.py の ui_state 生成に
    ' text_cols=[2] を渡して揃えてある。
    ws.Columns(2).NumberFormat = "@"
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim r As Long: r = 0
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(CStr(ws.Cells(i, 1).Value), keyName, vbTextCompare) = 0 Then
            r = i
            Exit For
        End If
    Next i
    If r = 0 Then
        r = lastRow + 1
        If r < 1 Then r = 1
        ws.Cells(r, 1).Value = keyName
    End If
    ws.Cells(r, 2).Value = valText
    On Error GoTo 0
End Sub
