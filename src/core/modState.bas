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
