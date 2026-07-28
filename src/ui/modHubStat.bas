Attribute VB_Name = "modHubStat"
Option Explicit

' 更新保留チャンネルの表示ラベル。"|"区切りの保留リストから、
' 1件ならその名前、2件以上なら「先頭ほか N件」を返す。
' 2026-07-28(レビュー H-13): 保留の実体を見ずにアクティブ部門名を出して
' いたため、「押しても『既に最新です』でバッジが消えない」という
' 説明不能な状態になっていた。名前は必ず保留リストから取る。
Public Function PendingLabel(ByVal pendList As String) As String
    On Error Resume Next
    Dim s As String: s = Trim$(pendList)
    If LenB(s) = 0 Then Exit Function
    Dim parts() As String: parts = Split(s, "|")
    Dim n As Long: n = UBound(parts) - LBound(parts) + 1
    If n <= 1 Then
        PendingLabel = Trim$(parts(LBound(parts)))
    Else
        PendingLabel = Trim$(parts(LBound(parts))) & " ほか" & (n - 1) & "部門"
    End If
    On Error GoTo 0
End Function

' ========================================
' modHubStat - Hub(ホーム)が出す数値の取得・整形と、Hub図形の一括削除
'
' modHub から切り出した裏方。タイルに出る数字が変な話と、タイルの位置が
' ずれる話を別々に追えるようにする。RemoveHubShapes は旧ホーム画面の
' btn_/lbl_ まで消す再描画の要で、描画側と一緒にいると見落とされやすい。
'
' 切り出しの理由(2026-07-28): modHub が契約上限30,000字に対し残り712字で、
' 更新バッジの修正(レビュー H-13)を入れる余裕が乏しかった(レビュー I-2)。
' ========================================

' nx_hub_ に加え旧ホーム画面のbtn_/lbl_も消す(残ると上に浮く)。
Public Sub RemoveHubShapes(ByVal ws As Worksheet)
    Dim names() As String
    ReDim names(0 To ws.Shapes.Count)
    Dim n As Long
    Dim shp As Shape
    For Each shp In ws.Shapes
        Dim nm As String: nm = shp.Name
        If Left$(nm, 7) = "nx_hub_" Or Left$(nm, 4) = "btn_" Or Left$(nm, 4) = "lbl_" Then
            names(n) = nm
            n = n + 1
        End If
    Next shp
    Dim i As Long
    For i = 0 To n - 1
        On Error Resume Next
        ws.Shapes(names(i)).Delete
        On Error GoTo 0
    Next i
End Sub

' 数値を必ず表示できる文字列にする(空欄にしない)。
Public Function NumText(ByVal v As Long) As String
    NumText = CStr(v)
    If LenB(NumText) = 0 Then NumText = "0"
End Function

Public Function SafeStat(ByVal key As String) As Long
    On Error Resume Next
    SafeStat = modStats.GetStat(key)
    On Error GoTo 0
End Function

' 質問回数は quick/deep 別カウンタの合算(単一のquestion_totalキーは存在しない)。
Public Function AskTotal() As Long
    AskTotal = SafeStat("ask_quick_total") + SafeStat("ask_deep_total")
End Function

Public Function SafeSavedMinutes() As Long
    On Error Resume Next
    SafeSavedMinutes = modStats.SavedMinutesEstimate()
    On Error GoTo 0
End Function

Public Function SafeChunks() As Long
    On Error Resume Next
    SafeChunks = modShelf.TotalChunks()
    On Error GoTo 0
End Function

' 本棚の使用量。実数だけ出しても上限が分からないので割合で見せる。
' 部門チャンネルを増やすほど埋まるので、増やしてよいかの判断材料になる。
Public Function ChunkUsage() As String
    Dim pct As Long
    On Error Resume Next
    pct = modChannel.ChunkUsagePercent()
    On Error GoTo 0
    ChunkUsage = pct & "%"
End Function

Public Function FmtMin(ByVal minutes As Long) As String
    If minutes < 60 Then
        FmtMin = CStr(minutes) & "分"
    Else
        FmtMin = CStr(minutes \ 60) & "時間"
        If (minutes Mod 60) > 0 Then FmtMin = FmtMin & CStr(minutes Mod 60) & "分"
    End If
    If LenB(FmtMin) = 0 Then FmtMin = "0分"
End Function

Public Function OrgMin(ByVal period As String) As String
    Dim v As Long
    On Error Resume Next
    If period = "d" Then
        v = modBoard.OrgMinutesDay()
    Else
        v = modBoard.OrgMinutesMon()
    End If
    On Error GoTo 0
    OrgMin = FmtMin(v)
End Function
