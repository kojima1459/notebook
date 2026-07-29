Attribute VB_Name = "modDashStat"
Option Explicit

' ========================================
' modDashStat - ダッシュボードが表示する数値の集計・整形
'
' modDash から切り出した「数字を作る」層。描画(Shape)を一切触らないので
' 数字が合わない話と、絵がずれる話を別々に追える。
'
' 切り出しの理由(2026-07-28): modDash が契約上限30,000字に対し残り67字で、
' 描画の修正で1行足すこともできなくなっていた(レビュー I-2)。
' ========================================

' 自己解決1件あたりの節約時間(分)。modStats の換算値と同じ値
' (向こうは Private のため複製。値を変えるときは必ず両方を直すこと)。
' 2026-07-28: modDash からここへ切り出した際、これを向こうに置いたままにして
' コンパイルエラーになっていた。使う側であるここが持つ。
Private Const MINUTES_PER_SELFSOLVE As Long = 15

' ----------------------------------------------------------------------------
' 内部: 前月比(取り戻した時間)の集計
' ----------------------------------------------------------------------------

' 今月/先月の"feedback_green"件数から前月比の表示文字列を作る。
'   isUp: ▲(上昇/新記録)ならTrue、▼/―ならFalse(アクセント色の切替に使用)。
Public Function SavedTimeDeltaLabel(ByRef isUp As Boolean) As String
    Dim thisMonthCount As Long: thisMonthCount = CountUsageEvent("feedback_green", False)
    Dim lastMonthCount As Long: lastMonthCount = CountUsageEvent("feedback_green", True)
    Dim thisMin As Long: thisMin = thisMonthCount * MINUTES_PER_SELFSOLVE
    Dim lastMin As Long: lastMin = lastMonthCount * MINUTES_PER_SELFSOLVE

    isUp = False
    If lastMin > 0 Then
        Dim pct As Long: pct = CLng(Round((thisMin - lastMin) / lastMin * 100, 0))
        If pct >= 0 Then
            isUp = True
            SavedTimeDeltaLabel = ChrW(&H25B2) & pct & "%"
        Else
            SavedTimeDeltaLabel = ChrW(&H25BC) & Abs(pct) & "%"
        End If
    ElseIf thisMin > 0 Then
        isUp = True
        SavedTimeDeltaLabel = "新記録"
    Else
        SavedTimeDeltaLabel = ChrW(&H2014)
    End If
End Function

' usage_logの指定イベント名を、今月(wantLastMonth=False)または先月
'   (wantLastMonth=True)に絞って件数集計する(modUIDashboard.MonthlyAskCountの一般化)。
Public Function CountUsageEvent(ByVal eventName As String, ByVal wantLastMonth As Boolean) As Long
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_USAGE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim arr As Variant
    If lastRow = 2 Then
        ' 単一行は.Valueがスカラーになるため個別に読む
        Dim tmp(1 To 1, 1 To 2) As Variant
        tmp(1, 1) = ws.Cells(2, 1).Value
        tmp(1, 2) = ws.Cells(2, 2).Value
        arr = tmp
    Else
        arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 2)).Value
    End If

    Dim n As Long: n = 0
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim monthMatch As Boolean
        If wantLastMonth Then
            monthMatch = IsLastMonthStamp(arr(i, 1))
        Else
            monthMatch = IsThisMonthStamp(arr(i, 1))
        End If
        If monthMatch Then
            If StrComp(CStr(arr(i, 2)), eventName, vbTextCompare) = 0 Then
                n = n + 1
            End If
        End If
    Next i
    CountUsageEvent = n
End Function

' timestampセルが「今月」かどうか(modUIDashboard.IsThisMonthStampを移植)。
' 日付型ならYear/Monthで直接比較し、文字列のままなら前置き比較にフォールバック。
Public Function IsThisMonthStamp(ByVal v As Variant) As Boolean
    If IsDate(v) Then
        Dim d As Date: d = CDate(v)
        IsThisMonthStamp = (Year(d) = Year(Now) And Month(d) = Month(Now))
    Else
        IsThisMonthStamp = (Left$(CStr(v), 7) = Format$(Now, "yyyy-mm"))
    End If
End Function

' timestampセルが「先月」かどうか(IsThisMonthStampと同じ考え方)。
Public Function IsLastMonthStamp(ByVal v As Variant) As Boolean
    Dim refDate As Date: refDate = DateAdd("m", -1, Now)
    If IsDate(v) Then
        Dim d As Date: d = CDate(v)
        IsLastMonthStamp = (Year(d) = Year(refDate) And Month(d) = Month(refDate))
    Else
        IsLastMonthStamp = (Left$(CStr(v), 7) = Format$(refDate, "yyyy-mm"))
    End If
End Function

' ----------------------------------------------------------------------------
' 内部: 蔵書チャンク数の使用率バー("▓▓▓░░░ 42%"形式)
' ----------------------------------------------------------------------------
Public Function UsageBarText(ByVal ratio As Double) As String
    Dim r As Double: r = ratio
    If r < 0 Then r = 0
    If r > 1 Then r = 1

    Dim totalBlocks As Long: totalBlocks = 10
    Dim filled As Long: filled = CLng(Round(r * totalBlocks, 0))
    If filled < 0 Then filled = 0
    If filled > totalBlocks Then filled = totalBlocks

    Dim pctInt As Long: pctInt = CLng(Round(r * 100, 0))
    UsageBarText = String$(filled, ChrW(&H2593)) & String$(totalBlocks - filled, ChrW(&H2591)) & " " & pctInt & "%"
End Function

' 分を「n分」「n時間m分」形式に整形する(modUIDashboard.FormatMinutesを移植)。
Public Function FormatMinutes(ByVal minutes As Long) As String
    If minutes < 60 Then
        FormatMinutes = minutes & "分"
    Else
        Dim h As Long: h = minutes \ 60
        Dim m As Long: m = minutes Mod 60
        If m = 0 Then
            FormatMinutes = h & "時間"
        Else
            FormatMinutes = h & "時間" & m & "分"
        End If
    End If
End Function

' ----------------------------------------------------------------------------
' 内部: 各種統計APIの安全ラッパー(未初期化/読取失敗時は既定値)
' ----------------------------------------------------------------------------
Public Function SafeGetStat(ByVal key As String) As Long
    On Error Resume Next
    SafeGetStat = modStats.GetStat(key)
    On Error GoTo 0
End Function

Public Function SafeSavedMinutes() As Long
    On Error Resume Next
    SafeSavedMinutes = modStats.SavedMinutesEstimate()
    On Error GoTo 0
End Function

Public Function SafeTotalChunks() As Long
    On Error Resume Next
    SafeTotalChunks = modShelf.TotalChunks()
    On Error GoTo 0
End Function

Public Function SafeShelfMax() As Long
    On Error Resume Next
    SafeShelfMax = modConfig.GetLong("shelf_max_chunks", 20000)
    On Error GoTo 0
    If SafeShelfMax <= 0 Then SafeShelfMax = 20000
End Function

Public Function SafeLevel() As Long
    On Error Resume Next
    SafeLevel = modStats.Level()
    On Error GoTo 0
    If SafeLevel < 1 Then SafeLevel = 1
End Function

Public Function SafeExpTotal() As Long
    On Error Resume Next
    SafeExpTotal = modStats.ExpTotal()
    On Error GoTo 0
End Function

Public Function SafeExpFloorForLevel(ByVal lv As Long) As Long
    On Error Resume Next
    SafeExpFloorForLevel = modStats.ExpFloorForLevel(lv)
    On Error GoTo 0
End Function

Public Function SafeLevelProgress() As Double
    On Error Resume Next
    SafeLevelProgress = modStats.LevelProgress()
    On Error GoTo 0
End Function
