Attribute VB_Name = "modStats"
Option Explicit

' ============================================================================
' modStats - 個人統計カウンタ・連続利用日数・バッジ判定(MASTER_SPEC §7.5)
' ----------------------------------------------------------------------------
' 役割:
'   my_stats(key, value, updated_at)への読み書きを一元化する。ダッシュボード
'   (modUIDashboard、担当外)はこのモジュールのGetStat/SavedMinutesEstimate等
'   を読むだけで統計タイルを描画できるようにする。
'
' 設計判断:
'   ・usage_logへの書込は modLog.LogUsage が既に契約として持っているため、
'     ここでは重複実装しない(MASTER_SPEC本文の指示どおり)。modStatsは
'     あくまでmy_stats(集計値)の管理に専念する。
'   ・Bump/GetStatは汎用カウンタ用途(value列がLong)。streak_days/
'     last_used_dateは「増分ではなく値そのものを置き換える」必要があるため、
'     Bumpとは別にPrivateなSetStatValue(汎用upsert)を用意し、TouchTodayは
'     こちらを使う。badge:<id>(取得日文字列)も同じSetStatValueで書く。
'   ・TouchTodayのstreak計算(§7.5): last_used_dateが昨日なら+1、今日なら
'     不変、それ以外は1にリセット。「今日呼ばれたがstreak_daysがまだ0
'     (=初回起動)」のケースだけは1に初期化する(0連続はUIで不自然なため)。
'   ・EvaluateBadgesの§9バッジ条件のうち shelf10/shelf30(本棚10冊/30冊)は
'     「現在の本棚の資料数」であり、my_stats の ingest_files_total(取込
'     操作の累計回数。削除しても減らない)とは意味が異なる。modShelf.
'     SourceList を呼ぶことも検討したが、件数を得るためだけに2つの配列
'     ByRef引数を渡す設計は過剰なため、my_manifestの行数を直接数える
'     小さなPrivateヘルパー(ShelfSourceCount)を持つ(modShelf.TotalChunks
'     等、各モジュールが同種の小さな自前カウントヘルパーを持つ既存の
'     慣習に合わせた)。
'   ・新規バッジ獲得時のみお祝いMsgBoxを出す(badge:<id>が未記録のときだけ)。
'     MsgBox文字列に絵文字は入れない(§12)。バッジ自体の表示装飾
'     (色付き絵文字等)はダッシュボードのセル/Shape側(modUIDashboard、
'     担当外)の仕事であり、ここでは日付を記録するだけ。
'   ・my_stats は MASTER_SPEC §4 で可視性 hidden(veryHiddenではない)と
'     定義されているため、新規作成時は xlSheetHidden(=0)を使う。
' ============================================================================

Private Const MINUTES_PER_SELFSOLVE As Long = 15   ' 自己解決1件=15分換算(§7.5)

' ----------------------------------------------------------------------------
' Bump - my_stats upsert(現在値にdeltaを加算)
' ----------------------------------------------------------------------------
Public Sub Bump(ByVal key As String, Optional ByVal delta As Long = 1)
    Dim ws As Worksheet: Set ws = EnsureStatsSheet()
    If ws Is Nothing Then Exit Sub

    Dim r As Long: r = FindKeyRow(ws, key)
    Dim cur As Long
    If r = 0 Then
        r = NextRow(ws)
        cur = 0
    Else
        cur = SafeCLng(ws.Cells(r, 2).Value)
    End If

    ws.Cells(r, 1).Value = key
    ws.Cells(r, 2).Value = cur + delta
    ws.Cells(r, 3).Value = modUtil.NowStamp()
End Sub

' ----------------------------------------------------------------------------
' GetStat - 現在値(数値以外/未登録は0)
' ----------------------------------------------------------------------------
Public Function GetStat(ByVal key As String) As Long
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_STATS)
    If ws Is Nothing Then Exit Function
    Dim r As Long: r = FindKeyRow(ws, key)
    If r = 0 Then Exit Function
    GetStat = SafeCLng(ws.Cells(r, 2).Value)
End Function

' ----------------------------------------------------------------------------
' TouchToday - streak_days/last_used_date 更新(連続利用日数)
'   last_used_dateが昨日なら+1、今日なら不変、それ以外は1にリセット。
' ----------------------------------------------------------------------------
Public Sub TouchToday()
    Dim todayStr As String: todayStr = Format$(Date, "yyyy-mm-dd")
    Dim yesterdayStr As String: yesterdayStr = Format$(Date - 1, "yyyy-mm-dd")
    Dim lastStr As String: lastStr = GetStatValueString("last_used_date")

    If lastStr = todayStr Then
        If GetStat("streak_days") < 1 Then SetStatValue "streak_days", 1
    ElseIf lastStr = yesterdayStr Then
        SetStatValue "streak_days", GetStat("streak_days") + 1
    Else
        SetStatValue "streak_days", 1
    End If

    SetStatValue "last_used_date", todayStr
End Sub

' ----------------------------------------------------------------------------
' EvaluateBadges - §9のバッジ条件を検査。新規獲得は badge:<id> に日付を記録し、
'   お祝いMsgBoxを1回だけ出す(絵文字はMsgBoxに入れない)。
' ----------------------------------------------------------------------------
Public Sub EvaluateBadges()
    CheckBadge "first_ingest", GetStat("ingest_files_total") >= 1, "はじめての資料取込"
    CheckBadge "shelf10", ShelfSourceCount() >= 10, "本棚に10冊"
    CheckBadge "shelf30", ShelfSourceCount() >= 30, "本棚に30冊"
    CheckBadge "first_pack_out", GetStat("pack_export_total") >= 1, "はじめてのパック共有"
    CheckBadge "first_pack_in", GetStat("pack_import_total") >= 1, "はじめてのパック取込"
    CheckBadge "solve10", GetStat("selfsolve_total") >= 10, "自己解決10件"
    CheckBadge "solve50", GetStat("selfsolve_total") >= 50, "自己解決50件"
    CheckBadge "streak7", GetStat("streak_days") >= 7, "7日連続利用"
End Sub

' ----------------------------------------------------------------------------
' SavedMinutesEstimate - 自己解決1件=15分換算(config化不要・定数で明示)
' ----------------------------------------------------------------------------
Public Function SavedMinutesEstimate() As Long
    SavedMinutesEstimate = GetStat("selfsolve_total") * MINUTES_PER_SELFSOLVE
End Function

' ----------------------------------------------------------------------------
' EXP/レベル(ゲーミフィケーション)
' ----------------------------------------------------------------------------
' 設計:
'   ・EXPは exp_total(単調増加カウンタ)に集約。加点イベントは4種類のみで、
'     各イベントの発火点(質問成功/取込成功/🟢自己解決/パック出力)で AddExp を
'     1回だけ呼ぶ。二重加算はそれぞれの発火点の既存ガード(FeedbackAccepted・
'     IngestFileの戻り値判定等)が防ぐため、ここでは素直に Bump するだけ。
'   ・加点量とレベル除数は config 可変(既定 5/20/10/30、除数100)。configが
'     読めない環境でも既定値で動くよう GetLong の第2引数に同じ既定を置く。
'   ・レベルは Lv = Int(√(EXP / 除数)) + 1(平方根カーブ=最初は上がりやすく
'     だんだん重くなる、体感の良い成長曲線)。EXP=0→Lv1, 100→Lv2, 400→Lv3…。
' ----------------------------------------------------------------------------
Public Sub AddExp(ByVal eventType As String)
    Dim amt As Long
    Select Case LCase$(eventType)
        Case "question":   amt = modConfig.GetLong("exp_question", 5)
        Case "register":   amt = modConfig.GetLong("exp_register", 20)
        Case "thumbup":    amt = modConfig.GetLong("exp_thumbup", 10)
        Case "pack_share": amt = modConfig.GetLong("exp_pack_share", 30)
        Case Else:         amt = 0
    End Select
    If amt <> 0 Then Bump "exp_total", amt
End Sub

' 累計EXP
Public Function ExpTotal() As Long
    ExpTotal = GetStat("exp_total")
End Function

' 現在のレベル Lv = Int(√(EXP / 除数)) + 1(除数<1はガードして100扱い)
Public Function Level() As Long
    Dim divisor As Long: divisor = LevelDivisor()
    Dim e As Long: e = GetStat("exp_total")
    If e < 0 Then e = 0
    Level = Int(Sqr(CDbl(e) / divisor)) + 1
End Function

' 指定レベルに到達するのに必要な累計EXPの下限 = 除数 * (Lv-1)^2
'   (Lv計算の逆関数。ダッシュボードの進捗バー算出に使う)
Public Function ExpFloorForLevel(ByVal lv As Long) As Long
    If lv < 1 Then lv = 1
    ExpFloorForLevel = LevelDivisor() * (lv - 1) * (lv - 1)
End Function

' 現レベル内の進捗(0.0〜1.0)。次レベルまであとどれくらいか、をバーで見せる用。
Public Function LevelProgress() As Double
    Dim lv As Long: lv = Level()
    Dim curFloor As Long: curFloor = ExpFloorForLevel(lv)
    Dim nextFloor As Long: nextFloor = ExpFloorForLevel(lv + 1)
    Dim span As Long: span = nextFloor - curFloor
    If span <= 0 Then Exit Function
    Dim into As Long: into = GetStat("exp_total") - curFloor
    If into < 0 Then into = 0
    LevelProgress = into / span
    If LevelProgress > 1 Then LevelProgress = 1
End Function

Private Function LevelDivisor() As Long
    Dim divisor As Long: divisor = modConfig.GetLong("exp_level_divisor", 100)
    If divisor < 1 Then divisor = 100
    LevelDivisor = divisor
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

Private Sub CheckBadge(ByVal badgeId As String, ByVal achieved As Boolean, ByVal label As String)
    If Not achieved Then Exit Sub
    Dim key As String: key = "badge:" & badgeId
    If LenB(GetStatValueString(key)) > 0 Then Exit Sub   ' 既に取得済み

    SetStatValue key, Format$(Date, "yyyy-mm-dd")

    MsgBox "新しいバッジを獲得しました: " & label & vbLf & _
        "これからも「マイ本棚AI」を使ってみてください。", vbInformation, modAppDef.APP_NAME
End Sub

Private Function ShelfSourceCount() As Long
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_MANIFEST)
    If ws Is Nothing Then Exit Function
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then Exit Function
    ShelfSourceCount = lastR - 1
End Function

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function EnsureStatsSheet() As Worksheet
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_STATS)
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = modAppDef.SH_STATS
        ws.Cells(1, 1).Value = "key"
        ws.Cells(1, 2).Value = "value"
        ws.Cells(1, 3).Value = "updated_at"
        On Error Resume Next
        ws.Visible = 0   ' xlSheetHidden(§4: my_statsはhidden)
        On Error GoTo 0
    End If
    Set EnsureStatsSheet = ws
    Exit Function
Fail:
    Set EnsureStatsSheet = Nothing
End Function

Private Sub SetStatValue(ByVal key As String, ByVal value As Variant)
    Dim ws As Worksheet: Set ws = EnsureStatsSheet()
    If ws Is Nothing Then Exit Sub
    Dim r As Long: r = FindKeyRow(ws, key)
    If r = 0 Then r = NextRow(ws)
    ws.Cells(r, 1).Value = key
    ws.Cells(r, 2).Value = value
    ws.Cells(r, 3).Value = modUtil.NowStamp()
End Sub

Private Function GetStatValueString(ByVal key As String) As String
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_STATS)
    If ws Is Nothing Then Exit Function
    Dim r As Long: r = FindKeyRow(ws, key)
    If r = 0 Then Exit Function
    GetStatValueString = CStr(ws.Cells(r, 2).Value)
End Function

Private Function FindKeyRow(ByVal ws As Worksheet, ByVal key As String) As Long
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then Exit Function
    If lastR = 2 Then
        If StrComp(CStr(ws.Cells(2, 1).Value), key, vbTextCompare) = 0 Then FindKeyRow = 2
        Exit Function
    End If

    Dim arr As Variant: arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastR, 1)).Value
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(CStr(arr(i, 1)), key, vbTextCompare) = 0 Then
            FindKeyRow = i + 1
            Exit Function
        End If
    Next i
End Function

Private Function NextRow(ByVal ws As Worksheet) As Long
    Dim r As Long: r = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If r < 1 Then r = 1
    NextRow = r + 1
End Function

Private Function SafeCLng(ByVal v As Variant) As Long
    If IsNumeric(v) Then SafeCLng = CLng(v)
End Function
