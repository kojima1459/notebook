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

' 2026-07-31(R7実装後の追加是正・発見事項1): 215pt×4枚+間隔だと902ptになり、
' 実画面幅(他画面の実測で約597〜650pt)を大幅に超えて3・4枚目が画面外に出る。
' A-1ではヘッダー帯だけを幅内に収めたが、本文のカード自体は放置されていた。
' 4枚とも1画面に収まる130ptへ縮小する(バッジも同じ行幅に揃える)。
' KPI_CARD_W/KPI_GAP は ROW_WIDTH の定数式が参照するため Public。
' (LibreOffice のコンパイルは Public Const が Private Const を参照すると
'  応答不能になる。実Excelでは私有のままでも通るが、モード2の構文チェックを
'  通すために公開側へ揃える。2026-07-31 R11-F1で実測。)
Public Const KPI_CARD_W As Double = 130
Private Const KPI_CARD_H As Double = 92
Public Const KPI_GAP As Double = 10
Public Const KPI_X0 As Double = 20
' 2026-07-31(R7 A-1): ヘッダー帯(48pt)+サブタイトルの下から本文を始める。
Private Const KPI_Y0 As Double = 80
Public Const ROW_WIDTH As Double = KPI_CARD_W * 4 + KPI_GAP * 3

Private Const EXPBAR_H As Double = 12
Private Const EXPBAR_Y As Double = KPI_Y0 + KPI_CARD_H + 18
Private Const EXPLABEL_Y As Double = EXPBAR_Y + EXPBAR_H + 4

' 2026-07-31: KPIカードと同じ行幅(130*4+10*3+X0*2=590pt)に揃える。
Private Const BADGE_W As Double = 130
Private Const BADGE_H As Double = 54
Private Const BADGE_GAP_X As Double = 10
Private Const BADGE_GAP_Y As Double = 10
Private Const BADGES_PER_ROW As Long = 4
Private Const BADGE_HEAD_Y As Double = EXPLABEL_Y + 22
Private Const BADGE_GRID_Y As Double = BADGE_HEAD_Y + 24

' 2026-07-28(解説書 §11-11): バッジが8種から12種に増えた。
' 以前は「2行ぶん」と決め打ちしていたため、増えた行がチャート枠と重なる。
' VBAの定数式では関数を呼べないので、位置は実行時に ChartNoteY() で求める。
Private Const BADGE_ROWS_FALLBACK As Long = 3

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
        IsThisMonthStamp = (Left$(CStr(v), 7) = Left$(modUtilText.IsoDate(Date), 7))
    End If
End Function

' timestampセルが「先月」かどうか(IsThisMonthStampと同じ考え方)。
Public Function IsLastMonthStamp(ByVal v As Variant) As Boolean
    Dim refDate As Date: refDate = DateAdd("m", -1, Now)
    If IsDate(v) Then
        Dim d As Date: d = CDate(v)
        IsLastMonthStamp = (Year(d) = Year(refDate) And Month(d) = Month(refDate))
    Else
        IsLastMonthStamp = (Left$(CStr(v), 7) = Left$(modUtilText.IsoDate(refDate), 7))
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
    SafeShelfMax = modConfig.GetLong("shelf_max_chunks", modAppDef.DEFAULT_SHELF_MAX_CHUNKS)
    On Error GoTo 0
    If SafeShelfMax <= 0 Then SafeShelfMax = modAppDef.DEFAULT_SHELF_MAX_CHUNKS
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

' ----------------------------------------------------------------------------
' ダッシュボード本文の描画(KPIカード4枚 / 経験値バー / バッジ棚 / クラスタ地図)。
'   2026-07-31(R11-F1)に modDash から移設した。modDash が27,961字でWARN帯
'   (28,000字)の直前におり、R11-F2で足す「?」ヘルプ入口が入らない状態
'   だったため(憲章§4-6)。modDash 側にはシート取得・ヘッダー・管理者
'   セクション・ハンドラが残る。数値の取り出し(SafeGetStat等)は元から
'   本モジュールにあり、その値を「どう見せるか」がここに揃った形になる。
'   KPI_X0 / ROW_WIDTH は本文とヘッダー・管理者行が共有する版面の基準なので
'   Public にし、modDash から modDashStat.KPI_X0 として参照する。
' ----------------------------------------------------------------------------



' ---- KPIカード4枚 ----

Public Sub DrawKpiRow(ByVal ws As Worksheet)
    ' Card0: 取り戻した時間(+前月比)
    Dim savedMinutes As Long: savedMinutes = modDashStat.SafeSavedMinutes()
    Dim isUp As Boolean
    Dim deltaText As String: deltaText = modDashStat.SavedTimeDeltaLabel(isUp)
    Dim deltaColor As Long
    If isUp Then
        deltaColor = modUI.UiColor("accent")
    Else
        deltaColor = modUI.UiColor("muted")
    End If
    DrawKpiCard ws, 0, KPI_X0, KPI_Y0, KPI_CARD_W, KPI_CARD_H, _
        "取り戻した時間", modDashStat.FormatMinutes(savedMinutes), deltaText, deltaColor

    ' Card1: 登録ナレッジ数
    Dim ingestTotal As Long: ingestTotal = modDashStat.SafeGetStat("ingest_files_total")
    DrawKpiCard ws, 1, KPI_X0 + (KPI_CARD_W + KPI_GAP), KPI_Y0, KPI_CARD_W, KPI_CARD_H, _
        "登録ナレッジ数", ingestTotal & "件", "あなたが登録した資料"

    ' Card2: 蔵書チャンク数
    Dim totalChunks As Long: totalChunks = modDashStat.SafeTotalChunks()
    Dim shelfMax As Long: shelfMax = modDashStat.SafeShelfMax()
    Dim ratio As Double
    If shelfMax > 0 Then
        ratio = CDbl(totalChunks) / CDbl(shelfMax)
    Else
        ratio = 0
    End If
    DrawKpiCard ws, 2, KPI_X0 + 2 * (KPI_CARD_W + KPI_GAP), KPI_Y0, KPI_CARD_W, KPI_CARD_H, _
        "蔵書チャンク数", totalChunks & " / " & shelfMax, modDashStat.UsageBarText(ratio)

    ' Card3: レベル
    Dim lv As Long: lv = modDashStat.SafeLevel()
    Dim expTotalV As Long: expTotalV = modDashStat.SafeExpTotal()
    Dim remain As Long: remain = modDashStat.SafeExpFloorForLevel(lv + 1) - expTotalV
    If remain < 0 Then remain = 0
    DrawKpiCard ws, 3, KPI_X0 + 3 * (KPI_CARD_W + KPI_GAP), KPI_Y0, KPI_CARD_W, KPI_CARD_H, _
        "レベル", "Lv." & lv, "EXP " & expTotalV & " ・ 次まで" & remain & "EXP"
End Sub

' 単一Shape=1カード(caption/value/label の3段を1テキストで結合し段別書式)。
'   段1=caption(小・muted)、段2=value(大・太・text)、段3=label(小・指定色)。
' 2026-07-31(R11-E M-4): 従来はOERNが一切無く、1枚のShape生成/書式設定の
' 失敗が例外として呼び出し元(4枚ループ)へ飛び、残りのKPIカードごと本文が
' 全滅していた。カード単位でOERNを閉じ、Paragraphs.Count不足も個別に弾く。
Private Sub DrawKpiCard(ByVal ws As Worksheet, ByVal idx As Long, ByVal x As Double, ByVal y As Double, _
                        ByVal cardW As Double, ByVal cardH As Double, ByVal captionText As String, _
                        ByVal valueText As String, ByVal labelText As String, _
                        Optional ByVal labelColor As Long = -1)
    On Error Resume Next
    Dim card As Shape
    Set card = ws.Shapes.AddShape(5, x, y, cardW, cardH)
    If card Is Nothing Then GoTo Done
    card.Name = "nxd_kpi_" & idx
    card.Adjustments(1) = 0.08
    card.Fill.ForeColor.RGB = modUI.UiColor("surface")
    card.Line.ForeColor.RGB = modUI.UiColor("border")
    card.Line.Weight = 0.75
    card.Shadow.Visible = 0

    Dim finalLabelColor As Long
    If labelColor = -1 Then
        finalLabelColor = modUI.UiColor("muted")
    Else
        finalLabelColor = labelColor
    End If

    ' ParagraphsはvbCr区切りでしか分かれない(vbLfだと1段落のままParagraphs(2)が
    ' 範囲外例外になる。実機報告「KPIカードが1枚しか出ない」の原因)。
    Dim body As String: body = captionText & vbCr & valueText & vbCr & labelText

    With card.TextFrame2
        .WordWrap = -1
        .MarginLeft = 12: .MarginRight = 10: .MarginTop = 9: .MarginBottom = 8
        .TextRange.Text = body
        .TextRange.Font.Size = 9
        .TextRange.Font.Fill.ForeColor.RGB = finalLabelColor
        .VerticalAnchor = 1
        If .TextRange.Paragraphs.Count >= 2 Then
            ' 段1: caption(小・muted)
            With .TextRange.Paragraphs(1).Font
                .Size = 9
                .Fill.ForeColor.RGB = modUI.UiColor("muted")
            End With
            ' 段2: value(大・太・text)
            With .TextRange.Paragraphs(2).Font
                .Size = 17
                .Bold = -1
                .Fill.ForeColor.RGB = modUI.UiColor("text")
            End With
        End If
    End With
Done:
    Err.Clear
    On Error GoTo 0
End Sub

' ---- EXP進捗バー ----

Public Sub DrawExpBar(ByVal ws As Worksheet)
    Dim trackW As Double: trackW = ROW_WIDTH
    Dim prog As Double: prog = modDashStat.SafeLevelProgress()
    If prog < 0 Then prog = 0
    If prog > 1 Then prog = 1

    Dim fillW As Double: fillW = trackW * prog
    If prog > 0 And fillW < 2 Then fillW = 2

    Dim bgBar As Shape
    Set bgBar = ws.Shapes.AddShape(5, KPI_X0, EXPBAR_Y, trackW, EXPBAR_H)
    bgBar.Name = "nxd_expbar_bg"
    bgBar.Adjustments(1) = 0.5
    bgBar.Fill.ForeColor.RGB = modUI.UiColor("border")
    bgBar.Line.Visible = 0
    bgBar.Shadow.Visible = 0

    If fillW > 0 Then
        Dim fgBar As Shape
        Set fgBar = ws.Shapes.AddShape(5, KPI_X0, EXPBAR_Y, fillW, EXPBAR_H)
        fgBar.Name = "nxd_expbar_fg"
        fgBar.Adjustments(1) = 0.5
        fgBar.Fill.ForeColor.RGB = modUI.UiColor("primary")
        fgBar.Line.Visible = 0
        fgBar.Shadow.Visible = 0
    End If

    Dim lv As Long: lv = modDashStat.SafeLevel()
    Dim remain As Long: remain = modDashStat.SafeExpFloorForLevel(lv + 1) - modDashStat.SafeExpTotal()
    If remain < 0 Then remain = 0

    Dim lbl As Shape
    Set lbl = ws.Shapes.AddShape(1, KPI_X0, EXPLABEL_Y, trackW, 16)
    lbl.Name = "nxd_exp_label"
    lbl.Line.Visible = 0
    lbl.Fill.Visible = 0
    With lbl.TextFrame2
        .WordWrap = -1
        .TextRange.Text = "Lv." & (lv + 1) & " まであと " & remain & " EXP"
        .TextRange.Font.Size = 8.5
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
        .VerticalAnchor = 3
        .MarginLeft = 10: .MarginRight = 10: .MarginTop = 6: .MarginBottom = 6
    End With
End Sub

' ---- バッジ棚(8種・4列×2行) ----

' バッジ棚の下端 = チャート類の開始位置。バッジ件数から行数を出すので、
' バッジを増やしてもレイアウト定数を直す必要がない(解説書 §11-11)。
Public Function ChartNoteY() As Double
    Dim rowsN As Long: rowsN = BADGE_ROWS_FALLBACK
    On Error Resume Next
    Dim ids() As String, titles() As String, shorts() As String, conditions() As String
    Dim n As Long: n = modStats.BadgeCatalog(ids, titles, shorts, conditions)
    If n > 0 Then
        rowsN = (n + BADGES_PER_ROW - 1) \ BADGES_PER_ROW
    End If
    Err.Clear
    On Error GoTo 0
    If rowsN < 1 Then rowsN = 1
    ChartNoteY = BADGE_GRID_Y + rowsN * (BADGE_H + BADGE_GAP_Y) + 20
End Function

Public Sub DrawBadgeShelf(ByVal ws As Worksheet)
    Dim headShp As Shape
    Set headShp = ws.Shapes.AddShape(1, KPI_X0, BADGE_HEAD_Y, 300, 20)
    headShp.Name = "nxd_badge_head"
    headShp.Line.Visible = 0
    headShp.Fill.Visible = 0
    With headShp.TextFrame2
        .TextRange.Text = ChrW(&HD83C) & ChrW(&HDFC5) & " バッジ"
        .TextRange.Font.Size = 11.5
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
        .VerticalAnchor = 3
    End With

    ' 2026-07-28(解説書 §11-11): バッジ表を自前で持たない。
    ' 判定している modStats から受け取る(表を2つ持つと必ずズレる)。
    Dim ids() As String, titles() As String, shorts() As String, conditions() As String
    Dim badgeN As Long
    badgeN = modStats.BadgeCatalog(ids, titles, shorts, conditions)
    If badgeN < 1 Then Exit Sub

    Dim i As Long
    For i = 0 To badgeN - 1
        Dim col As Long: col = i Mod BADGES_PER_ROW
        Dim rowN As Long: rowN = i \ BADGES_PER_ROW
        Dim cardX As Double: cardX = KPI_X0 + col * (BADGE_W + BADGE_GAP_X)
        Dim cardY As Double: cardY = BADGE_GRID_Y + rowN * (BADGE_H + BADGE_GAP_Y)

        Dim dt As String: dt = modStats.BadgeEarnedOn(ids(i))
        Dim earned As Boolean: earned = (LenB(dt) > 0)

        Dim line1 As String, line2 As String
        If earned Then
            line1 = ChrW(&HD83C) & ChrW(&HDFC5) & " " & titles(i)
            line2 = dt
        Else
            line1 = ChrW(&HD83D) & ChrW(&HDD12) & " " & titles(i)
            line2 = modUtil.SafeLeft(conditions(i), 22)
        End If

        DrawBadgeCard ws, i, cardX, cardY, earned, line1, line2
    Next i
End Sub

' 2026-07-31(R11-E M-4): KpiCardと同じ理由でOERN+Err.Clearに統一(1枚の
' 失敗でバッジ棚全体が空になるのを防ぐ)。
Private Sub DrawBadgeCard(ByVal ws As Worksheet, ByVal idx As Long, ByVal x As Double, ByVal y As Double, _
                          ByVal earned As Boolean, ByVal line1 As String, ByVal line2 As String)
    On Error Resume Next
    Dim card As Shape
    Set card = ws.Shapes.AddShape(5, x, y, BADGE_W, BADGE_H)
    If card Is Nothing Then GoTo Done
    card.Name = "nxd_badge_" & idx
    card.Adjustments(1) = 0.14
    card.Line.Weight = 0.75
    card.Shadow.Visible = 0
    If earned Then
        card.Fill.ForeColor.RGB = modUI.UiColor("surface")
        card.Line.ForeColor.RGB = modUI.UiColor("accent")
    Else
        card.Fill.ForeColor.RGB = modUI.UiColor("bg")
        card.Line.ForeColor.RGB = modUI.UiColor("border")
    End If

    Dim body As String: body = line1 & vbLf & line2
    With card.TextFrame2
        .WordWrap = -1
        .MarginLeft = 10: .MarginRight = 10: .MarginTop = 6: .MarginBottom = 6
        .TextRange.Text = body
        .TextRange.Font.Size = 8.5   ' R12-7-4: 7.5pt→8.5pt(a11y監査Med)
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
        .VerticalAnchor = 1
        If .TextRange.Paragraphs.Count >= 1 Then
            With .TextRange.Paragraphs(1).Font
                .Size = 9
                .Bold = -1
                If earned Then
                    .Fill.ForeColor.RGB = modUI.UiColor("text")
                Else
                    .Fill.ForeColor.RGB = modUI.UiColor("muted")
                End If
            End With
        End If
    End With
Done:
    Err.Clear
    On Error GoTo 0
End Sub


' ---- 育ちぐあいの将来枠(クラスタチャートのプレースホルダ) ----

Public Sub DrawChartPlaceholder(ByVal ws As Worksheet)
    ' ナレッジ地図(K-Meansクラスタ可視化)を描く。データ不足時は0が返るので
    ' 案内テキストを出す(いずれも nxd_ 接頭辞=次回描画で一括削除される)。
    Dim drawn As Long
    On Error Resume Next
    drawn = modCluster.DrawClusterMap(ws, KPI_X0, ChartNoteY(), ROW_WIDTH, 250)
    On Error GoTo 0
    If drawn > 0 Then Exit Sub

    Dim note As Shape
    Set note = ws.Shapes.AddShape(1, KPI_X0, ChartNoteY(), ROW_WIDTH, 24)
    note.Name = "nxd_chart_note"
    note.Line.Visible = 0
    note.Fill.Visible = 0
    With note.TextFrame2
        .TextRange.Text = "ナレッジ地図: 資料をもう少し登録すると、似た資料のかたまりが表示されます。"
        .TextRange.Font.Size = 9
        .TextRange.Font.Italic = -1
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
        .VerticalAnchor = 3
    End With
End Sub
