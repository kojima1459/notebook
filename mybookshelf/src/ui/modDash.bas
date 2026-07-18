Attribute VB_Name = "modDash"
Option Explicit

' ============================================================================
' modDash - Nexusダッシュボード画面(KPIカード・EXPバー・バッジ棚)
' ----------------------------------------------------------------------------
' 役割:
'   modAppDef.SH_NEXUS_DASH("Dashboard")シートに、統計サマリー(4枚のKPI
'   カード)・EXP進捗バー・バッジ棚(8種)を描画する「見るだけ」の画面。
'   ナビゲーション配線・シート定数・テーマゲッターは呼び出し側(modApp/
'   modAppDef/modUI)で用意済みのため、本モジュールは純粋な描画/集計のみ
'   を担当する(closed契約: ShowDashboard/OnDashBackToChat/OnDashRefreshの3本)。
'
' 設計判断:
'   ・Shape命名規約: 全て "nxd_" 接頭辞。RemoveShapesByPrefixで毎回全消去
'     してから再構築する(冪等描画・Shape増殖なし)。
'   ・バッジ判定(BadgeCatalog/BadgeDate)・月次判定(IsThisMonthStamp)・
'     分数フォーマット(FormatMinutes)はmodUIDashboardの実証済みロジックを
'     移植(modUIDashboardのヘルパーはPrivateで直接呼べないため)。
'   ・前月比の判定にはmodUIDashboardに無い「先月」判定(IsLastMonthStamp)
'     を追加で持つ(DateAdd("m",-1,Now)基準)。
' ============================================================================

Private Const KPI_CARD_W As Double = 215
Private Const KPI_CARD_H As Double = 92
Private Const KPI_GAP As Double = 14
Private Const KPI_X0 As Double = 20
Private Const KPI_Y0 As Double = 64
Private Const ROW_WIDTH As Double = KPI_CARD_W * 4 + KPI_GAP * 3

Private Const EXPBAR_H As Double = 12
Private Const EXPBAR_Y As Double = KPI_Y0 + KPI_CARD_H + 18
Private Const EXPLABEL_Y As Double = EXPBAR_Y + EXPBAR_H + 4

Private Const BADGE_W As Double = 208
Private Const BADGE_H As Double = 54
Private Const BADGE_GAP_X As Double = 12
Private Const BADGE_GAP_Y As Double = 10
Private Const BADGES_PER_ROW As Long = 4
Private Const BADGE_HEAD_Y As Double = EXPLABEL_Y + 22
Private Const BADGE_GRID_Y As Double = BADGE_HEAD_Y + 24

Private Const CHART_NOTE_Y As Double = BADGE_GRID_Y + 2 * (BADGE_H + BADGE_GAP_Y) + 20

Private Const HEADER_TITLE_Y As Double = 14
Private Const HEADER_SUB_Y As Double = 40
Private Const HEADER_BTN_Y As Double = 14
Private Const HEADER_BTN_H As Double = 26
Private Const BTN_CHAT_W As Double = 118
Private Const BTN_REFRESH_W As Double = 96
Private Const BTN_GAP As Double = 10
Private Const CONTENT_RIGHT As Double = KPI_X0 + ROW_WIDTH
Private Const BTN_REFRESH_X As Double = CONTENT_RIGHT - BTN_REFRESH_W
Private Const BTN_CHAT_X As Double = BTN_REFRESH_X - BTN_GAP - BTN_CHAT_W
Private Const BTN_EXPORT_W As Double = 128
Private Const BTN_EXPORT_X As Double = BTN_CHAT_X - BTN_GAP - BTN_EXPORT_W

Private Const MINUTES_PER_SELFSOLVE As Long = 15   ' modStatsの換算値と同じ(先方はPrivateのため複製)

' ----------------------------------------------------------------------------
' ShowDashboard - シートを取得/生成し、描画してSPA遷移する(公開エントリ)
' ----------------------------------------------------------------------------
Public Sub ShowDashboard()
    Dim ws As Worksheet
    Set ws = GetOrCreateDashSheet()
    If ws Is Nothing Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False

    DrawDashboard ws

    ws.Visible = -1   ' xlSheetVisible
    ws.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayWorkbookTabs = False
    On Error GoTo 0

    Application.ScreenUpdating = True
    Exit Sub

Fail:
    On Error Resume Next
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' OnDashBackToChat - チャット画面(Nexusシート)へ戻る
' ----------------------------------------------------------------------------
Public Sub OnDashBackToChat()
    On Error Resume Next
    ThisWorkbook.Worksheets("Nexus").Activate
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' OnDashRefresh - 既存のダッシュボードシートを再描画する
' ----------------------------------------------------------------------------
Public Sub OnDashRefresh()
    Dim ws As Worksheet
    Set ws = GetDashSheet()
    If ws Is Nothing Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    DrawDashboard ws
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    On Error Resume Next
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 内部: 画面全体の描画(冪等・毎回全再構築)
' ----------------------------------------------------------------------------
Private Sub DrawDashboard(ByVal ws As Worksheet)
    RemoveShapesByPrefix ws, "nxd_"
    ws.Cells.Clear
    ws.Cells.Font.Name = "Yu Gothic UI"
    ws.Cells.Interior.Color = modUI.UiColor("bg")
    ws.Columns("A:T").ColumnWidth = 9

    DrawHeader ws
    DrawKpiRow ws
    DrawExpBar ws
    DrawBadgeShelf ws
    DrawChartPlaceholder ws
    modUI.FreezeShapePlacement ws   ' 全Shape(クラスタ円含む)を絶対配置に固定
End Sub

' ---- ヘッダー ----

Private Sub DrawHeader(ByVal ws As Worksheet)
    Dim titleShp As Shape
    Set titleShp = ws.Shapes.AddShape(1, KPI_X0, HEADER_TITLE_Y, 400, 26)
    titleShp.Name = "nxd_title"
    titleShp.Line.Visible = 0
    titleShp.Fill.Visible = 0
    With titleShp.TextFrame2
        .TextRange.Text = ChrW(&H1F4CA) & " ダッシュボード"
        .TextRange.Font.Size = 15
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
        .VerticalAnchor = 3
    End With

    Dim subShp As Shape
    Set subShp = ws.Shapes.AddShape(1, KPI_X0, HEADER_SUB_Y, 400, 18)
    subShp.Name = "nxd_subtitle"
    subShp.Line.Visible = 0
    subShp.Fill.Visible = 0
    With subShp.TextFrame2
        .TextRange.Text = "あなたの本棚とAI活用の記録"
        .TextRange.Font.Size = 9.5
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
        .VerticalAnchor = 3
    End With

    DrawHeaderButton ws, "nxd_btn_export", BTN_EXPORT_X, HEADER_BTN_Y, BTN_EXPORT_W, HEADER_BTN_H, _
        ChrW(&H1F4E5) & " 分析用ログ出力", "modAnalytics.ExportAnalyticsCsv"
    DrawHeaderButton ws, "nxd_btn_chat", BTN_CHAT_X, HEADER_BTN_Y, BTN_CHAT_W, HEADER_BTN_H, _
        ChrW(&H1F4AC) & " チャットへ", "modDash.OnDashBackToChat"
    DrawHeaderButton ws, "nxd_btn_refresh", BTN_REFRESH_X, HEADER_BTN_Y, BTN_REFRESH_W, HEADER_BTN_H, _
        ChrW(&H1F504) & " 更新", "modDash.OnDashRefresh"
End Sub

Private Sub DrawHeaderButton(ByVal ws As Worksheet, ByVal shapeName As String, ByVal x As Double, _
                             ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                             ByVal caption As String, ByVal handlerName As String)
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(5, x, y, w, h)
    btn.Name = shapeName
    btn.Adjustments(1) = 0.35
    btn.Fill.ForeColor.RGB = modUI.UiColor("surface")
    btn.Line.ForeColor.RGB = modUI.UiColor("border")
    btn.Line.Weight = 0.75
    btn.Shadow.Visible = 0
    With btn.TextFrame2
        .TextRange.Text = caption
        .TextRange.Font.Size = 9
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 2: .MarginRight = 2
    End With
    btn.OnAction = handlerName
End Sub

' ---- KPIカード4枚 ----

Private Sub DrawKpiRow(ByVal ws As Worksheet)
    ' Card0: 取り戻した時間(+前月比)
    Dim savedMinutes As Long: savedMinutes = SafeSavedMinutes()
    Dim isUp As Boolean
    Dim deltaText As String: deltaText = SavedTimeDeltaLabel(isUp)
    Dim deltaColor As Long
    If isUp Then
        deltaColor = modUI.UiColor("accent")
    Else
        deltaColor = modUI.UiColor("muted")
    End If
    DrawKpiCard ws, 0, KPI_X0, KPI_Y0, KPI_CARD_W, KPI_CARD_H, _
        "取り戻した時間", FormatMinutes(savedMinutes), deltaText, deltaColor

    ' Card1: 登録ナレッジ数
    Dim ingestTotal As Long: ingestTotal = SafeGetStat("ingest_files_total")
    DrawKpiCard ws, 1, KPI_X0 + (KPI_CARD_W + KPI_GAP), KPI_Y0, KPI_CARD_W, KPI_CARD_H, _
        "登録ナレッジ数", ingestTotal & "件", "あなたが登録した資料"

    ' Card2: 蔵書チャンク数
    Dim totalChunks As Long: totalChunks = SafeTotalChunks()
    Dim shelfMax As Long: shelfMax = SafeShelfMax()
    Dim ratio As Double
    If shelfMax > 0 Then
        ratio = CDbl(totalChunks) / CDbl(shelfMax)
    Else
        ratio = 0
    End If
    DrawKpiCard ws, 2, KPI_X0 + 2 * (KPI_CARD_W + KPI_GAP), KPI_Y0, KPI_CARD_W, KPI_CARD_H, _
        "蔵書チャンク数", totalChunks & " / " & shelfMax, UsageBarText(ratio)

    ' Card3: レベル
    Dim lv As Long: lv = SafeLevel()
    Dim expTotalV As Long: expTotalV = SafeExpTotal()
    Dim remain As Long: remain = SafeExpFloorForLevel(lv + 1) - expTotalV
    If remain < 0 Then remain = 0
    DrawKpiCard ws, 3, KPI_X0 + 3 * (KPI_CARD_W + KPI_GAP), KPI_Y0, KPI_CARD_W, KPI_CARD_H, _
        "レベル", "Lv." & lv, "EXP " & expTotalV & " ・ 次まで" & remain & "EXP"
End Sub

' 単一Shape=1カード(caption/value/label の3段を1テキストで結合し段別書式)。
'   段1=caption(小・muted)、段2=value(大・太・text)、段3=label(小・指定色)。
Private Sub DrawKpiCard(ByVal ws As Worksheet, ByVal idx As Long, ByVal x As Double, ByVal y As Double, _
                        ByVal cardW As Double, ByVal cardH As Double, ByVal captionText As String, _
                        ByVal valueText As String, ByVal labelText As String, _
                        Optional ByVal labelColor As Long = -1)
    Dim card As Shape
    Set card = ws.Shapes.AddShape(5, x, y, cardW, cardH)
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

    Dim body As String: body = captionText & vbLf & valueText & vbLf & labelText

    With card.TextFrame2
        .WordWrap = -1
        .MarginLeft = 12: .MarginRight = 10: .MarginTop = 9: .MarginBottom = 8
        .TextRange.Text = body
        .TextRange.Font.Size = 9
        .TextRange.Font.Fill.ForeColor.RGB = finalLabelColor
        .VerticalAnchor = 1
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
    End With
End Sub

' ---- EXP進捗バー ----

Private Sub DrawExpBar(ByVal ws As Worksheet)
    Dim trackW As Double: trackW = ROW_WIDTH
    Dim prog As Double: prog = SafeLevelProgress()
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

    Dim lv As Long: lv = SafeLevel()
    Dim remain As Long: remain = SafeExpFloorForLevel(lv + 1) - SafeExpTotal()
    If remain < 0 Then remain = 0

    Dim lbl As Shape
    Set lbl = ws.Shapes.AddShape(1, KPI_X0, EXPLABEL_Y, trackW, 16)
    lbl.Name = "nxd_exp_label"
    lbl.Line.Visible = 0
    lbl.Fill.Visible = 0
    With lbl.TextFrame2
        .TextRange.Text = "Lv." & (lv + 1) & " まであと " & remain & " EXP"
        .TextRange.Font.Size = 8.5
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
        .VerticalAnchor = 3
    End With
End Sub

' ---- バッジ棚(8種・4列×2行) ----

Private Sub DrawBadgeShelf(ByVal ws As Worksheet)
    Dim headShp As Shape
    Set headShp = ws.Shapes.AddShape(1, KPI_X0, BADGE_HEAD_Y, 300, 20)
    headShp.Name = "nxd_badge_head"
    headShp.Line.Visible = 0
    headShp.Fill.Visible = 0
    With headShp.TextFrame2
        .TextRange.Text = ChrW(&H1F3C5) & " バッジ"
        .TextRange.Font.Size = 11.5
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
        .VerticalAnchor = 3
    End With

    Dim ids() As String, titles() As String, conditions() As String
    BadgeCatalog ids, titles, conditions

    Dim i As Long
    For i = 0 To UBound(ids)
        Dim col As Long: col = i Mod BADGES_PER_ROW
        Dim rowN As Long: rowN = i \ BADGES_PER_ROW
        Dim cardX As Double: cardX = KPI_X0 + col * (BADGE_W + BADGE_GAP_X)
        Dim cardY As Double: cardY = BADGE_GRID_Y + rowN * (BADGE_H + BADGE_GAP_Y)

        Dim dt As String: dt = BadgeDate(ids(i))
        Dim earned As Boolean: earned = (LenB(dt) > 0)

        Dim line1 As String, line2 As String
        If earned Then
            line1 = ChrW(&H1F3C5) & " " & titles(i)
            line2 = dt
        Else
            line1 = ChrW(&H1F512) & " " & titles(i)
            line2 = modUtil.SafeLeft(conditions(i), 22)
        End If

        DrawBadgeCard ws, i, cardX, cardY, earned, line1, line2
    Next i
End Sub

Private Sub DrawBadgeCard(ByVal ws As Worksheet, ByVal idx As Long, ByVal x As Double, ByVal y As Double, _
                          ByVal earned As Boolean, ByVal line1 As String, ByVal line2 As String)
    Dim card As Shape
    Set card = ws.Shapes.AddShape(5, x, y, BADGE_W, BADGE_H)
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
        .MarginLeft = 8: .MarginRight = 8: .MarginTop = 6: .MarginBottom = 6
        .TextRange.Text = body
        .TextRange.Font.Size = 7.5
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
        .VerticalAnchor = 1
        With .TextRange.Paragraphs(1).Font
            .Size = 9
            .Bold = -1
            If earned Then
                .Fill.ForeColor.RGB = modUI.UiColor("text")
            Else
                .Fill.ForeColor.RGB = modUI.UiColor("muted")
            End If
        End With
    End With
End Sub

' MASTER_SPEC §9のバッジ定義(modUIDashboard.BadgeCatalogを移植)。
Private Sub BadgeCatalog(ByRef ids() As String, ByRef titles() As String, ByRef conditions() As String)
    ids = Split("first_ingest,shelf10,shelf30,first_pack_out,first_pack_in,solve10,solve50,streak7", ",")
    titles = Split("初めての取込,本棚10冊,本棚30冊,初パック共有,初パック取込,自己解決10件,自己解決50件,7日連続利用", ",")
    conditions = Split( _
        "資料を1つ本棚に追加すると獲得|" & _
        "資料を10冊集めると獲得|" & _
        "資料を30冊集めると獲得|" & _
        "資料をパックとして誰かに渡すと獲得|" & _
        "誰かのパックを取り込むと獲得|" & _
        "🟢解決したが10回になると獲得|" & _
        "🟢解決したが50回になると獲得|" & _
        "7日連続で使うと獲得", "|")
End Sub

' badge:<id> の獲得日文字列(未獲得は"")。my_statsを直接読む(modUIDashboardと同じ理由)。
Private Function BadgeDate(ByVal badgeId As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_STATS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim key As String
    key = "badge:" & badgeId

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim i As Long
    For i = 2 To lastRow
        If StrComp(CStr(ws.Cells(i, 1).Value), key, vbTextCompare) = 0 Then
            BadgeDate = Trim$(CStr(ws.Cells(i, 2).Value))
            Exit Function
        End If
    Next i
End Function

' ---- 育ちぐあいの将来枠(クラスタチャートのプレースホルダ) ----

Private Sub DrawChartPlaceholder(ByVal ws As Worksheet)
    ' ナレッジ地図(K-Meansクラスタ可視化)を描く。データ不足時は0が返るので
    ' 案内テキストを出す(いずれも nxd_ 接頭辞=次回描画で一括削除される)。
    Dim drawn As Long
    On Error Resume Next
    drawn = modCluster.DrawClusterMap(ws, KPI_X0, CHART_NOTE_Y, ROW_WIDTH, 250)
    On Error GoTo 0
    If drawn > 0 Then Exit Sub

    Dim note As Shape
    Set note = ws.Shapes.AddShape(1, KPI_X0, CHART_NOTE_Y, ROW_WIDTH, 24)
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

' ----------------------------------------------------------------------------
' 内部: 前月比(取り戻した時間)の集計
' ----------------------------------------------------------------------------

' 今月/先月の"feedback_green"件数から前月比の表示文字列を作る。
'   isUp: ▲(上昇/新記録)ならTrue、▼/—ならFalse(アクセント色の切替に使用)。
Private Function SavedTimeDeltaLabel(ByRef isUp As Boolean) As String
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
Private Function CountUsageEvent(ByVal eventName As String, ByVal wantLastMonth As Boolean) As Long
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
Private Function IsThisMonthStamp(ByVal v As Variant) As Boolean
    If IsDate(v) Then
        Dim d As Date: d = CDate(v)
        IsThisMonthStamp = (Year(d) = Year(Now) And Month(d) = Month(Now))
    Else
        IsThisMonthStamp = (Left$(CStr(v), 7) = Format$(Now, "yyyy-mm"))
    End If
End Function

' timestampセルが「先月」かどうか(IsThisMonthStampと同じ考え方)。
Private Function IsLastMonthStamp(ByVal v As Variant) As Boolean
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
Private Function UsageBarText(ByVal ratio As Double) As String
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
Private Function FormatMinutes(ByVal minutes As Long) As String
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
Private Function SafeGetStat(ByVal key As String) As Long
    On Error Resume Next
    SafeGetStat = modStats.GetStat(key)
    On Error GoTo 0
End Function

Private Function SafeSavedMinutes() As Long
    On Error Resume Next
    SafeSavedMinutes = modStats.SavedMinutesEstimate()
    On Error GoTo 0
End Function

Private Function SafeTotalChunks() As Long
    On Error Resume Next
    SafeTotalChunks = modShelf.TotalChunks()
    On Error GoTo 0
End Function

Private Function SafeShelfMax() As Long
    On Error Resume Next
    SafeShelfMax = modConfig.GetLong("shelf_max_chunks", 20000)
    On Error GoTo 0
    If SafeShelfMax <= 0 Then SafeShelfMax = 20000
End Function

Private Function SafeLevel() As Long
    On Error Resume Next
    SafeLevel = modStats.Level()
    On Error GoTo 0
    If SafeLevel < 1 Then SafeLevel = 1
End Function

Private Function SafeExpTotal() As Long
    On Error Resume Next
    SafeExpTotal = modStats.ExpTotal()
    On Error GoTo 0
End Function

Private Function SafeExpFloorForLevel(ByVal lv As Long) As Long
    On Error Resume Next
    SafeExpFloorForLevel = modStats.ExpFloorForLevel(lv)
    On Error GoTo 0
End Function

Private Function SafeLevelProgress() As Double
    On Error Resume Next
    SafeLevelProgress = modStats.LevelProgress()
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' 内部: シート取得/生成・Shape一括削除
' ----------------------------------------------------------------------------
Private Function GetDashSheet() As Worksheet
    On Error Resume Next
    Set GetDashSheet = ThisWorkbook.Worksheets(modAppDef.SH_NEXUS_DASH)
    On Error GoTo 0
End Function

Private Function GetOrCreateDashSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = GetDashSheet()
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = modAppDef.SH_NEXUS_DASH
        On Error GoTo 0
    End If
    Set GetOrCreateDashSheet = ws
    Exit Function
Fail:
    Set GetOrCreateDashSheet = Nothing
End Function

Private Sub RemoveShapesByPrefix(ByVal ws As Worksheet, ByVal prefix As String)
    Dim names() As String
    ReDim names(0 To ws.Shapes.Count)
    Dim n As Long: n = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, Len(prefix)) = prefix Then
            names(n) = shp.Name
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
