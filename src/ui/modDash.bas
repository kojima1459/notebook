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
'   ・バッジ判定(BadgeCatalog/BadgeDate)・月次判定(modDashStat.IsThisMonthStamp)・
'     分数フォーマット(modDashStat.FormatMinutes)はmodUIDashboardの実証済みロジックを
'     移植(modUIDashboardのヘルパーはPrivateで直接呼べないため)。
'   ・前月比の判定にはmodUIDashboardに無い「先月」判定(modDashStat.IsLastMonthStamp)
'     を追加で持つ(DateAdd("m",-1,Now)基準)。
' ============================================================================

Private Const KPI_CARD_W As Double = 215
Private Const KPI_CARD_H As Double = 92
Private Const KPI_GAP As Double = 14
Private Const KPI_X0 As Double = 20
' 2026-07-31(R7 A-1): ヘッダー帯(48pt)+サブタイトルの下から本文を始める。
Private Const KPI_Y0 As Double = 80
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

' 2026-07-28(解説書 §11-11): バッジが8種から12種に増えた。
' 以前は「2行ぶん」と決め打ちしていたため、増えた行がチャート枠と重なる。
' VBAの定数式では関数を呼べないので、位置は実行時に ChartNoteY() で求める。
Private Const BADGE_ROWS_FALLBACK As Long = 3

' 2026-07-31(R7 A-1): ヘッダーは他画面(Hub/チャット/ナレッジ)と同じ形にする。
'   ・濃色バー(全幅)+タイトル
'   ・ナビピル(← Hub / 💬チャット / 📚ナレッジと本棚)
'   ・共通ユーティリティ(🔄更新 / 📥ログ / 🎨着せ替え / 🚪保存して終わる)
' 旧実装はボタンを CONTENT_RIGHT(=KPI_X0+ROW_WIDTH=922pt)へ右寄せしていた。
' このアプリの画面幅は約600pt(チャットヘッダーの実測597pt・Hubの帯626pt)
' なので、右寄せしたボタンは【画面外】に置かれていた。実機報告
' 「ダッシュボード上部が真っ白・ナビが無い」の正体はこれで、ボタンは
' 描かれていたが誰にも見えていなかった。操作系は必ず HDR_CTL_COLS 列の
' 内側(=他画面と同じ幅)に置き、はみ出す分は modChrome.FlowLeft が段を
' 増やして受ける(画面外へ描かない)。
Private Const HDR_BAR_H As Double = 48
Private Const HDR_PILL_H As Double = 26
Private Const HDR_PILL_GAP As Double = 5
Private Const HDR_PILL_PITCH As Double = 9
Private Const HDR_PILL_PAD As Double = 14
Private Const HDR_PILL_MIN As Double = 40
Private Const HDR_TITLE_RESERVE As Double = 140   ' タイトル「📊 ダッシュボード」用
Private Const HDR_CTL_COLS As String = "A1:L1"    ' 操作系を置いてよい範囲(約600pt)
Private Const HDR_ITEMS As Long = 7
Private Const HEADER_SUB_Y As Double = 54


Private Const ADMIN_ROW_H As Double = 26
Private Const ADMIN_MAX_ROWS As Long = 12

Private mAdminExclNames() As String
Private mAdminExclCount As Long
Private mDashStep As String   ' DrawDashboard失敗箇所の特定用(Fail:から参照)

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
    ActiveWindow.DisplayWorkbookTabs = False
    On Error GoTo 0
    ' R7 A-2: 全画面/数式バー/罫線/スクロール位置/等倍をまとめて自己修復する
    ' (「閉じる→キャンセル」で壊れた表示が、この画面を開くだけで戻る)。
    On Error Resume Next
    modUI.EnsureAppView
    On Error GoTo 0

    Application.ScreenUpdating = True
    Exit Sub

Fail:
    Dim failNum As Long, failDesc As String
    failNum = Err.Number: failDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup0
FailCleanup0:
    On Error Resume Next
    Application.ScreenUpdating = True
    modLog.LogError "E0801", "modDash.ShowDashboard", "[" & mDashStep & "] " & failDesc, failNum
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' OnDashBackToChat - チャット画面(Nexusシート)へ戻る
' ----------------------------------------------------------------------------
Public Sub OnDashBackToChat()
    If modUiLock.BlockIfIngesting() Then Exit Sub   ' R7 B-2
    modUI.GoToNexus "modDash.OnDashBackToChat"
End Sub

' ----------------------------------------------------------------------------
' OnDashRefresh - 既存のダッシュボードシートを再描画する
' ----------------------------------------------------------------------------
Public Sub OnDashRefresh()
    If modUiLock.BlockIfIngesting() Then Exit Sub   ' R7 B-2
    Dim ws As Worksheet
    Set ws = GetDashSheet()
    If ws Is Nothing Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    DrawDashboard ws
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    Dim failNum As Long, failDesc As String
    failNum = Err.Number: failDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup2
FailCleanup2:
    On Error Resume Next
    Application.ScreenUpdating = True
    modLog.LogError "E0801", "modDash.OnDashRefresh", "[" & mDashStep & "] " & failDesc, failNum
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' OnDashRestore - 管理者用「復帰」ボタンのハンドラ(組織的除外を解除する)
' ----------------------------------------------------------------------------
Public Sub OnDashRestore()
    Dim callerName As String
    On Error Resume Next
    callerName = CStr(Application.Caller)
    On Error GoTo 0
    If Left$(callerName, Len("nxd_adm_btn_")) <> "nxd_adm_btn_" Then Exit Sub

    Dim idx As Long
    idx = CLng(Val(Mid$(callerName, Len("nxd_adm_btn_") + 1)))
    If idx < 0 Or idx >= mAdminExclCount Then Exit Sub

    Dim src As String
    src = mAdminExclNames(idx)

    Dim answer As Long
    answer = MsgBox("『" & src & "』の組織的除外を解除しますか?(全ユーザーの検索に復帰します)", _
                     vbYesNo + vbQuestion, modAppDef.APP_NAME)
    If answer <> vbYes Then Exit Sub

    If modP2P.ClearNoise(src) Then
        On Error Resume Next
        modP2P.CollectNoiseVotes True   ' ローカルの除外状態を今すぐ再計算(一覧から消す)
        On Error GoTo 0
        MsgBox "『" & src & "』を復帰しました。次回以降の同期で全ユーザーに反映されます。", _
               vbInformation, modAppDef.APP_NAME
        OnDashRefresh
    Else
        MsgBox "解除に失敗しました(管理者権限または共有フォルダをご確認ください)。", _
               vbExclamation, modAppDef.APP_NAME
    End If
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

    mDashStep = "DrawHeader": DrawHeader ws
    mDashStep = "DrawKpiRow": DrawKpiRow ws
    mDashStep = "DrawExpBar": DrawExpBar ws
    mDashStep = "DrawBadgeShelf": DrawBadgeShelf ws
    mDashStep = "DrawChartPlaceholder": DrawChartPlaceholder ws
    mDashStep = "DrawAdminSection": DrawAdminSection ws
    mDashStep = "FreezeShapePlacement"
    modUI.FreezeShapePlacement ws   ' 全Shape(クラスタ円含む)を絶対配置に固定
    modSkin.BeautifyAll ws          ' フォント統一(Yu Gothic UI)+固定クロムに柔らかい影
End Sub

' ---- ヘッダー ----

Private Sub DrawHeader(ByVal ws As Worksheet)
    Dim L As Double, W As Double
    L = ws.Range("A1").Left
    W = ws.Range("A1:T1").Width                      ' 帯そのものは画面いっぱい

    Dim caps(0 To HDR_ITEMS - 1) As String
    Dim acts(0 To HDR_ITEMS - 1) As String
    Dim nms(0 To HDR_ITEMS - 1) As String
    Dim wds(0 To HDR_ITEMS - 1) As Double
    HeaderSpec caps, acts, nms, wds

    ' 配置の算数は modChrome に任せる(枠外へ描く置き方が存在しない形にする)。
    Dim xs() As Double, rws() As Long, uws() As Double
    Dim rowN As Long
    rowN = modChrome.FlowLeft(wds, HDR_ITEMS, L + HDR_TITLE_RESERVE, _
                              L + ws.Range(HDR_CTL_COLS).Width - 8, HDR_PILL_GAP, _
                              xs, rws, uws)
    If rowN < 1 Then rowN = 1
    Dim barH As Double: barH = HDR_BAR_H + (rowN - 1) * (HDR_PILL_H + 4)

    Dim hdr As Shape
    Set hdr = ws.Shapes.AddShape(5, L, 0, W, barH)
    hdr.Name = "nxd_hdr"
    hdr.Line.Visible = 0
    hdr.Adjustments(1) = 0.02
    hdr.Fill.ForeColor.RGB = modUI.UiColor("sidebar")
    modSkin.ApplyHeaderDepth hdr                      ' §9: 濃紺の2色グラデーション
    With hdr.TextFrame2
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCCA) & " ダッシュボード"
        .TextRange.Font.Size = 14
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .MarginLeft = 14
        .VerticalAnchor = 3
    End With

    Dim i As Long
    For i = 0 To HDR_ITEMS - 1
        HeaderPill ws, nms(i), caps(i), acts(i), xs(i), _
                   (HDR_BAR_H - HDR_PILL_H) / 2 + rws(i) * (HDR_PILL_H + 4), uws(i)
    Next i

    Dim subShp As Shape
    Set subShp = ws.Shapes.AddShape(1, KPI_X0, barH + HEADER_SUB_Y - HDR_BAR_H, 400, 18)
    subShp.Name = "nxd_subtitle"
    subShp.Line.Visible = 0
    subShp.Fill.Visible = 0
    With subShp.TextFrame2
        .TextRange.Text = "あなたの本棚とAI活用の記録"
        .TextRange.Font.Size = 9.5
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
        .VerticalAnchor = 3
    End With
End Sub

' ヘッダーに載せるものの唯一の定義。ナビ3枚→ユーティリティ4枚の順で、
' 並びとアイコンはHub・チャット画面に合わせる(いちばん右が🚪)。
' 「💬 チャットへ」が2つあった問題(旧ボタンとナビの重複)はナビ側へ一本化した。
Private Sub HeaderSpec(ByRef caps() As String, ByRef acts() As String, _
                       ByRef nms() As String, ByRef wds() As Double)
    caps(0) = ChrW(&H2190) & " Hub":                          nms(0) = "nxd_nav_hub"
    acts(0) = "modKnowledge.OnBackHub"
    caps(1) = ChrW(&HD83D) & ChrW(&HDCAC) & " チャット":      nms(1) = "nxd_nav_chat"
    acts(1) = "modDash.OnDashBackToChat"
    caps(2) = ChrW(&HD83D) & ChrW(&HDCDA) & " ナレッジと本棚": nms(2) = "nxd_nav_vault"
    acts(2) = "modHub.OnGoVault"
    caps(3) = ChrW(&HD83D) & ChrW(&HDD04) & " 更新":          nms(3) = "nxd_btn_refresh"
    acts(3) = "modDash.OnDashRefresh"
    caps(4) = ChrW(&HD83D) & ChrW(&HDCE5) & " ログ":          nms(4) = "nxd_btn_export"
    acts(4) = "modAnalytics.ExportAnalyticsCsv"
    caps(5) = ChrW(&HD83C) & ChrW(&HDFA8) & " 着せ替え":      nms(5) = "nxd_btn_skin"
    acts(5) = "modHub.OnThemeToggle"
    caps(6) = ChrW(&HD83D) & ChrW(&HDEAA) & " 終了":          nms(6) = "nxd_btn_exit"
    acts(6) = "modApp.OnSaveAndExit"

    Dim i As Long
    For i = 0 To HDR_ITEMS - 1
        wds(i) = modChrome.PillWidth(caps(i), HDR_PILL_PITCH, HDR_PILL_PAD, HDR_PILL_MIN)
    Next i
End Sub

' 帯の上に置くピル1個(1個の失敗で残りを道連れにしない)。
Private Sub HeaderPill(ByVal ws As Worksheet, ByVal shapeName As String, _
                       ByVal caption As String, ByVal handlerName As String, _
                       ByVal x As Double, ByVal y As Double, ByVal w As Double)
    On Error Resume Next
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(5, x, y, w, HDR_PILL_H)
    If btn Is Nothing Then Exit Sub
    btn.Name = shapeName
    btn.Adjustments(1) = 0.35
    btn.Line.Visible = 0
    btn.Shadow.Visible = 0
    btn.Fill.ForeColor.RGB = modUI.UiColor("sidebarActive")
    With btn.TextFrame2
        .WordWrap = -1
        .TextRange.Text = caption
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
    End With
    btn.OnAction = handlerName
    On Error GoTo 0
End Sub

' ---- KPIカード4枚 ----

Private Sub DrawKpiRow(ByVal ws As Worksheet)
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
Private Function ChartNoteY() As Double
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

Private Sub DrawBadgeShelf(ByVal ws As Worksheet)
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
        .MarginLeft = 10: .MarginRight = 10: .MarginTop = 6: .MarginBottom = 6
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


' ---- 育ちぐあいの将来枠(クラスタチャートのプレースホルダ) ----

Private Sub DrawChartPlaceholder(ByVal ws As Worksheet)
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

' ---- 管理者専用: 組織的除外の管理セクション(非管理者には何も描かない) ----

Private Sub DrawAdminSection(ByVal ws As Worksheet)
    If Not modP2P.IsAdmin() Then Exit Sub

    Dim topY As Double: topY = ChartNoteY() + 270   ' クラスタ地図(~250pt)の下に確保

    Dim headShp As Shape
    Set headShp = ws.Shapes.AddShape(1, KPI_X0, topY, ROW_WIDTH, 20)
    headShp.Name = "nxd_adm_head"
    headShp.Line.Visible = 0
    headShp.Fill.Visible = 0
    With headShp.TextFrame2
        .TextRange.Text = ChrW(&H26A0) & " 組織的除外の管理(管理者)"
        .TextRange.Font.Size = 12
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
        .VerticalAnchor = 3
    End With

    Dim ex As Object
    On Error Resume Next
    Set ex = modStats.GlobalExcludedSources()
    On Error GoTo 0
    If ex Is Nothing Then Exit Sub

    Dim exCount As Long: exCount = ex.Count
    If exCount > 0 Then
        ReDim mAdminExclNames(0 To exCount - 1)
    Else
        ReDim mAdminExclNames(0 To 0)
    End If
    mAdminExclCount = 0

    Dim k As Variant
    For Each k In ex.Keys
        mAdminExclNames(mAdminExclCount) = CStr(k)
        mAdminExclCount = mAdminExclCount + 1
    Next k

    If mAdminExclCount = 0 Then
        Dim emptyShp As Shape
        Set emptyShp = ws.Shapes.AddShape(1, KPI_X0, topY + 28, ROW_WIDTH, 20)
        emptyShp.Name = "nxd_adm_empty"
        emptyShp.Line.Visible = 0
        emptyShp.Fill.Visible = 0
        With emptyShp.TextFrame2
            .TextRange.Text = "現在、組織的に除外されているナレッジはありません。"
            .TextRange.Font.Size = 9
            .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
            .VerticalAnchor = 3
        End With
        Exit Sub
    End If

    Dim shownRows As Long: shownRows = mAdminExclCount
    If shownRows > ADMIN_MAX_ROWS Then shownRows = ADMIN_MAX_ROWS

    Dim i As Long
    For i = 0 To shownRows - 1
        Dim rowY As Double: rowY = topY + 28 + i * ADMIN_ROW_H

        Dim lbl As Shape
        Set lbl = ws.Shapes.AddShape(1, KPI_X0, rowY, ROW_WIDTH - 100, ADMIN_ROW_H - 4)
        lbl.Name = "nxd_adm_lbl_" & i
        lbl.Line.Visible = 0
        lbl.Fill.Visible = 0
        With lbl.TextFrame2
            .WordWrap = -1
            .MarginLeft = 8: .MarginRight = 8
            .TextRange.Text = modUtil.SafeLeft(mAdminExclNames(i), 60)
            .TextRange.Font.Size = 9
            .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
            .VerticalAnchor = 3
        End With

        Dim btn As Shape
        Set btn = ws.Shapes.AddShape(5, KPI_X0 + ROW_WIDTH - 80, rowY, 72, 22)
        btn.Name = "nxd_adm_btn_" & i
        btn.Adjustments(1) = 0.3
        btn.Line.Visible = 0
        btn.Shadow.Visible = 0
        btn.Fill.ForeColor.RGB = modUI.UiColor("primary")
        With btn.TextFrame2
            .WordWrap = -1
            .MarginLeft = 8: .MarginRight = 8
            .TextRange.Text = ChrW(&H21A9) & " 復帰"
            .TextRange.Font.Size = 8.5
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = 2
            .VerticalAnchor = 3
        End With
        btn.OnAction = "modDash.OnDashRestore"
    Next i
End Sub


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
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup16
FailCleanup16:
    If Not ws Is Nothing Then
        On Error Resume Next
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        On Error GoTo 0
    End If
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
