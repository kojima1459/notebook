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

' 2026-07-31(R7 A-1): ヘッダーは他画面(Hub/チャット/ナレッジ)と同じ形にする。
'   ・濃色バー(全幅)+タイトル
'   ・ナビピル(← Hub / 💬チャット / 📚ナレッジと本棚)
'   ・共通ユーティリティ(🔄更新 / 📥ログ / 🎨着せ替え / 🚪保存して終わる)
' 旧実装はボタンを CONTENT_RIGHT(=modDashStat.KPI_X0+modDashStat.ROW_WIDTH=922pt)へ右寄せしていた。
' このアプリの画面幅は約600pt(チャットヘッダーの実測597pt・Hubの帯626pt)
' なので、右寄せしたボタンは【画面外】に置かれていた。実機報告
' 「ダッシュボード上部が真っ白・ナビが無い」の正体はこれで、ボタンは
' 描かれていたが誰にも見えていなかった。操作系は必ず実可視幅
' (R11-B: HDR_CTL_COLSの決め打ち"A1:L1"を廃止しmodChrome.BarWidth+
' modUIMain.ViewportWidthへ)の内側に置き、はみ出す分はFlowLeftが段を
' 増やして受ける(画面外へ描かない)。
Private Const HDR_BAR_H As Double = 48
Private Const HDR_PILL_H As Double = 26
Private Const HDR_PILL_GAP As Double = 5
Private Const HDR_PILL_PITCH As Double = 9
Private Const HDR_PILL_PAD As Double = 14
Private Const HDR_PILL_MIN As Double = 40
Private Const HDR_TITLE_RESERVE As Double = 140   ' タイトル「📊 ダッシュボード」用
' 2026-07-31(R11-F2): 「?」ヘルプを足して8個(監査1 L-3。ダッシュボードから
' ヘルプへ行く手段が無かった。R11-Eでは容量不足で保留になっていた)。
Private Const HDR_ITEMS As Long = 8
Private Const HEADER_SUB_Y As Double = 54


Private Const ADMIN_ROW_H As Double = 26
Private Const ADMIN_MAX_ROWS As Long = 12

' R18-3a/3b: この画面が使うセル範囲。本文はすべてpt座標のShapeで積むため
' 行高の既定がフォント依存でぶれると「pt→行」の換算ができない。行高を
' DASH_ROW_H に固定する。DASH_ROWS は行の【上限】(120行=1,800pt)で、
' 実際に塗る行数は内容の実下端から決める(R19-1b。従来は常に120行=1,800ptを
' 塗っていたが内容は400〜700ptしかなく、空塗りが最大の画面だった)。
Private Const DASH_ROWS As Long = 120
Private Const DASH_ROW_H As Double = 15
' R19-1b: 列は A:J(内容幅ぶん)+ 吸収列K。旧実装は A:T=約1,020pt を9ptで
' 敷き詰めていたが、内容の実右端は590ptしかなく、右へ430ptの空塗りが
' 残っていた(実機報告「特にダッシュボード」の正体)。
Private Const DASH_BAND As String = "A1:K1"
Private Const DASH_PAD_COL As String = "K"

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

    ' R21-S1(実機第8報⑦の独立欠陥その2): 「描く→活性化→表示状態」を
    ' 【活性化→表示状態→描く】へ反転する。R20-1b で全画面化だけは前へ出した
    ' が、タブ・罫線・見出し・水平スクロールバーは依然として描画の【後】に
    ' 確定していた。とくに水平スクロールバーは可視高を約24pt食うため、
    ' 描画時に測った窓高が常に24pt過大になり、境界が必ず窓を超えていた。
    ' 測ってよいのは modViewport2.EnsureViewState を通った後だけ。
    On Error Resume Next
    modUI.EnsureAppView
    On Error GoTo Fail

    ws.Visible = -1   ' xlSheetVisible
    If modUI.ActivateSheetRobust(ws, "modDash.ShowDashboard") Then
        On Error Resume Next
        modViewport2.EnsureViewState ws
        On Error GoTo Fail
    Else
        modUI.RestoreExcelUI
    End If

    Dim vw0 As Double, vh0 As Double
    modViewport2.MarkView vw0, vh0
    modUiLock.AlertsOff   ' R25-1a-1: DrawDashboard内のMerge警告を出さない(対はAlertsOn)
    DrawDashboard ws
    modUiLock.AlertsOn

    ' R7 A-2: 表示状態の自己修復(他画面から引き継いだ崩れを戻す)。
    On Error Resume Next
    modUI.EnsureAppView
    On Error GoTo 0

    Application.ScreenUpdating = True
    ' R21-S1の保険: 描画中に窓が動いていたら1回だけ組み直す(ワンショット)。
    modViewport2.ReflowIfMoved vw0, vh0
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
    modUiLock.AlertsOn   ' R25-1a-1: DrawDashboard中断時もAlertsOffの対を必ず戻す
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
' 「?」= 使い方。ヘルプカードはNexus(チャット)シート上に描かれるため、
' 先にチャットへ移ってから modHelp.OnHelpClick を呼ぶ。OnHelpClick 自身が
' modUiLock を取るので、ここではロックを取らない(2026-07-31 R11-F2)。
Public Sub OnHelp()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    On Error Resume Next
    modUI.GoToNexus "modDash.OnHelp"
    If Err.Number <> 0 Then modLog.LogError "E0801", "modDash.OnHelp", Err.Description, Err.Number
    On Error GoTo 0
    modHelp.OnHelpClick
End Sub

Public Sub OnDashRefresh()
    If modUiLock.BlockIfIngesting() Then Exit Sub   ' R7 B-2
    Dim ws As Worksheet
    Set ws = GetDashSheet()
    If ws Is Nothing Then Exit Sub

    ' 2026-07-31(レビュー R8 F9): ビーコン集計には10分のTTLが入っている。
    ' 利用者が明示的に「更新」を押したときだけはTTLを無視して取り直す
    '(押しても値が変わらないのは、ボタンが壊れているのと区別が付かない)。
    On Error Resume Next
    modBoard.ForceRefreshBoard
    On Error GoTo 0

    On Error GoTo Fail
    Application.ScreenUpdating = False
    modUiLock.AlertsOff   ' R25-1a-1: DrawDashboard内のMerge警告を出さない(対はAlertsOn)
    DrawDashboard ws
    modUiLock.AlertsOn
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
    modUiLock.AlertsOn   ' R25-1a-1: DrawDashboard中断時もAlertsOffの対を必ず戻す
    Application.ScreenUpdating = True
    modLog.LogError "E0801", "modDash.OnDashRefresh", "[" & mDashStep & "] " & failDesc, failNum
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' OnDashRestore - 管理者用「復帰」ボタンのハンドラ(組織的除外を解除する)
' ----------------------------------------------------------------------------
Public Sub OnDashRestore()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    Dim callerName As String
    On Error Resume Next
    callerName = CStr(Application.Caller)
    On Error GoTo 0
    If Left$(callerName, Len("nxd_adm_btn_")) <> "nxd_adm_btn_" Then
        modLog.LogUsage "caller_mismatch", "modDash.OnDashRestore", callerName
        Exit Sub
    End If

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
    ' R18-3a(実機第5報②): 全域(ws.Cells)への書式はUsedRangeをシート最大へ
    ' 膨らませる(この画面は行高を一切設定していないのに「下に無限へ
    ' スクロールできる」と報告された。調査agent2 §1.3)。実使用範囲だけに
    ' 当てる。行高も明示して幾何を確定させる(pt→行の換算をここで固定する)。
    ' R25-2a(FA-R25-2a): 吸収列K(DASH_PAD_COL)もここでリセットする。以前は
    ' A:Jのみを固定していたため、K(吸収列)は前回描画で確定した幅を引き継いだ
    ' ままFitBandToViewportへ渡り、目標が「前回の帯幅」を自己参照する形になって
    ' いた(SAFE_MARGIN 2ptが再描画のたびに複利で効き、実測1010→988の縮小と
    ' 一致)。K込みでリセットすることでチャット(modUI.bas:100のC:K/L)と同じ
    ' 「固定側は毎回既知の初期値に戻す」形にし、ラチェットを断つ。
    ws.Columns("A:K").ColumnWidth = 9
    ' R27 F2-2: 帯(A:K)の右外を幅1にして、右端の部分可視列がキャップから
    ' 丸ごと差し引かれる(=帯が可視幅より短くなる)のを断つ。Hub実証済み。
    modChrome.SetupDashColumns ws
    ' 余りをK列に吸わせて A:K の合計を可視幅ぴったりにする。内容の実右端
    ' (最小版面590pt)より狭い窓では内容側を優先する(ボタンが境界の外に出ない)。
    modViewport.FitBandToViewport ws, DASH_BAND, DASH_PAD_COL, MinContentRightX()
    ' R20-1b(層1): 合わせ終えた帯の【実幅】からKPI/バッジのカード幅を決め直す。
    ' これより後に描くもの(ヘッダー・KPI・EXPバー・バッジ・管理者行)は全て
    ' modDashStat.KpiCardW()/RowWidth() を見るので、本文がまるごと窓幅へ追随する。
    modDashStat.SetBandWidth modViewport.ContentRight(ws, DASH_BAND, 0) - ws.Range("A1").Left
    ' R21-S5: 窓高に収まらないときだけ縦を圧縮する(KPIカード高・バッジ高)。
    ' 係数は modViewport2 が1本で持ち、実際に掛けるのは modViewport2.SY() だけ。
    ' R21H F3: needを固定/可変に分けて渡す(まとめて渡すと分母が薄まり圧縮不足になる)。
    modViewport2.SetScaleY modViewport2.CompressFactor( _
        modViewport.ViewportHeight(), modDashStat.NeedYFixed(), modDashStat.NeedYVariable())
    ' R20-1d(層2): 行高を明示する範囲は40行まで。120行(1,800pt=3画面ぶん)を
    ' 毎回「使用済み」にしていたのが、下へ延々スクロールできる状態の正体。
    ' 実下端が確定した ApplyDashScrollBound が、それ以深を既定へ戻す。
    ws.Rows("1:40").RowHeight = DASH_ROW_H
    ' 書式(フォント・塗り)は描き終えてから実下端ぶんだけ当てる
    ' (ApplyDashScrollBound)。全域・広域の書式はUsedRangeを膨らませる。

    mDashStep = "DrawHeader": DrawHeader ws
    mDashStep = "DrawKpiRow": modDashStat.DrawKpiRow ws
    mDashStep = "DrawExpBar": modDashStat.DrawExpBar ws
    mDashStep = "DrawBadgeShelf": modDashStat.DrawBadgeShelf ws
    mDashStep = "DrawAdminSection": DrawAdminSection ws
    mDashStep = "ApplyScrollBound": ApplyDashScrollBound ws
    mDashStep = "FreezeShapePlacement"
    modUI.FreezeShapePlacement ws   ' 全Shapeを絶対配置に固定
    modSkin.BeautifyAll ws          ' フォント統一(Yu Gothic UI)+固定クロムに柔らかい影
End Sub

' 本文は全てpt座標のShapeで積まれ、内容(バッジ件数・管理者一覧の件数)で
' 下端が変わるため、固定の行数では決められない。描き終えた実物のShapeから
' 下端を実測し、塗り・ScrollArea をそこまで(+24pt、最低1画面)にする
' (R18-3b → R19-1b。右端は吸収列Kで可視幅に一致しているので実測不要)。
Private Sub ApplyDashScrollBound(ByVal ws As Worksheet)
    Dim bottomY As Double
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Top + shp.Height > bottomY Then bottomY = shp.Top + shp.Height
    Next shp
    On Error GoTo 0
    If bottomY < 200 Then bottomY = 200

    Dim bnd As String, paintAddr As String
    bnd = modViewport.BoundAddr(ws, DASH_PAD_COL, bottomY, DASH_ROWS, paintAddr)
    On Error Resume Next
    ' R21H F6: 塗りはpaintAddr(切り上げ=部分行を含む)。ScrollAreaはbnd
    ' (切り下げ)のまま=窓下端の未塗り帯(darkテーマの白帯)を防ぐ。
    ws.Range(paintAddr).Font.Name = "Yu Gothic UI"
    ws.Range(paintAddr).Interior.Color = modUI.UiColor("bg")
    On Error GoTo 0
    modViewport.ApplyScrollBound ws, bnd
    ' R20-1d: 境界の【下】に行高カスタムを1行も残さない(受け入れ基準)。
    ' bnd は必ず A1 起点なので、行数がそのまま下端行になる。
    On Error Resume Next
    modViewport.ResetRowsBelow ws, ws.Range(bnd).Rows.Count + 1, DASH_ROWS
    ' R27 F2-3(埋め草): 境界は切り下げなので最終行の下端は必ず窓高より上に
    ' 残る。その差だけ最終行を高くして、塗りの下端を窓下端へ届かせる
    ' (行数は増やさない=縦スクロールを生き返らせない)。
    modViewport2.PadRowToWindow ws, ws.Range(bnd).Rows.Count
    On Error GoTo 0
    ' R21-S7: 窓幅/可視幅/帯実幅/中身右端/境界下端の5値観測点。
    modViewport2.LogFit ws, "dash", DASH_BAND, bnd
End Sub

' 帯を可視幅へ合わせるときの下限(pt)。KPIカードが【最小幅】130ptで4枚並ぶ
' 版面より狭くはできない(それ以上詰めるとカードが画面外へ出る)。
' R20-1b: 実際の版面幅は帯が決まってから modDashStat.RowWidth() が返す。
' ここで現在値(RowWidth)を返してはいけない ―― 前回の広い窓で決めた幅が
' 次回の狭い窓で下限として効き、帯が可視幅を超えてしまう。
Private Function MinContentRightX() As Double
    MinContentRightX = modDashStat.KPI_X0 * 2 + modDashStat.ROW_WIDTH
End Function

' ---- ヘッダー ----

Private Sub DrawHeader(ByVal ws As Worksheet)
    Dim L As Double, W As Double
    L = ws.Range("A1").Left
    ' R19-1b: 帯は可視幅いっぱい、操作系の右端はその内側8pt。どちらも
    ' modViewport.ContentRight 1本から取る(画面ごとの別式を作らない)。
    W = modViewport.ContentRight(ws, DASH_BAND, 0) - L

    Dim caps(0 To HDR_ITEMS - 1) As String
    Dim acts(0 To HDR_ITEMS - 1) As String
    Dim nms(0 To HDR_ITEMS - 1) As String
    Dim wds(0 To HDR_ITEMS - 1) As Double
    HeaderSpec caps, acts, nms, wds

    ' 配置の算数は modChrome に任せる(枠外へ描く置き方が存在しない形にする)。
    Dim xs() As Double, rws() As Long, uws() As Double
    Dim rowN As Long
    rowN = modChrome.FlowLeft(wds, HDR_ITEMS, L + HDR_TITLE_RESERVE, _
                              modViewport.ContentRight(ws, DASH_BAND, 8), HDR_PILL_GAP, _
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

    ' R21-S3: 版面のX原点は RowX0() 1本(センタリング廃止で全要素が同じ左端に
    ' 揃う)。ここだけ KPI_X0 の定数を直に見ていたのを他と同じ口へ寄せる。
    Dim subShp As Shape
    Set subShp = ws.Shapes.AddShape(1, modDashStat.RowX0(), barH + HEADER_SUB_Y - HDR_BAR_H, 400, 18)
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
    ' ヘルプカードはNexus(チャット)シート上に描く作りなので、押すとチャットへ
    ' 移ってから開く(modHub.OnHelp/modKnowledge.OnHelp と同型)。
    caps(7) = ChrW(&H2753):                                   nms(7) = "nxd_btn_help"
    acts(7) = "modDash.OnHelp"

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
    ' R24-2b: 白枠0.75ptは視認不可と実機判定。塗り白+濃緑文字+濃緑枠へ反転。
    btn.Line.Visible = -1
    btn.Line.ForeColor.RGB = RGB(1, 77, 68)
    btn.Line.Weight = 0.75
    btn.Shadow.Visible = 0
    btn.Fill.ForeColor.RGB = RGB(255, 255, 255)
    With btn.TextFrame2
        .WordWrap = -1
        .TextRange.Text = caption
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(1, 77, 68)
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
    End With
    btn.OnAction = handlerName
    On Error GoTo 0
End Sub

' ---- 管理者専用: 組織的除外の管理セクション(非管理者には何も描かない) ----

Private Sub DrawAdminSection(ByVal ws As Worksheet)
    If Not modP2P.IsAdmin() Then Exit Sub

    ' R18-4: ここは ChartNoteY() + 270(クラスタ地図~250ptぶんの確保)だった。
    ' 地図を撤去したので跡地を詰める。+20 はバッジ棚との間の余白のみ。
    Dim topY As Double: topY = modDashStat.ChartNoteY() + 20

    ' R20H FA-7: KPI_X0(左寄せ固定の定数)のままだと、帯を広げてKPI行/EXP
    ' バーが中央寄せへ動いたときに管理者行だけ左端に取り残され版面が割れて
    ' 見えていた。KPI行と同じRowX0()(帯内中央寄せの左端X)へ揃える。
    Dim headShp As Shape
    Set headShp = ws.Shapes.AddShape(1, modDashStat.RowX0(), topY, modDashStat.RowWidth(), 20)
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
        Set emptyShp = ws.Shapes.AddShape(1, modDashStat.RowX0(), topY + 28, modDashStat.RowWidth(), 20)
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
        Set lbl = ws.Shapes.AddShape(1, modDashStat.RowX0(), rowY, modDashStat.RowWidth() - 100, ADMIN_ROW_H - 4)
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
        Set btn = ws.Shapes.AddShape(5, modDashStat.RowX0() + modDashStat.RowWidth() - 80, rowY, 72, 22)
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
