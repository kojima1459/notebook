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

' R13-7d: 自己解決1件あたりの節約時間(分)は modP2PIo.MinutesPerSelfsolve()
' (config minutes_per_selfsolve、既定15)へ統合した。modStats/modBoardと
' 3重複していたPrivate Constはここも削除する(憲章§4-5)。

' 2026-07-31(R7実装後の追加是正・発見事項1): 215pt×4枚+間隔だと902ptになり、
' 実画面幅(他画面の実測で約597〜650pt)を大幅に超えて3・4枚目が画面外に出る。
' A-1ではヘッダー帯だけを幅内に収めたが、本文のカード自体は放置されていた。
' 4枚とも1画面に収まる130ptへ縮小する(バッジも同じ行幅に揃える)。
' KPI_CARD_W/KPI_GAP は ROW_WIDTH の定数式が参照するため Public。
' (LibreOffice のコンパイルは Public Const が Private Const を参照すると
'  応答不能になる。実Excelでは私有のままでも通るが、モード2の構文チェックを
'  通すために公開側へ揃える。2026-07-31 R11-F1で実測。)
' 2026-08-06(R20-1b・実機第7報⑦の層1): 130pt は【最小】幅になった。R19 までは
' 帯(セル列 A:K)だけを可視幅へ合わせていて、本文のKPIカードは4枚とも130pt
' 固定のまま。窓を広げるほど「帯だけが伸びて、中身は左の590ptで止まっている」
' 画面になり、右に数百ptの空塗りが残っていた(実機写真の主症状)。帯の実幅から
' カード幅を出し直し、本文を窓幅へ追随させる。上限220ptは、1枚の数字カードが
' それ以上広がっても情報密度が下がるだけ(4枚で920pt=一般的な窓幅の上限)という
' 判断。KPI_CARD_W/KPI_GAP は ROW_WIDTH の定数式が参照するため Public。
Public Const KPI_CARD_W As Double = 130
Public Const KPI_CARD_MAX_W As Double = 220
Private Const KPI_CARD_H As Double = 92
Public Const KPI_GAP As Double = 10
Public Const KPI_X0 As Double = 20
' 2026-07-31(R7 A-1): ヘッダー帯(48pt)+サブタイトルの下から本文を始める。
Private Const KPI_Y0 As Double = 80
' ROW_WIDTH は「最小版面」(=130×4+10×3=550pt。左右余白40ptを足した帯590ptが
' 帯幅の下限)。帯を可視幅へ合わせるときの下限としてだけ使い、
' 実際の版面幅は RowWidth() を見る(R20-1b)。
Public Const ROW_WIDTH As Double = KPI_CARD_W * 4 + KPI_GAP * 3

' 直近に FitBandToViewport で確定した帯幅から決めたカード幅(pt)。
' 0 のあいだは最小幅(KPI_CARD_W)で描く=従来と同じ絵になる。
Private mCardW As Double
' 追加D-1(2026-08-06・センタリング): 版面(RowWidth)の左端X。0のあいだは
' KPI_X0(左寄せ・従来と同じ絵)。SetBandWidthが帯幅から中央寄せの値へ更新する。
Private mX0 As Double

Private Const EXPBAR_H As Double = 12
Private Const EXPBAR_Y As Double = KPI_Y0 + KPI_CARD_H + 18
Private Const EXPLABEL_Y As Double = EXPBAR_Y + EXPBAR_H + 4

' 2026-07-31: KPIカードと同じ行幅(130*4+10*3+X0*2=590pt)に揃える。
' R20-1b: 幅は KpiCardW()(帯幅から決まる可変値)へ移した。BADGE_W は
' 「KPIカードと同じ幅に揃える」という約束を示す名前として残していたが、
' 2箇所に数字を持つと必ずズレるので削除する(憲章§4-5)。
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
' 版面(KPIカード4枚+バッジカード4枚)の幅を帯幅から決める(R20-1b)
' ----------------------------------------------------------------------------

' CardWidthFor - 帯幅 bandW のときの1枚あたりのカード幅(pt)。純関数
'   (ゴールデン対象。境界: 帯590pt未満は130で頭打ち / 950pt超は220で頭打ち)。
'   左右の余白 KPI_X0 を2つ、カード間の隙間 KPI_GAP を3つ引いて4等分する。
Public Function CardWidthFor(ByVal bandW As Double) As Double
    CardWidthFor = modViewport.ClampD((bandW - 2 * KPI_X0 - 3 * KPI_GAP) / 4, _
                                      KPI_CARD_W, KPI_CARD_MAX_W)
End Function

' SetBandWidth - 描画の直前に、帯の【実幅】(FitBandToViewport 後)を渡す。
'   ここより後に走る DrawKpiRow/DrawExpBar/DrawBadgeShelf と modDash の
'   ヘッダー・管理者行が、全て同じ CardW()/RowWidth() を見る。
' 追加D-1: 同時に版面の左端X(RowX0)も帯幅から決め直す(センタリング)。
'   modDash.MinContentRightX() はこのSubの【前】に呼ばれるため、そちらは
'   従来どおりKPI_X0(定数=20)のまま据え置く(mX0を参照させると前回描画の
'   古い値を読み、狭い窓で帯を縮められなくなる=RowWidth()staleness警告と
'   同型の回帰を招くため、意図的に触れていない)。
Public Sub SetBandWidth(ByVal bandW As Double)
    mCardW = CardWidthFor(bandW)
    mX0 = CenterX0(bandW, RowWidth())
End Sub

' KpiCardW - 現在のカード幅(pt)。SetBandWidth 前は最小幅。
Public Function KpiCardW() As Double
    KpiCardW = mCardW
    If KpiCardW < KPI_CARD_W Then KpiCardW = KPI_CARD_W
    If KpiCardW > KPI_CARD_MAX_W Then KpiCardW = KPI_CARD_MAX_W
End Function

' RowWidth - 現在の版面幅(pt)。カード4枚+隙間3つ。
Public Function RowWidth() As Double
    RowWidth = KpiCardW() * 4 + KPI_GAP * 3
End Function

' CenterX0(追加D-1) - 帯内中央寄せの左端X(純関数・ゴールデン対象)。帯幅
'   bandWが版面幅rowWidthより広いときだけ余白を等分する(下限minX0=20pt)。
'   bandW<=rowWidthの狭い窓ではminX0のまま(左寄せ・従来どおりの絵)。
'   左右対称の余白は「デザイン」に見え、左だけの片寄り余白は「バグ」に
'   見える、という趣旨(広窓でKPI行/EXPバー/バッジ群が帯内で中央に来る)。
Public Function CenterX0(ByVal bandW As Double, ByVal rowW As Double, _
                         Optional ByVal minX0 As Double = 20) As Double
    Dim c As Double: c = (bandW - rowW) / 2
    If c < minX0 Then c = minX0
    CenterX0 = c
End Function

' RowX0 - 版面(KPI行/EXPバー/バッジ群)の左端X。SetBandWidth前・未初期化
'   時はKPI_X0(=20・左寄せ)を返す。
Public Function RowX0() As Double
    RowX0 = mX0
    If RowX0 < KPI_X0 Then RowX0 = KPI_X0
End Function

' ----------------------------------------------------------------------------
' 内部: 前月比(節約した時間)の集計
' ----------------------------------------------------------------------------

' 今月/先月の"feedback_green"件数から前月比の表示文字列を作る。
'   isUp: ▲(上昇/新記録)ならTrue、▼/―ならFalse(アクセント色の切替に使用)。
Public Function SavedTimeDeltaLabel(ByRef isUp As Boolean) As String
    Dim thisMonthCount As Long: thisMonthCount = CountUsageEvent("feedback_green", False)
    Dim lastMonthCount As Long: lastMonthCount = CountUsageEvent("feedback_green", True)
    Dim perSolve As Long: perSolve = modP2PIo.MinutesPerSelfsolve()
    Dim thisMin As Long: thisMin = thisMonthCount * perSolve
    Dim lastMin As Long: lastMin = lastMonthCount * perSolve

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
' ダッシュボード本文の描画(KPIカード4枚 / 経験値バー / バッジ棚)。
'   ※「クラスタ地図」は R18-4 で撤去済み(2026-08-05 R18H FB-3 / A-L10:
'     残骸のコメントが「まだ在る」と誤読させるため文言から除去した)。
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
    ' Card0: 節約した時間(+前月比)
    Dim savedMinutes As Long: savedMinutes = modDashStat.SafeSavedMinutes()
    Dim isUp As Boolean
    Dim deltaText As String: deltaText = modDashStat.SavedTimeDeltaLabel(isUp)
    Dim deltaColor As Long
    If isUp Then
        deltaColor = modUI.UiColor("accent")
    Else
        deltaColor = modUI.UiColor("muted")
    End If
    Dim cw As Double: cw = KpiCardW()                 ' R20-1b: 帯幅から決めた実カード幅
    Dim x0 As Double: x0 = RowX0()                     ' 追加D-1: 帯内中央寄せの左端X
    DrawKpiCard ws, 0, x0, KPI_Y0, cw, KPI_CARD_H, _
        "節約した時間", modDashStat.FormatMinutes(savedMinutes), deltaText, deltaColor

    ' Card1: 登録ナレッジ数
    Dim ingestTotal As Long: ingestTotal = modDashStat.SafeGetStat("ingest_files_total")
    DrawKpiCard ws, 1, x0 + (cw + KPI_GAP), KPI_Y0, cw, KPI_CARD_H, _
        "登録ナレッジ数", ingestTotal & "件", "あなたが登録した資料"

    ' Card2: 資料の分量(2026-08-06 R20H FA-16: 生ジャーゴン「蔵書チャンク数」を平易化)
    Dim totalChunks As Long: totalChunks = modDashStat.SafeTotalChunks()
    Dim shelfMax As Long: shelfMax = modDashStat.SafeShelfMax()
    Dim ratio As Double
    If shelfMax > 0 Then
        ratio = CDbl(totalChunks) / CDbl(shelfMax)
    Else
        ratio = 0
    End If
    DrawKpiCard ws, 2, x0 + 2 * (cw + KPI_GAP), KPI_Y0, cw, KPI_CARD_H, _
        "資料の分量", totalChunks & " / " & shelfMax, modDashStat.UsageBarText(ratio)

    ' Card3: レベル
    Dim lv As Long: lv = modDashStat.SafeLevel()
    Dim expTotalV As Long: expTotalV = modDashStat.SafeExpTotal()
    Dim remain As Long: remain = modDashStat.SafeExpFloorForLevel(lv + 1) - expTotalV
    If remain < 0 Then remain = 0
    DrawKpiCard ws, 3, x0 + 3 * (cw + KPI_GAP), KPI_Y0, cw, KPI_CARD_H, _
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
    Dim trackW As Double: trackW = RowWidth()   ' R20-1b: 版面幅に追随
    Dim x0 As Double: x0 = RowX0()               ' 追加D-1: 帯内中央寄せの左端X
    Dim prog As Double: prog = modDashStat.SafeLevelProgress()
    If prog < 0 Then prog = 0
    If prog > 1 Then prog = 1

    Dim fillW As Double: fillW = trackW * prog
    If prog > 0 And fillW < 2 Then fillW = 2

    Dim bgBar As Shape
    Set bgBar = ws.Shapes.AddShape(5, x0, EXPBAR_Y, trackW, EXPBAR_H)
    bgBar.Name = "nxd_expbar_bg"
    bgBar.Adjustments(1) = 0.5
    bgBar.Fill.ForeColor.RGB = modUI.UiColor("border")
    bgBar.Line.Visible = 0
    bgBar.Shadow.Visible = 0

    If fillW > 0 Then
        Dim fgBar As Shape
        Set fgBar = ws.Shapes.AddShape(5, x0, EXPBAR_Y, fillW, EXPBAR_H)
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
    Set lbl = ws.Shapes.AddShape(1, x0, EXPLABEL_Y, trackW, 16)
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
    Dim x0 As Double: x0 = RowX0()   ' 追加D-1: 帯内中央寄せの左端X
    Dim headShp As Shape
    Set headShp = ws.Shapes.AddShape(1, x0, BADGE_HEAD_Y, 300, 20)
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
        Dim cardX As Double: cardX = x0 + col * (KpiCardW() + BADGE_GAP_X)
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
    Set card = ws.Shapes.AddShape(5, x, y, KpiCardW(), BADGE_H)
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

' ---- ナレッジ地図(撤去済み) ----
' 2026-08-05(R18-4・実機第5報③): DrawChartPlaceholder(modCluster.DrawClusterMap
' の呼び出しと、点が足りないときの案内テキスト)をここから削除した。跡地は
' 詰めてあり、バッジ棚の下端(ChartNoteY)の直下に管理者セクションが続く
' (modDash.DrawAdminSection の「+270」も同時に外した)。
