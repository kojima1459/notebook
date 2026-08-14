Attribute VB_Name = "modViewport2"
Option Explicit

' ============================================================================
' modViewport2 - 「測ってよいタイミング」と「窓に収める算数」(2026-08-07 R21-1)
' ----------------------------------------------------------------------------
' なぜ新設したか(憲章§4-6の容量分割):
'   R21-1(実機第8報⑦)で足すロジック ―― 表示状態の単一化(EnsureViewState)・
'   スクロールバー幅の校正・Fitの事後検証・弾性グリッド・窓高適応圧縮・
'   再フィット経路の穴埋め ―― を全て modViewport(21,524字)へ入れると
'   30,000字の上限に届く。幾何の【原始的な算数】は従来どおり modViewport が
'   持ち、こちらは「いつ測ってよいか」と「窓に収めるための調整」を持つ。
'
' 5回目を出さないための設計の要点(調査A班の数値トレースに対応):
'   (1) 測定の前に表示状態(全画面/罫線/見出し/タブ/水平スクロールバー)を
'       必ず最終形へ置く。UsableWidth/UsableHeight はこれら全てで変わるため、
'       「描いてから全画面にする」と初回だけ古い窓で幾何が決まる。
'   (2) 帯の目標幅は「可視セル幅 − 安全余裕」。可視セル幅は
'       UsableWidth − 縦スクロールバー幅で、バー幅は決め打ちせず実測から校正する。
'   (3) 収まらないときだけ縦を圧縮する(係数は1本・SY()経由でのみ適用)。
'   全て On Error Resume Next 配下(失敗しても画面を落とさない)。
'
' 依存: modViewport / modUIMain / modLog / modKnowledge / modUIShelf /
'       modUINexusDraw / modUI(いずれも同じUI層)。
' 純関数(HScrollNeeded / SbWidthFrom / ViewMoved / CompressFactor /
'   GridColsFor / GridCardW / RightGapExceeds)は他を呼ばず、LibreOffice の
'   純ロジックテスト(modTestsPure22)で境界を固定する。
' ============================================================================

' 水平スクロールバーを安全弁として出す窓幅の境目(pt)。これより狭い窓では
' 帯を可視幅まで詰めても内容の最小幅に負ける(#30の安全弁)ので、横へ
' 逃げられる手段を残す。これ以上の窓では「帯≤可視幅」が構造的に成立するので
' 出さない(出ているだけで可視高が約24pt減り、境界計算が窓を超える原因になる)。
Private Const NARROW_W As Double = 625

' 帯の目標幅に持たせる安全余裕(pt)。ColumnWidth の丸めで1pt前後は必ず
' ぶれるので、ぴったりを狙うと「1ptだけ超える」が常態化する。
Private Const SAFE_MARGIN As Double = 2

' 縦スクロールバー幅の既定(pt)。校正できない端末でだけ使う。
Private Const SB_FALLBACK As Double = 12
' 校正値として信じてよい差の上限(pt)。VisibleRange は列単位で切り上がるため、
' これを超える差は「バー幅」ではなく「セル1個ぶんの端数」。
Private Const SB_MAX As Double = 30

' 描画の前後で可視サイズがこれ以上動いたら、1回だけ組み直す(pt)。
Private Const VIEW_TOL As Double = 4

' 窓高適応圧縮の下限。これ以上詰めると文字が箱からはみ出す。
Private Const MIN_SCALE As Double = 0.78

' 埋め草(R27 F2-3)の上限/下限(pt)。詳細は PadRowDelta の直前を参照。
'   MAX: 1行ぶん(既定18pt前後)を大きく超える差は構造問題なので埋めない。
'   MIN: RowHeight の丸めで毎描画わずかに伸び続けるのを止める足切り。
Private Const PAD_ROW_MAX As Double = 48
Private Const PAD_ROW_MIN As Double = 1

' チャットのバンド下端の下に残す余裕行(R30 W1)。0にすると境界ちょうどで行が
' 終わり、ホイールの1刻みで必ず「行の無い所」へ出る。3行=54ptはホイール
' 1〜2刻みぶんの遊び。詳細は FitChatRows の直前の見出しコメント参照。
Private Const CHAT_ROW_SLACK As Long = 3
' R30 F7: modSkin.ExtendChatBandの縮小分岐は容量都合でこの定数を参照できず、
' リテラル3(CHAT_ROW_SLACKと同値)を直書きしている。値を変えるときは両方直す。

' 校正済みの縦スクロールバー幅(pt)。0=未校正。EnsureViewState で捨てる。
Private mSbW As Double

' 窓高適応圧縮の係数(1=圧縮なし)。SY() だけがこれを掛ける。
Private mScaleY As Double

' 描画末尾のワンショット再描画中フラグ(無限ループ禁止・S1の保険)。
Private mReflowGuard As Boolean

' ----------------------------------------------------------------------------
' S1: 測定の単一化と順序固定
' ----------------------------------------------------------------------------
' EnsureViewState - 表示状態を最終形へ置く。【これを通した後だけ】測ってよい。
'   全画面の描画エントリ(modHub/modDash/modVaultGallery/modUIShelf)が
'   Activate の直後・幾何を決める前に1回だけ呼ぶ。冪等。
'   自分のブックが前面でない/そのシートが前面でないときは何もしない
'   (他人の窓の設定を書き換えない。EnsureAppView と同じ作法)。
Public Sub EnsureViewState(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If Not (ThisWorkbook.ActiveSheet Is ws) Then Exit Sub

    If Not Application.DisplayFullScreen Then Application.DisplayFullScreen = True
    If Application.DisplayFormulaBar Then Application.DisplayFormulaBar = False

    Dim win As Object
    Set win = ActiveWindow
    If win Is Nothing Then Exit Sub
    If win.DisplayGridlines Then win.DisplayGridlines = False
    If win.DisplayHeadings Then win.DisplayHeadings = False
    If win.DisplayWorkbookTabs Then win.DisplayWorkbookTabs = False
    If win.Zoom <> 100 Then win.Zoom = 100
    ' 水平スクロールバーは可視【高】を約24pt食う。ここで最終形にしてから
    ' 測らないと、境界計算がその24ptぶんだけ必ず窓を超える(実機第8報⑦の
    ' 独立欠陥その2)。狭い窓でだけ安全弁として残す。
    Dim wantH As Boolean: wantH = HScrollNeeded(modUIMain.ViewportWidth())
    If win.DisplayHorizontalScrollBar <> wantH Then win.DisplayHorizontalScrollBar = wantH
    If win.ScrollColumn <> 1 Then win.ScrollColumn = 1
    If win.ScrollRow <> 1 Then win.ScrollRow = 1

    mSbW = 0            ' 表示が変わった=バー幅の校正はやり直し
    On Error GoTo 0
End Sub

' HScrollNeeded - 水平スクロールバーを出すべきか。純関数(ゴールデン対象)。
'   #30(操作系が画面外)の安全弁は狭い窓にだけ要る。広い窓で出しっぱなしに
'   すると、可視高だけが減って「境界>窓高」を作り続ける。
Public Function HScrollNeeded(ByVal viewportW As Double) As Boolean
    HScrollNeeded = (viewportW < NARROW_W)
End Function

' WantHScroll - 今の窓幅での答え(modUI.EnsureAppView の一律ONを条件化する口)。
Public Function WantHScroll() As Boolean
    WantHScroll = HScrollNeeded(modUIMain.ViewportWidth())
End Function

' ----------------------------------------------------------------------------
' S2: 帯は必ず可視幅以下(目標幅の単一情報源)
' ----------------------------------------------------------------------------
' SbWidthFrom - 縦スクロールバーの実幅(pt)。純関数(ゴールデン対象)。
'   UsableWidth(バーを含む内側の幅)と VisibleRange.Width(セルとして
'   見えている幅)の差がバー幅。ただし VisibleRange は列単位で切り上がるので、
'   差が負や大きすぎる値になったら実測を信じずに fallback へ倒す。
Public Function SbWidthFrom(ByVal usableW As Double, ByVal visibleW As Double, _
                            ByVal fallback As Double) As Double
    SbWidthFrom = fallback
    If usableW <= 0 Then Exit Function
    If visibleW <= 0 Then Exit Function
    Dim d As Double: d = usableW - visibleW
    If d < 0 Then Exit Function
    If d > SB_MAX Then Exit Function
    SbWidthFrom = d
End Function

' ScrollbarW - 校正済みの縦スクロールバー幅(pt)。1セッション/1表示状態に1回。
Public Function ScrollbarW() As Double
    ScrollbarW = mSbW
    If ScrollbarW > 0 Then Exit Function
    Dim uw As Double, vw As Double
    On Error Resume Next
    If ActiveWorkbook Is ThisWorkbook Then
        uw = ActiveWindow.UsableWidth
        vw = ActiveWindow.VisibleRange.Width
    End If
    On Error GoTo 0
    ScrollbarW = SbWidthFrom(uw, vw, SB_FALLBACK)
    mSbW = ScrollbarW
End Function

' VisibleCellW - セルを置ける実可視幅(pt)= 可視幅 − 縦スクロールバー幅。
Public Function VisibleCellW() As Double
    VisibleCellW = modUIMain.ViewportWidth() - ScrollbarW()
    If VisibleCellW < 1 Then VisibleCellW = 1
End Function

' FitTarget - 帯(A列～吸収列)の合計幅の目標(pt)。単一情報源。
' R21H F8(敵対的レビュー確定): ActiveWindow.VisibleRangeは右端が半分だけ
' 見えている列も【全幅】で数えるため、SbWidthFrom(usableW,visibleW,…)の
' 差dがその半端ぶんだけ小さく出る環境がある(d∈(0,30]の許容内に収まって
' しまうと既定値へ倒れず、そのまま採用されてしまう)。ScrollbarWが小さく
' 出るとVisibleCellW/FitTargetは実際より大きくなり、帯が可視幅を最大で
' 半端列1つぶん(実測13pt級)超える。FitTargetを「VisibleRangeの完全表示
' 部分(右端列が食み出していれば除いた幅)」で頭打ちする。
Public Function FitTarget() As Double
    FitTarget = VisibleCellW() - SAFE_MARGIN
    If FitTarget < 1 Then FitTarget = 1
    Dim capW As Double: capW = FullyVisibleWidth()
    If capW > 0 Then
        Dim capTarget As Double: capTarget = capW - SAFE_MARGIN
        If capTarget < 1 Then capTarget = 1
        If FitTarget > capTarget Then FitTarget = capTarget
    End If
End Function

' FullyVisibleWidth - VisibleRangeのうち右端列が完全に収まっているところ
'   までの幅(pt)。右端列がUsableWidthを食み出していなければVisibleRange.
'   Widthそのまま、食み出していればその列ぶんを除いて返す(測れなければ0=
'   FitTarget側は頭打ちを適用しない)。
Private Function FullyVisibleWidth() As Double
    On Error Resume Next
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Function
    Dim vr As Range: Set vr = ActiveWindow.VisibleRange
    If vr Is Nothing Then Exit Function
    Dim uw As Double: uw = ActiveWindow.UsableWidth
    Dim w As Double: w = vr.Width
    Dim lastCol As Range: Set lastCol = vr.Columns(vr.Columns.Count)
    If lastCol.Left + lastCol.Width > uw Then w = w - lastCol.Width
    FullyVisibleWidth = w
    On Error GoTo 0
End Function

' FitVerify - Fit の事後検証。帯の実幅が target を超えていたら、超過ぶんだけ
'   吸収列を詰め直す(最大2回)。換算は modViewport.PadUnitsRefine(アフィン
'   2点補正)を流用する ―― 「1回設定して実測し直し、差分だけ足す」の縮小版。
'   2回で収まらない場合は固定列だけで target を超えている(=内容のほうが
'   広い正当なケース)なので、それ以上は触らない。
Public Sub FitVerify(ByVal ws As Worksheet, ByVal bandAddr As String, _
                     ByVal padColLetter As String, ByVal target As Double)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Dim i As Long
    For i = 1 To 2
        Dim bandW As Double: bandW = ws.Range(bandAddr).Width
        If bandW <= target Then Exit For
        Dim u1 As Double: u1 = ws.Columns(padColLetter).ColumnWidth
        Dim w1 As Double: w1 = ws.Columns(padColLetter).Width
        Dim needPt As Double: needPt = w1 - (bandW - target)
        If needPt < 0.5 Then Exit For      ' 吸収列では吸い切れない(固定列が広い)
        ws.Columns(padColLetter).ColumnWidth = 0.05
        Dim u0 As Double: u0 = ws.Columns(padColLetter).ColumnWidth
        Dim w0 As Double: w0 = ws.Columns(padColLetter).Width
        ws.Columns(padColLetter).ColumnWidth = _
            modViewport.PadUnitsRefine(needPt, u0, w0, u1, w1)
    Next i
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' S1の保険: 描いた後に窓が動いていたら1回だけ組み直す(ワンショット)
' ----------------------------------------------------------------------------
' ViewMoved - 可視サイズが動いたか。純関数(ゴールデン対象)。
Public Function ViewMoved(ByVal w0 As Double, ByVal h0 As Double, _
                          ByVal w1 As Double, ByVal h1 As Double, _
                          ByVal tol As Double) As Boolean
    ViewMoved = (Abs(w1 - w0) > tol) Or (Abs(h1 - h0) > tol)
End Function

' MarkView - 描画の直前に可視サイズを控える(ByRefで2値を受け取る)。
Public Sub MarkView(ByRef w As Double, ByRef h As Double)
    On Error Resume Next
    w = modUIMain.ViewportWidth()
    h = modViewport.ViewportHeight()
    On Error GoTo 0
End Sub

' ReflowIfMoved - 描画の末尾で再測定し、4pt超動いていたら【1回だけ】組み直す。
'   2回目は mReflowGuard が必ず止める(無限ループ禁止)。組み直しの分岐は
'   modViewport.RefitActiveScreen 1本に集約する(2つ持つとズレる)。
Public Sub ReflowIfMoved(ByVal w0 As Double, ByVal h0 As Double)
    If mReflowGuard Then Exit Sub
    On Error Resume Next
    If Not ViewMoved(w0, h0, modUIMain.ViewportWidth(), modViewport.ViewportHeight(), VIEW_TOL) Then
        On Error GoTo 0
        Exit Sub
    End If
    mReflowGuard = True
    modViewport.RefitActiveScreen
    Err.Clear
    mReflowGuard = False
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' S5: 窓高適応圧縮(内容が収まらないときだけ縦を詰める)
' ----------------------------------------------------------------------------
' CompressFactor - 圧縮係数。純関数(ゴールデン対象)。
'   2026-08-07(R21H F3・敵対的レビュー確定): 引数を need 1本から
'   needFixed(SY()を通らない固定要素)/needVariable(SY()で縮む可変要素)の
'   2本へ分けた。旧実装は s=(viewH-8)/need を固定chrome込みの合計needで
'   割っており、SY()は可変要素にしか掛からないのに分母だけ固定要素を含む
'   ため、実際に縮む量が必要量より系統的に不足していた(Hub必要19pt削減に
'   対し実際は12.6ptしか縮まない/gallery は適用後もなお窓を超える)。
'   固定要素は圧縮しても縮まないので、可変要素だけで不足ぶんを吸わせる:
'     fixed + s*variable = avail  ⇔  s = (avail - fixed) / variable
'   need(fixed+variable)が可視高に収まるなら1(何もしない)。
'   ※既定値 0.78 は MIN_SCALE と同値(VBAのOptional既定値は定数式のみ)。
Public Function CompressFactor(ByVal viewH As Double, ByVal needFixed As Double, _
                               ByVal needVariable As Double, _
                               Optional ByVal minScale As Double = 0.78) As Double
    CompressFactor = 1
    If needVariable <= 0 Then Exit Function
    Dim avail As Double: avail = viewH - 8
    If avail >= needFixed + needVariable Then Exit Function
    CompressFactor = modViewport.ClampD((avail - needFixed) / needVariable, minScale, 1)
End Function

' SetScaleY - その画面の圧縮係数を確定させる(描画の一番手前で1回)。
Public Sub SetScaleY(ByVal s As Double)
    mScaleY = modViewport.ClampD(s, MIN_SCALE, 1)
End Sub

' ScaleY - 現在の圧縮係数(未設定は1)。
Public Function ScaleY() As Double
    ScaleY = mScaleY
    If ScaleY <= 0 Then ScaleY = 1
    If ScaleY > 1 Then ScaleY = 1
End Function

' SY - 縦の可変量に係数を掛ける唯一の口。呼び出し側は定数の参照を SY(定数)へ
'   置き換えるだけで済む(掛け算をここ1箇所に閉じる)。
Public Function SY(ByVal v As Double) As Double
    SY = v * ScaleY()
End Function

' HubNeedYFixed / HubNeedYVariable - Hubが s=1 のときに縦へ積む合計(pt)を
'   CompressFactor(F3)の分母に合わせて固定/可変へ分解したもの。純関数(ゴールデン対象)。
'   固定=ヘッダー実高+プロフィール上余白12+行境界の丸め1行15pt(バッジ帯は
'   セル行なのでSY()が掛からない)+バッジ帯(見出し+4行=75pt)
'   +フッター本体16と下余白8。
'   可変=プロフィールカード高cardH(SY(CARD_H))+下余白18(SY(18))
'   +統計タイルtilesH(SY(TilesHeight()))+フッター上余白10(SY(10))。
'   実体を modHub ではなくここに置いたのは容量(WARN帯)と、この式を
'   ゴールデンで固定したいため。呼び出し側は定数を渡すだけ。
Public Function HubNeedYFixed(ByVal hdrH As Double) As Double
    HubNeedYFixed = hdrH + 12 + 15 + 75 + 24
End Function

Public Function HubNeedYVariable(ByVal cardH As Double, ByVal tilesH As Double) As Double
    HubNeedYVariable = cardH + 18 + tilesH + 10
End Function

' HubNeedY - 後方互換の合計(pt)。s=1のときにHubが縦へ積む総量そのもの
'   (「収まるかどうか」の目視確認・旧テストとの対で残す。CompressFactorへは
'   渡さない=固定/可変を分けずに渡すと再びF3のバグに戻るため直接使わないこと)。
Public Function HubNeedY(ByVal hdrH As Double, ByVal cardH As Double, _
                         ByVal tilesH As Double) As Double
    HubNeedY = HubNeedYFixed(hdrH) + HubNeedYVariable(cardH, tilesH)
End Function

' BadgeRowsFor - Hubのバッジ帯に使う行数。純関数(ゴールデン対象)。
'   バッジ帯は【セルの行】(15pt固定)なので SY() では縮まない。圧縮が
'   かかっている画面では段数そのものを1つ落とす。
'   2026-08-07(R21H F3再校正): 閾値0.95はCompressFactorの旧式(固定chrome込み
'   need で割る)が返す値に合わせて選んでいた。新式は同じ実窓高でも系統的に
'   小さい係数を返す(分母を可変要素だけに絞ったため)ので、旧式のまま流用すると
'   実機の窓499pt(旧式0.963→新式0.943)でも行が落ちてしまい、意図せず表示が
'   後退する。実機499ptで従来どおり4行を保つ境界として0.92へ再校正した
'   (Hub 1段ヘッダーの新式s: 499pt→0.9435・450pt→0.7976。499ptは4行のまま・
'   450ptは3行への実際の落差と整合する)。
'   2026-08-10(R23c F4): 引数名を scale から sc へ改名。Scale はMS-VBAL仕様の
'   reserved-name かつ special-form(VB伝統のグラフィック命令)であり、本物の
'   Excel VBAパーサは識別子として拒否する。LibreOffice Basicは通してしまうため
'   LO検査では検出できず、実機でのみ「BadgeRowsFor が見つかりません/構文エラー」
'   として表面化していた(再発防止は tools/vba_lint.py の予約名検査)。
Public Function BadgeRowsFor(ByVal sc As Double) As Long
    BadgeRowsFor = 4
    If sc < 0.92 Then BadgeRowsFor = 3
End Function

' ----------------------------------------------------------------------------
' S3: 弾性グリッド(カードが右端まで張る列数とカード幅)
' ----------------------------------------------------------------------------
' GridColsFor - 幅 availW・左余白 leftX の領域に、最小幅 cardMinW のカードが
'   何列入るか。純関数(ゴールデン対象)。最後のカードの右には隙間を残さない
'   前提なので、隙間の数は (cols-1) ―― 数えるときは gap を1つ足して割る。
'   modViewport.GalleryColsFor(固定カード幅・末尾に隙間が残る旧式)の後継。
Public Function GridColsFor(ByVal availW As Double, ByVal leftX As Double, _
                            ByVal cardMinW As Double, ByVal gap As Double, _
                            ByVal minCols As Long, ByVal maxCols As Long) As Long
    GridColsFor = minCols
    If cardMinW + gap <= 0 Then Exit Function
    Dim n As Long
    n = Int((availW - leftX + gap) / (cardMinW + gap))
    If n < minCols Then n = minCols
    If n > maxCols Then n = maxCols
    GridColsFor = n
End Function

' GridCardW - その列数で右端まで張るカード幅。純関数(ゴールデン対象)。
'   上限 maxW を超える窓では上限で頭打ちにする(1枚が広がりすぎると
'   一覧としての情報密度が落ちる)。
Public Function GridCardW(ByVal availW As Double, ByVal leftX As Double, _
                          ByVal gap As Double, ByVal cols As Long, _
                          ByVal maxW As Double, ByVal minW As Double) As Double
    GridCardW = minW
    If cols < 1 Then Exit Function
    GridCardW = modViewport.ClampD((availW - leftX - gap * (cols - 1)) / cols, minW, maxW)
End Function

' GridGapFor - カードが上限で頭打ちになってなお余る幅を隙間へ配分する(pt)。
'   純関数(ゴールデン対象)。modDashStat.CardGapFor と同じ考え方で、これが
'   無いと「カード幅の上限に達した広い窓」でだけ右端が帯に届かない
'   (1600pt窓のギャラリーで54pt残っていた)。余りが無ければ既定 gap のまま。
Public Function GridGapFor(ByVal availW As Double, ByVal leftX As Double, _
                           ByVal cardW As Double, ByVal cols As Long, _
                           ByVal gap As Double) As Double
    GridGapFor = gap
    If cols < 2 Then Exit Function
    Dim g As Double: g = (availW - leftX - cols * cardW) / (cols - 1)
    If g > gap Then GridGapFor = g
End Function

' RightGapExceeds - 本文の実右端が ContentRight から離れすぎていないか。
'   純関数(ゴールデン対象)。S3の不変条件「本文要素の右端 == ContentRight」を
'   机上で検算するための検出器で、12pt を超える差は設計の破れとみなす。
Public Function RightGapExceeds(ByVal contentMaxRight As Double, _
                                ByVal contentRight As Double, _
                                ByVal tol As Double) As Boolean
    RightGapExceeds = (Abs(contentRight - contentMaxRight) > tol)
End Function

' ----------------------------------------------------------------------------
' S6: 再フィット経路の穴埋め
' ----------------------------------------------------------------------------
' ShelfPadCol - 「マイ本棚」シートで余りを吸う列(R21-S3)。純関数。
'   一覧表(table)は【本文の最終列】メモJが吸う=本文の右端そのものが窓幅へ
'   追随する。gallery(本文がpt座標のShapeで別に伸びる)は帯の最終列
'   (defaultCol)が吸う。従来は一覧表だけ「EnsureLayoutでJ→DrawChromeでN」の
'   2段階Fitで、2回目の丸めで帯が可視幅を数pt超える経路が残っていた。
'   吸収列をモードごとに1つ選び、Fitは【1回だけ】通す。
'   2026-08-07(R21H F7): sharedはgalleryと同じ「カード系」に分類していたが
'   実際はtableと同じセル値主体の一覧(質問文はD:J列そのもの)で、Shapeで
'   本文が別に伸びる画面ではない。modShared.Showがtableと同型の外部Fit(J)を
'   打った上でDrawChromeが再びN側でFitし直す二段Fitになっていたので、
'   tableと同じJへ倒し、Fitは1回だけにする。
'   defaultCol を引数で受けるのは、帯の最終列の単一情報源
'   (modKnowledge.SHELF_PAD_COL)を2箇所に持たないため ―― 同時に、この関数が
'   他モジュールを呼ばない純関数になり LibreOffice のゴールデンで固定できる。
'   実体をここに置いたのは modKnowledge の容量(WARN帯)を守るため。
Public Function ShelfPadCol(ByVal mode As String, ByVal defaultCol As String) As String
    Dim m As String: m = LCase$(Trim$(mode))
    If m = "table" Or m = "shared" Then
        ShelfPadCol = "J"
    Else
        ShelfPadCol = defaultCol
    End If
End Function

' RefitShelfTable - 一覧表モードの再フィット(R21-S6)。
'   窓リサイズの再フィットは modKnowledge.RefreshCurrent 経由で
'   modUIShelf.RenderShelf に落ちるが、RenderShelf は【帯(列幅)を一切
'   触らない】ため、一覧の右端が前の窓幅のまま取り残されていた。
'   シートの作り直し(EnsureLayout 全体)まで戻さず、「表示状態→列幅→
'   共通クロム(この中で Fit)→本文」だけをやり直す軽量経路。
Public Sub RefitShelfTable(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    EnsureViewState ws
    modUIShelf.RefitColumns ws
    modKnowledge.DrawChrome ws, "table"
    modUIShelf.RenderShelf
    On Error GoTo 0
End Sub

' RefitChatBand - チャット画面の帯・ヘッダー・入力欄を今の窓幅へ合わせ直す。
'   modUI.Repaint は再彩色しかしておらず、窓を広げても列幅(帯)と入力欄の
'   幅が前の窓のままだった(R21-S6の穴その2)。modUI に容量が無いため実体を
'   ここに置き、Repaint からは1行呼ぶ。
'   nx_top_add / nx_top_send は modUINexusDraw の掃除対象から外れている
'   (ヘッダー再描画で消さない約束)ので、二重生成しないよう先に落とす。
'   会話バブルそのものの再フロー(過去バブルの幅)は対象外(R21で次期送り)。
Public Sub RefitChatBand(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    modViewport.FitBandToViewport ws, modUINexusDraw.NEXUS_BAND, modUI.NEXUS_INPUT_PAD_COL
    DropShape ws, "nx_top_add"
    DropShape ws, "nx_top_send"
    ' R21H F1: DrawChatHeaderを直呼びすると同名Shape(nx_top_bg等)がClearを
    ' 経ずに積み上がる(ExcelはShapeの同名重複を許すため、Repaint/テーマ切替/
    ' リサイズのたびに1個ずつ増える)。RedrawChatHeaderはClearChatHeader→
    ' DrawChatHeader→行高→前面化を1本にまとめた契約なので、必ずこちらを通す。
    modUINexusDraw.RedrawChatHeader
    modUINexusDraw.DrawInputArea ws
    ' R30 F2: ApplyThemeはmChatBandRowだけ先へ進め、新規bound内行のRowHeight=18を
    ' 誰も設定しないままにする(ScrollToBottomの18pt換算がずれる)。ここで回復する。
    FitChatRows ws
    Dim bnd As String: bnd = modUINexusDraw.NexusBound(ws)
    modViewport.ApplyScrollBound ws, bnd
    LogFit ws, "chat", modUINexusDraw.NEXUS_BAND, bnd
    On Error GoTo 0
End Sub

' DropShape - 名前が一致するShapeを1つ落とす(無ければ何もしない)。
Private Sub DropShape(ByVal ws As Worksheet, ByVal shapeName As String)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    Err.Clear
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' S7: 可観測性(フィット直後の5値ログ)
' ----------------------------------------------------------------------------
' LogFit - 「窓幅/可視セル幅/帯実幅/中身右端/境界下端」を1行で残す。
'   次の実機報告で、余白がどの層(帯・中身・境界)の破れかが一意に決まる。
'   実体(重複抑止つき)は modViewport.LogViewport。
Public Sub LogFit(ByVal ws As Worksheet, ByVal screenName As String, _
                  ByVal bandAddr As String, ByVal boundAddr As String)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Dim bandW As Double, rightX As Double, botY As Double
    bandW = ws.Range(bandAddr).Width
    rightX = modViewport.ContentRight(ws, bandAddr, 8)
    If LenB(boundAddr) > 0 Then botY = ws.Range(boundAddr).Height
    modViewport.LogViewport screenName, bandW, rightX, botY
    On Error GoTo 0
End Sub

' LogChat - チャットの初期描画(modUI.EnsureLayout)からの1行呼び出し。
Public Sub LogChat(ByVal ws As Worksheet)
    LogFit ws, "chat", modUINexusDraw.NEXUS_BAND, modUINexusDraw.NexusBound(ws)
End Sub

' ----------------------------------------------------------------------------
' F2-3(R27・実機第12報①): 埋め草 ―― 境界の最終行を窓下端まで届かせる
' ----------------------------------------------------------------------------
' 内容が窓に収まる画面では、境界の最終行は modViewport.RowAtFloor(切り下げ)で
' 決まる。切り下げなので【最終行の下端は必ず窓高より上】で、その差(最大で
' 1行ぶん)は塗りもScrollAreaも届かない帯になる ―― ホイールで転がったときに
' bg色ではなく素の白が見える下端の余白の一部はこれ。切り上げにすると必ず窓を
' 1行ぶん超え、その1行のぶんだけ縦スクロールが生き残る(R21-1の裁定)ので、
' 行数は増やさず【最終行そのものを差分ぶん高くする】。
' 上限クランプの意図: 差が1行ぶん(既定18pt前後)を大きく超えるのは、境界が
' 実下端とずれている等の構造問題であって埋め草の出番ではない。PAD_ROW_MAX を
' 超える差は何もせずに残す ―― 埋めてしまうと症状だけが消えて原因の観測が
' できなくなる(定数はモジュール先頭の宣言部)。
'
' PadRowDelta - 最終行に足すべき高さ(pt)。純関数(modTestsPure24が固定)。
'   boundBottomY : 境界最終行の下端Y(pt)
'   viewportH    : 窓の可視高(pt)
'   maxPad       : 埋めてよい差の上限(pt)。これを超えたら 0(=何もしない)
Public Function PadRowDelta(ByVal boundBottomY As Double, ByVal viewportH As Double, _
                            ByVal maxPad As Double) As Double
    Dim d As Double: d = viewportH - boundBottomY
    If d < PAD_ROW_MIN Then Exit Function   ' 既に届いている/丸め未満の差は触らない
    If d > maxPad Then Exit Function        ' 構造問題は隠さない
    PadRowDelta = d
End Function

' PadRowToWindow - 境界最終行の高さを足して、塗りの下端を窓下端へ届かせる。
'   boundLastRow: 境界(A1起点)の最終行番号。呼び出し側は Range(bnd).Rows.Count。
'   冪等: 一度届かせると次回の差は0になり、二度と足さない。
Public Sub PadRowToWindow(ByVal ws As Worksheet, ByVal boundLastRow As Long)
    If ws Is Nothing Then Exit Sub
    If boundLastRow < 1 Then Exit Sub
    On Error Resume Next
    Dim botY As Double
    botY = ws.Cells(boundLastRow, 1).Top + ws.Cells(boundLastRow, 1).Height
    Dim d As Double
    d = PadRowDelta(botY, modViewport.ViewportHeight(), PAD_ROW_MAX)
    If d >= PAD_ROW_MIN Then
        ws.Rows(boundLastRow).RowHeight = ws.Rows(boundLastRow).RowHeight + d
    End If
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' W1(R30・実機第15報): チャットの行高と「使用済み行」の解放
' ----------------------------------------------------------------------------
' 行高の明示設定(RowHeight=18)は、その行をExcelの内部使用範囲へ焼き付ける。
' 焼き付いた行は ClearFormats でも保存でも消えず、Rows.Delete だけが即時に
' 解放できる(R30実機実証)。Nexus は FreezePanes 併用でホイールが ScrollArea
' を素通りするため、焼き付いた行の末尾までいくらでも転がれる ―― 旧
' modUI.InitUI の Rows("1:400").RowHeight=18 が「約10画面ぶんの下余白」の正体。
'
' 以降のW1実装が共有する不変式:
'   行1〜4  : 固定領域(ヘッダー/入力欄/ヒント)。高さは modUI.InitUI が個別に持つ。
'   行5〜B  : バンド。全行18pt(modUI.ScrollToBottom の18pt換算はこれに依存)。
'   行B+1〜 : 存在しない(解放済み)。地は modChrome.ApplyNormalStyleBg の
'             Normalスタイル地色で見えるので、塗る必要は無い。
'   B = バンド下端 + CHAT_ROW_SLACK。
' 実体をここへ置く理由: modUI 残716字 / modSkin 残163字では入らない(憲章§4-6)。

' FitChatRows - チャットの18pt行を「必要な範囲だけ」に確定させる(冪等)。
'   二段構え: (1) 窓を必ず覆う暫定範囲を18ptにしてから、(2) 実測で確定した
'   バンド下端+CHAT_ROW_SLACK までを残して下を解放する。順序が逆だと、
'   既定行高(端末依存で15pt〜18.75pt)のまま測ったバンドへ後から18ptを当てる
'   ことになり、バンドの実下端が窓とずれる。
'   暫定行数 窓高/18 は「18pt行だけで窓を覆うのに要る行数」の上界
'   (行1〜4が108pt前後を占めるぶん必ず余る)。
'   2026-08-13(R30 F2-1・敵対的レビュー2周目MAJOR裁定): 暫定行数(窓高/18)は
'   バンド下端+CHAT_ROW_SLACKより構造的に大きい(窓を覆うための上界であり、
'   実内容の下端とは無関係)。旧実装はこれをそのまま焼いてから解放していたため、
'   「焼いた分をすぐ削除する」が【毎回】起き、Rows.Delete+usage_log書き込みが
'   冪等にならなかった。ChatSeedRowで焼き範囲の下端をバンド下端+SLACKへ
'   クランプし、通常経路では「焼く=残す範囲」を一致させて削除を不発にする
'   (旧ブックの過去の焼き付け掃除は ReleaseRowsBelow が ws.UsedRange の実測値
'   から独立に判定するため、クランプの影響を受けず引き続き機能する)。
Public Sub FitChatRows(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Dim firstRow As Long: firstRow = modUINexusDraw.INPUT_ROW + 2
    Dim boundRow As Long: boundRow = ws.Range(modUINexusDraw.NexusBound(ws)).Rows.Count + CHAT_ROW_SLACK
    Dim seedRow As Long
    seedRow = ChatSeedRow(firstRow, modViewport.ViewportHeight(), _
                          modUINexusDraw.NEXUS_MAX_ROW, boundRow)
    ws.Rows(firstRow & ":" & seedRow).RowHeight = 18
    ReleaseRowsBelow ws, boundRow
    ' R30 F8: ClearChat経路はScrollArea設定後に行削除が走るため、ここで掛け直す。
    ' InitUI経路では後続のDrawInputAreaが再設定するので二重適用でも冪等。
    modViewport.ApplyScrollBound ws, modUINexusDraw.NexusBound(ws)
    On Error GoTo 0
End Sub

' ChatSeedRow - FitChatRowsの暫定焼き範囲の下端行。純関数(modTestsPure29が固定)。
'   firstRow  : 焼き範囲の上端(固定領域の直下)。
'   viewportH : 窓の可視高(pt)。/18 が「18pt行だけで窓を覆う行数」の上界。
'   maxRowCap : NEXUS_MAX_ROW。旧来からの絶対上限(取り残し防止)。
'   boundRow  : バンド下端+CHAT_ROW_SLACK。R30 F2-1の頭打ち先 ―― これを
'               超えて焼くと、その分がそのまま ReleaseRowsBelow で削除される
'               (焼く=残す範囲を一致させて冪等にする)。
'   下限もfirstRowでクランプする: boundRowがfirstRowを下回る異常値でも
'   Rows("5:3")のような反転範囲文字列(行3〜4=固定領域まで巻き込む)を
'   作らないため。
Public Function ChatSeedRow(ByVal firstRow As Long, ByVal viewportH As Double, _
                            ByVal maxRowCap As Long, ByVal boundRow As Long) As Long
    ChatSeedRow = firstRow + CLng(viewportH / 18)
    If ChatSeedRow > maxRowCap Then ChatSeedRow = maxRowCap
    If ChatSeedRow > boundRow Then ChatSeedRow = boundRow
    If ChatSeedRow < firstRow Then ChatSeedRow = firstRow
End Function

' ReleaseRowsBelow - チャット(Nexus)の使用済み行を解放する(冪等)。
'   R31 W2-1: 中身は汎用の ReleaseSheetRowsBelow へ移した(Nexus定数の直書きを
'   引数へ出しただけで挙動は同一。ログのラベルも "chat" のまま)。
'   ScrollArea は常に boundRow 以内なので掛け直しは不要(boundAddr は空)。
Public Sub ReleaseRowsBelow(ByVal ws As Worksheet, ByVal boundRow As Long)
    modViewport.ReleaseSheetRowsBelow ws, boundRow, modUINexusDraw.NEXUS_MAX_ROW, "", _
                          "chat", modUINexusDraw.INPUT_ROW + 2
End Sub

' SeedRowCap - 「セッション初回の高水位フォールバック」の行。純関数
'   (modTestsPure29 が固定)。R31 W2-4: 本棚は mShelfRowHigh がセッション毎に
'   0へ戻るため、初回だけ SHELF_MAX_ROW=412(約8画面)を均し・クリアの範囲に
'   していた。焼いてすぐ削除する非冪等な往復(R30 F2-1と同型)になるので、
'   「窓を覆う行数」と「実際の使用済み下端」の大きい方へ頭打ちする。
'   firstRow  : 範囲の上端行。戻り値はここを下回らない。
'   viewportH : 窓の可視高(pt)。/rowH が「窓を覆う行数」の上界。
'   rowH      : 均した後の行高(pt)。0以下なら15ptとみなす。
'   lastUsed  : 現在の使用済み最終行。旧ブックの焼き付け救済のため、これが
'               窓ぶんより大きければそちらを採る。
'   maxRowCap : 絶対上限(SHELF_MAX_ROW)。
Public Function SeedRowCap(ByVal firstRow As Long, ByVal viewportH As Double, _
                           ByVal rowH As Double, ByVal lastUsed As Long, _
                           ByVal maxRowCap As Long) As Long
    Dim h As Double: h = rowH
    If h < 1 Then h = 15
    SeedRowCap = firstRow + CLng(viewportH / h)
    If lastUsed > SeedRowCap Then SeedRowCap = lastUsed
    If SeedRowCap > maxRowCap Then SeedRowCap = maxRowCap
    If SeedRowCap < firstRow Then SeedRowCap = firstRow
End Function

' ShelfSeedRow - 本棚の高水位フォールバック(セッション初回)。SeedRowCap に
'   窓高と ws.UsedRange の実測を与えるだけの薄い口(modKnowledge 残160字対策)。
Public Function ShelfSeedRow(ByVal ws As Worksheet, ByVal maxRowCap As Long) As Long
    Dim lastUsed As Long
    On Error Resume Next
    If Not ws Is Nothing Then lastUsed = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    On Error GoTo 0
    ShelfSeedRow = SeedRowCap(1, modViewport.ViewportHeight(), 15, lastUsed, maxRowCap)
End Function

' ReleaseRange - 解放する行範囲を決める。純関数(modTestsPure29が固定)。
'   boundRow : 残す下端行(バンド下端+余裕)
'   lastUsed : 現在の使用済み最終行(ws.UsedRange の下端)
'   minRow   : 固定領域の直下。boundRow がこれを下回っても必ずここまでは残す
'              ―― 行1〜4を消すと FreezePanes ごと画面が壊れる
'   maxRow   : 旧版が焼き得た上限(NEXUS_MAX_ROW)。使用済みがそこまで届いて
'              いなくても、この行までは消す(旧版の焼き付けの取りこぼし防止)
'   戻り値   : True=削除が要る(fromRow/toRow に範囲を返す)。
'              False のとき fromRow/toRow の値は意味を持たない。
Public Function ReleaseRange(ByVal boundRow As Long, ByVal lastUsed As Long, _
                             ByVal minRow As Long, ByVal maxRow As Long, _
                             ByRef fromRow As Long, ByRef toRow As Long) As Boolean
    Dim keepRow As Long: keepRow = boundRow
    If keepRow < minRow Then keepRow = minRow
    fromRow = keepRow + 1
    toRow = lastUsed
    If toRow < maxRow Then toRow = maxRow
    ' 旧世代ブックのUsedRangeが病的に大きい場合(例: 破損・外部貼付けの取りこぼし)、
    ' Rows.Delete が数十万行に及び32bitで凍結し得る(R30 F1)。maxRow(NEXUS_MAX_ROW)の
    ' 4倍を上限に頭打ちする ―― 通常経路では届かない余裕を持たせつつ、病的値だけ抑える。
    If toRow > maxRow * 4 Then toRow = maxRow * 4
    ' R30 F2-2(敵対的レビュー2周目MINOR裁定): F1のmaxRow*4クランプ導入で、
    ' boundRow(ひいてはkeepRow)がmaxRow*4を上回る異常値のとき、頭打ち後の
    ' toRowがfromRow(keepRow+1)より小さい逆転範囲が理論上作れる状態になった
    ' (現状の呼び出し経路では到達不能だが、不変式として保護する)。
    If toRow < fromRow Then
        ReleaseRange = False
        Exit Function
    End If
    ReleaseRange = (lastUsed > keepRow)
End Function
