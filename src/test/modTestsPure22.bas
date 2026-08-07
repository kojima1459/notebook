Attribute VB_Name = "modTestsPure22"
Option Explicit

' ============================================================================
' modTestsPure22 - R21-1(実機第8報⑦・余白の構造完治)の純ロジックゴールデン。
'   入口は modTestsPure21.RunAll21 の末尾から呼ばれる RunAll22 の1本。
'   modTestsPure21(6,692字)はR20Hの3点固定が主題で、余白の物理とは別の話
'   なので分割先を新設した(憲章§4-6)。
' ----------------------------------------------------------------------------
' ここで固定するもの(S1〜S5の受け入れ基準を机上で検算できる形にする):
'   S1 modViewport2.HScrollNeeded : 水平スクロールバーは狭窓の安全弁だけ
'      modViewport2.ViewMoved     : 描画前後で窓が動いたかの4pt判定(ワンショット)
'   S2 modViewport2.SbWidthFrom   : 縦スクロールバー幅の校正(決め打ち12ptの撤廃)
'   S3 modViewport2.GridColsFor / GridCardW / GridGapFor / RightGapExceeds
'      modDashStat.CardGapFor     : 余りを隙間へ配分して右端まで張る
'   S4 modViewport.FitsInView     : 内容が窓に収まるか(境界を切り下げてよいか)
'   S5 modViewport2.CompressFactor / HubNeedY / BadgeRowsFor
'   S3 modViewport2.ShelfPadCol   : モードごとの吸収列(2段階Fitの廃止)
'
' 実測の前提(調査A班の数値トレース。実機の窓は 1024×499pt):
'   可視セル幅 = 1024 − 縦スクロールバー12 = 1012
'   帯の目標   = 1012 − 安全余裕2 = 1010
'   帯の右端(操作系) = 1010 − 8 = 1002
' 旧実装ではダッシュの中身の右端が961pt・ギャラリーが916ptで止まっており、
' その差(41pt/86pt)が写真に写っていた右の白帯だった。下のゴールデンは
' 「新式なら 1002 に一致し、旧式なら一致しない」を両方固定する。
'
' Worksheet が要る modViewport.RowAtFloor(実セル幾何を読む)は純ロジックでは
' 回せないため、ここでは RowAtFloor に渡す判断(FitsInView)だけを固定する。
' ============================================================================

' ----------------------------------------------------------------------------
' S1: 水平スクロールバーは狭窓(<625pt)の安全弁だけ(modViewport2.HScrollNeeded)
' ----------------------------------------------------------------------------
' 出しっぱなしにすると可視高が約24pt減り、境界計算が必ず窓を超える
' (実機第8報⑦の独立欠陥その2)。広い窓では「帯≤可視幅」が成立するので不要。
Private Sub TestHScrollNeeded()
    modTestRunner.Check "水平バー_窓624ptは出す(境界の内側)", _
        (modViewport2.HScrollNeeded(624) = True)
    modTestRunner.Check "水平バー_窓625ptは出さない(境界)", _
        (modViewport2.HScrollNeeded(625) = False)
    modTestRunner.Check "水平バー_実機の窓1024ptは出さない", _
        (modViewport2.HScrollNeeded(1024) = False)
    modTestRunner.Check "水平バー_極端に狭い窓320ptは出す", _
        (modViewport2.HScrollNeeded(320) = True)
End Sub

' ----------------------------------------------------------------------------
' S1: 描画の前後で窓が動いたか(modViewport2.ViewMoved・許容4pt)
' ----------------------------------------------------------------------------
Private Sub TestViewMoved()
    modTestRunner.Check "窓移動_同じなら動いていない", _
        (modViewport2.ViewMoved(1024, 499, 1024, 499, 4) = False)
    modTestRunner.Check "窓移動_高さ+4ptは許容内(境界)", _
        (modViewport2.ViewMoved(1024, 499, 1024, 503, 4) = False)
    modTestRunner.Check "窓移動_高さ+5ptは組み直す(境界の外)", _
        (modViewport2.ViewMoved(1024, 499, 1024, 504, 4) = True)
    modTestRunner.Check "窓移動_幅-4ptは許容内(境界)", _
        (modViewport2.ViewMoved(1024, 499, 1020, 499, 4) = False)
    modTestRunner.Check "窓移動_幅-5ptは組み直す", _
        (modViewport2.ViewMoved(1024, 499, 1019, 499, 4) = True)
    ' 水平スクロールバーが消えたぶん(約24pt)は必ず検知できること。
    modTestRunner.Check "窓移動_水平バー消滅の24ptは必ず検知", _
        (modViewport2.ViewMoved(1024, 475, 1024, 499, 4) = True)
End Sub

' ----------------------------------------------------------------------------
' S2: 縦スクロールバー幅の校正(modViewport2.SbWidthFrom)
' ----------------------------------------------------------------------------
' UsableWidth と VisibleRange.Width の差がバー幅。ただし VisibleRange は列単位で
' 切り上がるので、負や大きすぎる差は「バー幅」ではない=既定値へ倒す。
Private Sub TestSbWidthFrom()
    modTestRunner.Check "バー幅_1024と1012の差12ptを採る", _
        (modViewport2.SbWidthFrom(1024, 1012, 12) = 12)
    modTestRunner.Check "バー幅_差18ptも実測どおり採る(端末差)", _
        (modViewport2.SbWidthFrom(1024, 1006, 12) = 18)
    modTestRunner.Check "バー幅_差0(バー無し)も採る", _
        (modViewport2.SbWidthFrom(1024, 1024, 12) = 0)
    modTestRunner.Check "バー幅_差30ptは採る(上限の境界)", _
        (modViewport2.SbWidthFrom(1024, 994, 12) = 30)
    modTestRunner.Check "バー幅_差31ptは信じず既定12へ(境界の外)", _
        (modViewport2.SbWidthFrom(1024, 993, 12) = 12)
    modTestRunner.Check "バー幅_負の差(切り上がり)は既定12へ", _
        (modViewport2.SbWidthFrom(1024, 1030, 12) = 12)
    modTestRunner.Check "バー幅_測れない(0)ときは既定12へ", _
        (modViewport2.SbWidthFrom(0, 1012, 12) = 12 And _
         modViewport2.SbWidthFrom(1024, 0, 12) = 12)
    ' 実機1024pt窓の帯の目標: 可視セル幅1012 − 安全余裕2 = 1010 ≤ 1012。
    Dim visW As Double: visW = 1024 - modViewport2.SbWidthFrom(1024, 1012, 12)
    modTestRunner.Check "帯目標_1024pt窓では1010ptで可視セル幅1012を超えない", _
        (visW = 1012 And (visW - 2) = 1010 And (visW - 2) <= visW)
End Sub

' ----------------------------------------------------------------------------
' S3: ギャラリーの弾性グリッド(modViewport2.GridColsFor / GridCardW)
' ----------------------------------------------------------------------------
' 実機1024pt窓: 帯の右端1002pt・カード左端14pt・GAP14pt・最小幅170pt。
'   cols  = int((1002 - 14 + 14) / (170 + 14)) = int(5.44) = 5
'   cardW = (1002 - 14 - 14×4) / 5 = 932 / 5 = 186.4
'   右端  = 14 + 5×186.4 + 4×14 = 1002 = 帯の右端(不変条件)
Private Sub TestGalleryGrid()
    modTestRunner.Check "列数_1024pt窓(avail1002)は5列", _
        (modViewport2.GridColsFor(1002, 14, 170, 14, 3, 6) = 5), _
        "実際=" & modViewport2.GridColsFor(1002, 14, 170, 14, 3, 6)
    modTestRunner.Check "列数_avail736ptでちょうど4列(境界)", _
        (modViewport2.GridColsFor(736, 14, 170, 14, 3, 6) = 4)
    modTestRunner.Check "列数_avail735ptは3列(境界の外)", _
        (modViewport2.GridColsFor(735, 14, 170, 14, 3, 6) = 3)
    modTestRunner.Check "列数_狭い窓でも下限3列(画面が空にならない)", _
        (modViewport2.GridColsFor(400, 14, 170, 14, 3, 6) = 3)
    modTestRunner.Check "列数_広い窓でも上限6列(一覧性を保つ)", _
        (modViewport2.GridColsFor(2000, 14, 170, 14, 3, 6) = 6)
    modTestRunner.Check "列数_退化入力(カード幅+GAPが0)は下限3列", _
        (modViewport2.GridColsFor(1000, 0, 0, 0, 3, 6) = 3)

    Dim cw As Double: cw = modViewport2.GridCardW(1002, 14, 14, 5, 240, 170)
    modTestRunner.Check "カード幅_1024pt窓の5列は186.4pt", _
        (Abs(cw - 186.4) < 0.001), "実際=" & cw
    modTestRunner.Check "カード幅_下限170を割らない(狭い窓)", _
        (modViewport2.GridCardW(400, 14, 14, 3, 240, 170) = 170)
    modTestRunner.Check "カード幅_上限240で頭打ち(広い窓)", _
        (modViewport2.GridCardW(2000, 14, 14, 6, 240, 170) = 240)
    modTestRunner.Check "カード幅_列数0の退化入力は下限170", _
        (modViewport2.GridCardW(1000, 14, 14, 0, 240, 170) = 170)

    ' 不変条件: 本文(最後のカード)の右端 == 帯の右端(ContentRight)。
    Dim gp As Double: gp = modViewport2.GridGapFor(1002, 14, cw, 5, 14)
    modTestRunner.Check "隙間_余りが無いので既定の14ptのまま", (gp = 14), "実際=" & gp
    Dim rightX As Double: rightX = 14 + 5 * cw + 4 * gp
    modTestRunner.Check "S3不変条件_ギャラリーの右端は帯の右端1002ptに一致", _
        (modViewport2.RightGapExceeds(rightX, 1002, 12) = False), _
        "実際の右端=" & rightX

    ' 1600pt窓(avail1578): 6列でカード幅が上限240に張り付き、余り124ptは
    ' 隙間5つへ配分されて右端が帯へ届く(配分が無いと54pt足りなかった)。
    Dim wCols As Long: wCols = modViewport2.GridColsFor(1578, 14, 170, 14, 3, 6)
    Dim wCw As Double: wCw = modViewport2.GridCardW(1578, 14, 14, wCols, 240, 170)
    Dim wGp As Double: wGp = modViewport2.GridGapFor(1578, 14, wCw, wCols, 14)
    modTestRunner.Check "広い窓_6列・カード幅240で頭打ち", (wCols = 6 And wCw = 240)
    modTestRunner.Check "広い窓_余り124ptを隙間5つへ配分して24.8pt", _
        (Abs(wGp - 24.8) < 0.001), "実際=" & wGp
    modTestRunner.Check "S3不変条件_1600pt窓でも右端は帯の右端1578ptに一致", _
        (Abs(14 + wCols * wCw + (wCols - 1) * wGp - 1578) < 0.001)
    modTestRunner.Check "隙間_1列以下の退化入力は既定のまま", _
        (modViewport2.GridGapFor(1578, 14, 240, 1, 14) = 14)
    ' 旧実装(カード幅215pt固定・列数だけ弾性)は4列で右端916pt=帯より86pt手前。
    ' これが実機写真の右の白帯で、退行したら必ずここで落ちる。
    Dim oldCols As Long
    oldCols = modViewport.GalleryColsFor(1002, 14, 215, 14, 3, 6)
    Dim oldRight As Double: oldRight = 14 + oldCols * 215 + (oldCols - 1) * 14
    modTestRunner.Check "S3退行検知_旧式(215pt固定)の右端916ptは不合格", _
        (oldCols = 4 And oldRight = 916 And _
         modViewport2.RightGapExceeds(oldRight, 1002, 12) = True), _
        "旧式の列数=" & oldCols & " 右端=" & oldRight
End Sub

' ----------------------------------------------------------------------------
' S3: ダッシュの余り配分(modDashStat.CardGapFor)
' ----------------------------------------------------------------------------
' 実機1024pt窓: 帯 A:K = 1010pt。
'   cardW = clamp((1010 - 40 - 30)/4, 130, 260) = 235
'   gap   = (1010 - 40 - 4×235)/3 = 30/3 = 10
'   右端  = 20 + 4×235 + 3×10 = 990 = 1010 − 右余白20(不変条件)
' 旧実装は上限220 + センタリングで右端961pt(=帯より49pt手前)だった。
Private Sub TestDashGap()
    Dim cw As Double: cw = modDashStat.CardWidthFor(1010)
    Dim gp As Double: gp = modDashStat.CardGapFor(1010, cw)
    modTestRunner.Check "ダッシュ_1024pt窓のカード幅は235pt", _
        (cw = 235), "実際=" & cw
    modTestRunner.Check "ダッシュ_1024pt窓の隙間は既定の10pt(余りゼロ)", _
        (gp = 10), "実際=" & gp
    Dim rightX As Double: rightX = 20 + 4 * cw + 3 * gp
    modTestRunner.Check "S3不変条件_ダッシュの右端は帯の内側右余白20ptに一致", _
        (rightX = 1010 - 20), "実際の右端=" & rightX

    ' カードが上限260で頭打ちになる広い帯では、余りが隙間へ回って右端まで張る。
    Dim cw2 As Double: cw2 = modDashStat.CardWidthFor(1600)
    Dim gp2 As Double: gp2 = modDashStat.CardGapFor(1600, cw2)
    modTestRunner.Check "ダッシュ_帯1600ptはカード260で頭打ち", (cw2 = 260)
    modTestRunner.Check "ダッシュ_帯1600ptの余りは隙間へ(右端が帯へ届く)", _
        (Abs(20 + 4 * cw2 + 3 * gp2 - (1600 - 20)) < 0.001), _
        "実際の右端=" & (20 + 4 * cw2 + 3 * gp2)
    ' 狭い帯では隙間が既定を下回らない(カードが重ならない)。
    modTestRunner.Check "ダッシュ_最小版面590ptでも隙間は10ptを割らない", _
        (modDashStat.CardGapFor(590, modDashStat.CardWidthFor(590)) = 10)
    modTestRunner.Check "ダッシュ_退化入力(帯0pt)でも隙間は10pt", _
        (modDashStat.CardGapFor(0, 130) = 10)
End Sub

' ----------------------------------------------------------------------------
' S4: 内容が窓に収まるか(modViewport.FitsInView)
' ----------------------------------------------------------------------------
' True のときだけ境界を RowAtFloor(窓高) へ【切り下げ】てよい。従来は
' 「+24pt → 行の切り上げ → 最低1画面」の三重切り上げで、収まっている画面でも
' 境界が必ず窓の外に出ていた(縦スクロールが常に生き残る)。
Private Sub TestFitsInView()
    modTestRunner.Check "収まる_Hub実測478pt/窓499ptは収まる", _
        (modViewport.FitsInView(478, 499, 8) = True)
    modTestRunner.Check "収まる_491pt/窓499ptは境界ちょうどで収まる", _
        (modViewport.FitsInView(491, 499, 8) = True)
    modTestRunner.Check "収まる_492pt/窓499ptは収まらない(境界の外)", _
        (modViewport.FitsInView(492, 499, 8) = False)
    modTestRunner.Check "収まる_管理者一覧が伸びた800ptは収まらない", _
        (modViewport.FitsInView(800, 499, 8) = False)
    ' 収まらない側は従来どおり BoundBottomY(内容+8pt)を使う。
    modTestRunner.Check "収まらない側_下端は内容+8pt(最低1画面の底上げは効かない)", _
        (modViewport.BoundBottomY(800, 499, 8) = 808)
End Sub

' ----------------------------------------------------------------------------
' S5: 窓高適応圧縮(modViewport2.CompressFactor / HubNeedY / BadgeRowsFor)
' ----------------------------------------------------------------------------
' Hubが s=1 で積む縦の合計は、1段ヘッダー(48pt)で 510pt:
'   48 + 12 + 68(プロフィール) + 18 + 240(タイル4段) + 15(行境界の丸め)
'   + 75(バッジ見出し+4行) + 10(フッター上) + 24(フッター+下余白) = 510
Private Sub TestCompress()
    modTestRunner.Check "Hub必要量_1段ヘッダーで510pt", _
        (modViewport2.HubNeedY(48, 68, 240) = 510), _
        "実際=" & modViewport2.HubNeedY(48, 68, 240)
    modTestRunner.Check "Hub必要量_2段ヘッダー(96pt)では558pt", _
        (modViewport2.HubNeedY(96, 68, 240) = 558)

    modTestRunner.Check "圧縮_窓700ptには収まるので係数1(何もしない)", _
        (modViewport2.CompressFactor(700, 510) = 1)
    modTestRunner.Check "圧縮_窓518ptは境界ちょうど(510=510)で係数1", _
        (modViewport2.CompressFactor(518, 510) = 1)
    modTestRunner.Check "圧縮_実機の窓499ptでは 491/510 へ縮める", _
        (Abs(modViewport2.CompressFactor(499, 510) - 491 / 510) < 0.000001), _
        "実際=" & modViewport2.CompressFactor(499, 510)
    modTestRunner.Check "圧縮_窓450ptでは 442/510 へ縮める", _
        (Abs(modViewport2.CompressFactor(450, 510) - 442 / 510) < 0.000001)
    modTestRunner.Check "圧縮_極端に低い窓300ptでも下限0.78で止める", _
        (Abs(modViewport2.CompressFactor(300, 510) - 0.78) < 0.000001)
    modTestRunner.Check "圧縮_必要量0や負(退化入力)は係数1", _
        (modViewport2.CompressFactor(499, 0) = 1 And _
         modViewport2.CompressFactor(499, -10) = 1)

    ' バッジ帯はセルの行(15pt固定)なので係数では縮まない。段数を落とす。
    modTestRunner.Check "バッジ段数_圧縮なし(1.0)は4行", _
        (modViewport2.BadgeRowsFor(1) = 4)
    modTestRunner.Check "バッジ段数_0.95は4行(境界)", _
        (modViewport2.BadgeRowsFor(0.95) = 4)
    modTestRunner.Check "バッジ段数_0.94は3行(境界の外)", _
        (modViewport2.BadgeRowsFor(0.94) = 3)
    modTestRunner.Check "バッジ段数_下限0.78でも3行(2行までは削らない)", _
        (modViewport2.BadgeRowsFor(0.78) = 3)
End Sub

' ----------------------------------------------------------------------------
' S3: モードごとの吸収列(modViewport2.ShelfPadCol・2段階Fitの廃止)
' ----------------------------------------------------------------------------
Private Sub TestShelfPadCol()
    modTestRunner.Check "吸収列_一覧表は本文の最終列J", _
        (modViewport2.ShelfPadCol("table", "N") = "J")
    modTestRunner.Check "吸収列_大文字/前後空白でも一覧表はJ", _
        (modViewport2.ShelfPadCol("TABLE", "N") = "J" And _
         modViewport2.ShelfPadCol(" table ", "N") = "J")
    modTestRunner.Check "吸収列_ギャラリーは帯の最終列N", _
        (modViewport2.ShelfPadCol("gallery", "N") = "N")
    modTestRunner.Check "吸収列_みんなの解決事例も帯の最終列N", _
        (modViewport2.ShelfPadCol("shared", "N") = "N")
    modTestRunner.Check "吸収列_未描画(空文字)はNへ倒す", _
        (modViewport2.ShelfPadCol("", "N") = "N")
End Sub

Public Sub RunAll22()
    On Error GoTo HScrollFail22
    TestHScrollNeeded
NextMoved22:
    On Error GoTo MovedFail22
    TestViewMoved
NextSb22:
    On Error GoTo SbFail22
    TestSbWidthFrom
NextGrid22:
    On Error GoTo GridFail22
    TestGalleryGrid
NextGap22:
    On Error GoTo GapFail22
    TestDashGap
NextFits22:
    On Error GoTo FitsFail22
    TestFitsInView
NextComp22:
    On Error GoTo CompFail22
    TestCompress
NextPad22:
    On Error GoTo PadFail22
    TestShelfPadCol
NextRun23:
    On Error GoTo Run23Fail22
    modTestsPure23.RunAll23
NextDone22:
    On Error GoTo 0
    Exit Sub

HScrollFail22:
    modTestRunner.Check "TestHScrollNeeded(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextMoved22
MovedFail22:
    modTestRunner.Check "TestViewMoved(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextSb22
SbFail22:
    modTestRunner.Check "TestSbWidthFrom(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGrid22
GridFail22:
    modTestRunner.Check "TestGalleryGrid(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGap22
GapFail22:
    modTestRunner.Check "TestDashGap(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextFits22
FitsFail22:
    modTestRunner.Check "TestFitsInView(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextComp22
CompFail22:
    modTestRunner.Check "TestCompress(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextPad22
PadFail22:
    modTestRunner.Check "TestShelfPadCol(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRun23
Run23Fail22:
    modTestRunner.Check "modTestsPure23.RunAll23(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone22
End Sub
