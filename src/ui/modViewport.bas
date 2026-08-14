Attribute VB_Name = "modViewport"
Option Explicit

' ============================================================================
' modViewport - 画面の「見える範囲」に幾何を合わせる共通部品
'               (2026-08-05 R18-3b → 2026-08-06 R19-1a/1d で全面改訂)。
' ----------------------------------------------------------------------------
' ホイールの停止線は何で決まるか(2026-08-14 R32実機プローブで確定。
' R19/R29/R31の記述をここで訂正する。法則が2回反転しているので、次に
' 書き換えるときは必ず実機プローブの生データを添えること):
'   停止線 S = B + k
'     B = 焼き付いた UsedRange の末尾行(行高を明示した行・塗った行は
'         Excel から見れば「使用済み」)
'     k = 1画面ぶんの行数(最終使用行が【画面上端】に来るまで転がれる)
'   実機値: Hub境界32行→約65行目で停止(報告61)/ Dash境界34行→約68行目(63)。
'
'   【ScrollArea はホイールを止めない】(R32実機実証。R29の「止める」は誤帰属)。
'   "A1:L15" にしても60行付近まで転がる。確実に止まるのはセル選択と
'   スクロールバーだけ。掛ける価値はあるが余白対策として当てにしない。
'
'   【塗りでは原理的に解けない】: S − B = k は塗る深さ d に依存しない。
'   深く塗れば B が増え停止線 S も同じだけ下がる=未塗り到達域は常に k行残る。
'   R18〜R27の7ラウンドが同じ壁に当たり続けた真因。打てる手は2つだけ:
'     (1) B を小さくする=内容の実下端+1行より下を削除(解放できるのは
'         Rows.Delete だけ。ClearFormats も保存も UseStandardHeight も効かない
'         =R30実機実証)―― ReleaseSheetRowsBelow
'     (2) セルを1つも使わずに地を色づける=背景画像(modBackdrop。R32 W4-5)。
'         使用済み範囲を増やさないので B が増えない。
'   ThisWorkbook.Styles("Normal")/("標準") 方式は実機で 1004(使用不能)。
'
' 右の余白は寸法の問題:
'   列幅の合計がウィンドウの可視幅より狭ければ、その差は必ず白く残る。
'   消せるのは列幅を可視幅に合わせることだけ(FitBandToViewport)。
'
' 設計の鉄則:
'   ・右端の値は ContentRight 1本を単一情報源にする。画面ごとに別の式を
'     書かない(帯・ピル・ボタンの右端が3段にズレていた原因)。
'   ・pt と ColumnWidth(文字数単位)の換算はフォント・DPIで変わる。
'     机上計算せず、必ず実セル(Columns(n).Width / .ColumnWidth)から取る。
'     Shape座標を必ずセル幾何から出すのと同じ理由。
'   ・範囲は「実際に描いた最大到達点+余白」で決める。きつく締めると
'     ボタンが境界の外に取り残され、憲章§3-1「押せるものは必ず反応する」を
'     破る(見えない=押せない)。迷ったら広い側へ倒す。
'   ・失敗しても画面を落とさない(全て On Error Resume Next 配下)。
'
' 依存: modUIMain.ViewportWidth / modChrome.BarWidth / modLog.LogUsage。
'   純関数(PadPtNeeded / PadUnitsRefine / BoundBottomY / RightEdgeAt)は他を呼ばず、
'   LibreOffice の純ロジックテストで固定する。
' ============================================================================

' 縦スクロールバーが食う幅(pt)。R21-S2で決め打ちをやめ、実測から校正する
' modViewport2.ScrollbarW() が単一情報源になった(ここは残していない)。
' 吸収列の最小幅(pt)。0にすると列が消えて「最終列の右は灰色」になる。
Private Const MIN_PAD_PT As Double = 8
' 内容下端に足す余白(pt)。R21-S4: 24→8。24ptは「+24pt→行切り上げ→最低1画面」の
' 三重切り上げの1段目で、内容が窓に収まっている画面でも境界を必ず窓の外へ
' 押し出していた(実機第8報⑦の独立欠陥その3)。
Private Const BOTTOM_PAD As Double = 8
' 自分のブックが前面でないときに ViewportHeight が返す既定値(pt)。R19H FB-5。
Private Const DEFAULT_VIEW_H As Double = 600

' 行解放が何があっても残す最上部の行数。どの画面もヘッダー/ツールバーが
' 行1〜12に載るので、boundRow が異常値(0や負)で降ってきてもここから上は
' 絶対に消さない(ReleaseRange の minRow へ渡す下限)。
Private Const MIN_KEEP_ROW As Long = 12

' usage_log("viewport") を画面ごとに1セッション1回だけ出すためのメモ(R19-1e)。
Private mLoggedScreens As String

' RowAt の増分走査メモ(R19H FB-7)。詳細は RowAt の見出しコメント参照。
Private mRowMemoSheet As String
Private mRowMemoRow As Long
Private mRowMemoTop As Double

' 窓リサイズ後の再フィット(R20-1f)。詳細は最下部の OnWindowResized 一式を参照。
' デバウンス間隔(秒)。ドラッグ中は WindowResize が毎フレーム飛ぶので、
' 手が止まってから組み直す。0.7秒は「もう動かさない」と判断できて、かつ
' 待たされた感じがしない値。
Private Const REFIT_DELAY_SEC As Double = 0.7

Private mRefitAt As Date        ' 予約済み OnTime の時刻(0=予約なし)
Private mRefitArmed As Boolean  ' 予約が生きているか(mRefitAtの有効フラグ)
Private mRefitRunning As Boolean ' 再フィット実行中(再入抑止)
Private mRefitWaited As Boolean  ' Busyで1回だけ待ち直したか

' ----------------------------------------------------------------------------
' ApplyScrollBound - ws.ScrollArea を設定する。空文字なら制限を外す。
' ----------------------------------------------------------------------------
'   例) modViewport.ApplyScrollBound ws, "A1:L60"(areaAddr が空なら制限解除)
'   ScrollArea はブックに保存されないセッション限定のプロパティ。本アプリは
'   画面を出すたびに構築Subを必ず通る(冪等な全再構築)ので、そこに置けば
'   再適用の配線は要らない。シート保護下でも設定できるが、順序の事故
'   (保護→選択不能→FreezePanes失敗)を避けるため、呼び出し側は「幾何と
'   Shapeを置き終えた直後・保護をかける前」に呼ぶこと。
Public Sub ApplyScrollBound(ByVal ws As Worksheet, ByVal areaAddr As String)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.ScrollArea = areaAddr
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' FitBandToViewport - 帯(A列～最終列)の合計幅を可視幅ぴったりに合わせる。
' ----------------------------------------------------------------------------
'   bandAddr    : 画面が使う列帯の1行ぶん(例 "A1:L1")。吸収列も【含む】。
'   padColLetter: 余りを吸わせる列(例 "L")。帯の中にあれば位置は問わない。
'   minRightX   : 内容の実右端(pt)。可視幅がこれより狭くても帯はここまで
'                 確保する(内容がScrollAreaの外へ出て押せなくなるのを防ぐ)。
'
'   固定側が連続範囲とは限らない(チャットは A:J と L:M が固定で、間の K が
'   吸収列)ため、引数は「帯ぜんぶ」を受け取り、吸収列の現在幅を引いて
'   固定側の合計を出す。調査①班6章の案を非連続に対応させたもの。
Public Sub FitBandToViewport(ByVal ws As Worksheet, ByVal bandAddr As String, _
                             ByVal padColLetter As String, _
                             Optional ByVal minRightX As Double = 0)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next

    Dim u0 As Double: u0 = ws.Columns(padColLetter).ColumnWidth
    If u0 <= 0 Then
        ws.Columns(padColLetter).ColumnWidth = 1
        u0 = ws.Columns(padColLetter).ColumnWidth
    End If
    Dim w0 As Double: w0 = ws.Columns(padColLetter).Width
    Dim ptPerUnit As Double
    If u0 > 0 Then ptPerUnit = w0 / u0
    If ptPerUnit <= 0 Then Exit Sub          ' 換算できない端末では何もしない

    Dim fixedW As Double
    fixedW = ws.Range(bandAddr).Width - w0

    ' R21-S2: 目標は「可視セル幅 − 安全余裕2pt」。可視セル幅は決め打ちの
    ' SCROLLBAR_W=12 ではなく、UsableWidth と VisibleRange.Width の突合から
    ' 校正したバー幅で出す(modViewport2 が単一情報源)。
    Dim target As Double
    target = modViewport2.FitTarget()
    If target < minRightX Then target = minRightX

    ' R19H FA-1(A-H①): pt と ColumnWidth の関係は比例ではなく【アフィン】
    ' (Width = 傾き×ColumnWidth + セルの内側余白ぶんの下駄)。1点の実測から
    ' 出した比 w0/u0 には下駄が丸ごと乗っているため、基準列が狭いほど比が
    ' 大きく出て、必要な幅を吸い切れない(Hub の L=1.5 では必要幅の約70%
    ' しか吸わず、可視幅との差が50pt以上残っていた)。
    ' そこで「1回設定 → 実測を読み直す → 目標との差分だけ足し直す」の2段にする。
    ' 差分の換算には、2点(設定前・設定後)の実測から出した傾きを使う ――
    ' 下駄は差を取った時点で消えるので、この1回の補正で目標へ収束する。
    Dim need As Double: need = PadPtNeeded(target, fixedW, MIN_PAD_PT)
    ws.Columns(padColLetter).ColumnWidth = need / ptPerUnit
    Dim u1 As Double: u1 = ws.Columns(padColLetter).ColumnWidth   ' Excelが丸めた実値
    Dim w1 As Double: w1 = ws.Columns(padColLetter).Width
    ws.Columns(padColLetter).ColumnWidth = PadUnitsRefine(need, u0, w0, u1, w1)
    ' R21-S2: 事後検証。丸め・最小幅の都合で帯が target を超えたら差分で詰め直す
    ' (「帯実幅 ≤ 可視幅」を出口で必ず成立させる=横スクロールが構造的に不能)。
    modViewport2.FitVerify ws, bandAddr, padColLetter, target
    On Error GoTo 0
End Sub

' PadUnitsRefine - 1回目の設定結果から補正後の ColumnWidth を出す。純関数
'   (ゴールデン対象。狭い基準列1.5と広い基準列13の両方を modTestsPure16 が固定)。
'   needPt : 吸収列に要る幅(pt)= PadPtNeeded の戻り値
'   u0/w0  : 設定【前】の (ColumnWidth, 実測Width)
'   u1/w1  : 設定【後】の (ColumnWidth, 実測Width)
'   傾き = (w1-w0)/(u1-u0)。アフィンの下駄が差で消えるので、これが唯一
'   正しい換算係数になる。2点が使えない(1回目で幅が動かなかった端末・
'   等しいColumnWidth)ときだけ、従来どおりの実測比 w1/u1 へ退化する
'   (補正が効かないだけで、従来より悪くはならない)。
Public Function PadUnitsRefine(ByVal needPt As Double, _
                               ByVal u0 As Double, ByVal w0 As Double, _
                               ByVal u1 As Double, ByVal w1 As Double) As Double
    PadUnitsRefine = u1
    If u1 <= 0 Then Exit Function
    Dim slope As Double
    If Abs(u1 - u0) > 0.001 Then slope = (w1 - w0) / (u1 - u0)
    If slope <= 0 Then slope = w1 / u1
    If slope <= 0 Then Exit Function
    Dim u As Double: u = u1 + (needPt - w1) / slope
    ' 0にすると列が消え、最終列の右が灰色の非セル領域になる(PadPtNeededと同じ理由)。
    If u < 0.05 Then u = 0.05
    ' R20-1c: ColumnWidth の上限は255(Excel仕様)。本文の最終列に余りを吸わせる
    ' 使い方(本棚のメモ列J・登録フォームのH列)を足したので、超広い窓では
    ' 255を超える値が出うる。超えた値の代入は1004で、On Error Resume Next 配下の
    ' ここでは「列幅が前回のまま」という無言の失敗になる。手前で頭打ちにする。
    If u > 255 Then u = 255
    PadUnitsRefine = u
End Function

' PadPtNeeded - 吸収列に要る幅(pt)。純関数(ゴールデン対象)。
'   帯の合計を targetW にするには吸収列を targetW-fixedW にすればよい。
'   ただし負や極小にはしない(列が消えると最終列の右が灰色の非セル領域に
'   なり、余白の見た目はむしろ悪化する)。
Public Function PadPtNeeded(ByVal targetW As Double, ByVal fixedW As Double, _
                            ByVal minPad As Double) As Double
    PadPtNeeded = targetW - fixedW
    If PadPtNeeded < minPad Then PadPtNeeded = minPad
End Function

' ----------------------------------------------------------------------------
' ContentRight - 右端の単一情報源(pt)。帯・ピル・主ボタンは全部これを見る。
' ----------------------------------------------------------------------------
'   bandAddr : 帯の1行ぶん(例 "A1:M1")
'   rightPad : 帯の内側に空ける余白。0なら帯の右端そのもの。
'   modKnowledge.ToolbarContentRight(R14-2a)と同じ思想の汎用版。
Public Function ContentRight(ByVal ws As Worksheet, ByVal bandAddr As String, _
                             ByVal rightPad As Double) As Double
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Dim rng As Range: Set rng = ws.Range(bandAddr)
    If rng Is Nothing Then Exit Function
    ContentRight = RightEdgeAt(rng.Left, rng.Width, modUIMain.ViewportWidth(), rightPad)
    On Error GoTo 0
End Function

' RightEdgeAt - ContentRight の算数だけを取り出した純関数(ゴールデン対象)。
'   帯が可視幅より広い端末(列幅を詰め切れなかった場合)では可視幅で頭打ちに
'   する。頭打ちの算数は modChrome.BarWidth が単一情報源(2つ持つとズレる)。
Public Function RightEdgeAt(ByVal bandLeft As Double, ByVal bandW As Double, _
                            ByVal viewportW As Double, ByVal rightPad As Double) As Double
    RightEdgeAt = bandLeft + modChrome.BarWidth(bandW, viewportW, rightPad) - rightPad
End Function

' ----------------------------------------------------------------------------
' BoundAddr - 画面の実使用範囲 "A1:<吸収列><行>" を実測で組み立てる。
' ----------------------------------------------------------------------------
'   padColLetter : 帯の最終列(FitBandToViewport の吸収列と同じ)
'   contentBottom: 描き終えた内容の実下端(pt)
'   maxRow       : その画面で許す最大行(行高が壊れていても暴走しない)
'   塗り範囲・ScrollArea・Lockedの範囲は全てこの1本から取る。
'
'   R21-S4(実機第8報⑦の独立欠陥その3): 従来は内容が窓に収まる画面でも
'   「+24pt → RowAt(=切り上げ) → 最低1画面」の三重切り上げを通っており、
'   境界の下端が必ず窓高より下になっていた ―― 内容が1画面に収まっているのに
'   毎回スクロール余地が残る、を構造的に作っていた。
'   収まる画面は【切り下げ】(RowAtFloor)で窓高ちょうどに止め、収まらない
'   画面だけ従来どおり内容の実下端(+8pt)まで伸ばす。
'   R21H F6(敵対的レビュー確定): 戻り値(ScrollArea用)は上のとおりRowAtFloor
'   (切り下げ)でよいが、【塗り】まで同じ切り下げ値を使うと、窓高が行境界に
'   ちょうど乗らない端末で最後の部分行(最大1行ぶん=13〜15pt)が塗り残る
'   (darkテーマで白帯として見える)。塗りはRowAt(切り上げ)で部分行を含めて
'   よい ―― ScrollAreaと違い、塗りが1行分広くてもホイールで到達可能な範囲は
'   増えない。Optional paintAddr へ「塗り用(ceil)」を別途返す(戻り値=
'   ScrollArea用は変えない。呼ばない既存呼び出し元はそのまま=後方互換)。
Public Function BoundAddr(ByVal ws As Worksheet, ByVal padColLetter As String, _
                          ByVal contentBottom As Double, ByVal maxRow As Long, _
                          Optional ByRef paintAddr As String) As String
    If ws Is Nothing Then Exit Function
    Dim viewH As Double: viewH = ViewportHeight()
    If FitsInView(contentBottom, viewH, BOTTOM_PAD) Then
        BoundAddr = "A1:" & padColLetter & RowAtFloor(ws, viewH, maxRow)
        paintAddr = "A1:" & padColLetter & RowAt(ws, viewH, maxRow)
    Else
        BoundAddr = "A1:" & padColLetter & _
                    RowAt(ws, BoundBottomY(contentBottom, viewH, BOTTOM_PAD), maxRow)
        paintAddr = BoundAddr
    End If
End Function

' PrimePaint - Hub描画前の先行塗り(R31 F11)。modViewport2.SeedBurn(hdrH渡し)
'   が行1をヘッダー実高へ合わせ済みの前提で呼ぶこと(呼ぶ前が15pt均しの
'   ままだと境界を切り下げすぎ、本描画後の実境界より深く塗ってUsedRangeが
'   毎描画で最終境界を上回るF2-2を再発する)。塗りはBoundAddrのpaintAddr
'   (切り上げ)を使い、塗り下端が常にRowAtFloor(viewH)以上になることを
'   式で保証する(切り下げだと窓下端に未塗り帯=darkテーマの白帯が出る)。
Public Sub PrimePaint(ByVal ws As Worksheet, ByVal padCol As String, ByVal maxRow As Long)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Dim addr As String, paintAddr As String
    addr = BoundAddr(ws, padCol, 0, maxRow, paintAddr)
    ws.Range(paintAddr).Font.Name = "Yu Gothic UI"
    ws.Range(paintAddr).Font.Size = 10
    ws.Range(paintAddr).Interior.Color = modUI.UiColor("bg")
    On Error GoTo 0
End Sub

' FitsInView - 内容が窓に収まっているか。純関数(ゴールデン対象)。
'   収まっている=境界を窓高ちょうどに切り下げてよい(スクロール余地ゼロ)。
Public Function FitsInView(ByVal contentBottom As Double, ByVal viewportH As Double, _
                           ByVal pad As Double) As Boolean
    FitsInView = (contentBottom <= viewportH - pad)
End Function

' RowAtFloor - 下端が y 【以下】に収まる最後の行(部分行を含まない)。
'   RowAt(yを含む最小の行=切り上げ)の対で、こちらは切り下げ。
'   境界を窓高に合わせるときに使う ―― 切り上げると必ず窓を1行ぶん超え、
'   その1行があるだけで縦スクロールが生きてしまう。
'   行高は端末依存なので、ここも RowAt と同じく実セル幾何だけを読む。
'   1行目の下端すら y を超える(窓が極端に低い)ときは 1 を返す。
Public Function RowAtFloor(ByVal ws As Worksheet, ByVal y As Double, ByVal maxRow As Long) As Long
    RowAtFloor = 1
    If maxRow < 1 Then maxRow = 1
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Dim i As Long
    For i = 1 To maxRow
        If ws.Cells(i, 1).Top + ws.Cells(i, 1).Height > y Then Exit For
        RowAtFloor = i
    Next i
    On Error GoTo 0
End Function

' BoundBottomY - 境界の下端Y(pt)。純関数(ゴールデン対象)。
'   内容+余白。ただし最低でも1画面ぶん確保する ―― 内容が1画面に満たない
'   ときに下へ1ptも行けない画面は窮屈で、かつホイールで行けてしまう先が
'   塗られていない「白い断崖」になる。
Public Function BoundBottomY(ByVal contentBottom As Double, ByVal viewportH As Double, _
                             ByVal pad As Double) As Double
    BoundBottomY = contentBottom + pad
    If BoundBottomY < viewportH Then BoundBottomY = viewportH
End Function

' RowAt - y(pt)を含む最小の行番号。上限に張り付いたら上限を返す。
'   行高はフォントやDPIで変わり、ptから行番号を机上計算することはできない
'   (「1行=何pt」は端末依存)。実際のセル幾何(Top/Height)を読んで決める
'   ―― Shape座標をセル幾何から出すのと同じ理由。
'
'   R19H FB-7(A-L⑭): 毎回1行目から数え直していた。チャットは【バブル1個ごと】
'   に ExtendChatBand→BoundAddr→ここ を通るので、会話が伸びるほど1発言あたり
'   数百回のCOM往復になる(NEXUS_MAX_ROW ぶんの Cells(i,1).Top/.Height)。
'   前回返した行を覚えておき、今回の y がその行の上端より下なら【そこから】
'   再開する(会話は下へ伸びる一方なので、実際の走査は数行で終わる)。
'   メモが今の行高と食い違っていないかは、再開の前に1回だけ実測で確かめる
'   (EnsureLayout が行高を組み直した直後は先頭から数え直す)。
'
'   R20-1e: Public 化した。Hub のバッジ帯は「行高15pt固定」を前提に
'   CLng(下端pt / 15) + 2 で行を出していたが、Hub の行1はヘッダー帯と同じ
'   48pt(ナビが2段に折り返すと96pt)で、15pt換算では毎回2〜3行ぶん下へ
'   ずれる。pt→行の換算を机上でやってよい場所は1つも無いので、実測を
'   持っているここを唯一の口にする(モジュール冒頭の設計の鉄則)。
Public Function RowAt(ByVal ws As Worksheet, ByVal y As Double, ByVal maxRow As Long) As Long
    If maxRow < 1 Then maxRow = 1
    RowAt = maxRow
    If ws Is Nothing Then Exit Function
    On Error Resume Next

    Dim startRow As Long: startRow = 1
    If mRowMemoRow >= 1 And mRowMemoRow <= maxRow Then
        If StrComp(mRowMemoSheet, ws.Name, vbTextCompare) = 0 And mRowMemoTop < y Then
            ' 「上端 < y」なら答えはメモの行以降にしか無い(それより上の行は
            ' 下端が上端以下=y未満で、条件を満たさない)。等号を含めないのは、
            ' ちょうど境目のときに1行手前が正解になるため。
            If ws.Cells(mRowMemoRow, 1).Top = mRowMemoTop Then startRow = mRowMemoRow
        End If
    End If

    Dim i As Long
    For i = startRow To maxRow
        If ws.Cells(i, 1).Top + ws.Cells(i, 1).Height >= y Then
            RowAt = i
            mRowMemoSheet = ws.Name
            mRowMemoRow = i
            mRowMemoTop = ws.Cells(i, 1).Top
            Exit For
        End If
    Next i
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' ViewportHeight - ウィンドウの実可視高(pt)。ViewportWidth(modUIMain)の縦版。
' ----------------------------------------------------------------------------
'   異常値でレイアウト計算全体を壊さないよう 200〜2000pt へクランプする
'   (R20-1a: 上限は1200ptだった。4Kの縦置きや高さ1440pxのモニタでは実可視高が
'    1200ptを超え、そこで頭打ちにすると BoundBottomY の「最低1画面」が
'    画面より短くなって、下端に白い断崖が残る)。
'
'   R19H FB-5(A-L⑫): 他のブックが前面のときは【一切測らない】。従来は
'   UsableHeight だけを ThisWorkbook で守り、その下の VisibleRange.Height は
'   無条件に走っていたため、作業用Excelなど別ブックが前面のまま起動時の
'   突合や自動同期からここへ来ると、他人の窓の高さで自分の塗り・境界を
'   決めてしまっていた(同居時=実機第6報⑤では常態)。測れないときは既定値
'   (DEFAULT_VIEW_H)を返す。画面は次に前面へ来たとき必ず冪等に組み直される
'   ので、ここで嘘の実測を返すより既定値のほうが安全側。
Public Function ViewportHeight() As Double
    Dim h As Double
    On Error Resume Next
    If ActiveWorkbook Is ThisWorkbook Then
        h = ActiveWindow.UsableHeight
        If h <= 0 Then h = ActiveWindow.VisibleRange.Height
    End If
    On Error GoTo 0
    If h <= 0 Then h = DEFAULT_VIEW_H
    If h < 200 Then h = 200
    If h > 2000 Then h = 2000
    ViewportHeight = h
End Function

' ----------------------------------------------------------------------------
' R31 W2-1: 境界より下の「使用済み行」を解放する(Hub/Dash/本棚/チャット共通)
' ----------------------------------------------------------------------------
' R20-1d の ResetRowsBelow(UseStandardHeight=True)はここに在ったが、R30の
' 実機実証で「行高の明示を取り消しても焼き付きは解放されない」ことが確定した
' ため撤去した。解放できるのは Rows.Delete だけ(冒頭の見出し参照)。
' 実体をここへ置く理由: 本体を modViewport2 へ入れると同モジュールが
' 28,000字の警告線を越える(憲章§4-6。純関数のSeedRowCapだけ向こうに置いた)。
' 呼び出し側(modHub/modDash/modKnowledge)は1行で呼ぶ。

' ReleaseSheetRowsBelow - boundRow より下の使用済み行を Rows.Delete で解放(冪等)。
'   ws        : 対象シート
'   boundRow  : 残す下端行(境界の最終行+余裕1行)。これより下を消す。
'   maxRow    : 旧版が焼き得た上限(Hub=200 / Dash=DASH_ROWS / 本棚=412 /
'               チャット=NEXUS_MAX_ROW)。使用済みがそこまで届いていなくても
'               この行までは消す(旧版の焼き付けの取りこぼし救済)。
'   scrollAddr: 空でなければ削除の後に ScrollArea を掛け直す(R30 F8の作法。
'               行削除が ScrollArea 設定の【後】に走る経路のため)。
'   logTag    : usage_log の区分。空ならシート名(どの画面かをログで識別する)。
'   minRow    : これより上は何があっても消さない(既定=MIN_KEEP_ROW)。
'   Rows.Delete はモーダルを出さないので modUiLock.AlertsOff/On は要らない。
Public Sub ReleaseSheetRowsBelow(ByVal ws As Worksheet, ByVal boundRow As Long, _
                                 ByVal maxRow As Long, Optional ByVal scrollAddr As String = "", _
                                 Optional ByVal logTag As String = "", _
                                 Optional ByVal minRow As Long = MIN_KEEP_ROW)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Dim lastUsed As Long
    lastUsed = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    Dim fromRow As Long, toRow As Long
    If modViewport2.ReleaseRange(boundRow, lastUsed, minRow, maxRow, fromRow, toRow) Then
        ' 削除範囲に掛かるShapeが「移動/サイズ変更する」設定のままだと、行削除で
        ' 縮む・消える。描画側のFreezeShapePlacementは削除より後に走る経路
        ' (Hub/Dash/gallery)があるので、ここで先に絶対配置へ固定する(冪等)。
        modUI.FreezeShapePlacement ws
        ws.Rows(fromRow & ":" & toRow).Delete
        ' R31 Fix波F9: ok=の判定をErr.Numberの直接参照からRows.Delete直後の
        ' 退避へ変更(CLAUDE.mdの「ログの前にErr退避」作法)。従来は
        ' FreezeShapePlacementの内部On Errorが先にErrをクリアするという
        ' 他関数の実装詳細に依存していた(敵対的レビュー1周目MINOR裁定)。
        Dim delOk As Boolean: delOk = (Err.Number = 0)
        ' R30 F9: 削除実行時のみの観測ログ(冪等な通常起動では出ない)。
        Dim tg As String: tg = logTag
        If LenB(tg) = 0 Then tg = ws.Name
        modLog.LogUsage "row_release", tg, "from=" & fromRow & " to=" & toRow & _
            " lastUsed=" & lastUsed & " ok=" & delOk
        Err.Clear
        ' 削除だけでは内部使用範囲(xlCellTypeLastCell)が縮まらない端末がある。
        ' UsedRange を1回参照して再計算させる(戻り値は捨てる)。
        lastUsed = ws.UsedRange.Rows.Count
        If LenB(scrollAddr) > 0 Then ApplyScrollBound ws, scrollAddr
    End If
    On Error GoTo 0
End Sub

' ClampD - 下限/上限で挟む(pt・幅・カード数の共通算数)。純関数。
'   lo > hi の壊れた指定でも lo を返す(呼び出し側の防御を1つ減らす)。
Public Function ClampD(ByVal v As Double, ByVal lo As Double, ByVal hi As Double) As Double
    ClampD = v
    If ClampD > hi Then ClampD = hi
    If ClampD < lo Then ClampD = lo
End Function

' ----------------------------------------------------------------------------
' GalleryColsFor - 幅 bandW の帯に、左端 leftX から等幅カードが何列入るか。
' ----------------------------------------------------------------------------
'   【非推奨】R21-S3 でギャラリーは列数とカード幅の両方を弾性にしたため
'   (modViewport2.GridColsFor / GridCardW)、この「カード幅固定・最後の右に
'   GAP が残る」式の呼び出し元は無い。算数と境界ゴールデン(modTestsPure19)は
'   固定カード幅の画面が再び要るときのために残す。
'   純関数(ゴールデン対象)。カードは左寄せのまま列数だけ増やす方式なので、
'   最後のカードの右に GAP ぶんの余白が残る前提で数える。
'   minCols/maxCols で挟むのは、狭い窓で0列になって画面が空になるのと、
'   広い窓でカードが並びすぎて一覧性が落ちるのを同時に防ぐため。
Public Function GalleryColsFor(ByVal bandW As Double, ByVal leftX As Double, _
                               ByVal cardW As Double, ByVal gap As Double, _
                               ByVal minCols As Long, ByVal maxCols As Long) As Long
    GalleryColsFor = minCols
    If cardW + gap <= 0 Then Exit Function
    Dim n As Long
    n = Int((bandW - leftX - gap) / (cardW + gap))
    If n < minCols Then n = minCols
    If n > maxCols Then n = maxCols
    GalleryColsFor = n
End Function

' ----------------------------------------------------------------------------
' LogViewport - 実機の可視幅×可視高を1画面につき1セッション1回だけ記録する。
' ----------------------------------------------------------------------------
' R19-1e: 全ての幾何判断は「実機の可視幅は600〜900pt」という推定の上に
' 立っているが、その推定を裏付ける実測値を一度も採っていない(憲章§4-2)。
'
' R21-S7: 記録を「窓幅×窓高」の2値から【5値】へ拡げた。余白の報告が来たとき、
' 帯・中身・境界のどの層が破れたのかを次の1報で一意に決めるため:
'   w   = 窓の可視幅(ViewportWidth)
'   vis = セルを置ける可視幅(可視幅 − 縦スクロールバー幅。Fitの目標の素)
'   band= 帯の実幅   … w/vis より大きければ S2 の破れ(横スクロールが生きる)
'   right= 中身の右端 … band と離れていれば S3 の破れ(中身が張っていない)
'   h/bottom = 窓高と境界の高さ … bottom>h なら S4 の破れ(縦スクロールが残る)
' フィット直後に毎回呼んでよい。同じ5値の重複だけを抑止するので、幾何が
' 動かない再描画ではログが増えない(1画面1セッション1回の旧方式では、
' 窓をリサイズした後の値が二度と記録されなかった)。
' 2026-08-10(R25-4b FA-R25-4b): uw(ActiveWindow.UsableWidthの生値)を末尾へ
' 追加。vis(=ViewportWidth−ScrollbarW)がフォールバック値(modViewport2.
' SB_FALLBACK)で計算されたのか、実測のスクロールバー幅で計算されたのかは
' 5値だけでは切り分けられない(実機で vis=1008/996 が混在した報告の原因が
' 特定できなかった)。UsableWidthの生値があれば、vis側の式(SbWidthFrom)を
' 逆算して実測/fallbackのどちらかを事後判定できる。取得できない環境
' (ActiveWorkbookがThisWorkbookでない等)では0を記録する(modViewport2.
' ScrollbarWと同じOn Error Resume Next+ActiveWorkbook Is ThisWorkbook guard)。
Public Sub LogViewport(ByVal screenName As String, Optional ByVal bandW As Double = 0, _
                       Optional ByVal rightX As Double = 0, _
                       Optional ByVal boundBottom As Double = 0)
    On Error Resume Next
    Dim uw As Double: uw = 0
    If ActiveWorkbook Is ThisWorkbook Then uw = ActiveWindow.UsableWidth
    If Err.Number <> 0 Then uw = 0
    Err.Clear
    Dim d As String
    d = "w=" & CLng(modUIMain.ViewportWidth()) & " vis=" & CLng(modViewport2.VisibleCellW()) & _
        " band=" & CLng(bandW) & " right=" & CLng(rightX) & _
        " h=" & CLng(ViewportHeight()) & " bottom=" & CLng(boundBottom) & _
        " uw=" & CLng(uw)
    Dim key As String: key = "|" & screenName & " " & d & "|"
    If InStr(mLoggedScreens, key) > 0 Then Exit Sub
    If Len(mLoggedScreens) > 4000 Then mLoggedScreens = ""   ' 際限なく持たない
    mLoggedScreens = mLoggedScreens & key
    modLog.LogUsage "viewport", screenName, d
    On Error GoTo 0
End Sub

' ============================================================================
' R20-1f: 窓のリサイズ後に、今出ている画面を自動で組み直す(デバウンス付き)
' ============================================================================
' なぜ要るか: このアプリの幾何は全て「今の可視幅・可視高」から決まる
' (FitBandToViewport / ContentRight / BoundAddr)。ところが再計算の機会は
' 「画面を開いたとき」だけで、開いたあとに窓を広げても列幅もカードも
' 動かない ―― R19 までの修正が実機で効いて見えなかった最大の理由がこれ。
'
' 実体をここに置く理由: 呼び口は modUI.OnWindowResized が素直だが、modUI は
' 30,000字上限に対する余裕が乏しい。ThisWorkbook のイベントから直接
' modViewport.OnWindowResized を呼ぶ(仕様 R20-1f)。
'
' 危険なのは Application.OnTime の予約が残ること。予約を残したままブックを
' 閉じると、Excel は【時刻が来たときにそのブックを開き直して】まで実行しようと
' する(§9・modShelfSync の自動同期で実際に踏んだ事故)。だから:
'   (i)   Auto_Close の経路(modUI.RestoreExcelUI)から必ず CancelRefit を呼ぶ
'   (ii)  予約名はブック名で修飾する(新旧2版が併存すると無修飾では
'         名前解決先が決まらない。modShelfSync.TickProcName と同じ作法)
'   (iii) 他のブックが前面のときは予約もしないし実行もしない
'   (iv)  連続リサイズでは【必ず先に解除してから】入れ直す(多重登録ゼロ)
' 32bit Excel でも OnTime/Window イベントの作法は同じで、必要なのは
' 「予約時刻を覚えて完全一致で解除する」ことだけ(APIは一切使わない)。

' Busy3 - tick経路がRefitActionへ渡すisBusyの組み立て(純関数・R20H FA-2の
'   ゴールデン対象)。単純なOrだが、呼び出し側(ViewportRefitTick)が3情報源
'   (modUiLock/modShelf/modShelfSync)のうち1つでも足し忘れると再フィットが
'   busy中に走ってしまう契約なので、組み立てそのものを1本にまとめてここで
'   固定する(呼び出し側は3引数をそのまま渡すだけになる)。
Public Function Busy3(ByVal lockBusy As Boolean, ByVal shelfBusy As Boolean, _
                      ByVal shelfSyncBusy As Boolean) As Boolean
    Busy3 = lockBusy Or shelfBusy Or shelfSyncBusy
End Function

' RefitAction - 状態遷移だけを取り出した純関数(ゴールデン対象)。
'   OnTime も ActiveWorkbook も見ないので、LibreOffice の純ロジックテストで
'   「多重登録しない」「Busyでは1回だけ待ち直す」「他ブック前面では何もしない」
'   「自分の再描画が起こしたリサイズは無視する」を固定できる。
'   ev       : "resize"(窓が動いた) / "tick"(予約時刻が来た)
'   isMine   : 自分のブックが前面か
'   isRunning: 再フィットの実行中か(再入)
'   isBusy   : 取込・同期・仕上げ・回答生成のいずれかで処理中か(2026-08-06
'              R20H FA-2: modUiLock.IsBusy()単独では取込・同期・仕上げ中の
'              DoEventsで再入し得るTrueを取りこぼす。呼び出し側はBusy3で
'              modUiLock/modShelf/modShelfSyncの3情報源をORして渡す。
'              本関数自体は単一のBoolean判定のまま無変更)
'   waited   : Busyのために既に1回待ち直したか
'   戻り値   : "none"(何もしない) / "schedule"(予約し直す) / "run"(組み直す)
Public Function RefitAction(ByVal ev As String, ByVal isMine As Boolean, _
                            ByVal isRunning As Boolean, ByVal isBusy As Boolean, _
                            ByVal waited As Boolean) As String
    RefitAction = "none"
    If isRunning Then Exit Function
    If Not isMine Then Exit Function
    Dim e As String: e = LCase$(Trim$(ev))
    If e = "resize" Then
        RefitAction = "schedule"
        Exit Function
    End If
    If e <> "tick" Then Exit Function
    If isBusy Then
        ' 待つのは1回だけ。際限なく再予約すると、長い取込の間じゅう0.7秒おきに
        ' OnTime を張り替え続けることになる(予約1本の原則は守れても無駄が残る)。
        If Not waited Then RefitAction = "schedule"
        Exit Function
    End If
    RefitAction = "run"
End Function

' OnWindowResized - ThisWorkbook.Workbook_WindowResize から呼ばれる唯一の口。
'   ここでは【予約するだけ】。描画は必ず OnTime 経由に落とす(イベント中に
'   Shapeを作り直すとExcelが不安定になるため)。
Public Sub OnWindowResized()
    On Error Resume Next
    Dim act As String
    act = RefitAction("resize", (ActiveWorkbook Is ThisWorkbook), mRefitRunning, False, False)
    If act <> "schedule" Then Exit Sub
    mRefitWaited = False
    ScheduleRefit
    On Error GoTo 0
End Sub

' ScheduleRefit - 予約を1本だけ持つ(入れ直しは必ず解除してから)。
Private Sub ScheduleRefit()
    On Error Resume Next
    CancelRefit
    Dim t As Date: t = Now + REFIT_DELAY_SEC / 86400#
    Application.OnTime EarliestTime:=t, Procedure:=RefitProcName()
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub                            ' 予約できないのは非致命(次回に委ねる)
    End If
    mRefitAt = t
    mRefitArmed = True
    On Error GoTo 0
End Sub

' CancelRefit - 未消化の予約を取り消す。modUI.RestoreExcelUI(=Auto_Close の
'   経路)から必ず呼ぶ。予約が無い/既に発火済みでも安全(1004は握る)。
Public Sub CancelRefit()
    On Error Resume Next
    If mRefitArmed Then
        Application.OnTime EarliestTime:=mRefitAt, Procedure:=RefitProcName(), Schedule:=False
        Err.Clear
    End If
    mRefitArmed = False
    On Error GoTo 0
End Sub

' RefitProcName - 予約/解除で必ず同じ文字列になるように1箇所で組む。
Private Function RefitProcName() As String
    RefitProcName = "'" & ThisWorkbook.Name & "'!modViewport.ViewportRefitTick"
End Function

' ViewportRefitTick - OnTime のコールバック本体。Public 必須(OnTime は
'   Application.Run と同じ遅延バインドで Private を呼べない)。
Public Sub ViewportRefitTick()
    On Error Resume Next
    mRefitArmed = False
    ' 他ブックが前面なら何もしない(他人の窓幅で自分の帯を決めない)。取込・
    ' 同期・仕上げ・回答生成の最中も組み直さない(modApp.OnRefreshUI が
    ' BlockIfIngesting を置いているのと同じ理由)。busy判定は modUiLock だけ
    ' では取込中のDoEventsで漏れる(2026-08-06 R20H FA-2)ため、
    ' modShelf.IsBusy/modShelfSync.IsBusy も併せて見る。判断は
    ' RefitAction 1本に集約する(純関数側は無変更)。
    Dim act As String
    act = RefitAction("tick", (ActiveWorkbook Is ThisWorkbook), mRefitRunning, _
                      Busy3(modUiLock.IsBusy(), modShelf.IsBusy(), modShelfSync.IsBusy()), _
                      mRefitWaited)
    If act = "schedule" Then
        mRefitWaited = True
        ScheduleRefit
        Exit Sub
    End If
    If act <> "run" Then Exit Sub
    ' 2026-08-10(R27波3-2): ここから先の組み直しは【全アクションの関所】である
    ' modUiLock の内側で行う。mRefitRunning が止められるのは再フィット同士の
    ' 入れ子だけで、RefitActiveScreen(=Repaint/EnsureHubLayout)の中の DoEvents
    ' で押されたボタン(OnAction)が描画の途中から入れ子で走る窓が残っていた
    ' (窓リサイズ起点の入れ子再描画)。Enter を取っておけば、その OnAction 側が
    ' 先頭の Enter で弾かれる。
    ' Enter を【判定の後】に置く理由: 手続きの先頭で取ると直上の Busy3 が
    ' 自分自身のロックを busy と読んでしまい、再フィットが二度と走らない。
    ' さらに busy 時の作法は「1回だけ待ち直す」(schedule)であって拒否ではなく、
    ' そこを Enter の拒否に替えると、押してもいない拒否トーストが取込中に湧く。
    ' ここまで来た時点で3情報源とも busy でないことは判定済みで、Enter は
    ' 判定との隙間で誰かが取った場合の取りこぼしだけを拾う。
    If Not modUiLock.Enter() Then Exit Sub
    mRefitWaited = False
    mRefitRunning = True
    RefitActiveScreen
    mRefitRunning = False
    modUiLock.Leave
    On Error GoTo 0
End Sub

' RefitActiveScreen - 今前面にある画面だけを冪等に組み直す。
'   分岐は modApp.OnRefreshUI(🔄再描画)と同じ形にする(2つ持つとズレる)。
'   チャットは会話が伸びる設計で、幾何は modUI.Repaint が持つ。
'   R21-S1: 描画末尾のワンショット再描画(modViewport2.ReflowIfMoved)からも
'   ここへ落とすため Public 化した(組み直しの分岐を2本持たない)。
'   R21-S6: 本棚の一覧表モードは RefreshCurrent→RenderShelf に落ちるが
'   RenderShelf は帯を触らない(=窓を広げても右端が前の窓のまま)。
'   列幅とクロムまでやり直す軽量経路(modViewport2.RefitShelfTable)へ回す。
Public Sub RefitActiveScreen()
    Dim nm As String
    On Error Resume Next
    nm = ThisWorkbook.ActiveSheet.Name
    If LenB(nm) = 0 Then Exit Sub
    Select Case nm
        Case "Nexus":                  modUI.Repaint
        Case "Dashboard":              modDash.ShowDashboard
        Case modAppDef.SH_HOME:        modHub.EnsureHubLayout
        Case modAppDef.SH_SHELF
            If modKnowledge.IsTableMode() Then
                modViewport2.RefitShelfTable ThisWorkbook.ActiveSheet
            Else
                modKnowledge.RefreshCurrent
            End If
        Case Else:                     Exit Sub   ' 素のシート(config等)は触らない
    End Select
    If Err.Number <> 0 Then
        modLog.LogError "E0801", "modViewport.ViewportRefitTick", Err.Description, Err.Number
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' ColLetter - 列番号(1始まり)を列名へ。1→"A" / 26→"Z" / 27→"AA"。
' ----------------------------------------------------------------------------
'   Excelに聞けば(Cells(1,n).Address)取れるが、範囲文字列の組み立てのたびに
'   COMを往復させたくないうえ、ここだけは実行テストで固定できる純粋な算数
'   なので自前で持つ。範囲外(0以下)は "A" に丸める(壊れた値でRangeを
'   落とさない=呼び出し側の防御を1つ減らす)。
Public Function ColLetter(ByVal colIdx As Long) As String
    Dim n As Long: n = colIdx
    If n < 1 Then n = 1
    Dim s As String
    Do While n > 0
        Dim r As Long: r = (n - 1) Mod 26
        s = Chr$(65 + r) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function
