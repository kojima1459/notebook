Attribute VB_Name = "modViewport"
Option Explicit

' ============================================================================
' modViewport - 画面の「見える範囲」に幾何を合わせる共通部品
'               (2026-08-05 R18-3b → 2026-08-06 R19-1a/1d で全面改訂)。
' ----------------------------------------------------------------------------
' ScrollArea の効能(R19-1d・実機第6報①の物理裏取りによる訂正):
'   Worksheet.ScrollArea が保証するのは「セル選択の制限」と「スクロールバー
'   移動の制限」だけで、マウスホイールは仕様上そのまま素通りする
'   (Microsoft Learn の定義文は "Cells outside the scroll area cannot be
'   selected." としか書いていない)。R18 はこれを「ここから先へはスクロール
'   させない宣言」と読んで設計していた ―― 公式の保証範囲を超えた誤読で、
'   実機で効かなかった直接の原因。
'
' では余白は何が消すのか:
'   右の余白はスクロールの問題ではなく【寸法】の問題。列幅の合計が
'   ウィンドウの可視幅より狭ければ、その差は必ず白く残る(ScrollArea では
'   原理的に一切消えない)。消せるのは列幅を可視幅に合わせることだけで、
'   それを行うのが FitBandToViewport。下の余白も同じで、塗り/行高の下端を
'   内容の実下端(+1画面)まで縮めるのが正攻法(BoundAddr)。
'   ScrollArea は補助(スクロールバーとキー移動は確実に止まる)として残す。
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

' 縦スクロールバーが食う幅(pt)。ここを差し引かないと最終列が半分隠れる。
Private Const SCROLLBAR_W As Double = 12
' 吸収列の最小幅(pt)。0にすると列が消えて「最終列の右は灰色」になる。
Private Const MIN_PAD_PT As Double = 8
' 内容下端に足す余白(pt)。すぐ下で切ると窮屈に見える。
Private Const BOTTOM_PAD As Double = 24
' 自分のブックが前面でないときに ViewportHeight が返す既定値(pt)。R19H FB-5。
Private Const DEFAULT_VIEW_H As Double = 600

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

    Dim target As Double
    target = modUIMain.ViewportWidth() - SCROLLBAR_W
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
Public Function BoundAddr(ByVal ws As Worksheet, ByVal padColLetter As String, _
                          ByVal contentBottom As Double, ByVal maxRow As Long) As String
    If ws Is Nothing Then Exit Function
    BoundAddr = "A1:" & padColLetter & _
                RowAt(ws, BoundBottomY(contentBottom, ViewportHeight(), BOTTOM_PAD), maxRow)
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
' ResetRowsBelow - 指定行より下に残った「行高カスタム」を既定へ戻す(R20-1d)。
' ----------------------------------------------------------------------------
' 実機第7報⑦の層2。Hub は Rows("1:60")、Dash は Rows("1:120")、本棚3モードは
' Rows("7:412") へ毎回いっせいに行高を代入していた。行高を明示した行は
' Excel から見れば「使用済み」なので、内容が20行しか無い画面でも常に
' 60〜412行ぶん(最大6,200pt=8画面)の下スクロール域が残る ―― 塗りと
' ScrollArea をいくら実下端まで縮めても、この焼き付きが消えない限り
' ホイールは下まで転がる(ScrollArea はホイールを止めない・冒頭参照)。
'   UseStandardHeight=True は「行高の明示を取り消す」操作で、値の代入とは
'   逆に使用済みフラグを増やさない。既存ブックに既に焼き付いている行も
'   ここで毎回畳む(冪等。だから救済のために毎回呼んでよい)。
Public Sub ResetRowsBelow(ByVal ws As Worksheet, ByVal fromRow As Long, ByVal toRow As Long)
    If ws Is Nothing Then Exit Sub
    Dim r0 As Long: r0 = fromRow
    If r0 < 1 Then r0 = 1
    If toRow < r0 Then Exit Sub
    On Error Resume Next
    ws.Rows(r0 & ":" & toRow).UseStandardHeight = True
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
' 各画面の EnsureLayout から1行呼ぶだけで、次回の実機報告で境界値を校正
' できるようにする。毎回書くとログが埋まるので画面ごとに1回だけ。
Public Sub LogViewport(ByVal screenName As String)
    On Error Resume Next
    If InStr(mLoggedScreens, "|" & screenName & "|") > 0 Then Exit Sub
    mLoggedScreens = mLoggedScreens & "|" & screenName & "|"
    modLog.LogUsage "viewport", screenName, _
        CLng(modUIMain.ViewportWidth()) & "x" & CLng(ViewportHeight())
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
    mRefitWaited = False
    mRefitRunning = True
    RefitActiveScreen
    mRefitRunning = False
    On Error GoTo 0
End Sub

' RefitActiveScreen - 今前面にある画面だけを冪等に組み直す。
'   分岐は modApp.OnRefreshUI(🔄再描画)と同じ形にする(2つ持つとズレる)。
'   チャットは会話が伸びる設計で、幾何は modUI.Repaint が持つ。
Private Sub RefitActiveScreen()
    Dim nm As String
    On Error Resume Next
    nm = ThisWorkbook.ActiveSheet.Name
    If LenB(nm) = 0 Then Exit Sub
    Select Case nm
        Case "Nexus":                  modUI.Repaint
        Case "Dashboard":              modDash.ShowDashboard
        Case modAppDef.SH_HOME:        modHub.EnsureHubLayout
        Case modAppDef.SH_SHELF:       modKnowledge.RefreshCurrent
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
