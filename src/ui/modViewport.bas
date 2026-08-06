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
'   純関数(PadPtNeeded / BoundBottomY / RightEdgeAt)は他モジュールを呼ばず、
'   LibreOffice の純ロジックテストで固定する。
' ============================================================================

' 縦スクロールバーが食う幅(pt)。ここを差し引かないと最終列が半分隠れる。
Private Const SCROLLBAR_W As Double = 12
' 吸収列の最小幅(pt)。0にすると列が消えて「最終列の右は灰色」になる。
Private Const MIN_PAD_PT As Double = 8
' 内容下端に足す余白(pt)。すぐ下で切ると窮屈に見える。
Private Const BOTTOM_PAD As Double = 24

' usage_log("viewport") を画面ごとに1セッション1回だけ出すためのメモ(R19-1e)。
Private mLoggedScreens As String

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

    Dim units As Double: units = ws.Columns(padColLetter).ColumnWidth
    If units <= 0 Then
        ws.Columns(padColLetter).ColumnWidth = 1
        units = ws.Columns(padColLetter).ColumnWidth
    End If
    Dim ptPerUnit As Double
    If units > 0 Then ptPerUnit = ws.Columns(padColLetter).Width / units
    If ptPerUnit <= 0 Then Exit Sub          ' 換算できない端末では何もしない

    Dim fixedW As Double
    fixedW = ws.Range(bandAddr).Width - ws.Columns(padColLetter).Width

    Dim target As Double
    target = modUIMain.ViewportWidth() - SCROLLBAR_W
    If target < minRightX Then target = minRightX

    ws.Columns(padColLetter).ColumnWidth = PadPtNeeded(target, fixedW, MIN_PAD_PT) / ptPerUnit
    On Error GoTo 0
End Sub

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

' ----------------------------------------------------------------------------
' BoundFor - 実測の右下端(pt)を含む最小の "A1:<列><行>" を組み立てる。
' ----------------------------------------------------------------------------
'   吸収列方式(BoundAddr)に移行できない画面のための旧口。列も実測で決める。
Public Function BoundFor(ByVal ws As Worksheet, ByVal rightX As Double, _
                         ByVal bottomY As Double, ByVal maxCol As Long, _
                         ByVal maxRow As Long) As String
    If ws Is Nothing Then Exit Function
    BoundFor = "A1:" & ColLetter(ColAt(ws, rightX, maxCol)) & RowAt(ws, bottomY, maxRow)
End Function

' ColAt / RowAt - pt座標を含む最小の列/行番号。上限に張り付いたら上限を返す。
'   列幅・行高はフォントやDPIで変わり、ptから列番号を机上計算することは
'   できない(「9文字幅=何pt」は端末依存)。実際のセル幾何(Left/Width,
'   Top/Height)を読んで決める ―― Shape座標をセル幾何から出すのと同じ理由。
Private Function ColAt(ByVal ws As Worksheet, ByVal x As Double, ByVal maxCol As Long) As Long
    If maxCol < 1 Then maxCol = 1
    ColAt = maxCol
    If ws Is Nothing Then Exit Function
    Dim i As Long
    On Error Resume Next
    For i = 1 To maxCol
        If ws.Cells(1, i).Left + ws.Cells(1, i).Width >= x Then
            ColAt = i
            Exit For
        End If
    Next i
    On Error GoTo 0
End Function

Private Function RowAt(ByVal ws As Worksheet, ByVal y As Double, ByVal maxRow As Long) As Long
    If maxRow < 1 Then maxRow = 1
    RowAt = maxRow
    If ws Is Nothing Then Exit Function
    Dim i As Long
    On Error Resume Next
    For i = 1 To maxRow
        If ws.Cells(i, 1).Top + ws.Cells(i, 1).Height >= y Then
            RowAt = i
            Exit For
        End If
    Next i
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' ViewportHeight - ウィンドウの実可視高(pt)。ViewportWidth(modUIMain)の縦版。
' ----------------------------------------------------------------------------
'   異常値でレイアウト計算全体を壊さないよう 200〜1200pt へクランプする。
Public Function ViewportHeight() As Double
    Dim h As Double
    On Error Resume Next
    If ActiveWorkbook Is ThisWorkbook Then h = ActiveWindow.UsableHeight
    If h <= 0 Then h = ActiveWindow.VisibleRange.Height
    On Error GoTo 0
    If h < 200 Then h = 200
    If h > 1200 Then h = 1200
    ViewportHeight = h
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
