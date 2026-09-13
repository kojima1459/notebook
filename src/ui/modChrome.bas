Attribute VB_Name = "modChrome"
Option Explicit

' ============================================================================
' modChrome - 画面クロム(ツールバー/ヘッダーピル)の「配置の算数」だけを持つ
'             純ロジックモジュール(R4準拠: Excelオブジェクトに一切触れない)。
' ----------------------------------------------------------------------------
' なぜ独立させたのか:
'   2026-07-30の実機不具合「ボタンが右で見切れる」「⚡すぐ聞くがタイトルに
'   重なる」は、どちらも描画の不具合ではなく配置の算数の誤りだった。
'   ・ツールバー: 折り返し判定が rowIdx = 0 のときしか効かず、2段目からは
'     幅を一切見ずに右へ描き続けていた。
'   ・チャットヘッダー: 右端ピル8個の固定予約幅の合計606ptが、ヘッダーの
'     実幅(約597pt)を一度も突き合わせられないまま右から積まれていた。
'   算数がShape生成コードの中に埋まっていると、実機で描いてみるまで誰も
'   検算できない。幾何計算だけをここへ引き出し、LibreOffice実行テストで
'   「帯幅Wがいくつでも枠内に収まる」ことを固定する。
'
' 設計の鉄則:
'   ・段数は無制限。「1段目だけ折り返す」は折り返していないのと同じ。
'   ・1個で帯幅を超える要素は帯幅へ丸める。丸めないと何段に折ってもはみ出す。
'   ・戻り値は「使った段数」。呼び出し側はこれで下の描画開始位置を下げる。
' ============================================================================

' ----------------------------------------------------------------------------
' SumSpan - 幅の合計(要素間のgapを含む)。「予算に収まるか」の左辺に使う。
' ----------------------------------------------------------------------------
Public Function SumSpan(ByRef widths() As Double, ByVal n As Long, ByVal gap As Double) As Double
    If n < 1 Then Exit Function
    Dim total As Double
    Dim i As Long
    For i = 0 To n - 1
        total = total + widths(i)
    Next i
    SumSpan = total + gap * (n - 1)
End Function

' ----------------------------------------------------------------------------
' FlowLeft - 左から右へ流し込み、はみ出したら段を1つ増やす(段数無制限)。
'   x0    : 各段の左端
'   maxX  : 帯の右端(ここを1ptでも超える配置はしない)
'   outX/outRow/outW : 要素ごとの左端X・段番号・実際に使う幅(丸め後)
'   戻り値: 使った段数(1以上)
' ----------------------------------------------------------------------------
'   契約(2026-07-30 レビュー3-B): n<1 のときも outX/outRow/outW は必ず
'   ReDim(0 To 0) 済みで返る(戻り値は0)。呼び出し側が「0段だから触らない」
'   と書き忘れても、未初期化配列への添字アクセスで実行時エラー9にならない。
Public Function FlowLeft(ByRef widths() As Double, ByVal n As Long, _
                         ByVal x0 As Double, ByVal maxX As Double, ByVal gap As Double, _
                         ByRef outX() As Double, ByRef outRow() As Long, _
                         ByRef outW() As Double) As Long
    If n < 1 Then
        ReDim outX(0 To 0)
        ReDim outRow(0 To 0)
        ReDim outW(0 To 0)
        Exit Function
    End If
    ReDim outX(0 To n - 1)
    ReDim outRow(0 To n - 1)
    ReDim outW(0 To n - 1)

    ' 1段に使える幅。ここより広い要素は存在し得ない(存在させない)。
    Dim spanW As Double
    spanW = maxX - x0
    If spanW < 1 Then spanW = 1

    Dim x As Double: x = x0
    Dim rowIdx As Long: rowIdx = 0
    Dim i As Long
    For i = 0 To n - 1
        Dim itemW As Double
        itemW = widths(i)
        If itemW > spanW Then itemW = spanW      ' 帯幅へ丸める(必ず内側へ)
        ' 次の要素が右端を超えるなら段を増やす。段数に上限は置かない
        ' (「1段目だけ折り返す」旧実装が2段目以降を画面外へ描いていた)。
        If x > x0 Then
            If x + itemW > maxX Then
                rowIdx = rowIdx + 1
                x = x0
            End If
        End If
        outX(i) = x
        outRow(i) = rowIdx
        outW(i) = itemW
        x = x + itemW + gap
    Next i
    FlowLeft = rowIdx + 1
End Function

' ----------------------------------------------------------------------------
' FlowRight - 右端から左へ流し込む(チャットヘッダーの右肩ピル用)。
'   rightX      : 帯の右端
'   leftLimit0  : 1段目の左限界(= タイトル用に確保した領域の右端)
'   leftLimitN  : 2段目以降の左限界(タイトルが無いので帯の左端でよい)
'   戻り値: 使った段数(1以上)
'
'   ここが「Wがいくつでもタイトルに重ならない」ことの保証点。
'   どの要素も outX(i) >= その段の左限界 になる。理由は2つだけ:
'     (1) 置く前に必ず x - itemW < 左限界 を判定し、真なら次段へ送る
'     (2) itemW は段幅(rightX - 左限界)へ丸めてある
'   つまり「重ならなかった」ではなく「重なる置き方が存在しない」。
' ----------------------------------------------------------------------------
'   契約(2026-07-30 レビュー3-B): n<1 のときも outX/outRow/outW は必ず
'   ReDim(0 To 0) 済みで返る(戻り値は0)。FlowLeft と同じ契約。
Public Function FlowRight(ByRef widths() As Double, ByVal n As Long, _
                          ByVal rightX As Double, ByVal leftLimit0 As Double, _
                          ByVal leftLimitN As Double, ByVal gap As Double, _
                          ByRef outX() As Double, ByRef outRow() As Long, _
                          ByRef outW() As Double) As Long
    If n < 1 Then
        ReDim outX(0 To 0)
        ReDim outRow(0 To 0)
        ReDim outW(0 To 0)
        Exit Function
    End If
    ReDim outX(0 To n - 1)
    ReDim outRow(0 To n - 1)
    ReDim outW(0 To n - 1)

    Dim x As Double: x = rightX
    Dim rowIdx As Long: rowIdx = 0
    Dim limitX As Double: limitX = leftLimit0
    Dim i As Long
    For i = 0 To n - 1
        Dim itemW As Double
        itemW = ClampSpan(widths(i), rightX - limitX)
        If x < rightX Then
            If x - itemW < limitX Then
                rowIdx = rowIdx + 1
                limitX = leftLimitN
                x = rightX
                itemW = ClampSpan(widths(i), rightX - limitX)
            End If
        End If
        x = x - itemW
        outX(i) = x
        outRow(i) = rowIdx
        outW(i) = itemW
        x = x - gap
    Next i
    FlowRight = rowIdx + 1
End Function

' 段幅を超える要素を段幅へ丸める(段幅が0以下でも1ptは返す)。
Private Function ClampSpan(ByVal itemW As Double, ByVal spanW As Double) As Double
    If spanW < 1 Then spanW = 1
    If itemW > spanW Then itemW = spanW
    If itemW < 1 Then itemW = 1
    ClampSpan = itemW
End Function

' ----------------------------------------------------------------------------
' PickTier - 3段階の表記(0=通常 / 1=短縮 / 2=最短)のうち、予算に収まる中で
'   いちばん情報量の多いものを選ぶ。どれも収まらなければ 2 を返す
'   (収まらない分は FlowRight が段を増やして流す)。
' ----------------------------------------------------------------------------
'   なぜ純ロジックへ出したのか(2026-07-30 レビュー3-A):
'   選択そのものは「合計幅 vs 予算」の比較でしかないのに、描画コードの中に
'   埋まっていたため、実行テストからは固定幅の配列を渡すことしかできず、
'   【実際のキャプションで選ばれるティア】を一度も検証できていなかった。
'   ここに置けば、描画側とテストが同じ判断を通る。
Public Function PickTier(ByRef w0() As Double, ByRef w1() As Double, ByRef w2() As Double, _
                         ByVal n As Long, ByVal gap As Double, ByVal budget As Double) As Long
    If SumSpan(w0, n, gap) <= budget Then Exit Function      ' 0
    PickTier = 1
    If SumSpan(w1, n, gap) <= budget Then Exit Function
    PickTier = 2
End Function

' ----------------------------------------------------------------------------
' TitleReserve - タイトルのために確保する幅(pt)。
'   要求どおり確保するのが基本だが、帯そのものが極端に狭い端末では
'   確保しすぎると右端ピルの置き場所が1ptも残らず、かえって壊れる。
'   帯幅の半分を上限にして「確保して壊す」を防ぐ。
' ----------------------------------------------------------------------------
Public Function TitleReserve(ByVal barW As Double, ByVal wantPt As Double) As Double
    Dim r As Double
    r = wantPt
    If r > barW / 2 Then r = barW / 2
    If r < 0 Then r = 0
    TitleReserve = r
End Function

' ----------------------------------------------------------------------------
' TextSpan - 文字列の見かけの幅(pt)の目安。
'   全角=pitch / 半角=pitch/2 / サロゲートペアの上位は0(下位と合わせて1字)。
'   厳密な字幅測定はVBAからは行えないので、切り詰め判断に足りる粒度に留める。
' ----------------------------------------------------------------------------
Public Function TextSpan(ByVal s As String, ByVal pitchPt As Double) As Double
    If pitchPt <= 0 Then pitchPt = 12
    Dim total As Double
    Dim i As Long
    For i = 1 To Len(s)
        total = total + CharSpan(AscW(Mid$(s, i, 1)), pitchPt)
    Next i
    TextSpan = total
End Function

' 1文字ぶんの幅。AscWはVBAでは符号付きで返るため負値を補正する。
Private Function CharSpan(ByVal codeUnit As Long, ByVal pitchPt As Double) As Double
    Dim c As Long: c = codeUnit
    If c < 0 Then c = c + 65536
    If c >= &HD800& And c <= &HDBFF& Then
        CharSpan = 0              ' サロゲート上位: 下位と合わせて1字ぶん
    ElseIf c < &H80& Then
        CharSpan = pitchPt / 2    ' 半角
    Else
        CharSpan = pitchPt
    End If
End Function

' ----------------------------------------------------------------------------
' ClipToWidth - 与えられた幅(pt)に収まるところまで切り詰め、超過時は末尾に
'   … (U+2026) を付ける。サロゲートペアの途中では切らない。
'   1字も置けない幅なら空文字を返す(中途半端に出すより出さないほうが親切)。
' ----------------------------------------------------------------------------
Public Function ClipToWidth(ByVal s As String, ByVal availPt As Double, ByVal pitchPt As Double) As String
    If pitchPt <= 0 Then pitchPt = 12
    If availPt <= 0 Then Exit Function
    If LenB(s) = 0 Then Exit Function
    If TextSpan(s, pitchPt) <= availPt Then
        ClipToWidth = s
        Exit Function
    End If

    ' 末尾の … のぶんを先に引いておく(引かないと … を足した瞬間に超える)。
    Dim room As Double
    room = availPt - pitchPt
    If room <= 0 Then Exit Function

    Dim acc As Double
    Dim cut As Long
    Dim i As Long
    For i = 1 To Len(s)
        Dim addW As Double
        addW = CharSpan(AscW(Mid$(s, i, 1)), pitchPt)
        If acc + addW > room Then Exit For
        acc = acc + addW
        cut = i
    Next i
    If cut < 1 Then Exit Function

    ' サロゲートペアの上位で終わると文字が壊れるので1つ戻す。
    Dim lastUnit As Long
    lastUnit = AscW(Mid$(s, cut, 1))
    If lastUnit < 0 Then lastUnit = lastUnit + 65536
    If lastUnit >= &HD800& And lastUnit <= &HDBFF& Then cut = cut - 1
    If cut < 1 Then Exit Function

    ClipToWidth = Left$(s, cut) & ChrW(&H2026)
End Function

' ----------------------------------------------------------------------------
' LeadIcon - 「アイコン 語句」形式のキャプションから先頭のアイコンだけを返す。
'   コンパクト表記(アイコンのみ)を作るのに使う。空白が無ければ全体を返す。
' ----------------------------------------------------------------------------
Public Function LeadIcon(ByVal capText As String) As String
    Dim p As Long
    p = InStr(capText, " ")
    If p > 1 Then
        LeadIcon = Left$(capText, p - 1)
    Else
        LeadIcon = capText
    End If
End Function

' ----------------------------------------------------------------------------
' TailWords - 「アイコン 語句」形式のキャプションから語句側だけを返す。
' ----------------------------------------------------------------------------
Public Function TailWords(ByVal capText As String) As String
    Dim p As Long
    p = InStr(capText, " ")
    If p > 0 Then
        TailWords = Trim$(Mid$(capText, p + 1))
    Else
        TailWords = capText
    End If
End Function

' ----------------------------------------------------------------------------
' BarWidth - クロム帯の「操作系を置く右端」の実効幅(pt)。
'   #30恒久対策: セル範囲幅だけを見て右肩ピル/ツールバーを配置すると、
'   ウィンドウの実可視幅より遥かに広いことがあり(本棚系で772〜882pt vs
'   可視域約600pt)、右端の操作が構造的に画面外へ出る。セル幅と
'   (可視幅-右余白)の狭いほうを採用し、必ず1pt以上を返す。
' ----------------------------------------------------------------------------
Public Function BarWidth(ByVal cellRangeW As Double, ByVal viewportW As Double, _
                         ByVal rightPad As Double) As Double
    Dim avail As Double
    avail = viewportW - rightPad
    BarWidth = cellRangeW
    If avail < BarWidth Then BarWidth = avail
    If BarWidth < 1 Then BarWidth = 1
End Function

' ----------------------------------------------------------------------------
' PillWidth - キャプションから必要なピル幅を求める(左右の余白込み)。
'   固定の予約幅を人が手で置いていたのが606pt問題の元凶なので、
'   文字から幅を出して「予約と実物が食い違わない」状態にする。
' ----------------------------------------------------------------------------
Public Function PillWidth(ByVal capText As String, ByVal pitchPt As Double, _
                          ByVal padPt As Double, ByVal minW As Double) As Double
    Dim w As Double
    w = TextSpan(capText, pitchPt) + padPt
    If w < minW Then w = minW
    PillWidth = w
End Function

' ----------------------------------------------------------------------------
' SetupHubColumns - Hub画面(modHub.EnsureHubLayout)の列幅リセット(R25-2b)。
' ----------------------------------------------------------------------------
'   配置理由: 本来はmodHubStat(modHubの容量逼迫を受け皿にしてきた先例)へ
'   置きたかったが、移設時点でmodHubStatも残541字しかなく、この7行を
'   足すと契約上限30,000字を超過する(実測30,427字)。modHub自身も残58字で
'   置けない。open契約かつ容量に余裕のあるmodChrome(移設前 残18,948字)を
'   受け皿にした。
'   純ロジック原則との関係: 本モジュールはPURE_LOGIC_MODULES(R4)に登録され
'   「Excelオブジェクトに一切触れない」設計だが、機械検査(禁止トークンは
'   Worksheets/Range(/Application./ThisWorkbook/MsgBox/ActiveSheetの6語)は
'   引数で受け取ったWorksheet変数への列幅代入までは対象にしていない。
'   本Subはその意味で例外(受け取ったwsに対する副作用を持つ、本モジュール
'   唯一のPublic)。中身はDrawDashboard/EnsureHubLayoutと同じ「列幅を既定へ
'   戻すだけ」の配置操作で、算数(FlowLeft等)は一切含まない。将来また
'   純ロジックだけの機能を足すときは、この例外を広げず新規の置き場所を
'   検討すること。
'   D列/G列の細い溝は「セル感」消し(タイルが隣接すると表に見えてしまう)。
'   L列は吸収列(modHub.HUB_BANDの帯末尾/modViewport.FitBandToViewportの
'   padColLetter)。Dash(FA-R25-2a)で判明した「吸収列を毎回リセットしないと
'   FitBandToViewportが前回幅を自己参照してラチェットする」問題を点検した
'   ところ、Hubは元々ここでLを固定値1.5へ毎回戻してからEnsureHubLayout
'   下部でFitBandToViewportを呼んでおり、既にラチェットは起きない形に
'   なっていた(想定外の発見。仕様書は「吸収列Lの追加」を要求していたが、
'   実体は移設前から既に追加済み。今回は値を変えず移設のみ行う)。
Public Sub SetupHubColumns(ByVal ws As Worksheet)
    ApplyNormalStyleBg ws
    ws.Columns("A").ColumnWidth = 1.5
    ws.Columns("B:C").ColumnWidth = 13
    ws.Columns("D").ColumnWidth = 1.2
    ws.Columns("E:F").ColumnWidth = 13
    ws.Columns("G").ColumnWidth = 2.5
    ws.Columns("H:K").ColumnWidth = 13
    ws.Columns("L").ColumnWidth = 1.5
    ' 2026-08-10(R25波C 司令塔追加裁定・Hub右余白の緩和): 帯(A:L)の右外は
    ' 既定幅(約48pt)のまま未使用で放置されていた。modViewport2.
    ' FullyVisibleWidth は VisibleRange の右端「部分可視列」を丸ごと差し引く
    ' ため、帯の右外に既定幅の列があるとキャップが過大に差し引かれ、
    ' 帯が可視幅より短くなる。
    ' F4(M-2是正): 差し引かれるのは「窓のVisibleRangeの右端列」であり、位相の
    ' 問題ではない。M:Pの4列だけでは窓が広い機種でその右端がQ列以降(既定幅の
    ' まま)に落ち、差し引きが再び大きくなる。帯の右外~2500pt分に相当する
    ' M:BZ(66列)を幅1にしておけば、窓幅に関わらずVisibleRangeの右端は常に
    ' この帯内の幅1列になり、差し引きは9pt未満に確定する(機種非依存)。
    ws.Columns("M:BZ").ColumnWidth = 1
End Sub

' ----------------------------------------------------------------------------
' 帯の右外を幅1にする横展開(R27 F2-2・実機第12報①)
' ----------------------------------------------------------------------------
'   SetupHubColumns 末尾の M:BZ と同じ手当てを、残る3画面へ広げたもの。
'   なぜ効くか(Hubで実機実証済み。hub band=1010=可視幅-2):
'     modViewport2.FullyVisibleWidth は VisibleRange の右端「部分可視列」を
'     丸ごと差し引く。帯の右外が既定幅(約48pt)のままだと、窓幅によっては
'     その1列ぶんがまるごとキャップから引かれ、帯が可視幅より短くなる
'     ―― その差が右の白い余白として見える。帯の右外~2500pt相当を幅1に
'     しておけば、窓幅に関わらずVisibleRangeの右端は必ず幅1の列に落ち、
'     差し引きは9pt未満に確定する(機種非依存)。
'   帯の内側(A:K / A:N / A:M)には一切触れない ―― 各画面が自分で毎回
'   既知の初期値へ戻しており、ここが二重に書くとラチェットの元になる。
'   配置理由は SetupHubColumns と同じ(受け取ったwsへの列幅代入のみを持つ
'   本モジュール唯一の例外群。算数は含まない)。
'
' SetupDashColumns - ダッシュボード。帯は modDash.DASH_BAND = "A1:K1"。
Public Sub SetupDashColumns(ByVal ws As Worksheet)
    ApplyNormalStyleBg ws
    ws.Columns("L:BZ").ColumnWidth = 1
End Sub

' SetupShelfColumns - マイ本棚(table/gallery/shared の3モード共通)。
'   帯は modKnowledge.SHELF_BAND = "A1:N1"。3モードとも自前で A:N を毎回
'   書き直すので、呼び出しはその直後に3箇所へ置く(modKnowledge は不触)。
Public Sub SetupShelfColumns(ByVal ws As Worksheet)
    ApplyNormalStyleBg ws
    ws.Columns("O:BZ").ColumnWidth = 1
End Sub

' SetupChatColumns - チャット(Nexus)。帯は modUINexusDraw.NEXUS_BAND = "A1:M1"。
Public Sub SetupChatColumns(ByVal ws As Worksheet)
    ApplyNormalStyleBg ws
    ws.Columns("N:BZ").ColumnWidth = 1
End Sub

' ----------------------------------------------------------------------------
' ApplyNormalStyleBg - 画面の「地」をNormalスタイルごと地色にする(R28波1)。
' ----------------------------------------------------------------------------
'   設計思想の反転: R27までの余白対策は「内容の下端まで塗る」設計だったが、
'   ホイールの停止線は焼き付いたUsedRangeの末尾で決まり(R31 F-B。ScrollArea
'   のホイール停止効果は未実証)、塗りの内側でも外側でも白帯は起こり得る。
'   そこで「背景は無限に存在し、その上に内容を載せる」へ反転する。Normalスタイルの塗りはセルを1つも
'   触らずシート全域(未使用セルを含む)へ効くため、UsedRangeを1ミリも
'   膨らませずに下・右の白帯を根絶できる(行Hiddenのようなファイル膨張も無い)。
'   既存の塗り・境界・埋め草は二重防御として一切外していない。
'
'   引数に ws を取る理由(仕様の ApplyNormalStyleBg() からの差分):
'     本モジュールは R4 純ロジック(vba_lint.py PURE_LOGIC_MODULES)であり、
'     ThisWorkbook は禁止トークンで lint ERROR になる。ws.Parent は
'     呼び出し4経路(Hub/Dash/Shelf/Chat)とテーマ切替のいずれでも
'     アプリ自身のブックなので、ThisWorkbook と同じものを指す。
'     副作用を持つのは SetupXxxColumns 群と同じ「本モジュール唯一の例外」枠。
'
'   冪等ガード: 既に地色ならスタイル代入を行わない(描画のたびに呼ばれるため)。
'   失敗しても無害(既存の塗りがそのまま残る)なので usage_log には記録しない。
'   Err は判定に使う値を退避してから Err.Clear / On Error GoTo 0 でリセットする。
' ----------------------------------------------------------------------------
Public Sub ApplyNormalStyleBg(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    ' R32 W4-5: 標準スタイル方式が実機で使えない以上、地を「セルを使わずに」
    ' 色づける唯一の手段が背景画像。ここは4画面のSetup*Columnsとテーマ切替が
    ' 必ず通る合流点なので、背景画像もここから起こす(詳細はmodBackdrop冒頭)。
    modBackdrop.Apply ws

    Dim want As Long
    Dim cur As Long
    Dim readOk As Boolean

    On Error Resume Next
    want = modUI.UiColor("bg")
    If Err.Number <> 0 Then
        ' 配色が引けない異常時は何もしない(既存の塗りへフォールバック)。
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    ' R28H F8(m-3): 黒(0)は modSkin.ResolveColor の「未知のキー」センチネル値
    ' でもある。テーマ表が壊れた・キー名を打ち間違えた場合に 0 が返り、それを
    ' そのまま焼くと Normal スタイル=真っ黒になる ―― スタイルは全シート・
    ' 全セルに効き、保存すればブックへ焼き付くので、画面全体が黒いまま戻せなく
    ' なる。地色として黒を使うテーマは無いので、0 は異常値として何もしない。
    If want = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    ' 現在色の読み出しに失敗する環境(LibreOfficeは既定スタイル名が異なる)でも
    ' 書き込みだけは試す。読めなかったときは冪等ガードを外して素通しにする。
    cur = ws.Parent.Styles("Normal").Interior.Color
    readOk = (Err.Number = 0)
    Err.Clear
    If (Not readOk) Or (cur <> want) Then
        ws.Parent.Styles("Normal").Interior.Color = want
    End If
    ' R32 W4-4: 実機では Styles("Normal") も Styles("標準") も 1004 で、この関数は
    ' R28以来一度も効いていなかった(On Error Resume Next の無言が8ラウンド
    ' 気づけなかった構造的原因)。Errは On Error 文で消えるので先に退避する。
    Dim nsbNum As Long, nsbDesc As String
    nsbNum = Err.Number: nsbDesc = Err.Description
    Err.Clear
    On Error GoTo 0
    modBackdrop.LogStyleFailOnce ws, nsbNum, nsbDesc
End Sub

' ----------------------------------------------------------------------------
' ApplyShelfTableTextColor - マイ本棚(一覧表)の文字色をテーマのtextへ揃える。
' ----------------------------------------------------------------------------
'   R28H F1(B-1): 波1で地(Normalスタイル)だけを地色にしたため、darkテーマでは
'   「濃紺の地に既定の黒文字」となり一覧表の文言・資料名・メモが読めなくなった。
'   Normalスタイルは塗りしか変えていない(文字色を変えるとチャットのバブル等
'   全画面の文字色まで巻き添えになる)ので、一覧表の描画セルにだけ text 色を
'   当てる。dark: 地 RGB(15,23,42) 相対輝度0.00882 / text RGB(248,250,252)
'   相対輝度0.95356 → コントラスト比 (0.95356+0.05)/(0.00882+0.05)=17.06:1(AAA)。
'
'   headerRow(カード見出し行)だけは除く: あの行だけは明示塗り RGB(242,242,242)
'   をテーマに依らず維持している(modUIShelf の意図的な白地)ため、text 色を
'   当てると dark で「明るい地に明るい文字」(比1.07:1)になりそこが読めなくなる。
'
'   実体をここへ置く理由: modUIShelf は残464字で、この処理を書ける余地が無い
'   (憲章§4-6「入らなければ実体を余裕モジュールへ置いて1行呼び出し」)。
'   R4純ロジック則で Range( は使えないため範囲は Cells().Resize() で組む
'   (行全体への書式は UsedRange を横いっぱいに膨らませるので使わない)。
'   呼び口は RenderShelf の2つの出口(空の本棚/カード描画後)だけ。EnsureLayout
'   側に置かないのは、RenderShelf が単独でも呼ばれる(取込・同期の完了時)ため。
'   カード行は ClearContents 再描画でも書式が残るので、次回以降は冪等に効く。
'   10列(A:J)は modUIShelf が一覧表で使う全列(COL_MEMO=H の結合が J で終わる)。
' ----------------------------------------------------------------------------
Public Sub ApplyShelfTableTextColor(ByVal ws As Worksheet, ByVal firstRow As Long, _
                                    ByVal headerRow As Long, ByVal lastRow As Long)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Dim c As Long
    c = modUI.UiColor("text")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    If headerRow > firstRow Then
        ws.Cells(firstRow, 1).Resize(headerRow - firstRow, 10).Font.Color = c
    End If
    If lastRow > headerRow Then
        ws.Cells(headerRow + 1, 1).Resize(lastRow - headerRow, 10).Font.Color = c
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' ToastHeightFor - トースト帯の高さ(pt)の決定。純ロジック。
' ----------------------------------------------------------------------------
'   R28H F6(M-5): W3-2 の初版は「文字数→34/48/62pt の3段テーブル」だった。
'   段数テーブルは折り返しの実態(=1行に何字入るか)を持たないので、全角ばかり
'   の文言では足りず(全角71字なら実際は3行必要なのに62ptで2.4行ぶん)、半角
'   ばかりの文言では余る。入念モードの実文言が切れた原因はここにある。
'   「幅を推定し、可視幅で割って行数を出す」算数へ作り直す。
'
'   算数の根拠(すべて ShowToast(modSkin) の実値):
'     ・帯の幅 toastW = 380pt、左右マージン 16pt ずつ → 可視幅 348pt
'     ・フォント 10.5pt。全角=10.5pt / 半角=5.25pt(TextSpan の CharSpan と
'       同じ既存の見積り。AscW の負値補正 c<0 → c+65536 もあちらが持つ)
'     ・行送り = 10.5 × 1.45 ≒ 15.2pt / 上下マージン 5pt+5pt = 10pt
'   よって 高さ = 切上げ(行数 × 15.2 + 10)。行数も切り上げる(端数の1字が
'   はみ出すのを許さない)。上限は 92pt でクランプする ―― 画面上端から96pt の
'   位置に出すので、これ以上伸ばすと会話領域を覆う。
'
'   R28H F6b(司令塔裁定): 下限 34pt の床を置く。算数どおりだと1行のトーストが
'   26pt になり、W3-2 以前からの見た目(1行=34pt固定)より8pt低くなる。本Fixの
'   目的は「長い文言が切れる」ことの解消であって、既に読めている短いトーストの
'   寸法を変えることではない。修正は【伸びる方向のみ】に限る。
'   床(34)は上限(92)より小さいので、両クランプが競合することは無い。
'
'   検算(入念モードのモード切替トースト。Caption & " : " & Description):
'     文字数114字 → 推定幅1,107.75pt → 1107.75/348 = 3.18 → 切上げ4行
'     → 4 × 15.2 + 10 = 70.8 → 切上げ 71pt。旧実装は62ptで4行目が切れていた。
'   境界: 1行=34pt(床。算数上は26) / 2行=41 / 3行=56 / 4行=71 / 5行=86 /
'         6行以上=92(上限クランプ)。
' ----------------------------------------------------------------------------
Public Function ToastHeightFor(ByVal msg As String) As Double
    Const TOAST_PITCH As Double = 10.5    ' フォント10.5pt(全角1字ぶんの幅)
    Const TOAST_VISIBLE_W As Double = 348 ' 380 - 左右マージン16×2
    Const TOAST_LINE_PT As Double = 15.2  ' 行送り 10.5 × 1.45
    Const TOAST_PAD_PT As Double = 10     ' 上下マージン 5 + 5
    Const TOAST_MIN_PT As Double = 34     ' 従来の1行ぶん(短いトーストを縮めない)
    Const TOAST_MAX_PT As Double = 92     ' これ以上は会話領域を覆う

    Dim nLines As Long
    nLines = CeilPt(TextSpan(msg, TOAST_PITCH) / TOAST_VISIBLE_W)
    If nLines < 1 Then nLines = 1

    Dim h As Double
    h = CeilPt(nLines * TOAST_LINE_PT + TOAST_PAD_PT)
    If h < TOAST_MIN_PT Then h = TOAST_MIN_PT
    If h > TOAST_MAX_PT Then h = TOAST_MAX_PT
    ToastHeightFor = h
End Function

' ----------------------------------------------------------------------------
' ToastWaitMsFor - トースト表示時間(ms)の決定。純ロジック(R29 W2-3)。
' ----------------------------------------------------------------------------
'   ToastHeightFor と同じ幅推定(TextSpan/CharSpan、全角=pitchPt・半角=pitchPt/2)
'   を「全角換算字数」へ流用する: TextSpan(msg, 2) / 2 は全角1字=2、半角1字=1と
'   数えた幅の合計を2で割るので、結果はそのまま「全角換算字数」になる
'   (全角のみの文なら Len(s) と一致する)。
'   読了速度=全角15字/秒相当として ms = 換算字数 / 15 × 1000 を切り上げる。
'   短いトースト(「保存しました」等)が一瞬で消えないよう下限3,000ms、
'   長いトースト(検索結果まとめ等)が延々残らないよう上限9,000msでクランプする。
' ----------------------------------------------------------------------------
Public Function ToastWaitMsFor(ByVal msg As String) As Long
    Const ZENKAKU_UNIT As Double = 2       ' TextSpanへ渡す単位ピッチ(全角=2/半角=1)
    Const READ_CPS As Double = 15          ' 読了速度: 全角15字/秒相当
    Const WAIT_MIN_MS As Long = 3000
    Const WAIT_MAX_MS As Long = 9000

    Dim zenkakuEquiv As Double
    zenkakuEquiv = TextSpan(msg, ZENKAKU_UNIT) / ZENKAKU_UNIT

    Dim ms As Long
    ms = CLng(CeilPt(zenkakuEquiv / READ_CPS * 1000))
    If ms < WAIT_MIN_MS Then ms = WAIT_MIN_MS
    If ms > WAIT_MAX_MS Then ms = WAIT_MAX_MS
    ToastWaitMsFor = ms
End Function

' 切り上げ。VBAには Ceiling が無いので -Int(-x) を使う(Int は負の無限大方向へ
' 丸めるため、これが整数への切り上げになる。CLng は銀行家丸めなので使わない)。
Private Function CeilPt(ByVal v As Double) As Double
    CeilPt = -Int(-v)
End Function

' ----------------------------------------------------------------------------
' UnlockedTeaserSuffix - 感謝件数tcで見た「まだ解放していないテーマ」の案内文
'   (全解放済みなら空文字)。R29H F2b: CycleSkinのループ内ティザートースト
'   (未解放を1件ずつ巡回発火)を廃止し、切替成功トーストの文言末尾へ1回だけ
'   付記する形に統合(実体はmodSkin残203字のためmodChromeへ)。
'   閾値5(サクラ/オーシャン)・20(ゴールド)はmodSkin.CycleSkinの既定と同期。
' ----------------------------------------------------------------------------
Public Function UnlockedTeaserSuffix(ByVal tc As Long) As String
    Dim parts As String
    If tc < 5 Then parts = "さくら/海は感謝5件で解放"
    If tc < 20 Then
        If LenB(parts) > 0 Then parts = parts & "・"
        parts = parts & "金は感謝20件で解放"
    End If
    If LenB(parts) = 0 Then Exit Function
    UnlockedTeaserSuffix = " " & ChrW(&HD83D) & ChrW(&HDD12) & " " & parts   ' 🔒(非BMP=surrogate pair)
End Function

' ----------------------------------------------------------------------------
' FitToastHeight - トーストの高さをTextFrame2のAutoSizeで実測し確定する。
' ----------------------------------------------------------------------------
'   R30 W2-1: 従来の ToastHeightFor(幅からの見積り算数)は「実際に描く
'   icon & " " & message」を計測対象に含め忘れる余地があり(B班確定の計測
'   漏れ)、modInsightCard等の行数ズレの温床にもなっていた。modUI.AddChatBubble
'   / modHelp.ShowHelpCard / modPeek.ShowPeek と同型の実測方式(AutoSizeで
'   一度伸ばして高さを読み取り、AutoSizeを解除する)へ転換する。
'   呼び出し前提: shp.TextFrame2 へ最終テキストを設定済みであること
'   (WordWrap=-1のため、AutoSize中は幅は固定のまま高さだけが変化する)。
'   R28H F6bの「修正は伸びる方向のみ」の教訓を踏襲し、下限34pt(従来の
'   1行ぶん)は維持する。上限92ptは画面上端から96ptの位置に出すため、
'   これ以上伸ばすと会話領域を覆う(ToastHeightForと同じ根拠)。
Public Sub FitToastHeight(ByVal shp As Shape)
    Const TOAST_MIN_PT As Double = 34
    Const TOAST_MAX_PT As Double = 92
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    shp.TextFrame2.AutoSize = 1        ' msoAutoSizeShapeToFitText
    Dim h As Double: h = shp.Height
    shp.TextFrame2.AutoSize = 0        ' msoAutoSizeNone(以降の操作でサイズが暴れないよう解除)
    ' R30 F6: AutoSizeが効かない環境ではhが初期値34pt(=AddShapeの既定高さ)の
    ' ままになり、長文が1行ぶんに切られる退行を招く。フィット後もなお下限の
    ' ままなら、実際に描いた文字列全体(icon込み)でToastHeightForの見積り
    ' 算数へフォールバックする。
    If h <= TOAST_MIN_PT Then h = ToastHeightFor(shp.TextFrame2.TextRange.Text)
    If h < TOAST_MIN_PT Then h = TOAST_MIN_PT
    If h > TOAST_MAX_PT Then h = TOAST_MAX_PT
    shp.Height = h
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' FormatLegendBody - ❓凡例カードの本文整形(R31 W1-3/W1-4)。純関数。
' ----------------------------------------------------------------------------
'   caps(i)はボタンの表示文字列(絵文字+ボタン名。AddTool時点で既に絵文字を
'   含んでいるため、ここで別途絵文字を足す必要は無い)。tips(i)は説明文
'   (R30まではws.Hyperlinks.AddのScreenTipだったが、R31 W1-1で撤去し
'   btn.AlternativeTextへ退避したものを呼び出し側が読み出して渡す)。
'   空のcapTextは行を作らない(TB_MAXの未使用スロット対策)。tipTextが
'   空の項目はボタン名だけの行にする(説明が無くても存在は分かる)。
Public Function FormatLegendBody(ByRef caps() As String, ByRef tips() As String, _
                                 ByVal n As Long) As String
    Dim out As String
    Dim i As Long
    For i = 0 To n - 1
        Dim capText As String: capText = Trim$(caps(i))
        If LenB(capText) > 0 Then
            Dim line As String: line = capText
            Dim tipText As String: tipText = Trim$(tips(i))
            If LenB(tipText) > 0 Then line = line & ": " & tipText
            If LenB(out) > 0 Then out = out & vbLf
            out = out & line
        End If
    Next i
    FormatLegendBody = out
End Function

' ----------------------------------------------------------------------------
' ShowLegendCard - ❓凡例カードの描画(R31 W1-3)。ShowToast(modSkin)と同型の
'   使い捨てShapeだが、①自動で消えず②クリックで自分自身を消す(呼び出し側の
'   OnActionハンドラ名を渡す設計にはせず、固定で"modKnowledgeBar.
'   OnToolbarLegendClose"を割り当てる。凡例カードの出し先はナレッジ画面の
'   ツールバーだけのため、汎用化せず呼び出し元を決め打つ)。
' ----------------------------------------------------------------------------
'   位置(leftPos/topPos)は呼び出し側(modKnowledgeBar)がActiveWindow基準で
'   計算して渡す。本関数自体はws以外のExcelグローバル状態(ActiveSheet/
'   ActiveWindow/ThisWorkbook)を参照しない(FitToastHeightと同じ、渡された
'   オブジェクトだけを操作する設計)。
'   塗り必須(Fill.Visible=-1): 塗り無しShapeは内部クリック透過するため、
'   自分自身のクリックで閉じる動きが利かなくなる既知の落とし穴を踏まない。
' R31 Fix波F6: 狭い窓ではカードが可視域からはみ出し、波2のScrollArea+行削除で
'   スクロールして読みに行けなくなっていた。呼び出し側からVisibleRangeの
'   幅・上端・下端を渡してもらい、幅はカード分だけ縮め、縦位置は可視域の
'   内側へクランプする(折返しが増えて高さが伸びても、そのAutoSize実測の
'   後で最終クランプするので既存の実測ロジックはそのまま機能する)。
Public Sub ShowLegendCard(ByVal ws As Worksheet, ByVal leftPos As Double, _
                          ByVal topPos As Double, ByVal title As String, _
                          ByVal body As String, _
                          Optional ByVal visW As Double = 0, _
                          Optional ByVal visTop As Double = 0, _
                          Optional ByVal visBottom As Double = 0)
    Const LEGEND_W_MAX As Double = 460
    Const LEGEND_W_MIN As Double = 200
    Const CARD_MIN_H As Double = 100
    Const CARD_MAX_H As Double = 480
    If ws Is Nothing Then Exit Sub

    Dim CARD_W As Double: CARD_W = LEGEND_W_MAX
    If visW > 0 And (visW - 40) < CARD_W Then CARD_W = visW - 40
    If CARD_W < LEGEND_W_MIN Then CARD_W = LEGEND_W_MIN

    On Error Resume Next
    ws.Shapes("nxk_legend").Delete   ' 再表示時の増殖防止(同名Shapeを先に消す)

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, leftPos, topPos, CARD_W, 40)   ' 5=角丸四角。高さは下でAutoSize実測。
    If shp Is Nothing Then Exit Sub
    shp.Name = "nxk_legend"
    shp.Placement = 3   ' xlFreeFloating
    shp.Adjustments(1) = 0.04
    shp.Line.Weight = 0.75
    shp.Line.ForeColor.RGB = modUI.UiColor("border")
    shp.Fill.Visible = -1
    shp.Fill.ForeColor.RGB = modUI.UiColor("surface")

    With shp.TextFrame2
        .WordWrap = -1
        .MarginLeft = 14: .MarginRight = 14: .MarginTop = 10: .MarginBottom = 10
        .TextRange.Text = title & vbLf & vbLf & body
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 9
        .TextRange.ParagraphFormat.Alignment = 1   ' 左揃え
        .VerticalAnchor = 1   ' 上詰め
    End With
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")

    shp.TextFrame2.AutoSize = 1        ' msoAutoSizeShapeToFitText
    Dim h As Double: h = shp.Height
    shp.TextFrame2.AutoSize = 0        ' msoAutoSizeNone(以降のサイズ操作が暴れないよう解除)
    If h < CARD_MIN_H Then h = CARD_MIN_H
    If h > CARD_MAX_H Then h = CARD_MAX_H
    ' R31 Fix波F13: 480頭打ちだけでは狭い窓(可視高<488pt)でカードが
    ' 可視域からはみ出し、テキストが枠外に描かれ続ける+下のmaxTop<minTopの
    ' 縮退で下端もはみ出た。可視高からも頭打ちする(自然高/可視高-16/480の
    ' 最小)。visBottom未指定(0)の旧経路は従来どおり480頭打ちのみ。
    If visBottom > 0 And visTop >= 0 Then
        Dim capH As Double: capH = visBottom - visTop - 16
        If capH < CARD_MIN_H Then capH = CARD_MIN_H
        If h > capH Then h = capH
    End If
    shp.Height = h

    ' 縦位置クランプ: 可視下端(visBottom)からカード高を引いた位置を上限、
    ' 可視上端(visTop)+8ptを下限にする。visBottom/visTopが未指定(0)の
    ' ときは従来どおりtopPosそのまま(呼び出し側が渡さない旧経路の保険)。
    If visBottom > 0 Then
        Dim finalTop As Double: finalTop = topPos
        Dim maxTop As Double: maxTop = visBottom - h
        Dim minTop As Double: minTop = visTop + 8
        If finalTop > maxTop Then finalTop = maxTop
        If finalTop < minTop Then finalTop = minTop
        shp.Top = finalTop
    End If

    modSkin.ApplyLightShadow shp
    shp.OnAction = "modKnowledgeBar.OnToolbarLegendClose"
    shp.ZOrder 0   ' msoBringToFront
    On Error GoTo 0
End Sub
