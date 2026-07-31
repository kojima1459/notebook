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
