Attribute VB_Name = "modTestsPure30"
Option Explicit

' ============================================================================
' modTestsPure30 - R31波3(実機第16報F-C・案C「伸縮両端揃え」)の純ロジック
'   回帰。modTestsPure29(残3,669字)には新テストを載せる余地が足りない
'   ため新設した分割先(憲章§4-6。modTestsPure21/25/26/28/29と同型)。
'   入口は modTestsPure29.RunAll29 の末尾から呼ばれる RunAll30 の1本。
' ----------------------------------------------------------------------------
' 固定するもの: modKnowledgeBar.StretchToolbarRows(段ごとの比例配分伸縮。
'   Worksheet非依存の純関数)。
'   (a) 1段構成: 唯一の段が伸縮し、Δが伸び率上限+25%の内側なら比例配分の
'       合計が目標右端にちょうど一致する。
'   (b) 伸び率上限+25%: 目標が遠すぎると各ボタンの増分はwidths(i)*0.25で
'       頭打ちになり、目標には届かない(無理に配りきらない=許容)。
'   (c) 最終段は伸縮しない(段が2つ以上のとき)。最終段以外は伸縮する。
'   (d) Δ<=0(既に目標へ届いている・帯が異常に狭い)では縮めない。
' ============================================================================

Private Sub TestStretchToolbarRows30()
    Dim x0 As Double: x0 = 18   ' L+TB_PAD相当
    Dim gap As Double: gap = 5

    ' (a) 1段構成: Δが上限内なら比例配分の合計が目標右端にちょうど一致する。
    Dim w1(0 To 2) As Double: w1(0) = 40: w1(1) = 60: w1(2) = 20
    Dim r1(0 To 2) As Long: r1(0) = 0: r1(1) = 0: r1(2) = 0
    Dim xs1(0 To 2) As Double, uw1(0 To 2) As Double
    xs1(0) = x0: uw1(0) = 40
    xs1(1) = x0 + 40 + gap: uw1(1) = 60
    xs1(2) = x0 + 40 + gap + 60 + gap: uw1(2) = 20
    Dim target1 As Double: target1 = x0 + 40 + 60 + 20 + 2 * gap + 10   ' Δ=10(上限120*0.25=30の内側)
    modKnowledgeBar.StretchToolbarRows w1, r1, 3, 1, x0, target1, gap, xs1, uw1
    Dim right1 As Double: right1 = xs1(2) + uw1(2)
    modTestRunner.Check "R31-W3-1_1段構成は唯一の段を伸縮する(目標に一致)", _
        (Abs(right1 - target1) < 0.001), "right=" & right1 & " target=" & target1

    ' (b) 伸び率上限+25%: 目標が遠すぎると、各ボタンの増分は元の幅の25%で頭打ちになる。
    Dim w2(0 To 1) As Double: w2(0) = 100: w2(1) = 100
    Dim r2(0 To 1) As Long: r2(0) = 0: r2(1) = 0
    Dim xs2(0 To 1) As Double, uw2(0 To 1) As Double
    xs2(0) = x0: uw2(0) = 100
    xs2(1) = x0 + 100 + gap: uw2(1) = 100
    Dim target2 As Double: target2 = x0 + 100 + 100 + gap + 5000   ' Δ=5000(はるかに上限超過)
    modKnowledgeBar.StretchToolbarRows w2, r2, 2, 1, x0, target2, gap, xs2, uw2
    modTestRunner.Check "R31-W3-1_伸び率上限25%で頭打ち(ボタン0)", _
        (Abs(uw2(0) - 125) < 0.001), "useW(0)=" & uw2(0)
    modTestRunner.Check "R31-W3-1_伸び率上限25%で頭打ち(ボタン1)", _
        (Abs(uw2(1) - 125) < 0.001), "useW(1)=" & uw2(1)
    modTestRunner.Check "R31-W3-1_上限頭打ちは目標に届かない", _
        ((xs2(1) + uw2(1)) < target2), "right=" & (xs2(1) + uw2(1)) & " target=" & target2

    ' (c) 最終段は伸縮しない(2段以上のとき)。1段目は伸び、2段目(最終段)は不変。
    Dim w3(0 To 3) As Double: w3(0) = 40: w3(1) = 40: w3(2) = 30: w3(3) = 30
    Dim r3(0 To 3) As Long: r3(0) = 0: r3(1) = 0: r3(2) = 1: r3(3) = 1
    Dim xs3(0 To 3) As Double, uw3(0 To 3) As Double
    xs3(0) = x0: uw3(0) = 40
    xs3(1) = x0 + 40 + gap: uw3(1) = 40
    xs3(2) = x0: uw3(2) = 30                 ' 2段目は左端からやり直し(FlowLeft作法)
    xs3(3) = x0 + 30 + gap: uw3(3) = 30
    Dim origXs3_2 As Double: origXs3_2 = xs3(2)
    Dim origUw3_2 As Double: origUw3_2 = uw3(2)
    Dim origXs3_3 As Double: origXs3_3 = xs3(3)
    Dim origUw3_3 As Double: origUw3_3 = uw3(3)
    Dim target3 As Double: target3 = x0 + 300   ' どちらの段にとっても届かない遠い目標
    modKnowledgeBar.StretchToolbarRows w3, r3, 4, 2, x0, target3, gap, xs3, uw3
    modTestRunner.Check "R31-W3-1_最終段(2段目)は伸縮しない(xs不変)", _
        (xs3(2) = origXs3_2 And xs3(3) = origXs3_3), _
        "xs2=" & xs3(2) & "(元" & origXs3_2 & ") xs3=" & xs3(3) & "(元" & origXs3_3 & ")"
    modTestRunner.Check "R31-W3-1_最終段(2段目)は伸縮しない(useW不変)", _
        (uw3(2) = origUw3_2 And uw3(3) = origUw3_3), _
        "uw2=" & uw3(2) & "(元" & origUw3_2 & ") uw3=" & uw3(3) & "(元" & origUw3_3 & ")"
    modTestRunner.Check "R31-W3-1_最終段以外(1段目)は伸縮する", _
        (uw3(0) > 40), "useW(0)=" & uw3(0)

    ' (d) Δ<=0(既に目標へ届いている・帯が異常に狭い)では縮めない。
    Dim w4(0 To 0) As Double: w4(0) = 50
    Dim r4(0 To 0) As Long: r4(0) = 0
    Dim xs4(0 To 0) As Double, uw4(0 To 0) As Double
    xs4(0) = x0: uw4(0) = 50
    Dim target4 As Double: target4 = x0 + 10   ' 現在の右端(x0+50)より手前=Δ<0
    modKnowledgeBar.StretchToolbarRows w4, r4, 1, 1, x0, target4, gap, xs4, uw4
    modTestRunner.Check "R31-W3-1_Δ<=0では縮めない", (uw4(0) = 50), "useW(0)=" & uw4(0)
End Sub

' R31 Fix波 F1: rowN=3(段0/段1/段2)構成でのネガティブ回帰。
'   StretchToolbarRowsのDimはPrologue専用(手続きスコープ)のため、段ループ内で
'   rowRightをsumWと一緒に0リセットしないと、段0のrowRightが段1の計算に
'   持ち越り、段1(2段目以降)の伸縮量が過小になる(「不発」)。
'   本テストは段0(natural右端280)・段1(natural右端250)・段2(除外)の3段で、
'   段0のnatural右端(280)が段1のnatural右端(250)より大きい構成を用いる。
'   リセット漏れがあると、段1のdeltaが「target-段0のnatural右端(280)」=20
'   と誤計算され(正しくは target-250=50)、段1の右端はtarget未達のまま
'   (250+20=270≠300)になる。
Private Sub TestStretchToolbarRows30_ThreeRows()
    Dim x0 As Double: x0 = 18
    Dim gap As Double: gap = 5
    Dim target As Double: target = 300

    Dim w(0 To 4) As Double
    w(0) = 130: w(1) = 127          ' 段0: natural右端=280
    w(2) = 113: w(3) = 114          ' 段1: natural右端=250(段0より小さい)
    w(4) = 50                       ' 段2(最終段・除外): 内容は無関係

    Dim rws(0 To 4) As Long
    rws(0) = 0: rws(1) = 0: rws(2) = 1: rws(3) = 1: rws(4) = 2

    Dim xs(0 To 4) As Double, uw(0 To 4) As Double
    xs(0) = x0: uw(0) = w(0)
    xs(1) = x0 + w(0) + gap: uw(1) = w(1)
    xs(2) = x0: uw(2) = w(2)
    xs(3) = x0 + w(2) + gap: uw(3) = w(3)
    xs(4) = x0: uw(4) = w(4)

    modKnowledgeBar.StretchToolbarRows w, rws, 5, 3, x0, target, gap, xs, uw

    Dim right0 As Double: right0 = xs(1) + uw(1)
    Dim right1 As Double: right1 = xs(3) + uw(3)

    modTestRunner.Check "R31-Fix-F1_rowN3_段0の右端が目標に一致する", _
        (Abs(right0 - target) < 0.001), "right0=" & right0 & " target=" & target
    modTestRunner.Check "R31-Fix-F1_rowN3_段1(2段目)の右端が目標に一致する", _
        (Abs(right1 - target) < 0.001), "right1=" & right1 & " target=" & target
End Sub

' ----------------------------------------------------------------------------
' R31 Fix検証パスF12: Hub/DashのSeedBurn×境界(RowAtFloor相当)往復ゼロを
'   「実装を通る」形で固定する。旧F2テストはSeedBurnを経由せずReleaseRangeへ
'   手で選んだ一致値を渡すだけで、実装を旧コード(CLng丸め/maxRowCap不一致/
'   hdrH無視)へ戻しても落ちない恒真に近いテストだった。
'   ここではModelRowAtFloorでmodViewport.RowAtFloorの実アルゴリズム(行1〜N
'   の累積Top+Heightをyで切り下げ)をWorksheet非依存に再現し、SeedRowCapが
'   実際に呼ばれる形(modViewport2.SeedBurnの内部呼び出しと同じ引数の組み方)
'   で出す値と、独立に計算した「実境界(ModelRowAtFloor)+1」が一致することを
'   確認する ―― SeedRowCapとModelRowAtFloorは別実装なので、丸め方式や
'   maxRowCapの受け渡しがズレれば必ず不一致で落ちる。
' ----------------------------------------------------------------------------
Private Function ModelRowAtFloor(ByVal hdrH As Double, ByVal rowH As Double, _
                                 ByVal y As Double, ByVal maxRow As Long) As Long
    Dim result As Long: result = 1
    Dim acc As Double: acc = 0
    Dim i As Long, h As Double
    For i = 1 To maxRow
        If i = 1 Then h = hdrH Else h = rowH
        If acc + h > y Then Exit For
        acc = acc + h
        result = i
    Next i
    ModelRowAtFloor = result
End Function

Private Sub TestHubDashSeedBurnIdempotent_F12()
    ' (a) 縦に高い窓(Dash相当・行1もrowHの一様モデル・hdrH=rowH)。
    '     viewH=1000/rowH=15/maxRowCap=120(DASH_ROWS)は余裕を持って
    '     頭打ちに掛からない ―― 焼く行(seed)は実境界+1(keepRow)に一致し、
    '     ReleaseRangeは削除0(往復ゼロ)。
    Dim keepA As Long: keepA = ModelRowAtFloor(15, 15, 1000, 120) + 1
    Dim seedA As Long: seedA = modViewport2.SeedRowCap(1, 1000, 15, 0, 120)
    modTestRunner.Check "R31-F12a_縦に高い窓(1000pt)でseed=実境界+1に一致", _
        (seedA = keepA), "seed=" & seedA & " keep=" & keepA
    Dim fromA As Long, toA As Long, hitA As Boolean
    hitA = modViewport2.ReleaseRange(keepA, seedA, 4, 120, fromA, toA)
    modTestRunner.Check "R31-F12a_縦に高い窓は焼き≦残しで削除0", _
        (hitA = False), "hit=" & hitA

    ' (b) Hub相当(行1=ヘッダー48pt・行2以降15pt)。窓588pt(48+15*36の
    '     割り切れる境界)で塗り下端(RowAt=ModelRowAtFloorが割り切れる場合は
    '     ceilもfloorと同値)=keepRow-1になることを固定しつつ、レビューの
    '     実測値である窓600pt(割り切れない・keepRow=38)でも往復ゼロを取る。
    Dim keepB1 As Long: keepB1 = ModelRowAtFloor(48, 15, 588, 60) + 1
    Dim seedB1 As Long: seedB1 = modViewport2.SeedRowCap(2, 588 - 48, 15, 0, 60)
    modTestRunner.Check "R31-F12b_窓588pt(割り切れる)でseed=keepRowに一致(37行)", _
        (keepB1 = 38) And (seedB1 = keepB1), _
        "keep=" & keepB1 & " seed=" & seedB1
    modTestRunner.Check "R31-F12b_窓588ptの塗り下端はkeepRow-1(37行)", _
        (ModelRowAtFloor(48, 15, 588, 60) = keepB1 - 1), _
        "floor=" & ModelRowAtFloor(48, 15, 588, 60)

    Dim keepB2 As Long: keepB2 = ModelRowAtFloor(48, 15, 600, 60) + 1
    Dim seedB2 As Long: seedB2 = modViewport2.SeedRowCap(2, 600 - 48, 15, 0, 60)
    modTestRunner.Check "R31-F12b_窓600pt(実機第16報の実測値・keepRow=38)でseed一致", _
        (keepB2 = 38) And (seedB2 = keepB2), _
        "keep=" & keepB2 & " seed=" & seedB2
    Dim fromB As Long, toB As Long, hitB As Boolean
    hitB = modViewport2.ReleaseRange(keepB2, seedB2, 4, 60, fromB, toB)
    modTestRunner.Check "R31-F12b_窓600pt(Hub相当)は往復ゼロ(削除0)", _
        (hitB = False), "hit=" & hitB

    ' (c) maxRowCap頭打ち: 窓が絶対上限(60)を超えて要求しても、焼く行は
    '     60を超えない(旧セッションの焼き付き救済がmaxRowCapを破らない)。
    Dim seedC As Long: seedC = modViewport2.SeedRowCap(2, 5000, 15, 0, 60)
    modTestRunner.Check "R31-F12c_巨大な窓でもmaxRowCap(60)で頭打ちする", _
        (seedC = 60), "seed=" & seedC
    Dim seedC2 As Long: seedC2 = modViewport2.SeedRowCap(2, 5000, 15, 9999, 60)
    modTestRunner.Check "R31-F12c_使用済み下端の救済もmaxRowCapを破らない", _
        (seedC2 = 60), "seed=" & seedC2
End Sub

' ============================================================================
' R32 波1(W1-11): 「みんなの困りごと」の純ロジック回帰【ゼロからの新設】
' ----------------------------------------------------------------------------
' この機能は【一度も通しで動作確認されていない】。IsMine により自分の投稿は
' 自分の板に出ないため、1人で試している限り板は常に空で、実機では確かめ
' ようがない。守る自動テストも1件も無かった。ここの網羅性がそのまま品質になる。
' 固定するもの(いずれも modInsight の純関数):
'   GapListBuild(件数上限20/合計820字/「…ほか N 件」/1件100字クランプ)
'   SortGapDesc(created_at降順・4列が同じ行のまま動く・部分ソート)
'   ReasonText(全コード+未知コード+空)/ NormKey / IsGapRow / GapAged /
'   WithinWindow(連投抑止)/ NonceKey・NonceIsKnown(重複ガード)
' ============================================================================

' 板の入力配列(0=質問 1=部署 2=created_at 3=reason)へ1件書く小道具。
Private Sub SetGap32(ByRef g() As String, ByVal i As Long, ByVal q As String, _
                     ByVal dept As String, ByVal created As String, ByVal reason As String)
    g(0, i) = q
    g(1, i) = dept
    g(2, i) = created
    g(3, i) = reason
End Sub

' 同じ文字をn個並べた文字列(String$の第2引数の解釈差を避けて自前で作る)。
Private Function Rep32(ByVal ch As String, ByVal n As Long) As String
    Dim sb As String, i As Long
    For i = 1 To n
        sb = sb & ch
    Next i
    Rep32 = sb
End Function

' (1) 件数上限20件 +「…ほか N 件」。1件あたりが最小でも21件目は出さない。
Private Sub TestGapListBuild32_ItemCap()
    Dim g() As String: ReDim g(0 To 3, 0 To 29)
    Dim i As Long
    For i = 0 To 29
        SetGap32 g, i, "", "", "", ""
    Next i
    Dim s As String: s = modInsight.GapListBuild(g, 30)
    modTestRunner.Check "R32-W1-3_30件届いても表示は20件まで", _
        (InStr(s, "20. ") > 0 And InStr(s, "21. ") = 0), "len=" & Len(s)
    modTestRunner.Check "R32-W1-3_件数上限で切った残りは「ほか N 件」で残す", _
        (InStr(s, "ほか 10 件") > 0), "len=" & Len(s)
    modTestRunner.Check "R32-F3_合計820字以内(MsgBoxの無言欠落を作らない)", _
        (Len(s) <= 820), "len=" & Len(s)
    modTestRunner.Check "R32-W1-8_部署が空でも氏名は出さず「部署の記録なし」と書く", _
        (InStr(s, "部署の記録なし") > 0), "len=" & Len(s)
End Sub

' (2) 合計820字クランプ。長文が届いても字数で打ち切り、残数は文言で残す。
'   2026-08-14(R32 F3): 上限を900→820へ下げた(呼び出し側の案内文が実測162字で、
'   1,024−162=862 が本文に使える上限だったため)。打ち切り位置が6件目→5件目へ
'   下がるのが正しい追随で、ここの期待値もそれに合わせてある。
Private Sub TestGapListBuild32_CharCap()
    Dim g() As String: ReDim g(0 To 3, 0 To 9)
    Dim longQ As String: longQ = Rep32("あ", 500)
    Dim i As Long
    For i = 0 To 9
        SetGap32 g, i, longQ, "営業部", "2026-08-01 10:00", "no_hit"
    Next i
    Dim s As String: s = modInsight.GapListBuild(g, 10)
    modTestRunner.Check "R32-F3_1件500字が10件来ても820字以内に収める", _
        (Len(s) <= 820), "len=" & Len(s)
    modTestRunner.Check "R32-W1-3_字数で打ち切った残りも「ほか N 件」で残す", _
        (InStr(s, "ほか 5 件") > 0), "len=" & Len(s)
    modTestRunner.Check "R32-F3_打ち切りは5件目まで(6件目は出さない)", _
        (InStr(s, vbLf & "5. ") > 0 And InStr(s, vbLf & "6. ") = 0), "len=" & Len(s)
    modTestRunner.Check "R32-W1-3_1件の質問文は100字で切る", _
        (InStr(s, Rep32("あ", 100)) > 0 And InStr(s, Rep32("あ", 101)) = 0), "len=" & Len(s)
End Sub

' (3) created_at の降順(「新しい順」の名乗りと実際を一致させる)。
Private Sub TestGapListBuild32_SortOrder()
    Dim g() As String: ReDim g(0 To 3, 0 To 2)
    SetGap32 g, 0, "古いほうの質問", "営業部", "2026-08-01 09:00", "no_hit"
    SetGap32 g, 1, "新しいほうの質問", "商品部", "2026-08-03 09:00", "wrong"
    SetGap32 g, 2, "真ん中の質問", "人事部", "2026-08-02 09:00", "low_conf"
    Dim s As String: s = modInsight.GapListBuild(g, 3)
    Dim p1 As Long: p1 = InStr(s, "新しいほうの質問")
    Dim p2 As Long: p2 = InStr(s, "真ん中の質問")
    Dim p3 As Long: p3 = InStr(s, "古いほうの質問")
    modTestRunner.Check "R32-W1-7_受信順ではなくcreated_atの降順で並ぶ", _
        (p1 > 0 And p2 > p1 And p3 > p2), "p=" & p1 & "/" & p2 & "/" & p3
    modTestRunner.Check "R32-W1-3_全部入りきるなら「ほか」は付けない", _
        (InStr(s, "ほか") = 0), s
    modTestRunner.Check "R32-m4_理由コードは日本語で出す(生英語を出さない)", _
        (InStr(s, "回答が違うと報告") > 0 And InStr(s, "wrong") = 0), s
End Sub

' (4) 並べ替えそのもの。4列が同じ行のまま動くこと・部分ソートの保証範囲。
Private Sub TestSortGapDesc32()
    Dim g() As String: ReDim g(0 To 3, 0 To 3)
    SetGap32 g, 0, "Q-A", "部A", "2026-08-01 00:00", "no_hit"
    SetGap32 g, 1, "Q-B", "部B", "2026-08-04 00:00", "wrong"
    SetGap32 g, 2, "Q-C", "部C", "2026-08-02 00:00", "low_conf"
    SetGap32 g, 3, "Q-D", "部D", "2026-08-03 00:00", "correction"
    modInsight.SortGapDesc g, 4, 4
    modTestRunner.Check "R32-W1-7_created_atが降順に並ぶ", _
        (g(2, 0) > g(2, 1) And g(2, 1) > g(2, 2) And g(2, 2) > g(2, 3)), _
        g(2, 0) & "/" & g(2, 1) & "/" & g(2, 2) & "/" & g(2, 3)
    modTestRunner.Check "R32-W1-7_並べ替えても質問・部署・理由が同じ行のまま動く", _
        (g(0, 0) = "Q-B" And g(1, 0) = "部B" And g(3, 0) = "wrong" And _
         g(0, 3) = "Q-A" And g(1, 3) = "部A" And g(3, 3) = "no_hit"), _
        g(0, 0) & "/" & g(1, 0) & "/" & g(3, 0) & " … " & g(0, 3)

    ' 部分ソート: topK=1 でも先頭だけは必ず最新になる(表示は最大20件なので
    ' 500行を全部泡立てない設計。保証するのは上位topK件だけ)。
    Dim h() As String: ReDim h(0 To 3, 0 To 3)
    SetGap32 h, 0, "Q-A", "部A", "2026-08-01 00:00", "no_hit"
    SetGap32 h, 1, "Q-B", "部B", "2026-08-04 00:00", "wrong"
    SetGap32 h, 2, "Q-C", "部C", "2026-08-02 00:00", "low_conf"
    SetGap32 h, 3, "Q-D", "部D", "2026-08-03 00:00", "no_hit"
    modInsight.SortGapDesc h, 4, 1
    modTestRunner.Check "R32-W1-7_topK=1でも先頭は最新になる", _
        (h(2, 0) = "2026-08-04 00:00" And h(0, 0) = "Q-B"), h(2, 0) & "/" & h(0, 0)

    ' 1件・0件で落ちない(受信箱が空・1件だけの日は毎日ある)。
    Dim z() As String: ReDim z(0 To 3, 0 To 0)
    SetGap32 z, 0, "Q-Z", "部Z", "2026-08-01 00:00", "no_hit"
    modInsight.SortGapDesc z, 1, 1
    modTestRunner.Check "R32-W1-7_1件でも壊れない", (z(0, 0) = "Q-Z"), z(0, 0)
End Sub

' (5) 理由コードの日本語訳(未知コードを生英語で出さない)。
Private Sub TestReasonText32()
    modTestRunner.Check "R32-m4_no_hit", _
        (modInsight.ReasonText("no_hit") = "本棚に該当資料なし"), modInsight.ReasonText("no_hit")
    modTestRunner.Check "R32-m4_low_conf", _
        (modInsight.ReasonText("low_conf") = "根拠が薄い"), modInsight.ReasonText("low_conf")
    modTestRunner.Check "R32-m4_wrong", _
        (modInsight.ReasonText("wrong") = "回答が違うと報告"), modInsight.ReasonText("wrong")
    modTestRunner.Check "R32-m4_correctionが生英語で露出しない", _
        (modInsight.ReasonText("correction") = "回答への訂正"), modInsight.ReasonText("correction")
    modTestRunner.Check "R32-m4_空は「理由の記録なし」", _
        (modInsight.ReasonText("") = "理由の記録なし"), modInsight.ReasonText("")
    modTestRunner.Check "R32-m4_未知コードも日本語の器に入れて出す", _
        (modInsight.ReasonText("weird_code") = "その他(weird_code)"), modInsight.ReasonText("weird_code")
    modTestRunner.Check "R32-m4_前後空白と大文字を吸収する", _
        (modInsight.ReasonText(" NO_HIT ") = "本棚に該当資料なし"), modInsight.ReasonText(" NO_HIT ")
    modTestRunner.Check "R32-m4_長い未知コードは20字で切る(板の桁を壊さない)", _
        (Len(modInsight.ReasonText(Rep32("z", 100))) = Len("その他()") + 20), _
        modInsight.ReasonText(Rep32("z", 100))
End Sub

' (6) 連投抑止の照合キー(解決済みQ&A側の SameQuestionCount と同じ正規化)。
Private Sub TestNormKey32()
    modTestRunner.Check "R32-W1-6_空白と疑問符の違いは同じキーになる", _
        (modInsight.NormKey("在庫の 締め日は?") = modInsight.NormKey("在庫の締め日は" & ChrW(&HFF1F))), _
        modInsight.NormKey("在庫の 締め日は?") & " vs " & modInsight.NormKey("在庫の締め日は" & ChrW(&HFF1F))
    modTestRunner.Check "R32-W1-6_全角スペース・読点・句点・全角括弧を落とす", _
        (modInsight.NormKey("a" & ChrW(&H3000) & "b" & ChrW(&H3001) & "c" & ChrW(&H3002) & _
                            ChrW(&HFF08) & "d" & ChrW(&HFF09)) = "abcd"), _
        modInsight.NormKey("a" & ChrW(&H3000) & "b" & ChrW(&H3001) & "c" & ChrW(&H3002) & _
                           ChrW(&HFF08) & "d" & ChrW(&HFF09))
    modTestRunner.Check "R32-W1-6_大文字小文字を区別しない", _
        (modInsight.NormKey("ABC") = "abc"), modInsight.NormKey("ABC")
    modTestRunner.Check "R32-W1-6_見るのは先頭40字だけ", _
        (Len(modInsight.NormKey(Rep32("x", 100))) = 40), _
        CStr(Len(modInsight.NormKey(Rep32("x", 100))))
    modTestRunner.Check "R32-W1-6_別の質問は別のキーになる(抑止しすぎない)", _
        (modInsight.NormKey("在庫の締め日") <> modInsight.NormKey("納期の締め日")), _
        modInsight.NormKey("在庫の締め日") & " vs " & modInsight.NormKey("納期の締め日")
End Sub

' (7) 板に出す行の判定(件数と一覧の食い違い・旧訂正データの混入)。
Private Sub TestIsGapRow32()
    modTestRunner.Check "R32-r3_未取込のgapは板に出す", _
        modInsight.IsGapRow("gap", "no_hit", "")
    modTestRunner.Check "R32-r3_取込済み(consumed=1)は板に出さない", _
        (Not modInsight.IsGapRow("gap", "no_hit", "1"))
    modTestRunner.Check "R32-W1-1_旧データの訂正(kind=gap/reason=correction)は板に出さない", _
        (Not modInsight.IsGapRow("gap", "correction", ""))
    modTestRunner.Check "R32-W1-1_大文字混じり・前後空白のcorrectionも弾く", _
        (Not modInsight.IsGapRow("gap", " Correction ", ""))
    modTestRunner.Check "R32-r3_解決済みQ&A(qa)は板に出さない", _
        (Not modInsight.IsGapRow("qa", "", ""))
    modTestRunner.Check "R32-W1-1_新しいkind=correctionも板に出さない", _
        (Not modInsight.IsGapRow("correction", "", ""))
End Sub

' (8) 保持期間(件数が永久に減らない問題の止め方)。
Private Sub TestGapAged32()
    Const CUT As String = "2026-07-15"
    modTestRunner.Check "R32-W1-4_期限より古いgapは落とす", _
        modInsight.GapAged("gap", "2026-07-14 09:00", CUT)
    modTestRunner.Check "R32-W1-4_期限当日は残す(境界は残す側)", _
        (Not modInsight.GapAged("gap", "2026-07-15 09:00", CUT))
    modTestRunner.Check "R32-W1-4_新しいgapは残す", _
        (Not modInsight.GapAged("gap", "2026-08-01 09:00", CUT))
    modTestRunner.Check "R32-W1-4_訂正も同じ保持期間で落とす(二度と減らない行を作らない)", _
        modInsight.GapAged("correction", "2026-07-14 09:00", CUT)
    modTestRunner.Check "R32-W1-4_解決済みQ&A(qa)は保持期間の対象外(読む前に捨てない)", _
        (Not modInsight.GapAged("qa", "2020-01-01 09:00", CUT))
    modTestRunner.Check "R32-W1-4_created_atが無い行は落とさない(判定材料が無い)", _
        (Not modInsight.GapAged("gap", "", CUT))
    modTestRunner.Check "R32-W1-4_保持期間が無効(cutoff空)なら1件も落とさない", _
        (Not modInsight.GapAged("gap", "2000-01-01 09:00", ""))
    modTestRunner.Check "R32-W1-4_セルの日付型で崩れたyyyy/mm/dd表記でも同じ判定になる", _
        modInsight.GapAged("gap", "2026/07/14 09:00", CUT)
End Sub

' (9) 連投抑止の期間判定(ISO文字列の辞書順比較。CDateを通さない)。
Private Sub TestWithinWindow32()
    Const LIM As String = "2026-08-13 10:00:00"
    modTestRunner.Check "R32-W1-6_抑止期間の内側なら再送しない", _
        modInsight.WithinWindow("2026-08-14 09:00:00", LIM)
    modTestRunner.Check "R32-W1-6_抑止期間を過ぎたら再送する", _
        (Not modInsight.WithinWindow("2026-08-13 09:59:59", LIM))
    modTestRunner.Check "R32-W1-6_境界ちょうどは期間外(送る側)へ倒す", _
        (Not modInsight.WithinWindow(LIM, LIM))
    modTestRunner.Check "R32-W1-6_記録が無ければ再送する(黙って捨てない)", _
        (Not modInsight.WithinWindow("", LIM))
    modTestRunner.Check "R32-W1-6_旧形式の数値(Bumpが書いた1)は再送する", _
        (Not modInsight.WithinWindow("1", LIM))
    modTestRunner.Check "R32-W1-6_期限側が空なら判定しない", _
        (Not modInsight.WithinWindow("2026-08-14 09:00:00", ""))
End Sub

' (10) nonce重複ガード(67日後の一斉再受信で本棚が二重登録される事故の止め方)。
Private Sub TestNonceGuard32()
    modTestRunner.Check "R32-W1-2_nonceキーは大文字小文字と前後空白を吸収する", _
        (modInsight.NonceKey("  User-20260814-00001  ") = "user-20260814-00001"), _
        modInsight.NonceKey("  User-20260814-00001  ")
    modTestRunner.Check "R32-W1-2_空のnonceはキーにならない(重複扱いにしない)", _
        (LenB(modInsight.NonceKey("   ")) = 0), "[" & modInsight.NonceKey("   ") & "]"
    ' 2026-08-14(R32 F11): ここは以前「集合が作れない環境(known=Nothing)なら
    ' False」しか撃っておらず、B1の核心である【集合に在るときTrueを返して
    ' 二重登録を止める】陽性経路が一度も検証されていなかった。判定を
    ' 文字列版(NonceMemoAdd / NonceIsKnownIn)へ移したので、両方を固定する。
    Dim memo As String
    modTestRunner.Check "R32-F11_空のメモには何も当たらない(初回は必ず足す)", _
        (Not modInsight.NonceIsKnownIn(memo, "user-20260814-00001")), "[" & memo & "]"

    memo = modInsight.NonceMemoAdd(memo, "User-20260814-00001")
    modTestRunner.Check "R32-F11_足した直後は【当たる】(陽性経路=二重登録を止める本体)", _
        (modInsight.NonceIsKnownIn(memo, "user-20260814-00001")), "[" & memo & "]"
    modTestRunner.Check "R32-F11_大文字小文字と前後空白が違っても当たる", _
        (modInsight.NonceIsKnownIn(memo, "  USER-20260814-00001 ")), "[" & memo & "]"
    modTestRunner.Check "R32-F11_別のnonceには当たらない(取りこぼしを作らない)", _
        (Not modInsight.NonceIsKnownIn(memo, "user-20260814-00002")), "[" & memo & "]"

    memo = modInsight.NonceMemoAdd(memo, "user-20260814-00002")
    modTestRunner.Check "R32-F11_2件目も当たり、1件目も残る", _
        (modInsight.NonceIsKnownIn(memo, "user-20260814-00001") And _
         modInsight.NonceIsKnownIn(memo, "user-20260814-00002")), "[" & memo & "]"
    ' 前方一致の取り違え防止: "user-20260814-0000" は "…-00001" に当たらない
    ' (キーの両端を "|" で挟んでいることの確認。modBackdrop.MemoKey と同型)。
    modTestRunner.Check "R32-F11_前方一致では誤ヒットしない", _
        (Not modInsight.NonceIsKnownIn(memo, "user-20260814-0000")), "[" & memo & "]"
    ' 同じnonceを二度足してもメモは太らない(受信箱に同じ行が2つあっても同じ)。
    Dim before As Long: before = Len(memo)
    memo = modInsight.NonceMemoAdd(memo, "USER-20260814-00001")
    modTestRunner.Check "R32-F11_同じnonceを足してもメモは伸びない", _
        (Len(memo) = before), "before=" & before & " after=" & Len(memo)
    ' 空のnonceはメモを壊さない(空キーが入ると以後の判定が全部当たる)。
    memo = modInsight.NonceMemoAdd(memo, "   ")
    modTestRunner.Check "R32-F11_空のnonceはメモに入れない", _
        (Len(memo) = before And Not modInsight.NonceIsKnownIn(memo, "")), "[" & memo & "]"
End Sub

' ----------------------------------------------------------------------------
' GroupFail30: グループ単位の実行時エラーを1件の失敗として記録する。
'   Err はハンドラを抜けると消えるため、呼び出し側で Err.Number/Err.Description
'   を実引数として渡し切ってから記録する(CLAUDE.md「ログの前にErrを退避」)。
'   ハンドラ本体をここへ寄せているのは容量のため(ラベルだけ増やす方式)。
' ----------------------------------------------------------------------------
Private Sub GroupFail30(ByVal nm As String, ByVal errNo As Long, ByVal errText As String)
    modTestRunner.Check nm & "(グループ全体)", False, _
        "実行時エラー: " & errText & " (Err=" & errNo & ")"
End Sub

' ============================================================================
' RunAll30 — 2026-08-15(R33波1 W1-3): 群ごとにハンドラを張り替える。
'   旧実装は先頭の1本(On Error GoTo StretchFail30)で13群すべてを包んでおり、
'   n本目で実行時エラーが起きると Resume NextDone30 で末尾へ飛ぶため
'   n+1本目以降が1件も実行されず、PASSにもFAILにも数えられずに総数が
'   黙って減っていた。実測(2026-08-15): 先頭の1群を落とすと後続12群の
'   75アサートが消え、報告は PASS 2421/FAIL 1 になった(健全時は PASS 2497)。
'   さらに失敗名が固定文字列で、どの群で落ちても先頭Subの名を騙っていた。
'   RunAll19/RunAll22 と同じ「群ごとに On Error GoTo → Next ラベル」へ戻す。
' ============================================================================
Public Sub RunAll30()
    On Error GoTo G01Fail30
    TestStretchToolbarRows30
G02Next30:
    On Error GoTo G02Fail30
    TestStretchToolbarRows30_ThreeRows
G03Next30:
    On Error GoTo G03Fail30
    TestHubDashSeedBurnIdempotent_F12
G04Next30:
    On Error GoTo G04Fail30
    TestGapListBuild32_ItemCap
G05Next30:
    On Error GoTo G05Fail30
    TestGapListBuild32_CharCap
G06Next30:
    On Error GoTo G06Fail30
    TestGapListBuild32_SortOrder
G07Next30:
    On Error GoTo G07Fail30
    TestSortGapDesc32
G08Next30:
    On Error GoTo G08Fail30
    TestReasonText32
G09Next30:
    On Error GoTo G09Fail30
    TestNormKey32
G10Next30:
    On Error GoTo G10Fail30
    TestIsGapRow32
G11Next30:
    On Error GoTo G11Fail30
    TestGapAged32
G12Next30:
    On Error GoTo G12Fail30
    TestWithinWindow32
G13Next30:
    On Error GoTo G13Fail30
    TestNonceGuard32
NextDone30:
    On Error GoTo 0
    Exit Sub

G01Fail30:
    GroupFail30 "TestStretchToolbarRows30", Err.Number, Err.Description
    Resume G02Next30
G02Fail30:
    GroupFail30 "TestStretchToolbarRows30_ThreeRows", Err.Number, Err.Description
    Resume G03Next30
G03Fail30:
    GroupFail30 "TestHubDashSeedBurnIdempotent_F12", Err.Number, Err.Description
    Resume G04Next30
G04Fail30:
    GroupFail30 "TestGapListBuild32_ItemCap", Err.Number, Err.Description
    Resume G05Next30
G05Fail30:
    GroupFail30 "TestGapListBuild32_CharCap", Err.Number, Err.Description
    Resume G06Next30
G06Fail30:
    GroupFail30 "TestGapListBuild32_SortOrder", Err.Number, Err.Description
    Resume G07Next30
G07Fail30:
    GroupFail30 "TestSortGapDesc32", Err.Number, Err.Description
    Resume G08Next30
G08Fail30:
    GroupFail30 "TestReasonText32", Err.Number, Err.Description
    Resume G09Next30
G09Fail30:
    GroupFail30 "TestNormKey32", Err.Number, Err.Description
    Resume G10Next30
G10Fail30:
    GroupFail30 "TestIsGapRow32", Err.Number, Err.Description
    Resume G11Next30
G11Fail30:
    GroupFail30 "TestGapAged32", Err.Number, Err.Description
    Resume G12Next30
G12Fail30:
    GroupFail30 "TestWithinWindow32", Err.Number, Err.Description
    Resume G13Next30
G13Fail30:
    GroupFail30 "TestNonceGuard32", Err.Number, Err.Description
    Resume NextDone30
End Sub
