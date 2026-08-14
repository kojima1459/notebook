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
Public Sub RunAll30()
    On Error GoTo StretchFail30
    TestStretchToolbarRows30
    TestStretchToolbarRows30_ThreeRows
    TestHubDashSeedBurnIdempotent_F12
NextDone30:
    On Error GoTo 0
    Exit Sub

StretchFail30:
    modTestRunner.Check "TestStretchToolbarRows30(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone30
End Sub
