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

' ============================================================================
Public Sub RunAll30()
    On Error GoTo StretchFail30
    TestStretchToolbarRows30
    TestStretchToolbarRows30_ThreeRows
NextDone30:
    On Error GoTo 0
    Exit Sub

StretchFail30:
    modTestRunner.Check "TestStretchToolbarRows30(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone30
End Sub
