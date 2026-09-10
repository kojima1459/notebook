Attribute VB_Name = "modTestsPure47"
Option Explicit

' ============================================================================
' modTestsPure47 - R43 波B(UI/UXポリッシュ)の純ロジック回帰。
'   modTestRunner.RunAllPureTests から直接呼ばれる独立の別枝
'   (modTestsPure31〜45と同型)。
' ----------------------------------------------------------------------------
' 【何を固定するか(手計算の根拠。仕様 spec_20260910_R43 §2-1/2-10 対応)】
'
'   A modKnowledgeBar.OnAccentColor(accent面に乗せる文字色の輝度判定):
'     WCAG相対輝度 L = 0.2126*R'+0.7152*G'+0.0722*B'(sRGB線形化後)。
'     しきい値0.179は「白字コントラスト(1.05/(L+0.05))=濃色字コントラスト
'     ((L+0.05)/0.05)」となる分岐点(L=sqrt(1.05*0.05)-0.05)。
'     A1 dark/sakura/light共通accent RGB(0,168,89): L=0.2873(手計算)
'       →濃色RGB(17,24,39)との対比5.70:1・白との対比3.11:1 → 濃色を選ぶ。
'     A2 msad accent RGB(7,169,99): L=0.2932(手計算)
'       →濃色対比5.80:1・白対比3.06:1(既存コメントの実測値と一致) → 濃色。
'     A3 ocean accent RGB(0,145,200): L=0.2442(手計算)
'       →濃色対比4.97:1・白対比3.57:1 → 濃色(4.5:1をわずかに上回る境界例)。
'     A4 gold accent RGB(230,196,80)(「エグゼクティブ・ゴールド」・実機報告の
'       1.70:1=ほぼ読めない元凶): L=0.5688(手計算)
'       →白対比1.70:1(実機報告の数値と一致)・濃色対比10.46:1 → 濃色。
'     A5 中間グレーRGB(105,105,105): L=0.1413(手計算・0.179未満)
'       →白対比5.49:1・濃色対比3.23:1 → 白を選ぶ(白branchの検証)。
'     A6 純黒RGB(0,0,0): L=0 → 白を選ぶ(自明な境界)。
'     A7 純白RGB(255,255,255): L=1 → 濃色を選ぶ(自明な境界)。
'
'   B modHubBadge.BadgeSpans(獲得済みバッジの文字区間):
'     呼び出し元(modHub.DrawBadges)の連結規約 sb = sb & mark(2文字) & " "
'     (1文字) & title & "   "(区切り3文字) を1文字も変えず踏襲する。
'     区間は「mark+空白+資料名」まで(末尾の区切り3文字は含めない)。
'     B1 titles=["初回質問","10問達成","伝道師"] earned=(True,False,True):
'       badge0: segLen=3+Len("初回質問")=3+4=7 → 獲得→span(1,7)。
'         pos = 1+7+3 = 11。
'       badge1: segLen=3+Len("10問達成")=3+5=8 → 未獲得→span無し。
'         pos = 11+8+3 = 22。
'       badge2: segLen=3+Len("伝道師")=3+3=6 → 獲得→span(22,6)。
'       → cnt=2, starts=(1,22), lens=(7,6)。
'     B2 全件獲得 titles=("AB","C") earned=(True,True):
'       badge0: segLen=3+2=5→span(1,5)。pos=1+5+3=9。
'       badge1: segLen=3+1=4→span(9,4)。
'       → cnt=2, starts=(1,9), lens=(5,4)。
'     B3 全件未獲得 titles=("X","YY") earned=(False,False) → cnt=0。
'     B4 n=0 → cnt=0(配列に触れず早期return)。
' ============================================================================

Private Sub ChkLong47(ByVal label As String, ByVal got As Long, ByVal want As Long)
    modTestRunner.Check "R43-47-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub TestOnAccentColor47()
    ChkLong47 "A1_dark_sakura_light", modKnowledgeBar.OnAccentColor(RGB(0, 168, 89)), RGB(17, 24, 39)
    ChkLong47 "A2_msad", modKnowledgeBar.OnAccentColor(RGB(7, 169, 99)), RGB(17, 24, 39)
    ChkLong47 "A3_ocean", modKnowledgeBar.OnAccentColor(RGB(0, 145, 200)), RGB(17, 24, 39)
    ChkLong47 "A4_gold", modKnowledgeBar.OnAccentColor(RGB(230, 196, 80)), RGB(17, 24, 39)
    ChkLong47 "A5_midgray_white", modKnowledgeBar.OnAccentColor(RGB(105, 105, 105)), RGB(255, 255, 255)
    ChkLong47 "A6_black_white", modKnowledgeBar.OnAccentColor(RGB(0, 0, 0)), RGB(255, 255, 255)
    ChkLong47 "A7_white_dark", modKnowledgeBar.OnAccentColor(RGB(255, 255, 255)), RGB(17, 24, 39)

    ' A8〜A11(司令塔 Fix): primary/danger の地。旧しきい値方式(0.179)では
    ' A8 が濃色に倒れて 3.93:1(AA未達)になっていた。白と濃色の実コントラストを
    ' 比べる方式なら、白 4.52:1 が選ばれる。数値はすべて python で再計算した実測。
    '   A8  light primary  RGB(0,137,62)   白 4.52 / 濃色 3.93 → 白
    '   A9  gold  primary  RGB(212,175,55) 白 2.10 / 濃色 8.44 → 濃色
    '   A10 gold  danger   RGB(248,113,113) 白 2.77 / 濃色 6.41 → 濃色
    '   A11 sakura danger  RGB(185,28,28)  白 6.47 / 濃色 2.74 → 白
    ChkLong47 "A8_light_primary_white", modKnowledgeBar.OnAccentColor(RGB(0, 137, 62)), RGB(255, 255, 255)
    ChkLong47 "A9_gold_primary_dark", modKnowledgeBar.OnAccentColor(RGB(212, 175, 55)), RGB(17, 24, 39)
    ChkLong47 "A10_gold_danger_dark", modKnowledgeBar.OnAccentColor(RGB(248, 113, 113)), RGB(17, 24, 39)
    ChkLong47 "A11_sakura_danger_white", modKnowledgeBar.OnAccentColor(RGB(185, 28, 28)), RGB(255, 255, 255)
End Sub

Private Sub TestBadgeSpans47()
    Dim starts() As Long, lens() As Long, cnt As Long

    Dim t1(0 To 2) As String, e1(0 To 2) As Boolean
    t1(0) = "初回質問": t1(1) = "10問達成": t1(2) = "伝道師"
    e1(0) = True: e1(1) = False: e1(2) = True
    cnt = modHubBadge.BadgeSpans(t1, e1, 3, starts, lens)
    ChkLong47 "B1_cnt", cnt, 2
    ChkLong47 "B1_start0", starts(0), 1
    ChkLong47 "B1_len0", lens(0), 7
    ChkLong47 "B1_start1", starts(1), 22
    ChkLong47 "B1_len1", lens(1), 6

    Dim t2(0 To 1) As String, e2(0 To 1) As Boolean
    t2(0) = "AB": t2(1) = "C"
    e2(0) = True: e2(1) = True
    cnt = modHubBadge.BadgeSpans(t2, e2, 2, starts, lens)
    ChkLong47 "B2_cnt", cnt, 2
    ChkLong47 "B2_start0", starts(0), 1
    ChkLong47 "B2_len0", lens(0), 5
    ChkLong47 "B2_start1", starts(1), 9
    ChkLong47 "B2_len1", lens(1), 4

    Dim t3(0 To 1) As String, e3(0 To 1) As Boolean
    t3(0) = "X": t3(1) = "YY"
    e3(0) = False: e3(1) = False
    cnt = modHubBadge.BadgeSpans(t3, e3, 2, starts, lens)
    ChkLong47 "B3_cnt", cnt, 0

    Dim t4() As String, e4() As Boolean
    cnt = modHubBadge.BadgeSpans(t4, e4, 0, starts, lens)
    ChkLong47 "B4_cnt", cnt, 0
End Sub

Public Sub RunAll47()
    On Error GoTo H01Fail47
    TestOnAccentColor47
H02Next47:
    On Error GoTo H02Fail47
    TestBadgeSpans47
H01Done47:
    On Error GoTo 0
    Exit Sub

H01Fail47:
    modTestRunner.Check "TestOnAccentColor47(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H02Next47
H02Fail47:
    modTestRunner.Check "TestBadgeSpans47(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done47
End Sub
