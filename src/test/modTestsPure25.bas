Attribute VB_Name = "modTestsPure25"
Option Explicit

' ============================================================================
' modTestsPure25 - R27H(敵対的レビュー裁定Fix波)の純ロジック回帰テスト。
'   modTestsPure24(28,068字)は30,000字上限まで残り1,932字しかなく、F3(私用
'   領域の化け検出)とF5(行頭インデント温存)の追随でその枠を使い切るため、
'   憲章§4-6に従って新設した分割先(modTestsPure21がR20H Fix波用に新設された
'   のと同型の理由)。入口は modTestsPure24.RunAll24 の末尾から呼ばれる
'   RunAll25 の1本。
' ----------------------------------------------------------------------------
' ここで固定するもの:
'   ・F1: modSparse.DiversitySwapPick — 最終hitsへの「最小介入」の3分岐。
'         前身の modSparse.DiversityOrder(候補プールの全面再配列)は
'           (a) 再ランクを通さない⚡すぐ聞く経路で topK 枠が各資料の代表で
'               埋まり、上位チャンクを最大 topK-1 件押し出す
'           (b) 分散判定(modClarify.HasScoreDispersion)は順序を見ないので、
'               狙いだった「資料が2種類以上」には一切効かない
'         と裁定され、modAskRetrieve からの呼び出しを撤去した(純関数と
'         modTestsPure24 の DiversityOrder 3群は将来用+テスト資産として残置)。
'         代わりに入れたのが「最終hitsが1資料へ収束した時だけ最下位1件を、
'         pool内の別資料の最高スコアへ替える」という1件だけの介入で、
'         その添字計算が DiversitySwapPick。
' ============================================================================

' ----------------------------------------------------------------------------
' F1: (i)2資料混在は無介入 / (ii)1資料hits+pool別資料ありは1件だけ差し替え /
'     (iii)pool別資料なしは無介入。加えて「hitsが1件のときは動かさない」
'     (替えると全体1位が消える=撤去した押し出しの再現になる)。
' ----------------------------------------------------------------------------
Private Sub TestDiversitySwapPick25()
    ' 実機の形をそのまま写す: 105チャンクある「就業規則」がpoolを占有し、
    ' 少数派の「興行中止保険特約」がpoolに残っている状態。
    Dim ps(1 To 4) As String
    Dim sc(1 To 4) As Double
    ps(1) = "就業規則.pdf": sc(1) = 0.9
    ps(2) = "就業規則.pdf": sc(2) = 0.7
    ps(3) = "興行中止保険特約.doc": sc(3) = 0.6
    ps(4) = "興行中止保険特約.doc": sc(4) = 0.8

    ' (i) hitsに2資料が混ざっている = 選択肢は既に出ているので動かさない。
    Dim hs(1 To 3) As String
    hs(1) = "就業規則.pdf": hs(2) = "就業規則.pdf": hs(3) = "興行中止保険特約.doc"
    Dim p1 As Long: p1 = modSparse.DiversitySwapPick(hs, 3, ps, sc, 4)
    modTestRunner.Check "R27H-F1_(i)2資料が混ざるhitsは無介入(0)", (p1 = 0), "pick=" & p1

    ' (ii) 1資料へ収束 → pool内の【別資料の最高スコア】(添字4=0.8)を指す。
    '      添字3(0.6)ではないこと=同一資料内の順位を正しく比べている証拠。
    hs(3) = "就業規則.pdf"
    Dim p2 As Long: p2 = modSparse.DiversitySwapPick(hs, 3, ps, sc, 4)
    modTestRunner.Check "R27H-F1_(ii)1資料hitsは次点資料の最高1件を指す(添字4)", _
        (p2 = 4), "pick=" & p2

    ' (iii) poolが同じ資料だけ = 差し替え先が無いので動かさない。
    Dim ps2(1 To 2) As String
    Dim sc2(1 To 2) As Double
    ps2(1) = "就業規則.pdf": sc2(1) = 0.9
    ps2(2) = "就業規則.pdf": sc2(2) = 0.5
    Dim p3 As Long: p3 = modSparse.DiversitySwapPick(hs, 3, ps2, sc2, 2)
    modTestRunner.Check "R27H-F1_(iii)poolに別資料が無ければ無介入(0)", (p3 = 0), "pick=" & p3

    ' hitsが1件しか無いときは動かさない(替えると全体1位そのものが消える)。
    Dim p4 As Long: p4 = modSparse.DiversitySwapPick(hs, 1, ps, sc, 4)
    modTestRunner.Check "R27H-F1_hitsが1件なら無介入(全体1位を守る)", (p4 = 0), "pick=" & p4

    ' 資料名が空のhitsは「どの資料へ収束したか」が決まらないので無介入。
    Dim hb(1 To 2) As String
    hb(1) = "": hb(2) = ""
    Dim p5 As Long: p5 = modSparse.DiversitySwapPick(hb, 2, ps, sc, 4)
    modTestRunner.Check "R27H-F1_資料名が空のhitsは無介入(0)", (p5 = 0), "pick=" & p5

    ' pool側の空の資料名は差し替え先にならない(空へ替えると出典が消える)。
    Dim ps3(1 To 2) As String
    Dim sc3(1 To 2) As Double
    ps3(1) = "就業規則.pdf": sc3(1) = 0.5
    ps3(2) = "": sc3(2) = 0.99
    Dim p6 As Long: p6 = modSparse.DiversitySwapPick(hs, 3, ps3, sc3, 2)
    modTestRunner.Check "R27H-F1_pool側の空資料名は差し替え先にしない(0)", (p6 = 0), "pick=" & p6
End Sub

' ============================================================================
Public Sub RunAll25()
    On Error GoTo SwapFail25
    TestDiversitySwapPick25
NextDone25:
    On Error GoTo 0
    Exit Sub

SwapFail25:
    modTestRunner.Check "TestDiversitySwapPick25(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone25
End Sub
