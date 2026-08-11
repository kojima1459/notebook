Attribute VB_Name = "modTestsPure26"
Option Explicit

' ============================================================================
' modTestsPure26 - R26H F8レビューFix(提案a)の純ロジック回帰テスト。
'   modTestsPure25(残650字)には収まらないため、憲章§4-6に従って新設した
'   分割先(modTestsPure21/25がそれぞれR20H/R27H Fix波用に新設されたのと同型)。
'   入口は modTestsPure25.RunAll25 の末尾から呼ばれる RunAll26 の1本。
' ----------------------------------------------------------------------------
' ここで固定するもの:
'   ・modInsightCard.PickFullerAnswer — 考察メモの回答をどちらから採るかの
'     判定。nexus_hist_a を第一情報源にしたF8のR26H実装は、
'     modApp.SaveTurnForRestore が保存時に回答を700字で切る
'     (Replace(modUtil.SafeLeft(ans, 700), ";;;", " "))ため、考察メモの本文が
'     700字ちょうどで無言に途切れる副作用を持っていた。修正後は
'     modAsk.LastAnswerText()(画面のバブル本文=フル)が hist_a を先頭部分と
'     して含む(=同一ターンの回答である)場合だけフル本文を採用し、含まない
'     (=別ターンの回答が混ざっている)場合は従来どおり hist_a を採る
'     (別ターン混成の防止は崩さない)。
'   固定する3分岐:
'     (1) hist_aが700字で切られていて、fullAがその続きを持つ同一ターン
'         → フル本文を採用(700字切れが解消される)。
'     (2) fullAがhist_aを含まない別ターン(モード切替直後などのすれ違い)
'         → hist_aを採用(別ターン混成は防止されたまま)。
'     (3) hist_aが空(そもそも比較できない) → フル本文を採用。
' ============================================================================

' ----------------------------------------------------------------------------
' (1) 700字で切られたhist_a + それを先頭に含むフル本文 → フル採用。
'     実機の形をそのまま写す: 800字の回答をSaveTurnForRestoreが700字に切って
'     hist_aへ保存した後、同じターンでLastAnswerText()がフル800字を返す状態。
' ----------------------------------------------------------------------------
Private Sub TestPickFullerAnswerTruncatedSameTurn26()
    Dim fullA As String: fullA = String$(800, "あ")
    Dim histA As String: histA = Left$(fullA, 700)

    Dim picked As String: picked = modInsightCard.PickFullerAnswer(histA, fullA)
    modTestRunner.Check "R26H-F8Fix_700字切れは同一ターンならフル本文を採用", _
        (picked = fullA), "len=" & Len(picked) & "(期待800)"
    modTestRunner.Check "R26H-F8Fix_フル採用で700字の無言切れが解消される", _
        (Len(picked) > 700), "len=" & Len(picked)
End Sub

' ----------------------------------------------------------------------------
' (2) 別ターン: fullAがhist_aを先頭部分として含まない → hist_aを採用
'     (別ターン混成の防止はR26H F8の実装意図そのものなので、ここが緩むと
'     退行になる)。
' ----------------------------------------------------------------------------
Private Sub TestPickFullerAnswerDifferentTurn26()
    Dim histA As String: histA = "前のターンの回答です。"
    Dim fullA As String: fullA = "別のモードで出た全く別内容の回答です。"

    Dim picked As String: picked = modInsightCard.PickFullerAnswer(histA, fullA)
    modTestRunner.Check "R26H-F8Fix_別ターンはhist_aを採用(混成防止は維持)", _
        (picked = histA), picked

    ' 前方一致が偶然そこそこ長くても、完全な先頭一致でなければ別ターン扱い。
    Dim histA2 As String: histA2 = "免責事由の説明はこちらです"
    Dim fullA2 As String: fullA2 = "免責事由の説明はこちらとは違います(別ターン)"
    Dim picked2 As String: picked2 = modInsightCard.PickFullerAnswer(histA2, fullA2)
    modTestRunner.Check "R26H-F8Fix_途中から食い違う前方一致は別ターン扱い", _
        (picked2 = histA2), picked2
End Sub

' ----------------------------------------------------------------------------
' (3) hist_aが空 → フル本文を採用(比較できないので迷わずフルを使う)。
' ----------------------------------------------------------------------------
Private Sub TestPickFullerAnswerEmptyHistA26()
    Dim fullA As String: fullA = "hist_aがまだ無いときのフル本文。"
    Dim picked As String: picked = modInsightCard.PickFullerAnswer("", fullA)
    modTestRunner.Check "R26H-F8Fix_hist_aが空ならフル本文を採用", _
        (picked = fullA), picked

    ' 両方空でも例外を出さず空を返す(呼び出し元LastTurnがLenB=0で保存を断る)。
    Dim pickedBoth As String: pickedBoth = modInsightCard.PickFullerAnswer("", "")
    modTestRunner.Check "R26H-F8Fix_両方空なら空を返す(例外にならない)", _
        (LenB(pickedBoth) = 0), "[" & pickedBoth & "]"

    ' hist_aが700字未満(=切られていないケース)でfullAと完全一致するときも、
    ' 先頭一致かつ同じ長さでフル採用の分岐へ入ることを確認する(仕様注記の
    ' 「Len比較で長い方を採る」が退行しないための固定)。
    Dim shortA As String: shortA = "短い回答なので切られていません。"
    Dim picked3 As String: picked3 = modInsightCard.PickFullerAnswer(shortA, shortA)
    modTestRunner.Check "R26H-F8Fix_切られていない完全一致でもフル本文相当を採用", _
        (picked3 = shortA), picked3
End Sub

' ============================================================================
Public Sub RunAll26()
    On Error GoTo TruncFail26
    TestPickFullerAnswerTruncatedSameTurn26
NextDiffTurn26:
    On Error GoTo DiffTurnFail26
    TestPickFullerAnswerDifferentTurn26
NextEmpty26:
    On Error GoTo EmptyFail26
    TestPickFullerAnswerEmptyHistA26
NextDone26:
    On Error GoTo 0
    Exit Sub

TruncFail26:
    modTestRunner.Check "TestPickFullerAnswerTruncatedSameTurn26(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDiffTurn26
DiffTurnFail26:
    modTestRunner.Check "TestPickFullerAnswerDifferentTurn26(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextEmpty26
EmptyFail26:
    modTestRunner.Check "TestPickFullerAnswerEmptyHistA26(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone26
End Sub
