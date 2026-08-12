Attribute VB_Name = "modTestsPure28"
Option Explicit

' ============================================================================
' modTestsPure28 - R28波2(実機第13報②「逆質問の断線」)の純ロジック回帰。
'   modTestsPure26(残り僅少)には収まらないため新設した分割先(憲章§4-6。
'   modTestsPure21/25/26 がそれぞれのFix波用に新設されたのと同型)。
'   入口は modTestsPure26.RunAll26 の末尾から呼ばれる RunAll28 の1本。
' ----------------------------------------------------------------------------
' ここで固定するもの:
'   (A) modClarify.ParseClarifyReply — 逆質問への返事から
'       「資料番号」「意図番号」を読み取る規則。R28以前の PickSource は
'       半角数字の出現位置しか見ておらず、全角「２」や丸数字だけで答えた人の
'       資料選択が丸ごと消えていた(合成後の質問から「対象の資料:」が落ちる)。
'       半角/全角/丸数字の全組合せと、範囲外番号の棄却を固定する。
'   (B) modFollowup.ScopeNeedsWiden — 逆質問で選ばれた資料に絞って検索した
'       結果を採用するか、本棚全体へ広げ直すかの境界(0/1件は広げる・2件で
'       採用・-1=埋め込み失敗は広げない)。ここが緩むと、スコープ内で1件しか
'       当たらない質問が「資料が見つかりません」で終わる。
' ============================================================================

' 期待値を1組ずつ確かめる小さな道具。srcIdx/intentIdx の両方を1回で見る。
Private Sub CheckReply(ByVal label As String, ByVal reply As String, ByVal nSrc As Long, _
                       ByVal expSrc As Long, ByVal expIntent As Long)
    Dim s As Long, t As Long
    modClarify.ParseClarifyReply reply, nSrc, s, t
    modTestRunner.Check "R28-W2-1_" & label, _
        (s = expSrc And t = expIntent), _
        "src=" & s & "(期待" & expSrc & ") intent=" & t & "(期待" & expIntent & ")"
End Sub

' ----------------------------------------------------------------------------
' (A-1) 資料+意図を2つ打つ形。半角・全角・丸数字の全組合せ。
' ----------------------------------------------------------------------------
Private Sub TestParseClarifyReplyPairs28()
    ' 教えている打ち方そのもの(「2-②」)。
    CheckReply "半角数字と丸数字_2-②", "2-" & ChrW(&H2461), 4, 2, 2
    ' 全角で打つ人(「１－①」)。旧実装はここで資料選択が消えていた。
    CheckReply "全角数字と丸数字_１－①", ChrW(&HFF11) & ChrW(&HFF0D) & ChrW(&H2460), 4, 1, 1
    ' 丸数字が打てない環境("2-3")。
    CheckReply "半角のみ_2-3", "2-3", 4, 2, 3
    ' 全角のみ("２－３")。
    CheckReply "全角のみ_２－３", ChrW(&HFF12) & ChrW(&HFF0D) & ChrW(&HFF13), 4, 2, 3
    ' 同じ数字を2つ("2-2")。modRagParse.ParseChoiceNumbers は重複を落とすため、
    ' 資料側と意図側を割らずに渡すと意図が消える(この1件が割る理由そのもの)。
    CheckReply "同番号_2-2は意図が消えない", "2-2", 4, 2, 2
    ' 長音を全角ハイフンの代わりに打つ癖("1ー3")。
    CheckReply "長音区切り_1ー3", "1" & ChrW(&H30FC) & "3", 4, 1, 3
    ' 区切り無しの並べ打ち("2 3")。
    CheckReply "空白区切り_2 3", "2 3", 4, 2, 3
End Sub

' ----------------------------------------------------------------------------
' (A-2) 片方だけ打つ形。丸数字は【常に意図番号】が逆質問の表記契約なので、
'       「①」1つの返事を資料番号1と読んではならない。
' ----------------------------------------------------------------------------
Private Sub TestParseClarifyReplySingle28()
    CheckReply "丸数字だけ_①は意図1(資料ではない)", ChrW(&H2460), 4, 0, 1
    CheckReply "丸数字だけ_④は意図4", ChrW(&H2463), 4, 0, 4
    CheckReply "半角数字だけ_3は資料3", "3", 4, 3, 0
    CheckReply "全角数字だけ_２は資料2", ChrW(&HFF12), 4, 2, 0
    ' 「0. この中にない / わからない」= 資料を選ばない意思表示。
    CheckReply "ゼロ_0は資料も意図も選ばれない", "0", 4, 0, 0
    CheckReply "全角ゼロ_０も同じ", ChrW(&HFF10), 4, 0, 0

    ' 戻り値(どちらか読めたか)も固定する。
    Dim s As Long, t As Long
    modTestRunner.Check "R28-W2-1_戻り値_0は False", _
        (modClarify.ParseClarifyReply("0", 4, s, t) = False), "s=" & s & " t=" & t
    modTestRunner.Check "R28-W2-1_戻り値_丸数字だけでも True", _
        (modClarify.ParseClarifyReply(ChrW(&H2462), 4, s, t) = True), "s=" & s & " t=" & t
End Sub

' ----------------------------------------------------------------------------
' (A-3) 範囲外の棄却。提示していない番号を黙って採ると、利用者が選んでいない
'       資料で検索スコープを立てる(W2-2)ことになり、実害が検索側へ伝播する。
' ----------------------------------------------------------------------------
Private Sub TestParseClarifyReplyOutOfRange28()
    ' 資料は2件しか出していないのに「5」= 採らない。
    CheckReply "資料番号が範囲外_nSrc2に5", "5", 2, 0, 0
    ' 「5-1」は資料5(範囲外)だけ捨て、意図1は生かす。
    CheckReply "資料だけ範囲外_5-1", "5-1", 2, 0, 1
    ' 意図は①～⑤の5つだけ。9は採らない(資料3は生きる)。
    CheckReply "意図番号が範囲外_3-9", "3-9", 4, 3, 0
    ' ⑥は意図の範囲外。丸数字は資料番号にはならないので何も選ばれない。
    CheckReply "丸数字が範囲外_⑥", ChrW(&H2465), 4, 0, 0
    ' 2桁は選択肢に存在しない(旧 PickSource は "12" の中の "1" を拾っていた)。
    CheckReply "2桁_12は資料1と読まない", "12", 4, 0, 0
    ' 資料の選択肢を出していない逆質問(nSrc=0)では資料番号は成立しない。
    CheckReply "nSrcが0_資料番号は成立しない", "3", 0, 0, 0

    ' 番号ではない文章。modRagParse.ParseChoiceNumbers が【書き直し】とみなして
    ' 空を返すため、こちらも何も選ばない(打った文章が黙って捨てられない)。
    CheckReply "文章_第1条の適用範囲は?", "第1条の適用範囲は?", 4, 0, 0
    CheckReply "空文字", "", 4, 0, 0
End Sub

' ----------------------------------------------------------------------------
' (B) スコープ内ヒット不足の境界(modFollowup.ScopeNeedsWiden)。
'     RunDeepScoped(modAskRetrieve)の既存作法と同じ「2件で採用」。
' ----------------------------------------------------------------------------
Private Sub TestScopeNeedsWiden28()
    modTestRunner.Check "R28-W2-2_0件は本棚全体へ広げ直す", _
        (modFollowup.ScopeNeedsWiden(0) = True), ""
    modTestRunner.Check "R28-W2-2_1件も広げ直す", _
        (modFollowup.ScopeNeedsWiden(1) = True), ""
    modTestRunner.Check "R28-W2-2_2件は採用(広げない)", _
        (modFollowup.ScopeNeedsWiden(2) = False), ""
    modTestRunner.Check "R28-W2-2_5件は採用(広げない)", _
        (modFollowup.ScopeNeedsWiden(5) = False), ""
    ' -1 は埋め込み失敗。広げ直しても同じ結果にしかならないので広げない
    ' (ここが True になると、失敗するだけの検索をもう1回払うことになる)。
    modTestRunner.Check "R28-W2-2_埋め込み失敗(-1)は広げない", _
        (modFollowup.ScopeNeedsWiden(-1) = False), ""

    ' スコープが作れなかった端末(Dictionary不可)は Nothing で来る。
    ' 「採用」にしてしまうと、絞れていない検索結果を絞れたことにする。
    modTestRunner.Check "R28-W2-2_スコープNothingは採用しない", _
        (modFollowup.ClarifyScopeKept(Nothing, 5) = False), ""
End Sub

' ============================================================================
Public Sub RunAll28()
    On Error GoTo PairsFail28
    TestParseClarifyReplyPairs28
NextSingle28:
    On Error GoTo SingleFail28
    TestParseClarifyReplySingle28
NextRange28:
    On Error GoTo RangeFail28
    TestParseClarifyReplyOutOfRange28
NextWiden28:
    On Error GoTo WidenFail28
    TestScopeNeedsWiden28
NextDone28:
    On Error GoTo 0
    Exit Sub

PairsFail28:
    modTestRunner.Check "TestParseClarifyReplyPairs28(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextSingle28
SingleFail28:
    modTestRunner.Check "TestParseClarifyReplySingle28(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRange28
RangeFail28:
    modTestRunner.Check "TestParseClarifyReplyOutOfRange28(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextWiden28
WidenFail28:
    modTestRunner.Check "TestScopeNeedsWiden28(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone28
End Sub
