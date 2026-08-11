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
'   ・R26-1: modGenPipe — 一般アシスタント3段化の純関数3本。
'         ParseVerdict(検証応答の判定)/ShouldRunVerifyLoop(周回上限)/
'         PlanFor(モード分岐表)。ここが緩むと入念モードが「永久に上限まで
'         回る」か「1周目で誤ってPASS扱いになる」のどちらかへ倒れ、
'         どちらも画面上は同じ「入念に聞く」に見えるため誰も気付けない。
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

' ----------------------------------------------------------------------------
' R26-1(1): modGenPipe.ParseVerdict — 検証応答の判定。
'   期待値は modGenPipe の公開Constではなく【文字列リテラル】で書く。
'   Constと突き合わせると、Constごと書き換えられたときにテストも一緒に
'   ずれて何も守らない(恒真アサートの一種)。
' ----------------------------------------------------------------------------
Private Sub TestGenParseVerdict25()
    Dim f As String
    Dim v As String

    ' (i) 素直なPASS。findings は必ず空(空でないと改稿へ進んでしまう)。
    v = modGenPipe.ParseVerdict("verdict:PASS", f)
    modTestRunner.Check "R26-1_verdict素直なPASS", (v = "PASS"), "v=" & v
    modTestRunner.Check "R26-1_PASS時のfindingsは空", (LenB(f) = 0), "f=[" & f & "]"

    ' (ii) 前後の空白・空行・全角コロン・装飾が混ざってもPASSと読む。
    '      ここが厳密一致だけだと、約束が1文字ぶれた日から入念は永久に
    '      上限まで回り続ける(利用者からは「ただ遅い」としか見えない)。
    v = modGenPipe.ParseVerdict("  " & vbCrLf & " verdict：PASS 。" & vbLf & vbLf, f)
    modTestRunner.Check "R26-1_verdict前後空白+全角コロン+句点でもPASS", (v = "PASS"), "v=" & v
    v = modGenPipe.ParseVerdict("**verdict: PASS**" & vbLf & "(問題ありません)", f)
    modTestRunner.Check "R26-1_verdict装飾付きでもPASS", (v = "PASS"), "v=" & v

    ' (iii) 反証: 「verdict:PASSではない」を PASS と読んではならない。
    '       「PASSを含むか」で判定すると、査読が否定形で書いた瞬間に
    '       指摘を丸ごと捨てて早期終了する。
    v = modGenPipe.ParseVerdict("verdict:PASSではない。第2段落の断定に根拠がない。", f)
    modTestRunner.Check "R26-1_verdictPASSではないをPASSと読まない", (v = "FINDINGS"), "v=" & v
    modTestRunner.Check "R26-1_指摘ありのfindingsは本文を持つ", (InStr(f, "根拠がない") > 0), "f=[" & f & "]"

    ' (iv) 指摘あり(通常形)。前後の空行は落として渡す。
    v = modGenPipe.ParseVerdict(vbLf & "・免責の範囲を断定している" & vbLf & "・例外条項の見落とし" & vbLf, f)
    modTestRunner.Check "R26-1_verdict指摘ありはFINDINGS", (v = "FINDINGS"), "v=" & v
    modTestRunner.Check "R26-1_findingsの先頭に空行を残さない", (Left$(f, 1) = "・"), "f=[" & Left$(f, 4) & "]"

    ' (v) 形式崩れ(空・空白のみ・#ERR)は FAIL。呼び出し側はここで周回を
    '     やめて【いま手元にある回答】を返す(利用者を待たせて壊さない)。
    v = modGenPipe.ParseVerdict("", f)
    modTestRunner.Check "R26-1_verdict空応答はFAIL", (v = "FAIL"), "v=" & v
    v = modGenPipe.ParseVerdict("  " & vbCrLf & vbTab & " ", f)
    modTestRunner.Check "R26-1_verdict空白だけの応答はFAIL", (v = "FAIL"), "v=" & v
    v = modGenPipe.ParseVerdict("#ERR:E0202:応答が空でした", f)
    modTestRunner.Check "R26-1_verdictエラー応答はFAIL", (v = "FAIL"), "v=" & v
    modTestRunner.Check "R26-1_FAIL時のfindingsは空", (LenB(f) = 0), "f=[" & f & "]"
End Sub

' ----------------------------------------------------------------------------
' R26-1(2): modGenPipe.ShouldRunVerifyLoop — 検証→改稿の周回上限。
'   上限を超えて回ると「答えは出ているのに何分も待たされる」になり、
'   0以下で回さないと入念がしっかりと同じものに退化する。両端を固定する。
' ----------------------------------------------------------------------------
Private Sub TestGenVerifyLoopBound25()
    modTestRunner.Check "R26-1_周回0/上限2は回す", _
        (modGenPipe.ShouldRunVerifyLoop(0, 2) = True), "0/2"
    modTestRunner.Check "R26-1_周回1/上限2は回す", _
        (modGenPipe.ShouldRunVerifyLoop(1, 2) = True), "1/2"
    modTestRunner.Check "R26-1_周回2/上限2で打ち切る", _
        (modGenPipe.ShouldRunVerifyLoop(2, 2) = False), "2/2"

    ' 上限0=検証しない(エスケープハッチ)。負値も同じ扱いへ倒す。
    modTestRunner.Check "R26-1_上限0は1周も回さない", _
        (modGenPipe.ShouldRunVerifyLoop(0, 0) = False), "0/0"
    modTestRunner.Check "R26-1_上限が負値でも回さない", _
        (modGenPipe.ShouldRunVerifyLoop(0, -5) = False), "0/-5"

    ' configに壊れた値(999)が入っても4周で頭打ち。ここが無いと
    ' 「設定を間違えた1回」でExcelが数十分固まる。
    modTestRunner.Check "R26-1_上限999でも3周目までは回す", _
        (modGenPipe.ShouldRunVerifyLoop(3, 999) = True), "3/999"
    modTestRunner.Check "R26-1_上限999でも4周で頭打ち", _
        (modGenPipe.ShouldRunVerifyLoop(4, 999) = False), "4/999"
End Sub

' ----------------------------------------------------------------------------
' R26-1(3): modGenPipe.PlanFor — 一般アシスタントのモード分岐表。
'   読めない値は必ず「すぐ聞く相当の1回呼び出し」へ倒す(不明を多段へ
'   倒すと、ui_stateが壊れた1回のために数分待たされる)。
'   併せて ResetTurn がフッター用の記録を捨てることを固定する
'   (捨て損ねると前ターンの「検証2回」が次の回答へ漏れる)。
' ----------------------------------------------------------------------------
Private Sub TestGenPlanTable25()
    modTestRunner.Check "R26-1_planすぐ聞く=single", _
        (modGenPipe.PlanFor("quick") = "single"), modGenPipe.PlanFor("quick")
    modTestRunner.Check "R26-1_planしっかり=single_deep", _
        (modGenPipe.PlanFor("deep") = "single_deep"), modGenPipe.PlanFor("deep")
    modTestRunner.Check "R26-1_plan入念=pipeline", _
        (modGenPipe.PlanFor("thorough") = "pipeline"), modGenPipe.PlanFor("thorough")

    ' 大文字・前後空白は modMode.Normalize を通って吸収される。
    modTestRunner.Check "R26-1_plan大文字と前後空白でも入念", _
        (modGenPipe.PlanFor("  THOROUGH ") = "pipeline"), modGenPipe.PlanFor("  THOROUGH ")

    ' 空・未知の値はフェイルセーフで現行動作(1回呼び出し)へ。
    modTestRunner.Check "R26-1_plan空文字はsingleへ倒す", _
        (modGenPipe.PlanFor("") = "single"), modGenPipe.PlanFor("")
    modTestRunner.Check "R26-1_plan未知の値はsingleへ倒す", _
        (modGenPipe.PlanFor("ちょうねんいり") = "single"), modGenPipe.PlanFor("ちょうねんいり")

    ' ResetTurn 後は「検証n回」を出さない。入念の次にすぐ聞くを撃っても
    ' 前の回の周回数がフッターへ残らないこと(モード間リークの防止)。
    modGenPipe.ResetTurn "thorough"
    modTestRunner.Check "R26-1_入念でも0周ならフッター注記は出さない", _
        (LenB(modGenPipe.VerifyFooterNote()) = 0), "[" & modGenPipe.VerifyFooterNote() & "]"
    modGenPipe.ResetTurn "quick"
    modTestRunner.Check "R26-1_すぐ聞くのフッター注記は必ず空", _
        (LenB(modGenPipe.VerifyFooterNote()) = 0), "[" & modGenPipe.VerifyFooterNote() & "]"
    modTestRunner.Check "R26-1_ResetTurn後のusage_log detailは初期値", _
        (modGenPipe.TurnDetail() = "loops=0 pass=0 parse_fail=0"), modGenPipe.TurnDetail()

    ' モード説明は3段の【両方の画面】の差を語ること。R21-2 D4 の
    ' 「一般アシスタントでは使われない」トースト注記を撤去した以上、
    ' 説明が社内ナレッジ検索の話だけに戻ると、一般アシスタント側の3段化が
    ' 利用者から見て存在しないものになる(既存のmodTestsPure8/12は
    ' 社内ナレッジ検索側の語(会話の流れ/本棚全体/多段検証)を固定している)。
    modTestRunner.Check "R26-1_しっかりの説明が一般アシスタントにも触れる", _
        (InStr(modMode.Description("deep"), "一般アシスタント") > 0), modMode.Description("deep")
    modTestRunner.Check "R26-1_入念の説明が一般アシスタントにも触れる", _
        (InStr(modMode.Description("thorough"), "一般アシスタント") > 0), modMode.Description("thorough")
End Sub

' ----------------------------------------------------------------------------
' R26-2(1): modConvBridge.TruncateTail — 4,000字境界の切り詰め(全角安全)。
'   「超過分は先頭を落とす」=末尾側(直近の内容)を優先して残す。境界を
'   跨いだ位置にサロゲートペア(絵文字)が来ると、ペアの片割れ(低位
'   サロゲート)だけが残ってはならない(表示が化ける・二度と結合できない)。
' ----------------------------------------------------------------------------
Private Sub TestConvBridgeTruncate25()
    ' (i) 上限ちょうど・上限未満は無介入。
    Dim exact As String: exact = String(4000, ChrW(&H3042))   ' "あ"×4000
    modTestRunner.Check "R26-2_切詰め_ちょうど4000字は無変化", _
        (modConvBridge.TruncateTail(exact, 4000) = exact), "len=" & Len(modConvBridge.TruncateTail(exact, 4000))
    Dim under As String: under = String(3999, ChrW(&H3044))   ' "い"×3999
    modTestRunner.Check "R26-2_切詰め_上限未満は無変化", _
        (modConvBridge.TruncateTail(under, 4000) = under), "len"

    ' (ii) 超過は末尾4000字だけ残す(先頭側を落とす)。マーカーで前後を区別する。
    Dim mixed As String
    mixed = String(5, ChrW(&H3042)) & String(4000, ChrW(&H3044))   ' 先頭"あ"×5+末尾"い"×4000=4005字
    Dim got As String: got = modConvBridge.TruncateTail(mixed, 4000)
    modTestRunner.Check "R26-2_切詰め_超過分は末尾4000字だけ残す", _
        (got = String(4000, ChrW(&H3044))), "len=" & Len(got)
    modTestRunner.Check "R26-2_切詰め_落ちた先頭の「あ」は含まない", _
        (InStr(got, ChrW(&H3042)) = 0), got

    ' (iii) 全角安全: 境界にサロゲートペア(絵文字)が跨がると、低位サロゲート
    '       だけの片割れを残さない(道連れでもう1字落とす)。emoji=🧩(2コード
    '       単位)を文字列の先頭2字に置き、合計4001字にして境界をペアの
    '       ちょうど真ん中(高位側)に当てる。
    Dim emoji As String: emoji = ChrW(&HD83E) & ChrW(&HDDE9)   ' 🧩(U+1F9E9)
    Dim withEmoji As String: withEmoji = emoji & String(3999, "A")   ' 2+3999=4001字
    Dim gotE As String: gotE = modConvBridge.TruncateTail(withEmoji, 4000)
    modTestRunner.Check "R26-2_切詰め_境界の絵文字は片割れを残さず道連れに落とす", _
        (gotE = String(3999, "A")), "len=" & Len(gotE) & " [" & Left$(gotE, 3) & "]"
    ' 低位サロゲート単独(&HDC00~&HDFFF)が先頭に残っていないことも直接確認する。
    Dim c0 As Long: c0 = AscW(Left$(gotE, 1))
    If c0 < 0 Then c0 = c0 + 65536
    modTestRunner.Check "R26-2_切詰め_結果の先頭は低位サロゲートではない", _
        (c0 < &HDC00& Or c0 > &HDFFF&), "code=&H" & Hex$(c0)

    ' (iv) 上限0/負値は空文字列(エスケープハッチ)。
    modTestRunner.Check "R26-2_切詰め_上限0は空文字列", _
        (modConvBridge.TruncateTail("何か文字列", 0) = ""), "[" & modConvBridge.TruncateTail("何か文字列", 0) & "]"
    modTestRunner.Check "R26-2_切詰め_上限が負値でも空文字列", _
        (modConvBridge.TruncateTail("何か文字列", -3) = ""), "[" & modConvBridge.TruncateTail("何か文字列", -3) & "]"
End Sub

' ----------------------------------------------------------------------------
' R26-2(2): modConvBridge.WithBridgeHeader — 出所ヘッダーの付与と冪等性。
'   モード切替を行き来しても【直前の○○での文脈】が積み重ならないこと
'   (二重付与しない)。
' ----------------------------------------------------------------------------
Private Sub TestConvBridgeHeader25()
    Dim h1 As String
    h1 = modConvBridge.WithBridgeHeader("社内ナレッジ検索", "回答本文です")
    modTestRunner.Check "R26-2_ヘッダー_先頭に付与される", _
        (Left$(h1, 1) = ChrW(&H3010)), h1   ' 【
    modTestRunner.Check "R26-2_ヘッダー_元モード名を含む", _
        (InStr(h1, "社内ナレッジ検索") > 0), h1
    modTestRunner.Check "R26-2_ヘッダー_本文は保たれる", _
        (InStr(h1, "回答本文です") > 0), h1

    ' 冪等性: 既にヘッダー済みの文字列へ「別モード名」で付け直そうとしても、
    ' 二重に積み上がらない(反証: 同じ関数を2回連続で呼んでも1回分のまま)。
    Dim h2 As String
    h2 = modConvBridge.WithBridgeHeader("一般アシスタント", h1)
    modTestRunner.Check "R26-2_ヘッダー_二重付与しない(冪等)", (h2 = h1), h2
    modTestRunner.Check "R26-2_ヘッダー_冪等時もヘッダーは1個だけ", _
        (CountOccurrences25(h2, ChrW(&H3010) & "直前の") = 1), "n=" & CountOccurrences25(h2, ChrW(&H3010) & "直前の")
End Sub

' "直前の"ヘッダーの出現回数を数える(冪等性の反証用)。
Private Function CountOccurrences25(ByVal s As String, ByVal needle As String) As Long
    Dim p As Long: p = 1
    Dim n As Long
    Do
        p = InStr(p, s, needle)
        If p = 0 Then Exit Do
        n = n + 1
        p = p + Len(needle)
    Loop
    CountOccurrences25 = n
End Function

' ----------------------------------------------------------------------------
' R26-2(3): modConvBridge.ComputeBridgeCore — 橋渡しの真理表(状態非依存)。
'   conv_bridge=off時は何もしない(off時の非動作)ことを、configシートに
'   依存せず直接固定する(ComputeBridgeCoreはenabledを引数で受け取るだけの
'   純関数。理由は modConvBridge.bas冒頭コメント参照)。
'
'   R26H F3(M-1)で仕様が「置換」から「先頭差し込み」へ変わった。旧テストは
'   (v)(vi)で outQ が橋渡しの1件【だけ】になることを固定しており、まさに
'   「切替先の既存の記憶が消える」挙動を守っていた。期待値を新仕様
'   (既存の記憶が残る・先頭に入る・上限で丸まる)へ書き換える。
' ----------------------------------------------------------------------------
Private Sub TestConvBridgeCore25()
    Dim oq As String, oa As String
    Dim ok As Boolean

    ' (i) conv_bridge=off: 記憶があっても一切橋渡ししない。
    ok = modConvBridge.ComputeBridgeCore("rag", "normal", False, "Q1", "A1", "", "", 3, oq, oa)
    modTestRunner.Check "R26-2_off時は橋渡ししない(戻り値False)", (ok = False), "ok=" & ok
    modTestRunner.Check "R26-2_off時はoutQ/outAも空", (LenB(oq) = 0 And LenB(oa) = 0), "[" & oq & "][" & oa & "]"

    ' (ii) 同一モードは何もしない(切替が起きていない)。
    ok = modConvBridge.ComputeBridgeCore("rag", "rag", True, "Q1", "A1", "", "", 3, oq, oa)
    modTestRunner.Check "R26-2_同一モードは橋渡ししない", (ok = False), "ok=" & ok

    ' (iii) 未知のモード値の組合せはフェイルセーフでFalse。
    ok = modConvBridge.ComputeBridgeCore("foo", "bar", True, "Q1", "A1", "", "", 3, oq, oa)
    modTestRunner.Check "R26-2_未知のモード組合せは橋渡ししない", (ok = False), "ok=" & ok

    ' (iv) 切替元の記憶が空なら橋渡し不要(トースト等の無駄打ちを避ける)。
    ok = modConvBridge.ComputeBridgeCore("rag", "normal", True, "", "", "既存Q", "既存A", 3, oq, oa)
    modTestRunner.Check "R26-2_切替元の記憶が空なら橋渡ししない", (ok = False), "ok=" & ok

    ' (v) RAG→一般: 引き継ぐのは直前1往復だけ(多往復の履歴から先頭のみ)で、
    '     回答側だけにヘッダーを付ける(質問側には付けない)。切替先が空のとき
    '     はその1件だけになる。
    ok = modConvBridge.ComputeBridgeCore("rag", "normal", True, _
        "Q最新;;;Q古い", "A最新;;;A古い", "", "", 3, oq, oa)
    modTestRunner.Check "R26-2_RAGから一般への橋渡しは成功", (ok = True), "ok=" & ok
    modTestRunner.Check "R26-2_質問側は直前1件のみでヘッダー無し", (oq = "Q最新"), oq
    modTestRunner.Check "R26-2_回答側は社内ナレッジ検索のヘッダー付き", _
        (InStr(oa, "社内ナレッジ検索") > 0 And Left$(oa, 1) = ChrW(&H3010)), oa
    modTestRunner.Check "R26-2_回答側は直前1件のみ(古い方は含まない)", _
        (InStr(oa, "A最新") > 0 And InStr(oa, "A古い") = 0), oa

    ' (vi) 一般→RAG: 逆方向はモード名が「一般アシスタント」で出ること。
    ok = modConvBridge.ComputeBridgeCore("normal", "rag", True, "Q", "A", "", "", 3, oq, oa)
    modTestRunner.Check "R26-2_一般からRAGへの橋渡しは成功", (ok = True), "ok=" & ok
    modTestRunner.Check "R26-2_回答側は一般アシスタントのヘッダー付き", _
        (InStr(oa, "一般アシスタント") > 0), oa

    ' (vii) R26H F3: 切替先に既存の記憶があるとき、それを【消さず】に先頭へ
    '       差し込む。ここが旧仕様(置換)との分かれ目で、反証条件は
    '       「既存の記憶が残っているか」。
    ok = modConvBridge.ComputeBridgeCore("rag", "normal", True, _
        "橋Q", "橋A", "既存Q1;;;既存Q2", "既存A1;;;既存A2", 3, oq, oa)
    modTestRunner.Check "R26H_F3_差し込みは成功する", (ok = True), "ok=" & ok
    modTestRunner.Check "R26H_F3_質問側は橋渡し分が先頭", (Left$(oq, 5) = "橋Q;;;"), oq
    modTestRunner.Check "R26H_F3_質問側は既存の記憶を消さない", _
        (InStr(oq, "既存Q1") > 0 And InStr(oq, "既存Q2") > 0), oq
    modTestRunner.Check "R26H_F3_回答側も既存の記憶を消さない", _
        (InStr(oa, "既存A1") > 0 And InStr(oa, "既存A2") > 0), oa
    modTestRunner.Check "R26H_F3_回答側の先頭は橋渡しのヘッダー付き1件", _
        (Left$(oa, 1) = ChrW(&H3010) And InStr(oa, "橋A") > 0), oa

    ' (viii) 上限で丸める。maxPairs=2 に既存2件へ1件差し込めば、いちばん古い
    '        1件が落ちて2件になる(modFollowup.KeepNewestPairs へ委譲)。
    ok = modConvBridge.ComputeBridgeCore("rag", "normal", True, _
        "橋Q", "橋A", "既存Q1;;;既存Q2", "既存A1;;;既存A2", 2, oq, oa)
    modTestRunner.Check "R26H_F3_上限2で3件目は落ちる", _
        (oq = "橋Q;;;既存Q1"), oq
    modTestRunner.Check "R26H_F3_上限で落ちるのは最も古い1件", _
        (InStr(oq, "既存Q2") = 0), oq

    ' (ix) 同じ往復が既に先頭にあるなら差し込まない(切替を往復させただけの
    '      とき、同じ内容が2件並ばない/トーストも出ない)。反証は「2回目の
    '      呼び出しでも oq が1回目と同じであること」ではなく【Falseを返す】こと。
    Dim oq2 As String, oa2 As String
    ok = modConvBridge.ComputeBridgeCore("rag", "normal", True, _
        "橋Q", "橋A", "", "", 3, oq2, oa2)
    modTestRunner.Check "R26H_F3_1回目は差し込む", (ok = True), "ok=" & ok
    ok = modConvBridge.ComputeBridgeCore("rag", "normal", True, _
        "橋Q", "橋A", oq2, oa2, 3, oq, oa)
    modTestRunner.Check "R26H_F3_同じ往復が先頭なら差し込まない", (ok = False), "ok=" & ok
    modTestRunner.Check "R26H_F3_差し込まないときのoutQ/outAは空", _
        (LenB(oq) = 0 And LenB(oa) = 0), "[" & oq & "][" & oa & "]"

    ' (x) 片方だけ一致は「同じ往復」ではない(回答が変わっていれば差し込む)。
    ok = modConvBridge.ComputeBridgeCore("rag", "normal", True, _
        "橋Q", "別のA", oq2, oa2, 3, oq, oa)
    modTestRunner.Check "R26H_F3_回答が違えば同じ往復と見なさない", (ok = True), "ok=" & ok
End Sub

' ----------------------------------------------------------------------------
' R26H F3: 差し込みの部品(InsertAtHead / AlreadyAtHead)を単体で固定する。
'   ComputeBridgeCore 経由だけだと、上限の丸めが KeepNewestPairs 側の
'   仕様変更で静かにずれても気付けない。
' ----------------------------------------------------------------------------
Private Sub TestConvBridgeInsert25()
    modTestRunner.Check "R26H_F3_既存が空なら差し込んだ1件だけ", _
        (modConvBridge.InsertAtHead("新", "", 3) = "新"), _
        modConvBridge.InsertAtHead("新", "", 3)
    modTestRunner.Check "R26H_F3_既存の先頭へ入る", _
        (modConvBridge.InsertAtHead("新", "旧1;;;旧2", 3) = "新;;;旧1;;;旧2"), _
        modConvBridge.InsertAtHead("新", "旧1;;;旧2", 3)
    modTestRunner.Check "R26H_F3_上限を超えた分は末尾から落ちる", _
        (modConvBridge.InsertAtHead("新", "旧1;;;旧2", 2) = "新;;;旧1"), _
        modConvBridge.InsertAtHead("新", "旧1;;;旧2", 2)
    ' maxPairs<=0(履歴を持たない設定)でも、押した操作そのものは無効にしない。
    modTestRunner.Check "R26H_F3_上限0以下でも差し込んだ1件は残る", _
        (modConvBridge.InsertAtHead("新", "旧1;;;旧2", 0) = "新"), _
        modConvBridge.InsertAtHead("新", "旧1;;;旧2", 0)

    modTestRunner.Check "R26H_F3_両方一致で既に先頭", _
        (modConvBridge.AlreadyAtHead("Q", "A", "Q;;;古Q", "A;;;古A") = True), ""
    modTestRunner.Check "R26H_F3_質問だけ一致は先頭扱いしない", _
        (modConvBridge.AlreadyAtHead("Q", "A", "Q;;;古Q", "別A;;;古A") = False), ""
    modTestRunner.Check "R26H_F3_切替先が空なら先頭一致ではない", _
        (modConvBridge.AlreadyAtHead("Q", "A", "", "") = False), ""
    ' 2件目以降に同じ往復があっても「先頭」ではない(差し込む)。
    modTestRunner.Check "R26H_F3_2件目に同じ往復があっても先頭扱いしない", _
        (modConvBridge.AlreadyAtHead("Q", "A", "別Q;;;Q", "別A;;;A") = False), ""
End Sub

' ----------------------------------------------------------------------------
' R26-3(1): modInsightCard — 資料名と既定題名の生成。
'   資料名の先頭が💭考察メモ_であることは、凍結制約下(modRetrieve/modPrompts
'   不触)で唯一成立する汚染防止の防御層そのもの。ここが緩むと、AIとの対話から
'   生まれた考察が一次資料と同じ顔で検索結果・出典に並ぶ。
' ----------------------------------------------------------------------------
Private Sub TestInsightDocName25()
    Dim memo As String: memo = ChrW(&HD83D) & ChrW(&HDCAD) & "考察メモ_"
    modTestRunner.Check "R26-3_資料名は考察メモ絵文字+考察メモ_題名", _
        (modInsightCard.DocNameFor("契約の免責") = memo & "契約の免責"), _
        modInsightCard.DocNameFor("契約の免責")
    ' 先頭2文字(サロゲートペア)が💭であること=本棚カード・出典表示に出る印。
    modTestRunner.Check "R26-3_資料名の先頭は考察メモ絵文字(サロゲートペア)", _
        (Left$(modInsightCard.DocNameFor("X"), 2) = ChrW(&HD83D) & ChrW(&HDCAD)), _
        modInsightCard.DocNameFor("X")
    ' 資料名の生成も禁止文字処理を必ず通る(題名を素通しさせない)。
    modTestRunner.Check "R26-3_資料名の題名も禁止文字が置換される", _
        (modInsightCard.DocNameFor("A/B") = memo & "A_B"), _
        modInsightCard.DocNameFor("A/B")

    ' 既定題名=質問の先頭24字(spec §3)。
    Dim q As String
    q = "約款の免責事由について教えてください。特に地震の扱いが知りたいです。"
    modTestRunner.Check "R26-3_既定題名は質問の先頭24字と一致", _
        (modInsightCard.TitleFromQuestion(q) = Left$(q, 24)), _
        modInsightCard.TitleFromQuestion(q)
    modTestRunner.Check "R26-3_既定題名の長さは24字", _
        (Len(modInsightCard.TitleFromQuestion(q)) = 24), _
        "len=" & Len(modInsightCard.TitleFromQuestion(q))
    modTestRunner.Check "R26-3_24字以下の質問はそのまま既定題名になる", _
        (modInsightCard.TitleFromQuestion("免責とは何ですか") = "免責とは何ですか"), _
        modInsightCard.TitleFromQuestion("免責とは何ですか")
    ' 改行・タブは空白へ潰す(InputBoxの既定値が複数行になると編集できない)。
    modTestRunner.Check "R26-3_既定題名の改行は空白へ潰す", _
        (modInsightCard.TitleFromQuestion("免責とは" & vbLf & "何ですか") = "免責とは 何ですか"), _
        modInsightCard.TitleFromQuestion("免責とは" & vbLf & "何ですか")
End Sub

' ----------------------------------------------------------------------------
' R26-3(2): modInsightCard.SafeTitle — 題名の禁止文字処理。
'   題名はそのまま一時ファイル名になり(modVault.RegisterKnowledgeText の
'   docBase)、そのファイル名が本棚の資料名になる。禁止文字が1つ残るだけで
'   書き出しが失敗し、保存が丸ごと落ちる。
' ----------------------------------------------------------------------------
Private Sub TestInsightSafeTitle25()
    modTestRunner.Check "R26-3_ファイル名禁止文字は全て_へ", _
        (modInsightCard.SafeTitle("A\B/C:D*E?F""G<H>I|J") = "A_B_C_D_E_F_G_H_I_J"), _
        modInsightCard.SafeTitle("A\B/C:D*E?F""G<H>I|J")
    modTestRunner.Check "R26-3_タブ・改行も_へ", _
        (modInsightCard.SafeTitle("A" & vbTab & "B" & vbLf & "C") = "A_B_C"), _
        modInsightCard.SafeTitle("A" & vbTab & "B" & vbLf & "C")
    ' Windowsは末尾がドット/空白のファイルを作れない(作れても開けない)。
    modTestRunner.Check "R26-3_末尾のドットと空白は落とす", _
        (modInsightCard.SafeTitle("  メモ. . ") = "メモ"), _
        "[" & modInsightCard.SafeTitle("  メモ. . ") & "]"
    modTestRunner.Check "R26-3_ドットだけの題名は空になる(呼び出し側が断る)", _
        (LenB(modInsightCard.SafeTitle("...")) = 0), _
        "[" & modInsightCard.SafeTitle("...") & "]"
    ' 40字上限(接頭辞7字を足しても modVault.SanitizeName の60字で切られない)。
    Dim long50 As String: long50 = String$(50, "あ")
    modTestRunner.Check "R26-3_題名は40字で切る", _
        (Len(modInsightCard.SafeTitle(long50)) = 40), _
        "len=" & Len(modInsightCard.SafeTitle(long50))
    modTestRunner.Check "R26-3_40字以下の題名は切らない", _
        (modInsightCard.SafeTitle("免責事由") = "免責事由"), _
        modInsightCard.SafeTitle("免責事由")
End Sub

' ----------------------------------------------------------------------------
' R26-3(3): modInsightCard — 汚染防止の定型文(冪等)と本文の組み立て。
'   同じ会話を2回保存しても定型文が積み重ならないこと=modConvBridge の
'   出所ヘッダーと同じ性質。積み重なると、本文の頭が注意書きで埋まって
'   1チャンク目の中身が痩せる。
' ----------------------------------------------------------------------------
Private Sub TestInsightNotice25()
    Dim notice As String: notice = modInsightCard.MemoNotice()
    modTestRunner.Check "R26-3_定型文は一次資料ではないと明言する", _
        (InStr(notice, "一次資料ではありません") > 0 And InStr(notice, "出典として扱わず") > 0), _
        notice

    Dim once As String: once = modInsightCard.WithMemoNotice("本文")
    modTestRunner.Check "R26-3_定型文は本文の冒頭に付く", _
        (Left$(once, Len(notice)) = notice And InStr(once, "本文") > 0), once
    modTestRunner.Check "R26-3_定型文の付与は冪等(2回でも1回ぶん)", _
        (modInsightCard.WithMemoNotice(once) = once), _
        "len=" & Len(modInsightCard.WithMemoNotice(once)) & "/" & Len(once)

    Dim body As String: body = modInsightCard.MemoBody("免責とは?", "免責は…です。")
    modTestRunner.Check "R26-3_本文は定型文で始まる", _
        (Left$(body, Len(notice)) = notice), Left$(body, 20)
    modTestRunner.Check "R26-3_本文に質問と回答の両方が入る", _
        (InStr(body, "免責とは?") > 0 And InStr(body, "免責は…です。") > 0), body
    modTestRunner.Check "R26-3_本文の見出しは質問と回答を分ける", _
        (InStr(body, "■ 質問") > 0 And InStr(body, "■ AIの回答") > 0), body
    ' 二重保存でも定型文は1つ(MemoBodyの結果をもう一度包んでも増えない)。
    modTestRunner.Check "R26-3_本文を再度包んでも定型文は増えない", _
        (modInsightCard.WithMemoNotice(body) = body), _
        "len=" & Len(modInsightCard.WithMemoNotice(body)) & "/" & Len(body)
End Sub

' ============================================================================
Public Sub RunAll25()
    On Error GoTo SwapFail25
    TestDiversitySwapPick25
NextVerdict25:
    On Error GoTo VerdictFail25
    TestGenParseVerdict25
NextLoopBound25:
    On Error GoTo LoopBoundFail25
    TestGenVerifyLoopBound25
NextPlan25:
    On Error GoTo PlanFail25
    TestGenPlanTable25
NextBridgeTrunc25:
    On Error GoTo BridgeTruncFail25
    TestConvBridgeTruncate25
NextBridgeHeader25:
    On Error GoTo BridgeHeaderFail25
    TestConvBridgeHeader25
NextBridgeCore25:
    On Error GoTo BridgeCoreFail25
    TestConvBridgeCore25
NextBridgeInsert25:
    On Error GoTo BridgeInsertFail25
    TestConvBridgeInsert25
NextInsightName25:
    On Error GoTo InsightNameFail25
    TestInsightDocName25
NextInsightTitle25:
    On Error GoTo InsightTitleFail25
    TestInsightSafeTitle25
NextInsightNotice25:
    On Error GoTo InsightNoticeFail25
    TestInsightNotice25
NextDone25:
    On Error GoTo 0
    Exit Sub

SwapFail25:
    modTestRunner.Check "TestDiversitySwapPick25(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextVerdict25
VerdictFail25:
    modTestRunner.Check "TestGenParseVerdict25(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextLoopBound25
LoopBoundFail25:
    modTestRunner.Check "TestGenVerifyLoopBound25(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextPlan25
PlanFail25:
    modTestRunner.Check "TestGenPlanTable25(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBridgeTrunc25
BridgeTruncFail25:
    modTestRunner.Check "TestConvBridgeTruncate25(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBridgeHeader25
BridgeHeaderFail25:
    modTestRunner.Check "TestConvBridgeHeader25(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBridgeCore25
BridgeCoreFail25:
    modTestRunner.Check "TestConvBridgeCore25(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBridgeInsert25
BridgeInsertFail25:
    modTestRunner.Check "TestConvBridgeInsert25(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextInsightName25
InsightNameFail25:
    modTestRunner.Check "TestInsightDocName25(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextInsightTitle25
InsightTitleFail25:
    modTestRunner.Check "TestInsightSafeTitle25(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextInsightNotice25
InsightNoticeFail25:
    modTestRunner.Check "TestInsightNotice25(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone25
End Sub
