Attribute VB_Name = "modTestsPure40"
Option Explicit

' ============================================================================
' modTestsPure40 - R36 波2(是正の仕組み)の純ロジック回帰テスト。
'   既存チェーン(modTestsPure.RunAll→…)へは繋がず、
'   modTestRunner.RunAllPureTests から直接呼ばれる RunAll40 の1本が入口
'   (modTestsPure31〜38 と同型の別枝)。
' ----------------------------------------------------------------------------
' 【何を固定するか】
'   C1 modCorrect.BuildMemoBody: 是正メモの本文の書式そのもの。1行目の
'     「【是正メモ】この質問には、下の「正しい内容」で答えること。」が
'     【本文側にも】載ること ―― この1行がそのまま LLM への指示として働く
'     ので、資料名側にだけ置くと注入しても何も起きない。質問が「質問: 」行
'     として入り、ExtractQuestionLine で往復できること。
'   C2 modCorrect.ExtractQuestionLine: 取込時に full_text の先頭へ焼かれる
'     breadcrumb 行【資料名>章>条】(modShelf.bas:615-630)があっても質問行を
'     拾えること。vbCrLf 改行。質問行が無い本文・空文字では空を返すこと
'     (=是正メモとして扱わない=注入しない)。
'   C3 modCorrect.MatchLevel: 2=完全一致 / 1=語一致 / 0=不一致 の3値。
'     完全一致は modInsight.NormKey(空白・句読点・?・括弧を落として先頭40字を
'     小文字化)を物差しにするので、「?」や空白の有無で揺れない。
'   C4 modCorrect.KeyMatchPct: 語一致率の算数。100%/0% の両端で固定する。
'     ここを両端で押さえておくと、C3 の語一致分岐が config correct_key_min
'     (既定60)の値に依存せず判定できる ―― 閾値をいくつにしても 100 は通り
'     0 は通らないため、テストが config の実行環境差で揺れない。
'   C5 modCorrect.MemoDocBase: 接頭辞と24字クランプ、そして
'     modVault.SanitizeName の置換表を【先に】通してある冪等性。
'     modAppAct.RecordCorrection は「同名の既存資料を消してから登録し直す」
'     ために source 名(= MemoDocBase & ".txt")を自力で組み立てる。冪等性が
'     崩れると「?」を含む質問で名前が食い違い、消し忘れ=是正メモが本棚に
'     増え続ける(R36 §2 が止めようとしている当のもの)。
'
' 【なぜ InjectHits / PrependHit のテストが無いか】
'   どちらも Hit(Public Type)の配列を跨ぐ。LO ではモジュール間で UDT 配列を
'   受け渡せず、書いても [SKIP] に落ちるだけで SKIP 上限14を押し上げる
'   (run_lo_tests.py:129-132 / modTestsPure38.bas:27 と同型の既知の死角)。
'   注入の実挙動は R36 §6 の実機受入③(❌違う→正しい内容→同じ質問で是正メモが
'   出典の先頭に出る)で確かめる。
' ============================================================================

' ---- C1: BuildMemoBody の書式 ----------------------------------------------
'   discriminate:
'   ・1行目の指示文を落とす(質問と正しい内容だけを書く)実装にすると
'     「指示文が本文の先頭にある」の2本が落ちる。
'   ・質問行の頭を "質問:" 以外(例 "Q: ")にすると往復の1本が落ちる。
'   ・正しい内容・誤答抜粋の行を落とすと、それぞれの InStr 判定が落ちる。
'   ・行区切りを vbLf 以外にすると ExtractQuestionLine 往復が落ちる
'     (ExtractQuestionLine は vbCr を落としてから vbLf で割るため、
'      vbCrLf にしても通る=そこは意図した許容範囲)。
Private Sub TestBuildMemoBody40()
    Dim body As String
    body = modCorrect.BuildMemoBody("退職金の計算方法は?", "勤続年数×基本給×0.6です。", "退職金は一律50万円です")

    ' (a) 指示文が【本文の先頭】にある。ここが LLM への指示そのもの。
    ChkBool40 "C1_本文の先頭が是正メモの指示文", _
        (InStr(1, body, "【是正メモ】この質問には、下の「正しい内容」で答えること。", vbBinaryCompare) = 1), True

    ' (b) 質問・正しい内容・誤答抜粋の3つが全部載る。
    ChkBool40 "C1_質問行が載る", (InStr(1, body, vbLf & "質問: 退職金の計算方法は?", vbBinaryCompare) > 0), True
    ChkBool40 "C1_正しい内容が載る", (InStr(1, body, "正しい内容: 勤続年数×基本給×0.6です。", vbBinaryCompare) > 0), True
    ChkBool40 "C1_誤答の抜粋が載る", (InStr(1, body, "退職金は一律50万円です", vbBinaryCompare) > 0), True

    ' (c) 書いた質問を ExtractQuestionLine がそのまま取り戻せる(往復)。
    ChkStr40 "C1_質問の往復", modCorrect.ExtractQuestionLine(body), "退職金の計算方法は?"

    ' (d) 空入力でも書式は壊れない(質問が空なら往復も空)。
    Dim empt As String: empt = modCorrect.BuildMemoBody("", "", "")
    ChkBool40 "C1_空入力でも指示文は残る", _
        (InStr(1, empt, "【是正メモ】", vbBinaryCompare) = 1), True
    ChkStr40 "C1_空質問の往復は空", modCorrect.ExtractQuestionLine(empt), ""
End Sub

' ---- C2: ExtractQuestionLine ------------------------------------------------
'   discriminate:
'   ・行単位ではなく本文全体へ InStr する実装にすると、(b)の breadcrumb の
'     中に「質問:」を含む資料名が来たとき壊れる。ここでは行頭一致であることを
'     breadcrumb 付きケースで押さえる。
'   ・vbCr を落とさない実装にすると(c)の戻り値に vbCr が残って落ちる。
'   ・"質問:" が無いときに1行目や空でない何かを返す実装にすると(d)(e)が落ちる。
Private Sub TestExtractQuestionLine40()
    ' (a) 素直な本文。
    Dim plain As String
    plain = "【是正メモ】この質問には、下の「正しい内容」で答えること。" & vbLf & _
            "質問: 育休の申請期限は" & vbLf & "正しい内容: 1か月前です"
    ChkStr40 "C2_素直な本文", modCorrect.ExtractQuestionLine(plain), "育休の申請期限は"

    ' (b) 取込後の full_text は先頭に breadcrumb 行が焼かれている。
    '     breadcrumb の【中に】"質問:" を仕込んであるので、行頭一致ではなく
    '     本文全体への InStr で実装するとここで "ダミー】" を返して落ちる。
    Dim crumbed As String
    crumbed = "【是正メモ_育休>質問:ダミー】" & vbLf & plain
    ChkStr40 "C2_breadcrumb付きでも拾える", modCorrect.ExtractQuestionLine(crumbed), "育休の申請期限は"

    ' (c) vbCrLf 改行(TEMP へ書いた txt の読み戻しで起きうる)。
    Dim crlf As String
    crlf = "【是正メモ】指示" & vbCrLf & "質問: 経費精算の締切" & vbCrLf & "正しい内容: 月末"
    ChkStr40 "C2_vbCrLf改行でも拾える", modCorrect.ExtractQuestionLine(crlf), "経費精算の締切"

    ' (d) 質問行が無い本文 → 空(是正メモとして扱わない)。
    ChkStr40 "C2_質問行が無ければ空", _
        modCorrect.ExtractQuestionLine("【修正ナレッジ】" & vbLf & "正しい内容: なにか"), ""

    ' (e) 空文字 → 空。
    ChkStr40 "C2_空文字は空", modCorrect.ExtractQuestionLine(""), ""
End Sub

' ---- C3: MatchLevel の3値 ---------------------------------------------------
'   discriminate:
'   ・完全一致の判定を NormKey ではなく生文字列の比較にすると、(b)の
'     「空白と?の有無だけが違う質問」が 2 を返さず落ちる。
'   ・質問行の無い本文に対して 1 以上を返す実装(本文全体を語一致の照合先に
'     する等)にすると(e)が落ちる。
'   ・空クエリを弾かない実装にすると(f)が落ちる。
'   ・語一致の照合先を「是正メモの本文全体」にすると、(d)の無関係な質問が
'     「正しい内容」側の語に当たって 1 を返し落ちうる(照合先は質問行だけ)。
Private Sub TestMatchLevel40()
    Dim memo As String
    ' 「正しい内容」の側に、(d)で使う質問の効く語(育児休業・申請期限)を
    ' わざと両方入れてある。語一致の照合先を【質問行だけ】ではなく本文全体に
    ' した実装だと、(d)が一致率100%になって 1 を返し、そこで落ちる。
    memo = modCorrect.BuildMemoBody("退職金の計算方法は?", _
                                    "勤続年数と基本給から算定します。育児休業の申請期限とは無関係です。", _
                                    "退職金は一律50万円です")

    ' (a) そのままの質問 → 完全一致。
    ChkLong40 "C3_同じ質問は完全一致2", _
        modCorrect.MatchLevel(modSparse.NormalizeForSearch("退職金の計算方法は?"), memo), 2

    ' (b) 空白と「?」の有無だけが違う → NormKey が同じなので完全一致のまま。
    ChkLong40 "C3_空白と?の揺れは完全一致のまま2", _
        modCorrect.MatchLevel(modSparse.NormalizeForSearch("退職金の 計算方法は"), memo), 2

    ' (c) 末尾の「は?」だけを削った言い回し → NormKey は一致しない(2ではない)が、
    '     効く語「退職金」「計算方法」が両方とも是正メモの質問行にあるので
    '     一致率100% → 語一致。100% は閾値(config correct_key_min)を
    '     いくつにしても通るので、この期待値は設定値に依存しない。
    ChkLong40 "C3_語が全部重なる言い回しは語一致1", _
        modCorrect.MatchLevel(modSparse.NormalizeForSearch("退職金の計算方法"), memo), 1

    ' (d) 無関係な質問 → 不一致。是正メモの「正しい内容」側に「育児休業」が
    '     入れてあるので、照合先を本文全体にした実装はここで落ちる。
    ChkLong40 "C3_無関係な質問は0", _
        modCorrect.MatchLevel(modSparse.NormalizeForSearch("育児休業の申請期限は?"), memo), 0

    ' (e) 質問行の無い本文(旧「修正ナレッジ」など)は是正メモではない → 0。
    ChkLong40 "C3_質問行の無い本文は0", _
        modCorrect.MatchLevel(modSparse.NormalizeForSearch("退職金の計算方法は?"), _
                              "【修正ナレッジ】" & vbLf & "正しい内容: 退職金の計算方法"), 0

    ' (f) 空の質問では何とも一致しない(全是正メモが先頭へ来る事故を防ぐ)。
    ChkLong40 "C3_空クエリは0", modCorrect.MatchLevel("", memo), 0
End Sub

' ---- C4: KeyMatchPct の両端 -------------------------------------------------
'   discriminate:
'   ・分母を「是正メモ側の語数」にすると(a)が100を返さなくなる
'     (分母は【質問側】の効く語の数)。
'   ・部分一致ではなく完全一致で数える実装にすると(a)が落ちる
'     (「退職金」は「退職金の計算方法は」の中に部分一致で入っている)。
'   ・照合先の空白を落とさない実装だと、空白入りの質問行で取りこぼす(b)。
'   ・語が1つも重ならないのに 0 より大きい値を返す実装は(c)で落ちる。
Private Sub TestKeyMatchPct40()
    ' (a) 質問の効く語(DistinctiveKeys は「退職金」「計算方法」の2語に割れる。
    '     ひらがな「の」が語の区切りになるため)がすべて是正メモの質問行に
    '     部分一致で現れる → 100%。
    ChkLong40 "C4_全語一致は100", _
        modCorrect.KeyMatchPct(modSparse.NormalizeForSearch("退職金の計算方法"), _
                               modSparse.NormalizeForSearch("退職金の計算方法は?")), 100

    ' (b) 照合先に空白が混ざっていても落とすので 100% のまま。
    ChkLong40 "C4_照合先の空白は無視して100", _
        modCorrect.KeyMatchPct(modSparse.NormalizeForSearch("退職金の計算方法"), _
                               modSparse.NormalizeForSearch("退職金 の 計算方法 は")), 100

    ' (c) 語が1つも重ならない → 0%。
    ChkLong40 "C4_無関係は0", _
        modCorrect.KeyMatchPct(modSparse.NormalizeForSearch("育児休業 申請期限"), _
                               modSparse.NormalizeForSearch("退職金の計算方法は?")), 0

    ' (d) 照合先が空 → 0%(0除算も例外も起こさない)。
    ChkLong40 "C4_照合先が空なら0", _
        modCorrect.KeyMatchPct(modSparse.NormalizeForSearch("退職金の計算方法"), ""), 0

    ' (e) 質問側から効く語が1つも取れない → 0%。
    ChkLong40 "C4_効く語が無ければ0", _
        modCorrect.KeyMatchPct("", modSparse.NormalizeForSearch("退職金の計算方法は?")), 0
End Sub

' ---- C5: MemoDocBase の接頭辞・クランプ・冪等性 ------------------------------
'   discriminate:
'   ・接頭辞を変える/落とすと(a)が落ち、InjectHits の走査(source が
'     "是正メモ_" で始まる行だけを見る)も同時に空振りするので、ここが
'     接頭辞の単一情報源の見張りになる。
'   ・24字クランプを外すと(b)が落ちる。
'   ・置換を modVault.SanitizeName 任せにして自分では置換しない実装にすると
'     (c)(d)が落ちる ―― これが崩れると source 名が食い違い、
'     modAppAct.RecordCorrection の DeleteSource が空振りして是正メモが
'     本棚に増え続ける(旧「修正ナレッジ」と同じ壊れ方に戻る)。
Private Sub TestMemoDocBase40()
    ' (a) 接頭辞。
    ChkStr40 "C5_接頭辞と本体", modCorrect.MemoDocBase("退職金の計算方法"), "是正メモ_退職金の計算方法"

    ' (b) 24字でクランプ。"あ"×30 → 先頭24字だけ。
    Dim long30 As String: long30 = String$(30, ChrW(&H3042))
    ChkStr40 "C5_24字クランプ", modCorrect.MemoDocBase(long30), "是正メモ_" & String$(24, ChrW(&H3042))

    ' (c) SanitizeName の禁止文字は【この関数の中で】潰す。日本語の質問では
    '     「?」が普通に出るので、ここが最も踏まれる経路。
    ChkStr40 "C5_疑問符は先に潰す", modCorrect.MemoDocBase("退職金は?"), "是正メモ_退職金は_"
    ChkStr40 "C5_パス区切りも潰す", modCorrect.MemoDocBase("A/B:C*D"), "是正メモ_A_B_C_D"

    ' (d) 冪等性: 出力をもう一度通しても値が変わらない(= modVault.SanitizeName を
    '     通しても変わらない、の代理検査)。ここが崩れると DeleteSource が空振りする。
    Dim once As String: once = modCorrect.MemoDocBase("退職金は?")
    ChkStr40 "C5_置換は冪等", modCorrect.MemoDocBase(Mid$(once, Len("是正メモ_") + 1)), once

    ' (e) 空の質問でも接頭辞だけは残る(名前が空文字にならない)。
    ChkStr40 "C5_空質問でも接頭辞は残る", modCorrect.MemoDocBase(""), "是正メモ_"
End Sub

' ---- 判定ヘルパー -----------------------------------------------------------
Private Sub ChkBool40(ByVal label As String, ByVal got As Boolean, ByVal want As Boolean)
    modTestRunner.Check "R36-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkLong40(ByVal label As String, ByVal got As Long, ByVal want As Long)
    modTestRunner.Check "R36-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkStr40(ByVal label As String, ByVal got As String, ByVal want As String)
    modTestRunner.Check "R36-" & label, (StrComp(got, want, vbBinaryCompare) = 0), _
        "実際=[" & got & "] 期待=[" & want & "]"
End Sub

Public Sub RunAll40()
    On Error GoTo H01Fail40
    TestBuildMemoBody40
H02Next40:
    On Error GoTo H02Fail40
    TestExtractQuestionLine40
H03Next40:
    On Error GoTo H03Fail40
    TestMatchLevel40
H04Next40:
    On Error GoTo H04Fail40
    TestKeyMatchPct40
H05Next40:
    On Error GoTo H05Fail40
    TestMemoDocBase40
H01Done40:
    On Error GoTo 0
    Exit Sub

H01Fail40:
    modTestRunner.Check "TestBuildMemoBody40(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H02Next40
H02Fail40:
    modTestRunner.Check "TestExtractQuestionLine40(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H03Next40
H03Fail40:
    modTestRunner.Check "TestMatchLevel40(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H04Next40
H04Fail40:
    modTestRunner.Check "TestKeyMatchPct40(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H05Next40
H05Fail40:
    modTestRunner.Check "TestMemoDocBase40(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done40
End Sub
