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
'   C5 modCorrect.MemoDocBase: 接頭辞・24字クランプ・末尾4桁、そして
'     modVault.SanitizeName の置換表を【先に】通してある冪等性。
'     modAppAct.RecordCorrection は「同名の既存資料を消してから登録し直す」
'     ために source 名(= MemoDocBase & ".txt")を自力で組み立てる。冪等性が
'     崩れると「?」を含む質問で名前が食い違い、消し忘れ=是正メモが本棚に
'     増え続ける(R36 §2 が止めようとしている当のもの)。
'     R36 Fix M2/A-M7: 先頭24字が同じ別の質問が同名になり、後の是正が前の
'     是正を DeleteSource で消していた。末尾4桁で分ける。
'     R36 Fix A-r5: 冪等性の検査を「SanitizeName の置換表を再現して掛ける」
'     形へ書き換えた(旧 (d) は出力を再投入するだけで実質恒真だった)。
'   C6 modCorrect.QuestionTag4: 資料名の末尾4桁。手計算できる期待値で固定し、
'     AscW の負値補正(U+8000以降)とゼロ埋めを押さえる。
'   C7 modCorrect.MatchLevel の完全一致は40字で切らない(R36 Fix A-M6)。
'     先頭40字が同じで41字目以降が違う別の質問が score 1.0 で先頭へ入る
'     経路を、正例(同一質問=2)と反例(41字目違い=2ではない)の2本で挟む。
'   C8 modCorrect.BuildMemoBody の4引数書式(R37 §2-2): 第4引数
'     sourcesLine が非空のとき「誤答の根拠: …」行が末尾に載り
'     ExtractWrongSources で往復できること。空なら行そのものが出ないこと
'     (既存3引数呼び出しと同じ書式を保つ)。
'   C9 modCorrect.ExtractWrongSources: 「誤答の根拠:」行を行単位で拾えること
'     (有・複数・無・空の4パターン)。
'   C10 modCorrect.WrongSourceMatches: (source,page)一致判定の純関数。
'     資料名は大小無視・ページは数値一致・"|"区切りの2件目以降も拾えること。
'
' 【なぜ InjectHits / PrependHit / DemoteWrongHits のテストが無いか】
'   いずれも Hit(Public Type)の配列を跨ぐ。LO ではモジュール間で UDT 配列を
'   受け渡せず、書いても [SKIP] に落ちるだけで SKIP 上限14を押し上げる
'   (run_lo_tests.py:129-132 / modTestsPure38.bas:27 と同型の既知の死角)。
'   注入・降格の実挙動は R36 §6 の実機受入③(❌違う→正しい内容→同じ質問で
'   是正メモが出典の先頭に出る)と R37 §2-1(誤根拠が末尾へ下がる)で確かめる。
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
    body = modCorrect.BuildMemoBody("退職金の計算方法は?", "勤続年数×基本給×0.6です。", "退職金は一律50万円です", "")

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
    Dim empt As String: empt = modCorrect.BuildMemoBody("", "", "", "")
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
                                    "退職金は一律50万円です", "")

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

    ' (c) 語が1つも重ならない → 0%。効く語は2語(育児休業/申請期限)取れるので、
    '     (f)の total<2 ルールではなく【重なりが無いこと】で 0 になる
    '     (「育児休業 申請期限」と空白で書くと DistinctiveKeys が空白を落として
    '      漢字ラン1本=1語になり、判別力の無いテストになる。助詞で割る)。
    ChkLong40 "C4_無関係は0", _
        modCorrect.KeyMatchPct(modSparse.NormalizeForSearch("育児休業の申請期限"), _
                               modSparse.NormalizeForSearch("退職金の計算方法は?")), 0

    ' (d) 照合先が空 → 0%(0除算も例外も起こさない)。
    ChkLong40 "C4_照合先が空なら0", _
        modCorrect.KeyMatchPct(modSparse.NormalizeForSearch("退職金の計算方法"), ""), 0

    ' (e) 質問側から効く語が1つも取れない → 0%。
    ChkLong40 "C4_効く語が無ければ0", _
        modCorrect.KeyMatchPct("", modSparse.NormalizeForSearch("退職金の計算方法は?")), 0

    ' (f) R36 Fix N6: 効く語が【1語しか】取れない質問は語一致に使わない。
    '     「有給」は漢字ラン1本=キー1語なので、従来はその1語が当たるだけで
    '     一致率100%になり、「有給の繰越は?」にも「有給の申請先は?」にも
    '     同じ是正メモが先頭で刺さっていた(誤爆の主要経路)。分母が2未満なら 0。
    ChkLong40 "C4_効く語が1語なら0(N6・誤爆防止)", _
        modCorrect.KeyMatchPct(modSparse.NormalizeForSearch("有給"), _
                               modSparse.NormalizeForSearch("有給の繰越はいつまでですか")), 0

    ' (g) (f)の対照。照合先は同じで、質問側の効く語が2語(有給/繰越)になれば
    '     従来どおり率を返す。N6 が語一致そのものを殺していないことの確認
    '     ―― この2本は「1語か2語か」だけが違うので、total<2 の判定を
    '     外すと(f)が100に化けて落ち、判定を強くしすぎると(g)が0に化けて落ちる。
    ChkLong40 "C4_効く語が2語なら従来どおり100", _
        modCorrect.KeyMatchPct(modSparse.NormalizeForSearch("有給の繰越"), _
                               modSparse.NormalizeForSearch("有給の繰越はいつまでですか")), 100
End Sub

' ---- C6: QuestionTag4(資料名の4桁識別子) ------------------------------------
'   R36 Fix M2/A-M7 の要。期待値は手計算できる形で固定してある:
'   ・"A/B:C*D" は NormKey で小文字化されるだけ(除去対象の文字を含まない)。
'     a97 + /47 + b98 + :58 + c99 + *42 + d100 = 541 → "0541"。
'   ・"退職金は?" は「?」が落ちて 退(U+9000=36864) + 職(U+8077=32887) +
'     金(U+91D1=37329) + は(U+306F=12399) = 119479 → mod 10000 = 9479。
'     退・職・金はいずれも U+8000 以降で AscW が【負値】を返すため、
'     +65536 の補正(CLAUDE.md §10)を外すとこの1本が落ちる。
'   discriminate:
'   ・+65536 補正を外す → (b)が落ちる((a)は ASCII だけなので落ちない)。
'   ・NormKey を通さず生文字列で足す → (c)が落ちる(「?」の有無で値が変わる)。
'   ・4桁ゼロ埋めを忘れる → (d)が "0" になって落ちる。
Private Sub TestQuestionTag4_40()
    ChkStr40 "C6_ASCIIの和", modCorrect.QuestionTag4("A/B:C*D"), "0541"
    ChkStr40 "C6_U+8000以降を含む和(負値補正)", modCorrect.QuestionTag4("退職金は?"), "9479"
    ChkStr40 "C6_疑問符の有無で値が変わらない", _
        modCorrect.QuestionTag4("退職金は"), modCorrect.QuestionTag4("退職金は?")
    ChkStr40 "C6_空はゼロ埋め4桁", modCorrect.QuestionTag4(""), "0000"
End Sub

' ---- C7: MatchLevel の完全一致は40字で切らない(R36 Fix A-M6) ----------------
'   従来 level 2 の物差しは modInsight.NormKey(先頭40字クランプ)だった。
'   41字目以降だけが違う別の質問が「完全一致」と判定され、score 1.0 で
'   出典の先頭に差し込まれる(gap の重複判定は「似た質問をまとめる」のが
'   目的なので40字で正しいが、是正の完全一致は同じ質問でしか成立しない)。
'   discriminate:
'   ・NormFull を NormKey へ戻す(=40字クランプを復活させる)と(b)が 2 を
'     返して落ちる。(a)は戻しても通るので、2本で挟んで初めて効く。
'   ・NormFull から記号除去を落とすと(c)が落ちる。
Private Sub TestMatchLevel40Boundary()
    Dim head40 As String: head40 = String$(40, ChrW(&H3042))   ' 「あ」×40=クランプ長ちょうど
    Dim qA As String: qA = head40 & "退職金の計算方法"
    Dim qB As String: qB = head40 & "育児休業の申請期限"

    Dim memo As String
    memo = modCorrect.BuildMemoBody(qA, "勤続年数×基本給×0.6です。", "誤答", "")

    ' (a) 同じ質問はもちろん完全一致。
    ChkLong40 "C7_41字目まで同じ同一質問は2", _
        modCorrect.MatchLevel(modSparse.NormalizeForSearch(qA), memo), 2

    ' (b) 先頭40字が同じで、41字目以降が違う別の質問 → 完全一致ではない。
    '     語一致も成立しない(「育児休業」「申請期限」は是正メモの質問行に無い)。
    ChkLong40 "C7_先頭40字だけ同じ別質問は2ではない", _
        modCorrect.MatchLevel(modSparse.NormalizeForSearch(qB), memo), 0

    ' (c) 40字クランプは無くしたが、記号・空白の揺れは従来どおり吸収する。
    ChkLong40 "C7_記号と空白の揺れは吸収したまま2", _
        modCorrect.MatchLevel(modSparse.NormalizeForSearch(qA & "?"), memo), 2
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
    ' (a) 接頭辞 + 本体 + "_" + 4桁。4桁の値そのものは C6 が固定するので、
    '     ここは【組み立ての形】を見る(2箇所で同じ数字を書き写さない)。
    ChkStr40 "C5_接頭辞と本体と4桁", modCorrect.MemoDocBase("退職金の計算方法"), _
        "是正メモ_退職金の計算方法_" & modCorrect.QuestionTag4("退職金の計算方法")
    ' 4桁は 0355(C6 と同じ手計算の物差し)。組み立てと値の両方を1本で押さえる。
    ChkStr40 "C5_実値(退職金の計算方法)", modCorrect.MemoDocBase("退職金の計算方法"), _
        "是正メモ_退職金の計算方法_0355"

    ' (b) 24字でクランプ。"あ"×30 → 本体は先頭24字だけ。4桁は【クランプ前の
    '     質問全体】(正確には NormKey の40字)から作るので、24字を超える差も
    '     名前へ反映される ―― これが(f)の衝突回避の仕組み。
    Dim long30 As String: long30 = String$(30, ChrW(&H3042))
    ChkStr40 "C5_24字クランプ", modCorrect.MemoDocBase(long30), _
        "是正メモ_" & String$(24, ChrW(&H3042)) & "_0620"

    ' (c) SanitizeName の禁止文字は【この関数の中で】潰す。日本語の質問では
    '     「?」が普通に出るので、ここが最も踏まれる経路。
    ChkStr40 "C5_疑問符は先に潰す", modCorrect.MemoDocBase("退職金は?"), "是正メモ_退職金は__9479"
    ChkStr40 "C5_パス区切りも潰す", modCorrect.MemoDocBase("A/B:C*D"), "是正メモ_A_B_C_D_0541"

    ' (d) R36 Fix A-r5: 冪等性を【modVault.SanitizeName の置換表を再現して】
    '     検査する。旧テスト(出力の本体部分をもう一度 MemoDocBase に通す)は
    '     4桁が付いた時点で成り立たず、そもそも「SanitizeName を通しても
    '     変わらない」を測っていなかった(実質恒真)。
    '     ここが崩れると source 名が食い違い、modAppAct.RecordCorrection の
    '     DeleteSource が空振りして是正メモが本棚に増え続ける。
    '     禁止文字(? / :)を全部含む質問で見る。
    Dim qBad As String: qBad = "有給休暇の申請は?いつまで/どこへ:提出しますか"
    Dim baseName As String: baseName = modCorrect.MemoDocBase(qBad)
    ChkStr40 "C5_SanitizeNameを通しても変わらない(冪等)", SanitizeLikeVault40(baseName), baseName
    ' SanitizeName は60字で切る。切られると冪等が崩れるので、長さも押さえる
    ' (接頭辞5 + 本体24 + "_" + 4桁 = 最大34字。上限まで26字の余裕がある)。
    ChkBool40 "C5_資料名は60字以内(SanitizeNameのクランプに掛からない)", _
        (Len(baseName) <= 60), True
    ' 禁止文字が1つも残っていない(置換漏れの直接検査)。
    ChkBool40 "C5_禁止文字が残らない", _
        (InStr(1, baseName, "?", vbBinaryCompare) = 0 And _
         InStr(1, baseName, "/", vbBinaryCompare) = 0 And _
         InStr(1, baseName, ":", vbBinaryCompare) = 0), True

    ' (e) 空の質問でも接頭辞だけは残る(名前が空文字にならない)。
    ChkStr40 "C5_空質問でも接頭辞は残る", modCorrect.MemoDocBase(""), "是正メモ__0000"

    ' (f) R36 Fix M2/A-M7: 先頭24字が同じ【別の質問】は、別の資料名になる。
    '     旧実装は24字だけで名前を決めていたので、後から書いた是正メモが
    '     DeleteSource で前のものを消していた(2つ目の是正が1つ目を殺す)。
    '     4桁を落とす/24字だけに戻すと、この1本が必ず落ちる。
    Dim q24 As String: q24 = String$(24, ChrW(&H3042))
    ChkStr40 "C5_先頭24字が同じ別質問A", modCorrect.MemoDocBase(q24 & "退職金"), _
        "是正メモ_" & q24 & "_3576"
    ChkStr40 "C5_先頭24字が同じ別質問B", modCorrect.MemoDocBase(q24 & "育児休業"), _
        "是正メモ_" & q24 & "_7488"
    ChkBool40 "C5_先頭24字が同じ別質問は資料名が衝突しない", _
        (modCorrect.MemoDocBase(q24 & "退職金") <> modCorrect.MemoDocBase(q24 & "育児休業")), True
End Sub

' modVault.SanitizeName(modVault.bas:370-378)の再現。テストから Private は
' 呼べないので、置換表と60字クランプをここに写す。片方を変えたら両方直すこと
' ―― 写しがズレたらこのテストは「自分の写しに対して冪等」を測るだけになる。
Private Function SanitizeLikeVault40(ByVal s As String) As String
    Dim bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)
    Dim t As String: t = s
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        t = Replace(t, CStr(bad(i)), "_")
    Next i
    SanitizeLikeVault40 = modUtil.SafeLeft(Trim$(t), 60)
End Function

' ---- C8: BuildMemoBody の4引数書式(R37 §2-2「誤答の根拠」行) ----------------
'   discriminate:
'   ・sourcesLine を無視する実装にすると(a)(b)が落ちる。
'   ・「誤答の根拠: 」のヘッダーを付けない/別の文字列にする実装にすると
'     (a)の InStr 判定と(b)の ExtractWrongSources 往復の両方が落ちる。
'   ・sourcesLine が空でも常に行を足す実装にすると(c)(d)が落ちる
'     (「sourcesLine は空でもよい」= 空なら行そのものを出さない)。
Private Sub TestBuildMemoBodySourcesLine40()
    ' (a)(b) sourcesLine が非空 → 末尾に「誤答の根拠:」行が載り、往復も取れる。
    Dim body As String
    body = modCorrect.BuildMemoBody("退職金の計算方法は?", "勤続年数×基本給×0.6です。", _
                                    "退職金は一律50万円です", "就業規則.txt p.3 | 給与規程.txt p.5")
    ChkBool40 "C8_誤答の根拠行が末尾に載る", _
        (InStr(1, body, vbLf & "誤答の根拠: 就業規則.txt p.3 | 給与規程.txt p.5", vbBinaryCompare) > 0), True
    ChkStr40 "C8_誤答の根拠行を往復で取り出せる", _
        modCorrect.ExtractWrongSources(body), "誤答の根拠: 就業規則.txt p.3 | 給与規程.txt p.5"
    ' 既存の書式(質問・正しい内容)は壊れていない。
    ChkStr40 "C8_4引数でも質問の往復は変わらない", _
        modCorrect.ExtractQuestionLine(body), "退職金の計算方法は?"

    ' (c)(d) sourcesLine が空 → 行そのものが現れない(既存3引数相当の書式)。
    Dim bodyNoSrc As String
    bodyNoSrc = modCorrect.BuildMemoBody("退職金の計算方法は?", "勤続年数×基本給×0.6です。", _
                                        "退職金は一律50万円です", "")
    ChkBool40 "C8_空なら誤答の根拠行は出ない", _
        (InStr(1, bodyNoSrc, "誤答の根拠", vbBinaryCompare) = 0), True
    ChkStr40 "C8_空のときExtractWrongSourcesも空", modCorrect.ExtractWrongSources(bodyNoSrc), ""
End Sub

' ---- C9: ExtractWrongSources -------------------------------------------------
'   discriminate:
'   ・行単位ではなく本文全体を返す実装にすると(a)(b)が余分な行を含んで落ちる。
'   ・"誤答の根拠:" が無いときに空以外を返す実装にすると(c)(d)が落ちる。
Private Sub TestExtractWrongSources40()
    ' (a) 単独1件。
    ChkStr40 "C9_単独1件", _
        modCorrect.ExtractWrongSources("質問: x" & vbLf & "誤答の根拠: 就業規則.txt p.3"), _
        "誤答の根拠: 就業規則.txt p.3"

    ' (b) 複数件("|"区切り)。前後に他の行があっても行単位で拾える。
    ChkStr40 "C9_複数件", _
        modCorrect.ExtractWrongSources("質問: x" & vbLf & _
            "誤答の根拠: A.txt p.1 | B.txt p.2 | C.txt p.9" & vbLf & "正しい内容: y"), _
        "誤答の根拠: A.txt p.1 | B.txt p.2 | C.txt p.9"

    ' (c) 行が無い本文 → 空(古い是正メモ。sourcesLine が空だったケース)。
    ChkStr40 "C9_行が無ければ空", _
        modCorrect.ExtractWrongSources("質問: x" & vbLf & "正しい内容: y"), ""

    ' (d) 空文字 → 空。
    ChkStr40 "C9_空文字は空", modCorrect.ExtractWrongSources(""), ""
End Sub

' ---- C10: WrongSourceMatches((source,page)一致判定) --------------------------
'   discriminate:
'   ・資料名を大小区別で比較する実装にすると(b)が落ちる。
'   ・ページを文字列比較にする実装にすると、桁の異なる数値表記で(c)相当が
'     誤って一致扱いになりうる(ここでは数値一致のみで判定することを(a)(c)で
'     押さえる)。
'   ・"|" 区切りの2件目以降を見ない実装にすると(e)が落ちる。
'   ・空行/空資料名を無条件 True にする実装にすると(f)が落ちる。
Private Sub TestWrongSourceMatches40()
    Dim line As String: line = "誤答の根拠: 就業規則.txt p.3 | 給与規程.txt p.12"

    ' (a) 資料名・ページとも一致。
    ChkBool40 "C10_資料名とページが一致", _
        modCorrect.WrongSourceMatches(line, "就業規則.txt", 3), True

    ' (b) 資料名の大小文字の揺れは無視する。
    ChkBool40 "C10_資料名は大小無視", _
        modCorrect.WrongSourceMatches("誤答の根拠: RULE.TXT p.3", "rule.txt", 3), True

    ' (c) ページが違えば不一致(数値一致で判定)。
    ChkBool40 "C10_ページが違えば不一致", _
        modCorrect.WrongSourceMatches(line, "就業規則.txt", 4), False

    ' (d) 資料名が違えば不一致。
    ChkBool40 "C10_資料名が違えば不一致", _
        modCorrect.WrongSourceMatches(line, "無関係の資料.txt", 3), False

    ' (e) "|"区切りの2件目だけに一致するケースも拾える。
    ChkBool40 "C10_複数件のうち2件目に一致", _
        modCorrect.WrongSourceMatches(line, "給与規程.txt", 12), True

    ' (f) 行が空/資料名が空なら不一致(恒真にならないことの確認)。
    ChkBool40 "C10_行が空なら不一致", modCorrect.WrongSourceMatches("", "就業規則.txt", 3), False
    ChkBool40 "C10_資料名が空なら不一致", modCorrect.WrongSourceMatches(line, "", 3), False

    ' (g) R41 §1 A: Excel由来の新表記「シートN」も受ける(旧表記" p."と混在する行)。
    Dim lineX As String: lineX = "誤答の根拠: 売上.xlsx シート2 | 就業規則.txt p.3"
    ChkBool40 "C10_シート表記が一致", _
        modCorrect.WrongSourceMatches(lineX, "売上.xlsx", 2), True
    ChkBool40 "C10_シート表記はページ違いで不一致", _
        modCorrect.WrongSourceMatches(lineX, "売上.xlsx", 3), False
    ChkBool40 "C10_混在行の旧表記p.も一致", _
        modCorrect.WrongSourceMatches(lineX, "就業規則.txt", 3), True

    ' (h) R41 Fix m4: 接尾辞あり項目の【後】に接尾辞なし項目が来ても、前の
    '     項目の切り分け位置を持ち越さない(Dim はループで再実行されない)。
    '     持ち越すと "ab1234" が entSrc="ab"/pageStr="34" に化けて誤一致する。
    Dim lineY As String: lineY = "誤答の根拠: x.pdf p.1 | ab1234"
    ChkBool40 "C10_接尾辞なし項目は前の位置を持ち越さない", _
        modCorrect.WrongSourceMatches(lineY, "ab", 34), False
    ChkBool40 "C10_接尾辞なし項目の前の項目は一致", _
        modCorrect.WrongSourceMatches(lineY, "x.pdf", 1), True
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
H06Next40:
    On Error GoTo H06Fail40
    TestQuestionTag4_40
H07Next40:
    On Error GoTo H07Fail40
    TestMatchLevel40Boundary
H08Next40:
    On Error GoTo H08Fail40
    TestBuildMemoBodySourcesLine40
H09Next40:
    On Error GoTo H09Fail40
    TestExtractWrongSources40
H10Next40:
    On Error GoTo H10Fail40
    TestWrongSourceMatches40
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
    Resume H06Next40
H06Fail40:
    modTestRunner.Check "TestQuestionTag4_40(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H07Next40
H07Fail40:
    modTestRunner.Check "TestMatchLevel40Boundary(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H08Next40
H08Fail40:
    modTestRunner.Check "TestBuildMemoBodySourcesLine40(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H09Next40
H09Fail40:
    modTestRunner.Check "TestExtractWrongSources40(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H10Next40
H10Fail40:
    modTestRunner.Check "TestWrongSourceMatches40(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done40
End Sub
