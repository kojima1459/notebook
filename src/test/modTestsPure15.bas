Attribute VB_Name = "modTestsPure15"
Option Explicit

' ============================================================================
' modTestsPure15 - R16-3B/3C/3D(複合質問の精度・波3)の純ロジック回帰テスト
' ----------------------------------------------------------------------------
' なぜ新設したか(憲章§4-6):
'   modTestsPure14 が24,825字で、ここの真理表を足すと30,000字上限を超える。
'   14を新設したときと同じ線で分割する。
'   入口は modTestsPure14.RunAll14 の末尾から呼ばれる RunAll15 の1本だけ。
'
' ここで固定するもの:
'   ・modRagParse.ParseChoiceNumbers(3B): 番号選択の読み取り。「1」「1と3」
'     「①と③」を拾い、【数字と区切り以外の文字が1つでも混じれば拾わない】。
'     ここが甘いと、聞き返しのあと自分の言葉で書き直した人の文章が黙って
'     捨てられる(2026-07-28 レビュー H-12 の実バグと同型)。
'   ・modRagParse.ParseOptions(3B): 段0の <options> の読み取り(欠落・空・切詰)。
'   ・modClarify.MergeTopicAnswer(3B): 選んだ読み方と元質問の合成。単一/複数/
'     書き直し/範囲外。IsPendingExpired はTTL30分の境界。
'   ・modRagParse.BuildErrorAnswer(R16H FA-2): #ERR: は従来の友好文へ、それ以外
'     (逆質問などの非回答テキスト)は一字一句そのまま素通し、空は従来どおり。
'   ・modAskFocus.ParseChunkKey / NeighborIdList(3C): chunk_id からの文書順
'     復元と前後radiusの取り方。ページ跨ぎで seq がリセットされる並び、
'     radius=0/1/2、窓の重なり、資料の境界、重複排除。真理表は【実データ形状】
'     =行ごとに異なるハッシュ+同一 source で書く(R16H FA-1。理由は
'     TestNeighborIdList の冒頭)。
'   ・modFollowup.DemoteOrder / RememberUsedIds(3D): 既出チャンク降格の規則。
'     未出優先・安定・keepK境界・全既出でも件数が減らないこと(埋め戻し)。
'     Hit配列そのものの入れ替え(DemoteUsed)は、別モジュールで定義した Type の
'     配列を LibreOffice が扱えないためここでは動かせない(modTestsPure 冒頭の
'     既知制約)。規則は DemoteOrder が単一情報源なので、ここを固定すれば
'     「どれを残すか」の判断は守れる。並べ替えの実行そのものは実機スモーク
'     テスト項目27で確認する。
' ============================================================================

' ----------------------------------------------------------------------------
' R16-3B(1): 番号選択の読み取り。
' ----------------------------------------------------------------------------
Private Sub TestParseChoiceNumbers()
    modTestRunner.Check "選択番号_単一", _
        (modRagParse.ParseChoiceNumbers("1", 3) = "1")
    modTestRunner.Check "選択番号_とで複数", _
        (modRagParse.ParseChoiceNumbers("1と3", 3) = "1,3")
    modTestRunner.Check "選択番号_丸数字で複数", _
        (modRagParse.ParseChoiceNumbers(ChrW(&H2460) & "と" & ChrW(&H2462), 3) = "1,3")
    modTestRunner.Check "選択番号_半角カンマ", _
        (modRagParse.ParseChoiceNumbers("1,3", 3) = "1,3")
    modTestRunner.Check "選択番号_読点", _
        (modRagParse.ParseChoiceNumbers("2、3", 3) = "2,3")
    modTestRunner.Check "選択番号_半角スペース", _
        (modRagParse.ParseChoiceNumbers("1 3", 3) = "1,3")
    modTestRunner.Check "選択番号_全角数字と全角スペース", _
        (modRagParse.ParseChoiceNumbers(ChrW(&HFF11) & ChrW(&H3000) & ChrW(&HFF13), 3) = "1,3")

    ' 範囲外は無視する(表示していない番号を当てはめない)。
    modTestRunner.Check "選択番号_範囲外は落とす", _
        (modRagParse.ParseChoiceNumbers("1と5", 3) = "1")
    ' 連続した桁は1つの数として読む。maxN=3 なら 13 は範囲外=0件。
    modTestRunner.Check "選択番号_13は十三(範囲外なら空)", _
        (modRagParse.ParseChoiceNumbers("13", 3) = "")
    modTestRunner.Check "選択番号_13は十三(範囲内なら採用)", _
        (modRagParse.ParseChoiceNumbers("13", 20) = "13")
    modTestRunner.Check "選択番号_重複は1回だけ", _
        (modRagParse.ParseChoiceNumbers("1と1と2", 3) = "1,2")
    ' 重複の畳み込みは【出現順】を保つ(後から出た番号を前へ繰り上げない)。
    modTestRunner.Check "選択番号_重複を畳んでも出現順", _
        (modRagParse.ParseChoiceNumbers("3と1と3", 3) = "3,1"), _
        "実際=" & modRagParse.ParseChoiceNumbers("3と1と3", 3)

    ' R16H FA-7(A-M7/B-M4): 資料の聞き返し(modClarify)は「2-②」と打つよう
    ' 教えている。同じ癖で「1-3」と打った人の返事を書き直し扱いで捨てないよう、
    ' ハイフン類と括弧も区切りとして認める(範囲指定ではなく複数選択の意味)。
    modTestRunner.Check "選択番号_半角ハイフン区切り", _
        (modRagParse.ParseChoiceNumbers("1-3", 3) = "1,3"), _
        "実際=" & modRagParse.ParseChoiceNumbers("1-3", 3)
    modTestRunner.Check "選択番号_資料型の癖(2-丸2)", _
        (modRagParse.ParseChoiceNumbers("2-" & ChrW(&H2461), 3) = "2"), _
        "実際=" & modRagParse.ParseChoiceNumbers("2-" & ChrW(&H2461), 3)
    modTestRunner.Check "選択番号_括弧付き", _
        (modRagParse.ParseChoiceNumbers("(1)(3)", 3) = "1,3"), _
        "実際=" & modRagParse.ParseChoiceNumbers("(1)(3)", 3)
    modTestRunner.Check "選択番号_全角括弧付き", _
        (modRagParse.ParseChoiceNumbers(ChrW(&HFF08) & "1" & ChrW(&HFF09) & _
            ChrW(&HFF08) & "3" & ChrW(&HFF09), 3) = "1,3")
    modTestRunner.Check "選択番号_長音での打ち間違い", _
        (modRagParse.ParseChoiceNumbers("1" & ChrW(&H30FC) & "3", 3) = "1,3"), _
        "実際=" & modRagParse.ParseChoiceNumbers("1" & ChrW(&H30FC) & "3", 3)
    modTestRunner.Check "選択番号_全角ハイフン", _
        (modRagParse.ParseChoiceNumbers("1" & ChrW(&HFF0D) & "3", 3) = "1,3")
    modTestRunner.Check "選択番号_マイナス記号", _
        (modRagParse.ParseChoiceNumbers("1" & ChrW(&H2212) & "3", 3) = "1,3")

    ' ここからが本丸。数字以外の語が混じったら【書き直し】として空を返す。
    modTestRunner.Check "選択番号_数字と語の混在は空", _
        (modRagParse.ParseChoiceNumbers("1番のほうを詳しく", 3) = "")
    modTestRunner.Check "選択番号_第1条の質問は空", _
        (modRagParse.ParseChoiceNumbers("第1条の適用範囲は?", 3) = "")
    modTestRunner.Check "選択番号_3日以内は空", _
        (modRagParse.ParseChoiceNumbers("3日以内に出す必要ある?", 3) = "")
    modTestRunner.Check "選択番号_空文字は空", _
        (modRagParse.ParseChoiceNumbers("", 3) = "")
    modTestRunner.Check "選択番号_空白だけは空", _
        (modRagParse.ParseChoiceNumbers("  ", 3) = "")
    modTestRunner.Check "選択番号_maxN0は空", _
        (modRagParse.ParseChoiceNumbers("1", 0) = "")
    ' 数字が1つも無い区切りだけの返事も空(有効番号0件=書き直し扱い)。
    modTestRunner.Check "選択番号_区切りだけは空", _
        (modRagParse.ParseChoiceNumbers("と、", 3) = "")
End Sub

' ----------------------------------------------------------------------------
' R16-3B(2): <options> のパース(ParseParts と同型の寛容退化)。
' ----------------------------------------------------------------------------
Private Sub TestParseOptions()
    Dim opts() As String
    Dim n As Long

    n = modRagParse.ParseOptions( _
        "<verdict>clarify</verdict><options>契約の解約手続き | 解約金の計算 | 解約後の補償</options>", _
        3, opts)
    modTestRunner.Check "options_正常3件", (n = 3), "実際=" & n
    If n = 3 Then
        modTestRunner.Check "options_前後の空白が落ちる", (opts(0) = "契約の解約手続き"), "実際=" & opts(0)
        modTestRunner.Check "options_3件目", (opts(2) = "解約後の補償"), "実際=" & opts(2)
    End If

    n = modRagParse.ParseOptions("<verdict>clarify</verdict>", 3, opts)
    modTestRunner.Check "options_タグ欠落は0件", (n = 0), "実際=" & n

    n = modRagParse.ParseOptions("<options></options>", 3, opts)
    modTestRunner.Check "options_空タグは0件", (n = 0), "実際=" & n

    n = modRagParse.ParseOptions("<options>a | b | c | d</options>", 3, opts)
    modTestRunner.Check "options_maxNで切り詰め", (n = 3), "実際=" & n

    n = modRagParse.ParseOptions("#ERR:E0202:失敗", 3, opts)
    modTestRunner.Check "options_ERR応答は0件", (n = 0), "実際=" & n
End Sub

' ----------------------------------------------------------------------------
' R16-3B(3): 選んだ読み方と元質問の合成 + 保留のTTL。
' ----------------------------------------------------------------------------
Private Sub TestMergeTopicAnswer()
    Dim q As String: q = "解約について教えて"
    Dim o As String: o = "解約手続きの流れ|解約金の計算|解約後の補償"

    modTestRunner.Check "読み方合成_単一", _
        (modClarify.MergeTopicAnswer(q, o, "2") = q & " / 解約金の計算"), _
        "実際=" & modClarify.MergeTopicAnswer(q, o, "2")

    ' 複数選択はそのまま並べる(次ターンの段0が parts へ割る)。
    modTestRunner.Check "読み方合成_複数は列挙", _
        (modClarify.MergeTopicAnswer(q, o, "1と3") = q & " / 解約手続きの流れ / 解約後の補償"), _
        "実際=" & modClarify.MergeTopicAnswer(q, o, "1と3")

    ' 番号ではなく書き直した人の文章は、そのまま新しい質問として扱う。
    modTestRunner.Check "読み方合成_書き直しは打った文を優先", _
        (modClarify.MergeTopicAnswer(q, o, "解約金の上限はいくらですか") = "解約金の上限はいくらですか")

    ' 短い上に番号でもない返事は捨てずに元質問へ足す。
    modTestRunner.Check "読み方合成_短文は元質問へ足す", _
        (modClarify.MergeTopicAnswer(q, o, "違う") = q & " 違う")

    ' 範囲外だけの返事は番号として成立しない=書き直し扱い(短いので合成)。
    modTestRunner.Check "読み方合成_範囲外だけは書き直し扱い", _
        (modClarify.MergeTopicAnswer(q, o, "9") = q & " 9")

    ' 選択肢が無い保留(壊れた状態)でも落ちない。
    modTestRunner.Check "読み方合成_選択肢なしでも落ちない", _
        (modClarify.MergeTopicAnswer(q, "", "1") = q & " 1")

    ' TTL: 30分ちょうどは有効、31分で失効。読めない時刻は失効、空は有効。
    Dim t0 As Date: t0 = CDate("2026-08-05 10:00:00")
    modTestRunner.Check "保留TTL_29分は有効", _
        (modClarify.IsPendingExpired("2026-08-05 09:31:00", t0) = False)
    modTestRunner.Check "保留TTL_30分ちょうどは有効", _
        (modClarify.IsPendingExpired("2026-08-05 09:30:00", t0) = False)
    modTestRunner.Check "保留TTL_31分で失効", _
        (modClarify.IsPendingExpired("2026-08-05 09:29:00", t0) = True)
    modTestRunner.Check "保留TTL_空は有効(旧データ)", _
        (modClarify.IsPendingExpired("", t0) = False)
    modTestRunner.Check "保留TTL_読めない時刻は失効", _
        (modClarify.IsPendingExpired("きのう", t0) = True)
End Sub

' ----------------------------------------------------------------------------
' R16-3B(4): 逆質問の本文と発動条件。
' ----------------------------------------------------------------------------
Private Sub TestClarifyAsk()
    modTestRunner.Check "逆質問_offなら出さない", _
        (modAskMulti.ShouldClarify("off", 3) = False)
    modTestRunner.Check "逆質問_autoで2件なら出す", _
        (modAskMulti.ShouldClarify("auto", 2) = True)
    modTestRunner.Check "逆質問_1件では出さない(選べない)", _
        (modAskMulti.ShouldClarify("auto", 1) = False)
    modTestRunner.Check "逆質問_0件では出さない", _
        (modAskMulti.ShouldClarify("auto", 0) = False)
    modTestRunner.Check "逆質問_知らない値はauto扱い", _
        (modAskMulti.ShouldClarify("yes", 2) = True)

    Dim opts(0 To 2) As String
    opts(0) = "解約手続きの流れ"
    opts(1) = "解約金の計算"
    opts(2) = "解約後の補償"
    Dim s As String: s = modAskMulti.BuildClarifyAsk(opts, 3)
    modTestRunner.Check "逆質問文_番号付きで列挙", (InStr(s, "2) 解約金の計算") > 0), "実際=" & s
    modTestRunner.Check "逆質問文_複数選択の例を示す", (InStr(s, "1と3") > 0), "実際=" & s
    modTestRunner.Check "逆質問文_書き直しの道を残す", (InStr(s, "書き直") > 0), "実際=" & s
    ' R16H FB-6(B-M5): 保留はTTL30分で失効し、以後の「1」は普通の新しい質問に
    ' なる。書いておかないと、席へ戻って番号を打った人に理由の分からない答えが
    ' 返る(TTLそのものの境界は TestMergeTopicAnswer の IsPendingExpired 側)。
    modTestRunner.Check "逆質問文_30分で無効になると明記", _
        (InStr(s, "30分") > 0), "実際=" & s
End Sub

' ----------------------------------------------------------------------------
' R16H FA-2: BuildErrorAnswer の素通し(逆質問を偽E0202にしない)。
' ----------------------------------------------------------------------------
'   入念モードの逆質問は ok=False で返る=modAsk の thorough 分岐では
'   BuildErrorAnswer を必ず通る。ここで #ERR以外まで E0202 の定型文へ丸めると、
'   利用者には「番号の選択肢」ではなく「うまく回答をまとめられませんでした」が
'   出て、見えない保留に次の発言が吸収される。素通しはその一点を守る契約なので、
'   #ERR: の従来動作と一緒に固定しておく。
' ----------------------------------------------------------------------------
Private Sub TestBuildErrorAnswer()
    ' (1) #ERR: は従来どおりコードを読んで利用者向けの文言+コード表記にする。
    Dim s As String
    s = modRagParse.BuildErrorAnswer("#ERR:E0303:画像PDFです")
    modTestRunner.Check "エラー回答_ERRはコードの友好文になる", _
        (InStr(s, "(コード: E0303)") > 0), "実際=" & s
    modTestRunner.Check "エラー回答_ERRは元の生文字列を出さない", _
        (InStr(s, "#ERR:") = 0), "実際=" & s
    ' コード部が読めない #ERR: は従来どおり E0202 へ丸める。
    modTestRunner.Check "エラー回答_コード無しERRはE0202", _
        (InStr(modRagParse.BuildErrorAnswer("#ERR:") , "(コード: E0202)") > 0), _
        "実際=" & modRagParse.BuildErrorAnswer("#ERR:")

    ' (2) #ERR以外は一字一句そのまま返る(逆質問・非回答テキストの素通し)。
    Dim ask As String
    ask = "ご質問はいくつかの読み方ができます。どれを調べますか?" & vbLf & _
          "  1) 解約手続きの流れ" & vbLf & "  2) 解約金の計算"
    modTestRunner.Check "エラー回答_非ERRは一字一句そのまま", _
        (modRagParse.BuildErrorAnswer(ask) = ask), _
        "実際=" & modRagParse.BuildErrorAnswer(ask)
    ' 先頭が "#" でも "#ERR:" でなければ本文扱い(判定は前方一致だけ)。
    modTestRunner.Check "エラー回答_ERR風の本文も素通し", _
        (modRagParse.BuildErrorAnswer("#ERROR風の見出し") = "#ERROR風の見出し")

    ' (3) 空文字・空白だけは従来どおり E0202(素通しすると無言の回答になる)。
    modTestRunner.Check "エラー回答_空文字は従来どおりE0202", _
        (InStr(modRagParse.BuildErrorAnswer(""), "(コード: E0202)") > 0), _
        "実際=" & modRagParse.BuildErrorAnswer("")
    modTestRunner.Check "エラー回答_空白だけも従来どおりE0202", _
        (InStr(modRagParse.BuildErrorAnswer("   "), "(コード: E0202)") > 0), _
        "実際=" & modRagParse.BuildErrorAnswer("   ")
End Sub

' ----------------------------------------------------------------------------
' R16-3C(1): chunk_id "bs::ハッシュ::pN::cM" の分解。
'   d(docKey)は「末尾2要素を除いた部分」=チャンク本文のハッシュであって、
'   資料キーではない(R16H FA-1)。ここで固定するのは形式の読み取りと
'   (page, seq) の取り出しだけで、資料の同一性は source 側が持つ。
' ----------------------------------------------------------------------------
Private Sub TestParseChunkKey()
    Dim d As String, p As Long, s As Long

    modTestRunner.Check "chunk_id_正常", _
        (modAskFocus.ParseChunkKey("bs::a1b2::p3::c2", d, p, s) = True)
    modTestRunner.Check "chunk_id_正常_docKey", (d = "bs::a1b2"), "実際=" & d
    modTestRunner.Check "chunk_id_正常_page", (p = 3), "実際=" & p
    modTestRunner.Check "chunk_id_正常_seq", (s = 2), "実際=" & s

    ' ハッシュ部に "::" が入っていても末尾2要素から読むので壊れない。
    modTestRunner.Check "chunk_id_ハッシュ内に区切りがあっても読める", _
        (modAskFocus.ParseChunkKey("bs::a::b::p10::c1", d, p, s) = True)
    modTestRunner.Check "chunk_id_その場合のdocKey", (d = "bs::a::b"), "実際=" & d
    modTestRunner.Check "chunk_id_その場合のpage", (p = 10), "実際=" & p

    modTestRunner.Check "chunk_id_要素不足は不可", _
        (modAskFocus.ParseChunkKey("bs::a1b2::p3", d, p, s) = False)
    modTestRunner.Check "chunk_id_pもcも無いのは不可", _
        (modAskFocus.ParseChunkKey("bs::a::x3::y2", d, p, s) = False)
    modTestRunner.Check "chunk_id_数字でないのは不可", _
        (modAskFocus.ParseChunkKey("bs::a::pX::c1", d, p, s) = False)
    modTestRunner.Check "chunk_id_空は不可", _
        (modAskFocus.ParseChunkKey("", d, p, s) = False)
    modTestRunner.Check "chunk_id_無関係な文字列は不可", _
        (modAskFocus.ParseChunkKey("c1", d, p, s) = False)
    ' 失敗時は出力を汚さない(前の値が残ると隣を1つずらす)。
    modTestRunner.Check "chunk_id_失敗時はpage0", (p = 0), "実際=" & p
End Sub

' ----------------------------------------------------------------------------
' R16-3C(2): 文書順の復元と前後radiusの取り方。
' ----------------------------------------------------------------------------
' 2026-08-05(R16H FA-1 / A-H1): ここの形を【実データの形】へ入れ替えた。
'   旧テストは "bs::A::p1::c1" と "bs::A::p1::c2" のように同じ資料の行が
'   ハッシュ部を共有する形で書かれていて、ハッシュ部を資料キーにする実装が
'   そのまま通っていた。実データの chunk_id は
'   "bs::" & Fnv1a64Hex(チャンク本文) & "::pN::cM"(modShelf)で、
'   ハッシュは【行ごとに違う】。そのため実機では同じ資料の隣同士が別資料と
'   判定され、近傍が恒久的に0件だった(真理表は通るのに実機では一度も
'   効かない、という一番たちの悪い外し方)。
'   以後この真理表は「行ごとに異なるハッシュ + 同一 source」で書く。
' ----------------------------------------------------------------------------
Private Sub TestNeighborIdList()
    ' 資料A(約款.pdf): p1に2チャンク、p2に2チャンク、p3に1チャンク。
    ' ハッシュは実データ同様に全行バラバラ。資料の同一性は source 側だけが持つ。
    Const A1 As String = "bs::7c1f::p1::c1"
    Const A2 As String = "bs::9b04::p1::c2"
    Const A3 As String = "bs::22ae::p2::c1"
    Const A4 As String = "bs::e5d7::p2::c2"
    Const A5 As String = "bs::0f38::p3::c1"
    Const SRC_A As String = "yakkan.pdf"

    ' わざとシート上の並びを乱して渡す(文書順は (source, page, seq) で作り直す)。
    Dim ids As String, srcs As String
    ids = A4 & vbLf & A1 & vbLf & A5 & vbLf & A2 & vbLf & A3
    srcs = SRC_A & vbLf & SRC_A & vbLf & SRC_A & vbLf & SRC_A & vbLf & SRC_A
    ' 文書順: A1(p1c1), A2(p1c2), A3(p2c1), A4(p2c2), A5(p3c1)

    modTestRunner.Check "近傍_radius0は無し", _
        (modAskFocus.NeighborIdList(ids, srcs, A3, 0) = "")

    ' radius=1: p2c1 の前後 = p1c2 と p2c2。文書順で返る。
    ' ハッシュが全部違っても同じ資料として並ぶこと(FA-1の本丸)。
    modTestRunner.Check "近傍_radius1は前後1つずつ(ハッシュは行ごとに別)", _
        (modAskFocus.NeighborIdList(ids, srcs, A3, 1) = A2 & vbLf & A4), _
        "実際=" & Replace(modAskFocus.NeighborIdList(ids, srcs, A3, 1), vbLf, "/")

    ' radius=2: ページ跨ぎでも seq のリセットに引きずられない。
    modTestRunner.Check "近傍_radius2はページを跨ぐ", _
        (modAskFocus.NeighborIdList(ids, srcs, A3, 2) = _
         A1 & vbLf & A2 & vbLf & A4 & vbLf & A5), _
        "実際=" & Replace(modAskFocus.NeighborIdList(ids, srcs, A3, 2), vbLf, "/")

    ' 端は切り詰める(存在しない前後を作らない)。
    modTestRunner.Check "近傍_先頭は前が無い", _
        (modAskFocus.NeighborIdList(ids, srcs, A1, 1) = A2)
    modTestRunner.Check "近傍_末尾は後ろが無い", _
        (modAskFocus.NeighborIdList(ids, srcs, A5, 1) = A4)

    ' 窓が重なっても重複しない。ヒット自身は結果に入らない。
    Dim two As String: two = A2 & vbLf & A3
    modTestRunner.Check "近傍_窓の重なりはマージされる", _
        (modAskFocus.NeighborIdList(ids, srcs, two, 1) = A1 & vbLf & A4), _
        "実際=" & Replace(modAskFocus.NeighborIdList(ids, srcs, two, 1), vbLf, "/")

    ' 資料の境界は越えない(別の資料の先頭が「前のチャンク」にならない)。
    ' 判定は source。ページ番号もチャンク番号も同じ2資料で確かめる。
    Dim mIds As String, mSrcs As String
    mIds = A1 & vbLf & A2 & vbLf & "bs::4d90::p1::c1" & vbLf & "bs::1a55::p1::c2"
    mSrcs = SRC_A & vbLf & SRC_A & vbLf & "manual.docx" & vbLf & "manual.docx"
    modTestRunner.Check "近傍_資料の境界を越えない(sourceで切る)", _
        (modAskFocus.NeighborIdList(mIds, mSrcs, "bs::4d90::p1::c1", 1) = "bs::1a55::p1::c2"), _
        "実際=" & Replace(modAskFocus.NeighborIdList(mIds, mSrcs, "bs::4d90::p1::c1", 1), vbLf, "/")

    ' 読めない chunk_id は並びから外す(混ぜると隣が1つずれる)。
    Dim dIds As String, dSrcs As String
    dIds = A1 & vbLf & "kowareta" & vbLf & A2
    dSrcs = SRC_A & vbLf & SRC_A & vbLf & SRC_A
    modTestRunner.Check "近傍_壊れたidは無視する", _
        (modAskFocus.NeighborIdList(dIds, dSrcs, A1, 1) = A2)

    ' source が無い行も並びから外す(どの資料の隣か決められない)。
    Dim nIds As String, nSrcs As String
    nIds = A1 & vbLf & A2 & vbLf & A3
    nSrcs = SRC_A & vbLf & "" & vbLf & SRC_A
    modTestRunner.Check "近傍_source空の行は並びから外す", _
        (modAskFocus.NeighborIdList(nIds, nSrcs, A1, 1) = A3), _
        "実際=" & Replace(modAskFocus.NeighborIdList(nIds, nSrcs, A1, 1), vbLf, "/")
    ' source の本数が足りないときも、余ったidを無所属で混ぜない。
    modTestRunner.Check "近傍_source不足の行は無視する", _
        (modAskFocus.NeighborIdList(nIds, SRC_A, A1, 1) = "")

    ' 入力が空・ヒットが候補に無いときは何も足さない。
    modTestRunner.Check "近傍_候補が空なら何も足さない", _
        (modAskFocus.NeighborIdList("", "", A1, 1) = "")
    modTestRunner.Check "近傍_ヒットが空なら何も足さない", _
        (modAskFocus.NeighborIdList(ids, srcs, "", 1) = "")
    modTestRunner.Check "近傍_候補に無いヒットは無視", _
        (modAskFocus.NeighborIdList(ids, srcs, "bs::ffff::p1::c1", 1) = "")
End Sub

' ----------------------------------------------------------------------------
' R16-3D: 既出チャンクの降格(未出優先・安定・埋め戻し)。
' ----------------------------------------------------------------------------
Private Sub TestDemoteUsed()
    ' 既出降格の規則は modFollowup.DemoteOrder(純関数)が単一情報源で、
    ' DemoteUsed はその順に Hit を入れ替える係。ここでは規則そのものを固定する。
    ' Hit配列を直接組まないのは、別モジュールで定義した Type の配列が
    ' LibreOffice の実行テスト環境で扱えないため(modTestsPure 冒頭の既知制約。
    ' 配列の入れ替え自体は実機スモークテスト項目27で確認する)。
    Dim ids As String
    ids = "bs::A::p1::c1" & vbLf & "bs::A::p1::c2" & vbLf & _
          "bs::A::p2::c1" & vbLf & "bs::A::p2::c2"

    ' 会話をまたぐ状態を持つので、必ず新規質問(=リセット)から始める。
    modFollowup.SetFollowupTurn False

    ' (1) 既出が1件も無いターンは並べ替えず、keepK へ切るだけ。
    modTestRunner.Check "降格_既出なしは元の順序のまま", _
        (modFollowup.DemoteOrder(ids, 4) = "1,2,3,4"), _
        "実際=" & modFollowup.DemoteOrder(ids, 4)
    modTestRunner.Check "降格_既出なしでもkeepKで切る", _
        (modFollowup.DemoteOrder(ids, 2) = "1,2"), _
        "実際=" & modFollowup.DemoteOrder(ids, 2)

    ' (2) 1件目と3件目を既出にする → 未出(2,4)が前へ、既出(1,3)が後ろへ。
    modFollowup.SetFollowupTurn True
    modFollowup.RememberUsedIds "bs::A::p1::c1" & vbLf & "bs::A::p2::c1"
    modTestRunner.Check "降格_未出が前へ既出が後ろへ(安定)", _
        (modFollowup.DemoteOrder(ids, 4) = "2,4,1,3"), _
        "実際=" & modFollowup.DemoteOrder(ids, 4)

    ' (3) keepK 境界: 2件へ切ると未出だけが残る。
    modTestRunner.Check "降格_keepKで切ると未出だけ残る", _
        (modFollowup.DemoteOrder(ids, 2) = "2,4"), _
        "実際=" & modFollowup.DemoteOrder(ids, 2)
    ' 未出が2件しか無いのに3件要るときは既出で埋め戻す(空にしない保証)。
    modTestRunner.Check "降格_未出が足りなければ既出で埋め戻す", _
        (modFollowup.DemoteOrder(ids, 3) = "2,4,1"), _
        "実際=" & modFollowup.DemoteOrder(ids, 3)

    ' (4) 全部が既出でも件数は減らず、順序は元のまま(除外ではなく降格)。
    modFollowup.RememberUsedIds ids
    modTestRunner.Check "降格_全既出でも空にならない", _
        (modFollowup.DemoteOrder(ids, 3) = "1,2,3"), _
        "実際=" & modFollowup.DemoteOrder(ids, 3)
    modTestRunner.Check "降格_全既出なら元の順序のまま", _
        (modFollowup.DemoteOrder(ids, 4) = "1,2,3,4"), _
        "実際=" & modFollowup.DemoteOrder(ids, 4)

    ' (5) chunk_id が空のヒットは既出にできない=常に未出として前に残る。
    Dim mixed As String
    mixed = "bs::A::p1::c1" & vbLf & "" & vbLf & "bs::A::p2::c1"
    modTestRunner.Check "降格_id空は未出扱いで先頭へ", _
        (modFollowup.DemoteOrder(mixed, 3) = "2,1,3"), _
        "実際=" & modFollowup.DemoteOrder(mixed, 3)
    modFollowup.RememberUsedIds ""      ' 空行は覚えない(記憶は変わらない)
    modTestRunner.Check "降格_空idを覚えても記憶は変わらない", _
        (modFollowup.DemoteOrder(ids, 4) = "1,2,3,4"), _
        "実際=" & modFollowup.DemoteOrder(ids, 4)

    ' (6) 空集合・keepK 0以下・1件だけでも落ちない。
    modTestRunner.Check "降格_空入力は空", (modFollowup.DemoteOrder("", 3) = "")
    modTestRunner.Check "降格_keepK0は全件扱い", _
        (modFollowup.DemoteOrder(ids, 0) = "1,2,3,4"), _
        "実際=" & modFollowup.DemoteOrder(ids, 0)
    modTestRunner.Check "降格_keepKが件数超なら全件", _
        (modFollowup.DemoteOrder(ids, 99) = "1,2,3,4"), _
        "実際=" & modFollowup.DemoteOrder(ids, 99)

    ' (7) 新規質問(SetFollowupTurn False)で既出メモリが消える。
    modFollowup.SetFollowupTurn False
    modTestRunner.Check "降格_新規質問で旗が降りる", (modFollowup.IsFollowupTurn() = False)
    modTestRunner.Check "降格_新規質問後は並べ替えない", _
        (modFollowup.DemoteOrder(ids, 4) = "1,2,3,4"), _
        "実際=" & modFollowup.DemoteOrder(ids, 4)
End Sub

Public Sub RunAll15()
    On Error GoTo ChoiceFail15
    TestParseChoiceNumbers
NextOptions15:
    On Error GoTo OptionsFail15
    TestParseOptions
NextTopic15:
    On Error GoTo TopicFail15
    TestMergeTopicAnswer
NextAsk15:
    On Error GoTo AskFail15
    TestClarifyAsk
NextErrAns15:
    On Error GoTo ErrAnsFail15
    TestBuildErrorAnswer
NextKey15:
    On Error GoTo KeyFail15
    TestParseChunkKey
NextNeighbor15:
    On Error GoTo NeighborFail15
    TestNeighborIdList
NextDemote15:
    On Error GoTo DemoteFail15
    TestDemoteUsed
NextDone15:
    On Error GoTo 0
    Exit Sub

ChoiceFail15:
    modTestRunner.Check "TestParseChoiceNumbers(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextOptions15
OptionsFail15:
    modTestRunner.Check "TestParseOptions(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextTopic15
TopicFail15:
    modTestRunner.Check "TestMergeTopicAnswer(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextAsk15
AskFail15:
    modTestRunner.Check "TestClarifyAsk(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextErrAns15
ErrAnsFail15:
    modTestRunner.Check "TestBuildErrorAnswer(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextKey15
KeyFail15:
    modTestRunner.Check "TestParseChunkKey(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextNeighbor15
NeighborFail15:
    modTestRunner.Check "TestNeighborIdList(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDemote15
DemoteFail15:
    modTestRunner.Check "TestDemoteUsed(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone15
End Sub
