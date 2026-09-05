Attribute VB_Name = "modTestsPure42"
Option Explicit

' ============================================================================
' modTestsPure42 - R37 波B(資料間リンク)の純ロジック回帰テスト。
'   既存チェーン(modTestsPure.RunAll→…)へは繋がず、
'   modTestRunner.RunAllPureTests から直接呼ばれる RunAll42 の1本が入口
'   (modTestsPure31〜41 と同型の別枝)。
' ----------------------------------------------------------------------------
' 【何を固定するか】
'   A modXDoc.CosineCsv        : 同一=1 / 直交=0 / 次元不一致=0 / 未正規化でも
'                                コサインになる(正規化を忘れた実装を落とす)。
'   B modXDocStore.MeanNormalizedCsv : 2本の平均・長さ1・次元違いは数に入れない。
'   C modXDocStore.TopNLinks   : 上位3件・同点先勝ち・keep>n・keep<=0。
'   D modXDoc.PickLinked       : 閾値(xdoc_min_sim)・自資料除外・最大N・
'                                起点キー不一致は無視・重複除去。
'   E modXDoc.ChapterOf        : 章キーの取り方(第1要素・目次行の末尾正規化・
'                                リーダー1字は触らない)。期待値は定数で書く。
'
' 【なぜ RememberPool/Expand/BuildFor のテストが無いか】
'   RememberPool/Expand は Hit 型(Public Type)を跨ぐため LibreOffice の
'   実行テストからは呼べない(modCorrect.InjectHits/PrependHit と同じ既知の
'   死角)。BuildFor/シートI/O は Worksheet を持つ。どちらも R37 §3-2 の
'   実機受入(①doc_links に行ができ📖にメタ行が出る ②A の言葉で聞いて B が
'   末尾に出る ③無関係な資料が混ざらない)に委ねる。
' ============================================================================

' Double の比較幅。重心CSVは6桁丸めで保存するので、正規化してから丸めた
' ベクトルの長さは 1 から最大で 1e-6 程度ずれる。1e-7 では「丸めのせい」で
' 落ちるので 1e-5 を取る(このテストが見分けたい違い ―― 0.5 と 0.707107、
' 内積 3 とコサイン 1 ―― はどれも桁が違うので鈍りはしない)。
Private Const EPS42 As Double = 0.00001

' ---- A: CosineCsv ------------------------------------------------------------
'   discriminate:
'   ・分母(ノルム)で割り忘れた実装は (e) が 3 を返して落ちる。
'   ・次元不一致を 0 にせず短いほうに合わせる実装は (c) が 1 を返して落ちる。
'   ・ゼロベクトルを 0/0 のまま計算する実装は (d) で例外か NaN になり落ちる。
Private Sub TestCosineCsv42()
    ' (a) 同一ベクトル=1。
    ChkNear42 "A_同一は1", modXDoc.CosineCsv("0.6,0.8", "0.6,0.8"), 1#

    ' (b) 直交=0。
    ChkNear42 "A_直交は0", modXDoc.CosineCsv("1,0", "0,1"), 0#

    ' (c) 次元不一致=0(短いほうに揃えない)。
    ChkNear42 "A_次元不一致は0", modXDoc.CosineCsv("1,0,0", "1,0"), 0#

    ' (d) ゼロベクトル=0。
    ChkNear42 "A_ゼロベクトルは0", modXDoc.CosineCsv("0,0", "1,0"), 0#

    ' (e) 未正規化でもコサイン(内積 3 ではなく 1)。
    ChkNear42 "A_未正規化でもコサイン", modXDoc.CosineCsv("3,0", "1,0"), 1#

    ' (f) 空文字=0(読めないCSVは「近くない」を返す)。
    ChkNear42 "A_空は0", modXDoc.CosineCsv("", "1,0"), 0#

    ' (g) 正規化済み2本では「内積=CosineCsv」。doc_centroids は書込み時に
    '     L2正規化しており、取込の総当たり(modXDocBuild.MatchChapters)は
    '     その性質を使って内積だけを回す。その等価性をここで固定する。
    Dim v1() As Double, v2() As Double
    Dim ok1 As Boolean: ok1 = modUtil.CsvToVector("0.6,0.8", v1)
    Dim ok2 As Boolean: ok2 = modUtil.CsvToVector("0.8,0.6", v2)
    ChkBool42 "A_2本ともパースできる", (ok1 And ok2), True
    ChkNear42 "A_単位ベクトル同士の値", modXDoc.CosineCsv("0.6,0.8", "0.8,0.6"), 0.96
    ChkNear42 "A_正規化済みなら内積=コサイン", _
        modXDoc.CosineCsv("0.6,0.8", "0.8,0.6"), modUtil.DotProduct(v1, v2)
End Sub

' ---- B: MeanNormalizedCsv ----------------------------------------------------
'   discriminate:
'   ・平均を取らず先頭行をそのまま返す実装は (a) の第1成分が 1 になって落ちる。
'   ・L2正規化を忘れた実装は (a) が 0.5 のままで落ちる(長さ 0.7071 ではない)。
'   ・次元の違う行を混ぜて足す実装は (b) の成分数が 3 になって落ちる。
'   ・ゼロベクトル/壊れたCSVに空文字を返さない実装は (c)(d) が落ちる。
Private Sub TestMeanNormalizedCsv42()
    ' (a) (1,0) と (0,1) の平均は (0.5,0.5)、正規化して (0.7071,0.7071)。
    Dim rows2(1 To 2) As String
    rows2(1) = "1,0"
    rows2(2) = "0,1"
    Dim got As String: got = modXDocStore.MeanNormalizedCsv(rows2, 2)
    Dim v() As Double
    Dim okA As Boolean: okA = modUtil.CsvToVector(got, v)
    ChkBool42 "B_平均が読める", okA, True
    ChkLong42 "B_成分数は2", UBound(v) - LBound(v) + 1, 2
    ChkNear42 "B_第1成分は0.707107", v(LBound(v)), 0.707107
    ChkNear42 "B_第2成分は0.707107", v(LBound(v) + 1), 0.707107
    ChkNear42 "B_長さは1", Sqr(modUtil.DotProduct(v, v)), 1#

    ' (b) 次元の違う行は数に入れない(最初に読めた行の次元で確定)。
    '     2本目を "0,1,0" にしてあるのは意図的。"1,0,0" だと、次元違いを
    '     混ぜて足す誤実装でも足される成分が (1,0) で、平均→正規化のあと
    '     結果が (1,0) のまま一致してしまい見分けられない(ネガティブ確認で
    '     実測)。"0,1,0" なら混ぜた瞬間 (0.707,0.707) になって落ちる。
    Dim rows3(1 To 2) As String
    rows3(1) = "1,0"
    rows3(2) = "0,1,0"
    Dim got2 As String: got2 = modXDocStore.MeanNormalizedCsv(rows3, 2)
    Dim v2() As Double
    Dim okB As Boolean: okB = modUtil.CsvToVector(got2, v2)
    ChkBool42 "B_次元違い混在でも読める", okB, True
    ChkLong42 "B_次元違いは数に入れない", UBound(v2) - LBound(v2) + 1, 2
    ChkNear42 "B_残った1本がそのまま", v2(LBound(v2)), 1#

    ' (c) 全部ゼロ → 空文字(重心が作れない)。
    Dim rowsZ(1 To 1) As String
    rowsZ(1) = "0,0"
    ChkStr42 "B_ゼロベクトルは空", modXDocStore.MeanNormalizedCsv(rowsZ, 1), ""

    ' (d) 壊れたCSV → 空文字。
    Dim rowsB(1 To 1) As String
    rowsB(1) = "abc"
    ChkStr42 "B_壊れたCSVは空", modXDocStore.MeanNormalizedCsv(rowsB, 1), ""

    ' (e) n=0 → 空文字。
    ChkStr42 "B_0本は空", modXDocStore.MeanNormalizedCsv(rows2, 0), ""
End Sub

' ---- C: TopNLinks ------------------------------------------------------------
'   discriminate:
'   ・昇順に並べる実装は (a) の1行目が A(0.1)になって落ちる。
'   ・keep を無視して全件返す実装は (a) の行数が 5 になって落ちる。
'   ・同点で後から来たほうを採る実装(比較に >= を使う実装)は (b) が
'     "P2|P3" になって落ちる。
'   ・keep>n を丸めない実装は (c) で添字の範囲外になる。
Private Sub TestTopNLinks42()
    Dim p(1 To 5) As String
    Dim s(1 To 5) As Double
    p(1) = "A|1章": s(1) = 0.1
    p(2) = "B|1章": s(2) = 0.9
    p(3) = "C|1章": s(3) = 0.5
    p(4) = "D|1章": s(4) = 0.7
    p(5) = "E|1章": s(5) = 0.3

    ' (a) 上位3件が sim 降順で返る。
    Dim got As String: got = modXDocStore.TopNLinks(p, s, 5, 3)
    Dim ln() As String: ln = Split(got, vbLf)
    ChkLong42 "C_上位3件", UBound(ln) - LBound(ln) + 1, 3
    ChkStr42 "C_1位はB", PairOf42(ln(0)), "B|1章"
    ChkStr42 "C_2位はD", PairOf42(ln(1)), "D|1章"
    ChkStr42 "C_3位はC", PairOf42(ln(2)), "C|1章"
    ChkNear42 "C_1位のsimが行に載る", SimOf42(ln(0)), 0.9

    ' (b) 同点は先に並んでいたほうが勝つ。
    Dim q(1 To 3) As String
    Dim t(1 To 3) As Double
    q(1) = "P1|1章": t(1) = 0.5
    q(2) = "P2|1章": t(2) = 0.5
    q(3) = "P3|1章": t(3) = 0.5
    Dim got2 As String: got2 = modXDocStore.TopNLinks(q, t, 3, 2)
    Dim ln2() As String: ln2 = Split(got2, vbLf)
    ChkLong42 "C_同点でも2件", UBound(ln2) - LBound(ln2) + 1, 2
    ChkStr42 "C_同点は先勝ち1", PairOf42(ln2(0)), "P1|1章"
    ChkStr42 "C_同点は先勝ち2", PairOf42(ln2(1)), "P2|1章"

    ' (c) keep が件数より多いときは件数ぶん。
    Dim got3 As String: got3 = modXDocStore.TopNLinks(q, t, 3, 99)
    Dim ln3() As String: ln3 = Split(got3, vbLf)
    ChkLong42 "C_keep過大は件数ぶん", UBound(ln3) - LBound(ln3) + 1, 3

    ' (d) keep<=0 / n<=0 は空文字。
    ChkStr42 "C_keep0は空", modXDocStore.TopNLinks(q, t, 3, 0), ""
    ChkStr42 "C_n0は空", modXDocStore.TopNLinks(q, t, 0, 3), ""
End Sub

' ---- D: PickLinked -----------------------------------------------------------
'   linksCsv の1行 = "src_a|chap_a|src_b|chap_b|sim"
'   discriminate:
'   ・閾値を見ない実装は (a) が 0.75 の行まで拾って落ちる。
'   ・自資料(起点hitsの資料)を除外しない実装は (b) が A を拾って落ちる。
'   ・maxN を無視する実装は (c) の件数が 3 になって落ちる。
'   ・起点キー(src_a|chap_a)を照合しない実装は (d) が拾って落ちる。
'   ・重複を落とさない実装は (e) が2件返して落ちる。
Private Sub TestPickLinked42()
    Dim seeds(1 To 1) As String
    seeds(1) = "A|第1章"

    ' (a) 閾値: 0.85 は通り 0.75 は通らない(既定 minSim=80)。
    Dim csv1 As String
    csv1 = "A|第1章|B|第3章|0.85" & vbLf & "A|第1章|C|第2章|0.75"
    ChkStr42 "D_閾値未満は落とす", modXDoc.PickLinked(seeds, csv1, 80, 5), "B|第3章"

    ' (a2) 閾値ちょうど(0.80)は通す(二進小数の丸め落ちで消えない)。
    Dim csv1b As String: csv1b = "A|第1章|B|第3章|0.8"
    ChkStr42 "D_閾値ちょうどは通す", modXDoc.PickLinked(seeds, csv1b, 80, 5), "B|第3章"

    ' (b) 自資料は除外(同じ資料の別の章は「Aで聞いてBに届く」にならない)。
    Dim csv2 As String
    csv2 = "A|第1章|A|第9章|0.99" & vbLf & "A|第1章|B|第3章|0.90"
    ChkStr42 "D_自資料は除外", modXDoc.PickLinked(seeds, csv2, 80, 5), "B|第3章"

    ' (c) 最大N。
    Dim csv3 As String
    csv3 = "A|第1章|B|第3章|0.95" & vbLf & "A|第1章|C|第2章|0.90" & vbLf & _
           "A|第1章|D|第1章|0.85"
    ChkStr42 "D_最大N", modXDoc.PickLinked(seeds, csv3, 80, 2), "B|第3章" & vbLf & "C|第2章"

    ' (d) 起点キーに一致しない行は無視。
    Dim csv4 As String: csv4 = "Z|第1章|B|第3章|0.95"
    ChkStr42 "D_起点キー不一致は無視", modXDoc.PickLinked(seeds, csv4, 80, 5), ""

    ' (e) 同じ(資料|章)は1回だけ。
    Dim csv5 As String
    csv5 = "A|第1章|B|第3章|0.95" & vbLf & "A|第1章|B|第3章|0.90"
    ChkStr42 "D_重複は1回だけ", modXDoc.PickLinked(seeds, csv5, 80, 5), "B|第3章"

    ' (f) maxN<=0 / 空のlinksCsv は空文字。
    ChkStr42 "D_maxN0は空", modXDoc.PickLinked(seeds, csv3, 80, 0), ""
    ChkStr42 "D_空linksは空", modXDoc.PickLinked(seeds, "", 80, 5), ""
End Sub

' ---- E: ChapterOf ------------------------------------------------------------
'   discriminate(期待値は modOutlineBuild.ChapterKeyOf の振る舞いをそのまま
'   書き下した【定数】で、あちらを呼び直して突き合わせる恒真比較にはしない):
'   ・">" の前だけを取らない実装は (a) が条まで含んで落ちる。
'   ・目次行の末尾(リーダー記号+頁番号)を落とさない自前の式を書いた実装は
'     (c) が "第3章 総則……12" を返して落ちる ―― これが「章要約の章と
'     リンクの章が静かに食い違う」事故を機械で止める線。
Private Sub TestChapterOf42()
    ' (a) section_path の第1要素。
    ChkStr42 "E_第1要素を取る", modXDoc.ChapterOf("第3章 総則>第12条(免責)"), "第3章 総則"

    ' (b) 空は空(章が分からない、という正しい答え)。
    ChkStr42 "E_空は空", modXDoc.ChapterOf(""), ""

    ' (c) 目次行(「第3章 総則……12」)も本文見出しと同じキーへ寄る。
    Dim toc As String: toc = "第3章 総則" & ChrW(&H2026) & ChrW(&H2026) & "12"
    ChkStr42 "E_目次行の末尾を落とす", modXDoc.ChapterOf(toc), "第3章 総則"

    ' (d) リーダー記号が1字だけなら触らない(章名を誤って削らない保守側)。
    ChkStr42 "E_リーダー1字は触らない", modXDoc.ChapterOf("第3章 総則.12"), "第3章 総則.12"
End Sub

' ---- 判定ヘルパー -----------------------------------------------------------
' TopNLinks の1行 "src|chap|sim" から "src|chap" だけを取る。
Private Function PairOf42(ByVal line0 As String) As String
    Dim p As Long: p = InStrRev(line0, "|")
    If p < 2 Then Exit Function
    PairOf42 = Left$(line0, p - 1)
End Function

' TopNLinks の1行 "src|chap|sim" から sim を取る(Str$の書式に依存しない)。
Private Function SimOf42(ByVal line0 As String) As Double
    Dim p As Long: p = InStrRev(line0, "|")
    If p < 1 Then Exit Function
    SimOf42 = Val(Mid$(line0, p + 1))
End Function

Private Sub ChkBool42(ByVal label As String, ByVal got As Boolean, ByVal want As Boolean)
    modTestRunner.Check "R37-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkLong42(ByVal label As String, ByVal got As Long, ByVal want As Long)
    modTestRunner.Check "R37-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkStr42(ByVal label As String, ByVal got As String, ByVal want As String)
    modTestRunner.Check "R37-" & label, (StrComp(got, want, vbBinaryCompare) = 0), _
        "実際=[" & Replace(got, vbLf, "\n") & "] 期待=[" & Replace(want, vbLf, "\n") & "]"
End Sub

Private Sub ChkNear42(ByVal label As String, ByVal got As Double, ByVal want As Double)
    modTestRunner.Check "R37-" & label, (Abs(got - want) < EPS42), _
        "実際=" & got & " 期待=" & want & " (許容" & EPS42 & ")"
End Sub

Public Sub RunAll42()
    On Error GoTo H01Fail42
    TestCosineCsv42
H02Next42:
    On Error GoTo H02Fail42
    TestMeanNormalizedCsv42
H03Next42:
    On Error GoTo H03Fail42
    TestTopNLinks42
H04Next42:
    On Error GoTo H04Fail42
    TestPickLinked42
H05Next42:
    On Error GoTo H05Fail42
    TestChapterOf42
H01Done42:
    On Error GoTo 0
    Exit Sub

H01Fail42:
    modTestRunner.Check "TestCosineCsv42(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H02Next42
H02Fail42:
    modTestRunner.Check "TestMeanNormalizedCsv42(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H03Next42
H03Fail42:
    modTestRunner.Check "TestTopNLinks42(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H04Next42
H04Fail42:
    modTestRunner.Check "TestPickLinked42(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H05Next42
H05Fail42:
    modTestRunner.Check "TestChapterOf42(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done42
End Sub
