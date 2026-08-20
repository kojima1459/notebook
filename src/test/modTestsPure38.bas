Attribute VB_Name = "modTestsPure38"
Option Explicit

' ============================================================================
' modTestsPure38 - R34波2(層1の5件)の純ロジック回帰テスト。
'   modTestRunner.RunAllPureTests から直接呼ばれる RunAll38 の1本が入口
'   (modTestsPure31〜37 と同型の別枝。既存チェーンへは繋がない)。
' ----------------------------------------------------------------------------
' 【何を守るのか】
'   B1 出典突合を ⚡すぐ聞く / 🔍しっかり調べる にも効かせる:
'     ・modMode.ShouldAnnotate … 走ってよいモードの線引き。ここが緩むと
'       入念モードで「(出典確認できず)」が二重に付き、逆に厳しすぎると
'       quick/deep の幻覚出典が素通りする。
'     ・modMode.CiteTagFrom … 突合表に積むタグの書式。**modPrompts.SourceTag
'       と一字一句同じ**でなければ、正しい出典まで全件「確認できず」になる
'       (最も危険な壊れ方=正しい回答が全部疑わしく見える)。書式そのものは
'       リテラルで、SourceTag との同形性は実際の戻り値との直接比較で固定する
'       (後者は Hit 型を通るため LO では SKIP。下の「LO の死角」を参照)。
'     ・modMode.CiteIndexAdd … 突合表の積み方(正規化・重複排除)。
'     ・上の3本を組んで modAskThorough.AnnotateCitations へ通したときの
'       ふるまい(実在タグ→無加工 / 幻覚タグ→注記 / ヒット0件→無加工)。
'   B3 前後チャンク結合のゲート(R34 F1 で発火点が modAsk.RunDeepFlow へ移った
'   あとも、ゲートそのものはこの関数のまま=テストは無改修で有効):
'     ・modMode.UseNeighborExpand … deep だけ True。thorough で True を返すと
'       modAskThorough/modAskMulti の自前呼び出しと合わせて二重結合になる。
'
' 【LO の死角 = SKIP になる2群(2026-08-20 R34波2 Fix・実測)】
'   LibreOffice は別モジュールで定義した Public Type(ここでは Hit)を関数の
'   引数として受け渡せない。初版はこれを踏まず、
'     TestCiteTagFormat38 → Err=91(Object variable not set)
'     TestRerankExcerpt38 → Err=420(Invalid object reference)
'   で【群まるごと】NG になった。既存の同型前例(modTestsPure2「modPrompts
'   (Hit配列を使う3関数)」/ modTestsPure8「BuildRerankPrompt<->ParseRankOrder
'   往復」/ modTestsPure12「出典突合表(Hit配列)」)と同じ作法で、Hit に触る
'   アサートだけを CanUseTypeArrays38 のガードへ分離した。
'   ・書式そのもの(CiteTagFrom のリテラル4本)は Hit を使わないので LO でも走る。
'     SourceTag との突合だけが SKIP になる ―― 守れる範囲を減らしていない。
'   ・B2 の抜粋700字は Hit 抜きで観測する手段が無いため群ごと SKIP。恒真
'     アサートで見た目のPASS数を水増ししない。
'   Excel 実機ではどちらも走る(§11.3 の受入チェックで必ず確認すること)。
'
' 【守れないもの(正直に書く。恒真アサートで埋めない)】
'   ・modMode.AnnotateIfNeeded 本体。modAsk のモジュール状態(直近ヒット)を
'     読むため、LO からは入力を作れない。中身は上の純関数3本+
'     AnnotateCitations の合成に過ぎず、その合成そのものは
'     TestAnnotatePipeline38 が同じ順序で組み直して固定している。
'     残るのは「modAsk のアクセサが0始まりであること」の1点で、これは
'     modAsk 側のコメント(bas:343-353)と実機観点に委ねる。
'   ・B0(SaveTurnForRestore の移設)は modState への書き込みそのものなので
'     純テスト対象外。移設先で私有状態への依存が無いことはレビューで確認済み
'     (使うのは同居する TrimPairs と modUtil/modState のみ)。
'   ・B2(再ランク抜粋 300→700字)は modPrompts.BuildRerankPrompt の数値1個で、
'     TestRerankExcerpt38 が「700字の本文が途中で切られずに載ること」と
'     「lim を超えたら (以下省略) で打ち切られること」の両方を固定する
'     (ただし LO では下記の Hit 制限により群ごと SKIP。実機でのみ走る)。
'   ・C(max_context_chars 既定 60,000)はビルド時のconfig既定値なので、
'     ビルド自己検証の担当(テスト対象外)。
' ============================================================================

' ---- B1(1): 走ってよいモードの線引き ---------------------------------------
'   discriminate の作り:
'   ・"thorough" を True 側へ入れると「入念は二重付与しない」1本が落ちる。
'   ・Normalize を通す実装へ変えると "" と "general" の2本が落ちる
'     (Normalize は未知の文字列を quick へ丸めるため)。
'   ・LCase/Trim を外すと大小・空白の2本が落ちる。
Private Sub TestShouldAnnotate38()
    ChkBool38 "B1_quickは突合する", modMode.ShouldAnnotate("quick"), True
    ChkBool38 "B1_deepは突合する", modMode.ShouldAnnotate("deep"), True
    ChkBool38 "B1_thoroughは突合しない(生成の内側で済み)", _
        modMode.ShouldAnnotate("thorough"), False
    ChkBool38 "B1_一般アシスタントは突合しない", _
        modMode.ShouldAnnotate("general"), False
    ChkBool38 "B1_空文字(空質問ターン)は突合しない", _
        modMode.ShouldAnnotate(""), False
    ChkBool38 "B1_知らないモード名は突合しない", _
        modMode.ShouldAnnotate("turbo"), False
    ChkBool38 "B1_大文字でも同じ", modMode.ShouldAnnotate("DEEP"), True
    ChkBool38 "B1_前後の空白があっても同じ", modMode.ShouldAnnotate("  quick "), True
End Sub

' ---- B1(2): タグ書式そのもの(Hit型を使わない=LOでも必ず走る) ---------------
'   CiteTagFrom を1文字でも変えると(全角コロン・p.の省略・空白の増減)、
'   正しい出典まで citeIndex と一致しなくなり全件へ「(出典確認できず)」が付く。
'   CiteTagFrom は素の3値を取る純文字列関数なので、ここは Hit 型に一切触れず
'   リテラルで固定できる。**LO の Public Type 制限に左右されない側**。
Private Sub TestCiteTagFormat38()
    ChkStr38 "B1_本棚形の実物", _
        modMode.CiteTagFrom("就業規則.pdf", 12, "shelf"), "[本棚:就業規則.pdf p.12]"
    ChkStr38 "B1_パック形の実物(ページ番号は入らない)", _
        modMode.CiteTagFrom("経費規程.docx", 3, "pack:山田太郎"), _
        "[パック(山田太郎):経費規程.docx]"
    ' origin の大小はどちらの実装も LCase で見る(PACK: でもパック形)。
    ChkStr38 "B1_originの大文字PACK:もパック形", _
        modMode.CiteTagFrom("手順書.xlsx", 1, "PACK:Sales"), "[パック(Sales):手順書.xlsx]"
    ' origin が空(旧データ)なら本棚形。ここが pack 側へ倒れると
    ' ページ番号が落ちて既存の突合まで壊れる。
    ChkStr38 "B1_origin空は本棚形", _
        modMode.CiteTagFrom("議事録.txt", 0, ""), "[本棚:議事録.txt p.0]"
End Sub

' ---- B1(3): 上の書式が modPrompts.SourceTag と一字一句同じであること --------
'   ここが本丸。上の群がリテラルを固定しても、**modPrompts.SourceTag 側だけを
'   変えられたら気付けない**(両方が同じ形であることに意味がある)ので、
'   期待値をリテラルではなく SourceTag の【実際の戻り値】に取って両側を同時に
'   見張る。ただし SourceTag は Hit 型を引数に取るため、LibreOffice の
'   Public Type 制限(別モジュールで定義した Type の受け渡しが Err=91/420 に
'   なる)にかかる。同じ理由の SKIP は modTestsPure2 / 8 / 12 に前例がある。
Private Sub TestCiteTagVsSourceTag38()
    If Not CanUseTypeArrays38() Then
        modTestRunner.Check _
            "[SKIP] 出典タグのSourceTag突合(Hit型): LO環境の既知の制限によりスキップ", True, _
            "modPrompts.SourceTag は Hit 型を引数に取るため、別モジュール定義の " & _
            "Public Type を渡せない環境(LibreOffice)では呼べない。書式そのものは " & _
            "TestCiteTagFormat38 がリテラルで固定済みで、両者が同形であることは " & _
            "コードレビューで確認済み(modPrompts の SourceTag 606-614行 と " & _
            "modMode の CiteTagFrom)。" & _
            "Excel実機受入チェック(§11.3)で必ず再確認すること。"
        Exit Sub
    End If

    Dim h As Hit
    h.source = "就業規則.pdf": h.page = 12: h.origin = "shelf"
    ChkStr38 "B1_本棚形がSourceTagと一致", _
        modMode.CiteTagFrom(h.source, h.page, h.origin), modPrompts.SourceTag(h)

    Dim p As Hit
    p.source = "経費規程.docx": p.page = 3: p.origin = "pack:山田太郎"
    ChkStr38 "B1_パック形がSourceTagと一致", _
        modMode.CiteTagFrom(p.source, p.page, p.origin), modPrompts.SourceTag(p)

    Dim u As Hit
    u.source = "手順書.xlsx": u.page = 1: u.origin = "PACK:Sales"
    ChkStr38 "B1_originの大文字PACK:もSourceTagと一致", _
        modMode.CiteTagFrom(u.source, u.page, u.origin), modPrompts.SourceTag(u)

    Dim z As Hit
    z.source = "議事録.txt": z.page = 0: z.origin = ""
    ChkStr38 "B1_origin空は本棚形でSourceTagと一致", _
        modMode.CiteTagFrom(z.source, z.page, z.origin), modPrompts.SourceTag(z)
End Sub

' modTestsPure2.CanUseTypeArrays / modTestsPure12.CanUseTypeArrays11 と同じ
' 実測プローブ(別モジュールの Private は呼べないため。テスト限定の軽微な
' 重複で、公開契約を増やすより実害が小さい ―― modTestsPure12 の判断を踏襲)。
Private Function CanUseTypeArrays38() As Boolean
    On Error Resume Next
    Err.Clear
    Dim probe() As ShelfChunk
    ReDim probe(0 To 0)
    CanUseTypeArrays38 = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' ---- B1(4): 突合表の積み方(正規化・重複排除) ------------------------------
Private Sub TestCiteIndexAdd38()
    ' 積むときに NormalizeCiteTag を通すので、表の中では空白が落ちた形になる
    ' (照合する側=TagIsKnown も同じ正規化を通るので、この非対称は正しい)。
    Dim idx As String
    idx = modMode.CiteIndexAdd("", "[本棚:A.pdf p.1]")
    ChkStr38 "B1_1件目は|タグ|の形で積む(空白は落ちる)", idx, "|[本棚:A.pdfp.1]|"

    idx = modMode.CiteIndexAdd(idx, "[本棚:B.pdf p.2]")
    ChkStr38 "B1_2件目は連結する", idx, "|[本棚:A.pdfp.1]||[本棚:B.pdfp.2]|"

    ChkStr38 "B1_同じタグは二度積まない", _
        modMode.CiteIndexAdd(idx, "[本棚:A.pdf p.1]"), idx
    ChkStr38 "B1_空白違いも同じタグとして積まない", _
        modMode.CiteIndexAdd(idx, "[本棚: A.pdf  p. 1]"), idx
    ChkStr38 "B1_空タグは積まない", modMode.CiteIndexAdd(idx, ""), idx

    ' 空白入りで積んでも表は正規形(半角・全角・タブのいずれも落ちる)。
    ChkStr38 "B1_積むときに空白を落とす", _
        modMode.CiteIndexAdd("", "[本棚: C.pdf p. 9]"), "|[本棚:C.pdfp.9]|"
End Sub

' ---- B1(5): 3本を組んで AnnotateCitations へ通したときのふるまい ------------
'   AnnotateIfNeeded 本体と同じ順序(CiteTagFrom → CiteIndexAdd →
'   AnnotateCitations)で組み直す。ここが quick/deep の実際の効き目。
Private Sub TestAnnotatePipeline38()
    Dim idx As String
    idx = modMode.CiteIndexAdd(idx, modMode.CiteTagFrom("就業規則.pdf", 12, "shelf"))
    idx = modMode.CiteIndexAdd(idx, modMode.CiteTagFrom("経費規程.docx", 3, "pack:山田太郎"))

    Dim n As Long

    ' (a) 実在タグだけの回答は1文字も変わらない。
    Dim okAns As String
    okAns = "有給は入社半年で付与されます。[本棚:就業規則.pdf p.12]"
    ChkStr38 "B1_実在タグは無加工", _
        modAskThorough.AnnotateCitations(okAns, idx, n), okAns
    ChkLong38 "B1_実在タグの不一致件数は0", n, 0

    ' パック形も同じ表で通ること(書式が片方だけズレていれば落ちる)。
    Dim okPack As String
    okPack = "上限は月3万円です。[パック(山田太郎):経費規程.docx]"
    ChkStr38 "B1_パック形の実在タグも無加工", _
        modAskThorough.AnnotateCitations(okPack, idx, n), okPack
    ChkLong38 "B1_パック形の不一致件数は0", n, 0

    ' (b) 検索結果に無いタグ(幻覚)には注記が付く。
    Dim ngAns As String
    ngAns = "上限は月5万円です。[本棚:就業規則.pdf p.99]"
    ChkStr38 "B1_幻覚タグに注記が付く", _
        modAskThorough.AnnotateCitations(ngAns, idx, n), _
        ngAns & modAskThorough.UNVERIFIED_MARK
    ChkLong38 "B1_幻覚タグの不一致件数は1", n, 1

    ' ページ番号だけ違うものも「別の出典」として扱う(±0の完全一致)。
    ChkStr38 "B1_ページ違いも幻覚として扱う", _
        modAskThorough.AnnotateCitations("[本棚:就業規則.pdf p.13]", idx, n), _
        "[本棚:就業規則.pdf p.13]" & modAskThorough.UNVERIFIED_MARK

    ' (c) ヒット0件のターンは突合表が空=無加工で返る(AnnotateIfNeeded は
    '     そもそも LastHitCount()=0 で手前から抜けるが、万一通っても無害)。
    ChkStr38 "B1_突合表が空なら無加工", _
        modAskThorough.AnnotateCitations(ngAns, "", n), ngAns

    ' 出典タグを1つも含まない回答も無加工(一般的な言い回しへ注記しない)。
    ChkStr38 "B1_タグの無い回答は無加工", _
        modAskThorough.AnnotateCitations("資料からは判断できません。", idx, n), _
        "資料からは判断できません。"
End Sub

' ---- B3: 前後チャンク結合のゲート ------------------------------------------
'   discriminate の作り:
'   ・"thorough" を True 側へ入れると1本が落ちる(=二重結合の検知)。
'   ・"quick" を True 側へ入れると1本が落ちる(GOの範囲外)。
'   ・Normalize を外して素の比較にすると「知らない値」の1本が落ちる。
Private Sub TestUseNeighborExpand38()
    ChkBool38 "B3_deepでだけ前後結合する", modMode.UseNeighborExpand("deep"), True
    ChkBool38 "B3_thoroughは自前で呼ぶので掛けない", _
        modMode.UseNeighborExpand("thorough"), False
    ChkBool38 "B3_quickは対象外", modMode.UseNeighborExpand("quick"), False
    ChkBool38 "B3_空文字はquickへ丸まる", modMode.UseNeighborExpand(""), False
    ChkBool38 "B3_知らない値もquickへ丸まる", modMode.UseNeighborExpand("neighbor"), False
End Sub

' ---- B2: 再ランク候補の抜粋が700字まで載ること -----------------------------
'   300 のままだと「700字の本文が丸ごと載る」1本が落ちる(実際に何字載ったかを
'   数えるので、閾値を700未満へ下げれば必ず落ちる)。
'   lim による打ち切り(BuildRerankPrompt の残量管理)も同時に見る。ここが
'   効いていないと、抜粋を厚くした瞬間にプロンプトが max_context_chars を
'   超える(B2 の安全弁そのもの)。
Private Sub TestRerankExcerpt38()
    ' BuildRerankPrompt は hits() As Hit を取るため、LibreOffice の Public Type
    ' 制限にかかる(modTestsPure8.TestRerankPromptRoundTrip と同じ理由・同じ作法)。
    ' 抜粋長は SafeLeft の第2引数1個で、Hit を経由せずに外から観測する手段が
    ' 無い(SourceBody が Private・プロンプト組み立てが唯一の出口)ため、
    ' この群は丸ごとガードする。恒真アサートでの水増しはしない。
    If Not CanUseTypeArrays38() Then
        modTestRunner.Check _
            "[SKIP] 再ランク抜粋700字(Hit配列): LO環境の既知の制限によりスキップ", True, _
            "modPrompts.BuildRerankPrompt は Hit() 配列を引数に取るため、別モジュール" & _
            "定義の Public Type を渡せない環境(LibreOffice)では呼べない。R34 B2 の" & _
            "変更は modPrompts の BuildRerankPrompt 内 SafeLeft 第2引数 300→700 の" & _
            "1箇所のみ(223行)で、lim による残量管理(225-230行)は無改修。" & _
            "Excel実機受入チェック(§11.3)で必ず再確認すること。"
        Exit Sub
    End If

    Dim hits(1 To 1) As Hit
    hits(1).source = "規程.pdf"
    hits(1).page = 5
    hits(1).origin = "shelf"
    ' 目印の文字は「ゑ」。BuildRerankPrompt の定型文・出典タグ・質問文の
    ' どこにも現れないので、数えた個数がそのまま抜粋の長さになる
    ' (「あ」は定型文の先頭「あなたは…」に出るので使えない)。
    hits(1).full_text = String$(900, "ゑ")

    Dim s As String
    s = modPrompts.BuildRerankPrompt("上限は?", hits, 1, 60000)
    ChkLong38 "B2_抜粋は700字まで載る", CountChar38(s, "ゑ"), 700

    ' lim が小さいときは候補を積まずに (以下省略) で打ち切る(残量管理は不変)。
    Dim t As String
    t = modPrompts.BuildRerankPrompt("上限は?", hits, 1, 200)
    ChkLong38 "B2_limを超える候補は積まない", CountChar38(t, "ゑ"), 0
    ChkBool38 "B2_打ち切ったら(以下省略)を書く", (InStr(1, t, "(以下省略)") > 0), True
End Sub

' 指定した1文字が何個含まれるかを数える(Replace の長さ差で数える定石)。
Private Function CountChar38(ByVal s As String, ByVal c As String) As Long
    CountChar38 = Len(s) - Len(Replace(s, c, ""))
End Function

Private Sub ChkLong38(ByVal label As String, ByVal got As Long, ByVal want As Long)
    modTestRunner.Check "R34-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkBool38(ByVal label As String, ByVal got As Boolean, ByVal want As Boolean)
    modTestRunner.Check "R34-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkStr38(ByVal label As String, ByVal got As String, ByVal want As String)
    modTestRunner.Check "R34-" & label, (StrComp(got, want, vbBinaryCompare) = 0), _
        "実際=[" & got & "] 期待=[" & want & "]"
End Sub

' 群ごとに独立のハンドラを持たせるのは R33 の教訓(単一の巨大ハンドラだと
' 1群の例外で以降が丸ごと無言で消える)。Hit 型に触る2群は群の【先頭】で
' CanUseTypeArrays38 のガードを通し、例外ではなく SKIP として集計へ載せる。
Public Sub RunAll38()
    On Error GoTo H01Fail38
    TestShouldAnnotate38
H02Next38:
    On Error GoTo H02Fail38
    TestCiteTagFormat38
H03Next38:
    On Error GoTo H03Fail38
    TestCiteTagVsSourceTag38
H04Next38:
    On Error GoTo H04Fail38
    TestCiteIndexAdd38
H05Next38:
    On Error GoTo H05Fail38
    TestAnnotatePipeline38
H06Next38:
    On Error GoTo H06Fail38
    TestUseNeighborExpand38
H07Next38:
    On Error GoTo H07Fail38
    TestRerankExcerpt38
H01Done38:
    On Error GoTo 0
    Exit Sub

H01Fail38:
    modTestRunner.Check "TestShouldAnnotate38(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H02Next38
H02Fail38:
    modTestRunner.Check "TestCiteTagFormat38(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H03Next38
H03Fail38:
    modTestRunner.Check "TestCiteTagVsSourceTag38(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H04Next38
H04Fail38:
    modTestRunner.Check "TestCiteIndexAdd38(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H05Next38
H05Fail38:
    modTestRunner.Check "TestAnnotatePipeline38(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H06Next38
H06Fail38:
    modTestRunner.Check "TestUseNeighborExpand38(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H07Next38
H07Fail38:
    modTestRunner.Check "TestRerankExcerpt38(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done38
End Sub
