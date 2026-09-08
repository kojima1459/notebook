Attribute VB_Name = "modTestsPure43"
Option Explicit

' ============================================================================
' modTestsPure43 - R38 入念1回読み(modAskOnePass)の純関数回帰テスト。
'   modTestRunner.RunAllPureTests から呼ばれる RunAll43 が単独入口。
' ----------------------------------------------------------------------------
' 【何を固定するか】
'   A modAskOnePass.IsOn:
'     空・on・true・1・yes(前後空白・大小無視)は True。
'     off・false・0・no(前後空白・大小無視)は False。
'     未定義文字列は既定の True 側へ倒すこと。
'   B modAskOnePass.DedupeKeys:
'     初出順の維持、大文字小文字の同一視、空行の排除、指定 maxN での切り詰め、
'     maxN<1 の空返却、前後空白トリム、先頭末尾に余分な vbLf を付けないこと。
'   C modAskOnePass.PerChapterCap:
'     total \ nChap の等分値、minCap 保証(minCap 未満なら minCap)、
'     nChap<1 で total を返すこと、total<1 で minCap を返すこと。
'   D modAskOnePass.CountLines:
'     空文字で 0、単一行で 1、空行混じりの複数行で非空行数のみを正しく数えること。
'   E modAskOnePass.BuildOnePassPrompt:
'     各節の出現順序(役割 -> 手順 -> 規則 -> 履歴 -> 本文 -> 重点 -> 質問 -> 深掘り)、
'     ansTags=True/False での手順節の有無、strictG=True/False での厳守節の有無、
'     history の空/非空での節挿入、focusBlock の空/非空での節挿入、
'     FOLLOWUP 指定行が末尾にあること、質問本文が抜粋本文の後にあること。
'
' 【どう壊すと落ちるか (discriminate)】
'   ・IsOn で空文字を False にすると A_空はTrue が落ちる。
'   ・IsOn で "OFF " などの前後空白をトリムし忘れると A_off空白が落ちる。
'   ・IsOn で未知の設定値を False に倒すと A_未知値はTrue が落ちる。
'   ・DedupeKeys で初出順を崩したりソートしたりすると B_初出順維持 が落ちる。
'   ・DedupeKeys で大文字小文字を別物と扱うと B_大小無視重複 が落ちる。
'   ・DedupeKeys で maxN 切り詰めを忘れると B_maxN切り詰め が落ちる。
'   ・PerChapterCap で minCap の下限保護を省くと C_minCap下限保護 が落ちる。
'   ・BuildOnePassPrompt で質問を本文の前に置くと E_質問は本文の後 が落ちる。
'   ・BuildOnePassPrompt で ansTags=False なのに手順を出すと E_ansTags偽で手順なし が落ちる。
'   ・BuildOnePassPrompt で strictG=False なのに厳守を出すと E_strictG偽で厳守なし が落ちる。
'
' 【なぜ TryOnePass / SourceBlockOf のテストが無いか】
'   Hit 型(Public Type)の配列渡しを伴う関数は、LibreOffice 実行テスト環境の
'   モジュール境界制限に抵触するため純テストから除外する(仕様 §2-4/§4)。
'   これらの複合動作は実機受入テスト(§5)にて担保する。
' ============================================================================

Private Sub ChkBool43(ByVal label As String, ByVal got As Boolean, ByVal want As Boolean)
    modTestRunner.Check "R38-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkLong43(ByVal label As String, ByVal got As Long, ByVal want As Long)
    modTestRunner.Check "R38-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkStr43(ByVal label As String, ByVal got As String, ByVal want As String)
    modTestRunner.Check "R38-" & label, (StrComp(got, want, vbBinaryCompare) = 0), _
        "実際=[" & Replace(got, vbLf, "\n") & "] 期待=[" & Replace(want, vbLf, "\n") & "]"
End Sub

' ---- A: IsOn -----------------------------------------------------------------
Private Sub TestIsOn43()
    ' (a) 空文字は True(既定有効)
    ChkBool43 "A_空はTrue", modAskOnePass.IsOn(""), True

    ' (b) "on" は True
    ChkBool43 "A_小文字on", modAskOnePass.IsOn("on"), True

    ' (c) "ON " 大小・前後空白混じりは True
    ChkBool43 "A_大文字空白ON", modAskOnePass.IsOn("  ON  "), True

    ' (d) "true", "1", "yes" は True
    ChkBool43 "A_true文字列", modAskOnePass.IsOn("true"), True
    ChkBool43 "A_数字1", modAskOnePass.IsOn("1"), True
    ChkBool43 "A_yes文字列", modAskOnePass.IsOn("yes"), True

    ' (e) "off", "false", "0", "no" は False
    ChkBool43 "A_小文字off", modAskOnePass.IsOn("off"), False
    ChkBool43 "A_off空白", modAskOnePass.IsOn("  OFF "), False
    ChkBool43 "A_false文字列", modAskOnePass.IsOn("FALSE"), False
    ChkBool43 "A_数字0", modAskOnePass.IsOn("0"), False
    ChkBool43 "A_no文字列", modAskOnePass.IsOn("No"), False

    ' (f) 未定義の文字列は既定(True)に倒す
    ChkBool43 "A_未知値はTrue", modAskOnePass.IsOn("hoge"), True
End Sub

' ---- B: DedupeKeys -----------------------------------------------------------
Private Sub TestDedupeKeys43()
    ' (a) 初出順の維持
    Dim raw1 As String: raw1 = "規程A::第1章" & vbLf & "規程B::第2章" & vbLf & "規程C::第3章"
    ChkStr43 "B_初出順維持", modAskOnePass.DedupeKeys(raw1, 5), _
        "規程A::第1章" & vbLf & "規程B::第2章" & vbLf & "規程C::第3章"

    ' (b) 大小無視・重複排除
    Dim raw2 As String: raw2 = "doc::chap1" & vbLf & "DOC::CHAP1" & vbLf & "doc::chap2"
    ChkStr43 "B_大小無視重複", modAskOnePass.DedupeKeys(raw2, 5), _
        "doc::chap1" & vbLf & "doc::chap2"

    ' (c) 空行の無視
    Dim raw3 As String: raw3 = vbLf & "A::1" & vbLf & vbLf & "B::2" & vbLf
    ChkStr43 "B_空行無視", modAskOnePass.DedupeKeys(raw3, 5), "A::1" & vbLf & "B::2"

    ' (d) maxN による切り詰め
    Dim raw4 As String: raw4 = "A::1" & vbLf & "B::2" & vbLf & "C::3"
    ChkStr43 "B_maxN切り詰め", modAskOnePass.DedupeKeys(raw4, 2), "A::1" & vbLf & "B::2"

    ' (e) maxN < 1 は空
    ChkStr43 "B_maxNゼロは空", modAskOnePass.DedupeKeys(raw4, 0), ""
    ChkStr43 "B_maxN負は空", modAskOnePass.DedupeKeys(raw4, -1), ""

    ' (f) 前後空白のトリム
    Dim raw5 As String: raw5 = "  A::1  " & vbLf & "B::2"
    ChkStr43 "B_前後空白トリム", modAskOnePass.DedupeKeys(raw5, 5), "A::1" & vbLf & "B::2"

    ' (g) 先頭末尾に余分な vbLf が無いこと
    Dim gotG As String: gotG = modAskOnePass.DedupeKeys("X::1", 3)
    ChkBool43 "B_先頭末尾vbLf無し", (InStr(gotG, vbLf) = 0 And gotG = "X::1"), True
End Sub

' ---- C: PerChapterCap --------------------------------------------------------
Private Sub TestPerChapterCap43()
    ' (a) 正常等分
    ChkLong43 "C_正常等分", modAskOnePass.PerChapterCap(30000, 3, 2000), 10000

    ' (b) minCap 下限保護
    ChkLong43 "C_minCap下限保護", modAskOnePass.PerChapterCap(3000, 3, 2000), 2000

    ' (c) nChap < 1 は total
    ChkLong43 "C_章数ゼロはtotal", modAskOnePass.PerChapterCap(50000, 0, 2000), 50000

    ' (d) total < 1 は minCap
    ChkLong43 "C_totalゼロはminCap", modAskOnePass.PerChapterCap(0, 5, 2000), 2000

    ' (e) 端数切り捨て確認
    ChkLong43 "C_端数切り捨て", modAskOnePass.PerChapterCap(10005, 2, 2000), 5002
End Sub

' ---- D: CountLines -----------------------------------------------------------
Private Sub TestCountLines43()
    ' (a) 空文字は 0
    ChkLong43 "D_空文字は0", modAskOnePass.CountLines(""), 0

    ' (b) 単一行は 1
    ChkLong43 "D_単一行は1", modAskOnePass.CountLines("A::1"), 1

    ' (c) 空行混じりのカウント
    Dim box As String: box = vbLf & "A::1" & vbLf & "  " & vbLf & "B::2" & vbLf
    ChkLong43 "D_空行除外カウント", modAskOnePass.CountLines(box), 2
End Sub

' ---- E: BuildOnePassPrompt ---------------------------------------------------
Private Sub TestBuildOnePassPrompt43()
    Dim pFull As String
    pFull = modAskOnePass.BuildOnePassPrompt("私の質問", "CTX_TEXT", "FOCUS_TEXT", "HIST_TEXT", "日本語", True, True)

    ' (a) 節の出現順序の固定(InStr位置比較)
    Dim posRole As Long: posRole = InStr(pFull, "あなたは社内資料の調査アシスタントです")
    Dim posProc As Long: posProc = InStr(pFull, "<thinking>")
    Dim posRule As Long: posRule = InStr(pFull, "・根拠にした箇所には必ず出典タグ")
    Dim posHist As Long: posHist = InStr(pFull, "## これまでの会話")
    Dim posCtx As Long: posCtx = InStr(pFull, "## 本棚抜粋(関連する章の本文・文書順)")
    Dim posFoc As Long: posFoc = InStr(pFull, "## 重点抜粋(検索で最も近かった箇所。上の本文と重複あり)")
    Dim posQ As Long: posQ = InStr(pFull, "## 質問")
    Dim posFoll As Long: posFoll = InStr(pFull, "[[FOLLOWUP: 候補1 | 候補2]]")

    Dim orderOk As Boolean
    orderOk = (posRole > 0) And (posProc > posRole) And (posRule > posProc) And _
              (posHist > posRule) And (posCtx > posHist) And (posFoc > posCtx) And _
              (posQ > posFoc) And (posFoll > posQ)
    ChkBool43 "E_各節の並び順序固定", orderOk, True

    ' (b) ansTags=False で手順(<thinking>)が出ないこと
    Dim pNoTags As String
    pNoTags = modAskOnePass.BuildOnePassPrompt("Q", "C", "F", "H", "日本語", False, True)
    ChkBool43 "E_ansTags偽で手順なし", (InStr(pNoTags, "<thinking>") = 0), True

    ' (c) ansTags=True で手順(<thinking>)が出ること
    ChkBool43 "E_ansTags真で手順あり", (InStr(pFull, "<thinking>") > 0), True

    ' (d) strictG=False で厳守節が出ないこと
    Dim pNoStrict As String
    pNoStrict = modAskOnePass.BuildOnePassPrompt("Q", "C", "F", "H", "日本語", True, False)
    ChkBool43 "E_strictG偽で厳守なし", (InStr(pNoStrict, "【厳守】本棚抜粋に書かれた情報のみで回答し") = 0), True

    ' (e) strictG=True で厳守節が出ること
    ChkBool43 "E_strictG真で厳守あり", (InStr(pFull, "【厳守】本棚抜粋に書かれた情報のみで回答し") > 0), True

    ' (f) history 空で会話履歴の見出しが出ないこと
    Dim pNoHist As String
    pNoHist = modAskOnePass.BuildOnePassPrompt("Q", "C", "F", "", "日本語", True, True)
    ChkBool43 "E_history空で履歴節なし", (InStr(pNoHist, "## これまでの会話") = 0), True

    ' (g) focusBlock 空で重点抜粋の見出しが出ないこと
    Dim pNoFocus As String
    pNoFocus = modAskOnePass.BuildOnePassPrompt("Q", "C", "", "H", "日本語", True, True)
    ChkBool43 "E_focus空で重点節なし", (InStr(pNoFocus, "## 重点抜粋") = 0), True

    ' (h) FOLLOWUP 行が末尾に存在すること
    Dim endFollow As Boolean
    endFollow = (Right$(pFull, Len("[[FOLLOWUP: 候補1 | 候補2]]")) = "[[FOLLOWUP: 候補1 | 候補2]]")
    ChkBool43 "E_FOLLOWUP末尾配置", endFollow, True

    ' (j) R38 Fix2(2周目 MAJOR-1): 見出し記号 ■ が出力規則に残っていること。
    '     表示側は行頭 ■ の段落だけを見出しにするので、■ を落とす実装は落ちる。
    ChkBool43 "E_見出し記号■が残る", (InStr(pFull, "・■見出しで構造化") > 0), True

    ' (i) 質問本文が抜粋本文の後ろにあること
    Dim pCtxPos As Long: pCtxPos = InStr(pFull, "CTX_TEXT")
    Dim pQPos As Long: pQPos = InStr(pFull, "私の質問")
    ChkBool43 "E_質問は本文の後", (pQPos > pCtxPos And pCtxPos > 0), True
End Sub

' ----------------------------------------------------------------------------
' RunAll43 - modTestRunner から呼ばれる総合エントリーポイント
' ----------------------------------------------------------------------------
' ---- F: R40 F2 CiteTagSpans / F3 PageLabel・CellAddressOf --------------------
Private Sub TestR40Display43()
    Dim st() As Long, ln() As Long
    Dim s As String
    ' 2つのタグ(本棚・パック)。位置は1始まり、長さは "[" から "]" まで。
    s = "上限はない。[本棚:約款.pdf p.12] また" & vbCr & "別紙参照 [パック(山田):別紙.docx]。"
    ChkLong43 "R40F2_2件", modLive.CiteTagSpans(s, st, ln), 2
    ChkLong43 "R40F2_1件目開始", st(1), InStr(s, "[本棚:")
    ChkStr43 "R40F2_1件目切り出し", Mid$(s, st(1), ln(1)), "[本棚:約款.pdf p.12]"
    ChkStr43 "R40F2_2件目切り出し", Mid$(s, st(2), ln(2)), "[パック(山田):別紙.docx]"
    ' 閉じ括弧が段落を跨ぐものはタグではない(次の候補は拾う)。
    s = "[本棚:壊れ" & vbCr & "た] 本文 [本棚:a.pdf p.1]"
    ChkLong43 "R40F2_段落跨ぎは除外", modLive.CiteTagSpans(s, st, ln), 1
    ChkStr43 "R40F2_跨ぎ後の正常タグ", Mid$(s, st(1), ln(1)), "[本棚:a.pdf p.1]"
    ' タグ無し・空文字は0件。閉じ括弧が無ければ0件。
    ChkLong43 "R40F2_無し", modLive.CiteTagSpans("出典なし [[FOLLOWUP: a | b]]", st, ln), 0
    ChkLong43 "R40F2_空", modLive.CiteTagSpans("", st, ln), 0
    ChkLong43 "R40F2_閉じ無し", modLive.CiteTagSpans("[本棚:a.pdf p.1", st, ln), 0
    ' 隣接するタグも別々に拾う。
    ChkLong43 "R40F2_隣接2件", modLive.CiteTagSpans("[本棚:a.pdf p.1][本棚:b.pdf p.2]", st, ln), 2
    ' 同じ段落内で閉じていない開始は、後続タグの ] に食いつかない(レビュー R40 m1)。
    s = "参考は [本棚:規程 切れた 本文 [本棚:a.pdf p.1] 続き"
    ChkLong43 "R40F2_未閉じは後続へ食いつかない", modLive.CiteTagSpans(s, st, ln), 1
    ChkStr43 "R40F2_未閉じ後の正常タグ", Mid$(s, st(1), ln(1)), "[本棚:a.pdf p.1]"
    s = "[パック(山田):切れた [本棚:a.pdf p.1]"
    ChkLong43 "R40F2_未閉じパック→本棚", modLive.CiteTagSpans(s, st, ln), 1
    ChkStr43 "R40F2_未閉じパック後の正常タグ", Mid$(s, st(1), ln(1)), "[本棚:a.pdf p.1]"

    ' PageLabel: Excel 系はシート、それ以外は p.、0は空。
    ChkStr43 "R40F3_pdf", modLive.PageLabel("約款.pdf", 12), " p.12"
    ChkStr43 "R40F3_xlsx", modLive.PageLabel("売上.xlsx", 1), " シート1"
    ChkStr43 "R40F3_XLSM大文字", modLive.PageLabel("一覧.XLSM", 3), " シート3"
    ChkStr43 "R40F3_page0", modLive.PageLabel("売上.xlsx", 0), ""
    ChkStr43 "R40F3_拡張子なし", modLive.PageLabel("メモ", 2), " p.2"

    ' CellAddressOf / RowPrefix
    ChkStr43 "R40F3_A1", modExtractorExcel.CellAddressOf(1, 1), "A1"
    ChkStr43 "R40F3_Z10", modExtractorExcel.CellAddressOf(26, 10), "Z10"
    ChkStr43 "R40F3_AA6", modExtractorExcel.CellAddressOf(27, 6), "AA6"
    ChkStr43 "R40F3_IV", modExtractorExcel.CellAddressOf(256, 1), "IV1"
    ChkStr43 "R40F3_範囲外", modExtractorExcel.CellAddressOf(0, 5), ""
    ChkStr43 "R40F3_prefix", modExtractorExcel.RowPrefix(3, 6), "[C6] "
    ChkStr43 "R40F3_prefix範囲外は印無し", modExtractorExcel.RowPrefix(0, 5), ""
End Sub

' ---- G: R41 §1 A PageTagPart(単一情報源)の直接検査 --------------------------
Private Sub TestR41PageTagPart43()
    ChkStr43 "R41_xls小文字", modMode.PageTagPart("a.xls", 1), " シート1"
    ChkStr43 "R41_XLSB大文字", modMode.PageTagPart("B.XLSB", 4), " シート4"
    ChkStr43 "R41_pdfはp", modMode.PageTagPart("c.pdf", 2), " p.2"
    ' page=0でも空にしない(PageLabelとの違い。CiteTagFrom("議事録.txt",0,"")の
    ' 既存挙動=modTestsPure38を保つため)。
    ChkStr43 "R41_page0でも空にしない", modMode.PageTagPart("d.xlsx", 0), " シート0"
End Sub

' ---- H: R41 §3 C2 CiteTagSpans の閉じ位置の伸長(]入りファイル名) ------------
Private Sub TestR41CiteTagSpansExtend43()
    Dim st() As Long, ln() As Long
    Dim s As String

    ' (a) 資料名に "]" を含む(report[1].pdf)。最初の "]" では閉じない
    '     (直前が数字でも、その前が p./シートでない)ので伸ばして最後まで拾う。
    s = "[本棚:report[1].pdf p.3] 本文"
    ChkLong43 "R41C2_伸長1件", modLive.CiteTagSpans(s, st, ln), 1
    ChkStr43 "R41C2_伸長後の切り出し", Mid$(s, st(1), ln(1)), "[本棚:report[1].pdf p.3]"

    ' (b) シート表記でも同様に伸びる(末尾まで=trailing無し)。
    s = "[本棚:計算[改定].xlsx シート2]"
    ChkLong43 "R41C2_シート表記伸長1件", modLive.CiteTagSpans(s, st, ln), 1
    ChkStr43 "R41C2_シート表記伸長後の切り出し", Mid$(s, st(1), ln(1)), s

    ' (c) 伸ばしても該当が無ければ従来どおり最初の "]"(空振りで壊さない)。
    s = "[本棚:x[1] 本文] 続き"
    ChkLong43 "R41C2_該当無しは従来どおり1件", modLive.CiteTagSpans(s, st, ln), 1
    ChkStr43 "R41C2_該当無しは最初の]で切る", Mid$(s, st(1), ln(1)), "[本棚:x[1]"

    ' (d) 2件連続でも、それぞれ正しく切り出せる。
    s = "[本棚:a[1].pdf p.1][本棚:b.pdf p.2]"
    ChkLong43 "R41C2_2件連続", modLive.CiteTagSpans(s, st, ln), 2
    ChkStr43 "R41C2_2件連続の1件目", Mid$(s, st(1), ln(1)), "[本棚:a[1].pdf p.1]"
    ChkStr43 "R41C2_2件連続の2件目", Mid$(s, st(2), ln(2)), "[本棚:b.pdf p.2]"
End Sub

Public Sub RunAll43()
    On Error GoTo H00Fail43
    TestR40Display43
H01Next43:
    On Error GoTo H01Fail43
    TestIsOn43
H02Next43:
    On Error GoTo H02Fail43
    TestDedupeKeys43
H03Next43:
    On Error GoTo H03Fail43
    TestPerChapterCap43
H04Next43:
    On Error GoTo H04Fail43
    TestCountLines43
H05Next43:
    On Error GoTo H05Fail43
    TestBuildOnePassPrompt43
H06Next43:
    On Error GoTo H06Fail43
    TestR41PageTagPart43
H07Next43:
    On Error GoTo H07Fail43
    TestR41CiteTagSpansExtend43
H01Done43:
    On Error GoTo 0
    Exit Sub

H00Fail43:
    modTestRunner.Check "TestR40Display43(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Next43
H01Fail43:
    modTestRunner.Check "TestIsOn43(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H02Next43
H02Fail43:
    modTestRunner.Check "TestDedupeKeys43(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H03Next43
H03Fail43:
    modTestRunner.Check "TestPerChapterCap43(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H04Next43
H04Fail43:
    modTestRunner.Check "TestCountLines43(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H05Next43
H05Fail43:
    modTestRunner.Check "TestBuildOnePassPrompt43(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H06Next43
H06Fail43:
    modTestRunner.Check "TestR41PageTagPart43(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H07Next43
H07Fail43:
    modTestRunner.Check "TestR41CiteTagSpansExtend43(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done43
End Sub
