Attribute VB_Name = "modTestsPure46"
Option Explicit

' ============================================================================
' modTestsPure46 - R43 波A(UI/UXポリッシュ)の純ロジック回帰。
'   modTestRunner.RunAllPureTests から直接呼ばれる独立の別枝
'   (modTestsPure31〜45と同型)。
' ----------------------------------------------------------------------------
' 【何を固定するか(手計算の根拠。仕様 spec_20260910_R43 §1 対応)】
'
'   A modLiveStyle.LeadParaIndex(text): 先頭段落を「結論」として強調して
'     よいか(1=対象/0=対象外)。
'     A1 通常の結論文(vbCr区切り2段落・1段落目が結論): 1段落目が「■」でも
'       💭でもないので対象=1。
'     A2 ■見出しで始まる回答: LTrim後の先頭1文字が"■"なので対象外=0。
'     A3 聞き返し(💭=ChrW(&HD83D)&ChrW(&HDCAD)で始まる): LTrim後の先頭2文字
'       (サロゲートペア)が💭と一致するので対象外=0。
'     A4 空文字: LenB=0で即0。
'     A5 先頭に半角空白が付く結論文: LTrimで空白を落としてから判定するので
'       対象=1(空白自体は■でも💭でもない)。
'     A6 改行が無い単一段落: brk=0となりfirstPara=text全体。結論扱いなので
'       対象=1。
'     A7 "■"1文字だけの段落: Left$(firstPara,1)="■"なので対象外=0。
'     A8 vbLf区切り(Shape読み戻し前の直接呼び出しを想定): brkはvbCrが
'       見つからずvbLfにフォールバックするので、vbCr区切りと同じ判定になる
'       (このケースは1段落目が結論文なので対象=1)。
'
'   B modLiveStyle.CellRefSpans(s): "(<列A-Z1文字以上><行0-9数字1文字以上>
'     付近)" の位置・長さを列挙する。位置・長さはpython3でCellRefCloseAtと
'     同じ規則を再実装して実測(結果はこのファイルの各アサートに転記)。
'     B1 "根拠は(A6 付近)にあります。": "根拠は"が3字なので"("は位置4。
'       "(A6 付近)"は "(" A 6 空白 付 近 ")" の7字→開始4・長さ7。
'     B2 2文字列(AA12): "(AA12 付近)"は全体が対象で開始1・長さ9
'       ("(" A A 1 2 空白 付 近 ")" =9字)。
'     B3 " 付近"が無い"(A6)": 数字の直後が")"で" 付近"と一致しないので0件。
'     B4 空白を欠いた"(A6付近)": 数字の直後3字が"付近)"で" 付近"と不一致
'       (先頭が空白でない)ため0件。
'     B5 2件連続(全角読点で区切る): "(A6 付近)、また(B12 付近)にも記載が
'       あります。" → 1件目は開始1・長さ7("(A6 付近)")。続く"、また"が3字
'       (位置8-10)なので2件目の"("は位置11、"(B12 付近)"は
'       "(" B 1 2 空白 付 近 ")" =8字→開始11・長さ8。
'     B6 空文字: 0件。
'     B7 列の文字が無い"(6 付近)": 数字の前に大文字が1つも無いので0件。
'     B8 閉じ括弧が無い"(A6 付近": ")"が続かないので0件。
'
'   C modPeek.ChipDocLabel(s,n): n字以下はそのまま、超えたときだけ
'     SafeLeft(s,n)&"…"。境界(ちょうどn字)は切らない。
'   D modPeek.ChipOverflowLabel(totalUnique,shownCount): 超過分が無ければ
'     空文字、超過があれば"ほか N件"(N=totalUnique-shownCount)。
'   E modPeek.PeekSpans(s): 見出し(1行目)/breadcrumb(見出し直後が「【」の
'     行)/行頭"[...]"タグ(複数可)/閉じる案内(最後の改行より後ろ全部)の
'     位置・長さ。位置はpython3でPeekSpansと同じ規則を再実装して実測。
'     E1(breadcrumb+タグ2件・vbLf区切り):
'       "見出し行\n【crumb】\nDIVIDER\n[A6] 本文1\n\n[B7] 本文2\nDIVIDER\n
'       閉じる案内" (33文字境界は下記個々のアサート参照)。
'       見出し=開始1・長さ4("見出し行")。breadcrumb=開始6・長さ7
'       ("【crumb】"=【crumbの6字+】=7字)。1件目タグ=開始22・長さ4
'       ("[A6]")。2件目タグ=開始32・長さ4("[B7]")。閉じる案内=開始49・
'       長さ5("閉じる案内")。件数=2。
'     E2(breadcrumb無し・タグ無し): "見出し行2\nDIVIDER\n本文のみ\nDIVIDER\n
'       閉じる案内2" → 見出し=開始1・長さ5("見出し行2")。breadcrumb長さ0。
'       件数=0。閉じる案内=開始28・長さ6("閉じる案内2")。
'     E3(空文字): 全て0。
'     E4(行頭以外の"["は拾わない): "見出し\nabc[C9] middle\n閉じ" →
'       件数=0(「[」の直前が改行でも文字列先頭でもないため)。見出し=
'       開始1・長さ3。閉じる案内=開始20・長さ2("閉じ")。
'
' 【どう壊すと落ちるか (discriminate)】
'   ・LeadParaIndexで■判定をLTrim前の文字列に対して行うと、先頭に空白が
'     付くだけの通常の結論文(A5)まで誤って対象外になる。
'   ・LeadParaIndexで💭判定を1文字だけ見ると、サロゲート上位だけ一致する
'     別の絵文字と誤認する(A3はサロゲートペア2文字一致を要求している)。
'   ・CellRefSpansで" 付近"の前の空白を省くと、B4(空白無し)も誤って
'     1件と数えてしまう。
'   ・ChipDocLabelでLen(s)>=nにすると、ちょうどn字の資料名まで…が付く
'     (境界のオフバイワン)。
'   ・PeekSpansで行頭判定を「直前が改行」だけにして先頭位置を見落とすと、
'     1行目("["で始まる異常な見出し)がタグとして誤検出される。E4は
'     「["が行の途中にある場合は拾わない」ことを固定している。
' ============================================================================

Private Sub ChkStr46(ByVal label As String, ByVal got As String, ByVal want As String)
    modTestRunner.Check "R43-46-" & label, (StrComp(got, want, vbBinaryCompare) = 0), _
        "実際=[" & got & "] 期待=[" & want & "]"
End Sub

Private Sub ChkLong46(ByVal label As String, ByVal got As Long, ByVal want As Long)
    modTestRunner.Check "R43-46-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

' ---- A: modLiveStyle.LeadParaIndex -------------------------------------------
Private Sub TestLeadParaIndex46()
    ' A1: 通常の結論文(2段落・vbCr区切り)
    ChkLong46 "LeadParaIndex_通常結論文", _
        modLiveStyle.LeadParaIndex("結論です。" & vbCr & "■ 詳細"), 1

    ' A2: ■見出しで始まる回答は対象外
    ChkLong46 "LeadParaIndex_見出しで始まるは対象外", _
        modLiveStyle.LeadParaIndex("■ 見出し" & vbCr & "本文"), 0

    ' A3: 聞き返し(💭で始まる)は対象外
    Dim pin As String: pin = ChrW(&HD83D) & ChrW(&HDCAD)
    ChkLong46 "LeadParaIndex_聞き返しは対象外", _
        modLiveStyle.LeadParaIndex(pin & " ご質問は複数の資料に該当します。" & vbCr & _
            "■ どの資料のお話でしょうか"), 0

    ' A4: 空文字は対象外
    ChkLong46 "LeadParaIndex_空文字は0", modLiveStyle.LeadParaIndex(""), 0

    ' A5: 先頭の半角空白はLTrimしてから判定(結論文のまま=対象)
    ChkLong46 "LeadParaIndex_先頭空白はLTrim後判定", _
        modLiveStyle.LeadParaIndex("  結論です" & vbCr & "本文"), 1

    ' A6: 改行が無い単一段落はそのまま結論扱い
    ChkLong46 "LeadParaIndex_単一段落は対象", _
        modLiveStyle.LeadParaIndex("結論だけの一文です"), 1

    ' A7: "■"1文字だけの段落も対象外
    ChkLong46 "LeadParaIndex_見出し記号のみも対象外", modLiveStyle.LeadParaIndex("■"), 0

    ' A8: vbLf区切り(Shape読み戻し前を想定)でも同じ判定になる
    ChkLong46 "LeadParaIndex_vbLf区切りでも判定同じ", _
        modLiveStyle.LeadParaIndex("結論" & vbLf & "■ 見出し"), 1
End Sub

' ---- B: modLiveStyle.CellRefSpans ---------------------------------------------
Private Sub TestCellRefSpans46()
    Dim st() As Long, ln() As Long
    Dim s As String

    ' B1: 1文字列の単純なセル番地
    s = "根拠は(A6 付近)にあります。"
    ChkLong46 "CellRefSpans_単純1件", modLiveStyle.CellRefSpans(s, st, ln), 1
    ChkLong46 "CellRefSpans_単純1件_開始", st(1), 4
    ChkStr46 "CellRefSpans_単純1件_切り出し", Mid$(s, st(1), ln(1)), "(A6 付近)"

    ' B2: 2文字の列(AA)
    s = "(AA12 付近)"
    ChkLong46 "CellRefSpans_2文字列1件", modLiveStyle.CellRefSpans(s, st, ln), 1
    ChkStr46 "CellRefSpans_2文字列1件_切り出し", Mid$(s, st(1), ln(1)), "(AA12 付近)"

    ' B3: " 付近"が無ければ0件
    ChkLong46 "CellRefSpans_付近無しは0件", modLiveStyle.CellRefSpans("(A6)", st, ln), 0

    ' B4: 数字の直後に空白が無ければ0件(" 付近"の空白必須)
    ChkLong46 "CellRefSpans_空白無しは0件", modLiveStyle.CellRefSpans("(A6付近)", st, ln), 0

    ' B5: 2件連続(全角読点区切り)
    s = "(A6 付近)、また(B12 付近)にも記載があります。"
    ChkLong46 "CellRefSpans_2件連続", modLiveStyle.CellRefSpans(s, st, ln), 2
    ChkStr46 "CellRefSpans_2件連続_1件目", Mid$(s, st(1), ln(1)), "(A6 付近)"
    ChkLong46 "CellRefSpans_2件連続_2件目開始", st(2), 11
    ChkStr46 "CellRefSpans_2件連続_2件目", Mid$(s, st(2), ln(2)), "(B12 付近)"

    ' B6: 空文字は0件
    ChkLong46 "CellRefSpans_空は0件", modLiveStyle.CellRefSpans("", st, ln), 0

    ' B7: 列の文字が無ければ0件
    ChkLong46 "CellRefSpans_列無しは0件", modLiveStyle.CellRefSpans("(6 付近)", st, ln), 0

    ' B8: 閉じ括弧が無ければ0件
    ChkLong46 "CellRefSpans_閉じ無しは0件", modLiveStyle.CellRefSpans("(A6 付近", st, ln), 0
End Sub

' ---- C: modPeek.ChipDocLabel --------------------------------------------------
Private Sub TestChipDocLabel46()
    ' (a) n字ちょうどは切らない(境界値)
    ChkStr46 "ChipDocLabel_境界はそのまま", modPeek.ChipDocLabel(String(16, "あ"), 16), _
        String(16, "あ")

    ' (b) n字未満はそのまま
    ChkStr46 "ChipDocLabel_短い名前はそのまま", modPeek.ChipDocLabel("短い名前.pdf", 16), _
        "短い名前.pdf"

    ' (c) n字を超えたら切って…を付ける
    ChkStr46 "ChipDocLabel_超過は省略記号付き", modPeek.ChipDocLabel(String(20, "あ"), 16), _
        String(16, "あ") & ChrW(&H2026)
End Sub

' ---- D: modPeek.ChipOverflowLabel ----------------------------------------------
Private Sub TestChipOverflowLabel46()
    ' (a) 超過が無ければ空文字
    ChkStr46 "ChipOverflowLabel_超過無しは空", modPeek.ChipOverflowLabel(4, 4), ""
    ChkStr46 "ChipOverflowLabel_総数が少なくても空", modPeek.ChipOverflowLabel(3, 4), ""

    ' (b) 超過があれば「ほかN件」
    ChkStr46 "ChipOverflowLabel_3件超過", modPeek.ChipOverflowLabel(7, 4), "ほか 3件"
    ChkStr46 "ChipOverflowLabel_1件超過", modPeek.ChipOverflowLabel(5, 4), "ほか 1件"
End Sub

' ---- E: modPeek.PeekSpans ------------------------------------------------------
Private Sub TestPeekSpans46()
    Dim hS As Long, hL As Long, cS As Long, cL As Long, clS As Long, clL As Long
    Dim rSt() As Long, rLn() As Long
    Dim s As String
    Dim n As Long

    ' E1: breadcrumb有り+行頭タグ2件(vbLf区切り)
    s = "見出し行" & vbLf & "【crumb】" & vbLf & "DIVIDER" & vbLf & _
        "[A6] 本文1" & vbLf & vbLf & "[B7] 本文2" & vbLf & "DIVIDER" & vbLf & "閉じる案内"
    n = modPeek.PeekSpans(s, hS, hL, cS, cL, rSt, rLn, clS, clL)
    ChkLong46 "PeekSpans_E1_件数", n, 2
    ChkLong46 "PeekSpans_E1_見出し開始", hS, 1
    ChkLong46 "PeekSpans_E1_見出し長さ", hL, 4
    ChkStr46 "PeekSpans_E1_見出し切り出し", Mid$(s, hS, hL), "見出し行"
    ChkLong46 "PeekSpans_E1_crumb開始", cS, 6
    ChkLong46 "PeekSpans_E1_crumb長さ", cL, 7
    ChkStr46 "PeekSpans_E1_crumb切り出し", Mid$(s, cS, cL), "【crumb】"
    ChkLong46 "PeekSpans_E1_タグ1開始", rSt(1), 22
    ChkStr46 "PeekSpans_E1_タグ1切り出し", Mid$(s, rSt(1), rLn(1)), "[A6]"
    ChkLong46 "PeekSpans_E1_タグ2開始", rSt(2), 32
    ChkStr46 "PeekSpans_E1_タグ2切り出し", Mid$(s, rSt(2), rLn(2)), "[B7]"
    ChkLong46 "PeekSpans_E1_閉じる開始", clS, 49
    ChkStr46 "PeekSpans_E1_閉じる切り出し", Mid$(s, clS, clL), "閉じる案内"

    ' E2: breadcrumb無し・タグ無し
    s = "見出し行2" & vbLf & "DIVIDER" & vbLf & "本文のみ" & vbLf & "DIVIDER" & vbLf & "閉じる案内2"
    n = modPeek.PeekSpans(s, hS, hL, cS, cL, rSt, rLn, clS, clL)
    ChkLong46 "PeekSpans_E2_件数", n, 0
    ChkStr46 "PeekSpans_E2_見出し切り出し", Mid$(s, hS, hL), "見出し行2"
    ChkLong46 "PeekSpans_E2_crumb長さゼロ", cL, 0
    ChkStr46 "PeekSpans_E2_閉じる切り出し", Mid$(s, clS, clL), "閉じる案内2"

    ' E3: 空文字は全て0
    n = modPeek.PeekSpans("", hS, hL, cS, cL, rSt, rLn, clS, clL)
    ChkLong46 "PeekSpans_E3_件数ゼロ", n, 0
    ChkLong46 "PeekSpans_E3_見出し長さゼロ", hL, 0
    ChkLong46 "PeekSpans_E3_閉じる長さゼロ", clL, 0

    ' E4: 行の途中の"["はタグとして拾わない
    s = "見出し" & vbLf & "abc[C9] middle" & vbLf & "閉じ"
    n = modPeek.PeekSpans(s, hS, hL, cS, cL, rSt, rLn, clS, clL)
    ChkLong46 "PeekSpans_E4_行途中の角括弧は除外", n, 0
    ChkStr46 "PeekSpans_E4_見出し切り出し", Mid$(s, hS, hL), "見出し"
    ChkStr46 "PeekSpans_E4_閉じる切り出し", Mid$(s, clS, clL), "閉じ"
End Sub

' ----------------------------------------------------------------------------
' RunAll46 - modTestRunner から呼ばれる総合エントリーポイント
' ----------------------------------------------------------------------------
Public Sub RunAll46()
    On Error GoTo H01Fail46
    TestLeadParaIndex46
H02Next46:
    On Error GoTo H02Fail46
    TestCellRefSpans46
H03Next46:
    On Error GoTo H03Fail46
    TestChipDocLabel46
H04Next46:
    On Error GoTo H04Fail46
    TestChipOverflowLabel46
H05Next46:
    On Error GoTo H05Fail46
    TestPeekSpans46
H01Done46:
    On Error GoTo 0
    Exit Sub

H01Fail46:
    modTestRunner.Check "TestLeadParaIndex46(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H02Next46
H02Fail46:
    modTestRunner.Check "TestCellRefSpans46(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H03Next46
H03Fail46:
    modTestRunner.Check "TestChipDocLabel46(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H04Next46
H04Fail46:
    modTestRunner.Check "TestChipOverflowLabel46(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H05Next46
H05Fail46:
    modTestRunner.Check "TestPeekSpans46(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done46
End Sub
