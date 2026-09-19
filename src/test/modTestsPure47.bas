Attribute VB_Name = "modTestsPure47"
Option Explicit

' ============================================================================
' modTestsPure47 - R43 波B(UI/UXポリッシュ)の純ロジック回帰。
'   modTestRunner.RunAllPureTests から直接呼ばれる独立の別枝
'   (modTestsPure31〜45と同型)。
' ----------------------------------------------------------------------------
' 【何を固定するか(手計算の根拠。仕様 spec_20260910_R43 §2-1/2-10 対応)】
'
'   A modKnowledgeBar.OnAccentColor(accent面に乗せる文字色の輝度判定):
'     WCAG相対輝度 L = 0.2126*R'+0.7152*G'+0.0722*B'(sRGB線形化後)。
'     判定は白と濃色の実コントラストを比べて大きい方を採る
'     (司令塔 Fix。旧しきい値0.179方式は light primary を 3.93=AA未達にした)。
'     A1 dark/sakura/light共通accent RGB(0,168,89): L=0.2873(手計算)
'       →濃色RGB(17,24,39)との対比5.70:1・白との対比3.11:1 → 濃色を選ぶ。
'     A2 msad accent RGB(7,169,99): L=0.2932(手計算)
'       →濃色対比5.80:1・白対比3.06:1(既存コメントの実測値と一致) → 濃色。
'     A3 ocean accent RGB(0,145,200): L=0.2442(手計算)
'       →濃色対比4.97:1・白対比3.57:1 → 濃色(4.5:1をわずかに上回る境界例)。
'     A4 gold accent RGB(230,196,80)(「エグゼクティブ・ゴールド」・実機報告の
'       1.70:1=ほぼ読めない元凶): L=0.5688(手計算)
'       →白対比1.70:1(実機報告の数値と一致)・濃色対比10.46:1 → 濃色。
'     A5 中間グレーRGB(105,105,105): L=0.1413(手計算・0.179未満)
'       →白対比5.49:1・濃色対比3.23:1 → 白を選ぶ(白branchの検証)。
'     A6 純黒RGB(0,0,0): L=0 → 白を選ぶ(自明な境界)。
'     A7 純白RGB(255,255,255): L=1 → 濃色を選ぶ(自明な境界)。
'
'   B modHubBadge.BadgeSpans(獲得済みバッジの文字区間):
'     呼び出し元(modHub.DrawBadges)の連結規約 sb = sb & mark(2文字) & " "
'     (1文字) & title & "   "(区切り3文字) を1文字も変えず踏襲する。
'     区間は「mark+空白+資料名」まで(末尾の区切り3文字は含めない)。
'     B1 titles=["初回質問","10問達成","伝道師"] earned=(True,False,True):
'       badge0: segLen=3+Len("初回質問")=3+4=7 → 獲得→span(1,7)。
'         pos = 1+7+3 = 11。
'       badge1: segLen=3+Len("10問達成")=3+5=8 → 未獲得→span無し。
'         pos = 11+8+3 = 22。
'       badge2: segLen=3+Len("伝道師")=3+3=6 → 獲得→span(22,6)。
'       → cnt=2, starts=(1,22), lens=(7,6)。
'     B2 全件獲得 titles=("AB","C") earned=(True,True):
'       badge0: segLen=3+2=5→span(1,5)。pos=1+5+3=9。
'       badge1: segLen=3+1=4→span(9,4)。
'       → cnt=2, starts=(1,9), lens=(5,4)。
'     B3 全件未獲得 titles=("X","YY") earned=(False,False) → cnt=0。
'     B4 n=0 → cnt=0(配列に触れず早期return)。
' ============================================================================

Private Sub ChkLong47(ByVal label As String, ByVal got As Long, ByVal want As Long)
    modTestRunner.Check "R43-47-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkStr47(ByVal label As String, ByVal got As String, ByVal want As String)
    modTestRunner.Check "R48-47-" & label, (got = want), "実際=[" & got & "] 期待=[" & want & "]"
End Sub

Private Sub ChkBool47(ByVal label As String, ByVal got As Boolean, ByVal want As Boolean)
    modTestRunner.Check "R48-47-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

' ----------------------------------------------------------------------------
' R群(R48) modRibbonFail - 社内AIリボンが返す「失敗の語彙」の判定
' ----------------------------------------------------------------------------
' なぜここを固定するか:
'   リボンちゃんは通信に失敗しても例外を出さず【文字列を戻り値で返す】。
'   R47 まで modGateway.CallLLM はこれを1つも検査しておらず、利用者には
'   「AIの回答」として確信度バッジと出典つきで表示されていた。
'   一方で、部分一致(InStr)で雑に拾うと【正当な回答をエラーに差し替える】
'   誤爆になる ―― このリポジトリは R21 D2 と R30 W2-2 で2回それをやっている。
'   よって「拾うべきものを拾う」と「拾ってはいけないものを拾わない」を
'   必ず対で固定する。R7〜R10 が後者で、こちらの方が実害が重い。
Private Sub TestRibbonFail47()
    ' --- 拾うべきもの(リボンの実ソース log.bas / GPT.bas から採取した実文字列)
    ChkStr47 "R1_http429_limit", _
        modRibbonFail.FailKind("(error:429)Requests to the ChatCompletions_Create Operation have exceeded call rate limit"), "limit"
    ChkStr47 "R2_http404", modRibbonFail.FailKind("(error:404)DeploymentNotFound"), "http"
    ChkStr47 "R3_conn", modRibbonFail.FailKind("接続切れ"), "conn"
    ChkStr47 "R4_filter", modRibbonFail.FailKind("content_filterに該当しました"), "filter"
    ChkStr47 "R5_parse_head", _
        modRibbonFail.FailKind("レスポンスから当該テキストを抽出できませんChatGPTの仕様が変更となった可能性がありますので、AIリボンをダウンロードしたホームページの情報をご確認ください"), "parse"
    ChkStr47 "R6_json_cut", modRibbonFail.FailKind("レスポンスのJSON文字列が途中で終了しています。"), "parse"
    ' 前後の空白は Trim$ で落として完全一致させる
    ChkStr47 "R3b_conn_padded", modRibbonFail.FailKind("  接続切れ  "), "conn"

    ' --- 拾ってはいけないもの(誤爆の反例。ここが本丸)
    ' 「接続切れ」を部分一致で拾うと、通信障害の手順書を引いた回の正当な回答が
    ' まるごとエラーに化ける。完全一致でしか拾わないことを固定する。
    ChkStr47 "R7_conn_substring_is_answer", _
        modRibbonFail.FailKind("接続切れの場合は、担当部署へご連絡ください。"), ""
    ChkStr47 "R7b_conn_substring_with_tag", _
        modRibbonFail.FailKind("接続切れの場合は、担当部署へご連絡ください。[本棚: 手順書.pdf p.3]"), ""
    ' R21 D2 で実機が踏んだ「正当な短文回答」。移設後も守られることを固定する。
    ChkStr47 "R8_limit_false_positive", _
        modRibbonFail.FailKind("請求回数の上限はありません。[本棚: 約款.pdf p.12]"), ""
    ChkStr47 "R9_normal_answer", _
        modRibbonFail.FailKind("保険料は年額3,000円です。[本棚: 約款.pdf p.5]"), ""
    ChkStr47 "R10_empty", modRibbonFail.FailKind(""), ""
    ' content_filter を語っただけの回答(完全一致でないので失敗ではない)
    ChkStr47 "R10b_filter_substring", _
        modRibbonFail.FailKind("content_filterに該当しましたという表示が出た場合の対処は次のとおりです。"), ""

    ' --- HTTPステータスの取り出し(IsNumeric は "1e5"/"-1" を通すので使わない)
    ChkLong47 "R11_status429", modRibbonFail.HttpStatusOf("(error:429)x"), 429
    ChkLong47 "R12_status_empty", modRibbonFail.HttpStatusOf("(error:)x"), 0
    ChkLong47 "R13_status_alpha", modRibbonFail.HttpStatusOf("(error:abc)x"), 0
    ChkLong47 "R14_status_not_err", modRibbonFail.HttpStatusOf("ふつうの回答です"), 0
    ChkLong47 "R14b_status_no_paren", modRibbonFail.HttpStatusOf("(error:429"), 0

    ' --- 種別 → エラーコード
    ChkStr47 "R15a_code_limit", modRibbonFail.CodeFor("limit"), "E0204"
    ChkStr47 "R15b_code_http", modRibbonFail.CodeFor("http"), "E0205"
    ChkStr47 "R15c_code_conn", modRibbonFail.CodeFor("conn"), "E0205"
    ChkStr47 "R15d_code_parse", modRibbonFail.CodeFor("parse"), "E0205"
    ChkStr47 "R15e_code_filter", modRibbonFail.CodeFor("filter"), "E0206"
    ChkStr47 "R15f_code_none", modRibbonFail.CodeFor(""), ""

    ' --- 取込の打ち切り判定(#ERR 化の前後どちらでも効くこと)
    ChkBool47 "R16_islimit_code", modRibbonFail.IsLimitErr("#ERR:E0204:(error:429)long english message that exceeds one hundred and twenty characters in total length for sure padding padding"), True
    ChkBool47 "R17_islimit_other_code", modRibbonFail.IsLimitErr("#ERR:E0205:接続切れ"), False
    ChkBool47 "R18_islimit_raw", modRibbonFail.IsLimitErr("申し訳ございません。本日の利用上限に達しました。"), True

    ' --- 移設でロジックが変わっていないことの確認(modGateway 時代の代表2件)
    ChkBool47 "R19a_moved_limit_true", _
        modRibbonFail.LooksLikeLimitError("申し訳ございません。本日の利用上限に達しました。"), True
    ChkBool47 "R19b_moved_limit_false", _
        modRibbonFail.LooksLikeLimitError("請求回数の上限はありません。[本棚: 約款.pdf p.12]"), False

    ' --- 入念が「しっかり」へ静かに降格する罠の固定(裁定書 R48 §3-4)
    ' modMode.UseVerify は thorough でも True を返す。入念が RunDeepFlow へ
    ' 落ちないのは modAsk.AnswerWithContext の分岐順序だけが理由。
    ' この戻り値が False に変わったら、分岐順序に依存した設計そのものが
    ' 変わったということなので、必ず modAsk 側も見直すこと。
    ChkBool47 "R20_useverify_thorough_is_true", modMode.UseVerify("thorough"), True
    ChkBool47 "R21_useverify_deep_is_true", modMode.UseVerify("deep"), True
    ChkBool47 "R22_useverify_quick_is_false", modMode.UseVerify("quick"), False

    ' --- 新エラーコードが modLog.FriendlyMessage の Case 追加漏れで
    '     「予期しない問題が発生しました」の既定値へ落ちていないこと。
    '     doc_gate はコードの存在は見るが、文言の中身までは見ない。
    '     R48 で E0205/E0206 を、R48 Fix で E0207 を足した。
    Dim m5 As String: m5 = modLog.FriendlyMessage("E0205")
    Dim m6 As String: m6 = modLog.FriendlyMessage("E0206")
    Dim m7 As String: m7 = modLog.FriendlyMessage("E0207")
    ChkBool47 "R23_e0205_not_fallback", (InStr(m5, "予期しない問題") = 0), True
    ChkBool47 "R24_e0206_not_fallback", (InStr(m6, "予期しない問題") = 0), True
    ChkBool47 "R25_e0207_not_fallback", (InStr(m7, "予期しない問題") = 0), True
    ' 3つが互いに違う文言であること(同じ文言なら分けた意味が無い)
    ChkBool47 "R26_e0205_ne_e0206", (m5 = m6), False
    ChkBool47 "R27_e0205_ne_e0207", (m5 = m7), False
    ChkBool47 "R28_e0206_ne_e0207", (m6 = m7), False
    ' 未定義コードは既定値へ落ちること(この検査自体が恒真でないことの担保)
    ChkBool47 "R29_unknown_is_fallback", _
        (InStr(modLog.FriendlyMessage("E9999"), "予期しない問題") > 0), True

    ' --- R48 Fix2: 中断の案内を1つに保つ(縫い目の固定)
    ' ESC は2つの経路で観測される ―― 段の境界(DoEvents → Err 18 → modAsk.bas:249)と、
    ' LLM 呼び出しの内側(modGateway が #ERR:E0207 へ変換)。利用者にとっては
    ' 【同じ操作】なので、違う案内を出してはならない。初版は「質問するボタンを
    ' 押してください」と「そのまま質問を送り直してください」で指示が食い違っていた。
    ' 逐語で固定する。**modAsk.bas:249 の文言を変えるときは、ここも必ず同時に変える。**
    ChkStr47 "R30_e0207_wording_matches_modask", modLog.FriendlyMessage("E0207"), _
        "操作を中断しました。もう一度質問するときは、質問するボタンを押してください。"

    ' --- R48 Fix3: 中断の伝播(IsCancelled)
    ' 多段の経路は #ERR を受けても次の段へ進むように作られているため、
    ' E0207 を作っただけでは「止めるときは ESC キー」の約束が果たせなかった。
    ' いまは5箇所(modAskMulti の論点ループ / Integrate の統合・自己点検 /
    ' modAskGlobal の回答段 / modAskThorough の要点整理・自己点検)がこれを見て
    ' 段を畳む。**通信失敗(E0205)を中断と取り違えると、直すべき障害が
    ' 「利用者が止めた」ことにされて消える**ので、その境目を対で固定する。
    ChkBool47 "R31_cancel_e0207", _
        modRibbonFail.IsCancelled("#ERR:E0207:利用者の操作で中断しました"), True
    ChkBool47 "R32_cancel_not_e0205", _
        modRibbonFail.IsCancelled("#ERR:E0205:接続切れ"), False
    ChkBool47 "R33_cancel_not_e0204", _
        modRibbonFail.IsCancelled("#ERR:E0204:(error:429)rate limit"), False
    ChkBool47 "R34_cancel_not_plain", _
        modRibbonFail.IsCancelled("保険料は年額3,000円です。[本棚: 約款.pdf p.5]"), False
    ChkBool47 "R35_cancel_not_empty", modRibbonFail.IsCancelled(""), False
End Sub

Private Sub TestOnAccentColor47()
    ChkLong47 "A1_dark_sakura_light", modKnowledgeBar.OnAccentColor(RGB(0, 168, 89)), RGB(17, 24, 39)
    ChkLong47 "A2_msad", modKnowledgeBar.OnAccentColor(RGB(7, 169, 99)), RGB(17, 24, 39)
    ChkLong47 "A3_ocean", modKnowledgeBar.OnAccentColor(RGB(0, 145, 200)), RGB(17, 24, 39)
    ChkLong47 "A4_gold", modKnowledgeBar.OnAccentColor(RGB(230, 196, 80)), RGB(17, 24, 39)
    ChkLong47 "A5_midgray_white", modKnowledgeBar.OnAccentColor(RGB(105, 105, 105)), RGB(255, 255, 255)
    ChkLong47 "A6_black_white", modKnowledgeBar.OnAccentColor(RGB(0, 0, 0)), RGB(255, 255, 255)
    ChkLong47 "A7_white_dark", modKnowledgeBar.OnAccentColor(RGB(255, 255, 255)), RGB(17, 24, 39)

    ' A8〜A11(司令塔 Fix): primary/danger の地。旧しきい値方式(0.179)では
    ' A8 が濃色に倒れて 3.93:1(AA未達)になっていた。白と濃色の実コントラストを
    ' 比べる方式なら、白 4.52:1 が選ばれる。数値はすべて python で再計算した実測。
    '   A8  light primary  RGB(0,137,62)   白 4.52 / 濃色 3.93 → 白
    '   A9  gold  primary  RGB(212,175,55) 白 2.10 / 濃色 8.44 → 濃色
    '   A10 gold  danger   RGB(248,113,113) 白 2.77 / 濃色 6.41 → 濃色
    '   A11 sakura danger  RGB(185,28,28)  白 6.47 / 濃色 2.74 → 白
    ChkLong47 "A8_light_primary_white", modKnowledgeBar.OnAccentColor(RGB(0, 137, 62)), RGB(255, 255, 255)
    ChkLong47 "A9_gold_primary_dark", modKnowledgeBar.OnAccentColor(RGB(212, 175, 55)), RGB(17, 24, 39)
    ChkLong47 "A10_gold_danger_dark", modKnowledgeBar.OnAccentColor(RGB(248, 113, 113)), RGB(17, 24, 39)
    ChkLong47 "A11_sakura_danger_white", modKnowledgeBar.OnAccentColor(RGB(185, 28, 28)), RGB(255, 255, 255)
End Sub

Private Sub TestBadgeSpans47()
    Dim starts() As Long, lens() As Long, cnt As Long

    Dim t1(0 To 2) As String, e1(0 To 2) As Boolean
    t1(0) = "初回質問": t1(1) = "10問達成": t1(2) = "伝道師"
    e1(0) = True: e1(1) = False: e1(2) = True
    cnt = modHubBadge.BadgeSpans(t1, e1, 3, starts, lens)
    ChkLong47 "B1_cnt", cnt, 2
    ChkLong47 "B1_start0", starts(0), 1
    ChkLong47 "B1_len0", lens(0), 7
    ChkLong47 "B1_start1", starts(1), 22
    ChkLong47 "B1_len1", lens(1), 6

    Dim t2(0 To 1) As String, e2(0 To 1) As Boolean
    t2(0) = "AB": t2(1) = "C"
    e2(0) = True: e2(1) = True
    cnt = modHubBadge.BadgeSpans(t2, e2, 2, starts, lens)
    ChkLong47 "B2_cnt", cnt, 2
    ChkLong47 "B2_start0", starts(0), 1
    ChkLong47 "B2_len0", lens(0), 5
    ChkLong47 "B2_start1", starts(1), 9
    ChkLong47 "B2_len1", lens(1), 4

    Dim t3(0 To 1) As String, e3(0 To 1) As Boolean
    t3(0) = "X": t3(1) = "YY"
    e3(0) = False: e3(1) = False
    cnt = modHubBadge.BadgeSpans(t3, e3, 2, starts, lens)
    ChkLong47 "B3_cnt", cnt, 0

    Dim t4() As String, e4() As Boolean
    cnt = modHubBadge.BadgeSpans(t4, e4, 0, starts, lens)
    ChkLong47 "B4_cnt", cnt, 0
End Sub

' ----------------------------------------------------------------------------
' S群(R49 監査H-7) modInsightGate.PiiScanHit - 共有フォルダへ出す前の個人情報走査
' ----------------------------------------------------------------------------
' なぜここを固定するか:
'   解決済みQ&A(EmitVerifiedQA)は、質問全文・回答全文・実名の3点セットを
'   部内の共有フォルダへ置く経路。共有フォルダに一度書いたものは取り消せない。
'   R49 までこの経路だけ走査が1回も無く、困りごと(EmitGap)・訂正
'   (EmitCorrection)の2経路だけが PiiBlocked を通っていた。
'
'   固定するのは判定部品 PiiScanHit(通知もログも出さない純判定)。
'   陽性(=送らない)だけでなく【陰性の反例】を必ず置く ―― この関門は
'   過検知すると部内ナレッジのフライホイールを止めるので、
'   「日付が入っているだけの質問」「社内資料由来の短い数字」が通ることまで
'   固定しないと、次の誰かが閾値を触ったときに静かに壊れる。
Private Sub TestPiiScan47()
    ' --- 陽性: 送ってはいけないもの -------------------------------------
    ' 携帯番号(modPii が継続文字として数える "-" 込みで11桁)
    ChkBool47 "S1_携帯番号", _
        modInsightGate.PiiScanHit("取引先の田中様の携帯 090-1234-5678 へ連絡"), True
    ' 区切り無しの長い数字列
    ChkBool47 "S2_長い数字列", _
        modInsightGate.PiiScanHit("口座 1234567890123 の扱いは?"), True

    ' --- 陰性: 止めてはいけないもの(誤検知の反例) -----------------------
    ' 日付+時刻。StripDateLike が無いと 4+2+2+2+2=12桁として一発で当たり、
    ' 「日時をひとつ書いただけの質問」が永久に部内へ出なくなる(R32 F4 の事故)。
    ChkBool47 "S3_日時は素通り", _
        modInsightGate.PiiScanHit("2026-08-14 10:00 時点の取扱いを教えて"), False
    ' 全角の日付(R33 W2-7。NormalizeWidth を先頭で通していないと陽性になる)
    ChkBool47 "S4_全角日時も素通り", _
        modInsightGate.PiiScanHit(ChrW(&HFF12) & ChrW(&HFF10) & ChrW(&HFF12) & _
            ChrW(&HFF16) & ChrW(&HFF0D) & ChrW(&HFF10) & ChrW(&HFF18) & _
            ChrW(&HFF0D) & ChrW(&HFF11) & ChrW(&HFF14) & " の改定について"), False
    ' 業務で普通に出る短い数字(約款の条番号・金額)は通す
    ChkBool47 "S5_短い数字は素通り", _
        modInsightGate.PiiScanHit("第12条の免責金額 50000 円の根拠は?"), False
    ' 数字を含まない普通の質問
    ChkBool47 "S6_平文は素通り", _
        modInsightGate.PiiScanHit("団体扱いの解約返戻金の計算方法を教えてください"), False
    ' 空・空白だけは走査対象外(資料名が空のときに伏せ字化しないため)
    ChkBool47 "S7_空は素通り", modInsightGate.PiiScanHit(""), False
    ChkBool47 "S8_空白のみは素通り", modInsightGate.PiiScanHit("   "), False

    ' --- 改行の扱い(R32 F17 と同じ真因の再発防止) -----------------------
    ' 走査は ScanClean を通すので、改行で分かれた別々の数字列は繋がらない。
    ' ここを Clean1(改行→半角スペース)で走査すると "12345" と "67890" が
    ' 1本の10桁として誤検知する。
    ChkBool47 "S9_改行で数字は繋がらない", _
        modInsightGate.PiiScanHit("整理番号 12345" & vbLf & "67890 の件"), False
End Sub

Public Sub RunAll47()
    On Error GoTo H01Fail47
    TestOnAccentColor47
H02Next47:
    On Error GoTo H02Fail47
    TestBadgeSpans47
H03Next47:
    On Error GoTo H03Fail47
    TestRibbonFail47
H04Next47:
    On Error GoTo H04Fail47
    TestPiiScan47
H01Done47:
    On Error GoTo 0
    Exit Sub

H04Fail47:
    modTestRunner.Check "TestPiiScan47(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done47

H01Fail47:
    modTestRunner.Check "TestOnAccentColor47(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H02Next47
H02Fail47:
    modTestRunner.Check "TestBadgeSpans47(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H03Next47
H03Fail47:
    modTestRunner.Check "TestRibbonFail47(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H04Next47
End Sub
