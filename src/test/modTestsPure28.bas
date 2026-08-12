Attribute VB_Name = "modTestsPure28"
Option Explicit

' ============================================================================
' modTestsPure28 - R28波2(実機第13報②「逆質問の断線」)+波3の純ロジック回帰。
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
'   (C) modChrome.ToastHeightFor — トースト高さ3段の文字数境界(R28 W3-2)。
'   (D) modVaultGallery.PageCapFor — 列数からページ枚数を出す算数(R28 W3-3)。
'   (E) modConvBridge.ComputeBridgeCore — 全往復運搬・総量クランプ・
'       二重挿入防止・クリア後空(R28 W3-6b)。
'   (F) modAsk.SetPrevMemory / ResetPrevMemory — 橋渡しと会話クリアが
'       modAsk の会話メモリへ直接届くこと(R28 W4-1/W4-2)。
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
    ' R28H F7(m-1): 資料は2件しか出していないのに「5」。資料番号としては
    ' 採れないが、他に数が無いので意図5として読む(旧仕様は 0,0 で全部捨てた)。
    CheckReply "資料番号が範囲外_nSrc2に5は意図5として拾う", "5", 2, 0, 5
    ' 「5-1」は資料5(範囲外)だけ捨て、意図1は生かす。2個目の数があるので
    ' F7の拾い直しは働かない(1個目を意図に読み替えない)。
    CheckReply "資料だけ範囲外_5-1", "5-1", 2, 0, 1
    ' 意図は①～⑤の5つだけ。9は採らない(資料3は生きる)。
    CheckReply "意図番号が範囲外_3-9", "3-9", 4, 3, 0
    ' ⑥は意図の範囲外。丸数字は資料番号にはならないので何も選ばれない
    ' (丸数字が在る=circled<>0 なので F7 の拾い直しも働かない)。
    CheckReply "丸数字が範囲外_⑥", ChrW(&H2465), 4, 0, 0
    ' 2桁は選択肢に存在しない(旧 PickSource は "12" の中の "1" を拾っていた)。
    ' ParseChoiceNumbers が空を返す=n1も0なので、F7でも拾わない。
    CheckReply "2桁_12は資料1と読まない", "12", 4, 0, 0
    ' R28H F7(m-2): 資料の選択肢を出していない逆質問(nSrc=0。意図だけを聞く型)。
    ' 資料番号は成立しないので、その数は意図番号として読む。
    CheckReply "nSrcが0_3は意図3として拾う", "3", 0, 0, 3
    CheckReply "nSrcが0_丸数字③も意図3", ChrW(&H2462), 0, 0, 3
    ' nSrc=0 でも「0」は「この中にない/わからない」。意図には読み替えない。
    CheckReply "nSrcが0_0は何も選ばれない", "0", 0, 0, 0
    ' 意図の上限(MAX_INTENT=5)を超える数は拾い直しの対象外。
    CheckReply "nSrcが0_6は範囲外で何も選ばれない", "6", 0, 0, 0
    ' 2つ打ってあれば従来どおり「1個目=資料/2個目=意図」。nSrc=0でも
    ' 1個目を意図に読み替えない(2個目が意図として既に読めている)。
    CheckReply "nSrcが0_2-3は意図3のみ(1個目は読み替えない)", "2-3", 0, 0, 3

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

' ----------------------------------------------------------------------------
' (C) modChrome.ToastHeightFor — R28H F6 で「文字数→3段テーブル」から
'     「推定幅→行数→高さ」の算数へ作り直した境界を固定する。
'     幅の見積り: 全角10.5pt / 半角5.25pt(TextSpan)。可視幅348pt で割って
'     行数を切り上げ、高さ = 切上げ(行数×15.2 + 10)、上限92ptでクランプ。
'     1行あたりの容量は 全角33.14字 / 半角66.28字。
' ----------------------------------------------------------------------------
Private Function Zen28(ByVal n As Long) As String
    Zen28 = String(n, ChrW(&H3042))   ' 「あ」= 全角(幅10.5pt)
End Function

Private Sub TestToastHeightFor28()
    ' 空文字は行数0だが、1行へ丸める(高さ0の帯は描けない)。
    modTestRunner.Check "R28H-F6_0字は1行26pt", (modChrome.ToastHeightFor("") = 26), ""

    ' --- 全角のみ。348/10.5 = 33.14字で改行。
    modTestRunner.Check "R28H-F6_全角33字は1行26pt(境界)", _
        (modChrome.ToastHeightFor(Zen28(33)) = 26), "h=" & modChrome.ToastHeightFor(Zen28(33))
    modTestRunner.Check "R28H-F6_全角34字は2行41pt(境界+1)", _
        (modChrome.ToastHeightFor(Zen28(34)) = 41), "h=" & modChrome.ToastHeightFor(Zen28(34))
    modTestRunner.Check "R28H-F6_全角66字は2行41pt(境界)", _
        (modChrome.ToastHeightFor(Zen28(66)) = 41), "h=" & modChrome.ToastHeightFor(Zen28(66))
    modTestRunner.Check "R28H-F6_全角67字は3行56pt(境界+1)", _
        (modChrome.ToastHeightFor(Zen28(67)) = 56), "h=" & modChrome.ToastHeightFor(Zen28(67))

    ' --- 半角のみ。全角の倍の66.28字で改行(旧テーブルはここを潰していた)。
    modTestRunner.Check "R28H-F6_半角66字は1行26pt(境界)", _
        (modChrome.ToastHeightFor(String(66, "a")) = 26), _
        "h=" & modChrome.ToastHeightFor(String(66, "a"))
    modTestRunner.Check "R28H-F6_半角67字は2行41pt(境界+1)", _
        (modChrome.ToastHeightFor(String(67, "a")) = 41), _
        "h=" & modChrome.ToastHeightFor(String(67, "a"))

    ' --- 混在。全角30字(315pt)+半角7字(36.75pt)=351.75pt > 348 で2行。
    modTestRunner.Check "R28H-F6_混在_全角30半角6は1行26pt", _
        (modChrome.ToastHeightFor(Zen28(30) & String(6, "a")) = 26), _
        "h=" & modChrome.ToastHeightFor(Zen28(30) & String(6, "a"))
    modTestRunner.Check "R28H-F6_混在_全角30半角7は2行41pt", _
        (modChrome.ToastHeightFor(Zen28(30) & String(7, "a")) = 41), _
        "h=" & modChrome.ToastHeightFor(Zen28(30) & String(7, "a"))

    ' --- 実機報告の文言(入念モードの切替トースト・114字相当)。
    '     推定幅1,107.75pt → 4行 → 4×15.2+10 = 70.8 → 71pt。旧実装は62ptで切れた。
    modTestRunner.Check "R28H-F6_全角105字相当は4行71pt(入念モードの実文言)", _
        (modChrome.ToastHeightFor(Zen28(105)) = 71), "h=" & modChrome.ToastHeightFor(Zen28(105))

    ' --- 上限クランプ。5行=86ptまでは伸び、6行ぶん以上は92ptで止まる。
    modTestRunner.Check "R28H-F6_全角165字は5行86pt", _
        (modChrome.ToastHeightFor(Zen28(165)) = 86), "h=" & modChrome.ToastHeightFor(Zen28(165))
    modTestRunner.Check "R28H-F6_全角166字は6行ぶんだが92ptでクランプ", _
        (modChrome.ToastHeightFor(Zen28(166)) = 92), "h=" & modChrome.ToastHeightFor(Zen28(166))
    modTestRunner.Check "R28H-F6_全角1000字でも92ptを超えない", _
        (modChrome.ToastHeightFor(Zen28(1000)) = 92), "h=" & modChrome.ToastHeightFor(Zen28(1000))
End Sub

' ----------------------------------------------------------------------------
' (D) modVaultGallery.PageCapFor — cols*2。cols<1は1へクランプしてから*2。
' ----------------------------------------------------------------------------
Private Sub TestPageCapFor28()
    modTestRunner.Check "R28-W3-3_cols3は6枚", (modVaultGallery.PageCapFor(3) = 6), ""
    modTestRunner.Check "R28-W3-3_cols4は8枚", (modVaultGallery.PageCapFor(4) = 8), ""
    modTestRunner.Check "R28-W3-3_cols6は12枚", (modVaultGallery.PageCapFor(6) = 12), ""
    modTestRunner.Check "R28-W3-3_cols0は1扱いで2枚", (modVaultGallery.PageCapFor(0) = 2), ""
    modTestRunner.Check "R28-W3-3_cols負値も1扱いで2枚", (modVaultGallery.PageCapFor(-1) = 2), ""
End Sub

' ----------------------------------------------------------------------------
' (E) modConvBridge.ComputeBridgeCore — 全往復運搬・総量クランプ・
'     二重挿入防止・クリア後空(R28 W3-6b)。
' ----------------------------------------------------------------------------
Private Function PairsCount28(ByVal joined As String) As Long
    If LenB(joined) = 0 Then
        PairsCount28 = 0
    Else
        Dim p() As String: p = Split(joined, ";;;")
        PairsCount28 = UBound(p) - LBound(p) + 1
    End If
End Function

' (E-1) 5往復すべてが運搬される(FirstPair1件のみだった旧実装からの拡張)。
Private Sub TestConvBridgeCarryFivePairs28()
    Dim srcQ As String, srcA As String
    srcQ = "q1;;;q2;;;q3;;;q4;;;q5"
    srcA = "a1;;;a2;;;a3;;;a4;;;a5"
    Dim outQ As String, outA As String
    Dim ok As Boolean
    ok = modConvBridge.ComputeBridgeCore("rag", "normal", True, srcQ, srcA, "", "", 5, outQ, outA)
    modTestRunner.Check "R28-W3-6_5往復_戻り値True", ok, ""
    modTestRunner.Check "R28-W3-6_5往復_Q件数5", (PairsCount28(outQ) = 5), "outQ=" & outQ
    modTestRunner.Check "R28-W3-6_5往復_A件数5", (PairsCount28(outA) = 5), "outA=" & outA
    modTestRunner.Check "R28-W3-6_5往復_Q先頭はq1", (modConvBridge.FirstPair(outQ) = "q1"), ""
    modTestRunner.Check "R28-W3-6_5往復_Qにq5が含まれる(最古も運ばれる)", _
        (InStr(outQ, "q5") > 0), "outQ=" & outQ
    modTestRunner.Check "R28-W3-6_5往復_Aは出所ヘッダー付き", _
        (Left$(modConvBridge.FirstPair(outA), Len("【直前の")) = "【直前の"), "outA=" & outA
End Sub

' (E-2) 運搬総量20,016字クランプ(R28H F9で20,000から引き上げ): 1往復3,500字級を
'       6件渡すと総量を超える最も古い1件(P6)が切り捨てられ5件になる
'       (5件で 3500 + 3503×4 = 17,512字。6件目は +3,503 で 21,015字 > 20,016)。
'       A側は短文のみでmaxPairs6を満たし(クランプ非発動)、Q/Aそれぞれ独立に
'       クランプすることを確かめる。
Private Function PairText28(ByVal idx As Long, ByVal totalLen As Long) As String
    Dim tag As String: tag = "P" & idx & "_"
    PairText28 = tag & String(totalLen - Len(tag), "x")
End Function

Private Sub TestConvBridgeTotalClamp28()
    Dim srcQ As String
    Dim i As Long
    For i = 1 To 6
        If i > 1 Then srcQ = srcQ & ";;;"
        srcQ = srcQ & PairText28(i, 3500)
    Next i
    Dim srcA As String
    srcA = "a1;;;a2;;;a3;;;a4;;;a5;;;a6"   ' 短文6件(クランプに触れない側)

    Dim outQ As String, outA As String
    Dim ok As Boolean
    ok = modConvBridge.ComputeBridgeCore("rag", "normal", True, srcQ, srcA, "", "", 6, outQ, outA)
    modTestRunner.Check "R28-W3-6_総量クランプ_戻り値True", ok, ""
    modTestRunner.Check "R28-W3-6_総量クランプ_Qは5件に切り詰め", _
        (PairsCount28(outQ) = 5), "count=" & PairsCount28(outQ)
    modTestRunner.Check "R28-W3-6_総量クランプ_最古P6は含まれない", _
        (InStr(outQ, "P6_") = 0), "outQ長さ=" & Len(outQ)
    modTestRunner.Check "R28-W3-6_総量クランプ_新しいP1は含まれる", _
        (InStr(outQ, "P1_") > 0), ""
    modTestRunner.Check "R28-W3-6_総量クランプ_Aはクランプ非該当で6件", _
        (PairsCount28(outA) = 6), "count=" & PairsCount28(outA)

    ' R28H F9(m-5): 既定の上限いっぱい(1往復MAX_CARRY_CHARS=4,000字 × 5往復)が
    ' 総量クランプに弾かれずに入ること。旧値20,000では必要量20,012字
    ' (4,000×5 + 区切り3字×4)に12字足りず、5件目が黙って落ちていた。
    Dim srcQ2 As String
    For i = 1 To 6
        If i > 1 Then srcQ2 = srcQ2 & ";;;"
        srcQ2 = srcQ2 & PairText28(i, 4000)
    Next i
    ok = modConvBridge.ComputeBridgeCore("rag", "normal", True, srcQ2, "a1", "", "", 6, outQ, outA)
    modTestRunner.Check "R28H-F9_4000字×5往復が総量クランプに落ちない", _
        (PairsCount28(outQ) = 5), "count=" & PairsCount28(outQ)
    modTestRunner.Check "R28H-F9_運搬総量はちょうど20012字", _
        (Len(outQ) = 20012), "len=" & Len(outQ)
    modTestRunner.Check "R28H-F9_20016でも6件目は入らない", _
        (InStr(outQ, "P6_") = 0), "len=" & Len(outQ)
End Sub

' (E-3) 二重挿入防止: 切替先の先頭に運搬内容と同じ往復が既にあれば挿入しない
'       (往来を繰り返しただけで増殖・トースト連発しない。判定は先頭往復のみ)。
Private Sub TestConvBridgeAlreadyAtHead28()
    Dim srcQ As String, srcA As String
    srcQ = "q1;;;q2"
    srcA = "a1;;;a2"
    Dim headA As String
    headA = modConvBridge.WithBridgeHeader("社内ナレッジ検索", "a1")

    Dim dstQ As String, dstA As String
    dstQ = "q1;;;old_q"
    dstA = headA & ";;;old_a"

    Dim outQ As String, outA As String
    Dim ok As Boolean
    ok = modConvBridge.ComputeBridgeCore("rag", "normal", True, srcQ, srcA, dstQ, dstA, 5, outQ, outA)
    modTestRunner.Check "R28-W3-6_二重挿入防止_戻り値False", (ok = False), ""
    modTestRunner.Check "R28-W3-6_二重挿入防止_outQは空", (LenB(outQ) = 0), "outQ=" & outQ
    modTestRunner.Check "R28-W3-6_二重挿入防止_outAは空", (LenB(outA) = 0), "outA=" & outA
End Sub

' (E-4) クリア後(会話履歴が空)は運搬内容が無い=Falseで何も足さない。
Private Sub TestConvBridgeClearedIsEmpty28()
    Dim outQ As String, outA As String
    Dim ok As Boolean
    ok = modConvBridge.ComputeBridgeCore("rag", "normal", True, "", "", "keep_q", "keep_a", 5, outQ, outA)
    modTestRunner.Check "R28-W3-6_クリア後空_戻り値False", (ok = False), ""
    modTestRunner.Check "R28-W3-6_クリア後空_outQは空", (LenB(outQ) = 0), "outQ=" & outQ
    modTestRunner.Check "R28-W3-6_クリア後空_outAは空", (LenB(outA) = 0), "outA=" & outA
End Sub

' (E-5) 逆流(戻り)の重複除外(R28H F5/M-4)。
'       RAGで1問 → 一般へ切替 → 一般で1問 → RAGへ戻す、という実際の使い方で
'       切替先に既にある往復が2件目として差し込まれないことを固定する。
'       旧実装は先頭往復(q2)だけを見ていたため [q2, q1, q1] になっていた。
Private Function CountOccur28(ByVal hay As String, ByVal needle As String) As Long
    If LenB(needle) = 0 Then Exit Function
    Dim p As Long: p = InStr(1, hay, needle, vbBinaryCompare)
    Do While p > 0
        CountOccur28 = CountOccur28 + 1
        p = InStr(p + Len(needle), hay, needle, vbBinaryCompare)
    Loop
End Function

Private Sub TestConvBridgeNoBackflowDup28()
    Dim outQ As String, outA As String
    Dim ok As Boolean

    ' (1) RAG[q1/a1] → 一般(空)。一般側は [q1] になる。
    ok = modConvBridge.ComputeBridgeCore("rag", "normal", True, "q1", "a1", "", "", 5, outQ, outA)
    modTestRunner.Check "R28H-F5_往1_戻り値True", ok, ""
    Dim genQ As String, genA As String
    genQ = outQ: genA = outA

    ' (2) 一般で1問(q2/a2)。履歴は新しい順に積む。
    genQ = "q2;;;" & genQ
    genA = "a2;;;" & genA

    ' (3) 一般 → RAG へ戻す。RAG側は依然 [q1/a1] のまま。
    ok = modConvBridge.ComputeBridgeCore("normal", "rag", True, genQ, genA, "q1", "a1", 5, outQ, outA)
    modTestRunner.Check "R28H-F5_往2_戻り値True", ok, ""
    modTestRunner.Check "R28H-F5_往2_Qは2件(q1が重複しない)", _
        (PairsCount28(outQ) = 2), "outQ=" & outQ
    modTestRunner.Check "R28H-F5_往2_Q先頭はq2", _
        (modConvBridge.FirstPair(outQ) = "q2"), "outQ=" & outQ
    modTestRunner.Check "R28H-F5_往2_q1の出現は1回だけ", _
        (CountOccur28(outQ, "q1") = 1), "outQ=" & outQ

    ' (4) 何も足さずにもう一度往復させても増えない。ここは AlreadyAtHead では
    '     止まらない(A側にだけ出所ヘッダーが付き、先頭同士が文字列一致しない)。
    '     Q側の除外が全件を落とし、運ぶ物が無い=Falseで止まることを固定する。
    Dim ragQ As String, ragA As String
    ragQ = outQ: ragA = outA
    ok = modConvBridge.ComputeBridgeCore("rag", "normal", True, ragQ, ragA, genQ, genA, 5, outQ, outA)
    modTestRunner.Check "R28H-F5_往3_足す物が無ければFalse", (ok = False), "outQ=" & outQ
    modTestRunner.Check "R28H-F5_往3_outQは空", (LenB(outQ) = 0), "outQ=" & outQ
End Sub

' ----------------------------------------------------------------------------
' (F) modAsk.SetPrevMemory / ResetPrevMemory(R28 W4-1/W4-2)。
'     橋渡し(一般→RAG)が ui_state へ書くだけでは、modAsk.CanFollowup の
'     遅延ロードが「mPrevU が空のときしか読まない」ため効かない。直接受け渡し口
'     を通せば CanFollowup が True になり、ResetPrevMemory で False へ戻ること
'     (会話クリア後の亡霊=HANDOFF M-4 が消えること)を固定する。
'     LO では modState.LoadState が ui_state シート不在で既定値("")を返すため、
'     Reset 後の遅延ロードは空のまま=False になる(実Excelでも OnClearChat が
'     同じ2キーを "" にしてから呼ぶので同値)。
' ----------------------------------------------------------------------------
Private Sub TestAskPrevMemory28()
    modAsk.SetPrevMemory "前の質問", "前の回答"
    modTestRunner.Check "R28-W4_SetPrevMemory後はCanFollowup=True", _
        (modAsk.CanFollowup() = True), "CanFollowup=" & modAsk.CanFollowup()

    modAsk.ResetPrevMemory
    modTestRunner.Check "R28-W4_ResetPrevMemory後はCanFollowup=False", _
        (modAsk.CanFollowup() = False), "CanFollowup=" & modAsk.CanFollowup()

    ' 上書きできること(切替のたびに最新の橋渡し内容へ差し替わる)。
    modAsk.SetPrevMemory "新しい質問", "新しい回答"
    modTestRunner.Check "R28-W4_再Setで再びCanFollowup=True", _
        (modAsk.CanFollowup() = True), "CanFollowup=" & modAsk.CanFollowup()

    ' 後片付け: 以降のテストへ会話メモリを残さない。
    modAsk.ResetPrevMemory
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
NextToast28:
    On Error GoTo ToastFail28
    TestToastHeightFor28
NextPageCap28:
    On Error GoTo PageCapFail28
    TestPageCapFor28
NextBridgeFive28:
    On Error GoTo BridgeFiveFail28
    TestConvBridgeCarryFivePairs28
NextBridgeClamp28:
    On Error GoTo BridgeClampFail28
    TestConvBridgeTotalClamp28
NextBridgeDup28:
    On Error GoTo BridgeDupFail28
    TestConvBridgeAlreadyAtHead28
NextBridgeClear28:
    On Error GoTo BridgeClearFail28
    TestConvBridgeClearedIsEmpty28
NextBridgeBackflow28:
    On Error GoTo BridgeBackflowFail28
    TestConvBridgeNoBackflowDup28
NextAskMem28:
    On Error GoTo AskMemFail28
    TestAskPrevMemory28
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
    Resume NextToast28
ToastFail28:
    modTestRunner.Check "TestToastHeightFor28(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextPageCap28
PageCapFail28:
    modTestRunner.Check "TestPageCapFor28(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBridgeFive28
BridgeFiveFail28:
    modTestRunner.Check "TestConvBridgeCarryFivePairs28(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBridgeClamp28
BridgeClampFail28:
    modTestRunner.Check "TestConvBridgeTotalClamp28(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBridgeDup28
BridgeDupFail28:
    modTestRunner.Check "TestConvBridgeAlreadyAtHead28(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBridgeClear28
BridgeClearFail28:
    modTestRunner.Check "TestConvBridgeClearedIsEmpty28(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBridgeBackflow28
BridgeBackflowFail28:
    modTestRunner.Check "TestConvBridgeNoBackflowDup28(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextAskMem28
AskMemFail28:
    modTestRunner.Check "TestAskPrevMemory28(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone28
End Sub
