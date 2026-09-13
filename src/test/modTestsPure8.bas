Attribute VB_Name = "modTestsPure8"
Option Explicit

' ============================================================================
' modTestsPure8 - R12-8(テスト補強)の続き。modTestsPure7の容量逼迫による分割先
' ----------------------------------------------------------------------------
' なぜ新しいモジュールなのか:
'   modTestsPure7がFnv1a64Hexゴールデン値+modRagParse4関数の追加で30,000字
'   上限に対する余裕が乏しくなったため、憲章§4-6「WARN帯のモジュールに機能を
'   足さない。足す前に分割を裁定する」に従って分割先を新設した。
'   入口は modTestsPure7.RunAll7 の末尾から呼ばれる Public Sub RunAll8()。
'   modTestRunner.RunAllPureTests は modTestsPure.RunAll だけを呼ぶ契約なので、
'   ここへの導線は modTestsPure7 内の1行だけ。消すとテストが「全部PASS」の
'   まま実行されなくなる(既存の分割群と同型の理由)。
'
' 固定する事実(監査「testmeta」の高リスク未テスト39件のうち残り分):
'   ・modBitwiseOpt: QuantizeToLongs/PopcountLong/HammingLongs の bit31(符号
'     ビット)境界。モジュール冒頭コメントは「Pythonリファレンスと一致検証済み
'     (オフライン)」とだけ書かれ、検証がリポジトリに残っていなかった。
'   ・modFollowup: SplitFollowupTrailer([[FOLLOWUP:/]]欠落・「なし」)、
'     KeepNewestPairs(件数境界)、SanitizeForFollowupHistory(縮約ループの
'     停止性)の全3関数。
'   ・modUtil.TruncateAndRenorm(次元打切り+再正規化。単位ベクトル化の誤差
'     1e-9以内)/HasVector(配列の割当てだけを見る契約。値がゼロでもTrue)。
'   ・modUtilText.ElapsedMsSince(Timer日跨ぎ境界)/BlendPerItemMs(初回/2回目
'     混合比)。
'   ・modClarify.MergeAnswer(LO環境で到達できる範囲の安全側フォールバック。
'     番号選択合成の全体はExcel実機専用である理由をTestMergeAnswerFallback
'     直前のコメントに明記)。
'   ・modPrompts.BuildRerankPrompt と modRagParse.ParseRankOrder の番号規約
'     (生成した候補番号がパースで往復すること)。
'
' ■ CanUseTypeArraysの複製について(modTestsPure2/3/4/7の冒頭コメントと同じ
'   理由): 他モジュールのPrivateは呼べないため、軽量な実測プローブを複製する。
'   Hit()配列(modPrompts.BuildRerankPrompt)を使うテストだけをこれで守る。
' ============================================================================

Private Function CanUseTypeArrays() As Boolean
    On Error Resume Next
    Err.Clear
    Dim probe() As Hit
    ReDim probe(0 To 0)
    CanUseTypeArrays = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' modBitwiseOpt(R12-8-2続き)- bit31(符号ビット)境界のゴールデン値
' ----------------------------------------------------------------------------
Private Sub TestBitwiseBoundaries()
    Dim vAllPos(0 To 31) As Double, vAllNeg(0 To 31) As Double
    Dim vBit31Only(0 To 31) As Double, vAlt(0 To 31) As Double
    Dim i As Long
    For i = 0 To 31
        vAllPos(i) = 1#
        vAllNeg(i) = -1#
        vBit31Only(i) = IIf(i = 31, 1#, -1#)
        vAlt(i) = IIf(i Mod 2 = 0, 1#, -1#)
    Next i

    Dim c1() As Long, c2() As Long, c3() As Long, c4() As Long
    modBitwiseOpt.QuantizeToLongs vAllPos, c1
    modBitwiseOpt.QuantizeToLongs vAllNeg, c2
    modBitwiseOpt.QuantizeToLongs vBit31Only, c3
    modBitwiseOpt.QuantizeToLongs vAlt, c4

    modTestRunner.Check "QuantizeToLongs_全次元非負→全bit1(&HFFFFFFFF)", _
        (c1(0) = &HFFFFFFFF), "実際=" & c1(0)
    modTestRunner.Check "QuantizeToLongs_全次元負→0", (c2(0) = 0), "実際=" & c2(0)
    modTestRunner.Check "QuantizeToLongs_bit31のみ非負→&H80000000", _
        (c3(0) = &H80000000), "実際=" & c3(0)
    modTestRunner.Check "QuantizeToLongs_偶数次元のみ非負→&H55555555", _
        (c4(0) = &H55555555), "実際=" & c4(0)

    ' 0.0は"v(i)>=0#"の境界で非負扱い(1次元だけで直接確認)
    Dim vZero(0 To 0) As Double
    vZero(0) = 0#
    Dim cZero() As Long
    modBitwiseOpt.QuantizeToLongs vZero, cZero
    modTestRunner.Check "QuantizeToLongs_0.0は非負としてbit0が立つ", (cZero(0) = 1), "実際=" & cZero(0)

    ' 40次元(2ワードにまたがる)量子化: 2ワード目にも正しく詰まる
    Dim v40(0 To 39) As Double
    For i = 0 To 39
        If i < 32 Then
            v40(i) = 1#
        Else
            v40(i) = IIf((i - 32) Mod 2 = 0, 1#, -1#)
        End If
    Next i
    Dim c40() As Long
    modBitwiseOpt.QuantizeToLongs v40, c40
    modTestRunner.Check "QuantizeToLongs_dims40_1ワード目は全1", (c40(0) = &HFFFFFFFF), "実際=" & c40(0)
    modTestRunner.Check "QuantizeToLongs_dims40_2ワード目は&H55(85)", (c40(1) = 85), "実際=" & c40(1)

    modTestRunner.Check "PopcountLong_全bit1(-1)は32", (modBitwiseOpt.PopcountLong(&HFFFFFFFF) = 32), ""
    modTestRunner.Check "PopcountLong_bit31のみ(&H80000000)は1", _
        (modBitwiseOpt.PopcountLong(&H80000000) = 1), ""
    modTestRunner.Check "PopcountLong_bit31以外全部(&H7FFFFFFF)は31", _
        (modBitwiseOpt.PopcountLong(&H7FFFFFFF) = 31), ""
    modTestRunner.Check "PopcountLong_0は0", (modBitwiseOpt.PopcountLong(0) = 0), ""
    modTestRunner.Check "PopcountLong_正の交互パターン(&H55555555)は16", _
        (modBitwiseOpt.PopcountLong(&H55555555) = 16), ""
    modTestRunner.Check "PopcountLong_負の交互パターン(&HAAAAAAAA・bit31側)は16", _
        (modBitwiseOpt.PopcountLong(&HAAAAAAAA) = 16), ""

    Dim ha(0 To 1) As Long, hb(0 To 1) As Long
    ha(0) = &HFFFFFFFF: ha(1) = 0
    hb(0) = &H55555555: hb(1) = &HAAAAAAAA
    modTestRunner.Check "HammingLongs_2ワードのXOR合算は32", _
        (modBitwiseOpt.HammingLongs(ha, hb) = 32), "実際=" & modBitwiseOpt.HammingLongs(ha, hb)
    modTestRunner.Check "HammingLongs_同一配列同士は0", (modBitwiseOpt.HammingLongs(ha, ha) = 0), ""
End Sub

' ----------------------------------------------------------------------------
' modFollowup(R12-8-2続き)- 「続けて質問」パース/履歴整形の全3関数
' ----------------------------------------------------------------------------
Private Sub TestFollowupSplitTrailer()
    Dim body As String, cands As String

    modFollowup.SplitFollowupTrailer "回答本文です。" & vbLf & "[[FOLLOWUP: 質問1 | 質問2]]", body, cands
    modTestRunner.Check "SplitFollowupTrailer_正常_本文からマーカー行を除去", _
        (body = "回答本文です。"), "body=[" & body & "]"
    modTestRunner.Check "SplitFollowupTrailer_正常_候補2件をvbLf区切りで抽出", _
        (cands = "質問1" & vbLf & "質問2"), "cands=[" & cands & "]"

    modFollowup.SplitFollowupTrailer "マーカーが無い普通の回答", body, cands
    modTestRunner.Check "SplitFollowupTrailer_マーカー欠落は本文そのまま", _
        (body = "マーカーが無い普通の回答" And cands = ""), "body=[" & body & "]"

    modFollowup.SplitFollowupTrailer "本文" & vbLf & "[[FOLLOWUP: なし]]", body, cands
    modTestRunner.Check "SplitFollowupTrailer_なし指定は候補なし・マーカー行は除去", _
        (body = "本文" And cands = ""), "body=[" & body & "] cands=[" & cands & "]"

    modFollowup.SplitFollowupTrailer "本文" & vbLf & "[[FOLLOWUP: 閉じ忘れ", body, cands
    modTestRunner.Check "SplitFollowupTrailer_]]欠落でも例外なく本文は残る", _
        (body = "本文" And cands = ""), "body=[" & body & "] cands=[" & cands & "]"

    modFollowup.SplitFollowupTrailer "", body, cands
    modTestRunner.Check "SplitFollowupTrailer_空入力は空のまま", (body = "" And cands = ""), ""

    modFollowup.SplitFollowupTrailer "本文" & vbLf & "[[FOLLOWUP:  Q1 ||  Q2  ]]" & vbLf & vbLf, body, cands
    modTestRunner.Check "SplitFollowupTrailer_空要素除去+末尾空行の刈込", _
        (body = "本文" And cands = "Q1" & vbLf & "Q2"), "body=[" & body & "] cands=[" & cands & "]"
End Sub

Private Sub TestFollowupKeepNewestPairs()
    modTestRunner.Check "KeepNewestPairs_0件指定は空文字", _
        (modFollowup.KeepNewestPairs("a;;;b;;;c", 0, ";;;") = ""), ""
    modTestRunner.Check "KeepNewestPairs_件数ちょうどは無変更", _
        (modFollowup.KeepNewestPairs("a;;;b", 2, ";;;") = "a;;;b"), ""
    modTestRunner.Check "KeepNewestPairs_件数未満は無変更", _
        (modFollowup.KeepNewestPairs("a", 5, ";;;") = "a"), ""
    modTestRunner.Check "KeepNewestPairs_上限超過は先頭N件のみ", _
        (modFollowup.KeepNewestPairs("a;;;b;;;c;;;d", 2, ";;;") = "a;;;b"), _
        "実際=[" & modFollowup.KeepNewestPairs("a;;;b;;;c;;;d", 2, ";;;") & "]"
    modTestRunner.Check "KeepNewestPairs_空入力+0件は空文字", _
        (modFollowup.KeepNewestPairs("", 0, ";;;") = ""), ""
    modTestRunner.Check "KeepNewestPairs_空入力+1件以上は空文字のまま", _
        (modFollowup.KeepNewestPairs("", 3, ";;;") = ""), ""
End Sub

Private Sub TestFollowupSanitizeHistory()
    modTestRunner.Check "SanitizeForFollowupHistory_区切り無しは無変更", _
        (modFollowup.SanitizeForFollowupHistory("hello world", ";;;") = "hello world"), ""
    modTestRunner.Check "SanitizeForFollowupHistory_単発の区切りは2文字に縮約", _
        (modFollowup.SanitizeForFollowupHistory("a;;;b", ";;;") = "a;;b"), _
        "実際=[" & modFollowup.SanitizeForFollowupHistory("a;;;b", ";;;") & "]"

    Dim r6 As String: r6 = modFollowup.SanitizeForFollowupHistory("a;;;;;;b", ";;;")
    modTestRunner.Check "SanitizeForFollowupHistory_連続6個は収束して区切りが残らない", _
        (InStr(r6, ";;;") = 0), "実際=[" & r6 & "]"
    modTestRunner.Check "SanitizeForFollowupHistory_連続6個の収束値", (r6 = "a;;b"), "実際=[" & r6 & "]"

    ' 縮約ループの停止性(各回で必ず短くなるので有限回で終わる契約)。
    ' 20個連続でもハングせず収束することを確認する。
    Dim manySep As String: manySep = "a" & String$(20, ";") & "b"
    Dim r20 As String: r20 = modFollowup.SanitizeForFollowupHistory(manySep, ";;;")
    modTestRunner.Check "SanitizeForFollowupHistory_20個連続でも有限回で収束", _
        (InStr(r20, ";;;") = 0), "実際=[" & r20 & "]"
    modTestRunner.Check "SanitizeForFollowupHistory_20個連続の収束値", (r20 = "a;;b"), "実際=[" & r20 & "]"
End Sub

' ----------------------------------------------------------------------------
' modUtil.TruncateAndRenorm / HasVector(R12-8-2続き)
' ----------------------------------------------------------------------------
Private Sub TestTruncateAndRenorm()
    ' R39 F003(外部受入テスト・実Excel で再現): 実装は打切り時に vec = cut と
    ' 配列を丸ごと代入するので、固定長配列を渡すと実Excel は Err 13(型が一致
    ' しません)。LibreOffice は通すため気付けなかった。打切りが起きる入力は
    ' 動的配列で作る(v2〜v5 は打切りが起きず代入に到達しないので従来どおり)。
    Dim v1() As Double: ReDim v1(0 To 4)
    v1(0) = 3#: v1(1) = 4#: v1(2) = 0#: v1(3) = 0#: v1(4) = 0#
    Dim ok1 As Boolean
    ok1 = modUtil.TruncateAndRenorm(v1, 2)
    modTestRunner.Check "TruncateAndRenorm_次元打切り(5→2)は成功", ok1, ""
    modTestRunner.Check "TruncateAndRenorm_打切り後は2要素", _
        (UBound(v1) - LBound(v1) + 1 = 2), "実際=" & (UBound(v1) - LBound(v1) + 1)
    modTestRunner.Check "TruncateAndRenorm_単位ベクトル化誤差1e-9以内(x=0.6)", _
        (Abs(v1(0) - 0.6) < 0.000000001), "v1(0)=" & v1(0)
    modTestRunner.Check "TruncateAndRenorm_単位ベクトル化誤差1e-9以内(y=0.8)", _
        (Abs(v1(1) - 0.8) < 0.000000001), "v1(1)=" & v1(1)
    modTestRunner.Check "TruncateAndRenorm_単位ベクトル化誤差1e-9以内(ノルム=1)", _
        (Abs(Sqr(v1(0) * v1(0) + v1(1) * v1(1)) - 1#) < 0.000000001), ""

    Dim v2(0 To 1) As Double
    v2(0) = 1#: v2(1) = 1#
    Dim ok2 As Boolean
    ok2 = modUtil.TruncateAndRenorm(v2, 2)
    modTestRunner.Check "TruncateAndRenorm_dims=n(打切りなし)は再正規化のみ", _
        (ok2 And Abs(v2(0) - 0.707106781) < 0.000001), "v2(0)=" & v2(0)

    Dim v3(0 To 1) As Double
    v3(0) = 3#: v3(1) = 4#
    Dim ok3 As Boolean
    ok3 = modUtil.TruncateAndRenorm(v3, 10)
    modTestRunner.Check "TruncateAndRenorm_dims>n(打切りなし)は再正規化のみ", _
        (ok3 And Abs(v3(0) - 0.6) < 0.000000001 And Abs(v3(1) - 0.8) < 0.000000001), ""

    Dim v4(0 To 2) As Double
    Dim ok4 As Boolean
    ok4 = modUtil.TruncateAndRenorm(v4, 5)
    modTestRunner.Check "TruncateAndRenorm_ゼロベクトル(dims>=n)はFalse・無変更", _
        (ok4 = False And v4(0) = 0# And v4(1) = 0# And v4(2) = 0#), ""

    Dim v5(0 To 2) As Double
    v5(0) = 1#: v5(1) = 2#: v5(2) = 3#
    Dim ok5 As Boolean
    ok5 = modUtil.TruncateAndRenorm(v5, 0)
    modTestRunner.Check "TruncateAndRenorm_dims<1はFalse・無変更", _
        (ok5 = False And v5(0) = 1# And v5(1) = 2# And v5(2) = 3#), ""

    Dim v6() As Double
    Dim ok6 As Boolean
    ok6 = modUtil.TruncateAndRenorm(v6, 3)
    modTestRunner.Check "TruncateAndRenorm_未初期化配列はFalse", (ok6 = False), ""
End Sub

Private Sub TestHasVector()
    Dim vUnassigned() As Double
    modTestRunner.Check "HasVector_未初期化配列はFalse", (modUtil.HasVector(vUnassigned) = False), ""

    ' 要素は既定で0.0(値がゼロでも配列としては存在する)。HasVectorは割当てだけを
    ' 見る契約(modRetrieve.bas:83/350のゼロベクトル関門)なので、これはTrueが正しい。
    Dim vZero(0 To 2) As Double
    modTestRunner.Check "HasVector_全要素0でも配列があればTrue(値ではなく割当てを見る)", _
        (modUtil.HasVector(vZero) = True), ""

    Dim vNonZero(0 To 2) As Double
    vNonZero(0) = 0.1: vNonZero(1) = 0.2: vNonZero(2) = 0.3
    modTestRunner.Check "HasVector_非ゼロベクトルはTrue", (modUtil.HasVector(vNonZero) = True), ""

    Dim vOne(0 To 0) As Double
    modTestRunner.Check "HasVector_1要素でもTrue", (modUtil.HasVector(vOne) = True), ""
End Sub

' ----------------------------------------------------------------------------
' modUtilText.ElapsedMsSince / BlendPerItemMs(R12-8-2続き)
' ----------------------------------------------------------------------------
' ElapsedMsSince は内部で Timer() を直接呼ぶため、テスト側から「今」を差し込む
' ことはできない。そこで t0 を「呼び出し時点のTimer()より確実に未来」の値に
' 固定する手法を使う: (Timer-t0)は実際の壁時計の値に関わらず必ず負になるため、
' 日跨ぎ救済(+MS_PER_DAY)の分岐を実行時刻に依存せず確定的に踏める。
Private Sub TestElapsedMsSinceDayCross()
    Dim tNear As Double: tNear = Timer + 5#   ' 5秒だけ未来(=日跨ぎ相当)
    Dim msNear As Double: msNear = modUtilText.ElapsedMsSince(tNear)
    modTestRunner.Check "ElapsedMsSince_日跨ぎ相当のt0は正の大きな値(丸1日弱)を返す", _
        (msNear > 80000000#), "実際=" & msNear
    modTestRunner.Check "ElapsedMsSince_日跨ぎ救済後も1日(86400000ms)を超えない", _
        (msNear <= 86400000#), "実際=" & msNear

    ' 1日を大きく超える未来(現実にはあり得ない壊れた時刻)を渡すと、日跨ぎ救済を
    ' 1回足してもなお負のままなので、最終防御(0未満なら0)で0に丸まる。
    Dim tFar As Double: tFar = Timer + 90000#   ' 25時間ぶん未来
    Dim msFar As Double: msFar = modUtilText.ElapsedMsSince(tFar)
    modTestRunner.Check "ElapsedMsSince_1日超未来のt0は最終防御で0", (msFar = 0#), "実際=" & msFar
End Sub

Private Sub TestBlendPerItemMs()
    ' itemCount<1は経過時間計算そのものに触れず即curMsを素通しする(短絡)。
    modTestRunner.Check "BlendPerItemMs_itemCount0は素通し", _
        (modUtilText.BlendPerItemMs(42#, 999#, 0) = 42#), ""
    modTestRunner.Check "BlendPerItemMs_itemCount負も素通し", _
        (modUtilText.BlendPerItemMs(7#, 0#, -1) = 7#), ""

    ' 経過0msを確定的に作る(t0をTimer()の1日以上未来にして最終防御で0にする。
    ' 上のElapsedMsSince境界テストと同じ手法)。これで初回/2回目の混合式そのものを
    ' 壁時計の揺れなしに検証できる。
    Dim t0Zero As Double: t0Zero = Timer + 90000#

    modTestRunner.Check "BlendPerItemMs_初回(curMs<=0)はthisMsをそのまま採用", _
        (modUtilText.BlendPerItemMs(0#, t0Zero, 3) = 0#), ""
    modTestRunner.Check "BlendPerItemMs_初回はcurMsが負でも同じ枝", _
        (modUtilText.BlendPerItemMs(-5#, t0Zero, 2) = 0#), ""
    modTestRunner.Check "BlendPerItemMs_2回目以降は(旧値+今回)/2の移動平均", _
        (modUtilText.BlendPerItemMs(200#, t0Zero, 1) = 100#), _
        "実際=" & modUtilText.BlendPerItemMs(200#, t0Zero, 1)
End Sub

' ----------------------------------------------------------------------------
' modClarify.MergeAnswer(R12-8-2続き)
' ----------------------------------------------------------------------------
' MergeAnswerは保留質問をmodState経由(ThisWorkbook.Worksheets)で読むため、
' Excelを持たないLO純ロジック実行では読み込みが必ず失敗する(On Error Resume
' Nextで自衛済み)。この環境で確実に検証できるのは「保留が読めない=origQが
' 空」の安全側フォールバック(返答をTrimしてそのまま新しい質問として通す)
' だけであり、番号選択の合成(PickSource/PickIntentの実行)自体はExcel実機
' でしか到達しない(両関数ともPrivateで直接テストできない)。監査(testmeta)
' 指摘のとおりこれはLO制約であり、実機側は docs/40 の実操作確認でカバーする。
Private Sub TestMergeAnswerFallback()
    Dim r As String
    r = modClarify.MergeAnswer("  普通の返答です  ")
    modTestRunner.Check "MergeAnswer_LO環境では保留なし扱い・Trimして素通し", _
        (r = "普通の返答です"), "実際=[" & r & "]"

    Dim r2 As String
    r2 = modClarify.MergeAnswer("2")
    modTestRunner.Check "MergeAnswer_番号だけの返答でも保留が読めなければ素通し", _
        (r2 = "2"), "実際=[" & r2 & "]"

    modTestRunner.Check "MergeAnswer_保留読込失敗でも例外を出さず戻る", True, ""
End Sub

' ----------------------------------------------------------------------------
' modPrompts.BuildRerankPrompt <-> modRagParse.ParseRankOrder の番号規約突合
' ----------------------------------------------------------------------------
Private Sub TestRerankPromptRoundTrip()
    If Not CanUseTypeArrays() Then
        modTestRunner.Check "[SKIP] BuildRerankPrompt<->ParseRankOrder往復はLO環境の制限によりスキップ", True, _
            "Hit()配列のReDimがエラーになる環境。番号規約はParseRankOrder単体側で固定済み。"
        Exit Sub
    End If

    Dim hits(1 To 3) As Hit
    hits(1).chunk_id = "c1": hits(1).source = "資料A": hits(1).page = 1
    hits(1).origin = "": hits(1).full_text = "第1条の本文"
    hits(2).chunk_id = "c2": hits(2).source = "資料B": hits(2).page = 5
    hits(2).origin = "": hits(2).full_text = "第2条の本文"
    hits(3).chunk_id = "c3": hits(3).source = "資料C": hits(3).page = 9
    hits(3).origin = "": hits(3).full_text = "第3条の本文"

    Dim prompt As String
    prompt = modPrompts.BuildRerankPrompt("質問文", hits, 3, 4000)
    modTestRunner.Check "BuildRerankPrompt_候補番号は1起点で連番", _
        (InStr(prompt, "[1] ") > 0 And InStr(prompt, "[2] ") > 0 And InStr(prompt, "[3] ") > 0), _
        "prompt=[" & prompt & "]"

    ' LLMが忠実にこの番号規約で<rank>…</rank>を返したと仮定し、ParseRankOrderが
    ' 同じ1..nHits範囲として正しく往復することを確認する(両側無検証だった対)。
    Dim order() As Long, cnt As Long
    cnt = modRagParse.ParseRankOrder("<rank>3,1,2</rank>", 3, order)
    modTestRunner.Check "ParseRankOrder_BuildRerankPromptの番号規約と往復一致", _
        (cnt = 3 And order(0) = 3 And order(1) = 1 And order(2) = 2), "cnt=" & cnt
End Sub

' ----------------------------------------------------------------------------
' 会話出典メモリ(2026-08-03 R13-5b)
' ----------------------------------------------------------------------------
' 「深掘り=会話の流れ(引用済み資料)の中を深く」の材料そのもの。
' 順番(新しい順)と上限(8件)が崩れると、深掘りのスコープが古い話題を
' 引きずったり本棚全体に近づいたりして、モードの区別が静かに消える。
Private Sub TestCitedSourceMemory()
    ' 入力は modAskRetrieve.HitSourceList と同じ "|" 区切り。
    Dim a As String
    a = modFollowup.MergeCitedSources("", "資料A|資料B", 8)
    modTestRunner.Check "会話出典_初回は与えた順に積む", _
        (a = "資料A" & vbLf & "資料B"), "got=[" & Replace(a, vbLf, "/") & "]"

    ' 直近ターンの資料が先頭へ来る(最後に引用した資料が最優先で残る)。
    Dim b As String
    b = modFollowup.MergeCitedSources(a, "資料C|資料A", 8)
    modTestRunner.Check "会話出典_新しい順で重複を除く", _
        (b = "資料C" & vbLf & "資料A" & vbLf & "資料B"), "got=[" & Replace(b, vbLf, "/") & "]"

    Dim c As String
    c = modFollowup.MergeCitedSources("", "s1|s2|s3|s4|s5|s6|s7|s8|s9", 8)
    modTestRunner.Check "会話出典_上限8件で打ち切る", _
        (UBound(Split(c, vbLf)) = 7 And InStr(vbLf & c & vbLf, vbLf & "s9" & vbLf) = 0), _
        "got=[" & Replace(c, vbLf, "/") & "]"

    modTestRunner.Check "会話出典_空入力は空のまま", _
        (modFollowup.MergeCitedSources("", "", 8) = ""), "空にならない"
    modTestRunner.Check "会話出典_空白だけの名前は捨てる", _
        (modFollowup.MergeCitedSources("", " | ", 8) = ""), "空白が残る"
End Sub

' ----------------------------------------------------------------------------
' スコープ判定と同値性(2026-08-03 R13-5a)
' ----------------------------------------------------------------------------
' modRetrieve は許可Dictionaryの .Exists で行を弾く。Scripting.Dictionary は
' 実機(Windows)専用でLOからは作れないため、ここでは【判定の意味】を
' modFollowup.InCitedScope(検索側と同じ完全一致の定義。Dictionaryを作る
' ScopeDictFrom もこの規則で名前を積む)で固定する。
' 検収条件は「スコープを掛けた結果 = 掛けない結果からスコープ外を抜いたもの」
' で、順位が変わらないこと。
Private Sub TestCitedScopeFilter()
    Dim scopeLine As String
    scopeLine = "資料A" & vbLf & "資料C"

    modTestRunner.Check "スコープ_完全一致なら入る", _
        modFollowup.InCitedScope(scopeLine, "資料A"), ""
    modTestRunner.Check "スコープ_前方一致では入らない", _
        (modFollowup.InCitedScope(scopeLine, "資料AB") = False), "部分一致が通っている"
    modTestRunner.Check "スコープ_未収録の資料は落ちる", _
        (modFollowup.InCitedScope(scopeLine, "資料B") = False), ""
    modTestRunner.Check "スコープ_空リストは何も通さない(無制限はNothingで表す)", _
        (modFollowup.InCitedScope("", "資料A") = False), ""

    Dim ranked As Variant
    ranked = Array("資料B", "資料A", "資料D", "資料C", "資料A")
    Dim unscoped As String, scoped As String
    Dim i As Long
    For i = LBound(ranked) To UBound(ranked)
        If LenB(unscoped) > 0 Then unscoped = unscoped & "|"
        unscoped = unscoped & CStr(ranked(i))
        If modFollowup.InCitedScope(scopeLine, CStr(ranked(i))) Then
            If LenB(scoped) > 0 Then scoped = scoped & "|"
            scoped = scoped & CStr(ranked(i))
        End If
    Next i
    modTestRunner.Check "スコープ同値性_落ちるのはスコープ外だけ・順位は不変", _
        (scoped = "資料A|資料C|資料A"), "unscoped=[" & unscoped & "] scoped=[" & scoped & "]"
End Sub

' ----------------------------------------------------------------------------
' 段階ナレーションの番号(2026-08-03 R13-9b)
' ----------------------------------------------------------------------------
' 番号は「その回に実際に通す段の数」から作る。総数を4に固定すると、
' 拡張も再ランクも通さない既定構成で「(3/4)で終わる」という嘘になる。
Private Sub TestAskStageNumbering()
    modTestRunner.Check "段階_すぐ聞くは番号を出さない", _
        (modMode.AskStageTotal(False, False, False) = 0), _
        "total=" & modMode.AskStageTotal(False, False, False)
    modTestRunner.Check "段階_すぐ聞くのラベルは素のまま", _
        (modMode.AskStageText(modMode.AskStageIndex("quick", False, False), 0, _
                              modMode.AskStageLabel("quick")) = "回答を作成中…"), ""

    modTestRunner.Check "段階_拡張も再ランクも無い深掘りは2段", _
        (modMode.AskStageTotal(False, False, True) = 2), ""
    modTestRunner.Check "段階_2段構成の下書きは(1/2)", _
        (modMode.AskStageText(modMode.AskStageIndex("draft", False, False), 2, _
                              modMode.AskStageLabel("draft")) = "(1/2) 下書きを作成中…"), ""
    modTestRunner.Check "段階_2段構成の検証は(2/2)", _
        (modMode.AskStageText(modMode.AskStageIndex("verify", False, False), 2, _
                              modMode.AskStageLabel("verify")) = "(2/2) 検証中…"), ""

    Dim t As Long: t = modMode.AskStageTotal(True, True, True)
    modTestRunner.Check "段階_全段構成は4段", (t = 4), "total=" & t
    modTestRunner.Check "段階_(1/4)質問を分解中", _
        (modMode.AskStageText(modMode.AskStageIndex("expand", True, True), t, _
                              modMode.AskStageLabel("expand")) = "(1/4) 質問を分解中…"), ""
    modTestRunner.Check "段階_(2/4)資料を照合中", _
        (modMode.AskStageText(modMode.AskStageIndex("rerank", True, True), t, _
                              modMode.AskStageLabel("rerank")) = "(2/4) 資料を照合中…"), ""
    modTestRunner.Check "段階_(3/4)下書きを作成中", _
        (modMode.AskStageText(modMode.AskStageIndex("draft", True, True), t, _
                              modMode.AskStageLabel("draft")) = "(3/4) 下書きを作成中…"), ""
    modTestRunner.Check "段階_(4/4)検証中", _
        (modMode.AskStageText(modMode.AskStageIndex("verify", True, True), t, _
                              modMode.AskStageLabel("verify")) = "(4/4) 検証中…"), ""

    ' 計画に入っていない段は番号を持たない(素のラベルへ退化する)。
    modTestRunner.Check "段階_通さない段には番号を付けない", _
        (modMode.AskStageIndex("rerank", True, False) = 0), ""
End Sub

' ----------------------------------------------------------------------------
' 3モードの役割説明(2026-08-03 R13-5d)
' ----------------------------------------------------------------------------
' deep と thorough が「同じことをパラメータ違いでやる」状態を直した以上、
' 説明も役割で言い分ける。ここが元に戻ると、利用者から見て3つ目のモードは
' 「ただ遅いだけの選択肢」に戻る。
Private Sub TestModeRoleWording()
    modTestRunner.Check "モード説明_すぐ聞く=まず速く", _
        (InStr(modMode.Description("quick"), "まず速く") > 0), modMode.Description("quick")
    modTestRunner.Check "モード説明_しっかり調べる=会話の流れの中を深く", _
        (InStr(modMode.Description("deep"), "会話の流れ") > 0), modMode.Description("deep")
    modTestRunner.Check "モード説明_入念に調べる=本棚全体を広く", _
        (InStr(modMode.Description("thorough"), "本棚全体") > 0), modMode.Description("thorough")
End Sub

Public Sub RunAll8()
    On Error GoTo BitwiseFail
    TestBitwiseBoundaries
NextSplitTrailer:
    On Error GoTo SplitTrailerFail
    TestFollowupSplitTrailer
NextKeepPairs:
    On Error GoTo KeepPairsFail
    TestFollowupKeepNewestPairs
NextSanitizeHistory:
    On Error GoTo SanitizeHistoryFail
    TestFollowupSanitizeHistory
NextTruncate:
    On Error GoTo TruncateFail
    TestTruncateAndRenorm
NextHasVector:
    On Error GoTo HasVectorFail
    TestHasVector
NextElapsed:
    On Error GoTo ElapsedFail
    TestElapsedMsSinceDayCross
NextBlend:
    On Error GoTo BlendFail
    TestBlendPerItemMs
NextMerge:
    On Error GoTo MergeFail
    TestMergeAnswerFallback
NextRerank:
    On Error GoTo RerankFail
    TestRerankPromptRoundTrip
NextCited:
    On Error GoTo CitedFail
    TestCitedSourceMemory
NextScope:
    On Error GoTo ScopeFail
    TestCitedScopeFilter
NextStageNum:
    On Error GoTo StageNumFail
    TestAskStageNumbering
NextModeWord:
    On Error GoTo ModeWordFail
    TestModeRoleWording
NextRun9:
    ' 2026-08-01(R12-4): 容量のための分割先(modTestsPure9)。ここが唯一の
    ' 導線で、消すとR12-4のテストが「実行されないまま」全部PASSに見える。
    On Error GoTo Run9Fail
    modTestsPure9.RunAll9
NextDone8:
    On Error GoTo 0
    Exit Sub

BitwiseFail:
    modTestRunner.Check "TestBitwiseBoundaries(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextSplitTrailer
SplitTrailerFail:
    modTestRunner.Check "TestFollowupSplitTrailer(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextKeepPairs
KeepPairsFail:
    modTestRunner.Check "TestFollowupKeepNewestPairs(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextSanitizeHistory
SanitizeHistoryFail:
    modTestRunner.Check "TestFollowupSanitizeHistory(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextTruncate
TruncateFail:
    modTestRunner.Check "TestTruncateAndRenorm(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextHasVector
HasVectorFail:
    modTestRunner.Check "TestHasVector(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextElapsed
ElapsedFail:
    modTestRunner.Check "TestElapsedMsSinceDayCross(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBlend
BlendFail:
    modTestRunner.Check "TestBlendPerItemMs(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextMerge
MergeFail:
    modTestRunner.Check "TestMergeAnswerFallback(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRerank
RerankFail:
    modTestRunner.Check "TestRerankPromptRoundTrip(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextCited
CitedFail:
    modTestRunner.Check "TestCitedSourceMemory(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextScope
ScopeFail:
    modTestRunner.Check "TestCitedScopeFilter(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextStageNum
StageNumFail:
    modTestRunner.Check "TestAskStageNumbering(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextModeWord
ModeWordFail:
    modTestRunner.Check "TestModeRoleWording(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRun9
Run9Fail:
    modTestRunner.Check "modTestsPure9.RunAll9(モジュール全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone8
End Sub
