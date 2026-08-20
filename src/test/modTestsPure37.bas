Attribute VB_Name = "modTestsPure37"
Option Explicit

' ============================================================================
' modTestsPure37 - R34波1(外部レビュー裁定・層2の3件)の純ロジック回帰テスト。
'   modTestsPure36 と同じく既存チェーン(modTestsPure.RunAll→…)へは繋がず、
'   modTestRunner.RunAllPureTests から直接呼ばれる RunAll37 の1本が入口
'   (modTestsPure31〜36 と同型の別枝)。
' ----------------------------------------------------------------------------
' 【何を固定するか】
'   A1 modGenPipe.ParseVerdict: 1行目がPASSでも2行目以降(trailing)を無言で
'     捨てない。101字以上の懸念文はFINDINGSへ昇格、それ未満でも findings へ
'     載せて可観測にする(R34裁定)。
'   A2 modTelemetry.BakIsStale: summary_*.bak を消してよいか(24時間境界)を
'     決める純関数。ファイルI/O本体(BoardSweepStaleBak)はLO対象外のため
'     ここでは境界判定だけを固定する(実機観点として記録)。
'   A3 modChunker.bas:606 の Left$→SafeLeft 置換: AppendStructChunk は
'     Private かつ、呼び出し元の2箇所(FlushBlock直呼び出し/スライディング
'     窓)がいずれも渡す body の長さを room 以下に事前クランプしているため、
'     「Len(b) > room」の分岐は公開API(ChunkPagesEx)からは実質到達しない
'     防御的な保険であることを確認した(詳細はR34最終報告)。したがって
'     専用テストは追加せず、既存の modUtil.SafeLeft サロゲート境界テスト
'     (modTestsPure4.bas:392 TestSafeLeftSurrogate)で置換後の挙動は
'     足りている。
' ============================================================================

' ---- A1: ParseVerdict の PASS+trailing 可視化 ------------------------------
'   discriminate:
'   ・trailingを見ない(findingsを空のまま返す)実装だと(b)(a)双方が落ちる。
'   ・閾値を100以下(例:99)にすると、既存R26-1の9字装飾ケースに影響は
'     出ないが、本モジュールの100字ちょうどケースがFINDINGSへ倒れて落ちる。
'   ・閾値を101以上(例:102)にすると、本モジュールの101字ケースがPASS側へ
'     倒れて落ちる。
'   ・PASSのときfindingsへ全文(s)を詰める実装にすると、(b)(c)のfindings
'     一致チェックが落ちる(findings=trailingであってsではない)。
Private Sub TestParseVerdictTrailing37()
    Dim v As String, f As String

    ' (a) PASS+101字の懸念文→FINDINGSへ昇格。findingsは全文(s)。
    Dim long101 As String: long101 = String$(101, ChrW(&H3042)) ' "あ"×101
    Dim resp101 As String: resp101 = "verdict:PASS" & vbLf & long101
    v = modGenPipe.ParseVerdict(resp101, f)
    ChkStr37 "A1_PASS+101字はFINDINGSへ昇格", v, modGenPipe.VERDICT_FINDINGS
    ChkBool37 "A1_101字昇格時のfindingsは全文", (InStr(f, "verdict:PASS") > 0 And InStr(f, long101) > 0), True

    ' 境界: ちょうど100字はPASSのまま(閾値の根拠="100を超えたら")。
    Dim exact100 As String: exact100 = String$(100, ChrW(&H3042))
    Dim resp100 As String: resp100 = "verdict:PASS" & vbLf & exact100
    v = modGenPipe.ParseVerdict(resp100, f)
    ChkStr37 "A1_PASS+ちょうど100字はPASSのまま", v, modGenPipe.VERDICT_PASS
    ChkStr37 "A1_100字境界のfindingsはtrailingそのもの", f, exact100

    ' (b) PASS+短い装飾文→PASSのままだがfindingsに装飾文を載せる(無言消失にしない)。
    Dim resp20 As String: resp20 = "verdict:PASS" & vbLf & "(補足: 参考文献のページ番号だけ再確認)"
    v = modGenPipe.ParseVerdict(resp20, f)
    ChkStr37 "A1_PASS+短い装飾はPASSのまま", v, modGenPipe.VERDICT_PASS
    ChkStr37 "A1_短い装飾時のfindingsは装飾文そのもの", f, "(補足: 参考文献のページ番号だけ再確認)"

    ' (c) PASSのみ(2行目が無い)→従来通りPASSかつfindingsは空。
    v = modGenPipe.ParseVerdict("verdict:PASS", f)
    ChkStr37 "A1_PASSのみはPASSのまま", v, modGenPipe.VERDICT_PASS
    ChkBool37 "A1_PASSのみのfindingsは空", (LenB(f) = 0), True

    ' 回帰: R26-1の9字装飾ケース(既存テストの前提=このモジュールでも保つ)。
    v = modGenPipe.ParseVerdict("**verdict: PASS**" & vbLf & "(問題ありません)", f)
    ChkStr37 "A1_回帰_9字装飾でもPASS", v, modGenPipe.VERDICT_PASS
    ChkStr37 "A1_回帰_9字装飾はfindingsに残る", f, "(問題ありません)"
End Sub

' ---- A2: BakIsStale の24時間境界 -------------------------------------------
'   discriminate:
'   ・`>` を `>=` にすると「ちょうど24時間は消さない」が落ちる。
'   ・86400秒ではなく86400分等スケールを間違えると全ケースが逆転して落ちる。
'   ・引数の順序(fileTime, nowTime)を入れ替えると、新しいbakを「古い」と
'     判定するようになり(3)(4)が入れ替わって落ちる。
Private Sub TestBakIsStale37()
    Dim baseNow As Date: baseNow = DateSerial(2026, 8, 20) + TimeSerial(12, 0, 0)

    ' (1) ちょうど24時間前=消さない(境界は「超えたら」であって「以上」ではない)。
    ChkBool37 "A2_ちょうど24時間前は消さない", _
        modTelemetry.BakIsStale(baseNow - 1, baseNow), False

    ' (2) 24時間+1秒前=消す。
    ChkBool37 "A2_24時間と1秒前は消す", _
        modTelemetry.BakIsStale(baseNow - (1 + 1 / 86400#), baseNow), True

    ' (3) 24時間-1秒前=消さない。
    ChkBool37 "A2_24時間より1秒新しいと消さない", _
        modTelemetry.BakIsStale(baseNow - (1 - 1 / 86400#), baseNow), False

    ' (4) 作られたばかり(差0)=消さない。
    ChkBool37 "A2_作られたばかりは消さない", _
        modTelemetry.BakIsStale(baseNow, baseNow), False

    ' (5) 極端に古い(30日前)=消す。
    ChkBool37 "A2_30日前は消す", _
        modTelemetry.BakIsStale(baseNow - 30, baseNow), True
End Sub

Private Sub ChkBool37(ByVal label As String, ByVal got As Boolean, ByVal want As Boolean)
    modTestRunner.Check "R34-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkStr37(ByVal label As String, ByVal got As String, ByVal want As String)
    modTestRunner.Check "R34-" & label, (StrComp(got, want, vbBinaryCompare) = 0), _
        "実際=[" & got & "] 期待=[" & want & "]"
End Sub

Public Sub RunAll37()
    On Error GoTo H01Fail37
    TestParseVerdictTrailing37
H02Next37:
    On Error GoTo H02Fail37
    TestBakIsStale37
H01Done37:
    On Error GoTo 0
    Exit Sub

H01Fail37:
    modTestRunner.Check "TestParseVerdictTrailing37(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H02Next37
H02Fail37:
    modTestRunner.Check "TestBakIsStale37(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done37
End Sub
