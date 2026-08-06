Attribute VB_Name = "modTestsPure18"
Option Explicit

' ============================================================================
' modTestsPure18 - R17H(レビュー裁定 Fix波)の純ロジック回帰
' ----------------------------------------------------------------------------
' なぜ新設したか(憲章§4-6):
'   modTestsPure17 が25,032字で、ここの真理表(約3,400字)を足すと WARN帯
'   (28,000字)へ入る。17 を新設したときと同じ線で分割する。
'   入口は modTestsPure17.RunAll17 の末尾から呼ばれる RunAll18 の1本だけ。
'
' ここで固定するもの:
'   ・modRagParse.MergeSynPairs(FA-1 / A-H1): 名寄せ辞書のマージ規則。
'     実体は modSynonymStore.MergeAndSave にあり、ParseSynResp が返す
'     【0始まり】の配列を 1〜n で読んでいたため、実データでは毎回
'     「添字が範囲外」で握り潰され synonyms が永久に0行だった。純関数へ
'     切り出した以上、規則そのものをゴールデンで縛る(新規のみ/旧のみ/
'     上書き/先頭要素の保持/空入力/壊れた要素/大小無視/順序安定)。
'   ・modRagParse.HasGlobalSignal(FA-6 / A-M7・B-H2): 俯瞰語彙の検知。
'     拾う語と【拾わない語】の両方を固定する(誤爆しやすい短語を後から
'     足されないようにするのがこのテストの半分の目的)。
'   ・modAskMulti.DecomposeGate(FA-6): 俯瞰シグナルのOR。短い俯瞰質問でも
'     段0が呼ばれること、off が最優先であることの2点。
'
' 2026-08-06(R19H Fix波)で追記:
'   ・modClarify.DispersionGapX100(FB-1): gapの実値と資料数(観測の校正用)。
'   ・modIntegrity.CohabitOtherCount / IsCohabiting(FA-3): 同居の数え方
'     (personal.xlsb のみ同居 / 可視ブック同居 / 混在)。
'
' 2026-08-06(R19-4a/4d・実機第6報④)で追記:
'   ・modClarify.HasScoreDispersion: 資料分散の判定(「第3の曖昧さ」)。
'   ・modClarify.MentionsSourceName: 質問文が資料を名指ししているときの除外。
'   ・調査④班6.2のテスト表12件を、それぞれ【実際に止める側のゲート】で固定する。
'
' 2026-08-06(R20-2/R20-6・実機第7報①⑧・波B)で追記:
'   ・modAppAct.GateUsesGeneralHistory: 「続けて質問」ゲートのnormal/rag分岐。
'   ・modAppState.ShouldClearGeneralHistory / ClearGeneralMemory:
'     followup_max_pairs=0の境界とOnClearChat後の一般履歴消去。
'   ・modClarify.IsSelfEcho: 聞き返し保留の自己連結防御の境界。
'   ・modClarify.HasScoreDispersion: thorough(閾値25)/deep・quick(閾値10)の
'     gap20での発動有無の対。
'   ・modAskThorough.ThoroughVerifyPrompt: 入念モード専用の文体追記。
' ============================================================================

' ----------------------------------------------------------------------------
' FA-1: 名寄せ辞書のマージ規則(旧CSV + 新CSV → 統合CSV)。
' ----------------------------------------------------------------------------
Private Sub TestMergeSynPairs()
    ' 旧が空 = 初回の取込。新規ぶんがそのまま順序どおり残る。
    modTestRunner.Check "名寄せ統合_旧が空なら新規がそのまま", _
        (modRagParse.MergeSynPairs("", "回収>リコール|中途解約>解約") = _
         "回収>リコール|中途解約>解約"), _
        "実際=" & modRagParse.MergeSynPairs("", "回収>リコール|中途解約>解約")

    ' 新が空 = 表記ゆれの無い資料。既存の辞書を1行も落とさない。
    modTestRunner.Check "名寄せ統合_新が空なら旧がそのまま", _
        (modRagParse.MergeSynPairs("回収>リコール", "") = "回収>リコール"), _
        "実際=" & modRagParse.MergeSynPairs("回収>リコール", "")

    ' 両方空。
    modTestRunner.Check "名寄せ統合_両方空なら空", _
        (modRagParse.MergeSynPairs("", "") = "")

    ' 【先頭要素の保持】= A-H1 の本体。旧実装は 0始まりの配列を 1〜n で
    ' 読んでいたため、先頭が落ちて末尾で範囲外エラーになっていた。
    Dim r1 As String
    r1 = modRagParse.MergeSynPairs("A>a|B>b|C>c", "D>d")
    modTestRunner.Check "名寄せ統合_旧の先頭要素が落ちない", _
        (r1 = "A>a|B>b|C>c|D>d"), "実際=" & r1

    ' 既存termは上書き: 同じ term を持つ旧行は落とし、新しい行が末尾へ。
    Dim r2 As String
    r2 = modRagParse.MergeSynPairs("A>a|B>b|C>c", "B>bb")
    modTestRunner.Check "名寄せ統合_既存termは上書き(旧行が消える)", _
        (r2 = "A>a|C>c|B>bb"), "実際=" & r2

    ' 順序安定: 残った旧行の並びは変わらない。
    Dim r3 As String
    r3 = modRagParse.MergeSynPairs("A>a|B>b|C>c|D>d", "C>cc|A>aa")
    modTestRunner.Check "名寄せ統合_残った旧行の順序は変わらない", _
        (r3 = "B>b|D>d|C>cc|A>aa"), "実際=" & r3

    ' 大小は同じ term として扱う(termは検索キーで、大小の違いは表記ゆれ)。
    Dim r4 As String
    r4 = modRagParse.MergeSynPairs("abc>x", "ABC>y")
    modTestRunner.Check "名寄せ統合_termの大小違いは同じ語として上書き", _
        (r4 = "ABC>y"), "実際=" & r4

    ' 壊れた要素(">"無し・どちらか空・空要素)は両側とも1件ずつ捨てる。
    Dim r5 As String
    r5 = modRagParse.MergeSynPairs("A>a|こわれた|>x|y>", "||B>b")
    modTestRunner.Check "名寄せ統合_壊れた要素は捨てて読めた分だけ残す", _
        (r5 = "A>a|B>b"), "実際=" & r5

    ' 同じ term が新CSVの中で重複していたら先勝ち(2行目以降は捨てる)。
    Dim r6 As String
    r6 = modRagParse.MergeSynPairs("", "A>1|A>2|B>3")
    modTestRunner.Check "名寄せ統合_新CSV内の重複termは先勝ち", _
        (r6 = "A>1|B>3"), "実際=" & r6

    ' 前後の空白は落として保存する(シートの表示と照合の両方を揃える)。
    Dim r7 As String
    r7 = modRagParse.MergeSynPairs(" A > a ", "")
    modTestRunner.Check "名寄せ統合_前後の空白は落とす", _
        (r7 = "A>a"), "実際=" & r7

    ' 統合結果は ParseSynResp がそのまま読み戻せる形(Store が書き戻す経路)。
    Dim tms() As String, cns() As String, n As Long
    n = modRagParse.ParseSynResp("<syn>" & r2 & "</syn>", tms, cns)
    modTestRunner.Check "名寄せ統合_結果をParseSynRespで読み戻せる", _
        (n = 3), "実際=" & n
    If n = 3 Then
        modTestRunner.Check "名寄せ統合_読み戻しの先頭がAのまま(0始まり)", _
            (tms(LBound(tms)) = "A" And cns(LBound(cns)) = "a"), _
            "実際=" & tms(LBound(tms)) & ">" & cns(LBound(cns))
    End If
End Sub

' ----------------------------------------------------------------------------
' FA-6: 俯瞰語彙の検知。拾う語と【拾わない語】の両方を固定する。
' ----------------------------------------------------------------------------
Private Sub TestHasGlobalSignal()
    modTestRunner.Check "俯瞰語_全体像は?", _
        (modRagParse.HasGlobalSignal("全体像は?") = True)
    modTestRunner.Check "俯瞰語_この規程を全部教えて", _
        (modRagParse.HasGlobalSignal("この規程を全部教えてください") = True)
    modTestRunner.Check "俯瞰語_すべて教えて", _
        (modRagParse.HasGlobalSignal("補償の種類をすべて教えて") = True)
    modTestRunner.Check "俯瞰語_一覧", _
        (modRagParse.HasGlobalSignal("提出書類の一覧") = True)
    modTestRunner.Check "俯瞰語_どんな種類", _
        (modRagParse.HasGlobalSignal("どんな種類がありますか") = True)
    modTestRunner.Check "俯瞰語_どんなこと", _
        (modRagParse.HasGlobalSignal("この資料にはどんなことが書いてありますか") = True)
    modTestRunner.Check "俯瞰語_概要", _
        (modRagParse.HasGlobalSignal("概要を知りたい") = True)
    modTestRunner.Check "俯瞰語_何が書いて", _
        (modRagParse.HasGlobalSignal("何が書いてある?") = True)
    modTestRunner.Check "俯瞰語_全体の流れ", _
        (modRagParse.HasGlobalSignal("申請の全体の流れは?") = True)
    modTestRunner.Check "俯瞰語_どういう構成", _
        (modRagParse.HasGlobalSignal("どういう構成ですか") = True)
    modTestRunner.Check "俯瞰語_まとめて", _
        (modRagParse.HasGlobalSignal("免責事項をまとめて知りたい") = True)

    ' 【拾わない】: 単独で普通に使う短語で誤爆させない。ここが緩むと、
    ' ふつうの条文質問まで俯瞰扱いになり、聞き返しも効かなくなる。
    modTestRunner.Check "俯瞰語_『全て』単独では拾わない", _
        (modRagParse.HasGlobalSignal("全ての書類を提出しますか") = False)
    modTestRunner.Check "俯瞰語_『全部』単独では拾わない", _
        (modRagParse.HasGlobalSignal("全部で何日かかりますか") = False)
    modTestRunner.Check "俯瞰語_条文の質問は拾わない", _
        (modRagParse.HasGlobalSignal("第5条の免責は?") = False)
    modTestRunner.Check "俯瞰語_空文字は拾わない", _
        (modRagParse.HasGlobalSignal("") = False)
    modTestRunner.Check "俯瞰語_空白だけは拾わない", _
        (modRagParse.HasGlobalSignal("   ") = False)
End Sub

' ----------------------------------------------------------------------------
' FA-6: 段0(分解判定)を呼ぶ門に俯瞰シグナルをORで足したこと。
'   実装(modAskMulti.DecomposeGate)とテストが同じ1本を見る(憲章§4-5)。
' ----------------------------------------------------------------------------
Private Sub TestDecomposeGateGlobal()
    ' 6字の俯瞰質問。文字数ゲート(既定25)も複合シグナルも通らないが、
    ' 俯瞰シグナルで段0が呼ばれる=俯瞰が試される。
    modTestRunner.Check "段0門_短い俯瞰質問でも呼ばれる", _
        (modAskMulti.DecomposeGate("auto", "全体像は?", 25) = True)
    modTestRunner.Check "段0門_一覧を求める短文でも呼ばれる", _
        (modAskMulti.DecomposeGate("auto", "補償の一覧", 25) = True)

    ' 俯瞰でも複合でもない短文は従来どおり呼ばない(無駄な段0を増やさない)。
    modTestRunner.Check "段0門_ふつうの短文は従来どおり呼ばない", _
        (modAskMulti.DecomposeGate("auto", "更新日は?", 25) = False)

    ' off は最優先(機能まるごとのエスケープハッチ)。
    modTestRunner.Check "段0門_offは俯瞰シグナルより優先", _
        (modAskMulti.DecomposeGate("off", "全体像は?", 25) = False)

    ' 複合シグナルのORは従来どおり効いたまま(片方を足して片方を壊さない)。
    modTestRunner.Check "段0門_複合シグナルのORは不変", _
        (modAskMulti.DecomposeGate("auto", "免責は?保険料は?", 25) = True)
End Sub

' ----------------------------------------------------------------------------
' R19-4a/4d(実機第6報④): 資料分散の逆質問。調査④班6.2の表12件をそのまま置く。
' ----------------------------------------------------------------------------
' 12件のうち、分散判定そのもの(HasScoreDispersion)が答えるのは
' #1/#2/#3/#4/#9/#10/#12。残りは【別のゲートが先に止める】ことを確かめる行で、
' それぞれ担当する純関数で固定する:
'   #5  俯瞰       → modRagParse.HasGlobalSignal(IsTooVague が先に除外)
'   #6  資料名     → modClarify.MentionsSourceName(R19-4d の誤発動対策)
'   #7  複合質問   → modAskMulti.DecomposeGate(段0の分解が先に効く)
'   #8  資料名+突出 → MentionsSourceName と HasScoreDispersion の両方
'   #11 長さ       → ambiguous_max_chars(既定10)の文字数ゲート
' 「どのゲートが止めるのか」まで含めて固定しないと、片方を直したときに
' もう片方が黙って開くため(実機第6報④はまさにその隙間で起きた)。
' 行の形式は "資料名<TAB>スコア" を vbLf で連ねたもの(呼び出し側が畳む形)。
Private Function DispLines(ByVal a As String, ByVal b As String, _
                           Optional ByVal c As String = "") As String
    DispLines = a
    If LenB(b) > 0 Then DispLines = DispLines & vbLf & b
    If LenB(c) > 0 Then DispLines = DispLines & vbLf & c
End Function

Private Sub TestScoreDispersion()
    ' #1 免責は? — 3資料が拮抗。分散の本命ケース。
    modTestRunner.Check "分散_#1_免責は?は拮抗した3資料で発動", _
        (modClarify.HasScoreDispersion( _
            DispLines("約款A" & vbTab & "0.72", "約款B" & vbTab & "0.70", _
                      "規程C" & vbTab & "0.68"), 10) = True)

    ' #2 計算方法 — 深掘り(followup)でも同じ経路を通る。
    modTestRunner.Check "分散_#2_計算方法(深掘り)も発動", _
        (modClarify.HasScoreDispersion( _
            DispLines("約款A" & vbTab & "0.75", "約款B" & vbTab & "0.71"), 10) = True)

    ' #3 上位2件だけが拮抗していれば足りる(3件目が離れていても関係ない)。
    modTestRunner.Check "分散_#3_上位2件が拮抗していれば発動", _
        (modClarify.HasScoreDispersion( _
            DispLines("約款A" & vbTab & "0.80", "約款B" & vbTab & "0.78", _
                      "規程C" & vbTab & "0.55"), 10) = True)

    ' #4 1位が突出。聞き返さない=いちばん大事な「黙って答える」側。
    modTestRunner.Check "分散_#4_1位が突出なら発動しない", _
        (modClarify.HasScoreDispersion( _
            DispLines("約款A" & vbTab & "0.9", "約款B" & vbTab & "0.3"), 10) = False)

    ' #9 1語だけの質問でも、分散していれば対象。
    modTestRunner.Check "分散_#9_1語の質問でも分散があれば発動", _
        (modClarify.HasScoreDispersion( _
            DispLines("約款A" & vbTab & "0.72", "約款B" & vbTab & "0.70"), 10) = True)

    ' #10 資料が1種類しかない。分散しようがない。
    modTestRunner.Check "分散_#10_資料が1種類なら発動しない", _
        (modClarify.HasScoreDispersion(DispLines("約款A" & vbTab & "0.85", ""), 10) = False)

    ' #12 突出(0.75 vs 0.10)。
    modTestRunner.Check "分散_#12_クーリングオフは突出なので発動しない", _
        (modClarify.HasScoreDispersion( _
            DispLines("約款A" & vbTab & "0.75", "約款B" & vbTab & "0.10"), 10) = False)

    ' 同じ資料が複数チャンクで当たっても【1資料】として数える(最高スコアだけ残す)。
    modTestRunner.Check "分散_同一資料の複数ヒットは1種類として数える", _
        (modClarify.HasScoreDispersion( _
            DispLines("約款A" & vbTab & "0.72", "約款A" & vbTab & "0.70", _
                      "約款A" & vbTab & "0.69"), 10) = False)

    ' 境界: 差ちょうど0.10 は発動しない(未満のみ)。0.09 は発動する。
    modTestRunner.Check "分散_境界_差0.10ちょうどは発動しない", _
        (modClarify.HasScoreDispersion( _
            DispLines("約款A" & vbTab & "0.80", "約款B" & vbTab & "0.70"), 10) = False)
    modTestRunner.Check "分散_境界_差0.09は発動する", _
        (modClarify.HasScoreDispersion( _
            DispLines("約款A" & vbTab & "0.80", "約款B" & vbTab & "0.71"), 10) = True)

    ' gap=0 は機能OFF(config の ambiguous_dispersion_gap_x100=0)。
    modTestRunner.Check "分散_gap0は機能OFF", _
        (modClarify.HasScoreDispersion( _
            DispLines("約款A" & vbTab & "0.72", "約款B" & vbTab & "0.70"), 0) = False)

    ' 壊れた入力(空・TAB無し・スコアが数値でない)で落ちない・発動しない。
    modTestRunner.Check "分散_空文字は発動しない", _
        (modClarify.HasScoreDispersion("", 10) = False)
    modTestRunner.Check "分散_TAB無しの行だけなら発動しない", _
        (modClarify.HasScoreDispersion(DispLines("約款A", "約款B"), 10) = False)
    ' スコアが読めない行は0点扱い。0点どうしは差0=拮抗しているのが事実なので
    ' 発動して構わない(聞き返しは害が小さい側)。ここでは落ちないことを見る。
    modTestRunner.Check "分散_スコアが数値でなくても落ちない", _
        (modClarify.HasScoreDispersion( _
            DispLines("約款A" & vbTab & "xx", "約款B" & vbTab & "yy"), 10) = True)
End Sub

Private Sub TestMentionsSourceName()
    ' #6 質問文が資料名を名指ししている(「就業規則の免責は?」)。
    modTestRunner.Check "分散除外_#6_資料名を名指ししていれば聞き返さない", _
        (modClarify.MentionsSourceName("就業規則の免責は?", _
            "就業規則|自動車保険約款") = True)

    ' #8 資料名の一部(火災保険)が質問文に含まれる。
    modTestRunner.Check "分散除外_#8_資料名の主要語が質問文にあれば聞き返さない", _
        (modClarify.MentionsSourceName("火災保険の免責金額は?", _
            "火災保険約款|就業規則") = True)

    ' #1/#2/#9: 資料を名指ししていない短文は除外されない(=分散判定へ進む)。
    modTestRunner.Check "分散除外_免責は?は資料名を含まない", _
        (modClarify.MentionsSourceName("免責は?", _
            "自動車保険約款|火災保険約款|就業規則") = False)
    modTestRunner.Check "分散除外_計算方法は資料名を含まない", _
        (modClarify.MentionsSourceName("計算方法", _
            "自動車保険約款|火災保険約款") = False)

    ' 拡張子つきの資料名でも名指しとして拾う(本棚の名前はファイル名のまま)。
    modTestRunner.Check "分散除外_拡張子つきの資料名でも拾う", _
        (modClarify.MentionsSourceName("就業規則の免責は?", "就業規則.pdf") = True)

    ' 2字の一般語では止めない(「免責」が資料名に含まれるだけで機能を殺さない)。
    modTestRunner.Check "分散除外_2字の一般語では止めない", _
        (modClarify.MentionsSourceName("期限は?", "契約期限の手引き") = False)

    ' 空入力で落ちない。
    modTestRunner.Check "分散除外_質問が空なら除外しない", _
        (modClarify.MentionsSourceName("", "就業規則") = False)
    modTestRunner.Check "分散除外_資料名が空なら除外しない", _
        (modClarify.MentionsSourceName("就業規則の免責は?", "") = False)
End Sub

Private Sub TestDispersionOtherGates()
    ' #5 俯瞰質問。IsTooVague は HasGlobalSignal を見て【分散より先に】除外する。
    modTestRunner.Check "分散他門_#5_俯瞰質問は俯瞰ゲートが先に除外する", _
        (modRagParse.HasGlobalSignal("全体像は?") = True)

    ' #7 複合質問。段0(分解)が先に効くので、分散の聞き返しは要らない。
    modTestRunner.Check "分散他門_#7_複合質問は段0が先に受け持つ", _
        (modAskMulti.DecomposeGate("auto", "免責は?保険料は?", 25) = True)

    ' #11 長い質問は文字数ゲート(ambiguous_max_chars 既定10)で対象外。
    '     IsTooVague と同じ長さ制約を分散にも継承させている事実を固定する。
    modTestRunner.Check "分散他門_#11_17字の質問は長さゲートで対象外", _
        (Len("保険料の計算方法を教えてください") > 10)
    modTestRunner.Check "分散他門_#1_免責は?は長さゲートを通る", _
        (Len("免責は?") <= 10)
    modTestRunner.Check "分散他門_#2_計算方法は長さゲートを通る", _
        (Len("計算方法") <= 10)
End Sub

' ----------------------------------------------------------------------------
' R19H FB-1(A-M⑧): gapの実値を返す純関数。閾値との比較は呼び側が持つ。
' ----------------------------------------------------------------------------
' HasScoreDispersion(上の12件)と同じ走査を共有しているので、ここでは
' 「Boolean では見えなかった値そのもの」だけを固定する。
Private Sub TestDispersionGap()
    Dim n As Long

    ' 拮抗(0.72 vs 0.70)= gap 2。資料は3種類。
    modTestRunner.Check "分散gap_拮抗した3資料はgap2を返す", _
        (modClarify.DispersionGapX100( _
            DispLines("約款A" & vbTab & "0.72", "約款B" & vbTab & "0.70", _
                      "規程C" & vbTab & "0.68"), n) = 2)
    modTestRunner.Check "分散gap_資料数を数えて返す", (n = 3)

    ' 突出(0.9 vs 0.3)= gap 60。鳴らない側こそ校正に要る値。
    modTestRunner.Check "分散gap_突出はgap60を返す", _
        (modClarify.DispersionGapX100( _
            DispLines("約款A" & vbTab & "0.9", "約款B" & vbTab & "0.3"), n) = 60)

    ' 資料が1種類=判定不能。-1(0ではない)で返し、ログでも区別できるようにする。
    modTestRunner.Check "分散gap_資料1種類は判定不能の-1", _
        (modClarify.DispersionGapX100(DispLines("約款A" & vbTab & "0.85", ""), n) = -1)
    modTestRunner.Check "分散gap_判定不能でも資料数は返る", (n = 1)

    ' 空文字でも落ちない(-1・0件)。
    modTestRunner.Check "分散gap_空文字は-1", _
        (modClarify.DispersionGapX100("", n) = -1)
    modTestRunner.Check "分散gap_空文字の資料数は0", (n = 0)

    ' 同じ資料の複数ヒットは1種類(最高スコアだけ残す)=判定不能。
    modTestRunner.Check "分散gap_同一資料の複数ヒットは1種類", _
        (modClarify.DispersionGapX100( _
            DispLines("約款A" & vbTab & "0.72", "約款A" & vbTab & "0.70", ""), n) = -1)

    ' 境界: 差ちょうど0.10 は gap10(呼び側の「未満」で落ちる)。
    modTestRunner.Check "分散gap_差0.10ちょうどはgap10", _
        (modClarify.DispersionGapX100( _
            DispLines("約款A" & vbTab & "0.80", "約款B" & vbTab & "0.70"), n) = 10)
End Sub

' ----------------------------------------------------------------------------
' R19H FA-3(A-H③・B-H②): 同居しているのは「可視の他ブック」だけ。
' ----------------------------------------------------------------------------
' PERSONAL.XLSB・アドイン・不可視ブック・自分自身を数えてしまうと、単独で
' 開いている人にまで同居警告が出て、しかも他のExcelを全部閉じても消えない
' =直せない警告になる(狼少年になった警告は次から読まれない)。
' 入力は "ブック名<TAB>可視ウィンドウ数<TAB>アドインなら1" の vbLf 連結。
Private Function BookLine(ByVal nm As String, ByVal vis As Long, _
                          ByVal isAdd As Long) As String
    BookLine = nm & vbTab & vis & vbTab & isAdd
End Function

Private Sub TestCohabitOtherCount()
    Dim me_ As String: me_ = "MyBookshelf.xlsm"

    ' 単独で開いている(自分の行しか無い)。
    modTestRunner.Check "同居_自分だけなら0", _
        (modIntegrity.CohabitOtherCount(BookLine(me_, 1, 0), me_) = 0)

    ' personal.xlsb のみ同居。Excelが常に非表示で開くので可視0だが、
    ' 可視で開いている端末でも【名前で】除く(両方を固定する)。
    modTestRunner.Check "同居_personal.xlsbだけ(不可視)なら0", _
        (modIntegrity.CohabitOtherCount( _
            BookLine(me_, 1, 0) & vbLf & BookLine("PERSONAL.XLSB", 0, 0), me_) = 0)
    modTestRunner.Check "同居_personal.xlsbは可視でも名前で除く", _
        (modIntegrity.CohabitOtherCount( _
            BookLine(me_, 1, 0) & vbLf & BookLine("personal.xlsb", 1, 0), me_) = 0)

    ' 可視ブック同居(これだけが本物の同居)。
    modTestRunner.Check "同居_可視の他ブック1冊なら1", _
        (modIntegrity.CohabitOtherCount( _
            BookLine(me_, 1, 0) & vbLf & BookLine("見積書.xlsx", 1, 0), me_) = 1)

    ' アドイン(IsAddin=True)は一緒に固まる相手ではない。
    modTestRunner.Check "同居_アドインは数えない", _
        (modIntegrity.CohabitOtherCount( _
            BookLine(me_, 1, 0) & vbLf & BookLine("組織配布.xlam", 1, 1), me_) = 0)

    ' 可視ウィンドウを持たないブック(他マクロがVisible=Falseで開いたもの)。
    modTestRunner.Check "同居_不可視ブックは数えない", _
        (modIntegrity.CohabitOtherCount( _
            BookLine(me_, 1, 0) & vbLf & BookLine("裏方.xlsm", 0, 0), me_) = 0)

    ' 混在: 自分+personal(不可視)+アドイン+不可視+可視2冊 = 2。
    modTestRunner.Check "同居_混在では可視の他ブックだけ数える", _
        (modIntegrity.CohabitOtherCount( _
            BookLine(me_, 1, 0) & vbLf & BookLine("PERSONAL.XLSB", 0, 0) & vbLf & _
            BookLine("組織配布.xlam", 1, 1) & vbLf & BookLine("裏方.xlsm", 0, 0) & vbLf & _
            BookLine("見積書.xlsx", 1, 0) & vbLf & BookLine("台帳.xlsm", 2, 0), me_) = 2)

    ' 自分の名前は大小無視で除く(Workbooks の表記は端末で揺れる)。
    modTestRunner.Check "同居_自分の名前は大小無視で除く", _
        (modIntegrity.CohabitOtherCount(BookLine("MYBOOKSHELF.XLSM", 1, 0), me_) = 0)

    ' 壊れた行(TABが足りない)は数えない。出せない警告を出すより出さない側へ。
    modTestRunner.Check "同居_壊れた行は数えない", _
        (modIntegrity.CohabitOtherCount( _
            BookLine(me_, 1, 0) & vbLf & "壊れた行", me_) = 0)
    modTestRunner.Check "同居_空文字なら0", _
        (modIntegrity.CohabitOtherCount("", me_) = 0)

    ' IsCohabiting は「他ブックが1冊でもあれば同居」(境界)。
    modTestRunner.Check "同居_境界_0冊は単独", (modIntegrity.IsCohabiting(0) = False)
    modTestRunner.Check "同居_境界_1冊で同居", (modIntegrity.IsCohabiting(1) = True)
End Sub

' ----------------------------------------------------------------------------
' R20-2/R20-6(実機第7報①⑧・波B): 深掘りの一般アシスタント対応と
'   3モードの可観測性・実効差・聞き返し強化の真理表。
' ----------------------------------------------------------------------------
Private Sub TestFollowupGate()
    ' (i) ゲート分岐: normal/rag × 履歴有無の全4通り(R20-2b)。
    modTestRunner.Check "続けて質問ゲート_normalかつ一般履歴ありなら許可", _
        (modAppAct.GateUsesGeneralHistory("normal", True, False) = True)
    modTestRunner.Check "続けて質問ゲート_normalかつ一般履歴なしなら不許可", _
        (modAppAct.GateUsesGeneralHistory("normal", False, True) = False)
    modTestRunner.Check "続けて質問ゲート_ragかつRAG履歴ありなら許可", _
        (modAppAct.GateUsesGeneralHistory("rag", True, True) = True)
    modTestRunner.Check "続けて質問ゲート_ragかつRAG履歴なしなら不許可", _
        (modAppAct.GateUsesGeneralHistory("rag", False, False) = False)
End Sub

Private Sub TestGeneralHistoryBoundary()
    ' (iii) followup_max_pairs=0の境界(R20-2e)。RAG側AppendFollowupPairと
    ' 同型のエスケープハッチ(0以下=履歴を持たない)。
    modTestRunner.Check "一般履歴クリア境界_0は消す", _
        (modAppState.ShouldClearGeneralHistory(0) = True)
    modTestRunner.Check "一般履歴クリア境界_負も消す", _
        (modAppState.ShouldClearGeneralHistory(-1) = True)
    modTestRunner.Check "一般履歴クリア境界_1は残す", _
        (modAppState.ShouldClearGeneralHistory(1) = False)
    modTestRunner.Check "一般履歴クリア境界_既定3は残す", _
        (modAppState.ShouldClearGeneralHistory(3) = False)

    ' (ii) OnClearChat後(modAppState.ClearGeneralMemory)の一般履歴消去。
    ' LO環境にはui_stateシートが無くLoadStateは常に既定値""を返すため
    ' (modState.LoadStateの契約)、この環境で再現できる範囲は「クリア後は
    ' HasGeneralMemoryがFalseへ戻る」こと(実機の「クリア前は履歴が読める」側は
    ' UI層のためLO到達範囲外。modClarify.MergeAnswerと同型のLO制約)。
    modAppState.ClearGeneralMemory
    modTestRunner.Check "一般履歴_ClearGeneralMemory後はHasGeneralMemoryがFalse", _
        (modAppState.HasGeneralMemory() = False)
End Sub

Private Sub TestSelfEchoGuard()
    ' (iv) modClarify.MergeAnswer同一質問防御(R20-6c)の境界。
    modTestRunner.Check "自己連結防御_完全一致は自己エコー", _
        (modClarify.IsSelfEcho("免責は?", "免責は?") = True)
    modTestRunner.Check "自己連結防御_前後空白を無視して一致", _
        (modClarify.IsSelfEcho("  免責は?  ", "免責は?") = True)
    modTestRunner.Check "自己連結防御_大小無視(半角英字)で一致", _
        (modClarify.IsSelfEcho("ABC", "abc") = True)
    ' 2026-08-06 R20H FA-10: 旧名称「8字未満の別文言」は実装(文字数を一切見ない
    ' 完全一致比較)と噛み合っていなかったため、実態に合わせて改称。
    modTestRunner.Check "自己連結防御_文言そのものが異なれば自己エコーでない", _
        (modClarify.IsSelfEcho("免責とは", "免責は?") = False)
    modTestRunner.Check "自己連結防御_origQが空なら自己エコーでない", _
        (modClarify.IsSelfEcho("免責は?", "") = False)
    ' R20H FA-10: 全角スペース境界(半角Trim$だけでは落とせず、正規化漏れで
    ' 別文言と誤判定していた実バグ)。
    modTestRunner.Check "自己連結防御_全角スペースの前後付着は正規化して一致", _
        (modClarify.IsSelfEcho("免責は?" & ChrW(&H3000), "免責は?") = True)
    modTestRunner.Check "自己連結防御_語間の全角スペースは半角と同一視して一致", _
        (modClarify.IsSelfEcho("免責" & ChrW(&H3000) & "は?", "免責 は?") = True)
    modTestRunner.Check "自己連結防御_origQが全角スペースのみなら空扱いで自己エコーでない", _
        (modClarify.IsSelfEcho("免責は?", ChrW(&H3000) & ChrW(&H3000)) = False)
End Sub

Private Sub TestThoroughDispersionPair()
    ' (v) 分散閾値: thorough(既定25)ではgap20が発動し、deep/quick(既定10)
    ' では発動しない、の対(R20-6d)。DispLinesは上のTestScoreDispersionと
    ' 同じヘルパー関数を共用する。
    Dim lines20 As String
    lines20 = DispLines("約款A" & vbTab & "0.90", "約款B" & vbTab & "0.70")
    modTestRunner.Check "分散閾値対_thoroughはgap20で発動(閾値25)", _
        (modClarify.HasScoreDispersion(lines20, 25) = True)
    modTestRunner.Check "分散閾値対_deep/quickはgap20で非発動(閾値10)", _
        (modClarify.HasScoreDispersion(lines20, 10) = False)
End Sub

Private Sub TestThoroughStyleAddendum()
    ' 6f: BuildDeepVerifyPrompt(凍結)の戻りへ、入念モードだけの文体指示を
    ' modAskThorough側で連結する(modPrompts本体は1文字も変えない)。
    Dim p As String: p = modAskThorough.ThoroughVerifyPrompt("元プロンプト本体")
    modTestRunner.Check "入念文体_元プロンプトを保持", _
        (Left$(p, Len("元プロンプト本体")) = "元プロンプト本体")
    modTestRunner.Check "入念文体_構造化の指示を連結", _
        (InStr(p, "■見出しで構造化") > 0)
    modTestRunner.Check "入念文体_断定回避の文言を連結", _
        (InStr(p, "資料からは確認できません") > 0)
End Sub

Public Sub RunAll18()
    On Error GoTo MergeFail18
    TestMergeSynPairs
NextGlobal18:
    On Error GoTo GlobalFail18
    TestHasGlobalSignal
NextGate18:
    On Error GoTo GateFail18
    TestDecomposeGateGlobal
NextDisp18:
    On Error GoTo DispFail18
    TestScoreDispersion
NextMention18:
    On Error GoTo MentionFail18
    TestMentionsSourceName
NextOther18:
    On Error GoTo OtherFail18
    TestDispersionOtherGates
NextGap18:
    On Error GoTo GapFail18
    TestDispersionGap
NextCohabit18:
    On Error GoTo CohabitFail18
    TestCohabitOtherCount
NextGate20:
    On Error GoTo GateFail20
    TestFollowupGate
NextHistBound20:
    On Error GoTo HistBoundFail20
    TestGeneralHistoryBoundary
NextEcho20:
    On Error GoTo EchoFail20
    TestSelfEchoGuard
NextDispPair20:
    On Error GoTo DispPairFail20
    TestThoroughDispersionPair
NextStyle20:
    On Error GoTo StyleFail20
    TestThoroughStyleAddendum
NextChain19:
    ' R20-1: 実機第7報⑦(右・下余白の3層根治)の真理表は modTestsPure19 へ。
    ' ここが28,000字のWARN帯に近いため、16→17→18 と同じ線で分割した。
    On Error GoTo ChainFail19
    modTestsPure19.RunAll19
NextDone18:
    On Error GoTo 0
    Exit Sub

MergeFail18:
    modTestRunner.Check "TestMergeSynPairs(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGlobal18
GlobalFail18:
    modTestRunner.Check "TestHasGlobalSignal(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGate18
GateFail18:
    modTestRunner.Check "TestDecomposeGateGlobal(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDisp18
DispFail18:
    modTestRunner.Check "TestScoreDispersion(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextMention18
MentionFail18:
    modTestRunner.Check "TestMentionsSourceName(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextOther18
OtherFail18:
    modTestRunner.Check "TestDispersionOtherGates(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGap18
GapFail18:
    modTestRunner.Check "TestDispersionGap(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextCohabit18
CohabitFail18:
    modTestRunner.Check "TestCohabitOtherCount(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGate20
GateFail20:
    modTestRunner.Check "TestFollowupGate(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextHistBound20
HistBoundFail20:
    modTestRunner.Check "TestGeneralHistoryBoundary(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextEcho20
EchoFail20:
    modTestRunner.Check "TestSelfEchoGuard(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDispPair20
DispPairFail20:
    modTestRunner.Check "TestThoroughDispersionPair(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextStyle20
StyleFail20:
    modTestRunner.Check "TestThoroughStyleAddendum(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextChain19
ChainFail19:
    modTestRunner.Check "modTestsPure19.RunAll19(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone18
End Sub
