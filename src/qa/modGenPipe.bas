Attribute VB_Name = "modGenPipe"
Option Explicit

' ============================================================================
' modGenPipe - 一般アシスタント(本棚を使わない回答)の3段化(2026-08-11 R26-1)
' ----------------------------------------------------------------------------
' なぜ新しいモジュールなのか:
'   ヘッダーの速さトグル(すぐ聞く/しっかり/入念)は、これまで社内ナレッジ検索
'   側だけに効き、一般アシスタントは【常に】quick 相当の1回呼び出しだった
'   (modApp が sendSpeed を捨てて AskGeneral(q, "") を呼んでいた)。
'   利用者から見ると「入念にしても一般アシスタントだけ何も変わらない」で、
'   トグルの意味がモードによって消える。R26-1 はここを埋める。
'
'   置き場をここにした理由は2つ:
'     ・modAppState(UI層)は状態の窓口であって生成の道筋を持つ場所ではない。
'     ・多段の道筋(起草→検証→改稿)は modAskThorough と同じ性質の qa 層の
'       仕事で、Excel オブジェクトに一切触れない(実況は modUIMain.SetStage
'       へ委ねる。R1 の UI通知コールバック例外)。
'
' 3段の設計(仕様 spec_20260810_R26 §1):
'   すぐ聞く   … 1回呼び出し。現行と1文字も変えない(quick_effort/verbosity)。
'   しっかり   … 1回呼び出し + 構造化指示。effort=medium / verbosity=high。
'                (Webリサーチ確定: high と medium の差は一般タスクでは薄く、
'                 high が効くのは多段推論。しっかりは medium + プロンプト
'                 構造化が費用対効果の最良点。)
'   入念       … 起草(high) → 検証(別ペルソナ・別呼び出し) → 改稿 を最大2周。
'                verdict:PASS で早期終了(CoVe型)。
'
' 検証を「別の呼び出し・別のペルソナ」にしているのは、自己採点の楽観バイアス
' 対策。同じ呼び出しの中で「点検して直せ」と言うと、モデルは自分の文章を守る
' ために指摘を軽くする(modAskThorough の(4)自己批判で得た知見と同じ)。
'
' 壊さないための約束:
'   ・検証応答が読めない/落ちたときは、起草(または直前の改稿)をそのまま返す。
'     利用者を待たせた末に壊すより、1段ぶん退行して答えを出すほうがよい。
'     痕跡は usage_log(gen_verify_parse_fail)へ必ず残す(無言の失敗禁止)。
'   ・LLM 呼び出しは全て modGateway.CallLLM 経由(mock_llm=TRUE の dev ビルド
'     でも道筋がそのまま動く)。
'   ・直近ターンの記録(mLastMode/mLastLoops)は AskGeneral が毎ターン先頭で
'     ResetTurn を呼んで捨てる。捨てないと「前のターンの検証回数」が次の
'     回答フッターへ漏れる(R20H FA-8 のモード間リークと同型の事故)。
' ============================================================================

Private Const MODE_QUICK As String = "quick"
Private Const MODE_DEEP As String = "deep"
Private Const MODE_THOROUGH As String = "thorough"

' 実行計画(PlanFor の戻り値)。AskGeneral はこの3値だけを見て分岐する。
Public Const PLAN_SINGLE As String = "single"
Public Const PLAN_SINGLE_DEEP As String = "single_deep"
Public Const PLAN_PIPELINE As String = "pipeline"

' ParseVerdict の戻り値。
Public Const VERDICT_PASS As String = "PASS"
Public Const VERDICT_FINDINGS As String = "FINDINGS"
Public Const VERDICT_FAIL As String = "FAIL"

' 査読者が「問題なし」を宣言する機械可読な合図(正規化後の比較用)。
Private Const PASS_TOKEN As String = "verdict:pass"

' 検証→改稿の周回数の絶対上限。config が壊れた値(999等)でも、利用者を
' 何十分も待たせないための最後の歯止め。実証知見でも自己改善は2～4周で
' 頭打ちになるので、4より上を許す意味が無い。
Private Const LOOPS_HARD_MAX As Long = 4

' prevU/prevA の往復区切り(modAppState.TrimPairs / modApp が保存に使う形と同じ)。
Private Const HIST_SEP As String = ";;;"

' 会話ブロックの見出し。プロンプト内の指示文と綴りを揃えるため定数にする。
Private Const HIST_HEAD As String = "## 先行する会話(古い順)"

' --- 直近1ターンの記録(フッター表示と usage_log の材料) --------------------
Private mLastMode As String       ' 正規化済みモード("" = まだ1度も答えていない)
Private mLastLoops As Long        ' 検証を回した回数
Private mLastPass As Boolean      ' verdict:PASS に到達したか
Private mLastParseFail As Long    ' 検証応答を読めなかった回数

' ----------------------------------------------------------------------------
' ResetTurn - 一般アシスタントの1ターンの開始。前ターンの記録を必ず捨てる。
'   modAppState.AskGeneral が【モードに関係なく毎回】先頭で呼ぶ。
' ----------------------------------------------------------------------------
Public Sub ResetTurn(ByVal normalizedMode As String)
    mLastMode = normalizedMode
    mLastLoops = 0
    mLastPass = False
    mLastParseFail = 0
End Sub

Public Function LastLoops() As Long
    LastLoops = mLastLoops
End Function

Public Function LastPass() As Boolean
    LastPass = mLastPass
End Function

' usage_log gen_mode の detail 欄(検証周回数・PASS到達・パース失敗回数)。
Public Function TurnDetail() As String
    TurnDetail = "loops=" & mLastLoops & " pass=" & IIf(mLastPass, 1, 0) & _
                 " parse_fail=" & mLastParseFail
End Function

' ----------------------------------------------------------------------------
' VerifyFooterNote - 回答フッター(R20-6a の既存機構)へ足す「検証n回」。
' ----------------------------------------------------------------------------
' 一般アシスタントの入念【だけ】が返す。他モード・社内ナレッジ検索側では
' 必ず空を返す(RAG側の段数表示 modAskRetrieve.LastStageTotal とは別系統で、
' どちらか一方だけが出る。混ざると「何段だったのか」がどのモードの数字か
' 分からなくなる=R20H FA-8 で塞いだのと同じ種類のリーク)。
Public Function VerifyFooterNote() As String
    If mLastMode <> MODE_THOROUGH Then Exit Function
    If mLastLoops < 1 Then Exit Function
    VerifyFooterNote = " ・ 検証" & mLastLoops & "回"
End Function

' ----------------------------------------------------------------------------
' PlanFor - モード文字列 -> 実行計画。3段の分岐表の単一情報源。
' ----------------------------------------------------------------------------
' 読めない値・空文字は必ず PLAN_SINGLE(=現行動作)へ倒す。ここを「不明なら
' 多段」に倒すと、ui_state が壊れた1回のために数分待たされる。
Public Function PlanFor(ByVal mode As String) As String
    Select Case modMode.Normalize(mode)
        Case MODE_THOROUGH: PlanFor = PLAN_PIPELINE
        Case MODE_DEEP:     PlanFor = PLAN_SINGLE_DEEP
        Case Else:          PlanFor = PLAN_SINGLE
    End Select
End Function

' ----------------------------------------------------------------------------
' ShouldRunVerifyLoop - 検証→改稿を、あと1周まわしてよいか。
'   doneLoops: すでに回した周回数(0起点) / loopsMax: config の上限。
' ----------------------------------------------------------------------------
Public Function ShouldRunVerifyLoop(ByVal doneLoops As Long, ByVal loopsMax As Long) As Boolean
    Dim mx As Long: mx = loopsMax
    If mx < 0 Then mx = 0                              ' 負値=検証しない(エスケープハッチ)
    If mx > LOOPS_HARD_MAX Then mx = LOOPS_HARD_MAX    ' 壊れたconfigでも待たせすぎない
    ShouldRunVerifyLoop = (doneLoops < mx)
End Function

' ----------------------------------------------------------------------------
' ParseVerdict - 検証応答の判定(純関数)。
'   戻り値: VERDICT_PASS / VERDICT_FINDINGS / VERDICT_FAIL
'   findings: 指摘本文(PASS・FAIL のときは空)
' ----------------------------------------------------------------------------
' 「問題なければ1行目に verdict:PASS とだけ書く」という約束は守られない日が
' 必ず来る(全角コロン・前後の空行・**装飾**・末尾の句点)。判定を厳密一致
' だけにすると、その日から入念モードは永久に2周まわす。逆に「PASS を含む」に
' すると「verdict:PASS ではない」で誤って早期終了する。そこで
' 【1行目を正規化して1語と厳密一致】という中間を取る。
Public Function ParseVerdict(ByVal resp As String, ByRef findings As String) As String
    findings = ""

    Dim s As String
    s = TrimWs(resp)
    If LenB(s) = 0 Then
        ParseVerdict = VERDICT_FAIL      ' 空応答=形式崩れ(起草をそのまま返す側へ)
        Exit Function
    End If
    If modRagParse.IsErrorResponse(s) Then
        ParseVerdict = VERDICT_FAIL      ' "#ERR:…" は査読結果ではない
        Exit Function
    End If

    If IsPassLine(NormalizeVerdictLine(FirstLineOf(s))) Then
        ' R34 A1: 1行目がPASSでも2行目以降(trailing)を無言で捨てない。
        ' 101字以上の懸念文はFINDINGSへ昇格し、それ未満(空・短い装飾)は
        ' 従来通りPASSだが、trailingが非空ならfindingsへ載せて可観測にする。
        Dim trailing As String
        trailing = TrimWs(TrailingOf(s))
        If Len(trailing) > 100 Then
            ' R34 F3: findings は【指摘そのもの】だけにする。ここで verdict 行
            ' 込みの全文(s)を返していたため、改稿プロンプトの「指摘事項:」に
            ' "verdict:PASS" が先頭で刺さり、書き直す側が「問題なしと言われた」
            ' と読める材料を渡していた。両方の経路で trailing に揃える。
            findings = trailing
            ParseVerdict = VERDICT_FINDINGS
            Exit Function
        End If
        findings = trailing
        ParseVerdict = VERDICT_PASS
        Exit Function
    End If

    findings = s
    ParseVerdict = VERDICT_FINDINGS
End Function

' ----------------------------------------------------------------------------
' IsPassLine - 正規化済みの1行を「問題なし宣言」と読んでよいか(純関数)。
' ----------------------------------------------------------------------------
' R26H F7(m-3): 厳密一致だけだと「verdict:PASSです」「verdict:PASSでした」の
' ように丁寧語が1語付いた日から、入念モードは永久に上限まで回る(利用者からは
' 「ただ遅い」としか見えない)。そこで PASS_TOKEN で始まり、残りが肯定的な
' 装飾(です・でした・だ・である・ね・よ。句読点と記号は正規化で既に落ちている)
' 【だけ】のときも PASS と読む。
' ただし「verdict:PASSではない」を PASS と読むのは、査読の指摘を丸ごと捨てる
' 最悪の誤読なので、残りに否定語が1つでも混ざったら PASS にしない。否定判定を
' 先に置くのは、"で"+"はない" のように装飾の剥がし方によっては否定形が
' 装飾の連なりに見えうるため(短絡しない And を避け、判定を段で分ける)。
Private Function IsPassLine(ByVal normLine As String) As Boolean
    If normLine = PASS_TOKEN Then
        IsPassLine = True
        Exit Function
    End If
    If Left$(normLine, Len(PASS_TOKEN)) <> PASS_TOKEN Then Exit Function

    Dim rest As String
    rest = Mid$(normLine, Len(PASS_TOKEN) + 1)
    If HasNegation(rest) Then Exit Function
    IsPassLine = (LenB(StripAffirmative(rest)) = 0)
End Function

' 否定語(1つでもあれば PASS と読まない)。「ない」「ぬ」「不可」「否」「not」。
Private Function HasNegation(ByVal s As String) As Boolean
    Dim ng As Variant
    ng = Array("ない", "無い", "なし", "ません", "ぬ", "否", "不可", "違", "not", "no")
    Dim i As Long
    For i = LBound(ng) To UBound(ng)
        If InStr(1, s, CStr(ng(i)), vbTextCompare) > 0 Then
            HasNegation = True
            Exit Function
        End If
    Next i
End Function

' 肯定的な装飾を頭から剥がす。剥がしきれた(空になった)ら装飾だけだった。
' 各周で必ず短くなるので停止する。
Private Function StripAffirmative(ByVal s As String) As String
    Dim ok As Variant
    ok = Array("でした", "である", "です", "ます", "だ", "ね", "よ", "、", "。")
    Dim t As String: t = s
    Dim hit As Boolean
    Dim i As Long
    Do
        hit = False
        For i = LBound(ok) To UBound(ok)
            Dim w As String: w = CStr(ok(i))
            If Left$(t, Len(w)) = w Then
                t = Mid$(t, Len(w) + 1)
                hit = True
                Exit For
            End If
        Next i
    Loop While hit And LenB(t) > 0
    StripAffirmative = t
End Function

' 先頭行(改行はvbLfへ寄せてある前提)。
Private Function FirstLineOf(ByVal s As String) As String
    Dim p As Long: p = InStr(s, vbLf)
    If p > 0 Then
        FirstLineOf = Left$(s, p - 1)
    Else
        FirstLineOf = s
    End If
End Function

' 2行目以降(先頭行を除いた残り。先頭行しか無ければ空文字)。R34 A1。
Private Function TrailingOf(ByVal s As String) As String
    Dim p As Long: p = InStr(s, vbLf)
    If p > 0 Then
        TrailingOf = Mid$(s, p + 1)
    Else
        TrailingOf = ""
    End If
End Function

' 前後の空白・全角空白・改行・タブを落とす(Trim$ は空白しか落とさない)。
Private Function TrimWs(ByVal s As String) As String
    Dim t As String
    t = Replace(Replace(s, vbCr, vbLf), vbTab, " ")
    Dim c As String
    Do While LenB(t) > 0
        c = Left$(t, 1)
        If c <> " " And c <> vbLf And c <> ChrW(&H3000) Then Exit Do
        t = Mid$(t, 2)
    Loop
    Do While LenB(t) > 0
        c = Right$(t, 1)
        If c <> " " And c <> vbLf And c <> ChrW(&H3000) Then Exit Do
        t = Left$(t, Len(t) - 1)
    Loop
    TrimWs = t
End Function

' 1行を比較用の1語へ畳む(小文字化・全角コロン・空白・装飾記号・末尾句読点)。
Private Function NormalizeVerdictLine(ByVal ln As String) As String
    Dim t As String
    t = LCase$(ln)
    t = Replace(t, ChrW(&HFF1A), ":")     ' 全角コロン
    t = Replace(t, ChrW(&H3000), "")      ' 全角空白
    t = Replace(t, " ", "")
    t = Replace(t, vbTab, "")

    Dim c As String
    Do While LenB(t) > 0                  ' 行頭の "**" "- " "#" 等の装飾
        c = Left$(t, 1)
        If InStr("*-#>", c) = 0 Then Exit Do
        t = Mid$(t, 2)
    Loop
    Do While LenB(t) > 0                  ' 行末の装飾・句読点・括弧閉じ
        c = Right$(t, 1)
        If InStr("*。.、,!!?？:：)）」』】", c) = 0 Then Exit Do
        t = Left$(t, Len(t) - 1)
    Loop

    NormalizeVerdictLine = t
End Function

' ============================================================================
' 会話履歴のプロンプト整形(R29 W1-1)
' ============================================================================

' ----------------------------------------------------------------------------
' GenHistoryBlock - 会話履歴(prevU/prevA)を読める塊へ整形する(純関数)。
' ----------------------------------------------------------------------------
' modAsk.HistoryBlock(modAsk.bas:723-735)と【同型】。あちらは modAsk が1本の
' 文字列として持つ履歴を畳む専用で、こちらは一般アシスタントの prevU/prevA
' (質問側・回答側が別々の ";;;" 区切り文字列)を受ける。modAsk は凍結モジュール
' のため参照のみとし、ここへコピー実装する(共通化は凍結解除を要するので採らない)。
'
' 保存形式は【新しい順】。ここで【古い→新しい】へ並べ直すのは、読む側のモデルが
' 「直前の話題」を取り違えないため(新しい順のまま渡すと、一番上=最新を会話の
' 始まりだと読む事故が起きる)。
'
' prevA 側の要素数が prevU 側と食い違う場合は、件数を prevU 側で決め、対応が
' 取れない位置は空のまま出す。R29H F3で保存側(modAppState.AskGeneral)が
' 保存直前に本文の ";;;" を半角スペースへ退避するようになったため、新規保存分で
' 「回答本文に ";;;" が混ざって往復がずれる」事態は起きない想定。ここでの保険は
' 置換前に書かれた旧state等を読んだ場合の防御として残す。プロンプト本文を
' 作るだけの関数なので、落ちる・件数がずれるより「壊れずに出す」を採る。
'
' クランプはしない(保存側が既に1往復あたり質問3,000字級・回答2,000字で
' 切っている。二重に切ると、どちらが効いたのか追えなくなる)。
' ----------------------------------------------------------------------------
Public Function GenHistoryBlock(ByVal prevU As String, ByVal prevA As String) As String
    If LenB(prevU) = 0 Then Exit Function

    Dim qArr() As String
    Dim aArr() As String
    qArr = Split(prevU, HIST_SEP)
    aArr = Split(prevA, HIST_SEP)

    Dim out As String
    Dim i As Long, n As Long
    For i = UBound(qArr) To LBound(qArr) Step -1
        n = n + 1
        out = out & "【会話" & n & "】" & vbLf & _
              "Q: " & qArr(i) & vbLf & _
              "A: " & PartAt(aArr, i) & vbLf & vbLf
    Next i
    GenHistoryBlock = out
End Function

' 配列の i 番目(範囲外なら空文字)。上の食い違いに対する保険。境界判定と参照を
' 分けているのは、VBA の And が短絡しないため(1本の If に畳むと添字エラー)。
Private Function PartAt(ByRef arr() As String, ByVal i As Long) As String
    If i < LBound(arr) Then Exit Function
    If i > UBound(arr) Then Exit Function
    PartAt = arr(i)
End Function

' ----------------------------------------------------------------------------
' HistorySection - 会話ブロックへ見出しを付けた「プロンプトへ差し込む塊」。
'   履歴が空なら空文字を返す=差し込んでも R28 以前と1文字も変わらない。
' ----------------------------------------------------------------------------
Private Function HistorySection(ByVal prevU As String, ByVal prevA As String) As String
    Dim blk As String
    blk = GenHistoryBlock(prevU, prevA)
    If LenB(blk) = 0 Then Exit Function
    HistorySection = HIST_HEAD & vbLf & blk
End Function

' ============================================================================
' プロンプト(日本語。modAppState が組み立てる共通の口調ルールへ継ぎ足す形)
' ============================================================================

' すぐ聞く の起草に足す指示(R28 W3-1a)。しっかり/入念は DeepRules で共通
' ルールを上書きしているが、すぐ聞くはこれまで共通ルールのみで専用ルールが
' 無かった(3モード中ここだけ非対称)。速さを保ったまま「結論が最初の1文で
' 出る」を徹底するための追加指示。
Public Function QuickRules() As String
    QuickRules = _
        "・結論を最初の1文で言い切る。前置きや質問の復唱はしない。" & vbLf & _
        "・全体は200字以内。" & vbLf & _
        "・箇条書きを使う場合は3点まで。"
End Function

' ----------------------------------------------------------------------------
' CarryRules - 社内ナレッジから引き継いだ文脈の扱い方(2026-09-10 R45)。純関数。
' ----------------------------------------------------------------------------
' なぜ要るか(R44 の解剖で敵対的検証を生き残った所見):
'   一般モード(壁打ち)の指示文には事実面のガードが1本も入らない。そこへ
'   モード切替の橋渡し(modConvBridge)が【直前の社内ナレッジ検索での文脈】と
'   して社内資料由来の回答本文を流し込む。しかも運ばれるのは modAsk が
'   控えた本文(mLastCleanAnswer)で、R44 の常設ガード「※…」も低関連度の⚠も
'   その【後】に足す表示専用の飾りなので含まれない。つまり画面には
'   「社内資料は未使用」と出ているのに、中身は社内回答を引き継いでいる。
'
' 文言の要件(反証班の指摘をそのまま守る):
'   ・「社内資料は一切使えない」とは言わない。conv_bridge は既定 TRUE なので、
'     そう書くと直前に出典付きで確定した数字まで撤回させる
'     (R29 W1-2 で潰した文脈断線と同型の事故になる)。
'   ・「会話に出ている情報を根拠として使うこと自体は問題ない」を明示する。
'   ・止めるのは【断定して言い直すこと】だけに絞る。
' 呼ぶのは modAppState が【橋渡しヘッダーがあるターンだけ】(無条件に入れると
' 存在しない会話を指す指示が残り、モデルがその会話について語り出す。
' 同じ型の実害が本モジュール:409-415 に記録されている)。
Public Function CarryRules() As String
    CarryRules = _
        "・直前の会話に「【直前の社内ナレッジ検索での文脈】」が含まれる場合、" & _
        "そこに出てくる会社固有の条件・金額・期限・条文番号・様式は、" & vbLf & _
        "  自分の知識で断定し直さず「社内ナレッジ検索で確認した内容です。" & _
        "原文は出典でお確かめください」と一言を添える。" & vbLf & _
        "・会話に出ている情報を根拠として使うこと自体は問題ない。"
End Function

' しっかり聞く/入念の起草に共通する「深く・構造立てて」の指示。
' 一般アシスタントの共通ルールは「全体はおおむね200～400字」と言っているので、
' ここで明示的に上書きしないと、深掘りを指示しても字数の壁で潰れる。
Public Function DeepRules() As String
    DeepRules = _
        "・まず質問の前提と論点を1～2行で整理してから、構造立てて答える。" & vbLf & _
        "・表面的な一般論で止まらず、判断基準・具体例・注意点まで踏み込む。" & vbLf & _
        "・最後に「次に考えるべき問い」を1つ添える。" & vbLf & _
        "・字数は上の200～400字ではなく600～900字程度を目安にする(この指示が優先)。"
End Function

' 入念の起草だけに足す指示(自分から反証材料を出させる)。R28 W3-1b: しっかり
' (deep、600～900字目安)との差別化を明確にするため、構造(①②③④)と字数目安を
' 明示指定する。指定が無いと DeepRules の600～900字がそのまま効いてしまい、
' 「入念」を選ぶ意味がしっかりと区別できない。
Private Function ThoroughDraftRules() As String
    ThoroughDraftRules = _
        "・反対説・例外・限界条件も自ら挙げる(都合の良い前提だけで書かない)。" & vbLf & _
        "・構成は①前提→②本論→③実務上の注意→④次の一手、の順で書く。" & vbLf & _
        "・字数は800～1200字程度を目安にする(しっかり聞くの600～900字より深く書く)。"
End Function

' 入念の検証(査読者ペルソナ)。起草者とは別の呼び出しで読ませる。
' R29 W1-2/W1-3: 会話ブロックと「文脈を根拠不明扱いしない」指示を足す。
' 追加分は履歴があるときだけ足す(履歴が空なら rule も hist も空文字=R28 と
' 1文字も変わらないプロンプトになる)。存在しない会話を指す指示だけが残ると、
' 査読者が「上記の会話が見当たらない」を指摘として書き始める。
Private Function VerifyPrompt(ByVal q As String, ByVal answerText As String, _
                              ByVal prevU As String, ByVal prevA As String) As String
    Dim hist As String
    hist = HistorySection(prevU, prevA)
    Dim rule As String
    If LenB(hist) > 0 Then
        rule = "・利用者と起草者の間には下記「先行する会話」があり、その内容が" & _
               "回答案に反映されている。会話に現れる情報を「根拠不明」" & _
               "「前提が示されていない」として指摘しないこと。" & vbLf
    End If

    VerifyPrompt = _
        "あなたは起草者とは別の査読者です。以下の質問と回答案を読み、" & _
        "回答案の主な主張について検証質問を3～5個つくり、それぞれに自分で答え、" & _
        "誤り・根拠不足・見落としを具体的に列挙してください。" & vbLf & _
        "・列挙は「・」で始まる1行1件。どの記述についての指摘かが分かるように書く。" & vbLf & _
        "・言い回しの好みや体裁の話は書かない(内容の誤りだけを書く)。" & vbLf & _
        "・問題が無ければ、1行目に verdict:PASS とだけ書いて終わる" & _
        "(この場合は他に何も書かない)。" & vbLf & rule & vbLf & _
        hist & _
        "## 質問" & vbLf & q & vbLf & vbLf & _
        "## 回答案" & vbLf & answerText
End Function

' 入念の改稿(指摘を反映して書き直させる)。口調ルールは起草と同じものを渡す。
' R29 W1-2/W1-4: 会話ブロックと「会話の情報は根拠として維持してよい」指示。
' 査読が文脈由来の記述を落としても、改稿がここで踏みとどまれるようにする。
Private Function RevisePrompt(ByVal sysBase As String, ByVal q As String, _
                              ByVal answerText As String, ByVal findings As String, _
                              ByVal prevU As String, ByVal prevA As String) As String
    Dim hist As String
    hist = HistorySection(prevU, prevA)
    Dim rule As String
    If LenB(hist) > 0 Then
        rule = "・下記「先行する会話」に現れる情報は根拠として維持してよい" & _
               "(会話で共有済みの前提を、根拠不明として削らない)。" & vbLf
    End If

    RevisePrompt = sysBase & vbLf & DeepRules() & vbLf & _
        "・以下の査読指摘を反映して回答を書き直す。指摘に反論できる場合は" & _
        "その根拠を1行で示し、直せない点は「確認が必要」と明示する。" & vbLf & _
        "・査読のやり取り自体は書かない(利用者が読むのは書き直した回答だけ)。" & vbLf & _
        rule & vbLf & _
        hist & _
        "## 質問" & vbLf & q & vbLf & vbLf & _
        "## 回答案" & vbLf & answerText & vbLf & vbLf & _
        "## 査読指摘" & vbLf & findings
End Function

' ============================================================================
' RunThorough - 入念に聞く の本体(起草 -> 検証 -> 改稿・最大2周)
' ----------------------------------------------------------------------------
'   q       : 質問文
'   sysBase : modAppState.AskGeneral が組み立てた共通の口調ルール+追加制約
'   prevU/prevA : 一般アシスタントの会話履歴(;;;区切り・新しい順)
'
'   【R29 裁定変更】R26-1 当初は「起草にだけ渡す。検証・改稿は目の前の回答案
'   だけを見ればよく、履歴を混ぜると査読者が前のターンの話題へ引きずられる」
'   としていた。実機第14報で、この設計が文脈断線の直接原因と確定したため撤回する。
'   起草は履歴を見て「さっきの件」を答えるのに、査読者はその履歴を知らないので
'   文脈由来の記述を「根拠不明」と指摘し、改稿がそれを削って一般論へ戻していた
'   (失敗ターンの usage_log は全て loops=2/pass=0 で機序と整合)。
'   R29 W1-2 以降は【起草・検証・改稿の3プロンプト本文すべて】へ会話ブロックを
'   差し込む。危惧した「前ターンへの引きずり」は、W1-3/W1-4 の指示文
'   (会話は前提であって回答対象ではない)で抑える。
'   なお CallLLM のリボン引数 prevU/prevA は従来どおり起草にだけ渡す(本文へ
'   埋め込むことで、リボン側が ";;;" を往復へ割り戻すかどうかに依存しなくなる)。
'   戻り値  : 最終回答本文、または起草が落ちたときの "#ERR:…"
' ============================================================================
Public Function RunThorough(ByVal q As String, ByVal sysBase As String, _
                            ByVal prevU As String, ByVal prevA As String) As String
    Dim mdl As String: mdl = modConfig.GetString("quick_model", "gpt-5.6-luna")
    Dim eff As String: eff = modConfig.GetString("gen_thorough_effort", "high")
    Dim vrb As String: vrb = modConfig.GetString("gen_deep_verbosity", "high")
    Dim lat As Long

    ' --- (1) 起草 -----------------------------------------------------------
    ShowGenStage "draft", 0
    Dim best As String
    best = modGateway.CallLLM( _
        sysBase & vbLf & DeepRules() & vbLf & ThoroughDraftRules() & _
        vbLf & vbLf & HistorySection(prevU, prevA) & _
        "## 質問" & vbLf & q, _
        "gen_thorough_draft", eff, vrb, mdl, lat, prevU, prevA)
    If modRagParse.IsErrorResponse(best) Then
        ' 起草が無ければ検証も改稿も意味が無い。呼び出し元(AskGeneral)の
        ' 既存の "#ERR:" 経路へそのまま返す(失敗の作法を二重に持たない)。
        RunThorough = best
        Exit Function
    End If

    Dim loopsMax As Long: loopsMax = modConfig.GetLong("gen_thorough_loops_max", 2)
    Dim nLoop As Long: nLoop = 0
    Dim vr As String
    Dim rv As String
    Dim findings As String
    Dim verdictKind As String

    Do While ShouldRunVerifyLoop(nLoop, loopsMax)
        ' --- (2) 検証 -------------------------------------------------------
        ShowGenStage "verify", nLoop + 1
        vr = modGateway.CallLLM(VerifyPrompt(q, best, prevU, prevA), "gen_thorough_verify", _
            modConfig.GetString("gen_thorough_verify_effort", "medium"), _
            modConfig.GetString("gen_thorough_verify_verbosity", "medium"), mdl, lat)
        nLoop = nLoop + 1
        mLastLoops = nLoop

        verdictKind = ParseVerdict(vr, findings)
        If verdictKind = VERDICT_FAIL Then
            ' 査読結果が読めない。ここで止めて【いま手元にある回答】を返す。
            mLastParseFail = mLastParseFail + 1
            On Error Resume Next
            modLog.LogUsage "gen_verify_parse_fail", MODE_THOROUGH, _
                "loop=" & nLoop & " resp=" & modUtil.SafeLeft(vr, 120)
            On Error GoTo 0
            Exit Do
        End If
        If verdictKind = VERDICT_PASS Then
            mLastPass = True
            If LenB(findings) > 0 Then
                ' R34 A1: PASSでもtrailingが残っていた(無言消失にしない可観測化)。
                On Error Resume Next
                modLog.LogUsage "gen_verify_pass_trailing", MODE_THOROUGH, _
                    "loop=" & nLoop & " trailing=" & modUtil.SafeLeft(findings, 120)
                On Error GoTo 0
            End If
            Exit Do
        End If

        ' --- (3) 改稿 -------------------------------------------------------
        ShowGenStage "revise", nLoop
        rv = modGateway.CallLLM(RevisePrompt(sysBase, q, best, findings, prevU, prevA), _
            "gen_thorough_revise", eff, vrb, mdl, lat)
        If modRagParse.IsErrorResponse(rv) Then
            ' 改稿だけが落ちた場合は、査読前の本文で確定する(退行して答えを出す)。
            On Error Resume Next
            modLog.LogUsage "gen_revise_fail", MODE_THOROUGH, _
                "loop=" & nLoop & " resp=" & modUtil.SafeLeft(rv, 120)
            On Error GoTo 0
            Exit Do
        End If
        best = rv
    Loop

    RunThorough = best
End Function

' ----------------------------------------------------------------------------
' ShowGenStage - 段の実況。既存の進捗表示機構(modUIMain.SetStage ->
'   modLive.PaintStage で「考えています…」の仮バブルを差し替える)へ流すだけ。
'   新しいバナー機構は作らない(R1 の UI通知コールバック例外の範囲内)。
' ----------------------------------------------------------------------------
Private Sub ShowGenStage(ByVal kind As String, ByVal roundNo As Long)
    On Error Resume Next
    modUIMain.SetStage GenStageText(kind, roundNo)
    On Error GoTo 0
End Sub

' 実況文(利用者の言葉。「起草」「検証」「改稿」という業界語は出さない)。
' 2周目に入ったことは伝える。同じ文言が黙って2回出ると、止まったように見える。
Private Function GenStageText(ByVal kind As String, ByVal roundNo As Long) As String
    Dim lbl As String
    Select Case kind
        Case "draft":  lbl = "下書きを作成中…"
        Case "verify": lbl = "下書きを検証中…"
        Case Else:     lbl = "指摘を反映して書き直し中…"
    End Select
    If roundNo >= 2 Then lbl = lbl & "(" & roundNo & "周目)"
    GenStageText = lbl
End Function
