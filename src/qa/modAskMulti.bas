Attribute VB_Name = "modAskMulti"
Option Explicit

' ============================================================================
' modAskMulti - 複合質問の分解 → 論点ごとの調査 → 統合(2026-08-05 R16-3A)
' ----------------------------------------------------------------------------
' なぜ要るのか(R16 要望③・仕様 spec_20260805_R16 §3):
'   「AとBの違いと、Cの申請手続きは?」のような質問は、1本のクエリで検索すると
'   【全部の論点に薄く当たった資料】が上位に来る。入念モードは6段かけて丁寧に
'   検証するが、材料そのものが「どの論点も詰め切れていない16件」なので、
'   時間をかけたぶんだけ丁寧に薄い答えが出る。現行の expand(言い換え拡張)は
'   角度を増やすための機能で、論点を割る機能ではない。
'
'   そこで入念モードにだけ段0を1回足す。質問が複数論点なら論点ごとに
'   検索+下書きし、最後に1本へ統合してから自己点検・検証・出典突合へ回す。
'   AI呼び出しの回数は【どこまでを数えるか】で2つの数がある(R16H FB-5。
'   数え方を書かずに数字だけ出すと、docs と実況とログで別の数が並ぶ):
'     ・このモジュールが出す分 = 1(判定) + 論点数 + 3(統合・自己点検・検証)
'       = 3論点で7回。実況の分母(n/7段)はこちら。
'     ・質問全体 = 論点数 + 6回(検索の拡張・再ランク・判定込み。3論点で9回)。
'       docs の「AI呼び出し」はこちら=現行入念6回 + 論点数、が正しい見積もり。
'
' 設計判断:
'   ・不発は必ず「従来どおり」へ落ちる。判定が #ERR・タグ崩れ・single・
'     論点1本のときは False を返し、呼び出し元(modAsk)がそのまま
'     modAskThorough.RunThoroughFlow を実行する。分解は上積みであって
'     代替ではないので、読めない応答で分解へ倒すのは筋が悪い。
'   ・quick / deep は1行も通らない(modAsk の thorough 分岐からしか呼ばない)。
'   ・段0で verdict=clarify(読み方が定まらない)なら、薄い回答を作らずに
'     「どれを調べますか?」を番号選択肢で返す(R16-3B)。LLMの追加呼び出しは
'     0回=段0の1回だけで返るので、待たせずに主導権を利用者へ返せる。
'     選択肢が2件未満しか読めなければ逆質問は成立しないので従来フローへ落とす。
'   ・中断(ESC=Err18)は論点の境界だけで拾う。LLM呼び出しの内側は
'     modGateway が Err18 を握り潰す既知の制約があり、そこは変えられない。
'     拾ったら【完成済みの論点だけで】統合して正常返却する。調べ終わった
'     ぶんを捨てて「中断しました」だけ返すのは、待った時間の全損になる。
'   ・出典突合(modAskThorough.AnnotateAgainstHits)は、全論点のヒットを
'     chunk_id で重複排除して連結した1本に対して1回だけ掛ける。論点ごとに
'     掛けると、論点2の資料を引いた文が論点1の索引に無くて「確認できず」に
'     なる(索引の作り方だけで嘘の警告が出る)。
'   ・Excelオブジェクトには触れない。実況は modUIMain.SetStage を直接呼ぶ
'     (段番号の分母が single 経路と違うため modMode.AskStageIndex は使わない。
'      あちらを触ると single 経路の番号まで動く)。
' ============================================================================

Private Const MODE_THOROUGH As String = "thorough"

' 論点数の天井。config decompose_max_parts をこの範囲へ丸める。
' 上限5は「1質問で7回のLLM呼び出し」を超えないための安全弁。
Private Const PARTS_MIN As Long = 2
Private Const PARTS_MAX As Long = 5

' 分解経路を通ったターンだけ True(検証注記をどちらから取るかの印)。
Private mHandled As Boolean
Private mNote As String

' 直近ターンの論点数。usage_log の ask_steps へ足す "dec=N" の材料で、
' modAsk が読んだら消す(ConsumeStepBuf と同じ作法)。
Private mParts As Long

' このターンの開始時刻(Timer値・秒)。実況へ「経過M分S秒」を出すためだけに
' 持つ(R16H FB-8 / B-M7)。分解した入念は数分ブロックし、その間ウィンドウは
' 「応答なし」に見える。段番号だけでは【進んでいるのか止まっているのか】が
' 分からず、待てる人まで強制終了してしまう。0=未計測(TryDecomposed を通って
' いないターン。深夜0時ちょうどに始まった場合も未計測扱いになるが、実害は
' 経過表示が1ターンぶん出ないことだけ)。
Private mT0 As Double

' ----------------------------------------------------------------------------
' TryDecomposed - 入念モードの入口(modAsk の thorough 分岐から最初に呼ばれる)
' ----------------------------------------------------------------------------
'   True  = この関数が回答を作った。result / ok / hits / nHits を書き換え済みで、
'           呼び出し元は従来どおり Done: の共通後処理へ流すだけでよい
'           (hits には全論点のヒットを重複排除で連結した unionHits が入る。
'            RememberCitedSources も ApplyLowHitWarning も出典チップも、
'            この配列を見れば分解後の実態と一致する)。
'   False = 何もしていない。呼び出し元は従来の RunThoroughFlow を実行する。
'   ok=False で True を返すのは【まだ答えていないターン】で、2種類ある:
'     ・全論点が失敗した → result は "#ERR:…"(呼び出し元の
'       modRagParse.BuildErrorAnswer が利用者向け文言へ直す既存の契約)
'     ・逆質問を出した(R16H FA-2) → result は逆質問の本文。BuildErrorAnswer は
'       #ERR以外を素通しするので、そのまま画面に出る。
'   どちらも ok=False なので、履歴・会話出典・共有・感謝状・節約時間・
'   低関連度警告は全て不発火になる(=答えていないものを答えたことにしない)。
'   会話履歴は modAsk.HistoryBlock() を直接読む(modAskRetrieve が拡張段で
'   やっているのと同じ。呼び出し元の1行を短く保つため)。
' ----------------------------------------------------------------------------
Public Function TryDecomposed(ByVal q As String, ByRef hits() As Hit, ByRef nHits As Long, _
                              ByRef ok As Boolean, ByRef result As String, _
                              ByVal prevU As String, ByVal prevA As String) As Boolean
    mHandled = False
    mNote = ""
    mParts = 0
    mT0 = Timer                ' 経過表示の基準(FB-8)。段0の判定から数え始める

    ' R18-7b(実機第5報⑦)/ R18H FB-2: 発動条件は DecomposeGate 1本に閉じる
    ' (実装とテストが同じ式を見るため。従来はテスト側が同じ式を書き写しており、
    ' 片方だけ直しても誰も気付けなかった=憲章§4-5)。
    If Not DecomposeGate(modConfig.GetString("decompose_mode", "auto"), q, _
                         modConfig.GetLong("decompose_min_chars", 25)) Then Exit Function

    Dim parts() As String
    Dim nParts As Long
    Dim opts() As String
    Dim nOpts As Long
    Dim verdict As String
    verdict = RunDecideStage(q, modAsk.HistoryBlock(), parts, nParts, opts, nOpts)

    ' R16-3B: 読み方が定まらない質問は、薄い回答を作らずに番号で選ばせる。
    ' 選択肢が読めなかった(タグ崩れ・1件だけ)ときは逆質問が成立しないので、
    ' 黙って従来の入念フローへ落ちる=「聞き返しに失敗して何も返らない」を作らない。
    '
    ' 2026-08-05(R16H FA-2 / A-H2): ok=False で返す。逆質問は【まだ答えていない】
    ' ターンであって成功した回答ではない。ok=True で返していた間は、✅(解決した)を
    ' 押すと聞き返しの文面が「解決済みQ&A」として部内へ共有され、資料の作者へ感謝状
    ' まで飛んでいた(modAsk の共有条件は mLastCleanAnswer の有無=ok の有無)。
    ' 併せて会話履歴・会話出典・節約時間・低関連度警告も不発火になる。
    ' TryDecomposed は True のまま(False にすると modAsk が RunThoroughFlow を
    ' 走らせ、聞き返したうえで薄い回答まで作ってしまう)。逆質問の本文は
    ' modRagParse.BuildErrorAnswer の素通し(FA-2)でそのまま画面へ出る。
    ' hits/nHits は呼び出し前の値のまま触らない(出典チップは検索結果の実態)。
    If verdict = "clarify" Then
        If ShouldClarify(modConfig.GetString("clarify_mode", "auto"), nOpts) Then
            result = AskWhichTopic(q, opts, nOpts)
            ok = False
            mHandled = True
            TryDecomposed = True
        End If
        Exit Function
    End If

    If verdict <> "parts" Then Exit Function
    If nParts < PARTS_MIN Then Exit Function
    mParts = nParts        ' ここから先は必ず分解ターン(中断・全滅でも dec= を残す)

    Dim strictG As Boolean: strictG = modConfig.GetBool("strict_grounding", False)
    Dim ansTags As Boolean: ansTags = modConfig.GetBool("answer_tags", False)
    Dim mdl As String: mdl = modConfig.GetString("recommended_model", "gpt-5.5")

    Dim total As Long: total = 1 + nParts + 3
    Dim topKPer As Long
    topKPer = PerPartTopK(modConfig.GetLong("topk_thorough", 16), nParts)

    Dim uHits() As Hit
    Dim uN As Long: uN = 0
    Dim sect() As String
    ReDim sect(1 To nParts)
    Dim okN As Long: okN = 0
    Dim lastErr As String
    Dim aborted As Boolean
    Dim missingAt As Long
    Dim bodyI As String
    Dim i As Long

    On Error GoTo PartFail
    For i = 1 To nParts
        ' 論点の境界に1つだけ置く DoEvents。ESC の Err18 がこの窓から入る
        ' (LLM呼び出しの内側では modGateway が握り潰すので届かない)。
        DoEvents
        Stage nParts, 1 + i, total, "論点" & i & "を調査中…"
        bodyI = ""
        bodyI = OnePart(q, parts(i), i, topKPer, uHits, uN, strictG, ansTags, mdl, lastErr)
        sect(i) = BuildPartSection(i, parts(i), bodyI)
        If LenB(bodyI) > 0 Then okN = okN + 1
    Next i
    GoTo PartsDone

PartFail:
    If Err.Number = 18 Then
        aborted = True
        Err.Clear
        Resume PartsDone
    End If
    ' 想定外の失敗はその論点だけ落として次へ(1論点の事故で全部を失わない)。
    lastErr = "#ERR:E0602:" & Err.Description
    Err.Clear
    Resume Next

PartsDone:
    On Error GoTo 0

    ' 1論点も終わらないうちに中断された = 統合する材料が無い。握り潰さず
    ' 素の中断として呼び出し元へ返す(modAsk が「操作を中断しました」を出す)。
    If okN = 0 And aborted Then Err.Raise 18

    If okN = 0 Then
        ' 全滅。従来のエラー回答へ渡す(文言化は modAsk 側の既存経路)。
        If LenB(lastErr) = 0 Then
            lastErr = "#ERR:E0202:分解した論点のどれでも資料を確認できませんでした"
        End If
        On Error Resume Next
        modLog.LogUsage "multi_all_failed", MODE_THOROUGH, "parts=" & nParts
        On Error GoTo 0
        result = lastErr
        ok = False
        mHandled = True
        TryDecomposed = True
        Exit Function
    End If

    ' 節を並べつつ、中断で節そのものが作れなかった最初の論点を控えておく。
    ' 中断した番号(i)をそのまま使うと、節が出来た直後に中断されたときに1つずれて
    ' 「調べてある論点を未調査と言う」ことになる。本文に出ていない論点だけを言う。
    Dim sections As String
    For i = 1 To nParts
        If LenB(sect(i)) > 0 Then
            If LenB(sections) > 0 Then sections = sections & vbLf & vbLf
            sections = sections & sect(i)
        ElseIf missingAt = 0 Then
            missingAt = i
        End If
    Next i

    ' R16-3C: 精読(根拠チャンクの前後を束ねる)は、全論点を1本にした unionHits へ
    ' 1回だけ掛ける。論点ごとに掛けると同じ資料の同じ近傍を論点数ぶん読み直す。
    ' 増えた件数は uUse にだけ載せ、呼び出し元へ返す nHits(出典チップ・信頼度の
    ' 材料)は検索ヒットのままにする。近傍は根拠であってヒットではない。
    Dim uUse As Long: uUse = uN
    modAskFocus.NeighborExpand uHits, uUse, modConfig.GetLong("deep_neighbor", 2)

    Dim body As String
    body = Integrate(q, sections, uHits, uUse, nParts, total, strictG, ansTags, mdl, prevU, prevA)

    ' 出典突合は全論点ぶんの索引で1回だけ(理由はモジュール冒頭の設計判断)。
    ' 索引は【精読で足した近傍を含んだ】側から作る。近傍を根拠に書いた文へ
    ' 「(出典確認できず)」が付くと、正しい引用を嘘だと言うことになる。
    body = modAskThorough.AnnotateAgainstHits(body, uHits, uUse)

    If aborted And missingAt > 0 Then
        body = body & vbLf & vbLf & "※中断されたため論点" & missingAt & "以降は未調査です"
    End If

    On Error Resume Next
    modLog.LogUsage "decomposed", MODE_THOROUGH, _
        "parts=" & nParts & " ok=" & okN & " hits=" & uN, 0, uN
    On Error GoTo 0

    hits = uHits
    nHits = uN
    result = body
    ok = True
    mHandled = True
    TryDecomposed = True
End Function

' ----------------------------------------------------------------------------
' VerifyNote - 直近の入念ターンの検証注記(無ければ空)。
'   分解経路を通ったターンは自前の注記を、通らなかったターンは
'   modAskThorough の注記をそのまま返す。窓口を1つにしておかないと、
'   分解したターンに【前回の single 経路で立った注記】が付いて嘘になる
'   (modAskThorough 側の注記は RunThoroughFlow を通るまで消えないため)。
' ----------------------------------------------------------------------------
Public Function VerifyNote() As String
    If mHandled Then
        VerifyNote = mNote
    Else
        VerifyNote = modAskThorough.VerifyNote()
    End If
End Function

' ----------------------------------------------------------------------------
' StepNote - usage_log の ask_steps へ足す1語(分解した論点数)。
'   分解しなかったターンは空文字=従来の1行と1字も変わらない。
'   読んだら消す(modGateway.ConsumeStepBuf と同じ作法。消さないと次の
'   quick の1行にまで前回の dec= が残る)。
' ----------------------------------------------------------------------------
Public Function StepNote() As String
    If mParts > 0 Then StepNote = ";dec=" & mParts
    mParts = 0
End Function

' ============================================================================
' 純ロジック(LibreOfficeの実行テストで固定する)
' ============================================================================

' ----------------------------------------------------------------------------
' ShouldDecompose - 段0の判定を呼ぶかどうか(config decompose_mode + 文字数)。
'   off    = 呼ばない(機能まるごとのエスケープハッチ)
'   always = 文字数を見ずに必ず呼ぶ(検証用)
'   auto   = 既定。minChars 未満の短い質問では呼ばない
'            (「更新日は?」に判定を1回足しても、割る論点が無い)
'   知らない値は auto 扱い(config の打ち間違いで機能が黙って止まらない)。
' ----------------------------------------------------------------------------
Public Function ShouldDecompose(ByVal modeCfg As String, ByVal qLen As Long, _
                                ByVal minChars As Long) As Boolean
    Dim m As String: m = LCase$(Trim$(modeCfg))
    If m = "off" Then Exit Function
    If m = "always" Then
        ShouldDecompose = True
        Exit Function
    End If

    Dim lim As Long: lim = minChars
    If lim < 1 Then lim = 25
    ShouldDecompose = (qLen >= lim)
End Function

' ----------------------------------------------------------------------------
' DecomposeGate - 段0(分解の判定)を呼ぶかどうかの最終判定(2026-08-05 R18H
'   FB-2 / A-L9)。文字数ゲート(ShouldDecompose)に複合シグナルをORで足した1本。
' ----------------------------------------------------------------------------
' 「免責は?保険料は?」(9字)のような短くても強く当たる複合質問は、文字数
' だけを見る ShouldDecompose を素通りしていた(decompose_min_chars既定25未満。
' 実機第5報⑦)。R18-7b でORを足したが、実装は TryDecomposed の中に直書きで、
' テスト(modTestsPure16)は同じ式を独立に書き写していた。片方だけ直せば
' 「テストは通るのに実機では発動しない」が作れる状態=憲章§4-5。ここへ切り出し、
' TryDecomposed とテストの双方がこの1本だけを呼ぶ。
'   off    : 複合シグナルより優先(呼ばない=機能まるごとのエスケープハッチ)
'   always : ShouldDecompose が既に True を返すのでORの出番は無い(不変)
' Excel/COMには触れない純ロジック(modRagParse.HasCompoundSignal も同じ)。
Public Function DecomposeGate(ByVal modeCfg As String, ByVal q As String, _
                              ByVal minChars As Long) As Boolean
    Dim g As Boolean
    g = ShouldDecompose(modeCfg, Len(q), minChars)
    If Not g And LCase$(Trim$(modeCfg)) <> "off" Then g = modRagParse.HasCompoundSignal(q)
    DecomposeGate = g
End Function

' ----------------------------------------------------------------------------
' BuildPartSection - 統合段への入力になる「論点1つぶんの節」。
'   body が空 = その論点は資料を確認できなかった(検索0件 or 生成失敗)。黙って
'   節ごと落とすと「聞いたのに答えが無い論点」に気付けない(憲章§3-3)。文言は
'   modPrompts.PART_FAIL_TEXT に固定する(統合段へ「この文言は消すな」と指示する
'   のと同じ文字列でないと、指示と実物がずれる)。
'   置き場が modPrompts ではなくここなのは、これがプロンプトではなく統合入力の
'   組み立てで、呼ぶのもこのモジュールだけだから(modPromptsの残枠も守れる)。
' ----------------------------------------------------------------------------
Public Function BuildPartSection(ByVal idx As Long, ByVal part As String, _
                                 ByVal body As String) As String
    Dim t As String: t = Trim$(body)
    If LenB(t) = 0 Then
        BuildPartSection = "■論点" & idx & ": " & modPrompts.PART_FAIL_TEXT
        Exit Function
    End If

    Dim head As String: head = Trim$(part)
    If LenB(head) = 0 Then head = "論点" & idx
    BuildPartSection = "■" & head & vbLf & t
End Function

' ----------------------------------------------------------------------------
' ShouldClarify - 逆質問(番号選択肢)を出すかどうか(R16-3B・純ロジック)。
'   off      = 出さない(機能まるごとのエスケープハッチ。従来どおり回答を作る)
'   auto既定 = 選択肢が2件以上読めたときだけ出す
'   知らない値は auto 扱い(config の打ち間違いで機能が黙って止まらない)。
'
'   1件のときに出さないのは「選べない選択肢」だからで、これは逆質問ではなく
'   ただの足止めになる。読み方が1つに決まったのなら、そのまま調べたほうが早い。
' ----------------------------------------------------------------------------
Public Function ShouldClarify(ByVal modeCfg As String, ByVal nOpts As Long) As Boolean
    If LCase$(Trim$(modeCfg)) = "off" Then Exit Function
    ShouldClarify = (nOpts >= 2)
End Function

' ----------------------------------------------------------------------------
' BuildClarifyAsk - 逆質問の本文(R16-3B・純ロジック)。
'   LLMは呼ばない(段0の1回で材料は揃っている)ので、この経路は数秒で返る。
'   「番号で返信」と「書き直してもOK」を必ず並記する。番号だけを案内すると、
'   選択肢がどれも違う人が行き止まりになる(modClarify.BuildClarifyPrompt と
'   同じ約束)。例に "1と3" を出すのは、複数選べることが文面から読めないと
'   誰も試さないため(実装があっても使われない機能になる)。
'   有効期限も書く(R16H FB-6 / B-M5)。保留はブックを閉じてもTTL30分だけ生きて
'   いて、過ぎると番号だけの返事が【普通の新しい質問】として扱われる。書いて
'   おかないと、席へ戻って「1」と打った人には理由の分からない答えが返る。
' ----------------------------------------------------------------------------
Public Function BuildClarifyAsk(opts() As String, ByVal nOpts As Long) As String
    Dim sb As String
    sb = ChrW(&HD83D) & ChrW(&HDCAD) & " ご質問はいくつかの読み方ができます。" & _
         "どれを調べますか?" & vbLf & vbLf

    Dim i As Long
    For i = 1 To nOpts
        Dim t As String: t = Trim$(opts(LBound(opts) + i - 1))
        If LenB(t) = 0 Then t = "(読み方" & i & ")"
        sb = sb & "  " & i & ") " & t & vbLf
    Next i

    sb = sb & vbLf & "番号で返信してください(例: 1 / 1と3)。" & _
         "質問を書き直していただいてもOKです。" & vbLf & _
         "(この聞き返しは30分で無効になります)"
    BuildClarifyAsk = sb
End Function

' ----------------------------------------------------------------------------
' PerPartTopK - 論点1つあたりに渡すチャンク数。
'   合計を論点数で割るが、下限6は割らない。論点あたり2～3件まで絞ると、
'   その論点の根拠が1資料に依存して「資料には見当たらない」が量産される
'   (max_context_chars はLLM呼び出しごとに独立して効くので、合計が増えても
'    1回ぶんの上限には当たらない)。
' ----------------------------------------------------------------------------
Public Function PerPartTopK(ByVal thoroughK As Long, ByVal partN As Long) As Long
    Dim k As Long
    If partN > 0 Then k = thoroughK \ partN
    If k < 6 Then k = 6
    PerPartTopK = k
End Function

' ============================================================================
' 内部ヘルパー
' ============================================================================

' ----------------------------------------------------------------------------
' RunDecideStage - 段0(1回だけのLLM呼び出し)。戻り値は
'   "single" / "parts" / "clarify"。parts のときだけ parts()/nParts が埋まる。
'   判定が読めない(#ERR・タグ崩れ)ときは single=従来フローへ寛容退化する。
' ----------------------------------------------------------------------------
Private Function RunDecideStage(ByVal q As String, ByVal history As String, _
                                ByRef parts() As String, ByRef nParts As Long, _
                                ByRef opts() As String, ByRef nOpts As Long) As String
    RunDecideStage = "single"
    nParts = 0
    nOpts = 0

    Dim maxP As Long
    maxP = modConfig.GetLong("decompose_max_parts", 3)
    If maxP < PARTS_MIN Then maxP = PARTS_MIN
    If maxP > PARTS_MAX Then maxP = PARTS_MAX

    ' この時点では論点数が未定=分母を出さない(出せば必ず後でずれる)。
    Stage 0, 0, 0, "質問の論点を確認中…"

    Dim lat As Long
    Dim resp As String
    resp = modGateway.CallLLM(modPrompts.BuildDecomposePrompt(q, history, maxP), "decompose", _
        modConfig.GetString("expand_effort", "low"), _
        modConfig.GetString("expand_verbosity", "low"), DecideModel(), lat)
    If modRagParse.IsErrorResponse(resp) Then Exit Function

    Dim v As String
    v = modRagParse.ParseDecomposeVerdict(resp)
    If v = "parts" Then
        nParts = modRagParse.ParseParts(resp, maxP, parts)
        If nParts < PARTS_MIN Then v = "single"      ' 1本しか割れなかった=分解の意味が無い
    ElseIf v = "clarify" Then
        nOpts = modRagParse.ParseOptions(resp, maxP, opts)
    End If
    RunDecideStage = v
End Function

' ----------------------------------------------------------------------------
' AskWhichTopic - 逆質問を提示し、保留状態(元質問+選択肢)を残す(R16-3B)。
'   保留の置き場は modClarify に一本化する(資料の聞き返しと同じ ui_state・
'   同じTTL30分・同じ HasPending/MergeAnswer の配線に乗る)。ここで別の場所へ
'   置くと、modApp の送信経路に2つ目の「保留の途中か?」判定が要る。
'   このターンは ok=False(R16H FA-2)なので会話履歴には積まれない。それで
'   困らないのは、次のターンで送られる合成質問を modClarify.MergeTopicAnswer が
'   ui_state の保留(元質問+選択肢)だけから組み立てるため=履歴に依存しない。
' ----------------------------------------------------------------------------
Private Function AskWhichTopic(ByVal q As String, opts() As String, ByVal nOpts As Long) As String
    Dim joined As String
    Dim i As Long
    For i = 1 To nOpts
        If LenB(joined) > 0 Then joined = joined & "|"
        joined = joined & Trim$(opts(LBound(opts) + i - 1))
    Next i

    On Error Resume Next
    modClarify.SaveTopicPending q, joined
    modLog.LogUsage "clarify_topic", MODE_THOROUGH, _
        "opts=" & nOpts & " q=" & modUtil.SafeLeft(q, 80)
    modUIMain.SetStage ""
    On Error GoTo 0

    AskWhichTopic = BuildClarifyAsk(opts, nOpts)
End Function

' 段0のモデル。拡張段と同じ選び方(判定だけの軽い段に本命モデルは要らない)。
Private Function DecideModel() As String
    DecideModel = modConfig.GetString("expand_model", "")
    If LenB(DecideModel) = 0 Then DecideModel = modConfig.GetString("quick_model", "gpt-5.5")
End Function

' ----------------------------------------------------------------------------
' OnePart - 1論点ぶんの検索+副下書き。戻り値=下書き本文(失敗時は空文字)。
'   uHits/uN へは chunk_id で重複排除しながら足し込む(出典突合の材料)。
'   検索は軽量(skipExpand=True=拡張段も再ランク段も通さない。裁定1)。既に1論点
'   まで割ってあるものを更にばらすと論点の外の資料が混ざり、論点あたりtopK=6～8
'   では再ランクの並べ替える余地も無い。この2段を省くことで、質問全体のAI呼び
'   出しは 論点数+6 回(3論点で9回)に収まる。
' ----------------------------------------------------------------------------
Private Function OnePart(ByVal q As String, ByVal part As String, ByVal idx As Long, _
                         ByVal topKPer As Long, ByRef uHits() As Hit, ByRef uN As Long, _
                         ByVal strictG As Boolean, ByVal ansTags As Boolean, _
                         ByVal mdl As String, ByRef lastErr As String) As String
    Dim pHits() As Hit
    Dim n As Long
    n = modAskRetrieve.RunMultiRetrieve(part, MODE_THOROUGH, topKPer, pHits, Nothing, 0, True)
    If n < 1 Then
        ' -1 = 埋め込み失敗(E0203)。全論点で同じ理由で落ちたときに E0202(API失敗)
        ' の案内へ丸めると、直す手順が違う障害を同じ文面で案内することになる。
        If n = -1 Then lastErr = "#ERR:E0203:埋め込みに失敗しました"
        On Error Resume Next
        modLog.LogUsage "multi_part_nohit", MODE_THOROUGH, _
            "n=" & idx & " ret=" & n & " q=" & modUtil.SafeLeft(part, 80)
        On Error GoTo 0
        Exit Function
    End If

    MergeHits uHits, uN, pHits, n

    Dim lat As Long
    Dim draft As String
    draft = modGateway.CallLLM( _
        modPrompts.BuildPartDraftPrompt(q, part, idx, pHits, n, strictG, ansTags), _
        "multi_draft", modConfig.GetString("thorough_draft_effort", "high"), _
        "medium", mdl, lat)
    If modRagParse.IsErrorResponse(draft) Then
        lastErr = draft
        On Error Resume Next
        modLog.LogUsage "multi_part_err", MODE_THOROUGH, modUtil.SafeLeft(draft, 120)
        On Error GoTo 0
        Exit Function
    End If

    OnePart = Trim$(modAsk.ApplyAnswerTags(draft))
End Function

' ----------------------------------------------------------------------------
' Integrate - 統合(1回)+自己点検(1回)+検証(1回)。戻り値=最終本文。
'   どこで失敗しても、論点ごとの下書きを並べた sections までは必ず返す。
'   統合段の事故で「調べ終わった内容ごと消える」のがいちばん惜しい壊れ方で、
'   ここは新機能側のサーキットブレーカー(CONTRIBUTING §2.4)そのもの。
'   ESC(Err18)もここでは同じ扱い=その時点の成果を返して静かに終わる。
' ----------------------------------------------------------------------------
Private Function Integrate(ByVal q As String, ByVal sections As String, uHits() As Hit, _
                           ByVal uN As Long, ByVal partN As Long, ByVal total As Long, _
                           ByVal strictG As Boolean, ByVal ansTags As Boolean, _
                           ByVal mdl As String, ByVal prevU As String, _
                           ByVal prevA As String) As String
    On Error GoTo Bail

    Dim lat As Long
    Stage partN, 2 + partN, total, "論点ごとの答えを統合中…"
    Dim merged As String
    merged = modGateway.CallLLM(modPrompts.BuildMergePrompt(q, sections), "multi_merge", _
        modConfig.GetString("thorough_draft_effort", "high"), _
        modConfig.GetString("thorough_draft_verbosity", "high"), mdl, lat, prevU, prevA)
    If modRagParse.IsErrorResponse(merged) Then
        LogSkip "multi_merge_skip", merged
        merged = sections          ' 下書きは完成している。捨てずにそのまま並べる
    Else
        merged = modAsk.ApplyAnswerTags(merged)
    End If

    Stage partN, 3 + partN, total, "統合した回答を自己点検中…"
    Dim critique As String
    critique = modGateway.CallLLM(modPrompts.BuildCritiquePrompt(q, merged, uHits, uN), _
        "multi_critique", modConfig.GetString("thorough_critique_effort", "medium"), _
        "low", mdl, lat)
    If modRagParse.IsErrorResponse(critique) Then
        LogSkip "multi_critique_skip", critique
        critique = ""
    Else
        critique = modAsk.ApplyAnswerTags(critique)
    End If

    Stage partN, 4 + partN, total, "検証中…"
    Dim verified As String
    verified = modGateway.CallLLM( _
        modPrompts.BuildDeepVerifyPrompt(q, merged, uHits, uN, strictG, ansTags, critique), _
        "multi_verify", modConfig.GetString("thorough_verify_effort", "high"), _
        modConfig.GetString("deep_verify_verbosity", "medium"), mdl, lat)

    If modRagParse.IsErrorResponse(verified) Then
        ' 注記は本文へ混ぜない(modAsk が整形の後に足す=履歴と部内共有へ入らない)。
        mNote = vbLf & vbLf & _
            "(注: 検証段階でエラーが発生したため、統合した下書きの内容を表示しています。)"
        Integrate = merged
    Else
        Integrate = modAsk.ApplyAnswerTags(verified)
    End If
    Exit Function

Bail:
    Err.Clear
    Resume BailOut
BailOut:
    ' 2026-08-05(R16H FA-8 / A-L14): 論点ごとの下書きをそのまま出すことを
    ' 本文でも言う。ここへ落ちるのは統合段の事故(ESC=Err18を含む)で、
    ' 黙って渡すと「■論点1/■論点2…」が並んだ回答が【統合・検証まで通った
    ' 完成品】に見える。見た目が完成品なのに検証していない回答は、その差が
    ' 一番危ない(憲章§3-3: 分からないことは分からないと言う)。
    LogSkip "multi_integrate_fallback", "parts=" & partN
    Integrate = sections & vbLf & vbLf & _
        "※中断のため論点別の回答をそのまま表示しています(統合・検証は未実施)"
End Function

' 補助段が落ちたことの記録(憲章§4-1: 止めないが、必ず痕跡は残す)。
Private Sub LogSkip(ByVal tag As String, ByVal detail As String)
    On Error Resume Next
    modLog.LogUsage tag, MODE_THOROUGH, modUtil.SafeLeft(detail, 120)
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' MergeHits - 全論点のヒットを chunk_id で重複排除しながら1本へ連結する。
'   chunk_id が空のヒット(旧データ・テストダブル)は重複判定の材料が無いので
'   必ず足す側へ倒す。落とす側へ倒すと、根拠に使った資料が索引から消えて
'   「(出典確認できず)」が嘘で付く。
' ----------------------------------------------------------------------------
Private Sub MergeHits(ByRef uHits() As Hit, ByRef uN As Long, srcHits() As Hit, _
                      ByVal srcN As Long)
    Dim i As Long, j As Long
    For i = 1 To srcN
        Dim dup As Boolean: dup = False
        If LenB(srcHits(i).chunk_id) > 0 Then
            For j = 1 To uN
                If StrComp(uHits(j).chunk_id, srcHits(i).chunk_id, vbTextCompare) = 0 Then
                    dup = True
                    Exit For
                End If
            Next j
        End If

        If Not dup Then
            If uN = 0 Then
                ReDim uHits(1 To 1)
            Else
                ReDim Preserve uHits(1 To uN + 1)
            End If
            uN = uN + 1
            uHits(uN) = srcHits(i)
        End If
    Next i
End Sub

' ----------------------------------------------------------------------------
' Stage - 分解経路専用の実況。分母は 1(判定)+論点数+3(統合・点検・検証)で、
'   single 経路の (n/6) とは別物なので modMode の番号計算は使わない
'   (あちらへ論点数を持ち込むと、分解しないときの番号まで動く)。
'   段0の時点では論点数が分からない=分母も出さない(嘘の分母を出すくらいなら
'   番号を出さない、は modMode.AskStageText と同じ判断)。
'   「※応答なし表示でも処理中」は R16-2c と同じ一文。入念は数分ブロックし、
'   その間 Windows が「応答なし」と出すのは正常だと先に言っておく。
'   経過時間(R16H FB-8)も添える。「正常だ」と書いてあっても、動いている証拠が
'   何も動かない画面では、待ってよい時間の見当が付かない。秒が増えていく表示は
'   「止まっていない」ことをいちばん確実に伝える。
' ----------------------------------------------------------------------------
Private Sub Stage(ByVal partN As Long, ByVal idx As Long, ByVal total As Long, _
                  ByVal label As String)
    Dim head As String
    head = ChrW(&HD83E) & ChrW(&HDDEC) & " 入念"
    If partN > 0 Then head = head & "(" & partN & "論点)"
    If total > 0 And idx > 0 Then head = head & " " & idx & "/" & total & "段:"

    On Error Resume Next
    modUIMain.SetStage head & " " & label & ElapsedText() & " ※応答なし表示でも処理中"
    On Error GoTo 0
End Sub

' 開始からの経過。1分未満は「経過S秒」、それ以上は「経過M分S秒」。
' 日跨ぎ(深夜0時でTimerが0へ戻る)は +86400 で補正する(modWorkExcel.Debounced
' と同じ作法)。補正しないと負の秒数が出て、進んでいるのに戻って見える。
Private Function ElapsedText() As String
    If mT0 = 0 Then Exit Function
    Dim s As Double: s = Timer - mT0
    If s < 0 Then s = s + 86400
    Dim n As Long: n = CLng(Int(s))
    If n < 60 Then
        ElapsedText = " 経過" & n & "秒"
    Else
        ElapsedText = " 経過" & (n \ 60) & "分" & (n Mod 60) & "秒"
    End If
End Function
