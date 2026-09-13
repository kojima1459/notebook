Attribute VB_Name = "modAsk"
Option Explicit

' modAsk - 2速QA(すぐ聞く/しっかり調べる)のオーケストレーション。
' modRetrieve.Search→modPrompts.Build*Prompt→modGateway.CallLLM の順に処理し、
' 進捗をmodUIMain経由で実況する(MASTER_SPEC §7.3)。
'
' 設計判断:
'   ・quick=Search→出典先出し→CallLLM / deep=+検証段。
'   ・Answer()は最終回答テキストのみ返す契約。hits等はmLast*に保持しUIへ渡す。
'   ・検索0件はLLM未呼び出しで定型文。埋め込み失敗はE0203案内。
'   ・質問3000字打ち切り、連打防止はmAsking。ESCはErr18で「操作を中断」。
'   ・[[FOLLOWUP:...]]は本文から除去し「深掘り候補」ブロックへ整形する。

Private Const MODE_QUICK As String = "quick"
Private Const MODE_DEEP As String = "deep"
' モードの方針は modMode が単一情報源(実測の根拠もそちら)。
Private Const MODE_THOROUGH As String = "thorough"
Private Const MAX_QUESTION_CHARS As Long = 3000
Private Const QUESTION_RANGE_NAME As String = "mb_question"
Private Const UI_STATE_MODE_KEY As String = "mode"
Private Const EMPTY_QUESTION_MESSAGE As String = _
    "質問が入力されていません。ホームの入力欄に質問を書いてから、もう一度お試しください。"

Private Const HISTORY_MAX_TURNS As Long = 3
Private Const HISTORY_SEP As String = "<<<__QA_TURN__>>>"

' 続けて質問(D11): prevU/prevA の履歴区切り(新しい順;;;区切り)。
Private Const FOLLOWUP_PAIR_SEP As String = ";;;"

Private mAsking As Boolean
Private mHistory As String

' 続けて質問(D11)用の履歴。mLastCleanAnswerは深掘り候補除去後の本文。
Private mPrevU As String
Private mPrevA As String
Private mLastCleanAnswer As String

Private mLastQuestion As String
Private mLastAnswer As String
Private mLastMode As String
' 感想は1回の回答につき1回だけ記録する(多重カウント防止)。
Private mFeedbackDone As Boolean
Private mLastHits() As Hit
Private mLastNHits As Long
' R34 F1: 生成に渡した件数(契約は LastGenHitCount の見出し)。
Private mLastGenN As Long
Private mLastSeconds As Long

' AskFromUI - ホームの質問セル+モード(ui_state)を読み、Answer実行→
'   RenderAnswer(MASTER_SPEC §7.3)。連打防止ガード付き。
Public Sub AskFromUI()
    If mAsking Then
        On Error Resume Next
        modUIMain.SetStage "処理中です。少しお待ちください…"
        On Error GoTo 0
        Exit Sub
    End If

    mAsking = True
    On Error GoTo Fail

    Dim q As String
    q = ReadQuestionFromHome()
    Dim mode As String
    mode = ReadModeFromUiState()

    Dim ans As String
    ans = Answer(q, mode)

    modUIMain.RenderAnswer ans, mLastHits, mLastNHits, mLastMode, mLastSeconds

    mAsking = False
    Exit Sub

Fail:
    modLog.LogError "E0602", "modAsk.AskFromUI", Err.Description
    Err.Clear
    On Error GoTo 0
    mAsking = False
End Sub

' Answer - 質問文とモードから回答テキストを組み立てて返す(MASTER_SPEC §7.3)。
'   契約どおりの公開API。会話履歴なしの単発質問として実行する(実体は
'   AnswerWithContext。裁定D11のAskFollowupと本体を共有する)。
Public Function Answer(ByVal question As String, ByVal mode As String) As String
    Answer = AnswerWithContext(question, mode, "", "", False)
End Function

' CanFollowup - 追質問できる直近回答があるか(D11)。followup_max_pairs<=0で
'   常にFalse=機能停止のエスケープハッチを兼ねる。
Public Function CanFollowup() As Boolean
    If LenB(mPrevU) = 0 Then
        mPrevU = modState.LoadState("nexus_ask_prevu", "")
        mPrevA = modState.LoadState("nexus_ask_preva", "")
    End If
    CanFollowup = (LenB(mPrevU) > 0)
End Function

' AskFollowup - 会話履歴を添えた追質問(D11)。Searchも再実行し、mPrevU/mPrevAを
'   CallLLMのprevU/prevAへ渡す。履歴が空でも単発質問として正常に退化する。
Public Sub AskFollowup(ByVal followupText As String)
    If mAsking Then
        On Error Resume Next
        modUIMain.SetStage "処理中です。少しお待ちください…"
        On Error GoTo 0
        Exit Sub
    End If

    mAsking = True
    On Error GoTo Fail

    Dim mode As String
    mode = ReadModeFromUiState()

    ' mPrevU/mPrevAはByValスナップショットで渡す(AnswerWithContextの成功時に
    ' 今回のターンが履歴へ追記されても、今回の呼び出し自体には影響しない)。
    Dim ans As String
    ans = AnswerWithContext(followupText, mode, mPrevU, mPrevA, True)

    modUIMain.RenderAnswer ans, mLastHits, mLastNHits, mLastMode, mLastSeconds

    mAsking = False
    Exit Sub

Fail:
    modLog.LogError "E0602", "modAsk.AskFollowup", Err.Description
    Err.Clear
    On Error GoTo 0
    mAsking = False
End Sub

' AnswerWithContext - Answer/AskFollowup共通の回答生成本体。prevU/prevAは
'   会話履歴(空=なし)。isFollowupはusage_logの識別用。
Private Function AnswerWithContext(ByVal question As String, ByVal mode As String, _
                                   ByVal prevU As String, ByVal prevA As String, _
                                   ByVal isFollowup As Boolean) As String
    Dim tStart As Double
    tStart = Timer

    Dim q As String
    q = modUtil.SafeLeft(Trim$(question), MAX_QUESTION_CHARS)
    Dim mdMode As String
    mdMode = NormalizeMode(mode)
    modAskRetrieve.PlanAskStages mdMode   ' R13-9b: 何段通すかを先に決めてから実況する
    mLastGenN = 0                         ' R34 F1

    If LenB(q) = 0 Then
        ' 空質問でもmLast*を必ず更新する。しないと前回のヒットが残り、空クリックが
        ' その資料で回答したように見える(Wave4実バグ)。mLastMode=""は「検索も回答
        ' 生成もしなかった」印でmodUIMain.RenderAnswerも使う。
        Dim emptyHits() As Hit
        mLastQuestion = q
        mLastAnswer = EMPTY_QUESTION_MESSAGE
        mLastMode = ""
        mLastHits = emptyHits
        mLastNHits = 0
        mLastSeconds = 0
        AnswerWithContext = EMPTY_QUESTION_MESSAGE
        Exit Function
    End If

    Dim hits() As Hit
    Dim nHits As Long
    Dim result As String
    Dim ok As Boolean
    ok = False

    ' 2026-07-28(レビュー M-4): 回答バッファを毎ターン空にする。設定するのは
    ' DecorateWithFollowups(成功ターンだけ)なので、ここで消さないと 0件回答・
    ' 聞き返しのターンでも前回の回答が residual として残り、✅を押したときに
    ' 噛み合わないQ&Aが部内へ流れる。
    mLastCleanAnswer = ""

    On Error Resume Next
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo 0

    On Error GoTo Fail

    modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDD0D) & " 検索中…"
    Dim topK As Long
    topK = TopKFor(mdMode)

    ' 多段RAG(§C)。retrieve_mode=singleで従来の単段Searchへ完全退化(RunUnscoped)。
    ' R13-5c: 「続けて質問」かつ「しっかり調べる」のときだけ、会話で引用済みの
    ' 資料へスコープを絞って掘り下げる。新規質問の deep は従来どおり。
    If isFollowup And mdMode = MODE_DEEP Then
        nHits = modAskRetrieve.RunDeepScoped(q, mdMode, topK, hits)
    Else
        nHits = modAskRetrieve.RunUnscoped(q, mdMode, topK, hits)
    End If

    If nHits = -1 Then
        result = modLog.FriendlyMessage("E0203") & vbLf & "(コード: E0203)"
    ElseIf nHits = 0 Then
        modLog.LogError "E0601", "modAsk.Answer", "query=" & modUtil.SafeLeft(q, 200)
        ' 「見つかりません」で終わらせず、探すべき資料の種類まで示す。
        ' 本棚が空では回答自体が成立しない以上、次の一手を出すのが唯一の親切。
        result = modClarify.MissingDocGuide(q)
        On Error Resume Next
        modInsightIo.EmitGap q, "no_hit"    ' 資料が無い領域として部内に共有する
        On Error GoTo Fail
    Else
        modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDCC4) & " " & nHits & "件の資料がヒット"
        modUIMain.RenderSourcesPreview hits, nHits

        If modAskRetrieve.IsTooVague(q, hits, nHits) Then
            result = modClarify.BuildClarifyPrompt(q, modAskRetrieve.HitSourceList(hits, nHits))
            On Error Resume Next
            modLog.LogUsage "ambiguous_clarify", mdMode, modUtil.SafeLeft(q, 80)
            On Error GoTo Fail
        ElseIf mdMode = MODE_THOROUGH Then
            ' R14-8a: 入念だけ専用の6段。R16-3A: 先に論点分解を試し不発なら6段へ。
            If Not modAskMulti.TryDecomposed(q, hits, nHits, ok, result, prevU, prevA) Then
                result = modAskThorough.RunThoroughFlow(q, hits, nHits, ok, HistoryBlock(), prevU, prevA)
            End If
            If ok Then
                ' 検証段の注記は整形の【後】(deepと同型。履歴と共有へ混ぜない)。
                result = DecorateWithFollowups(result) & modAskMulti.VerifyNote()
            Else
                result = modRagParse.BuildErrorAnswer(result)
            End If
        ElseIf modMode.UseVerify(mdMode) Then
            result = RunDeepFlow(q, hits, nHits, ok, prevU, prevA, mdMode)
        Else
            result = RunQuickFlow(q, hits, nHits, ok, prevU, prevA)
        End If
    End If

    modUIMain.SetStage ""
    GoTo Done

Fail:
    Dim errNum As Long
    errNum = Err.Number
    Dim errDesc As String
    errDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup4
FailCleanup4:
    Err.Clear
    On Error GoTo 0

    If errNum = 18 Then
        result = "操作を中断しました。もう一度質問するときは、質問するボタンを押してください。"
    Else
        ' 出所ラベルは公開APIのAnswerで記録(内部関数名を文字列で書くとlintの参照検査が誤検知するため)。
        modLog.LogError "E0602", "modAsk.Answer", "共通本体(AnswerWithContext) mode=" & mdMode & " err=" & errDesc
        result = modLog.FriendlyMessage("E0602") & vbLf & "(コード: E0602)"
    End If

    On Error Resume Next
    modUIMain.SetStage ""
    ' 同上。Resume で抜けないと Done: の後始末のエラーが呼び出し元へ素通りする。
    Resume Done

Done:
    ' 経過時間は modUtilText.ElapsedMsSince に一本化(Timerは0時に0へ戻るため、
    ' 素の引き算だと日付をまたいだ質問で負の値になる。2026-07-31 R11-F2)。
    Dim elapsedMs As Long
    elapsedMs = CLng(modUtilText.ElapsedMsSince(tStart))
    Dim elapsedSec As Long
    elapsedSec = CLng(elapsedMs / 1000#)

    ' 成功ターンのみ履歴に積む。積むのは深掘り候補ブロックを含まない
    ' 除去後の本文(mLastCleanAnswer。DecorateWithFollowupsが設定)。
    If ok Then
        AppendHistory q, mLastCleanAnswer
        AppendFollowupPair q, mLastCleanAnswer
        ' R13-5b: この回答が根拠にした資料を「会話の出典」として覚える。新規質問
        ' (非followup)なら覚え直す(=前の話題を引きずらない)。R13 L-batch: 会話の
        ' 出典メモリだけは最大8件まで覚える(逆質問の材料は従来どおり4件)。深掘りの
        ' スコープはここが元になるので、4件で切ると「引用したのに次で対象外」が起きる。
        If nHits > 0 Then modFollowup.RememberCitedSources modAskRetrieve.HitSourceList(hits, nHits, 8), Not isFollowup
    End If

    ' 低関連度警告(表示専用): 履歴(AppendHistory/mLastCleanAnswer)は上で確定
    ' 済みのため、ここでresultに警告を足しても履歴側には混入しない。
    ' R46: nHits=0 の回にもガード(※)を付ける。旧条件 nHits>0 だと、本棚から
    ' 1件も引けなかった=いちばん根拠の薄い回答にだけ注意書きが出なかった。
    ' 低関連度の⚠は hits を見るので従来どおり nHits>0 のときだけ。
    ' R46: 回答の数字が資料に実在するかを照合して注記を足す。常設ガードより
    ' 【前】に置く(ガードは最終段落という R44 の裁定)。
    ' 2026-09-12(R47 自己訂正): R46 でここへ足した「nHits=0 なら GuardOnly」の
    ' Else は到達不能かつ前提も誤りだった。経緯は modGround の冒頭に記録。
    If ok And nHits > 0 Then
        result = modGround.AppendGroundNote(result, hits, nHits, q)
        result = modAskRetrieve.ApplyLowHitWarning(result, hits, nHits)
    End If

    ' チャット履歴シート記録(modChatLog、core層。書込失敗で死なない設計)。
    If LenB(q) > 0 Then
        On Error Resume Next
        modChatLog.LogTurn q, result, mdMode
        On Error GoTo 0
    End If

    mLastQuestion = q
    mLastAnswer = result
    mLastMode = modMode.NoteAnswered(ok, mdMode)   ' R33H F7(抑止の実体はmodMode)
    mLastHits = hits
    mLastNHits = nHits
    If mLastGenN < nHits Then mLastGenN = nHits   ' R34 F1
    mLastSeconds = elapsedSec
    mFeedbackDone = False   ' 新しい回答に対する感想を受付可能にする

    Dim logDetail As String
    logDetail = "q=" & modUtil.SafeLeft(q, 200)
    If isFollowup Then logDetail = "followup " & logDetail
    modLog.LogUsage "ask", mdMode, logDetail, elapsedMs, nHits

    ' R13 F8: 段ごとの所要時間は質問1回につき【1行】にまとめて書き出す。1段
    ' 1行だと1問で10行前後になり、2,000行で回る usage_log が約180問で一周して
    ' feedback_green 等の履歴を押し出す(modDashStat の前月比が静かに壊れる)。
    ' Done: は成否どちらの経路も必ず通るので、書き出しとバッファの掃除はこの
    ' 1点だけでよい(ConsumeStepBufは読んだら空にし次の質問へ持ち越さない)。
    On Error Resume Next
    Dim stepBuf As String
    stepBuf = modGateway.ConsumeStepBuf()
    stepBuf = stepBuf & modAskMulti.StepNote()   ' R16-3A: 分解時だけ "dec=論点数"
    If LenB(stepBuf) > 0 Then
        modLog.LogUsage "ask_steps", mdMode, stepBuf, elapsedMs
    End If
    On Error GoTo 0

    modStats.Bump "ask_" & mdMode & "_total"
    modStats.AddExp "question"

    AnswerWithContext = result
End Function

' 直近回答の本文(整形済み)を返す(Nexus UI等の外部表示用ゲッター)。
Public Function LastAnswerText() As String
    LastAnswerText = mLastAnswer
End Function

' 直近回答で最上位スコアのソース名(P2P感謝状の宛先解決用)。
Public Function LastTopSource() As String
    ' mLastHitsは ReDim(1 To n)。0始まり走査は添字エラーになる(2026-07-20実バグ)。
    Dim bestI As Long: bestI = -1
    Dim bestScore As Double: bestScore = -1E+30
    Dim i As Long
    For i = 1 To mLastNHits
        If mLastHits(i).score > bestScore Then
            bestScore = mLastHits(i).score
            bestI = i
        End If
    Next i
    If bestI >= 1 Then LastTopSource = mLastHits(bestI).source
End Function

' Peek View用の読み取り専用アクセサ(添字0始まり)。内部状態は変更しない。
'
' 2026-07-28(レビュー C-2): 公開契約は「0始まり」で、呼び出し側3箇所
' (modPeek.RenderCitations / modPeek.ShowPeek / modMentor.FindExpert /
'  modLive.UniqueSourceCount)はすべて For i = 0 To n - 1 で回すが、内部の
' mLastHits は ReDim(1 To n)。i をそのまま添字にすると i=0 で実行時エラー9
' (添字が範囲外)。呼び出し側が On Error で握るため落ちず、代わりに ・出典
' チップが1枚も出ない(Peek View が丸ごと死ぬ) ・「この分野は さんが詳しい
' です」という空名ボタンが出る、という「静かに壊れている」状態になっていた
' (すぐ上の LastTopSource には同じ罠のコメントが残るのにアクセサ側だけ
' 直し漏れ)。ここで 1 始まりへ変換する。
Public Function LastHitCount() As Long
    LastHitCount = mLastNHits
End Function
' R34 F1: 生成に渡した件数(deepの近傍込み)。読んでよいのは出典突合(modMode)だけ。
Public Function LastGenHitCount() As Long
    LastGenHitCount = mLastGenN
    If LastGenHitCount < mLastNHits Then LastGenHitCount = mLastNHits
End Function
Public Function LastHitSource(ByVal i As Long) As String
    If i >= 0 And i < LastGenHitCount() Then LastHitSource = mLastHits(i + 1).source
End Function
Public Function LastHitPage(ByVal i As Long) As Long
    If i >= 0 And i < LastGenHitCount() Then LastHitPage = mLastHits(i + 1).page
End Function
Public Function LastHitOrigin(ByVal i As Long) As String
    If i >= 0 And i < LastGenHitCount() Then LastHitOrigin = mLastHits(i + 1).origin
End Function
Public Function LastHitPeek(ByVal i As Long) As String
    If i >= 0 And i < LastGenHitCount() Then LastHitPeek = mLastHits(i + 1).full_text
End Function

' 回答の信頼度(2=根拠あり/1=部分的/0=乏しい)。検索スコアを人間に見える形に
' して「確認すべきときだけ確認させる」。閾値=config confidence_score_x100。
' R46 A-4: 判定の実体は modMode.ConfidenceOf(スケール非依存の coverage と
' flatness で決める)。旧実装は score>=0.55 の【件数】だけを見ており、
' 0.55 は埋め込みの絶対スケールに依存するため固有名詞の一致だけで超えた
' (「三井住友の株価」で🟢)。理由と実測は modMode 側の注記が正。
Public Function LastConfidence() As Long
    If mLastNHits < 1 Then Exit Function
    Dim body As String, sc() As Double
    ReDim sc(1 To mLastNHits)
    Dim i As Long
    On Error Resume Next
    For i = 1 To mLastNHits
        body = body & mLastHits(i).full_text & vbLf
        sc(i) = mLastHits(i).score
    Next i
    On Error GoTo 0
    LastConfidence = modMode.ConfidenceOf(mLastQuestion, body, sc, mLastNHits)
End Function

' 信頼度の説明文(バッジ用)。検索しなかったターンでは空文字=バッジを出さない。
Public Function LastConfidenceText() As String
    If LenB(mLastMode) = 0 Then Exit Function
    Select Case LastConfidence()
        Case 2
            LastConfidenceText = ChrW(&HD83D) & ChrW(&HDFE2) & _
                " 本棚の資料と強く一致(" & mLastNHits & "件)"
        Case 1
            LastConfidenceText = ChrW(&HD83D) & ChrW(&HDFE1) & _
                " 部分的に一致 ― 下の出典で原文をご確認ください"
        Case Else
            LastConfidenceText = ChrW(&HD83D) & ChrW(&HDD34) & _
                " 本棚に十分な根拠なし ― 内容をうのみにしないでください"
    End Select
End Function

' NoteGeneralAnswered - 一般アシスタントの回答を「直近の回答」として記録する
'   (R14-1b / RC2)。nHits=0・本文空が modMode.ShouldEmitInsight の
'   「発信しない」根拠(mLastHits は残るが nHits=0 では誰も走査しない)。
'   a=その回答本文(前のRAG回答が残ると訂正共有が誤爆する。R14-G1)。
Public Sub NoteGeneralAnswered(ByVal q As String, ByVal a As String)
    mLastQuestion = q
    mLastAnswer = a
    mLastMode = "general"
    mLastCleanAnswer = ""
    mLastNHits = 0
    mLastGenN = 0                         ' 一般回答後に前ターンの出典チップ導線(ShowPeek/OnOpenSource)が生き返るのを断つ(R34H)
    mFeedbackDone = False
End Sub

' NoteAnswerFailed - 回答を作れなかったターン(R14-G2)。感想ボタンは
'   「まず質問して回答を受け取ってから」で断る=残留状態で撃たせない。
Public Sub NoteAnswerFailed()
    mLastQuestion = ""
    mFeedbackDone = True
End Sub

' CanShareInsight - 直近回答を部内へ発信してよいか(真理表と門は
'   modMode.EmitInsightAllowed)。UI層(modAppAct)の訂正共有も同じ窓口を通す。
Public Function CanShareInsight() As Boolean
    CanShareInsight = modMode.EmitInsightAllowed(mLastMode, mLastNHits)
End Function

Public Sub FeedbackGreen()
    If Not FeedbackAccepted() Then Exit Sub
    ' selfsolve_totalは個人統計のみ。感謝EXPは自己申告では付けず、P2Pで他者の感謝状を受領した時だけ(modP2P)。
    modStats.Bump "selfsolve_total"
    ' 節約時間の日付キー蓄積。日/月/年キーなので跨げば自動リセット、過去キーが
    ' そのまま履歴になる(modBoardのウィジェット/ビーコンが読む)。発火点は
    ' FeedbackAcceptedガードの内側=多重カウント不可。R13 L-batch: 1解決の分数は
    ' modP2PIo.MinutesPerSelfsolve()(config minutes_per_selfsolve、既定15)へ
    ' 一本化済み。ここだけ 15 の直値が残り、係数を変えた組織では「加算は15分・
    ' 表示は新係数」という食い違いが積み上がっていた。加算側も同じ窓口から読む。
    Dim perSolve As Long: perSolve = modP2PIo.MinutesPerSelfsolve()
    On Error Resume Next
    modStats.Bump "sv:d:" & modUtilText.IsoDateCompact(Date), perSolve
    modStats.Bump "sv:m:" & modUtilText.IsoYm(Date), perSolve
    modStats.Bump "sv:y:" & modUtilText.IsoYear(Date), perSolve
    ' 2026-07-31(レビュー R8 F8): 加算した「今日の節約時間」は、この直後に共有
    ' フォルダのビーコンへ反映する必要がある(起動時にしか発信していないため、
    ' 全員のビーコンが「今日 0分」のまま置かれていた)。ただし modBoard は UI層で、
    ' ここ(qa層)から呼ぶと層の向きが逆になる。発信は UI層の呼び出し元
    ' (modAppAct.OnActResolve)が担当する。
    On Error GoTo 0
    modLog.LogUsage "feedback_green", mLastMode, "q=" & modUtil.SafeLeft(mLastQuestion, 200)

    ' 部内への発信は「本棚の資料を根拠に答えたターン」だけ(理由と真理表は
    ' modMode.ShouldEmitInsight)。R14-1bの一般モード解禁もM-4の誤爆も同型。
    ' R14-G2: 本文が空=生成に失敗したターン。従来は感謝状だけが飛び、使われていない
    ' 資料の作者へ嘘が届いた。共有と同じ条件へ揃える(憲章§3-3)。
    Dim mayEmit As Boolean
    mayEmit = CanShareInsight() And (LenB(Trim$(mLastCleanAnswer)) > 0)
    On Error Resume Next
    If mayEmit Then
        modP2P.EmitThanksForLastAnswer   ' 他者の共有ナレッジ由来なら作者へ感謝状(自作/出所不明は送らない)
        ' 共有知フライホイール: 人が正しいと確認したQ&Aは組織の一次情報になる。
        modInsightIo.EmitVerifiedQA mLastQuestion, mLastCleanAnswer, LastTopSource()
    End If
    On Error GoTo 0

    ' 文面も判定に合わせる(共有していないのに共有したと言わない。憲章§3-3)。
    If mayEmit Then
        MsgBox "ありがとうございます。" & vbCrLf & _
               "この質問と回答は「解決済みQ&A」として部内に共有され、" & vbCrLf & _
               "同じことで困っている人がすぐ答えにたどり着けるようになります。", _
               vbInformation, modAppDef.APP_NAME
    Else
        MsgBox "ありがとうございます。解決できたことを記録しました。" & vbCrLf & _
               "(この回答は社内資料を根拠にしていないため、部内への共有は行いません。)", _
               vbInformation, modAppDef.APP_NAME
    End If
End Sub

Public Sub FeedbackYellow()
    If Not FeedbackAccepted() Then Exit Sub
    modStats.Bump "hint_total"
    modLog.LogUsage "feedback_yellow", mLastMode, "q=" & modUtil.SafeLeft(mLastQuestion, 200)
    MsgBox "ありがとうございます。次はもっと的確に答えられるよう活かします。", vbInformation, modAppDef.APP_NAME
End Sub

Public Sub FeedbackRed()
    If Not FeedbackAccepted() Then Exit Sub
    modStats.Bump "fail_total"
    modLog.LogUsage "feedback_red", mLastMode, "q=" & modUtil.SafeLeft(mLastQuestion, 200)
    ' 共有知フライホイール: 答えられなかった質問は「組織に文書が無い領域」の一次
    ' 情報。資料を書ける人の画面へ自動で流す。発信は緑と同じ門を通す(一般モードの
    ' 雑談まで「資料が足りない領域」として流さない)。
    On Error Resume Next
    If CanShareInsight() Then modInsightIo.EmitGap mLastQuestion, "wrong"
    On Error GoTo 0
    MsgBox "教えていただきありがとうございます。" & vbCrLf & _
           "この質問は「まだ答えを用意できていない質問」として記録し、" & vbCrLf & _
           "資料を作れる担当者の画面に届きます。", vbInformation, modAppDef.APP_NAME
End Sub

' 🤔 微妙(判断がつかない)。入力を一切求めず1クリックで終わる。
' 「正しいか分からないから何も押さない」を無くすための逃げ道。
Public Sub FeedbackUnsure()
    If Not FeedbackAccepted() Then Exit Sub
    modStats.Bump "unsure_total"
    modLog.LogUsage "feedback_unsure", mLastMode, "q=" & modUtil.SafeLeft(mLastQuestion, 200)
    On Error Resume Next
    If CanShareInsight() Then modInsightIo.EmitGap mLastQuestion, "low_conf"
    On Error GoTo 0
    MsgBox "ありがとうございます。" & vbCrLf & _
           "「判断がつかない」も立派な情報です。この質問は資料が不足している" & vbCrLf & _
           "可能性が高い領域として記録しました。", vbInformation, modAppDef.APP_NAME
End Sub

' 感想を記録してよい状態かの共通判定(多重カウント防止・回答前クリック防止)。
' Trueを返した時点で「記録済み」に倒す(呼び出し側は必ず記録する前提)。
Private Function FeedbackAccepted() As Boolean
    If LenB(mLastQuestion) = 0 Then
        MsgBox "まず質問して回答を受け取ってから、感想ボタンを押してください。", _
               vbInformation, modAppDef.APP_NAME
        Exit Function
    End If
    If mFeedbackDone Then
        MsgBox "この回答への感想はすでに記録されています。" & vbLf & _
               "(次の質問の回答から、また感想を送れます)", vbInformation, modAppDef.APP_NAME
        Exit Function
    End If
    mFeedbackDone = True
    FeedbackAccepted = True
End Function

' 内部ヘルパー(公開契約の一覧は vba_lint.py の CONTRACT が単一情報源)


' answer_tags時: <answer>抽出+タグ外FOLLOWUP救出+thinkingデバッグ記録。
' R14-8a でPublic(入念モードの各段も同じ規約で取り出すため)。
Public Function ApplyAnswerTags(ByVal resp As String) As String
    If Not modConfig.GetBool("answer_tags", True) Then
        ApplyAnswerTags = resp
        Exit Function
    End If

    Dim thinkingTxt As String, answerTxt As String
    modRagParse.ExtractAnswer resp, thinkingTxt, answerTxt

    ' FOLLOWUPは</answer>の外契約(§D-1)。answer側に無ければ元応答から救出。
    If InStr(answerTxt, "[[FOLLOWUP") = 0 Then
        Dim fp As Long
        fp = InStr(resp, "[[FOLLOWUP")
        If fp > 0 Then
            Dim fe As Long
            fe = InStr(fp, resp, "]]")
            If fe > 0 Then answerTxt = answerTxt & vbLf & Mid$(resp, fp, fe - fp + 2)
        End If
    End If

    If LenB(thinkingTxt) > 0 And modConfig.GetBool("debug_mode", False) Then
        On Error Resume Next
        modLog.LogUsage "thinking", "", modUtil.SafeLeft(thinkingTxt, 500)
        On Error GoTo 0
    End If

    ApplyAnswerTags = answerTxt
End Function

Private Function RunQuickFlow(ByVal q As String, hits() As Hit, ByVal nHits As Long, _
                              ByRef ok As Boolean, ByVal prevU As String, ByVal prevA As String) As String
    modAskRetrieve.ShowAskStage "quick"

    Dim prompt As String
    prompt = modPrompts.BuildQuickPrompt(q, hits, nHits, _
        modConfig.GetBool("strict_grounding", True), modConfig.GetBool("answer_tags", True))

    Dim eff As String: eff = modConfig.GetString("quick_effort", "low")
    Dim vrb As String: vrb = modConfig.GetString("quick_verbosity", "low")
    Dim mdl As String: mdl = modConfig.GetString("quick_model", "gpt-5.5")
    Dim latency As Long

    Dim resp As String
    resp = modGateway.CallLLM(prompt, "quick_draft", eff, vrb, mdl, latency, prevU, prevA)

    If modRagParse.IsErrorResponse(resp) Then
        ok = False
        RunQuickFlow = modRagParse.BuildErrorAnswer(resp)
    Else
        ok = True
        RunQuickFlow = DecorateWithFollowups(ApplyAnswerTags(resp))
    End If
End Function

' R34 F1: 精読は入念(modAskThorough:82-93)と同型。増やすのは nUse だけ。
Private Function RunDeepFlow(ByVal q As String, hits() As Hit, ByVal nHits As Long, _
                             ByRef ok As Boolean, ByVal prevU As String, ByVal prevA As String, _
                             ByVal mdMode As String) As String
    modAskRetrieve.ShowAskStage "draft"

    Dim strictG As Boolean: strictG = modConfig.GetBool("strict_grounding", True)
    Dim ansTags As Boolean: ansTags = modConfig.GetBool("answer_tags", True)

    Dim nUse As Long: nUse = nHits
    If modMode.UseNeighborExpand(mdMode) Then _
        modAskFocus.NeighborExpand hits, nUse, modConfig.GetLong("deep_neighbor", 2)
    mLastGenN = nUse

    Dim draftPrompt As String
    draftPrompt = modPrompts.BuildDeepDraftPrompt(q, hits, nUse, HistoryBlock(), strictG, ansTags)

    Dim dEff As String: dEff = modConfig.GetString("deep_draft_effort", "medium")
    Dim dVrb As String: dVrb = modConfig.GetString("deep_draft_verbosity", "high")
    Dim mdl As String: mdl = modConfig.GetString("recommended_model", "gpt-5.5")
    Dim latency As Long

    ' 会話履歴(prevU/prevA)は下書き段のみに渡す。検証段は本棚抜粋との照合に専念のため渡さない(裁定D11)。
    Dim draft As String
    draft = modGateway.CallLLM(draftPrompt, "deep_draft", dEff, dVrb, mdl, latency, prevU, prevA)

    If modRagParse.IsErrorResponse(draft) Then
        ok = False
        RunDeepFlow = modRagParse.BuildErrorAnswer(draft)
        Exit Function
    End If

    modAskRetrieve.ShowAskStage "verify"

    Dim draftBody As String
    draftBody = ApplyAnswerTags(draft)

    Dim verifyPrompt As String
    verifyPrompt = modPrompts.BuildDeepVerifyPrompt(q, draftBody, hits, nUse, strictG, ansTags)

    Dim vEff As String: vEff = modConfig.GetString("deep_verify_effort", "high")
    Dim vVrb As String: vVrb = modConfig.GetString("deep_verify_verbosity", "medium")

    Dim verified As String
    verified = modGateway.CallLLM(verifyPrompt, "deep_verify", vEff, vVrb, mdl, latency)

    ok = True
    If modRagParse.IsErrorResponse(verified) Then
        RunDeepFlow = DecorateWithFollowups(draftBody) & vbLf & vbLf & _
            "(注: 検証段階でエラーが発生したため、下書きの内容を表示しています。)"
    Else
        RunDeepFlow = DecorateWithFollowups(ApplyAnswerTags(verified))
    End If
End Function


Private Function NormalizeMode(ByVal mode As String) As String
    NormalizeMode = modMode.Normalize(mode)
End Function

Private Function TopKFor(ByVal mdMode As String) As Long
    TopKFor = modMode.TopK(mdMode, modConfig.GetLong("topk_quick", 6), _
                           modConfig.GetLong("topk_deep", 12), _
                           modConfig.GetLong("topk_thorough", 16))
End Function

' 定義済み名前(Excel Name)"mb_question"経由でホームの質問セルを読む
' (UI層連携アサンプション。上部の設計判断コメント参照)。未定義なら空文字列。
Private Function ReadQuestionFromHome() As String
    On Error Resume Next
    Dim v As Variant
    v = ThisWorkbook.Names(QUESTION_RANGE_NAME).RefersToRange.Value
    On Error GoTo 0
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    ReadQuestionFromHome = CStr(v)
End Function

' ui_stateシート(A=key, B=value)の key="mode" を読む("quick"/"deep"、既定quick)。
Private Function ReadModeFromUiState() As String
    ReadModeFromUiState = MODE_QUICK

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_UISTATE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(CStr(ws.Cells(i, 1).Value), UI_STATE_MODE_KEY, vbTextCompare) = 0 Then
            Dim v As String
            v = LCase$(Trim$(CStr(ws.Cells(i, 2).Value)))
            If v = MODE_DEEP Then ReadModeFromUiState = MODE_DEEP
            If v = MODE_THOROUGH Then ReadModeFromUiState = MODE_THOROUGH
            Exit Function
        End If
    Next i
End Function

' 会話履歴(直近HISTORY_MAX_TURNS往復)。
Private Sub AppendHistory(ByVal q As String, ByVal a As String)
    Dim turn As String
    turn = "Q: " & q & vbLf & "A: " & modUtil.SafeLeft(a, 1500)

    If LenB(mHistory) = 0 Then
        mHistory = turn
    Else
        mHistory = mHistory & HISTORY_SEP & turn
    End If

    Dim parts() As String
    parts = Split(mHistory, HISTORY_SEP)
    Dim n As Long
    n = UBound(parts) - LBound(parts) + 1
    If n > HISTORY_MAX_TURNS Then
        Dim keep As String
        Dim i As Long
        For i = UBound(parts) - HISTORY_MAX_TURNS + 1 To UBound(parts)
            If LenB(keep) > 0 Then keep = keep & HISTORY_SEP
            keep = keep & parts(i)
        Next i
        mHistory = keep
    End If
End Sub

' modAskRetrieve からも使う(多段検索の拡張プロンプトに直近履歴を載せる)。
Public Function HistoryBlock() As String
    If LenB(mHistory) = 0 Then Exit Function

    Dim parts() As String
    parts = Split(mHistory, HISTORY_SEP)
    Dim out As String
    Dim i As Long, n As Long
    For i = LBound(parts) To UBound(parts)
        n = n + 1
        out = out & "【会話" & n & "】" & vbLf & parts(i) & vbLf & vbLf
    Next i
    HistoryBlock = out
End Function

' [[FOLLOWUP:...]]を分離し「深掘り候補」ブロックを末尾に付けた表示用文字列を返す。
' 除去後の本文はmLastCleanAnswerへ。マーカー無し・形式崩れでも壊れない。
Private Function DecorateWithFollowups(ByVal resp As String) As String
    Dim body As String
    Dim cands As String
    modFollowup.SplitFollowupTrailer resp, body, cands
    mLastCleanAnswer = body

    If LenB(cands) = 0 Then
        DecorateWithFollowups = body
        Exit Function
    End If

    Dim disp As String
    disp = body & vbLf & vbLf & ChrW(&HD83D) & ChrW(&HDCA1) & " さらに深掘り(『深掘り』ボタンでそのまま聞けます):"
    Dim fl() As String
    fl = Split(cands, vbLf)
    Dim i As Long
    For i = LBound(fl) To UBound(fl)
        If LenB(Trim$(fl(i))) > 0 Then disp = disp & vbLf & "  ・" & fl(i)
    Next i
    DecorateWithFollowups = disp
End Function

' prevU/prevA用履歴(D11)。mHistory(プロンプト内履歴)とは別物。
Private Sub AppendFollowupPair(ByVal q As String, ByVal a As String)
    Dim maxPairs As Long
    maxPairs = modConfig.GetLong("followup_max_pairs", 3)
    If maxPairs <= 0 Then
        ' 0以下=履歴を持たない(CanFollowup=Falseになる)エスケープハッチ。
        ' 2026-07-28(レビュー L-2): メモリだけ消して ui_state を消していなかった
        ' ため、無効化したはずなのに保存済みの古い履歴が読み直され、LLMへ送られ
        ' 続けていた。保存側も空にする。
        mPrevU = ""
        mPrevA = ""
        modState.SaveState "nexus_ask_prevu", ""
        modState.SaveState "nexus_ask_preva", ""
        Exit Sub
    End If

    ' 2026-07-28(レビュー L-3): 書く前に必ず読む。VBAリセット等でモジュール変数が
    ' 消えた直後に新規質問が来ると、mPrevU が空のまま「先頭に積む」処理が走り、
    ' 保存済みの履歴を1ターンで上書き消去していた。CanFollowup と同じ遅延ロードを
    ' ここでも通す。
    If LenB(mPrevU) = 0 Then
        mPrevU = modState.LoadState("nexus_ask_prevu", "")
        mPrevA = modState.LoadState("nexus_ask_preva", "")
    End If

    ' 回答は1500字で打ち切る(AppendHistoryと同じ判断: 履歴でトークンを食い過ぎ
    ' ない)。";;;"はリボン側の履歴区切り文字のため、本文中に現れた場合は";;"へ
    ' 縮めて区切りの誤認を防ぐ(SanitizeForFollowupHistory)。
    Dim qs As String
    qs = modFollowup.SanitizeForFollowupHistory(q, FOLLOWUP_PAIR_SEP)
    Dim ans As String
    ans = modFollowup.SanitizeForFollowupHistory(modUtil.SafeLeft(a, 1500), FOLLOWUP_PAIR_SEP)

    ' 新しい順: 先頭に積む。
    If LenB(mPrevU) = 0 Then
        mPrevU = qs
    Else
        mPrevU = qs & FOLLOWUP_PAIR_SEP & mPrevU
    End If
    If LenB(mPrevA) = 0 Then
        mPrevA = ans
    Else
        mPrevA = ans & FOLLOWUP_PAIR_SEP & mPrevA
    End If

    mPrevU = modFollowup.KeepNewestPairs(mPrevU, maxPairs, FOLLOWUP_PAIR_SEP)
    mPrevA = modFollowup.KeepNewestPairs(mPrevA, maxPairs, FOLLOWUP_PAIR_SEP)
    modState.SaveState "nexus_ask_prevu", mPrevU
    modState.SaveState "nexus_ask_preva", mPrevA
End Sub

' SetPrevMemory - 会話メモリの直接受け渡し口(R28-W4)。SetGeneralMemory と
'   対称の単純代入。橋渡しが state へ書くだけだと mPrevU 非空時に CanFollowup
'   の遅延ロードが再読込せず、切替前の古い会話が深掘りに残る穴を塞ぐ。
Public Sub SetPrevMemory(ByVal u As String, ByVal a As String)
    mPrevU = u
    mPrevA = a
End Sub

' ResetPrevMemory - 会話クリア時の亡霊(HANDOFF M-4)対策。state だけ消しても
'   mPrevU が残ると深掘りに古い文脈が乗るため、モジュール変数も空へ戻す。
Public Sub ResetPrevMemory()
    mPrevU = ""
    mPrevA = ""
End Sub
