Attribute VB_Name = "modAsk"
Option Explicit

' ============================================================================
' modAsk - 2速QA(⚡すぐ聞く/🔍しっかり調べる)のオーケストレーション
' ----------------------------------------------------------------------------
' 役割:
'   ホームの質問+モードを読み(またはAnswer(question, mode)として直接
'   呼ばれ)、modRetrieve.Searchで検索→modPrompts.Build*Promptで組み立て→
'   modGateway.CallLLMで回答生成、という一連の流れを実行し、進捗を
'   modUIMain経由で実況する(MASTER_SPEC §7.3)。
'
' 設計判断:
'   ・quick: SetStage(検索中)→Search(topk_quick)→出典先出し表示→
'     SetStage(回答作成中)→CallLLM(quick_draft)の2段。
'     deep : Search(topk_deep)→出典先出し→draft(deep_draft_*)→
'     SetStage(検証中)→verify(deep_verify_*)の3段。
'   ・UI連携はSetStage/RenderSourcesPreview/RenderAnswerの3本のみ
'     (§7.3/§7.6準拠。R1例外はWave3で整合済み)。
'   ・Answer(question, mode)は最終回答テキストのみを返す契約(§7.3)なので、
'     直近の検索結果(hits/nHits)・所要秒・実際に使ったモードは
'     モジュール変数(mLast*)に保持し、同一モジュール内のAskFromUIが
'     それを読んでmodUIMain.RenderAnswerに渡す(RenderAnswerの二重呼び
'     出しを避けるため、Answer自身はRenderAnswerを呼ばない)。
'   ・質問セルは定義済み名前"mb_question"経由で読む(未定義=空扱い)。
'   ・モードは ui_state シート(A=key, B=value)の key="mode" を読む
'     ("quick"/"deep"、既定quick。指示どおり)。
'   ・検索0件はLLMを呼ばず定型文+資料追加の案内を返す(E0601はログのみ、
'     ShowErrorのようなMsgBoxポップアップは出さない)。埋め込み失敗
'     (Search=-1)はE0203の案内文を返す(実際のLogErrorはmodGateway.
'     GetEmbedding内で完了済みなのでここでは二重に記録しない)。
'   ・CallLLMが返す"#ERR:E02xx:..."は、コード部分を取り出してmodLog.
'     FriendlyMessageの文面へ変換して表示する(生の#ERR文字列をそのまま
'     ユーザーに見せない)。
'   ・質問は3000字で打ち切り(modUtil.SafeLeft)、空質問はLLMを呼ばずに
'     案内文を返す。連打防止は module変数のBooleanフラグ(mAsking)で
'     AskFromUI側にガードする(Answer自体は直接テスト呼び出しされる
'     ことも想定し、ガードしない)。
'   ・ESC対応: Application.EnableCancelKey = xlErrorHandler を設定し、
'     Err.Number=18(Ctrl+Break/中断)を検知したら通常のE0602ではなく
'     「操作を中断しました」という案内にする(modEmbed.EmbedPendingと
'     同じ考え方: MASTER_SPEC §7.2)。
'   ・会話履歴: 直近3往復をmodule変数(mHistory)に保持する、V2
'     src/chatbot_v2/modBoot.bas の HistoryBlock/AppendHistory の
'     簡易版をmodAsk内に内蔵したもの(V2から流用可: MASTER_SPEC発注時の
'     指示どおり)。quick/deep問わず、実際にLLMから回答が得られた
'     (エラーでない)ターンのみ履歴に積む。
'   ・続けて質問(裁定D11): 成功した各ターンのQ&Aを、上記mHistory
'     (プロンプト内履歴)とは別に、リボンChatGPT()の確定引数prevU/prevA
'     (台帳§1 #1 第7・8引数)へそのまま渡せる形式=「新しい順;;;区切り」の
'     mPrevU/mPrevAとしてもセッション保持する(最大 config
'     followup_max_pairs 既定3ペア)。AskFollowupは既存のAnswer系フローを
'     そのまま再利用し、追質問文でmodRetrieve.Searchも再実行したうえで、
'     この履歴をmodGateway.CallLLMのprevU/prevAに添えて出典付き回答を
'     返す(2速モードは既存どおりui_stateの設定に従う)。
'   ・深掘り候補(裁定D11): LLM応答末尾の [[FOLLOWUP: 候補1 | 候補2]] を
'     パースして本文から除去し、「深掘り候補(『続けて質問』でそのまま
'     聞けます)」ブロックとして本文末尾に整形追記してからRenderAnswerへ
'     渡す(RenderAnswerの契約は不変)。パースはV2実証済みの
'     modPipeline.ParseTrailersと同じ流儀の寛容実装で、マーカーが無い応答
'     (mockLLM応答等)・候補なし・形式崩れでも壊れない(その場合は候補
'     ブロックなしで正常動作)。履歴(mHistory/mPrevU/mPrevA)には候補
'     ブロックを含まない除去後の本文だけを積む。
' ============================================================================

Private Const MODE_QUICK As String = "quick"
Private Const MODE_DEEP As String = "deep"
Private Const MAX_QUESTION_CHARS As Long = 3000
Private Const QUESTION_RANGE_NAME As String = "mb_question"
Private Const UI_STATE_MODE_KEY As String = "mode"
Private Const EMPTY_QUESTION_MESSAGE As String = _
    "質問が入力されていません。ホームの入力欄に質問を書いてから、もう一度お試しください。"

Private Const HISTORY_MAX_TURNS As Long = 3
Private Const HISTORY_SEP As String = "<<<__QA_TURN__>>>"

' 続けて質問(裁定D11): リボンChatGPT()のprevU/prevA引数の履歴区切り文字
' (確定: 新しい順;;;区切り。台帳§1 #1+裁定D11)。
Private Const FOLLOWUP_PAIR_SEP As String = ";;;"

Private mAsking As Boolean
Private mHistory As String

' 続けて質問(裁定D11)用のセッション履歴: prevU=過去の質問/prevA=過去の
' 回答(どちらも新しい順;;;区切り・最大 config followup_max_pairs ペア)。
' mLastCleanAnswerは深掘り候補ブロック除去後の直近回答本文(履歴保存用)。
Private mPrevU As String
Private mPrevA As String
Private mLastCleanAnswer As String

Private mLastQuestion As String
Private mLastAnswer As String
Private mLastMode As String
' フィードバックの多重カウント防止(2026-07-16): 🟢🟡🔴ボタンは何度でも
' 押せてしまい、押すたびにselfsolve_total等が加算されて「取り戻した時間」も
' 実態とズレていた。1回の回答につき感想は1回だけ記録する。
Private mFeedbackDone As Boolean
Private mLastHits() As Hit
Private mLastNHits As Long
Private mLastSeconds As Long

' ----------------------------------------------------------------------------
' AskFromUI - ホームの質問セル+モード(ui_state)を読み、Answer実行→
'   RenderAnswer(MASTER_SPEC §7.3)。連打防止ガード付き。
' ----------------------------------------------------------------------------
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

' ----------------------------------------------------------------------------
' Answer - 質問文とモードから回答テキストを組み立てて返す(MASTER_SPEC §7.3)。
'   契約どおりの公開API。会話履歴なしの単発質問として実行する(実体は
'   AnswerWithContext。裁定D11のAskFollowupと本体を共有する)。
' ----------------------------------------------------------------------------
Public Function Answer(ByVal question As String, ByVal mode As String) As String
    Answer = AnswerWithContext(question, mode, "", "", False)
End Function

' ----------------------------------------------------------------------------
' CanFollowup - 「続けて質問」できる直近回答が存在するか(裁定D11)。
'   このセッションで成功した回答が1件でもあればTrue。UI側(modUIMainの
'   OnFollowupButton)がFalse時に「まず質問してから」の丁寧な案内を出す。
'   config followup_max_pairs を0以下にすると履歴を持たなくなるため、
'   常にFalse(=機能無効)になるエスケープハッチを兼ねる。
' ----------------------------------------------------------------------------
Public Function CanFollowup() As Boolean
    CanFollowup = (LenB(mPrevU) > 0)
End Function

' ----------------------------------------------------------------------------
' AskFollowup - 直近の会話履歴を添えて追質問を実行する(裁定D11)。
'   既存のAnswer系フローを再利用: 追質問文でmodRetrieve.Searchも再実行し、
'   出典付き回答をRenderAnswerで表示する(2速モードは既存どおりui_stateの
'   設定に従う)。履歴はmPrevU/mPrevA(新しい順;;;区切り)をCallLLMの
'   prevU/prevA(台帳§1 #1 第7・8引数)へそのまま渡す。履歴が空
'   (CanFollowup=False)のまま呼ばれた場合は通常の単発質問と同じ動作に
'   自然に退化する(UI側が事前案内する契約だが、直接呼ばれても壊れない防御)。
' ----------------------------------------------------------------------------
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

' ----------------------------------------------------------------------------
' AnswerWithContext - Answer/AskFollowup共通の回答生成本体(Private)。
'   prevU/prevAはリボンChatGPT()へ渡す会話履歴(空文字=履歴なし)。
'   isFollowupはusage_logの識別用(event="ask"のdetail先頭に"followup "を付す)。
' ----------------------------------------------------------------------------
Private Function AnswerWithContext(ByVal question As String, ByVal mode As String, _
                                   ByVal prevU As String, ByVal prevA As String, _
                                   ByVal isFollowup As Boolean) As String
    Dim tStart As Double
    tStart = Timer

    Dim q As String
    q = modUtil.SafeLeft(Trim$(question), MAX_QUESTION_CHARS)
    Dim mdMode As String
    mdMode = NormalizeMode(mode)

    If LenB(q) = 0 Then
        ' Wave4修正: 以前はここでmLast*を更新せずExit Functionしていたため、
        ' 直前の(本物の)質問のヒット/所要秒がmodUIMain.RenderAnswerの
        ' footer・出典欄に残り、まるで空クリックがその資料で回答したかの
        ' ように見える不具合があった。mLastMode=""は「検索・回答生成を
        ' 一切行わなかった」ことを示す目印としてmodUIMain.RenderAnswer側
        ' でも使う(footer/出典欄を出さない判定)。
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

    On Error Resume Next
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo 0

    On Error GoTo Fail

    modUIMain.SetStage "🔍 検索中…"
    Dim topK As Long
    topK = TopKFor(mdMode)

    ' 多段RAG(§C)。retrieve_mode=singleで従来の単段Searchへ完全退化。
    If LCase$(modConfig.GetString("retrieve_mode", "single")) = "multi" Then
        nHits = RunMultiRetrieve(q, mdMode, topK, hits)
    Else
        nHits = modRetrieve.Search(q, topK, hits)
    End If

    If nHits = -1 Then
        result = modLog.FriendlyMessage("E0203") & vbLf & "(コード: E0203)"
    ElseIf nHits = 0 Then
        modLog.LogError "E0601", "modAsk.Answer", "query=" & modUtil.SafeLeft(q, 200)
        result = modLog.FriendlyMessage("E0601") & vbLf & vbLf & _
            "『マイ本棚』タブから資料を追加すると、次から答えられるようになります。"
    Else
        modUIMain.SetStage "📄 " & nHits & "件の資料がヒット"
        modUIMain.RenderSourcesPreview hits, nHits

        If mdMode = MODE_DEEP Then
            result = RunDeepFlow(q, hits, nHits, ok, prevU, prevA)
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
    Err.Clear
    On Error GoTo 0

    If errNum = 18 Then
        result = "操作を中断しました。もう一度質問するときは、質問するボタンを押してください。"
    Else
        ' 出所ラベルは公開APIのAnswerで記録する(内部関数名を "modAsk.Xxx" 形式の
        ' 文字列で書くとlintのモジュール間参照検査が実在Publicと照合して誤検知する
        ' ため。詳細detailに共通本体である旨を残す)。
        modLog.LogError "E0602", "modAsk.Answer", "共通本体(AnswerWithContext) mode=" & mdMode & " err=" & errDesc
        result = modLog.FriendlyMessage("E0602") & vbLf & "(コード: E0602)"
    End If

    On Error Resume Next
    modUIMain.SetStage ""
    On Error GoTo 0

Done:
    Dim elapsedMs As Long
    elapsedMs = CLng((Timer - tStart) * 1000)
    Dim elapsedSec As Long
    elapsedSec = CLng(Timer - tStart)

    ' 成功ターンのみ履歴に積む。積むのは深掘り候補ブロックを含まない
    ' 除去後の本文(mLastCleanAnswer。DecorateWithFollowupsが設定)。
    If ok Then
        AppendHistory q, mLastCleanAnswer
        AppendFollowupPair q, mLastCleanAnswer
    End If

    mLastQuestion = q
    mLastAnswer = result
    mLastMode = mdMode
    mLastHits = hits
    mLastNHits = nHits
    mLastSeconds = elapsedSec
    mFeedbackDone = False   ' 新しい回答に対する感想を受付可能にする

    Dim logDetail As String
    logDetail = "q=" & modUtil.SafeLeft(q, 200)
    If isFollowup Then logDetail = "followup " & logDetail
    modLog.LogUsage "ask", mdMode, logDetail, elapsedMs, nHits
    modStats.Bump "ask_" & mdMode & "_total"

    AnswerWithContext = result
End Function

Public Sub FeedbackGreen()
    If Not FeedbackAccepted() Then Exit Sub
    modStats.Bump "selfsolve_total"
    modLog.LogUsage "feedback_green", mLastMode, "q=" & modUtil.SafeLeft(mLastQuestion, 200)
    MsgBox "ありがとうございます。解決に役立てて何よりです。", vbInformation, modAppDef.APP_NAME
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
    MsgBox "ご意見ありがとうございます。改善の参考にします。", vbInformation, modAppDef.APP_NAME
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

' ----------------------------------------------------------------------------
' 内部ヘルパー(すべてPrivate: modAskの公開契約は上記7本のみ)
' ----------------------------------------------------------------------------

' 多段RAG検索段(§C): 拡張→マルチクエリ→再ランク。全段とも失敗時は
' 単段Search(q)へ安全退化(mock/タグ欠落でも壊れない)。
Private Function RunMultiRetrieve(ByVal q As String, ByVal mdMode As String, _
                                  ByVal topK As Long, ByRef hits() As Hit) As Long
    On Error GoTo FallbackSingle

    ' 1) クエリ拡張
    Dim queries() As String
    Dim queryN As Long: queryN = 0
    ReDim queries(0 To 8)

    Dim standalone As String, hyde As String
    Dim subs() As String
    standalone = "": hyde = ""

    If modConfig.GetBool("expand_enabled", False) Then
        modUIMain.SetStage "🧭 質問を分析中…"
        Dim lightMode As Boolean
        lightMode = (mdMode <> MODE_DEEP) And modConfig.GetBool("quick_expand_light", True)

        Dim exPrompt As String
        exPrompt = modPrompts.BuildExpandPrompt(q, HistoryBlock(), _
            modConfig.GetLong("expand_subqueries", 3), lightMode)

        Dim exModel As String: exModel = modConfig.GetString("expand_model", "")
        If LenB(exModel) = 0 Then exModel = modConfig.GetString("quick_model", "gpt-5.5")
        Dim exLat As Long
        Dim exResp As String
        exResp = modGateway.CallLLM(exPrompt, "expand", _
            modConfig.GetString("expand_effort", "low"), _
            modConfig.GetString("expand_verbosity", "low"), exModel, exLat)

        If Not IsErrorResponse(exResp) Then
            modRagParse.ParseExpand exResp, standalone, subs, hyde
        End If
    End If

    If LenB(Trim$(standalone)) = 0 Then standalone = q
    queries(queryN) = standalone: queryN = queryN + 1

    Dim si As Long
    On Error Resume Next
    For si = LBound(subs) To UBound(subs)
        If queryN <= 7 And LenB(Trim$(subs(si))) > 0 Then
            queries(queryN) = subs(si): queryN = queryN + 1
        End If
    Next si
    On Error GoTo FallbackSingle
    If LenB(Trim$(hyde)) > 0 And queryN <= 8 Then
        queries(queryN) = hyde: queryN = queryN + 1
    End If
    ReDim Preserve queries(0 To queryN - 1)

    ' 2) マルチクエリ検索(候補プール)
    Dim poolK As Long: poolK = modConfig.GetLong("multi_candidates", 40)
    If poolK < topK Then poolK = topK
    Dim poolHits() As Hit
    Dim poolN As Long
    poolN = modRetrieve.SearchExpanded(queries, poolK, poolHits)
    If poolN <= 0 Then GoTo FallbackSingle

    ' 3) 再ランク(候補がtopKより多いときだけ意味がある)
    Dim orderN As Long: orderN = 0
    Dim rankOrder() As Long
    If modConfig.GetBool("rerank_enabled", False) And poolN > topK Then
        modUIMain.SetStage "🧮 関連度を精査中…"
        Dim rkPrompt As String
        rkPrompt = modPrompts.BuildRerankPrompt(q, poolHits, poolN, _
            modConfig.GetLong("max_context_chars", 40000))
        Dim rkModel As String: rkModel = modConfig.GetString("rerank_model", "")
        If LenB(rkModel) = 0 Then rkModel = modConfig.GetString("quick_model", "gpt-5.5")
        Dim rkLat As Long
        Dim rkResp As String
        rkResp = modGateway.CallLLM(rkPrompt, "rerank", _
            modConfig.GetString("rerank_effort", "low"), _
            modConfig.GetString("rerank_verbosity", "low"), rkModel, rkLat)
        If Not IsErrorResponse(rkResp) Then
            orderN = modRagParse.ParseRankOrder(rkResp, poolN, rankOrder)
        End If
    End If

    ' 4) 最終topKへ絞り込み(再ランク順 or スコア順)
    Dim outN As Long
    If orderN > 0 Then
        outN = orderN
    Else
        outN = poolN
    End If
    If outN > topK Then outN = topK

    ReDim hits(1 To outN)
    Dim oi As Long
    For oi = 1 To outN
        If orderN > 0 Then
            hits(oi) = poolHits(rankOrder(oi - 1))
        Else
            hits(oi) = poolHits(oi)
        End If
    Next oi
    RunMultiRetrieve = outN
    Exit Function

FallbackSingle:
    Err.Clear
    On Error GoTo 0
    RunMultiRetrieve = modRetrieve.Search(q, topK, hits)
End Function

' answer_tags時: <answer>抽出+タグ外FOLLOWUP救出+thinkingデバッグ記録。
Private Function ApplyAnswerTags(ByVal resp As String) As String
    If Not modConfig.GetBool("answer_tags", False) Then
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
    modUIMain.SetStage "✍️ 回答作成中…"

    Dim prompt As String
    prompt = modPrompts.BuildQuickPrompt(q, hits, nHits, _
        modConfig.GetBool("strict_grounding", False), modConfig.GetBool("answer_tags", False))

    Dim eff As String: eff = modConfig.GetString("quick_effort", "low")
    Dim vrb As String: vrb = modConfig.GetString("quick_verbosity", "low")
    Dim mdl As String: mdl = modConfig.GetString("quick_model", "gpt-5.5")
    Dim latency As Long

    Dim resp As String
    resp = modGateway.CallLLM(prompt, "quick_draft", eff, vrb, mdl, latency, prevU, prevA)

    If IsErrorResponse(resp) Then
        ok = False
        RunQuickFlow = BuildErrorAnswer(resp)
    Else
        ok = True
        RunQuickFlow = DecorateWithFollowups(ApplyAnswerTags(resp))
    End If
End Function

Private Function RunDeepFlow(ByVal q As String, hits() As Hit, ByVal nHits As Long, _
                             ByRef ok As Boolean, ByVal prevU As String, ByVal prevA As String) As String
    modUIMain.SetStage "✍️ 回答を下書き中…"

    Dim strictG As Boolean: strictG = modConfig.GetBool("strict_grounding", False)
    Dim ansTags As Boolean: ansTags = modConfig.GetBool("answer_tags", False)

    Dim draftPrompt As String
    draftPrompt = modPrompts.BuildDeepDraftPrompt(q, hits, nHits, HistoryBlock(), strictG, ansTags)

    Dim dEff As String: dEff = modConfig.GetString("deep_draft_effort", "medium")
    Dim dVrb As String: dVrb = modConfig.GetString("deep_draft_verbosity", "high")
    Dim mdl As String: mdl = modConfig.GetString("recommended_model", "gpt-5.5")
    Dim latency As Long

    ' 会話履歴(prevU/prevA)は、追質問の意図解釈を担う下書き段のみに渡す。
    ' 検証段は「下書きと本棚抜粋の照合」に専念する役割のため渡さない(裁定D11
    ' の適用判断。過去の会話が検証の判断材料に混ざるのを避ける)。
    Dim draft As String
    draft = modGateway.CallLLM(draftPrompt, "deep_draft", dEff, dVrb, mdl, latency, prevU, prevA)

    If IsErrorResponse(draft) Then
        ok = False
        RunDeepFlow = BuildErrorAnswer(draft)
        Exit Function
    End If

    modUIMain.SetStage "✅ 検証中…"

    Dim draftBody As String
    draftBody = ApplyAnswerTags(draft)

    Dim verifyPrompt As String
    verifyPrompt = modPrompts.BuildDeepVerifyPrompt(q, draftBody, hits, nHits, strictG, ansTags)

    Dim vEff As String: vEff = modConfig.GetString("deep_verify_effort", "high")
    Dim vVrb As String: vVrb = modConfig.GetString("deep_verify_verbosity", "medium")

    Dim verified As String
    verified = modGateway.CallLLM(verifyPrompt, "deep_verify", vEff, vVrb, mdl, latency)

    ok = True
    If IsErrorResponse(verified) Then
        RunDeepFlow = DecorateWithFollowups(draftBody) & vbLf & vbLf & _
            "(注: 検証段階でエラーが発生したため、下書きの内容を表示しています。)"
    Else
        RunDeepFlow = DecorateWithFollowups(ApplyAnswerTags(verified))
    End If
End Function

Private Function NormalizeMode(ByVal mode As String) As String
    If LCase$(Trim$(mode)) = MODE_DEEP Then
        NormalizeMode = MODE_DEEP
    Else
        NormalizeMode = MODE_QUICK
    End If
End Function

Private Function TopKFor(ByVal mdMode As String) As Long
    If mdMode = MODE_DEEP Then
        TopKFor = modConfig.GetLong("topk_deep", 12)
    Else
        TopKFor = modConfig.GetLong("topk_quick", 6)
    End If
End Function

Private Function IsErrorResponse(ByVal s As String) As Boolean
    IsErrorResponse = (Left$(s, 5) = "#ERR:")
End Function

Private Function BuildErrorAnswer(ByVal errResp As String) As String
    Dim code As String
    code = ExtractErrorCode(errResp)
    If LenB(code) = 0 Then code = "E0202"
    BuildErrorAnswer = modLog.FriendlyMessage(code) & vbLf & "(コード: " & code & ")"
End Function

' "#ERR:E0202:説明..." -> "E0202"
Private Function ExtractErrorCode(ByVal s As String) As String
    If Left$(s, 5) <> "#ERR:" Then Exit Function
    Dim rest As String
    rest = Mid$(s, 6)
    Dim p As Long
    p = InStr(rest, ":")
    If p = 0 Then
        ExtractErrorCode = rest
    Else
        ExtractErrorCode = Left$(rest, p - 1)
    End If
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
            Exit Function
        End If
    Next i
End Function

' ----------------------------------------------------------------------------
' 会話履歴(直近HISTORY_MAX_TURNS往復): V2 modBoot.HistoryBlock/AppendHistory の
' 簡易版をmodAsk内に内蔵したもの。
' ----------------------------------------------------------------------------
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

Private Function HistoryBlock() As String
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

' ----------------------------------------------------------------------------
' 深掘り候補(裁定D11): [[FOLLOWUP: 候補1 | 候補2]] のパースと表示整形。
' V2実証済みの src/chatbot_v2/modPipeline.bas ParseTrailers(FOLLOWUP部)と
' modChatUI.bas の候補表示を、本モジュール用に移植したもの。
' ----------------------------------------------------------------------------

' LLM応答からマーカーを分離し、候補があれば「深掘り候補」ブロックを本文
' 末尾に整形追記した表示用文字列を返す(RenderAnswerへはこの戻り値が渡る)。
' マーカー除去後の本文(履歴保存用)はmLastCleanAnswerに保持する。
' マーカーなし・候補なし・形式崩れでも壊れない(候補ブロックなしで本文を
' そのまま返すだけ。mockLLM応答にマーカーが無い場合もこの経路で正常動作)。
Private Function DecorateWithFollowups(ByVal resp As String) As String
    Dim body As String
    Dim cands As String
    SplitFollowupTrailer resp, body, cands
    mLastCleanAnswer = body

    If LenB(cands) = 0 Then
        DecorateWithFollowups = body
        Exit Function
    End If

    Dim disp As String
    disp = body & vbLf & vbLf & "🔎 深掘り候補(『続けて質問』でそのまま聞けます):"
    Dim fl() As String
    fl = Split(cands, vbLf)
    Dim i As Long
    For i = LBound(fl) To UBound(fl)
        If LenB(Trim$(fl(i))) > 0 Then disp = disp & vbLf & "  ・" & fl(i)
    Next i
    DecorateWithFollowups = disp
End Function

' 応答aから [[FOLLOWUP: ...]] をパースし、body=マーカー行除去後の本文 /
' candidates=候補(vbLf区切り。無ければ空)に分離する。V2と同じ寛容実装:
' "]]"が見つからない・中身が空・「なし」の場合は候補なし扱いとし、
' マーカーを含む行だけを本文から取り除く(パース失敗でも例外は出さない)。
Private Sub SplitFollowupTrailer(ByVal a As String, ByRef body As String, ByRef candidates As String)
    body = a
    candidates = ""
    If LenB(a) = 0 Then Exit Sub

    Dim fp As Long
    fp = InStr(1, a, "[[FOLLOWUP:", vbTextCompare)
    If fp = 0 Then Exit Sub

    Dim fq As Long
    fq = InStr(fp, a, "]]")
    If fq > 0 Then
        Dim fv As String
        fv = Trim$(Mid$(a, fp + Len("[[FOLLOWUP:"), fq - fp - Len("[[FOLLOWUP:")))
        If StrComp(fv, "なし", vbTextCompare) <> 0 And LenB(fv) > 0 Then
            Dim parts() As String
            parts = Split(fv, "|")
            Dim out As String
            Dim i As Long
            For i = LBound(parts) To UBound(parts)
                Dim t As String
                t = Trim$(parts(i))
                If LenB(t) > 0 Then
                    If LenB(out) > 0 Then out = out & vbLf
                    out = out & t
                End If
            Next i
            candidates = out
        End If
    End If

    ' マーカーを含む行を本文から除去し、末尾の空行・空白を刈り込む
    ' (V2 modPipeline.ParseTrailersと同じ流儀)。
    Dim lines() As String
    lines = Split(a, vbLf)
    Dim keep As String
    Dim j As Long
    For j = LBound(lines) To UBound(lines)
        If InStr(lines(j), "[[FOLLOWUP:") = 0 Then
            If LenB(keep) > 0 Then keep = keep & vbLf
            keep = keep & lines(j)
        End If
    Next j
    Do While Len(keep) > 0 And (Right$(keep, 1) = vbLf Or Right$(keep, 1) = vbCr Or Right$(keep, 1) = " ")
        keep = Left$(keep, Len(keep) - 1)
    Loop
    body = keep
End Sub

' ----------------------------------------------------------------------------
' prevU/prevA用履歴(裁定D11): 成功した各ターンのQ&Aを「新しい順;;;区切り」で
' セッション保持する(最大 config followup_max_pairs 既定3ペア)。
' mHistory(プロンプト内履歴)とは別物: こちらはリボンChatGPT()の確定引数
' prevU/prevA(台帳§1 #1 第7・8引数)へそのまま渡すための形式。
' ----------------------------------------------------------------------------
Private Sub AppendFollowupPair(ByVal q As String, ByVal a As String)
    Dim maxPairs As Long
    maxPairs = modConfig.GetLong("followup_max_pairs", 3)
    If maxPairs <= 0 Then
        ' 0以下=履歴を持たない(CanFollowup=Falseになる)エスケープハッチ。
        mPrevU = ""
        mPrevA = ""
        Exit Sub
    End If

    ' 回答は1500字で打ち切る(AppendHistoryと同じ判断: 履歴でトークンを
    ' 食い過ぎない)。";;;"はリボン側の履歴区切り文字のため、本文中に現れた
    ' 場合は";;"へ縮めて区切りの誤認を防ぐ(SanitizeForFollowupHistory)。
    Dim qs As String
    qs = SanitizeForFollowupHistory(q)
    Dim ans As String
    ans = SanitizeForFollowupHistory(modUtil.SafeLeft(a, 1500))

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

    mPrevU = KeepNewestPairs(mPrevU, maxPairs)
    mPrevA = KeepNewestPairs(mPrevA, maxPairs)
End Sub

' 「新しい順;;;区切り」文字列の先頭からmaxPairs件だけを残す。
Private Function KeepNewestPairs(ByVal joined As String, ByVal maxPairs As Long) As String
    Dim parts() As String
    parts = Split(joined, FOLLOWUP_PAIR_SEP)
    Dim n As Long
    n = UBound(parts) - LBound(parts) + 1
    If n <= maxPairs Then
        KeepNewestPairs = joined
        Exit Function
    End If

    Dim out As String
    Dim i As Long
    For i = LBound(parts) To LBound(parts) + maxPairs - 1
        If LenB(out) > 0 Then out = out & FOLLOWUP_PAIR_SEP
        out = out & parts(i)
    Next i
    KeepNewestPairs = out
End Function

' 履歴に積む文字列から区切り文字";;;"を除去する。単純な1回のReplaceでは
' ";;;;;;"→";;;;"のように置換結果へ再び";;;"が現れ得るため、無くなるまで
' 繰り返す(各回で必ず短くなるので有限回で終わる)。
Private Function SanitizeForFollowupHistory(ByVal s As String) As String
    Dim t As String
    t = s
    Do While InStr(t, FOLLOWUP_PAIR_SEP) > 0
        t = Replace(t, FOLLOWUP_PAIR_SEP, ";;")
    Loop
    SanitizeForFollowupHistory = t
End Function
