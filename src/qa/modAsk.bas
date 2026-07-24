Attribute VB_Name = "modAsk"
Option Explicit

' ============================================================================
' modAsk - 2速QA(⚡すぐ聞く/🔍しっかり調べる)のオーケストレーション
' ----------------------------------------------------------------------------
' 役割: ホームの質問+モード(またはAnswer(question, mode)の直接呼び出し)を
'   modRetrieve.Search→modPrompts.Build*Prompt→modGateway.CallLLMの順に処理し、
'   進捗をmodUIMain経由で実況する(MASTER_SPEC §7.3)。
'
' 設計判断:
'   ・quick: Search(topk_quick)→出典先出し→CallLLM(quick_draft)の2段。
'     deep : Search(topk_deep)→出典先出し→draft→検証(verify)の3段。
'   ・UI連携はSetStage/RenderSourcesPreview/RenderAnswerの3本のみ。
'   ・Answer(question, mode)は最終回答テキストのみを返す契約(§7.3)。
'     hits/nHits・所要秒・実モードはmLast*に保持し、AskFromUIがRenderAnswerへ渡す。
'   ・質問セルは定義済み名前"mb_question"、モードはui_state(key="mode")から読む。
'   ・検索0件はLLM未呼び出しで定型文+資料追加案内(E0601はログのみ)。
'     埋め込み失敗(Search=-1)はE0203案内(LogErrorはGetEmbedding内で完了済み)。
'   ・CallLLMの"#ERR:E02xx:..."はコードを抽出しmodLog.FriendlyMessageへ変換。
'   ・質問は3000字打ち切り、空質問は案内のみ。連打防止はmAskingでAskFromUI側がガード。
'   ・ESC: EnableCancelKey=xlErrorHandler、Err.Number=18は「操作を中断しました」。
'   ・会話履歴: 直近3往復をmHistoryに保持(V2 modBoot流用)。成功ターンのみ積む。
'   ・続けて質問(裁定D11): 成功ターンのQ&Aをリボン確定引数prevU/prevA形式
'     (新しい順;;;区切り)でmPrevU/mPrevAにも保持(最大followup_max_pairsペア)。
'     AskFollowupはAnswer系フローを再利用しSearchも再実行、prevU/prevAを添えて返す。
'     mPrevU/mPrevAはVBAリセットに備えmodStateでui_state退避・復元する。
'   ・深掘り候補(裁定D11): 応答末尾[[FOLLOWUP: 候補1 | 候補2]]をパースして本文から
'     除去し、「深掘り候補」ブロックとして末尾に整形追記(V2 ParseTrailers流儀の
'     寛容実装、マーカーなしでも壊れない)。履歴には除去後の本文のみ積む。
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
    If LenB(mPrevU) = 0 Then
        mPrevU = modState.LoadState("nexus_ask_prevu", "")
        mPrevA = modState.LoadState("nexus_ask_preva", "")
    End If
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

    modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDD0D) & " 検索中…"
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
        modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDCC4) & " " & nHits & "件の資料がヒット"
        modUIMain.RenderSourcesPreview hits, nHits

        ' 曖昧クエリ検知: 質問が短く、かつトップヒットの関連度が低い場合
        ' LLMを呼ばずに確認プロンプトを返す(API節約+精度向上)
        If Len(q) <= 10 And nHits >= 1 Then
            If hits(1).score < 0.72 Then
                result = ChrW(&HD83D) & ChrW(&HDCAD) & " もう少し詳しく教えていただけますか？" & vbLf & vbLf & _
                    "たとえば:" & vbLf & _
                    "・どの資料について？（約款 / マニュアル / 規程）" & vbLf & _
                    "・どんな状況で？（契約者対応 / 社内手続き / 研修）" & vbLf & _
                    "・知りたい結論は？（必要書類 / 所要日数 / 保険料への影響）" & vbLf & vbLf & _
                    "具体的に書くほど、精度の高い回答が得られます。"
                modLog.LogUsage "ambiguous_clarify", "", "q=" & modUtil.SafeLeft(q, 50) & " score=" & Format$(hits(1).score, "0.000")
                GoTo Done
            End If
        End If

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
        ' 出所ラベルは公開APIのAnswerで記録(内部関数名を文字列で書くとlintの参照検査が誤検知するため)。
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

    ' 低関連度警告(表示専用): 履歴(AppendHistory/mLastCleanAnswer)は上で
    ' 既に確定済みのため、ここでresultに警告を足しても履歴側には混入しない。
    If ok And nHits > 0 Then
        result = ApplyLowHitWarning(result, hits, nHits)
    End If

    ' チャット履歴シート記録(modChatLog、core層。書込失敗で死なない設計)。
    If LenB(q) > 0 Then
        On Error Resume Next
        modChatLog.LogTurn q, result, mdMode
        On Error GoTo 0
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
    modStats.AddExp "question"

    AnswerWithContext = result
End Function

' 直近回答の本文(整形済み)を返す(Nexus UI等の外部表示用ゲッター)。
Public Function LastAnswerText() As String
    LastAnswerText = mLastAnswer
End Function

' 直近回答で最上位スコアのソース名(P2P感謝状の宛先解決用)。
Public Function LastTopSource() As String
    ' 【2026-07-20修正】mLastHitsは ReDim(1 To n)(modRetrieve/RunMultiRetrieve共通)
    ' なのに、ここだけ 0始まりで走査していたため初回の mLastHits(0) で
    ' 添字エラー(9)→呼び出し元EmitThanksForLastAnswerのOn Error Resume Nextに
    ' 握りつぶされ、感謝状が一度も発行されない実バグだった。1始まりに修正。
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

' ----------------------------------------------------------------------------
' Peek View(出典ポップアップ)用の読み取り専用アクセサ。直近回答が根拠にした
' 出典(source/page/origin/本文)をUI層へ公開する。添字は0始まり(0..LastHitCount-1)で
' LastTopSourceと同一規約。内部状態は一切変更しない(検索/回答ロジックに影響なし)。
' ----------------------------------------------------------------------------
Public Function LastHitCount() As Long
    LastHitCount = mLastNHits
End Function
Public Function LastHitSource(ByVal i As Long) As String
    If i >= 0 And i < mLastNHits Then LastHitSource = mLastHits(i).source
End Function
Public Function LastHitPage(ByVal i As Long) As Long
    If i >= 0 And i < mLastNHits Then LastHitPage = mLastHits(i).page
End Function
Public Function LastHitOrigin(ByVal i As Long) As String
    If i >= 0 And i < mLastNHits Then LastHitOrigin = mLastHits(i).origin
End Function
Public Function LastHitPeek(ByVal i As Long) As String
    If i >= 0 And i < mLastNHits Then LastHitPeek = mLastHits(i).full_text
End Function

Public Sub FeedbackGreen()
    If Not FeedbackAccepted() Then Exit Sub
    ' selfsolve_totalは個人統計のみ。感謝EXPは自己申告では付けず、P2Pで他者の感謝状を受領した時だけ(modP2P)。
    modStats.Bump "selfsolve_total"
    ' 節約時間の日付キー蓄積(1解決=15分)。日/月/年キーなので跨げば自動リセット、
    ' 過去キーがそのまま履歴になる(modBoardのウィジェット/ビーコンが読む)。
    ' 発火点はFeedbackAcceptedガードの内側=多重カウント不可。
    On Error Resume Next
    modStats.Bump "sv:d:" & Format$(Date, "yyyymmdd"), 15
    modStats.Bump "sv:m:" & Format$(Date, "yyyymm"), 15
    modStats.Bump "sv:y:" & Format$(Date, "yyyy"), 15
    On Error GoTo 0
    modLog.LogUsage "feedback_green", mLastMode, "q=" & modUtil.SafeLeft(mLastQuestion, 200)
    On Error Resume Next
    modP2P.EmitThanksForLastAnswer   ' 他者の共有ナレッジ由来なら作者へ感謝状(自作/出所不明は送らない)
    On Error GoTo 0

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
        modUIMain.SetStage "" & ChrW(&HD83E) & ChrW(&HDDED) & " 質問を分析中…"
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
        modUIMain.SetStage "" & ChrW(&HD83E) & ChrW(&HDDEE) & " 関連度を精査中…"
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

    ' 会話履歴(prevU/prevA)は下書き段のみに渡す。検証段は本棚抜粋との照合に専念のため渡さない(裁定D11)。
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

' 低関連度警告(表示専用): 検索ヒットの最高スコアがconfig low_hit_warn_score
' (既定0.3・コサイン類似度+キーワード/2gramボーナスのスケール。modRetrieve.bas
' 冒頭コメント参照)未満なら、表示用resultの先頭に注意書きを付ける。
' 0以下の設定値は「無効化」として扱う(閾値なしで常時警告になるのを防ぐ)。
Private Function ApplyLowHitWarning(ByVal result As String, hits() As Hit, ByVal nHits As Long) As String
    ApplyLowHitWarning = result
    Dim threshold As Double
    threshold = modConfig.GetDouble("low_hit_warn_score", 0.3)
    If threshold <= 0 Then Exit Function

    Dim maxScore As Double
    maxScore = -1E+30
    Dim i As Long
    For i = 1 To nHits
        If hits(i).score > maxScore Then maxScore = hits(i).score
    Next i
    If maxScore >= threshold Then Exit Function

    ApplyLowHitWarning = "⚠️ 手元の資料との関連が薄い可能性があります。回答は参考程度にご覧ください。" & _
        vbLf & vbLf & result
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

' ----------------------------------------------------------------------------
' prevU/prevA用履歴(裁定D11): 成功した各ターンのQ&Aを「新しい順;;;区切り」で
' セッション保持する(最大 config followup_max_pairs 既定3ペア)。
' mHistory(プロンプト内履歴)とは別物: こちらはリボンChatGPT()の確定引数
' prevU/prevA(台帳§1 #1 第7・8引数)へそのまま渡すための形式。
' パース/整形の純関数(SplitFollowupTrailer/KeepNewestPairs/
' SanitizeForFollowupHistory)は modFollowup へ分離済み(文字数上限対策)。
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

