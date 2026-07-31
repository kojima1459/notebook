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

    If LenB(q) = 0 Then
        ' 空質問でもmLast*を必ず更新する。しないと前回のヒットが残り、空クリックが
        ' その資料で回答したように見える(Wave4実バグ)。mLastMode=""は
        ' 「検索も回答生成もしなかった」印でmodUIMain.RenderAnswerも使う。
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

    ' 2026-07-28(レビュー M-4): 回答バッファを毎ターン空にする。
    ' 設定するのは DecorateWithFollowups(成功ターンだけ)なので、
    ' ここで消しておかないと 0件回答・聞き返しのターンでも前回の回答が
    ' residual として残り、✅を押したときに噛み合わないQ&Aが部内へ流れる。
    mLastCleanAnswer = ""

    On Error Resume Next
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo 0

    On Error GoTo Fail

    modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDD0D) & " 検索中…"
    Dim topK As Long
    topK = TopKFor(mdMode)

    ' 多段RAG(§C)。retrieve_mode=singleで従来の単段Searchへ完全退化。
    If LCase$(modConfig.GetString("retrieve_mode", "single")) = "multi" Then
        nHits = modAskRetrieve.RunMultiRetrieve(q, mdMode, topK, hits)
    Else
        nHits = modRetrieve.Search(q, topK, hits)
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
        ElseIf modMode.UseVerify(mdMode) Then
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
    ' Resume で抜けてハンドラ実行中の状態を解除する(On Error GoTo 0 では
    ' 解除されず、Done: の後始末で起きたエラーが呼び出し元へ素通りする)。
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
    End If

    ' 低関連度警告(表示専用): 履歴(AppendHistory/mLastCleanAnswer)は上で
    ' 既に確定済みのため、ここでresultに警告を足しても履歴側には混入しない。
    If ok And nHits > 0 Then
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
'  modLive.UniqueSourceCount)はすべて For i = 0 To n - 1 で回している。
' しかし内部の mLastHits は ReDim(1 To n) なので、i をそのまま添字に使うと
' i=0 で実行時エラー9(添字が範囲外)になっていた。
' 呼び出し側は On Error でエラーを握る作りのため落ちはせず、代わりに
'   ・出典チップが1枚も出ない(Peek View 機能が丸ごと死ぬ)
'   ・「この分野は さんが詳しいです」という空名ボタンが出る
' という「静かに壊れている」状態になっていた。
' すぐ上の LastTopSource には同じ罠のコメントが残っているのに、
' アクセサ側だけ直し漏れていた。ここで 1 始まりへ変換する。
Public Function LastHitCount() As Long
    LastHitCount = mLastNHits
End Function
Public Function LastHitSource(ByVal i As Long) As String
    If i >= 0 And i < mLastNHits Then LastHitSource = mLastHits(i + 1).source
End Function
Public Function LastHitPage(ByVal i As Long) As Long
    If i >= 0 And i < mLastNHits Then LastHitPage = mLastHits(i + 1).page
End Function
Public Function LastHitOrigin(ByVal i As Long) As String
    If i >= 0 And i < mLastNHits Then LastHitOrigin = mLastHits(i + 1).origin
End Function
Public Function LastHitPeek(ByVal i As Long) As String
    If i >= 0 And i < mLastNHits Then LastHitPeek = mLastHits(i + 1).full_text
End Function

' 回答の信頼度(2=根拠あり/1=部分的/0=乏しい)。検索スコアを人間に見える形に
' して「確認すべきときだけ確認させる」。閾値=config confidence_score_x100。
Public Function LastConfidence() As Long
    If mLastNHits < 1 Then Exit Function

    Dim thr As Double
    On Error Resume Next
    thr = CDbl(modConfig.GetLong("confidence_score_x100", 55)) / 100#
    On Error GoTo 0
    If thr <= 0# Then thr = 0.55

    Dim best As Double, strong As Long
    Dim i As Long
    On Error Resume Next
    For i = 1 To mLastNHits
        If mLastHits(i).score > best Then best = mLastHits(i).score
        If mLastHits(i).score >= thr Then strong = strong + 1
    Next i
    On Error GoTo 0

    If strong >= 2 Then
        LastConfidence = 2
    ElseIf best >= thr Then
        LastConfidence = 1
    End If
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
    ' 2026-07-31(レビュー R8 F8): 加算した「今日の節約時間」は、この直後に
    ' 共有フォルダのビーコンへ反映する必要がある(起動時にしか発信していない
    ' ため、全員のビーコンが「今日 0分」のまま置かれていた)。
    ' ただし modBoard は UI層で、ここ(qa層)から呼ぶと層の向きが逆になる。
    ' 発信は UI層の呼び出し元(modAppAct.OnActResolve)が担当する。
    On Error GoTo 0
    modLog.LogUsage "feedback_green", mLastMode, "q=" & modUtil.SafeLeft(mLastQuestion, 200)
    On Error Resume Next
    modP2P.EmitThanksForLastAnswer   ' 他者の共有ナレッジ由来なら作者へ感謝状(自作/出所不明は送らない)
    ' 共有知フライホイール: 人が正しいと確認したQ&Aは組織の一次情報になる。
    ' 社内ナレッジ検索の回答のときだけ発信する(一般アシスタントの雑談は流さない)。
    ' 2026-07-28(レビュー M-4): 回答本文は成功ターンでしか更新されないのに、
    ' 質問は毎ターン更新される。そのため「0件回答」や「聞き返し」の直後に
    ' ✅を押すと、【今回の質問 + 前回成功ターンの回答】という噛み合わない
    ' ペアが「人が確認したQ&A」として部内へ配信されていた。
    ' 中身が揃っているときだけ発信する。
    If LenB(mLastMode) > 0 And LenB(Trim$(mLastCleanAnswer)) > 0 Then
        modInsightIo.EmitVerifiedQA mLastQuestion, mLastCleanAnswer, LastTopSource()
    End If
    On Error GoTo 0

    MsgBox "ありがとうございます。" & vbCrLf & _
           "この質問と回答は「解決済みQ&A」として部内に共有され、" & vbCrLf & _
           "同じことで困っている人がすぐ答えにたどり着けるようになります。", _
           vbInformation, modAppDef.APP_NAME
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
    ' 共有知フライホイール: 答えられなかった質問は「組織に文書が無い領域」の
    ' 一次情報。資料を書ける人の画面へ自動で流す。
    On Error Resume Next
    modInsightIo.EmitGap mLastQuestion, "wrong"
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
    modInsightIo.EmitGap mLastQuestion, "low_conf"
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

' 内部ヘルパー(すべてPrivate: modAskの公開契約は上記7本のみ)


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
    modUIMain.SetStage ChrW(&H270D) & ChrW(&HFE0F) & " 回答作成中…"

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
    modUIMain.SetStage ChrW(&H270D) & ChrW(&HFE0F) & " 回答を下書き中…"

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

    modUIMain.SetStage ChrW(&H2705) & " 検証中…"

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
    NormalizeMode = modMode.Normalize(mode)
End Function

Private Function TopKFor(ByVal mdMode As String) As Long
    TopKFor = modMode.TopK(mdMode, modConfig.GetLong("topk_quick", 6), _
                           modConfig.GetLong("topk_deep", 12), _
                           modConfig.GetLong("topk_thorough", 16))
End Function

' modAskRetrieve からも使う(#ERR: 応答を検索の途中で捨てる判定)。
Public Function IsErrorResponse(ByVal s As String) As Boolean
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
        ' 2026-07-28(レビュー L-2): メモリだけ消して ui_state を消していな
        ' かったため、無効化したはずなのに保存済みの古い履歴が読み直され、
        ' LLMへ送られ続けていた。保存側も空にする。
        mPrevU = ""
        mPrevA = ""
        modState.SaveState "nexus_ask_prevu", ""
        modState.SaveState "nexus_ask_preva", ""
        Exit Sub
    End If

    ' 2026-07-28(レビュー L-3): 書く前に必ず読む。VBAリセット等でモジュール
    ' 変数が消えた直後に新規質問が来ると、mPrevU が空のまま「先頭に積む」
    ' 処理が走り、保存済みの履歴を1ターンで上書き消去していた。
    ' CanFollowup と同じ遅延ロードをここでも通す。
    If LenB(mPrevU) = 0 Then
        mPrevU = modState.LoadState("nexus_ask_prevu", "")
        mPrevA = modState.LoadState("nexus_ask_preva", "")
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

