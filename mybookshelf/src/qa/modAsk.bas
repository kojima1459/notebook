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
'   ・【既知のR1緊張関係・懸念事項】MASTER_SPEC §3 R1は「機能層→
'     modUIMain.SetStageのみ許可」と明記する一方、§7.3/§7.6は
'     modAsk.Answerが検索中の「出典先出し表示」でmodUIMain.
'     RenderSourcesPreviewを、modAsk.AskFromUIが最後にmodUIMain.
'     RenderAnswerを呼ぶことを文章で明示している。両者は字面上矛盾する。
'     本実装は後者(§7.3/§7.6の具体的シーケンス記述)を優先し、
'     SetStage/RenderSourcesPreview/RenderAnswerの3つをmodUIMain経由で
'     呼ぶ。現時点(modUIMain.bas未実装)ではvba_lint.pyのR1チェックは
'     「未実装モジュール参照」としてSKIP扱いになり自担当分のERRORは
'     出ないが、2-F(UI)がmodUIMainを実装した時点でR1チェックが
'     ERRORへ転じる可能性が高い。Wave3統合時にMASTER_SPEC側のR1例外の
'     文言を「SetStage/RenderSourcesPreview/RenderAnswer」へ広げるか、
'     lint側の例外リストを広げるかの判断を仰ぐこと(本モジュールを
'     勝手に作り変えて対処しない)。
'   ・出典先出しプレビュー(RenderSourcesPreview)は「回答本文が出来上がる
'     前に、ヒットした資料が見えると安心できる」という§8.1のUI意図を
'     素直に実現するために検索直後に呼ぶ。
'   ・Answer(question, mode)は最終回答テキストのみを返す契約(§7.3)なので、
'     直近の検索結果(hits/nHits)・所要秒・実際に使ったモードは
'     モジュール変数(mLast*)に保持し、同一モジュール内のAskFromUIが
'     それを読んでmodUIMain.RenderAnswerに渡す(RenderAnswerの二重呼び
'     出しを避けるため、Answer自身はRenderAnswerを呼ばない)。
'   ・ホームの質問セルは、まだ実装されていないmodUIMain.EnsureLayoutが
'     所有するレイアウトの一部であり、本モジュールはそのセル番地を
'     決め打ちしない。代わりに定義済み名前(Excel Name)
'     "mb_question"(modUIMain.EnsureLayoutが指す先を管理する想定)を
'     経由して読む。名前が未定義の場合は空文字列として扱い、
'     「質問が入力されていません」の案内を返すだけで例外にはしない。
'     これはopt層の「SIGNATURE ASSUMPTION」コメントと同じ考え方の、
'     UI層との連携アサンプションである。Wave3で2-Fと突き合わせること。
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

Private mAsking As Boolean
Private mHistory As String

Private mLastQuestion As String
Private mLastAnswer As String
Private mLastMode As String
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
' ----------------------------------------------------------------------------
Public Function Answer(ByVal question As String, ByVal mode As String) As String
    Dim tStart As Double
    tStart = Timer

    Dim q As String
    q = modUtil.SafeLeft(Trim$(question), MAX_QUESTION_CHARS)
    Dim mdMode As String
    mdMode = NormalizeMode(mode)

    If LenB(q) = 0 Then
        Answer = EMPTY_QUESTION_MESSAGE
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
    nHits = modRetrieve.Search(q, topK, hits)

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
            result = RunDeepFlow(q, hits, nHits, ok)
        Else
            result = RunQuickFlow(q, hits, nHits, ok)
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
        modLog.LogError "E0602", "modAsk.Answer", "mode=" & mdMode & " err=" & errDesc
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

    If ok Then AppendHistory q, result

    mLastQuestion = q
    mLastAnswer = result
    mLastMode = mdMode
    mLastHits = hits
    mLastNHits = nHits
    mLastSeconds = elapsedSec

    modLog.LogUsage "ask", mdMode, "q=" & modUtil.SafeLeft(q, 200), elapsedMs, nHits
    modStats.Bump "ask_" & mdMode & "_total"

    Answer = result
End Function

Public Sub FeedbackGreen()
    modStats.Bump "selfsolve_total"
    modLog.LogUsage "feedback_green", mLastMode, "q=" & modUtil.SafeLeft(mLastQuestion, 200)
    MsgBox "ありがとうございます。解決に役立てて何よりです。", vbInformation, modAppDef.APP_NAME
End Sub

Public Sub FeedbackYellow()
    modStats.Bump "hint_total"
    modLog.LogUsage "feedback_yellow", mLastMode, "q=" & modUtil.SafeLeft(mLastQuestion, 200)
    MsgBox "ありがとうございます。次はもっと的確に答えられるよう活かします。", vbInformation, modAppDef.APP_NAME
End Sub

Public Sub FeedbackRed()
    modStats.Bump "fail_total"
    modLog.LogUsage "feedback_red", mLastMode, "q=" & modUtil.SafeLeft(mLastQuestion, 200)
    MsgBox "ご意見ありがとうございます。改善の参考にします。", vbInformation, modAppDef.APP_NAME
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー(すべてPrivate: modAskの公開契約は上記5本のみ)
' ----------------------------------------------------------------------------

Private Function RunQuickFlow(ByVal q As String, hits() As Hit, ByVal nHits As Long, ByRef ok As Boolean) As String
    modUIMain.SetStage "✍️ 回答作成中…"

    Dim prompt As String
    prompt = modPrompts.BuildQuickPrompt(q, hits, nHits)

    Dim eff As String: eff = modConfig.GetString("quick_effort", "low")
    Dim vrb As String: vrb = modConfig.GetString("quick_verbosity", "low")
    Dim mdl As String: mdl = modConfig.GetString("quick_model", "gpt-5.5")
    Dim latency As Long

    Dim resp As String
    resp = modGateway.CallLLM(prompt, "quick_draft", eff, vrb, mdl, latency)

    If IsErrorResponse(resp) Then
        ok = False
        RunQuickFlow = BuildErrorAnswer(resp)
    Else
        ok = True
        RunQuickFlow = resp
    End If
End Function

Private Function RunDeepFlow(ByVal q As String, hits() As Hit, ByVal nHits As Long, ByRef ok As Boolean) As String
    modUIMain.SetStage "✍️ 回答を下書き中…"

    Dim draftPrompt As String
    draftPrompt = modPrompts.BuildDeepDraftPrompt(q, hits, nHits, HistoryBlock())

    Dim dEff As String: dEff = modConfig.GetString("deep_draft_effort", "medium")
    Dim dVrb As String: dVrb = modConfig.GetString("deep_draft_verbosity", "high")
    Dim mdl As String: mdl = modConfig.GetString("recommended_model", "gpt-5.5")
    Dim latency As Long

    Dim draft As String
    draft = modGateway.CallLLM(draftPrompt, "deep_draft", dEff, dVrb, mdl, latency)

    If IsErrorResponse(draft) Then
        ok = False
        RunDeepFlow = BuildErrorAnswer(draft)
        Exit Function
    End If

    modUIMain.SetStage "✅ 検証中…"

    Dim verifyPrompt As String
    verifyPrompt = modPrompts.BuildDeepVerifyPrompt(q, draft, hits, nHits)

    Dim vEff As String: vEff = modConfig.GetString("deep_verify_effort", "high")
    Dim vVrb As String: vVrb = modConfig.GetString("deep_verify_verbosity", "medium")

    Dim verified As String
    verified = modGateway.CallLLM(verifyPrompt, "deep_verify", vEff, vVrb, mdl, latency)

    ok = True
    If IsErrorResponse(verified) Then
        RunDeepFlow = draft & vbLf & vbLf & _
            "(注: 検証段階でエラーが発生したため、下書きの内容を表示しています。)"
    Else
        RunDeepFlow = verified
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
