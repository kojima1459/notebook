Attribute VB_Name = "modApp"
Option Explicit

' ============================================================================
' modApp - Nexus Agent Controller(DOCS_NEXUS_SPEC Phase 1-2)
' ----------------------------------------------------------------------------
' UI(modUI)とエンジン(modAsk=多段RAG/modGateway/modShelf)を結合する制御層。
' ・モード切替: 社内ナレッジ検索(RAG=modAsk.Answer) / 一般アシスタント(CallLLM直)
' ・コンテキスト・アクション(裁定②): 固定アクションバーは「選択中のAIバブル」
'   (クリックでOnSelectBubbleが記録。未選択時は最新のAIバブル)に対して発火
' ・P2P共有パスは config nexus_share_path(既定は下記定数)で差替え可能(裁定③)
' ============================================================================

Private Const SHARE_PATH_DEFAULT As String = "\\pgiofs01\Nexus_Share\"
Private Const MODE_KEY As String = "nexus_mode"      ' rag / normal

Private mActiveBubble As String
Private mBusy As Boolean
Private mGenPrevU As String   ' 一般モードの会話履歴(新しい順;;;区切り)
Private mGenPrevA As String

' ----------------------------------------------------------------------------
' LaunchNexus - Nexus UIの起動(modBootから呼ばれる)
' ----------------------------------------------------------------------------
Public Sub LaunchNexus()
    modUI.InitUI
    modUI.AddChatBubble "ai", _
        "こんにちは。Nexus Agentです。" & vbLf & _
        "上のモードボタンで「社内ナレッジ検索」(本棚の資料から出典付きで回答)と" & _
        "「一般アシスタント」を切り替えられます。メッセージを入力して送信してください。"
End Sub

' ----------------------------------------------------------------------------
' OnSend - 送信ボタン。入力セル(nx_input)を読み、モードに応じて回答生成。
' ----------------------------------------------------------------------------
Public Sub OnSend()
    If mBusy Then Exit Sub
    mBusy = True
    On Error GoTo Fail

    Dim q As String
    q = ReadInputCell()
    If LenB(Trim$(q)) = 0 Then
        mBusy = False
        MsgBox "メッセージを入力してから送信してください。", vbInformation, "Nexus Agent"
        Exit Sub
    End If

    modUI.AddChatBubble "user", q
    ClearInputCell

    Dim ans As String
    If CurrentMode() = "normal" Then
        ans = AskGeneral(q)
    Else
        ans = modAsk.Answer(q, RagSpeed())
    End If

    Dim bubbleName As String
    bubbleName = modUI.AddChatBubble("ai", ans)
    mActiveBubble = bubbleName
    modUI.MarkActiveBubble bubbleName

    mBusy = False
    Exit Sub

Fail:
    Dim failDesc As String: failDesc = Err.Description
    Err.Clear
    On Error Resume Next
    modLog.LogError "E0602", "modApp.OnSend", failDesc
    modUI.AddChatBubble "ai", "エラーが発生しました。もう一度お試しください。(" & failDesc & ")"
    On Error GoTo 0
    mBusy = False
End Sub

' ----------------------------------------------------------------------------
' OnSelectBubble - AIバブルのクリック(コンテキスト・アクションの対象指定)
' ----------------------------------------------------------------------------
Public Sub OnSelectBubble()
    Dim callerName As String
    On Error Resume Next
    callerName = CStr(Application.Caller)
    On Error GoTo 0
    If LenB(callerName) = 0 Then Exit Sub

    mActiveBubble = callerName
    modUI.MarkActiveBubble callerName
End Sub

' ----------------------------------------------------------------------------
' フローティング・アクションバー(裁定②): 選択中バブルに対して発火
' ----------------------------------------------------------------------------
Public Sub OnActGood()
    If Not HasTarget() Then Exit Sub
    On Error Resume Next
    modStats.Bump "hint_total"
    modLog.LogUsage "feedback_good", CurrentMode(), modUtil.SafeLeft(TargetText(), 120)
    On Error GoTo 0
    MsgBox "ありがとうございます。評価を記録しました。", vbInformation, "Nexus Agent"
End Sub

Public Sub OnActBad()
    If Not HasTarget() Then Exit Sub
    On Error Resume Next
    modStats.Bump "fail_total"
    modLog.LogUsage "feedback_bad", CurrentMode(), modUtil.SafeLeft(TargetText(), 120)
    On Error GoTo 0

    ' RLHF簡易版(Phase 2): 正しい内容を教えてもらい、ナレッジとして学習する
    Dim fix As String
    fix = InputBox("この回答の正しい内容・修正点を教えてください。" & vbCrLf & _
                   "入力いただいた内容はナレッジとして学習し、次回から回答に反映されます。" & vbCrLf & _
                   "(空欄のまま閉じると記録のみ行います)", "Nexus Agent - 自己学習")
    If LenB(Trim$(fix)) = 0 Then Exit Sub

    Dim body As String
    body = "【修正ナレッジ】" & vbLf & _
           "対象の回答(抜粋): " & modUtil.SafeLeft(TargetText(), 400) & vbLf & vbLf & _
           "正しい内容: " & fix
    If modVault.RegisterKnowledgeText("修正ナレッジ", body, "修正,フィードバック") Then
        MsgBox "学習しました。次回の回答から反映されます。", vbInformation, "Nexus Agent"
    Else
        MsgBox "学習の保存に失敗しました。マイ本棚の一覧をご確認ください。", vbExclamation, "Nexus Agent"
    End If
End Sub

Public Sub OnActDrill()
    Dim q As String
    q = InputBox("さらに深掘りしたい内容を入力してください。" & vbCrLf & _
                 "(直前までの会話を踏まえて回答します)", "Nexus Agent - 深掘り")
    If LenB(Trim$(q)) = 0 Then Exit Sub

    If mBusy Then Exit Sub
    mBusy = True
    On Error GoTo Fail

    modUI.AddChatBubble "user", ChrW(&H1F50D) & " " & q

    Dim ans As String
    If modAsk.CanFollowup() Then
        modAsk.AskFollowup q
        ans = modAsk.LastAnswerText()
    Else
        ans = modAsk.Answer(q, RagSpeed())
    End If

    Dim bubbleName As String
    bubbleName = modUI.AddChatBubble("ai", ans)
    mActiveBubble = bubbleName
    modUI.MarkActiveBubble bubbleName
    mBusy = False
    Exit Sub

Fail:
    Err.Clear
    On Error GoTo 0
    mBusy = False
End Sub

Public Sub OnActResolve()
    If Not HasTarget() Then Exit Sub
    modAsk.FeedbackGreen   ' selfsolve_total加算+多重防止は既存ガードに従う
End Sub

Public Sub OnActHq()
    MsgBox "本社システムへの照会は準備中です。" & vbLf & _
           "(Phase 4で共有フォルダ " & SharePath() & " 連携として実装予定)", _
           vbInformation, "Nexus Agent"
End Sub

Public Sub OnActWord()
    If Not HasTarget() Then Exit Sub
    Dim answerBody As String
    answerBody = TargetText()

    Dim result As Variant
    result = modFeatures.InvokeFeature("markdown", "ExportAnswerAsDoc", Array(answerBody, ""))
    If VarType(result) = vbString Then
        If Left$(CStr(result), 5) = "#ERR:" Then
            MsgBox "Word出力は現在利用できません(管理者が有効化すると使えます)。", _
                   vbInformation, "Nexus Agent"
        End If
    End If
End Sub

' ----------------------------------------------------------------------------
' OnAttachImage - 📎 クリップボード画像でVisionチャット(GPTV連携)
' ----------------------------------------------------------------------------
Public Sub OnAttachImage()
    If modConfig.GetBool("mock_llm", True) Then
        MsgBox "画像チャットは本番環境(AIリボンあり)でのみ動作します。", vbInformation, "Nexus Agent"
        Exit Sub
    End If

    Dim hasImg As Variant
    hasImg = modFeatures.InvokeFeature("vision", "HasClipboardImage", Array())
    If VarType(hasImg) = vbString Or Not CBool(hasImg) Then
        MsgBox "クリップボードに画像がありません。" & vbCrLf & _
               "画面をコピー(Win+Shift+S等)してから、もう一度押してください。", _
               vbInformation, "Nexus Agent"
        Exit Sub
    End If

    Dim prompt As String
    prompt = ReadInputCell()
    If LenB(Trim$(prompt)) = 0 Then prompt = "この画像の内容を読み取り、要点を説明してください。"

    If mBusy Then Exit Sub
    mBusy = True
    On Error GoTo Fail

    modUI.AddChatBubble "user", ChrW(&H1F4CE) & "(画像) " & prompt
    ClearInputCell

    Dim b64 As Variant
    b64 = modGateway.TryRibbonRun("Base64FromCB", Array())
    Dim ans As String
    If VarType(b64) = vbString And LenB(CStr(b64)) > 0 And Left$(CStr(b64), 5) <> "#ERR:" Then
        Dim resp As Variant
        resp = modGateway.TryRibbonRun("ChatGPTV", Array(prompt, CStr(b64), "", "high", "マイ本棚AI:nexus_vision"))
        ans = CStr(resp)
        If LenB(ans) = 0 Or Left$(ans, 5) = "#ERR:" Then
            ans = "画像の解析に失敗しました。もう一度お試しください。"
        End If
    Else
        ans = "画像の取得に失敗しました。画像をコピーし直してからお試しください。"
    End If

    Dim bubbleName As String
    bubbleName = modUI.AddChatBubble("ai", ans)
    mActiveBubble = bubbleName
    modUI.MarkActiveBubble bubbleName
    mBusy = False
    Exit Sub

Fail:
    Err.Clear
    On Error GoTo 0
    mBusy = False
End Sub

' ----------------------------------------------------------------------------
' ナビゲーション(SPA遷移)・モード/言語トグル
' ----------------------------------------------------------------------------
Public Sub OnNavChat()
    On Error Resume Next
    ThisWorkbook.Worksheets("Nexus").Activate
    On Error GoTo 0
End Sub

Public Sub OnNavVault()
    modVault.ShowVaultGallery
End Sub

Public Sub OnNavDash()
    modDash.ShowDashboard
End Sub

Public Sub OnToggleMode()
    Dim newMode As String
    If CurrentMode() = "normal" Then
        newMode = "rag"
    Else
        newMode = "normal"
    End If
    WriteUiState MODE_KEY, newMode
    UpdateModeButton
End Sub

' 回答言語の巡回切替(日本語→English→中文→Tiếng Việt)。
' answer_languageは既存プロンプト(modPrompts)がそのまま使用する。
Public Sub OnLangCycle()
    Dim cur As String
    cur = modConfig.GetString("answer_language", "日本語")

    Dim nextLang As String
    Select Case cur
        Case "日本語": nextLang = "English"
        Case "English": nextLang = "中文"
        Case "中文": nextLang = "Tiếng Việt"
        Case Else: nextLang = "日本語"
    End Select
    modConfig.SetValue "answer_language", nextLang

    On Error Resume Next
    ThisWorkbook.Worksheets("Nexus").Shapes("nx_top_lang").TextFrame2.TextRange.Text = _
        ChrW(&H1F310) & " " & nextLang & "で回答"
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' 対象バブル(選択中→無ければ最新のAI回答)があるか。無ければ案内してFalse。
Private Function HasTarget() As Boolean
    If LenB(TargetBubbleName()) = 0 Then
        MsgBox "対象のAI回答がありません。まず質問して回答を受け取ってください。" & vbCrLf & _
               "(過去の回答に対して操作する場合は、その吹き出しをクリックして選択してから押してください)", _
               vbInformation, "Nexus Agent"
        Exit Function
    End If
    HasTarget = True
End Function

Private Function TargetBubbleName() As String
    If LenB(mActiveBubble) > 0 Then
        If LenB(modUI.BubbleTextOf(mActiveBubble)) > 0 Then
            TargetBubbleName = mActiveBubble
            Exit Function
        End If
    End If
    TargetBubbleName = modUI.LatestAiBubbleName()
End Function

Private Function TargetText() As String
    TargetText = modUI.BubbleTextOf(TargetBubbleName())
End Function

Private Function SharePath() As String
    SharePath = modConfig.GetString("nexus_share_path", SHARE_PATH_DEFAULT)
End Function

' 一般アシスタントモード: 本棚を介さずCallLLM直(会話履歴つき)。
Private Function AskGeneral(ByVal q As String) As String
    Dim sys As String
    sys = "あなたは親切で有能な社内アシスタントです。" & _
          modConfig.GetString("answer_language", "日本語") & "で、簡潔かつ正確に回答してください。"

    Dim lat As Long
    Dim resp As String
    resp = modGateway.CallLLM(sys & vbLf & vbLf & "## 質問" & vbLf & q, "nexus_general", _
        modConfig.GetString("quick_effort", "low"), _
        modConfig.GetString("quick_verbosity", "low"), _
        modConfig.GetString("quick_model", "gpt-5.5"), lat, mGenPrevU, mGenPrevA)

    If Left$(resp, 5) = "#ERR:" Then
        AskGeneral = "回答の作成に失敗しました。時間を置いてもう一度お試しください。"
        Exit Function
    End If

    ' 会話履歴(新しい順;;;区切り・最大followup_max_pairsペア)
    Dim maxPairs As Long
    maxPairs = modConfig.GetLong("followup_max_pairs", 3)
    If maxPairs > 0 Then
        mGenPrevU = TrimPairs(q & IIf(LenB(mGenPrevU) > 0, ";;;" & mGenPrevU, ""), maxPairs)
        mGenPrevA = TrimPairs(modUtil.SafeLeft(resp, 2000) & IIf(LenB(mGenPrevA) > 0, ";;;" & mGenPrevA, ""), maxPairs)
    End If
    AskGeneral = resp
End Function

' ";;;"区切り文字列を先頭maxN件へ切り詰める。
Private Function TrimPairs(ByVal s As String, ByVal maxN As Long) As String
    Dim parts() As String: parts = Split(s, ";;;")
    Dim n As Long: n = UBound(parts) - LBound(parts) + 1
    If n <= maxN Then
        TrimPairs = s
        Exit Function
    End If
    Dim keep() As String: ReDim keep(0 To maxN - 1)
    Dim i As Long
    For i = 0 To maxN - 1
        keep(i) = parts(LBound(parts) + i)
    Next i
    TrimPairs = Join(keep, ";;;")
End Function

' RAGモードの速度(既存ui_stateのquick/deep設定を流用。既定quick)。
Private Function RagSpeed() As String
    RagSpeed = ReadUiState("mode", "quick")
    If RagSpeed <> "deep" Then RagSpeed = "quick"
End Function

Private Function CurrentMode() As String
    CurrentMode = ReadUiState(MODE_KEY, "rag")
    If CurrentMode <> "normal" Then CurrentMode = "rag"
End Function

Private Sub UpdateModeButton()
    Dim caption As String
    If CurrentMode() = "normal" Then
        caption = ChrW(&H1F310) & " 一般アシスタント"
    Else
        caption = ChrW(&H1F3E2) & " 社内ナレッジ検索"
    End If
    On Error Resume Next
    ThisWorkbook.Worksheets("Nexus").Shapes("nx_top_mode").TextFrame2.TextRange.Text = caption
    On Error GoTo 0
End Sub

Private Function ReadInputCell() As String
    On Error Resume Next
    Dim v As Variant
    v = ThisWorkbook.Names("nx_input").RefersToRange.Value
    On Error GoTo 0
    If IsEmpty(v) Or IsError(v) Then Exit Function
    ReadInputCell = modUtil.SafeLeft(CStr(v), 3000)
End Function

Private Sub ClearInputCell()
    On Error Resume Next
    ThisWorkbook.Names("nx_input").RefersToRange.Value = ""
    On Error GoTo 0
End Sub

Private Function ReadUiState(ByVal keyName As String, ByVal defaultVal As String) As String
    ReadUiState = defaultVal
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_UISTATE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(CStr(ws.Cells(i, 1).Value), keyName, vbTextCompare) = 0 Then
            Dim v As String
            v = Trim$(CStr(ws.Cells(i, 2).Value))
            If LenB(v) > 0 Then ReadUiState = LCase$(v)
            Exit Function
        End If
    Next i
End Function

Private Sub WriteUiState(ByVal keyName As String, ByVal valText As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_UISTATE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim r As Long: r = 0
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(CStr(ws.Cells(i, 1).Value), keyName, vbTextCompare) = 0 Then
            r = i
            Exit For
        End If
    Next i
    If r = 0 Then
        r = lastRow + 1
        If r < 1 Then r = 1
        ws.Cells(r, 1).Value = keyName
    End If
    ws.Cells(r, 2).Value = valText
End Sub
