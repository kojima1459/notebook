Attribute VB_Name = "modApp"
Option Explicit

' modApp - Nexus Agent Controller。UI(modUI)とエンジン(modAsk/modGateway/
' modShelf)を結合する制御層。固定アクションバーは「選択中のAIバブル」
' (未選択時は最新)に対して発火。P2P共有パスはconfig nexus_share_pathで差替可。

Private Const SHARE_PATH_DEFAULT As String = "\\pgiofs01\Nexus_Share\"
Private Const MODE_KEY As String = "nexus_mode"      ' rag / normal
Private Const MAX_INPUT_CHARS As Long = 2000         ' A3: 入力の最大文字数(超過はカット+警告)

' 連打/多重発火はmodUiLockへ一本化(Enter/Leave対で必ずLeave到達)。
Private mActiveBubble As String
Private mGenPrevU As String   ' 一般モードの会話履歴(新しい順;;;区切り)
Private mGenPrevA As String

' LaunchNexus - Nexus UIの起動(modBootから呼ばれる)
Public Sub LaunchNexus()
    modUI.InitUI
    RestoreLastConversation      ' ④前回の続きを薄く復元(失敗しても挨拶へ進む)
    modUI.AddChatBubble "ai", _
        TimeGreeting() & " Nexus Agentです。" & vbLf & _
        "上のモードボタンで「社内ナレッジ検索」(本棚の資料から出典付きで回答)と" & _
        "「一般アシスタント」を切り替えられます。メッセージを入力して送信してください。" & vbLf & _
        "(送信は Ctrl+Enter、呼び出しはどこからでも Ctrl+Shift+Q が使えます)"
    On Error Resume Next         ' 以降は追加機能のフック(各自が内部で握るが二重に防護)
    modBoard.BootBoard           ' チーム連帯ボード: ビーコン発信+集計+サイドバーウィジェット
    modMentor.CollectQuestions   ' Mentor受信: 自分宛の質問を回収
    modHelp.EnsureHelpButton     ' ヘルプ(?)ボタン
    DrawSidebarExtras            ' 質問テンプレチップ+ナレッジガチャ(白紙の恐怖対策)
    modTour.StartTourIfFirstRun  ' 初回オンボーディングツアー
    On Error GoTo 0
End Sub

' サイドバー下部: 質問テンプレチップ3つ+🎲今日のナレッジガチャ。
' nx_sb_接頭辞なので既存のZ-Order/テーマ再彩色ループが自動で面倒を見る。冪等。
Private Sub DrawSidebarExtras()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub

    Dim nm As Variant
    For Each nm In Array("nx_sb_qa1", "nx_sb_qa2", "nx_sb_qa3", "nx_sb_gacha")
        ws.Shapes(CStr(nm)).Delete
    Next nm

    Dim caps As Variant
    caps = Array(ChrW(&HD83D) & ChrW(&HDCAC) & " 改定ポイントを教えて", _
                 ChrW(&HD83D) & ChrW(&HDCAC) & " 用語をやさしく解説", _
                 ChrW(&HD83D) & ChrW(&HDCAC) & " 手続きの流れを知りたい")
    Dim i As Long
    For i = 0 To 2
        Dim chip As Shape
        Set chip = ws.Shapes.AddShape(5, 10, 482 + i * 28, 175, 24)
        chip.Name = "nx_sb_qa" & (i + 1)
        chip.Adjustments(1) = 0.4
        chip.Line.Visible = 0
        chip.Fill.ForeColor.RGB = modUI.UiColor("sidebarActive")
        With chip.TextFrame2
            .WordWrap = -1
            .TextRange.Text = CStr(caps(i))
            .TextRange.Font.Name = "Yu Gothic UI"
            .TextRange.Font.Size = 8.5
            .VerticalAnchor = 3
            .MarginLeft = 10: .MarginTop = 0: .MarginBottom = 0
        End With
        chip.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("sidebarText")
        chip.OnAction = "modApp.OnQuickAsk"
        chip.Placement = 3
    Next i

    Dim g As Shape
    Set g = ws.Shapes.AddShape(5, 10, 574, 175, 26)
    g.Name = "nx_sb_gacha"
    g.Adjustments(1) = 0.4
    g.Line.Visible = -1
    g.Line.Weight = 0.75
    g.Line.ForeColor.RGB = modUI.UiColor("accent")
    g.Fill.Visible = 0
    With g.TextFrame2
        .WordWrap = -1
        .TextRange.Text = ChrW(&HD83C) & ChrW(&HDFB2) & " 今日のワンポイント"
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = -1
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    g.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("accent")
    g.OnAction = "modApp.OnGacha"
    g.Placement = 3
    On Error GoTo 0
End Sub

' テンプレチップのクリック: 入力欄へ雛形を流し込むだけ(送信しない=補助輪)。
Public Sub OnQuickAsk()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    Dim tpl As String
    Select Case CStr(Application.Caller)
        Case "nx_sb_qa1": tpl = "【知りたい改定】: (資料名や年度を記入)" & vbLf & _
            "【気になる点】: (例: 保険料への影響)" & vbLf & _
            "【知りたい結論】: (例: 何がいつから変わるか)"
        Case "nx_sb_qa2": tpl = "【わからない用語】: (ここに記入)" & vbLf & _
            "【それを見た場所】: (例: 〇〇約款 第4条)" & vbLf & _
            "【どこまで理解したいか】: (例: お客様に説明できるレベル)"
        Case "nx_sb_qa3": tpl = "【手続き名】: (ここに記入)" & vbLf & _
            "【お客様/自分の状況】: (例: 契約者が死亡、受取人が海外在住)" & vbLf & _
            "【知りたい結論】: (例: 必要書類と所要日数)"
        Case Else: GoTo Done
    End Select
    On Error Resume Next
    ThisWorkbook.Names("nx_input").RefersToRange.Value = tpl
    On Error GoTo Done
    modSkin.ShowToast "(ここに記入)の部分を埋めて Ctrl+Enter で送信してください。", "info"
Done:
    modUiLock.Leave
End Sub

' 🎲 ナレッジガチャ: my_knowledgeからランダムに1件を「今日のワンポイント」として
' バブル表示(API非通信・完全ローカル)。偶然の学びのエンタメ化。
Public Sub OnGacha()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then GoTo Done

    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then
        modSkin.ShowToast "まだ資料がありません。「ナレッジ倉庫」から登録すると、ここで豆知識が引けます。", "info"
        GoTo Done
    End If

    Randomize
    Dim r As Long: r = 2 + Int(Rnd * (lastR - 1))
    Dim src As String: src = CStr(ws.Cells(r, 2).Value)
    Dim summ As String: summ = CStr(ws.Cells(r, 5).Value)
    Dim body As String: body = CStr(ws.Cells(r, 7).Value)

    modUI.AddChatBubble "ai", _
        ChrW(&HD83C) & ChrW(&HDFB2) & " 今日のワンポイント" & vbLf & _
        "【" & modUtil.SafeLeft(src, 40) & "】" & IIf(LenB(summ) > 0, " " & summ, "") & vbLf & _
        modUtil.SafeLeft(body, 300) & IIf(Len(body) > 300, "…", "") & vbLf & _
        "(もう一度引く: サイドバーの「" & ChrW(&HD83C) & ChrW(&HDFB2) & " 今日のワンポイント」)"
    On Error Resume Next
    modLog.LogUsage "gacha", "", modUtil.SafeLeft(src, 80)
    On Error GoTo Done
Done:
    modUiLock.Leave
End Sub

' ④会話の記憶: 直近2往復をui_stateへ保存し、次回起動時に薄く復元する。
' 区切りはAskGeneral履歴と同じ";;;"(質問/回答に含まれる場合は改行1個に置換して保護)。
Private Sub SaveTurnForRestore(ByVal q As String, ByVal ans As String)
    On Error Resume Next
    Dim u As String, a As String
    u = Replace(modUtil.SafeLeft(q, 300), ";;;", " ")
    a = Replace(modUtil.SafeLeft(ans, 700), ";;;", " ")
    Dim prevU As String: prevU = modState.LoadState("nexus_hist_u", "")
    Dim prevA As String: prevA = modState.LoadState("nexus_hist_a", "")
    modState.SaveState "nexus_hist_u", TrimPairs(u & IIf(LenB(prevU) > 0, ";;;" & prevU, ""), 2)
    modState.SaveState "nexus_hist_a", TrimPairs(a & IIf(LenB(prevA) > 0, ";;;" & prevA, ""), 2)
    On Error GoTo 0
End Sub

Private Sub RestoreLastConversation()
    On Error Resume Next
    Dim histU As String: histU = modState.LoadState("nexus_hist_u", "")
    Dim histA As String: histA = modState.LoadState("nexus_hist_a", "")
    If LenB(histU) = 0 Or LenB(histA) = 0 Then Exit Sub

    Dim us() As String: us = Split(histU, ";;;")
    Dim aas() As String: aas = Split(histA, ";;;")
    Dim n As Long: n = UBound(us)
    If UBound(aas) < n Then n = UBound(aas)

    ' 保存は新しい順なので、古い方から描く(チャットは下が最新)
    Dim i As Long
    For i = n To 0 Step -1
        If LenB(Trim$(us(i))) > 0 Then
            modUI.AddChatBubble "user", us(i)
            modUI.AddChatBubble "ai", ChrW(&HD83D) & ChrW(&HDCDC) & "(前回の回答) " & aas(i)
        End If
    Next i
    On Error GoTo 0
End Sub

' OnSend - 送信ボタン。入力セル(nx_input)を読み、モードに応じて回答生成。
Public Sub OnSend()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Fail
    modPeek.HideCitations   ' 前回回答の出典チップ/ポップアップを消す(最新回答の下だけに出す)
    modMentor.ClearMentor   ' Mentorボタンも同時に掃除(内部On Error Resume Next=安全弁)

    Dim q As String
    q = ReadInputCell()
    If LenB(Trim$(q)) = 0 Then
        modUiLock.Leave
        modSkin.ShowToast "はじめにメッセージをご入力ください。ご質問をお待ちしています。", "info"
        Exit Sub
    End If

    ' A3: 異常な文字数の入力を防ぐ。数千文字の貼り付けはShapeの高さ計算限界や
    '     APIのトークン上限溢れでクラッシュ/エラーを招くため、上限で切って警告する。
    If Len(q) > MAX_INPUT_CHARS Then
        q = Left$(q, MAX_INPUT_CHARS)
        MsgBox "入力が長いため、先頭 " & MAX_INPUT_CHARS & " 文字だけを送信します。" & vbLf & _
               "長い資料は「ナレッジ倉庫」に取り込んでから質問すると、全文を対象に回答できます。", _
               vbInformation, "Nexus Agent"
    End If

    ' 遊び心: 弱音キーワードはAPIに投げず、関西弁コンシェルジュが即座に労う
    ' (意図的なタイミング限定・完全ローカルなので事故りようがない)
    If IsTiredWords(q) Then
        modUI.AddChatBubble "user", q
        ClearInputCell
        modUI.AddChatBubble "ai", ComfortMessage()
        modUiLock.Leave
        Exit Sub
    End If

    modUI.AddChatBubble "user", q
    ClearInputCell

    ' 体感速度ハック: 待ち時間の無反応(壊れた?)を防ぐため、考え中バブルを即時表示。
    ' 回答が来たら削除して本物を追加する(in-place置換はバブル高さ管理と衝突するため
    ' 削除→追加方式。小さな余白が残るだけで崩れない)。
    Dim phName As String
    phName = modUI.AddChatBubble("ai", ChrW(&HD83D) & ChrW(&HDCAD) & " 考えています…")
    DoEvents

    Dim ans As String
    If CurrentMode() = "normal" Then
        ans = AskGeneral(q)
    Else
        ans = modAsk.Answer(q, RagSpeed())
    End If

    On Error Resume Next
    If LenB(phName) > 0 Then ThisWorkbook.Worksheets("Nexus").Shapes(phName).Delete
    On Error GoTo Fail

    Dim bubbleName As String
    bubbleName = modUI.AddChatBubble("ai", ans)
    mActiveBubble = bubbleName
    modUI.MarkActiveBubble bubbleName
    SaveTurnForRestore q, ans   ' ④記憶の継続: 次回起動時の「前回の続き」復元用に保存

    ' Peek View: RAG(社内ナレッジ検索)回答のときだけ、出典チップを回答直下に描画する
    ' (一般アシスタントは出典が無いので出さない=古いチップの誤表示も防ぐ)。
    If CurrentMode() <> "normal" Then
        modPeek.RenderCitations bubbleName
        modMentor.OfferMentor bubbleName   ' Mentor: 専門家ボタン(失敗しても出ないだけ=安全弁内蔵)
    End If

    ' 爆速証明(狂気案Lv.1): binary_rag_debug=TRUEのとき、直近ハイブリッド検索の所要msを
    ' Toastで見せる(qa層のperfログをUI層で取り出す=R1レイヤリングを守る)。
    If modConfig.GetBool("binary_rag_debug", False) Then
        Dim perf As String: perf = modBitwiseOpt.ConsumePerfLog()
        If LenB(perf) > 0 Then modSkin.ShowToast perf, "info"
    End If

    modUiLock.Leave
    Exit Sub

Fail:
    Dim failDesc As String: failDesc = Err.Description
    Err.Clear
    On Error Resume Next
    modLog.LogError "E0602", "modApp.OnSend", failDesc
    modUI.AddChatBubble "ai", "エラーが発生しました。もう一度お試しください。(" & failDesc & ")"
    On Error GoTo 0
    modUiLock.Leave
End Sub

' OnSelectBubble - AIバブルのクリック(コンテキスト・アクションの対象指定)
Public Sub OnSelectBubble()
    Dim callerName As String
    On Error Resume Next
    callerName = CStr(Application.Caller)
    On Error GoTo 0
    If LenB(callerName) = 0 Then Exit Sub

    mActiveBubble = callerName
    modUI.MarkActiveBubble callerName
End Sub

' Peek View(出典ポップアップ): 出典チップ/ポップアップのクリック受け。
' 出典チップ(nx_cite_<i>)のクリック → そのチャンク本文をポップアップ表示。
Public Sub OnPeek()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    Dim caller As String
    caller = CStr(Application.Caller)
    If Left$(caller, 8) = "nx_cite_" Then
        modPeek.ShowPeek CLng(Val(Mid$(caller, 9)))
    End If
Done:
    modUiLock.Leave
End Sub

' 出典ポップアップ(nx_peek)のクリック → 閉じる。
Public Sub OnPeekClose()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modPeek.HidePeek
    On Error GoTo 0
    modUiLock.Leave
End Sub

' フローティング・アクションバー(裁定②): 選択中バブルに対して発火
Public Sub OnActGood()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    If Not HasTarget() Then GoTo Done
    On Error Resume Next
    modStats.Bump "hint_total"
    modLog.LogUsage "feedback_good", CurrentMode(), modUtil.SafeLeft(TargetText(), 120)
    On Error GoTo Done
    modSkin.ShowToast "フィードバックありがとうございます。今後の回答の質に活かします。", "success"
Done:
    modUiLock.Leave
End Sub

Public Sub OnActBad()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    If Not HasTarget() Then GoTo Done
    On Error Resume Next
    modStats.Bump "fail_total"
    modLog.LogUsage "feedback_bad", CurrentMode(), modUtil.SafeLeft(TargetText(), 120)
    On Error GoTo Done

    ' RLHF簡易版(Phase 2): 正しい内容を教えてもらい、ナレッジとして学習する
    Dim fixText As String
    fixText = InputBox("この回答の正しい内容・修正点を教えてください。" & vbCrLf & _
                   "入力いただいた内容はナレッジとして学習し、次回から回答に反映されます。" & vbCrLf & _
                   "(空欄のまま閉じると記録のみ行います)", "Nexus Agent - 自己学習")
    If LenB(Trim$(fixText)) = 0 Then GoTo Done

    Dim body As String
    body = "【修正ナレッジ】" & vbLf & _
           "対象の回答(抜粋): " & modUtil.SafeLeft(TargetText(), 400) & vbLf & vbLf & _
           "正しい内容: " & fixText
    If modVault.RegisterKnowledgeText("修正ナレッジ", body, "修正,フィードバック") Then
        modSkin.ShowToast "教えていただきありがとうございます。次回の回答から反映します。", "success"
    Else
        MsgBox "学習の保存に失敗しました。マイ本棚の一覧をご確認ください。", vbExclamation, "Nexus Agent"
    End If
Done:
    modUiLock.Leave
End Sub

Public Sub OnActDrill()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Fail
    modPeek.HideCitations   ' 前回の出典チップ/ポップアップを消す
    modMentor.ClearMentor   ' Mentorボタンも掃除(安全弁内蔵)

    Dim q As String
    q = InputBox("さらに深掘りしたい内容を入力してください。" & vbCrLf & _
                 "(直前までの会話を踏まえて回答します)", "Nexus Agent - 深掘り")
    If LenB(Trim$(q)) = 0 Then
        modUiLock.Leave
        Exit Sub
    End If
    If Len(q) > MAX_INPUT_CHARS Then q = Left$(q, MAX_INPUT_CHARS)   ' A3: 上限で切る

    modUI.AddChatBubble "user", ChrW(&HD83D) & ChrW(&HDD0D) & " " & q

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
    modPeek.RenderCitations bubbleName   ' Peek View: 深掘り回答の出典チップ
    modMentor.OfferMentor bubbleName     ' Mentor: 専門家ボタン(安全弁内蔵)
    modUiLock.Leave
    Exit Sub

Fail:
    Err.Clear
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnActResolve()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    If Not HasTarget() Then GoTo Done
    modAsk.FeedbackGreen   ' selfsolve_total加算+多重防止は既存ガードに従う
    On Error Resume Next
    modBoard.DrawWidget   ' 節約時間ウィジェット再描画(起動時のみで固まっていた対策)
    On Error GoTo Done
Done:
    modUiLock.Leave
End Sub

Public Sub OnActHq()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next   ' 何が起きてもLeaveへ到達させる(ロック取りっぱなし=永久フリーズ防止)
    modSkin.ShowToast "本社への照会機能は準備中です。公開までいましばらくお待ちください。", "info"
    On Error GoTo 0
    modUiLock.Leave
End Sub

' ホームのOnOpenWordButtonと同じInputBoxを挟む(以前は指示文なし=""固定だった)。
Public Sub OnActWord()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    If Not HasTarget() Then GoTo Done
    Dim answerBody As String
    answerBody = TargetText()

    Dim resp As Variant
    resp = Application.InputBox( _
        Prompt:="どんな文書に仕上げますか?" & vbCrLf & _
                "(例: お客様向けの回答文書風に / 社内回覧用の要約に)" & vbCrLf & _
                "※空欄ならそのまま転記", _
        Title:=modAppDef.APP_NAME & " - Wordで開く", Default:="", Type:=2)
    If VarType(resp) = vbBoolean Then GoTo Done   ' キャンセル→何もしない
    Dim instruction As String
    instruction = Trim$(CStr(resp))

    Dim result As Variant
    result = modFeatures.InvokeFeature("markdown", "ExportAnswerAsDoc", Array(answerBody, instruction))
    If VarType(result) = vbString Then
        If Left$(CStr(result), 5) = "#ERR:" Then
            MsgBox "Word出力は現在利用できません(管理者が有効化すると使えます)。", _
                   vbInformation, "Nexus Agent"
        End If
    End If
Done:
    modUiLock.Leave
End Sub

' 📋 コピー(盲点C2/D6): 選択中(無ければ最新)のAIバブル本文をクリップボードへ。
' Shape(図形)の文字は手で綺麗にコピーできないため明示ボタンを用意し、文字化けしない
' Unicode方式(modClip)で格納する。
Public Sub OnActCopy()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    If Not HasTarget() Then GoTo Done
    Dim t As String: t = TargetText()
    If LenB(t) = 0 Then GoTo Done
    If modClip.SetClipboardText(t) Then
        modSkin.ShowToast "回答をコピーしました。Ctrl+V でどこへでも貼り付けできます。", "success"
    Else
        MsgBox "コピーに失敗しました。お使いの環境では手動での選択をお試しください。", _
               vbExclamation, "Nexus Agent"
    End If
Done:
    modUiLock.Leave
End Sub

' OnAttachImage - 📎 クリップボード画像でVisionチャット(GPTV連携)
Public Sub OnAttachImage()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Fail

    If modConfig.GetBool("mock_llm", True) Then
        modUiLock.Leave
        MsgBox "画像チャットは本番環境(AIリボンあり)でのみ動作します。", vbInformation, "Nexus Agent"
        Exit Sub
    End If

    Dim hasImg As Variant
    hasImg = modFeatures.InvokeFeature("vision", "HasClipboardImage", Array())
    If VarType(hasImg) = vbString Or Not CBool(hasImg) Then
        modUiLock.Leave
        MsgBox "クリップボードに画像がありません。" & vbCrLf & _
               "画面をコピー(Win+Shift+S等)してから、もう一度押してください。", _
               vbInformation, "Nexus Agent"
        Exit Sub
    End If

    Dim prompt As String
    prompt = ReadInputCell()
    If LenB(Trim$(prompt)) = 0 Then prompt = "この画像の内容を読み取り、要点を説明してください。"
    If Len(prompt) > MAX_INPUT_CHARS Then prompt = Left$(prompt, MAX_INPUT_CHARS)   ' A3: 上限で切る

    modUI.AddChatBubble "user", ChrW(&HD83D) & ChrW(&HDCCE) & "(画像) " & prompt
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
    modUiLock.Leave
    Exit Sub

Fail:
    Err.Clear
    On Error GoTo 0
    modUiLock.Leave
End Sub

' ナビゲーション(SPA遷移)・モード/言語トグル
Public Sub OnNavChat()
    If Not modUiLock.Enter() Then Exit Sub
    modUI.GoToNexus "modApp.OnNavChat"
    modUiLock.Leave
End Sub

Public Sub OnNavHome()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modHub.EnsureHubLayout activate:=True   ' 描画と遷移を必ずセットで行う
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnNavShelf()
    If Not modUiLock.Enter() Then Exit Sub
    modUI.GoToNativeSheet modAppDef.SH_SHELF, "modApp.OnNavShelf"
    modUiLock.Leave
End Sub

Public Sub OnNavVault()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modVault.ShowVaultGallery
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnNavDash()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modDash.ShowDashboard
    On Error GoTo 0
    modUiLock.Leave
End Sub

' 🔄 画面を再描画(盲点B2/C5): ウィンドウのリサイズ・Alt+Tab復帰・マルチモニタ間の
' 移動でShapeがゴースト化/ズレたとき、ユーザーが1クリックで現在の画面を作り直す。
' 自己インストーラ配布版ではWorkbook_WindowActivate等が発火しない制約があるため、
' 自動ではなく明示的なリフレッシュ手段を提供する。アクティブな画面に応じて振り分け:
'   Nexus     → 会話履歴を壊さず視覚不変条件だけ再適用(modUI.Repaint)
'   Dashboard → データから再構築(modDash)
'   その他    → ナレッジ倉庫をデータから再構築(modVault)
Public Sub OnRefreshUI()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    Select Case ActiveSheet.Name
        Case "Nexus":     modUI.Repaint: modBoard.DrawWidget
        Case "Dashboard": modDash.ShowDashboard
        Case Else:        modVault.ShowVaultGallery
    End Select
    On Error GoTo 0
    modUiLock.Leave
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

' すぐ聞く/しっかり調べる切替。ホームと同じui_state "mode"キーを共有。
Public Sub OnToggleSpeed()
    Dim newSpeed As String
    If ReadUiState("mode", "quick") = "deep" Then
        newSpeed = "quick"
    Else
        newSpeed = "deep"
    End If
    WriteUiState "mode", newSpeed
    On Error Resume Next
    ThisWorkbook.Worksheets("Nexus").Shapes("nx_top_speed").TextFrame2.TextRange.Text = SpeedCaption()
    On Error GoTo 0
End Sub

Public Function SpeedCaption() As String
    If ReadUiState("mode", "quick") = "deep" Then
        SpeedCaption = ChrW(&HD83D) & ChrW(&HDD0D) & " しっかり調べる"
    Else
        SpeedCaption = ChrW(&H26A1) & " すぐ聞く"
    End If
End Function

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
        Case "Tiếng Việt": nextLang = "関西弁"   ' 遊び心: シークレット・オプション
        Case Else: nextLang = "日本語"
    End Select
    modConfig.SetValue "answer_language", nextLang

    On Error Resume Next
    ThisWorkbook.Worksheets("Nexus").Shapes("nx_top_lang").TextFrame2.TextRange.Text = _
        ChrW(&HD83C) & ChrW(&HDF10) & " " & nextLang & "で回答"
    On Error GoTo 0
End Sub

' ホットキー(modBootが登録/解除): Ctrl+Shift+Q=一撃召喚 / Ctrl+Enter=送信
' Ctrl+Shift+Q: どのブック・シートで作業中でも一瞬でNexusへ(軽量Activateのみ。
' LaunchNexusのフル再描画は呼ばない=速い&会話を消さない)。
Public Sub SummonNexus()
    On Error Resume Next
    ThisWorkbook.Activate
    On Error GoTo 0
    modUI.GoToNexus "modApp.SummonNexus"
    On Error Resume Next
    modUI.ParkFocus
    On Error GoTo 0
End Sub

' Ctrl+Enter: Nexus画面がアクティブな時だけ送信を発火。他のブック上では何もしない
' (副作用: 他ブックでのCtrl+Enter一括入力は本ブックを開いている間は効かなくなる。
' 稀用途とのトレードオフとしてオーナー承認済み)。
Public Sub HotSend()
    On Error Resume Next
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If ActiveSheet.Name <> "Nexus" Then Exit Sub
    On Error GoTo 0
    OnSend
End Sub

' 会話をクリアして新しい挨拶を出す(実機要望: 長い会話をリセットしたい)。
Public Sub OnClearChat()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUI.ClearChat
    modUI.AddChatBubble "ai", TimeGreeting() & " 会話をクリアしました。新しい質問をどうぞ。"
    On Error GoTo 0
    modUiLock.Leave
End Sub

' 保存して(このファイルだけ)閉じる(実機要望: 安全な終了方法が分からない)。
Public Sub OnSaveAndExit()
    If Not modUiLock.Enter() Then Exit Sub
    Dim resp As VbMsgBoxResult
    resp = MsgBox("保存してこのファイルを閉じますか?", vbYesNoCancel + vbQuestion, modAppDef.APP_NAME)
    If resp = vbCancel Then
        modUiLock.Leave
        Exit Sub
    End If
    modUiLock.Leave
    ThisWorkbook.Close SaveChanges:=(resp = vbYes)
End Sub

' 遊び心(血の通った余白): 時間帯挨拶/弱音への関西弁コンシェルジュ
Private Function TimeGreeting() As String
    Dim h As Long: h = Hour(Now)
    If h >= 5 And h < 10 Then
        TimeGreeting = "おはようございます。今日もスムーズにいきましょう。"
    ElseIf h >= 20 Or h < 5 Then
        TimeGreeting = "こんな時間までお疲れ様です。キリのいいところで切り上げてくださいね。"
    Else
        TimeGreeting = "こんにちは。"
    End If
End Function

Private Function IsTiredWords(ByVal q As String) As Boolean
    Dim t As String: t = Trim$(q)
    If Len(t) > 12 Then Exit Function   ' 長文は業務の質問(誤発動防止)
    IsTiredWords = (InStr(t, "疲れた") > 0 Or InStr(t, "つかれた") > 0 Or _
                    InStr(t, "しんどい") > 0 Or InStr(t, "眠い") > 0)
End Function

Private Function ComfortMessage() As String
    Dim pick As Long: pick = (Minute(Now) Mod 3)   ' 乱数を使わない決定的な出し分け
    Select Case pick
        Case 0
            ComfortMessage = "お疲れ様です！今日はずいぶん頑張ってはりますね。" & vbLf & _
                "温かいお茶でも飲んで、ちょっと一息つきましょか。" & ChrW(&HD83C) & ChrW(&HDF75)
        Case 1
            ComfortMessage = "ようやってはりますよ、ほんまに。" & vbLf & _
                "5分だけ肩の力抜いて、深呼吸してからまたいきましょ。" & ChrW(&H2615)
        Case Else
            ComfortMessage = "無理は禁物でっせ。仕事は明日も待ってくれます。" & vbLf & _
                "今日はここまでにして、はよ休んでくださいね。" & ChrW(&HD83C) & ChrW(&HDF19)
    End Select
End Function

' 内部ヘルパー

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
    If LenB(mGenPrevU) = 0 Then
        mGenPrevU = modState.LoadState("nexus_gen_prevu", "")
        mGenPrevA = modState.LoadState("nexus_gen_preva", "")
    End If

    Dim sys As String
    sys = "あなたはMS&ADの最上位ナレッジコンシェルジュです。プロフェッショナルで簡潔、温かく頼りになるトーンで、" & _
          modConfig.GetString("answer_language", "日本語") & "で回答してください。" & vbLf & _
          "・必ず最初の1〜2行で結論を言い切る(前置き・挨拶から始めない)。" & vbLf & _
          "・Markdown記号(#、**、`、表)は使わない(この画面では装飾されない)。" & _
          "見出しは「■ 」、箇条書きは「・」、最重要語だけ【 】で囲む。1ブロック3行以内。" & vbLf & _
          "・全体はおおむね200〜400字。言い換えの繰り返しや締めの挨拶は書かない。" & vbLf & _
          "・専門用語には短い補足を()で添え、初めて読む人にも一度で伝わる言葉を選ぶ。"

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

    ' チャット履歴シート記録(modChatLog、core層。書込失敗で死なない設計)。
    On Error Resume Next
    modChatLog.LogTurn q, resp, "general"
    On Error GoTo 0

    ' 会話履歴(新しい順;;;区切り・最大followup_max_pairsペア)
    Dim maxPairs As Long
    maxPairs = modConfig.GetLong("followup_max_pairs", 3)
    If maxPairs > 0 Then
        mGenPrevU = TrimPairs(q & IIf(LenB(mGenPrevU) > 0, ";;;" & mGenPrevU, ""), maxPairs)
        mGenPrevA = TrimPairs(modUtil.SafeLeft(resp, 2000) & IIf(LenB(mGenPrevA) > 0, ";;;" & mGenPrevA, ""), maxPairs)
    End If
    modState.SaveState "nexus_gen_prevu", mGenPrevU
    modState.SaveState "nexus_gen_preva", mGenPrevA
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
        caption = ChrW(&HD83C) & ChrW(&HDF10) & " 一般アシスタント"
    Else
        caption = ChrW(&HD83C) & ChrW(&HDFE2) & " 社内ナレッジ検索"
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
