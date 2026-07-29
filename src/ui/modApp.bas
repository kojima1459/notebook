Attribute VB_Name = "modApp"
Option Explicit

' modApp - Nexus Agent Controller。UI(modUI)とエンジン(modAsk/modGateway/
' modShelf)を結合する制御層。固定アクションバーは「選択中のAIバブル」
' (未選択時は最新)に対して発火。P2P共有パスはconfig nexus_share_pathで差替可。

Private Const MAX_INPUT_CHARS As Long = 2000         ' A3: 入力の最大文字数(超過はカット+警告)

' 連打/多重発火はmodUiLockへ一本化(Enter/Leave対で必ずLeave到達)。

' LaunchNexus - Nexus UIの起動(modBootから呼ばれる)
Public Sub LaunchNexus()
    modUI.InitUI
    RestoreLastConversation      ' ④前回の続きを薄く復元(失敗しても挨拶へ進む)
    ' 最初の1言はできるだけ短く。以前はここでモード切替の説明と
    ' ショートカット2つを並べていたが、まだ何の役にも立っていない段階で
    ' 設定の話をされても頭に入らない。操作説明は「?」に置いてある。
    modUI.AddChatBubble "ai", modLive.OpeningLine()
    On Error Resume Next         ' 以降は追加機能のフック(各自が内部で握るが二重に防護)
    modBoard.BootBoard           ' チーム連帯ボード: ビーコン発信+集計(表示はHubのタイル)
    modMentor.CollectQuestions   ' Mentor受信: 自分宛の質問を回収
    modHelp.EnsureHelpButton     ' ヘルプ(?)ボタン

    ' 起動処理中に届いていた「ありがとう」を、ここで会話として伝える。
    ' このアプリで唯一、機械ではなく人が相手にいる瞬間なので、統計タイルの
    ' 数字ではなく、名前と資料名のまま出す。
    Dim thx As String
    thx = modP2P.NoticeText()
    If LenB(thx) > 0 Then modUI.AddChatBubble "ai", thx

    ' 押すだけで試せる質問を並べる。初手で「何を聞こう」と考えさせない。
    modStarter.Draw

    modTour.StartTourIfFirstRun  ' 初回オンボーディングツアー
    On Error GoTo 0
End Sub

' 2026-07-26 再設計: 質問テンプレチップ/ナレッジガチャはサイドバー廃止に伴い
' Hub画面(modHub.DrawExtras)へ移した。チャット画面は会話だけを担う。

' ④会話の記憶: 直近2往復をui_stateへ保存し、次回起動時に薄く復元する。
' 区切りはAskGeneral履歴と同じ";;;"(質問/回答に含まれる場合は改行1個に置換して保護)。
Private Sub SaveTurnForRestore(ByVal q As String, ByVal ans As String)
    On Error Resume Next
    Dim u As String, a As String
    u = Replace(modUtil.SafeLeft(q, 300), ";;;", " ")
    a = Replace(modUtil.SafeLeft(ans, 700), ";;;", " ")
    Dim prevU As String: prevU = modState.LoadState("nexus_hist_u", "")
    Dim prevA As String: prevA = modState.LoadState("nexus_hist_a", "")
    modState.SaveState "nexus_hist_u", modAppState.TrimPairs(u & IIf(LenB(prevU) > 0, ";;;" & prevU, ""), 2)
    modState.SaveState "nexus_hist_a", modAppState.TrimPairs(a & IIf(LenB(prevA) > 0, ";;;" & prevA, ""), 2)
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
    ClearActions            ' 文脈アクションも消す(質問中はボタン0個=入力に集中)
    ClearConfidence         ' 信頼度バッジはnx_act_ではないので個別に消す
    modStarter.Clear        ' 質問例も消す(会話が始まったら役目は終わり)

    Dim q As String
    q = modAppState.ReadInputCell()
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
    If modLive.IsTiredWords(q) Then
        modUI.AddChatBubble "user", q
        modAppState.ClearInputCell
        modUI.AddChatBubble "ai", modLive.ComfortMessage()
        modUiLock.Leave
        Exit Sub
    End If

    modUI.AddChatBubble "user", q
    modAppState.ClearInputCell

    ' 逆質問の途中なら、返事(番号選択 or 書き直し)を元の質問と合成して
    ' 完全な質問文に組み立て直す。利用者は番号を打つだけでよい。
    On Error Resume Next
    If modClarify.HasPending() Then q = modClarify.MergeAnswer(q)
    On Error GoTo Fail

    ' 体感速度ハック: 待ち時間の無反応(壊れた?)を防ぐため、考え中バブルを即時表示。
    ' 回答が来たら削除して本物を追加する(in-place置換はバブル高さ管理と衝突するため
    ' 削除→追加方式。小さな余白が残るだけで崩れない)。
    '
    ' さらに、このバブルを modUIMain へ「実況先」として預ける。検索は最初の
    ' 1～2秒で終わっているので、どの資料に答えがあるかはその時点で分かる。
    ' 預けておけば modAsk の進捗通知(SetStage/RenderSourcesPreview)がそのまま
    ' ここへ流れ込み、「考えています…」が「もう見つけた。いま書いている」へ
    ' 変わる。待ち時間そのものは1秒も縮まないが、体感は完全に別物になる。
    Dim phName As String
    phName = modUI.AddChatBubble("ai", ChrW(&HD83D) & ChrW(&HDCAD) & " 考えています…")
    On Error Resume Next
    modLive.Begin phName
    On Error GoTo Fail
    DoEvents

    ' 2026-07-28(レビュー M-26): モードと速さは【送信の入口で一度だけ】
    ' 確定させ、以降はこのローカル変数だけを見る。
    ' 従来は OnSend の中で CurrentMode()/RagSpeed() を都度読み直していたため、
    ' LLMの応答待ち(数十秒)の間にヘッダーのトグルを押されると、
    ' 「一般回答の下に前回RAGの出典チップが出る」という【出典の誤提示】が
    ' 起きた。回答は根拠と対で意味を持つので、これは表示崩れでは済まない。
    Dim sendMode As String: sendMode = modAppState.CurrentMode()
    Dim sendSpeed As String: sendSpeed = modAppState.RagSpeed()

    ' 入念は数分かかる。始まる前に「何をするか・どれくらいかかるか」を
    ' 必ず出す。黙って数分止まると、利用者は固まったと判断して閉じる。
    On Error Resume Next
    If sendSpeed = "thorough" Then
        modLive.PaintStage "入念に調べます。多方向から検索して、資料と1行ずつ照合します…"
    End If
    On Error GoTo Fail

    Dim t0 As Double: t0 = Timer

    Dim ans As String
    Dim grounded As Boolean
    If sendMode = "normal" Then
        ans = modAppState.AskGeneral(q, "")
    ElseIf modAppState.ShelfIsEmpty() Then
        ' 資料が1件も無いのに検索へ行くと、埋め込みAPIを1往復使ったうえで
        ' 「資料がありません」とだけ返る。いちばん遅い経路が、いちばん
        ' 価値の無い返事に着く。空だと分かっているなら聞くまでもない。
        ans = modAppState.AnswerWithoutShelf(q)
    Else
        ans = modAsk.Answer(q, sendSpeed)
        grounded = True
    End If

    Dim secs As Double: secs = Timer - t0

    On Error Resume Next
    modLive.Finish
    If LenB(phName) > 0 Then ThisWorkbook.Worksheets("Nexus").Shapes(phName).Delete
    On Error GoTo Fail

    ' 速さと調べた量は、言わなければ伝わらない。「15秒待たされた」と
    ' 「3冊の資料から12秒で根拠付き」は同じ時間の別の体験になる。
    Dim bubbleName As String
    bubbleName = modUI.AddChatBubble("ai", ans & vbCr & modLive.Footer(secs, grounded))
    modLive.StyleFooter bubbleName
    modAppState.SetActiveBubble bubbleName
    modUI.MarkActiveBubble bubbleName
    SaveTurnForRestore q, ans   ' ④記憶の継続: 次回起動時の「前回の続き」復元用に保存

    ' 積む順は 回答 → 信頼度 → 出典 → 評価。根拠を見る前に評価させない。
    ' (一般アシスタントは出典が無いので信頼度・出典は出さない=誤表示も防ぐ)
    If sendMode <> "normal" Then
        DrawConfidence bubbleName
        modPeek.RenderCitations bubbleName
    End If
    DrawActions bubbleName
    If sendMode <> "normal" Then
        modMentor.OfferMentor bubbleName   ' Mentor: 専門家ボタン(失敗しても出ないだけ=安全弁内蔵)
    End If
    On Error Resume Next
    modUI.SettleChat        ' 次のバブルがアクション/出典に重ならないよう下端を確定
    On Error GoTo Fail

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
    modLive.Finish   ' 実況先を必ず手放す(次のターンへ持ち越さない)
    modLog.LogError "E0602", "modApp.OnSend", failDesc
    ' 2026-07-28(レビュー L-22): 失敗したら入力欄へ書き戻す。
    ' 送信直後に入力欄をクリアする作りなので、長文を書いて送って落ちると
    ' 打った文章が丸ごと消えていた。「もう一度お試しください」と言われても、
    ' もう一度打ち直すところからになる。
    If LenB(q) > 0 Then modAppState.RestoreInputCell q
    modUI.AddChatBubble "ai", "エラーが発生しました。入力欄に文章を戻しましたので、" & _
        "もう一度お試しください。(" & failDesc & ")"
    On Error GoTo 0
    modUiLock.Leave
End Sub

' 文脈アクションの描画/消去。対象は常に「最新のAI回答」なので、描いたバブル名を
' mActiveBubbleにも記録し、各OnAct*が同じものを見るようにする。
Private Sub DrawActions(ByVal bubbleName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub
    modUINexusDraw.DrawContextActions ws, bubbleName
    On Error GoTo 0
End Sub

Private Sub ClearConfidence()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub
    modUINexusDraw.ClearConfidence ws
    On Error GoTo 0
End Sub

Private Sub DrawConfidence(ByVal bubbleName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub
    modUINexusDraw.DrawConfidence ws, bubbleName
    On Error GoTo 0
End Sub

Private Sub ClearActions()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub
    modUINexusDraw.ClearContextActions ws
    On Error GoTo 0
End Sub

' モードボタンの表示文字列(ヘッダー描画とトグルの両方が使う単一情報源)。
Public Function ModeCaption() As String
    If modAppState.CurrentMode() = "normal" Then
        ModeCaption = ChrW(&HD83C) & ChrW(&HDF10) & " 一般アシスタント"
    Else
        ModeCaption = ChrW(&HD83C) & ChrW(&HDFE2) & " 社内ナレッジ検索"
    End If
End Function


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


' ❌ 違う: まずシグナルだけ1クリックで確定させ、修正入力は任意で聞く。
' 入力を先に要求すると、面倒が勝って誰も押さなくなる(旧実装の失敗)。
' 書いてくれた人にはEXPとバッジで明確に報いる。
Public Sub OnActBad()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    If Not modAppState.HasTarget() Then GoTo Done

    modAsk.FeedbackRed        ' 記録+知識の穴として部内共有(ここまでは1クリック)

    Dim fixText As String
    ' 「EXP +20」「フィードバックキングのバッジ」で釣るのをやめた。
    ' 誤った回答を訂正するのは、客先事故を1件止めうる専門職の仕事であって、
    ' ポイントで報いる対象ではない。42歳の課長がこの文面を横から見たとき、
    ' 「若手向けのおもちゃ」と判定された時点で、その部署では終わる。
    ' 本当の見返り(自分の訂正が次から全員の答えになる)は既に実装済み
    ' (modInsight.EmitCorrection)。それをそのまま書けばいい。
    fixText = InputBox( _
        "もしお分かりでしたら、正しい内容を教えてください。" & vbCrLf & _
        "書いていただいた内容は、次から同じ質問をした人全員の答えになります。" & vbCrLf & _
        "空欄のまま閉じても、記録は済んでいます。", _
        "Nexus Agent - 正しい内容を教える")
    If LenB(Trim$(fixText)) = 0 Then GoTo Done

    RecordCorrection fixText
Done:
    modUiLock.Leave
End Sub

' 修正入力の共通処理(👎/🤔の両方から呼ぶ)。ナレッジ化+EXP+カウント。
Private Sub RecordCorrection(ByVal fixText As String)
    On Error Resume Next
    Dim body As String
    body = "【修正ナレッジ】" & vbLf & _
           "対象の回答(抜粋): " & modUtil.SafeLeft(modAppState.TargetText(), 400) & vbLf & vbLf & _
           "正しい内容: " & fixText

    If modVault.RegisterKnowledgeText("修正ナレッジ", body, "修正,フィードバック") Then
        modStats.Bump "correction_total"
        modStats.AddExp "correction"
        modStats.EvaluateBadges
        ' 修正内容も部内へ共有する。1人の訂正が全員の訂正になる。
        modInsight.EmitCorrection modAsk.LastAnswerText(), fixText
        modSkin.ShowToast "ありがとうございます。次に同じ質問をした人から、この内容で答えます。", "success"
    Else
        MsgBox "学習の保存に失敗しました。マイ本棚の一覧をご確認ください。", vbExclamation, "Nexus Agent"
    End If
    On Error GoTo 0
End Sub

Public Sub OnActDrill()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Fail
    modPeek.HideCitations   ' 前回の出典チップ/ポップアップを消す
    modMentor.ClearMentor   ' Mentorボタンも掃除(安全弁内蔵)
    ClearActions
    ClearConfidence

    Dim q As String
    q = InputBox("さらに深掘りしたい内容を入力してください。" & vbCrLf & _
                 "(直前までの会話を踏まえて回答します)", "Nexus Agent - 深掘り")
    If LenB(Trim$(q)) = 0 Then
        modUiLock.Leave
        Exit Sub
    End If
    If Len(q) > MAX_INPUT_CHARS Then q = Left$(q, MAX_INPUT_CHARS)   ' A3: 上限で切る

    modUI.AddChatBubble "user", ChrW(&HD83D) & ChrW(&HDD0D) & " " & q

    ' 深掘りも本流の送信と同じ待ち時間が発生する。同じように実況する。
    Dim phName As String
    phName = modUI.AddChatBubble("ai", ChrW(&HD83D) & ChrW(&HDCAD) & " 考えています…")
    On Error Resume Next
    modLive.Begin phName
    On Error GoTo Fail
    DoEvents

    Dim ans As String
    If modAsk.CanFollowup() Then
        modAsk.AskFollowup q
        ans = modAsk.LastAnswerText()
    Else
        ans = modAsk.Answer(q, modAppState.RagSpeed())
    End If

    On Error Resume Next
    modLive.Finish
    If LenB(phName) > 0 Then ThisWorkbook.Worksheets("Nexus").Shapes(phName).Delete
    On Error GoTo Fail

    Dim bubbleName As String
    bubbleName = modUI.AddChatBubble("ai", ans)
    modAppState.SetActiveBubble bubbleName
    modUI.MarkActiveBubble bubbleName
    DrawConfidence bubbleName            ' 信頼度 → 出典 → 評価 の順に積む
    modPeek.RenderCitations bubbleName   ' Peek View: 深掘り回答の出典チップ
    DrawActions bubbleName               ' 文脈アクション(根拠より下)
    modMentor.OfferMentor bubbleName     ' Mentor: 専門家ボタン(安全弁内蔵)
    On Error Resume Next
    modUI.SettleChat
    On Error GoTo 0
    modUiLock.Leave
    Exit Sub

Fail:
    ' 2026-07-28(レビュー L-21): 無言で終わらない。
    ' ここは Err.Clear してログも通知も出さずに Leave していたため、
    ' 深掘りが失敗すると【押したのに何も起きない】だけになっていた。
    ' 利用者はボタンが壊れたと判断し、二度と押さない。
    ' OnSend の Fail と同じく、記録とエラーバブルの両方を出す。
    Dim drillDesc As String: drillDesc = Err.Description
    Err.Clear
    On Error Resume Next
    modLive.Finish   ' 実況先を必ず手放す(次のターンへ持ち越さない)
    modLog.LogError "E0602", "modApp.OnActDrill", drillDesc
    modUI.AddChatBubble "ai", "深掘りに失敗しました。もう一度お試しください。(" & drillDesc & ")"
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnActResolve()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    If Not modAppState.HasTarget() Then GoTo Done
    modAsk.FeedbackGreen   ' selfsolve_total加算+多重防止は既存ガードに従う
    ' 統計の表示先はHubの統計タイルへ移した。加算直後に描き直して
    ' 「押しても0のまま」を防ぐ(activate:=Falseなので画面は移動しない)。
    On Error Resume Next
    modHub.EnsureHubLayout
    On Error GoTo Done
Done:
    modUiLock.Leave
End Sub

' 🤔 微妙: 入力を一切求めない1クリック評価。「正しいか分からないから何も
' 押さない」を無くすための逃げ道であり、同時に「資料が足りない領域」の
' シグナルとして共有される(modAsk.FeedbackUnsure)。
Public Sub OnActUnsure()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    If Not modAppState.HasTarget() Then GoTo Done

    modAsk.FeedbackUnsure     ' ここまでは1クリックで完結

    Dim hint As String
    hint = InputBox( _
        "どのあたりが引っかかりましたか?(任意)" & vbCrLf & _
        "一言でも書いていただければ、同じ引っかかりを次の人が踏まずに済みます。" & vbCrLf & _
        "空欄のまま閉じても構いません。", _
        "Nexus Agent - どこが気になりましたか")
    If LenB(Trim$(hint)) = 0 Then GoTo Done

    RecordCorrection hint
Done:
    modUiLock.Leave
End Sub

' 「🏘 本社照会」は削除した(2026-07-27)。押しても常に
' 「準備中です」としか返らないボタンが、すべての回答の下に永久に並んでいた。
' 動かないものが1つ混じっているだけで、利用者は「他も見せかけかもしれない」と
' 学習する。出せる目処が立つまでは、出さないほうが信用は減らない。

' ホームのOnOpenWordButtonと同じInputBoxを挟む(以前は指示文なし=""固定だった)。
Public Sub OnActWord()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    If Not modAppState.HasTarget() Then GoTo Done
    Dim answerBody As String
    answerBody = modAppState.TargetText()

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
    If Not modAppState.HasTarget() Then GoTo Done
    Dim t As String: t = modAppState.TargetText()
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

' 📎 クリップボード画像チャット(OnAttachImage)は削除した(2026-07-27)。
' 入力欄左の一等地を「📁 資料を入れる」に譲った時点でボタンが無くなり、
' どこからも呼べない死んだ経路になっていた。画像の取込自体はナレッジ画面の
' 「📸 スクショ取込」(modUIShelf.OnIngestScreenshot)に生きている。

' ----------------------------------------------------------------------------
' OnAddDocs - 入力欄の左「📁 資料を入れる」。チャットから離れずに資料を入れる。
' ----------------------------------------------------------------------------
' 画面を移動させないのが肝。資料を入れる目的はたいてい「いま聞きたいことが
' ある」からで、別画面へ飛ばされると質問のほうを見失う。取り込みが終わったら
' 会話の中に結果を出し、そのまま次の一言を打てる状態に戻す。
Public Sub OnAddDocs()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done

    Dim before As Long
    On Error Resume Next
    before = modShelf.TotalChunks()
    On Error GoTo Done

    modShelf.AddFilesViaDialog

    Dim after As Long
    On Error Resume Next
    after = modShelf.TotalChunks()
    modUI.GoToNexus "modApp.OnAddDocs"   ' 取込は別シートを触るので必ず戻す
    On Error GoTo Done

    If after > before Then
        modUI.AddChatBubble "ai", _
            ChrW(&H2705) & " 資料を取り込みました。" & vbLf & _
            "これで、この資料の中身について「どのページに書いてあるか」まで付けてお答えできます。" & vbLf & _
            "さっそく、いま知りたいことをそのまま聞いてみてください。"
    Else
        ' 0件のときに黙って戻ると「壊れた?」になる。取り消したのか失敗したのかを
        ' 言い切らず、次の一手だけ示す。
        modUI.AddChatBubble "ai", _
            "資料は追加されませんでした。" & vbLf & _
            "対応しているのは PDF / Word / Excel / テキスト です。" & vbLf & _
            "資料が無くても、一般的な内容ならこのままお答えできます。"
    End If
    On Error Resume Next
    modUI.SettleChat
    On Error GoTo Done
Done:
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
        Case "Nexus":     modUI.Repaint
        Case "Dashboard": modDash.ShowDashboard
        Case modAppDef.SH_HOME: modHub.EnsureHubLayout
        Case Else:        modVault.ShowVaultGallery
    End Select
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnToggleMode()
    ' 2026-07-28(レビュー M-26): 送信処理中はモード類を切り替えさせない。
    ' このトグル4本だけ modUiLock を通っておらず、LLMの応答待ち中に
    ' 押せてしまうため、待っている回答と表示の前提がずれる。
    ' ロックは取らない(この操作自体は一瞬で終わる)。busy かどうかだけ見る。
    If modUiLock.IsBusy() Then
        On Error Resume Next
        modSkin.ShowToast "回答の生成中です。終わってから切り替えてください。", "info"
        On Error GoTo 0
        Exit Sub
    End If

    Dim newMode As String
    If modAppState.CurrentMode() = "normal" Then
        newMode = "rag"
    Else
        newMode = "normal"
    End If
    modAppState.WriteUiState modAppState.MODE_KEY, newMode
    modAppState.UpdateModeButton
End Sub

' すぐ聞く/しっかり調べる切替。ホームと同じui_state "mode"キーを共有。
Public Sub OnToggleSpeed()
    ' 2026-07-28(レビュー M-26): 送信処理中はモード類を切り替えさせない。
    ' このトグル4本だけ modUiLock を通っておらず、LLMの応答待ち中に
    ' 押せてしまうため、待っている回答と表示の前提がずれる。
    ' ロックは取らない(この操作自体は一瞬で終わる)。busy かどうかだけ見る。
    If modUiLock.IsBusy() Then
        On Error Resume Next
        modSkin.ShowToast "回答の生成中です。終わってから切り替えてください。", "info"
        On Error GoTo 0
        Exit Sub
    End If

    ' すぐ聞く → しっかり調べる → 入念に調べる → すぐ聞く の巡回。
    ' 押すたびに何が変わるかをトーストで必ず出す。モード名だけでは
    ' 「押したら遅くなった」としか分からず、選ぶ理由が伝わらない。
    Dim newSpeed As String
    newSpeed = modMode.NextMode(modAppState.ReadUiState("mode", "quick"))
    modAppState.WriteUiState "mode", newSpeed
    On Error Resume Next
    ThisWorkbook.Worksheets("Nexus").Shapes("nx_top_speed").TextFrame2.TextRange.Text = _
        modMode.Caption(newSpeed)
    modSkin.ShowToast modMode.Caption(newSpeed) & " ： " & modMode.Description(newSpeed), "info"
    On Error GoTo 0
End Sub

Public Function SpeedCaption() As String
    SpeedCaption = modMode.Caption(modAppState.ReadUiState("mode", "quick"))
End Function

' 回答言語の巡回切替(日本語→English→中文→Tiếng Việt)。
' answer_languageは既存プロンプト(modPrompts)がそのまま使用する。
Public Sub OnLangCycle()
    ' 2026-07-28(レビュー M-26): 送信処理中はモード類を切り替えさせない。
    ' このトグル4本だけ modUiLock を通っておらず、LLMの応答待ち中に
    ' 押せてしまうため、待っている回答と表示の前提がずれる。
    ' ロックは取らない(この操作自体は一瞬で終わる)。busy かどうかだけ見る。
    If modUiLock.IsBusy() Then
        On Error Resume Next
        modSkin.ShowToast "回答の生成中です。終わってから切り替えてください。", "info"
        On Error GoTo 0
        Exit Sub
    End If

    Dim cur As String
    cur = modConfig.GetString("answer_language", "日本語")

    ' 2026-07-28(レビュー H-16): "Tiếng Việt" をリテラルで書けない。
    ' このブックは起動のたび vba_src のソースを VBE へ注入するが、VBE は
    ' コードを CP932 で保持するため、CP932 に無い ế/ệ はリテラル "?" として
    ' 保存される。結果、ボタン表示も config に保存される値も "Ti?ng Vi?t" に
    ' なり、LLM への言語指定まで壊れていた(vbaProject.bin の実バイトで確認済み)。
    ' ChrW で組み立てれば、ソースは ASCII のまま実行時に正しい文字になる。
    Dim viet As String
    viet = "Ti" & ChrW(&H1EBF) & "ng Vi" & ChrW(&H1EC7) & "t"

    Dim nextLang As String
    Select Case cur
        Case "日本語": nextLang = "English"
        Case "English": nextLang = "中文"
        Case "中文": nextLang = viet
        ' 既に "Ti?ng Vi?t" で保存されてしまった config からの移行も拾う。
        Case viet, "Ti?ng Vi?t": nextLang = "関西弁"   ' 遊び心: シークレット・オプション
        Case Else: nextLang = "日本語"
    End Select
    modConfig.SetValue "answer_language", nextLang

    On Error Resume Next
    ThisWorkbook.Worksheets("Nexus").Shapes("nx_top_lang").TextFrame2.TextRange.Text = _
        ChrW(&HD83C) & ChrW(&HDF10) & " " & nextLang
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
    modClarify.ClearPending
    ' 2026-07-28(レビュー L-18): クリアで消えていなかったものを片付ける。
    '   ・出典チップ / 専門家ボタン … 消した会話の下にボタンだけ残っていた
    '   ・続けて質問の履歴 / 復元用の直近ターン … 残っていると、再起動時に
    '     「クリアしたはずの会話」が復元される(利用者から見れば消えていない)
    ' 「消した」と言った以上は、次に開いたときも消えていなければならない。
    modPeek.HideCitations
    modMentor.ClearMentor
    ClearActions
    ClearConfidence
    modState.SaveState "nexus_ask_prevu", ""
    modState.SaveState "nexus_ask_preva", ""
    modState.SaveState "nexus_hist_u", ""
    modState.SaveState "nexus_hist_a", ""
    modUI.AddChatBubble "ai", modLive.TimeGreeting() & " 会話をクリアしました。新しい質問をどうぞ。"
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

' 内部ヘルパー

