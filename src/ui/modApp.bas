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
    ClearActions            ' 文脈アクションも消す(質問中はボタン0個=入力に集中)
    ClearConfidence         ' 信頼度バッジはnx_act_ではないので個別に消す

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
    If modLive.IsTiredWords(q) Then
        modUI.AddChatBubble "user", q
        ClearInputCell
        modUI.AddChatBubble "ai", modLive.ComfortMessage()
        modUiLock.Leave
        Exit Sub
    End If

    modUI.AddChatBubble "user", q
    ClearInputCell

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
    ' 1〜2秒で終わっているので、どの資料に答えがあるかはその時点で分かる。
    ' 預けておけば modAsk の進捗通知(SetStage/RenderSourcesPreview)がそのまま
    ' ここへ流れ込み、「考えています…」が「もう見つけた。いま書いている」へ
    ' 変わる。待ち時間そのものは1秒も縮まないが、体感は完全に別物になる。
    Dim phName As String
    phName = modUI.AddChatBubble("ai", ChrW(&HD83D) & ChrW(&HDCAD) & " 考えています…")
    On Error Resume Next
    modLive.Begin phName
    On Error GoTo Fail
    DoEvents

    Dim t0 As Double: t0 = Timer

    Dim ans As String
    Dim grounded As Boolean
    If CurrentMode() = "normal" Then
        ans = AskGeneral(q, "")
    ElseIf ShelfIsEmpty() Then
        ' 資料が1件も無いのに検索へ行くと、埋め込みAPIを1往復使ったうえで
        ' 「資料がありません」とだけ返る。いちばん遅い経路が、いちばん
        ' 価値の無い返事に着く。空だと分かっているなら聞くまでもない。
        ans = AnswerWithoutShelf(q)
    Else
        ans = modAsk.Answer(q, RagSpeed())
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
    mActiveBubble = bubbleName
    modUI.MarkActiveBubble bubbleName
    SaveTurnForRestore q, ans   ' ④記憶の継続: 次回起動時の「前回の続き」復元用に保存

    ' 積む順は 回答 → 信頼度 → 出典 → 評価。根拠を見る前に評価させない。
    ' (一般アシスタントは出典が無いので信頼度・出典は出さない=誤表示も防ぐ)
    If CurrentMode() <> "normal" Then
        DrawConfidence bubbleName
        modPeek.RenderCitations bubbleName
    End If
    DrawActions bubbleName
    If CurrentMode() <> "normal" Then
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
    modUI.AddChatBubble "ai", "エラーが発生しました。もう一度お試しください。(" & failDesc & ")"
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
    If CurrentMode() = "normal" Then
        ModeCaption = ChrW(&HD83C) & ChrW(&HDF10) & " 一般アシスタント"
    Else
        ModeCaption = ChrW(&HD83C) & ChrW(&HDFE2) & " 社内ナレッジ検索"
    End If
End Function

' OnSelectBubble - 旧UIのバブル選択(現在は未使用。文脈アクションが常に最新の
' 回答へ紐づくため対象切り替えの概念自体を廃止した)。外部からの誤呼び出しに
' 備えて残す。
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

' ❌ 違う: まずシグナルだけ1クリックで確定させ、修正入力は任意で聞く。
' 入力を先に要求すると、面倒が勝って誰も押さなくなる(旧実装の失敗)。
' 書いてくれた人にはEXPとバッジで明確に報いる。
Public Sub OnActBad()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    If Not HasTarget() Then GoTo Done

    modAsk.FeedbackRed        ' 記録+知識の穴として部内共有(ここまでは1クリック)

    Dim fixText As String
    fixText = InputBox( _
        "もしお分かりでしたら、正しい内容を教えてください。" & vbCrLf & _
        "書いていただくと EXP +20、修正が貯まると「フィードバックキング」の" & vbCrLf & _
        "バッジがもらえます。空欄のまま閉じても記録は済んでいます。", _
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
           "対象の回答(抜粋): " & modUtil.SafeLeft(TargetText(), 400) & vbLf & vbLf & _
           "正しい内容: " & fixText

    If modVault.RegisterKnowledgeText("修正ナレッジ", body, "修正,フィードバック") Then
        modStats.Bump "correction_total"
        modStats.AddExp "correction"
        modStats.EvaluateBadges
        ' 修正内容も部内へ共有する。1人の訂正が全員の訂正になる。
        modInsight.EmitCorrection modAsk.LastAnswerText(), fixText
        modSkin.ShowToast "ありがとうございます。EXP +20。次回から反映します。", "success"
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
        ans = modAsk.Answer(q, RagSpeed())
    End If

    On Error Resume Next
    modLive.Finish
    If LenB(phName) > 0 Then ThisWorkbook.Worksheets("Nexus").Shapes(phName).Delete
    On Error GoTo Fail

    Dim bubbleName As String
    bubbleName = modUI.AddChatBubble("ai", ans)
    mActiveBubble = bubbleName
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
    Err.Clear
    On Error Resume Next
    modLive.Finish   ' 実況先を必ず手放す(次のターンへ持ち越さない)
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnActResolve()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    If Not HasTarget() Then GoTo Done
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
    If Not HasTarget() Then GoTo Done

    modAsk.FeedbackUnsure     ' ここまでは1クリックで完結

    Dim hint As String
    hint = InputBox( _
        "どのあたりが引っかかりましたか?(任意)" & vbCrLf & _
        "一言でも書いていただくと EXP +20 です。空欄のまま閉じても構いません。", _
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
    DrawActions bubbleName
    On Error Resume Next
    modUI.SettleChat
    On Error GoTo 0
    modUiLock.Leave
    Exit Sub

Fail:
    Err.Clear
    On Error GoTo 0
    modUiLock.Leave
End Sub

' ----------------------------------------------------------------------------
' OnAddDocs - 入力欄の左「📁 資料を入れる」。チャットから離れずに資料を入れる。
' ----------------------------------------------------------------------------
' 画面を移動させないのが肝。資料を入れる目的は、たいてい「いま聞きたいこと
' がある」からで、別画面へ飛ばされると質問のほうを見失う。取り込みが終わったら
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
    On Error GoTo Done

    ' 取り込み処理は別シートを触ることがあるので、必ずチャットへ戻す。
    On Error Resume Next
    modUI.GoToNexus "modApp.OnAddDocs"
    On Error GoTo Done

    If after > before Then
        modUI.AddChatBubble "ai", _
            ChrW(&H2705) & " 資料を取り込みました。" & vbLf & _
            "これで、この資料の中身について「どのページに書いてあるか」まで付けてお答えできます。" & vbLf & _
            "さっそく、いま知りたいことをそのまま聞いてみてください。"
    Else
        ' 0件のときに黙って戻ると「壊れた?」になる。取り消したのか
        ' 失敗したのかを言い切らず、次の一手だけ示す。
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
        Case "Nexus":     modUI.Repaint
        Case "Dashboard": modDash.ShowDashboard
        Case modAppDef.SH_HOME: modHub.EnsureHubLayout
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

' 本棚が空のときは「答えない」のではなく「何に基づく答えかをはっきり
' させて答える」。文面と安全指示は modLive が持つ(§modLive参照)。
Private Function ShelfIsEmpty() As Boolean
    On Error Resume Next
    ShelfIsEmpty = (modShelf.TotalChunks() = 0)
    On Error GoTo 0
End Function

Private Function AnswerWithoutShelf(ByVal q As String) As String
    On Error Resume Next
    modLive.PaintStage "一般知識でお答えしています…"
    On Error GoTo 0
    AnswerWithoutShelf = modLive.EmptyShelfWrap(AskGeneral(q, modLive.EmptyShelfGuard()))
End Function

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
' extraRules: 呼び出し文脈ごとの追加制約(空可)。本棚が空のときの
' 「社内固有の数字を断定させない」制約はここから注入される。
Private Function AskGeneral(ByVal q As String, ByVal extraRules As String) As String
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
    If LenB(extraRules) > 0 Then sys = sys & vbLf & extraRules

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
    On Error Resume Next
    ThisWorkbook.Worksheets("Nexus").Shapes("nx_top_mode").TextFrame2.TextRange.Text = ModeCaption()
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
