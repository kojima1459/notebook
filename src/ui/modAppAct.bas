Attribute VB_Name = "modAppAct"
Option Explicit

' ============================================================================
' modAppAct - 回答に付く「文脈アクション行」と信頼度バッジの描画/消去、および
'   そのボタン(解決した/微妙/違う/深掘り/コピー/Word)のハンドラ。
' ----------------------------------------------------------------------------
' 2026-07-31(R11-F1): modApp が30,000字上限まで残り160字となり、バグ修正
'   1件も入らない状態だったため、憲章§4-6に基づき分割した。
'   切り口は「1回の質問→回答を作る流れ(modApp: OnSend/取込/ナビ/終了)」と
'   「出来上がった回答に対して利用者が起こす操作(本モジュール)」。
'   後者は modAppState.HasTarget()/TargetBubble() が対象バブルの唯一の情報源で、
'   modApp と共有するモジュールレベル状態を持たないため縫い目が残らない。
'   描画ラッパー4本(DrawActions/ClearActions/DrawConfidence/ClearConfidence)は
'   この行を出し入れする道具なので同時に移し、modApp 側から Public で呼ぶ。
' ============================================================================

' 文脈アクションの描画/消去。対象は常に「最新のAI回答」なので、描いたバブル名を
' mActiveBubbleにも記録し、各OnAct*が同じものを見るようにする。
Public Sub DrawActions(ByVal bubbleName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub
    modUINexusDraw.DrawContextActions ws, bubbleName
    On Error GoTo 0
End Sub

Public Sub ClearConfidence()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub
    modUINexusDraw.ClearConfidence ws
    On Error GoTo 0
End Sub

Public Sub DrawConfidence(ByVal bubbleName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub
    modUINexusDraw.DrawConfidence ws, bubbleName
    On Error GoTo 0
End Sub

Public Sub ClearActions()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub
    modUINexusDraw.ClearContextActions ws
    On Error GoTo 0
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
    ' (modInsightIo.EmitCorrection)。それをそのまま書けばいい。
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
        modInsightIo.EmitCorrection modAsk.LastAnswerText(), fixText
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
    If Len(q) > modApp.MAX_INPUT_CHARS Then q = Left$(q, modApp.MAX_INPUT_CHARS)   ' A3: 上限で切る

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
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup13
FailCleanup13:
    Err.Clear
    On Error Resume Next
    modLive.Finish   ' 実況先を必ず手放す(次のターンへ持ち越さない)
    modLog.LogError "E0602", "modAppAct.OnActDrill", drillDesc
    modUI.AddChatBubble "ai", "深掘りに失敗しました。もう一度お試しください。(" & drillDesc & ")"
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnActResolve()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    If Not modAppState.HasTarget() Then GoTo Done
    modAsk.FeedbackGreen   ' selfsolve_total加算+多重防止は既存ガードに従う
    ' 2026-07-31(レビュー R8 F8): 加算した「今日の節約時間」を、その場で
    ' 共有フォルダのビーコンへ反映する。従来は起動時(BootBoard)にしか
    ' 発信しておらず、「解決した」を押すのは起動よりずっと後なので、全員のビーコンが
    ' 「今日 0分」のまま置かれていた。結果、Hubの「みんなの節約(今日)」が
    ' 構造的にほぼ常に0で、共有そのものが動いていないように見えていた。
    ' 中で共有到達ガードと10分スロットルが効くので、連打しても重くならない。
    On Error Resume Next
    modBoard.PublishBeacon
    On Error GoTo Done
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
