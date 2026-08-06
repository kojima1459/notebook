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
'
' 2026-08-03(R13-6a): 「続けて質問/深掘り」の armed followup 一式もここに置く
'   (modUIMain は残383字で呼び出し1行しか入らないため。憲章§4-6)。
' ============================================================================

' 「続きの質問」チップのShape名。nx_ 接頭辞なので、Nexus画面の全面再描画
' (modUI の nx_ 一括削除)で自動的に消える=取り残しが構造的に起きない。
Private Const FCHIP_NAME As String = "nx_fchip"

' armed状態(セッション変数)。State Lossで消えたら「武装していない」に倒れる。
Private mArmedFollowup As Boolean

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
        modAppDef.APP_NAME & " - 正しい内容を教える")
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
        ' ただし発信してよいのは「本棚の資料を根拠に答えたターン」だけ
        ' (真理表は modMode.ShouldEmitInsight、窓口は modAsk.CanShareInsight)。
        ' 一般アシスタントで答えた直後にここを撃つと、【前のRAG質問の回答】が
        ' 「訂正の対象になった回答」として部内へ配られる(実機第3報 RC2 と
        ' 同型の誤爆)。修正ナレッジの個人保存・EXPは判定と無関係に必ず行う。
        If modAsk.CanShareInsight() Then
            modInsightIo.EmitCorrection modAsk.LastAnswerText(), fixText
        End If
        modSkin.ShowToast "ありがとうございます。次に同じ質問をした人から、この内容で答えます。", "success"
    Else
        MsgBox "学習の保存に失敗しました。マイ本棚の一覧をご確認ください。", vbExclamation, modAppDef.APP_NAME
    End If
    On Error GoTo 0
End Sub

' 🔍 深掘り: 2026-08-03(R13-6a)から InputBox をやめ「armed followup」にした。
' 旧実装は幅を制御できない素のInputBoxで長文を書かせ、しかも押した瞬間に
' 「その時のモード」が暗黙で確定していた(実機第2報 RC7)。
' いまは「次の送信を続きの質問として扱う」印を立てるだけにして、入力は
' 既存の幅広い入力欄(nx_input)で受け、送信ボタンの1本道へ合流させる。
' 実際の質問処理は modApp.OnSend が行う(=経路が1つになり、実況・出典・
' 評価ボタンの積み方が深掘りだけ別実装、という重複も消える)。
Public Sub OnActDrill()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Fail
    ' R20-2a: ゲート判定を【先消しより前】にする。従来は出典/Mentor/アクション/
    ' 信頼度を無条件で消してからArmFollowupのゲートに落ちていたため、失敗時に
    ' ボタン列が消えたまま何も起きない(「押したのに何も起きない」の再発)。
    If Not CanArmFollowup() Then
        modSkin.ShowToast "まず質問して回答を受け取ってから使ってください。", "info"
        modUiLock.Leave
        Exit Sub
    End If
    modPeek.HideCitations   ' 前回の出典チップ/ポップアップを消す
    modMentor.ClearMentor   ' Mentorボタンも掃除(安全弁内蔵)
    ClearActions
    ClearConfidence
    ArmFollowup
    modUiLock.Leave
    Exit Sub

Fail:
    ' 2026-07-28(レビュー L-21): 無言で終わらない。押したのに何も起きない、を作らない。
    Dim drillDesc As String: drillDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup13
FailCleanup13:
    Err.Clear
    On Error Resume Next
    modLog.LogError "E0602", "modAppAct.OnActDrill", drillDesc
    modUI.AddChatBubble "ai", "深掘りの準備に失敗しました。もう一度お試しください。(" & drillDesc & ")"
    On Error GoTo 0
    modUiLock.Leave
End Sub

' ============================================================================
' armed followup(2026-08-03 R13-6a/6b)
' ----------------------------------------------------------------------------
' 「続けて質問」「深掘り」を押すと、
'   (i)   セッション変数に「次の送信は続きの質問」という印を立て、
'   (ii)  入力欄(nx_input)へフォーカスし、
'   (iii) 入力欄の下に小さなチップを1枚出す。
' この状態で送信すると modApp.OnSend が modAsk.AskFollowup へ回す。
' チップをクリックすると解除(=押せるものは必ず反応する。憲章§3-1)。
'
' 状態はモジュール変数なので、VBAのState Lossで自然に消える。消えたときに
' チップだけが画面に残らないよう、再描画(RedrawFollowupChip)は「必ず消して
' から、armedのときだけ描く」形にし、Nexus画面の描画(modUI.InitUI)からも呼ぶ。
'
' 6b(エフォートの明示): チップに出すモード名は【押した時点】のもの。
' 実際に使われるのは【送信時点】のトグル値(modAsk.AskFollowup が ui_state を
' 読む既存の作り)。つまり利用者はチップを出したままモードを切り替えてから
' 送れる ―― これが「深掘りの力の入れ方を選べる」の実装であり、
' そのためにチップへ「モード: 〜」を書いて選択肢の存在を見せている。
' 宣言(FCHIP_NAME / mArmedFollowup)はモジュール先頭にある。
' ============================================================================

' R20-2b: ゲートのモード分岐。normalは一般アシスタントの会話履歴
' (modAppState.HasGeneralMemory)、それ以外はRAGの履歴(modAsk.CanFollowup)を見る。
' modAskは触らずmodAppStateへ新設した窓口を読む。判定式そのもの(どちらの
' 真偽値を見るか)は純関数GateUsesGeneralHistoryへ出し、LOテストで
' 「normal/rag × 履歴有無」の4通りを固定する(実際の真偽値はExcel状態を
' 読むためLOでは再現できないが、注入した2値からの決定規則は再現できる)。
Public Function GateUsesGeneralHistory(ByVal chatMode As String, _
                                       ByVal hasGeneralMemory As Boolean, _
                                       ByVal hasRagMemory As Boolean) As Boolean
    If chatMode = "normal" Then
        GateUsesGeneralHistory = hasGeneralMemory
    Else
        GateUsesGeneralHistory = hasRagMemory
    End If
End Function

Private Function CanArmFollowup() As Boolean
    CanArmFollowup = GateUsesGeneralHistory(modAppState.CurrentMode(), _
        modAppState.HasGeneralMemory(), modAsk.CanFollowup())
End Function

' 「続きの質問」を武装する(modUIMain.OnFollowupButton / OnActDrill の共通実体)。
Public Sub ArmFollowup()
    If Not CanArmFollowup() Then
        ' R20-2a: MsgBoxをやめToastへ(ボタン消失の再発防止と歩調を揃える)。
        modSkin.ShowToast "まず質問して回答を受け取ってから使ってください。", "info"
        Exit Sub
    End If

    mArmedFollowup = True
    On Error Resume Next
    modUI.GoToNexus "modAppAct.ArmFollowup"   ' ホーム画面から押されたときはチャットへ
    RedrawFollowupChip
    modUI.ParkFocus                           ' 入力欄(nx_input)へフォーカス
    On Error GoTo 0
End Sub

' 送信時に1回だけTrueを返し、同時に解除+チップを消す(modApp.OnSend から)。
Public Function ConsumeArmedFollowup() As Boolean
    If Not mArmedFollowup Then Exit Function
    mArmedFollowup = False
    RedrawFollowupChip
    ConsumeArmedFollowup = True
End Function

' チップ([×])のクリック。Shape.OnAction の宛先なので Public。
Public Sub OnFollowupChipOff()
    mArmedFollowup = False
    RedrawFollowupChip
End Sub

' チップの再描画=掃除。armedでなければ消すだけ(State Lossで印だけ消え、
' チップが残ったまま…を構造的に起こさない)。
Public Sub RedrawFollowupChip()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub
    ws.Shapes(FCHIP_NAME).Delete
    Err.Clear
    If mArmedFollowup Then DrawFollowupChip ws
    On Error GoTo 0
End Sub

Private Sub DrawFollowupChip(ByVal ws As Worksheet)
    On Error Resume Next
    Dim rowIdx As Long: rowIdx = modUINexusDraw.INPUT_ROW + 1
    Dim r As String: r = CStr(rowIdx)
    Dim anchor As Range: Set anchor = ws.Range("C" & r)
    If anchor Is Nothing Then Exit Sub

    ' R13 L-batch(チップの幾何): チップはヒント行(入力欄の1つ下)に
    ' ぴったり収める。従来は「行の1pt上・高さ15固定」だったが、この行の
    ' 高さは12ptなので下へ約2pt はみ出し、直下の会話領域の最初のバブルに
    ' 重なっていた。行の実寸から取れば、行高を変えても重なりは起きない。
    Dim chipH As Double: chipH = ws.Rows(rowIdx).Height
    If chipH < 10 Then chipH = 10     ' 行が極端に低いときの下限(文字が潰れない高さ)

    Dim chip As Shape
    Set chip = ws.Shapes.AddShape(5, anchor.Left, ws.Rows(rowIdx).Top, _
                                  ws.Range("C" & r & ":K" & r).Width, chipH)
    If chip Is Nothing Then Exit Sub
    chip.Name = FCHIP_NAME
    chip.Placement = 3
    chip.Adjustments(1) = 0.4
    chip.Line.Visible = 0
    chip.Fill.ForeColor.RGB = modUI.UiColor("accent")
    With chip.TextFrame2
        .TextRange.Text = ChrW(&H21B3) & " 続きの質問(前回の会話を引き継ぐ)  [" & _
                          ChrW(&HD7) & "]  モード: " & modApp.SpeedCaption()
        .TextRange.Font.Size = 8
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .VerticalAnchor = 3
        .MarginLeft = 6: .MarginRight = 4: .MarginTop = 0: .MarginBottom = 0
    End With
    chip.OnAction = "modAppAct.OnFollowupChipOff"
    chip.AlternativeText = "クリックすると「続きの質問」をやめます"
    On Error GoTo 0
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
        modAppDef.APP_NAME & " - どこが気になりましたか")
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
                   vbInformation, modAppDef.APP_NAME
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
    ' R14-8c: クリップボードの改行はWindows標準のCRLFで渡す(TargetTextはvbLf
    ' 統一済み)。LFのままだとメモ帳など古い貼り付け先で1行に潰れて見える。
    If modClip.SetClipboardText(Replace(t, vbLf, vbCrLf)) Then
        modSkin.ShowToast "回答をコピーしました。Ctrl+V でどこへでも貼り付けできます。", "success"
    Else
        MsgBox "コピーに失敗しました。お使いの環境では手動での選択をお試しください。", _
               vbExclamation, modAppDef.APP_NAME
    End If
Done:
    modUiLock.Leave
End Sub
