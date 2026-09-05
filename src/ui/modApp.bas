Attribute VB_Name = "modApp"
Option Explicit

' modApp - Nexus Agent Controller。UI(modUI)とエンジン(modAsk/modGateway/
' modShelf)を結合する制御層。固定アクションバーは「選択中のAIバブル」
' (未選択時は最新)に対して発火。P2P共有パスはconfig nexus_share_pathで差替可。

' A3: 入力の最大文字数(超過はカット+警告)。2026-07-31(R11-F1)の modAppAct 分離で
' Public にした。2026-08-03(R13-6a)に深掘りも入力欄+送信の1本道へ合流したため、
' 上限を掛ける場所はこの OnSend だけになっている(同じ数字を2箇所に置かない)。
Public Const MAX_INPUT_CHARS As Long = 2000

' 連打/多重発火はmodUiLockへ一本化(Enter/Leave対で必ずLeave到達)。

' R18H FB-5: 直前に「入念に調べる」を案内した質問文。同じ質問を連打しても
' 同じ案内を毎回出さないためだけの1つ前の記憶(モジュールレベル宣言は
' プロシージャより前に置く。実機VBAの制約)。
Private mLastCompoundQ As String

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

' R34 B0: ④会話の記憶(SaveTurnForRestore)の実体は modAppState へ移設した
' (modApp 残36字で B1 の配線1行が入らなかったため)。呼び出しは241行目の1箇所のみ。

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
    ' R14-G12: 復元したAI回答も、その場で生成した回答と同じ整形を通す
    ' (modLive.AnswerParagraphs=記法の保険変換+vbCr段落)。通さないと、
    ' 前回の会話だけが "## " や "**" の生記号・1段落ベタ組みで表示され、
    ' 起動直後の画面が一番壊れて見える(実機第3報 RC9と同じ見え方)。
    Dim i As Long
    For i = n To 0 Step -1
        If LenB(Trim$(us(i))) > 0 Then
            modUI.AddChatBubble "user", us(i)
            Dim bn As String
            bn = modUI.AddChatBubble("ai", ChrW(&HD83D) & ChrW(&HDCDC) & "(前回の回答) " & _
                modLive.AnswerParagraphs(aas(i)))
            modLive.StyleAnswerParas bn   ' ■見出しの太字も生成時と同じにする
        End If
    Next i
    On Error GoTo 0
End Sub

' OnSend - 送信ボタン。入力セル(nx_input)を読み、モードに応じて回答生成。
Public Sub OnSend()
    If modUiLock.BlockIfIngesting() Then Exit Sub   ' R13-4c: 取込中は質問を受けない
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Fail
    Dim stage As String: stage = "検証"   ' R25-1b: 中断時のask_abort段階名
    modPeek.HideCitations   ' 前回回答の出典チップ/ポップアップを消す(最新回答の下だけに出す)
    modMentor.ClearMentor   ' Mentorボタンも同時に掃除(内部On Error Resume Next=安全弁)
    modAppAct.ClearActions            ' 文脈アクションも消す(質問中はボタン0個=入力に集中)
    modAppAct.ClearConfidence         ' 信頼度バッジはnx_act_ではないので個別に消す
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
               vbInformation, modAppDef.APP_NAME
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
    stage = "受付"   ' F8(m-8): 仮バブル生成前の段階を記録(ask_abort粒度)
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
        modLive.PaintStage modAppState.ThoroughPreNotice(sendMode)   ' R26-1同梱2
    End If
    On Error GoTo Fail

    ' R13 F3(武装したフォローアップの漏れ止め): 「続きの質問」の武装は
    ' 【モード分岐より前に】必ず1回で消費する。R20-2c: 一般アシスタントは
    ' AskGeneral自身が会話履歴(mGenPrevU/mGenPrevA)を保持しているため、
    ' armedならそのまま続きの質問としてAskGeneralを呼べばよい(RAG専用の
    ' isFollowup分岐/modFollowupには合流させない=握りつぶしではなく退避)。
    Dim isFollowup As Boolean
    isFollowup = modAppAct.ConsumeArmedFollowup()
    Dim wasArmedGeneral As Boolean
    wasArmedGeneral = (isFollowup And sendMode = "normal")
    If sendMode = "normal" Then isFollowup = False
    ' R16-3D: 深掘りかどうかを qa層へ1回だけ預ける(下書き指示の出し分けと
    ' 既出チャンクメモリの寿命が、どちらもこの旗を見る)。
    modFollowup.SetFollowupTurn isFollowup

    stage = "検索/生成"
    Dim t0 As Double: t0 = Timer

    Dim ans As String
    Dim grounded As Boolean
    If sendMode = "normal" Then
        ' R20-2c: armed(=深掘り/続けて質問)だったときだけ実況で伝える
        ' (履歴自体はAskGeneralが常に保持済みなので分岐は表示だけでよい)。
        If wasArmedGeneral Then
            On Error Resume Next
            modLive.PaintStage "続きの質問として回答中…"
            On Error GoTo Fail
        End If
        ans = modAppState.AskGeneral(q, "", sendSpeed)
    ElseIf isFollowup Then
        ' R13-6a: 「続けて質問/深掘り」で武装済み。会話を引き継いで答える
        ' (エフォートは AskFollowup が送信時点のトグル値を読む=6b)。
        ' 本棚が空かどうかより先に判定する。ここを後ろに置くと、資料を
        ' 全部消した直後の1回だけ武装が解けずチップが残る。
        modAsk.AskFollowup q
        ans = modAsk.LastAnswerText()
        grounded = True
    ElseIf modAppState.ShelfIsEmpty() Then
        ' 資料が1件も無いのに検索へ行くと、埋め込みAPIを1往復使ったうえで
        ' 「資料がありません」とだけ返る。いちばん遅い経路が、いちばん
        ' 価値の無い返事に着く。空だと分かっているなら聞くまでもない。
        ans = modAppState.AnswerWithoutShelf(q)
    Else
        ans = modAsk.Answer(q, sendSpeed)
        grounded = True
    End If
    ' R34 B1: 両経路の合流点で出典突合を1回だけ通す(quick/deepのみ。実体は
    ' modMode。thoroughは生成の内側で突合済みなので中で弾かれる=二重付与なし)。
    If grounded Then ans = modMode.AnnotateIfNeeded(ans)

    Dim secs As Double: secs = Timer - t0
    stage = "描画/保存"

    On Error Resume Next
    modLive.Finish
    If LenB(phName) > 0 Then ThisWorkbook.Worksheets("Nexus").Shapes(phName).Delete
    On Error GoTo Fail

    ' 速さと調べた量は、言わなければ伝わらない。「15秒待たされた」と
    ' 「3冊の資料から12秒で根拠付き」は同じ時間の別の体験になる。
    Dim bubbleName As String
    ' R14-8c: 本文は modLive で記法を整え、段落区切りを vbCr にしてから書く
    ' (Shape の Paragraphs は vbCr でしか分かれない。フッターの装飾も同じ前提)。
    ' R20-6a: フッターへモード名を先頭表示(3モード同一に見える問題の可観測化)。
    bubbleName = modUI.AddChatBubble("ai", _
        modLive.AnswerParagraphs(ans) & vbCr & modLive.Footer(secs, grounded, sendMode, sendSpeed))
    modLive.StyleFooter bubbleName
    modLive.StyleAnswerParas bubbleName   ' ■見出しの段落だけ太字(AI回答のみ)
    modAppState.SetActiveBubble bubbleName
    modUI.MarkActiveBubble bubbleName
    modAppState.SaveTurnKind sendMode, isFollowup   ' R36 Fix M3: 是正メモ経路のゲート
    modAppState.SaveTurnForRestore q, ans   ' ④記憶の継続: 次回起動時の復元用

    ' 積む順は 回答 → 信頼度 → 出典 → 評価。根拠を見る前に評価させない。
    ' (一般アシスタントは出典が無いので信頼度・出典は出さない=誤表示も防ぐ)
    If sendMode <> "normal" Then
        modAppAct.DrawConfidence bubbleName
        modPeek.RenderCitations bubbleName
    End If
    modAppAct.DrawActions bubbleName
    If sendMode <> "normal" Then
        modMentor.OfferMentor bubbleName   ' Mentor: 専門家ボタン(失敗しても出ないだけ=安全弁内蔵)
    End If
    On Error Resume Next
    modUI.SettleChat        ' 次のバブルがアクション/出典に重ならないよう下端を確定
    OfferThoroughForCompound sendMode, sendSpeed, q
    On Error GoTo Fail

    ' 爆速証明(狂気案Lv.1): binary_rag_debug=TRUEのとき、直近ハイブリッド検索の所要msを
    ' Toastで見せる(qa層のperfログをUI層で取り出す=R1レイヤリングを守る)。
    If modConfig.GetBool("binary_rag_debug", False) Then
        Dim perf As String: perf = modBitwiseOpt.ConsumePerfLog()
        If LenB(perf) > 0 Then modSkin.ShowToast perf, "info"
    End If

    modUINexusDraw.RedrawInputHint   ' R29 W2-2b
    modUiLock.Leave
    Exit Sub

Fail:
    Dim failNum As Long, failDesc As String
    failNum = Err.Number: failDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup3
FailCleanup3:
    Err.Clear
    On Error Resume Next
    modLive.Finish   ' 実況先を必ず手放す(次のターンへ持ち越さない)
    ' R25-1b: 「考えています…」の仮バブルが残ったまま(=無言の砂時計)に
    ' ならないよう、中断時もここで確実に消す(従来は成功時のみ削除していた)。
    If LenB(phName) > 0 Then ThisWorkbook.Worksheets("Nexus").Shapes(phName).Delete
    modLog.LogError "E0602", "modApp.OnSend", failDesc
    modLog.LogUsage "ask_abort", stage, "Err#" & failNum & ": " & failDesc   ' R25-1b: 可観測性
    ' 2026-07-28(レビュー L-22): 失敗したら入力欄へ書き戻す。
    ' 送信直後に入力欄をクリアする作りなので、長文を書いて送って落ちると
    ' 打った文章が丸ごと消えていた。「もう一度お試しください」と言われても、
    ' もう一度打ち直すところからになる。
    If LenB(q) > 0 Then modAppState.RestoreInputCell q
    modUI.AddChatBubble "ai", "エラーが発生しました。入力欄に文章を戻しましたので、" & _
        "もう一度お試しください。(" & failDesc & ")"
    On Error GoTo 0
    modUINexusDraw.RedrawInputHint   ' R29H F1
    modUiLock.Leave
End Sub

' ----------------------------------------------------------------------------
' OfferThoroughForCompound - 複合質問を入念モードへ案内する(R18H FB-5 / B-M4)。
' ----------------------------------------------------------------------------
' 論点ごとの分解・番号での聞き返しは【入念に調べる】だけの機能で、すぐ聞く/
' しっかり調べるでは複合質問を1本のクエリのまま検索する。どの論点にも薄く
' 当たった資料が上位に来て、丁寧に薄い答えが返る(R16-3A)。利用者からは
' 「聞き方が悪かったのか、ツールが弱いのか」が区別できないので、複合の合図
' (modRagParse.HasCompoundSignal)が立った回だけ、回答を出した【あとで】
' 入念モードの存在を1回伝える(質問の前に説教しない=答えを先に出す)。
' 一般アシスタント(normal)には検索も分解も無いので出さない。
' 同じ質問を連打したときの連発は直前の質問文で抑える(ガードは1つ前だけ。
' セッション中に一度きりにすると、別の複合質問のときに案内が届かない)。
' 表示の失敗が回答を壊してはならないので呼び出し側の OERN 配下に置く(§4-4)。
Private Sub OfferThoroughForCompound(ByVal sendMode As String, _
                                     ByVal sendSpeed As String, ByVal q As String)
    If sendMode = "normal" Then Exit Sub
    If sendSpeed = "thorough" Then Exit Sub
    If Not modRagParse.HasCompoundSignal(q) Then Exit Sub
    If StrComp(q, mLastCompoundQ, vbTextCompare) = 0 Then Exit Sub
    mLastCompoundQ = q
    modSkin.ShowToast "複数の論点を含む質問は「入念に調べる」で論点ごとに" & _
        "分けて回答できます", "info"
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
    If modUiLock.BlockIfIngesting() Then Exit Sub   ' R13-4c: 取込中は出典を開かない
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
' 2026-07-30(R4要件E)の作り直し:
'   旧実装は TotalChunks の差分だけで成否を判定していた。ところが「既に
'   入っている資料をもう一度入れた」ときは取込自体は成功(status=done)なのに
'   チャンクは1件も増えない。その結果、成功しているのに
'   「資料は追加されませんでした」と言い、同時にmodShelf側のMsgBoxが
'   「1件を本棚に追加しました」と言う ―― 正反対の報告が2つ同時に出ていた。
'   R2で追加された modShelfBatch.AddFilesResult(showMsgBox:=False) を使い、
'   報告をチャットバブル1本に統一して、結果ごとに正直な文言を出す。
Public Sub OnAddDocs()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done

    Dim res As String
    On Error Resume Next
    res = modShelfBatch.AddFilesResult(showMsgBox:=False)
    modUI.GoToNexus "modApp.OnAddDocs"   ' 取込は別シートを触るので必ず戻す
    On Error GoTo Done

    Dim say As String
    If LenB(res) = 0 Then
        ' 4-B: AddFilesResult が例外で抜けた(戻り値が入らなかった)場合。
        ' キャンセルは "ok=0;ng=0;..." が必ず入るので区別できる。
        ' ここで黙ると、押したのに何も起きない=最悪の無反応になる。
        say = ChrW(&H26A0) & " 取込処理でエラーが発生しました。" & vbLf & _
              "err_log シートをご確認ください(管理者にこの画面を見せてください)。"
    Else
        say = AddDocsMessage(res)
    End If
    ' ファイル選択をキャンセルしただけのときは何も言わない
    ' (押し間違いに見えるだけで、利用者に伝えることが無い)。
    If LenB(say) > 0 Then modUI.AddChatBubble "ai", say

    On Error Resume Next
    modUI.SettleChat
    On Error GoTo Done
Done:
    modUiLock.Leave
End Sub

' 取込結果("ok=..;ng=..;capped=..;chunks=..;reasons=<code>:<n>,..")を
' 利用者向けの1本の文面へ翻訳する。空文字=何も言わない(キャンセル)。
Private Function AddDocsMessage(ByVal res As String) As String
    Dim okN As Long, ngN As Long, capN As Long, chunkN As Long
    okN = ResultNum(res, "ok")
    ngN = ResultNum(res, "ng")
    capN = ResultNum(res, "capped")
    chunkN = ResultNum(res, "chunks")

    Dim reasons As String
    reasons = ResultText(res, "reasons")

    ' 2026-07-31(R11-H Med1): 「取込中で受け付けなかった」をキャンセルと
    ' 区別する。どちらも件数は全部0なので、reasons=busy でしか見分けられない。
    If ReasonCount(reasons, "busy") > 0 Then
        AddDocsMessage = ChrW(&H23F3) & " いま別の取り込みが動いています。" & vbLf & _
            "完了後にもう一度お試しください(進み具合は画面上部の帯に出ています)。"
        Exit Function
    End If

    ' R15-FixB(FB-5): 見送り(確認で「いいえ」)は ng に入らなくなったので、
    ' ここで一緒に見ないと【1件だけ選んで見送った】ときが「キャンセル」と同じ
    ' 無言になる。自分で見送ったことすら画面から消えるのは黙り過ぎ(§4-1)。
    Dim decN As Long: decN = ReasonCount(reasons, "declined")
    If okN = 0 And ngN = 0 And capN = 0 And decN = 0 Then Exit Function   ' キャンセル

    Dim say As String
    If okN > 0 And chunkN > 0 Then
        say = ChrW(&H2705) & " 資料を取り込みました。" & vbLf & _
            "これで、この資料の中身について「どのページに書いてあるか」まで付けてお答えできます。" & vbLf & _
            "さっそく、いま知りたいことをそのまま聞いてみてください。"
    ElseIf okN > 0 Then
        ' 取込は成功したのにチャンクが増えていない。4-C: 「すでに入っています」と
        ' 言い切ると【更新版を入れ直した人に嘘をつく】ことになる。更新取込では
        ' 変わった部分だけが入れ替わり、同じ部分は重複排除で増えないので、
        ' 総数が増えないことは珍しくない(反映はされている)。
        say = ChrW(&H2705) & " 取り込みました。前回から内容が変わっていない部分は増やしていません" & _
            "(更新は反映されています)。" & vbLf & _
            "そのまま、この資料について聞いていただけます。"
    End If

    Dim dupN As Long, imgN As Long
    dupN = ReasonCount(reasons, "E0504")
    imgN = ReasonCount(reasons, "image_pdf")

    If dupN > 0 Then
        say = AppendBlock(say, ChrW(&H26A0) & " 同じ名前の資料が、別の場所からすでに登録されています(" & dupN & "件)。" & vbLf & _
            "ファイル名を変えて入れ直すか、マイ本棚で古いほうを削除してから、もう一度お試しください。")
    End If
    If imgN > 0 Then
        say = AppendBlock(say, ChrW(&HD83D) & ChrW(&HDDBC) & " 画像として保存されたPDFで、文字を取り出せませんでした(" & imgN & "件)。" & vbLf & _
            "Ghostscriptを置くとAIが1ページずつ読み取れます(手順は「43_画像PDFのOCR取込設定」)。" & vbLf & _
            "すぐ試すなら、その画面をコピー(Win+Shift+S)して、ナレッジ画面の「" & _
            ChrW(&HD83D) & ChrW(&HDCF8) & " スクショ取込」からどうぞ。")
    End If

    ' R15-FixB(FB-5): 見送りは「読み取れませんでした」の仲間ではない。上の
    ' image_pdf の文(Ghostscriptを置けば読める)を見送った資料にまで出すと、
    ' 何も壊れていないのに設定作業を促すことになる。件数は declined として
    ' 別に届くので、専用の1文で言う。
    If decN > 0 Then
        say = AppendBlock(say, ChrW(&H23F3) & " 時間がかかるため取り込みませんでした(" & decN & "件)。" & vbLf & _
            "確認で「いいえ」を選んだ資料です。もう一度選んで「はい」を選ぶと取り込めます。")
    End If

    Dim otherN As Long
    otherN = ngN - dupN - imgN
    If otherN > 0 Then
        say = AppendBlock(say, otherN & "件、取り込めませんでした。" & vbLf & _
            "マイ本棚の一覧で、その資料の状態欄をご確認ください。")
    End If
    If capN > 0 Then
        say = AppendBlock(say, capN & "件は本棚の上限に達したため見送りました。" & vbLf & _
            "configシートの shelf_max_chunks を大きくすると上限を増やせます。")
    End If

    If LenB(say) = 0 Then
        say = "資料は追加されませんでした。" & vbLf & _
            "対応しているのは PDF / Word / Excel / テキスト です。" & vbLf & _
            "資料が無くても、一般的な内容ならこのままお答えできます。"
    End If
    AddDocsMessage = say
End Function

' "ok=1;ng=0;..." から key の値(文字列)を取り出す。
Private Function ResultText(ByVal res As String, ByVal keyName As String) As String
    Dim parts() As String
    parts = Split(res, ";")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If Left$(parts(i), Len(keyName) + 1) = keyName & "=" Then
            ResultText = Mid$(parts(i), Len(keyName) + 2)
            Exit Function
        End If
    Next i
End Function

Private Function ResultNum(ByVal res As String, ByVal keyName As String) As Long
    Dim s As String
    s = ResultText(res, keyName)
    If IsNumeric(s) Then ResultNum = CLng(s)
End Function

' reasons="E0504:2,image_pdf:1" から特定コードの件数を取り出す。
Private Function ReasonCount(ByVal reasons As String, ByVal codeName As String) As Long
    If LenB(reasons) = 0 Then Exit Function
    Dim parts() As String
    parts = Split(reasons, ",")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim kv() As String
        kv = Split(parts(i), ":")
        If (UBound(kv) - LBound(kv)) >= 1 Then
            If kv(0) = codeName Then
                If IsNumeric(kv(1)) Then ReasonCount = CLng(kv(1))
                Exit Function
            End If
        End If
    Next i
End Function

' 文面のブロック連結(1つ目のブロックの前に空行を作らない)。
Private Function AppendBlock(ByVal say As String, ByVal extra As String) As String
    If LenB(say) = 0 Then
        AppendBlock = extra
    Else
        AppendBlock = say & vbLf & vbLf & extra
    End If
End Function

' ナビゲーション(SPA遷移)・モード/言語トグル
Public Sub OnNavChat()
    If modUiLock.BlockIfIngesting() Then Exit Sub   ' R7 B-2
    If Not modUiLock.Enter() Then Exit Sub
    ' R27波3-9: ロックを取りながらハンドラが無かった。ここで例外が出ると
    ' Leaveに到達せず、全ボタンが自動解除(10分)まで無音で死ぬ(OnNavShelfも同型)。
    On Error GoTo Done
    modUI.GoToNexus "modApp.OnNavChat"
Done:
    modUiLock.Leave
End Sub

Public Sub OnNavHome()
    If modUiLock.BlockIfIngesting() Then Exit Sub   ' R7 B-2
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modHub.EnsureHubLayout activate:=True   ' 描画と遷移を必ずセットで行う
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnNavShelf()
    If modUiLock.BlockIfIngesting() Then Exit Sub   ' R7 B-2
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done   ' R27波3-9(OnNavChatと同型)
    modUI.GoToNativeSheet modAppDef.SH_SHELF, "modApp.OnNavShelf"
Done:
    modUiLock.Leave
End Sub


' 🔄 画面を再描画(盲点B2/C5): ウィンドウのリサイズ・Alt+Tab復帰・マルチモニタ間の
' 移動でShapeがゴースト化/ズレたとき、ユーザーが1クリックで現在の画面を作り直す。
' 自己インストーラ配布版ではWorkbook_WindowActivate等が発火しない制約があるため、
' 自動ではなく明示的なリフレッシュ手段を提供する。アクティブな画面に応じて振り分け:
'   Nexus     → 会話履歴を壊さず視覚不変条件だけ再適用(modUI.Repaint)
'   Dashboard → データから再構築(modDash)
'   その他    → マイ本棚シートを「今のモードのまま」データから再構築
'               (2026-07-30 R4要件F: ギャラリー/一覧/解決事例は同じシートに
'                描くようになったので、決め打ちでギャラリーへ戻すと
'                一覧を見ていた人の画面が勝手に変わってしまう)
Public Sub OnRefreshUI()
    ' R15-2a(実機第4報 RC8): 再描画は画面を丸ごと組み直す重い処理で、
    ' 取込が掴んでいるシート状態と噛み合うと入れ子で崩れる。他のナビ系
    ' ハンドラと同じく BlockIfIngesting を Enter の前に置く(後ろに置くと
    ' ロックを取ったまま Exit してしまい、全ボタンが10分死ぬ)。
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUI.EnsureAppView          ' R7 A-2: 表示状態の自己修復を再描画にも載せる
    Select Case ActiveSheet.Name
        Case "Nexus":     modUI.Repaint
        Case "Dashboard": modDash.ShowDashboard
        Case modAppDef.SH_HOME: modHub.EnsureHubLayout
        Case Else:        modKnowledge.RefreshCurrent
    End Select
    If Err.Number <> 0 Then modLog.LogError "E0801", "modApp.OnRefreshUI", Err.Description, Err.Number
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

    Dim oldMode As String: oldMode = modAppState.CurrentMode()
    Dim newMode As String
    If oldMode = "normal" Then
        newMode = "rag"
    Else
        newMode = "normal"
    End If
    modAppState.WriteUiState modAppState.MODE_KEY, newMode
    modAppState.UpdateModeButton
    On Error Resume Next
    modAppState.BridgeConvMemory oldMode, newMode   ' R26-2: 会話メモリの橋渡し
    modClarify.ClearPending   ' R28H F4: モードを跨いだら逆質問は破棄
    modUINexusDraw.RedrawInputHint   ' R29H F1
    On Error GoTo 0
    ' R13 F3: モードを跨いだら「続きの質問」の武装は捨てる。一般アシスタントには
    ' 会話引き継ぎの意味論が無く、RAGへ戻したときに前の話題が生き返るのも
    ' 利用者の意図ではない。切替が成功したこの位置(busyガードの後)でだけ解除する。
    On Error Resume Next
    modAppAct.OnFollowupChipOff
    On Error GoTo 0
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
    ' 3-A(1): ピルへ直接書かず、幅計算を通してヘッダーごと描き直す。
    modUINexusDraw.RedrawChatHeader
    modUINexusDraw.ReapplyChatBound   ' R27 F2-1c: 行1の高さが変わると境界がずれる
    ' R13 F3: 力の入れ方(effort)の切替では武装は保つ ―― 「チップを出したまま
    ' モードを選んで送る」は R13-6b の設計そのもの。ただしチップに書いてある
    ' モード名は【押した時点】の文字列なので、切り替えたら描き直さないと
    ' 表示だけが古い名前のまま残る(実際に使われるのは送信時点の値)。
    modAppAct.RedrawFollowupChip
    ' R26-1: 一般アシスタントも3段化したのでR21-2 D4の「使われない」注記は撤去。
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
    ' R35(方式B)以降、ビルドがモジュール本文を CP932 で
    ' vbaProject.bin へ焼き込むため、CP932 に無い ế/ệ はリテラル "?" として
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
    ' R27H F6(項目13裁定): 成否を見てから知らせる。SetValue は config シートを
    ' 見つけられないと黙って何もしない(戻り値も無い)ので、書いた値を読み直して
    ' 一致を確かめる(modConfig はキャッシュを持たない=これが実測になる)。
    ' modHub 側の無条件「成功」トーストは撤去し、出す場所をここ1箇所にした。
    Dim langOK As Boolean
    langOK = (StrComp(modConfig.GetString("answer_language", ""), nextLang, vbBinaryCompare) = 0)

    On Error Resume Next
    ' 3-A(1): ピルへ直接書かず、幅計算を通してヘッダーごと描き直す。
    modUINexusDraw.RedrawChatHeader
    modUINexusDraw.ReapplyChatBound   ' R27 F2-1c: 行1の高さが変わると境界がずれる
    If langOK Then
        modSkin.ShowToast "回答言語: " & nextLang, "success"
    Else
        modSkin.ShowToast "回答言語を切り替えられませんでした(設定を保存できません)。", "error"
    End If
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
    modUINexusDraw.RedrawInputHint   ' R29H F1
    ' 2026-07-28(レビュー L-18): クリアで消えていなかったものを片付ける。
    '   ・出典チップ / 専門家ボタン … 消した会話の下にボタンだけ残っていた
    '   ・続けて質問の履歴 / 復元用の直近ターン … 残っていると、再起動時に
    '     「クリアしたはずの会話」が復元される(利用者から見れば消えていない)
    ' 「消した」と言った以上は、次に開いたときも消えていなければならない。
    modPeek.HideCitations
    modMentor.ClearMentor
    modAppAct.ClearActions
    modAppAct.ClearConfidence
    modAppAct.OnFollowupChipOff   ' R13-6a: 武装したままの「続きの質問」も解除する
    modFollowup.ClearCitedSources ' R13-5b: 会話の出典メモリも消す(会話リセット)
    modAppState.ClearGeneralMemory ' R20-2d: 一般アシスタントの会話履歴も消す(新発見バグ)
    ' R27波3-6: 消したQ&Aが部内へ共有される穴を塞ぐ。🟢🔴🟡(btn_fb_*)は
    ' 常設Shapeで OnAction が modAsk.FeedbackGreen 等へ直結しており、クリア後に
    ' 押すと消したはずのQ&AがEmitVerifiedQAで共有フォルダへ出ていた。modAskは
    ' 凍結だが同用途(残留状態で感想を撃たせない)のPublicが既にある。
    ' NoteAnswerFailed が mLastQuestion を空にし mFeedbackDone を立てるので、
    ' 以後 FeedbackAccepted が3ボタンとも断る。👎/🤔の訂正共有は対象バブルが
    ' 消えて HasTarget=False になるため従来どおり撃てない。
    modAsk.NoteAnswerFailed
    modUIMain.ClearLastAnswerState  ' 「Wordで開く」が消した回答を出さないように
    modAsk.ResetPrevMemory   ' R28-W4: state だけでは mPrevU 亡霊が残る(M-4)
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
    ' R11-A C1 / R13-4d: 取込中は「無反応」でも「黙って終了」でもなく、
    ' 何が起きているかを伝えて選ばせる(いいえ=終了中止)。
    If Not modUiLock.ConfirmCloseDuringIngest() Then Exit Sub
    If Not modUiLock.Enter() Then
        ' 承諾はしたが、別の処理中で終了に進めなかった。この承諾は使われないまま
        ' 60秒生き残り、その間の×クリックが無確認で閉じる(R13 F5)。捨てる。
        modUiLock.CancelCloseOk
        Exit Sub
    End If
    Dim resp As VbMsgBoxResult
    resp = MsgBox("保存してこのファイルを閉じますか?" & vbCrLf & vbCrLf & _
        "  [はい] 保存して閉じます" & vbCrLf & _
        "  [いいえ] 保存せずに閉じます(今回取り込んだ資料は消えます)" & vbCrLf & _
        "  [キャンセル] 閉じずに元の画面へ戻ります", _
        vbYesNoCancel + vbQuestion, modAppDef.APP_NAME)
    If resp = vbCancel Then
        ' 終了は取りやめ。取込中の終了承諾も一緒に取り消す(R13 F5)。
        ' ここを消し忘れると「やめる」と答えた直後の60秒だけ、×クリックが
        ' 確認なしで閉じる窓が開く。
        modUiLock.CancelCloseOk
        modUiLock.Leave
        Exit Sub
    End If
    modUiLock.Leave

    ' 2026-08-01(R12-1-2): 終了時の後始末を明示的に実行してから閉じる。
    ' 本番のThisWorkbookはWindowResize転送のみ(R35)で
    ' Workbook_BeforeClose を持たず、かつ Excel は「VBAから呼んだ Close」では
    ' Auto_Close マクロを実行しない。つまりこの終了ボタン経由で閉じると
    ' 後始末が一切走らず、
    '   ・Application.OnKey "^z", "" が残る(閉じた後もExcel全体でCtrl+Zが死ぬ)
    '   ・^+q / ^~ / ^{ENTER} が閉じたブックのマクロに紐付いたまま残る
    '   ・リボン非表示・全画面・数式バー非表示がExcel全体に残る
    '   ・modTelemetry.Publish(送信は終了時の設計)が実行されない
    '   ・sync_interval_min>0 なら AutoSyncTick の OnTime 予約が残り、
    '     閉じたブックをExcelが勝手に開き直す
    ' という状態でExcelに戻ることになる。Auto_Close は解除系と上書き系だけで
    ' 構成されていて冪等なので、開発構成(ThisWorkbook.cls あり)で
    ' Workbook_BeforeClose と二重に走っても害はない。
    On Error Resume Next
    modTextView.CleanupSheet "modApp.OnSaveAndExit"   ' R36 Fix A-r4: text_viewを保存物に残さない
    modBoot.Auto_Close
    On Error GoTo 0

    ThisWorkbook.Close SaveChanges:=(resp = vbYes)
End Sub

' 内部ヘルパー

