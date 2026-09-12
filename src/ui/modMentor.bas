Attribute VB_Name = "modMentor"
Option Explicit

' ============================================================================
' modMentor - 専門家召喚プラグイン(Mentor)。完全疎結合のオプトイン機能。
' ----------------------------------------------------------------------------
' 役割:
'   RAG回答の根拠チャンクの出所タグ(origin="pack:<作者>")から「この分野に最も
'   詳しい人(専門家)」を特定し、回答・出典チップの下に
'   「💡 この分野は〇〇さんが詳しいです [質問を送る]」ボタンを描画する。
'   クリックで質問を入力させ、既存P2Pと同じ共有フォルダ機構(ファイルベースI/O)で
'   専門家宛てにサイレント送信し、Toastで完了を通知する。AIが答えるだけでなく、
'   社内の「人」へ繋ぐハブになる組織エコシステムの最終ピース。
'
' 防衛設計(4条項の実装):
'   1. モジュール隔離: ロジックは本モジュールに完結。既存コードへの変更は
'      modApp内のフック4行(ClearMentor/OfferMentor各2箇所)のみ。
'   2. サーキットブレーカー: エントリポイント(OfferMentor/ClearMentor)の最上部で
'      On Error Resume Next。本機能のどんな失敗も「ボタンが出ないだけ」に留め、
'      回答バブル表示という中核機能へ絶対に波及させない。
'   3. グローバル汚染ゼロ: 既存のUI状態・RAGスコア変数には一切書き込まない。
'      読むのは modAsk.LastHit*(読み取り専用アクセサ)と modP2P.CurrentUserId、
'      modShareRule.OriginKind/OriginName、modP2P.ResolveAuthorId(いずれも
'      読み取り専用)、modConfig のみ。自前状態は本モジュールPrivate
'      (mExpert/mExpertId/mExpertKind/mTopSource)に完結。
'   4. サニタイズ徹底: 専門家名・質問者名は modP2PIo.SanitizeId で禁止文字を
'      除去し、さらにファイル名には modP2PIo.IdHash(16桁)のみを使う
'      (MAX_PATH対策も同時達成。payload側に生のsanitize済IDを保持)。
'      R33H F30: 送信側だけが均していた複製をやめ、送受信とも IdHash を通す。
'   ・送信I/OはmodMentor内に自前カプセル化(ADODB.Stream+リトライ+Set=Nothing)。
'     R33 W5-11 で1点だけ例外を作った: 宛先IDの解決だけは modP2P.ResolveAuthorId
'     (このとき Private → Public 化)を通す。表示名を宛先にすると受信側の照合
'     キー(ユーザID)と一致せず誰にも届かないため、宛先解決の情報源は感謝状と
'     1本に揃える必要がある。I/O は従来どおり自前カプセル化のまま。
'   ・共有フォルダ書式: <nexus_share_path>\questions\q_<宛先hash>_<nonce>.txt
'     内容はTSV(nonce/from/to/質問/出典/時刻)。将来の受信機能(CollectQuestions)が
'     感謝状と同じcollect-then-process方式で読める形式にしてある。
' ============================================================================

Private Const QUESTIONS_SUBDIR As String = "questions"
Private Const BTN_NAME As String = "nx_mentor_btn"
Private Const MAX_Q_CHARS As Long = 1000       ' 質問文の上限(Shape/ファイル肥大防止)

' 本モジュール内で完結する状態(防衛条項3: 既存グローバルには一切触れない)
Private mExpert As String      ' 直近OfferMentorで特定した専門家の【表示名】
Private mExpertId As String    ' R33 W5-11: その解決済み【宛先ID】(送信に使うのは必ずこちら)
Private mExpertKind As String  ' R33 W5-10: 出所の種別("pack" / "channel")
Private mTopSource As String   ' その専門家の代表ソース名(質問の文脈として同送)
Private mLastAsker As String   ' 直近に受信した質問の差出人(返信ボタンの宛先)

' ----------------------------------------------------------------------------
' OfferMentor - エントリポイント(modApp.OnSend/modAppAct.OnActDrill末尾から1行フック)。
'   直近RAG回答の出典から専門家を特定できたときだけボタンを描く。
'   失敗しても静かに何もしない(サーキットブレーカー)。
' ----------------------------------------------------------------------------
Public Sub OfferMentor(ByVal bubbleName As String)
    On Error Resume Next   ' 安全弁: 本機能の失敗を絶対にメインへ波及させない
    ClearMentor

    ' R33H F8: 回答が成立しなかったターン(API失敗の #ERR・逆質問)には出さない。
    ' 信頼度バッジ(modAppAct.DrawConfidence)と出典チップ(modPeek.RenderCitations)
    ' は W5-21/F7 の門をくぐったのに、ここだけ LastHitCount しか見ていなかった。
    ' そのため「AIとの通信に失敗しました」の直下に
    ' 「この分野は○○さんが詳しいです[質問を送る]」だけが残り、答えが出ていない
    ' のに人を呼ぶ導線だけが立つ。門は modApp 側ではなくここに置く ――
    ' modApp は残36字で1行も足せず、ここに置けば全ての呼び出し元に効く。
    ' ClearMentor の【後】に置くのは、前のターンのボタンを必ず消すため。
    If Not modMode.GroundingAllowed() Then Exit Sub

    Dim expert As String, topSource As String
    Dim kind As String, targetId As String
    If Not FindExpert(expert, topSource, kind, targetId) Then GoTo Done
    ' R33 W5-11: 宛先IDが解決できないヒットは FindExpert が候補にしないが、
    ' 冒頭の On Error Resume Next で途中の例外が握り潰される可能性があるため、
    ' ここでも最終防衛線を張る。届けられない相手に「送りました」と言わない。
    targetId = Trim$(targetId)
    If LenB(targetId) = 0 Then GoTo Done
    ' 2026-07-28(レビュー C-2): 冒頭の On Error Resume Next のせいで、
    ' FindExpert 内で例外が起きても「見つかった」扱いのまま空文字で先へ進み、
    ' 「この分野は さんが詳しいです」という宛先の無いボタンが出ていた。
    ' 押しても何も起きないので、利用者から見ると壊れたボタンでしかない。
    ' 名前が取れていないなら出さない、を最終防衛線として置く。
    expert = Trim$(expert)
    If LenB(expert) = 0 Then GoTo Done

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then GoTo Done

    ' 描画位置: 文脈アクション(nx_act_*)と出典チップ(nx_cite_*)の最下端の下。
    ' どちらも無ければバブルの直下。
    Dim anchor As Shape
    Set anchor = ws.Shapes(bubbleName)
    If anchor Is Nothing Then GoTo Done
    Dim y As Double: y = anchor.Top + anchor.Height + 6
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 8) = "nx_cite_" Or Left$(shp.Name, 7) = "nx_act_" Then
            If shp.Top + shp.Height + 6 > y Then y = shp.Top + shp.Height + 6
        End If
    Next shp

    Dim btn As Shape
    Set btn = SafeRoundedRect(ws, anchor.Left, y, 360, 26)
    btn.Name = BTN_NAME
    btn.Adjustments(1) = 0.4
    btn.Line.Visible = -1
    btn.Line.Weight = 1#
    btn.Line.ForeColor.RGB = modUI.UiColor("accent")           ' MS&ADグリーン
    btn.Fill.ForeColor.RGB = modUI.UiColor("surface")
    ' 称号(偽装不可): 感謝受領数ベースの絶対評価(💡5+/🌟20+)を名前の頭に自動付与。
    ' (本Subは冒頭のOn Error Resume Nextが活性のためTitleFor失敗時はhonor=""のまま)
    ' R33 W5-11: 称号は【解決済みID】で引く(mTitles も CurrentUserId 系のIDが
    ' キーなので、表示名で引いていた従来は常に空だった)。
    Dim honor As String
    Dim who As String
    If StrComp(kind, "channel", vbTextCompare) = 0 Then
        ' 部門名は人ではないので「〇〇さん」とは呼ばない(R33 W5-10)。
        who = "「" & expert & "」の発行者が詳しいです"
    Else
        honor = modBoard.TitleFor(targetId)
        who = honor & expert & " さんが詳しいです"
    End If
    With btn.TextFrame2
        .WordWrap = -1
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCA1) & " この分野は " & who & " [質問を送る]"
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = -1
        .TextRange.ParagraphFormat.Alignment = 2   ' 中央
        .VerticalAnchor = 3
        .MarginLeft = 8: .MarginRight = 8: .MarginTop = 0: .MarginBottom = 0
    End With
    btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("accent")
    btn.OnAction = "modMentor.OnAskExpert"
    btn.Placement = 3   ' xlFreeFloating(絶対配置)
    modSkin.ApplySoftShadow btn

    mExpert = expert
    mExpertId = targetId
    mExpertKind = kind
    mTopSource = topSource
Done:
    ' 2026-07-31 R11-D(監査2 指摘3): 冒頭の On Error Resume Next で本機能の
    ' 失敗をメインへ波及させない設計(サーキットブレーカー)は維持したまま、
    ' 【痕跡だけ】を残す。利用者への通知は増やさない(静かに諦めるのが正しい)。
    LogMentorFail "modMentor.OfferMentor"
End Sub

' ClearMentor - ボタン削除(送信の先頭フック/OfferMentor冒頭から。孤児Shape防止)。
Public Sub ClearMentor()
    On Error Resume Next   ' 安全弁
    ThisWorkbook.Worksheets("Nexus").Shapes(BTN_NAME).Delete
    ThisWorkbook.Worksheets("Nexus").Shapes("nx_mentor_reply").Delete
    ' R33 W5-11: ボタンを消したら宛先も捨てる(前のターンの宛先が残ったまま
    ' 次のターンのボタンに使い回されることが無いようにする)。
    mExpert = ""
    mExpertId = ""
    mExpertKind = ""
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' OnAskExpert - ボタンクリック。質問を入力させ、共有フォルダの questions\ へ
'   専門家宛てにサイレント送信。完了はToast(MsgBox禁止)。
' ----------------------------------------------------------------------------
Public Sub OnAskExpert()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done

    If LenB(mExpert) = 0 Then GoTo Done
    ' R33 W5-11: 宛先IDが無いまま送ると、受信側は q_<Fnv1a64Hex(自分のID)>_*.txt
    ' しか列挙しないので誰にも届かない。届かないと分かっているなら送らず、
    ' その旨を伝える(「送りました」とだけ言って沈黙するのが最悪の壊れ方)。
    If LenB(mExpertId) = 0 Then
        modSkin.ShowToast "この資料の発行者を特定できないため、質問を送れません。" & _
                          "資料を配った方へ直接ご連絡ください。", "error"
        GoTo Done
    End If

    Dim toWhom As String: toWhom = ExpertLabel()
    Dim q As String
    q = InputBox(toWhom & "への質問を入力してください。" & vbCrLf & _
                 "(共有フォルダ経由で届きます。関連資料: " & modUtil.SafeLeft(mTopSource, 40) & ")", _
                 modAppDef.APP_NAME & " - 専門家へ質問")
    q = Trim$(q)
    If LenB(q) = 0 Then GoTo Done
    If Len(q) > MAX_Q_CHARS Then q = Left$(q, MAX_Q_CHARS)

    If SendQuestion(mExpertId, q, mTopSource) Then
        modSkin.ShowToast toWhom & "へ質問を送りました。回答をお待ちください。", "success"
        On Error Resume Next
        modLog.LogUsage "mentor_ask", "", "to=" & mExpertId & " q=" & modUtil.SafeLeft(q, 120)
        On Error GoTo Done
    Else
        modSkin.ShowToast SendFailText(), "error"   ' R27波3-10: 原因に応じた文言
    End If
Done:
    LogMentorFail "modMentor.OnAskExpert"
    modUiLock.Leave
End Sub

' ----------------------------------------------------------------------------
' CollectQuestions - 受信側(往復ループの完結)。自分宛の質問(q_<自分hash>_*.txt)を
'   回収し、チャットバブル+Toastで受け取る。感謝状と同型のcollect-then-process
'   +nonce重複排除(modStats "mq:"キー)。エントリポイントはmodApp.LaunchNexus末尾
'   の1行フック(UI初期化後=バブル/Toastが確実に描ける唯一のタイミング)。
'   本Sub自身がOn Error GoTo Doneで全障害を握るため、呼び出し元へは波及しない。
' ----------------------------------------------------------------------------
Public Sub CollectQuestions(Optional ByVal silent As Boolean = False)
    On Error GoTo Done   ' 安全弁: 受信の失敗でメイン(起動)を絶対に止めない

    Dim folderPath As String: folderPath = QuestionsDir()
    If LenB(folderPath) = 0 Then Exit Sub
    If LenB(Dir(folderPath, vbDirectory)) = 0 Then Exit Sub   ' 共有未達→静かに撤退

    Dim myId As String
    On Error Resume Next
    myId = modP2P.CurrentUserId()
    On Error GoTo Done
    If LenB(myId) = 0 Then Exit Sub

    ' 1) 宛先=自分のファイル名を全部集める(Dir列挙中にKillしない=列挙破壊防止)
    Dim names() As String: ReDim names(0 To 31)
    Dim nFiles As Long: nFiles = 0
    ' R33H F30: 送信側と同じ modP2PIo.IdHash を通す。従来は受信側だけが生の
    ' CurrentUserId() をハッシュしており、CN のエスケープ(`\`)や64字超で
    ' 送信側(SanitizeId 済み)と別のファイル名になっていた。
    Dim fn As String: fn = Dir(folderPath & "q_" & modP2PIo.IdHash(myId) & "_*.txt")
    Do While LenB(fn) > 0
        If nFiles > UBound(names) Then ReDim Preserve names(0 To UBound(names) + 32)
        names(nFiles) = fn
        nFiles = nFiles + 1
        fn = Dir()
    Loop

    ' 2) 読取り→nonce重複排除→バブル表示(最大3件)→表示できた分だけGC
    ' 2026-08-10(R27波3-11): 従来は【表示しなかった4件目以降も】nonceを
    ' 立ててファイルを消していた。同僚が送った質問が、こちらの画面に一度も
    ' 出ないまま共有フォルダから永久に消える(送った側は届いたと思っている)。
    ' 4件目以降は何も記録せずファイルも残し、次回の起動で改めて表示する。
    ' 件数(newCount)は残した分も数える=「N件届いています」は嘘にならない。
    Dim newCount As Long: newCount = 0
    Dim shown As Long: shown = 0
    Dim i As Long
    For i = 0 To nFiles - 1
        Dim full As String: full = folderPath & names(i)
        Dim rec As String
        If ReadUtf8WithRetry(full, rec) Then
            Dim f() As String: f = Split(rec, vbTab)
            If UBound(f) >= 5 Then
                If StrComp(f(2), myId, vbTextCompare) = 0 Then
                    If modStats.GetStat("mq:" & f(0)) = 0 Then
                        newCount = newCount + 1
                        If shown < 3 Then
                            shown = shown + 1
                            modStats.Bump "mq:" & f(0)
                            mLastAsker = f(1)   ' 返信ボタンの宛先(最後に表示した質問の差出人)
                            On Error Resume Next
                            modUI.AddChatBubble "ai", _
                                ChrW(&HD83D) & ChrW(&HDCEE) & " " & f(1) & " さんからあなた宛の質問が届いています。" & vbLf & _
                                "「" & modUtil.SafeLeft(f(3), 400) & "」" & vbLf & _
                                "(関連資料: " & modUtil.SafeLeft(f(4), 60) & " / " & f(5) & ")"
                            modLog.LogUsage "mentor_recv", "", "from=" & f(1) & " q=" & modUtil.SafeLeft(f(3), 120)
                            On Error GoTo Done
                            KillWithRetry full   ' 表示できた分だけGC
                        End If
                    Else
                        KillWithRetry full   ' 既知(=過去に表示済み)はGC
                    End If
                End If
            End If
        End If
    Next i

    If newCount > 0 And Not silent Then
        modUI.AddChatBubble "ai", "あなた宛の質問が " & newCount & " 件届いています。チャット欄をご確認ください。"   ' R29H F2b
        DrawReplyButton   ' 往復→対話へ: 最後の質問の差出人へ返信するボタン
    End If
Done:
    LogMentorFail "modMentor.CollectQuestions"
End Sub

' 返信ボタン(nx_mentor_reply)を最下端バブルの下に描く。質問と同機構で逆向きに送る。
Private Sub DrawReplyButton()
    If LenB(mLastAsker) = 0 Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Shapes("nx_mentor_reply").Delete
    On Error GoTo 0

    Dim anchorName As String: anchorName = modUI.LatestAiBubbleName()
    If LenB(anchorName) = 0 Then Exit Sub
    Dim anchor As Shape
    On Error Resume Next
    Set anchor = ws.Shapes(anchorName)
    On Error GoTo 0
    If anchor Is Nothing Then Exit Sub

    Dim btn As Shape
    Set btn = SafeRoundedRect(ws, anchor.Left, anchor.Top + anchor.Height + 6, 300, 26)
    btn.Name = "nx_mentor_reply"
    btn.Adjustments(1) = 0.4
    btn.Line.Visible = -1
    btn.Line.Weight = 1#
    btn.Line.ForeColor.RGB = modUI.UiColor("accent")
    btn.Fill.ForeColor.RGB = modUI.UiColor("surface")
    With btn.TextFrame2
        .WordWrap = -1
        .TextRange.Text = modEmj.Mail() & " " & mLastAsker & " さんへ返信する"
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = -1
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("accent")
    btn.OnAction = "modMentor.OnReplyQuestion"
    btn.Placement = 3
    modSkin.ApplySoftShadow btn
End Sub

' 返信ボタンのクリック。質問送信と同じカプセル化I/Oで逆向きに送る。
Public Sub OnReplyQuestion()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    If LenB(mLastAsker) = 0 Then GoTo Done

    Dim q As String
    q = InputBox("「" & mLastAsker & "」さんへの返信を入力してください。", _
                 modAppDef.APP_NAME & " - 返信")
    q = Trim$(q)
    If LenB(q) = 0 Then GoTo Done
    If Len(q) > MAX_Q_CHARS Then q = Left$(q, MAX_Q_CHARS)

    If SendQuestion(mLastAsker, q, "(返信)") Then
        modSkin.ShowToast mLastAsker & " さんへ返信を送りました。", "success"
        On Error Resume Next
        ThisWorkbook.Worksheets("Nexus").Shapes("nx_mentor_reply").Delete
        modLog.LogUsage "mentor_reply", "", "to=" & mLastAsker
        On Error GoTo Done
    Else
        modSkin.ShowToast SendFailText(), "error"   ' R27波3-10: 原因に応じた文言
    End If
Done:
    LogMentorFail "modMentor.OnReplyQuestion"
    modUiLock.Leave
End Sub

' 読取りリトライ(AVロック耐性。COMは両経路Set=Nothing)。
Private Function ReadUtf8WithRetry(ByVal filePath As String, ByRef outText As String) As Boolean
    Dim attempt As Long
    For attempt = 1 To 3
        If TryReadOnce(filePath, outText) Then
            ReadUtf8WithRetry = True
            Exit Function
        End If
        MentorWait 250 * attempt
    Next attempt
End Function

' UTF-8読み取りの実体は modUtilText.ReadTextFileUtf8(2026-07-31 R11-F2)。
Private Function TryReadOnce(ByVal filePath As String, ByRef outText As String) As Boolean
    TryReadOnce = modUtilText.ReadTextFileUtf8(filePath, outText)
End Function

' 削除リトライ(並行GC耐性: 既に無い=達成として即成功)。
Private Function KillWithRetry(ByVal filePath As String) As Boolean
    Dim attempt As Long
    For attempt = 1 To 3
        On Error Resume Next
        Err.Clear
        If LenB(Dir(filePath)) = 0 Then
            On Error GoTo 0
            KillWithRetry = True
            Exit Function
        End If
        Kill filePath
        If Err.Number = 0 Then
            On Error GoTo 0
            KillWithRetry = True
            Exit Function
        End If
        On Error GoTo 0
        MentorWait 250 * attempt
    Next attempt
    ' 3回とも消せなかった。nonce重複排除があるので二重表示にはならないが、
    ' 質問ファイルが共有フォルダに残り続ける(掃除役は他にいない)。R11-D。
    On Error Resume Next
    modLog.LogUsage "mentor_kill_failed", "", _
        "質問ファイルを削除できませんでした: " & modUtil.SafeLeft(filePath, 200)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' 内部: 専門家の特定(スコア上位の出典から順に、pack作者を探す)
' ----------------------------------------------------------------------------
' 直近回答のヒットはスコア降順(modRetrieveがソート済み)なので、先頭から走査し
' 最初に見つかった「自分以外・不明以外」の出所を専門家とする。
'
' R33 W5-10(到達条件): 従来は origin が "pack:" で始まるヒットだけを候補に
'   していた。ところが origin に入りうる値は "self"(自力取込)/"channel:<部門名>"
'   (部門チャンネル同期)/"pack:<作者>"(手渡しパックを自分でファイル選択して
'   取り込んだときだけ)の3種類しかない。DESIGN_v3 が「既定は全チャンネル購読。
'   誰も操作しなくても自動的に届く」と定める主配布路を通った知識には "pack:" が
'   絶対に付かないので、標準的な運用の端末ではこの機能が原理的に一度も出ない
'   (出ないだけでエラーにならないため実機報告にも上がらない)。
'   modShareRule.OriginKind を通し、channel 由来も候補にする。
'
' R33 W5-11(宛先): 従来は origin から表示名を切り出しただけの文字列を宛先に
'   していた。受信側 CollectQuestions は q_<Fnv1a64Hex(CurrentUserId())>_*.txt
'   しか列挙しないため、表示名(初回起動で本人が打つ pack_author)とユーザID
'   (AD の CN / %USERNAME%)が偶然一致しない限り、書込みには成功するのに
'   誰にも届かず「送りました」とだけ表示されていた。感謝状(modP2P.EmitThanks)が
'   H-5/R8 F1 で既に確立した解決 ―― my_stats の pkauth:/chauth: を引く
'   modP2P.ResolveAuthorId ―― を必ず通し、解決できないときは候補にしない
'   (=ボタンを出さない。届かないものを「届いた」と言わないため)。
'   自己除外も、表示名どうしではなく解決済みIDどうしで比較する。
'   outExpert   : 画面に出す表示名(パック=作者名 / チャンネル=部門名)
'   outKind     : "pack" / "channel"(呼び方の出し分けに使う)
'   outTargetId : 実際の送信先ID(解決済み・サニタイズ済み)
Private Function FindExpert(ByRef outExpert As String, ByRef outSource As String, _
                            ByRef outKind As String, ByRef outTargetId As String) As Boolean
    Dim n As Long: n = modAsk.LastHitCount()
    If n <= 0 Then Exit Function

    Dim myId As String
    On Error Resume Next
    myId = modP2P.CurrentUserId()
    On Error GoTo 0

    Dim i As Long
    For i = 0 To n - 1
        Dim origin As String: origin = modAsk.LastHitOrigin(i)
        Dim kind As String: kind = modShareRule.OriginKind(origin)
        If LenB(kind) > 0 Then
            Dim shown As String: shown = Trim$(modShareRule.OriginName(origin))
            If LenB(shown) > 0 Then
                If StrComp(shown, "不明", vbTextCompare) <> 0 Then
                    Dim targetId As String
                    targetId = ""
                    On Error Resume Next
                    targetId = SanitizeId(Trim$(modP2P.ResolveAuthorId(origin, shown)))
                    On Error GoTo 0
                    If LenB(targetId) > 0 Then
                        If StrComp(targetId, myId, vbTextCompare) <> 0 Then
                            outExpert = SanitizeId(shown)
                            outKind = kind
                            outTargetId = targetId
                            outSource = modAsk.LastHitSource(i)
                            FindExpert = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Function

' ExpertLabel - 宛先の呼び方。パックは人名なので「〇〇 さん」、チャンネルは
'   部門名(人ではない)なので「「営業部」の発行者」と出す。R33 W5-10。
Private Function ExpertLabel() As String
    If StrComp(mExpertKind, "channel", vbTextCompare) = 0 Then
        ExpertLabel = "「" & mExpert & "」の発行者"
    Else
        ExpertLabel = mExpert & " さん"
    End If
End Function

' ----------------------------------------------------------------------------
' 内部: 質問ファイルの送信(共有フォルダI/O。modMentor内に完全カプセル化)
' ----------------------------------------------------------------------------
Private Function SendQuestion(ByVal expert As String, ByVal q As String, _
                              ByVal topSource As String) As Boolean
    ' R11-D(監査2 指摘3): ここは唯一ハンドラの無い入口だった。共有フォルダの
    ' MkDir や書込みで例外が出ると、呼び出し元(OnAskExpert/OnReplyQuestion)の
    ' Done: へ飛んで「送信できませんでした」だけが出て、原因が何も残らない。
    On Error GoTo Fail
    Dim folderPath As String: folderPath = QuestionsDir()
    If LenB(folderPath) = 0 Then
        ' 2026-08-10(R27波3-10): 共有フォルダが未設定(または未到達)。ここだけは
        ' 「失敗した理由が分かっている」唯一の分岐なのに、痕跡も残さず False を
        ' 返していたため、呼び出し元は原因の違う「ネットワーク接続を確認して」
        ' しか言えなかった。設定さえすれば直る人に、直せない案内を出さない。
        LogMentorErr "modMentor(SendQuestion)", 0, _
            "共有フォルダが未設定または未到達のため質問を送れませんでした"
        Exit Function
    End If
    On Error Resume Next
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    On Error GoTo 0

    Dim myId As String
    On Error Resume Next
    myId = modP2P.CurrentUserId()
    On Error GoTo 0
    If LenB(myId) = 0 Then myId = "user"

    ' ファイル名はハッシュのみ(サニタイズ+MAX_PATH二重達成)。payloadに生IDを保持。
    Dim nonce As String
    nonce = modP2PIo.IdHash(myId) & "-" & Format$(Now, "yyyymmddhhnnss") & "-" & _
            Format$(Int(Timer * 1000) Mod 100000, "00000")
    Dim rowText As String
    rowText = nonce & vbTab & myId & vbTab & expert & vbTab & _
              SanitizeField(q) & vbTab & SanitizeField(topSource) & vbTab & modUtil.NowStamp()

    SendQuestion = WriteUtf8WithRetry( _
        folderPath & "q_" & modP2PIo.IdHash(expert) & "_" & nonce & ".txt", rowText)
    If Not SendQuestion Then
        LogMentorErr "modMentor(SendQuestion)", 0, _
            "共有フォルダへ質問ファイルを書けませんでした(3回リトライ後)"
    End If
    Exit Function
Fail:
    LogMentorErr "modMentor(SendQuestion)", Err.Number, Err.Description
    SendQuestion = False
End Function

' ----------------------------------------------------------------------------
' LogMentorFail / LogMentorErr - 無言の失敗を1行だけ残す(R11-D / 監査2 指摘3)。
'   本モジュールは「失敗しても利用者には静かに何もしない」サーキットブレーカー
'   設計で、それ自体は正しい。しかし痕跡まで無いと、専門家への質問が届かない
'   ことに誰も気付けない(憲章§4-1)。通知は増やさず、ログだけを足す。
'   LogMentorFail は Err が立っているときだけ書く(正常な取りやめでは書かない)。
'   どちらも別Subなので、稼働中のハンドラの中から呼んでも On Error Resume Next
'   が正しく効く。
' ----------------------------------------------------------------------------
Private Sub LogMentorFail(ByVal context As String)
    If Err.Number = 0 Then Exit Sub
    Dim n As Long: n = Err.Number
    Dim d As String: d = Err.Description
    Err.Clear
    LogMentorErr context, n, d
End Sub

Private Sub LogMentorErr(ByVal context As String, ByVal errNum As Long, _
                         ByVal errDesc As String)
    On Error Resume Next
    modLog.LogError "E0801", context, _
        "err#" & errNum & ": " & modUtil.SafeLeft(errDesc, 300)
    On Error GoTo 0
End Sub

' 2026-07-31(レビュー R8 F2): modShare の関所を通す。自前で
' nexus_share_path を読むと、届かない共有でも毎回パスが組み上がり、
' 書込みリトライ(3回×バックオフ)を丸ごと払ってから失敗する。
' 到達判定とルート解決は modShare だけが行う(modShare 冒頭「唯一性の原則」)。
Private Function QuestionsDir() As String
    QuestionsDir = modShare.SubDir(QUESTIONS_SUBDIR)
End Function

' SendFailText - 送信に失敗したときに出す一文(2026-08-10 R27波3-10)。
'   共有フォルダが未設定・未到達なら、いくら通信を確かめても直らない。
'   実際に辿れる導線(❓ヘルプ →「⚙ 共有フォルダ設定」)を名指しで示す。
'   OnAskExpert/OnReplyQuestion の両方から同じ文を出す(2箇所で分岐を持たない)。
Private Function SendFailText() As String
    If LenB(QuestionsDir()) = 0 Then
        SendFailText = "共有フォルダが未設定です。" & ChrW(&H2753) & _
            "ヘルプの共有フォルダ設定から登録してください。"
    Else
        SendFailText = "送信できませんでした。ネットワーク接続を確認して、もう一度お試しください。"
    End If
End Function

' AVロック(エラー70等)に耐える書込みリトライ。modP2Pと同仕様の自前実装
' (Privateへ依存せず完全隔離。COMは両経路でSet=Nothing)。
Private Function WriteUtf8WithRetry(ByVal filePath As String, ByVal content As String) As Boolean
    Dim attempt As Long
    For attempt = 1 To 3
        If TryWriteOnce(filePath, content) Then
            WriteUtf8WithRetry = True
            Exit Function
        End If
        MentorWait 250 * attempt
    Next attempt
End Function

' UTF-8書き出しの実体は modUtilText.WriteTextFileUtf8(2026-07-31 R11-F2)。
Private Function TryWriteOnce(ByVal filePath As String, ByVal content As String) As Boolean
    TryWriteOnce = modUtilText.WriteTextFileUtf8(filePath, content)
End Function

Private Sub MentorWait(ByVal ms As Long)
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000# < ms
        DoEvents
        If Timer < t0 Then Exit Do   ' 深夜0時のTimerロールオーバーガード
    Loop
End Sub

' R33H F30: 「意図的な軽量重複」だった Private の複製を消し、modP2PIo.SanitizeId
' へ寄せた。ID の均し方が2箇所にあると、いつか必ず片方だけが更新される ――
' 実際、送信側だけがこれを通し受信側は生 ID をハッシュしていたため、CN の
' エスケープ(`\`)や64字超で「送りました」が嘘になっていた。
Private Function SanitizeId(ByVal s As String) As String
    SanitizeId = modP2PIo.SanitizeId(s)
End Function

' TSVフィールド安全化(タブ/改行→空白。modP2P.SanitizeFieldと同一仕様の複製)。
Private Function SanitizeField(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace(t, vbTab, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    SanitizeField = t
End Function

' 実機防衛(2026-07-21): modUIMain/modUIShelf.SafeRoundedRectと同じ理由・同じ
' 実装。座標(anchor.Left/.Top+.Height由来)を1未満に落ちないようクランプし、
' DoEvents+1回リトライ付きでAddShapeする(詳細はmodUIMain.bas側参照)。
Private Function SafeCoord(ByVal v As Double) As Double
    If v < 1 Then v = 1
    SafeCoord = v
End Function

Private Function SafeRoundedRect(ByVal ws As Worksheet, ByVal L As Double, ByVal T As Double, _
                                 ByVal W As Double, ByVal H As Double) As Shape
    L = SafeCoord(L): T = SafeCoord(T): W = SafeCoord(W): H = SafeCoord(H)
    On Error GoTo Retry
    Set SafeRoundedRect = ws.Shapes.AddShape(5, L, T, W, H)   ' 5=msoShapeRoundedRectangle(リテラル)
    Exit Function
Retry:
    DoEvents
    Set SafeRoundedRect = ws.Shapes.AddShape(5, L, T, W, H)
End Function
