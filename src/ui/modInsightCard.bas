Attribute VB_Name = "modInsightCard"
Option Explicit

' ============================================================================
' modInsightCard - 洞察カード保存(2026-08-11 R26-3 波C)
' ----------------------------------------------------------------------------
' 「いま得た答え」を本棚のストックに足す導線。回答直下の文脈アクション行に
' 追加した 💾保存 から呼ばれ、直前の1往復(質問+回答)を通常の取込パイプライン
' へ流して my_knowledge へ登録する(RAG回答・一般アシスタント回答のどちらでも
' 保存できる)。以後は検索でヒットし、出典としても引かれる。
'
' 取込経路(調査の結論): テキストを直接ナレッジ化する口は既に1本ある
' ―― modVault.RegisterKnowledgeText。TEMPへUTF-8のtxtを書き出し、
' modShelf.IngestFile("self") へ渡す(=抽出→チャンク分割(modChunker)→
' 埋め込み(modEmbed)→my_knowledge/my_vectors/manifest登録まで既存の一本道)。
' 👎の修正ナレッジ(modAppAct)と共有フォルダ取込(modShared)が既に同じ口を
' 使っている。ここを再利用し、新しい取込実装は1行も書かない
' (modChunker/modShelf/modShelfStore は凍結・逼迫のため不触)。
'
' 汚染防止(spec §3): 保存した考察は【一次資料ではない】。凍結制約下
' (modRetrieve/modPrompts不触)で成立する防御は出所ラベル層だけなので、
'   (1) 資料名を「考察メモ_題名」に固定する(先頭の💭が検索結果・出典表示・
'       本棚カードの全てに出る=一目で一次資料と区別が付く)
'   (2) 本文の冒頭に「一次資料ではない/出典として扱わない」定型文を置く
' の2枚を必ず重ねる。(1)は資料名=取り込んだファイル名という既存仕様に
' 乗るため、modVault.RegisterKnowledgeText の第4引数(docBase)で決める。
'
' 境界の関所(R27 §7 繋ぎ目1): 💾は文脈アクション行(nx_act_*)の7個目として
' 描く。文脈アクション行は modApp.OnSend の中で SettleChat より前に描かれ、
' 境界(ExtendChatBand)の内側に必ず入る=「押せない・スクロールで消える」の
' 逆走が構造的に起きない。だから新しいShape機構は作らない。
' ============================================================================

' 資料名の接頭辞。💭(U+1F4AD)は非BMPのためConstに置けない(サロゲートペア)。
' 本棚カード・検索結果・出典表示のいずれも資料名の先頭から描くので、この
' 2文字が「これは一次資料ではない」の一次シグナルになる。
Private Const MEMO_LABEL As String = "考察メモ_"

' 本文冒頭の定型文(汚染防止)。LLMがこのメモを出典へ昇格させるのを抑止する。
Private Const MEMO_NOTICE As String = _
    "[このメモはAIとの対話から保存された考察であり、一次資料ではありません。" & _
    "出典として扱わず、根拠は原典を確認すること]"

' 題名の既定値に使う質問文の長さ(spec §3: 先頭24字)。
Private Const TITLE_HEAD_CHARS As Long = 24

' 題名の上限。資料名は「💭考察メモ_」+題名で、modVault.SanitizeName が
' 60字で切るため、接頭辞(7字)を足しても切られない長さに抑える。
Private Const TITLE_MAX_CHARS As Long = 40

' ----------------------------------------------------------------------------
' MemoPrefix - 資料名の接頭辞「💭考察メモ_」。
' ----------------------------------------------------------------------------
Public Function MemoPrefix() As String
    MemoPrefix = ChrW(&HD83D) & ChrW(&HDCAD) & MEMO_LABEL
End Function

' ----------------------------------------------------------------------------
' DocNameFor - 題名から資料名(=取り込むファイルのベース名)を作る。
'   題名は必ず SafeTitle を通す(ファイル名禁止文字・長さ)。
' ----------------------------------------------------------------------------
Public Function DocNameFor(ByVal titleText As String) As String
    DocNameFor = MemoPrefix() & SafeTitle(titleText)
End Function

' ----------------------------------------------------------------------------
' TitleFromQuestion - 題名の既定値(質問文の先頭24字)。改行・タブは空白へ
'   潰す(InputBoxの既定値が複数行になると編集できない)。全角安全な切り出しは
'   modUtil.SafeLeft(サロゲートの片割れを残さない)。
' ----------------------------------------------------------------------------
Public Function TitleFromQuestion(ByVal q As String) As String
    Dim t As String
    t = Replace(Replace(Replace(q, vbCr, " "), vbLf, " "), vbTab, " ")
    t = Trim$(t)
    TitleFromQuestion = Trim$(modUtil.SafeLeft(t, TITLE_HEAD_CHARS))
End Function

' ----------------------------------------------------------------------------
' SafeTitle - 題名をファイル名・資料名として安全な形へ整える。
'   ・ファイル名禁止文字(\ / : * ? " < > |)と改行・タブは "_" へ
'   ・末尾の "." と空白は落とす(Windowsは末尾ドットのファイルを作れない)
'   ・TITLE_MAX_CHARS で全角安全に切る
'   数式インジェクション(先頭 = + - @)への配慮は不要: 資料名は必ず
'   MemoPrefix() の💭で始まるため、セルへ書かれる値が数式化する形にならない。
' ----------------------------------------------------------------------------
Public Function SafeTitle(ByVal s As String) As String
    Dim bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)
    Dim t As String: t = s
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        t = Replace(t, CStr(bad(i)), "_")
    Next i
    t = Trim$(t)
    ' 末尾のドット/空白は落とす(1文字ずつ。Doで回すのは"..."対策)
    Do While Len(t) > 0
        If Right$(t, 1) <> "." And Right$(t, 1) <> " " Then Exit Do
        t = Left$(t, Len(t) - 1)
    Loop
    SafeTitle = Trim$(modUtil.SafeLeft(t, TITLE_MAX_CHARS))
End Function

' ----------------------------------------------------------------------------
' MemoNotice - 本文冒頭の定型文(汚染防止)。
' ----------------------------------------------------------------------------
Public Function MemoNotice() As String
    MemoNotice = MEMO_NOTICE
End Function

' ----------------------------------------------------------------------------
' WithMemoNotice - 本文の冒頭へ定型文を付ける。既に付いていれば付け直さない
'   (冪等: 同じ会話を2回保存しても定型文が積み重ならない)。
' ----------------------------------------------------------------------------
Public Function WithMemoNotice(ByVal bodyText As String) As String
    If Left$(bodyText, Len(MEMO_NOTICE)) = MEMO_NOTICE Then
        WithMemoNotice = bodyText
        Exit Function
    End If
    WithMemoNotice = MEMO_NOTICE & vbLf & vbLf & bodyText
End Function

' ----------------------------------------------------------------------------
' MemoBody - 保存する本文(定型文 + 質問 + 回答)。見出しは画面と同じ「■ 」
'   記法で、再取込しても読みやすい素のテキストにする。
' ----------------------------------------------------------------------------
Public Function MemoBody(ByVal q As String, ByVal a As String) As String
    MemoBody = WithMemoNotice("■ 質問" & vbLf & q & vbLf & vbLf & _
                              "■ AIの回答" & vbLf & a)
End Function

' ----------------------------------------------------------------------------
' SaveLastTurn - 💾保存の実体(再入の関所は呼び出し元 modAppAct.OnActSaveInsight
'   が持つ。ここは Enter/Leave を触らない)。
'   失敗は全て usage_log へ残す(無言失敗禁止・憲章§4-1)。
' ----------------------------------------------------------------------------
Public Sub SaveLastTurn()
    Dim mode As String: mode = ModeTag()
    Dim t0 As Double: t0 = Timer

    ' 対象のAI回答が無ければ HasTarget が案内を出して False(既存3ボタンと同じ門)。
    If Not modAppState.HasTarget() Then
        LogFail mode, "対象のAI回答が無い"
        Exit Sub
    End If

    Dim q As String, a As String
    If Not LastTurn(q, a) Then
        LogFail mode, "直前の1往復を取得できない"
        modSkin.ShowToast "保存できる会話が見つかりませんでした。もう一度質問してからお試しください。", "info"
        Exit Sub
    End If

    ' ここから先はダイアログ・取込を伴う。例外は必ず1箇所(Fail)で受けて
    ' 進捗バナーを消す(R19③: バナーの残留は「固まった」に見える)。
    On Error GoTo Fail

    Dim titleText As String
    ' 既定値=質問の先頭24字。ここは絵文字を出さない(ネイティブダイアログは
    ' 非BMP文字を"?"に化けさせる。lint検査 check_msgbox_nonbmp)。
    titleText = InputBox( _
        "この会話を本棚へ保存します。メモの題名を入力してください。" & vbCrLf & _
        "本棚には「考察メモ_題名」として並び、AIの答えの材料になります。" & vbCrLf & _
        "(一次資料ではない考察メモとして、資料名で区別されます)", _
        modAppDef.APP_NAME & " - この会話を本棚に保存", TitleFromQuestion(q))

    If LenB(Trim$(titleText)) = 0 Then
        On Error Resume Next
        modLog.LogUsage "insight_cancel", mode, "題名の入力を中断"
        On Error GoTo 0
        Exit Sub
    End If

    titleText = SafeTitle(titleText)
    If LenB(titleText) = 0 Then
        LogFail mode, "題名が禁止文字だけ"
        modSkin.ShowToast "その題名は使えません。文字を変えてもう一度お試しください。", "info"
        Exit Sub
    End If

    Dim docName As String: docName = DocNameFor(titleText)
    ' 同じ題名の考察メモが既にあると、取込は「同じ資料の入れ直し」になり
    ' 前のメモが置き換わる(黙って消えたように見える)。必ず先に聞く。
    ' 資料名は非BMPの💭で始まるのでダイアログ本文には出さない(化ける)。
    If HasSameMemo(docName) Then
        If MsgBox("同じ題名の考察メモが既に本棚にあります。" & vbCrLf & _
                  "上書きして更新しますか?" & vbCrLf & _
                  "(いいえ=保存しません。題名を変えてもう一度お試しください)", _
                  vbQuestion + vbYesNo + vbDefaultButton2, modAppDef.APP_NAME) <> vbYes Then
            On Error Resume Next
            modLog.LogUsage "insight_cancel", mode, "同名メモの上書きを中止"
            On Error GoTo 0
            Exit Sub
        End If
    End If

    Dim body As String: body = MemoBody(q, a)
    Dim chunksBefore As Long: chunksBefore = SafeTotalChunks()

    modUIMain.ShowProgress "考察メモを本棚へ保存しています…(埋め込みの作成に少し時間がかかります)"
    Dim ok As Boolean
    ok = modVault.RegisterKnowledgeText(titleText, body, "考察メモ", docName)
    modUIMain.HideProgress

    If Not ok Then
        LogFail mode, "取込パイプラインが失敗(title=" & titleText & ")"
        modSkin.ShowToast "保存に失敗しました。マイ本棚の一覧で状態をご確認ください。", "error"
        Exit Sub
    End If

    Dim added As Long: added = SafeTotalChunks() - chunksBefore
    On Error Resume Next
    modLog.LogUsage "insight_saved", mode, _
        "title=" & titleText & " chars=" & Len(body) & " chunks=" & added, _
        CLng(modUtilText.ElapsedMsSince(t0)), added
    On Error GoTo 0
    modSkin.ShowToast ChrW(&HD83D) & ChrW(&HDCAD) & " 考察メモとして本棚に保存しました。" & _
        "次からの検索にも使われます。", "success"
    Exit Sub

Fail:
    ' Err は Resume/On Error 文でクリアされる。後始末の前に必ず退避する。
    Dim failNum As Long: failNum = Err.Number
    Dim failDesc As String: failDesc = Err.Description
    Resume FailCleanup
FailCleanup:
    On Error Resume Next
    modUIMain.HideProgress
    modLog.LogUsage "insight_save_fail", mode, _
        "例外: " & failDesc & " (Err=" & failNum & ")"
    modLog.LogError "E0801", "modInsightCard.SaveLastTurn", failDesc, failNum
    modSkin.ShowToast "保存に失敗しました。時間を置いてもう一度お試しください。", "error"
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' LastTurn - 直前1往復(質問/回答)を取り出す。凍結モジュール(modAsk)の
'   mLast系にはPublicアクセサ経由でしか触らない。
'
' R26H F8(m-4): 質問と回答の第一情報源を nexus_hist_u / nexus_hist_a に統一した。
'   このキーは modApp.SaveTurnForRestore が毎ターン【対で】書く唯一のキーで、
'   質問と回答が必ず同じターンのものになる。旧実装は質問=モード別の記憶キー・
'   回答=modAsk.LastAnswerText() という別々の情報源から取っており、
'   conv_bridge=off でモードを切り替えた直後などに、別ターンの質問と別ターンの
'   回答が1枚の考察メモへ混ざりうる(本棚に入って検索・出典に出るものなので、
'   混ざったまま残ると後から誰も気付けない)。
'   フォールバックは従来のまま:
'     質問: モード別の記憶キー(nexus_ask_prevu / nexus_gen_prevu)の先頭1件。
'     回答: modAsk.LastAnswerText() → 画面のバブル本文(modAppState.TargetText)。
'   区切りの解釈は modConvBridge.FirstPair が単一情報源。
'
' R26H F8 レビューFix(提案a): nexus_hist_a を第一情報源にした副作用として、
'   modApp.SaveTurnForRestore が保存時に回答を700字で切る(SafeLeft(ans,700))
'   ため、考察メモの本文が700字ちょうどで無言に途切れていた。修正後は
'   modAsk.LastAnswerText()(画面のバブル本文=フル)が hist_a を先頭部分として
'   含む場合―つまり同一ターンの回答である場合―だけフル本文を採用する。
'   含まない(=別ターンの回答が入っている)場合は従来どおり hist_a を採用し、
'   別ターン混成の防止は崩さない。判定は PickFullerAnswer が単一情報源。
' ----------------------------------------------------------------------------
Private Function LastTurn(ByRef q As String, ByRef a As String) As Boolean
    On Error Resume Next
    q = modConvBridge.FirstPair(modState.LoadState("nexus_hist_u", ""))
    a = modConvBridge.FirstPair(modState.LoadState("nexus_hist_a", ""))

    If LenB(q) = 0 Then
        Dim keyU As String
        If modAppState.CurrentMode() = "normal" Then
            keyU = "nexus_gen_prevu"
        Else
            keyU = "nexus_ask_prevu"
        End If
        q = modConvBridge.FirstPair(modState.LoadState(keyU, ""))
    End If

    a = PickFullerAnswer(a, modAsk.LastAnswerText())
    If LenB(a) = 0 Then a = modAppState.TargetText()
    On Error GoTo 0
    ' 回答が取れないなら保存する意味が無い(質問だけのメモは資料にならない)。
    LastTurn = (LenB(a) > 0)
End Function

' ----------------------------------------------------------------------------
' PickFullerAnswer - hist_a(nexus_hist_a由来。700字で切られている場合がある)と
'   fullA(modAsk.LastAnswerText()。画面のバブル本文=フル)のどちらを採用するか
'   を決める純関数。
'   ・histA が空 → fullA を採用(そもそも比較できない)。
'   ・fullA が histA を先頭部分として含む(Left$(fullA, Len(histA)) = histA)
'     → 同一ターンの回答とみなし、長い方(=通常fullA。等しければどちらでも
'       同じ内容)を採用する。histA が700字未満(=切られていない)なら
'       Len(fullA) = Len(histA) の完全一致になるので、この分岐でも安全に
'       fullA を返せる。
'   ・含まない → 別ターンの回答が fullA に残っている(モード切替直後などの
'     すれ違い)とみなし、hist_a を採用する(別ターン混成の防止)。
' ----------------------------------------------------------------------------
Public Function PickFullerAnswer(ByVal histA As String, ByVal fullA As String) As String
    If LenB(histA) = 0 Then
        PickFullerAnswer = fullA
        Exit Function
    End If
    If LenB(fullA) = 0 Then
        PickFullerAnswer = histA
        Exit Function
    End If
    If Left$(fullA, Len(histA)) = histA Then
        If Len(fullA) >= Len(histA) Then
            PickFullerAnswer = fullA
        Else
            PickFullerAnswer = histA
        End If
    Else
        PickFullerAnswer = histA
    End If
End Function

' 同じ資料名(=同じ題名)の考察メモが既に本棚にあるか。manifest の走査は
' modShelfStore が単一情報源で、newPath="" を渡すと「そのファイル名を持つ
' self行のパス」がそのまま返る(空パスの行は存在しないため存在判定に使える)。
Private Function HasSameMemo(ByVal docName As String) As Boolean
    On Error Resume Next
    HasSameMemo = (LenB(modShelfStore.FindConflictingManifestPath(docName & ".txt", "")) > 0)
    On Error GoTo 0
End Function

' 本棚の総チャンク数(失敗しても0で続ける)。保存前後の差分=このメモの
' チャンク数として usage_log に残す。
Private Function SafeTotalChunks() As Long
    On Error Resume Next
    SafeTotalChunks = modShelf.TotalChunks()
    On Error GoTo 0
End Function

' 記録用のモード名("rag"/"normal")。
Private Function ModeTag() As String
    On Error Resume Next
    ModeTag = modAppState.CurrentMode()
    On Error GoTo 0
    If LenB(ModeTag) = 0 Then ModeTag = "rag"
End Function

Private Sub LogFail(ByVal mode As String, ByVal reason As String)
    On Error Resume Next
    modLog.LogUsage "insight_save_fail", mode, reason
    On Error GoTo 0
End Sub
