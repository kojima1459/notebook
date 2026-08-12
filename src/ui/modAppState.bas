Attribute VB_Name = "modAppState"
Option Explicit

' ========================================
' modAppState - Nexus画面の状態の読み書きと、本棚を使わない一般回答
'
' modApp から切り出した裏方。ui_state シートの読み書き、入力セルの読み書き、
' モード表示、返信対象バブルの解決、本棚が空のときの一般回答。
' modApp 側はイベントハンドラ(OnSend/OnClearChat 等)の並びだけになる。
'
' 切り出しの理由(2026-07-28): modApp が契約上限30,000字に対し残り1,160字で、
' UIロック・言語表示・失敗時の入力復元(レビュー M-24/M-26/L-18/L-21/L-22)を
' まとめて入れる余裕が無かった(レビュー I-2)。
' ========================================

' 2026-07-28: modApp からここへ切り出した際、下記の定数とモジュール変数を
' 向こうに置いたままにしてしまい、実機で「変数が定義されていません」という
' コンパイルエラーになった。使う側であるここが持つのが正しい。
Private Const SHARE_PATH_DEFAULT As String = "\\pgiofs01\Nexus_Share\"
Public Const MODE_KEY As String = "nexus_mode"      ' rag / normal

' 一般モードの会話履歴(新しい順;;;区切り)。AskGeneral だけが読み書きする。
Private mGenPrevU As String
Private mGenPrevA As String

' 文脈アクション(👍/解決/深掘り等)の対象になる直近のAI回答バブル名。
' 【共有状態】書くのは modApp(回答を描いた直後)、読むのはここ。
' 変数を両方に置くと片方の書き込みがもう片方に見えないので、
' 持ち主をここ1つに決めて、modApp は SetActiveBubble 経由で書く。
Private mActiveBubble As String

' modApp が回答バブルを描いた直後に呼ぶ。
Public Sub SetActiveBubble(ByVal bubbleName As String)
    mActiveBubble = bubbleName
End Sub

' 本棚が空のときは「答えない」のではなく「何に基づく答えかをはっきり
' させて答える」。文面と安全指示は modLive が持つ(§modLive参照)。
Public Function ShelfIsEmpty() As Boolean
    On Error Resume Next
    ShelfIsEmpty = (modShelf.TotalChunks() = 0)
    On Error GoTo 0
End Function

Public Function AnswerWithoutShelf(ByVal q As String) As String
    On Error Resume Next
    modLive.PaintStage "一般知識でお答えしています…"
    On Error GoTo 0
    AnswerWithoutShelf = modLive.EmptyShelfWrap(AskGeneral(q, modLive.EmptyShelfGuard()))
End Function

' R20-2b: 一般アシスタントの「続けて質問」ゲート。modAsk.CanFollowup()と同型
' (遅延ロード)だがRAG側のmPrevUではなくmGenPrevUを見る。深掘りボタンが
' normalモードで常に「まず質問して…」になっていた不具合の是正。
Public Function HasGeneralMemory() As Boolean
    If LenB(mGenPrevU) = 0 Then
        mGenPrevU = modState.LoadState("nexus_gen_prevu", "")
    End If
    HasGeneralMemory = (LenB(mGenPrevU) > 0)
End Function

' R20-2d(新発見バグ): OnClearChatはRAG側の会話履歴は消していたが、一般
' アシスタント側(mGenPrevU/mGenPrevA)を消し忘れていた。会話クリア後も
' 一般アシスタントだけ前の話題を覚えている、という穴を塞ぐ。
Public Sub ClearGeneralMemory()
    mGenPrevU = ""
    mGenPrevA = ""
    On Error Resume Next
    modState.SaveState "nexus_gen_prevu", ""
    modState.SaveState "nexus_gen_preva", ""
    On Error GoTo 0
End Sub

' R26-2: modConvBridge(qa層・計算専任)の受け口。AskGeneralは毎回この変数を
' 読むため、ここで書けば次の一般アシスタント回答へそのまま効く
' (follow-up武装は不要)。
' R26H F3: 渡ってくるのは【橋渡しの1往復を先頭へ差し込んだ後の履歴全体】
' (";;;"区切り・新しい順)であって、1往復だけではない。以前の記憶を消さない
' のはmodConvBridge側の責務で、ここは計算済みの結果をそのまま持つだけ。
Public Sub SetGeneralMemory(ByVal u As String, ByVal a As String)
    mGenPrevU = u
    mGenPrevA = a
End Sub

' R26-2: モード切替の確定点(modApp.OnToggleMode)から1行で呼ぶ。実体の
' 計算は modConvBridge(qa層。modState/modConfigしか読まない計算専任)に
' 置き、ここでは戻り値を使ってUI層の書き込み(モジュール変数・
' modState.SaveState)とトースト通知だけを行う(modConvBridgeはmid層のため
' modSkin.ShowToastを直接呼べない=R1。呼び出し元をUI層側へ寄せる設計)。
' クリアの完全性: ここが書くキー(nexus_gen_prevu/preva・nexus_ask_prevu/
' preva)はいずれもmodApp.OnClearChatが既存のClearGeneralMemory呼び出しと
' 直接SaveStateで無条件に空へ戻す対象と完全に一致する(新規キーを増やして
' いない)。橋渡し後にクリアしても前の文脈は復活しない。
' R28-W4: RAG向きは modAsk のモジュール変数にも直接書くようになったため、
' OnClearChat 側も modAsk.ResetPrevMemory を呼ぶ(state だけ消すと亡霊が残る)。
Public Sub BridgeConvMemory(ByVal fromMode As String, ByVal toMode As String)
    On Error Resume Next
    Dim outQ As String, outA As String
    If Not modConvBridge.ComputeBridge(fromMode, toMode, outQ, outA) Then Exit Sub

    If toMode = "normal" Then
        SetGeneralMemory outQ, outA
        modState.SaveState "nexus_gen_prevu", outQ
        modState.SaveState "nexus_gen_preva", outA
    Else
        ' R28-W4(実機第13報⑥): SaveStateだけでは効かない。modAsk.CanFollowupの
        ' 遅延ロードは mPrevU が空のときしか ui_state を読まないため、切替前の
        ' RAG会話がメモリに残っていると、橋渡しした文脈ではなく古い会話が
        ' 深掘りに使われていた。モジュール変数へ直接渡す。SaveStateは
        ' 再起動後の復元用として維持する(消さない)。
        modAsk.SetPrevMemory outQ, outA
        modState.SaveState "nexus_ask_prevu", outQ
        modState.SaveState "nexus_ask_preva", outA
    End If

    ' chars は【差し込み後の記憶全体】の字数(R26H F3で置換→先頭差し込みに
    ' 変わったため、橋渡しした1往復だけの字数ではない)。
    modLog.LogUsage "memory_carry", fromMode & "->" & toMode, "chars=" & (Len(outQ) + Len(outA))
    modSkin.ShowToast BridgeToastText(toMode), "info", True
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' BridgeToastText - 橋渡しの通知文(向きで言えることが違う)。
' ----------------------------------------------------------------------------
' R26H F4(M-2 A): 一般→社内ナレッジ検索の向きでも「引き継ぎました」と出して
' いたが、これは実態と違う。modAsk.Answer(単発質問)は毎回 prevU/prevA へ
' 空文字を明示的に渡すため、保存した1往復が実際に効くのは🔍深掘り(続けて
' 質問)を押して AskFollowup を通したときだけ(modConvBridge 冒頭の既知の
' 非対称。凍結モジュール制約下の現実解)。引き継がれたと思って質問した人が
' 「覚えていない」と感じるのは、この一文が先に嘘をついているため。
' できることだけを言い、次の一手を名指しする。
'
' R29 W1-6: 「直前の1往復」という件数の明示をやめた。R28 で運搬は
' followup_max_pairs 件(既定5往復)まで広がったのに、この文言だけが1往復の
' ままで残り、利用者に「1往復しか渡らない」と読ませていた(誤表記)。件数は
' config で動くので、文言では件数を言わない。シグネチャは不変。
Public Function BridgeToastText(ByVal toMode As String) As String
    If toMode = "rag" Then
        BridgeToastText = "直前の会話を保存しました。続きとして聞くには " & _
            ChrW(&HD83D) & ChrW(&HDD0D) & " 深掘り(続けて質問)を押してください"
    Else
        BridgeToastText = "直前の会話を引き継ぎました。新しく始めるなら " & _
            ChrW(&HD83D) & ChrW(&HDDD1) & " クリア"
    End If
End Function

' R26-1同梱2(波A裁定): 入念モード開始前の事前案内はRAG専用の文言で固定
' されていた(「多方向から検索して…」)。一般アシスタントは検索も出典照合も
' 行わないため、sendMode別に出し分ける(modAppは容量残り僅かのためここへ
' 置き、呼び出し側は1行にする)。
Public Function ThoroughPreNotice(ByVal sendMode As String) As String
    If sendMode = "normal" Then
        ThoroughPreNotice = "多段の自己検証で答えを練ります。数分かかることがあります…"
    Else
        ThoroughPreNotice = "入念に調べます。多方向から検索して、資料と1行ずつ照合します…"
    End If
End Function

' 対象バブル(選択中→無ければ最新のAI回答)があるか。無ければ案内してFalse。
Public Function HasTarget() As Boolean
    If LenB(TargetBubbleName()) = 0 Then
        ' 2026-07-28(レビュー M-25): 「吹き出しをクリックして選択」という
        ' 案内をやめた。図形保護を掛けたので選択できないし、そもそも
        ' 選択できた頃はDeleteで回答が消える操作へ誘導していた。
        MsgBox "対象のAI回答がありません。まず質問して回答を受け取ってください。", _
               vbInformation, modAppDef.APP_NAME
        Exit Function
    End If
    HasTarget = True
End Function

Public Function TargetBubbleName() As String
    If LenB(mActiveBubble) > 0 Then
        If LenB(modUI.BubbleTextOf(mActiveBubble)) > 0 Then
            TargetBubbleName = mActiveBubble
            Exit Function
        End If
    End If
    TargetBubbleName = modUI.LatestAiBubbleName()
End Function

Public Function TargetText() As String
    ' R14-8c: バブルの段落区切りは vbCr(Shape の Paragraphs は vbCr でしか
    ' 分かれないため)。ここから先はコピー・Word化・確認プレビューといった
    ' 「ただの文章」として扱われるので、読み出した時点で vbLf へ戻す。
    ' 戻さないと Word へ渡すプロンプトや貼り付け先で改行が1行に潰れて見える。
    TargetText = Replace(Replace(modUI.BubbleTextOf(TargetBubbleName()), vbCrLf, vbLf), vbCr, vbLf)
End Function

Public Function SharePath() As String
    SharePath = modConfig.GetString("nexus_share_path", SHARE_PATH_DEFAULT)
End Function

' R20-2e: maxPairs<=0(機能OFFのエスケープハッチ)の境界判定だけを純関数に
' 出す(LOテストで0/負/正の境界を固定する)。
Public Function ShouldClearGeneralHistory(ByVal maxPairs As Long) As Boolean
    ShouldClearGeneralHistory = (maxPairs <= 0)
End Function

' 一般アシスタントモード: 本棚を介さずCallLLM直(会話履歴つき)。
' extraRules: 呼び出し文脈ごとの追加制約(空可)。本棚が空のときの
' 「社内固有の数字を断定させない」制約はここから注入される。
'
' R26-1(2026-08-11): 第3引数 mode でヘッダーの速さトグルを受ける。
'   quick    … 従来と1文字も変えない1回呼び出し(既定。引数を渡さない
'              既存の呼び出し=AnswerWithoutShelf もここへ落ちる)
'   deep     … 1回呼び出し + 構造化指示(gen_deep_effort/verbosity)
'   thorough … modGenPipe.RunThorough(起草→検証→改稿)へ委譲
' 分岐表そのものは modGenPipe.PlanFor が単一情報源(UI層に表を持たない)。
Public Function AskGeneral(ByVal q As String, ByVal extraRules As String, _
                           Optional ByVal mode As String = "quick") As String
    ' 前ターンの検証周回数を必ず捨てる。捨てないと、入念の次に出した
    ' すぐ聞くの回答フッターへ「検証2回」が漏れる(R20H FA-8 と同型)。
    Dim genMode As String: genMode = modMode.Normalize(mode)
    modGenPipe.ResetTurn genMode
    Dim t0 As Double: t0 = Timer

    If LenB(mGenPrevU) = 0 Then
        mGenPrevU = modState.LoadState("nexus_gen_prevu", "")
        mGenPrevA = modState.LoadState("nexus_gen_preva", "")
    End If

    ' R20-2e: followup_max_pairs<=0はRAG側AppendFollowupPairと同型の
    ' エスケープハッチ。まだ送っていない古い履歴を握ったままではmaxPairsを
    ' 絞った意味が無いので、このターンの送信に含める【前】に消す。
    Dim maxPairs As Long
    maxPairs = modConfig.GetLong("followup_max_pairs", 3)
    If ShouldClearGeneralHistory(maxPairs) Then
        mGenPrevU = ""
        mGenPrevA = ""
        On Error Resume Next
        modState.SaveState "nexus_gen_prevu", ""
        modState.SaveState "nexus_gen_preva", ""
        On Error GoTo 0
    End If

    Dim sys As String
    sys = "あなたはMS&ADの最上位ナレッジコンシェルジュです。プロフェッショナルで簡潔、温かく頼りになるトーンで、" & _
          modConfig.GetString("answer_language", "日本語") & "で回答してください。" & vbLf & _
          "・必ず最初の1～2行で結論を言い切る(前置き・挨拶から始めない)。" & vbLf & _
          "・Markdown記号(#、**、`、表)は使わない(この画面では装飾されない)。" & _
          "見出しは「■ 」、箇条書きは「・」、最重要語だけ【 】で囲む。1ブロック3行以内。" & vbLf & _
          "・全体はおおむね200～400字。言い換えの繰り返しや締めの挨拶は書かない。" & vbLf & _
          "・専門用語には短い補足を()で添え、初めて読む人にも一度で伝わる言葉を選ぶ。" & vbLf & _
          "・直前の会話が示されている場合は、その続きとして答える。" & _
          "「これ」「その」「さっきの」等の指示語は直前の会話の内容を指すものとして解釈する。"
    If LenB(extraRules) > 0 Then sys = sys & vbLf & extraRules

    Dim lat As Long
    Dim resp As String
    Select Case modGenPipe.PlanFor(genMode)
        Case modGenPipe.PLAN_PIPELINE
            ' 入念: 起草→検証→改稿。段の実況も modGenPipe が出す。
            resp = modGenPipe.RunThorough(q, sys, mGenPrevU, mGenPrevA)
        Case modGenPipe.PLAN_SINGLE_DEEP
            ' しっかり: 呼び出しは1回のまま、指示とeffort/verbosityだけ差し替える
            ' (high と medium の差は一般タスクでは薄く、構造化のほうが効く)。
            resp = modGateway.CallLLM(sys & vbLf & modGenPipe.DeepRules() & _
                vbLf & vbLf & "## 質問" & vbLf & q, "nexus_general_deep", _
                modConfig.GetString("gen_deep_effort", "medium"), _
                modConfig.GetString("gen_deep_verbosity", "high"), _
                modConfig.GetString("quick_model", "gpt-5.5"), lat, mGenPrevU, mGenPrevA)
        Case Else
            resp = modGateway.CallLLM(sys & vbLf & modGenPipe.QuickRules() & _
                vbLf & vbLf & "## 質問" & vbLf & q, "nexus_general", _
                modConfig.GetString("quick_effort", "low"), _
                modConfig.GetString("quick_verbosity", "low"), _
                modConfig.GetString("quick_model", "gpt-5.5"), lat, mGenPrevU, mGenPrevA)
    End Select

    ' モードごとの実測(所要ms・検証周回数・PASS到達)は成否に関わらず毎回残す。
    ' 「3モードが同じことをしている」という疑いは、記録が無ければ永久に晴れない。
    On Error Resume Next
    modLog.LogUsage "gen_mode", genMode, modGenPipe.TurnDetail(), _
        CLng(modUtilText.ElapsedMsSince(t0)), modGenPipe.LastLoops()
    On Error GoTo 0

    If Left$(resp, 5) = "#ERR:" Then
        ' R14-G2: 失敗したターンは「直近の回答」を持たない。ここで印を消さないと、
        ' この後に押された感想ボタンが【前の質問】の状態(mLastQuestion/mLastHits)
        ' で受理され、失敗した回答に対する「解決した」が前の資料の感謝状や
        ' 部内共有まで撃ってしまう。回答前と同じ「まず質問して…」で断らせる。
        On Error Resume Next
        modAsk.NoteAnswerFailed
        On Error GoTo 0
        AskGeneral = "回答の作成に失敗しました。時間を置いてもう一度お試しください。"
        Exit Function
    End If

    ' チャット履歴シート記録(modChatLog、core層。書込失敗で死なない設計)。
    On Error Resume Next
    modChatLog.LogTurn q, resp, "general"
    On Error GoTo 0

    ' 会話履歴(新しい順;;;区切り・最大followup_max_pairsペア。maxPairsは冒頭で読済み)
    If maxPairs > 0 Then
        mGenPrevU = TrimPairs(q & IIf(LenB(mGenPrevU) > 0, ";;;" & mGenPrevU, ""), maxPairs)
        mGenPrevA = TrimPairs(modUtil.SafeLeft(resp, 2000) & IIf(LenB(mGenPrevA) > 0, ";;;" & mGenPrevA, ""), maxPairs)
    End If
    modState.SaveState "nexus_gen_prevu", mGenPrevU
    modState.SaveState "nexus_gen_preva", mGenPrevA
    ' R14-1b: 一般モードで答えたことを「直近の回答」として qa 層へ知らせる。
    ' これが無いと、この直後の「解決した」が拒否されるか、前のRAG質問の
    ' 回答を解決したことにされる(実機第3報 RC2)。
    ' R14-G1: 回答本文も一緒に渡す(残った前回のRAG回答を読む経路を潰す)。
    modAsk.NoteGeneralAnswered q, resp
    AskGeneral = resp
End Function

' ";;;"区切り文字列を先頭maxN件へ切り詰める。
Public Function TrimPairs(ByVal s As String, ByVal maxN As Long) As String
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
Public Function RagSpeed() As String
    ' 3モードの正規化は modMode が単一情報源(quick/deep/thorough)。
    RagSpeed = modMode.Normalize(ReadUiState("mode", "quick"))
End Function

Public Function CurrentMode() As String
    CurrentMode = ReadUiState(MODE_KEY, "rag")
    If CurrentMode <> "normal" Then CurrentMode = "rag"
End Function

' 2026-07-30(レビュー3-A(1)): ピルへ直接テキストを書き込むのをやめ、
' ヘッダー全体を描き直す。文字だけを差し替えると幅計算(ティア選択と
' FlowRight)を通らないため、長い語に変わった瞬間にピルからはみ出し、
' 隣のピルやタイトルへ重なる(実機のヘッダー崩れの直接原因)。
Public Sub UpdateModeButton()
    On Error Resume Next
    modUINexusDraw.RedrawChatHeader
    ' R27 F2-1c(実機第12報①): ヘッダーの行高が変わると、ptから換算した境界
    ' (NexusBound)が指す行がずれる。描き直したら必ず貼り直す。
    modUINexusDraw.ReapplyChatBound
    On Error GoTo 0
End Sub

Public Function ReadInputCell() As String
    On Error Resume Next
    Dim v As Variant
    v = ThisWorkbook.Names("nx_input").RefersToRange.Value
    On Error GoTo 0
    If IsEmpty(v) Or IsError(v) Then Exit Function
    ReadInputCell = modUtil.SafeLeft(CStr(v), 3000)
End Function

' 入力欄へ文章を書き戻す(送信に失敗したときの復元。レビュー L-22)。
' 読み書きは ClearInputCell/ReadInputCell と同じ名前定義 nx_input を使う
' (座標を別に持つと、レイアウト変更で片方だけずれる)。
Public Sub RestoreInputCell(ByVal s As String)
    On Error Resume Next
    ThisWorkbook.Names("nx_input").RefersToRange.Value = s
    On Error GoTo 0
End Sub

Public Sub ClearInputCell()
    On Error Resume Next
    ThisWorkbook.Names("nx_input").RefersToRange.Value = ""
    On Error GoTo 0
End Sub

Public Function ReadUiState(ByVal keyName As String, ByVal defaultVal As String) As String
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

Public Sub WriteUiState(ByVal keyName As String, ByVal valText As String)
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
