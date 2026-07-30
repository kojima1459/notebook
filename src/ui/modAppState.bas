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

' 対象バブル(選択中→無ければ最新のAI回答)があるか。無ければ案内してFalse。
Public Function HasTarget() As Boolean
    If LenB(TargetBubbleName()) = 0 Then
        ' 2026-07-28(レビュー M-25): 「吹き出しをクリックして選択」という
        ' 案内をやめた。図形保護を掛けたので選択できないし、そもそも
        ' 選択できた頃はDeleteで回答が消える操作へ誘導していた。
        MsgBox "対象のAI回答がありません。まず質問して回答を受け取ってください。", _
               vbInformation, "Nexus Agent"
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
    TargetText = modUI.BubbleTextOf(TargetBubbleName())
End Function

Public Function SharePath() As String
    SharePath = modConfig.GetString("nexus_share_path", SHARE_PATH_DEFAULT)
End Function

' 一般アシスタントモード: 本棚を介さずCallLLM直(会話履歴つき)。
' extraRules: 呼び出し文脈ごとの追加制約(空可)。本棚が空のときの
' 「社内固有の数字を断定させない」制約はここから注入される。
Public Function AskGeneral(ByVal q As String, ByVal extraRules As String) As String
    If LenB(mGenPrevU) = 0 Then
        mGenPrevU = modState.LoadState("nexus_gen_prevu", "")
        mGenPrevA = modState.LoadState("nexus_gen_preva", "")
    End If

    Dim sys As String
    sys = "あなたはMS&ADの最上位ナレッジコンシェルジュです。プロフェッショナルで簡潔、温かく頼りになるトーンで、" & _
          modConfig.GetString("answer_language", "日本語") & "で回答してください。" & vbLf & _
          "・必ず最初の1～2行で結論を言い切る(前置き・挨拶から始めない)。" & vbLf & _
          "・Markdown記号(#、**、`、表)は使わない(この画面では装飾されない)。" & _
          "見出しは「■ 」、箇条書きは「・」、最重要語だけ【 】で囲む。1ブロック3行以内。" & vbLf & _
          "・全体はおおむね200～400字。言い換えの繰り返しや締めの挨拶は書かない。" & vbLf & _
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
