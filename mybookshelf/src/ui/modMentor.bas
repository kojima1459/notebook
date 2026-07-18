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
'      modConfig のみ。自前状態は本モジュールPrivate(mExpert/mTopSource)に完結。
'   4. サニタイズ徹底: 専門家名・質問者名はSanitizeId(modP2Pと同一仕様の複製)で
'      禁止文字を除去し、さらにファイル名にはFnv1a64Hex(16桁)のみを使う
'      (MAX_PATH対策も同時達成。payload側に生のsanitize済IDを保持)。
'   ・送信I/OはmodMentor内に自前カプセル化(ADODB.Stream+リトライ+Set=Nothing)。
'     modP2PのPrivate関数には依存しない=modP2Pも1バイトも変更しない。
'   ・共有フォルダ書式: <nexus_share_path>\questions\q_<宛先hash>_<nonce>.txt
'     内容はTSV(nonce/from/to/質問/出典/時刻)。将来の受信機能(CollectQuestions)が
'     感謝状と同じcollect-then-process方式で読める形式にしてある。
' ============================================================================

Private Const QUESTIONS_SUBDIR As String = "questions"
Private Const BTN_NAME As String = "nx_mentor_btn"
Private Const MAX_Q_CHARS As Long = 1000       ' 質問文の上限(Shape/ファイル肥大防止)

' 本モジュール内で完結する状態(防衛条項3: 既存グローバルには一切触れない)
Private mExpert As String      ' 直近OfferMentorで特定した専門家(生sanitize済ID)
Private mTopSource As String   ' その専門家の代表ソース名(質問の文脈として同送)

' ----------------------------------------------------------------------------
' OfferMentor - エントリポイント(modApp.OnSend/OnActDrill末尾から1行フック)。
'   直近RAG回答の出典から専門家を特定できたときだけボタンを描く。
'   失敗しても静かに何もしない(サーキットブレーカー)。
' ----------------------------------------------------------------------------
Public Sub OfferMentor(ByVal bubbleName As String)
    On Error Resume Next   ' 安全弁: 本機能の失敗を絶対にメインへ波及させない
    ClearMentor

    Dim expert As String, topSource As String
    If Not FindExpert(expert, topSource) Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub

    ' 描画位置: 出典チップ(nx_cite_*)の最下端の下。チップが無ければバブルの下。
    Dim anchor As Shape
    Set anchor = ws.Shapes(bubbleName)
    If anchor Is Nothing Then Exit Sub
    Dim y As Double: y = anchor.Top + anchor.Height + 6
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 8) = "nx_cite_" Then
            If shp.Top + shp.Height + 6 > y Then y = shp.Top + shp.Height + 6
        End If
    Next shp

    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(5, anchor.Left, y, 360, 26)   ' 5=角丸四角
    btn.Name = BTN_NAME
    btn.Adjustments(1) = 0.4
    btn.Line.Visible = -1
    btn.Line.Weight = 1#
    btn.Line.ForeColor.RGB = modUI.UiColor("accent")           ' MS&ADグリーン
    btn.Fill.ForeColor.RGB = modUI.UiColor("surface")
    With btn.TextFrame2
        .WordWrap = -1
        .TextRange.Text = ChrW(&H1F4A1) & " この分野は " & expert & " さんが詳しいです [質問を送る]"
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
    mTopSource = topSource
End Sub

' ClearMentor - ボタン削除(送信の先頭フック/OfferMentor冒頭から。孤児Shape防止)。
Public Sub ClearMentor()
    On Error Resume Next   ' 安全弁
    ThisWorkbook.Worksheets("Nexus").Shapes(BTN_NAME).Delete
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

    Dim q As String
    q = InputBox("「" & mExpert & "」さんへの質問を入力してください。" & vbCrLf & _
                 "(共有フォルダ経由で届きます。関連資料: " & modUtil.SafeLeft(mTopSource, 40) & ")", _
                 "Nexus Agent - 専門家へ質問")
    q = Trim$(q)
    If LenB(q) = 0 Then GoTo Done
    If Len(q) > MAX_Q_CHARS Then q = Left$(q, MAX_Q_CHARS)

    If SendQuestion(mExpert, q, mTopSource) Then
        modSkin.ShowToast mExpert & " さんへ質問を送りました。回答をお待ちください。", "success"
        On Error Resume Next
        modLog.LogUsage "mentor_ask", "", "to=" & mExpert & " q=" & modUtil.SafeLeft(q, 120)
        On Error GoTo Done
    Else
        modSkin.ShowToast "送信できませんでした。ネットワーク接続を確認して、もう一度お試しください。", "error"
    End If
Done:
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
    Dim fn As String: fn = Dir(folderPath & "q_" & modUtil.Fnv1a64Hex(myId) & "_*.txt")
    Do While LenB(fn) > 0
        If nFiles > UBound(names) Then ReDim Preserve names(0 To UBound(names) + 32)
        names(nFiles) = fn
        nFiles = nFiles + 1
        fn = Dir()
    Loop

    ' 2) 読取り→nonce重複排除→バブル表示(最大3件。以降は件数のみ)→GC
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
                        modStats.Bump "mq:" & f(0)
                        newCount = newCount + 1
                        If shown < 3 Then
                            shown = shown + 1
                            On Error Resume Next
                            modUI.AddChatBubble "ai", _
                                ChrW(&H1F4EE) & " " & f(1) & " さんからあなた宛の質問が届いています。" & vbLf & _
                                "「" & modUtil.SafeLeft(f(3), 400) & "」" & vbLf & _
                                "(関連資料: " & modUtil.SafeLeft(f(4), 60) & " / " & f(5) & ")"
                            modLog.LogUsage "mentor_recv", "", "from=" & f(1) & " q=" & modUtil.SafeLeft(f(3), 120)
                            On Error GoTo Done
                        End If
                    End If
                    KillWithRetry full   ' 処理済み(既知含む)はGC。nonceで二重表示は防止済み
                End If
            End If
        End If
    Next i

    If newCount > 0 And Not silent Then
        modSkin.ShowToast "あなた宛の質問が " & newCount & " 件届いています。チャット欄をご確認ください。", "success"
    End If
Done:
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

Private Function TryReadOnce(ByVal filePath As String, ByRef outText As String) As Boolean
    Dim st As Object
    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile filePath
    outText = CStr(st.ReadText(-1))
    st.Close
    Set st = Nothing
    TryReadOnce = True
    Exit Function
Fail:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Set st = Nothing
    On Error GoTo 0
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
End Function

' ----------------------------------------------------------------------------
' 内部: 専門家の特定(スコア上位の出典から順に、pack作者を探す)
' ----------------------------------------------------------------------------
' 直近回答のヒットはスコア降順(modRetrieveがソート済み)なので、先頭から走査し
' 最初に見つかった「自分以外・不明以外のpack作者」を専門家とする。
Private Function FindExpert(ByRef outExpert As String, ByRef outSource As String) As Boolean
    Dim n As Long: n = modAsk.LastHitCount()
    If n <= 0 Then Exit Function

    Dim myId As String
    On Error Resume Next
    myId = modP2P.CurrentUserId()
    On Error GoTo 0

    Dim i As Long
    For i = 0 To n - 1
        Dim origin As String: origin = modAsk.LastHitOrigin(i)
        If LCase$(Left$(origin, 5)) = "pack:" Then
            Dim author As String: author = SanitizeId(Trim$(Mid$(origin, 6)))
            If LenB(author) > 0 Then
                If StrComp(author, "不明", vbTextCompare) <> 0 Then
                    If StrComp(author, myId, vbTextCompare) <> 0 Then
                        outExpert = author
                        outSource = modAsk.LastHitSource(i)
                        FindExpert = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

' ----------------------------------------------------------------------------
' 内部: 質問ファイルの送信(共有フォルダI/O。modMentor内に完全カプセル化)
' ----------------------------------------------------------------------------
Private Function SendQuestion(ByVal expert As String, ByVal q As String, _
                              ByVal topSource As String) As Boolean
    Dim folderPath As String: folderPath = QuestionsDir()
    If LenB(folderPath) = 0 Then Exit Function
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
    nonce = modUtil.Fnv1a64Hex(myId) & "-" & Format$(Now, "yyyymmddhhnnss") & "-" & _
            Format$(Int(Timer * 1000) Mod 100000, "00000")
    Dim rowText As String
    rowText = nonce & vbTab & myId & vbTab & expert & vbTab & _
              SanitizeField(q) & vbTab & SanitizeField(topSource) & vbTab & modUtil.NowStamp()

    SendQuestion = WriteUtf8WithRetry( _
        folderPath & "q_" & modUtil.Fnv1a64Hex(expert) & "_" & nonce & ".txt", rowText)
End Function

Private Function QuestionsDir() As String
    Dim basePath As String: basePath = modConfig.GetString("nexus_share_path", "")
    If LenB(basePath) = 0 Then Exit Function
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
    QuestionsDir = basePath & QUESTIONS_SUBDIR & "\"
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

Private Function TryWriteOnce(ByVal filePath As String, ByVal content As String) As Boolean
    Dim st As Object
    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2          ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText content
    st.SaveToFile filePath, 2   ' adSaveCreateOverWrite
    st.Close
    Set st = Nothing
    TryWriteOnce = True
    Exit Function
Fail:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Set st = Nothing
    On Error GoTo 0
End Function

Private Sub MentorWait(ByVal ms As Long)
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000# < ms
        DoEvents
        If Timer < t0 Then Exit Do   ' 深夜0時のTimerロールオーバーガード
    Loop
End Sub

' Windowsファイル名禁止文字の除去(modP2P.SanitizeIdと同一仕様の複製。Private依存を
' 避けるための意図的な軽量重複)。最大64字。
Private Function SanitizeId(ByVal s As String) As String
    Dim bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", ",", vbTab, vbCr, vbLf)
    Dim t As String: t = s
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        t = Replace(t, CStr(bad(i)), "_")
    Next i
    SanitizeId = modUtil.SafeLeft(Trim$(t), 64)
End Function

' TSVフィールド安全化(タブ/改行→空白。modP2P.SanitizeFieldと同一仕様の複製)。
Private Function SanitizeField(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace(t, vbTab, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    SanitizeField = t
End Function
