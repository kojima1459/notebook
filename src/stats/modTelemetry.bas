Attribute VB_Name = "modTelemetry"
Option Explicit

' ============================================================================
' modTelemetry - 利用データの集計と、匿名フィードバック
' ----------------------------------------------------------------------------
' 何を集めるか(プロダクト改善のため):
'   ・画面別の表示回数(Hub/チャット/ナレッジ/ダッシュボード)
'   ・質問回数、モード別、解決/微妙/違うの比率
'   ・所属部署(config user_department)と利用者名
'
' 何を集めないか(ここが信頼の分かれ目):
'   ・質問の全文は共有側へ送らない。先頭40字までに切る。
'     傾向の把握には十分で、個別の中身を覗く用途には足りない粒度にしてある。
'   ・共有側に送っていることは「使い方」シートに明記する。黙って集めない。
'   ・ローカルの usage_log には従来どおり全文が残る(自分の履歴は自分のもの)。
'
' 送り方:
'   1人1ファイルを日次で上書きするだけ(telemetry\<user>_<yyyymm>.txt)。
'   追記ではなく上書きなので、何度書いても行が増えず、競合も起きない。
'   朝の一斉起動と重ならないよう、送信は終了時(Auto_Close)に行う。
'
' 匿名フィードバック:
'   既存のご意見箱はメール送信のため実名になる。率直な意見は実名では出にくい。
'   feedback\ に、利用者IDを含まないファイル名(乱数)で書き込む。
'   誰が書いたかは記録しないし、復元もできない。
' ============================================================================

Private Const TELE_SUBDIR As String = "telemetry"
Private Const FB_SUBDIR As String = "feedback"
Private Const Q_PREVIEW_CHARS As Long = 40

' ----------------------------------------------------------------------------
' TrackScreen - 画面表示を1件記録する(ローカルのmy_statsに貯めるだけ)。
'   共有フォルダI/Oはここでは行わない。画面遷移のたびにネットワークへ
'   書きに行くと、遷移が目に見えて遅くなる。
' ----------------------------------------------------------------------------
Public Sub TrackScreen(ByVal screenName As String)
    On Error Resume Next
    modStats.Bump "scr:" & screenName
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Publish - 集めた統計を共有フォルダへ1ファイル上書きする。
'   Auto_Close から呼ぶ(終了時なので数秒かかっても業務を止めない)。
' ----------------------------------------------------------------------------
Public Sub Publish()
    On Error Resume Next
    If Not modConfig.GetBool("telemetry_enabled", True) Then Exit Sub

    Dim dirPath As String: dirPath = SubDir(TELE_SUBDIR)
    If LenB(dirPath) = 0 Then Exit Sub
    EnsureDir dirPath

    Dim uid As String: uid = SafeUserId()
    If LenB(uid) = 0 Then Exit Sub

    Dim sb As String
    sb = "v1" & vbTab & uid & vbTab & AuthorName() & vbTab & DeptName() & vbTab & _
         Format$(Now, "yyyy-mm-dd hh:nn") & vbLf
    sb = sb & "ask_quick" & vbTab & modStats.GetStat("ask_quick_total") & vbLf
    sb = sb & "ask_deep" & vbTab & modStats.GetStat("ask_deep_total") & vbLf
    sb = sb & "solved" & vbTab & modStats.GetStat("selfsolve_total") & vbLf
    sb = sb & "unsure" & vbTab & modStats.GetStat("unsure_total") & vbLf
    sb = sb & "wrong" & vbTab & modStats.GetStat("fail_total") & vbLf
    sb = sb & "correction" & vbTab & modStats.GetStat("correction_total") & vbLf
    sb = sb & "saved_min" & vbTab & modStats.SavedMinutesEstimate() & vbLf
    sb = sb & "streak" & vbTab & modStats.GetStat("streak_days") & vbLf
    sb = sb & "level" & vbTab & modStats.Level() & vbLf
    sb = sb & "scr_hub" & vbTab & modStats.GetStat("scr:hub") & vbLf
    sb = sb & "scr_chat" & vbTab & modStats.GetStat("scr:chat") & vbLf
    sb = sb & "scr_knowledge" & vbTab & modStats.GetStat("scr:knowledge") & vbLf
    sb = sb & "scr_dash" & vbTab & modStats.GetStat("scr:dash") & vbLf

    ' 2026-07-28(レビュー M-19): ファイル名から月を外した。
    ' 送っているのは【生涯累計】のスナップショットなのに、ファイルを
    ' 「ユーザー×月」で分けていたため、2ヶ月目以降は同じ人の累計が
    ' 月数ぶん足し合わされ、利用者数・質問総数・節約時間が過大に出ていた
    ' (集計側は素直に全ファイルを足す)。1人1ファイルの上書きにする。
    Dim outPath As String
    outPath = dirPath & SafeName(uid) & ".txt"
    WriteShared outPath, sb

    ' 旧形式(<uid>_yyyymm.txt)が残っていると二重計上が続くので、
    ' 自分のぶんだけ片付ける。他人のファイルには触らない。
    CleanupLegacyMonthlyFiles dirPath, SafeName(uid)
    On Error GoTo 0
End Sub

' 旧「ユーザー×月」形式の自分のファイルを消す(移行用・レビュー M-19)。
' 列挙中に Kill すると列挙が壊れるので「集めてから消す」。
Private Sub CleanupLegacyMonthlyFiles(ByVal dirPath As String, ByVal uidSafe As String)
    On Error Resume Next
    Dim names() As String: ReDim names(0 To 63)
    Dim n As Long
    Dim fn As String: fn = Dir(dirPath & uidSafe & "_*.txt")
    Do While LenB(fn) > 0
        If n > UBound(names) Then ReDim Preserve names(0 To UBound(names) + 64)
        names(n) = fn
        n = n + 1
        fn = Dir()
    Loop
    Dim i As Long
    For i = 0 To n - 1
        Kill dirPath & names(i)
        Err.Clear
    Next i
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' SendAnonymousFeedback - 匿名の投書。利用者IDは一切書かない。
'   ファイル名も乱数なので、誰が書いたかは後から辿れない。
' ----------------------------------------------------------------------------
Public Function SendAnonymousFeedback(ByVal bodyText As String) As Boolean
    On Error Resume Next
    If LenB(Trim$(bodyText)) = 0 Then Exit Function

    Dim dirPath As String: dirPath = SubDir(FB_SUBDIR)
    If LenB(dirPath) = 0 Then Exit Function
    EnsureDir dirPath

    ' 本文と日付だけ。所属も名前も書かない(書けば匿名ではなくなる)。
    Dim sb As String
    sb = "v1" & vbTab & Format$(Now, "yyyy-mm-dd hh:nn") & vbTab & _
         Clean1(bodyText)

    Randomize
    Dim fname As String
    fname = "fb_" & Format$(Now, "yyyymmdd") & "_" & _
            Format$(Int(Rnd() * 1000000), "000000") & ".txt"
    SendAnonymousFeedback = WriteShared(dirPath & fname, sb)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' オーナー向け集計。共有フォルダのtelemetryを読んで要約テキストを返す。
' ----------------------------------------------------------------------------
Public Function SummaryText() As String
    On Error Resume Next
    Dim dirPath As String: dirPath = SubDir(TELE_SUBDIR)
    If LenB(dirPath) = 0 Then
        SummaryText = "共有フォルダが未設定です。"
        Exit Function
    End If
    If Len(Dir(dirPath, vbDirectory)) = 0 Then
        SummaryText = "まだ利用データが届いていません。"
        Exit Function
    End If

    Dim users As Long, askTotal As Long, solved As Long, savedMin As Long
    Dim wrong As Long, corr As Long
    Dim deptSummary As String

    Dim f As String: f = Dir(dirPath & "*.txt")
    Dim guard As Long
    Do While LenB(f) > 0 And guard < 500
        guard = guard + 1
        Dim txt As String
        If ReadShared(dirPath & f, txt) Then
            users = users + 1
            askTotal = askTotal + Val(FieldOf(txt, "ask_quick")) + Val(FieldOf(txt, "ask_deep"))
            solved = solved + Val(FieldOf(txt, "solved"))
            wrong = wrong + Val(FieldOf(txt, "wrong"))
            corr = corr + Val(FieldOf(txt, "correction"))
            savedMin = savedMin + Val(FieldOf(txt, "saved_min"))

            Dim d As String: d = HeaderDept(txt)
            If LenB(d) > 0 Then
                If InStr(1, deptSummary, "[" & d & "]") = 0 Then deptSummary = deptSummary & "[" & d & "]"
            End If
        End If
        f = Dir()
    Loop

    Dim sb As String
    sb = "■ 利用状況(共有フォルダに届いている分)" & vbLf & _
         "  利用者数: " & users & " 名" & vbLf & _
         "  質問の総数: " & askTotal & " 件" & vbLf & _
         "  解決した: " & solved & " 件" & vbLf & _
         "  違うと報告: " & wrong & " 件" & vbLf & _
         "  正しい内容の提供: " & corr & " 件" & vbLf & _
         "  取り戻した時間の合計: " & (savedMin \ 60) & " 時間 " & (savedMin Mod 60) & " 分" & vbLf & vbLf
    If askTotal > 0 Then
        sb = sb & "  解決率: " & CLng(solved * 100# / askTotal) & "%" & vbLf & vbLf
    End If
    sb = sb & "■ 参加している部署" & vbLf & "  " & Replace(deptSummary, "][", "] [") & vbLf
    SummaryText = sb
    On Error GoTo 0
End Function

' 匿名フィードバックの一覧(新しい順・最大20件)。
Public Function FeedbackText() As String
    On Error Resume Next
    Dim dirPath As String: dirPath = SubDir(FB_SUBDIR)
    If LenB(dirPath) = 0 Then Exit Function
    If Len(Dir(dirPath, vbDirectory)) = 0 Then
        FeedbackText = "まだ届いていません。"
        Exit Function
    End If

    Dim names() As String: ReDim names(0 To 255)
    Dim n As Long
    Dim f As String: f = Dir(dirPath & "fb_*.txt")
    Do While LenB(f) > 0 And n < 256
        names(n) = f
        n = n + 1
        f = Dir()
    Loop
    If n = 0 Then
        FeedbackText = "まだ届いていません。"
        Exit Function
    End If

    Dim a As Long, b As Long
    For a = 0 To n - 2
        For b = 0 To n - 2 - a
            If names(b) < names(b + 1) Then
                Dim t As String: t = names(b): names(b) = names(b + 1): names(b + 1) = t
            End If
        Next b
    Next a

    Dim sb As String
    Dim i As Long
    For i = 0 To n - 1
        If i >= 20 Then Exit For
        Dim txt As String
        If ReadShared(dirPath & names(i), txt) Then
            Dim parts() As String: parts = Split(txt, vbTab)
            If UBound(parts) >= 2 Then
                sb = sb & "・" & parts(1) & vbLf & "  " & parts(2) & vbLf & vbLf
            End If
        End If
    Next i
    FeedbackText = sb
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' 内部
' ----------------------------------------------------------------------------
Private Function FieldOf(ByVal txt As String, ByVal key As String) As String
    Dim lines_() As String: lines_ = Split(Replace(txt, vbCrLf, vbLf), vbLf)
    Dim i As Long
    For i = LBound(lines_) To UBound(lines_)
        If Left$(lines_(i), Len(key) + 1) = key & vbTab Then
            FieldOf = Mid$(lines_(i), Len(key) + 2)
            Exit Function
        End If
    Next i
End Function

Private Function HeaderDept(ByVal txt As String) As String
    Dim lines_() As String: lines_ = Split(Replace(txt, vbCrLf, vbLf), vbLf)
    If UBound(lines_) < 0 Then Exit Function
    Dim parts() As String: parts = Split(lines_(0), vbTab)
    If UBound(parts) >= 3 Then HeaderDept = Trim$(parts(3))
End Function

' 到達性の判定は modShare が1セッション1回だけ行う(レビュー M-20)。
' 終了時のテレメトリ送信がここで数十秒ブロックすると、
' 「閉じたのにExcelが残る」という最悪の別れ方になる。
Private Function SubDir(ByVal leaf As String) As String
    SubDir = modShare.SubDir(leaf)
End Function

Private Sub EnsureDir(ByVal folderPath As String)
    On Error Resume Next
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    On Error GoTo 0
End Sub

Private Function SafeUserId() As String
    On Error Resume Next
    SafeUserId = modP2P.CurrentUserId()
    On Error GoTo 0
End Function

Private Function AuthorName() As String
    On Error Resume Next
    AuthorName = Trim$(modConfig.GetString("pack_author", ""))
    On Error GoTo 0
End Function

Private Function DeptName() As String
    On Error Resume Next
    DeptName = Trim$(modConfig.GetString("user_department", ""))
    On Error GoTo 0
End Function

Private Function SafeName(ByVal s As String) As String
    Dim t As String: t = s
    Dim bad As Variant
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|", " ")
        t = Replace(t, CStr(bad), "_")
    Next bad
    SafeName = modUtil.SafeLeft(t, 32)
End Function

' 共有側へ出す文字列は改行とタブを潰し、長さも切る。
' 質問文の全文が共有フォルダに残らないようにするための最後の関門。
Private Function Clean1(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace(t, vbTab, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    Clean1 = modUtil.SafeLeft(Trim$(t), 1000)
End Function

' 質問文を共有用に切り詰める(先頭40字)。傾向把握には足り、
' 個別の中身を覗くには足りない粒度。
Public Function QuestionPreview(ByVal q As String) As String
    QuestionPreview = modUtil.SafeLeft(Trim$(q), Q_PREVIEW_CHARS)
End Function

Private Function WriteShared(ByVal filePath As String, ByVal content As String) As Boolean
    Dim attempt As Long
    For attempt = 1 To 3
        If TryWrite(filePath, content) Then
            WriteShared = True
            Exit Function
        End If
        Wait_ 200 * attempt
    Next attempt
End Function

Private Function TryWrite(ByVal filePath As String, ByVal content As String) As Boolean
    Dim st As Object
    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText content
    st.SaveToFile filePath, 2
    st.Close
    Set st = Nothing
    TryWrite = True
    Exit Function
Fail:
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup17
FailCleanup17:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Set st = Nothing
    On Error GoTo 0
End Function

Private Function ReadShared(ByVal filePath As String, ByRef outText As String) As Boolean
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
    ReadShared = True
    Exit Function
Fail:
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup18
FailCleanup18:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Set st = Nothing
    On Error GoTo 0
End Function

Private Sub Wait_(ByVal ms As Long)
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000# < ms
        DoEvents
        If Timer < t0 Then Exit Do
    Loop
End Sub
