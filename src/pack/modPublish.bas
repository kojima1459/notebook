Attribute VB_Name = "modPublish"
Option Explicit

' ============================================================================
' modPublish - 部門正典の発行(手を動かすのは役職なしの若手という前提)
' ----------------------------------------------------------------------------
' 設計思想:
'   発行者は専門家ではない。ミスは必ず起きる前提で作る。
'   目指すのは「間違えられない設計」ではなく「間違えても数十秒で戻せる設計」。
'   承認フローを挟むと、事故が起きたときに止められなくなる(承認待ちの間、
'   誤った正典が全社に配信され続ける)。だから承認は入れず、
'   代わりに【誰でも即座に巻き戻せる】ことに全振りしている。
'
' 発行の手順(5段):
'   1. チャンネル選択 … 部門名そのまま。実在する組織と1対1にする
'   2. 内容確認     … 何件出るのかを発行前に見せる
'   3. PII走査      … modPackの既存走査を必ず通す(個人情報があれば中止)
'   4. 発行キー確認 … config publish_key と一致しなければ発行しない
'   5. 発行         … 旧版を_archiveへ退避 → pack.xlsx → 最後にversion.txt
'
' なぜ version.txt を最後に書くか(重要):
'   購読者は version.txt だけを見て「新しい版が来た」と判断する。
'   pack.xlsx を書き終える前に version.txt を更新すると、
'   「新しい版番号なのに中身が古い/書きかけ」を掴む事故が起きる。
'   最後に書けば、この不整合は原理的に発生しない。
'
' 巻き戻し:
'   Rollback は _archive の直前版を pack.xlsx へ戻し、version.txt も
'   その版へ戻す。購読者は次回起動時に「版が変わった」と判定して
'   旧版へ置換する(誤った内容は各PCから消える)。
'
' 発行ログ:
'   channels\<名前>\publish_log.txt に日時・発行者・件数・版を追記。
'   誰がいつ何を出したかが常に追え、責任の所在が曖昧にならない。
' ============================================================================

Private Const CH_SUBDIR As String = "channels"
Private Const PACK_NAME As String = "pack.xlsx"
Private Const VER_NAME As String = "version.txt"
Private Const LOG_NAME As String = "publish_log.txt"
Private Const ARCHIVE_DIR As String = "_archive"

' ----------------------------------------------------------------------------
' CanPublish - 発行キーが設定されているか(発行画面を出してよいか)。
'   キーが空の配布物では発行機能そのものを見せない。一般利用者の画面に
'   「押してはいけないボタン」を置かないため。
' ----------------------------------------------------------------------------
Public Function CanPublish() As Boolean
    On Error Resume Next
    CanPublish = (LenB(Trim$(modConfig.GetString("publish_key", ""))) > 0)
    On Error GoTo 0
End Function

' 入力されたキーが正しいか。
Public Function VerifyKey(ByVal entered As String) As Boolean
    On Error Resume Next
    Dim k As String: k = Trim$(modConfig.GetString("publish_key", ""))
    If LenB(k) = 0 Then Exit Function
    VerifyKey = (StrComp(Trim$(entered), k, vbBinaryCompare) = 0)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' PrepareDir - チャンネル用フォルダ一式を用意し、そのパスを返す。
' ----------------------------------------------------------------------------
Public Function PrepareDir(ByVal chName As String) As String
    On Error Resume Next
    Dim baseDir As String: baseDir = ChannelsDir()
    If LenB(baseDir) = 0 Then Exit Function
    EnsureDir baseDir
    EnsureDir baseDir & chName & "\"
    EnsureDir baseDir & chName & "\" & ARCHIVE_DIR & "\"
    PrepareDir = baseDir & chName & "\"
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' ArchiveCurrent - 今ある pack.xlsx を _archive へ退避する。
'   戻り値 = 退避したファイル名(無ければ空)。発行の前に必ず呼ぶ。
' ----------------------------------------------------------------------------
Public Function ArchiveCurrent(ByVal chName As String) As String
    On Error Resume Next
    Dim d As String: d = PrepareDir(chName)
    If LenB(d) = 0 Then Exit Function
    If LenB(Dir(d & PACK_NAME)) = 0 Then Exit Function

    Dim stamp As String: stamp = Format$(Now, "yyyymmdd-hhnnss")
    Dim dest As String: dest = d & ARCHIVE_DIR & "\pack_" & stamp & ".xlsx"
    FileCopy d & PACK_NAME, dest
    If LenB(Dir(dest)) > 0 Then ArchiveCurrent = "pack_" & stamp & ".xlsx"

    ' version.txt も同じ名前で残す(巻き戻しで版番号ごと戻すため)。
    If LenB(Dir(d & VER_NAME)) > 0 Then
        FileCopy d & VER_NAME, d & ARCHIVE_DIR & "\ver_" & stamp & ".txt"
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' FinalizePublish - pack.xlsx を置いた後に呼ぶ。version.txt を最後に更新し、
'   発行ログを追記する。ここを通って初めて購読者に配信される。
' ----------------------------------------------------------------------------
Public Function FinalizePublish(ByVal chName As String, ByVal chunkCount As Long) As String
    On Error Resume Next
    Dim d As String: d = PrepareDir(chName)
    If LenB(d) = 0 Then Exit Function

    ' pack.xlsx が実在しない状態で version だけ進めない(不整合の予防)。
    If LenB(Dir(d & PACK_NAME)) = 0 Then Exit Function

    Dim author As String: author = PublisherName()
    Dim ver As String
    ver = Format$(Now, "yyyymmdd-hhnn") & "|" & Format$(Date, "yyyy-mm-dd") & "|" & author

    If Not WriteShared(d & VER_NAME, ver) Then Exit Function

    AppendLog chName, "発行" & vbTab & ver & vbTab & chunkCount & "件"
    modLog.LogUsage "channel_publish", chName, "version=" & ver & " chunks=" & chunkCount
    FinalizePublish = ver
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' ArchiveList - _archive にある過去版を新しい順に | 区切りで返す(最大10件)。
' ----------------------------------------------------------------------------
Public Function ArchiveList(ByVal chName As String) As String
    On Error Resume Next
    Dim d As String: d = ChannelsDir() & chName & "\" & ARCHIVE_DIR & "\"
    If Len(Dir(d, vbDirectory)) = 0 Then Exit Function

    Dim names() As String: ReDim names(0 To 63)
    Dim n As Long
    Dim f As String: f = Dir(d & "pack_*.xlsx")
    Do While LenB(f) > 0 And n < 64
        names(n) = f
        n = n + 1
        f = Dir()
    Loop
    If n = 0 Then Exit Function

    ' ファイル名に時刻が入っているので、名前の降順=新しい順。
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
        If i >= 10 Then Exit For
        If LenB(sb) > 0 Then sb = sb & "|"
        sb = sb & names(i)
    Next i
    ArchiveList = sb
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' Rollback - 過去版を現行に戻す。誤った正典の配信を止める最終手段。
'   archiveName は ArchiveList が返すファイル名(pack_yyyymmdd-hhnnss.xlsx)。
'   購読者は次回起動時に版の変化を検知し、旧版へ置き換える。
' ----------------------------------------------------------------------------
Public Function Rollback(ByVal chName As String, ByVal archiveName As String) As Boolean
    On Error Resume Next
    Dim d As String: d = PrepareDir(chName)
    If LenB(d) = 0 Then Exit Function
    Dim src As String: src = d & ARCHIVE_DIR & "\" & archiveName
    If LenB(Dir(src)) = 0 Then Exit Function

    ' 巻き戻す前に、今の版も退避しておく(巻き戻し自体を取り消せるように)。
    ArchiveCurrent chName

    FileCopy src, d & PACK_NAME
    If Err.Number <> 0 Then Exit Function

    ' 版番号は必ず「新しい値」にする。古い版番号に戻すと、既にその版を
    ' 取り込み済みの人が「変化なし」と判定して巻き戻しが届かない。
    Dim ver As String
    ver = Format$(Now, "yyyymmdd-hhnn") & "|" & Format$(Date, "yyyy-mm-dd") & "|" & _
          PublisherName() & "(巻き戻し:" & archiveName & ")"
    If Not WriteShared(d & VER_NAME, ver) Then Exit Function

    AppendLog chName, "巻き戻し" & vbTab & ver & vbTab & archiveName
    modLog.LogUsage "channel_rollback", chName, archiveName
    Rollback = True
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' PackDestPath - 発行先の pack.xlsx フルパス(UI層が書き出し先に使う)。
' ----------------------------------------------------------------------------
Public Function PackDestPath(ByVal chName As String) As String
    On Error Resume Next
    Dim d As String: d = PrepareDir(chName)
    If LenB(d) = 0 Then Exit Function
    PackDestPath = d & PACK_NAME
    On Error GoTo 0
End Function

' 直近の発行ログ(新しい順に最大10行)。発行画面に出して履歴を見せる。
Public Function RecentLog(ByVal chName As String) As String
    On Error Resume Next
    Dim p As String: p = ChannelsDir() & chName & "\" & LOG_NAME
    If LenB(Dir(p)) = 0 Then Exit Function
    Dim txt As String
    If Not ReadShared(p, txt) Then Exit Function

    Dim lines_() As String: lines_ = Split(Replace(txt, vbCrLf, vbLf), vbLf)
    Dim sb As String, cnt As Long
    Dim i As Long
    For i = UBound(lines_) To LBound(lines_) Step -1
        If LenB(Trim$(lines_(i))) > 0 Then
            sb = sb & lines_(i) & vbLf
            cnt = cnt + 1
            If cnt >= 10 Then Exit For
        End If
    Next i
    RecentLog = sb
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' 内部
' ----------------------------------------------------------------------------
Private Sub AppendLog(ByVal chName As String, ByVal line_ As String)
    On Error Resume Next
    Dim p As String: p = PrepareDir(chName) & LOG_NAME
    If LenB(p) = 0 Then Exit Sub

    Dim prev As String
    ReadShared p, prev
    Dim body As String
    body = prev & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & PublisherName() & _
           vbTab & line_ & vbLf
    WriteShared p, body
    On Error GoTo 0
End Sub

Private Function PublisherName() As String
    On Error Resume Next
    PublisherName = Trim$(modConfig.GetString("pack_author", ""))
    If LenB(PublisherName) = 0 Or PublisherName = "名称未設定" Then
        PublisherName = modP2P.CurrentUserId()
    End If
    On Error GoTo 0
End Function

Private Function ChannelsDir() As String
    Dim basePath As String
    On Error Resume Next
    basePath = modConfig.GetString("nexus_share_path", "")
    On Error GoTo 0
    If LenB(basePath) = 0 Then Exit Function
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
    ChannelsDir = basePath & CH_SUBDIR & "\"
End Function

Private Sub EnsureDir(ByVal folderPath As String)
    On Error Resume Next
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    On Error GoTo 0
End Sub

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
