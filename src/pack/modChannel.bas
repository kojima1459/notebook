Attribute VB_Name = "modChannel"
Option Explicit

' ============================================================================
' modChannel - 部門チャンネル(正典ナレッジの発行と購読)
' ----------------------------------------------------------------------------
' 解決する構造問題:
'   これまでの設計は「12,000人がそれぞれ自分の本棚を組み立てる」だった。
'   これは3つの意味で破綻する。
'     ・同じ資料が12,000回ベクトル化される(費用も時間も無駄)
'     ・人によって本棚が違うので、同じ質問に違う答えが出る(金融では致命的)
'     ・「何を入れたらいいか分からない」は個人の怠慢ではなく、個人に判断させる
'       設計そのものの欠陥
'   そこで配布の単位を個人から部門へ移す。各部門が自分の正典を1つ発行し、
'   利用者は必要なチャンネルを購読するだけでよい。開いた初日から答えが返る。
'
' 併存の原則(重要):
'   チャンネルはマイ本棚を置き換えない。足すだけである。
'     ・チャンネル = 部門が保守する共通の正典(商品/システム/人事/法務…)
'     ・マイ本棚   = 自分の担当領域・自分だけが持つ資料・個人メモ
'   検索は両方を等しく対象にする(my_knowledge に同居させるだけなので、
'   検索・回答・出典の仕組みには一切手を入れていない)。
'   ポータルを探し回る手間が消えるのは前者、自分固有の資料が効くのは後者で、
'   どちらが欠けても実務のペインは消えない。
'
' 2万チャンク上限をどう越えるか(部門が5つ以上ある前提での設計):
'   1) 購読制: 全社共通と自部門は既定で購読、他部門は必要な人だけ。
'      「全部入れる」を既定にしない。
'   2) 正典は原文全部ではなく「確認済みQ&A+要点」に絞る運用を前提にする。
'      100ページの約款をそのまま入れれば1部門で数千チャンク食うが、
'      正典Q&Aなら1件1〜2チャンクで済む。原文が要る人は自分の本棚へ入れる。
'   3) 予算の可視化: modChannel.ChunkUsagePercent がHubに使用率を出す。
'      8割を超えたら警告し、出所別の内訳から丸ごと外せるようにする。
'   4) 版管理: 同じチャンネルの新版が出たら、旧版のチャンクを消してから
'      入れ直す(差分ではなく置換)。放置すると版が重なって際限なく増える。
'
' 共有フォルダの構造:
'   <nexus_share_path>\channels\<チャンネル名>\
'        pack.xlsx        … 正典パック本体(modPack形式)
'        version.txt      … 版番号+発行日+発行者(1行)
'        _archive\        … 過去版(誤削除・誤発行からの復旧用)
'
' 事故対策(共有フォルダが消えた/壊れた場合):
'   ・発行時に必ず _archive\pack_<版>.xlsx を残す(世代バックアップ)
'   ・購読者のローカル本棚に取り込み済みの内容は消えない。つまり配布済みの
'     正典は各PCに自然に複製されており、共有フォルダ消失=全損にはならない
'   ・発行者は自分の本棚から再発行できる(pack.xlsxは本棚から再生成できる)
'   ・書き換えリスクは「共有フォルダ側の書き込み権限を発行者に限る」ことで
'     運用でヘッジする(このアプリからは正典を上書きできない設計にしてある。
'     PublishChannel は必ず新しい版番号で書き、既存版は消さない)。
' ============================================================================

Private Const CH_SUBDIR As String = "channels"
Private Const PACK_NAME As String = "pack.xlsx"
Private Const VER_NAME As String = "version.txt"
Private Const ARCHIVE_DIR As String = "_archive"
Private Const MAX_CHANNELS As Long = 40

' ----------------------------------------------------------------------------
' ListChannels - 共有フォルダにある部門チャンネル名を | 区切りで返す。
' ----------------------------------------------------------------------------
Public Function ListChannels() As String
    On Error Resume Next
    Dim baseDir As String: baseDir = ChannelsDir()
    If LenB(baseDir) = 0 Then Exit Function
    If Len(Dir(baseDir, vbDirectory)) = 0 Then Exit Function

    Dim sb As String, n As Long
    Dim nm As String
    nm = Dir(baseDir & "*", vbDirectory)
    Do While LenB(nm) > 0 And n < MAX_CHANNELS
        If nm <> "." And nm <> ".." And Left$(nm, 1) <> "_" Then
            ' フォルダかどうかは version.txt の有無で判定する(GetAttrは
            ' ネットワークドライブで例外を投げることがあるため使わない)。
            If LenB(Dir(baseDir & nm & "\" & VER_NAME)) > 0 Then
                If LenB(sb) > 0 Then sb = sb & "|"
                sb = sb & nm
                n = n + 1
            End If
        End If
        nm = Dir()
    Loop
    ListChannels = sb
    On Error GoTo 0
End Function

' 購読中か(config subscribed_channels のカンマ区切りに載っているか)。
Public Function IsSubscribed(ByVal chName As String) As Boolean
    On Error Resume Next
    Dim subs As String: subs = "," & modConfig.GetString("subscribed_channels", "") & ","
    IsSubscribed = (InStr(1, subs, "," & chName & ",", vbTextCompare) > 0)
    On Error GoTo 0
End Function

Public Sub Subscribe(ByVal chName As String)
    On Error Resume Next
    If IsSubscribed(chName) Then Exit Sub
    Dim subs As String: subs = Trim$(modConfig.GetString("subscribed_channels", ""))
    If LenB(subs) > 0 Then subs = subs & ","
    modConfig.SetValue "subscribed_channels", subs & chName
    On Error GoTo 0
End Sub

Public Sub Unsubscribe(ByVal chName As String)
    On Error Resume Next
    Dim subs As String: subs = modConfig.GetString("subscribed_channels", "")
    Dim parts() As String: parts = Split(subs, ",")
    Dim sb As String
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim one As String: one = Trim$(parts(i))
        If LenB(one) > 0 And StrComp(one, chName, vbTextCompare) <> 0 Then
            If LenB(sb) > 0 Then sb = sb & ","
            sb = sb & one
        End If
    Next i
    modConfig.SetValue "subscribed_channels", sb
    On Error GoTo 0
End Sub

' チャンネルの現在の版(共有フォルダ側)。"版番号|発行日|発行者" 形式。
Public Function RemoteVersion(ByVal chName As String) As String
    On Error Resume Next
    Dim p As String: p = ChannelsDir() & chName & "\" & VER_NAME
    If LenB(p) = 0 Then Exit Function
    Dim txt As String
    If ReadShared(p, txt) Then RemoteVersion = Trim$(Replace(Replace(txt, vbCr, ""), vbLf, ""))
    On Error GoTo 0
End Function

' 自分が取り込み済みの版(my_stats に記録)。
Public Function LocalVersion(ByVal chName As String) As String
    On Error Resume Next
    LocalVersion = modStats.GetStatText("ch:" & LCase$(chName))
    On Error GoTo 0
End Function

' 更新があるチャンネル名を | 区切りで返す(購読中のものだけ)。
Public Function PendingUpdates() As String
    On Error Resume Next
    Dim all As String: all = ListChannels()
    If LenB(all) = 0 Then Exit Function

    Dim parts() As String: parts = Split(all, "|")
    Dim sb As String
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If IsSubscribed(parts(i)) Then
            If RemoteVersion(parts(i)) <> LocalVersion(parts(i)) Then
                If LenB(RemoteVersion(parts(i))) > 0 Then
                    If LenB(sb) > 0 Then sb = sb & "|"
                    sb = sb & parts(i)
                End If
            End If
        End If
    Next i
    PendingUpdates = sb
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' SyncChannel - 1チャンネルを最新版へ更新する。
'   版が変わっているときは、まず旧版で取り込んだチャンクを削除してから
'   入れ直す(差分ではなく置換)。これをやらないと改定のたびに古い記述が
'   残り、AIが古い条文を根拠に答えるという最悪の事故になる。
'   戻り値 = 取り込んだチャンク数(0=更新なし/失敗)。
' ----------------------------------------------------------------------------
Public Function SyncChannel(ByVal chName As String) As Long
    On Error Resume Next
    Dim remote As String: remote = RemoteVersion(chName)
    If LenB(remote) = 0 Then Exit Function
    If remote = LocalVersion(chName) Then Exit Function

    Dim packPath As String: packPath = ChannelsDir() & chName & "\" & PACK_NAME
    If LenB(Dir(packPath)) = 0 Then Exit Function

    ' 旧版の掃除。origin が "pack:<チャンネル名>" のチャンクをまとめて消す。
    PurgeChannelChunks chName

    Dim got As Long
    got = modPack.ImportPackFile(packPath, True)
    If got > 0 Then
        modStats.SetStatText "ch:" & LCase$(chName), remote
        modLog.LogUsage "channel_sync", chName, "version=" & remote & " chunks=" & got
    End If
    SyncChannel = got
    On Error GoTo 0
End Function

' 購読中の全チャンネルを同期する(利用者が明示的に押したときだけ呼ぶ)。
Public Function SyncSubscribed() As Long
    On Error Resume Next
    Dim pend As String: pend = PendingUpdates()
    If LenB(pend) = 0 Then Exit Function
    Dim parts() As String: parts = Split(pend, "|")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        SyncSubscribed = SyncSubscribed + SyncChannel(parts(i))
    Next i
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' チャンク予算。上限に当たってから慌てないよう、常に使用率を見せる。
' ----------------------------------------------------------------------------
Public Function ChunkLimit() As Long
    On Error Resume Next
    ChunkLimit = modConfig.GetLong("chunk_limit", 20000)
    On Error GoTo 0
    If ChunkLimit < 1000 Then ChunkLimit = 20000
End Function

Public Function ChunkUsagePercent() As Long
    Dim used As Long
    On Error Resume Next
    used = modShelf.TotalChunks()
    On Error GoTo 0
    If ChunkLimit() < 1 Then Exit Function
    ChunkUsagePercent = CLng(used * 100# / ChunkLimit())
End Function

' 上限に近いか(8割超)。Hubの警告表示とツールバーの棚卸し導線に使う。
Public Function IsBudgetTight() As Boolean
    IsBudgetTight = (ChunkUsagePercent() >= 80)
End Function

' ----------------------------------------------------------------------------
' PurgeChannelChunks - あるチャンネル由来のチャンクを本棚から取り除く。
'   購読解除と版の入れ替えの両方で使う。my_knowledge と my_vectors の
'   両方を同じ判定で消さないとベクトルだけが残って検索が壊れる。
' ----------------------------------------------------------------------------
Public Function PurgeChannelChunks(ByVal chName As String) As Long
    On Error Resume Next
    Dim wsK As Worksheet, wsV As Worksheet
    Set wsK = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    Set wsV = ThisWorkbook.Worksheets(modAppDef.SH_VECTORS)
    If wsK Is Nothing Then Exit Function

    Dim tag As String: tag = "pack:" & chName
    Dim lastR As Long: lastR = wsK.Cells(wsK.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Function

    ' 消すchunk_idを先に集める(行を消しながら走査すると添字がずれる)。
    Dim ids() As String: ReDim ids(0 To lastR)
    Dim n As Long
    Dim r As Long
    For r = 2 To lastR
        If StrComp(Trim$(CStr(wsK.Cells(r, 3).Value)), tag, vbTextCompare) = 0 Then
            ids(n) = CStr(wsK.Cells(r, 1).Value)
            n = n + 1
        End If
    Next r
    If n = 0 Then Exit Function

    ' 後ろから消す。
    For r = lastR To 2 Step -1
        If StrComp(Trim$(CStr(wsK.Cells(r, 3).Value)), tag, vbTextCompare) = 0 Then
            wsK.Rows(r).Delete
        End If
    Next r

    If Not wsV Is Nothing Then
        Dim lastV As Long: lastV = wsV.Cells(wsV.Rows.Count, 1).End(xlUp).Row
        Dim v As Long, i As Long
        For v = lastV To 2 Step -1
            Dim cid As String: cid = CStr(wsV.Cells(v, 1).Value)
            For i = 0 To n - 1
                If ids(i) = cid Then
                    wsV.Rows(v).Delete
                    Exit For
                End If
            Next i
        Next v
    End If

    modLog.LogUsage "channel_purge", chName, "removed=" & n
    PurgeChannelChunks = n
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' PublishChannel - 自分の本棚の内容を、部門の正典として発行する(発行者用)。
'   既存の pack.xlsx は消さず、_archive へ版を残してから置き換える。
'   誤発行しても過去版から戻せるようにするための世代管理。
'   実際のパック生成は modPack のダイアログ経路を使う(PII走査を必ず通す
'   ため。無言でPII入りの資料を全社配布する経路は作らない)。
' ----------------------------------------------------------------------------
Public Function PrepareChannelDir(ByVal chName As String) As String
    On Error Resume Next
    Dim d As String: d = ChannelsDir()
    If LenB(d) = 0 Then Exit Function
    EnsureDir d
    EnsureDir d & chName & "\"
    EnsureDir d & chName & "\" & ARCHIVE_DIR & "\"
    PrepareChannelDir = d & chName & "\"
    On Error GoTo 0
End Function

' 新しい版番号を書き込む(発行の最後に呼ぶ)。既存版は _archive に退避済み前提。
Public Sub WriteVersion(ByVal chName As String, ByVal author As String)
    On Error Resume Next
    Dim d As String: d = PrepareChannelDir(chName)
    If LenB(d) = 0 Then Exit Sub
    Dim ver As String
    ver = Format$(Now, "yyyymmdd-hhnn") & "|" & Format$(Date, "yyyy-mm-dd") & "|" & author
    WriteShared d & VER_NAME, ver
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 内部
' ----------------------------------------------------------------------------
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
    WriteShared = True
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
