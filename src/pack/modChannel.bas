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
'   1) 既定で全チャンネルを購読する(2026-07-26 設計変更)。
'      「必要なものだけ購読」は誤りだった。利用者の行動は
'      「今日は商品、明日はシステム、明後日は経費精算」であり、聞く前に
'      分野を判断して購読操作をさせた時点で、ポータルを探し回るのと同じ
'      認知負荷が発生する。解決したかった問題を作り直してしまう。
'      成立の根拠: 正典はQ&A粒度(1件1～2チャンク)なので、6部門×500件でも
'      3,000～6,000チャンク。個人の本棚と合わせても上限に届かない。
'      ただし「正典に原文マニュアルを丸ごと入れない」ことが絶対条件。
'   2) 正典は原文全部ではなく「確認済みQ&A+要点」に絞る運用を前提にする。
'      100ページの約款をそのまま入れれば1部門で数千チャンク食うが、
'      正典Q&Aなら1件1～2チャンクで済む。原文が要る人は自分の本棚へ入れる。
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
Private Const MIGRATION_KEY As String = "origin_ns_migrated"
Private Const PACK_NAME As String = "pack.xlsx"
Private Const VER_NAME As String = "version.txt"
Private Const ARCHIVE_DIR As String = "_archive"
Private Const MAX_CHANNELS As Long = 40

' 直近の ListChannels で上限を超えて読み込めなかった部門数(R8 F13)。
' Hub の受信箱がこれを見て注意を出す。VBAはモジュールレベル宣言を
' プロシージャより前に置く必要がある。
Private mLastOverflow As Long

' ----------------------------------------------------------------------------
' ListChannels - 共有フォルダにある部門チャンネル名を | 区切りで返す。
' ----------------------------------------------------------------------------
Public Function ListChannels() As String
    On Error Resume Next
    Dim baseDir As String: baseDir = ChannelsDir()
    If LenB(baseDir) = 0 Then Exit Function
    If Len(Dir(baseDir, vbDirectory)) = 0 Then Exit Function

    ' 【重要】Dir()は「1つの列挙状態」しか持たない。フォルダ列挙のループの
    ' 中で別の Dir(パス) を呼ぶと、その瞬間に列挙が乗っ取られ、続く Dir() は
    ' 内側のパターンを返す。結果、チャンネルが1つしか見つからない/無限ループ
    ' になる。必ず「先に名前を全部集める → あとで中身を確認する」の2段にする。
    Dim names() As String
    ReDim names(0 To MAX_CHANNELS)
    Dim n As Long
    Dim overflow As Long
    Dim nm As String
    nm = Dir(baseDir & "*", vbDirectory)
    Do While LenB(nm) > 0
        If nm <> "." And nm <> ".." And Left$(nm, 1) <> "_" Then
            If n < MAX_CHANNELS Then
                names(n) = nm
                n = n + 1
            Else
                ' 2026-07-31(レビュー R8 F13): 41件目以降は静かに捨てられて
                ' いた。上限に当たったこと自体が誰にも見えないため、
                ' 「うちの部門だけ正典が届かない」という調べようのない
                ' 不具合になる。落とした件数を数えて必ず記録する。
                overflow = overflow + 1
            End If
        End If
        nm = Dir()
    Loop
    mLastOverflow = overflow
    If overflow > 0 Then
        modLog.LogError "E0806", "modChannel.ListChannels", _
            "部門チャンネルが上限(" & MAX_CHANNELS & ")を超えました。" & _
            overflow & "部門を読み込み対象から外しています。MAX_CHANNELS の引き上げが必要です。"
    End If

    ' ここから先は列挙が終わっているので Dir() を自由に使える。
    ' フォルダかどうかは version.txt の有無で判定する(GetAttrは
    ' ネットワークドライブで例外を投げることがあるため使わない)。
    Dim sb As String
    Dim i As Long
    For i = 0 To n - 1
        If LenB(Dir(baseDir & names(i) & "\" & VER_NAME)) > 0 Then
            If LenB(sb) > 0 Then sb = sb & "|"
            sb = sb & names(i)
        End If
    Next i
    ListChannels = sb
    On Error GoTo 0
End Function

' OverflowCount - 直近の ListChannels で上限超過により切り捨てた部門数(R8 F13)。
' 0 なら正常。Hub の受信箱がこれを見て注意表示を出す。
Public Function OverflowCount() As Long
    OverflowCount = mLastOverflow
End Function

' 購読中か。既定は「全部購読」なので、判定は除外リスト(config
' unsubscribed_channels)に載っていないこと。新しい部門が正典を発行したら
' 誰も操作しなくても自動的に届く。これが v3 の要。
' 2026-08-16(R33H F4): 区切りはカンマから "|" へ。Windows のフォルダ名に
' カンマは使えるため、「品質保証,監査」を解除すると「品質保証」と「監査」の
' 両方が届かなくなっていた。読み書きと旧形式からの移行は modShareRule 側
' (純ロジック=modTestsPure35 が境界を固定する)。
Public Function IsSubscribed(ByVal chName As String) As Boolean
    On Error Resume Next
    IsSubscribed = Not modShareRule.UnsubListHas( _
        modConfig.GetString("unsubscribed_channels", ""), chName)
    On Error GoTo 0
End Function

' 購読を戻す(除外リストから外す)。
' 2026-07-28(レビュー I-14/I-4): 現在この2本を呼ぶUIは無い。ただし
' unsubscribed_channels 自体は IsSubscribed → PendingUpdates 経由で
' Hub の更新通知に効いており「死に設定」ではない。設定UIが無いだけなので、
' config を手編集する運用のための対として残す(消すと設定を安全に
' 書き換える手段が無くなる)。
Public Sub Subscribe(ByVal chName As String)
    On Error Resume Next
    modConfig.SetValue "unsubscribed_channels", modShareRule.UnsubListRemove( _
        modConfig.GetString("unsubscribed_channels", ""), chName)
    On Error GoTo 0
End Sub

' 購読をやめる(除外リストへ追加)。戻り値=以後この部門が届かない状態か。
' 2026-08-16(R33H F3): Sub のままだと modConfig.SetValue の失敗(読み取り専用・
' シート保護)が呼び出し元へ伝わらず、無条件に「今後届きません」と断言して
' usage_log にも unsubscribed=yes を残していた。書いたあとに IsSubscribed を
' 読み直して検算した結果を返す(W3-8 KillGsTree と同型の手術)。
Public Function Unsubscribe(ByVal chName As String) As Boolean
    On Error Resume Next
    If LenB(Trim$(chName)) = 0 Then Exit Function
    If Not IsSubscribed(chName) Then
        Unsubscribe = True          ' 既に届かない=求められた状態には既にある
        Exit Function
    End If
    modConfig.SetValue "unsubscribed_channels", modShareRule.UnsubListAdd( _
        modConfig.GetString("unsubscribed_channels", ""), chName)
    Unsubscribe = Not IsSubscribed(chName)
    On Error GoTo 0
End Function

' チャンネルの現在の版(共有フォルダ側)。"版番号|発行日|発行者" 形式。
Public Function RemoteVersion(ByVal chName As String) As String
    On Error Resume Next
    ' 2026-07-31(レビュー R8・低): 判定は【連結する前】に行う。
    ' 連結後の LenB(p) は "\version.txt" が付いている分、共有が未設定でも
    ' 必ず 0 より大きく、ガードとして一度も働かなかった(潜在欠陥の芽)。
    Dim baseDir As String: baseDir = ChannelsDir()
    If LenB(baseDir) = 0 Then Exit Function
    Dim p As String: p = baseDir & chName & "\" & VER_NAME
    Dim txt As String
    If ReadShared(p, txt) Then RemoteVersion = Trim$(Replace(Replace(txt, vbCr, ""), vbLf, ""))
    On Error GoTo 0
End Function

' 自分が取り込み済みの版(my_stats に記録)。
Public Function LocalVersion(ByVal chName As String) As String
    On Error Resume Next
    LocalVersion = modStats.GetStatText(modShareRule.ChannelStatKey(chName))
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
            ' 2026-07-31(レビュー R8 F3): RemoteVersion は共有フォルダの
            ' version.txt を ADODB.Stream で開いて読む【実I/O】。同じ部門に
            ' つき2回呼んでいたため、部門数×2回の共有アクセスを毎回払っていた。
            ' 変数へ1回だけ受ける(結果は同じ)。
            Dim remote As String: remote = RemoteVersion(parts(i))
            If LenB(remote) > 0 And remote <> LocalVersion(parts(i)) Then
                If LenB(sb) > 0 Then sb = sb & "|"
                sb = sb & parts(i)
            End If
        End If
    Next i
    PendingUpdates = sb
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' SyncChannel - 1チャンネルを最新版へ更新する。
'   版が変わっているときは、旧版で取り込んだチャンクを削除してから入れ直す
'   (差分ではなく置換)。これをやらないと改定のたびに古い記述が残り、
'   AIが古い条文を根拠に答えるという最悪の事故になる。
'   戻り値 = 取り込んだチャンク数(0=更新なし/失敗、-1は使わない)。
'   2026-07-31(R11-H Med5): 戻り値0が「更新が無かった(成功)」と
'   「取り込めなかった(失敗)」の両方を意味していたため、呼び出し側は
'   全部まとめて失敗として数えるしかなかった(実機で「更新はありませんでした
'   (1部門は取り込めませんでした)」という自己矛盾した通知が出る)。
'   outOk に「処理として成功したか」を、outErrDesc に失敗の理由を返す。
'
' leavingChannel (2026-07-28 レビュー H-14):
'   部門を切り替えるときに「離れる部門」を渡す。旧部門の削除は
'   SwitchTo ではなくここで行う。SwitchTo が先に消していた頃は、
'   共有瞬断や発行者の上書きロックで pack のコピーに失敗すると
'   「旧部門は消えた・新部門は入らない・active_channel は空」という
'   三重苦で終わっていた。破壊的な操作は、失敗しうる I/O を全部
'   終えてから始める。
' ----------------------------------------------------------------------------
Public Function SyncChannel(ByVal chName As String, Optional ByVal leavingChannel As String = "", _
                           Optional ByRef outOk As Boolean, _
                           Optional ByRef outErrDesc As String) As Long
    On Error Resume Next
    outOk = False
    outErrDesc = ""
    Dim remote As String: remote = RemoteVersion(chName)
    If LenB(remote) = 0 Then
        outErrDesc = "版番号(version.txt)を読めませんでした"
        Exit Function
    End If
    ' 切替時(leavingChannel あり)は「版が同じ」でも処理を通す。旧部門を
    ' 消して新部門を入れ直す必要があるため。
    If LenB(leavingChannel) = 0 Then
        If remote = LocalVersion(chName) Then
            outOk = True        ' 更新なし=成功で0件(失敗ではない)
            Exit Function
        End If
    End If

    ' 連結前にガードする(R8・低)。共有が未設定/到達不能なら ChannelsDir は
    ' 空文字を返すので、そのまま連結すると相対パスを掴みに行くことになる。
    Dim baseDir As String: baseDir = ChannelsDir()
    If LenB(baseDir) = 0 Then
        outErrDesc = "共有フォルダが未設定または到達できません"
        Exit Function
    End If
    Dim packPath As String: packPath = baseDir & chName & "\" & PACK_NAME
    If LenB(Dir(packPath)) = 0 Then
        outErrDesc = "部門のパックファイルが見つかりません"
        Exit Function
    End If

    ' 朝の一斉起動で数千人が同じファイルを掴みに行くため、共有側に読み取り
    ' ロックを残さないよう %TEMP% へコピーしてから開く。発行者が書き換えても
    ' こちらが読むのはコピーなので壊れたファイルを掴まない。
    Dim localPath As String
    localPath = CopyToTemp(packPath, chName)
    If LenB(localPath) = 0 Then
        outErrDesc = "パックを一時フォルダへコピーできませんでした"
        Exit Function
    End If

    ' 消す対象のタグを組み立てる。実際の削除は ImportPackFile の中、
    ' 「パックを読み終えて書き込む直前」に行われる(同居時間ゼロのまま、
    ' 読み込み失敗で旧版まで失う窓を無くす)。
    Dim purgeTags As String
    purgeTags = ChannelOriginTag(chName)
    If LenB(leavingChannel) > 0 Then
        If StrComp(leavingChannel, chName, vbTextCompare) <> 0 Then
            purgeTags = purgeTags & "|" & ChannelOriginTag(leavingChannel)
        End If
    End If

    Dim got As Long
    got = modPack.ImportPackFile(localPath, True, ChannelOriginTag(chName), purgeTags)
    On Error Resume Next
    Kill localPath
    On Error GoTo 0

    If got > 0 Then
        ' 離れた部門の版数は空へ。「アクティブな部門だけが版数を持つ」という
        ' 不変条件を保つ(PendingUpdates と SwitchTo の両方がこれに依存する)。
        If LenB(leavingChannel) > 0 Then
            If StrComp(leavingChannel, chName, vbTextCompare) <> 0 Then
                modStats.SetStatText modShareRule.ChannelStatKey(leavingChannel), ""
            End If
        End If
        ' 2026-08-01(R12-2-1): remote は共有フォルダの version.txt から読んだ
        ' 未信頼テキスト。my_stats のセルへ書く前に数式インジェクション対策を
        ' 通す(セキュリティ監査3。modStats.SetStatValue自体は汎用関数のため
        ' この呼出元で適用する)。
        modStats.SetStatText modShareRule.ChannelStatKey(chName), modUtilText.SanitizeForCell(remote)
        modLog.LogUsage "channel_sync", chName, "version=" & remote & " chunks=" & got
        outOk = True
    Else
        outErrDesc = "パックを読み込めませんでした(取り込めたまとまりが0件)"
    End If
    SyncChannel = got
    On Error GoTo 0
End Function

' 部門チャンネル由来のチャンクに付ける origin タグ。
' 書く側(modPack.ImportChunksDedup)と消す側(PurgeChannelChunks)が
' 必ず同じ文字列を使うよう、組み立てはこの1関数に集約する(レビュー C-1)。
Public Function ChannelOriginTag(ByVal chName As String) As String
    If LenB(Trim$(chName)) = 0 Then Exit Function
    ChannelOriginTag = "channel:" & Trim$(chName)
End Function

' 購読中の全チャンネルを同期する(利用者が明示的に押したときだけ呼ぶ)。
' outFailN に「更新があるのに取り込めなかった部門数」を返す(R8 F11)。
' 呼び出し側が「0件=すべて最新」と誤って報告しないための材料。
' 2026-07-31(R11-H Med5): 「成功だが取り込むものが無かった」部門を失敗として
' 数えていたため、通知が自己矛盾していた。SyncChannel の outOk で分ける。
' outLastErr には最後に起きた失敗の理由を返す(呼び出し側のE0801の材料)。
Public Function SyncSubscribed(Optional ByRef outFailN As Long = 0, _
                               Optional ByRef outLastErr As String) As Long
    On Error Resume Next
    outFailN = 0
    outLastErr = ""
    Dim pend As String: pend = PendingUpdates()
    If LenB(pend) = 0 Then Exit Function
    Dim parts() As String: parts = Split(pend, "|")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim chOk As Boolean, chErr As String
        Dim got As Long: got = SyncChannel(parts(i), "", chOk, chErr)
        If got > 0 Then SyncSubscribed = SyncSubscribed + got
        If Not chOk Then
            outFailN = outFailN + 1
            outLastErr = parts(i) & ": " & chErr
        End If
    Next i
    On Error GoTo 0
End Function

' 共有フォルダのファイルを%TEMP%へコピーする(3回までリトライ)。
' サーバの同時接続上限やAVスキャンで一時的に失敗することがあるため、
' 失敗しても例外を出さず空文字を返して静かに次回へ回す。
Private Function CopyToTemp(ByVal src As String, ByVal tag As String) As String
    Dim tmpDir As String: tmpDir = Environ$("TEMP")
    If LenB(tmpDir) = 0 Then tmpDir = Environ$("TMP")
    If LenB(tmpDir) = 0 Then Exit Function
    If Right$(tmpDir, 1) <> "\" Then tmpDir = tmpDir & "\"

    ' 2026-07-31(レビュー R8・低): 一時名に Timer 由来のサフィックスを足す。
    ' 秒精度だけだと、同じ秒に2つのチャンネルを同期したとき(全部門一括取込の
    ' 実行中は普通に起こる)同じ一時ファイル名になり、片方が他方の Kill と
    ' 競合する。部門名も入っているので衝突は稀だが、稀な失敗ほど原因が追えない。
    Dim dest As String
    dest = tmpDir & "nexus_ch_" & SafeName(tag) & "_" & Format$(Now, "hhnnss") & _
           "_" & Format$(Int(Timer * 1000) Mod 100000, "00000") & ".xlsx"

    Dim attempt As Long
    For attempt = 1 To 3
        On Error Resume Next
        Err.Clear
        FileCopy src, dest
        If Err.Number = 0 Then
            On Error GoTo 0
            CopyToTemp = dest
            Exit Function
        End If
        On Error GoTo 0
        Wait_ 200 * attempt
    Next attempt
End Function

Private Function SafeName(ByVal s As String) As String
    Dim t As String: t = s
    Dim bad As Variant
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|", " ")
        t = Replace(t, CStr(bad), "_")
    Next bad
    SafeName = modUtil.SafeLeft(t, 24)
End Function

Private Sub Wait_(ByVal ms As Long)
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000# < ms
        DoEvents
        If Timer < t0 Then Exit Do
    Loop
End Sub

' 起動直後の一斉アクセスを散らすためのランダム待機(0～jitterMs)。
' 全員が同一ミリ秒に共有フォルダへ殺到する状況を構造的に避ける。
Public Sub StartupJitter()
    On Error Resume Next
    Dim maxMs As Long: maxMs = modConfig.GetLong("startup_jitter_ms", 3000)
    If maxMs <= 0 Then Exit Sub
    Randomize
    Wait_ CLng(Rnd() * maxMs)
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' チャンク予算。上限に当たってから慌てないよう、常に使用率を見せる。
' ----------------------------------------------------------------------------
Public Function ChunkLimit() As Long
    On Error Resume Next
    ChunkLimit = modConfig.GetLong("chunk_limit", 20000)
    On Error GoTo 0
    If ChunkLimit < 1000 Then ChunkLimit = 20000
End Function

' R33 W5-17: 使用率の分母(ハード上限)。取込の実上限は shelf_max_chunks なので、
' 見せる分母もこれに揃える(Hubとダッシュボードで同じ数字になる唯一の情報源)。
Public Function ShelfMaxChunks() As Long
    On Error Resume Next
    ShelfMaxChunks = modConfig.GetLong("shelf_max_chunks", modAppDef.DEFAULT_SHELF_MAX_CHUNKS)
    On Error GoTo 0
    If ShelfMaxChunks <= 0 Then ShelfMaxChunks = modAppDef.DEFAULT_SHELF_MAX_CHUNKS
End Function

Public Function ChunkUsagePercent() As Long
    Dim used As Long
    On Error Resume Next
    used = modShelf.TotalChunks()
    On Error GoTo 0
    ' 算数と 0..100 クランプは modShareRule(純ロジック)が単一情報源。R33 W5-17。
    ChunkUsagePercent = modShareRule.UsagePercentOf(used, ShelfMaxChunks())
End Function

' 上限に近いか(chunk_limit の8割超)。Hubの警告表示とツールバーの棚卸し導線に使う。
' R33 W5-17: 表示のパーセント(分母=shelf_max_chunks)からは導かない。導くと
' shelf_max_chunks だけ広げた組織で警告が永久に出なくなる。
Public Function IsBudgetTight() As Boolean
    Dim used As Long
    On Error Resume Next
    used = modShelf.TotalChunks()
    On Error GoTo 0
    IsBudgetTight = modShareRule.IsBudgetTightAt(used, ChunkLimit())
End Function

' ----------------------------------------------------------------------------
' PurgeChannelChunks - あるチャンネル由来のチャンクを本棚から取り除く。
'   チャンネル切替と版の入れ替えの両方で使う。戻り値=消した件数。
'
' 2026-07-28(レビュー C-1): 探すタグが "pack:"&部門名 だったのに対し、
' 取込側が書いていたのは "pack:"&作者名 で、部門名≠作者名である限り
' 削除は常に0件だった。行削除そのものは modShelfStore にも同型の実装が
' あり「同じ概念の実装が2つある」ことがタグずれの温床だったため、
' 実処理は modShelfStore.RemoveRowsByOrigin へ寄せ、ここはタグの決定と
' 画面更新の抑止だけを持つ。
' ----------------------------------------------------------------------------
' outOk: 書き戻しが全行成功したか(R33H F2)。False なら呼び出し元は
'   「削除しました」と言ってはいけない(旧実装は失敗を握り潰していた)。
Public Function PurgeChannelChunks(ByVal chName As String, _
                                   Optional ByRef outOk As Boolean = True) As Long
    On Error Resume Next
    outOk = True
    Dim tag As String: tag = ChannelOriginTag(chName)
    If LenB(tag) = 0 Then Exit Function

    Application.ScreenUpdating = False
    Dim prevCalc As Long: prevCalc = Application.Calculation
    Application.Calculation = -4135          ' xlCalculationManual

    Dim n As Long
    n = modShelfStore.RemoveRowsByOrigin(tag, outOk)

    Application.Calculation = prevCalc
    Application.ScreenUpdating = True

    If n > 0 Then modLog.LogUsage "channel_purge", chName, "removed=" & n
    PurgeChannelChunks = n
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' PurgeLegacyPackChunks - 旧仕様(origin="pack:<作者名>")で取り込まれた
'   チャンネル残骸をまとめて掃除する。戻り値=消した件数。
'
' レビュー C-1 の修正前に部門チャンネルへつないだ端末には、消せないまま
' 溜まった "pack:" 行が残っている。新旧のタグが混在すると
' 「切り替えたのに前の部門の出典で答える」が直らないため、移行時に一度だけ
' 全部消して、正しいタグで入れ直す。
'
' 注意: 手渡しパック(ImportPackDialog で取り込んだもの)も同じ "pack:" 
' 名前空間にいるため、一緒に消える。これは区別する手段が本棚側に無いため
' の割り切りで、消えた手渡しパックは再取込で戻せる(呼び出し側が利用者へ
' その旨を伝えること)。
' ----------------------------------------------------------------------------
Public Function PurgeLegacyPackChunks() As Long
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim prevCalc As Long: prevCalc = Application.Calculation
    Application.Calculation = -4135

    Dim n As Long
    n = modShelfStore.RemoveRowsByOriginPrefix("pack:")

    Application.Calculation = prevCalc
    Application.ScreenUpdating = True

    If n > 0 Then modLog.LogUsage "channel_purge_legacy", "", "removed=" & n
    PurgeLegacyPackChunks = n
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' アクティブチャンネル(都度切替方式)
' ----------------------------------------------------------------------------
'   実機要望(2026-07-26): 1部門の正典が1,000チャンクで収まる想定は甘かった。
'   商品部だけで教えてBOX 1年分のQ&Aに加え、マニュアル・約款の読み方・
'   商品パンフレットの解説まで入れると5,000チャンクを軽く超える。
'   全部門を常駐させると本棚が埋まり、マイ本棚に使える枠が無くなる。
'
'   そこで「常駐は1部門だけ」にする。利用者は聞きたい分野の部門をクリックし、
'   同期が終わったら質問する。別の分野に移るときは部門を切り替える。
'   切替時に前の部門のチャンクは丸ごと消えるので、本棚は常に
'   「マイ本棚 + いま選んでいる1部門」だけになる。
'
'   トレードオフ(利用者に必ず伝えること):
'     ・切替に時間がかかる(5,000チャンクで数十秒～)
'     ・複数部門にまたがる質問は一度にできない
'     ・「今どこにつないでいるか」が分からないと混乱する
'       → Hubとチャット画面のヘッダーに常時表示して迷子を防ぐ
Public Function ActiveChannel() As String
    On Error Resume Next
    ActiveChannel = Trim$(modConfig.GetString("active_channel", ""))
    On Error GoTo 0
End Function

' SwitchTo - 常駐チャンネルを切り替える。戻り値=取り込んだチャンク数
'   (-1 = 既に最新、0 = 失敗)。
'
' 2026-07-28(レビュー H-14): 以前はここで「先に旧部門を消してから
' SyncChannel を呼ぶ」形になっていた。版の読み取り・pack の存在確認・
' %TEMP% へのコピー(リトライ込み)は全部 SyncChannel の中にあるので、
' 共有の瞬断や発行者の上書きロックでコピーに失敗すると
'   旧部門は消えた / 新部門は入らない / active_channel は空
' の三重苦で終わっていた。旧部門の削除は SyncChannel の中の
' 「読み込み完了後・書き込み直前」へ移し、ここは順序を決めるだけにする。
' active_channel も、取り込みが成功してから書き換える。
'
' 2026-07-31(レビュー R8・低): 【この関数は現在どこからも呼ばれていない】。
' 「1部門だけを常駐させて切り替える」という v2 の考え方の名残で、v3 では
' SubscribeAllAvailable(全部門をまとめて購読する)へ移行済み。
' 削除しないのは、過去に配布したブックの Shape に OnAction として
' "modChannel.SwitchTo" が残っている可能性を潰しきれていないため
'(残っていると押した瞬間に「マクロが見つかりません」で止まる)。
' OnAction 残存の再調査が済むまでは、消さずにこの注記だけ置く。
Public Function SwitchTo(ByVal chName As String) As Long
    On Error Resume Next
    Dim cur As String: cur = ActiveChannel()

    ' 同じ部門を選び直した場合は、更新があるときだけ入れ替える。
    If StrComp(cur, chName, vbTextCompare) = 0 Then
        If RemoteVersion(chName) = LocalVersion(chName) Then
            SwitchTo = -1          ' -1 = 既に最新(呼び出し側が案内に使う)
            Exit Function
        End If
    End If

    Dim got As Long
    got = SyncChannel(chName, cur)
    If got > 0 Then
        modConfig.SetValue "active_channel", chName
    End If
    ' 失敗時は active_channel を触らない。旧部門の中身も消えていないので、
    ' 「つないだつもりで実は空」にはならず、前の状態のまま留まる。
    SwitchTo = got
    On Error GoTo 0
End Function

' 今つないでいる部門の表示用ラベル(ヘッダーに常時出す)。
Public Function ActiveLabel() As String
    On Error Resume Next
    Dim have As Long, all As Long
    have = SubscribedCount(all)
    ActiveLabel = ActiveLabelText(have, all)
    On Error GoTo 0
End Function

' ActiveLabelText(R20-4c) - 表示文言の組み立てだけを分離した純関数
'   (ゴールデン対象)。have<=0は「未読込」ではなく「未設定」とし、共有
'   フォルダ配下の部門チャンネルが未接続であることを明示する。
'   2026-08-06 R20H FA-6: このラベルShapeにクリックハンドラは無い(下記
'   【未配線・報告事項】)。「クリックで設定」は実行不能な案内だったため、
'   実際に辿れる導線(❓ヘルプの⚙から設定)へ差し替える。
Public Function ActiveLabelText(ByVal have As Long, ByVal total As Long) As String
    If have <= 0 Then
        ActiveLabelText = ChrW(&HD83D) & ChrW(&HDCDA) & " 部門資料: 未設定(" & _
            ChrW(&H2753) & "ヘルプの" & ChrW(&H2699) & "から設定)"
    ElseIf have >= total Then
        ActiveLabelText = ChrW(&HD83D) & ChrW(&HDCDA) & " 全" & total & "部門"
    Else
        ActiveLabelText = ChrW(&HD83D) & ChrW(&HDCDA) & " " & have & "/" & total & "部門"
    End If
End Function

' 【未配線・報告事項】(R20-4c): このラベルは現在modUINexusDraw.DrawHeaderが
'   nx_top_bgのTextFrame2へ直接書き込む文字列で、Shapeそのものへの
'   クリックハンドラ配線が無い。部門チャンネルは共有フォルダ配下のため
'   設定UIはmodHelp.OnShareSetupと共通にできるはずだが、(a)配線先の
'   modUINexusDraw.basはR20-4波の許可ファイル一覧に無く、(b)そもそも
'   modChannel(src/pack=部品層)からはUI層(modHelp/modUiLock)を直接
'   参照できない(R1層規約)。ハンドラ自体もUI層側に置く必要があるため、
'   本波では文言変更のみに留め、配線は司令塔判断へ委ねる(完了報告に記載)。

' SubscribedCount - 読み込み済みの部門数(totalに発見できた総数を返す)。
Public Function SubscribedCount(ByRef total As Long) As Long
    On Error Resume Next
    Dim all As String: all = ListChannels()
    If LenB(all) = 0 Then Exit Function
    Dim parts() As String: parts = Split(all, "|")
    total = UBound(parts) - LBound(parts) + 1
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If LenB(LocalVersion(parts(i))) > 0 Then SubscribedCount = SubscribedCount + 1
    Next i
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' MigrateOriginNamespace - 旧タグ("pack:<作者名>")で入った部門チャンクを掃除する
'   一度だけの移行処理。戻り値=消した件数(0=移行不要または対象なし)。
'
' 背景(レビュー C-1): 修正前のビルドで部門チャンネルへつないだ端末には、
' origin="pack:<作者名>" のチャンクが消せないまま残っている。新タグ
' ("channel:<部門名>")で入れ直しても、旧タグの行は誰も消さないので
' 「切り替えたのに前の部門の出典で答える」「使用量だけ増える」が続く。
'
' 判定を「チャンネルを使ったことがある端末」に限っているのは、手渡しパック
' (ImportPackDialog)も同じ "pack:" 名前空間にいるため。チャンネルを一度も
' 使っていない端末で消すのは、ただのデータ破壊にしかならない。
'
' 実行後は全チャンネルのローカル版数を空にして、次の同期で正しいタグで
' 入れ直させる。
' ----------------------------------------------------------------------------
Public Function MigrateOriginNamespace() As Long
    On Error Resume Next
    If LenB(modStats.GetStatText(MIGRATION_KEY)) > 0 Then Exit Function

    Dim all As String: all = ListChannels()
    If LenB(all) = 0 Then
        ' 共有フォルダが見えない/チャンネルが無い端末。今回は判断できないので
        ' フラグも立てない(次にチャンネルが見えたときに改めて判定する)。
        Exit Function
    End If

    Dim parts() As String: parts = Split(all, "|")
    Dim usedChannels As Boolean
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If LenB(LocalVersion(parts(i))) > 0 Then usedChannels = True
    Next i

    If usedChannels Then
        MigrateOriginNamespace = PurgeLegacyPackChunks()
        For i = LBound(parts) To UBound(parts)
            modStats.SetStatText modShareRule.ChannelStatKey(parts(i)), ""
        Next i
        modLog.LogUsage "channel_origin_migration", "", _
            "removed=" & MigrateOriginNamespace & " channels=" & UBound(parts) - LBound(parts) + 1
    End If

    modStats.SetStatText MIGRATION_KEY, modUtil.NowStamp()
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' SubscribeAllAvailable - 見つかった部門を「全部」読み込む/更新する。
' ----------------------------------------------------------------------------
' 本モジュール冒頭のコメントは、こう書いてある:
'   「必要なものだけ購読」は誤りだった。聞く前に分野を判断して購読操作を
'   させた時点で、ポータルを探し回るのと同じ認知負荷が発生する。
'   解決したかった問題を作り直してしまう。既定で全チャンネルを購読する。
' ところが実装は SwitchTo 一本で、1部門だけを常駐させ、切り替えのたびに
' 前の部門を消していた。書き残した失敗モードを、そのまま作ってしまっていた。
' (DESIGN_v3 §9 は SubscribeAllAvailable が実装済みと書いていたが、
'  この関数は存在しなかった。ここで実際に用意する。)
'
' 全部入れて成立するのか:
'   チャンク上限は modAppDef.DEFAULT_SHELF_MAX_CHUNKS(20,500。config
'   shelf_max_chunks で変更可)。2026-08-01(R12-3-11): ここは 20,000 という
'   旧値のままだった(R12-1-8 で3箇所に散っていたフォールバックを定数へ
'   統一した際、この説明文だけが残った)。正典はQ&A粒度で
'   1件1～2チャンクなので、6部門×500件でも 3,000～6,000。個人の本棚
'   (20冊で～2,800)を足しても 9,000 前後で収まる。数字は足りている。
'
' データ構造上も問題ない:
'   各チャンネルのチャンクは origin="channel:<部門名>" で区別され
'   (2026-07-31 R8・低: ここは "pack:<部門名>" と書いてあった。レビュー C-1 の
'    名前空間分離より前の記述で、実装は ChannelOriginTag が返す
'    "channel:<部門名>"。誤ったまま残すと、次にここを読んだ人が
'    「タグは pack: だ」と信じて F1 と同型のバグを作り直す)、
'   PurgeChannelChunks はその部門の分だけを消す。SyncChannel も取り込み前に
'   自分の旧版だけを掃除する。つまり複数部門の同居は元から成立していて、
'   単一常駐を強制していたのは SwitchTo の「前の部門を消す」1行だけだった。
'
' 予算保護:
'   上限の9割に達したらそこで止め、何を入れて何を見送ったかを返す。
'   黙って打ち切らない(残りは版が変わったときに改めて入る)。
Public Function SubscribeAllAvailable() As String
    On Error Resume Next
    Dim all As String: all = ListChannels()
    If LenB(all) = 0 Then
        SubscribeAllAvailable = ""
        Exit Function
    End If

    Dim parts() As String: parts = Split(all, "|")
    Dim okN As Long, skipN As Long, chunkN As Long
    Dim doneNames As String, skipNames As String

    Dim guard As Long: guard = CLng(ChunkLimit() * 0.9)
    If guard < 1 Then guard = 18000

    ' 2026-07-31(レビュー R8 F11): 失敗件数(failN)を数える。
    ' 従来は「取り込めた部門が0件」を無条件に「すべて最新です」と報告して
    ' いた。発行者が pack.xlsx を書き換えている最中(=読めない)や共有の
    ' 瞬断でも同じ文言が出るため、利用者は【古い正典のまま】仕事を続ける。
    ' 「更新が無かった」のか「読めなかった」のかは、正直に区別して伝える。
    Dim failN As Long
    Dim i As Long
    ' 2026-07-31(R11-E H-3): 部門ぶんの読み込みは数分かかることがあるのに
    ' 進捗が何も出ていなかった(監査1)。実際に同期する部門だけ「N/M部門」を
    ' 進捗バナーに出す(mid層からのShowProgress/HideProgressはlint許容済み)。
    Dim nSub As Long: nSub = UBound(parts) - LBound(parts) + 1
    For i = LBound(parts) To UBound(parts)
        Dim nm As String: nm = parts(i)
        If LenB(nm) > 0 And IsSubscribed(nm) Then
            If RemoteVersion(nm) <> LocalVersion(nm) Then
                If modShelf.TotalChunks() >= guard Then
                    skipN = skipN + 1
                    skipNames = skipNames & IIf(LenB(skipNames) > 0, "/", "") & nm
                Else
                    modUIMain.ShowProgress modUtil.ProgressText(i - LBound(parts) + 1, nSub, "") & _
                        " " & nm & " を読み込み中…"
                    Dim got As Long
                    Dim chOk2 As Boolean, chErr2 As String
                    got = SyncChannel(nm, "", chOk2, chErr2)
                    If got > 0 Then
                        okN = okN + 1
                        chunkN = chunkN + got
                        doneNames = doneNames & IIf(LenB(doneNames) > 0, "/", "") & nm
                    ElseIf Not chOk2 Then
                        ' 2026-07-31(R11-H Med5): 「成功で0件」は失敗に数えない。
                        failN = failN + 1
                    End If
                End If
            End If
        End If
    Next i
    modUIMain.HideProgress

    ' 「どの部門につないでいるか」という単一の概念は無くなったが、
    ' 既存UI(ヘッダー表示・お知らせ)が空文字を「未接続」と解釈するため、
    ' 1つでも入っていることを示す目印として先頭の部門名を入れておく。
    If okN > 0 Then modConfig.SetValue "active_channel", parts(LBound(parts))

    ' 文面の組み立ては純ロジック(modShareRule.SyncSummaryText)へ寄せて
    ' 「okN=0 かつ failN>0 のときに『すべて最新です』と言わない」をテストで固定する。
    Dim sb As String
    sb = modShareRule.SyncSummaryText(okN, failN, chunkN, doneNames)
    If skipN > 0 Then
        sb = sb & vbLf & "本棚の空きが足りないため " & skipN & "部門は見送りました(" & skipNames & ")。"
    End If
    SubscribeAllAvailable = sb
    On Error GoTo 0
End Function


' ----------------------------------------------------------------------------
' 内部
' ----------------------------------------------------------------------------
' 到達性の判定は modShare が1セッション1回だけ行う(レビュー M-20)。
' 共有が死んでいる日に、起動のたび複数箇所で OS のタイムアウトを
' 払い直すのをやめる(「開かない」「閉じたのに数十秒残る」の主因)。
Private Function ChannelsDir() As String
    ChannelsDir = modShare.SubDir(CH_SUBDIR)
End Function

Private Sub EnsureDir(ByVal folderPath As String)
    On Error Resume Next
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    On Error GoTo 0
End Sub

' UTF-8書き出しの実体は modUtilText.WriteTextFileUtf8(2026-07-31 R11-F2で
' 9箇所の同型実装を1本化)。戻り値の意味は従来どおり(成功=True/失敗=False・
' 例外は外へ出さない)。
Private Function WriteShared(ByVal filePath As String, ByVal content As String) As Boolean
    WriteShared = modUtilText.WriteTextFileUtf8(filePath, content)
End Function

' UTF-8読み取りの実体は modUtilText.ReadTextFileUtf8(2026-07-31 R11-F2で
' 10箇所の同型実装を1本化)。
Private Function ReadShared(ByVal filePath As String, ByRef outText As String) As Boolean
    ReadShared = modUtilText.ReadTextFileUtf8(filePath, outText)
End Function
