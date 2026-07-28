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
    Dim nm As String
    nm = Dir(baseDir & "*", vbDirectory)
    Do While LenB(nm) > 0 And n < MAX_CHANNELS
        If nm <> "." And nm <> ".." And Left$(nm, 1) <> "_" Then
            names(n) = nm
            n = n + 1
        End If
        nm = Dir()
    Loop

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

' 購読中か。既定は「全部購読」なので、判定は除外リスト(config
' unsubscribed_channels)に載っていないこと。新しい部門が正典を発行したら
' 誰も操作しなくても自動的に届く。これが v3 の要。
Public Function IsSubscribed(ByVal chName As String) As Boolean
    On Error Resume Next
    Dim ex As String: ex = "," & modConfig.GetString("unsubscribed_channels", "") & ","
    IsSubscribed = (InStr(1, ex, "," & chName & ",", vbTextCompare) = 0)
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
    Dim ex As String: ex = modConfig.GetString("unsubscribed_channels", "")
    Dim parts() As String: parts = Split(ex, ",")
    Dim sb As String
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim one As String: one = Trim$(parts(i))
        If LenB(one) > 0 And StrComp(one, chName, vbTextCompare) <> 0 Then
            If LenB(sb) > 0 Then sb = sb & ","
            sb = sb & one
        End If
    Next i
    modConfig.SetValue "unsubscribed_channels", sb
    On Error GoTo 0
End Sub

' 購読をやめる(除外リストへ追加)。
Public Sub Unsubscribe(ByVal chName As String)
    On Error Resume Next
    If Not IsSubscribed(chName) Then Exit Sub
    Dim ex As String: ex = Trim$(modConfig.GetString("unsubscribed_channels", ""))
    If LenB(ex) > 0 Then ex = ex & ","
    modConfig.SetValue "unsubscribed_channels", ex & chName
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
'   版が変わっているときは、旧版で取り込んだチャンクを削除してから入れ直す
'   (差分ではなく置換)。これをやらないと改定のたびに古い記述が残り、
'   AIが古い条文を根拠に答えるという最悪の事故になる。
'   戻り値 = 取り込んだチャンク数(0=更新なし/失敗、-1は使わない)。
'
' leavingChannel (2026-07-28 レビュー H-14):
'   部門を切り替えるときに「離れる部門」を渡す。旧部門の削除は
'   SwitchTo ではなくここで行う。SwitchTo が先に消していた頃は、
'   共有瞬断や発行者の上書きロックで pack のコピーに失敗すると
'   「旧部門は消えた・新部門は入らない・active_channel は空」という
'   三重苦で終わっていた。破壊的な操作は、失敗しうる I/O を全部
'   終えてから始める。
' ----------------------------------------------------------------------------
Public Function SyncChannel(ByVal chName As String, Optional ByVal leavingChannel As String = "") As Long
    On Error Resume Next
    Dim remote As String: remote = RemoteVersion(chName)
    If LenB(remote) = 0 Then Exit Function
    ' 切替時(leavingChannel あり)は「版が同じ」でも処理を通す。旧部門を
    ' 消して新部門を入れ直す必要があるため。
    If LenB(leavingChannel) = 0 Then
        If remote = LocalVersion(chName) Then Exit Function
    End If

    Dim packPath As String: packPath = ChannelsDir() & chName & "\" & PACK_NAME
    If LenB(Dir(packPath)) = 0 Then Exit Function

    ' 朝の一斉起動で数千人が同じファイルを掴みに行くため、共有側に読み取り
    ' ロックを残さないよう %TEMP% へコピーしてから開く。発行者が書き換えても
    ' こちらが読むのはコピーなので壊れたファイルを掴まない。
    Dim localPath As String
    localPath = CopyToTemp(packPath, chName)
    If LenB(localPath) = 0 Then Exit Function

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
                modStats.SetStatText "ch:" & LCase$(leavingChannel), ""
            End If
        End If
        modStats.SetStatText "ch:" & LCase$(chName), remote
        modLog.LogUsage "channel_sync", chName, "version=" & remote & " chunks=" & got
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

' 共有フォルダのファイルを%TEMP%へコピーする(3回までリトライ)。
' サーバの同時接続上限やAVスキャンで一時的に失敗することがあるため、
' 失敗しても例外を出さず空文字を返して静かに次回へ回す。
Private Function CopyToTemp(ByVal src As String, ByVal tag As String) As String
    Dim tmpDir As String: tmpDir = Environ$("TEMP")
    If LenB(tmpDir) = 0 Then tmpDir = Environ$("TMP")
    If LenB(tmpDir) = 0 Then Exit Function
    If Right$(tmpDir, 1) <> "\" Then tmpDir = tmpDir & "\"

    Dim dest As String
    dest = tmpDir & "nexus_ch_" & SafeName(tag) & "_" & Format$(Now, "hhnnss") & ".xlsx"

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
    On Error Resume Next
    IsBudgetTight = (ChunkUsagePercent() >= 80)
    On Error GoTo 0
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
Public Function PurgeChannelChunks(ByVal chName As String) As Long
    On Error Resume Next
    Dim tag As String: tag = ChannelOriginTag(chName)
    If LenB(tag) = 0 Then Exit Function

    Application.ScreenUpdating = False
    Dim prevCalc As Long: prevCalc = Application.Calculation
    Application.Calculation = -4135          ' xlCalculationManual

    Dim n As Long
    n = modShelfStore.RemoveRowsByOrigin(tag)

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
    If have <= 0 Then
        ActiveLabel = ChrW(&HD83D) & ChrW(&HDCDA) & " 部門: 未読込"
    ElseIf have >= all Then
        ActiveLabel = ChrW(&HD83D) & ChrW(&HDCDA) & " 全" & all & "部門"
    Else
        ActiveLabel = ChrW(&HD83D) & ChrW(&HDCDA) & " " & have & "/" & all & "部門"
    End If
    On Error GoTo 0
End Function

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
            modStats.SetStatText "ch:" & LCase$(parts(i)), ""
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
'   チャンク上限は 20,000(config shelf_max_chunks)。正典はQ&A粒度で
'   1件1～2チャンクなので、6部門×500件でも 3,000～6,000。個人の本棚
'   (20冊で～2,800)を足しても 9,000 前後で収まる。数字は足りている。
'
' データ構造上も問題ない:
'   各チャンネルのチャンクは origin="pack:<部門名>" で区別され、
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

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim nm As String: nm = parts(i)
        If LenB(nm) > 0 And IsSubscribed(nm) Then
            If RemoteVersion(nm) <> LocalVersion(nm) Then
                If modShelf.TotalChunks() >= guard Then
                    skipN = skipN + 1
                    skipNames = skipNames & IIf(LenB(skipNames) > 0, "/", "") & nm
                Else
                    Dim got As Long
                    got = SyncChannel(nm)
                    If got > 0 Then
                        okN = okN + 1
                        chunkN = chunkN + got
                        doneNames = doneNames & IIf(LenB(doneNames) > 0, "/", "") & nm
                    End If
                End If
            End If
        End If
    Next i

    ' 「どの部門につないでいるか」という単一の概念は無くなったが、
    ' 既存UI(ヘッダー表示・お知らせ)が空文字を「未接続」と解釈するため、
    ' 1つでも入っていることを示す目印として先頭の部門名を入れておく。
    If okN > 0 Then modConfig.SetValue "active_channel", parts(LBound(parts))

    Dim sb As String
    If okN > 0 Then
        sb = okN & "部門を読み込みました(" & doneNames & " / 合計" & chunkN & "件)。"
    Else
        sb = "すべて最新です。読み込み直すものはありませんでした。"
    End If
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
