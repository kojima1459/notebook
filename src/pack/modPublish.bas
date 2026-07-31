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
Private Const LOCK_NAME As String = "publish.lock"

' publish.lock を「前回の異常終了の残骸」とみなすまでの時間(分)。
' 発行そのものは数十秒で終わるので、10分もあれば正常な発行は必ず終わっている。
' 短すぎると大きな本棚の発行中に他の人が割り込み、長すぎると Excel が落ちた
' 端末のロック1つで部門の発行が数時間止まる(R8 F12)。
Private Const LOCK_STALE_MIN As Double = 10#

' いま自分が置いている publish.lock のフルパス(2026-07-31 R8b B7a)。
' 解放を ChannelsDir() から組み直すと、発行中に共有が落ちて modShare が
' セッションを降格したとき SubDir が空になり、自分のロックを外せなくなる。
' VBAはモジュールレベル宣言をプロシージャより前に置く必要がある。
Private mLockPath As String

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

' ============================================================================
' 同時発行の見張り(2026-07-31 レビュー R8 F12)
' ----------------------------------------------------------------------------
' 何が起きていたか:
'   発行は「旧版を退避 → pack.xlsx を共有へ直接 SaveAs → version.txt」の順で、
'   pack.xlsx の書込みが共有フォルダ上で数十秒続く。その間に別の担当者が
'   同じ部門を発行すると、
'     ・Aの pack を書いている最中に B が退避・上書きを始める
'     ・A の version.txt と B の pack.xlsx が組み合わさる
'   という、どちらの発行者にもエラーが出ないまま中身と版番号が食い違う
'   壊れ方をする。購読者は「新しい版だから」と信じて取り込む。
'
' 対策は2つ:
'   (1) 書込み窓を短くする。pack.xlsx はまず %TEMP% へ書き、出来上がった
'       ものを共有へ FileCopy する(共有側が中途半端な状態でいる時間が、
'       数十秒から1回のコピーに縮む)。→ StagePackPath / CommitPack
'   (2) 発行の間だけ publish.lock を置く。他の人はそれを見て中断する。
'       10分を超えた残骸は無視して続行する(ロック1つで発行機能が
'       永久に死ぬのを防ぐ)。→ AcquireLock / ReleaseLock
' ============================================================================

' AcquireLock - 発行ロックを取る。取れたら True。
'   取れなかったときは outHolder に「誰がいつ始めたか」を返す(案内に使う)。
Public Function AcquireLock(ByVal chName As String, ByRef outHolder As String) As Boolean
    On Error Resume Next
    outHolder = ""
    mLockPath = ""
    Dim d As String: d = PrepareDir(chName)
    If LenB(d) = 0 Then Exit Function

    Dim lockPath As String: lockPath = d & LOCK_NAME
    Dim hasLock As Boolean: hasLock = (LenB(Dir(lockPath)) > 0)

    Dim ageMin As Double
    If hasLock Then
        Err.Clear
        ageMin = (Now - FileDateTime(lockPath)) * 1440#
        ' 2026-07-31(R8b B7b): タイムスタンプが読めないときは 0 分(=新しい)
        ' ではなく、残骸側へ倒す。読めないロックで永久に発行できなくなるより、
        ' 上書きして続行する方が害が小さい(判断材料は err_log に残す)。
        If Err.Number <> 0 Then ageMin = LOCK_STALE_MIN
        Err.Clear
        Dim body As String
        If ReadShared(lockPath, body) Then outHolder = Trim$(Replace(body, vbTab, " / "))
    End If

    Select Case modShareRule.PublishLockAction(hasLock, ageMin, LOCK_STALE_MIN)
        Case "wait"
            ' 2026-07-31(C3): 呼び出し側は outHolder が空かどうかで
            ' 「他の人が発行中(待てば解ける)」と「ロックを作れなかった
            '  (待っても解けない)」を区別する。ここは前者なので、
            ' ロック本文が読めなかった(書込み中でロックされている等)場合でも
            ' 必ず何かを返す。空のまま返すと後者と誤診断される。
            If LenB(outHolder) = 0 Then outHolder = "(お名前を読み取れませんでした)"
            Exit Function
        Case "stale"
            ' 残骸とみなして続行する。誰がいつ残したかは記録に残す。
            modLog.LogUsage "publish_lock_stale", chName, _
                "古い publish.lock を上書きしました(経過" & CLng(ageMin) & "分): " & _
                modUtil.SafeLeft(outHolder, 120)
            ' 2026-07-31(R8b B7b): 時計ズレ(経過が負)で stale になった場合は、
            ' 原因を追えるように err_log にも残す。端末の時計が共有サーバより
            ' 進んでいると、正常な発行のロックまで残骸に見える。
            If ageMin < 0 Then
                modLog.LogError "E0807", "modPublish.AcquireLock", _
                    "publish.lock の更新時刻が未来です(端末の時計が共有サーバとずれている" & _
                    "可能性があります)。残骸とみなして発行を続行しました: " & chName
            End If
            outHolder = ""
    End Select

    ' 2026-07-31(R8b B2): ロックの作成に失敗したら【取れなかった】と答える。
    ' 従来は WriteShared の戻り値を捨てて無条件に True を返していたため、
    ' 権限不足や共有断でロックが1バイトも書けていないのに「取れた」と報告し、
    ' 2人が同時に発行へ進める = 見張りが存在しないのと同じ状態だった。
    ' しかも「誰もロックを見ない」ので、事故が起きたことすら誰も気付けない。
    If Not WriteShared(lockPath, PublisherName() & vbTab & Format$(Now, "yyyy-mm-dd hh:nn:ss")) Then
        modLog.LogError "E0807", "modPublish.AcquireLock", _
            "publish.lock を作成できませんでした(共有フォルダへの書き込み権限を" & _
            "ご確認ください)。発行は中止しました: " & chName
        Exit Function
    End If

    ' 2026-07-31(R8b B7a): 解放に使うフルパスをここで控える。
    ' ReleaseLock が ChannelsDir() から組み直していると、発行中に共有が
    ' 一時的に落ちて modShare がセッションを降格した場合、SubDir が空文字を
    ' 返して【自分が置いたロックを外せなくなる】。その部門は次の10分間
    ' 誰も発行できない。取得時のパスをそのまま Kill する。
    mLockPath = lockPath
    AcquireLock = True
    On Error GoTo 0
End Function

' ReleaseLock - 発行ロックを外す。発行が成功しても失敗しても必ず呼ぶ。
'   chName は記録用。実際に消すのは AcquireLock 成功時に控えたフルパス
'   (2026-07-31 R8b B7a。理由は AcquireLock 側のコメント参照)。
Public Sub ReleaseLock(ByVal chName As String)
    On Error Resume Next
    If LenB(mLockPath) = 0 Then Exit Sub
    Dim released As String: released = mLockPath
    Kill released
    Err.Clear
    mLockPath = ""

    ' 2026-07-31(C6): どの部門のロックをいつ外したかを usage_log に残す。
    ' publish.lock は「残ってしまうと10分間その部門の発行が止まる」ものなので、
    ' 取得(publish_lock_stale)と解放が対で追えないと、実機で
    ' 「発行できない」と言われたときに何が起きたのか復元できない。
    ' 引数 chName はこれまで使っておらず、記録の宛名として使う
    '(シグネチャと呼び出し元は変えない)。
    modLog.LogUsage "publish_lock_release", chName, _
        "発行の見張りを外しました: " & modUtil.SafeLeft(released, 160)
    On Error GoTo 0
End Sub

' StagePackPath - pack.xlsx を一旦置くローカル(%TEMP%)のパス。
'   共有フォルダ上で「書きかけの pack.xlsx」が見えている時間を短くするため、
'   重い書き出しは必ずローカルで済ませる(R8 F12)。
'   同じ秒に2人が同じ端末を使うことは無いが、同一端末での連続発行に備えて
'   Timer 由来のサフィックスも足す(modChannel.CopyToTemp と同じ作法)。
Public Function StagePackPath(ByVal chName As String) As String
    On Error Resume Next
    Dim tmpDir As String: tmpDir = Environ$("TEMP")
    If LenB(tmpDir) = 0 Then tmpDir = Environ$("TMP")
    If LenB(tmpDir) = 0 Then Exit Function
    If Right$(tmpDir, 1) <> "\" Then tmpDir = tmpDir & "\"

    StagePackPath = tmpDir & "nexus_pub_" & modUtil.Fnv1a64Hex(chName) & "_" & _
                    Format$(Now, "hhnnss") & "_" & _
                    Format$(Int(Timer * 1000) Mod 100000, "00000") & ".xlsx"
    On Error GoTo 0
End Function

' CommitPack - %TEMP% に出来上がったパックを共有の pack.xlsx へ据える。
'   コピーが済んだら一時ファイルは消す(%TEMP%に発行物を残さない)。
Public Function CommitPack(ByVal chName As String, ByVal stagePath As String) As Boolean
    On Error Resume Next
    If LenB(stagePath) = 0 Then Exit Function
    If LenB(Dir(stagePath)) = 0 Then Exit Function
    Dim d As String: d = PrepareDir(chName)
    If LenB(d) = 0 Then Exit Function

    CommitPack = CopyWithRetry(stagePath, d & PACK_NAME)
    Kill stagePath
    Err.Clear
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' ArchiveCurrent - 今ある pack.xlsx を _archive へ退避する。
'   戻り値 = 退避したファイル名(無ければ空)。発行の前に必ず呼ぶ。
'
'   outFailed(2026-07-31 R11-A C3): 「退避すべき旧版が実在したのに退避
'   できなかった」ときだけ True。旧実装は失敗も「旧版なし」も同じ空文字で
'   返していたため、退避に失敗しても発行はそのまま進み、pack.xlsx が
'   上書きされて【戻せる版が消えた】。発行画面が利用者に約束している
'   「まちがえても前の版に戻せます」が黙って不履行になる経路だった。
'   共有フォルダを準備できなかった場合(d が空)は失敗として扱わない。
'   その場合はこの後の配置も必ず失敗し、旧版は上書きされないため。
' ----------------------------------------------------------------------------
Public Function ArchiveCurrent(ByVal chName As String, _
                               Optional ByRef outFailed As Boolean = False) As String
    outFailed = False
    On Error Resume Next
    Dim d As String: d = PrepareDir(chName)
    If LenB(d) = 0 Then Exit Function
    If LenB(Dir(d & PACK_NAME)) = 0 Then Exit Function

    Dim stamp As String: stamp = Format$(Now, "yyyymmdd-hhnnss")
    Dim dest As String: dest = d & ARCHIVE_DIR & "\pack_" & stamp & ".xlsx"
    Err.Clear
    FileCopy d & PACK_NAME, dest
    Dim copyErr As String: copyErr = "err#" & Err.Number & " " & Err.Description
    Err.Clear
    If LenB(Dir(dest)) > 0 Then
        ArchiveCurrent = "pack_" & stamp & ".xlsx"
    Else
        outFailed = True
        modLog.LogError "E0808", "modPublish.ArchiveCurrent", _
            "旧版の退避に失敗(戻せる版が作れていない): " & _
            modUtil.SafeLeft(dest, 200) & " / " & copyErr
    End If

    ' version.txt も同じ名前で残す(巻き戻しで版番号ごと戻すため)。
    If LenB(Dir(d & VER_NAME)) > 0 Then
        Err.Clear
        FileCopy d & VER_NAME, d & ARCHIVE_DIR & "\ver_" & stamp & ".txt"
        If Err.Number <> 0 Then
            ' 巻き戻しは pack_*.xlsx だけで成立する(Rollback は版番号を
            ' 新しく書き直す)ので、ここの失敗では中止しない。記録は残す。
            modLog.LogUsage "publish_archive_ver_fail", chName, _
                "err#" & Err.Number & " " & Err.Description
            Err.Clear
        End If
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
    ' 2026-07-28(レビュー M-18): 版番号は秒精度。分精度だと
    ' 「発行 → 巻き戻し → 再発行」を同一分内でやったときに版文字列が
    ' 初回と一致してしまい、購読者側は「更新なし」と判断して配信されない。
    ' テストで連続発行すると現実に踏む。
    ver = Format$(Now, "yyyymmdd-hhnnss") & "|" & Format$(Date, "yyyy-mm-dd") & "|" & author

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
    ' 判定は【連結する前】に行う(R8・低)。連結後だと "<部門>\_archive\" が
    ' 付いている分だけ必ず長さ0より大きく、ガードとして働かない。
    Dim baseDir As String: baseDir = ChannelsDir()
    If LenB(baseDir) = 0 Then Exit Function
    Dim d As String: d = baseDir & chName & "\" & ARCHIVE_DIR & "\"
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
'   outArcFailed(2026-07-31 R11-D・A波発見事項3): 巻き戻す【前】に今の版を
'   退避する処理が失敗したとき True。E0808はArchiveCurrentが記録済みだが、
'   利用者には何も伝わっていなかった。この状態で巻き戻すと「巻き戻しを
'   取り消して元に戻す」ことができなくなる(いま配っている版が控え無しで
'   消える)ので、呼び出し元が必ず知らせる。巻き戻し自体は続行してよい
'   (誤った正典の配信を止めることの方が優先度が高い)。
Public Function Rollback(ByVal chName As String, ByVal archiveName As String, _
                         Optional ByRef outArcFailed As Boolean = False) As Boolean
    outArcFailed = False
    On Error Resume Next
    Dim d As String: d = PrepareDir(chName)
    If LenB(d) = 0 Then Exit Function
    Dim src As String: src = d & ARCHIVE_DIR & "\" & archiveName
    If LenB(Dir(src)) = 0 Then Exit Function

    ' 巻き戻す前に、今の版も退避しておく(巻き戻し自体を取り消せるように)。
    ArchiveCurrent chName, outArcFailed

    ' 2026-07-28(レビュー L-14): FileCopy をリトライする。
    ' 共有フォルダは他の人が読んでいる最中だと一時的に掴めない。
    ' 1回で諦めると「戻せませんでした」だけが出る。
    If Not CopyWithRetry(src, d & PACK_NAME) Then Exit Function

    ' 版番号は必ず「新しい値」にする。古い版番号に戻すと、既にその版を
    ' 取り込み済みの人が「変化なし」と判定して巻き戻しが届かない。
    Dim ver As String
    ' 2026-07-28(レビュー M-18): 版番号は秒精度。分精度だと
    ' 「発行 → 巻き戻し → 再発行」を同一分内でやったときに版文字列が
    ' 初回と一致してしまい、購読者側は「更新なし」と判断して配信されない。
    ' テストで連続発行すると現実に踏む。
    ver = Format$(Now, "yyyymmdd-hhnnss") & "|" & Format$(Date, "yyyy-mm-dd") & "|" & _
          PublisherName() & "(巻き戻し:" & archiveName & ")"
    If Not WriteShared(d & VER_NAME, ver) Then
        ' 2026-07-28(レビュー L-14): pack は既に旧版へ置き換わっているのに
        ' 版番号だけ書けなかった状態。「戻せませんでした」と言うと嘘になる
        ' (中身は戻っている。届かないだけ)。実状をそのまま記録して返す。
        modLog.LogError "E0801", "modPublish.Rollback", _
            "pack は旧版へ置換済みだが version.txt を更新できなかった: " & chName
        Exit Function
    End If

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
    ' 連結前判定(R8・低)。
    Dim baseDir As String: baseDir = ChannelsDir()
    If LenB(baseDir) = 0 Then Exit Function
    Dim p As String: p = baseDir & chName & "\" & LOG_NAME
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
    ' 2026-07-31(レビュー R8・低): 判定は【連結する前】に行う。
    ' 連結後の LenB(p) は "publish_log.txt" が付いている分、共有が未設定でも
    ' 必ず 0 より大きく、このガードは一度も働いたことが無かった。
    ' 空パス + ファイル名 = カレントフォルダへ Open するので、共有未設定の
    ' 端末が自分のどこかに発行ログを作ってしまう(潜在欠陥の芽)。
    Dim baseDir As String: baseDir = PrepareDir(chName)
    If LenB(baseDir) = 0 Then Exit Sub
    Dim p As String: p = baseDir & LOG_NAME

    ' 2026-07-28(レビュー L-16): OSの追記モードで書く。
    ' 従来は「全部読む → 末尾に足す → 全部書く」だったため、2人が同時刻に
    ' 発行すると後から書いた方が相手の行を消していた(発行履歴は
    ' 「誰がいつ何を配ったか」の唯一の記録なので、静かに欠けると困る)。
    ' FreeFile は Integer を返すが、契約(Integer型禁止)に合わせて Long で受ける。
    Dim fnum As Long
    fnum = FreeFile
    Open p For Append As #fnum
    Print #fnum, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & PublisherName() & vbTab & line_
    Close #fnum
    On Error GoTo 0
End Sub

' 共有フォルダ上のファイルコピーを3回まで試す(他者が読んでいる最中は
' 一時的に掴めないため。レビュー L-14)。
Private Function CopyWithRetry(ByVal srcPath As String, ByVal destPath As String) As Boolean
    Dim attempt As Long
    For attempt = 1 To 3
        On Error Resume Next
        Err.Clear
        FileCopy srcPath, destPath
        If Err.Number = 0 Then
            On Error GoTo 0
            CopyWithRetry = True
            Exit Function
        End If
        On Error GoTo 0
        WaitMs 200 * attempt
    Next attempt
End Function

' 指定ミリ秒だけ待つ(Declareを増やさずTimerで回す。共有I/Oのリトライ間隔
' なので精度は要らない。日跨ぎでTimerが0へ戻ったら即抜ける)。
Private Sub WaitMs(ByVal ms As Long)
    On Error Resume Next
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000 < ms
        DoEvents
        If Timer < t0 Then Exit Do
    Loop
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

' 2026-07-31(レビュー R8 F2): modShare の関所を通す。発行は共有フォルダへの
' 書込みが本体なので、届かない共有では「重い書き出しを全部やってから失敗」
' という一番悪い順序になっていた。到達判定とルート解決は modShare だけが行う
'(modShare 冒頭「唯一性の原則」)。届かなければ空文字 = 既存の
' 「パスが空なら共有フォルダ未設定の案内を出す」経路にそのまま乗る。
Private Function ChannelsDir() As String
    On Error Resume Next
    ChannelsDir = modShare.SubDir(CH_SUBDIR)
    On Error GoTo 0
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
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup16
FailCleanup16:
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
    Resume FailCleanup17
FailCleanup17:
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
