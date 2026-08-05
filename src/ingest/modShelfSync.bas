Attribute VB_Name = "modShelfSync"
Option Explicit

' ============================================================================
' modShelfSync - 本棚フォルダの差分同期(MASTER_SPEC §7.2)
' ----------------------------------------------------------------------------
' 役割:
'   config shelf_folder で指定されたフォルダの第1階層を Dir() で走査し、
'   my_manifest(origin=self・このフォルダ配下の行のみ)と突き合わせて
'   新規/変更/消失を判定し、それぞれ IngestFile / DeleteSource へつなぐ。
'   sync_interval_min(config)>0のときは Application.OnTime で自己再帰予約
'   する定期自動同期にも対応する。
'
' 設計判断(R15-FixA: 同量圧縮。事実は落とさず言い方だけ縮めた):
'   ・同期スコープの限定: my_manifest には「＋資料を追加」から shelf_folder 外の
'     ファイルを取り込んだ行も混在する。「消失→削除」を適用してよいのは
'     shelf_folder 配下由来の行だけなので、file_path の親ディレクトリが
'     shelf_folder と一致する行だけを比較対象にする
'     (modShelfScan.LoadManifestScope)。フォルダ外の資料を誤って消さない。
'   ・拡張子フィルタ: Dir()走査には FileDialog の Filters が無いので、
'     modExtractor.SupportedExts() 外(Thumbs.db/.ini 等)は最初から除外する
'     (取込失敗の空騒ぎ・manifestの無駄な失敗行を防ぐ)。
'   ・差分判定(新規/変更/据え置き)は純関数 DiffDecision へ切り出しテスト可能に
'     する(MASTER_SPEC §7.8)。
'   ・フォルダごと見つからない場合(削除・リネーム・OneDriveオフライン)は個々の
'     ファイルの消失と区別する。全資料の DeleteSource は復旧不能なので、E0502を
'     出したうえで対象スコープを status="missing" にするだけに留める(§13)。
'     フォルダが復活すれば次回同期で通常の keep/replace 判定に戻る。
'   ・再入防止: mSyncRunning は「手動🔄連打」「自動同期の発火中に手動ボタン」を
'     防ぐ。VBAはシングルスレッドだが、長いループ中の DoEvents でShapeクリック
'     (OnAction)がその場で再入発火するため実際に起こり得る。行削除・圧縮と
'     EmbedPending の行インデックス食い違いは modEmbed 側で解決済み(冒頭参照)。
'   ・OnTime予約: 予約時刻(mNextRunTime)を保持し、CancelAutoSyncは同時刻を
'     指定して解除(Excel仕様・§12)。AutoSyncTickはSyncNow後に自己再予約する
'     自己再帰。Public必須なのはOnTimeがApplication.Runと同じ遅延バインドで
'     Private Subを解決できないため(発火時「マクロを実行できません」で自動
'     同期が永久に死ぬ)。CONTRACT/MASTER_SPEC §7.2 も同じ契約へ更新済み。
'   ・進捗実況は modEmbed/modEnrich と同じ作法で、1行スコープの On Error Resume
'     Next 越しに呼ぶ(表示が失敗しても同期自体は止めない)。
'   ・FileDialog/EnableCancelKey 等の名前付き定数は使わずリテラル値を使う
'     (LibreOffice互換のためのV2以来の慣習)。
' ============================================================================

Private mSyncRunning As Boolean     ' 再入防止
' ガードの自己回復(2026-07-16): 焼き付きで以降の同期が全て無言スキップに
' なるのを防ぐ開始時刻(失効判定は R15-1a でビート基準へ)。
Private mSyncRunningSince As Date
Private Const GUARD_EXPIRY_MIN As Long = 30
Private mScheduled As Boolean       ' OnTime予約中かどうか
Private mNextRunTime As Date        ' 予約時刻(Cancelは同時刻指定のため保持)
' 予約時刻の控え(ui_state)。VBAリセットでモジュール変数が消えても
' Auto_Close が解除できるようにするため(レビュー L-23)。
Private Const SCHED_KEY As String = "autosync_next"

' ----------------------------------------------------------------------------
' PickShelfFolder - フォルダ選択→config shelf_folder 保存→即SyncNow
' ----------------------------------------------------------------------------
Public Sub PickShelfFolder()
    Dim fd As Object
    Set fd = Application.FileDialog(4)   ' msoFileDialogFolderPicker(名前付き定数は使わない)
    fd.Title = "本棚フォルダを選んでください"

    If fd.Show <> -1 Then Exit Sub   ' キャンセル
    If fd.SelectedItems.count < 1 Then Exit Sub

    Dim folder As String: folder = CStr(fd.SelectedItems(1))
    modConfig.SetValue "shelf_folder", folder

    SyncNow
End Sub

' ----------------------------------------------------------------------------
' SyncNow - Dir()走査(第1階層のみ) vs my_manifest の差分同期
'   silent(既定False): Trueのとき、未設定/フォルダ不存在のE0502と完了サマリを
'   MsgBoxで出さず、SetStage(状態表示行+StatusBar)だけに留める。
'   modBoot.Boot(sync_on_open)とAutoSyncTick(定期自動同期)は利用者の操作
'   なしで走るので、対話ダイアログでフォーカスを奪ってはならない(§8)。配布
'   直後は shelf_folder 未設定のまま起動するため、名前を入力した直後に必ず
'   「フォルダが見つかりません」が出ていた。手動🔄・PickShelfFolder 直後は
'   従来どおり silent=False(操作への応答なので出した方がよい)。
' ----------------------------------------------------------------------------
Public Sub SyncNow(Optional ByVal silent As Boolean = False)
    ' 再入防止(連打・重複)。R15-1a: 失効は最後のビートから。
    If mSyncRunning Then
        If modShelfBatch.GuardExpiredNow(mSyncRunningSince, GUARD_EXPIRY_MIN) Then
            On Error Resume Next
            modLog.LogUsage "guard_recover", "sync", _
                "無音" & GUARD_EXPIRY_MIN & "分超で自動解除"
            On Error GoTo 0
            mSyncRunning = False
        End If
    End If
    If mSyncRunning Then Exit Sub

    ' R12-H-2a: 自動同期(silent)は他の処理中(UIロック取得中)には始めない
    ' (検索の重ループ中DoEventsに割り込むと結果が欠ける。手動🔄は対象外)。
    ' R12-4: 始める側は検索用ベクトルキャッシュを解放しピークを重ねない。
    ' R15-FixA(FA-3iii): modShelf.IsBusy()も判定に足す。modUiLock.IsBusyだけでは
    ' 取込ループ中(ロック非取得)の自動同期tickで二重取込+中断印の誤消去が
    ' 起きていた(取込中の入れ子自動同期。A-H3)。
    On Error Resume Next
    If silent Then
        If modUiLock.IsBusy() Or modShelf.IsBusy() Then
            modLog.LogUsage "autosync_deferred", "sync", "他の処理中のため見送りました"
            Exit Sub
        End If
    End If
    modVecCache.ResetVecCache
    On Error GoTo 0

    mSyncRunning = True
    mSyncRunningSince = Now
    ' R15波2の発見6(2026-08-04): 中断の印を同期の入口でも必ず下ろす。
    ' 残っていると、押した覚えの無い同期が最初の頁境界でいきなり止まる。
    ' R15-FixA(FA-3iv): ただし取込が動いているあいだは下ろさない。手動同期は
    ' silentではないので上の見送りを通らず、取込中でもここまで来られる。
    ' そこで印を下ろすと、利用者が押した中断が【無かったこと】にされる。
    If Not modShelf.IsBusy() Then modShelfBatch.ResetCancel

    ' 大量シート書換え中のイベント連鎖を抑止。Finishで必ずTrueへ戻す(死の連鎖防止)。
    On Error Resume Next
    Application.EnableEvents = False
    On Error GoTo 0

    ' 2026-07-16 恒久対策: どこで実行時エラーが起きても必ず Finish
    ' (mSyncRunningの解除)へ合流させる。従来は本体を覆うハンドラが無く、
    ' 例外が出ると mSyncRunning=True のまま抜けて同期が永久に走らなくなった。
    Dim uiStep As String
    On Error GoTo Failed

    uiStep = "状態表示の更新"
    On Error Resume Next
    modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDD04) & " 同期を確認しています…"
    On Error GoTo Failed

    uiStep = "本棚フォルダ設定の確認"
    Dim folder As String: folder = Trim$(modConfig.GetString("shelf_folder", ""))
    If LenB(folder) = 0 Then
        ' shelf_folder未設定は障害ではない(初回起動直後の正常な状態)。
        ' silent時はerr_logを汚さず、手動🔄のときだけ案内する(毎起動E0502が
        ' 3件ずつ積もり本当の障害が埋もれた実機報告への対応)。
        ' 【2026-07-16 袋小路の解消】partial資料のメモは「🔄同期で続きから
        ' 再開」と案内するのに、フォルダ未設定だとE0502で弾かれ再開手段が
        ' 無かった。未設定でも未完了のベクトル化(EmbedPending)だけは実行する。
        uiStep = "未完了ベクトル化の再開(フォルダ未設定)"
        Dim orphanPending As Long
        orphanPending = 0
        On Error Resume Next
        orphanPending = modEmbed.PendingCount()
        On Error GoTo Failed
        If orphanPending > 0 Then
            On Error Resume Next
            modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDCE5) & " 未完了の埋め込みを再開しています…"
            On Error GoTo Failed
            Dim orphanDone As Long
            orphanDone = modEmbed.EmbedPending()
            On Error Resume Next
            modUIMain.SetStage ""
            On Error GoTo Failed
            If Not silent Then
                MsgBox "同期フォルダはまだ設定されていませんが、未完了だった資料の変換(ベクトル化)を" & _
                       orphanDone & "件再開しました。" & vbLf & vbLf & _
                       "フォルダごと自動同期したい場合は「フォルダを選ぶ」から設定できます。", _
                       vbInformation, modAppDef.APP_NAME
            End If
        ElseIf Not silent Then
            modLog.ShowError "E0502", "modShelfSync.SyncNow", "shelf_folderが未設定です"
        Else
            On Error Resume Next
            modUIMain.SetStage ""
            On Error GoTo Failed
        End If
        GoTo Finish
    End If

    ' SharePoint/WebのURLはDir()で走査できない(実機ログ2026-07-16: httpsの
    ' URLがshelf_folderに設定されE0502になっていた)。原因が分かる案内を出す。
    ' OneDrive同期済みのローカルフォルダパスを使ってもらう。
    uiStep = "フォルダ形式の確認"
    If LCase$(Left$(folder, 4)) = "http" Then
        modLog.LogError "E0502", "modShelfSync.SyncNow", "URLは同期不可: " & modUtil.SafeLeft(folder, 200)
        If Not silent Then
            MsgBox "インターネット上のアドレス(SharePoint等のURL)は、本棚フォルダに設定できません。" & vbLf & vbLf & _
                   "SharePointのフォルダを使いたい場合は、まずOneDriveの「同期」ボタンで" & _
                   "パソコンのフォルダとして同期し、そのフォルダ(例: C:\Users\…\OneDrive - 会社名\…)を" & _
                   "「フォルダを選ぶ」から選んでください。" & vbLf & _
                   "(コード: E0502)", vbExclamation, modAppDef.APP_NAME
        End If
        On Error Resume Next
        modUIMain.SetStage ""
        On Error GoTo Failed
        GoTo Finish
    End If

    uiStep = "フォルダ存在の確認"
    If Not FolderExists(folder) Then
        If silent Then
            modLog.LogError "E0502", "modShelfSync.SyncNow", folder & "(silent)"
        Else
            modLog.ShowError "E0502", "modShelfSync.SyncNow", folder
        End If
        modShelfScan.MarkFolderScopeMissing folder
        On Error Resume Next
        modUIMain.SetStage ""
        On Error GoTo Failed
        GoTo Finish
    End If

    Dim folderNorm As String: folderNorm = modShelfScan.EnsureTrailingSlash(folder)

    uiStep = "フォルダ内ファイルの一覧取得"
    Dim diskNames() As String, diskCount As Long
    Dim enumFailed As Boolean
    modShelfScan.EnumFolderFiles folderNorm, diskNames, diskCount, enumFailed
    If enumFailed Then
        ' 列挙自体に失敗した。中身が分からない以上、消失判定は絶対にしない
        ' (レビュー H-7)。フォルダが消えたときと同じ扱いにして、
        ' 次回の同期で復帰させる。
        uiStep = "フォルダ列挙の失敗"
        modShelfScan.MarkFolderScopeMissing folderNorm
        modLog.LogError "E0502", "modShelfSync.SyncNow", _
            "フォルダを列挙できませんでした(権限/ネットワーク/ウイルス対策の可能性): " & folderNorm
        ' R11-D: エラー未発生のためFailed:(Resumeがerr#20を誤発生させる)へは
        ' 飛ばさず後始末へ直行(挙動不変・記録が実態どおりになるだけ)。
        GoTo FailedCleanup1
    End If

    uiStep = "資料台帳スコープの読込"
    Dim wsM As Worksheet: Set wsM = GetSheet(modAppDef.SH_MANIFEST)
    Dim mPaths() As String, mNames() As String, mModified() As Date
    Dim mSize() As Double, mStatus() As String, mCount As Long
    modShelfScan.LoadManifestScope wsM, folderNorm, mPaths, mNames, mModified, mSize, mStatus, mCount

    Dim diskDict As Object: Set diskDict = CreateObject("Scripting.Dictionary")
    Dim ingestedN As Long, replacedN As Long, deletedN As Long, failedN As Long
    Dim declinedN As Long: declinedN = 0   ' R18-6d: 見送り(declined)件数
    Dim resumeNeeded As Boolean: resumeNeeded = False

    ' 本棚上限(2026-07-16 実機対応): 大容量フォルダを同期すると、上限到達後の
    ' 残り全ファイルがIngestFile内のE0501ダイアログを1件ずつ出し続ける
    ' 「ダイアログ地獄」になっていた(実機で35連発)。上限はこのループ側で
    ' 先に判定し、到達後は取込を静かに見送って件数だけ数え、最後に1回だけ
    ' まとめて案内する。
    Dim capMax As Long: capMax = modConfig.GetLong("shelf_max_chunks", modAppDef.DEFAULT_SHELF_MAX_CHUNKS)
    If capMax < 1 Then capMax = modAppDef.DEFAULT_SHELF_MAX_CHUNKS
    Dim cappedN As Long: cappedN = 0
    ' R15-FixA(FA-3i): 中断のため手を付けずに見送ったファイル数。
    Dim cancelSkipN As Long: cancelSkipN = 0

    uiStep = "新規・更新の差分判定"
    Dim i As Long
    For i = 0 To diskCount - 1
        ' R15-FixA(FA-3i・レビューA-H3): ファイル境界の中断確認。同期にだけ
        ' これが無く、⏹中断を押しても残り全部が取り込まれ続けていた。しかも
        ' 途中で止まった資料は manifest が pending/partial のまま残り、次の同期は
        ' サイズも日時も変わっていないので "keep" と判定して二度と直さない
        ' (=恒久汚染)。残りは【manifestに一切触れず】見送り、件数だけ言う。
        If modShelfBatch.CancelRequested() Then
            cancelSkipN = diskCount - i
            Exit For
        End If

        If Not diskDict.Exists(LCase$(diskNames(i))) Then diskDict.Add LCase$(diskNames(i)), True

        ' R10-5: SetStageはShowProgress内部で呼ぶので二重呼び出しにしない。
        ' R15-FixA(FA-6): 中断できるのは取込・同期だけなので中断ボタン付きで出す。
        On Error Resume Next
        modShelfBatch.ShowIngestBanner "" & ChrW(&HD83D) & ChrW(&HDD04) & " 同期中 " & (i + 1) & "/" & diskCount & " …"
        On Error GoTo Failed

        ' 2026-07-31(R7 B-2): 1ファイルごとに1回だけメッセージを捌く。
        ' ここは「前のファイルの取込が完全に終わり、次のファイルの
        ' 抽出(COM)がまだ始まっていない」唯一の合間で、Word/Excelの
        ' オブジェクトを1つも掴んでいない。取込中に押されたボタンは
        ' modUiLock.BlockIfIngesting が実況を出して受け流す。
        DoEvents

        Dim path As String: path = folderNorm & diskNames(i)
        Dim mi As Long: mi = modShelfScan.FindManifestIndexByName(mNames, mCount, diskNames(i))
        Dim existsInManifest As Boolean: existsInManifest = (mi >= 0)

        Dim sizeChanged As Boolean: sizeChanged = False
        Dim timeChanged As Boolean: timeChanged = False
        If existsInManifest Then
            Dim curSize As Double: curSize = modShelfScan.SafeFileLen(path)
            Dim curMod As Date: curMod = modShelfScan.SafeFileDateTime(path)
            sizeChanged = (mSize(mi) <> curSize)
            timeChanged = Not modUtil.IsSameTimestamp(mModified(mi), curMod)
        End If

        Dim decision As String: decision = DiffDecision(existsInManifest, sizeChanged, timeChanged)
        ' status="failed"/"missing"は次回同期で自動的にリカバリさせる
        ' (§13「OneDriveオフライン…次回再試行」「フォルダ削除・リネーム…
        ' フォルダが復活すれば次回同期時に通常のkeep/replace判定に戻る」)。
        ' DiffDecision自体のシグネチャ(§7.8契約)は変えず、ResolveDecision
        ' (同じく純関数・§7.8)で呼び出し側からのみ上書きする。
        Dim curStatus As String: curStatus = ""
        If existsInManifest Then curStatus = mStatus(mi)
        decision = ResolveDecision(decision, existsInManifest, curStatus)

        ' 2026-07-28(レビュー L-11): 同期は silent:=True で取り込み結果を戻り値で
        ' 数える。従来は(a)同名衝突(E0504)が silent を無視してモーダルを出し、
        ' 誰も見ていない朝の同期がそこで止まる (b)失敗も「新規/更新」に足すので
        ' サマリが「新規12件」でも本棚は増えない、の2つの嘘があった。
        Dim st As String
        Dim errCd As String
        ' R18-6d: interactive:=Not silent(手動🔄/📁フォルダのみOCR事前確認対象。
        ' 無人同期は従来どおりFalse)。errCdでdeclined(「いいえ」)をfailedNと分離。
        Select Case decision
            Case "ingest"
                If modShelf.TotalChunks() >= capMax Then
                    cappedN = cappedN + 1
                Else
                    errCd = ""
                    st = modShelf.IngestFile(path, "self", True, errCd, Not silent)
                    If st = "done" Or st = "partial" Then
                        ingestedN = ingestedN + 1
                    ElseIf errCd = "declined" Then
                        declinedN = declinedN + 1
                    Else
                        failedN = failedN + 1
                    End If
                End If
            Case "replace"
                If modShelf.TotalChunks() >= capMax Then
                    cappedN = cappedN + 1
                Else
                    errCd = ""
                    st = modShelf.IngestFile(path, "self", True, errCd, Not silent)
                    If st = "done" Or st = "partial" Then
                        replacedN = replacedN + 1
                    ElseIf errCd = "declined" Then
                        declinedN = declinedN + 1
                    Else
                        failedN = failedN + 1
                    End If
                End If
            Case "keep"
                If mi >= 0 Then
                    If mStatus(mi) = "pending" Or mStatus(mi) = "partial" Then
                        resumeNeeded = True
                    End If
                End If
        End Select
    Next i

    ' 消失検知→削除(フォルダ配下由来の行のみ対象)。R15-FixA(FA-3i): 中断で
    ' 打ち切った同期はdiskDictが不完全なため「消えた」誤判定を避け、削除判定
    ' ごと見送る(列挙失敗時に消失判定しないのと同じ理由)。
    uiStep = "消失資料の削除判定"
    If cancelSkipN = 0 Then
        For i = 0 To mCount - 1
            If Not diskDict.Exists(LCase$(mNames(i))) Then
                modShelf.DeleteSource mNames(i)
                deletedN = deletedN + 1
            End If
        Next i
    End If

    uiStep = "未完了の埋め込み再開"
    Dim resumedCount As Long: resumedCount = 0
    If resumeNeeded Then
        On Error Resume Next
        modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDCE5) & " 未完了の埋め込みを再開しています…"
        On Error GoTo Failed
        resumedCount = modEmbed.EmbedPending()
    End If

    Dim summaryLine As String
    summaryLine = "新規" & ingestedN & "件・更新" & replacedN & "件・削除" & deletedN & "件"
    If failedN > 0 Then summaryLine = summaryLine & "・失敗" & failedN & "件"
    If declinedN > 0 Then summaryLine = summaryLine & "・見送り" & declinedN & "件"   ' R18-6d

    If resumeNeeded Then summaryLine = summaryLine & "・再開" & resumedCount & "件"
    If cappedN > 0 Then summaryLine = summaryLine & "・上限見送り" & cappedN & "件"
    ' R15-FixA(FA-3i): 中断は必ず言う(黙って件数が減るのが一番困る)。
    If cancelSkipN > 0 Then _
        summaryLine = summaryLine & "・中断のため" & cancelSkipN & "件を見送りました"

    If cappedN > 0 Then
        On Error Resume Next
        modLog.LogError "E0501", "modShelfSync.SyncNow", _
            "上限(" & capMax & "チャンク)到達により" & cappedN & "件を見送り"
        On Error GoTo 0
    End If

    ' 2026-07-28(レビュー I-11): 本棚カードの描き直しは同期の最後に1回だけ。
    ' IngestFile は silent のとき描かないので、ここで必ず1回呼ぶ。
    On Error Resume Next
    modUIShelf.RenderShelf
    On Error GoTo Failed

    If silent Then
        ' バックグラウンド同期(sync_on_open/自動同期)は対話ダイアログで
        ' フォーカスを奪わない。状態表示行+StatusBarのみで完了を知らせる
        ' (§8 UXレビュー: 定期的なMsgBoxが作業を中断する問題への対応)。
        On Error Resume Next
        modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDD04) & " 同期が完了しました(" & summaryLine & ")"
        On Error GoTo 0
    ElseIf cappedN = 0 Then
        ' ユーザーの追加アクションが不要な正常完了はMsgBoxを出さず、
        ' silent時と同じくSetStageのみで知らせる(§機能6・MsgBox削減)。
        On Error Resume Next
        modUIMain.SetStage ChrW(&H2705) & " 同期が完了しました(" & summaryLine & ")"
        ' R10-5: トーストで1回知らせる(kind:="success"がChrW(&H2705)を自前で
        ' 付けるため文言側では重ねない)。R10c(M4): バナーを先に閉じてから出す
        ' (逆順だと1.1秒2枚並び紛らわしい)。Finish:のHideProgressは異常系用(無害)。
        modUIMain.HideProgress
        modSkin.ShowToast "同期が完了しました(" & summaryLine & ")", "success"
        On Error GoTo 0
    Else
        Dim summary As String
        summary = "同期が完了しました。" & vbLf & _
            "新規: " & ingestedN & "件 / 更新: " & replacedN & "件 / 削除: " & deletedN & "件"
        If resumeNeeded Then
            summary = summary & vbLf & "未完了だった埋め込みを" & resumedCount & "件再開しました。"
        End If
        summary = summary & vbLf & vbLf & _
            "※本棚の上限(" & capMax & "チャンク)に達したため、" & cappedN & "件は取込を見送りました。" & vbLf & _
            "もっと入れたい場合は、configシートの shelf_max_chunks の数字を大きくしてから、" & _
            "もう一度「フォルダと同期」を押してください。"

        On Error Resume Next
        modUIMain.SetStage ""
        On Error GoTo 0

        MsgBox summary, vbInformation, modAppDef.APP_NAME
    End If

    modLog.LogUsage "sync", "", "ingest=" & ingestedN & " replace=" & replacedN & _
        " delete=" & deletedN & " resumed=" & resumedCount

    GoTo Finish

Failed:
    ' 同期中の想定外エラー。詳細をerr_logへ残し、mSyncRunningを確実に解除する
    ' (焼き付き防止)。同期はバックグラウンド処理なのでダイアログは出さず、
    ' 状態表示だけ元に戻す。
    Dim failNum As Long: failNum = Err.Number
    Dim failDesc As String: failDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailedCleanup1
FailedCleanup1:
    On Error Resume Next
    ' failNum=0 は「例外ではない打ち切り」でここへ直行してきた場合(列挙失敗
    ' など。その経路は自分で E0502 等を既に記録済み)。中身の無い E0801 を
    ' 重ねると、err_logで本当の原因が埋もれる(R11-D・A波発見事項2)。
    If failNum <> 0 Then
        modLog.LogError "E0801", "modShelfSync.SyncNow", _
            "[" & uiStep & "] err#" & failNum & ": " & failDesc, failNum
    End If
    modUIMain.SetStage ""
    ' Resume で抜けてハンドラ実行中の状態を解除する(On Error GoTo 0 では
    ' 解除されず、Finish: の後始末で起きたエラーが呼び出し元へ素通りする)。
    Resume Finish

Finish:
    mSyncRunning = False
    On Error Resume Next
    Application.EnableEvents = True   ' 抑止したイベントを必ず復帰
    modUIMain.HideProgress   ' R10-5: 正常/異常どちらの経路でも進捗バナーを必ず閉じる
    On Error GoTo 0
    ' P2P: 共有フォルダの感謝状(他者の✅由来)を回収して感謝EXPを加算する。
    ' shelf_folder未設定でもここは通る(P2P共有はnexus_share_pathで独立)。
    On Error Resume Next
    modP2P.CollectThanks silent
    modP2P.CollectNoiseVotes silent   ' 品質報告を集計し組織的除外(gexcl)を再計算
    On Error GoTo 0
    ' R15-3a: 変わったものがある同期だけ保存(RC9)。
    modShelfBatch.SaveCheckpoint ingestedN + replacedN + deletedN + resumedCount + orphanDone
    ' R15-FixA(FA-3ii): 中断の印を同期の外へ持ち越さない。取込側
    ' (AddFilesResult)が全ての出口で下ろしているのと同じ作法。下ろさないと、
    ' 同期の途中で押した中断が次に押した「資料を追加」まで生き残る。
    modShelfBatch.ResetCancel
End Sub

' ----------------------------------------------------------------------------
' IsBusy - 同期が走っているか(R7 B-2)。DoEventsで発火したクリックを画面遷移側
'   の入口で受け流すための判定。R15-1a: 失効判定は SyncNow と同じ。
' ----------------------------------------------------------------------------
Public Function IsBusy() As Boolean
    If Not mSyncRunning Then Exit Function
    On Error Resume Next
    IsBusy = Not modShelfBatch.GuardExpiredNow(mSyncRunningSince, GUARD_EXPIRY_MIN)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' ScheduleAutoSync - sync_interval_min>0ならApplication.OnTimeで次回予約(自己再帰)
' ----------------------------------------------------------------------------
Public Sub ScheduleAutoSync()
    ' 2026-08-01(R12-3-8): 二重予約の防止。画面遷移ごとの自己修復
    ' (modUI.EnsureSessionResources)からも呼ばれるようになったため、
    ' 「予約が生きているなら何もしない」を関数の側で保証する。予約が2本に
    ' なると CancelAutoSync が最後の1本しか解除できず、閉じたブックを
    ' Excelが勝手に開き直す事故(§9)の芽になる。
    ' AutoSyncTick は自分の発火時に mScheduled=False にしてから再予約するので、
    ' 自己再帰は従来どおり回る。
    If mScheduled Then Exit Sub

    Dim minutes As Long: minutes = modConfig.GetLong("sync_interval_min", 0)
    If minutes < 1 Then Exit Sub

    Dim nextTime As Date: nextTime = Now + TimeSerial(0, minutes, 0)

    On Error GoTo Fail
    ' Procedure は "'ブック名'!" 修飾(R12-3-8)。新旧2版が別名で併存したとき、
    ' 無修飾だと発火時の名前解決がどちらのブックへ向くか決まらない。解除側
    ' (CancelAutoSync)も同じ式で組み立てるので、文字列は必ず一致する。
    Application.OnTime EarliestTime:=nextTime, Procedure:=TickProcName()
    mNextRunTime = nextTime
    mScheduled = True
    ' 2026-07-28(レビュー L-23): 予約時刻を ui_state にも残す。モジュール変数
    ' だけだとVBAリセットで消え、Auto_Close が予約を解除できなくなる=「閉じた
    ' のに数分後にExcelが勝手に開き直す」事故(§9が防ぐと明記した事故)になる。
    ' R12-5-10: CStr(CDbl)の15桁精度欠落を避け、数値セルへ直書き(下のSaveSchedTime)。
    On Error Resume Next
    SaveSchedTime nextTime
    On Error GoTo Fail
    Exit Sub

Fail:
    ' OnTime予約自体の失敗は非致命(次回起動時のBoot経由の再予約に委ねる)。
    ' R12-5-8: Debug.Printは実機で誰も見ないため modLog.LogUsage(OERN下)へ。
    ' ハンドラ稼働中はResume Nextが効かないため、Resumeで抜けてから記録する
    ' (modMigrate等と同じ作法。2026-07-30実機err#462)。
    Dim schedErrDesc As String: schedErrDesc = Err.Description
    Resume FailCleanup
FailCleanup:
    On Error Resume Next
    modLog.LogUsage "autosync_schedule_failed", "", schedErrDesc
    On Error GoTo 0
    mScheduled = False
End Sub

' ----------------------------------------------------------------------------
' CancelAutoSync - 予約解除(必ずOn Error握り: 予約なしでも安全)
' ----------------------------------------------------------------------------
Public Sub CancelAutoSync()
    On Error Resume Next
    ' モジュール変数が生きていればそれで解除する。
    If mScheduled Then
        Application.OnTime EarliestTime:=mNextRunTime, Procedure:=TickProcName(), Schedule:=False
        Err.Clear
    End If

    ' 2026-07-28(レビュー L-23): モジュール変数が消えていても、
    ' ui_state に残した予約時刻で解除を試みる(R12-5-10: 数値セル直書き)。
    Dim schedFound As Boolean
    Dim schedT As Double: schedT = LoadSchedTime(schedFound)
    If schedFound Then
        Application.OnTime EarliestTime:=CDate(schedT), Procedure:=TickProcName(), Schedule:=False
        Err.Clear
        ClearSchedTime
    End If
    mScheduled = False
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' DiffDecision - 純関数(§7.8): manifest差分ロジックをテスト可能に切り出す。
'   existsInManifest=False              -> "ingest"
'   existsInManifest かつ サイズ/時刻変化あり -> "replace"
'   それ以外(既知・変化なし)              -> "keep"
' ----------------------------------------------------------------------------
Public Function DiffDecision(ByVal existsInManifest As Boolean, ByVal sizeChanged As Boolean, _
                             ByVal timeChanged As Boolean) As String
    If Not existsInManifest Then
        DiffDecision = "ingest"
    ElseIf sizeChanged Or timeChanged Then
        DiffDecision = "replace"
    Else
        DiffDecision = "keep"
    End If
End Function

' ----------------------------------------------------------------------------
' ResolveDecision - 純関数(§7.8): DiffDecisionの結果に、manifestの現在status
'   による上書きルールを適用する(SyncNowから切り出し・テスト可能にする)。
'   decision="keep" かつ status="failed"  -> "replace"(前回失敗を次回で再試行)
'   decision="keep" かつ status="missing" -> "replace"(フォルダ復活時に復帰。
'     Wave4修正: 従来はmissing行がkeepのまま固定され、フォルダが戻っても
'     カードが「削除待ち」表示のまま・埋め込み再開も走らない不具合があった)
'   decision="keep" かつ status="failed_permanent" -> "keep" のまま
'     (2026-08-01 R12-3-3。連続3回失敗したファイルは自動同期の対象から外す。
'      毎回の同期時間と err_log を食い続けるのを止めるため。ファイル自体が
'      更新されれば DiffDecision が "replace" を返すので自動で再試行になり、
'      利用者が「資料を追加」で選び直せば modShelfStore.ResetFailCountForPath
'      が status を "failed" に戻して通常の再試行経路へ復帰する)
'   それ以外はdecisionをそのまま返す。
' ----------------------------------------------------------------------------
Public Function ResolveDecision(ByVal decision As String, ByVal existsInManifest As Boolean, _
                                ByVal currentStatus As String) As String
    ResolveDecision = decision
    If decision = "keep" And existsInManifest Then
        If currentStatus = "failed" Or currentStatus = "missing" Then
            ResolveDecision = "replace"
        End If
    End If
End Function

' ----------------------------------------------------------------------------
' AutoSyncTick - OnTimeコールバック本体。SyncNow実行後に次回分を再予約する。
'   Public必須(Wave4修正): Application.OnTimeのProcedure引数はApplication.Run
'   と同じ遅延バインドであり、Private Subは解決できない(発火時に「マクロを
'   実行できません」という未処理エラーになり、自動同期が機能しなくなる)。
'   バックグラウンド呼び出しのためSyncNowはsilent:=Trueで呼ぶ(完了ダイアログ
'   でユーザー操作を中断しない)。
' ----------------------------------------------------------------------------
Public Sub AutoSyncTick()
    mScheduled = False
    On Error Resume Next
    SyncNow silent:=True
    On Error GoTo 0
    ScheduleAutoSync
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

' OnTime へ渡すプロシージャ名。予約側と解除側で必ず同じ文字列にするため、
' 組み立てはここ1箇所に置く(R12-3-8)。ブック名は実行時に決まる
' (配布更新で「MyBookshelf (1).xlsm」等へ変わり得る)ので定数にはできない。
Private Function TickProcName() As String
    TickProcName = "'" & ThisWorkbook.Name & "'!AutoSyncTick"
End Function

' SCHED_KEY 保存(R12-5-10)。modState経由(CStr(CDbl))は最大15桁の精度欠落で
' EarliestTime完全一致解除が壊れ得るため、ui_stateのB列へ数値を直書きする
' (インストーラのE1=CDbl直書きと同型)。この用途だけ直接シート操作する。
Private Function FindStateRow(ByVal ws As Worksheet, ByVal keyName As String) As Long
    On Error Resume Next
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(CStr(ws.Cells(i, 1).Value), keyName, vbTextCompare) = 0 Then
            FindStateRow = i
            Exit Function
        End If
    Next i
    On Error GoTo 0
End Function

Private Sub SaveSchedTime(ByVal t As Date)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_UISTATE)
    If ws Is Nothing Then Exit Sub
    Dim r As Long: r = FindStateRow(ws, SCHED_KEY)
    If r = 0 Then
        r = ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 1
        If r < 1 Then r = 1
        ws.Cells(r, 1).Value = SCHED_KEY
    End If
    ws.Cells(r, 2).Value = CDbl(t)
    On Error GoTo 0
End Sub

Private Function LoadSchedTime(ByRef found As Boolean) As Double
    found = False
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_UISTATE)
    If ws Is Nothing Then Exit Function
    Dim r As Long: r = FindStateRow(ws, SCHED_KEY)
    If r = 0 Then Exit Function
    Dim v As Variant: v = ws.Cells(r, 2).Value
    If IsNumeric(v) Then
        LoadSchedTime = CDbl(v)
        found = True
    End If
    On Error GoTo 0
End Function

Private Sub ClearSchedTime()
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_UISTATE)
    If ws Is Nothing Then Exit Sub
    Dim r As Long: r = FindStateRow(ws, SCHED_KEY)
    If r > 0 Then ws.Cells(r, 2).Value = ""
    On Error GoTo 0
End Sub

' フォルダの存在確認。Dir(path, vbDirectory)は末尾に区切り文字が付いていると
' 正しく判定できないため、まず末尾のスラッシュ・バックスラッシュを取り除く。
Private Function FolderExists(ByVal path As String) As Boolean
    Dim p As String: p = path
    Do While Len(p) > 0 And (Right$(p, 1) = "\" Or Right$(p, 1) = "/")
        p = Left$(p, Len(p) - 1)
    Loop
    If LenB(p) = 0 Then Exit Function

    On Error Resume Next
    Dim nm As String: nm = Dir$(p, vbDirectory)
    On Error GoTo 0
    FolderExists = (LenB(nm) > 0)
End Function


