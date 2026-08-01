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
' 設計判断:
'   ・同期スコープの限定: my_manifest には「＋資料を追加」ダイアログから
'     shelf_folder外のファイルを取り込んだ行も混在しうる。SyncNowが
'     「消失→削除」を適用してよいのは shelf_folder 配下由来の行だけであり、
'     フォルダ外から個別に追加した資料をSyncNowが誤って削除しないよう、
'     manifestの file_path の親ディレクトリが shelf_folder と一致する行
'     だけを比較対象スコープとする(modShelfScan.LoadManifestScope)。
'   ・拡張子フィルタ: フォルダ内には Thumbs.db や .ini 等、対応外の
'     ファイルが混在しうる。AddFilesViaDialog はFileDialogのFiltersで
'     絞り込んでいるが、Dir()走査にはその仕組みが無いため、
'     modExtractor.SupportedExts() に含まれない拡張子は同期の対象外として
'     最初から除外する(取込失敗の空騒ぎ・manifestの無駄な失敗行を防ぐ)。
'   ・差分判定(新規/変更/据え置き)は純関数 DiffDecision に切り出し、
'     モジュールのExcel依存部分と分離してテスト可能にする(MASTER_SPEC §7.8)。
'   ・フォルダそのものが見つからない場合(削除・リネーム・OneDriveオフライン
'     で丸ごと見えない等)は、個々のファイルの消失と区別する。誤って
'     全資料をDeleteSourceしてしまうと復旧不能なため、E0502を表示した上で
'     対象スコープのmanifest行を status="missing" に更新するだけに留め、
'     実削除はしない(§13「フォルダ削除・リネーム(E0502+missing)」)。
'     フォルダが復活すれば次回同期時に通常のkeep/replace判定に戻る。
'   ・再入防止: mSyncRunning は「手動🔄同期ボタン連打」や「自動同期タイマー
'     発火中に手動ボタンが押される」ケースの多重実行を防ぐ。
'     【Wave4訂正】以前はここで「Excel/VBAはシングルスレッドなので、1つの
'     マクロ実行中は他のボタンクリックはExcelが自動的に遅延させる」と
'     説明していたが、これは不正確だった: modEmbed.EmbedPendingは
'     ベクトル化ループの中でDoEvents/SleepMs(DoEvents呼び出し)を使って
'     ESC中断とUI応答性を確保しており、DoEvents実行中はメッセージキュー上の
'     Shapeクリック(OnAction)がその場で再入的に発火しうる(「現在のマクロが
'     終わるまで待つ」わけではない)。したがって「取込/埋め込み中にユーザーが
'     🔄同期ボタンを押す」という再入は実際に起こり得る。この場合に
'     my_knowledge/my_vectorsの行削除・圧縮(DeleteSource等)と、
'     EmbedPendingが保持する行インデックスが食い違って誤った行へ書き込む
'     事故を防ぐため、modEmbed側でchunk_id起点に書込み先を都度再解決する
'     自己防衛(re-resolve)を入れている(modEmbed.bas冒頭コメント参照)。
'     modShelfSync自体に新しいPublicの相互排他フラグを追加する設計変更は
'     影響範囲が大きいため見送り、実害(誤った行への書込み)を確実に断つ
'     modEmbed側の対策で十分と判断した。
'   ・OnTime予約: 予約時刻(mNextRunTime)をモジュール変数に保持し、
'     CancelAutoSyncは同時刻を指定して解除する(Excel仕様・§12)。
'     コールバック先(AutoSyncTick)はSyncNow実行後に自分自身を再度
'     ScheduleAutoSyncする「自己再帰」で次回分を予約する。
'     【Wave4修正】AutoSyncTickは当初「§7.2の公開契約に無い名前だから」
'     という理由でPrivateにしていたが、これはVBAの実際の挙動と矛盾する
'     誤った判断だった: Application.OnTimeのProcedure引数はApplication.Run
'     と同じ遅延バインド(文字列からの実行時解決)であり、Private Subは
'     解決できない(予約自体はOn Errorに掛からず成立してしまうが、発火時に
'     「マクロ'AutoSyncTick'を実行できません」という素のランタイムエラーが
'     出て自動同期が永久に機能しなくなる)。そのためAutoSyncTickはPublicに
'     変更し、CONTRACT(tools/vba_lint.py)・MASTER_SPEC §7.2の両方を
'     「OnTimeコールバックとして公開が必須」という契約に合わせて更新した
'     (契約側の欠陥が原因のケースとしてMASTER_SPEC本文も修正)。
'   ・進捗実況は他のingest層モジュール(modEmbed/modEnrich)と同じ作法で
'     modUIMain.SetStage を1行スコープの On Error Resume Next で呼ぶ
'     (modUIMain未実装/実行時エラーでも同期処理自体は止めない)。
'   ・FileDialog/EnableCancelKey等のOffice名前付き定数は使わずリテラル値
'     を使う(modShelf.bas/modLog.bas と同じ、LibreOffice互換のための
'     V2以来の慣習)。
' ============================================================================

Private mSyncRunning As Boolean     ' 再入防止
' ガードの自己回復(2026-07-16): modShelf.mIngestingと同じ理由(強制停止で
' フラグが焼き付くと以降の同期が全て無言スキップになる)のタイムスタンプ。
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
'   【Wave4追加】modBoot.Boot(sync_on_open)とAutoSyncTick(定期自動同期)は
'   ユーザー操作なしで走るバックグラウンド処理のため、対話ダイアログで
'   フォーカスを奪うべきではない(§8 UXレビュー指摘)。特に配布直後は
'   shelf_folderが未設定("")のままsync_on_open=TRUEで起動するため、名前を
'   入力した直後に「フォルダが見つかりません」という警告が必ず出てしまう
'   問題があった。手動🔄ボタン(modKnowledge.OnSync)・PickShelfFolder直後は
'   従来どおりsilent=Falseでダイアログ表示する(ユーザー操作への応答なので
'   フィードバックがあった方がよい)。
' ----------------------------------------------------------------------------
Public Sub SyncNow(Optional ByVal silent As Boolean = False)
    ' 再入防止(手動連打・自動同期との重複)。焼き付いたガードは自動解除。
    If mSyncRunning Then
        If DateDiff("n", mSyncRunningSince, Now) >= GUARD_EXPIRY_MIN Then
            On Error Resume Next
            modLog.LogUsage "guard_recover", "sync", _
                "前回の同期ガードが" & GUARD_EXPIRY_MIN & "分以上残留していたため自動解除"
            On Error GoTo 0
            mSyncRunning = False
        End If
    End If
    If mSyncRunning Then Exit Sub
    mSyncRunning = True
    mSyncRunningSince = Now

    ' 自動同期(OnTime)/手動同期中の大量シート書換え中にイベント連鎖が起きない
    ' よう抑止。Finishで必ずTrueへ戻す(死の連鎖防止)。
    On Error Resume Next
    Application.EnableEvents = False
    On Error GoTo 0

    ' 2026-07-16 恒久対策: 同期処理のどこで実行時エラーが起きても、必ず
    ' Finish(mSyncRunningの解除)へ合流させる。従来は本体を覆うエラー
    ' ハンドラが無く、フォルダ走査やmanifest突合で例外が出るとmSyncRunning=True
    ' のまま抜けて以降の同期が永久に走らなくなる危険があった(IngestFileの
    ' 焼き付きと同種の予防)。
    Dim uiStep As String
    On Error GoTo Failed

    uiStep = "状態表示の更新"
    On Error Resume Next
    modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDD04) & " 同期を確認しています…"
    On Error GoTo Failed

    uiStep = "本棚フォルダ設定の確認"
    Dim folder As String: folder = Trim$(modConfig.GetString("shelf_folder", ""))
    If LenB(folder) = 0 Then
        ' 初回起動直後(shelf_folder未設定)は「まだ何も設定していない正常な
        ' 状態」であって障害ではない。バックグラウンド同期(silent)のときは
        ' err_logを汚さない(実機で毎起動E0502が3件ずつ記録され、本当の
        ' 障害が埋もれる問題への対応)。手動🔄時のみ丁寧に案内する。
        '
        ' 【2026-07-16 袋小路の解消】「＋資料を追加」で入れた資料がpartial
        ' (ベクトル化未完了)のとき、カードのメモは「🔄フォルダと同期を押すと
        ' 続きから再開します」と案内するのに、フォルダ未設定だとE0502で
        ' 弾かれて再開手段が無かった。フォルダが未設定でも、未完了の
        ' ベクトル化(EmbedPending)だけは実行して「続きから再開」を成立させる。
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
        ' 2026-07-31 R11-D(A波発見事項2): ここは【エラーが起きていない】
        ' 正常な打ち切り経路なので、Failed: へ飛ばしてはいけない。Failed: の
        ' Resume はエラー未発生の状態で実行されると err#20(Resume without
        ' error)を起こし、それが On Error GoTo Failed に捕まって2周目で
        ' 初めて成立する。その2周目で failNum/failDesc が 20 /「Resume without
        ' error」に上書きされるため、err_log には上のE0502ではなく自作自演の
        ' err#20 だけが残っていた。後始末(FailedCleanup1)へ直行させる。
        ' 挙動は不変(記録内容が実態どおりになるだけ)。modPublishUI.OnPublish
        ' で同型を是正したのと同じ構造。
        GoTo FailedCleanup1
    End If

    uiStep = "資料台帳スコープの読込"
    Dim wsM As Worksheet: Set wsM = GetSheet(modAppDef.SH_MANIFEST)
    Dim mPaths() As String, mNames() As String, mModified() As Date
    Dim mSize() As Double, mStatus() As String, mCount As Long
    modShelfScan.LoadManifestScope wsM, folderNorm, mPaths, mNames, mModified, mSize, mStatus, mCount

    Dim diskDict As Object: Set diskDict = CreateObject("Scripting.Dictionary")
    Dim ingestedN As Long, replacedN As Long, deletedN As Long, failedN As Long
    Dim resumeNeeded As Boolean: resumeNeeded = False

    ' 本棚上限(2026-07-16 実機対応): 大容量フォルダを同期すると、上限到達後の
    ' 残り全ファイルがIngestFile内のE0501ダイアログを1件ずつ出し続ける
    ' 「ダイアログ地獄」になっていた(実機で35連発)。上限はこのループ側で
    ' 先に判定し、到達後は取込を静かに見送って件数だけ数え、最後に1回だけ
    ' まとめて案内する。
    Dim capMax As Long: capMax = modConfig.GetLong("shelf_max_chunks", modAppDef.DEFAULT_SHELF_MAX_CHUNKS)
    If capMax < 1 Then capMax = modAppDef.DEFAULT_SHELF_MAX_CHUNKS
    Dim cappedN As Long: cappedN = 0

    uiStep = "新規・更新の差分判定"
    Dim i As Long
    For i = 0 To diskCount - 1
        If Not diskDict.Exists(LCase$(diskNames(i))) Then diskDict.Add LCase$(diskNames(i)), True

        ' R10-5: SetStageはShowProgress内部で呼ぶので二重呼び出しにしない。
        On Error Resume Next
        modUIMain.ShowProgress "" & ChrW(&HD83D) & ChrW(&HDD04) & " 同期中 " & (i + 1) & "/" & diskCount & " …"
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

        ' 2026-07-28(レビュー L-11): 同期は silent:=True で取り込み、
        ' 結果を戻り値で数える。従来は
        '   ・同名衝突(E0504)が silent を無視してモーダルを出す
        '     → 誰も見ていない朝の同期がそこで止まる
        '   ・失敗も「新規/更新」の件数に足す
        '     → サマリが「新規12件」と言うのに本棚は増えていない
        ' という2つの嘘があった。数えるのは実際に入ったものだけにする。
        Dim st As String
        Select Case decision
            Case "ingest"
                If modShelf.TotalChunks() >= capMax Then
                    cappedN = cappedN + 1
                Else
                    st = modShelf.IngestFile(path, "self", True)
                    If st = "done" Or st = "partial" Then
                        ingestedN = ingestedN + 1
                    Else
                        failedN = failedN + 1
                    End If
                End If
            Case "replace"
                If modShelf.TotalChunks() >= capMax Then
                    cappedN = cappedN + 1
                Else
                    st = modShelf.IngestFile(path, "self", True)
                    If st = "done" Or st = "partial" Then
                        replacedN = replacedN + 1
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

    ' 消失検知→削除(このフォルダ配下由来の行のみが対象スコープ)
    uiStep = "消失資料の削除判定"
    For i = 0 To mCount - 1
        If Not diskDict.Exists(LCase$(mNames(i))) Then
            modShelf.DeleteSource mNames(i)
            deletedN = deletedN + 1
        End If
    Next i

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

    If resumeNeeded Then summaryLine = summaryLine & "・再開" & resumedCount & "件"
    If cappedN > 0 Then summaryLine = summaryLine & "・上限見送り" & cappedN & "件"

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
        ' R10-5: 手動同期(silent=False)は現行何も出ないため、トーストで1回知らせる。
        ' kind:="success"はShowToast側が自前でChrW(&H2705)を付けるため、文言に
        ' 重ねて絵文字を書かない。
        ' R10c(M4): 進捗バナーを先に閉じてからトーストを出す(modShelfと同順)。
        ' 逆順だと1.1秒のあいだ2枚がならび、どちらが今の状態か分からなくなる。
        ' Finish: の HideProgress は異常系用に残す(二重呼び出しは無害)。
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
            "もう一度「" & ChrW(&HD83D) & ChrW(&HDD04) & " フォルダと同期」を押してください。"

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
End Sub

' ----------------------------------------------------------------------------
' IsBusy - 同期が走っているか(2026-07-31 R7 B-2)。
'   ファイルループの DoEvents で発火したクリックを、画面遷移側の入口で
'   受け流すための判定。再入ガード mSyncRunning をそのまま公開する。
'   焼き付いたガードで全ボタンが無反応になるのを避けるため、SyncNow と同じ
'   期限(GUARD_EXPIRY_MIN)を過ぎたものは busy とみなさない。
' ----------------------------------------------------------------------------
Public Function IsBusy() As Boolean
    If Not mSyncRunning Then Exit Function
    On Error Resume Next
    IsBusy = (DateDiff("n", mSyncRunningSince, Now) < GUARD_EXPIRY_MIN)
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
    ' 2026-07-28(レビュー L-23): 予約時刻を ui_state にも残す。
    ' モジュール変数だけに持っていると、VBAリセット(エラー中断・VBEでの
    ' リセット)で消え、Auto_Close が予約を解除できなくなる。
    ' 解除できない OnTime は「閉じたのに数分後にExcelが勝手に開き直す」
    ' という事故になる(§9 が防ぐと明記している事故そのもの)。
    On Error Resume Next
    modState.SaveState SCHED_KEY, CStr(CDbl(nextTime))
    On Error GoTo Fail
    Exit Sub

Fail:
    ' OnTime予約自体の失敗は非致命(次回起動時のBoot経由の再予約に委ねる)。
    ' ログ書き込み失敗と同じ扱いでDebug.Printのみに留める(R5の「ログで
    ' 死なない」運用と同様、背景処理のスケジューリング失敗を理由にアプリ
    ' 全体を止めたくないための設計判断)。
    Debug.Print "[modShelfSync.ScheduleAutoSync:予約失敗] " & Err.Description
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
    ' ui_state に残した予約時刻で解除を試みる。
    Dim s As String: s = modState.LoadState(SCHED_KEY, "")
    If LenB(s) > 0 Then
        If IsNumeric(s) Then
            Application.OnTime EarliestTime:=CDate(CDbl(s)), Procedure:=TickProcName(), Schedule:=False
            Err.Clear
        End If
        modState.SaveState SCHED_KEY, ""
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


