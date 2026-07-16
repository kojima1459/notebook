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
'     だけを比較対象スコープとする(LoadManifestScope)。
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
'   問題があった。手動🔄ボタン(modUIShelf.OnSyncNow)・PickShelfFolder直後は
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

    ' 2026-07-16 恒久対策: 同期処理のどこで実行時エラーが起きても、必ず
    ' Finish(mSyncRunningの解除)へ合流させる。従来は本体を覆うエラー
    ' ハンドラが無く、フォルダ走査やmanifest突合で例外が出るとmSyncRunning=True
    ' のまま抜けて以降の同期が永久に走らなくなる危険があった(IngestFileの
    ' 焼き付きと同種の予防)。
    Dim uiStep As String
    On Error GoTo Failed

    uiStep = "状態表示の更新"
    On Error Resume Next
    modUIMain.SetStage "🔄 同期を確認しています…"
    On Error GoTo Failed

    uiStep = "本棚フォルダ設定の確認"
    Dim folder As String: folder = Trim$(modConfig.GetString("shelf_folder", ""))
    If LenB(folder) = 0 Then
        ' 初回起動直後(shelf_folder未設定)は「まだ何も設定していない正常な
        ' 状態」であって障害ではない。バックグラウンド同期(silent)のときは
        ' err_logを汚さない(実機で毎起動E0502が3件ずつ記録され、本当の
        ' 障害が埋もれる問題への対応)。手動🔄時のみ丁寧に案内する。
        If Not silent Then
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
                   "「📁 フォルダを選ぶ」から選んでください。" & vbLf & _
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
        MarkFolderScopeMissing folder
        On Error Resume Next
        modUIMain.SetStage ""
        On Error GoTo Failed
        GoTo Finish
    End If

    Dim folderNorm As String: folderNorm = EnsureTrailingSlash(folder)

    uiStep = "フォルダ内ファイルの一覧取得"
    Dim diskNames() As String, diskCount As Long
    EnumFolderFiles folderNorm, diskNames, diskCount

    uiStep = "資料台帳スコープの読込"
    Dim wsM As Worksheet: Set wsM = GetSheet(modAppDef.SH_MANIFEST)
    Dim mPaths() As String, mNames() As String, mModified() As Date
    Dim mSize() As Double, mStatus() As String, mCount As Long
    LoadManifestScope wsM, folderNorm, mPaths, mNames, mModified, mSize, mStatus, mCount

    Dim diskDict As Object: Set diskDict = CreateObject("Scripting.Dictionary")
    Dim ingestedN As Long, replacedN As Long, deletedN As Long
    Dim resumeNeeded As Boolean: resumeNeeded = False

    uiStep = "新規・更新の差分判定"
    Dim i As Long
    For i = 0 To diskCount - 1
        If Not diskDict.Exists(LCase$(diskNames(i))) Then diskDict.Add LCase$(diskNames(i)), True

        On Error Resume Next
        modUIMain.SetStage "🔄 同期中 " & (i + 1) & "/" & diskCount & " …"
        On Error GoTo Failed

        Dim path As String: path = folderNorm & diskNames(i)
        Dim mi As Long: mi = FindManifestIndexByName(mNames, mCount, diskNames(i))
        Dim existsInManifest As Boolean: existsInManifest = (mi >= 0)

        Dim sizeChanged As Boolean: sizeChanged = False
        Dim timeChanged As Boolean: timeChanged = False
        If existsInManifest Then
            Dim curSize As Double: curSize = SafeFileLen(path)
            Dim curMod As Date: curMod = SafeFileDateTime(path)
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

        Select Case decision
            Case "ingest"
                modShelf.IngestFile path, "self"
                ingestedN = ingestedN + 1
            Case "replace"
                modShelf.IngestFile path, "self"
                replacedN = replacedN + 1
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
        modUIMain.SetStage "📥 未完了の埋め込みを再開しています…"
        On Error GoTo Failed
        resumedCount = modEmbed.EmbedPending()
    End If

    Dim summaryLine As String
    summaryLine = "新規" & ingestedN & "件・更新" & replacedN & "件・削除" & deletedN & "件"
    If resumeNeeded Then summaryLine = summaryLine & "・再開" & resumedCount & "件"

    If silent Then
        ' バックグラウンド同期(sync_on_open/自動同期)は対話ダイアログで
        ' フォーカスを奪わない。状態表示行+StatusBarのみで完了を知らせる
        ' (§8 UXレビュー: 定期的なMsgBoxが作業を中断する問題への対応)。
        On Error Resume Next
        modUIMain.SetStage "🔄 同期が完了しました(" & summaryLine & ")"
        On Error GoTo 0
    Else
        Dim summary As String
        summary = "同期が完了しました。" & vbLf & _
            "新規: " & ingestedN & "件 / 更新: " & replacedN & "件 / 削除: " & deletedN & "件"
        If resumeNeeded Then
            summary = summary & vbLf & "未完了だった埋め込みを" & resumedCount & "件再開しました。"
        End If

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
    On Error Resume Next
    modLog.LogError "E0801", "modShelfSync.SyncNow", "[" & uiStep & "] err#" & failNum & ": " & failDesc
    modUIMain.SetStage ""
    On Error GoTo 0

Finish:
    mSyncRunning = False
End Sub

' ----------------------------------------------------------------------------
' ScheduleAutoSync - sync_interval_min>0ならApplication.OnTimeで次回予約(自己再帰)
' ----------------------------------------------------------------------------
Public Sub ScheduleAutoSync()
    Dim minutes As Long: minutes = modConfig.GetLong("sync_interval_min", 0)
    If minutes < 1 Then Exit Sub

    Dim nextTime As Date: nextTime = Now + TimeSerial(0, minutes, 0)

    On Error GoTo Fail
    Application.OnTime EarliestTime:=nextTime, Procedure:="AutoSyncTick"
    mNextRunTime = nextTime
    mScheduled = True
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
    If Not mScheduled Then Exit Sub
    On Error Resume Next
    Application.OnTime EarliestTime:=mNextRunTime, Procedure:="AutoSyncTick", Schedule:=False
    On Error GoTo 0
    mScheduled = False
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

Private Function EnsureTrailingSlash(ByVal path As String) As String
    Dim p As String: p = path
    If Len(p) = 0 Then Exit Function
    If Right$(p, 1) <> "\" And Right$(p, 1) <> "/" Then p = p & "\"
    EnsureTrailingSlash = p
End Function

' pathからファイル名部分を除いたディレクトリ部分(末尾区切り文字含む)を返す。
Private Function DirOf(ByVal path As String) As String
    Dim fn As String: fn = modUtil.FileNameOf(path)
    If Len(fn) >= Len(path) Then
        DirOf = ""
    Else
        DirOf = Left$(path, Len(path) - Len(fn))
    End If
End Function

' folderNorm配下(第1階層のみ・対応拡張子のみ)のファイル名一覧を返す。
' サブフォルダはDir()の既定動作(vbDirectory属性を指定しない)により
' 自動的に除外される。
Private Sub EnumFolderFiles(ByVal folderNorm As String, ByRef fileNames() As String, ByRef fileCount As Long)
    Dim names() As String: ReDim names(0 To 63)
    Dim cnt As Long: cnt = 0

    On Error Resume Next
    Dim nm As String: nm = Dir$(folderNorm & "*.*")
    On Error GoTo 0

    Do While LenB(nm) > 0
        If IsSupportedExtLocal(nm) Then
            If cnt > UBound(names) Then ReDim Preserve names(0 To (UBound(names) + 1) * 2 - 1)
            names(cnt) = nm
            cnt = cnt + 1
        End If
        nm = Dir$()
    Loop

    If cnt = 0 Then
        ReDim fileNames(0 To 0)
    Else
        ReDim Preserve names(0 To cnt - 1)
        fileNames = names
    End If
    fileCount = cnt
End Sub

Private Function IsSupportedExtLocal(ByVal fileName As String) As Boolean
    Dim ext As String: ext = modUtil.ExtOf(fileName)
    If LenB(ext) = 0 Then Exit Function
    Dim supported As String: supported = modExtractor.SupportedExts()
    IsSupportedExtLocal = (InStr(1, "," & supported & ",", "," & ext & ",", vbTextCompare) > 0)
End Function

' my_manifestのうち origin=self かつ folderNorm配下の行だけを抜き出す
' (SyncNowの比較・消失検知スコープ)。一括読み取り→配列フィルタで
' ループ内Range直接アクセスを避ける(§12)。
Private Sub LoadManifestScope(ByVal wsM As Worksheet, ByVal folderNorm As String, _
        ByRef outPaths() As String, ByRef outNames() As String, ByRef outModified() As Date, _
        ByRef outSize() As Double, ByRef outStatus() As String, ByRef outCount As Long)
    outCount = 0
    ReDim outPaths(0 To 0)
    ReDim outNames(0 To 0)
    ReDim outModified(0 To 0)
    ReDim outSize(0 To 0)
    ReDim outStatus(0 To 0)
    If wsM Is Nothing Then Exit Sub

    Dim lastM As Long: lastM = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
    If lastM < 2 Then Exit Sub

    Dim arr As Variant: arr = wsM.Range(wsM.Cells(2, 1), wsM.Cells(lastM, 9)).Value
    Dim n As Long: n = UBound(arr, 1) - LBound(arr, 1) + 1

    Dim tmpPaths() As String: ReDim tmpPaths(0 To n - 1)
    Dim tmpNames() As String: ReDim tmpNames(0 To n - 1)
    Dim tmpModified() As Date: ReDim tmpModified(0 To n - 1)
    Dim tmpSize() As Double: ReDim tmpSize(0 To n - 1)
    Dim tmpStatus() As String: ReDim tmpStatus(0 To n - 1)
    Dim cnt As Long: cnt = 0

    Dim folderLower As String: folderLower = LCase$(folderNorm)

    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(CStr(arr(i, 9)), "self", vbTextCompare) = 0 Then
            Dim p As String: p = CStr(arr(i, 1))
            If LCase$(EnsureTrailingSlash(DirOf(p))) = folderLower Then
                tmpPaths(cnt) = p
                tmpNames(cnt) = CStr(arr(i, 2))

                On Error Resume Next
                tmpModified(cnt) = CDate(arr(i, 3))
                On Error GoTo 0

                If IsNumeric(arr(i, 4)) Then tmpSize(cnt) = CDbl(arr(i, 4))
                tmpStatus(cnt) = CStr(arr(i, 6))
                cnt = cnt + 1
            End If
        End If
    Next i

    If cnt > 0 Then
        ReDim Preserve tmpPaths(0 To cnt - 1)
        ReDim Preserve tmpNames(0 To cnt - 1)
        ReDim Preserve tmpModified(0 To cnt - 1)
        ReDim Preserve tmpSize(0 To cnt - 1)
        ReDim Preserve tmpStatus(0 To cnt - 1)
        outPaths = tmpPaths
        outNames = tmpNames
        outModified = tmpModified
        outSize = tmpSize
        outStatus = tmpStatus
    End If
    outCount = cnt
End Sub

Private Function FindManifestIndexByName(names() As String, ByVal cnt As Long, ByVal fileName As String) As Long
    FindManifestIndexByName = -1
    Dim i As Long
    For i = 0 To cnt - 1
        If StrComp(names(i), fileName, vbTextCompare) = 0 Then
            FindManifestIndexByName = i
            Exit Function
        End If
    Next i
End Function

' shelf_folder自体が見つからない(削除・リネーム・OneDriveオフライン等)ときに、
' そのフォルダ配下由来のmanifest行だけを status="missing" に更新する
' (実削除はしない。§13「フォルダ削除・リネーム(E0502+missing)」)。
Private Sub MarkFolderScopeMissing(ByVal folder As String)
    Dim wsM As Worksheet: Set wsM = GetSheet(modAppDef.SH_MANIFEST)
    If wsM Is Nothing Then Exit Sub

    Dim lastM As Long: lastM = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
    If lastM < 2 Then Exit Sub

    Dim folderLower As String: folderLower = LCase$(EnsureTrailingSlash(folder))

    Dim arr As Variant: arr = wsM.Range(wsM.Cells(2, 1), wsM.Cells(lastM, 9)).Value
    Dim changed As Boolean: changed = False

    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(CStr(arr(i, 9)), "self", vbTextCompare) = 0 Then
            Dim p As String: p = CStr(arr(i, 1))
            If LCase$(EnsureTrailingSlash(DirOf(p))) = folderLower Then
                If CStr(arr(i, 6)) <> "missing" Then
                    arr(i, 6) = "missing"
                    changed = True
                End If
            End If
        End If
    Next i

    If changed Then
        wsM.Range(wsM.Cells(2, 1), wsM.Cells(lastM, 9)).Value = arr
    End If
End Sub

Private Function SafeFileLen(ByVal path As String) As Double
    On Error GoTo Fail
    SafeFileLen = CDbl(FileLen(path))
    Exit Function
Fail:
    SafeFileLen = 0
End Function

Private Function SafeFileDateTime(ByVal path As String) As Date
    On Error GoTo Fail
    SafeFileDateTime = FileDateTime(path)
    Exit Function
Fail:
    SafeFileDateTime = Now
End Function
