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
'     発火中に手動ボタンが押される」ケースの多重実行を防ぐ。なお
'     「自動同期中にユーザーが質問実行」(§13)については、Excel/VBAの
'     実行モデルがそもそもシングルスレッドであり、1つのマクロ実行中は
'     OnTime予約や他のボタンクリックはExcelが自動的に現在のマクロの終了
'     まで遅延させるため、modShelfSync側で追加の相互排他は不要と判断した
'     (modAsk側との明示的なフラグ共有はMASTER_SPEC契約に無く、勝手な
'     Public APIの追加はcontract closedのため許されない)。
'   ・OnTime予約: 予約時刻(mNextRunTime)をモジュール変数に保持し、
'     CancelAutoSyncは同時刻を指定して解除する(Excel仕様・§12)。
'     コールバック先(AutoSyncTick)はSyncNow実行後に自分自身を再度
'     ScheduleAutoSyncする「自己再帰」で次回分を予約する。
'     AutoSyncTickはMASTER_SPEC §7.2の公開契約(closed)に無い名前のため
'     Publicにはしない(Private)。Application.OnTimeのProcedure引数は
'     モジュール修飾なしの裸の手続き名("AutoSyncTick")を渡す
'     (同一VBAプロジェクト内でPrivate Subも名前解決して呼び出せるVBAの
'     挙動を利用。"modShelfSync.AutoSyncTick"のように修飾すると静的Lintの
'     モジュール間参照チェックが「外部から呼べるPublic APIの呼び出し」と
'     誤認してERRORになるため、あえて非修飾名にしている)。
'   ・進捗実況は他のingest層モジュール(modEmbed/modEnrich)と同じ作法で
'     modUIMain.SetStage を1行スコープの On Error Resume Next で呼ぶ
'     (modUIMain未実装/実行時エラーでも同期処理自体は止めない)。
'   ・FileDialog/EnableCancelKey等のOffice名前付き定数は使わずリテラル値
'     を使う(modShelf.bas/modLog.bas と同じ、LibreOffice互換のための
'     V2以来の慣習)。
' ============================================================================

Private mSyncRunning As Boolean     ' 再入防止
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
' ----------------------------------------------------------------------------
Public Sub SyncNow()
    If mSyncRunning Then Exit Sub   ' 再入防止(手動連打・自動同期との重複)
    mSyncRunning = True

    On Error Resume Next
    modUIMain.SetStage "🔄 同期を確認しています…"
    On Error GoTo 0

    Dim folder As String: folder = Trim$(modConfig.GetString("shelf_folder", ""))
    If LenB(folder) = 0 Then
        modLog.ShowError "E0502", "modShelfSync.SyncNow", "shelf_folderが未設定です"
        GoTo Finish
    End If

    If Not FolderExists(folder) Then
        modLog.ShowError "E0502", "modShelfSync.SyncNow", folder
        MarkFolderScopeMissing folder
        GoTo Finish
    End If

    Dim folderNorm As String: folderNorm = EnsureTrailingSlash(folder)

    Dim diskNames() As String, diskCount As Long
    EnumFolderFiles folderNorm, diskNames, diskCount

    Dim wsM As Worksheet: Set wsM = GetSheet(modAppDef.SH_MANIFEST)
    Dim mPaths() As String, mNames() As String, mModified() As Date
    Dim mSize() As Double, mStatus() As String, mCount As Long
    LoadManifestScope wsM, folderNorm, mPaths, mNames, mModified, mSize, mStatus, mCount

    Dim diskDict As Object: Set diskDict = CreateObject("Scripting.Dictionary")
    Dim ingestedN As Long, replacedN As Long, deletedN As Long
    Dim resumeNeeded As Boolean: resumeNeeded = False

    Dim i As Long
    For i = 0 To diskCount - 1
        If Not diskDict.Exists(LCase$(diskNames(i))) Then diskDict.Add LCase$(diskNames(i)), True

        On Error Resume Next
        modUIMain.SetStage "🔄 同期中 " & (i + 1) & "/" & diskCount & " …"
        On Error GoTo 0

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
        ' status="failed"(例: OneDriveプレースホルダでOpen失敗)は次回同期で
        ' 自動再試行する(§13「OneDriveオフライン…次回再試行」)。DiffDecision
        ' 自体のシグネチャ(§7.8契約)は変えず、呼び出し側でのみ上書きする。
        If decision = "keep" And existsInManifest Then
            If mStatus(mi) = "failed" Then decision = "replace"
        End If

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
    For i = 0 To mCount - 1
        If Not diskDict.Exists(LCase$(mNames(i))) Then
            modShelf.DeleteSource mNames(i)
            deletedN = deletedN + 1
        End If
    Next i

    Dim resumedCount As Long: resumedCount = 0
    If resumeNeeded Then
        On Error Resume Next
        modUIMain.SetStage "📥 未完了の埋め込みを再開しています…"
        On Error GoTo 0
        resumedCount = modEmbed.EmbedPending()
    End If

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

    modLog.LogUsage "sync", "", "ingest=" & ingestedN & " replace=" & replacedN & _
        " delete=" & deletedN & " resumed=" & resumedCount

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
' AutoSyncTick - OnTimeコールバック本体。SyncNow実行後に次回分を再予約する。
'   MASTER_SPEC §7.2の公開契約(closed)に無いためPrivateとする。
' ----------------------------------------------------------------------------
Private Sub AutoSyncTick()
    mScheduled = False
    On Error Resume Next
    SyncNow
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
        ReDim fileNames(0 To -1)
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
    ReDim outPaths(0 To -1)
    ReDim outNames(0 To -1)
    ReDim outModified(0 To -1)
    ReDim outSize(0 To -1)
    ReDim outStatus(0 To -1)
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
