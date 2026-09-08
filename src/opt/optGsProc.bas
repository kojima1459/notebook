Attribute VB_Name = "optGsProc"
Option Explicit

' ============================================================================
' optGsProc - Ghostscriptプロセスの起動と停止(opt機能・2026-08-03 R13-F2)
' ----------------------------------------------------------------------------
' 役割:
'   GSを非表示・非同期で起動してPIDを控える(RunGsAsync)、終了を検知して
'   完了フラグを書く(SyncDoneFlag)、真のハングで止める(KillGsTree)。
'   呼び出し元は optGsTxt(txtwrite経路)と optOcrPage(画像PDFのOCR経路)。
'
' 2026-09-08(R40 F1・実機報告): 「悪意のあるマクロが検出されました。データを
'   保護するため、Office を終了します」で PDF 取込中に Excel が強制終了した。
'   これは Defender の AMSI(マクロの実行時挙動監視)で、R13〜R39 の起動経路
'   【Office マクロ → WMI Win32_Process.Create → cmd.exe /s /c "(…) 1>… 2>&1
'   & (if errorlevel …) >flag"】は、マクロ型マルウェアの典型手口
'   (WMI 経由の cmd 起動・出力リダイレクト・条件分岐付きバッチ)と字面が
'   同じになる。信頼できる場所の追加も情シス連絡もできない(台帳)ので、
'   【疑われる部品そのものを無くす】: cmd.exe も WMI も WScript.Shell も
'   使わず、VBA の Shell 関数で gswin32c.exe を直接起動し、kernel32 の
'   OpenProcess/GetExitCodeProcess で終了を見張り、完了フラグは VBA 自身が
'   書く(フラグの中身=GSの終了コード。読む側 GsExitCodeFromFlag は不変)。
'   GS の標準出力は GS 自身の -sstdout=<gs_out.log> でファイルへ落とす
'   (optOcrCore.BuildRunCommand)。停止は taskkill ではなく TerminateProcess。
'   起動時に得たハンドルを終了まで保持するので PID は再利用されない
'   (R13-F2 の WMI 名前照合は不要になった)。
'
' 設計判断:
'   ・Shell が失敗したとき(実行ファイル無し・AppLocker 等)は Err が上がるので
'     errNum/errDesc へ写して False(呼び出し元が E0302 に記録)。
'   ・完了フラグの書き手はこのモジュールだけになった(cmd 時代の「& で無条件に
'     作る」保証が消えた)ので、退路を2段持つ(レビュー R40-B1):
'       1) 起動直後に OpenProcess が失敗しても、SyncDoneFlag が呼ばれるたびに
'          取り直す(EDR の一時的な拒否・起動直後の競合に効く)。
'       2) 終了コードを読む権限が無くても SYNCHRONIZE だけで開けるなら
'          WaitForSingleObject(0) で「終わった」ことだけを見て、フラグには
'          "?"(=GsExitCodeFromFlag が -1・失敗と決めつけない側)を書く。
'     それでも開けない(PID がもう無い=起動直後に終了した等)ときは、
'     OpenProcess の LastDllError が 87(無効な PID)なら終了とみなして "?"。
'   ・フラグは【書いてから】ハンドルを離す。書けなければ次の呼び出しで再試行。
'   ・ハンドルは必ず CloseHandle(SyncDoneFlag/KillGsTree/次の起動時)。
' ============================================================================

#If Win64 Then
    Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
    Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByRef lpExitCode As Long) As Long
    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
    Private Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
#Else
    Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
#End If

Private Const PROCESS_TERMINATE As Long = &H1
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000
Private Const SYNCHRONIZE_ACCESS As Long = &H100000
Private Const STILL_ACTIVE As Long = 259
Private Const WAIT_OBJECT_0 As Long = 0
Private Const ERROR_INVALID_PARAMETER As Long = 87
' Shell 関数の第2引数(vbHide=0)。LO では定数名が無いので数値で書く。
Private Const SHELL_HIDE As Long = 0
' 終了コードが読めないときにフラグへ書く印(GsExitCodeFromFlag → -1 = 不明)。
Private Const FLAG_UNKNOWN As String = "?"

' 直近に起動した GS(1本しか同時に動かさない。optGsTxt/optOcrPage は直列)。
' ハンドルは Win64 でも下位32bitに収まる(カーネルハンドルの公式仕様)ので Long。
' mSyncOnly=True は「SYNCHRONIZE だけで開けた」(終了コードは読めない)。
Private mPid As Long
Private mHandle As Long
Private mSyncOnly As Boolean
' OpenProcess の直近の失敗理由(Err.LastDllError)。0=成功。
Private mLastOpenErr As Long

' R13-1e: 停止手段が無いことを1セッション1回だけ記録するフラグ。
Private mKillUnavailableLogged As Boolean

Public Function Ping() As Boolean
    Ping = True
End Function

' 非同期起動(待たない)。gswin32c.exe を Shell で直接・非表示で起こす。
' 失敗時はErr.Number/Descriptionを呼び出し元へByRefで返す(戻り値はBoolean)。
' pidOut: 取れたPID(0=不明)。
Public Function RunGsAsync(ByVal runCmd As String, ByRef errNum As Long, _
                            ByRef errDesc As String, _
                            Optional ByRef pidOut As Long = 0) As Boolean
    pidOut = 0
    ReleaseHandle

    Dim taskId As Double
    On Error GoTo NoRun
    taskId = Shell(runCmd, SHELL_HIDE)
    On Error GoTo 0
    If taskId <= 0 Then
        errNum = 5
        errDesc = "Shell がプロセスIDを返しませんでした"
        Exit Function
    End If

    pidOut = CLng(taskId)
    mPid = pidOut
    TryOpenHandle
    If mHandle = 0 Then NoteKillUnavailable
    RunGsAsync = True
    Exit Function
NoRun:
    errNum = Err.Number
    errDesc = Err.Description
    RunGsAsync = False
End Function

' 直近に起動した GS が終わっていれば、終了コードを完了フラグへ書いて True。
' まだ動いている・問い合わせに失敗、は False(呼び出し元の待ちループはそのまま
' 回る)。フラグが既に在れば上書きしない。書けなかったときはハンドルを離さず
' 次の呼び出しでもう一度書く(R40 レビュー M3)。
Public Function SyncDoneFlag(ByVal flagPath As String) As Boolean
    If mPid = 0 Then Exit Function
    Dim body As String
    If mHandle = 0 Then
        ' 起動直後に開けなかった端末: 呼ばれるたびに取り直す(R40 レビュー B1)。
        TryOpenHandle
        If mHandle = 0 Then
            ' PID がもう無い(起動直後に終了)ときだけ「終わった・コード不明」。
            If mLastOpenErr <> ERROR_INVALID_PARAMETER Then Exit Function
            body = FLAG_UNKNOWN
        End If
    End If
    If LenB(body) = 0 Then
        If mSyncOnly Then
            If WaitForSingleObject(mHandle, 0) <> WAIT_OBJECT_0 Then Exit Function
            body = FLAG_UNKNOWN
        Else
            Dim code As Long: code = STILL_ACTIVE
            If GetExitCodeProcess(mHandle, code) = 0 Then Exit Function
            If code = STILL_ACTIVE Then Exit Function
            body = CStr(code)
        End If
    End If

    Dim written As Boolean
    If optVision.PathExists(flagPath) Then
        written = True
    Else
        written = WriteFlag(flagPath, body)
    End If
    If Not written Then Exit Function
    ReleaseHandle
    SyncDoneFlag = True
End Function

' 真のハング(アイドル上限超過)でGSを止める(R13-1e)。
' 戻り値の意味は R33H F20 のまま【もうGSは動いていないと言い切れる】:
'   既に終了していた → True(止める必要が無かった=成功側)
'   TerminateProcess 成功 → True
'   ハンドルが無い/別のPID/失敗 → False(生きているか分からないので
'   作業フォルダを残す側。誤爆より取りこぼしを選ぶ)
' R40 F1: ハンドルを保持している限り PID は再利用されないので、WMI の
' 名前照合と taskkill(cmd 経由)は不要になった。
Public Function KillGsTree(ByVal pid As Long) As Boolean
    If pid <= 0 Then Exit Function
    If pid = mPid And mHandle = 0 Then TryOpenHandle
    If pid <> mPid Or mHandle = 0 Then
        On Error Resume Next
        modLog.LogUsage "gs_kill_skipped", "", "pid=" & pid & " (ハンドル無し)"
        On Error GoTo 0
        Exit Function
    End If

    If HasExited() Then
        ReleaseHandle
        On Error Resume Next
        modLog.LogUsage "gs_kill_not_needed", "", "pid=" & pid & " (既に終了)"
        On Error GoTo 0
        KillGsTree = True
        Exit Function
    End If

    If TerminateProcess(mHandle, 1) <> 0 Then
        KillGsTree = True
        On Error Resume Next
        modLog.LogUsage "gs_killed_on_hang", "", "pid=" & pid
        On Error GoTo 0
    Else
        On Error Resume Next
        modLog.LogUsage "gs_kill_failed", "", "pid=" & pid
        On Error GoTo 0
    End If
    ReleaseHandle
End Function

' 保持ハンドルのプロセスが終わっているか(終了コードの読めない SYNCHRONIZE
' だけのハンドルにも答える)。判定できないときは False(生きている側)。
Private Function HasExited() As Boolean
    If mHandle = 0 Then Exit Function
    If mSyncOnly Then
        HasExited = (WaitForSingleObject(mHandle, 0) = WAIT_OBJECT_0)
    Else
        Dim code As Long: code = STILL_ACTIVE
        If GetExitCodeProcess(mHandle, code) = 0 Then Exit Function
        HasExited = (code <> STILL_ACTIVE)
    End If
End Function

' mPid のハンドルを取る。終了コードが読める権限を優先し、駄目なら
' SYNCHRONIZE だけ(mSyncOnly=True)。全部駄目なら mHandle=0 のまま。
Private Sub TryOpenHandle()
    If mPid = 0 Then Exit Sub
    mSyncOnly = False
    mLastOpenErr = 0
    On Error Resume Next
    mHandle = CLng(OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION Or PROCESS_TERMINATE, 0, mPid))
    If mHandle = 0 Then mHandle = CLng(OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, 0, mPid))
    If mHandle = 0 Then mHandle = CLng(OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, 0, mPid))
    If mHandle = 0 Then
        mHandle = CLng(OpenProcess(SYNCHRONIZE_ACCESS, 0, mPid))
        If mHandle <> 0 Then mSyncOnly = True
    End If
    If mHandle = 0 Then mLastOpenErr = Err.LastDllError
    On Error GoTo 0
End Sub

' 完了フラグを書く(cmd.exe の echo の代わり)。書けたら True。
' Open の後に落ちてもファイル番号を必ず閉じる(R40 レビュー M3)。
Private Function WriteFlag(ByVal flagPath As String, ByVal body As String) As Boolean
    Dim fn As Long: fn = 0
    On Error GoTo WriteFail
    fn = FreeFile
    Open flagPath For Output As #fn
    Print #fn, body
    Close #fn
    WriteFlag = True
    Exit Function
WriteFail:
    Resume WriteCleanup
WriteCleanup:
    On Error Resume Next
    If fn <> 0 Then Close #fn
    On Error GoTo 0
    WriteFlag = False
End Function

Private Sub ReleaseHandle()
    If mHandle <> 0 Then
        On Error Resume Next
        CloseHandle mHandle
        On Error GoTo 0
    End If
    mHandle = 0
    mPid = 0
    mSyncOnly = False
End Sub

' 停止手段が無いことの記録(1セッション1回)。「なぜ止められなかったのか」
' が後から分かるようにする(憲章§4-1 無言の失敗禁止)。
Private Sub NoteKillUnavailable()
    If mKillUnavailableLogged Then Exit Sub
    mKillUnavailableLogged = True
    On Error Resume Next
    modLog.LogUsage "gs_kill_unavailable", "", _
        "起動直後にプロセスハンドル未取得(err=" & mLastOpenErr & "。待ちループで取り直す。取れないままなら停止できず、終了もアイドル上限で見切る)"
    On Error GoTo 0
End Sub
