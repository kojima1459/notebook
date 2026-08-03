Attribute VB_Name = "optGsProc"
Option Explicit

' ============================================================================
' optGsProc - Ghostscriptプロセスの起動と停止(opt機能・2026-08-03 R13-F2)
' ----------------------------------------------------------------------------
' 役割:
'   GSを非表示・非同期で起動してPIDを控える(RunGsAsync)、真のハングで
'   その系統を止める(KillGsTree)、の2つだけ。呼び出し元は optGsTxt
'   (txtwrite経路)と optVision(画像PDFのOCR経路)。
'
' なぜ optGsTxt から分けたのか:
'   PID再利用よけの本人確認(WMI Win32_Process の名前照合)と、rc=0でPIDが
'   読めなかったときの扱いを足した結果、optGsTxt が28,000字のWARN帯に入った
'   (憲章§4-6 / R13容量ルール)。「プロセスを起こす・殺す」はGSの出力を
'   読む処理とは別の関心事で、切り出しても呼び出しは各1行で済む。
'
' 設計判断:
'   ・起動はWMI(Win32_Process.Create)優先。PIDが無いとハングを止められず、
'     GS/cmd.exe がログオン中ずっと残る(実機第2報 RC3)。WMIがEDR/ポリシーで
'     塞がれている端末では従来のWScript.Shellへ必ず退避する(取込が動かなく
'     なる方が害が大きい)。退避したセッションでは停止できないことを
'     usage_log へ1回だけ残す。
'   ・停止は必ず本人確認つき。PIDは再利用されうるので、名前を確かめずに
'     taskkill /T /F を撃つのは無関係なプロセスへの誤爆になる。
'   ・COM参照はどの失敗経路でも必ず解放する。ハンドラ稼働中は
'     On Error Resume Next が効かないので、Resumeでハンドラを抜けてから
'     後始末する(opt層GS系モジュール共通の作法)。
' ============================================================================

' R13-1e: WMI起動が使えずPIDを取れなかったことを1セッション1回だけ記録する
' ためのフラグ(EDR等でWMIが塞がれている端末では毎回失敗するので、
' 記録が usage_log を埋め尽くさないようにする)。
Private mKillUnavailableLogged As Boolean

Public Function Ping() As Boolean
    Ping = True
End Function

' 非同期起動(待たない)。0=ウィンドウ非表示 / False=完了を待たない。
' R10-2: 失敗時はErr.Number/Descriptionを呼び出し元へByRefで返す(戻り値は
' Booleanのまま・モジュール変数を増やさない最小構成)。呼び出し元がerr_logの
' detailへ含めることで、WScript.Shell自体がポリシーでブロックされる端末を
' 「パスは合っているのにGSが動かない」から切り分けられるようにする。
' R13-1e: まずWMI(Win32_Process.Create)で起動してPIDを控える。PIDが無いと
' 真のハングを止める手段が無く、GS/cmd.exe がログオン中ずっと残る(RC3)。
' WMIはEDR/ポリシーで塞がれる端末があるので、失敗したら【必ず】従来の
' WScript.Shell へ退避する(取込が動かなくなる方が害が大きい)。退避した
' セッションではkillできないことを1回だけ usage_log へ残す。
' pidOut: 取れたPID(0=不明。省略可なので既存の3引数呼び出しはそのまま動く)。
Public Function RunGsAsync(ByVal runCmd As String, ByRef errNum As Long, _
                            ByRef errDesc As String, _
                            Optional ByRef pidOut As Long = 0) As Boolean
    pidOut = 0
    If RunGsViaWmi(runCmd, pidOut) Then
        RunGsAsync = True
        Exit Function
    End If
    NoteKillUnavailable

    Dim wsh As Object
    On Error GoTo NoRun
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run runCmd, 0, False
    Set wsh = Nothing
    RunGsAsync = True
    Exit Function
NoRun:
    errNum = Err.Number
    errDesc = Err.Description
    RunGsAsync = False
End Function

' WMI経由の非表示起動。成功でTrue+PID、どこかで失敗したらFalse(呼び出し元が
' 従来経路へ退避)。GetObject/Get/SpawnInstance_/Create のどれもポリシーで
' 落ちうるので1本のハンドラで受け、Resumeでハンドラを抜けてからCOM参照を
' 必ず解放する(本モジュール共通の作法。ハンドラ稼働中はOERNが効かない)。
Private Function RunGsViaWmi(ByVal runCmd As String, ByRef pidOut As Long) As Boolean
    Dim svc As Object
    Dim startCls As Object
    Dim cfg As Object
    Dim proc As Object
    Dim newPid As Variant
    Dim rc As Long

    On Error GoTo WmiFail
    Set svc = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set startCls = svc.Get("Win32_ProcessStartup")
    Set cfg = startCls.SpawnInstance_()
    cfg.ShowWindow = 0                 ' 0=SW_HIDE(従来のwsh.Run(...,0,False)と同じ)
    Set proc = svc.Get("Win32_Process")
    rc = proc.Create(runCmd, Null, cfg, newPid)

    Set proc = Nothing
    Set cfg = Nothing
    Set startCls = Nothing
    Set svc = Nothing
    On Error GoTo 0

    ' rc<>0 は「起動できなかった」。ここだけがFalse(=wsh.Runへ退避)でよい。
    If rc <> 0 Then Exit Function

    ' R13-L12: rc=0 なら GS は【もう動いている】。PIDが読めなかったからと
    ' False を返すと、呼び出し元が同じ出力先へ2本目のGSを起動して二重書き込み
    ' になる。PIDが取れないときは「停止できない」だけなので、その事実を
    ' 1回記録して True で返す。
    If IsNumeric(newPid) Then
        If CLng(newPid) > 0 Then pidOut = CLng(newPid)
    End If
    If pidOut = 0 Then NoteKillUnavailable

    RunGsViaWmi = True
    Exit Function

WmiFail:
    Resume WmiCleanup
WmiCleanup:
    On Error Resume Next
    Set proc = Nothing
    Set cfg = Nothing
    Set startCls = Nothing
    Set svc = Nothing
    On Error GoTo 0
    RunGsViaWmi = False
End Function

' PIDが取れなかったことの記録(1セッション1回)。「なぜ止められなかったのか」
' が後から分かるようにする(憲章§4-1 無言の失敗禁止)。
Private Sub NoteKillUnavailable()
    If mKillUnavailableLogged Then Exit Sub
    mKillUnavailableLogged = True
    On Error Resume Next
    modLog.LogUsage "gs_kill_unavailable", "", _
        "WMI起動が使えずWScript.Shellへ退避(PID未取得のためハング時に停止できません)"
    On Error GoTo 0
End Sub

' 真のハング(アイドル上限超過)でGSを系統ごと止める(R13-1e)。
' /T で子プロセス(cmd.exe が起動した gswin32c.exe)まで、/F で強制終了。
' PIDが無い(WScript.Shell退避)ときは何もしない。
' R13-F2: 止める前に「そのPIDが本当に自分の起動したGSか」をWMIで確かめる。
' 待っている間に対象が終了していればWindowsはPIDを再利用しうるので、確認
' 抜きの taskkill /T /F は無関係なプロセス(最悪は業務アプリ)を巻き添えに
' する。名前が cmd.exe / gswin32c.exe のときだけ実行し、確認できなければ
' 「止めなかった」ことを記録して黙って帰る(誤爆より取りこぼしを選ぶ)。
Public Sub KillGsTree(ByVal pid As Long)
    If pid <= 0 Then Exit Sub

    Dim nm As String: nm = ProcNameOf(pid)
    Dim low As String: low = LCase$(Trim$(nm))
    If low <> "cmd.exe" And low <> "gswin32c.exe" Then
        On Error Resume Next
        modLog.LogUsage "gs_kill_skipped", "", "pid=" & pid & " name=[" & nm & "]"
        On Error GoTo 0
        Exit Sub
    End If

    Dim wsh As Object
    On Error GoTo KillFail
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "taskkill /T /F /PID " & CStr(pid), 0, False
    Set wsh = Nothing
    On Error GoTo 0
    On Error Resume Next
    modLog.LogUsage "gs_killed_on_hang", "", "pid=" & pid
    On Error GoTo 0
    Exit Sub
KillFail:
    Resume KillCleanup
KillCleanup:
    On Error Resume Next
    Set wsh = Nothing
    modLog.LogUsage "gs_kill_failed", "", "pid=" & pid
    On Error GoTo 0
End Sub

' PIDに今ぶら下がっているプロセス名をWMIで引く(R13-F2)。
' 見つからない(すでに終了している)ときは ""、WMIそのものが使えない/
' 問い合わせに失敗したときは "?" を返す。どちらも呼び出し元はkillしない。
' PIDはWMIのCreateで得たものなので、この端末でWMIが通ること自体は既知。
Private Function ProcNameOf(ByVal pid As Long) As String
    Dim svc As Object
    Dim col As Object
    Dim p As Object

    On Error GoTo NameFail
    Set svc = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set col = svc.ExecQuery("SELECT Name FROM Win32_Process WHERE ProcessId=" & CStr(pid))
    For Each p In col
        ProcNameOf = CStr(p.Name)
        Exit For
    Next p
    Set p = Nothing
    Set col = Nothing
    Set svc = Nothing
    On Error GoTo 0
    Exit Function

NameFail:
    Resume NameCleanup
NameCleanup:
    On Error Resume Next
    Set p = Nothing
    Set col = Nothing
    Set svc = Nothing
    On Error GoTo 0
    ProcNameOf = "?"
End Function
