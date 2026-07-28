Attribute VB_Name = "modShare"
Option Explicit

' ============================================================================
' modShare - 共有フォルダへの到達性を1セッションに1回だけ確かめる関所
' ----------------------------------------------------------------------------
' 解決する問題(2026-07-28 レビュー M-20):
'   共有フォルダ(nexus_share_path)が届かないとき、このアプリは
'   「届かない」と判定するまで OS の SMB タイムアウトに委ねていた。
'   存在しないホストへの最初の UNC アクセスは 10〜30 秒ブロックし得る。
'   ところが共有I/Oは
'     起動時: 感謝状の回収 → 品質報告の回収 → 共有Q&Aの回収 → チャンネル確認
'     終了時: テレメトリ送信
'   と直列に走るため、1回30秒×複数箇所で「開かない」「閉じたのに数十秒残る」
'   という体感になる。しかも共有が死んでいる日は全員が同じ目に遭う。
'
' 方針:
'   最初の1回だけ実際に触りに行き、届かなければ【そのセッションの間は
'   共有I/Oを全部あきらめる】。届かない共有に何度も挨拶しに行かない。
'   届いた場合はキャッシュして、以後は判定コストをゼロにする。
'
'   「今日はたまたま遅かっただけ」は次にブックを開けば作り直される。
'   セッションを跨いで諦め続けることはしない。
'
' 使い方:
'   各モジュールの <なんとか>Dir() は、パスを組み立てる前に Reachable() を
'   通す。届かなければ空文字を返し、呼び出し側は既存の「パスが空なら
'   何もしない」経路にそのまま乗る(新しい失敗経路を増やさない)。
' ============================================================================

' 0 = 未判定 / 1 = 到達可 / 2 = 到達不可(このセッションは以後スキップ)
Private mState As Long
Private Const ST_UNKNOWN As Long = 0
Private Const ST_OK As Long = 1
Private Const ST_NG As Long = 2

' ----------------------------------------------------------------------------
' BasePath - nexus_share_path を末尾"\"付きで返す。未設定なら空。
'   到達性は見ない(設定されているかどうかだけ)。
' ----------------------------------------------------------------------------
Public Function BasePath() As String
    On Error Resume Next
    Dim p As String
    p = Trim$(modConfig.GetString("nexus_share_path", ""))
    If LenB(p) = 0 Then Exit Function
    If Right$(p, 1) <> "\" Then p = p & "\"
    BasePath = p
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' Reachable - 共有フォルダに届くか。判定は1セッションに1回だけ。
' ----------------------------------------------------------------------------
Public Function Reachable() As Boolean
    If mState = ST_OK Then
        Reachable = True
        Exit Function
    End If
    If mState = ST_NG Then Exit Function

    Dim p As String: p = BasePath()
    If LenB(p) = 0 Then
        ' 未設定。届かないのではなく「使わない」設定なので、
        ' ログには残さず静かに NG にする。
        mState = ST_NG
        Exit Function
    End If

    On Error Resume Next
    Dim probe As String
    probe = Dir(p, vbDirectory)      ' ここで OS のタイムアウトを1回だけ払う
    Dim probeErr As Long: probeErr = Err.Number
    Err.Clear
    On Error GoTo 0

    If probeErr = 0 And LenB(probe) > 0 Then
        mState = ST_OK
        Reachable = True
    Else
        mState = ST_NG
        On Error Resume Next
        modLog.LogUsage "share_unreachable", "", _
            "このセッションは共有フォルダへのアクセスを見送ります: " & modUtil.SafeLeft(p, 200)
        On Error GoTo 0
    End If
End Function

' ----------------------------------------------------------------------------
' SubDir - <nexus_share_path>\<name>\ を返す。届かないときは空文字。
' ----------------------------------------------------------------------------
Public Function SubDir(ByVal leaf As String) As String
    If Not Reachable() Then Exit Function
    Dim p As String: p = BasePath()
    If LenB(p) = 0 Then Exit Function
    SubDir = p & leaf & "\"
End Function

' ----------------------------------------------------------------------------
' ResetProbe - 到達性の判定をやり直させる(共有パスを設定し直した直後など)。
' ----------------------------------------------------------------------------
Public Sub ResetProbe()
    mState = ST_UNKNOWN
End Sub
