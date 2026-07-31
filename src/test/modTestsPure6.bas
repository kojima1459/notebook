Attribute VB_Name = "modTestsPure6"
Option Explicit

' ============================================================================
' modTestsPure6 - modTestsPure5の分割先(2026-07-31 R8b: 敵対的レビュー対応)
' ----------------------------------------------------------------------------
' 役割:
'   R8の敵対的レビューで見つかった穴(B1〜B16)のうち、純ロジックで固定できる
'   ものをここで押さえる。入口は modTestsPure5.RunAll5 の末尾から呼ばれる
'   Public Sub RunAll6()。
'   (modTestsPure5 が30,000字上限まで残り約200字となったための分割。
'    modTestsPure3→4→5 と同じ中継方式。)
'
' ここに集めたものの共通点:
'   「一度も再現していないが、起きたら取り返しがつかない」種類の穴。
'   ・B1  失効判定の戻り値が想定外(空文字など)のときに知識を消してしまう
'   ・B7b 端末の時計が共有サーバとずれている人が永久に発行できない
'   ・B10 UNC共有ルートを直接指した人が永久に到達不能になる
'   どれも「特定の環境の人だけが、原因の分からないまま被害を受ける」形で、
'   実機で踏んでも報告として上がってきにくい。だから机上で固定する。
' ============================================================================

' ----------------------------------------------------------------------------
' B1: 失効判定の戻り値のホワイトリスト
' ----------------------------------------------------------------------------
' modGuard.EnforceExpiry は WipeKnowledge の直前で
'   If act <> "wipe" Then Exit Function
' と【ホワイトリスト】で守る。従来はここが「1つずつ弾くブラックリスト」で、
' 既定の振る舞いが「消す」だった。EnforceExpiry の冒頭は On Error Resume Next
' なので、my_stats のセルがエラー値(#REF! 等)で GetStatText が失敗すると
' act が空文字のまま消去へ落ちる。
'
' ここでは「ExpiryDecision が返し得る値の集合」を固定する。呼び出し側の
' ホワイトリストは、この集合のうち "wipe" だけを通す実装であることが前提。
' 集合に知らない値が増えたらこのテストが落ち、ホワイトリストの見直しを促す。
Private Sub TestExpiryDecisionVocabulary()
    Dim reach As String: reach = "2026-06-01"
    Dim seenWipe As Boolean
    Dim bad As String
    Dim L As Long, d As Long, w As Long

    ' limitDays / daysSince / 予告済み / 消去済み を広く振って、戻り値が
    ' 想定した6語のいずれかであることを確かめる。
    For L = 0 To 35
        For d = 0 To 40
            For w = 0 To 1
                Dim warned As String
                Dim wiped As String
                If w = 0 Then
                    warned = ""
                    wiped = ""
                Else
                    warned = CStr(d)
                    wiped = ""
                End If
                Dim act As String
                act = modShareRule.ExpiryDecision(reach, d, L, warned, wiped)
                Select Case act
                    Case "off", "never", "none", "warn", "warn_first"
                        ' 想定内。消さない側。
                    Case "wipe"
                        seenWipe = True
                    Case Else
                        If LenB(bad) = 0 Then _
                            bad = "limit=" & L & " days=" & d & " warned=" & warned & " -> [" & act & "]"
                End Select
            Next w
        Next d
    Next L

    modTestRunner.Check "B1: ExpiryDecisionは想定した6語しか返さない", _
        (LenB(bad) = 0), "想定外の戻り値=" & bad
    modTestRunner.Check "B1: 上の総当たりに wipe が実際に含まれる(テストが空回りしていない)", _
        (seenWipe = True), "wipeが1度も出ていない = 条件の作り方が間違っている"

    ' 消去してよい唯一の値は "wipe"。空文字は【絶対に】消去側ではない。
    ' modGuard 側は If act <> "wipe" Then Exit Function で守っているので、
    ' 空文字が来ても消えない。ここではその前提(空はwipeではない)を明示する。
    modTestRunner.Check "B1: 空文字は wipe ではない(ホワイトリストの前提)", _
        (LenB("") = 0 And "" <> "wipe"), ""
    modTestRunner.Check "B1: 未到達端末は総当たりでも never のみ", _
        (modShareRule.ExpiryDecision("", 0, 30, "", "") = "never" And _
         modShareRule.ExpiryDecision("", 400, 1, "400", "") = "never"), ""
End Sub

' ----------------------------------------------------------------------------
' 同時発行の見張り(F12)+ B7b: 発行ロックの時計ズレ
' ----------------------------------------------------------------------------
' (F12のロック境界テストは modTestsPure5 の容量逼迫によりこちらへ移設。
'  同じ関数の境界は1箇所にまとめておく方が、次に触る人が見落とさない。)
'
' B7b: ファイルサーバと端末の時計が数分ずれるのは社内では普通にある。
' その端末から見ると他人のロックは常に「未来の時刻」に見え、経過時間が負になる。
' これを wait にすると、その端末だけが永久に発行できない
' (しかも画面には「いま別の方が発行中です」と出るので、本人は待てば直ると思う)。
' 残骸側へ倒し、呼び出し側で err_log(E0807)へ1行残す。
Private Sub TestPublishLockClockSkew()
    ' --- F12: ロックが無い / 新しい / 古い の基本 ---
    modTestRunner.Check "発行ロック: ロック無しなら go", _
        (modShareRule.PublishLockAction(False, 0, 10) = "go"), ""
    modTestRunner.Check "発行ロック: ロック無しは経過時間を見ない", _
        (modShareRule.PublishLockAction(False, 999, 10) = "go"), ""
    modTestRunner.Check "発行ロック: 直後(0分)は wait", _
        (modShareRule.PublishLockAction(True, 0, 10) = "wait"), ""
    modTestRunner.Check "発行ロック: 9.9分は wait", _
        (modShareRule.PublishLockAction(True, 9.9, 10) = "wait"), ""
    ' 境界: ちょうど10分は残骸とみなす(発行は数十秒で終わるため)。
    modTestRunner.Check "発行ロック: ちょうど10分は stale", _
        (modShareRule.PublishLockAction(True, 10, 10) = "stale"), _
        "実際=" & modShareRule.PublishLockAction(True, 10, 10)
    modTestRunner.Check "発行ロック: 60分は stale", _
        (modShareRule.PublishLockAction(True, 60, 10) = "stale"), ""

    ' --- B7b: 経過が負(ロックの更新時刻が未来)は残骸扱いへ変更した ---
    modTestRunner.Check "B7b: 経過が負(-1分)は stale", _
        (modShareRule.PublishLockAction(True, -1, 10) = "stale"), _
        "実際=" & modShareRule.PublishLockAction(True, -1, 10)
    modTestRunner.Check "B7b: 経過が負(-5分)は stale", _
        (modShareRule.PublishLockAction(True, -5, 10) = "stale"), _
        "実際=" & modShareRule.PublishLockAction(True, -5, 10)
    modTestRunner.Check "B7b: 経過が負(-600分)は stale", _
        (modShareRule.PublishLockAction(True, -600, 10) = "stale"), ""

    ' 待たせるのは「ロックがあり、経過が0以上 staleMinutes 未満」のときだけ。
    Dim m As Long
    Dim badWait As Boolean
    For m = -20 To 20
        Dim act As String
        act = modShareRule.PublishLockAction(True, CDbl(m), 10)
        If act = "wait" Then
            If m < 0 Or m >= 10 Then badWait = True
        End If
    Next m
    modTestRunner.Check "B7b: waitになるのは0分以上10分未満のときだけ", _
        (badWait = False), "時計ズレや残骸で待たせている"
End Sub

' ----------------------------------------------------------------------------
' B10: UNC共有ルート直指定のときのプローブ再試行
' ----------------------------------------------------------------------------
' ProbeTargetPath は GetAttr のために末尾の "\" を1つ落とす。ところが
' 共有ルートを直接指した場合、環境によっては "\\srv\share" では
' 実行時エラー52/76 が返り、"\\srv\share\" でしか通らない。
' 1回目が失敗したときだけ、末尾を残した形でもう一度試す。
' 「形が変わらないなら再試行しない」= 遅い共有でタイムアウトを二重に払わない。
Private Sub TestProbeRetryPath()
    ' 末尾 "\" を落とした = 形が変わった → もう1回試す価値がある。
    modTestRunner.Check "B10: UNC共有ルートは末尾\付きで再試行する", _
        (modShareRule.ProbeRetryPath("\\srv\share\") = "\\srv\share\"), _
        "実際=[" & modShareRule.ProbeRetryPath("\\srv\share\") & "]"
    modTestRunner.Check "B10: 深いUNCパスも末尾\付きで再試行する", _
        (modShareRule.ProbeRetryPath("\\srv\share\Nexus\") = "\\srv\share\Nexus\"), _
        "実際=[" & modShareRule.ProbeRetryPath("\\srv\share\Nexus\") & "]"
    modTestRunner.Check "B10: ローカルの深いパスも再試行対象", _
        (modShareRule.ProbeRetryPath("C:\data\nexus\") = "C:\data\nexus\"), ""

    ' 末尾に "\" が無い = 1回目がそのままの形 → 再試行しない。
    modTestRunner.Check "B10: 末尾\が無ければ再試行しない", _
        (LenB(modShareRule.ProbeRetryPath("\\srv\share")) = 0), _
        "実際=[" & modShareRule.ProbeRetryPath("\\srv\share") & "]"
    ' ドライブルート "C:\" は ProbeTargetPath が "\" を落とさない
    ' = 1回目と同じ実引数になるので、再試行しても意味が無い。
    modTestRunner.Check "B10: ドライブルート C:\ は再試行しない(1回目と同形)", _
        (LenB(modShareRule.ProbeRetryPath("C:\")) = 0), _
        "実際=[" & modShareRule.ProbeRetryPath("C:\") & "]"
    ' "\\host\" も同様(ホスト名だけになるので ProbeTargetPath が落とさない)。
    modTestRunner.Check "B10: \\host\ は再試行しない(1回目と同形)", _
        (LenB(modShareRule.ProbeRetryPath("\\host\")) = 0), _
        "実際=[" & modShareRule.ProbeRetryPath("\\host\") & "]"
    modTestRunner.Check "B10: 空文字は再試行しない", _
        (LenB(modShareRule.ProbeRetryPath("")) = 0), ""

    ' 再試行の実引数は「1回目と必ず違う」ことが存在意義。同じものを2回投げると
    ' 遅い共有に対して OS のタイムアウトを二重に払うだけになる。
    Dim samples As Variant
    samples = Array("\\srv\share\", "\\srv\share\Nexus\", "C:\data\nexus\", _
                    "\\srv\share", "C:\", "\\host\", "", "   ")
    Dim i As Long
    Dim badSame As Boolean
    For i = LBound(samples) To UBound(samples)
        Dim rp As String: rp = modShareRule.ProbeRetryPath(CStr(samples(i)))
        If LenB(rp) > 0 Then
            If StrComp(rp, modShareRule.ProbeTargetPath(CStr(samples(i))), vbBinaryCompare) = 0 Then
                badSame = True
            End If
        End If
    Next i
    modTestRunner.Check "B10: 再試行パスは1回目の実引数と必ず異なる", _
        (badSame = False), "同じ実引数で2回GetAttrを呼ぶ = 無駄なタイムアウト"

    ' 再試行する場合、その形は必ず末尾 "\" 付き(=元の設定値そのもの)。
    modTestRunner.Check "B10: 再試行パスは末尾\付き", _
        (Right$(modShareRule.ProbeRetryPath("\\srv\share\"), 1) = "\"), ""
End Sub

Public Sub RunAll6()
    On Error GoTo VocabFail
    TestExpiryDecisionVocabulary
NextSkew:
    On Error GoTo SkewFail
    TestPublishLockClockSkew
NextRetry:
    On Error GoTo RetryFail
    TestProbeRetryPath
NextDone6:
    On Error GoTo 0
    Exit Sub

VocabFail:
    modTestRunner.Check "TestExpiryDecisionVocabulary(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextSkew
SkewFail:
    modTestRunner.Check "TestPublishLockClockSkew(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRetry
RetryFail:
    modTestRunner.Check "TestProbeRetryPath(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone6
End Sub
