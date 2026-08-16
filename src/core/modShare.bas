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
'
' 唯一性の原則(2026-07-31 R8 F2):
'   共有ルートの解決と到達判定は【このモジュールだけ】が行う。
'   modBoard/modPublish/modMentor は自前で nexus_share_path を読んで
'   パスを組み立てており、modShare の関所を素通りしていた。結果、
'   届かない共有に対して起動のたび OS のタイムアウトを払い直していたうえ、
'   「modShare は NG と言っているのに他は触りに行く」という矛盾した状態に
'   なっていた。パスの組み立ては必ず SubDir() を通すこと。
' ============================================================================

' 0 = 未判定 / 1 = 到達可 / 2 = 到達不可(このセッションは以後スキップ)
Private mState As Long
Private Const ST_UNKNOWN As Long = 0
Private Const ST_OK As Long = 1
Private Const ST_NG As Long = 2

' 到達OKと判定したあとの連続失敗回数(2026-07-31 R8・F7の最小実装)。
' セッション途中のVPN切断で「届く」と言い張り続けないための受け口。
' 完全な降格機構(状態機械)は次期。ここは連続 FAIL_TO_NG 回で NG へ倒すだけ。
Private mFailStreak As Long
Private Const FAIL_TO_NG As Long = 5

' 標準サブフォルダの初期化を1セッションに1回だけ行うための印。
Private mDirsEnsured As Boolean

' 集約スナップショット(board\summary.txt)1行目の種別印(2026-08-16 R33 W6-1)。
' 書式の詳細はモジュール末尾の BoardHeadText 直上を参照。
Public Const BOARD_HEAD_TAG As String = "nxboard1"

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

    If ProbePath(p) Then
        mState = ST_OK
        mFailStreak = 0
        Reachable = True
        EnsureStandardDirs p
    Else
        mState = ST_NG
        On Error Resume Next
        modLog.LogUsage "share_unreachable", "", _
            "このセッションは共有フォルダへのアクセスを見送ります: " & modUtil.SafeLeft(p, 200)
        On Error GoTo 0
    End If
End Function

' ----------------------------------------------------------------------------
' ProbePath - 指定したパスへ実際に届くかを1回だけ確かめる(キャッシュしない)。
'
' 判定の中身(R8 F6 / R8b B10):
'   ・プローブは「中身が1件以上あるか」ではなく「ディレクトリとして存在するか」。
'     旧実装は Dir(p, vbDirectory) の戻り文字列を見ていたが、Dir にフォルダを
'     渡すと返るのは【そのフォルダの中の最初のエントリ】で、空フォルダでは
'     空文字になる。つまり「共有ルートは正常に作られているが、まだ誰も何も
'     置いていない」という健全な初期状態が到達不能と判定されていた。
'     しかも共有ルートの下に thanks\ などを作るのは全部 modShare 経由なので、
'     誰も最初の1ファイルを置けない=永久に空=永久に到達不能、という
'     デッドロックになる(PoC開始直後がまさにこの状態)。
'     GetAttr は「存在するか」だけを答えるので、空フォルダでも真になる。
'   ・1回目は末尾の "\" を落とした形("\\srv\share")で試す。ところが環境に
'     よってはそこで実行時エラー52/76が返る(共有ルートは "\\srv\share\" の
'     形でしか受け付けない実装がある)。失敗したときだけ、末尾を残した形で
'     もう一度だけ試す。成功する見込みが無い(形が変わらない)ときは
'     ProbeRetryPath が空文字を返すので、無駄なタイムアウトは払わない。
'
' なぜ Public な関数に切り出すか(2026-07-31 C1):
'   この2段プローブは Reachable() と、共有パスの設定画面
'   (modHelp.OnShareSetup の保存前チェック)の【2箇所】で必要になる。
'   設定画面側にコピーを置いていたところ、B10の再試行を入れたのが
'   Reachable() だけだったため、
'     「UNC共有ルート(\\srv\share\)を入力 → 設定画面が『見つかりません』と
'       拒否 → そもそも保存できないので Reachable まで到達しない」
'   という、片側だけ直っても意味が無い形になっていた。
'   判定式が2箇所にあると、いつか必ず片方だけが更新される。
'   関数を1本にして「2箇所で違う答えを出さない」を構造で保証する。
' ----------------------------------------------------------------------------
Public Function ProbePath(ByVal p As String) As Boolean
    Dim target As String: target = Trim$(p)
    If LenB(target) = 0 Then Exit Function

    On Error Resume Next
    Err.Clear
    Dim t0 As Double: t0 = Timer
    Dim attrVal As Long
    attrVal = GetAttr(modShareRule.ProbeTargetPath(target))  ' OSのタイムアウトはここで払う
    Dim probeErr As Long: probeErr = Err.Number
    ' 1回目にかかった時間を測る(2026-08-16 R33 W4-6)。この値だけが
    ' 「構文で即座に弾かれた」と「死んだホストのタイムアウトを払った」を
    ' 区別できる。測らずに再試行すると、届かない共有へフルのタイムアウトを
    ' 2回払い、起動の無反応が倍(20〜60秒)になる。
    Dim firstMs As Long: firstMs = CLng(modUtilText.ElapsedMsSince(t0))
    Err.Clear

    If Not modShareRule.ProbeIsReachable(probeErr, attrVal) Then
        If modShareRule.ShouldRetryProbe(probeErr, firstMs) Then
            Dim retryPath As String: retryPath = modShareRule.ProbeRetryPath(target)
            If LenB(retryPath) > 0 Then
                Err.Clear
                attrVal = GetAttr(retryPath)
                probeErr = Err.Number
                Err.Clear
            End If
        End If
    End If
    On Error GoTo 0

    ProbePath = modShareRule.ProbeIsReachable(probeErr, attrVal)
End Function

' ----------------------------------------------------------------------------
' EnsureStandardDirs - 初回OK時に標準サブフォルダを1回だけ用意する。
'   これが「共有フォルダの初期化の入口」で、アプリ内にここ以外は作らない
'   (2026-07-31 R8 F6)。各モジュールが自分のフォルダを作るだけだと、
'   その機能を一度も使わない限りフォルダが生まれず、他端末から見ると
'   「まだ何も無い=壊れている?」と区別が付かない。
' ----------------------------------------------------------------------------
Private Sub EnsureStandardDirs(ByVal rootPath As String)
    If mDirsEnsured Then Exit Sub
    mDirsEnsured = True
    On Error Resume Next
    Dim leafs() As String: leafs = Split(modShareRule.StandardSubDirs(), "|")
    Dim i As Long
    For i = LBound(leafs) To UBound(leafs)
        Dim d As String: d = rootPath & leafs(i) & "\"
        If Len(Dir(d, vbDirectory)) = 0 Then MkDir d
        Err.Clear
    Next i
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' ReportFailure - 共有I/Oが失敗したことの受け口(2026-07-31 R8・F7最小実装)。
'   到達OKと判定した後にVPNが切れた場合、各所のリトライ(3回)を延々と
'   払い続けることになる。連続5回の失敗でこのセッションを NG へ倒し、
'   以降の共有I/Oを見送る。1回でも成功すれば連続回数は0へ戻る
'   (たまたまの1回でNGに倒さない)。
' ----------------------------------------------------------------------------
Public Sub ReportFailure(ByVal src As String)
    If mState <> ST_OK Then Exit Sub
    mFailStreak = mFailStreak + 1
    If mFailStreak < FAIL_TO_NG Then Exit Sub
    mState = ST_NG
    On Error Resume Next
    modLog.LogUsage "share_degraded", "", _
        "共有I/Oが" & FAIL_TO_NG & "回続けて失敗したため、このセッションは以後見送ります: " & _
        modUtil.SafeLeft(src, 120)
    On Error GoTo 0
End Sub

' ReportSuccess - 共有I/Oが成功したときに連続失敗回数を戻す。
Public Sub ReportSuccess()
    mFailStreak = 0
End Sub

' ----------------------------------------------------------------------------
' SubDir - <nexus_share_path>\<name>\ を返す。届かないときは空文字。
' ----------------------------------------------------------------------------
Public Function SubDir(ByVal leaf As String) As String
    If Not Reachable() Then Exit Function
    Dim p As String: p = BasePath()
    If LenB(p) = 0 Then Exit Function
    SubDir = p & leaf & "\"
End Function

' ============================================================================
' 集約スナップショット(board\summary.txt)の書式と鮮度判定(2026-08-16 R33 W6-1)
' ----------------------------------------------------------------------------
' なぜ modShare に置くか:
'   実体は「共有フォルダに置く1本のファイルの取り決め」で、書く側(発行者
'   端末)と読む側(全端末)が【同じ答え】を出さなければ集計が黙って狂う。
'   共有まわりの取り決めを1箇所へ寄せるのが冒頭「唯一性の原則」。
'   純関数なので modTestsPure34 からゴールデン固定できる(modBoard 本体は
'   Excel を触るのでテストへ載せられない)。
'
' 1行目(ヘッダ)はタブ区切り10列:
'   0 種別印 BOARD_HEAD_TAG / 1 作成時刻(ISO "yyyy-mm-dd hh:nn:ss") /
'   2 日キー / 3 今日の分 / 4 月キー / 5 今月の分 / 6 年キー / 7 今年の分 /
'   8 集計に入れた人数 / 9 打ち切り("1"=上限で切った概算)
' 2行目以降は "D<TAB>部コード<TAB>今月の分" / "T<TAB>利用者id<TAB>感謝受領数"。
' 種別印 BOARD_HEAD_TAG はモジュール先頭の宣言部にある(VBAはモジュール
' レベル宣言をプロシージャより後に置けない)。
' ============================================================================

Public Function BoardHeadText(ByVal stampText As String, ByVal dk As String, _
        ByVal dMin As Long, ByVal mk As String, ByVal mMin As Long, _
        ByVal yk As String, ByVal yMin As Long, ByVal userN As Long, _
        ByVal capped As Boolean) As String
    BoardHeadText = BOARD_HEAD_TAG & vbTab & stampText & vbTab & _
        dk & vbTab & dMin & vbTab & mk & vbTab & mMin & vbTab & _
        yk & vbTab & yMin & vbTab & userN & vbTab & IIf(capped, "1", "0")
End Function

' ヘッダの1フィールドを取り出す(範囲外・壊れた行は空文字)。
Public Function BoardHeadField(ByVal headLine As String, ByVal idx As Long) As String
    On Error GoTo Bad
    If idx < 0 Then Exit Function
    Dim f() As String: f = Split(headLine, vbTab)
    If idx > UBound(f) Then Exit Function
    BoardHeadField = Trim$(f(idx))
Bad:
End Function

' ----------------------------------------------------------------------------
' BoardHeadStatus - このスナップショットを数字として画面に出してよいか。
'   "ok" / "stale"(古すぎる) / "broken"(読めない・書きかけ)。
'
'   時刻はISO文字列のまま辞書順で比べる(CDateを通さない。和暦カレンダー
'   端末で年が化けるため。modInsight.WithinWindow と同じ作法)。
'   書きかけのファイルを読むと列が足りないので、列数と種別印の両方を見て
'   "broken" へ倒す ―― 壊れた集計を「0分」として出すと、共有が動いて
'   いないのか本当に0分なのかを利用者が区別できない。
'   書き手の時計が進んでいる場合は cutoff より後になるので "ok" 側へ倒れる
'   (時計のずれで集計を消さない)。
' ----------------------------------------------------------------------------
Public Function BoardHeadStatus(ByVal headLine As String, ByVal cutoffStamp As String) As String
    BoardHeadStatus = "broken"
    On Error GoTo Bad
    Dim f() As String: f = Split(headLine, vbTab)
    If UBound(f) < 9 Then Exit Function
    If StrComp(Trim$(f(0)), BOARD_HEAD_TAG, vbBinaryCompare) <> 0 Then Exit Function
    Dim st As String: st = Trim$(f(1))
    If Len(st) < 10 Then Exit Function
    If LenB(cutoffStamp) > 0 Then
        If st < cutoffStamp Then
            BoardHeadStatus = "stale"
            Exit Function
        End If
    End If
    BoardHeadStatus = "ok"
Bad:
End Function

' ----------------------------------------------------------------------------
' BoardHeadMin - 日/月/年の合計。スナップショットのキーと今のキーが一致する
'   ときだけ返す(昨日書かれた集計の「今日の分」を今日の数字として出さない。
'   ビーコン側と同じ日付キー方式)。壊れた値・巨大な値は0へ捨てる
'   (modBoard.SafeNum と同じ上限。1億分=約190年ぶんは現実の値ではない)。
' ----------------------------------------------------------------------------
Public Function BoardHeadMin(ByVal headLine As String, ByVal kindText As String, _
        ByVal curKey As String) As Long
    On Error GoTo Bad
    If LenB(curKey) = 0 Then Exit Function
    Dim f() As String: f = Split(headLine, vbTab)
    If UBound(f) < 9 Then Exit Function
    Dim ki As Long
    Select Case kindText
        Case "d": ki = 2
        Case "m": ki = 4
        Case "y": ki = 6
        Case Else: Exit Function
    End Select
    If StrComp(Trim$(f(ki)), curKey, vbBinaryCompare) <> 0 Then Exit Function
    Dim v As Double: v = Val(f(ki + 1))
    If v < 0 Or v > 100000000# Then Exit Function
    BoardHeadMin = CLng(v)
Bad:
End Function

' ----------------------------------------------------------------------------
' ResetProbe - 到達性の判定をやり直させる(共有パスを設定し直した直後など)。
' ----------------------------------------------------------------------------
Public Sub ResetProbe()
    mState = ST_UNKNOWN
    mFailStreak = 0
    mDirsEnsured = False
End Sub
