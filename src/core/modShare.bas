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
' 集約スナップショットのファイル名。書く側(modShare)と読む側(modBoard)で
' 名前が割れないよう、共有フォルダ側の取り決めはここが唯一の情報源。
Public Const BOARD_SUMMARY_NAME As String = "summary.txt"
' 終端行の印(2026-08-16 R33H F15)。最終行は必ず "E<TAB>データ行数"。
' 詳細は BoardEndCount の見出しコメント。
Public Const BOARD_END_TAG As String = "E"
' ビーコン名の列挙の上限(R33H F18。ファイルは開かないのでメモリの安全弁)。
Private Const BOARD_ENUM_CAP As Long = 50000

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

    If ProbePath(p, True) Then
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
' allowSlowRetry(2026-08-16 R33H F21): 冷えたVPN・スリープ復帰では1回目の
'   GetAttr が数秒かけて失敗しうる。その1回を「到達不能」で確定させると
'   共有機能がセッション丸ごと死に、W2-2 の reachableNow=False 経由で知識の
'   消去にまで連鎖する。取り違えの代償が大きい呼び口(起動時の Reachable と
'   modGuard の消去直前の再確認)だけ True を渡し、少し長めの1回目まで
'   再試行を許す。設定画面の入力検査は False のまま(利用者が目の前で
'   待っているので、速く「見つかりません」と返す方がよい)。
Public Function ProbePath(ByVal p As String, _
                          Optional ByVal allowSlowRetry As Boolean = False) As Boolean
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
        If modShareRule.ShouldRetryProbe(probeErr, firstMs, allowSlowRetry) Then
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
' 1行目(ヘッダ)はタブ区切り11列:
'   0 種別印 BOARD_HEAD_TAG / 1 作成時刻(ISO "yyyy-mm-dd hh:nn:ss") /
'   2 日キー / 3 今日の分 / 4 月キー / 5 今月の分 / 6 年キー / 7 今年の分 /
'   8 集計に入れた人数 / 9 打ち切り("1"=上限で切った概算) /
'   10 在ったビーコンの本数(=母数。R33H F18。旧10列の版は空で読まれる)
' 2行目以降は "D<TAB>部コード<TAB>今月の分" / "T<TAB>利用者id<TAB>感謝受領数"、
' 最終行は終端行 "E<TAB>データ行数"(R33H F15)。
' 種別印 BOARD_HEAD_TAG はモジュール先頭の宣言部にある(VBAはモジュール
' レベル宣言をプロシージャより後に置けない)。
' ============================================================================

' totalN は Optional(既定0)。既存の9引数呼び出しは1文字も挙動が変わらない。
Public Function BoardHeadText(ByVal stampText As String, ByVal dk As String, _
        ByVal dMin As Long, ByVal mk As String, ByVal mMin As Long, _
        ByVal yk As String, ByVal yMin As Long, ByVal userN As Long, _
        ByVal capped As Boolean, Optional ByVal totalN As Long = 0) As String
    BoardHeadText = BOARD_HEAD_TAG & vbTab & stampText & vbTab & _
        dk & vbTab & dMin & vbTab & mk & vbTab & mMin & vbTab & _
        yk & vbTab & yMin & vbTab & userN & vbTab & IIf(capped, "1", "0") & _
        vbTab & totalN
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

' 本文(ヘッダ+部別行+称号行+終端行)の組み立て。辞書の値は Long 前提
' (呼び出し側が SafeNum を通してから入れる)。キーは SanitizeId 済み=
' タブ・改行を含まない。
' R33H F15: 最後に必ず終端行 "E<TAB>データ行数" を足す。理由は BoardEndCount。
Public Function BoardBodyText(ByVal headLine As String, ByRef deptAgg As Object, _
                              ByRef titleAgg As Object) As String
    Dim s As String: s = headLine
    Dim n As Long
    Dim k As Variant
    ' R33H F22: 行の組み立ても modUtilText.BoardRowText へ寄せた(読む側の
    '   BoardRowParse と対。書式が2箇所に分かれると往復が壊れても気付けない)。
    If Not deptAgg Is Nothing Then
        For Each k In deptAgg.Keys
            s = s & vbLf & modUtilText.BoardRowText("D", CStr(k), CStr(CLng(deptAgg(k))))
            n = n + 1
        Next k
    End If
    If Not titleAgg Is Nothing Then
        For Each k In titleAgg.Keys
            s = s & vbLf & modUtilText.BoardRowText("T", CStr(k), CStr(CLng(titleAgg(k))))
            n = n + 1
        Next k
    End If
    BoardBodyText = s & vbLf & BOARD_END_TAG & vbTab & CStr(n)
End Function

' ----------------------------------------------------------------------------
' BoardEndCount - 終端行まで揃った完全なスナップショットか(2026-08-16 R33H F15)。
'   戻り値: データ行数(0以上)= 完全 / -1 = 途中までしか無い(broken)。
'
' なぜ要るか【この波の主題の1つ】:
'   置き換えに FileCopy を使っていた頃は、宛先をいったん切り詰めてから先頭
'   へ書くため、読み手は「ヘッダだけ揃った状態」を掴み得た。ヘッダは80〜100
'   バイトしかないので【部分読みの大多数はヘッダが完全】になり、列数と種別印
'   しか見ない BoardHeadStatus はそれを "ok" と判定する。結果、部の合算が0・
'   他人の称号が全滅したまま「集計時点/N名ぶん」と自信を持って表示し、
'   TTL(600秒)のあいだそれが続いた ―― 数字が嘘であることを誰も知らせない。
'   置き換え自体は rename 方式(BoardSwap)へ変えたが、共有フォルダのファイルは
'   誰にでも触れるので、読む側にも「最後まで揃っているか」の検査を必ず置く。
'   行数まで突き合わせるのは、途中の行が欠けた形も同時に弾けるため。
'   終端行より後ろは読まない(旧版の残骸が末尾に残っても影響を受けない)。
' ----------------------------------------------------------------------------
Public Function BoardEndCount(ByVal fileText As String) As Long
    BoardEndCount = -1
    On Error GoTo Bad
    Dim rows() As String: rows = Split(Replace$(fileText, vbCrLf, vbLf), vbLf)
    Dim i As Long, n As Long
    Dim declared As Long: declared = -1
    For i = 1 To UBound(rows)
        Dim c() As String: c = Split(rows(i), vbTab)
        If c(0) = BOARD_END_TAG Then
            If UBound(c) >= 1 Then declared = CLng(Val(c(1)))
            Exit For
        ElseIf UBound(c) >= 2 Then
            If c(0) = "D" Or c(0) = "T" Then n = n + 1
        End If
    Next i
    If declared < 0 Then Exit Function
    If declared <> n Then Exit Function
    BoardEndCount = n
Bad:
End Function

' ============================================================================
' ビーコンの走査範囲と称号の引き継ぎ(2026-08-16 R33H F18)
' ----------------------------------------------------------------------------
' 直したこと【この波でいちばん人に関わる欠陥】:
'   ビーコンのファイル名は stats_<不変ハッシュ>.txt で、NTFS は名前順に返す。
'   よって「先頭から SCAN_CAP 本で打ち切る」旧実装では、打ち切られる501人目
'   以降が【毎回まったく同じ顔ぶれ】になる。その人たちは組織合計に一生入らず、
'   感謝を20件集めても称号が誰の画面にも付かない ―― 「感謝ベースの称号は
'   偽装不能」という設計の売りが、ハッシュのくじ引きに化けていた。
'
' 直し方(2つ組で「恒久除外」を消す):
'   (1) 走査の起点を日ごとにずらす(BoardScanStart)。窓は capN 本ぶんで、
'       起点も1日 capN 本ずつ進むので、窓は輪をきれいに敷き詰める ――
'       ceil(本数/capN) 日で【全員が必ず1回は読まれる】(12,000人・500本なら
'       24日)。名前の列挙(ファイルを開かない)は全件やるので、母数も正確。
'   (2) 称号行は前回のスナップショットから引き継ぐ(BoardCarryTitles)。
'       (1)だけだと、その日の窓に入らなかった人の称号が1日ごとに点いたり
'       消えたりする。感謝の受領数は減らない量なので、引き継いで上書きして
'       いく形が事実に合う。結果、一巡すれば全員の称号が載り、以後は消えない。
' 残る限界(記録): 名前の列挙そのものは BOARD_ENUM_CAP 本で止める(メモリの
'   安全弁)。設計目標の12,000人に対して4倍の余裕を取ってあるが、これを
'   超える規模では再び末尾が落ちる。そこまで行ったら部署ごとに集計を割る。
' ============================================================================

' ----------------------------------------------------------------------------
' BoardScanStart - 今回の走査を何番目から始めるか【純関数】。
'   nFiles=在るビーコンの本数 / capN=1回に開く上限 / daySerial=CLng(Date)。
'   nFiles <= capN なら常に0(全部読むので回す意味が無い=従来と同じ答え)。
'   超えるときだけ、1日につき capN 本ぶん起点を進める(mod で輪にする)。
' ----------------------------------------------------------------------------
Public Function BoardScanStart(ByVal nFiles As Long, ByVal capN As Long, _
                               ByVal daySerial As Long) As Long
    If nFiles <= 0 Then Exit Function
    If capN <= 0 Then Exit Function
    If nFiles <= capN Then Exit Function
    Dim d As Long: d = daySerial
    If d < 0 Then d = -d
    BoardScanStart = ((d Mod nFiles) * capN) Mod nFiles
End Function

' ----------------------------------------------------------------------------
' BoardListBeacons - 今回開くビーコンのファイル名を vbLf 区切りで返す。
'   folderPath=board フォルダ / capN=1回に開く上限 / outTotal=在った本数(母数)。
'   collect-then-process(Dir の列挙中に他の Dir を呼ばない)。呼び出し側は
'   この関数を呼ぶ【前】に BoardCarryTitles を済ませること(あちらも Dir を
'   使うため。VBA の Dir はプロセス全体で状態を1つしか持たない)。
' ----------------------------------------------------------------------------
Public Function BoardListBeacons(ByVal folderPath As String, ByVal capN As Long, _
                                 ByRef outTotal As Long) As String
    outTotal = 0
    If LenB(folderPath) = 0 Then Exit Function
    On Error GoTo Bad
    Dim all() As String: ReDim all(0 To 63)
    Dim n As Long
    Dim fn As String: fn = Dir(folderPath & "stats_*.txt")
    Do While LenB(fn) > 0
        If n > UBound(all) Then ReDim Preserve all(0 To (UBound(all) + 1) * 2 - 1)
        all(n) = fn
        n = n + 1
        If n >= BOARD_ENUM_CAP Then Exit Do
        fn = Dir()
    Loop
    outTotal = n
    If n = 0 Then Exit Function

    Dim take As Long: take = capN
    If take <= 0 Or take > n Then take = n
    Dim st As Long: st = BoardScanStart(n, capN, CLng(Date))
    Dim sb As String
    Dim i As Long
    For i = 0 To take - 1
        If i > 0 Then sb = sb & vbLf
        sb = sb & all((st + i) Mod n)
    Next i
    BoardListBeacons = sb
Bad:
End Function

' ----------------------------------------------------------------------------
' BoardCarryTitles - 前回のスナップショットの称号行を titlesOut へ引き継ぐ。
'   戻り値は引き継いだ件数(引き継げなければ0)。完全な(終端行のある)
'   スナップショットだけを材料にする ―― 半端なファイルから称号を拾うと、
'   欠けた人の称号を「無かったこと」にして書き戻してしまう。
' ----------------------------------------------------------------------------
Public Function BoardCarryTitles(ByVal folderPath As String, ByRef titlesOut As Object) As Long
    If titlesOut Is Nothing Then Exit Function
    If LenB(folderPath) = 0 Then Exit Function
    On Error GoTo Bad
    Dim p As String: p = folderPath & BOARD_SUMMARY_NAME
    If LenB(Dir(p)) = 0 Then Exit Function
    Dim rec As String
    If Not modUtilText.ReadTextFileUtf8(p, rec) Then Exit Function
    If BoardEndCount(rec) < 0 Then Exit Function
    Dim ignored As Long
    ignored = BoardReadRows(rec, "", titlesOut)
    BoardCarryTitles = titlesOut.Count
Bad:
End Function

' ----------------------------------------------------------------------------
' BoardApproxSuffix - Hubの「みんなの節約」タイルに付ける概算の断り(R33H F18)。
'   打ち切っているのに「(概算)」としか言わないと、どれだけの人数ぶんなのかが
'   画面から分からない。母数が取れるときは「概算500/12000名」と出す。
' ----------------------------------------------------------------------------
Public Function BoardApproxSuffix(ByVal userN As Long, ByVal totalN As Long, _
                                  ByVal approx As Boolean) As String
    If Not approx Then Exit Function
    If totalN > userN And userN > 0 Then
        BoardApproxSuffix = "(概算" & userN & "/" & totalN & "名)"
    Else
        BoardApproxSuffix = "(概算)"
    End If
End Function

' ----------------------------------------------------------------------------
' BoardTextStatus - 読む側の唯一の入口(2026-08-16 R33H F15)。ファイル全文を
'   受け取り、(1)終端行まで揃っているか (2)ヘッダが使えるか の順で判定する。
'   終端行が無ければ、ヘッダがどれだけ綺麗でも "broken"(合格側へ倒さない)。
'   BoardHeadStatus は「ヘッダ1行だけ」の判定として従来どおり残す(呼ぶのは
'   ここと既存テスト)。modBoard.LoadSnapshot はこの関数だけを見て、"ok" 以外
'   なら BoardReadRows へ進まないので、壊れた本文が読まれる経路は無くなる。
' ----------------------------------------------------------------------------
Public Function BoardTextStatus(ByVal fileText As String, ByVal cutoffStamp As String) As String
    BoardTextStatus = "broken"
    If BoardEndCount(fileText) < 0 Then Exit Function
    BoardTextStatus = BoardHeadStatus(BoardHeadLine(fileText), cutoffStamp)
End Function

' 1行目(ヘッダ)だけを取り出す。CRLF/LFのどちらで書かれていても同じ答え。
Public Function BoardHeadLine(ByVal fileText As String) As String
    Dim rows() As String: rows = Split(Replace$(fileText, vbCrLf, vbLf), vbLf)
    BoardHeadLine = rows(0)
End Function

' ----------------------------------------------------------------------------
' BoardReadRows - 2行目以降を読む。称号行("T")は titlesOut(Dictionary)へ入れ、
'   戻り値は myDept の今月合計("D"行)。myDept が空なら0(部の行を出さない)。
'   数値は上限で弾く(壊れた1行で桁が跳ねない)。
' ----------------------------------------------------------------------------
Public Function BoardReadRows(ByVal fileText As String, ByVal myDept As String, _
                              ByRef titlesOut As Object) As Long
    ' R33H F22: 行の読み分けは modUtilText.BoardRowParse(純関数)へ出した。
    '   ここに残るのは「辞書へ入れる/自部署なら返す」の振り分けだけ。
    Dim rows() As String: rows = Split(Replace$(fileText, vbCrLf, vbLf), vbLf)
    Dim i As Long
    For i = 1 To UBound(rows)
        Dim p As String: p = modUtilText.BoardRowParse(rows(i))
        If LenB(p) > 0 Then
            Dim c() As String: c = Split(p, vbTab)
            If c(0) = "T" Then
                If Not titlesOut Is Nothing Then titlesOut(c(1)) = BoardNum(c(2))
            ElseIf LenB(myDept) > 0 Then
                If StrComp(c(1), myDept, vbTextCompare) = 0 Then BoardReadRows = BoardNum(c(2))
            End If
        End If
    Next i
End Function

' ----------------------------------------------------------------------------
' BoardNum - 数値欄の共通ガード。共有フォルダのファイルは誰でも書ける以上、
'   壊れた値・巨大な値・負の値が来る前提で扱う。Long の範囲を超える値で CLng が
'   オーバーフローすると集計ループがその1本で止まり、以降のビーコンが【全部】
'   欠けたまま画面に出る。1億分(=約190年ぶん)超と負値は現実の節約時間では
'   ないので0として捨てる。読めない文字列は Val が0を返すのでそのまま0。
'   R33H F17: modBoard.SafeNum が同じ判定の複製だった(コメント自身が同値だと
'   認めていた)ため、こちらへ寄せて modBoard 側を消した ―― 同じ「壊れた値の
'   線引き」が2箇所にあると、いつか必ず片方だけが更新される。
' ----------------------------------------------------------------------------
Public Function BoardNum(ByVal s As String) As Long
    On Error GoTo Bad
    Dim v As Double: v = Val(s)
    If v < 0 Or v > 100000000# Then Exit Function
    BoardNum = CLng(v)
Bad:
End Function

' ----------------------------------------------------------------------------
' BoardWriteSummary - 集約スナップショットを置き換える(2026-08-16 R33 W6-1)。
'   書きかけを他端末に読ませないため、いったん自分専用の一時ファイル
'   (summary_<自分hash>.tmp)へ書いてから上書きコピーする ―― 共有への
'   配布物を一時ファイル経由で置く modPublish と同じ作法。一時ファイル名に
'   hashが入るので複数の発行者端末が同時に書いても一時ファイルは衝突せず、
'   最後のコピーが競合しても失敗するだけで既存の summary.txt は壊れない
'   (次の集計でどちらかが書き直す)。
'   書き込み自体は AV ロックに備えて3回だけ試す(modBoard のビーコン I/O と
'   同じ考え方)。
'
'   2026-08-16(R33H F15/F16): 置き換えを FileCopy から【rename 方式】へ変え、
'   置き換えも3回試すようにした。
'   ・FileCopy は宛先を切り詰めてから先頭へ書くので、読み手が「ヘッダだけ
'     揃った壊れた集計」を掴み得た(それを "ok" と誤判定していた。BoardEndCount
'     の見出し参照)。rename は【中身の完全な1本を、その名前へ差し替える】
'     操作で、読み手が半端な内容を見る瞬間が存在しない ―― 見えるのは
'     「旧版」「新版」「(一瞬だけ)ファイルが無い」の3つだけで、無いときは
'     読む側が "none"(集計はまだありません)へ倒れる。これが原子性の根拠。
'   ・Windows の Name は【宛先が既に在ると失敗する】ので、先に宛先を .bak へ
'     退避してから tmp を宛先名へ改名し、成功したら .bak を捨てる。改名に
'     失敗したら .bak を宛先名へ戻す(旧版を失わない)。読み手が宛先を掴んで
'     いて退避に失敗した場合は、宛先が在るので次の改名も失敗し、この回は
'     何も壊さずに False で帰る(次の集計で書き直す)。
'   ・置き換えの1発勝負をやめた理由(F16): 宛先は業務時間中ほぼ常に誰かが
'     読んでおり、落とすと次の機会は TTL の10分後、24時間過ぎれば全端末が
'     「更新されていません」に倒れる。
' ----------------------------------------------------------------------------
Public Function BoardWriteSummary(ByVal folderPath As String, ByVal myHash As String, _
                                  ByVal content As String) As Boolean
    If LenB(folderPath) = 0 Or LenB(myHash) = 0 Then Exit Function
    Dim tmpPath As String: tmpPath = folderPath & "summary_" & myHash & ".tmp"
    Dim dstPath As String: dstPath = folderPath & BOARD_SUMMARY_NAME
    Dim bakPath As String: bakPath = folderPath & "summary_" & myHash & ".bak"

    Dim wroteOk As Boolean
    Dim attempt As Long
    For attempt = 1 To 3
        If modUtilText.WriteTextFileUtf8(tmpPath, content) Then
            wroteOk = True
            Exit For
        End If
    Next attempt
    If Not wroteOk Then Exit Function

    For attempt = 1 To 3
        If BoardSwap(tmpPath, dstPath, bakPath) Then
            BoardWriteSummary = True
            Exit Function
        End If
        BoardPause 200 * attempt
    Next attempt

    ' 3回とも差し替えられなかった。一時ファイルを残すと共有に増え続けるので
    ' 消す(中身は次の集計で作り直せる)。宛先は旧版のまま無傷。
    On Error Resume Next
    Kill tmpPath
    Err.Clear
    On Error GoTo 0
End Function

' BoardSwap - tmp を dst へ差し替える1回ぶん(R33H F15)。手順と根拠は
'   BoardWriteSummary の見出しコメント。壊さないことを最優先に、失敗したら
'   旧版を必ず戻す。
Private Function BoardSwap(ByVal tmpPath As String, ByVal dstPath As String, _
                           ByVal bakPath As String) As Boolean
    On Error Resume Next
    Err.Clear
    Kill bakPath              ' 前回の失敗で残った退避があれば捨てる
    Err.Clear
    Name dstPath As bakPath   ' 宛先が無ければ 53 で失敗するだけ(初回)
    Dim hadOld As Boolean: hadOld = (Err.Number = 0)
    Err.Clear
    Name tmpPath As dstPath
    Dim swapErr As Long: swapErr = Err.Number
    Err.Clear
    If swapErr <> 0 Then
        If hadOld Then Name bakPath As dstPath   ' 旧版を戻す(消したままにしない)
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    If hadOld Then Kill bakPath
    Err.Clear
    On Error GoTo 0
    BoardSwap = True
End Function

' 置き換えの再試行のあいだだけ待つ(modBoard.BoardWait と同じ作り。
' Timer は日跨ぎで0へ戻るのでその場合は即抜ける)。
Private Sub BoardPause(ByVal ms As Long)
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000# < ms
        DoEvents
        If Timer < t0 Then Exit Do
    Loop
End Sub

' ----------------------------------------------------------------------------
' BoardStateText - 集計を数字として出せないときに、その理由を利用者の言葉で
'   1文にする(BoardHeadStatus の戻り値+"none"に対応)。usage_log と画面の
'   ポップアップが【同じ文言】を使うための単一情報源: 利用者が「集計が出ない」
'   と言ってきたとき、画面に出ている文とログの文が同じでなければ、どの端末の
'   どの状態の話なのかを突き合わせられない。
' ----------------------------------------------------------------------------
' ----------------------------------------------------------------------------
' BoardOrgBlock - 「みんな(組織全体)」ブロックの本文(2026-08-16 R33 W6-1)。
'   使える集計が無いときは「0分」を並べず理由を出す ―― 共有が動いていない
'   のか本当に0分なのかを、利用者が区別できるようにするため。
'   "ok" のときは必ず「いつの・何名ぶんの・概算かどうか」を1行添える
'   (数字だけ見せて出どころを言わないのは、古い集計を黙って出すのと同罪)。
'   純関数なので modTestsPure34 が全状態を固定できる。
' ----------------------------------------------------------------------------
Public Function BoardOrgBlock(ByVal stateText As String, ByVal dayText As String, _
        ByVal monText As String, ByVal yearText As String, ByVal deptLine As String, _
        ByVal stampText As String, ByVal userN As Long, ByVal approx As Boolean, _
        ByVal capN As Long, ByVal maxAgeHours As Long) As String
    If StrComp(stateText, "ok", vbBinaryCompare) <> 0 Then
        BoardOrgBlock = "  " & BoardStateText(stateText, maxAgeHours)
        Exit Function
    End If
    BoardOrgBlock = "  今日: " & dayText & "  /  今月: " & monText & _
        "  /  今年: " & yearText & deptLine & vbLf & _
        BoardStampNote(stampText, userN, approx, capN)
End Function

' 集計の出どころの1行(BoardOrgBlock専用)。
Private Function BoardStampNote(ByVal stampText As String, ByVal userN As Long, _
                               ByVal approx As Boolean, ByVal capN As Long) As String
    BoardStampNote = "  (集計時点: " & stampText & " / " & userN & "名ぶん" & _
        IIf(approx, " / 上限" & capN & "件までの概算", "") & ")"
End Function

' BoardForceWaitText - 「更新」を下限間隔の中で押したときの1行(R33H F19)。
'   何も起きずに帰ると、ボタンが壊れているのと区別が付かない(R8 F9 の裁定)。
Public Function BoardForceWaitText(ByVal minSec As Long) As String
    BoardForceWaitText = "組織の集計は最短" & minSec & _
        "秒おきに取り直します。いまは前回の数字を表示しています。"
End Function

' ----------------------------------------------------------------------------
' BoardDeptLine - ポップアップの「部(○○)で今月 約N」1行【純関数】。
'   自分の部が分かっていて、かつ今月の合算が1分でもあるときだけ足す
'   (タイル新設はしない。R13-7c)。R13 L-batch: 60分未満を時間へ丸めると
'   「約0時間」になる ―― 1分でも貯まっているから出している行なのに「0」と
'   書くのは、事実としても労いとしても間違っている。60分未満は分のまま出す。
'   R33H F18: 容量(modBoard 残173字)のため modBoard.DeptLineForPopup の実体を
'   こちらへ移した。純関数になったので文言と丸めをテストで固定できる。
' ----------------------------------------------------------------------------
Public Function BoardDeptLine(ByVal deptCode As String, ByVal monMin As Long) As String
    If LenB(deptCode) = 0 Then Exit Function
    If monMin <= 0 Then Exit Function
    Dim amt As String
    If monMin < 60 Then
        amt = monMin & "分"
    Else
        amt = CLng(Round(monMin / 60, 0)) & "時間"
    End If
    BoardDeptLine = vbLf & "  部(" & deptCode & ")で今月 約" & amt
End Function

Public Function BoardStateText(ByVal stateText As String, ByVal maxAgeHours As Long) As String
    Select Case stateText
        Case "stale"
            BoardStateText = "組織の集計が" & maxAgeHours & _
                "時間以上更新されていないため表示していません(発行者用ブックの端末を1度起動すると作り直されます)。"
        Case "broken"
            BoardStateText = "組織の集計ファイルが読める形になっていません(書き込み中の可能性。次の更新で回復します)。"
        Case Else
            BoardStateText = "組織の集計はまだありません(発行者用ブックの端末が起動すると作られます)。"
    End Select
End Function

' ----------------------------------------------------------------------------
' ResetProbe - 到達性の判定をやり直させる(共有パスを設定し直した直後など)。
' ----------------------------------------------------------------------------
Public Sub ResetProbe()
    mState = ST_UNKNOWN
    mFailStreak = 0
    mDirsEnsured = False
End Sub
