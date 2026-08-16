Attribute VB_Name = "modShareRule"
Option Explicit

' ============================================================================
' modShareRule - 共有(P2P)まわりの「判定式」だけを集めた純ロジック(R4準拠)
' ----------------------------------------------------------------------------
' なぜ切り出すか(2026-07-31 R8):
'   P2P/共有系の実害級バグは、どれも I/O ではなく【判定式】の間違いだった。
'     ・感謝状: origin が "channel:<部門>" のとき作者を解決できず即Exit(F1)
'     ・失効:   共有へ一度も到達したことが無い端末まで30日で全消去(F4)
'     ・到達性: フォルダの「存在」ではなく「中身1件以上」を見ていた(F6)
'   いずれも実機(Windows+共有フォルダ2台)でしか踏めない形で埋まっていて、
'   本リポジトリのLOテストでは一度も触れられなかった。判定式そのものを
'   副作用ゼロの関数へ引き剥がせば、境界値をテストで固定できる。
'
'   ここに置くのは「文字列と数値だけで答えが決まるもの」に限る。
'   実際のファイル操作・シート操作は呼び出し側(modShare/modP2P/modGuard)に残す。
' ============================================================================

' GetAttr の戻り値と突き合わせるディレクトリビット。VBA定数 vbDirectory と
' 同値だが、純ロジックとして値を明示しておく(処理系差の影響を受けない)。
Public Const ATTR_DIRECTORY As Long = 16

' 到達性プローブの1回目が「構文起因で即座に弾かれた」と言える上限(ミリ秒)。
' 構文拒否はネットワークを1往復もしないので実測は0〜十数ms。一方、
' 名前解決/SMBのタイムアウトは秒単位。1秒に置けば構文拒否には桁違いの
' 余裕があり、二重タイムアウト(10〜30秒級)は確実に切り落とせる。
' 詳細は ShouldRetryProbe のコメント(R33 W4-6)。
Public Const PROBE_FAST_FAIL_MS As Long = 1000

' ----------------------------------------------------------------------------
' origin の名前空間(レビュー C-1 で "pack:" と "channel:" に分離済み)
' ----------------------------------------------------------------------------
' OriginKind - origin 文字列の種別。"pack" / "channel" / ""(自作・不明)。
Public Function OriginKind(ByVal origin As String) As String
    Dim s As String: s = LTrim$(origin)
    If LCase$(Left$(s, 5)) = "pack:" Then
        OriginKind = "pack"
    ElseIf LCase$(Left$(s, 8)) = "channel:" Then
        OriginKind = "channel"
    End If
End Function

' OriginName - origin の接頭辞を外した名前(パックなら作者表示名、
'   チャンネルなら部門名)。種別が判らないときは空文字。
Public Function OriginName(ByVal origin As String) As String
    Dim s As String: s = LTrim$(origin)
    Select Case OriginKind(origin)
        Case "pack":    OriginName = Trim$(Mid$(s, 6))
        Case "channel": OriginName = Trim$(Mid$(s, 9))
    End Select
End Function

' AuthorStatKey - 「表示名 → 発行者ID」の対応を控える my_stats のキー。
'   パック取込は "pkauth:<作者表示名>"(H-5 で導入済み)、
'   チャンネル取込は "chauth:<部門名>"(R8 F1 で追加)。
'   名前が空のときはキーを作らない(空キーは my_stats を汚すだけ)。
Public Function AuthorStatKey(ByVal origin As String) As String
    Dim nm As String: nm = OriginName(origin)
    If LenB(nm) = 0 Then Exit Function
    Select Case OriginKind(origin)
        Case "pack":    AuthorStatKey = "pkauth:" & LCase$(nm)
        Case "channel": AuthorStatKey = "chauth:" & LCase$(nm)
    End Select
End Function

' NeedsResolvedId - 発行者IDが引けなかったときに「表示名で代用してよいか」。
'   パックは代用してよい(author_id を持たない旧パックとの過渡期互換。H-5)。
'   チャンネルは代用してはいけない。部門名は人ではないので、代用すると
'   「人事部」という宛先の感謝状が永久に誰にも届かず溜まり続ける。
'   裁定 R8 F1:「発行者IDが取れない旧パックは従来どおり感謝なし(静かにExit)」。
Public Function NeedsResolvedId(ByVal origin As String) As Boolean
    NeedsResolvedId = (OriginKind(origin) = "channel")
End Function

' ----------------------------------------------------------------------------
' 到達性プローブ(R8 F6)
' ----------------------------------------------------------------------------
' ProbeTargetPath - GetAttr へ渡す形へ整える。GetAttr は末尾に "\" が付いた
'   パスを嫌う(実行時エラー52/76)ため、区切りを1つだけ落とす。ただし
'   ルート("C:\" / "\\host\share\")は "\" を落とすと別物になるので残す。
Public Function ProbeTargetPath(ByVal basePath As String) As String
    Dim p As String: p = Trim$(basePath)
    If LenB(p) = 0 Then Exit Function
    If Right$(p, 1) <> "\" Then
        ProbeTargetPath = p
        Exit Function
    End If

    Dim body As String: body = Left$(p, Len(p) - 1)
    ' "C:" だけ、あるいは "\\host\share" より浅い形になるなら落とさない。
    If Len(body) <= 2 Then
        ProbeTargetPath = p
        Exit Function
    End If
    If Left$(body, 2) = "\\" Then
        ' \\host\share までは1つの単位。区切りが2つ(= \\host\share)未満なら
        ' 末尾の "\" を落とすとホスト名だけになるので残す。
        If SepCount(Mid$(body, 3)) < 1 Then
            ProbeTargetPath = p
            Exit Function
        End If
    End If
    ProbeTargetPath = body
End Function

Private Function SepCount(ByVal s As String) As Long
    Dim i As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) = "\" Then SepCount = SepCount + 1
    Next i
End Function

' ProbeIsReachable - GetAttr の結果から到達可否を決める。
'
'   旧実装は Dir(p, vbDirectory) の戻り文字列が空かどうかを見ていた。
'   Dir にフォルダを渡すと「そのフォルダの中の最初のエントリ」を返す実装差が
'   あり、【空だが健全な共有ルート】が「届かない」と判定されていた。
'   共有ルートを作る側も全員 modShare 経由なので、誰も初期化できない
'   デッドロックになる(R8 F6)。
'   新判定は「GetAttr が成功し、かつディレクトリ属性が立っていること」。
'     ・空フォルダ  → GetAttr 成功・ディレクトリビット有り → True
'     ・不存在      → GetAttr が実行時エラー(53/76)→ errNo<>0 → False
'     ・同名ファイル → ディレクトリビットが立たない → False
Public Function ProbeIsReachable(ByVal probeErrNo As Long, ByVal attrValue As Long) As Boolean
    If probeErrNo <> 0 Then Exit Function
    ProbeIsReachable = ((attrValue And ATTR_DIRECTORY) <> 0)
End Function

' ProbeRetryPath - 1回目のプローブが失敗したときに、もう1回だけ試す形を返す。
'   もう試す価値が無ければ空文字(呼び出し側は再試行しない)。
'
' なぜ必要か(2026-07-31 R8b B10):
'   ProbeTargetPath は GetAttr のために末尾の "\" を1つ落とす。ところが
'   UNC の共有ルートを直接指した場合、
'       設定値      \\srv\share\
'       1回目の実引数 \\srv\share      ← "\" を落とした形
'   となり、Windows/ネットワークリダイレクタの組み合わせによっては
'   ここで実行時エラー52(不正なファイル名)や76(パスが見つかりません)が返る。
'   共有ルートは「\\srv\share\」の形でしか受け付けない実装があるため。
'   その結果、【共有フォルダを正しく指定しているのに永久に到達不能】になり、
'   しかも症状は「何も起きない」なので利用者にはまず原因が分からない。
'   1回目が失敗したときだけ、末尾 "\" を残した形でもう一度だけ試す。
'
'   再試行するのは「落とした結果が元と違う」ときだけ。落としていない
'   (元から "\" が無い / ルートなので残した)なら同じ実引数を2回投げるだけで、
'   遅い共有に対して OS のタイムアウトを二重に払うことになる。
Public Function ProbeRetryPath(ByVal basePath As String) As String
    Dim p As String: p = Trim$(basePath)
    If LenB(p) = 0 Then Exit Function
    If Right$(p, 1) <> "\" Then Exit Function        ' 落としていない = 再試行しない
    If StrComp(ProbeTargetPath(p), p, vbBinaryCompare) = 0 Then Exit Function
    ProbeRetryPath = p
End Function

' ShouldRetryProbe - 1回目の失敗が「構文起因」なら再試行してよい、の判定。
'
' なぜ必要か(2026-08-16 R33 W4-6):
'   ProbeRetryPath の判定材料は【パスの形】だけで、失敗の【理由】を見ない。
'   BasePath() は必ず末尾に "\" を付けて返す(modShare.bas:63)ので、UNC の
'   共有ルート "\\srv\share\" は常に「落とした形と元が違う」を満たす。
'   つまり1回目が失敗すれば必ず2回目が走る。B10 が救いたいのは
'   「末尾 "\" を落とすと構文的に受け付けられず 52/76 が【即座に】返る」
'   ケースなのに、ホスト停止・VPN未接続で1回目が長いタイムアウトの末に
'   失敗した場合も同じ経路へ入り、死んだホストへもう一度フルの
'   タイムアウトを払っていた。共有が落ちている日の起動が
'   10〜30秒ではなく20〜60秒の無反応になる ―― modShare 冒頭が
'   「解決する問題」に挙げた症状そのものを、救済コードが倍化していた。
'
' 判定:
'   構文起因の拒否はネットワークを1往復もしないので即座に返る。到達性の
'   問題は名前解決/SMBのタイムアウトを払うので秒単位になる。両者を分ける
'   のは【経過時間】であってエラー番号ではない(76 はどちらでも返りうる)。
'   よって「1回目が PROBE_FAST_FAIL_MS 以内に失敗した」ときだけ再試行する。
'   閾値は1秒 ―― 構文拒否には桁違いに余裕があり、かつ二重タイムアウト
'   (10〜30秒級)は確実に切り落とせる幅。
'   errNo は見ない。GetAttr が成功しても属性がディレクトリでなければ
'   ProbeIsReachable は偽になるが(同名ファイル)、その失敗も即座に返る
'   ので再試行のコストは無視できる。ここで errNo を絞ると、環境差で
'   別の番号が返る端末で B10 の救済が丸ごと効かなくなる方が怖い。
'   ElapsedMsSince は日跨ぎを補正して必ず0以上を返すので、負値は来ない。
Public Function ShouldRetryProbe(ByVal probeErrNo As Long, ByVal elapsedMs As Long) As Boolean
    ' probeErrNo は判定に使わない(理由は上のコメント)。エラー番号での
    ' 絞り込みを足したくなったときの受け口として引数だけ持っている。
    ' 「番号が違っても答えは経過時間だけで決まる」ことは
    ' modTestsPure34 が対で固定している。
    If elapsedMs > PROBE_FAST_FAIL_MS Then Exit Function
    ShouldRetryProbe = True
End Function

' 共有ルート直下に必ず在ってほしい標準サブフォルダ(R8 F6)。
' 「初期化の入口」を1箇所に集めるための一覧。順序は表示都合のみ。
Public Function StandardSubDirs() As String
    StandardSubDirs = "thanks|noise|insight|telemetry|board|channels|questions"
End Function

' ----------------------------------------------------------------------------
' 部門名のパス脱出検査(R33 W2-9)
' ----------------------------------------------------------------------------
' ChannelPathEscapeReason - 部門名が「フォルダの外」を指していないか。
'   使えない理由(利用者へそのまま見せる文)を返す。問題なければ空文字。
'
'   なぜ要るか: 部門名は共有フォルダの中で【部門ごとのフォルダ名】として
'   そのまま連結される(modPublish.PrepareDir が baseDir & chName & "\")。
'   W2-5 で単独の "\" を弾いたので "..\thanks" のような名前は止まるように
'   なったが、**".." 単体はまだ通る** ―― <共有>\channels\..\ は共有ルートに
'   解決するため、pack.xlsx / version.txt / publish_log.txt / publish.lock が
'   channels\ の外へ書き出される。どのチャンネル一覧にも現れないゴミが
'   共有ルートに残り、以後誰も掃除できない。
'
'   先頭・末尾の空白とピリオドも弾く: Windows はフォルダ名の先頭・末尾の
'   空白とピリオドを【黙って落とす】ので、"商品部." は "商品部" として
'   作られる。別名を入れたつもりの発行者が、気付かないまま既存部門の正典を
'   上書き発行することになる。
'
'   VBA の Trim$ は半角空白しか落とさないため、日本語IMEが出す全角空白
'   (U+3000)とタブは自分で見る。非ASCIIは ChrW で組む(CP932 の往復事故を
'   避けるための本ラウンド共通の作法)。
'
'   純関数として modShareRule に置くのは、判定をテストで固定するため
'   (ExpiryDecision を modGuard から切り出したのと同じ理由)。呼び出し側は
'   modPublishUI.OnPublish の入力検証1箇所。
' ----------------------------------------------------------------------------
Public Function ChannelPathEscapeReason(ByVal s As String) As String
    If LenB(s) = 0 Then Exit Function

    If InStr(1, s, "..") > 0 Then
        ChannelPathEscapeReason = "「..」は1つ上のフォルダを指す記号なので使えません"
        Exit Function
    End If

    If IsEdgeJunkChar(Left$(s, 1)) Then
        ChannelPathEscapeReason = "先頭の空白・タブ・ピリオドは使えません"
        Exit Function
    End If

    If IsEdgeJunkChar(Right$(s, 1)) Then
        ChannelPathEscapeReason = "末尾の空白・タブ・ピリオドは使えません"
    End If
End Function

' フォルダ名の端に置けない文字か(半角空白/全角空白/タブ/ピリオド)。
Private Function IsEdgeJunkChar(ByVal ch As String) As Boolean
    If LenB(ch) = 0 Then Exit Function
    IsEdgeJunkChar = (ch = " " Or ch = "." Or ch = vbTab Or ch = ChrW(&H3000&))
End Function

' ----------------------------------------------------------------------------
' 端末失効(R8 F4)
' ----------------------------------------------------------------------------
' ExpiryDecision - 失効タイマーが今回とるべき行動を決める。
'   戻り値:
'     "off"        … 機能無効(limitDays<=0)
'     "never"      … 一度も共有へ到達した記録が無い端末。【絶対に消さない】
'     "none"       … 何もしない(まだ猶予がある / 予告済み / 消去済み)
'     "warn"       … 残り日数つきの予告を出す
'     "warn_first" … 期限は過ぎているが予告が一度も出ていない。予告だけ出す
'     "wipe"       … 消去する
'
'   「一度も届いていない端末は消さない」が最優先(データ喪失防止)。
'   共有パスにダミー値が入っているだけ・共有をまだ作っていないPoC初期状態で
'   30日後に全端末の知識が消える、という事故を構造的に不可能にする。
'   消してよいのは「かつて届いていたのに30日届かない」端末だけ、が設計意図。
'
'   2026-07-31(司令塔裁定・R8後追い): limitDays が小さいとき(管理者が
'   knowledge_expire_days に7以下を入れたとき)の端を純ロジックで塞いだ。
'   旧実装は予告フェーズの条件が「1日以上かつ期限未満」だけだったため、
'     ・limitDays<=7 では daysSince=0(=今日共有へ届いている端末)が
'       予告フェーズに入らず、そのまま期限超過側の分岐へ落ちていた
'     ・expiry_warned_days に古い値が残っていれば、0日目でも wipe に届いた
'   つまり「今日つながっている端末の知識が消える」経路が存在した。
'   守る不変条件を3つに固定する:
'     (1) daysSince=0 では warn/warn_first/wipe を【決して】返さない
'     (2) 予告窓は max(1, limitDays-7) 〜 limitDays-1 に clamp する
'         (limitDays がいくつでも、最低1日は予告のためだけの猶予が残る)
'     (3) wipe は「daysSince >= limitDays かつ 予告済み」のときだけ。
'         予告済みフラグが古くても daysSince < limitDays なら消さない
'
'   2026-08-16(R33 W2-2・データ喪失): 不変条件を4つ目まで増やした。
'     (4) reachableNow=True(=今この瞬間、共有フォルダへ実際に届いている)なら
'         【何があっても "none"】。warn も warn_first も wipe も返さない。
'   なぜ要るか: 旧実装が受け取る「つながっているか」の材料は daysSince
'   (guard_last_reach からの経過日数=前回セッションまでの記録)だけで、
'   『今この瞬間届いているか』は入力に存在しなかった。呼び出し側の順序が
'   その穴を実害にしていた ―― 消去を実行する modGuard.EnforceExpiry は
'   Boot の前半で走るのに、今日の到達を記録する TouchReach は Boot の後半で
'   しか呼ばれない。結果、期限日以降に社内ネットワークへ復帰して開いた端末は
'   「届いているのに、届いていないことになっている」状態で判定され、
'   予告文が唯一の解除手段として案内した操作(社内NWに接続して一度開く)を
'   実行したその瞬間に本棚を消される、という裏切りが起きていた。
'   本体の修正は呼び出し側(modGuard が判定より先に到達性を確かめ、
'   届いていれば TouchReach を打つ)だが、判定式にも同じ不変条件を焼き、
'   将来また順序が入れ替わっても「つながっているのに消える」が原理的に
'   起こらないようにする(二重防御)。
'   引数は Optional。既存の5引数呼び出しは【1文字も挙動が変わらない】
'   (省略時 False = 従来どおり daysSince だけで判定する)。
Public Function ExpiryDecision(ByVal lastReachRaw As String, ByVal daysSince As Long, _
                               ByVal limitDays As Long, ByVal warnedRaw As String, _
                               ByVal wipedRaw As String, _
                               Optional ByVal reachableNow As Boolean = False) As String
    If limitDays <= 0 Then
        ExpiryDecision = "off"
        Exit Function
    End If
    If LenB(Trim$(lastReachRaw)) = 0 Then
        ExpiryDecision = "never"
        Exit Function
    End If

    ExpiryDecision = "none"

    ' (4) 今まさに共有へ届いている端末には何もしない(R33 W2-2)。
    '     予告済み・消去済みの印がどう残っていても、ここで必ず止まる。
    If reachableNow Then Exit Function

    ' (1) 今日(あるいは未来の日付)に到達している端末には何もしない。
    '     ここを通す限り「つながっているのに消えた」は起こり得ない。
    If daysSince <= 0 Then Exit Function

    ' (2) 予告窓の下限。limitDays<=7 でも 1 未満へは下げない。
    Dim warnFrom As Long
    warnFrom = limitDays - 7
    If warnFrom < 1 Then warnFrom = 1

    If daysSince < limitDays Then
        ' 期限内。ここでは【何があっても消さない】(3)。
        If daysSince < warnFrom Then Exit Function
        ' 予告フェーズ。同じ日数で繰り返し出さない(レビュー L-7)。
        If StrComp(Trim$(warnedRaw), CStr(daysSince), vbTextCompare) = 0 Then Exit Function
        ExpiryDecision = "warn"
        Exit Function
    End If

    ' 期限超過。予告を一度も出せていない(端末の時計が前へ飛んだ等)なら、
    ' まず予告だけ出して1回見送る。黙って消さない。
    If LenB(Trim$(warnedRaw)) = 0 Then
        ExpiryDecision = "warn_first"
        Exit Function
    End If
    ' 既に消したあとなら、開くたびに同じ通知を出さない。
    If LenB(Trim$(wipedRaw)) > 0 Then Exit Function
    ExpiryDecision = "wipe"
End Function

' ----------------------------------------------------------------------------
' TTLキャッシュの判定(R8 F3/F9)
' ----------------------------------------------------------------------------
' CacheIsFresh - 前回集計時刻(Timer秒)から ttlSec 以内か。
'   nowSec/lastSec は VBA の Timer(0:00からの経過秒)。日跨ぎで Timer が
'   0へ戻ると nowSec < lastSec になるので、そのときは無条件に「古い」とする
'   (24時間近く前の値を新鮮扱いする方が害が大きい)。
'   lastSec <= 0 は「まだ一度も集計していない」= 古い。
Public Function CacheIsFresh(ByVal lastSec As Double, ByVal nowSec As Double, _
                             ByVal ttlSec As Double) As Boolean
    If lastSec <= 0 Then Exit Function
    If nowSec < lastSec Then Exit Function
    CacheIsFresh = ((nowSec - lastSec) < ttlSec)
End Function

' ----------------------------------------------------------------------------
' 同期結果の報告(R8 F11)
' ----------------------------------------------------------------------------
' SyncSummaryText - チャンネル一括同期の結果文。
'   okN=0 かつ failN>0 のときに「すべて最新です」と言わない、が要点。
'   取込失敗を「最新」と報告するのは、利用者が古い正典で仕事を続ける原因になる。
Public Function SyncSummaryText(ByVal okN As Long, ByVal failN As Long, _
                                ByVal chunkN As Long, ByVal doneNames As String) As String
    If okN > 0 Then
        SyncSummaryText = okN & "部門を読み込みました(" & doneNames & " / 合計" & chunkN & "件)。"
        If failN > 0 Then
            SyncSummaryText = SyncSummaryText & vbLf & _
                failN & "部門は読み込めませんでした(発行中の可能性。少し待って再実行してください)。"
        End If
        Exit Function
    End If
    If failN > 0 Then
        SyncSummaryText = failN & "部門を読み込めませんでした" & vbLf & _
            "(発行中の可能性があります。少し待ってから、もう一度お試しください)"
        Exit Function
    End If
    SyncSummaryText = "すべて最新です。読み込み直すものはありませんでした。"
End Function

' ----------------------------------------------------------------------------
' 同時発行の見張り(R8 F12)
' ----------------------------------------------------------------------------
' PublishLockAction - publish.lock の状態から、発行を続けてよいかを決める。
'   戻り値:
'     "go"    … ロックが無い。自分がロックを取って発行してよい
'     "wait"  … 他の人が発行中(ロックが新しい)。断って案内する
'     "stale" … ロックはあるが古い / 更新時刻が未来。前回の発行が異常終了した
'               残骸とみなし、上書きして続行する
'
'   ageMinutes が負(= ロックの更新時刻が未来)のとき(2026-07-31 R8b B7b / C2):
'   当初は「判らないなら待つ」として一律 wait にしていたが、これは
'   【端末の時計が共有サーバより遅れている人が、永久に発行できなくなる】。
'   ファイルサーバと端末の時計が数分ずれるのは社内では普通にあり、
'   その端末から見ると他人のロックは常に未来の時刻に見える。
'   逆に一律 stale にすると、今度は
'   【時計が1分ずれた端末が、本当に発行中の人のロックを踏み潰す】。
'   負の側も大きさで分ける(C2):
'     ・-staleMinutes < age < 0 … 小さなズレ。「たった今誰かが作った」の方が
'       ありそうなので wait(待てば数分で解ける)
'     ・age <= -staleMinutes    … 時計が大きく狂っているか、未来の日付が
'       書かれた残骸。放置すると永久に発行できないので stale
'   どちらに倒したかは呼び出し側が err_log に残し、原因を追えるようにする。
Public Function PublishLockAction(ByVal hasLock As Boolean, ByVal ageMinutes As Double, _
                                  ByVal staleMinutes As Double) As String
    If Not hasLock Then
        PublishLockAction = "go"
        Exit Function
    End If
    If ageMinutes < 0 Then
        If ageMinutes <= -staleMinutes Then
            PublishLockAction = "stale"
        Else
            PublishLockAction = "wait"
        End If
        Exit Function
    End If
    If ageMinutes >= staleMinutes Then
        PublishLockAction = "stale"
        Exit Function
    End If
    PublishLockAction = "wait"
End Function
