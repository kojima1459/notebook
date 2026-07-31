Attribute VB_Name = "modTestsPure5"
Option Explicit

' ============================================================================
' modTestsPure5 - modTestsPure4の分割先(2026-07-31 R8: P2P/共有系の判定式)
' ----------------------------------------------------------------------------
' 役割:
'   modShareRule(共有まわりの判定式だけを集めた純ロジック)の境界値を固定する。
'   入口は modTestsPure4.RunAll4 の末尾から呼ばれる Public Sub RunAll5()。
'
' なぜここを機械で固定するのか:
'   R8のレビューで見つかった実害級の不具合は、どれもI/Oではなく【判定式】の
'   間違いだった。しかも全部「Windows + 共有フォルダ2台」でしか症状が出ない
'   形で埋まっていて、このリポジトリのテストでは一度も触れられていなかった。
'     ・F1 感謝状: origin が "channel:<部門>" のとき作者を解決できず即Exit
'     ・F4 失効:   共有へ一度も到達したことが無い端末まで30日で全消去
'     ・F6 到達性: フォルダの「存在」ではなく「中身1件以上」を見ていた
'   判定式を副作用ゼロの関数へ引き剥がした以上、ここで境界を押さえておかないと
'   同じ形の事故がまた実機で見つかることになる。
'
'   とくに ExpiryDecision は【利用者の資料を消す】唯一の判定なので、
'   「lastReach が空なら絶対に wipe を返さない」を最優先で固定する。
'   ここが壊れると、失うのは表示ではなく人の仕事の成果物になる。
' ============================================================================

' ----------------------------------------------------------------------------
' origin の名前空間(F1)
' ----------------------------------------------------------------------------
Private Sub TestOriginKind()
    modTestRunner.Check "OriginKind: pack接頭辞", _
        (modShareRule.OriginKind("pack:山田太郎") = "pack"), _
        "実際=" & modShareRule.OriginKind("pack:山田太郎")
    modTestRunner.Check "OriginKind: channel接頭辞", _
        (modShareRule.OriginKind("channel:人事部") = "channel"), _
        "実際=" & modShareRule.OriginKind("channel:人事部")

    ' 大文字小文字は問わない(config や旧パックの綴りゆれを吸収する)。
    modTestRunner.Check "OriginKind: 大文字のPACK:も pack と見なす", _
        (modShareRule.OriginKind("PACK:山田") = "pack"), ""
    modTestRunner.Check "OriginKind: 大文字のChannel:も channel と見なす", _
        (modShareRule.OriginKind("Channel:人事部") = "channel"), ""

    ' 自作(self)・空・接頭辞に似ているだけの文字列は種別なし。
    modTestRunner.Check "OriginKind: self は種別なし", _
        (LenB(modShareRule.OriginKind("self")) = 0), ""
    modTestRunner.Check "OriginKind: 空文字は種別なし", _
        (LenB(modShareRule.OriginKind("")) = 0), ""
    ' "package:" は先頭5文字が "packa" なので pack: とは一致しない。
    modTestRunner.Check "OriginKind: package: は種別なし(前方一致で誤爆しない)", _
        (LenB(modShareRule.OriginKind("package:x")) = 0), _
        "実際=" & modShareRule.OriginKind("package:x")
    modTestRunner.Check "OriginKind: channels: は種別なし", _
        (LenB(modShareRule.OriginKind("channels:x")) = 0), _
        "実際=" & modShareRule.OriginKind("channels:x")

    ' 先頭の空白は落とす(シートから読んだ値に混ざりやすい)。
    modTestRunner.Check "OriginKind: 先頭空白つきでも判定できる", _
        (modShareRule.OriginKind("  channel:営業部") = "channel"), ""
End Sub

Private Sub TestOriginName()
    modTestRunner.Check "OriginName: packは作者表示名を返す", _
        (modShareRule.OriginName("pack:山田太郎") = "山田太郎"), _
        "実際=" & modShareRule.OriginName("pack:山田太郎")
    modTestRunner.Check "OriginName: channelは部門名を返す", _
        (modShareRule.OriginName("channel:人事部") = "人事部"), _
        "実際=" & modShareRule.OriginName("channel:人事部")
    modTestRunner.Check "OriginName: 前後の空白は落とす", _
        (modShareRule.OriginName("channel:  人事部  ") = "人事部"), ""
    modTestRunner.Check "OriginName: 種別不明は空文字", _
        (LenB(modShareRule.OriginName("self")) = 0), ""
    modTestRunner.Check "OriginName: 名前が無い channel: は空文字", _
        (LenB(modShareRule.OriginName("channel:")) = 0), ""
End Sub

' AuthorStatKey は「書く側(modPack)」と「読む側(modP2P)」の唯一の合意点。
' ここがずれると感謝状が永久に届かない(症状は"何も起きない"だけ)。
Private Sub TestAuthorStatKey()
    modTestRunner.Check "AuthorStatKey: packは pkauth: + 小文字名", _
        (modShareRule.AuthorStatKey("pack:Yamada") = "pkauth:yamada"), _
        "実際=" & modShareRule.AuthorStatKey("pack:Yamada")
    modTestRunner.Check "AuthorStatKey: channelは chauth: + 小文字名", _
        (modShareRule.AuthorStatKey("channel:HR部") = "chauth:hr部"), _
        "実際=" & modShareRule.AuthorStatKey("channel:HR部")
    modTestRunner.Check "AuthorStatKey: 日本語の部門名もそのまま使える", _
        (modShareRule.AuthorStatKey("channel:人事部") = "chauth:人事部"), ""

    ' 空キーは my_stats を汚すだけなので作らない。
    modTestRunner.Check "AuthorStatKey: 名前が空ならキーを作らない", _
        (LenB(modShareRule.AuthorStatKey("channel:")) = 0), ""
    modTestRunner.Check "AuthorStatKey: 種別不明ならキーを作らない", _
        (LenB(modShareRule.AuthorStatKey("self")) = 0), ""
    modTestRunner.Check "AuthorStatKey: 空文字ならキーを作らない", _
        (LenB(modShareRule.AuthorStatKey("")) = 0), ""

    ' 書く側と読む側で同じ結果になること(F1の再発防止そのもの)。
    modTestRunner.Check "AuthorStatKey: 同じoriginなら常に同じキー", _
        (modShareRule.AuthorStatKey("channel:人事部") = _
         modShareRule.AuthorStatKey("Channel: 人事部 ")), ""
End Sub

' 部門名は人ではない。IDが引けないチャンネルで表示名を代用すると、
' 「人事部」宛の感謝状が誰にも受け取られず共有フォルダに溜まり続ける。
Private Sub TestNeedsResolvedId()
    modTestRunner.Check "NeedsResolvedId: channelはID必須(表示名で代用しない)", _
        (modShareRule.NeedsResolvedId("channel:人事部") = True), ""
    modTestRunner.Check "NeedsResolvedId: packは代用可(旧パックとの互換)", _
        (modShareRule.NeedsResolvedId("pack:山田太郎") = False), ""
    modTestRunner.Check "NeedsResolvedId: 種別不明は代用可の側へ倒す", _
        (modShareRule.NeedsResolvedId("self") = False), ""
End Sub

' ----------------------------------------------------------------------------
' 到達性プローブ(F6)
' ----------------------------------------------------------------------------
' 受入条件:「空フォルダでもOK」「不存在でNG」の根拠をここで固定する。
Private Sub TestProbeIsReachable()
    ' GetAttr が成功し、ディレクトリ属性が立っている = 到達可。
    ' 中に1件もファイルが無くても真になる(ここが旧実装との決定的な違い)。
    modTestRunner.Check "プローブ: 空フォルダ(err=0/属性16)は到達可", _
        (modShareRule.ProbeIsReachable(0, 16) = True), ""
    modTestRunner.Check "プローブ: 読み取り専用+ディレクトリ(err=0/属性17)も到達可", _
        (modShareRule.ProbeIsReachable(0, 17) = True), ""
    modTestRunner.Check "プローブ: 隠し+システム+ディレクトリ(err=0/属性22)も到達可", _
        (modShareRule.ProbeIsReachable(0, 22) = True), ""

    ' 不存在・権限なしは GetAttr が実行時エラー(53/76/75等)を出す。
    modTestRunner.Check "プローブ: 不存在(err=53)は到達不可", _
        (modShareRule.ProbeIsReachable(53, 0) = False), ""
    modTestRunner.Check "プローブ: パスが見つからない(err=76)は到達不可", _
        (modShareRule.ProbeIsReachable(76, 0) = False), ""
    ' エラー番号が立っていれば、属性値が何であれ到達不可(値は信用しない)。
    modTestRunner.Check "プローブ: err<>0なら属性16でも到達不可", _
        (modShareRule.ProbeIsReachable(70, 16) = False), ""

    ' 同名の【ファイル】が置かれている場合。属性は立つがディレクトリではない。
    modTestRunner.Check "プローブ: 通常ファイル(属性0)は到達不可", _
        (modShareRule.ProbeIsReachable(0, 0) = False), ""
    modTestRunner.Check "プローブ: 書庫属性のファイル(属性32)は到達不可", _
        (modShareRule.ProbeIsReachable(0, 32) = False), ""
End Sub

' GetAttr は末尾に "\" が付いたパスを嫌う。ただしルートの "\" は落とせない。
Private Sub TestProbeTargetPath()
    modTestRunner.Check "プローブ整形: 末尾の\を1つだけ落とす", _
        (modShareRule.ProbeTargetPath("\\srv\share\Nexus\") = "\\srv\share\Nexus"), _
        "実際=" & modShareRule.ProbeTargetPath("\\srv\share\Nexus\")
    modTestRunner.Check "プローブ整形: 末尾に\が無ければそのまま", _
        (modShareRule.ProbeTargetPath("\\srv\share\Nexus") = "\\srv\share\Nexus"), ""

    ' "C:\" から "\" を落とすと "C:"(=カレントディレクトリ)になり別物。
    modTestRunner.Check "プローブ整形: ドライブルート C:\ は落とさない", _
        (modShareRule.ProbeTargetPath("C:\") = "C:\"), _
        "実際=" & modShareRule.ProbeTargetPath("C:\")
    ' "\\host\share\" から "\" を落とすと "\\host\share"(共有ルート)で、
    ' これは GetAttr に渡して差し支えない。
    modTestRunner.Check "プローブ整形: 共有ルート \\host\share\ は落としてよい", _
        (modShareRule.ProbeTargetPath("\\host\share\") = "\\host\share"), _
        "実際=" & modShareRule.ProbeTargetPath("\\host\share\")
    ' "\\host\" まで削るとホスト名だけになり意味が変わるので残す。
    modTestRunner.Check "プローブ整形: \\host\ はホスト名だけになるので落とさない", _
        (modShareRule.ProbeTargetPath("\\host\") = "\\host\"), _
        "実際=" & modShareRule.ProbeTargetPath("\\host\")

    modTestRunner.Check "プローブ整形: 空文字は空文字", _
        (LenB(modShareRule.ProbeTargetPath("")) = 0), ""
    modTestRunner.Check "プローブ整形: 前後の空白は落とす", _
        (modShareRule.ProbeTargetPath("  \\srv\s\n\  ") = "\\srv\s\n"), _
        "実際=" & modShareRule.ProbeTargetPath("  \\srv\s\n\  ")
End Sub

' 共有ルートの初期化に使う標準サブフォルダ一覧。1箇所しか無いことが要点。
Private Sub TestStandardSubDirs()
    Dim s As String: s = modShareRule.StandardSubDirs()
    Dim parts() As String: parts = Split(s, "|")

    modTestRunner.Check "標準サブフォルダ: 空でない", (LenB(s) > 0), ""
    modTestRunner.Check "標準サブフォルダ: 7件ある", _
        ((UBound(parts) - LBound(parts) + 1) = 7), "実際=" & s

    ' 各モジュールが実際に使う名前が全部そろっていること。
    ' 1つでも欠けると、その機能だけ「フォルダが無いので静かに何も起きない」
    ' 状態になり、利用者からは壊れているのか未使用なのか区別が付かない。
    modTestRunner.Check "標準サブフォルダ: thanks(感謝状)を含む", _
        (InStr(s, "thanks") > 0), "実際=" & s
    modTestRunner.Check "標準サブフォルダ: noise(品質報告)を含む", _
        (InStr(s, "noise") > 0), ""
    modTestRunner.Check "標準サブフォルダ: insight(共有知)を含む", _
        (InStr(s, "insight") > 0), ""
    modTestRunner.Check "標準サブフォルダ: telemetry(利用状況)を含む", _
        (InStr(s, "telemetry") > 0), ""
    modTestRunner.Check "標準サブフォルダ: board(節約時間ビーコン)を含む", _
        (InStr(s, "board") > 0), ""
    modTestRunner.Check "標準サブフォルダ: channels(部門正典)を含む", _
        (InStr(s, "channels") > 0), ""
    modTestRunner.Check "標準サブフォルダ: questions(専門家への質問)を含む", _
        (InStr(s, "questions") > 0), ""

    ' 名前に空要素・パス区切りが混ざっていないこと(MkDirがそのまま使う)。
    Dim i As Long
    Dim bad As Boolean
    For i = LBound(parts) To UBound(parts)
        If LenB(Trim$(parts(i))) = 0 Then bad = True
        If InStr(parts(i), "\") > 0 Then bad = True
        If InStr(parts(i), "/") > 0 Then bad = True
    Next i
    modTestRunner.Check "標準サブフォルダ: 空要素もパス区切りも含まない", _
        (bad = False), "実際=" & s
End Sub

' ----------------------------------------------------------------------------
' 端末失効(F4): 【データ喪失に直結するので最優先で固定する】
' ----------------------------------------------------------------------------
' 最重要: guard_last_reach が空(=一度も共有へ届いた記録が無い)なら、
' 日数がいくつであろうと絶対に "wipe" を返さない。
' 共有パスにダミー値が入っているだけの端末・共有をまだ作っていないPoC初期
' 状態で、利用者が自分で取り込んだ資料まで30日後に全部消える事故を、
' 構造的に不可能にする。
Private Sub TestExpiryNeverReached()
    modTestRunner.Check "失効: 未到達(空)+0日 → never", _
        (modShareRule.ExpiryDecision("", 0, 30, "", "") = "never"), _
        "実際=" & modShareRule.ExpiryDecision("", 0, 30, "", "")
    modTestRunner.Check "失効: 未到達(空)+29日 → never", _
        (modShareRule.ExpiryDecision("", 29, 30, "", "") = "never"), ""
    modTestRunner.Check "失効: 未到達(空)+30日(期限ちょうど) → never", _
        (modShareRule.ExpiryDecision("", 30, 30, "", "") = "never"), _
        "実際=" & modShareRule.ExpiryDecision("", 30, 30, "", "")
    modTestRunner.Check "失効: 未到達(空)+9999日 → never", _
        (modShareRule.ExpiryDecision("", 9999, 30, "", "") = "never"), ""
    ' 予告済みでも、消した記録があっても、未到達なら never のまま。
    modTestRunner.Check "失効: 未到達+予告済み+9999日でも never", _
        (modShareRule.ExpiryDecision("", 9999, 30, "9999", "") = "never"), ""
    ' 空白だけの値も「記録が無い」と同じ扱い(シート由来の空セル対策)。
    modTestRunner.Check "失効: 未到達(空白のみ)でも never", _
        (modShareRule.ExpiryDecision("   ", 9999, 30, "9999", "") = "never"), ""

    ' 総当たりで「未到達なら wipe を返さない」を押さえる。
    ' 個別の条件を足すたびにここが崩れていないことを機械で確かめる。
    Dim d As Long
    Dim anyWipe As Boolean
    For d = 0 To 400
        If modShareRule.ExpiryDecision("", d, 30, "", "") = "wipe" Then anyWipe = True
        If modShareRule.ExpiryDecision("", d, 30, CStr(d), "") = "wipe" Then anyWipe = True
        If modShareRule.ExpiryDecision("", d, 1, "1", "") = "wipe" Then anyWipe = True
    Next d
    modTestRunner.Check "失効: 未到達なら0～400日のどの条件でもwipeを返さない", _
        (anyWipe = False), "未到達端末を消す分岐が復活している"
End Sub

' knowledge_expire_days<=0 は機能OFF。PoC中はこれを推奨する運用なので、
' 「0にしたのに消えた」が起きないことを固定する。
Private Sub TestExpiryOff()
    modTestRunner.Check "失効: limitDays=0 → off", _
        (modShareRule.ExpiryDecision("2026-01-01", 9999, 0, "", "") = "off"), _
        "実際=" & modShareRule.ExpiryDecision("2026-01-01", 9999, 0, "", "")
    modTestRunner.Check "失効: limitDays=-1 → off", _
        (modShareRule.ExpiryDecision("2026-01-01", 9999, -1, "", "") = "off"), ""
    ' offの判定は「未到達かどうか」より先に来る(機能を切ってあるなら
    ' そもそも何も判断しない、が素直)。
    modTestRunner.Check "失効: limitDays=0は未到達より優先して off", _
        (modShareRule.ExpiryDecision("", 9999, 0, "", "") = "off"), ""
End Sub

' 到達した記録がある端末の通常の遷移。
Private Sub TestExpiryNormalFlow()
    Dim reach As String: reach = "2026-06-01"

    ' 猶予中(期限-7日より前)は何も出さない。
    modTestRunner.Check "失効: 到達済み+0日 → none", _
        (modShareRule.ExpiryDecision(reach, 0, 30, "", "") = "none"), ""
    modTestRunner.Check "失効: 到達済み+22日(予告開始の1日前) → none", _
        (modShareRule.ExpiryDecision(reach, 22, 30, "", "") = "none"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 22, 30, "", "")

    ' 予告フェーズ(期限-7日～期限の前日)。
    modTestRunner.Check "失効: 到達済み+23日(予告開始) → warn", _
        (modShareRule.ExpiryDecision(reach, 23, 30, "", "") = "warn"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 23, 30, "", "")
    modTestRunner.Check "失効: 到達済み+29日(期限の前日) → warn", _
        (modShareRule.ExpiryDecision(reach, 29, 30, "", "") = "warn"), ""

    ' 同じ日数で二度は出さない(レビューL-7)。
    modTestRunner.Check "失効: 同じ日数で予告済みなら none", _
        (modShareRule.ExpiryDecision(reach, 29, 30, "29", "") = "none"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 29, 30, "29", "")
    modTestRunner.Check "失効: 別の日数の予告済みなら改めて warn", _
        (modShareRule.ExpiryDecision(reach, 29, 30, "28", "") = "warn"), ""

    ' 期限超過。予告が一度も出ていなければ、まず予告だけ出して1回見送る
    ' (端末の時計が前へ飛んだときに黙って消さないための最後の砦)。
    modTestRunner.Check "失効: 期限超過+予告なし → warn_first(まだ消さない)", _
        (modShareRule.ExpiryDecision(reach, 30, 30, "", "") = "warn_first"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 30, 30, "", "")
    modTestRunner.Check "失効: 大幅超過+予告なしでも warn_first", _
        (modShareRule.ExpiryDecision(reach, 999, 30, "", "") = "warn_first"), ""

    ' 予告済みで期限超過 → ここで初めて wipe。
    modTestRunner.Check "失効: 期限超過+予告済み → wipe", _
        (modShareRule.ExpiryDecision(reach, 30, 30, "29", "") = "wipe"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 30, 30, "29", "")

    ' 消したあとは黙って通す(開くたび同じ通知を出さない)。
    modTestRunner.Check "失効: 消去済みなら none", _
        (modShareRule.ExpiryDecision(reach, 40, 30, "29", "2026-07-01 09:00") = "none"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 40, 30, "29", "2026-07-01 09:00")

    ' warn_first → wipe の遷移(予告を出した翌起動で消える)を1本で示す。
    Dim first As String, second As String
    first = modShareRule.ExpiryDecision(reach, 31, 30, "", "")
    second = modShareRule.ExpiryDecision(reach, 31, 30, "31", "")
    modTestRunner.Check "失効: warn_first の次は wipe へ遷移する", _
        (first = "warn_first" And second = "wipe"), _
        "1回目=" & first & " / 2回目=" & second
End Sub

' limitDays が小さいとき(limitDays-7 が0以下になる領域)の境界。
' 管理者が knowledge_expire_days に7以下を入れると必ずここを通る。
' 守るべき不変条件は3つ(modShareRule.ExpiryDecision のコメントと対):
'   (1) daysSince=0(今日つながっている)では warn/warn_first/wipe を返さない
'   (2) 予告窓は max(1, limitDays-7) ～ limitDays-1
'   (3) wipe は daysSince >= limitDays かつ予告済みのときだけ
' ここが崩れると「つながっているのに知識が消えた」が起きる。
Private Sub TestExpirySmallLimit()
    Dim reach As String: reach = "2026-06-01"

    ' --- (1) 0日目は、どの limitDays でも・予告済みでも何もしない ---
    modTestRunner.Check "失効: limitDays=1・0日目は none", _
        (modShareRule.ExpiryDecision(reach, 0, 1, "", "") = "none"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 0, 1, "", "")
    modTestRunner.Check "失効: limitDays=2・0日目は none", _
        (modShareRule.ExpiryDecision(reach, 0, 2, "", "") = "none"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 0, 2, "", "")
    modTestRunner.Check "失効: limitDays=7・0日目は none", _
        (modShareRule.ExpiryDecision(reach, 0, 7, "", "") = "none"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 0, 7, "", "")
    modTestRunner.Check "失効: limitDays=8・0日目は none", _
        (modShareRule.ExpiryDecision(reach, 0, 8, "", "") = "none"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 0, 8, "", "")
    ' 古い予告済みフラグが残っていても、0日目は消さない(旧実装の抜け穴)。
    modTestRunner.Check "失効: limitDays=1・0日目+予告済みでも none", _
        (modShareRule.ExpiryDecision(reach, 0, 1, "99", "") = "none"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 0, 1, "99", "")

    ' 0日目の総当たり(limitDays 0～40 × 予告済み有無)で none/off だけを返す。
    Dim L As Long
    Dim badDay0 As Boolean
    Dim act0 As String
    For L = 0 To 40
        act0 = modShareRule.ExpiryDecision(reach, 0, L, "", "")
        If act0 <> "none" And act0 <> "off" Then badDay0 = True
        act0 = modShareRule.ExpiryDecision(reach, 0, L, "0", "")
        If act0 <> "none" And act0 <> "off" Then badDay0 = True
        act0 = modShareRule.ExpiryDecision(reach, 0, L, "99", "9")
        If act0 <> "none" And act0 <> "off" Then badDay0 = True
    Next L
    modTestRunner.Check "失効: 0日目はlimitDays0～40のどの条件でもnone/offのみ", _
        (badDay0 = False), "今日つながっている端末に警告か消去を返す経路がある"

    ' --- (2) 予告窓の clamp: limitDays<=8 では1日目から予告 ---
    modTestRunner.Check "失効: limitDays=2・1日目は warn", _
        (modShareRule.ExpiryDecision(reach, 1, 2, "", "") = "warn"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 1, 2, "", "")
    modTestRunner.Check "失効: limitDays=7・1日目は warn", _
        (modShareRule.ExpiryDecision(reach, 1, 7, "", "") = "warn"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 1, 7, "", "")
    modTestRunner.Check "失効: limitDays=7・6日目(期限の前日)は warn", _
        (modShareRule.ExpiryDecision(reach, 6, 7, "", "") = "warn"), ""
    modTestRunner.Check "失効: limitDays=8・1日目は warn(8-7=1が窓の下限)", _
        (modShareRule.ExpiryDecision(reach, 1, 8, "", "") = "warn"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 1, 8, "", "")
    modTestRunner.Check "失効: limitDays=3・1日目/2日目は warn", _
        (modShareRule.ExpiryDecision(reach, 1, 3, "", "") = "warn" And _
         modShareRule.ExpiryDecision(reach, 2, 3, "", "") = "warn"), ""

    ' --- (3) 期限内は予告済みフラグが古くても決して wipe しない ---
    modTestRunner.Check "失効: limitDays=7・6日目+古い予告済みでも wipe しない", _
        (modShareRule.ExpiryDecision(reach, 6, 7, "1", "") = "warn"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 6, 7, "1", "")
    modTestRunner.Check "失効: limitDays=8・7日目+古い予告済みでも wipe しない", _
        (modShareRule.ExpiryDecision(reach, 7, 8, "2", "") = "warn"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 7, 8, "2", "")

    ' 期限ちょうどで初めて消える(予告が出ていれば)。
    modTestRunner.Check "失効: limitDays=1・1日目+予告なしは warn_first", _
        (modShareRule.ExpiryDecision(reach, 1, 1, "", "") = "warn_first"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 1, 1, "", "")
    modTestRunner.Check "失効: limitDays=2・2日目+予告済みは wipe", _
        (modShareRule.ExpiryDecision(reach, 2, 2, "1", "") = "wipe"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 2, 2, "1", "")
    modTestRunner.Check "失効: limitDays=3・3日目+予告済みは wipe", _
        (modShareRule.ExpiryDecision(reach, 3, 3, "2", "") = "wipe"), ""
    modTestRunner.Check "失効: limitDays=7・7日目+予告済みは wipe", _
        (modShareRule.ExpiryDecision(reach, 7, 7, "6", "") = "wipe"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 7, 7, "6", "")
    modTestRunner.Check "失効: limitDays=8・8日目+予告済みは wipe", _
        (modShareRule.ExpiryDecision(reach, 8, 8, "7", "") = "wipe"), ""

    ' 期限内(daysSince < limitDays)からは、どの予告済み値でも wipe が出ない。
    Dim d As Long
    Dim badInside As Boolean
    For L = 1 To 40
        For d = 0 To L - 1
            If modShareRule.ExpiryDecision(reach, d, L, "999", "") = "wipe" Then badInside = True
            If modShareRule.ExpiryDecision(reach, d, L, CStr(d), "") = "wipe" Then badInside = True
        Next d
    Next L
    modTestRunner.Check "失効: 期限内(日数<limitDays)は総当たりでwipeを返さない", _
        (badInside = False), "期限前に消す経路がある"

    ' 未到達の端末は、limitDays をどれだけ小さくしても never のまま。
    modTestRunner.Check "失効: limitDays=1でも未到達なら never", _
        (modShareRule.ExpiryDecision("", 0, 1, "", "") = "never"), ""
    modTestRunner.Check "失効: limitDays=1・未到達・1日目でも never", _
        (modShareRule.ExpiryDecision("", 1, 1, "1", "") = "never"), ""
End Sub

' ----------------------------------------------------------------------------
' TTLキャッシュ(F3/F9)
' ----------------------------------------------------------------------------
Private Sub TestCacheIsFresh()
    modTestRunner.Check "TTL: 未集計(last=0)は古い", _
        (modShareRule.CacheIsFresh(0, 100, 600) = False), ""
    modTestRunner.Check "TTL: 負のlastも古い", _
        (modShareRule.CacheIsFresh(-1, 100, 600) = False), ""
    modTestRunner.Check "TTL: 直後(差0秒)は新しい", _
        (modShareRule.CacheIsFresh(100, 100, 600) = True), ""
    modTestRunner.Check "TTL: 599秒後は新しい", _
        (modShareRule.CacheIsFresh(100, 699, 600) = True), ""
    ' 境界: ちょうどTTL秒後は「古い」側へ倒す(取り直す方が安全)。
    modTestRunner.Check "TTL: ちょうど600秒後は古い", _
        (modShareRule.CacheIsFresh(100, 700, 600) = False), _
        "境界はTTL未満のみ新鮮"
    modTestRunner.Check "TTL: 601秒後は古い", _
        (modShareRule.CacheIsFresh(100, 701, 600) = False), ""
    ' 日跨ぎで Timer が0へ戻ると now < last になる。24時間近く前の値を
    ' 新鮮扱いする方が害が大きいので、無条件に古いとする。
    modTestRunner.Check "TTL: 日跨ぎ(now<last)は古い", _
        (modShareRule.CacheIsFresh(86000, 10, 600) = False), ""
    modTestRunner.Check "TTL: ttl=0なら常に古い", _
        (modShareRule.CacheIsFresh(100, 100, 0) = False), ""
End Sub

' ----------------------------------------------------------------------------
' 同期結果の報告(F11)
' ----------------------------------------------------------------------------
' 「取り込めなかった」を「すべて最新です」と報告しない、が唯一の要点。
' 古い正典のまま仕事を続けさせるのが、このアプリで最も高くつく間違い。
Private Sub TestSyncSummaryText()
    Dim s As String

    s = modShareRule.SyncSummaryText(0, 0, 0, "")
    modTestRunner.Check "同期報告: 成功0・失敗0なら『すべて最新』", _
        (InStr(s, "すべて最新") > 0), "実際=" & s

    s = modShareRule.SyncSummaryText(0, 3, 0, "")
    modTestRunner.Check "同期報告: 成功0・失敗3なら『すべて最新』と言わない", _
        (InStr(s, "すべて最新") = 0), "実際=" & s
    modTestRunner.Check "同期報告: 成功0・失敗3なら失敗件数を出す", _
        (InStr(s, "3部門") > 0), "実際=" & s
    modTestRunner.Check "同期報告: 成功0・失敗3なら再実行を促す", _
        (InStr(s, "もう一度") > 0), "実際=" & s

    s = modShareRule.SyncSummaryText(2, 0, 40, "人事部/商品部")
    modTestRunner.Check "同期報告: 成功2件は部門名と件数を出す", _
        (InStr(s, "2部門") > 0 And InStr(s, "人事部/商品部") > 0 And InStr(s, "40") > 0), _
        "実際=" & s
    modTestRunner.Check "同期報告: 成功のみなら失敗の行は出さない", _
        (InStr(s, "読み込めませんでした") = 0), "実際=" & s

    s = modShareRule.SyncSummaryText(2, 1, 40, "人事部/商品部")
    modTestRunner.Check "同期報告: 一部成功なら成功と失敗の両方を出す", _
        (InStr(s, "2部門を読み込みました") > 0 And InStr(s, "1部門は読み込めませんでした") > 0), _
        "実際=" & s
End Sub

' ----------------------------------------------------------------------------
' Ghostscript解決候補の組み立て(R9・optOcrCore.GsCandidatePaths/
' GsCandidatesForFolder)。空要素スキップ・末尾\正規化・優先順位の3点が
' ここで壊れると、実機では「config通りに書いたのに見つからない」という
' 再現しづらい不具合になる(パス文字列だけの純ロジックなので機械で固定する)。
' ----------------------------------------------------------------------------
Private Sub TestGsCandidatePaths()
    Dim s As String

    ' 優先順位: cfgPath→同梱→search_dirs の順で先頭から並ぶこと。
    s = optOcrCore.GsCandidatePaths("C:\GS\gswin32c.exe", "C:\Book", "C:\Tools\GS")
    Dim parts() As String: parts = Split(s, "|")
    modTestRunner.Check "GS候補: 1件目はcfgPathそのまま", _
        (parts(0) = "C:\GS\gswin32c.exe"), "実際=" & s
    modTestRunner.Check "GS候補: 2件目は同梱(wbDir\Ghostscript\gswin32c.exe)", _
        (parts(1) = "C:\Book\Ghostscript\gswin32c.exe"), "実際=" & s
    modTestRunner.Check "GS候補: 3件目はsearch_dirs直下", _
        (parts(2) = "C:\Tools\GS\gswin32c.exe"), "実際=" & s
    modTestRunner.Check "GS候補: 4件目はsearch_dirsのbin直下", _
        (parts(3) = "C:\Tools\GS\bin\gswin32c.exe"), "実際=" & s
    modTestRunner.Check "GS候補: 合計4件(空要素なし)", _
        ((UBound(parts) - LBound(parts) + 1) = 4), "実際=" & s

    ' cfgPathが空なら候補に加えない(空文字列の候補を作らない)。
    s = optOcrCore.GsCandidatePaths("", "C:\Book", "")
    modTestRunner.Check "GS候補: cfgPath空はスキップし同梱1件のみ", _
        (s = "C:\Book\Ghostscript\gswin32c.exe"), "実際=" & s

    ' 同梱パスの末尾\は正規化(二重\にならない)。
    s = optOcrCore.GsCandidatePaths("", "C:\Book\", "")
    modTestRunner.Check "GS候補: wbDirの末尾\は正規化", _
        (s = "C:\Book\Ghostscript\gswin32c.exe"), "実際=" & s

    ' search_dirsはセミコロン区切り。空要素(連続";"・前後空白)はスキップする。
    s = optOcrCore.GsCandidatePaths("", "", " C:\A ;;D:\B\ ")
    parts = Split(s, "|")
    modTestRunner.Check "GS候補: search_dirsの空要素はスキップされ2フォルダ×2=4件", _
        ((UBound(parts) - LBound(parts) + 1) = 4), "実際=" & s
    modTestRunner.Check "GS候補: search_dirsの前後空白は落ちる", _
        (parts(0) = "C:\A\gswin32c.exe"), "実際=" & s
    modTestRunner.Check "GS候補: search_dirsの末尾\は正規化(二重\にならない)", _
        (parts(2) = "D:\B\gswin32c.exe"), "実際=" & s

    ' 3引数すべて空なら候補ゼロ("")。
    modTestRunner.Check "GS候補: 全て空なら空文字列", _
        (LenB(optOcrCore.GsCandidatePaths("", "", "")) = 0), ""

    ' GsCandidatesForFolder単体: 直下→bin直下の順で2件。空白のみは""。
    modTestRunner.Check "フォルダ候補: 直下とbin直下の2件", _
        (optOcrCore.GsCandidatesForFolder("D:\GS") = "D:\GS\gswin32c.exe|D:\GS\bin\gswin32c.exe"), ""
    modTestRunner.Check "フォルダ候補: 末尾\は正規化", _
        (optOcrCore.GsCandidatesForFolder("D:\GS\") = "D:\GS\gswin32c.exe|D:\GS\bin\gswin32c.exe"), ""
    modTestRunner.Check "フォルダ候補: 空白のみは空文字列", _
        (LenB(optOcrCore.GsCandidatesForFolder("   ")) = 0), ""
End Sub

Public Sub RunAll5()
    On Error GoTo OriginKindFail
    TestOriginKind
NextOriginName:
    On Error GoTo OriginNameFail
    TestOriginName
NextAuthorKey:
    On Error GoTo AuthorKeyFail
    TestAuthorStatKey
NextNeedsId:
    On Error GoTo NeedsIdFail
    TestNeedsResolvedId
NextProbe:
    On Error GoTo ProbeFail
    TestProbeIsReachable
NextProbePath:
    On Error GoTo ProbePathFail
    TestProbeTargetPath
NextSubDirs:
    On Error GoTo SubDirsFail
    TestStandardSubDirs
NextExpNever:
    On Error GoTo ExpNeverFail
    TestExpiryNeverReached
NextExpOff:
    On Error GoTo ExpOffFail
    TestExpiryOff
NextExpFlow:
    On Error GoTo ExpFlowFail
    TestExpiryNormalFlow
NextExpSmall:
    On Error GoTo ExpSmallFail
    TestExpirySmallLimit
NextCache:
    On Error GoTo CacheFail
    TestCacheIsFresh
NextSyncSum:
    On Error GoTo SyncSumFail
    TestSyncSummaryText
NextGsCand:
    On Error GoTo GsCandFail
    TestGsCandidatePaths
NextPure6:
    ' 2026-07-31 R8b: 敵対的レビュー対応(B1/B7b/B10)のテストは
    ' modTestsPure6 へ置いた。ここが唯一の導線なので消さないこと。
    On Error GoTo Pure6Fail
    modTestsPure6.RunAll6
NextDone5:
    On Error GoTo 0
    Exit Sub

OriginKindFail:
    modTestRunner.Check "TestOriginKind(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextOriginName
OriginNameFail:
    modTestRunner.Check "TestOriginName(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextAuthorKey
AuthorKeyFail:
    modTestRunner.Check "TestAuthorStatKey(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextNeedsId
NeedsIdFail:
    modTestRunner.Check "TestNeedsResolvedId(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextProbe
ProbeFail:
    modTestRunner.Check "TestProbeIsReachable(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextProbePath
ProbePathFail:
    modTestRunner.Check "TestProbeTargetPath(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextSubDirs
SubDirsFail:
    modTestRunner.Check "TestStandardSubDirs(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextExpNever
ExpNeverFail:
    modTestRunner.Check "TestExpiryNeverReached(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextExpOff
ExpOffFail:
    modTestRunner.Check "TestExpiryOff(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextExpFlow
ExpFlowFail:
    modTestRunner.Check "TestExpiryNormalFlow(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextExpSmall
ExpSmallFail:
    modTestRunner.Check "TestExpirySmallLimit(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextCache
CacheFail:
    modTestRunner.Check "TestCacheIsFresh(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextSyncSum
SyncSumFail:
    modTestRunner.Check "TestSyncSummaryText(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGsCand
GsCandFail:
    modTestRunner.Check "TestGsCandidatePaths(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextPure6
Pure6Fail:
    modTestRunner.Check "modTestsPure6.RunAll6(モジュール全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone5
End Sub
