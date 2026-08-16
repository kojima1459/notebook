Attribute VB_Name = "modGuard"
Option Explicit

' ドメイン不一致の連続回数を控えるmy_statsのキー(レビュー M-6)。
Private Const STREAK_KEY As String = "domain_block_streak"
' 失効の予告を出した日数(同じ日数で繰り返し出さないため)と、
' 実際に消した記録(消したあと毎回通知しないため)。レビュー L-7。
Private Const WARNED_KEY As String = "expiry_warned_days"
Private Const WIPED_KEY As String = "expiry_wiped_at"
' 「一度も共有へ到達していないので失効判定を見送った」ことを usage_log へ
' 1行だけ残すための印(R8 F4)。毎起動で同じ行を積まないための重複防止。
Private Const NEVER_LOGGED_KEY As String = "expiry_never_reached_logged"

' ============================================================================
' modGuard - 端末セキュリティ(PC紛失・持ち出し対策)
' ----------------------------------------------------------------------------
' 先に限界を明記しておく:
'   xlsm という配布形態である以上、完全な防御は不可能。VBAプロジェクトを
'   開ける人にはブック暗号化も無意味で、「守れている」と言い切るのは嘘になる。
'   ここでできるのは次の2つだけであり、それ以上を約束しない。
'     ・拾った人が中身を読むまでの手間を上げる
'     ・時間が経てば自動的に空になる(被害の時間的な上限を作る)
'
' 実装する2層:
'   1) 所属チェック
'      起動時に Windows のユーザードメインを確認する。config allowed_domain
'      に値が入っていて、かつ一致しない端末では知識シートを一切開かず案内
'      だけ出す。私物PCへコピーされた場合の一次防御。
'
'   2) 有効期限(kill switch)
'      共有フォルダに最後に到達できた日から config knowledge_expire_days
'      (出荷既定0=無効。運用で日数を入れて初めて有効になる)が過ぎたら、
'      my_knowledge / my_vectors の中身を消す。
'      社内ネットワークから切り離された端末は、放っておけば空になる。
'      「盗まれた瞬間に守る」のではなく「持ち出し続けさせない」設計。
'
' 誤爆させないための配慮(ここを間違えると業務が止まる):
'   ・長期休暇・出張・在宅で30日超えることは普通にある。消す前に必ず警告を
'     出し、残り日数を見せる(7日前から)。黙って消すことは絶対にしない。
'   ・allowed_domain が空(既定)なら所属チェックは行わない。設定した組織
'     だけが有効になる。初期状態で誰かの業務を止める作りにはしない。
'   ・消えるのは取り込んだ知識だけ。設定・統計・履歴は残す。共有フォルダに
'     つながればチャンネルから再取得できるので、実害は再同期の待ち時間だけ。
'   ・【一度も共有フォルダへ到達した記録が無い端末は絶対に消さない】
'     (2026-07-31 R8 F4)。共有パスにダミー値が入っているだけ、共有をまだ
'     作っていない、という状態で30日経つと全端末の知識が消えていた。
'     「つながらなくなった」と「そもそも一度もつながっていない」は別物で、
'     後者はこの機能が想定している脅威(持ち出された端末)ではない。
'     判定は modShareRule.ExpiryDecision に切り出し、境界値をテストで固定した。
'   ・knowledge_expire_days は config から変更できる(0で無効。PoC中は0推奨)。
' ============================================================================

Private Const K_LAST_REACH As String = "guard_last_reach"   ' 共有フォルダ最終到達日

' ----------------------------------------------------------------------------
' CheckDomain - 所属チェック。使ってよい端末なら True。
'   config allowed_domain が空なら常に True(既定では無効)。
' ----------------------------------------------------------------------------
Public Function CheckDomain() As Boolean
    CheckDomain = True
    On Error Resume Next
    Dim allowed As String
    allowed = Trim$(modConfig.GetString("allowed_domain", ""))
    If LenB(allowed) = 0 Then Exit Function

    ' 2026-07-28(レビュー M-6): 判定を2変数の突き合わせにした。
    ' 従来は USERDOMAIN(NetBIOS名)しか実質見ておらず、FQDN で
    ' allowed_domain を設定した組織や Azure AD 参加端末では
    ' 【正規の端末が全滅】する。不一致の帰結が予告なしの即ワイプなので、
    ' 誤判定のコストが極端に高い。どちらかが一致すれば許可する。
    Dim domNb As String: domNb = Trim$(Environ$("USERDOMAIN"))
    Dim domDns As String: domDns = Trim$(Environ$("USERDNSDOMAIN"))

    ' 2026-07-31(R8b B6): ドメイン名が【両方とも取れない】ときは
    ' 「不一致」ではなく「シグナル無し」として扱い、許可する。
    ' 従来はここで False を返していたが、不一致の帰結は数回の観測で
    ' 知識の消去(DomainBlockShouldWipe)である。環境変数が取れない状況は
    '   ・ワークグループ端末(社内でも一定数ある)
    '   ・サービス/タスクスケジューラ起動でユーザー環境が薄い
    '   ・プロファイル読み込み前・一時的な環境変数の欠落
    ' など、【所属が違うこと以外の理由】で普通に起こる。
    ' 「持ち出された端末を止める」という目的に対し、情報が無いことを
    ' 有罪の証拠にしてはいけない。実際に別ドメイン名が取れたときだけ
    ' 不一致とみなす(その分岐は下の突き合わせで従来どおり動く)。
    If LenB(domNb) = 0 And LenB(domDns) = 0 Then Exit Function

    ' 複数許可はカンマ区切り。config は人が手で書くため、半角/全角の
    ' スペースとカンマの表記ゆれを正規化してから比較する
    ' (全角カンマ1つで全社員が締め出される、という事故を作らない)。
    Dim allowNorm As String: allowNorm = "," & NormalizeDomainList(allowed) & ","
    CheckDomain = MatchesDomain(allowNorm, domNb) Or MatchesDomain(allowNorm, domDns)
    On Error GoTo 0
End Function

' allowed_domain の表記ゆれを吸収する(全角カンマ/全角スペース/半角スペース)。
Private Function NormalizeDomainList(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace(t, ChrW(&HFF0C), ",")      ' 全角カンマ
    t = Replace(t, ChrW(&H3001), ",")      ' 読点
    t = Replace(t, ChrW(&H3000), "")       ' 全角スペース
    t = Replace(t, " ", "")
    t = Replace(t, vbTab, "")
    NormalizeDomainList = t
End Function

' 許可リスト(前後をカンマで囲った文字列)に dom が含まれるか。
' FQDN の先頭ラベル("CORP.EXAMPLE.CO.JP" に対する "CORP")でも一致させる。
Private Function MatchesDomain(ByVal allowNorm As String, ByVal dom As String) As Boolean
    Dim d As String: d = NormalizeDomainList(dom)
    If LenB(d) = 0 Then Exit Function
    If InStr(1, allowNorm, "," & d & ",", vbTextCompare) > 0 Then
        MatchesDomain = True
        Exit Function
    End If
    Dim p As Long: p = InStr(d, ".")
    If p > 1 Then
        MatchesDomain = (InStr(1, allowNorm, "," & Left$(d, p - 1) & ",", vbTextCompare) > 0)
    End If
End Function

' 所属チェックに落ちたときの案内(起動を止めるのではなく、知識を出さない)。
Public Sub ShowDomainBlocked()
    On Error Resume Next
    MsgBox "この端末では社内ナレッジを利用できません。" & vbCrLf & vbCrLf & _
           "会社から貸与された端末で、社内ネットワークにログインした状態で" & vbCrLf & _
           "開いてください。" & vbCrLf & vbCrLf & _
           "(このファイルには社内の業務知識が含まれるため、" & vbCrLf & _
           " 許可された端末以外では内容を表示しない設定になっています)", _
           vbExclamation, modAppDef.APP_NAME
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' TouchReach - 共有フォルダに到達できたことを記録する。
'   チャンネルの版確認など、共有フォルダを触る処理が成功したときに呼ぶ。
' ----------------------------------------------------------------------------
Public Sub TouchReach()
    On Error Resume Next
    modStats.SetStatText K_LAST_REACH, modUtilText.IsoDate(Date)
    ' つながったので失効まわりの記録はリセットする(レビュー L-7)。
    ' これをしないと、復帰後も「予告済み」「消去済み」の印が残り、
    ' 次に離れたときの予告が出なくなる。
    If LenB(modStats.GetStatText(WARNED_KEY)) > 0 Then modStats.SetStatText WARNED_KEY, ""
    If LenB(modStats.GetStatText(WIPED_KEY)) > 0 Then modStats.SetStatText WIPED_KEY, ""
    On Error GoTo 0
End Sub

' 最終到達日からの経過日数(記録が無ければ0=今日として扱い、誤爆を防ぐ)。
'
' 2026-07-31(レビュー R8 F4): 記録が無いときに TouchReach を呼んで
' 【到達したことにしていた】のをやめた。共有フォルダへ一度も届いていない
' 端末が「今日届いた」ことになり、そこから30日を数え始めてしまう。
' 失効の最終防壁は「一度も到達した記録が無い端末は消さない」なので、
' その記録を勝手に作ってしまうと防壁ごと無効になる。
' 記録が無ければ 0 を返すだけにし、書き込みは実際に到達した TouchReach
' だけが行う(EnforceExpiry は LastReachRaw() の空判定を先に見る)。
Public Function DaysSinceReach() As Long
    On Error Resume Next
    Dim s As String: s = LastReachRaw()
    If LenB(s) = 0 Then Exit Function
    ' 2026-07-31(R8b B13): CLng は【四捨五入】する。guard_last_reach に
    ' 時刻付きの値が入っていると Date - CDate(s) が例えば 29.6 になり、
    ' CLng だと 30 へ切り上がって【1日早く消える】。経過日数は切り捨てが正しい
    ' (「まだ29日と14時間」を30日経過とは呼ばない)。
    DaysSinceReach = CLng(Int(Date - CDate(s)))
    If DaysSinceReach < 0 Then DaysSinceReach = 0
    On Error GoTo 0
End Function

' 共有フォルダへ最後に到達した日の生値(未到達なら空文字)。
Public Function LastReachRaw() As String
    On Error Resume Next
    LastReachRaw = Trim$(modStats.GetStatText(K_LAST_REACH))
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' EnforceExpiry - 期限切れなら知識を消す。7日前から警告を出す。
'   戻り値 True = 消去した。
' ----------------------------------------------------------------------------
Public Function EnforceExpiry() As Boolean
    On Error Resume Next
    Dim limitDays As Long
    ' 2026-08-01(R12-1-1): 既定値は 0(=失効させない)でなければならない。
    ' ここが 30 だと、config の knowledge_expire_days セルが空になった/文字列に
    ' なった/シートが読めなかった、というだけで「30日で全知識を消すタイマー」が
    ' 勝手に起動する。出荷既定は 0(build_mybookshelf.py)なので、読めなかった
    ' ときに出荷既定より危険な側へ倒れるのは逆向きの既定。読めない時は
    ' 「何もしない」へ倒す(憲章§3-5: データ保全は全機能に優先する)。
    limitDays = modConfig.GetLong("knowledge_expire_days", 0)

    ' 共有フォルダを使っていない環境ではこの機能自体を動かさない
    ' (単独利用の人の知識を、つながる先が無いという理由で消すのは理不尽)。
    If LenB(Trim$(modConfig.GetString("nexus_share_path", ""))) = 0 Then Exit Function

    ' 判定式そのものは純ロジック(modShareRule.ExpiryDecision)へ切り出した。
    ' 境界値をテストで固定するため、およびここでの分岐を読める長さに保つため。
    ' 2026-07-28(レビュー L-7): 3つの雑さを直してある。
    '  (a) expire_days<=7 だと limitDays-7<=0 になり、初日から毎回警告が出る
    '      → 「1日でも猶予がある」ときだけ予告する。
    '  (b) 期限超過後は開くたびにワイプ通知が出ていた
    '      → 一度消したら記録し、二度目からは黙って通す(もう消すものが無い)。
    '  (c) 端末の時計が前へ飛ぶと、予告を一度も出さずにいきなり消えていた
    '      → 予告を出した記録が無ければ、まず予告だけ出して1回見送る。
    ' ------------------------------------------------------------------------
    ' 2026-08-16(R33 W2-2・データ喪失): 【消去を判定する前に、今つながって
    ' いるかを確かめる】。
    '
    ' ここに到達性の確認が無かったせいで、次の事故が成立していた:
    '   期限日以降に社内ネットワークへ復帰して開いた端末が、共有フォルダへ
    '   実際に届いているのに本棚を全消去される。判定材料の guard_last_reach を
    '   今日へ更新する TouchReach は Boot の【後半】でしか呼ばれず、この関数が
    '   走る時点の記録は「前回セッション=社外にいたとき」のままだったため。
    '   予告文(下の warn/warn_first)が唯一の解除手段として案内しているのが
    '   「社内ネットワークに接続して一度開く」ことなので、案内どおりに操作した
    '   利用者が、その操作の瞬間に資産を失う形になっていた(憲章§3-5違反)。
    '
    ' 直し方は順序の入れ替え1本: 判定より先に到達性を見て、届いていれば
    ' 先に TouchReach を打つ(guard_last_reach=今日 / 予告・消去の印もリセット)。
    ' その後で読む lastReach / d は「今日・0日」になるので、判定は自然に
    ' "none" へ落ちる。判定式側にも同じ不変条件を入れてある(二重防御。
    ' modShareRule.ExpiryDecision の不変条件(4))。
    '
    ' 【本来消すべきケースは従来どおり消える】: Reachable() が False の端末
    ' (本当に届かない=社外のまま期限超過)では reachableNow=False となり、
    ' TouchReach も打たれないので、材料も判定も従来と1ビットも変わらない。
    ' Reachable() は1セッション1回のキャッシュ(modShare.mState)なので、
    ' Boot 後半の呼び出しと合わせてもプローブは1回のまま増えない。
    ' なお、この関数は nexus_share_path が空なら既に Exit しているので、
    ' 共有を使っていない端末がここでプローブの待ち時間を払うことは無い。
    ' ------------------------------------------------------------------------
    Err.Clear
    Dim reachableNow As Boolean
    reachableNow = modShare.Reachable()
    If Err.Number <> 0 Then
        Dim probeErrNum As Long: probeErrNum = Err.Number
        Dim probeErrDesc As String: probeErrDesc = Err.Description
        Err.Clear
        reachableNow = False
        modLog.LogUsage "guard_reach_probe_fail", "", _
            "失効判定の前に共有到達性を確認できませんでした(err#" & probeErrNum & _
            " " & probeErrDesc & ")。到達性は未確認として判定します。"
        Err.Clear
    End If
    If reachableNow Then TouchReach

    Dim lastReach As String: lastReach = LastReachRaw()
    Dim d As Long: d = DaysSinceReach()
    Err.Clear
    Dim act As String
    act = modShareRule.ExpiryDecision(lastReach, d, limitDays, _
                                      modStats.GetStatText(WARNED_KEY), _
                                      modStats.GetStatText(WIPED_KEY), _
                                      reachableNow)
    ' 2026-07-31(R11-A): 判定材料(config・my_stats・判定式)の取得が失敗しても、
    ' この関数は冒頭の On Error Resume Next で黙って先へ進む。act は空のまま
    ' になり、下のホワイトリスト(act<>"wipe" なら何もしない)で消去は
    ' 防がれるが、【なぜ判定できなかったのか】がどこにも残らない。
    ' 失効判定はデータ喪失に直結する処理なので、材料が欠けたことは必ず残す。
    If Err.Number <> 0 Then
        modLog.LogUsage "guard_material_fail", "", _
            "失効判定の材料を取得できませんでした(err#" & Err.Number & " " & _
            Err.Description & ")。判定結果=" & act
        Err.Clear
    End If

    ' 2026-07-31(レビュー R8 F4・データ喪失): 【一度も共有へ到達した記録が
    ' 無い端末は、絶対に消さない】。
    ' 共有パスにダミー値が入っているだけの端末・共有フォルダをまだ作って
    ' いないPoC初期状態では、guard_last_reach が空のまま日数だけが進む。
    ' 旧実装はそれを「30日つながっていない」と解釈して、部門の正典どころか
    ' 【利用者が自分で取り込んだ資料まで】全消去していた。
    ' 消してよいのは「かつて届いていたのに30日届かない」端末だけ、が設計意図。
    ' 予告も出さない(まだ一度も繋がっていない人に「消えます」と脅す意味が無い)。
    ' 見送ったことは usage_log に1行だけ残し、原因調査の手掛かりにする。
    If act = "never" Then
        If LenB(modStats.GetStatText(NEVER_LOGGED_KEY)) = 0 Then
            modStats.SetStatText NEVER_LOGGED_KEY, modUtil.NowStamp()
            modLog.LogUsage "guard_expiry_skipped", "", _
                "共有フォルダへ一度も到達した記録が無いため、失効判定を行いません" & _
                "(guard_last_reach が空)。共有パスの設定・到達性をご確認ください。"
        End If
        Exit Function
    End If
    If act = "off" Then Exit Function        ' knowledge_expire_days<=0 で機能OFF
    If act = "none" Then Exit Function

    If act = "warn" Then
        modStats.SetStatText WARNED_KEY, CStr(d)
        ' 予告。黙って消さない。
        MsgBox "社内ネットワークに " & d & " 日間つながっていません。" & vbCrLf & _
               "あと " & (limitDays - d) & " 日つながらないと、" & vbCrLf & _
               "安全のため取り込んだ知識を自動的に消去します。" & vbCrLf & vbCrLf & _
               "社内ネットワークに接続して一度開いていただければ解除されます。" & vbCrLf & _
               "(設定や統計は消えません。知識は接続後に自動で戻せます)", _
               vbExclamation, modAppDef.APP_NAME
        Exit Function
    End If

    If act = "warn_first" Then
        modStats.SetStatText WARNED_KEY, CStr(d)
        MsgBox "社内ネットワークに " & d & " 日間つながっていません。" & vbCrLf & _
               "次に開いたときも接続できていない場合、安全のため" & vbCrLf & _
               "取り込んだ知識を消去します。" & vbCrLf & vbCrLf & _
               "社内ネットワークに接続して一度開いていただければ解除されます。", _
               vbExclamation, modAppDef.APP_NAME
        Exit Function
    End If

    ' ------------------------------------------------------------------------
    ' 2026-07-31(R8b B1・最重要/データ喪失封鎖): 消去は【ホワイトリスト】で守る。
    '
    ' ここまでの分岐は "never"/"off"/"none"/"warn"/"warn_first" を1つずつ
    ' 弾く【ブラックリスト】で、最後に残ったものを消していた。つまり
    ' 既定の振る舞いが「消す」になっている。この関数の冒頭は
    ' On Error Resume Next なので、
    '   ・my_stats のセルがエラー値(#REF! 等)で GetStatText が失敗する
    '   ・modShareRule の呼び出しが何らかの理由で失敗する
    ' といったときに act が空文字のままここへ落ち、【利用者の資料が消える】。
    ' 判定式がどれだけ正しくても、呼び出し側の既定が「消す」ならいつか消える。
    ' 消してよいのは act が明示的に "wipe" のときだけ、と裏返す。
    If act <> "wipe" Then Exit Function
    ' ------------------------------------------------------------------------

    ' 知識だけを消す。
    Dim removed As Long
    Dim leftRows As Long
    removed = WipeKnowledge(leftRows)
    modStats.SetStatText WIPED_KEY, modUtil.NowStamp()
    modLog.LogUsage "guard_wipe", "", "days=" & d & " removed=" & removed & " left=" & leftRows
    If leftRows > 0 Then
        ' 消し切れていないのに「消去しました」と言わない(R11-A)。
        MsgBox "社内ネットワークに " & d & " 日間つながらなかったため、" & vbCrLf & _
               "取り込んだ知識の消去を行いましたが、" & vbCrLf & _
               "一部の資料を消去できませんでした。" & vbCrLf & vbCrLf & _
               "ファイルを開き直してからもう一度お試しください。" & vbCrLf & _
               "それでも残る場合は、このツールの管理担当者にご連絡ください。", _
               vbExclamation, modAppDef.APP_NAME
    Else
        MsgBox "社内ネットワークに " & d & " 日間つながらなかったため、" & vbCrLf & _
               "安全のため取り込んだ知識を消去しました。" & vbCrLf & vbCrLf & _
               "社内ネットワークに接続して開き直すと、部門チャンネルから" & vbCrLf & _
               "自動的に取り込み直せます。設定・統計・履歴は残っています。", _
               vbInformation, modAppDef.APP_NAME
    End If
    EnforceExpiry = True
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' WipeKnowledge - my_knowledge / my_vectors の中身を消す(見出し行は残す)。
'   チャンネルの取り込み済み版番号も消す。次回接続時に再取得させるため。

' ----------------------------------------------------------------------------
' DomainBlockShouldWipe - ドメイン不一致を何回続けて観測したら消すか。
'
' 2026-07-28(レビュー M-6): 端末失効(EnforceExpiry)は7日前から予告するのに、
' ドメイン不一致は【予告なしの即ワイプ】だった。しかも判定は環境変数1つに
' 依存していて、VPN未接続・一時的なプロファイル不整合・FQDN設定の違いでも
' 不一致になり得る。H-8(ワイプ後に自分の資料が復元されない)と重なると
' 実質不可逆だった。
'
' 1回目は表示ブロックだけにして、連続して規定回数に達したときだけ消す。
' 一度でも一致したらカウンタは0へ戻す(たまたまの1回で近づかない)。
' 既定3回。config domain_wipe_after_n_boots で調整でき、0以下なら
' 「絶対に消さない(表示ブロックのみ)」になる。
' ----------------------------------------------------------------------------
Public Function DomainBlockShouldWipe() As Boolean
    On Error Resume Next
    Dim threshold As Long
    threshold = modConfig.GetLong("domain_wipe_after_n_boots", 3)

    ' 数値統計(Bump)は減算・リセットの口が無いので、文字列統計を使う。
    Dim n As Long
    n = CLngSafeStat(modStats.GetStatText(STREAK_KEY)) + 1
    modStats.SetStatText STREAK_KEY, CStr(n)

    If threshold <= 0 Then Exit Function
    DomainBlockShouldWipe = (n >= threshold)
    On Error GoTo 0
End Function

' ドメイン確認に通ったので、不一致の連続回数をリセットする。
Private Function CLngSafeStat(ByVal s As String) As Long
    On Error Resume Next
    If IsNumeric(s) Then CLngSafeStat = CLng(Val(s))
    On Error GoTo 0
End Function

Public Sub ClearDomainBlockStreak()
    On Error Resume Next
    If LenB(modStats.GetStatText(STREAK_KEY)) > 0 Then
        modStats.SetStatText STREAK_KEY, "0"
    End If
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
Public Function WipeKnowledge(Optional ByRef outLeftRows As Long = 0) As Long
    outLeftRows = 0
    On Error Resume Next
    Dim wsK As Worksheet, wsV As Worksheet
    Set wsK = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    Set wsV = ThisWorkbook.Worksheets(modAppDef.SH_VECTORS)

    If Not wsK Is Nothing Then
        Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.Count, 1).End(xlUp).Row
        If lastK >= 2 Then
            WipeKnowledge = lastK - 1
            wsK.Rows("2:" & lastK).Delete
        End If
    End If
    If Not wsV Is Nothing Then
        Dim lastV As Long: lastV = wsV.Cells(wsV.Rows.Count, 1).End(xlUp).Row
        If lastV >= 2 Then wsV.Rows("2:" & lastV).Delete
    End If
    ' R12-H-6: 失効ワイプで全ベクトルが消えた。セッション内キャッシュ
    ' (32bit Excelで最大126MB)を抱えたままにする理由はもう無いので解放し、
    ' 世代も進めて「消したはずの本棚で検索できる」状態を作らない。
    modVecCache.ResetVecCache
    modVecCache.BumpGeneration

    ' 取り込み済み版の記録を消し、次回接続時に全チャンネルを取り直させる。
    Dim wsS As Worksheet
    Set wsS = ThisWorkbook.Worksheets(modAppDef.SH_STATS)
    If Not wsS Is Nothing Then
        Dim lastS As Long: lastS = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row
        Dim r As Long
        For r = lastS To 2 Step -1
            If Left$(CStr(wsS.Cells(r, 1).Value), 3) = "ch:" Then wsS.Rows(r).Delete
        Next r
    End If

    ' 2026-07-28(レビュー H-8): my_manifest も status="failed" に倒す。
    '
    ' チャンネル分は上の ch: 版クリアで next sync に復元されるが、
    ' 自分で取り込んだ資料は manifest が status="done" のまま残るため、
    ' 次回の同期が「変化なし=keep」でスキップし、shelf_folder が健在でも
    ' 二度と自動復元されなかった。失効ダイアログは
    ' 「接続して開き直すと自動的に取り込み直せます」と案内しているのに、
    ' 実装がそうなっていなかった。本棚UIはカードを出し続けるのに検索は
    ' 0件、という説明できない不整合にもなる。
    ' 行は消さずに status を倒すのは、DiffDecision が failed/missing を
    ' replace 経路へ乗せる既存ルール(ResolveDecision)にそのまま乗るため。
    ' パスと元ファイル名の記録が残るので、何が復元されるかも追える。
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Worksheets(modAppDef.SH_MANIFEST)
    If Not wsM Is Nothing Then
        Dim lastM As Long: lastM = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).Row
        If lastM >= 2 Then
            Dim arr As Variant
            arr = wsM.Range(wsM.Cells(2, 5), wsM.Cells(lastM, 6)).Value
            Dim i As Long
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, 1) = 0                ' chunk_count(消した実態に合わせる)
                arr(i, 2) = "failed"         ' status
            Next i
            wsM.Range(wsM.Cells(2, 5), wsM.Cells(lastM, 6)).Value = arr

            ' 2026-08-01(同梱-8): fail_count(10列目)もここでリセットする。
            ' status は上で"failed"へ戻したのに fail_count(連続失敗回数)だけ
            ' ワイプ前の値を引き継ぐと、ワイプ後に社内NWへ再接続して最初の
            ' 取込がたまたま1回失敗しただけで MAX_FAIL_STREAK(3)に達し
            ' failed_permanent へ倒れる(恒久失敗バックオフの芽が育ってしまう)。
            ' ワイプは全資料を仕切り直す操作なので、連続失敗もゼロから数え直す。
            wsM.Range(wsM.Cells(2, 10), wsM.Cells(lastM, 10)).Value = 0
        End If
    End If

    ' 2026-07-31(R11-A): 消えたことを確かめてから「消しました」と言う。
    ' ここは全体が On Error Resume Next のため、シートの保護・行の削除拒否・
    ' 参照エラーで Delete が空振りしても、そのまま「消去しました」という
    ' 完了案内だけが出ていた(利用者は消えたと信じるが実際は残っている)。
    ' 残った行数を数えて呼び出し元へ返し、文言を切り替えられるようにする。
    Dim leftK As Long, leftV As Long
    If Not wsK Is Nothing Then
        leftK = wsK.Cells(wsK.Rows.Count, 1).End(xlUp).Row - 1
        If leftK < 0 Then leftK = 0
    End If
    If Not wsV Is Nothing Then
        leftV = wsV.Cells(wsV.Rows.Count, 1).End(xlUp).Row - 1
        If leftV < 0 Then leftV = 0
    End If
    outLeftRows = leftK + leftV
    If outLeftRows > 0 Then
        modLog.LogUsage "guard_wipe_incomplete", "", _
            "消去後も行が残っています(my_knowledge=" & leftK & _
            " / my_vectors=" & leftV & ")。シート保護・削除拒否の可能性。"
    End If
    On Error GoTo 0
End Function
