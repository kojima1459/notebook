Attribute VB_Name = "modGuard"
Option Explicit

' ドメイン不一致の連続回数を控えるmy_statsのキー(レビュー M-6)。
Private Const STREAK_KEY As String = "domain_block_streak"
' 失効の予告を出した日数(同じ日数で繰り返し出さないため)と、
' 実際に消した記録(消したあと毎回通知しないため)。レビュー L-7。
Private Const WARNED_KEY As String = "expiry_warned_days"
Private Const WIPED_KEY As String = "expiry_wiped_at"

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
'      (既定30日)が過ぎたら、my_knowledge / my_vectors の中身を消す。
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

    ' ドメイン名が取れない環境(ワークグループ等)は許可しない。
    ' allowed_domain を設定した組織は「ドメイン参加端末でのみ使う」意思表示。
    If LenB(domNb) = 0 And LenB(domDns) = 0 Then
        CheckDomain = False
        Exit Function
    End If

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
    modStats.SetStatText K_LAST_REACH, Format$(Date, "yyyy-mm-dd")
    ' つながったので失効まわりの記録はリセットする(レビュー L-7)。
    ' これをしないと、復帰後も「予告済み」「消去済み」の印が残り、
    ' 次に離れたときの予告が出なくなる。
    If LenB(modStats.GetStatText(WARNED_KEY)) > 0 Then modStats.SetStatText WARNED_KEY, ""
    If LenB(modStats.GetStatText(WIPED_KEY)) > 0 Then modStats.SetStatText WIPED_KEY, ""
    On Error GoTo 0
End Sub

' 最終到達日からの経過日数(記録が無ければ0=今日として扱い、誤爆を防ぐ)。
Public Function DaysSinceReach() As Long
    On Error Resume Next
    Dim s As String: s = modStats.GetStatText(K_LAST_REACH)
    If LenB(s) = 0 Then
        ' 初回起動や共有フォルダ未設定。記録を作って0日扱いにする。
        TouchReach
        Exit Function
    End If
    DaysSinceReach = CLng(Date - CDate(s))
    If DaysSinceReach < 0 Then DaysSinceReach = 0
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' EnforceExpiry - 期限切れなら知識を消す。7日前から警告を出す。
'   戻り値 True = 消去した。
' ----------------------------------------------------------------------------
Public Function EnforceExpiry() As Boolean
    On Error Resume Next
    Dim limitDays As Long
    limitDays = modConfig.GetLong("knowledge_expire_days", 30)
    If limitDays <= 0 Then Exit Function      ' 0以下で機能OFF

    ' 共有フォルダを使っていない環境ではこの機能自体を動かさない
    ' (単独利用の人の知識を、つながる先が無いという理由で消すのは理不尽)。
    If LenB(Trim$(modConfig.GetString("nexus_share_path", ""))) = 0 Then Exit Function

    Dim d As Long: d = DaysSinceReach()
    If d < limitDays - 7 Then Exit Function

    ' 2026-07-28(レビュー L-7): 3つの雑さを直す。
    '  (a) expire_days<=7 だと limitDays-7<=0 になり、初日から毎回警告が出る
    '      → 「1日でも猶予がある」ときだけ予告する。
    '  (b) 期限超過後は開くたびにワイプ通知が出ていた
    '      → 一度消したら記録し、二度目からは黙って通す(もう消すものが無い)。
    '  (c) 端末の時計が前へ飛ぶと、予告を一度も出さずにいきなり消えていた
    '      → 予告を出した記録が無ければ、まず予告だけ出して1回見送る。
    If d >= 1 And d < limitDays Then
        If StrComp(modStats.GetStatText(WARNED_KEY), CStr(d), vbTextCompare) = 0 Then Exit Function
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

    ' 期限超過。
    ' 予告を一度も出せていない(時計の前進ジャンプ等)なら、まず予告だけ出す。
    If LenB(modStats.GetStatText(WARNED_KEY)) = 0 Then
        modStats.SetStatText WARNED_KEY, CStr(d)
        MsgBox "社内ネットワークに " & d & " 日間つながっていません。" & vbCrLf & _
               "次に開いたときも接続できていない場合、安全のため" & vbCrLf & _
               "取り込んだ知識を消去します。" & vbCrLf & vbCrLf & _
               "社内ネットワークに接続して一度開いていただければ解除されます。", _
               vbExclamation, modAppDef.APP_NAME
        Exit Function
    End If

    ' 既に消したあとなら、開くたびに同じ通知を出さない。
    If LenB(modStats.GetStatText(WIPED_KEY)) > 0 Then Exit Function

    ' 知識だけを消す。
    Dim removed As Long
    removed = WipeKnowledge()
    modStats.SetStatText WIPED_KEY, modUtil.NowStamp()
    modLog.LogUsage "guard_wipe", "", "days=" & d & " removed=" & removed
    MsgBox "社内ネットワークに " & d & " 日間つながらなかったため、" & vbCrLf & _
           "安全のため取り込んだ知識を消去しました。" & vbCrLf & vbCrLf & _
           "社内ネットワークに接続して開き直すと、部門チャンネルから" & vbCrLf & _
           "自動的に取り込み直せます。設定・統計・履歴は残っています。", _
           vbInformation, modAppDef.APP_NAME
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
Public Function WipeKnowledge() As Long
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
        End If
    End If
    On Error GoTo 0
End Function
