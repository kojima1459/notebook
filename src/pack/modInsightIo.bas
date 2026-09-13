Attribute VB_Name = "modInsightIo"
Option Explicit

' ============================================================================
' modInsightIo - 部内の知恵(解決済みQ&A/答えの無かった質問/訂正)を共有フォルダへ
'   出す(発信)/拾って insight_inbox シートへ入れる(収集)/古い残骸と既読印を
'   掃除する(GC)の3役と、その低水準ファイルI/O。
' ----------------------------------------------------------------------------
' 2026-07-31(R11-F1): modInsight が30,000字上限まで残り363字となり、修正が
'   1件も入らない状態だったため、憲章§4-6に基づき分割した。
'   切り口は「共有フォルダとのやり取り(本モジュール)」と「受信箱シートの
'   参照・選択・取り込み済み管理(modInsight)」。両者をまたぐ呼び出しは
'   CollectInsights → modInsight.EnsureSheet の1本だけで、共有する
'   モジュールレベル状態は無い(定数も用途ごとに完全に分かれていた)。
' ============================================================================

Private Const INSIGHT_SUBDIR As String = "insight"
Private Const QA_SUBDIR As String = "qa"
Private Const GAP_SUBDIR As String = "gap"
' 2026-08-14(R32 W1-1 B2): 訂正(correction)専用のサブフォルダ。従来は訂正も
' GAP_SUBDIR へ書いており、CollectFrom は【フォルダ単位で kind を決める】ため
' 訂正が kind="gap" として受信箱に入り、訂正本文(最長2,000字)が8列目
' =source_or_dept へ落ちて板の表示そのものを壊していた(理由も生英語で露出)。
' 出す場所を分けるのが根治。既に届いた旧データは modInsight.IsGapRow が救済。
Private Const CORR_SUBDIR As String = "correction"

' 2026-08-14(R32 W1-8(a) M7): 困りごと・訂正の発信に氏名を載せない
' 【ユーザー裁定=匿名化して続ける】。部署(DeptName)は残す ―― 資料を書ける人が
' 「どの部署で足りていないか」を掴む最小限で、個人は特定されない。解決済みQ&A
' (EmitVerifiedQA)は本人が✅解決したを押して共有すると分かっている経路なので
' 従来どおり氏名を載せる(取り込み本文にも出す設計。QABodyText)。
Private Const ANON_AUTHOR As String = "(匿名)"
Private Const MAX_COLLECT As Long = 60      ' 1回の起動で読むファイル数の上限
Private Const FIELD_SEP As String = vbTab

' 既読印(my_stats)の接頭辞。modP2P の "thx:" と同じ作法(R8 F10)。
Private Const INS_PREFIX As String = "ins:"

Private Const INBOX_COLS As Long = 10       ' A..J(J=選択状態)

' ----------------------------------------------------------------------------
' 発信: 解決済みQ&A(✅解決したの発火点から呼ばれる)
' ----------------------------------------------------------------------------
Public Sub EmitVerifiedQA(ByVal q As String, ByVal ans As String, ByVal src As String)
    On Error Resume Next
    ' 2026-08-01(R12-2-2): 質問全文・氏名・部署の共有フォルダ送信を止める
    ' プライバシーのエスケープハッチ(既定TRUE=共有知は維持)。docs/30 §9-1参照。
    If Not modConfig.GetBool("insight_share_enabled", True) Then Exit Sub
    If LenB(Trim$(q)) = 0 Or LenB(Trim$(ans)) = 0 Then Exit Sub

    ' R36 Fix A-M8: 是正メモが出典の先頭に来たターンで ✅解決した を押すと、
    ' 「もとになった資料」として "是正メモ_退職金は退職金は…_1234.txt" という
    ' 【他人の質問文がそのまま入ったファイル名】が部内へ配られる。資料名では
    ' なく個人の入力なので、共有本文には出さず固定文言へ置き換える。
    If modCorrect.IsMemoSource(src) Then src = "(利用者の是正メモ)"

    ' R34 F4: 部内へ出す本文にも出典突合の注記を通す。⚡/🔍 の注記は modApp の
    ' 画面用 ans にだけ付き、ここへ来る mLastCleanAnswer は未注記なので、
    ' 「画面では(出典確認できず)が付いていた回答が、知恵袋では無印で配られる」
    ' という非対称になっていた。部内一次情報になる本文こそ印が要る。
    ' 呼び出し元(modAsk の✅解決した)は凍結モジュールなので、共有本文の
    ' 境界であるここで掛ける。入念は生成の内側で注記済み=ShouldAnnotate=False
    ' で素通り(二重付与なし)。⚡/🔍 の mLastCleanAnswer は常に未注記=冪等。
    ans = modMode.AnnotateIfNeeded(ans)

    Dim dirPath As String: dirPath = SubDir(QA_SUBDIR)
    If LenB(dirPath) = 0 Then Exit Sub
    EnsureDir dirPath

    Dim myId As String: myId = SafeUserId()
    If LenB(myId) = 0 Then Exit Sub

    Dim body As String
    body = "v1" & FIELD_SEP & myId & FIELD_SEP & AuthorName() & FIELD_SEP & _
           Left$(modUtilText.IsoDateTime(Now), 16) & FIELD_SEP & _
           Clean1(q) & FIELD_SEP & Clean1(ans) & FIELD_SEP & Clean1(src)

    ' 2026-08-10(R25-3a FA-R25-3a): qa_shared_total は「解決済みQ&Aを部内へ
    ' 実際に送り出せた」ことの計測(=フライホイールへの実貢献)なので、
    ' 加算は共有フォルダへの書き込みが成功したときだけに限る(失敗経路で
    ' 加算しない。バッジ qa_share10 が獲得不可能だった不具合の修理)。
    If WriteShared(dirPath & MakeNonce(myId) & ".txt", body) Then
        modStats.Bump "qa_shared_total"
    End If
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 発信: 答えられなかった質問(根拠が薄い/違うと言われた)
'   reason: "no_hit"(本棚に無い) / "low_conf"(根拠が薄い) / "wrong"(回答が違う)
' ----------------------------------------------------------------------------
Public Sub EmitGap(ByVal q As String, ByVal reason As String)
    On Error Resume Next
    ' 2026-08-01(R12-2-2): 検索0件時は利用者操作を介さず自動発火するため、
    ' このゲートが最も効く経路(既定TRUE。docs/30 §9-1参照)。
    If Not modConfig.GetBool("insight_share_enabled", True) Then Exit Sub
    If LenB(Trim$(q)) = 0 Then Exit Sub
    ' R32 W1-8(b): 個人情報を含む可能性があるなら部内へ出さない。
    ' このゲートは波2の pii_scan_enabled(パック発行側のオフスイッチ)の
    ' 影響を受けず【常に走る】。理由は modInsightGate.PiiBlocked のコメント。
    ' R32 F4: 走査するのは【実際に送る文字列と同じ内容】。
    ' R32 マイクロ修正波 F17[MAJOR]: ただし Clean1 済みの文字列をそのまま
    ' 渡すと、改行が半角スペースへ均されて別々の数字列が繋がり誤検知する
    ' (詳細は modInsightGate.ScanClean のコメント参照)。走査だけは
    ' ScanClean(改行等を modPii の継続文字でない "," へ落とす)を通した q を
    ' 使い、送信本文(下の body)は現状どおり Clean1(q) のまま変えない。
    If modInsightGate.PiiBlocked(modInsightGate.ScanClean(q), "gap") Then Exit Sub
    ' R32 W1-6(M4): 同じ趣旨の質問の連投を止める(判定は modInsightGate 側)。
    If modInsightGate.GapDupBlocked(q) Then Exit Sub

    Dim dirPath As String: dirPath = SubDir(GAP_SUBDIR)
    If LenB(dirPath) = 0 Then Exit Sub
    EnsureDir dirPath

    ' R32 F2: 名乗るのはハッシュ化した匿名ID(ファイル名にもこれが載る)。
    Dim myId As String: myId = AnonUserId()
    If LenB(myId) = 0 Then Exit Sub

    ' R32 W1-8(a): 氏名は送らない(部署だけ)。困りごとの板は「誰が困ったか」
    ' ではなく「何の資料が足りないか」を見る場所で、氏名は要らないうえ、
    ' 部内の全員に「この人はこれを知らない」と配ることになる。
    Dim body As String
    body = "v1" & FIELD_SEP & myId & FIELD_SEP & ANON_AUTHOR & FIELD_SEP & _
           Left$(modUtilText.IsoDateTime(Now), 16) & FIELD_SEP & _
           Clean1(q) & FIELD_SEP & reason & FIELD_SEP & DeptName()

    ' 抑止キーを記録するのは【実際に部内へ出せたとき】だけ(共有フォルダへ
    ' 書けなかった投稿まで「もう送った」ことにすると、その質問は24時間
    ' 誰にも届かない)。qa_shared_total / gapfill_total と同じ考え方。
    If WriteShared(dirPath & MakeNonce(myId) & ".txt", body) Then
        modInsightGate.MarkGapEmitted q
    Else
        modInsightGate.NotifySkip "write"   ' R32 F5: 書けなかったことを黙らない
    End If
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 発信: 修正内容(誰かが「違う」を押して正しい内容を書いてくれた)
'   1人の訂正を全員の訂正にする。ただし個人の申告が即座に正になるのは危険なので、
'   受け取り側では「参考情報」として扱い、同じ趣旨の訂正が複数集まったものだけを
'   商品部が正式なナレッジに昇格させる運用を前提にしている(CorrectionAgreeCount)。
' ----------------------------------------------------------------------------
Public Sub EmitCorrection(ByVal answerText As String, ByVal fixText As String)
    On Error Resume Next
    ' 2026-08-01(R12-2-2): Emit系入口共通のゲート(既定TRUE。docs/30 §9-1参照)。
    If Not modConfig.GetBool("insight_share_enabled", True) Then Exit Sub
    If LenB(Trim$(fixText)) = 0 Then Exit Sub
    ' R32 W1-8: 訂正も同じ扱い(実名を出さない・PII検知で見送る)。訂正本文は
    ' 利用者が自由に書ける欄なので、質問文よりむしろ個人情報が入りやすい。
    ' R32 F4: 走査は【実際に送る2本の文字列と同じ内容】を【別々に】行う。
    ' 連結すると片方の末尾の数字ともう片方の先頭の数字が繋がって偽の長い
    ' 数字列になり(半角スペースは modPii がランの継続として数える)、
    ' しかも走査だけ全文・送信は先頭200字という食い違いもあった。
    ' R32 マイクロ修正波 F17[MAJOR]: 走査に Clean1 済みの文字列を渡すと、
    ' 改行が半角スペースへ均されて同じ理由で誤検知する(gap側F17と同じ
    ' 真因)。走査だけ modInsightGate.ScanClean を通した「素の文字列」を使い、
    ' 送信本文の qField/fField は Clean1 のまま変えない。
    Dim qRaw As String: qRaw = "【訂正】" & modUtil.SafeLeft(answerText, 200)
    Dim qField As String: qField = Clean1(qRaw)
    Dim fField As String: fField = Clean1(fixText)
    If modInsightGate.PiiBlocked(modInsightGate.ScanClean(qRaw), "correction") Then Exit Sub
    If modInsightGate.PiiBlocked(modInsightGate.ScanClean(fixText), "correction") Then Exit Sub

    ' R32 W1-1: 困りごとの板に混ざらないよう専用フォルダへ出す。
    Dim dirPath As String: dirPath = SubDir(CORR_SUBDIR)
    If LenB(dirPath) = 0 Then Exit Sub
    EnsureDir dirPath

    Dim myId As String: myId = AnonUserId()      ' R32 F2(訂正も匿名IDで出す)
    If LenB(myId) = 0 Then Exit Sub

    Dim body As String
    body = "v1" & FIELD_SEP & myId & FIELD_SEP & ANON_AUTHOR & FIELD_SEP & _
           Left$(modUtilText.IsoDateTime(Now), 16) & FIELD_SEP & _
           qField & FIELD_SEP & "correction" & FIELD_SEP & fField

    ' 2026-08-10(R25-3a FA-R25-3a): gapfill_total は「みんなの困りごと」への
    ' 回答(=1人の訂正を全員の訂正にする、この投稿)が実際に部内へ届いた
    ' ときだけ加算する(qa_shared_total と同じ考え方。EmitGap は困りごと
    ' そのものの報告であって回答ではないため対象外。correction_total
    ' (fb10/fb50)は本人のナレッジ登録成功で加算する別カウンタで、
    ' こちらは共有フォルダへの書き込み成功が条件)。バッジ gapfill が
    ' 獲得不可能だった不具合の修理。
    If WriteShared(dirPath & MakeNonce(myId) & ".txt", body) Then
        modStats.Bump "gapfill_total"
    Else
        modInsightGate.NotifySkip "write"   ' R32 F5: 書けなかったことを黙らない
    End If
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 受信: 共有フォルダの新着をinsight_inboxへ取り込む(APIは呼ばない=速い)。
'   戻り値 = 新しく受け取った件数。
' ----------------------------------------------------------------------------
Public Function CollectInsights() As Long
    On Error Resume Next
    Dim ws As Worksheet: Set ws = modInsight.EnsureSheet()
    If ws Is Nothing Then Exit Function

    ' 収集の前に受信箱を掃除する(R12-3-5)。ここが「行番号を握っている
    ' 利用者操作が1つも走っていない」と言える唯一の場所(起動時の収集)で、
    ' 取り込みループの最中に行を消すと outRow の指す先がずれる。
    modInsight.TrimInboxRows ws

    ' 2026-07-31(レビュー R8 F10): 既読印(my_stats の "ins:" 行)を
    ' 【1回の一括読み】で集合(Dictionary)にしてから照合する。
    ' 従来は1ファイルにつき modStats.GetStat("ins:"&nc) を呼んでいた。
    ' GetStat は my_stats を上から探す実装なので、共有フォルダのファイル数
    ' × my_stats の行数の計算量になり、しかも my_stats はこの機能自体が
    ' 際限なく行を足す(下の GcOldNonces まで掃除する仕組みが無かった)。
    ' 使うほど起動が重くなる形だった。
    Dim seen As Object: Set seen = LoadSeenSet()
    ' R32 W1-2(B1): 既読印とは別に、受信箱に在る nonce そのものを突き合わせる。
    ' 理由は modInsight.InboxNonceMemo の説明。
    Dim known As String: known = modInsight.InboxNonceMemo()

    Dim qaDir As String: qaDir = SubDir(QA_SUBDIR)
    Dim gapDir As String: gapDir = SubDir(GAP_SUBDIR)
    Dim corrDir As String: corrDir = SubDir(CORR_SUBDIR)
    CollectInsights = CollectFrom(ws, qaDir, "qa", seen, known) + _
                      CollectFrom(ws, gapDir, "gap", seen, known) + _
                      CollectFrom(ws, corrDir, "correction", seen, known)

    ' GC(R8 F10)。my_stats 側の既読印は【全端末】が自分のブックを掃除する。
    ' 自分のシートが太るのは自分の問題なので、誰がやっても構わない。
    GcOldNonces
    ' R13-7e: sv:d:(節約時間の日別キー)のGCも同じタイミングに相乗りする。
    ' ここ(CollectInsights)はEnsureSheetがローカル処理のみで共有フォルダの
    ' 設定・到達性に関係なく毎起動走る(modBoot.Boot内で無条件に呼ばれる)。
    ' sv:d:はP2P/共有を使わない端末でも「✅解決」のたびに増えるローカル値
    ' なので、共有I/Oを前提にした呼び出し口(modP2P.CollectThanksは
    ' 共有パス未設定だと最初のExit Functionで素通りしてしまう)には乗せられない。
    GcOldSavedDays
    ' R32 W1-6: 連投抑止キー(gapq:)のGCも同じタイミングで。実体は
    ' modInsightGate 側(キー名・書き込み・掃除を1モジュールに集める)。
    modInsightGate.GcGapDupKeys

    ' 2026-07-31(R8b B12): 共有フォルダ側(qa/gap の実ファイル)を消すのは
    ' 【発行者端末だけ】に限定する。
    ' 全端末が消しに行くと、次の2つが同時に起きる:
    '   ・朝の一斉起動で数十台が同じフォルダに対して Kill を撃ち合う
    '     (消えた直後のファイルへの Kill でエラー、列挙と削除の競合)
    '   ・「まだ誰も受け取っていない投稿」を、たまたま最初に起動した1台の
    '     時計や設定(thanks_gc_days)の都合で全員から奪える
    ' 共有フォルダの寿命管理は、本来1箇所が責任を持つべき運用作業に近い。
    ' 判定は既存の発行者判定(modPublish.CanPublish = config publish_key が
    ' 入っている = --publisher 付きでビルドした発行者用ブック)を再利用する。
    ' 新しい設定キーを増やさない(増やすほど運用が壊れやすくなる)。
    Dim isPublisher As Boolean
    isPublisher = False
    On Error Resume Next
    isPublisher = modPublish.CanPublish()
    On Error GoTo 0
    If Not isPublisher Then Exit Function

    Dim killedN As Long
    killedN = GcOldInsights(qaDir) + GcOldInsights(gapDir) + GcOldInsights(corrDir)
    If killedN > 0 Then
        On Error Resume Next
        modLog.LogUsage "insight_gc", "", _
            "共有フォルダの古い共有知を " & killedN & "件片付けました" & _
            "(発行者端末のみが実行。保持日数は config thanks_gc_days)。"
        On Error GoTo 0
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' LoadSeenSet - my_stats の "ins:" 行を1回で読んで既読集合にする(R8 F10)。
'   キーは nonce(ファイル名から拡張子を除いたもの)。値は使わない。
' ----------------------------------------------------------------------------
Private Function LoadSeenSet() As Object
    Dim d As Object
    On Error Resume Next
    Set d = CreateObject("Scripting.Dictionary")
    Set LoadSeenSet = d
    If d Is Nothing Then Exit Function

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_STATS)
    If ws Is Nothing Then Exit Function
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Function

    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastR, 1)).Value

    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim k As String: k = CStr(arr(i, 1))
        If Left$(k, Len(INS_PREFIX)) = INS_PREFIX Then
            d(LCase$(Mid$(k, Len(INS_PREFIX) + 1))) = 1
        End If
    Next i
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' GcOldInsights - N日以上前の qa/gap ファイルを共有フォルダから消す(R8 F10)。
'   戻り値 = 実際に消した件数(R8b B12: 呼び出し側が usage_log に残す)。
'   modP2P.GcOldThanks と同じ作法(集めてから消す。列挙中に Kill しない)。
'   保持日数は感謝状と揃えて config thanks_gc_days(既定60)。0以下でGC無効。
'   ここに掃除が無かったため、共有の insight\qa\ は永久に増え続け、
'   起動のたびの Dir 列挙が全員ぶん重くなっていた。
'   【呼ぶのは発行者端末だけ】(R8b B12。理由は CollectInsights 側のコメント)。
' ----------------------------------------------------------------------------
Private Function GcOldInsights(ByVal dirPath As String) As Long
    If LenB(dirPath) = 0 Then Exit Function
    On Error Resume Next
    Dim days As Long: days = modConfig.GetLong("thanks_gc_days", 60)
    If days <= 0 Then Exit Function
    If Len(Dir(dirPath, vbDirectory)) = 0 Then Exit Function

    Dim old() As String: ReDim old(0 To 63)
    Dim n As Long
    Dim limit As Date: limit = DateAdd("d", -days, Now)

    Dim fn As String: fn = Dir(dirPath & "*.txt")
    Do While LenB(fn) > 0
        If n > UBound(old) Then ReDim Preserve old(0 To UBound(old) + 64)
        old(n) = fn
        n = n + 1
        fn = Dir()
    Loop

    Dim i As Long
    For i = 0 To n - 1
        Dim full As String: full = dirPath & old(i)
        Dim stamp As Date
        Err.Clear
        stamp = FileDateTime(full)
        If Err.Number = 0 Then
            If stamp < limit Then
                If modP2PIo.KillRetry(full) Then GcOldInsights = GcOldInsights + 1
            End If
        End If
        Err.Clear
    Next i
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' GcOldNonces - 期限を過ぎた "ins:" 行を my_stats から取り除く(R8 F10)。
'   modP2P.GcOldNonces と同じ作法。共有側のファイルを N日で消す以上、
'   既読印だけ永久に残すと my_stats が無限に伸びる(=起動が重くなる)。
'   保持は共有側+7日(modInsightGate.NonceKeepDays)。受信箱側(InboxKeepDays)は
'   必ずこれより長い ―― 両方が同時に消える日を作らないため(R32 F1)。
'
'   旧い端末の既読印は Bump で書かれた数値("1")で、日付として読めない。
'   modP2P は読めない値を掃除対象にしているが、ここで同じことをすると
'   【まだ共有に残っている投稿を二度取り込む】ことになり、受信箱に同じ
'   Q&Aが並ぶ。読めない値は今日の日付へ書き直すだけにして、次の周回から
'   正しく期限管理する(移行は1回で終わる)。
' ----------------------------------------------------------------------------
Private Sub GcOldNonces()
    On Error Resume Next
    Dim keepDays As Long
    keepDays = modInsightGate.NonceKeepDays(modConfig.GetLong("thanks_gc_days", 60))
    If keepDays < 1 Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_STATS)
    If ws Is Nothing Then Exit Sub
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Sub

    Dim limit As Date: limit = DateAdd("d", -keepDays, Date)
    Dim today As String: today = modUtilText.IsoDate(Date)
    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastR, 2)).Value

    ' 下から消す(上から消すと行番号がずれる)。
    Dim i As Long
    For i = UBound(arr, 1) To LBound(arr, 1) Step -1
        Dim k As String: k = CStr(arr(i, 1))
        If Left$(k, Len(INS_PREFIX)) = INS_PREFIX Then
            Dim v As String: v = Trim$(CStr(arr(i, 2)))
            ' 旧形式は Bump が書いた数値("1")。CDate("1") は 1899-12-31 という
            ' 「有効な日付」になってしまうので、数値かどうかを先に見る
            ' (ここを間違えると、旧端末の既読印が全部期限切れ扱いで消える)。
            Dim isOldForm As Boolean
            isOldForm = (LenB(v) = 0)
            If Not isOldForm Then isOldForm = IsNumeric(v)
            If isOldForm Then
                ws.Cells(i + 1, 2).Value = today     ' 旧形式("1")の移行
            Else
                Dim d As Date
                Err.Clear
                d = CDate(v)
                If Err.Number = 0 Then
                    If d < limit Then ws.Rows(i + 1).Delete
                End If
                Err.Clear
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' GcOldSavedDays(2026-08-03 R13-7e): sv:d:yyyymmdd(節約時間の【日別】キー)を
'   400日超で掃除する。sv:m:/sv:y:(月次/年次)は履歴として残す価値があるが、
'   日別は modBoard が「今日」と「直近7日」までしか読まず、無期限に伸ばす
'   理由が無い(HANDOFF §2「my_statsの行数が育つほどFindKeyRowの線形探索が
'   重くなる」と同じ課題)。GcOldNonces(このモジュール上・modP2P上とも)と
'   同じ「集めてから下から消す」作法。日付はキー側(A列)に埋め込まれている
'   ため、値側(B列)を見るGcOldNoncesとは判定材料が異なる。
' ----------------------------------------------------------------------------
Private Sub GcOldSavedDays()
    On Error Resume Next
    Const KEEP_DAYS As Long = 400
    Const SV_D_PREFIX As String = "sv:d:"

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_STATS)
    If ws Is Nothing Then Exit Sub
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Sub

    Dim limit As Date: limit = DateAdd("d", -KEEP_DAYS, Date)
    Dim arr As Variant
    ' A:B の2列で読む(1行しか無いときA列単独だと配列ではなくスカラーが
    ' 返るVBAの仕様を避けるため。GcOldNoncesと同じ回避策)。
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastR, 2)).Value

    Dim i As Long
    For i = UBound(arr, 1) To LBound(arr, 1) Step -1
        Dim k As String: k = CStr(arr(i, 1))
        If Left$(k, Len(SV_D_PREFIX)) = SV_D_PREFIX Then
            Dim svDatePart As String: svDatePart = Mid$(k, Len(SV_D_PREFIX) + 1)
            Dim drop As Boolean: drop = True   ' 読めない形式も掃除対象
            If Len(svDatePart) = 8 And IsNumeric(svDatePart) Then
                Err.Clear
                Dim d As Date
                d = DateSerial(CInt(Left$(svDatePart, 4)), CInt(Mid$(svDatePart, 5, 2)), CInt(Mid$(svDatePart, 7, 2)))
                If Err.Number = 0 Then drop = (d < limit)
                Err.Clear
            End If
            If drop Then ws.Rows(i + 1).Delete
        End If
    Next i
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 内部: 共有フォルダ→シート
' ----------------------------------------------------------------------------
Private Function CollectFrom(ByVal ws As Worksheet, ByVal dirPath As String, _
                             ByVal kind As String, ByVal seen As Object, _
                             ByRef known As String) As Long
    If LenB(dirPath) = 0 Then Exit Function
    On Error Resume Next
    If Len(Dir(dirPath, vbDirectory)) = 0 Then Exit Function

    Dim myId As String: myId = SafeUserId()
    ' R32 F2: 自分発の判定は【生ID】と【匿名ID】の両方で行う。新しい困りごと・
    ' 訂正はハッシュ名で出るが、解決済みQ&Aは従来どおり生ID名で出るうえ、
    ' 共有フォルダには移行前に出した旧ファイル(生ID名)も残っているため。
    Dim myHash As String: myHash = AnonUserId()

    ' 1) まずファイル名を全部集める。Dir()は列挙状態を1つしか持たないので、
    '    ループの中で読み書きをすると列挙が壊れる(他モジュールと同じ2段方式)。
    Dim names() As String: ReDim names(0 To 255)
    Dim nFiles As Long
    Dim fname As String
    fname = Dir(dirPath & "*.txt")
    Do While LenB(fname) > 0
        If nFiles > UBound(names) Then ReDim Preserve names(0 To UBound(names) + 256)
        names(nFiles) = fname
        nFiles = nFiles + 1
        fname = Dir()
    Loop

    ' 2) 未読だけを処理する。
    '
    ' 2026-07-28(レビュー H-15)で直した3点:
    '   (a) 読み取りに失敗しても既読にしていた
    '       → その投稿はその端末に二度と入らない。既読化は成功パスへ移す。
    '   (b) 上限60を「列挙した件数」で数えていた
    '       → 既読・自分発のファイルも枠を食うため、部内の累計投稿が60件を
    '         超えた日から新着が誰にも届かなくなる(エラー表示も無し)。
    '         共有知フライホイールが静かに止まる。数えるのは「実際に
    '         処理した未読の件数」にする。
    '   (c) 自分発の判定が前方一致だった(レビュー M-22)
    '       → ID "鈴木" が "鈴木一" さんの投稿まで自分発と誤認して受信しない。
    '         ファイル名は "<ID>-<nonce>" なので、区切りまで含めて比較する。
    Dim processed As Long
    Dim i As Long
    For i = 0 To nFiles - 1
        If processed >= MAX_COLLECT Then Exit For
        Dim nc As String: nc = Left$(names(i), Len(names(i)) - 4)
        If Not (IsMine(nc, myId) Or IsMine(nc, myHash)) Then
            ' 既読判定は一括読みした集合で行う(R8 F10)。my_stats を
            ' ファイル1件ごとに走査しない。
            If Not IsSeen(seen, nc) Then
                processed = processed + 1
                Dim raw As String
                If ReadShared(dirPath & names(i), raw) Then
                    ' 2026-08-14(R32 W1-9 M8): 既読にしてよいのは「受信箱に
                    ' その行が確実に在る」ときだけ。従来は ReadShared 成功だけで
                    ' 印を書いており、書き込み途中(=フィールド不足)のファイルを
                    ' 読んだ瞬間にその投稿はこの端末から永久消失していた。
                    Dim added As Boolean
                    If AppendRow(ws, kind, nc, raw, known, added) Then
                        If added Then CollectFrom = CollectFrom + 1
                        ' 値は日付にしておく(GcOldNonces が期限を判定できる
                        ' ようにするため。R8 F10)。
                        modStats.SetStatText INS_PREFIX & nc, modUtilText.IsoDate(Date)
                        If Not seen Is Nothing Then seen(LCase$(nc)) = 1
                    End If
                End If
            End If
        End If
    Next i
    On Error GoTo 0
End Function

' 既読集合の照合(R8 F10)。集合が作れなかった環境(Scripting.Dictionary が
' 使えない等)では、従来どおり my_stats を1件ずつ引く経路へ落とす。
' 「速くするための仕組みが無いと動かない」は、機能そのものを壊す作りなので避ける。
Private Function IsSeen(ByVal seen As Object, ByVal nc As String) As Boolean
    If seen Is Nothing Then
        IsSeen = (LenB(modStats.GetStatText(INS_PREFIX & nc)) > 0)
        Exit Function
    End If
    IsSeen = seen.Exists(LCase$(nc))
End Function

' ファイル名の先頭が自分のIDか。前方一致だと "鈴木" が "鈴木一" を巻き込むため、
' 「ID そのもの」か「ID + 区切り」でだけ一致とみなす(レビュー M-22)。
Private Function IsMine(ByVal nameCore As String, ByVal myId As String) As Boolean
    If LenB(myId) = 0 Then Exit Function
    If StrComp(nameCore, myId, vbTextCompare) = 0 Then
        IsMine = True
        Exit Function
    End If
    IsMine = (StrComp(Left$(nameCore, Len(myId) + 1), myId & "-", vbTextCompare) = 0)
End Function

' insight_inbox の列: nonce / kind / user_id / author / created_at /
'                     question / answer / source / consumed
' 戻り値 = 「この nonce の行が受信箱に確実に在る」(=既読印を書いてよい)。
'          新規に足したときだけ outAdded=True(件数はこちらで数える)。
' 2026-08-14(R32 W1-2/W1-9): 「足せたか」と「既読にしてよいか」は別の問い
' なので分けた。壊れたファイル(フィールド不足)は両方False=次回また読む。
Private Function AppendRow(ByVal ws As Worksheet, ByVal kind As String, _
                           ByVal nc As String, ByVal raw As String, _
                           ByRef known As String, ByRef outAdded As Boolean) As Boolean
    outAdded = False
    Dim f() As String
    f = Split(raw, FIELD_SEP)
    If UBound(f) < 5 Then Exit Function     ' 書き込み途中・壊れたファイル

    ' R32 W1-2(B1): 同じ nonce の行が既に在るなら足さない(冪等化)。
    If modInsight.NonceIsKnownIn(known, nc) Then
        AppendRow = True                    ' 受信箱には在る=既読にしてよい
        Exit Function
    End If

    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = nc
    ws.Cells(r, 2).Value = kind
    ' 2026-08-01(R12-2-1): f(1)〜f(6) は共有フォルダの他人のファイルから
    ' 読んだ未信頼テキスト。セルへ書く前に数式インジェクション対策を通す
    ' (セキュリティ監査3)。
    ws.Cells(r, 3).Value = modUtilText.SanitizeForCell(f(1))
    ws.Cells(r, 4).Value = modUtilText.SanitizeForCell(f(2))
    ws.Cells(r, 5).Value = modUtilText.SanitizeForCell(f(3))
    ws.Cells(r, 6).Value = modUtilText.SanitizeForCell(f(4))
    If kind = "qa" Then
        ws.Cells(r, 7).Value = modUtilText.SanitizeForCell(f(5))
        If UBound(f) >= 6 Then ws.Cells(r, 8).Value = modUtilText.SanitizeForCell(f(6))
    Else
        ' kind="gap":        f(5)=reason / f(6)=部署
        ' kind="correction": f(5)="correction" / f(6)=訂正本文(R32 W1-1で
        '   フォルダを分けたので、この行が困りごとの板へ出ることはもう無い。
        '   電文の並びは旧版と互換のまま置く=旧gapフォルダに残っている訂正も
        '   同じ列に落ち、reason列ガード(modInsight.IsGapRow)1本で弾ける)。
        ws.Cells(r, 7).Value = modUtilText.SanitizeForCell(f(5))
        If UBound(f) >= 6 Then ws.Cells(r, 8).Value = modUtilText.SanitizeForCell(f(6))
    End If
    ws.Cells(r, 9).Value = ""
    known = modInsight.NonceMemoAdd(known, nc)
    outAdded = True
    AppendRow = True
End Function

' ----------------------------------------------------------------------------
' 内部: 共有フォルダI/O(modBoard.WriteBeacon と同仕様の意図的な複製。
'   モジュール完全隔離のため、他モジュールのPrivateには依存しない)
' ----------------------------------------------------------------------------
Private Function WriteShared(ByVal filePath As String, ByVal content As String) As Boolean
    Dim attempt As Long
    For attempt = 1 To 3
        If TryWrite(filePath, content) Then
            WriteShared = True
            Exit Function
        End If
        ShortWait 200 * attempt
    Next attempt
End Function

' UTF-8書き出しの実体は modUtilText.WriteTextFileUtf8(2026-07-31 R11-F2)。
Private Function TryWrite(ByVal filePath As String, ByVal content As String) As Boolean
    TryWrite = modUtilText.WriteTextFileUtf8(filePath, content)
End Function

' 共有フォルダからの読み取り。AVスキャン中・書込み直後のロックで一時的に
' 失敗することがあるため、modP2P と同じく3回まで待って試す(レビュー H-15)。
' 1回で諦めると、その投稿は既読化されないまま毎回失敗し続けるか、
' 修正前のように既読化されて永久に落ちる。
Private Function ReadShared(ByVal filePath As String, ByRef outText As String) As Boolean
    Dim attempt As Long
    For attempt = 1 To 3
        If ReadSharedOnce(filePath, outText) Then
            ReadShared = True
            Exit Function
        End If
        WaitMs 150 * attempt
    Next attempt
End Function

' 指定ミリ秒だけ待つ(Sleep宣言を増やさないためTimerで回す。共有I/Oの
' リトライ間隔なので精度は要らない)。
Private Sub WaitMs(ByVal ms As Long)
    On Error Resume Next
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000 < ms
        DoEvents
        If Timer < t0 Then Exit Do      ' 日付をまたいだ(Timerが0へ戻る)
    Loop
    On Error GoTo 0
End Sub

' UTF-8読み取りの実体は modUtilText.ReadTextFileUtf8(2026-07-31 R11-F2)。
Private Function ReadSharedOnce(ByVal filePath As String, ByRef outText As String) As Boolean
    ReadSharedOnce = modUtilText.ReadTextFileUtf8(filePath, outText)
End Function

Private Sub ShortWait(ByVal ms As Long)
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000# < ms
        DoEvents
        If Timer < t0 Then Exit Do
    Loop
End Sub

' ----------------------------------------------------------------------------
' 内部: 小物
' ----------------------------------------------------------------------------
' 到達性の判定は modShare が1セッション1回だけ行う(レビュー M-20)。
Private Function SubDir(ByVal leaf As String) As String
    Dim root As String: root = modShare.SubDir(INSIGHT_SUBDIR)
    If LenB(root) = 0 Then Exit Function
    EnsureDir root
    SubDir = root & leaf & "\"
End Function

Private Sub EnsureDir(ByVal folderPath As String)
    On Error Resume Next
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    On Error GoTo 0
End Sub

Private Function MakeNonce(ByVal myId As String) As String
    MakeNonce = myId & "-" & Format$(Now, "yyyymmddhhnnss") & "-" & _
            Format$(Int(Timer * 1000) Mod 100000, "00000")
End Function

Private Function SafeUserId() As String
    On Error Resume Next
    SafeUserId = modP2P.CurrentUserId()
    On Error GoTo 0
End Function

' R32 F2: 困りごと・訂正を出すときに名乗るID。実体と旧データ互換の考え方は
' modInsightGate.AnonId(ハッシュ化する理由・IsMineが成立する理由をそこに集約)。
Private Function AnonUserId() As String
    AnonUserId = modInsightGate.AnonId(SafeUserId())
End Function

Private Function AuthorName() As String
    On Error Resume Next
    AuthorName = Trim$(modConfig.GetString("pack_author", ""))
    On Error GoTo 0
    If LenB(AuthorName) = 0 Or AuthorName = "名称未設定" Then AuthorName = "どなたか"
End Function

Private Function DeptName() As String
    On Error Resume Next
    DeptName = Trim$(modConfig.GetString("user_department", ""))
    On Error GoTo 0
End Function

' タブ・改行はフィールド区切りと衝突するため潰す。
Private Function Clean1(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace(t, vbTab, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    Clean1 = modUtil.SafeLeft(Trim$(t), 2000)
End Function
