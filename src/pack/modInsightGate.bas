Attribute VB_Name = "modInsightGate"
Option Explicit

' ============================================================================
' modInsightGate - 「部内へ出してよいか」を決める発信側の関所(2026-08-14 R32)
' ----------------------------------------------------------------------------
' なぜ独立したモジュールなのか(容量と主題の両方):
'   R32波1では、この関所の実体を modInsight へ間借りさせていた
'   (発信の持ち主 modInsightIo が30,000字上限まで残り1,800字を切っていたため。
'    憲章§4-6「入らなければ実体を余裕モジュールへ置いて1行呼び出し」)。
'   Fix波でF2(匿名ID)・F4(日付誤検知の前処理)・F5(見送りの通知)を足した
'   ところで modInsight も30,000字を超えたため、間借りをやめて
'   「出してよいかの判断」だけを持つ置き場所を独立させた。
'   modInsight は受信箱(insight_inbox)の参照・集計・表示、modInsightIo は
'   共有フォルダとの実I/O、本モジュールは【出す前の判断】と役割が割れる。
'
' 持ち分:
'   AnonId          … 発信者IDの匿名化(ハッシュ化)。F2
'   PiiBlocked      … 個人情報らしき文字列を含むなら出さない。W1-8(b)/F4
'   StripDateLike   … PII走査へ渡す前に日付・時刻の並びを潰す前処理。F4/F16
'   ScanClean       … PII走査へ渡す前に改行等をカンマへ潰す前処理。F17
'   GapDupBlocked / MarkGapEmitted / GcGapDupKeys
'                   … 同じ趣旨の質問の連投抑止(既定24時間)とそのGC。W1-6
'   NotifySkip      … 見送ったことを利用者へ伝えるトースト。F5
'   NonceKeepDays / InboxKeepDays
'                   … 既読印GCと受信箱保持の不等式(定数の唯一の持ち主)。F1
'
' 時刻の窓の判定(WithinWindow)だけは modInsight に残してある ―― 日付文字列の
' 一般的な比較で、受信箱側(GapAged)と同じ「ISO文字列の辞書順で比べる」作法の
' 一員だからで、modTestsPure30 が既にそこを固定している。
' ============================================================================

' 連投抑止キー(my_stats)の接頭辞。modP2P の "thx:" / 本機能の "ins:" と同作法。
' 2026-08-14(R32 W1-6)。
Private Const GAPQ_PREFIX As String = "gapq:"
' 抑止キーそのものの保持日数(抑止期間を過ぎたキーを my_stats に残す理由は無い)。
Private Const GAPQ_KEEP_DAYS As Long = 7

' ----------------------------------------------------------------------------
' AnonId - 発信者IDを匿名化する(純関数・2026-08-14 R32 F2)。
' ----------------------------------------------------------------------------
'   W1-8で電文の【著者名】だけを "(匿名)" にしたが、匿名化としては穴だらけだった:
'     (a) 共有フォルダに置くファイル名が MakeNonce(myId) =
'         "山田 太郎-20260814093012-01234.txt"。フォルダを開けば実名が並ぶ。
'     (b) 電文の2番目のフィールド(user_id)が modP2P.CurrentUserId() で、
'         これはActive DirectoryのCN(=実名)。受信箱のC列にそのまま入る。
'   氏名の欄だけ伏せても、ファイル名とID欄から誰の投稿かは一目で分かる。
'   そこで【発信時のIDそのもの】を FNV-1a 64bit のハッシュ(16桁の16進)へ
'   置き換える。modMentor が質問ファイルの宛先に使っているのと同じ作法で、
'   逆算はできず、同じ人からは常に同じ値になる(=自分の投稿の判別は保てる)。
'
'   IsMine がハッシュ同士の比較で成立する理由: modInsightIo.CollectFrom は
'   ファイル名の先頭が「自分のID」かで自分発を弾く。発信側がハッシュを使えば
'   ファイル名は "<ハッシュ>-<時刻>-<連番>.txt" になり、受信側も同じ関数で
'   作ったハッシュと突き合わせるので、同じ端末では必ず一致する。
'
'   【旧データ互換】ハッシュ化するのは【新しく出す投稿だけ】。
'     ・既に共有フォルダに在る旧ファイルは名前が生ID。受信側は生IDでも
'       ハッシュでも自分発を弾けるよう、両方で照合する(CollectFrom)。
'     ・既読印("ins:"+nonce)は nonce = ファイル名そのままなので、旧ファイルの
'       既読印は旧nonceのまま有効。読み取り側は一切変更が要らない。
'     ・解決済みQ&A(EmitVerifiedQA)は「本人が✅解決したを押して共有すると
'       分かっている」経路で氏名も本文に載せる設計なので、従来どおり生IDのまま
'       (ここをハッシュにすると誰の知恵か分からなくなり、機能の意味が消える)。
'   空文字を渡されたら空文字を返す ―― 呼び出し側が「IDが取れなかった」として
'   発信を止められるようにするため(ここで生IDへ落とすと匿名化が漏れる)。
' ----------------------------------------------------------------------------
Public Function AnonId(ByVal rawId As String) As String
    If LenB(rawId) = 0 Then Exit Function
    On Error Resume Next
    AnonId = modUtil.Fnv1a64Hex(rawId)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' PiiBlocked - 共有フォルダへ出す前の個人情報走査(R32 W1-8(b) M7)。
'   True を返したら【送らない】。
'   困りごとの投稿は質問の全文がそのまま部内へ出ていく経路で、しかも
'   no_hit は利用者の操作を介さず自動で発火する(押した覚えのないまま
'   「取引先の田中様の携帯 090-xxxx-xxxx は…」が部内に配られる)。
'   共有フォルダに一度書いたものは取り消せないため、ここは検知したら黙って
'   見送る(利用者の作業は止めない。痕跡は usage_log に1行残す)。
'
'   【重要】この走査は、パック発行側の個人情報チェックのオフスイッチ
'   (config pii_scan_enabled・R32 波2)の影響を受けない。あちらは自分で
'   選んで自分の資料を出す操作で、利用者が中身を分かって押している。
'   こちらは自動発火かつ他人の目に触れる経路なので、常に走らせる。
' ----------------------------------------------------------------------------
Public Function PiiBlocked(ByVal s As String, ByVal what As String) As Boolean
    On Error Resume Next
    Dim hit As String
    ' R33 W2-7: 前処理の【いちばん先頭】で幅を均す。StripDateLike は半角数字
    ' 専用なので、幅を均すのを ScanText の内側だけに置くと、全角で書かれた
    ' 日付("２０２６－０８－１４ １０:００")が日付潰しを素通りしたまま
    ' ScanText の内側で半角化され、12桁ランとして検知されてしまう
    ' ―― R32 F4 が半角側で潰した誤検知の全角版で、質問が無言で共有見送りに
    ' なる。均しは冪等(出力に全角は残らない)なので、ScanText 内側の均しと
    ' 二重に通っても結果は変わらない。
    ' R32 F4(b): 走査の前に日付・時刻の並びを潰す(理由は StripDateLike)。
    hit = modPii.ScanText(StripDateLike(modPii.NormalizeWidth(s)))
    If LenB(hit) = 0 Then Exit Function
    PiiBlocked = True
    modLog.LogUsage "insight_pii_skip", what, _
        "個人情報を含む可能性があるため部内共有を見送りました(" & hit & ")"
    ' R32 F5: 黙って落とさない。利用者へ「今回は共有していない」と理由を伝える。
    NotifySkip "pii"
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' StripDateLike - 日付・時刻らしき並びを潰す前処理(純関数・2026-08-14 R32 F4)。
' ----------------------------------------------------------------------------
'   【必ず誤検知していたこと】modPii.HasLongDigitRun は "-"・半角スペース・
'   U+2010 を「数字のランの継続」として数える(電話番号 090-1234-5678 を
'   1本の12桁として拾うための仕様)。ところが ISO のタイムスタンプ
'   "2026-08-14 10:00" も同じ規則で 4+2+2+2+2 = 12桁と数えられ、【10桁以上】に
'   一発で当たる。共有発信側のPII走査は常時ONなので、質問文に日時がひとつ
'   入っているだけで、その質問は永久に部内へ出ない(しかも利用者には見えない)。
'
'   【なぜ modPii 本体を直さないのか】modPii は パック書き出し(📦/📤)の
'   PII検知と共用で、そちらは R32 波2 の対象(既定オフへ倒したばかり)。
'   桁数閾値・区切り文字の扱いを変えると、オフ→オンへ戻したときの挙動まで
'   同時に変わる ―― 波2の裁定(「PII緩和ロジックはオンに戻す際に再検討」)を
'   このFix波が先回りして踏み越えることになる。よって共有発信側だけに
'   前処理を置き、modPii の判定規則そのものには一切触らない。
'
'   潰し方: 日付/時刻とみなせる並びを "," 1文字へ置き換える。"," は modPii の
'   継続文字ではないのでランがそこで切れる。前後に残った数字は独立に数えられる
'   ので、日時の【隣に】本物の電話番号があれば従来どおり検知できる。
' ----------------------------------------------------------------------------
Public Function StripDateLike(ByVal s As String) As String
    Dim n As Long: n = Len(s)
    If n = 0 Then Exit Function
    Dim sb As String, i As Long, take As Long
    i = 1
    Do While i <= n
        take = DateLikeLen(s, i)
        If take > 0 Then
            sb = sb & ","
            i = i + take
        Else
            sb = sb & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    StripDateLike = sb
End Function

' i文字目から始まる「日付/時刻とみなせる並び」の長さ(無ければ0)。純関数。
'   日付: 4桁 + 区切り(- / .)+ 1〜2桁 + 同じ区切り + 1〜2桁   例 2026-08-14
'   時刻: 1〜2桁 + ":" + 2桁 [+ ":" + 2桁]                     例 10:00:00
' 直前が数字なら開始しない ―― 長い数字列(20260814100000)の途中を日付と
' 見なして切ってしまうと、本物の長い数字列を見逃す。
'
' R32 マイクロ修正波 F16[MAJOR]: 直後が数字のときも同じ理由で日付と見なさない。
'   【退行していたこと】市外局番4桁の固定電話(0463-12-3456 等)は
'   "4桁+区切り+1〜2桁+区切り+1〜2桁" の日付パターンにそのまま一致する
'   (0463-12-34 が「日付」、残る "56" は素通り)。潰すと "," に化けて
'   数字ランが切れ、本物のPII(10桁の電話番号)が検知されなくなっていた
'   (F4前は検知・F4後は不検知の退行)。
'   直し方: 一致した並びの直後がまだ数字なら、それは日付ではなく長い数字列の
'   一部という判断に倒す(直前を見る既存のガードと対称)。日付/時刻の後ろに
'   さらに数字が続く実例は無く(西暦の前・秒の後に数字が続く自然な日本語は
'   考えにくい)、既存の14ケース(2周目で実測検証済み)はいずれも直後が
'   非数字(空白・全角文字・文字列末尾)なので影響しない。
'   日付側の分岐は途中で Exit Function するため、判定はその出口と関数末尾
'   (時刻側)の両方に置く。
Private Function DateLikeLen(ByVal s As String, ByVal i As Long) As Long
    If i > 1 Then
        If IsDigitAt(s, i - 1) Then Exit Function
    End If
    If Not IsDigitAt(s, i) Then Exit Function

    Dim d1 As Long: d1 = DigitRun(s, i, 4)
    If d1 = 4 Then
        Dim sep As String: sep = Mid$(s, i + 4, 1)
        If sep = "-" Or sep = "/" Or sep = "." Then
            Dim p As Long: p = i + 5
            Dim d2 As Long: d2 = DigitRun(s, p, 2)
            If d2 >= 1 Then
                If Mid$(s, p + d2, 1) = sep Then
                    Dim d3 As Long: d3 = DigitRun(s, p + d2 + 1, 2)
                    If d3 >= 1 Then DateLikeLen = 4 + 1 + d2 + 1 + d3
                End If
            End If
        End If
    End If
    If DateLikeLen > 0 Then
        If IsDigitAt(s, i + DateLikeLen) Then DateLikeLen = 0   ' R32 F16
        Exit Function
    End If

    Dim h As Long: h = DigitRun(s, i, 2)
    If h < 1 Then Exit Function
    If Mid$(s, i + h, 1) <> ":" Then Exit Function
    If DigitRun(s, i + h + 1, 2) <> 2 Then Exit Function
    DateLikeLen = h + 1 + 2
    If Mid$(s, i + DateLikeLen, 1) = ":" Then
        If DigitRun(s, i + DateLikeLen + 1, 2) = 2 Then DateLikeLen = DateLikeLen + 3
    End If
    If IsDigitAt(s, i + DateLikeLen) Then DateLikeLen = 0        ' R32 F16
End Function

' i文字目から続く数字の個数(最大maxN)。純関数。
Private Function DigitRun(ByVal s As String, ByVal i As Long, ByVal maxN As Long) As Long
    Dim k As Long
    Do While k < maxN
        If Not IsDigitAt(s, i + k) Then Exit Do
        k = k + 1
    Loop
    DigitRun = k
End Function

' i文字目が半角数字か(範囲外はFalse)。純関数。
Private Function IsDigitAt(ByVal s As String, ByVal i As Long) As Boolean
    If i < 1 Or i > Len(s) Then Exit Function
    Dim c As String: c = Mid$(s, i, 1)
    IsDigitAt = (c >= "0" And c <= "9")
End Function

' ----------------------------------------------------------------------------
' ScanClean - 走査専用の改行つぶし(純関数・R32 マイクロ修正波 F17[MAJOR])。
' ----------------------------------------------------------------------------
'   【新たな偽陽性だったこと】送信本文を作る modInsightIo.Clean1 は、電文の
'   フィールド区切り(FIELD_SEP=タブ)と衝突しないよう vbLf/vbCr/vbTab を
'   【半角スペース】へ均す。ところが modPii.HasLongDigitRun は半角スペースを
'   「数字ランの継続」として数えるため、本来は改行で切れていた別々の数字列
'   (例: 1行目末尾の "12345" と2行目冒頭の "67890")が Clean1 を通した瞬間に
'   1本の10桁ランへ繋がり、無関係な数字の並びが個人情報として誤検知される
'   (例:"伝票番号を教えてください\n12345\n67890" が F4前=不検知→F4後=検知)。
'   PiiBlocked は「実際に送る内容」を走査する設計(F4)なので、走査対象は
'   送信本文と【同じ文面】であるべきだが、Clean1 が選んだ「半角スペースへ
'   均す」という【送信フォーマット上の都合】まで走査に持ち込む必要は無い。
'
'   直し方: 走査の直前だけもう1段整形を分ける。改行・タブは modPii の継続
'   文字("-"・半角スペース・U+2010)に含まれない "," へ落とす ―― これで
'   「行が変わる=数字列が別物」という事実を走査側に残したまま、"," の前後の
'   数字は独立に数えられるので本物のPII(1行の中に長い数字列)は従来どおり
'   検知できる。送信本文(Clean1)は【一切変えない】 ―― 電文はFIELD_SEPが
'   タブなので、本文中の改行をそのまま送るとフィールドがずれる。
'   modPii 本体には一切触れない(理由は StripDateLike のコメント参照)。
' ----------------------------------------------------------------------------
Public Function ScanClean(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace(t, vbCrLf, ",")
    t = Replace(t, vbCr, ",")
    t = Replace(t, vbLf, ",")
    t = Replace(t, vbTab, ",")
    ScanClean = t
End Function

' ----------------------------------------------------------------------------
' GapDupBlocked - 同じ趣旨の質問を短時間に連投していないか(R32 W1-6 M4)。
'   EmitGap には重複抑止が無く、しかも no_hit 経路は感想ボタンを通らないので
'   「1ターン1回」のガードすら掛からない。言い換えて3回聞けば3件投稿され、
'   板が1人の質問で埋まる(見る側からは同じ質問が3行並ぶ)。
'   照合は解決済みQ&A側の SameQuestionCount と同じ NormKey(記号と空白を
'   落とした先頭40字)。思想を揃えておかないと、片方で「同じ」もう片方で
'   「別」という食い違いが必ず出る。
'   抑止時間は config gap_dup_hours(既定24時間)。0以下で無効。
' ----------------------------------------------------------------------------
Public Function GapDupBlocked(ByVal qText As String) As Boolean
    On Error Resume Next
    Dim hours As Long: hours = modConfig.GetLong("gap_dup_hours", 24)
    If hours <= 0 Then Exit Function
    Dim k As String: k = modInsight.NormKey(qText)
    If LenB(k) = 0 Then Exit Function
    GapDupBlocked = modInsight.WithinWindow(modStats.GetStatText(GAPQ_PREFIX & k), _
                                 modUtilText.IsoDateTime(DateAdd("h", -hours, Now)))
    ' R32 F5: 抑止したことを利用者へ伝える(黙って落とさない)。
    If GapDupBlocked Then NotifySkip "dup"
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' NotifySkip - 部内共有を見送ったことを利用者へ知らせる(2026-08-14 R32 F5)。
' ----------------------------------------------------------------------------
'   modAsk の「だめだった」「微妙」は、押した直後に
'   「資料を作れる担当者の画面に届きます」と【断言する】MsgBox を出す。
'   ところが EmitGap は (1)PII検知で見送り (2)24時間の連投抑止
'   (3)共有フォルダへ書けない の3経路で無言のまま落ちるので、
'   「届きます」と言われたのに誰にも届いていない、が普通に起こる。
'   modAsk は凍結モジュール(CLAUDE.md)なのであの文言には触れない。代わりに
'   【見送ったときだけ】発信側からトーストを出して事実を上書きする。
'
'   waitless:=True にする理由: この関数は modAsk の MsgBox の直前、
'   および検索0件(no_hit)の自動発火の途中で呼ばれる。既定の待って消す
'   トーストは文字量に応じて3〜9秒ブロックするので、感想ボタンを押した人を
'   その分だけ待たせ、回答の描画も遅らせてしまう(残ったトーストは次の
'   ShowToast/PaintProgress が掃除する。modSkin.ShowToast のコメント参照)。
'   modSkin は30,000字上限まで残り27字のため、呼び出すだけで実体は置かない。
' ----------------------------------------------------------------------------
Public Sub NotifySkip(ByVal why As String)
    On Error Resume Next
    Dim msg As String
    Select Case why
        Case "pii"
            msg = "個人情報が含まれているかもしれないので、この質問は部内へ共有しませんでした。"
        Case "dup"
            msg = "同じ質問は24時間に1回だけ部内へ共有します(今回は共有していません)。"
        Case Else
            msg = "共有フォルダへ書き込めなかったため、今回は部内へ共有できませんでした。"
    End Select
    modSkin.ShowToast msg, "info", True
    On Error GoTo 0
End Sub

' 発信できたことを記録する(次の GapDupBlocked がこれを見る)。
Public Sub MarkGapEmitted(ByVal qText As String)
    On Error Resume Next
    Dim k As String: k = modInsight.NormKey(qText)
    If LenB(k) = 0 Then Exit Sub
    modStats.SetStatText GAPQ_PREFIX & k, modUtilText.IsoDateTime(Now)
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' GcGapDupKeys - 抑止期間を過ぎた連投抑止キーを my_stats から取り除く。
'   modInsightIo.GcOldNonces と同じ「下から消す」作法。掃除しないと、質問の
'   種類ぶんだけ my_stats が伸び、FindKeyRow の線形探索が重くなる
'   (HANDOFF §2 の既知課題・R32 r5 と同型)。
' ----------------------------------------------------------------------------
Public Sub GcGapDupKeys()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_STATS)
    If ws Is Nothing Then Exit Sub
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Sub

    Dim limit As String
    limit = modUtilText.IsoDateTime(DateAdd("d", -GAPQ_KEEP_DAYS, Now))
    ' A:B の2列で読む(1行しか無いときA列単独だと配列ではなくスカラーが返る
    ' VBAの仕様を避けるため。GcOldNonces と同じ回避策)。
    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastR, 2)).Value

    Dim i As Long
    For i = UBound(arr, 1) To LBound(arr, 1) Step -1
        Dim k As String: k = CStr(arr(i, 1))
        If Left$(k, Len(GAPQ_PREFIX)) = GAPQ_PREFIX Then
            If Not modInsight.WithinWindow(CStr(arr(i, 2)), limit) Then ws.Rows(i + 1).Delete
        End If
    Next i
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' NonceKeepDays / InboxKeepDays - 既読印と受信箱の保持日数(純関数・R32 F1)。
' ----------------------------------------------------------------------------
'   【BLOCKER だったこと】受信箱の行は INBOX_KEEP_DAYS=60日 の固定値で消えるのに、
'   既読印("ins:")は thanks_gc_days+7=67日 まで残る作りだった。61〜67日目は
'   「行は無いが既読印はある」で守られるが、68日目には【どちらも無い】瞬間が来る。
'   共有フォルダのファイルが残っていれば(発行担当が不在の組織では永久に残る)、
'   その日に全端末が過去分を一斉に再受信する。W1-2で入れた nonce 冪等化は
'   【受信箱にその行が在ること】を根拠にしているので、行が先に消える限り空振りし、
'   取り込み済みだったQ&Aが「未取込」に戻って本棚へ二重登録される。
'
'   直し方は「受信箱を既読印より長く持つ」の一点。既読印が切れて再読みされた
'   瞬間に行がまだ在れば、NonceIsKnownIn が重複を止め、そのついでに既読印が
'   書き直されて延命する(CollectFrom は AppendRow が True を返したら必ず
'   既読印を書く)。
'
'   【R32マイクロ修正波 F21訂正: これは「穴が消えた」のではない】延命される
'   のは既読印(この端末の収集時刻が起点)だけで、行の created_at(発信者側の
'   絶対時刻)は最初の1回から動かない。行は必ずいずれ消える ――延命は
'   その消える瞬間を「1周ぶん遠のける」だけで、恒久に閉じるわけではない。
'   延命後に既読印が次に切れる日が来れば(=行はとうに消えている)、また
'   「どちらも無い」瞬間に戻る。既定(thanks_gc_days=60)では実測で約day134に
'   再び穴が開き、thanks_gc_days<=0(例:0)なら約day15に前倒しになる。
'   B1の症状(移設直後の即時再来)を防ぐという当初の目的は満たすが、「穴を
'   永久に塞ぐ」設計はこのFix波の範囲外(必要ならnonceの延命ではなく
'   行そのものの生存期間を延ばす別設計が要る)。
'
'   【不変条件】任意の gcDays について
'       InboxKeepDays(gcDays) > NonceKeepDays(gcDays)
'   +7 / +14 という差の付け方は、config thanks_gc_days をいくつに変えても
'   この不等式が自動的に保たれるようにするため(片方だけ固定値にすると、
'   設定を変えた組織でだけ穴が開く)。定数をいじって不等式を壊す改修を
'   機械で止めるため、modTestsPure31 が境界値でこの不等式そのものを検算する。
' ----------------------------------------------------------------------------
Public Function NonceKeepDays(ByVal gcDays As Long) As Long
    NonceKeepDays = gcDays + 7
End Function

Public Function InboxKeepDays(ByVal gcDays As Long) As Long
    InboxKeepDays = gcDays + 14
End Function
