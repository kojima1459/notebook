Attribute VB_Name = "modTestsPure31"
Option Explicit

' ============================================================================
' modTestsPure31 - R32 Fix波(敵対的レビュー1周目の裁定)の純ロジック回帰テスト。
'   既存チェーン(modTestsPure.RunAll→…→modTestsPure30.RunAll30)へは繋がず、
'   modTestRunner.RunAllPureTests から直接呼ばれる RunAll31 の1本が入口
'   (modTestsPure32 と同型の別枝)。
' ----------------------------------------------------------------------------
' 【このモジュールが何を守るのか】
'   R32波2の時点では、ここは config の既定値3件(pii_scan_enabled /
'   gap_keep_days / gap_dup_hours)を modConfig.GetBool/GetLong で読んで
'   固定していた。しかしそれは【恒真】だった: LibreOffice の純ロジック実行
'   環境には config シートが存在せず、GetBool/GetLong はどのキーでも必ず
'   Fallback(=呼び出し側が渡した既定値)を返す。
'   `GetBool("pii_scan_enabled", False) = False` は「False = False」を
'   比べているだけで、config行を消しても値を変えても落ちない。
'   本当に守りたい「出荷する xlsm の config シートに正しい値の行が実在する
'   こと」は Python 側でしか確かめられないため、2026-08-14(R32 Fix波 F6)で
'   build/build_mybookshelf.py の自己検証(_verify_r32_config_defaults)へ
'   移し、ここは【VBAでしか書けない検査】だけを持つ形へ入れ替えた。
'
'   守る対象:
'     (1) 既読印GCと受信箱保持の不等式(F1。定数変更で不等式そのものが壊れる
'         事故を止める。ただしこの不等式が守るのは「穴を1周ぶん遠のける」
'         ところまでで、穴を恒久に閉じる保証ではない。詳細は
'         TestKeepDaysInequality31 のコメント参照。R32マイクロ修正波 F21)
'     (2) 共有発信側のPII前処理(F4/F16。日付・時刻の誤検知は通し、市外局番
'         4桁の固定電話のような本物の長い数字列は取り違えず止める)
'     (3) 発信者IDの匿名化(F2。実名がそのまま出ていないこと)
'     (4) 走査専用の改行つぶし(F17。改行をまたいだ別々の数字列を1本の
'         ランへ誤って繋げて誤検知しないこと)
' ============================================================================

' ----------------------------------------------------------------------------
' (1) R32 F1: 「既読印GC日数 < 受信箱保持日数」の不等式そのものを検算する。
'   この不等式が破れると、両方が同時に消える日が生まれ、共有ファイルが残って
'   いる組織で過去分の一斉再受信 → 本棚への二重登録が起きる(B1の再来)。
'   config thanks_gc_days は組織が自由に変えられるので、代表値だけでなく
'   境界(0・1)と極端な値でも成り立つことを確かめる。
'
'   【R32マイクロ修正波 F21訂正】この不等式は「穴が消える」ことの保証では
'   ない。既読印は「切れて再読みされた瞬間に行がまだ在れば書き直されて
'   延命する」(modInsightGate.NonceKeepDays/InboxKeepDays のコメント参照)が、
'   行の created_at(発信者側の絶対時刻)は最初の1回から動かないため、行は
'   いずれ必ず消える。延命でずれた既読印の次の期限が、行が消えた後にも
'   残っていれば当面は守られるが、その期限も来ればまた「どちらも無い」瞬間
'   (=穴)に戻る。この不等式が保証するのは「穴を(既読印の保持日数ぶん)
'   1周遠のける」ことまでで、既定(thanks_gc_days=60)では約day134に、
'   thanks_gc_days<=0(例:0)なら約day15に穴が再び開く。したがって以下の
'   検算は「不等式そのものが崩れていないか」(=B1の即時再来を防げているか)
'   だけを見るもので、「穴が二度と開かない」ことまでは検算していない。
' ----------------------------------------------------------------------------
Private Sub TestKeepDaysInequality31()
    Dim vals As Variant
    vals = Array(0, 1, 7, 30, 60, 90, 365, 3650)
    Dim i As Long
    For i = LBound(vals) To UBound(vals)
        Dim gc As Long: gc = CLng(vals(i))
        Dim nk As Long: nk = modInsightGate.NonceKeepDays(gc)
        Dim ik As Long: ik = modInsightGate.InboxKeepDays(gc)
        modTestRunner.Check "R32-F1_受信箱保持は既読印GCより長い(thanks_gc_days=" & gc & ")", _
            (ik > nk), "既読印=" & nk & " 受信箱=" & ik
    Next i

    ' 既定値(thanks_gc_days=60)での実数。67日 < 74日。
    modTestRunner.Check "R32-F1_既定60日なら既読印67日・受信箱74日", _
        (modInsightGate.NonceKeepDays(60) = 67 And modInsightGate.InboxKeepDays(60) = 74), _
        "既読印=" & modInsightGate.NonceKeepDays(60) & " 受信箱=" & modInsightGate.InboxKeepDays(60)
End Sub

' ----------------------------------------------------------------------------
' (2) R32 F4/F16: 共有発信側のPII前処理。
'   modPii.HasLongDigitRun は "-" と半角スペースを数字ランの継続として数える
'   ため、ISOのタイムスタンプ "2026-08-14 10:00" が 4+2+2+2+2=12桁として
'   【必ず】検知される。共有発信側の走査は常時ONなので、質問文に日時が1つ
'   入っているだけでその質問は永久に部内へ出ない(利用者には見えない)。
'   StripDateLike を通したあとの ScanText を直接撃つ ―― PiiBlocked 本体は
'   usage_log とトーストへ到達するのでここからは呼ばない。
'
'   【R32マイクロ修正波 F16[MAJOR]】市外局番4桁の固定電話(0463-12-3456等)は
'   「4桁+区切り+1〜2桁+区切り+1〜2桁」という日付パターンにそのまま一致して
'   しまい、先頭10字が "," へ潰されて本物のPIIが検知されなくなる退行が
'   あった(F4前=検知・F4後=不検知)。DateLikeLen に「一致の直後がまだ数字
'   なら日付ではない」判定を足して直した ―― 以下でこの3パターンが検知へ
'   戻ることと、上の13ケースが変わらないことの両方を確かめる。
' ----------------------------------------------------------------------------
Private Sub TestStripDateLike31()
    ' 通ってほしいもの(誤検知しない)
    CheckPass31 "ISOのタイムスタンプ", "2026-08-14 10:00"
    CheckPass31 "秒まで付いたタイムスタンプ", "2026-08-14 10:00:00"
    CheckPass31 "日本語に埋め込んだ日時", "打合せは2026-08-14 10:00からです"
    CheckPass31 "スラッシュ区切りの日付", "2026/08/14 09:30 に確認しました"
    CheckPass31 "日付が2つ並ぶ", "2026-08-01 から 2026-08-31 まで"
    CheckPass31 "年月日と時刻", "2026年8月14日 10:00 の会議"
    CheckPass31 "日付のあとに短い番号", "2026-08-14 の第3号"

    ' 止まってほしいもの(本物は見逃さない)
    CheckBlock31 "ハイフン区切りの携帯番号", "連絡先は090-1234-5678です"
    CheckBlock31 "空白区切りの電話番号", "03 1234 5678 へお願いします"
    CheckBlock31 "12桁の数字列", "契約番号123456789012を確認"
    CheckBlock31 "メールアドレス", "tanaka@example.co.jp へ送付"
    ' 日付の【隣に】本物があるときも見逃さない(前処理がランを切るだけで、
    ' 前後の数字を消してしまわないことの確認)。
    CheckBlock31 "日時と電話番号が同居", "2026-08-14 10:00 に 090-1234-5678 へ架電"

    ' R32マイクロ修正波 F16: 市外局番4桁の固定電話は日付と誤認しない
    ' (退行前の状態=本物のPIIとして検知される、へ戻ったことの確認)。
    CheckBlock31 "市外局番4桁の固定電話(0463)", "固定電話は0463-12-3456です"
    CheckBlock31 "市外局番4桁の固定電話(0267)", "固定電話は0267-42-1234です"
    CheckBlock31 "市外局番4桁の固定電話(0120フリーダイヤル)", "フリーダイヤルは0120-12-3456です"
    ' 前処理そのものの形も固定する(直後が数字なら "," へ潰さず素通しする)。
    modTestRunner.Check "R32-F16_市外局番4桁の固定電話は前処理で変化しない(0463)", _
        (modInsightGate.StripDateLike("0463-12-3456") = "0463-12-3456"), _
        "実際=[" & modInsightGate.StripDateLike("0463-12-3456") & "]"
    modTestRunner.Check "R32-F16_市外局番4桁の固定電話は前処理で変化しない(0267)", _
        (modInsightGate.StripDateLike("0267-42-1234") = "0267-42-1234"), _
        "実際=[" & modInsightGate.StripDateLike("0267-42-1234") & "]"
    modTestRunner.Check "R32-F16_市外局番4桁の固定電話は前処理で変化しない(0120)", _
        (modInsightGate.StripDateLike("0120-12-3456") = "0120-12-3456"), _
        "実際=[" & modInsightGate.StripDateLike("0120-12-3456") & "]"

    ' 前処理そのものの形も1つ固定する(日付が消えて前後の文字は残る)。
    modTestRunner.Check "R32-F4_日付は区切り文字へ置き換わり前後は残る", _
        (modInsightGate.StripDateLike("A2026-08-14B") = "A,B"), _
        "実際=[" & modInsightGate.StripDateLike("A2026-08-14B") & "]"
    ' 長い数字列の途中を日付と誤認して切らない(切ると本物を見逃す)。
    modTestRunner.Check "R32-F4_長い数字列の途中は日付と見なさない", _
        (modInsightGate.StripDateLike("20260814100000") = "20260814100000"), _
        "実際=[" & modInsightGate.StripDateLike("20260814100000") & "]"
End Sub

Private Sub CheckPass31(ByVal label As String, ByVal s As String)
    Dim hit As String: hit = modPii.ScanText(modInsightGate.StripDateLike(s))
    modTestRunner.Check "R32-F4_誤検知しない: " & label, _
        (LenB(hit) = 0), "検知=" & hit & " / 前処理後=[" & modInsightGate.StripDateLike(s) & "]"
End Sub

Private Sub CheckBlock31(ByVal label As String, ByVal s As String)
    Dim hit As String: hit = modPii.ScanText(modInsightGate.StripDateLike(s))
    modTestRunner.Check "R32-F4_本物は止める: " & label, _
        (LenB(hit) > 0), "前処理後=[" & modInsightGate.StripDateLike(s) & "]"
End Sub

' ----------------------------------------------------------------------------
' (3) R32 F2: 発信者IDの匿名化。共有フォルダのファイル名と電文の user_id 欄に
'   使う値なので、「実名がそのまま出ない」「同じ人からは常に同じ」
'   「別人とは別の値」の3つが崩れると、匿名化かIsMineのどちらかが壊れる。
' ----------------------------------------------------------------------------
Private Sub TestAnonId31()
    Const NAME1 As String = "山田 太郎"
    Const NAME2 As String = "山田 次郎"
    Dim a As String: a = modInsightGate.AnonId(NAME1)

    modTestRunner.Check "R32-F2_実名がそのまま出ない", _
        (InStr(a, NAME1) = 0 And a <> NAME1), "実際=" & a
    ' modUtil.U32ToHex8 は【小文字】の16進を返す(Fnv1a64Hex はそれを2本連結)。
    ' 大文字前提で書くとここで落ちる ―― 実際にこのテストがそれを捕まえた。
    modTestRunner.Check "R32-F2_16桁の16進になる(ファイル名に使える文字だけ)", _
        (Len(a) = 16 And IsHex31(a)), "実際=" & a
    modTestRunner.Check "R32-F2_同じ人からは常に同じ値(IsMineが成立する条件)", _
        (modInsightGate.AnonId(NAME1) = a), "実際=" & modInsightGate.AnonId(NAME1)
    modTestRunner.Check "R32-F2_別人は別の値(他人の投稿を自分発と誤認しない)", _
        (modInsightGate.AnonId(NAME2) <> a), "実際=" & modInsightGate.AnonId(NAME2)
    modTestRunner.Check "R32-F2_IDが取れないときは空を返す(生IDへ落とさない)", _
        (LenB(modInsightGate.AnonId("")) = 0), "実際=[" & modInsightGate.AnonId("") & "]"
End Sub

' 16進16桁か(0-9a-f。大文字小文字は問わない)。ファイル名に使える文字だけで
' できていること=共有フォルダへ置くファイル名として安全であることの確認。
Private Function IsHex31(ByVal s As String) As Boolean
    Dim t As String: t = LCase$(s)
    Dim i As Long, c As String
    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If Not ((c >= "0" And c <= "9") Or (c >= "a" And c <= "f")) Then Exit Function
    Next i
    IsHex31 = (Len(t) > 0)
End Function

' ----------------------------------------------------------------------------
' (4) R32マイクロ修正波 F17[MAJOR]: 走査専用の改行つぶし(ScanClean)。
'   modInsightIo.Clean1 は送信本文を作るために vbLf/vbCr/vbTab を半角
'   スペースへ均すが、modPii.HasLongDigitRun は半角スペースを数字ランの
'   継続として数えるため、Clean1 済みの文字列をそのままPII走査へ渡すと、
'   改行で本来切れていた別々の数字列が1本のランに繋がって新たな偽陽性に
'   なる(例:「12345」の次行に「67890」があるだけの質問文が検知される)。
'   ScanClean は改行等を modPii の継続文字でない "," へ落とすことでこれを防ぐ。
'   ここでは ScanClean そのものの形と、ScanClean後の ScanText の判定結果
'   (誤検知しない/本物は止める)の両方を確かめる。
' ----------------------------------------------------------------------------
Private Sub TestScanClean31()
    ' ScanCleanの形そのものを固定する(改行・復帰・タブが "," へ落ちる)。
    Dim mixed As String
    mixed = "A" & vbLf & "B" & vbCr & "C" & vbTab & "D" & vbCrLf & "E"
    modTestRunner.Check "R32-F17_ScanCleanは改行等をカンマへ落とす", _
        (modInsightGate.ScanClean(mixed) = "A,B,C,D,E"), _
        "実際=[" & modInsightGate.ScanClean(mixed) & "]"

    ' 改行をまたいだ別々の数字列は繋がず、誤検知しない(F17の本題)。
    CheckPassScan31 "改行で分かれた数字列(LF)", _
        "伝票番号を教えてください" & vbLf & "12345" & vbLf & "67890"
    CheckPassScan31 "改行で分かれた数字列(CRLF)", _
        "在庫が" & vbCrLf & "12345" & vbCrLf & "67890" & vbCrLf & "個ありません"
    CheckPassScan31 "タブで分かれた数字列", _
        "コード" & vbTab & "12345" & vbTab & "67890"

    ' 1行の中の本物の長い数字列は、改行が別の場所にあっても見逃さない。
    CheckBlockScan31 "改行の別の行に本物の電話番号", _
        "在庫確認をお願いします" & vbLf & "担当は090-1234-5678です"
    CheckBlockScan31 "1行内の12桁の数字列(改行なし)", "契約番号123456789012を確認"
End Sub

' PiiBlockedの実引数と同じ経路(ScanClean→StripDateLike→ScanText)を
' modInsightGate.PiiBlockedへは通さず直接確認する(あちらは usage_log/
' トーストへ到達するため)。
Private Sub CheckPassScan31(ByVal label As String, ByVal s As String)
    Dim cleaned As String: cleaned = modInsightGate.ScanClean(s)
    Dim hit As String: hit = modPii.ScanText(modInsightGate.StripDateLike(cleaned))
    modTestRunner.Check "R32-F17_誤検知しない: " & label, _
        (LenB(hit) = 0), "検知=" & hit & " / 走査整形後=[" & cleaned & "]"
End Sub

Private Sub CheckBlockScan31(ByVal label As String, ByVal s As String)
    Dim cleaned As String: cleaned = modInsightGate.ScanClean(s)
    Dim hit As String: hit = modPii.ScanText(modInsightGate.StripDateLike(cleaned))
    modTestRunner.Check "R32-F17_本物は止める: " & label, _
        (LenB(hit) > 0), "走査整形後=[" & cleaned & "]"
End Sub

' ============================================================================
Public Sub RunAll31()
    On Error GoTo Fail31
    TestKeepDaysInequality31
    TestStripDateLike31
    TestAnonId31
    TestScanClean31
NextDone31:
    On Error GoTo 0
    Exit Sub

Fail31:
    modTestRunner.Check "TestKeepDaysInequality31/TestStripDateLike31/TestAnonId31/TestScanClean31(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone31
End Sub
