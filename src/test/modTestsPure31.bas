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
'     (1) 既読印GCと受信箱保持の不等式(F1。定数変更で穴が開くのを止める)
'     (2) 共有発信側のPII前処理(F4。日付・時刻の誤検知を通し、本物は止める)
'     (3) 発信者IDの匿名化(F2。実名がそのまま出ていないこと)
' ============================================================================

' ----------------------------------------------------------------------------
' (1) R32 F1: 「既読印GC日数 < 受信箱保持日数」の不等式そのものを検算する。
'   この不等式が破れると、両方が同時に消える日が生まれ、共有ファイルが残って
'   いる組織で過去分の一斉再受信 → 本棚への二重登録が起きる(B1の再来)。
'   config thanks_gc_days は組織が自由に変えられるので、代表値だけでなく
'   境界(0・1)と極端な値でも成り立つことを確かめる。
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
' (2) R32 F4: 共有発信側のPII前処理。
'   modPii.HasLongDigitRun は "-" と半角スペースを数字ランの継続として数える
'   ため、ISOのタイムスタンプ "2026-08-14 10:00" が 4+2+2+2+2=12桁として
'   【必ず】検知される。共有発信側の走査は常時ONなので、質問文に日時が1つ
'   入っているだけでその質問は永久に部内へ出ない(利用者には見えない)。
'   StripDateLike を通したあとの ScanText を直接撃つ ―― PiiBlocked 本体は
'   usage_log とトーストへ到達するのでここからは呼ばない。
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

' ============================================================================
Public Sub RunAll31()
    On Error GoTo Fail31
    TestKeepDaysInequality31
    TestStripDateLike31
    TestAnonId31
NextDone31:
    On Error GoTo 0
    Exit Sub

Fail31:
    modTestRunner.Check "TestKeepDaysInequality31/TestStripDateLike31/TestAnonId31(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone31
End Sub
