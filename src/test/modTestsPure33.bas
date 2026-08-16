Attribute VB_Name = "modTestsPure33"
Option Explicit

' ============================================================================
' modTestsPure33 - R33波2(個人情報・共有の関所)の純ロジック回帰テスト。
'   既存チェーン(modTestsPure.RunAll→…→modTestsPure30.RunAll30)へは繋がず、
'   modTestRunner.RunAllPureTests から直接呼ばれる RunAll33 の1本が入口
'   (modTestsPure31 / modTestsPure32 と同型の別枝)。
'   置き場が新規モジュールなのは容量: modTestsPure は30,000字上限まで
'   残り1,087字しかなく、ここのゴールデン一式(約6,000字)が入らない。
' ----------------------------------------------------------------------------
' 【このモジュールが何を守るのか】
'   (1) W2-1: modPii の全角ゴールデン。全角数字の電話番号・全角ハイフン類・
'       全角＠のメールが、半角で書いたときと同じように検知されること。
'       あわせて「判定閾値(10桁)と既存の区切り規則は変えていない」ことを
'       負例(全角9桁は非検知)で固定する。
'   (2) W2-2: modShareRule.ExpiryDecision の不変条件(4)。今つながっている
'       端末は【絶対に消さない】。同時に、逆向きの検算 ―― 本当に届かない
'       端末では従来どおり "wipe" が出る ―― も総当たりで固定する。
'       この2つは対で意味を持つ(片方だけだと「何も消さない」実装でも
'       通ってしまう / 「常に消す」実装でも通ってしまう)。
'
' 【非ASCIIリテラルの書き方】
'   区切り文字のうち U+2212(MINUS SIGN)は CP932 では U+FF0D と同じバイト
'   0x817C へ潰れる。ソースへ生の文字を置くと、VBEへ注入して読み戻した時点で
'   U+FF0D に化けてしまい、「U+2212 を検査したつもりで U+FF0D を2回検査して
'   いる」という嘘のテストになる。よって区切り文字は全て ChrW で組む。
'   全角数字・全角＠は CP932 で一意に往復するのでリテラルで書く。
' ============================================================================

' ----------------------------------------------------------------------------
' (1) W2-1: modPii の全角ゴールデン。
' ----------------------------------------------------------------------------
Private Sub TestPiiWide33()
    ' --- 全角数字の電話番号(区切りは4種類の全角ハイフン類すべて) -----------
    CheckHit33 "全角電話番号_全角ハイフンマイナス(U+FF0D)", _
        "田中様の携帯 ０９０" & ChrW(&HFF0D&) & "１２３４" & ChrW(&HFF0D&) & "５６７８ に折り返す", _
        "長い数字列"
    CheckHit33 "全角電話番号_MINUS SIGN(U+2212)", _
        "０９０" & ChrW(&H2212&) & "１２３４" & ChrW(&H2212&) & "５６７８", _
        "長い数字列"
    CheckHit33 "全角電話番号_長音符(U+30FC)", _
        "０９０" & ChrW(&H30FC&) & "１２３４" & ChrW(&H30FC&) & "５６７８", _
        "長い数字列"
    CheckHit33 "全角電話番号_全角チルダ(U+FF5E)", _
        "０９０" & ChrW(&HFF5E&) & "１２３４" & ChrW(&HFF5E&) & "５６７８", _
        "長い数字列"

    ' --- 区切り無しの全角10桁(マイナンバー等) ------------------------------
    CheckHit33 "全角10桁_区切り無し", "マイナンバーは１２３４５６７８９０です", "長い数字列"

    ' --- 半角と全角が混ざった実際の打ち間違い ------------------------------
    CheckHit33 "全角と半角の混在", "090" & ChrW(&HFF0D&) & "１２３４-5678", "長い数字列"

    ' --- 全角＠のメール ----------------------------------------------------
    CheckHit33 "全角＠のメール", "連絡は taro＠example.co.jp まで", "メールアドレス"
    CheckHit33 "全角＠のメール_全角ハイフン混在", _
        "taro.yamada＠example.co.jp", "メールアドレス"

    ' --- 負例: 判定閾値(10桁)を変えていないこと ----------------------------
    '   全角9桁は【検知しない】。ここが検知になったら、幅を均すついでに
    '   閾値か区切り規則を触っている(R32波2の裁定を先回りしている)。
    CheckNone33 "全角9桁は非検知(閾値10桁は不変)", "商品コード ００１２３４５６７"
    CheckNone33 "全角9桁_全角ハイフン区切りでも非検知", _
        "００１" & ChrW(&HFF0D&) & "２３４" & ChrW(&HFF0D&) & "５６７"

    ' --- 負例: 全角＠が無ければメール扱いしない ----------------------------
    CheckNone33 "全角のふつうの日本語文は非検知", "本日の会議は１０時から会議室Ａで行います。"

    ' --- 半角側の既存挙動が1つも変わっていないこと -------------------------
    '   (正規化は全角だけを均す。半角入力の答えは従来と同一でなければならない)
    CheckHit33 "半角_既存の電話番号は従来どおり検知", "電話番号: 090-1234-5678 です", "長い数字列"
    CheckHit33 "半角_既存のメールは従来どおり検知", "taro.yamada@example.co.jp", "メールアドレス"
    CheckNone33 "半角_9桁は従来どおり非検知", "商品コード: 123456789"
    CheckNone33 "半角_通常文章は従来どおり非検知", "本日の会議は10時から会議室Aで行います。"
End Sub

Private Sub CheckHit33(ByVal label As String, ByVal s As String, ByVal wantCat As String)
    Dim hit As String: hit = modPii.ScanText(s)
    modTestRunner.Check "R33-W2-1_検知する: " & label, _
        (InStr(hit, wantCat) > 0), "検知=[" & hit & "] 期待に含む=[" & wantCat & "]"
End Sub

Private Sub CheckNone33(ByVal label As String, ByVal s As String)
    Dim hit As String: hit = modPii.ScanText(s)
    modTestRunner.Check "R33-W2-1_検知しない: " & label, _
        (LenB(hit) = 0), "検知=[" & hit & "]"
End Sub

' ----------------------------------------------------------------------------
' (2) W2-2: 今つながっている端末は消さない(不変条件4)。
'   実機で起きていた事故そのものを1本目に置く ―― 管理者が
'   knowledge_expire_days=30 を設定 / 利用者が26日目に社外で予告を受け
'   (warned="26") / 34日目に出社し社内NWに接続した状態で開く。
'   旧実装はここで "wipe" を返し、本棚が全消去されていた。
' ----------------------------------------------------------------------------
Private Sub TestExpiryReachable33()
    Dim reach As String: reach = "2026-07-01"

    ' --- 事故の再現ケース --------------------------------------------------
    modTestRunner.Check "R33-W2-2_復帰した端末は消さない(実機再現)", _
        (modShareRule.ExpiryDecision(reach, 34, 30, "26", "", True) = "none"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 34, 30, "26", "", True)

    ' --- 同じ材料で reachableNow だけを落とすと従来どおり "wipe" ----------
    '   (この対が無いと「何も消さない」実装でも上の1本は通ってしまう)
    modTestRunner.Check "R33-W2-2_到達不能なら従来どおり消す(明示False)", _
        (modShareRule.ExpiryDecision(reach, 34, 30, "26", "", False) = "wipe"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 34, 30, "26", "", False)
    modTestRunner.Check "R33-W2-2_到達不能なら従来どおり消す(引数省略)", _
        (modShareRule.ExpiryDecision(reach, 34, 30, "26", "") = "wipe"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 34, 30, "26", "")

    ' --- つながっていれば予告も出さない ------------------------------------
    modTestRunner.Check "R33-W2-2_つながっていればwarnも出さない", _
        (modShareRule.ExpiryDecision(reach, 29, 30, "", "", True) = "none"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 29, 30, "", "", True)
    modTestRunner.Check "R33-W2-2_同じ材料でFalseならwarn", _
        (modShareRule.ExpiryDecision(reach, 29, 30, "", "", False) = "warn"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 29, 30, "", "", False)
    modTestRunner.Check "R33-W2-2_つながっていればwarn_firstも出さない", _
        (modShareRule.ExpiryDecision(reach, 30, 30, "", "", True) = "none"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 30, 30, "", "", True)
    modTestRunner.Check "R33-W2-2_同じ材料でFalseならwarn_first", _
        (modShareRule.ExpiryDecision(reach, 30, 30, "", "", False) = "warn_first"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 30, 30, "", "", False)

    ' --- 先に立つ2語("off"/"never")は reachableNow に食われない -----------
    '   "never" は「一度も到達した記録が無い」という別の防壁で、呼び出し側が
    '   usage_log へ1行残す分岐を持つ。ここが "none" に化けるとその記録が
    '   消え、共有パス誤設定の調査手掛かりが失われる。
    modTestRunner.Check "R33-W2-2_neverはreachableNowより優先", _
        (modShareRule.ExpiryDecision("", 34, 30, "26", "", True) = "never"), _
        "実際=" & modShareRule.ExpiryDecision("", 34, 30, "26", "", True)
    modTestRunner.Check "R33-W2-2_offはreachableNowより優先", _
        (modShareRule.ExpiryDecision(reach, 34, 0, "26", "", True) = "off"), _
        "実際=" & modShareRule.ExpiryDecision(reach, 34, 0, "26", "", True)

    ' --- 総当たり(1): つながっている限り wipe も warn も絶対に出ない -------
    Dim L As Long, d As Long
    Dim badWhenReachable As String
    For L = 1 To 60
        For d = 0 To 200
            Dim a1 As String
            a1 = modShareRule.ExpiryDecision(reach, d, L, CStr(d), "", True)
            If a1 <> "none" Then
                If LenB(badWhenReachable) = 0 Then badWhenReachable = "L=" & L & " d=" & d & " -> " & a1
            End If
            a1 = modShareRule.ExpiryDecision(reach, d, L, "999", "", True)
            If a1 <> "none" Then
                If LenB(badWhenReachable) = 0 Then badWhenReachable = "L=" & L & " d=" & d & " warned=999 -> " & a1
            End If
        Next d
    Next L
    modTestRunner.Check "R33-W2-2_到達中は全組合せで none のみ", _
        (LenB(badWhenReachable) = 0), "初回の違反: " & badWhenReachable

    ' --- 総当たり(2): 逆向きの検算 -----------------------------------------
    '   「本来消すべきケースを消さなくなっていないか」。到達不能・期限超過・
    '   予告済み・未消去 の4条件が揃った端末は、従来どおり必ず "wipe"。
    Dim missedWipe As String
    For L = 1 To 60
        For d = L To L + 60
            Dim a2 As String
            a2 = modShareRule.ExpiryDecision(reach, d, L, CStr(d - 1), "", False)
            If a2 <> "wipe" Then
                If LenB(missedWipe) = 0 Then missedWipe = "L=" & L & " d=" & d & " -> " & a2
            End If
        Next d
    Next L
    modTestRunner.Check "R33-W2-2_到達不能で期限超過かつ予告済みなら必ずwipe", _
        (LenB(missedWipe) = 0), "初回の取りこぼし: " & missedWipe
End Sub

' ----------------------------------------------------------------------------
' 入口。群ごとにハンドラを分ける(1本のハンドラだと最初の群で落ちた時点で
' 残りが無言で消える。R33波1 W1-3 で実害が出た型)。
' ----------------------------------------------------------------------------
Public Sub RunAll33()
    On Error GoTo G01Fail33
    TestPiiWide33
G02Next33:
    On Error GoTo G02Fail33
    TestExpiryReachable33
NextDone33:
    On Error GoTo 0
    Exit Sub

G01Fail33:
    GroupFail33 "TestPiiWide33", Err.Number, Err.Description
    Resume G02Next33
G02Fail33:
    GroupFail33 "TestExpiryReachable33", Err.Number, Err.Description
    Resume NextDone33
End Sub

Private Sub GroupFail33(ByVal groupName As String, ByVal errNum As Long, ByVal errDesc As String)
    modTestRunner.Check groupName & "(グループ全体)", False, _
        "群の実行中に例外: " & errDesc & " (Err=" & errNum & ")"
End Sub
