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

    ' --- W2-8: 全角英字・全角ピリオドで書かれたメール ----------------------
    '   ＠だけ半角にしても、前後が全角のままでは
    '   `Like "*[A-Za-z0-9.]@[A-Za-z0-9.]*"` を満たせず素通りしていた。
    CheckHit33 "全角のメール_小文字", "ｔａｒｏ＠ｅｘａｍｐｌｅ．ｃｏｍ", "メールアドレス"
    CheckHit33 "全角のメール_大文字", "ＴＡＲＯ＠ＥＸＡＭＰＬＥ．ＣＯＭ", "メールアドレス"
    CheckHit33 "全角のメール_半角と混在", "連絡先 taro＠ｅｘａｍｐｌｅ.co.jp", "メールアドレス"
    CheckHit33 "全角のメール_数字入りローカル部", "ｙ１２３＠ｅｘ．ｊｐ", "メールアドレス"

    ' --- W2-8 の負例: 英字を均しても判定規則そのものは動かないこと --------
    '   (1) ＠が無ければメール扱いしない
    CheckNone33 "全角英字だけでは非検知(＠が無い)", "ｅｘａｍｐｌｅ ｃｏｍ の資料"
    '   (2) 英字は10桁ランを作らない=閾値と区切り規則は不変
    CheckNone33 "全角英字26字でも数字列にはならない", "ＡＢＣＤＥＦＧＨＩＪＫＬＭＮ"
    '   (3) 全角ピリオドは区切り文字ではない(ランを切る側のまま)
    CheckNone33 "全角ピリオド区切りの数字は繋がらない", _
        "２３４５６．６７８９０ の伝票"
    '   (4) ＠の直前が日本語なら従来どおり非検知(単価表記の素通りを確認)
    CheckNone33 "＠の直前が日本語なら非検知", "りんご＠１２０円 で仕入れ"

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
' (3) W2-7: 共有発信側の関所で、全角の日付を誤検知しない。
' ----------------------------------------------------------------------------
'   W2-1(ScanText の内側で幅を均す)が作った退行を止める群。
'   modInsightGate.PiiBlocked の前処理は「幅を均す→StripDateLike→ScanText」で、
'   StripDateLike は半角数字専用。均しを ScanText の内側だけに置くと、
'   全角の日付が日付潰しを素通りしたまま内側で半角化され、4+2+2+2+2=12桁の
'   ランとして検知される ―― R32 F4 が半角側で潰した誤検知の全角版で、
'   質問が【無言で】共有見送りになる。
'
'   【PiiBlocked を直接叩く群と、前処理の並びを写す群を分けている理由】
'   PiiBlocked は「検知しなかった」経路では 92行目の
'   `If LenB(hit) = 0 Then Exit Function` で即抜けるので副作用がゼロ
'   (ログもトーストも通らない)。よって【見送られないこと】の検査は
'   実物を直接叩ける ―― 前処理の並びが将来また入れ替わってもここが落ちる。
'   一方【止まること】の検査で実物を叩くと NotifySkip → modSkin.ShowToast
'   (Shape描画)まで走ってしまうため、そちらは modTestsPure31 の
'   CheckPass31/CheckBlock31 と同じ作法で前処理の並びを写して検査する。
' ----------------------------------------------------------------------------
Private Sub TestGateWideDate33()
    ' --- 実物を叩く: 全角の日付を含む質問は見送られない -------------------
    modTestRunner.Check "R33-W2-7_全角日付の質問は見送られない(実物)", _
        (modInsightGate.PiiBlocked("２０２６" & ChrW(&HFF0D&) & "０８" & ChrW(&HFF0D&) & _
            "１４ １０:００ の締切について教えてください", "gap") = False), _
        "PiiBlocked が True を返した(=共有が無言で見送られる)"
    ' 区切りが U+2212(MINUS SIGN)でも同じこと(日付潰しの手前で均されている)。
    modTestRunner.Check "R33-W2-7_全角日付(U+2212区切り)も見送られない(実物)", _
        (modInsightGate.PiiBlocked("２０２６" & ChrW(&H2212&) & "０８" & ChrW(&H2212&) & _
            "１４ ０９:３０ の議事録はありますか", "gap") = False), _
        "PiiBlocked が True を返した"
    ' 半角の日付(R32 F4 が守っている側)が巻き添えで壊れていないこと。
    '   この1本は W2-7 の均しを外しても通る ―― 半角側は元から
    '   StripDateLike が拾えるため。W2-7 を discriminate する検査ではなく、
    '   「全角対応の巻き添えで半角側を壊していないこと」を見る番人。
    modTestRunner.Check "R33-W2-7_半角日付も従来どおり見送られない(実物)", _
        (modInsightGate.PiiBlocked("2026-08-14 10:00 の締切について教えてください", "gap") = False), _
        "PiiBlocked が True を返した"

    ' --- 前処理の並びを写す: 本物の全角電話番号は従来どおり止まる ---------
    '   日付潰しが電話番号まで食べていたら、ここが落ちる。
    CheckGateBlock33 "全角の携帯番号", _
        "田中様の携帯 ０９０" & ChrW(&HFF0D&) & "１２３４" & ChrW(&HFF0D&) & "５６７８ に折り返す"
    CheckGateBlock33 "全角の固定電話(市外局番4桁)", _
        "０１２０" & ChrW(&HFF0D&) & "１２３" & ChrW(&HFF0D&) & "４５６７ へ連絡"
    CheckGateBlock33 "全角＠のメール", "連絡は taro＠example.co.jp まで"

    ' --- 冪等性: 二重に均しても結果は同じ ---------------------------------
    '   PiiBlocked は入口で1回、ScanText の内側でもう1回通す。ここが冪等で
    '   ないと「1回目と2回目で答えが変わる」ため、内側の均しを残せない。
    CheckIdem33 "全角電話番号", _
        "０９０" & ChrW(&HFF0D&) & "１２３４" & ChrW(&H2212&) & "５６７８"
    CheckIdem33 "全角日付", "２０２６" & ChrW(&HFF0D&) & "０８" & ChrW(&HFF0D&) & "１４ １０:００"
    CheckIdem33 "全角＠と長音符", "taro＠ex.com サーバー障害"
    CheckIdem33 "全角の英字・ピリオド(W2-8)", "ｔａｒｏ＠ｅｘａｍｐｌｅ．ｃｏｍ"
    CheckIdem33 "半角のみ", "090-1234-5678 / a@b.com"
End Sub

Private Sub CheckGateBlock33(ByVal label As String, ByVal s As String)
    ' modInsightGate.PiiBlocked の前処理と同じ並び(実物は副作用があるため写す)。
    Dim hit As String
    hit = modPii.ScanText(modInsightGate.StripDateLike(modPii.NormalizeWidth(s)))
    modTestRunner.Check "R33-W2-7_関所で止まる: " & label, _
        (LenB(hit) > 0), _
        "前処理後=[" & modInsightGate.StripDateLike(modPii.NormalizeWidth(s)) & "]"
End Sub

Private Sub CheckIdem33(ByVal label As String, ByVal s As String)
    Dim once_ As String: once_ = modPii.NormalizeWidth(s)
    Dim twice_ As String: twice_ = modPii.NormalizeWidth(once_)
    modTestRunner.Check "R33-W2-7_NormalizeWidthは冪等: " & label, _
        (once_ = twice_), "1回=[" & once_ & "] 2回=[" & twice_ & "]"
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
' (4) W2-9: 部門名でフォルダの外へ逃げられないこと。
' ----------------------------------------------------------------------------
'   W2-5 で単独の "\" を弾いたが ".." 単体はまだ通り、
'   <共有>\channels\..\ = 共有ルートへ正典一式を書き出せた。同じ「発行先を
'   共有ルートへ逃がされる」穴なので、片方だけでは意味が無い。
'   先頭・末尾の空白/ピリオドは、Windowsが黙って落とすことによる
'   「別名のつもりで既存部門を上書き発行」を止めるための検査。
' ----------------------------------------------------------------------------
Private Sub TestChannelNameEscape33()
    ' --- 弾くべき名前 ------------------------------------------------------
    CheckEscBad33 "..単体", ".."
    CheckEscBad33 "..を含む(前)", "..\商品部"
    CheckEscBad33 "..を含む(後ろ)", "商品部..仮"
    CheckEscBad33 "先頭ピリオド", ".商品部"
    CheckEscBad33 "末尾ピリオド", "商品部."
    CheckEscBad33 "単独ピリオド", "."
    CheckEscBad33 "先頭が全角空白", ChrW(&H3000&) & "商品部"
    CheckEscBad33 "末尾が全角空白", "商品部" & ChrW(&H3000&)
    CheckEscBad33 "先頭が半角空白", " 商品部"
    CheckEscBad33 "末尾が半角空白", "商品部 "
    CheckEscBad33 "末尾がタブ", "商品部" & vbTab

    ' --- 通すべき名前(正当な部門名を巻き添えにしない) ---------------------
    CheckEscOk33 "ふつうの部門名", "商品部"
    CheckEscOk33 "課まで入った名前", "商品部商品1課"
    CheckEscOk33 "中黒や括弧を含む名前", "商品部(第2)"
    CheckEscOk33 "英字と数字の名前", "Sales2026"
    CheckEscOk33 "途中のピリオド1個", "商品部.仮"
    CheckEscOk33 "途中の空白", "商品 部"
    CheckEscOk33 "空文字は他の検査の担当", ""
End Sub

Private Sub CheckEscBad33(ByVal label As String, ByVal s As String)
    Dim r As String: r = modShareRule.ChannelPathEscapeReason(s)
    modTestRunner.Check "R33-W2-9_弾く: " & label, _
        (LenB(r) > 0), "理由が空だった(=この名前で発行できてしまう)"
End Sub

Private Sub CheckEscOk33(ByVal label As String, ByVal s As String)
    Dim r As String: r = modShareRule.ChannelPathEscapeReason(s)
    modTestRunner.Check "R33-W2-9_通す: " & label, _
        (LenB(r) = 0), "理由=[" & r & "]"
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
G03Next33:
    On Error GoTo G03Fail33
    TestGateWideDate33
G04Next33:
    On Error GoTo G04Fail33
    TestChannelNameEscape33
NextDone33:
    On Error GoTo 0
    Exit Sub

G01Fail33:
    GroupFail33 "TestPiiWide33", Err.Number, Err.Description
    Resume G02Next33
G02Fail33:
    GroupFail33 "TestExpiryReachable33", Err.Number, Err.Description
    Resume G03Next33
G03Fail33:
    GroupFail33 "TestGateWideDate33", Err.Number, Err.Description
    Resume G04Next33
G04Fail33:
    GroupFail33 "TestChannelNameEscape33", Err.Number, Err.Description
    Resume NextDone33
End Sub

Private Sub GroupFail33(ByVal groupName As String, ByVal errNum As Long, ByVal errDesc As String)
    modTestRunner.Check groupName & "(グループ全体)", False, _
        "群の実行中に例外: " & errDesc & " (Err=" & errNum & ")"
End Sub
