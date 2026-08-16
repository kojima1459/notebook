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
' (5) W2-10: 「＠＝単価」の日本語表記をメールと誤検知しない。
' ----------------------------------------------------------------------------
'   日本語の業務文書は＠を「単価」の意味で日常的に使う("商品Ａ＠１００円")。
'   旧判定は＠の前後1文字しか見ていなかったので、＠の直前が英数字になる
'   書き方をした瞬間に個人情報として誤検知していた。W2-10 で
'   「＠より後ろにドットがあること」を条件に足して切り分ける。
'
'   この群は両方の呼び出し経路を撃つ:
'     ・発行側  = modPackExport.ScanChunksForPii が1チャンクごとに呼ぶ
'                 modPii.ScanText そのもの(CheckHit33/CheckNone33)
'     ・共有発信側 = modInsightGate.PiiBlocked(実物。検知しない経路は
'                 副作用ゼロなので直接叩ける)
' ----------------------------------------------------------------------------
Private Sub TestEmailUnitPrice33()
    ' --- 検知されなければならないもの(狙った検知を巻き添えにしない) -------
    CheckHit33 "W2-10_全角のメール", "ｔａｒｏ＠ｅｘａｍｐｌｅ．ｃｏｍ", "メールアドレス"
    CheckHit33 "W2-10_半角のメール", "taro@example.com", "メールアドレス"
    CheckHit33 "W2-10_＠の前にもドットがある形", "taro.yamada@example.co.jp", "メールアドレス"
    ' ドットが＠より【前】にしか無い形は通さない、を裏から確かめる番人:
    ' 下の "資料.pdf" 付き単価表記(非検知)と対になる。
    CheckHit33 "W2-10_文末がドットで終わるアドレス", _
        "Contact taro@example.com.", "メールアドレス"

    ' --- 検知されてはならないもの(今回の主目的) ---------------------------
    CheckNone33 "W2-10_単価表記_全角英字ラベル", "商品Ａ＠１００円"
    CheckNone33 "W2-10_単価表記_全角英字2字ラベルと全角カンマ", "型番ＸＹ＠２，０００円"
    '   これは W2-1 より前から誤検知していた分(＠の直前が半角数字)。
    CheckNone33 "W2-10_単価表記_半角数字ラベル(W2-1以前からの誤検知)", "商品1＠100円"
    '   ＠の前にドットがあっても、＠より後ろに無ければ通さない。
    CheckNone33 "W2-10_単価表記_＠より前にドットがある文", "資料.pdf の 商品Ａ＠１００円"

    ' --- 従来どおり非検知のまま変わらないことの番人 -----------------------
    CheckNone33 "W2-10_番人_りんご単価", "りんご＠１２０円"
    CheckNone33 "W2-10_番人_サイズ単価", "Ｓサイズ＠５００"
    CheckNone33 "W2-10_番人_時刻と場所", "打合せは１０時＠会議室Ａ"

    ' --- 共有発信側(実物の PiiBlocked)でも単価表記が通ること --------------
    modTestRunner.Check "R33-W2-10_単価表記の質問は見送られない(実物)", _
        (modInsightGate.PiiBlocked("商品Ａ＠１００円 の値上げ手順を教えてください", "gap") = False), _
        "PiiBlocked が True を返した(=共有が無言で見送られる)"
    modTestRunner.Check "R33-W2-10_半角の単価表記も見送られない(実物)", _
        (modInsightGate.PiiBlocked("商品1＠100円 の在庫はどこで見られますか", "gap") = False), _
        "PiiBlocked が True を返した"
    ' 本物のアドレスは共有発信側でも従来どおり止まる(巻き添えの確認)。
    CheckGateBlock33 "W2-10_全角のメールは関所で止まる", "ｔａｒｏ＠ｅｘａｍｐｌｅ．ｃｏｍ へ連絡"
End Sub

' ----------------------------------------------------------------------------
' (6) W3-1: Excel取込で、行内の空セルを詰めない(列位置=タブ位置を保つ)。
' ----------------------------------------------------------------------------
'   旧実装は「値のあったセル」だけを先頭から詰めて Join していたため、
'   見出し行「氏名/部署/内線」に対しデータ行「山田/(空)/1234」が
'   `山田<TAB>1234` になり、内線番号が【部署の位置】へ入っていた。
'   取込は成功扱いでログにも残らないので、検索が静かに嘘をつく。
'
'   【この群が discriminate であることの根拠】
'   旧実装ではデータ行のタブ数が見出し行より少なくなる。よって
'   「タブ数が揃うこと」と「k列目の値が Split の添字 k-1 に来ること」の
'   両方が旧実装で必ず落ちる(恒真ではない)。逆に「常に cols-1 個のタブを
'   吐く」だけの実装で通らないよう、末尾の空欄を落とす検査も対で置く。
' ----------------------------------------------------------------------------
Private Sub TestExcelRowAlign33()
    Dim arr(1 To 6, 1 To 3) As Variant
    ' 1行目=見出し(全列に値)
    arr(1, 1) = "氏名": arr(1, 2) = "部署": arr(1, 3) = "内線"
    ' 2行目=事故の再現(真ん中が空)
    arr(2, 1) = "山田": arr(2, 3) = "1234"
    ' 3行目=行頭が空
    arr(3, 2) = "総務": arr(3, 3) = "5678"
    ' 4行目=末尾が空(尻尾のタブは落とす)
    arr(4, 1) = "佐藤"
    ' 5行目=丸ごと空(行ごと捨てる)
    ' 6行目=数式ブランク("")だけ。値としては存在するので行は残る。
    arr(6, 2) = ""

    Dim has As Boolean
    Dim head As String: head = modExtractorExcel.RowTextFrom(arr, 1, 3, has)
    Dim r2 As String: r2 = modExtractorExcel.RowTextFrom(arr, 2, 3, has)
    Dim r3 As String: r3 = modExtractorExcel.RowTextFrom(arr, 3, 3, has)

    Dim hp() As String: hp = Split(head, vbTab)
    Dim p2() As String: p2 = Split(r2, vbTab)
    Dim p3() As String: p3 = Split(r3, vbTab)

    ' --- 本命: 見出しとデータのタブ数が揃う ------------------------------
    '   見出し側が3列であること自体も見る(両方とも空文字なら「揃っている」
    '   ことになってしまい、恒真化するため)。
    Dim okTabs As Boolean: okTabs = (UBound(hp) = 2)
    If okTabs Then okTabs = (UBound(p2) = 2)
    modTestRunner.Check "R33-W3-1_空セルの行でもタブ数が見出しと揃う", okTabs, _
        "見出し=[" & head & "](" & UBound(hp) & ") データ=[" & r2 & "](" & UBound(p2) & ")"

    ' --- 本命: 3列目の値が3列目の位置に来る(旧実装では2列目に来ていた) ---
    '   旧実装では UBound(p2)=1 なので、添字2の参照そのものが落ちる。
    '   よって UBound を先に見てから中身を見る(短絡しない And を避ける)。
    Dim okCol3 As Boolean: okCol3 = False
    Dim okCol2 As Boolean: okCol2 = False
    Dim okCol1 As Boolean: okCol1 = False
    If UBound(p2) = 2 Then
        okCol3 = (p2(2) = "1234")
        okCol2 = (LenB(p2(1)) = 0)
        okCol1 = (p2(0) = "山田")
    End If
    modTestRunner.Check "R33-W3-1_内線が内線の位置に来る", okCol3, "データ=[" & r2 & "]"
    modTestRunner.Check "R33-W3-1_空けた2列目は空文字のまま", okCol2, "データ=[" & r2 & "]"
    modTestRunner.Check "R33-W3-1_1列目は動かない", okCol1, "データ=[" & r2 & "]"

    ' --- 行頭が空でも後続がずり上がらない --------------------------------
    Dim okHead As Boolean: okHead = False
    If UBound(p3) = 2 Then
        okHead = (LenB(p3(0)) = 0)
        If okHead Then okHead = (p3(1) = "総務")
        If okHead Then okHead = (p3(2) = "5678")
    End If
    modTestRunner.Check "R33-W3-1_行頭が空でも列位置を保つ", okHead, "データ=[" & r3 & "]"

    ' --- 逆向きの検算: 末尾の空欄はタブごと落とす ------------------------
    '   (これが無いと「常に cols-1 個のタブを吐く」だけの実装でも上が通る)
    Dim r4 As String: r4 = modExtractorExcel.RowTextFrom(arr, 4, 3, has)
    modTestRunner.Check "R33-W3-1_末尾の空欄は尻尾のタブごと落とす", _
        (r4 = "佐藤"), "データ=[" & r4 & "]"

    ' --- 丸ごと空の行は行ごと捨てる(従来の cellCount>0 と同じ判定) -------
    Dim r5 As String: r5 = modExtractorExcel.RowTextFrom(arr, 5, 3, has)
    Dim okDrop As Boolean: okDrop = (has = False)
    If okDrop Then okDrop = (LenB(r5) = 0)
    modTestRunner.Check "R33-W3-1_丸ごと空の行は捨てる", okDrop, _
        "has=" & has & " データ=[" & r5 & "]"

    ' --- 数式ブランク("")は「値のあるセル」として行を残す ----------------
    '   IsEmpty ではないので従来も行は残っていた。ここが False になると
    '   =IF(...,"",...) だらけのシートで行数が静かに減る。
    Dim r6 As String: r6 = modExtractorExcel.RowTextFrom(arr, 6, 3, has)
    modTestRunner.Check "R33-W3-1_数式ブランクの行は残す", _
        (has = True), "has=" & has & " データ=[" & r6 & "]"

    ' --- 総当たり: 最終列に値がある限り、全行のタブ数は必ず cols-1 -------
    '   空セルの並び方(2^4=16通り)を全部撃つ。1つでも詰めたら落ちる。
    Dim g(1 To 16, 1 To 5) As Variant
    Dim bit_(1 To 4) As Long
    bit_(1) = 1: bit_(2) = 2: bit_(3) = 4: bit_(4) = 8
    Dim mask As Long, col As Long
    For mask = 0 To 15
        For col = 1 To 4
            If (mask And bit_(col)) <> 0 Then g(mask + 1, col) = "v" & col
        Next col
        g(mask + 1, 5) = "END"
    Next mask
    Dim badMask As String
    Dim line_ As String
    Dim parts_() As String
    Dim want_ As String
    For mask = 0 To 15
        line_ = modExtractorExcel.RowTextFrom(g, mask + 1, 5, has)
        parts_ = Split(line_, vbTab)
        If UBound(parts_) <> 4 Then
            If LenB(badMask) = 0 Then badMask = "mask=" & mask & " タブ数=" & UBound(parts_)
        ElseIf parts_(4) <> "END" Then
            If LenB(badMask) = 0 Then badMask = "mask=" & mask & " 最終列=[" & parts_(4) & "]"
        Else
            For col = 1 To 4
                If (mask And bit_(col)) <> 0 Then want_ = "v" & col Else want_ = ""
                If parts_(col - 1) <> want_ Then
                    If LenB(badMask) = 0 Then _
                        badMask = "mask=" & mask & " " & col & "列目=[" & parts_(col - 1) & _
                                  "] 期待=[" & want_ & "]"
                End If
            Next col
        End If
    Next mask
    modTestRunner.Check "R33-W3-1_空セルの並び16通りで列位置が保たれる", _
        (LenB(badMask) = 0), "初回の違反: " & badMask
End Sub

' ----------------------------------------------------------------------------
' (7) W3-3: CP932の資料を「読み込み成功・中身は全滅」で取り込まない。
' ----------------------------------------------------------------------------
'   ADODB.Stream の UTF-8 デコードは不正バイト列で例外を投げず U+FFFD(置換
'   文字)へ落とす。よって「読めた/読めない」だけを見ている呼び出し側からは
'   化けが見えない。判定を1本にした modUtilText.ReadTextFileAuto は、この
'   U+FFFD の比率を根拠に CP932 での読み直しへ切り替える。
'   最後の関門である modExtractor.GarbleRatio も U+FFFD を数える。
'
'   【discriminate の作り】
'   ・ReplacementRatio: U+FFFD を数えない実装(=直す前の世界)にすると
'     「CP932をUTF-8で読んだ本文」の比率が 0 になり、下の 0.02 超の検査が
'     落ちる。同時に「正しいUTF-8は 0」の検査があるので、常に 1 を返す
'     実装でも通らない。
'   ・GarbleRatio: U+FFFD の枝を外すと 0.2 超の検査が落ちる。日本語本文が
'     0 のままであることを対で見るので、常に化け扱いにする実装も通らない。
'   ADODB.Stream 自体は LO で動かせないため、ここで固定するのは
'   【判定の算数】だけ(ファイルI/Oは実機スモークの担当)。
' ----------------------------------------------------------------------------
Private Sub TestGarbleFffd33()
    Dim fffd As String: fffd = ChrW(&HFFFD&)

    ' --- CP932の日本語をUTF-8で読んだときの姿(ほぼ全部が置換文字) --------
    Dim broken As String: broken = Rep33(fffd, 40) & "ABC"
    Dim rBroken As Double: rBroken = modUtilText.ReplacementRatio(broken)
    modTestRunner.Check "R33-W3-3_置換文字だらけの本文は読み直しの線(2%)を超える", _
        (rBroken > 0.02), "ratio=" & Format$(rBroken, "0.0000")

    ' --- 正しい日本語(UTF-8)は 0 ------------------------------------------
    modTestRunner.Check "R33-W3-3_正しい日本語の本文は置換文字ゼロ", _
        (modUtilText.ReplacementRatio("本日の会議は10時から会議室Aで行います。") = 0#), _
        "ratio=" & Format$(modUtilText.ReplacementRatio("本日の会議は10時から会議室Aで行います。"), "0.0000")

    ' --- 本物のU+FFFDが数文字混ざるだけのUTF-8は読み直さない(誤爆防止) ---
    '   1000字の本文に1字だけ混ざっても 0.001 で、2%の線を超えない。
    Dim rare As String: rare = Rep33("あ", 999) & fffd
    modTestRunner.Check "R33-W3-3_置換文字が1字だけなら読み直さない", _
        (modUtilText.ReplacementRatio(rare) <= 0.02), _
        "ratio=" & Format$(modUtilText.ReplacementRatio(rare), "0.0000")

    ' --- 空白類は分母に入れない(GarbleRatio と同じ数え方) ----------------
    modTestRunner.Check "R33-W3-3_空白だけの本文は0(0除算にしない)", _
        (modUtilText.ReplacementRatio("   " & vbTab & vbLf) = 0#), _
        "ratio=" & Format$(modUtilText.ReplacementRatio("   " & vbTab & vbLf), "0.0000")
    '   空白を分母に入れていたら、下の比率は 20/40 = 0.5 まで落ちる。
    Dim padded As String: padded = Rep33(fffd & " ", 20)
    modTestRunner.Check "R33-W3-3_空白は分母に入れない", _
        (modUtilText.ReplacementRatio(padded) = 1#), _
        "ratio=" & Format$(modUtilText.ReplacementRatio(padded), "0.0000")

    ' --- 最後の関門: 化け検知が置換文字を数える --------------------------
    '   呼び出し側(modExtractorPdf.DropGarbledPages)の閾値は 0.2 超。
    modTestRunner.Check "R33-W3-3_化け検知は置換文字を化けとして数える", _
        (modExtractor.GarbleRatio(Rep33("あ", 60) & Rep33(fffd, 40)) > 0.2), _
        "ratio=" & Format$(modExtractor.GarbleRatio(Rep33("あ", 60) & Rep33(fffd, 40)), "0.0000")
    ' 対の番人: 日本語本文は従来どおり 0(置換文字を足した巻き添えが無い)。
    modTestRunner.Check "R33-W3-3_日本語本文は従来どおり化け0", _
        (modExtractor.GarbleRatio(Rep33("本日の会議は10時から", 10)) = 0#), _
        "ratio=" & Format$(modExtractor.GarbleRatio(Rep33("本日の会議は10時から", 10)), "0.0000")
End Sub

Private Function Rep33(ByVal s As String, ByVal n As Long) As String
    Dim sb As String
    Dim i As Long
    For i = 1 To n
        sb = sb & s
    Next i
    Rep33 = sb
End Function

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
G05Next33:
    On Error GoTo G05Fail33
    TestEmailUnitPrice33
G06Next33:
    On Error GoTo G06Fail33
    TestExcelRowAlign33
G07Next33:
    On Error GoTo G07Fail33
    TestGarbleFffd33
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
    Resume G05Next33
G05Fail33:
    GroupFail33 "TestEmailUnitPrice33", Err.Number, Err.Description
    Resume G06Next33
G06Fail33:
    GroupFail33 "TestExcelRowAlign33", Err.Number, Err.Description
    Resume G07Next33
G07Fail33:
    GroupFail33 "TestGarbleFffd33", Err.Number, Err.Description
    Resume NextDone33
End Sub

Private Sub GroupFail33(ByVal groupName As String, ByVal errNum As Long, ByVal errDesc As String)
    modTestRunner.Check groupName & "(グループ全体)", False, _
        "群の実行中に例外: " & errDesc & " (Err=" & errNum & ")"
End Sub
