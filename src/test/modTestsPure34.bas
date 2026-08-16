Attribute VB_Name = "modTestsPure34"
Option Explicit

' ============================================================================
' modTestsPure34 - R33波3(取込・データ整合)の純ロジック回帰テスト・その2。
'   modTestsPure33 が残997字になったための分割先。既存チェーン
'   (modTestsPure.RunAll→…→modTestsPure30.RunAll30)へは繋がず、
'   modTestRunner.RunAllPureTests から直接呼ばれる RunAll34 の1本が入口
'   (modTestsPure31 / 32 / 33 と同型の別枝)。
' ----------------------------------------------------------------------------
' 【このモジュールが何を守るのか】(詳細は各テストの直上コメント)
'   W3-9  modLog.FeatureErrMessage: 「機能が無効なとき」だけ固定文言を出し、
'         理由付きの #ERR は理由(またはコードの案内)をそのまま届けること。
'   W3-11 modShelfScan.EncryptedByHeader: 暗号化/非暗号化/判別不能を取り違えず、
'         .xls/.doc/PDF を暗号化と断じないこと。
'   W4-1  modConfig.ParseBoolText: 真/偽/解釈不能(=既定値)の3分岐。全角も。
'   W4-6  modShareRule.ShouldRetryProbe: 届かない共有へタイムアウトを2回払わない。
'   W5-2  modBackdrop.MemoDrop: 該当シートの1件だけ落とす(前方一致で巻き込まない)。
'   W5-1  modBackdrop.CfStartRow / CfRowsAddr / CfFormula: 条件付き書式を使用済みの
'         行に被せない算数。※実際にFormatConditionsが通るかはLOでは検証不能
'         (tools/README §4)。恒真アサートで代替しない。
'   W6-1  modShare.BoardHead* / BoardReadRows / BoardOrgBlock: 組織集計の
'         集約スナップショットの書式・鮮度・打ち切り(TestBoardSnap34)。
' ============================================================================

' ----------------------------------------------------------------------------
' W3-9: 理由付きの #ERR を「管理者が有効化すると使えます」で潰さない。
' ----------------------------------------------------------------------------
'   ExportAnswerAsDoc は失敗理由ごとに違う #ERR を返す設計で、
'   modFeatures.InvokeFeature も R6要件9 で「#ERR: を一律 FEATURE_UNAVAILABLE
'   へ潰すのをやめる」よう直されている。ところが最終利用者に見せる2箇所
'   (modUIMain / modAppAct)が接頭辞の有無だけを見て中身を捨てていたため、
'   R6要件9 の修正は最後の1ホップで無効化されていた。feature_markdown は
'   dev/prod とも既定TRUEで、ボタン自体が FeatureEnabled が真のときしか
'   描かれないので、「管理者が有効化すると使えます」が正しい状況は構造的に
'   存在しない=案内は必ず的外れになる。
'
'   【discriminate の作り】
'   ・「常に固定文言」に戻すと mock/コード/素の説明の3件が落ちる。
'   ・「常に中身をそのまま」にすると FEATURE_UNAVAILABLE の2件が落ちる。
'   両方向を対で置いているので、片側だけの実装は通らない。
' ----------------------------------------------------------------------------
Private Const FIXED34 As String = "管理者が有効化すると使えます"

' W3-11 のゴールデンで使う実測した先頭バイト(推測ではない):
'   openpyxl で作った .xlsx / zipfile で作った .docx = 50 4B 03 04 14 00 00 00
'   OLE複合ドキュメントの署名(olefile.MAGIC)      = D0 CF 11 E0 A1 B1 1A E1
Private Const ZIP34 As String = "504B030414000000"
Private Const OLE34 As String = "D0CF11E0A1B11AE1"

Private Sub TestFeatureErrMessage34()
    ' --- 本当に機能が無効なときだけ従来の固定文言 -------------------------
    CheckHas34 "無効(FEATURE_UNAVAILABLE)は固定文言", _
        "#ERR:FEATURE_UNAVAILABLE", FIXED34
    CheckHas34 "無効(小文字でも同じ)", "#ERR:feature_unavailable", FIXED34
    CheckHas34 "理由が空の#ERRも固定文言", "#ERR:", FIXED34

    ' --- 事故の再現: mockモードの案内が消えていた -------------------------
    Dim mockErr As String
    mockErr = "#ERR:mockモード(mock_llm=TRUE)ではWord連携を利用できません。" & _
              "config の mock_llm を FALSE にすると実際にWordで開けるようになります。"
    CheckHas34 "mockの理由がそのまま届く", mockErr, "mock_llm"
    CheckNot34 "mockのときは固定文言を出さない", mockErr, FIXED34

    ' --- エラーコード付きは、そのコードの案内へ委譲する -------------------
    Dim limitErr As String: limitErr = "#ERR:E0204:rate limit"
    CheckHas34 "コード付きはコードを添える", limitErr, "E0204"
    CheckNot34 "コード付きのときは固定文言を出さない", limitErr, FIXED34
    '   本文は modLog.FriendlyMessage("E0204") と同じでなければならない
    '   (同じコードに2つの文面を持たない=単一情報源)。
    modTestRunner.Check "R33-W3-9_コード付きの本文はFriendlyMessageと同一", _
        (InStr(modLog.FeatureErrMessage(limitErr), modLog.FriendlyMessage("E0204")) = 1), _
        "実際=[" & modLog.FeatureErrMessage(limitErr) & "]"

    ' --- コードに見えるだけの文字列はコード扱いしない ---------------------
    '   "E02:…" や "ABCDE:…" を FriendlyMessage へ渡すと Case Else の
    '   「予期しない問題」に化け、opt側が書いた本当の理由が消える。
    CheckHas34 "5字でも英字ならコード扱いしない", "#ERR:ABCDE:本当の理由", "本当の理由"
    CheckHas34 "桁が足りなければコード扱いしない", "#ERR:E02:本当の理由", "本当の理由"
    CheckNot34 "コードでない#ERRに予期しない問題を出さない", _
        "#ERR:ABCDE:本当の理由", "予期しない問題"

    ' --- コロンの無い素の説明はそのまま ------------------------------------
    CheckHas34 "コロン無しの理由はそのまま", _
        "#ERR:Ghostscript が見つかりません", "Ghostscript"

    ' --- #ERR: で始まらない文字列は素通し(呼び出し側は渡さない想定の保険) -
    modTestRunner.Check "R33-W3-9_#ERR以外はそのまま返す", _
        (modLog.FeatureErrMessage("ふつうの本文") = "ふつうの本文"), _
        "実際=[" & modLog.FeatureErrMessage("ふつうの本文") & "]"
End Sub

Private Sub CheckHas34(ByVal label As String, ByVal errText As String, ByVal want As String)
    Dim m As String: m = modLog.FeatureErrMessage(errText)
    modTestRunner.Check "R33-W3-9_含む: " & label, _
        (InStr(m, want) > 0), "実際=[" & m & "] 期待に含む=[" & want & "]"
End Sub

Private Sub CheckNot34(ByVal label As String, ByVal errText As String, ByVal ng As String)
    Dim m As String: m = modLog.FeatureErrMessage(errText)
    modTestRunner.Check "R33-W3-9_含まない: " & label, _
        (InStr(m, ng) = 0), "実際=[" & m & "] 含んではいけない=[" & ng & "]"
End Sub

' ----------------------------------------------------------------------------
' W3-11: 暗号化ブック/文書を、開く前に先頭バイトで見分ける。
' ----------------------------------------------------------------------------
'   暗号化ファイルへ「パスワード引数なしの Open」を投げるとモーダルが出て、
'   無人の自動同期がそこで永久に止まる。一方その再試行は 2026-07-29 の実機
'   事故(Password 引数を付けると Open 自体が失敗する端末)への保険なので
'   消せない。そこで【開く前に判別し、暗号化と分かったものだけ】を落とす。
'
'   ここで固定するのは判定の算数だけ(ADODB.Stream の実読みは LO で動かない)。
'
'   【最重要】判別できないものを「暗号化ではない」と言わないこと。
'   .xls / .doc は暗号化の有無に関わらず常に OLE、PDF は暗号化でも "%PDF" の
'   ままで、どれも区別できない。ここが "enc" に倒れると正常な .xls / .doc /
'   PDF が1本残らず取り込めなくなる(PDF経路は modExtractorPdf が
'   modExtractorWord.Extract を PDF のパスで呼ぶため直撃する)。
'
'   【discriminate の作り】
'   ・「常に enc」に壊すと、判別不能群と plain 群が落ちる。
'   ・「常に ""(判別不能)」に壊すと、enc 群が落ちる。
'   両方向を対で置いているので、片側だけの実装は通らない。
' ----------------------------------------------------------------------------
Private Sub TestEncryptedByHeader34()
    ' --- 暗号化(拡張子はOOXMLなのに中身がOLE)-----------------------------
    CheckHdr34 "暗号化xlsx", "xlsx", OLE34, "enc"
    CheckHdr34 "暗号化xlsm", "xlsm", OLE34, "enc"
    CheckHdr34 "暗号化docx", "docx", OLE34, "enc"
    CheckHdr34 "暗号化docm", "docm", OLE34, "enc"
    '   拡張子・16進の大小は揃えない書き方でも同じ答えになること。
    CheckHdr34 "大文字拡張子と小文字16進", "XLSX", "d0cf11e0a1b11ae1", "enc"

    ' --- 通常のOOXML(ZIP)は従来どおり開く経路へ ---------------------------
    CheckHdr34 "通常xlsx", "xlsx", ZIP34, "plain"
    CheckHdr34 "通常docx", "docx", ZIP34, "plain"
    '   空アーカイブ(50 4B 05 06)も ZIP。ZIPと分かれば暗号化ではない。
    CheckHdr34 "空アーカイブのZIP", "xlsx", "504B0506" & "00000000", "plain"

    ' --- 【最重要】旧形式は判別不能。従来経路へ流す -----------------------
    '   .xls / .doc は暗号化の有無に関わらず常に OLE。ここが "enc" に倒れると
    '   正常な旧形式ファイルが1本残らず取り込めなくなる。
    CheckHdr34 "xlsは判別不能(暗号化と断じない)", "xls", OLE34, ""
    CheckHdr34 "docは判別不能(暗号化と断じない)", "doc", OLE34, ""
    '   PDF は暗号化されていても先頭は "%PDF"(25 50 44 46)のまま。
    '   modExtractorPdf は PDF のパスで modExtractorWord.Extract を呼ぶので、
    '   ここが判別不能でないと PDF 取込が全滅する。
    CheckHdr34 "pdfは判別不能", "pdf", "255044462D312E37", ""
    '   平文の txt/md/csv もこの判定の対象外。
    CheckHdr34 "txtは判別不能", "txt", "E38182E38184", ""
    CheckHdr34 "拡張子なしは判別不能", "", OLE34, ""

    ' --- 壊れた入力で落ちない ----------------------------------------------
    CheckHdr34 "空ファイル(先頭が読めない)", "xlsx", "", ""
    CheckHdr34 "2バイトしか無いファイル", "xlsx", "504B", ""
    CheckHdr34 "3バイトしか無いファイル", "xlsx", "504B03", ""
    '   ちょうど4バイト読めれば判定できる。
    CheckHdr34 "ちょうど4バイトのZIP", "xlsx", "504B0304", "plain"
    CheckHdr34 "ちょうど4バイトのOLE", "xlsx", "D0CF11E0", "enc"
    '   OOXMLの拡張子だが中身がどちらでもない(PDFを改名した等)。
    CheckHdr34 "中身がZIPでもOLEでもない", "xlsx", "255044462D312E37", ""
End Sub

Private Sub CheckHdr34(ByVal label As String, ByVal ext As String, _
                       ByVal headHex As String, ByVal want As String)
    Dim r As String: r = modShelfScan.EncryptedByHeader(ext, headHex)
    modTestRunner.Check "R33-W3-11_" & label, _
        (r = want), "実際=[" & r & "] 期待=[" & want & "]"
End Sub

' ----------------------------------------------------------------------------
' W4-1: config の真偽値は、読めなければ既定値へ落ちる。
' ----------------------------------------------------------------------------
'   configは非エンジニアが直接編集する設計なので、「値を消すつもりで
'   スペースキーを押した」「日本語IMEを全角のまま ＴＲＵＥ と打った」は
'   現実に起きる。従来の GetBool はこれらを無条件 False にしており、
'   既定TRUEのキー(insight_share_enabled / sync_on_open / conv_bridge 等)が
'   無言で全部オフに倒れていた。
'
'   【discriminate の作り】3群それぞれが別の壊し方を捕まえる:
'   ・真トークン群は既定 False で呼ぶ → 「常に既定値」に壊すと落ちる。
'   ・偽トークン群は既定 True で呼ぶ  → 「真トークン以外は既定値」の
'     2分岐に縮めると落ちる(明示的な FALSE が効かなくなる事故)。
'   ・解釈不能群は既定 True / False の両方で呼ぶ → 修正前の実装
'     (「真トークンでなければ False」)に戻すと既定 True 側が全部落ちる。
'   ・全角群を消すと、全角入力だけが解釈不能群へ落ちて期待と食い違う。
' ----------------------------------------------------------------------------
Private Sub TestParseBoolText34()
    ' --- 真トークン(既定 False で呼ぶ: 既定値では説明できない True) --------
    CheckBoolT34 "true", "true"
    CheckBoolT34 "大文字TRUE", "TRUE"
    CheckBoolT34 "先頭大文字True", "True"
    CheckBoolT34 "前後に空白", "  true  "
    CheckBoolT34 "1", "1"
    CheckBoolT34 "yes", "yes"
    CheckBoolT34 "on", "ON"

    ' --- 偽トークン(既定 True で呼ぶ: 明示的なFALSEが既定に勝つこと) ------
    '   運用保守ガイドが指示する insight_share_enabled=FALSE の手入力は、
    '   ここが効かないと無効になる。
    CheckBoolF34 "false", "false"
    CheckBoolF34 "大文字FALSE", "FALSE"
    CheckBoolF34 "0", "0"
    CheckBoolF34 "no", "no"
    CheckBoolF34 "off", "OFF"
    CheckBoolF34 "前後に空白のFALSE", " FALSE "

    ' --- 全角で書かれた設定値(日本語IMEの取りこぼし) ----------------------
    CheckBoolT34 "全角ＴＲＵＥ", "ＴＲＵＥ"
    CheckBoolT34 "全角小文字ｔｒｕｅ", "ｔｒｕｅ"
    CheckBoolT34 "全角ＹＥＳ", "ＹＥＳ"
    CheckBoolT34 "全角ＯＮ", "ＯＮ"
    CheckBoolT34 "全角の１", "１"
    CheckBoolF34 "全角ＦＡＬＳＥ", "ＦＡＬＳＥ"
    CheckBoolF34 "全角小文字ｆａｌｓｅ", "ｆａｌｓｅ"
    CheckBoolF34 "全角ＮＯ", "ＮＯ"
    CheckBoolF34 "全角ＯＦＦ", "ＯＦＦ"
    CheckBoolF34 "全角の０", "０"
    '   全角スペースで囲まれていても読めること(Trim$ は U+3000 を落とさない
    '   ので、幅を均してから Trim する順序でないと解釈不能へ落ちる)。
    CheckBoolT34 "全角スペースで囲まれたＴＲＵＥ", _
        ChrW(&H3000&) & "ＴＲＵＥ" & ChrW(&H3000&)

    ' --- 解釈不能(既定値へ落ちる。両方向で確かめる) -----------------------
    '   セルが空に見える2種。Delete(=Empty)は GetBool 側の分岐で既定値に
    '   なるが、スペースキーで消したセルはここへ来る。
    CheckBoolDef34 "空文字", ""
    CheckBoolDef34 "半角スペース1個", " "
    CheckBoolDef34 "全角スペース1個", ChrW(&H3000&)
    CheckBoolDef34 "タブのみ", vbTab
    '   日本語で書いた/打ち間違えた設定値。
    CheckBoolDef34 "はい", "はい"
    CheckBoolDef34 "いいえ", "いいえ"
    CheckBoolDef34 "オン", "オン"
    CheckBoolDef34 "打ち間違いtru", "tru"
    CheckBoolDef34 "打ち間違いonn", "onn"
    CheckBoolDef34 "余計な語を伴うyes", "yes please"
    CheckBoolDef34 "トークンでない数値2", "2"
    CheckBoolDef34 "トークンでない数値-1", "-1"
    CheckBoolDef34 "trueを含むだけの文", "it is true"
End Sub

'   真トークン: 既定 False で True になること(既定値では説明できない)。
Private Sub CheckBoolT34(ByVal label As String, ByVal raw As String)
    modTestRunner.Check "R33-W4-1_真トークン: " & label, _
        (modConfig.ParseBoolText(raw, False) = True), _
        "入力=[" & raw & "] 実際=" & modConfig.ParseBoolText(raw, False) & " 期待=True"
End Sub

'   偽トークン: 既定 True で False になること(明示的なFALSEが既定に勝つ)。
Private Sub CheckBoolF34(ByVal label As String, ByVal raw As String)
    modTestRunner.Check "R33-W4-1_偽トークン: " & label, _
        (modConfig.ParseBoolText(raw, True) = False), _
        "入力=[" & raw & "] 実際=" & modConfig.ParseBoolText(raw, True) & " 期待=False"
End Sub

'   解釈不能: 既定値がそのまま返ること(True/False の両方向で確認)。
Private Sub CheckBoolDef34(ByVal label As String, ByVal raw As String)
    modTestRunner.Check "R33-W4-1_既定へ落ちる(既定True): " & label, _
        (modConfig.ParseBoolText(raw, True) = True), _
        "入力=[" & raw & "] 実際=" & modConfig.ParseBoolText(raw, True) & " 期待=True"
    modTestRunner.Check "R33-W4-1_既定へ落ちる(既定False): " & label, _
        (modConfig.ParseBoolText(raw, False) = False), _
        "入力=[" & raw & "] 実際=" & modConfig.ParseBoolText(raw, False) & " 期待=False"
End Sub

' ----------------------------------------------------------------------------
' W4-6: 到達不能な共有へ、OSのタイムアウトを2回払わない。
' ----------------------------------------------------------------------------
'   ProbeRetryPath の判定材料は【パスの形】だけで、失敗の【理由】を見ない。
'   BasePath() は必ず末尾 "\" を付けるので UNC 共有ルートは常に再試行条件を
'   満たし、ホスト停止・VPN未接続で1回目が長いタイムアウトの末に失敗しても
'   同じ死んだホストへもう一度フルのタイムアウトを払っていた(起動の無反応が
'   10〜30秒→20〜60秒)。ShouldRetryProbe が「1回目が即座に失敗したか」だけを
'   見て、構文起因の救済(R8b B10)は残したまま二重待ちを切る。
'
'   【discriminate の作り】
'   ・「常に再試行」に戻すと、遅い失敗群(2件以上)が落ちる。
'   ・「再試行しない」に倒すと、即失敗群が落ちる=B10の救済が死ぬ。
'   ・エラー番号で絞り込む実装(例: 52/76 のときだけ再試行)を足すと、
'     errNo=0/53 の即失敗ケースが落ちる。
'   閾値は境界の前後(999/1000/1001)を直値で固定しているので、定数を
'   黙って動かすとここが落ちる。
' ----------------------------------------------------------------------------
Private Sub TestShouldRetryProbe34()
    ' --- 即座に失敗した=構文起因。B10の救済を残す ------------------------
    CheckRetry34 "0msで失敗(52)", 52, 0, True
    CheckRetry34 "0msで失敗(76)", 76, 0, True
    CheckRetry34 "数msで失敗", 76, 5, True
    CheckRetry34 "境界の1つ手前(999ms)", 76, 999, True
    CheckRetry34 "境界ちょうど(1000ms)", 76, 1000, True

    ' --- 時間がかかった=到達性の問題。2回目は払わない --------------------
    CheckRetry34 "境界の1つ先(1001ms)", 76, 1001, False
    CheckRetry34 "3秒かかって失敗", 76, 3000, False
    CheckRetry34 "名前解決のタイムアウト(15秒)", 53, 15000, False
    CheckRetry34 "SMBのタイムアウト(30秒)", 76, 30000, False

    ' --- 答えを決めるのは経過時間だけで、エラー番号ではない ---------------
    '   76 は「構文で弾かれた」ときにも「ホストが死んでいる」ときにも返る。
    '   番号で絞ると、別の番号を返す端末で B10 の救済が丸ごと効かなくなる。
    CheckRetry34 "errNo=0(同名ファイル)でも即失敗なら再試行", 0, 0, True
    CheckRetry34 "errNo=53でも即失敗なら再試行", 53, 10, True
    CheckRetry34 "errNo=52でも遅ければ再試行しない", 52, 20000, False
    CheckRetry34 "errNo=0でも遅ければ再試行しない", 0, 20000, False
End Sub

Private Sub CheckRetry34(ByVal label As String, ByVal errNo As Long, _
                         ByVal ms As Long, ByVal want As Boolean)
    Dim got As Boolean: got = modShareRule.ShouldRetryProbe(errNo, ms)
    modTestRunner.Check "R33-W4-6_" & label, (got = want), _
        "errNo=" & errNo & " 経過=" & ms & "ms 実際=" & got & " 期待=" & want
End Sub

' ----------------------------------------------------------------------------
' W5-2: 背景画像の敷設に失敗したら、貼ってある画像を剥がしてメモも落とす。
' ----------------------------------------------------------------------------
'   剥がす1行(ws.SetBackgroundPicture "")は Excel 依存なので LO では実行
'   できない。ここで固定できるのは【メモ側の後始末】=MemoDrop の契約だけ。
'   なぜメモを落とす必要があるのか(このテストが守っている不変式):
'     mApplied は「どのシートにどの色を敷いたか」の唯一の記録で、Apply は
'     一致したら敷き直しを丸ごと弾く。失敗して画像を剥がしたのに記録が
'     「敷いてある」のままだと、テーマを元の色へ戻したときに弾かれて
'     そのシートだけ白いまま二度と戻らない。
'   discriminate の作り: 「常に空文字を返す」実装にすると他シートの分まで
'   消えるので2件が落ち、「常に元の文字列を返す」実装にすると3件が落ちる。
' ----------------------------------------------------------------------------
Private Sub TestMemoDrop34()
    modTestRunner.Check "R33-W5-2_該当シートの1件だけ落ちる", _
        (modBackdrop.MemoDrop("|ホーム=1|Dashboard=2|", "ホーム") = "|Dashboard=2|"), _
        "実際=" & modBackdrop.MemoDrop("|ホーム=1|Dashboard=2|", "ホーム")

    modTestRunner.Check "R33-W5-2_他シートの分は残す", _
        (modBackdrop.MemoDrop("|ホーム=1|Dashboard=2|", "Dashboard") = "|ホーム=1|"), _
        "実際=" & modBackdrop.MemoDrop("|ホーム=1|Dashboard=2|", "Dashboard")

    ' 前方一致で巻き込まないこと("ホーム" が "ホーム2" を巻き込むと、
    ' 別シートの記録まで消えて無駄な敷き直しが走る)。
    modTestRunner.Check "R33-W5-2_前方一致の別名を巻き込まない", _
        (modBackdrop.MemoDrop("|ホーム2=9|ホーム=1|", "ホーム") = "|ホーム2=9|"), _
        "実際=" & modBackdrop.MemoDrop("|ホーム2=9|ホーム=1|", "ホーム")

    modTestRunner.Check "R33-W5-2_無い名前を渡しても壊さない", _
        (modBackdrop.MemoDrop("|Dashboard=2|", "ホーム") = "|Dashboard=2|"), _
        "実際=" & modBackdrop.MemoDrop("|Dashboard=2|", "ホーム")

    ' 落とした直後は Apply の照合(InStr)が必ず外れる=次の描画で敷き直す。
    Dim m As String
    m = modBackdrop.MemoPut("", "ホーム", 2758415)
    modTestRunner.Check "R33-W5-2_落とすとMemoKeyの照合が外れる", _
        (InStr(1, modBackdrop.MemoDrop(m, "ホーム"), _
               modBackdrop.MemoKey("ホーム", 2758415), vbBinaryCompare) = 0), _
        "実際=" & modBackdrop.MemoDrop(m, "ホーム")

    ' 既存の MemoPut 契約(R32 W4-5)が MemoDrop 切り出しで変わっていないこと。
    modTestRunner.Check "R33-W5-2_MemoPutの上書き契約は不変", _
        (modBackdrop.MemoPut("|ホーム=1|", "ホーム", 7) = "|ホーム=7|"), _
        "実際=" & modBackdrop.MemoPut("|ホーム=1|", "ホーム", 7)
End Sub

' ----------------------------------------------------------------------------
' W5-1: 条件付き書式を張る範囲の算数(純関数だけ)。
' ----------------------------------------------------------------------------
'   【LOで守れること / 守れないこと】
'   守れる : どの行からどの行まで張るか(CfStartRow / CfRowsAddr)、
'            数式が常に真で目印を含むこと(CfFormula)。
'   守れない: FormatConditions.Add が通るか、UsedRange が伸びないか、
'            ルールの Interior が実際に見えるか。これは LO headless では
'            原理的に検証できない(tools/README §4)ので、実装側が
'            【実行時に自分で検算して、伸びたら剥がす】形にしてある。
'            ここで恒真アサートを作って「守っているふり」をしない。
'
'   このテストが守っている不変式は1つ:【使用済みの行に絶対に被せない】。
'   被せると本文がルールの地色で塗り潰され、カードや見出しが消える。
' ----------------------------------------------------------------------------
Private Sub TestCfRange34()
    ' 数式のゴールデン。目印が消えると ClearOwnCF が自分のルールを
    ' 見分けられなくなり、張り替えのたびにルールが積み上がる。
    modTestRunner.Check "R33-W5-1_数式は常に真かつ目印入り", _
        (modBackdrop.CfFormula() = "=ISTEXT(""MBSBG"")"), _
        "実際=" & modBackdrop.CfFormula()

    ' 境界と使用済みが一致していれば、その次の行から。
    CheckCfStart34 "境界=使用済み+1なら境界から", 31, 30, 31
    ' 解放しきれず使用済みが境界より下に残っている場合は、そちらを優先する
    ' (境界から張ると 31〜45 の実コンテンツが地色で潰れる)。
    CheckCfStart34 "使用済みが境界より下なら使用済みの次から", 31, 45, 46
    ' 行1は必ずヘッダーなので、そこへ掛かる指定は張らない。
    CheckCfStart34 "行1に掛かる指定は張らない", 1, 0, 0
    CheckCfStart34 "行2からは張ってよい", 2, 0, 2
    CheckCfStart34 "0行の異常指定は張らない", 0, 0, 0

    CheckCfAddr34 "通常は開始行から深さぶん", 31, 400, "31:430"
    CheckCfAddr34 "最終行を超えたらクランプ", 1048500, 400, "1048500:1048576"
    CheckCfAddr34 "開始行0は空(何もしない)", 0, 400, ""
    CheckCfAddr34 "行1は空(何もしない)", 1, 400, ""
    CheckCfAddr34 "深さ0は空(何もしない)", 31, 0, ""
    CheckCfAddr34 "深さ1でも1行だけ張る", 31, 1, "31:31"

    ' 深さの設計不変式: 停止線 S=B+k の k(=1画面ぶんの行数)より十分深いこと。
    ' 実機窓高の上限は約750pt、既定行高18pt換算で k は約42行。100行を下回る
    ' 設定に変えたら、深さ不足で下端に未塗り帯が残りうる。
    modTestRunner.Check "R33-W5-1_深さは1画面(約42行)より十分深い", _
        (modBackdrop.CF_DEPTH_ROWS >= 100), _
        "実際=" & modBackdrop.CF_DEPTH_ROWS

    modTestRunner.Check "R33-W5-1_最終行はExcel2007以降の1048576", _
        (modBackdrop.CF_MAX_ROW = 1048576), _
        "実際=" & modBackdrop.CF_MAX_ROW
End Sub

Private Sub CheckCfStart34(ByVal label As String, ByVal boundRow As Long, _
                           ByVal usedLast As Long, ByVal want As Long)
    Dim got As Long: got = modBackdrop.CfStartRow(boundRow, usedLast)
    modTestRunner.Check "R33-W5-1_" & label, (got = want), _
        "bound=" & boundRow & " used=" & usedLast & " 実際=" & got & " 期待=" & want
End Sub

Private Sub CheckCfAddr34(ByVal label As String, ByVal startRow As Long, _
                          ByVal depth As Long, ByVal want As String)
    Dim got As String: got = modBackdrop.CfRowsAddr(startRow, depth)
    modTestRunner.Check "R33-W5-1_" & label, (got = want), _
        "start=" & startRow & " depth=" & depth & " 実際=[" & got & "] 期待=[" & want & "]"
End Sub

' ----------------------------------------------------------------------------
' 入口。群ごとにハンドラを分ける(1本のハンドラだと最初の群で落ちた時点で
' 残りが無言で消える。R33波1 W1-3 で実害が出た型)。
' ----------------------------------------------------------------------------
' ============================================================================
' R33波5b(本棚・ギャラリー・Hub/Dash の縫い目+小物)の純ロジック回帰。
' ----------------------------------------------------------------------------
' ここに載るのは「Excelオブジェクトに触れずに答えが決まる」ものだけ。
' W5-5/6/7/8/9/10/11/13/14/16/19 は Shape・Range・共有フォルダ・直近RAGヒット
' が要るため LO では撃てない(恒真アサートで代替しない。実機確認へ回す)。
' ============================================================================

' ---- W5-12: 質問例は件数ぶんだけ描く(6個固定で巻き戻さない) ----------------
'   discriminate: 常に pageSize を返す実装は total=5 の行で落ち、
'   常に total を返す実装は total=7 の行で落ちる。両方向を対で置いてある。
Private Sub TestShownCount34()
    ChkLong34 "W5-12_プールが少ないと件数ぶん(5<6)", modStarter.ShownCountFor(5, 6), 5
    ChkLong34 "W5-12_LLM生成の上限5件でも重複しない", modStarter.ShownCountFor(3, 6), 3
    ChkLong34 "W5-12_ちょうどなら満杯", modStarter.ShownCountFor(6, 6), 6
    ChkLong34 "W5-12_多いときは1画面ぶんで打ち切る", modStarter.ShownCountFor(7, 6), 6
    ChkLong34 "W5-12_0件は0", modStarter.ShownCountFor(0, 6), 0
    ChkLong34 "W5-12_ページ幅0は0(ゼロ除算・負ループを作らない)", _
        modStarter.ShownCountFor(5, 0), 0
End Sub

' ---- W5-15: 半角カナは半角幅(5.5pt)で数える -------------------------------
'   16進リテラルに & が無いと &HFF61/&HFF9F が Integer 負値になり、条件が
'   恒偽の死に枝になる。半角カナ10文字は 55pt(=収まる)のはずが 105pt
'   (=収まらない)と見積もられ、進捗バーが不要に2行へ伸びていた。
'   discriminate: 全角10文字は【直しても直さなくても】105pt で伸びる側に
'   落ちるので、閾値そのものが効いていることを同時に固定できる。
Private Sub TestKanaWidth34()
    Dim kana As String, ascii34 As String, zen As String
    Dim i As Long
    For i = 1 To 10
        kana = kana & ChrW(&HFF71&)      ' ｱ 半角カナ(U+FF71)
        ascii34 = ascii34 & "a"
        zen = zen & ChrW(&H3042&)        ' あ 全角
    Next i
    ChkTrue34 "W5-15_半角カナ10字は1行に収まる(55pt<=80pt)", _
        (modProgressBar.BarHeightFor(kana, 80) = 30), _
        "実際=" & modProgressBar.BarHeightFor(kana, 80)
    ChkTrue34 "W5-15_ASCII10字も1行(既存の枝を壊していない)", _
        (modProgressBar.BarHeightFor(ascii34, 80) = 30), _
        "実際=" & modProgressBar.BarHeightFor(ascii34, 80)
    ChkTrue34 "W5-15_全角10字は2行へ伸びる(閾値が効いている)", _
        (modProgressBar.BarHeightFor(zen, 80) = 46), _
        "実際=" & modProgressBar.BarHeightFor(zen, 80)
End Sub

' ---- W5-17: 本棚の使用率は1つの分母・0..100クランプ -----------------------
'   Hub は chunk_limit、Dash は shelf_max_chunks を分母にしていて、取込の
'   ハード上限は後者。Hub だけが 100% を超えられた。
'   discriminate: クランプを外すと 146% の行が落ち、警告のしきい値を
'   「表示のパーセント>=80」から導くと 25000/20000 の行が落ちる。
Private Sub TestUsagePercent34()
    ChkLong34 "W5-17_ちょうど半分", modShareRule.UsagePercentOf(10250, 20500), 50
    ChkLong34 "W5-17_上限を広げた組織でも実比率", modShareRule.UsagePercentOf(24000, 40000), 60
    ChkLong34 "W5-17_100%超はクランプ", modShareRule.UsagePercentOf(30000, 20500), 100
    ChkLong34 "W5-17_分母0は0(ゼロ除算しない)", modShareRule.UsagePercentOf(100, 0), 0
    ChkLong34 "W5-17_0件は0", modShareRule.UsagePercentOf(0, 20500), 0
    ChkTrue34 "W5-17_警告線は chunk_limit の8割ちょうどで立つ", _
        modShareRule.IsBudgetTightAt(16000, 20000), ""
    ChkTrue34 "W5-17_8割に1件足りなければ立たない", _
        (modShareRule.IsBudgetTightAt(15999, 20000) = False), ""
    ChkTrue34 "W5-17_上限を広げても警告は消えない(表示%から導かない)", _
        modShareRule.IsBudgetTightAt(25000, 20000), ""
    ChkTrue34 "W5-17_しきい値0は立たない", _
        (modShareRule.IsBudgetTightAt(100, 0) = False), ""
End Sub

' ---- W5-18: 本文の起点はヘッダー帯の実高から出す -------------------------
'   サブタイトルは帯の下端+6pt に高さ18ptで置かれるので、下端は barH+24。
'   本文(KPIカード)の開始Yがそれ以上でなければ、後から描かれる不透明な
'   カードがサブタイトルを覆う。定数80固定だと barH=78 のとき 80 < 102 で
'   必ず覆っていた。3通りの帯高で不等式を固定する。
'   discriminate: BodyY0 を定数80へ戻すと barH=78/108 の2行が落ちる。
Private Sub TestDashBodyY34()
    modDashStat.SetHeaderH 48
    ChkTrue34 "W5-18_帯が最小(48)なら従来と同じ80", (modDashStat.BodyY0() = 80), _
        "実際=" & modDashStat.BodyY0()
    ChkTrue34 "W5-18_帯48でサブタイトル下端(72)を割らない", _
        (modDashStat.BodyY0() >= 48 + 24), "実際=" & modDashStat.BodyY0()

    modDashStat.SetHeaderH 78          ' ピル2段(可視幅590〜618ptの窓)
    ChkTrue34 "W5-18_帯78でサブタイトル下端(102)を割らない", _
        (modDashStat.BodyY0() >= 78 + 24), "実際=" & modDashStat.BodyY0()

    modDashStat.SetHeaderH 108         ' 3段
    ChkTrue34 "W5-18_帯108でサブタイトル下端(132)を割らない", _
        (modDashStat.BodyY0() >= 108 + 24), "実際=" & modDashStat.BodyY0()

    modDashStat.SetHeaderH 0           ' 未測定でも下限48で守る
    ChkTrue34 "W5-18_未測定なら最小帯として扱う", (modDashStat.BodyY0() = 80), _
        "実際=" & modDashStat.BodyY0()
    modDashStat.SetHeaderH 48          ' 後続テストへ状態を持ち越さない
End Sub

' ---- W5-20: 実況の言い換えで経過秒と段番号を捨てない ----------------------
'   入念(分解経路)の実況は先頭がサロゲートなので旧実装の head 抽出が発火せず、
'   本文まるごとが固定文へ差し替えられていた。
'   discriminate: 「当たったら全文置換」へ戻すと段番号・経過秒・注記の3行が
'   落ち、逆に「言い換えを一切しない」実装にすると点検の行が落ちる。
Private Sub TestHumanizeKeepsTail34()
    Dim src As String, r As String
    src = ChrW(&HD83E) & ChrW(&HDDEC) & " 入念(3論点) 6/7段: " & _
          "統合した回答を自己点検中… 経過2分13秒 ※応答なし表示でも処理中"
    r = modLive.Humanize(src)
    ChkTrue34 "W5-20_段番号が残る", (InStr(r, "6/7段") > 0), "実際=" & r
    ChkTrue34 "W5-20_論点数が残る", (InStr(r, "3論点") > 0), "実際=" & r
    ChkTrue34 "W5-20_経過秒が残る", (InStr(r, "経過2分13秒") > 0), "実際=" & r
    ChkTrue34 "W5-20_応答なし注記が残る", (InStr(r, "応答なし表示でも処理中") > 0), "実際=" & r
    ChkTrue34 "W5-20_ラベル本体は言い換わる", (InStr(r, "点検") > 0), "実際=" & r
    ChkTrue34 "W5-20_生ラベルは残らない(素通しではない)", _
        (InStr(r, "統合した回答を自己点検中") = 0), "実際=" & r

    ' 単段経路(既存の形)を壊していないこと。
    r = modLive.Humanize("(4/6) 検証中…")
    ChkTrue34 "W5-20_単段の番号は従来どおり", (Left$(r, 6) = "(4/6) "), "実際=" & r
    ChkTrue34 "W5-20_単段も言い換わる", (InStr(r, "突き合わせ") > 0), "実際=" & r
End Sub

' ---- W5-21: 回答が成立しなかったターンはモードを覚えない ------------------
'   信頼度バッジ(modUINexusDraw)も出典チップ(modPeek)も、この1式が返す
'   モード名が空かどうかを最終的な門にしている。
'   discriminate: ok を無視して modeName をそのまま返す実装(=旧コード)は
'   下2行が落ちる。常に空を返す実装は上2行が落ちる。
Private Sub TestAnsweredMode34()
    ChkTrue34 "W5-21_成立したターンはモードを返す", _
        (modMode.AnsweredMode(True, "deep") = "deep"), ""
    ChkTrue34 "W5-21_成立ターンは発信の門も開く", _
        modMode.ShouldEmitInsight(modMode.AnsweredMode(True, "deep"), 3), ""
    ChkTrue34 "W5-21_API失敗・逆質問のターンは空", _
        (modMode.AnsweredMode(False, "deep") = ""), _
        "実際=" & modMode.AnsweredMode(False, "deep")
    ChkTrue34 "W5-21_不成立ターンは発信の門も閉じる", _
        (modMode.ShouldEmitInsight(modMode.AnsweredMode(False, "deep"), 3) = False), ""
End Sub

Private Sub ChkLong34(ByVal label As String, ByVal got As Long, ByVal want As Long)
    modTestRunner.Check "R33-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkTrue34(ByVal label As String, ByVal cond As Boolean, ByVal detail As String)
    modTestRunner.Check "R33-" & label, cond, detail
End Sub

' W6-1(冒頭の一覧参照): 壊れた/古い集計を数字にしない・日付キーが変われば今日の
'   値として使わない・概算が読む側へ伝わる。全件走査へ戻さない経路自体は共有I/O。
Private Sub TestBoardSnap34()
    Dim h As String, b As String, g As String, td As Object
    h = modShare.BoardHeadText("2026-08-16 09:00:00", "20260816", 120, _
        "202608", 900, "2026", 5000, 42, True)
    b = h & vbLf & "D" & vbTab & "営業部" & vbTab & "300" & vbLf & "T" & vbTab & "u1" & vbTab & "7"
    g = modShare.BoardOrgBlock("ok", "9分", "8分", "7分", "", "s", 42, True, 500, 24)
    ChkTrue34 "W6-1 新しければok", modShare.BoardHeadStatus(h, "2026-08-15 09:00:00") = "ok", h
    ChkTrue34 "W6-1 古ければstale", modShare.BoardHeadStatus(h, "2026-08-16 09:00:01") = "stale", h
    ChkTrue34 "W6-1 印違い/列不足はbroken", modShare.BoardHeadStatus("zz" & Mid$(h, 9), "") = "broken" _
        And modShare.BoardHeadStatus(modShare.BOARD_HEAD_TAG & vbTab & "2026-08-16", "") = "broken", ""
    ChkTrue34 "W6-1 ヘッダの値", modShare.BoardHeadMin(h, "d", "20260816") = 120 And _
        modShare.BoardHeadMin(h, "m", "202608") = 900 And modShare.BoardHeadMin(h, "y", "2026") = 5000 _
        And modShare.BoardHeadField(h, 8) = "42" And modShare.BoardHeadField(h, 9) = "1", h
    ChkTrue34 "W6-1 日が変われば今日は0", modShare.BoardHeadMin(h, "d", "20260817") = 0, h
    ChkTrue34 "W6-1 自部署の行だけ", modShare.BoardReadRows(b, "営業部", td) = 300 And _
        modShare.BoardReadRows(b, "総務部", td) = 0, b
    ChkTrue34 "W6-1 集計無しは数字を出さず/okは出どころ付き", InStr(g, "9分") > 0 And _
        InStr(g, "42名ぶん") > 0 And InStr(g, "概算") > 0 And _
        InStr(modShare.BoardOrgBlock("none", "9分", "", "", "", "", 0, False, 500, 24), "9分") = 0, g
End Sub

Public Sub RunAll34()
    On Error GoTo H01Fail34
    TestFeatureErrMessage34
H02Next34:
    On Error GoTo H02Fail34
    TestEncryptedByHeader34
H03Next34:
    On Error GoTo H03Fail34
    TestParseBoolText34
H04Next34:
    On Error GoTo H04Fail34
    TestShouldRetryProbe34
H05Next34:
    On Error GoTo H05Fail34
    TestMemoDrop34
H06Next34:
    On Error GoTo H06Fail34
    TestCfRange34
H07Next34:
    On Error GoTo H07Fail34
    TestShownCount34
H08Next34:
    On Error GoTo H08Fail34
    TestKanaWidth34
H09Next34:
    On Error GoTo H09Fail34
    TestUsagePercent34
H10Next34:
    On Error GoTo H10Fail34
    TestDashBodyY34
H11Next34:
    On Error GoTo H11Fail34
    TestHumanizeKeepsTail34
H12Next34:
    On Error GoTo H12Fail34
    TestAnsweredMode34
H13Next34:
    On Error GoTo H13Fail34
    TestBoardSnap34
H01Done34:
    On Error GoTo 0
    Exit Sub

H01Fail34:
    modTestRunner.Check "TestFeatureErrMessage34(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H02Next34
H02Fail34:
    modTestRunner.Check "TestEncryptedByHeader34(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H03Next34
H03Fail34:
    modTestRunner.Check "TestParseBoolText34(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H04Next34
H04Fail34:
    modTestRunner.Check "TestShouldRetryProbe34(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H05Next34
H05Fail34:
    modTestRunner.Check "TestMemoDrop34(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H06Next34
H06Fail34:
    modTestRunner.Check "TestCfRange34(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H07Next34
H07Fail34:
    modTestRunner.Check "TestShownCount34(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H08Next34
H08Fail34:
    modTestRunner.Check "TestKanaWidth34(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H09Next34
H09Fail34:
    modTestRunner.Check "TestUsagePercent34(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H10Next34
H10Fail34:
    modTestRunner.Check "TestDashBodyY34(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H11Next34
H11Fail34:
    modTestRunner.Check "TestHumanizeKeepsTail34(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H12Next34
H12Fail34:
    modTestRunner.Check "TestAnsweredMode34(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H13Next34
H13Fail34:
    modTestRunner.Check "TestBoardSnap34(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done34
End Sub
