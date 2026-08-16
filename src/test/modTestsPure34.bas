Attribute VB_Name = "modTestsPure34"
Option Explicit

' ============================================================================
' modTestsPure34 - R33波3(取込・データ整合)の純ロジック回帰テスト・その2。
'   modTestsPure33 が残997字になったための分割先。既存チェーン
'   (modTestsPure.RunAll→…→modTestsPure30.RunAll30)へは繋がず、
'   modTestRunner.RunAllPureTests から直接呼ばれる RunAll34 の1本が入口
'   (modTestsPure31 / 32 / 33 と同型の別枝)。
' ----------------------------------------------------------------------------
' 【このモジュールが何を守るのか】
'   W3-9: opt機能が返す "#ERR:…" の理由を、UIが握り潰さないこと。
'     modLog.FeatureErrMessage が「本当に機能が無効なとき」だけ固定文言を出し、
'     それ以外は理由(またはエラーコードに対応する案内)を返す。
'   W3-11: 暗号化された Office 文書を、開く前に先頭バイトで見分けること。
'     modShelfScan.EncryptedByHeader が「暗号化」「非暗号化」「判別不能」を
'     取り違えないこと。とくに .xls/.doc/PDF を暗号化と断じないこと。
'   W4-1: config の真偽値が、読めないときに既定値へ落ちること。
'     modConfig.ParseBoolText が「真トークン/偽トークン/解釈不能」を
'     取り違えないこと。全角で書かれた TRUE/FALSE も取りこぼさないこと。
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
' 入口。群ごとにハンドラを分ける(1本のハンドラだと最初の群で落ちた時点で
' 残りが無言で消える。R33波1 W1-3 で実害が出た型)。
' ----------------------------------------------------------------------------
Public Sub RunAll34()
    On Error GoTo H01Fail34
    TestFeatureErrMessage34
H02Next34:
    On Error GoTo H02Fail34
    TestEncryptedByHeader34
H03Next34:
    On Error GoTo H03Fail34
    TestParseBoolText34
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
    Resume H01Done34
End Sub
