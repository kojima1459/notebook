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
' 入口。群ごとにハンドラを分ける(1本のハンドラだと最初の群で落ちた時点で
' 残りが無言で消える。R33波1 W1-3 で実害が出た型)。
' ----------------------------------------------------------------------------
Public Sub RunAll34()
    On Error GoTo H01Fail34
    TestFeatureErrMessage34
H02Next34:
    On Error GoTo H02Fail34
    TestEncryptedByHeader34
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
    Resume H01Done34
End Sub
