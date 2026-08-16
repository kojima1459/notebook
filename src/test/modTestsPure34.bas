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
' 入口。群ごとにハンドラを分ける(1本のハンドラだと最初の群で落ちた時点で
' 残りが無言で消える。R33波1 W1-3 で実害が出た型)。
' ----------------------------------------------------------------------------
Public Sub RunAll34()
    On Error GoTo H01Fail34
    TestFeatureErrMessage34
H01Done34:
    On Error GoTo 0
    Exit Sub

H01Fail34:
    modTestRunner.Check "TestFeatureErrMessage34(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done34
End Sub
