Attribute VB_Name = "modTestsPure10"
Option Explicit

' ============================================================================
' modTestsPure10 - R13-7c(チーム/部・共有ビーコンのteam列)の純ロジック回帰テスト
' ----------------------------------------------------------------------------
' なぜ新しいモジュールなのか:
'   modTestsPure9(26,883字)は次の追加を入れると28,000字のWARN帯に触れる。
'   憲章§4-6「WARN帯のモジュールに機能を足さない。足す前に分割を裁定する」
'   に従い、新規分割先とした。入口は modTestsPure9.RunAll9 の末尾から呼ばれる
'   Public Sub RunAll10()。modTestRunner.RunAllPureTests は modTestsPure.RunAll
'   だけを呼ぶ契約なので、ここへの導線は modTestsPure9 内の1行だけ。消すと
'   このモジュールのテストは「実行されないまま」全部PASSに見える。
'
' 固定する事実:
'   ・modP2PIo.TeamCodeOf: ユーザーIDの末尾 "_"+英大文字/数字4〜6桁だけを
'     チームコードとみなす(実機で観測された「氏名_E2T22」規約)。小文字は
'     不一致、アンダースコアが複数あるときは【最後の】"_"を基準にする
'     (末尾優先で、手前の区切りにフォールバックしない)。
'   ・modP2PIo.DeptOf: チームコードの先頭3字。空入力は空を返す。
'   ・modP2PIo.BeaconDataText/BeaconTeamField: 共有ビーコン(タブ区切り1行)の
'     組み立てとteam列の取り出し。team列が無い旧形式ファイルを新しい読み手が
'     読んでも、次の列(送信時刻)をteamと誤読しないこと(R13-7cの本旨=
'     「旧読み手は新形式の余分な列を無視する/新読み手は列数で分岐する」の
'     後半を固定する)。
'   ・optOcrCore.GsTotalPagesFromLog を読ませる【窓の大きさ】(R13-F1):
'     起動バナーと xref 修復警告で "Processing pages" が先頭300字から押し
'     出される実機相当のログを作り、300字窓では gsfail、4000字窓では image
'     と分類が入れ替わることを固定する。
'   ・modLog.FriendlyFailMsg の462分岐(R13-F11): 導入句を持たない
'     「<アプリ名>を操作できませんでした」の案内も汎用文言より優先すること。
'   ・modExtractorPdf.IsThinExtract の拡張子足切り(R13-F7): 薄い抽出の関門は
'     PDF専用で、docx/xlsxには当てない(閾値の真理表は modTestsPure9)。
'   ・modP2PIo.IsTeamCode / BeaconDataText のサニタイズ(R13 F12): 自由記述の
'     user_department(「営業部」等)をチームコードとして採用しないこと、
'     team値へのタブ/改行注入でビーコン行の列がズレないこと。
'   ・modUtilText.AppendStepBuf の書式(R13 F8): 段ごとの所要時間を
'     "expand=3x900;rerank=900" の形へ畳む規則(1質問=1行にまとめるため)。
'   ・modBitwiseOpt.ShouldPrefilterScoped(R13 F9): スコープ検索では
'     粗選別を必ず使わない(全体上位N件とスコープの積が枯れるため)。
' ============================================================================

Private Sub TestTeamCodeOf()
    ' 正常: 「氏名_E2T22」規約どおりの末尾チームコード。
    modTestRunner.Check "TeamCodeOf_正常", _
        modP2PIo.TeamCodeOf("山田太郎_E2T22") = "E2T22"

    ' サフィックス無し("_"自体が無い)。
    modTestRunner.Check "TeamCodeOf_サフィックス無し", _
        modP2PIo.TeamCodeOf("山田太郎") = ""

    ' 空入力・末尾が"_"だけ(サフィックスが空)。
    modTestRunner.Check "TeamCodeOf_空入力", modP2PIo.TeamCodeOf("") = ""
    modTestRunner.Check "TeamCodeOf_末尾アンダースコアのみ", _
        modP2PIo.TeamCodeOf("山田太郎_") = ""

    ' 小文字は不一致(規約は英大文字/数字のみ)。
    modTestRunner.Check "TeamCodeOf_小文字は不一致", _
        modP2PIo.TeamCodeOf("山田太郎_e2t22") = ""

    ' 長さ境界: 4桁(下限)・6桁(上限)は一致、3桁・7桁は不一致。
    modTestRunner.Check "TeamCodeOf_長さ境界4は一致", _
        modP2PIo.TeamCodeOf("山田_AB12") = "AB12"
    modTestRunner.Check "TeamCodeOf_長さ境界6は一致", _
        modP2PIo.TeamCodeOf("山田_ABC123") = "ABC123"
    modTestRunner.Check "TeamCodeOf_長さ境界3は不一致", _
        modP2PIo.TeamCodeOf("山田_AB1") = ""
    modTestRunner.Check "TeamCodeOf_長さ境界7は不一致", _
        modP2PIo.TeamCodeOf("山田_ABCD123") = ""

    ' アンダースコア複数は【末尾】優先。手前の区切りにフォールバックしない
    ' ことも合わせて確認する(末尾サフィックスが規約外なら全体で不一致)。
    modTestRunner.Check "TeamCodeOf_複数アンダースコアは末尾優先", _
        modP2PIo.TeamCodeOf("山田_太郎_E2T22") = "E2T22"
    modTestRunner.Check "TeamCodeOf_末尾が規約外なら手前へフォールバックしない", _
        modP2PIo.TeamCodeOf("E2T22_ab") = ""
End Sub

Private Sub TestDeptOf()
    modTestRunner.Check "DeptOf_先頭3字", modP2PIo.DeptOf("E2T22") = "E2T"
    modTestRunner.Check "DeptOf_空入力は空", modP2PIo.DeptOf("") = ""
    modTestRunner.Check "DeptOf_3字未満はそのまま", modP2PIo.DeptOf("AB") = "AB"
End Sub

' ビーコン行の組み立て→(WriteBeaconが末尾に送信時刻を足す想定)→Split→
' team列の取り出し、を実際の書式のまま往復させる。
Private Sub TestBeaconTeamRoundTrip()
    ' team列あり(新形式)。RefreshBoardが読む主要列(f(0)・f(3)・f(5))も
    ' ズレていないことを合わせて確認する。
    Dim dataText As String
    dataText = modP2PIo.BeaconDataText("alice", 3, "20260803", 15, "202608", 45, "2026", 200, "E2T22")
    Dim recNew As String: recNew = dataText & vbTab & "2026-08-03 10:00:00"
    Dim fNew() As String: fNew = Split(recNew, vbTab)

    modTestRunner.Check "Beacon_新形式_myId", fNew(0) = "alice"
    modTestRunner.Check "Beacon_新形式_今日分", CLng(Val(fNew(3))) = 15
    modTestRunner.Check "Beacon_新形式_今月分", CLng(Val(fNew(5))) = 45
    modTestRunner.Check "Beacon_新形式_team列を取り出せる", _
        modP2PIo.BeaconTeamField(fNew) = "E2T22"

    ' team列なし(旧形式。BeaconDataTextを使わず素の8列+送信時刻を組み立てる
    ' ことで、実際に流通しうる旧バージョンのファイルを模す)。
    Dim recOld As String
    recOld = "bob" & vbTab & "1" & vbTab & "20260803" & vbTab & "10" & vbTab & _
             "202608" & vbTab & "30" & vbTab & "2026" & vbTab & "120" & vbTab & _
             "2026-08-03 09:00:00"
    Dim fOld() As String: fOld = Split(recOld, vbTab)

    modTestRunner.Check "Beacon_旧形式_列数は9(idx0-8)", UBound(fOld) = 8
    modTestRunner.Check "Beacon_旧形式_送信時刻をteamと誤読しない", _
        modP2PIo.BeaconTeamField(fOld) = ""

    ' team列ありの新形式は列数が10(idx0-9)になっていること自体も固定する
    ' (この差が新旧の分岐条件そのものなので、境界がズレたら即検知できる)。
    modTestRunner.Check "Beacon_新形式_列数は10(idx0-9)", UBound(fNew) = 9
End Sub

' ----------------------------------------------------------------------------
' R13-F1: gs_out.log から総ページ数を読む窓の大きさ。
'   Ghostscriptは処理の頭に起動バナー(版名・著作権・利用条件)を約210バイト
'   出し、壊れかけのPDFではそのあとに xref 修復の警告が数行続く。従来の
'   「先頭300字」ではその時点で窓が尽き、"Processing pages 1 through 44." が
'   窓の外へ押し出されて「ページ処理の証拠なし」と読めてしまう。
'   そうなると空出力のスキャンPDFが image ではなく gsfail に分類され、
'   Word経路のゴミ本文で「登録成功」になる実機第2報RC1が再発する。
'   ここでは同じ文字列を300字窓と4000字窓で読ませ、窓の大きさだけが
'   分類を変える事実を固定する(判定関数そのものは変えていない)。
' ----------------------------------------------------------------------------
Private Sub TestGsLogHeadWindow()
    ' 実機のログを模す: 起動バナー + xref修復の警告 + Processing行。
    Dim banner As String
    banner = "GPL Ghostscript 10.02.1 (2023-11-01) " & _
             "Copyright (C) 2023 Artifex Software, Inc.  All rights reserved. " & _
             "This software is supplied under the GNU AGPLv3 and comes with " & _
             "NO WARRANTY: see the file COPYING for details. "
    Dim warn As String
    warn = "**** Error: An error occurred while reading an XREF table. " & _
           "**** The file has been damaged.  This may have been caused " & _
           "**** by a problem while converting or transfering the file. " & _
           "**** Ghostscript will attempt to recover the data. "
    Dim logText As String
    logText = banner & warn & "Processing pages 1 through 44. Page 1 Page 2 "

    ' 前置きが300字を超えていること自体を明示しておく(前提が崩れたら気付く)。
    modTestRunner.Check "GsLog窓_前置きが300字を超える実機相当", _
        Len(banner & warn) > 300, "実際=" & Len(banner & warn)

    ' 旧実装の窓(先頭300字)では総ページ数を読めない=退行の再現。
    modTestRunner.Check "GsLog窓_300字では総ページ数を読めない", _
        optOcrCore.GsTotalPagesFromLog(Left$(logText, 300)) = 0, _
        "実際=" & optOcrCore.GsTotalPagesFromLog(Left$(logText, 300))

    ' 新実装の窓(先頭4000字)なら読める。
    modTestRunner.Check "GsLog窓_4000字なら総ページ数を読める", _
        optOcrCore.GsTotalPagesFromLog(Left$(logText, 4000)) = 44, _
        "実際=" & optOcrCore.GsTotalPagesFromLog(Left$(logText, 4000))

    ' 分類への影響まで通しで固定する。証拠が見えるかどうかだけで
    ' image(→OCR)と gsfail(→Word経路)が入れ替わる。
    Dim seen300 As Boolean
    seen300 = (optOcrCore.GsTotalPagesFromLog(Left$(logText, 300)) > 0)
    Dim seen4000 As Boolean
    seen4000 = (optOcrCore.GsTotalPagesFromLog(Left$(logText, 4000)) > 0)
    modTestRunner.Check "GsLog窓_300字だとgsfailへ落ちる(退行の姿)", _
        optOcrCore.ClassifyGsTextResult(0, 0, seen300) = "gsfail", _
        "実際=" & optOcrCore.ClassifyGsTextResult(0, 0, seen300)
    modTestRunner.Check "GsLog窓_4000字ならimage(OCRへ)", _
        optOcrCore.ClassifyGsTextResult(0, 0, seen4000) = "image", _
        "実際=" & optOcrCore.ClassifyGsTextResult(0, 0, seen4000)
End Sub

' ----------------------------------------------------------------------------
' R13-F11: 462(相手のCOMサーバが居ない/セキュリティ製品に止められた)の
'   案内文は「<アプリ名>を操作できませんでした。」で始まり、他の分岐が持つ
'   導入句(「この端末では」「この環境では」)を持たない。modLog.ActionableHint
'   が導入句だけを目印にしていたため、実機で最も多いブロックの案内が汎用文言
'   (E0302=Ghostscript前提)に潰されて利用者に届いていなかった。
'   検体は modUtil.DescribeComError の実出力から組み立てる(文言を写経すると
'   本体を変えたときにテストだけが古いまま通ってしまうため)。
' ----------------------------------------------------------------------------
Private Sub TestActionableHint462()
    Dim raw As String
    raw = modUtil.DescribeComError(462, "オートメーション エラーです。", "Word")

    ' 前提: 462の案内は導入句を持たない(持ち始めたらこのテストの意味が変わる)。
    modTestRunner.Check "Hint462_案内は導入句を持たない", _
        (InStr(raw, "この端末では") = 0) And (InStr(raw, "この環境では") = 0), _
        "実際=" & raw

    ' 実機の err_log と同じく、前後に診断情報が付いた形で渡す。
    Dim detail As String
    detail = "[Word起動/開き方1] " & raw & " [localcopy=ok]"
    Dim msg As String: msg = modLog.FriendlyFailMsg("E0302", detail, "docx")

    modTestRunner.Check "Hint462_案内文を優先採用する", _
        InStr(msg, "Wordを操作できませんでした") > 0, "実際=" & msg
    modTestRunner.Check "Hint462_セキュリティ製品の可能性まで残す", _
        InStr(msg, "セキュリティ製品") > 0, "実際=" & msg
    modTestRunner.Check "Hint462_技術情報を混ぜない", _
        (InStr(msg, "(詳細:") = 0) And (InStr(msg, "localcopy") = 0) And _
        (InStr(msg, "開き方1") = 0), "実際=" & msg
    modTestRunner.Check "Hint462_コードを添える", _
        InStr(msg, "(コード: E0302)") > 0, "実際=" & msg

    ' 導入句つきの案内(429・相乗り)と同時に出ても、先に現れた方を採る。
    Dim both As String
    both = "GS: 応答なし / この端末ではExcelからWordを起動できません / " & raw
    modTestRunner.Check "Hint462_先に現れた案内を採る", _
        InStr(modLog.FriendlyFailMsg("E0302", both, "pdf"), "この端末では") > 0, _
        "実際=" & modLog.FriendlyFailMsg("E0302", both, "pdf")

    ' 「操作できませんでした」の手前にアプリ名が無い文は拾わない
    ' (診断文を巻き込んで意味の通らない案内を出さないための歯止め)。
    Dim noApp As String: noApp = "[取込] を操作できませんでした。 [localcopy=ok]"
    modTestRunner.Check "Hint462_アプリ名が無ければ拾わない", _
        InStr(modLog.FriendlyFailMsg("E0302", noApp, "pdf"), "Ghostscript") > 0, _
        "実際=" & modLog.FriendlyFailMsg("E0302", noApp, "pdf")
End Sub

' ----------------------------------------------------------------------------
' R13-F7: 薄い抽出の関門(modExtractorPdf.IsThinExtract)はPDF専用。
'   閾値そのものの真理表は modTestsPure9.TestIsThinExtract が持ち、ここは
'   「拡張子による足切り」だけを固定する(modTestsPure9 は30,000字上限まで
'   残りが少ないため置き場を分けた。憲章§4-6)。
'   docxの「ページ」もxlsxの「シート」も同じ ExtractedPage 配列に入るので、
'   拡張子を見ないとスライド資料や表計算まで partial(=本文を取り出せて
'   いない可能性があります/OCRを試します)と言われてしまう。
' ----------------------------------------------------------------------------
Private Sub TestThinExtractExt()
    ' 同じ数字でも、PDFなら薄い/PDF以外は薄くない。
    modTestRunner.Check "ThinExt_pdfは薄いと判定する", _
        modExtractorPdf.IsThinExtract("pdf", 44, 1, 30000), "Falseになった"
    modTestRunner.Check "ThinExt_docxは同じ数字でも薄くない", _
        Not modExtractorPdf.IsThinExtract("docx", 44, 1, 30000), "Trueになった"
    modTestRunner.Check "ThinExt_xlsxは同じ数字でも薄くない", _
        Not modExtractorPdf.IsThinExtract("xlsx", 40, 2, 99999), "Trueになった"
    modTestRunner.Check "ThinExt_拡張子が空なら薄くない", _
        Not modExtractorPdf.IsThinExtract("", 44, 1, 0), "Trueになった"

    ' 大文字・前後の空白で取りこぼさない(呼び出し元の渡し方に依存しない)。
    modTestRunner.Check "ThinExt_PDF大文字でも判定する", _
        modExtractorPdf.IsThinExtract(" PDF ", 44, 1, 30000), "Falseになった"
End Sub

' ----------------------------------------------------------------------------
' R13 F12: チームコードとして採用してよい文字列か(modP2PIo.IsTeamCode)と、
'   ビーコンteam列のサニタイズ(modP2PIo.BeaconDataText)。
'   config user_department は自由記述で、実機には「営業部」のような部署名が
'   そのまま入る。無検証で採用すると team 列に日本語が乗り、DeptOf がその
'   先頭3字を部コードとして扱う ―― 誰も気付かないまま部別集計だけが狂う。
'   また team 値にタブが混じるとタブ区切り行の列がまるごとズレ、他人の
'   節約時間が team として読まれる。どちらも「1文字で集計が壊れる」ので
'   境界をここで固定する。
' ----------------------------------------------------------------------------
Private Sub TestIsTeamCode()
    ' 規約どおり(3字部コード+2桁)は採用する。
    modTestRunner.Check "IsTeamCode_規約どおりは採用", _
        modP2PIo.IsTeamCode("E2T22"), "Falseになった"
    modTestRunner.Check "IsTeamCode_長さ境界4は採用", modP2PIo.IsTeamCode("AB12")
    modTestRunner.Check "IsTeamCode_長さ境界6は採用", modP2PIo.IsTeamCode("ABC123")

    ' 自由記述の部署名(F12の本命)。日本語はそもそも規約外。
    modTestRunner.Check "IsTeamCode_日本語の部署名は不採用", _
        Not modP2PIo.IsTeamCode("営業部"), "Trueになった"
    modTestRunner.Check "IsTeamCode_日本語の長い部署名は不採用", _
        Not modP2PIo.IsTeamCode("第二技術部"), "Trueになった"

    ' 英字だけ・数字だけは部署名やコード断片の混入とみなして採用しない。
    modTestRunner.Check "IsTeamCode_英字のみは不採用", Not modP2PIo.IsTeamCode("SALES")
    modTestRunner.Check "IsTeamCode_数字のみは不採用", Not modP2PIo.IsTeamCode("12345")

    ' 長さ境界の外・小文字・空白混じり・空文字。
    modTestRunner.Check "IsTeamCode_3字は不採用", Not modP2PIo.IsTeamCode("AB1")
    modTestRunner.Check "IsTeamCode_7字は不採用", Not modP2PIo.IsTeamCode("ABCD123")
    modTestRunner.Check "IsTeamCode_小文字は不採用", Not modP2PIo.IsTeamCode("e2t22")
    modTestRunner.Check "IsTeamCode_空文字は不採用", Not modP2PIo.IsTeamCode("")
    modTestRunner.Check "IsTeamCode_内部の空白は不採用", Not modP2PIo.IsTeamCode("AB 12")

    ' 前後の空白は落としてから判定する(configの入力ゆれで落とさない)。
    modTestRunner.Check "IsTeamCode_前後空白は無視して採用", _
        modP2PIo.IsTeamCode("  E2T22 "), "Falseになった"

    ' TeamCodeOf が返す値は、そのまま採用可能でなければ意味が通らない。
    modTestRunner.Check "IsTeamCode_TeamCodeOfの戻り値は採用できる", _
        modP2PIo.IsTeamCode(modP2PIo.TeamCodeOf("山田太郎_E2T22"))
End Sub

' team列へのタブ/改行注入が、行の列構成を壊さないこと。
Private Sub TestBeaconTeamSanitize()
    Dim evil As String: evil = "E2T22" & vbTab & "9999"
    Dim rec As String
    rec = modP2PIo.BeaconDataText("mallory", 0, "20260803", 0, "202608", 0, "2026", 0, evil) & _
          vbTab & "2026-08-03 11:00:00"
    Dim f() As String: f = Split(rec, vbTab)

    ' 列数は新形式のまま(タブが1本増えていたらUBoundが10になる)。
    modTestRunner.Check "Beacon_タブ注入でも列数は10(idx0-9)", UBound(f) = 9, _
        "UBound=" & UBound(f)
    ' team列にはタブが残らず、値は1本にまとまっている。
    modTestRunner.Check "Beacon_タブ注入は無害化される", _
        modP2PIo.BeaconTeamField(f) = "E2T22_9999", _
        "team=[" & modP2PIo.BeaconTeamField(f) & "]"
    ' 送信時刻の列が押し出されていない(押し出されると集計の日付がズレる)。
    modTestRunner.Check "Beacon_タブ注入でも送信時刻の位置は不動", _
        f(9) = "2026-08-03 11:00:00"

    ' 改行注入も同じく1行のまま(行が割れると次の行が別レコードに見える)。
    Dim rec2 As String
    rec2 = modP2PIo.BeaconDataText("mallory", 0, "20260803", 0, "202608", 0, "2026", 0, _
           "E2T" & vbLf & "22")
    modTestRunner.Check "Beacon_改行注入は無害化される", _
        InStr(rec2, vbLf) = 0, "改行が残った"
End Sub

' ----------------------------------------------------------------------------
' R13 F8: 段(step)ごとの所要時間バッファの書式(modUtilText.AppendStepBuf)。
'   1段=1行で usage_log へ書くと1問で10行前後になり、2,000行ローテーションが
'   約180問で一周して feedback_green 等の履歴を押し出す(前月比が静かに壊れる)。
'   1問=1行にまとめるための畳み込み規則を、ここで書式ごと固定する。
' ----------------------------------------------------------------------------
Private Sub TestStepBuf()
    Dim b As String
    b = modUtilText.AppendStepBuf("", "expand", 1200, 400)
    modTestRunner.Check "StepBuf_初出は名前=ms", b = "expand=1200", "[" & b & "]"

    b = modUtilText.AppendStepBuf(b, "rerank", 900, 400)
    modTestRunner.Check "StepBuf_2件目は;で連結", b = "expand=1200;rerank=900", "[" & b & "]"

    ' 同じ段の再登場は「回数x平均」へ畳む(1200と800 -> 2x1000)。
    b = modUtilText.AppendStepBuf(b, "expand", 800, 400)
    modTestRunner.Check "StepBuf_同名は回数x平均へ畳む", _
        b = "expand=2x1000;rerank=900", "[" & b & "]"

    ' 3回目も平均が正しく更新される(1200,800,700 -> 3x900)。
    b = modUtilText.AppendStepBuf(b, "expand", 700, 400)
    modTestRunner.Check "StepBuf_3回目の平均も正しい", _
        b = "expand=3x900;rerank=900", "[" & b & "]"

    ' 畳んでも他の段の値を壊さない。
    modTestRunner.Check "StepBuf_畳み込みで他段を壊さない", InStr(b, "rerank=900") > 0

    ' 区切り文字を含む段名は無害化する(1つの段名で行の書式が壊れないこと)。
    Dim c As String
    c = modUtilText.AppendStepBuf("", "a;b=c", 100, 400)
    modTestRunner.Check "StepBuf_区切り文字を含む段名を無害化", _
        c = "a_b_c=100", "[" & c & "]"

    ' 空の段名でも壊れない(名前なしのトークンを作らない)。
    c = modUtilText.AppendStepBuf("", "", 50, 400)
    modTestRunner.Check "StepBuf_空の段名は既定名になる", c = "step=50", "[" & c & "]"

    ' 負のmsは0として扱う(壊れた計測値で平均を汚さない)。
    c = modUtilText.AppendStepBuf("", "x", -5, 400)
    modTestRunner.Check "StepBuf_負のmsは0", c = "x=0", "[" & c & "]"

    ' 上限に達したら【新しい段名だけ】を足さない。既出の畳み込みは続く。
    Dim big As String: big = "aaa=1"
    Dim i As Long
    For i = 1 To 40
        big = modUtilText.AppendStepBuf(big, "s" & i, 100, 60)
    Next i
    modTestRunner.Check "StepBuf_上限で新しい段名を足さない", Len(big) <= 80, _
        "len=" & Len(big)
    big = modUtilText.AppendStepBuf(big, "aaa", 3, 60)
    modTestRunner.Check "StepBuf_上限でも既出段は畳み続ける", _
        InStr(big, "aaa=2x2") = 1, "[" & Left$(big, 20) & "]"
End Sub

' ----------------------------------------------------------------------------
' R13 F9: スコープ検索では粗選別(binary_rag)を使わない
'   (modBitwiseOpt.ShouldPrefilterScoped)。
'   粗選別は本棚【全体】のハミング距離上位N件を候補にする。そのあとで
'   スコープ辞書による絞り込みが走るため、本棚が大きいほど
'   (全体の上位N件) ∩ (スコープ内の行) がほぼ空になり、深掘りが
'   「候補ゼロ」でスコープ無しへ落ちる(R12 High-1 と同じ構造の欠陥)。
'   スコープ検索は辞書判定でベクトル計算の前に対象外行を捨てるので、
'   そもそも粗選別の節約が要らない。
' ----------------------------------------------------------------------------
Private Sub TestShouldPrefilterScoped()
    ' スコープなしのときは従来判定と完全に一致する(挙動を変えていない)。
    modTestRunner.Check "PrefilterScoped_非スコープは従来どおり有効", _
        modBitwiseOpt.ShouldPrefilterScoped(10000, 5000, False, True, False), "Falseになった"
    modTestRunner.Check "PrefilterScoped_非スコープ_明示ONも有効", _
        modBitwiseOpt.ShouldPrefilterScoped(10000, 5000, True, False, False)
    modTestRunner.Check "PrefilterScoped_非スコープ_小規模は無効", _
        Not modBitwiseOpt.ShouldPrefilterScoped(100, 5000, True, True, False)

    ' スコープありなら、件数や設定に関わらず必ず粗選別を使わない。
    modTestRunner.Check "PrefilterScoped_スコープありは常に無効", _
        Not modBitwiseOpt.ShouldPrefilterScoped(10000, 5000, False, True, True), "Trueになった"
    modTestRunner.Check "PrefilterScoped_スコープあり_明示ONでも無効", _
        Not modBitwiseOpt.ShouldPrefilterScoped(1000000, 5000, True, True, True)
    modTestRunner.Check "PrefilterScoped_スコープあり_小規模でも無効", _
        Not modBitwiseOpt.ShouldPrefilterScoped(10, 5000, True, True, True)

    ' 従来関数そのものは変わっていないことも合わせて固定する。
    modTestRunner.Check "PrefilterScoped_従来ShouldPrefilterは不変", _
        modBitwiseOpt.ShouldPrefilter(10000, 5000, False, True)
End Sub

Public Sub RunAll10()
    On Error GoTo TeamCodeFail
    TestTeamCodeOf
NextDeptOf:
    On Error GoTo DeptOfFail
    TestDeptOf
NextBeaconRoundTrip:
    On Error GoTo BeaconRoundTripFail
    TestBeaconTeamRoundTrip
NextGsLogWindow:
    On Error GoTo GsLogWindowFail
    TestGsLogHeadWindow
NextHint462:
    On Error GoTo Hint462Fail
    TestActionableHint462
NextThinExt:
    On Error GoTo ThinExtFail
    TestThinExtractExt
NextIsTeamCode:
    On Error GoTo IsTeamCodeFail
    TestIsTeamCode
NextBeaconSanitize:
    On Error GoTo BeaconSanitizeFail
    TestBeaconTeamSanitize
NextStepBuf:
    On Error GoTo StepBufFail
    TestStepBuf
NextPrefilterScoped:
    On Error GoTo PrefilterScopedFail
    TestShouldPrefilterScoped
NextRun11:
    ' 2026-08-03(R14-3): 容量のための分割先(modTestsPure11)。ここが唯一の
    ' 導線で、消すとコピー失敗判定・OCRバッチ境界・上限メモのテストが
    ' 「実行されないまま」全部PASSに見える。
    On Error GoTo Run11Fail
    modTestsPure11.RunAll11
NextDone10:
    On Error GoTo 0
    Exit Sub

TeamCodeFail:
    modTestRunner.Check "TestTeamCodeOf(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDeptOf
DeptOfFail:
    modTestRunner.Check "TestDeptOf(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBeaconRoundTrip
BeaconRoundTripFail:
    modTestRunner.Check "TestBeaconTeamRoundTrip(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGsLogWindow
GsLogWindowFail:
    modTestRunner.Check "TestGsLogHeadWindow(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextHint462
Hint462Fail:
    modTestRunner.Check "TestActionableHint462(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextThinExt
ThinExtFail:
    modTestRunner.Check "TestThinExtractExt(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextIsTeamCode
IsTeamCodeFail:
    modTestRunner.Check "TestIsTeamCode(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBeaconSanitize
BeaconSanitizeFail:
    modTestRunner.Check "TestBeaconTeamSanitize(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextStepBuf
StepBufFail:
    modTestRunner.Check "TestStepBuf(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextPrefilterScoped
PrefilterScopedFail:
    modTestRunner.Check "TestShouldPrefilterScoped(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRun11
Run11Fail:
    modTestRunner.Check "modTestsPure11.RunAll11(モジュール全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone10
End Sub
