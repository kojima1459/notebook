Attribute VB_Name = "modTestsPure13"
Option Explicit

' ============================================================================
' modTestsPure13 - R15波2(OCRの正直さ強化 / 総頁の早期確定とETA / 中断ボタン)の
'                  純ロジック回帰テスト
' ----------------------------------------------------------------------------
' なぜ新しいモジュールなのか:
'   modTestsPure11 は27,888字、modTestsPure12 は27,001字で、どちらも
'   28,000字のWARN帯の目前にある。憲章§4-6「WARN帯のモジュールに機能を
'   足さない。足す前に分割を裁定する」に従い、10 → 11 → 12 と同じ形で
'   分割先を新設した。入口は modTestsPure12.RunAll12 の末尾から呼ばれる
'   Public Sub RunAll13()。導線はその1行だけで、消すとここのテストは
'   「実行されないまま」全部PASSに見える。
'
' 固定する事実(いずれも実機第4報の根本原因に直結する):
'   ・optOcrEta.OcrPartialMemoFor(R15-4a / RC6): 全ページ描けたのに一部の頁
'     だけ読めなかったときの正直なメモ。従来この形は failN を数えるだけで
'     status="done" のまま本棚に並び、頁が【無言で】欠けていた。
'   ・optOcrEta.OcrAbortMemoFor(R15-4b / R15-6b / RC3・RC6): 打ち切りの理由
'     (利用者の中断 / AI利用の上限 / 変換エラー・時間切れ)で言うべきことは
'     まるで違う。上限に当たったなら打てる手は「時間をおく」だけで、今すぐ
'     再取込を繰り返しても同じ結果にしかならない。hasCache は次波(R15-7)で
'     True になる口で、False のあいだ「続きから再開します」と書かないこと。
'   ・optOcrEta.OcrCapMemoFor(R15-4c / RC4): 上限は【実際に効く値】で言う。
'     呼び出し元(optVision)が SafeMaxPages を掛けるようになったので、
'     ここでは「掛かった値を渡されたら正しく言う」ことだけを固定する。
'   ・optOcrEta.RemainingText / OcrPageBanner(R15-5b / RC2): 残り時間の
'     秒/分の境界と、終了目安 HH:MM の書式。現在時刻を引数で受け取る設計に
'     したので、実行時刻に左右されずに1文字単位で固定できる。
'   ・optOcrEta.BatchLabel(R15-5a/5b): 総頁が判明した取込では画像化フェーズ
'     でも「バッチ2/13 (全254頁)」と言い切る。判明していないときの見え方は
'     従来と1文字も変えない(嘘の分母を出さない=憲章§4-1)。
' ============================================================================

' ----------------------------------------------------------------------------
' R15-4a(RC6): 頁が欠けたときの正直なメモ。okN×failN の真理表。
' ----------------------------------------------------------------------------
Private Sub TestOcrPartialMemo()
    ' 失敗0なら何も言わない(全部読めた取込に注意書きを出さない)。
    modTestRunner.Check "頁欠けメモ_失敗0は空", _
        (LenB(optOcrEta.OcrPartialMemoFor(20, 0)) = 0)
    modTestRunner.Check "頁欠けメモ_失敗が負でも空", _
        (LenB(optOcrEta.OcrPartialMemoFor(20, -1)) = 0)

    ' 20頁読めて3頁失敗 = 全23頁中3頁が読めなかった。
    Dim p1 As String: p1 = optOcrEta.OcrPartialMemoFor(20, 3)
    modTestRunner.Check "頁欠けメモ_全体と失敗数を数字で言う", _
        (InStr(p1, "全23頁中3頁が読み取れませんでした") > 0), "実際=" & p1
    modTestRunner.Check "頁欠けメモ_再取込で再試行と言う", _
        (InStr(p1, "もう一度取り込むと再試行します") > 0), "実際=" & p1
    modTestRunner.Check "頁欠けメモ_上限の話をしない", _
        (InStr(p1, "上限") = 0), "実際=" & p1
    modTestRunner.Check "頁欠けメモ_続きから再開とは言わない", _
        (InStr(p1, "続きから") = 0), "実際=" & p1

    ' 1頁も読めていない(okN=0)場合でも分母は失敗数と一致する。
    ' ※この経路は実際には呼ばれない(okN=0はE0303で返す)が、数字の
    '   整合だけは崩れないことを確かめる。
    Dim p2 As String: p2 = optOcrEta.OcrPartialMemoFor(0, 5)
    modTestRunner.Check "頁欠けメモ_ok0でも分母が合う", _
        (InStr(p2, "全5頁中5頁") > 0), "実際=" & p2

    ' okNが壊れた負値でも0として扱う(嘘の分母を作らない)。
    Dim p3 As String: p3 = optOcrEta.OcrPartialMemoFor(-4, 2)
    modTestRunner.Check "頁欠けメモ_負のokは0として扱う", _
        (InStr(p3, "全2頁中2頁") > 0), "実際=" & p3
End Sub

' ----------------------------------------------------------------------------
' R15-4b / R15-6b: 打ち切りメモの真理表(理由 × キャッシュ有無)。
' ----------------------------------------------------------------------------
Private Sub TestOcrAbortMemoReasons()
    ' (1) 利用者の中断 × キャッシュ無し(本波の呼び出しは必ずこれ)。
    Dim c0 As String: c0 = optOcrEta.OcrAbortMemoFor(37, "cancel", False)
    modTestRunner.Check "中断メモ_利用者の操作と言い切る", _
        (InStr(c0, "利用者の操作で中断しました") > 0), "実際=" & c0
    modTestRunner.Check "中断メモ_保存できた頁数を言う", _
        (InStr(c0, "ここまでの37頁を保存しました") > 0), "実際=" & c0
    modTestRunner.Check "中断メモ_キャッシュ無しは最初から再試行", _
        (InStr(c0, "最初から再試行します") > 0), "実際=" & c0
    modTestRunner.Check "中断メモ_キャッシュ無しで続きからとは言わない", _
        (InStr(c0, "続きから") = 0), "実際=" & c0
    modTestRunner.Check "中断メモ_エラーのせいにしない", _
        (InStr(c0, "変換エラー") = 0), "実際=" & c0

    ' (2) 利用者の中断 × キャッシュ有り(R15-7で有効になる口)。
    Dim c1 As String: c1 = optOcrEta.OcrAbortMemoFor(37, "cancel", True)
    modTestRunner.Check "中断メモ_キャッシュ有りは続きから再開", _
        (InStr(c1, "続きから再開します") > 0), "実際=" & c1
    modTestRunner.Check "中断メモ_キャッシュ有無で文面が変わる", (c0 <> c1)

    ' (3) AI利用の上限(頁失敗が3連続)。打てる手は「時間をおく」だけ。
    Dim l0 As String: l0 = optOcrEta.OcrAbortMemoFor(12, "limit", False)
    modTestRunner.Check "上限メモ_上限の可能性を言う", _
        (InStr(l0, "AI利用の上限に達した可能性があります") > 0), "実際=" & l0
    modTestRunner.Check "上限メモ_時間をおくよう案内する", _
        (InStr(l0, "時間をおいて再度取り込んでください") > 0), "実際=" & l0
    modTestRunner.Check "上限メモ_ここまでの頁数を言う", _
        (InStr(l0, "ここまでの12頁") > 0), "実際=" & l0
    modTestRunner.Check "上限メモ_利用者のせいにしない", _
        (InStr(l0, "利用者の操作") = 0), "実際=" & l0

    ' (4) 変換エラー・時間切れ(R14からの既存文面。理由の指定が無い場合も同じ)。
    Dim e0 As String: e0 = optOcrEta.OcrAbortMemoFor(37, "error", False)
    Dim e1 As String: e1 = optOcrEta.OcrAbortMemoFor(37, "", False)
    modTestRunner.Check "エラーメモ_R14の文面を保つ", _
        (InStr(e0, "読み取りを37頁で中断しました") > 0 And _
         InStr(e0, "変換エラーまたは時間切れ") > 0), "実際=" & e0
    modTestRunner.Check "エラーメモ_理由未指定はエラー扱い", (e0 = e1), "実際=" & e1
    modTestRunner.Check "エラーメモ_3つの理由は全て別の文", _
        (c0 <> l0 And l0 <> e0 And c0 <> e0)

    ' 負の頁数でも壊れない(0として扱う)。
    modTestRunner.Check "中断メモ_負の頁数でも壊れない", _
        (InStr(optOcrEta.OcrAbortMemoFor(-9, "cancel", False), "ここまでの0頁") > 0), _
        "実際=" & optOcrEta.OcrAbortMemoFor(-9, "cancel", False)
End Sub

' ----------------------------------------------------------------------------
' R15-4c(RC4): 上限メモは【実際に効く上限】で言う。
'   optVision が optOcrCore.SafeMaxPages を掛けてから渡すようになったので、
'   ここでは「クランプ後の値を渡されたときに正しく言えるか」と、
'   クランプそのものの表(SafeMaxPages)の両方を固定する。
' ----------------------------------------------------------------------------
Private Sub TestOcrCapMemoClamped()
    ' クランプ表: config が壊れていても 1..200 の外へ出ない。
    modTestRunner.Check "上限クランプ_既定100はそのまま", _
        (optOcrCore.SafeMaxPages(100) = 100)
    modTestRunner.Check "上限クランプ_ハード上限200で頭打ち", _
        (optOcrCore.SafeMaxPages(300) = 200), _
        "実際=" & optOcrCore.SafeMaxPages(300)
    modTestRunner.Check "上限クランプ_0以下は1へ", _
        (optOcrCore.SafeMaxPages(0) = 1 And optOcrCore.SafeMaxPages(-5) = 1)
    modTestRunner.Check "上限クランプ_境界200は通す", _
        (optOcrCore.SafeMaxPages(200) = 200)
    modTestRunner.Check "上限クランプ_境界201は200へ", _
        (optOcrCore.SafeMaxPages(201) = 200)

    ' config=300 の端末で146頁だけ入った場合。クランプ後の200で言うので
    ' 「54ページは読み取れませんでした」。従来は生の300で計算して
    ' 「154ページ」という、絶対に到達できない数を出していた。
    Dim m As String
    m = optOcrEta.OcrCapMemoFor(True, 146, optOcrCore.SafeMaxPages(300))
    modTestRunner.Check "上限メモ_クランプ後の上限で言う", _
        (InStr(m, "設定上限200ページのうち先頭146ページ") > 0), "実際=" & m
    modTestRunner.Check "上限メモ_読めなかった頁数もクランプ後で計算する", _
        (InStr(m, "54ページは読み取れませんでした") > 0), "実際=" & m
    modTestRunner.Check "上限メモ_到達不能な数を出さない", _
        (InStr(m, "154ページ") = 0), "実際=" & m
End Sub

' ----------------------------------------------------------------------------
' R15-5b(RC2): 残り時間の秒/分の境界と、終了目安 HH:MM。
'   現在時刻は引数なので、ここは実行時刻に一切左右されない。
' ----------------------------------------------------------------------------
Private Sub TestRemainingText()
    Dim noTime As Date: noTime = 0

    ' 60秒未満は秒。ちょうど60秒から分。
    modTestRunner.Check "残り_59秒は秒で言う", _
        (optOcrEta.RemainingText(59, noTime) = "残り約59秒"), _
        "実際=" & optOcrEta.RemainingText(59, noTime)
    modTestRunner.Check "残り_60秒は1分", _
        (optOcrEta.RemainingText(60, noTime) = "残り約1分"), _
        "実際=" & optOcrEta.RemainingText(60, noTime)
    ' 四捨五入: 89秒→1分(1.48分) / 90秒→2分(1.5分)。
    modTestRunner.Check "残り_89秒は約1分", _
        (optOcrEta.RemainingText(89, noTime) = "残り約1分"), _
        "実際=" & optOcrEta.RemainingText(89, noTime)
    modTestRunner.Check "残り_90秒は約2分", _
        (optOcrEta.RemainingText(90, noTime) = "残り約2分"), _
        "実際=" & optOcrEta.RemainingText(90, noTime)
    ' 254頁の実機規模(4,823秒=約80分)でも秒のまま垂れ流さない。
    modTestRunner.Check "残り_80分規模は分で言う", _
        (optOcrEta.RemainingText(4823, noTime) = "残り約80分"), _
        "実際=" & optOcrEta.RemainingText(4823, noTime)
    ' 0や負でも壊れない(最低1秒扱い)。
    modTestRunner.Check "残り_0秒でも壊れない", _
        (optOcrEta.RemainingText(0, noTime) = "残り約1秒"), _
        "実際=" & optOcrEta.RemainingText(0, noTime)

    ' 時刻が分かっていないなら終了目安を出さない(嘘の時刻を出さない)。
    modTestRunner.Check "終了目安_時刻不明なら出さない", _
        (InStr(optOcrEta.RemainingText(600, noTime), "終了目安") = 0), _
        "実際=" & optOcrEta.RemainingText(600, noTime)

    ' 9時5分0秒 + 3600秒 = 10:05。2桁ゼロ埋め。
    Dim t1 As Date: t1 = TimeSerial(9, 5, 0)
    modTestRunner.Check "終了目安_1時間後を hh:mm で言う", _
        (InStr(optOcrEta.RemainingText(3600, t1), "終了目安 10:05") > 0), _
        "実際=" & optOcrEta.RemainingText(3600, t1)
    ' 0時台・1桁の分はゼロ埋めされる(0:7 のような形にしない)。
    Dim t2 As Date: t2 = TimeSerial(0, 3, 0)
    modTestRunner.Check "終了目安_0時台もゼロ埋め", _
        (InStr(optOcrEta.RemainingText(240, t2), "終了目安 00:07") > 0), _
        "実際=" & optOcrEta.RemainingText(240, t2)
    ' 秒も足して繰り上がる(23:59:30 + 60秒 = 翌日00:00)。日跨ぎで折り返す。
    Dim t3 As Date: t3 = TimeSerial(23, 59, 30)
    modTestRunner.Check "終了目安_日跨ぎで折り返す", _
        (InStr(optOcrEta.RemainingText(60, t3), "終了目安 00:00") > 0), _
        "実際=" & optOcrEta.RemainingText(60, t3)
End Sub

' ----------------------------------------------------------------------------
' R15-5b: 進捗バナー。総頁の確定/未確定 × レートの実測/過去値/無し。
' ----------------------------------------------------------------------------
Private Sub TestOcrPageBannerEta()
    Dim t As Date: t = TimeSerial(14, 0, 0)

    ' 総頁が判明していれば1頁目からETAが出る(R15-5aの成果。従来は
    ' 24頁中20頁が終わるまで何も出なかった)。
    ' 残り23頁 × 5000ms = 115秒 → 約2分、終了目安 14:01。
    Dim b1 As String: b1 = optOcrEta.OcrPageBanner(1, 24, 1, 2, 5000#, t)
    modTestRunner.Check "バナー_1頁目から分母が出る", _
        (InStr(b1, "1/24頁") > 0), "実際=" & b1
    modTestRunner.Check "バナー_バッチの分母も出る", _
        (InStr(b1, "(バッチ 1/2)") > 0), "実際=" & b1
    modTestRunner.Check "バナー_1頁目から残り時間が出る", _
        (InStr(b1, "残り約2分") > 0), "実際=" & b1
    modTestRunner.Check "バナー_1頁目から終了目安が出る", _
        (InStr(b1, "終了目安 14:01") > 0), "実際=" & b1

    ' レートが0(過去実績も無い初回端末)なら残り時間そのものを出さない。
    Dim b2 As String: b2 = optOcrEta.OcrPageBanner(1, 24, 1, 2, 0#, t)
    modTestRunner.Check "バナー_レート無しは残り時間を出さない", _
        (InStr(b2, "残り") = 0), "実際=" & b2
    modTestRunner.Check "バナー_レート無しでも分母は出す", _
        (InStr(b2, "1/24頁") > 0), "実際=" & b2

    ' 総頁が未確定なら分母もETAも出さない(嘘の数字を出さない)。
    Dim b3 As String: b3 = optOcrEta.OcrPageBanner(5, 0, 1, 0, 5000#, t)
    modTestRunner.Check "バナー_総頁未確定は頁目だけ", _
        (InStr(b3, "5頁目") > 0 And InStr(b3, "/") = 0), "実際=" & b3
    modTestRunner.Check "バナー_総頁未確定は終了目安も出さない", _
        (InStr(b3, "終了目安") = 0), "実際=" & b3

    ' 最終頁では残りを出さない(0分と言われても意味が無い)。
    modTestRunner.Check "バナー_最終頁は残りを出さない", _
        (InStr(optOcrEta.OcrPageBanner(24, 24, 2, 2, 5000#, t), "残り") = 0), _
        "実際=" & optOcrEta.OcrPageBanner(24, 24, 2, 2, 5000#, t)
End Sub

' ----------------------------------------------------------------------------
' R15-5a/5b: 画像化フェーズのバッチラベル。
' ----------------------------------------------------------------------------
Private Sub TestBatchLabel()
    ' 総頁が未確定のときの見え方は従来と1文字も変えない。
    modTestRunner.Check "バッチ表示_未確定は番号だけ", _
        (optOcrEta.BatchLabel(2, 0, 20) = "バッチ2"), _
        "実際=" & optOcrEta.BatchLabel(2, 0, 20)
    modTestRunner.Check "バッチ表示_バッチ幅が壊れていたら番号だけ", _
        (optOcrEta.BatchLabel(2, 254, 0) = "バッチ2"), _
        "実際=" & optOcrEta.BatchLabel(2, 254, 0)

    ' 総頁が判明していれば総バッチ数と総頁数まで言い切る(254頁=13バッチ)。
    Dim s As String: s = optOcrEta.BatchLabel(2, 254, 20)
    modTestRunner.Check "バッチ表示_総バッチ数を出す", _
        (InStr(s, "バッチ2/13") > 0), "実際=" & s
    modTestRunner.Check "バッチ表示_総頁数を出す", _
        (InStr(s, "(全254頁)") > 0), "実際=" & s

    ' 端数の無い割り切れる頁数(40頁=2バッチ)。
    modTestRunner.Check "バッチ表示_割り切れる頁数", _
        (optOcrEta.BatchLabel(1, 40, 20) = "バッチ1/2 (全40頁)"), _
        "実際=" & optOcrEta.BatchLabel(1, 40, 20)
End Sub

' ----------------------------------------------------------------------------
' R14-F6(移設の追随): 画像化待ち予算は【1資料あたり】。移設先でも同じ表。
' ----------------------------------------------------------------------------
Private Sub TestRemainingWaitAfterMove()
    modTestRunner.Check "予算_使っていなければ全額", _
        (optOcrEta.RemainingWaitSec(1200, 0) = 1200)
    modTestRunner.Check "予算_使ったぶんだけ減る", _
        (optOcrEta.RemainingWaitSec(1200, 500) = 700)
    modTestRunner.Check "予算_使い切ったら0", _
        (optOcrEta.RemainingWaitSec(1200, 1200) = 0)
    modTestRunner.Check "予算_残りが短くても10秒は待つ", _
        (optOcrEta.RemainingWaitSec(1200, 1197) = 10)
End Sub

Public Sub RunAll13()
    On Error GoTo PartialFail
    TestOcrPartialMemo
NextAbort13:
    On Error GoTo AbortFail
    TestOcrAbortMemoReasons
NextCap13:
    On Error GoTo CapFail
    TestOcrCapMemoClamped
NextRemain13:
    On Error GoTo RemainFail
    TestRemainingText
NextBanner13:
    On Error GoTo BannerFail
    TestOcrPageBannerEta
NextBatch13:
    On Error GoTo BatchFail
    TestBatchLabel
    TestRemainingWaitAfterMove
NextDone13:
    On Error GoTo 0
    Exit Sub

PartialFail:
    modTestRunner.Check "TestOcrPartialMemo(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextAbort13
AbortFail:
    modTestRunner.Check "TestOcrAbortMemoReasons(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextCap13
CapFail:
    modTestRunner.Check "TestOcrCapMemoClamped(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRemain13
RemainFail:
    modTestRunner.Check "TestRemainingText(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBanner13
BannerFail:
    modTestRunner.Check "TestOcrPageBannerEta(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBatch13
BatchFail:
    modTestRunner.Check "TestBatchLabel(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone13
End Sub
