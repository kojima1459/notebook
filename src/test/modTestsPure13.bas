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
'   ・modUtilText.IngestChunksDetail(R15-8a/RC1): 画面の「127」とusage_logの
'     「125」の食い違いの正体(生成chunkN vs 重複排除後accepted)をログへ
'     残す書式。重複0のときは従来どおり"chunks=N"のみ(ログ互換優先)。
'   ・modUIShelf.ParseStats(R15-8b/RC1): modVaultGalleryの単純Split重複実装を
'     置換した共用パース。正常系に加え、error_note自体に"|"が混じる異常系
'     でも ingested_at/chunk_count がズレないことを固定する。
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
    ' クランプ表: config が壊れていても 1..300 の外へ出ない
    ' (R15-7a でハード上限を 200 → 300 へ引き上げた。254頁の資料を分割
    '  せずに取り込むため。既定 vision_pdf_max_pages も 100 → 300)。
    modTestRunner.Check "上限クランプ_既定300はそのまま", _
        (optOcrCore.SafeMaxPages(300) = 300)
    modTestRunner.Check "上限クランプ_ハード上限300で頭打ち", _
        (optOcrCore.SafeMaxPages(500) = 300), _
        "実際=" & optOcrCore.SafeMaxPages(500)
    modTestRunner.Check "上限クランプ_0以下は1へ", _
        (optOcrCore.SafeMaxPages(0) = 1 And optOcrCore.SafeMaxPages(-5) = 1)
    modTestRunner.Check "上限クランプ_境界300は通す", _
        (optOcrCore.SafeMaxPages(300) = 300)
    modTestRunner.Check "上限クランプ_境界301は300へ", _
        (optOcrCore.SafeMaxPages(301) = 300)
    modTestRunner.Check "上限クランプ_254頁はそのまま通る", _
        (optOcrCore.SafeMaxPages(254) = 254)

    ' config=500 の端末で246頁だけ入った場合。クランプ後の300で言うので
    ' 「54ページは読み取れませんでした」。従来は生の500で計算して
    ' 「254ページ」という、絶対に到達できない数を出していた。
    Dim m As String
    m = optOcrEta.OcrCapMemoFor(True, 246, optOcrCore.SafeMaxPages(500))
    modTestRunner.Check "上限メモ_クランプ後の上限で言う", _
        (InStr(m, "設定上限300ページのうち先頭246ページ") > 0), "実際=" & m
    modTestRunner.Check "上限メモ_読めなかった頁数もクランプ後で計算する", _
        (InStr(m, "54ページは読み取れませんでした") > 0), "実際=" & m
    modTestRunner.Check "上限メモ_到達不能な数を出さない", _
        (InStr(m, "254ページ") = 0), "実際=" & m
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


' ----------------------------------------------------------------------------
' R15-7d(RC7): 頁キャッシュの鍵。ここが緩いと【中身が差し替わった資料に
'   古い本文が混ざる】(規程・料金表が同名で更新されるのがこの製品の主対象
'   なので、静かな取り違えは最悪の壊れ方になる)。逆に厳しすぎると毎回
'   全頁を読み直すだけで、失うのは速さだけ。安全側は「厳しい」方。
' ----------------------------------------------------------------------------
Private Sub TestOcrCacheKey()
    Dim baseKey As String
    baseKey = optOcrCache.CacheKeyFor("C:\docs\a.pdf", 12345, "2026-08-04 10:00:00", 7)

    modTestRunner.Check "鍵_同じ資料の同じ頁は同じ鍵", _
        (baseKey = optOcrCache.CacheKeyFor("C:\docs\a.pdf", 12345, "2026-08-04 10:00:00", 7))
    modTestRunner.Check "鍵_頁番号は末尾に |p<頁> で付く", _
        (Right$(baseKey, 3) = "|p7"), "実際=" & baseKey
    modTestRunner.Check "鍵_頁が違えば別の鍵", _
        (baseKey <> optOcrCache.CacheKeyFor("C:\docs\a.pdf", 12345, "2026-08-04 10:00:00", 8))

    ' 差し替え検知の3条件(どれか1つでも違えば古い本文は使わない)。
    modTestRunner.Check "鍵_パスが違えば別の鍵", _
        (baseKey <> optOcrCache.CacheKeyFor("C:\docs\b.pdf", 12345, "2026-08-04 10:00:00", 7))
    modTestRunner.Check "鍵_サイズが違えば別の鍵(差し替え検知)", _
        (baseKey <> optOcrCache.CacheKeyFor("C:\docs\a.pdf", 12346, "2026-08-04 10:00:00", 7))
    modTestRunner.Check "鍵_更新日時が違えば別の鍵(差し替え検知)", _
        (baseKey <> optOcrCache.CacheKeyFor("C:\docs\a.pdf", 12345, "2026-08-04 10:00:01", 7))

    ' 頁より前の部分は資料1件を指す。頁の鍵はそこへ "|p" を足しただけ。
    Dim pre As String
    pre = optOcrCache.DocPrefixFor("C:\docs\a.pdf", 12345, "2026-08-04 10:00:00")
    modTestRunner.Check "鍵_資料の鍵に頁を足したものが行の鍵", _
        (baseKey = pre & "|p7"), "実際=" & baseKey

    ' セルの上限は32,767字。控えは32,000字で切る(切っても取込は壊れない)。
    Dim long1 As String: long1 = String$(33000, "a")
    modTestRunner.Check "控え本文_32000字で切る", _
        (Len(optOcrCache.CacheTextFor(long1)) = 32000), _
        "実際=" & Len(optOcrCache.CacheTextFor(long1))
    modTestRunner.Check "控え本文_短い本文はそのまま", _
        (optOcrCache.CacheTextFor("abc") = "abc")
End Sub

' ----------------------------------------------------------------------------
' R15-7b(RC4/RC7): 何時間もかかる資料は始める前に一度だけ聞く。
'   見積もりは【これから読む頁数】×レート。前回の続きから復元できる頁は
'   待ち時間0なので数えない(全頁揃っていれば確認そのものが出ない)。
' ----------------------------------------------------------------------------
Private Sub TestOcrConfirmEstimate()
    ' 実測レートが無い端末は 25秒/頁 と仮定する(実機第4報の観測値)。
    modTestRunner.Check "見積_254頁は約106分(実測なし=25秒/頁)", _
        (optOcrEta.OcrEstMinutes(254, 0) = 106), _
        "実際=" & optOcrEta.OcrEstMinutes(254, 0)
    modTestRunner.Check "見積_読む頁が0なら0分(全頁が前回の続きにある)", _
        (optOcrEta.OcrEstMinutes(0, 0) = 0)
    modTestRunner.Check "見積_1頁でも最低1分と言う(0分と言わない)", _
        (optOcrEta.OcrEstMinutes(1, 0) = 1)
    modTestRunner.Check "見積_実測があればそちらを使う(10秒/頁×120頁=20分)", _
        (optOcrEta.OcrEstMinutes(120, 10000#) = 20), _
        "実際=" & optOcrEta.OcrEstMinutes(120, 10000#)

    ' ocr_confirm_min_minutes(既定15)の境界。ちょうどは【聞く】。
    modTestRunner.Check "確認_見積14分は聞かない(15分未満)", _
        (LenB(optOcrEta.OcrConfirmAskFor(30, 14, 15)) = 0)
    modTestRunner.Check "確認_見積ちょうど15分は聞く", _
        (LenB(optOcrEta.OcrConfirmAskFor(36, 15, 15)) > 0)
    modTestRunner.Check "確認_総頁が不明(0)なら聞かない", _
        (LenB(optOcrEta.OcrConfirmAskFor(0, 106, 15)) = 0)
    modTestRunner.Check "確認_しきい値0以下は確認しない(0=無効の慣習)", _
        (LenB(optOcrEta.OcrConfirmAskFor(254, 106, 0)) = 0)

    Dim ask As String: ask = optOcrEta.OcrConfirmAskFor(254, 106, 15)
    modTestRunner.Check "確認文_総頁を名乗る", _
        (InStr(ask, "全254頁") > 0), "実際=" & ask
    modTestRunner.Check "確認文_推定時間を分で言う", _
        (InStr(ask, "推定約106分") > 0), "実際=" & ask
    modTestRunner.Check "確認文_途中で止められることを言う", _
        (InStr(ask, "中断") > 0), "実際=" & ask
    modTestRunner.Check "確認文_続きから再開できることを言う", _
        (InStr(ask, "続きから再開") > 0), "実際=" & ask
    modTestRunner.Check "確認文_最後に取り込むかを聞く", _
        (InStr(ask, "取り込みますか") > 0), "実際=" & ask

    ' 「いいえ」を選んだ資料は、失敗ではなく【見送り】として理由を残す。
    Dim no1 As String: no1 = optOcrEta.OcrDeclineMemoFor(106)
    modTestRunner.Check "見送りメモ_推定時間を理由として言う", _
        (InStr(no1, "推定約106分のため取込を見送りました") > 0), "実際=" & no1
    modTestRunner.Check "見送りメモ_次の一手を必ず言う", _
        (InStr(no1, "再度取り込むと実行します") > 0), "実際=" & no1
End Sub

' ----------------------------------------------------------------------------
' R15-7c(RC10): 画像化待ちの絶対上限を頁数へ連動させる。
'   既定1200秒のまま254頁を13バッチ描くと、後半は必ず「残り10秒」になり、
'   描けているのに時間切れで partial になる。
' ----------------------------------------------------------------------------
Private Sub TestGsBudgetByPages()
    modTestRunner.Check "予算_総頁が不明ならconfigの値のまま", _
        (optOcrEta.GsBudgetSec(1200, 0) = 1200)
    modTestRunner.Check "予算_少ない頁数ではconfigが勝つ(100頁=800秒)", _
        (optOcrEta.GsBudgetSec(1200, 100) = 1200), _
        "実際=" & optOcrEta.GsBudgetSec(1200, 100)
    modTestRunner.Check "予算_境界150頁はちょうど1200秒(configと同値)", _
        (optOcrEta.GsBudgetSec(1200, 150) = 1200)
    modTestRunner.Check "予算_151頁から頁数が勝つ", _
        (optOcrEta.GsBudgetSec(1200, 151) = 1208), _
        "実際=" & optOcrEta.GsBudgetSec(1200, 151)
    modTestRunner.Check "予算_254頁は2032秒を確保する", _
        (optOcrEta.GsBudgetSec(1200, 254) = 2032), _
        "実際=" & optOcrEta.GsBudgetSec(1200, 254)
    modTestRunner.Check "予算_configを大きくしてある端末はそちらが勝つ", _
        (optOcrEta.GsBudgetSec(9000, 254) = 9000)
    modTestRunner.Check "予算_壊れた負のconfigでも負を返さない", _
        (optOcrEta.GsBudgetSec(-5, 0) = 0)
End Sub

' ----------------------------------------------------------------------------
' R15-7d: 復元したことを黙らない/キャッシュがあるときだけ「続きから」と書く。
' ----------------------------------------------------------------------------
Private Sub TestOcrResumeAndCacheMemo()
    modTestRunner.Check "再開メモ_復元0なら1文字も足さない", _
        (optOcrEta.OcrResumeMemo("全23頁中3頁が読み取れませんでした", 0) = _
         "全23頁中3頁が読み取れませんでした")

    Dim r1 As String
    r1 = optOcrEta.OcrResumeMemo("全23頁中3頁が読み取れませんでした", 20)
    modTestRunner.Check "再開メモ_復元があれば冒頭に付ける", _
        (Left$(r1, 14) = "前回の続きから再開しました。"), "実際=" & r1
    modTestRunner.Check "再開メモ_元のメモは消さない", _
        (InStr(r1, "全23頁中3頁が読み取れませんでした") > 0), "実際=" & r1
    modTestRunner.Check "再開メモ_メモが空でも事実だけは残す", _
        (optOcrEta.OcrResumeMemo("", 5) = "前回の続きから再開しました。")

    ' hasCache=True の側(波2では常にFalseだった口が、本波で初めてTrueになる)。
    Dim c1 As String: c1 = optOcrEta.OcrAbortMemoFor(120, "cancel", True)
    modTestRunner.Check "中断メモ_控えがあるなら続きから再開すると言う", _
        (InStr(c1, "もう一度取り込むと続きから再開します") > 0), "実際=" & c1
    modTestRunner.Check "中断メモ_控えがあるとき最初からとは言わない", _
        (InStr(c1, "最初から") = 0), "実際=" & c1
    Dim c2 As String: c2 = optOcrEta.OcrAbortMemoFor(120, "error", True)
    modTestRunner.Check "中断メモ_時間切れでも控えがあれば続きから", _
        (InStr(c2, "続きから再開します") > 0), "実際=" & c2
    ' 控えが無い(書込みに失敗した)ときに「続きから」と書いたら嘘になる。
    Dim c3 As String: c3 = optOcrEta.OcrAbortMemoFor(120, "cancel", False)
    modTestRunner.Check "中断メモ_控えが無ければ最初から再試行と言う", _
        (InStr(c3, "もう一度取り込むと最初から再試行します") > 0), "実際=" & c3
End Sub

' ----------------------------------------------------------------------------
' R15-8a(RC1): 取込内訳の可視化。dup(=gen-accepted)が0のときは既存ログとの
'   互換を優先して"chunks=N"のみ。差があるときだけ内訳を添える。
' ----------------------------------------------------------------------------
Private Sub TestIngestChunksDetail()
    modTestRunner.Check "取込内訳_重複0は従来どおりchunks=Nのみ", _
        (modUtilText.IngestChunksDetail(125, 125) = "chunks=125"), _
        "実際=" & modUtilText.IngestChunksDetail(125, 125)
    ' 生成127・保存125=重複2件。画面の「127」と記録の「125」の食い違いの
    ' 正体をログへ残す。
    modTestRunner.Check "取込内訳_重複ありは内訳を添える", _
        (modUtilText.IngestChunksDetail(125, 127) = "chunks=125 (gen=127 dup=2)"), _
        "実際=" & modUtilText.IngestChunksDetail(125, 127)
    modTestRunner.Check "取込内訳_全滅でも壊れない", _
        (modUtilText.IngestChunksDetail(0, 0) = "chunks=0")
    ' acceptedがgenを超えることは実際には起こらないが、負dupという
    ' 意味の無い表示を出さないことだけ確かめる。
    modTestRunner.Check "取込内訳_acceptedがgenを超えても負dupを出さない", _
        (modUtilText.IngestChunksDetail(5, 3) = "chunks=5"), _
        "実際=" & modUtilText.IngestChunksDetail(5, 3)
End Sub

' ----------------------------------------------------------------------------
' R15-8b(RC1): modVaultGallery.DrawOneCardが使っていた単純Split(statLine,"|")
'   の重複実装を、modUIShelf.ParseStats(頑健パース)へ共用した。正常系に加え、
'   error_note自体に"|"が混じる異常系でも、両モジュールが実際に使う
'   ingested_at/chunk_count がズレないことをここで固定する(挙動不変の証拠)。
' ----------------------------------------------------------------------------
Private Sub TestParseStatsShared()
    Dim status As String, ingestedAt As String, chunkCount As String
    Dim errorNote As String, origin As String

    ' 正常系: modShelf.SourceListの契約どおり5要素(status|ingested_at|
    ' chunk_count|error_note|origin)。
    modUIShelf.ParseStats "done|2026-08-04 09:00:00|125||self", _
        status, ingestedAt, chunkCount, errorNote, origin
    modTestRunner.Check "統計パース_正常系status", (status = "done")
    modTestRunner.Check "統計パース_正常系ingestedAt", _
        (ingestedAt = "2026-08-04 09:00:00"), "実際=" & ingestedAt
    modTestRunner.Check "統計パース_正常系chunkCount", (chunkCount = "125")
    modTestRunner.Check "統計パース_正常系errorNoteは空", (LenB(errorNote) = 0)
    modTestRunner.Check "統計パース_正常系origin", (origin = "self")

    ' 異常系: error_note自体に「|」が混じる(単純Splitでは中身がズレる形)。
    modUIShelf.ParseStats _
        "partial|2026-08-04 09:00:00|125|上限に達しました|続けて再試行してください|self", _
        status, ingestedAt, chunkCount, errorNote, origin
    modTestRunner.Check "統計パース_異常系でもingestedAtはズレない", _
        (ingestedAt = "2026-08-04 09:00:00"), "実際=" & ingestedAt
    modTestRunner.Check "統計パース_異常系でもchunkCountはズレない", _
        (chunkCount = "125"), "実際=" & chunkCount
    modTestRunner.Check "統計パース_異常系errorNoteはパイプごと復元", _
        (errorNote = "上限に達しました|続けて再試行してください"), "実際=" & errorNote
    modTestRunner.Check "統計パース_異常系originは末尾のまま", _
        (origin = "self"), "実際=" & origin
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
NextCache13:
    On Error GoTo CacheFail
    TestOcrCacheKey
NextConfirm13:
    On Error GoTo ConfirmFail
    TestOcrConfirmEstimate
NextBudget13:
    On Error GoTo BudgetFail
    TestGsBudgetByPages
    TestOcrResumeAndCacheMemo
NextR158_13:
    On Error GoTo R158Fail13
    TestIngestChunksDetail
    TestParseStatsShared
NextFixA13:
    On Error GoTo FixAFail13
    modTestsPure14.RunAll14
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
    Resume NextCache13
CacheFail:
    modTestRunner.Check "TestOcrCacheKey(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextConfirm13
ConfirmFail:
    modTestRunner.Check "TestOcrConfirmEstimate(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBudget13
BudgetFail:
    modTestRunner.Check "TestGsBudgetByPages(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextR158_13
R158Fail13:
    modTestRunner.Check "TestIngestChunksDetail/TestParseStatsShared(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextFixA13
FixAFail13:
    ' R15-FixA: 続きは modTestsPure14(本モジュールは29,989字まで埋まった)。
    modTestRunner.Check "modTestsPure14.RunAll14(モジュール全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone13
End Sub
