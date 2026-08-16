Attribute VB_Name = "modUtilText"
Option Explicit

' ============================================================================
' modUtilText - テキストファイル入出力と、複数モジュールに散っていた
'   「同じ算数」を1本にまとめるための共通部品置き場(基盤層)。
' ----------------------------------------------------------------------------
' 2026-07-31(R11-F1/F2): 憲章§4-5「同型の問題は共通部品で一度だけ解決する。
'   2箇所で違う答えを出す実装を残さない」に基づく統合先。modUtil は27,822字で
'   余地が少ないため、新しい共通部品はこちらへ置く。
'
' 置くもの / 置かないもの:
'   ・置く: 2箇所以上で同じ答えを出す必要がある処理。ADODB.Stream による
'     UTF-8の読み書き、経過ミリ秒の算出(Timerの日跨ぎ)、1件あたり所要時間の
'     移動平均、GS txtwrite 出力のページ分割添字。
'   ・置かない: 画面(Shape/Range)を触るもの、業務ロジックの判断。
'     基盤層なので上位層(ingest/qa/pack/stats/ui)を参照してはならない。
' ============================================================================

' 1日のミリ秒。VBAのTimerは0時に0へ戻るため、経過が負になったらこれを足す。
Private Const MS_PER_DAY As Double = 86400000#

' 「見送りました」のメモの先頭句(2026-08-04 R15-FixB FB-5)。実体の説明と
' 判定は末尾の IsDeclineNote を参照(モジュールレベル宣言は実機VBAの制約で
' プロシージャ定義より前に置かなければならないため、宣言だけがここにある)。
Public Const DECLINE_MEMO_HEAD As String = "取込を見送りました"

' 文字コード自動判定(ReadTextFileAuto)のしきい値。実体の説明はその関数の
' 直前を参照(モジュールレベル宣言は実機VBAの制約でプロシージャ定義より前に
' 置かなければならないため、宣言だけがここにある)。
'   FFFD_LIMIT: 置換文字(U+FFFD)の比率がこれを超えたら「UTF-8として読めて
'     いない」とみなして CP932 で読み直す。CP932 の日本語を UTF-8 で読むと
'     本文のほぼ全部が U+FFFD になる一方、正しい UTF-8 に本物の U+FFFD が
'     混ざるのは数文字なので、2% の線がこの2つを余裕をもって分ける。
'   FFFD_SAMPLE: 比率を測る先頭文字数。文字コードの取り違えは【ファイル全体
'     の性質】なので先頭だけ見れば足りる(上限50MBの本文を1文字ずつ回すと
'     実用にならない)。
Private Const FFFD_LIMIT As Double = 0.02
Private Const FFFD_SAMPLE As Long = 20000

' ----------------------------------------------------------------------------
' ReadTextFileUtf8 - UTF-8テキストファイルを読む(ADODB.Stream)。
'   戻り値: 成功=True。失敗時は False を返し、outErrNum/outErrDesc に理由を
'   入れる(例外は外へ伝播させない=呼び出し側が必ず分岐で受けられる)。
'   先頭にBOM(U+FEFF)が残っていたら落とす。ADODB.Stream は Charset="utf-8"
'   のとき通常はBOMを落とすが、環境差で残ることが実機で確認されている
'   (旧 optGsTxt.ExtractPdfTextNoOcr の個別対処をここへ集約した)。
'
'   2026-07-31(R11-F2): 同じ形の実装が10箇所にあった(modExtractor/optDiffDoc/
'   optGsTxt/modChannel/modP2PIo/modPublish/modTelemetry/modBoard/modMentor/
'   modInsightIo)。COMの後始末(Resumeでハンドラを抜けてからClose/解放)を
'   1箇所でだけ正しく書けばよい状態にする。
' ----------------------------------------------------------------------------
Public Function ReadTextFileUtf8(ByVal filePath As String, ByRef outText As String, _
                                 Optional ByRef outErrNum As Long, _
                                 Optional ByRef outErrDesc As String) As Boolean
    ReadTextFileUtf8 = ReadWithCharset(filePath, "utf-8", outText, outErrNum, outErrDesc)
End Function

' ----------------------------------------------------------------------------
' ReadWithCharset - 文字コードを指定してテキストファイルを読む(実体)。
'   ReadTextFileUtf8 / ReadTextFileAuto の共通の下回り。例外は外へ出さない。
' ----------------------------------------------------------------------------
Private Function ReadWithCharset(ByVal filePath As String, ByVal charsetName As String, _
                                 ByRef outText As String, ByRef outErrNum As Long, _
                                 ByRef outErrDesc As String) As Boolean
    Dim st As Object
    Dim txt As String

    outText = ""
    outErrNum = 0
    outErrDesc = ""

    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2          ' adTypeText
    st.Charset = charsetName
    st.Open
    st.LoadFromFile filePath
    txt = CStr(st.ReadText(-1))   ' -1 = adReadAll
    st.Close
    Set st = Nothing

    If Len(txt) > 0 Then
        If Left$(txt, 1) = ChrW(&HFEFF) Then txt = Mid$(txt, 2)
    End If
    outText = txt
    ReadWithCharset = True
    Exit Function

Fail:
    ' Err は Resume でクリアされるので先に控える。
    outErrNum = Err.Number
    outErrDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きたエラーは
    ' 呼び出し元へ飛んで本来の原因を上書きする。後始末の前に Resume で
    ' ハンドラを抜ける(2026-07-30 実機err#462)。
    Resume ReadCleanup
ReadCleanup:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Set st = Nothing     ' COM解放(異常パス=半開きも確実に解放)
    On Error GoTo 0
    outText = ""
    ReadWithCharset = False
End Function

' ----------------------------------------------------------------------------
' ReadTextFileAuto - 文字コードを見分けてテキストファイルを読む。
'   2026-08-16(R33波3 W3-3): txt/md/csv を UTF-8 固定で読んでいたため、
'   CP932(Shift_JIS)で保存された社内資料が【読み込み成功・中身は全滅】に
'   なっていた。ADODB.Stream の UTF-8 デコードは不正バイト列で例外を投げず
'   U+FFFD(置換文字)へ落とすので、呼び出し側は失敗に気付けない。
'   さらに化け検知(modExtractor.GarbleRatio)は U+FFFD を数えていなかったため
'   どの関門も止めず、意味を失った本文がそのまま索引に入っていた。
'
'   判定の順番(BOM が最優先。BOM があるものを内容から推測し直さない):
'     1. 先頭バイトの BOM …… UTF-8 / UTF-16LE / UTF-16BE をそのまま採用
'     2. BOM 無し …… まず UTF-8 で読む。置換文字の比率が FFFD_LIMIT を
'        超えたら CP932 で読み直し、置換文字が【減ったときだけ】採用する
'        (減らないなら UTF-8 の結果を残す=推測で悪化させない)
'   どちらでも決め切れなかった本文は U+FFFD を含んだまま返る。呼び出し側の
'   化け検知が U+FFFD を数えるので、そこで拾われる(仕様「判定不能時は
'   化け検知に必ずかける」)。
' ----------------------------------------------------------------------------
Public Function ReadTextFileAuto(ByVal filePath As String, ByRef outText As String, _
                                 Optional ByRef outErrNum As Long, _
                                 Optional ByRef outErrDesc As String) As Boolean
    outText = ""
    outErrNum = 0
    outErrDesc = ""

    Dim bom As String: bom = BomCharset(filePath)
    If LenB(bom) > 0 Then
        ReadTextFileAuto = ReadWithCharset(filePath, bom, outText, outErrNum, outErrDesc)
        Exit Function
    End If

    If Not ReadWithCharset(filePath, "utf-8", outText, outErrNum, outErrDesc) Then Exit Function
    ReadTextFileAuto = True

    Dim r1 As Double: r1 = ReplacementRatio(outText)
    If Not AutoNeedsRetry(r1) Then Exit Function

    Dim alt As String, e2 As Long, d2 As String
    Dim altOk As Boolean
    altOk = ReadWithCharset(filePath, "shift_jis", alt, e2, d2)
    Dim r2 As Double
    If altOk Then r2 = ReplacementRatio(alt)
    If AutoTextPick(r1, altOk, r2) = "cp932" Then outText = alt
End Function

' ----------------------------------------------------------------------------
' AutoNeedsRetry / AutoTextPick - ReadTextFileAuto の採否判断【純関数】
'                                 (2026-08-16 R33H F31)
' ----------------------------------------------------------------------------
' なぜ切り出すか: W3-3 の「判定を1本にした」の【判断】は ADODB.Stream と同じ
'   関数の中にあり LO の実行テストが一度も撃てなかった(閾値も不等号も無検査
'   なのに、テストの見出しは守られていると読める)。引数だけで決まる形へ出す。
'   判断そのものは上の仕様どおりで変更なし。境界の意味だけ書き残す:
'   ・2%ちょうどは読み直さない側 ―― 本物の U+FFFD が数文字混ざる UTF-8 を
'     CP932 だと言い出さないための線。
'   ・同率は採らない。それは「CP932 と分かった」ではなく「判定できなかった」
'     なので答えを動かさない(化けは呼び出し側の化け検知が拾う)。
' ----------------------------------------------------------------------------
Public Function AutoNeedsRetry(ByVal utf8Ratio As Double) As Boolean
    AutoNeedsRetry = (utf8Ratio > FFFD_LIMIT)
End Function

' 戻り値は採用する読み方の名前: "utf8" / "cp932"。
'   altOk=False のとき altRatio は見ない(読めていない値に意味は無い)。
Public Function AutoTextPick(ByVal utf8Ratio As Double, ByVal altOk As Boolean, _
                             ByVal altRatio As Double) As String
    AutoTextPick = "utf8"
    If Not AutoNeedsRetry(utf8Ratio) Then Exit Function
    If Not altOk Then Exit Function
    If altRatio < utf8Ratio Then AutoTextPick = "cp932"
End Function

' ----------------------------------------------------------------------------
' BomCharset - 先頭バイトのBOMから ADODB.Stream の Charset 名を返す。
'   BOMが無い/読めないときは ""(呼び出し側は内容から推測する)。
'   "unicode"=UTF-16LE / "unicodeFFFE"=UTF-16BE は ADODB の登録名。
' ----------------------------------------------------------------------------
Private Function BomCharset(ByVal filePath As String) As String
    Dim st As Object
    Dim b As Variant
    Dim n As Long: n = -1

    On Error GoTo BomFail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1          ' adTypeBinary
    st.Open
    st.LoadFromFile filePath
    b = st.Read(3)
    st.Close
    Set st = Nothing

    On Error Resume Next
    n = UBound(b)
    Err.Clear
    On Error GoTo BomFail

    If n >= 2 Then
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then
            BomCharset = "utf-8"
            Exit Function
        End If
    End If
    If n >= 1 Then
        ' And は短絡しないので、要素の参照は境界チェックと分けている。
        If b(0) = &HFF Then
            If b(1) = &HFE Then BomCharset = "unicode"
        ElseIf b(0) = &HFE Then
            If b(1) = &HFF Then BomCharset = "unicodeFFFE"
        End If
    End If
    Exit Function

BomFail:
    Resume BomCleanup
BomCleanup:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Set st = Nothing
    On Error GoTo 0
    BomCharset = ""
End Function

' ----------------------------------------------------------------------------
' ReplacementRatio - 先頭 FFFD_SAMPLE 文字に占める U+FFFD の割合(0.0〜1.0)。
'   空白類は分母に入れない(modExtractor.GarbleRatio と同じ数え方)。
' ----------------------------------------------------------------------------
Public Function ReplacementRatio(ByVal s As String) As Double
    Dim n As Long: n = Len(s)
    If n > FFFD_SAMPLE Then n = FFFD_SAMPLE
    If n < 1 Then Exit Function

    Dim bad As Long, tot As Long
    Dim i As Long
    For i = 1 To n
        Dim c As Long: c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536          ' AscWの符号付き戻りを補正
        If c = 32 Or c = 9 Or c = 10 Or c = 13 Or c = &H3000 Then
            ' 空白類は分母に入れない
        Else
            tot = tot + 1
            If c = &HFFFD& Then bad = bad + 1
        End If
    Next i
    If tot > 0 Then ReplacementRatio = bad / tot
End Function

' ----------------------------------------------------------------------------
' WriteTextFileUtf8 - UTF-8テキストファイルを書く(ADODB.Stream・既存は上書き)。
'   戻り値: 成功=True。失敗時は False を返し、outErrNum/outErrDesc に理由を
'   入れる。BOMは ADODB.Stream の既定どおり付く(Excelが日本語CSVを正しく
'   開けるのはこのBOMのおかげなので、落とさない。modAnalytics のCSV書き出しが
'   これに依存している)。
'
'   2026-07-31(R11-F2): 同じ形の実装が9箇所にあった(modChannel/modP2PIo/
'   modPublish/modTelemetry/modAnalytics/modBoard/modMentor/modVault/
'   modInsightIo)。
' ----------------------------------------------------------------------------
Public Function WriteTextFileUtf8(ByVal filePath As String, ByVal content As String, _
                                  Optional ByRef outErrNum As Long, _
                                  Optional ByRef outErrDesc As String) As Boolean
    Dim st As Object

    outErrNum = 0
    outErrDesc = ""

    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2          ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText content
    st.SaveToFile filePath, 2   ' adSaveCreateOverWrite
    st.Close
    Set st = Nothing     ' COM解放(正常パス)
    WriteTextFileUtf8 = True
    Exit Function

Fail:
    outErrNum = Err.Number
    outErrDesc = Err.Description
    Resume WriteCleanup
WriteCleanup:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Set st = Nothing     ' COM解放(異常パス=半開きも確実に解放)
    On Error GoTo 0
    WriteTextFileUtf8 = False
End Function

' ----------------------------------------------------------------------------
' ElapsedMsSince - t0(Timerの値)からの経過ミリ秒。
'   VBAの Timer は0時に0へ戻るため、素の (Timer - t0) は日付をまたいだ瞬間に
'   大きな負の値になる。latency_ms が負になるとログの集計が壊れ、進捗の残り
'   時間見積りも跳ねる。負になったら1日ぶんを足して実経過へ戻す。
'   2026-07-31(R11-F2): modGateway 12箇所・modAsk 2箇所が素の式で、その他
'   9モジュールは各自でガードしていた(答えが2種類ある状態)。ここへ一本化する。
' ----------------------------------------------------------------------------
Public Function ElapsedMsSince(ByVal t0 As Double) As Double
    Dim ms As Double
    ms = (Timer - t0) * 1000#
    If ms < 0 Then ms = ms + MS_PER_DAY     ' 日跨ぎ(Timerが0へ戻った)
    If ms < 0 Then ms = 0                   ' それでも負なら0(時計の巻き戻し)
    ElapsedMsSince = ms
End Function

' ----------------------------------------------------------------------------
' GuardExpired - 再入ガードの失効判定(2026-08-04 R15-1a・実機第4報 RC5)。
'   startAt   : ガードを立てた時刻
'   beatAt    : 最終ハートビート(まだ1度も打たれていなければ0でよい)
'   nowAt     : 現在時刻
'   limitMin  : 無音がこの分数だけ続いたら失効とみなす(既定30分)
'   absLimitMin: 開始からの絶対上限(既定480分=8時間)。ビートがどれだけ
'                新しくても、これを超えたら無条件で失効させる(R15波1b裁定c)。
'                0以下は【無効】(この製品の 0=無効 の慣習。R15-FixA FA-9)。
'                8時間を超える正当なバッチ(254頁の資料を複数まとめて取込む)
'                でガードが解け、二重取込の入口が開くのを止めるため、
'                呼び出し側(modShelfBatch.GuardExpiredNow)は0を渡す。
'                ビート源が全て実作業由来になった(FA-8でTouchBusyを中断
'                ハンドラから外した)ので、延命ループの心配はもう無い。
'
'   従来は「開始から30分」で自動失効させていた。OCR付きの取込は実機で
'   85〜127分かかるため、t=30分の時点で取込中にもかかわらず全ボタンが
'   解放され、二重取込が構造的に起こり得た(憲章§3-5)。かといって単純に
'   上限を延ばすと、焼き付いたガードの自己回復がその分だけ遅れる。
'   「最後に生きている印(ビート)から数える」ことで、動いている間は
'   何時間でも守り、止まった瞬間から30分で自己回復するようになる。
'
'   基準は max(startAt, beatAt)。前回の取込で残ったビートが次の取込の開始より
'   古い場合に「昔のビート」で判定してしまわないよう、必ず新しい方を採る。
'   startAt が0(=時刻不明)なら基準も0となり必ず失効側へ倒れる。判断できない
'   ガードで全ボタンを殺し続けるより奪い返す方が害が小さい(modUiLockと同じ)。
'   時計が巻き戻った場合(nowAt < 基準)は DateDiff が負になり失効しない。
'   これは変更前の式と同じ挙動で、進んだ時計が戻るまでの間だけ守りが続く。
'
'   絶対上限を足した理由(R15波1の敵対的自己点検): modUiLock.BlockIfIngesting
'   自身の実況が SetStage 経由でビートを打つと、ガードが焼き付いた状態でも
'   利用者が押すたびに失効がlimitMin分だけ延び続け、永久に解けなくなる
'   (a対応でこの経路自体は塞いだが、他の未知の経路が同じ穴を再現しても
'   ここで必ず止まるよう、時間で数える最終防衛線を独立に置く)。
' ----------------------------------------------------------------------------
Public Function GuardExpired(ByVal startAt As Date, ByVal beatAt As Date, _
                             ByVal nowAt As Date, ByVal limitMin As Long, _
                             Optional ByVal absLimitMin As Long = 480) As Boolean
    Dim baseAt As Date
    baseAt = startAt
    If beatAt > baseAt Then baseAt = beatAt
    If DateDiff("n", baseAt, nowAt) >= limitMin Then
        GuardExpired = True
        Exit Function
    End If
    If absLimitMin <= 0 Then Exit Function
    GuardExpired = (DateDiff("n", startAt, nowAt) >= absLimitMin)
End Function

' ----------------------------------------------------------------------------
' IsoDate / IsoDateTime - カレンダー設定に左右されないグレゴリオ暦の
'   日付文字列("yyyy-mm-dd" / "yyyy-mm-dd hh:nn:ss")を作る。
'
'   2026-08-01(R12-1-4): Windowsの地域設定「カレンダーの種類=和暦」の端末では
'   VBAの Format$(Date, "yyyy-mm-dd") が元号年を返す(令和8年→"0008-08-01")。
'   本製品は保険営業向けで和暦運用が濃い業種であり、この設定は日本語Windowsの
'   正規の選択肢である。日付を文字列で永続化している箇所(端末失効タイマーの
'   guard_last_reach、連続利用日数の last_used_date/last_streak_date、既読印の
'   日付 等)が元号年で書かれると、
'     ・読み側 CDate("0008-08-01") が過去年に解釈されれば失効判定が暴発し
'       知識が消える(憲章§3-5に反する最悪の壊れ方)
'     ・変換エラーになれば失効タイマーが永久に発火しない(無言の機能喪失)
'   という両方向の事故になる。左辺の Date はシリアル値で常に西暦なので、
'   文字列側だけが元号化する非対称が問題の核である。
'
'   Year()/Month()/Day()/Hour()/Minute()/Second() はカレンダー設定に依存せず
'   常に西暦(グレゴリオ暦)の数値を返すため、数値から自前で連結すれば
'   設定に関わらず同じ文字列になる。日付を「文字列として書く・比べる」箇所は
'   Format$ を使わず必ずこの2関数を通すこと(憲章§4-5: 答えは1つ)。
' ----------------------------------------------------------------------------
Public Function IsoDate(ByVal d As Date) As String
    IsoDate = Pad0(Year(d), 4) & "-" & Pad0(Month(d), 2) & "-" & Pad0(Day(d), 2)
End Function

Public Function IsoDateTime(ByVal dt As Date) As String
    IsoDateTime = IsoDate(dt) & " " & Pad0(Hour(dt), 2) & ":" & _
                  Pad0(Minute(dt), 2) & ":" & Pad0(Second(dt), 2)
End Function

' ----------------------------------------------------------------------------
' IsoDateCompact / IsoYm / IsoYear - 区切り無しの日・月・年キー(2026-08-01
'   R12-3-10)。統計キー("sv:d:20260801" 等)や1日1回の判定キーに使う。
'
'   IsoDate と同じ理由(和暦カレンダー端末で Format$(Date,"yyyymmdd") が元号年に
'   なる)だが、こちらは失効ではなく【集計の分断】として現れる: 元号年の
'   "00080801" と西暦の "20260801" が別キーになり、端末の設定を変えた日を境に
'   その日の記録が二重化し、連番も昇順に並ばなくなる(気付ける人がいない)。
'
'   対象は「キーとして保存・比較する日付」だけ。ファイル名や nonce に使う
'   Format$(Now,"yyyymmddhhnnss") は【対象外】のままにしている。あれは一意性の
'   ためのラベルで、和暦になっても衝突しない限り実害が無く、むしろ既存
'   ファイル名との照合規則を変える方が危ないため(判断根拠を残す)。
' ----------------------------------------------------------------------------
Public Function IsoDateCompact(ByVal d As Date) As String
    IsoDateCompact = Pad0(Year(d), 4) & Pad0(Month(d), 2) & Pad0(Day(d), 2)
End Function

Public Function IsoYm(ByVal d As Date) As String
    IsoYm = Pad0(Year(d), 4) & Pad0(Month(d), 2)
End Function

Public Function IsoYear(ByVal d As Date) As String
    IsoYear = Pad0(Year(d), 4)
End Function

' ----------------------------------------------------------------------------
' NormalizeIsoDate - セルへ書いた日付文字列が「セルの日付型自動変換」で
'   ロケール短形式("2026/08/01")へ化けて戻ってくる現象を吸収し、
'   "yyyy-mm-dd" へ戻す。日付として読めない文字列はそのまま返す
'   (日付ではないキーや値を壊さないため)。
'
'   2026-08-01(R12-1-3): 連続利用日数(streak)が実機で永久に1のままだった
'   原因がこれ。my_stats の値セルは書式指定なし(General)なので、書いた
'   "2026-08-01" がExcelの日付型セルになり、CStr で読むと "2026/08/01" に
'   なる。表示系(modUIShelf.ShortDate 等)は既にIsDate分岐で対処済みだが、
'   modStats.TouchToday の等値比較だけが素の文字列比較で、毎回「連続が
'   途切れた」と判定していた。
' ----------------------------------------------------------------------------
Public Function NormalizeIsoDate(ByVal s As String) As String
    NormalizeIsoDate = s

    Dim t As String
    t = Trim$(s)
    If LenB(t) = 0 Then Exit Function

    ' 年から始まる表記(この製品が書く形と、日本語Windowsの短い日付形式
    ' "yyyy/mm/dd" の両方)は数字だけを見て組み直す。CDate を通さないので
    ' ロケールにもカレンダー設定にも一切依存しない(和暦端末で CDate が
    ' 元号として解釈しても、こちらの経路なら影響を受けない)。
    Dim iso As String
    iso = YearFirstIso(t)
    If LenB(iso) > 0 Then
        NormalizeIsoDate = iso
        Exit Function
    End If

    ' 年から始まらない表記(他ロケールの短い日付形式など)は最後の手段として
    ' CDate に任せる。読めなければ元の文字列をそのまま返す
    ' (my_stats には日付でない値も入るので、勝手に書き換えてはならない)。
    If Not IsDate(t) Then Exit Function

    On Error Resume Next
    Err.Clear
    Dim d As Date
    d = CDate(t)
    If Err.Number = 0 Then NormalizeIsoDate = IsoDate(d)
    Err.Clear
    On Error GoTo 0
End Function

' "yyyy-m-d" / "yyyy/mm/dd" / "yyyy.mm.dd"(後ろに時刻が付いていてもよい)を
' "yyyy-mm-dd" にする。当てはまらなければ空文字を返す。
Private Function YearFirstIso(ByVal t As String) As String
    Dim head As String
    head = t
    Dim sp As Long
    sp = InStr(head, " ")
    If sp > 0 Then head = Left$(head, sp - 1)      ' 時刻部は落とす
    head = Replace(Replace(head, "/", "-"), ".", "-")

    Dim parts() As String
    parts = Split(head, "-")
    If UBound(parts) <> 2 Then Exit Function

    If Len(parts(0)) <> 4 Then Exit Function
    If Not AllDigits(parts(0)) Then Exit Function
    If Len(parts(1)) < 1 Or Len(parts(1)) > 2 Then Exit Function
    If Not AllDigits(parts(1)) Then Exit Function
    If Len(parts(2)) < 1 Or Len(parts(2)) > 2 Then Exit Function
    If Not AllDigits(parts(2)) Then Exit Function

    Dim mn As Long: mn = CLng(parts(1))
    Dim dn As Long: dn = CLng(parts(2))
    If mn < 1 Or mn > 12 Then Exit Function
    If dn < 1 Or dn > 31 Then Exit Function

    YearFirstIso = parts(0) & "-" & Pad0(mn, 2) & "-" & Pad0(dn, 2)
End Function

' 半角数字だけで出来ているか(IsNumericは "1e2" や " 1 " も通すので使わない)。
Private Function AllDigits(ByVal s As String) As Boolean
    If LenB(s) = 0 Then Exit Function
    Dim i As Long
    For i = 1 To Len(s)
        Dim c As String: c = Mid$(s, i, 1)
        If c < "0" Or c > "9" Then Exit Function
    Next i
    AllDigits = True
End Function

' 数値を左ゼロ詰めの固定桁文字列にする(IsoDate/IsoDateTime専用の内部部品)。
Private Function Pad0(ByVal n As Long, ByVal digits As Long) As String
    Dim t As String
    If n < 0 Then
        t = "0"          ' 起こり得ないが、桁数の約束だけは壊さない
    Else
        t = CStr(n)
    End If
    Do While Len(t) < digits
        t = "0" & t
    Loop
    Pad0 = t
End Function

' ----------------------------------------------------------------------------
' BlendPerItemMs - 1件あたり所要ミリ秒の更新(R7 B-1)。直近バッチの実測と
'   現在値の平均(=直近2バッチの移動平均)を返す。
'   2026-07-31(R11-F2): modEmbed と modEnrich に同じ実装が2本あった。
'   経過時間の算出は ElapsedMsSince に任せる(日跨ぎもそこで吸収される)。
' ----------------------------------------------------------------------------
Public Function BlendPerItemMs(ByVal curMs As Double, ByVal t0 As Double, _
                               ByVal itemCount As Long) As Double
    BlendPerItemMs = curMs
    If itemCount < 1 Then Exit Function

    Dim thisMs As Double
    thisMs = ElapsedMsSince(t0) / CDbl(itemCount)
    If curMs <= 0 Then
        BlendPerItemMs = thisMs
    Else
        BlendPerItemMs = (curMs + thisMs) / 2#
    End If
End Function

' ----------------------------------------------------------------------------
' SanitizeForCell - 数式インジェクション対策。未信頼テキスト(共有フォルダ・
'   部門パック・他モジュール由来のファイル内容等、自分以外が書いたテキスト)
'   をセルへ .Value = で書く直前に必ず通す。先頭文字が =/+/-/@ だとExcelが
'   数式として解釈してしまうため、アポストロフィを前置してテキスト強制する。
'
'   2026-08-01(R12-2-1): 元は modChatLog の Private 実装(チャット履歴シート
'   のみに適用)だったが、未信頼テキストをセルへ書く経路が他に3つ
'   (modInsightIo.AppendRow / modPack.ImportChunksDedup / modChannel.SyncChannel
'   経由の modStats.SetStatValue)あり、そこは無防備だった(セキュリティ監査3)。
'   同じ答えを2箇所以上に書かないという憲章§4-5に基づき、ここへ1本化する。
'   modChatLog はこちらへ委譲する(重複実装を残さない)。
' ----------------------------------------------------------------------------
Public Function SanitizeForCell(ByVal s As String) As String
    If LenB(s) = 0 Then
        SanitizeForCell = s
        Exit Function
    End If
    Dim c As String
    c = Left$(s, 1)
    If c = "=" Or c = "+" Or c = "-" Or c = "@" Then
        SanitizeForCell = "'" & s
    Else
        SanitizeForCell = s
    End If
End Function

' ----------------------------------------------------------------------------
' GsPageIsBlank - 空白類しか無いページか。
'   Trim$ は半角スペースしか落とさないため使わない(改行だけのページを
'   「中身あり」と数えると出典ページ番号が丸ごと1つずれる)。
' ----------------------------------------------------------------------------
Public Function GsPageIsBlank(ByVal s As String) As Boolean
    GsPageIsBlank = (CleanTextLen(s) = 0)
End Function

' ----------------------------------------------------------------------------
' CleanTextLen - 空白類(改行/タブ/改ページ/半角空白/全角空白)を除いた文字数。
' ----------------------------------------------------------------------------
Public Function CleanTextLen(ByVal s As String) As Long
    Dim t As String

    t = Replace(s, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, vbTab, "")
    t = Replace(t, Chr$(12), "")   ' 改ページ(GSのtxtwriteはページをChr(12)で区切る)
    t = Replace(t, " ", "")
    t = Replace(t, ChrW(&H3000), "")
    CleanTextLen = Len(t)
End Function

' ----------------------------------------------------------------------------
' GsPageBounds - Ghostscript txtwrite の出力を改ページで割ったとき、
'   実質何ページぶんあるかと、その範囲を返す。
'   戻り値: ページ数(中身が全く無ければ0)。
'   firstIdx/lastIdx: Split(txt, Chr(12)) の結果配列における先頭/末尾の添字
'   (先頭側・末尾側の空ページを対称に読み飛ばした位置)。中間の空ページは
'   落とさない(本当に白紙のページがあり得るので、落とすと以降のページ番号が
'   全部ずれる)。
'
'   2026-07-31(R11-F2): この添字計算は optOcrCore.GsPageCount(採否しきい値の
'   分母)と modExtractor.BuildPagesFromGsText(出典ページ番号)に同じものが
'   2本あり、ズレると「出典 p.5 を開いても別のページが出る」という誰も
'   気付けない壊れ方をするため、突き合わせテストで守っていた。実装を1本に
'   したので、突き合わせではなくこの関数の単体テストで固定する。
' ----------------------------------------------------------------------------
Public Function GsPageBounds(ByVal txt As String, ByRef firstIdx As Long, _
                             ByRef lastIdx As Long) As Long
    firstIdx = 0
    lastIdx = -1
    If CleanTextLen(txt) = 0 Then Exit Function   ' 全体が空白類だけ=0ページ

    Dim parts() As String
    parts = Split(txt, Chr$(12))
    firstIdx = LBound(parts)
    lastIdx = UBound(parts)

    ' 上の早期returnにより、必ずどこかに中身のあるページがある=ループは必ず解ける。
    Do While firstIdx < lastIdx
        If CleanTextLen(parts(firstIdx)) > 0 Then Exit Do
        firstIdx = firstIdx + 1
    Loop
    Do While lastIdx > firstIdx
        If CleanTextLen(parts(lastIdx)) > 0 Then Exit Do
        lastIdx = lastIdx - 1
    Loop

    GsPageBounds = lastIdx - firstIdx + 1
End Function

' ----------------------------------------------------------------------------
' AppendStepBuf - 段(step)ごとの所要時間バッファへ1件足す(2026-08-03 R13 F8)
' ----------------------------------------------------------------------------
' 質問1回のあいだに走った各段(expand/emb/rerank/draft/verify…)の所要msを、
' 1本の短い文字列へ畳んでいくための純ロジック。呼ぶのは modGateway だけだが、
' 中身は文字列処理しかないのでここへ置く(modGatewayは28,000字帯に近く、
' またこの書式はLOの実行テストから直接固定したい)。
'
' 書式: "expand=1200;emb=6x830;rerank=900"
'   ・初出の段は "<名前>=<ms>"。
'   ・同じ段が再登場したら "<回数>x<平均ms>" へ畳む(段の並びではなく
'     「何回・平均どれくらい」が読みたい情報。多段検索の emb がこれ)。
'   ・区切り文字(";" "=")は名前から追い出し、名前は32字で打ち切る。
'     ここを素通しにすると、段名1つで行の書式全体が壊れる。
'   ・全体が maxChars 以上になったら【新しい段名の追加だけ】をやめる
'     (既出の段の畳み込みは続けるので、回数と平均は最後まで正しい)。
' ----------------------------------------------------------------------------
Public Function AppendStepBuf(ByVal buf As String, ByVal stepName As String, _
                              ByVal ms As Long, ByVal maxChars As Long) As String
    Dim nm As String
    nm = Replace(Replace(Trim$(stepName), ";", "_"), "=", "_")
    If LenB(nm) = 0 Then nm = "step"
    If Len(nm) > 32 Then nm = Left$(nm, 32)
    Dim v As Long: v = ms
    If v < 0 Then v = 0
    Dim lim As Long: lim = maxChars
    If lim < 32 Then lim = 32

    AppendStepBuf = buf
    Dim parts() As String
    Dim i As Long
    If LenB(buf) > 0 Then
        parts = Split(buf, ";")
        For i = LBound(parts) To UBound(parts)
            If InStr(parts(i), nm & "=") = 1 Then
                Dim cnt As Long, tot As Long
                DecodeStepPart Mid$(parts(i), Len(nm) + 2), cnt, tot
                cnt = cnt + 1
                tot = tot + v
                parts(i) = nm & "=" & cnt & "x" & CLng(tot / cnt)
                AppendStepBuf = Join(parts, ";")
                Exit Function
            End If
        Next i
    End If

    ' 初出。上限に達していたら足さない(既出ぶんの精度は落とさない)。
    If Len(buf) >= lim Then Exit Function
    If LenB(buf) = 0 Then
        AppendStepBuf = nm & "=" & v
    Else
        AppendStepBuf = buf & ";" & nm & "=" & v
    End If
End Function

' ----------------------------------------------------------------------------
' IngestChunksDetail - usage_log "ingest" 行のdetailに、生成/重複の内訳を
'   添える(2026-08-04 R15-8a・実機第4報 RC1)。
'   画面の「127」と記録の「125」が食い違って見える主因は、生成チャンク数
'   (gen=modChunker.ChunkPagesEx の戻り値)とハッシュ重複排除後の保存数
'   (accepted)の乖離だが、この差分がログにもUIにも一切残らず診断できな
'   かった。dup(=gen-accepted)が0のときは従来どおり "chunks=N" のみを返し、
'   既存ログの書式・grep習慣との互換を優先する(内訳は差分がある時だけ)。
' ----------------------------------------------------------------------------
Public Function IngestChunksDetail(ByVal accepted As Long, ByVal gen As Long) As String
    Dim dup As Long
    dup = gen - accepted
    If dup < 0 Then dup = 0   ' 起こり得ないが、負の重複という無意味な表示は避ける

    If dup = 0 Then
        IngestChunksDetail = "chunks=" & accepted
    Else
        IngestChunksDetail = "chunks=" & accepted & " (gen=" & gen & " dup=" & dup & ")"
    End If
End Function

' "830"(1回) / "6x830"(6回・平均830) のどちらの形も回数と合計へ戻す。
' 読めない値は0件0msとして扱う(壊れた値で以降の平均を汚さない)。
Private Sub DecodeStepPart(ByVal part As String, ByRef outCount As Long, ByRef outTotal As Long)
    outCount = 0
    outTotal = 0
    On Error Resume Next
    Dim p As Long: p = InStr(part, "x")
    If p > 0 Then
        outCount = CLng(Val(Left$(part, p - 1)))
        outTotal = outCount * CLng(Val(Mid$(part, p + 1)))
    Else
        outCount = 1
        outTotal = CLng(Val(part))
    End If
    If outCount < 0 Then outCount = 0
    If outTotal < 0 Then outTotal = 0
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' BuildWorkExcelCmd - 「作業用Excelを開く」ボタン(R16-2a・modWorkExcel)が
'   起動するコマンド行の組み立て(純粋な文字列処理)。appPathは
'   Application.Path(EXCEL.EXEがあるフォルダ)を想定するが、この関数自体は
'   ただの文字列連結でExcelには依存しない。末尾に区切り文字が付いていても
'   (環境差で付くことがある。ドライブ直下等)全て取り除いてから
'   "\EXCEL.EXE" を1つだけ足す(二重にしない)。全体を""で囲み、末尾に
'   " /x"(新規プロセスで空のExcelを開くスイッチ。ドキュメント既定を開き
'   直さない・別プロセスなのでこのブックの応答なし状態の影響を受けない)を
'   付ける。パスの空白・日本語はそのまま素通しでよい(""で囲むため)。
'   区切り文字はChr$(92)で比較する(ソース中に区切り文字を直書きした文字列
'   リテラルの末尾へ置くとLO構文チェッカーが沈黙ハングするため。
'   docs/dev/EDGE_CASES.md §1.3)。
' ----------------------------------------------------------------------------
Public Function BuildWorkExcelCmd(ByVal appPath As String) As String
    Dim p As String: p = appPath
    Do While Len(p) > 0 And Right$(p, 1) = Chr$(92)
        p = Left$(p, Len(p) - 1)
    Loop
    BuildWorkExcelCmd = """" & p & "\EXCEL.EXE"" /x"
End Function

' ----------------------------------------------------------------------------
' DECLINE_MEMO_HEAD / IsDeclineNote - 「見送りました」のメモかどうか
'   (2026-08-04 R15-FixB FB-5・レビューB-M)。
' ----------------------------------------------------------------------------
' 事前確認(R15-7b)で利用者が「いいえ」を選んだ資料は、取込に【失敗した】
' わけではない。それなのに一括取込の集計・トースト・モーダルでは失敗と同じ
' 数に混ざり、「1件は取り込めませんでした。状態をご確認ください」と、
' 自分で見送っただけの利用者に不具合を疑わせていた(憲章§4-1)。
' 見送りを別に数えるには「このメモは見送りか」を機械で言えなければならない。
' 判定の一次情報は先頭句のみで、文面の残り(推定分数・次の一手)は
' optOcrEta.OcrDeclineMemoFor が持つ。同じ文字列を2箇所に書かないため、
' 先頭句だけを【opt層とコア層の両方から見える基盤層】に1つ置く。
' 先頭一致にするのは、分数のような可変部分が入っても判定が揺れないように
' するため(文中の部分一致にすると、利用者が付けた説明文と衝突し得る)。
Public Function IsDeclineNote(ByVal note As String) As Boolean
    Dim s As String: s = LTrim$(note)
    If LenB(s) = 0 Then Exit Function
    IsDeclineNote = (Left$(s, Len(DECLINE_MEMO_HEAD)) = DECLINE_MEMO_HEAD)
End Function
