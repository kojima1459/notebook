Attribute VB_Name = "optOcrEta"
Option Explicit

' ============================================================================
' optOcrEta - 画像PDF OCRの「進捗の見せ方」と「打ち切りの言い方」だけを集めた
'             純ロジックモジュール(2026-08-04 R15-5b・R4準拠)
' ----------------------------------------------------------------------------
' 役割:
'   OCR中の進捗バナーの文面(OcrPageBanner / BatchLabel / RemainingText)と、
'   途中で止まった・頁が欠けたときに本棚カードへ出す正直なメモ
'   (OcrCapMemoFor / OcrAbortMemoFor / OcrPartialMemoFor)、および1資料
'   あたりの画像化待ち予算の残り配分(RemainingWaitSec)。
'   いずれも副作用ゼロ(Worksheets/Range/Application/ThisWorkbook/MsgBoxに
'   触れない)で、文字列と数値だけで完結する。
'
' なぜ optOcrCore から分けたのか(憲章§4-6):
'   optOcrCore が30,000字上限まで残り2,068字となり、R15-5(ETAと終了時刻)
'   R15-6(中断メモ)の追記が入らなくなった。GSコマンドの組み立て・ページ上限の
'   算数(=Ghostscriptの都合)と、進捗の文面・打ち切りのメモ(=利用者への
'   伝え方)は別の関心事なので、後者をまとめてこちらへ【移設】した
'   (optOcrCore は純減。あちらへの追記は禁止)。
'
' 設計判断:
'   ・現在時刻は引数(nowAt)で受け取る。Now をこの中で呼ぶと「終了目安」の
'     ゴールデンテストが実行時刻に依存して壊れるため、副作用の入口は
'     呼び出し元(optOcrPage)に置く。nowAt=0 は「時刻不明」として終了目安を
'     出さない(嘘の時刻を出すくらいなら黙る。憲章§4-1)。
'   ・時刻の書式は Format$(…, "hh:mm") と同じ結果を、Hour/Minute/Second の
'     整数計算だけで組み立てる。Format の書式コードの解釈は環境(ロケール・
'     LibreOffice)で揺れるが、整数計算なら揺れようがない。
'   ・「残り」は60秒未満だけ秒、それ以上は分。秒のまま「残り約4823秒」と
'     出しても人は時間の見当が付かない(実機第4報 RC2の要望)。
' ============================================================================

' R15-7b: 1頁あたりの所要ミリ秒が1度も実測されていない端末で、事前確認の
' 推定に使う仮のレート。実機第4報の254頁の資料は「1頁25秒前後」で進んで
' いた(76頁で約32分)。低く見積もると「15分で終わる」と言って2時間かかる
' ことになるので、実測が無いあいだは実機で観測できた値をそのまま使う。
Private Const DEFAULT_PAGE_MS As Double = 25000#

' R15-7c: 総頁が分かっているときに1頁あたり最低限見込むGS描画秒数。
' 実機第4報 RC10: 254頁を20頁ずつ13バッチ描くのに、1資料あたりの絶対上限が
' 1200秒(既定)しかなく、後半のバッチが「残り10秒」で必ず時間切れになる。
' 描画は150dpiで1頁あたり1〜3秒だが、EDRのスキャンが挟まる端末では数倍に
' なるため、8秒/頁を「頁数に連動して自動で確保する下限」として使う。
Private Const GS_SEC_PER_PAGE As Long = 8

' R16-2c(2026-08-05): バナー末尾に付ける「フリーズは正常」の一言。定数化して
' OcrPageBannerの2つのreturn文で同じ文字列を重複させない(憲章§4-5)。
Private Const FREEZE_NOTE As String = "※応答なし表示でも処理中"

Public Function Ping() As Boolean
    Ping = True
End Function

' ----------------------------------------------------------------------------
' BatchLabel - 画像化フェーズの進捗ラベル(2026-08-04 R15-5b。optOcrPage の
'   Private 実装を、テストできる純関数としてここへ移設した)。
'     総頁が未確定: 「バッチ2」          … 分母もページ数も名乗らない
'     総頁が確定  : 「バッチ2/13 (全254頁)」
'   総頁は R15-5a により txtwrite 分類パスの gs_out.log から取込開始前に
'   判明することが多い。判明していない場合の見え方は従来と1文字も変えない。
' ----------------------------------------------------------------------------
Public Function BatchLabel(ByVal batchIdx As Long, ByVal knownTotal As Long, _
                           ByVal batchSize As Long) As String
    BatchLabel = "バッチ" & batchIdx
    If knownTotal <= 0 Or batchSize < 1 Then Exit Function
    BatchLabel = BatchLabel & "/" & optOcrCore.BatchCountFor(knownTotal, batchSize) & _
        " (全" & knownTotal & "頁)"
End Function

' ----------------------------------------------------------------------------
' OcrPageBanner - OCR中の進捗バナーの文面(R14-4a / R14-F9 / R15-5b)。
'   総頁が【確定してから】だけ分母とETAを出す:
'     未確定: 「OCR中… 3頁目 (バッチ 1)」    分母も総バッチ数も出さない
'     確定後: 「OCR中… 3/44頁 (バッチ 1/3) 残り約2分・終了目安 14:05」
'   totalPages が現在頁より小さい値は【未確定】として扱う。
'   avgMsPerPage は1頁あたりの実測(まだ実測が無いときは過去の実績。0なら
'   残り時間そのものを出さない=跳ね回る表示を作らない)。
'   nowAt は呼び出し元が渡す現在時刻(0=不明なら終了目安を出さない)。
' ----------------------------------------------------------------------------
Public Function OcrPageBanner(ByVal pageNo As Long, ByVal totalPages As Long, _
                              ByVal batchIdx As Long, ByVal batchCount As Long, _
                              ByVal avgMsPerPage As Double, ByVal nowAt As Date) As String
    Dim p As Long: p = pageNo
    If p < 1 Then p = 1

    If totalPages < p Or batchCount < batchIdx Then
        OcrPageBanner = "OCR中… " & p & "頁目 (バッチ " & batchIdx & ")" & " " & FREEZE_NOTE
        Exit Function
    End If

    Dim n As Long: n = totalPages
    Dim s As String
    s = "OCR中… " & p & "/" & n & "頁 (バッチ " & batchIdx & "/" & batchCount & ")"

    If avgMsPerPage > 0 And n > p Then
        Dim remSec As Long
        remSec = CLng(Int((CDbl(n - p) * avgMsPerPage) / 1000# + 0.5))
        If remSec < 1 Then remSec = 1
        s = s & " " & RemainingText(remSec, nowAt)
    End If

    OcrPageBanner = s & " " & FREEZE_NOTE
End Function

' ----------------------------------------------------------------------------
' RemainingText - 残り時間と終了目安の文面(2026-08-04 R15-5b)。
'   60秒未満 : 「残り約45秒」
'   60秒以上 : 「残り約3分」(四捨五入。最低1分)
'   nowAt が入っていれば「・終了目安 HH:MM」を足す。
'   実機第4報 RC2: 「あと何分で終わるのか」「何時に終わるのか」が最後まで
'   分からないため、利用者は席を立てず、止まっているのかも判断できなかった。
' ----------------------------------------------------------------------------
Public Function RemainingText(ByVal remSec As Long, ByVal nowAt As Date) As String
    Dim r As Long: r = remSec
    If r < 1 Then r = 1

    Dim s As String
    If r < 60 Then
        s = "残り約" & r & "秒"
    Else
        Dim m As Long: m = CLng(Int(CDbl(r) / 60# + 0.5))
        If m < 1 Then m = 1
        s = "残り約" & m & "分"
    End If

    If CDbl(nowAt) > 0# Then s = s & "・終了目安 " & EndClock(nowAt, r)
    RemainingText = s
End Function

' ----------------------------------------------------------------------------
' OcrCapMemoFor - 上限ページで打ち切ったときに本棚カードへ出す正直なメモ
'   (2026-08-03 R14-4c / R14-F7。R15-5bで optOcrCore から移設)。
'   打ち切っていなければ ""。
'   truncated=上限で打ち切ったか / keptN=取り込めた頁数 /
'   capN=【実際に効く上限】(config vision_pdf_max_pages にハード上限の
'   クランプを掛けたあとの値。R15-4c: 掛けずに config の生値を出していたため、
'   config=300・ハード上限200 のときに「100ページは読み取れませんでした」と
'   いう事実と違うメモが出ていた)。総ページ数は不明なので名乗らない。
' ----------------------------------------------------------------------------
Public Function OcrCapMemoFor(ByVal truncated As Boolean, ByVal keptN As Long, _
                              ByVal capN As Long) As String
    If Not truncated Then Exit Function
    Dim k As Long: k = keptN
    If k < 0 Then k = 0
    Dim c As Long: c = capN
    If c < k Then c = k          ' 上限が不明・壊れた値なら取り込めた数で言う

    OcrCapMemoFor = "設定上限" & c & "ページのうち先頭" & k & _
        "ページを取り込みました(設定 vision_pdf_max_pages で変更できます)"
    If c > k Then
        OcrCapMemoFor = OcrCapMemoFor & "。" & (c - k) & "ページは読み取れませんでした"
    End If
End Function

' ----------------------------------------------------------------------------
' OcrAbortMemoFor - 途中で読み取りを打ち切ったときのメモ
'   (2026-08-03 R14-F2 / 2026-08-04 R15-4b・R15-6b)。
'   keptN    : ここまでに読めた頁数
'   reason   : "cancel"=利用者が中断ボタンを押した
'              "limit" =AI利用の上限に当たったとみられる(頁失敗が3連続)
'              それ以外(既定 "error")=変換エラーまたは時間切れ
'   hasCache : 読めた頁を次回まで持ち越せるか。True なら「続きから再開」、
'              False なら「最初から再試行」。R15波2の呼び出し元は常に False
'              (頁キャッシュは次波 R15-7 の担当。嘘を言わないための口だけ用意)。
'   なぜ理由で言い分けるのか: 「変換エラーまたは時間切れ」と書かれたカードを
'   見た利用者は、自分が中断ボタンを押したことと結び付けられない。上限に
'   当たったのなら打てる手は「時間をおく」だけで、再取込を今すぐ繰り返しても
'   同じ結果にしかならない(実機第4報 RC3/RC6)。
' ----------------------------------------------------------------------------
Public Function OcrAbortMemoFor(ByVal keptN As Long, ByVal reason As String, _
                                ByVal hasCache As Boolean) As String
    Dim k As Long: k = keptN
    If k < 0 Then k = 0

    ' 次に何が起きるかは1回だけ・同じ言い方で言う。キャッシュが無いのに
    ' 「続きから再開します」と書いたら嘘になる(R14-F2で一度直した嘘なので
    ' 二度と戻さない)。中断だけ「最初から」と明示するのは、押した本人が
    ' 「途中まで入ったぶんは活きるのか」を最も気にするため。
    Dim again As String
    If hasCache Then
        again = "もう一度取り込むと続きから再開します"
    ElseIf reason = "cancel" Then
        again = "もう一度取り込むと最初から再試行します"
    Else
        again = "もう一度取り込むと再試行します"
    End If

    If reason = "cancel" Then
        ' R15-FixB(FB-14): 1頁も読めていないうちに押された中断で「ここまでの
        ' 0頁を保存しました」と言うのは日本語として不自然なうえ、保存したもの
        ' など何も無いのに保存を名乗る(憲章§4-1)。頁数の話ごと落として、
        ' 次に何が起きるかだけを言う。
        If k <= 0 Then
            OcrAbortMemoFor = "利用者の操作で中断しました(" & again & ")"
        Else
            OcrAbortMemoFor = "利用者の操作で中断しました。ここまでの" & k & _
                "頁を保存しました(" & again & ")"
        End If
    ElseIf reason = "limit" Then
        OcrAbortMemoFor = "AI利用の上限に達した可能性があります。" & _
            "時間をおいて再度取り込んでください(ここまでの" & k & "頁を保存しました)"
    Else
        OcrAbortMemoFor = "読み取りを" & k & "頁で中断しました" & _
            "(変換エラーまたは時間切れ)。" & again
    End If
End Function

' ----------------------------------------------------------------------------
' OcrPartialMemoFor - 全ページ描けたのに一部の頁だけ読み取れなかったときの
'   メモ(2026-08-04 R15-4a・実機第4報 RC6)。
'   okN=読めた頁数 / failN=読めなかった頁数。失敗0なら ""。
'   従来この形は failN を数えるだけで status="done" のまま本棚に並び、
'   頁が【無言で】欠けていた(憲章§4-1違反)。何頁のうち何頁が欠けたのかを
'   数字で言い、もう一度取り込めば再試行されることまで書く。
' ----------------------------------------------------------------------------
'   totalKnown(2026-08-04 R15-FixA FA-4): 資料の総頁数が分かっているなら
'   それを分母にする。0(不明)なら従来どおり okN+failN。上限で打ち切られた
'   資料では okN+failN は「読もうとした頁」でしかなく、254頁の資料で
'   「全200頁中3頁が…」と名乗ると、残り54頁の存在が消える。
Public Function OcrPartialMemoFor(ByVal okN As Long, ByVal failN As Long, _
                                  Optional ByVal totalKnown As Long = 0) As String
    If failN <= 0 Then Exit Function
    Dim k As Long: k = okN
    If k < 0 Then k = 0
    Dim f As Long: f = failN
    Dim n As Long: n = k + f
    If totalKnown > n Then n = totalKnown
    OcrPartialMemoFor = "全" & n & "頁中" & f & _
        "頁が読み取れませんでした(もう一度取り込むと再試行します)"
End Function

' ----------------------------------------------------------------------------
' ComposeOcrMemo - 本棚カードへ出すメモの組み立てを1箇所に閉じる
'   (2026-08-04 R15-FixA FA-4・レビューA-H4/B-M)。
'   resumeHead  : 「前回の続きから再開しました。」(復元が無ければ "")
'   partialBody : 中断・上限・頁欠けの理由(OcrAbortMemoFor/OcrPartialMemoFor)
'   capBody     : 設定上限で打ち切ったことの説明(OcrCapMemoFor)
'
'   守る契約は2つだけ:
'     ・resumeHead は【必ず前置き】。あとから来たメモで置き換えない。
'     ・partialBody と capBody は【両方あれば連結】する。片方が片方を
'       消してはならない。
'   従来はこの3つが2箇所(optOcrPage と optVision.OcrCapMemo)で別々に
'   組み立てられ、しかも後から来たものが前のものを丸ごと【置換】していた。
'   その結果、上限で打ち切られた資料の一部の頁が読めなかった場合、
'   カードには「頁が欠けた」か「上限で切った」のどちらか一方しか出ず、
'   もう一方の事実は誰にも届かないまま消えていた(憲章§4-1)。
'   復元の事実に至っては、欠けも打ち切りも無い正常な再開取込では
'   truncated=False のためカードへ渡す経路自体が無かった。
' ----------------------------------------------------------------------------
Public Function ComposeOcrMemo(ByVal resumeHead As String, ByVal partialBody As String, _
                               ByVal capBody As String) As String
    Dim body As String: body = Trim$(partialBody)
    Dim cap As String: cap = Trim$(capBody)

    If LenB(body) = 0 Then
        body = cap
    ElseIf LenB(cap) > 0 Then
        body = body & "。" & cap
    End If

    ComposeOcrMemo = Trim$(resumeHead) & body
End Function

' ----------------------------------------------------------------------------
' RemainingWaitSec - 1資料あたりの絶対上限(config gs_abs_timeout_sec)の
'   【残り】を返す(2026-08-03 R14-F6。R15-5bで optOcrCore から移設)。
'   0=もう待てない(打ち切り)。バッチ描画で絶対上限が「1バッチあたり」の
'   意味になっていた(20頁×6バッチなら最悪1200秒×6)。資料単位の壁へ戻す。
'   残りが短くても10秒は待つ(1秒の待ちは必ず時間切れになるだけで得が無い)。
' ----------------------------------------------------------------------------
Public Function RemainingWaitSec(ByVal absSec As Long, ByVal usedSec As Long) As Long
    If absSec < 1 Then Exit Function
    If usedSec >= absSec Then Exit Function
    Dim r As Long: r = absSec - usedSec
    If r < 10 Then r = 10
    RemainingWaitSec = r
End Function

' ----------------------------------------------------------------------------
' BatchWaitSec - 1バッチの画像化に待ってよい秒数(2026-08-04 R15-FixA FA-5iii)。
'   = min(資料あたりの残り予算, 1バッチの上限 max(120, batchPages×16))
'   0 = もう待てない(呼び出し元は打ち切る)。
'   なぜ1バッチにも上限が要るか: 残り予算をそのまま1バッチへ渡していたため、
'   254頁の資料(予算2032秒)の1バッチ目でGSがハングすると、そこで34分待ち
'   続けて他の12バッチには1秒も残らない。頁数から素直に見積もれば20頁は
'   長くても数分なので、1バッチ320秒(20×16)を超えたら、そのバッチだけを
'   時間切れにして先へ進む方が、資料としては多くの頁が残る。
'   下限120秒は「1頁でも描けないほど短い待ち」を作らないため。
' ----------------------------------------------------------------------------
Public Function BatchWaitSec(ByVal absSec As Long, ByVal usedSec As Long, _
                             ByVal batchPages As Long) As Long
    Dim rest As Long: rest = RemainingWaitSec(absSec, usedSec)
    If rest <= 0 Then Exit Function

    Dim capSec As Long: capSec = 120
    If batchPages > 0 And batchPages < 100000 Then
        If batchPages * 16& > capSec Then capSec = batchPages * 16&
    End If

    If rest > capSec Then rest = capSec
    BatchWaitSec = rest
End Function

' ----------------------------------------------------------------------------
' OcrEstMinutes - これから読む頁数から所要時間(分)を見積もる
'   (2026-08-04 R15-7b・実機第4報 RC4/RC7)。
'   freshPages   : 【これから実際にOCRする】頁数。前回の続きから再開できる頁
'                  (キャッシュ済み)は数えない=待ち時間はかからない。
'   avgMsPerPage : 1頁あたりの実測(modState "ocr_avg_page_ms")。
'                  0以下(=この端末でまだ1度も読んでいない)なら DEFAULT_PAGE_MS。
'   四捨五入で最低1分。0頁なら0分(=確認そのものを出さない材料になる)。
' ----------------------------------------------------------------------------
Public Function OcrEstMinutes(ByVal freshPages As Long, ByVal avgMsPerPage As Double) As Long
    If freshPages <= 0 Then Exit Function
    Dim msPer As Double: msPer = avgMsPerPage
    If msPer <= 0# Then msPer = DEFAULT_PAGE_MS
    Dim m As Long
    m = CLng(Int((CDbl(freshPages) * msPer) / 60000# + 0.5))
    If m < 1 Then m = 1
    OcrEstMinutes = m
End Function

' ----------------------------------------------------------------------------
' OcrConfirmAskFor - 取り込む前に一度だけ聞く確認文(2026-08-04 R15-7b)。
'   確認が要らないときは ""(呼び出し元はダイアログを出さない)。
'   totalPages : 資料の総頁数(0=不明なら何も聞かない。嘘の頁数で驚かせない)
'   estMin     : OcrEstMinutes の見積もり(分)
'   minMinutes : config ocr_confirm_min_minutes。この分数以上かかるときだけ聞く。
'                0以下は「確認しない」(この製品の 0=無効 の慣習に合わせる)。
'   なぜ聞くのか: 254頁のスキャンPDFは1〜2時間かかる。何も言わずに始めると、
'   利用者は「フリーズした」と判断してExcelを強制終了する(実機第4報 RC7で
'   実際に起きた)。かかる時間と、途中で止められること・続きから再開できる
'   ことを先に伝えれば、待つか後回しにするかを利用者が選べる(憲章§3-2)。
'   R17H FA-10(B-H3): 「検索用の準備」だけでは、テキストPDFでも取込後に
'   章の要約(章数ぶんのLLM呼び出し)で数分かかることが伝わらない。何分続くのか
'   分からない待ちは「終わったはずなのに固まっている」と読まれる(憲章§3-2)。
'   R15-FixB(FB-12): 中断ボタンの見た目(■ U+25A0)は【CP932 にある】文字で、
'   CP932 で書けないから ChrW にしているのではない(旧コメントの事実誤り)。
'   ChrW で書くのは、UTF-8のソースをCP932の実行文へ変換して注入するビルドの
'   往復で、この1文字が化けても意図が読めなくならないようにするため。
' ----------------------------------------------------------------------------
Public Function OcrConfirmAskFor(ByVal totalPages As Long, ByVal estMin As Long, _
                                 ByVal minMinutes As Long) As String
    If totalPages <= 0 Then Exit Function
    If minMinutes <= 0 Then Exit Function
    If estMin < minMinutes Then Exit Function

    OcrConfirmAskFor = "全" & totalPages & "頁のスキャンPDFです。" & _
        "読み取りに推定約" & estMin & "分かかります。" & vbLf & _
        "処理中も" & ChrW(&H25A0) & "中断で止められ、次回は続きから再開できます。" & vbLf & _
        "処理中このExcelは操作できません。他の仕事はバナー内の" & _
        "「作業用Excel」ボタンからどうぞ。" & vbLf & _
        "(完了後に検索用の準備と章の要約が続きます。数分かかることがあります)" & vbLf & _
        vbLf & "取り込みますか?"
End Function

' ----------------------------------------------------------------------------
' RenderWaitBanner - 画像化(GS描画)の待ちに出す実況(R15-FixA FA-5i)。
'   「画像化中… バッチ3/13 (全254頁) 経過42秒」
'   バッチのラベルは BatchLabel が作る(総頁が未確定なら「バッチ3」だけ)。
'   経過秒を出すのは、この待ちがVisionの応答待ちと違って【画面が完全に
'   止まって見える】区間だから。数字が1秒ごとに動いていれば、止まって
'   いないことだけは分かる(憲章§3-2)。0以下の経過秒は出さない。
' ----------------------------------------------------------------------------
Public Function RenderWaitBanner(ByVal labelText As String, ByVal elapsedSec As Long) As String
    RenderWaitBanner = "画像化中… " & labelText
    If elapsedSec > 0 Then _
        RenderWaitBanner = RenderWaitBanner & " 経過" & elapsedSec & "秒"
End Function

' ----------------------------------------------------------------------------
' OcrDeclineMemoFor - 確認で「いいえ」を選んだ資料に残すメモ(R15-7b)。
'   見送りは失敗ではないが、本棚に何も無いまま黙って消えるのが一番困る。
'   「なぜ入っていないのか」と「どうすれば入るのか」を1文で言う。
'   R15-FixB(FB-5): 文の【先頭】を分数に依らない固定句にする。見送りを失敗と
'   分けて数えるには「このメモは見送りか」を機械で判定する必要があり、判定は
'   modUtilText.IsDeclineNote(先頭一致)が1本だけ持つ。同じ文字列を2箇所に
'   書かないため、先頭句そのものは modUtilText.DECLINE_MEMO_HEAD から借りる
'   (opt層とコア層の両方から見えるのは基盤層だけ。文面の残りはここが一次情報)。
' ----------------------------------------------------------------------------
Public Function OcrDeclineMemoFor(ByVal estMin As Long) As String
    Dim m As Long: m = estMin
    If m < 1 Then m = 1
    OcrDeclineMemoFor = modUtilText.DECLINE_MEMO_HEAD & "(推定約" & m & "分)。" & _
        "再度取り込むと実行します"
End Function

' ----------------------------------------------------------------------------
' ClampPageMs - 1頁あたり所要ミリ秒の過去実績を、常識の範囲へ丸める
'   (2026-08-04 R15-FixB FB-4・レビューB-M)。0以下(=実測なし)は0のまま。
'   なぜ要るのか: この値は ui_state に永続化され、次の資料のETAの土台になる。
'   バッチ1つでも異常な値(GSのハングを含んだ描画時間、時計の巻き戻し)が
'   混ざると、以後ずっと「残り約3秒」や「残り約8時間」と言い続ける端末が
'   できる。人が見て意味のある下限3秒・上限120秒(=2分/頁)で頭打ちにする。
'   実機第4報の実測は1頁25秒前後で、この幅の内側に十分収まる。
' ----------------------------------------------------------------------------
Public Function ClampPageMs(ByVal ms As Double) As Double
    If ms <= 0# Then Exit Function
    ClampPageMs = ms
    If ClampPageMs < 3000# Then ClampPageMs = 3000#
    If ClampPageMs > 120000# Then ClampPageMs = 120000#
End Function

' ----------------------------------------------------------------------------
' GsBudgetSec - 画像化待ちの絶対上限を総頁数に連動させる(R15-7c・RC10)。
'   = max(baseSec, totalPages × 8秒)。totalPages が不明(0以下)なら
'   config の値(baseSec)をそのまま使う=従来動作。
'   config の gs_abs_timeout_sec は「1資料あたり」の絶対上限で、254頁の
'   資料でも既定1200秒しか無い。13バッチに配分すると後半は必ず残り10秒に
'   なり、描けているのに時間切れで partial になる(利用者からは「毎回
'   途中で止まる」としか見えない)。頁数が分かっているときだけ、頁数ぶんの
'   予算を自動で確保する。config を大きく設定してある端末ではそちらが勝つ。
'   Long のオーバーフロー(×8で溢れる)を避けるため、上げ幅は頭打ちにする。
' ----------------------------------------------------------------------------
Public Function GsBudgetSec(ByVal baseSec As Long, ByVal totalPages As Long) As Long
    GsBudgetSec = baseSec
    If GsBudgetSec < 0 Then GsBudgetSec = 0
    If totalPages <= 0 Then Exit Function

    Dim need As Long
    If totalPages > 1000000 Then
        need = 8000000
    Else
        need = totalPages * GS_SEC_PER_PAGE
    End If
    If need > GsBudgetSec Then GsBudgetSec = need
End Function

' ----------------------------------------------------------------------------
' OcrResumeMemo - 前回の続きから再開した取込のメモに、その事実を足す
'   (2026-08-04 R15-7d)。cachedN=0(復元なし)なら1文字も変えない。
'   利用者から見ると「同じ資料を取り込み直したのに、前より早く終わった」の
'   説明がここにしか無い。黙って速いのは、黙って遅いのと同じくらい不安。
' ----------------------------------------------------------------------------
Public Function OcrResumeMemo(ByVal baseMemo As String, ByVal cachedN As Long) As String
    OcrResumeMemo = baseMemo
    If cachedN <= 0 Then Exit Function
    OcrResumeMemo = "前回の続きから再開しました。" & baseMemo
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' nowAt から remSec 秒後の時刻を "HH:MM"(24時制・2桁ゼロ埋め)で返す。
' Format$(nowAt + remSec / 86400#, "hh:mm") と同じ結果を整数計算だけで作る
' (書式コードの解釈が環境で揺れないようにするため)。日を跨いだら折り返す。
' Hour() の戻り値は Integer なので、3600 を掛ける前に必ず Long へ広げる
' (23 * 3600 = 82800 は Integer の範囲を超えオーバーフローする)。
' R15-FixB(FB-11): 折り返したときは「翌 」を前置する。23:50 に始めた取込の
' 終了目安が「00:40」とだけ出ると、40分後ではなく【もう過ぎた時刻】に読める。
' 折り返していない普通のケースでは1文字も足さない(判定はここで完結し、
' 呼び出し元の RemainingText は今までどおり戻り値をそのまま使う)。
Private Function EndClock(ByVal nowAt As Date, ByVal remSec As Long) As String
    Dim t As Long
    t = CLng(Hour(nowAt)) * 3600& + CLng(Minute(nowAt)) * 60& + CLng(Second(nowAt))
    t = (t + remSec) \ 60&
    Dim dayShift As Long: dayShift = t \ 1440&
    t = t - dayShift * 1440&
    If t < 0 Then
        t = t + 1440&
        dayShift = dayShift - 1
    End If
    EndClock = Pad2(t \ 60&) & ":" & Pad2(t - (t \ 60&) * 60&)
    If dayShift > 0 Then EndClock = "翌 " & EndClock
End Function

Private Function Pad2(ByVal n As Long) As String
    Pad2 = Right$("0" & CStr(n), 2)
End Function
