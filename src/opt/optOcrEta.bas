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
        OcrPageBanner = "OCR中… " & p & "頁目 (バッチ " & batchIdx & ")"
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

    OcrPageBanner = s
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
        OcrAbortMemoFor = "利用者の操作で中断しました。ここまでの" & k & _
            "頁を保存しました(" & again & ")"
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
Public Function OcrPartialMemoFor(ByVal okN As Long, ByVal failN As Long) As String
    If failN <= 0 Then Exit Function
    Dim k As Long: k = okN
    If k < 0 Then k = 0
    Dim f As Long: f = failN
    OcrPartialMemoFor = "全" & (k + f) & "頁中" & f & _
        "頁が読み取れませんでした(もう一度取り込むと再試行します)"
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
' 内部ヘルパー
' ----------------------------------------------------------------------------

' nowAt から remSec 秒後の時刻を "HH:MM"(24時制・2桁ゼロ埋め)で返す。
' Format$(nowAt + remSec / 86400#, "hh:mm") と同じ結果を整数計算だけで作る
' (書式コードの解釈が環境で揺れないようにするため)。日を跨いだら折り返す。
' Hour() の戻り値は Integer なので、3600 を掛ける前に必ず Long へ広げる
' (23 * 3600 = 82800 は Integer の範囲を超えオーバーフローする)。
Private Function EndClock(ByVal nowAt As Date, ByVal remSec As Long) As String
    Dim t As Long
    t = CLng(Hour(nowAt)) * 3600& + CLng(Minute(nowAt)) * 60& + CLng(Second(nowAt))
    t = (t + remSec) \ 60&
    t = t - (t \ 1440&) * 1440&
    If t < 0 Then t = t + 1440&
    EndClock = Pad2(t \ 60&) & ":" & Pad2(t - (t \ 60&) * 60&)
End Function

Private Function Pad2(ByVal n As Long) As String
    Pad2 = Right$("0" & CStr(n), 2)
End Function
