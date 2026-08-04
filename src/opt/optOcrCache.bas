Attribute VB_Name = "optOcrCache"
Option Explicit

' ============================================================================
' optOcrCache - 画像PDF OCRの頁チェックポイント(opt機能・2026-08-04 R15-7d)
' ----------------------------------------------------------------------------
' 役割:
'   1頁読むごとに、その頁の本文を隠しシート "ocr_cache" へ控えておき、
'   次に同じ資料を取り込んだときは【読めている頁を読み直さない】。
'
' なぜ要るのか(実機第4報 RC7):
'   Vision(ChatGPTV)は同期呼び出しで、1頁の応答が返らない限りVBAからは
'   何もできない。実機では254頁のスキャンPDFの途中でハングし、利用者は
'   Excelを強制終了するしか無かった。ハングそのものは【こちらから制御
'   できない】(RIBBON_API_CONFIRMED.md:29 に待ち秒数の引数が無い)ので、
'   打てる手は「被害を限定する」ことだけになる。読めた頁が次回まで残れば、
'   200頁目で落ちても失うのは1頁ぶんで、もう一度取り込めば続きから進む。
'   中断ボタン(R15-6)で自分から止めた場合も同じ理屈で続きから再開できる。
'
' 設計判断:
'   ・キャッシュは【常に任意の高速化】。シートが作れない・書けない・壊れて
'     いる、のどれが起きても取込は最後まで走り切る(全損しない)。書込みに
'     失敗したことは usage_log "ocr_cache_fail" に1行だけ残して先へ進む。
'   ・鍵 = Fnv1a64Hex(元ファイルのフルパス) & "|" & FileLen & "|" &
'     FileDateTime & "|p" & 頁番号。サイズか更新日時が1つでも違えば別の鍵に
'     なる=【中身が差し替わった資料に古い本文を混ぜない】。同名で上書き
'     更新される規程・料金表はこの製品の主対象なので、ここは絶対に譲れない。
'     鍵に使うのは【元のパス】で、取込が実際に読む一時コピー(mbtmp_*)では
'     ない(コピー先の名前と日時は取込のたびに変わり、二度と一致しない)。
'   ・本文は先頭に番兵1字("t")を付けて書き、読むときに剥ぐ(R15-FixB FB-3)。
'     本文が "=" や "+" で始まると、セルはそれを【数式】として解釈する。
'     取り込んだ本文が数式に化ければ #NAME? が並び、読み直したときに元の
'     文字は二度と戻らない(黙って中身が変わる=憲章§4-1の最悪形)。番兵が
'     あれば必ず文字列として入り、剥げば完全に元へ戻る。さらに「本文が空の
'     頁」("t"だけの行)と「まだ読んでいない頁」(行が無い)を区別できる。
'   ・本文は modUtil.SafeLeft(text, 32000)。セルの上限は32,767字で、
'     サロゲートペアの途中で切ると壊れた文字が残る(SafeLeftはそこも見る)。
'     ただし【切り詰めが起きる頁は控えない】(FB-10): 途中で切れた本文を
'     次回そのまま復元すると、読み直せば全部入るはずの頁が永久に欠ける。
'   ・書き込みは【バッチ単位で1回】(1行ずつのFind/書込みはしない)。
'     20頁ぶんを1つの配列にして1回のRange代入で置く(§12の一括I/O規約)。
'   ・読み出しも1回。シート全体を配列で受け取り、当該資料の鍵で始まる行だけ
'     をメモリへ拾う。復元した本文はメモリからすぐ捨てる(pages()に移った
'     あとまで二重に持たない)。
'   ・後始末は2段: 取込が【本棚に done として並んだ】ら、その資料の行を全部
'     消す(R15-FixB FB-1: 判断できるのはコア層の modShelf だけ。OCRが読み
'     切っても、そのあとのチャンク化・ベクトル化で落ちれば資料は partial の
'     まま残り、控えはまだ要る)。それ以外の理由で残った孤児行は、起動時GC
'     (modBoot)から2日で消す。放っておくとブックが太り続ける(憲章§3-5)。
'   ・opt層でシートを触るのはこのモジュールだけ。作法(ThisWorkbookからの
'     取得・xlSheetVeryHidden・見出し3列)は modEmbed.EnsureVectorSheet を
'     そのまま踏襲する。opt機能を撤去するときは、このシートが残っても
'     コアは1行も壊れない(誰も読まない veryHidden シートが残るだけ)。
' ============================================================================

Private Const SH_OCR_CACHE As String = "ocr_cache"

' セル1つの上限は32,767字。SafeLeft で切る幅をここ1箇所に持つ。
Private Const TEXT_MAX As Long = 32000

' 孤児行(取込が完走しないまま残った行)を起動時GCで消すまでの日数。
' R15-FixB(FB-1): 7日 → 2日。1頁32,000字×300頁で1資料あたり最大約10MBに
' なり得るのに、7日ぶん溜め込む理由が無い。中断した取込を再開するのは
' 「思い出したその日か翌日」で、2日あれば実利用は十分に覆える。
Private Const CACHE_KEEP_DAYS As Long = 2

' 本文セルの先頭に必ず付ける番兵1字(R15-FixB FB-3)。"=" 始まりの本文が
' 数式として解釈されるのを防ぎ、読み出しで剥げば完全に元へ戻る。
Private Const TEXT_SENTINEL As String = "t"

' 1頁あたり所要ミリ秒の過去実績(ui_state)。事前確認の見積もりに使う。
' 【optOcrPage.RATE_KEY と必ず対で直すこと】(同じキーを2箇所で読む)。
Private Const RATE_KEY As String = "ocr_avg_page_ms"

' 列(見出し行は1行目固定)。
Private Const COL_KEY As Long = 1
Private Const COL_TEXT As Long = 2
Private Const COL_SAVED As Long = 3

' 頁の配列を伸ばす単位。
Private Const GROW_N As Long = 64

' 取込1件のあいだだけ持つ状態(次の資料に持ち越さない。BeginDocで必ず初期化)。
Private mDocPrefix As String      ' 鍵の頁番号より前の部分("" = 鍵を作れない)
Private mPageNos() As Long        ' 既知の頁番号(復元できる頁+今回書いた頁)
Private mTexts() As String        ' 復元用の本文(復元したら空にして手放す)
Private mN As Long                ' mPageNos に入っている件数
Private mCachedN As Long          ' そのうち前回から復元できた頁数(記録用に不変)
Private mSavedN As Long           ' 今回の取込で書けた頁数
Private mFailLogged As Boolean    ' 書込み失敗の記録は資料ごとに1行だけ
Private mLastEstMin As Long       ' 直近の事前確認で見積もった分数(拒否メモ用)

Public Function Ping() As Boolean
    Ping = True
End Function

' ----------------------------------------------------------------------------
' DocPrefixFor / CacheKeyFor - 鍵の組み立て(純ロジック・テストで固定する)。
'   DocPrefixFor : 資料1件を指す部分。ハッシュ|サイズ|更新日時。
'   CacheKeyFor  : そこへ頁番号を足した1行ぶんの鍵。
'   パス・サイズ・更新日時のどれか1つでも違えば別の鍵になり、古い本文は
'   二度と使われない(差し替え検知)。ハッシュを使うのはパスが長くても
'   鍵が短く収まるようにするため(セルに入れる文字列なので短いほど良い)。
' ----------------------------------------------------------------------------
Public Function DocPrefixFor(ByVal pdfPath As String, ByVal fileSize As Long, _
                             ByVal stampText As String) As String
    DocPrefixFor = modUtil.Fnv1a64Hex(pdfPath) & "|" & fileSize & "|" & stampText
End Function

Public Function CacheKeyFor(ByVal pdfPath As String, ByVal fileSize As Long, _
                            ByVal stampText As String, ByVal pageNo As Long) As String
    CacheKeyFor = DocPrefixFor(pdfPath, fileSize, stampText) & "|p" & pageNo
End Function

' セルへ入れてよい長さに切る(純ロジック。切り幅を2箇所に書かないための1本)。
Public Function CacheTextFor(ByVal s As String) As String
    CacheTextFor = modUtil.SafeLeft(s, TEXT_MAX)
End Function

' ----------------------------------------------------------------------------
' BeginDoc - これから取り込む資料のキャッシュを読み込む(戻り値=復元できる頁数)。
'   pdfPath は【元のファイル】のフルパス(一時コピーではない)。
'   鍵が作れない(ファイルが消えた等)・シートが無い・読めない場合は0を返し、
'   以後この取込ではキャッシュを一切使わない=従来どおりの取込になる。
' ----------------------------------------------------------------------------
Public Function BeginDoc(ByVal pdfPath As String) As Long
    ResetDoc
    mDocPrefix = DocPrefixNow(pdfPath)
    If LenB(mDocPrefix) = 0 Then Exit Function
    LoadDoc
    BeginDoc = mCachedN
End Function

' ----------------------------------------------------------------------------
' CachedText - この頁の本文が前回から残っているか。
'   outHit : True=前回の控えがあった(本文が空でも True)。
'   一度返した本文はメモリから手放す(呼び出し元の pages() へ移ったあとまで
'   二重に持たない。300頁×32,000字を二重に抱えると32bit Excelでは効く)。
'   頁番号の管理(HasPage)は残るので、そのあとの重複書込み防止は効き続ける。
'   R15-FixB(FB-3): 戻り値だけでは「控えがあって中身が空だった頁」と
'   「まだ読んでいない頁」を区別できず、前者を毎回読み直していた(白紙頁・
'   Visionが1文字も返さなかった頁は、何度取り込んでも必ず読み直しになる)。
'   在/ 不在は outHit だけが答える。
' ----------------------------------------------------------------------------
Public Function CachedText(ByVal pageNo As Long, ByRef outHit As Boolean) As String
    outHit = False
    Dim i As Long
    For i = 0 To mN - 1
        If mPageNos(i) = pageNo Then
            CachedText = mTexts(i)
            mTexts(i) = ""
            outHit = True
            Exit Function
        End If
    Next i
End Function

' ----------------------------------------------------------------------------
' SaveRange - 直前のバッチで読めた頁をまとめて書き込む(R15-7d)。
'   pages(fromIdx..toIdx) のうち、まだシートに無い頁だけを1回のRange代入で
'   追記する(復元しただけの頁は既に載っているので書かない)。
'   失敗しても取込は止めない。usage_log "ocr_cache_fail" を資料ごとに1行だけ
'   残す(毎バッチ書くと、本当の失敗が行数ローテで流れる)。
' ----------------------------------------------------------------------------
Public Sub SaveRange(ByRef pages() As ExtractedPage, ByVal fromIdx As Long, _
                     ByVal toIdx As Long)
    If LenB(mDocPrefix) = 0 Then Exit Sub
    If toIdx < fromIdx Then Exit Sub

    ' シートが作れないときは Fail へ飛ばさない。エラーは起きていないので
    ' そこの Resume が実行時エラー20(Resume にエラーがありません)になる。
    Dim ws As Worksheet: Set ws = EnsureCacheSheet()
    If ws Is Nothing Then
        NoteCacheFail "保存用のシートを用意できませんでした"
        Exit Sub
    End If

    On Error GoTo Fail
    ' 1周目: 何行書くかを数える(配列の大きさを行数ぴったりにするため)。
    Dim n As Long: n = 0
    Dim i As Long
    For i = fromIdx To toIdx
        If Savable(pages, i) Then n = n + 1
    Next i
    If n = 0 Then Exit Sub

    Dim buf() As Variant
    ReDim buf(1 To n, 1 To 3)
    Dim savedAt As Date: savedAt = Now
    Dim k As Long: k = 0
    For i = fromIdx To toIdx
        If Savable(pages, i) Then
            Dim p As Long: p = pages(i).page
            k = k + 1
            buf(k, COL_KEY) = mDocPrefix & "|p" & p
            ' R15-FixB(FB-3): 番兵つきで書く(読み出しで剥ぐ)。
            buf(k, COL_TEXT) = TEXT_SENTINEL & CacheTextFor(pages(i).Text)
            buf(k, COL_SAVED) = savedAt
            NotePage p, ""      ' 以後の重複書込みを防ぐ(本文は持たない)
        End If
    Next i

    Dim lastRow As Long: lastRow = LastRowOf(ws)
    ws.Range(ws.Cells(lastRow + 1, COL_KEY), ws.Cells(lastRow + n, COL_SAVED)).Value = buf
    mSavedN = mSavedN + n

    ' R15-FixA(FA-2・レビューA-H2/B-H2): ここまでで控えは【メモリ上のブック】に
    ' しか無い。Excelが落ちれば、あるいは強制終了されれば、シートへ書いた行ごと
    ' 消える=「続きから再開できる」という約束(R15-7d)がその瞬間だけ嘘になる。
    ' 実機第4報で実際に起きたのは、まさにその強制終了だった。書けた直後に
    ' ディスクへ落とす。ただし毎バッチ保存すると数百KBの書き戻しが積み上がる
    ' 端末があるので120秒のスロットルを掛ける(失うのは最大2分ぶん)。
    modShelfBatch.SaveCheckpoint n, 120
    Exit Sub

Fail:
    Dim failDesc As String: failDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かない。Resume で抜けてから
    ' 記録する(opt層GS系モジュール共通の作法)。
    Resume SaveFailed
SaveFailed:
    NoteCacheFail failDesc
End Sub

' ----------------------------------------------------------------------------
' Savable - この頁を控えてよいか(R15-FixB FB-10)。
'   頁番号が正で / まだ書いていなくて / セルに丸ごと収まる本文であること。
'   32,000字を超える頁を SafeLeft で切って控えると、次の取込では【切れた
'   ままの本文】が復元され、読み直せば全部入るはずの後半が永久に失われる。
'   控えないでおけば次回もう一度読むだけ(遅くなるだけで、欠けない)。
' ----------------------------------------------------------------------------
Private Function Savable(ByRef pages() As ExtractedPage, ByVal i As Long) As Boolean
    If pages(i).page <= 0 Then Exit Function
    If HasPage(pages(i).page) Then Exit Function
    Savable = (Len(pages(i).Text) <= TEXT_MAX)
End Function

' 書込み失敗を usage_log へ1行だけ残す(資料ごとに1回。取込は止めない)。
Private Sub NoteCacheFail(ByVal why As String)
    If mFailLogged Then Exit Sub
    mFailLogged = True
    On Error Resume Next
    modLog.LogUsage "ocr_cache_fail", "", modUtil.SafeLeft( _
        "読み取り途中の保存に失敗しました(取込は続行): " & why, 300)
    On Error GoTo 0
End Sub

' 1頁でも「次回に続きから再開できる」状態になっているか(R15-7d)。
' 中断メモで「もう一度取り込むと続きから再開します」と書いてよいかの唯一の
' 根拠。書けていないのに書いたら嘘になる(憲章§4-1)。
Public Function HasSaved() As Boolean
    HasSaved = (mSavedN > 0 Or mCachedN > 0)
End Function

' ----------------------------------------------------------------------------
' PurgeFor - その資料の控えを全部消す(R15-7d。R15-FixB FB-1 で呼び時機を変更)。
'   pdfPath : 【元のファイル】のフルパス(鍵と同じもの)。
'   従来は OCR を読み切った時点(FinishDoc)で消していた。しかし読み切った
'   あとにもチャンク分割・ベクトル化・保存が控えており、そこで落ちれば資料は
'   partial のまま本棚に残る——控えだけ先に消えているので、次の取込は
'   【最初から】300頁を読み直すことになる(控えの存在意義そのものを失う)。
'   消してよいと言えるのは「本棚に done として並んだ」ことを知っている
'   コア層(modShelf)だけなので、そこから呼べる口をここに置く。
'   OCR を1度も通っていない資料(通常のテキストPDF・Word等)から呼ばれても、
'   その鍵で始まる行が1つも無いので RewriteKeeping は1セルも触らずに戻る。
'   戻り値は modFeatures.InvokeFeature 経由で呼ぶための契約合わせ(常に "")。
' ----------------------------------------------------------------------------
Public Function PurgeFor(ByVal pdfPath As String) As String
    On Error GoTo Quiet
    Dim pfx As String: pfx = DocPrefixNow(pdfPath)
    If LenB(pfx) = 0 Then Exit Function
    RewriteKeeping pfx & "|p", 0
    ' いま取り込み終えた資料そのものなら、メモリ側の控えも捨てる
    ' (次の資料へ持ち越さない。持ち越すと復元件数の記録がズレる)。
    If pfx = mDocPrefix Then ResetDoc
    Exit Function
Quiet:
    Exit Function          ' 掃除に失敗しても取込の結果は変えない
End Function

' ----------------------------------------------------------------------------
' FinishDoc - 取込1件ぶんの記録と、復元の事実のメモ付け(R15-7d)。
'   baseMemo : 呼び出し元(optOcrPage)が組み立てた正直なメモ
'   freshN   : 今回いくつの頁を実際に読んだか(usage_log の内訳用)
'   戻り値   : カードへ出すメモ。復元があった取込は冒頭に
'              「前回の続きから再開しました。」が付く。
'   R15-FixB(FB-1): 控えの削除(旧 okAll → PurgeDoc)はここから外した。
'   OCRが読み切ったことと、資料が本棚に done として並ぶことは別の事実で、
'   後者を知っているのは modShelf だけ(上の PurgeFor の注記)。
' ----------------------------------------------------------------------------
Public Function FinishDoc(ByVal baseMemo As String, ByVal freshN As Long) As String
    FinishDoc = baseMemo
    Dim cachedN As Long: cachedN = mCachedN
    If cachedN <= 0 Then Exit Function

    FinishDoc = optOcrEta.OcrResumeMemo(baseMemo, cachedN)
    On Error Resume Next
    modLog.LogUsage "ocr_resume", "", "cached=" & cachedN & ";fresh=" & freshN
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' GcOldRows - 孤児行の起動時GC(R15-7d)。保存から CACHE_KEEP_DAYS 日を超えた行を
'   消す。取込が完走しなかった資料の行は PurgeFor を通らないので、ここが
'   唯一の回収口になる。戻り値は modFeatures.InvokeFeature 経由で呼ぶための
'   契約合わせ(常に "")。日付として読めない行も消す(壊れた行を永久に
'   抱え続けない。消えて困るのは「次回が少し速い」だけ)。
' ----------------------------------------------------------------------------
Public Function GcOldRows() As String
    On Error Resume Next
    RewriteKeeping "", CACHE_KEEP_DAYS
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' ConfirmAskFor - 取込前の事前確認の文面(R15-7b)。確認不要なら ""。
'   総頁数は txtwrite の分類パスで既に読めている(optGsTxt.LastGsTotalPages)。
'   見積もりは【これから実際に読む頁数】(総頁 - 復元できる頁)×レートで出す。
'   全頁がキャッシュに揃っている資料は待ち時間がほぼ0なので何も聞かない
'   (R15-7d と 7b の整合)。
'   maxPages は config の上限(ここでは実際に効く値へ丸めてから使う)。
' ----------------------------------------------------------------------------
Public Function ConfirmAskFor(ByVal pdfPath As String, ByVal maxPages As Long) As String
    mLastEstMin = 0
    On Error GoTo Quiet

    Dim total As Long: total = optGsTxt.LastGsTotalPages()
    If total <= 0 Then Exit Function           ' 頁数が分からないなら聞かない

    Dim plan As Long: plan = total
    Dim cap As Long: cap = optOcrCore.SafeMaxPages(maxPages)
    If plan > cap Then plan = cap              ' 上限までしか読まない=そのぶんの時間

    Dim fresh As Long: fresh = plan - BeginDoc(pdfPath)
    If fresh < 0 Then fresh = 0

    mLastEstMin = optOcrEta.OcrEstMinutes(fresh, Val(modState.LoadState(RATE_KEY, "0")))
    ConfirmAskFor = optOcrEta.OcrConfirmAskFor(total, mLastEstMin, ConfirmMinMinutes())
    Exit Function

Quiet:
    ' ハンドラ稼働中は On Error Resume Next が効かない(NoteCacheFail の中で
    ' 転ぶと呼び出し元へ飛ぶ)。Resume で抜けてから記録する=この層の共通作法。
    Resume AskQuiet
AskQuiet:
    ConfirmAskFor = ""      ' 見積もれないときは従来どおり黙って取り込む
    ' R15-FixB(FB-3): 「聞かなかった」ことの理由は残す。ここが無言だと、
    ' 2時間の取込が確認なしで始まった原因(見積もり経路の故障)がどこにも
    ' 出ないまま、確認機能そのものが死んでいても気付けない。
    NoteCacheFail "取込前の見積もりに失敗しました"
End Function

' 直前の ConfirmAskFor の見積もりで作る「見送りました」のメモ(R15-7b)。
Public Function DeclineMemo() As String
    DeclineMemo = optOcrEta.OcrDeclineMemoFor(mLastEstMin)
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' config ocr_confirm_min_minutes(既定15分。0以下=確認しない)。
Private Function ConfirmMinMinutes() As Long
    ConfirmMinMinutes = modConfig.GetLong("ocr_confirm_min_minutes", 15)
End Function

Private Sub ResetDoc()
    mDocPrefix = ""
    mN = 0
    mCachedN = 0
    mSavedN = 0
    mFailLogged = False
    ReDim mPageNos(0 To GROW_N - 1)
    ReDim mTexts(0 To GROW_N - 1)
End Sub

' 元ファイルの実体(サイズ・更新日時)から鍵の前半を作る。読めなければ ""。
' R15-FixB(FB-8): 更新日時は modUtilText.IsoDateTime で固定書式にする。
' CStr(Date) はロケール・カレンダー設定(和暦)で表記が変わるため、和暦の
' 端末では鍵が "R8/08/04..." になり、設定を切り替えた瞬間に全ての控えが
' 別物になって二度と使われなくなる(そこまで気付ける手がかりも残らない)。
Private Function DocPrefixNow(ByVal pdfPath As String) As String
    On Error GoTo NoKey
    DocPrefixNow = DocPrefixFor(pdfPath, FileLen(pdfPath), _
        modUtilText.IsoDateTime(FileDateTime(pdfPath)))
    Exit Function
NoKey:
    DocPrefixNow = ""
End Function

' シートから当該資料の行だけをメモリへ拾う(1回の一括読み)。
Private Sub LoadDoc()
    On Error GoTo Quiet
    Dim ws As Worksheet: Set ws = GetSheet(SH_OCR_CACHE)
    If ws Is Nothing Then Exit Sub

    Dim arr As Variant
    If Not ReadAllRows(ws, arr) Then Exit Sub

    Dim tag As String: tag = mDocPrefix & "|p"
    Dim tagLen As Long: tagLen = Len(tag)
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim rowKey As String: rowKey = CStr(arr(i, COL_KEY))
        If Left$(rowKey, tagLen) = tag Then
            Dim p As Long: p = CLng(Val(Mid$(rowKey, tagLen + 1)))
            If p > 0 Then
                If Not HasPage(p) Then
                    NotePage p, BodyOf(CStr(arr(i, COL_TEXT)))
                    mCachedN = mCachedN + 1
                End If
            End If
        End If
    Next i
    Exit Sub

Quiet:
    Resume LoadQuiet        ' ハンドラを抜けてから記録する(共通作法)
LoadQuiet:
    ' 読めないキャッシュは【無いのと同じ】に倒す(取込は続く)。途中まで
    ' 拾えていた頁も捨てる: 復元はしたのに「復元0頁」と記録するズレを作らない。
    mN = 0
    mCachedN = 0
    ' R15-FixB(FB-3): 従来ここは完全な無言だった。控えが読めていないのに
    ' 何度取り込んでも毎回最初からになる理由が、利用者にも調査者にも1行も
    ' 残らない(黙って遅いのは、黙って壊れているのと見分けが付かない)。
    NoteCacheFail "前回の控えを読み出せませんでした"
End Sub

' 番兵を剥いだ本文(R15-FixB FB-3)。空セル(番兵ごと消えている壊れた行)は
' 空本文として扱う=行が在ることは事実なので、復元済みの頁として数える。
Private Function BodyOf(ByVal cellText As String) As String
    If Len(cellText) >= 1 Then BodyOf = Mid$(cellText, 2)
End Function

' 既知の頁か(復元できる頁+今回書いた頁)。
Private Function HasPage(ByVal pageNo As Long) As Boolean
    Dim i As Long
    For i = 0 To mN - 1
        If mPageNos(i) = pageNo Then
            HasPage = True
            Exit Function
        End If
    Next i
End Function

Private Sub NotePage(ByVal pageNo As Long, ByVal bodyText As String)
    If mN > UBound(mPageNos) Then
        ReDim Preserve mPageNos(0 To UBound(mPageNos) + GROW_N)
        ReDim Preserve mTexts(0 To UBound(mTexts) + GROW_N)
    End If
    mPageNos(mN) = pageNo
    mTexts(mN) = bodyText
    mN = mN + 1
End Sub

' ----------------------------------------------------------------------------
' RewriteKeeping - 「残す行だけ」を書き直す形の削除(R15-7d)。
'   dropPrefix   : この文字列で始まる鍵の行を消す("" なら鍵では消さない)
'   dropOlderDays: 保存からこの日数を超えた行を消す(0以下なら日数では消さない)
'   1行ずつ Delete すると、消す行数ぶんだけシートの再計算が走る(数百行で
'   目に見えて固まる)。残す行を配列に集め、まとめて置き直す方が速く、
'   途中で失敗しても「消えすぎる」壊れ方をしない。
' ----------------------------------------------------------------------------
Private Sub RewriteKeeping(ByVal dropPrefix As String, ByVal dropOlderDays As Long)
    On Error GoTo Quiet
    Dim ws As Worksheet: Set ws = GetSheet(SH_OCR_CACHE)
    If ws Is Nothing Then Exit Sub

    Dim arr As Variant
    If Not ReadAllRows(ws, arr) Then Exit Sub

    Dim total As Long: total = UBound(arr, 1) - LBound(arr, 1) + 1
    Dim keep() As Long: ReDim keep(0 To total - 1)
    Dim keepN As Long: keepN = 0
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If Not DropRow(arr, i, dropPrefix, dropOlderDays) Then
            keep(keepN) = i
            keepN = keepN + 1
        End If
    Next i
    If keepN = total Then Exit Sub          ' 消す行が無いなら1セルも触らない

    Dim lastRow As Long: lastRow = LastRowOf(ws)
    ws.Range(ws.Cells(2, COL_KEY), ws.Cells(lastRow, COL_SAVED)).ClearContents
    If keepN = 0 Then Exit Sub

    Dim buf() As Variant
    ReDim buf(1 To keepN, 1 To 3)
    Dim k As Long
    For k = 0 To keepN - 1
        buf(k + 1, COL_KEY) = arr(keep(k), COL_KEY)
        buf(k + 1, COL_TEXT) = arr(keep(k), COL_TEXT)
        buf(k + 1, COL_SAVED) = arr(keep(k), COL_SAVED)
    Next k
    ws.Range(ws.Cells(2, COL_KEY), ws.Cells(1 + keepN, COL_SAVED)).Value = buf
    Exit Sub

Quiet:
    Resume RewriteQuiet     ' ハンドラを抜けてから記録する(共通作法)
RewriteQuiet:
    ' 掃除に失敗しても取込・起動は続ける。ただし黙らない(R15-FixB FB-3):
    ' 消せない行が溜まり続けてブックが太ることに、誰も気付けなくなる(§4-1)。
    NoteCacheFail "控えの整理に失敗しました"
End Sub

' この行を消すか。鍵の一致(資料単位の削除)または古すぎる(孤児のGC)。
Private Function DropRow(ByRef arr As Variant, ByVal i As Long, _
                         ByVal dropPrefix As String, ByVal dropOlderDays As Long) As Boolean
    If LenB(dropPrefix) > 0 Then
        If Left$(CStr(arr(i, COL_KEY)), Len(dropPrefix)) = dropPrefix Then
            DropRow = True
            Exit Function
        End If
    End If
    If dropOlderDays <= 0 Then Exit Function

    Dim savedAt As Date
    On Error GoTo Broken
    savedAt = CDate(arr(i, COL_SAVED))
    On Error GoTo 0
    DropRow = (DateDiff("d", savedAt, Now) > dropOlderDays)
    Exit Function
Broken:
    ' 日付として読めない行は捨てる(壊れた行を永久に抱え続けない)。
    DropRow = True
End Function

' 見出しを除く全行を配列で受け取る(1行も無ければ False)。
Private Function ReadAllRows(ByVal ws As Worksheet, ByRef arr As Variant) As Boolean
    Dim lastRow As Long: lastRow = LastRowOf(ws)
    If lastRow < 2 Then Exit Function
    arr = ws.Range(ws.Cells(2, COL_KEY), ws.Cells(lastRow, COL_SAVED)).Value
    ReadAllRows = True
End Function

Private Function LastRowOf(ByVal ws As Worksheet) As Long
    LastRowOf = ws.Cells(ws.Rows.count, COL_KEY).End(xlUp).row
    If LastRowOf < 1 Then LastRowOf = 1
End Function

' modEmbed.EnsureVectorSheet と同じ作法(ThisWorkbook から取り、無ければ
' 末尾に作って veryHidden にする)。作れなければ Nothing(呼び出し元は諦める)。
' R15-FixB(FB-2): 通常このシートはビルドが最初から入れてあり(build の
' headers-only 生成)、ここを通るのは配布前の古いブックだけになった。それでも
' 経路は残す(消えていても取込が止まらない自己修復)。Worksheets.Add は
' 【追加したシートをアクティブにする】ので、取込の途中で画面が知らない
' シートへ飛び、進捗バナー(ActiveSheet に描く)も行き先を見失う。
' 追加の前後でアクティブシートを退避・復元する。復元の失敗は無視してよい
' (見えているシートが変わるだけで、控えも取込も壊れない)。
Private Function EnsureCacheSheet() As Worksheet
    Dim ws As Worksheet: Set ws = GetSheet(SH_OCR_CACHE)
    If ws Is Nothing Then
        Dim prevActive As Object: Set prevActive = Nothing
        On Error Resume Next
        Set prevActive = ActiveSheet
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = SH_OCR_CACHE
        ws.Cells(1, COL_KEY).Value = "key"
        ws.Cells(1, COL_TEXT).Value = "text"
        ws.Cells(1, COL_SAVED).Value = "saved_at"
        On Error Resume Next
        ws.Visible = 2   ' xlSheetVeryHidden
        If Not prevActive Is Nothing Then prevActive.Activate
        On Error GoTo 0
    End If
    Set EnsureCacheSheet = ws
    Exit Function
Fail:
    Set EnsureCacheSheet = Nothing
End Function

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function
