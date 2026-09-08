Attribute VB_Name = "optGsTxt"
Option Explicit

' ============================================================================
' optGsTxt - Ghostscript実行の共通道具 + テキストPDFのtxtwrite抽出
'            (opt機能・R10-3 / R10-3b)
' ----------------------------------------------------------------------------
' 役割:
'   (1) PDFの本文抽出の【第1選択】。同梱の gswin32c.exe に -sDEVICE=txtwrite
'       を渡し、PDFに埋まっている文字をそのまま書き出させて読む。COM
'       (Word/Acrobat)を一切使わないので、CreateObjectがポリシーで塞がれた
'       管理端末でも動く。入口は ExtractPdfTextNoOcr。
'   (2) Ghostscriptを動かすための共通道具(一時フォルダ・完了待ち・出力の
'       読み取り・後始末)の置き場。txtwrite経路と、optVisionの画像PDF OCR
'       経路(jpeg化)の両方がここを使う(R10-3bでoptVisionから移設)。
'       プロセスの起動と停止だけは optGsProc(R13-F2で容量のため分離)。
'
' なぜ第1選択なのか(2026-07-31 実機報告): 管理端末では Word/Acrobat の
'   CreateObject が塞がれテキストPDFが1本も入らず(実機初報B・E0302)、塞がれて
'   いない端末でもWordのPDF Reflowが「OLE操作の完了を待っています」を頻発させて
'   いた。GSは別プロセスで完了フラグ監視のタイムアウトが効き、数秒で返る。
'   駄目だったときだけ Word → Acrobat の既存連鎖へ落ちる
'   (modExtractor.ExtractPdfWithFallback)。
'
' 設計判断:
'   ・置き場所: optVision に容量が無く R10-3/R10-3b で分離(optGsTxt=GS実行、
'     optVision=Vision API+OCR統括、optGsProc=プロセス起動停止、
'     optOcrCore=コマンド組み立てと採否判定の純ロジック)。
'   ・入口は modFeatures.InvokeFeature("vision","ExtractPdfTextNoOcr",…)。
'     vision の行き先は optVision 固定なので、あちらの薄い受け口から転送される。
'   ・GS未検出でも案内カード(フォルダ選択)は出さない。ここは全PDF取込が毎回
'     通る道で、モーダルは取込のたびに割り込む。静かな候補探索
'     (optVision.FindGsExeByCandidates)だけを使い、無ければ黙ってWordへ譲る。
'     カードは利用者がOCRを期待している画像PDF経路だけの役目のまま。
'   ・mock_llm=TRUE でも動かす(txtwriteはLLMもAIリボンも使わない)。
'   ・戻り値の規約(optVisionの各関数と同じ "#ERR:" 文字列):
'       成功          : 抽出したテキスト(ページ区切りは改ページ文字のまま)
'       文字層ゼロ    : "#ERR:E0303:…" スキャンPDF確定。呼び出し元は即OCRへ
'       文字が薄い    : "#ERR:E0302:GS_SPARSE:…" 呼び出し元は Word → Acrobat を
'                       先に試し、両方失敗したときだけOCR(E0303)へ回す
'       タイムアウト  : "#ERR:E0302:GS_TIMEOUT:…" 途中まで読めていても採用しない
'       その他の失敗  : "#ERR:E0302:…" 呼び出し元は Word → Acrobat へ落ちる
'   ・失敗を握りつぶさない。WScript.Shell まで塞がれた端末があり得るので、
'     GS起動・出力読取で拾った Err.Number/Description は必ず err_log の detail
'     へ残す。ただし「GSが置かれていないだけ」は想定内かつ全PDFで毎回起きる
'     ので書かず、呼び出し元が3者併記のdetailでまとめて記録する。
' ============================================================================

Private Const OUT_TXT_NAME As String = "gstext.txt"
Private Const ERR_302 As String = "#ERR:E0302:"
Private Const ERR_303 As String = "#ERR:E0303:"

' R10c: 呼び出し元(modExtractor)が挙動を変える必要のある2つの失敗だけ、
' E0302系のうちに識別用トークンを立てる。どちらも「Word/Acrobatへ譲る」点は
' 通常のE0302と同じで、SPARSE だけが「Word/Acrobatも失敗したらE0303へ」と
' 最終コードを変える。文言の先頭に置くのでSafeLeftで落ちない。
Private Const TOKEN_TIMEOUT As String = "#ERR:E0302:GS_TIMEOUT:"
Private Const TOKEN_SPARSE As String = "#ERR:E0302:GS_SPARSE:"

' R13-1a: 完了フラグは「出来た瞬間はまだ空」でありうる(cmd.exe の echo は
' ファイル作成と書込みが別操作)。存在を見た後、中身が数字として読めるまで
' 100ms x 20 = 最大2秒だけ粘る。実機第2報の "rc=? flag=[]" はこのTOCTOU。
Private Const FLAG_RETRY_N As Long = 20

' R13-1c: 生存監視の刻み。待ちループは1秒ごとに回り(R13-M2)、2秒ごとに進捗
'(gs_out.log の "Page N" と gstext.txt のサイズ)を見て、1秒ごとにバナーを
' 更新する。数字を3箇所に散らさないためここに置く。
Private Const POLL_SEC As Double = 2#
Private Const BANNER_SEC As Double = 1#

' R13-F1: gs_out.log から "Processing pages 1 through N" を探す窓の大きさ。
' 300字では足りない。GSの起動バナー(版名・著作権・警告)だけで約210バイト
' あり、xref修復の警告が数行出ると Processing 行は窓の外へ押し出される。
' そうなると「ページ処理あり」の証拠を見失い、スキャンPDFが image ではなく
' gsfail に落ちてWord経路へ流れる(実機第2報 RC1 の再発)。
Private Const LOG_HEAD_CHARS As Long = 4000

' R15-5a(実機第4報 RC2): 直近の txtwrite 実行で判明した総ページ数(0=不明)。
' GSは処理の最初に "Processing pages 1 through N" を出すので、この分類パスを
' 通った時点で総ページ数は既に分かっている。ところが後続のOCR経路は自分で
' 描いた枚数からしか総頁を知る手段を持たず、最終バッチに入るまで分母もETAも
' 出せなかった(24頁中20頁・76頁中60頁が無表示)。ここへ置いておけば、
' 同じ資料のOCRが opt層内の参照1本で受け取れる。
Private mLastTotalPages As Long

Public Function Ping() As Boolean
    Ping = True
End Function

' ----------------------------------------------------------------------------
' LastGsTotalPages - 直近の ExtractPdfTextNoOcr で読めた総ページ数(R15-5a)。
'   読めなかった実行のあとは0。ExtractPdfTextNoOcr は入口で必ず0へ戻すので、
'   前の資料の値が次の資料へ持ち越されることはない(嘘の分母を出さない)。
'   PDFがOCR経路(E0303)へ回るのは、必ずこの関数の実行を通った【あと】である
'   (modExtractorPdf.ExtractPdfWithFallback が最初に txtwrite を試す)。
' ----------------------------------------------------------------------------
Public Function LastGsTotalPages() As Long
    LastGsTotalPages = mLastTotalPages
End Function

' ----------------------------------------------------------------------------
' ExtractPdfTextNoOcr - テキストPDFの本文をCOM無しで抜き出す(R10-3)。
'   path : PDFのフルパス(modExtractorが用意した一時コピーのパスが渡る)
'   戻り値: 本文 / "#ERR:E0303:…"(画像PDF疑い) / "#ERR:E0302:…"(その他失敗)
'
'   ・待ち方は ExtractPdfOcrPagedText と同じ規約: 非同期起動+完了フラグの
'     監視。壊れたPDF1つでExcelが永久に固まらないようにするため、GSの終了は
'     ブロッキングで待たない。2026-08-03(R13-1c)から待ち方は生存監視型で、
'     config vision_pdf_timeout_sec は「無進捗を許容する秒数(アイドル上限)」、
'     config gs_abs_timeout_sec(既定1200)が絶対上限。
'   ・一時フォルダは成功・失敗どちらの経路でも必ず片付ける。後始末は
'     R6規約に従って別Sub(CleanupTxtFolder)へ切り出してある(稼働中の
'     エラーハンドラの中では On Error Resume Next が効かないため)。
' ----------------------------------------------------------------------------
Public Function ExtractPdfTextNoOcr(ByVal path As String) As String
    Dim folderPath As String
    Dim outTxt As String
    Dim gsExe As String
    Dim flagPath As String
    Dim runCmd As String
    Dim waitSec As Long
    Dim gsErrNum As Long
    Dim gsErrDesc As String
    Dim finished As Boolean
    Dim txt As String
    Dim verdict As String

    folderPath = ""
    outTxt = ""
    gsErrNum = 0
    gsErrDesc = ""
    ' R15-5a: 新しい実行のたびに0へ戻す。ここで戻さないと、GSが動かなかった
    ' 資料のOCRが【前の資料の】総頁数を分母に出すことになる。
    mLastTotalPages = 0

    On Error GoTo Fail

    If modUtil.ExtOf(path) <> "pdf" Then
        ExtractPdfTextNoOcr = ERR_302 & "この処理はPDF専用です。"
        Exit Function
    End If

    ' 案内カードを出さない静かな解決(候補の実在確認だけ)。見つからなければ
    ' 黙って諦め、呼び出し元が Word → Acrobat の既存連鎖へ落ちる。
    gsExe = optVision.FindGsExeByCandidates()
    If LenB(gsExe) = 0 Then
        ExtractPdfTextNoOcr = ERR_302 & "Ghostscript が見つからないため、PDFの" & _
            "文字取り出し(COM不要の経路)は使えませんでした。"
        Exit Function
    End If

    folderPath = MakeOcrFolder()
    If LenB(folderPath) = 0 Then
        ExtractPdfTextNoOcr = ERR_302 & "作業用の一時フォルダを作成できませんでした。"
        Exit Function
    End If

    outTxt = folderPath & "\" & OUT_TXT_NAME
    flagPath = optOcrCore.DoneFlagFor(folderPath)
    runCmd = optOcrCore.BuildRunCommand( _
        optOcrCore.BuildGsTextCommand(gsExe, path, outTxt), flagPath, _
        optOcrCore.GsLogFor(folderPath))

    ' R13-1c: vision_pdf_timeout_sec の意味を「無進捗を許容する秒数(アイドル
    ' 上限)」へ変えた。ページが進んでいる限り待ち続け、何も進まなくなってから
    ' この秒数で見切る。暴走・ハングに備えた絶対上限が gs_abs_timeout_sec。
    waitSec = modConfig.GetLong("vision_pdf_timeout_sec", 120)
    If waitSec < 10 Then waitSec = 10
    Dim absSec As Long: absSec = modConfig.GetLong("gs_abs_timeout_sec", 1200)
    If absSec < waitSec Then absSec = waitSec

    modUIMain.SetStage "PDFの文字を取り出しています…"

    Dim gsPid As Long: gsPid = 0
    If Not optGsProc.RunGsAsync(runCmd, gsErrNum, gsErrDesc, gsPid) Then
        modUIMain.SetStage ""
        CleanupTxtFolder folderPath, outTxt
        ' R40 F1: Shell 直起動の失敗(実行ファイル無し・AppLocker 等)はここに出る。
        modLog.LogError "E0302", "optGsTxt.ExtractPdfTextNoOcr", modUtil.SafeLeft( _
            "GS起動失敗 err#" & gsErrNum & ": " & gsErrDesc & " " & path, 2000)
        ExtractPdfTextNoOcr = ERR_302 & "PDFから文字を取り出す処理を開始できませんでした。"
        Exit Function
    End If

    ' R13-1a/1c/1d: 完了フラグの中身まで確かめ、無進捗の秒数で見切り、
    ' 待っているあいだ1秒ごとにページ進捗をバナーへ出す待ち方。
    Dim gsRc As Long: gsRc = -1
    Dim pagesSeen As Long: pagesSeen = 0
    Dim totalPages As Long: totalPages = 0
    finished = WaitGsTextDone(folderPath, outTxt, waitSec, absSec, gsPid, _
                              gsRc, pagesSeen, totalPages)
    ' R15-5a: 総ページ数はこの1点でしか分からない。採否(ok/sparse/image)に
    ' 関わらず控える。画像PDFと分類された資料こそOCR経路がこれを必要とする。
    If totalPages > 0 Then mLastTotalPages = totalPages

    ' R10c(H1): タイムアウトは【失敗】。完了フラグが出ていない間、GSはまだ
    ' 出力txtを書いている途中なので、その時点で40字以上読めたとしても
    ' 「本文の途中まで」でしかない。採用すると資料が黙って欠けたまま本棚に
    ' 入り、以後の回答が虫食いの本文に基づくものになる(誰も気付けない)。
    ' Word→Acrobat連鎖へ譲るため E0302 系トークンで返す。
    ' R11-D(監査3 H-1): タイムアウト時は作業フォルダを【消さない】。
    ' 書きかけの gs_out.log と出力txtこそが「なぜ終わらないのか」の唯一の
    ' 手がかりで、消してしまうと二度と調べられない。残骸は次回起動時の
    ' GC(modBoot)が24時間後に片付けるので溜まり続けることはない。
    ' R13-1c: 「何秒で切れたか」だけでなく「どこまで進んでいたか」を必ず言う。
    ' 44頁中12頁で止まったのか、1頁も始まっていないのかで次の一手が変わる。
    If Not finished Then
        modUIMain.SetStage ""
        modLog.LogError "E0302", "optGsTxt.ExtractPdfTextNoOcr", modUtil.SafeLeft( _
            "GS_TIMEOUT idle=" & waitSec & "s abs=" & absSec & "s pages=" & pagesSeen & _
            "/" & totalPages & " pid=" & gsPid & " " & path & _
            " work=" & folderPath & " " & GsFailureDetail(folderPath), 2000)
        ExtractPdfTextNoOcr = TOKEN_TIMEOUT & "PDFからの文字取り出しが進まなくなりました" & _
            PagesSeenText(pagesSeen, totalPages) & "。" & waitSec & "秒間まったく進まなかったため中止しました。"
        Exit Function
    End If

    ' 2026-09-07(R39 F001): 単一gstext.txtの読み取りから、ページ別出力
    ' ("gstext_%04d.txt")を連結して読む形へ(BuildGsTextCommand側の変更に対応)。
    txt = optOcrCore.ReadPageFilesJoined(folderPath, totalPages, gsErrNum, gsErrDesc)

    ' R11-D(監査3 H-2): 終了コードとGS出力ログは【後始末の前】に読む。
    ' フォルダを消してから読もうとしても何も残っていない。
    ' R13-1a: 終了コードは待ちループが完了フラグの中身から読み終えている
    ' (gsRc。-1=2秒粘っても中身が空=flag_delayed)。
    Dim gsDetail As String
    Dim hasPages As Boolean: hasPages = (totalPages > 0)
    If gsRc <> 0 Or LenB(txt) = 0 Then
        gsDetail = GsFailureDetail(folderPath)
        If Not hasPages Then hasPages = GsOutHasPages(folderPath)
    End If

    modUIMain.SetStage ""
    CleanupTxtFolder folderPath, outTxt

    ' 先頭BOMの除去は modUtilText.ReadTextFileUtf8 が行う(2026-07-31 R11-F2で
    ' ここにあった個別対処を共通部品側へ集約した)。

    ' R13-1b(実機第2報 RC1): 「完了したのに本文が空」を一律のE0302にしない。
    ' 従来はここが設計済みの画像PDF判定(GsTextVerdict→E0303→OCR)より先に
    ' 発火し、スキャンPDFがWord経路へ流れてゴミ本文のまま「登録成功」になって
    ' いた。空のときは終了コードとGS出力から本当の理由を分類する。
    ' R39 レビュー B1: ページ別出力の連結は空ページにも Chr(12) を足すので LenB は
    ' 0 にならない。空白類だけ(=本文なし)を CleanTextLen で見る。
    If modUtilText.CleanTextLen(txt) = 0 Then
        Dim cls As String
        ' R14-3c: 入力サイズも材料に渡す。0バイトのPDFはGSが「空のPS」として
        ' rc=0+バナーだけで即終了するため、rcとログだけでは image と見分けが
        ' つかず、スキャンPDF扱いでOCRへ回って「描画0枚」に化けていた(RC3)。
        cls = optOcrCore.ClassifyGsTextResult(gsRc, 0, hasPages, InputFileBytes(path))

        If cls = "emptyinput" Then
            modLog.LogError "E0302", "optGsTxt.ExtractPdfTextNoOcr", modUtil.SafeLeft( _
                "empty_input: 0バイトの入力 " & path & " " & gsDetail, 2000)
            ExtractPdfTextNoOcr = ERR_302 & "empty_input:" & modLog.SharedReadFailMsg()
            Exit Function
        End If

        If cls = "image" Then
            ' GSは正常終了しページ処理も走った=文字層が無い。OCRへ回す。
            modLog.LogUsage "gs_txt_empty_image", "", modUtil.SafeLeft( _
                modUtil.FileNameOf(path) & " " & gsDetail, 500)
            ExtractPdfTextNoOcr = ERR_303 & "このPDFには文字データがほとんど入って" & _
                "いません(画像として保存されたPDFとみられます)。"
            Exit Function
        End If

        Dim clsNote As String: clsNote = " class=" & cls
        If cls = "flagdelay" Then clsNote = clsNote & " flag_delayed"
        modLog.LogError "E0302", "optGsTxt.ExtractPdfTextNoOcr", modUtil.SafeLeft( _
            "txtwrite出力を読めず rc=" & gsRc & clsNote & " err#" & gsErrNum & ": " & _
            gsErrDesc & " " & path & " " & gsDetail, 2000)
        ExtractPdfTextNoOcr = ERR_302 & "PDFから文字を取り出せませんでした" & _
            "(出力ファイルを読めませんでした)。"
        Exit Function
    End If

    ' 終了コードが非0でも本文が取れているなら採用する(GSは軽微な警告でも
    ' 非0を返すことがあり、ここで捨てると取り込めるPDFが取り込めなくなる)。
    ' ただし「非0なのに通っている」事実は次の調査の手がかりなので必ず残す。
    If gsRc > 0 Then
        modLog.LogUsage "gs_txt_rc_nonzero", "", modUtil.SafeLeft(gsDetail, 500)
    End If

    ' R10c(M1/M2): 採否は3値。文字層ゼロ(image)は即OCRへ、薄すぎる(sparse)は
    ' Word/Acrobatに先を譲り、両方失敗したときだけ呼び出し元がOCRへ回す。
    verdict = optOcrCore.GsTextVerdict(optOcrCore.CleanTextLen(txt), _
                                       optOcrCore.GsPageCount(txt))
    If verdict = "image" Then
        ExtractPdfTextNoOcr = ERR_303 & "このPDFには文字データがほとんど入って" & _
            "いません(画像として保存されたPDFとみられます)。"
        Exit Function
    ElseIf verdict = "sparse" Then
        ExtractPdfTextNoOcr = TOKEN_SPARSE & "このPDFの文字データはページ数に対して" & _
            "薄すぎます(" & optOcrCore.GsPageCount(txt) & "ページ)。"
        Exit Function
    End If

    modLog.LogUsage "pdf_text_gs", "", modUtil.SafeLeft(modUtil.FileNameOf(path), 120) & _
        " " & CStr(Len(txt)) & "字"
    ExtractPdfTextNoOcr = txt
    Exit Function

Fail:
    gsErrNum = Err.Number
    gsErrDesc = Err.Description
    Err.Clear
    On Error GoTo 0
    modUIMain.SetStage ""
    CleanupTxtFolder folderPath, outTxt
    modLog.LogError "E0302", "optGsTxt.ExtractPdfTextNoOcr", modUtil.SafeLeft( _
        "err#" & gsErrNum & ": " & gsErrDesc & " " & path, 2000)
    ExtractPdfTextNoOcr = ERR_302 & "PDFの文字取り出しでエラーが発生しました: " & gsErrDesc
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' 出力txtと一時フォルダの後始末。CleanupOcrFolder は *.jpg と
' *.flag しか消さない(OCR経路の規約)ので、txtは先に自分で消してから
' フォルダごと片付ける。ロック中で消せなくても無視する(TEMPなのでOSが
' 後で片付ける)。
Private Sub CleanupTxtFolder(ByVal folderPath As String, ByVal outTxt As String)
    If LenB(folderPath) = 0 Then Exit Sub
    On Error Resume Next
    If LenB(outTxt) > 0 Then Kill outTxt
    Kill folderPath & "\gstext_*.txt"   ' R39 F001: ページ別出力の後始末
    On Error GoTo 0
    CleanupOcrFolder folderPath
End Sub

' 2026-09-07(R39 F001): ReadUtf8Text(旧・単一gstext.txtの読み取り)は削除。
' 読み取りは optOcrCore.ReadPageFilesJoined(内部で modUtilText.ReadTextFileUtf8
' をページごとに呼ぶ)へ一本化した。

' 「(44ページ中12ページまで)」。総ページ数が読めていないときは分母を出さない。
' 何も分からないときは空文字(嘘の数字を出すくらいなら黙る)。
Private Function PagesSeenText(ByVal pagesSeen As Long, ByVal totalPages As Long) As String
    If totalPages > 0 Then
        PagesSeenText = "(" & totalPages & "ページ中" & pagesSeen & "ページまで)"
    ElseIf pagesSeen > 0 Then
        PagesSeenText = "(" & pagesSeen & "ページまで)"
    End If
End Function

' gs_out.log の先頭に "Processing pages 1 through N" があったか(R13-1b)。
' 「GSがPDFを開いてページ処理まで進んだ」ことの唯一の証拠で、これが無いのに
' 出力が空なら文字層の有無ではなく実行そのものの失敗(EDRブロック等)。
' 読む窓は LOG_HEAD_CHARS(300字では足りない。理由は定数のコメント参照)。
' R15-5a: 読めた総ページ数はここでも控える。小さなPDFは完了フラグが2秒の
' ポーリング周期より先に出るため、待ちループが1度も総頁を読まないまま
' 終わることがある(そのままではOCR経路が分母を受け取れない)。
Private Function GsOutHasPages(ByVal folderPath As String) As Boolean
    Dim n As Long
    n = optOcrCore.GsTotalPagesFromLog( _
        ReadTextHead(optOcrCore.GsLogFor(folderPath), LOG_HEAD_CHARS))
    ' R40 F1(レビュー M2): ログは -sstdout 経由になり、環境により空でありうる。
    ' ページ別ファイル gstext_%04d.txt が1つでも在れば GS はページを処理した
    ' (=画像PDFの分類と総頁の分母を失わない)。
    If n <= 0 Then n = optOcrCore.MaxPageFileIndex(folderPath, 0)
    If n > mLastTotalPages Then mLastTotalPages = n
    GsOutHasPages = (n > 0)
End Function

' ----------------------------------------------------------------------------
' WaitGsTextDone - txtwrite の完了待ち(2026-08-03 R13-1c/1d)。
'   従来の「固定120秒で完了フラグを待つだけ」を生存監視型の二段構えへ。
'     ・2秒ごとに gs_out.log の "Page N" 最大値と gstext.txt のサイズを見て、
'       どちらかが進んでいればアイドル起点を今へ戻す(まだ動いている)。
'     ・idleSec(=vision_pdf_timeout_sec)は無進捗の許容秒数、
'       absSec(=gs_abs_timeout_sec)は進んでいても打ち切る絶対上限。
'     ・1秒ごとに進捗バナーを更新(見えている時だけ。StageBanner側の規約)。
'   戻り値 True=完了フラグを確認 / False=タイムアウト。ByRefで終了コード
'   (-1=不明)・到達ページ・総ページを返す。タイムアウト時はPIDが取れて
'   いれば系統ごと停止する(1e。放置するとGS/cmd.exeが残り続ける=RC3)。
' ----------------------------------------------------------------------------
Private Function WaitGsTextDone(ByVal folderPath As String, ByVal outTxt As String, _
                                ByVal idleSec As Long, ByVal absSec As Long, _
                                ByVal pid As Long, ByRef rcOut As Long, _
                                ByRef pagesSeen As Long, ByRef totalPages As Long) As Boolean
    Dim flagPath As String: flagPath = optOcrCore.DoneFlagFor(folderPath)
    Dim logPath As String: logPath = optOcrCore.GsLogFor(folderPath)

    Dim t0 As Double: t0 = Timer
    Dim tIdle As Double: tIdle = t0
    Dim tPoll As Double: tPoll = t0
    Dim tBanner As Double: tBanner = t0
    Dim lastSize As Double: lastSize = -1#

    ' R13-F13b: どの生存信号が実機で本当に効いたかの計数(最後に1回だけ記録)。
    Dim pageAdv As Long, sizeAdv As Long, noneAdv As Long

    rcOut = -1
    pagesSeen = 0
    totalPages = 0

    Do
        optGsProc.SyncDoneFlag flagPath   ' R40 F1: GS終了→VBAがフラグを書く
        If optVision.PathExists(flagPath) Then
            rcOut = ReadFlagRcRetry(flagPath)
            WaitGsTextDone = True
            Exit Do
        End If

        DoEvents
        SleepOneSec

        Dim nowT As Double: nowT = Timer
        ' 日跨ぎでTimerが0へ戻ったら、全ての起点を今へ寄せて待ち続ける
        ' (「進んでいるのに0時を回ったから打ち切る」を起こさない)。
        If nowT < t0 Then
            t0 = nowT: tIdle = nowT: tPoll = nowT: tBanner = nowT
        End If

        If (nowT - tPoll) >= POLL_SEC Then
            tPoll = nowT
            If totalPages <= 0 Then
                ' R13-F1: ここも300字では起動バナーに押し出される(定数参照)。
                totalPages = optOcrCore.GsTotalPagesFromLog(ReadTextHead(logPath, LOG_HEAD_CHARS))
            End If
            Dim newPages As Long: newPages = optOcrCore.GsPagesFromLog(ReadTextTail(logPath, 500))
            ' 2026-09-07(R39 F001): 「1本のgstext.txt」ではなく最新のページ
            ' ファイルのサイズを見る(FileSizeOf/LofSizeOfは不要になり削除)。
            Dim newSize As Double: newSize = optOcrCore.LatestPageFileSize(folderPath)
            If newPages > pagesSeen Or newSize > lastSize Then
                If newPages > pagesSeen Then
                    pagesSeen = newPages
                    pageAdv = pageAdv + 1
                End If
                If newSize > lastSize Then
                    lastSize = newSize
                    sizeAdv = sizeAdv + 1
                End If
                tIdle = nowT
            Else
                noneAdv = noneAdv + 1
            End If
        End If

        If (nowT - tBanner) >= BANNER_SEC Then
            tBanner = nowT
            ' 表示の失敗が取込を壊してはならない(憲章§4-4)。
            On Error Resume Next
            modShelfBatch.StageBanner optOcrCore.GsWaitBanner( _
                pagesSeen, totalPages, CLng(Int(nowT - t0)))
            On Error GoTo 0
        End If

        If (nowT - tIdle) >= idleSec Then Exit Do
        If (nowT - t0) >= absSec Then Exit Do
    Loop

    ' R13-F13b: 成功・タイムアウトのどちらでも1回だけ、どの信号が生きていたかを
    ' 残す。cmd のブロックバッファリングで gs_out.log が遅れる端末、開いたまま
    ' のファイルで FileLen が固まる端末では、どちらか(最悪は両方)が none に
    ' 倒れる。次の実機テストで「効いた信号」を事実として拾うための計測。
    On Error Resume Next
    modLog.LogUsage "gs_progress_signal", "", "page=" & pageAdv & _
        ";size=" & sizeAdv & ";none=" & noneAdv
    On Error GoTo 0

    ' W3-8でFunction化。ここは戻り値不要(元々消さない設計)。
    If Not WaitGsTextDone Then optGsProc.KillGsTree pid
End Function

' 完了フラグの中身(終了コード)を最大2秒粘って読む(R13-1a)。
' cmd.exe の `echo %^ERRORLEVEL% >flag` はファイル作成と書込みが別操作なので、
' 「存在するが空」の一瞬がある。そこで読んで諦めていたのが rc=? flag=[] の正体。
' 2秒たっても数字にならなければ -1(不明)を返す。呼び出し元は image と
' 決めつけず flag_delayed として扱う。
Private Function ReadFlagRcRetry(ByVal flagPath As String) As Long
    Dim i As Long
    For i = 1 To FLAG_RETRY_N
        Dim rc As Long
        rc = optOcrCore.GsExitCodeFromFlag(ReadTextHead(flagPath, 80))
        If rc >= 0 Then
            ReadFlagRcRetry = rc
            Exit Function
        End If
        DoEvents
        SleepTick
    Next i
    ReadFlagRcRetry = -1
End Function

' 待ちループ1周ぶんの間引き(約100ms)。Application.Wait が使えない環境でも
' 待たずに回るだけで壊れない(R7 B-2の判断をそのまま踏襲)。
Private Sub SleepTick()
    On Error Resume Next
    Application.Wait Now + 0.1 / 86400#
    On Error GoTo 0
End Sub

' 完了待ちループ1周ぶんの休止(約1秒。R13-M2)。
' Now は秒単位までしか持たないため Now + 0.1/86400# は【過去】を指し、
' Application.Wait は一切眠らずに返っていた。結果、待ちループは最悪20分間
' 全速で Dir$ を叩き続ける(EDR配下では特に高くつく)。TimeSerial で必ず
' 1秒進める。Application.Wait が使えない環境では従来どおり空回りする。
Private Sub SleepOneSec()
    On Error Resume Next
    Application.Wait Now + TimeSerial(0, 0, 1)
    On Error GoTo 0
End Sub

' 2026-09-07(R39 F001): FileSizeOf/LofSizeOfは削除(単一gstext.txtのサイズ
' 監視は不要になった。呼び出し元は optOcrCore.LatestPageFileSize を使う)。

' ----------------------------------------------------------------------------
' Ghostscript実行の共通道具(R10-3bで optVision から移設)
' ----------------------------------------------------------------------------
' 元は optVision の内部ヘルパーだったが、R10-3でtxtwrite経路からも使うように
' なり、optVisionが30,000字上限まで残り15字という状態になった。GS実行の
' 道具一式(一時フォルダ・完了待ち・後始末)はどちらの経路
' (txtwrite / OCRのjpeg化)からも同じ形で使うものなので、ここへまとめて
' 置き、optVision 側は optGsTxt.〜 で呼ぶ。中身は移設前と1文字も変えていない。

' 一時フォルダ(TEMP\nxocr_<一意名>)を作って返す。失敗時は""。
Public Function MakeOcrFolder() As String
    Dim tempRoot As String: tempRoot = Environ$("TEMP")
    If LenB(tempRoot) = 0 Then tempRoot = Environ$("TMP")
    If LenB(tempRoot) = 0 Then Exit Function

    Dim uniqueName As String
    uniqueName = Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Rnd() * 9000) + 1000)

    Dim folderPath As String
    folderPath = optOcrCore.TempFolderFor(tempRoot, uniqueName)

    On Error Resume Next
    MkDir folderPath
    On Error GoTo 0

    If FolderExists(folderPath) Then MakeOcrFolder = folderPath
End Function

' 完了フラグの出現をDoEventsつきで待つ。Trueで完了、Falseでタイムアウト。
' 2026-07-31(R7 B-2)/2026-08-03(R13-M2): 間引きは【1秒】。DoEvents専業だと
' GSが画像を書いている数分のあいだCPUを1コア回し切る。待っているのはファイルの
' 出現で、詰めても早くは終わらない(Application.Wait が無い環境でも壊れない)。
' 2026-08-03(R13-1a): フラグの「存在」だけで完了とみなすのをやめた。
' 存在を見た後、中身が終了コードとして読めるまで最大2秒粘る(ReadFlagRcRetry)。
' rcOut は省略可なので、既存の2引数呼び出し(optVisionのOCR経路)はそのまま動く。
' 2026-08-04(R15-FixA FA-5i): この待ちは実機で数分に達するが、従来ここには
' 中断の確認も実況も1つも無かった。利用者から見ると「⏹中断を押しても何分も
' 何も起きない・画面も動かない」区間で、そこで強制終了されていた(B-H3)。
'   ・中断が申し出られていたら早期に False で戻る(呼び出し元が時間切れと
'     混同しないよう、あちらが自分でもう一度 CancelRequested を見る)。
'   ・1秒ごとにバナーへ経過秒を流す(bannerLabel が空なら従来どおり無言)。
Public Function WaitForDoneFlag(ByVal flagPath As String, ByVal timeoutSec As Long, _
                                Optional ByRef rcOut As Long = -1, _
                                Optional ByVal bannerLabel As String = "") As Boolean
    Dim t0 As Double: t0 = Timer
    Dim tBanner As Double: tBanner = t0
    Do
        optGsProc.SyncDoneFlag flagPath   ' R40 F1: GS終了→VBAがフラグを書く
        If optVision.PathExists(flagPath) Then
            rcOut = ReadFlagRcRetry(flagPath)
            WaitForDoneFlag = True
            Exit Function
        End If
        If CancelWanted() Then Exit Function
        DoEvents
        SleepOneSec
        If Timer < t0 Then                 ' 日跨ぎでTimerが0へ戻った場合の保険
            t0 = Timer
            tBanner = t0
        End If
        If LenB(bannerLabel) > 0 And (Timer - tBanner) >= BANNER_SEC Then
            tBanner = Timer
            ' 表示の失敗が取込を壊してはならない(憲章§4-4)。
            On Error Resume Next
            modShelfBatch.StageBanner optOcrEta.RenderWaitBanner( _
                bannerLabel, CLng(Int(Timer - t0)))
            On Error GoTo 0
        End If
    Loop While (Timer - t0) < timeoutSec
End Function

' 利用者が進捗バナーの中断ボタンを押したか(R15-FixA FA-5i)。印の実体は
' modShelfBatch(取込の入口でリセットされる)にあり、ここは読むだけ。
' 問い合わせで例外が出ても待ちを壊さない=「押されていない」に倒す。
Private Function CancelWanted() As Boolean
    On Error Resume Next
    CancelWanted = modShelfBatch.CancelRequested()
    On Error GoTo 0
End Function

' 一時フォルダの後始末(成功・失敗の両経路から呼ぶ。R6規約により別Sub)。
' ロック中でKillに失敗しても無視する(TEMPなのでOSが後で片付ける)。
Public Sub CleanupOcrFolder(ByVal folderPath As String)
    If LenB(folderPath) = 0 Then Exit Sub
    On Error Resume Next
    Kill folderPath & "\*.jpg"
    Kill folderPath & "\*.flag"
    Kill folderPath & "\*.log"   ' R11-D: GSの出力ログ(gs_out.log)も片付ける
    RmDir folderPath
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' GsFailureDetail - 完了フラグとGS出力ログから「なぜ画像/テキストが出ないのか」
'   を1本の文字列にまとめる(2026-07-31 R11-D・監査3 H-2)。
'   err_log の detail へそのまま入れる用途。読めないものは黙って飛ばす
'   (観測のための処理が取込を壊してはならない)。
'   戻り値の形: "rc=<終了コード or ?> flag=[<生の中身>] gs_out=[<末尾500字>]"
'   2026-08-03(R13-1f): gs_out は【末尾】500字にした。エラーはログの最後に
'   出るのに先頭を切り出していたため、全ページ完走していても "Processing pages"
'   の途中で切れた同じ景色しか見えず、実機第2報では切り分けができなかった。
'   ページ進捗の行頭が切れても構わない(そちらは進捗監視が別に見ている)。
' ----------------------------------------------------------------------------
Public Function GsFailureDetail(ByVal folderPath As String) As String
    Dim flagText As String: flagText = ReadTextHead(optOcrCore.DoneFlagFor(folderPath), 80)
    Dim logText As String: logText = ReadTextTail(optOcrCore.GsLogFor(folderPath), 500)

    Dim rc As Long: rc = optOcrCore.GsExitCodeFromFlag(flagText)
    Dim rcText As String
    If rc < 0 Then rcText = "?" Else rcText = CStr(rc)

    GsFailureDetail = "rc=" & rcText & " flag=[" & Trim$(flagText) & "]" & _
        " gs_out=[" & logText & "]"
End Function

' 2026-08-03(R14-F13): GsExitCode をここから削除した。終了コードは待ちループ
' (ReadFlagRcRetry)と GsFailureDetail が読んでおり、外からの呼び出しは
' 1件も無かった(契約に載っているだけの公開API)。

' GSへ渡した入力ファイルの実バイト数(取れなければ -1 = 不明。R14-3c)。
' FileLen は実体が無いと実行時エラー53を出すので握る(観測用の材料であって、
' ここで失敗しても本処理を止めてはならない)。
Private Function InputFileBytes(ByVal p As String) As Long
    On Error Resume Next
    InputFileBytes = -1
    InputFileBytes = FileLen(p)
    On Error GoTo 0
End Function

' テキストファイルの先頭 maxChars 字だけを読む(改行はスペースへ潰す)。
' 失敗しても例外を出さず ""(観測用なので絶対に本処理を止めない)。
' GSのログはCP932(コンソールのコードページ)なので、ADODB.Streamではなく
' 素のOpen/Inputで読む(Charset指定を誤ると文字化けだけが残るため)。
Private Function ReadTextHead(ByVal filePath As String, ByVal maxChars As Long) As String
    Dim fn As Long: fn = 0
    Dim buf As String

    On Error GoTo NoRead
    If LenB(Dir$(filePath)) = 0 Then Exit Function
    fn = FreeFile
    Open filePath For Binary Access Read As #fn
    Dim n As Double: n = LOF(fn)
    If n > maxChars Then n = maxChars
    If n > 0 Then
        buf = Space$(CLng(n))
        Get #fn, 1, buf
    End If
    Close #fn
    fn = 0

    buf = Replace(buf, vbCr, " ")
    buf = Replace(buf, vbLf, " ")
    ReadTextHead = buf
    Exit Function
NoRead:
    ' ハンドラ稼働中は On Error Resume Next が効かない。後始末は Resume で
    ' ハンドラを抜けてから、別Subで行う(本モジュール共通の作法)。
    Resume ReadHeadCleanup
ReadHeadCleanup:
    CloseFileNumberSafely fn
    ReadTextHead = ""
End Function

' テキストファイルの【末尾】 maxChars 字だけを読む(2026-08-03 R13-1f)。
' ReadTextHead と対の道具で、読み方(素のOpen/Binary・CP932のバイト列を
' そのまま文字数として扱う・改行はスペースへ潰す)は完全に同一にしてある。
' Ghostscriptのログは異常終了の理由を最後に書くので診断はこちらを使う。
' 失敗しても例外を出さず ""(観測用なので絶対に本処理を止めない)。
Private Function ReadTextTail(ByVal filePath As String, ByVal maxChars As Long) As String
    Dim fn As Long: fn = 0
    Dim buf As String

    On Error GoTo NoReadTail
    If LenB(Dir$(filePath)) = 0 Then Exit Function
    fn = FreeFile
    Open filePath For Binary Access Read As #fn
    Dim total As Double: total = LOF(fn)
    Dim n As Double: n = total
    If n > maxChars Then n = maxChars
    Dim startPos As Double: startPos = total - n + 1#
    If startPos < 1# Then startPos = 1#
    If n > 0 Then
        buf = Space$(CLng(n))
        Get #fn, CLng(startPos), buf
    End If
    Close #fn
    fn = 0

    buf = Replace(buf, vbCr, " ")
    buf = Replace(buf, vbLf, " ")
    ReadTextTail = buf
    Exit Function
NoReadTail:
    ' ReadTextHead と同じ作法: ハンドラを Resume で抜けてから後始末する。
    Resume ReadTailCleanup
ReadTailCleanup:
    CloseFileNumberSafely fn
    ReadTextTail = ""
End Function

' 開きかけのファイル番号を確実に閉じる(0=未取得は何もしない)。
Private Sub CloseFileNumberSafely(ByVal fn As Long)
    If fn = 0 Then Exit Sub
    On Error Resume Next
    Close #fn
    On Error GoTo 0
End Sub

Private Function FolderExists(ByVal p As String) As Boolean
    On Error Resume Next
    FolderExists = (LenB(Dir$(p, vbDirectory)) > 0)
    On Error GoTo 0
End Function

