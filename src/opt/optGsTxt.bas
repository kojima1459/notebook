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
'   (2) Ghostscriptを動かすための共通道具(一時フォルダ・非同期起動・完了
'       待ち・後始末)の置き場。txtwrite経路と、optVisionの画像PDF OCR経路
'       (jpeg化)の両方がここを使う(R10-3bでoptVisionから移設)。
'
' なぜ第1選択なのか(2026-07-31 実機報告):
'   ・会社の管理端末では Word・Acrobat の CreateObject が塞がれており、
'     テキストPDFが1本も取り込めなかった(実機初報B・E0302)。
'   ・塞がれていない端末でも、WordのPDF Reflowは「'Word'がOLE操作を完了する
'     のを待っています」ダイアログを頻発させ、数分たっても終わりが見えない
'     体験になっていた(追加報告)。
'   GSは別プロセスで、完了フラグの監視によるタイムアウト制御が効き、
'   テキストPDFなら数秒で返る。ここを先に通すことでOLE待ちダイアログという
'   最大のUX問題ごと消す。GSで駄目だったときだけ Word → Acrobat の既存連鎖へ
'   落ちる(modExtractor.ExtractPdfWithFallback)。
'
' 設計判断:
'   ・置き場所: 当初は道具一式のある optVision へ足す想定だったが、optVision
'     は30,000字上限まで残りが無く入らなかった。R10-3で実行部を本モジュール
'     へ分け(道具はoptVision側をPublic化して共用)、R10-3bで道具そのものを
'     こちらへ移設して役割を整理した(optGsTxt=GS実行、optVision=Vision API
'     +OCRオーケストレーション)。コマンド文字列の組み立てと採否判定は
'     optOcrCore(純ロジック)。
'   ・入口は modFeatures.InvokeFeature("vision","ExtractPdfTextNoOcr",…)。
'     vision機能の行き先は optVision 固定(modFeatures.ModuleNameOf)なので、
'     optVision側に同名の薄い受け口を置き、そこからここへ転送している。
'   ・GS未検出でも案内カード(フォルダ選択ダイアログ)は出さない。ここは
'     「全てのPDF取込で毎回通る道」なので、モーダルが出ると取込のたびに
'     割り込むことになる。候補探索だけの静かな解決(optVision.
'     FindGsExeByCandidates)を使い、見つからなければ黙ってWord経路へ譲る。
'     案内カードは従来どおり画像PDFのOCR経路(ExtractPdfOcrPagedText)だけの
'     役目のまま(そちらは利用者がOCRを期待している場面なので割り込んでよい)。
'   ・mock_llm=TRUE でも動かす。txtwriteはLLMもAIリボンも使わないため、
'     mockで止める理由が無い(ExtractPdfOcrPagedTextのmockガードはChatGPTVを
'     呼ぶからであって、ここには当てはまらない)。
'   ・戻り値の規約(optVisionの各関数と同じ "#ERR:" 文字列):
'       成功          : 抽出したテキスト(ページ区切りは改ページ文字のまま)
'       文字層ゼロ    : "#ERR:E0303:…" スキャンPDF確定。呼び出し元は即OCRへ
'       文字が薄い    : "#ERR:E0302:GS_SPARSE:…" 呼び出し元は Word → Acrobat を
'                       先に試し、両方失敗したときだけOCR(E0303)へ回す
'       タイムアウト  : "#ERR:E0302:GS_TIMEOUT:…" 途中まで読めていても採用しない
'       その他の失敗  : "#ERR:E0302:…" 呼び出し元は Word → Acrobat へ落ちる
'   ・失敗を握りつぶさない。CreateObject("WScript.Shell") まで塞がれている
'     端末があり得るので、GS起動・出力読取で拾ったErr.Number/Descriptionは
'     必ずerr_logのdetailへ残す(切り分けの唯一の手がかりになる)。ただし
'     「GSが置かれていないだけ」は想定内かつ全PDFで毎回起きるので、err_logは
'     書かず戻り値の文面だけで伝える(呼び出し元がWord/Acrobatも失敗した
'     ときに3者併記のdetailとしてまとめて記録する)。
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

' R13-1c: 生存監視の刻み。待ちループは100msごとに回り、2秒ごとに進捗
'(gs_out.log の "Page N" と gstext.txt のサイズ)を見て、1秒ごとにバナーを
' 更新する。数字を3箇所に散らさないためここに置く。
Private Const POLL_SEC As Double = 2#
Private Const BANNER_SEC As Double = 1#

' R13-1e: WMI起動が使えずPIDを取れなかったことを1セッション1回だけ記録する
' ためのフラグ(EDR等でWMIが塞がれている端末では毎回失敗するので、
' 記録が usage_log を埋め尽くさないようにする)。
Private mKillUnavailableLogged As Boolean

Public Function Ping() As Boolean
    Ping = True
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
    If Not RunGsAsync(runCmd, gsErrNum, gsErrDesc, gsPid) Then
        modUIMain.SetStage ""
        CleanupTxtFolder folderPath, outTxt
        ' WScript.Shell自体がポリシーで塞がれている端末をここで切り分ける。
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

    txt = ReadUtf8Text(outTxt, gsErrNum, gsErrDesc)

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
    If LenB(txt) = 0 Then
        Dim cls As String
        cls = optOcrCore.ClassifyGsTextResult(gsRc, 0, hasPages)

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
    On Error GoTo 0
    CleanupOcrFolder folderPath
End Sub

' txtwriteの出力を ADODB.Stream(Charset "utf-8")で読む
' (modExtractor.ExtractPlainText と同じ作法)。読めなければ ""。
' 失敗理由は呼び出し元のerr_log用にByRefで返す。
' UTF-8読み取りの実体は modUtilText.ReadTextFileUtf8(2026-07-31 R11-F2)。
' 先頭BOMの除去も向こうが行う(旧実装はここの呼び出し元で個別に落としていた)。
Private Function ReadUtf8Text(ByVal txtPath As String, ByRef errNum As Long, _
                              ByRef errDesc As String) As String
    Dim txt As String
    If modUtilText.ReadTextFileUtf8(txtPath, txt, errNum, errDesc) Then ReadUtf8Text = txt
End Function

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
Private Function GsOutHasPages(ByVal folderPath As String) As Boolean
    GsOutHasPages = (optOcrCore.GsTotalPagesFromLog( _
        ReadTextHead(optOcrCore.GsLogFor(folderPath), 300)) > 0)
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

    rcOut = -1
    pagesSeen = 0
    totalPages = 0

    Do
        If optVision.PathExists(flagPath) Then
            rcOut = ReadFlagRcRetry(flagPath)
            WaitGsTextDone = True
            Exit Function
        End If

        DoEvents
        SleepTick

        Dim nowT As Double: nowT = Timer
        ' 日跨ぎでTimerが0へ戻ったら、全ての起点を今へ寄せて待ち続ける
        ' (「進んでいるのに0時を回ったから打ち切る」を起こさない)。
        If nowT < t0 Then
            t0 = nowT: tIdle = nowT: tPoll = nowT: tBanner = nowT
        End If

        If (nowT - tPoll) >= POLL_SEC Then
            tPoll = nowT
            If totalPages <= 0 Then
                totalPages = optOcrCore.GsTotalPagesFromLog(ReadTextHead(logPath, 300))
            End If
            Dim newPages As Long: newPages = optOcrCore.GsPagesFromLog(ReadTextTail(logPath, 500))
            Dim newSize As Double: newSize = FileSizeOf(outTxt)
            If newPages > pagesSeen Or newSize > lastSize Then
                If newPages > pagesSeen Then pagesSeen = newPages
                If newSize > lastSize Then lastSize = newSize
                tIdle = nowT
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

    KillGsTree pid
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

' 書きかけファイルのサイズ(読めなければ -1)。観測用なので絶対に落とさない。
Private Function FileSizeOf(ByVal filePath As String) As Double
    FileSizeOf = -1#
    On Error Resume Next
    FileSizeOf = CDbl(FileLen(filePath))
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' Ghostscript実行の共通道具(R10-3bで optVision から移設)
' ----------------------------------------------------------------------------
' 元は optVision の内部ヘルパーだったが、R10-3でtxtwrite経路からも使うように
' なり、optVisionが30,000字上限まで残り15字という状態になった。GS実行の
' 道具一式(一時フォルダ・非同期起動・完了待ち・後始末)はどちらの経路
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

' 非同期起動(待たない)。0=ウィンドウ非表示 / False=完了を待たない。
' R10-2: 失敗時はErr.Number/Descriptionを呼び出し元へByRefで返す(戻り値は
' Booleanのまま・モジュール変数を増やさない最小構成)。呼び出し元がerr_logの
' detailへ含めることで、WScript.Shell自体がポリシーでブロックされる端末を
' 「パスは合っているのにGSが動かない」から切り分けられるようにする。
' R13-1e: まずWMI(Win32_Process.Create)で起動してPIDを控える。PIDが無いと
' 真のハングを止める手段が無く、GS/cmd.exe がログオン中ずっと残る(RC3)。
' WMIはEDR/ポリシーで塞がれる端末があるので、失敗したら【必ず】従来の
' WScript.Shell へ退避する(取込が動かなくなる方が害が大きい)。退避した
' セッションではkillできないことを1回だけ usage_log へ残す。
' pidOut: 取れたPID(0=不明。省略可なので既存の3引数呼び出しはそのまま動く)。
Public Function RunGsAsync(ByVal runCmd As String, ByRef errNum As Long, _
                            ByRef errDesc As String, _
                            Optional ByRef pidOut As Long = 0) As Boolean
    pidOut = 0
    If RunGsViaWmi(runCmd, pidOut) Then
        RunGsAsync = True
        Exit Function
    End If
    NoteKillUnavailable

    Dim wsh As Object
    On Error GoTo NoRun
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run runCmd, 0, False
    Set wsh = Nothing
    RunGsAsync = True
    Exit Function
NoRun:
    errNum = Err.Number
    errDesc = Err.Description
    RunGsAsync = False
End Function

' WMI経由の非表示起動。成功でTrue+PID、どこかで失敗したらFalse(呼び出し元が
' 従来経路へ退避)。GetObject/Get/SpawnInstance_/Create のどれもポリシーで
' 落ちうるので1本のハンドラで受け、Resumeでハンドラを抜けてからCOM参照を
' 必ず解放する(本モジュール共通の作法。ハンドラ稼働中はOERNが効かない)。
Private Function RunGsViaWmi(ByVal runCmd As String, ByRef pidOut As Long) As Boolean
    Dim svc As Object
    Dim startCls As Object
    Dim cfg As Object
    Dim proc As Object
    Dim newPid As Variant
    Dim rc As Long

    On Error GoTo WmiFail
    Set svc = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set startCls = svc.Get("Win32_ProcessStartup")
    Set cfg = startCls.SpawnInstance_()
    cfg.ShowWindow = 0                 ' 0=SW_HIDE(従来のwsh.Run(...,0,False)と同じ)
    Set proc = svc.Get("Win32_Process")
    rc = proc.Create(runCmd, Null, cfg, newPid)

    Set proc = Nothing
    Set cfg = Nothing
    Set startCls = Nothing
    Set svc = Nothing
    On Error GoTo 0

    If rc <> 0 Then Exit Function
    If Not IsNumeric(newPid) Then Exit Function
    If CLng(newPid) <= 0 Then Exit Function

    pidOut = CLng(newPid)
    RunGsViaWmi = True
    Exit Function

WmiFail:
    Resume WmiCleanup
WmiCleanup:
    On Error Resume Next
    Set proc = Nothing
    Set cfg = Nothing
    Set startCls = Nothing
    Set svc = Nothing
    On Error GoTo 0
    RunGsViaWmi = False
End Function

' PIDが取れなかったことの記録(1セッション1回)。「なぜ止められなかったのか」
' が後から分かるようにする(憲章§4-1 無言の失敗禁止)。
Private Sub NoteKillUnavailable()
    If mKillUnavailableLogged Then Exit Sub
    mKillUnavailableLogged = True
    On Error Resume Next
    modLog.LogUsage "gs_kill_unavailable", "", _
        "WMI起動が使えずWScript.Shellへ退避(PID未取得のためハング時に停止できません)"
    On Error GoTo 0
End Sub

' 真のハング(アイドル上限超過)でGSを系統ごと止める(R13-1e)。
' /T で子プロセス(cmd.exe が起動した gswin32c.exe)まで、/F で強制終了。
' PIDが無い(WScript.Shell退避)ときは何もしない。
Private Sub KillGsTree(ByVal pid As Long)
    If pid <= 0 Then Exit Sub

    Dim wsh As Object
    On Error GoTo KillFail
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "taskkill /T /F /PID " & CStr(pid), 0, False
    Set wsh = Nothing
    On Error GoTo 0
    On Error Resume Next
    modLog.LogUsage "gs_killed_on_hang", "", "pid=" & pid
    On Error GoTo 0
    Exit Sub
KillFail:
    Resume KillCleanup
KillCleanup:
    On Error Resume Next
    Set wsh = Nothing
    modLog.LogUsage "gs_kill_failed", "", "pid=" & pid
    On Error GoTo 0
End Sub

' 完了フラグの出現をDoEventsつきで待つ。Trueで完了、Falseでタイムアウト。
' 2026-07-31(R7 B-2ついで): このループは DoEvents 専業で、Ghostscript が
' ページ画像を書いている数分のあいだCPUを1コア回し切っていた(R6の報告)。
' 待っているのはファイルの出現であって、詰めても早くは終わらない。
' 1周ごとに約100ms止めて間引く(挙動は不変。判定間隔が0.1秒になるだけ)。
' Application.Wait が使えない環境でも待たずに回るだけで壊れない。
' 2026-08-03(R13-1a): フラグの「存在」だけで完了とみなすのをやめた。
' 存在を見た後、中身が終了コードとして読めるまで最大2秒粘る(ReadFlagRcRetry)。
' rcOut は省略可なので、既存の2引数呼び出し(optVisionのOCR経路)はそのまま動く。
Public Function WaitForDoneFlag(ByVal flagPath As String, ByVal timeoutSec As Long, _
                                Optional ByRef rcOut As Long = -1) As Boolean
    Dim t0 As Double: t0 = Timer
    Do
        If optVision.PathExists(flagPath) Then
            rcOut = ReadFlagRcRetry(flagPath)
            WaitForDoneFlag = True
            Exit Function
        End If
        DoEvents
        SleepTick
        If Timer < t0 Then t0 = Timer      ' 日跨ぎでTimerが0へ戻った場合の保険
    Loop While (Timer - t0) < timeoutSec
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

' ----------------------------------------------------------------------------
' GsExitCode - 完了フラグの中身から終了コードを読む(0=正常 / -1=判定不能)。
' ----------------------------------------------------------------------------
Public Function GsExitCode(ByVal folderPath As String) As Long
    GsExitCode = optOcrCore.GsExitCodeFromFlag( _
        ReadTextHead(optOcrCore.DoneFlagFor(folderPath), 80))
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

