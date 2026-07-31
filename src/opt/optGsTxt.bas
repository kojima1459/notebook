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
'     ブロッキングで待たない。タイムアウトは config vision_pdf_timeout_sec。
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
        optOcrCore.BuildGsTextCommand(gsExe, path, outTxt), flagPath)

    waitSec = modConfig.GetLong("vision_pdf_timeout_sec", 120)
    If waitSec < 10 Then waitSec = 10

    modUIMain.SetStage "PDFの文字を取り出しています…"

    If Not RunGsAsync(runCmd, gsErrNum, gsErrDesc) Then
        modUIMain.SetStage ""
        CleanupTxtFolder folderPath, outTxt
        ' WScript.Shell自体がポリシーで塞がれている端末をここで切り分ける。
        modLog.LogError "E0302", "optGsTxt.ExtractPdfTextNoOcr", modUtil.SafeLeft( _
            "GS起動失敗 err#" & gsErrNum & ": " & gsErrDesc & " " & path, 2000)
        ExtractPdfTextNoOcr = ERR_302 & "PDFから文字を取り出す処理を開始できませんでした。"
        Exit Function
    End If

    finished = WaitForDoneFlag(flagPath, waitSec)

    ' R10c(H1): タイムアウトは【失敗】。完了フラグが出ていない間、GSはまだ
    ' 出力txtを書いている途中なので、その時点で40字以上読めたとしても
    ' 「本文の途中まで」でしかない。採用すると資料が黙って欠けたまま本棚に
    ' 入り、以後の回答が虫食いの本文に基づくものになる(誰も気付けない)。
    ' Word→Acrobat連鎖へ譲るため E0302 系トークンで返す。
    If Not finished Then
        modUIMain.SetStage ""
        CleanupTxtFolder folderPath, outTxt
        modLog.LogError "E0302", "optGsTxt.ExtractPdfTextNoOcr", modUtil.SafeLeft( _
            "GS_TIMEOUT " & waitSec & "秒以内に完了フラグが出ませんでした " & path, 2000)
        ExtractPdfTextNoOcr = TOKEN_TIMEOUT & "PDFからの文字取り出しが" & waitSec & _
            "秒以内に終わりませんでした。"
        Exit Function
    End If

    txt = ReadUtf8Text(outTxt, gsErrNum, gsErrDesc)

    modUIMain.SetStage ""
    CleanupTxtFolder folderPath, outTxt

    ' ADODB.StreamはUTF-8のBOMを落とすが、環境差で残ることがあるので念のため。
    If Len(txt) > 0 Then
        If Left$(txt, 1) = ChrW(&HFEFF) Then txt = Mid$(txt, 2)
    End If

    If LenB(txt) = 0 Then
        modLog.LogError "E0302", "optGsTxt.ExtractPdfTextNoOcr", modUtil.SafeLeft( _
            "txtwrite出力を読めず err#" & gsErrNum & ": " & gsErrDesc & " " & path, 2000)
        ExtractPdfTextNoOcr = ERR_302 & "PDFから文字を取り出せませんでした" & _
            "(出力ファイルを読めませんでした)。"
        Exit Function
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
Private Function ReadUtf8Text(ByVal txtPath As String, ByRef errNum As Long, _
                              ByRef errDesc As String) As String
    Dim st As Object

    On Error GoTo Failed
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2          ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile txtPath
    ReadUtf8Text = st.ReadText
    st.Close
    Set st = Nothing
    Exit Function

Failed:
    errNum = Err.Number
    errDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きたエラーは
    ' 呼び出し元へ飛んで本来の原因を上書きする。後始末の前に Resume で
    ' ハンドラを抜ける(modExtractor.ExtractPlainText と同型)。
    Resume ReadCleanup
ReadCleanup:
    If Not st Is Nothing Then
        On Error Resume Next
        st.Close
        On Error GoTo 0
    End If
    Set st = Nothing
    ReadUtf8Text = ""
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
Public Function RunGsAsync(ByVal runCmd As String, ByRef errNum As Long, _
                            ByRef errDesc As String) As Boolean
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

' 完了フラグの出現をDoEventsつきで待つ。Trueで完了、Falseでタイムアウト。
' 2026-07-31(R7 B-2ついで): このループは DoEvents 専業で、Ghostscript が
' ページ画像を書いている数分のあいだCPUを1コア回し切っていた(R6の報告)。
' 待っているのはファイルの出現であって、詰めても早くは終わらない。
' 1周ごとに約100ms止めて間引く(挙動は不変。判定間隔が0.1秒になるだけ)。
' Application.Wait が使えない環境でも待たずに回るだけで壊れない。
Public Function WaitForDoneFlag(ByVal flagPath As String, ByVal timeoutSec As Long) As Boolean
    Dim t0 As Double: t0 = Timer
    Do
        If optVision.PathExists(flagPath) Then
            WaitForDoneFlag = True
            Exit Function
        End If
        DoEvents
        On Error Resume Next
        Application.Wait Now + 0.1 / 86400#
        On Error GoTo 0
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
    RmDir folderPath
    On Error GoTo 0
End Sub

Private Function FolderExists(ByVal p As String) As Boolean
    On Error Resume Next
    FolderExists = (LenB(Dir$(p, vbDirectory)) > 0)
    On Error GoTo 0
End Function

