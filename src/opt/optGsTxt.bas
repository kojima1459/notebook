Attribute VB_Name = "optGsTxt"
Option Explicit

' ============================================================================
' optGsTxt - テキストPDFをGhostscript(txtwrite)でCOM無しに読む(opt機能・R10-3)
' ----------------------------------------------------------------------------
' 役割:
'   PDFの本文抽出の【第1選択】。同梱の gswin32c.exe に -sDEVICE=txtwrite を
'   渡し、PDFに埋まっている文字をそのまま書き出させて読む。COM(Word/Acrobat)
'   を一切使わないので、CreateObjectがポリシーで塞がれた管理端末でも動く。
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
'   ・置き場所: 道具一式(GS起動・完了待ち・一時フォルダ)がある optVision へ
'     足すのが素直だが、optVisionは30,000字上限まで残りが無く入らない。
'     要件R10-3の事前承認に従って実行部だけを本モジュールへ分けた。道具は
'     optVision側をPublic化して共用する(重複実装はしない)。コマンド文字列の
'     組み立てと採否判定は optOcrCore(純ロジック)。
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
'       ほぼ空        : "#ERR:E0303:…" 画像PDF疑い。呼び出し元はOCR経路へ回す
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

    folderPath = optVision.MakeOcrFolder()
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

    If Not optVision.RunGsAsync(runCmd, gsErrNum, gsErrDesc) Then
        modUIMain.SetStage ""
        CleanupTxtFolder folderPath, outTxt
        ' WScript.Shell自体がポリシーで塞がれている端末をここで切り分ける。
        modLog.LogError "E0302", "optGsTxt.ExtractPdfTextNoOcr", modUtil.SafeLeft( _
            "GS起動失敗 err#" & gsErrNum & ": " & gsErrDesc & " " & path, 2000)
        ExtractPdfTextNoOcr = ERR_302 & "PDFから文字を取り出す処理を開始できませんでした。"
        Exit Function
    End If

    finished = optVision.WaitForDoneFlag(flagPath, waitSec)
    txt = ReadUtf8Text(outTxt, gsErrNum, gsErrDesc)

    modUIMain.SetStage ""
    CleanupTxtFolder folderPath, outTxt

    ' ADODB.StreamはUTF-8のBOMを落とすが、環境差で残ることがあるので念のため。
    If Len(txt) > 0 Then
        If Left$(txt, 1) = ChrW(&HFEFF) Then txt = Mid$(txt, 2)
    End If

    If LenB(txt) = 0 Then
        modLog.LogError "E0302", "optGsTxt.ExtractPdfTextNoOcr", modUtil.SafeLeft( _
            "txtwrite出力を読めず finished=" & finished & " err#" & gsErrNum & _
            ": " & gsErrDesc & " " & path, 2000)
        ExtractPdfTextNoOcr = ERR_302 & "PDFから文字を取り出せませんでした" & _
            "(変換が終わらなかったか、出力を読めませんでした)。"
        Exit Function
    End If

    If optOcrCore.GsTextVerdict(CleanTextLen(txt)) = "image" Then
        ExtractPdfTextNoOcr = ERR_303 & "このPDFには文字データがほとんど入って" & _
            "いません(画像として保存されたPDFとみられます)。"
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

' 出力txtと一時フォルダの後始末。optVision.CleanupOcrFolder は *.jpg と
' *.flag しか消さない(OCR経路の規約)ので、txtは先に自分で消してから
' フォルダごと片付ける。ロック中で消せなくても無視する(TEMPなのでOSが
' 後で片付ける)。
Private Sub CleanupTxtFolder(ByVal folderPath As String, ByVal outTxt As String)
    If LenB(folderPath) = 0 Then Exit Sub
    On Error Resume Next
    If LenB(outTxt) > 0 Then Kill outTxt
    On Error GoTo 0
    optVision.CleanupOcrFolder folderPath
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

' 空白類(半角/全角スペース・タブ・改行・改ページ)を除いた文字数。
' txtwriteは画像PDFに対しても改ページ文字だけは律儀に書き出すので、素のLenで
' 判定すると「文字がある」と誤認する。1文字ずつ走査するとページ数の多いPDFで
' 無駄に重いため、Replaceで落としてから数える。
Private Function CleanTextLen(ByVal s As String) As Long
    Dim t As String

    t = Replace(s, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, vbTab, "")
    t = Replace(t, Chr$(12), "")
    t = Replace(t, " ", "")
    t = Replace(t, ChrW(&H3000), "")
    CleanTextLen = Len(t)
End Function
