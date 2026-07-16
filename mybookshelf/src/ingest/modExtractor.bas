Attribute VB_Name = "modExtractor"
Option Explicit

' ============================================================================
' modExtractor - 抽出ディスパッチャ(拡張子ごとに専門モジュールへ振り分け)
' ----------------------------------------------------------------------------
' 役割:
'   ファイルパスの拡張子を見て、専門モジュール(txt/md/csv=自前読込、
'   pdf=Word→Acrobatフォールバック、docx/doc=Word、xlsx/xls/xlsm=Excel)へ
'   処理を振り分け、結果を ExtractedPage() の配列に統一して返す。
'
' 流用元: /home/user/notebook/src/admin/modExtractor.bas(V2チャットボット資産。
'   コピー元として精読のみ・変更禁止)。
'
' 適応要件(MASTER_SPEC §7.2):
'   ・V1は失敗時 Err.Raise で例外を投げていたが、本モジュールは
'     Boolean 戻り値 + errCode/errDetail の ByRef 引数へ変換する
'     (呼び出し側 modShelf.IngestFile が例外ハンドリングを書かずに済むように)。
'     errCodeは E0301(対応外拡張子)/E0302(開けない)/E0303(画像PDF)/
'     E0304(抽出結果が空)のいずれか。成功時は "" のまま。
'   ・画像PDF検知: 抽出後の総文字数が「抽出できたページ数×10」未満なら
'     E0303(画像PDFの可能性が高いと判断)。PDF以外の拡張子には適用しない。
'   ・max_pages_per_file(config)を超えるページ/シートを持つファイルは、
'     各アダプタモジュール側で先頭 maxPages 件だけを抽出して打ち切る
'     (全ページ抽出してから捨てるのではなく、抽出自体を早期に止めることで
'     COM呼び出しコストを抑える)。打ち切りが発生した場合は戻り値 True の
'     まま errCode="PARTIAL_PAGES" を返し、呼び出し側が manifest の
'     status="partial" 相当の扱いをする(MASTER_SPEC §7.2)。
'   ・COM呼び出し(ADODB.Stream / Word.Application 等)はモジュール内で
'     必ず On Error により捕捉し、例外を外へ伝播させない。実際のCOM操作は
'     modExtractorWord/modExtractorExcel/modExtractorAcrobat 側が担当し、
'     各モジュールが自分の開いたプロセス(Word/Excelブック)を成功・失敗
'     いずれの経路でも必ず Quit/Close する(V1のFailedハンドラの作法を踏襲)。
'   ・Mac等COM非対応環境の案内は、各アダプタモジュールが実行時エラー
'     429(ActiveX component can't create object)を検知した場合に、
'     errDetail へ「Windows版Excelでお試しください」という丁寧な文言を
'     含める(§13)。errCodeそのものはE0302のまま(契約の4コード以外は
'     増やさない)。
'   ・失敗の記録(R5)は本モジュールが一元的に modLog.LogError を呼ぶ。
'     アダプタ側(modExtractorWord等)ではログを書かない(1件の失敗に対して
'     複数モジュールが二重にログを書く事態を避けるため、ログ責務は
'     ディスパッチャに集約する設計判断)。
' ============================================================================

Private Const SUPPORTED_EXTS As String = "txt,md,csv,pdf,docx,doc,xlsx,xls,xlsm"

' ----------------------------------------------------------------------------
' ExtractFile - 拡張子で振り分けて抽出する。
'   成功: True(pagesに1件以上、またはPARTIAL_PAGES時は打ち切り分を格納)
'   失敗: False + errCode(E0301/E0302/E0303/E0304) + errDetail(内部詳細)
' ----------------------------------------------------------------------------
Public Function ExtractFile(ByVal path As String, ByRef pages() As ExtractedPage, _
                            ByRef errCode As String, ByRef errDetail As String) As Boolean
    errCode = ""
    errDetail = ""
    ReDim pages(0 To -1)   ' 呼び出し前に必ず空配列へリセットしておく

    Dim ext As String: ext = modUtil.ExtOf(path)
    If Not IsSupportedExt(ext) Then
        errCode = "E0301"
        errDetail = "対応外の拡張子です: " & ext
        modLog.LogError "E0301", "modExtractor.ExtractFile", modUtil.SafeLeft(path, 500)
        ExtractFile = False
        Exit Function
    End If

    Dim maxPages As Long: maxPages = modConfig.GetLong("max_pages_per_file", 300)
    If maxPages < 1 Then maxPages = 300

    ' 2026-07-16 実機対策: ネットワーク共有(\\server\share\...)やクラウド同期
    ' フォルダ上のOfficeファイルは、実機WindowsのWord/Excelが「保護ビュー」で
    ' 開くため、COM自動化での本文抽出が失敗したり実行時エラー(添字範囲外等)に
    ' なることがある(LibreOfficeでは再現しない実機固有の挙動)。安全のため、
    ' Office系(pdf/docx/doc/xlsx/xls/xlsm)はローカル一時フォルダへコピーした
    ' コピーを開く。コピー失敗時は元パスのままフォールバックする(悪化させない)。
    Dim workPath As String: workPath = path
    Dim tmpCopy As String: tmpCopy = ""
    Select Case ext
        Case "pdf", "docx", "doc", "xlsx", "xls", "xlsm"
            tmpCopy = CopyToLocalTemp(path)
            If LenB(tmpCopy) > 0 Then workPath = tmpCopy
    End Select

    Dim ok As Boolean
    Dim truncated As Boolean
    Dim adapterErr As String

    Select Case ext
        Case "txt", "md", "csv"
            ok = ExtractPlainText(workPath, pages, adapterErr)
        Case "pdf"
            ok = ExtractPdfWithFallback(workPath, maxPages, pages, truncated, adapterErr)
        Case "docx", "doc"
            ok = modExtractorWord.Extract(workPath, maxPages, pages, truncated, adapterErr)
        Case "xlsx", "xls", "xlsm"
            ok = modExtractorExcel.Extract(workPath, maxPages, pages, truncated, adapterErr)
    End Select

    ' 一時コピーは用が済んだら消す(失敗しても無視)。
    If LenB(tmpCopy) > 0 Then
        On Error Resume Next
        Kill tmpCopy
        On Error GoTo 0
    End If

    If Not ok Then
        errCode = "E0302"
        errDetail = adapterErr
        modLog.LogError "E0302", "modExtractor.ExtractFile", modUtil.SafeLeft(path & " : " & adapterErr, 500)
        ExtractFile = False
        Exit Function
    End If

    Dim totalChars As Long: totalChars = SumPageChars(pages)
    Dim pageCount As Long: pageCount = PageArrayCount(pages)

    ' 画像PDF検知(PDFのみ・§7.2): 抽出総文字数 < ページ数×10
    If ext = "pdf" Then
        If totalChars < pageCount * 10 Then
            errCode = "E0303"
            errDetail = "画像PDFの可能性(総文字数=" & totalChars & " ページ数=" & pageCount & ")"
            modLog.LogError "E0303", "modExtractor.ExtractFile", modUtil.SafeLeft(path, 500)
            ExtractFile = False
            Exit Function
        End If
    End If

    If totalChars = 0 Then
        errCode = "E0304"
        errDetail = "抽出結果が空でした"
        modLog.LogError "E0304", "modExtractor.ExtractFile", modUtil.SafeLeft(path, 500)
        ExtractFile = False
        Exit Function
    End If

    If truncated Then
        errCode = "PARTIAL_PAGES"
        errDetail = "ページ数が上限(" & maxPages & ")を超えたため、先頭" & maxPages & "ページのみ取込みました"
    End If

    ExtractFile = True
End Function

Public Function SupportedExts() As String
    SupportedExts = SUPPORTED_EXTS
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

Private Function IsSupportedExt(ByVal ext As String) As Boolean
    IsSupportedExt = (InStr(1, "," & SUPPORTED_EXTS & ",", "," & ext & ",", vbTextCompare) > 0)
End Function

' path を %TEMP% 配下へコピーし、そのコピー先パスを返す(成功時)。
' コピーできなければ空文字列を返す(呼び出し側は元パスのまま処理を続ける)。
' 拡張子は必ず元と同じにする(抽出は拡張子で振り分けるため)。
' Word/Excelの保護ビュー(ネットワーク/クラウド上のファイルで発動)を避け、
' かつ抽出中の元ファイルロック・ネットワーク瞬断の影響を受けないようにする。
Private Function CopyToLocalTemp(ByVal path As String) As String
    On Error GoTo Fail

    Dim tempDir As String: tempDir = Environ$("TEMP")
    If LenB(tempDir) = 0 Then tempDir = Environ$("TMP")
    If LenB(tempDir) = 0 Then Exit Function
    If Right$(tempDir, 1) <> "\" Then tempDir = tempDir & "\"

    Dim baseName As String: baseName = modUtil.FileNameOf(path)
    If LenB(baseName) = 0 Then Exit Function

    ' 衝突回避のため連番を付ける(同時取込や前回の消し残りに備える)。
    Dim dest As String
    Dim n As Long: n = 0
    Do
        If n = 0 Then
            dest = tempDir & "mbtmp_" & baseName
        Else
            dest = tempDir & "mbtmp" & n & "_" & baseName
        End If
        If LenB(Dir$(dest)) = 0 Then Exit Do
        n = n + 1
        If n > 500 Then Exit Function   ' 異常時の暴走防止
    Loop

    FileCopy path, dest
    CopyToLocalTemp = dest
    Exit Function

Fail:
    CopyToLocalTemp = ""
End Function

' pdfはWordのPDF Reflowを優先し、失敗時のみAcrobat COMへフォールバックする
' (V1 modExtractor.ExtractPdfWithFallback を踏襲)。
Private Function ExtractPdfWithFallback(ByVal path As String, ByVal maxPages As Long, _
                                        ByRef pages() As ExtractedPage, ByRef truncated As Boolean, _
                                        ByRef errDetail As String) As Boolean
    Dim wordErr As String
    If modExtractorWord.Extract(path, maxPages, pages, truncated, wordErr) Then
        ExtractPdfWithFallback = True
        Exit Function
    End If

    Dim acroErr As String
    If modExtractorAcrobat.Extract(path, maxPages, pages, truncated, acroErr) Then
        ExtractPdfWithFallback = True
        Exit Function
    End If

    errDetail = "Word: " & wordErr & " / Acrobat: " & acroErr
    ExtractPdfWithFallback = False
End Function

' txt/md/csv はADODB.StreamでUTF-8として読み込む(常に1ページ扱い)。
Private Function ExtractPlainText(ByVal path As String, ByRef pages() As ExtractedPage, _
                                  ByRef errDetail As String) As Boolean
    Dim st As Object

    On Error GoTo Failed
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2          ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    Dim txt As String: txt = st.ReadText
    st.Close
    Set st = Nothing

    Dim tmp(0 To 0) As ExtractedPage
    tmp(0).page = 1
    tmp(0).Text = txt
    pages = tmp

    ExtractPlainText = True
    Exit Function

Failed:
    errDetail = DescribeComError(Err.Number, Err.Description)
    If Not st Is Nothing Then
        On Error Resume Next
        st.Close
        On Error GoTo 0
    End If
    ExtractPlainText = False
End Function

Private Function SumPageChars(pages() As ExtractedPage) As Long
    Dim n As Long: n = PageArrayCount(pages)
    If n = 0 Then Exit Function
    Dim lo As Long: lo = LBound(pages)
    Dim total As Long
    Dim i As Long
    For i = 0 To n - 1
        total = total + Len(pages(lo + i).Text)
    Next i
    SumPageChars = total
End Function

' 未初期化/空配列でも例外にせず0件として扱う(LBound/UBoundの実行時エラー9対策)。
Private Function PageArrayCount(pages() As ExtractedPage) As Long
    On Error GoTo Empty0
    PageArrayCount = UBound(pages) - LBound(pages) + 1
    Exit Function
Empty0:
    PageArrayCount = 0
End Function

' Mac等COM不可環境向けの丁寧な案内文を生成する(§13)。
' errNum=429は「ActiveX component can't create object」= COM未対応環境の定番エラー。
Private Function DescribeComError(ByVal errNum As Long, ByVal desc As String) As String
    If errNum = 429 Then
        DescribeComError = "この環境ではOffice連携(COM)が利用できません。" & _
            "Mac版ExcelやCOM未対応環境の可能性があります。Windows版Excelでお試しください。" & _
            "(詳細: " & desc & ")"
    Else
        DescribeComError = desc
    End If
End Function
