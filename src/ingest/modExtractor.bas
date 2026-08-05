Attribute VB_Name = "modExtractor"
Option Explicit

' ============================================================================
' modExtractor - 抽出ディスパッチャ(拡張子ごとに専門モジュールへ振り分け)
' ----------------------------------------------------------------------------
' 役割:
'   ファイルパスの拡張子を見て、専門モジュール(txt/md/csv=自前読込、
'   pdf=Ghostscript→Word→Acrobatフォールバック、docx/doc=Word、
'   xlsx/xls/xlsm=Excel)へ処理を振り分け、結果を ExtractedPage() の配列に
'   統一して返す。
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
'
' 2026-08-03(R13 Phase 0): 28,000字のWARN帯に達したため、PDFの3経路
'   フォールバック・ローカル一時コピー・文字化けページ除去を
'   modExtractorPdf へ純移設した。振り分けと判定(画像PDF/空/打ち切り)と
'   ページ配列のユーティリティはここに残る。
' ============================================================================

Private Const SUPPORTED_EXTS As String = "txt,md,csv,pdf,docx,doc,xlsx,xls,xlsm"

' txt/md/csv の入力上限50MB(R12-3-4。理由は ExtractPlainText の見出し)。
Private Const PLAINTEXT_MAX_BYTES As Double = 52428800#

' ----------------------------------------------------------------------------
' ExtractFile - 拡張子で振り分けて抽出する。
'   成功: True(pagesに1件以上、またはPARTIAL_PAGES時は打ち切り分を格納)
'   失敗: False + errCode(E0301/E0302/E0303/E0304) + errDetail(内部詳細)
' ----------------------------------------------------------------------------
Public Function ExtractFile(ByVal path As String, ByRef pages() As ExtractedPage, _
                            ByRef errCode As String, ByRef errDetail As String) As Boolean
    errCode = ""
    errDetail = ""
    ReDim pages(0 To 0)   ' 呼び出し前に必ず空配列へリセットしておく

    Dim ext As String: ext = modUtil.ExtOf(path)
    If Not IsSupportedExt(ext) Then
        errCode = "E0301"
        errDetail = "対応外の拡張子です: " & ext
        ' 2026-07-29: 画像(png/jpg/jpeg)は「対応外」で正しいが、呼び出し側
        ' (modShelf.IngestFile)がこの後 vision へ回して成功させる。
        ' 成功する経路の途中経過を err_log へエラーとして書くと、
        ' スクショを取り込むたびに err_log が伸び、本当の障害が埋もれる。
        ' 画像だけは usage_log 側へ回す(記録は残すが障害ではない)。
        If IsImageExt(ext) Then
            modLog.LogUsage "extract_to_vision", "", _
                "画像なので画像解析へ回します: " & modUtil.SafeLeft(modUtil.FileNameOf(path), 200)
        Else
            modLog.LogError "E0301", "modExtractor.ExtractFile", modUtil.SafeLeft(path, 500)
        End If
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
    ' 実機報告(2026-07-21)「取込に失敗(#462等)」対策: 各アダプタは自分の
    ' COM呼び出しを捕捉する契約だが、抜けがあれば生の実行時エラーが素通りして
    ' しまう(呼び出し元へ意味不明な番号だけが伝わる)。ここでも保険的に捕捉し、
    ' 必ずE0302+説明文に変換してから返す(R5: 原因不明のエラーを見せない)。
    ' 2026-07-22実機再発: 前回のトラップはSelect Case本体だけを覆っており、
    ' その直前のCopyToLocalTemp呼び出しがトラップ外だったため、そちら側で
    ' 何か漏れれば依然としてExtractFile自体が未捕捉のまま呼び出し元
    ' (modShelf.IngestFile)まで抜けてしまう隙間があった。関数全体を覆うように
    ' ここへ引き上げる。
    On Error GoTo ExtractFailed

    Dim workPath As String: workPath = path
    Dim tmpCopy As String: tmpCopy = ""
    Dim triedLocalCopy As Boolean: triedLocalCopy = False
    Dim copyReason As String: copyReason = ""
    Select Case ext
        Case "xlsx", "xls", "xlsm"
            triedLocalCopy = True
            tmpCopy = modExtractorPdf.CopyToLocalTemp(path, copyReason)
            If LenB(tmpCopy) > 0 Then
                workPath = tmpCopy
            Else
                ' 2026-07-28(レビュー H-11)→2026-07-30(要件D)で対象をExcel系
                ' だけへ絞った。危険が実在するのは「同一Excelインスタンスで
                ' 開かれるxls/xlsx/xlsm」だけである。コピーに失敗したまま
                ' 原本を開くと、Excelは既存のブックオブジェクトを返してしまい、
                ' 抽出後の wb.Close False が【利用者が編集中のブックを破棄して
                ' 閉じる】ことになる。取り込めないと伝える方が、他人の作業を
                ' 消すよりましである(pdf/doc/docxはこの危険が無いので下のCase
                ' で原本フォールバックを許す。H-11の過剰一般化を正した)。
                errCode = "E0302"
                errDetail = "ファイルを一時フォルダへコピーできませんでした。" & _
                    "そのファイルを開いている場合は閉じてから、もう一度お試しください"
                modLog.LogError "E0302", "modExtractor.ExtractFile", _
                    modUtil.SafeLeft(path & " : localcopy=failed:" & copyReason & _
                    "(原本フォールバックは行わない)", 500)
                ExtractFile = False
                Exit Function
            End If
        Case "pdf", "docx", "doc"
            triedLocalCopy = True
            tmpCopy = modExtractorPdf.CopyToLocalTemp(path, copyReason)
            If LenB(tmpCopy) > 0 Then
                workPath = tmpCopy
            ElseIf ext = "pdf" And modExtractorPdf.IsUnreadableCopyReason(copyReason) Then
                ' 2026-08-03(R14-3b・実機第3報 RC3): コピーが「元を読めていない」
                ' 形で失敗したPDFを、元パスのままGhostscriptへ渡してはならない。
                ' NFD分解濁点等を含む共有上の名前はGSに /undefinedfilename で
                ' 拒否され、結果は「描画0枚」だけ。利用者には理由が1文字も
                ' 残らず、しかも localcopy=ok と嘘が記録されていた。
                ' ここで正直に止め、その場で打てる一手(名前を変える)を伝える。
                errCode = "E0302"
                errDetail = modExtractorPdf.CopyFailMsgFor(copyReason)
                modLog.LogError "E0302", "modExtractor.ExtractFile", _
                    modUtil.SafeLeft(path & " : localcopy=failed:" & copyReason, 500)
                ExtractFile = False
                Exit Function
            Else
                ' 2026-07-30(要件D): Word COMはReadOnlyで開くだけで、Excelと
                ' 違って利用者の編集セッションに触れる経路が無い(既存インス
                ' タンスのブックを掴んで閉じる、という事故が起きない)。
                ' コピーが失敗する主因は「そのファイルが今そのPCで開かれて
                ' いる(ロック中)」で、「Wordを開いたまま使う」運用の回避策が
                ' 定着したため頻発する(実機ログ: 悪天候の定義.docx :
                ' localcopy=failed)。取り込めないより原本で読める方が良いので、
                ' 原本パスのまま続行する。次回の切り分け用にログへ残す。
                workPath = path
                On Error Resume Next
                modLog.LogUsage "extract_localcopy_fallback", "", _
                    modUtil.SafeLeft(path & " : localcopy=failed:" & copyReason & _
                    "(原本で続行)", 500)
                On Error GoTo ExtractFailed
            End If
    End Select

    ' 診断用: ローカル一時コピーの成否を憶えておく(失敗時のerrDetailへ含め、
    ' 「コピー自体が失敗してUNCパスのままWordへ渡った」のか「コピーは成功
    ' したのにローカルコピーでも失敗した」のかを次回のログで切り分ける)。
    ' R14-3a: 成功は【サイズ突合まで通った】ことを意味する(verified)。
    ' 旧実装は0バイトの嘘コピーでも localcopy=ok と書いていた(RC3)。
    Dim copyNote As String
    If triedLocalCopy Then
        If LenB(tmpCopy) > 0 Then
            copyNote = "localcopy=ok(verified)"
        Else
            copyNote = "localcopy=failed:" & copyReason
        End If
    Else
        copyNote = "localcopy=n/a"
    End If

    Dim ok As Boolean
    Dim truncated As Boolean
    Dim adapterErr As String
    ' R10-3: PDF経路だけが「E0302ではなくE0303で返したい」場合がある
    ' (Ghostscriptが『文字がほとんど入っていない』と判定した=画像PDF)。
    Dim pdfRouteCode As String

    Select Case ext
        Case "txt", "md", "csv"
            ok = ExtractPlainText(workPath, pages, adapterErr)
        Case "pdf"
            ' 原本パスも渡す(レビュー4-D)。Word側が「利用者が編集中の文書を
            ' 掴んでいないか」を、一時コピー・原本の両方のパスで調べる。
            ok = modExtractorPdf.ExtractPdfWithFallback(workPath, maxPages, pages, truncated, _
                                                       adapterErr, pdfRouteCode, path)
        Case "docx", "doc"
            ok = modExtractorWord.Extract(workPath, maxPages, pages, truncated, adapterErr, path)
        Case "xlsx", "xls", "xlsm"
            ok = modExtractorExcel.Extract(workPath, maxPages, pages, truncated, adapterErr)
    End Select

    ' 一時コピーは用が済んだら消す(失敗しても無視)。GoTo ExtractFailedで
    ' 関数全体のトラップを維持したまま一時的にResume Nextへ切り替える。
    If LenB(tmpCopy) > 0 Then
        On Error Resume Next
        Kill tmpCopy
        On Error GoTo ExtractFailed
    End If

    If Not ok Then
        errCode = "E0302"
        ' R10-3: GS(txtwrite)が画像PDFと判定した場合だけE0303へ振り替え、
        ' modShelfVision のOCR経路へ回す。それ以外は従来どおりE0302。
        If LenB(pdfRouteCode) > 0 Then errCode = pdfRouteCode
        errDetail = adapterErr & " [" & copyNote & "]"
        ' R18-10(是正・実機第5報): 実機のerr_logに赤字で並んでいたのはこの
        ' 「gs_image:」経路(GSが文字層ゼロと即断して画像PDFと分類した)で、
        ' thin:経路と同じく【失敗ではない】(OCRフォールバックへ回すだけの
        ' 正常な分類)。usage_logへ格下げする。3経路全滅の allfail: は
        ' Word/Acrobatまで失敗した本物の失敗なのでerr_logのまま。
        If Left$(adapterErr, 9) = "gs_image:" Then
            ' R18H FB-1(A-L8): detail はファイル名ではなくパス全体(200字)。
            ' 同名の資料が別フォルダに複数あると、どれの話か特定できなかった。
            modLog.LogUsage "image_pdf_detected", "", _
                "gs_image: " & modUtil.SafeLeft(path, 200)
        Else
            modLog.LogError errCode, "modExtractor.ExtractFile", modUtil.SafeLeft(path & " : " & errDetail, 500)
        End If
        ExtractFile = False
        Exit Function
    End If

    ' 文字化けページの除去(2026-07-27追加。2026-07-30 要件Aで配列圧縮化)。
    '
    ' 実物の約款PDFで確認: 表紙・裏の連絡先ページが装飾用の埋め込みフォント
    ' (ToUnicodeマップ無し)で作られていると、抽出結果が制御文字とギリシャ/
    ' キリル文字の羅列になる。例) "ঝೝʣ(/$ʢʣ ϛη"
    ' 本文は正常なので既存の画像PDF判定(総文字数)には引っかからず、化けた行が
    ' そのままチャンクになり、プロンプトに載り、検索の邪魔をする。
    ' 「答えが的外れ」の原因として最悪の部類で、しかも誰にも見えない。
    ' ページ単位で落とせば、本文は1文字も失わずにノイズだけ消える。
    '
    ' 2026-07-30実機: 旧実装は化けページの.Textを""にするだけで配列を詰め
    ' なかった。実機ログで extract_garbled N件 と chunk_skipped_pages N件が
    ' 全6ペアで完全一致しており、空文字ページが後段でerr#9を出しページ単位
    ' スキップに化けていた。1ページしかない.docだと全文が消えtotalChars=0→
    ' E0304「抽出結果が空」で取込不能になっていた(実機:【個別特約】興行中止
    ' 保険特約(関係者読替).doc)。DropGarbledPages側で配列そのものを詰め、
    ' 空文字ページを後段へ渡さないようにする。
    Dim droppedPages As Long
    Dim allGarbled As Boolean
    droppedPages = modExtractorPdf.DropGarbledPages(pages, allGarbled)
    If droppedPages > 0 Then
        On Error Resume Next
        modLog.LogUsage "extract_garbled", "", _
            modUtil.SafeLeft(modUtil.FileNameOf(path), 120) & " 化けページ" & droppedPages & "件を除外"
        On Error GoTo ExtractFailed
    End If

    ' 2026-07-31(R6追補): 全ページが化け判定になるPDFは、実機ログを見ると
    ' 「WordのリフローがゴミWord文字しか返せていない画像PDF」だった
    ' (総文字数は多いので既存のE0303判定=文字数の少なさには引っかからない)。
    ' 化けたまま取り込んでも検索の邪魔にしかならないので、PDFに限り
    ' E0303として返し、OCRフォールバック(modShelfVision)へ回す。
    ' PDF以外(doc/docx/xls系)はGhostscriptで画像化できないため従来どおり続行。
    If allGarbled Then
        Dim routeCode As String
        routeCode = GarbledRouteCode(ext, allGarbled)
        On Error Resume Next
        If LenB(routeCode) > 0 Then
            modLog.LogUsage "extract_garbled_ocr_route", "", _
                modUtil.SafeLeft(modUtil.FileNameOf(path), 120) & _
                " 全" & PageArrayCount(pages) & "件が化け判定のためOCR経路へ回します"
        Else
            modLog.LogUsage "extract_garbled_kept", "", _
                modUtil.SafeLeft(modUtil.FileNameOf(path), 120) & _
                " 全" & PageArrayCount(pages) & "件が化け判定のため除外せず続行"
        End If
        On Error GoTo ExtractFailed
        If LenB(routeCode) > 0 Then
            errCode = routeCode
            errDetail = "全ページが文字化け判定(Wordリフロー不能の画像PDFとみなす)"
            modLog.LogError routeCode, "modExtractor.ExtractFile", modUtil.SafeLeft(path, 500)
            ExtractFile = False
            Exit Function
        End If
    End If

    Dim totalChars As Long: totalChars = SumPageChars(pages)
    Dim pageCount As Long: pageCount = PageArrayCount(pages)

    ' 画像PDF検知(PDFのみ・§7.2): 抽出総文字数 < ページ数×10
    If ext = "pdf" Then
        If totalChars < pageCount * 10 Then
            ' R13-3c: E0303は3箇所から同じコードで発報される。detailの先頭に
            ' 発報点を書き、診断者が「GSの即断」と取り違えないようにする。
            errCode = "E0303"
            errDetail = "thin: 画像PDFの可能性(総文字数=" & totalChars & " ページ数=" & pageCount & ")"
            ' R18-10(実機第5報⑧・agent1調査): この分類は失敗ではない
            ' (抽出自体は成功しており、OCRフォールバックへ回すだけの正常な
            ' 判定)。err_logの赤字は「利用者のエラー一覧」に並んでしまうため、
            ' usage_logへ格下げする。真の失敗(OCR経路のTryVisionFallback失敗
            ' 等)は従来どおりerr_log(呼び出し元modShelfVision側で別途発報)。
            ' R18H FB-1(A-L8): detail はパス全体(200字)。上の gs_image: と同じ。
            modLog.LogUsage "image_pdf_detected", "", _
                "thin: " & modUtil.SafeLeft(path, 200)
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
    Exit Function

ExtractFailed:
    Dim leakNum As Long, leakDesc As String
    leakNum = Err.Number: leakDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume ExtractFailedCleanup0
ExtractFailedCleanup0:
    On Error Resume Next
    If LenB(tmpCopy) > 0 Then Kill tmpCopy
    On Error GoTo 0
    If LenB(copyNote) = 0 Then copyNote = "localcopy=n/a(コピー処理到達前に失敗)"
    errCode = "E0302"
    errDetail = "err#" & leakNum & ": " & leakDesc & " [" & copyNote & "]"
    modLog.LogError "E0302", "modExtractor.ExtractFile", modUtil.SafeLeft(path & " : " & errDetail, 500)
    ExtractFile = False
End Function

Public Function SupportedExts() As String
    SupportedExts = SUPPORTED_EXTS
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' 画像の拡張子か(vision へ回す対象。modShelf.IsImageExtension と同じ判定)。
Private Function IsImageExt(ByVal ext As String) As Boolean
    Dim e As String: e = LCase$(ext)
    IsImageExt = (e = "png" Or e = "jpg" Or e = "jpeg")
End Function

Private Function IsSupportedExt(ByVal ext As String) As Boolean
    IsSupportedExt = (InStr(1, "," & SUPPORTED_EXTS & ",", "," & ext & ",", vbTextCompare) > 0)
End Function

' ----------------------------------------------------------------------------
' GarbledRouteCode - 「全ページが化け判定」だったときの行き先(純ロジック)。
'   2026-07-31 R6追補。実機ログでは画像PDFの典型症状が「E0303(文字が少ない)」
'   ではなく「WordリフローがゴミWord文字を返し、全ページが化け判定になる」形
'   だった。この形はPDFならOCRで救えるので E0303 を返してフォールバックを
'   発火させる。doc/docx/xls系はGhostscriptで画像化できない(PDF専用)ため
'   従来どおり "" =化けたまま続行にする。
'   分岐だけを純関数として切り出し、modTestsPure4で固定する。
' ----------------------------------------------------------------------------
Public Function GarbledRouteCode(ByVal ext As String, ByVal allGarbled As Boolean) As String
    If Not allGarbled Then Exit Function
    If LCase$(Trim$(ext)) = "pdf" Then GarbledRouteCode = "E0303"
End Function

' 2026-08-03(R14-F13): SharedCopyNextChunkLen をここから削除した。
' R14-3a で共有読みコピーが ADODB.Stream 一本になった時点で呼び出し元が
' 消え、契約とテストだけが残っていた。R14-F3 でクラシックの1MB分割コピーを
' 復活させたが、分割の算数は modExtractorPdf.CopyClassicShared のループに
' 3行で書いてある(モジュールを跨いだPublicを1つ減らす方を採った)。

' 空白類しか無いページか。Trim$ は半角スペースしか落とさないため使わない
' (改行だけのページを中身ありと数えると出典ページ番号が全部1つずれる)。
' optOcrCore.CleanTextLen と同じ規約(R2で2箇所に分かれる。突合はmodTestsPure6)。
Private Function IsBlankGsPage(ByVal s As String) As Boolean
    IsBlankGsPage = modUtilText.GsPageIsBlank(s)
End Function

' ----------------------------------------------------------------------------
' BuildPagesFromGsText - Ghostscript(txtwrite)の出力を ExtractedPage 配列へ。
'   txtwriteはページの区切りに改ページ文字 Chr(12) を挟むので、そこで割れば
'   Word/Acrobat経路と同じページ番号つきで取り込める(出典表示のため
'   ページ番号は落とせない)。上限ページを超えた場合は先頭maxPagesだけ残して
'   truncated=True にする(他アダプタと同じ規約)。中身が空なら False を返し、
'   呼び出し元は Word → Acrobat の連鎖へ落ちる。
'
'   R10c(L1): 空ページの読み飛ばしを先頭側・末尾側で対称にした。従来は末尾しか
'   見ておらず、先頭が改ページで始まるPDFでは1ページ目が空になり、以降の出典
'   ページ番号が丸ごと1つズレていた(「出典 p.5」で別ページが開く壊れ方)。
'   Public なのは modTestsPure6 から境界を検証するため(lintのCONTRACTに明記)。
' ----------------------------------------------------------------------------
Public Function BuildPagesFromGsText(ByVal txt As String, ByVal maxPages As Long, _
                                     ByRef pages() As ExtractedPage, _
                                     ByRef truncated As Boolean) As Boolean
    Dim parts() As String
    Dim firstIdx As Long
    Dim lastIdx As Long
    Dim keptN As Long
    Dim i As Long

    truncated = False
    If IsBlankGsPage(txt) Then Exit Function

    ' 先頭側・末尾側の空ページ(空白類だけの要素)を対称に読み飛ばす添字計算は
    ' modUtilText.GsPageBounds が唯一の実装(2026-07-31 R11-F2)。従来は
    ' optOcrCore.GsPageCount と同じ規約を2箇所で書いており、ズレると出典ページ
    ' 番号が丸ごとずれる=「出典 p.5 を開くと別のページが出る」壊れ方をした。
    keptN = modUtilText.GsPageBounds(txt, firstIdx, lastIdx)
    If keptN < 1 Then Exit Function
    parts = Split(txt, Chr$(12))
    If maxPages > 0 And keptN > maxPages Then
        keptN = maxPages
        truncated = True
    End If

    ' 受け取った配列を直接ReDimする(一時配列からの代入はLO実行テストで
    ' ユーザー定義型の配列代入が420になるため使わない。全件コピーも省ける)。
    '
    ' R12-3-6: ページ番号は【物理ページ】。txtwriteは1ページごとに末尾へ改ページを
    ' 出すので、0起点のSplit添字がそのまま物理ページになる。従来は読み飛ばした先頭の
    ' 空ページを詰めてi+1を振り直しており、表紙が画像だけのPDF(保険のパンフレット・
    ' 装飾表紙の約款)で全チャンクのページ番号が表紙の枚数ぶん手前にずれた=
    ' 「出典 p.5 を開くと別ページ」。OCR経路は元から物理番号。
    ReDim pages(0 To keptN - 1)
    For i = 0 To keptN - 1
        pages(i).page = firstIdx + i + 1
        pages(i).Text = parts(firstIdx + i)
    Next i

    BuildPagesFromGsText = True
End Function

' txt/md/csv はADODB.StreamでUTF-8として読み込む(常に1ページ扱い)。
'
' R12-3-4: 読む前にサイズを見る。txt/md/csvは「1ファイル=全文1ページ」で丸ごと
' メモリに載せるため、上限無しだと32bitExcelは巨大ファイル1本でerr7や無応答に陥る。
' 力尽きるのは読込の後(チャンク化・埋込)で原因が見えない。入口で明示的に止める。
Private Function ExtractPlainText(ByVal path As String, ByRef pages() As ExtractedPage, _
                                  ByRef errDetail As String) As Boolean
    Dim txt As String
    Dim readErrNum As Long
    Dim readErrDesc As String

    Dim sizeBytes As Double
    On Error Resume Next
    sizeBytes = CDbl(FileLen(path))
    On Error GoTo 0
    If sizeBytes > PLAINTEXT_MAX_BYTES Then
        errDetail = "ファイルが大きすぎます(上限50MB)。分割してお試しください。" & _
                    "(" & modUtil.HumanBytes(sizeBytes) & ")"
        Exit Function
    End If

    ' UTF-8読み取りの実体は modUtilText.ReadTextFileUtf8(2026-07-31 R11-F2で
    ' 10箇所の同型実装を1本化)。失敗時の利用者向け文言(DescribeComError)は
    ' ここが従来どおり組み立てる=挙動は変えない。
    If Not modUtilText.ReadTextFileUtf8(path, txt, readErrNum, readErrDesc) Then
        errDetail = modUtil.DescribeComError(readErrNum, readErrDesc, "Office")
        Exit Function
    End If

    Dim tmp(0 To 0) As ExtractedPage
    tmp(0).page = 1
    tmp(0).Text = txt
    pages = tmp

    ExtractPlainText = True
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
' Public なのは modExtractorPdf.DropGarbledPages が同じ数え方を使うため
' (2026-08-03 Phase 0 の分割で2箇所に同じ実装を置かないための公開)。
Public Function PageArrayCount(pages() As ExtractedPage) As Long
    On Error GoTo Empty0
    PageArrayCount = UBound(pages) - LBound(pages) + 1
    Exit Function
Empty0:
    PageArrayCount = 0
End Function

