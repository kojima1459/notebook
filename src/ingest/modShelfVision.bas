Attribute VB_Name = "modShelfVision"
Option Explicit

' ============================================================================
' modShelfVision - 取込が失敗したときの vision フォールバックの一本化
' ----------------------------------------------------------------------------
' 役割:
'   modExtractor が抽出に失敗した資料を、画像解析(opt機能 vision)へ回すか
'   どうかの判断と、その結果の取り出しを引き受ける。modShelf.IngestFile は
'   TryVisionFallback を1回呼ぶだけでよくなる(§7.2の取込フロー本体を
'   これ以上太らせないための切り出し。2026-07-31 R6)。
'
'   受け持つのは2経路:
'     E0301(未対応形式)かつ拡張子が png/jpg/jpeg
'        → ExtractImagePdfText(画像1枚をそのまま文字起こし。裁定D13)
'     E0303(画像PDF)
'        → ExtractPdfOcrPagedText(Ghostscriptでページ毎にJPEG化してOCR。R6)
'          結果は modUtil.SplitPagedText でページ配列へ戻す。上限で打ち切られた
'          場合は truncated=True を返し、呼び出し元の partial 判定へ合流する。
'
' 設計判断:
'   ・opt モジュール名はここにも書かない(R2)。呼び出しは全て
'     modFeatures.InvokeFeature("vision", ...) の文字列経由。
'   ・失敗理由(outNote)は「利用者がその場で次の一手を打てる文」をそのまま
'     返す。Ghostscript未配置の案内など、opt側が持っている具体的な対処が
'     manifest の失敗メモまで届くようにするため、"#ERR:" の接頭辞だけを外して
'     本文を通す(modFeatures.InvokeFeature が元の文言を潰さなくなったのも
'     同じ目的。R6要件9)。
'   ・例外は外へ出さない。判断に失敗したら False を返し、呼び出し元は
'     従来どおり失敗として扱えばよい。
'
' v1で扱わないこと(拡張余地。2026-07-31 R6追補):
'   ・一部のページだけが文字化けするPDF(実機例: リコールマニュアル.pdf の
'     化けページ3件)は、modExtractor が該当ページを除外して続行するため、
'     ここには来ない(=そのページの内容は本棚に入らない)。
'     **v2: 除外されたページ番号だけを個別にOCRして本文へ差し込む**のが
'     自然な拡張で、そのときは modExtractor から「落としたページ番号」を
'     受け取り、optOcrCore の -dFirstPage/-dLastPage をページ単位で回せばよい。
'   ・全ページが化け判定になるPDFは modExtractor 側で E0303 に振り替えられ、
'     この関数のOCR経路に載る(R6追補1)。
' ============================================================================

Private Const VISION_FEATURE As String = "vision"

' ----------------------------------------------------------------------------
' TryVisionFallback - vision で救えるなら救う。
'   path      : 対象ファイルのフルパス
'   errCode   : modExtractor が返したエラーコード(E0301 / E0303 等)
'   pages     : 成功時に本文を格納する(ページ番号つき)
'   truncated : 上限ページで打ち切られた場合に True(呼び出し元は partial へ)
'   outNote   : 失敗時に manifest へ残す利用者向けの説明(空なら既定文を使う)
'   silent    : True=無人経路(フォルダ同期など)。誰も見ていない画面で
'               モーダルを出すと、そこで同期が朝まで止まる(R11-A C4)。
'               opt側へそのまま渡し、案内ダイアログを出さない解決だけを
'               させる。手動取込(False)では従来どおり案内を出す。
'   戻り値    : True=pagesに本文が入った / False=救えなかった
' ----------------------------------------------------------------------------
Public Function TryVisionFallback(ByVal path As String, ByVal errCode As String, _
                                  ByRef pages() As ExtractedPage, ByRef truncated As Boolean, _
                                  ByRef outNote As String, _
                                  Optional ByVal silent As Boolean = False) As Boolean
    outNote = ""
    Dim tmpCopy As String: tmpCopy = ""
    Dim okRet As Boolean: okRet = False

    Dim isImageFile As Boolean: isImageFile = IsImageFileExt(path)
    If errCode <> "E0303" And Not (errCode = "E0301" And isImageFile) Then Exit Function

    If Not modFeatures.FeatureEnabled(VISION_FEATURE) Then
        If isImageFile Then
            outNote = "画像の取込には画像解析機能の有効化が必要です。" & _
                "configシートの feature_vision を TRUE にしてから、もう一度お試しください。" & _
                "(コード: " & errCode & ")"
        Else
            outNote = "画像PDFの読み取り(OCR)は画像解析機能が有効なときだけ使えます。" & _
                "configシートの feature_vision を TRUE にし、Ghostscript を配置してから" & _
                "もう一度お試しください。(コード: " & errCode & ")"
        End If
        Exit Function
    End If

    On Error GoTo VisionFailed

    ' ------------------------------------------------------------------
    ' 2026-08-03(R13-3d): OCR/画像経路だけが安全な一時コピーを通っていなかった。
    ' modExtractor は Office系を必ず modExtractorPdf.CopyToLocalTemp 経由で
    ' 開く(R13-2でASCII安全名になった)のに、抽出に失敗してここへ落ちてくると
    ' modShelf は【元のパス】を渡してくる。その結果、Mac由来のNFD分解濁点
    ' (U+3099)等を含むファイル名がそのまま Ghostscript へ渡り、
    ' /undefinedfilename になっていた(実機第2報 RC4と同じ穴のOCR側)。
    ' ここでも安全なコピーを作り、抽出にはそれを使う。元のパスは
    ' 表示・ログ・メタデータ側でだけ使い続ける(利用者に見えるのは元の名前)。
    ' コピーできなくても「元を読めなかった」以外の理由(一時フォルダが無い等)
    ' なら従来どおり元パスで続行する(悪化させない)。
    ' 作ったコピーは【全ての出口】で消す(下の Finish が一手に引き受ける)。
    ' ------------------------------------------------------------------
    ' 2026-08-03(R14-3b): コピーが「元を読めていない」形で失敗したときは、
    ' 元パスのまま続行してはならない。NFD分解名のままGSへ渡ると
    ' /undefinedfilename で「描画0枚」になり、利用者には理由が何も残らない
    ' (実機第3報 RC3のOCR側)。正直に止めて、その場で打てる一手を伝える。
    Dim workPath As String: workPath = path
    Dim copyReason As String: copyReason = ""
    tmpCopy = modExtractorPdf.CopyToLocalTemp(path, copyReason)
    If LenB(tmpCopy) > 0 Then
        workPath = tmpCopy
    ElseIf modExtractorPdf.IsUnreadableCopyReason(copyReason) Then
        outNote = modLog.FriendlyFailMsg(errCode, _
            modExtractorPdf.CopyFailMsgFor(copyReason), modUtil.ExtOf(path))
        modLog.LogError errCode, "modShelfVision.TryVisionFallback", _
            "localcopy=failed:" & copyReason & " " & modUtil.SafeLeft(path, 200)
        GoTo Finish
    End If

    ' OCR経路(ExtractPdfOcrPagedText)だけは silent を渡す。Ghostscript が
    ' 見つからないときに opt 側が案内カード(モーダル)を出すのはこの経路
    ' だけで、無人の同期からそれが出ると同期がそこで止まる(R11-A C4)。
    Dim procName As String
    Dim callArgs As Variant
    If isImageFile Then
        procName = "ExtractImagePdfText"
        callArgs = workPath
    Else
        procName = "ExtractPdfOcrPagedText"
        callArgs = Array(workPath, silent)
    End If

    Dim raw As String
    raw = ResultToText(modFeatures.InvokeFeature(VISION_FEATURE, procName, callArgs))

    If LenB(Trim$(raw)) = 0 Or Left$(raw, 5) = "#ERR:" Then
        outNote = FailureNote(raw, errCode)
        modLog.LogError errCode, "modShelfVision.TryVisionFallback", _
            procName & " 失敗 " & modUtil.SafeLeft(path, 200) & " : " & modUtil.SafeLeft(raw, 300)
        GoTo Finish
    End If

    ' ページ付きで返ってきたら復号する。そうでなければ全文1ページとして扱う。
    If modUtil.SplitPagedText(raw, pages, truncated) Then
        okRet = True
        ' R14-4c: 上限で打ち切られたときは【何ページ入ったか】を正直に残す。
        ' 従来この partial は本棚カードで「同期を押すと続きから再開します」に
        ' なっていたが、OCRの続きを再開するロジックは存在しない(RC4の嘘)。
        If truncated Then outNote = OcrCapNote(pages)
        GoTo Finish
    End If

    truncated = False
    ReDim pages(0 To 0)
    pages(0).page = 1
    pages(0).Text = raw
    okRet = True
    GoTo Finish

VisionFailed:
    Dim failNum As Long: failNum = Err.Number
    Dim failDesc As String: failDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きたエラーは
    ' 呼び出し元へ飛んで本来の原因を上書きする。後始末の前に Resume で
    ' ハンドラを抜ける(modShelf.IngestFile と同じ作法)。
    Resume VisionCleanup
VisionCleanup:
    On Error Resume Next
    modLog.LogError errCode, "modShelfVision.TryVisionFallback", _
        "err#" & failNum & ": " & failDesc & " " & modUtil.SafeLeft(path, 200)
    If LenB(outNote) = 0 Then _
        outNote = modLog.FriendlyMessage(errCode) & "(コード: " & errCode & ")"
    On Error GoTo 0
    okRet = False
    ' Finish へ落ちる(一時コピーの後始末は1箇所だけに持つ)。

Finish:
    If LenB(tmpCopy) > 0 Then
        On Error Resume Next
        Kill tmpCopy
        On Error GoTo 0
    End If
    TryVisionFallback = okRet
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' ----------------------------------------------------------------------------
' OcrCapNote - 上限ページで打ち切られたときの正直なメモ(R14-4c)。
'   文言と算数は opt 層の純ロジック(optOcrCore.OcrCapMemoFor)が持ち、ここは
'   modFeatures 経由で受け取るだけ(R2: コアに opt モジュール名を書かない)。
'   総ページ数はGSに上限+1ページまでしか描かせていない以上こちらでは分からない
'   ので 0(不明)を渡す。取り込めた枚数だけは pages から確実に分かる。
'   取れなければ ""(メモが無い=従来どおりの partial 表示に落ちるだけ)。
' ----------------------------------------------------------------------------
Private Function OcrCapNote(ByRef pages() As ExtractedPage) As String
    On Error Resume Next
    Dim keptN As Long: keptN = modExtractor.PageArrayCount(pages)
    Dim note As String
    note = ResultToText(modFeatures.InvokeFeature(VISION_FEATURE, "OcrCapMemo", _
        Array(True, keptN, 0)))
    If Left$(note, 5) = "#ERR:" Then note = ""
    OcrCapNote = note
    On Error GoTo 0
End Function

' 画像拡張子か(vision委譲判定・裁定D13)。optVisionの対応形式と揃える。
Private Function IsImageFileExt(ByVal path As String) As Boolean
    Dim e As String
    e = LCase$(modUtil.ExtOf(path))
    IsImageFileExt = (e = "png" Or e = "jpg" Or e = "jpeg")
End Function

Private Function ResultToText(ByVal result As Variant) As String
    On Error GoTo NotText
    If IsError(result) Then Exit Function
    If VarType(result) <> vbString Then Exit Function
    ResultToText = CStr(result)
    Exit Function
NotText:
    ResultToText = ""
End Function

' opt側の失敗文字列を、そのまま利用者に見せられる1文へ整える。
'   "#ERR:E0303:<本文>" -> "<本文>(コード: E0303)"
'   "#ERR:FEATURE_UNAVAILABLE" のような機械的な文字列は既定文へ落とす。
Private Function FailureNote(ByVal raw As String, ByVal errCode As String) As String
    Dim body As String: body = raw
    If Left$(body, 5) = "#ERR:" Then body = Mid$(body, 6)
    If Left$(body, 6) = errCode & ":" Then body = Mid$(body, 7)
    body = Trim$(body)

    ' 文になっていない機械的な文字列(FEATURE_UNAVAILABLE等)は既定文へ落とす。
    If LenB(body) = 0 Or InStr(body, "。") = 0 Then body = modLog.FriendlyMessage(errCode)
    FailureNote = body & "(コード: " & errCode & ")"
End Function
