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
'   戻り値    : True=pagesに本文が入った / False=救えなかった
' ----------------------------------------------------------------------------
Public Function TryVisionFallback(ByVal path As String, ByVal errCode As String, _
                                  ByRef pages() As ExtractedPage, ByRef truncated As Boolean, _
                                  ByRef outNote As String) As Boolean
    outNote = ""

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

    Dim procName As String
    If isImageFile Then
        procName = "ExtractImagePdfText"
    Else
        procName = "ExtractPdfOcrPagedText"
    End If

    Dim raw As String
    raw = ResultToText(modFeatures.InvokeFeature(VISION_FEATURE, procName, path))

    If LenB(Trim$(raw)) = 0 Or Left$(raw, 5) = "#ERR:" Then
        outNote = FailureNote(raw, errCode)
        modLog.LogError errCode, "modShelfVision.TryVisionFallback", _
            procName & " 失敗 " & modUtil.SafeLeft(path, 200) & " : " & modUtil.SafeLeft(raw, 300)
        Exit Function
    End If

    ' ページ付きで返ってきたら復号する。そうでなければ全文1ページとして扱う。
    If modUtil.SplitPagedText(raw, pages, truncated) Then
        TryVisionFallback = True
        Exit Function
    End If

    truncated = False
    ReDim pages(0 To 0)
    pages(0).page = 1
    pages(0).Text = raw
    TryVisionFallback = True
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

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
