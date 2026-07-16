Attribute VB_Name = "optVision"
Option Explicit

' ============================================================================
' optVision - 画像ファイルの文字抽出(opt機能・MASTER_SPEC §7.7)
' ----------------------------------------------------------------------------
' 役割:
'   通常抽出(modExtractor)がE0303(画像PDF・文字が取れない)を返した資料や、
'   画像ファイル(png/jpg/jpeg)そのものについて、AIリボンのVision機能
'   (ChatGPTV・確定済み)へ渡して文字起こしする。この層が無くても製品は
'   成立し、modules.jsonから1行削除するだけで撤去できる(§7.7)。
'
' 設計判断:
'   ・対応形式は画像ファイル(png/jpg/jpeg)のみ。PDFの直渡し
'     (旧config vision_pdf_direct 路線)は、ChatGPTVのimageInputsが
'     「Base64エンコードした画像」に限定されると確定した
'     (出典: RIBBON_API_CONFIRMED.md §1 #5)ため完全撤去した。
'     PDFが渡された場合は丁寧な案内文を含む失敗として扱う。
'   ・抽出経路は2段階の確定規約:
'       (1) Base64FromFile(filePath) で画像ファイルをBase64文字列化
'       (2) ChatGPTV(Text, imageInputs, roleSystem, resolution, toolN) で文字起こし
'     どちらも modGateway.TryRibbonRun のみ経由で呼ぶ(R3)。
'   ・失敗しても例外を外に出さない。ExtractImagePdfはBoolean(§7.7契約の
'     とおり)、ExtractImagePdfTextは "#ERR:..." 文字列 または 全文。
'   ・スクリーンショット取込(裁定D13): クリップボードの画像を本棚へ取り込む
'     導線のために HasClipboardImage / SaveClipboardImage を公開する。
'     確定関数 IsImageInCB / Base64FromCB(Ptn=1) の唯一の利用箇所
'     (呼び出しはTryRibbonRun経由・R3)。UI導線はmodUIShelf.OnIngestScreenshot。
'   ・mock_llm=TRUE のときはリボンを呼ばない: HasClipboardImageはFalse、
'     SaveClipboardImageは親切な "#ERR:mockモード…" 案内を返す(optMarkdownの
'     mock時の流儀に合わせる)。
'   ・ExtractImagePdfの失敗理由をExtractImagePdfTextへ伝えるため、
'     モジュール内Private変数(mLastErrorMsg)に直近の失敗メッセージを
'     保持する(VBAは単一スレッドで実行されるため競合の心配がない)。
'   ・modUIMain.SetStage 以外のコア参照は行わない(基盤層のみ参照可・§7.7)。
' ============================================================================

' === 確定済みシグネチャ(出典: RIBBON_API_CONFIRMED.md §1 #5, #8 / 裁定D5) ===
'   Base64FromFile(filePath As String) -> String
'     PNG/JPEG等の画像ファイルをそのままBase64文字列にして返す(#8)。
'   ChatGPTV(Text, imageInputs, [roleSystem], [resolution], [toolN]) -> String
'     imageInputs = Base64FromFileで得たBase64文字列(複数はカンマ区切り、
'     最大10枚)。resolution="high" で高精細解析(#5)。
'     本モジュールは (VISION_PROMPT, b64, "", "high", "マイ本棚AI:vision") で呼ぶ。
'   戻り値の扱い(§0: ラッパーはAPIレスポンスをそのまま文字列で返す):
'     失敗時は "#ERR:" 始まり(TryRibbonRun側)/空文字列/上限系メッセージが
'     返り得るため、IsVisionError() で失敗と判定する。
'   PDF直渡しは公式仕様上不可能と確定(imageInputsは画像のBase64限定)。
'   PDF→画像変換の手段はVBA単体に無いため、PDFは案内文つきの失敗とする。
'   IsImageInCB() -> Boolean(台帳§1 #6)
'     クリップボードに画像があるかどうかを即時判定する。
'   Base64FromCB([Ptn=0]) -> String(台帳§1 #7)
'     Ptn=0でBase64文字列、Ptn=1で「Tempに保存したjpgのパス」を返す。
'     本モジュール(SaveClipboardImage)は Ptn=1 のjpgパス取得のみ使う。
' ============================================================================

Private Const RIBBON_FUNC_NAME As String = "ChatGPTV"
Private Const BASE64_FUNC_NAME As String = "Base64FromFile"
Private Const VISION_RESOLUTION As String = "high"
Private Const VISION_TOOL_NAME As String = "マイ本棚AI:vision"
Private Const IMAGE_EXTS As String = "png,jpg,jpeg"
' 2026-07-16 実機フィードバック(「どんなスクショでも1チャンクにしかならず
' 精度が悪い」)を受けて逐語書き起こしの指示を強化: 表の構造・数値・単位の
' 厳密な転記と、見出しによる構造化を明示的に要求する(書き起こしが厚くなる
' ほどチャンク分割・検索の材料が増える)。
Private Const VISION_PROMPT As String = _
    "この画像に写っている文字情報を、一字一句省略せずにすべて書き起こしてください。" & vbLf & _
    "・見出しやセクション名は行頭に「# 」を付けて構造を保つこと" & vbLf & _
    "・表は1行=1レコードの形で、列名と値の対応が分かるように書き出すこと" & vbLf & _
    "・数値・金額・日付・単位・記号は画像のとおり正確に転記すること(要約や丸めは禁止)" & vbLf & _
    "・図やグラフは、読み取れる軸・凡例・数値を含めて内容を文章で説明すること" & vbLf & _
    "・画像に無い情報を補って書かないこと"

Private mLastErrorMsg As String

Public Function Ping() As Boolean
    Ping = True
End Function

' ----------------------------------------------------------------------------
' HasClipboardImage - クリップボードに画像があるか(裁定D13)。
'   確定関数 IsImageInCB()(台帳§1 #6)のラッパー。mock_llm=TRUE・リボン
'   不在・失敗時はFalse(失敗の記録はTryRibbonRun側がE0202で行うため、
'   ここでの追加ログは不要=握りつぶしではない)。
' ----------------------------------------------------------------------------
Public Function HasClipboardImage() As Boolean
    HasClipboardImage = False
    If modConfig.GetBool("mock_llm", True) Then Exit Function
    Dim res As Variant
    res = modGateway.TryRibbonRun("IsImageInCB", Array())
    If VarType(res) = vbBoolean Then HasClipboardImage = CBool(res)
End Function

' ----------------------------------------------------------------------------
' SaveClipboardImage - クリップボードの画像をTempへjpg保存しパスを返す
'   (裁定D13)。確定関数 Base64FromCB(Ptn=1)(台帳§1 #7: Ptn=1で
'   「Tempに保存したjpgのパス」が返る)のラッパー。
'   成功: jpgのフルパス / 失敗: "#ERR:..."(例外は出さない)。
' ----------------------------------------------------------------------------
Public Function SaveClipboardImage() As String
    If modConfig.GetBool("mock_llm", True) Then
        SaveClipboardImage = "#ERR:mockモード(mock_llm=TRUE)ではクリップボード取込を利用できません。" & _
            "config の mock_llm を FALSE にすると実際に取り込めるようになります。"
        Exit Function
    End If

    Dim res As Variant
    res = modGateway.TryRibbonRun("Base64FromCB", Array(1&))
    Dim s As String: s = CStr(res)

    If LenB(s) = 0 Then
        SaveClipboardImage = "#ERR:クリップボード画像の保存に失敗しました(応答が空)"
    Else
        ' TryRibbonRun失敗時の "#ERR:E0202:..." はそのまま呼び出し元へ渡す
        SaveClipboardImage = s
    End If
End Function

' ----------------------------------------------------------------------------
' ExtractImagePdf - path(画像ファイル)からVisionで文字を読み取り、
'   ExtractedPage 1件(page=1、全文)としてpagesへ格納する(MASTER_SPEC §7.7)。
'   成功: True(pagesに1件) / 失敗: False(pagesは空配列のまま)。
'   ※関数名の「Pdf」は§7.7契約の歴史的経緯によるもの。実際の対応形式は
'     png/jpg/jpeg のみで、PDFは案内文つきの失敗になる(裁定D5)。
' ----------------------------------------------------------------------------
Public Function ExtractImagePdf(ByVal path As String, ByRef pages() As ExtractedPage) As Boolean
    mLastErrorMsg = ""
    ReDim pages(0 To 0)

    On Error GoTo Fail

    Dim ext As String
    ext = modUtil.ExtOf(path)

    If Not IsImageExt(ext) Then
        If ext = "pdf" Then
            mLastErrorMsg = "#ERR:E0303:" & modLog.FriendlyMessage("E0303") & _
                "(社内AIリボンの画像解析はPDFを直接読み取れない仕様です。お手数ですが、" & _
                "該当ページをスクリーンショット等で画像(png/jpg)にしてからお試しください。)"
        Else
            mLastErrorMsg = "#ERR:E0301:このファイル形式(" & ext & ")はVisionで直接読み取れません。" & _
                "対応形式は画像(png/jpg/jpeg)のみです。"
        End If
        modLog.LogError "E0303", "optVision.ExtractImagePdf", "未対応ext=" & ext & " path=" & modUtil.SafeLeft(path, 300)
        ExtractImagePdf = False
        Exit Function
    End If

    modUIMain.SetStage "🖼 画像から文字を読み取っています…"

    ' (1) 画像ファイル → Base64文字列(確定: Base64FromFile / 台帳§1 #8)
    Dim b64 As String
    b64 = SafeResultToString(modGateway.TryRibbonRun(BASE64_FUNC_NAME, Array(path)))

    If LenB(Trim$(b64)) = 0 Or Left$(b64, 5) = "#ERR:" Then
        modUIMain.SetStage ""
        mLastErrorMsg = "#ERR:E0303:画像ファイルの読み込み(Base64変換)に失敗しました。" & _
            "ファイルが開けるか、壊れていないかをご確認ください。" & _
            IIf(LenB(b64) > 0, "(詳細: " & modUtil.SafeLeft(b64, 200) & ")", "")
        modLog.LogError "E0303", "optVision.ExtractImagePdf", "Base64FromFile失敗 path=" & _
            modUtil.SafeLeft(path, 300) & " : " & modUtil.SafeLeft(b64, 200)
        ExtractImagePdf = False
        Exit Function
    End If

    ' (2) Base64 → 文字起こし(確定: ChatGPTV / 台帳§1 #5・裁定D5)
    Dim result As Variant
    result = modGateway.TryRibbonRun(RIBBON_FUNC_NAME, _
        Array(VISION_PROMPT, b64, "", VISION_RESOLUTION, VISION_TOOL_NAME))

    modUIMain.SetStage ""

    Dim s As String
    s = SafeResultToString(result)

    If IsVisionError(s) Then
        mLastErrorMsg = "#ERR:E0303:文字の読み取りに失敗しました。" & _
            "画像が不鮮明か、対応外の内容の可能性があります。" & _
            IIf(LenB(s) > 0, "(詳細: " & modUtil.SafeLeft(s, 200) & ")", "")
        modLog.LogError "E0303", "optVision.ExtractImagePdf", modUtil.SafeLeft(path & " : " & s, 300)
        ExtractImagePdf = False
        Exit Function
    End If

    Dim tmp(0 To 0) As ExtractedPage
    tmp(0).page = 1
    tmp(0).Text = s
    pages = tmp

    ExtractImagePdf = True
    Exit Function

Fail:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    On Error GoTo 0
    modUIMain.SetStage ""
    modLog.LogError "E0303", "optVision.ExtractImagePdf", errDesc
    mLastErrorMsg = "#ERR:E0303:文字の読み取り処理でエラーが発生しました: " & errDesc
    ExtractImagePdf = False
End Function

' ----------------------------------------------------------------------------
' ExtractImagePdfText - ExtractImagePdfのText版ラッパ(MASTER_SPEC §7.7)。
'   modFeatures.InvokeFeature("vision","ExtractImagePdfText", path) から
'   文字列一枚で呼べるようにするための契約(戻り値: 全文 または "#ERR:..." )。
' ----------------------------------------------------------------------------
Public Function ExtractImagePdfText(ByVal path As String) As String
    Dim pages() As ExtractedPage
    If ExtractImagePdf(path, pages) Then
        ExtractImagePdfText = JoinPagesText(pages)
    ElseIf LenB(mLastErrorMsg) > 0 Then
        ExtractImagePdfText = mLastErrorMsg
    Else
        ExtractImagePdfText = "#ERR:E0303:画像からの文字抽出に失敗しました"
    End If
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー(すべてPrivate: optVisionの公開契約はPing/ExtractImagePdf/
' ExtractImagePdfTextのみ)
' ----------------------------------------------------------------------------

Private Function IsImageExt(ByVal ext As String) As Boolean
    IsImageExt = (InStr(1, "," & IMAGE_EXTS & ",", "," & ext & ",", vbTextCompare) > 0)
End Function

' 失敗判定(確定仕様準拠: ラッパーはエラーも文字列で返す。台帳§0)。
'   "#ERR:" 始まり = TryRibbonRun側の失敗 / 空・"error" = リボン側の失敗 /
'   上限系メッセージ = E0204の可能性(LooksLikeLimitError)。
Private Function IsVisionError(ByVal s As String) As Boolean
    If LenB(Trim$(s)) = 0 Then
        IsVisionError = True
    ElseIf Left$(s, 5) = "#ERR:" Then
        IsVisionError = True
    ElseIf StrComp(Trim$(s), "error", vbTextCompare) = 0 Then
        IsVisionError = True
    Else
        IsVisionError = modGateway.LooksLikeLimitError(s)
    End If
End Function

Private Function SafeResultToString(ByVal result As Variant) As String
    On Error GoTo Empty0
    If IsError(result) Then Exit Function
    SafeResultToString = CStr(result)
    Exit Function
Empty0:
    SafeResultToString = ""
End Function

Private Function JoinPagesText(pages() As ExtractedPage) As String
    On Error GoTo Empty0
    Dim n As Long
    n = UBound(pages) - LBound(pages) + 1
    If n <= 0 Then Exit Function

    Dim parts() As String
    ReDim parts(0 To n - 1)
    Dim lo As Long
    lo = LBound(pages)
    Dim i As Long
    For i = 0 To n - 1
        parts(i) = pages(lo + i).Text
    Next i
    JoinPagesText = Join(parts, vbLf)
    Exit Function
Empty0:
    JoinPagesText = ""
End Function
