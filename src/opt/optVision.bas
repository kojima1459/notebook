Attribute VB_Name = "optVision"
Option Explicit

' ============================================================================
' optVision - Vision API(文字起こし)+画像PDF OCRのオーケストレーション
'             (opt機能・MASTER_SPEC §7.7)
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
'   ・ExtractImagePdfの失敗理由をExtractImagePdfTextへ伝えるため、モジュール内
'     Private変数(mLastErrorMsg)に直近の失敗メッセージを保持する(VBAは単一
'     スレッドなので競合しない)。
'   ・modUIMain.SetStage 以外のコア参照は行わない(基盤層のみ参照可・§7.7)。
'   ・画像PDF(E0303)のOCR取込(2026-07-31 R6): 会社公式ツール同梱の
'     Ghostscript でページ毎にJPEG化してから、上と同じ経路(Base64FromFile→
'     ChatGPTV)で1ページずつ読む。「PDF→画像変換の手段がVBA単体に無い」という
'     唯一の穴を、公式配布物の再利用で塞ぐ。入口は ExtractPdfOcrPagedText。
'     コマンド文字列の組み立ては optOcrCore(純ロジック)に切り出してある。
'     GSを動かす道具(一時フォルダ・非同期起動・完了待ち・後始末)は
'     optGsTxt へ移設済み(R10-3b)、「何ページ描いて何をOCRするか」の段取りは
'     optOcrPage へ移設済み(R14-4a)で、ここは入口と前後処理だけを持つ。
'   ・Ghostscriptの配布・自動検出(2026-07-31 R9): dist/Ghostscript を常置
'     配布物に同梱する方針(要件書R9)に伴い、ResolveGsExeの解決順を
'     「config ghostscript_path → 同梱(ThisWorkbook.Path\Ghostscript\
'     gswin32c.exe)→ config ghostscript_search_dirs(IT焼き込み用)→
'     案内カード(1回きり・赤エラーでなく)」の4段へ拡張した。
'     候補パスの組み立ては optOcrCore.GsCandidatePaths / GsCandidatesForFolder
'     (純ロジック)に切り出し、実在確認(Dir$)とApplication.FileDialogは副作用
'     なのでここに残す(§7.7契約: コア参照は modUIMain.SetStage のみ)。
'     解決結果はセッション内にキャッシュし、案内カードはこのブックを開いて
'     いる間は1回しか出さない(毎回ダイアログで割り込まない)。
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
' 2026-08-07(R21-3 E1・実機第8報②): 全見出しに一律「# 」を付けさせていたため
' modChunker.ClassifyLineが目次の項目行・小見出しまで無条件に章境界(lbl=1)
' 扱いし、60章上限に張り付いて俯瞰の出典が目次ページへ偏っていた。階層指示
' (章のみ単一#・節以下は##・目次行には付けない)へ変更する(対はClassifyLine
' 側の#個数判定とLooksLikeTocPage)。再取込した資料にのみ効く(breadcrumbは
' 取込時焼き込みのため)。
Private Const VISION_PROMPT As String = _
    "この画像に写っている文字情報を、一字一句省略せずにすべて書き起こしてください。" & vbLf & _
    "・見出しは階層を区別すること: 章・部・編の最上位タイトルだけ「# 」、節・条・小見出しは「## 」を付けること" & vbLf & _
    "・目次ページの項目行には「# 」「## 」いずれも付けないこと" & vbLf & _
    "・表は1行=1レコードの形で、列名と値の対応が分かるように書き出すこと" & vbLf & _
    "・数値・金額・日付・単位・記号は画像のとおり正確に転記すること(要約や丸めは禁止)" & vbLf & _
    "・図やグラフは、読み取れる軸・凡例・数値を含めて内容を文章で説明すること" & vbLf & _
    "・画像に無い情報を補って書かないこと"

Private mLastErrorMsg As String

' R9: Ghostscript解決のセッションキャッシュと案内カードの1回きりフラグ。
Private mGsExeCache As String        ' 解決済みパス(mGsResolved=Trueのときのみ有効)
Private mGsResolved As Boolean       ' Trueなら以後mGsExeCacheをそのまま使う
Private mGsGuidanceShown As Boolean  ' 案内カード(FileDialog)はセッション中1回だけ

' R10-2: GS解決の観測性(実機初報A)。E0303のerr_log detailへ載せるための
' 直近状態を保持する(いずれも副作用ゼロの記録用途のみ)。
Private mGsLastCands As String  ' FindGsExeByCandidatesが最後に試した候補(|区切り)
Private mGsCardState As String  ' 案内カードの結果: ""/"cancel"/"nofind"/"saved"

' R14-F2: 直近のOCRが「途中で中断した」のか「上限まで読み切った」のか。
' コア層へ渡せるのは InvokeFeature の文字列1本だけで理由を運ぶ余地が無い。
' 取込1件のあいだ ExtractPdfOcrPagedText → OcrCapMemo が必ずこの順で呼ばれる
' ので、その間だけ覚えておく(入口で必ずFalseへ戻す=持ち越さない)。
Private mLastOcrAborted As Boolean

' R15-4a/4b/6b: 直近のOCRが用意した「正直なメモ」。理由(利用者操作/AI上限/
' 変換エラー/頁欠け)を知っているのは optOcrPage だけなので、文面はあちらで
' 組み立ててもらい、ここは OcrCapMemo が返すまで預かるだけ。
Private mLastOcrMemo As String

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
    If modConfig.GetBool("mock_llm", False) Then Exit Function
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
    If modConfig.GetBool("mock_llm", False) Then
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

    modUIMain.SetStage "" & modEmj.Picture() & " 画像から文字を読み取っています…"

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

' ============================================================================
' ExtractPdfOcrPagedText - 画像PDFをGhostscriptでページ毎にJPEG化し、
'   1ページ=1回のChatGPTV(既存のVISION_PROMPT・resolution="high")で
'   文字起こしして、modUtil.JoinPagedText の1本の文字列にして返す(R6)。
' ----------------------------------------------------------------------------
'   成功: ページ付きテキスト(先頭行が@@NEXUS_TRUNCATED@@なら上限で打ち切り)
'   失敗: "#ERR:E0303:<何が起きたか+どうすればよいか>"
'
'   ・GSの起動はブロッキング待ちにしない(壊れたPDF1つでExcelが永久に固まる
'     のを避ける)。完了フラグの出現をDoEventsつきで監視し、config
'     gs_abs_timeout_sec(既定1200秒)を超えたら失敗として返す(R13-F4)。
'   ・2026-08-03(R14-4a): 描画とOCRの本体は optOcrPage へ移した。GSは20ページ
'     ずつ複数回起動し、そのバッチのOCRが終わったら即座に画像を消す。ここは
'     GS解決・一時フォルダ・後始末という「前後」だけを持つ。
'   ・一時フォルダは成功・失敗どちらの経路でも必ず片付ける。後始末はR6規約に
'     従って別Sub(optGsTxt.CleanupOcrFolder)へ切り出してある(稼働中のエラー
'     ハンドラの中では On Error Resume Next が効かないため)。ただしタイム
'     アウトで1枚も描けなかったときだけは、書きかけの gs_out.log を残すために
'     消さない(optOcrPage が keepWork=True で知らせる。R11-D H-1)。
'   ・将来課題: ChatGPTVは最大10枚のバッチ入力に対応しているが、まとめて渡すと
'     ページ境界が崩れて出典ページ番号が信用できなくなるためv1では使わない。
'
'   silent(2026-07-31 R11-A C4): True=無人経路(フォルダ同期・起動時同期)。
'     Ghostscript が見つからないときの案内カード(MsgBox+フォルダ選択)を出さず
'     静かな解決だけを行う。誰も見ていない画面でモーダルが開くと同期はそこで
'     止まり、翌朝まで誰も気付けない。呼び出しは InvokeFeature 経由なので
'     省略可能な引数にしてある(省略時=従来どおり案内を出す)。
' ============================================================================
'   origPath(R15-7d): 【元のファイル】のフルパス。path は一時コピー
'     (mbtmp_*)で名前も日時も毎回変わり、頁キャッシュの鍵にできない。
Public Function ExtractPdfOcrPagedText(ByVal path As String, _
                                       Optional ByVal silent As Boolean = False, _
                                       Optional ByVal origPath As String = "") As String
    Dim folderPath As String: folderPath = ""
    mLastOcrAborted = False        ' R14-F2: 前の資料の結果を持ち越さない
    mLastOcrMemo = ""              ' R15-4a: メモも同じく持ち越さない

    On Error GoTo Fail

    If modUtil.ExtOf(path) <> "pdf" Then
        ExtractPdfOcrPagedText = "#ERR:E0303:この処理はPDF専用です(画像ファイルは" & _
            "そのまま取り込めます)。"
        Exit Function
    End If

    If modConfig.GetBool("mock_llm", False) Then
        ExtractPdfOcrPagedText = "#ERR:E0303:mockモード(mock_llm=TRUE)では画像PDFの" & _
            "OCR取込を利用できません。config の mock_llm を FALSE にすると実際に" & _
            "読み取れるようになります。"
        Exit Function
    End If

    Dim gsExe As String: gsExe = ResolveGsExe(silent)
    If LenB(gsExe) = 0 Then
        ' R10-2: 「GS未検出」は実機初報Aの本体だが従来は完全無言だった。
        ' 候補パス・案内カードの状態を必ずerr_logへ残す(観測性ゼロの解消)。
        modLog.LogError "E0303", "optVision.ExtractPdfOcrPagedText", modUtil.SafeLeft( _
            "GS未解決 wbPath=" & ThisWorkbook.path & " cand=" & mGsLastCands & _
            " card=" & mGsCardState, 2000)
        ExtractPdfOcrPagedText = "#ERR:E0303:画像PDFを読み取るための Ghostscript が" & _
            "見つかりませんでした。配布zip(リポジトリのZIPをDL→解凍したもの)は " & _
            "Ghostscript フォルダを同梱済みのはずです。このファイルと同じ場所に " & _
            "Ghostscript フォルダがあるかご確認いただくか、config シートの " & _
            "ghostscript_path に gswin32c.exe のフルパスを設定してください。" & vbLf & _
            "このファイルの場所: " & ThisWorkbook.path & vbLf & _
            "zipの中から直接開くと動きません。zipを右クリック→[すべて展開]で展開し、" & _
            "展開先の MyBookshelf.xlsm を開いてください。"
        Exit Function
    End If

    folderPath = optGsTxt.MakeOcrFolder()
    If LenB(folderPath) = 0 Then
        ExtractPdfOcrPagedText = "#ERR:E0303:作業用の一時フォルダを作成できませんでした。" & _
            "一時フォルダ(TEMP)の空き容量と書き込み権限をご確認ください。"
        Exit Function
    End If

    Dim maxPages As Long: maxPages = OcrMaxPages()
    Dim dpi As Long: dpi = modConfig.GetLong("vision_pdf_dpi", 150)
    ' R14-F6: gs_abs_timeout_sec は【1資料あたり】の画像化待ちの絶対上限。
    ' バッチごとの残り時間の配分は optOcrPage が行う。
    Dim absSec As Long: absSec = modConfig.GetLong("gs_abs_timeout_sec", 1200)
    If absSec < 10 Then absSec = 10

    ' R14-4a: 「何ページ描いて、どれをOCRして、いつ画像を捨てるか」は
    ' optOcrPage が持つ(20ページずつのバッチ制御・ページ間DoEvents・
    ' バッチ完了ごとのJPEG削除)。ここは前後(GS解決・一時フォルダ・後始末)
    ' だけを持つ。keepWork=True で返ったらフォルダを消さない(R11-D H-1 /
    ' R14-F5: 止められなかったGSが書いている最中のフォルダも消さない)。
    ' R15-5a: 総頁数は、この資料が【OCRへ回る前に】必ず通っている txtwrite の
    ' 分類パス(optGsTxt)がGSに出させたログから既に読めていることが多い。
    ' 従来 knownTotal は最終バッチに入るまで0のままで、24頁中20頁が終わるまで
    ' 分母もETAも出なかった(実機第4報 RC2)。ここで受け取って先へ渡す。
    ' opt層内の参照なのでR2に触れない(コア層を1つも経由しない=
    ' modFeatures.InvokeFeature の往復を1つも増やさない。RC8の火種を増やさない)。
    ' 取れなければ0で、従来どおり最終バッチでの確定へ自然に縮退する。
    Dim gsTotal As Long: gsTotal = 0
    On Error Resume Next
    gsTotal = optGsTxt.LastGsTotalPages()
    On Error GoTo Fail

    ' R15-7c(RC10): 総頁が分かっているなら画像化待ちの絶対上限も頁数へ連動
    ' させる(既定1200秒では254頁の後半が必ず時間切れになる)。configを
    ' 大きくしてある端末はそちらが勝つ。他経路の予算には触らない。
    absSec = optOcrEta.GsBudgetSec(absSec, gsTotal)

    ' R15-7d: 前回の続きから再開できる頁を読み込む(鍵は【元のパス】)。
    Dim keyPath As String: keyPath = origPath
    If LenB(keyPath) = 0 Then keyPath = path
    optOcrCache.BeginDoc keyPath

    Dim keepWork As Boolean: keepWork = False
    Dim aborted As Boolean: aborted = False
    Dim ocrMemo As String: ocrMemo = ""
    Dim ocrText As String
    ocrText = optOcrPage.OcrPdfByBatch(gsExe, path, folderPath, dpi, maxPages, _
                                       absSec, VISION_PROMPT, gsTotal, _
                                       keepWork, aborted, ocrMemo)
    mLastOcrAborted = aborted
    mLastOcrMemo = ocrMemo

    modUIMain.SetStage ""
    If Not keepWork Then optGsTxt.CleanupOcrFolder folderPath

    ' 1枚も描けなかった/全ページ読めなかったときだけ "#ERR:..." が返る。
    ExtractPdfOcrPagedText = ocrText
    Exit Function

Fail:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    On Error GoTo 0
    modUIMain.SetStage ""
    optGsTxt.CleanupOcrFolder folderPath
    modLog.LogError "E0303", "optVision.ExtractPdfOcrPagedText", errDesc
    ExtractPdfOcrPagedText = "#ERR:E0303:画像PDFの読み取り処理でエラーが発生しました: " & errDesc
End Function

' ----------------------------------------------------------------------------
' ResetGsGuidance - GS未検出の案内カード(セッション1回きり)を再提示可能に
'   戻す(R10-2)。mGsGuidanceShownのみFalseへ戻し、解決済みキャッシュ
'   (mGsResolved/mGsExeCache)には触れない(設置済みのGSは再解決不要)。
'   利用者が能動的に「資料を追加」を実行する入口(modShelfBatch.AddFilesResult)
'   から modFeatures.InvokeFeature("vision","ResetGsGuidance",…) 経由で
'   呼ばれる想定。自動同期(modShelfSync)からは呼ばれない=起動時に
'   モーダルが出ない現行性質を維持する。戻り値は契約合わせの空文字列。
' ----------------------------------------------------------------------------
Public Function ResetGsGuidance() As String
    mGsGuidanceShown = False
    ' R10c(L6): カードの結果も一緒に戻す。残したままだと次にE0303が出たとき、
    ' err_logのdetailに前回の card=cancel 等が載って調査を誤らせる。
    mGsCardState = ""
    ResetGsGuidance = ""
End Function

' R10-3: テキストPDFのCOM無し抽出(実体はoptGsTxt・容量都合の分割)。
' vision機能の行き先はoptVision固定なので受け口だけ置く。
Public Function ExtractPdfTextNoOcr(ByVal path As String) As String
    ExtractPdfTextNoOcr = optGsTxt.ExtractPdfTextNoOcr(path)
End Function

' ----------------------------------------------------------------------------
' OcrCapMemo - OCRが打ち切られたときに本棚カードへ出す正直なメモ(R14-4c)。
'   文言と算数は optOcrCore(純ロジック)が持ち、ここは
'   modFeatures.InvokeFeature("vision","OcrCapMemo",…) から呼べるようにする
'   ための受け口。コア層は opt モジュール名を書けない(R2)ので、この1本が
'   「打ち切りの説明」をコアへ渡す唯一の経路になる。truncated=False なら空。
'   R14-F2: 打ち切りには2種類あり言うべきことが違う(上限まで読んだ=設定を
'   変えれば続きも読める / 途中で中断した=再試行される)。どちらかを知って
'   いるのは直前の ExtractPdfOcrPagedText だけ(mLastOcrAborted)。
'   R14-F7: capN は【設定の上限】。呼び出し元は 0(不明)を渡してよく、
'   そのときはここが config から解決する(既定値を2箇所に持たない)。
' ----------------------------------------------------------------------------
'   R15-4a/4b/6b: 打ち切りの種類はさらに増えた(利用者の中断・AI利用の上限・
'   一部の頁だけ読めなかった)。理由を知っているのは optOcrPage だけなので、
'   文面はあちらが optOcrEta の純関数で作って渡してくる(mLastOcrMemo)。
'   R15-FixA(FA-4): mLastOcrMemo があるときは【そのまま】返す。以前はここで
'   truncated を条件に握りつぶし、上限メモと二者択一で【置換】していたため、
'   (a)再開だけで読み切った取込の「前回の続きから再開しました。」がカードへ
'   一度も届かず、(b)頁欠けを言うと上限の説明が消える、の2つの穴があった。
'   3事実の連結は optOcrEta.ComposeOcrMemo が1箇所で済ませてある。下の2経路は
'   OcrPdfByBatch を通らなかったとき(旧経路・例外)の保険。
'   R15-4c: capN には必ず optOcrCore.SafeMaxPages を掛ける(config=300でも
'   ハード上限で頭打ちになるのに生値で計算し、事実と違う数字を出していた)。
Public Function OcrCapMemo(ByVal truncated As Boolean, ByVal keptN As Long, _
                           ByVal capN As Long) As String
    If LenB(mLastOcrMemo) > 0 Then
        OcrCapMemo = mLastOcrMemo
        Exit Function
    End If
    If Not truncated Then Exit Function

    If mLastOcrAborted Then
        OcrCapMemo = optOcrEta.OcrAbortMemoFor(keptN, "error", False)
        Exit Function
    End If

    Dim cap As Long: cap = capN
    If cap <= 0 Then cap = OcrMaxPages()
    OcrCapMemo = optOcrEta.OcrCapMemoFor(truncated, keptN, optOcrCore.SafeMaxPages(cap))
End Function

' 取込前の事前確認の受け口(R15-7b)。コア層は opt モジュール名を書けない
' (R2)ので InvokeFeature("vision", …) から呼べる窓口をここに置く。
'   OcrConfirmAsk : "" なら確認しない/文字列ならそれを見せて Yes/No を聞く
'                   (pdfPath は【元のファイル】=キャッシュの鍵と同じもの)。
'   OcrDeclineMemo: 直前の見積もりで作る「見送りました」のメモ。
Public Function OcrConfirmAsk(ByVal pdfPath As String) As String
    OcrConfirmAsk = optOcrCache.ConfirmAskFor(pdfPath, OcrMaxPages())
End Function

Public Function OcrDeclineMemo() As String
    OcrDeclineMemo = optOcrCache.DeclineMemo()
End Function

' 頁キャッシュの孤児行(2日超)の起動時GC(R15-7d)。modBoot の既存GC群
' (nxocr_*/mbtmp_*)と同じ線で1回だけ呼ばれる。vision無効なら呼ばれないが、
' そのときはキャッシュ自体が1行も作られないので何も溜まらない。
Public Function OcrCacheGc() As String
    OcrCacheGc = optOcrCache.GcOldRows()
End Function

' 本棚に done として並んだ資料の頁控えを消す受け口(R15-FixB FB-1)。
' 呼ぶのは modShelf(status確定の直後)で、OCRを1度も通っていない資料から
' 呼ばれても該当行が無いので何も起きない。
Public Function OcrCachePurge(ByVal pdfPath As String) As String
    OcrCachePurge = optOcrCache.PurgeFor(pdfPath)
End Function

' OCRの1資料あたりページ上限(config vision_pdf_max_pages)。既定値をここ
' 1箇所だけに持つ(R14-F7: メモに出す上限と実際に使う上限を絶対に割らない)。
' R15-7a: 既定 100 → 300(254頁のスキャンPDFを分割せずに取り込めるように
' する。ハード上限 optOcrCore.PAGES_MAX も 200 → 300)。
Private Function OcrMaxPages() As Long
    OcrMaxPages = modConfig.GetLong("vision_pdf_max_pages", 300)
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー(GS実行まわりのみoptGsTxt/optOcrPage用にPublic・R10-3/R14-4a)
' ----------------------------------------------------------------------------

' ----------------------------------------------------------------------------
' ResolveGsExe - Ghostscript実行ファイルを解決する(R9: 4段階)。
'   (1) config ghostscript_path(明示フルパス)
'   (2) 同梱(ThisWorkbook.Path\Ghostscript\gswin32c.exe)
'   (3) config ghostscript_search_dirs(セミコロン区切り。IT焼き込み用)
'   (4) 見つからない場合: 1回きりの案内カード→フォルダ選択→即続行
'   候補の組み立ては optOcrCore.GsCandidatePaths(純ロジック)、実在確認は
'   ここ(Dir$)。解決済みパスはセッション内(モジュール変数)にキャッシュする。
'   silent=True(無人経路。R11-A C4)のときは (4) を丸ごと飛ばし、(1)-(3) の
'   静かな解決だけを行う。「1回きり」の権利(mGsGuidanceShown)も消費しない
'   ため、あとで人が手動で取り込んだときに案内カードは従来どおり出る。
' ----------------------------------------------------------------------------
Private Function ResolveGsExe(ByVal silent As Boolean) As String
    If mGsResolved Then
        ResolveGsExe = mGsExeCache
        Exit Function
    End If

    Dim found As String
    found = FindGsExeByCandidates()
    If LenB(found) > 0 Then
        CacheGsExe found
        ResolveGsExe = found
        Exit Function
    End If

    ' (4) 案内カード。見つからなかった/カード自体を出したことが既にある
    '     場合は、今回はダイアログを出さず未解決のまま返す(1回きり)。
    '     mGsResolvedはTrueにしない: config書き換え・フォルダ設置が後から
    '     行われた場合に、次回呼び出しで(3)までを毎回再チェックできるように
    '     しておくため(ダイアログさえ出さなければ再探索自体は軽い)。
    If silent Then
        ' 無人経路。ダイアログは出さず、何が起きたかだけ残す(呼び出し元は
        ' このあと E0303 を err_log へ書き、資料は image_pdf として残る)。
        mGsCardState = "skip_silent"
        On Error Resume Next
        modLog.LogUsage "gs_guidance_card", mGsCardState, _
            "無人の取込のため Ghostscript の案内は出しませんでした"
        On Error GoTo 0
        Exit Function
    End If

    If Not mGsGuidanceShown Then
        mGsGuidanceShown = True
        Dim viaCard As String
        viaCard = OfferGsFolderPicker()
        If LenB(viaCard) > 0 Then
            CacheGsExe viaCard
            ResolveGsExe = viaCard
            Exit Function
        End If
    End If

    ResolveGsExe = ""
End Function

' config ghostscript_path / 同梱 / ghostscript_search_dirs の順で候補を
' 実在確認する(optOcrCore.GsCandidatePathsが組み立てた候補文字列を使う)。
' R10-3でPublic化: 案内カードを出せないoptGsTxt用の「静かな解決」がこの関数。
Public Function FindGsExeByCandidates() As String
    Dim cfgPath As String
    cfgPath = modConfig.GetString("ghostscript_path", "")
    Dim searchDirs As String
    searchDirs = modConfig.GetString("ghostscript_search_dirs", "")

    Dim candStr As String
    candStr = optOcrCore.GsCandidatePaths(cfgPath, ThisWorkbook.path, searchDirs)
    mGsLastCands = candStr   ' R10-2: E0303のerr_log detailで参照するため毎回更新
    FindGsExeByCandidates = FirstExistingCandidate(candStr)
End Function

' "|" 区切りの候補文字列から、実在する最初の1件を返す(無ければ空文字列)。
Private Function FirstExistingCandidate(ByVal candStr As String) As String
    If LenB(candStr) = 0 Then Exit Function
    Dim cands() As String: cands = Split(candStr, "|")
    Dim i As Long
    For i = LBound(cands) To UBound(cands)
        If PathExists(cands(i)) Then
            FirstExistingCandidate = cands(i)
            Exit Function
        End If
    Next i
End Function

Private Sub CacheGsExe(ByVal resolvedPath As String)
    mGsExeCache = resolvedPath
    mGsResolved = True
End Sub

' ----------------------------------------------------------------------------
' OfferGsFolderPicker - 1回きりの丁寧な案内(赤エラーでなく)+フォルダ選択。
'   選択フォルダ直下または bin\ 直下に gswin32c.exe があれば
'   config ghostscript_path へ保存して即そのパスを返す。無ければ
'   modUIMain.SetStage で1行案内して ""(未解決)を返す(§7.7契約により
'   opt層からのUI参照はSetStageのみ許可)。
'   ※LibreOfficeテスト環境にはApplication.FileDialogが無いため、この関数と
'   PickGsFolder(実際のダイアログ呼び出し)はLOの純ロジックテスト(モード1)
'   からは一切呼ばれない(候補の組み立てはoptOcrCore側の純ロジックとして
'   分離済み・そちらだけが境界値テストの対象)。
' ----------------------------------------------------------------------------
Private Function OfferGsFolderPicker() As String
    Dim resp As VbMsgBoxResult
    resp = MsgBox( _
        "画像PDFの読み取りには Ghostscript が必要です。" & vbLf & _
        "配布zipの Ghostscript フォルダをこのファイルの隣に置くか、" & vbLf & _
        "場所を指定してください。" & vbLf & vbLf & _
        "[OK] を押すとフォルダを選ぶ画面が開きます。", _
        vbInformation + vbOKCancel, modAppDef.APP_NAME)
    If resp <> vbOK Then
        mGsCardState = "cancel"
        modLog.LogUsage "gs_guidance_card", mGsCardState, ""
        Exit Function
    End If

    Dim picked As String
    picked = PickGsFolder()
    If LenB(picked) = 0 Then
        mGsCardState = "cancel"
        modLog.LogUsage "gs_guidance_card", mGsCardState, ""
        Exit Function
    End If

    Dim candStr As String: candStr = optOcrCore.GsCandidatesForFolder(picked)
    Dim found As String: found = FirstExistingCandidate(candStr)
    If LenB(found) = 0 Then
        modUIMain.SetStage "選んだフォルダに gswin32c.exe が見つかりませんでした。"
        mGsCardState = "nofind"
        modLog.LogUsage "gs_guidance_card", mGsCardState, modUtil.SafeLeft(picked, 200)
        Exit Function
    End If

    modConfig.SetValue "ghostscript_path", found
    mGsCardState = "saved"
    modLog.LogUsage "gs_guidance_card", mGsCardState, modUtil.SafeLeft(found, 200)
    OfferGsFolderPicker = found
End Function

' フォルダ選択ダイアログ(msoFileDialogFolderPicker=4)。名前付き定数は使わない
' (modShelfSync.PickShelfFolderと同じ慣習・LibreOffice互換のため)。
' キャンセル・実行時エラー(FileDialog非対応環境)はどちらも空文字列で返す。
Private Function PickGsFolder() As String
    On Error GoTo NoDialog
    Dim fd As Object
    Set fd = Application.FileDialog(4)   ' msoFileDialogFolderPicker
    fd.Title = "Ghostscript フォルダを選択してください"
    If fd.Show <> -1 Then Exit Function   ' キャンセル
    If fd.SelectedItems.count < 1 Then Exit Function
    PickGsFolder = CStr(fd.SelectedItems(1))
    Exit Function
NoDialog:
    PickGsFolder = ""
End Function

' R10-3bで optGsTxt へ移設したGS実行の道具(MakeOcrFolder/
' WaitForDoneFlag/CleanupOcrFolder)からも使うためPublic。
Public Function PathExists(ByVal p As String) As Boolean
    On Error Resume Next
    PathExists = (LenB(Dir$(p)) > 0)
    On Error GoTo 0
End Function

Private Function IsImageExt(ByVal ext As String) As Boolean
    IsImageExt = (InStr(1, "," & IMAGE_EXTS & ",", "," & ext & ",", vbTextCompare) > 0)
End Function

' 失敗判定(確定仕様準拠: ラッパーはエラーも文字列で返す。台帳§0)。
'   "#ERR:" 始まり = TryRibbonRun側の失敗 / 空・"error" = リボン側の失敗 /
'   上限系メッセージ = E0204の可能性(LooksLikeLimitError)。
' R14-4a: ページOCRのループを optOcrPage へ移したので、同じ判定を2箇所に
' 持たないようPublicにしてある(憲章§4-5)。opt層内の参照なのでR2に触れない。
Public Function IsVisionError(ByVal s As String) As Boolean
    If LenB(Trim$(s)) = 0 Then
        IsVisionError = True
    ElseIf Left$(s, 5) = "#ERR:" Then
        IsVisionError = True
    ElseIf StrComp(Trim$(s), "error", vbTextCompare) = 0 Then
        IsVisionError = True
    Else
        IsVisionError = modRibbonFail.LooksLikeLimitError(s)
    End If
End Function

' R14-4a: optOcrPage(ページOCRのループ)からも使うのでPublic。
Public Function SafeResultToString(ByVal result As Variant) As String
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
