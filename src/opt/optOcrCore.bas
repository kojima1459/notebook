Attribute VB_Name = "optOcrCore"
Option Explicit

' ============================================================================
' optOcrCore - 画像PDFのOCR取込(R6)で使う「文字列の組み立てだけ」を集めた
'              純ロジックモジュール(R4準拠)
' ----------------------------------------------------------------------------
' 役割:
'   Ghostscript(gswin32c.exe)へ渡すコマンドライン、一時フォルダ・出力ファイル
'   名の規約、ページ上限と切り詰め判定の算数を持つ。実際にプロセスを起動したり
'   ファイルを読み書きするのは optVision 側の仕事で、ここは一切の副作用を
'   持たない(Worksheets/Range/Application/ThisWorkbook/MsgBoxに触れない)。
'
' 設計判断:
'   ・引用符の付け方こそがこの機能の一番の地雷なので、そこだけを切り出して
'     ゴールデンテスト(modTestsPure4)で1文字単位に固定する。会社公式ツールの
'     帳票OCR版(chouhyou)は gsPath を引用符で囲んでおらず、ユーザー名に空白が
'     入る端末では起動すらしなかった(docs/dev/reference_公式OCRツール調査_
'     20260731.md §5)。同じ轍は踏まない: **すべてのパスを二重引用符で囲む**。
'     Windowsのファイル名に " は使えないため、引用符のエスケープは考えなくてよい。
'   ・GS実行はブロッキング待ちにしない(壊れたPDF1つでExcelが永久に固まる)。
'     そのため「GS本体のコマンド」と「完了フラグ作成まで含めたcmd.exeラップ」を
'     BuildGsCommand / BuildRunCommand の2段に分けている。フラグの作成は `&`
'     (無条件)で連結するので、GSが異常終了しても監視ループは必ず解ける。
'   ・ページ上限は「上限+1ページまで描画してみて、上限+1枚できていたら
'     切り詰めあり」と判定する(総ページ数を知る手段がGS抜きには無いため)。
'     RenderCapFor / IsTruncatedCount / KeepPageCount がその算数。
'   ・optVision からのみ呼ばれる。opt層内の参照なので依存ルール上も問題ない。
'   ・Ghostscript実行ファイルの解決候補列挙(2026-07-31 R9): 同梱配布(dist/
'     Ghostscript同封)・IT焼き込み(ghostscript_search_dirs)の2経路を
'     追加するにあたり、「候補パス文字列をどの順で・どう組み立てるか」を
'     ここへ切り出した(GsCandidatePaths / GsCandidatesForFolder)。
'     実際のファイル存在確認(Dir$)とApplication.FileDialogは副作用そのもの
'     なのでoptVision側の仕事のまま。戻り値は "|" 区切りの1本の文字列にした
'     (run_lo_tests.pyの既知の制約により `As String()` という配列戻り値の
'     関数宣言はLibreOffice Basicでコンパイルがハングするため。modUtil.
'     SplitKeepNonEmptyと同じ理由・同じ回避策)。
'   ・テキストPDFのCOM無し抽出(2026-07-31 R10-3): Word/AcrobatのCOMが端末
'     ポリシーで塞がれていてもテキストPDFを取り込めるよう、同じGhostscriptの
'     txtwriteデバイスを使う経路を足した。その「コマンドの組み立て」と
'     「抜き出せた文字数から採否を決める判定」だけをここへ置く
'     (BuildGsTextCommand / GsTextVerdict)。実行するのは optGsTxt。
' ============================================================================

' 公式ツール(gazou版)と同じ規約。ページ番号3桁ゼロ埋め。
Private Const JPG_PREFIX As String = "page_"
Private Const JPG_SUFFIX As String = ".jpg"
Private Const JPG_PATTERN As String = "page_%03d.jpg"
Private Const DONE_FLAG_NAME As String = "done.flag"
Private Const TEMP_DIR_PREFIX As String = "nxocr_"

' dpiの安全範囲(既定150は公式帳票OCR版と同値。300は倍重い)。
Private Const DPI_MIN As Long = 72
Private Const DPI_MAX As Long = 600
Private Const DPI_DEFAULT As Long = 150

' 1ファイルあたりのページ上限の安全範囲(configの値が壊れていても暴走しない)。
Private Const PAGES_MIN As Long = 1
Private Const PAGES_MAX As Long = 200

' txtwrite救済(R10-3)の採否しきい値。空白類を除いた抽出文字数がこれ未満なら
' 「文字が薄すぎる」とみなす。テキストPDFなら1ページでも普通は数百字ある。
' R10c(M1/M2): 固定40字だけでは「100ページのスキャンPDFに透明テキストが
' 50字だけ乗っている」ケースを採用してしまい、OCRへ回らず中身の無い資料が
' 本棚に入る。ページ数に比例する下限(1ページ10字。modExtractorの既存の
' 画像PDF判定と同じ係数)と、固定40字の大きい方をしきい値にする。
Private Const GS_TEXT_MIN_CHARS As Long = 40
Private Const GS_TEXT_MIN_PER_PAGE As Long = 10

Public Function Ping() As Boolean
    Ping = True
End Function

' ----------------------------------------------------------------------------
' BuildGsCommand - Ghostscript本体のコマンドラインを組み立てる。
'   gsExe      : gswin32c.exe のフルパス
'   pdfPath    : 変換元のPDFのフルパス
'   outPattern : 出力パターン(例 C:\Temp\nxocr_x\page_%03d.jpg)
'   dpi        : 解像度(範囲外は既定150へ丸める)
'   lastPage   : 描画する最終ページ(1未満は1へ丸める)
'   戻り値の形(パスは全て二重引用符で囲む):
'     "<gs>" -dSAFER -dNOPAUSE -dBATCH -sDEVICE=jpeg -r150
'     -dTextAlphaBits=4 -dGraphicsAlphaBits=4 -dFirstPage=1 -dLastPage=21
'     -sOutputFile="<out>" "<pdf>"
' ----------------------------------------------------------------------------
Public Function BuildGsCommand(ByVal gsExe As String, ByVal pdfPath As String, _
                               ByVal outPattern As String, ByVal dpi As Long, _
                               ByVal lastPage As Long) As String
    Dim r As Long: r = SafeDpi(dpi)
    Dim lastN As Long: lastN = lastPage
    If lastN < 1 Then lastN = 1

    BuildGsCommand = Quoted(gsExe) & _
        " -dSAFER -dNOPAUSE -dBATCH" & _
        " -sDEVICE=jpeg -r" & CStr(r) & _
        " -dTextAlphaBits=4 -dGraphicsAlphaBits=4" & _
        " -dFirstPage=1 -dLastPage=" & CStr(lastN) & _
        " -sOutputFile=" & Quoted(outPattern) & _
        " " & Quoted(pdfPath)
End Function

' ----------------------------------------------------------------------------
' BuildRunCommand - GSコマンドを cmd.exe でラップし、終了後に完了フラグ
'   ファイルを作らせる(非同期起動の完了検知に使う)。
'   `/s` を付けると cmd.exe は「最初と最後の二重引用符だけを外し、残りを
'   そのまま実行する」ので、内側のパスの引用符が壊れない。
'   連結は `&`(無条件)なのでGSが失敗しても必ずフラグが作られ、
'   監視ループがタイムアウトを待たずに解ける。
' ----------------------------------------------------------------------------
Public Function BuildRunCommand(ByVal gsCommand As String, ByVal doneFlagPath As String) As String
    BuildRunCommand = "cmd.exe /s /c " & Chr$(34) & gsCommand & _
        " & echo done>" & Quoted(doneFlagPath) & Chr$(34)
End Function

' ----------------------------------------------------------------------------
' BuildGsTextCommand - Ghostscriptの txtwrite デバイスで「PDFに埋まっている
'   文字をそのまま抜き出す」コマンドラインを組み立てる(2026-07-31 R10-3)。
'   gsExe   : gswin32c.exe のフルパス
'   pdfPath : 変換元のPDFのフルパス
'   outTxt  : 書き出し先テキストファイルのフルパス
'   戻り値の形(パスは全て二重引用符で囲む・BuildGsCommandと同じ規約):
'     "<gs>" -dSAFER -dNOPAUSE -dBATCH -sDEVICE=txtwrite
'     -sOutputFile="<out>" "<pdf>"
'   ページ指定はしない(全ページを1本のテキストへ書き出す。txtwriteはページの
'   区切りに改ページ文字 Chr(12) を挟むので、ページ分けは読み手の仕事)。
'   出力先の指定は -o ではなく -sOutputFile + -dNOPAUSE -dBATCH にした。
'   -o は「-sOutputFile と -dBATCH -dNOPAUSE をまとめた省略形」で意味は同じだが、
'   BuildGsCommand(jpeg側)と字面を揃えておく方が、両者を見比べたときに
'   引用符の付け忘れ等の差分が目で見つけやすい。
'   完了フラグ付きの cmd.exe ラップは BuildRunCommand をそのまま共用する。
' ----------------------------------------------------------------------------
Public Function BuildGsTextCommand(ByVal gsExe As String, ByVal pdfPath As String, _
                                   ByVal outTxt As String) As String
    BuildGsTextCommand = Quoted(gsExe) & _
        " -dSAFER -dNOPAUSE -dBATCH" & _
        " -sDEVICE=txtwrite" & _
        " -sOutputFile=" & Quoted(outTxt) & _
        " " & Quoted(pdfPath)
End Function

' ----------------------------------------------------------------------------
' GsTextVerdict - txtwriteで抜き出せた文字数から採否を決める(R10-3 / R10c)。
'   cleanLen  : 空白類(スペース・タブ・改行・改ページ)を除いた抽出文字数の
'               全ページ合計(CleanTextLen の戻り値)
'   pageCount : 抜き出せたページ数(GsPageCount の戻り値)
'   戻り値の3値:
'     "image"  文字層が1文字も無い=スキャンPDF確定。呼び出し元は即E0303
'              (OCR経路)へ回してよい。Word/Acrobatに渡しても必ず空振りする。
'     "sparse" 文字はあるがページ数に対して薄すぎる。判断が割れる領域なので
'              「まずWord/Acrobatに任せ、両方失敗したときだけE0303」に回す。
'              ・薄い透明テキスト付きスキャンPDF → Word/Acrobatも失敗 → OCRへ
'              ・表紙だけの短い正当なPDF        → Word/Acrobatが成功 → 救済
'              R10cの敵対的レビュー M1/M2: 固定40字だけで判定していたため、
'              100ページのスキャンPDFに透明テキストが50字乗っているだけの
'              資料を「本文あり」として取り込んでいた(中身の無い資料が
'              本棚に入り、検索の邪魔にしかならない)。
'     "ok"     テキストPDFとして採用してよい。
'   しきい値は Max(GS_TEXT_MIN_CHARS, pageCount × GS_TEXT_MIN_PER_PAGE)。
'   後者の係数10は modExtractor の既存の画像PDF判定(総文字数 < ページ数×10)
'   と同じで、2箇所で別々の常識を持たないようにしている。
' ----------------------------------------------------------------------------
Public Function GsTextVerdict(ByVal cleanLen As Long, ByVal pageCount As Long) As String
    If cleanLen <= 0 Then
        GsTextVerdict = "image"
        Exit Function
    End If

    Dim need As Long: need = GS_TEXT_MIN_CHARS
    Dim byPage As Long: byPage = 0
    If pageCount > 0 Then byPage = pageCount * GS_TEXT_MIN_PER_PAGE
    If byPage > need Then need = byPage

    If cleanLen >= need Then
        GsTextVerdict = "ok"
    Else
        GsTextVerdict = "sparse"
    End If
End Function

' ----------------------------------------------------------------------------
' CleanTextLen - 空白類(半角/全角スペース・タブ・改行・改ページ)を除いた
'   文字数(R10cで optGsTxt から移設。純ロジックなのでここが正しい置き場)。
'   txtwriteは画像PDFに対しても改ページ文字だけは律儀に書き出すので、素のLenで
'   判定すると「文字がある」と誤認する。1文字ずつ走査するとページ数の多いPDFで
'   無駄に重いため、Replaceで落としてから数える。
' ----------------------------------------------------------------------------
Public Function CleanTextLen(ByVal s As String) As Long
    Dim t As String

    t = Replace(s, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, vbTab, "")
    t = Replace(t, Chr$(12), "")
    t = Replace(t, " ", "")
    t = Replace(t, ChrW(&H3000), "")
    CleanTextLen = Len(t)
End Function

' ----------------------------------------------------------------------------
' GsPageCount - txtwrite出力が実質何ページ分あるかを数える(R10c)。
'   改ページ文字 Chr(12) で割り、先頭側・末尾側の空ページ(空白類だけの要素)を
'   落とした残りの件数を返す。中身が全く無ければ0。
'   modExtractor.BuildPagesFromGsText と【同じ規約】でなければならない
'   (あちらはExtractedPage配列を組むためコア層に置かざるを得ず、R2により
'   optOcrCoreを直接呼べない)。両者がズレるとGsTextVerdictの分母と実際の
'   ページ数が食い違うので、modTestsPure6 で突き合わせテストを固定してある。
' ----------------------------------------------------------------------------
'   空判定に Trim$ は使わない: VBAのTrim$は半角スペースしか落とさないため、
'   改行だけのページ(GSが白紙ページに対して吐く典型)を「中身あり」と数えて
'   しまう。CleanTextLen と同じ「空白類を全部落として残るか」で判定する。
Public Function GsPageCount(ByVal txt As String) As Long
    If CleanTextLen(txt) = 0 Then Exit Function   ' 全体が空白類だけ=0ページ

    Dim parts() As String: parts = Split(txt, Chr$(12))
    Dim firstIdx As Long: firstIdx = LBound(parts)
    Dim lastIdx As Long: lastIdx = UBound(parts)

    ' 上の早期returnにより、必ずどこかに中身のあるページがある=ループは必ず解ける。
    Do While firstIdx < lastIdx
        If CleanTextLen(parts(firstIdx)) > 0 Then Exit Do
        firstIdx = firstIdx + 1
    Loop
    Do While lastIdx > firstIdx
        If CleanTextLen(parts(lastIdx)) > 0 Then Exit Do
        lastIdx = lastIdx - 1
    Loop

    GsPageCount = lastIdx - firstIdx + 1
End Function

' 一時フォルダのフルパス(例: C:\Users\x\AppData\Local\Temp\nxocr_20260731_101112_437)。
Public Function TempFolderFor(ByVal tempRoot As String, ByVal uniqueName As String) As String
    TempFolderFor = TrimTrailingSep(tempRoot) & "\" & TEMP_DIR_PREFIX & uniqueName
End Function

' GSへ渡す出力パターン(フォルダ + page_%03d.jpg)。
Public Function OutPatternFor(ByVal folderPath As String) As String
    OutPatternFor = TrimTrailingSep(folderPath) & "\" & JPG_PATTERN
End Function

' 完了フラグファイルのフルパス。
Public Function DoneFlagFor(ByVal folderPath As String) As String
    DoneFlagFor = TrimTrailingSep(folderPath) & "\" & DONE_FLAG_NAME
End Function

' 出力される1ページ分のファイル名(GSの%03dと同じ形)。
Public Function PageJpgName(ByVal idx As Long) As String
    Dim n As Long: n = idx
    If n < 0 Then n = 0
    Dim s As String: s = CStr(n)
    If Len(s) < 3 Then s = Right$("000" & s, 3)
    PageJpgName = JPG_PREFIX & s & JPG_SUFFIX
End Function

' ----------------------------------------------------------------------------
' RenderCapFor - 実際に描画するページ数(=上限+1)。
'   総ページ数を知らずに「切り詰めが起きたか」を判定するため、上限より1枚
'   多く描かせる。上限+1枚できていれば「まだ先がある」と確定できる。
' ----------------------------------------------------------------------------
Public Function RenderCapFor(ByVal maxPages As Long) As Long
    RenderCapFor = SafeMaxPages(maxPages) + 1
End Function

' 描画できた枚数が上限を超えていれば切り詰めあり。
Public Function IsTruncatedCount(ByVal foundCount As Long, ByVal maxPages As Long) As Boolean
    IsTruncatedCount = (foundCount > SafeMaxPages(maxPages))
End Function

' 実際にOCRへ回す枚数(上限で頭打ち。余分な1枚はここで捨てる)。
Public Function KeepPageCount(ByVal foundCount As Long, ByVal maxPages As Long) As Long
    Dim cap As Long: cap = SafeMaxPages(maxPages)
    If foundCount < 0 Then Exit Function
    If foundCount > cap Then
        KeepPageCount = cap
    Else
        KeepPageCount = foundCount
    End If
End Function

' configの値が壊れていても暴走しないための丸め(公開: 診断・テスト用)。
Public Function SafeDpi(ByVal dpi As Long) As Long
    SafeDpi = dpi
    If SafeDpi < DPI_MIN Or SafeDpi > DPI_MAX Then SafeDpi = DPI_DEFAULT
End Function

Public Function SafeMaxPages(ByVal maxPages As Long) As Long
    SafeMaxPages = maxPages
    If SafeMaxPages < PAGES_MIN Then SafeMaxPages = PAGES_MIN
    If SafeMaxPages > PAGES_MAX Then SafeMaxPages = PAGES_MAX
End Function

' ----------------------------------------------------------------------------
' GsCandidatePaths - Ghostscript実行ファイル(gswin32c.exe)の解決候補を、
'   要件書R9の優先順位どおりに1本の文字列(候補ごとに "|" 区切り)で返す。
'   実在確認(Dir$)は行わない・呼び出し側(optVision)が先頭から順に試す。
'     (1) cfgPath    : config ghostscript_path(明示フルパス。空なら候補なし)
'     (2) wbDir       : wbDir\Ghostscript\gswin32c.exe(同梱・既定経路)
'     (3) searchDirs  : config ghostscript_search_dirs(セミコロン区切り)。
'                        各ディレクトリを GsCandidatesForFolder で2候補
'                        (直下 / bin直下)へ展開して順に追加する。
'   空要素(空文字列・空白のみ)はスキップし、各パスの末尾の \ と / は
'   正規化してから連結する(configの書式ゆれを吸収するため)。
'   R10-2: wbDir が "http" 始まり(大小無視。OneDriveの共有URL等で
'   ThisWorkbook.Path が "https://..." を返すケース)のときは、(2)の
'   同梱候補を組み立てない。URLへ "\Ghostscript\gswin32c.exe" を単純連結
'   しても実在確認(Dir$)が必ず失敗する無意味な候補になるだけなので省く。
' ----------------------------------------------------------------------------
Public Function GsCandidatePaths(ByVal cfgPath As String, ByVal wbDir As String, _
                                 ByVal searchDirs As String) As String
    Dim result As String: result = ""

    Dim c As String: c = Trim$(cfgPath)
    If LenB(c) > 0 Then result = AppendCandidate(result, c)

    Dim wb As String: wb = TrimTrailingSep(Trim$(wbDir))
    If LenB(wb) > 0 Then
        If LCase$(Left$(wb, 4)) <> "http" Then
            result = AppendCandidate(result, wb & "\Ghostscript\gswin32c.exe")
        End If
    End If

    Dim dirs() As String: dirs = Split(searchDirs, ";")
    Dim i As Long
    For i = LBound(dirs) To UBound(dirs)
        Dim forDir As String: forDir = GsCandidatesForFolder(dirs(i))
        If LenB(forDir) > 0 Then result = AppendCandidate(result, forDir)
    Next i

    GsCandidatePaths = result
End Function

' ----------------------------------------------------------------------------
' GsCandidatesForFolder - 1つのフォルダから「直下」「bin直下」の2候補を
'   "|" 区切りで返す(R9)。案内カードでフォルダを選んでもらった直後の確認と、
'   GsCandidatePathsのsearchDirs展開の両方から使う共通ロジック(規約は1つ:
'   直下 → 無ければbin直下)。空・空白のみのフォルダは候補を作らず ""。
' ----------------------------------------------------------------------------
Public Function GsCandidatesForFolder(ByVal folderPath As String) As String
    Dim d As String: d = TrimTrailingSep(Trim$(folderPath))
    If LenB(d) = 0 Then Exit Function
    GsCandidatesForFolder = d & "\gswin32c.exe|" & d & "\bin\gswin32c.exe"
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

Private Function AppendCandidate(ByVal existing As String, ByVal newPart As String) As String
    If LenB(existing) = 0 Then
        AppendCandidate = newPart
    Else
        AppendCandidate = existing & "|" & newPart
    End If
End Function

' パスを二重引用符で囲む(Windowsのファイル名に " は使えないのでエスケープ不要)。
Private Function Quoted(ByVal s As String) As String
    Quoted = Chr$(34) & s & Chr$(34)
End Function

Private Function TrimTrailingSep(ByVal s As String) As String
    Dim t As String: t = s
    Do While Len(t) > 0
        If Right$(t, 1) = "\" Or Right$(t, 1) = "/" Then
            t = Left$(t, Len(t) - 1)
        Else
            Exit Do
        End If
    Loop
    TrimTrailingSep = t
End Function
