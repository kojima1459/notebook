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
'   ・Ghostscript実行ファイルの解決候補列挙(2026-07-31 R9): 同梱配布・IT
'     焼き込み(ghostscript_search_dirs)の候補パスを「どの順で・どう組み立てる
'     か」だけを切り出した(GsCandidatePaths / GsCandidatesForFolder)。実在
'     確認(Dir$)とFileDialogは副作用なのでoptVision側の仕事のまま。戻り値が
'     "|" 区切りの1本なのは、LibreOffice Basicが `As String()` の関数宣言で
'     コンパイルごとハングするため(modUtil.SplitKeepNonEmptyと同じ回避策)。
'   ・テキストPDFのCOM無し抽出(2026-07-31 R10-3): Word/AcrobatのCOMが端末
'     ポリシーで塞がれていてもテキストPDFを取り込めるよう、txtwriteデバイスの
'     「コマンド組み立て」と「採否判定」をここへ置いた(BuildGsTextCommand /
'     GsTextVerdict)。実行するのは optGsTxt。
' ============================================================================

' 公式ツール(gazou版)と同じ規約。ページ番号3桁ゼロ埋め。
Private Const JPG_PREFIX As String = "page_"
Private Const JPG_SUFFIX As String = ".jpg"
Private Const JPG_PATTERN As String = "page_%03d.jpg"
Private Const DONE_FLAG_NAME As String = "done.flag"
Private Const TEMP_DIR_PREFIX As String = "nxocr_"
' R11-D(監査3 H-2): Ghostscriptの標準出力・標準エラーの受け皿。
' MOTW/AppLocker/EDRに実行を止められた場合、GSは何も出力せず終了コードだけを
' 返すため、従来は「フラグは出たのに1枚もできていない」としか観測できなかった。
Private Const GS_LOG_NAME As String = "gs_out.log"

' dpiの安全範囲(既定150は公式帳票OCR版と同値。300は倍重い)。
Private Const DPI_MIN As Long = 72
Private Const DPI_MAX As Long = 600
Private Const DPI_DEFAULT As Long = 150

' 1ファイルあたりのページ上限の安全範囲(configの値が壊れていても暴走しない)。
Private Const PAGES_MIN As Long = 1
Private Const PAGES_MAX As Long = 300

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
'   firstPage  : 描画する先頭ページ(省略時1。R14-4aのバッチ描画で使う)
'   戻り値の形(パスは全て二重引用符で囲む):
'     "<gs>" -dSAFER -dNOPAUSE -dBATCH -sDEVICE=jpeg -r150
'     -dTextAlphaBits=4 -dGraphicsAlphaBits=4 -dFirstPage=1 -dLastPage=21
'     -sOutputFile="<out>" "<pdf>"
'   2026-08-03(R14-4a): 100ページを一度に描かせると、画像が全部できるまで
'   OCRが1文字も始まらず、途中経過も出せない。20ページずつ複数回起動する
'   ため firstPage を足した(省略時の字面は従来と1文字も変えない=
'   modTestsPure4 のゴールデンテストがそのまま生きる)。
' ----------------------------------------------------------------------------
Public Function BuildGsCommand(ByVal gsExe As String, ByVal pdfPath As String, _
                               ByVal outPattern As String, ByVal dpi As Long, _
                               ByVal lastPage As Long, _
                               Optional ByVal firstPage As Long = 1) As String
    Dim r As Long: r = SafeDpi(dpi)
    Dim firstN As Long: firstN = firstPage
    If firstN < 1 Then firstN = 1
    Dim lastN As Long: lastN = lastPage
    If lastN < firstN Then lastN = firstN

    BuildGsCommand = Quoted(gsExe) & _
        " -dSAFER -dNOPAUSE -dBATCH" & _
        " -sDEVICE=jpeg -r" & CStr(r) & _
        " -dTextAlphaBits=4 -dGraphicsAlphaBits=4" & _
        " -dFirstPage=" & CStr(firstN) & " -dLastPage=" & CStr(lastN) & _
        " -sOutputFile=" & Quoted(outPattern) & _
        " " & Quoted(pdfPath)
End Function

' ----------------------------------------------------------------------------
' BuildRunCommand - 実行するコマンドライン(GS本体に標準出力のログ付けを足す)。
'   R11-D〜R39 は cmd.exe /s /c "(gs) 1>log 2>&1 & (if errorlevel …) >flag" で
'   ログと完了フラグを cmd に作らせていたが、2026-09-08 R40 F1 で廃止
'   (AMSI がマクロ型マルウェアの手口と誤検知し Office が強制終了。optGsProc
'   冒頭)。今は GS 本体のコマンドに -sstdout=<log>(GS 自身の機能)を実行
'   ファイル直後(入力PDFより前)へ挿すだけ。stderr は落ちない(記録・R41)。
'   完了フラグ(doneFlagPath)は optGsProc.SyncDoneFlag が GS の終了コードで
'   書くので、ここでは使わない(引数は呼び出し側との契約維持のため残す)。
' ----------------------------------------------------------------------------
Public Function BuildRunCommand(ByVal gsCommand As String, ByVal doneFlagPath As String, _
                                ByVal logPath As String) As String
    gsCommand = LTrim$(gsCommand)
    Dim ins As String: ins = " -sstdout=" & Quoted(logPath)
    Dim p As Long
    If Left$(gsCommand, 1) = Chr$(34) Then
        p = InStr(2, gsCommand, Chr$(34) & " ")
        If p > 0 Then p = p + 1
    Else
        p = InStr(gsCommand, " ")
    End If
    If p = 0 Then
        BuildRunCommand = gsCommand & ins
    Else
        BuildRunCommand = Left$(gsCommand, p - 1) & ins & Mid$(gsCommand, p)
    End If
End Function

' ----------------------------------------------------------------------------
' GsLogFor - GSの標準出力/標準エラーを落とすログファイルのフルパス(R11-D)。
' ----------------------------------------------------------------------------
Public Function GsLogFor(ByVal folderPath As String) As String
    GsLogFor = TrimTrailingSep(folderPath) & "\" & GS_LOG_NAME
End Function

' ----------------------------------------------------------------------------
' GsExitCodeFromFlag - 完了フラグの中身からGSの終了コードを読む(R11-D)。
'   0以上   : その終了コード(0=正常終了)
'   -1      : 判定不能(空・数字以外・旧形式の "done" など)。
'             呼び出し元は「不明」として扱い、失敗と決めつけない
'             (フラグの形式が変わっても取込を壊さないための安全側)。
'   前後の空白・改行は落としてから判定する(echo は末尾に空白と改行を残す)。
'   負の終了コード("-1073741819" 等のクラッシュ系)は「非0の失敗」として
'   255 に丸める。生のフラグ内容は呼び出し元が err_log へそのまま残すので、
'   実際の値が分からなくなることはない。
' ----------------------------------------------------------------------------
Public Function GsExitCodeFromFlag(ByVal flagText As String) As Long
    GsExitCodeFromFlag = -1

    Dim t As String: t = flagText
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, vbTab, "")
    t = Trim$(t)
    If LenB(t) = 0 Then Exit Function

    If Left$(t, 1) = "-" Then
        Dim body As String: body = Mid$(t, 2)
        If LenB(body) = 0 Then Exit Function
        Dim k As Long
        For k = 1 To Len(body)
            If InStr("0123456789", Mid$(body, k, 1)) = 0 Then Exit Function
        Next k
        GsExitCodeFromFlag = 255
        Exit Function
    End If

    Dim i As Long
    For i = 1 To Len(t)
        If InStr("0123456789", Mid$(t, i, 1)) = 0 Then Exit Function
    Next i
    If Len(t) > 9 Then Exit Function        ' Longの桁あふれ防止

    GsExitCodeFromFlag = CLng(t)
End Function

' ----------------------------------------------------------------------------
' BuildGsTextCommand - Ghostscriptの txtwrite デバイスで「PDFに埋まっている
'   文字をそのまま抜き出す」コマンドラインを組み立てる(2026-07-31 R10-3)。
'   gsExe   : gswin32c.exe のフルパス
'   pdfPath : 変換元のPDFのフルパス
'   outTxt  : 呼び出し側の従来の書き出し先("<folder>\gstext.txt")。
'             フォルダ部だけを使う(2026-09-07 R39 F001)。
'   戻り値の形(パスは全て二重引用符で囲む・BuildGsCommandと同じ規約):
'     "<gs>" -dSAFER -dNOPAUSE -dBATCH -sDEVICE=txtwrite
'     -sOutputFile="<folder>\gstext_%04d.txt" "<pdf>"
'   2026-09-07(R39 F001): 同梱GSのtxtwriteは改ページ文字Chr(12)を書き出さず、
'   Split(txt,Chr(12))前提のGsPageBoundsが全ページp.1に潰れていた(受入テスト
'   実測)。ページごとに別ファイル("%04d"はGSがページ番号で置換)へ出させ、
'   読む側(ReadPageFilesJoined)がChr(12)を足しながら連結する形へ変えた。
'   連結後の判定(GsPageBounds等)は不変。
'   出力先の指定は -o ではなく -sOutputFile + -dNOPAUSE -dBATCH にした。
'   -o は「-sOutputFile と -dBATCH -dNOPAUSE をまとめた省略形」で意味は同じだが、
'   BuildGsCommand(jpeg側)と字面を揃えておく方が、両者を見比べたときに
'   引用符の付け忘れ等の差分が目で見つけやすい。
'   標準出力のログ付け(-sstdout)は BuildRunCommand をそのまま共用する(R40 F1)。
' ----------------------------------------------------------------------------
Public Function BuildGsTextCommand(ByVal gsExe As String, ByVal pdfPath As String, _
                                   ByVal outTxt As String) As String
    Dim outPattern As String: outPattern = FolderOfPath(outTxt) & "\gstext_%04d.txt"
    BuildGsTextCommand = Quoted(gsExe) & _
        " -dSAFER -dNOPAUSE -dBATCH" & _
        " -sDEVICE=txtwrite" & _
        " -sOutputFile=" & Quoted(outPattern) & _
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
' ClassifyGsTextResult - txtwrite が「完了したのに本文が空」だったときの
'   本当の理由を1語に落とす(2026-08-03 R13-1b)。
'   rc            : 完了フラグから読んだGSの終了コード(-1=不明)
'   txtLen        : 読み出せた出力テキストの長さ(0=空)
'   gsOutHasPages : gs_out.log に "Processing pages 1 through N" があったか
'   戻り値の4値:
'     "ok"        本文が取れている。分類の対象外(呼び出し元は通常経路へ)。
'     "image"     GSは正常終了(rc=0)し、ページ処理も走ったのに文字が1字も
'                 出なかった=文字層の無いスキャンPDF。GsTextVerdict の
'                 "image" と同じ扱いで E0303(OCR経路)へ回す。
'     "gsfail"    GSの実行そのものが失敗した(rc<>0)。従来どおりE0302で
'                 Word/Acrobat連鎖へ譲る。gs_outのtailを添えて記録する。
'     "flagdelay" 完了フラグは出たが中身(終了コード)が最後まで読めなかった。
'                 rcが不明なので image と決めつけてはならない(誤ってOCRへ
'                 回すと、実際にはGSが落ちていた資料が「画像PDF」として
'                 記録され原因が消える)。E0302 + detail に flag_delayed。
'   なぜ切り出したか(RC1): 実機第2報では「txtwrite出力を読めず」の
'   E0302 が画像PDF判定より先に発火し、44頁の約款がWordのゴミ本文で
'   「登録成功」になった。分類の順序こそが事故の本体なので、副作用ゼロの
'   関数にして真理表をLOテストで固定する(憲章§4-5)。
' ----------------------------------------------------------------------------
'     "emptyinput" 渡されたPDFそのものが0バイトだった(2026-08-03 R14-3c)。
'                 GSは空の入力を「空のPS」として rc=0/バナーだけで即終了する
'                 ため、rcとログだけを見ると image と見分けがつかない。
'                 実機第3報 RC3(0バイトコピー)では、これが「画像PDF」として
'                 OCRへ回り、描画0枚だけが記録されて原因が消えていた。
'                 材料(入力サイズ)があるときは他のどの分岐よりも先に言い切る。
'   gsInputBytes  : GSへ渡した入力ファイルの実バイト数(-1=不明・既定)。
' ----------------------------------------------------------------------------
Public Function ClassifyGsTextResult(ByVal rc As Long, ByVal txtLen As Long, _
                                     ByVal gsOutHasPages As Boolean, _
                                     Optional ByVal gsInputBytes As Long = -1) As String
    If txtLen > 0 Then
        ClassifyGsTextResult = "ok"
        Exit Function
    End If

    ' 入力が0バイトなら、rcもログも見る意味が無い(GSは何も読んでいない)。
    If gsInputBytes = 0 Then
        ClassifyGsTextResult = "emptyinput"
        Exit Function
    End If

    If rc < 0 Then
        ClassifyGsTextResult = "flagdelay"
        Exit Function
    End If

    If rc = 0 And gsOutHasPages Then
        ClassifyGsTextResult = "image"
        Exit Function
    End If

    ClassifyGsTextResult = "gsfail"
End Function

' ----------------------------------------------------------------------------
' GsTotalPagesFromLog - gs_out.log の "Processing pages 1 through 44." から
'   総ページ数(44)を読む(2026-08-03 R13-1c/1d)。見つからなければ 0。
'   GSはこの行を処理の【最初】に出すので、呼び出し元はログの先頭を渡す。
'   ログはCP932(コンソールのコードページ)の生バイトを素のOpenで読んだもので、
'   改行はスペースへ潰されている前提。大小は区別せず探す(GSの版差の保険)。
' ----------------------------------------------------------------------------
Public Function GsTotalPagesFromLog(ByVal logText As String) As Long
    Dim marker As String: marker = "through"
    Dim p As Long: p = InStr(1, logText, marker, vbTextCompare)
    If p = 0 Then Exit Function
    ' "Processing pages" が前にあることまで確かめる(別の through との取り違え防止)。
    If InStr(1, Left$(logText, p), "Processing pages", vbTextCompare) = 0 Then Exit Function
    GsTotalPagesFromLog = FirstNumberFrom(logText, p + Len(marker))
End Function

' ----------------------------------------------------------------------------
' GsPagesFromLog - gs_out.log に現れる "Page N" のうち最大のNを返す
'   (2026-08-03 R13-1c)。見つからなければ 0。
'   進捗が動いているかの唯一の観測点なので、末尾だけを渡されても
'   (途中で切れた数字があっても)最大値だけを拾えば足りる。
'   "Processing pages" の "pages" を拾わないよう、比較は大小を区別する
'   (GSは進捗行を必ず "Page " と大文字で始める)。
' ----------------------------------------------------------------------------
Public Function GsPagesFromLog(ByVal logText As String) As Long
    Dim marker As String: marker = "Page "
    Dim p As Long: p = InStr(1, logText, marker, vbBinaryCompare)
    Do While p > 0
        Dim n As Long: n = FirstNumberFrom(logText, p + Len(marker))
        If n > GsPagesFromLog Then GsPagesFromLog = n
        p = InStr(p + Len(marker), logText, marker, vbBinaryCompare)
    Loop
End Function

' ----------------------------------------------------------------------------
' GsWaitBanner - 本文抽出の待ち時間に出す進捗バナーの文面(R13-1d)。
'   「本文抽出中… 12/44ページ (経過 35秒) 残り約93秒」
'   ・総ページ数が読めるまでは経過秒だけを出す(嘘の分母を出さない)。
'   ・残り時間はページが3枚以上進んでから出す(1枚目だけの実績で出すと
'     表示が跳ね回り、かえって「動いていない」不安を生む)。
'   表示専用の純関数。ここで組み立てるのは文字列だけで、描くのは呼び出し元。
' ----------------------------------------------------------------------------
Public Function GsWaitBanner(ByVal pagesSeen As Long, ByVal totalPages As Long, _
                             ByVal elapsedSec As Long) As String
    Dim s As String: s = "本文抽出中… "
    Dim el As Long: el = elapsedSec
    If el < 0 Then el = 0

    If totalPages > 0 Then
        Dim p As Long: p = pagesSeen
        If p < 0 Then p = 0
        If p > totalPages Then p = totalPages
        s = s & p & "/" & totalPages & "ページ "
    End If

    s = s & "(経過 " & el & "秒)"

    If totalPages > 0 And pagesSeen >= 3 And pagesSeen < totalPages And el > 0 Then
        Dim remain As Double
        remain = (CDbl(totalPages - pagesSeen) * CDbl(el)) / CDbl(pagesSeen)
        Dim remSec As Long: remSec = CLng(Int(remain + 0.5))
        If remSec < 1 Then remSec = 1
        s = s & " 残り約" & remSec & "秒"
    End If

    GsWaitBanner = s
End Function

' ----------------------------------------------------------------------------
' CleanTextLen - 空白類(半角/全角スペース・タブ・改行・改ページ)を除いた
'   文字数(R10cで optGsTxt から移設。純ロジックなのでここが正しい置き場)。
'   txtwriteは画像PDFに対しても改ページ文字だけは律儀に書き出すので、素のLenで
'   判定すると「文字がある」と誤認する。1文字ずつ走査するとページ数の多いPDFで
'   無駄に重いため、Replaceで落としてから数える。
' ----------------------------------------------------------------------------
Public Function CleanTextLen(ByVal s As String) As Long
    CleanTextLen = modUtilText.CleanTextLen(s)
End Function

' ----------------------------------------------------------------------------
' GsPageCount - txtwrite出力が実質何ページ分あるかを数える(R10c)。
'   2026-07-31(R11-F2): 添字計算の実体を modUtilText.GsPageBounds へ一本化した。
'   従来は modExtractor.BuildPagesFromGsText と同じ規約を2箇所で書いており、
'   ズレると GsTextVerdict の分母と出典ページ番号が食い違うため、
'   modTestsPure6 の突き合わせテストで守っていた(実装が1本になったので、
'   突き合わせではなく modUtilText.GsPageBounds の単体テストで固定する)。
' ----------------------------------------------------------------------------
Public Function GsPageCount(ByVal txt As String) As Long
    Dim firstIdx As Long, lastIdx As Long
    GsPageCount = modUtilText.GsPageBounds(txt, firstIdx, lastIdx)
End Function

' 2026-09-07(R39 F001): ページ別出力("gstext_%04d.txt")を連結する側。無い・
' 読めないページは空扱いで末尾にChr(12)を足す。連結後(GsPageBounds等)は不変。
' ファイルI/Oを伴う3関数はPureテスト対象外(PageTxtName/JoinPageTextsのみ対象)。

' GSの"%04d"と同じ4桁ゼロ埋め(桁あふれはそのまま桁数増)。
Public Function PageTxtName(ByVal folderPath As String, ByVal i As Long) As String
    PageTxtName = folderPath & "\gstext_" & Format$(i, "0000") & ".txt"
End Function

' parts(1..n)を、各要素末尾にChr(12)を付けて連結(n<1は空。配列は1始まり)。
Public Function JoinPageTexts(ByRef parts() As String, ByVal n As Long) As String
    If n < 1 Then Exit Function
    ' R39 M4: 逐次連結は O(n²)。先頭 n 要素だけを Join で1回に。
    Dim i As Long
    Dim tmp() As String: ReDim tmp(1 To n)
    For i = 1 To n
        tmp(i) = parts(i)
    Next i
    JoinPageTexts = Join(tmp, Chr$(12)) & Chr$(12)
End Function

' 実在する"gstext_*.txt"の最大番号(hintNより大きければそちら)。
Public Function MaxPageFileIndex(ByVal folderPath As String, ByVal hintN As Long) As Long
    Dim maxN As Long: maxN = hintN
    If maxN < 0 Then maxN = 0
    Dim fn As String
    On Error Resume Next
    fn = Dir$(folderPath & "\gstext_*.txt")
    On Error GoTo 0
    Do While LenB(fn) > 0
        Dim p As Long: p = InStr(1, fn, "gstext_", vbTextCompare)
        If p > 0 Then
            Dim v As Long: v = FirstNumberFrom(fn, p + 7)
            If v > maxN Then maxN = v
        End If
        ' R39 M3: Dir$() 失敗時に fn が残って無限ループしないよう先に空へ。
        fn = ""
        On Error Resume Next
        fn = Dir$()
        On Error GoTo 0
    Loop
    MaxPageFileIndex = maxN
End Function

' N=Max(totalPages,実在最大番号)まで1..Nを読み連結(N=0は空)。無い・読めない
' ページは空扱い、最初の失敗だけerrNum/errDescへ返す。
Public Function ReadPageFilesJoined(ByVal folderPath As String, ByVal totalPages As Long, _
                                    ByRef errNum As Long, ByRef errDesc As String) As String
    errNum = 0
    errDesc = ""
    Dim n As Long: n = totalPages
    If n < 0 Then n = 0
    Dim maxFound As Long: maxFound = MaxPageFileIndex(folderPath, n)
    If maxFound > n Then n = maxFound
    If n < 1 Then Exit Function
    ' R39 M4: ログが化けて総頁が9桁でも配列を確保しない安全弁(打切りは連結後)。
    If n > 5000 Then n = 5000

    Dim parts() As String: ReDim parts(1 To n)
    Dim gotErr As Boolean
    Dim i As Long
    For i = 1 To n
        Dim t As String: t = ""
        Dim en As Long, ed As String
        If Not modUtilText.ReadTextFileUtf8(PageTxtName(folderPath, i), t, en, ed) Then
            t = ""
            If Not gotErr Then errNum = en: errDesc = ed: gotErr = True
        End If
        parts(i) = t
    Next i
    ReadPageFilesJoined = JoinPageTexts(parts, n)
End Function

' 最大番号のページファイルのサイズ(無ければ0)。WaitGsTextDoneの進捗監視用
' (optGsTxt側はPrivateで貸せないためここでFileLenを使う)。
Public Function LatestPageFileSize(ByVal folderPath As String) As Double
    ' R39 M2/B2: 進捗信号は単調増加が要る(newSize > lastSize)。全ページの合計に、
    ' 書き込み中で FileLen が更新されないぶんを最新ファイルの LOF で補う。
    Dim maxN As Long: maxN = MaxPageFileIndex(folderPath, 0)
    If maxN <= 0 Then Exit Function
    Dim total As Double: total = 0#
    Dim i As Long
    For i = 1 To maxN
        Dim p As String: p = PageTxtName(folderPath, i)
        Dim n As Double: n = 0#
        On Error Resume Next
        n = CDbl(FileLen(p))
        Err.Clear                              ' B2: 直前の Err を持ち越さない
        If i = maxN Then
            Dim fn As Long: fn = FreeFile
            Open p For Binary Access Read As #fn
            If Err.Number = 0 Then
                If CDbl(LOF(fn)) > n Then n = CDbl(LOF(fn))
                Close #fn                      ' Open が成功した経路では必ず閉じる
            End If
            Err.Clear
        End If
        On Error GoTo 0
        total = total + n
    Next i
    LatestPageFileSize = total
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

' ----------------------------------------------------------------------------
' BatchCountFor / BatchBoundsFor - ページ描画を何回に分けるか(R14-4a)。
'   totalPages : 描画したい総ページ数(=RenderCapFor の戻り値)
'   batchSize  : 1回のGS起動で描かせるページ数(20)
'   batchIdx   : 1始まりのバッチ番号
'   BatchBoundsFor の戻り値: "<先頭ページ>|<最終ページ>"(範囲外なら "")。
'   配列を返せない(LibreOffice Basicが As String() の宣言でハングする既知の
'   制約。GsCandidatePaths と同じ回避策)ので "|" 区切りの1本にしている。
'   なぜ分けるのか(実機第3報 RC4): 100ページを1回で描かせると全部できるまで
'   OCRが始まらず、%TEMP% にも100枚が同時に載る。20ページずつなら最初の20枚は
'   数十秒で読み始められ、一時領域も20枚ぶんで頭打ちになる。
' ----------------------------------------------------------------------------
Public Function BatchCountFor(ByVal totalPages As Long, ByVal batchSize As Long) As Long
    If totalPages < 1 Or batchSize < 1 Then Exit Function
    BatchCountFor = ((totalPages + batchSize - 1) \ batchSize)
End Function

Public Function BatchBoundsFor(ByVal totalPages As Long, ByVal batchSize As Long, _
                               ByVal batchIdx As Long) As String
    If totalPages < 1 Or batchSize < 1 Or batchIdx < 1 Then Exit Function

    Dim firstP As Long: firstP = (batchIdx - 1) * batchSize + 1
    If firstP > totalPages Then Exit Function

    Dim lastP As Long: lastP = firstP + batchSize - 1
    If lastP > totalPages Then lastP = totalPages

    BatchBoundsFor = CStr(firstP) & "|" & CStr(lastP)
End Function

' 2026-08-04(R15-5b): OcrPageBanner / OcrCapMemoFor / OcrAbortMemoFor /
' RemainingWaitSec は optOcrEta へ【移設】した。進捗の見せ方と打ち切りの
' 言い方(=利用者への伝え方)は、GSコマンドの組み立てとページ上限の算数
' (=Ghostscriptの都合)とは別の関心事で、本モジュールが30,000字上限まで
' 残り2,068字となり追記できなくなっていたため(憲章§4-6)。ここは純減のみ。

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

' 位置 startPos 以降で最初に現れる連続した半角数字を整数として返す(無ければ0)。
' 数字の前の空白は読み飛ばすが、数字以外の文字に当たった時点で諦める
' (「Page abc」のような想定外の行を無理に解釈しないため)。
' 桁あふれ防止に9桁で打ち切る(ページ数がそれを超えることはない)。
Private Function FirstNumberFrom(ByVal s As String, ByVal startPos As Long) As Long
    Dim i As Long: i = startPos
    If i < 1 Then i = 1
    Do While i <= Len(s)
        If Mid$(s, i, 1) <> " " Then Exit Do
        i = i + 1
    Loop

    Dim digits As String: digits = ""
    Do While i <= Len(s)
        Dim c As String: c = Mid$(s, i, 1)
        If InStr("0123456789", c) = 0 Then Exit Do
        digits = digits & c
        If Len(digits) >= 9 Then Exit Do
        i = i + 1
    Loop

    If LenB(digits) > 0 Then FirstNumberFrom = CLng(digits)
End Function

' パスの末尾要素を除いたフォルダ部("C:\t\gstext.txt" -> "C:\t")。
Private Function FolderOfPath(ByVal p As String) As String
    Dim i As Long: i = InStrRev(p, "\")
    If i > 0 Then FolderOfPath = Left$(p, i - 1) Else FolderOfPath = p
End Function

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
