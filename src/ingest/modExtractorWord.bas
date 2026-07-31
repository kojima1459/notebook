Attribute VB_Name = "modExtractorWord"
Option Explicit

' 文書の開き方。
'   1 = 自分でWordを起動(非表示) + パスワード引数
'   2 = 自分でWordを起動(非表示)
'   3 = 自分でWordを起動(表示)
'   4 = 既に起動しているWordへ相乗りする(GetObject。自分では起動しない)
' 実機で通った開き方をセッション中おぼえておく(0=未確定)。
Private mPreferredMode As Long
Private Const OPEN_MODE_MAX As Long = 4
Private Const MODE_ATTACH As Long = 4

' どのCOM呼び出しで落ちたかを表す段階名(err_log / usage_log に出す)。
' 実機でしか再現しない失敗を、次に推測ではなく事実から追うための手掛かり。
Private Const STEP_CREATE As String = "Word起動"
Private Const STEP_VISIBLE As String = "表示切替"
Private Const STEP_OPEN As String = "文書オープン"
Private Const STEP_PAGES As String = "ページ数取得"
Private Const STEP_TEXT As String = "本文取り出し"
Private Const STEP_CLOSE As String = "後始末"

' ============================================================================
' modExtractorWord - PDF/Word抽出(Word.Application COM経由)
' ----------------------------------------------------------------------------
' 役割:
'   Word 2013以降はPDFを「リフロー」して編集可能なテキストへ変換して開ける。
'   複雑な表組み・多段組みの約款PDFではリフローが崩れることがあるが、
'   埋め込みに使える程度の平文は十分に回収できる。
'   docx/docは素直にWordがネイティブに開く。
'   ページ単位のテキストは Document.Range をページ区切りまでスキャンして
'   取り出す(チャンクに正しいページ番号を付けるため)。
'
' 流用元: /home/user/notebook/src/admin/modExtractorWord.bas(V2資産・コピー元。
'   変更禁止)。
'
' 適応要件(MASTER_SPEC §7.2):
'   ・V1は Err.Raise で失敗を通知していたが、本モジュールは
'     Boolean 戻り値 + errDetail(ByRef)へ変換する。呼び出し元 modExtractor
'     がE0302等の契約エラーコードへ変換する(本モジュールはエラーコードを
'     持たない・生の詳細文字列のみ返す)。
'   ・maxPages(呼び出し側が modConfig.max_pages_per_file から解決した値)を
'     超えるページを持つ文書は、超過分を抽出せず先頭 maxPages ページのみ
'     処理し、truncated(ByRef)を True にする(§7.2 PARTIAL_PAGES)。
'   ・Word.Application は必ず On Error で包み、成功・失敗いずれの経路でも
'     doc.Close / word.Quit を試みる(V1のFailedハンドラの作法をそのまま
'     踏襲。プロセスが残留するとユーザーPCにWordプロセスが積み上がる事故に
'     なるため必須)。
'   ・Mac等COM非対応環境(CreateObject失敗)は実行時エラー429を検知して
'     丁寧な案内文を errDetail に含める(§13)。
' ============================================================================

' 【2026-07-16 実機err#462対応】Word COMは「リモート サーバーがないか、
' 使用できる状態ではありません」(実行時エラー462)で散発的に失敗する
' (直前のWord.Quitの残骸プロセスや、起動直後のRPC未確立が原因の
' 実機Windows特有の揺らぎ。LibreOfficeでは再現しない)。1回の失敗で
' 資料を「取込失敗」にせず、少し待ってからまっさらなWordで1回だけ
' やり直す(2回目も失敗したら本当の失敗として報告する)。
' 【2026-07-30 実機ログからの確定診断】.doc/.docx/.pdf を全滅させていた
' err#462 は原因ではなく【後始末の失敗】だった。
'
' VBAには「有効なハンドラ」と「稼働中のハンドラ」の区別があり、ハンドラへ
' 飛んでから Resume / Exit までの間に起きたエラーは、そのプロシージャでは
' 捕まえられず呼び出し元へ投げ返される。On Error Resume Next を書いても効かない。
' 下の Failed: は doc.Close / word.Quit を On Error Resume Next で囲んでいたが
' それが無効だったため、
'   ・本当の最初のエラー(errDetail に正しく入れていた)は返らずに捨てられ、
'   ・後始末で死んだWordを叩いた err#462 だけが modExtractor.ExtractFile まで
'     飛んで記録され、
'   ・その結果 Extract の開き方フォールバックも、PDFの Acrobat
'     フォールバックも【一度も走っていなかった】。
' ログ上の根拠: .pdf の失敗が "Word: … / Acrobat: …"(ExtractPdfWithFallback
' の書式)ではなく "err#462: … [localcopy=ok]"(ExtractFile のハンドラの書式)
' で出ていた。つまり戻り値ではなく例外で抜けている。
'
' 同時に分かること: word が Nothing なら後始末は何も叩かないので上書きは
' 起きない。上書きが起きていた以上、CreateObject は成功している。
' つまりこの端末では【Wordは起動するが、その直後に死ぬ/届かなくなる】。
' 自分で起動したWordが使えない環境(ライセンスの初回起動待ち、EDR/DLPが
' Officeの子プロセス起動を止める等)は実在するので、起動しない道を用意する。
'
' 開き方を「静かな順」に並べて、通るまで降りていく。
'   1. 自分で起動(非表示)+パスワード引数 … いちばん行儀が良い
'   2. 自分で起動(非表示)                … 引数が原因のときここで通る
'   3. 自分で起動(表示)                  … 非表示だと落ちる環境向け
'   4. 起動中のWordへ相乗り(GetObject)   … 自分では起動できない環境向け
' 4 は利用者のWordなので Quit しない(編集中のWordを閉じたら大事故)。
' 起動そのものに失敗したときは 2・3 を試すのは無駄なので 4 へ直行する。
' 一度成功した開き方はセッション中おぼえておき、次からそこから始める
' (毎回1から試すと、100件のフォルダ同期で失敗待ちの時間を100回払う)。
'
' origPath(2026-07-30 レビュー4-D): modExtractor は保護ビュー対策として
' ローカル一時コピーを開くため、path には一時コピーのパスが来る。利用者が
' 編集中の文書を掴んでいないかを調べるには原本のパスも要るので、呼び出し元
' から受け取る(省略時は path だけで判定する)。
Public Function Extract(ByVal path As String, ByVal maxPages As Long, _
                        ByRef pages() As ExtractedPage, ByRef truncated As Boolean, _
                        ByRef errDetail As String, _
                        Optional ByVal origPath As String = "") As Boolean
    Dim startMode As Long
    startMode = mPreferredMode
    If startMode < 1 Or startMode > OPEN_MODE_MAX Then startMode = 1

    ' どのCOM呼び出しで落ちたかを errDetail の先頭に付け、err_log へ残す。
    ' 実機でしか出ない失敗を、次から推測ではなく事実で追えるようにするため。
    Dim failedStep As String
    Dim mode As Long
    mode = startMode
    Do While mode <= OPEN_MODE_MAX
        Dim attempt As Long
        For attempt = 1 To 2
            If TryExtractOnce(path, maxPages, pages, truncated, errDetail, mode, failedStep, origPath) Then
                If mPreferredMode <> mode Then
                    mPreferredMode = mode
                    On Error Resume Next
                    modLog.LogUsage "word_open_mode", "", _
                        "この端末では開き方" & mode & "(" & OpenModeName(mode) & ")で成功しました"
                    On Error GoTo 0
                End If
                Extract = True
                Exit Function
            End If
            If attempt = 1 Then WaitBriefly 800
        Next attempt
        If failedStep = STEP_CREATE And mode < MODE_ATTACH Then
            ' Word を起動できないなら、開き方(表示・引数)を変えても届かない。
            ' 起動方法そのものが違う「相乗り」へ直行する(待ち時間の節約)。
            mode = MODE_ATTACH
        Else
            mode = mode + 1
        End If
    Loop

    ' 何をどこまで試して駄目だったのかを、あとから追える形で1行残す。
    On Error Resume Next
    modLog.LogUsage "word_open_fail", "", _
        "開き方" & startMode & "から" & OPEN_MODE_MAX & "まで試して失敗(最後に落ちた段階: " & failedStep & ")"
    On Error GoTo 0

    ' 一度おぼえた開き方でも駄目だったなら、環境が変わった可能性がある。
    ' 次回は1から試し直す。
    mPreferredMode = 0
    Extract = False
End Function

' 開き方の名前(ログ用)。
Private Function OpenModeName(ByVal mode As Long) As String
    Select Case mode
        Case 1: OpenModeName = "非表示+パスワード無効化"
        Case 2: OpenModeName = "非表示"
        Case 3: OpenModeName = "表示"
        Case Else: OpenModeName = "起動中のWordへ相乗り"
    End Select
End Function

Private Function TryExtractOnce(ByVal path As String, ByVal maxPages As Long, _
                                ByRef pages() As ExtractedPage, ByRef truncated As Boolean, _
                                ByRef errDetail As String, ByVal openMode As Long, _
                                ByRef failedStep As String, _
                                Optional ByVal origPath As String = "") As Boolean
    truncated = False

    Dim word As Object
    Dim doc As Object
    Dim stepName As String
    ' このWordを自分で起動したのか(相乗りなら Quit してはいけない)。
    Dim ownsApp As Boolean
    ' この文書を自分で開いたのか(相乗り先で既に開かれていた文書なら
    ' 閉じてはいけない。閉じると利用者の未保存の編集がその場で消える)。
    Dim ownsDoc As Boolean

    On Error GoTo Failed
    stepName = STEP_CREATE
    If openMode >= MODE_ATTACH Then
        ' 自分では起動しない。既に開いているWordへ相乗りする。
        ' 起動できない端末(ライセンス初回起動待ち、EDR/DLPがOfficeの
        ' 子プロセス起動を止める等)でも、人が開いたWordなら使える。
        ' Wordが1つも開いていなければ 429 になり、案内文でそれを伝える。
        Set word = GetObject(, "Word.Application")
        ownsApp = False
    Else
        Set word = CreateObject("Word.Application")
        ownsApp = True
    End If

    ' 開き方は openMode で切り替える(Extract のコメント参照)。
    stepName = STEP_VISIBLE
    Dim showWord As Boolean
    If ownsApp Then
        showWord = (openMode >= 3)
        word.Visible = showWord
    Else
        ' 相乗り先は利用者のWord。表示状態は触らない(勝手に隠したら事故)。
        showWord = True
    End If
    word.DisplayAlerts = 0   ' wdAlertsNone

    ' 取り込む文書のマクロを走らせない(レビュー H-11)。
    ' 3 = msoAutomationSecurityForceDisable。出所の分からないファイルを
    ' 取り込むのは日常操作なので、そこが任意コード実行の経路になっていては
    ' いけない。
    ' 相乗り(mode 4)のときは利用者のWordの設定なので、必ず元へ戻す
    ' (自分で起動した分は Quit するので戻す必要はない)。
    Dim prevSecurity As Long: prevSecurity = -1
    On Error Resume Next
    If Not ownsApp Then prevSecurity = word.AutomationSecurity
    word.AutomationSecurity = 3
    Err.Clear
    On Error GoTo Failed

    ' ConfirmConversions:=False でPDFリフロー確認ダイアログを抑止する。
    ' PasswordDocument にダミーを渡すのは、暗号化文書に当たったときに
    ' パスワード入力ダイアログでフリーズさせないため(レビュー M-11)。
    ' ただしこの引数自体が環境によっては Open を失敗させうるので、
    ' openMode=1 のときだけ付ける(駄目なら 2 以降で外して試す)。
    ' 【2026-07-30 レビュー4-D】開く前に「もう開かれていないか」を必ず調べる。
    ' 相乗り(mode 4)先は利用者のWordなので、取り込もうとしている資料を本人が
    ' 開いたまま作業していることが普通にある。その状態で Documents.Open を
    ' 呼ぶと、Wordは新しい文書ではなく【その開いている文書オブジェクト】を
    ' 返す。抽出後の後始末 doc.Close 0 は wdDoNotSaveChanges なので、
    ' 未保存の編集が黙って消える(データ喪失)。
    ' 見つかったら開き直さず・閉じもせず、その文書からテキストだけ取る。
    stepName = STEP_OPEN
    Set doc = FindOpenDocument(word, path, origPath)
    ownsDoc = (doc Is Nothing)
    If ownsDoc Then
        If openMode <= 1 Then
            Set doc = word.Documents.Open( _
                FileName:=path, _
                ConfirmConversions:=False, _
                ReadOnly:=True, _
                AddToRecentFiles:=False, _
                PasswordDocument:="__mybookshelf_no_password__", _
                Visible:=showWord)
        Else
            Set doc = word.Documents.Open( _
                FileName:=path, _
                ConfirmConversions:=False, _
                ReadOnly:=True, _
                AddToRecentFiles:=False, _
                Visible:=showWord)
        End If
    Else
        ' 何が起きたのかを1行残す(実機で「開いたまま取り込んだ」ことを
        ' あとから確認できるようにする。障害ではないので usage_log 側)。
        On Error Resume Next
        modLog.LogUsage "word_reuse_open_doc", "", _
            "既に開かれている文書をそのまま読みました(閉じません): " & _
            modUtil.SafeLeft(modUtil.FileNameOf(path), 120)
        On Error GoTo Failed
    End If

    stepName = STEP_PAGES
    Dim pageCount As Long
    pageCount = doc.ComputeStatistics(2)   ' wdStatisticPages
    If pageCount < 1 Then pageCount = 1

    Dim loopCount As Long: loopCount = pageCount
    If loopCount > maxPages Then
        loopCount = maxPages
        truncated = True
    End If

    stepName = STEP_TEXT
    Dim tmp() As ExtractedPage: ReDim tmp(0 To loopCount - 1)
    Dim i As Long
    For i = 1 To loopCount
        tmp(i - 1).page = i
        tmp(i - 1).Text = ExtractPageText(doc, i, pageCount)
        ' 2026-07-31(R7 B-2): 1ページ分のCOMシーケンス(GoTo→範囲確定→Text)が
        ' 完全に終わった【あと】で1回だけメッセージを捌く。ページの途中に
        ' 置くと、Rangeを掴んだままイベントへ抜けることになり、その間に
        ' 文書が閉じられると掴んでいる参照が無効になる。300ページの約款でも
        ' 追加コストはページ数回のDoEventsだけ。
        DoEvents
    Next i

    stepName = STEP_CLOSE
    Cleanup doc, word, ownsApp, prevSecurity, ownsDoc

    pages = tmp
    TryExtractOnce = True
    Exit Function

Failed:
    failedStep = stepName
    errDetail = "[" & stepName & "/開き方" & openMode & "] " & _
                DescribeComError(Err.Number, Err.Description, openMode)
    ' 後始末は Cleanup(別Sub)に任せる。ハンドラ稼働中は On Error Resume Next
    ' が効かないため、ここで直接 doc.Close / word.Quit を叩くと、その失敗が
    ' 呼び出し元へ飛んで【本来の原因を上書きする】。呼ばれた側は新しい
    ' エラー文脈を持つので、Cleanup 内の On Error Resume Next は正しく働く。
    ' 2026-07-30 実機: この上書きで err#462 だけが記録され、開き方の
    ' フォールバックもPDFのAcrobatフォールバックも一度も走っていなかった。
    Cleanup doc, word, ownsApp, prevSecurity, ownsDoc
    TryExtractOnce = False
End Function

' ----------------------------------------------------------------------------
' Cleanup - 文書とWordの後始末。何が起きても黙って終わる。
'   稼働中のエラーハンドラから呼ばれる前提のため、必ず別プロシージャに置く
'   (同一プロシージャ内の On Error Resume Next は効かない)。
'   ownsApp=False は利用者のWordへ相乗りした場合。Quit してはいけないので
'   自分が開いた文書だけ閉じ、変更した設定を元へ戻す。
'   ownsDoc=False は「開く前から開かれていた文書」(レビュー4-D)。
'   doc.Close 0 は wdDoNotSaveChanges なので、閉じた瞬間に利用者の
'   未保存の編集が消える。自分が開いた文書だけを閉じる。
' ----------------------------------------------------------------------------
Private Sub Cleanup(ByRef doc As Object, ByRef app As Object, _
                    ByVal ownsApp As Boolean, ByVal prevSecurity As Long, _
                    ByVal ownsDoc As Boolean)
    On Error Resume Next
    If ownsDoc Then
        If Not doc Is Nothing Then doc.Close 0   ' wdDoNotSaveChanges
    End If
    Set doc = Nothing
    If Not app Is Nothing Then
        If ownsApp Then
            app.Quit 0
        ElseIf prevSecurity >= 0 Then
            app.AutomationSecurity = prevSecurity
        End If
    End If
    Set app = Nothing
    Err.Clear
End Sub

' ----------------------------------------------------------------------------
' FindOpenDocument - 対象ファイルが【そのWordで既に開かれている】なら、その
'   Document を返す(開かれていなければ Nothing)。レビュー4-D の判定本体。
' ----------------------------------------------------------------------------
'   比較は Document.FullName の大文字小文字無視。原本パスと一時コピーパスの
'   両方を見る(modExtractor は保護ビュー対策でローカル一時コピーを開くが、
'   コピーに失敗した資料は原本パスのまま渡ってくる。どちらの経路でも
'   利用者が編集中の文書を掴みうるので、片方だけ見ても穴が残る)。
'   走査そのものが失敗しても取込を止めない(Nothing を返して従来どおり
'   Open へ進む)。ここは「壊さないための保険」であって主経路ではない。
Private Function FindOpenDocument(ByVal app As Object, ByVal path1 As String, _
                                  ByVal path2 As String) As Object
    On Error Resume Next
    Dim n As Long: n = 0
    n = app.Documents.count
    Dim i As Long
    For i = 1 To n
        Dim fullName As String: fullName = ""
        fullName = app.Documents(i).fullName
        If SamePath(fullName, path1) Or SamePath(fullName, path2) Then
            Set FindOpenDocument = app.Documents(i)
            Exit For
        End If
    Next i
    Err.Clear
    On Error GoTo 0
End Function

' パスの一致判定(大文字小文字無視。空文字はどれとも一致させない)。
Private Function SamePath(ByVal a As String, ByVal b As String) As Boolean
    If LenB(a) = 0 Or LenB(b) = 0 Then Exit Function
    SamePath = (StrComp(a, b, vbTextCompare) = 0)
End Function

' Declareを使わない短い待機(リトライ前にWordプロセスの後始末を待つ)。
' Timerは0時に0へリセットされるため、経過時間が負になったら日跨ぎとみなし
' 即座に抜ける(旧実装は日跨ぎで最大24時間ループするバグがあった)。
Private Sub WaitBriefly(ByVal ms As Long)
    Dim t0 As Double: t0 = Timer
    Dim elapsedMs As Double
    Do
        elapsedMs = (Timer - t0) * 1000#
        If elapsedMs < 0 Then Exit Do
        If elapsedMs >= ms Then Exit Do
        DoEvents
    Loop
End Sub

' 指定ページの本文を、次ページ開始直前までの範囲として取り出す
' (最終ページは文書末尾まで)。
Private Function ExtractPageText(ByVal doc As Object, ByVal pageNum As Long, ByVal totalPages As Long) As String
    ' wdGoToPage = 1, wdGoToAbsolute = 1
    Dim startRange As Object, endRange As Object
    Set startRange = doc.GoTo(What:=1, Which:=1, count:=pageNum)
    If pageNum < totalPages Then
        Set endRange = doc.GoTo(What:=1, Which:=1, count:=pageNum + 1)
        ' 2026-07-28(レビュー M-14): ここは endRange.Start - 1 だった。
        ' 次ページの開始位置の1つ手前まで、という意図だが、Word の Range は
        ' End が排他(終端の1つ先)なので、-1 すると【各ページ末尾の1文字が
        ' 欠落】する。「100万円」の「円」が消える類で、300ページの約款なら
        ' 最大299箇所。改ページ文字が混じる分は文字列側で落とす。
        startRange.End = endRange.Start
    Else
        startRange.End = doc.Content.End
    End If
    ExtractPageText = TrimPageBreak(startRange.Text)
End Function

' Mac等COM不可環境向けの丁寧な案内文を生成する(§13)。
Private Function DescribeComError(ByVal errNum As Long, ByVal desc As String, _
                                  ByVal openMode As Long) As String
    If errNum = 429 And openMode >= MODE_ATTACH Then
        ' 相乗りモードでの429 = Wordが1つも起動していない。
        ' この端末は自分でWordを起動できないので、人が開いてもらう必要がある。
        DescribeComError = "この端末ではExcelからWordを起動できないため、" & _
            "すでに開いているWordを使おうとしましたが、Wordが開いていませんでした。" & _
            "Wordを開いたままにして、もう一度お試しください。" & _
            "(詳細: " & desc & ")"
    ElseIf errNum = 429 Then
        DescribeComError = "この環境ではWord連携(COM)が利用できません。" & _
            "Mac版ExcelやCOM未対応環境の可能性があります。Windows版Excel+Wordでお試しください。" & _
            "(詳細: " & desc & ")"
    ElseIf errNum = 462 Then
        ' 462 は「相手のCOMサーバが居ない/応答しない」。実機では
        ' Wordが未インストール、セキュリティ製品にプロセス起動を止められて
        ' いる、起動したWordが即座に落ちている、のいずれかであることが多い。
        ' 生の文言は原因を何も示さないので、確かめる先を書く。
        DescribeComError = "Wordを操作できませんでした。" & _
            "Wordがインストールされているか、" & _
            "Wordを手で起動できるか(スタートメニューから)をご確認ください。" & _
            "手で起動できるのにここで失敗する場合、" & _
            "セキュリティ製品が他アプリからのWord操作を止めている可能性があります。" & _
            "(詳細: " & desc & ")"
    Else
        DescribeComError = desc
    End If
End Function

' ページ末尾に付く改ページ文字(Chr(12))と、その直前の改行を落とす。
' 本文の1文字を守るために End を1つ伸ばした副作用の後始末(レビュー M-14)。
Private Function TrimPageBreak(ByVal s As String) As String
    Dim t As String: t = s
    Do While Len(t) > 0
        Dim lastCh As String: lastCh = Right$(t, 1)
        If lastCh = Chr$(12) Or lastCh = vbCr Or lastCh = vbLf Then
            t = Left$(t, Len(t) - 1)
        Else
            Exit Do
        End If
    Loop
    TrimPageBreak = t
End Function
