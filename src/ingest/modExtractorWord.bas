Attribute VB_Name = "modExtractorWord"
Option Explicit

' 文書の開き方。1=非表示+パスワード無効化 / 2=非表示 / 3=表示。
' 実機で通った開き方をセッション中おぼえておく(0=未確定)。
Private mPreferredMode As Long
Private Const OPEN_MODE_MAX As Long = 3

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
' 2026-07-29(実機事故): 取込が .doc / .pdf で全滅した(err#462
' 「リモート サーバーがないか、使用できる状態ではありません」)。
'
' 原因は 2026-07-28 の修正で Word を Visible=False に戻したこと。
' あの Visible=True は「診断コードの消し忘れ」ではなく、
' 【非表示のWordが起動直後に落ちる環境で、実際に効いていた回避策】だった。
' コメントにも「実機err#462はローカルコピーでも再現し」と書いてあったのに、
' レビューの「診断コードの残存」という指摘をそのまま適用してしまった。
' 動いているものを、理由を確かめずに外した。
'
' 直し方: どちらか一方を選ぶのをやめる。
' 開き方を「静かな順」に並べて、通るまで降りていく。
'   1. 非表示 + パスワード無効化   … いちばん行儀が良い
'   2. 非表示(パスワード引数なし)  … 引数が原因のときここで通る
'   3. 表示                        … 実機で動いていた形。最後の砦
' 一度成功した開き方はセッション中おぼえておき、次からそこから始める
' (毎回1から試すと、100件のフォルダ同期で失敗待ちの時間を100回払う)。
' どの開き方で通ったかは usage_log に残す。次はこれで推測ではなく事実から
' 判断できる。
Public Function Extract(ByVal path As String, ByVal maxPages As Long, _
                        ByRef pages() As ExtractedPage, ByRef truncated As Boolean, _
                        ByRef errDetail As String) As Boolean
    Dim startMode As Long
    startMode = mPreferredMode
    If startMode < 1 Or startMode > OPEN_MODE_MAX Then startMode = 1

    Dim mode As Long
    For mode = startMode To OPEN_MODE_MAX
        Dim attempt As Long
        For attempt = 1 To 2
            If TryExtractOnce(path, maxPages, pages, truncated, errDetail, mode) Then
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
    Next mode

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
        Case Else: OpenModeName = "表示"
    End Select
End Function

Private Function TryExtractOnce(ByVal path As String, ByVal maxPages As Long, _
                                ByRef pages() As ExtractedPage, ByRef truncated As Boolean, _
                                ByRef errDetail As String, ByVal openMode As Long) As Boolean
    truncated = False

    Dim word As Object
    Dim doc As Object

    On Error GoTo Failed
    Set word = CreateObject("Word.Application")
    ' 開き方は openMode で切り替える(Extract のコメント参照)。
    ' 3 = 表示。実機で err#462 を回避できていた形なので最後の砦にする。
    Dim showWord As Boolean: showWord = (openMode >= 3)
    word.Visible = showWord
    word.DisplayAlerts = 0   ' wdAlertsNone

    ' 取り込む文書のマクロを走らせない(レビュー H-11)。
    ' 3 = msoAutomationSecurityForceDisable。出所の分からないファイルを
    ' 取り込むのは日常操作なので、そこが任意コード実行の経路になっていては
    ' いけない。この設定自体は err#462 の原因ではないので全モードで掛ける。
    ' このWordインスタンスは最後に Quit するので元へ戻す必要はない。
    On Error Resume Next
    word.AutomationSecurity = 3
    Err.Clear
    On Error GoTo Failed

    ' ConfirmConversions:=False でPDFリフロー確認ダイアログを抑止する。
    ' PasswordDocument にダミーを渡すのは、暗号化文書に当たったときに
    ' パスワード入力ダイアログでフリーズさせないため(レビュー M-11)。
    ' ただしこの引数自体が環境によっては Open を失敗させうるので、
    ' openMode=1 のときだけ付ける(駄目なら 2 以降で外して試す)。
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

    Dim pageCount As Long
    pageCount = doc.ComputeStatistics(2)   ' wdStatisticPages
    If pageCount < 1 Then pageCount = 1

    Dim loopCount As Long: loopCount = pageCount
    If loopCount > maxPages Then
        loopCount = maxPages
        truncated = True
    End If

    Dim tmp() As ExtractedPage: ReDim tmp(0 To loopCount - 1)
    Dim i As Long
    For i = 1 To loopCount
        tmp(i - 1).page = i
        tmp(i - 1).Text = ExtractPageText(doc, i, pageCount)
    Next i

    doc.Close 0    ' wdDoNotSaveChanges
    word.Quit 0
    Set doc = Nothing
    Set word = Nothing

    pages = tmp
    TryExtractOnce = True
    Exit Function

Failed:
    errDetail = DescribeComError(Err.Number, Err.Description)
    If Not doc Is Nothing Then
        On Error Resume Next
        doc.Close 0
        On Error GoTo 0
    End If
    If Not word Is Nothing Then
        On Error Resume Next
        word.Quit 0
        On Error GoTo 0
    End If
    Set doc = Nothing
    Set word = Nothing
    TryExtractOnce = False
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
Private Function DescribeComError(ByVal errNum As Long, ByVal desc As String) As String
    If errNum = 429 Then
        DescribeComError = "この環境ではWord連携(COM)が利用できません。" & _
            "Mac版ExcelやCOM未対応環境の可能性があります。Windows版Excel+Wordでお試しください。" & _
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
