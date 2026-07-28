Attribute VB_Name = "modExtractorWord"
Option Explicit

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
Public Function Extract(ByVal path As String, ByVal maxPages As Long, _
                        ByRef pages() As ExtractedPage, ByRef truncated As Boolean, _
                        ByRef errDetail As String) As Boolean
    Dim attempt As Long
    For attempt = 1 To 2
        If TryExtractOnce(path, maxPages, pages, truncated, errDetail) Then
            Extract = True
            Exit Function
        End If
        If attempt = 1 Then WaitBriefly 800
    Next attempt
    Extract = False
End Function

Private Function TryExtractOnce(ByVal path As String, ByVal maxPages As Long, _
                                ByRef pages() As ExtractedPage, ByRef truncated As Boolean, _
                                ByRef errDetail As String) As Boolean
    truncated = False

    Dim word As Object
    Dim doc As Object

    On Error GoTo Failed
    Set word = CreateObject("Word.Application")
    ' 2026-07-28(レビュー M-12): 2026-07-22の診断コード(Visible=True)が
    ' 本番ビルドに残っていた。PDF/docx を取り込むたびに Word のウィンドウが
    ' 開いては閉じ、利用者がそこを触ると COM エラーや WINWORD の残留を招く。
    ' 診断は終わったので非表示へ戻す。
    word.Visible = False
    word.DisplayAlerts = 0   ' wdAlertsNone

    ' 2026-07-28(レビュー H-11): 取り込む文書のマクロを走らせない。
    ' 3 = msoAutomationSecurityForceDisable。出所の分からないファイルを
    ' 取り込むのは日常操作なので、そこが任意コード実行の経路になっていては
    ' いけない。このWordインスタンスは最後に Quit するので元へ戻す必要はない。
    On Error Resume Next
    word.AutomationSecurity = 3
    Err.Clear
    On Error GoTo Failed

    ' ConfirmConversions:=False でPDFリフロー確認ダイアログを抑止する。
    ' PasswordDocument にダミーを渡すのは、暗号化文書に当たったときに
    ' パスワード入力ダイアログでフリーズさせないため(レビュー M-11)。
    ' 誤ったパスワードは即エラーになるので、E0302 として扱える。
    Set doc = word.Documents.Open( _
        FileName:=path, _
        ConfirmConversions:=False, _
        ReadOnly:=True, _
        AddToRecentFiles:=False, _
        PasswordDocument:="__mybookshelf_no_password__", _
        Visible:=False)

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
        startRange.End = endRange.Start - 1
    Else
        startRange.End = doc.Content.End
    End If
    ExtractPageText = startRange.Text
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
