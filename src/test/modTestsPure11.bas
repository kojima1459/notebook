Attribute VB_Name = "modTestsPure11"
Option Explicit

' ============================================================================
' modTestsPure11 - R14-3/R14-4(共有読みコピーの根治とOCR全量取込)の
'                  純ロジック回帰テスト
' ----------------------------------------------------------------------------
' なぜ新しいモジュールなのか:
'   modTestsPure9 は28,000字のWARN帯まで残り578字しかなく、真理表を1行足す
'   だけで超える。憲章§4-6「WARN帯のモジュールに機能を足さない。足す前に
'   分割を裁定する」に従い、新規分割先とした。入口は modTestsPure10.RunAll10
'   の末尾から呼ばれる Public Sub RunAll11()。導線はその1行だけで、消すと
'   ここのテストは「実行されないまま」全部PASSに見える。
'
' 固定する事実:
'   ・modExtractorPdf.IsUnreadableCopyReason / CopyFailMsgFor(R14-3b):
'     一時コピーの失敗理由のうち「元ファイルを読めていない」3種
'     (src_empty / size_mismatch / load_fail)だけを止める判定。ここが
'     ゆるいと、読めていない共有ファイルが元パスのままGhostscriptへ渡り
'     /undefinedfilename → 「描画0枚」だけが残る(実機第3報 RC3の本体)。
'     逆に厳しすぎると、単に一時フォルダが無いだけの端末で取込が止まる。
'   ・modLog.FriendlyFailMsg(R14-3b): その1文が汎用文言(Ghostscriptの話)に
'     潰されず、診断用の [reason] は利用者へ見せないこと。
'   ・optOcrCore.ClassifyGsTextResult の emptyinput(R14-3c): 入力0バイトは
'     rc・ログより先に言い切る。0バイトのPDFがGSの「空のPS即終了」で
'     画像PDF扱いになり、OCRで描画0枚に化ける経路を塞ぐ。
'   ・optOcrCore.BatchCountFor / BatchBoundsFor(R14-4a): 20ページずつの
'     バッチ境界。ここがズレるとページが飛ぶ(取り込めたつもりで欠ける)か、
'     同じページを二度OCRする(コストが倍)。
'   ・optOcrCore.OcrPageBanner(R14-4a): 進捗の文面とETAの出し方。
'   ・optOcrCore.OcrCapMemoFor(R14-4c): 上限打ち切りの正直なメモ。
'     総ページ数が不明(0や上限以下)のときに数字をでっち上げないこと。
' ============================================================================

' ----------------------------------------------------------------------------
' コピー失敗理由の判定と文面(R14-3b)
' ----------------------------------------------------------------------------
Private Sub TestCopyFailReason()
    ' 元を読めていない3種は必ず止める。
    modTestRunner.Check "CopyReason_src_emptyは読めていない", _
        modExtractorPdf.IsUnreadableCopyReason("src_empty")
    modTestRunner.Check "CopyReason_size_mismatchは読めていない", _
        modExtractorPdf.IsUnreadableCopyReason("size_mismatch src=1234 dest=0")
    modTestRunner.Check "CopyReason_load_failは読めていない", _
        modExtractorPdf.IsUnreadableCopyReason("load_fail err#3004")

    ' 「読めなかった」わけではない失敗は止めない(原本での続行を妨げない)。
    modTestRunner.Check "CopyReason_no_tempは止めない", _
        (modExtractorPdf.IsUnreadableCopyReason("no_temp") = False)
    modTestRunner.Check "CopyReason_name_busyは止めない", _
        (modExtractorPdf.IsUnreadableCopyReason("name_busy") = False)
    modTestRunner.Check "CopyReason_空は止めない", _
        (modExtractorPdf.IsUnreadableCopyReason("") = False)

    ' 文面: 止める理由のときだけ出る。診断用の理由は角括弧で末尾に付く。
    Dim m As String: m = modExtractorPdf.CopyFailMsgFor("size_mismatch src=10 dest=0")
    modTestRunner.Check "CopyFailMsg_本文が入る", _
        (InStr(m, "共有のファイルを読み取れませんでした") > 0), "実際=" & m
    modTestRunner.Check "CopyFailMsg_次の一手が入る", _
        (InStr(m, "ファイル名を変えて") > 0), "実際=" & m
    modTestRunner.Check "CopyFailMsg_理由が角括弧で付く", _
        (InStr(m, " [size_mismatch") > 0), "実際=" & m
    modTestRunner.Check "CopyFailMsg_止めない理由では空", _
        (LenB(modExtractorPdf.CopyFailMsgFor("no_temp")) = 0)
End Sub

' ----------------------------------------------------------------------------
' 利用者に届く文面(R14-3b): FriendlyFailMsg が汎用文言に潰さないこと。
' ----------------------------------------------------------------------------
Private Sub TestSharedReadFriendlyMsg()
    ' modExtractor.ExtractFile が作る形(理由つき)。
    Dim detail As String
    detail = modExtractorPdf.CopyFailMsgFor("load_fail err#3004")
    Dim msg As String: msg = modLog.FriendlyFailMsg("E0302", detail, "pdf")

    modTestRunner.Check "SharedRead_案内文がそのまま届く", _
        (InStr(msg, "ファイル名を変えて") > 0), "実際=" & msg
    modTestRunner.Check "SharedRead_Ghostscriptの汎用文言に潰されない", _
        (InStr(msg, "Ghostscript") = 0), "実際=" & msg
    modTestRunner.Check "SharedRead_診断用の理由は見せない", _
        (InStr(msg, "err#3004") = 0), "実際=" & msg
    modTestRunner.Check "SharedRead_コードは付く", _
        (InStr(msg, "(コード: E0302)") > 0), "実際=" & msg

    ' GS/Word/Acrobatの3者併記の中に埋まっていても拾える(optGsTxtのempty_input経路)。
    Dim mixed As String
    mixed = "GS: #ERR:E0302:empty_input:" & modLog.SharedReadFailMsg() & _
        " / Word: 型が一致しません。 / Acrobat: (応答なし)"
    Dim msg2 As String: msg2 = modLog.FriendlyFailMsg("E0302", mixed, "pdf")
    modTestRunner.Check "SharedRead_3者併記でも案内文を拾う", _
        (InStr(msg2, "ファイル名を変えて") > 0), "実際=" & msg2
    modTestRunner.Check "SharedRead_3者併記で後ろの技術情報を切る", _
        (InStr(msg2, "Word:") = 0), "実際=" & msg2
End Sub

' ----------------------------------------------------------------------------
' 空入力の分類(R14-3c): 入力0バイトは rc・ログより先に言い切る。
' ----------------------------------------------------------------------------
Private Sub TestClassifyEmptyInput()
    ' 0バイト入力は、従来なら image(rc=0+ページ処理あり)に見える組合せでも
    ' emptyinput が勝つ。ここを image に倒すと、0バイトのコピーが「画像PDF」
    ' としてOCRへ回り、描画0枚だけが残って原因が消える(RC3)。
    modTestRunner.Check "Classify_0バイト入力はemptyinput(rc0+ページ処理あり)", _
        (optOcrCore.ClassifyGsTextResult(0, 0, True, 0) = "emptyinput"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(0, 0, True, 0)
    modTestRunner.Check "Classify_0バイト入力はemptyinput(rc非0)", _
        (optOcrCore.ClassifyGsTextResult(1, 0, False, 0) = "emptyinput"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(1, 0, False, 0)
    modTestRunner.Check "Classify_0バイト入力はemptyinput(rc不明)", _
        (optOcrCore.ClassifyGsTextResult(-1, 0, True, 0) = "emptyinput"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(-1, 0, True, 0)

    ' 本文が取れているなら入力サイズは関係ない(okが最優先のまま)。
    modTestRunner.Check "Classify_本文ありは0バイトでもok", _
        (optOcrCore.ClassifyGsTextResult(0, 1234, True, 0) = "ok"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(0, 1234, True, 0)

    ' サイズ不明(-1)と実サイズありは従来の3値のまま(既存の真理表を壊さない)。
    modTestRunner.Check "Classify_サイズ不明は従来どおりimage", _
        (optOcrCore.ClassifyGsTextResult(0, 0, True, -1) = "image"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(0, 0, True, -1)
    modTestRunner.Check "Classify_サイズありは従来どおりimage", _
        (optOcrCore.ClassifyGsTextResult(0, 0, True, 51200) = "image"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(0, 0, True, 51200)
    modTestRunner.Check "Classify_サイズありrc非0は従来どおりgsfail", _
        (optOcrCore.ClassifyGsTextResult(1, 0, True, 51200) = "gsfail"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(1, 0, True, 51200)
    modTestRunner.Check "Classify_サイズありrc不明は従来どおりflagdelay", _
        (optOcrCore.ClassifyGsTextResult(-1, 0, True, 51200) = "flagdelay"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(-1, 0, True, 51200)
    modTestRunner.Check "Classify_引数省略時は従来どおりimage", _
        (optOcrCore.ClassifyGsTextResult(0, 0, True) = "image"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(0, 0, True)
End Sub

' ----------------------------------------------------------------------------
' バッチ境界(R14-4a): ページが飛ばない・重ならない・はみ出さない。
' ----------------------------------------------------------------------------
Private Sub TestBatchBounds()
    ' 既定形(上限100 → 描画は101ページ → 20ページずつで6回。最後は1ページ)。
    modTestRunner.Check "BatchCount_101を20ずつは6回", _
        (optOcrCore.BatchCountFor(101, 20) = 6), _
        "実際=" & optOcrCore.BatchCountFor(101, 20)
    modTestRunner.Check "BatchBounds_101の1回目は1-20", _
        (optOcrCore.BatchBoundsFor(101, 20, 1) = "1|20"), _
        "実際=" & optOcrCore.BatchBoundsFor(101, 20, 1)
    modTestRunner.Check "BatchBounds_101の2回目は21-40", _
        (optOcrCore.BatchBoundsFor(101, 20, 2) = "21|40"), _
        "実際=" & optOcrCore.BatchBoundsFor(101, 20, 2)
    modTestRunner.Check "BatchBounds_101の6回目は101-101", _
        (optOcrCore.BatchBoundsFor(101, 20, 6) = "101|101"), _
        "実際=" & optOcrCore.BatchBoundsFor(101, 20, 6)
    modTestRunner.Check "BatchBounds_101の7回目は無い", _
        (LenB(optOcrCore.BatchBoundsFor(101, 20, 7)) = 0), _
        "実際=" & optOcrCore.BatchBoundsFor(101, 20, 7)

    ' 割り切れる形(20ページ丁度=1回)と、1バッチに満たない形。
    modTestRunner.Check "BatchCount_20を20ずつは1回", _
        (optOcrCore.BatchCountFor(20, 20) = 1), _
        "実際=" & optOcrCore.BatchCountFor(20, 20)
    modTestRunner.Check "BatchBounds_20の1回目は1-20", _
        (optOcrCore.BatchBoundsFor(20, 20, 1) = "1|20"), _
        "実際=" & optOcrCore.BatchBoundsFor(20, 20, 1)
    modTestRunner.Check "BatchBounds_3ページは1-3で終わる", _
        (optOcrCore.BatchBoundsFor(3, 20, 1) = "1|3"), _
        "実際=" & optOcrCore.BatchBoundsFor(3, 20, 1)

    ' 壊れた値でも暴走しない(0回・空文字)。
    modTestRunner.Check "BatchCount_0ページは0回", _
        (optOcrCore.BatchCountFor(0, 20) = 0)
    modTestRunner.Check "BatchCount_バッチ幅0は0回", _
        (optOcrCore.BatchCountFor(101, 0) = 0)
    modTestRunner.Check "BatchBounds_バッチ番号0は空", _
        (LenB(optOcrCore.BatchBoundsFor(101, 20, 0)) = 0)
    modTestRunner.Check "BatchBounds_バッチ幅0は空", _
        (LenB(optOcrCore.BatchBoundsFor(101, 0, 1)) = 0)

    ' 全バッチを連結すると1..totalが飛びも重なりも無く1回ずつ現れること。
    Dim total As Long: total = 101
    Dim size As Long: size = 20
    Dim expectNext As Long: expectNext = 1
    Dim contiguous As Boolean: contiguous = True
    Dim b As Long
    For b = 1 To optOcrCore.BatchCountFor(total, size)
        Dim parts() As String
        parts = Split(optOcrCore.BatchBoundsFor(total, size, b), "|")
        If UBound(parts) - LBound(parts) <> 1 Then
            contiguous = False
        Else
            If CLng(Val(parts(LBound(parts)))) <> expectNext Then contiguous = False
            expectNext = CLng(Val(parts(LBound(parts) + 1))) + 1
        End If
    Next b
    modTestRunner.Check "BatchBounds_全バッチで1..101が連続する", _
        (contiguous And expectNext = total + 1), "次に期待した頁=" & expectNext

    ' GSコマンドは先頭ページを省略すると従来と1文字も変わらない(ゴールデン維持)、
    ' 指定すると -dFirstPage がその値になる。
    Dim c1 As String
    c1 = optOcrCore.BuildGsCommand("C:\gs\gswin32c.exe", "C:\a.pdf", "C:\t\page_%03d.jpg", 150, 21)
    modTestRunner.Check "BuildGsCommand_省略時は先頭1のまま", _
        (InStr(c1, "-dFirstPage=1 -dLastPage=21") > 0), "実際=" & c1
    Dim c2 As String
    c2 = optOcrCore.BuildGsCommand("C:\gs\gswin32c.exe", "C:\a.pdf", "C:\t\page_%03d.jpg", 150, 40, 21)
    modTestRunner.Check "BuildGsCommand_バッチ指定が効く", _
        (InStr(c2, "-dFirstPage=21 -dLastPage=40") > 0), "実際=" & c2
    Dim c3 As String
    c3 = optOcrCore.BuildGsCommand("C:\gs\gswin32c.exe", "C:\a.pdf", "C:\t\page_%03d.jpg", 150, 5, 9)
    modTestRunner.Check "BuildGsCommand_逆転した範囲は1ページへ丸める", _
        (InStr(c3, "-dFirstPage=9 -dLastPage=9") > 0), "実際=" & c3
End Sub

' ----------------------------------------------------------------------------
' 進捗バナー(R14-4a)と上限メモ(R14-4c)。
' ----------------------------------------------------------------------------
Private Sub TestOcrBannerAndCapMemo()
    ' 実測が無いうちは残り時間を出さない(跳ね回る表示を作らない)。
    Dim b0 As String: b0 = optOcrCore.OcrPageBanner(1, 100, 1, 5, 0#)
    modTestRunner.Check "OcrBanner_頁とバッチが出る", _
        (InStr(b0, "1/100頁") > 0 And InStr(b0, "(バッチ 1/5)") > 0), "実際=" & b0
    modTestRunner.Check "OcrBanner_実測前は残り時間を出さない", _
        (InStr(b0, "残り") = 0), "実際=" & b0

    ' 実測が入ったら残り時間を出す(残り98頁 × 1000ms = 98秒)。
    Dim b1 As String: b1 = optOcrCore.OcrPageBanner(2, 100, 1, 5, 1000#)
    modTestRunner.Check "OcrBanner_実測後は残り時間を出す", _
        (InStr(b1, "残り約98秒") > 0), "実際=" & b1

    ' 最終頁では残りを出さない。分母が現在頁より小さい壊れた値は寄せる。
    Dim b2 As String: b2 = optOcrCore.OcrPageBanner(100, 100, 5, 5, 1000#)
    modTestRunner.Check "OcrBanner_最終頁は残りを出さない", _
        (InStr(b2, "残り") = 0), "実際=" & b2
    Dim b3 As String: b3 = optOcrCore.OcrPageBanner(7, 3, 1, 1, 0#)
    modTestRunner.Check "OcrBanner_壊れた分母は現在頁へ寄せる", _
        (InStr(b3, "7/7頁") > 0), "実際=" & b3

    ' 上限メモ: 打ち切っていなければ何も言わない。
    modTestRunner.Check "OcrCapMemo_打ち切りなしは空", _
        (LenB(optOcrCore.OcrCapMemoFor(False, 30, 0)) = 0)

    ' 総ページ数が分からないとき(0や上限以下)は数字をでっち上げない。
    Dim m1 As String: m1 = optOcrCore.OcrCapMemoFor(True, 100, 0)
    modTestRunner.Check "OcrCapMemo_総頁不明は先頭Nのみと言う", _
        (InStr(m1, "上限100ページのため、先頭100ページのみ取り込みました") > 0), "実際=" & m1
    modTestRunner.Check "OcrCapMemo_設定名を案内する", _
        (InStr(m1, "vision_pdf_max_pages") > 0), "実際=" & m1
    modTestRunner.Check "OcrCapMemo_再開できるとは言わない", _
        (InStr(m1, "再開") = 0), "実際=" & m1
    modTestRunner.Check "OcrCapMemo_総頁が上限以下なら不明扱い", _
        (optOcrCore.OcrCapMemoFor(True, 100, 100) = m1), _
        "実際=" & optOcrCore.OcrCapMemoFor(True, 100, 100)

    ' 総ページ数が分かっているときは x/y で言う。
    Dim m2 As String: m2 = optOcrCore.OcrCapMemoFor(True, 100, 240)
    modTestRunner.Check "OcrCapMemo_総頁既知はy中xと言う", _
        (InStr(m2, "240ページ中100ページのみ取り込みました") > 0), "実際=" & m2
End Sub

Public Sub RunAll11()
    On Error GoTo CopyReasonFail
    TestCopyFailReason
NextSharedMsg:
    On Error GoTo SharedMsgFail
    TestSharedReadFriendlyMsg
NextEmptyInput:
    On Error GoTo EmptyInputFail
    TestClassifyEmptyInput
NextBatch:
    On Error GoTo BatchFail
    TestBatchBounds
NextBanner:
    On Error GoTo BannerFail
    TestOcrBannerAndCapMemo
NextDone11:
    On Error GoTo 0
    Exit Sub

CopyReasonFail:
    modTestRunner.Check "TestCopyFailReason(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextSharedMsg
SharedMsgFail:
    modTestRunner.Check "TestSharedReadFriendlyMsg(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextEmptyInput
EmptyInputFail:
    modTestRunner.Check "TestClassifyEmptyInput(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBatch
BatchFail:
    modTestRunner.Check "TestBatchBounds(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBanner
BannerFail:
    modTestRunner.Check "TestOcrBannerAndCapMemo(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone11
End Sub
