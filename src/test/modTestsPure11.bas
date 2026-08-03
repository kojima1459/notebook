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
'   ・modExtractorPdf.IsUnreadableCopyReason / CopyReasonKind / CopyFailMsgFor
'     (R14-3b / R14-F3・F11): 一時コピーの失敗のうち「中身を手にできて
'     いない」6種(src_empty / size_mismatch / load_fail / locked / too_big /
'     name_busy)を止める判定と、種類ごとの文面。ここがゆるいと、読めていない
'     共有ファイルが元パスのままGhostscriptへ渡り /undefinedfilename →
'     「描画0枚」だけが残る(実機第3報 RC3の本体)。逆に厳しすぎると、単に
'     一時フォルダが無いだけの端末で取込が止まる。
'   ・modLog.FriendlyFailMsg / CopyFailMsgOf(R14-3b / R14-F3): 種類別の1文が
'     汎用文言(Ghostscriptの話)に潰されず、診断用の [reason] は見せないこと。
'   ・optOcrCore.ClassifyGsTextResult の emptyinput(R14-3c): 入力0バイトは
'     rc・ログより先に言い切る。0バイトのPDFがGSの「空のPS即終了」で
'     画像PDF扱いになり、OCRで描画0枚に化ける経路を塞ぐ。
'   ・optOcrCore.BatchCountFor / BatchBoundsFor(R14-4a): 20ページずつの
'     バッチ境界。ここがズレるとページが飛ぶ(取り込めたつもりで欠ける)か、
'     同じページを二度OCRする(コストが倍)。
'   ・optOcrCore.OcrPageBanner(R14-4a / R14-F9): 進捗の文面とETAの出し方。
'     総頁が確定するまで分母もバッチ総数もETAも出さないこと(嘘の分母禁止)。
'   ・optOcrCore.OcrCapMemoFor / OcrAbortMemoFor(R14-4c / R14-F2・F7):
'     上限は【設定の値】、取り込めた頁数は別、読めなかった頁数も別。途中で
'     中断したときは上限の話をせず「もう一度取り込むと再試行」と言うこと。
'   ・modMode.ShouldEmitInsight(R14-1b): 「解決した」で部内へ発信してよい
'     回答かの真理表(実機第3報 RC2)。
'
' 2026-08-03(R14-8): 本モジュールも28,000字のWARN帯へ入ったため、R14-8
'   (入念モードの本格強化+回答可読性)のテストは modTestsPure12 へ分割した。
'   末尾の RunAll11 が RunAll12 を呼ぶ。
' ============================================================================

' ----------------------------------------------------------------------------
' コピー失敗理由の判定と文面(R14-3b)
' ----------------------------------------------------------------------------
Private Sub TestCopyFailReason()
    ' コピーの中身を手にできていない理由は必ず止める(R14-F3/F11で6種)。
    modTestRunner.Check "CopyReason_src_emptyは読めていない", _
        modExtractorPdf.IsUnreadableCopyReason("src_empty")
    modTestRunner.Check "CopyReason_size_mismatchは読めていない", _
        modExtractorPdf.IsUnreadableCopyReason("size_mismatch src=1234 dest=0")
    modTestRunner.Check "CopyReason_load_failは読めていない", _
        modExtractorPdf.IsUnreadableCopyReason("load_fail err#3004")
    modTestRunner.Check "CopyReason_lockedは読めていない", _
        modExtractorPdf.IsUnreadableCopyReason("locked err#70")
    modTestRunner.Check "CopyReason_too_bigは読めていない", _
        modExtractorPdf.IsUnreadableCopyReason("too_big err#7")
    ' R14-F11: 一時名が枯渇したときもコピーは1つも作れていない。原本のまま
    ' 進めるとNFD分解名がそのままGhostscriptへ渡る(RC3と同じ穴)ので止める。
    modTestRunner.Check "CopyReason_name_busyも止める", _
        modExtractorPdf.IsUnreadableCopyReason("name_busy")

    ' 一時フォルダがそもそも無いだけなら止めない(原本での続行を妨げない)。
    modTestRunner.Check "CopyReason_no_tempは止めない", _
        (modExtractorPdf.IsUnreadableCopyReason("no_temp") = False)
    modTestRunner.Check "CopyReason_空は止めない", _
        (modExtractorPdf.IsUnreadableCopyReason("") = False)

    ' 理由 → 文言の種類(R14-F3)。ここがズレると全部が特殊文字の話になる。
    modTestRunner.Check "CopyKind_src_empty", _
        (modExtractorPdf.CopyReasonKind("src_empty") = "src_empty")
    modTestRunner.Check "CopyKind_locked", _
        (modExtractorPdf.CopyReasonKind("locked err#3002") = "locked")
    modTestRunner.Check "CopyKind_too_big", _
        (modExtractorPdf.CopyReasonKind("too_big err#3004") = "too_big")
    modTestRunner.Check "CopyKind_name_busy", _
        (modExtractorPdf.CopyReasonKind("name_busy") = "name_busy")
    modTestRunner.Check "CopyKind_size_mismatchは既定(特殊文字の文面)", _
        (LenB(modExtractorPdf.CopyReasonKind("size_mismatch src=1 dest=0")) = 0)
    modTestRunner.Check "CopyKind_load_failは既定(特殊文字の文面)", _
        (LenB(modExtractorPdf.CopyReasonKind("load_fail err#5")) = 0)

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

    ' R14-F3/F11: 種類ごとに「その場で打てる一手」が違う。名前を変えろと
    ' 言ってよいのは、実体に届かなかった系(size_mismatch / load_fail)だけ。
    Dim mEmpty As String: mEmpty = modExtractorPdf.CopyFailMsgFor("src_empty")
    modTestRunner.Check "CopyFailMsg_空ファイルは0バイトと言う", _
        (InStr(mEmpty, "空(0バイト)") > 0), "実際=" & mEmpty
    modTestRunner.Check "CopyFailMsg_空ファイルに名前変更を求めない", _
        (InStr(mEmpty, "ファイル名を変えて") = 0), "実際=" & mEmpty

    Dim mLock As String: mLock = modExtractorPdf.CopyFailMsgFor("locked err#70")
    modTestRunner.Check "CopyFailMsg_ロックは閉じてからと言う", _
        (InStr(mLock, "他のアプリで開かれている") > 0 And _
         InStr(mLock, "閉じてから") > 0), "実際=" & mLock

    Dim mBig As String: mBig = modExtractorPdf.CopyFailMsgFor("too_big err#7")
    modTestRunner.Check "CopyFailMsg_大きすぎるはそう言う", _
        (InStr(mBig, "大きすぎて") > 0), "実際=" & mBig

    Dim msgBusy As String: msgBusy = modExtractorPdf.CopyFailMsgFor("name_busy")
    modTestRunner.Check "CopyFailMsg_一時名の枯渇は待つよう案内", _
        (InStr(msgBusy, "一時ファイルが混み合っています") > 0 And _
         InStr(msgBusy, "しばらくして") > 0), "実際=" & msgBusy
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

    ' R14-F3/F11: 種類別の文言も同じように生き残ること。1つでも
    ' ActionableHint の目印から漏れると、その種類だけE0302の汎用文言
    ' (Ghostscriptの話)に潰されて次の一手が消える。
    Dim kinds As Variant
    kinds = Array("src_empty", "locked err#70", "too_big err#7", "name_busy")
    Dim heads As Variant
    heads = Array("空(0バイト)", "他のアプリで開かれている", "大きすぎて", _
                  "一時ファイルが混み合っています")
    Dim i As Long
    For i = LBound(kinds) To UBound(kinds)
        Dim d As String: d = modExtractorPdf.CopyFailMsgFor(CStr(kinds(i)))
        Dim mk As String: mk = modLog.FriendlyFailMsg("E0302", d, "pdf")
        modTestRunner.Check "CopyFailMsg_" & CStr(kinds(i)) & "_案内が届く", _
            (InStr(mk, CStr(heads(i))) > 0), "実際=" & mk
        modTestRunner.Check "CopyFailMsg_" & CStr(kinds(i)) & "_汎用文言に潰されない", _
            (InStr(mk, "Ghostscript") = 0), "実際=" & mk
        modTestRunner.Check "CopyFailMsg_" & CStr(kinds(i)) & "_診断用の理由は見せない", _
            (InStr(mk, " [") = 0), "実際=" & mk
    Next i
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
    ' R14-F9: 総頁が確定してからだけ分母を出す。確定前に上限を分母として
    ' 出していたため、44頁のPDFで「1/100頁」という嘘が最後まで残っていた。
    Dim u0 As String: u0 = optOcrCore.OcrPageBanner(3, 0, 1, 0, 0#)
    modTestRunner.Check "OcrBanner_総頁未確定は頁目だけを言う", _
        (InStr(u0, "3頁目") > 0), "実際=" & u0
    modTestRunner.Check "OcrBanner_総頁未確定は分母を出さない", _
        (InStr(u0, "/") = 0), "実際=" & u0
    modTestRunner.Check "OcrBanner_総頁未確定はバッチ総数を出さない", _
        (InStr(u0, "(バッチ 1)") > 0), "実際=" & u0
    modTestRunner.Check "OcrBanner_総頁未確定はETAを出さない", _
        (InStr(optOcrCore.OcrPageBanner(3, 0, 1, 0, 1000#), "残り") = 0), _
        "実際=" & optOcrCore.OcrPageBanner(3, 0, 1, 0, 1000#)

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

    ' 最終頁では残りを出さない。分母が現在頁より小さい壊れた値は未確定扱い。
    Dim b2 As String: b2 = optOcrCore.OcrPageBanner(100, 100, 5, 5, 1000#)
    modTestRunner.Check "OcrBanner_最終頁は残りを出さない", _
        (InStr(b2, "残り") = 0), "実際=" & b2
    Dim b3 As String: b3 = optOcrCore.OcrPageBanner(7, 3, 1, 1, 0#)
    modTestRunner.Check "OcrBanner_壊れた分母は未確定として扱う", _
        (InStr(b3, "7頁目") > 0 And InStr(b3, "/") = 0), "実際=" & b3

    ' 上限メモ: 打ち切っていなければ何も言わない。
    modTestRunner.Check "OcrCapMemo_打ち切りなしは空", _
        (LenB(optOcrCore.OcrCapMemoFor(False, 30, 100)) = 0)

    ' R14-F7: 上限は【設定の値】。取り込めた頁数を上限として言わない。
    Dim m1 As String: m1 = optOcrCore.OcrCapMemoFor(True, 100, 100)
    modTestRunner.Check "OcrCapMemo_設定上限と取込頁数を言う", _
        (InStr(m1, "設定上限100ページのうち先頭100ページを取り込みました") > 0), "実際=" & m1
    modTestRunner.Check "OcrCapMemo_設定名を案内する", _
        (InStr(m1, "vision_pdf_max_pages") > 0), "実際=" & m1
    modTestRunner.Check "OcrCapMemo_再開できるとは言わない", _
        (InStr(m1, "再開") = 0), "実際=" & m1
    modTestRunner.Check "OcrCapMemo_全部読めたなら失敗頁は言わない", _
        (InStr(m1, "読み取れませんでした") = 0), "実際=" & m1

    ' 頁OCRの失敗で kept が上限に届かなかったぶんは、必ず数字で言う。
    Dim m2 As String: m2 = optOcrCore.OcrCapMemoFor(True, 97, 100)
    modTestRunner.Check "OcrCapMemo_上限は設定値のまま", _
        (InStr(m2, "設定上限100ページのうち先頭97ページ") > 0), "実際=" & m2
    modTestRunner.Check "OcrCapMemo_読めなかった頁数を言う", _
        (InStr(m2, "3ページは読み取れませんでした") > 0), "実際=" & m2

    ' 上限が不明・壊れた値(0や負)でも、取り込めた数より小さい嘘は出さない。
    Dim m3 As String: m3 = optOcrCore.OcrCapMemoFor(True, 20, 0)
    modTestRunner.Check "OcrCapMemo_上限不明は取込数で言い切る", _
        (InStr(m3, "設定上限20ページのうち先頭20ページ") > 0), "実際=" & m3

    ' R14-F2: 途中で中断したときは上限の話をしない(まだ先がある)。
    Dim a1 As String: a1 = optOcrCore.OcrAbortMemoFor(37)
    modTestRunner.Check "OcrAbortMemo_何頁で止まったかを言う", _
        (InStr(a1, "37頁で中断しました") > 0), "実際=" & a1
    modTestRunner.Check "OcrAbortMemo_理由の範囲を言う", _
        (InStr(a1, "変換エラーまたは時間切れ") > 0), "実際=" & a1
    modTestRunner.Check "OcrAbortMemo_再取込で再試行と言う", _
        (InStr(a1, "もう一度取り込むと再試行します") > 0), "実際=" & a1
    modTestRunner.Check "OcrAbortMemo_続きから再開とは言わない", _
        (InStr(a1, "続きから") = 0), "実際=" & a1
    modTestRunner.Check "OcrAbortMemo_上限の話をしない", _
        (InStr(a1, "上限") = 0), "実際=" & a1
    modTestRunner.Check "OcrAbortMemo_負の頁数でも壊れない", _
        (InStr(optOcrCore.OcrAbortMemoFor(-3), "0頁で中断") > 0), _
        "実際=" & optOcrCore.OcrAbortMemoFor(-3)
End Sub

' ----------------------------------------------------------------------------
' R14-1b: 「解決した」で部内へ発信してよい回答かの真理表(実機第3報 RC2)
'   緩むと (a)一般アシスタントの雑談が「人が確認した社内Q&A」として配信され
'   (b)残留出典へ無関係な感謝状が飛ぶ。配信も感謝状も取り消せない。
' ----------------------------------------------------------------------------
Private Sub TestShouldEmitInsight()
    ' 本棚の資料を根拠に答えたターンだけ発信してよい。
    modTestRunner.Check "発信判定_すぐ聞くで出典あり", _
        (modMode.ShouldEmitInsight("quick", 1) = True)
    modTestRunner.Check "発信判定_深掘りで出典あり", _
        (modMode.ShouldEmitInsight("deep", 3) = True)
    modTestRunner.Check "発信判定_入念で出典あり", _
        (modMode.ShouldEmitInsight("thorough", 12) = True)

    ' 一般モードは出典の有無に関わらず発信しない(残留 hits があっても)。
    modTestRunner.Check "発信判定_一般モードは出典0でも発信しない", _
        (modMode.ShouldEmitInsight("general", 0) = False)
    modTestRunner.Check "発信判定_一般モードは残留出典があっても発信しない", _
        (modMode.ShouldEmitInsight("general", 5) = False)

    ' 検索も生成もしなかったターン(空質問・0件・聞き返し)は mode="" になる。
    modTestRunner.Check "発信判定_モード空は発信しない", _
        (modMode.ShouldEmitInsight("", 4) = False)
    modTestRunner.Check "発信判定_出典0件は発信しない", _
        (modMode.ShouldEmitInsight("deep", 0) = False)
    modTestRunner.Check "発信判定_出典が負でも発信しない", _
        (modMode.ShouldEmitInsight("deep", -1) = False)

    ' 表記ゆれで判定が反転しないこと(モード名はui_state由来)。
    modTestRunner.Check "発信判定_大文字小文字を無視", _
        (modMode.ShouldEmitInsight("GENERAL", 3) = False)
    modTestRunner.Check "発信判定_前後空白を無視", _
        (modMode.ShouldEmitInsight("  general  ", 3) = False)
End Sub

' ----------------------------------------------------------------------------
' R14-2a: ツールバーの実際の右端(実機第3報 RC10「ヘッダー右ズレ」)。
'   Shapeを一切作らずToolbarSpec+FlowLeftの算数だけを走らせる。ToolbarSpecが
'   呼ぶ modPublish.CanPublish/modFeatures.FeatureEnabled は On Error Resume
'   Next配下のため、両モジュール未注入でもFalseへ倒れるだけで数値は決め打ち
'   できる(発行キー未設定・画像解析無効の環境と同じボタン構成になる)。
' ----------------------------------------------------------------------------
Private Sub TestToolbarContentRight()
    Dim L As Double: L = 10

    ' 帯を広げるほど右端は後退しない(段が減って詰まるだけ)。
    Dim rTable50 As Double: rTable50 = modKnowledgeBar.ToolbarContentRight(True, False, L, 50)
    Dim rTable200 As Double: rTable200 = modKnowledgeBar.ToolbarContentRight(True, False, L, 200)
    Dim rTable600 As Double: rTable600 = modKnowledgeBar.ToolbarContentRight(True, False, L, 600)
    Dim rTable900 As Double: rTable900 = modKnowledgeBar.ToolbarContentRight(True, False, L, 900)
    Dim rTable2000 As Double: rTable2000 = modKnowledgeBar.ToolbarContentRight(True, False, L, 2000)
    modTestRunner.Check "ツールバー右端_単調非減少(50→200)", (rTable200 >= rTable50), _
        "50=" & rTable50 & " 200=" & rTable200
    modTestRunner.Check "ツールバー右端_単調非減少(200→600)", (rTable600 >= rTable200), _
        "200=" & rTable200 & " 600=" & rTable600
    modTestRunner.Check "ツールバー右端_単調非減少(600→900)", (rTable900 >= rTable600), _
        "600=" & rTable600 & " 900=" & rTable900
    ' 全ボタンが1段に収まる幅を超えたら、帯を広げても右端は増えない
    ' (ボタンが伸びるわけではないため=頭打ち)。
    modTestRunner.Check "ツールバー右端_1段に収まったら頭打ち", (rTable2000 = rTable900), _
        "900=" & rTable900 & " 2000=" & rTable2000

    ' フォールバック境界(modKnowledge.DrawChromeがL+200未満でL+W-8へ戻す)。
    Dim rShared50 As Double: rShared50 = modKnowledgeBar.ToolbarContentRight(False, True, L, 50)
    modTestRunner.Check "ツールバー右端_狭い帯はフォールバック域", _
        (rShared50 > 0 And rShared50 < L + 200), "実際=" & rShared50
    Dim rShared900 As Double: rShared900 = modKnowledgeBar.ToolbarContentRight(False, True, L, 900)
    modTestRunner.Check "ツールバー右端_広い帯は通常域", _
        (rShared900 >= L + 200), "実際=" & rShared900

    ' ギャラリー(検索ボタンが増える分、一覧表より右端が広がる)。
    Dim rGallery2000 As Double: rGallery2000 = modKnowledgeBar.ToolbarContentRight(False, False, L, 2000)
    modTestRunner.Check "ツールバー右端_ギャラリーは一覧表より広い", _
        (rGallery2000 > rTable2000), "gallery=" & rGallery2000 & " table=" & rTable2000

    ' 帯幅0でも例外にならない(FlowLeft側の下限クランプが効く)。
    modTestRunner.Check "ツールバー右端_帯幅0でも例外にならない", _
        (modKnowledgeBar.ToolbarContentRight(True, False, L, 0) >= 0), ""
End Sub

' ----------------------------------------------------------------------------
' R14-5a/5c: 拡張子別チャンク設定のキー名(実機第3報 RC7「チャンク粗さ」)。
' ----------------------------------------------------------------------------
Private Sub TestChunkKeyOrder()
    modTestRunner.Check "チャンクキー_PDF拡張子で専用キー名", _
        (modShelf.ChunkKeyOrder("chunk_target_chars", "pdf") = "chunk_target_chars_pdf")
    modTestRunner.Check "チャンクキー_docx拡張子で専用キー名", _
        (modShelf.ChunkKeyOrder("chunk_max_chars", "docx") = "chunk_max_chars_docx")
    ' modUtil.ExtOfは常に小文字だが、ChunkKeyOrder自身も念のため小文字化する
    ' (呼び出し元がExtOfを経由し忘れても壊れない二重の安全)。
    modTestRunner.Check "チャンクキー_大文字拡張子も小文字化", _
        (modShelf.ChunkKeyOrder("chunk_overlap_chars", "PDF") = "chunk_overlap_chars_pdf")
    modTestRunner.Check "チャンクキー_前後空白を無視", _
        (modShelf.ChunkKeyOrder("chunk_max_chars", "  pdf  ") = "chunk_max_chars_pdf")

    ' modUtil.ExtOf(実際にIngestFileが渡す形)と組み合わせても一致すること。
    modTestRunner.Check "チャンクキー_ExtOfの戻り値と組み合う", _
        (modShelf.ChunkKeyOrder("chunk_target_chars", modUtil.ExtOf("C:\書類\規約.PDF")) _
         = "chunk_target_chars_pdf")
    modTestRunner.Check "チャンクキー_拡張子なしパスは末尾アンダースコアのみ", _
        (modShelf.ChunkKeyOrder("chunk_target_chars", modUtil.ExtOf("C:\書類\readme")) _
         = "chunk_target_chars_")
End Sub

' ----------------------------------------------------------------------------
' R14-7a: 質問例オンデマンド生成の応答パーサ(modRagParse.ParseQuestionLines)。
' ----------------------------------------------------------------------------
Private Sub TestParseQuestionLines()
    Dim r1 As String
    r1 = modRagParse.ParseQuestionLines( _
        "休業補償の対象は?" & vbLf & "免責期間は何日?" & vbLf & "更新手続きの締切は?", 5)
    modTestRunner.Check "質問パース_3行を3件へ", (UBound(Split(r1, "|")) = 2), "実際=" & r1
    modTestRunner.Check "質問パース_1件目がそのまま入る", _
        (Split(r1, "|")(0) = "休業補償の対象は?"), "実際=" & r1

    ' 番号・箇条書き記号は剥がす。空行は捨てる。
    Dim r2 As String
    r2 = modRagParse.ParseQuestionLines( _
        "1. 保険金の請求方法は?" & vbLf & vbLf & "2) 免責金額はいくら?" & vbLf & _
        ChrW(&H30FB) & "解約の手続きは?" & vbLf & "- 更新は自動?", 5)
    modTestRunner.Check "質問パース_番号を剥がす(ピリオド)", _
        (InStr(r2, "保険金の請求方法は?") > 0 And InStr(r2, "1.") = 0), "実際=" & r2
    modTestRunner.Check "質問パース_番号を剥がす(括弧)", _
        (InStr(r2, "免責金額はいくら?") > 0 And InStr(r2, "2)") = 0), "実際=" & r2
    modTestRunner.Check "質問パース_箇条書き記号(・)を剥がす", _
        (InStr(r2, ChrW(&H30FB) & "解約") = 0 And InStr(r2, "解約の手続きは?") > 0), "実際=" & r2
    modTestRunner.Check "質問パース_箇条書き記号(-)を剥がす", _
        (InStr(r2, "- 更新") = 0 And InStr(r2, "更新は自動?") > 0), "実際=" & r2
    modTestRunner.Check "質問パース_空行は数えない", (UBound(Split(r2, "|")) = 3), "実際=" & r2

    ' 上限件数(5件目までで打ち切り)。
    Dim many As String
    many = "Q1" & vbLf & "Q2" & vbLf & "Q3" & vbLf & "Q4" & vbLf & "Q5" & vbLf & "Q6" & vbLf & "Q7"
    Dim r3 As String: r3 = modRagParse.ParseQuestionLines(many, 5)
    modTestRunner.Check "質問パース_maxNで打ち切る", (UBound(Split(r3, "|")) = 4), "実際=" & r3
    modTestRunner.Check "質問パース_6件目は含まれない", (InStr(r3, "Q6") = 0), "実際=" & r3

    ' CRLF/CRも同じ結果になること(LLM応答の改行コードは保証されない)。
    Dim r4 As String: r4 = modRagParse.ParseQuestionLines("A" & vbCrLf & "B" & vbCr & "C", 5)
    modTestRunner.Check "質問パース_CRLF/CRもLF同様に分割", (r4 = "A|B|C"), "実際=" & r4

    ' 空文字・maxN<1は空文字(壊れた応答/設定で例外を出さない)。
    modTestRunner.Check "質問パース_空応答は空文字", (LenB(modRagParse.ParseQuestionLines("", 5)) = 0)
    modTestRunner.Check "質問パース_maxN0は空文字", _
        (LenB(modRagParse.ParseQuestionLines("A" & vbLf & "B", 0)) = 0)
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
NextEmit:
    On Error GoTo EmitFail
    TestShouldEmitInsight
NextToolbarRight:
    On Error GoTo ToolbarRightFail
    TestToolbarContentRight
NextChunkKey:
    On Error GoTo ChunkKeyFail
    TestChunkKeyOrder
NextQParse:
    On Error GoTo QParseFail
    TestParseQuestionLines
NextPure12:
    ' R14-8(入念モードの本格強化+回答可読性)のテストは modTestsPure12 へ
    ' 分割した(本モジュールが28,000字のWARN帯へ入ったため。憲章§4-6)。
    ' この1行を消すと向こうのテストは「実行されないまま」全部PASSに見える。
    On Error GoTo Pure12Fail
    modTestsPure12.RunAll12
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
    Resume NextEmit
EmitFail:
    modTestRunner.Check "TestShouldEmitInsight(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextToolbarRight
ToolbarRightFail:
    modTestRunner.Check "TestToolbarContentRight(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextChunkKey
ChunkKeyFail:
    modTestRunner.Check "TestChunkKeyOrder(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextQParse
QParseFail:
    modTestRunner.Check "TestParseQuestionLines(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextPure12
Pure12Fail:
    modTestRunner.Check "modTestsPure12(モジュール全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone11
End Sub
