Attribute VB_Name = "modTestsPure14"
Option Explicit

' ============================================================================
' modTestsPure14 - R15-FixA(敵対的レビュー裁定 Fix-A)の純ロジック回帰テスト
' ----------------------------------------------------------------------------
' なぜ新設したか(憲章§4-6):
'   modTestsPure13 が25,138字で、ここの真理表とゴールデン文字列を足すと
'   30,000字上限まで残り11字になった。次のバグ修正どころか1行の追記も入らない
'   ので、R15波2で13を新設したときと同じ線で分割する。
'   入口は modTestsPure13.RunAll13 の末尾から呼ばれる RunAll14 の1本だけ。
'
' ここで固定するもの:
'   ・optOcrEta.ComposeOcrMemo(FA-4): 本棚カードのメモの組み立て契約。
'     「復元の冒頭文は前置きであって置換ではない」「頁欠け/中断の理由と
'     設定上限の説明は、両方あれば連結する」。従来は optOcrPage と
'     optVision.OcrCapMemo が別々に組み立て、あとから来たものが前のものを
'     丸ごと置換していたため、事実が片方ずつ黙って消えていた(A-H4)。
'   ・optOcrEta.OcrPartialMemoFor の分母(FA-4): 総頁が分かっているときは
'     それを使う。上限で切られた254頁の資料で「全200頁中3頁が…」と名乗ると、
'     残り54頁の存在がカードから消える。
'   ・optOcrEta.BatchWaitSec(FA-5iii): 1バッチの画像化待ちの上限。
'     資料あたりの残り予算をそのまま1バッチへ渡していたため、1バッチ目の
'     ハングが予算を全部食い潰し、残り12バッチには1秒も残らなかった。
'   ・optOcrEta.RenderWaitBanner(FA-5i): 画像化待ちの実況(経過秒つき)。
'     この区間は画面が完全に止まって見えるので、数字が動くことだけが
'     「止まっていない」証拠になる(憲章§3-2)。
'   ・optOcrEta.OcrConfirmAskFor(FA-1): 確認文は埋め込み時間にも正直に言う。
'
' R15-FixB(2026-08-04 レビュー裁定 Fix-B)で追加した分:
'   ・optOcrEta.ClampPageMs(FB-4): ui_state に永続化される1頁あたり実績の
'     丸め。異常値が1度混ざると、以後ずっと嘘のETAを出し続ける端末になる。
'   ・modUtilText.IsDeclineNote(FB-5): 「見送り」を失敗と分けて数えるための
'     唯一の判定。ここが外れると、自分で見送っただけの利用者に
'     「取り込めませんでした。状態をご確認ください」と不具合を疑わせる。
' ============================================================================

' ----------------------------------------------------------------------------
' FA-4(レビューA-H4/B-M): メモ合成の契約。
' ----------------------------------------------------------------------------
Private Sub TestComposeOcrMemo()
    Dim head As String: head = "前回の続きから再開しました。"
    Dim cap As String: cap = optOcrEta.OcrCapMemoFor(True, 200, 254)
    Dim part As String: part = optOcrEta.OcrPartialMemoFor(197, 3, 254)

    ' (1) 復元 + 上限打切り。冒頭は復元、そのあとに上限の説明。
    Dim m1 As String: m1 = optOcrEta.ComposeOcrMemo(head, "", cap)
    modTestRunner.Check "メモ合成_復元とcap_冒頭は復元の一文", _
        (Left$(m1, Len(head)) = head), "実際=" & m1
    modTestRunner.Check "メモ合成_復元とcap_上限の説明が消えない", _
        (InStr(m1, "設定上限254ページのうち先頭200ページ") > 0), "実際=" & m1

    ' (2) 復元のみで完了(頁欠けも打ち切りも無い)。冒頭文だけが残る。
    '     この形が truncated=False で起きる=従来はカードへ届く経路が
    '     1つも無く、「同じ資料が前より早く終わった」説明がどこにも無かった。
    Dim m2 As String: m2 = optOcrEta.ComposeOcrMemo(head, "", "")
    modTestRunner.Check "メモ合成_復元のみ完了は冒頭文だけ", (m2 = head), "実際=" & m2

    ' (3) 部分失敗 + cap + 復元 の3連結。3つの事実が全部残ること。
    Dim m3 As String: m3 = optOcrEta.ComposeOcrMemo(head, part, cap)
    modTestRunner.Check "メモ合成_3連結_復元が冒頭", _
        (Left$(m3, Len(head)) = head), "実際=" & m3
    modTestRunner.Check "メモ合成_3連結_頁欠けが残る", _
        (InStr(m3, "全254頁中3頁が読み取れませんでした") > 0), "実際=" & m3
    modTestRunner.Check "メモ合成_3連結_上限も残る", _
        (InStr(m3, "設定上限254ページ") > 0), "実際=" & m3
    modTestRunner.Check "メモ合成_3連結_順序は頁欠け→上限", _
        (InStr(m3, "全254頁中3頁") < InStr(m3, "設定上限254ページ")), "実際=" & m3
    modTestRunner.Check "メモ合成_3連結_文が地続きにならない(区切りがある)", _
        (InStr(m3, "再試行します)。設定上限") > 0), "実際=" & m3

    ' 何も無ければ空(正常に読み切った資料のカードに注意書きを出さない)。
    modTestRunner.Check "メモ合成_全部空なら空", _
        (LenB(optOcrEta.ComposeOcrMemo("", "", "")) = 0)
    ' 片方だけのときは余計な区切りを付けない。
    modTestRunner.Check "メモ合成_本体だけなら本体そのまま", _
        (optOcrEta.ComposeOcrMemo("", part, "") = part)
    modTestRunner.Check "メモ合成_capだけならcapそのまま", _
        (optOcrEta.ComposeOcrMemo("", "", cap) = cap)
    ' 中断の理由も同じ口を通る(置換されずに冒頭の復元と共存する)。
    Dim m4 As String
    m4 = optOcrEta.ComposeOcrMemo(head, optOcrEta.OcrAbortMemoFor(120, "cancel", True), "")
    modTestRunner.Check "メモ合成_中断メモも復元を消さない", _
        (Left$(m4, Len(head)) = head And InStr(m4, "利用者の操作で中断しました") > 0), _
        "実際=" & m4
End Sub

' ----------------------------------------------------------------------------
' FA-4: 頁欠けメモの分母(総頁が分かっているならそれで言う)。
' ----------------------------------------------------------------------------
Private Sub TestPartialMemoTotalKnown()
    modTestRunner.Check "頁欠けメモ_総頁が分かればそれを分母にする", _
        (InStr(optOcrEta.OcrPartialMemoFor(197, 3, 254), "全254頁中3頁") > 0), _
        "実際=" & optOcrEta.OcrPartialMemoFor(197, 3, 254)
    modTestRunner.Check "頁欠けメモ_総頁0なら従来どおりok+failが分母", _
        (InStr(optOcrEta.OcrPartialMemoFor(197, 3, 0), "全200頁中3頁") > 0), _
        "実際=" & optOcrEta.OcrPartialMemoFor(197, 3, 0)
    modTestRunner.Check "頁欠けメモ_総頁が小さすぎるときは合計を採る", _
        (InStr(optOcrEta.OcrPartialMemoFor(197, 3, 10), "全200頁中3頁") > 0)
    modTestRunner.Check "頁欠けメモ_失敗0なら総頁を渡しても空のまま", _
        (LenB(optOcrEta.OcrPartialMemoFor(200, 0, 254)) = 0)
End Sub

' ----------------------------------------------------------------------------
' FA-5(レビューB-H3): 画像化フェーズの待ちの上限と実況。
' FA-1(B-M): 確認文は埋め込み時間にも正直に言う。
' ----------------------------------------------------------------------------
Private Sub TestBatchWaitAndBanner()
    ' 1バッチ上限 = max(120, 20頁×16秒) = 320秒。
    modTestRunner.Check "バッチ待ち_残り予算が大きくても1バッチ320秒で頭打ち", _
        (optOcrEta.BatchWaitSec(2032, 0, 20) = 320), _
        "実際=" & optOcrEta.BatchWaitSec(2032, 0, 20)
    modTestRunner.Check "バッチ待ち_残りが上限より短ければ残りが勝つ", _
        (optOcrEta.BatchWaitSec(2032, 1900, 20) = 132), _
        "実際=" & optOcrEta.BatchWaitSec(2032, 1900, 20)
    modTestRunner.Check "バッチ待ち_予算を使い切ったら0(打ち切り)", _
        (optOcrEta.BatchWaitSec(1200, 1200, 20) = 0)
    modTestRunner.Check "バッチ待ち_下限120秒(頁数が極端に小さくても)", _
        (optOcrEta.BatchWaitSec(2032, 0, 1) = 120), _
        "実際=" & optOcrEta.BatchWaitSec(2032, 0, 1)
    modTestRunner.Check "バッチ待ち_残り10秒の床は従来どおり効く", _
        (optOcrEta.BatchWaitSec(1200, 1197, 20) = 10)
    modTestRunner.Check "バッチ待ち_予算0なら0(RemainingWaitSecと同じ答え)", _
        (optOcrEta.BatchWaitSec(0, 0, 20) = 0)

    ' 画像化の実況。総頁が分からないときはバッチ番号だけ(嘘の分母を出さない)。
    modTestRunner.Check "画像化実況_バッチと経過秒を言う", _
        (optOcrEta.RenderWaitBanner(optOcrEta.BatchLabel(3, 254, 20), 42) = _
         "画像化中… バッチ3/13 (全254頁) 経過42秒"), _
        "実際=" & optOcrEta.RenderWaitBanner(optOcrEta.BatchLabel(3, 254, 20), 42)
    modTestRunner.Check "画像化実況_経過0秒は言わない", _
        (optOcrEta.RenderWaitBanner(optOcrEta.BatchLabel(2, 0, 20), 0) = _
         "画像化中… バッチ2"), _
        "実際=" & optOcrEta.RenderWaitBanner(optOcrEta.BatchLabel(2, 0, 20), 0)

    ' FA-1(B-M): 読み取りが終わってもベクトル化が続くことを先に言う。
    Dim ask As String: ask = optOcrEta.OcrConfirmAskFor(254, 106, 15)
    modTestRunner.Check "確認文_完了後にも準備が続くことを言う", _
        (InStr(ask, "完了後に検索用の準備が続きます") > 0), "実際=" & ask
    modTestRunner.Check "確認文_従来の要素(頁数・推定・中断・再開)も残る", _
        (InStr(ask, "全254頁") > 0 And InStr(ask, "推定約106分") > 0 And _
         InStr(ask, "中断") > 0 And InStr(ask, "続きから再開") > 0), "実際=" & ask
End Sub

' ----------------------------------------------------------------------------
' FB-4(レビューB-M): 1頁あたり実績の丸め。
'   0(=この端末でまだ1度も読んでいない)は0のまま返すことが最も大事:
'   ここで3秒へ押し上げると、実測が1つも無い端末が「残り約3秒」と言い始める。
' ----------------------------------------------------------------------------
Private Sub TestClampPageMs()
    modTestRunner.Check "レート丸め_実測なし(0)は0のまま", _
        (optOcrEta.ClampPageMs(0#) = 0#)
    modTestRunner.Check "レート丸め_負の値も0(壊れたui_state)", _
        (optOcrEta.ClampPageMs(-500#) = 0#)
    modTestRunner.Check "レート丸め_下限3秒(1ms/頁は有り得ない)", _
        (optOcrEta.ClampPageMs(1#) = 3000#), _
        "実際=" & optOcrEta.ClampPageMs(1#)
    modTestRunner.Check "レート丸め_上限120秒(GSのハングを含んだ値を持ち越さない)", _
        (optOcrEta.ClampPageMs(600000#) = 120000#), _
        "実際=" & optOcrEta.ClampPageMs(600000#)
    modTestRunner.Check "レート丸め_境界3000はそのまま", _
        (optOcrEta.ClampPageMs(3000#) = 3000#)
    modTestRunner.Check "レート丸め_境界120000はそのまま", _
        (optOcrEta.ClampPageMs(120000#) = 120000#)
    modTestRunner.Check "レート丸め_実機実測(25秒/頁)は素通し", _
        (optOcrEta.ClampPageMs(25000#) = 25000#)
    ' 丸めた値でETAが破綻しないこと(丸め→見積もりの経路をつなぐ)。
    modTestRunner.Check "レート丸め_丸めた値で見積もりが出る(120頁×3秒=6分)", _
        (optOcrEta.OcrEstMinutes(120, optOcrEta.ClampPageMs(1#)) = 6), _
        "実際=" & optOcrEta.OcrEstMinutes(120, optOcrEta.ClampPageMs(1#))
End Sub

' ----------------------------------------------------------------------------
' FB-5(レビューB-M): 見送りメモの判定。文面を作る側(optOcrEta)と数える側
'   (modShelf/modShelfBatch)が同じ答えを出すことを、実際に作った文で固定する。
' ----------------------------------------------------------------------------
Private Sub TestIsDeclineNote()
    modTestRunner.Check "見送り判定_実際の見送りメモを見送りと判定する", _
        (modUtilText.IsDeclineNote(optOcrEta.OcrDeclineMemoFor(106)) = True), _
        "実際=" & optOcrEta.OcrDeclineMemoFor(106)
    modTestRunner.Check "見送り判定_分数が変わっても判定は揺れない", _
        (modUtilText.IsDeclineNote(optOcrEta.OcrDeclineMemoFor(1)) = True And _
         modUtilText.IsDeclineNote(optOcrEta.OcrDeclineMemoFor(9999)) = True)
    ' 中断・頁欠け・上限のメモは【見送りではない】(失敗として数える側へ回る)。
    modTestRunner.Check "見送り判定_中断メモは見送りではない", _
        (modUtilText.IsDeclineNote(optOcrEta.OcrAbortMemoFor(37, "cancel", True)) = False)
    modTestRunner.Check "見送り判定_頁欠けメモは見送りではない", _
        (modUtilText.IsDeclineNote(optOcrEta.OcrPartialMemoFor(197, 3, 254)) = False)
    modTestRunner.Check "見送り判定_上限メモは見送りではない", _
        (modUtilText.IsDeclineNote(optOcrEta.OcrCapMemoFor(True, 200, 254)) = False)
    modTestRunner.Check "見送り判定_空文字は見送りではない", _
        (modUtilText.IsDeclineNote("") = False)
    modTestRunner.Check "見送り判定_無関係な失敗文は見送りではない", _
        (modUtilText.IsDeclineNote("ファイルが大きすぎます。") = False)
    ' 先頭の空白は無視する(manifestを経由して前後が整形されても判定を保つ)。
    modTestRunner.Check "見送り判定_先頭の空白があっても判定できる", _
        (modUtilText.IsDeclineNote("  " & optOcrEta.OcrDeclineMemoFor(30)) = True)
    ' 文中に句が現れるだけの文は見送りにしない(先頭一致であることの確認)。
    modTestRunner.Check "見送り判定_文中の一致では見送りにしない", _
        (modUtilText.IsDeclineNote("前回の続きから再開しました。" & _
            modUtilText.DECLINE_MEMO_HEAD) = False)
End Sub

' ----------------------------------------------------------------------------
' R16-2a(2026-08-05): modUtilText.BuildWorkExcelCmd の完全一致ゴールデン。
'   空文字/末尾に区切り文字があるパス/スペース含みパス/日本語パスを網羅する。
'   区切り文字を含む入力は、ソース中の文字列リテラル末尾へ直書きすると
'   LO構文チェッカーが沈黙ハングするため(docs/dev/EDGE_CASES.md §1.3)、
'   Chr$(92)との連結で作る(関数側の実装も同じ理由でChr$(92)比較にしてある)。
' ----------------------------------------------------------------------------
Private Sub TestBuildWorkExcelCmd()
    Dim noTrail As String: noTrail = "C:\Program Files\Microsoft Office\root\Office16"
    Dim withTrail As String: withTrail = noTrail & Chr$(92)
    Dim wantNormal As String
    wantNormal = """C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"" /x"

    modTestRunner.Check "作業用Excelコマンド_通常パス", _
        (modUtilText.BuildWorkExcelCmd(noTrail) = wantNormal), _
        "実際=" & modUtilText.BuildWorkExcelCmd(noTrail)
    modTestRunner.Check "作業用Excelコマンド_末尾の区切り文字を二重にしない", _
        (modUtilText.BuildWorkExcelCmd(withTrail) = wantNormal), _
        "実際=" & modUtilText.BuildWorkExcelCmd(withTrail)

    modTestRunner.Check "作業用Excelコマンド_空文字でも壊れない", _
        (modUtilText.BuildWorkExcelCmd("") = """\EXCEL.EXE"" /x"), _
        "実際=" & modUtilText.BuildWorkExcelCmd("")

    Dim driveRoot As String: driveRoot = "C:" & Chr$(92)
    modTestRunner.Check "作業用Excelコマンド_ドライブ直下(区切り文字1個のみ)", _
        (modUtilText.BuildWorkExcelCmd(driveRoot) = """C:\EXCEL.EXE"" /x"), _
        "実際=" & modUtilText.BuildWorkExcelCmd(driveRoot)

    Dim withSpace As String: withSpace = "C:\Program Files (x86)\Office"
    modTestRunner.Check "作業用Excelコマンド_スペース含みパス", _
        (modUtilText.BuildWorkExcelCmd(withSpace) = _
         """C:\Program Files (x86)\Office\EXCEL.EXE"" /x"), _
        "実際=" & modUtilText.BuildWorkExcelCmd(withSpace)

    Dim withJp As String: withJp = "C:\日本語フォルダ\Office16"
    modTestRunner.Check "作業用Excelコマンド_日本語パス", _
        (modUtilText.BuildWorkExcelCmd(withJp) = _
         """" & withJp & "\EXCEL.EXE"" /x"), _
        "実際=" & modUtilText.BuildWorkExcelCmd(withJp)
End Sub

Public Sub RunAll14()
    On Error GoTo ComposeFail14
    TestComposeOcrMemo
NextPartial14:
    On Error GoTo PartialFail14
    TestPartialMemoTotalKnown
NextWait14:
    On Error GoTo WaitFail14
    TestBatchWaitAndBanner
NextClamp14:
    On Error GoTo ClampFail14
    TestClampPageMs
NextDecline14:
    On Error GoTo DeclineFail14
    TestIsDeclineNote
NextWorkExcel14:
    On Error GoTo WorkExcelFail14
    TestBuildWorkExcelCmd
NextDone14:
    On Error GoTo 0
    Exit Sub

ComposeFail14:
    modTestRunner.Check "TestComposeOcrMemo(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextPartial14
PartialFail14:
    modTestRunner.Check "TestPartialMemoTotalKnown(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextWait14
WaitFail14:
    modTestRunner.Check "TestBatchWaitAndBanner(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextClamp14
ClampFail14:
    modTestRunner.Check "TestClampPageMs(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDecline14
DeclineFail14:
    modTestRunner.Check "TestIsDeclineNote(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextWorkExcel14
WorkExcelFail14:
    modTestRunner.Check "TestBuildWorkExcelCmd(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone14
End Sub
