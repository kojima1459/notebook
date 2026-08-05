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
    ' 2026-08-05(R17H FA-10 / B-H3): 「準備」だけでは何分続くのか伝わらない。
    ' 章の要約も続くことと、数分かかり得ることを確認文の時点で言う。
    modTestRunner.Check "確認文_完了後にも準備が続くことを言う", _
        (InStr(ask, "完了後に検索用の準備と章の要約が続きます") > 0), "実際=" & ask
    modTestRunner.Check "確認文_その準備が数分かかり得ることも言う", _
        (InStr(ask, "数分かかることがあります") > 0), "実際=" & ask
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

' ----------------------------------------------------------------------------
' R16-3A(2026-08-05): 複合質問の分解→統合。
' ----------------------------------------------------------------------------
' 段0の応答パーサ。ここが「読めない応答で parts へ倒す」と、割れていない
' 論点で検索と生成をN回して、遅くなったうえに薄い答えが出る。全ての異常は
' single(=従来の入念フロー)へ落ちることを固定する。
' ----------------------------------------------------------------------------
Private Sub TestParseDecomposeVerdict()
    Dim v As String
    ' (1) 正常3種。
    modTestRunner.Check "分解判定_single", _
        (modRagParse.ParseDecomposeVerdict("<verdict>single</verdict>") = "single")
    modTestRunner.Check "分解判定_parts", _
        (modRagParse.ParseDecomposeVerdict("<verdict>parts</verdict><parts>a|b</parts>") = "parts")
    modTestRunner.Check "分解判定_clarify", _
        (modRagParse.ParseDecomposeVerdict("<verdict>clarify</verdict><options>x|y</options>") = "clarify")

    ' (2) 大文字・前後空白のゆれ(モデルの書き方は毎回ぶれる)。
    modTestRunner.Check "分解判定_大文字でも読む", _
        (modRagParse.ParseDecomposeVerdict("<VERDICT> PARTS </VERDICT>") = "parts")

    ' (3) タグ欠落・空・不正な語・#ERR はすべて single へ寛容退化。
    modTestRunner.Check "分解判定_タグ欠落はsingle", _
        (modRagParse.ParseDecomposeVerdict("分けたほうがよいと思います") = "single")
    modTestRunner.Check "分解判定_空文字はsingle", _
        (modRagParse.ParseDecomposeVerdict("") = "single")
    modTestRunner.Check "分解判定_中身が空のタグはsingle", _
        (modRagParse.ParseDecomposeVerdict("<verdict></verdict>") = "single")
    modTestRunner.Check "分解判定_知らない語はsingle", _
        (modRagParse.ParseDecomposeVerdict("<verdict>maybe</verdict>") = "single")
    v = modRagParse.ParseDecomposeVerdict("#ERR:E0202:応答が空でした")
    modTestRunner.Check "分解判定_エラー応答はsingle", (v = "single"), "実際=" & v
    ' 閉じタグが無い応答でも「開始タグ以降の全部」を読んで判定できる。
    modTestRunner.Check "分解判定_閉じタグ欠落でも読む", _
        (modRagParse.ParseDecomposeVerdict("<verdict>parts") = "parts")
End Sub

Private Sub TestParseParts()
    Dim p() As String
    Dim n As Long

    ' (1) 正常。区切りは "|"。前後の空白は落ちる。
    n = modRagParse.ParseParts("<parts> A の違い | B の手続き </parts>", 3, p)
    modTestRunner.Check "論点分割_2件", (n = 2), "実際=" & n
    modTestRunner.Check "論点分割_前後空白は落ちる", _
        (n = 2 And p(0) = "A の違い" And p(1) = "B の手続き"), "実際=" & p(0) & "/" & p(1)

    ' (2) 上限で切り詰める(4分割→3)。増やすほどAI呼び出しが増える段なので、
    '     ここが効かないと1質問で何回でも呼べてしまう。
    n = modRagParse.ParseParts("<parts>a|b|c|d</parts>", 3, p)
    modTestRunner.Check "論点分割_上限3で切り詰め", (n = 3), "実際=" & n
    modTestRunner.Check "論点分割_切り詰めても先頭から順に残る", _
        (n = 3 And p(0) = "a" And p(2) = "c"), "実際=" & p(0) & p(1) & p(2)

    ' (3) 空要素は数えない(「a||b」を3件と数えると空の論点を1回調べる)。
    n = modRagParse.ParseParts("<parts>a||b</parts>", 3, p)
    modTestRunner.Check "論点分割_空要素は除去", (n = 2), "実際=" & n

    ' (4) タグ欠落・空・#ERR は0件(呼び出し側が single へ落ちる)。
    modTestRunner.Check "論点分割_タグ欠落は0件", _
        (modRagParse.ParseParts("<verdict>parts</verdict>", 3, p) = 0)
    modTestRunner.Check "論点分割_空タグは0件", _
        (modRagParse.ParseParts("<parts></parts>", 3, p) = 0)
    modTestRunner.Check "論点分割_区切りだけは0件", _
        (modRagParse.ParseParts("<parts> | | </parts>", 3, p) = 0)
    modTestRunner.Check "論点分割_エラー応答は0件", _
        (modRagParse.ParseParts("#ERR:E0202:応答が空でした", 3, p) = 0)
    modTestRunner.Check "論点分割_上限0なら0件", _
        (modRagParse.ParseParts("<parts>a|b</parts>", 0, p) = 0)

    ' (5) 1件しか割れなかった応答も「読めている」こと(切るのは呼び出し側の判断)。
    n = modRagParse.ParseParts("<parts>ひとつだけ</parts>", 3, p)
    modTestRunner.Check "論点分割_1件は1件として返す", (n = 1), "実際=" & n
End Sub

' 発動ゲート(off/auto/always と文字数)と、論点あたりtopKの下限。
Private Sub TestDecomposeGate()
    modTestRunner.Check "分解ゲート_offは呼ばない", _
        (modAskMulti.ShouldDecompose("off", 200, 25) = False)
    modTestRunner.Check "分解ゲート_alwaysは短くても呼ぶ", _
        (modAskMulti.ShouldDecompose("always", 1, 25) = True)
    modTestRunner.Check "分解ゲート_autoは既定25字未満で呼ばない", _
        (modAskMulti.ShouldDecompose("auto", 24, 25) = False)
    modTestRunner.Check "分解ゲート_autoは境界25字ちょうどで呼ぶ", _
        (modAskMulti.ShouldDecompose("auto", 25, 25) = True)
    ' config の打ち間違いで機能が黙って止まらない(未知の値は auto 扱い)。
    modTestRunner.Check "分解ゲート_知らない値はauto扱い", _
        (modAskMulti.ShouldDecompose("ｵﾝ", 30, 25) = True And _
         modAskMulti.ShouldDecompose("", 10, 25) = False)
    modTestRunner.Check "分解ゲート_大文字と前後空白を吸収", _
        (modAskMulti.ShouldDecompose(" OFF ", 200, 25) = False)
    ' 0以下の閾値は既定25へ倒す(configを0にして全質問で呼ばせない)。
    modTestRunner.Check "分解ゲート_閾値0以下は既定25", _
        (modAskMulti.ShouldDecompose("auto", 24, 0) = False And _
         modAskMulti.ShouldDecompose("auto", 25, 0) = True)

    modTestRunner.Check "論点topK_16を3論点なら下限6", _
        (modAskMulti.PerPartTopK(16, 3) = 6), "実際=" & modAskMulti.PerPartTopK(16, 3)
    modTestRunner.Check "論点topK_16を2論点なら8", _
        (modAskMulti.PerPartTopK(16, 2) = 8), "実際=" & modAskMulti.PerPartTopK(16, 2)
    modTestRunner.Check "論点topK_大きい設定はそのまま割る", _
        (modAskMulti.PerPartTopK(30, 3) = 10)
    modTestRunner.Check "論点topK_論点0でも下限6(0除算にしない)", _
        (modAskMulti.PerPartTopK(16, 0) = 6)
End Sub

' ----------------------------------------------------------------------------
' 統合段への入力(節の組み立て)。ここが崩れると、統合LLMへ渡す前の時点で
' 出典タグが落ちたり、調べられなかった論点が黙って消えたりする。
' ----------------------------------------------------------------------------
Private Sub TestPartSection()
    Dim s As String

    ' (1) 出典タグは節の組み立てで一字一句そのまま通る。
    s = modAskMulti.BuildPartSection(1, "Aの違いは?", _
        "結論です。[本棚:規約集 p.12]と[パック(山田):承認フロー]が根拠です。")
    Dim head1 As String: head1 = "■Aの違いは?"
    modTestRunner.Check "統合入力_見出しは副質問", _
        (Left$(s, Len(head1)) = head1), "実際=" & s
    modTestRunner.Check "統合入力_本棚タグがそのまま残る", _
        (InStr(s, "[本棚:規約集 p.12]") > 0), "実際=" & s
    modTestRunner.Check "統合入力_パックタグがそのまま残る", _
        (InStr(s, "[パック(山田):承認フロー]") > 0), "実際=" & s

    ' (2) 部分失敗の文言は固定(統合段へ「この文言は消すな」と指示する文字列と
    '     同じでなければ、指示と実物がずれて黙って言い換えられる)。
    s = modAskMulti.BuildPartSection(2, "Bの手続きは?", "")
    modTestRunner.Check "統合入力_失敗節の文言が固定", _
        (s = "■論点2: " & modPrompts.PART_FAIL_TEXT), "実際=" & s
    modTestRunner.Check "統合入力_失敗節は空白だけの本文でも同じ", _
        (modAskMulti.BuildPartSection(2, "Bの手続きは?", "   ") = s), _
        "実際=" & modAskMulti.BuildPartSection(2, "Bの手続きは?", "   ")

    ' (3) 副質問が空でも見出しは必ず立つ(見出しが消えると統合段が節を融かす)。
    modTestRunner.Check "統合入力_副質問が空なら論点nを見出しに", _
        (Left$(modAskMulti.BuildPartSection(3, "", "本文"), 4) = "■論点3"), _
        "実際=" & modAskMulti.BuildPartSection(3, "", "本文")

    ' (4) 統合プロンプトは、渡した節をそのまま抱え、タグ保持を明示する。
    Dim sections As String
    sections = modAskMulti.BuildPartSection(1, "Aの違いは?", "甲です。[本棚:規約集 p.12]") & _
        vbLf & vbLf & modAskMulti.BuildPartSection(2, "Bの手続きは?", "")
    Dim mp As String
    mp = modPrompts.BuildMergePrompt("AとBについて教えて", sections)
    modTestRunner.Check "統合プロンプト_節の出典タグが本文に載る", _
        (InStr(mp, "[本棚:規約集 p.12]") > 0), ""
    modTestRunner.Check "統合プロンプト_一字一句残せと書いてある", _
        (InStr(mp, "一字一句そのまま残すこと") > 0), ""
    modTestRunner.Check "統合プロンプト_失敗節の文言を残せと書いてある", _
        (InStr(mp, modPrompts.PART_FAIL_TEXT) > 0), ""
    modTestRunner.Check "統合プロンプト_元の質問も載る", _
        (InStr(mp, "AとBについて教えて") > 0), ""
    ' 全滅(節が1つも無い)ときは、そもそも統合を呼ばずエラー回答へ倒す設計。
    ' ここでは「空を渡しても壊れない」ことだけを固定する。
    modTestRunner.Check "統合プロンプト_節が空でも壊れない", _
        (Len(modPrompts.BuildMergePrompt("q", "")) > 0)

    ' (5) 段0プロンプトは3タグの出力契約を必ず含む(パーサと対になる規約)。
    Dim dp As String
    dp = modPrompts.BuildDecomposePrompt("AとBの違いとCの手続きは?", "", 3)
    modTestRunner.Check "分解プロンプト_verdictタグを要求", (InStr(dp, "<verdict>") > 0), ""
    modTestRunner.Check "分解プロンプト_partsタグを要求", (InStr(dp, "<parts>") > 0), ""
    modTestRunner.Check "分解プロンプト_optionsタグを要求", (InStr(dp, "<options>") > 0), ""
    modTestRunner.Check "分解プロンプト_上限件数を書く", (InStr(dp, "最大3個") > 0), "実際=" & dp
    modTestRunner.Check "分解プロンプト_上限は2から5へ丸める", _
        (InStr(modPrompts.BuildDecomposePrompt("q", "", 99), "最大5個") > 0 And _
         InStr(modPrompts.BuildDecomposePrompt("q", "", 0), "最大2個") > 0)
End Sub

' ----------------------------------------------------------------------------
' 統合後テキストに対する出典突合のゴールデン(■見出し複数+複数論点のタグ+
' 1件だけ偽タグ)。分解経路では索引を unionHits から作るので、論点2の資料を
' 引いた文が論点1の索引に無くて「確認できず」になる、という壊れ方をしない
' ことをここで固定する。
' ----------------------------------------------------------------------------
Private Sub TestMergedAnnotate()
    ' unionHits から作られる突合表(空白は落ちた形)。論点1=規約集、
    ' 論点2=承認フロー(パック)+規約集p.13。
    Dim idx As String
    idx = "|[本棚:規約集p.12]||[本棚:規約集p.13]||[パック(山田):承認フロー]|"

    Dim src As String
    src = "【結論】AとBのどちらも手続きが要ります。" & vbLf & vbLf & _
          "■Aの違いは?" & vbLf & _
          "・甲は乙より優先されます。[本棚:規約集 p.12]" & vbLf & vbLf & _
          "■Bの手続きは?" & vbLf & _
          "・課長承認が必要です。[パック(山田):承認フロー]" & vbLf & _
          "・提出期限は5日以内。[本棚:規約集 p.13]" & vbLf & _
          "・例外は年度末のみ。[本棚:存在しない手引き.pdf p.9]"

    Dim mism As Long
    Dim r As String
    r = modAskThorough.AnnotateCitations(src, idx, mism)

    modTestRunner.Check "統合後突合_偽タグは1件だけ", (mism = 1), "実際=" & mism
    modTestRunner.Check "統合後突合_偽タグの直後に付記", _
        (InStr(r, "[本棚:存在しない手引き.pdf p.9]" & modAskThorough.UNVERIFIED_MARK) > 0), _
        "実際=" & r
    ' 論点をまたいだタグに付記が付かないこと(索引を論点ごとに作ると全滅する)。
    modTestRunner.Check "統合後突合_論点1のタグは無傷", _
        (InStr(r, "[本棚:規約集 p.12]" & modAskThorough.UNVERIFIED_MARK) = 0), "実際=" & r
    modTestRunner.Check "統合後突合_論点2のパックタグは無傷", _
        (InStr(r, "[パック(山田):承認フロー]" & modAskThorough.UNVERIFIED_MARK) = 0), _
        "実際=" & r
    modTestRunner.Check "統合後突合_論点2の本棚タグは無傷", _
        (InStr(r, "[本棚:規約集 p.13]" & modAskThorough.UNVERIFIED_MARK) = 0), "実際=" & r
    ' ■見出しは1つも欠けない(突合は本文を線形に写すだけで構造を壊さない)。
    modTestRunner.Check "統合後突合_見出しが2つとも残る", _
        (InStr(r, "■Aの違いは?") > 0 And InStr(r, "■Bの手続きは?") > 0), "実際=" & r
    modTestRunner.Check "統合後突合_付記1件ぶんだけ長くなる", _
        (Len(r) = Len(src) + Len(modAskThorough.UNVERIFIED_MARK)), _
        "実際=" & Len(r) & " 元=" & Len(src)

    ' 部分失敗の節が混じっていても、そこにはタグが無いので何も付かない。
    Dim src2 As String
    src2 = "■Aの違いは?" & vbLf & "・甲です。[本棚:規約集 p.12]" & vbLf & vbLf & _
           "■論点2: " & modPrompts.PART_FAIL_TEXT
    r = modAskThorough.AnnotateCitations(src2, idx, mism)
    modTestRunner.Check "統合後突合_失敗節には何も付かない", (mism = 0), "実際=" & mism
    modTestRunner.Check "統合後突合_失敗節の文言が残る", _
        (InStr(r, modPrompts.PART_FAIL_TEXT) > 0), "実際=" & r
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
NextVerdict14:
    On Error GoTo VerdictFail14
    TestParseDecomposeVerdict
NextParts14:
    On Error GoTo PartsFail14
    TestParseParts
NextGate14:
    On Error GoTo GateFail14
    TestDecomposeGate
NextSection14:
    On Error GoTo SectionFail14
    TestPartSection
NextMerged14:
    On Error GoTo MergedFail14
    TestMergedAnnotate
NextDone14:
    On Error GoTo 0
    ' R16波3: 容量逼迫による分割先(modTestsPure15)を末尾から連鎖で呼ぶ。
    modTestsPure15.RunAll15
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
    Resume NextVerdict14
VerdictFail14:
    modTestRunner.Check "TestParseDecomposeVerdict(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextParts14
PartsFail14:
    modTestRunner.Check "TestParseParts(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGate14
GateFail14:
    modTestRunner.Check "TestDecomposeGate(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextSection14
SectionFail14:
    modTestRunner.Check "TestPartSection(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextMerged14
MergedFail14:
    modTestRunner.Check "TestMergedAnnotate(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone14
End Sub
