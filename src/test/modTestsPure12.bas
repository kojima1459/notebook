Attribute VB_Name = "modTestsPure12"
Option Explicit

' ============================================================================
' modTestsPure12 - R14-8(入念モードの本格強化 + 回答の可読性)の
'                  純ロジック回帰テスト
' ----------------------------------------------------------------------------
' なぜ新しいモジュールなのか:
'   R14-8 のテストを modTestsPure11 へ足すと 30,000字上限を超えた。
'   憲章§4-6「WARN帯のモジュールに機能を足さない。足す前に分割を裁定する」に
'   従い、modTestsPure10 → 11 と同じ形で分割先を新設した。入口は
'   modTestsPure11.RunAll11 の末尾から呼ばれる Public Sub RunAll12()。
'   導線はその1行だけで、消すとここのテストは「実行されないまま」
'   全部PASSに見える。
'
' 固定する事実:
'   ・modMode.Description(R14-8b): 3モードの説明。入念は所要時間(2～4分)を
'     隠さない(隠すと「固まった」と思われる。憲章§3-2)。
'   ・modMode.RerankEffort(R14-8a): 入念だけ再ランクの effort を分ける。
'   ・modMode.AskStage*(R14-8a): 入念=最大6段の誠実な段数。深掘り側の番号が
'     1つも動かないこと(動くと R13-9b の回帰)。
'   ・modAskThorough の出典突合(R14-8a-6): 回答中の出典タグを検索結果と
'     機械的に照合する純ロジック。ここがズレると「存在しない資料名やページを
'     引用しても誰も気付けない」状態(実機第3報 RC8)へ戻る。
'   ・modLive.NormalizeAnswerText / AnswerParagraphs / Humanize(R14-8c):
'     混入したMarkdownの保険変換と、実況の言い換えで段数の番号を落とさないこと。
'   ・optOcrEta.RemainingWaitSec(R14-F6): gs_abs_timeout_sec は【1資料
'     あたり】の絶対上限。バッチ描画で「1バッチあたり」に化けていたので、
'     何バッチ回しても合計が上限を超えないことをここで固定する
'     (modTestsPure11 が28,000字のWARN帯に入るため置き場はこちら)。
' ============================================================================

' R14-8b: モード説明(押すたびにトーストで出る唯一の説明)。
Private Sub TestModeDescriptionsR14()
    Dim dq As String: dq = modMode.Description("quick")
    Dim dd As String: dd = modMode.Description("deep")
    Dim dt As String: dt = modMode.Description("thorough")

    modTestRunner.Check "モード説明_すぐ聞くは速さ優先", (InStr(dq, "速さ優先") > 0), "実際=" & dq
    modTestRunner.Check "モード説明_深掘りは会話の流れ", (InStr(dd, "会話の流れ") > 0), "実際=" & dd
    modTestRunner.Check "モード説明_入念は多段検証", (InStr(dt, "多段検証") > 0), "実際=" & dt
    ' 所要時間を隠さない(隠すと「固まった」と思われる。憲章§3-2)。
    modTestRunner.Check "モード説明_入念は2～4分と明言", (InStr(dt, "2～4分") > 0), "実際=" & dt
    modTestRunner.Check "モード説明_3つとも別の文", _
        (dq <> dd And dd <> dt And dq <> dt), ""
End Sub

' R14-8a: 入念モードだけ再ランクの effort を分ける。
Private Sub TestRerankEffortThorough()
    modTestRunner.Check "再ランクeffort_入念は専用値", _
        (modMode.RerankEffort("thorough", "low", "medium") = "medium")
    modTestRunner.Check "再ランクeffort_深掘りは従来値", _
        (modMode.RerankEffort("deep", "low", "medium") = "low")
    modTestRunner.Check "再ランクeffort_すぐ聞くは従来値", _
        (modMode.RerankEffort("quick", "low", "medium") = "low")
    ' configを空にしたら従来値へ倒れる(設定ミスで空文字を送らない)。
    modTestRunner.Check "再ランクeffort_専用値が空なら従来値", _
        (modMode.RerankEffort("thorough", "low", "") = "low")
    modTestRunner.Check "再ランクeffort_専用値が空白のみなら従来値", _
        (modMode.RerankEffort("thorough", "low", "   ") = "low")
End Sub

' ----------------------------------------------------------------------------
' R14-8a: 段階ナレーションの誠実な段数(入念=最大6段)。総数を決め打ちに
'   すると「(3/4)で終わる」という嘘になる。入念で増えた2段が総数と番号の
'   両方へ効くこと、深掘り側の番号が1つも動かないこと(R13-9bの回帰)。
' ----------------------------------------------------------------------------
Private Sub TestAskStageThorough()
    ' 総数: 拡張+再ランク+要点整理+下書き+自己批判+検証 = 6
    modTestRunner.Check "段階_入念フル構成は6段", _
        (modMode.AskStageTotal(True, True, True, True) = 6)
    modTestRunner.Check "段階_入念で拡張も再ランクも無ければ4段", _
        (modMode.AskStageTotal(False, False, True, True) = 4)
    ' 深掘りは従来どおり(R13-9bの回帰検知)。
    modTestRunner.Check "段階_深掘りフル構成は4段のまま", _
        (modMode.AskStageTotal(True, True, True, False) = 4)
    modTestRunner.Check "段階_検証を通さないなら番号を出さない", _
        (modMode.AskStageTotal(False, False, False, True) = 0)

    ' 番号: 拡張1 → 再ランク2 → 要点整理3 → 下書き4 → 自己批判5 → 検証6
    modTestRunner.Check "段階_入念の要点整理は3", _
        (modMode.AskStageIndex("digest", True, True, True) = 3)
    modTestRunner.Check "段階_入念の下書きは4", _
        (modMode.AskStageIndex("draft", True, True, True) = 4)
    modTestRunner.Check "段階_入念の自己点検は5", _
        (modMode.AskStageIndex("critique", True, True, True) = 5)
    modTestRunner.Check "段階_入念の検証は最後の6", _
        (modMode.AskStageIndex("verify", True, True, True) = 6)

    ' 深掘り(入念ではない)の番号は1つも動かない。
    modTestRunner.Check "段階_深掘りの下書きは3のまま", _
        (modMode.AskStageIndex("draft", True, True) = 3)
    modTestRunner.Check "段階_深掘りの検証は4のまま", _
        (modMode.AskStageIndex("verify", True, True) = 4)
    ' 計画に入っていない段は番号を出さない(嘘の番号を出すより出さない)。
    modTestRunner.Check "段階_入念以外に要点整理の番号は無い", _
        (modMode.AskStageIndex("digest", True, True) = 0)
    modTestRunner.Check "段階_入念以外に自己点検の番号は無い", _
        (modMode.AskStageIndex("critique", True, True) = 0)

    ' ラベルは利用者の言葉で、段ごとに違うこと。
    Dim lbl As String: lbl = modMode.AskStageLabel("critique")
    modTestRunner.Check "段階_要点整理のラベル", _
        (InStr(modMode.AskStageLabel("digest"), "要点") > 0)
    modTestRunner.Check "段階_自己点検のラベル", (InStr(lbl, "自己点検") > 0), "実際=" & lbl
    modTestRunner.Check "段階_6段の実況テキストは番号つき", _
        (Left$(modMode.AskStageText(5, 6, lbl), 6) = "(5/6) "), _
        "実際=" & modMode.AskStageText(5, 6, lbl)
End Sub

' ----------------------------------------------------------------------------
' R14-8a(6): 出典タグの抽出と突合(実機第3報 RC8「機械的な突合が皆無」)。
'   モデルに「この出典は正しいか」と聞いても自分の書いたタグを肯定する。
'   この表がズレると、全件一致(存在しない引用が素通り)か全件不一致
'   (正しい回答が「確認できず」だらけ)のどちらかへ静かに退化する。
' ----------------------------------------------------------------------------
Private Sub TestCiteTagParsing()
    ' 正規化: 空白(半角/全角)は突合の邪魔にしかならないので全部落とす。
    modTestRunner.Check "出典タグ正規化_半角空白を落とす", _
        (modAskThorough.NormalizeCiteTag("[本棚: 規約集 p. 12 ]") = "[本棚:規約集p.12]"), _
        "実際=" & modAskThorough.NormalizeCiteTag("[本棚: 規約集 p. 12 ]")
    modTestRunner.Check "出典タグ正規化_全角空白を落とす", _
        (modAskThorough.NormalizeCiteTag("[本棚:規約集" & ChrW(&H3000) & "p.12]") = "[本棚:規約集p.12]"), _
        "実際=" & modAskThorough.NormalizeCiteTag("[本棚:規約集" & ChrW(&H3000) & "p.12]")

    ' 出典タグらしさ: §7.3の2形式だけを拾う。
    modTestRunner.Check "出典タグ判定_本棚形式", _
        (modAskThorough.IsCiteTag("[本棚:規約集 p.12]") = True)
    modTestRunner.Check "出典タグ判定_パック形式", _
        (modAskThorough.IsCiteTag("[パック(山田):承認フロー]") = True)
    modTestRunner.Check "出典タグ判定_深掘り候補は出典ではない", _
        (modAskThorough.IsCiteTag("[[FOLLOWUP: 候補1 | 候補2]") = False)
    modTestRunner.Check "出典タグ判定_ただの角括弧は出典ではない", _
        (modAskThorough.IsCiteTag("[注記]") = False)

    ' 抽出: 出現順に、閉じ括弧の位置つきで返す。
    Dim tags() As String
    Dim tagEnd() As Long
    Dim body As String
    body = "結論です。[本棚:規約集 p.12]" & vbLf & "・補足[パック(山田):承認フロー]です。"
    Dim n As Long
    n = modAskThorough.ExtractCiteTags(body, tags, tagEnd)
    modTestRunner.Check "出典タグ抽出_2件", (n = 2), "実際=" & n
    modTestRunner.Check "出典タグ抽出_1件目は本棚", _
        (tags(0) = "[本棚:規約集 p.12]"), "実際=" & tags(0)
    modTestRunner.Check "出典タグ抽出_2件目はパック", _
        (tags(1) = "[パック(山田):承認フロー]"), "実際=" & tags(1)
    modTestRunner.Check "出典タグ抽出_閉じ位置は昇順", (tagEnd(1) > tagEnd(0)), ""

    ' タグが1つも無い回答(一般的な説明だけの回答)で壊れないこと。
    modTestRunner.Check "出典タグ抽出_タグ無しは0件", _
        (modAskThorough.ExtractCiteTags("出典の無い回答です。", tags, tagEnd) = 0), ""
    modTestRunner.Check "出典タグ抽出_空文字は0件", _
        (modAskThorough.ExtractCiteTags("", tags, tagEnd) = 0), ""
    ' 閉じ括弧が無い壊れた応答でも止まらないこと(無限ループにしない)。
    modTestRunner.Check "出典タグ抽出_閉じ括弧なしでも0件で返る", _
        (modAskThorough.ExtractCiteTags("[本棚:規約集 p.12 が根拠です", tags, tagEnd) = 0), ""
End Sub

Private Sub TestAnnotateCitations()
    ' 検索結果(hits)から作った「正しい出典タグの集合」。
    Dim idx As String
    idx = "|[本棚:規約集p.12]||[パック(山田):承認フロー]|"

    Dim mism As Long
    Dim r As String

    ' 一致するタグには何も足さない。
    r = modAskThorough.AnnotateCitations("根拠です。[本棚:規約集 p.12]", idx, mism)
    modTestRunner.Check "出典突合_一致は付記しない", _
        (InStr(r, modAskThorough.UNVERIFIED_MARK) = 0), "実際=" & r
    modTestRunner.Check "出典突合_一致は不一致0件", (mism = 0), "実際=" & mism

    ' 空白の入り方が違っても一致(モデルの書き方は毎回ぶれる)。
    r = modAskThorough.AnnotateCitations("根拠です。[本棚: 規約集  p. 12]", idx, mism)
    modTestRunner.Check "出典突合_空白ゆれは一致扱い", (mism = 0), "実際=" & r

    ' ページ違いは不一致(±0で見る。原文を開いて「書いていない」が一番効く裏切り)。
    r = modAskThorough.AnnotateCitations("根拠です。[本棚:規約集 p.13]", idx, mism)
    modTestRunner.Check "出典突合_ページ違いは不一致", (mism = 1), "実際=" & mism
    modTestRunner.Check "出典突合_付記はタグの直後", _
        (InStr(r, "[本棚:規約集 p.13]" & modAskThorough.UNVERIFIED_MARK) > 0), "実際=" & r

    ' 資料名そのものが検索結果に無い(でっち上げ)。
    r = modAskThorough.AnnotateCitations("[本棚:存在しない資料.pdf p.1]です", idx, mism)
    modTestRunner.Check "出典突合_知らない資料は不一致", (mism = 1), "実際=" & mism

    ' 複数タグ: 正しいものは触らず、間違ったものだけに付ける。
    r = modAskThorough.AnnotateCitations( _
        "A[本棚:規約集 p.12]B[本棚:規約集 p.99]C[パック(山田):承認フロー]D", idx, mism)
    modTestRunner.Check "出典突合_複数タグでも不一致は1件", (mism = 1), "実際=" & mism
    modTestRunner.Check "出典突合_誤りタグにだけ付記", _
        (InStr(r, "[本棚:規約集 p.99]" & modAskThorough.UNVERIFIED_MARK) > 0), "実際=" & r
    modTestRunner.Check "出典突合_正しいタグの直後には付けない", _
        (InStr(r, "[本棚:規約集 p.12]" & modAskThorough.UNVERIFIED_MARK) = 0), "実際=" & r
    modTestRunner.Check "出典突合_本文は欠けない", _
        (InStr(r, "A") > 0 And InStr(r, "B") > 0 And InStr(r, "C") > 0 And InStr(r, "D") > 0), _
        "実際=" & r

    ' タグが無い回答はそのまま返す。
    r = modAskThorough.AnnotateCitations("出典の無い回答です。", idx, mism)
    modTestRunner.Check "出典突合_タグ無しは原文のまま", _
        (r = "出典の無い回答です。" And mism = 0), "実際=" & r

    ' 深掘り候補のマーカーは出典ではないので触らない(触ると候補が汚れる)。
    r = modAskThorough.AnnotateCitations("本文[[FOLLOWUP: 候補1 | 候補2]]", idx, mism)
    modTestRunner.Check "出典突合_深掘り候補には付記しない", (mism = 0), "実際=" & r

    ' 突合の材料が無いときは何も言わない(検査したふりをしない)。
    r = modAskThorough.AnnotateCitations("[本棚:規約集 p.12]", "", mism)
    modTestRunner.Check "出典突合_材料が無ければ付記しない", _
        (r = "[本棚:規約集 p.12]" And mism = 0), "実際=" & r
End Sub

' ----------------------------------------------------------------------------
' R14-G3: 資料名に "]" を含むファイル名(report[1].pdf 等。ブラウザやメールの
'   重複ダウンロードが普通に作る)。タグの抽出は最初の "]" で切るため、
'   突合表と一致せず【全件が不一致】になり、しかも付記がタグの途中へ入って
'   本文に ".pdf p.3]" が取り残されていた。次の "]" まで伸ばして再照合する。
' ----------------------------------------------------------------------------
Private Sub TestCiteTagWithBracketName()
    ' 突合表は CiteIndexFrom が作る形(NormalizeCiteTag 済み=空白は落ちている)。
    Dim idx As String
    idx = "|[本棚:report[1].pdfp.3]||[パック(山田):月報[2].docx]||[本棚:規約集p.12]|"
    Dim mism As Long
    Dim r As String

    ' 伸ばせば一致する: 付記なし・不一致0件・本文が欠けない。
    r = modAskThorough.AnnotateCitations("結論です。[本棚:report[1].pdf p.3]以上。", idx, mism)
    modTestRunner.Check "角括弧名_伸ばして一致すれば付記しない", _
        (InStr(r, modAskThorough.UNVERIFIED_MARK) = 0), "実際=" & r
    modTestRunner.Check "角括弧名_不一致は0件", (mism = 0), "実際=" & mism
    modTestRunner.Check "角括弧名_本文は原文のまま", _
        (r = "結論です。[本棚:report[1].pdf p.3]以上。"), "実際=" & r

    ' パック形式("(" と "]" の両方が混じる形)でも同じ。
    r = modAskThorough.AnnotateCitations("A[パック(山田):月報[2].docx]B", idx, mism)
    modTestRunner.Check "角括弧名_パック形式も一致", (mism = 0), "実際=" & r

    ' 伸ばしても知らない資料は不一致。付記は【本当のタグ末尾】の直後へ入る
    ' (途中へ入れると ".pdf p.9]" が本文に取り残されて読めなくなる)。
    r = modAskThorough.AnnotateCitations("根拠[本棚:報告[9].pdf p.9]です", idx, mism)
    modTestRunner.Check "角括弧名_知らない資料は不一致1件", (mism = 1), "実際=" & mism
    modTestRunner.Check "角括弧名_付記は最初の閉じ括弧の直後(伸ばさない)", _
        (InStr(r, "[本棚:報告[9]" & modAskThorough.UNVERIFIED_MARK) > 0), "実際=" & r

    ' 伸ばす回数には上限がある(壊れた応答で延々と伸ばさない)。
    r = modAskThorough.AnnotateCitations("[本棚:a]b]c]d]e] p.1]", idx, mism)
    modTestRunner.Check "角括弧名_伸ばす上限を超えても止まる", (mism = 1), "実際=" & mism

    ' 正しいタグと角括弧名タグが混在しても、正しい側は触らない。
    r = modAskThorough.AnnotateCitations( _
        "X[本棚:規約集 p.12]Y[本棚:report[1].pdf p.3]Z", idx, mism)
    modTestRunner.Check "角括弧名_混在でも不一致0件", (mism = 0), "実際=" & r
    modTestRunner.Check "角括弧名_混在でも本文は欠けない", _
        (InStr(r, "X") > 0 And InStr(r, "Y") > 0 And InStr(r, "Z") > 0), "実際=" & r
End Sub

' 検索結果から作る突合表が、LLMへ指示しているタグの形と同じであること。
' ここが別実装になると、検査は全件一致か全件不一致へ静かに退化する。
Private Sub TestCiteIndexFromHits()
    If Not CanUseTypeArrays11() Then
        modTestRunner.Check "[SKIP] 出典突合表(Hit配列): LO環境の既知の制限によりスキップ", True, ""
        Exit Sub
    End If

    Dim h(1 To 2) As Hit
    h(1).source = "規約集": h(1).page = 12: h(1).origin = "self"
    h(2).source = "承認フロー": h(2).page = 1: h(2).origin = "pack:山田"

    Dim idx As String
    idx = modAskThorough.CiteIndexFrom(h, 2)
    modTestRunner.Check "出典突合表_本棚はページ込み", _
        (InStr(idx, "|[本棚:規約集p.12]|") > 0), "実際=" & idx
    modTestRunner.Check "出典突合表_パックは作成者込み", _
        (InStr(idx, "|[パック(山田):承認フロー]|") > 0), "実際=" & idx
    modTestRunner.Check "出典突合表_回答中のタグと突き合う", _
        (modAskThorough.TagIsKnown("[本棚:規約集 p.12]", idx) = True), "実際=" & idx
    modTestRunner.Check "出典突合表_ページ違いは突き合わない", _
        (modAskThorough.TagIsKnown("[本棚:規約集 p.11]", idx) = False), "実際=" & idx
End Sub

' modTestsPure2 の CanUseTypeArrays と同じ実測プローブ(別モジュールのPrivateは
' 呼べないため。テスト限定の軽微な重複で、公開契約を増やすより実害が小さい)。
Private Function CanUseTypeArrays11() As Boolean
    On Error Resume Next
    Err.Clear
    Dim probe() As ShelfChunk
    ReDim probe(0 To 0)
    CanUseTypeArrays11 = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' R14-8c: 回答本文の読みやすさ(実機第3報 RC9)。この画面はMarkdownを描画
'   しない。禁じても長い回答ほどモデルは地の癖で "## " や "**強調**" を
'   混ぜ、記号が生のまま出て壊れて見える。プロンプトは確率・ここは保険。
' ----------------------------------------------------------------------------
Private Sub TestNormalizeAnswerText()
    modTestRunner.Check "可読性_見出し1は■へ", (NA("# 見出し") = "■ 見出し"), "実際=" & NA("# 見出し")
    modTestRunner.Check "可読性_見出し2は■へ", (NA("## 見出し") = "■ 見出し"), "実際=" & NA("## 見出し")
    modTestRunner.Check "可読性_見出し3も■へ", (NA("### 小") = "■ 小"), "実際=" & NA("### 小")
    ' 空白の無い # はタグや番号なので触らない(誤変換のほうが害が大きい)。
    modTestRunner.Check "可読性_空白の無い#は変えない", (NA("#タグ") = "#タグ"), "実際=" & NA("#タグ")

    modTestRunner.Check "可読性_強調は【】へ", (NA("**重要**です") = "【重要】です"), "実際=" & NA("**重要**です")
    modTestRunner.Check "可読性_強調が複数でも全部変換", (NA("**A**と**B**") = "【A】と【B】"), _
        "実際=" & NA("**A**と**B**")

    modTestRunner.Check "可読性_ハイフン箇条書きは・へ", (NA("- 項目") = "・項目"), "実際=" & NA("- 項目")
    modTestRunner.Check "可読性_アスタリスク箇条書きは・へ", (NA("* 項目") = "・項目"), "実際=" & NA("* 項目")

    ' R14-G9: 空行が2つ以上続くのは体裁の事故。1つへ畳む(閾値が「3つ以上」
    ' だったため、一番よく出る「空行2つ」だけが素通りしていた)。
    Dim many As String: many = "本文1" & vbLf & vbLf & vbLf & vbLf & "本文2"
    modTestRunner.Check "可読性_空行3つ以上は1つへ", _
        (NA(many) = "本文1" & vbLf & vbLf & "本文2"), "実際=" & Replace(NA(many), vbLf, "<LF>")
    Dim two As String: two = "本文1" & vbLf & vbLf & vbLf & "本文2"
    modTestRunner.Check "可読性_空行2つも1つへ", _
        (NA(two) = "本文1" & vbLf & vbLf & "本文2"), "実際=" & Replace(NA(two), vbLf, "<LF>")
    ' 空行1つ(=段落の正しい区切り)は畳まない。
    Dim one As String: one = "本文1" & vbLf & vbLf & "本文2"
    modTestRunner.Check "可読性_空行1つはそのまま", _
        (NA(one) = "本文1" & vbLf & vbLf & "本文2"), "実際=" & Replace(NA(one), vbLf, "<LF>")

    ' R14-G10: 中身が空の "** **" で走査を打ち切らない(打ち切ると、それ以降の
    ' 正しい強調が全部 "**" の生記号のまま画面に出る)。
    modTestRunner.Check "可読性_空の強調記号の後も変換を続ける", _
        (NA("** **と**重要**") = "** **と【重要】"), "実際=" & NA("** **と**重要**")
    modTestRunner.Check "可読性_空の強調記号だけなら素通り", _
        (NA("****") = "****"), "実際=" & NA("****")
    ' 既に【】で囲まれている中身は二重に囲まない(【【重要】】は読みにくい)。
    modTestRunner.Check "可読性_既に【】なら二重にしない", _
        (NA("**【重要】**です") = "【重要】です"), "実際=" & NA("**【重要】**です")
    modTestRunner.Check "可読性_二重防止の後も次の強調を変換", _
        (NA("**【A】**と**B**") = "【A】と【B】"), "実際=" & NA("**【A】**と**B**")

    ' 見出しの前は必ず1行空ける(直前の本文とくっつくと見出しに見えない)。
    Dim h1 As String: h1 = "本文" & vbLf & "## 次の話"
    modTestRunner.Check "可読性_見出しの前に空行を入れる", _
        (NA(h1) = "本文" & vbLf & vbLf & "■ 次の話"), "実際=" & Replace(NA(h1), vbLf, "<LF>")
    ' 先頭が見出しのときは、頭に空行を作らない。
    Dim h2 As String: h2 = "## 冒頭" & vbLf & "本文"
    modTestRunner.Check "可読性_先頭の見出しの前は空けない", _
        (NA(h2) = "■ 冒頭" & vbLf & "本文"), "実際=" & Replace(NA(h2), vbLf, "<LF>")

    ' 改行コードの混在(CRLF/CR)でも同じ結果になること。
    modTestRunner.Check "可読性_CRLFもLFへ揃える", (NA("A" & vbCrLf & "B") = "A" & vbLf & "B")
    modTestRunner.Check "可読性_空文字は空文字", (NA("") = "")

    ' バブルへ書く直前の形: 段落区切りは vbCr(Shapeはこれでしか段落が分かれない)。
    Dim para As String
    para = modLive.AnswerParagraphs("## 見出し" & vbLf & "本文")
    modTestRunner.Check "可読性_バブル本文はCR区切り", (InStr(para, vbCr) > 0)
    modTestRunner.Check "可読性_バブル本文にLFを残さない", (InStr(para, vbLf) = 0)
    modTestRunner.Check "可読性_バブル本文も■へ変換済み", (InStr(para, "■ 見出し") > 0), "実際=" & para
End Sub

' この節だけで20回以上呼ぶための短い別名(テスト限定)。
Private Function NA(ByVal s As String) As String
    NA = modLive.NormalizeAnswerText(s)
End Function

' R14-G5: 質問例は "|" 区切りで1本の文字列に畳んでボタンへ配る。質問文
'   そのものに "|" が入っていると1問が2つのボタンへ割れ、後半は文の途中から
'   始まる意味不明なボタンになる。区切りに使う以上、要素側からは必ず落とす。
Private Sub TestQuestionLinePipeStrip()
    Dim r As String
    r = modRagParse.ParseQuestionLines("A|B は何ですか" & vbLf & "次の質問", 5)
    modTestRunner.Check "質問パース_本文の縦棒は区切りにしない", _
        (r = "A" & ChrW(&HFF0F) & "B は何ですか|次の質問"), "実際=" & r
    modTestRunner.Check "質問パース_縦棒を含んでも2件のまま", _
        (UBound(Split(r, "|")) = 1), "実際=" & r

    ' 縦棒だけの行は捨てずに全角へ置換して残る(空行判定と混同しない)。
    r = modRagParse.ParseQuestionLines("|" & vbLf & "質問1", 5)
    modTestRunner.Check "質問パース_縦棒だけの行も1件として扱う", _
        (UBound(Split(r, "|")) = 1), "実際=" & r
End Sub

Private Sub TestHumanizeKeepsStageNumber()
    Dim r As String
    r = modLive.Humanize("(3/6) 資料の要点を整理中…")
    modTestRunner.Check "実況_番号を残す", (Left$(r, 6) = "(3/6) "), "実際=" & r
    modTestRunner.Check "実況_要点整理は言い換える", (InStr(r, "要点") > 0), "実際=" & r

    r = modLive.Humanize("(5/6) 下書きを自己点検中…")
    modTestRunner.Check "実況_自己点検の番号を残す", (Left$(r, 6) = "(5/6) "), "実際=" & r
    ' 「下書きを自己点検中」が「下書き」の分岐に食われないこと(語順の罠)。
    modTestRunner.Check "実況_自己点検は下書きの言い換えに食われない", _
        (InStr(r, "点検") > 0), "実際=" & r

    r = modLive.Humanize("(4/6) 下書きを作成中…")
    modTestRunner.Check "実況_下書きの番号を残す", (Left$(r, 6) = "(4/6) "), "実際=" & r

    ' 番号が付いていない実況(すぐ聞く)は従来どおり素の言い換え。
    r = modLive.Humanize("検索中…")
    modTestRunner.Check "実況_番号なしはそのまま言い換える", _
        (Left$(r, 1) <> "(" And InStr(r, "本棚") > 0), "実際=" & r
End Sub

' ----------------------------------------------------------------------------
' R14-F6: gs_abs_timeout_sec は【1資料あたり】の絶対上限。バッチごとに
'   残りだけを渡す(バッチ数×上限まで待てていたのが元の姿)。
' ----------------------------------------------------------------------------
Private Sub TestRemainingWaitSec()
    modTestRunner.Check "残り待ち_未使用なら全額", _
        (optOcrEta.RemainingWaitSec(1200, 0) = 1200), _
        "実際=" & optOcrEta.RemainingWaitSec(1200, 0)
    modTestRunner.Check "残り待ち_使ったぶんだけ減る", _
        (optOcrEta.RemainingWaitSec(1200, 500) = 700), _
        "実際=" & optOcrEta.RemainingWaitSec(1200, 500)
    modTestRunner.Check "残り待ち_使い切ったら0(=打ち切り)", _
        (optOcrEta.RemainingWaitSec(1200, 1200) = 0), _
        "実際=" & optOcrEta.RemainingWaitSec(1200, 1200)
    modTestRunner.Check "残り待ち_超過しても0", _
        (optOcrEta.RemainingWaitSec(1200, 1500) = 0), _
        "実際=" & optOcrEta.RemainingWaitSec(1200, 1500)
    modTestRunner.Check "残り待ち_残りが僅かでも10秒は待つ", _
        (optOcrEta.RemainingWaitSec(1200, 1197) = 10), _
        "実際=" & optOcrEta.RemainingWaitSec(1200, 1197)
    modTestRunner.Check "残り待ち_壊れた上限は0", _
        (optOcrEta.RemainingWaitSec(0, 0) = 0)

    ' 20頁バッチを6回まわしても、合計の待ちは絶対上限を超えない。
    Dim used As Long: used = 0
    Dim total As Long: total = 0
    Dim k As Long
    For k = 1 To 6
        Dim w As Long: w = optOcrEta.RemainingWaitSec(300, used)
        If w <= 0 Then Exit For
        total = total + w
        used = used + w          ' 最悪ケース(毎回待ち切る)
    Next k
    modTestRunner.Check "残り待ち_合計は絶対上限を超えない", (total <= 300), _
        "合計=" & total
End Sub

' ----------------------------------------------------------------------------
' R15-1c(実機第4報 RC5)+R15波1b裁定c: 再入ガードの失効判定
'   modUtilText.GuardExpired。従来は「開始から30分」で自動失効させていたため、
'   OCR付きの取込(実機で85〜127分)の途中で全ボタンが解放され二重取込が
'   構造的に起こり得た。「最後のビートから30分」へ変えたので、(a)開始が古くても
'   ビートが新しければ生存 (b)双方が古ければ失効(焼き付きは自己回復する)の
'   両方を同時に満たすことを固定する。片方だけの実装(常に生存/常に失効)は
'   どちらも実害が大きい。絶対上限(既定480分)の境界と、R15-FixA FA-9 で
'   取込ガードが渡すようになった absLimitMin=0(=無効)も合わせて固定する。
' ----------------------------------------------------------------------------
Private Sub TestGuardExpiredHeartbeat()
    Dim nowAt As Date
    nowAt = DateSerial(2026, 8, 4) + TimeSerial(10, 0, 0)

    ' 開始は90分前=旧実装なら失効。ビートが5分前なので生きている。
    modTestRunner.Check "ガード失効_古い開始でもビートが新しければ生存", _
        (modUtilText.GuardExpired(nowAt - TimeSerial(1, 30, 0), _
                                  nowAt - TimeSerial(0, 5, 0), nowAt, 30) = False)
    ' 双方古い(開始90分前・ビート40分前)=無音が続いた焼き付き。失効させる。
    modTestRunner.Check "ガード失効_開始もビートも古ければ失効", _
        (modUtilText.GuardExpired(nowAt - TimeSerial(1, 30, 0), _
                                  nowAt - TimeSerial(0, 40, 0), nowAt, 30) = True)

    ' ビート未打鍵(0)は開始時刻だけで判定する(取込開始直後の窓)。
    modTestRunner.Check "ガード失効_ビート無しで10分経過は生存", _
        (modUtilText.GuardExpired(nowAt - TimeSerial(0, 10, 0), 0, nowAt, 30) = False)
    modTestRunner.Check "ガード失効_ビート無しで31分経過は失効", _
        (modUtilText.GuardExpired(nowAt - TimeSerial(0, 31, 0), 0, nowAt, 30) = True)

    ' 境界: ちょうど30分は失効側(旧実装の >= と同じ向きを保つ)。
    modTestRunner.Check "ガード失効_ちょうど30分は失効", _
        (modUtilText.GuardExpired(nowAt - TimeSerial(0, 30, 0), 0, nowAt, 30) = True)
    modTestRunner.Check "ガード失効_29分は生存", _
        (modUtilText.GuardExpired(nowAt - TimeSerial(0, 29, 0), 0, nowAt, 30) = False)

    ' 前の取込が残した【古いビート】で今の判定を狂わせない(基準は新しい方)。
    modTestRunner.Check "ガード失効_開始より古いビートは延命に使わない", _
        (modUtilText.GuardExpired(nowAt - TimeSerial(0, 40, 0), _
                                  nowAt - TimeSerial(1, 40, 0), nowAt, 30) = True)
    ' 時刻不明(0)のガードは失効扱い=奪い返す(modUiLock.LockExpiredと同じ判断)。
    modTestRunner.Check "ガード失効_開始時刻が不明なら失効", _
        (modUtilText.GuardExpired(0, 0, nowAt, 30) = True)
    ' 時計が巻き戻った端末では失効させない(旧式と同じ挙動)。
    modTestRunner.Check "ガード失効_時計巻き戻しでは失効しない", _
        (modUtilText.GuardExpired(nowAt + TimeSerial(1, 0, 0), 0, nowAt, 30) = False)

    ' 上限値そのものは呼び出し側の定数。60分運用でも同じ式で効くこと。
    modTestRunner.Check "ガード失効_上限60分なら45分は生存", _
        (modUtilText.GuardExpired(nowAt - TimeSerial(0, 45, 0), 0, nowAt, 60) = False)

    ' R15波1b(裁定c): 絶対上限480分(8時間)。BlockIfIngesting自身の実況が
    ' ビートを打ってしまうと(a対応で塞いだが)、ビートがどれだけ新しくても
    ' 開始からabsLimitMinを超えたら無条件で失効させる最終防衛線。
    modTestRunner.Check "ガード失効_ビート新しくても開始8時間超は失効", _
        (modUtilText.GuardExpired(nowAt - TimeSerial(8, 10, 0), _
                                  nowAt - TimeSerial(0, 1, 0), nowAt, 30) = True)
    ' 境界: ちょうど480分は失効側(既定値の境界も >= と同じ向きを保つ)。
    modTestRunner.Check "ガード失効_ビート新しくても開始ちょうど480分は失効", _
        (modUtilText.GuardExpired(nowAt - TimeSerial(8, 0, 0), _
                                  nowAt - TimeSerial(0, 1, 0), nowAt, 30) = True)
    ' 479分・ビート新しければ絶対上限にはまだ触れず生存する。
    modTestRunner.Check "ガード失効_ビート新しく開始479分は生存", _
        (modUtilText.GuardExpired(nowAt - TimeSerial(7, 59, 0), _
                                  nowAt - TimeSerial(0, 1, 0), nowAt, 30) = False)
    ' absLimitMinを明示指定した運用でも同じ式で効くこと(既定値だけの偶然でない)。
    modTestRunner.Check "ガード失効_absLimitMin指定運用でも境界どおり", _
        (modUtilText.GuardExpired(nowAt - TimeSerial(2, 1, 0), _
                                  nowAt - TimeSerial(0, 1, 0), nowAt, 30, 120) = True)

    ' R15-FixA(FA-9・レビューA-M2): absLimitMin=0 は【無効】。8時間を超える
    ' 取込は正当に起こり得る(254頁の資料を何件もまとめて選ぶ)。そこでガードが
    ' 解けると、動いている取込の上に二重取込を開く入口ができる。ビートを打つのは
    ' 実作業だけになった(FA-8)ので、絶対上限という最終防衛線はもう要らない。
    ' 0 を「即失効」と解釈する実装(DateDiff >= 0 が常にTrue)になっていたら
    ' 全てのガードが常時失効=最悪の壊れ方をするので、ここで必ず止める。
    modTestRunner.Check "ガード失効_absLimitMin0はビート新しければ8時間超でも生存", _
        (modUtilText.GuardExpired(nowAt - TimeSerial(9, 0, 0), _
                                  nowAt - TimeSerial(0, 1, 0), nowAt, 30, 0) = False)
    modTestRunner.Check "ガード失効_absLimitMin0でも無音30分は従来どおり失効", _
        (modUtilText.GuardExpired(nowAt - TimeSerial(9, 0, 0), _
                                  nowAt - TimeSerial(0, 31, 0), nowAt, 30, 0) = True)
    modTestRunner.Check "ガード失効_absLimitMin0で開始直後は生存(常時失効ではない)", _
        (modUtilText.GuardExpired(nowAt - TimeSerial(0, 1, 0), 0, nowAt, 30, 0) = False)
    modTestRunner.Check "ガード失効_absLimitMin負も0と同じく無効", _
        (modUtilText.GuardExpired(nowAt - TimeSerial(9, 0, 0), _
                                  nowAt - TimeSerial(0, 1, 0), nowAt, 30, -1) = False)
End Sub

' ----------------------------------------------------------------------------
' R15-3c(実機第4報 RC9): 保存まわりの文言。どちらも「起きたこと」と
'   「次に打てる一手」を必ず含む(憲章§3-1: 無反応・無説明は故障と同じ)。
'   E0805 は今回から使い始めるコードで、コード表と起動時の案内が
'   1つの文言関数を共有していること(二重管理にしないこと)も固定する。
' ----------------------------------------------------------------------------
Private Sub TestSaveAndReadOnlyMessages()
    Dim ro As String: ro = modLog.ReadOnlyWarnMsg()
    modTestRunner.Check "読み取り専用文言_状態を言う", (InStr(ro, "読み取り専用") > 0), "実際=" & ro
    modTestRunner.Check "読み取り専用文言_保存されないと言う", _
        (InStr(ro, "保存されません") > 0), "実際=" & ro
    modTestRunner.Check "読み取り専用文言_確かめ方を言う", _
        (InStr(ro, "別のExcel") > 0), "実際=" & ro
    modTestRunner.Check "読み取り専用文言_E0805のコード表と同一", _
        (modLog.FriendlyMessage("E0805") = ro)
    ' 未知コードの汎用文言(「予期しない問題」)へ落ちていないこと。
    modTestRunner.Check "読み取り専用文言_汎用文言に落ちていない", _
        (InStr(modLog.FriendlyMessage("E0805"), "予期しない問題") = 0)

    Dim sf As String: sf = modLog.SaveFailMsg()
    modTestRunner.Check "保存失敗文言_起きたことを言う", _
        (InStr(sf, "保存に失敗しました") > 0), "実際=" & sf
    modTestRunner.Check "保存失敗文言_次の一手を言う", _
        (InStr(sf, "終了ボタン") > 0), "実際=" & sf
End Sub

' ----------------------------------------------------------------------------
' R15-2d(実機第4報 RC8): E0202の案内。実体は「AI通信の不調」ではなく
'   【実行中の処理と操作が重なったこと】であることが多い。AIの話だけで
'   終わらせると、利用者は待っても直らない別の問題として受け取り、
'   同じ操作を繰り返す(それがまた重なって E0202 を増やす)。
' ----------------------------------------------------------------------------
Private Sub TestE0202FriendlyMentionsOverlap()
    Dim m As String: m = modLog.FriendlyMessage("E0202")
    modTestRunner.Check "E0202文言_重なりに言及する", (InStr(m, "重なった") > 0), "実際=" & m
    modTestRunner.Check "E0202文言_少し待つよう案内する", _
        (InStr(m, "少し待って") > 0), "実際=" & m
    modTestRunner.Check "E0202文言_利用者を責めない", _
        (InStr(m, "あなたの操作に問題はありません") > 0), "実際=" & m
End Sub

Public Sub RunAll12()
    On Error GoTo ModeDescFail
    TestModeDescriptionsR14
    TestRerankEffortThorough
NextStage:
    On Error GoTo StageFail
    TestAskStageThorough
NextCite:
    On Error GoTo CiteFail
    TestCiteTagParsing
    TestAnnotateCitations
    TestCiteTagWithBracketName
    TestCiteIndexFromHits
NextRead:
    On Error GoTo ReadFail
    TestNormalizeAnswerText
    TestQuestionLinePipeStrip
    TestHumanizeKeepsStageNumber
NextRemainWait:
    On Error GoTo RemainWaitFail
    TestRemainingWaitSec
NextGuard15:
    On Error GoTo Guard15Fail
    TestGuardExpiredHeartbeat
NextSaveMsg15:
    On Error GoTo SaveMsg15Fail
    TestSaveAndReadOnlyMessages
    TestE0202FriendlyMentionsOverlap
NextPure13:
    ' R15波2: 分割先の入口。この1行が導線の全て(消すと modTestsPure13 の
    ' テストは「実行されないまま」全部PASSに見える)。
    On Error GoTo Pure13Fail
    modTestsPure13.RunAll13
NextDone12:
    On Error GoTo 0
    Exit Sub

ModeDescFail:
    modTestRunner.Check "TestModeDescriptionsR14(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextStage
StageFail:
    modTestRunner.Check "TestAskStageThorough(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextCite
CiteFail:
    modTestRunner.Check "TestCiteTagParsing(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRead
ReadFail:
    modTestRunner.Check "TestNormalizeAnswerText(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRemainWait
RemainWaitFail:
    modTestRunner.Check "TestRemainingWaitSec(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGuard15
Guard15Fail:
    modTestRunner.Check "TestGuardExpiredHeartbeat(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextSaveMsg15
SaveMsg15Fail:
    modTestRunner.Check "TestSaveAndReadOnlyMessages(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextPure13
Pure13Fail:
    modTestRunner.Check "modTestsPure13.RunAll13(モジュール全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone12
End Sub
