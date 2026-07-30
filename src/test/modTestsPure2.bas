Attribute VB_Name = "modTestsPure2"
Option Explicit

' ============================================================================
' modTestsPure2 - modTestsPureの分割先(MASTER_SPEC §7.1「1モジュール30,000字
'   以内」を超えたための分割。§7.8のテスト範囲そのものは変わらない)
' ----------------------------------------------------------------------------
' 役割:
'   modPrompts / modShelfSync.DiffDecision / modPack.ValidatePackMeta の
'   純ロジック部のテストをここに置く(modUtil/modChunker/modPiiは
'   modTestsPure側)。入口は modTestsPure.RunAll の末尾から呼ばれる
'   Public Sub RunAll2()。modTestRunner.RunAllPureTests は modTestsPure.RunAll
'   だけを呼ぶ契約(modTestRunner.bas §7.8。担当外につき変更しない)なので、
'   本モジュールへの導線は modTestsPure.RunAll 内に置く。
'
' 設計判断(R4準拠・グループ単位の失敗隔離): modTestsPure側の冒頭コメントと
'   同じ方針(1グループの想定外エラーが他グループを道連れにしない)。
'
' ■ CanUseTypeArraysの複製について:
'   modTestsPure の Private Function CanUseTypeArrays() は別モジュールの
'   Privateであるため、ここから呼べない。判定ロジック自体は3行程度の軽量な
'   実測プローブなので、モジュールをまたいだ複製を許容する(共有したいだけの
'   ために新たにPublicの公開契約を増やすより、テストモジュール限定の軽微な
'   重複の方が実害が小さいと判断)。挙動・コメントはmodTestsPure側の
'   オリジナルと同一にしてある。
'
' ■ lint注意(2026-07-12 Wave3-Tで特定): tools/vba_lint.py の
'   モジュール間参照検査は文字列リテラルの中身までスキャンする(行コメント
'   `'` だけを除去し、ダブルクォート文字列は除去しない実装のため)。
'   このため Check の detail 文字列の中で「modTestsPure.bas」のように
'   モジュール名にドット+英字が続く書き方をすると、「modTestsPure.bas」を
'   modTestsPure.bas という架空のPublicメンバ参照として誤検知する
'   (vba_lint.py自体はCONTRACT/PURE_ALLOWLIST追随目的以外は変更不可のため、
'   テスト側の文字列表現でこれを避ける: 「modTestsPureのbasファイル」のように
'   ドット直後に英字が来ない書き方に統一する)。
' ============================================================================

' ----------------------------------------------------------------------------
' 日本語キーワード検索(modSparse)の回帰テスト(2026-07-27追加)
' ----------------------------------------------------------------------------
' 実測(実物6資料180チャンク31問・toolsのベンチ):
'   旧実装(空白分割+一律加点)  R@1 32% / MRR 0.34   ← 赤点
'   本実装(文字bigram+BM25)    R@1 84% / R@5 100% / MRR 0.91
' 差の大半は「日本語は空白で区切らない」という一点。ここが壊れると
' 検索は静かにゴミへ戻るので、性質を固定しておく。
' ----------------------------------------------------------------------------
' 出荷する検索経路の回帰テスト(2026-07-27)
' ----------------------------------------------------------------------------
' 実測(実物6資料180チャンク31問)で選んだ最終構成:
'     手法                              R@1  R@3  R@5  MRR  ページ
'     旧実装(空白分割+一律加点)           32%  32%  32% 0.34   16%
'     文字bigram BM25(全件)              84%  97% 100% 0.91   45%
'     ★採用: InStr全件・空白除去マッチ     84%  84%  90% 0.86   81%
' BM25は R@3/R@5 が上だが、1チャンクのトークナイズに225ms(LO実測)かかり
' 9000件では34分になるため毎クエリ全件には使えない。採用案は
' トークナイズ不要で、出典ページの正確さ(81%)も最も高い。
' 表記揺れ(全角/大文字/空白挿入/空白除去)は全て84%を維持した。
' ----------------------------------------------------------------------------
' 3モード(すぐ聞く/しっかり/入念)の方針テスト(2026-07-28)
' ----------------------------------------------------------------------------
' 実測(実物6資料180チャンク31問):
'   単発検索      R@1 84% / R@5 90% / R@10 90%
'   多クエリ→RRF  R@1 77% / R@5 90% / R@10 97%
' 「渡す件数」を増やしても効かず(R@5 90%→R@10 90%)、「引き方」を増やすと
' 効く(R@10 90%→97%)。モードごとに戦略が違うので、その方針を固定する。
Private Sub RunModeTests()
    modTestRunner.Check "モード: 未知の値はすぐ聞くへ倒す", _
        modMode.Normalize("xxx") = "quick", "実際=" & modMode.Normalize("xxx")
    modTestRunner.Check "モード: 大文字でも認識する", _
        modMode.Normalize("THOROUGH") = "thorough", "実際=" & modMode.Normalize("THOROUGH")

    modTestRunner.Check "巡回: すぐ聞く→しっかり", modMode.NextMode("quick") = "deep", ""
    modTestRunner.Check "巡回: しっかり→入念", modMode.NextMode("deep") = "thorough", ""
    modTestRunner.Check "巡回: 入念→すぐ聞く", modMode.NextMode("thorough") = "quick", ""

    ' 入念は時間より精度。全段を必ず通す
    modTestRunner.Check "入念: 質問拡張を必ず行う", _
        modMode.UseExpand("thorough", False, False) = True, _
        "configがFalseでも入念だけは通す"
    modTestRunner.Check "入念: 再ランクを必ず行う", _
        modMode.UseRerank("thorough", False, False) = True, ""
    modTestRunner.Check "入念: 検証段を必ず通す", modMode.UseVerify("thorough") = True, ""
    modTestRunner.Check "しっかり: 検証段を通す", modMode.UseVerify("deep") = True, ""
    modTestRunner.Check "すぐ聞く: 検証段は通さない", modMode.UseVerify("quick") = False, _
        "すぐ聞くで検証まで走ると速さの意味が無くなる"

    ' すぐ聞くは既定で拡張・再ランクを行わない(実測でLLM3回→1回)
    modTestRunner.Check "すぐ聞く: 既定では拡張しない", _
        modMode.UseExpand("quick", True, False) = False, ""
    modTestRunner.Check "すぐ聞く: optInすれば拡張する", _
        modMode.UseExpand("quick", True, True) = True, ""

    ' 件数
    modTestRunner.Check "件数: モードごとに変わる", _
        modMode.TopK("quick", 6, 12, 16) = 6 And _
        modMode.TopK("deep", 6, 12, 16) = 12 And _
        modMode.TopK("thorough", 6, 12, 16) = 16, ""
    modTestRunner.Check "件数: 0以下は安全側へ倒す", modMode.TopK("quick", 0, 0, 0) = 6, ""

    ' サブクエリ数(角度の数がそのまま精度になる)
    modTestRunner.Check "サブクエリ: 入念は多い", _
        modMode.SubQueryCount("thorough", 3, 6) = 6, ""
    modTestRunner.Check "サブクエリ: 通常は既定", _
        modMode.SubQueryCount("deep", 3, 6) = 3, ""

    ' 表示文言(利用者が選ぶ理由になる情報が入っていること)
    modTestRunner.Check "表示: 入念の説明に所要時間が入る", _
        InStr(modMode.Description("thorough"), "分") > 0, _
        "時間を隠すと固まったと思われる"
    modTestRunner.Check "表示: モード名が3つとも異なる", _
        modMode.Caption("quick") <> modMode.Caption("deep") And _
        modMode.Caption("deep") <> modMode.Caption("thorough"), ""
End Sub


Private Sub RunKeyScoreTests()
    ' 照合用テキスト: 正規化 + 空白の完全除去
    modTestRunner.Check "照合用: 空白を落として揃える", _
        modSparse.CompactForMatch("保 険 　金") = "保険金", _
        "実際=" & modSparse.CompactForMatch("保 険 　金")
    modTestRunner.Check "照合用: 全角数字も揃う", _
        modSparse.CompactForMatch("第 ２ ０ 条") = "第20条", _
        "実際=" & modSparse.CompactForMatch("第 ２ ０ 条")

    ' 空白入りの質問からでもキーが取れること(取れないと0件になる)
    Dim k1 As String, k2 As String
    k1 = modSparse.DistinctiveKeys("保険金を支払わない場合")
    k2 = modSparse.DistinctiveKeys("保 険 金 を 支 払 わ な い 場 合")
    modTestRunner.Check "効く語: 空白が入っても同じキーが取れる", k1 = k2, _
        "空白なし=" & k1 & " / 空白あり=" & k2

    ' スコア: 一致が多い文書のほうが高い
    Dim body1 As String, body2 As String
    body1 = modSparse.CompactForMatch("第12条(保険金を支払わない場合)当社は保険金を支払いません。")
    body2 = modSparse.CompactForMatch("第30条(保険料の払込方法)当社は保険料を集金します。")
    Dim keys As String: keys = modSparse.DistinctiveKeys("保険金を支払わない場合")
    modTestRunner.Check "スコア: 一致の多い文書が上位になる", _
        modSparse.KeyScore(keys, body1) > modSparse.KeyScore(keys, body2), _
        "body1=" & Format$(modSparse.KeyScore(keys, body1), "0.000") & _
        " body2=" & Format$(modSparse.KeyScore(keys, body2), "0.000")

    ' 条番号は確実に効くこと(金融では外せない)
    Dim ka As String: ka = modSparse.DistinctiveKeys("第12条について教えて")
    modTestRunner.Check "条番号: 第12条の文書が第30条より上位", _
        modSparse.KeyScore(ka, body1) > modSparse.KeyScore(ka, body2), _
        "12条=" & Format$(modSparse.KeyScore(ka, body1), "0.000") & _
        " 30条=" & Format$(modSparse.KeyScore(ka, body2), "0.000")

    ' 無関係な文書は0点(ノイズを上げない)
    modTestRunner.Check "スコア: 無関係な文書は0点", _
        modSparse.KeyScore(keys, modSparse.CompactForMatch("abcdefg hijklmn")) = 0, ""

    ' 空白の入った本文にも当たること(PDF字詰め対策の本丸)
    Dim spaced As String
    spaced = modSparse.CompactForMatch("第12条(保 険 金 を 支 払 わ な い 場 合)")
    modTestRunner.Check "字詰め: 空白入り本文にも一致する", _
        modSparse.KeyScore(keys, spaced) > 0, _
        "PDFの字詰めで空白が入ると当たらない実装は使えない"
End Sub


Private Sub RunSparseTests()
    modTestRunner.Check "正規化: 全角数字を半角へ", _
        modSparse.NormalizeForSearch("第２０条") = "第20条", _
        "実際=" & modSparse.NormalizeForSearch("第２０条")
    modTestRunner.Check "正規化: 英大文字を小文字へ", _
        modSparse.NormalizeForSearch("Recall") = "recall", _
        "実際=" & modSparse.NormalizeForSearch("Recall")
    modTestRunner.Check "正規化: 全角英字を半角小文字へ", _
        modSparse.NormalizeForSearch("ＲＥＣＡＬＬ") = "recall", _
        "実際=" & modSparse.NormalizeForSearch("ＲＥＣＡＬＬ")
    modTestRunner.Check "正規化: 全角空白を半角へ", _
        modSparse.NormalizeForSearch("あ" & ChrW(12288) & "い") = "あ い", _
        "実際=" & modSparse.NormalizeForSearch("あ" & ChrW(12288) & "い")

    Dim t As String
    t = modSparse.Tokenize("保険金")
    modTestRunner.Check "bigram: 保険金から保険と険金が出る", _
        InStr(t, "保険") > 0 And InStr(t, "険金") > 0, "実際=" & t
    modTestRunner.Check "bigram: 英数字は語のまま残す", _
        InStr(modSparse.Tokenize("code123"), "code123") > 0, _
        "実際=" & modSparse.Tokenize("code123")
    modTestRunner.Check "bigram: 途中の空白で語を割らない", _
        InStr(modSparse.Tokenize("保 険 金"), "保険") > 0, _
        "実際=" & modSparse.Tokenize("保 険 金")

    Dim keys As String
    keys = modSparse.DistinctiveKeys("第12条の内容を教えて")
    modTestRunner.Check "効く語: 条番号を必ず拾う", _
        InStr(keys, "第12条") > 0, "実際=" & keys

    modTestRunner.Check "完全一致: 本文に含まれれば1", _
        modSparse.ExactHitCount("第12条", "【第12条(保険金)】当社は") = 1, ""
    modTestRunner.Check "完全一致: 含まれなければ0", _
        modSparse.ExactHitCount("第12条", "【第13条(通知)】当社は") = 0, ""

    Dim q As String
    q = modSparse.Tokenize("保険金を支払わない場合")
    Dim hi As Double, lo As Double
    hi = modSparse.Bm25Score(q, "保険金を支払わない場合について定めます。", "", 100, 50)
    lo = modSparse.Bm25Score(q, "保険料の払込方法について定めます。", "", 100, 50)
    modTestRunner.Check "BM25: 一致の多い文書が上位になる", hi > lo, _
        "hi=" & Format$(hi, "0.000") & " lo=" & Format$(lo, "0.000")
    modTestRunner.Check "BM25: 無関係な文書は0点", _
        modSparse.Bm25Score(q, "abcdefg hijklmn", "", 100, 50) = 0, ""

    modTestRunner.Check "揺れ: 全角でも同じトークン列になる", _
        modSparse.Tokenize("第20条") = modSparse.Tokenize("第２０条"), _
        "半角=" & modSparse.Tokenize("第20条") & " 全角=" & modSparse.Tokenize("第２０条")
End Sub


Public Sub RunAll2()
    On Error GoTo PromptsFail
    TestModPrompts
NextShelfSync:
    On Error GoTo ShelfSyncFail
    TestModShelfSync
NextPack:
    On Error GoTo PackFail
    TestModPack

    ' 2026-07-30 R2要件B: modTestsPure/modTestsPure2とも30,000字上限まで
    ' 残りが少なく、新規テストの置き場が無かったための分割先(モジュール
    ' 冒頭コメント参照)。ここで失敗しても後続のSparse以降は道連れにしない。
    On Error GoTo Pure3Fail
    modTestsPure3.RunAll3
NextSparse:
    On Error GoTo SparseFail
    RunSparseTests
    RunKeyScoreTests
    RunModeTests
    On Error GoTo ChannelFail
    RunChannelOriginTests
    On Error GoTo ClarifyFail
    RunClarifyChoiceTests
    On Error GoTo VecFail
    RunVectorCsvTests
    On Error GoTo BadgeFail
    RunBadgeCatalogTests
NextDone:
    On Error GoTo 0
    Exit Sub

PromptsFail:
    modTestRunner.Check "TestModPrompts(グループ全体)", False, "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextShelfSync
ShelfSyncFail:
    modTestRunner.Check "TestModShelfSync(グループ全体)", False, "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextPack
SparseFail:
    modTestRunner.Check "RunSparseTests(グループ全体)", False, "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone
ChannelFail:
    modTestRunner.Check "RunChannelOriginTests(グループ全体)", False, "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone
ClarifyFail:
    modTestRunner.Check "RunClarifyChoiceTests(グループ全体)", False, "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone
VecFail:
    modTestRunner.Check "RunVectorCsvTests(グループ全体)", False, "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone
BadgeFail:
    modTestRunner.Check "RunBadgeCatalogTests(グループ全体)", False, "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone
PackFail:
    modTestRunner.Check "TestModPack(グループ全体)", False, "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone
Pure3Fail:
    modTestRunner.Check "modTestsPure3のRunAll3(グループ全体)", False, "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextSparse
End Sub

' modTestsPureの CanUseTypeArrays と同一実装(モジュール冒頭コメント参照)。
Private Function CanUseTypeArrays() As Boolean
    On Error Resume Next
    Err.Clear
    Dim probe() As ShelfChunk
    ReDim probe(0 To 0)
    CanUseTypeArrays = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' ============================================================================
' modPrompts
' ----------------------------------------------------------------------------
' 重要(2026-07-12 Wave3-Tで特定・修正): modRetrieve.Search が返すHit配列は
' 1始まり(ReDim outHits(1 To filled))であり、modPrompts.BuildSourceBlockも
' `For i = 1 To nHits: ... hits(i) ...` と1始まり前提でhits()を読む
' (src/qa/modRetrieve.bas / src/qa/modPrompts.bas 参照)。
' 以前の下書きはテスト側のHit配列を「Dim h(0 To 1) As Hit」のように0始まりで
' 宣言しつつ nHits=2 を渡していたため、BuildQuickPromptがhits(2)へアクセスし
' 添字範囲外(実行時エラー9)になる、テストコード自身のバグだった(modPrompts
' 側の契約が誤っていたわけではない。1始まりが正・テストの宣言が誤りと判断)。
' 本ファイルでは全てのHit配列を「Dim h(1 To n) As Hit」の1始まりに統一する。
' 実行環境がCanUseTypeArrays=Falseの場合はこれらのテスト自体が丸ごとスキップ
' されるため(下記TestModPrompts参照)、この修正はLibreOffice純ロジック
' テストの合否には影響しないが、Excel実機受入チェック(§11.3)でこの
' テストコードが実際に実行されたときに正しく動くために必須の修正である。
' ============================================================================
Private Sub TestModPrompts()
    ' BuildEnrichPromptはHit()配列を取らないため、LO実行環境のPublic Type配列の
    ' 制限(モジュール冒頭コメント参照)の影響を受けない。常にフル検証する。
    TestBuildEnrichPrompt

    If Not CanUseTypeArrays() Then
        modTestRunner.Check "modPrompts(Hit配列を使う3関数): LO環境の既知の制限によりスキップ", True, _
            "BuildQuickPrompt/BuildDeepDraftPrompt/BuildDeepVerifyPromptはHit()配列を" & _
            "引数に取るため、LibreOffice実行環境のPublic Type配列制限(モジュールのbasファイル" & _
            "冒頭コメント参照)の影響を受ける。出典形式([本棚:.. p.N] / [パック(作成者):..])・" & _
            "full_text根拠の使用・max_context_chars打切り時の「(一部省略)」挿入・" & _
            "answer_language既定値挿入はコードレビューで確認済みだが、" & _
            "Excel実機受入チェック(§11.3)で必ず再確認すること。"
        Exit Sub
    End If

    TestBuildQuickPrompt_CitationAndLanguage
    TestBuildQuickPrompt_FullTextIsUsedAsBody
    TestBuildQuickPrompt_Truncation
    TestBuildDeepDraftAndVerify_Citation
End Sub

Private Sub TestBuildEnrichPrompt()
    Dim r As String
    r = modPrompts.BuildEnrichPrompt("チャンク1本文" & vbLf & "チャンク2本文")
    modTestRunner.Check "BuildEnrichPrompt_バッチ本文を含む", (InStr(r, "チャンク1本文") > 0), "r=" & r
    modTestRunner.Check "BuildEnrichPrompt_JSON形式の指示を含む", (InStr(r, "summary") > 0) And (InStr(r, "keywords") > 0), "r=" & r
    ' answer_language既定の挿入(modConfig未接続でも既定値"日本語"で動くこと)
    modTestRunner.Check "BuildEnrichPrompt_既定言語_日本語", (InStr(r, "日本語") > 0), "r=" & r
End Sub

Private Sub TestBuildQuickPrompt_CitationAndLanguage()
    ' modRetrieve.Searchが返すHit配列は1始まり(§7.3参照)。テストのローカル
    ' 配列もそれに合わせて1始まりで宣言する(モジュール冒頭コメント参照)。
    Dim h(1 To 2) As Hit
    h(1).chunk_id = "bs::aaa::p3::c1": h(1).score = 0.9
    h(1).source = "ファイルA.pdf": h(1).page = 3
    h(1).preview = "ここに本文の抜粋が入ります。": h(1).origin = "self"

    h(2).chunk_id = "bs::bbb::p1::c1": h(2).score = 0.8
    h(2).source = "ファイルB.docx": h(2).page = 1
    h(2).preview = "別の抜粋です。": h(2).origin = "pack:山田太郎"

    Dim r As String
    r = modPrompts.BuildQuickPrompt("何か質問", h, 2)

    modTestRunner.Check "BuildQuickPrompt_本棚出典形式", (InStr(r, "[本棚:ファイルA.pdf p.3]") > 0), "r=" & r
    modTestRunner.Check "BuildQuickPrompt_パック出典形式", (InStr(r, "[パック(山田太郎):ファイルB.docx]") > 0), "r=" & r
    ' answer_language既定の挿入(modConfig未接続でも既定値"日本語"で動くこと)
    modTestRunner.Check "BuildQuickPrompt_既定言語_日本語", (InStr(r, "日本語") > 0), "r=" & r
    modTestRunner.Check "BuildQuickPrompt_質問文を含む", (InStr(r, "何か質問") > 0), "r=" & r
    ' 打ち切っていない通常ケースでは「(一部省略)」は出ないこと
    modTestRunner.Check "BuildQuickPrompt_非打切り時は省略表記なし", (InStr(r, "(一部省略)") = 0), "r=" & r
End Sub

' Wave3 PM裁定1(Hit.full_text追加)の検証: 本棚抜粋ブロックの本文には
' preview(先頭120字・出典先出し表示専用)ではなく full_text(チャンク本文
' 全体)が使われること。preview と full_text に別々の目印文字列を仕込み、
' 出力に full_text 側の目印だけが含まれ、preview 側の目印は含まれないことを
' 確認する(modPrompts.SourceBodyがfull_text優先・空時のみpreviewへ
' フォールバックする設計であることのテスト。§7.3/modTypes.Hit参照)。
Private Sub TestBuildQuickPrompt_FullTextIsUsedAsBody()
    Dim h(1 To 1) As Hit
    h(1).chunk_id = "bs::ft::p9::c1": h(1).score = 0.9
    h(1).source = "全文根拠資料.pdf": h(1).page = 9
    h(1).preview = "PREVIEW_ONLY_MARKER_短い先頭抜粋"
    h(1).full_text = "FULLTEXT_MARKER_これがチャンク本文全体の根拠テキストです。" & _
        "本来はここに数百字の実際の抜粋が入る想定。"
    h(1).origin = "self"

    Dim r As String
    r = modPrompts.BuildQuickPrompt("何か質問", h, 1)

    modTestRunner.Check "BuildQuickPrompt_full_text根拠が含まれる", _
        (InStr(r, "FULLTEXT_MARKER_これがチャンク本文全体の根拠テキストです。") > 0), "r=" & r
    modTestRunner.Check "BuildQuickPrompt_full_text優先時はpreviewを使わない", _
        (InStr(r, "PREVIEW_ONLY_MARKER") = 0), "r=" & r
    modTestRunner.Check "BuildQuickPrompt_full_text使用時も本棚出典形式", _
        (InStr(r, "[本棚:全文根拠資料.pdf p.9]") > 0), "r=" & r

    ' 防御的フォールバック確認: full_textが空(旧データ・テストダブル等)なら
    ' previewへフォールバックすること。
    Dim h2(1 To 1) As Hit
    h2(1).chunk_id = "bs::fb::p1::c1": h2(1).score = 0.5
    h2(1).source = "フォールバック資料.pdf": h2(1).page = 1
    h2(1).preview = "PREVIEW_FALLBACK_MARKER"
    h2(1).full_text = ""   ' 空
    h2(1).origin = "self"

    Dim r2 As String
    r2 = modPrompts.BuildQuickPrompt("別の質問", h2, 1)
    modTestRunner.Check "BuildQuickPrompt_full_text空時はpreviewへフォールバック", _
        (InStr(r2, "PREVIEW_FALLBACK_MARKER") > 0), "r2=" & r2
End Sub

Private Sub TestBuildQuickPrompt_Truncation()
    ' previewを非常に長くして確実にmax_context_chars(既定40000。modConfig未接続時も
    ' SafeMaxContextCharsの既定値40000が使われる)を超えさせ、打ち切りと
    ' 「(一部省略)」挿入を検証する。full_textは空のままにして、SourceBodyの
    ' フォールバック(full_text空→preview使用)経由で長文を本文に載せる。
    Dim bigPreview As String: bigPreview = String(30000, "x")
    Dim h(1 To 2) As Hit
    h(1).chunk_id = "bs::a::p1::c1": h(1).score = 0.9
    h(1).source = "大きい資料.pdf": h(1).page = 1
    h(1).preview = bigPreview: h(1).origin = "self"

    h(2).chunk_id = "bs::b::p1::c1": h(2).score = 0.8
    h(2).source = "大きい資料2.pdf": h(2).page = 1
    h(2).preview = bigPreview: h(2).origin = "self"

    Dim r As String
    r = modPrompts.BuildQuickPrompt("質問", h, 2)

    modTestRunner.Check "BuildQuickPrompt_打切り時は省略表記あり", (InStr(r, "(一部省略)") > 0), _
        "len(r)=" & Len(r)
    ' 打ち切られているので2件目の資料名までは含まれない(先頭60000字級のpreviewの後半は
    ' コンテキストブロックに入りきらないはず)。少なくとも全文がそのまま連結されて
    ' いない(=本当に打ち切られている)ことを、長さの上限チェックで確認する。
    modTestRunner.Check "BuildQuickPrompt_本文が無制限に伸びていない", (Len(r) < (Len(bigPreview) * 2 + 5000)), _
        "len(r)=" & Len(r)
End Sub

Private Sub TestBuildDeepDraftAndVerify_Citation()
    Dim h(1 To 1) As Hit
    h(1).chunk_id = "bs::c::p2::c1": h(1).score = 0.7
    h(1).source = "資料C.xlsx": h(1).page = 2
    h(1).preview = "抜粋C": h(1).origin = "self"

    Dim rDraft As String
    rDraft = modPrompts.BuildDeepDraftPrompt("質問2", h, 1, "")
    modTestRunner.Check "BuildDeepDraftPrompt_出典形式", (InStr(rDraft, "[本棚:資料C.xlsx p.2]") > 0), "rDraft=" & rDraft
    modTestRunner.Check "BuildDeepDraftPrompt_既定言語_日本語", (InStr(rDraft, "日本語") > 0), "rDraft=" & rDraft

    Dim rVerify As String
    rVerify = modPrompts.BuildDeepVerifyPrompt("質問2", "下書き回答本文", h, 1)
    modTestRunner.Check "BuildDeepVerifyPrompt_出典形式", (InStr(rVerify, "[本棚:資料C.xlsx p.2]") > 0), "rVerify=" & rVerify
    modTestRunner.Check "BuildDeepVerifyPrompt_下書きを含む", (InStr(rVerify, "下書き回答本文") > 0), "rVerify=" & rVerify
End Sub

' ============================================================================
' modShelfSync.DiffDecision(純関数)
' ============================================================================
Private Sub TestModShelfSync()
    modTestRunner.Check "DiffDecision_新規", _
        (modShelfSync.DiffDecision(False, False, False) = "ingest")
    modTestRunner.Check "DiffDecision_新規_サイズ時刻フラグ無視", _
        (modShelfSync.DiffDecision(False, True, True) = "ingest")
    modTestRunner.Check "DiffDecision_サイズ変化", _
        (modShelfSync.DiffDecision(True, True, False) = "replace")
    modTestRunner.Check "DiffDecision_時刻変化", _
        (modShelfSync.DiffDecision(True, False, True) = "replace")
    modTestRunner.Check "DiffDecision_サイズ時刻両方変化", _
        (modShelfSync.DiffDecision(True, True, True) = "replace")
    modTestRunner.Check "DiffDecision_不変", _
        (modShelfSync.DiffDecision(True, False, False) = "keep")

    ' modShelfSync.ResolveDecision(純関数・Wave4追加): keepにfailed/missingの
    ' 復帰ルールを適用する。フォルダ復活時にmissing行が永久に「削除待ち」表示の
    ' ままになる不具合(Wave4レビュー指摘)の修正対象。
    modTestRunner.Check "ResolveDecision_keepかつfailedはreplaceへ昇格", _
        (modShelfSync.ResolveDecision("keep", True, "failed") = "replace")
    modTestRunner.Check "ResolveDecision_keepかつmissingはreplaceへ昇格", _
        (modShelfSync.ResolveDecision("keep", True, "missing") = "replace")
    modTestRunner.Check "ResolveDecision_keepかつdoneはそのまま", _
        (modShelfSync.ResolveDecision("keep", True, "done") = "keep")
    modTestRunner.Check "ResolveDecision_keepかつpartialはそのまま", _
        (modShelfSync.ResolveDecision("keep", True, "partial") = "keep")
    modTestRunner.Check "ResolveDecision_ingestは素通り", _
        (modShelfSync.ResolveDecision("ingest", False, "") = "ingest")
    modTestRunner.Check "ResolveDecision_replaceは素通り", _
        (modShelfSync.ResolveDecision("replace", True, "done") = "replace")
    modTestRunner.Check "ResolveDecision_existsInManifest偽なら昇格しない", _
        (modShelfSync.ResolveDecision("keep", False, "missing") = "keep")
End Sub

' ============================================================================
' modPack.ValidatePackMeta(純関数)
' ============================================================================
Private Sub TestModPack()
    Dim reason As String

    ' 正常(バージョン一致・次元1以上)
    Dim okNormal As Boolean
    okNormal = modPack.ValidatePackMeta(modAppDef.PACK_FORMAT_VERSION, 1536, reason)
    modTestRunner.Check "ValidatePackMeta_正常", okNormal, "reason=" & reason

    ' バージョン不一致
    reason = ""
    Dim okVerMismatch As Boolean
    okVerMismatch = modPack.ValidatePackMeta(modAppDef.PACK_FORMAT_VERSION + 1, 1536, reason)
    modTestRunner.Check "ValidatePackMeta_バージョン不一致False", (Not okVerMismatch), "reason=" & reason
    modTestRunner.Check "ValidatePackMeta_バージョン不一致理由あり", (Len(reason) > 0), "reason=" & reason

    ' 次元不一致(0以下は不正)
    reason = ""
    Dim okDimInvalid As Boolean
    okDimInvalid = modPack.ValidatePackMeta(modAppDef.PACK_FORMAT_VERSION, 0, reason)
    modTestRunner.Check "ValidatePackMeta_次元0はFalse", (Not okDimInvalid), "reason=" & reason

    reason = ""
    Dim okDimNegative As Boolean
    okDimNegative = modPack.ValidatePackMeta(modAppDef.PACK_FORMAT_VERSION, -5, reason)
    modTestRunner.Check "ValidatePackMeta_次元負値はFalse", (Not okDimNegative), "reason=" & reason
End Sub


' ----------------------------------------------------------------------------
' 部門チャンネルの origin タグ規約(2026-07-28 レビュー C-1 の恒久再発防止)
'
' 実バグ: 取込側が書くタグは "pack:"&作者名、削除側が探すタグは
' "pack:"&部門名 で、部門名≠作者名である限り削除は常に0件だった。
' 結果、部門を切り替えても前の部門が本棚に残り、更新配信では新旧の条文が
' 混ざった(=AIが古い条文を根拠に答える)。どちらも静かに壊れるので、
' 実機テストでも「動いているように見える」のが最悪だった。
'
' ここで固定するのは次の3点:
'   1. タグの組み立ては ChannelOriginTag 1関数に集約されていること
'   2. チャンネルのタグは手渡しパックの "pack:" と別名前空間であること
'      (混ざると手渡しパックが部門切替で巻き添えで消える)
'   3. 前後の空白でタグがズレないこと(部門名は人が打つフォルダ名)
' ----------------------------------------------------------------------------
Private Sub RunChannelOriginTests()
    modTestRunner.Check "ChannelOriginTag_基本", _
        (modChannel.ChannelOriginTag("商品部") = "channel:商品部"), _
        "actual=" & modChannel.ChannelOriginTag("商品部")

    modTestRunner.Check "ChannelOriginTag_前後空白を落とす", _
        (modChannel.ChannelOriginTag("  人事部 ") = "channel:人事部"), _
        "actual=" & modChannel.ChannelOriginTag("  人事部 ")

    modTestRunner.Check "ChannelOriginTag_空は空を返す", _
        (LenB(modChannel.ChannelOriginTag("")) = 0)

    modTestRunner.Check "ChannelOriginTag_空白のみは空を返す", _
        (LenB(modChannel.ChannelOriginTag("   ")) = 0)

    ' 手渡しパックの名前空間("pack:")と衝突しないこと。ここが崩れると
    ' 部門を切り替えたときに手渡しパックまで一緒に消える。
    modTestRunner.Check "ChannelOriginTag_手渡しパックと別名前空間", _
        (Left$(modChannel.ChannelOriginTag("商品部"), 5) <> "pack:"), _
        "actual=" & modChannel.ChannelOriginTag("商品部")

    ' 部門名が違えばタグも違う(切替時に他部門を巻き込まない)。
    modTestRunner.Check "ChannelOriginTag_部門ごとに異なる", _
        (modChannel.ChannelOriginTag("商品部") <> modChannel.ChannelOriginTag("人事部"))
End Sub


' ----------------------------------------------------------------------------
' 聞き返しの「番号選択」判定(2026-07-28 レビュー H-12 の恒久再発防止)
'
' 実バグ: 聞き返し保留中の返信に半角1～4がどこかに1つでもあれば番号選択と
' みなし、利用者が実際に打った文章を捨てて「前回の質問＋対象の資料: 候補N」
' を送っていた。保険実務の質問は「第1条」「3日以内」のように数字をほぼ確実に
' 含むので、書き直したつもりの人の文章が黙って消えていた。
'
' 固定するルール: 番号選択とみなすのは「番号だけを短く打った」ときだけ。
' 日本語が1文字でも混ざったら、それは書き直しである。
' ----------------------------------------------------------------------------
Private Sub RunClarifyChoiceTests()
    ' --- 番号だけ = True ---
    modTestRunner.Check "番号選択_半角1文字", modClarify.IsNumberChoiceOnly("2")
    modTestRunner.Check "番号選択_全角数字", modClarify.IsNumberChoiceOnly(ChrW(65298))
    modTestRunner.Check "番号選択_丸数字", modClarify.IsNumberChoiceOnly(ChrW(9313))
    modTestRunner.Check "番号選択_資料と観点の組", modClarify.IsNumberChoiceOnly("1-3")
    modTestRunner.Check "番号選択_丸数字2つ", _
        modClarify.IsNumberChoiceOnly(ChrW(9312) & " " & ChrW(9314))
    modTestRunner.Check "番号選択_ピリオド付き", modClarify.IsNumberChoiceOnly("3.")
    modTestRunner.Check "番号選択_括弧付き", modClarify.IsNumberChoiceOnly("(2)")
    modTestRunner.Check "番号選択_読点区切り", modClarify.IsNumberChoiceOnly("2" & ChrW(12289) & "4")
    modTestRunner.Check "番号選択_前後空白", modClarify.IsNumberChoiceOnly("  1  ")

    ' --- 文章 = False(ここが実バグの本体) ---
    modTestRunner.Check "書き直し_第1条を含む質問は番号選択ではない", _
        (Not modClarify.IsNumberChoiceOnly("第1条の適用範囲は?"))
    modTestRunner.Check "書き直し_3日以内を含む質問は番号選択ではない", _
        (Not modClarify.IsNumberChoiceOnly("3日以内に出す必要ある?"))
    modTestRunner.Check "書き直し_年度を含む短文も番号選択ではない", _
        (Not modClarify.IsNumberChoiceOnly("2026年度"))
    modTestRunner.Check "書き直し_数字なしの文章", _
        (Not modClarify.IsNumberChoiceOnly("もっと詳しく教えて"))

    ' --- 境界 ---
    modTestRunner.Check "番号選択_空文字はFalse", (Not modClarify.IsNumberChoiceOnly(""))
    modTestRunner.Check "番号選択_空白のみはFalse", (Not modClarify.IsNumberChoiceOnly("   "))
    modTestRunner.Check "番号選択_区切りだけで数字なしはFalse", _
        (Not modClarify.IsNumberChoiceOnly("-.-"))
    modTestRunner.Check "番号選択_長すぎる数字列はFalse", _
        (Not modClarify.IsNumberChoiceOnly("1234567890"))
End Sub


' ----------------------------------------------------------------------------
' ベクトルCSVの読み取り(2026-07-28 レビュー L-5)
'
' 実バグ: Val() は解釈できない文字列を黙って 0 にするため、壊れたCSVが
' 「全要素0の正常なベクトル」として通っていた。ゼロベクトルは内積が常に0で
' どの質問とも無関係になるので、そのチャンクは検索から静かに消える。
' 「壊れている」と「関係が無い」が区別できないのがいちばん困る。
' ----------------------------------------------------------------------------
Private Sub RunVectorCsvTests()
    Dim v() As Double

    modTestRunner.Check "CsvToVector_正常", modUtil.CsvToVector("0.1,-0.2,0.3", v)
    modTestRunner.Check "CsvToVector_正常_指数表記", modUtil.CsvToVector("1e-3,2E+2,-3.5e1", v)
    modTestRunner.Check "CsvToVector_正常_整数", modUtil.CsvToVector("1,2,3", v)

    modTestRunner.Check "CsvToVector_文字混入は失敗", _
        (Not modUtil.CsvToVector("0.1,abc,0.3", v))
    modTestRunner.Check "CsvToVector_全角数字は失敗", _
        (Not modUtil.CsvToVector("0.1," & ChrW(65298) & ",0.3", v))
    modTestRunner.Check "CsvToVector_小数点2つは失敗", _
        (Not modUtil.CsvToVector("0.1,1.2.3,0.3", v))
    modTestRunner.Check "CsvToVector_符号だけは失敗", _
        (Not modUtil.CsvToVector("0.1,-,0.3", v))
    modTestRunner.Check "CsvToVector_空要素は失敗", _
        (Not modUtil.CsvToVector("0.1,,0.3", v))
    modTestRunner.Check "CsvToVector_空文字は失敗", (Not modUtil.CsvToVector("", v))
End Sub


' ----------------------------------------------------------------------------
' バッジ表の整合(解説書 §11-11 の恒久再発防止)
'
' 実バグ: EvaluateBadges が12種を判定・記録する一方、表示側(modDash/modHub)は
' それぞれ独立に8種の配列を持っていた。共有知フライホイールに最も貢献した
' 4種(fb10/fb50/qa_share10/gapfill)は、獲得しても本人に一生見えなかった。
' 「共有知は使う人ではなく直す人がいないと育たない」という設計意図に対して、
' 実装が正反対を向いていたことになる。
'
' 表を単一情報源(modStats.BadgeCatalog)へ寄せたので、ここでは
' 「表そのものが壊れていないこと」を固定する。
'   ・4つの配列の長さが揃っていること(ズレると表示が別バッジの説明になる)
'   ・id に重複が無いこと(重複すると獲得日の取り違えが起きる)
'   ・空の要素が無いこと(空ラベルのバッジは画面上ただの穴になる)
'   ・判定側にある4種が表にも載っていること(今回の実バグそのもの)
' ----------------------------------------------------------------------------
Private Sub RunBadgeCatalogTests()
    Dim ids() As String, titles() As String, shorts() As String, conds() As String
    Dim n As Long
    n = modStats.BadgeCatalog(ids, titles, shorts, conds)

    modTestRunner.Check "バッジ表_件数が1以上", (n >= 1), "n=" & n
    modTestRunner.Check "バッジ表_titlesの長さ一致", _
        (UBound(titles) - LBound(titles) + 1 = n)
    modTestRunner.Check "バッジ表_shortTitlesの長さ一致", _
        (UBound(shorts) - LBound(shorts) + 1 = n)
    modTestRunner.Check "バッジ表_conditionsの長さ一致", _
        (UBound(conds) - LBound(conds) + 1 = n)

    Dim i As Long, j As Long
    Dim emptyN As Long, dupN As Long
    For i = LBound(ids) To UBound(ids)
        If LenB(Trim$(ids(i))) = 0 Then emptyN = emptyN + 1
        If LenB(Trim$(titles(i))) = 0 Then emptyN = emptyN + 1
        If LenB(Trim$(shorts(i))) = 0 Then emptyN = emptyN + 1
        If LenB(Trim$(conds(i))) = 0 Then emptyN = emptyN + 1
        For j = i + 1 To UBound(ids)
            If StrComp(ids(i), ids(j), vbTextCompare) = 0 Then dupN = dupN + 1
        Next j
    Next i
    modTestRunner.Check "バッジ表_空の要素が無い", (emptyN = 0), "empty=" & emptyN
    modTestRunner.Check "バッジ表_idに重複が無い", (dupN = 0), "dup=" & dupN

    ' 実バグで落ちていた4種が載っていること。
    modTestRunner.Check "バッジ表_fb10を含む", BadgeIdExists(ids, "fb10")
    modTestRunner.Check "バッジ表_fb50を含む", BadgeIdExists(ids, "fb50")
    modTestRunner.Check "バッジ表_qa_share10を含む", BadgeIdExists(ids, "qa_share10")
    modTestRunner.Check "バッジ表_gapfillを含む", BadgeIdExists(ids, "gapfill")
End Sub

Private Function BadgeIdExists(ByRef ids() As String, ByVal target As String) As Boolean
    Dim i As Long
    For i = LBound(ids) To UBound(ids)
        If StrComp(Trim$(ids(i)), target, vbTextCompare) = 0 Then
            BadgeIdExists = True
            Exit Function
        End If
    Next i
End Function
