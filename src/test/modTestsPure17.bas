Attribute VB_Name = "modTestsPure17"
Option Explicit

' ============================================================================
' modTestsPure17 - R17 Phase1(構造メタデータ+参照エッジ)の純ロジック回帰
' ----------------------------------------------------------------------------
' なぜ新設したか(憲章§4-6):
'   modTestsPure16 が20,739字で、ここの真理表(約7,000字)を足すと WARN帯
'   (28,000字)へ入る。16 を新設したときと同じ線で分割する。
'   入口は modTestsPure16.RunAll16 の末尾から呼ばれる RunAll17 の1本だけ。
'
' ここで固定するもの:
'   ・modChunkMeta.ExtractSectionPath: breadcrumb からの「章>条」抽出
'     (正常/資料名だけ/閉じ括弧無し/breadcrumb 無し/全角数字/2行目の【…】)。
'   ・modChunkMeta.ExtractRefs: 「第N条/第N項/第N章/別表N/様式N」の抽出
'     (複数/重複/参照なし/全半角/別表第N/上限24件/32,000字級)。
'   ・modChunkMeta.RefLabelsFor: 自分自身の見出し番号を落とす(自己参照の無害化)。
'   ・modChunkMeta.PathHasLabel: 第1条が第12条に誤って当たらないこと。
'   ・modChunkMeta.GraphActive: 【フェイルセーフの単一情報源】。chunk_meta が
'     0行なら参照展開も条番号の直接ヒット保証も無操作=既存本棚は従来動作。
'   ・modChunkMeta.MetaOf: 2本を1回で呼ぶ入口が個別呼びと同じ答えを返すこと。
'   ・modSparse.DistinctiveKeys: 「別表N」「様式N」を条番号と同格で拾うこと
'     (R17 Phase1 の拡張。既存の第N条の挙動は変えない)。
' ============================================================================

' ----------------------------------------------------------------------------
' section_path の抽出。
' ----------------------------------------------------------------------------
' modChunker.FlushBlock は各チャンクの先頭へ「【〔資料〕 > 章 > 条】+改行」を
' 置く。資料名の部分は my_knowledge.source が持っている情報なので必ず捨てる
' (混ぜると同じ条文が資料名の表記ゆれで別ラベルになる)。
Private Sub TestExtractSectionPath()
    Dim c1 As String
    c1 = "【〔資料〕 > 第3章 総則 > 第12条(免責)】" & vbLf & "当社は保険金を支払いません。"
    modTestRunner.Check "section_path_章と条を採り資料名は捨てる", _
        (modChunkMeta.ExtractSectionPath(c1) = "第3章 総則>第12条(免責)"), _
        "実際=" & modChunkMeta.ExtractSectionPath(c1)

    ' 章だけ(条の見出しがまだ来ていないブロック)。
    modTestRunner.Check "section_path_章だけでも取れる", _
        (modChunkMeta.ExtractSectionPath("【〔資料〕 > 第1章 目的】" & vbLf & "本文") = "第1章 目的"), _
        "実際=" & modChunkMeta.ExtractSectionPath("【〔資料〕 > 第1章 目的】" & vbLf & "本文")

    ' 資料名だけ=構造が取れなかったチャンク。空文字が正しい答え。
    modTestRunner.Check "section_path_資料名だけなら空", _
        (modChunkMeta.ExtractSectionPath("【〔資料〕】" & vbLf & "本文") = "")

    ' 形式崩れ(閉じ括弧が無い)は空。壊れた行から無理に読まない。
    modTestRunner.Check "section_path_閉じ括弧が無ければ空", _
        (modChunkMeta.ExtractSectionPath("【〔資料〕 > 第3章 総則") = "")
    modTestRunner.Check "section_path_中身が空の括弧も空", _
        (modChunkMeta.ExtractSectionPath("【】" & vbLf & "本文") = "")

    ' breadcrumb が無い(chunk_mode=structure 以外・パック由来 等)。
    modTestRunner.Check "section_path_breadcrumbが無ければ空", _
        (modChunkMeta.ExtractSectionPath("第12条(免責)当社は…") = "")
    modTestRunner.Check "section_path_空文字入力は空", _
        (modChunkMeta.ExtractSectionPath("") = "")

    ' 2行目以降の【…】は本文。1行目だけを見る。
    Dim c2 As String
    c2 = "本文の途中" & vbLf & "【〔資料〕 > 第9章】" & vbLf & "続き"
    modTestRunner.Check "section_path_2行目の括弧は見ない", _
        (modChunkMeta.ExtractSectionPath(c2) = "")

    ' 全角数字は modSparse.NormalizeForSearch で半角へ寄る(質問側と同じ式)。
    ' ここが揃っていないと「第１２条」の資料が「第12条」の質問で外れる。
    modTestRunner.Check "section_path_全角数字は半角へ寄る", _
        (modChunkMeta.ExtractSectionPath("【〔資料〕 > 第３章 > 第１２条】" & vbLf & "本文") _
         = "第3章>第12条"), _
        "実際=" & modChunkMeta.ExtractSectionPath("【〔資料〕 > 第３章 > 第１２条】" & vbLf & "本文")
End Sub

' ----------------------------------------------------------------------------
' refs_out(明示参照)の抽出。
' ----------------------------------------------------------------------------
Private Sub TestExtractRefs()
    Dim body As String
    body = "第6条(免責)本条の適用については第8条および別表2による。手続は様式3を用いる。"
    modTestRunner.Check "refs_複数の参照を順に拾う", _
        (modChunkMeta.ExtractRefs(body) = "第6条|第8条|別表2|様式3"), _
        "実際=" & modChunkMeta.ExtractRefs(body)

    ' 同じ参照が何度出ても1件(refs_out はラベルの集合)。
    modTestRunner.Check "refs_重複は1件に畳む", _
        (modChunkMeta.ExtractRefs("第8条による。第8条の定めにより。") = "第8条"), _
        "実際=" & modChunkMeta.ExtractRefs("第8条による。第8条の定めにより。")

    ' 参照が1つも無い本文は空(検索側は何もしない)。
    modTestRunner.Check "refs_参照が無ければ空", _
        (modChunkMeta.ExtractRefs("当社は保険金をお支払いします。") = "")
    modTestRunner.Check "refs_空文字入力は空", (modChunkMeta.ExtractRefs("") = "")

    ' 全角数字でも同じラベルになる(取込側と質問側で式が1本だから成立する)。
    modTestRunner.Check "refs_全角数字は半角ラベルへ", _
        (modChunkMeta.ExtractRefs("第８条および別表２のとおり") = "第8条|別表2"), _
        "実際=" & modChunkMeta.ExtractRefs("第８条および別表２のとおり")

    ' 「別表第2」は原文どおりのラベル(同じ資料の中では表記が揃っている前提。
    ' 勝手に「第」を落として寄せると、見出し側の表記と一致しなくなる)。
    modTestRunner.Check "refs_別表第Nは原文どおり", _
        (modChunkMeta.ExtractRefs("詳細は別表第2による") = "別表第2"), _
        "実際=" & modChunkMeta.ExtractRefs("詳細は別表第2による")

    ' 「号」は拾わない(条の内側の細目で、章>条の粒度の section_path と
    ' 噛み合わない=引いても当たらないラベルを refs_out に溜めない)。
    modTestRunner.Check "refs_号は拾わない", _
        (modChunkMeta.ExtractRefs("前項第3号に掲げる") = "")

    ' 上限24件。目次のように条番号が延々並ぶチャンクで1セルを埋めない。
    Dim many As String
    Dim i As Long
    For i = 1 To 40
        many = many & "第" & i & "条 "
    Next i
    Dim got As String: got = modChunkMeta.ExtractRefs(many)
    modTestRunner.Check "refs_上限24件で打ち切る", _
        (CountPipe17(got) = 24), "実際=" & CountPipe17(got) & "件 / " & got

    ' 32,000字級(1セルの上限)でも落ちず、末尾の参照まで届くこと。
    Dim big As String
    big = String$(31900, "x") & "第8条による。"
    modTestRunner.Check "refs_32000字級でも末尾の参照を拾う", _
        (modChunkMeta.ExtractRefs(big) = "第8条"), _
        "実際=" & modChunkMeta.ExtractRefs(big)
End Sub

' "|" 区切りの件数(空なら0)。
Private Function CountPipe17(ByVal s As String) As Long
    If LenB(s) = 0 Then Exit Function
    CountPipe17 = Len(s) - Len(Replace(s, "|", "")) + 1
End Function

' ----------------------------------------------------------------------------
' 自己参照の無害化と、ラベルの当て方。
' ----------------------------------------------------------------------------
' 第6条の本文には見出し行ごと入るので「第6条」自身が必ず refs_out に混ざる。
' 落とさないと、参照展開の枠(既定8件)を「自分と同じ条の別チャンク」で
' 使い切り、本来足したい第8条・別表2へ届かない。
Private Sub TestRefLabelsFor()
    modTestRunner.Check "自己参照_自分の条番号は落とす", _
        (modChunkMeta.RefLabelsFor("第6条|第8条|別表2", "第3章>第6条(免責)") = "第8条|別表2"), _
        "実際=" & modChunkMeta.RefLabelsFor("第6条|第8条|別表2", "第3章>第6条(免責)")
    modTestRunner.Check "自己参照_章も落とす(同じ章の全条を巻き込まない)", _
        (modChunkMeta.RefLabelsFor("第3章|第8条", "第3章>第6条") = "第8条"), _
        "実際=" & modChunkMeta.RefLabelsFor("第3章|第8条", "第3章>第6条")
    modTestRunner.Check "自己参照_section_pathが空なら全部残る", _
        (modChunkMeta.RefLabelsFor("第6条|第8条", "") = "第6条|第8条")
    modTestRunner.Check "自己参照_refsが空なら空", _
        (modChunkMeta.RefLabelsFor("", "第3章>第6条") = "")
    modTestRunner.Check "自己参照_全部が自分自身なら空", _
        (modChunkMeta.RefLabelsFor("第6条", "第3章>第6条") = "")

    ' 数字の直後に単位漢字が来る形なので、部分一致でも取り違えない。
    modTestRunner.Check "ラベル_第1条は第12条に当たらない", _
        (modChunkMeta.PathHasLabel("第3章>第12条", "第1条") = False)
    modTestRunner.Check "ラベル_第12条は第12条に当たる", _
        (modChunkMeta.PathHasLabel("第3章>第12条", "第12条") = True)
    modTestRunner.Check "ラベル_枝番の見出しにも当たる", _
        (modChunkMeta.PathHasLabel("第3章>第12条の2", "第12条") = True)
    modTestRunner.Check "ラベル_section_pathが空なら当たらない", _
        (modChunkMeta.PathHasLabel("", "第12条") = False)
    modTestRunner.Check "ラベル_ラベルが空なら当たらない", _
        (modChunkMeta.PathHasLabel("第3章>第12条", "") = False)
End Sub

' ----------------------------------------------------------------------------
' 既存本棚との共存(フェイルセーフ)。
' ----------------------------------------------------------------------------
' R17設計書§3は「既存資料は再取込で生成(移行処理は書かない)」と明言している。
' つまり chunk_meta が0行の本棚は【恒久的に】あり得る。そのとき
' modAskFocus.RefsExpand / ArticleEnsure は1行も足さずに戻り、回答は R16 までと
' 完全に同じになる。その条件を持っているのがこの1本(2箇所に書き写さない)。
Private Sub TestGraphActive()
    modTestRunner.Check "フェイルセーフ_chunk_meta0行なら働かない", _
        (modChunkMeta.GraphActive(0, 5) = False)
    modTestRunner.Check "フェイルセーフ_ヒット0件でも働かない", _
        (modChunkMeta.GraphActive(120, 0) = False)
    modTestRunner.Check "フェイルセーフ_負の件数でも働かない", _
        (modChunkMeta.GraphActive(-1, 5) = False)
    modTestRunner.Check "フェイルセーフ_両方あるときだけ働く", _
        (modChunkMeta.GraphActive(120, 5) = True)
    modTestRunner.Check "フェイルセーフ_1行1件の最小でも働く", _
        (modChunkMeta.GraphActive(1, 1) = True)

    ' 2026-08-05(R17H FA-8 / A-M9): 0件許容版。検索が1件も当たらなくても、
    ' 質問が条番号を名指ししているときだけ直接キーで材料を用意する経路が
    ' この引数を True で渡す。chunk_meta が0行なら【許容しても働かない】=
    ' 既存本棚の従来動作はこの引数では絶対に崩れない。
    modTestRunner.Check "0件許容_ヒット0でも働く(条番号の直接キー経路)", _
        (modChunkMeta.GraphActive(120, 0, True) = True)
    modTestRunner.Check "0件許容でもchunk_meta0行なら働かない", _
        (modChunkMeta.GraphActive(0, 0, True) = False)
    modTestRunner.Check "既定(省略)は従来どおり0件を許さない", _
        (modChunkMeta.GraphActive(120, 0) = False)
End Sub

' MetaOf(取込ループから1行で呼ぶ入口)が個別呼びと同じ答えを返すこと。
' ここがずれると、片方だけ ApplyCrumb 後のテキストから採るような取り違えが
' 起きても誰も気付けない。
Private Sub TestMetaOf()
    Dim raw As String
    raw = "【〔資料〕 > 第2章 > 第7条】" & vbLf & "第9条の定めを準用する。"
    Dim p As String, r As String
    modChunkMeta.MetaOf raw, p, r
    modTestRunner.Check "MetaOf_section_pathが個別呼びと一致", _
        (p = modChunkMeta.ExtractSectionPath(raw)), "実際=" & p
    modTestRunner.Check "MetaOf_refs_outが個別呼びと一致", _
        (r = modChunkMeta.ExtractRefs(raw)), "実際=" & r
    modTestRunner.Check "MetaOf_参照先(第9条)を拾えている", (InStr(r, "第9条") > 0), "実際=" & r
End Sub

' ----------------------------------------------------------------------------
' modSparse.DistinctiveKeys の「別表N」「様式N」拡張(R17 Phase1)。
' ----------------------------------------------------------------------------
' 従来は「別表」だけが2文字の漢字ランとして残り、番号が落ちていた。
' 「別表2の料率は?」と聞かれてどの別表も同点になる=条番号と同じ事故。
Private Sub TestDistinctiveTableKeys()
    Dim k As String
    k = modSparse.DistinctiveKeys("別表2の料率を教えて")
    modTestRunner.Check "効く語_別表Nを拾う", (InStr(k, "別表2") > 0), "実際=" & k

    k = modSparse.DistinctiveKeys("様式第3の書き方を教えて")
    modTestRunner.Check "効く語_様式第Nを原文どおり拾う", (InStr(k, "様式第3") > 0), "実際=" & k

    k = modSparse.DistinctiveKeys("別表２について")
    modTestRunner.Check "効く語_全角の別表２も半角ラベルへ", (InStr(k, "別表2") > 0), "実際=" & k

    ' 同じ別表を2回書いても1件(キーが重複すると採れる語の枠を食う)。
    k = modSparse.DistinctiveKeys("別表2と別表2の違い")
    modTestRunner.Check "効く語_別表の重複は1件", _
        (Len(k) - Len(Replace(k, "別表2", "")) = Len("別表2")), "実際=" & k

    ' 条番号と同格=どちらも先頭側に置かれる。
    k = modSparse.DistinctiveKeys("第5条と別表2の関係")
    modTestRunner.Check "効く語_条番号と別表が両方入る", _
        (InStr(k, "第5条") > 0 And InStr(k, "別表2") > 0), "実際=" & k

    ' 既存の挙動は変えない(R16までの回帰)。
    k = modSparse.DistinctiveKeys("第12条について教えて")
    modTestRunner.Check "効く語_条番号は従来どおり拾う", (InStr(k, "第12条") > 0), "実際=" & k
    k = modSparse.DistinctiveKeys("保険金を支払わない場合")
    modTestRunner.Check "効く語_別表を含まない質問は従来どおり", _
        (InStr(k, "別表") = 0 And InStr(k, "様式") = 0), "実際=" & k
End Sub

' ============================================================================
' R17 Phase2(章単位要約=疑似グローバル検索)
' ============================================================================

' ----------------------------------------------------------------------------
' 章グルーピングキー(modOutlineBuild.ChapterKeyOf)。
' ----------------------------------------------------------------------------
' section_path の第1要素をそのまま使う(正規化を足さない)。ここが取込側と
' 照合側で1文字でも違うと、doc_outline に貯めた章キーと chunk_meta の章キーが
' 一致せず、章を選べたのに本文が1件も引けない(=俯瞰が無音で死ぬ)。
Private Sub TestChapterKeyOf()
    modTestRunner.Check "章キー_章>条なら章だけを採る", _
        (modOutlineBuild.ChapterKeyOf("第3章 総則>第12条(免責)") = "第3章 総則"), _
        "実際=" & modOutlineBuild.ChapterKeyOf("第3章 総則>第12条(免責)")

    ' 章見出しの無い資料は第1要素が条になる(=条単位の要約)。これが正しい
    ' 保守的動作で、無い章立てを推測するより外れ方が小さい(R17 Phase2 裁定)。
    modTestRunner.Check "章キー_章が無ければ条単位になる", _
        (modOutlineBuild.ChapterKeyOf("第12条(免責)") = "第12条(免責)")

    ' 3階層以上でも先頭だけ(将来 section_path が伸びても章の粒度は変わらない)。
    modTestRunner.Check "章キー_3階層でも先頭だけ", _
        (modOutlineBuild.ChapterKeyOf("第1章>第2節>第3条") = "第1章")

    modTestRunner.Check "章キー_空文字は空", (modOutlineBuild.ChapterKeyOf("") = "")
    modTestRunner.Check "章キー_区切りだけは空", (modOutlineBuild.ChapterKeyOf(">") = "")
    modTestRunner.Check "章キー_先頭が空の階層は空", (modOutlineBuild.ChapterKeyOf(">第3条") = "")
    modTestRunner.Check "章キー_前後の空白は落とす", _
        (modOutlineBuild.ChapterKeyOf("  第1章 目的 > 第1条 ") = "第1章 目的")
End Sub

' ----------------------------------------------------------------------------
' 1章の本文の予算打ち切り(modOutlineBuild.BudgetTake)。
' ----------------------------------------------------------------------------
' 1回のLLM呼び出しに載る量は決まっている(max_context_chars)。手前で切らないと
' LLM側で切られ、章の後半が理由も分からず消える。俯瞰(modAskGlobal)では
' 章ごとに予算を等分するので、ここが甘いと先頭の章だけで予算を使い切り、
' 後ろの章が丸ごと落ちた「俯瞰したのに1章しか読んでいない」回答になる。
Private Sub TestBudgetTake()
    modTestRunner.Check "予算_余裕があれば全部入る", (modOutlineBuild.BudgetTake(0, 100, 1000) = 100)
    modTestRunner.Check "予算_残りぶんだけ入る", (modOutlineBuild.BudgetTake(950, 100, 1000) = 50)
    modTestRunner.Check "予算_ちょうど使い切ったら0", (modOutlineBuild.BudgetTake(1000, 100, 1000) = 0)
    modTestRunner.Check "予算_超過していても負にならない", _
        (modOutlineBuild.BudgetTake(1200, 100, 1000) = 0)
    modTestRunner.Check "予算_境界(残りと同じ長さ)は全部入る", _
        (modOutlineBuild.BudgetTake(900, 100, 1000) = 100)
    modTestRunner.Check "予算_上限0なら1字も入らない", (modOutlineBuild.BudgetTake(0, 100, 0) = 0)
    modTestRunner.Check "予算_上限が負でも0", (modOutlineBuild.BudgetTake(0, 100, -5) = 0)
    modTestRunner.Check "予算_足す長さが0なら0", (modOutlineBuild.BudgetTake(0, 0, 1000) = 0)
End Sub

' ----------------------------------------------------------------------------
' 章要約の応答パーサ(modRagParse.ParseOutlineResp)。
' ----------------------------------------------------------------------------
Private Sub TestParseOutlineResp()
    Dim sm As String, kw As String
    Dim okFlag As Boolean

    okFlag = modRagParse.ParseOutlineResp( _
        "<summary>この章は免責事由を定める。</summary><keywords>免責|支払|例外</keywords>", sm, kw)
    modTestRunner.Check "章要約_正常に読める", okFlag
    modTestRunner.Check "章要約_summaryが取れる", (sm = "この章は免責事由を定める。"), "実際=" & sm
    modTestRunner.Check "章要約_keywordsが取れる", (kw = "免責|支払|例外"), "実際=" & kw

    ' 読点・カンマ区切りで返すモデルがあるので "|" へ寄せる(空要素は落とす)。
    okFlag = modRagParse.ParseOutlineResp( _
        "<summary>要約</summary><keywords>免責、支払, 例外|</keywords>", sm, kw)
    modTestRunner.Check "章要約_読点とカンマは|へ寄せる", (kw = "免責|支払|例外"), "実際=" & kw

    ' keywords が無くても summary が読めれば成功(キーワードは選択段の補助)。
    okFlag = modRagParse.ParseOutlineResp("<summary>要約だけ</summary>", sm, kw)
    modTestRunner.Check "章要約_keywords欠落でも成功", (okFlag And sm = "要約だけ" And kw = "")

    ' summary が無い/空/タグ崩れ/#ERR は失敗=呼び出し元が「(要約失敗)」を書く。
    okFlag = modRagParse.ParseOutlineResp("<keywords>a|b</keywords>", sm, kw)
    modTestRunner.Check "章要約_summary欠落は失敗", (okFlag = False And sm = "")
    okFlag = modRagParse.ParseOutlineResp("<summary>   </summary>", sm, kw)
    modTestRunner.Check "章要約_中身が空白だけは失敗", (okFlag = False)
    okFlag = modRagParse.ParseOutlineResp("この章は免責を定めます(タグ無し)", sm, kw)
    modTestRunner.Check "章要約_タグ無しは失敗", (okFlag = False)
    okFlag = modRagParse.ParseOutlineResp("#ERR:E0202:応答が空でした", sm, kw)
    modTestRunner.Check "章要約_#ERRは失敗", (okFlag = False)
    ' 失敗時は出力を必ず空へ戻す(前の章の要約が残ると別の章の要約として保存される)。
    modTestRunner.Check "章要約_失敗時は出力が空へ戻る", (sm = "" And kw = "")
End Sub

' ----------------------------------------------------------------------------
' 章選択の応答パーサ(modRagParse.ParseChapterPick)。
' ----------------------------------------------------------------------------
Private Sub TestParseChapterPick()
    Dim picks() As String
    Dim n As Long

    n = modRagParse.ParseChapterPick( _
        "<pick>規程.pdf::第1章 総則|規程.pdf::第3章 保険金</pick>", 4, picks)
    modTestRunner.Check "章選択_2件読める", (n = 2), "実際=" & n
    modTestRunner.Check "章選択_1件目が原文どおり", _
        (picks(LBound(picks)) = "規程.pdf::第1章 総則"), "実際=" & picks(LBound(picks))

    ' 空=1章も選ばれなかった(mockの<pick></pick>もここ)。俯瞰は不発で
    ' 従来の入念フローへ落ちる=「関係ない章の本文で回答を作る」を作らない。
    n = modRagParse.ParseChapterPick("<pick></pick>", 4, picks)
    modTestRunner.Check "章選択_空は0件", (n = 0), "実際=" & n
    n = modRagParse.ParseChapterPick("該当する章はありません", 4, picks)
    modTestRunner.Check "章選択_タグ無しは0件", (n = 0), "実際=" & n
    n = modRagParse.ParseChapterPick("#ERR:E0201:AIリボンが見つかりません", 4, picks)
    modTestRunner.Check "章選択_#ERRは0件", (n = 0), "実際=" & n

    ' "::" が無い指定は捨てる(資料名か章キーが欠けた指定は、別の資料の
    ' 同じ章名に当たり得る。「第1章 総則」はどの規程にもある)。
    n = modRagParse.ParseChapterPick("<pick>第1章 総則|規程.pdf::第2章</pick>", 4, picks)
    modTestRunner.Check "章選択_資料名の無い指定は捨てる", (n = 1), "実際=" & n
    modTestRunner.Check "章選択_残るのは正しい形の方", _
        (picks(LBound(picks)) = "規程.pdf::第2章"), "実際=" & picks(LBound(picks))
    n = modRagParse.ParseChapterPick("<pick>規程.pdf::|::第2章</pick>", 4, picks)
    modTestRunner.Check "章選択_片側が空の指定は捨てる", (n = 0), "実際=" & n

    ' 上限。5件返ってきても4件まで(1回のプロンプトへ全章の本文を載せるため、
    ' 章が増えるほど1章あたりの取り分が減ってどの章も途中で切れる)。
    n = modRagParse.ParseChapterPick( _
        "<pick>a.pdf::1|a.pdf::2|a.pdf::3|a.pdf::4|a.pdf::5</pick>", 4, picks)
    modTestRunner.Check "章選択_上限4件で切る", (n = 4), "実際=" & n
    ' 呼び出し側が大きい上限を渡しても4を超えない(上限の実体はパーサ側)。
    n = modRagParse.ParseChapterPick( _
        "<pick>a.pdf::1|a.pdf::2|a.pdf::3|a.pdf::4|a.pdf::5</pick>", 99, picks)
    modTestRunner.Check "章選択_上限99を渡しても4件", (n = 4), "実際=" & n
End Sub

' ----------------------------------------------------------------------------
' 俯瞰のフェイルセーフ(modAskGlobal.OutlineActive)と verdict=global。
' ----------------------------------------------------------------------------
' doc_outline が0行(=章要約をまだ作っていない本棚・graph_outline=off のまま
' 使ってきた本棚・R16以前からの本棚)なら俯瞰は一切動かず、回答は従来と
' 1文字も変わらない。GraphActive(Phase1)と同じ「条件は1本だけ」の作法。
Private Sub TestOutlineFailsafe()
    modTestRunner.Check "俯瞰_doc_outline0行なら働かない", _
        (modAskGlobal.OutlineActive(0) = False)
    modTestRunner.Check "俯瞰_負の件数でも働かない", _
        (modAskGlobal.OutlineActive(-1) = False)
    modTestRunner.Check "俯瞰_1行あれば働く", (modAskGlobal.OutlineActive(1) = True)
    modTestRunner.Check "俯瞰_多数行でも働く", (modAskGlobal.OutlineActive(320) = True)

    ' 段0の判定語彙。global を知らない応答(旧プロンプト・旧モデル・タグ崩れ)は
    ' 従来どおり single へ落ちる=俯瞰は上積みであって置き換えではない。
    modTestRunner.Check "段0_globalを読める", _
        (modRagParse.ParseDecomposeVerdict("<verdict>global</verdict>") = "global")
    modTestRunner.Check "段0_GLOBAL大文字も読める", _
        (modRagParse.ParseDecomposeVerdict("<verdict>GLOBAL</verdict>") = "global")
    modTestRunner.Check "段0_知らない語はsingle", _
        (modRagParse.ParseDecomposeVerdict("<verdict>overview</verdict>") = "single")
    modTestRunner.Check "段0_既存の3語は不変", _
        (modRagParse.ParseDecomposeVerdict("<verdict>parts</verdict>") = "parts" And _
         modRagParse.ParseDecomposeVerdict("<verdict>clarify</verdict>") = "clarify" And _
         modRagParse.ParseDecomposeVerdict("#ERR:E0202:x") = "single")
End Sub

' ============================================================================
' R17 Phase3(用語の名寄せ)
' ============================================================================

' ----------------------------------------------------------------------------
' 名寄せ応答の解析(modRagParse.ParseSynResp)。
'   出力契約 <syn>表記>正規形|表記>正規形</syn>。
' ----------------------------------------------------------------------------
Private Sub TestParseSynResp()
    Dim terms() As String, canons() As String, n As Long

    n = modRagParse.ParseSynResp("<syn>回収>リコール|解約>解約</syn>", terms, canons)
    modTestRunner.Check "名寄せ応答_正常系は2件", (n = 2), "実際=" & n
    modTestRunner.Check "名寄せ応答_1件目の表記と正規形", _
        (terms(0) = "回収" And canons(0) = "リコール"), _
        "実際=" & terms(0) & "/" & canons(0)
    modTestRunner.Check "名寄せ応答_2件目も読める", _
        (terms(1) = "解約" And canons(1) = "解約"), _
        "実際=" & terms(1) & "/" & canons(1)

    ' タグ欠落(<syn>が無い応答)は0件。
    n = modRagParse.ParseSynResp("説明だけの応答です", terms, canons)
    modTestRunner.Check "名寄せ応答_タグ欠落は0件", (n = 0), "実際=" & n
    n = modRagParse.ParseSynResp("<syn></syn>", terms, canons)
    modTestRunner.Check "名寄せ応答_空タグも0件", (n = 0), "実際=" & n

    ' #ERR: はCallLLM失敗の応答契約。全て0件(modOutlineBuild/enrich系と同じ寛容退化)。
    n = modRagParse.ParseSynResp("#ERR:E0201:AIリボンが見つかりません", terms, canons)
    modTestRunner.Check "名寄せ応答_ERRは0件", (n = 0), "実際=" & n

    ' 不正ペア破棄: ">"が無い/片側が空、は1件ずつ捨てて読めた分だけ残す。
    n = modRagParse.ParseSynResp("<syn>回収リコール|解約>|>正規形|表記>正規形</syn>", terms, canons)
    modTestRunner.Check "名寄せ応答_不正ペアを捨てて1件残る", (n = 1), "実際=" & n
    modTestRunner.Check "名寄せ応答_残った1件の中身", _
        (terms(0) = "表記" And canons(0) = "正規形"), _
        "実際=" & terms(0) & "/" & canons(0)
End Sub

' ----------------------------------------------------------------------------
' クエリ展開(modRagParse.ExpandQueryBySyn)。
'   mapCsv は modSynonymStore.ReadMapCsv が返す形("term>canonical|…")のまま
'   純関数へ渡す(シートI/Oはこの関数に無いのでLOでも直接検証できる)。
'   「同義語展開_空マップは無操作」が【ReadMapCsvの契約=0行なら空文字列】を
'   受けたときの振る舞いをそのまま固定する(ReadMapCsv自体はシートI/Oのため
'   ReadOutline/ReadAllMeta と同じくLO純ロジックテストの対象に出来ない)。
' ----------------------------------------------------------------------------
Private Sub TestExpandQueryBySyn()
    Dim q As String

    ' 一致0(その1): 空マップ=ReadMapCsvの「0行なら空文字」契約どおりの入力。
    q = modRagParse.ExpandQueryBySyn("回収の保険は?", "", 3)
    modTestRunner.Check "同義語展開_空マップは無操作", (q = "回収の保険は?"), "実際=" & q

    ' 一致0(その2): マップは有るが質問中のどの語にも当たらない。
    q = modRagParse.ExpandQueryBySyn("解約の手続きは?", "回収>リコール", 3)
    modTestRunner.Check "同義語展開_不一致なら無操作", (q = "解約の手続きは?"), "実際=" & q

    ' 1件一致(表記→正規形)。
    q = modRagParse.ExpandQueryBySyn("回収の保険は?", "回収>リコール", 3)
    modTestRunner.Check "同義語展開_1件一致で追記", (q = "回収の保険は? リコール"), "実際=" & q

    ' 双方向(正規形→表記): 質問側が正規形でも、対応する表記が足される。
    q = modRagParse.ExpandQueryBySyn("リコールの保険は?", "回収>リコール", 3)
    modTestRunner.Check "同義語展開_逆方向でも追記", (q = "リコールの保険は? 回収"), "実際=" & q

    ' 3件上限: 4件一致しても3件で打ち切る(足す語は質問文に無い字を選び、
    ' 「について」のような助詞の中に偶然含まれて二重追記防止に弾かれないようにする)。
    q = modRagParse.ExpandQueryBySyn("AとBとCとDの内容", "A>ア|B>イ|C>ウ|D>エ", 3)
    modTestRunner.Check "同義語展開_上限3件で打ち切る", _
        (q = "AとBとCとDの内容 ア イ ウ"), "実際=" & q

    ' 全角/半角: 質問側が全角・マップ側が半角でもNormalizeForSearch経由で一致。
    q = modRagParse.ExpandQueryBySyn("ＡＢＣの申請", "abc>エービーシー", 3)
    modTestRunner.Check "同義語展開_全角半角の表記ゆれでも一致", _
        (q = "ＡＢＣの申請 エービーシー"), "実際=" & q

    ' 自己一致除外: termとcanonicalが同じ(正規化後)ペアは何も足さない。
    q = modRagParse.ExpandQueryBySyn("回収の保険は?", "回収>回収", 3)
    modTestRunner.Check "同義語展開_自己一致は除外", (q = "回収の保険は?"), "実際=" & q

    ' 質問に両方の語が既にあれば追記しない(重複を増やさない)。
    q = modRagParse.ExpandQueryBySyn("回収とリコールの違いは?", "回収>リコール", 3)
    modTestRunner.Check "同義語展開_既に両方あれば追記しない", _
        (q = "回収とリコールの違いは?"), "実際=" & q

    ' maxAdd<1は無操作。
    q = modRagParse.ExpandQueryBySyn("回収の保険は?", "回収>リコール", 0)
    modTestRunner.Check "同義語展開_maxAdd0は無操作", (q = "回収の保険は?"), "実際=" & q
End Sub

Public Sub RunAll17()
    On Error GoTo PathFail17
    TestExtractSectionPath
NextRefs17:
    On Error GoTo RefsFail17
    TestExtractRefs
NextSelf17:
    On Error GoTo SelfFail17
    TestRefLabelsFor
NextGraph17:
    On Error GoTo GraphFail17
    TestGraphActive
NextMetaOf17:
    On Error GoTo MetaOfFail17
    TestMetaOf
NextTable17:
    On Error GoTo TableFail17
    TestDistinctiveTableKeys
NextChapKey17:
    On Error GoTo ChapKeyFail17
    TestChapterKeyOf
NextBudget17:
    On Error GoTo BudgetFail17
    TestBudgetTake
NextOutline17:
    On Error GoTo OutlineFail17
    TestParseOutlineResp
NextPick17:
    On Error GoTo PickFail17
    TestParseChapterPick
NextGlobal17:
    On Error GoTo GlobalFail17
    TestOutlineFailsafe
NextSynParse17:
    On Error GoTo SynParseFail17
    TestParseSynResp
NextSynExpand17:
    On Error GoTo SynExpandFail17
    TestExpandQueryBySyn
NextChain18:
    ' 2026-08-05(R17H): modTestsPure17 に R17H の真理表(名寄せマージ・俯瞰
    ' シグナル)を足すと WARN帯へ入るため 18 を新設した。連鎖の入口はここ1本。
    On Error GoTo ChainFail18
    modTestsPure18.RunAll18
NextDone17:
    On Error GoTo 0
    Exit Sub

PathFail17:
    modTestRunner.Check "TestExtractSectionPath(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRefs17
RefsFail17:
    modTestRunner.Check "TestExtractRefs(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextSelf17
SelfFail17:
    modTestRunner.Check "TestRefLabelsFor(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGraph17
GraphFail17:
    modTestRunner.Check "TestGraphActive(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextMetaOf17
MetaOfFail17:
    modTestRunner.Check "TestMetaOf(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextTable17
TableFail17:
    modTestRunner.Check "TestDistinctiveTableKeys(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextChapKey17
ChapKeyFail17:
    modTestRunner.Check "TestChapterKeyOf(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBudget17
BudgetFail17:
    modTestRunner.Check "TestBudgetTake(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextOutline17
OutlineFail17:
    modTestRunner.Check "TestParseOutlineResp(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextPick17
PickFail17:
    modTestRunner.Check "TestParseChapterPick(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGlobal17
GlobalFail17:
    modTestRunner.Check "TestOutlineFailsafe(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextSynParse17
SynParseFail17:
    modTestRunner.Check "TestParseSynResp(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextSynExpand17
SynExpandFail17:
    modTestRunner.Check "TestExpandQueryBySyn(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextChain18
ChainFail18:
    modTestRunner.Check "modTestsPure18.RunAll18(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone17
End Sub
