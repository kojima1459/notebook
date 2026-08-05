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
    Resume NextDone17
End Sub
