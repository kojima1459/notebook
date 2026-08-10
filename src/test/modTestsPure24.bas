Attribute VB_Name = "modTestsPure24"
Option Explicit

' ============================================================================
' modTestsPure24 - R25-3(バッジ16枠化・実機第11報⑥)の純ロジック回帰。
'   入口は modTestsPure23.RunAll23 の末尾から呼ばれる RunAll24 の1本
'   (modTestsPure20〜23と同型の連鎖)。
' ----------------------------------------------------------------------------
' 背景(docs/dev/spec_20260810_R25_実機第11報.md R25-3):
'   ・FA-R25-3a: qa_share10(qa_shared_total)/gapfill(gapfill_total)は
'     加算箇所がリポジトリに存在せず獲得不可能だった。modInsightIo.
'     EmitVerifiedQA/EmitCorrection の書込み成功時にBumpを配線した
'     (加算そのものはWorksheetを書くためPureテストの対象外。ここでは
'     「バッジ表に載っていること」を固定する)。
'   ・FA-R25-3b: 新規3種(thanks5/streak30/thorough10)を追加し、
'     13+3=16(ダッシュ4列×4段がちょうど埋まる)にした。
'   modTestsPure2.RunBadgeCatalogTests が表の一般的な整合(長さ一致・
'   空要素なし・重複なし・旧4種の存在)を既に固定しているため、ここでは
'   「16枠化」という今回の変更そのもの(件数=16固定・新3種の存在・
'   閾値を表す数字が条件文に含まれること)に絞る。
' ============================================================================

Private Sub TestBadgeCatalogCount16()
    Dim ids() As String, titles() As String, shorts() As String, conds() As String
    Dim n As Long
    n = modStats.BadgeCatalog(ids, titles, shorts, conds)

    ' 恒真防止: 13(旧)+3(新規)=16をリテラルで固定する(将来ここが動いたら
    ' 「意図した増減か」をレビューで必ず立ち止まらせるための固定値)。
    modTestRunner.Check "R25-3_BadgeCatalog件数はちょうど16(ダッシュ4x4を埋める)", _
        (n = 16), "n=" & n
End Sub

Private Sub TestBadgeCatalogNewIdsExist()
    Dim ids() As String, titles() As String, shorts() As String, conds() As String
    Dim n As Long
    n = modStats.BadgeCatalog(ids, titles, shorts, conds)

    ' 新3種(FA-R25-3b)。
    modTestRunner.Check "R25-3_バッジ表_thanks5を含む", BadgeIdExists24(ids, "thanks5")
    modTestRunner.Check "R25-3_バッジ表_streak30を含む", BadgeIdExists24(ids, "streak30")
    modTestRunner.Check "R25-3_バッジ表_thorough10を含む", BadgeIdExists24(ids, "thorough10")

    ' 修理した死にバッジ2種(FA-R25-3a)。modTestsPure2側でも固定済みだが、
    ' 「16枠が旧13種+新3種で構成される」という今回の変更の前提そのものを
    ' このテストファイル単体でも確認できるよう重ねて固定する。
    modTestRunner.Check "R25-3_バッジ表_qa_share10を含む(修理対象)", BadgeIdExists24(ids, "qa_share10")
    modTestRunner.Check "R25-3_バッジ表_gapfillを含む(修理対象)", BadgeIdExists24(ids, "gapfill")
End Sub

Private Sub TestBadgeCatalogNoDuplicateIds()
    Dim ids() As String, titles() As String, shorts() As String, conds() As String
    Dim n As Long
    n = modStats.BadgeCatalog(ids, titles, shorts, conds)

    Dim i As Long, j As Long, dupN As Long
    For i = LBound(ids) To UBound(ids)
        For j = i + 1 To UBound(ids)
            If StrComp(ids(i), ids(j), vbTextCompare) = 0 Then dupN = dupN + 1
        Next j
    Next i
    modTestRunner.Check "R25-3_バッジ表_16種にidの重複が無い(新3種を足しても)", _
        (dupN = 0), "dup=" & dupN
End Sub

' 新3種の条件文(長文説明)に、閾値を表す数字が正しく入っていること
' (コピペ改変で閾値の数字だけ取り違える事故を防ぐ)。EvaluateBadges本体は
' GetStat経由でWorksheetに触れるためPureテストから直接は呼べない
' (modStats冒頭・modOutlineBuild等の既存注記と同型)。ここではBadgeCatalog
' (Worksheetに触れない純関数)側の文言が閾値と矛盾していないことだけを固定する。
Private Sub TestNewBadgeConditionsMentionThreshold()
    Dim ids() As String, titles() As String, shorts() As String, conds() As String
    Dim n As Long
    n = modStats.BadgeCatalog(ids, titles, shorts, conds)

    Dim idxThanks As Long, idxStreak As Long, idxThorough As Long
    idxThanks = IndexOf24(ids, "thanks5")
    idxStreak = IndexOf24(ids, "streak30")
    idxThorough = IndexOf24(ids, "thorough10")

    modTestRunner.Check "R25-3_境界_thanks5の条件文に閾値5が入っている", _
        (idxThanks >= 0), "idx=" & idxThanks
    If idxThanks >= 0 Then
        modTestRunner.Check "R25-3_境界_thanks5条件文=『感謝を5件受け取ると獲得』", _
            (InStr(conds(idxThanks), "5") > 0), "cond=" & conds(idxThanks)
    End If

    modTestRunner.Check "R25-3_境界_streak30の条件文に閾値30が入っている", _
        (idxStreak >= 0), "idx=" & idxStreak
    If idxStreak >= 0 Then
        modTestRunner.Check "R25-3_境界_streak30条件文に『30』が入っている(29/30の境界を表す数字)", _
            (InStr(conds(idxStreak), "30") > 0), "cond=" & conds(idxStreak)
    End If

    modTestRunner.Check "R25-3_境界_thorough10の条件文に閾値10が入っている", _
        (idxThorough >= 0), "idx=" & idxThorough
    If idxThorough >= 0 Then
        modTestRunner.Check "R25-3_境界_thorough10条件文に『10』が入っている(9/10の境界を表す数字)", _
            (InStr(conds(idxThorough), "10") > 0), "cond=" & conds(idxThorough)
    End If
End Sub

' F7(m-5): バッジ絵文字二重の是正。旧チェックは新3種の先頭だけを見ていたが、
' 本来固定すべきは「長い名前(titles)に絵文字を含まない」という表全体の
' 意匠統一(modHub.DrawBadges/modDash側が既に🏅/🔒を前置するため、titles
' 自身にも絵文字があると二重表示になる)。全16種を対象に、サロゲートペア
' (絵文字はUTF-16で上位D800-DBFF+下位DC00-DFFFの2コードから成る)の
' 有無で判定する。
Private Sub TestBadgeTitlesHaveNoSurrogate24()
    Dim ids() As String, titles() As String, shorts() As String, conds() As String
    Dim n As Long
    n = modStats.BadgeCatalog(ids, titles, shorts, conds)

    Dim i As Long, badN As Long
    For i = LBound(titles) To UBound(titles)
        If ContainsSurrogate24(titles(i)) Then badN = badN + 1
    Next i
    modTestRunner.Check "R25-3_バッジ表_長い名前16種に絵文字(サロゲート)を含まない", _
        (badN = 0), "badN=" & badN
End Sub

' ============================================================================
' R27 波1(実機第12報②「特約が検索から消える」の根治)の純ロジック回帰。
' ============================================================================
' F1-1: modSparse.CapKeyScore — キーワード加点の頭打ち。
'   実機ではKeyScoreが330まで伸び、SPARSE_WEIGHT 0.06 を掛けた加点20が
'   cos類似度(-1〜1)を押し流していた。cap=10(加点0.6)でcosと同じ土俵へ戻す。
'   cap未満/ちょうど/超過 の3点を固定する(cap+10 は「越えた分は切る」側)。
Private Sub TestCapKeyScoreBoundary24()
    Dim cap As Double: cap = 10#

    ' cap-1: 頭打ちの手前は素通し(値が変わってはいけない)。
    modTestRunner.Check "R27-F1-1_cap-1は素通し(9→9)", _
        (modSparse.CapKeyScore(cap - 1#, cap) = cap - 1#), _
        "実際=" & Format$(modSparse.CapKeyScore(cap - 1#, cap), "0.000")

    ' cap ちょうど: 境界は「切らない」側(> で比較しているため)。
    modTestRunner.Check "R27-F1-1_capちょうどは素通し(10→10)", _
        (modSparse.CapKeyScore(cap, cap) = cap), _
        "実際=" & Format$(modSparse.CapKeyScore(cap, cap), "0.000")

    ' cap+10: 超過分は必ず落ちる(ここが効かないとR27以前へ逆戻り)。
    modTestRunner.Check "R27-F1-1_cap+10は頭打ち(20→10)", _
        (modSparse.CapKeyScore(cap + 10#, cap) = cap), _
        "実際=" & Format$(modSparse.CapKeyScore(cap + 10#, cap), "0.000")

    ' 実機で観測された生スコア330も同じ1本の式で10へ落ちること。
    modTestRunner.Check "R27-F1-1_実機の暴走値330も10へ落ちる", _
        (modSparse.CapKeyScore(330#, cap) = cap), _
        "実際=" & Format$(modSparse.CapKeyScore(330#, cap), "0.000")

    ' cap<=0 は「頭打ちなし」= 旧挙動へ戻すエスケープハッチ。
    modTestRunner.Check "R27-F1-1_cap=0は頭打ちなし(330がそのまま返る)", _
        (modSparse.CapKeyScore(330#, 0#) = 330#), _
        "実際=" & Format$(modSparse.CapKeyScore(330#, 0#), "0.000")
End Sub

' F1-2: modSparse.KeyLenWeight — 語長重み Len^1.5 の Len を8でclamp。
'   7→8 は増え、8→9 以降は増えない。9字で 9^1.5=27 が返ったら clamp が
'   効いていない(=長い資料名キー1本で順位が決まる旧挙動)。
Private Sub TestKeyLenWeightClamp24()
    Dim w7 As Double, w8 As Double, w9 As Double, w16 As Double
    w7 = modSparse.KeyLenWeight(7)
    w8 = modSparse.KeyLenWeight(8)
    w9 = modSparse.KeyLenWeight(9)
    w16 = modSparse.KeyLenWeight(16)

    ' clampの手前(7字)は8字より必ず軽い=「長い語ほど重い」性質は残す。
    modTestRunner.Check "R27-F1-2_7字は8字より軽い(clamp手前は従来どおり)", _
        (w7 < w8), "w7=" & Format$(w7, "0.000") & " w8=" & Format$(w8, "0.000")

    ' 境界(8字)の値そのもの: 8^1.5 = 22.627…(恒真化を避けるため実値で固定)
    modTestRunner.Check "R27-F1-2_8字の重みは8^1.5=22.627", _
        (Abs(w8 - 22.6274169979695) < 0.000001), "w8=" & Format$(w8, "0.000000")

    ' clamp本体: 9字は8字と同じ値へ寝る(増えない)。
    modTestRunner.Check "R27-F1-2_9字は8字と同値(clampが効いている)", _
        (w9 = w8), "w9=" & Format$(w9, "0.000") & " w8=" & Format$(w8, "0.000")

    ' clampが無ければ 9^1.5 = 27 になる。27未満であることを別途固定する
    ' (上の等値だけだと「両方27」でも通ってしまうため)。
    modTestRunner.Check "R27-F1-2_9字の重みは27(=9^1.5)より小さい", _
        (w9 < 27#), "w9=" & Format$(w9, "0.000")

    ' 資料名級の長語(16字)でも8字止まり。16^1.5=64 が入るのが実機の暴走源。
    modTestRunner.Check "R27-F1-2_16字の資料名キーも8字止まり(64点にならない)", _
        (w16 = w8), "w16=" & Format$(w16, "0.000")
End Sub

' F1-1 の結線確認: KeyScore 本体が cap を通っていること。
' 9字キーが1回だけ出る短い本文の生スコアは 22.4 前後(clamp後22.627を
' 文書長正規化1.0099で割った値)で、cap(既定10)が効いていれば10になる。
' 既定値はビルド時 config(sparse_keyscore_cap=10)およびLO実行テスト
' (modConfig未注入 → 既定値へフォールバック)の双方で10。
Private Sub TestKeyScoreIsCapped24()
    Dim key As String: key = "特約条項変更届出書"          ' 9字
    Dim doc As String: doc = modSparse.CompactForMatch(key)
    Dim v As Double: v = modSparse.KeyScore(key, doc)

    modTestRunner.Check "R27-F1-1_KeyScore本体が上限10を超えない", _
        (v <= 10.000001), "v=" & Format$(v, "0.000")
    ' 0点に落ちていない(頭打ちと取りこぼしを取り違えない)。
    modTestRunner.Check "R27-F1-1_頭打ちしても加点は消えない(v>0)", _
        (v > 0#), "v=" & Format$(v, "0.000")
End Sub

' F1-3: modSparse.DiversityOrder — 候補プールの資料多様性。
'   実機の形をそのまま写す: 105チャンクある「就業規則」がpoolを埋め、
'   少数派の「興行中止保険特約」が1件だけ後方にいる。並べ替え後は
'   少数派資料の最高スコア1件が【先頭群】へ上がっていなければならない
'   (逆質問はsrc>=2が必須なので、ここが効かないと構造的に鳴らない)。
Private Sub TestDiversityOrderMinoritySurvives24()
    ' 添字1..6。1位資料=就業規則(0.90/0.88/0.86/0.84)、少数派=特約(0.80)を5番目に。
    Dim src(1 To 6) As String
    Dim sc(1 To 6) As Double
    src(1) = "就業規則.pdf": sc(1) = 0.9
    src(2) = "就業規則.pdf": sc(2) = 0.88
    src(3) = "就業規則.pdf": sc(3) = 0.86
    src(4) = "就業規則.pdf": sc(4) = 0.84
    src(5) = "興行中止保険特約.doc": sc(5) = 0.8
    src(6) = "就業規則.pdf": sc(6) = 0.78

    Dim ord() As Long
    Dim n As Long: n = modSparse.DiversityOrder(src, sc, 6, ord)

    modTestRunner.Check "R27-F1-3_件数は増減しない(6→6)", (n = 6), "n=" & n
    If n <> 6 Then Exit Sub

    ' 全体1位(就業規則の0.90)は動かない。
    modTestRunner.Check "R27-F1-3_全体1位は先頭のまま(元添字1)", _
        (ord(1) = 1), "ord(1)=" & ord(1)
    ' 少数派資料の代表(元添字5)が2番目=先頭群へ繰り上がる。
    modTestRunner.Check "R27-F1-3_少数派資料の最高1件が2番目へ繰り上がる(元添字5)", _
        (ord(2) = 5), "ord(2)=" & ord(2)
    ' 残りは元の順序のまま(2,3,4,6)。
    modTestRunner.Check "R27-F1-3_残りは元順のまま(2,3,4,6)", _
        (ord(3) = 2 And ord(4) = 3 And ord(5) = 4 And ord(6) = 6), _
        "ord=" & ord(3) & "," & ord(4) & "," & ord(5) & "," & ord(6)
    ' 添字の欠落・重複が無い(1..6が1回ずつ)。
    modTestRunner.Check "R27-F1-3_元添字1..6が過不足なく1回ずつ", _
        OrderIsPermutation24(ord, 6), "ord=" & JoinOrder24(ord, 6)

    ' 上位2件だけを見ても資料が2種類ある = 逆質問(src>=2)が成立する。
    modTestRunner.Check "R27-F1-3_上位2件で資料が2種類になる(逆質問が鳴る条件)", _
        (StrComp(src(ord(1)), src(ord(2)), vbTextCompare) <> 0), _
        "1=" & src(ord(1)) & " 2=" & src(ord(2))
End Sub

' 1資料しか無いプールでは順序が1つも動かないこと(副作用を作らない)。
Private Sub TestDiversityOrderSingleSource24()
    Dim src(1 To 3) As String
    Dim sc(1 To 3) As Double
    src(1) = "就業規則.pdf": sc(1) = 0.9
    src(2) = "就業規則.pdf": sc(2) = 0.5
    src(3) = "就業規則.pdf": sc(3) = 0.1

    Dim ord() As Long
    Dim n As Long: n = modSparse.DiversityOrder(src, sc, 3, ord)
    modTestRunner.Check "R27-F1-3_1資料のみなら順序は不変(1,2,3)", _
        (n = 3 And ord(1) = 1 And ord(2) = 2 And ord(3) = 3), _
        "n=" & n & " ord=" & JoinOrder24(ord, n)
End Sub

' 資料名が空の要素は代表になれない(どの資料の代表か決まらないため)。
' 空名だけのプールでは並べ替えが起きず、元の順序で全件返ること。
Private Sub TestDiversityOrderBlankSource24()
    Dim src(1 To 3) As String
    Dim sc(1 To 3) As Double
    src(1) = "": sc(1) = 0.9
    src(2) = "特約.doc": sc(2) = 0.5
    src(3) = "": sc(3) = 0.1

    Dim ord() As Long
    Dim n As Long: n = modSparse.DiversityOrder(src, sc, 3, ord)
    modTestRunner.Check "R27-F1-3_空の資料名は代表にならない(特約が先頭へ)", _
        (n = 3 And ord(1) = 2), "n=" & n & " ord=" & JoinOrder24(ord, n)
    modTestRunner.Check "R27-F1-3_空名の要素も落とさず全件返す", _
        OrderIsPermutation24(ord, 3), "ord=" & JoinOrder24(ord, 3)
End Sub

' F1-4: modExtractor.GarbleRatio — 化け判定の誤爆是正。
'   閾値は呼び出し側(DropGarbledPages)の 0.2 超で「化け」。
'   (c-1) 表のセル終端 Chr(7) を2割含む .doc 風テキストは OK 判定のまま
'   (c-2) キリル文字2割は従来どおり NG 判定
Private Sub TestGarbleRatioWordControls24()
    ' 日本語80字 + Chr(7)20個 = 全100字のうち2割がWordの表セル終端。
    ' 構造制御文字は分母からも外すので比率は0(=OK判定)でなければならない。
    Dim tbl As String
    tbl = Repeat24("保険金を支払わない場合", 8)          ' 11字×8 = 88字
    tbl = Left$(tbl, 80) & Repeat24(Chr$(7), 20)
    Dim rTbl As Double: rTbl = modExtractor.GarbleRatio(tbl)
    modTestRunner.Check "R27-F1-4_表のセル終端Chr(7)2割はOK判定(比率0)", _
        (rTbl = 0#), "ratio=" & Format$(rTbl, "0.000")
    modTestRunner.Check "R27-F1-4_表のセル終端Chr(7)2割は閾値0.2を超えない", _
        (rTbl <= 0.2), "ratio=" & Format$(rTbl, "0.000")

    ' 12種すべてを混ぜても同じ(1,2,5,7,11,12,14,19,20,21,30,31)。
    Dim mix As String
    mix = Left$(Repeat24("保険金を支払わない場合", 8), 80) & _
          Chr$(1) & Chr$(2) & Chr$(5) & Chr$(7) & Chr$(11) & Chr$(12) & _
          Chr$(14) & Chr$(19) & Chr$(20) & Chr$(21) & Chr$(30) & Chr$(31)
    modTestRunner.Check "R27-F1-4_Word構造制御12種を混ぜてもOK判定(比率0)", _
        (modExtractor.GarbleRatio(mix) = 0#), _
        "ratio=" & Format$(modExtractor.GarbleRatio(mix), "0.000")
End Sub

Private Sub TestGarbleRatioCyrillicStillNg24()
    ' 本文80字 + キリル20字 = 2割。除外対象ではないので比率0.2で、
    ' 「0.2を超える」ではないため境界の1字ぶんを足して超えさせる。
    Dim body As String: body = Left$(Repeat24("保険金を支払わない場合", 8), 80)
    Dim cyr As String: cyr = Repeat24(ChrW$(&H414), 20)      ' Д
    Dim r20 As Double: r20 = modExtractor.GarbleRatio(body & cyr)
    modTestRunner.Check "R27-F1-4_キリル2割はちょうど0.2(境界の値そのもの)", _
        (Abs(r20 - 0.2) < 0.000001), "ratio=" & Format$(r20, "0.0000")

    Dim r21 As Double: r21 = modExtractor.GarbleRatio(body & cyr & ChrW$(&H414))
    modTestRunner.Check "R27-F1-4_キリルが2割を超えたらNG判定(>0.2)", _
        (r21 > 0.2), "ratio=" & Format$(r21, "0.0000")

    ' ギリシャ文字も従来どおり化けとして数える。
    Dim grk As Double
    grk = modExtractor.GarbleRatio(Left$(body, 50) & Repeat24(ChrW$(&H3B1), 50))
    modTestRunner.Check "R27-F1-4_ギリシャ文字5割はNG判定(>0.2)", _
        (grk > 0.2), "ratio=" & Format$(grk, "0.0000")
End Sub

' AscW は U+8000 以降を負値で返すVBAの仕様がある(modSparse.NormalizeForSearch /
' modChrome / modClarify 等が同じ補正を持つ既存の作法)。旧実装はその負値を
' 「c < 32 = 制御文字」に掛けていたため、U+8000〜U+9FFF に住む常用漢字が
' まるごと化けとして数えられていた(険・関・金・通・者・除・認・説・語…)。
' 実測: 約款風の日本語166字のうち13.3%が誤ってbadに入る。
' 【重要な注記】LibreOffice Basic の AscW は同じ文字を正値で返すため、
' この符号補正の有無はLO実行テストでは差が出ない(下の3件はLO上では
' 補正を外しても通る)。それでも置くのは、実機VBAでのみ起きるこの誤爆が
' 再発したときに、実機のテスト実行(ブック同梱のmodTestRunner)で必ず
' 赤くなるようにするため。LOで捕まえられるのはWord構造制御文字側
' (TestGarbleRatioWordControls24)で、そちらは補正を外すと実際に落ちる。
Private Sub TestGarbleRatioHighKanjiNotGarbled24()
    Dim s As String
    s = "保険金額の関係者への通知は説明責任の観点から適切に行う。認識の相違を除く。"
    Dim r As Double: r = modExtractor.GarbleRatio(s)
    modTestRunner.Check "R27-F1-4_U+8000以降の常用漢字は化けではない(比率0)", _
        (r = 0#), "ratio=" & Format$(r, "0.0000")

    ' U+9FFF側の端(關 U+95DC 相当の常用字)も単独で化けにならない。
    modTestRunner.Check "R27-F1-4_『険』(U+967A)単独で化け判定にならない", _
        (modExtractor.GarbleRatio(Repeat24(ChrW$(&H967A), 60)) = 0#), _
        "ratio=" & Format$(modExtractor.GarbleRatio(Repeat24(ChrW$(&H967A), 60)), "0.0000")

    ' 本物の制御文字(Wordの構造制御文字ではないもの)は従来どおり化け。
    ' Chr(3)/Chr(4)/Chr(6) は除外リストに載っていない。
    Dim ctl As String
    ctl = Left$(s, 30) & Repeat24(Chr$(3), 20)
    modTestRunner.Check "R27-F1-4_除外外の制御文字Chr(3)は従来どおり化け(>0.2)", _
        (modExtractor.GarbleRatio(ctl) > 0.2), _
        "ratio=" & Format$(modExtractor.GarbleRatio(ctl), "0.0000")
End Sub

' F1-5: modExtractor.StripControlChars — 本文サニタイズ。
'   この文字列がそのまま embed / my_knowledge.norm_text / プロンプトへ流れる。
'   タブ・改行(9/10/13)は行と列の区切りとして下流が読むので必ず残す。
'   波3-16: 除去ではなく【半角スペースへ置換】(表のセルが直結して実在しない
'   語になるのを防ぐ)。連続する空白は1つへ圧縮。
Private Sub TestStripControlChars24()
    Dim outS As String

    ' 表のセル終端Chr(7)はスペースになる(セルが直結しない)。
    Dim n1 As Long
    n1 = modExtractor.StripControlChars("項目" & Chr$(7) & "値" & Chr$(7), outS)
    modTestRunner.Check "R27-F1-5_Chr(7)は2字とも置換される", (n1 = 2), "replaced=" & n1
    modTestRunner.Check "R27波3-16_セルは直結せず『項目 値 』になる", _
        (outS = "項目 値 "), "out=[" & outS & "]"

    ' 連続した制御文字はスペース1つへ(空白だけが延々と続かない)。
    Dim n1b As Long
    n1b = modExtractor.StripControlChars("A" & Chr$(7) & Chr$(7) & Chr$(7) & "B", outS)
    modTestRunner.Check "R27波3-16_連続制御文字は3字とも数え空白1つへ圧縮", _
        (n1b = 3 And outS = "A B"), "replaced=" & n1b & " out=[" & outS & "]"

    ' 先頭の制御文字は行頭に無意味な空白を作らない。
    modTestRunner.Check "R27波3-16_先頭の制御文字はスペースを生まない", _
        (modExtractor.StripControlChars(Chr$(7) & "本文", outS) = 1 And outS = "本文"), _
        "out=[" & outS & "]"

    ' 既存スペースとの連続もまとめて1つ(置換が起きたときだけ働く)。
    modTestRunner.Check "R27波3-16_制御文字の隣の既存スペースも1つにまとまる", _
        (modExtractor.StripControlChars("A " & Chr$(7) & " B", outS) = 1 And outS = "A B"), _
        "out=[" & outS & "]"

    ' 制御文字が無ければ、連続スペースは1つも触らない(副作用ゼロ)。
    modTestRunner.Check "R27波3-16_制御文字が無ければ連続スペースは温存", _
        (modExtractor.StripControlChars("A  B", outS) = 0 And outS = "A  B"), _
        "out=[" & outS & "]"

    ' 9/10/13 は残す(チャンク分割と見出し判定が読んでいる)。
    Dim keep As String: keep = "行1" & vbTab & "列2" & vbCrLf & "行3" & Chr$(10)
    Dim n2 As Long: n2 = modExtractor.StripControlChars(keep, outS)
    modTestRunner.Check "R27-F1-5_タブ/CR/LFは1字も落とさない", (n2 = 0), "removed=" & n2
    modTestRunner.Check "R27-F1-5_落とすものが無ければ文字列は同一", _
        (outS = keep), "out=[" & outS & "]"

    ' F1-4で除外した12種は全部落ちる(除外=化け判定に数えないだけで、本文には残さない)。
    Dim ctl As String
    ctl = "本" & Chr$(1) & Chr$(2) & Chr$(5) & Chr$(7) & Chr$(11) & Chr$(12) & _
          "文" & Chr$(14) & Chr$(19) & Chr$(20) & Chr$(21) & Chr$(30) & Chr$(31)
    Dim n3 As Long: n3 = modExtractor.StripControlChars(ctl, outS)
    modTestRunner.Check "R27-F1-5_Word構造制御12種は全て置換される", (n3 = 12), "replaced=" & n3
    modTestRunner.Check "R27波3-16_本文2字が空白1つずつで区切られる", _
        (outS = "本 文 "), "out=[" & outS & "]"

    ' 空文字は空文字のまま(0字除去)。
    modTestRunner.Check "R27-F1-5_空文字は0字除去で空のまま", _
        (modExtractor.StripControlChars("", outS) = 0 And outS = ""), "out=[" & outS & "]"

    ' 日本語(U+8000以降を含む)は1字も落とさない。AscWの符号補正が抜けると
    ' 「険」「金」等が制御文字扱いで消え、本文が丸ごと壊れる。
    Dim jp As String: jp = "保険金額の関係者へ通知する"
    Dim n4 As Long: n4 = modExtractor.StripControlChars(jp, outS)
    modTestRunner.Check "R27-F1-5_日本語本文は1字も落とさない", _
        (n4 = 0 And outS = jp), "removed=" & n4 & " out=[" & outS & "]"
End Sub

' F1-9: modClarify.MentionsSourceName — 逆質問の名指し誤判定の緩和。
'   規則(3)「資料名の側が質問の主要語を含む」は、該当資料が2件以上あるなら
'   利用者はどちらの資料かを指定していない=聞き返すべき場面。
'   規則(1)(2)(利用者が資料名を実際に打った証拠)は1件でも従来どおり名指し。
Private Sub TestMentionsSourceNameMultiHit24()
    ' (3)が1件だけ当たる: 従来どおり名指し扱い(聞き返さない)。
    modTestRunner.Check "R27-F1-9_主要語を含む資料が1件なら名指し扱い(従来どおり)", _
        modClarify.MentionsSourceName("火災保険の免責は?", "火災保険約款.pdf|自動車保険約款.pdf"), _
        "1件でも聞き返しを止めないと、名指しした人にもう一度選ばせることになる"

    ' (3)が2件当たる: 名指しではない(=逆質問を出してよい)。ここがR27の変更点。
    modTestRunner.Check "R27-F1-9_主要語を含む資料が2件なら名指しではない", _
        (modClarify.MentionsSourceName("団体保険の免責は?", "団体保険約款.pdf|団体保険特約.pdf") = False), _
        "2件あるのに聞き返しを止めると、どちらの資料か永久に確認できない"

    ' 3件でも同じ(2件以上は一律で名指しではない)。
    ' 資料名は「主要語(連続する漢字/カタカナの最長列)が質問文に出てこない」
    ' ものを選ぶ。例えば『団体保険しおり』は主要語が「団体保険」になり、
    ' 規則(2)(資料名の主要語が質問文に含まれる)で先に名指し扱いになるため
    ' この検証には使えない(規則(2)はR27の変更対象外)。
    modTestRunner.Check "R27-F1-9_主要語を含む資料が3件でも名指しではない", _
        (modClarify.MentionsSourceName("団体保険の免責は?", _
            "団体保険約款.pdf|団体保険特約.pdf|団体保険規程.pdf") = False), ""

    ' 規則(1): 質問文が資料名まるごとを含むなら、他に何件並んでいても名指し。
    modTestRunner.Check "R27-F1-9_規則(1)資料名まるごとは2件並んでいても名指し", _
        modClarify.MentionsSourceName("団体保険約款の免責は?", "団体保険約款.pdf|団体保険特約.pdf"), _
        "利用者が資料名を打っている以上、聞き返す理由が無い"

    ' 無関係な質問は従来どおり名指しではない。
    modTestRunner.Check "R27-F1-9_主要語がどの資料名にも無ければ名指しではない", _
        (modClarify.MentionsSourceName("駐車場の使い方は?", "団体保険約款.pdf|団体保険特約.pdf") = False), ""
End Sub

' F2-3: modViewport2.PadRowDelta — 埋め草の差分計算と上限クランプ。
'   境界の最終行は modViewport.RowAtFloor(切り下げ)で決まるため、その下端は
'   必ず窓高より上に残る。差のぶんだけ最終行を高くして塗りを窓下端へ届かせる
'   のが埋め草で、この純関数は「いくつ足すか」だけを決める。
'   上限(48pt=1行ぶんを大きく超える差)は構造問題なので埋めない ―― 埋めると
'   症状だけが消えて原因を観測できなくなる。仕様の境界0/10/48/49を固定する。
Private Sub TestPadRowDeltaBoundary24()
    Const MX As Double = 48

    ' 差0: 既に窓下端へ届いている=足さない。
    modTestRunner.Check "R27-F2-3_差0なら足さない", _
        (modViewport2.PadRowDelta(700, 700, MX) = 0), _
        "d=" & modViewport2.PadRowDelta(700, 700, MX)

    ' 差10: そのまま10pt足す(戻り値を実数で固定=「0を返すだけ」の実装を落とす)。
    modTestRunner.Check "R27-F2-3_差10なら10pt足す", _
        (modViewport2.PadRowDelta(690, 700, MX) = 10), _
        "d=" & modViewport2.PadRowDelta(690, 700, MX)

    ' 差48(上限ちょうど): まだ埋める。クランプは「超えたら」なので48は含む。
    modTestRunner.Check "R27-F2-3_差48(上限ちょうど)は埋める", _
        (modViewport2.PadRowDelta(652, 700, MX) = 48), _
        "d=" & modViewport2.PadRowDelta(652, 700, MX)

    ' 差49(上限超): 何もしない。
    modTestRunner.Check "R27-F2-3_差49(上限超)は何もしない", _
        (modViewport2.PadRowDelta(651, 700, MX) = 0), _
        "d=" & modViewport2.PadRowDelta(651, 700, MX)

    ' ネガティブ: 既に窓を超えている(差が負)なら足さない。足すと境界が
    ' 窓を1行ぶん超え、消したはずの縦スクロールが生き返る。
    modTestRunner.Check "R27-F2-3_窓を超えていたら足さない", _
        (modViewport2.PadRowDelta(760, 700, MX) = 0), _
        "d=" & modViewport2.PadRowDelta(760, 700, MX)

    ' 丸め未満(1pt未満)の差は触らない。触ると「足す→RowHeightの丸めで少し
    ' 足りない→また足す」で毎描画わずかに伸び続ける。
    modTestRunner.Check "R27-F2-3_1pt未満の差は触らない", _
        (modViewport2.PadRowDelta(699.5, 700, MX) = 0), _
        "d=" & modViewport2.PadRowDelta(699.5, 700, MX)

    ' 恒真防止: 上の「=0」群は常に0を返す実装でも全部通ってしまう。上限を
    ' またぐ隣り合う2点で戻り値が実際に変わることを直接確かめる。
    modTestRunner.Check "R27-F2-3_上限の前後で戻り値が変わる(恒真でない)", _
        (modViewport2.PadRowDelta(652, 700, MX) <> modViewport2.PadRowDelta(651, 700, MX)), _
        "48pt=" & modViewport2.PadRowDelta(652, 700, MX) & _
        " / 49pt=" & modViewport2.PadRowDelta(651, 700, MX)

    ' 上限は引数なので、呼び出し側が別の上限を渡したらそれに従う
    ' (定数の焼き付きではないことの確認)。
    modTestRunner.Check "R27-F2-3_上限は引数で決まる(maxPad=8なら差10は埋めない)", _
        (modViewport2.PadRowDelta(690, 700, 8) = 0), _
        "d=" & modViewport2.PadRowDelta(690, 700, 8)
End Sub

Private Function Repeat24(ByVal unit As String, ByVal times As Long) As String
    Dim sb As String
    Dim i As Long
    For i = 1 To times
        sb = sb & unit
    Next i
    Repeat24 = sb
End Function

Private Function OrderIsPermutation24(ByRef ord() As Long, ByVal n As Long) As Boolean
    Dim seen() As Boolean: ReDim seen(1 To n)
    Dim i As Long
    For i = 1 To n
        If ord(i) < 1 Or ord(i) > n Then Exit Function
        If seen(ord(i)) Then Exit Function
        seen(ord(i)) = True
    Next i
    OrderIsPermutation24 = True
End Function

Private Function JoinOrder24(ByRef ord() As Long, ByVal n As Long) As String
    Dim sb As String
    Dim i As Long
    For i = 1 To n
        If LenB(sb) > 0 Then sb = sb & ","
        sb = sb & ord(i)
    Next i
    JoinOrder24 = sb
End Function

Private Function BadgeIdExists24(ByRef ids() As String, ByVal target As String) As Boolean
    Dim i As Long
    For i = LBound(ids) To UBound(ids)
        If StrComp(Trim$(ids(i)), target, vbTextCompare) = 0 Then
            BadgeIdExists24 = True
            Exit Function
        End If
    Next i
End Function

Private Function IndexOf24(ByRef ids() As String, ByVal target As String) As Long
    IndexOf24 = -1
    Dim i As Long
    For i = LBound(ids) To UBound(ids)
        If StrComp(Trim$(ids(i)), target, vbTextCompare) = 0 Then
            IndexOf24 = i
            Exit Function
        End If
    Next i
End Function

' サロゲートペア(絵文字)の有無を判定する。AscWは符号付きLongを返すため
' 負値(&H8000以上)を補正してから上位/下位サロゲート範囲と比較する。
Private Function ContainsSurrogate24(ByVal s As String) As Boolean
    Dim i As Long, c As Long
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536
        ' &H8000以上の16進リテラルは"&"サフィックス無しだとIntegerの符号付き
        ' (負値)として解釈される罠がある。上のc補正(正値0-65535)と比較が
        ' 噛み合うよう、境界値は必ずLong強制の"&"サフィックス付きで書く。
        If c >= &HD800& And c <= &HDFFF& Then
            ContainsSurrogate24 = True
            Exit Function
        End If
    Next i
End Function

' ============================================================================
Public Sub RunAll24()
    On Error GoTo Count16Fail24
    TestBadgeCatalogCount16
NextNewIds24:
    On Error GoTo NewIdsFail24
    TestBadgeCatalogNewIdsExist
NextNoDup24:
    On Error GoTo NoDupFail24
    TestBadgeCatalogNoDuplicateIds
NextThreshold24:
    On Error GoTo ThresholdFail24
    TestNewBadgeConditionsMentionThreshold
NextNoEmoji24:
    On Error GoTo NoEmojiFail24
    TestBadgeTitlesHaveNoSurrogate24
NextCapKey24:
    On Error GoTo CapKeyFail24
    TestCapKeyScoreBoundary24
NextLenClamp24:
    On Error GoTo LenClampFail24
    TestKeyLenWeightClamp24
NextKeyScoreCap24:
    On Error GoTo KeyScoreCapFail24
    TestKeyScoreIsCapped24
NextDiv24:
    On Error GoTo DivFail24
    TestDiversityOrderMinoritySurvives24
NextDivSingle24:
    On Error GoTo DivSingleFail24
    TestDiversityOrderSingleSource24
NextDivBlank24:
    On Error GoTo DivBlankFail24
    TestDiversityOrderBlankSource24
NextGarbleWord24:
    On Error GoTo GarbleWordFail24
    TestGarbleRatioWordControls24
NextGarbleCyr24:
    On Error GoTo GarbleCyrFail24
    TestGarbleRatioCyrillicStillNg24
NextGarbleKanji24:
    On Error GoTo GarbleKanjiFail24
    TestGarbleRatioHighKanjiNotGarbled24
NextStrip24:
    On Error GoTo StripFail24
    TestStripControlChars24
NextMention24:
    On Error GoTo MentionFail24
    TestMentionsSourceNameMultiHit24
NextPadRow24:
    On Error GoTo PadRowFail24
    TestPadRowDeltaBoundary24
NextRun25:
    On Error GoTo Run25Fail24
    modTestsPure25.RunAll25
NextDone24:
    On Error GoTo 0
    Exit Sub

Count16Fail24:
    modTestRunner.Check "TestBadgeCatalogCount16(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextNewIds24
NewIdsFail24:
    modTestRunner.Check "TestBadgeCatalogNewIdsExist(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextNoDup24
NoDupFail24:
    modTestRunner.Check "TestBadgeCatalogNoDuplicateIds(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextThreshold24
ThresholdFail24:
    modTestRunner.Check "TestNewBadgeConditionsMentionThreshold(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextNoEmoji24
NoEmojiFail24:
    modTestRunner.Check "TestBadgeTitlesHaveNoSurrogate24(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextCapKey24
CapKeyFail24:
    modTestRunner.Check "TestCapKeyScoreBoundary24(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextLenClamp24
LenClampFail24:
    modTestRunner.Check "TestKeyLenWeightClamp24(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextKeyScoreCap24
KeyScoreCapFail24:
    modTestRunner.Check "TestKeyScoreIsCapped24(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDiv24
DivFail24:
    modTestRunner.Check "TestDiversityOrderMinoritySurvives24(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDivSingle24
DivSingleFail24:
    modTestRunner.Check "TestDiversityOrderSingleSource24(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDivBlank24
DivBlankFail24:
    modTestRunner.Check "TestDiversityOrderBlankSource24(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGarbleWord24
GarbleWordFail24:
    modTestRunner.Check "TestGarbleRatioWordControls24(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGarbleCyr24
GarbleCyrFail24:
    modTestRunner.Check "TestGarbleRatioCyrillicStillNg24(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGarbleKanji24
GarbleKanjiFail24:
    modTestRunner.Check "TestGarbleRatioHighKanjiNotGarbled24(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextStrip24
StripFail24:
    modTestRunner.Check "TestStripControlChars24(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextMention24
MentionFail24:
    modTestRunner.Check "TestMentionsSourceNameMultiHit24(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextPadRow24
PadRowFail24:
    modTestRunner.Check "TestPadRowDeltaBoundary24(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRun25
Run25Fail24:
    modTestRunner.Check "modTestsPure25.RunAll25(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone24
End Sub
