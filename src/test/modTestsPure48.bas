Attribute VB_Name = "modTestsPure48"
Option Explicit

' R44 の検算に使う定数(モジュール宣言部。プロシージャより後に置くと実機VBAで
' コンパイルエラーになる=R42 で一度踏んだ型)。
Private Const SPARSE_WEIGHT_44 As Double = 0.06   ' modRetrieve.bas:75 と同値
Private Const CAP_DEFAULT_44 As Double = 3#       ' build_config_rows の既定
Private Const WARN_THRESHOLD_44 As Double = 0.3   ' low_hit_warn_score の既定
Private Const CONF_THRESHOLD_44 As Double = 0.55  ' confidence_score_x100/100


' ============================================================================
' modTestsPure48 - R43 §4(実機報告「Excel の出典が A 列ばかりになる」)の
'   modExtractorExcel.RowTextFrom セル単位番地付けの純ロジック回帰。
'   modTestRunner.RunAllPureTests から直接呼ばれる独立の別枝
'   (modTestsPure31〜45と同型)。既存の RowTextFrom 検査(空セルの位置保持・
'   末尾トリム等)は modTestsPure33(今回の触ってよいファイル範囲外)に
'   既にあり、R43 で追加した rowIdx/baseCol(いずれも Optional・既定0)の
'   挙動だけをここで固定する。4引数のみで呼ぶ既存呼び出し(modTestsPure33)
'   が番地無しの従来出力のまま変わらないことも、後方互換の検算として
'   1件含める(CLAUDE.md §9 のシグネチャ変更対応をOptional化で満たした
'   ことの直接証跡)。
' ----------------------------------------------------------------------------
' 【手計算の根拠】
'   CellAddressOf(colIdx,rowIdx) は 1→A, 26→Z, 27→AA, 28→AB の26進表記
'   (modTestsPure43 の R40F3_A1/Z10/AA6 で既に固定済み)。RowPrefix(c,r) は
'   "[" & CellAddressOf(c,r) & "] "。RowTextFrom は値のあるセル c だけに
'   RowPrefix(baseCol+c-1, rowIdx) & 値 を置き、空セルは空文字のまま残す
'   (位置保持。R33波3 W3-1 の規則は変えていない)。区切りは全セル共通で
'   vbTab(セル区切り文字は変更なし)。
'
'   T1(値のあるセルだけに番地・空セルは番地無しで位置保持):
'     3列中2列目だけ空。baseCol=1(A列起点)・rowIdx=5。
'     c1="山田"→[A5] 山田 / c2=Empty→"" / c3="1234"→[C5] 1234。
'     Join結果=「[A5] 山田」&TAB&""&TAB&「[C5] 1234」
'     =「[A5] 山田<TAB><TAB>[C5] 1234」。
'
'   T2(A列が空でB列起点のシート): UsedRangeの左端がB列という状況を
'     baseCol=2で再現(1列だけの行・rowIdx=10)。
'     colIdx=baseCol+1-1=2→"B"。結果=「[B10] X」(Aではなく必ずBになる)。
'
'   T3(複数列で列が1つずつ進む): 4列すべて値あり・baseCol=1・rowIdx=7。
'     colIdx=1,2,3,4→A,B,C,D。結果=
'     「[A7] v1」&TAB&「[B7] v2」&TAB&「[C7] v3」&TAB&「[D7] v4」。
'
'   T4(行番号が正しい・複数行): 1列3行・baseCol=5(E列)。
'     AppendBlockLinesの実際の呼び方(firstRow+r-1)を模して、r=1→
'     rowIdx=20、r=2→21、r=3→22を渡す。結果はE列のまま行番号だけが
'     20→21→22と進む(「[E20] p」/「[E21] q」/「[E22] r」)。
'
'   T5(26列超・AA/AB): 2列・baseCol=27(AA列起点)・rowIdx=3。
'     colIdx=27,28→AA,AB(modTestsPure43のCellAddressOf(27,6)="AA6"と
'     同じ26進表記)。結果=「[AA3] x1」&TAB&「[AB3] x2」。
'
'   T6(空行): 2列とも空。rowIdx=50・baseCol=3を渡しても値が無いので
'     hasCell=Falseかつ戻り値は空文字(番地は一切出ない)。
'
'   T7(後方互換・4引数呼び出しは番地無しのまま): rowIdx/baseColを省略
'     (既定0)して1列1行"foo"を渡すと、RowPrefix(0,0)が
'     CellAddressOfのcolIdx<1 Or rowIdx<1ガードで空文字を返すため、
'     結果はプレーンな「foo」のまま(modTestsPure33が4引数のまま
'     RowTextFromを呼び続けても出力もタブ数も変わらないことの直接検算)。
' ============================================================================

Private Sub ChkStr48(ByVal label As String, ByVal got As String, ByVal want As String)
    modTestRunner.Check "R43-48-" & label, (StrComp(got, want, vbBinaryCompare) = 0), _
        "実際=[" & got & "] 期待=[" & want & "]"
End Sub

Private Sub ChkBool48(ByVal label As String, ByVal got As Boolean, ByVal want As Boolean)
    modTestRunner.Check "R43-48-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

' ---- T1: 値のあるセルだけに番地(空セルは番地無しで位置保持) --------------
Private Sub TestCellOnlyAddress48()
    Dim arr() As Variant: ReDim arr(1 To 1, 1 To 3)
    arr(1, 1) = "山田"
    arr(1, 3) = "1234"

    Dim has As Boolean
    Dim got As String: got = modExtractorExcel.RowTextFrom(arr, 1, 3, has, 5, 1)
    ChkBool48 "T1_hasCell", has, True
    ChkStr48 "T1_値のあるセルだけに番地", got, "[A5] 山田" & vbTab & vbTab & "[C5] 1234"

    ' T1b(レビュー R43 m2): 数式が返した空文字("")には番地を付けない。
    ' 付けると [B5] だけが残り、値の無いセルを AI が引用しうる。
    ' ただし【行は残す】(hasCell=True)。R33 W3-1「数式ブランクの行は残す」を
    ' 壊さないため、番地の有無と行の存否は別の判定にしてある
    ' (一緒にすると、数式で空になった行が丸ごと落ちる ―― 実際に LO の
    '  R33-W3-1 が落ちて気付いた回帰)。
    Dim arr2() As Variant: ReDim arr2(1 To 1, 1 To 3)
    arr2(1, 1) = "山田"
    arr2(1, 2) = ""
    arr2(1, 3) = "1234"
    Dim has2 As Boolean
    Dim got2 As String: got2 = modExtractorExcel.RowTextFrom(arr2, 1, 3, has2, 5, 1)
    ChkBool48 "T1b_数式の空文字でも行は残る", has2, True
    ChkStr48 "T1b_数式の空文字には番地を付けない", got2, "[A5] 山田" & vbTab & vbTab & "[C5] 1234"

    ' T1c: 全セルが数式ブランクの行も残す(R33 W3-1 の直接検算)。
    Dim arr3() As Variant: ReDim arr3(1 To 1, 1 To 2)
    arr3(1, 1) = ""
    arr3(1, 2) = ""
    Dim has3 As Boolean
    Dim got3 As String: got3 = modExtractorExcel.RowTextFrom(arr3, 1, 2, has3, 7, 1)
    ChkBool48 "T1c_全セル数式ブランクでも行は残る", has3, True
    ChkStr48 "T1c_番地は付かない", got3, "" & vbTab & ""
End Sub

' ---- T2: A列が空でB列起点のシートで正しくBが出る --------------------------
Private Sub TestBColStart48()
    Dim arr() As Variant: ReDim arr(1 To 1, 1 To 1)
    arr(1, 1) = "X"

    Dim has As Boolean
    Dim got As String: got = modExtractorExcel.RowTextFrom(arr, 1, 1, has, 10, 2)
    ChkBool48 "T2_hasCell", has, True
    ChkStr48 "T2_A列空でB起点はBが出る", got, "[B10] X"
End Sub

' ---- T3: 複数列の行で列が1つずつ進む --------------------------------------
Private Sub TestColAdvance48()
    Dim arr() As Variant: ReDim arr(1 To 1, 1 To 4)
    arr(1, 1) = "v1": arr(1, 2) = "v2": arr(1, 3) = "v3": arr(1, 4) = "v4"

    Dim has As Boolean
    Dim got As String: got = modExtractorExcel.RowTextFrom(arr, 1, 4, has, 7, 1)
    ChkStr48 "T3_列が1つずつ進む", got, _
        "[A7] v1" & vbTab & "[B7] v2" & vbTab & "[C7] v3" & vbTab & "[D7] v4"
End Sub

' ---- T4: 行番号が正しい(AppendBlockLinesの firstRow+r-1 を模した複数行) ---
Private Sub TestRowNumberAdvance48()
    Dim arr() As Variant: ReDim arr(1 To 3, 1 To 1)
    arr(1, 1) = "p": arr(2, 1) = "q": arr(3, 1) = "r"

    Dim has As Boolean
    Dim firstRow As Long: firstRow = 20
    Dim got1 As String, got2 As String, got3 As String
    got1 = modExtractorExcel.RowTextFrom(arr, 1, 1, has, firstRow + 1 - 1, 5)
    got2 = modExtractorExcel.RowTextFrom(arr, 2, 1, has, firstRow + 2 - 1, 5)
    got3 = modExtractorExcel.RowTextFrom(arr, 3, 1, has, firstRow + 3 - 1, 5)
    ChkStr48 "T4_行1", got1, "[E20] p"
    ChkStr48 "T4_行2", got2, "[E21] q"
    ChkStr48 "T4_行3", got3, "[E22] r"
End Sub

' ---- T5: 26列を超える列(AA・AB)の番地 --------------------------------------
Private Sub TestBeyondZ48()
    Dim arr() As Variant: ReDim arr(1 To 1, 1 To 2)
    arr(1, 1) = "x1": arr(1, 2) = "x2"

    Dim has As Boolean
    Dim got As String: got = modExtractorExcel.RowTextFrom(arr, 1, 2, has, 3, 27)
    ChkStr48 "T5_AA_AB", got, "[AA3] x1" & vbTab & "[AB3] x2"
End Sub

' ---- T6: 空行の扱い(値が無ければ番地を渡されても何も出ない) --------------
Private Sub TestEmptyRow48()
    Dim arr() As Variant: ReDim arr(1 To 1, 1 To 2)
    ' 両セルともEmptyのまま(代入しない)。

    Dim has As Boolean
    Dim got As String: got = modExtractorExcel.RowTextFrom(arr, 1, 2, has, 50, 3)
    ChkBool48 "T6_hasCell", has, False
    ChkStr48 "T6_空行は空文字", got, ""
End Sub

' ---- T7: 後方互換(rowIdx/baseCol省略時は番地無しの従来出力のまま) --------
Private Sub TestBackwardCompat48()
    Dim arr() As Variant: ReDim arr(1 To 1, 1 To 1)
    arr(1, 1) = "foo"

    Dim has As Boolean
    Dim got As String: got = modExtractorExcel.RowTextFrom(arr, 1, 1, has)
    ChkBool48 "T7_hasCell", has, True
    ChkStr48 "T7_4引数呼び出しは番地無し", got, "foo"
End Sub

' ============================================================================
' R44 追加分 - 表示専用の注意書き(modMode.DisplayNotes)と、キーワード加点の
'   上限が安全装置を殺さないことの算術検算。
' ----------------------------------------------------------------------------
' 【なぜこの検算が要るか】
'   検索スコアは「ベクトル類似度 + KeyScore*0.06」の合計(modRetrieve.bas:280)。
'   信頼度バッジ(modAsk.LastConfidence)は合計を confidence_score_x100/100 と、
'   低関連度警告(modAskRetrieve)は合計を low_hit_warn_score と比べる。
'   つまり【加点の最大値が両しきい値未満】でなければ、意味が一致していなくても
'   キーワードが数個かぶるだけで両方を突破し、安全装置が常時オフになる。
'   R44 実測: 上限10(旧既定)のとき加点は最大0.6で、0.3も0.55も無条件に超えた
'   (10,996かけらの本棚で、1質問あたり中央403かけらが0.55超え=🟢常時点灯)。
'   ここは「上限×0.06 < 0.3」という不等式そのものをテストで固定する。
'   config を 5 以上へ戻すとこのテストが落ちる=気付ける、という関所にする。
Private Sub ChkStr44(ByVal label As String, ByVal got As String, ByVal want As String)
    modTestRunner.Check "R44-48-" & label, (StrComp(got, want, vbBinaryCompare) = 0), _
        "実際=[" & got & "] 期待=[" & want & "]"
End Sub

Private Sub ChkBool44(ByVal label As String, ByVal got As Boolean, ByVal want As Boolean)
    modTestRunner.Check "R44-48-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

' ---- G1: 加点の上限が両しきい値を超えないこと(不等式そのものの固定) --------
Private Sub TestBoostBelowThresholds44()
    Dim maxBoost As Double
    maxBoost = modSparse.CapKeyScore(1000#, CAP_DEFAULT_44) * SPARSE_WEIGHT_44
    ChkBool44 "G1a_加点だけでは低関連度警告のしきい値に届かない", (maxBoost < WARN_THRESHOLD_44), True
    ChkBool44 "G1b_加点だけでは信頼度バッジのしきい値に届かない", (maxBoost < CONF_THRESHOLD_44), True
    ' 旧既定10なら両方を突破していたことも同時に固定する(回帰の向きを明示)。
    Dim oldBoost As Double
    oldBoost = modSparse.CapKeyScore(1000#, 10#) * SPARSE_WEIGHT_44
    ChkBool44 "G1c_旧既定10は警告のしきい値を突破していた", (oldBoost > WARN_THRESHOLD_44), True
    ChkBool44 "G1d_旧既定10は信頼度のしきい値も突破していた", (oldBoost > CONF_THRESHOLD_44), True
End Sub

' ---- G2: 常設ガードは末尾に1つだけ付く ------------------------------------
Private Sub TestGuardSuffix44()
    Dim body As String: body = "結論です。[本棚:規約集 p.1]"
    Dim got As String
    got = modMode.DisplayNotes(body, 0.8, True, 0.3, True)
    ChkStr44 "G2a_高スコアなら警告無しでガードだけ", got, _
        body & vbLf & vbLf & modMode.GuardNoteText()
    ' 二度通しても増えない(復元表示・再描画で本文を通し直す経路がある)
    ChkStr44 "G2b_二度通しても増えない", modMode.DisplayNotes(got, 0.8, True, 0.3, True), got
    ' guardOn=False で消せる
    ChkStr44 "G2c_設定で消せる", modMode.DisplayNotes(body, 0.8, True, 0.3, False), body
    ' 空の本文にはガードを付けない(空バブルに注意書きだけ出るのを防ぐ)
    ChkStr44 "G2d_空本文には付けない", modMode.DisplayNotes("", 0.8, True, 0.3, True), ""
End Sub

' ---- G3: 低関連度の⚠は先頭・ガードは末尾(順序と共存) ----------------------
Private Sub TestWarnAndGuard44()
    Dim body As String: body = "たぶんこうです。"
    Dim got As String
    got = modMode.DisplayNotes(body, 0.1, True, 0.3, True)
    ChkStr44 "G3a_低スコアは低関連度の印が先頭でガードが末尾", got, _
        modMode.LowHitNoteText() & vbLf & vbLf & body & vbLf & vbLf & modMode.GuardNoteText()
    ' 俯瞰ターン(warnOn=False)は⚠を出さないが、ガードは出す
    ChkStr44 "G3b_俯瞰ターンは低関連度の印を出さずガードだけ", _
        modMode.DisplayNotes(body, 0#, False, 0.3, True), _
        body & vbLf & vbLf & modMode.GuardNoteText()
    ' threshold<=0 は警告の無効化(既存仕様)。ガードは無効化されない。
    ChkStr44 "G3c_しきい値0は警告のみ無効", _
        modMode.DisplayNotes(body, 0#, True, 0#, True), _
        body & vbLf & vbLf & modMode.GuardNoteText()
End Sub

' ---- G4: 装飾側が拾う先頭文字の契約(変えたら両方直す、の固定) --------------
Private Sub TestNoteMarks44()
    ChkStr44 "G4a_ガードは※で始まる", Left$(modMode.GuardNoteText(), 1), ChrW(&H203B)
    ChkStr44 "G4b_低関連度の印は所定の記号で始まる", Left$(modMode.LowHitNoteText(), 1), ChrW(&H26A0)
End Sub

' ---- G6: 精査モードの既定経路にドメインガードが入っていること ---------------
' R44: modAskOnePass は thorough_onepass=on(出荷既定)のとき精査モードで必ず
' 通る経路で、成功したら modAskThorough は Exit Function する(4段へ行かない)。
' ここに DomainGuardInstruction が無いと、いちばん外したくないモードでだけ
' 数値の捏造ガードが外れる ―― R44 で実際にそうなっていた。
' FOLLOWUP 行より前にあることも見る(modTestsPure43 が末尾を固定しているため)。
Private Sub TestOnePassHasGuard44()
    Dim p As String
    p = modAskOnePass.BuildOnePassPrompt("等級はどうなりますか", "## 本棚抜粋" & vbLf & "本文", _
                                         "", "", "日本語", False, False)
    ChkBool44 "G6a_数値の厳格性が入っている", (InStr(p, "【数値の厳格性】") > 0), True
    ChkBool44 "G6b_確認マークが入っている", (InStr(p, "(要確認)") > 0), True
    ChkBool44 "G6c_断定の禁止が入っている", (InStr(p, "【断定の禁止】") > 0), True
    ChkBool44 "G6d_ガードはFOLLOWUPより前", _
        (InStr(p, "【数値の厳格性】") < InStr(p, "[[FOLLOWUP:")), True
    ' Quick と同じ文言であること(複製ではなく単一情報源を使っている証跡)
    ChkBool44 "G6e_文言はmodPromptsの単一情報源", _
        (InStr(p, modPrompts.DomainGuardInstruction()) > 0), True
End Sub

' ---- G7: Excel のシート見出しが「章」として認識されること(R45) --------------
' R44 の解剖: Excel は行がタブ区切り(表行)か、値1個でも [A5] 前置があるため
' ClassifyLine が章にも節にも一度も分類せず、ブック全体が1ブロックになっていた。
' → chapter/section が空 → breadcrumb が「【〔資料〕】」だけ → section_path が空
' → 章要約・俯瞰・参照展開・条番号保証が Excel に対して全滅。
' シート1行目を「# シート: 名前」にして【シート＝章】を立てるのが R45 の直し。
' 抽出側の実物(SheetHeadingLine)を分類器に食わせるので、表記を変えると落ちる。
Private Sub TestSheetHeading45()
    Dim h As String: h = modExtractorExcel.SheetHeadingLine("自動車保険(個人)_1")
    ChkBool44 "G7a_シート見出しは章(1)に分類される", (modChunker.ClassifyLine(h) = 1), True
    ' 見出しの印は StripHeadingMark が外せる "#" であること(breadcrumb が
    ' 「【〔資料〕 > 【シート: 名前】】」と二重括弧にならないための契約)
    ChkStr44 "G7b_印はシャープと空白", Left$(h, 2), "# "
    ChkBool44 "G7c_シート名が入っている", (InStr(h, "自動車保険(個人)_1") > 0), True
    ' データ行は従来どおり。2セル以上=表行(4)、値1個=本文(0)。どちらも見出しでない
    ChkBool44 "G7d_複数セルの行は表行(4)", _
        (modChunker.ClassifyLine("[A5] 山田" & vbTab & "[B5] 1234") = 4), True
    ChkBool44 "G7e_単一セルの行は本文(0)", _
        (modChunker.ClassifyLine("[A5] 見出しらしき文字列") = 0), True
    ' 章が立つと breadcrumb に入る(section_path の材料になる)
    Dim crumb As String
    crumb = modChunker.BuildBreadcrumb(modChunker.CRUMB_PLACEHOLDER, "シート: 自動車保険(個人)_1", "")
    ChkBool44 "G7f_breadcrumbにシート名が入る", (InStr(crumb, "シート: 自動車保険(個人)_1") > 0), True
    ChkStr44 "G7g_section_pathはシート名になる", _
        modChunkMeta.ExtractSectionPath(crumb & vbLf & "本文"), "シート: 自動車保険(個人)_1"
End Sub

' ---- G8: 一般モードへ引き継いだ文脈の扱い(R45) ------------------------------
' 橋渡しヘッダーがあるターンだけ CarryRules を足す。無条件に入れると
' 「存在しない会話」を指す指示が残る(modGenPipe.bas:409-415 の実害記録)。
Private Sub TestCarryGuard45()
    Dim hdr As String: hdr = modConvBridge.WithBridgeHeader("社内ナレッジ検索", "回答本文")
    ChkBool44 "G8a_橋渡し直後は検出できる", modConvBridge.HasBridgeHeader(hdr), True
    ' 一般モードで1往復進むと新しい回答が先頭へ積まれ、ヘッダーは途中へ移る。
    ' それでも運搬内容は送られ続けるので、検出も続かなければならない。
    ChkBool44 "G8b_先頭でなくても検出できる", _
        modConvBridge.HasBridgeHeader("新しい回答;;;" & hdr), True
    ChkBool44 "G8c_橋渡しが無ければ検出しない", _
        modConvBridge.HasBridgeHeader("ふつうの会話です。直前の話の続きで。"), False
    ChkBool44 "G8d_空文字は検出しない", modConvBridge.HasBridgeHeader(""), False
    ' 接頭辞だけの偶然一致で誤検知しない(前後2つの印を両方見る契約)
    ChkBool44 "G8e_接頭辞だけでは検出しない", _
        modConvBridge.HasBridgeHeader("【直前の資料を見てください】"), False
    ' 文言の要件: 撤回させない(会話の情報は根拠として使ってよい)ことを明示
    ChkBool44 "G8f_会話の情報は使ってよいと書いてある", _
        (InStr(modGenPipe.CarryRules(), "根拠として使うこと自体は問題ない") > 0), True
    ChkBool44 "G8g_原文の確認を促している", _
        (InStr(modGenPipe.CarryRules(), "原文は出典でお確かめください") > 0), True
End Sub

' ---- G5: 結論の12pt太字(LeadParaIndex)をガードが奪わない -------------------
Private Sub TestLeadUnaffected44()
    Dim body As String: body = "結論です。" & vbLf & "■ 詳細" & vbLf & "・あれこれ"
    Dim withGuard As String: withGuard = modMode.DisplayNotes(body, 0.8, True, 0.3, True)
    ChkBool44 "G5a_ガードを付けても結論は第1段落のまま", _
        (modLiveStyle.LeadParaIndex(withGuard) = 1), True
    ' ⚠が先頭に付いた回は従来どおり結論強調をしない(R43 M1 の割り切りを維持)
    Dim withWarn As String: withWarn = modMode.DisplayNotes(body, 0.1, True, 0.3, True)
    ChkBool44 "G5b_低関連度の印が先頭の回は結論強調しない", _
        (modLiveStyle.LeadParaIndex(withWarn) = 0), True
End Sub

Private Sub ChkStr46(ByVal label As String, ByVal got As String, ByVal want As String)
    modTestRunner.Check "R46-48-" & label, (StrComp(got, want, vbBinaryCompare) = 0), _
        "実際=[" & got & "] 期待=[" & want & "]"
End Sub

' ---- G9: 吹き出し本文の整形(R46 B-2・modPeek.PeekBodyText) ------------------
' 実機報告「Excelが出典元の場合、なんかずれて表示されて、変な余白がある」。
' Excel 由来の行は modChunker が 100% 表行(タブあり)に落とすので、本文行の
' CollapseSpaces を迂回してタブが生で届く。散文(PDF/Word)との扱いを分ける。
Private Sub TestPeekBody46()
    Dim tb As String
    tb = "# " & ChrW(&H30B7) & ChrW(&H30FC) & ChrW(&H30C8) & ": " & "売上" & vbLf & _
         "[A1] 日付" & vbTab & "[B1] 得意先" & vbLf & _
         "[A2] 4/1" & vbTab & vbTab & "[D2] 120"
    Dim got As String: got = modPeek.PeekBodyText(tb)
    ' (1) 先頭の "# " は落ちる(breadcrumb に同じシート名が既に出ているため)
    ChkBool46 "G9a_先頭の#が落ちる", (Left$(got, 1) = "#"), False
    ' (2) タブは固定幅の区切りへ。生のタブは1つも残らない
    ChkBool46 "G9b_タブが残らない", (InStr(got, vbTab) > 0), False
    ChkBool46 "G9c_区切りが入る", (InStr(got, " " & ChrW(&H2502) & " ") > 0), True
    ' (3) 空セルの連続タブは区切り1つへ畳む(区切りが2つ続かない)
    ChkBool46 "G9d_連続タブを畳む", _
        (InStr(got, ChrW(&H2502) & "  " & ChrW(&H2502)) > 0), False
    ' (4) 表由来では改行を二重化しない(1行=1シート行で空行だらけになるため)
    ChkBool46 "G9e_表では空行を作らない", (InStr(got, vbLf & vbLf) > 0), False

    ' 散文(タブなし)は従来どおり二重化して読みやすくする
    Dim prose As String: prose = "第1条 目的" & vbLf & "この約款は…"
    Dim got2 As String: got2 = modPeek.PeekBodyText(prose)
    ChkBool46 "G9f_散文は従来どおり二重化する", (InStr(got2, vbLf & vbLf) > 0), True
    ChkStr46 "G9g_散文の本文は変えない", Replace(got2, vbLf & vbLf, vbLf), prose
End Sub

' ---- G10: 前の回答をコピーの1行見出し(R46・modClip.HeadLineOf) --------------
Private Sub TestHeadLine46()
    ChkStr46 "G10a_先頭行だけを取る", _
        modClip.HeadLineOf("結論です。" & vbLf & "つづき"), "結論です。"
    ChkStr46 "G10b_vbCr区切りでも先頭行", _
        modClip.HeadLineOf("結論です。" & vbCr & "つづき"), "結論です。"
    ChkStr46 "G10c_空なら印を出す", modClip.HeadLineOf(""), "(空の回答)"
    ChkStr46 "G10d_空白だけでも印を出す", modClip.HeadLineOf("   " & vbLf & "x"), "(空の回答)"
    ' 28字を超えたら … を付けて切る(切れたことが読み手に分かる形)
    Dim long1 As String: long1 = String$(40, "あ")
    Dim got As String: got = modClip.HeadLineOf(long1)
    ChkBool46 "G10e_長い行は切る", (Len(got) = 29), True
    ChkBool46 "G10f_切ったら…を付ける", (Right$(got, 1) = ChrW(&H2026)), True
    ' 28字ちょうどは切らない(境界)
    ChkStr46 "G10g_28字ちょうどは切らない", _
        modClip.HeadLineOf(String$(28, "い")), String$(28, "い")
End Sub

Private Sub ChkBool46(ByVal label As String, ByVal got As Boolean, ByVal want As Boolean)
    modTestRunner.Check "R46-48-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

' ---- G11: 確信度(R46 A-4・modMode.ConfidenceLevel/CoverageOf/FlatnessOf) ----
' しきい値は build_config_rows の既定と同値(conf_cov_green=1.0 /
' conf_flat_green=1.6 / conf_cov_amber=0.9)。実測の根拠は modMode の注記。
Private Sub TestConfidence46()
    Const CG As Double = 1#, FG As Double = 1.6, CA As Double = 0.9
    ' 答えがある型: 語が全部当たり、1位が突出している
    ChkBool46 "G11a_全語一致かつ突出なら緑", _
        (modMode.ConfidenceLevel(1#, 1.8, CG, FG, CA) = 2), True
    ' 「三井住友の株価」型: 述語が本棚に無い(cov 0.33)・平坦(flat 1.05)
    ChkBool46 "G11b_語が欠けて平坦なら赤", _
        (modMode.ConfidenceLevel(0.33, 1.05, CG, FG, CA) = 0), True
    ' 語は全部当たるが平坦: 特定の資料が当たっていない → 緑にしない
    ChkBool46 "G11c_全語一致でも平坦なら緑にしない", _
        (modMode.ConfidenceLevel(1#, 1.2, CG, FG, CA) = 1), True
    ' 突出しているが語が欠ける → 緑にしない(旧実装はここを緑にしていた)
    ChkBool46 "G11d_突出でも語が欠ければ緑にしない", _
        (modMode.ConfidenceLevel(0.5, 2.5, CG, FG, CA) = 0), True
    ' 黄の境界(cov=0.9 ちょうどは黄、その下は赤)
    ChkBool46 "G11e_黄の境界ちょうどは黄", _
        (modMode.ConfidenceLevel(0.9, 1#, CG, FG, CA) = 1), True
    ChkBool46 "G11f_黄の境界未満は赤", _
        (modMode.ConfidenceLevel(0.89, 1#, CG, FG, CA) = 0), True

    ' CoverageOf: 語が実在する割合。modSparse が語を抜く。
    Dim body As String
    body = "人身傷害保険の搭乗中のみ担保特約について、補償範囲は…"
    ChkBool46 "G11g_本文にある語は数える", _
        (modMode.CoverageOf("人身傷害保険の補償範囲は", body) >= 0.99), True
    ChkBool46 "G11h_本文に無い語は数えない", _
        (modMode.CoverageOf("三井住友の株価は", body) < 0.5), True

    ' FlatnessOf: 1位÷平均。全部同じなら1.0、1位だけ高ければ大きくなる。
    Dim flatSc() As Double: ReDim flatSc(1 To 3)
    flatSc(1) = 0.4: flatSc(2) = 0.4: flatSc(3) = 0.4
    ChkBool46 "G11i_横並びは1.0", (Abs(modMode.FlatnessOf(flatSc, 3) - 1#) < 0.001), True
    Dim peakSc() As Double: ReDim peakSc(1 To 3)
    peakSc(1) = 0.9: peakSc(2) = 0.3: peakSc(3) = 0.3
    ChkBool46 "G11j_1位が突出すると大きくなる", _
        (modMode.FlatnessOf(peakSc, 3) > 1.7), True
    Dim zeroSc() As Double: ReDim zeroSc(1 To 2)
    zeroSc(1) = 0#: zeroSc(2) = 0#
    ChkBool46 "G11k_全部0なら平坦扱い", (modMode.FlatnessOf(zeroSc, 2) = 1#), True

    ' 出典帯の見出しがバッジに連動する(🔴のとき「出典」と呼ばない)
    ChkBool46 "G11l_赤では出典と呼ばない", _
        (InStr(modPeek.CiteLabelText(0), "出典") > 0), False
    ChkBool46 "G11m_赤では根拠でないと書く", _
        (InStr(modPeek.CiteLabelText(0), "根拠ではありません") > 0), True
    ChkBool46 "G11n_黄以上は出典と呼ぶ", _
        (InStr(modPeek.CiteLabelText(1), "出典") > 0), True
End Sub

' ---- G12: 回答本文の数字の実在チェック(R46・modGround) --------------------
' 金額・第N条・％のような「間違うと実害が出る値」が、渡した資料と質問文に
' 実在するかを照合する。誤検知(実在するのに指摘する)を出すと警告が日常化して
' 誰も読まなくなるので、拾わない側の線引きも固定する。
Private Sub TestGround46()
    Dim src As String
    src = "第12条 保険金の限度額は 1,000,000円 とする。免責は 30日。" & vbLf & _
          "支払割合は 80% とし、対象は 3件 までとする。"

    ' 実在する数字は指摘しない
    ChkStr46 "G12a_実在する数字は指摘しない", _
        modGround.UngroundedNumbers("限度額は1,000,000円、免責は30日です。", src), ""
    ChkStr46 "G12b_条番号も実在すれば指摘しない", _
        modGround.UngroundedNumbers("第12条をご確認ください。", src), ""
    ' 実在しない金額は捕まえる
    ChkStr46 "G12c_実在しない金額を捕まえる", _
        modGround.UngroundedNumbers("限度額は 7,654,321円 です。", src), "7,654,321円"
    ChkStr46 "G12d_実在しない条番号を捕まえる", _
        modGround.UngroundedNumbers("第99条が適用されます。", src), "第99条"

    ' 質問文に出ていれば引き写しなので指摘しない
    ChkStr46 "G12e_質問に出た数字は指摘しない", _
        modGround.UngroundedNumbers("ご質問の 500万円 について…", src & vbLf & "500万円の車両保険"), ""

    ' 自分で付けたタグの中の数字は数えない
    ChkStr46 "G12f_出典タグの数字は数えない", _
        modGround.UngroundedNumbers("根拠です [本棚:約款.pdf p.123]", src), ""
    ChkStr46 "G12g_セル番地は数えない", _
        modGround.UngroundedNumbers("こちらです [B4567] の行", src), ""
    ChkStr46 "G12h_付近表記も数えない", _
        modGround.UngroundedNumbers("値は (C8912 付近) にあります", src), ""

    ' 行頭の箇条書き番号は主張ではない
    ChkStr46 "G12i_箇条書き番号は数えない", _
        modGround.UngroundedNumbers("まとめ" & vbLf & "1. あれ" & vbLf & "2. これ", src), ""
    ' 1〜2桁の裸の数字は拾わない(数え上げで頻出するため)
    ChkStr46 "G12j_短い裸の数字は拾わない", _
        modGround.UngroundedNumbers("方法は 2 つあります。", src), ""

    ' 上限で切って「ほかN件」を出す
    Dim many As String
    many = "111111円 222222円 333333円 444444円 555555円 666666円"
    Dim got As String: got = modGround.UngroundedNumbers(many, src, 4)
    ChkBool46 "G12k_上限4件で切る", (UBound(Split(got, "|")) + 1 = 4), True
    ChkBool46 "G12l_ほかN件を出す", _
        (InStr(modGround.GroundNoteText(got, 6), "ほか2件") > 0), True
    ChkStr46 "G12m_指摘が無ければ注記も出さない", modGround.GroundNoteText("", 0), ""
    ChkBool46 "G12n_注記は断定しない", _
        (InStr(modGround.GroundNoteText("999999円", 1), "可能性があります") > 0), True
End Sub

' ----------------------------------------------------------------------------
' RunAll48 - modTestRunner から呼ばれる総合エントリーポイント
' ----------------------------------------------------------------------------
Public Sub RunAll48()
    On Error GoTo H01Fail48
    TestCellOnlyAddress48
H02Next48:
    On Error GoTo H02Fail48
    TestBColStart48
H03Next48:
    On Error GoTo H03Fail48
    TestColAdvance48
H04Next48:
    On Error GoTo H04Fail48
    TestRowNumberAdvance48
H05Next48:
    On Error GoTo H05Fail48
    TestBeyondZ48
H06Next48:
    On Error GoTo H06Fail48
    TestEmptyRow48
H07Next48:
    On Error GoTo H07Fail48
    TestBackwardCompat48
H08Next48:
    On Error GoTo H08Fail48
    TestBoostBelowThresholds44
H09Next48:
    On Error GoTo H09Fail48
    TestGuardSuffix44
H10Next48:
    On Error GoTo H10Fail48
    TestWarnAndGuard44
H11Next48:
    On Error GoTo H11Fail48
    TestNoteMarks44
H12Next48:
    On Error GoTo H12Fail48
    TestLeadUnaffected44
H13Next48:
    On Error GoTo H13Fail48
    TestOnePassHasGuard44
H14Next48:
    On Error GoTo H14Fail48
    TestSheetHeading45
H15Next48:
    On Error GoTo H15Fail48
    TestCarryGuard45
H16Next48:
    On Error GoTo H16Fail48
    TestPeekBody46
H17Next48:
    On Error GoTo H17Fail48
    TestHeadLine46
H18Next48:
    On Error GoTo H18Fail48
    TestConfidence46
H19Next48:
    On Error GoTo H19Fail48
    TestGround46
H01Done48:
    On Error GoTo 0
    Exit Sub

H01Fail48:
    modTestRunner.Check "TestCellOnlyAddress48(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H02Next48
H02Fail48:
    modTestRunner.Check "TestBColStart48(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H03Next48
H03Fail48:
    modTestRunner.Check "TestColAdvance48(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H04Next48
H04Fail48:
    modTestRunner.Check "TestRowNumberAdvance48(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H05Next48
H05Fail48:
    modTestRunner.Check "TestBeyondZ48(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H06Next48
H06Fail48:
    modTestRunner.Check "TestEmptyRow48(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H07Next48
H07Fail48:
    modTestRunner.Check "TestBackwardCompat48(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H08Next48
H08Fail48:
    modTestRunner.Check "TestBoostBelowThresholds44(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H09Next48
H09Fail48:
    modTestRunner.Check "TestGuardSuffix44(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H10Next48
H10Fail48:
    modTestRunner.Check "TestWarnAndGuard44(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H11Next48
H11Fail48:
    modTestRunner.Check "TestNoteMarks44(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H12Next48
H12Fail48:
    modTestRunner.Check "TestLeadUnaffected44(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H13Next48
H13Fail48:
    modTestRunner.Check "TestOnePassHasGuard44(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H14Next48
H14Fail48:
    modTestRunner.Check "TestSheetHeading45(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H15Next48
H15Fail48:
    modTestRunner.Check "TestCarryGuard45(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H16Next48
H16Fail48:
    modTestRunner.Check "TestPeekBody46(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H17Next48
H17Fail48:
    modTestRunner.Check "TestHeadLine46(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H18Next48
H18Fail48:
    modTestRunner.Check "TestConfidence46(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H19Next48
H19Fail48:
    modTestRunner.Check "TestGround46(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done48
End Sub
