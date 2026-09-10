Attribute VB_Name = "modTestsPure48"
Option Explicit

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
    Resume H01Done48
End Sub
