Attribute VB_Name = "modTestsPure39"
Option Explicit

' ============================================================================
' modTestsPure39 - R36波1(§3スクショ本文表示・§4画像📁)の純ロジック回帰。
'   modTestsPure31〜38 と同じく既存チェーン(modTestsPure.RunAll→…)へは繋がず、
'   modTestRunner.RunAllPureTests から直接呼ばれる RunAll39 の1本が入口
'   (同型の別枝)。
' ----------------------------------------------------------------------------
' 【何を固定するか】
'   A modTextView.StripBreadcrumb: 取込時に焼かれた先頭行【資料名>章>条】の
'     剥がし方。modVaultGallery.MakePreviewText(既存踏襲)と同じく、
'     (1)先頭が【でvbLfがあれば最初のvbLfまでを落とす
'     (2)先頭が【でないなら無変更
'     (3)空文字は空文字のまま
'     (4)「【だけ」(改行が無い不完全な行)は無変更(誤って全部消さないため)
'     の4パターンを固定する。
'   B modTextView.SortPagesStable: pages()/texts()並列配列のpage昇順・安定
'     ソート(挿入ソート)。
'     (1)ばらばらの順序→昇順に並ぶ
'     (2)page値が同じ要素が複数あるとき、元の相対順序を保つ(安定性)
'     (3)n=0/n=1は無変更(早期return)
'     (4)LBoundが1始まりの配列でも正しく動く(UDTを跨がない配列2本渡しの
'        契約どおり)
'     を固定する。
'   C modShelfVision.ImageDialogPattern: 📁追加のFileDialogフィルタへ足す
'     画像パターン文字列("*.png;*.jpg;*.jpeg")の内容一致を固定する。
' ============================================================================

' ---- A: StripBreadcrumb ------------------------------------------------------
'   discriminate:
'   ・vbLfの位置を1つずらす(Mid$の第2引数をlfPos+2等にする)と、(1)の
'     先頭1文字が本文側に残る/消えすぎるかたちで落ちる。
'   ・「先頭が【なら常に剥がす」実装(lfPos>0のガードを外す)にすると、
'     (4)「【だけ」が空文字になり落ちる。
'   ・「先頭が【」の判定を削ると、(2)は通るが本来剥がすべき(1)が
'     無変更のまま返り落ちる。
Private Sub TestStripBreadcrumb39()
    ' (1) breadcrumb行1本 + 本文複数行 → breadcrumb行だけ落ちる。
    ChkStr39 "A_breadcrumb行を剥がす", _
        modTextView.StripBreadcrumb("【資料名>章>条】" & vbLf & "本文1行目" & vbLf & "本文2行目"), _
        "本文1行目" & vbLf & "本文2行目"

    ' (2) 先頭が【でない → 無変更。
    ChkStr39 "A_先頭が【でなければ無変更", _
        modTextView.StripBreadcrumb("本文だけ(先頭に【なし)"), _
        "本文だけ(先頭に【なし)"

    ' (3) 空文字 → 空文字のまま。
    ChkStr39 "A_空文字は空文字のまま", modTextView.StripBreadcrumb(""), ""

    ' (4) 「【だけ」(改行なし) → vbLfが無いので無変更(既存MakePreviewText踏襲)。
    ChkStr39 "A_【だけ(改行なし)は無変更", _
        modTextView.StripBreadcrumb("【だけ"), "【だけ"
End Sub

' ---- B: SortPagesStable ------------------------------------------------------
'   discriminate:
'   ・比較を `<` にすると(等しい要素も動かす)、(2)の安定性が崩れて
'     元の相対順序が入れ替わり落ちる。
'   ・挿入位置を1つずらす(j+1をjにする等)と(1)(2)とも配列が壊れて落ちる。
'   ・n<=1の早期returnを外しても実害は薄いが、LBoundをハードコード0にすると
'     (4)の1始まり配列テストで先頭要素が欠落し落ちる。
Private Sub TestSortPagesStable39()
    ' (1) ばらばらの順序 → page昇順。
    Dim p1() As Long: p1 = Array(3, 1, 2)
    Dim t1() As String: t1 = Array("c", "a", "b")
    modTextView.SortPagesStable p1, t1, 3
    ChkLong39 "B_昇順_page0", p1(0), 1
    ChkLong39 "B_昇順_page1", p1(1), 2
    ChkLong39 "B_昇順_page2", p1(2), 3
    ChkStr39 "B_昇順_text0", t1(0), "a"
    ChkStr39 "B_昇順_text1", t1(1), "b"
    ChkStr39 "B_昇順_text2", t1(2), "c"

    ' (2) page値が同じ要素の安定性: [2(x1), 1(y), 2(x2)] → [1(y), 2(x1), 2(x2)]
    '     (page=2の2件は元の相対順序=x1が先→x2が後 を保つ)。
    Dim p2() As Long: p2 = Array(2, 1, 2)
    Dim t2() As String: t2 = Array("x1", "y", "x2")
    modTextView.SortPagesStable p2, t2, 3
    ChkLong39 "B_安定性_page0", p2(0), 1
    ChkStr39 "B_安定性_text0", t2(0), "y"
    ChkLong39 "B_安定性_page1", p2(1), 2
    ChkStr39 "B_安定性_同点1件目が先", t2(1), "x1"
    ChkLong39 "B_安定性_page2", p2(2), 2
    ChkStr39 "B_安定性_同点2件目が後", t2(2), "x2"

    ' (3) n=0/n=1は無変更(早期return)。
    Dim p3() As Long: p3 = Array(9)
    Dim t3() As String: t3 = Array("z")
    modTextView.SortPagesStable p3, t3, 1
    ChkLong39 "B_n1は無変更", p3(0), 9
    modTextView.SortPagesStable p3, t3, 0
    ChkLong39 "B_n0は無変更", p3(0), 9

    ' (4) LBoundが1始まりの配列でも正しく動く(UDTを跨がない配列2本渡しの契約)。
    Dim p4(1 To 3) As Long
    Dim t4(1 To 3) As String
    p4(1) = 5: p4(2) = 3: p4(3) = 4
    t4(1) = "e": t4(2) = "c": t4(3) = "d"
    modTextView.SortPagesStable p4, t4, 3
    ChkLong39 "B_1始まり配列_page1", p4(1), 3
    ChkLong39 "B_1始まり配列_page2", p4(2), 4
    ChkLong39 "B_1始まり配列_page3", p4(3), 5
    ChkStr39 "B_1始まり配列_text1", t4(1), "c"
    ChkStr39 "B_1始まり配列_text2", t4(2), "d"
    ChkStr39 "B_1始まり配列_text3", t4(3), "e"
End Sub

' ---- C: modShelfVision.ImageDialogPattern -----------------------------------
'   discriminate:
'   ・区切り文字を","に変えるとFileDialog.Filtersの書式(";"区切り)と
'     食い違い、この一致テストで落ちる。
'   ・拡張子を1つでも落とす/増やすと文字列比較で落ちる。
Private Sub TestImageDialogPattern39()
    ChkStr39 "C_画像フィルタ文字列", modShelfVision.ImageDialogPattern(), "*.png;*.jpg;*.jpeg"
End Sub

Private Sub ChkLong39(ByVal label As String, ByVal got As Long, ByVal want As Long)
    modTestRunner.Check "R36-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkStr39(ByVal label As String, ByVal got As String, ByVal want As String)
    modTestRunner.Check "R36-" & label, (StrComp(got, want, vbBinaryCompare) = 0), _
        "実際=[" & got & "] 期待=[" & want & "]"
End Sub

Public Sub RunAll39()
    On Error GoTo H01Fail39
    TestStripBreadcrumb39
H02Next39:
    On Error GoTo H02Fail39
    TestSortPagesStable39
H03Next39:
    On Error GoTo H03Fail39
    TestImageDialogPattern39
H01Done39:
    On Error GoTo 0
    Exit Sub

H01Fail39:
    modTestRunner.Check "TestStripBreadcrumb39(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H02Next39
H02Fail39:
    modTestRunner.Check "TestSortPagesStable39(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H03Next39
H03Fail39:
    modTestRunner.Check "TestImageDialogPattern39(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done39
End Sub
