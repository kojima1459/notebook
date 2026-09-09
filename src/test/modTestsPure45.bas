Attribute VB_Name = "modTestsPure45"
Option Explicit

' ============================================================================
' modTestsPure45 - R42 §1(A1/A2・受入FAIL PDF-F01/RT-27/28/37/スモーク5・
'   PDF-F02/RT-22/23)の modChunkPage 純ロジック回帰。
'   modTestRunner.RunAllPureTests から直接呼ばれる独立の別枝
'   (modTestsPure31〜44と同型)。
' ----------------------------------------------------------------------------
' 【何を固定するか(手計算の根拠。仕様 spec_20260909_R42 §1-5 対応)】
'
'   A LineStarts: Join(buf(1..n), vbLf) の各行開始位置。
'     3行("abc","de","f") → 行1開始=1、行2開始=1+Len("abc")+1=5、
'     行3開始=5+Len("de")+1=8 → "1,5,8"。1行("x") → "1"。
'
'   B PageAtPos: starts=(1,5,8) pg=(1,2,5)(=行1:p1/行2:p2/行3:p5相当)。
'     pos1→starts(1)<=1のみ→pg(1)=1。pos4→starts(2)=5>4なのでpg(1)=1のまま。
'     pos5→starts(2)=5<=5→pg(2)=2。pos7→pg(2)=2のまま(starts(3)=8>7)。
'     pos8→starts(3)=8<=8→pg(3)=5。pos100→同じくpg(3)=5。
'     pos0→starts(1)=1>0(=pos<starts(1))→pg(1)=1。n=0→0。
'
'   C FirstInkPos: 空白類(半角/全角スペース・vbLf・vbCr・vbTab)でない最初の
'     位置。"  x"(startPos=1)→位置3の"x"。vbLf&"x"(startPos=1)→位置2の"x"。
'     "x"(startPos=1)→位置1。"abc   "をstartPos=4から探すと全部空白→
'     見つからずstartPos=4を返す。全角空白(U+3000)&"x"→位置2。
'
'   D PlanWindows: modChunker.FlushBlock:582-600 のスライディング規則を1文字
'     も変えず移植(FindSentenceBoundaryを呼ぶ)。
'     D1(境界なし): "あ"×30字・tgt=10・ov=3・room=100。
'       startPos=1: windowEnd=10(句読点なし・total30未満なのでFindSentence
'       Boundary(body,10,10\4=2)を引くが0=境界なし)→"1-10"。
'       nextStart=10-3+1=8。startPos=8: windowEnd=17(同様に境界なし)→
'       "8-17"。nextStart=17-3+1=15。startPos=15: windowEnd=24→"15-24"。
'       nextStart=24-3+1=22。startPos=22: windowEnd=22+9=31>30→30に
'       クランプ(windowEnd=total なので境界探索なし)→"22-30"。
'       windowEnd>=total(30)で終了。→ "1-10|8-17|15-24|22-30"。
'     D2(句点で境界スナップ): "あ"×10字+"。"+"あ"×9字(計20字。句点は
'       位置11)。startPos=1: 素のwindowEnd=10・FindSentenceBoundary
'       (body,10,slack=10\4=2)は位置10..12を走査し位置11の"。"を検出→
'       windowEnd=11→"1-11"。nextStart=11-3+1=9。startPos=9:
'       windowEnd=9+9=18・slack探索(18..20)に句読点なし→"9-18"。
'       nextStart=18-3+1=16。startPos=16: windowEnd=16+9=25>20→20に
'       クランプ(windowEnd=totalなので境界探索なし)→"16-20"。終了。
'       → "1-11|9-18|16-20"。
'
'   E PhysicalKeep(firstIdx,keptN,maxPages): 上限は物理ページ番号に当てる。
'     (0,5,2)→firstIdx+keptN=5>2→truncated=True・戻り=2-0=2。
'     (1,2,2)→3>2→truncated=True・戻り=2-1=1(先頭白紙1枚+本文2枚・上限2
'     ならp2のみ残る)。(2,1,2)→3>2→truncated=True・戻り=2-2=0(本文なし)。
'     (0,2,0)→maxPages<=0→truncated=False・戻り=keptNそのまま=2。
'     (1,1,3)→2<=3→truncated=False・戻り=keptNそのまま=1。
'
'   F StartPagesOf(受入報告の再現ゴールデン): 見出しの無い5ページ・各ページ
'     2行(1行5字の"あ")をBlockAddPgで組む(pg=1,1,2,2,3,3,4,4,5,5)。
'     Join後の行開始位置(1行=5字+区切りvbLf1字=6字間隔): 1,7,13,19,25,31,
'     37,43,49,55(全10行・総字数59=5*10+9)。tgt=15・ov=4・room=1000で
'     PlanWindowsすると(句読点なしなので境界スナップ無し):
'       "1-15|12-26|23-37|34-48|45-59" の5窓。
'     各窓の開始ページ(FirstInkPos+PageAtPos):
'       s=1  → 位置1は行1の先頭(あ) → page=1。
'       s=12 → 位置12は行2直後の区切り(vbLf・空白類)→FirstInkPosが
'              位置13(行3の先頭)まで進む → page=2。
'       s=23 → 位置23は行4の最終文字(あ)そのもの(空白ではない)→page=2。
'              ※オーバーラップが改ページをまたぐ例: この窓(23-37)の本体は
'              大半が3ページ目(位置25-35)〜4ページ目(位置37)だが、開始
'              位置23が2ページ目の末尾行(行4)の中にあるため「前ページ」
'              (page=2)に付く(仕様§1-3の規則どおり)。
'       s=34 → 位置34は行6の中(あ)→page=3(この窓34-48も同様に大半が
'              4ページ目だが開始位置は3ページ目に属するため前ページ)。
'       s=45 → 位置45は行8の中(あ)→page=4(同型。大半は5ページ目)。
'     → "1,2,2,3,4"。
'
'   G FlushPaged(UDT直接検査・CanUseTypeArrays()ガード下・Excel実機のみ):
'     Fの同じ10行・5ページ構成にtgt=15・ov=4・mx=20(chapter/section空)を
'     渡す。crumb=BuildBreadcrumb("〔資料〕","","")&vbLf="【〔資料〕】"&vbLf
'     (7字)。atomicLimit=mx-7=13<body59字なのでPlanWindows経路になり、
'     Fと同じ5窓・同じページ列(1,2,2,3,4)のShelfChunkが積まれる
'     (room=32000-7=31993なので切詰めは発生しない)。呼び出し後n=0。
'
' 【なぜ FlushPaged の直接検査だけ CanUseTypeArrays() ガード下か】
'   ShelfChunk配列(Public Type)はLibreOffice実行環境ではReDimが実行時
'   エラー420になる既知の制限(modChunker冒頭コメント・modTestsPure3等と
'   同型)。StartPagesOf(F)は文字列とLong配列だけを扱う検証口なのでLOでも
'   フルに実行できる。LOでは本群だけ[SKIP]1件になる
'   (tools/run_lo_tests.py の EXPECTED_SKIP_MAX を15へ)。
' ============================================================================

Private Sub ChkStr45(ByVal label As String, ByVal got As String, ByVal want As String)
    modTestRunner.Check "R42-45-" & label, (StrComp(got, want, vbBinaryCompare) = 0), _
        "実際=[" & got & "] 期待=[" & want & "]"
End Sub

Private Sub ChkLong45(ByVal label As String, ByVal got As Long, ByVal want As Long)
    modTestRunner.Check "R42-45-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkBool45(ByVal label As String, ByVal got As Boolean, ByVal want As Boolean)
    modTestRunner.Check "R42-45-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

' CanUseTypeArrays - modTypes(別モジュール)のPublic Typeの配列をReDimできる
'   実行環境かどうかを実測で判定する(modTestsPure/2/3/4/6/7/8/23と同じ複製。
'   Privateは別モジュールから呼べないためモジュールごとに複製する契約)。
Private Function CanUseTypeArrays() As Boolean
    On Error Resume Next
    Err.Clear
    Dim probe() As ShelfChunk
    ReDim probe(0 To 0)
    CanUseTypeArrays = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' ---- A: LineStarts --------------------------------------------------------
Private Sub TestLineStarts45()
    Dim buf3() As String: ReDim buf3(1 To 3)
    buf3(1) = "abc": buf3(2) = "de": buf3(3) = "f"
    ChkStr45 "LineStarts_3行", modChunkPage.LineStarts(buf3, 3), "1,5,8"

    Dim buf1() As String: ReDim buf1(1 To 1)
    buf1(1) = "x"
    ChkStr45 "LineStarts_1行", modChunkPage.LineStarts(buf1, 1), "1"
End Sub

' ---- B: PageAtPos ----------------------------------------------------------
Private Sub TestPageAtPos45()
    Dim starts() As Long: ReDim starts(1 To 3)
    starts(1) = 1: starts(2) = 5: starts(3) = 8
    Dim pg() As Long: ReDim pg(1 To 3)
    pg(1) = 1: pg(2) = 2: pg(3) = 5

    ChkLong45 "PageAtPos_pos1", modChunkPage.PageAtPos(starts, pg, 3, 1), 1
    ChkLong45 "PageAtPos_pos4", modChunkPage.PageAtPos(starts, pg, 3, 4), 1
    ChkLong45 "PageAtPos_pos5", modChunkPage.PageAtPos(starts, pg, 3, 5), 2
    ChkLong45 "PageAtPos_pos7", modChunkPage.PageAtPos(starts, pg, 3, 7), 2
    ChkLong45 "PageAtPos_pos8", modChunkPage.PageAtPos(starts, pg, 3, 8), 5
    ChkLong45 "PageAtPos_pos100", modChunkPage.PageAtPos(starts, pg, 3, 100), 5
    ChkLong45 "PageAtPos_pos0", modChunkPage.PageAtPos(starts, pg, 3, 0), 1

    Dim emptyStarts() As Long, emptyPg() As Long
    ChkLong45 "PageAtPos_n0", modChunkPage.PageAtPos(emptyStarts, emptyPg, 0, 5), 0
End Sub

' ---- C: FirstInkPos ---------------------------------------------------------
Private Sub TestFirstInkPos45()
    ChkLong45 "FirstInkPos_先頭2空白", modChunkPage.FirstInkPos("  x", 1), 3
    ChkLong45 "FirstInkPos_改行", modChunkPage.FirstInkPos(vbLf & "x", 1), 2
    ChkLong45 "FirstInkPos_先頭から本文", modChunkPage.FirstInkPos("x", 1), 1
    ChkLong45 "FirstInkPos_末尾空白のみ", modChunkPage.FirstInkPos("abc   ", 4), 4
    ChkLong45 "FirstInkPos_全角空白", modChunkPage.FirstInkPos(ChrW$(&H3000) & "x", 1), 2
End Sub

' ---- D: PlanWindows ----------------------------------------------------------
Private Sub TestPlanWindows45()
    Dim bodyA As String: bodyA = String(30, "あ")
    ChkStr45 "PlanWindows_境界なし", _
        modChunkPage.PlanWindows(Len(bodyA), 10, 3, 100, bodyA), "1-10|8-17|15-24|22-30"

    Dim bodyB As String: bodyB = String(10, "あ") & "。" & String(9, "あ")
    modTestRunner.Check "R42-45-PlanWindows_前提_bodyB20字", (Len(bodyB) = 20), "len=" & Len(bodyB)
    ChkStr45 "PlanWindows_句点で境界スナップ", _
        modChunkPage.PlanWindows(Len(bodyB), 10, 3, 100, bodyB), "1-11|9-18|16-20"
End Sub

' ---- E: PhysicalKeep ----------------------------------------------------------
Private Sub TestPhysicalKeep45()
    Dim tr As Boolean

    tr = False
    ChkLong45 "PhysicalKeep_A_keptN", modChunkPage.PhysicalKeep(0, 5, 2, tr), 2
    ChkBool45 "PhysicalKeep_A_truncated", tr, True

    tr = False
    ChkLong45 "PhysicalKeep_B_keptN", modChunkPage.PhysicalKeep(1, 2, 2, tr), 1
    ChkBool45 "PhysicalKeep_B_truncated", tr, True

    tr = False
    ChkLong45 "PhysicalKeep_C_keptN", modChunkPage.PhysicalKeep(2, 1, 2, tr), 0
    ChkBool45 "PhysicalKeep_C_truncated", tr, True

    tr = True
    ChkLong45 "PhysicalKeep_D_keptN", modChunkPage.PhysicalKeep(0, 2, 0, tr), 2
    ChkBool45 "PhysicalKeep_D_truncated", tr, False

    tr = True
    ChkLong45 "PhysicalKeep_E_keptN", modChunkPage.PhysicalKeep(1, 1, 3, tr), 1
    ChkBool45 "PhysicalKeep_E_truncated", tr, False
End Sub

' 5ページ・見出し無し・各ページ2行(1行5字の"あ")をBlockAddPgで組む共通部品
' (F/Gの両方で使う。呼ぶたびに新しい配列を作るのでFとGは互いに干渉しない)。
Private Sub BuildFivePageBlock45(ByRef buf() As String, ByRef pg() As Long, ByRef n As Long)
    n = 0
    Dim li As Long
    For li = 1 To 10
        Dim pageNo As Long: pageNo = (li - 1) \ 2 + 1   ' 1,1,2,2,3,3,4,4,5,5
        modChunkPage.BlockAddPg buf, pg, n, String(5, "あ"), pageNo
    Next li
End Sub

' ---- F: StartPagesOf(受入報告の再現ゴールデン) --------------------------------
Private Sub TestStartPagesOfGolden45()
    Dim buf() As String, pg() As Long, n As Long
    BuildFivePageBlock45 buf, pg, n
    modTestRunner.Check "R42-45-StartPagesOf_前提_10行", (n = 10), "n=" & n

    ChkStr45 "StartPagesOf_5ページ_見出し無し_開始ページ列", _
        modChunkPage.StartPagesOf(buf, pg, n, 15, 4, 1000), "1,2,2,3,4"
End Sub

' ---- G: FlushPaged(UDT直接検査・Excel実機のみ) ------------------------------
Private Sub TestFlushPagedDirect45()
    If Not CanUseTypeArrays() Then
        modTestRunner.Check "[SKIP] FlushPaged(UDT)直接検査: LO環境の既知の制限によりスキップ", True, _
            "ShelfChunk(Public Type)配列のReDimはLibreOffice実行環境で実行時エラー420に" & _
            "なることを確認済み(modChunker冒頭コメント参照)。ページ列(1,2,2,3,4)自体は" & _
            "本モジュールのTestStartPagesOfGolden45(F)がLO上で固定しており、ここで確認する" & _
            "のはFlushPagedがそのページ列でShelfChunk配列を正しく積むことだけなので、" & _
            "Excel実機受入チェックで必ず再確認すること。"
        Exit Sub
    End If

    Dim buf() As String, pg() As Long, n As Long
    BuildFivePageBlock45 buf, pg, n

    Dim outArr() As ShelfChunk: ReDim outArr(0 To 15)
    Dim outCount As Long: outCount = 0
    modChunkPage.FlushPaged buf, pg, n, "", "", 15, 4, 20, outArr, outCount

    ChkLong45 "FlushPaged_outCount", outCount, 5
    ChkLong45 "FlushPaged_n0リセット", n, 0
    If outCount = 5 Then
        ChkLong45 "FlushPaged_page1", outArr(0).page, 1
        ChkLong45 "FlushPaged_page2", outArr(1).page, 2
        ChkLong45 "FlushPaged_page3", outArr(2).page, 2
        ChkLong45 "FlushPaged_page4", outArr(3).page, 3
        ChkLong45 "FlushPaged_page5", outArr(4).page, 4

        Dim crumb As String
        crumb = modChunker.BuildBreadcrumb(modChunker.CRUMB_PLACEHOLDER, "", "") & vbLf
        modTestRunner.Check "R42-45-FlushPaged_crumb前置", _
            (Left$(outArr(0).full_text, Len(crumb)) = crumb), "full_text=" & outArr(0).full_text
        ChkLong45 "FlushPaged_1件目本文長15字", Len(outArr(0).full_text) - Len(crumb), 15
    End If
End Sub

' ----------------------------------------------------------------------------
' RunAll45 - modTestRunner から呼ばれる総合エントリーポイント
' ----------------------------------------------------------------------------
Public Sub RunAll45()
    On Error GoTo H01Fail45
    TestLineStarts45
H02Next45:
    On Error GoTo H02Fail45
    TestPageAtPos45
H03Next45:
    On Error GoTo H03Fail45
    TestFirstInkPos45
H04Next45:
    On Error GoTo H04Fail45
    TestPlanWindows45
H05Next45:
    On Error GoTo H05Fail45
    TestPhysicalKeep45
H06Next45:
    On Error GoTo H06Fail45
    TestStartPagesOfGolden45
H07Next45:
    On Error GoTo H07Fail45
    TestFlushPagedDirect45
H01Done45:
    On Error GoTo 0
    Exit Sub

H01Fail45:
    modTestRunner.Check "TestLineStarts45(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H02Next45
H02Fail45:
    modTestRunner.Check "TestPageAtPos45(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H03Next45
H03Fail45:
    modTestRunner.Check "TestFirstInkPos45(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H04Next45
H04Fail45:
    modTestRunner.Check "TestPlanWindows45(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H05Next45
H05Fail45:
    modTestRunner.Check "TestPhysicalKeep45(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H06Next45
H06Fail45:
    modTestRunner.Check "TestStartPagesOfGolden45(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H07Next45
H07Fail45:
    modTestRunner.Check "TestFlushPagedDirect45(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done45
End Sub
