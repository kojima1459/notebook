Attribute VB_Name = "modTestsPure44"
Option Explicit

' ============================================================================
' modTestsPure44 - R39 F001(受入指摘§0 F001)のページ別txtwrite出力を固定する
'   純関数回帰テスト。modTestRunner.RunAllPureTests から呼ばれる RunAll44 が
'   単独入口。
' ----------------------------------------------------------------------------
' 【何を固定するか】
'   A optOcrCore.PageTxtName:
'     i=1 で "\gstext_0001.txt"(4桁ゼロ埋め)、i=12 で "0012"、
'     i=10000 で "10000"(桁あふれはそのまま桁数が増える)。
'     folderPath の末尾に "\" が無い前提で連結すること。
'   B optOcrCore.JoinPageTexts:
'     n=0 は空文字、n=1 は "a"&Chr$(12)、n=3(中間が空)は
'     "a"&FF&""&FF&"c"&FF(各要素の末尾に必ず Chr$(12) が付くこと)。
'   C JoinPageTexts の結果を modUtilText.GsPageBounds に通したときの境界:
'     ["a","","c"] -> keptN=3・firstIdx=0・lastIdx=2(中間の空ページは残る)、
'     ["","b","c"] -> keptN=2・firstIdx=1(先頭の空だけ読み飛ばす)、
'     ["a","b",""] -> keptN=2・lastIdx=1(末尾の空だけ読み飛ばす)、
'     ["","",""]   -> keptN=0(全部空白類)。
'   D optOcrCore.BuildGsTextCommand("g.exe","in.pdf","C:\t\gstext.txt"):
'     "gstext_%04d.txt" を含み、旧ファイル名 "gstext.txt" は含まず、
'     "-sDEVICE=txtwrite" を含むこと。
'
' 【どう壊すと落ちるか (discriminate)】
'   ・PageTxtName でゼロ埋め桁数を3桁にすると A_i1 が "0001" ではなく
'     "001" になり落ちる。
'   ・JoinPageTexts で最後の要素にだけ Chr$(12) を付け忘れると B_n3中間空 の
'     末尾FFが1個減って落ちる。
'   ・JoinPageTexts で空要素をまるごとスキップすると(結果に "a"&FF&"c"&FF
'     しか出ない)、C の GsPageBounds が keptN=2 に化けて落ちる。
'   ・BuildGsTextCommand が outTxt をそのまま -sOutputFile= に使い続けると
'     D_旧ファイル名なし が落ちる(退行の直接検知)。
'
' 【なぜ BuildPagesFromGsText / ReadPageFilesJoined / MaxPageFileIndex を
'   呼ばないか】
'   modExtractor.BuildPagesFromGsText は ExtractedPage 型(Public Type の
'   配列)を戻すため、LibreOffice 実行テストのモジュール境界越えの制約に
'   抵触する(modTestsPure43 の TryOnePass/SourceBlockOf と同型の死角)。
'   ReadPageFilesJoined/MaxPageFileIndex/LatestPageFileSize はファイルI/Oを
'   伴う(Dir$/FileLen/ファイル読み取り)ため純関数ではなく、LOのテスト環境に
'   実ファイルを用意する意味が薄い。これらは実機受入テスト(仕様§4)で
'   確かめる。
' ============================================================================

Private Sub ChkStr44(ByVal label As String, ByVal got As String, ByVal want As String)
    modTestRunner.Check "R39-" & label, (StrComp(got, want, vbBinaryCompare) = 0), _
        "実際=[" & Replace(got, Chr$(12), "<FF>") & "] 期待=[" & Replace(want, Chr$(12), "<FF>") & "]"
End Sub

Private Sub ChkLong44(ByVal label As String, ByVal got As Long, ByVal want As Long)
    modTestRunner.Check "R39-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkBool44(ByVal label As String, ByVal got As Boolean, ByVal want As Boolean)
    modTestRunner.Check "R39-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

' ---- A: PageTxtName -----------------------------------------------------------
Private Sub TestPageTxtName44()
    ChkStr44 "A_i1", optOcrCore.PageTxtName("C:\t", 1), "C:\t\gstext_0001.txt"
    ChkStr44 "A_i12", optOcrCore.PageTxtName("C:\t", 12), "C:\t\gstext_0012.txt"
    ChkStr44 "A_i10000", optOcrCore.PageTxtName("C:\t", 10000), "C:\t\gstext_10000.txt"
End Sub

' ---- B: JoinPageTexts ----------------------------------------------------------
Private Sub TestJoinPageTexts44()
    Dim empty0() As String
    ChkStr44 "B_n0は空", optOcrCore.JoinPageTexts(empty0, 0), ""

    Dim one(1 To 1) As String: one(1) = "a"
    ChkStr44 "B_n1", optOcrCore.JoinPageTexts(one, 1), "a" & Chr$(12)

    Dim mid3(1 To 3) As String: mid3(1) = "a": mid3(2) = "": mid3(3) = "c"
    ChkStr44 "B_n3中間空", optOcrCore.JoinPageTexts(mid3, 3), _
        "a" & Chr$(12) & "" & Chr$(12) & "c" & Chr$(12)

    ' 各要素の末尾に必ずChr$(12)が付く=全体をFF区切りで割ると要素数+1になる
    ' (末尾FFの後に空文字が1つ残る)。
    Dim parts44() As String: parts44 = Split(optOcrCore.JoinPageTexts(mid3, 3), Chr$(12))
    ChkLong44 "B_末尾FF後に空要素", UBound(parts44) - LBound(parts44) + 1, 4
End Sub

' ---- C: JoinPageTexts -> modUtilText.GsPageBounds ------------------------------
Private Sub TestJoinThenBounds44()
    Dim firstIdx As Long, lastIdx As Long

    Dim p1(1 To 3) As String: p1(1) = "a": p1(2) = "": p1(3) = "c"
    ChkLong44 "C_中間空_keptN", modUtilText.GsPageBounds( _
        optOcrCore.JoinPageTexts(p1, 3), firstIdx, lastIdx), 3
    ChkLong44 "C_中間空_firstIdx", firstIdx, 0
    ChkLong44 "C_中間空_lastIdx", lastIdx, 2

    Dim p2(1 To 3) As String: p2(1) = "": p2(2) = "b": p2(3) = "c"
    ChkLong44 "C_先頭空_keptN", modUtilText.GsPageBounds( _
        optOcrCore.JoinPageTexts(p2, 3), firstIdx, lastIdx), 2
    ChkLong44 "C_先頭空_firstIdx", firstIdx, 1

    Dim p3(1 To 3) As String: p3(1) = "a": p3(2) = "b": p3(3) = ""
    ChkLong44 "C_末尾空_keptN", modUtilText.GsPageBounds( _
        optOcrCore.JoinPageTexts(p3, 3), firstIdx, lastIdx), 2
    ChkLong44 "C_末尾空_lastIdx", lastIdx, 1

    Dim p4(1 To 3) As String: p4(1) = "": p4(2) = "": p4(3) = ""
    ChkLong44 "C_全部空_keptN", modUtilText.GsPageBounds( _
        optOcrCore.JoinPageTexts(p4, 3), firstIdx, lastIdx), 0
End Sub

' ---- D: BuildGsTextCommand -----------------------------------------------------
Private Sub TestBuildGsTextCommand44()
    Dim cmd As String
    cmd = optOcrCore.BuildGsTextCommand("g.exe", "in.pdf", "C:\t\gstext.txt")

    ChkBool44 "D_新ファイル名を含む", (InStr(cmd, "gstext_%04d.txt") > 0), True
    ChkBool44 "D_旧ファイル名なし", (InStr(cmd, "gstext.txt") = 0), True
    ChkBool44 "D_txtwriteデバイス", (InStr(cmd, "-sDEVICE=txtwrite") > 0), True
End Sub

' ----------------------------------------------------------------------------
' RunAll44 - modTestRunner から呼ばれる総合エントリーポイント
' ----------------------------------------------------------------------------
Public Sub RunAll44()
    On Error GoTo H01Fail44
    TestPageTxtName44
H02Next44:
    On Error GoTo H02Fail44
    TestJoinPageTexts44
H03Next44:
    On Error GoTo H03Fail44
    TestJoinThenBounds44
H04Next44:
    On Error GoTo H04Fail44
    TestBuildGsTextCommand44
H01Done44:
    On Error GoTo 0
    Exit Sub

H01Fail44:
    modTestRunner.Check "TestPageTxtName44(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H02Next44
H02Fail44:
    modTestRunner.Check "TestJoinPageTexts44(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H03Next44
H03Fail44:
    modTestRunner.Check "TestJoinThenBounds44(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H04Next44
H04Fail44:
    modTestRunner.Check "TestBuildGsTextCommand44(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done44
End Sub
