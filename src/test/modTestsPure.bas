Attribute VB_Name = "modTestsPure"
Option Explicit

' ============================================================================
' modTestsPure - 純ロジックモジュールのユニットテスト(MASTER_SPEC §7.8/§11)
' ----------------------------------------------------------------------------
' 役割:
'   modUtil / modChunker / modPii の「純ロジック部」を、Excelなしで
'   (tools/run_lo_tests.py のモード1 = LibreOffice headless)検証する。
'   modPrompts / modShelfSync.DiffDecision / modPack.ValidatePackMeta は
'   1モジュール30,000字上限(§7.1)を超えたため modTestsPure2.bas に分割した
'   (モジュール名一覧を参照。修正の経緯はそちらの冒頭コメントに記載)。
'   入口は modTestRunner.RunAllPureTests から呼ばれる Public Sub RunAll()。
'   RunAllは末尾で modTestsPure2.RunAll2 を呼び、分割先のテストも実行する
'   (modTestRunner.bas §7.8はmodTestsPure.RunAllだけを呼ぶ契約であり、担当外
'   につき変更しないため、分割先への導線はここに置く)。
'
' 設計判断(R4準拠): Worksheets/Range/Application/ThisWorkbook/MsgBox/
'   ActiveSheet には一切触れない。改行は vbLf 基準(§12)。
'
' ■ グループ単位の失敗隔離
'   RunAll は各テストグループ(TestModUtil/TestModChunker/…)を
'   "On Error GoTo <Label> / Resume <NextLabel>" で1グループずつ囲む。
'   modTestRunner.RunAllPureTests 側は modTestsPure.RunAll 全体を1個の
'   On Error Resume Nextで包むだけなので、RunAll内で無防備に例外が起きると
'   その時点で以降のグループが一切実行されずレポートが失われる。
'   グループ単位の隔離により「1グループの想定外エラー」が他グループの
'   結果を道連れにしないようにする(1行スコープではないが、これはR5が
'   守ろうとしている「本番コードでのエラー握りつぶし」の話ではなく、
'   テストランナー自身の頑健性のための意図的な設計であり、各グループの
'   捕捉時に必ずCheckで可視化している=握りつぶしていない)。同じ方針で
'   modTestsPure2.RunAll2の呼び出しも1グループとして隔離する。
'
' ■ 重要な既知の制約(実測で確認・tools/run_lo_tests.pyは変更不可のため回避不能):
'   LibreOffice Basic(本ハーネスの "Option VBASupport 1" 注入環境)では、
'   modTypes で定義された Public Type(ExtractedPage/ShelfChunk/Hit)の
'   「配列」を ReDim した瞬間に実行時エラー420 "Invalid object reference"
'   になることを実験で確認した(スカラー変数 "Dim h As Hit" は問題なく動く。
'   ローカルで定義したTypeの配列も問題なく動く。問題になるのは
'   「別モジュール(modTypes)で定義されたPublic Typeの配列」だけ)。
'   これは tools/README.md 記載の技術メモ(#2: スカラーのクロスモジュール
'   Type参照がOption VBASupport 1で解消する話)ではカバーされていない、
'   配列特有の別の制約であることを実験で切り分け済み。
'   影響を受けるのは modChunker.ChunkPages(ExtractedPage()/ShelfChunk())と
'   modPrompts.BuildQuickPrompt/BuildDeepDraftPrompt/BuildDeepVerifyPrompt
'   (Hit())。BuildEnrichPromptはHit配列を取らないため影響を受けない。
'   tools/run_lo_tests.py はこのテストハーネス担当(1-I)の完成物であり
'   本Wave(2-X)からは変更できないため、CanUseTypeArrays()で実行時に
'   検知し、影響を受ける環境では「LO環境の既知の制限によりスキップ」の
'   単一Checkに倒す(テストコード自体は完全な形で書いており、Excel実機や
'   将来ハーネスが直った環境では即座に本来の検証が有効になる)。
'   Excel実機受入チェックリストでの確認が必須(§11.3)。
'
' ■ もう1つの既知の制約(こちらは修正済み): modPrompts が modConfig.GetString/
'   GetLong を無条件に呼んでいたため、modConfigモジュールが注入されていない
'   本ハーネスの純ロジック用一時ライブラリでは実行時エラー420になっていた
'   (これは環境制限ではなくmodPrompts側の明白なバグと判断し、1行スコープの
'   On Error Resume Next + 既定値フォールバックで最小修正した。詳細は
'   src/qa/modPrompts.bas の SafeAnswerLanguage/SafeMaxContextChars 参照。
'   本ファイルのdeviationsとして最終報告に記載する)。
' ============================================================================

' ----------------------------------------------------------------------------
' 取込正規化の回帰テスト(2026-07-27追加)
' ----------------------------------------------------------------------------
' 実物の約款PDFで、条見出しの20.2%(252行中51行)を取りこぼしていた。
' 原因はPDFの字詰めで「第１ ０条」のように数字のあいだへ空白が入ること。
' 第1条～第9条は通り、第10条以降が全滅する ―― 約款の中身の大半である。
' 見出しを落とすとその条は隣の条の本文に吸収され、「第12条は?」に別の条が
' 返る。しかも利用者には理由が永久に分からない。
' 入力の欠陥は出力の全部に効くので、二度と戻さないための固定テスト。
Private Sub RunIngestNormalizeTests()
    modTestRunner.Check "桁割れ: 第1 0条", _
        modChunker.JoinSplitNumbers("第1 0条(定義)") = "第10条(定義)", _
        "実際=" & modChunker.JoinSplitNumbers("第1 0条(定義)")
    modTestRunner.Check "桁割れ: 全角 第１ ０ 条", _
        modChunker.JoinSplitNumbers("第１ ０ 条") = "第10条", _
        "実際=" & modChunker.JoinSplitNumbers("第１ ０ 条")
    modTestRunner.Check "桁割れ: 第 ２ ０ ０ 条", _
        modChunker.JoinSplitNumbers("第 ２ ０ ０ 条") = "第200条", _
        "実際=" & modChunker.JoinSplitNumbers("第 ２ ０ ０ 条")
    modTestRunner.Check "桁割れ: 章にも効く", _
        modChunker.JoinSplitNumbers("第 １ 章 総則") = "第1章 総則", _
        "実際=" & modChunker.JoinSplitNumbers("第 １ 章 総則")

    modTestRunner.Check "桁割れ: 本文の空白は壊さない", _
        modChunker.JoinSplitNumbers("この 保険 は 次の とおり") = "この 保険 は 次の とおり", _
        "過剰置換は本文を壊す"
    modTestRunner.Check "桁割れ: 単位が続かない第は不変", _
        modChunker.JoinSplitNumbers("第 3 者に対して") = "第 3 者に対して", _
        "「第3者」は条番号ではない"

    modTestRunner.Check "正規化後は条見出しとして認識される", _
        modChunker.ClassifyLine(modChunker.JoinSplitNumbers("第１ ２条(保険契約の無効)")) = 2, _
        "2桁の条を拾えないと、その条は隣の条に吸収される"

    modTestRunner.Check "ページ番号行: - 19 -", _
        modChunker.IsPageNumberLine("- 19 -") = True, ""
    modTestRunner.Check "ページ番号行: 全角ダッシュ", _
        modChunker.IsPageNumberLine(ChrW(&H2014) & "19" & ChrW(&H2014)) = True, ""
    modTestRunner.Check "ページ番号行: 本文は落とさない", _
        modChunker.IsPageNumberLine("- 保険金の支払 -") = False, _
        "本文を誤って落とすと資料が欠ける"
    modTestRunner.Check "ページ番号行: 空行は対象外", _
        modChunker.IsPageNumberLine("") = False, ""

    Dim src As String
    src = "- 12 -" & vbLf & "第１ ５条(保険契約の取消)" & vbLf & "当社は、次の場合は取り消します。"
    Dim got As String
    got = modChunker.NormalizeForIngest(src)
    modTestRunner.Check "取込正規化: ページ番号行が消える", InStr(got, "- 12 -") = 0, "実際=" & got
    modTestRunner.Check "取込正規化: 条番号が繋がる", InStr(got, "第15条") > 0, "実際=" & got
    modTestRunner.Check "取込正規化: 本文は残る", InStr(got, "取り消します") > 0, "実際=" & got
End Sub


Public Sub RunAll()
    On Error GoTo UtilFail
    TestModUtil
NextChunker:
    On Error GoTo ChunkerFail
    TestModChunker
NextPii:
    On Error GoTo PiiFail
    TestModPii
NextPure2:
    On Error GoTo Pure2Fail
    modTestsPure2.RunAll2

    RunIngestNormalizeTests
NextDone:
    On Error GoTo 0
    Exit Sub

UtilFail:
    modTestRunner.Check "TestModUtil(グループ全体)", False, "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextChunker
ChunkerFail:
    modTestRunner.Check "TestModChunker(グループ全体)", False, "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextPii
PiiFail:
    modTestRunner.Check "TestModPii(グループ全体)", False, "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextPure2
Pure2Fail:
    modTestRunner.Check "modTestsPure2.RunAll2(グループ全体)", False, "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone
End Sub

' ----------------------------------------------------------------------------
' CanUseTypeArrays - modTypes(別モジュール)のPublic Typeの配列をReDimできる
'   実行環境かどうかを実測で判定する(上部コメント参照)。
' ----------------------------------------------------------------------------
Private Function CanUseTypeArrays() As Boolean
    On Error Resume Next
    Err.Clear
    Dim probe() As ShelfChunk
    ReDim probe(0 To 0)
    CanUseTypeArrays = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' 【2026-07-16 訂正の記録】かつてここに CanUseEmptyArrayReDim という判定関数と
'   「"ReDim arr(0 To -1)" はExcel VBAの標準イディオムで、これをエラー9にする
'   のはLibreOffice側の制限」という長いコメントがあったが、実機Windows Excel
'   のerr_logで全面的に誤りだと証明された: ReDimの上限<下限は**実機Excel VBA
'   でも**実行時エラー9「インデックスが有効範囲にありません」になる
'   (0要素配列を宣言できるのはVB.NETであり、VBA/VB6ではできない)。
'   LOテストがこの文を「環境差」としてスキップしていたため、本番コード20箇所
'   以上に埋まった同イディオムが検出されず、実機で資料取込・同期・画面描画が
'   次々に失敗する事故になった。正しい0件表現は
'   ①count変数+ダミー1要素のReDim(0 To 0)、または
'   ②Split(vbNullString)(言語機能が返す正当な空配列。LBound=0/UBound=-1)。
'   本番コードは全箇所を①②へ移行済み。この教訓を消さないためコメントを残す。
' ----------------------------------------------------------------------------

' ============================================================================
' modUtil
' ============================================================================
Private Sub TestModUtil()
    TestFnv1a64Hex
    TestVectorCsvRoundtrip
    TestCsvToVectorInvalid
    TestL2Normalize
    TestDotProduct
    TestIsSameTimestamp
    TestSafeLeft
    TestFileNameOfExtOf
    TestSplitKeepNonEmpty
    TestHumanBytesSeconds
    TestNormalizeForHash
    TestDeobfuscateSecret
End Sub

Private Sub TestFnv1a64Hex()
    ' 決定性: 同一入力を2回計算して一致すること。
    Dim h1 As String, h2 As String
    h1 = modUtil.Fnv1a64Hex("hello world")
    h2 = modUtil.Fnv1a64Hex("hello world")
    modTestRunner.Check "Fnv1a64Hex_決定性_ASCII", (h1 = h2) And (h1 <> ""), "h1=" & h1 & " h2=" & h2

    ' 形式: 16桁の小文字16進文字列であること。
    Dim isHex16 As Boolean
    isHex16 = (Len(h1) = 16)
    If isHex16 Then
        Dim i As Long
        For i = 1 To 16
            Dim c As String: c = Mid$(h1, i, 1)
            If Not ((c >= "0" And c <= "9") Or (c >= "a" And c <= "f")) Then
                isHex16 = False
                Exit For
            End If
        Next i
    End If
    modTestRunner.Check "Fnv1a64Hex_形式_16桁小文字16進", isHex16, "h1=" & h1

    ' 空文字列: 決定性のみ確認(offset_basisそのものになるはず)。
    Dim hEmpty1 As String, hEmpty2 As String
    hEmpty1 = modUtil.Fnv1a64Hex("")
    hEmpty2 = modUtil.Fnv1a64Hex("")
    modTestRunner.Check "Fnv1a64Hex_決定性_空文字", (hEmpty1 = hEmpty2) And (Len(hEmpty1) = 16), "hEmpty1=" & hEmpty1

    ' 日本語: 決定性+ASCII結果との非衝突(ヒューリスティック確認)。
    Dim hJp1 As String, hJp2 As String
    hJp1 = modUtil.Fnv1a64Hex("こんにちは世界")
    hJp2 = modUtil.Fnv1a64Hex("こんにちは世界")
    modTestRunner.Check "Fnv1a64Hex_決定性_日本語", (hJp1 = hJp2) And (hJp1 <> h1), "hJp1=" & hJp1

    ' 絵文字混在(サロゲートペア): 決定性のみ確認(クラッシュしないことも兼ねる)。
    Dim hEmoji1 As String, hEmoji2 As String
    Dim emojiText As String: emojiText = "資料A" & ChrW(&HD83D) & ChrW(&HDE00) & "です"   ' U+1F600 のサロゲートペア
    hEmoji1 = modUtil.Fnv1a64Hex(emojiText)
    hEmoji2 = modUtil.Fnv1a64Hex(emojiText)
    modTestRunner.Check "Fnv1a64Hex_決定性_絵文字混在", (hEmoji1 = hEmoji2) And (Len(hEmoji1) = 16), "hEmoji1=" & hEmoji1

    ' 相互一致: 異なる入力からは(通常は)異なるハッシュが得られること(衝突の無いことの確認ではなく、
    ' 定数実装ミス=常に同じ値を返す、のような明白なバグを検出するための健全性チェック)。
    Dim allDifferent As Boolean
    allDifferent = (h1 <> hEmpty1) And (h1 <> hJp1) And (h1 <> hEmoji1) And (hEmpty1 <> hJp1) And (hEmpty1 <> hEmoji1) And (hJp1 <> hEmoji1)
    modTestRunner.Check "Fnv1a64Hex_異なる入力は異なる値", allDifferent, _
        "h1=" & h1 & " hEmpty1=" & hEmpty1 & " hJp1=" & hJp1 & " hEmoji1=" & hEmoji1

    ' 5000字の長文: MulU64ByPrime(桁上げ処理)が長い入力でも決定的であることの確認。
    ' Wave1申し送り事項(Hex$の負数Long挙動・非BMP文字のLen/Mid挙動)は、この長文が
    ' 絵文字混在チェック同様に大量のFNV反復(=多様なビットパターンのhHi/hLoが
    ' Hex$()に渡る)を実際に走らせることで、LO実行環境でも間接的に検証される
    ' (個々のビットパターンを狙って作り込むのではなく、大量反復による健全性確認)。
    Dim h5000_1 As String, h5000_2 As String
    Dim long5000 As String
    ' 日本語・数字・非BMP絵文字(サロゲートペア)・改行を含む7字単位を繰り返し連結し、
    ' Left$で厳密に5000字へ切り詰める(§12: &連鎖の長大化を避け、ループ+Left$で構成)。
    Dim repeatUnit As String: repeatUnit = "あい5" & ChrW(&HD83D) & ChrW(&HDE00) & vbLf & "x"
    Dim sb5000 As String
    Do While Len(sb5000) < 5000
        sb5000 = sb5000 & repeatUnit
    Loop
    long5000 = Left$(sb5000, 5000)
    modTestRunner.Check "Fnv1a64Hex_前提_5000字ちょうど", (Len(long5000) = 5000), "len=" & Len(long5000)

    h5000_1 = modUtil.Fnv1a64Hex(long5000)
    h5000_2 = modUtil.Fnv1a64Hex(long5000)
    modTestRunner.Check "Fnv1a64Hex_決定性_5000字", (h5000_1 = h5000_2) And (Len(h5000_1) = 16), "h5000_1=" & h5000_1

    Dim isHex16_5000 As Boolean: isHex16_5000 = (Len(h5000_1) = 16)
    If isHex16_5000 Then
        Dim j As Long
        For j = 1 To 16
            Dim cj As String: cj = Mid$(h5000_1, j, 1)
            If Not ((cj >= "0" And cj <= "9") Or (cj >= "a" And cj <= "f")) Then
                isHex16_5000 = False
                Exit For
            End If
        Next j
    End If
    modTestRunner.Check "Fnv1a64Hex_5000字_形式_16桁小文字16進", isHex16_5000, "h5000_1=" & h5000_1
End Sub

Private Sub TestVectorCsvRoundtrip()
    Dim v() As Double: ReDim v(0 To 3)
    v(0) = 1.5: v(1) = -2.25: v(2) = 0#: v(3) = 12345.6789

    Dim csv As String
    csv = modUtil.VectorToCsv(v)
    modTestRunner.Check "VectorToCsv_カンマ区切り", (InStr(csv, ",") > 0) And (LenB(csv) > 0), "csv=" & csv

    Dim back() As Double
    Dim ok As Boolean
    ok = modUtil.CsvToVector(csv, back)
    modTestRunner.Check "CsvToVector_往復成功", ok, "csv=" & csv

    Dim allMatch As Boolean: allMatch = True
    If ok Then
        Dim n As Long: n = UBound(back) - LBound(back) + 1
        If n <> 4 Then
            allMatch = False
        Else
            Dim i As Long
            For i = 0 To 3
                If Abs(back(i) - v(i)) > 0.0000001 Then allMatch = False
            Next i
        End If
    Else
        allMatch = False
    End If
    modTestRunner.Check "VectorToCsv_CsvToVector_往復一致", allMatch, "csv=" & csv
End Sub

Private Sub TestCsvToVectorInvalid()
    Dim back() As Double

    ' 空文字列は不正
    Dim ok1 As Boolean: ok1 = modUtil.CsvToVector("", back)
    modTestRunner.Check "CsvToVector_不正入力_空文字False", (Not ok1), "ok1=" & ok1

    ' 連続カンマ(空要素)は不正
    Dim ok2 As Boolean: ok2 = modUtil.CsvToVector("1,,3", back)
    modTestRunner.Check "CsvToVector_不正入力_連続カンマFalse", (Not ok2), "ok2=" & ok2

    ' 末尾カンマ(空要素)も不正
    Dim ok3 As Boolean: ok3 = modUtil.CsvToVector("1,2,", back)
    modTestRunner.Check "CsvToVector_不正入力_末尾カンマFalse", (Not ok3), "ok3=" & ok3

    ' 正常な1要素は成功する対照確認
    Dim ok4 As Boolean: ok4 = modUtil.CsvToVector("42", back)
    modTestRunner.Check "CsvToVector_正常1要素_True", ok4 And (UBound(back) - LBound(back) + 1 = 1), "ok4=" & ok4
End Sub

Private Sub TestL2Normalize()
    ' ゼロベクトルはFalse・値は変更されない
    Dim zeroVec() As Double: ReDim zeroVec(0 To 2)
    zeroVec(0) = 0: zeroVec(1) = 0: zeroVec(2) = 0
    Dim okZero As Boolean: okZero = modUtil.L2Normalize(zeroVec)
    modTestRunner.Check "L2Normalize_ゼロベクトルFalse", (Not okZero), "okZero=" & okZero

    ' 非ゼロベクトルはTrue・単位ノルムになる
    Dim v() As Double: ReDim v(0 To 2)
    v(0) = 3: v(1) = 4: v(2) = 0
    Dim okV As Boolean: okV = modUtil.L2Normalize(v)
    Dim sumSq As Double: sumSq = v(0) * v(0) + v(1) * v(1) + v(2) * v(2)
    modTestRunner.Check "L2Normalize_成功True", okV, "okV=" & okV
    modTestRunner.Check "L2Normalize_単位ノルム", (Abs(sumSq - 1#) < 0.0000001), "sumSq=" & sumSq
    modTestRunner.Check "L2Normalize_値が正しい比率", (Abs(v(0) - 0.6) < 0.0000001) And (Abs(v(1) - 0.8) < 0.0000001), _
        "v0=" & v(0) & " v1=" & v(1)
End Sub

Private Sub TestDotProduct()
    Dim a() As Double: ReDim a(0 To 2)
    a(0) = 1: a(1) = 2: a(2) = 3
    Dim b() As Double: ReDim b(0 To 2)
    b(0) = 4: b(1) = 5: b(2) = 6
    Dim d As Double: d = modUtil.DotProduct(a, b)
    ' 1*4+2*5+3*6 = 32
    modTestRunner.Check "DotProduct_既知の値", (Abs(d - 32#) < 0.0000001), "d=" & d

    ' 次元不一致は0を返す(呼び出し側が別途次元検査する規約)
    Dim c() As Double: ReDim c(0 To 1)
    c(0) = 1: c(1) = 1
    Dim d2 As Double: d2 = modUtil.DotProduct(a, c)
    modTestRunner.Check "DotProduct_次元不一致は0", (d2 = 0#), "d2=" & d2
End Sub

' 注意(2026-07-11 Wave3で特定): ローカル変数名に "base" を使うと、
' LibreOffice Basic(Option VBASupport 1環境)のコンパイルが応答不能になり
' run_lo_tests.pyがタイムアウト(exit 124)する、実験で切り分け済みの地雷が
' ある("Base" が "Option Base" 文のキーワードと衝突するためと推測される)。
' Excel VBAでは変数名として合法だが、本ハーネス互換のため "baseTs" を使う。
Private Sub TestIsSameTimestamp()
    Dim baseTs As Date: baseTs = DateSerial(2024, 1, 1)   ' 固定基準日時(0時0分0秒)

    ' 0秒差 -> True
    Dim same0 As Boolean: same0 = modUtil.IsSameTimestamp(baseTs, baseTs)
    modTestRunner.Check "IsSameTimestamp_0秒差True", same0, "same0=" & same0

    ' 1.9秒差(2秒丸めの範囲内) -> True
    Dim b19 As Date: b19 = baseTs + (1.9# / 86400#)
    Dim same19 As Boolean: same19 = modUtil.IsSameTimestamp(baseTs, b19)
    modTestRunner.Check "IsSameTimestamp_1_9秒差True", same19, "same19=" & same19

    ' 2.1秒差(2秒丸めの範囲外) -> False
    Dim b21 As Date: b21 = baseTs + (2.1# / 86400#)
    Dim same21 As Boolean: same21 = modUtil.IsSameTimestamp(baseTs, b21)
    modTestRunner.Check "IsSameTimestamp_2_1秒差False", (Not same21), "same21=" & same21

    ' 前後どちらが新しくても対称であること(bの方が過去)
    Dim same19rev As Boolean: same19rev = modUtil.IsSameTimestamp(b19, baseTs)
    modTestRunner.Check "IsSameTimestamp_対称性", same19rev, "same19rev=" & same19rev
End Sub

Private Sub TestSafeLeft()
    modTestRunner.Check "SafeLeft_短い文字列はそのまま", (modUtil.SafeLeft("abc", 10) = "abc")
    modTestRunner.Check "SafeLeft_ちょうどの長さ", (modUtil.SafeLeft("abcde", 5) = "abcde")
    modTestRunner.Check "SafeLeft_長い文字列は切り詰め", (modUtil.SafeLeft("abcdefgh", 3) = "abc")
    modTestRunner.Check "SafeLeft_負の上限は空文字扱い", (modUtil.SafeLeft("abc", -1) = "")
End Sub

Private Sub TestFileNameOfExtOf()
    ' \ と / が混在するパス
    modTestRunner.Check "FileNameOf_バックスラッシュ", _
        (modUtil.FileNameOf("C:\Users\test\資料フォルダ\file.txt") = "file.txt")
    modTestRunner.Check "FileNameOf_スラッシュ", _
        (modUtil.FileNameOf("C:/Users/test/file.txt") = "file.txt")
    modTestRunner.Check "FileNameOf_混在", _
        (modUtil.FileNameOf("C:\Users/test\sub/file.txt") = "file.txt")
    modTestRunner.Check "FileNameOf_区切りなし", (modUtil.FileNameOf("file.txt") = "file.txt")

    modTestRunner.Check "ExtOf_小文字化", (modUtil.ExtOf("C:\a\b\FILE.PDF") = "pdf")
    modTestRunner.Check "ExtOf_拡張子なし", (modUtil.ExtOf("C:\a\b\README") = "")
    modTestRunner.Check "ExtOf_末尾ドットのみ", (modUtil.ExtOf("C:\a\b\file.") = "")
    modTestRunner.Check "ExtOf_混在パス", (modUtil.ExtOf("C:/a\b/file.DOCX") = "docx")
End Sub

Private Sub TestSplitKeepNonEmpty()
    Dim r() As String
    r = modUtil.SplitKeepNonEmpty("a,,b, ,c", ",")
    Dim n As Long: n = UBound(r) - LBound(r) + 1
    modTestRunner.Check "SplitKeepNonEmpty_件数", (n = 3), "n=" & n
    If n = 3 Then
        modTestRunner.Check "SplitKeepNonEmpty_内容", (r(0) = "a") And (r(1) = "b") And (r(2) = "c"), _
            "r0=" & r(0) & " r1=" & r(1) & " r2=" & r(2)
    End If

    ' 空文字列入力(0件)の確認。SplitKeepNonEmptyはSplit(vbNullString)方式の
    ' 空配列(LBound=0/UBound=-1)を返すので、この確認はLO・実機Excelの両環境で
    ' 無条件に実行できる(2026-07-16: 誤った「LO制限スキップ」を撤廃)。
    Dim rEmpty() As String
    rEmpty = modUtil.SplitKeepNonEmpty("", ",")
    Dim nEmpty As Long: nEmpty = UBound(rEmpty) - LBound(rEmpty) + 1
    modTestRunner.Check "SplitKeepNonEmpty_空文字入力は0件", (nEmpty <= 0), "nEmpty=" & nEmpty

    ' 空白のみの入力(全要素が除去されて0件になるパス)も同様に確認。
    Dim rBlank() As String
    rBlank = modUtil.SplitKeepNonEmpty(" , , ", ",")
    Dim nBlank As Long: nBlank = UBound(rBlank) - LBound(rBlank) + 1
    modTestRunner.Check "SplitKeepNonEmpty_空白のみ入力は0件", (nBlank <= 0), "nBlank=" & nBlank
End Sub

' HumanBytes/HumanSecondsのテスト。
' 注意(LO実機確認事項): HumanBytesはKB以上でFormat$(v,"0.0")を使うため、
' LibreOffice実行環境のロケール設定次第で小数点区切りが "." ではなく ","
' になる可能性がある(tools/README.md「日付/ロケール/文字コード周りの細かな
' 挙動…はLOとExcelで一致しない場合がある」)。本テストは小数点区切りを
' "," -> "." に正規化してから比較することでこの差を吸収する
' (Excel実機では常に "." であることを別途確認すること)。
Private Sub TestHumanBytesSeconds()
    ' B単位はCStr(CLng(...))を使うため小数点の心配がない
    modTestRunner.Check "HumanBytes_B単位", (modUtil.HumanBytes(500) = "500 B"), "r=" & modUtil.HumanBytes(500)

    ' KB単位: ロケール差を吸収して比較
    Dim rKb As String: rKb = modUtil.HumanBytes(1536)   ' 1536/1024 = 1.5
    Dim rKbNorm As String: rKbNorm = Replace(rKb, ",", ".")
    modTestRunner.Check "HumanBytes_KB単位", (rKbNorm = "1.5 KB"), "r=" & rKb

    ' MB単位
    Dim rMb As String: rMb = modUtil.HumanBytes(1024# * 1024# * 2#)   ' 2.0 MB
    Dim rMbNorm As String: rMbNorm = Replace(rMb, ",", ".")
    modTestRunner.Check "HumanBytes_MB単位", (rMbNorm = "2.0 MB"), "r=" & rMb

    ' 負値は0扱い
    modTestRunner.Check "HumanBytes_負値は0扱い", (modUtil.HumanBytes(-100) = "0 B"), "r=" & modUtil.HumanBytes(-100)

    ' HumanSecondsは文字列連結のみでロケール非依存
    modTestRunner.Check "HumanSeconds_秒未満", (modUtil.HumanSeconds(45) = "約45秒"), "r=" & modUtil.HumanSeconds(45)
    modTestRunner.Check "HumanSeconds_分ちょうど", (modUtil.HumanSeconds(120) = "約2分"), "r=" & modUtil.HumanSeconds(120)
    modTestRunner.Check "HumanSeconds_分と秒", (modUtil.HumanSeconds(125) = "約2分5秒"), "r=" & modUtil.HumanSeconds(125)
    modTestRunner.Check "HumanSeconds_負値は0扱い", (modUtil.HumanSeconds(-5) = "約0秒"), "r=" & modUtil.HumanSeconds(-5)
End Sub

Private Sub TestNormalizeForHash()
    ' 空白圧縮(連続する半角スペース・タブを1個に)
    Dim r1 As String: r1 = modUtil.NormalizeForHash("a   b" & vbTab & vbTab & "c")
    modTestRunner.Check "NormalizeForHash_空白圧縮", (r1 = "a b c"), "r1=" & r1

    ' 改行統一(CRLF/CRをLFへ)。改行そのものは圧縮しない(段落境界保持)。
    Dim r2 As String: r2 = modUtil.NormalizeForHash("line1" & vbCrLf & "line2" & vbCr & "line3")
    modTestRunner.Check "NormalizeForHash_改行統一", (r2 = "line1" & vbLf & "line2" & vbLf & "line3"), "r2=" & r2

    ' Trim(前後の空白除去)
    Dim r3 As String: r3 = modUtil.NormalizeForHash("   前後空白あり   ")
    modTestRunner.Check "NormalizeForHash_Trim", (r3 = "前後空白あり"), "r3=" & r3

    ' 空文字列は空文字列のまま
    modTestRunner.Check "NormalizeForHash_空文字", (modUtil.NormalizeForHash("") = "")
End Sub

' build/build_mybookshelf.py obfuscate_secret() とVBA側DeobfuscateSecretが
' 同じXOR+16進アルゴリズムで対になっていることを固定値で検証する
' (Python側で生成した既知の変換結果をここへ転記。鍵やアルゴリズムを
' 片側だけ変えると密かに壊れるため、往復ではなく既知値の一致で検出する)。
Private Sub TestDeobfuscateSecret()
    modTestRunner.Check "DeobfuscateSecret_短い文字列", _
        (modUtil.DeobfuscateSecret("OBF1:260014191c") = "hello")

    modTestRunner.Check "DeobfuscateSecret_APIキー長の文字列", _
        (modUtil.DeobfuscateSecret("OBF1:7f014d41462055535f4171145d0a562c5b5f4512565547595c5f2f524e500904") _
            = "1d545a26153a4f2c990a543031d77b96")

    ' 接頭辞が無い値(手動でconfigに平文キーを入力した場合等)はそのまま返す
    modTestRunner.Check "DeobfuscateSecret_接頭辞なしはそのまま", _
        (modUtil.DeobfuscateSecret("plain-manual-key") = "plain-manual-key")

    ' 空文字列は空文字列のまま
    modTestRunner.Check "DeobfuscateSecret_空文字", (modUtil.DeobfuscateSecret("") = "")
End Sub

' ============================================================================
' modChunker
' ============================================================================
Private Sub TestModChunker()
    If Not CanUseTypeArrays() Then
        modTestRunner.Check "modChunker: LO環境の既知の制限によりスキップ", True, _
            "LibreOffice実行環境ではmodTypesのPublic Type配列(ExtractedPage()/" & _
            "ShelfChunk())をReDimすると実行時エラー420になることを確認済み" & _
            "(本モジュール冒頭コメント参照。tools/run_lo_tests.pyは担当外のため変更不可)。" & _
            "ChunkPagesの動作(空ページ→0/1文→1チャンク/長文+オーバーラップ/32000字上限/" & _
            "段落境界優先/複数ページのpage引継ぎ)はコードレビューで確認済みだが、" & _
            "Excel実機受入チェック(§11.3)で必ず再確認すること。"
        Exit Sub
    End If

    TestChunkPages_EmptyPages
    TestChunkPages_OneShortSentence
    TestChunkPages_LongTextWithOverlap
    TestChunkPages_MaxCharsGuarantee
    TestChunkPages_ParagraphBoundary
    TestChunkPages_MultiPagePageNumber
End Sub

Private Sub TestChunkPages_EmptyPages()
    Dim pages() As ExtractedPage
    Dim chunks() As ShelfChunk
    Dim n As Long
    n = modChunker.ChunkPages(pages, 700, 150, chunks)
    modTestRunner.Check "ChunkPages_空ページ配列は0件", (n = 0), "n=" & n
End Sub

Private Sub TestChunkPages_OneShortSentence()
    Dim pages(0 To 0) As ExtractedPage
    pages(0).page = 1
    pages(0).Text = "これは短い一文のテストです。"

    Dim chunks() As ShelfChunk
    Dim n As Long
    n = modChunker.ChunkPages(pages, 700, 150, chunks)
    modTestRunner.Check "ChunkPages_短文1つは1チャンク", (n = 1), "n=" & n
    If n = 1 Then
        modTestRunner.Check "ChunkPages_短文チャンクのpage引継ぎ", (chunks(0).page = 1), "page=" & chunks(0).page
        modTestRunner.Check "ChunkPages_短文チャンクの本文が含まれる", _
            (InStr(chunks(0).full_text, "短い一文") > 0), "full_text=" & chunks(0).full_text
    End If
End Sub

Private Sub TestChunkPages_LongTextWithOverlap()
    ' 句読点や改行を含まない(境界スナップが発生しない)数字の反復2000字を用意し、
    ' target=500 / overlap=100 でのウィンドウ位置を手計算どおりに検証する。
    Dim raw As String
    Dim parts(1 To 200) As String
    Dim i As Long
    For i = 1 To 200
        parts(i) = "0123456789"
    Next i
    raw = Join(parts, "")   ' 2000字ちょうど、句読点なし
    modTestRunner.Check "ChunkPages_前提_raw2000字", (Len(raw) = 2000), "len=" & Len(raw)

    Dim pages(0 To 0) As ExtractedPage
    pages(0).page = 7
    pages(0).Text = raw

    Dim chunks() As ShelfChunk
    Dim n As Long
    n = modChunker.ChunkPages(pages, 500, 100, chunks)

    ' 手計算: start=1/401/801/1201/1601 の5チャンクになるはず(モジュール冒頭コメント参照の
    ' アルゴリズムどおりの手計算。詳細はこのテストを書いた際の設計メモとしてWave3以降の
    ' レビューでも再現できるよう、コメントに計算過程を残す:
    '   1..500 / 401..900 / 801..1300 / 1201..1700 / 1601..2000)
    modTestRunner.Check "ChunkPages_長文は複数チャンク", (n = 5), "n=" & n

    If n >= 2 Then
        ' オーバーラップ実在検証: チャンク1の末尾100字とチャンク2の先頭100字が一致すること
        Dim tail1 As String: tail1 = Right$(chunks(0).full_text, 100)
        Dim head2 As String: head2 = Left$(chunks(1).full_text, 100)
        modTestRunner.Check "ChunkPages_オーバーラップ実在", (tail1 = head2) And (Len(tail1) = 100), _
            "tail1=" & tail1 & " head2=" & head2
        modTestRunner.Check "ChunkPages_チャンク1の長さ", (Len(chunks(0).full_text) = 500), "len=" & Len(chunks(0).full_text)
    End If
End Sub

Private Sub TestChunkPages_MaxCharsGuarantee()
    ' targetCharsに32000超の異常値を渡しても、full_textは絶対に32000字を超えない
    ' (modChunker冒頭コメント: MAX_CHUNK_CHARSによる二重防御)。33000字の1ページ入力。
    Dim raw As String
    raw = String(33000, "A")   ' 句読点なし=境界スナップも発生しない

    Dim pages(0 To 0) As ExtractedPage
    pages(0).page = 1
    pages(0).Text = raw

    Dim chunks() As ShelfChunk
    Dim n As Long
    n = modChunker.ChunkPages(pages, 50000, 0, chunks)   ' 50000は異常値(32000超)

    modTestRunner.Check "ChunkPages_33000字は複数チャンクに分割", (n >= 2), "n=" & n

    Dim allWithinLimit As Boolean: allWithinLimit = True
    Dim totalLen As Long: totalLen = 0
    Dim i As Long
    For i = 0 To n - 1
        If Len(chunks(i).full_text) > 32000 Then allWithinLimit = False
        totalLen = totalLen + Len(chunks(i).full_text)
    Next i
    modTestRunner.Check "ChunkPages_全チャンクが32000字以内", allWithinLimit, "n=" & n
    modTestRunner.Check "ChunkPages_33000字入力の合計文字数一致", (totalLen = 33000), "totalLen=" & totalLen
    If n >= 1 Then
        modTestRunner.Check "ChunkPages_先頭チャンクは32000字ちょうど", (Len(chunks(0).full_text) = 32000), _
            "len=" & Len(chunks(0).full_text)
    End If
End Sub

Private Sub TestChunkPages_ParagraphBoundary()
    ' 100字の"A"+句点+400字の"B"という構成で、target=90(句点の少し手前)を指定すると
    ' 文境界(句点)にスナップしてチャンク1が句点で終わるはず(modChunker冒頭コメントの
    ' 「段落・文境界優先」動作の確認)。
    Dim raw As String
    raw = String(100, "A") & ChrW(&H3002) & String(400, "B")   ' 。(全角句点)は位置101

    Dim pages(0 To 0) As ExtractedPage
    pages(0).page = 1
    pages(0).Text = raw

    Dim chunks() As ShelfChunk
    Dim n As Long
    n = modChunker.ChunkPages(pages, 90, 0, chunks)

    modTestRunner.Check "ChunkPages_段落境界_複数チャンク生成", (n >= 1), "n=" & n
    If n >= 1 Then
        modTestRunner.Check "ChunkPages_段落境界_句点で終わる", (Right$(chunks(0).full_text, 1) = ChrW(&H3002)), _
            "tail=" & Right$(chunks(0).full_text, 3)
        modTestRunner.Check "ChunkPages_段落境界_長さ101", (Len(chunks(0).full_text) = 101), _
            "len=" & Len(chunks(0).full_text)
    End If
End Sub

Private Sub TestChunkPages_MultiPagePageNumber()
    Dim pages(0 To 2) As ExtractedPage
    pages(0).page = 1: pages(0).Text = "1ページ目の短い本文です。"
    pages(1).page = 5: pages(1).Text = "5ページ目の短い本文です。"
    pages(2).page = 9: pages(2).Text = "9ページ目の短い本文です。"

    Dim chunks() As ShelfChunk
    Dim n As Long
    n = modChunker.ChunkPages(pages, 700, 150, chunks)

    modTestRunner.Check "ChunkPages_複数ページは各1チャンク", (n = 3), "n=" & n
    If n = 3 Then
        modTestRunner.Check "ChunkPages_page引継ぎ_1", (chunks(0).page = 1), "page=" & chunks(0).page
        modTestRunner.Check "ChunkPages_page引継ぎ_5", (chunks(1).page = 5), "page=" & chunks(1).page
        modTestRunner.Check "ChunkPages_page引継ぎ_9", (chunks(2).page = 9), "page=" & chunks(2).page
    End If
End Sub

' ============================================================================
' modPii
' ============================================================================
Private Sub TestModPii()
    ' 検知例: メールアドレス
    Dim r1 As String: r1 = modPii.ScanText("お問い合わせは taro.yamada@example.co.jp まで")
    modTestRunner.Check "ScanText_メールアドレス検知", (InStr(r1, "メールアドレス") > 0), "r1=" & r1

    ' 検知例: 10桁以上の連続数字(ハイフン区切りの電話番号)
    Dim r2 As String: r2 = modPii.ScanText("電話番号: 090-1234-5678 です")
    modTestRunner.Check "ScanText_長い数字列検知_ハイフン区切り", (InStr(r2, "長い数字列") > 0), "r2=" & r2

    ' 検知例: 10桁連続数字(区切りなし)
    Dim r3 As String: r3 = modPii.ScanText("マイナンバーは1234567890です")
    modTestRunner.Check "ScanText_長い数字列検知_区切りなし", (InStr(r3, "長い数字列") > 0), "r3=" & r3

    ' 両方検知される複合例
    Dim r4 As String: r4 = modPii.ScanText("連絡先: a@b.com / 09012345678")
    modTestRunner.Check "ScanText_複合検知_メール", (InStr(r4, "メールアドレス") > 0), "r4=" & r4
    modTestRunner.Check "ScanText_複合検知_数字列", (InStr(r4, "長い数字列") > 0), "r4=" & r4

    ' 非検知例: 普通の業務文章
    Dim r5 As String: r5 = modPii.ScanText("本日の会議は10時から会議室Aで行います。")
    modTestRunner.Check "ScanText_非検知_通常文章", (r5 = ""), "r5=" & r5

    ' 非検知例: 9桁の数字(10桁未満はセーフ)
    Dim r6 As String: r6 = modPii.ScanText("商品コード: 123456789")
    modTestRunner.Check "ScanText_非検知_9桁数字", (r6 = ""), "r6=" & r6

    ' 非検知例: 空文字列
    Dim r7 As String: r7 = modPii.ScanText("")
    modTestRunner.Check "ScanText_非検知_空文字", (r7 = ""), "r7=" & r7
End Sub
