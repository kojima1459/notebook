Attribute VB_Name = "modTestsPure41"
Option Explicit

' ============================================================================
' modTestsPure41 - R36 波4(版上げ時の自動引き継ぎ)の純ロジック回帰テスト。
'   既存チェーン(modTestsPure.RunAll→…)へは繋がず、
'   modTestRunner.RunAllPureTests から直接呼ばれる RunAll41 の1本が入口
'   (modTestsPure31〜40 と同型の別枝)。
' ----------------------------------------------------------------------------
' 【何を固定するか】
'   D1 modMigrateFrom.PickNewest: 「自分を除く・最新・同時刻は先勝ち・
'     候補0件は-1」の4条件。自分が全候補中いちばん新しい更新日時を
'     持っていても、自分自身は絶対に選ばれないことを押さえる(ここが
'     崩れると「前の版」のつもりで自分自身を Workbooks.Open してしまう
'     事故になる)。
'   D2 modMigrateFrom.OfferText: 確認ダイアログの文面にファイル名と
'     更新日時(yyyy-mm-dd hh:nn)が必ず載ること。
'
' 【なぜ FindPreviousBook/ImportFromBook/CopyUserData/OfferImportIfFirstRun の
'   テストが無いか】
'   いずれも Dir()・Workbooks.Open・シートI/Oを持ち、Excel/COMに実オブジェクトで
'   触れる(modShelfSync/modPack と同じ「モジュール全体はR4準拠ではないが、
'   テストが呼ぶ関数自体はExcel/COMに触れない」型のうち、今回テストに
'   採るのは純関数2本だけ)。配列操作の確認は実機(modTestsExcel系)と
'   R36 §6 の実機受入④に委ねる。
' ============================================================================

' ---- D1: PickNewest の4条件 --------------------------------------------------
'   discriminate:
'   ・自分を除く判定を落とす(全候補から単純に最新を選ぶ)実装にすると、
'     自分自身がいちばん新しいケース(a)が「自分」を返して落ちる。
'   ・同時刻の比較を ">" ではなく ">=" にすると、同時刻ケース(b)の
'     「先に見つかった方(idx=0)」が「後から見つかった方(idx=1)」に
'     置き換わって落ちる。
'   ・自分の判定を大小区別あり(vbBinaryCompare)にすると、大文字小文字が
'     違うだけの自分自身(c)を「他人」と誤認して選んでしまい落ちる
'     (Windowsのファイル名は大小を区別しないため、ここは vbTextCompare が正)。
'   ・候補0件(全部自分/配列が空)で -1 以外を返す実装は(d)(e)が落ちる。
Private Sub TestPickNewest41()
    Dim names(0 To 2) As String
    Dim stamps(0 To 2) As Double

    ' (a) 自分がいちばん新しくても、自分は絶対に選ばれない。
    names(0) = "MyBookshelf.xlsm": stamps(0) = 500      ' 自分・最新だが除外対象
    names(1) = "MyBookshelf_前回.xlsm": stamps(1) = 200
    names(2) = "MyBookshelf_old.xlsm": stamps(2) = 300
    ChkLong41 "D1_自分は最新でも除外し他人の最新(idx2)を選ぶ", _
        modMigrateFrom.PickNewest(names, stamps, 3, "MyBookshelf.xlsm"), 2

    ' (b) 同時刻は先勝ち(先に見つかったidx0を保つ。">="にすると1へずれる)。
    Dim names2(0 To 1) As String
    Dim stamps2(0 To 1) As Double
    names2(0) = "A.xlsm": stamps2(0) = 100
    names2(1) = "B.xlsm": stamps2(1) = 100
    ChkLong41 "D1_同時刻は先勝ち(idx0)", _
        modMigrateFrom.PickNewest(names2, stamps2, 2, "selfNotPresent.xlsm"), 0

    ' (c) 自分の判定は大小区別しない(Windowsのファイル名と同じ)。
    ChkLong41 "D1_自分判定は大小区別しない(大文字小文字違いも除外)", _
        modMigrateFrom.PickNewest(names, stamps, 3, "MYBOOKSHELF.XLSM"), 2

    ' (d) 候補が全部自分 → -1。
    Dim names3(0 To 0) As String
    Dim stamps3(0 To 0) As Double
    names3(0) = "MyBookshelf.xlsm": stamps3(0) = 999
    ChkLong41 "D1_全部自分なら-1", modMigrateFrom.PickNewest(names3, stamps3, 1, "MyBookshelf.xlsm"), -1

    ' (e) 候補0件(n=0) → -1。
    ChkLong41 "D1_候補0件(n=0)は-1", modMigrateFrom.PickNewest(names, stamps, 0, "MyBookshelf.xlsm"), -1
End Sub

' ---- D2: OfferText の文面 -----------------------------------------------------
'   discriminate:
'   ・ファイル名を文面に差し込み忘れる実装は(a)が落ちる。
'   ・Format$の書式を"yyyy-mm-dd hh:nn"以外(区切りや秒の有無違い)にすると
'     (b)が落ちる。
'   ・確認の問いそのものを削る/文言を変える実装は(c)が落ちる。
Private Sub TestOfferText41()
    Dim stamp As Double: stamp = CDbl(DateSerial(2026, 9, 5) + TimeSerial(14, 30, 0))
    Dim msg As String: msg = modMigrateFrom.OfferText("MyBookshelf_前回.xlsm", stamp)

    ' (a) ファイル名が載る。
    ChkBool41 "D2_ファイル名が載る", (InStr(1, msg, "MyBookshelf_前回.xlsm", vbBinaryCompare) > 0), True

    ' (b) 更新日時が "yyyy-mm-dd hh:nn" 形式で載る。
    ChkBool41 "D2_更新日時が載る", (InStr(1, msg, "2026-09-05 14:30", vbBinaryCompare) > 0), True

    ' (c) Yes/No で聞く確認文そのものが載る。
    ChkBool41 "D2_確認の問いが載る", _
        (InStr(1, msg, "この版へ本棚・実績・設定を引き継ぎますか?", vbBinaryCompare) > 0), True

    ' (d) 空のファイル名でも落ちない(空文字を返さない=文面の骨格は残る)。
    Dim emptMsg As String: emptMsg = modMigrateFrom.OfferText("", stamp)
    ChkBool41 "D2_ファイル名が空でも問いは残る", _
        (InStr(1, emptMsg, "この版へ本棚・実績・設定を引き継ぎますか?", vbBinaryCompare) > 0), True
End Sub

' ---- 判定ヘルパー -----------------------------------------------------------
Private Sub ChkBool41(ByVal label As String, ByVal got As Boolean, ByVal want As Boolean)
    modTestRunner.Check "R36-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkLong41(ByVal label As String, ByVal got As Long, ByVal want As Long)
    modTestRunner.Check "R36-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Public Sub RunAll41()
    On Error GoTo H01Fail41
    TestPickNewest41
H02Next41:
    On Error GoTo H02Fail41
    TestOfferText41
H01Done41:
    On Error GoTo 0
    Exit Sub

H01Fail41:
    modTestRunner.Check "TestPickNewest41(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H02Next41
H02Fail41:
    modTestRunner.Check "TestOfferText41(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done41
End Sub
