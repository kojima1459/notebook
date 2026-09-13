Attribute VB_Name = "modTestsPure41"
Option Explicit

' ============================================================================
' modTestsPure41 - R36 波4(版上げ時の自動引き継ぎ)の純ロジック回帰テスト。
'   既存チェーン(modTestsPure.RunAll→…)へは繋がず、
'   modTestRunner.RunAllPureTests から直接呼ばれる RunAll41 の1本が入口
'   (modTestsPure31〜40 と同型の別枝)。
' ----------------------------------------------------------------------------
' 【何を固定するか(R36 Fix2 §9-7で全面差し替え)】
'   D1 modMigrateFrom.PreviousBookName: "<stem>_旧版<拡張子>" の組み立て
'     (拡張子が無い/selfNameが空 → 空文字)。旧D1(PickNewest)はFix2で
'     候補が1個に絞られて役目を失ったため削除した(§9-7確定スコープ)。
'   D2 modMigrateFrom.StampText: "yyyy-mm-dd hh:nn" 形式(1桁の月/日/時/分の
'     ゼロ埋めを含む)。Format$の代わりにYear/Month/Day/Hour/Minuteで
'     組み立てる版(LibreOffice Basicが分の書式"n"を解釈しない死角の回避
'     ・§9-4)。
'   D3 modMigrateFrom.OfferText: 確認ダイアログの文面にファイル名・件数
'     ・更新日時(yyyy-mm-dd hh:nn)が必ず載ること。
'
' 【なぜ FindPreviousBook/ImportFromBook/CopyUserData/OfferImportIfFirstRun/
'   CountManifestInBook のテストが無いか】
'   いずれも Dir()・Workbooks.Open・シートI/Oを持ち、Excel/COMに実オブジェクトで
'   触れる(modShelfSync/modPack と同じ「モジュール全体はR4準拠ではないが、
'   テストが呼ぶ関数自体はExcel/COMに触れない」型のうち、今回テストに
'   採るのは純関数3本だけ)。配列操作の確認は実機(modTestsExcel系)と
'   R36 §6 の実機受入④に委ねる。
' ============================================================================

' ---- D1: PreviousBookName の組み立て ----------------------------------------
'   discriminate:
'   ・拡張子の位置(最後のドット)を探さず先頭から探す実装は、ファイル名に
'     ドットを含む(a)で壊れる。
'   ・"_旧版"を挿む位置を拡張子の【後ろ】にする実装は(a)(b)が
'     "MyBookshelf.xlsm_旧版"のような形になって落ちる。
'   ・拡張子が無い/空文字のガードを忘れると(c)(d)が例外や誤った値を返す。
Private Sub TestPreviousBookName41()
    ' (a) 標準名。
    ChkStr41 "D1_標準名", modMigrateFrom.PreviousBookName("MyBookshelf.xlsm"), _
        "MyBookshelf_旧版.xlsm"

    ' (b) dev バリアント(拡張子の直前に挿む。stem側にアンダースコアが
    '     あっても拡張子の位置だけで判定する)。
    ChkStr41 "D1_devバリアント", modMigrateFrom.PreviousBookName("MyBookshelf_dev.xlsm"), _
        "MyBookshelf_dev_旧版.xlsm"

    ' (c) 拡張子が無い → 空文字(版上げ対象として扱わない)。
    ChkStr41 "D1_拡張子なしは空", modMigrateFrom.PreviousBookName("MyBookshelf"), ""

    ' (d) 空文字 → 空文字。
    ChkStr41 "D1_空文字は空", modMigrateFrom.PreviousBookName(""), ""
End Sub

' ---- D2: StampText の書式 -----------------------------------------------------
'   discriminate:
'   ・Format$(…, "yyyy-mm-dd hh:nn")のまま(LO死角の再発)にすると、LOでは
'     "n"が解釈されず崩れた文字列になる=このテストはExcel/LO両方で
'     同じ文字列になることを固定する(実行系に依存しない組み立てかの検査)。
'   ・1桁の月/日/時/分をゼロ埋めし忘れると(b)が桁抜けの文字列になる。
Private Sub TestStampText41()
    ' (a) 2桁の月日時分。
    Dim stamp1 As Double: stamp1 = CDbl(DateSerial(2026, 9, 5) + TimeSerial(14, 30, 0))
    ChkStr41 "D2_2桁の月日時分", modMigrateFrom.StampText(stamp1), "2026-09-05 14:30"

    ' (b) 1桁の月/日/時/分がすべてゼロ埋めされる。
    Dim stamp2 As Double: stamp2 = CDbl(DateSerial(2026, 1, 2) + TimeSerial(3, 4, 0))
    ChkStr41 "D2_1桁の月日時分をゼロ埋め", modMigrateFrom.StampText(stamp2), "2026-01-02 03:04"
End Sub

' ---- D3: OfferText の文面 -----------------------------------------------------
'   discriminate:
'   ・ファイル名を文面に差し込み忘れる実装は(a)が落ちる。
'   ・件数(count)を差し込み忘れる実装は(b)が落ちる。
'   ・StampTextの書式を崩す実装は(c)が落ちる(D2と独立に、OfferText側の
'     結線=呼び忘れも検出する)。
'   ・確認の問いそのものを削る/文言を変える実装は(d)が落ちる。
Private Sub TestOfferText41()
    Dim stamp As Double: stamp = CDbl(DateSerial(2026, 9, 5) + TimeSerial(14, 30, 0))
    Dim msg As String: msg = modMigrateFrom.OfferText("MyBookshelf_旧版.xlsm", 7, stamp)

    ' (a) ファイル名が載る。
    ChkBool41 "D3_ファイル名が載る", (InStr(1, msg, "MyBookshelf_旧版.xlsm", vbBinaryCompare) > 0), True

    ' (b) 件数が載る。
    ChkBool41 "D3_件数が載る", (InStr(1, msg, "7件", vbBinaryCompare) > 0), True

    ' (c) 更新日時が "yyyy-mm-dd hh:nn" 形式で載る。
    ChkBool41 "D3_更新日時が載る", (InStr(1, msg, "2026-09-05 14:30", vbBinaryCompare) > 0), True

    ' (d) Yes/No で聞く確認文そのものが載る。
    ChkBool41 "D3_確認の問いが載る", _
        (InStr(1, msg, "この版へ本棚・実績・設定を引き継ぎますか?", vbBinaryCompare) > 0), True

    ' (e) 空のファイル名・0件でも落ちない(空文字を返さない=文面の骨格は残る)。
    Dim emptMsg As String: emptMsg = modMigrateFrom.OfferText("", 0, stamp)
    ChkBool41 "D3_ファイル名が空でも問いは残る", _
        (InStr(1, emptMsg, "この版へ本棚・実績・設定を引き継ぎますか?", vbBinaryCompare) > 0), True
End Sub

' ---- 判定ヘルパー -----------------------------------------------------------
Private Sub ChkBool41(ByVal label As String, ByVal got As Boolean, ByVal want As Boolean)
    modTestRunner.Check "R36-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkStr41(ByVal label As String, ByVal got As String, ByVal want As String)
    modTestRunner.Check "R36-" & label, (StrComp(got, want, vbBinaryCompare) = 0), _
        "実際=[" & got & "] 期待=[" & want & "]"
End Sub

Public Sub RunAll41()
    On Error GoTo H01Fail41
    TestPreviousBookName41
H02Next41:
    On Error GoTo H02Fail41
    TestStampText41
H03Next41:
    On Error GoTo H03Fail41
    TestOfferText41
H01Done41:
    On Error GoTo 0
    Exit Sub

H01Fail41:
    modTestRunner.Check "TestPreviousBookName41(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H02Next41
H02Fail41:
    modTestRunner.Check "TestStampText41(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H03Next41
H03Fail41:
    modTestRunner.Check "TestOfferText41(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done41
End Sub
