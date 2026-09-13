Attribute VB_Name = "modTestsPure21"
Option Explicit

' ============================================================================
' modTestsPure21 - R20H(レビュー裁定Fix波)の純ロジック回帰テスト。
'   modTestsPure20(7,311字)にこれを足すと同モジュールの主題(R20-4/R20-7/
'   追加D-1)からも外れるため、Fix波専用の分割先として新設した。入口は
'   modTestsPure20.RunAll20 の末尾から呼ばれる RunAll21 の1本。
'   末尾で R21-1(余白の構造完治)の modTestsPure22.RunAll22 へ繋ぐ。
' ----------------------------------------------------------------------------
' ここで固定するもの(FB-4「形だけ禁止」の指定3点):
'   ・FA-2: modViewport.Busy3           : tick経路がRefitActionへ渡す
'           isBusyの組み立て(modUiLock/modShelf/modShelfSyncの3情報源OR)。
'   ・FA-4: modUIShelf.ClearAreaLastRow : table/gallery切替をまたいだ
'           旧行の取りこぼし再現(table60冊→gallery→table絞り込み)。
'   ・FA-9: modSetupWizard.WizardShouldRun: ウィザード専用フラグの1回性。
' ============================================================================

' ----------------------------------------------------------------------------
' FA-2: tick経路のbusy判定に3情報源が正しくORされているか(modViewport.Busy3)。
'   RefitAction自体は無変更(isBusyを1個のBooleanで受けるだけ)なので、
'   ここで固定するのは「呼び出し側がmodUiLock単独に戻ってしまう」退行を
'   防ぐ、3引数OR組み立ての契約そのもの。
' ----------------------------------------------------------------------------
Private Sub TestBusy3()
    modTestRunner.Check "Busy3_全部False→False", _
        (modViewport.Busy3(False, False, False) = False)
    modTestRunner.Check "Busy3_lockのみTrue(旧実装相当)→True", _
        (modViewport.Busy3(True, False, False) = True)
    ' R20H FA-2の本丸: modUiLockはFalse(誰も明示ロックしていない)でも、
    ' 取込(modShelf)や同期(modShelfSync)のDoEvents中はTrueであるべき
    ' (旧実装はmodUiLock.IsBusy()単独だったため、ここがFalseのまま
    ' 再フィットが走ってしまっていた)。
    modTestRunner.Check "Busy3_shelfのみTrue(取込中・旧実装の穴)→True", _
        (modViewport.Busy3(False, True, False) = True)
    modTestRunner.Check "Busy3_shelfSyncのみTrue(同期中・旧実装の穴)→True", _
        (modViewport.Busy3(False, False, True) = True)
    modTestRunner.Check "Busy3_全部True→True", _
        (modViewport.Busy3(True, True, True) = True)
End Sub

' ----------------------------------------------------------------------------
' FA-4: table/gallery切替をまたいだ旧行の取りこぼし再現
'   (modUIShelf.ClearAreaLastRow)。
' ----------------------------------------------------------------------------
Private Sub TestClearAreaLastRow()
    ' 再現シナリオ: table60冊→gallery→table絞り込み。
    '   ①table60冊: FIRST_CARD_ROW=13ぶん+60冊で最終行72。この時点では
    '     共有高水位(modKnowledge.ShelfRowHigh)もtable自身の記憶も72で揃う。
    modTestRunner.Check "旧行再現_table60冊描画直後は共有もtable記憶も72で一致", _
        (modUIShelf.ClearAreaLastRow(72, 72) = 72)
    '   ②gallery切替: galleryが20冊ぶんで描き、共有高水位を33へ縮める。
    '     table自身の記憶(mTableLastRow)はgalleryの描画では一切更新されない
    '     ため72のまま残る。
    '   ③table絞り込み(5冊)へ戻る直前のClearCardAreaは、更新前の
    '     mTableLastRow=72をまだ持っている。旧実装(共有高水位だけを見る)
    '     ならここで33までしかクリアされず、34〜72行に旧table内容が
    '     取りこぼされて残る(実機で再現した不具合そのもの)。
    modTestRunner.Check "旧行再現_gallery経由で共有33に縮んでもtable記憶72を優先", _
        (modUIShelf.ClearAreaLastRow(33, 72) = 72)
    ' 逆方向(tableの記憶の方が小さい)は共有側を優先(Maxの対称性確認)。
    modTestRunner.Check "旧行再現_table記憶より共有高水位が大きければそちらを優先", _
        (modUIShelf.ClearAreaLastRow(80, 20) = 80)
    ' 上限クランプ: 本文最終行の上限412(先頭行13+最大400行-1)を超えない。
    modTestRunner.Check "旧行再現_上限412でクランプ", _
        (modUIShelf.ClearAreaLastRow(9999, 9999) = 412)
    ' 下限クランプ: 退化入力(0/負)でも本文の先頭行13を割らない。
    modTestRunner.Check "旧行再現_下限13でクランプ(退化入力)", _
        (modUIShelf.ClearAreaLastRow(0, -5) = 13)
End Sub

' ----------------------------------------------------------------------------
' FA-9: ウィザード専用フラグ(nexus_wizard_done)の1回性
'   (modSetupWizard.WizardShouldRun)。
' ----------------------------------------------------------------------------
Private Sub TestWizardShouldRun()
    modTestRunner.Check "ウィザード1回性_未設定(空文字)なら出す", _
        (modSetupWizard.WizardShouldRun("") = True)
    modTestRunner.Check "ウィザード1回性_フラグ1なら二度と出さない", _
        (modSetupWizard.WizardShouldRun("1") = False)
    ' 壊れた/未知の値は「出す」側に倒す(読み取り失敗で再出現を防ぐより、
    ' 誤って握り潰して二度と出せなくなる方を避ける設計)。
    modTestRunner.Check "ウィザード1回性_壊れた値(1以外)は出す側に倒す", _
        (modSetupWizard.WizardShouldRun("0") = True)
    modTestRunner.Check "ウィザード1回性_大文字/全角の1もどきは一致しない(出す)", _
        (modSetupWizard.WizardShouldRun(ChrW(&HFF11)) = True)
End Sub

Public Sub RunAll21()
    On Error GoTo Busy3Fail21
    TestBusy3
NextClear21:
    On Error GoTo ClearFail21
    TestClearAreaLastRow
NextWizard21:
    On Error GoTo WizardFail21
    TestWizardShouldRun
NextRun22:
    On Error GoTo Run22Fail21
    modTestsPure22.RunAll22
NextDone21:
    On Error GoTo 0
    Exit Sub

Busy3Fail21:
    modTestRunner.Check "TestBusy3(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextClear21
ClearFail21:
    modTestRunner.Check "TestClearAreaLastRow(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextWizard21
WizardFail21:
    modTestRunner.Check "TestWizardShouldRun(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRun22
Run22Fail21:
    modTestRunner.Check "modTestsPure22.RunAll22(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone21
End Sub
