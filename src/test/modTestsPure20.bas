Attribute VB_Name = "modTestsPure20"
Option Explicit

' ============================================================================
' modTestsPure20 - R20-4(実機第7報③④「設定の3ステップ化」)+
'   R20-7(⑥「MS&AD配色刷新」)+ 追加D-1(ダッシュのセンタリング)の純ロジック
'   回帰。入口は modTestsPure19.RunAll19 の末尾から呼ばれる RunAll20 の1本。
' ----------------------------------------------------------------------------
' ここで固定するもの:
'   ・modSetupWizard.DeptFromSelector : 部門ピッカーの数値→部門名の境界
'   ・modChannel.ActiveLabelText      : ヘッダー部門ラベルの文言分岐
'   ・modDashStat.CenterX0            : 帯内中央寄せのX原点計算(追加D-1)
'   ・modSkin.ResolveColor/EffectiveSkin/CurrentTheme: 既定テーマの色解決
' ============================================================================

' ----------------------------------------------------------------------------
' R20-4b: 部門ピッカーの入力検証境界(modSetupWizard.DeptFromSelector)
' ----------------------------------------------------------------------------
Private Sub TestDeptFromSelector()
    modTestRunner.Check "部門選択_1は商品部", _
        (modSetupWizard.DeptFromSelector("1") = "商品部")
    modTestRunner.Check "部門選択_2はリスコン部", _
        (modSetupWizard.DeptFromSelector("2") = "リスコン部")
    ' "0"(自由入力)はここでは空文字(呼び出し側PromptDeptが個別に扱う)。
    modTestRunner.Check "部門選択_0は自由入力へ委譲(空文字)", _
        (modSetupWizard.DeptFromSelector("0") = "")
    modTestRunner.Check "部門選択_未知の番号は空文字", _
        (modSetupWizard.DeptFromSelector("3") = "")
    modTestRunner.Check "部門選択_空文字は空文字", _
        (modSetupWizard.DeptFromSelector("") = "")
    modTestRunner.Check "部門選択_前後空白は無視して1と判定", _
        (modSetupWizard.DeptFromSelector("  1 ") = "商品部")
    modTestRunner.Check "部門選択_全角の1もどきは無効(半角のみ許容)", _
        (modSetupWizard.DeptFromSelector(ChrW(&HFF11)) = "")
    modTestRunner.Check "部門選択_数字以外の文字列は空文字", _
        (modSetupWizard.DeptFromSelector("商品部") = "")
End Sub

' ----------------------------------------------------------------------------
' R20-4c: チャット上部の部門ラベル文言分岐(modChannel.ActiveLabelText)
' ----------------------------------------------------------------------------
Private Sub TestActiveLabelText()
    ' 2026-08-06 R20H FA-6: 「クリックで設定」はハンドラ未配線のラベルに
    ' 対する実行不能な案内だったため撤去し、実際に辿れる導線
    ' (❓ヘルプの⚙から設定)へ差し替えた。旧文言が出ないことも併せて固定する。
    modTestRunner.Check "部門ラベル_0件は未設定(ヘルプの設定導線を案内)", _
        (InStr(modChannel.ActiveLabelText(0, 5), "未設定") > 0 And _
         InStr(modChannel.ActiveLabelText(0, 5), "ヘルプの") > 0 And _
         InStr(modChannel.ActiveLabelText(0, 5), "クリックで設定") = 0), _
        "実際=" & modChannel.ActiveLabelText(0, 5)
    modTestRunner.Check "部門ラベル_負数(異常値)も未設定扱い(境界)", _
        (InStr(modChannel.ActiveLabelText(-1, 5), "未設定") > 0)
    modTestRunner.Check "部門ラベル_totalも0なら未設定(退化入力)", _
        (InStr(modChannel.ActiveLabelText(0, 0), "未設定") > 0)
    modTestRunner.Check "部門ラベル_全件一致は「全N部門」", _
        (modChannel.ActiveLabelText(5, 5) = ChrW(&HD83D) & ChrW(&HDCDA) & " 全5部門")
    modTestRunner.Check "部門ラベル_部分一致はhave/total部門", _
        (modChannel.ActiveLabelText(2, 5) = ChrW(&HD83D) & ChrW(&HDCDA) & " 2/5部門")
    modTestRunner.Check "部門ラベル_have=1ちょうど(境界)は未設定にならない", _
        (InStr(modChannel.ActiveLabelText(1, 5), "未設定") = 0)
End Sub

' ----------------------------------------------------------------------------
' 追加D-1: ダッシュのカード群センタリング(modDashStat.CenterX0)
' ----------------------------------------------------------------------------
Private Sub TestCenterX0()
    ' 帯590pt・版面550pt(=KPI_CARD_W*4+GAP*3の最小版面): 余白40pt→左右20ptずつ
    ' =下限20ptとちょうど一致(境界)。
    modTestRunner.Check "センタリング_最小版面の帯590ptは下限20(境界)", _
        (modDashStat.CenterX0(590, 550) = 20), _
        "実際=" & modDashStat.CenterX0(590, 550)
    ' 帯1288pt(窓1300)・版面920pt(220*4+10*3=920): 余白368pt→左右184pt。
    modTestRunner.Check "センタリング_広い帯は左右均等(1288,920→184)", _
        (Abs(modDashStat.CenterX0(1288, 920) - 184) < 0.001), _
        "実際=" & modDashStat.CenterX0(1288, 920)
    ' 帯が版面よりちょうど40pt広い(下限ぴったり)境界。
    modTestRunner.Check "センタリング_帯=版面+40は下限20(境界)", _
        (modDashStat.CenterX0(920 + 40, 920) = 20)
    ' 帯=版面ちょうど(余白0)は下限20(左寄せ・従来どおり)。
    modTestRunner.Check "センタリング_帯=版面ちょうどは下限20", _
        (modDashStat.CenterX0(920, 920) = 20)
    ' 帯<版面(狭い窓)は下限20のまま(はみ出す側は左寄せで従来どおり)。
    modTestRunner.Check "センタリング_帯<版面は下限20(狭い窓)", _
        (modDashStat.CenterX0(500, 920) = 20)
    ' minX0を指定した場合はそちらが下限になる。
    modTestRunner.Check "センタリング_minX0指定はその値が下限", _
        (modDashStat.CenterX0(500, 920, 40) = 40)
    ' 退化入力(負の帯幅)でも下限を割らない。
    modTestRunner.Check "センタリング_負の帯幅でも下限20(退化入力)", _
        (modDashStat.CenterX0(-100, 920) = 20)
End Sub

' ----------------------------------------------------------------------------
' R20-7a: 既定テーマ(msad)の色解決(modSkin.ResolveColor/EffectiveSkin)
' ----------------------------------------------------------------------------
Private Sub TestMsadThemeResolve()
    ' 空文字・未知テーマ名は既定の"msad"へ落ちる(2026-08-06: 既定が
    ' "light"から変わった)。
    modTestRunner.Check "既定テーマ_空文字はmsad", _
        (modSkin.EffectiveSkin("") = "msad")
    modTestRunner.Check "既定テーマ_未知の名称もmsad", _
        (modSkin.EffectiveSkin("blueold") = "msad")
    modTestRunner.Check "既定テーマ_大文字MSADも正規化されmsad", _
        (modSkin.EffectiveSkin("MSAD") = "msad")
    ' "light"は着せ替えとして明示的に残る(壊さない)。
    modTestRunner.Check "既定テーマ_lightは着せ替えとして残る", _
        (modSkin.EffectiveSkin("light") = "light")
    modTestRunner.Check "既定テーマ_darkはそのまま", _
        (modSkin.EffectiveSkin("dark") = "dark")

    ' 確定パレット(仕様書): primary=#01675B accent=#07A963。
    modTestRunner.Check "MS&ADパレット_primaryは#01675B", _
        (modSkin.ResolveColor("primary", "msad") = RGB(1, 103, 91))
    modTestRunner.Check "MS&ADパレット_accentは#07A963", _
        (modSkin.ResolveColor("accent", "msad") = RGB(7, 169, 99))
    ' 自分バブル上段=PRIMARY_LIGHT(#0B7D6E)・白字固定。
    modTestRunner.Check "MS&ADパレット_userBubbleは#0B7D6E(グラデ上段)", _
        (modSkin.ResolveColor("userBubble", "msad") = RGB(11, 125, 110))
    modTestRunner.Check "MS&ADパレット_userBubbleTextは白固定", _
        (modSkin.ResolveColor("userBubbleText", "msad") = RGB(255, 255, 255))
    ' 相手バブルの文字色は他テーマと同じ(自分バブルだけが白字になる)。
    modTestRunner.Check "MS&ADパレット_lightのuserBubbleTextはtextと同色", _
        (modSkin.ResolveColor("userBubbleText", "light") = modSkin.ResolveColor("text", "light"))
    ' ダーク側もPRIMARY_LIGHTを基調に(白字4.5:1以上・機械計算済み)。
    modTestRunner.Check "ダークテーマ_primaryはPRIMARY_LIGHT(#0B7D6E)基調", _
        (modSkin.ResolveColor("primary", "dark") = RGB(11, 125, 110))
End Sub

Public Sub RunAll20()
    On Error GoTo DeptFail20
    TestDeptFromSelector
NextLabel20:
    On Error GoTo LabelFail20
    TestActiveLabelText
NextCenter20:
    On Error GoTo CenterFail20
    TestCenterX0
NextTheme20:
    On Error GoTo ThemeFail20
    TestMsadThemeResolve
NextR20H20:
    On Error GoTo R20HFail20
    ' R20H(レビュー裁定Fix波)の回帰は modTestsPure21 へ(波ごとの分割)。
    modTestsPure21.RunAll21
NextDone20:
    On Error GoTo 0
    Exit Sub

DeptFail20:
    modTestRunner.Check "TestDeptFromSelector(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextLabel20
LabelFail20:
    modTestRunner.Check "TestActiveLabelText(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextCenter20
CenterFail20:
    modTestRunner.Check "TestCenterX0(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextTheme20
ThemeFail20:
    modTestRunner.Check "TestMsadThemeResolve(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextR20H20
R20HFail20:
    modTestRunner.Check "modTestsPure21.RunAll21(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone20
End Sub
