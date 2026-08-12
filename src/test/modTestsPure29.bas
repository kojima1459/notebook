Attribute VB_Name = "modTestsPure29"
Option Explicit

' ============================================================================
' modTestsPure29 - R29波2(実機第14報・層2一括+視認性+文書訂正)の純ロジック
'   回帰。modTestsPure28(残り僅少・26,776字/上限30,000字)には収まらない
'   ため新設した分割先(憲章§4-6。modTestsPure21/25/26/28と同型)。
'   入口は modTestsPure28.RunAll28 の末尾から呼ばれる RunAll29 の1本。
' ----------------------------------------------------------------------------
' ここで固定するもの:
'   (A) modChrome.ToastWaitMsFor — トースト表示時間(R29 W2-3)。
'       ToastHeightForと同じ幅推定(全角=pitchPt/半角=pitchPt/2)を全角換算
'       字数へ流用し、読了速度=全角15字/秒相当でms換算する。下限3,000ms・
'       上限9,000msのクランプと、全角/半角の重みの違い(半角2字=全角1字)を
'       固定する。ここが緩むと、短いトーストが一瞬で消える/長いトーストが
'       読み切る前に消えるのRC(実機第14報)が再発する。
'   (B) modSparse.DiversitySwapPick — 相対スコア下限(R29 W2-5)。
'       pool側の差し替え候補スコアが現hits最下位スコアのminRatioX100%未満
'       なら差し替えない(69%不発火/70%発火)境界、minRatioX100<=0の旧挙動
'       (無条件差し替え。Optional省略時も同じ)、hitLowestScore<=0のときの
'       0除算防御(実行時エラー11を起こさず、制限せず差し替える)を固定する。
' ============================================================================

' ----------------------------------------------------------------------------
' (A) ToastWaitMsFor - 下限/上限クランプと全角換算の算数。
' ----------------------------------------------------------------------------
Private Sub TestToastWaitMsFor29()
    ' 短い文言(6字)は算数上400msだが、下限3,000msの床に張り付く。
    Dim short1 As Long: short1 = modChrome.ToastWaitMsFor("保存しました")
    modTestRunner.Check "R29-W2-3_短文は下限3000msの床", (short1 = 3000), "ms=" & short1

    ' 空文字でも実行時エラーにならず下限に張り付く(空メッセージの防御)。
    Dim empty1 As Long: empty1 = modChrome.ToastWaitMsFor("")
    modTestRunner.Check "R29-W2-3_空文字も下限3000ms(例外にならない)", _
        (empty1 = 3000), "ms=" & empty1

    ' 全角300字(換算300字)は 300/15*1000=20000ms だが上限9,000msでクランプ。
    Dim longMsg As String: longMsg = String(300, ChrW(&H3042))   ' 'あ'×300
    Dim long1 As Long: long1 = modChrome.ToastWaitMsFor(longMsg)
    modTestRunner.Check "R29-W2-3_長文は上限9000msでクランプ", (long1 = 9000), "ms=" & long1

    ' 中間値(クランプに触れない値): 全角101字→101/15*1000=6733.33…→切上げ6734ms。
    Dim midMsg As String: midMsg = String(101, ChrW(&H3042))
    Dim mid1 As Long: mid1 = modChrome.ToastWaitMsFor(midMsg)
    modTestRunner.Check "R29-W2-3_中間値は算数どおり(全角101字→6734ms)", _
        (mid1 = 6734), "ms=" & mid1

    ' 半角は全角の半分の重み(ToastHeightForと同じ幅推定)。半角202字は
    ' 全角101字と同じ換算字数(101)になり、待ち時間も一致するはず。
    Dim halfMsg As String: halfMsg = String(202, "A")
    Dim half1 As Long: half1 = modChrome.ToastWaitMsFor(halfMsg)
    modTestRunner.Check "R29-W2-3_半角は全角の半分換算(半角202字も6734ms)", _
        (half1 = mid1), "half=" & half1 & " mid=" & mid1
End Sub

' ----------------------------------------------------------------------------
' (B) DiversitySwapPick - 相対スコア下限の4分岐(R29 W2-5)。
'   hitN=2件とも"資料A"に収束・poolは"資料A"(現1位相当)と"資料B"(差し替え
'   候補)の2件。hitLowestScore=100を「現hitsの最下位スコア」に見立てる。
' ----------------------------------------------------------------------------
Private Sub TestDiversitySwapPick29()
    Dim hs(1 To 2) As String
    hs(1) = "資料A": hs(2) = "資料A"

    Dim ps(1 To 2) As String
    Dim sc(1 To 2) As Double
    ps(1) = "資料A": sc(1) = 90

    ' 69%(=69)は「未満」に該当し不発火(pick=0=既存の無介入と同じ形)。
    ps(2) = "資料B": sc(2) = 69
    Dim pick69 As Long
    pick69 = modSparse.DiversitySwapPick(hs, 2, ps, sc, 2, 100, 70)
    modTestRunner.Check "R29-W2-5_69%は不発火(下限未満)", (pick69 = 0), "pick=" & pick69

    ' ちょうど70%は「未満」に該当しないので発火(pick=2=poolの資料B)。
    ps(2) = "資料B": sc(2) = 70
    Dim pick70 As Long
    pick70 = modSparse.DiversitySwapPick(hs, 2, ps, sc, 2, 100, 70)
    modTestRunner.Check "R29-W2-5_70%は発火(下限以上)", (pick70 = 2), "pick=" & pick70

    ' minRatioX100<=0は旧挙動(スコアが1しかなくても無条件で差し替える)。
    ps(2) = "資料B": sc(2) = 1
    Dim pickOld As Long
    pickOld = modSparse.DiversitySwapPick(hs, 2, ps, sc, 2, 100, 0)
    modTestRunner.Check "R29-W2-5_minRatio0は旧挙動(無条件差し替え)", _
        (pickOld = 2), "pick=" & pickOld

    ' Optional省略(呼び出し元を増やさない後方互換)も同じ旧挙動になる。
    Dim pickDefault As Long
    pickDefault = modSparse.DiversitySwapPick(hs, 2, ps, sc, 2)
    modTestRunner.Check "R29-W2-5_新引数省略は旧挙動(後方互換)", _
        (pickDefault = 2), "pick=" & pickDefault

    ' hitLowestScore<=0は比較基準(分母)が無いため0除算せず制限しない
    ' (実行時エラー11を起こさず、pickはそのまま返る)。
    ps(2) = "資料B": sc(2) = 5
    Dim pickZeroDenom As Long
    pickZeroDenom = modSparse.DiversitySwapPick(hs, 2, ps, sc, 2, 0, 70)
    modTestRunner.Check "R29-W2-5_hits最下位0は除算防御(制限せず差し替え)", _
        (pickZeroDenom = 2), "pick=" & pickZeroDenom
End Sub

' ============================================================================
Public Sub RunAll29()
    On Error GoTo ToastFail29
    TestToastWaitMsFor29
NextDiversity29:
    On Error GoTo DiversityFail29
    TestDiversitySwapPick29
NextDone29:
    On Error GoTo 0
    Exit Sub

ToastFail29:
    modTestRunner.Check "TestToastWaitMsFor29(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDiversity29
DiversityFail29:
    modTestRunner.Check "TestDiversitySwapPick29(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone29
End Sub
