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

' ----------------------------------------------------------------------------
' (C) modAppState.AskGeneral 履歴保存のサニタイズ(R29H F3)。
'   AskGeneral自体はLLM呼び出しを含み純関数ではないため、保存直前に使う
'   合成式(Replace(本文,";;;"," ") & TrimPairs)そのものを、AskGeneralの
'   :318-319と同型でここに再現して固定する。回答本文へ区切り";;;"が
'   そのまま混ざっても、保存後の質問/回答の件数がずれない(往復ずれなし)
'   ことを見る。
' ----------------------------------------------------------------------------
Private Sub TestGenHistorySaveSanitize29()
    Dim prevU As String, prevA As String
    prevU = "": prevA = ""

    ' ターン1: 回答本文に区切り";;;"がそのまま混入したケース(実機で起こり得る)。
    Dim q1 As String: q1 = "質問その1"
    Dim a1 As String: a1 = "回答前半;;;回答後半"
    prevU = modAppState.TrimPairs(Replace(q1, ";;;", " ") & IIf(LenB(prevU) > 0, ";;;" & prevU, ""), 3)
    prevA = modAppState.TrimPairs(Replace(a1, ";;;", " ") & IIf(LenB(prevA) > 0, ";;;" & prevA, ""), 3)

    ' ターン2: 通常の質問/回答。
    Dim q2 As String: q2 = "質問その2"
    Dim a2 As String: a2 = "回答その2"
    prevU = modAppState.TrimPairs(Replace(q2, ";;;", " ") & IIf(LenB(prevU) > 0, ";;;" & prevU, ""), 3)
    prevA = modAppState.TrimPairs(Replace(a2, ";;;", " ") & IIf(LenB(prevA) > 0, ";;;" & prevA, ""), 3)

    ' サニタイズ後は件数が一致するはず(2ターン=2ペア)。
    Dim uParts() As String: uParts = Split(prevU, ";;;")
    Dim aParts() As String: aParts = Split(prevA, ";;;")
    modTestRunner.Check "R29H-F3_保存後の件数一致(往復ずれなし)", _
        (UBound(uParts) = UBound(aParts)), "u=" & (UBound(uParts) + 1) & " a=" & (UBound(aParts) + 1)

    ' ターン1の回答は";;;"がスペースへ退避され、区切りとして誤認されない
    ' (誤認されていれば余分な要素に割れ、aParts(1)は"回答前半"だけになる)。
    modTestRunner.Check "R29H-F3_区切り混入は退避され1ペアのまま", _
        (aParts(1) = "回答前半 回答後半"), "aParts(1)=[" & aParts(1) & "]"

    ' GenHistoryBlockへ渡しても新しい順→古い順の並べ直しが破綻しない
    ' (ターン1のQ/Aが対応する「会話2」として出る)。
    Dim blk As String
    blk = modGenPipe.GenHistoryBlock(prevU, prevA)
    modTestRunner.Check "R29H-F3_GenHistoryBlockで往復対応が保たれる", _
        (InStr(blk, "Q: 質問その1") > 0 And InStr(blk, "A: 回答前半 回答後半") > 0), _
        "blk=[" & blk & "]"
End Sub

' ============================================================================
Public Sub RunAll29()
    On Error GoTo ToastFail29
    TestToastWaitMsFor29
NextDiversity29:
    On Error GoTo DiversityFail29
    TestDiversitySwapPick29
NextGenSave29:
    On Error GoTo GenSaveFail29
    TestGenHistorySaveSanitize29
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
    Resume NextGenSave29
GenSaveFail29:
    modTestRunner.Check "TestGenHistorySaveSanitize29(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone29
End Sub
