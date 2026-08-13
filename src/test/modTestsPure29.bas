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
'   (B) modSparse.DiversitySwapPick — 相対スコア下限(R29 W2-5/R29H F4)。
'       pool側の差し替え候補スコアが現hits最下位スコアのminRatioX100%未満
'       なら差し替えない(69%不発火/70%発火)境界、minRatioX100<=0の旧挙動
'       (無条件差し替え。Optional省略時も同じ)、hitLowestScore<=0(0/負値)
'       のときは分母が無い異常値として安全側=差し替え不発火(0除算は
'       起こさない)を固定する。
'   (C) modAppState.AskGeneral 履歴保存のサニタイズ(R29H F3)。
'       保存直前に本文の区切り";;;"を半角スペースへ退避する合成式
'       (Replace+TrimPairs)を、区切り混入→保存→GenHistoryBlock往復の
'       形で固定する。
'   (D) modViewport2.ReleaseRange — 解放する行範囲の計算(R30 W1・実機第15報)。
'       行高18ptの明示設定が行をExcelの内部使用範囲へ焼き付け、FreezePanes
'       併用のホイールがそこまで転がれるのが下余白の真因だった。唯一の即時
'       解放手段 Rows.Delete の範囲計算を、固定領域(行1〜4)の防衛・冪等条件・
'       上限を超えた使用済みの3方向で固定する。
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

    ' hitLowestScore<=0は比較基準(分母)が無い異常値。R29H F4で旧挙動
    ' (制限せず無条件差し替え)から安全側=差し替え不発火(pick=0)へ反転した
    ' (0除算は起こさない)。
    ps(2) = "資料B": sc(2) = 5
    Dim pickZeroDenom As Long
    pickZeroDenom = modSparse.DiversitySwapPick(hs, 2, ps, sc, 2, 0, 70)
    modTestRunner.Check "R29H-F4_hits最下位0は安全側で不発火", _
        (pickZeroDenom = 0), "pick=" & pickZeroDenom

    ' 負値(hitLowestScore=-0.1)も同様に安全側で不発火(実行時エラー11を
    ' 起こさない)。
    ps(2) = "資料B": sc(2) = 5
    Dim pickNegDenom As Long
    pickNegDenom = modSparse.DiversitySwapPick(hs, 2, ps, sc, 2, -0.1, 70)
    modTestRunner.Check "R29H-F4_hits最下位が負値でも不発火", _
        (pickNegDenom = 0), "pick=" & pickNegDenom
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

' ----------------------------------------------------------------------------
' (D) modChrome.UnlockedTeaserSuffix - 未解放テーマ案内の付記(R29H F2b-2)。
'   CycleSkinのループ内ティザートースト(waitless連発)を廃止し、切替成功
'   トースト末尾へ1回だけ付記する形に統合した実体。tcの3区間を固定する。
' ----------------------------------------------------------------------------
Private Sub TestUnlockedTeaserSuffix29()
    ' tc=0: サクラ/海(5件)・金(20件)とも未解放。
    Dim s0 As String: s0 = modChrome.UnlockedTeaserSuffix(0)
    modTestRunner.Check "R29H-F2b-2_tc0は両方とも未解放を列挙", _
        (InStr(s0, "さくら/海は感謝5件で解放") > 0 And InStr(s0, "金は感謝20件で解放") > 0), _
        "s0=[" & s0 & "]"

    ' tc=4: 境界未満はまだ両方未解放(5件の壁に届いていない)。
    Dim s4 As String: s4 = modChrome.UnlockedTeaserSuffix(4)
    modTestRunner.Check "R29H-F2b-2_tc4はまだ両方未解放(5件未満)", _
        (InStr(s4, "さくら/海") > 0), "s4=[" & s4 & "]"

    ' tc=5: サクラ/海は解放済み、金だけ残る。
    Dim s5 As String: s5 = modChrome.UnlockedTeaserSuffix(5)
    modTestRunner.Check "R29H-F2b-2_tc5はさくら/海が解放済み(金のみ残)", _
        (InStr(s5, "さくら") = 0 And InStr(s5, "金は感謝20件で解放") > 0), "s5=[" & s5 & "]"

    ' tc=19: 金はまだ未解放(20件の壁に届いていない)。
    Dim s19 As String: s19 = modChrome.UnlockedTeaserSuffix(19)
    modTestRunner.Check "R29H-F2b-2_tc19は金がまだ未解放(20件未満)", _
        (InStr(s19, "金") > 0), "s19=[" & s19 & "]"

    ' tc=20: 全解放。付記なしの空文字(トーストに余計な文言を足さない)。
    Dim s20 As String: s20 = modChrome.UnlockedTeaserSuffix(20)
    modTestRunner.Check "R29H-F2b-2_tc20は全解放で空文字(ネガティブ確認込み)", _
        (LenB(s20) = 0), "s20=[" & s20 & "]"
End Sub

' ----------------------------------------------------------------------------
' (D) modViewport2.ReleaseRange - 解放する行範囲の計算(R30 W1-2/W1-4)。
' ----------------------------------------------------------------------------
'   余白の真因は「行高18ptの明示設定が行をExcelの内部使用範囲へ焼き付け、
'   FreezePanes併用のホイールがそこまで転がれる」ことだった。唯一の即時解放
'   手段が Rows.Delete なので、その範囲計算が緩むと (a) 固定領域(行1〜4)まで
'   消して画面が壊れる (b) 消し足りずに余白が残る (c) 消すものが無いのに毎回
'   削除して描画が重くなる、のいずれかが起きる。境界を1件ずつ固定する。
Private Sub TestReleaseRange30()
    Dim fromRow As Long, toRow As Long
    Dim hit As Boolean

    ' 旧版の焼き付け(1〜400行)からの移行掃除。バンド下端60+余裕3=63を残し、
    ' 64行目から NEXUS_MAX_ROW(2000)まで消す ―― 使用済みが400までしか無くても
    ' 2000まで消すのは、旧 ExtendChatBand が2000行まで塗り得たため。
    hit = modViewport2.ReleaseRange(63, 400, 5, 2000, fromRow, toRow)
    modTestRunner.Check "R30-W1_移行掃除は64:2000を削除", _
        (hit = True) And (fromRow = 64) And (toRow = 2000), _
        "hit=" & hit & " from=" & fromRow & " to=" & toRow

    ' 冪等: 掃除済み(使用済み下端=残す下端)なら削除しない。ここが True に
    ' なると、バブル1個ごとに1900行の Rows.Delete が走って描画が固まる。
    hit = modViewport2.ReleaseRange(63, 63, 5, 2000, fromRow, toRow)
    modTestRunner.Check "R30-W1_使用済み=下端なら削除しない(冪等)", _
        (hit = False), "hit=" & hit

    ' 使用済みが残す下端より上(バンドが窓より広い)ときも削除しない。
    hit = modViewport2.ReleaseRange(63, 20, 5, 2000, fromRow, toRow)
    modTestRunner.Check "R30-W1_使用済みが下端より上なら削除しない", _
        (hit = False), "hit=" & hit

    ' 1行だけはみ出した=削除する(冪等条件の境界。63/64の1行差)。
    hit = modViewport2.ReleaseRange(63, 64, 5, 2000, fromRow, toRow)
    modTestRunner.Check "R30-W1_1行超過でも削除する(境界)", _
        (hit = True) And (fromRow = 64), _
        "hit=" & hit & " from=" & fromRow

    ' 使用済みが上限を超えている(旧版の塗りが2000行より下まで届いた)場合は
    ' そこまで消す ―― 上限で頭打ちにすると余白が残る。
    hit = modViewport2.ReleaseRange(63, 5000, 5, 2000, fromRow, toRow)
    modTestRunner.Check "R30-W1_使用済みが上限超なら使用済みまで消す", _
        (hit = True) And (toRow = 5000), "to=" & toRow

    ' 固定領域の防衛。boundRow が異常に小さくても行1〜4(ヘッダー/入力欄/
    ' ヒント)は絶対に消さない ―― 消すと FreezePanes ごと画面が壊れる。
    hit = modViewport2.ReleaseRange(0, 400, 5, 2000, fromRow, toRow)
    modTestRunner.Check "R30-W1_boundRow=0でも行1-4は守る(from=6)", _
        (hit = True) And (fromRow = 6), "from=" & fromRow

    ' 負値(状態が壊れた場合の防衛)も同じ下限へ丸める。
    hit = modViewport2.ReleaseRange(-100, 400, 5, 2000, fromRow, toRow)
    modTestRunner.Check "R30-W1_boundRow負値でも行1-4は守る(from=6)", _
        (hit = True) And (fromRow = 6), "from=" & fromRow

    ' 下限クランプは冪等条件にも効く: 使用済みが minRow 以下なら削除しない。
    hit = modViewport2.ReleaseRange(0, 5, 5, 2000, fromRow, toRow)
    modTestRunner.Check "R30-W1_使用済み=minRowなら削除しない(クランプ後の冪等)", _
        (hit = False), "hit=" & hit

    ' 上限クランプ(R30 F1): 旧世代ブックのUsedRangeが病的に大きい(例:
    ' Excel全行1048576)場合、maxRow(2000)の4倍=8000で頭打ちする ―― さもなくば
    ' Rows("64:1048576").Delete が走り32bitで凍結し得る。
    hit = modViewport2.ReleaseRange(63, 1048576, 5, 2000, fromRow, toRow)
    modTestRunner.Check "R30-F1_使用済みが病的に大きくても8000で頭打ち", _
        (hit = True) And (toRow = 8000), "to=" & toRow
End Sub

' ----------------------------------------------------------------------------
' (E) modGateway.LooksLikeLimitError - <verdict>タグ救済(R30 W2-2・3件目)。
' ----------------------------------------------------------------------------
'   分解段(modPrompts.BuildDecomposePrompt)の <verdict>...</verdict> 応答が
'   「上限」等を含む論点(保険約款等)に触れただけでE0204(利用上限エラー)へ
'   誤爆する退行をR21 D2/R21H F4と同型で塞ぐ。117字級の短文decompose応答が
'   救済されること・本物の上限定型文は引き続きE0204のままであることの対。
Private Sub TestE0204DecomposeVerdictRescue30()
    Dim decomposeResp As String
    decomposeResp = "<verdict>single</verdict>" & vbLf & _
        "<reason>質問は保険契約における支払限度額の上限について尋ねる内容であり、" & _
        "対象範囲は単一の条文のみで十分に答えられるため、分解は不要と判断した。</reason>"
    modTestRunner.Check "R30-W2-2_decompose型verdict応答は救済(誤爆しない)", _
        (modGateway.LooksLikeLimitError(decomposeResp) = False), _
        "len=" & Len(decomposeResp)

    ' parts/clarify/global の他verdictでも同様に救済される(タグ判定は値を見ない)。
    modTestRunner.Check "R30-W2-2_verdict=clarifyでも救済", _
        (modGateway.LooksLikeLimitError("<verdict>clarify</verdict>上限の解釈が複数ありうる") = False)

    ' 退行検知(ネガティブ確認): <verdict>タグが無ければ、これまでどおり
    ' 本物の上限定型文は引き続きE0204のまま(タグ判定が全体を無条件で
    ' 救済してしまう退行を防ぐ)。
    modTestRunner.Check "R30-W2-2_タグ無しの本物の上限定型文は引き続きエラー扱い", _
        (modGateway.LooksLikeLimitError("申し訳ございません。本日の利用上限に達しました。") = True)
End Sub

' ----------------------------------------------------------------------------
' (F) modAskThorough.ExpandPromptWithSynonymHint - 表記揺れ道1(R30 W2-6)。
' ----------------------------------------------------------------------------
'   modPromptsは凍結のため不触。呼び出し側(modAskRetrieve)でBuildExpand
'   Promptの戻り値へ表記揺れヒントを連結する実体を固定する。連結される
'   こと・空でないこと・ネガティブ(元プロンプトが空文字でも例外にならない
'   こと)を見る。
Private Sub TestExpandPromptWithSynonymHint30()
    Dim basePr As String: basePr = "元のexpandプロンプト"
    Dim merged As String: merged = modAskThorough.ExpandPromptWithSynonymHint(basePr)

    modTestRunner.Check "R30-W2-6_連結後も元プロンプトを保持", _
        (Left$(merged, Len(basePr)) = basePr), "merged=[" & merged & "]"
    modTestRunner.Check "R30-W2-6_連結後に言い換え観点の文言を含む", _
        (InStr(merged, "言い換え") > 0), "merged=[" & merged & "]"
    modTestRunner.Check "R30-W2-6_連結後は元より長い(空の追記になっていない)", _
        (Len(merged) > Len(basePr)), "len=" & Len(merged)

    ' ネガティブ確認: 元プロンプトが空文字でも例外にならず、追記文だけが付く。
    Dim mergedEmpty As String: mergedEmpty = modAskThorough.ExpandPromptWithSynonymHint("")
    modTestRunner.Check "R30-W2-6_元が空文字でも例外にならず追記文が付く", _
        (LenB(mergedEmpty) > 0 And InStr(mergedEmpty, "言い換え") > 0), _
        "mergedEmpty=[" & mergedEmpty & "]"
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
NextTeaser29:
    On Error GoTo TeaserFail29
    TestUnlockedTeaserSuffix29
NextRelease30:
    On Error GoTo ReleaseFail30
    TestReleaseRange30
NextE0204D30:
    On Error GoTo E0204Fail30
    TestE0204DecomposeVerdictRescue30
NextSynHint30:
    On Error GoTo SynHintFail30
    TestExpandPromptWithSynonymHint30
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
    Resume NextTeaser29
TeaserFail29:
    modTestRunner.Check "TestUnlockedTeaserSuffix29(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRelease30
ReleaseFail30:
    modTestRunner.Check "TestReleaseRange30(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextE0204D30
E0204Fail30:
    modTestRunner.Check "TestE0204DecomposeVerdictRescue30(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextSynHint30
SynHintFail30:
    modTestRunner.Check "TestExpandPromptWithSynonymHint30(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone29
End Sub
