Attribute VB_Name = "modTestsPure24"
Option Explicit

' ============================================================================
' modTestsPure24 - R25-3(バッジ16枠化・実機第11報⑥)の純ロジック回帰。
'   入口は modTestsPure23.RunAll23 の末尾から呼ばれる RunAll24 の1本
'   (modTestsPure20〜23と同型の連鎖)。
' ----------------------------------------------------------------------------
' 背景(docs/dev/spec_20260810_R25_実機第11報.md R25-3):
'   ・FA-R25-3a: qa_share10(qa_shared_total)/gapfill(gapfill_total)は
'     加算箇所がリポジトリに存在せず獲得不可能だった。modInsightIo.
'     EmitVerifiedQA/EmitCorrection の書込み成功時にBumpを配線した
'     (加算そのものはWorksheetを書くためPureテストの対象外。ここでは
'     「バッジ表に載っていること」を固定する)。
'   ・FA-R25-3b: 新規3種(thanks5/streak30/thorough10)を追加し、
'     13+3=16(ダッシュ4列×4段がちょうど埋まる)にした。
'   modTestsPure2.RunBadgeCatalogTests が表の一般的な整合(長さ一致・
'   空要素なし・重複なし・旧4種の存在)を既に固定しているため、ここでは
'   「16枠化」という今回の変更そのもの(件数=16固定・新3種の存在・
'   閾値を表す数字が条件文に含まれること)に絞る。
' ============================================================================

Private Sub TestBadgeCatalogCount16()
    Dim ids() As String, titles() As String, shorts() As String, conds() As String
    Dim n As Long
    n = modStats.BadgeCatalog(ids, titles, shorts, conds)

    ' 恒真防止: 13(旧)+3(新規)=16をリテラルで固定する(将来ここが動いたら
    ' 「意図した増減か」をレビューで必ず立ち止まらせるための固定値)。
    modTestRunner.Check "R25-3_BadgeCatalog件数はちょうど16(ダッシュ4x4を埋める)", _
        (n = 16), "n=" & n
End Sub

Private Sub TestBadgeCatalogNewIdsExist()
    Dim ids() As String, titles() As String, shorts() As String, conds() As String
    Dim n As Long
    n = modStats.BadgeCatalog(ids, titles, shorts, conds)

    ' 新3種(FA-R25-3b)。
    modTestRunner.Check "R25-3_バッジ表_thanks5を含む", BadgeIdExists24(ids, "thanks5")
    modTestRunner.Check "R25-3_バッジ表_streak30を含む", BadgeIdExists24(ids, "streak30")
    modTestRunner.Check "R25-3_バッジ表_thorough10を含む", BadgeIdExists24(ids, "thorough10")

    ' 修理した死にバッジ2種(FA-R25-3a)。modTestsPure2側でも固定済みだが、
    ' 「16枠が旧13種+新3種で構成される」という今回の変更の前提そのものを
    ' このテストファイル単体でも確認できるよう重ねて固定する。
    modTestRunner.Check "R25-3_バッジ表_qa_share10を含む(修理対象)", BadgeIdExists24(ids, "qa_share10")
    modTestRunner.Check "R25-3_バッジ表_gapfillを含む(修理対象)", BadgeIdExists24(ids, "gapfill")
End Sub

Private Sub TestBadgeCatalogNoDuplicateIds()
    Dim ids() As String, titles() As String, shorts() As String, conds() As String
    Dim n As Long
    n = modStats.BadgeCatalog(ids, titles, shorts, conds)

    Dim i As Long, j As Long, dupN As Long
    For i = LBound(ids) To UBound(ids)
        For j = i + 1 To UBound(ids)
            If StrComp(ids(i), ids(j), vbTextCompare) = 0 Then dupN = dupN + 1
        Next j
    Next i
    modTestRunner.Check "R25-3_バッジ表_16種にidの重複が無い(新3種を足しても)", _
        (dupN = 0), "dup=" & dupN
End Sub

' 新3種の条件文(長文説明)に、閾値を表す数字が正しく入っていること
' (コピペ改変で閾値の数字だけ取り違える事故を防ぐ)。EvaluateBadges本体は
' GetStat経由でWorksheetに触れるためPureテストから直接は呼べない
' (modStats冒頭・modOutlineBuild等の既存注記と同型)。ここではBadgeCatalog
' (Worksheetに触れない純関数)側の文言が閾値と矛盾していないことだけを固定する。
Private Sub TestNewBadgeConditionsMentionThreshold()
    Dim ids() As String, titles() As String, shorts() As String, conds() As String
    Dim n As Long
    n = modStats.BadgeCatalog(ids, titles, shorts, conds)

    Dim idxThanks As Long, idxStreak As Long, idxThorough As Long
    idxThanks = IndexOf24(ids, "thanks5")
    idxStreak = IndexOf24(ids, "streak30")
    idxThorough = IndexOf24(ids, "thorough10")

    modTestRunner.Check "R25-3_境界_thanks5の条件文に閾値5が入っている", _
        (idxThanks >= 0), "idx=" & idxThanks
    If idxThanks >= 0 Then
        modTestRunner.Check "R25-3_境界_thanks5条件文=『感謝を5件受け取ると獲得』", _
            (InStr(conds(idxThanks), "5") > 0), "cond=" & conds(idxThanks)
    End If

    modTestRunner.Check "R25-3_境界_streak30の条件文に閾値30が入っている", _
        (idxStreak >= 0), "idx=" & idxStreak
    If idxStreak >= 0 Then
        modTestRunner.Check "R25-3_境界_streak30条件文に『30』が入っている(29/30の境界を表す数字)", _
            (InStr(conds(idxStreak), "30") > 0), "cond=" & conds(idxStreak)
    End If

    modTestRunner.Check "R25-3_境界_thorough10の条件文に閾値10が入っている", _
        (idxThorough >= 0), "idx=" & idxThorough
    If idxThorough >= 0 Then
        modTestRunner.Check "R25-3_境界_thorough10条件文に『10』が入っている(9/10の境界を表す数字)", _
            (InStr(conds(idxThorough), "10") > 0), "cond=" & conds(idxThorough)
    End If
End Sub

' F7(m-5): バッジ絵文字二重の是正。旧チェックは新3種の先頭だけを見ていたが、
' 本来固定すべきは「長い名前(titles)に絵文字を含まない」という表全体の
' 意匠統一(modHub.DrawBadges/modDash側が既に🏅/🔒を前置するため、titles
' 自身にも絵文字があると二重表示になる)。全16種を対象に、サロゲートペア
' (絵文字はUTF-16で上位D800-DBFF+下位DC00-DFFFの2コードから成る)の
' 有無で判定する。
Private Sub TestBadgeTitlesHaveNoSurrogate24()
    Dim ids() As String, titles() As String, shorts() As String, conds() As String
    Dim n As Long
    n = modStats.BadgeCatalog(ids, titles, shorts, conds)

    Dim i As Long, badN As Long
    For i = LBound(titles) To UBound(titles)
        If ContainsSurrogate24(titles(i)) Then badN = badN + 1
    Next i
    modTestRunner.Check "R25-3_バッジ表_長い名前16種に絵文字(サロゲート)を含まない", _
        (badN = 0), "badN=" & badN
End Sub

' ============================================================================
' R27 波1(実機第12報②「特約が検索から消える」の根治)の純ロジック回帰。
' ============================================================================
' F1-1: modSparse.CapKeyScore — キーワード加点の頭打ち。
'   実機ではKeyScoreが330まで伸び、SPARSE_WEIGHT 0.06 を掛けた加点20が
'   cos類似度(-1〜1)を押し流していた。cap=10(加点0.6)でcosと同じ土俵へ戻す。
'   cap未満/ちょうど/超過 の3点を固定する(cap+10 は「越えた分は切る」側)。
Private Sub TestCapKeyScoreBoundary24()
    Dim cap As Double: cap = 10#

    ' cap-1: 頭打ちの手前は素通し(値が変わってはいけない)。
    modTestRunner.Check "R27-F1-1_cap-1は素通し(9→9)", _
        (modSparse.CapKeyScore(cap - 1#, cap) = cap - 1#), _
        "実際=" & Format$(modSparse.CapKeyScore(cap - 1#, cap), "0.000")

    ' cap ちょうど: 境界は「切らない」側(> で比較しているため)。
    modTestRunner.Check "R27-F1-1_capちょうどは素通し(10→10)", _
        (modSparse.CapKeyScore(cap, cap) = cap), _
        "実際=" & Format$(modSparse.CapKeyScore(cap, cap), "0.000")

    ' cap+10: 超過分は必ず落ちる(ここが効かないとR27以前へ逆戻り)。
    modTestRunner.Check "R27-F1-1_cap+10は頭打ち(20→10)", _
        (modSparse.CapKeyScore(cap + 10#, cap) = cap), _
        "実際=" & Format$(modSparse.CapKeyScore(cap + 10#, cap), "0.000")

    ' 実機で観測された生スコア330も同じ1本の式で10へ落ちること。
    modTestRunner.Check "R27-F1-1_実機の暴走値330も10へ落ちる", _
        (modSparse.CapKeyScore(330#, cap) = cap), _
        "実際=" & Format$(modSparse.CapKeyScore(330#, cap), "0.000")

    ' cap<=0 は「頭打ちなし」= 旧挙動へ戻すエスケープハッチ。
    modTestRunner.Check "R27-F1-1_cap=0は頭打ちなし(330がそのまま返る)", _
        (modSparse.CapKeyScore(330#, 0#) = 330#), _
        "実際=" & Format$(modSparse.CapKeyScore(330#, 0#), "0.000")
End Sub

' F1-2: modSparse.KeyLenWeight — 語長重み Len^1.5 の Len を8でclamp。
'   7→8 は増え、8→9 以降は増えない。9字で 9^1.5=27 が返ったら clamp が
'   効いていない(=長い資料名キー1本で順位が決まる旧挙動)。
Private Sub TestKeyLenWeightClamp24()
    Dim w7 As Double, w8 As Double, w9 As Double, w16 As Double
    w7 = modSparse.KeyLenWeight(7)
    w8 = modSparse.KeyLenWeight(8)
    w9 = modSparse.KeyLenWeight(9)
    w16 = modSparse.KeyLenWeight(16)

    ' clampの手前(7字)は8字より必ず軽い=「長い語ほど重い」性質は残す。
    modTestRunner.Check "R27-F1-2_7字は8字より軽い(clamp手前は従来どおり)", _
        (w7 < w8), "w7=" & Format$(w7, "0.000") & " w8=" & Format$(w8, "0.000")

    ' 境界(8字)の値そのもの: 8^1.5 = 22.627…(恒真化を避けるため実値で固定)
    modTestRunner.Check "R27-F1-2_8字の重みは8^1.5=22.627", _
        (Abs(w8 - 22.6274169979695) < 0.000001), "w8=" & Format$(w8, "0.000000")

    ' clamp本体: 9字は8字と同じ値へ寝る(増えない)。
    modTestRunner.Check "R27-F1-2_9字は8字と同値(clampが効いている)", _
        (w9 = w8), "w9=" & Format$(w9, "0.000") & " w8=" & Format$(w8, "0.000")

    ' clampが無ければ 9^1.5 = 27 になる。27未満であることを別途固定する
    ' (上の等値だけだと「両方27」でも通ってしまうため)。
    modTestRunner.Check "R27-F1-2_9字の重みは27(=9^1.5)より小さい", _
        (w9 < 27#), "w9=" & Format$(w9, "0.000")

    ' 資料名級の長語(16字)でも8字止まり。16^1.5=64 が入るのが実機の暴走源。
    modTestRunner.Check "R27-F1-2_16字の資料名キーも8字止まり(64点にならない)", _
        (w16 = w8), "w16=" & Format$(w16, "0.000")
End Sub

' F1-1 の結線確認: KeyScore 本体が cap を通っていること。
' 9字キーが1回だけ出る短い本文の生スコアは 22.4 前後(clamp後22.627を
' 文書長正規化1.0099で割った値)で、cap(既定10)が効いていれば10になる。
' 既定値はビルド時 config(sparse_keyscore_cap=10)およびLO実行テスト
' (modConfig未注入 → 既定値へフォールバック)の双方で10。
Private Sub TestKeyScoreIsCapped24()
    Dim key As String: key = "特約条項変更届出書"          ' 9字
    Dim doc As String: doc = modSparse.CompactForMatch(key)
    Dim v As Double: v = modSparse.KeyScore(key, doc)

    modTestRunner.Check "R27-F1-1_KeyScore本体が上限10を超えない", _
        (v <= 10.000001), "v=" & Format$(v, "0.000")
    ' 0点に落ちていない(頭打ちと取りこぼしを取り違えない)。
    modTestRunner.Check "R27-F1-1_頭打ちしても加点は消えない(v>0)", _
        (v > 0#), "v=" & Format$(v, "0.000")
End Sub

Private Function BadgeIdExists24(ByRef ids() As String, ByVal target As String) As Boolean
    Dim i As Long
    For i = LBound(ids) To UBound(ids)
        If StrComp(Trim$(ids(i)), target, vbTextCompare) = 0 Then
            BadgeIdExists24 = True
            Exit Function
        End If
    Next i
End Function

Private Function IndexOf24(ByRef ids() As String, ByVal target As String) As Long
    IndexOf24 = -1
    Dim i As Long
    For i = LBound(ids) To UBound(ids)
        If StrComp(Trim$(ids(i)), target, vbTextCompare) = 0 Then
            IndexOf24 = i
            Exit Function
        End If
    Next i
End Function

' サロゲートペア(絵文字)の有無を判定する。AscWは符号付きLongを返すため
' 負値(&H8000以上)を補正してから上位/下位サロゲート範囲と比較する。
Private Function ContainsSurrogate24(ByVal s As String) As Boolean
    Dim i As Long, c As Long
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536
        ' &H8000以上の16進リテラルは"&"サフィックス無しだとIntegerの符号付き
        ' (負値)として解釈される罠がある。上のc補正(正値0-65535)と比較が
        ' 噛み合うよう、境界値は必ずLong強制の"&"サフィックス付きで書く。
        If c >= &HD800& And c <= &HDFFF& Then
            ContainsSurrogate24 = True
            Exit Function
        End If
    Next i
End Function

' ============================================================================
Public Sub RunAll24()
    On Error GoTo Count16Fail24
    TestBadgeCatalogCount16
NextNewIds24:
    On Error GoTo NewIdsFail24
    TestBadgeCatalogNewIdsExist
NextNoDup24:
    On Error GoTo NoDupFail24
    TestBadgeCatalogNoDuplicateIds
NextThreshold24:
    On Error GoTo ThresholdFail24
    TestNewBadgeConditionsMentionThreshold
NextNoEmoji24:
    On Error GoTo NoEmojiFail24
    TestBadgeTitlesHaveNoSurrogate24
NextCapKey24:
    On Error GoTo CapKeyFail24
    TestCapKeyScoreBoundary24
NextLenClamp24:
    On Error GoTo LenClampFail24
    TestKeyLenWeightClamp24
NextKeyScoreCap24:
    On Error GoTo KeyScoreCapFail24
    TestKeyScoreIsCapped24
NextDone24:
    On Error GoTo 0
    Exit Sub

Count16Fail24:
    modTestRunner.Check "TestBadgeCatalogCount16(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextNewIds24
NewIdsFail24:
    modTestRunner.Check "TestBadgeCatalogNewIdsExist(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextNoDup24
NoDupFail24:
    modTestRunner.Check "TestBadgeCatalogNoDuplicateIds(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextThreshold24
ThresholdFail24:
    modTestRunner.Check "TestNewBadgeConditionsMentionThreshold(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextNoEmoji24
NoEmojiFail24:
    modTestRunner.Check "TestBadgeTitlesHaveNoSurrogate24(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextCapKey24
CapKeyFail24:
    modTestRunner.Check "TestCapKeyScoreBoundary24(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextLenClamp24
LenClampFail24:
    modTestRunner.Check "TestKeyLenWeightClamp24(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextKeyScoreCap24
KeyScoreCapFail24:
    modTestRunner.Check "TestKeyScoreIsCapped24(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone24
End Sub
