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
    Resume NextDone24
End Sub
