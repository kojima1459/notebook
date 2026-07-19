Attribute VB_Name = "modFollowup"
Option Explicit

' ============================================================================
' modFollowup - 「続けて質問」(裁定D11)の応答パース/履歴整形の純関数群
' ----------------------------------------------------------------------------
' modAsk が文字数上限(実機VBAの1モジュール上限)に迫ったため、状態を持たない
' 純粋ヘルパーを機能単位でここへ退避した(behavior-preserving リファクタ)。
' 区切り文字はmodAsk側の確定値(FOLLOWUP_PAIR_SEP=";;;")を引数で受け取り、
' 本モジュール自身は他モジュールへ一切依存しない(String操作のみ)。
' ============================================================================

' 応答aから [[FOLLOWUP: ...]] をパースし、body=マーカー行除去後の本文 /
' candidates=候補(vbLf区切り。無ければ空)に分離する。V2と同じ寛容実装:
' "]]"が見つからない・中身が空・「なし」の場合は候補なし扱いとし、
' マーカーを含む行だけを本文から取り除く(パース失敗でも例外は出さない)。
Public Sub SplitFollowupTrailer(ByVal a As String, ByRef body As String, ByRef candidates As String)
    body = a
    candidates = ""
    If LenB(a) = 0 Then Exit Sub

    Dim fp As Long
    fp = InStr(1, a, "[[FOLLOWUP:", vbTextCompare)
    If fp = 0 Then Exit Sub

    Dim fq As Long
    fq = InStr(fp, a, "]]")
    If fq > 0 Then
        Dim fv As String
        fv = Trim$(Mid$(a, fp + Len("[[FOLLOWUP:"), fq - fp - Len("[[FOLLOWUP:")))
        If StrComp(fv, "なし", vbTextCompare) <> 0 And LenB(fv) > 0 Then
            Dim parts() As String
            parts = Split(fv, "|")
            Dim out As String
            Dim i As Long
            For i = LBound(parts) To UBound(parts)
                Dim t As String
                t = Trim$(parts(i))
                If LenB(t) > 0 Then
                    If LenB(out) > 0 Then out = out & vbLf
                    out = out & t
                End If
            Next i
            candidates = out
        End If
    End If

    ' マーカーを含む行を本文から除去し、末尾の空行・空白を刈り込む
    ' (V2 modPipeline.ParseTrailersと同じ流儀)。
    Dim lines() As String
    lines = Split(a, vbLf)
    Dim keep As String
    Dim j As Long
    For j = LBound(lines) To UBound(lines)
        If InStr(lines(j), "[[FOLLOWUP:") = 0 Then
            If LenB(keep) > 0 Then keep = keep & vbLf
            keep = keep & lines(j)
        End If
    Next j
    Do While Len(keep) > 0 And (Right$(keep, 1) = vbLf Or Right$(keep, 1) = vbCr Or Right$(keep, 1) = " ")
        keep = Left$(keep, Len(keep) - 1)
    Loop
    body = keep
End Sub

' 「新しい順<sep>区切り」文字列の先頭からmaxPairs件だけを残す。
Public Function KeepNewestPairs(ByVal joined As String, ByVal maxPairs As Long, ByVal sep As String) As String
    Dim parts() As String
    parts = Split(joined, sep)
    Dim n As Long
    n = UBound(parts) - LBound(parts) + 1
    If n <= maxPairs Then
        KeepNewestPairs = joined
        Exit Function
    End If

    Dim out As String
    Dim i As Long
    For i = LBound(parts) To LBound(parts) + maxPairs - 1
        If LenB(out) > 0 Then out = out & sep
        out = out & parts(i)
    Next i
    KeepNewestPairs = out
End Function

' 履歴に積む文字列から区切り<sep>(";;;")を除去する。単純な1回のReplaceでは
' ";;;;;;"→";;;;"のように置換結果へ再び";;;"が現れ得るため、無くなるまで
' 繰り返す(各回で必ず短くなるので有限回で終わる)。
Public Function SanitizeForFollowupHistory(ByVal s As String, ByVal sep As String) As String
    Dim t As String
    t = s
    Do While InStr(t, sep) > 0
        t = Replace(t, sep, ";;")
    Loop
    SanitizeForFollowupHistory = t
End Function
