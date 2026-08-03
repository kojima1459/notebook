Attribute VB_Name = "modFollowup"
Option Explicit

' ============================================================================
' modFollowup - 「続けて質問」(裁定D11)の応答パース/履歴整形の純関数群
' ----------------------------------------------------------------------------
' modAsk が文字数上限(実機VBAの1モジュール上限)に迫ったため、状態を持たない
' 純粋ヘルパーを機能単位でここへ退避した(behavior-preserving リファクタ)。
' 区切り文字はmodAsk側の確定値(FOLLOWUP_PAIR_SEP=";;;")を引数で受け取り、
' 本モジュール自身は他モジュールへ一切依存しない(String操作のみ)。
'
' 2026-08-03(R13-5b): これに加えて「会話出典メモリ」(直近ターンで引用した
' 資料名の短いリスト)を持つ。詳細はモジュール末尾の該当ブロックを参照。
' ============================================================================

' 覚えておく資料名の上限(R13-5b)。会話1本の文脈として意味を保てる範囲。
' 増やすほど「深掘り」が「本棚全体」に近づき、モードの区別が消える。
Private Const CITED_MAX As Long = 8

' 会話出典メモリ(vbLf区切り・新しい順・重複なし)。State Lossで消えても
' 「スコープ無し=従来動作」へ静かに戻るだけなので、保存はしない。
Private mCited As String

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
        Dim fval As String
        fval = Trim$(Mid$(a, fp + Len("[[FOLLOWUP:"), fq - fp - Len("[[FOLLOWUP:")))
        If StrComp(fval, "なし", vbTextCompare) <> 0 And LenB(fval) > 0 Then
            Dim parts() As String
            parts = Split(fval, "|")
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

' ============================================================================
' 会話出典メモリ(2026-08-03 R13-5b)
' ----------------------------------------------------------------------------
' 「深掘り=会話の流れ(引用済み資料)の中を深く掘る」を成立させるには、
' 直近のターンでどの資料を根拠にしたかを覚えておく場所が要る。それが無い
' ため、これまで deep と thorough はパラメータ量が違うだけの同じ検索だった
' (実機第2報 RC7)。
'
' ここに置く理由: 状態は「今の会話」に属し、ブックへ保存する類のものでは
' ない。VBAのState Lossで消えても「スコープ無し=従来動作」へ静かに戻るだけで
' 壊れない(消えて困る情報は持たない)。
' 保持形式は vbLf 区切りの資料名リスト(新しい順・重複なし・上限CITED_MAX件)。
' 宣言(CITED_MAX / mCited)はモジュール先頭にある。
' ============================================================================

' 直近ターンで引用した資料名を登録する。names は "|" 区切り(modAskRetrieve.
' HitSourceList の形式)でも改行区切りでも受ける。
' replaceAll=True(=followupでない新規質問)のときは覚え直す(上書き)。
Public Sub RememberCitedSources(ByVal names As String, Optional ByVal replaceAll As Boolean = False)
    If replaceAll Then mCited = ""
    mCited = MergeCitedSources(mCited, names, CITED_MAX)
End Sub

' 現在の会話出典(vbLf区切り)。空なら空文字列。
Public Function CitedSourcesLine() As String
    CitedSourcesLine = mCited
End Function

' 会話リセット(チャットのクリア)で消す。
Public Sub ClearCitedSources()
    mCited = ""
End Sub

' 検索へ渡す許可Dictionary。1件も覚えていなければ Nothing(=スコープ無し)。
Public Function CitedSourcesDict() As Object
    If LenB(mCited) = 0 Then Exit Function
    Set CitedSourcesDict = ScopeDictFrom(mCited)
End Function

' 資料名リスト(vbLf または "|" 区切り)から許可Dictionaryを作る。
' 中身が1件も無い/Dictionaryを作れない端末では Nothing を返し、
' 呼び出し側は「スコープ無し=従来動作」へ静かに退化する(無言の劣化はしない。
' 呼び出し側が usage_log へ記録する)。
Public Function ScopeDictFrom(ByVal namesLine As String) As Object
    If LenB(namesLine) = 0 Then Exit Function

    Dim d As Object
    On Error Resume Next
    Set d = CreateObject("Scripting.Dictionary")
    On Error GoTo 0
    If d Is Nothing Then Exit Function

    Dim parts() As String
    parts = Split(Replace(namesLine, "|", vbLf), vbLf)
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim nm As String
        nm = Trim$(parts(i))
        If LenB(nm) > 0 Then
            If Not d.Exists(nm) Then d.Add nm, True
        End If
    Next i
    If d.count = 0 Then Exit Function
    Set ScopeDictFrom = d
End Function

' 新しい順・重複なしで cap 件までに収めた資料名リスト(vbLf区切り)を返す。
' newNames が既存より前に来る(=最後に引用した資料が最優先で残る)。
Public Function MergeCitedSources(ByVal existing As String, ByVal newNames As String, _
                                  ByVal cap As Long) As String
    Dim lim As Long
    lim = cap
    If lim < 1 Then lim = 1

    Dim src As String
    src = Replace(newNames, "|", vbLf)
    If LenB(existing) > 0 Then
        If LenB(src) > 0 Then src = src & vbLf
        src = src & existing
    End If

    Dim parts() As String
    parts = Split(src, vbLf)
    Dim out As String
    Dim cnt As Long
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim nm As String
        nm = Trim$(parts(i))
        If LenB(nm) > 0 Then
            If Not InCitedScope(out, nm) Then
                If LenB(out) > 0 Then out = out & vbLf
                out = out & nm
                cnt = cnt + 1
                If cnt >= lim Then Exit For
            End If
        End If
    Next i
    MergeCitedSources = out
End Function

' 資料名がスコープ(vbLf区切りの許可リスト)に入っているか。
' 判定は【完全一致・大小区別あり】。検索側(modRetrieve)が使う
' Dictionary.Exists と同じ意味にすること。ここが非対称になると、
' 「スコープに入れたはずの資料が落ちる」「入れていない資料が混じる」の
' どちらかが静かに起きる。
Public Function InCitedScope(ByVal scopeLine As String, ByVal srcName As String) As Boolean
    If LenB(scopeLine) = 0 Then Exit Function
    If LenB(srcName) = 0 Then Exit Function
    InCitedScope = (InStr(1, vbLf & scopeLine & vbLf, vbLf & srcName & vbLf, vbBinaryCompare) > 0)
End Function
