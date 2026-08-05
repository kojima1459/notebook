Attribute VB_Name = "modCluster"
Option Explicit

' ============================================================================
' modCluster - 資料のクラスタリング(ガバナンス分析CSVのクラスタID列)
' ----------------------------------------------------------------------------
' 2026-08-05(R18-4・実機第5報③): ダッシュボードの「ナレッジ地図」(Shape円の
' 可視化)を撤去した。押しても何も起きない見るだけの絵に対し、開くたびに
' K-Meansを再計算し my_vectors を最大約800回のCOM単発読みで舐めていた。
' しかも共有フォルダ未設定の既定状態では点が MIN_POINTS に届かず、多くの端末で
' 案内文(プレースホルダ)にしかならない。オーナーからも「地図はいらない」。
' 撤去したのは可視化専用コードだけ(DrawClusterMap/DrawBubbles/ClusterColor と、
' 円の配置とラベルにしか使っていなかった MDS/Jacobi/BuildLabels 一式)。
' 再導入したくなったときの設計案は HANDOFF の次期課題に残してある。
'
' 残っているもの(= modAnalytics.ExportAnalyticsCsv のクラスタID列が現役で使う):
'   1) my_vectors から最大 SAMPLE_MAX 本を等間隔サンプリングし、先頭 CLUSTER_DIM
'      次元へ切詰め+L2再正規化(Matryoshka: 先頭次元に主要情報が乗る前提)。
'   2) 球面K-Means(単位ベクトル前提=内積が大きいほど近い)。初期化は最遠点法
'      (決定的・広く散らす)。
'   3) 資料ごとに、所属チャンクの多数決でクラスタIDを決める(SourceClusterMap)。
'
' 依存(全て公開API): modUtil.CsvToVector, modAppDef.SH_*。
' ============================================================================

Private Const SAMPLE_MAX As Long = 400      ' クラスタリングに使う最大ベクトル数
Private Const CLUSTER_DIM As Long = 96      ' 切詰め後の次元(速度優先。Matryoshka先頭)
Private Const MAX_K As Long = 7             ' クラスタ数の上限(可読性)
Private Const KMEANS_ITERS As Long = 12     ' K-Meansの最大反復
Private Const MIN_POINTS As Long = 8        ' これ未満はクラスタリングしない

' ----------------------------------------------------------------------------
' SourceClusterMap - 資料名→クラスタID(Long)の辞書を返す(分析CSV用)。
'   同じK-Means結果を使い、各資料は所属チャンクの多数決クラスタに割り当てる。
'   データ不足時は空の辞書を返す(呼び出し側は Exists で存在確認する)。
' ----------------------------------------------------------------------------
Public Function SourceClusterMap() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    On Error GoTo Done

    Dim vecs() As Double, src() As String
    Dim nPts As Long: nPts = LoadVectors(vecs, src)
    If nPts < MIN_POINTS Then GoTo Done

    Dim kk As Long: kk = ChooseK(nPts)
    Dim assign() As Long, centroids() As Double
    RunKMeans vecs, nPts, CLUSTER_DIM, kk, assign, centroids

    ' 資料×クラスタの出現回数を数える(キー: "資料名|クラスタ")
    Dim cnt As Object: Set cnt = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 0 To nPts - 1
        If LenB(src(i)) > 0 Then
            Dim ck As String: ck = src(i) & "|" & assign(i)
            If cnt.Exists(ck) Then cnt(ck) = cnt(ck) + 1 Else cnt(ck) = 1
            If Not d.Exists(src(i)) Then d(src(i)) = assign(i)
        End If
    Next i

    ' 多数決で各資料のクラスタを確定
    Dim sName As Variant
    For Each sName In d.Keys
        Dim bestC As Long: bestC = CLng(d(sName))
        Dim bestN As Long: bestN = -1
        Dim c As Long
        For c = 0 To kk - 1
            Dim ck2 As String: ck2 = sName & "|" & c
            If cnt.Exists(ck2) Then
                If cnt(ck2) > bestN Then
                    bestN = cnt(ck2)
                    bestC = c
                End If
            End If
        Next c
        d(sName) = bestC
    Next sName

Done:
    Set SourceClusterMap = d
End Function

' ----------------------------------------------------------------------------
' 1) ベクトル読込(サンプリング+切詰め+再正規化)
' ----------------------------------------------------------------------------
' 2026-08-05(R18H FB-3 / A-L10): kw()(keywords列)の読みを削除した。R18-4 で
' 可視化(BuildLabels/TokenizeKw)を撤去した時点で消費者が1つも居なくなり、
' my_knowledge の6列目を全行ぶん読むコストと配列だけが残っていた。
' あわせて索引の一括読みも1〜2列で足りるようになる(2万行で4列ぶんの純減)。
Private Function LoadVectors(ByRef vecs() As Double, ByRef src() As String) As Long
    Dim wsV As Worksheet, wsK As Worksheet
    On Error Resume Next
    Set wsV = ThisWorkbook.Worksheets(modAppDef.SH_VECTORS)
    Set wsK = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    On Error GoTo 0
    If wsV Is Nothing Then Exit Function

    ' my_knowledge: chunk_id(1) -> source(2) の索引を作る
    '
    ' 2026-07-28(レビュー M-5): ここは1セルずつ読んでいた。2万チャンクなら
    ' 3列×2万=6万回のCOM往復で、しかもダッシュボード表示と分析CSVで2回走る。
    ' 大規模な本棚でダッシュボードが数十秒フリーズする直接の原因だった。
    ' modRetrieve と同じく Range 一括読みへ揃える(MASTER_SPEC §12)。
    Dim meta As Object: Set meta = CreateObject("Scripting.Dictionary")
    If Not wsK Is Nothing Then
        Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
        If lastK >= 2 Then
            Dim arrK As Variant
            arrK = wsK.Range(wsK.Cells(2, 1), wsK.Cells(lastK, 2)).Value
            Dim r As Long
            For r = LBound(arrK, 1) To UBound(arrK, 1)
                Dim cid As String: cid = CStr(arrK(r, 1))
                If LenB(cid) > 0 And Not meta.Exists(cid) Then
                    meta(cid) = CStr(arrK(r, 2))
                End If
            Next r
        End If
    End If

    Dim lastV As Long: lastV = wsV.Cells(wsV.Rows.count, 1).End(xlUp).row
    If lastV < 2 Then Exit Function

    Dim total As Long: total = lastV - 1
    Dim stride As Long: stride = 1
    If total > SAMPLE_MAX Then stride = total \ SAMPLE_MAX
    If stride < 1 Then stride = 1

    ReDim vecs(0 To SAMPLE_MAX - 1, 0 To CLUSTER_DIM - 1)
    ReDim src(0 To SAMPLE_MAX - 1)

    Dim cnt As Long: cnt = 0
    Dim rowV As Long
    For rowV = 2 To lastV Step stride
        If cnt >= SAMPLE_MAX Then Exit For
        Dim full() As Double
        If modUtil.CsvToVector(CStr(wsV.Cells(rowV, 2).Value), full) Then
            Dim ubF As Long: ubF = UBound(full)
            Dim lastD As Long: lastD = CLUSTER_DIM - 1
            If ubF < lastD Then lastD = ubF
            Dim ss As Double: ss = 0
            Dim d As Long
            For d = 0 To lastD
                vecs(cnt, d) = full(d)
                ss = ss + full(d) * full(d)
            Next d
            For d = lastD + 1 To CLUSTER_DIM - 1
                vecs(cnt, d) = 0
            Next d
            If ss > 0 Then
                Dim inv As Double: inv = 1# / Sqr(ss)
                For d = 0 To CLUSTER_DIM - 1
                    vecs(cnt, d) = vecs(cnt, d) * inv
                Next d
            End If
            Dim cidV As String: cidV = CStr(wsV.Cells(rowV, 1).Value)
            If meta.Exists(cidV) Then src(cnt) = CStr(meta(cidV))
            cnt = cnt + 1
        End If
    Next rowV
    LoadVectors = cnt
End Function

Private Function ChooseK(ByVal nPts As Long) As Long
    Dim kk As Long: kk = Int(Sqr(CDbl(nPts) / 2#)) + 1
    If kk > MAX_K Then kk = MAX_K
    If kk < 2 Then kk = 2
    If kk > nPts Then kk = nPts
    ChooseK = kk
End Function

' ----------------------------------------------------------------------------
' 2) 球面K-Means
' ----------------------------------------------------------------------------
Private Sub RunKMeans(ByRef vecs() As Double, ByVal nPts As Long, ByVal nDim As Long, ByVal kk As Long, _
                      ByRef assign() As Long, ByRef centroids() As Double)
    ReDim centroids(0 To kk - 1, 0 To nDim - 1)
    ReDim assign(0 To nPts - 1)
    InitCentroidsFarthest vecs, nPts, nDim, kk, centroids

    Dim it As Long, i As Long, c As Long, d As Long
    For it = 1 To KMEANS_ITERS
        DoEvents   ' 反復ごとに息継ぎ(重い計算中もExcelを「応答なし」にしない)
        Dim changed As Boolean: changed = False

        ' 割当(内積最大の重心へ)
        For i = 0 To nPts - 1
            Dim best As Long: best = 0
            Dim bestDot As Double: bestDot = -2#
            For c = 0 To kk - 1
                Dim dp As Double: dp = 0
                For d = 0 To nDim - 1
                    dp = dp + vecs(i, d) * centroids(c, d)
                Next d
                If dp > bestDot Then
                    bestDot = dp
                    best = c
                End If
            Next c
            If assign(i) <> best Then changed = True
            assign(i) = best
        Next i

        ' 更新(平均→再正規化。空クラスタは重心据置き)
        Dim sums() As Double: ReDim sums(0 To kk - 1, 0 To nDim - 1)
        Dim counts() As Long: ReDim counts(0 To kk - 1)
        For i = 0 To nPts - 1
            Dim a As Long: a = assign(i)
            counts(a) = counts(a) + 1
            For d = 0 To nDim - 1
                sums(a, d) = sums(a, d) + vecs(i, d)
            Next d
        Next i
        For c = 0 To kk - 1
            If counts(c) > 0 Then
                Dim ss As Double: ss = 0
                For d = 0 To nDim - 1
                    Dim mv As Double: mv = sums(c, d) / counts(c)
                    centroids(c, d) = mv
                    ss = ss + mv * mv
                Next d
                If ss > 0 Then
                    Dim inv As Double: inv = 1# / Sqr(ss)
                    For d = 0 To nDim - 1
                        centroids(c, d) = centroids(c, d) * inv
                    Next d
                End If
            End If
        Next c

        If Not changed And it > 1 Then Exit For
    Next it
End Sub

' 最遠点法の初期重心(決定的): 先頭点→既存重心から最も遠い点、を繰り返す。
Private Sub InitCentroidsFarthest(ByRef vecs() As Double, ByVal nPts As Long, ByVal nDim As Long, _
                                  ByVal kk As Long, ByRef centroids() As Double)
    Dim d As Long
    For d = 0 To nDim - 1
        centroids(0, d) = vecs(0, d)
    Next d

    Dim c As Long
    For c = 1 To kk - 1
        Dim bestI As Long: bestI = 0
        Dim bestScore As Double: bestScore = -1#
        Dim i As Long
        For i = 0 To nPts - 1
            Dim mind As Double: mind = 2#
            Dim j As Long
            For j = 0 To c - 1
                Dim dp As Double: dp = 0
                For d = 0 To nDim - 1
                    dp = dp + vecs(i, d) * centroids(j, d)
                Next d
                Dim dist As Double: dist = 1# - dp
                If dist < mind Then mind = dist
            Next j
            If mind > bestScore Then
                bestScore = mind
                bestI = i
            End If
        Next i
        For d = 0 To nDim - 1
            centroids(c, d) = vecs(bestI, d)
        Next d
    Next c
End Sub
