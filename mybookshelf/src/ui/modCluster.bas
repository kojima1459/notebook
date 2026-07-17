Attribute VB_Name = "modCluster"
Option Explicit

' ============================================================================
' modCluster - ナレッジ地図(K-Meansクラスタリングの可視化)
' ----------------------------------------------------------------------------
' ダッシュボードに「似た資料のかたまり」をShape円で描く。ネイティブチャート
' (ChartObject/xlBubble)はExcelバージョン差が大きく本環境(LibreOffice構文
' チェックのみ)で実行検証できないため、実機で実証済みのShape APIのみで描画
' する(オーナー裁定: Shape円方式)。
'
' パイプライン:
'   1) my_vectors から最大 SAMPLE_MAX 本を等間隔サンプリングし、先頭 CLUSTER_DIM
'      次元へ切詰め+L2再正規化(Matryoshka: 先頭次元に主要情報が乗る前提)。
'   2) 球面K-Means(単位ベクトル前提=内積が大きいほど近い)。初期化は最遠点法
'      (決定的・広く散らす)。
'   3) K個の重心を古典的MDS(K×Kの二重中心化距離行列をJacobi固有分解)で2次元へ。
'   4) 各クラスタを円Shapeとして配置(サイズ=件数、位置=MDS座標、ラベル=代表
'      キーワード)。Shape名は "nxd_cluster_*"(modDashの nxd_ 一括削除で冪等)。
'
' 依存(全て公開API): modUtil.CsvToVector, modUI.UiColor, modAppDef.SH_*。
' ============================================================================

Private Const SAMPLE_MAX As Long = 400      ' クラスタリングに使う最大ベクトル数
Private Const CLUSTER_DIM As Long = 96      ' 切詰め後の次元(速度優先。Matryoshka先頭)
Private Const MAX_K As Long = 7             ' クラスタ数の上限(可読性)
Private Const KMEANS_ITERS As Long = 12     ' K-Meansの最大反復
Private Const MIN_POINTS As Long = 8        ' これ未満は地図を描かない

' ----------------------------------------------------------------------------
' DrawClusterMap - 指定領域にナレッジ地図を描く。戻り値=描いたクラスタ数
'   (0=データ不足等でスキップ。modDash側はこの戻り値でプレースホルダ切替)。
' ----------------------------------------------------------------------------
Public Function DrawClusterMap(ByVal ws As Worksheet, ByVal areaX As Double, ByVal areaY As Double, _
                               ByVal areaW As Double, ByVal areaH As Double) As Long
    On Error GoTo Fail

    Dim vecs() As Double, kw() As String, src() As String
    Dim nPts As Long
    nPts = LoadVectors(vecs, kw, src)
    If nPts < MIN_POINTS Then Exit Function

    Dim kk As Long: kk = ChooseK(nPts)

    Dim assign() As Long, centroids() As Double
    RunKMeans vecs, nPts, CLUSTER_DIM, kk, assign, centroids

    Dim sizes() As Long: ReDim sizes(0 To kk - 1)
    Dim i As Long
    For i = 0 To nPts - 1
        sizes(assign(i)) = sizes(assign(i)) + 1
    Next i

    Dim labels() As String: ReDim labels(0 To kk - 1)
    BuildLabels kk, assign, kw, src, nPts, centroids, vecs, labels

    Dim cx() As Double, cy() As Double
    MdsCoords centroids, kk, CLUSTER_DIM, cx, cy

    DrawBubbles ws, areaX, areaY, areaW, areaH, kk, cx, cy, sizes, labels

    DrawClusterMap = kk
    Exit Function

Fail:
    DrawClusterMap = 0
End Function

' ----------------------------------------------------------------------------
' 1) ベクトル読込(サンプリング+切詰め+再正規化)
' ----------------------------------------------------------------------------
Private Function LoadVectors(ByRef vecs() As Double, ByRef kw() As String, ByRef src() As String) As Long
    Dim wsV As Worksheet, wsK As Worksheet
    On Error Resume Next
    Set wsV = ThisWorkbook.Worksheets(modAppDef.SH_VECTORS)
    Set wsK = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    On Error GoTo 0
    If wsV Is Nothing Then Exit Function

    ' my_knowledge: chunk_id(1) -> source(2) & keywords(6) の索引を作る
    Dim meta As Object: Set meta = CreateObject("Scripting.Dictionary")
    If Not wsK Is Nothing Then
        Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
        Dim r As Long
        For r = 2 To lastK
            Dim cid As String: cid = CStr(wsK.Cells(r, 1).Value)
            If LenB(cid) > 0 And Not meta.Exists(cid) Then
                meta(cid) = CStr(wsK.Cells(r, 2).Value) & vbTab & CStr(wsK.Cells(r, 6).Value)
            End If
        Next r
    End If

    Dim lastV As Long: lastV = wsV.Cells(wsV.Rows.count, 1).End(xlUp).row
    If lastV < 2 Then Exit Function

    Dim total As Long: total = lastV - 1
    Dim stride As Long: stride = 1
    If total > SAMPLE_MAX Then stride = total \ SAMPLE_MAX
    If stride < 1 Then stride = 1

    ReDim vecs(0 To SAMPLE_MAX - 1, 0 To CLUSTER_DIM - 1)
    ReDim kw(0 To SAMPLE_MAX - 1)
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
            If meta.Exists(cidV) Then
                Dim parts() As String: parts = Split(CStr(meta(cidV)), vbTab)
                src(cnt) = parts(0)
                If UBound(parts) >= 1 Then kw(cnt) = parts(1)
            End If
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

' ----------------------------------------------------------------------------
' 3) 代表ラベル(クラスタ内で最頻のキーワード。無ければ重心最近傍の資料名)
' ----------------------------------------------------------------------------
Private Sub BuildLabels(ByVal kk As Long, ByRef assign() As Long, ByRef kw() As String, ByRef src() As String, _
                        ByVal nPts As Long, ByRef centroids() As Double, ByRef vecs() As Double, _
                        ByRef labels() As String)
    Dim c As Long
    For c = 0 To kk - 1
        Dim freq As Object: Set freq = CreateObject("Scripting.Dictionary")
        Dim i As Long
        For i = 0 To nPts - 1
            If assign(i) = c Then
                Dim toks() As String: toks = TokenizeKw(kw(i))
                Dim t As Long
                For t = LBound(toks) To UBound(toks)
                    Dim w As String: w = Trim$(toks(t))
                    If LenB(w) > 0 Then
                        If freq.Exists(w) Then
                            freq(w) = freq(w) + 1
                        Else
                            freq(w) = 1
                        End If
                    End If
                Next t
            End If
        Next i

        Dim bestW As String: bestW = ""
        Dim bestN As Long: bestN = 0
        Dim keyv As Variant
        For Each keyv In freq.Keys
            If freq(keyv) > bestN Then
                bestN = freq(keyv)
                bestW = CStr(keyv)
            End If
        Next keyv

        If LenB(bestW) = 0 Then bestW = NearestSource(c, assign, src, nPts, centroids, vecs)
        If LenB(bestW) = 0 Then bestW = "グループ" & (c + 1)
        labels(c) = modUtil.SafeLeft(bestW, 12)
    Next c
End Sub

Private Function TokenizeKw(ByVal s As String) As String()
    Dim t As String: t = s
    t = Replace(t, "、", ",")
    t = Replace(t, " ", ",")
    t = Replace(t, ChrW(&H3000), ",")   ' 全角スペース
    t = Replace(t, ";", ",")
    t = Replace(t, "/", ",")
    TokenizeKw = Split(t, ",")
End Function

Private Function NearestSource(ByVal c As Long, ByRef assign() As Long, ByRef src() As String, _
                               ByVal nPts As Long, ByRef centroids() As Double, ByRef vecs() As Double) As String
    Dim bestI As Long: bestI = -1
    Dim bestDot As Double: bestDot = -2#
    Dim i As Long
    For i = 0 To nPts - 1
        If assign(i) = c Then
            Dim dp As Double: dp = 0
            Dim d As Long
            For d = 0 To CLUSTER_DIM - 1
                dp = dp + vecs(i, d) * centroids(c, d)
            Next d
            If dp > bestDot Then
                bestDot = dp
                bestI = i
            End If
        End If
    Next i
    If bestI >= 0 Then NearestSource = src(bestI)
End Function

' ----------------------------------------------------------------------------
' 4) 古典的MDS(K×Kの二重中心化 → Jacobi固有分解 → 上位2軸)
' ----------------------------------------------------------------------------
Private Sub MdsCoords(ByRef centroids() As Double, ByVal kk As Long, ByVal nDim As Long, _
                      ByRef cx() As Double, ByRef cy() As Double)
    ReDim cx(0 To kk - 1)
    ReDim cy(0 To kk - 1)
    If kk <= 1 Then Exit Sub

    ' 二乗距離 D2(単位ベクトル: ||a-b||^2 = 2 - 2a・b)
    Dim d2() As Double: ReDim d2(0 To kk - 1, 0 To kk - 1)
    Dim a As Long, b As Long, d As Long
    For a = 0 To kk - 1
        For b = a + 1 To kk - 1
            Dim dp As Double: dp = 0
            For d = 0 To nDim - 1
                dp = dp + centroids(a, d) * centroids(b, d)
            Next d
            Dim dd As Double: dd = 2# - 2# * dp
            If dd < 0 Then dd = 0
            d2(a, b) = dd
            d2(b, a) = dd
        Next b
    Next a

    ' 二重中心化 B = -0.5 J D2 J
    Dim rowm() As Double: ReDim rowm(0 To kk - 1)
    Dim grand As Double: grand = 0
    For a = 0 To kk - 1
        Dim s As Double: s = 0
        For b = 0 To kk - 1
            s = s + d2(a, b)
        Next b
        rowm(a) = s / kk
        grand = grand + s
    Next a
    grand = grand / (CDbl(kk) * kk)

    Dim bmat() As Double: ReDim bmat(0 To kk - 1, 0 To kk - 1)
    For a = 0 To kk - 1
        For b = 0 To kk - 1
            bmat(a, b) = -0.5 * (d2(a, b) - rowm(a) - rowm(b) + grand)
        Next b
    Next a

    Dim eval() As Double, evec() As Double
    JacobiEigen bmat, kk, eval, evec

    Dim i1 As Long, i2 As Long
    TopTwo eval, kk, i1, i2
    Dim l1 As Double: l1 = eval(i1): If l1 < 0 Then l1 = 0
    Dim l2 As Double: l2 = eval(i2): If l2 < 0 Then l2 = 0
    Dim s1 As Double: s1 = Sqr(l1)
    Dim s2 As Double: s2 = Sqr(l2)
    For a = 0 To kk - 1
        cx(a) = evec(a, i1) * s1
        cy(a) = evec(a, i2) * s2
    Next a
End Sub

' 対称行列のJacobi固有分解(K<=7想定。evecの列=固有ベクトル)。
Private Sub JacobiEigen(ByRef ain() As Double, ByVal nn As Long, ByRef eval() As Double, ByRef evec() As Double)
    Dim m() As Double: ReDim m(0 To nn - 1, 0 To nn - 1)
    Dim i As Long, j As Long, p As Long, q As Long, k As Long
    For i = 0 To nn - 1
        For j = 0 To nn - 1
            m(i, j) = ain(i, j)
        Next j
    Next i

    ReDim evec(0 To nn - 1, 0 To nn - 1)
    For i = 0 To nn - 1
        For j = 0 To nn - 1
            evec(i, j) = 0
        Next j
        evec(i, i) = 1
    Next i

    Dim sweep As Long
    For sweep = 1 To 60
        Dim off As Double: off = 0
        For p = 0 To nn - 1
            For q = p + 1 To nn - 1
                off = off + Abs(m(p, q))
            Next q
        Next p
        If off < 0.000000001 Then Exit For

        For p = 0 To nn - 1
            For q = p + 1 To nn - 1
                If Abs(m(p, q)) > 0.000000000001 Then
                    Dim theta As Double: theta = (m(q, q) - m(p, p)) / (2# * m(p, q))
                    Dim sgn As Double: sgn = 1#
                    If theta < 0 Then sgn = -1#
                    Dim tt As Double: tt = sgn / (Abs(theta) + Sqr(theta * theta + 1#))
                    Dim cc As Double: cc = 1# / Sqr(tt * tt + 1#)
                    Dim sc As Double: sc = tt * cc

                    For k = 0 To nn - 1
                        Dim akp As Double: akp = m(k, p)
                        Dim akq As Double: akq = m(k, q)
                        m(k, p) = cc * akp - sc * akq
                        m(k, q) = sc * akp + cc * akq
                    Next k
                    For k = 0 To nn - 1
                        Dim apk As Double: apk = m(p, k)
                        Dim aqk As Double: aqk = m(q, k)
                        m(p, k) = cc * apk - sc * aqk
                        m(q, k) = sc * apk + cc * aqk
                    Next k
                    For k = 0 To nn - 1
                        Dim vkp As Double: vkp = evec(k, p)
                        Dim vkq As Double: vkq = evec(k, q)
                        evec(k, p) = cc * vkp - sc * vkq
                        evec(k, q) = sc * vkp + cc * vkq
                    Next k
                End If
            Next q
        Next p
    Next sweep

    ReDim eval(0 To nn - 1)
    For i = 0 To nn - 1
        eval(i) = m(i, i)
    Next i
End Sub

Private Sub TopTwo(ByRef eval() As Double, ByVal nn As Long, ByRef i1 As Long, ByRef i2 As Long)
    i1 = 0
    Dim i As Long
    For i = 1 To nn - 1
        If eval(i) > eval(i1) Then i1 = i
    Next i
    i2 = -1
    For i = 0 To nn - 1
        If i <> i1 Then
            If i2 = -1 Then
                i2 = i
            ElseIf eval(i) > eval(i2) Then
                i2 = i
            End If
        End If
    Next i
    If i2 = -1 Then i2 = i1
End Sub

' ----------------------------------------------------------------------------
' 5) 描画(タイトル+クラスタ円)
' ----------------------------------------------------------------------------
Private Sub DrawBubbles(ByVal ws As Worksheet, ByVal areaX As Double, ByVal areaY As Double, _
                        ByVal areaW As Double, ByVal areaH As Double, ByVal kk As Long, _
                        ByRef cx() As Double, ByRef cy() As Double, ByRef sizes() As Long, ByRef labels() As String)
    Dim titleShp As Shape
    Set titleShp = ws.Shapes.AddShape(1, areaX, areaY, 360, 20)
    titleShp.Name = "nxd_cluster_title"
    titleShp.Line.Visible = 0
    titleShp.Fill.Visible = 0
    With titleShp.TextFrame2
        .TextRange.Text = ChrW(&H1F5FA) & " ナレッジ地図(似た資料のかたまり)"
        .TextRange.Font.Size = 11.5
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
        .VerticalAnchor = 3
    End With

    Dim plotX As Double: plotX = areaX
    Dim plotY As Double: plotY = areaY + 28
    Dim plotW As Double: plotW = areaW
    Dim plotH As Double: plotH = areaH - 28
    If plotH < 90 Then plotH = 90

    Dim minX As Double: minX = cx(0)
    Dim maxX As Double: maxX = cx(0)
    Dim minY As Double: minY = cy(0)
    Dim maxY As Double: maxY = cy(0)
    Dim c As Long
    For c = 1 To kk - 1
        If cx(c) < minX Then minX = cx(c)
        If cx(c) > maxX Then maxX = cx(c)
        If cy(c) < minY Then minY = cy(c)
        If cy(c) > maxY Then maxY = cy(c)
    Next c
    Dim spanX As Double: spanX = maxX - minX
    If spanX < 0.000001 Then spanX = 1
    Dim spanY As Double: spanY = maxY - minY
    If spanY < 0.000001 Then spanY = 1

    Dim maxSize As Long: maxSize = 1
    For c = 0 To kk - 1
        If sizes(c) > maxSize Then maxSize = sizes(c)
    Next c

    Dim margin As Double: margin = 48
    For c = 0 To kk - 1
        Dim nx As Double: nx = (cx(c) - minX) / spanX
        Dim ny As Double: ny = (cy(c) - minY) / spanY
        Dim px As Double: px = plotX + margin + nx * (plotW - 2 * margin)
        Dim py As Double: py = plotY + margin + ny * (plotH - 2 * margin)
        Dim rad As Double: rad = 16 + 32 * Sqr(sizes(c) / maxSize)

        Dim ov As Shape
        Set ov = ws.Shapes.AddShape(9, px - rad, py - rad, rad * 2, rad * 2)   ' 9=楕円
        ov.Name = "nxd_cluster_b" & c
        ov.Fill.ForeColor.RGB = ClusterColor(c)
        On Error Resume Next
        ov.Fill.Transparency = 0.25
        On Error GoTo 0
        ov.Line.ForeColor.RGB = ClusterColor(c)
        ov.Line.Weight = 1
        ov.Shadow.Visible = 0
        With ov.TextFrame2
            .WordWrap = -1
            .TextRange.Text = labels(c) & vbLf & "(" & sizes(c) & ")"
            .TextRange.Font.Size = 8.5
            .TextRange.Font.Bold = -1
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = 2
            .VerticalAnchor = 3
            .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
        End With
    Next c
End Sub

Private Function ClusterColor(ByVal idx As Long) As Long
    Select Case idx Mod 7
        Case 0: ClusterColor = RGB(37, 99, 235)     ' 青
        Case 1: ClusterColor = RGB(16, 185, 129)    ' 緑
        Case 2: ClusterColor = RGB(245, 158, 11)    ' 橙
        Case 3: ClusterColor = RGB(139, 92, 246)    ' 紫
        Case 4: ClusterColor = RGB(236, 72, 153)    ' 桃
        Case 5: ClusterColor = RGB(14, 165, 233)    ' 空
        Case Else: ClusterColor = RGB(239, 68, 68)  ' 赤
    End Select
End Function
