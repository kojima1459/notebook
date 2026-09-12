Attribute VB_Name = "modHubStat"
Option Explicit

' 統計タイルの幾何(2026-07-31 R7 A-4で modHub から移設)。
' 高さ・段数はHubのバッジ帯の位置決め(modHub.StatTilesBottom)も使うので、
' 数字の出どころをここ1箇所にする。
Private Const TILE_H As Double = 52
Private Const TILE_GAP_Y As Double = 8
Private Const TILE_ROWS As Long = 4
Private Const TILE_COUNT As Long = 8

' Hub最下部のフッター(2026-08-05 R18-5b)。高さと、押すと開く社内ポータルの
' URL。URLはコード内でここ1箇所だけに持つ(2箇所に書くと必ず片方が古くなる)。
Private Const FOOTER_H As Double = 16
Private Const PORTAL_URL As String = _
    "http://www.portal.s1.ms-ad-ins.co.jp/loader/hp/OpenContents/A201203280048/toppage.html"

' ----------------------------------------------------------------------------
' 受信箱の共有問い合わせ(2026-07-31 R8 F3)
' ----------------------------------------------------------------------------
' modChannel.PendingUpdates は、購読中の全部門の version.txt を
' ADODB.Stream で開いて読む【実I/O】。Hubの受信箱はこれを描画のたびに
' 呼んでおり、しかも Hub の初期描画は起動シーケンスの真ん中にある。
' 遅い/届かない共有では、そこで起動そのものが止まる。
'   ・起動中(AllowShareQueries が呼ばれるまで)は一切問い合わせない。
'     受信箱は「押して確認」のプレースホルダを出す。
'   ・起動後の再描画では問い合わせるが、結果を10分だけ持ち回す。
'     Hubは画面を戻るたびに描き直されるので、TTLが無いと人数×往復が
'     そのままファイルサーバの負荷になる。
Private mShareQueryOk As Boolean
Private mPendCache As String
Private mPendAt As Double
Private Const PEND_TTL_SEC As Double = 600#

' AllowShareQueries - 起動シーケンスの完了を知らせる(modBoot が最後に呼ぶ)。
Public Sub AllowShareQueries()
    mShareQueryOk = True
End Sub

' ShareQueriesAllowed - 受信箱が「押して確認」を出すべきか判断するための印。
Public Function ShareQueriesAllowed() As Boolean
    ShareQueriesAllowed = mShareQueryOk
End Function

' PendingUpdatesCached - 更新のある部門一覧("|"区切り)。
'   起動中は必ず空文字(共有I/Oを走らせない)。
'   起動後は10分TTLのキャッシュ越しに modChannel へ問い合わせる。
Public Function PendingUpdatesCached() As String
    If Not mShareQueryOk Then Exit Function
    If modShareRule.CacheIsFresh(mPendAt, Timer, PEND_TTL_SEC) Then
        PendingUpdatesCached = mPendCache
        Exit Function
    End If
    On Error Resume Next
    mPendCache = modChannel.PendingUpdates()
    On Error GoTo 0
    mPendAt = Timer
    PendingUpdatesCached = mPendCache
End Function

' InvalidatePending - 取り込み直後など、キャッシュを捨てて次回に取り直させる。
Public Sub InvalidatePending()
    mPendCache = ""
    mPendAt = 0
End Sub

' ----------------------------------------------------------------------------
' OnSyncPending - 受信箱の「【部門】に更新があります」を押したときの実行口
'   (2026-07-31 R11-E・監査4)。従来はここが modKnowledge.OnChannels(全部門を
'   まとめて再取込する重い確認ダイアログ付き)に配線されていて、「更新のある
'   部門だけ静かに取り込む」というバッジの意味と実際の挙動がずれていた。
'   modChannel.SyncSubscribed は購読中で更新のある部門だけを対象にする軽量版
'   だが、これまで呼び出し元が1つも無かった(押下手段の無いバッジ)。
' ----------------------------------------------------------------------------
Public Sub OnSyncPending()
    ' R15-2a(実機第4報 RC8): 取込・同期の最中は受け付けない。ここは部門の
    ' 再取込(modChannel.SyncSubscribed)を始めるため、取込中に押されると
    ' 取込の途中から入れ子でもう1本の取込が走り出す(E0202の温床)。
    ' 順序は既存40箇所超と同型: BlockIfIngesting は Enter より前に置く
    ' (後ろに置くとロックを取ったまま Exit してUIが10分死ぬ)。
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUIMain.ShowProgress "部門の更新を取り込んでいます…"
    Dim failN As Long
    Dim lastErr As String
    Dim got As Long: got = modChannel.SyncSubscribed(failN, lastErr)
    modUIMain.HideProgress
    InvalidatePending
    modUiLock.Leave

    ' 2026-07-31(R11-H Med5): 取り込めなかった部門があったことを err_log にも
    ' 残す。従来はトーストで一瞬「N部門は取り込めませんでした」と流れるだけで、
    ' 「うちの部門にだけ正典が届かない」を後から調べる材料がゼロだった
    ' (憲章§4-1 無言の失敗禁止)。
    If failN > 0 Then
        modLog.LogError "E0801", "modHubStat.OnSyncPending", _
            failN & "部門の更新を取り込めませんでした。最後の理由: " & _
            modUtil.SafeLeft(lastErr, 300)
    End If

    ' 文言の整合(R11-H Med5): 「成功で0件(すべて最新)」と「失敗」を
    ' 分けて言う。両方0のときだけ「更新はありませんでした」。
    Dim msg As String
    If got > 0 Then
        msg = got & "件のまとまりを取り込みました。"
    ElseIf failN > 0 Then
        msg = "更新を取り込めませんでした。"
    Else
        msg = "更新はありませんでした(すべて最新です)。"
    End If
    If failN > 0 Then msg = msg & "(" & failN & "部門は取り込めませんでした)"
    modSkin.ShowToast msg, IIf(failN > 0, "error", "info")

    modHub.EnsureHubLayout
    On Error GoTo 0
End Sub

' 更新保留チャンネルの表示ラベル。"|"区切りの保留リストから、
' 1件ならその名前、2件以上なら「先頭ほか N件」を返す。
' 2026-07-28(レビュー H-13): 保留の実体を見ずにアクティブ部門名を出して
' いたため、「押しても『既に最新です』でバッジが消えない」という
' 説明不能な状態になっていた。名前は必ず保留リストから取る。
Public Function PendingLabel(ByVal pendList As String) As String
    On Error Resume Next
    Dim s As String: s = Trim$(pendList)
    If LenB(s) = 0 Then Exit Function
    Dim parts() As String: parts = Split(s, "|")
    Dim n As Long: n = UBound(parts) - LBound(parts) + 1
    If n <= 1 Then
        PendingLabel = Trim$(parts(LBound(parts)))
    Else
        PendingLabel = Trim$(parts(LBound(parts))) & " ほか" & (n - 1) & "部門"
    End If
    On Error GoTo 0
End Function

' ========================================
' modHubStat - Hub(ホーム)が出す数値の取得・整形と、Hub図形の一括削除
'
' modHub から切り出した裏方。タイルに出る数字が変な話と、タイルの位置が
' ずれる話を別々に追えるようにする。RemoveHubShapes は旧ホーム画面の
' btn_/lbl_ まで消す再描画の要で、描画側と一緒にいると見落とされやすい。
'
' 切り出しの理由(2026-07-28): modHub が契約上限30,000字に対し残り712字で、
' 更新バッジの修正(レビュー H-13)を入れる余裕が乏しかった(レビュー I-2)。
' ========================================

' nx_hub_ に加え旧ホーム画面のbtn_/lbl_も消す(残ると上に浮く)。
Public Sub RemoveHubShapes(ByVal ws As Worksheet)
    Dim names() As String
    ReDim names(0 To ws.Shapes.Count)
    Dim n As Long
    Dim shp As Shape
    For Each shp In ws.Shapes
        Dim nm As String: nm = shp.Name
        If Left$(nm, 7) = "nx_hub_" Or Left$(nm, 4) = "btn_" Or Left$(nm, 4) = "lbl_" Then
            names(n) = nm
            n = n + 1
        End If
    Next shp
    Dim i As Long
    For i = 0 To n - 1
        On Error Resume Next
        ws.Shapes(names(i)).Delete
        On Error GoTo 0
    Next i
End Sub

' ----------------------------------------------------------------------------
' ClearBadgeArea - 旧バッジ表示領域(B10:F60=HUB_BOUND上限。F6/m-4で余裕
'   ゼロだったB10:F56から拡張)のUnMerge+
'   ClearContentsを1回だけ行う(R25-1a-2)。modHub.DrawBadgesは実測r
'   (StatTilesBottomの実下端由来)へ毎回Mergeし直すため、前回描画からrが
'   動くと旧領域の値が新しい結合先セルに残ったまま結合され、Excel標準の
'   「複数の値を持つセル範囲」警告(→パイプライン凍結)を招く。DrawBadgesの
'   2つのMerge呼び出しより前に必ず1回呼ぶこと(modHubは容量逼迫のため実体を
'   ここへ置く。docs/dev/spec_20260810_R25_実機第11報.md FA-R25-1a-2)。
' ----------------------------------------------------------------------------
Public Sub ClearBadgeArea(ByVal ws As Worksheet)
    On Error Resume Next
    With ws.Range("B10:F60")
        .UnMerge
        .ClearContents
    End With
    On Error GoTo 0
End Sub

' 数値を必ず表示できる文字列にする(空欄にしない)。
Public Function NumText(ByVal v As Long) As String
    NumText = CStr(v)
    If LenB(NumText) = 0 Then NumText = "0"
End Function

Public Function SafeStat(ByVal key As String) As Long
    On Error Resume Next
    SafeStat = modStats.GetStat(key)
    On Error GoTo 0
End Function

' 質問回数はモード別カウンタの合算(単一のquestion_totalキーは存在しない)。
' 2026-08-03(R14-1a): 合算そのものは modStats.AskTotalAll が唯一の持ち主。
' ここで quick+deep と書いていたため入念モードの質問が Hub のタイルから
' 丸ごと落ちていた(実機第3報 RC1)。以後は1行の委譲に留める。
Public Function AskTotal() As Long
    On Error Resume Next
    AskTotal = modStats.AskTotalAll()
    On Error GoTo 0
End Function

Public Function SafeSavedMinutes() As Long
    On Error Resume Next
    SafeSavedMinutes = modStats.SavedMinutesEstimate()
    On Error GoTo 0
End Function

Public Function SafeChunks() As Long
    On Error Resume Next
    SafeChunks = modShelf.TotalChunks()
    On Error GoTo 0
End Function

' 本棚の使用量。実数だけ出しても上限が分からないので割合で見せる。
' 部門チャンネルを増やすほど埋まるので、増やしてよいかの判断材料になる。
Public Function ChunkUsage() As String
    Dim pct As Long
    On Error Resume Next
    pct = modChannel.ChunkUsagePercent()
    On Error GoTo 0
    ChunkUsage = pct & "%"
End Function

Public Function FmtMin(ByVal minutes As Long) As String
    If minutes < 60 Then
        FmtMin = CStr(minutes) & "分"
    Else
        FmtMin = CStr(minutes \ 60) & "時間"
        If (minutes Mod 60) > 0 Then FmtMin = FmtMin & CStr(minutes Mod 60) & "分"
    End If
    If LenB(FmtMin) = 0 Then FmtMin = "0分"
End Function

' タイル既定値(要件D・2026-07-30 R3)。DrawStatTilesが値のLenB=0を
' 検出したときに埋める既定文字列。0でも意味が通じるよう単位付きで返す
' (数字だけの"0"だと本当にゼロなのか値取得に失敗したのか画面上で
' 区別が付かないため)。idxはDrawStatTiles内のlabels/valsと同じ並び
' (0始まり、8枚固定)。modHubの30,000字上限に余裕が無かったため、
' 数値整形の裏方であるこちらへ置く(modHubStat切り出しの元々の理由と同じ)。
Public Function DefaultTileValue(ByVal idx As Long) As String
    Select Case idx
        Case 2, 4, 5   ' 自分の節約時間 / みんなの節約(今日) / みんなの節約(今月)
            DefaultTileValue = "0分"
        Case 6         ' 連続ログイン
            DefaultTileValue = "0日"
        Case 3         ' 本棚の使用量
            DefaultTileValue = "0%"
        Case Else      ' 質問した回数 / 自己解決 / パック共有
            DefaultTileValue = "0"
    End Select
End Function

Public Function OrgMin(ByVal period As String) As String
    Dim v As Long
    On Error Resume Next
    If period = "d" Then
        v = modBoard.OrgMinutesDay()
    Else
        v = modBoard.OrgMinutesMon()
    End If
    OrgMin = FmtMin(v) & modBoard.OrgApproxSuffix()   ' R33H F18: 母数つきの概算
    On Error GoTo 0
End Function

' タイル群が占める高さ。Hubのバッジ帯はこの下に置く。
Public Function TilesHeight() As Double
    TilesHeight = TILE_ROWS * (TILE_H + TILE_GAP_Y)
End Function

' ============================================================================
' DrawStatTiles - 統計タイル8枚(2026-07-31 R7 A-4で modHub から移設・再構成)
' ----------------------------------------------------------------------------
' 実機で「節約できた時間/みんな(今日)/みんな(今月)/連続ログイン」の4枚が
' ラベルだけになり、値が消えた。ところが hub_tile_empty は1件も出ていない
' =値は空ではなく、【描画の段階で見えなくなっている】。
'
' 旧実装は1枚のShapeに「ラベル vbCr 値」を入れ、TextFrame2.Paragraphs(2) に
' 別書式を当てていた。この段落2番目への書式適用は、Shapeの高さ・余白・
' フォント・DPIの組み合わせ次第で2行目が枠外へ送られる(はみ出した段落は
' 描画されない)うえ、Paragraphs(2)への代入自体が失敗しても
' On Error Resume Next で握られるため、失敗しても何も残らない。
'
' 原因の深追いより構造を変える。タイルは
'     背景Shape(nx_hub_tile<i>) + ラベルShape(nx_hub_tl<i>) + 値Shape(nx_hub_tv<i>)
' の分離構成にし、値は独立したテキストボックスへ直接書く(段落操作を全廃)。
' こうすると値の描画はラベルの高さにも段落書式にも依存しない。
'
' 併せて、描画のたびに8枚分の「idx=<n> len=<値の文字数>」を1行だけ
' usage_log へ残す。次に消えたときに「値が空だったのか、座標がおかしいのか」
' をログだけで切り分けられるようにするため。
' ============================================================================
Public Sub DrawStatTiles(ByVal ws As Worksheet, ByVal topY As Double)
    Dim labels As Variant, vals As Variant
    labels = Array("質問した回数", ChrW(&HD83D) & ChrW(&HDFE2) & " 自己解決", _
                   modEmj.Stopwatch() & " 自分の節約時間", ChrW(&HD83D) & ChrW(&HDCD6) & " 本棚の使用量", _
                   ChrW(&HD83C) & ChrW(&HDF0D) & " みんなの節約(今日)", ChrW(&HD83C) & ChrW(&HDF0D) & " みんなの節約(今月)", _
                   ChrW(&HD83D) & ChrW(&HDD25) & " 連続ログイン", ChrW(&HD83D) & ChrW(&HDCE6) & " パック共有")
    ' 実機報告(2026-07-27)「一部のタイルが真っ白」対策: Array()内で直接
    ' 関数を呼ぶと1つの失敗が空文字になる。1つずつ受けて必ず値を入れる。
    Dim vAsk As String, vSolve As String, vSaved As String, vUse As String
    Dim vOrgD As String, vOrgM As String, vStreak As String, vPack As String
    vAsk = NumText(AskTotal())
    vSolve = NumText(SafeStat("selfsolve_total"))
    vSaved = FmtMin(SafeSavedMinutes())
    vUse = ChunkUsage()
    ' R13-7b(正直な表示): 共有フォルダへ届いていないセッションでは「みんなの
    ' 節約」を0分と偽らず「未接続」と出す。modShare.Reachable()は1セッション
    ' 1回だけ実際に確かめてキャッシュする関所(2回目以降は判定コストゼロ)
    ' なので、ここで呼んでも再描画のたびに新しい網羅プローブは増えない。
    '
    ' R13 F6: ただし Reachable() は「未設定」と「設定済みだが届かない」の
    ' 両方で False を返す。配布既定の共有パスが空になった今、そのままでは
    ' 共有をまだ設定していない大多数の利用者に「未接続」+VPNの確認を促す
    ' ことになり、これは事実に反する案内(そもそも繋ぎに行っていない)。
    ' パスが空かどうかを先に見て、3状態(未設定 > 未接続 > 通常)に分ける。
    ' 判定材料は modShare.BasePath()(到達性を見ない設定値だけの窓口)。
    Dim shareSet As Boolean
    On Error Resume Next
    shareSet = (LenB(modShare.BasePath()) > 0)
    On Error GoTo 0
    If Not shareSet Then
        ' R20-4c(実機第7報③④): 「未設定」だけでは何をすればよいか分からない。
        ' タイル自体がクリック導線(modBoard.OnWidgetClick)なので、それを示す。
        vOrgD = "クリックで設定"
        vOrgM = "クリックで設定"
    ElseIf modShare.Reachable() Then
        vOrgD = OrgMin("d")
        vOrgM = OrgMin("m")
    Else
        vOrgD = "未接続"
        vOrgM = "未接続"
    End If
    vStreak = NumText(SafeStat("streak_days")) & "日"
    vPack = NumText(SafeStat("pack_export_total"))
    vals = Array(vAsk, vSolve, vSaved, vUse, vOrgD, vOrgM, vStreak, vPack)

    ' 要件D/E(2026-07-30 R3): 値のLenB=0を検出したら種別ごとの既定値へ置換して
    ' 描画を続け(タイルを絶対に空文字にしない)、事実だけ1行ログへ残す。
    Dim tileIdx As Long
    For tileIdx = 0 To TILE_COUNT - 1
        If LenB(CStr(vals(tileIdx))) = 0 Then
            On Error Resume Next
            modLog.LogUsage "hub_tile_empty", "", _
                "タイル" & tileIdx & ":" & CStr(labels(tileIdx)) & " の値が空でした"
            On Error GoTo 0
            vals(tileIdx) = DefaultTileValue(tileIdx)
        End If
    Next tileIdx

    ' 左ブロックの幾何。D列(溝)を挟んで B:C と E:F の2枚並び。
    Dim colL As Double, colW As Double, gutter As Double
    colL = ws.Range("B1").Left
    colW = ws.Range("B1:C1").Width
    gutter = ws.Range("D1").Width

    Dim drawLog As String
    Dim i As Long
    For i = 0 To TILE_COUNT - 1
        Dim x As Double, y As Double
        x = colL + (i Mod 2) * (colW + gutter)
        y = topY + (i \ 2) * modViewport2.SY(TILE_H + TILE_GAP_Y)

        Dim act As String
        act = ""
        ' みんなの節約(i=4,5)はクリックで部内ランキングを出す
        ' (旧サイドバーウィジェットのクリック機能の移設先)。
        If i = 4 Or i = 5 Then act = "modBoard.OnWidgetClick"

        DrawOneTile ws, i, x, y, colW, CStr(labels(i)), CStr(vals(i)), act

        If LenB(drawLog) > 0 Then drawLog = drawLog & " "
        drawLog = drawLog & "idx=" & i & " len=" & Len(CStr(vals(i)))
    Next i

    On Error Resume Next
    modLog.LogUsage "hub_tile_drawn", "", drawLog
    On Error GoTo 0
End Sub

' タイル1枚 = 背景 + ラベル + 値 の3Shape。1枚の失敗で残りを道連れにしない。
Private Sub DrawOneTile(ByVal ws As Worksheet, ByVal idx As Long, ByVal x As Double, _
                        ByVal y As Double, ByVal tileW As Double, ByVal labelText As String, _
                        ByVal valueText As String, ByVal action As String)
    On Error Resume Next
    Dim tile As Shape
    Set tile = ws.Shapes.AddShape(5, x, y, tileW, modViewport2.SY(TILE_H))
    If Not tile Is Nothing Then
        tile.Name = "nx_hub_tile" & idx
        tile.Adjustments(1) = 0.12
        tile.Line.Visible = -1
        tile.Line.Weight = 0.75
        tile.Line.ForeColor.RGB = modUI.UiColor("border")
        tile.Fill.ForeColor.RGB = modUI.UiColor("surface")
        modSkin.ApplyLightShadow tile
        If LenB(action) > 0 Then
            tile.OnAction = action
            tile.AlternativeText = "クリックで部内の節約ランキングを表示"
        End If
    End If
    Set tile = Nothing
    Err.Clear
    On Error GoTo 0

    TileText ws, "nx_hub_tl" & idx, x + 12, y + 6, tileW - 20, 13, labelText, _
             8, False, modUI.UiColor("muted"), action
    TileText ws, "nx_hub_tv" & idx, x + 12, y + 21, tileW - 20, 24, valueText, _
             16, True, modUI.UiColor("primary"), action
End Sub

' 枠も塗りも持たない素のテキストShape。値はここへ直接書く(段落操作をしない)。
Private Sub TileText(ByVal ws As Worksheet, ByVal shapeName As String, ByVal x As Double, _
                     ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                     ByVal bodyText As String, ByVal fontSize As Double, _
                     ByVal isBold As Boolean, ByVal rgbVal As Long, ByVal action As String)
    On Error Resume Next
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(1, x, y, w, h)
    If shp Is Nothing Then Exit Sub
    shp.Name = shapeName
    shp.Line.Visible = 0
    shp.Fill.Visible = 0
    With shp.TextFrame2
        .WordWrap = 0          ' 折り返さない(2行目送りで値が消えるのを防ぐ)
        .AutoSize = 0
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .TextRange.Text = bodyText
        .TextRange.Font.Size = fontSize
        If isBold Then .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = rgbVal
        .VerticalAnchor = 3
    End With
    ' 背景タイルと同じ行き先にする(ラベル/値の上を押しても反応しない、
    ' という「押せるのに押せない」状態を作らない)。
    If LenB(action) > 0 Then shp.OnAction = action
    Set shp = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 質問カード3枚(Hubの右カラム)。2026-08-05(R18-5a)に modHub から移設した。
' ----------------------------------------------------------------------------
' 実機第5報④: ここは「うっすら地色の平たいチップ」で、すぐ上のナビボタン
' だけがカードに見えるちぐはぐな状態だった。押せると気づかれなければ押せない
' のと同じ(憲章§3-1)。modHub.DrawNavButtons と同じ型へ揃える。
' 【3枚は同色にし、区別はアイコンで付ける】: パレットの意味づけ済みの色は
' primary/accent の2系統しかなく、3色に塗り分けると5テーマ分のコントラスト
' 検証が要るハードコードRGBへ逆戻りする(R12-7-4のa11y監査で通した線)。
' 幾何(L/W/topY/cardH)の持ち主はHub側。Shape名 nx_hub_qa0〜2 と OnAction は
' modHub.OnQuickAsk のCase分岐と対。
Public Sub DrawQuickAskCards(ByVal ws As Worksheet, ByVal L As Double, _
                             ByVal W As Double, ByVal topY As Double, _
                             ByVal cardH As Double)
    ' 見出しは置かない。カードの文面自体が「こう聞けばいい」の見本になっている。
    ' 1段目=何を聞くか(太字)、2段目=聞くとどう返るかの一言。
    Dim caps As Variant, subs As Variant
    caps = Array(ChrW(&HD83D) & ChrW(&HDCCC) & " 改定ポイント", _
                 ChrW(&HD83D) & ChrW(&HDCD6) & " 用語をやさしく", _
                 ChrW(&HD83D) & ChrW(&HDC63) & " 手続きの流れ")
    subs = Array("何が変わった?", "この言葉の意味?", "いつ何を出す?")
    Dim cardW As Double: cardW = (W - 12) / 3
    Dim i As Long
    For i = 0 To 2
        On Error Resume Next
        Dim card As Shape
        Set card = ws.Shapes.AddShape(5, L + i * (cardW + 6), topY, cardW, cardH)
        If Not card Is Nothing Then
            card.Name = "nx_hub_qa" & i
            card.Adjustments(1) = 0.14
            card.Line.Visible = -1
            card.Line.Weight = 1#
            card.Line.ForeColor.RGB = modUI.UiColor("accent")
            card.Fill.ForeColor.RGB = modUI.UiColor("surface")
            modSkin.ApplyLightShadow card
            ' 段落で書式を分けるため区切りはvbCr(vbLfだとParagraphs(2)が範囲外)。
            With card.TextFrame2
                .WordWrap = -1
                .TextRange.Text = CStr(caps(i)) & vbCr & CStr(subs(i))
                .TextRange.Font.Size = 8.5     ' 2段目(説明)の大きさ
                .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
                If .TextRange.Paragraphs.Count >= 1 Then
                    With .TextRange.Paragraphs(1).Font
                        .Size = 8.5            ' R12-7-4: 8.5pt下限(a11y監査Med)
                        .Bold = -1
                        .Fill.ForeColor.RGB = modUI.UiColor("text")
                    End With
                End If
                .TextRange.ParagraphFormat.Alignment = 2
                .VerticalAnchor = 3
                .MarginLeft = 3: .MarginRight = 3: .MarginTop = 3: .MarginBottom = 3
            End With
            card.OnAction = "modHub.OnQuickAsk"
        End If
        Set card = Nothing
        Err.Clear
        On Error GoTo 0
    Next i
End Sub

' ----------------------------------------------------------------------------
' 受信箱(Hubの「お知らせ」1枚)。2026-07-31(R11-F1)に modHub から移設した。
'   modHub が30,000字上限まで残り3字となり、修正が1文字も入らない状態だった
'   (憲章§4-6)。分岐の判定材料(PendingUpdatesCached/PendingLabel/
'   ShareQueriesAllowed)がすべて本モジュールにあり、modHub 側からは
'   DrawInbox 1本しか呼ばれていないため、ここが自然な置き場所になる。
' ----------------------------------------------------------------------------

' 共有知のお知らせ。通知が無い共有機能は使われないので、件数と行き先を出す。
' 戻り値(2026-08-05 R18-5b): 描いた箱の下端Y。箱の高さは文面の長さで変わる
' ため、フッターのY決めに必要な「右カラムの実下端」はここしか知らない。
Public Function DrawInbox(ByVal ws As Worksheet, ByVal L As Double, _
                      ByVal W As Double, ByVal T As Double) As Double
    Dim qaN As Long, gapN As Long
    On Error Resume Next
    qaN = modInsight.PendingQACount()
    gapN = modInsight.GapCount()
    On Error GoTo 0

    ' 共有フォルダが未設定だと、この機能は丸ごと沈黙する。黙って何も起きない
    ' のがいちばん不親切なので、まずそこを案内する。
    Dim shareOk As Boolean
    On Error Resume Next
    shareOk = (LenB(Trim$(modConfig.GetString("nexus_share_path", ""))) > 0)
    On Error GoTo 0

    ' 2026-07-31(レビュー R8 F3): ここで直接 modChannel.PendingUpdates を
    ' 呼ぶと、購読中の部門ぶんの version.txt をその場で読みに行く。Hubの
    ' 初期描画は起動シーケンスの途中なので、遅い共有では起動が止まる。
    ' 起動中は問い合わせず(空が返る)、起動後の再描画で10分TTL付きで引く。
    Dim chPend As String
    On Error Resume Next
    chPend = modHubStat.PendingUpdatesCached()
    On Error GoTo 0

    ' 上限超過(R8 F13)。41件目以降の部門は静かに捨てられていたため、
    ' 「うちの部門にだけ正典が届かない」が誰にも調べられない不具合だった。
    Dim chOver As Long
    On Error Resume Next
    chOver = modChannel.OverflowCount()
    On Error GoTo 0

    ' R13 F6: 分岐の優先順位は【未設定 > 未接続 > 通常】。未設定の端末は
    ' そもそも共有へ繋ぎに行っていないので、ネットワーク/VPNの確認を促すのは
    ' 事実に反する。未設定の案内(shareOk=False)は必ずこの位置=最初に置く。
    Dim cap As String, act As String
    If Not shareOk Then
        ' R20-4c: 「configのnexus_share_path」という実行不能な案内をやめ、
        ' クリックでそのまま設定UI(modHub.OnShareHelp→modHelp.OnShareSetup)へ。
        cap = modEmj.Warn() & " 部内の共有フォルダが未設定です" & vbCr & _
              "クリックして設定(フォルダを選ぶだけ)。みんなが解決したQ&Aが自動で届くようになります"
        act = "modHub.OnShareHelp"
    ElseIf LenB(modChannel.ActiveChannel()) = 0 Then
        ' まだどの部門にもつないでいない。ここを案内しないと、
        ' 「聞いても答えが返ってこない」理由が利用者に分からない。
        cap = ChrW(&HD83D) & ChrW(&HDCDA) & " 部門の公式ナレッジをまだ読み込んでいません" & vbCr & _
              "押すと全部門をまとめて取り込みます。以後は分野を選ばずそのまま聞けます"
        act = "modKnowledge.OnChannels"
    ElseIf LenB(chPend) > 0 Then
        ' 正典の改定は最優先で知らせる。古い版のまま使い続けると、AIが
        ' 古い条文を根拠に答えるという最悪の事故になる。
        '
        ' 2026-07-28(レビュー H-13): 表示は「アクティブ部門名」を出していた
        ' ため、更新があるのは人事部なのに「【商品部】に更新があります」と
        ' 出て、押しても商品部は最新なので何も起きない、という
        ' 消えないバッジになっていた。保留リストの実体をそのまま出す。
        cap = ChrW(&HD83D) & ChrW(&HDCE1) & " 【" & modHubStat.PendingLabel(chPend) & _
              "】に更新があります" & vbCr & _
              "押して読み込み直してください。古い内容で回答しないために早めの更新を"
        act = "modHubStat.OnSyncPending"
    ElseIf chOver > 0 Then
        ' 2026-07-31(レビュー R8 F13): 部門数が上限を超えると、41件目以降は
        ' 一覧から静かに落ちる。落ちた部門の正典は誰にも届かないのに、
        ' 画面にもログにも何も出ないため「うちの部門だけ届かない」という
        ' 調べようのない不具合になっていた。件数を必ず見せる。
        cap = modEmj.Warn() & " 部門が多すぎて " & chOver & "部門を読み込めていません" & vbCr & _
              "このツールの管理担当者にご連絡ください(部門数の上限を超えています)"
        act = "modKnowledge.OnChannels"
    ElseIf modChannel.IsBudgetTight() Then
        ' R33H F6: 警告文からパーセントを外す(文言は modShareRule が単一情報源)。
        cap = modShareRule.BudgetWarnCaption(modChannel.ChunkLimit())
        act = "modKnowledge.OnChannels"
    ElseIf qaN > 0 Then
        cap = ChrW(&HD83C) & ChrW(&HDF81) & " みんなが解決したQ&A " & qaN & "件が届いています" & vbCr & _
              "押すと一覧が開きます。要るものだけ選んで本棚に入れられます"
        act = "modKnowledge.OnGoShared"
    ElseIf gapN > 0 Then
        cap = ChrW(&HD83D) & ChrW(&HDCA1) & " まだ答えを用意できていない質問が " & gapN & "件" & vbCr & _
              "押すと一覧が開きます。答えられる資料を登録すると部内に行き渡ります"
        act = "modKnowledge.OnGapBoard"
    ElseIf Not modHubStat.ShareQueriesAllowed() Then
        ' 起動中は共有フォルダへ問い合わせない(R8 F3)。「更新はありません」と
        ' 言い切ると嘘になり得るので、まだ見ていないことをそのまま書く。
        '
        ' 2026-07-31(R8b B4): この分岐は qaN/gapN/IsBudgetTight の【後ろ】に置く。
        ' それらは insight_inbox シートと本棚のチャンク数を見るだけで共有
        ' フォルダに一切触らない、起動直後でも正しく出せる通知である。
        ' 前に置くと、せっかく受信済みの「みんなが解決したQ&A N件」が
        ' 起動直後は必ず「更新はまだ確認していません」に塗り潰されてしまい、
        ' 共有知フライホイールの入口(利用者が新着に気付く唯一の場所)が
        ' 事実上ふさがる。共有I/Oを要する通知だけを後回しにするのが趣旨。
        cap = ChrW(&HD83D) & ChrW(&HDCE1) & " 部門の更新はまだ確認していません" & vbCr & _
              "押すと今すぐ確認します(起動を軽くするため、開いた直後は確認しません)"
        act = "modHub.OnCheckUpdates"
    ElseIf Not modShare.Reachable() Then
        ' R13-7b: 「自動で行き来しています」は共有フォルダに実際に届いている
        ' ときだけ言ってよい嘘のない文言。Reachable()はセッション1回の
        ' キャッシュ判定なので、ここまでの分岐(shareOk等)を通り抜けた
        ' 時点で既に判定済み=このElseIf自体が新しい網羅プローブを増やさない。
        cap = modEmj.Warn() & " 部内の共有フォルダに届いていません" & vbCr & _
              "ネットワークまたはVPN接続をご確認ください(このセッションは自動更新を見送っています)"
        act = "modHub.OnShareHelp"
    Else
        cap = ChrW(&HD83D) & ChrW(&HDD01) & " 部内の知恵は自動で行き来しています" & vbCr & _
              ChrW(&H2705) & "解決した を押すとその答えが、答えが無かった質問は課題として共有されます"
        act = ""
    End If

    On Error Resume Next
    ' 2026-07-31(発見事項3): capは部門名・件数を含む可変長で、固定44ptの
    ' 箱に収まらないと2行目が枠外へ送られて消える(A-4と同型)。1行目の
    ' おおよその折返し行数を文字数から見積もり、箱を可変高にする。
    Dim estLines As Long: estLines = 1 + Int(Len(cap) / 46)
    If estLines < 2 Then estLines = 2
    Dim boxH As Double: boxH = 20 + estLines * 13
    If boxH < 44 Then boxH = 44

    Dim box As Shape
    Set box = ws.Shapes.AddShape(5, L, T, W, boxH)
    If box Is Nothing Then
        DrawInbox = T
        Exit Function
    End If
    box.Name = "nx_hub_inbox"
    box.Adjustments(1) = 0.08
    box.Line.Visible = -1
    box.Line.Weight = 0.75
    box.Line.ForeColor.RGB = modUI.UiColor("border")
    box.Fill.ForeColor.RGB = modUI.UiColor("surface")
    modSkin.ApplyLightShadow box
    ' 段落で書式を分けるため区切りはvbCr(vbLfだとParagraphs(2)が範囲外)。
    With box.TextFrame2
        .WordWrap = -1
        .MarginLeft = 12: .MarginRight = 10: .MarginTop = 6: .MarginBottom = 4
        .TextRange.Text = cap
        .TextRange.Font.Size = 9
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
        .TextRange.Paragraphs(1).Font.Bold = -1
        If .TextRange.Paragraphs.Count >= 2 Then
            .TextRange.Paragraphs(2).Font.Size = 8.5   ' R12-7-4: 7.5pt→8.5pt(a11y監査Med)
            .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
        End If
        .VerticalAnchor = 3
    End With
    If LenB(act) > 0 Then box.OnAction = act
    On Error GoTo 0
    DrawInbox = T + boxH
End Function

' ----------------------------------------------------------------------------
' Hub最下部のフッター(2026-08-05 R18-5b・実機第5報⑨)
' ----------------------------------------------------------------------------
' 「© リスクコンサルティング支援部」を最下部に置き、押すと社内ポータルを開く。
' 置き場所を modHub ではなくここにした理由: modHub が27,599字(WARN 28,000まで
' 残り401字)で、R18-5aのカード化と同じラウンドに両方は入らないため(憲章§4-6。
' DrawInbox を R11-F1 でここへ移したのと同じ判断)。
'
' Y座標は【固定値にしない】。左カラム(バッジ帯)も右カラム(お知らせ)も
' 件数・文面の長さで下端が動くため、固定Yだと内容が増えた端末でだけ本文に
' 重なる(modDash の ChartNoteY()+270 決め打ちが実際にそうなった。R18-4で
' 撤去した跡地)。呼び出し元(modHub)が左右の実下端の Max を渡し、ここでは
' ScrollArea の下端(R18-3b)を超えないことだけを保証する ―― 境界の外にある
' Shapeはスクロールで到達できない=押せない。
Public Function DrawFooter(ByVal ws As Worksheet, ByVal L As Double, _
                           ByVal W As Double, ByVal topY As Double, _
                           Optional ByRef trueBottom As Double) As Double
    On Error Resume Next
    ' R21-S4: 上限は「境界の下端」1本にする。従来は HUB_BOUND(60行)の実測と
    ' 可視高の底上げを重ねがけしており、modViewport.BoundAddr が決める本物の
    ' 境界とズレていた(「窓高で底上げ」が3箇所にあった)。境界と同じ
    ' RowAtFloor(窓高) から出せば、フッターは必ず境界の内側=押せる位置に来る。
    Dim viewH As Double: viewH = modViewport.ViewportHeight()
    Dim limitRow As Long: limitRow = modViewport.RowAtFloor(ws, viewH, 60)
    Dim limitY As Double
    limitY = ws.Rows(limitRow).Top + ws.Rows(limitRow).Height
    If limitY <= 0 Then limitY = viewH

    ' R21H F2(a)(c): 戻り値(y+FOOTER_H)は当時クランプされ常にlimitY以下=
    ' FitsInViewが常にTrueになり、狭窓で右カラムが境界の外に残っても検知
    ' できなかった。trueBottomは【クランプ前】の内容下端を別途返す(呼び出し
    ' 元がBoundAddrへ渡す用)。クランプが実際に発火した=本当は窓に収まって
    ' いないので、FitsInViewを確実にFalseへ倒す値(viewHを超える)にする。
    Dim y As Double: y = topY
    Dim clamped As Boolean
    If limitY > FOOTER_H + 8 Then clamped = (y + FOOTER_H > limitY - 8)
    If y < 0 Then y = 0
    ' R32 W4-3【確定バグ・R31 W2-2の副作用】: クランプ時は viewH+1 だけを返して
    ' いた。FitsInViewをFalseへ倒す番兵のつもりが、値としては実内容より【上】で、
    ' 境界を決めるBoundAddrがそこで切る。Hubのバッジ帯(Merge)が境界外に出て
    ' Rows.Deleteに巻き込まれ結合が縮んでいた(実機 lastUsed=34>境界32)。
    ' 番兵は【下限】として掛け、実内容の下端は常に含める。
    trueBottom = topY + FOOTER_H
    If clamped Then
        If trueBottom < viewH + 1 Then trueBottom = viewH + 1
    End If

    ' R32 F12: クランプした狭窓では【描かない】(上で位置も押し上げない)。
    ' 帯全体が当たり判定(下記)なので、境界外へ出たバッジ帯と重なると
    ' 押したものと違うもの=社内ポータルが反応する。置けないなら出さない。
    Dim fs As Shape
    If Not clamped Then Set fs = ws.Shapes.AddShape(1, L, y, W, FOOTER_H)
    If fs Is Nothing Then
        DrawFooter = topY
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    fs.Name = "nx_hub_footer"
    fs.Line.Visible = 0
    ' R18H FA-8(A-M7): 塗り無し(Fill.Visible=0)のShapeは、Excelでは
    ' 【文字の上だけ】がクリック領域になる。8.5ptの1行を狙って押させるのは
    ' 憲章§3-1(押せるものは必ず反応する)に反するので、完全透明の塗りを
    ' 敷いて帯全体を当たり判定にする(見た目は塗り無しと区別が付かない)。
    fs.Fill.Visible = -1
    fs.Fill.Transparency = 1
    With fs.TextFrame2
        .WordWrap = 0
        .TextRange.Text = ChrW(&HA9) & " リスクコンサルティング支援部"
        .TextRange.Font.Size = 8.5      ' R12-7-4のa11y下限
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
    End With
    fs.OnAction = "modHubStat.OnFooterPortal"
    fs.AlternativeText = "社内ポータルを開く"
    DrawFooter = y + FOOTER_H
    Set fs = Nothing
    Err.Clear
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' OnFooterPortal - フッターを押したら社内ポータルをブラウザで開く。
' ----------------------------------------------------------------------------
' 3段構え。どこまで落ちても「利用者が自分で辿り着ける状態」で終わらせる:
'   1. ThisWorkbook.FollowHyperlink(既存3箇所と同じ作法。事前トースト必須=
'      Officeの「このハイパーリンクを開きますか」を無説明で見せない)
'   2. WScript.Shell.Run(端末ポリシー(EDR/AppLocker)で塞がれていることが
'      多く、両方失敗する前提で設計する。modWorkExcel.DoOpenWorkExcel と同型)
'   3. URLをクリップボードへ入れて案内(modPeek の作法)+ E0905 を err_log へ
' 取込中は通常どおりブロックする(FollowHyperlinkでローカルを開く
' modPeek.OnOpenSource が R15波1の自己点検で allowlist から外され通常ガードに
' なった前例に倣う。ポータルは急がない)。
Public Sub OnFooterPortal()
    If modUiLock.BlockIfIngesting() Then Exit Sub

    Dim opened As Boolean: opened = True
    On Error Resume Next
    ' R27波3-14: waitless(直後に出る確認画面の予告。待たせる文ではない)。
    modSkin.ShowToast "ブラウザで社内ポータルを開きます(確認画面が出たら[はい])。", "info", True
    ThisWorkbook.FollowHyperlink PORTAL_URL
    If Err.Number <> 0 Then opened = False
    Err.Clear
    On Error GoTo 0
    If opened Then Exit Sub

    Dim wsh As Object
    On Error Resume Next
    Set wsh = CreateObject("WScript.Shell")
    If Not wsh Is Nothing Then
        wsh.Run PORTAL_URL, 1, False
        If Err.Number = 0 Then opened = True
    End If
    Set wsh = Nothing
    Err.Clear
    On Error GoTo 0
    If opened Then Exit Sub

    On Error Resume Next
    modLog.LogError "E0905", "modHubStat.OnFooterPortal", modUtil.SafeLeft(PORTAL_URL, 200)
    On Error GoTo 0
    modClip.CopyOrGuideBox PORTAL_URL, modLog.FriendlyMessage("E0905"), _
        modLog.PORTAL_FAIL_LEAD
End Sub
