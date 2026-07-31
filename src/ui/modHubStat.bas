Attribute VB_Name = "modHubStat"
Option Explicit

' 統計タイルの幾何(2026-07-31 R7 A-4で modHub から移設)。
' 高さ・段数はHubのバッジ帯の位置決め(modHub.StatTilesBottom)も使うので、
' 数字の出どころをここ1箇所にする。
Private Const TILE_H As Double = 52
Private Const TILE_GAP_Y As Double = 8
Private Const TILE_ROWS As Long = 4
Private Const TILE_COUNT As Long = 8

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
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUIMain.ShowProgress "部門の更新を取り込んでいます…"
    Dim failN As Long
    Dim got As Long: got = modChannel.SyncSubscribed(failN)
    modUIMain.HideProgress
    InvalidatePending
    modUiLock.Leave
    modSkin.ShowToast IIf(got > 0, got & "件のまとまりを取り込みました。", "更新はありませんでした。") & _
        IIf(failN > 0, "(" & failN & "部門は取り込めませんでした)", ""), "info"
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

' 質問回数は quick/deep 別カウンタの合算(単一のquestion_totalキーは存在しない)。
Public Function AskTotal() As Long
    AskTotal = SafeStat("ask_quick_total") + SafeStat("ask_deep_total")
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
        Case 2, 4, 5   ' 節約できた時間 / みんな(今日) / みんな(今月)
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
    On Error GoTo 0
    OrgMin = FmtMin(v)
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
                   ChrW(&H23F1) & " 節約できた時間", ChrW(&HD83D) & ChrW(&HDCD6) & " 本棚の使用量", _
                   ChrW(&HD83C) & ChrW(&HDF0D) & " みんな(今日)", ChrW(&HD83C) & ChrW(&HDF0D) & " みんな(今月)", _
                   ChrW(&HD83D) & ChrW(&HDD25) & " 連続ログイン", ChrW(&HD83D) & ChrW(&HDCE6) & " パック共有")
    ' 実機報告(2026-07-27)「一部のタイルが真っ白」対策: Array()内で直接
    ' 関数を呼ぶと1つの失敗が空文字になる。1つずつ受けて必ず値を入れる。
    Dim vAsk As String, vSolve As String, vSaved As String, vUse As String
    Dim vOrgD As String, vOrgM As String, vStreak As String, vPack As String
    vAsk = NumText(AskTotal())
    vSolve = NumText(SafeStat("selfsolve_total"))
    vSaved = FmtMin(SafeSavedMinutes())
    vUse = ChunkUsage()
    vOrgD = OrgMin("d")
    vOrgM = OrgMin("m")
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
        y = topY + (i \ 2) * (TILE_H + TILE_GAP_Y)

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
    Set tile = ws.Shapes.AddShape(5, x, y, tileW, TILE_H)
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
' 受信箱(Hubの「お知らせ」1枚)。2026-07-31(R11-F1)に modHub から移設した。
'   modHub が30,000字上限まで残り3字となり、修正が1文字も入らない状態だった
'   (憲章§4-6)。分岐の判定材料(PendingUpdatesCached/PendingLabel/
'   ShareQueriesAllowed)がすべて本モジュールにあり、modHub 側からは
'   DrawInbox 1本しか呼ばれていないため、ここが自然な置き場所になる。
' ----------------------------------------------------------------------------

' 共有知のお知らせ。通知が無い共有機能は使われないので、件数と行き先を出す。
Public Sub DrawInbox(ByVal ws As Worksheet, ByVal L As Double, _
                      ByVal W As Double, ByVal T As Double)
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

    Dim cap As String, act As String
    If Not shareOk Then
        cap = ChrW(&H26A0) & " 部内の共有フォルダが未設定です" & vbCr & _
              "設定すると、みんなが解決したQ&Aが自動で届くようになります(config の nexus_share_path)"
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
        cap = ChrW(&H26A0) & " 部門が多すぎて " & chOver & "部門を読み込めていません" & vbCr & _
              "このツールの管理担当者にご連絡ください(部門数の上限を超えています)"
        act = "modKnowledge.OnChannels"
    ElseIf modChannel.IsBudgetTight() Then
        cap = ChrW(&H26A0) & " 本棚の使用量が " & modChannel.ChunkUsagePercent() & "% です" & vbCr & _
              "使っていない資料を減らすと空きます(マイ本棚から削除できます)"
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
    If box Is Nothing Then Exit Sub
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
            .TextRange.Paragraphs(2).Font.Size = 7.5
            .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
        End If
        .VerticalAnchor = 3
    End With
    If LenB(act) > 0 Then box.OnAction = act
    On Error GoTo 0
End Sub
