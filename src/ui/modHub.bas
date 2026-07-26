Attribute VB_Name = "modHub"
Option Explicit

' modHub - Hub画面(ホームシート)。プロフィール+統計+ナビゲーションを1画面に
' まとめ、チャット画面から情報を追い出して「ごちゃつき」を解消する入口。
'
' 設計:
'   ・シートはSH_HOME。Shapeは全て nx_hub_ 接頭辞(冪等な全削除に使う)。
'   ・Shape座標はセル幾何(Range.Left/Top/Width)から算出する。pt決め打ちだと
'     列幅・行高の変更やDPIでセル文字とShapeがズレるため(実機で繰り返した事故)。
'   ・配色は modUI.UiColor() を単一情報源にする(ハードコードしない)。
'   ・TextFrame2.Paragraphsは vbCr でしか分割されない。複数段落に書式を当てる
'     ときは必ず vbCr を使う(vbLfだとParagraphs(2)が範囲外エラーになる)。

Private Const HDR_H As Double = 40
Private Const CARD_H As Double = 68
Private Const GAUGE_H As Double = 8
Private Const NAV_H As Double = 46
Private Const NAV_GAP As Double = 8
Private Const CHIP_H As Double = 22

' EnsureHubLayout - Hub画面を構築(冪等)。activate:=Trueで画面遷移も行う。
Public Sub EnsureHubLayout(Optional ByVal activate As Boolean = False)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_HOME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False

    If ws.Visible <> -1 Then ws.Visible = -1
    ' 保護されたままだとCells.Clear等が1004で全滅する(実機再発の教訓)。
    On Error Resume Next
    ws.Unprotect
    On Error GoTo Fail

    ' 旧ホーム画面(modUIMainのbtn_/lbl_)と前回のHub Shapeを両方消す。
    ' nx_hub_だけ消すと旧ボタンが上に浮いたまま残る。
    RemoveHubShapes ws

    ws.Cells.Clear
    ws.Cells.Font.Name = "Yu Gothic UI"
    ws.Cells.Font.Size = 10
    ws.Cells.Interior.Color = modUI.UiColor("bg")

    ' 幾何を確定させてからShapeを置く(順序が逆だと座標がズレる)。
    ' 「セル感」を消すための列取り: 統計タイル2枚の間にD列の細い溝を入れ、
    ' 左ブロック(B:F)と右ブロック(H:K)の間にもG列の溝を置く。タイル同士が
    ' 隣接していると罫線の有無に関わらず表に見えてしまう。
    ws.Columns("A").ColumnWidth = 1.5
    ws.Columns("B:C").ColumnWidth = 13
    ws.Columns("D").ColumnWidth = 1.2
    ws.Columns("E:F").ColumnWidth = 13
    ws.Columns("G").ColumnWidth = 2.5
    ws.Columns("H:K").ColumnWidth = 13
    ws.Columns("L").ColumnWidth = 1.5
    ws.Rows("1:60").RowHeight = 15

    If activate Then
        On Error Resume Next
        ws.Activate
        On Error GoTo Fail
    End If

    ' ActiveWindow系はそのシートが実際に前面のときだけ触る(別シートの
    ' 表示設定を巻き添えで変えてしまうため)。
    On Error Resume Next
    If ThisWorkbook.ActiveSheet Is ws Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
    End If
    On Error GoTo Fail

    On Error Resume Next
    modTelemetry.TrackScreen "hub"
    On Error GoTo Fail

    DrawHeader ws
    DrawProfileCard ws
    DrawStatTiles ws
    DrawNavButtons ws
    DrawExtras ws
    DrawBadges ws

    On Error Resume Next
    modUI.FreezeShapePlacement ws
    On Error GoTo 0

    Application.ScreenUpdating = True
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    modLog.LogError "E0801", "modHub.EnsureHubLayout", Err.Description, Err.Number
End Sub

' ヘッダーバー(全幅) + 右肩のユーティリティアイコン
Private Sub DrawHeader(ByVal ws As Worksheet)
    Dim L As Double, W As Double
    L = ws.Range("A1").Left
    W = ws.Range("A1:L1").Width

    Dim hdr As Shape
    Set hdr = ws.Shapes.AddShape(5, L, 0, W, HDR_H)
    hdr.Name = "nx_hub_hdr"
    hdr.Line.Visible = 0
    hdr.Adjustments(1) = 0.02
    hdr.Fill.ForeColor.RGB = modUI.UiColor("sidebar")
    modSkin.ApplyHeaderDepth hdr          ' §9: 濃紺の2色グラデーション
    With hdr.TextFrame2
        .TextRange.Text = ChrW(&H26A1) & " Nexus Agent"
        .TextRange.Font.Size = 14
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .MarginLeft = 14
        .VerticalAnchor = 3
    End With

    Dim icons As Variant, acts As Variant, tips As Variant
    ' 🔄=画面を再描画。旧サイドバーにあった復旧用ボタンの移設先
    ' (ウィンドウ移動やAlt+Tab復帰で表示が崩れたときの1クリック復旧手段)。
    ' 📮=匿名の投書箱。実名だと率直な意見は出てこないので、名前を一切
    ' 記録しない経路を別に用意する(既存のご意見箱はメール=実名)。
    icons = Array(ChrW(&HD83C) & ChrW(&HDF10), ChrW(&HD83C) & ChrW(&HDF19), _
                  ChrW(&HD83D) & ChrW(&HDD04), ChrW(&HD83D) & ChrW(&HDCEE), _
                  ChrW(&H2753), ChrW(&HD83D) & ChrW(&HDEAA))
    acts = Array("modHub.OnLangCycle", "modHub.OnThemeToggle", _
                 "modHub.OnRedraw", "modHub.OnAnonFeedback", _
                 "modHub.OnHelp", "modHub.OnSaveAndExit")
    tips = Array("回答言語を切り替える", "配色を切り替える", "画面を描き直す", _
                 "匿名で感想・要望を送る", "ヘルプ・使い方", "保存して閉じる")

    Dim xRight As Double: xRight = L + W - 8
    Dim i As Long
    For i = 5 To 0 Step -1
        xRight = xRight - 30
        Dim btn As Shape
        Set btn = ws.Shapes.AddShape(9, xRight, (HDR_H - 26) / 2, 26, 26)
        btn.Name = "nx_hub_ic" & i
        btn.Line.Visible = 0
        btn.Fill.ForeColor.RGB = modUI.UiColor("sidebarActive")
        With btn.TextFrame2
            .TextRange.Text = CStr(icons(i))
            .TextRange.Font.Size = 10
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = 2
            .VerticalAnchor = 3
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        End With
        btn.OnAction = CStr(acts(i))
        On Error Resume Next
        btn.AlternativeText = CStr(tips(i))
        On Error GoTo 0
    Next i
End Sub

' プロフィールカード + EXPゲージ(左カラムB:Eの幾何に合わせる)
Private Sub DrawProfileCard(ByVal ws As Worksheet)
    Dim L As Double, W As Double, T As Double
    L = ws.Range("B3").Left
    W = ws.Range("B3:F3").Width
    T = HDR_H + 12

    Dim card As Shape
    Set card = ws.Shapes.AddShape(5, L, T, W, CARD_H)
    card.Name = "nx_hub_profile"
    card.Adjustments(1) = 0.06
    card.Line.Visible = 0
    card.Fill.ForeColor.RGB = modUI.UiColor("surface")
    modSkin.ApplyLightShadow card

    Dim nm As String, dept As String
    Dim lvl As Long, ex As Long
    On Error Resume Next
    nm = Trim$(modConfig.GetString("pack_author", ""))
    dept = Trim$(modConfig.GetString("user_department", ""))
    lvl = modStats.Level()
    ex = modStats.ExpTotal()
    On Error GoTo 0
    If LenB(nm) = 0 Or nm = "名称未設定" Then nm = "ゲスト ユーザー"
    If lvl < 1 Then lvl = 1

    ' Paragraphsで書式を分けるため区切りはvbCr(vbLfだと1段落のまま)。
    With card.TextFrame2
        .WordWrap = -1
        .MarginLeft = 14: .MarginTop = 8: .MarginRight = 10
        .TextRange.Text = nm & IIf(LenB(dept) > 0, "  (" & dept & ")", "") & vbCr & _
            "Lv." & lvl & "   EXP " & ex
        .TextRange.Font.Size = 11
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
        On Error Resume Next
        .TextRange.Paragraphs(1).Font.Bold = -1
        .TextRange.Paragraphs(1).Font.Size = 12
        .TextRange.Paragraphs(2).Font.Size = 9
        .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
        On Error GoTo 0
    End With

    DrawExpGauge ws, L + 12, T + CARD_H - 14, W - 24, lvl, ex
End Sub

Private Sub DrawExpGauge(ByVal ws As Worksheet, ByVal L As Double, ByVal T As Double, _
                         ByVal W As Double, ByVal lvl As Long, ByVal ex As Long)
    Dim bg As Shape
    Set bg = ws.Shapes.AddShape(5, L, T, W, GAUGE_H)
    bg.Name = "nx_hub_expbg"
    bg.Adjustments(1) = 0.5
    bg.Line.Visible = 0
    bg.Fill.ForeColor.RGB = modUI.UiColor("border")

    Dim divisor As Long
    On Error Resume Next
    divisor = modConfig.GetLong("exp_level_divisor", 100)
    On Error GoTo 0
    If divisor < 1 Then divisor = 100

    ' Lv = Int(sqrt(EXP/divisor)) + 1 の逆算(modStats.Levelと同じ曲線)。
    Dim curFloor As Double, nextFloor As Double
    curFloor = CDbl(lvl - 1) * CDbl(lvl - 1) * divisor
    nextFloor = CDbl(lvl) * CDbl(lvl) * divisor

    Dim progress As Double
    If nextFloor > curFloor Then progress = (CDbl(ex) - curFloor) / (nextFloor - curFloor)
    If progress < 0# Then progress = 0#
    If progress > 1# Then progress = 1#

    Dim fillW As Double: fillW = W * progress
    If fillW < 3 Then Exit Sub          ' 0%は描かない(1pxの謎バーを出さない)

    Dim fl As Shape
    Set fl = ws.Shapes.AddShape(5, L, T, fillW, GAUGE_H)
    fl.Name = "nx_hub_expfill"
    fl.Adjustments(1) = 0.5
    fl.Line.Visible = 0
    fl.Fill.ForeColor.RGB = modUI.UiColor("accent")
    modSkin.ApplyGradient fl, RGB(245, 158, 11), RGB(251, 191, 36)   ' §9: amber
End Sub

' 統計タイル8枚。実機要望(2026-07-26)「セル感を消したい」への対処で、
' セルのMerge+罫線から角丸Shape+影に作り替えた。セルの矩形は等間隔・直角・
' 罫線という「表」の記号そのもので、罫線を消しても格子として認識されるため。
' 座標は左ブロック(B:F)の実測幾何から2列×4段で割り付ける。
Private Sub DrawStatTiles(ByVal ws As Worksheet)
    Dim labels As Variant, vals As Variant
    labels = Array("質問した回数", ChrW(&HD83D) & ChrW(&HDFE2) & " 自己解決", _
                   ChrW(&H23F1) & " 取り戻した時間", ChrW(&HD83D) & ChrW(&HDCD6) & " 本棚の使用量", _
                   ChrW(&HD83C) & ChrW(&HDF0D) & " みんな(今日)", ChrW(&HD83C) & ChrW(&HDF0D) & " みんな(今月)", _
                   ChrW(&HD83D) & ChrW(&HDD25) & " 連続ログイン", ChrW(&HD83D) & ChrW(&HDCE6) & " パック共有")
    vals = Array(CStr(AskTotal()), CStr(SafeStat("selfsolve_total")), _
                 FmtMin(SafeSavedMinutes()), ChunkUsage(), _
                 OrgMin("d"), OrgMin("m"), _
                 SafeStat("streak_days") & "日", CStr(SafeStat("pack_export_total")))

    ' 左ブロックの幾何。D列(溝)を挟んで B:C と E:F の2枚並び。
    Dim colL As Double, colW As Double, gutter As Double
    colL = ws.Range("B1").Left
    colW = ws.Range("B1:C1").Width
    gutter = ws.Range("D1").Width

    Dim tileH As Double: tileH = 52
    Dim gapY As Double: gapY = 8
    Dim topY As Double: topY = HDR_H + 12 + CARD_H + 18   ' プロフィールカードの下

    Dim i As Long
    For i = 0 To 7
        On Error Resume Next
        Dim x As Double, y As Double
        x = colL + (i Mod 2) * (colW + gutter)
        y = topY + (i \ 2) * (tileH + gapY)

        Dim tile As Shape
        Set tile = ws.Shapes.AddShape(5, x, y, colW, tileH)
        If Err.Number = 0 And Not tile Is Nothing Then
            tile.Name = "nx_hub_tile" & i
            tile.Adjustments(1) = 0.12
            tile.Line.Visible = -1
            tile.Line.Weight = 0.75
            tile.Line.ForeColor.RGB = modUI.UiColor("border")
            tile.Fill.ForeColor.RGB = modUI.UiColor("surface")
            modSkin.ApplyLightShadow tile
            ' 段落で書式を分けるため区切りはvbCr(vbLfだとParagraphs(2)が範囲外)。
            With tile.TextFrame2
                .WordWrap = -1
                .MarginLeft = 12: .MarginRight = 8: .MarginTop = 7: .MarginBottom = 5
                .TextRange.Text = CStr(labels(i)) & vbCr & CStr(vals(i))
                .TextRange.Font.Size = 8
                On Error Resume Next
                .TextRange.Paragraphs(1).Font.Size = 8
                .TextRange.Paragraphs(1).Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
                .TextRange.Paragraphs(2).Font.Size = 16
                .TextRange.Paragraphs(2).Font.Bold = -1
                .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = modUI.UiColor("primary")
                On Error GoTo 0
            End With
            ' みんなの節約(i=4,5)はクリックで部内ランキングを出す
            ' (旧サイドバーウィジェットのクリック機能の移設先)。
            If i = 4 Or i = 5 Then
                tile.OnAction = "modBoard.OnWidgetClick"
                tile.AlternativeText = "クリックで部内の節約ランキングを表示"
            End If
        End If
        Set tile = Nothing
        Err.Clear
        On Error GoTo 0
    Next i
End Sub

' 統計タイル群の下端(バッジ等をその下に置くために使う)。
Private Function StatTilesBottom() As Double
    StatTilesBottom = HDR_H + 12 + CARD_H + 18 + 4 * (52 + 8)
End Function

' ナビボタン4枚(右カラムG:Jの幾何に合わせる)
Private Sub DrawNavButtons(ByVal ws As Worksheet)
    Dim L As Double, W As Double, T As Double
    L = ws.Range("H3").Left
    W = ws.Range("H3:K3").Width
    T = HDR_H + 12

    Dim caps As Variant, acts As Variant, descs As Variant
    caps = Array(ChrW(&HD83D) & ChrW(&HDCAC) & " チャットで質問する", _
                 ChrW(&HD83D) & ChrW(&HDCDA) & " ナレッジ倉庫", _
                 ChrW(&HD83D) & ChrW(&HDCD6) & " マイ本棚", _
                 ChrW(&HD83D) & ChrW(&HDCE6) & " パック共有(P2P)", _
                 ChrW(&HD83D) & ChrW(&HDCCA) & " ダッシュボード")
    descs = Array("本棚の資料からAIが出典付きで回答", "資料の登録・検索・フォルダ同期", _
                  "取り込んだ資料の一覧と状態", "部内でナレッジを配る・受け取る", _
                  "バッジ・EXP・ナレッジ地図")
    acts = Array("modHub.OnGoChat", "modHub.OnGoVault", "modHub.OnGoShelf", _
                 "modHub.OnGoPack", "modHub.OnGoDash")

    Dim i As Long
    For i = 0 To 4
        Dim navTop As Double: navTop = T + i * (NAV_H + NAV_GAP)
        Dim btn As Shape
        Set btn = ws.Shapes.AddShape(5, L, navTop, W, NAV_H)
        btn.Name = "nx_hub_nav" & i
        btn.Adjustments(1) = 0.08
        btn.Line.Visible = -1
        btn.Line.Weight = 0.75
        btn.Line.ForeColor.RGB = modUI.UiColor("border")
        btn.Fill.ForeColor.RGB = modUI.UiColor("surface")
        modSkin.ApplyLightShadow btn
        With btn.TextFrame2
            .WordWrap = -1
            .MarginLeft = 14: .MarginTop = 5: .MarginRight = 8
            .TextRange.Text = CStr(caps(i)) & vbCr & CStr(descs(i))
            .TextRange.Font.Size = 10.5
            .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
            On Error Resume Next
            .TextRange.Paragraphs(1).Font.Bold = -1
            .TextRange.Paragraphs(2).Font.Size = 8
            .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
            On Error GoTo 0
            .VerticalAnchor = 3
        End With
        btn.OnAction = CStr(acts(i))
    Next i
End Sub

' テンプレチップ3つ + 今日のワンポイント(ナビボタンの直下に続けて置く)
Private Sub DrawExtras(ByVal ws As Worksheet)
    Dim L As Double, W As Double, T As Double
    L = ws.Range("H3").Left
    W = ws.Range("H3:K3").Width
    T = HDR_H + 12 + 5 * (NAV_H + NAV_GAP) + 10

    Dim lbl As Shape
    Set lbl = ws.Shapes.AddShape(1, L, T, W, 16)
    lbl.Name = "nx_hub_qalbl"
    lbl.Line.Visible = 0
    lbl.Fill.Visible = 0
    With lbl.TextFrame2
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCA1) & " こんなふうに聞いてみよう"
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
        .MarginLeft = 2
        .VerticalAnchor = 3
    End With

    Dim chips As Variant
    chips = Array("改定ポイントを教えて", "用語をやさしく解説", "手続きの流れを知りたい")
    Dim chipW As Double: chipW = (W - 12) / 3
    Dim i As Long
    For i = 0 To 2
        Dim chip As Shape
        Set chip = ws.Shapes.AddShape(5, L + i * (chipW + 6), T + 20, chipW, CHIP_H)
        chip.Name = "nx_hub_qa" & i
        chip.Adjustments(1) = 0.4
        chip.Line.Visible = 0
        chip.Fill.ForeColor.RGB = modUI.UiColor("bg")
        With chip.TextFrame2
            .WordWrap = -1
            .TextRange.Text = CStr(chips(i))
            .TextRange.Font.Size = 7.5
            .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
            .TextRange.ParagraphFormat.Alignment = 2
            .VerticalAnchor = 3
            .MarginLeft = 3: .MarginRight = 3: .MarginTop = 0: .MarginBottom = 0
        End With
        chip.OnAction = "modHub.OnQuickAsk"
    Next i

    Dim gacha As Shape
    Set gacha = ws.Shapes.AddShape(5, L, T + 20 + CHIP_H + 8, W, 26)
    gacha.Name = "nx_hub_gacha"
    gacha.Adjustments(1) = 0.1
    gacha.Fill.Visible = 0
    gacha.Line.Visible = -1
    gacha.Line.Weight = 1.25
    gacha.Line.ForeColor.RGB = modUI.UiColor("accent")
    On Error Resume Next
    gacha.Line.DashStyle = 4       ' msoLineDash(未対応環境では実線のまま)
    On Error GoTo 0
    With gacha.TextFrame2
        .TextRange.Text = ChrW(&HD83C) & ChrW(&HDFB2) & " 今日のワンポイントを引く"
        .TextRange.Font.Size = 9.5
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("accent")
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    gacha.OnAction = "modHub.OnGacha"

    DrawInbox ws, L, W, T + 20 + CHIP_H + 8 + 26 + 12
End Sub

' 共有知のお知らせ。届いた「解決済みQ&A」と「みんなの困りごと」の件数を出し、
' 1クリックでそれぞれの行き先へ送る。ここが無いと、せっかく届いた知見が
' 誰にも気づかれないまま眠る(通知が無い共有機能は使われない)。
Private Sub DrawInbox(ByVal ws As Worksheet, ByVal L As Double, _
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

    Dim chPend As String
    On Error Resume Next
    chPend = modChannel.PendingUpdates()
    On Error GoTo 0

    Dim cap As String, act As String
    If Not shareOk Then
        cap = ChrW(&H26A0) & " 部内の共有フォルダが未設定です" & vbCr & _
              "設定すると、みんなが解決したQ&Aが自動で届くようになります(config の nexus_share_path)"
        act = "modHub.OnShareHelp"
    ElseIf LenB(chPend) > 0 Then
        ' 正典の改定は最優先で知らせる。古い版のまま使い続けると、AIが
        ' 古い条文を根拠に答えるという最悪の事故になる。
        cap = ChrW(&HD83D) & ChrW(&HDCE1) & " 部門チャンネルに更新があります(" & _
              Replace(chPend, "|", " / ") & ")" & vbCr & _
              "押すと最新版に入れ替えます。古い内容で回答しないために早めの更新を"
        act = "modKnowledge.OnChannels"
    ElseIf modChannel.IsBudgetTight() Then
        cap = ChrW(&H26A0) & " 本棚の使用量が " & modChannel.ChunkUsagePercent() & "% です" & vbCr & _
              "押すと部門チャンネルの購読を見直せます(使っていないものを外すと空きます)"
        act = "modKnowledge.OnChannels"
    ElseIf qaN > 0 Then
        cap = ChrW(&HD83C) & ChrW(&HDF81) & " みんなが解決したQ&A " & qaN & "件が届いています" & vbCr & _
              "押すと一覧が開きます。要るものだけ選んで本棚に入れられます"
        act = "modKnowledge.OnGoShared"
    ElseIf gapN > 0 Then
        cap = ChrW(&HD83D) & ChrW(&HDCA1) & " まだ答えを用意できていない質問が " & gapN & "件" & vbCr & _
              "押すと一覧が開きます。答えられる資料を登録すると部内に行き渡ります"
        act = "modKnowledge.OnGapBoard"
    Else
        cap = ChrW(&HD83D) & ChrW(&HDD01) & " 部内の知恵は自動で行き来しています" & vbCr & _
              ChrW(&H2705) & "解決した を押すとその答えが、答えが無かった質問は課題として共有されます"
        act = ""
    End If

    On Error Resume Next
    Dim box As Shape
    Set box = ws.Shapes.AddShape(5, L, T, W, 44)
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
        .TextRange.Paragraphs(2).Font.Size = 7.5
        .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
        .VerticalAnchor = 3
    End With
    If LenB(act) > 0 Then box.OnAction = act
    On Error GoTo 0
End Sub

' バッジ(セル。獲得済みは🏅、未獲得は🔒)。統計タイルはShapeに変えたので
' この帯だけが左ブロックのセル表示になる。
Private Sub DrawBadges(ByVal ws As Worksheet)
    ' タイルの下端が入る行を実測で探す(行高15pt固定なので割り算で足りる)。
    Dim r As Long
    r = CLng(StatTilesBottom() / 15) + 2
    If r < 10 Then r = 10

    With ws.Range("B" & r & ":F" & r)
        .Merge
        .Value = ChrW(&HD83C) & ChrW(&HDFC5) & " バッジ"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = modUI.UiColor("muted")
        .VerticalAlignment = -4108
    End With

    Dim ids As Variant, titles As Variant
    ids = Split("first_ingest,shelf10,shelf30,first_pack_out,first_pack_in,solve10,solve50,streak7", ",")
    titles = Split("初取込,本棚10冊,本棚30冊,初パック出力,初パック取込,自己解決10,自己解決50,7日連続", ",")

    Dim sb As String
    Dim i As Long
    For i = 0 To UBound(ids)
        Dim mark As String
        If SafeStat("badge:" & CStr(ids(i))) > 0 Then
            mark = ChrW(&HD83C) & ChrW(&HDFC5)
        Else
            mark = ChrW(&HD83D) & ChrW(&HDD12)
        End If
        sb = sb & mark & " " & CStr(titles(i)) & "   "
    Next i

    With ws.Range("B" & (r + 1) & ":F" & (r + 3))
        .Merge
        .WrapText = True
        .Value = sb
        .Font.Size = 8.5
        .Font.Color = modUI.UiColor("text")
        .VerticalAlignment = -4160
    End With
End Sub

' ---- ボタンハンドラ ----

Public Sub OnGoChat()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUI.GoToNexus "modHub.OnGoChat"
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnGoVault()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modVault.ShowVaultGallery
    On Error GoTo 0
    modUiLock.Leave
End Sub

' マイ本棚=ナレッジ画面のテーブルモード。EnsureLayoutを先に通して共通クロム
' (ヘッダー+モード切替ピル+ツールバー)を確実に描いてから遷移する。
' modKnowledge.OnGoTableを呼ばないのは、あちらもmodUiLockを取るため
' ここで取得済みのロックと衝突して何も起きなくなるから。
Public Sub OnGoShelf()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUIShelf.EnsureLayout
    modUI.GoToNativeSheet modAppDef.SH_SHELF, "modHub.OnGoShelf"
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnGoDash()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modDash.ShowDashboard
    On Error GoTo 0
    modUiLock.Leave
End Sub

' パック共有: ナレッジ倉庫(パックの出力/取込ボタンがある画面)へ送る。
Public Sub OnGoPack()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modVault.ShowVaultGallery
    modSkin.ShowToast "画面下の「パック出力」「パック取込」で部内共有ができます。", "info"
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnQuickAsk()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    Dim tpl As String
    Select Case CStr(Application.Caller)
        Case "nx_hub_qa0"
            tpl = "【知りたい改定】: (資料名や年度)" & vbLf & "【気になる点】: (例: 保険料への影響)"
        Case "nx_hub_qa1"
            tpl = "【わからない用語】: (ここに記入)" & vbLf & "【どこまで知りたい】: (例: お客様に説明できるレベル)"
        Case "nx_hub_qa2"
            tpl = "【手続き名】: (ここに記入)" & vbLf & "【知りたい結論】: (例: 必要書類と所要日数)"
        Case Else
            GoTo Done
    End Select

    modUI.GoToNexus "modHub.OnQuickAsk"
    On Error Resume Next
    ThisWorkbook.Names("nx_input").RefersToRange.Value = tpl
    modSkin.ShowToast "(ここに記入)を埋めて送信してください。", "info"
    On Error GoTo Done
Done:
    modUiLock.Leave
End Sub

' 今日のワンポイントは既存のチャットバブル実装を再利用する
' (専用UserFormは実行時VBAプロジェクト書き換えが必要で実機リスクが高いため採らない)。
Public Sub OnGacha()
    On Error Resume Next
    modUI.GoToNexus "modHub.OnGacha"
    On Error GoTo 0
    modApp.OnGacha
End Sub

Public Sub OnLangCycle()
    On Error Resume Next
    modApp.OnLangCycle
    modSkin.ShowToast "回答言語: " & modConfig.GetString("answer_language", "日本語"), "success"
    On Error GoTo 0
End Sub

Public Sub OnThemeToggle()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modSkin.CycleSkin
    EnsureHubLayout
    On Error GoTo 0
    modUiLock.Leave
End Sub

' modHelp.OnHelpClick自身がmodUiLockを取る。modUiLockは非再入なので、
' ここで先に取ると内側のEnterがFalseになりヘルプが一切開かなくなる
' (Hub移植時に埋めてしまった不具合)。素通しにする。
' ヘルプカードはチャット画面(Nexus)に描かれるため、先にそちらへ遷移する。
Public Sub OnHelp()
    On Error Resume Next
    modUI.GoToNexus "modHub.OnHelp"
    On Error GoTo 0
    modHelp.OnHelpClick
End Sub

Public Sub OnSaveAndExit()
    modApp.OnSaveAndExit
End Sub

' 🔄 画面を再描画。ウィンドウのリサイズ・Alt+Tab復帰・マルチモニタ間の移動で
' Shapeがゴースト化/ズレたときの1クリック復旧手段(旧サイドバーから移設)。
' 会話は消さない。modApp.OnRefreshUIはロックを取るのでここでは取らない。
' 共有フォルダの設定手順を案内する。管理者がconfigに1回入れるだけで済むが、
' その1回が分からないまま放置されるのを防ぐ。
Public Sub OnShareHelp()
    MsgBox "部内で知恵を共有するには、共有フォルダを1回だけ設定します。" & vbCrLf & vbCrLf & _
        "【設定するもの】" & vbCrLf & _
        "  config シートの nexus_share_path に、部内の誰もが読み書きできる" & vbCrLf & _
        "  共有サーバー上のフォルダパスを入れてください。" & vbCrLf & _
        "  例) \\サーバー名\部門共有\Nexus_Share\" & vbCrLf & vbCrLf & _
        "【そのあと何が起きるか】" & vbCrLf & _
        "  ・必要なサブフォルダはこのアプリが自動で作ります" & vbCrLf & _
        "  ・誰かが「解決した」を押すと、その質問と答えがそこへ置かれます" & vbCrLf & _
        "  ・次に各自がこのファイルを開いたとき、自動で受け取ります" & vbCrLf & _
        "  ・受け取っただけでは本棚に入りません。必要なものを選んで取り込みます" & vbCrLf & vbCrLf & _
        "【注意】" & vbCrLf & _
        "  同じパスを部内の全員が設定してはじめて共有が成立します。" & vbCrLf & _
        "  配布用ファイルにあらかじめ入れておくのがいちばん確実です。", _
        vbInformation, modAppDef.APP_NAME
End Sub

' 匿名フィードバック。名前も所属も記録しないことを画面で明示する。
' 「匿名です」と書いてあるかどうかで、集まる本音の量が変わる。
' オーナー向け: 利用状況と匿名フィードバックの一覧。
' 発行キーを持つ端末(=運営側)にだけ意味がある機能なので、キー確認を挟む。
Public Sub OnOwnerReport()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done

    If Not modPublish.VerifyKey(InputBox( _
            "運営用の画面です。発行キーを入力してください。", _
            modAppDef.APP_NAME & " - 利用状況")) Then GoTo Done

    Dim body As String, fb As String
    On Error Resume Next
    body = modTelemetry.SummaryText()
    fb = modTelemetry.FeedbackText()
    On Error GoTo Done

    modUiLock.Leave
    MsgBox body & vbCrLf & "■ 匿名で届いた声(新しい順)" & vbCrLf & fb, _
           vbInformation, modAppDef.APP_NAME & " - 利用状況"
    Exit Sub
Done:
    modUiLock.Leave
End Sub

Public Sub OnAnonFeedback()
    Dim fb As String
    fb = InputBox( _
        "このツールへの感想・要望・不満を、匿名で送れます。" & vbCrLf & vbCrLf & _
        "  ・お名前も所属も記録しません(あとから誰が書いたかは分かりません)" & vbCrLf & _
        "  ・辛口で構いません。使いにくい点ほど価値があります" & vbCrLf & vbCrLf & _
        "例) 検索が遅い / ボタンの意味が分からない / この機能が欲しい", _
        modAppDef.APP_NAME & " - 匿名の投書箱")
    If LenB(Trim$(fb)) = 0 Then Exit Sub

    Dim ok As Boolean
    On Error Resume Next
    ok = modTelemetry.SendAnonymousFeedback(fb)
    On Error GoTo 0

    If ok Then
        On Error Resume Next
        modStats.AddExp "feedback"
        On Error GoTo 0
        MsgBox "ありがとうございます。匿名で届きました。" & vbCrLf & _
               "いただいた声は改善に使わせていただきます。(EXP +5)", _
               vbInformation, modAppDef.APP_NAME
    Else
        ' 共有フォルダに書けない環境ではメール経路へ逃がす(黙って捨てない)。
        On Error Resume Next
        modClip.SetClipboardText fb
        ThisWorkbook.FollowHyperlink _
            "mailto:m-kojima@aioinissaydowa.co.jp?subject=Nexus%20Agent%20feedback"
        On Error GoTo 0
        MsgBox "共有フォルダへ送れなかったため、メールの下書きを開きました。" & vbCrLf & _
               "本文はクリップボードに入っています(Ctrl+V で貼り付けてください)。", _
               vbInformation, modAppDef.APP_NAME
    End If
End Sub

Public Sub OnRedraw()
    On Error Resume Next
    modApp.OnRefreshUI
    On Error GoTo 0
End Sub

' ---- 内部ヘルパー ----

' 前回のHub Shapeに加え、旧ホーム画面(modUIMain)のbtn_/lbl_も消す。
' nx_hub_だけ消すと旧ボタンが新レイアウトの上に浮いたまま残る。
Private Sub RemoveHubShapes(ByVal ws As Worksheet)
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

Private Function SafeStat(ByVal key As String) As Long
    On Error Resume Next
    SafeStat = modStats.GetStat(key)
    On Error GoTo 0
End Function

' 質問回数は quick/deep 別カウンタの合算(単一のquestion_totalキーは存在しない)。
Private Function AskTotal() As Long
    AskTotal = SafeStat("ask_quick_total") + SafeStat("ask_deep_total")
End Function

Private Function SafeSavedMinutes() As Long
    On Error Resume Next
    SafeSavedMinutes = modStats.SavedMinutesEstimate()
    On Error GoTo 0
End Function

Private Function SafeChunks() As Long
    On Error Resume Next
    SafeChunks = modShelf.TotalChunks()
    On Error GoTo 0
End Function

' 本棚の使用量。実数だけ出しても上限が分からないので割合で見せる。
' 部門チャンネルを増やすほど埋まるので、増やしてよいかの判断材料になる。
Private Function ChunkUsage() As String
    Dim pct As Long
    On Error Resume Next
    pct = modChannel.ChunkUsagePercent()
    On Error GoTo 0
    ChunkUsage = pct & "%"
End Function

Private Function FmtMin(ByVal minutes As Long) As String
    If minutes < 60 Then
        FmtMin = minutes & "分"
    Else
        FmtMin = (minutes \ 60) & "時間"
        If (minutes Mod 60) > 0 Then FmtMin = FmtMin & (minutes Mod 60) & "分"
    End If
End Function

Private Function OrgMin(ByVal period As String) As String
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
