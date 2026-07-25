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
    ws.Columns("A").ColumnWidth = 1.5
    ws.Columns("B:E").ColumnWidth = 13
    ws.Columns("F").ColumnWidth = 2
    ws.Columns("G:J").ColumnWidth = 13
    ws.Columns("K").ColumnWidth = 1.5
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
    W = ws.Range("A1:K1").Width

    Dim hdr As Shape
    Set hdr = ws.Shapes.AddShape(5, L, 0, W, HDR_H)
    hdr.Name = "nx_hub_hdr"
    hdr.Line.Visible = 0
    hdr.Adjustments(1) = 0.02
    hdr.Fill.ForeColor.RGB = modUI.UiColor("sidebar")
    With hdr.TextFrame2
        .TextRange.Text = ChrW(&H26A1) & " Nexus Agent"
        .TextRange.Font.Size = 14
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .MarginLeft = 14
        .VerticalAnchor = 3
    End With

    Dim icons As Variant, acts As Variant, tips As Variant
    icons = Array(ChrW(&HD83C) & ChrW(&HDF10), ChrW(&HD83C) & ChrW(&HDF19), _
                  ChrW(&H2753), ChrW(&HD83D) & ChrW(&HDEAA))
    acts = Array("modHub.OnLangCycle", "modHub.OnThemeToggle", _
                 "modHub.OnHelp", "modHub.OnSaveAndExit")
    tips = Array("回答言語を切り替える", "配色を切り替える", "ヘルプ", "保存して閉じる")

    Dim xRight As Double: xRight = L + W - 8
    Dim i As Long
    For i = 3 To 0 Step -1
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
    W = ws.Range("B3:E3").Width
    T = HDR_H + 12

    Dim card As Shape
    Set card = ws.Shapes.AddShape(5, L, T, W, CARD_H)
    card.Name = "nx_hub_profile"
    card.Adjustments(1) = 0.06
    card.Line.Visible = 0
    card.Fill.ForeColor.RGB = modUI.UiColor("surface")
    SafeShadow card

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
End Sub

' 統計タイル8枚(セル。Merge+塗りで軽く保つ)
Private Sub DrawStatTiles(ByVal ws As Worksheet)
    Dim startRow As Long: startRow = 9
    Dim labels As Variant, vals As Variant
    labels = Array("質問した回数", ChrW(&HD83D) & ChrW(&HDFE2) & " 自己解決", _
                   ChrW(&H23F1) & " 取り戻した時間", ChrW(&HD83D) & ChrW(&HDCD6) & " 蔵書チャンク", _
                   ChrW(&HD83C) & ChrW(&HDF0D) & " みんな(今日)", ChrW(&HD83C) & ChrW(&HDF0D) & " みんな(今月)", _
                   ChrW(&HD83D) & ChrW(&HDD25) & " 連続ログイン", ChrW(&HD83D) & ChrW(&HDCE6) & " パック共有")
    vals = Array(CStr(AskTotal()), CStr(SafeStat("selfsolve_total")), _
                 FmtMin(SafeSavedMinutes()), CStr(SafeChunks()), _
                 OrgMin("d"), OrgMin("m"), _
                 SafeStat("streak_days") & "日", CStr(SafeStat("pack_export_total")))

    Dim i As Long
    For i = 0 To 7
        Dim r As Long: r = startRow + (i \ 2) * 3
        Dim c1 As String, c2 As String
        If (i Mod 2) = 0 Then
            c1 = "B": c2 = "C"
        Else
            c1 = "D": c2 = "E"
        End If

        With ws.Range(c1 & r & ":" & c2 & r)
            .Merge
            .Value = CStr(labels(i))
            .Font.Size = 8
            .Font.Color = modUI.UiColor("muted")
            .HorizontalAlignment = -4108
            .VerticalAlignment = -4108
        End With
        With ws.Range(c1 & (r + 1) & ":" & c2 & (r + 1))
            .Merge
            .Value = CStr(vals(i))
            .Font.Size = 15
            .Font.Bold = True
            .Font.Color = modUI.UiColor("primary")
            .HorizontalAlignment = -4108
            .VerticalAlignment = -4108
            .Interior.Color = modUI.UiColor("surface")
            .BorderAround LineStyle:=1, Weight:=2, Color:=modUI.UiColor("border")
        End With
        ws.Rows(r + 1).RowHeight = 26
    Next i
End Sub

' ナビボタン4枚(右カラムG:Jの幾何に合わせる)
Private Sub DrawNavButtons(ByVal ws As Worksheet)
    Dim L As Double, W As Double, T As Double
    L = ws.Range("G3").Left
    W = ws.Range("G3:J3").Width
    T = HDR_H + 12

    Dim caps As Variant, acts As Variant, descs As Variant
    caps = Array(ChrW(&HD83D) & ChrW(&HDCAC) & " チャットで質問する", _
                 ChrW(&HD83D) & ChrW(&HDCDA) & " ナレッジ倉庫", _
                 ChrW(&HD83D) & ChrW(&HDCD6) & " マイ本棚", _
                 ChrW(&HD83D) & ChrW(&HDCCA) & " ダッシュボード")
    descs = Array("本棚の資料からAIが出典付きで回答", "資料の登録・検索・フォルダ同期", _
                  "取り込んだ資料の一覧と状態", "バッジ・EXP・ナレッジ地図")
    acts = Array("modHub.OnGoChat", "modHub.OnGoVault", "modHub.OnGoShelf", "modHub.OnGoDash")

    Dim i As Long
    For i = 0 To 3
        Dim navTop As Double: navTop = T + i * (NAV_H + NAV_GAP)
        Dim btn As Shape
        Set btn = ws.Shapes.AddShape(5, L, navTop, W, NAV_H)
        btn.Name = "nx_hub_nav" & i
        btn.Adjustments(1) = 0.08
        btn.Line.Visible = -1
        btn.Line.Weight = 0.75
        btn.Line.ForeColor.RGB = modUI.UiColor("border")
        btn.Fill.ForeColor.RGB = modUI.UiColor("surface")
        SafeShadow btn
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
    L = ws.Range("G3").Left
    W = ws.Range("G3:J3").Width
    T = HDR_H + 12 + 4 * (NAV_H + NAV_GAP) + 10

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
End Sub

' バッジ(セル。獲得済みは🏅、未獲得は🔒)
Private Sub DrawBadges(ByVal ws As Worksheet)
    Dim r As Long: r = 22
    With ws.Range("B" & r & ":E" & r)
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

    With ws.Range("B" & (r + 1) & ":E" & (r + 3))
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

Public Sub OnGoShelf()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
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

Public Sub OnHelp()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modHelp.OnHelpClick
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnSaveAndExit()
    modApp.OnSaveAndExit
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

' 32bit環境ではShadow.Blurが未対応のことがある。個別にResume Nextで包む。
Private Sub SafeShadow(ByVal shp As Shape)
    On Error Resume Next
    shp.Shadow.Visible = -1
    shp.Shadow.OffsetX = 0
    shp.Shadow.OffsetY = 1.5
    shp.Shadow.Transparency = 0.88
    shp.Shadow.ForeColor.RGB = RGB(0, 0, 0)
    shp.Shadow.Blur = 5
    On Error GoTo 0
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
