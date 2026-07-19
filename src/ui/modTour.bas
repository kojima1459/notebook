Attribute VB_Name = "modTour"
Option Explicit

' ============================================================================
' modTour - 初回起動オンボーディングツアー(3ステップ)。完全疎結合の
'   オプトイン機能。UI層のみに閉じたフック・モジュール(modMentor/modPeekと
'   同型のパターン)。
' ----------------------------------------------------------------------------
' 役割:
'   初めてNexusを開いたユーザーへ、①質問の仕方 ②出典チップの確認方法
'   ③解決済みアクション(EXP/感謝)の3点を、画面上の該当箇所に寄り添う
'   小さなカードで順番に案内する。完了/スキップは ui_state シートへ
'   "nexus_tour_done"="1" として永続化し、以後は出さない。
'   ヘルプ(modHelp)から「ツアーをもう一度見る」で再起動できる。
'
' 防衛設計:
'   1. モジュール隔離: エントリはStartTourIfFirstRun(LaunchNexus末尾の1行
'      フック)のみ。既存モジュールへは一切書き込まない。
'   2. サーキットブレーカー: 全Publicエントリの最上部でOn Error Resume Next。
'      本機能の失敗は「カードが出ない/消えない」に留め、チャット本体機能へ
'      絶対に波及させない。
'   3. クリックハンドラ(OnTourNext/OnTourSkip)はmodUiLock.Enter/Leaveで
'      連打・多重発火を防ぐ(modMentor.OnAskExpertと同型)。
'   4. Shape名は全て "nx_tour_" 接頭辞で、既存の nx_sb_/nx_top_/nx_fab_/
'      nx_msg_/nx_thk_/nx_cite_/nx_peek/nx_mentor_/nx_toast と衝突しない。
' ============================================================================

Private mStep As Long   ' 現在のツアー段階(1〜3)。モジュール状態はこれ1個のみ。

' ----------------------------------------------------------------------------
' StartTourIfFirstRun - エントリポイント(架け元: modApp.LaunchNexus末尾の
'   1行フック、想定)。完了済みなら即撤退。未完了なら1段目から描画する。
' ----------------------------------------------------------------------------
Public Sub StartTourIfFirstRun()
    On Error Resume Next   ' 安全弁: 本機能の失敗を絶対にメインへ波及させない
    If modState.LoadState("nexus_tour_done", "") = "1" Then Exit Sub
    ClearTour
    DrawStep 1
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' OnTourNext - 「次へ」クリック。1→2→3と進め、3の次で完了扱いにする。
' ----------------------------------------------------------------------------
Public Sub OnTourNext()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done

    Dim nextStep As Long: nextStep = mStep + 1
    If nextStep > 3 Then
        ClearTour
        modState.SaveState "nexus_tour_done", "1"
        On Error Resume Next
        modSkin.ShowToast "ツアー完了です。さっそく質問してみましょう。", "success"
        On Error GoTo Done
    Else
        DrawStep nextStep
    End If
Done:
    modUiLock.Leave
End Sub

' ----------------------------------------------------------------------------
' OnTourSkip - 「スキップ」クリック。完了扱いにして即座に閉じる。
' ----------------------------------------------------------------------------
Public Sub OnTourSkip()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done

    ClearTour
    modState.SaveState "nexus_tour_done", "1"
    On Error Resume Next
    modSkin.ShowToast "ツアーはいつでもヘルプ(右上の?)から見直せます。", "info"
    On Error GoTo Done
Done:
    modUiLock.Leave
End Sub

' ----------------------------------------------------------------------------
' RestartTour - ヘルプ画面の「ツアーをもう一度見る」から呼ばれる(modHelp側で
'   既にmodUiLockを取得済みの文脈から呼ばれる想定のため、本Sub自身は
'   ロックを取らない。サーキットブレーカーのみ)。完了フラグを消して1段目から。
' ----------------------------------------------------------------------------
Public Sub RestartTour()
    On Error Resume Next
    modState.SaveState "nexus_tour_done", ""
    ClearTour
    DrawStep 1
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' ClearTour - nx_tour_* Shapeを全削除(孤児防止)。列挙中は削除しない
'   (名前を先に集めてから削除)。
' ----------------------------------------------------------------------------
Public Sub ClearTour()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub

    Dim names() As String
    ReDim names(0 To ws.Shapes.count)
    Dim n As Long: n = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 8) = "nx_tour_" Then
            names(n) = shp.Name
            n = n + 1
        End If
    Next shp

    Dim i As Long
    For i = 0 To n - 1
        ws.Shapes(names(i)).Delete
    Next i
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 内部: DrawStep - 段n(1〜3)のカード+ボタンを描画する(先にClearTour)。
'   説明対象の近くに置く: 1=入力欄付近、2/3=画面中段(出典/アクションバー付近)。
' ----------------------------------------------------------------------------
Private Sub DrawStep(ByVal n As Long)
    On Error Resume Next
    ClearTour

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub

    mStep = n

    Dim cardL As Double: cardL = 195 + 15
    Dim cardT As Double
    Select Case n
        Case 1: cardT = 100
        Case Else: cardT = 240
    End Select
    Dim cardW As Double: cardW = 400
    Dim cardH As Double: cardH = 130

    Dim card As Shape
    Set card = ws.Shapes.AddShape(5, cardL, cardT, cardW, cardH)   ' 5=角丸四角
    card.Name = "nx_tour_card"
    card.Adjustments(1) = 0.08
    card.Fill.ForeColor.RGB = modUI.UiColor("surface")
    card.Line.Visible = -1
    card.Line.Weight = 1.25
    card.Line.ForeColor.RGB = modUI.UiColor("accent")
    card.Placement = 3
    modSkin.ApplySoftShadow card
    card.ZOrder 0

    ' 段数表示("1 / 3")
    Dim stepLbl As Shape
    Set stepLbl = ws.Shapes.AddShape(1, cardL + cardW - 60, cardT + 8, 48, 14)
    stepLbl.Name = "nx_tour_step"
    stepLbl.Fill.Visible = 0
    stepLbl.Line.Visible = 0
    With stepLbl.TextFrame2
        .TextRange.Text = CStr(n) & " / 3"
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 8
        .TextRange.ParagraphFormat.Alignment = 3   ' 右
        .MarginLeft = 0: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
    End With
    stepLbl.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
    stepLbl.Placement = 3
    stepLbl.ZOrder 0

    Dim titleText As String, bodyText As String, nextLabel As String
    Select Case n
        Case 1
            titleText = ChrW(&H2460) & " 質問してみましょう"
            bodyText = "上の入力欄にメッセージを入力し「送信」。「社内ナレッジ検索」" & _
                "モードなら本棚の資料から出典付きで回答します。"
            nextLabel = "次へ " & ChrW(&H2192)
        Case 2
            titleText = ChrW(&H2461) & " 出典をワンクリック確認"
            bodyText = "AI回答の下の「" & ChrW(&H1F4C4) & " 出典チップ」を押すと、元の資料の" & _
                "該当箇所がその場で確認できます。AIの回答が正しいか、秒でチェック。"
            nextLabel = "次へ " & ChrW(&H2192)
        Case Else
            titleText = ChrW(&H2462) & " 役立ったら " & ChrW(&H2705) & "解決した"
            bodyText = "上のアクションバーで評価できます。「" & ChrW(&H2705) & "解決した」で" & _
                "EXPが貯まり、資料の作者へ「ありがとう」が自動で届きます。"
            nextLabel = "はじめる " & ChrW(&H2728)
    End Select

    Dim titleShp As Shape
    Set titleShp = ws.Shapes.AddShape(1, cardL + 14, cardT + 12, cardW - 28, 20)
    titleShp.Name = "nx_tour_title"
    titleShp.Fill.Visible = 0
    titleShp.Line.Visible = 0
    With titleShp.TextFrame2
        .WordWrap = -1
        .TextRange.Text = titleText
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 11
        .TextRange.Font.Bold = -1
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
    End With
    titleShp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
    titleShp.Placement = 3
    titleShp.ZOrder 0

    Dim bodyShp As Shape
    Set bodyShp = ws.Shapes.AddShape(1, cardL + 14, cardT + 38, cardW - 28, 60)
    bodyShp.Name = "nx_tour_body"
    bodyShp.Fill.Visible = 0
    bodyShp.Line.Visible = 0
    With bodyShp.TextFrame2
        .WordWrap = -1
        .TextRange.Text = bodyText
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 9.5
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
    End With
    bodyShp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
    bodyShp.Placement = 3
    bodyShp.ZOrder 0

    Dim btnW As Double: btnW = 78
    Dim btnH As Double: btnH = 24
    Dim btnT As Double: btnT = cardT + cardH - btnH - 10
    Dim nextL As Double: nextL = cardL + cardW - btnW - 10
    Dim skipL As Double: skipL = nextL - btnW - 8

    Dim skipBtn As Shape
    Set skipBtn = ws.Shapes.AddShape(5, skipL, btnT, btnW, btnH)
    skipBtn.Name = "nx_tour_skip"
    skipBtn.Adjustments(1) = 0.4
    skipBtn.Fill.ForeColor.RGB = modUI.UiColor("surface")
    skipBtn.Line.Visible = -1
    skipBtn.Line.Weight = 0.75
    skipBtn.Line.ForeColor.RGB = modUI.UiColor("border")
    With skipBtn.TextFrame2
        .TextRange.Text = "スキップ"
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 9
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
    End With
    skipBtn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
    skipBtn.OnAction = "modTour.OnTourSkip"
    skipBtn.Placement = 3
    skipBtn.ZOrder 0

    Dim nextBtn As Shape
    Set nextBtn = ws.Shapes.AddShape(5, nextL, btnT, btnW, btnH)
    nextBtn.Name = "nx_tour_next"
    nextBtn.Adjustments(1) = 0.4
    nextBtn.Fill.ForeColor.RGB = modUI.UiColor("primary")
    nextBtn.Line.Visible = 0
    With nextBtn.TextFrame2
        .TextRange.Text = nextLabel
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = -1
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
    End With
    nextBtn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    nextBtn.OnAction = "modTour.OnTourNext"
    nextBtn.Placement = 3
    nextBtn.ZOrder 0

    On Error GoTo 0
End Sub
