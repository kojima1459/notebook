Attribute VB_Name = "modHelp"
Option Explicit

' ============================================================================
' modHelp - ヘルプボタン+ヘルプ概要カード+詳細マニュアル(使い方シート)への
'   導線。完全疎結合のオプトイン機能。UI層のみに閉じたフック・モジュール
'   (modMentor/modPeek/modTourと同型のパターン)。
' ----------------------------------------------------------------------------
' 役割:
'   トップバー右端に丸い「?」ボタンを常設し、クリックすると概念ガイド
'   (チャット/出典チップ/アクションバー/ナレッジ倉庫/専門家/再描画)を
'   要約したカードを表示する。カードから①既存の「使い方」シート(詳細
'   マニュアル)を開く導線 ②modTourのオンボーディングを再生する導線、
'   の2つを提供する。
'
' 防衛設計:
'   1. モジュール隔離: エントリはEnsureHelpButton(LaunchNexus末尾の1行
'      フック)のみ。既存モジュールへは一切書き込まない。
'   2. サーキットブレーカー: 全Publicエントリの最上部でOn Error Resume Next。
'      本機能の失敗は「ボタン/カードが出ない」に留め、チャット本体機能へ
'      絶対に波及させない。
'   3. クリックハンドラは全てmodUiLock.Enter/Leaveで多重発火を防ぐ。
'      内部の削除処理(DoHideHelp)はPrivateに分離し、既にロックを保持した
'      文脈(OnOpenManual/OnRestartTour)からも再入デッドロックなしで
'      呼べるようにしてある(modUiLock.Enterは非再入のため)。
'   4. Shape名は全て "nx_help_" / "nxh_" 接頭辞で、既存の nx_sb_/nx_top_/
'      nx_fab_/nx_msg_/nx_thk_/nx_cite_/nx_peek/nx_mentor_/nx_toast/
'      nx_tour_ と衝突しない。
' ============================================================================

' ----------------------------------------------------------------------------
' EnsureHelpButton - エントリポイント(架け元: modApp.LaunchNexus末尾の
'   1行フック、想定)。トップバー右端に丸い「?」ボタンを描く(冪等)。
' ----------------------------------------------------------------------------
Public Sub EnsureHelpButton()
    On Error Resume Next   ' 安全弁: 本機能の失敗を絶対にメインへ波及させない
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub

    ws.Shapes("nx_help_btn").Delete   ' 冪等: 再呼び出しでの孤児/重複を防ぐ

    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(9, 195 + 662, 9, 27, 27)   ' 9=楕円
    btn.Name = "nx_help_btn"
    btn.Fill.ForeColor.RGB = modUI.UiColor("surface")
    btn.Line.Visible = -1
    btn.Line.Weight = 0.75
    btn.Line.ForeColor.RGB = modUI.UiColor("border")
    With btn.TextFrame2
        .TextRange.Text = "?"
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 12
        .TextRange.Font.Bold = -1
        .TextRange.ParagraphFormat.Alignment = 2   ' 中央
        .VerticalAnchor = 3
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
    End With
    btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
    btn.OnAction = "modHelp.OnHelpClick"
    btn.Placement = 3   ' xlFreeFloating
    btn.ZOrder 0        ' msoBringToFront
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' OnHelpClick - 「?」ボタンのクリック。概要カードを表示する。
' ----------------------------------------------------------------------------
Public Sub OnHelpClick()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    ShowHelpCard
Done:
    modUiLock.Leave
End Sub

' ----------------------------------------------------------------------------
' HideHelp - カード自身のクリック(閉じる)用のOnActionターゲット。
' ----------------------------------------------------------------------------
Public Sub HideHelp()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    DoHideHelp
Done:
    modUiLock.Leave
End Sub

' ----------------------------------------------------------------------------
' OnOpenManual - 「詳細マニュアルを開く」クリック。既存の使い方シートを
'   表示・アクティブ化し、そのシート上に「Nexusへ戻る」フローティング
'   ボタンを描く(冪等)。
' ----------------------------------------------------------------------------
Public Sub OnOpenManual()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done

    DoHideHelp

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_HOWTO)
    If ws Is Nothing Then GoTo Done

    ws.Visible = -1   ' xlSheetVisible
    ws.Activate
    DrawManualBackButton ws
Done:
    modUiLock.Leave
End Sub

' ----------------------------------------------------------------------------
' OnBackToNexus - 使い方シート上の「← Nexusへ戻る」クリック。
' ----------------------------------------------------------------------------
Public Sub OnBackToNexus()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    ThisWorkbook.Worksheets("Nexus").Activate
Done:
    modUiLock.Leave
End Sub

' ----------------------------------------------------------------------------
' OnRestartTour - 「ツアーをもう一度見る」クリック。ヘルプを閉じてから
'   modTour.RestartTourへ委譲する。
' ----------------------------------------------------------------------------
Public Sub OnRestartTour()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    DoHideHelp
    modTour.RestartTour
Done:
    modUiLock.Leave
End Sub

' ----------------------------------------------------------------------------
' 内部: ShowHelpCard - Nexusシート中央付近に概要カード+2つの導線ボタンを
'   描画する(先にDoHideHelpで前回分を消す)。
' ----------------------------------------------------------------------------
Private Sub ShowHelpCard()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub

    DoHideHelp

    Dim cardL As Double: cardL = 260
    Dim cardT As Double: cardT = 100
    Dim cardW As Double: cardW = 500

    Dim card As Shape
    Set card = ws.Shapes.AddShape(5, cardL, cardT, cardW, 60)   ' 5=角丸四角(高さはAutoSize)
    card.Name = "nx_help_card"
    card.Adjustments(1) = 0.05
    card.Fill.ForeColor.RGB = modUI.UiColor("surface")
    card.Line.Visible = -1
    card.Line.Weight = 1#
    card.Line.ForeColor.RGB = modUI.UiColor("primary")
    With card.TextFrame2
        .WordWrap = -1
        .AutoSize = 1   ' msoAutoSizeShapeToFitText
        .MarginLeft = 16: .MarginRight = 16: .MarginTop = 14: .MarginBottom = 14
        .TextRange.Text = HelpBodyText()
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 9.5
        .TextRange.ParagraphFormat.Alignment = 1   ' 左
    End With
    card.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
    card.OnAction = "modHelp.HideHelp"   ' クリックで閉じる
    card.Placement = 3
    modSkin.ApplySoftShadow card
    card.ZOrder 0

    ' AutoSize確定後の高さを使って下にボタンを並べる
    Dim belowT As Double: belowT = card.Top + card.Height + 8

    Dim manualBtn As Shape
    Set manualBtn = ws.Shapes.AddShape(5, cardL, belowT, 170, 28)
    manualBtn.Name = "nx_help_manual"
    manualBtn.Adjustments(1) = 0.3
    manualBtn.Fill.ForeColor.RGB = modUI.UiColor("primary")
    manualBtn.Line.Visible = 0
    With manualBtn.TextFrame2
        .WordWrap = -1
        .TextRange.Text = ChrW(&H1F4D6) & " 詳細マニュアルを開く"
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = -1
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
    End With
    manualBtn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    manualBtn.OnAction = "modHelp.OnOpenManual"
    manualBtn.Placement = 3
    manualBtn.ZOrder 0

    Dim tourBtn As Shape
    Set tourBtn = ws.Shapes.AddShape(5, cardL + 170 + 10, belowT, 170, 28)
    tourBtn.Name = "nx_help_tour"
    tourBtn.Adjustments(1) = 0.3
    tourBtn.Fill.ForeColor.RGB = modUI.UiColor("surface")
    tourBtn.Line.Visible = -1
    tourBtn.Line.Weight = 1#
    tourBtn.Line.ForeColor.RGB = modUI.UiColor("accent")
    With tourBtn.TextFrame2
        .WordWrap = -1
        .TextRange.Text = ChrW(&H2728) & " ツアーをもう一度見る"
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = -1
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
    End With
    tourBtn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("accent")
    tourBtn.OnAction = "modHelp.OnRestartTour"
    tourBtn.Placement = 3
    tourBtn.ZOrder 0

    On Error GoTo 0
End Sub

' ヘルプカードの本文(コンシェルジュ風の簡潔ガイド)。
Private Function HelpBodyText() As String
    Dim s As String
    s = ChrW(&H1F4D6) & " Nexus Agent かんたんガイド" & vbLf & vbLf & _
        ChrW(&H1F4AC) & " チャット: 入力欄に質問して「送信」。モードボタンで" & _
        "「社内ナレッジ検索」(出典付き)と「一般アシスタント」を切替。" & vbLf & _
        ChrW(&H1F4C4) & " 出典チップ: 回答下のチップをクリックすると原文をその場で確認できます。" & vbLf & _
        "アクションバー: " & ChrW(&H1F44D) & "/" & ChrW(&H1F44E) & "で評価、" & ChrW(&H1F50D) & _
        "深掘り、" & ChrW(&H2705) & "解決した(EXP+作者へ感謝が届く)、" & ChrW(&H1F4CB) & _
        "コピー、" & ChrW(&H1F4C4) & "Word出力。" & vbLf & _
        ChrW(&H1F4DA) & " ナレッジ倉庫: 資料の登録・検索・カード詳細。品質が低い資料は " & _
        ChrW(&H26A0) & "ノイズ報告 で検索から除外できます。" & vbLf & _
        ChrW(&H1F4A1) & " 専門家: 回答の下に「〇〇さんが詳しいです」と出たら、ボタンから" & _
        "直接質問を送れます。" & vbLf & _
        ChrW(&H1F504) & " 画面が乱れたら: サイドバーの「画面を再描画」。" & vbLf & vbLf & _
        "このカードはクリックで閉じます"
    HelpBodyText = s
End Function

' 使い方シート上の「← Nexusへ戻る」フローティングボタン(冪等・そのシート専用)。
Private Sub DrawManualBackButton(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Shapes("nxh_back").Delete

    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(5, 10, 8, 150, 28)   ' 5=角丸四角
    btn.Name = "nxh_back"
    btn.Adjustments(1) = 0.3
    btn.Fill.ForeColor.RGB = modUI.UiColor("primary")
    btn.Line.Visible = 0
    With btn.TextFrame2
        .WordWrap = -1
        .TextRange.Text = ChrW(&H2190) & " Nexusへ戻る"
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = -1
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
    End With
    btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    btn.OnAction = "modHelp.OnBackToNexus"
    btn.Placement = 3
    btn.ZOrder 0
    On Error GoTo 0
End Sub

' 内部: ヘルプカード+導線ボタンの削除(孤児防止)。ロックを取らない生の
' 削除処理として分離し、既にmodUiLockを保持している呼び出し元
' (OnOpenManual/OnRestartTour)からも再入デッドロックなしで呼べるようにする。
Private Sub DoHideHelp()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub
    ws.Shapes("nx_help_card").Delete
    ws.Shapes("nx_help_manual").Delete
    ws.Shapes("nx_help_tour").Delete
    On Error GoTo 0
End Sub
