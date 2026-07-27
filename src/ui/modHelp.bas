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
' EnsureHelpButton - 旧「?」浮きボタンの掃除(架け元: modApp.LaunchNexus)。
'   2026-07-26: 「?」はチャットのヘッダー(nx_top_help)とHubのヘッダー
'   アイコンへ統合した。旧座標(195+662)は廃止したサイドバー幅が前提で、
'   サイドバーを外した今は画面外/変な位置に浮くため、生成をやめて
'   既存ブックに残っている分の削除だけを行う。
' ----------------------------------------------------------------------------
Public Sub EnsureHelpButton()
    On Error Resume Next   ' 安全弁: 本機能の失敗を絶対にメインへ波及させない
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub
    ws.Shapes("nx_help_btn").Delete
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
    modUI.GoToNexus "modHelp.OnBackToNexus"
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

    ' サイドバー廃止(2026-07-26)で左端が変わったため、決め打ちのx=260ではなく
    ' チャット領域の実測幾何から中央寄せする。
    Dim cardW As Double: cardW = 500
    Dim cardL As Double
    Dim cardT As Double: cardT = modUINexusDraw.ChatTop(ws) + 20
    cardL = modUINexusDraw.ChatLeft(ws) + (modUINexusDraw.ChatWidth(ws) - cardW) / 2
    If cardL < 8 Then cardL = 8

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
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCD6) & " 詳細マニュアルを開く"
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

    ' 2段目: ご意見箱(EXP)+P2P接続設定
    Dim fbBtn As Shape
    Set fbBtn = ws.Shapes.AddShape(5, cardL, belowT + 34, 170, 28)
    fbBtn.Name = "nx_help_fb"
    fbBtn.Adjustments(1) = 0.3
    fbBtn.Fill.ForeColor.RGB = modUI.UiColor("surface")
    fbBtn.Line.Visible = -1
    fbBtn.Line.Weight = 1#
    fbBtn.Line.ForeColor.RGB = modUI.UiColor("accent")
    With fbBtn.TextFrame2
        .WordWrap = -1
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCEE) & " ご意見・不具合報告"
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 8.5
        .TextRange.Font.Bold = -1
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
    End With
    fbBtn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("accent")
    fbBtn.OnAction = "modHelp.OnFeedback"
    fbBtn.Placement = 3
    fbBtn.ZOrder 0

    Dim cfgBtn As Shape
    Set cfgBtn = ws.Shapes.AddShape(5, cardL + 170 + 10, belowT + 34, 170, 28)
    cfgBtn.Name = "nx_help_cfg"
    cfgBtn.Adjustments(1) = 0.3
    cfgBtn.Fill.ForeColor.RGB = modUI.UiColor("surface")
    cfgBtn.Line.Visible = -1
    cfgBtn.Line.Weight = 0.75
    cfgBtn.Line.ForeColor.RGB = modUI.UiColor("border")
    With cfgBtn.TextFrame2
        .WordWrap = -1
        .TextRange.Text = ChrW(&H2699) & " P2P接続設定(共有フォルダ)"
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 8.5
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
    End With
    cfgBtn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
    cfgBtn.OnAction = "modHelp.OnShareSetup"
    cfgBtn.Placement = 3
    cfgBtn.ZOrder 0

    ' 3段目: きせかえ(感謝数で解放されるスキン)
    Dim skinBtn As Shape
    Set skinBtn = ws.Shapes.AddShape(5, cardL, belowT + 68, 350, 28)
    skinBtn.Name = "nx_help_skin"
    skinBtn.Adjustments(1) = 0.3
    skinBtn.Fill.ForeColor.RGB = modUI.UiColor("surface")
    skinBtn.Line.Visible = -1
    skinBtn.Line.Weight = 0.75
    skinBtn.Line.ForeColor.RGB = modUI.UiColor("border")
    With skinBtn.TextFrame2
        .WordWrap = -1
        .TextRange.Text = ChrW(&HD83C) & ChrW(&HDFA8) & " きせかえ(「ありがとう」を集めると限定スキンが解放)"
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 8.5
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
    End With
    skinBtn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
    skinBtn.OnAction = "modHelp.OnCycleSkin"
    skinBtn.Placement = 3
    skinBtn.ZOrder 0

    On Error GoTo 0
End Sub

' きせかえボタン(ヘルプカードを閉じてからmodSkin.CycleSkinへ。CycleSkin自身が
' modUiLockを取るため、ここではロックを取らず閉じ処理のみ行う)。
Public Sub OnCycleSkin()
    On Error Resume Next
    DoHideHelp
    On Error GoTo 0
    modSkin.CycleSkin
End Sub

' ----------------------------------------------------------------------------
' OnFeedback - ご意見箱(B-2)。感想・不具合を入力→本文をクリップボードへ格納し、
'   作成者宛のOutlook新規メールを開く(mailto)。共有フォルダ設定に依存しない
'   最も確実な経路。送信協力へのお礼として1日1回EXP+5(バグバウンティ)。
' ----------------------------------------------------------------------------
Public Sub OnFeedback()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    DoHideHelp

    Dim fb As String
    fb = InputBox("Nexus Agentへのご意見・改善案・不具合(エラーの状況など)を教えてください。" & vbCrLf & _
                  "いただいた内容はすべて作成者(小島)が読み、改善に活かします。", _
                  "Nexus Agent - ご意見箱")
    fb = Trim$(fb)
    If LenB(fb) = 0 Then GoTo Done

    ' 本文(環境情報つき)をクリップボードへ。日本語のmailto本文は文字化けし得るため
    ' 「件名はmailtoで、本文はCtrl+V貼り付け」方式が最も確実。
    Dim body As String
    body = "【Nexus Agent ご意見・不具合報告】" & vbCrLf & fb & vbCrLf & vbCrLf & _
           "--- 環境情報(自動付記) ---" & vbCrLf & _
           "Ver: " & modAppDef.APP_VERSION & " / " & modUtil.NowStamp()
    modClip.SetClipboardText body

    On Error Resume Next
    ThisWorkbook.FollowHyperlink "mailto:m-kojima@aioinissaydowa.co.jp?subject=Nexus%20Agent%20feedback"
    On Error GoTo Done

    ' バグバウンティEXP(1日1回まで=空メール連打での稼ぎを防止)
    Dim dayKey As String: dayKey = "fb:" & Format$(Date, "yyyymmdd")
    If modStats.GetStat(dayKey) = 0 Then
        modStats.Bump dayKey
        modStats.AddExp "feedback"
        modSkin.ShowToast "ありがとうございます。メールが開くので Ctrl+V で本文を貼り付けて送信してください。", "success"
    Else
        modSkin.ShowToast "ありがとうございます。メールが開くので Ctrl+V で本文を貼り付けて送信してください。", "success"
    End If
    On Error Resume Next
    modLog.LogUsage "feedback_box", "", modUtil.SafeLeft(fb, 120)
    On Error GoTo Done
Done:
    modUiLock.Leave
End Sub

' ----------------------------------------------------------------------------
' OnShareSetup - P2P接続設定(B-1)。隠しconfigシートを触らせずに、共有フォルダの
'   パスをダイアログで設定できる唯一の窓口。保存後はボードを即時再構築。
' ----------------------------------------------------------------------------
Public Sub OnShareSetup()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    DoHideHelp

    Dim cur As String
    cur = modConfig.GetString("nexus_share_path", "")
    Dim p As String
    p = InputBox("P2P共有フォルダ(感謝状・専門家への質問・みんなの節約時間で使用)の" & vbCrLf & _
                 "パスを入力してください。チームで同じフォルダを指定します。" & vbCrLf & _
                 "例: \\サーバー名\共有\Nexus_Share\ (現在: " & IIf(LenB(cur) > 0, cur, "未設定") & ")", _
                 "Nexus Agent - P2P接続設定", cur)
    p = Trim$(p)
    If LenB(p) = 0 Then GoTo Done
    If Right$(p, 1) <> "\" Then p = p & "\"

    If Len(Dir(p, vbDirectory)) = 0 Then
        modSkin.ShowToast "そのフォルダが見つかりませんでした。パスをご確認ください(設定は変更していません)。", "error"
        GoTo Done
    End If

    modConfig.SetValue "nexus_share_path", p
    modSkin.ShowToast "接続しました。感謝状・専門家への質問・みんなの節約時間が使えます。", "success"
    On Error Resume Next
    modBoard.BootBoard   ' ウィジェットを即時再構築(次回起動を待たせない)
    On Error GoTo Done
Done:
    modUiLock.Leave
End Sub

' ヘルプカードの本文(コンシェルジュ風の簡潔ガイド)。
Private Function HelpBodyText() As String
    Dim s As String
    s = ChrW(&HD83D) & ChrW(&HDCD6) & " Nexus Agent かんたんガイド" & vbLf & vbLf & _
        ChrW(&HD83D) & ChrW(&HDCAC) & " チャット: 入力欄に質問して「送信」。モードボタンで" & _
        "「社内ナレッジ検索」(出典付き)と「一般アシスタント」を切替。" & vbLf & _
        ChrW(&HD83D) & ChrW(&HDCC4) & " 出典チップ: 回答下のチップをクリックすると原文をその場で確認できます。" & vbLf & _
        "アクションバー: " & ChrW(&HD83D) & ChrW(&HDC4D) & "/" & ChrW(&HD83D) & ChrW(&HDC4E) & "で評価、" & ChrW(&HD83D) & ChrW(&HDD0D) & _
        "深掘り、" & ChrW(&H2705) & "解決した(資料を書いた人へ感謝が届く)、" & ChrW(&HD83D) & ChrW(&HDCCB) & _
        "コピー、" & ChrW(&HD83D) & ChrW(&HDCC4) & "Word出力。" & vbLf & _
        ChrW(&HD83D) & ChrW(&HDCDA) & " ナレッジ倉庫: 資料の登録・検索・カード詳細。品質が低い資料は " & _
        ChrW(&H26A0) & "ノイズ報告 で検索から除外できます。" & vbLf & _
        ChrW(&HD83D) & ChrW(&HDCA1) & " 専門家: 回答の下に「〇〇さんが詳しいです」と出たら、ボタンから" & _
        "直接質問を送れます。" & vbLf & _
        ChrW(&HD83D) & ChrW(&HDD04) & " 画面が乱れたら: サイドバーの「画面を再描画」。" & vbLf & _
        ChrW(&H2328) & " ショートカット: Ctrl+Enter=送信 / Ctrl+Shift+Q=どこからでも呼び出し。" & vbLf & vbLf & _
        "作成: リスクコンサルティング支援部 ニューリスクG 小島正豪" & vbLf & _
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
    ws.Shapes("nx_help_fb").Delete
    ws.Shapes("nx_help_cfg").Delete
    ws.Shapes("nx_help_skin").Delete
    On Error GoTo 0
End Sub
