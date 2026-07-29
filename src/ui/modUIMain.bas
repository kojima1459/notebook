Attribute VB_Name = "modUIMain"
Option Explicit

' modUIMain - 「ホーム」画面(質問UI)の構築と描画(MASTER_SPEC §7.6/§8.1)
' 役割:
'   非エンジニアが説明書なしで使える「質問する画面」。EnsureLayoutが冪等に
'   Shape(ボタン)とセルレイアウトを再構築し、SetStage/RenderAnswer/
'   RenderSourcesPreviewが実行中の進捗と結果を表示する。modAskはこのモジュール
'   の名前付き範囲 "mb_question"(質問入力セル)を読み、SetStage/
'   RenderSourcesPreview/RenderAnswerを呼び返す(§7.3 modAskの設計コメント
'   のとおり)。
'
' 設計判断:
'   ・ラベルは原則セル値(V2 modChatUI.bas踏襲)とし、実際にクリックされる
'     ボタンだけをShape(msoShapeRoundedRectangle)にする。Shape名は
'     btn_/lbl_ プレフィクス(§12)。EnsureLayoutは既存の自前Shapesを
'     全削除してから再生成する(冪等。ユーザーが壊しても開き直せば直る)。
'   ・質問入力セルは "mb_question" という定義済み名前(Excel Name)を
'     A6セル(結合セルB6:H9ではなく左上のA6セルそのもの。結合セルの値は
'     左上セルに格納されるVBAの仕様を利用)に張る。modAsk.bas側の
'     アサンプションコメントと一致させること(Wave3で突き合わせ済み)。
'     プレースホルダ文字列は使わない(空のまま)。理由: プレースホルダ文字列
'     を入れると、ユーザーが消し忘れた場合にmodAsk側がそれをそのまま質問文
'     として送ってしまう(modAskはこのモジュールのプレースホルダ規約を
'     知らないため、二重管理を避けて「空セル=未入力」に統一した)。
'   ・SetStageは「セルへの書込み」「Application.StatusBarへの反映」を
'     必ず両方行う(§7.6契約)。処理中(msgが空でない)ときは待ち時間豆知識
'     ShowTipも合わせて更新し、ユーザーが手持ち無沙汰にならないようにする。
'   ・読み上げ(TTS)ボタンは置かない(裁定D10でボタンごと撤去。音声読み上げは
'     AIリボン本体でのみ利用可。出典: RIBBON_API_CONFIRMED.md §0/§2b D10)。
'   ・「続けて質問」ボタン(OnFollowupButton。裁定D11)は旧読み上げボタンの
'     位置(A31:D32)に置く。深掘りの実体(会話履歴の保持・再質問)はすべて
'     modAsk.CanFollowup/AskFollowupの責務で、ここはInputBoxで追質問を
'     受け取って渡すだけの薄いラッパーに徹する。
'   ・opt機能ボタン(Wordで開く)はmodFeatures.FeatureEnabledがTrueのときだけ
'     EnsureLayout内で生成する(§7.7)。押下時は「どんな文書に仕上げるか」の
'     指示文を尋ねてから(裁定D12: 対話型文書生成)、OnOpenWordButtonラッパー
'     経由でmodFeatures.InvokeFeature("markdown","ExportAnswerAsDoc",…)を
'     呼ぶ(opt直接参照はR2違反になるため、このモジュールにoptMarkdown等の
'     トークンは一切書かない)。
'   ・本棚が空のときの案内(§8.1「まず『マイ本棚』タブで資料を1つ追加して
'     みましょう →」)はShowEmptyShelfHintとして公開し、呼び出し判断(本棚が
'     空かどうか)はmodBoot側が行う(modUIMainはmodShelfの状態を勝手に
'     判定しない。UIモジュールは「表示のしかた」の責務に留める)。
'   ・OnAskButtonは質問実行後、EvaluateBadges/RenderDashboardの更新を
'     On Error Resume Next(1行スコープ)で試みる。ダッシュボードの更新に
'     失敗しても質問応答自体は成功として扱う(非致命)。
'   ・On Error Resume Next は §12の規約どおり必ず1行スコープ(直後に
'     On Error GoTo 0)で使う。複数行にまたがる保護が必要な場面では
'     個別のヘルパーSubに切り出して1行スコープを保つ。

Private Const RNG_STATUS As String = "A12:H12"
Private Const RNG_ANSWER As String = "A14:H25"
Private Const RNG_SOURCES As String = "A26:H27"
Private Const RNG_TIP As String = "A33:H34"
Private Const QUESTION_TOP_LEFT As String = "A6"
Private Const QUESTION_RANGE_NAME As String = "mb_question"

Private Const COLOR_SELECTED_BG As Long = 2039071    ' RGB(31,78,120) 濃紺
Private Const COLOR_SELECTED_FG As Long = 16777215   ' RGB(255,255,255) 白
' 配色は modUIMainShape(図形を描く側)を単一情報源にする。

Private mLastAnswerText As String

' 待ち時間の実況先は modLive が持つ。modAsk は進捗を SetStage /
' RenderSourcesPreview でここへ知らせる契約(§7.3・R1で許可された唯一のUI
' コールバック)だが、その宛先だった旧ホームシートのセルは、既定UIがNexus
' チャットになった今、利用者が一切見ていない。契約は変えずに転送先を足す。

' EnsureLayout - ホームを冪等再構築(既存Shapes全削除→再生成)
Public Sub EnsureLayout()
    Dim ws As Worksheet
    Set ws = modUIMainShape.GetOrCreateHomeSheet()
    If ws Is Nothing Then Exit Sub

    ' uiStep: 実機エラーの発生箇所を1回の報告で特定するための進捗マーカー。
    Dim uiStep As String
    On Error GoTo Fail

    Application.ScreenUpdating = False
    If ws.Visible <> -1 Then ws.Visible = -1   ' xlSheetVisible(非表示なら表示化)

    Dim prevActive As Object
    On Error Resume Next
    Set prevActive = ThisWorkbook.ActiveSheet
    On Error GoTo Fail

    uiStep = "既存ボタンの削除"
    modUIMainShape.RemoveManagedShapes ws
    DoEvents   ' 削除と追加の間でCOM/メモリを一拍解放する(先回り防衛#4)
    uiStep = "セルのクリア"
    ws.Cells.Clear

    uiStep = "既定フォント設定"
    ws.Cells.Font.Name = "游ゴシック"
    ws.Cells.Font.Size = 11

    ws.Columns("A:H").ColumnWidth = 12
    ws.Columns("I").ColumnWidth = 12

    ' ---- タイトル行 ----------------------------------------------------
    uiStep = "タイトル行"
    With ws.Range("A1:E1")
        .Merge
        .Value = "" & ChrW(&HD83D) & ChrW(&HDCDA) & " " & modAppDef.APP_NAME & "  v" & modAppDef.APP_VERSION
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = COLOR_SELECTED_BG
        .Font.Color = COLOR_SELECTED_FG
        .HorizontalAlignment = -4108   ' xlCenter
        .VerticalAlignment = -4108     ' xlCenter
    End With
    ws.Rows("1:2").RowHeight = 22

    uiStep = "サブタイトル行"
    With ws.Range("A2:E2")
        .Merge
        .Value = "「自分で入れた資料に、すぐ聞ける」"
        .HorizontalAlignment = -4108
        .Font.Size = 10
        .Font.Italic = True
    End With

    ' Activate失敗を致命的にしない(失敗しても続行)。
    uiStep = "描画前アクティブ化"
    Application.ScreenUpdating = True
    DoEvents
    On Error Resume Next
    ws.Activate
    If Err.Number <> 0 Then
        Dim actNum As Long: actNum = Err.Number
        modLog.LogError "E0801", "modUIMain.EnsureLayout", _
            "[描画前アクティブ化(許容続行)] ws.Visible=" & ws.Visible & _
            " ActiveSheet=" & ThisWorkbook.ActiveSheet.Name & _
            " AppWin=" & Application.Windows.Count & " WbWin=" & ThisWorkbook.Windows.Count, actNum
        Err.Clear
    End If
    On Error GoTo Fail
    Application.ScreenUpdating = False

    uiStep = "ボタン(使い方/診断)"
    modUIMainShape.AddButton ws, ws.Range("F1:G2"), "btn_howto", ChrW(&H2753) & " 使い方", "modUIMain.OnOpenHowto"
    modUIMainShape.AddButton ws, ws.Range("H1:H2"), "btn_diag", "" & ChrW(&HD83E) & ChrW(&HDE7A) & " 診断", "modUIMain.OnRunDiag"
    ' 2026-07-22実機報告対策: Nexus(チャット)へタブなしで戻れる導線
    modUIMainShape.AddButton ws, ws.Range("I1:I2"), "btn_back_chat", "" & ChrW(&HD83D) & ChrW(&HDCAC) & " チャットへ", "modUIMain.OnBackToChat"

    ' ---- モードトグル ----------------------------------------------------
    uiStep = "モードトグルボタン"
    ws.Rows("3:4").RowHeight = 20
    modUIMainShape.AddButton ws, ws.Range("A3:D4"), "btn_mode_quick", ModeCaption("quick"), "modUIMain.OnModeQuick"
    modUIMainShape.AddButton ws, ws.Range("E3:H4"), "btn_mode_deep", ModeCaption("deep"), "modUIMain.OnModeDeep"

    ' ---- 質問入力 ----------------------------------------------------
    uiStep = "質問入力欄の見出し"
    With ws.Range("A5:H5")
        .Merge
        .Value = "質問をここに入力してください(例: 〇〇の手続きに必要な書類は?)"
        .Font.Size = 10
        .Font.Bold = True
    End With
    ws.Rows("5").RowHeight = 18

    uiStep = "質問入力欄"
    With ws.Range("A6:H9")
        .Merge
        .WrapText = True
        .VerticalAlignment = -4160   ' xlTop
        .Interior.Color = 16513526   ' 薄い黄色
        .Borders.LineStyle = 1
    End With
    ws.Rows("6:9").RowHeight = 20
    uiStep = "質問入力欄への名前付け"
    EnsureQuestionName ws

    ' ---- 質問するボタン ----------------------------------------------------
    uiStep = "質問するボタン"
    ws.Rows("10:11").RowHeight = 20
    modUIMainShape.AddButton ws, ws.Range("C10:F11"), "btn_ask", "" & ChrW(&HD83D) & ChrW(&HDCAC) & "  質 問 す る", "modUIMain.OnAskButton"

    ' ---- 状態表示 ----------------------------------------------------
    uiStep = "状態表示行"
    With ws.Range(RNG_STATUS)
        .Merge
        .Value = "状態: 準備できています"
        .Font.Size = 9
        .Font.Color = 8421504   ' RGB(128,128,128)
    End With
    ws.Rows("12").RowHeight = 16

    uiStep = "回答見出し"
    With ws.Range("A13:H13")
        .Merge
        .Value = "─ 回答 ───────────────────────────"
        .Font.Bold = True
        .Font.Size = 10
    End With
    ws.Rows("13").RowHeight = 16

    ' ---- 回答本文 ----------------------------------------------------
    uiStep = "回答本文欄"
    With ws.Range(RNG_ANSWER)
        .Merge
        .Value = "（質問を入力して「" & ChrW(&HD83D) & ChrW(&HDCAC) & " 質問する」を押してください）"
        .WrapText = True
        .VerticalAlignment = -4160
        .Interior.Color = 15987946   ' 薄い水色
        .Borders.LineStyle = 1
        .Font.Size = 10
    End With
    ws.Rows("14:25").RowHeight = 15

    uiStep = "出典欄"
    With ws.Range(RNG_SOURCES)
        .Merge
        .Value = ""
        .WrapText = True
        .VerticalAlignment = -4160
        .Font.Size = 9
        .Font.Color = 4210752   ' 濃い灰
    End With
    ws.Rows("26:27").RowHeight = 15

    ' ---- フィードバック ----------------------------------------------------
    uiStep = "フィードバック見出し"
    With ws.Range("A28:H28")
        .Merge
        .Value = "この回答は役に立ちましたか?"
        .Font.Size = 10
    End With
    ws.Rows("28").RowHeight = 16

    uiStep = "フィードバックボタン"
    ws.Rows("29:30").RowHeight = 18
    modUIMainShape.AddButton ws, ws.Range("A29:C30"), "btn_fb_green", "" & ChrW(&HD83D) & ChrW(&HDFE2) & " 解決した!", "modAsk.FeedbackGreen"
    modUIMainShape.AddButton ws, ws.Range("D29:E30"), "btn_fb_yellow", "" & ChrW(&HD83D) & ChrW(&HDFE1) & " ヒントになった", "modAsk.FeedbackYellow"
    modUIMainShape.AddButton ws, ws.Range("F29:H30"), "btn_fb_red", "" & ChrW(&HD83D) & ChrW(&HDD34) & " だめだった", "modAsk.FeedbackRed"

    ' ---- 続けて質問+Wordで開く ------------------------------------------------
    ' 「続けて質問」(裁定D11)はコア機能(modAsk)への入口なので常時生成する。
    ' 「Wordで開く」はopt機能(markdown)なのでFeatureEnabledがTrueのときだけ
    ' 生成する(§7.7)。※読み上げボタンは裁定D10で撤去した(音声読み上げは
    ' AIリボン本体でのみ利用可)。
    uiStep = "続けて質問ボタン"
    ws.Rows("31:32").RowHeight = 18
    modUIMainShape.AddButton ws, ws.Range("A31:D32"), "btn_followup", "" & ChrW(&HD83D) & ChrW(&HDCAC) & " 続けて質問", "modUIMain.OnFollowupButton"
    uiStep = "Wordで開くボタン(有効判定)"
    If modFeatures.FeatureEnabled("markdown") Then
        uiStep = "Wordで開くボタン(生成)"
        modUIMainShape.AddButton ws, ws.Range("E31:H32"), "btn_word", "" & ChrW(&HD83D) & ChrW(&HDCDD) & " Wordで開く", "modUIMain.OnOpenWordButton"
    End If

    ' ---- 待ち時間豆知識 ----------------------------------------------------
    uiStep = "豆知識欄"
    With ws.Range(RNG_TIP)
        .Merge
        .Value = "" & ChrW(&HD83D) & ChrW(&HDCA1) & " 豆知識: "
        .WrapText = True
        .VerticalAlignment = -4160
        .Font.Size = 9
        .Font.Color = 8421504
    End With
    ws.Rows("33:34").RowHeight = 15

    uiStep = "モード配色の反映"
    ApplyModeColors CurrentMode()
    uiStep = "豆知識の表示"
    ShowTip

    On Error Resume Next
    If Not prevActive Is Nothing Then prevActive.Activate   ' 元のアクティブシートへ復帰
    On Error GoTo 0
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    ' どのブロックで型不一致等が起きたかをEnsureLayout自身が特定し、
    ' Descriptionに埋め込んでからmodBoot.Bootへ伝播させる(bootStageは
    ' 「ホーム画面の組み立て」としか分からないため、その内訳をここで補う)。
    Dim origNum As Long, origDesc As String
    origNum = Err.Number
    origDesc = Err.Description
    ' Err.Raise伝播に依存せず、失敗した瞬間にここで直接err_logへ書く。
    Dim diag As String: diag = ""
    On Error Resume Next
    diag = " ws.Visible=" & ws.Visible & " ActiveSheet=" & ThisWorkbook.ActiveSheet.Name
    modLog.LogError "E0801", "modUIMain.EnsureLayout", "[" & uiStep & "]" & diag, origNum
    If Not prevActive Is Nothing Then prevActive.Activate
    Application.ScreenUpdating = True
    On Error GoTo 0
    Err.Raise origNum, "modUIMain.EnsureLayout", "[" & uiStep & "] " & origDesc & diag
End Sub

' SetStage - ステータス行+Application.StatusBar 両方
Public Sub SetStage(ByVal msg As String)
    Dim displayMsg As String
    displayMsg = msg
    If LenB(displayMsg) = 0 Then displayMsg = "準備できています"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = modUIMainShape.GetHomeSheet()
    On Error GoTo 0

    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Range(RNG_STATUS).Value = "状態: " & displayMsg
        On Error GoTo 0
    End If

    On Error Resume Next
    Application.StatusBar = displayMsg
    On Error GoTo 0

    ' 空文字は「実況終わり」の合図。ここでバブルを消すと一瞬空白が出るので
    ' 触らない(生成中バブルは呼び出し側が本物の回答へ差し替える)。
    If LenB(msg) > 0 Then modLive.PaintStage msg

    If LenB(msg) > 0 Then ShowTip
End Sub

' RenderAnswer - 回答本文セル(SafeLeft)+出典ブロック+所要秒
Public Sub RenderAnswer(ByVal answerText As String, hits() As Hit, ByVal nHits As Long, _
                        ByVal mode As String, ByVal seconds As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = modUIMainShape.GetHomeSheet()
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ' Wave4修正: modAsk.Answerは空質問(未入力のまま「質問する」)のとき
    ' 検索を一切行わずmode=""で早期returnする契約にした(modAsk.bas参照)。
    ' mode=""は「実際の検索・回答生成が行われなかった」ことを示す唯一の
    ' 目印なので、このときは所要秒・出典欄(直前の質問の情報が残ったままに
    ' 見えてしまう)を付けず、案内文だけを表示する。mLastAnswerText(「Wordで
    ' 開く」の入力)も更新しない(案内文をWordで開いても意味がないため。
    ' 実際の回答のときだけこの下で更新する)。
    If LenB(mode) = 0 Then
        modUIMainShape.WriteSafe ws.Range(RNG_ANSWER), answerText
        modUIMainShape.WriteSafe ws.Range(RNG_SOURCES), ""

        On Error Resume Next
        ws.Range(RNG_STATUS).Value = "状態: 準備できています"
        On Error GoTo 0

        On Error Resume Next
        Application.StatusBar = False
        On Error GoTo 0
        Exit Sub
    End If

    mLastAnswerText = answerText

    Dim modeLabel As String
    If LCase$(mode) = "deep" Then
        modeLabel = "" & ChrW(&HD83D) & ChrW(&HDD0D) & " しっかり調べる"
    Else
        modeLabel = ChrW(&H26A1) & " すぐ聞く"
    End If

    Dim footer As String
    footer = vbLf & vbLf & "（" & modeLabel & " ・ 所要 " & modUtil.HumanSeconds(CDbl(seconds)) & "）"

    modUIMainShape.WriteSafe ws.Range(RNG_ANSWER), answerText & footer
    modUIMainShape.WriteSafe ws.Range(RNG_SOURCES), "" & ChrW(&HD83D) & ChrW(&HDCD6) & " この回答のもと: " & JoinSourceLabels(hits, nHits)

    On Error Resume Next
    ws.Range(RNG_STATUS).Value = "状態: 回答ができました。出典もあわせてご確認ください。"
    On Error GoTo 0

    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' RenderSourcesPreview - 出典先出し(ドラフト生成前に呼ぶ)。
'
' これはこのアプリで最も価値のある1秒。ここに来た時点で「どの資料に答えが
' あるか」は既に判明しており、あとはLLMが文章にするのを待つだけになっている。
' 従来はその事実を旧ホームシートのセルにだけ書いていたため、Nexusチャットを
' 見ている利用者には最後まで伝わらず、10～20秒がただの無反応だった。
' 実況先(生成中バブル)が預けられていれば、そちらへも必ず出す。
Public Sub RenderSourcesPreview(hits() As Hit, ByVal nHits As Long)
    ' 先に実況へ出す(こちらが利用者の見ている画面)。
    modLive.PaintSources hits, nHits

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = modUIMainShape.GetHomeSheet()
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim preview As String
    preview = "" & ChrW(&HD83D) & ChrW(&HDCC4) & " " & nHits & "件の資料がヒットしました: " & JoinSourceLabels(hits, nHits)
    modUIMainShape.WriteSafe ws.Range(RNG_SOURCES), preview
    modUIMainShape.WriteSafe ws.Range(RNG_ANSWER), "（資料を確認しました。ここから回答を作成します。もう少しお待ちください…）"
End Sub

' OnAskButton / OnModeQuick / OnModeDeep / OnOpenHowto / OnRunDiag
Public Sub OnAskButton()
    On Error GoTo Fail
    modAsk.AskFromUI
    RefreshBadgesAndDashboard
    Exit Sub

Fail:
    modLog.LogError "E0602", "modUIMain.OnAskButton", Err.Description
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub OnModeQuick()
    SetMode "quick"
End Sub

Public Sub OnModeDeep()
    SetMode "deep"
End Sub

Public Sub OnOpenHowto()
    On Error Resume Next
    ThisWorkbook.Worksheets(modAppDef.SH_HOWTO).Activate
    On Error GoTo 0
End Sub

' 2026-07-22実機報告対策: タブが隠れていてもNexus(チャット)へ戻れるように
' する(modUI.GoToNexusと同じ脱出路付き遷移をここから呼ぶだけ)。
Public Sub OnBackToChat()
    modUI.GoToNexus "modUIMain.OnBackToChat"
End Sub

Public Sub OnRunDiag()
    On Error GoTo Fail
    modDiag.RunDiagnostics
    AddDiagCopyErrorsButton
    Exit Sub
Fail:
    Err.Clear
    On Error GoTo 0
End Sub

' diag_reportシートに「📋直近のエラーをコピー」ボタンを追加(R1:基盤層は上位層を
' 呼ばないためUI操作はここで完結。シートは毎回作り直されるので毎回呼ぶ)。
Private Sub AddDiagCopyErrorsButton()
    On Error GoTo Fail
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("diag_report")
    modUIMainShape.AddButton ws, ws.Range("C1:D2"), "btn_diag_copy_errors", "" & ChrW(&HD83D) & ChrW(&HDCCB) & " 直近のエラーをコピー", "modUIMain.OnCopyRecentErrors"
    ' 2026-07-22実機報告対策: diag_reportに戻る導線が無くタブも消えていて
    ' 動けなくなっていた。チャットへの脱出路を追加。
    modUIMainShape.AddButton ws, ws.Range("E1:F2"), "btn_diag_back_chat", "" & ChrW(&HD83D) & ChrW(&HDCAC) & " チャットへ", "modUIMain.OnBackToChat"
    Exit Sub
Fail:
    Err.Clear
    On Error GoTo 0
End Sub

' 📋直近のエラーをコピー: err_logの直近5件を整形テキストにしてクリップボードへ。
Public Sub OnCopyRecentErrors()
    On Error GoTo Fail
    Dim text As String: text = modDiag.RecentErrorsForClipboard()
    If modClip.SetClipboardText(text) Then
        MsgBox "直近のエラーをコピーしました。そのまま貼り付けて共有してください。", _
               vbInformation, modAppDef.APP_NAME
    Else
        MsgBox "コピーに失敗しました。お手数ですが、diag_reportシートの内容を" & vbLf & _
               "スクリーンショットして共有してください。", vbExclamation, modAppDef.APP_NAME
    End If
    Exit Sub
Fail:
    Err.Clear
    On Error GoTo 0
End Sub

' ShowTip - 待ち時間豆知識(定型10本からRnd選択)
Public Sub ShowTip()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = modUIMainShape.GetHomeSheet()
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ' Wave6修正: Dim tips() As String に Variant配列(Array()の戻り値)を
    ' 直接代入すると、LibreOfficeでは黙って通るが実機Windows Excelでは
    ' 型不一致(実行時エラー13)になる(2026-07-16 実機E0801報告で確認)。
    ' 受け側をVariantにして暗黙変換を回避する。
    Dim tips As Variant
    tips = TipList()

    Randomize
    Dim idx As Long
    idx = Int(Rnd() * (UBound(tips) - LBound(tips) + 1)) + LBound(tips)

    modUIMainShape.WriteSafe ws.Range(RNG_TIP), "" & ChrW(&HD83D) & ChrW(&HDCA1) & " 豆知識: " & tips(idx)
End Sub

' OnFollowupButton - 「続けて質問」ボタン(裁定D11)。直近の回答を踏まえた
'   追加質問(深掘り)をInputBoxで受け取り、modAsk.AskFollowupへ渡す。
'   会話がまだ始まっていない(modAsk.CanFollowup=False)ときは丁寧な案内のみ。
'   文言はV2実証済みのmodChatUI.OnFollowupClickを踏襲(絵文字は使わない。§12)。
Public Sub OnFollowupButton()
    On Error GoTo Fail

    If Not modAsk.CanFollowup() Then
        MsgBox "まず質問して回答を受け取ってから使ってください。", _
               vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    Dim hint As String
    hint = "前回までの会話を踏まえて、追加の質問・深掘りを入力してください。" & vbCrLf & _
           "(何度でも続けられます。空欄のまま閉じると何もしません)"

    Dim followup As String
    followup = InputBox(hint, modAppDef.APP_NAME & " - 続けて質問", "")
    If LenB(Trim$(followup)) = 0 Then Exit Sub   ' キャンセル/空欄は何もしない

    modAsk.AskFollowup followup
    RefreshBadgesAndDashboard
    Exit Sub

Fail:
    modLog.LogError "E0602", "modUIMain.OnFollowupButton", Err.Description
    Err.Clear
    On Error GoTo 0
End Sub

' OnOpenWordButton - opt機能(Wordで開く)のUIラッパー(§7.7)。opt直接参照はしない。
'   裁定D12(対話型文書生成): 押下時に「どんな文書に仕上げるか」の指示文を
'   尋ね、直近回答テキストと合わせてmodFeatures.InvokeFeature経由で
'   optモジュール(ExportAnswerAsDoc)に渡す(引数は回答本文+指示文の2値を
'   Variant配列で。InvokeFeature→TryRibbonRunが展開する既存規約)。指示文が
'   空欄のときは整形せずそのままWordに転記される(opt側の契約)。
'   キャンセルと空欄OKは区別が必要(キャンセル=中止/空欄=そのまま転記)な
'   ため、VBAのInputBox(両者とも""が返り区別不能)ではなくApplication.InputBox
'   (Type:=2。キャンセル時はBooleanのFalseが返る)を使う。
Public Sub OnOpenWordButton()
    If LenB(mLastAnswerText) = 0 Then
        MsgBox "Wordで開く回答がありません。まず質問して、回答を受け取ってください。", _
               vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    Dim resp As Variant
    resp = Application.InputBox( _
        Prompt:="どんな文書に仕上げますか?" & vbCrLf & _
                "(例: お客様向けの回答文書風に / 社内回覧用の要約に)" & vbCrLf & _
                "※空欄ならそのまま転記", _
        Title:=modAppDef.APP_NAME & " - Wordで開く", Default:="", Type:=2)
    If VarType(resp) = vbBoolean Then Exit Sub   ' キャンセル→何もしない

    Dim instruction As String
    instruction = Trim$(CStr(resp))

    If LenB(instruction) > 0 Then
        SetStage "" & ChrW(&HD83D) & ChrW(&HDCDD) & " ご指定の形に整えて、Word文書を作成中…"
    Else
        SetStage "" & ChrW(&HD83D) & ChrW(&HDCDD) & " Word文書を作成中…"
    End If

    Dim result As Variant
    result = modFeatures.InvokeFeature("markdown", "ExportAnswerAsDoc", _
                                       Array(mLastAnswerText, instruction))
    SetStage ""

    ' ExportAnswerAsDocの契約: ""=成功 / "#ERR:..."=失敗(InvokeFeature側で
    ' "#ERR:FEATURE_UNAVAILABLE" に正規化される)。
    Dim isErr As Boolean
    isErr = False
    If VarType(result) = vbString Then
        If Left$(CStr(result), 5) = "#ERR:" Then isErr = True
    End If

    If isErr Then
        MsgBox "この機能は現在利用できません(管理者が有効化すると使えます)。", _
               vbInformation, modAppDef.APP_NAME
    End If
End Sub

' ShowEmptyShelfHint - 本棚が空のとき、回答エリアに常設案内を表示する。
'   呼び出し判断はmodBoot側の責務(本棚が空かどうかの判定はここではしない)。
Public Sub ShowEmptyShelfHint()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = modUIMainShape.GetHomeSheet()
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    modUIMainShape.WriteSafe ws.Range(RNG_ANSWER), _
        "まだ本棚に資料がありません。" & vbLf & _
        "まず『マイ本棚』タブで資料を1つ追加してみましょう " & ChrW(&H2192)
End Sub

' 内部ヘルパー

Private Sub RefreshBadgesAndDashboard()
    On Error Resume Next
    modStats.EvaluateBadges
    On Error GoTo 0
End Sub

Private Sub SetMode(ByVal mode As String)
    WriteUiStateMode mode
    ApplyModeColors mode
End Sub

Private Function CurrentMode() As String
    CurrentMode = "quick"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_UISTATE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row   ' xlUp
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(CStr(ws.Cells(i, 1).Value), "mode", vbTextCompare) = 0 Then
            Dim v As String
            v = LCase$(Trim$(CStr(ws.Cells(i, 2).Value)))
            If v = "deep" Then CurrentMode = "deep"
            Exit Function
        End If
    Next i
End Function

Private Sub WriteUiStateMode(ByVal mode As String)
    Dim ws As Worksheet
    Set ws = EnsureUiStateSheet()
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    Dim r As Long
    r = 0
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(CStr(ws.Cells(i, 1).Value), "mode", vbTextCompare) = 0 Then
            r = i
            Exit For
        End If
    Next i

    If r = 0 Then
        r = lastRow + 1
        If r < 1 Then r = 1
        ws.Cells(r, 1).Value = "mode"
    End If
    ws.Cells(r, 2).Value = mode
End Sub

Private Function EnsureUiStateSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_UISTATE)
    On Error GoTo 0

    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = modAppDef.SH_UISTATE
        ws.Cells(1, 1).Value = "key"
        ws.Cells(1, 2).Value = "value"
        On Error Resume Next
        ws.Visible = 2   ' xlSheetVeryHidden
        On Error GoTo 0
    End If
    Set EnsureUiStateSheet = ws
    Exit Function
Fail:
    Set EnsureUiStateSheet = Nothing
End Function

Private Sub ApplyModeColors(ByVal mode As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = modUIMainShape.GetHomeSheet()
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim quickSelected As Boolean
    quickSelected = (LCase$(mode) <> "deep")

    On Error Resume Next
    PaintModeButton ws, "btn_mode_quick", quickSelected
    On Error GoTo 0
    On Error Resume Next
    PaintModeButton ws, "btn_mode_deep", Not quickSelected
    On Error GoTo 0
End Sub

Private Sub PaintModeButton(ByVal ws As Worksheet, ByVal shapeName As String, ByVal selected As Boolean)
    Dim shp As Shape
    Set shp = ws.Shapes(shapeName)
    If selected Then
        shp.Fill.ForeColor.RGB = COLOR_SELECTED_BG
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_SELECTED_FG
    Else
        shp.Fill.ForeColor.RGB = modUIMainShape.COLOR_UNSELECTED_BG
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUIMainShape.COLOR_UNSELECTED_FG
    End If
End Sub

Private Function ModeCaption(ByVal mode As String) As String
    If mode = "deep" Then
        ModeCaption = "" & ChrW(&HD83D) & ChrW(&HDD0D) & " しっかり調べる (1～2分)"
    Else
        ModeCaption = ChrW(&H26A1) & " すぐ聞く (10～20秒)"
    End If
End Function

Private Sub EnsureQuestionName(ByVal ws As Worksheet)
    On Error Resume Next
    ThisWorkbook.Names(QUESTION_RANGE_NAME).Delete
    On Error GoTo 0

    On Error Resume Next
    ThisWorkbook.Names.Add Name:=QUESTION_RANGE_NAME, _
        RefersTo:="='" & ws.Name & "'!$" & Mid$(QUESTION_TOP_LEFT, 1, 1) & "$" & Mid$(QUESTION_TOP_LEFT, 2)
    On Error GoTo 0
End Sub

Private Function JoinSourceLabels(hits() As Hit, ByVal nHits As Long) As String
    If nHits <= 0 Then
        JoinSourceLabels = "(該当資料なし)"
        Exit Function
    End If

    Const MAX_SHOW As Long = 8
    Dim showCount As Long
    showCount = nHits
    If showCount > MAX_SHOW Then showCount = MAX_SHOW

    ' hits()はmodRetrieve.Search/modPromptsと同じ1-based配列(1 To nHits)。
    ' Wave4修正: ここが0-based(hits(0))になっていたため、ヒットが1件でも
    ' あると添字範囲エラー(実行時エラー9)でE0602に落ちていた。
    Dim parts() As String
    ReDim parts(1 To showCount)
    Dim i As Long
    For i = 1 To showCount
        parts(i) = OneSourceLabel(hits(i))
    Next i

    Dim joined As String
    joined = Join(parts, " / ")
    If nHits > MAX_SHOW Then
        joined = joined & " 他" & (nHits - MAX_SHOW) & "件"
    End If
    JoinSourceLabels = joined
End Function

Private Function OneSourceLabel(ByRef h As Hit) As String
    Dim label As String
    label = h.source & " p." & h.page
    If Left$(h.origin, 5) = "pack:" Then
        label = label & "（パック: " & Mid$(h.origin, 6) & "）"
    End If
    OneSourceLabel = label
End Function

Private Function TipList() As Variant
    TipList = Array( _
        "「" & ChrW(&HD83D) & ChrW(&HDD0D) & "しっかり調べる」は下書き" & ChrW(&H2192) & "検証の2段階なので少し時間がかかりますが、より丁寧な回答になります。", _
        "資料を追加すると本棚が育ち、答えられる質問がどんどん増えていきます。", _
        "資料は「マイ本棚」タブの「フォルダと同期」でまとめて自動追加できます。", _
        "回答の下にある出典を見れば、元の資料のどこに書いてあるかすぐ確認できます。", _
        "同じ資料を入れ直すと、自動的に新しい内容に置き換わります(重複しません)。", _
        "他の人が作った「パック」を取り込むと、自分で資料を集めなくても本棚が増やせます。", _
        "" & ChrW(&HD83D) & ChrW(&HDFE2) & "" & ChrW(&HD83D) & ChrW(&HDFE1) & "" & ChrW(&HD83D) & ChrW(&HDD34) & "のボタンで感想を送ると、ダッシュボードの記録に残ります。", _
        "困ったときは" & ChrW(&HD83E) & ChrW(&HDE7A) & "診断ボタンを押すと、今の状態が一目でわかります。", _
        "質問はできるだけ具体的に書くと、より的確な回答が返ってきます。", _
        "ダッシュボードでは、これまで取り戻した時間やバッジの獲得状況が見られます。" _
    )
End Function

