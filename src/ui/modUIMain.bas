Attribute VB_Name = "modUIMain"
Option Explicit

' ============================================================================
' modUIMain - 「ホーム」画面(質問UI)の構築と描画(MASTER_SPEC §7.6/§8.1)
' ----------------------------------------------------------------------------
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
' ============================================================================

Private Const RNG_STATUS As String = "A12:H12"
Private Const RNG_ANSWER As String = "A14:H25"
Private Const RNG_SOURCES As String = "A26:H27"
Private Const RNG_TIP As String = "A33:H34"
Private Const QUESTION_TOP_LEFT As String = "A6"
Private Const QUESTION_RANGE_NAME As String = "mb_question"

Private Const COLOR_SELECTED_BG As Long = 2039071    ' RGB(31,78,120) 濃紺
Private Const COLOR_SELECTED_FG As Long = 16777215   ' RGB(255,255,255) 白
Private Const COLOR_UNSELECTED_BG As Long = 15921906 ' RGB(242,242,242) 薄灰
Private Const COLOR_UNSELECTED_FG As Long = 0        ' RGB(0,0,0) 黒

Private mLastAnswerText As String

' ----------------------------------------------------------------------------
' EnsureLayout - ホームを冪等再構築(既存Shapes全削除→再生成)
' ----------------------------------------------------------------------------
Public Sub EnsureLayout()
    Dim ws As Worksheet
    Set ws = GetOrCreateHomeSheet()
    If ws Is Nothing Then Exit Sub

    ' uiStep: 実機でだけ起きるエラー(型不一致など)を1回の報告で特定できるよう、
    ' modBoot.bootStageと同じ考え方でブロック単位の進捗を追う
    ' (2026-07-15 実機E0801「ホーム画面の組み立て」報告への対策)。
    Dim uiStep As String
    On Error GoTo Fail

    Application.ScreenUpdating = False

    ' 実機防衛(2026-07-21): 1004誘発要因(非表示/非アクティブ/表示オフ/保護)を
    ' 先回り無効化する多層防御(詳細は末尾ヘルパー参照。保険として維持)。
    Dim sdState As Variant, prevActive As Object
    uiStep = "描画環境の整備"
    SafeBeginDraw ws, sdState, prevActive

    uiStep = "既存ボタンの削除"
    RemoveManagedShapes ws
    DoEvents   ' 削除と追加の間でCOM/メモリを一拍解放する(先回り防衛#4)
    uiStep = "セルのクリア"
    ws.Cells.Clear

    uiStep = "既定フォント設定"
    ws.Cells.Font.Name = "游ゴシック"
    ws.Cells.Font.Size = 11

    ws.Columns("A:H").ColumnWidth = 12

    ' ---- タイトル行 ----------------------------------------------------
    uiStep = "タイトル行"
    With ws.Range("A1:E1")
        .Merge
        .Value = "📚 " & modAppDef.APP_NAME & "  v" & modAppDef.APP_VERSION
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

    uiStep = "ボタン(使い方/診断)"
    AddButton ws, ws.Range("F1:G2"), "btn_howto", "❓ 使い方", "modUIMain.OnOpenHowto"
    AddButton ws, ws.Range("H1:H2"), "btn_diag", "🩺 診断", "modUIMain.OnRunDiag"

    ' ---- モードトグル ----------------------------------------------------
    uiStep = "モードトグルボタン"
    ws.Rows("3:4").RowHeight = 20
    AddButton ws, ws.Range("A3:D4"), "btn_mode_quick", ModeCaption("quick"), "modUIMain.OnModeQuick"
    AddButton ws, ws.Range("E3:H4"), "btn_mode_deep", ModeCaption("deep"), "modUIMain.OnModeDeep"

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
    AddButton ws, ws.Range("C10:F11"), "btn_ask", "💬  質 問 す る", "modUIMain.OnAskButton"

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
        .Value = "（質問を入力して「💬 質問する」を押してください）"
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
    AddButton ws, ws.Range("A29:C30"), "btn_fb_green", "🟢 解決した!", "modAsk.FeedbackGreen"
    AddButton ws, ws.Range("D29:E30"), "btn_fb_yellow", "🟡 ヒントになった", "modAsk.FeedbackYellow"
    AddButton ws, ws.Range("F29:H30"), "btn_fb_red", "🔴 だめだった", "modAsk.FeedbackRed"

    ' ---- 続けて質問+Wordで開く ------------------------------------------------
    ' 「続けて質問」(裁定D11)はコア機能(modAsk)への入口なので常時生成する。
    ' 「Wordで開く」はopt機能(markdown)なのでFeatureEnabledがTrueのときだけ
    ' 生成する(§7.7)。※読み上げボタンは裁定D10で撤去した(音声読み上げは
    ' AIリボン本体でのみ利用可)。
    uiStep = "続けて質問ボタン"
    ws.Rows("31:32").RowHeight = 18
    AddButton ws, ws.Range("A31:D32"), "btn_followup", "💬 続けて質問", "modUIMain.OnFollowupButton"
    uiStep = "Wordで開くボタン(有効判定)"
    If modFeatures.FeatureEnabled("markdown") Then
        uiStep = "Wordで開くボタン(生成)"
        AddButton ws, ws.Range("E31:H32"), "btn_word", "📝 Wordで開く", "modUIMain.OnOpenWordButton"
    End If

    ' ---- 待ち時間豆知識 ----------------------------------------------------
    uiStep = "豆知識欄"
    With ws.Range(RNG_TIP)
        .Merge
        .Value = "💡 豆知識: "
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

    uiStep = "描画環境の復元"
    SafeEndDraw ws, sdState, prevActive

    Application.ScreenUpdating = True
    Exit Sub

Fail:
    ' どのブロックで型不一致等が起きたかをEnsureLayout自身が特定し、
    ' Descriptionに埋め込んでからmodBoot.Bootへ伝播させる(bootStageは
    ' 「ホーム画面の組み立て」としか分からないため、その内訳をここで補う)。
    Dim origNum As Long, origDesc As String
    origNum = Err.Number
    origDesc = Err.Description
    On Error Resume Next
    SafeEndDraw ws, sdState, prevActive   ' 失敗時も環境・表示状態を完全復元
    Application.ScreenUpdating = True
    On Error GoTo 0
    Err.Raise origNum, "modUIMain.EnsureLayout", "[" & uiStep & "] " & origDesc
End Sub

' ----------------------------------------------------------------------------
' SetStage - ステータス行+Application.StatusBar 両方
' ----------------------------------------------------------------------------
Public Sub SetStage(ByVal msg As String)
    Dim displayMsg As String
    displayMsg = msg
    If LenB(displayMsg) = 0 Then displayMsg = "準備できています"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetHomeSheet()
    On Error GoTo 0

    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Range(RNG_STATUS).Value = "状態: " & displayMsg
        On Error GoTo 0
    End If

    On Error Resume Next
    Application.StatusBar = displayMsg
    On Error GoTo 0

    If LenB(msg) > 0 Then ShowTip
End Sub

' ----------------------------------------------------------------------------
' RenderAnswer - 回答本文セル(SafeLeft)+出典ブロック+所要秒
' ----------------------------------------------------------------------------
Public Sub RenderAnswer(ByVal answerText As String, hits() As Hit, ByVal nHits As Long, _
                        ByVal mode As String, ByVal seconds As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetHomeSheet()
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
        WriteSafe ws.Range(RNG_ANSWER), answerText
        WriteSafe ws.Range(RNG_SOURCES), ""

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
        modeLabel = "🔍 しっかり調べる"
    Else
        modeLabel = "⚡ すぐ聞く"
    End If

    Dim footer As String
    footer = vbLf & vbLf & "（" & modeLabel & " ・ 所要 " & modUtil.HumanSeconds(CDbl(seconds)) & "）"

    WriteSafe ws.Range(RNG_ANSWER), answerText & footer
    WriteSafe ws.Range(RNG_SOURCES), "📖 この回答のもと: " & JoinSourceLabels(hits, nHits)

    On Error Resume Next
    ws.Range(RNG_STATUS).Value = "状態: 回答ができました。出典もあわせてご確認ください。"
    On Error GoTo 0

    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' RenderSourcesPreview - 出典先出し(ドラフト生成前に呼ぶ)
' ----------------------------------------------------------------------------
Public Sub RenderSourcesPreview(hits() As Hit, ByVal nHits As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetHomeSheet()
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim preview As String
    preview = "📄 " & nHits & "件の資料がヒットしました: " & JoinSourceLabels(hits, nHits)
    WriteSafe ws.Range(RNG_SOURCES), preview
    WriteSafe ws.Range(RNG_ANSWER), "（資料を確認しました。ここから回答を作成します。もう少しお待ちください…）"
End Sub

' ----------------------------------------------------------------------------
' OnAskButton / OnModeQuick / OnModeDeep / OnOpenHowto / OnRunDiag
' ----------------------------------------------------------------------------
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
    AddButton ws, ws.Range("C1:D2"), "btn_diag_copy_errors", "📋 直近のエラーをコピー", "modUIMain.OnCopyRecentErrors"
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

' ----------------------------------------------------------------------------
' ShowTip - 待ち時間豆知識(定型10本からRnd選択)
' ----------------------------------------------------------------------------
Public Sub ShowTip()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetHomeSheet()
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

    WriteSafe ws.Range(RNG_TIP), "💡 豆知識: " & tips(idx)
End Sub

' ----------------------------------------------------------------------------
' OnFollowupButton - 「続けて質問」ボタン(裁定D11)。直近の回答を踏まえた
'   追加質問(深掘り)をInputBoxで受け取り、modAsk.AskFollowupへ渡す。
'   会話がまだ始まっていない(modAsk.CanFollowup=False)ときは丁寧な案内のみ。
'   文言はV2実証済みのmodChatUI.OnFollowupClickを踏襲(絵文字は使わない。§12)。
' ----------------------------------------------------------------------------
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

' ----------------------------------------------------------------------------
' OnOpenWordButton - opt機能(Wordで開く)のUIラッパー(§7.7)。opt直接参照はしない。
'   裁定D12(対話型文書生成): 押下時に「どんな文書に仕上げるか」の指示文を
'   尋ね、直近回答テキストと合わせてmodFeatures.InvokeFeature経由で
'   optモジュール(ExportAnswerAsDoc)に渡す(引数は回答本文+指示文の2値を
'   Variant配列で。InvokeFeature→TryRibbonRunが展開する既存規約)。指示文が
'   空欄のときは整形せずそのままWordに転記される(opt側の契約)。
'   キャンセルと空欄OKは区別が必要(キャンセル=中止/空欄=そのまま転記)な
'   ため、VBAのInputBox(両者とも""が返り区別不能)ではなくApplication.InputBox
'   (Type:=2。キャンセル時はBooleanのFalseが返る)を使う。
' ----------------------------------------------------------------------------
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
        SetStage "📝 ご指定の形に整えて、Word文書を作成中…"
    Else
        SetStage "📝 Word文書を作成中…"
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

' ----------------------------------------------------------------------------
' ShowEmptyShelfHint - 本棚が空のとき、回答エリアに常設案内を表示する。
'   呼び出し判断はmodBoot側の責務(本棚が空かどうかの判定はここではしない)。
' ----------------------------------------------------------------------------
Public Sub ShowEmptyShelfHint()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetHomeSheet()
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    WriteSafe ws.Range(RNG_ANSWER), _
        "まだ本棚に資料がありません。" & vbLf & _
        "まず『マイ本棚』タブで資料を1つ追加してみましょう →"
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

Private Sub RefreshBadgesAndDashboard()
    On Error Resume Next
    modStats.EvaluateBadges
    On Error GoTo 0

    On Error Resume Next
    modUIDashboard.RenderDashboard
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
    Set ws = GetHomeSheet()
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
        shp.Fill.ForeColor.RGB = COLOR_UNSELECTED_BG
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_UNSELECTED_FG
    End If
End Sub

Private Function ModeCaption(ByVal mode As String) As String
    If mode = "deep" Then
        ModeCaption = "🔍 しっかり調べる (1〜2分)"
    Else
        ModeCaption = "⚡ すぐ聞く (10〜20秒)"
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
        "「🔍しっかり調べる」は下書き→検証の2段階なので少し時間がかかりますが、より丁寧な回答になります。", _
        "資料を追加すると本棚が育ち、答えられる質問がどんどん増えていきます。", _
        "資料は「マイ本棚」タブの「フォルダと同期」でまとめて自動追加できます。", _
        "回答の下にある出典を見れば、元の資料のどこに書いてあるかすぐ確認できます。", _
        "同じ資料を入れ直すと、自動的に新しい内容に置き換わります(重複しません)。", _
        "他の人が作った「パック」を取り込むと、自分で資料を集めなくても本棚が増やせます。", _
        "🟢🟡🔴のボタンで感想を送ると、ダッシュボードの記録に残ります。", _
        "困ったときは🩺診断ボタンを押すと、今の状態が一目でわかります。", _
        "質問はできるだけ具体的に書くと、より的確な回答が返ってきます。", _
        "ダッシュボードでは、これまで取り戻した時間やバッジの獲得状況が見られます。" _
    )
End Function

Private Sub WriteSafe(ByVal cell As Range, ByVal text As String)
    Dim t As String
    t = modUtil.SafeLeft(text, 32000)
    On Error Resume Next
    cell.Value = t
    On Error GoTo 0
End Sub

Private Function GetHomeSheet() As Worksheet
    On Error Resume Next
    Set GetHomeSheet = ThisWorkbook.Worksheets(modAppDef.SH_HOME)
    On Error GoTo 0
End Function

Private Function GetOrCreateHomeSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_HOME)
    On Error GoTo 0
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = modAppDef.SH_HOME
        On Error GoTo 0
    End If
    Set GetOrCreateHomeSheet = ws
    Exit Function
Fail:
    Set GetOrCreateHomeSheet = Nothing
End Function

' マイクロログ(2026-07-21): 失敗箇所をuiStepでerr_logへネスト伝播する。
Private Sub AddButton(ByVal ws As Worksheet, ByVal rng As Range, ByVal shapeName As String, _
                      ByVal caption As String, ByVal action As String)
    Dim uiStep As String
    On Error GoTo Fail
    Dim shp As Shape
    ' ログと実呼出しに同じサニタイズ済み値を使う(生値だとログと実態が食い違う)。
    Dim sL As Double, sT As Double, sW As Double, sH As Double
    sL = SafeCoord(rng.Left): sT = SafeCoord(rng.Top)
    sW = SafeCoord(rng.Width): sH = SafeCoord(rng.Height)
    uiStep = "AddShape実行(Type=5 L=" & sL & " T=" & sT & " W=" & sW & " H=" & sH & ")"
    Set shp = SafeRoundedRect(ws, sL, sT, sW, sH)
    uiStep = "図形名設定"
    shp.Name = shapeName
    uiStep = "テキスト代入"
    shp.TextFrame2.TextRange.Text = caption
    uiStep = "フォント設定"
    shp.TextFrame2.WordWrap = -1   ' msoTrue
    shp.TextFrame2.TextRange.Font.Size = 11
    shp.TextFrame2.TextRange.Font.Bold = -1   ' msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = 2   ' msoAlignCenter
    shp.TextFrame2.VerticalAnchor = 3   ' msoAnchorMiddle
    uiStep = "色/塗りつぶし設定"
    shp.Fill.ForeColor.RGB = COLOR_UNSELECTED_BG
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_UNSELECTED_FG
    shp.Line.Visible = 0   ' msoFalse
    uiStep = "OnAction割当て"
    shp.OnAction = action
    Exit Sub
Fail:
    Err.Raise Err.Number, "AddButton", "[" & uiStep & "] " & Err.Description
End Sub

' 実機防衛(2026-07-21): 図形描画の1004を誘発しうるExcel環境要因を先回りで
' 全て無効化する。#1シート保護解除 / #3オブジェクト表示ON / 非表示→表示化 /
' 非アクティブ→Activate試行 を退避しつつ実施し、SafeEndDrawで完全復元する。
' 例外的メンバー(DisplayObjects/Protect系)は遅延バインドでLO互換も確保。
' 局所化のためmodUIShelfにも同一実装を複製している。
Private Sub SafeBeginDraw(ByVal ws As Worksheet, ByRef st As Variant, ByRef prevActive As Object)
    Dim savedVis As Long: savedVis = -1
    Dim savedDisp As Variant: savedDisp = Empty
    Dim sheetUnprotected As Boolean: sheetUnprotected = False
    Dim wb As Object: Set wb = ThisWorkbook
    Dim o As Object: Set o = ws

    On Error Resume Next
    Set prevActive = wb.ActiveSheet
    On Error GoTo 0

    On Error Resume Next
    Dim win As Object: Set win = wb.Windows(1)
    If Not win Is Nothing Then
        savedDisp = win.DisplayObjects
        win.DisplayObjects = -4104   ' xlDisplayShapes
    End If
    On Error GoTo 0

    On Error Resume Next
    savedVis = ws.Visible
    If ws.Visible <> -1 Then ws.Visible = -1   ' xlSheetVisible
    On Error GoTo 0

    On Error Resume Next
    If o.ProtectContents Or o.ProtectDrawingObjects Then
        o.Unprotect
        If Not (o.ProtectContents Or o.ProtectDrawingObjects) Then sheetUnprotected = True
    End If
    On Error GoTo 0

    On Error Resume Next
    ws.Activate
    On Error GoTo 0

    st = Array(savedVis, savedDisp, sheetUnprotected)
End Sub

Private Sub SafeEndDraw(ByVal ws As Worksheet, ByVal st As Variant, ByVal prevActive As Object)
    If Not IsArray(st) Then Exit Sub
    Dim savedVis As Long: savedVis = CLng(st(0))
    Dim savedDisp As Variant: savedDisp = st(1)
    Dim sheetUnprotected As Boolean: sheetUnprotected = CBool(st(2))
    Dim wb As Object: Set wb = ThisWorkbook
    Dim o As Object: Set o = ws

    On Error Resume Next
    If sheetUnprotected Then o.Protect
    On Error GoTo 0

    On Error Resume Next
    If Not prevActive Is Nothing Then prevActive.Activate
    On Error GoTo 0

    On Error Resume Next
    If ws.Visible <> savedVis Then ws.Visible = savedVis
    On Error GoTo 0

    On Error Resume Next
    If Not IsEmpty(savedDisp) Then
        Dim win As Object: Set win = wb.Windows(1)
        If Not win Is Nothing Then win.DisplayObjects = savedDisp
    End If
    On Error GoTo 0
End Sub

' #5座標サニタイズ(負/0→1)+ #4 DoEvents1回リトライ付きで角丸四角を追加。
Private Function SafeCoord(ByVal v As Double) As Double
    If v < 1 Then v = 1
    SafeCoord = v
End Function

' 呼び出し元がSafeCoordで既に安全化した値を渡す前提だが、直接呼ばれても
' 壊れないよう二重にクランプする(コストはほぼ無い)。
Private Function SafeRoundedRect(ByVal ws As Worksheet, ByVal L As Double, ByVal T As Double, _
                                 ByVal W As Double, ByVal H As Double) As Shape
    L = SafeCoord(L): T = SafeCoord(T): W = SafeCoord(W): H = SafeCoord(H)
    On Error GoTo Retry
    Set SafeRoundedRect = ws.Shapes.AddShape(5, L, T, W, H)   ' 5=msoShapeRoundedRectangle(リテラル)
    Exit Function
Retry:
    DoEvents
    Set SafeRoundedRect = ws.Shapes.AddShape(5, L, T, W, H)
End Function

Private Sub RemoveManagedShapes(ByVal ws As Worksheet)
    Dim names() As String
    ReDim names(0 To ws.Shapes.Count)
    Dim n As Long
    n = 0

    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 4) = "btn_" Or Left$(shp.Name, 4) = "lbl_" Then
            names(n) = shp.Name
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
