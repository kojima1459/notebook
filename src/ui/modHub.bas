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

Private Const HDR_H As Double = 48
Private Const CARD_H As Double = 68
' 2026-07-31(R7 A-3): ナビを4枚→3枚に減らしたときに縦幅の後始末が漏れ、
' 右列の下半分が丸ごと余白になっていた。3枚で従来4枚分をほぼ埋める寸法へ。
' (46*4+8*3=208 → 60*3+10*2=200。説明文の窮屈さも同時に解消する)
Private Const NAV_H As Double = 60
Private Const NAV_GAP As Double = 10
Private Const NAV_COUNT As Long = 3
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
    modHubStat.RemoveHubShapes ws

    ws.Cells.Clear
    ws.Cells.Font.Name = "Yu Gothic UI"
    ws.Cells.Font.Size = 10
    ws.Cells.Interior.Color = modUI.UiColor("bg")

    ' 幾何を確定させてからShapeを置く(順序が逆だと座標がズレる)。
    ' D列/G列の細い溝は「セル感」消し。タイルが隣接すると表に見えてしまう。
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

    ' 表示の共通儀式(左端へ戻す/等倍/旧Vaultシートの掃除)。R4要件B。
    modKnowledge.PrepareScreenView ws

    On Error Resume Next
    modTelemetry.TrackScreen "hub"
    On Error GoTo Fail

    DrawHeader ws
    DrawProfileCard ws
    ' 0点のスコアボードを初見の人に見せない。全部ゼロのタイル8枚と鍵つき
    ' バッジ8個は「ここまで来た」ではなく「まだ何もしていない」としか読めない。
    ' 1問でも通してから出す(§DrawFirstStep)。
    If HasAnyActivity() Then
        RefreshOrgTilesIfReachable
        modHubStat.DrawStatTiles ws, TilesTop()
        DrawBadges ws
    Else
        DrawFirstStep ws
        DrawBadges ws        ' 2-A: 初回からバッジ棚(welcome)を見せる
    End If
    DrawNavButtons ws
    DrawExtras ws

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
    ' 🔄=画面を再描画(Alt+Tab復帰等の崩れの1クリック復旧)。
    ' 📮=匿名の投書箱。実名だと率直な意見は出てこないので別経路を用意する。
    icons = Array(ChrW(&HD83C) & ChrW(&HDF10), ChrW(&HD83C) & ChrW(&HDF19), _
                  ChrW(&HD83D) & ChrW(&HDD04), ChrW(&HD83D) & ChrW(&HDCEE), _
                  ChrW(&H2753), ChrW(&HD83D) & ChrW(&HDEAA))
    acts = Array("modHub.OnLangCycle", "modHub.OnThemeToggle", _
                 "modHub.OnRedraw", "modHub.OnAnonFeedback", _
                 "modHub.OnHelp", "modHub.OnSaveAndExit")
    tips = Array("回答言語を切り替える", "配色を切り替える", "画面を描き直す", _
                 "匿名で感想・要望を送る", "ヘルプ・使い方", "保存して閉じる")
    Dim shortLabels As Variant
    shortLabels = Array("言語", "配色", "再描画", "ご意見", "使い方", "終了")

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

        ' Excelの図形はマウスを乗せても代替テキストがツールチップとして
        ' 出ない。アイコンだけでは何のボタンか分からないという実機報告に
        ' 対し、真下に小さな文字ラベルを必ず添える。
        On Error Resume Next
        Dim cap As Shape
        Set cap = ws.Shapes.AddShape(1, xRight - 10, HDR_H - 11, 46, 10)
        If Not cap Is Nothing Then
            cap.Name = "nx_hub_icl" & i
            cap.Line.Visible = 0
            cap.Fill.Visible = 0
            With cap.TextFrame2
                .WordWrap = 0
                .TextRange.Text = CStr(shortLabels(i))
                .TextRange.Font.Size = 6
                .TextRange.Font.Fill.ForeColor.RGB = RGB(190, 210, 235)
                .TextRange.ParagraphFormat.Alignment = 2
                .VerticalAnchor = 3
                .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            End With
        End If
        Set cap = Nothing
        Err.Clear
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
    On Error Resume Next
    nm = Trim$(modConfig.GetString("pack_author", ""))
    dept = Trim$(modConfig.GetString("user_department", ""))
    On Error GoTo 0
    If LenB(nm) = 0 Or nm = "名称未設定" Then nm = "ゲスト ユーザー"

    ' 2行目は「Lv.3  EXP 240」だった。琥珀色のゲージ付きで、Hubのいちばん
    ' 目立つ位置に常時出ていた。
    ' 42歳の課長が部下の画面を横から見て、レベルと経験値バーが載っている
    ' 業務ツールをどう判定するか ―― 「若手向けのおもちゃ」で終わる。
    ' 「すごいですね」と言われて使われない、あの反応の一因がここにある。
    '
    ' JTCで実際に価値のある通貨は、点数ではなく「同僚の役に立った事実」。
    ' 幸い thanks_received_total は既に集計されている(modP2P)。
    ' 同じ場所に、点数の代わりにそれを置く。こちらは自慢しても角が立たない。
    ' EXP自体は内部に残す(きせかえの解放条件に使っている)。表に出さないだけ。
    ' Paragraphsで書式を分けるため区切りはvbCr(vbLfだと1段落のまま)。
    With card.TextFrame2
        .WordWrap = -1
        .MarginLeft = 14: .MarginTop = 10: .MarginRight = 10
        .TextRange.Text = nm & IIf(LenB(dept) > 0, "  (" & dept & ")", "") & vbCr & _
            ContributionLine()
        .TextRange.Font.Size = 11
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
        On Error Resume Next
        .TextRange.Paragraphs(1).Font.Bold = -1
        .TextRange.Paragraphs(1).Font.Size = 12
        ' 2026-07-31(発見事項3): 段落数を数えてから触る。テキストが箱の
        ' 高さを超えて2段落目がはみ出す環境では、無条件の Paragraphs(2)
        ' アクセスがOn Error Resume Nextに握られ「気づかれない失敗」になる
        ' (A-4のHubタイルと同型の脆弱性)。
        If .TextRange.Paragraphs.Count >= 2 Then
            .TextRange.Paragraphs(2).Font.Size = 9
            .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
        End If
        On Error GoTo 0
    End With
End Sub

' プロフィール2行目。点数ではなく、その人がどう役に立ったかを書く。
' まだ何も無い人には、責める言葉にならないよう所属だけ/空にする。
Private Function ContributionLine() As String
    Dim thanks As Long, solved As Long
    thanks = modHubStat.SafeStat("thanks_received_total")
    solved = modHubStat.SafeStat("selfsolve_total")

    If thanks > 0 Then
        ContributionLine = ChrW(&HD83C) & ChrW(&HDF31) & _
            " あなたが入れた資料で、これまで " & thanks & "人が解決しました"
    ElseIf solved > 0 Then
        ContributionLine = ChrW(&H2705) & " これまで " & solved & "件を自分で解決しました"
    Else
        ContributionLine = "資料を入れて質問すると、ここに記録が残ります"
    End If
End Function


' 要件D(2026-07-30 R3): 「みんな(今日/今月)」タイルはmodBoard.BootBoard
' (LaunchNexus内=EnsureHubLayoutよりさらに後)で初めて集計されるため、
' 初回描画時点では前回セッションの値のまま(Hubは明示的に再描画されるまで
' そのまま。R3要件定義書 背景1)。共有フォルダに届く見込みがあるとき
' (modShare.Reachable。1セッション1回のキャッシュ判定なので、届かないと
' 分かった後の追加コストは実質ゼロ)だけ、タイル描画の直前に集計をやり直す。
' 未設定・到達不能なら呼ばず0表示のままにする(ブロッキングさせない)。
Private Sub RefreshOrgTilesIfReachable()
    On Error Resume Next
    If modShare.Reachable() Then modBoard.RefreshBoardTiles
    On Error GoTo 0
End Sub

' 統計タイル8枚の描画本体は modHubStat へ移した(2026-07-31 R7 A-4)。
' ここは「どこから始めるか」だけを渡す(幾何の持ち主はHub側のまま)。
Private Function TilesTop() As Double
    TilesTop = HDR_H + 12 + CARD_H + 18       ' プロフィールカードの下
End Function

' まだ何も起きていない状態か。質問も取込も0のときだけ「初回」とみなす。
Private Function HasAnyActivity() As Boolean
    If modHubStat.AskTotal() > 0 Then
        HasAnyActivity = True
        Exit Function
    End If
    HasAnyActivity = (modHubStat.SafeChunks() > 0)
End Function

' 初回のHub左半分。数字の代わりに、次にやる1つのことだけを大きく置く。
' ゲーミフィケーションは実績を語る道具で、実績ゼロの相手には逆に働く。
Private Sub DrawFirstStep(ByVal ws As Worksheet)
    Dim L As Double, W As Double, T As Double
    L = ws.Range("B1").Left
    W = ws.Range("B1:F1").Width
    T = HDR_H + 12 + CARD_H + 18

    On Error Resume Next
    Dim card As Shape
    Set card = ws.Shapes.AddShape(5, L, T, W, 128)
    If card Is Nothing Then Exit Sub
    card.Name = "nx_hub_first"
    card.Adjustments(1) = 0.06
    card.Line.Visible = -1
    card.Line.Weight = 1.25
    card.Line.ForeColor.RGB = modUI.UiColor("accent")
    card.Fill.ForeColor.RGB = modUI.UiColor("surface")
    modSkin.ApplyLightShadow card
    ' 段落で書式を分けるため区切りはvbCr(vbLfだとParagraphs(2)が範囲外)。
    With card.TextFrame2
        .WordWrap = -1
        .MarginLeft = 16: .MarginRight = 14: .MarginTop = 14: .MarginBottom = 10
        .TextRange.Text = _
            ChrW(&HD83D) & ChrW(&HDCAC) & " まず、1つ聞いてみてください" & vbCr & _
            "知りたいことを、ふだんの言葉のまま書くだけです。" & vbLf & _
            "資料をまだ入れていなくても、そのまま答えます。" & vbLf & vbLf & _
            "約款やマニュアルを入れると、「どの資料の何ページか」まで" & vbLf & _
            "付けて答えられるようになります。"
        .TextRange.Font.Size = 9.5
        .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
        On Error Resume Next
        .TextRange.Paragraphs(1).Font.Size = 13
        .TextRange.Paragraphs(1).Font.Bold = -1
        .TextRange.Paragraphs(1).Font.Fill.ForeColor.RGB = modUI.UiColor("text")
        On Error GoTo 0
    End With
    card.OnAction = "modHub.OnGoChat"
    Set card = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

' 統計タイル群の下端(バッジ等をその下に置くために使う)。
' 高さは描いている側(modHubStat)から取る。両方に数字を持つとズレる。
Private Function StatTilesBottom() As Double
    StatTilesBottom = TilesTop() + modHubStat.TilesHeight()
End Function

' ナビボタン4枚(右カラムG:Jの幾何に合わせる)
Private Sub DrawNavButtons(ByVal ws As Worksheet)
    Dim L As Double, W As Double, T As Double
    L = ws.Range("H3").Left
    W = ws.Range("H3:K3").Width
    T = HDR_H + 12

    Dim caps As Variant, acts As Variant, descs As Variant
    ' 2026-07-30(R4要件F): 「ナレッジ倉庫」と「マイ本棚」を1枚に統合した。
    ' 中身は元から同じデータ(modShelf.SourceList)で、違いは見せ方だけ
    ' だったのに、扉を2つ並べていたせいで「どっちに入れた資料か」を
    ' 利用者が悩んでいた。行き先も同じ1枚のシートになった。
    caps = Array(ChrW(&HD83D) & ChrW(&HDCAC) & " チャットで質問する", _
                 ChrW(&HD83D) & ChrW(&HDCDA) & " ナレッジと本棚", _
                 ChrW(&HD83D) & ChrW(&HDCCA) & " ダッシュボード")
    Dim chLbl As String
    On Error Resume Next
    chLbl = modChannel.ActiveLabel()
    On Error GoTo 0
    descs = Array("本棚の資料からAIが出典付きで回答 ・ " & chLbl, _
                  "資料の登録・検索・一覧 ・ 部内で配る/受け取る", _
                  "バッジ・EXP・ナレッジ地図")
    acts = Array("modHub.OnGoChat", "modHub.OnGoVault", "modHub.OnGoDash")

    Dim i As Long
    For i = 0 To NAV_COUNT - 1
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
            ' 2026-07-31(発見事項3): 段落数チェック(理由はDrawProfileCard参照)。
            If .TextRange.Paragraphs.Count >= 2 Then
                .TextRange.Paragraphs(2).Font.Size = 8
                .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
            End If
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
    ' 2026-07-31(R7 A-3・タスク#27): ここは 4 * (NAV_H + NAV_GAP) だった。
    ' ナビを4枚から3枚へ減らしたときの後始末漏れで、実枚数と食い違った
    ' 1枚ぶん(54pt)がそのまま右列の空白として残っていた。実枚数から出す。
    W = ws.Range("H3:K3").Width
    T = HDR_H + 12 + NAV_COUNT * (NAV_H + NAV_GAP) + 10

    ' 見出しは置かない。チップの文面自体が「こう聞けばいい」の見本になっている。
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

    DrawInbox ws, L, W, T + 20 + CHIP_H + 14
End Sub

' 共有知のお知らせ。通知が無い共有機能は使われないので、件数と行き先を出す。
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

    ' 2026-07-28(解説書 §11-11): バッジ表を自前で持たない。
    ' 判定している modStats から受け取る(表を2つ持つと必ずズレる。
    ' 実際、共有知に最も貢献した4種がここに無かったため、獲得しても
    ' 本人には一生見えなかった)。
    Dim ids() As String, titles() As String, longs_() As String, conds() As String
    Dim badgeN As Long
    badgeN = modStats.BadgeCatalog(ids, longs_, titles, conds)
    If badgeN < 1 Then Exit Sub

    Dim sb As String
    Dim i As Long
    For i = 0 To badgeN - 1
        Dim mark As String
        If LenB(modStats.BadgeEarnedOn(CStr(ids(i)))) > 0 Then
            mark = ChrW(&HD83C) & ChrW(&HDFC5)
        Else
            mark = ChrW(&HD83D) & ChrW(&HDD12)
        End If
        sb = sb & mark & " " & CStr(titles(i)) & "   "
    Next i

    With ws.Range("B" & (r + 1) & ":F" & (r + 4))
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
    If modUiLock.BlockIfIngesting() Then Exit Sub   ' R7 B-2
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUI.GoToNexus "modHub.OnGoChat"
    On Error GoTo 0
    modUiLock.Leave
End Sub

' 📚 ナレッジと本棚 = マイ本棚シートのギャラリーモード(R4要件F)。
' 一覧表・みんなの解決事例へは、開いた先の上部ピルで切り替える。
' 旧OnGoShelf(一覧表を直接開く導線)は、行き先が同じ1枚になったので廃止した。
Public Sub OnGoVault()
    If modUiLock.BlockIfIngesting() Then Exit Sub   ' R7 B-2
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modVault.ShowVaultGallery
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnGoDash()
    If modUiLock.BlockIfIngesting() Then Exit Sub   ' R7 B-2
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modDash.ShowDashboard
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnQuickAsk()
    If modUiLock.BlockIfIngesting() Then Exit Sub   ' R7 B-2
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

Public Sub OnLangCycle()
    On Error Resume Next
    modApp.OnLangCycle
    modSkin.ShowToast "回答言語: " & modConfig.GetString("answer_language", "日本語"), "success"
    On Error GoTo 0
End Sub

' modSkin.CycleSkin 自身が modUiLock を取る。modUiLockは非再入なので、
' ここで先に取ると内側のEnterがFalseになり、配色が一切変わらないまま
' ステータスバーの「処理中です...」だけが残る(実機で再現した不具合)。
' 素通しにして、CycleSkin の完了後にHubを描き直す。
Public Sub OnThemeToggle()
    ' 2026-07-28(レビュー M-26): 処理中の再描画は画面を壊す。
    If modUiLock.IsBusy() Then
        On Error Resume Next
        modSkin.ShowToast "処理中です。終わってから切り替えてください。", "info"
        On Error GoTo 0
        Exit Sub
    End If
    On Error Resume Next
    modSkin.CycleSkin
    EnsureHubLayout
    On Error GoTo 0
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
               "いただいた声は改善に使わせていただきます。", _
               vbInformation, modAppDef.APP_NAME
    Else
        ' 共有フォルダに書けない環境ではメール経路へ逃がす(黙って捨てない)。
        On Error Resume Next
        modClip.SetClipboardText fb
        Dim mailUrl As String: mailUrl = FeedbackMailto()
        If LenB(mailUrl) > 0 Then ThisWorkbook.FollowHyperlink mailUrl
        On Error GoTo 0
        If LenB(mailUrl) > 0 Then
            MsgBox "共有フォルダへ送れなかったため、メールの下書きを開きました。" & vbCrLf & _
                   "本文はクリップボードに入っています(Ctrl+V で貼り付けてください)。", _
                   vbInformation, modAppDef.APP_NAME
        Else
            MsgBox "共有フォルダへ送れませんでした。" & vbCrLf & _
                   "本文はクリップボードに入っています(Ctrl+V で貼り付けて、" & vbCrLf & _
                   "管理者へお送りください)。", vbInformation, modAppDef.APP_NAME
        End If
    End If
End Sub

Public Sub OnRedraw()
    On Error Resume Next
    modApp.OnRefreshUI
    On Error GoTo 0
End Sub

' ---- 内部ヘルパー ----

' 問い合わせ先のメールアドレス。2026-07-28(レビュー H-17): 個人の
' メールアドレスがソースへ直書きされていた。担当が変わるたびに再ビルドが
' 要るうえ、退職・異動で宛先が死ぬ。config へ出す(既定は空。空のときは
' メール経路そのものを出さない)。
Private Function FeedbackMailto() As String
    On Error Resume Next
    Dim addr As String
    addr = Trim$(modConfig.GetString("feedback_mail_to", ""))
    If LenB(addr) = 0 Then Exit Function
    FeedbackMailto = "mailto:" & addr & "?subject=Nexus%20Agent%20feedback"
    On Error GoTo 0
End Function
