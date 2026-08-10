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
' 2026-08-01(R12-7-4・a11y監査Med): 押せる質問チップの文字を8.5ptへ広げた。
' 2026-08-05(R18-5a): カード化(アイコン+太字+説明1行)で2段ぶんの高さが
' 要る(26→46)。この+20ptは下に続く modHubStat.DrawInbox のY
' (DrawExtras内で CHIP_H から算出)へ自動で伝わる。
Private Const CHIP_H As Double = 46

' R18-3a/3b: Hub画面が使えるセル範囲の【上限】(A:L / 1..60行=900pt)。実際に
' 塗る範囲と ScrollArea は modViewport.BoundAddr が実下端から求める(R19-1b)。
' R21-S4: フッターの境界チェックも modViewport.RowAtFloor(窓高)由来へ移した
' ため、この Const の参照は契約用に残るだけになった。
Public Const HUB_BOUND As String = "A1:L60"

' R19-1b → R19H FA-4: 帯(吸収列Lを含む列帯)。右端は全てここから取る
' (帯625 / ピル617 / ナビ・フッター614.6 の3段ズレは、同じ「右端」を3通りに
' 計算していたのが原因。基準を1本へ統一した)。帯は rightPad=0(帯の右端
' そのもの)、操作系は rightPad=8(内側8pt)= Nexus と同じ作法。
Private Const HUB_BAND As String = "A1:L1"

' R21-S5: 行1の実高(ピル2段で2倍)。本文の起点はここが単一情報源 ―― 定数
' HDR_H を直に見ていたため、2段の端末ではプロフィールカードがヘッダーへ
' 潜り込んでいた。SY()は窓高適応圧縮(実体 modViewport2.SY)。
Private mHdrH As Double

Private Function HeaderH() As Double
    HeaderH = mHdrH
    If HeaderH < HDR_H Then HeaderH = HDR_H
End Function

Private Function SY(ByVal v As Double) As Double
    SY = modViewport2.SY(v)
End Function


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
    ' R18-1e: 強制終了でブックに焼き付いた進捗バナーの孤児を掃除(取込中は何もしない)。
    modProgressBar.SweepOrphans

    ws.Cells.Clear

    ' 幾何を確定させてからShapeを置く(順序が逆だと座標がズレる)。
    ' D列/G列の細い溝は「セル感」消し。タイルが隣接すると表に見えてしまう。
    ws.Columns("A").ColumnWidth = 1.5
    ws.Columns("B:C").ColumnWidth = 13
    ws.Columns("D").ColumnWidth = 1.2
    ws.Columns("E:F").ColumnWidth = 13
    ws.Columns("G").ColumnWidth = 2.5
    ws.Columns("H:K").ColumnWidth = 13
    ws.Columns("L").ColumnWidth = 1.5
    ' R20-1d(実機第7報⑦の層2): ここは Rows("1:60") だった。行高を明示した行は
    ' Excelから見れば「使用済み」=下スクロール域なので、フッターの実下端に
    ' 収まる40行だけを明示し、それ以深は描画後に ResetRowsBelow で既定へ戻す。
    ws.Rows("1:40").RowHeight = 15

    ' R21-S1: 活性化と表示状態(罫線/見出し/タブ/水平バー)を幾何を測る【前】へ
    ' 移した。従来は Fit と BoundAddr が Activate より先に走り、初回だけ
    ' 「前の画面の表示状態の窓」で決まっていた(水平バーぶん可視高が24pt過大)。
    If activate Then
        If Not modUI.ActivateSheetRobust(ws, "modHub.EnsureHubLayout") Then modUI.RestoreExcelUI
    End If
    modViewport2.EnsureViewState ws
    ' 表示の共通儀式(左端へ戻す/等倍/旧Vaultシートの掃除)。R4要件B。
    modKnowledge.PrepareScreenView ws
    Dim vw0 As Double, vh0 As Double
    modViewport2.MarkView vw0, vh0

    ' R19-1b(実機第6報①): 右の余白はスクロールではなく寸法の問題。余りを
    ' L列に吸わせて A:L の合計を可視幅ぴったりにすると、白い余白が構造的に
    ' 消える(ScrollAreaでは原理的に消せない)。Shape座標は全てセル幾何から
    ' 取っているので、列幅が変わっても追従する。
    modViewport.FitBandToViewport ws, HUB_BAND, "L"

    ' R18-3a(実機第5報②): 書式は ws.Cells(全域)ではなく実使用範囲だけに
    ' 当てる。全域に一様な書式を置くとExcelがそこまで「使用済み」と見なし、
    ' UsedRange(=スクロールできる範囲)がシート最大まで膨らむ ―― 「右にも
    ' 下にも無限にスクロールできる」の主因(調査agent2 §1.3)。
    ' R19-1b: フォントは描画前(このあとセルへ書く文字が拾うため)、塗りと
    ' ScrollAreaは描画後に実下端から決める(EnsureHubLayout末尾)。
    Dim baseAddr As String
    baseAddr = modViewport.BoundAddr(ws, "L", 0, 60)
    ws.Range(baseAddr).Font.Name = "Yu Gothic UI"
    ws.Range(baseAddr).Font.Size = 10
    ws.Range(baseAddr).Interior.Color = modUI.UiColor("bg")

    On Error Resume Next
    modTelemetry.TrackScreen "hub"
    On Error GoTo Fail

    DrawHeader ws
    ' R21-S5: ヘッダーの実高が決まってから、縦が窓に収まるかを判定する。
    ' 収まらないときだけ圧縮係数を立て、可変な縦寸法は SY() 経由で縮む。
    ' R21H F3: needを固定/可変に分けて渡す(まとめて渡すと分母が薄まり圧縮不足になる)。
    ' R21H F2-b: 右カラム(ナビ+チップ+お知らせ)はHubNeedYに一切算入されて
    ' おらず、圧縮も掛かっていなかった(狭窓で右カラムだけ境界外に残る原因の
    ' 一つ)。ナビとチップは可変、お知らせは最小44pt(modHubStat.DrawInboxの
    ' 下限)を固定分として加える(実文面はこれより伸びうるがDrawFooter側の
    ' クランプ検知(F2-a/c)が最終防波堤になる)。
    Dim rightVar As Double: rightVar = NAV_COUNT * (NAV_H + NAV_GAP) + CHIP_H + 14
    modViewport2.SetScaleY modViewport2.CompressFactor(modViewport.ViewportHeight(), _
        modViewport2.HubNeedYFixed(HeaderH()) + 44, _
        modViewport2.HubNeedYVariable(CARD_H, modHubStat.TilesHeight()) + rightVar)
    DrawProfileCard ws
    ' 0点のスコアボードを初見の人に見せない。全部ゼロのタイル8枚と鍵つき
    ' バッジ8個は「ここまで来た」ではなく「まだ何もしていない」としか読めない。
    ' 1問でも通してから出す(§DrawFirstStep)。
    ' R18-5b: 左右カラムの実下端を集めてフッターのYにする(固定Y禁止)。
    Dim botY As Double
    If HasAnyActivity() Then
        RefreshOrgTilesIfReachable
        modHubStat.DrawStatTiles ws, TilesTop()
        botY = DrawBadges(ws)
    Else
        DrawFirstStep ws
        botY = DrawBadges(ws)   ' 2-A: 初回からバッジ棚(welcome)を見せる
    End If
    DrawNavButtons ws
    Dim rightBot As Double: rightBot = DrawExtras(ws)
    If rightBot > botY Then botY = rightBot
    ' フッターの右端もナビ・ピルと同じ ContentRight(FA-4)。
    Dim footL As Double: footL = ws.Range("B1").Left
    Dim footBot As Double, trueBot As Double
    ' R20-1e: 余白 +18 → +10。左右カラムの実下端はどちらも要素の外枠なので、
    ' そこからさらに18pt空けるとフッターだけが1行ぶん浮いて見えた。
    footBot = modHubStat.DrawFooter(ws, footL, _
        modViewport.ContentRight(ws, HUB_BAND, 8) - footL, botY + modViewport2.SY(10), trueBot)

    ' R19-1b/R21-S4: 塗りと ScrollArea はフッターの実下端から決める(収まる
    ' 画面は窓高ちょうどへ切り下げ)。60行=900ptを常に塗ると余計に転がれる。
    ' R21H F2-a: footBotはDrawFooter内でクランプ済み(常に窓高以下)なので
    ' FitsInViewの判定にはクランプ前のtrueBotを渡す(F2-a/c)。
    Dim bnd As String
    bnd = modViewport.BoundAddr(ws, "L", trueBot, 60)
    ws.Range(bnd).Interior.Color = modUI.UiColor("bg")
    modViewport.ApplyScrollBound ws, bnd
    ' R20-1d: 境界の【下】に行高カスタムを1行も残さない(受け入れ基準)。
    ' bnd は必ず A1 起点なので、行数がそのまま下端行になる。
    On Error Resume Next
    modViewport.ResetRowsBelow ws, ws.Range(bnd).Rows.Count + 1, 200
    On Error GoTo Fail
    ' R21-S7: 窓幅/可視幅/帯実幅/中身右端/境界下端の5値観測点。
    modViewport2.LogFit ws, "hub", HUB_BAND, bnd

    On Error Resume Next
    modUI.FreezeShapePlacement ws
    On Error GoTo 0

    Application.ScreenUpdating = True
    ' R21-S1の保険: 描画中に窓が動いていたら1回だけ組み直す(ワンショット)。
    modViewport2.ReflowIfMoved vw0, vh0
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    modLog.LogError "E0801", "modHub.EnsureHubLayout", Err.Description, Err.Number
End Sub

' ヘッダーバー(全幅) + 右肩のユーティリティアイコン
' R11-B(#30): 固定30pt×6の決め打ちをやめ、modChrome.PillWidth+FlowRightで
' 可視幅基準に並べ直す。右端は HUB_BAND の ContentRight 1本(FA-4)。
Private Sub DrawHeader(ByVal ws As Worksheet)
    Dim L As Double, cellW As Double
    L = ws.Range("A1").Left
    ' 帯は帯の右端そのもの、操作系はその内側8pt(モジュール冒頭の HUB_BAND)。
    cellW = modViewport.ContentRight(ws, HUB_BAND, 0) - L
    Dim rightX As Double
    rightX = modViewport.ContentRight(ws, HUB_BAND, 8)

    ' 右→左に置く並び(終了が最も右)。旧実装の decrement 順をそのまま
    ' 配列順にした(icons(0)が最初に置かれる=右端)。
    Dim icons As Variant, acts As Variant, tips As Variant, labels As Variant
    icons = Array(ChrW(&HD83D) & ChrW(&HDEAA), ChrW(&H2753), _
                  ChrW(&HD83D) & ChrW(&HDCEE), ChrW(&HD83D) & ChrW(&HDD04), _
                  ChrW(&HD83C) & ChrW(&HDF19), ChrW(&HD83C) & ChrW(&HDF10))
    acts = Array("modHub.OnSaveAndExit", "modHub.OnHelp", _
                 "modHub.OnAnonFeedback", "modHub.OnRedraw", _
                 "modHub.OnThemeToggle", "modHub.OnLangCycle")
    tips = Array("保存して閉じる", "ヘルプ・使い方", "匿名で感想・要望を送る", _
                 "画面を描き直す", "配色を切り替える", "回答言語を切り替える")
    labels = Array("終了", "使い方", "ご意見", "再描画", "配色", "言語")

    Dim widths(0 To 5) As Double
    Dim i As Long
    For i = 0 To 5
        ' 2026-08-01(R12-7-3): ラベルを6pt→8.5ptにするため、幅見積りの
        ' pitch/padも実フォントに合わせて引き上げる(9pt換算のpitch。
        ' modUINexusDraw.HDR_PILL_PITCHと同じ考え方)。FlowRightが折返しで
        ' 吸収するため、幅が多少増えてもレイアウトは崩れない。
        widths(i) = modChrome.PillWidth(CStr(labels(i)), 9, 8, 30)
    Next i
    Dim xs() As Double, rws() As Long, useW() As Double
    Dim rowN As Long
    rowN = modChrome.FlowRight(widths, 6, rightX, L + 140, L + 8, 0, xs, rws, useW)
    If rowN < 1 Then rowN = 1
    Dim hdrH As Double: hdrH = rowN * HDR_H
    If hdrH > 200 Then hdrH = 200
    ws.Rows(1).RowHeight = hdrH
    mHdrH = hdrH          ' R21-S5: 本文の起点はこの実高から出す(定数ではなく)

    Dim hdr As Shape
    Set hdr = ws.Shapes.AddShape(5, L, 0, cellW, hdrH)
    hdr.Name = "nx_hub_hdr"
    hdr.Line.Visible = 0
    hdr.Adjustments(1) = 0.02
    hdr.Fill.ForeColor.RGB = modUI.UiColor("sidebar")
    modSkin.ApplyHeaderDepth hdr          ' §9: 濃紺の2色グラデーション
    With hdr.TextFrame2
        .TextRange.Text = ChrW(&H26A1) & " " & modAppDef.APP_NAME
        .TextRange.Font.Size = 14
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .MarginLeft = 14
        .VerticalAnchor = 3
    End With

    For i = 0 To 5
        Dim rowTop As Double: rowTop = rws(i) * HDR_H
        Dim xCenter As Double: xCenter = xs(i) + (useW(i) - 26) / 2
        On Error Resume Next
        Dim btn As Shape
        Set btn = ws.Shapes.AddShape(9, xCenter, rowTop + (HDR_H - 26) / 2, 26, 26)
        If Not btn Is Nothing Then
            btn.Name = "nx_hub_ic" & i
            ' R24-2a: 実機で白枠0.75ptは26pt円では視認距離で消えると判定。
            ' 塗りを白へ反転し、枠をPRIMARY_DARKへ(白地に白枠は無意味)。
            btn.Line.Visible = -1
            btn.Line.ForeColor.RGB = RGB(1, 77, 68)
            btn.Line.Weight = 0.75
            btn.Fill.ForeColor.RGB = RGB(255, 255, 255)
            With btn.TextFrame2
                .TextRange.Text = CStr(icons(i))
                .TextRange.Font.Size = 10
                .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .TextRange.ParagraphFormat.Alignment = 2
                .VerticalAnchor = 3
                .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            End With
            btn.OnAction = CStr(acts(i))
            btn.AlternativeText = CStr(tips(i))
        End If
        Set btn = Nothing
        Err.Clear
        On Error GoTo 0

        ' Excelの図形はマウスを乗せても代替テキストがツールチップとして
        ' 出ない。アイコンだけでは何のボタンか分からないという実機報告に
        ' 対し、真下に小さな文字ラベルを必ず添える(幅はスロット幅=useWに揃える)。
        On Error Resume Next
        Dim cap As Shape
        ' 2026-08-01(R12-7-3・a11y監査Med): 6pt(実表示≒8px)は高齢の営業所員に
        ' は判読不能。8.5ptへ拡大し、収容できる高さも10→12へ広げる
        ' (Zoom=100固定で拡大による自衛ができないため)。
        Set cap = ws.Shapes.AddShape(1, xs(i), rowTop + HDR_H - 12, useW(i), 12)
        If Not cap Is Nothing Then
            cap.Name = "nx_hub_icl" & i
            cap.Line.Visible = 0
            cap.Fill.Visible = 0
            With cap.TextFrame2
                .WordWrap = 0
                .TextRange.Text = CStr(labels(i))
                .TextRange.Font.Size = 8.5
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
    T = HeaderH() + 12

    Dim card As Shape
    Set card = ws.Shapes.AddShape(5, L, T, W, SY(CARD_H))
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
    ' R20-4b: 部門未設定は空欄のままにせず、クリックで設定できることを示す。
    Dim deptTxt As String
    deptTxt = IIf(LenB(dept) > 0, "  (" & dept & ")", "  (部門を設定:クリック)")
    With card.TextFrame2
        .WordWrap = -1
        .MarginLeft = 14: .MarginTop = 10: .MarginRight = 10
        .TextRange.Text = nm & deptTxt & vbCr & ContributionLine()
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
    If LenB(dept) = 0 Then card.OnAction = "modSetupWizard.OnDeptSetup"
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
    TilesTop = HeaderH() + 12 + SY(CARD_H) + SY(18)   ' プロフィールカードの下
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
    T = TilesTop()          ' R21-S5: 起点の式を2つ持たない

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
    StatTilesBottom = TilesTop() + SY(modHubStat.TilesHeight())
End Function

' ナビボタン4枚(右カラムG:Jの幾何に合わせる)
Private Sub DrawNavButtons(ByVal ws As Worksheet)
    Dim L As Double, W As Double, T As Double
    L = ws.Range("H3").Left
    ' ナビの右端はヘッダーピル・フッターと同じ ContentRight(FA-4)。
    W = modViewport.ContentRight(ws, HUB_BAND, 8) - L
    T = HeaderH() + 12

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
    ' R18-4: 「ナレッジ地図」は撤去したので説明からも外す(画面に無いものを
    ' 案内し続けると、探して見つからない人が必ず出る)。
    descs = Array("本棚の資料からAIが出典付きで回答 ・ " & chLbl, _
                  "資料の登録・検索・一覧 ・ 部内で配る/受け取る", _
                  "節約した時間・バッジ・EXP")
    acts = Array("modHub.OnGoChat", "modHub.OnGoVault", "modHub.OnGoDash")

    Dim i As Long
    For i = 0 To NAV_COUNT - 1
        ' R21H F2-b: ナビの縦寸法もSY()で圧縮に追随させる(狭窓で右カラムだけ
        ' 圧縮されず境界の外に取り残されていた)。
        Dim navTop As Double: navTop = T + i * SY(NAV_H + NAV_GAP)
        Dim btn As Shape
        Set btn = ws.Shapes.AddShape(5, L, navTop, W, SY(NAV_H))
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

' 質問カード3枚 + お知らせ(ナビボタンの直下に続けて置く)。
' 幾何(どこから始めるか)はHub側が持ち、カードの描画そのものは modHubStat
' へ置く(DrawStatTiles/DrawInbox と同じ切り分け。R18-5aのカード化で
' modHub が上限まで残り11字になったため。憲章§4-6)。
' 戻り値(R18-5b): 右カラムの実下端Y(お知らせの箱は文面で高さが変わる)。
Private Function DrawExtras(ByVal ws As Worksheet) As Double
    Dim L As Double, W As Double, T As Double
    L = ws.Range("H3").Left
    ' 2026-07-31(R7 A-3・タスク#27): ここは 4 * (NAV_H + NAV_GAP) だった。
    ' ナビを4枚から3枚へ減らしたときの後始末漏れで、実枚数と食い違った
    ' 1枚ぶん(54pt)がそのまま右列の空白として残っていた。実枚数から出す。
    W = modViewport.ContentRight(ws, HUB_BAND, 8) - L
    ' R20-1e: NAV_COUNT*(NAV_H+NAV_GAP) には最後のナビの下のGAP(10pt)が既に
    ' 入っているのに、さらに +10 して、置くときにも +20 していた。合計40ptの
    ' 空隙(ナビ下端260pt → カード上端300pt)が右カラムの頭に開いていた。
    ' 二重計上をやめ、GAP 1つぶんだけ空ける。
    T = HeaderH() + 12 + NAV_COUNT * SY(NAV_H + NAV_GAP)

    modHubStat.DrawQuickAskCards ws, L, W, T, SY(CHIP_H)
    DrawExtras = modHubStat.DrawInbox(ws, L, W, T + SY(CHIP_H) + SY(14))
End Function

' バッジ(セル。獲得済みは🏅、未獲得は🔒)。統計タイルはShapeに変えたので
' この帯だけが左ブロックのセル表示になる。
' 戻り値(R18-5b): 左カラムの実下端Y(バッジ件数で使う行数が変わる)。
Private Function DrawBadges(ByVal ws As Worksheet) As Double
    ' R20-1e【確定バグ】: ここは r = CLng(StatTilesBottom()/15) + 2 だった。
    ' 「行高15pt固定なので割り算で足りる」という前提が誤りで、Hubの行1は
    ' ヘッダー帯と同じ48pt(ナビが2段に折り返せば96pt)。15pt換算では行が
    ' 2〜3行ぶん下へずれ、統計タイルの下に約60ptの空隙が開き、しかも
    ' 戻り値(下端pt)も同じ式で出していたためフッターが実際のバッジ帯へ
    ' 食い込んでいた。pt→行は実測(modViewport.RowAt)だけを使う。
    Dim r As Long
    r = modViewport.RowAt(ws, StatTilesBottom(), 60) + 1   ' タイル下端の次の行から
    If r < 10 Then r = 10
    If r > 52 Then r = 52          ' 見出し+4行が HUB_BOUND(60行)を割らないように
    ' バッジが1件も無いときの下端(見出し行のみ)も実測から返す。
    DrawBadges = ws.Rows(r).Top + ws.Rows(r).Height

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
    If badgeN < 1 Then Exit Function

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

    Dim br As Long: br = modViewport2.BadgeRowsFor(modViewport2.ScaleY())
    With ws.Range("B" & (r + 1) & ":F" & (r + br))
        .Merge
        .WrapText = True
        .Value = sb
        .Font.Size = 8.5
        .Font.Color = modUI.UiColor("text")
        .VerticalAlignment = -4160
    End With
    ' 4行ぶんの帯の下端。ここも (r+4)*15 の机上換算をやめて実測にする
    ' (行1が48ptあるぶん、旧式は実際より約33pt上を返していた=フッターが
    '  バッジ帯に重なる)。
    DrawBadges = ws.Rows(r + br).Top + ws.Rows(r + br).Height
End Function

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
    modVaultGallery.ShowVaultGallery
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
            modLog.LogUsage "caller_mismatch", "modHub.OnQuickAsk", CStr(Application.Caller)
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
    If modUiLock.BlockIfIngesting() Then Exit Sub
    On Error Resume Next
    modApp.OnLangCycle
    modSkin.ShowToast "回答言語: " & modConfig.GetString("answer_language", "日本語"), "success"
    On Error GoTo 0
End Sub

' CycleSkin自身がmodUiLockを取るため非再入で素通し(先取りすると変色しない)。
' R14-6a(実機第3報 RC5-A): EnsureHubLayoutはHub/Nexusシートしか再彩色せず、
' ナレッジ画面やダッシュボードを表示中に押しても「押した画面」自体は
' 変わらないまま(「わからない」の実体)。ActiveSheetで分岐し、その画面
' 自身も描き直す(押した画面が必ず変わる。憲章§3-1)。
Public Sub OnThemeToggle()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If modUiLock.IsBusy() Then
        On Error Resume Next
        modSkin.ShowToast "処理中です。終わってから切り替えてください。", "info"
        On Error GoTo 0
        Exit Sub
    End If
    On Error Resume Next
    modSkin.CycleSkin
    EnsureHubLayout
    Dim activeName As String
    If Not ActiveSheet Is Nothing Then activeName = ActiveSheet.Name
    Select Case activeName
        Case modAppDef.SH_SHELF
            modKnowledge.RefreshCurrent
        Case modAppDef.SH_NEXUS_DASH
            modDash.ShowDashboard
    End Select
    On Error GoTo 0
End Sub

' OnHelpClick自身がmodUiLockを取るため非再入で素通し。先にNexusへ遷移する。
Public Sub OnHelp()
    If modUiLock.BlockIfIngesting() Then Exit Sub
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
' OnCheckUpdates - 受信箱の「押すと今すぐ確認します」(2026-07-31 R8 F3)。
'   起動シーケンス中は共有フォルダへ問い合わせない代わりに、利用者が
'   自分の意思で確認できる入口をここに置く。押した時点で TTL キャッシュを
'   捨て、問い合わせを許可してから Hub を描き直す。
Public Sub OnCheckUpdates()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modHubStat.AllowShareQueries
    modHubStat.InvalidatePending
    On Error GoTo 0
    modUiLock.Leave
    On Error Resume Next
    EnsureHubLayout
    On Error GoTo 0
End Sub

' R20-4c(実機第7報③④): 「config を開け」という本アプリでは実行不能な手順を
' 案内するのをやめ、実際に動く設定UI(modHelp.OnShareSetup)へ直接つなぐ。
Public Sub OnShareHelp()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    On Error Resume Next
    modHelp.OnShareSetup
    On Error GoTo 0
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
    If modUiLock.BlockIfIngesting() Then Exit Sub
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
        If LenB(mailUrl) > 0 Then
            modSkin.ShowToast "Officeの確認画面が出たら[はい]を押してください。", "info"
            ThisWorkbook.FollowHyperlink mailUrl
        End If
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
    If modUiLock.BlockIfIngesting() Then Exit Sub
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
    FeedbackMailto = "mailto:" & addr & "?subject=" & Replace(modAppDef.APP_NAME, " ", "%20") & "%20feedback"
    On Error GoTo 0
End Function
