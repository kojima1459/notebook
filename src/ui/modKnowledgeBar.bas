Attribute VB_Name = "modKnowledgeBar"
Option Explicit

' ============================================================================
' modKnowledgeBar - ナレッジ画面のツールバー(何を出すか+ボタン1個の描画)。
' ----------------------------------------------------------------------------
' 2026-07-31(R11-F1): modKnowledge が30,000字上限まで残り110字となり、修正が
'   1件も入らない状態だったため、憲章§4-6に基づき分割した。
'   切り口は「画面の骨組み(modKnowledge: ヘッダー/右肩ピル/モード管理と
'   各ボタンのハンドラ)」と「ツールバーの並びと描画(本モジュール)」。
'   ToolbarSpec が「どのボタンを出すか」の唯一の場所であり、DrawChrome からは
'   DrawToolbar 1本しか呼ばれない。ボタンの OnAction 文字列は
'   modKnowledge 側のハンドラを指したまま(ハンドラは移していない)。
' ============================================================================

Public Const BAR_H As Double = 24       ' ツールバー1段の高さ(DrawChromeが行3の高さに使う)
Private Const TB_GAP As Double = 5      ' ツールバーのボタン間隔
Private Const TB_PAD As Double = 8      ' ツールバー帯の左右余白
Private Const TB_MAX As Long = 20       ' ツールバーに載りうるボタンの最大数

' R20-3(実機第7報②): 起動後の本棚初回表示で1回だけ「未仕上げ資料あり」を
' 告知したかどうか(セッション内で1回きり。ブックを開き直すと再び案内する)。
Private mBackfillToastDone As Boolean

' ----------------------------------------------------------------------------
' DrawToolbar - ツールバーを1本の流し込みレイアウトで描き、実使用高さを返す。
' ----------------------------------------------------------------------------
' 旧実装は「通常ボタン群」「正典を発行」「利用状況」「スクショ取込」の4箇所で
' 別々に折り返し判定を書いており、そのすべてが `rowIdx = 0` 限定だった。
' つまり2段目に入った瞬間から幅を一切見ずに右へ描き続けていた
' (全ボタン有効な端末では合計約1,068pt=帯幅の約1.8倍)。
' ボタンの並びを1本の配列にまとめ、段数無制限の流し込み(modChrome.FlowLeft)
' へ一本化する。判定が1箇所になったので「ここだけ直し忘れる」が起きない。
Public Function DrawToolbar(ByVal ws As Worksheet, ByVal isTable As Boolean, _
                             ByVal isShared As Boolean, ByVal L As Double, _
                             ByVal W As Double, ByVal barTop As Double) As Double
    Dim caps() As String, acts() As String, kinds() As String, tips() As String
    Dim n As Long
    Dim xs() As Double, rws() As Long, useW() As Double
    Dim rowN As Long
    ComputeToolbarLayout isTable, isShared, L, W, caps, acts, kinds, tips, n, xs, rws, useW, rowN
    MaybeShowBackfillToast isShared
    If n < 1 Then Exit Function

    Dim i As Long
    For i = 0 To n - 1
        ToolButton ws, "nxk_tb" & i, caps(i), acts(i), kinds(i), _
                   xs(i), barTop + rws(i) * (BAR_H + 2), useW(i), tips(i)
    Next i

    ' 実使用高さ。呼び出し元(DrawChrome)がこれを行3の高さに入れるので、
    ' 下の検索欄・カード領域は段数に応じて自動で下がる。
    DrawToolbar = rowN * (BAR_H + 2)
End Function

' ----------------------------------------------------------------------------
' ToolbarContentRight - ツールバーの実際の右端X(pt)。Shapeを一切生成せず、
'   DrawToolbarと同じ算数(ToolbarSpec+modChrome.FlowLeft)だけを走らせる。
' ----------------------------------------------------------------------------
' なぜ必要か(実機第3報 RC10「ヘッダー右ズレ」):
'   modKnowledge.DrawChrome の右肩ピルは modChrome.FlowRight で帯の右端
'   (L+W-8)へ密着させていたが、ツールバー自体は modChrome.FlowLeft で
'   左詰めに流し込むだけなので、ボタン数が少ない端末(発行キー未設定・
'   画像解析無効等でボタンが減る)ではツールバーの実際の右端が帯の右端まで
'   届かない。ピルだけが右端に密着し、ツールバー本体はそれより手前で
'   終わっているように見える非対称が実機の「ヘッダーがズレて見える」の
'   正体。ここでツールバーの「実際に使っている右端」を先に計算し、
'   DrawChrome側がピルの右アンカーをそこへ合わせる。
' 戻り値: 全段のうち最も右まで到達したボタンの右端X(pt)。ボタンが0個の
'   ときは0を返す(R14-G13: 呼び出し側は「L+TB_PAD以下=縮退」だけを
'   フォールバック条件にする。少ボタン構成でも整列の恩恵を受けられるよう、
'   「右端が手前にある」こと自体は異常値として扱わない)。
Public Function ToolbarContentRight(ByVal isTable As Boolean, ByVal isShared As Boolean, _
                                    ByVal L As Double, ByVal W As Double) As Double
    Dim caps() As String, acts() As String, kinds() As String, tips() As String
    Dim n As Long
    Dim xs() As Double, rws() As Long, useW() As Double
    Dim rowN As Long
    ComputeToolbarLayout isTable, isShared, L, W, caps, acts, kinds, tips, n, xs, rws, useW, rowN
    If n < 1 Then Exit Function

    Dim rightMost As Double
    Dim i As Long
    For i = 0 To n - 1
        Dim edge As Double: edge = xs(i) + useW(i)
        If edge > rightMost Then rightMost = edge
    Next i
    ToolbarContentRight = rightMost
End Function

' DrawToolbar/ToolbarContentRight共有の配置算数(ToolbarSpec+FlowLeftの1本化)。
' 2つが将来ズレて「描画とテストで別の右端を答える」事故を起こさないための
' 唯一の計算経路。
Private Sub ComputeToolbarLayout(ByVal isTable As Boolean, ByVal isShared As Boolean, _
                                 ByVal L As Double, ByVal W As Double, _
                                 ByRef caps() As String, ByRef acts() As String, _
                                 ByRef kinds() As String, ByRef tips() As String, _
                                 ByRef n As Long, _
                                 ByRef xs() As Double, ByRef rws() As Long, _
                                 ByRef useW() As Double, ByRef rowN As Long)
    Dim widths() As Double
    ToolbarSpec isTable, isShared, caps, acts, kinds, tips, widths, n
    If n < 1 Then
        rowN = 0
        Exit Sub
    End If
    rowN = modChrome.FlowLeft(widths, n, L + TB_PAD, L + W - TB_PAD, TB_GAP, _
                              xs, rws, useW)
    If rowN < 1 Then rowN = 1
End Sub

' ----------------------------------------------------------------------------
' ToolbarSpec - モードと端末の権限に応じたボタンの並びを1本の配列で組み立てる。
'   ここが「何を出すか」の唯一の場所。配置(どこへ置くか)は一切決めない。
' ----------------------------------------------------------------------------
Private Sub ToolbarSpec(ByVal isTable As Boolean, ByVal isShared As Boolean, _
                        ByRef caps() As String, ByRef acts() As String, _
                        ByRef kinds() As String, ByRef tips() As String, _
                        ByRef widths() As Double, ByRef n As Long)
    ReDim caps(0 To TB_MAX - 1)
    ReDim acts(0 To TB_MAX - 1)
    ReDim kinds(0 To TB_MAX - 1)
    ReDim tips(0 To TB_MAX - 1)
    ReDim widths(0 To TB_MAX - 1)
    n = 0

    ' みんなの解決事例モードは操作が全く違う(選択と取り込み)。資料管理用の
    ' ボタンを並べても押しどころが分からなくなるので、専用の並びにする。
    If isShared Then
        AddTool caps, acts, kinds, tips, widths, n, _
                ChrW(&H2713) & " 選択を取り込む", "modShared.OnImportSelected", "accent", 112, _
                "チェックを付けた解決事例を自分の本棚に取り込みます"
        ' R18-3c: 「チャットへ」はここから上段のピル列(modKnowledge.PillSpec)へ
        ' 移した。他画面(Hub/ダッシュボード)はチャットへの導線が常に上段に
        ' あるのに、ナレッジ画面だけがボタン数で位置の変わる下段の帯にあり、
        ' 「どこにあるか毎回探す」状態だった(実機第5報②)。
        AddTool caps, acts, kinds, tips, widths, n, "すべて選ぶ", "modShared.OnSelectAll", "plain", 72, _
                "一覧の解決事例をすべて選びます"
        AddTool caps, acts, kinds, tips, widths, n, "選択を解除", "modShared.OnSelectNone", "plain", 72, _
                "選択をすべて外します"
        AddTool caps, acts, kinds, tips, widths, n, ChrW(&H2190) & " 前", "modShared.OnPrevPage", "plain", 44, _
                "前のページに戻ります"
        AddTool caps, acts, kinds, tips, widths, n, "次 " & ChrW(&H2192), "modShared.OnNextPage", "plain", 44, _
                "次のページに進みます"
        AddTool caps, acts, kinds, tips, widths, n, _
                ChrW(&HD83D) & ChrW(&HDCA1) & " みんなの困りごと", "modKnowledge.OnGapBoard", "plain", 104, _
                "回答が見つからなかった質問の一覧を見られます"
        Exit Sub
    End If

    ' 検索はギャラリー専用、削除は一覧表専用(押しても何も起きないボタンを
    ' 見せない=実機報告「どっちで押せばいいか分からない」への対処)。
    If Not isTable Then
        AddTool caps, acts, kinds, tips, widths, n, _
                ChrW(&HD83D) & ChrW(&HDD0D) & " 検索", "modKnowledge.OnSearch", "plain", 62, _
                "資料をキーワードで検索します"
    End If
    AddTool caps, acts, kinds, tips, widths, n, _
            ChrW(&H2795) & " 登録", "modKnowledge.OnRegister", "plain", 58, _
            "文章を直接ここに書いてナレッジとして登録します"
    AddTool caps, acts, kinds, tips, widths, n, _
            ChrW(&HD83D) & ChrW(&HDCC1) & " 追加", "modKnowledge.OnAddFiles", "plain", 58, _
            "PDFやWordなどのファイルを資料として追加します"
    ' R20-3(実機第7報②): 再取込ゼロで旧形式の資料に俯瞰・条文参照・言い換え
    ' 検索を後付けする(modBackfill.OnBackfillClick=このモジュールの薄い
    ' ハンドラ。実体はmodBackfillへ)。
    AddTool caps, acts, kinds, tips, widths, n, _
            ChrW(&H26A1) & " 仕上げ", "modKnowledgeBar.OnBackfillClick", "plain", 64, _
            "昔の形式の資料に章の目次と要約を後付けして、俯瞰質問や条文参照を使えるようにします"
    ' R18-3c: 「チャットへ」は上段のピル列へ移設(理由は上のisShared分岐の
    ' コメント参照)。ここはボタン数で位置が動く帯なので、常設の移動導線に
    ' 向かない。
    AddTool caps, acts, kinds, tips, widths, n, _
            ChrW(&HD83D) & ChrW(&HDCE6) & " パック出力", "modKnowledge.OnPackOut", "plain", 80, _
            "この本棚の中身をファイルにまとめて書き出します"
    AddTool caps, acts, kinds, tips, widths, n, _
            ChrW(&HD83D) & ChrW(&HDCE5) & " パック取込", "modKnowledge.OnPackIn", "plain", 80, _
            "書き出しておいた本棚ファイルを読み込みます"
    AddTool caps, acts, kinds, tips, widths, n, _
            ChrW(&HD83D) & ChrW(&HDD04) & " 同期", "modKnowledge.OnSync", "plain", 54, _
            "共有フォルダの最新版と資料を同期します"
    AddTool caps, acts, kinds, tips, widths, n, _
            ChrW(&HD83D) & ChrW(&HDCC2) & " フォルダ", "modKnowledge.OnPickFolder", "plain", 68, _
            "資料を自動で取り込む共有フォルダを選びます"
    AddTool caps, acts, kinds, tips, widths, n, _
            ChrW(&HD83D) & ChrW(&HDCA1) & " みんなの困りごと", "modKnowledge.OnGapBoard", "plain", 104, _
            "回答が見つからなかった質問の一覧を見られます"
    AddTool caps, acts, kinds, tips, widths, n, _
            ChrW(&HD83D) & ChrW(&HDCE1) & " 部門チャンネル", "modKnowledge.OnChannels", "plain", 96, _
            "部門ごとの共有先を設定します"

    ' 発行ボタンは、発行キーが設定されている端末にだけ出す。
    ' 一般利用者の画面に「押してはいけないボタン」を置かない。
    Dim canPub As Boolean
    On Error Resume Next
    canPub = modPublish.CanPublish()
    On Error GoTo 0
    If canPub Then
        AddTool caps, acts, kinds, tips, widths, n, _
                ChrW(&HD83D) & ChrW(&HDCE4) & " 正典を発行", "modPublishUI.OnPublish", "primary", 104, _
                "この本棚を部門の正式な資料として発行します"
        ' 運営向けの利用状況。発行者=運営なので同じ条件で出す。
        AddTool caps, acts, kinds, tips, widths, n, _
                ChrW(&HD83D) & ChrW(&HDCCA) & " 利用状況", "modHub.OnOwnerReport", "plain", 88, _
                "誰がどれだけ使っているかを確認します"
    End If

    ' 画像解析が使える環境でだけスクショ取込を出す(無効環境で「押したら
    ' 断られるボタン」を見せない。既存modUIShelfの方針をそのまま踏襲)。
    Dim hasVision As Boolean
    On Error Resume Next
    hasVision = modFeatures.FeatureEnabled("vision")
    On Error GoTo 0
    If hasVision Then
        AddTool caps, acts, kinds, tips, widths, n, _
                ChrW(&HD83D) & ChrW(&HDCF8) & " スクショ取込", "modUIShelf.OnIngestScreenshot", "plain", 84, _
                "画面のスクリーンショットを資料として取り込みます"
    End If

    ' R29 W2-7(実機第14報・削除ボタンの視認性): 取込系(登録/追加/仕上げ/
    ' パック出力入/同期/フォルダ等)と混じっていた🗑削除を配列の最後尾へ
    ' 動かし、並びの右端(FlowLeftは左→右の流し込みなので最後に足した項目が
    ' 最も右/最終段に来る)へ分離する。kind="danger"で赤系配色にし、
    ' 「消す操作だけ色が違う」ことを見た目でも切り離す。一覧表専用は不変
    ' (isTableのときだけ足す)。
    If isTable Then
        AddTool caps, acts, kinds, tips, widths, n, _
                ChrW(&HD83D) & ChrW(&HDDD1) & " 削除", "modKnowledge.OnDelete", "danger", 58, _
                "選んだ資料を本棚から削除します"
    End If
End Sub

' ----------------------------------------------------------------------------
' OnBackfillClick - 「⚡仕上げ」ボタン(R20-3)。判定・確認ダイアログ・
'   実処理は modBackfill(ingest層)側に閉じており、ここは取込中ガード+
'   結果の表示(トースト)+一覧の再描画だけを持つ薄いハンドラ。
'   2026-08-06 R20H FA-3: modUiLock.Enter/Leaveが欠けており、連打での
'   二重実行や他ハンドラとの入れ子実行を遮断できていなかった
'   (modHelpの全ハンドラと同型: BlockIfIngestingをEnterの前に置く。
'   後ろに置くとロックを取ったままExitし、全ボタンが10分死ぬ)。
' ----------------------------------------------------------------------------
Public Sub OnBackfillClick()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done

    Dim raw As String
    raw = modBackfill.BackfillAll()

    Dim p As Long: p = InStr(raw, "|")
    Dim kind As String, msg As String
    If p > 0 Then
        kind = Left$(raw, p - 1)
        msg = Mid$(raw, p + 1)
    Else
        kind = raw
    End If
    If kind = modBackfill.OUTCOME_CANCELLED Then GoTo Done

    On Error Resume Next
    If LenB(msg) > 0 Then
        modSkin.ShowToast msg, IIf(kind = modBackfill.OUTCOME_NONE, "info", "success")
    End If
    modUIShelf.RenderShelf
    On Error GoTo 0

Done:
    modUiLock.Leave
End Sub

' 起動後、本棚ツールバーの初回描画時に1回だけ「未仕上げ資料あり」を告知する
' (R20-3・3b)。isSharedの並び(みんなの解決事例)には⚡仕上げボタンが出ない
' ためスキップし、通常の並びに来るまで「初回」を消費しない。
Private Sub MaybeShowBackfillToast(ByVal isShared As Boolean)
    If mBackfillToastDone Then Exit Sub
    If isShared Then Exit Sub
    mBackfillToastDone = True

    On Error Resume Next
    Dim cands As Collection: Set cands = modBackfill.DetectLegacyDocs()
    Dim needN As Long: needN = modBackfill.CountByStatus(cands, modBackfill.STATUS_NEEDS)
    If needN > 0 Then
        modSkin.ShowToast "旧形式の資料が" & needN & "冊あります。ツールバーの" & _
            ChrW(&H26A1) & "仕上げで俯瞰・条文参照が有効になります(再取込不要)。", "info"
    End If
    On Error GoTo 0
End Sub

' 並びへ1個足す(TB_MAXを超えたら黙って捨てる=配列外参照で全滅させない)。
' tipText(R30 W2-5): ホバー説明。Optional省略時は""(ツールチップ無し)。
Private Sub AddTool(ByRef caps() As String, ByRef acts() As String, _
                    ByRef kinds() As String, ByRef tips() As String, _
                    ByRef widths() As Double, ByRef n As Long, _
                    ByVal capText As String, ByVal actName As String, _
                    ByVal kind As String, ByVal itemW As Double, _
                    Optional ByVal tipText As String = "")
    If n >= TB_MAX Then Exit Sub
    caps(n) = capText
    acts(n) = actName
    kinds(n) = kind
    tips(n) = tipText
    widths(n) = itemW
    n = n + 1
End Sub

' ツールバーのボタン1個。kind: "plain"(白地) / "primary"(青地) / "accent"(強調) /
' "danger"(赤地。R29 W2-7=削除専用。全テーマ共通の固定色 #DC2626×白文字。
' コントラスト比4.83:1(WCAG AA 4.5:1以上)を検算済みで、テーマ配色に紐付く
' modUI.UiColorへは委ねず本モジュール内に直書きする=どのスキンでも
' 「消す操作だけは常に赤」という一貫した視覚合図にする)。
' tipText(R30 W2-5): ホバー説明。実機依存(Excelのみ確認可)。クリック挙動は
' OnActionがハイパーリンクに優先して動く既知手法(Address:=""・SubAddressは
' 自セルなので、万一ハイパーリンクが先に反応しても遷移しない)。
Private Sub ToolButton(ByVal ws As Worksheet, ByVal shapeName As String, _
                       ByVal capText As String, ByVal action As String, _
                       ByVal kind As String, ByVal x As Double, ByVal y As Double, _
                       ByVal w As Double, Optional ByVal tipText As String = "")
    ' Constは関数を持てないため通常のDimで組む(RGB()は関数呼び出し)。
    Dim dangerBg As Long: dangerBg = RGB(220, 38, 38)   ' #DC2626
    ' 1個の1004で残りを道連れにしない。
    On Error Resume Next
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(5, x, y, w, BAR_H)
    If Err.Number = 0 And Not btn Is Nothing Then
        btn.Name = shapeName
        ' 行3の高さは描いたあとに入れる(段数が決まるのが描画後のため)。
        ' 絶対配置にしておかないと行高の変更でボタンが伸縮する。
        btn.Placement = 3
        btn.Adjustments(1) = 0.35
        If kind = "primary" Or kind = "accent" Or kind = "danger" Then
            btn.Line.Visible = 0
            If kind = "primary" Then
                btn.Fill.ForeColor.RGB = modUI.UiColor("primary")
            ElseIf kind = "danger" Then
                btn.Fill.ForeColor.RGB = dangerBg
            Else
                btn.Fill.ForeColor.RGB = modUI.UiColor("accent")
            End If
        Else
            btn.Line.Visible = -1
            btn.Line.Weight = 0.75
            btn.Line.ForeColor.RGB = modUI.UiColor("border")
            btn.Fill.ForeColor.RGB = modUI.UiColor("surface")
        End If
        With btn.TextFrame2
            .WordWrap = -1
            .TextRange.Text = capText
            .TextRange.Font.Size = 8.5
            If kind = "primary" Or kind = "accent" Or kind = "danger" Then
                .TextRange.Font.Bold = -1
                .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            Else
                .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
            End If
            .TextRange.ParagraphFormat.Alignment = 2
            .VerticalAnchor = 3
            .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
        End With
        modSkin.ApplyLightShadow btn
        btn.OnAction = action
        ' R30 F3: danger(削除)ボタンにはツールチップを付けない。OnDeleteSourceが
        ' ActiveCell.Row依存で、Hyperlinks.Addの自セル選択("SubAddress"にホバーで
        ' 反応する端末がある)が起きると削除が恒久不能になるリスクがあるため。
        If LenB(tipText) > 0 And kind <> "danger" Then
            ws.Hyperlinks.Add Anchor:=btn, Address:="", _
                SubAddress:=btn.TopLeftCell.Address, ScreenTip:=tipText
        End If
    End If
    Set btn = Nothing
    Err.Clear
    On Error GoTo 0
End Sub
