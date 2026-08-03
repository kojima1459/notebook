Attribute VB_Name = "modKnowledge"
Option Explicit

' modKnowledge - ナレッジ画面の共通クロム(ヘッダー+モード切替+ツールバー)。
' 2026-07-26再設計: Vault(カードギャラリー)/マイ本棚(一覧表)の2シートに
' 同じヘッダー・ツールバーを描き、上部ピルで切り替える1画面2モードに見せる。
' ツールバーは既存のmodVault/modUIShelf/modPack/modShelfSyncへ配線するだけ。
' 設計の鉄則: Shape座標は実セル幾何から導く(旧実装の決め打ち座標が画面外に
' はみ出していた教訓)。配色はmodUI.UiColor()。Shape名はnxk_接頭辞。

Private Const HDR_H As Double = 40
Private Const PILL_H As Double = 26
' 右肩ピル(R7 A-5): モード3つ + 🎨着せ替え + 🚪終了 の5個。
' 2026-07-31(R11-F2): 「?」ヘルプを足して6個(監査1 L-3。ナレッジ画面から
' ヘルプへ行く手段が無く、困った人が詰む状態だった。R11-Eでは容量不足で保留)。
Private Const PILL_N As Long = 6
Private Const PILL_PITCH As Double = 9
Private Const PILL_PAD As Double = 14
Private Const PILL_MIN As Double = 44

' 旧ギャラリー/解決事例の描画先だった実行時生成シート。2026-07-30(R4要件A)で
' 描画先を「マイ本棚」へ統合したため、既存ブックに残っているものを消すためだけに
' 名前を持っている(新規に作ることは二度と無い)。
Private Const LEGACY_VAULT_SHEET As String = "Vault"

' 上部クロムが占める行(1..6)。本文はDrawChrome後の ContentTop から下に描く。
' 行3がツールバーの帯で、段数に応じて高さが伸びる(行数は変えない)。
Public Const CHROME_ROWS As Long = 6

' 直近に描いたモード("gallery"/"table"/"shared")。
' 3モードとも同じ「マイ本棚」シートに描くようになった(2026-07-30 R4要件A)ため、
' 「今どのモードか」はシート名からは分からない。ここが唯一の情報源。
Private mMode As String

' DrawChrome - ヘッダー+モードピル+ツールバーを描く(冪等)。
'   mode: "gallery"(ナレッジ倉庫のカード) / "table"(マイ本棚の一覧)
Public Sub DrawChrome(ByVal ws As Worksheet, ByVal mode As String)
    If ws Is Nothing Then Exit Sub
    On Error GoTo Fail

    RemoveChrome ws
    mMode = LCase$(mode)

    ' 幾何を先に確定させる(順序が逆だとShape座標がズレる)。
    ' 行1(ヘッダー)と行2(隙間)だけ先に決めれば 行3.Top は確定する
    ' (行3自身の高さはツールバーが何段になったかが分かってから入れる)。
    ws.Rows(1).RowHeight = HDR_H
    ws.Rows(2).RowHeight = 6

    ' 罫線と行列番号を隠す。ここを消さないと、どれだけ整えても
    ' 画面が「Excelのシート」にしか見えない(実機要望: エクセル感を消す)。
    On Error Resume Next
    If ThisWorkbook.ActiveSheet Is ws Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
    End If
    On Error GoTo Fail

    Dim L As Double, cellW As Double, W As Double
    L = ws.Range("A1").Left
    cellW = ws.Range("A1:N1").Width
    ' R11-B(#30本丸): セル範囲幅(cellW)基準のクロムは本棚系で772〜882ptに
    ' なり、実可視域(約600pt)を大幅超過して右肩ピルが画面外へ出ていた。
    ' 帯の背景(下のhdr)は従来どおりcellWいっぱいのまま、操作系を置く
    ' 右端(W)だけを実可視幅(modUIMain.ViewportWidth)へクランプする。
    W = modChrome.BarWidth(cellW, modUIMain.ViewportWidth(), 8)

    ' --- 右肩ピルの配置を先に決める(帯の高さがこれで決まるため) ---
    ' 2026-07-31(R7 A-5): 🎨着せ替え / 🚪終了 をここへ統合し、
    ' 「どの画面でも右上は同じ」にする(並びはチャット画面と同じで一番右が🚪)。
    ' 位置は固定座標をやめて modChrome.FlowRight に任せる。1段に入らなければ
    ' 段を増やし、帯の高さもそれに合わせて伸ばす=画面外や見切れが起きない。
    Dim pcaps(0 To PILL_N - 1) As String, pnames(0 To PILL_N - 1) As String
    Dim pacts(0 To PILL_N - 1) As String, pwid(0 To PILL_N - 1) As Double
    Dim pactive(0 To PILL_N - 1) As Boolean
    Dim md As String: md = LCase$(mode)
    Dim isTable As Boolean: isTable = (md = "table")
    Dim isShared As Boolean: isShared = (md = "shared")
    PillSpec md, isTable, isShared, pcaps, pnames, pacts, pactive, pwid

    ' R14-2a(実機第3報 RC10): ピルは帯の右端(L+W-8)へ密着させていたが、
    ' ツールバーはボタン数が減ると帯の右端まで届かず、両者の右端が
    ' 食い違って見えていた。ツールバーの実際の右端へピルの右アンカーを
    ' 合わせる(異常値=帯の右端付近まで届いていないときはL+W-8へ戻す。
    ' 200pt未満は「計算が壊れた」とみなせる下限)。isTable/isSharedは
    ' 下のDrawToolbar呼び出しと同じ値(このSub内で唯一の判定)。
    Dim tbRight As Double
    On Error Resume Next
    tbRight = modKnowledgeBar.ToolbarContentRight(isTable, isShared, L, W)
    On Error GoTo Fail
    If tbRight < L + 200 Then tbRight = L + W - 8

    Dim pxs() As Double, prows() As Long, pws() As Double
    Dim pillRowN As Long
    pillRowN = modChrome.FlowRight(pwid, PILL_N, tbRight, _
                                   L + modChrome.TitleReserve(W, 160), L + 8, 6, _
                                   pxs, prows, pws)
    If pillRowN < 1 Then pillRowN = 1
    Dim hdrH As Double: hdrH = HDR_H + (pillRowN - 1) * (PILL_H + 2)
    If hdrH > 200 Then hdrH = 200          ' 行高の異常値でDrawChrome全体を落とさない
    ws.Rows(1).RowHeight = hdrH

    ' --- ヘッダーバー ---
    Dim hdr As Shape
    Set hdr = ws.Shapes.AddShape(5, L, 0, cellW, hdrH)
    hdr.Name = "nxk_hdr"
    hdr.Adjustments(1) = 0.02
    hdr.Line.Visible = 0
    hdr.Fill.ForeColor.RGB = modUI.UiColor("sidebar")
    modSkin.ApplyHeaderDepth hdr          ' §9: 濃紺の2色グラデーション
    With hdr.TextFrame2
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCDA) & " ナレッジ"
        .TextRange.Font.Size = 12
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .MarginLeft = 82
        .VerticalAnchor = 3
    End With

    ' 実機報告(2026-07-27)「←Hubがヘッダーと同色で目立たない」対策。
    ' 戻り導線は一番見つけやすくなければならないので、白ピルで強調する。
    Pill ws, "nxk_back", ChrW(&H2190) & " Hub", L + 8, 72, 0, _
         "modKnowledge.OnBackHub", True

    ' --- モード切替ピル+共通ヘッダー機能(右肩) ---
    Dim pi As Long
    For pi = 0 To PILL_N - 1
        Pill ws, pnames(pi), pcaps(pi), pxs(pi), pws(pi), _
             prows(pi) * (PILL_H + 2), pacts(pi), pactive(pi)
    Next pi

    On Error Resume Next
    modTelemetry.TrackScreen "knowledge"
    On Error GoTo Fail

    ' --- ツールバー(行3の帯) ---
    ' 2026-07-30(R4要件C): 折り返しが「1段目だけ」だったため、2段目に
    ' あふれたボタンは幅を見ずに右へ描き続けられ、画面外で見切れていた
    ' (実機写真の「🗑削除が『除』しか見えない」)。段数無制限の流し込みに
    ' 統一し、実際に使った高さをそのまま帯(行3)の高さにする。
    ' こうすると検索欄(行5)もカード領域(行7以降)も自動で下がり、
    ' 何段になっても重ならない。
    Dim barH As Double
    barH = modKnowledgeBar.DrawToolbar(ws, isTable, isShared, L, W, ws.Rows(3).Top)
    If barH < modKnowledgeBar.BAR_H + 2 Then barH = modKnowledgeBar.BAR_H + 2
    ' 3-C: Excelの行高上限は409.5pt。段数無制限にしたので、極端に狭い帯
    ' (異常な列幅設定など)では計算上それを超えうる。超えた値を代入すると
    ' 1004になり、そこから先(検索欄・カード)が丸ごと描かれない。
    ' 手前でクランプして「崩れても描き切る」ほうを選ぶ。
    If barH > 400 Then barH = 400
    ws.Rows(3).RowHeight = barH
    ws.Rows(4).RowHeight = 2
    ws.Rows(5).RowHeight = 22
    ws.Rows(6).RowHeight = 8

    On Error Resume Next
    modUI.FreezeShapePlacement ws
    On Error GoTo 0
    Exit Sub

Fail:
    modLog.LogError "E0801", "modKnowledge.DrawChrome", Err.Description, Err.Number
End Sub

' ----------------------------------------------------------------------------
' PrepareScreenView - 「マイ本棚」「Hub」を表示・再描画するときの共通儀式。
' ----------------------------------------------------------------------------
' 2026-07-30(R4要件B): 実機で「🗑削除が『除』しか見えない」「ボタンが右で
' 切れている」と報告された画面は、多くが「一度右へスクロールした状態が
' 持ち越されているだけ」だった。Nexus起動時に水平スクロールバーを消して
' あるため、利用者には戻す手段が無い。表示のたびに必ず左上へ戻す。
' あわせて、旧"Vault"シートが残っているブックの移行掃除もここで行う
' (表示経路すべてがここを通るので、掃除の呼び忘れが起きない)。
' 罫線・行列番号の表示制御は各画面の既存処理に任せる(ここでは触らない)。
Public Sub PrepareScreenView(ByVal ws As Worksheet)
    PurgeLegacyVaultSheet
    ' 2026-07-31(R7 A-2): 全画面/数式バー/罫線の崩れをここでも自己修復する。
    ' 「閉じる→キャンセル」で壊れた表示が、次にどのボタンを押しても戻る。
    On Error Resume Next
    modUI.EnsureAppView
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ' ActiveWindow系は、そのシートが実際に前面のときだけ触る
    ' (別シートの表示状態を巻き添えで変えないため)。
    If ThisWorkbook.ActiveSheet Is ws Then
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
        ActiveWindow.Zoom = 100
    End If
    On Error GoTo 0
End Sub

' 旧"Vault"シートの移行削除。ギャラリーと解決事例の描画先を「マイ本棚」へ
' 統合した(R4要件A)ので、既存ブックに残る空の"Vault"タブは
' 「同じ資料が2つの画面にある」という元の混乱をそのまま残してしまう。
Private Sub PurgeLegacyVaultSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LEGACY_VAULT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ' 2026-07-30(レビュー2-D/7-B): DisplayAlerts は変更前の値を控えて戻す
    ' (True 決め打ちで戻すと、警告を切って一括処理している最中に呼ばれた
    ' ときに、その処理の途中から警告ダイアログが出るようになる)。
    ' また Delete に失敗した場合、可視化したままにすると旧"Vault"タブが
    ' 画面に出てしまう(消せなかったうえに混乱だけが増える)ので、
    ' まだ残っていたら veryHidden へ戻す。
    Dim prevAlerts As Boolean: prevAlerts = True
    On Error Resume Next
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ws.Visible = -1          ' veryHiddenのままだとDeleteが1004になる環境がある
    ws.Delete
    Application.DisplayAlerts = prevAlerts
    On Error GoTo 0

    ' 消えたかどうかを取り直して確認する(消えていれば参照は失敗する)。
    Dim still As Worksheet
    On Error Resume Next
    Set still = ThisWorkbook.Worksheets(LEGACY_VAULT_SHEET)
    On Error GoTo 0
    If still Is Nothing Then Exit Sub
    On Error Resume Next
    still.Visible = 2        ' xlSheetVeryHidden(消せなかったので隠し直す)
    modLog.LogError "E0801", "modKnowledge(PurgeLegacyVaultSheet)", _
        "旧Vaultシートを削除できなかったため veryHidden へ戻しました"
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' CurrentMode - 現在モードの正規化(このモジュールで唯一の解釈)。
' ----------------------------------------------------------------------------
' 2026-07-30(レビュー1-A): mMode が "" (起動直後・未描画)のときの解釈が
' IsTableMode(=一覧表とみなす)と RefreshCurrent(Case Else でギャラリー)で
' 食い違っていた。取込直後の自動再描画が「一覧表のつもりで書いた上に
' ギャラリーを描く」といった噛み合わない動きの原因になる。
' 既定の画面は一覧表なので、"" は "table" に寄せる。両方ここを使う。
Private Function CurrentMode() As String
    Dim m As String: m = LCase$(Trim$(mMode))
    If m = "gallery" Or m = "shared" Then
        CurrentMode = m
    Else
        CurrentMode = "table"
    End If
End Function

' 今「マイ本棚」シートに描かれているのが一覧表モードかどうか。
' 3モードが同じシートを共有する(R4要件A)ため、取込・同期の完了時に
' 走る自動再描画(modShelf/modShelfSync → modUIShelf.RenderShelf)が、
' ギャラリー表示中に一覧表をカードの裏へ書き込んでしまう経路ができた。
' RenderShelf側でここを見て空振りさせる。
' 未描画(起動直後)は一覧表とみなす ―― 既定の画面が一覧表だから。
Public Function IsTableMode() As Boolean
    IsTableMode = (CurrentMode() = "table")
End Function

' 本文(カード/表)を描き始めてよいY座標。
Public Function ContentTop(ByVal ws As Worksheet) As Double
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    ContentTop = ws.Rows(CHROME_ROWS + 1).Top
    On Error GoTo 0
End Function

' 検索キーワードのセル(ギャラリーだけが使う)。
Public Function SearchCellAddress() As String
    SearchCellAddress = "B5:E5"
End Function

' ----------------------------------------------------------------------------
' PillSpec - 右肩に並べるものの唯一の定義(2026-07-31 R7 A-5)。
'   右から: 🚪終了 / 🎨着せ替え / みんなの解決事例 / マイ本棚 / ギャラリー。
'   幅はキャプションの実文字から出す(固定幅の予約と実物が食い違うと、
'   実機で「🗑削除が『除』しか見えない」類の見切れになる。R4要件Cの教訓)。
' ----------------------------------------------------------------------------
Private Sub PillSpec(ByVal md As String, ByVal isTable As Boolean, ByVal isShared As Boolean, _
                     ByRef caps() As String, ByRef nms() As String, ByRef acts() As String, _
                     ByRef actives() As Boolean, ByRef widths() As Double)
    caps(0) = ChrW(&HD83D) & ChrW(&HDEAA) & " 終了"
    nms(0) = "nxk_exit": acts(0) = "modApp.OnSaveAndExit": actives(0) = False
    caps(1) = ChrW(&HD83C) & ChrW(&HDFA8) & " 着せ替え"
    nms(1) = "nxk_skin": acts(1) = "modHub.OnThemeToggle": actives(1) = False
    caps(2) = ChrW(&HD83C) & ChrW(&HDF81) & " みんなの解決事例"
    nms(2) = "nxk_m_shared": acts(2) = "modKnowledge.OnGoShared": actives(2) = isShared
    caps(3) = ChrW(&HD83D) & ChrW(&HDCCB) & " マイ本棚"
    nms(3) = "nxk_m_table": acts(3) = "modKnowledge.OnGoTable": actives(3) = isTable
    caps(4) = ChrW(&HD83C) & ChrW(&HDCCF) & " ギャラリー"
    nms(4) = "nxk_m_gallery": acts(4) = "modKnowledge.OnGoGallery"
    actives(4) = (Not isTable) And (Not isShared)
    ' 「?」= 使い方。ヘルプカードはチャット画面(Nexus)の上に描く作りなので、
    ' modHub.OnHelp と同じくチャットへ移ってから開く(2026-07-31 R11-F2)。
    caps(5) = ChrW(&H2753)
    nms(5) = "nxk_help": acts(5) = "modKnowledge.OnHelp": actives(5) = False

    Dim i As Long
    For i = 0 To PILL_N - 1
        widths(i) = modChrome.PillWidth(caps(i), PILL_PITCH, PILL_PAD, PILL_MIN)
    Next i
End Sub

' ヘッダー上のピル。active:=Trueで「今いるモード」を塗りつぶして示す。
' yOff は段送り(1段に入りきらなかったぶんを下の段へ置くための縦オフセット)。
Private Sub Pill(ByVal ws As Worksheet, ByVal shapeName As String, _
                 ByVal caption As String, ByVal x As Double, ByVal w As Double, _
                 ByVal yOff As Double, ByVal action As String, ByVal active As Boolean)
    On Error Resume Next
    Dim p As Shape
    Set p = ws.Shapes.AddShape(5, x, (HDR_H - PILL_H) / 2 + yOff, w, PILL_H)
    If p Is Nothing Then Exit Sub
    p.Name = shapeName
    p.Placement = 3          ' 行高を後から変えてもピルは動かさない
    p.Adjustments(1) = 0.35
    p.Line.Visible = 0
    If active Then
        p.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Else
        p.Fill.ForeColor.RGB = modUI.UiColor("sidebarActive")
    End If
    With p.TextFrame2
        .TextRange.Text = caption
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = -1
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
    End With
    If active Then
        p.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("sidebar")
    Else
        p.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End If
    p.OnAction = action
    On Error GoTo 0
End Sub

Private Sub RemoveChrome(ByVal ws As Worksheet)
    Dim names() As String
    ReDim names(0 To ws.Shapes.Count)
    Dim n As Long
    Dim shp As Shape
    For Each shp In ws.Shapes
        ' 旧マイ本棚のボタン(btn_/lbl_)も一緒に消す。残すと新ツールバーの上に
        ' 浮いたまま二重表示になる(Hub移植時に踏んだのと同じ罠)。
        ' 2026-07-30(R4要件A): 3モードが同じ「マイ本棚」シートを共有する
        ' ようになったので、前のモードのカード(nxg_)とチェックボックス(nxs_)も
        ' ここで必ず落とす。残すとギャラリーのカードの上に一覧表が重なる。
        If Left$(shp.Name, 4) = "nxk_" Or Left$(shp.Name, 4) = "btn_" _
           Or Left$(shp.Name, 4) = "lbl_" Or Left$(shp.Name, 4) = "nxg_" _
           Or Left$(shp.Name, 4) = "nxs_" Then
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

' ---- ツールバーのハンドラ(既存エンジンへの配線に徹する) ----
'
' 2026-07-31(R7 B-2): 画面遷移・資料操作の入口は必ず
' modUiLock.BlockIfIngesting を最初に通す。取込/同期の最中は
' 抽出ループのDoEventsでここが【取込の途中から入れ子で】走り出し、
' 実機の「取込中にボタンを押すとExcelが応答なし」になっていた。
' busy のときはモーダルを出さず、実況行だけ出して即Exitする。

Public Sub OnGoGallery()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    ' R14-2b: 既にギャラリー表示中の再押下は無反応に見えていた(押しても
    ' 見た目が変わらない=憲章§3-1違反)。押下前のモードを CurrentMode()
    ' (右肩ピルのactives()判定と同じ唯一の情報源)で控えておき、既に
    ' ギャラリーだった場合だけ再描画後にトーストで「効いた」ことを伝える。
    Dim wasGallery As Boolean: wasGallery = (CurrentMode() = "gallery")
    On Error Resume Next
    modVaultGallery.ShowVaultGallery
    If wasGallery Then modSkin.ShowToast "表示を更新しました", "info", True
    On Error GoTo 0
    modUiLock.Leave
End Sub

' みんなのQ&A(選択式取り込み)へ。
Public Sub OnGoShared()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    ' 2026-07-31(R7実装後の追加是正・発見事項4): Leave→Showの順で
    ' 順序が入れ替わっていたうえ、On Error Resume Nextの直後にGoTo 0が
    ' 無く、その後のExit Subまで丸ごと保護区間に入っていた。modShared.Show
    ' が失敗しても完全に無音になり、他の兄弟ハンドラ(OnGoGallery等)と
    ' 挙動が食い違っていた。
    On Error Resume Next
    modShared.Show
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnGoTable()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUIShelf.EnsureLayout
    modUI.GoToNativeSheet modAppDef.SH_SHELF, "modKnowledge.OnGoTable"
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnBackHub()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modHub.EnsureHubLayout activate:=True
    On Error GoTo 0
    modUiLock.Leave
End Sub

' 「?」= 使い方。ヘルプカードはNexus(チャット)シート上に描かれるため、
' 先にチャットへ移ってから modHelp.OnHelpClick を呼ぶ(modHub.OnHelpと同型)。
' OnHelpClick 自身が modUiLock を取るので、ここではロックを取らない。
Public Sub OnHelp()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    On Error Resume Next
    modUI.GoToNexus "modKnowledge.OnHelp"
    If Err.Number <> 0 Then modLog.LogError "E0801", "modKnowledge.OnHelp", Err.Description, Err.Number
    On Error GoTo 0
    modHelp.OnHelpClick
End Sub

Public Sub OnToChat()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUI.GoToNexus "modKnowledge.OnToChat"
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnSearch()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modVaultGallery.OnVaultSearch
    If Err.Number <> 0 Then modLog.LogError "E0801", "modKnowledge.OnSearch", Err.Description, Err.Number
    On Error GoTo 0
    modUiLock.Leave
End Sub

' ----------------------------------------------------------------------------
' 💡 みんなの困りごと: 組織で答えが見つからなかった質問の一覧。
'   資料を書ける人(商品部)がここを見て、その場でナレッジを書けるようにする。
'   営業の「分からない」が、商品部の「書くべきこと」に直結する一番短い経路。
' ----------------------------------------------------------------------------
Public Sub OnGapBoard()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done

    Dim body As String
    On Error Resume Next
    body = modInsight.GapListText()
    On Error GoTo Done

    Dim n As Long
    On Error Resume Next
    n = modInsight.GapCount()
    On Error GoTo Done

    Dim resp As VbMsgBoxResult
    resp = MsgBox( _
        "みんなが質問して、本棚に答えが無かった質問です(新しい順・最大20件)。" & vbCrLf & _
        "ここに並ぶ質問に答える資料を用意すると、部内の全員がすぐ答えを得られます。" & vbCrLf & vbCrLf & _
        body & vbCrLf & _
        "この内容に答える資料を、今すぐ登録しますか?", _
        vbYesNo + vbInformation, modAppDef.APP_NAME & " - みんなの困りごと (" & n & "件)")

    If resp = vbYes Then
        modUiLock.Leave
        modVault.ShowVaultInput
        Exit Sub
    End If
Done:
    modUiLock.Leave
End Sub


' ----------------------------------------------------------------------------
' 部門チャンネル: 購読・更新・チャンク予算をひとまとめに扱う入口。
'   全社共通/商品/システム/人事… と部門ごとに正典が発行される。ここで
'   必要なものだけ購読する。全部入れないのが既定なので、部門が増えても
'   ひとりのブックが際限なく膨らむことはない。
' ----------------------------------------------------------------------------
Public Sub OnChannels()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done

    Dim all As String
    On Error Resume Next
    all = modChannel.ListChannels()
    On Error GoTo Done

    If LenB(all) = 0 Then
        modUiLock.Leave
        MsgBox "部門の公式ナレッジがまだ1つも見つかりません。" & vbCrLf & vbCrLf & _
            "各部門が正典を発行すると、自動でここに現れます。" & vbCrLf & _
            "(共有フォルダの channels\<部門名>\ に pack.xlsx と version.txt が置かれる形です)" & vbCrLf & vbCrLf & _
            "共有フォルダ自体が未設定の場合は、Hubのお知らせから設定してください。", _
            vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    ' 2026-07-27: 「どの部門か」を利用者に選ばせるのをやめた。
    '
    ' 以前はここで、アプリが既に知っている部門名を一覧表示したうえで、その
    ' どれかを利用者にInputBoxへ打ち直させていた。さらに「常時つないでおける
    ' のは1部門だけ」「切り替えると前の部門は本棚から外れます」と説明していた。
    ' これは modChannel 冒頭のコメントが「誤りだった」と名指ししている設計
    ' そのもの ―― 聞く前に分野を自分で判断させた時点で、ポータルを探し回る
    ' のと同じ認知負荷が生まれる。実務は「今日は商品、明日はシステム」であり、
    ' 分野の判定は人間ではなく検索側の仕事。
    Dim parts() As String: parts = Split(all, "|")
    Dim n As Long: n = UBound(parts) - LBound(parts) + 1

    Dim listText As String
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        listText = listText & "  ・" & parts(i) & vbCrLf
    Next i

    modUiLock.Leave
    If MsgBox("見つかった " & n & " 部門の公式ナレッジを、まとめて本棚に読み込みます。" & vbCrLf & vbCrLf & _
              listText & vbCrLf & _
              "以後はどの分野の質問でも、部門を選ばずにそのまま聞けます。" & vbCrLf & _
              "(あなたが自分で入れた資料はそのまま残ります)" & vbCrLf & vbCrLf & _
              "分量によっては数分かかることがあります。よろしいですか?", _
              vbOKCancel + vbQuestion, modAppDef.APP_NAME & " - 部門の公式ナレッジ") <> vbOK Then Exit Sub

    ' 2026-07-28(レビュー M-21): 実処理の区間はロックを取り直す。
    ' 確認ダイアログの前に Leave しているため、ここは無防備だった。
    ' 取り込みは共有I/Oのリトライ待ちで DoEvents を回すので、その間に
    ' もう一度押されると同じ処理が入れ子で走り、purge と import が
    ' 交錯して本棚が中途半端な状態になる。
    If Not modUiLock.Enter() Then Exit Sub

    Dim result As String
    On Error Resume Next
    ' 取り込み中の ESC を実行時エラー18として捕捉できるようにする
    ' (捕捉しないと Excel が処理を強制中断し、purge 直後で止まり得る)。
    Application.EnableCancelKey = 2      ' xlErrorHandler
    modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDCE1) & " 部門の公式ナレッジを読み込んでいます…"
    result = modChannel.SubscribeAllAvailable()
    modUIMain.SetStage ""
    Application.EnableCancelKey = 1      ' xlInterrupt(既定へ戻す)
    On Error GoTo 0
    modUiLock.Leave

    On Error Resume Next
    modHub.EnsureHubLayout
    On Error GoTo 0

    MsgBox result & vbCrLf & vbCrLf & _
           "本棚の使用量: " & modChannel.ChunkUsagePercent() & "%", _
           vbInformation, modAppDef.APP_NAME
    Exit Sub
Done:
    modUiLock.Leave
End Sub


Public Sub OnRegister()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modVault.ShowVaultInput
    If Err.Number <> 0 Then modLog.LogError "E0801", "modKnowledge.OnRegister", Err.Description, Err.Number
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnAddFiles()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modShelfBatch.AddFilesViaDialog
    If Err.Number <> 0 Then modLog.LogError "E0801", "modKnowledge.OnAddFiles", Err.Description, Err.Number
    On Error GoTo 0
    RefreshCurrent
    modUiLock.Leave
End Sub

' OnPackOut/OnPackIn: 例外ハンドラ必須(監査1 M-9・VBA生ダイアログ防止)。
Public Sub OnPackOut()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Fail
    modPackExport.ExportPackDialog
    modUiLock.Leave
    Exit Sub
Fail:
    modUiLock.Leave
    modLog.ShowError "E0801", "modKnowledge.OnPackOut", Err.Description
End Sub

Public Sub OnPackIn()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Fail
    modPack.ImportPackDialog
    RefreshCurrent
    modUiLock.Leave
    Exit Sub
Fail:
    modUiLock.Leave
    modLog.ShowError "E0801", "modKnowledge.OnPackIn", Err.Description
End Sub

Public Sub OnSync()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modShelfSync.SyncNow
    If Err.Number <> 0 Then modLog.LogError "E0801", "modKnowledge.OnSync", Err.Description, Err.Number
    On Error GoTo 0
    RefreshCurrent
    modUiLock.Leave
End Sub

Public Sub OnPickFolder()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modShelfSync.PickShelfFolder
    If Err.Number <> 0 Then modLog.LogError "E0801", "modKnowledge.OnPickFolder", Err.Description, Err.Number
    On Error GoTo 0
    RefreshCurrent
    modUiLock.Leave
End Sub

Public Sub OnDelete()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUIShelf.OnDeleteSource
    If Err.Number <> 0 Then modLog.LogError "E0801", "modKnowledge.OnDelete", Err.Description, Err.Number
    On Error GoTo 0
    RefreshCurrent
    modUiLock.Leave
End Sub

' 今表示しているモードだけを描き直す(モード切替をまたいで表示がズレないよう、
' 資料を足した/消した直後は必ずここを通す)。
'
' 2026-07-30(R4要件A/F): 3モードとも描画先が「マイ本棚」シート1枚になったので、
' 旧実装のように ActiveSheet.Name では現在モードを判別できない。DrawChromeが
' 記録した mMode を唯一の情報源にする。modApp.OnRefreshUI(🔄再描画)の
' フォールバック先でもあるため Public。
Public Sub RefreshCurrent()
    On Error Resume Next
    ' モードの解釈は CurrentMode() 1箇所に集約する(レビュー1-A)。
    ' ここで mMode を直接見ると、"" の扱いが IsTableMode とずれる。
    Select Case CurrentMode()
        Case "shared"
            modShared.Show
        Case "gallery"
            modVaultGallery.ShowVaultGallery
        Case Else
            modUIShelf.RenderShelf
    End Select
    If Err.Number <> 0 Then modLog.LogError "E0801", "modKnowledge.RefreshCurrent", Err.Description, Err.Number
    On Error GoTo 0
End Sub

