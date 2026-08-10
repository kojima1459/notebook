Attribute VB_Name = "modVaultGallery"
Option Explicit

' ============================================================================
' modVaultGallery - ナレッジ倉庫ギャラリー(カード一覧・検索・ページング)。
'   2026-07-31(R11-F1): modVault が30,000字上限まで残り34字となり、保留中の
'   修正(検索/ページ送りの例外処理・ページ端トースト・0件時のページャ抑止)が
'   1件も入らない状態だったため、憲章§4-6に基づきギャラリー系を分離した。
'   modVault 側は「ナレッジ登録フォーム」(入力→RegisterKnowledgeText)だけを
'   持つ。両者は共有する内部状態を持たない(mGallery*/mPreview* はすべて
'   ギャラリー側でしか使われていなかった)。
' ============================================================================

' ナレッジ画面のカードに出す先頭チャンクの索引(レビュー M-16)。
' source名(小文字) -> プレビュー文字列。mPreviewRows は作り直し判定用の行数。
Private mPreview As Object
Private mPreviewRows As Long

' ナレッジ倉庫ギャラリー用の宣言(実機VBAは宣言部をモジュール先頭に集約する必要あり)
' 2026-07-30(R4要件A): 描画先を実行時生成の "Vault" シートから
' 「マイ本棚」(modAppDef.SH_SHELF)へ移した。タブに出る画面が
' 「マイ本棚」と「ナレッジ倉庫」の2枚に割れていて、中身が同じデータなのに
' 別物に見えていたのが実機の混乱の元だったため、シートを1枚に統合する。
Private Const CARDS_PER_PAGE As Long = 9
Private Const CARD_H As Double = 120
' R20-1c(実機第7報⑦の層1): カードは3列固定だった。3列=215*3+14*2=673pt で、
' 帯だけが可視幅へ伸びる窓(1300pt/1800pt)では右に600pt以上の空白が残る。
' 列数を帯幅から決める。
' R21-S3(実機第8報⑦): それでもカード幅が215pt固定だったため、列数が変わる
' 境目までの端数がそのまま右の余白として残っていた(1024pt窓では4列=884pt
' に対し中身の右端916pt・帯1004ptで、右に88ptの白)。列数【と】カード幅の
' 両方を弾性にし、最後のカードの右端を帯の右端(ContentRight)へ一致させる。
'   cols  = clamp(int((avail-leftX+GAP)/(CARD_W_MIN+GAP)), 3, 6)
'   cardW = clamp((avail-leftX-GAP*(cols-1))/cols, CARD_W_MIN, CARD_W_MAX)
' 上限6列は「1ページ9枚が2段で収まる」ところ。下限3列は従来の見た目。
Private Const CARD_W_MIN As Double = 170
Private Const CARD_W_MAX As Double = 240
Private Const CARD_GAP As Double = 14
Private Const CARD_COLS_MIN As Long = 3
Private Const CARD_COLS_MAX As Long = 6
' ページャの高さ(カードの下に置く帯)。窓高適応圧縮の必要量に入れる。
Private Const PAGER_H As Double = 28

Private mGalleryPage As Long
Private mGalleryNames() As String   ' 現在ページのカード順の資料名(クリック解決用)
Private mGalleryCount As Long

' 直近の描画で確定した最大ページ番号(0起点)。ページ端で「最後のページです」と
' 知らせるために持つ(2026-07-31 R11-F2。R11-Bで容量不足のため保留になっていた
' 監査1 H-2/M-6の回収)。
Private mGalleryMaxPage As Long

' ----------------------------------------------------------------------------
' ナレッジ倉庫ギャラリー(設計: 単一Shape=1カード・可変列グリッド(3〜6列。
' R20-1c で帯幅から決める)・ページング。
' Shape増殖なし=毎回同数のカードを描き直す)
' 宣言部(CARDS_PER_PAGE/CARD_W_MIN/CARD_H/mGallery*)はモジュール先頭に集約済み。
' 描画先は「マイ本棚」シート(R4要件A)。専用シートはもう作らない。
' ----------------------------------------------------------------------------

' ギャラリーを表示(SPA遷移)。検索語はシートのD3セル(検索バー)から読む。
' uiStep: modUIShelf.EnsureLayoutと同型(2026-07-31 R10-1)。Activate失敗の
' 無言スキップ(ギャラリー無反応の原因)を解消し、E0801を必ず1行残す。
' 戻り値(2026-08-03 R14-G8): 描き直しに成功したかどうか。呼び出し元
'   (modKnowledge.OnGoGallery)が「表示を更新しました」を出してよいのは成功時
'   だけで、失敗時のトースト(下のE0801経路)を上書きしてはならない。
'   Subのままだと呼び出し元からは成功も失敗も見分けが付かない。
Public Function ShowVaultGallery() As Boolean
    Dim ws As Worksheet
    Set ws = GetOrCreateGallerySheet()
    If ws Is Nothing Then
        ' R14-2b(実機第3報 RC10): 従来は無言Exitで、押しても何も起きない
        ' ように見えていた(憲章§3-1違反)。RedrawGalleryの失敗時の
        ' 案内と同型にする(E0801+利用者向けトースト)。
        modLog.LogError "E0801", "modVaultGallery.ShowVaultGallery", _
            "「マイ本棚」シートを取得/作成できませんでした"
        On Error Resume Next
        modSkin.ShowToast "一覧を描き直せませんでした。もう一度お試しください。", "error"
        On Error GoTo 0
        Exit Function
    End If

    Dim uiStep As String
    Dim gErrNum As Long, gErrDesc As String
    Dim vw0 As Double, vh0 As Double
    On Error GoTo Finish
    Application.ScreenUpdating = False
    modUiLock.AlertsOff   ' R25-1a-1(対はFinishCleanup4)

    ' R21-S1(実機第8報⑦): 「描く→活性化→表示状態」を【活性化→表示状態→
    ' 描く】へ反転する。従来は罫線・見出し・タブ・水平スクロールバーの確定が
    ' 描画の【後】にあり、カード列数と境界を「タブと横バーが出たままの窓」で
    ' 決めていた(可視高が約24pt過大・可視幅も別物)。測ってよいのは
    ' modViewport2.EnsureViewState を通った後だけ。
    ws.Visible = -1
    uiStep = "シートのアクティブ化(ActivateSheetRobust)"
    If modUI.ActivateSheetRobust(ws, "modVaultGallery.ShowVaultGallery") Then
        On Error Resume Next
        modViewport2.EnsureViewState ws
        On Error GoTo Finish
        ' 表示の共通儀式(左端へ戻す/等倍/旧Vaultシートの掃除)。R4要件B。
        uiStep = "表示の共通儀式(PrepareScreenView)"
        modKnowledge.PrepareScreenView ws
    End If
    ' Falseのときは両者をスキップ(ActiveWindowが別シートを向いたまま触らない)。

    modViewport2.MarkView vw0, vh0
    uiStep = "枠の描画(DrawGalleryFrame)"
    DrawGalleryFrame ws
    uiStep = "カードの描画(RenderGalleryCards)"
    RenderGalleryCards ws
    modUI.FreezeShapePlacement ws   ' 全Shapeを絶対配置に固定(ズレ防止)
    modSkin.BeautifyAll ws          ' フォント統一(Yu Gothic UI)+固定クロムに柔らかい影

    ' 正常系はハンドラ本体(Resume)を跨いで後始末へ入る
    ' (Resume はエラーが起きていないと実行時エラー20になる)。
    GoTo FinishCleanup4
Finish:
    gErrNum = Err.Number
    gErrDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FinishCleanup4
FinishCleanup4:
    On Error Resume Next
    modUiLock.AlertsOn
    If gErrNum <> 0 Then
        modLog.LogError "E0801", "modVaultGallery.ShowVaultGallery", "[" & uiStep & "] " & gErrDesc, gErrNum
    End If
    Application.ScreenUpdating = True   ' 例外時も必ず画面更新を戻す(暗転固定を防ぐ)
    ShowVaultGallery = (gErrNum = 0)
    On Error GoTo 0
    ' R21-S1の保険: 描画中に窓が動いていたら1回だけ組み直す(ワンショット)。
    modViewport2.ReflowIfMoved vw0, vh0
End Function

' ----------------------------------------------------------------------------
' 検索・ページ送りの3ハンドラ(2026-07-31 R11-F2)。
'   従来は例外保護が無く、描画の途中で1004等が出ると ScreenUpdating=False の
'   まま呼び出し元へ飛び、画面が固まったように見えた(暗転固定)。しかも
'   err_log に何も残らないため、実機で「押しても何も起きない」としか報告
'   できなかった(監査1 H-2。R11-Bで容量不足のため保留になっていた項目)。
'   3本とも RedrawGallery に集約し、そこで E0801 と画面更新の復帰を必ず行う。
' ----------------------------------------------------------------------------
Public Sub OnVaultSearch()
    mGalleryPage = 0
    RedrawGallery "modVaultGallery.OnVaultSearch"
End Sub

Public Sub OnVaultPrev()
    ' R15波1b裁定: 取込・同期中の再描画は入れ子で崩れ得る(既存40箇所超と
    ' 同型。lint検査13で現状追認だったものを保護へ切り替え)。
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If mGalleryPage <= 0 Then
        ' 端で無反応にしない。「押したのに何も起きない」は故障と同じ(憲章§3-1)。
        ' waitless: 端で押しただけの案内に1.1秒待たせない(R11-H Med4)。
        modSkin.ShowToast "最初のページです。", "info", True
        Exit Sub
    End If
    mGalleryPage = mGalleryPage - 1
    RedrawGallery "modVaultGallery.OnVaultPrev"
End Sub

Public Sub OnVaultNext()
    ' R15波1b裁定: OnVaultPrevと同型の保護(上記コメント参照)。
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If mGalleryPage >= mGalleryMaxPage Then
        modSkin.ShowToast "最後のページです。", "info", True
        Exit Sub
    End If
    mGalleryPage = mGalleryPage + 1
    RedrawGallery "modVaultGallery.OnVaultNext"
End Sub

' ページ位置を変えずに描き直す(カードからの削除・品質報告の後に使う)。
' context には呼び出し元の公開ハンドラ名を渡す(err_logから経路が追えるように)。
Private Sub OnVaultSearchKeepPage()
    RedrawGallery "modVaultGallery.OnVaultCardClick"
End Sub

' ギャラリーの再描画(唯一の入口)。例外が出ても必ず ScreenUpdating を戻し、
' 何が起きたかを E0801 に1行残す。
Private Sub RedrawGallery(ByVal context As String)
    Dim ws As Worksheet
    Dim gErrNum As Long, gErrDesc As String

    Set ws = GetGallerySheet()
    If ws Is Nothing Then
        modLog.LogError "E0801", context, "「マイ本棚」シートが見つかりませんでした"
        Exit Sub
    End If

    On Error GoTo Fail
    Application.ScreenUpdating = False
    RenderGalleryCards ws
    GoTo RedrawCleanup
Fail:
    gErrNum = Err.Number
    gErrDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かない。後始末の前に Resume で
    ' ハンドラを抜ける(2026-07-30 実機err#462と同型)。
    Resume RedrawCleanup
RedrawCleanup:
    On Error Resume Next
    Application.ScreenUpdating = True   ' 例外時も必ず戻す(暗転固定を防ぐ)
    If gErrNum <> 0 Then
        modLog.LogError "E0801", context, gErrDesc, gErrNum
        modSkin.ShowToast "一覧を描き直せませんでした。もう一度お試しください。", "error"
    End If
    On Error GoTo 0
End Sub


Public Sub OnVaultBackToChat()
    ' R15波1b裁定: 画面遷移は取込中の入れ子実行と衝突し得る(OnVaultPrevと
    ' 同型。lint検査13で現状追認だったものを保護へ切り替え)。
    If modUiLock.BlockIfIngesting() Then Exit Sub
    modUI.GoToNexus "modVaultGallery.OnVaultBackToChat"
End Sub

' カードクリック: 内容の先頭を表示し、削除も選べる
Public Sub OnVaultCardClick()
    ' R15波1b裁定: MsgBox+削除/報告(いずれもシート状態を操作)は取込中の
    ' 入れ子実行と衝突し得る(OnVaultPrevと同型)。
    If modUiLock.BlockIfIngesting() Then Exit Sub
    Dim callerName As String
    On Error Resume Next
    callerName = CStr(Application.Caller)
    On Error GoTo 0
    If Left$(callerName, 9) <> "nxg_card_" Then
        ' 想定外のShapeから呼ばれた(配線ミス/Shape名の変更)。無言で消えると
        ' 「押しても何も起きない」として二度と使われないため必ず1行残す。
        modLog.LogUsage "caller_mismatch", "modVaultGallery.OnVaultCardClick", callerName
        Exit Sub
    End If

    Dim idx As Long
    idx = CLng(Val(Mid$(callerName, 10)))
    If idx < 0 Or idx >= mGalleryCount Then Exit Sub

    Dim srcName As String
    srcName = mGalleryNames(idx)

    Dim answer As Long
    answer = MsgBox("『" & srcName & "』" & vbLf & vbLf & _
                    modUtil.SafeLeft(PreviewOf(srcName), 300) & vbLf & vbLf & _
                    "[はい]=削除  /  [いいえ]=" & ChrW(&H26A0) & "ノイズ報告(品質が低いと報告)  /  [キャンセル]=閉じる", _
                    vbYesNoCancel + vbQuestion + vbDefaultButton2, modAppDef.APP_NAME & " - ナレッジ詳細")
    If answer = vbYes Then
        modShelf.DeleteSource srcName
        OnVaultSearchKeepPage
    ElseIf answer = vbNo Then
        modStats.ReportNoise srcName          ' 個人ミュート(即時・自分の検索からのみ除外)
        On Error Resume Next
        modP2P.EmitNoiseVote srcName          ' 組織的除外への1票(共有フォルダ・同期時に集計)
        On Error GoTo 0
        MsgBox "『" & srcName & "』を品質報告しました。" & vbLf & vbLf & _
               "・あなたの検索からは今すぐ除外されます。" & vbLf & _
               "・異なる" & modStats.NoiseThreshold() & "人以上が報告すると、組織全体の検索から除外されます。", _
               vbInformation, modAppDef.APP_NAME
        OnVaultSearchKeepPage   ' 画面を再描画(既存のプライベートSubを呼ぶ)
    End If
End Sub

' ---- ギャラリー内部描画 ----

' ギャラリーの枠。ヘッダー/モード切替/ツールバーは modKnowledge が両モード
' 共通で描く(2026-07-26 再設計)。ここは検索欄とカード領域の下地だけを持つ。
Private Sub DrawGalleryFrame(ByVal ws As Worksheet)
    RemoveShapesByPrefix ws, "nxg_bar_"   ' 旧ツールバー(x決め打ち)の掃除
    ' 「マイ本棚」シートを3モードで共有する(R4要件A)ので、前のモードが
    ' 書いたセル(一覧表の結合・行高・値)を必ず消してから描く。消さないと
    ' カードの裏に一覧表が透けて残る。Shapeの掃除はDrawChromeが行う。
    ws.Cells.Clear
    ' A:N を全列ぶん明示する(一覧表モードがK/L未設定だったために、
    ' DrawChromeが使う W=A1:N1 の幅が機種・履歴依存でぶれていた)。
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B:N").ColumnWidth = 12
    ' R20-1d: 毎回406行(7:412)を書き直すのをやめ、前回使った行までに絞る
    ' (行高を明示した行はExcelから見れば「使用済み」=下スクロール域になる)。
    modKnowledge.NormalizeShelfRows ws, 7
    ' R18-3a: 全域(ws.Cells)への書式はUsedRangeをシート最大へ膨らませる
    ' (無限スクロールの主因・調査agent2 §1.3)。実使用範囲だけに当てる。
    ' R19-1b: ここではまだカードが無いので1画面ぶん。カードを描き終えた
    ' ApplyGalleryExtent が実下端まで伸ばす(旧 A1:N412 は8画面ぶんの空塗り)。
    ' 範囲は行高を戻した【後】に決める(前モードの行高のまま数えると足りない)。
    Dim bandAddr As String: bandAddr = modKnowledge.ShelfBound(ws, 0)
    ws.Range(bandAddr).Interior.Color = modUI.UiColor("bg")
    ws.Range(bandAddr).Font.Name = "Yu Gothic UI"

    modKnowledge.DrawChrome ws, "gallery"

    ' 検索バー(セル)。位置は modKnowledge のクロム行(1..6)のうち行5。
    With ws.Range(modKnowledge.SearchCellAddress())
        .Merge
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = 1
        .Borders.Color = modUI.UiColor("border")
        .IndentLevel = 1
    End With
    With ws.Range("F5")
        .Value = ChrW(&H2190) & " ここにキーワードを入れて「検索」を押す(例: 約款, 保険金)"
        .Font.Size = 9
        .Font.Color = modUI.UiColor("muted")
    End With
End Sub

' 検索→フィルタ→現在ページのカードだけを描く(カードShapeは毎回作り直すが
' 最大CARDS_PER_PAGE枚で一定=増殖しない)
Private Sub RenderGalleryCards(ByVal ws As Worksheet)
    RemoveShapesByPrefix ws, "nxg_card"
    RemoveShapesByPrefix ws, "nxg_pg_"
    RemoveShapesByPrefix ws, "nxg_empty"

    ' カード領域の起点。共通クロム(modKnowledgeのヘッダー+ツールバー+検索欄)の
    ' 直下から実測で求める(旧コードのy=84決め打ちだとクロムと重なる)。
    Dim cardL As Double: cardL = ws.Range("B1").Left
    Dim cardT As Double: cardT = modKnowledge.ContentTop(ws) + 6
    If cardT < 90 Then cardT = 90

    ' R20-1c/R21-S3: 帯の実右端(可視幅に合わせ済み)から、列数【と】カード幅を
    ' 出す。帯・ピル・ツールバーと同じ modViewport.ContentRight が単一情報源で、
    ' 最後のカードの右端 = cardL + cols*cardW + (cols-1)*GAP = bandRight になる。
    Dim bandRight As Double: bandRight = modViewport.ContentRight(ws, modKnowledge.SHELF_BAND, 8)
    Dim availW As Double: availW = bandRight - ws.Range("A1").Left
    Dim leftX As Double: leftX = cardL - ws.Range("A1").Left
    Dim cols As Long
    cols = modViewport2.GridColsFor(availW, leftX, CARD_W_MIN, CARD_GAP, _
                                    CARD_COLS_MIN, CARD_COLS_MAX)
    Dim cardW As Double
    cardW = modViewport2.GridCardW(availW, leftX, CARD_GAP, cols, CARD_W_MAX, CARD_W_MIN)
    ' カード幅が上限240で頭打ちになってなお余る幅は隙間へ配分する(ダッシュの
    ' modDashStat.CardGapFor と同型)。これが無いと広い窓でだけ右端が届かない。
    Dim cardGap As Double
    cardGap = modViewport2.GridGapFor(availW, leftX, cardW, cols, CARD_GAP)
    ' R21-S5: 1ページぶんの段数が窓高に収まらないときだけカード高を詰める。
    ' R24-1a: F3準拠で固定/可変に分ける(旧2引数のままの取り残し=実Excelでは
    ' 「引数は省略できません」でコンパイル不能)。段ピッチは SY(CARD_H)+CARD_GAP
    ' (:339 cardH=SY(CARD_H)、CARD_GAPは非適用)なので、可変= rowsNeed*CARD_H、
    ' 固定= cardT + rowsNeed*CARD_GAP + PAGER_H。modHub:139-141 と同型。
    Dim rowsNeed As Long: rowsNeed = (CARDS_PER_PAGE + cols - 1) \ cols
    modViewport2.SetScaleY modViewport2.CompressFactor(modViewport.ViewportHeight(), _
                cardT + rowsNeed * CARD_GAP + PAGER_H, _
                rowsNeed * CARD_H)
    Dim cardH As Double: cardH = modViewport2.SY(CARD_H)
    ' Empty State の透かし/文面もカード群と同じ幅(左端から帯の右端まで)に伸ばす。
    Dim emptyW As Double: emptyW = bandRight - cardL
    If emptyW < 320 Then emptyW = 320

    Dim keyword As String
    keyword = LCase$(Trim$(CStr(ws.Range("B5").Value)))

    Dim names() As String, stats() As String
    Dim total As Long
    total = modShelf.SourceList(names, stats)

    ' フィルタ(名前 or プレビューに部分一致)
    Dim fNames() As String, fStats() As String
    Dim fCount As Long: fCount = 0
    If total > 0 Then
        ReDim fNames(0 To total - 1)
        ReDim fStats(0 To total - 1)
        Dim i As Long
        For i = 0 To total - 1
            Dim hay As String
            hay = LCase$(names(i) & " " & PreviewOf(names(i)))
            If LenB(keyword) = 0 Or InStr(hay, keyword) > 0 Then
                fNames(fCount) = names(i)
                fStats(fCount) = stats(i)
                fCount = fCount + 1
            End If
        Next i
    End If

    ' ページ境界
    Dim maxPage As Long
    If fCount = 0 Then
        maxPage = 0
    Else
        maxPage = (fCount - 1) \ CARDS_PER_PAGE
    End If
    If mGalleryPage > maxPage Then mGalleryPage = maxPage
    If mGalleryPage < 0 Then mGalleryPage = 0
    mGalleryMaxPage = maxPage   ' ページ端トーストの判定材料(R11-F2)

    Dim startIdx As Long: startIdx = mGalleryPage * CARDS_PER_PAGE
    Dim endIdx As Long: endIdx = startIdx + CARDS_PER_PAGE - 1
    If endIdx > fCount - 1 Then endIdx = fCount - 1

    mGalleryCount = 0
    ReDim mGalleryNames(0 To CARDS_PER_PAGE - 1)

    If fCount = 0 Then
        ' Empty State(空の状態): 空白で放置せず、透かしアイコン+誘導CTAを配置する。
        On Error Resume Next
        ws.Range("B7:G8").UnMerge
        ws.Range("B7:G8").ClearContents
        On Error GoTo 0

        Dim icon As Shape
        Set icon = ws.Shapes.AddShape(1, cardL, cardT + 40, emptyW, 60)
        icon.Name = "nxg_empty_icon"
        icon.Fill.Visible = 0: icon.Line.Visible = 0
        With icon.TextFrame2
            .WordWrap = -1
            .TextRange.Text = ChrW(&HD83D) & ChrW(&HDD0D)   ' 虫めがね
            .TextRange.Font.Size = 40
            .TextRange.ParagraphFormat.Alignment = 2
        End With
        icon.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(148, 163, 184)

        Dim emsg As Shape
        Set emsg = ws.Shapes.AddShape(1, cardL, cardT + 104, emptyW, 46)
        emsg.Name = "nxg_empty_msg"
        emsg.Fill.Visible = 0: emsg.Line.Visible = 0
        With emsg.TextFrame2
            .WordWrap = -1
            If LenB(keyword) > 0 Then
                .TextRange.Text = "「" & ws.Range("B5").Value & "」に一致するナレッジが見つかりません。" & vbLf & _
                                  "AIにこの質問を投げて、新しいナレッジを作りませんか?"
            Else
                .TextRange.Text = "まだナレッジがありません。" & vbLf & _
                                  "「" & ChrW(&H2795) & " 登録」「" & ChrW(&HD83D) & ChrW(&HDCC1) & " 追加」で資料を取り込むか、AIに質問してみましょう。"
            End If
            .TextRange.Font.Size = 11
            .TextRange.ParagraphFormat.Alignment = 2
        End With
        emsg.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(107, 114, 128)

        Dim cta As Shape
        Set cta = ws.Shapes.AddShape(5, cardL + (emptyW - 140) / 2, cardT + 158, 140, 34)
        cta.Name = "nxg_empty_cta"
        cta.Adjustments(1) = 0.3
        cta.Line.Visible = 0
        cta.Fill.ForeColor.RGB = modUI.UiColor("primary")
        With cta.TextFrame2
            .WordWrap = -1
            .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCAC) & " AIに質問する"
            .TextRange.Font.Size = 10.5
            .TextRange.Font.Bold = -1
            .TextRange.ParagraphFormat.Alignment = 2
            .VerticalAnchor = 3
        End With
        cta.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        cta.OnAction = "modVaultGallery.OnVaultBackToChat"
    Else
        On Error Resume Next
        ws.Range("B7:G8").UnMerge
        ws.Range("B7:G8").ClearContents
        On Error GoTo 0

        Dim k As Long
        For k = startIdx To endIdx
            Dim slot As Long: slot = k - startIdx
            Dim col As Long: col = slot Mod cols
            Dim rowN As Long: rowN = slot \ cols
            DrawOneCard ws, slot, cardL + col * (cardW + cardGap), cardT + rowN * (cardH + CARD_GAP), _
                        cardW, cardH, fNames(k), fStats(k)
            mGalleryNames(slot) = fNames(k)
            mGalleryCount = mGalleryCount + 1
        Next k
    End If

    ' ページャ。0件のときは描かない(2026-07-31 R11-F2。監査1 M-6の回収)。
    ' 「1 / 1 ページ(全0件)」と前へ/次へだけが浮いている画面は、押しても
    ' 何も起きないボタンを2つ見せることになり、Empty Stateの案内を打ち消す。
    If fCount = 0 Then
        ApplyGalleryExtent ws
        Exit Sub
    End If

    ' R20-1c/1d: ページャは「実際に使った段数」の下へ。列数が増えれば
    ' 9枚が2段で収まるので、その分だけ画面が縦に詰まる(空スクロール域も減る)。
    Dim rowsUsed As Long: rowsUsed = ((endIdx - startIdx) \ cols) + 1
    If rowsUsed < 1 Then rowsUsed = 1
    Dim pgY As Double: pgY = cardT + rowsUsed * (cardH + CARD_GAP) + 6
    Dim prevBtn As Shape
    Set prevBtn = ws.Shapes.AddShape(5, cardL, pgY, 70, 22)
    prevBtn.Name = "nxg_pg_prev"
    ' 実機報告(2026-07-27)「ページ送りが真っ白で何か分からない」対策。
    ' 白地+薄グレー枠では背景に溶ける。塗りと文字色を明示する。
    prevBtn.Adjustments(1) = 0.3
    prevBtn.Fill.ForeColor.RGB = modUI.UiColor("primary")
    prevBtn.Line.Visible = 0
    prevBtn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    prevBtn.TextFrame2.TextRange.Font.Bold = -1
    prevBtn.TextFrame2.WordWrap = -1
    prevBtn.TextFrame2.TextRange.Text = ChrW(&H25C0) & " 前へ"
    prevBtn.TextFrame2.TextRange.Font.Size = 8.5
    prevBtn.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    prevBtn.TextFrame2.VerticalAnchor = 3
    prevBtn.TextFrame2.MarginLeft = 10: prevBtn.TextFrame2.MarginRight = 10
    prevBtn.TextFrame2.MarginTop = 6: prevBtn.TextFrame2.MarginBottom = 6
    prevBtn.OnAction = "modVaultGallery.OnVaultPrev"

    Dim pgInfo As Shape
    Set pgInfo = ws.Shapes.AddShape(1, cardL + 76, pgY, 146, 22)
    pgInfo.Name = "nxg_pg_info"
    pgInfo.Fill.Visible = 0
    pgInfo.Line.Visible = 0
    pgInfo.TextFrame2.WordWrap = -1
    pgInfo.TextFrame2.TextRange.Text = (mGalleryPage + 1) & " / " & (maxPage + 1) & " ページ(全" & fCount & "件)"
    pgInfo.TextFrame2.TextRange.Font.Size = 9
    pgInfo.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
    pgInfo.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    pgInfo.TextFrame2.VerticalAnchor = 3
    pgInfo.TextFrame2.MarginLeft = 10: pgInfo.TextFrame2.MarginRight = 10
    pgInfo.TextFrame2.MarginTop = 6: pgInfo.TextFrame2.MarginBottom = 6

    Dim nextBtn As Shape
    Set nextBtn = ws.Shapes.AddShape(5, cardL + 226, pgY, 70, 22)
    nextBtn.Name = "nxg_pg_next"
    nextBtn.Adjustments(1) = 0.3
    nextBtn.Fill.ForeColor.RGB = modUI.UiColor("primary")
    nextBtn.Line.Visible = 0
    nextBtn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    nextBtn.TextFrame2.TextRange.Font.Bold = -1
    nextBtn.TextFrame2.WordWrap = -1
    nextBtn.TextFrame2.TextRange.Text = "次へ " & ChrW(&H25B6)
    nextBtn.TextFrame2.TextRange.Font.Size = 8.5
    nextBtn.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    nextBtn.TextFrame2.VerticalAnchor = 3
    nextBtn.TextFrame2.MarginLeft = 10: nextBtn.TextFrame2.MarginRight = 10
    nextBtn.TextFrame2.MarginTop = 6: nextBtn.TextFrame2.MarginBottom = 6
    nextBtn.OnAction = "modVaultGallery.OnVaultNext"
    ApplyGalleryExtent ws
End Sub

' ApplyGalleryExtent - 描いたカード/ページャの実下端まで塗りと ScrollArea を
'   伸ばす(R19-1b)。カードはpt座標のShapeなので、下端はShapeから実測する。
'   窓が小さい端末ではカード3段が1画面に収まらないため、ここが無いと
'   下半分が「塗られていない白」になる(=白い断崖)。
Private Sub ApplyGalleryExtent(ByVal ws As Worksheet)
    On Error Resume Next
    Dim bottomY As Double
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 4) = "nxg_" Then
            If shp.Top + shp.Height > bottomY Then bottomY = shp.Top + shp.Height
        End If
    Next shp
    ' R20-1d: 塗り・ScrollArea・境界より下の行高リセットは modKnowledge へ集約。
    modKnowledge.ApplyShelfBound ws, bottomY
    On Error GoTo 0
End Sub

' 単一Shape=1カード(タイトル太字+プレビュー+日付を1テキストに結合し、
' 部分書式で表現。グループ化しない=軽量・増殖なし)
Private Sub DrawOneCard(ByVal ws As Worksheet, ByVal slot As Long, ByVal x As Double, _
                        ByVal y As Double, ByVal cardW As Double, ByVal cardH As Double, _
                        ByVal srcName As String, ByVal statLine As String)
    ' R15-8b(実機第4報 RC1): 単純Split(statLine,"|")の重複実装をやめ、
    ' error_note中の"|"にもズレないmodUIShelf.ParseStatsの頑健パースへ共用する
    ' (このカードで実際に使うのは addedAt/chunkN の2値のみ=挙動不変)。
    Dim pStatus As String, pIngestedAt As String, pChunkCount As String
    Dim pErrorNote As String, pOrigin As String
    modUIShelf.ParseStats statLine, pStatus, pIngestedAt, pChunkCount, pErrorNote, pOrigin
    Dim addedAt As String: addedAt = pIngestedAt
    Dim chunkN As String: chunkN = pChunkCount

    Dim titleText As String: titleText = modUtil.SafeLeft(srcName, 40)
    Dim previewText As String: previewText = modUtil.SafeLeft(PreviewOf(srcName), 90)
    Dim footText As String
    footText = ChrW(&HD83D) & ChrW(&HDCC5) & " " & ShortStamp(addedAt) & "  ・ " & chunkN & " chunks"

    Dim isExcluded As Boolean
    On Error Resume Next
    isExcluded = modStats.IsGloballyExcluded(srcName)
    On Error GoTo 0

    Dim card As Shape
    Set card = ws.Shapes.AddShape(5, x, y, cardW, cardH)
    card.Name = "nxg_card_" & slot
    card.Adjustments(1) = 0.08
    card.Line.Weight = 0.75
    card.Shadow.Visible = 0

    Dim body As String
    If isExcluded Then
        ' 組織的除外中: グレーアウト+警告行を追加(カードのサイズ/位置は変えない)
        card.Fill.ForeColor.RGB = modUI.UiColor("bg")
        card.Line.ForeColor.RGB = modUI.UiColor("border")
        body = titleText & vbLf & previewText & vbLf & footText & vbLf & _
               ChrW(&H26A0) & " 組織的除外(調査中)"
    Else
        card.Fill.ForeColor.RGB = RGB(255, 255, 255)
        card.Line.ForeColor.RGB = RGB(229, 231, 235)
        body = titleText & vbLf & previewText & vbLf & footText
    End If

    With card.TextFrame2
        .WordWrap = -1
        .MarginLeft = 10: .MarginRight = 10: .MarginTop = 8: .MarginBottom = 8
        .TextRange.Text = body
        .TextRange.Font.Size = 8.5
        If isExcluded Then
            .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
        Else
            .TextRange.Font.Fill.ForeColor.RGB = RGB(107, 114, 128)
        End If
        .VerticalAnchor = 1
        ' タイトル行のみ太字・大きめ(除外時は色もmutedに統一してグレーアウト表現)
        With .TextRange.Paragraphs(1).Font
            .Size = 10
            .Bold = -1
            If isExcluded Then
                .Fill.ForeColor.RGB = modUI.UiColor("muted")
            Else
                .Fill.ForeColor.RGB = RGB(17, 24, 39)
            End If
        End With
    End With
    card.OnAction = "modVaultGallery.OnVaultCardClick"
End Sub

' 資料の先頭チャンク本文(breadcrumb行を除去した150字)をプレビューとして返す。
Private Function PreviewOf(ByVal srcName As String) As String
    ' 2026-07-28(レビュー M-16): 資料ごとに my_knowledge を1セルずつ
    ' 走査していた。資料100件×1万チャンクなら最悪100万回のCOM往復で、
    ' ナレッジ画面が分単位でフリーズする(Static のキャッシュは直近1件しか
    ' 効かないので、カードを縦に並べる用途では毎回ミスする)。
    ' 「全資料ぶんの先頭チャンク」を一括読みで一度だけ索引化する。
    EnsurePreviewIndex
    If mPreview Is Nothing Then Exit Function
    If mPreview.Exists(LCase$(srcName)) Then PreviewOf = mPreview(LCase$(srcName))
End Function

' my_knowledge を一括で読み、資料ごとの先頭チャンクだけを拾って索引化する。
' 行数が変わったら作り直す(取込・削除のあとに古い内容を出さないため)。
Private Sub EnsurePreviewIndex()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then Exit Sub

    Dim lastK As Long
    lastK = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then
        Set mPreview = CreateObject("Scripting.Dictionary")
        mPreviewRows = 0
        Exit Sub
    End If

    If Not mPreview Is Nothing Then
        If mPreviewRows = lastK Then Exit Sub    ' 変化なし=作り直さない
    End If

    Set mPreview = CreateObject("Scripting.Dictionary")
    mPreviewRows = lastK

    ' source(2列目)と full_text(7列目)だけが要る。
    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, 2), ws.Cells(lastK, 7)).Value
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim nm As String: nm = LCase$(Trim$(CStr(arr(i, 1))))
        If LenB(nm) > 0 Then
            If Not mPreview.Exists(nm) Then
                mPreview(nm) = MakePreviewText(CStr(arr(i, 6)))
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

' full_text をカード用の1行プレビューへ整える(breadcrumb行【…】は剥がす)。
Private Function MakePreviewText(ByVal fullText As String) As String
    Dim txt As String: txt = fullText
    If Left$(txt, 1) = "【" Then
        Dim lfPos As Long: lfPos = InStr(txt, vbLf)
        If lfPos > 0 Then txt = Mid$(txt, lfPos + 1)
    End If
    MakePreviewText = modUtil.SafeLeft(Replace(txt, vbLf, " "), 150)
End Function

Private Function ShortStamp(ByVal stamp As String) As String
    If IsDate(stamp) Then
        Dim d As Date: d = CDate(stamp)
        ShortStamp = Year(d) & "/" & Month(d) & "/" & Day(d)
    Else
        ShortStamp = modUtil.SafeLeft(stamp, 10)
    End If
End Function

' ギャラリーの描画先は「マイ本棚」シート(2026-07-30 R4要件A)。
' このシートは起動時(modBoot→modUIShelf.EnsureLayout)に必ず作られるので、
' ここで作る必要はない。無い場合はmodUIShelf側に作らせる。
Private Function GetGallerySheet() As Worksheet
    On Error Resume Next
    Set GetGallerySheet = ThisWorkbook.Worksheets(modAppDef.SH_SHELF)
    On Error GoTo 0
End Function

Private Function GetOrCreateGallerySheet() As Worksheet
    Dim ws As Worksheet
    Set ws = GetGallerySheet()
    If ws Is Nothing Then
        On Error Resume Next
        modUIShelf.EnsureLayout        ' シートごと作る責務はmodUIShelfが持つ
        On Error GoTo 0
        Set ws = GetGallerySheet()
    End If
    Set GetOrCreateGallerySheet = ws
End Function

Private Sub RemoveShapesByPrefix(ByVal ws As Worksheet, ByVal prefix As String)
    Dim names() As String
    ReDim names(0 To ws.Shapes.count)
    Dim n As Long: n = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, Len(prefix)) = prefix Then
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
