Attribute VB_Name = "modVault"
Option Explicit

' ナレッジ画面のカードに出す先頭チャンクの索引(レビュー M-16)。
' source名(小文字) -> プレビュー文字列。mPreviewRows は作り直し判定用の行数。
Private mPreview As Object
Private mPreviewRows As Long

' ============================================================================
' modVault - ナレッジ登録フォーム(DOCS_NEXUS_SPEC Phase 2・裁定①)
' ----------------------------------------------------------------------------
' UserForm代替: 「登録フォーム用にデザインされた専用シート」をSPAの画面遷移
' としてアクティブ化する方式(Shape上のテキスト入力はフォーカス制御が不安定な
' ため、入力欄はセルで作る)。登録は既存の取込パイプライン(modShelf.IngestFile)
' を再利用: 入力内容を%TEMP%のUTF-8テキストに書き出して取り込むことで、
' 構造チャンク化・重複排除・バッチ埋め込みまで全て既存の実証済み経路に乗せる。
' ============================================================================

Private Const VAULT_SHEET As String = "VaultInput"
Private Const CELL_TITLE As String = "C6"
Private Const CELL_BODY As String = "C8"
Private Const CELL_TAGS As String = "C18"

' ナレッジ倉庫ギャラリー用の宣言(実機VBAは宣言部をモジュール先頭に集約する必要あり)
' 2026-07-30(R4要件A): 描画先を実行時生成の "Vault" シートから
' 「マイ本棚」(modAppDef.SH_SHELF)へ移した。タブに出る画面が
' 「マイ本棚」と「ナレッジ倉庫」の2枚に割れていて、中身が同じデータなのに
' 別物に見えていたのが実機の混乱の元だったため、シートを1枚に統合する。
Private Const CARDS_PER_PAGE As Long = 9
Private Const CARD_W As Double = 215
Private Const CARD_H As Double = 120

Private mGalleryPage As Long
Private mGalleryNames() As String   ' 現在ページのカード順の資料名(クリック解決用)
Private mGalleryCount As Long

' ----------------------------------------------------------------------------
' ShowVaultInput - 登録フォームを描画して表示(SPA遷移)
' ----------------------------------------------------------------------------
Public Sub ShowVaultInput()
    Dim ws As Worksheet
    Set ws = GetOrCreateVaultSheet()
    If ws Is Nothing Then Exit Sub

    On Error GoTo Finish
    Application.ScreenUpdating = False

    ' 冪等再構築
    RemoveVaultShapes ws
    ws.Cells.Clear
    ws.Cells.Font.Name = "Yu Gothic UI"
    ws.Cells.Interior.Color = RGB(243, 244, 246)

    ws.Columns("A").ColumnWidth = 4
    ws.Columns("B").ColumnWidth = 3
    ws.Columns("C:H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 3

    ' カード風の背景(白)。Shapeはセルより必ず手前に描画される(ZOrderは
    ' Shape同士の前後関係にしか効かない)ため、Shapeで背景を作るとラベル等の
    ' セル文字が完全に隠れてしまう(実機報告のバグ)。セルの塗りつぶしで代替する。
    ws.Range("C2:H18").Interior.Color = RGB(255, 255, 255)

    ' タイトル
    With ws.Range("C2:H2")
        .Merge
        .Value = ChrW(&H2795) & " 新規ナレッジの登録"
        .Font.Size = 14
        .Font.Bold = True
    End With
    With ws.Range("C3:H3")
        .Merge
        .Value = "登録した内容はベクトル化され、社内ナレッジ検索の回答に使われます。"
        .Font.Size = 9
        .Font.Color = RGB(107, 114, 128)
    End With

    ' タイトル入力
    With ws.Range("C5:H5")
        .Merge
        .Value = "タイトル(例: 漁船保険の特例について)"
        .Font.Size = 9.5
        .Font.Bold = True
    End With
    With ws.Range("C6:H6")
        .Merge
        .Interior.Color = RGB(249, 250, 251)
        .Borders.LineStyle = 1
        .Borders.Color = RGB(229, 231, 235)
    End With
    ws.Rows(6).RowHeight = 24

    ' 本文入力
    With ws.Range("C7:H7")
        .Merge
        .Value = "ナレッジ本文(またはAIへの修正指示)"
        .Font.Size = 9.5
        .Font.Bold = True
    End With
    With ws.Range("C8:H16")
        .Merge
        .WrapText = True
        .VerticalAlignment = -4160   ' xlTop
        .Interior.Color = RGB(249, 250, 251)
        .Borders.LineStyle = 1
        .Borders.Color = RGB(229, 231, 235)
    End With

    ' タグ入力
    With ws.Range("C17:H17")
        .Merge
        .Value = "タグ(カンマ区切り。例: 約款解釈,特約)"
        .Font.Size = 9.5
        .Font.Bold = True
    End With
    With ws.Range("C18:H18")
        .Merge
        .Interior.Color = RGB(249, 250, 251)
        .Borders.LineStyle = 1
        .Borders.Color = RGB(229, 231, 235)
    End With
    ws.Rows(18).RowHeight = 24

    ' ボタン(登録=青 / キャンセル=白)
    Dim submitBtn As Shape
    Set submitBtn = ws.Shapes.AddShape(5, 420, 388, 140, 32)
    submitBtn.Name = "nxv_submit"
    submitBtn.Line.Visible = 0
    submitBtn.Fill.ForeColor.RGB = RGB(37, 99, 235)
    With submitBtn.TextFrame2
        .TextRange.Text = "ベクトル化して登録"
        .TextRange.Font.Size = 10.5
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    submitBtn.OnAction = "modVault.OnVaultSubmit"

    Dim cancelBtn As Shape
    Set cancelBtn = ws.Shapes.AddShape(5, 320, 388, 90, 32)
    cancelBtn.Name = "nxv_cancel"
    cancelBtn.Fill.ForeColor.RGB = RGB(255, 255, 255)
    cancelBtn.Line.ForeColor.RGB = RGB(229, 231, 235)
    With cancelBtn.TextFrame2
        .TextRange.Text = "キャンセル"
        .TextRange.Font.Size = 10.5
        .TextRange.Font.Fill.ForeColor.RGB = RGB(17, 24, 39)
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    cancelBtn.OnAction = "modVault.OnVaultCancel"

    modUI.FreezeShapePlacement ws   ' 全Shapeを絶対配置に固定(ズレ防止)

    ' SPA遷移(表示してアクティブ化・枠線等は非表示)
    ws.Visible = -1   ' xlSheetVisible
    ws.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ws.Range(CELL_TITLE).Select
    On Error GoTo 0

    ' 正常系はハンドラ本体(Resume)を跨いで後始末へ入る
    ' (Resume はエラーが起きていないと実行時エラー20になる)。
    GoTo FinishCleanup0
Finish:
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FinishCleanup0
FinishCleanup0:
    On Error Resume Next
    Application.ScreenUpdating = True   ' 例外時も必ず画面更新を戻す(暗転固定を防ぐ)
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' OnVaultSubmit - 入力内容を既存取込パイプラインへ流して登録
' ----------------------------------------------------------------------------
Public Sub OnVaultSubmit()
    Dim ws As Worksheet
    Set ws = GetVaultSheet()
    If ws Is Nothing Then Exit Sub

    Dim titleText As String, bodyText As String, tagsText As String
    titleText = Trim$(CStr(ws.Range(CELL_TITLE).Value))
    bodyText = Trim$(CStr(ws.Range(CELL_BODY).Value))
    tagsText = Trim$(CStr(ws.Range(CELL_TAGS).Value))

    If LenB(titleText) = 0 Or LenB(bodyText) = 0 Then
        MsgBox "タイトルと本文を入力してください。", vbExclamation, "Nexus Agent"
        Exit Sub
    End If

    If RegisterKnowledgeText(titleText, TagsHeader(tagsText) & bodyText, tagsText) Then
        MsgBox "ナレッジデータベースに追加され、ベクトル化されました。", vbInformation, "Nexus Agent"
        ClearInputs ws
        CloseVault ws
    Else
        MsgBox "登録に失敗しました。マイ本棚の一覧で状態をご確認ください。", vbExclamation, "Nexus Agent"
    End If
End Sub

Public Sub OnVaultCancel()
    Dim ws As Worksheet
    Set ws = GetVaultSheet()
    If ws Is Nothing Then Exit Sub
    ClearInputs ws
    CloseVault ws
End Sub

' ----------------------------------------------------------------------------
' RegisterKnowledgeText - テキストナレッジを既存パイプラインで登録する共通口
'   (modAppの👎自己学習からも使う)。成功=True(done/partial)。
' ----------------------------------------------------------------------------
Public Function RegisterKnowledgeText(ByVal titleText As String, ByVal bodyText As String, _
                                      ByVal tagsText As String) As Boolean
    On Error GoTo Fail

    Dim tempDir As String: tempDir = Environ$("TEMP")
    If LenB(tempDir) = 0 Then tempDir = Environ$("TMP")
    If LenB(tempDir) = 0 Then Exit Function
    If Right$(tempDir, 1) <> "\" Then tempDir = tempDir & "\"

    Dim filePath As String
    filePath = tempDir & "ナレッジ_" & SanitizeName(titleText) & "_" & _
               Format$(Now, "yyyymmddhhnnss") & ".txt"

    Dim content As String
    content = "【" & titleText & "】" & vbLf & bodyText

    ' UTF-8で保存(modExtractorのtxt読取りがUTF-8のため整合)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2          ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText content
    st.SaveToFile filePath, 2   ' adSaveCreateOverWrite
    st.Close
    Set st = Nothing

    Dim resultStatus As String
    resultStatus = modShelf.IngestFile(filePath, "self")

    ' 一時ファイルは掃除(取込済みなので不要。失敗しても無視)
    On Error Resume Next
    Kill filePath
    On Error GoTo 0

    RegisterKnowledgeText = (resultStatus = "done" Or resultStatus = "partial")
    Exit Function

Fail:
    ' Err の内容は Resume でクリアされるので先に控える。
    Dim failDesc As String: failDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup3
FailCleanup3:
    On Error Resume Next
    modLog.LogError "E0801", "modVault.RegisterKnowledgeText", failDesc
    If Not st Is Nothing Then st.Close
    Set st = Nothing
    On Error GoTo 0
    RegisterKnowledgeText = False
End Function

' ----------------------------------------------------------------------------
' ナレッジ倉庫ギャラリー(設計: 単一Shape=1カード・3列グリッド・ページング。
' Shape増殖なし=毎回同数のカードを描き直す)
' 宣言部(CARDS_PER_PAGE/CARD_W/CARD_H/mGallery*)はモジュール先頭に集約済み。
' 描画先は「マイ本棚」シート(R4要件A)。専用シートはもう作らない。
' ----------------------------------------------------------------------------

' ギャラリーを表示(SPA遷移)。検索語はシートのD3セル(検索バー)から読む。
' uiStep: modUIShelf.EnsureLayoutと同型(2026-07-31 R10-1)。Activate失敗の
' 無言スキップ(ギャラリー無反応の原因)を解消し、E0801を必ず1行残す。
Public Sub ShowVaultGallery()
    Dim ws As Worksheet
    Set ws = GetOrCreateGallerySheet()
    If ws Is Nothing Then Exit Sub

    Dim uiStep As String
    Dim gErrNum As Long, gErrDesc As String
    On Error GoTo Finish
    Application.ScreenUpdating = False
    uiStep = "枠の描画(DrawGalleryFrame)"
    DrawGalleryFrame ws
    uiStep = "カードの描画(RenderGalleryCards)"
    RenderGalleryCards ws
    modUI.FreezeShapePlacement ws   ' 全Shapeを絶対配置に固定(ズレ防止)
    modSkin.BeautifyAll ws          ' フォント統一(Yu Gothic UI)+固定クロムに柔らかい影

    ws.Visible = -1
    uiStep = "シートのアクティブ化(ActivateSheetRobust)"
    If modUI.ActivateSheetRobust(ws, "modVault.ShowVaultGallery") Then
        On Error Resume Next
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.DisplayWorkbookTabs = False
        On Error GoTo Finish
        ' 表示の共通儀式(左端へ戻す/等倍/旧Vaultシートの掃除)。R4要件B。
        uiStep = "表示の共通儀式(PrepareScreenView)"
        modKnowledge.PrepareScreenView ws
    End If
    ' Falseのときは両者をスキップ(ActiveWindowが別シートを向いたまま触らない)。

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
    If gErrNum <> 0 Then
        modLog.LogError "E0801", "modVault.ShowVaultGallery", "[" & uiStep & "] " & gErrDesc, gErrNum
    End If
    Application.ScreenUpdating = True   ' 例外時も必ず画面更新を戻す(暗転固定を防ぐ)
    On Error GoTo 0
End Sub

Public Sub OnVaultSearch()
    mGalleryPage = 0
    Dim ws As Worksheet
    Set ws = GetGallerySheet()
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    RenderGalleryCards ws
    Application.ScreenUpdating = True
End Sub

Public Sub OnVaultPrev()
    If mGalleryPage > 0 Then mGalleryPage = mGalleryPage - 1
    OnVaultSearchKeepPage
End Sub

Public Sub OnVaultNext()
    mGalleryPage = mGalleryPage + 1
    OnVaultSearchKeepPage
End Sub

Private Sub OnVaultSearchKeepPage()
    Dim ws As Worksheet
    Set ws = GetGallerySheet()
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    RenderGalleryCards ws
    Application.ScreenUpdating = True
End Sub


Public Sub OnVaultBackToChat()
    modUI.GoToNexus "modVault.OnVaultBackToChat"
End Sub

' カードクリック: 内容の先頭を表示し、削除も選べる
Public Sub OnVaultCardClick()
    Dim callerName As String
    On Error Resume Next
    callerName = CStr(Application.Caller)
    On Error GoTo 0
    If Left$(callerName, 9) <> "nxg_card_" Then Exit Sub

    Dim idx As Long
    idx = CLng(Val(Mid$(callerName, 10)))
    If idx < 0 Or idx >= mGalleryCount Then Exit Sub

    Dim srcName As String
    srcName = mGalleryNames(idx)

    Dim answer As Long
    answer = MsgBox("『" & srcName & "』" & vbLf & vbLf & _
                    modUtil.SafeLeft(PreviewOf(srcName), 300) & vbLf & vbLf & _
                    "[はい]=削除  /  [いいえ]=" & ChrW(&H26A0) & "ノイズ報告(品質が低いと報告)  /  [キャンセル]=閉じる", _
                    vbYesNoCancel + vbQuestion + vbDefaultButton2, "Nexus Agent - ナレッジ詳細")
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
               vbInformation, "Nexus Agent"
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
    ws.Cells.Interior.Color = modUI.UiColor("bg")
    ws.Cells.Font.Name = "Yu Gothic UI"
    ' A:N を全列ぶん明示する(一覧表モードがK/L未設定だったために、
    ' DrawChromeが使う W=A1:N1 の幅が機種・履歴依存でぶれていた)。
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B:N").ColumnWidth = 12
    ws.Rows("7:412").RowHeight = 15       ' 一覧表モードの可変行高を戻す(1-D: 最終行412まで)

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
        Set icon = ws.Shapes.AddShape(1, cardL, cardT + 40, 640, 60)
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
        Set emsg = ws.Shapes.AddShape(1, cardL, cardT + 104, 640, 46)
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
        Set cta = ws.Shapes.AddShape(5, cardL + 250, cardT + 158, 140, 34)
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
        cta.OnAction = "modVault.OnVaultBackToChat"
    Else
        On Error Resume Next
        ws.Range("B7:G8").UnMerge
        ws.Range("B7:G8").ClearContents
        On Error GoTo 0

        Dim k As Long
        For k = startIdx To endIdx
            Dim slot As Long: slot = k - startIdx
            Dim col As Long: col = slot Mod 3
            Dim rowN As Long: rowN = slot \ 3
            DrawOneCard ws, slot, cardL + col * (CARD_W + 14), cardT + rowN * (CARD_H + 14), _
                        fNames(k), fStats(k)
            mGalleryNames(slot) = fNames(k)
            mGalleryCount = mGalleryCount + 1
        Next k
    End If

    ' ページャ
    Dim pgY As Double: pgY = cardT + 3 * (CARD_H + 14) + 6
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
    prevBtn.OnAction = "modVault.OnVaultPrev"

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
    nextBtn.OnAction = "modVault.OnVaultNext"
End Sub

' 単一Shape=1カード(タイトル太字+プレビュー+日付を1テキストに結合し、
' 部分書式で表現。グループ化しない=軽量・増殖なし)
Private Sub DrawOneCard(ByVal ws As Worksheet, ByVal slot As Long, ByVal x As Double, _
                        ByVal y As Double, ByVal srcName As String, ByVal statLine As String)
    Dim parts() As String: parts = Split(statLine, "|")
    Dim addedAt As String
    If UBound(parts) >= 1 Then addedAt = parts(1)
    Dim chunkN As String
    If UBound(parts) >= 2 Then chunkN = parts(2)

    Dim titleText As String: titleText = modUtil.SafeLeft(srcName, 40)
    Dim previewText As String: previewText = modUtil.SafeLeft(PreviewOf(srcName), 90)
    Dim footText As String
    footText = ChrW(&HD83D) & ChrW(&HDCC5) & " " & ShortStamp(addedAt) & "  ・ " & chunkN & " chunks"

    Dim isExcluded As Boolean
    On Error Resume Next
    isExcluded = modStats.IsGloballyExcluded(srcName)
    On Error GoTo 0

    Dim card As Shape
    Set card = ws.Shapes.AddShape(5, x, y, CARD_W, CARD_H)
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
    card.OnAction = "modVault.OnVaultCardClick"
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

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

Private Function TagsHeader(ByVal tagsText As String) As String
    If LenB(tagsText) = 0 Then Exit Function
    TagsHeader = "タグ: " & tagsText & vbLf & vbLf
End Function

Private Sub ClearInputs(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Range(CELL_TITLE).Value = ""
    ws.Range(CELL_BODY).Value = ""
    ws.Range(CELL_TAGS).Value = ""
    On Error GoTo 0
End Sub

' レビュー1-C: 順序を逆にした。アクティブなシートは隠せず1004になるため、
' 旧実装は「隠そうとして失敗→隠れないままギャラリーを描く」状態だった。
' 先に行き先(登録前に見ていたモード)を描いて前面へ出し、そのあとで隠す。
' ギャラリー固定をやめたのは、一覧表から登録した人が毎回ギャラリーへ
' 飛ばされていたため。
Private Sub CloseVault(ByVal ws As Worksheet)
    On Error Resume Next
    modKnowledge.RefreshCurrent
    ' RenderShelf(一覧表)はシートをアクティブにしないので必ず前面へ出す。
    modUI.GoToNativeSheet modAppDef.SH_SHELF, "modVault(CloseVault)"
    ws.Visible = 2   ' xlSheetVeryHidden
    On Error GoTo 0
End Sub

Private Function SanitizeName(ByVal s As String) As String
    Dim bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)
    Dim t As String: t = s
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        t = Replace(t, CStr(bad(i)), "_")
    Next i
    SanitizeName = modUtil.SafeLeft(Trim$(t), 60)
End Function

Private Function GetVaultSheet() As Worksheet
    On Error Resume Next
    Set GetVaultSheet = ThisWorkbook.Worksheets(VAULT_SHEET)
    On Error GoTo 0
End Function

Private Function GetOrCreateVaultSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = GetVaultSheet()
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = VAULT_SHEET
        On Error GoTo 0
    End If
    Set GetOrCreateVaultSheet = ws
    Exit Function
Fail:
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup26
FailCleanup26:
    If Not ws Is Nothing Then
        On Error Resume Next
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        On Error GoTo 0
    End If
    Set GetOrCreateVaultSheet = Nothing
End Function

Private Sub RemoveVaultShapes(ByVal ws As Worksheet)
    Dim names() As String
    ReDim names(0 To ws.Shapes.count)
    Dim n As Long: n = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 4) = "nxv_" Then
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
