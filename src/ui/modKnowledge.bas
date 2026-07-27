Attribute VB_Name = "modKnowledge"
Option Explicit

' modKnowledge - ナレッジ画面の共通クロム(ヘッダー+モード切替+ツールバー)。
'
' 2026-07-26 再設計(nexus-spec-v1 §2.3 / nexus-ui-final 画面3):
'   実機報告「ナレッジ倉庫とマイ本棚の違いが分からない」への対処。
'   中身は今までどおり2枚のシート(Vault=カードギャラリー / マイ本棚=一覧表)
'   だが、タブは隠してあるので利用者にシートの区別は見えない。両方に
'   まったく同じヘッダーとツールバーを描き、上部のピルで
'   「🃏 ギャラリー / 📋 マイ本棚」を切り替える1画面2モードとして見せる。
'
'   ツールバーの各ボタンは既存のmodVault/modUIShelf/modPack/modShelfSyncへ
'   そのまま配線するだけで、取込・同期・パックのロジックには一切触っていない
'   (仕様書の「既存コード変更最小化」原則)。これまで2画面に重複していた
'   パック出力/取込の入口も、この1本のツールバーに集約した。
'
' 設計の鉄則:
'   ・Shape座標は実セル幾何(Range.Left/.Width/Rows().Top)から導く。
'     旧modVaultのツールバーはx=390,460,528...と決め打ちで、列幅を変えると
'     すぐ画面外へはみ出していた。
'   ・配色は modUI.UiColor()。絵文字はChrW()で組み立てる。
'   ・Shape名は nxk_ 接頭辞(冪等な全削除に使う)。

Private Const HDR_H As Double = 40
Private Const BAR_H As Double = 24
Private Const PILL_W As Double = 88

' 上部クロムが占める行(1..6)。本文はDrawChrome後の ContentTop から下に描く。
Public Const CHROME_ROWS As Long = 6

' DrawChrome - ヘッダー+モードピル+ツールバーを描く(冪等)。
'   mode: "gallery"(ナレッジ倉庫のカード) / "table"(マイ本棚の一覧)
Public Sub DrawChrome(ByVal ws As Worksheet, ByVal mode As String)
    If ws Is Nothing Then Exit Sub
    On Error GoTo Fail

    RemoveChrome ws

    ' 幾何を先に確定させる(順序が逆だとShape座標がズレる)。
    ' 実機報告(2026-07-27)「ボタンが横に並びきらず、右へスクロールしないと
    ' 見えない」への対処。ツールバーを2段(行3・行4)に折り返す。
    ' 横スクロールを要求する時点でUIとして失格なので、幅に収まらなければ
    ' 必ず折り返す実装にしてある(ボタンが増えても破綻しない)。
    ws.Rows(1).RowHeight = HDR_H
    ws.Rows(2).RowHeight = 6
    ws.Rows(3).RowHeight = BAR_H
    ws.Rows(4).RowHeight = BAR_H + 2
    ws.Rows(5).RowHeight = 22
    ws.Rows(6).RowHeight = 8

    ' 罫線と行列番号を隠す。ここを消さないと、どれだけ整えても
    ' 画面が「Excelのシート」にしか見えない(実機要望: エクセル感を消す)。
    On Error Resume Next
    If ThisWorkbook.ActiveSheet Is ws Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
    End If
    On Error GoTo Fail

    Dim L As Double, W As Double
    L = ws.Range("A1").Left
    W = ws.Range("A1:N1").Width

    ' --- ヘッダーバー ---
    Dim hdr As Shape
    Set hdr = ws.Shapes.AddShape(5, L, 0, W, HDR_H)
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
    Pill ws, "nxk_back", ChrW(&H2190) & " Hub", L + 8, 72, _
         "modKnowledge.OnBackHub", True

    ' --- モード切替ピル(右肩) ---
    Dim md As String: md = LCase$(mode)
    Dim isTable As Boolean: isTable = (md = "table")
    Dim isShared As Boolean: isShared = (md = "shared")
    Dim px As Double: px = L + W - 8 - PILL_W
    Pill ws, "nxk_m_shared", ChrW(&HD83C) & ChrW(&HDF81) & " みんなの解決事例", px, PILL_W, _
         "modKnowledge.OnGoShared", isShared
    px = px - 6 - PILL_W
    Pill ws, "nxk_m_table", ChrW(&HD83D) & ChrW(&HDCCB) & " マイ本棚", px, PILL_W, _
         "modKnowledge.OnGoTable", isTable
    px = px - 6 - PILL_W
    Pill ws, "nxk_m_gallery", ChrW(&HD83C) & ChrW(&HDCCF) & " ギャラリー", px, PILL_W, _
         "modKnowledge.OnGoGallery", (Not isTable) And (Not isShared)

    On Error Resume Next
    modTelemetry.TrackScreen "knowledge"
    On Error GoTo Fail

    ' --- ツールバー(行3の帯) ---
    DrawToolbar ws, isTable, isShared, L, W

    On Error Resume Next
    modUI.FreezeShapePlacement ws
    On Error GoTo 0
    Exit Sub

Fail:
    modLog.LogError "E0801", "modKnowledge.DrawChrome", Err.Description, Err.Number
End Sub

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

Private Sub DrawToolbar(ByVal ws As Worksheet, ByVal isTable As Boolean, _
                        ByVal isShared As Boolean, ByVal L As Double, ByVal W As Double)
    Dim caps As Variant, acts As Variant, widths As Variant
    caps = Array(ChrW(&HD83D) & ChrW(&HDD0D) & " 検索", _
                 ChrW(&H2795) & " 登録", _
                 ChrW(&HD83D) & ChrW(&HDCC1) & " 追加", _
                 ChrW(&HD83D) & ChrW(&HDCE6) & " パック出力", _
                 ChrW(&HD83D) & ChrW(&HDCE5) & " パック取込", _
                 ChrW(&HD83D) & ChrW(&HDD04) & " 同期", _
                 ChrW(&HD83D) & ChrW(&HDCC2) & " フォルダ", _
                 ChrW(&HD83D) & ChrW(&HDDD1) & " 削除", _
                 ChrW(&HD83D) & ChrW(&HDCA1) & " みんなの困りごと", _
                 ChrW(&HD83D) & ChrW(&HDCE1) & " 部門チャンネル", _
                 ChrW(&HD83D) & ChrW(&HDCAC) & " チャットへ")
    acts = Array("OnSearch", "OnRegister", "OnAddFiles", "OnPackOut", "OnPackIn", _
                 "OnSync", "OnPickFolder", "OnDelete", "OnGapBoard", "OnChannels", "OnToChat")
    widths = Array(62, 58, 58, 80, 80, 54, 68, 58, 104, 96, 76)
    ' 発行ボタンは、発行キーが設定されている端末にだけ出す。
    ' 一般利用者の画面に「押してはいけないボタン」を置かない。
    Dim canPub As Boolean
    On Error Resume Next
    canPub = modPublish.CanPublish()
    On Error GoTo 0

    Dim barTop As Double: barTop = ws.Rows(3).Top
    Dim x As Double: x = L + 8
    Dim rowIdx As Long: rowIdx = 0
    Dim maxX As Double: maxX = L + W - 8

    ' みんなのQ&Aモードは操作が全く違う(選択と取り込み)。資料管理用の
    ' ボタンを並べても押しどころが分からなくなるので、専用の並びにする。
    If isShared Then
        SharedToolbar ws, barTop, L + 8, maxX
        Exit Sub
    End If

    Dim i As Long
    For i = 0 To 10
        ' 検索はギャラリー専用、削除は一覧表専用(押しても何も起きないボタンを
        ' 見せない=実機報告「どっちで押せばいいか分からない」への対処)。
        Dim skip As Boolean
        skip = (i = 0 And isTable) Or (i = 7 And Not isTable)
        If Not skip Then
            ' 幅に収まらなくなったら次の段へ折り返す(横スクロールさせない)。
            If x + CDbl(widths(i)) > maxX And rowIdx = 0 Then
                rowIdx = 1
                x = L + 8
            End If
            ' 1個の1004で残りを道連れにしない。
            On Error Resume Next
            Dim btn As Shape
            Set btn = ws.Shapes.AddShape(5, x, barTop + rowIdx * (BAR_H + 2), _
                                         CDbl(widths(i)), BAR_H)
            If Err.Number = 0 And Not btn Is Nothing Then
                btn.Name = "nxk_tb" & i
                btn.Adjustments(1) = 0.35
                btn.Line.Visible = -1
                btn.Line.Weight = 0.75
                btn.Line.ForeColor.RGB = modUI.UiColor("border")
                btn.Fill.ForeColor.RGB = modUI.UiColor("surface")
                With btn.TextFrame2
                    .WordWrap = -1
                    .TextRange.Text = CStr(caps(i))
                    .TextRange.Font.Size = 8.5
                    .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
                    .TextRange.ParagraphFormat.Alignment = 2
                    .VerticalAnchor = 3
                    .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
                End With
                modSkin.ApplyLightShadow btn
                btn.OnAction = "modKnowledge." & CStr(acts(i))
            End If
            Set btn = Nothing
            Err.Clear
            On Error GoTo 0
            x = x + CDbl(widths(i)) + 5
        End If
    Next i

    ' 発行者だけに見える「正典を発行」。
    If canPub Then
        If x + 104 > maxX And rowIdx = 0 Then
            rowIdx = 1
            x = L + 8
        End If
        On Error Resume Next
        Dim pb As Shape
        Set pb = ws.Shapes.AddShape(5, x, barTop + rowIdx * (BAR_H + 2), 104, BAR_H)
        If Not pb Is Nothing Then
            pb.Name = "nxk_tbpub"
            pb.Adjustments(1) = 0.35
            pb.Line.Visible = 0
            pb.Fill.ForeColor.RGB = modUI.UiColor("primary")
            With pb.TextFrame2
                .WordWrap = -1
                .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCE4) & " 正典を発行"
                .TextRange.Font.Size = 8.5
                .TextRange.Font.Bold = -1
                .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .TextRange.ParagraphFormat.Alignment = 2
                .VerticalAnchor = 3
                .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
            End With
            pb.OnAction = "modKnowledge.OnPublish"
        End If
        Set pb = Nothing
        Err.Clear
        On Error GoTo 0
        x = x + 109

        ' 運営向けの利用状況。発行者=運営なので同じ条件で出す。
        If x + 88 > maxX And rowIdx = 0 Then
            rowIdx = 1
            x = L + 8
        End If
        On Error Resume Next
        Dim rp As Shape
        Set rp = ws.Shapes.AddShape(5, x, barTop + rowIdx * (BAR_H + 2), 88, BAR_H)
        If Not rp Is Nothing Then
            rp.Name = "nxk_tbrep"
            rp.Adjustments(1) = 0.35
            rp.Line.Visible = -1
            rp.Line.Weight = 0.75
            rp.Line.ForeColor.RGB = modUI.UiColor("border")
            rp.Fill.ForeColor.RGB = modUI.UiColor("surface")
            With rp.TextFrame2
                .WordWrap = -1
                .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCCA) & " 利用状況"
                .TextRange.Font.Size = 8.5
                .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
                .TextRange.ParagraphFormat.Alignment = 2
                .VerticalAnchor = 3
                .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
            End With
            rp.OnAction = "modHub.OnOwnerReport"
        End If
        Set rp = Nothing
        Err.Clear
        On Error GoTo 0
        x = x + 93
    End If

    ' 画像解析が使える環境でだけスクショ取込を出す(無効環境で「押したら
    ' 断られるボタン」を見せない。既存modUIShelfの方針をそのまま踏襲)。
    On Error Resume Next
    If modFeatures.FeatureEnabled("vision") Then
        If x + 84 > maxX And rowIdx = 0 Then
            rowIdx = 1
            x = L + 8
        End If
        Dim sc As Shape
        Set sc = ws.Shapes.AddShape(5, x, barTop + rowIdx * (BAR_H + 2), 84, BAR_H)
        If Not sc Is Nothing Then
            sc.Name = "nxk_tbshot"
            sc.Adjustments(1) = 0.35
            sc.Line.Visible = -1
            sc.Line.Weight = 0.75
            sc.Line.ForeColor.RGB = modUI.UiColor("border")
            sc.Fill.ForeColor.RGB = modUI.UiColor("surface")
            With sc.TextFrame2
                .WordWrap = -1
                .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCF8) & " スクショ取込"
                .TextRange.Font.Size = 8.5
                .TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
                .TextRange.ParagraphFormat.Alignment = 2
                .VerticalAnchor = 3
                .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
            End With
            sc.OnAction = "modUIShelf.OnIngestScreenshot"
        End If
    End If
    On Error GoTo 0
End Sub

' みんなのQ&A専用ツールバー。
Private Sub SharedToolbar(ByVal ws As Worksheet, ByVal barTop As Double, _
                          ByVal x0 As Double, ByVal maxX As Double)
    Dim caps As Variant, acts As Variant, widths As Variant
    caps = Array(ChrW(&H2713) & " 選択を取り込む", "すべて選ぶ", "選択を解除", _
                 ChrW(&H2190) & " 前", "次 " & ChrW(&H2192), _
                 ChrW(&HD83D) & ChrW(&HDCA1) & " みんなの困りごと", _
                 ChrW(&HD83D) & ChrW(&HDCAC) & " チャットへ")
    acts = Array("modShared.OnImportSelected", "modShared.OnSelectAll", _
                 "modShared.OnSelectNone", "modShared.OnPrevPage", "modShared.OnNextPage", _
                 "modKnowledge.OnGapBoard", "modKnowledge.OnToChat")
    widths = Array(112, 72, 72, 44, 44, 104, 76)

    Dim x As Double: x = x0
    Dim rowIdx As Long
    Dim i As Long
    For i = 0 To 6
        If x + CDbl(widths(i)) > maxX And rowIdx = 0 Then
            rowIdx = 1
            x = x0
        End If
        On Error Resume Next
        Dim btn As Shape
        Set btn = ws.Shapes.AddShape(5, x, barTop + rowIdx * (BAR_H + 2), _
                                     CDbl(widths(i)), BAR_H)
        If Err.Number = 0 And Not btn Is Nothing Then
            btn.Name = "nxk_sb" & i
            btn.Adjustments(1) = 0.35
            btn.Line.Visible = -1
            btn.Line.Weight = 0.75
            btn.Line.ForeColor.RGB = modUI.UiColor("border")
            If i = 0 Then
                btn.Fill.ForeColor.RGB = modUI.UiColor("accent")
            Else
                btn.Fill.ForeColor.RGB = modUI.UiColor("surface")
            End If
            With btn.TextFrame2
                .WordWrap = -1
                .TextRange.Text = CStr(caps(i))
                .TextRange.Font.Size = 8.5
                If i = 0 Then
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
            btn.OnAction = CStr(acts(i))
        End If
        Set btn = Nothing
        Err.Clear
        On Error GoTo 0
        x = x + CDbl(widths(i)) + 5
    Next i
End Sub

' ヘッダー上のピル。active:=Trueで「今いるモード」を塗りつぶして示す。
Private Sub Pill(ByVal ws As Worksheet, ByVal shapeName As String, _
                 ByVal caption As String, ByVal x As Double, ByVal w As Double, _
                 ByVal action As String, ByVal active As Boolean)
    On Error Resume Next
    Dim p As Shape
    Set p = ws.Shapes.AddShape(5, x, (HDR_H - 26) / 2, w, 26)
    If p Is Nothing Then Exit Sub
    p.Name = shapeName
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
        ' 旧マイ本棚のボタン(btn_)も一緒に消す。残すと新ツールバーの上に
        ' 浮いたまま二重表示になる(Hub移植時に踏んだのと同じ罠)。
        If Left$(shp.Name, 4) = "nxk_" Or Left$(shp.Name, 4) = "btn_" Then
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

Public Sub OnGoGallery()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modVault.ShowVaultGallery
    On Error GoTo 0
    modUiLock.Leave
End Sub

' みんなのQ&A(選択式取り込み)へ。
Public Sub OnGoShared()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUiLock.Leave
    modShared.Show
    Exit Sub
End Sub

Public Sub OnGoTable()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUIShelf.EnsureLayout
    modUI.GoToNativeSheet modAppDef.SH_SHELF, "modKnowledge.OnGoTable"
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnBackHub()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modHub.EnsureHubLayout activate:=True
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnToChat()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUI.GoToNexus "modKnowledge.OnToChat"
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnSearch()
    modVault.OnVaultSearch
End Sub

' ----------------------------------------------------------------------------
' 💡 みんなの困りごと: 組織で答えが見つからなかった質問の一覧。
'   資料を書ける人(商品部)がここを見て、その場でナレッジを書けるようにする。
'   営業の「分からない」が、商品部の「書くべきこと」に直結する一番短い経路。
' ----------------------------------------------------------------------------
Public Sub OnGapBoard()
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
' 受け取った「解決済みQ&A」を自分の本棚へ取り込む。
'   埋め込みAPIを使う重い処理なので、起動時ではなく利用者が押したときだけ実行。
'   取り込むと、次から同じ質問に「人が確認済みの答え」で応えられるようになる。
' ----------------------------------------------------------------------------
Public Sub OnImportSharedQA_Legacy()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done

    Dim total As Long
    On Error Resume Next
    total = modInsight.PendingQACount()
    On Error GoTo Done

    If total < 1 Then
        modUiLock.Leave
        MsgBox "取り込める新しいQ&Aはありません。", vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    If MsgBox(total & " 件の「みんなが解決したQ&A」を本棚に取り込みます。" & vbCrLf & _
              "取り込むと、次から同じ内容を質問したときに出典つきで答えられます。" & vbCrLf & _
              "(件数によっては1〜2分かかります)", _
              vbOKCancel + vbQuestion, modAppDef.APP_NAME) <> vbOK Then GoTo Done

    Dim okN As Long, i As Long
    For i = total To 1 Step -1        ' 後ろから処理(消化印で番号が詰まるため)
        Dim author As String, qText As String, aText As String, srcText As String
        Dim rowIdx As Long
        On Error Resume Next
        If modInsight.PendingQAAt(i, author, qText, aText, srcText, rowIdx) Then
            If modVault.RegisterKnowledgeText( _
                   "解決済みQ&A: " & modUtil.SafeLeft(qText, 40), _
                   modInsight.QABodyText(author, qText, aText, srcText), _
                   "解決済みQ&A," & author) Then
                modInsight.MarkQAConsumed rowIdx
                okN = okN + 1
            End If
        End If
        On Error GoTo Done
    Next i

    modUiLock.Leave
    MsgBox okN & " 件を本棚に取り込みました。" & vbCrLf & _
           "同じことで困っている人が、次からはすぐ答えにたどり着けます。", _
           vbInformation, modAppDef.APP_NAME
    Exit Sub
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
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done

    Dim all As String, cur As String
    On Error Resume Next
    all = modChannel.ListChannels()
    cur = modChannel.ActiveChannel()
    On Error GoTo Done

    If LenB(all) = 0 Then
        modUiLock.Leave
        MsgBox "部門の公式ナレッジがまだ1つも見つかりません。" & vbCrLf & vbCrLf & _
            "各部門が正典を発行すると、ここに一覧が出ます。" & vbCrLf & _
            "(共有フォルダの channels\<部門名>\ に pack.xlsx と version.txt が置かれる形です)" & vbCrLf & vbCrLf & _
            "共有フォルダ自体が未設定の場合は、Hubのお知らせから設定してください。", _
            vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    Dim body As String
    body = "聞きたい分野の部門を1つ選んでください。" & vbCrLf & _
           "選ぶと、その部門の公式ナレッジが本棚に読み込まれます。" & vbCrLf & vbCrLf
    If LenB(cur) > 0 Then
        body = body & "  いま接続中: 【" & cur & "】" & vbCrLf & vbCrLf
    Else
        body = body & "  いま接続中: (なし)" & vbCrLf & vbCrLf
    End If

    body = body & "■ 選べる部門" & vbCrLf
    Dim parts() As String: parts = Split(all, "|")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim mark As String
        If StrComp(parts(i), cur, vbTextCompare) = 0 Then
            mark = ChrW(&H25CF) & " "
        Else
            mark = ChrW(&H25CB) & " "
        End If
        body = body & "  " & mark & parts(i)
        If StrComp(parts(i), cur, vbTextCompare) = 0 Then
            If modChannel.RemoteVersion(parts(i)) <> modChannel.LocalVersion(parts(i)) Then
                body = body & "  ← 更新があります"
            End If
        End If
        body = body & vbCrLf
    Next i

    body = body & vbCrLf & "■ 本棚の使用量: " & modChannel.ChunkUsagePercent() & "%" & vbCrLf & vbCrLf & _
           "【切り替えについて】" & vbCrLf & _
           "  ・常時つないでおけるのは1部門だけです" & vbCrLf & _
           "  ・切り替えると前の部門の内容は本棚から外れます" & vbCrLf & _
           "    (あなたが自分で入れた資料は消えません)" & vbCrLf & _
           "  ・分量によっては読み込みに1〜2分かかります" & vbCrLf & vbCrLf & _
           "接続する部門名を入力してください(空欄で閉じる)。"

    Dim ans As String
    ans = Trim$(InputBox(body, modAppDef.APP_NAME & " - 部門の公式ナレッジ"))
    If LenB(ans) = 0 Then GoTo Done

    ' 入力された名前が実在するか確認する(打ち間違いで無言で失敗させない)。
    Dim found As Boolean
    For i = LBound(parts) To UBound(parts)
        If StrComp(parts(i), ans, vbTextCompare) = 0 Then
            ans = parts(i)
            found = True
            Exit For
        End If
    Next i
    If Not found Then
        modUiLock.Leave
        MsgBox "「" & ans & "」という部門は見つかりませんでした。" & vbCrLf & _
               "一覧に表示されている名前をそのまま入力してください。", _
               vbExclamation, modAppDef.APP_NAME
        Exit Sub
    End If

    modUiLock.Leave
    DoSwitch ans
    Exit Sub
Done:
    modUiLock.Leave
End Sub

' 部門の切り替え。時間がかかる処理なので、始まる前に必ず予告する。
' 「押したあと固まったように見える」が離脱の最大要因なので、
' 待ち時間の見込みを先に言い、終わったら結果を必ず出す。
Private Sub DoSwitch(ByVal chName As String)
    On Error Resume Next

    If MsgBox("【" & chName & "】に接続します。" & vbCrLf & vbCrLf & _
              "読み込みの間、画面が止まったように見えることがあります。" & vbCrLf & _
              "そのままお待ちください(分量によっては1〜2分)。" & vbCrLf & vbCrLf & _
              "実行しますか?", vbOKCancel + vbQuestion, _
              modAppDef.APP_NAME) <> vbOK Then Exit Sub

    Application.Cursor = 2                      ' xlWait
    Application.StatusBar = "【" & chName & "】を読み込んでいます..."

    Dim got As Long
    got = modChannel.SwitchTo(chName)

    Application.Cursor = -4143                  ' xlDefault
    Application.StatusBar = False

    If got = -1 Then
        MsgBox "【" & chName & "】は既に最新の状態で接続されています。" & vbCrLf & _
               "そのまま質問できます。", vbInformation, modAppDef.APP_NAME
    ElseIf got > 0 Then
        On Error Resume Next
        modHub.EnsureHubLayout                  ' 使用量タイルと接続中表示を更新
        On Error GoTo 0
        MsgBox "【" & chName & "】に接続しました(" & got & " 件)。" & vbCrLf & vbCrLf & _
               "この分野の質問に、出典つきで答えられます。" & vbCrLf & _
               "別の分野を聞きたくなったら、また部門を切り替えてください。", _
               vbInformation, modAppDef.APP_NAME
    Else
        MsgBox "【" & chName & "】を読み込めませんでした。" & vbCrLf & vbCrLf & _
               "・共有フォルダにつながっているか" & vbCrLf & _
               "・その部門がまだ正典を発行していないか" & vbCrLf & _
               "をご確認ください。もう一度試すと成功することもあります。", _
               vbExclamation, modAppDef.APP_NAME
    End If
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 正典の発行(1操作で完結)。
'   以前は「保存先を手で貼り付けて保存 → もう一度発行を押す」の2段だった。
'   実機で共有フォルダに channels\<部門>\ が作られず止まった原因がこれ。
'   保存先はアプリが知っているのだから、人に貼らせる理由が無い。
'   ボタン1回で 書き出し → 旧版退避 → 配置 → version.txt まで通す。
' ----------------------------------------------------------------------------
Public Sub OnPublish()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done

    Dim chName As String
    chName = Trim$(InputBox( _
        "部門の公式ナレッジ(正典)を発行します。" & vbCrLf & vbCrLf & _
        "【ボタン1回で完了します】" & vbCrLf & _
        "  今の本棚の内容がそのまま部内へ配られます。" & vbCrLf & _
        "  保存先を選ぶ操作は要りません。" & vbCrLf & vbCrLf & _
        "【まちがえても大丈夫です】" & vbCrLf & _
        "  前の版は自動で保存されます。もう一度この画面を開いて" & vbCrLf & _
        "  「いいえ」を選ぶと、直前の版に戻せます。" & vbCrLf & vbCrLf & _
        "発行する部門名を入力してください(例: 商品部)。" & vbCrLf & _
        "※ 更新するときも、前と同じ名前を入れてください。", _
        modAppDef.APP_NAME & " - 正典を発行"))
    If LenB(chName) = 0 Then GoTo Done

    If Not modPublish.VerifyKey(InputBox( _
            "発行キーを入力してください。" & vbCrLf & _
            "(分からない場合は、このツールの管理担当者に確認してください)", _
            modAppDef.APP_NAME & " - 発行キー")) Then
        modUiLock.Leave
        MsgBox "発行キーが違います。発行は行いませんでした。", vbExclamation, modAppDef.APP_NAME
        Exit Sub
    End If

    ' 共有フォルダに書ける状態かを、書き出す前に確かめる。
    ' 権限が無いフォルダを指しているのが実機で一番多い失敗なので、
    ' 重い書き出しをやってから失敗させない。
    Dim dest As String
    On Error Resume Next
    dest = modPublish.PackDestPath(chName)
    On Error GoTo Done
    If LenB(dest) = 0 Then
        modUiLock.Leave
        MsgBox "共有フォルダが未設定です。" & vbCrLf & _
               "Hubのお知らせから設定してから、もう一度お試しください。", _
               vbExclamation, modAppDef.APP_NAME
        Exit Sub
    End If

    Dim total As Long
    On Error Resume Next
    total = modShelf.TotalChunks()
    Dim log_ As String: log_ = modPublish.RecentLog(chName)
    On Error GoTo Done

    Dim msg As String
    msg = "【" & chName & "】として発行します。" & vbCrLf & vbCrLf & _
          "  今の本棚: " & total & " チャンク" & vbCrLf & _
          "  配置先: " & dest & vbCrLf & vbCrLf & _
          "【発行前の確認】" & vbCrLf & _
          "  ・正典には『確認済みQ&A・要点』を入れてください" & vbCrLf & _
          "  ・個人情報が見つかった場合は自動で中止します" & vbCrLf & vbCrLf
    If LenB(log_) > 0 Then msg = msg & "【この部門の発行履歴】" & vbCrLf & log_ & vbCrLf
    msg = msg & "このまま発行しますか?" & vbCrLf & _
          "(「いいえ」= 直前の版に戻す操作に進みます)"

    Dim ans As VbMsgBoxResult
    ans = MsgBox(msg, vbYesNoCancel + vbQuestion, modAppDef.APP_NAME & " - 発行の確認")
    If ans = vbCancel Then GoTo Done
    If ans = vbNo Then
        modUiLock.Leave
        DoRollback chName
        Exit Sub
    End If

    ' ここから先は待たせるので、必ず砂時計と進捗を出す。
    On Error Resume Next
    Application.Cursor = 2
    Application.StatusBar = "【" & chName & "】を書き出しています..."
    modPublish.ArchiveCurrent chName          ' 旧版を退避(戻せるように)
    On Error GoTo Done

    Dim wrote As Long
    Dim ok As Boolean
    On Error Resume Next
    ok = modPack.ExportPackToFile(dest, "", True, wrote)
    Application.Cursor = -4143
    Application.StatusBar = False
    On Error GoTo Done

    If Not ok Then
        modUiLock.Leave
        MsgBox "発行できませんでした。" & vbCrLf & vbCrLf & _
               "・共有フォルダに書き込む権限があるか" & vbCrLf & _
               "・パスが正しいか(" & dest & ")" & vbCrLf & _
               "をご確認ください。", vbExclamation, modAppDef.APP_NAME
        Exit Sub
    End If

    ' pack.xlsx を置き終えてから version.txt を書く。順序が逆だと
    ' 「新しい版番号なのに中身が古い」を掴む人が出る。
    Dim ver As String
    On Error Resume Next
    ver = modPublish.FinalizePublish(chName, wrote)
    modStats.Bump "publish_total"
    On Error GoTo Done

    modUiLock.Leave
    If LenB(ver) > 0 Then
        MsgBox "発行しました。" & vbCrLf & vbCrLf & _
               "  部門: " & chName & vbCrLf & _
               "  件数: " & wrote & " チャンク" & vbCrLf & _
               "  版: " & ver & vbCrLf & vbCrLf & _
               "部内の全員が、次にファイルを開いたときに受け取れます。" & vbCrLf & _
               "内容に誤りが見つかったら、この画面から「いいえ」で戻せます。", _
               vbInformation, modAppDef.APP_NAME
    Else
        MsgBox "ファイルは置けましたが、版の記録に失敗しました。" & vbCrLf & _
               "共有フォルダの書き込み権限をご確認のうえ、もう一度発行してください。", _
               vbExclamation, modAppDef.APP_NAME
    End If
    Exit Sub
Done:
    On Error Resume Next
    Application.Cursor = -4143
    Application.StatusBar = False
    On Error GoTo 0
    modUiLock.Leave
End Sub

' 直前の版に戻す。事故を止める最終手段なので、操作は最短手数にする。
Private Sub DoRollback(ByVal chName As String)
    On Error Resume Next
    Dim list_ As String
    list_ = modPublish.ArchiveList(chName)
    If LenB(list_) = 0 Then
        MsgBox "戻せる過去版がありません(まだ1度も発行していない部門です)。", _
               vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    Dim parts() As String: parts = Split(list_, "|")
    Dim newest As String: newest = parts(0)

    If MsgBox("【" & chName & "】を直前の版に戻します。" & vbCrLf & vbCrLf & _
              "  戻す版: " & newest & vbCrLf & vbCrLf & _
              "戻すと、部内の全員が次にファイルを開いたときに" & vbCrLf & _
              "その版へ自動で置き換わります(誤った内容は各PCから消えます)。" & vbCrLf & vbCrLf & _
              "実行しますか?", vbOKCancel + vbExclamation, _
              modAppDef.APP_NAME & " - 直前の版に戻す") <> vbOK Then Exit Sub

    If modPublish.Rollback(chName, newest) Then
        MsgBox "戻しました。全員に自動で配信されます。", vbInformation, modAppDef.APP_NAME
    Else
        MsgBox "戻せませんでした。共有フォルダへの書き込み権限をご確認ください。", _
               vbExclamation, modAppDef.APP_NAME
    End If
    On Error GoTo 0
End Sub

Public Sub OnRegister()
    modVault.ShowVaultInput
End Sub

Public Sub OnAddFiles()
    On Error Resume Next
    modShelf.AddFilesViaDialog
    On Error GoTo 0
    RefreshCurrent
End Sub

Public Sub OnPackOut()
    modPack.ExportPackDialog
End Sub

Public Sub OnPackIn()
    modPack.ImportPackDialog
    RefreshCurrent
End Sub

Public Sub OnSync()
    modShelfSync.SyncNow
    RefreshCurrent
End Sub

Public Sub OnPickFolder()
    modShelfSync.PickShelfFolder
    RefreshCurrent
End Sub

Public Sub OnDelete()
    modUIShelf.OnDeleteSource
    RefreshCurrent
End Sub

' 今表示しているモードだけを描き直す(モード切替をまたいで表示がズレないよう、
' 資料を足した/消した直後は必ずここを通す)。
Private Sub RefreshCurrent()
    On Error Resume Next
    If ThisWorkbook.ActiveSheet.Name = modAppDef.SH_SHELF Then
        modUIShelf.RenderShelf
    Else
        modVault.ShowVaultGallery
    End If
    On Error GoTo 0
End Sub
