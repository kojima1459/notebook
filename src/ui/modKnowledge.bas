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
    ws.Rows(1).RowHeight = HDR_H
    ws.Rows(2).RowHeight = 6
    ws.Rows(3).RowHeight = BAR_H
    ws.Rows(4).RowHeight = 6
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

    Pill ws, "nxk_back", ChrW(&H2190) & " Hub", L + 8, 62, _
         "modKnowledge.OnBackHub", False

    ' --- モード切替ピル(右肩) ---
    Dim md As String: md = LCase$(mode)
    Dim isTable As Boolean: isTable = (md = "table")
    Dim isShared As Boolean: isShared = (md = "shared")
    Dim px As Double: px = L + W - 8 - PILL_W
    Pill ws, "nxk_m_shared", ChrW(&HD83C) & ChrW(&HDF81) & " みんなのQ&A", px, PILL_W, _
         "modKnowledge.OnGoShared", isShared
    px = px - 6 - PILL_W
    Pill ws, "nxk_m_table", ChrW(&HD83D) & ChrW(&HDCCB) & " マイ本棚", px, PILL_W, _
         "modKnowledge.OnGoTable", isTable
    px = px - 6 - PILL_W
    Pill ws, "nxk_m_gallery", ChrW(&HD83C) & ChrW(&HDCCF) & " ギャラリー", px, PILL_W, _
         "modKnowledge.OnGoGallery", (Not isTable) And (Not isShared)

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

    Dim barTop As Double: barTop = ws.Rows(3).Top
    Dim x As Double: x = L + 8

    ' みんなのQ&Aモードは操作が全く違う(選択と取り込み)。資料管理用の
    ' ボタンを並べても押しどころが分からなくなるので、専用の並びにする。
    If isShared Then
        SharedToolbar ws, barTop, x
        Exit Sub
    End If

    Dim i As Long
    For i = 0 To 10
        ' 検索はギャラリー専用、削除は一覧表専用(押しても何も起きないボタンを
        ' 見せない=実機報告「どっちで押せばいいか分からない」への対処)。
        Dim skip As Boolean
        skip = (i = 0 And isTable) Or (i = 7 And Not isTable)
        If Not skip Then
            ' 1個の1004で残りを道連れにしない。
            On Error Resume Next
            Dim btn As Shape
            Set btn = ws.Shapes.AddShape(5, x, barTop, CDbl(widths(i)), BAR_H)
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

    ' 画像解析が使える環境でだけスクショ取込を出す(無効環境で「押したら
    ' 断られるボタン」を見せない。既存modUIShelfの方針をそのまま踏襲)。
    On Error Resume Next
    If modFeatures.FeatureEnabled("vision") Then
        Dim sc As Shape
        Set sc = ws.Shapes.AddShape(5, x, barTop, 84, BAR_H)
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
Private Sub SharedToolbar(ByVal ws As Worksheet, ByVal barTop As Double, ByVal x0 As Double)
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
    Dim i As Long
    For i = 0 To 6
        On Error Resume Next
        Dim btn As Shape
        Set btn = ws.Shapes.AddShape(5, x, barTop, CDbl(widths(i)), BAR_H)
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

    Dim all As String, pend As String
    On Error Resume Next
    all = modChannel.ListChannels()
    pend = modChannel.PendingUpdates()
    On Error GoTo Done

    If LenB(all) = 0 Then
        modUiLock.Leave
        MsgBox "部門チャンネルが1つも見つかりませんでした。" & vbCrLf & vbCrLf & _
            "各部門が正典パックを発行すると、ここに一覧が出ます。" & vbCrLf & _
            "(共有フォルダの channels\<部門名>\ に pack.xlsx と version.txt を置く形です)" & vbCrLf & _
            "共有フォルダ自体が未設定の場合は、Hubのお知らせから設定してください。", _
            vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    Dim body As String
    body = "■ 使えるチャンネル(購読すると本棚に正典が入り、初日から答えが返ります)" & vbCrLf
    Dim parts() As String: parts = Split(all, "|")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim mark As String
        If modChannel.IsSubscribed(parts(i)) Then
            mark = "[購読中]"
        Else
            mark = "[未購読]"
        End If
        body = body & "  " & mark & " " & parts(i) & vbCrLf
    Next i

    body = body & vbCrLf & "■ 本棚の使用量: " & modChannel.ChunkUsagePercent() & "%"
    If modChannel.IsBudgetTight() Then
        body = body & "  ← 8割を超えています。使っていないチャンネルの購読を外してください。"
    End If
    body = body & vbCrLf & vbCrLf

    If LenB(pend) > 0 Then
        body = body & "■ 更新があります: " & Replace(pend, "|", " / ") & vbCrLf & vbCrLf & _
               "「はい」で最新版に更新します(古い版は自動で置き換えます)。"
        If MsgBox(body, vbYesNo + vbQuestion, modAppDef.APP_NAME & " - 部門チャンネル") = vbYes Then
            Dim got As Long
            modUiLock.Leave
            got = modChannel.SyncSubscribed()
            MsgBox got & " 件の内容を最新版に更新しました。", vbInformation, modAppDef.APP_NAME
            Exit Sub
        End If
        GoTo Done
    End If

    body = body & "購読を変更しますか?(チャンネル名を入力すると購読/解除が切り替わります)"
    Dim ans As String
    ans = InputBox(body, modAppDef.APP_NAME & " - 部門チャンネル")
    If LenB(Trim$(ans)) = 0 Then GoTo Done

    modUiLock.Leave
    ToggleChannel Trim$(ans)
    Exit Sub
Done:
    modUiLock.Leave
End Sub

' 購読の切り替え。解除時はそのチャンネル由来のチャンクを本棚から取り除く
' (残すと使っていない知識が検索を薄めるうえ、チャンク上限も食い続ける)。
Private Sub ToggleChannel(ByVal chName As String)
    On Error Resume Next
    If modChannel.IsSubscribed(chName) Then
        If MsgBox("「" & chName & "」の購読を解除します。" & vbCrLf & _
                  "このチャンネル由来の内容は本棚から取り除かれます" & vbCrLf & _
                  "(あなたが自分で入れた資料は消えません)。よろしいですか?", _
                  vbOKCancel + vbQuestion, modAppDef.APP_NAME) <> vbOK Then Exit Sub
        Dim removed As Long
        removed = modChannel.PurgeChannelChunks(chName)
        modChannel.Unsubscribe chName
        modStats.SetStatText "ch:" & LCase$(chName), ""
        MsgBox "購読を解除し、" & removed & " 件を本棚から取り除きました。", _
               vbInformation, modAppDef.APP_NAME
    Else
        modChannel.Subscribe chName
        Dim got As Long
        got = modChannel.SyncChannel(chName)
        If got > 0 Then
            MsgBox "「" & chName & "」を購読しました。" & vbCrLf & _
                   got & " 件の正典を本棚に取り込みました。" & vbCrLf & _
                   "すぐに質問できます。", vbInformation, modAppDef.APP_NAME
        Else
            MsgBox "「" & chName & "」を購読しました。" & vbCrLf & _
                   "(取り込む新しい内容はありませんでした)", vbInformation, modAppDef.APP_NAME
        End If
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
