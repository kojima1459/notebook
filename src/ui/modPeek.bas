Attribute VB_Name = "modPeek"
Option Explicit

' ============================================================================
' modPeek - ワンクリック出典ポップアップ(Peek View)。UI層。
' ----------------------------------------------------------------------------
' 役割:
'   AI回答の直下に「出典チップ」(📄 資料名 p.N)を並べ、クリックすると元の
'   チャンク本文が小さな吹き出し(ツールチップ)としてフワッと浮かぶ。PDFを開かず
'   審査員が「AIの回答が本当に約款に書いてあるか」を秒で照合でき、ハルシネーションを
'   人間が即検知・是正できる、というデモ最大のアピール機能。
'
' 設計判断:
'   ・出典データは modAsk の読み取り専用アクセサ(LastHit*)から取得する(検索/回答
'     ロジックには一切触れない)。UI層→qa層の下向き参照でR1レイヤリング準拠。
'   ・modUI.basは字数上限間際のため一切変更しない。描画は本モジュールに集約。
'   ・チップは「最新の回答」の下にだけ出す(次の送信でHideCitationsして描き直す)。
'     modUIのmChatBottom(private)を触らずに済ませるための割り切り。過去回答の
'     チップは残さない(Peekは今の回答の照合が目的)。
'   ・チップ/ポップアップはShape。クリック時はOnActionでmodApp.OnPeek/OnPeekCloseへ。
'     Shape名 nx_cite_<hitIndex> に0始まりのヒット添字を埋め、Peek本文を引く。
'   ・ポップアップは nx_peek。次のPeek表示・送信・閉じるクリックで消える(孤児防止に
'     都度Deleteしてから描く)。柔らかい影は modSkin.ApplySoftShadow を共用。
'   ・LibreOfficeは静的コンパイルのみ(実行しない)。Shapeプロパティは標準VBA。
' ============================================================================

Private Const NEXUS_SHEET As String = "Nexus"
Private Const MAX_CHIPS As Long = 4          ' 出典チップの最大数(横並び)
Private Const CHIP_W As Double = 188
Private Const CHIP_H As Double = 22
Private Const PEEK_W As Double = 470
Private Const PEEK_BODY_MAX As Long = 600    ' ポップアップ本文の最大文字数

' ----------------------------------------------------------------------------
' RenderCitations - 直近RAG回答の出典チップを、指定バブルの直下に描画する。
'   bubbleName: 直前に追加したAIバブルのShape名(位置決めの基準)。
' ----------------------------------------------------------------------------
Public Sub RenderCitations(ByVal bubbleName As String)
    On Error GoTo Done
    Dim ws As Worksheet
    Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub

    HideCitations   ' 前回のチップ/ポップアップを消す(最新回答の下だけに出す)

    Dim n As Long: n = modAsk.LastHitCount()
    If n <= 0 Then Exit Sub

    Dim anchor As Shape
    On Error Resume Next
    Set anchor = ws.Shapes(bubbleName)
    On Error GoTo Done
    If anchor Is Nothing Then Exit Sub

    Dim baseL As Double: baseL = anchor.Left
    Dim baseY As Double: baseY = anchor.Top + anchor.Height + 6

    ' 2026-07-27: 積む順を「回答 → 信頼度 → 出典 → 評価」に変えた。
    ' 以前は評価ボタン(nx_act_)の下に出典を置いていたため、利用者は
    ' 542pt分のボタンを越えないと根拠にたどり着けなかった。つまり
    ' 「確かめる前に評価しろ」という並びになっていた。順序は主張なので、
    ' 根拠を先に、上に出す。評価ボタン側がこちらの下端を見て下がる。
    On Error Resume Next
    Dim confBottom As Double
    confBottom = modUINexusDraw.ConfidenceBottom(ws)
    If confBottom + 6 > baseY Then baseY = confBottom + 6
    On Error GoTo Done

    ' 見出しラベル
    Dim lbl As Shape
    Set lbl = ws.Shapes.AddShape(1, baseL, baseY, 260, 16)
    lbl.Name = "nx_cite_lbl"
    lbl.Fill.Visible = 0: lbl.Line.Visible = 0
    With lbl.TextFrame2
        .WordWrap = -1
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDD0E) & " 出典(クリックで原文を確認):"
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 8.5
        .MarginLeft = 2: .MarginTop = 0: .MarginBottom = 0
    End With
    lbl.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
    lbl.Placement = 3

    ' 出典をソース名でユニーク化(先頭出現のヒット添字を保持)しつつチップ描画
    Dim seen As String: seen = "|"
    Dim x As Double: x = baseL
    Dim y As Double: y = baseY + 18
    Dim drawn As Long: drawn = 0
    Dim i As Long
    For i = 0 To n - 1
        Dim src As String: src = modAsk.LastHitSource(i)
        If LenB(src) = 0 Then GoTo NextHit
        Dim key As String: key = "|" & LCase$(src) & "|"
        If InStr(seen, key) > 0 Then GoTo NextHit   ' 同じ資料は1チップに集約
        seen = seen & LCase$(src) & "|"

        If x + CHIP_W > baseL + 640 Then   ' チャット幅で折り返し
            x = baseL
            y = y + CHIP_H + 6
        End If
        DrawChip ws, i, x, y, src, modAsk.LastHitPage(i)
        x = x + CHIP_W + 8
        drawn = drawn + 1
        If drawn >= MAX_CHIPS Then Exit For
NextHit:
    Next i

    FreezeAndFront ws
Done:
End Sub

Private Sub DrawChip(ByVal ws As Worksheet, ByVal hitIdx As Long, ByVal x As Double, _
                     ByVal y As Double, ByVal src As String, ByVal page As Long)
    Dim chip As Shape
    Set chip = ws.Shapes.AddShape(5, x, y, CHIP_W, CHIP_H)   ' 5=角丸四角
    chip.Name = "nx_cite_" & CStr(hitIdx)
    chip.Adjustments(1) = 0.5
    chip.Line.Visible = -1
    chip.Line.Weight = 0.75
    chip.Line.ForeColor.RGB = modUI.UiColor("primary")
    chip.Fill.ForeColor.RGB = modUI.UiColor("surface")

    Dim cap As String
    cap = ChrW(&HD83D) & ChrW(&HDCC4) & " " & modUtil.SafeLeft(src, 16)
    If page > 0 Then cap = cap & " p." & page
    With chip.TextFrame2
        .WordWrap = -1
        .TextRange.Text = cap
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 8.5
        .TextRange.ParagraphFormat.Alignment = 1   ' 左
        .VerticalAnchor = 3
        .MarginLeft = 8: .MarginRight = 6: .MarginTop = 0: .MarginBottom = 0
    End With
    chip.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("primary")
    chip.OnAction = "modApp.OnPeek"
    chip.Placement = 3
End Sub

' ----------------------------------------------------------------------------
' ShowPeek - 出典チップのクリックで、そのチャンク本文をポップアップ表示する。
'   idx: modAskの0始まりヒット添字(チップ名 nx_cite_<idx> 由来)。
' ----------------------------------------------------------------------------
Public Sub ShowPeek(ByVal idx As Long)
    On Error GoTo Done
    Dim ws As Worksheet
    Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub

    HidePeek

    Dim body As String: body = modAsk.LastHitPeek(idx)
    If LenB(body) = 0 Then body = "(この出典の本文プレビューは取得できませんでした)"
    Dim src As String: src = modAsk.LastHitSource(idx)
    Dim page As Long: page = modAsk.LastHitPage(idx)

    Dim leftPos As Double: leftPos = 280
    Dim topPos As Double: topPos = 110
    On Error Resume Next
    leftPos = ActiveWindow.VisibleRange.Left + (ActiveWindow.VisibleRange.Width - PEEK_W) / 2
    topPos = ActiveWindow.VisibleRange.Top + 96
    On Error GoTo Done

    Dim head As String
    head = ChrW(&HD83D) & ChrW(&HDCC4) & " " & src
    If page > 0 Then head = head & "  (p." & page & ")"

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, leftPos, topPos, PEEK_W, 60)   ' 高さはAutoSizeで伸ばす
    shp.Name = "nx_peek"
    shp.Adjustments(1) = 0.06
    shp.Line.Visible = -1
    shp.Line.Weight = 1#
    shp.Line.ForeColor.RGB = modUI.UiColor("primary")
    shp.Fill.ForeColor.RGB = modUI.UiColor("surface")
    With shp.TextFrame2
        .WordWrap = -1
        .AutoSize = 1   ' msoAutoSizeShapeToFitText
        .MarginLeft = 14: .MarginRight = 14: .MarginTop = 10: .MarginBottom = 10
        .TextRange.Text = head & vbLf & vbLf & _
                          modUtil.SafeLeft(body, PEEK_BODY_MAX) & vbLf & vbLf & _
                          ChrW(&H2715) & " クリックで閉じる"
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 10
        .TextRange.ParagraphFormat.Alignment = 1
    End With
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
    shp.OnAction = "modApp.OnPeekClose"
    shp.Placement = 3
    modSkin.ApplySoftShadow shp
    shp.ZOrder 0   ' msoBringToFront

    ' R27 F2-1b(実機第12報①): プレビューは modUI.RecalcChatBottom の許可
    ' リスト外(重ね表示)なので境界関所を通らない。AutoSizeで伸びた実下端で
    ' 関所を通す。「原文を開く」ボタンはこの吹き出しの内側に置くので、
    ' ここが常に最深(GoTo Doneで飛んだ経路でも取りこぼさない)。
    ' R27H F2(M-1裁定): 開いている間の下端を床として登録する(開いたまま
    ' 質問送信/トグルが走ると、別経路が会話の下端で境界を貼り直すため)。
    modSkin.SetOverlayFloor shp.Top + shp.Height + 12
    modSkin.ExtendChatBand ws, shp.Top + shp.Height + 12

    ' 「📂 原文を開く」。抜粋を読んで終わりではなく、実物の該当ページまで
    ' 連れて行く。ここまで来て初めて、人に見せられる根拠になる。
    ' 元ファイルの記録が無い資料(パック由来・手入力)には出さない。
    If LenB(SourcePath(src)) = 0 Then GoTo Done

    Dim openCap As String
    openCap = ChrW(&HD83D) & ChrW(&HDCC2) & " 原文を開く"
    If page > 0 Then openCap = openCap & "（p." & page & "）"

    Dim btn As Shape
    On Error Resume Next
    Set btn = ws.Shapes.AddShape(5, shp.Left + PEEK_W - 158, shp.Top + shp.Height - 34, 142, 26)
    If Err.Number = 0 And Not btn Is Nothing Then
        btn.Name = "nx_peek_open_" & CStr(idx)
        btn.Adjustments(1) = 0.4
        btn.Line.Visible = 0
        btn.Fill.ForeColor.RGB = modUI.UiColor("primary")
        With btn.TextFrame2
            .WordWrap = -1
            .TextRange.Text = openCap
            .TextRange.Font.Name = "Yu Gothic UI"
            .TextRange.Font.Size = 9
            .TextRange.Font.Bold = -1
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = 2
            .VerticalAnchor = 3
            .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
        End With
        btn.OnAction = "modPeek.OnOpenSource"
        btn.Placement = 3
        btn.ZOrder 0
    End If
    Set btn = Nothing
    Err.Clear
    On Error GoTo Done
Done:
End Sub

Public Sub HidePeek()
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub
    ws.Shapes("nx_peek").Delete
    ' 「原文を開く」ボタンはポップアップとは別Shapeなので、道連れにしないと
    ' 本文だけ消えてボタンだけが宙に浮いて残る。
    Dim names() As String
    ReDim names(0 To ws.Shapes.count)
    Dim n As Long: n = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 13) = "nx_peek_open_" Then
            names(n) = shp.Name
            n = n + 1
        End If
    Next shp
    Dim i As Long
    For i = 0 To n - 1
        ws.Shapes(names(i)).Delete
    Next i
    ' R27 F2-1b: 伸ばした境界を会話の下端へ戻す(戻さないと閉じたあとも
    ' 境界だけが下に残り、白い余白として見える)。
    ' R27H F2: 床を先に下ろす(残っていると縮む方向が効かない)。
    modSkin.ClearOverlayFloor
    modSkin.ExtendChatBand ws, modUI.ChatBottomFor(ws)
    On Error GoTo 0
End Sub

' 出典チップとポップアップを一括削除(次の送信の先頭・画面リセット時に呼ぶ)。
Public Sub HideCitations()
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub
    Dim names() As String
    ReDim names(0 To ws.Shapes.count)
    Dim n As Long: n = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 8) = "nx_cite_" Or shp.Name = "nx_peek" _
           Or Left$(shp.Name, 13) = "nx_peek_open_" Then
            names(n) = shp.Name
            n = n + 1
        End If
    Next shp
    Dim i As Long
    For i = 0 To n - 1
        ws.Shapes(names(i)).Delete
    Next i
    On Error GoTo 0
End Sub

' 出典チップ群の下端(無ければ0)。評価ボタンがこの下へ回り込むために使う。
Public Function CitationsBottom(ByVal ws As Worksheet) As Double
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 8) = "nx_cite_" Then
            If shp.Top + shp.Height > CitationsBottom Then CitationsBottom = shp.Top + shp.Height
        End If
    Next shp
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' OnOpenSource - ポップアップの「📂 原文を開く」。
' ----------------------------------------------------------------------------
' ここが無いせいで、このアプリは「retrievalのデモ」で止まっていた。
' 損保の実務で意味を持つ瞬間は、上司やお客さまに約款の該当ページそのものを
' 見せるとき。灰色のテキストボックスで終わっていては、その場に持っていけない。
' file_pathは my_manifest に既にある。開くだけでよかった。
Public Sub OnOpenSource()
    ' R15波1b裁定: 外部アプリで原文を開く際、取込中は同じファイルを掴んでいる
    ' 可能性がある(既存40箇所超と同型。lint検査13で現状追認だったものを
    ' 保護へ切り替え)。
    If modUiLock.BlockIfIngesting() Then Exit Sub
    On Error Resume Next
    Dim caller As String
    caller = CStr(Application.Caller)
    On Error GoTo 0
    If Left$(caller, 13) <> "nx_peek_open_" Then
        modLog.LogUsage "caller_mismatch", "modPeek.OnOpenSource", caller
        Exit Sub
    End If

    Dim idx As Long
    idx = CLng(Val(Mid$(caller, 14)))

    Dim src As String, page As Long
    On Error Resume Next
    src = modAsk.LastHitSource(idx)
    page = modAsk.LastHitPage(idx)
    On Error GoTo 0
    If LenB(src) = 0 Then Exit Sub

    Dim path As String
    path = SourcePath(src)

    If LenB(path) = 0 Then
        MsgBox "この資料の元ファイルの場所が記録されていません。" & vbCrLf & _
               "(パックで受け取った資料や、手入力で登録した内容には元ファイルがありません)", _
               vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    ' PDFはページ指定で開く。対応しないビューアでは先頭ページで開くだけで、
    ' 失敗はしない。開けなかったときはパスを見せる(手で辿れるようにする)。
    Dim target As String: target = path
    If page > 0 Then
        If LCase$(modUtil.ExtOf(path)) = "pdf" Then target = path & "#page=" & page
    End If

    Dim ok As Boolean: ok = True
    On Error Resume Next
    ' 2026-08-10(R27波3-14): waitless:=True。この一言は【この直後に出る
    ' Officeの確認画面】の予告であって、読ませるために待たせる文ではない。
    ' 既定のままだと 1.1 秒ぶん確認画面の到着が遅れ、その待ちの DoEvents 中に
    ' 別のクリックが割り込む窓も開く。描いたらすぐ開きにいく。
    modSkin.ShowToast "Officeの確認画面が出たら[はい]を押してください。", "info", True
    ThisWorkbook.FollowHyperlink target
    If Err.Number <> 0 Then ok = False
    Err.Clear
    On Error GoTo 0

    If Not ok Then
        On Error Resume Next
        modClip.SetClipboardText path
        On Error GoTo 0
        MsgBox "元ファイルを開けませんでした(移動または削除された可能性があります)。" & vbCrLf & _
               "場所をクリップボードにコピーしました:" & vbCrLf & path, _
               vbExclamation, modAppDef.APP_NAME
    End If
End Sub

' ----------------------------------------------------------------------------
' SourcePath - 資料名(= my_manifest の file_name)から元ファイルの実パスを引く。
'   見つからない/実体が消えている場合は空文字を返し、呼び出し側はボタンを出さない。
' ----------------------------------------------------------------------------
' パスは取込時から manifest 1列目に入っていたのに、そこへ到達する道が
' どこにも無かった。パック由来・手入力のナレッジは manifest に行が無いので空。
Private Function SourcePath(ByVal sourceName As String) As String
    Dim nm As String: nm = Trim$(sourceName)
    If LenB(nm) = 0 Then Exit Function

    On Error Resume Next
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Worksheets(modAppDef.SH_MANIFEST)
    On Error GoTo 0
    If wsM Is Nothing Then Exit Function

    Dim lastR As Long
    On Error Resume Next
    lastR = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
    On Error GoTo 0
    If lastR < 2 Then Exit Function

    Dim r As Long
    On Error Resume Next
    For r = 2 To lastR
        If StrComp(Trim$(CStr(wsM.Cells(r, 2).Value)), nm, vbTextCompare) = 0 Then
            Dim p As String: p = Trim$(CStr(wsM.Cells(r, 1).Value))
            ' 消えたファイルのパスを返すと「開けません」で終わる。存在確認まで
            ' 済ませてから返し、無ければボタン自体を出さない。
            If LenB(p) > 0 Then
                If LenB(Dir(p)) > 0 Then SourcePath = p
            End If
            Exit For
        End If
    Next r
    On Error GoTo 0
End Function

Private Function GetSheet() As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(NEXUS_SHEET)
    On Error GoTo 0
End Function

Private Sub FreezeAndFront(ByVal ws As Worksheet)
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 8) = "nx_cite_" Then shp.Placement = 3
    Next shp
    On Error GoTo 0
End Sub
