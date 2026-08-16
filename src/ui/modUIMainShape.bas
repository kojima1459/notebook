Attribute VB_Name = "modUIMainShape"
Option Explicit

' ========================================
' modUIMainShape - 旧ホーム画面のシート取得・図形生成・図形一括削除
'
' modUIMain から切り出した描画の下請け。座標のクランプ(SafeCoord)や
' 角丸生成の退避(SafeRoundedRect)といった「実機で落ちないための細工」が
' ここに集まるので、レイアウトの意図(modUIMain)と混ざらない。
'
' 切り出しの理由(2026-07-28): modUIMain が契約上限30,000字に対し残り545字
' しかなかった(レビュー I-2)。
' ========================================

' 旧ホーム画面のボタン配色。
' 2026-07-28: modUIMain から AddButton をここへ移した際、この2つを向こうに
' 置いたままにしてしまい、実機で「変数が定義されていません」というコンパイル
' エラーになった。複製せず、図形を描く側であるここを単一情報源にする
' (modUIMain.PaintModeButton は modUIMainShape.COLOR_... を参照する)。
Public Const COLOR_UNSELECTED_BG As Long = 15921906 ' RGB(242,242,242) 薄灰
Public Const COLOR_UNSELECTED_FG As Long = 0        ' RGB(0,0,0) 黒

Public Sub WriteSafe(ByVal cell As Range, ByVal text As String)
    Dim t As String
    t = modUtil.SafeLeft(text, 32000)
    On Error Resume Next
    ' 未Mergeのまま代入すると全セルに同じ値が複製される。都度Mergeする。
    ' R25H F1: cell.Mergeが出す「複数の値を持つセル範囲への結合」警告は
    ' On Error配下でも止められず、OK待ちでRAG回答描画(RenderSourcesPreview/
    ' RenderAnswer/ShowTip→ここ)を無言のまま固める(実機③④フリーズの本丸)。
    ' AlertsOff/AlertsOnで挟み、警告そのものを出さない(単一チョークポイント)。
    ' 直後の行は既にOn Error Resume Next配下なので、Mergeが例外を出しても
    ' 次の行(AlertsOn)へ進む=どちらの経路でも必ず対になる。
    If cell.Cells.Count > 1 Then
        modUiLock.AlertsOff
        cell.Merge
        modUiLock.AlertsOn
    End If
    cell.Value = t
    On Error GoTo 0
End Sub

Public Function GetHomeSheet() As Worksheet
    On Error Resume Next
    Set GetHomeSheet = ThisWorkbook.Worksheets(modAppDef.SH_HOME)
    On Error GoTo 0
End Function

Public Function GetOrCreateHomeSheet() As Worksheet
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
Public Sub AddButton(ByVal ws As Worksheet, ByVal rng As Range, ByVal shapeName As String, _
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
    ' Err.Raise伝播に依存せず直接err_logへ書く(この後のLogError呼び出しで
    ' Err自体が上書きされ得るため先に退避)。
    Dim btnErrNum As Long, btnErrDesc As String
    btnErrNum = Err.Number: btnErrDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup3
FailCleanup3:
    Dim diag As String: diag = ""
    On Error Resume Next
    diag = " ws.Visible=" & ws.Visible & " ws.ProtectContents=" & ws.ProtectContents & _
           " ActiveSheet=" & ThisWorkbook.ActiveSheet.Name & _
           " Interactive=" & Application.Interactive & _
           " AppWin=" & Application.Windows.Count & " WbWin=" & ThisWorkbook.Windows.Count
    modLog.LogError "E0801", "modUIMain.EnsureLayout", "AddButton [" & uiStep & "]" & diag, btnErrNum
    On Error GoTo 0
    ' 2026-07-22実機再発: 従来はErr.Raiseで再伝播していたが、それだと1個の
    ' ボタン生成失敗がEnsureLayout全体を中断させ、以降のボタン/レイアウトが
    ' 丸ごと描画されない(実機報告「画面がほぼ空っぽ」の直接原因)。ログは残す
    ' が再伝播はやめ、この1個だけ諦めて残りの描画を続けさせる。
    Err.Clear
End Sub

Public Function SafeCoord(ByVal v As Double) As Double
    If v < 1 Then v = 1
    SafeCoord = v
End Function

' 直接呼ばれても壊れないよう二重にクランプする。
Public Function SafeRoundedRect(ByVal ws As Worksheet, ByVal L As Double, ByVal T As Double, _
                                 ByVal W As Double, ByVal H As Double) As Shape
    L = SafeCoord(L): T = SafeCoord(T): W = SafeCoord(W): H = SafeCoord(H)
    On Error GoTo Retry
    Set SafeRoundedRect = ws.Shapes.AddShape(5, L, T, W, H)   ' 5=msoShapeRoundedRectangle(リテラル)
    Exit Function
Retry:
    DoEvents
    Set SafeRoundedRect = ws.Shapes.AddShape(5, L, T, W, H)
End Function

' HubIconLabel - Hub(SH_HOME)ヘッダーのアイコン下ラベルを1枚置く。R33 W5-16。
'   実体をここへ置く理由: modHub は残206字で分岐を書き足せない(R33容量裁定)。
'   幾何: HDR_H=48 のとき、アイコン(26pt円)は y∈[rowTop+11, rowTop+37]、この
'   ラベルは y∈[rowTop+36, rowTop+48]。重なりは1ptだけで、ラベルは実質
'   アイコンの【外】にある ―― つまり押してもアイコンには当たらない。
'   従来はここに OnAction が無く、しかも塗り無し(Fill.Visible=0)だったため、
'     ・文字グリフの上を押す → ラベルShapeが選択されるだけで何も起きない
'     ・文字の外(帯の余白)を押す → 下のセルへ透過して、やはり何も起きない
'   の2通りしかなかった。SH_HOME は DrawingObjects 保護が無い(EnsureHubLayout
'   が Unprotect 後に Protect し直さない)ので、白い選択ハンドルだけが残る。
'   アイコンと同じ行き先を与え、帯全体を当たり判定にする
'   (Fill.Visible=-1 + Transparency=1 は modHubStat の nx_hub_footer と同型)。
Public Sub HubIconLabel(ByVal ws As Worksheet, ByVal shapeName As String, _
                        ByVal x As Double, ByVal y As Double, ByVal w As Double, _
                        ByVal labelText As String, ByVal action As String)
    On Error Resume Next
    Dim cap As Shape
    Set cap = ws.Shapes.AddShape(1, x, y, w, 12)
    If Not cap Is Nothing Then
        cap.Name = shapeName
        cap.Line.Visible = 0
        cap.Fill.Visible = -1
        cap.Fill.Transparency = 1
        With cap.TextFrame2
            .WordWrap = 0
            .TextRange.Text = labelText
            .TextRange.Font.Size = 8.5      ' R12-7-3(a11y): 6ptは判読不能
            .TextRange.Font.Fill.ForeColor.RGB = RGB(190, 210, 235)
            .TextRange.ParagraphFormat.Alignment = 2
            .VerticalAnchor = 3
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        End With
        If LenB(action) > 0 Then cap.OnAction = action
    End If
    Set cap = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub RemoveManagedShapes(ByVal ws As Worksheet)
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
