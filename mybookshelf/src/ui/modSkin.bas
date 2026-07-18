Attribute VB_Name = "modSkin"
Option Explicit

' ============================================================================
' modSkin - デザインシステムの単一情報源(脱Excel/Windows95感のポリッシュ)。
' ----------------------------------------------------------------------------
' 役割:
'   ・BeautifyAll / StyleShape: 全nx_Shapeへ「Yu Gothic UI」フォントを徹底し、
'     固定クロム(トップバー/アクションボタン/サイドバー地)にだけ柔らかい影を付与。
'     チャットバブルはフラット(影は選択時のみ)=描画負荷のメリハリ(FPS低下防止)。
'   ・ShowToast: MsgBoxの代替(ハイブリッド運用)。完了/情報などの非ブロッキング通知を
'     画面上部にスッと出して自動で消すToast(細長Shape)で表示する。確認/入力/起動/
'     エラー/別シート時は呼び出し側でMsgBoxを維持する(モーダル性の担保)。
'
' 設計判断:
'   ・modUI.bas は文字数上限(30000字)に近いため、ポリッシュのロジックはこの新モジュールに
'     集約し、modUI側は BeautifyAll/StyleBubble/ApplySoftShadow の「呼び出し」だけを足す
'     (regression最小化)。配色はmodUI.UiColorを唯一の窓口として参照し重複させない。
'   ・影プロパティ(.Blur/.Transparency等)は環境・バージョン差で未対応があり得るため、
'     すべてOn Error Resume Next配下(見た目の劣化はあってもクラッシュさせない)。
'   ・LibreOfficeは静的コンパイルのみ(実行しない)。Shape/Windowプロパティは標準VBAの
'     ためコンパイル可能。
' ============================================================================

' ----------------------------------------------------------------------------
' BeautifyAll - シート上の全nx_Shapeにフォント統一+固定クロムへ柔らかい影。
'   modUI.InitUI/Repaint、および各画面(Vault/Dashboard)の描画終端から呼ぶ。
' ----------------------------------------------------------------------------
Public Sub BeautifyAll(ByVal ws As Worksheet)
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        Dim nm As String: nm = shp.Name
        If Left$(nm, 3) = "nx_" Then StyleShape shp, nm
    Next shp
    On Error GoTo 0
End Sub

' 1つのShapeへ: フォント統一(脱MS Pゴシック)+固定クロムにだけ影。
' バブル(nx_msg_)や思考(nx_thk_)の影はここでは触らない(フラット維持/選択時のみ浮遊)。
Public Sub StyleShape(ByVal shp As Shape, ByVal nm As String)
    On Error Resume Next
    shp.TextFrame2.TextRange.Font.Name = "Yu Gothic UI"
    If nm = "nx_top_bg" Or Left$(nm, 7) = "nx_fab_" Or nm = "nx_sb_bg" Then
        ApplySoftShadow shp
    End If
    On Error GoTo 0
End Sub

' StyleBubble - 新規バブル生成直後にフォントだけ適用(BeautifyAllの軽量版)。
Public Sub StyleBubble(ByVal shp As Shape)
    On Error Resume Next
    shp.TextFrame2.TextRange.Font.Name = "Yu Gothic UI"
    On Error GoTo 0
End Sub

' ふんわり柔らかいドロップシャドウ(透明度高め・ぼかし広め・ベタ塗りの黒影を避ける)。
' 固定バー/カード/選択中バブルにだけ使う(全面に付けると描画負荷でFPSが落ちるため)。
Public Sub ApplySoftShadow(ByVal shp As Shape)
    On Error Resume Next
    With shp.Shadow
        .Visible = -1        ' msoTrue
        .Blur = 8
        .OffsetX = 0
        .OffsetY = 2
        .Transparency = 0.86 ' 86%=ふんわり
        .Size = 100
        .ForeColor.RGB = RGB(15, 23, 42)
    End With
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' ShowToast - MsgBoxの代替(非ブロッキング通知)。画面上部中央に細長Shapeを出し、
'   短時間表示して自動で消す。kind: "success"/"error"/"info"。
' ----------------------------------------------------------------------------
Public Sub ShowToast(ByVal message As String, Optional ByVal kind As String = "info")
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    ws.Shapes("nx_toast").Delete   ' 前のToastが残っていれば消す(孤児防止)

    Dim toastW As Double: toastW = 380
    Dim leftPos As Double: leftPos = 280
    Dim topPos As Double: topPos = 96
    leftPos = ActiveWindow.VisibleRange.Left + (ActiveWindow.VisibleRange.Width - toastW) / 2
    topPos = ActiveWindow.VisibleRange.Top + 92

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, leftPos, topPos, toastW, 34)   ' 5=角丸四角
    shp.Name = "nx_toast"
    shp.Adjustments(1) = 0.35
    shp.Line.Visible = 0
    shp.Placement = 3   ' xlFreeFloating

    Dim bg As Long, fg As Long
    Select Case LCase$(kind)
        Case "success": bg = RGB(0, 168, 89):  fg = RGB(255, 255, 255)
        Case "error":   bg = RGB(220, 38, 38):  fg = RGB(255, 255, 255)
        Case Else:      bg = RGB(30, 41, 59):   fg = RGB(248, 250, 252)
    End Select
    shp.Fill.ForeColor.RGB = bg

    With shp.TextFrame2
        .WordWrap = -1
        .MarginLeft = 14: .MarginRight = 14: .MarginTop = 4: .MarginBottom = 4
        .TextRange.Text = message
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 10.5
        .TextRange.ParagraphFormat.Alignment = 2   ' 中央
        .VerticalAnchor = 3
    End With
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = fg

    ApplySoftShadow shp
    shp.ZOrder 0   ' msoBringToFront
    DoEvents
    ToastWait 1100
    ws.Shapes("nx_toast").Delete
    On Error GoTo 0
End Sub

' Timer基準の短時間待機(DoEventsで応答性維持。Sleep API宣言を避けbitness非依存)。
Private Sub ToastWait(ByVal ms As Long)
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000# < ms
        DoEvents
        If Timer < t0 Then Exit Do   ' 深夜0時のTimerロールオーバーガード
    Loop
End Sub
