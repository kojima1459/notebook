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

' R10-5: PaintProgress/ClearProgressが進捗バナーの直前描画シート名を控える
' (モジュールレベル宣言はプロシージャ定義より前に置く。実機VBAの制約)。
Private mProgressSheetName As String

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
    ' 影は固定クロム(ヘッダーバー)と文脈アクションpillにだけ。
    ' 2026-07-26: nx_fab_(旧・常設アクションバー)/nx_sb_(旧サイドバー)は廃止。
    If nm = "nx_top_bg" Or Left$(nm, 7) = "nx_act_" Then
        ApplySoftShadow shp
    End If
    ' 深み(Depth)の演出: 送信ボタンだけ同系グリーンの極微グラデーション
    ' (明るい緑→深い緑)。多用は描画負荷になるため主役の1ボタンに限定する。
    If nm = "nx_top_send" Then ApplyGreenDepth shp
    ' 仕様書§9(Apple風): 3画面のヘッダーバーは濃紺の2色グラデーション。
    ' ApplyThemeがベタ塗りに戻すため、BeautifyAll経由でここが必ず塗り直す。
    If nm = "nx_top_bg" Or nm = "nx_hub_hdr" Or nm = "nxk_hdr" Then
        ApplyHeaderDepth shp
    End If
    On Error GoTo 0
End Sub

' ヘッダーバーの2色グラデーション(#1a365d → #1f4e78)。32bit Excelで
' TwoColorGradientが失敗する環境ではベタ塗りのまま進む(平らに見えるだけ)。
Public Sub ApplyHeaderDepth(ByVal shp As Shape)
    On Error Resume Next
    With shp.Fill
        .TwoColorGradient 1, 1        ' msoGradientHorizontal, variant1
        .ForeColor.RGB = RGB(26, 54, 93)
        .BackColor.RGB = RGB(31, 78, 120)
    End With
    On Error GoTo 0
End Sub

' 汎用の2色グラデーション(ユーザーバブル・EXPゲージ等)。
Public Sub ApplyGradient(ByVal shp As Shape, ByVal c1 As Long, ByVal c2 As Long)
    On Error Resume Next
    With shp.Fill
        .TwoColorGradient 1, 1
        .ForeColor.RGB = c1
        .BackColor.RGB = c2
    End With
    On Error GoTo 0
End Sub

' カード/ボタン用の控えめな浮遊感(§9: Blur=4, OffsetY=1.5, Transparency=0.9)。
' ApplySoftShadowより弱く、要素が多い画面で影が重ならないようにする。
Public Sub ApplyLightShadow(ByVal shp As Shape)
    On Error Resume Next
    With shp.Shadow
        .Visible = -1
        .OffsetX = 0
        .OffsetY = 1.5
        .Transparency = 0.9
        .Size = 100
        .ForeColor.RGB = RGB(15, 23, 42)
        .Blur = 4
    End With
    On Error GoTo 0
End Sub

' MS&ADグリーンの微細な縦グラデーション(フラットの中の上質なマテリアル感)。
' テーマ再適用(ApplyTheme)でベタ塗りに戻ることがあるが、BeautifyAll経由で再適用される。
Public Sub ApplyGreenDepth(ByVal shp As Shape)
    On Error Resume Next
    With shp.Fill
        .TwoColorGradient 1, 1        ' msoGradientHorizontal, variant1(上→下)
        .ForeColor.RGB = RGB(22, 163, 88)    ' わずかに明るい緑
        .BackColor.RGB = RGB(0, 122, 55)     ' 深い緑
    End With
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
' スキン(きせかえ): 称号と同じ「感謝受領数」ゲートで解放されるアンロック方式。
'   light/darkは全員。sakura/oceanは感謝5件、goldは感謝20件で解放。
'   チート防御: 解放判定はResolveColor(全描画の色解決点)でも強制するため、
'   隠しシートのnexus_themeを手書きで"gold"にしても色はlightに落ちる。
'   感謝数はP2P受領のみが源泉(自己付与不可)なので、スキン自体が偽装不可の勲章になる。
' ----------------------------------------------------------------------------

' テーマ名を検証し、未解放/未知ならlightへ落とした正規名を返す。
Public Function EffectiveSkin(ByVal themeName As String) As String
    Dim t As String: t = LCase$(Trim$(themeName))
    Select Case t
        Case "dark":            EffectiveSkin = "dark"
        Case "sakura", "ocean": EffectiveSkin = IIf(ThanksCount() >= 5, t, "light")
        Case "gold":            EffectiveSkin = IIf(ThanksCount() >= 20, t, "light")
        Case Else:              EffectiveSkin = "light"
    End Select
End Function

Private Function ThanksCount() As Long
    On Error Resume Next
    ThanksCount = modStats.GetStat("thanks_received_total")
    On Error GoTo 0
End Function

' 全画面の配色解決点(modUI.ThemeColorから委譲される唯一の実装)。
Public Function ResolveColor(ByVal key As String, ByVal themeName As String) As Long
    Dim t As String: t = EffectiveSkin(themeName)
    Select Case t
        Case "dark"
            Select Case key
                Case "bg":            ResolveColor = RGB(15, 23, 42)
                Case "surface":       ResolveColor = RGB(30, 41, 59)
                Case "text":          ResolveColor = RGB(248, 250, 252)
                Case "muted":         ResolveColor = RGB(148, 163, 184)
                Case "border":        ResolveColor = RGB(51, 65, 85)
                Case "primary":       ResolveColor = RGB(52, 168, 96)
                Case "accent":        ResolveColor = RGB(0, 168, 89)
                Case "userBubble":    ResolveColor = RGB(51, 65, 85)
                Case "aiBubble":      ResolveColor = RGB(30, 41, 59)
                Case "sidebar":       ResolveColor = RGB(11, 15, 25)
                Case "sidebarText":   ResolveColor = RGB(209, 213, 219)
                Case "sidebarActive": ResolveColor = RGB(30, 41, 59)
                Case Else:            ResolveColor = RGB(0, 0, 0)
            End Select
        Case "sakura"   ' 感謝5件で解放: 少し華やかな春色
            Select Case key
                Case "bg":            ResolveColor = RGB(255, 241, 245)
                Case "surface":       ResolveColor = RGB(255, 255, 255)
                Case "text":          ResolveColor = RGB(66, 32, 44)
                Case "muted":         ResolveColor = RGB(164, 120, 136)
                Case "border":        ResolveColor = RGB(248, 214, 224)
                Case "primary":       ResolveColor = RGB(214, 51, 108)
                Case "accent":        ResolveColor = RGB(0, 168, 89)
                Case "userBubble":    ResolveColor = RGB(255, 228, 238)
                Case "aiBubble":      ResolveColor = RGB(255, 255, 255)
                Case "sidebar":       ResolveColor = RGB(84, 32, 52)
                Case "sidebarText":   ResolveColor = RGB(240, 210, 222)
                Case "sidebarActive": ResolveColor = RGB(112, 46, 72)
                Case Else:            ResolveColor = RGB(0, 0, 0)
            End Select
        Case "ocean"    ' 感謝5件で解放: 集中の青
            Select Case key
                Case "bg":            ResolveColor = RGB(240, 247, 255)
                Case "surface":       ResolveColor = RGB(255, 255, 255)
                Case "text":          ResolveColor = RGB(15, 36, 62)
                Case "muted":         ResolveColor = RGB(100, 126, 152)
                Case "border":        ResolveColor = RGB(208, 226, 244)
                Case "primary":       ResolveColor = RGB(2, 102, 190)
                Case "accent":        ResolveColor = RGB(0, 145, 200)
                Case "userBubble":    ResolveColor = RGB(224, 240, 255)
                Case "aiBubble":      ResolveColor = RGB(255, 255, 255)
                Case "sidebar":       ResolveColor = RGB(10, 35, 66)
                Case "sidebarText":   ResolveColor = RGB(198, 219, 240)
                Case "sidebarActive": ResolveColor = RGB(20, 56, 96)
                Case Else:            ResolveColor = RGB(0, 0, 0)
            End Select
        Case "gold"     ' 感謝20件で解放: 黒×金のエグゼクティブ
            Select Case key
                Case "bg":            ResolveColor = RGB(12, 12, 14)
                Case "surface":       ResolveColor = RGB(24, 24, 28)
                Case "text":          ResolveColor = RGB(240, 234, 216)
                Case "muted":         ResolveColor = RGB(162, 150, 120)
                Case "border":        ResolveColor = RGB(64, 58, 42)
                Case "primary":       ResolveColor = RGB(212, 175, 55)
                Case "accent":        ResolveColor = RGB(230, 196, 80)
                Case "userBubble":    ResolveColor = RGB(45, 42, 30)
                Case "aiBubble":      ResolveColor = RGB(24, 24, 28)
                Case "sidebar":       ResolveColor = RGB(6, 6, 8)
                Case "sidebarText":   ResolveColor = RGB(212, 175, 55)
                Case "sidebarActive": ResolveColor = RGB(38, 34, 24)
                Case Else:            ResolveColor = RGB(0, 0, 0)
            End Select
        Case Else       ' light = MS&ADスタンダード(従来値そのまま)
            Select Case key
                Case "bg":            ResolveColor = RGB(243, 244, 246)
                Case "surface":       ResolveColor = RGB(255, 255, 255)
                Case "text":          ResolveColor = RGB(17, 24, 39)
                Case "muted":         ResolveColor = RGB(107, 114, 128)
                Case "border":        ResolveColor = RGB(229, 231, 235)
                Case "primary":       ResolveColor = RGB(0, 137, 62)
                Case "accent":        ResolveColor = RGB(0, 168, 89)
                Case "userBubble":    ResolveColor = RGB(239, 246, 255)
                Case "aiBubble":      ResolveColor = RGB(255, 255, 255)
                Case "sidebar":       ResolveColor = RGB(17, 24, 39)
                Case "sidebarText":   ResolveColor = RGB(209, 213, 219)
                Case "sidebarActive": ResolveColor = RGB(31, 41, 55)
                Case Else:            ResolveColor = RGB(0, 0, 0)
            End Select
    End Select
End Function

' きせかえ切替(ヘルプの🎨ボタンから)。解放済みスキンを巡回し、未解放は
' 「あと◯件で解放」のティザーToastを出してスキップ(欲しくなる導線)。
Public Sub CycleSkin()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    Dim orderList As Variant
    orderList = Array("light", "dark", "sakura", "ocean", "gold")
    Dim labels As Variant
    labels = Array("MS&AD スタンダード", "ダークモード", "サクラ・ピンク", "オーシャン・ブルー", "エグゼクティブ・ゴールド")

    Dim cur As String: cur = EffectiveSkin(modUI.UiTheme())
    Dim curIdx As Long: curIdx = 0
    Dim i As Long
    For i = 0 To 4
        If CStr(orderList(i)) = cur Then curIdx = i
    Next i

    Dim tc As Long: tc = ThanksCount()
    Dim tried As Long
    For tried = 1 To 5
        Dim nx As Long: nx = (curIdx + tried) Mod 5
        Dim cand As String: cand = CStr(orderList(nx))
        Dim needN As Long
        needN = 0
        If cand = "sakura" Or cand = "ocean" Then needN = 5
        If cand = "gold" Then needN = 20
        If tc >= needN Then
            modState.SaveState "nexus_theme", cand
            modUI.Repaint
            modSkin.ShowToast "きせかえ: " & CStr(labels(nx)) & IIf(needN > 0, "(感謝" & needN & "件の限定スキン)", ""), "success"
            GoTo Done
        Else
            modSkin.ShowToast CStr(labels(nx)) & " は「ありがとう」を" & needN & "件受け取ると解放されます(現在" & tc & "件)。", "info"
        End If
    Next tried
Done:
    modUiLock.Leave
End Sub

' ----------------------------------------------------------------------------
' ShowToast - MsgBoxの代替(非ブロッキング通知)。画面上部中央に細長Shapeを出し、
'   短時間表示して自動で消す。kind: "success"/"error"/"info"。
' ----------------------------------------------------------------------------
Public Sub ShowToast(ByVal message As String, Optional ByVal kind As String = "info")
    On Error Resume Next
    ' 別ブック誤爆ガード: ユーザーが他の業務Excelを見ている間にToastを描くと、
    ' 他人のブックへShapeを生成して業務データを汚す。自ブックがアクティブな時だけ描く。
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
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

    ' 種別アイコン(視覚的認知スピード): 成功✅ / 注意⚠️ / 情報💡
    Dim bg As Long, fg As Long, icon As String
    Select Case LCase$(kind)
        Case "success": bg = RGB(0, 168, 89):  fg = RGB(255, 255, 255): icon = ChrW(&H2705)
        Case "error":   bg = RGB(220, 38, 38):  fg = RGB(255, 255, 255): icon = ChrW(&H26A0)
        Case Else:      bg = RGB(30, 41, 59):   fg = RGB(248, 250, 252): icon = ChrW(&HD83D) & ChrW(&HDCA1)
    End Select
    shp.Fill.ForeColor.RGB = bg

    With shp.TextFrame2
        .WordWrap = -1
        .MarginLeft = 16: .MarginRight = 16: .MarginTop = 5: .MarginBottom = 5
        .TextRange.Text = icon & " " & message
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

' ----------------------------------------------------------------------------
' PaintProgress / ClearProgress - 進捗バナー("nx_progress")の表示部(R10-5)。
'   ShowToastと違い待機ゼロ(ToastWaitは呼ばない。ファイル数×1.1秒の純増を
'   避けるのが要件)。更新後にDoEvents1回だけ挟んで再描画させる。
'   別ブック表示中は何もしない(ShowToastと同じ誤爆ガード)。直前と違うシートへ
'   移っていたら旧シートのShapeを消してから今のシートへ描き直す
'   (mProgressSheetNameでシート名を控える)。
' ----------------------------------------------------------------------------
Public Sub PaintProgress(ByVal message As String)
    On Error Resume Next
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    If LenB(mProgressSheetName) > 0 And mProgressSheetName <> ws.Name Then
        Dim wsOld As Worksheet
        Set wsOld = ThisWorkbook.Worksheets(mProgressSheetName)
        If Not wsOld Is Nothing Then wsOld.Shapes("nx_progress").Delete
    End If
    mProgressSheetName = ws.Name

    Dim barW As Double: barW = 380
    Dim leftPos As Double, topPos As Double
    leftPos = ActiveWindow.VisibleRange.Left + (ActiveWindow.VisibleRange.Width - barW) / 2
    ' R10c(M4): トースト(nx_toast)も同じ +92 に出るため、取込完了の瞬間だけ
    ' 2枚が完全に重なり、下になった方の文字が読めなくなっていた。バナーは
    ' トースト(高さ34)の下へずらす。92 + 34 + 余白4 = 130。
    topPos = ActiveWindow.VisibleRange.Top + 130

    Dim shp As Shape
    Set shp = ws.Shapes("nx_progress")
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(5, leftPos, topPos, barW, 30)   ' 5=角丸四角
        shp.Name = "nx_progress"
        shp.Adjustments(1) = 0.3
        shp.Line.Visible = 0
        shp.Placement = 3   ' xlFreeFloating
        shp.Fill.ForeColor.RGB = RGB(30, 41, 59)
        With shp.TextFrame2
            .WordWrap = -1
            .MarginLeft = 16: .MarginRight = 16: .MarginTop = 4: .MarginBottom = 4
            .TextRange.Font.Name = "Yu Gothic UI"
            .TextRange.Font.Size = 10
            .TextRange.ParagraphFormat.Alignment = 2   ' 中央
            .VerticalAnchor = 3
        End With
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(248, 250, 252)
        ApplySoftShadow shp
    Else
        shp.Left = leftPos
        shp.Top = topPos
    End If
    shp.TextFrame2.TextRange.Text = message
    shp.ZOrder 0   ' msoBringToFront
    DoEvents
    On Error GoTo 0
End Sub

Public Sub ClearProgress()
    On Error Resume Next
    If LenB(mProgressSheetName) > 0 Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(mProgressSheetName)
        If Not ws Is Nothing Then ws.Shapes("nx_progress").Delete
    End If
    mProgressSheetName = ""
    On Error GoTo 0
End Sub
