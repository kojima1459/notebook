Attribute VB_Name = "modSkin"
Option Explicit

' ============================================================================
' modSkin - デザインシステムの単一情報源(脱Excel/Windows95感のポリッシュ)。
' ----------------------------------------------------------------------------
' 役割:
'   ・BeautifyAll/StyleShape: 全nx_Shapeへ「Yu Gothic UI」を徹底し、固定クロム
'     (トップバー/アクションボタン/サイドバー地)にだけ柔らかい影を付与。
'     チャットバブルはフラット(影は選択時のみ、描画負荷にメリハリ)。
'   ・ShowToast: MsgBoxの代替(ハイブリッド)。完了/情報の非ブロッキング通知を
'     画面上部にToast(細長Shape)で出し自動で消す。確認/入力/起動/エラー/
'     別シート時は呼び出し側でMsgBoxを維持。
'
' 設計判断:
'   ・modUI.basが上限に近いためポリッシュはここへ集約、modUI側は呼び出しのみ。
'     配色はmodUI.UiColorを唯一の窓口とする。
'   ・影プロパティ(.Blur等)は環境差で未対応があり得るため全て
'     On Error Resume Next配下(劣化はあってもクラッシュさせない)。
'   ・LibreOfficeは静的コンパイルのみ(Shape/Windowプロパティは標準VBA)。
' ============================================================================

' 2026-08-05(R18-1a): 進捗バナー一式(PaintProgress/PaintCancelButton/
' ClearProgress/mProgressSheetName/Shape名2定数)は src/ui/modProgressBar.bas へ
' 移設した(本モジュールが上限まで残り10字で1行も直せなかったため。憲章§4-6)。

Private Const THEME_KEY As String = "nexus_theme"

' R19-1b: 直近にチャットの塗りを当てた最終行。会話が伸びたぶんだけ塗り足す
' ための目印で、同じ範囲を毎バブル塗り直さないためだけに持つ(状態の本体は
' modUI.mChatBottom。ここはその写像なので、ズレても塗りが1回増えるだけ)。
Private mChatBandRow As Long

' ----------------------------------------------------------------------------
' BeautifyAll - シート上の全nx_Shapeにフォント統一+固定クロムへ柔らかい影。
'   modUI.InitUI/Repaint、各画面(Vault/Dashboard)の描画終端から呼ぶ。
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

' 1つのShapeへ: フォント統一(脱MS Pゴシック)+固定クロムにだけ影。バブル
' (nx_msg_)や思考(nx_thk_)の影は触らない(フラット維持/選択時のみ)。
Public Sub StyleShape(ByVal shp As Shape, ByVal nm As String)
    On Error Resume Next
    shp.TextFrame2.TextRange.Font.Name = "Yu Gothic UI"
    ' 影は固定クロム(ヘッダーバー)と文脈アクションpillにだけ。
    ' 2026-07-26: nx_fab_(旧・常設アクションバー)/nx_sb_(旧サイド)は廃止。
    If nm = "nx_top_bg" Or Left$(nm, 7) = "nx_act_" Then
        ApplySoftShadow shp
    End If
    ' 深み(Depth)の演出: 送信ボタンだけ同系グリーンの極微グラデーション
    ' (明→深)。多用は描画負荷になるため主役の1ボタンに限定する。
    If nm = "nx_top_send" Then ApplyGreenDepth shp
    ' 仕様書§9(Apple風): 3画面のヘッダーバーは濃紺の2色グラデーション。
    ' ApplyThemeがベタ塗りに戻すため、BeautifyAll経由でここが塗り直す。
    If nm = "nx_top_bg" Or nm = "nx_hub_hdr" Or nm = "nxk_hdr" Then
        ApplyHeaderDepth shp
    End If
    On Error GoTo 0
End Sub

' ヘッダーバーの2色グラデーション(#1a365d → #1f4e78)。32bit Excelで
' TwoColorGradientが失敗する環境はベタ塗りのまま進む(平らに見えるだけ)。
Public Sub ApplyHeaderDepth(ByVal shp As Shape)
    On Error Resume Next
    With shp.Fill
        .TwoColorGradient 1, 1        ' msoGradientHorizontal, variant1
        .ForeColor.RGB = RGB(26, 54, 93)
        .BackColor.RGB = RGB(31, 78, 120)
    End With
    On Error GoTo 0
End Sub

' 汎用の2色グラデーション(バブル・EXPゲージ等)。
Public Sub ApplyGradient(ByVal shp As Shape, ByVal c1 As Long, ByVal c2 As Long)
    On Error Resume Next
    With shp.Fill
        .TwoColorGradient 1, 1
        .ForeColor.RGB = c1
        .BackColor.RGB = c2
    End With
    On Error GoTo 0
End Sub

' 2026-08-01(R12-7-2): RGBを各チャンネル一律 pct 暗くする(0〜1)。自分の発言
' バブルのグラデーション終端を userBubble の同系微差色にする内部部品
' (a11y監査Med: 終端をprimaryにすると下端の文字が沈む)。
Private Function DarkenRgb(ByVal rgbVal As Long, ByVal pct As Double) As Long
    Dim r As Long, g As Long, b As Long
    r = rgbVal Mod 256
    g = (rgbVal \ 256) Mod 256
    b = (rgbVal \ 65536) Mod 256
    DarkenRgb = RGB(CLng(r * (1 - pct)), CLng(g * (1 - pct)), CLng(b * (1 - pct)))
End Function

' カード/ボタン用の控えめな浮遊感(§9: Blur=4, OffsetY=1.5, Transparency=0.9)。
' ApplySoftShadowより弱く、要素が多い画面で影が重ならない。
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

' MS&ADグリーンの微細な縦グラデーション(フラットの中の上質な質感)。
' ApplyTheme でベタ塗りに戻ることがあるが BeautifyAll が再適用する。
Public Sub ApplyGreenDepth(ByVal shp As Shape)
    On Error Resume Next
    With shp.Fill
        .TwoColorGradient 1, 1        ' msoGradientHorizontal, variant1(上→下)
        .ForeColor.RGB = RGB(22, 163, 88)    ' わずかに明るい緑
        .BackColor.RGB = RGB(0, 122, 55)     ' 深い緑
    End With
    On Error GoTo 0
End Sub

' StyleBubble - 新規バブル生成直後にフォントだけ適用(BeautifyAll軽量版)。
Public Sub StyleBubble(ByVal shp As Shape)
    On Error Resume Next
    shp.TextFrame2.TextRange.Font.Name = "Yu Gothic UI"
    On Error GoTo 0
End Sub

' 柔らかいドロップシャドウ(透明度高め・ぼかし広め)。固定バー/カード/選択中
' バブルにだけ使う(全面に付けると描画負荷でFPSが落ちる)。
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

' スキン(きせかえ): 称号と同じ「感謝受領数」ゲート。light/darkは全員、
'   sakura/oceanは感謝5件、goldは20件。解放判定は色解決点(ResolveColor)でも
'   強制するので、隠しシートを手書きしても色は落ちる。

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
                ' 2026-08-01(R12-7-6・a11y監査Med): 旧bg比3.41/surface比3.74は
                ' 基準4.5未満(8pt多用でほぼ全ペア未達)。bg比5.3/surface比5.81へ(機械計算)。
                Case "muted":         ResolveColor = RGB(130, 90, 105)
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
                ' 2026-08-01(R12-7-6・a11y監査Med): 旧bg比3.91/surface比4.22は
                ' 基準4.5未満。bg比5.28/surface比5.7へ(機械計算)。
                Case "muted":         ResolveColor = RGB(80, 105, 130)
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
                ' 2026-08-01(R12-7-6・a11y監査Med): 旧bg比4.39は8pt多用時の
                ' 基準4.5に僅かに未達。bg比5.6/surface比6.16へ(機械計算)。
                Case "muted":         ResolveColor = RGB(90, 98, 110)
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

' きせかえ切替(ヘルプの🎨から)。解放済みを巡回し、未解放は「あと◯件で
' 解放」のティザーToastを出してスキップ(欲しくなる導線)。
Public Sub CycleSkin()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    Dim orderList As Variant
    orderList = Array("light", "dark", "sakura", "ocean", "gold")
    Dim labels As Variant
    labels = Array("MS&AD スタンダード", "ダークモード", "サクラ・ピンク", "オーシャン・ブルー", "エグゼクティブ・ゴールド")

    Dim cur As String: cur = EffectiveSkin(CurrentTheme())
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
            SaveTheme cand
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
' waitless(R11-H Med4): Trueなら1.1秒の待機・削除をせず描いたらすぐ戻る
' (「一言返すだけ」は待たせること自体が害)。残ったToastは次の
' ShowToast/PaintProgressの掃除で消える。既定は待って消す。
Public Sub ShowToast(ByVal message As String, Optional ByVal kind As String = "info", _
                     Optional ByVal waitless As Boolean = False)
    On Error Resume Next
    ' 別ブック誤爆ガード: 他の業務Excelを見ている間にToastを描くと、他人の
    ' ブックへShapeを生成して業務データを汚す。自ブックがアクティブな時だけ描く。
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
    If waitless Then
        On Error GoTo 0
        Exit Sub
    End If
    ToastWait 1100
    ws.Shapes("nx_toast").Delete
    On Error GoTo 0
End Sub

' Timer基準の短時間待機(DoEventsで応答性維持。Sleep宣言を避けbitness非依存)。
Private Sub ToastWait(ByVal ms As Long)
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000# < ms
        DoEvents
        If Timer < t0 Then Exit Do   ' 深夜0時のTimerロールオーバーガード
    Loop
End Sub

' ----------------------------------------------------------------------------
' テーマ(配色の適用)。2026-07-31(R11-F1)にmodUIから移設(modUIが上限まで
'   残り168字となり修正不能だったため、配色の単一情報源へテーマ塊を寄せた。
'   憲章§4-6)。modUI側にはUiColor/UiTheme/ToggleThemeの薄い委譲だけを残す。
' ----------------------------------------------------------------------------

' ----------------------------------------------------------------------------
' CurrentTheme / SaveTheme - 現在のテーマ(スキン)名の読み書き。
'   2026-07-31(R11-F1): ui_state走査の自前実装をmodState.LoadState/SaveStateへの
'   委譲に置換(同じ"nexus_theme"を2通りで読み書きし同型処理が2つあった=憲章§4-5)。
'   既定は"light"。値はスキン名も入る(検証はEffectiveSkin/ResolveColor側)。
' ----------------------------------------------------------------------------
Public Function CurrentTheme() As String
    CurrentTheme = LCase$(Trim$(modState.LoadState(THEME_KEY, "light")))
    If LenB(CurrentTheme) = 0 Then CurrentTheme = "light"
End Function

Public Sub SaveTheme(ByVal themeName As String)
    modState.SaveState THEME_KEY, themeName
End Sub

' 配色解決はmodSkin.ResolveColorへ委譲。
Public Function ThemeColor(ByVal key As String) As Long
    ThemeColor = ResolveColor(key, CurrentTheme())
End Function

' ----------------------------------------------------------------------------
' ExtendChatBand - チャットの塗りと ScrollArea を会話の実下端へ追随させる。
' ----------------------------------------------------------------------------
' R19-1b(実機第6報①): 旧実装は起動のたびに A1:P2000(約31,000pt=40画面ぶん・
' 32,000セル)を塗っていた。これは制限ではなく「無限スクロールの明文許可」で、
' 塗ったぶんだけブックにスタイルが焼き付いてもいた。会話は下へ伸び続けるので、
' 下端は mChatBottom を単一情報源にして【バブルを足すたびに伸ばす】。
' 固定領域(行1〜4=ヘッダー/入力欄)は ApplyTheme が別に塗るのでここでは触らない
' (触ると入力欄の白い箱を塗り潰す)。すでに足りているときは何もしない。
Public Sub ExtendChatBand(ByVal ws As Worksheet, ByVal bottomY As Double)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Dim addr As String: addr = modUINexusDraw.NexusBound(ws)
    Dim lastRow As Long, fromRow As Long
    lastRow = ws.Range(addr).Rows.Count
    If lastRow > mChatBandRow Then
        fromRow = mChatBandRow + 1
        If fromRow < modUINexusDraw.INPUT_ROW + 2 Then fromRow = modUINexusDraw.INPUT_ROW + 2
        ws.Range("A" & fromRow & ":" & _
                 modUINexusDraw.NEXUS_PAD_COL & lastRow).Interior.Color = ThemeColor("bg")
    End If
    mChatBandRow = lastRow
    modViewport.ApplyScrollBound ws, addr
    On Error GoTo 0
End Sub

' テーマ適用(背景+全nx_Shape再彩色)。
Public Sub ApplyTheme(ByVal ws As Worksheet)
    ' R18-3a/R19-1b: 全域(ws.Cells)への塗りはUsedRangeをシート最大へ膨らませ、
    ' 「右にも下にも無限にスクロールできる」状態を作る(調査agent2 §1.3)。
    ' 呼び出し元は全てNexus(チャット)シートなので、会話の実下端から決まる
    ' 実使用範囲(NexusBound)だけを塗る。フォントもここで当てる(旧 modUI.InitUI
    ' の A1:P2000 へのFont.Nameは廃止した)。
    Dim bandAddr As String: bandAddr = modUINexusDraw.NexusBound(ws)
    ws.Range(bandAddr).Interior.Color = ThemeColor("bg")
    On Error Resume Next
    ws.Range(bandAddr).Font.Name = "Yu Gothic UI"
    mChatBandRow = ws.Range(bandAddr).Rows.Count
    modViewport.ApplyScrollBound ws, bandAddr
    On Error GoTo 0

    ' 入力欄(C3:K3)は上の一括塗りで消えるため塗り直す(両端B/L列は
    ' あえて無地のまま=入力欄に見せない)。
    On Error Resume Next
    With ws.Range("C" & modUINexusDraw.INPUT_ROW & ":K" & modUINexusDraw.INPUT_ROW)
        .Interior.Color = RGB(255, 255, 255)
        .BorderAround LineStyle:=1, Weight:=2, Color:=ThemeColor("border")
    End With
    On Error GoTo 0

    ' R14-6c(実機第3報 RC5-E): ここに一括OERNが無かったため、1個のShapeで
    ' 例外(壊れた参照・保護シートでの書込不可等)が出るとループ全体が
    ' 止まり、残りのShapeが旧テーマ色のまま取り残されていた。モジュール
    ' 冒頭の方針(影プロパティはOERN配下)を、色のFill/Line代入にも広げる。
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        Dim nm As String: nm = shp.Name
        If Left$(nm, 3) <> "nx_" Then GoTo NextShp

        If Left$(nm, 7) = "nx_top_" Then
            If nm = "nx_top_bg" Then
                ' ヘッダーバーは濃色(ロゴ・操作pillの白文字が乗る)。
                shp.Fill.ForeColor.RGB = ThemeColor("sidebar")
                SetShapeTextColor shp, RGB(255, 255, 255)
            ElseIf nm = "nx_top_send" Or nm = "nx_top_add" Then
                shp.Fill.ForeColor.RGB = ThemeColor("accent")
                SetShapeTextColor shp, RGB(255, 255, 255)
            ElseIf nm = "nx_top_theme" Then
                shp.Fill.ForeColor.RGB = ThemeColor("sidebarActive")
                SetShapeTextColor shp, RGB(255, 255, 255)
                shp.TextFrame2.TextRange.Text = ThemeIcon()
            Else
                ' ヘッダー上の操作pill(back/clear/lang/mode/speed)。
                shp.Fill.ForeColor.RGB = ThemeColor("sidebarActive")
                SetShapeTextColor shp, RGB(255, 255, 255)
            End If
        ElseIf Left$(nm, 7) = "nx_act_" Then
            PaintActionButton shp, Mid$(nm, 8)
        ElseIf Left$(nm, 9) = "nx_msg_u_" Then
            PaintBubble shp, True
        ElseIf Left$(nm, 9) = "nx_msg_a_" Then
            PaintBubble shp, False
        ElseIf Left$(nm, 7) = "nx_thk_" Then
            SetShapeTextColor shp, ThemeColor("muted")
        ElseIf nm = "nx_fchip" Then
            ' R14-6c(RC5-D): modAppAct.DrawFollowupChipの生成色(accent塗り+
            ' 枠無し+白文字)と揃える。無いと着せ替え後もチップだけ旧色で残る。
            shp.Fill.ForeColor.RGB = ThemeColor("accent")
            shp.Line.Visible = 0
            SetShapeTextColor shp, RGB(255, 255, 255)
        ElseIf Left$(nm, 8) = "nx_help_" Then
            RecolorHelpShape shp, nm
        End If
NextShp:
    Next shp
    On Error GoTo 0
End Sub

' R14-6c(RC5-D): nx_help_*(modHelp.ShowHelpCard生成)の再彩色。生成時の
' 配色をそのまま踏襲する(modHelp側の配色を変えたらここも合わせること)。
Private Sub RecolorHelpShape(ByVal shp As Shape, ByVal nm As String)
    Select Case nm
        Case "nx_help_card"
            shp.Fill.ForeColor.RGB = ThemeColor("surface")
            shp.Line.ForeColor.RGB = ThemeColor("primary")
            SetShapeTextColor shp, ThemeColor("text")
        Case "nx_help_manual"
            shp.Fill.ForeColor.RGB = ThemeColor("primary")
            SetShapeTextColor shp, RGB(255, 255, 255)
        Case "nx_help_tour", "nx_help_fb"
            shp.Fill.ForeColor.RGB = ThemeColor("surface")
            shp.Line.ForeColor.RGB = ThemeColor("accent")
            SetShapeTextColor shp, ThemeColor("accent")
        Case Else
            ' cfg/skin/migout/migin/diag(AddHelpActionの共通配色)。
            shp.Fill.ForeColor.RGB = ThemeColor("surface")
            shp.Line.ForeColor.RGB = ThemeColor("border")
            SetShapeTextColor shp, ThemeColor("text")
    End Select
End Sub

Public Sub PaintBubble(ByVal shp As Shape, ByVal isUser As Boolean)
    On Error Resume Next   ' R14-6c: Fill/Line代入もOERN配下(モジュール方針)
    If isUser Then
        shp.Fill.ForeColor.RGB = ThemeColor("userBubble")
        shp.Line.Visible = 0
        ' §9(Apple風): 自分の発言だけ微グラデーションで奥行きを出す。
        ' 2026-08-01(R12-7-2): 終端は primary ではなく userBubble を14%
        ' 暗くした同系色にする。primary終端だとバブル下端が実質primary色に
        ' なり、全面に乗る text 色との対比が全テーマで3:1未満(酷いものは
        ' gold 1.75)まで落ちていた(a11y監査Med)。同系微差色なら
        ' text-on-userBubble の高い対比(9.1〜13.3、機械計算)をほぼ保てる。
        ApplyGradient shp, ThemeColor("userBubble"), DarkenRgb(ThemeColor("userBubble"), 0.14)
    Else
        shp.Fill.ForeColor.RGB = ThemeColor("aiBubble")
        shp.Line.Visible = -1
        shp.Line.ForeColor.RGB = ThemeColor("border")
        shp.Line.Weight = 0.75
    End If
    On Error GoTo 0
    SetShapeTextColor shp, ThemeColor("text")
End Sub

Public Sub PaintActionButton(ByVal shp As Shape, ByVal kind As String)
    On Error Resume Next   ' R14-6c: Fill/Line代入もOERN配下(モジュール方針)
    shp.Fill.ForeColor.RGB = ThemeColor("surface")
    shp.Line.Visible = -1
    shp.Line.Weight = 0.75
    Select Case kind
        Case "resolve"
            ' 2026-08-01(R12-7-5・a11y監査Med): 白surface上で比2.54(3:1未満)
            ' だった。フィードバック機構の主ボタンのため濃緑へ(白地で比5.48、
            ' 機械計算)。
            shp.Line.ForeColor.RGB = RGB(4, 120, 87)
            SetShapeTextColor shp, RGB(4, 120, 87)
        Case "unsure"
            shp.Line.ForeColor.RGB = RGB(245, 158, 11)
            SetShapeTextColor shp, RGB(180, 110, 8)
        Case "bad"
            ' 2026-08-01(R12-7-5): 枠が白surface上で比2.56(3:1未満)だった。
            ' 濃いスレートグレーへ(白地で比4.76、機械計算)。
            shp.Line.ForeColor.RGB = RGB(100, 116, 139)
            SetShapeTextColor shp, ThemeColor("muted")
        Case "conf"
            ' 信頼度バッジは枠も塗りも持たない文字だけの表示。
            shp.Line.Visible = 0
            shp.Fill.Visible = 0
            SetShapeTextColor shp, ThemeColor("muted")
        Case "hq"
            shp.Line.ForeColor.RGB = RGB(239, 68, 68)
            SetShapeTextColor shp, RGB(239, 68, 68)
        Case "word"
            shp.Line.ForeColor.RGB = ThemeColor("primary")
            SetShapeTextColor shp, ThemeColor("primary")
        Case Else
            shp.Line.ForeColor.RGB = ThemeColor("border")
            SetShapeTextColor shp, ThemeColor("text")
    End Select
    On Error GoTo 0
End Sub

Public Sub SetShapeTextColor(ByVal shp As Shape, ByVal rgbVal As Long)
    On Error Resume Next
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgbVal
    On Error GoTo 0
End Sub

Public Function ThemeIcon() As String
    If CurrentTheme() = "dark" Then
        ThemeIcon = ChrW(&H2600)    ' 太陽
    Else
        ThemeIcon = ChrW(&HD83C) & ChrW(&HDF19)   ' 月
    End If
End Function
