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
' 2026-09-08(R41 B1): ShowToast本体とToastWaitはsrc/ui/modToast.basへ移設した
' (本モジュールが上限まで残り27字で1行も直せなかったため。憲章§4-6)。

Private Const THEME_KEY As String = "nexus_theme"

' R19-1b: 直近にチャットの塗りを当てた最終行。会話が伸びたぶんだけ塗り足す
' ための目印で、同じ範囲を毎バブル塗り直さないためだけに持つ(状態の本体は
' modUI.mChatBottom。ここはその写像なので、ズレても塗りが1回増えるだけ)。
Private mChatBandRow As Long

' R27H F2(M-1裁定): 重ね表示(ヘルプ/プレビュー/ツアー)の下端。会話の下端とは
' 別に持ち、ExtendChatBand が下端の下限として使う。0=重ね表示なし。
' 開いている間に別経路(質問送信・トグル)が会話の下端で境界を貼り直しても、
' カードが境界の外へ落ちない(=下に白い余白が出ない)ようにするための床。
Private mOverlayFloor As Double

' 重ね表示を描いた側が「カード下端+12」を渡す。閉じる側は必ず Clear する。
Public Sub SetOverlayFloor(ByVal y As Double)
    mOverlayFloor = y
End Sub

Public Sub ClearOverlayFloor()
    mOverlayFloor = 0
End Sub

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

' ヘッダーバーの2色グラデーション。2026-08-06(R20-7a・実機第7報⑥):
' 旧#1a365d→#1f4e78(青系)をMS&AD確定パレットへ差し替え。テーマに関わらず
' 常設の"ブランドの色"(旧実装も着せ替えの影響を受けない固定色だった)。
' 上段PRIMARY_LIGHT(#0B7D6E・白字コントラスト5.0:1)→下段PRIMARY_DARK
' (#014D44・白字コントラスト9.8:1、どちらもAA合格・司令塔で計算済み)。
Public Sub ApplyHeaderDepth(ByVal shp As Shape)
    On Error Resume Next
    With shp.Fill
        .TwoColorGradient 1, 1        ' msoGradientHorizontal, variant1
        .ForeColor.RGB = RGB(11, 125, 110)   ' PRIMARY_LIGHT #0B7D6E
        .BackColor.RGB = RGB(1, 77, 68)      ' PRIMARY_DARK #014D44
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
' 2026-08-06(R20-7a): ACCENT(#07A963)は白字コントラストが3.06:1でAA未達
' (機械計算・司令塔のLINE緑不採用と同型の理由)のため、送信ボタンの白字と
' 組む面はPRIMARY系(白字5.0/6.8:1・AA合格)に統一する。
Public Sub ApplyGreenDepth(ByVal shp As Shape)
    On Error Resume Next
    With shp.Fill
        .TwoColorGradient 1, 1        ' msoGradientHorizontal, variant1(上→下)
        .ForeColor.RGB = RGB(11, 125, 110)   ' PRIMARY_LIGHT #0B7D6E
        .BackColor.RGB = RGB(1, 103, 91)     ' PRIMARY #01675B
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

' テーマ名を検証し、未解放/未知なら既定(msad)へ落とした正規名を返す。
' 2026-08-06(R20-7a): 既定を"light"→"msad"(MS&AD標準)へ。lightは着せ替えの
' 1枚として引き続き選べる(既存ユーザーの保存値を壊さない)。
Public Function EffectiveSkin(ByVal themeName As String) As String
    Dim t As String: t = LCase$(Trim$(themeName))
    Select Case t
        Case "dark":            EffectiveSkin = "dark"
        Case "light":           EffectiveSkin = "light"
        Case "sakura", "ocean": EffectiveSkin = IIf(ThanksCount() >= 5, t, "msad")
        Case "gold":            EffectiveSkin = IIf(ThanksCount() >= 20, t, "msad")
        Case Else:              EffectiveSkin = "msad"
    End Select
End Function

Private Function ThanksCount() As Long
    On Error Resume Next
    ThanksCount = modStats.GetStat("thanks_received_total")
    On Error GoTo 0
End Function

' 全画面の配色解決点(modUI.ThemeColorから委譲される唯一の実装)。
' userBubbleText(R20-7a追加): 自分バブルの文字色。msadは塗りが濃緑になる
' ため白固定、他は従来どおりtextと同色(見た目を変えない)。
Public Function ResolveColor(ByVal key As String, ByVal themeName As String) As Long
    ' R43 2-1: 実体はmodKnowledgeBar.OnAccentColorへ(modSkinの容量逼迫のため)。
    If key = "onAccent" Then
        ResolveColor = modKnowledgeBar.OnAccentColor(ResolveColor("accent", themeName))
        Exit Function
    End If
    Dim t As String: t = EffectiveSkin(themeName)
    Select Case t
        Case "dark"
            Select Case key
                Case "bg":            ResolveColor = RGB(15, 23, 42)
                Case "surface":       ResolveColor = RGB(30, 41, 59)
                Case "text":          ResolveColor = RGB(248, 250, 252)
                Case "muted":         ResolveColor = RGB(148, 163, 184)
                Case "border":        ResolveColor = RGB(51, 65, 85)
                ' 2026-08-06(R20-7b): PRIMARY_LIGHT(#0B7D6E)を基調に(白字5.0:1・AA合格)。
                Case "primary":       ResolveColor = RGB(11, 125, 110)
                Case "accent":        ResolveColor = RGB(0, 168, 89)
                Case "userBubble":    ResolveColor = RGB(51, 65, 85)
                Case "userBubbleText": ResolveColor = RGB(248, 250, 252)
                Case "aiBubble":      ResolveColor = RGB(30, 41, 59)
                Case "sidebar":       ResolveColor = RGB(11, 15, 25)
                Case "sidebarText":   ResolveColor = RGB(209, 213, 219)
                Case "sidebarActive": ResolveColor = RGB(30, 41, 59)
                ' R30 W2-3: 逆質問番号ヒントの赤太字用(暗bg比6.45:1・機械計算)。
                Case "danger":        ResolveColor = RGB(248, 113, 113)
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
                Case "userBubbleText": ResolveColor = RGB(66, 32, 44)
                Case "aiBubble":      ResolveColor = RGB(255, 255, 255)
                Case "sidebar":       ResolveColor = RGB(84, 32, 52)
                Case "sidebarText":   ResolveColor = RGB(240, 210, 222)
                Case "sidebarActive": ResolveColor = RGB(112, 46, 72)
                Case "danger":        ResolveColor = RGB(185, 28, 28)
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
                Case "userBubbleText": ResolveColor = RGB(15, 36, 62)
                Case "aiBubble":      ResolveColor = RGB(255, 255, 255)
                Case "sidebar":       ResolveColor = RGB(10, 35, 66)
                Case "sidebarText":   ResolveColor = RGB(198, 219, 240)
                Case "sidebarActive": ResolveColor = RGB(20, 56, 96)
                Case "danger":        ResolveColor = RGB(185, 28, 28)
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
                Case "userBubbleText": ResolveColor = RGB(240, 234, 216)
                Case "aiBubble":      ResolveColor = RGB(24, 24, 28)
                Case "sidebar":       ResolveColor = RGB(6, 6, 8)
                Case "sidebarText":   ResolveColor = RGB(212, 175, 55)
                Case "sidebarActive": ResolveColor = RGB(38, 34, 24)
                Case "danger":        ResolveColor = RGB(248, 113, 113)
                Case Else:            ResolveColor = RGB(0, 0, 0)
            End Select
        Case "light"    ' 旧既定。着せ替えの1枚として残す(値は従来のまま)。
            Select Case key
                Case "bg":            ResolveColor = RGB(243, 244, 246)
                Case "surface":       ResolveColor = RGB(255, 255, 255)
                Case "text":          ResolveColor = RGB(17, 24, 39)
                Case "muted":         ResolveColor = RGB(90, 98, 110)
                Case "border":        ResolveColor = RGB(229, 231, 235)
                Case "primary":       ResolveColor = RGB(0, 137, 62)
                Case "accent":        ResolveColor = RGB(0, 168, 89)
                Case "userBubble":    ResolveColor = RGB(239, 246, 255)
                Case "userBubbleText": ResolveColor = RGB(17, 24, 39)
                Case "aiBubble":      ResolveColor = RGB(255, 255, 255)
                Case "sidebar":       ResolveColor = RGB(17, 24, 39)
                Case "sidebarText":   ResolveColor = RGB(209, 213, 219)
                Case "sidebarActive": ResolveColor = RGB(31, 41, 55)
                Case "danger":        ResolveColor = RGB(185, 28, 28)
                Case Else:            ResolveColor = RGB(0, 0, 0)
            End Select
        Case Else       ' "msad" = 既定(実機第7報⑥・確定パレット)
            Select Case key
                Case "bg":            ResolveColor = RGB(243, 244, 246)   ' 既存の薄灰を維持
                Case "surface":       ResolveColor = RGB(255, 255, 255)
                Case "text":          ResolveColor = RGB(17, 24, 39)      ' 既存#1F2937系を維持
                Case "muted":         ResolveColor = RGB(90, 98, 110)
                Case "border":        ResolveColor = RGB(229, 231, 235)
                Case "primary":       ResolveColor = RGB(1, 103, 91)      ' PRIMARY #01675B(白字6.8:1)
                Case "accent":        ResolveColor = RGB(7, 169, 99)      ' ACCENT #07A963(使用は控えめに)
                Case "userBubble":    ResolveColor = RGB(11, 125, 110)    ' PRIMARY_LIGHT(グラデ上段)
                Case "userBubbleText": ResolveColor = RGB(255, 255, 255) ' 自分バブルは白字固定
                Case "aiBubble":      ResolveColor = RGB(255, 255, 255)
                Case "sidebar":       ResolveColor = RGB(1, 77, 68)       ' PRIMARY_DARK(ヘッダー系)
                Case "sidebarText":   ResolveColor = RGB(255, 255, 255)
                Case "sidebarActive": ResolveColor = RGB(1, 103, 91)      ' PRIMARY
                Case "danger":        ResolveColor = RGB(185, 28, 28)
                Case Else:            ResolveColor = RGB(0, 0, 0)
            End Select
    End Select
End Function

' きせかえ切替(ヘルプの🎨から)。解放済みを巡回し、未解放は「あと◯件で
' 解放」のティザーToastを出してスキップ(欲しくなる導線)。
Public Sub CycleSkin()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    ' 2026-08-06(R20-7a): 既定skinを"msad"に。旧既定"light"は着せ替えの1枚として残す。
    ' R29 W2-4: 巡回2番目を"dark"へ(1回押した時点で必ず視覚差が出るように)。
    Dim orderList As Variant
    orderList = Array("msad", "dark", "light", "sakura", "ocean", "gold")
    Dim labels As Variant
    labels = Array("MS&AD スタンダード", "ダークモード", "ライト(旧配色)", "サクラ・ピンク", "オーシャン・ブルー", "エグゼクティブ・ゴールド")

    Dim cur As String: cur = EffectiveSkin(CurrentTheme())
    Dim curIdx As Long: curIdx = 0
    Dim i As Long
    For i = 0 To 5
        If CStr(orderList(i)) = cur Then curIdx = i
    Next i

    Dim tc As Long: tc = ThanksCount()
    Dim tried As Long
    For tried = 1 To 6
        Dim nx As Long: nx = (curIdx + tried) Mod 6
        Dim cand As String: cand = CStr(orderList(nx))
        Dim needN As Long
        needN = 0
        If cand = "sakura" Or cand = "ocean" Then needN = 5
        If cand = "gold" Then needN = 20
        If tc >= needN Then
            SaveTheme cand
            modUI.Repaint
            ' R29H F2b: ループ内ティザートースト(未解放を1件ずつ巡回waitless発火)は
            ' 廃止し、切替成功トースト末尾へ未解放テーマの案内を1回だけ付記する。
            modSkin.ShowToast "きせかえ: " & CStr(labels(nx)) & _
                IIf(needN > 0, "(感謝" & needN & "件の限定スキン)", "") & _
                modChrome.UnlockedTeaserSuffix(tc), "success"
            GoTo Done
        End If
    Next tried
Done:
    modUiLock.Leave
End Sub

' ----------------------------------------------------------------------------
' ShowToast - MsgBoxの代替(非ブロッキング通知)。実体はmodToast.ShowToastへ
'   移設した(2026-09-08 R41 B1)。シグネチャは不変(lint契約のR1例外は
'   modSkin.ShowToastの名前で通っているため、mid層の呼び口は変えない)。
' ----------------------------------------------------------------------------
Public Sub ShowToast(ByVal message As String, Optional ByVal kind As String = "info", _
                     Optional ByVal waitless As Boolean = False)
    modToast.ShowToast message, kind, waitless
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
'   既定は"msad"(2026-08-06 R20-7a: 実機第7報⑥でMS&AD標準へ変更)。
'   値はスキン名も入る(検証はEffectiveSkin/ResolveColor側)。
' ----------------------------------------------------------------------------
Public Function CurrentTheme() As String
    CurrentTheme = LCase$(Trim$(modState.LoadState(THEME_KEY, "msad")))
    If LenB(CurrentTheme) = 0 Then CurrentTheme = "msad"
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
    ' R27H F2: 重ね表示中は、その下端より上へ境界を縮めない(境界衝突の封鎖)。
    If mOverlayFloor > bottomY Then bottomY = mOverlayFloor
    ' 下端は呼び出し側が渡す実測値(modUI.mChatBottom)から直に組む。
    ' modUINexusDraw.NexusBound と同じ算数だが、こちらは会話の下端が確定した
    ' 直後に呼ばれるので、値を取りに戻る往復を省く(式の持ち主は BoundAddr 1本)。
    Dim addr As String
    addr = modViewport.BoundAddr(ws, modUINexusDraw.NEXUS_PAD_COL, bottomY, _
                                 modUINexusDraw.NEXUS_MAX_ROW)
    Dim lastRow As Long, fromRow As Long
    lastRow = ws.Range(addr).Rows.Count
    If lastRow > mChatBandRow Then
        fromRow = mChatBandRow + 1
        If fromRow < modUINexusDraw.INPUT_ROW + 2 Then fromRow = modUINexusDraw.INPUT_ROW + 2
        ' R30W1-3: 新たにバンドへ入る行を18ptにしてから塗る(バンド内全行18pt
        ' の不変式。modUI.ScrollToBottom の18pt換算がこれに依存する)。
        ws.Rows(fromRow & ":" & lastRow).RowHeight = 18
        ws.Range("A" & fromRow & ":" & _
                 modUINexusDraw.NEXUS_PAD_COL & lastRow).Interior.Color = ThemeColor("bg")
    ElseIf lastRow < mChatBandRow Then
        ' R19H FA-5(i) → R28H F2(M-1) → R30W1-4。縮んだぶんを「塗り直す」のを
        ' やめ、【行ごと解放する】。行が残っている限り塗り直しても内部使用範囲は
        ' 縮まず、FreezePanes併用のホイールはそこまで転がれる(実機第15報の真因)。
        ' 行を消せば地は modChrome.ApplyNormalStyleBg のNormalスタイル地色に
        ' なるので、バンド外に塗りは要らない。行1〜4は解放側が下限で守る。
        modViewport2.ReleaseRowsBelow ws, lastRow + 3
    End If
    mChatBandRow = lastRow
    modViewport.ApplyScrollBound ws, addr
    On Error GoTo 0
End Sub

' テーマ適用(背景+全nx_Shape再彩色)。
Public Sub ApplyTheme(ByVal ws As Worksheet)
    modChrome.ApplyNormalStyleBg ws   ' R28波1: 地(Normalスタイル)も切替に追随
    ' R18-3a/R19-1b: 全域(ws.Cells)への塗りはUsedRangeをシート最大へ膨らませ、
    ' 「右にも下にも無限にスクロールできる」状態を作る(調査agent2 §1.3)。
    ' 呼び出し元は全てNexus(チャット)シートなので、会話の実下端から決まる
    ' 実使用範囲(NexusBound)だけを塗る。フォントもここで当てる(旧 modUI.InitUI
    ' の A1:P2000 へのFont.Nameは廃止した)。
    On Error Resume Next
    Dim bandAddr As String: bandAddr = modUINexusDraw.NexusBound(ws)
    ws.Range(bandAddr).Interior.Color = ThemeColor("bg")
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
                ' 2026-08-06(R20-7a): accentは白字コントラスト3.06:1でAA未達
                ' (機械計算)のため、白文字と組む塗りはprimary(6.8:1)にする。
                shp.Fill.ForeColor.RGB = ThemeColor("primary")
                SetShapeTextColor shp, RGB(255, 255, 255)
            ElseIf nm = "nx_top_theme" Then
                ' R24-2b白地仕様へ整合(単一情報源化)。
                shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
                shp.Line.ForeColor.RGB = RGB(1, 77, 68)
                SetShapeTextColor shp, RGB(1, 77, 68)
                shp.TextFrame2.TextRange.Text = ThemeIcon()
            Else
                ' ヘッダー上の操作pill(back/clear/lang/mode/speed)。R24-2b白地仕様。
                shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
                shp.Line.ForeColor.RGB = RGB(1, 77, 68)
                SetShapeTextColor shp, RGB(1, 77, 68)
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
            ' R14-6c(RC5-D): modAppAct.DrawFollowupChipの生成色に揃える
            ' (枠無し+白文字)。2026-08-06(R20-7a): 塗りはaccentではなく
            ' primaryへ(白字3.06→6.8:1。初回描画はmodAppAct側のaccentの
            ' ままだが、テーマ再適用のたびにここで正しい色へ揃う)。
            shp.Fill.ForeColor.RGB = ThemeColor("primary")
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
            ' 2026-08-06(R20-7a): accentは白地上の文字色としても3.06:1で
            ' AA未達(機械計算)のためprimary(6.8:1)へ。初回描画(modHelp側)は
            ' accentのままだが、再彩色のたびにここで正しい色へ揃う。
            shp.Fill.ForeColor.RGB = ThemeColor("surface")
            shp.Line.ForeColor.RGB = ThemeColor("primary")
            SetShapeTextColor shp, ThemeColor("primary")
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
        ' 2026-08-06(R20-7a): msadは仕様書確定の2段(上PRIMARY_LIGHT→下PRIMARY)を
        ' 固定で使う(userBubbleText=白と組んで6.8:1・AA合格)。他テーマは
        ' 従来どおりuserBubbleの自己微暗化を保つ(着せ替えの見た目を変えない)。
        Dim botC As Long
        If EffectiveSkin(CurrentTheme()) = "msad" Then
            botC = RGB(1, 103, 91)   ' PRIMARY #01675B
        Else
            botC = DarkenRgb(ThemeColor("userBubble"), 0.14)
        End If
        ApplyGradient shp, ThemeColor("userBubble"), botC
    Else
        shp.Fill.ForeColor.RGB = ThemeColor("aiBubble")
        shp.Line.Visible = -1
        shp.Line.ForeColor.RGB = ThemeColor("border")
        shp.Line.Weight = 0.75
    End If
    On Error GoTo 0
    SetShapeTextColor shp, IIf(isUser, ThemeColor("userBubbleText"), ThemeColor("text"))
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
