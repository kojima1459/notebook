Attribute VB_Name = "modUI"
Option Explicit

' modUI - Nexus Agent UIコア(Windows版Excel専用・SPA風UI)。ボタンはShape+
' OnAction。テーマはui_state(key=nexus_theme)。Shape命名はnx_接頭辞で分類。

Private Const NEXUS_SHEET As String = "Nexus"
Private Const THEME_KEY As String = "nexus_theme"
Private Const MSO_BRING_TO_FRONT As Long = 0   ' msoBringToFront(数値でLO互換)
Private Const MAX_BUBBLES As Long = 40         ' 32bitメモリ保護: 吹き出し保持上限

' 2026-07-26 再設計: サイドバー全廃。レイアウト定数(HDR_H/INPUT_ROW/ACT_H)と
' チャット領域の幾何(ChatLeft/ChatWidth/ChatTop)はmodUINexusDrawが単一情報源。
Private Const BUBBLE_RATIO As Double = 0.72  ' チャット幅に対するバブル最大幅
Private Const BUBBLE_GAP As Double = 14

Private mChatBottom As Double   ' 最後のバブルの下端(モジュール状態リセット時はRecalc)

' InitUI - ネイティブUI隠蔽+Nexus骨格描画
Public Sub InitUI()
    Dim ws As Worksheet
    Set ws = GetOrCreateNexusSheet()
    If ws Is Nothing Then Exit Sub

    ' 実機再発(2026-07-22): 保護状態(Protect)はブックの保存/再オープンをまたぐと
    ' UserInterfaceOnly=Trueが失効し、以降のCells.Clear等のマクロ操作が
    ' 「保護されたシート」err#1004で軒並み失敗し画面が真っ白になっていた。
    ' 必ず最初に解除してから組み立て、末尾で改めて保護し直す。
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    Application.ScreenUpdating = False

    On Error Resume Next
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    On Error GoTo 0
    On Error Resume Next
    Application.DisplayFormulaBar = False
    On Error GoTo 0
    On Error Resume Next
    Application.DisplayStatusBar = False
    On Error GoTo 0

    On Error Resume Next
    ws.Activate
    If Err.Number <> 0 Then
        modLog.LogError "E0801", "modApp.LaunchNexus", _
            "InitUI [ws.Activate] ws.Visible=" & ws.Visible & _
            " ActiveSheet=" & ThisWorkbook.ActiveSheet.Name & _
            " AppWin=" & Application.Windows.Count & " WbWin=" & ThisWorkbook.Windows.Count, Err.Number
        Err.Clear
    End If
    On Error GoTo 0

    ' 2026-07-28(レビュー L-17): ws.Activate に失敗していた場合、ここから
    ' 先の ActiveWindow 設定と FreezePanes は【別のシート】に適用される。
    ' 利用者から見ると「触っていない画面の見出しが消え、変な位置で固定
    ' された」という説明のつかない壊れ方になる。
    ' Repaint 側には既にあるシート一致ガードを、こちらにも入れる。
    Dim isFront As Boolean
    On Error Resume Next
    isFront = (ThisWorkbook.ActiveSheet Is ws)
    On Error GoTo 0
    If Not isFront Then Exit Sub

    On Error Resume Next
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
        .DisplayWorkbookTabs = False
        .DisplayHorizontalScrollBar = False
    End With
    On Error GoTo 0

    On Error Resume Next
    ActiveWorkbook.AutoSaveOn = False   ' D2: 自動保存がVBAへ割り込みクラッシュ/遅延するのを止める
    ActiveWindow.Zoom = 100             ' D4: Ctrl+ホイール等のズームでShape配置が崩れる基準を100%へ固定
    ws.EnableSelection = 1              ' C1: xlUnlockedCells(完全抑止はProtect併用時のみ。park運用と併せ誤選択を抑える)
    On Error GoTo 0

    ' Undoを封じる(Shapeと隠しDBの整合が崩れるため)。
    DisableUndoRedo

    ' --- キャンバス骨格 ---
    RemoveNexusShapes ws
    ws.Cells.Clear
    ws.Cells.Font.Name = "Yu Gothic UI"

    ' 幾何を先に確定させてからShapeを置く(順序が逆だと座標がズレる)。
    ' A=左余白 / B=📎 / C:K=入力欄(結合) / L=送信 / M=右余白。
    ws.Columns("A").ColumnWidth = 1.5
    ws.Columns("B").ColumnWidth = 4.5
    ws.Columns("C:K").ColumnWidth = 10.5
    ws.Columns("L").ColumnWidth = 11
    ws.Columns("M:P").ColumnWidth = 1.5
    ws.Rows("1:400").RowHeight = 18
    ws.Rows(1).RowHeight = modUINexusDraw.HDR_H
    ws.Rows(2).RowHeight = 8
    ws.Rows(modUINexusDraw.INPUT_ROW).RowHeight = 46
    ws.Rows(4).RowHeight = 12

    On Error Resume Next
    modUINexusDraw.DrawChatHeader ws
    If Err.Number <> 0 Then LogDrawStageError "DrawChatHeader", ws: Err.Clear
    On Error GoTo 0

    ' ヘッダーが右端ピルを2段目へ流した場合(極端に狭い列幅の端末)は、
    ' 行1をその実使用高さまで広げる。広げないと2段目が入力欄に重なる。
    ' ピルはPlacement=3で固定してあるので、行高を変えても位置は動かない。
    On Error Resume Next
    ws.Rows(1).RowHeight = modUINexusDraw.HeaderHeight()
    On Error GoTo 0

    On Error Resume Next
    modUINexusDraw.DrawInputArea ws
    If Err.Number <> 0 Then LogDrawStageError "DrawInputArea", ws: Err.Clear
    On Error GoTo 0

    ' 行1～4(ヘッダー+入力+ヒント)だけを固定。A5(=A列)を選ぶことで列は
    ' 固定しない(旧D8指定だとC列までが横方向にも固定されていた)。
    ' 保護をかける前に行う: EnableSelection=xlUnlockedCells下ではA5(ロック済み)を
    ' Selectできず、FreezePanesが黙って失敗するため。順序が意味を持つ。
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ws.Range("A5").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    ' 入力できるセルはC3:K3だけに絞る(UserInterfaceOnly=Trueなのでマクロは自由)。
    On Error Resume Next
    ws.Unprotect
    ws.Cells.Locked = True
    ws.Range("C" & modUINexusDraw.INPUT_ROW & ":K" & modUINexusDraw.INPUT_ROW).Locked = False
    ' 2026-07-28(レビュー M-25): 図形も保護する(DrawingObjects:=True)。
    ' False だとチャットの吹き出しをクリックすると白い選択ハンドルが付き、
    ' Delete キーで【回答が消える】。Ctrl+Z は無効化してあるので復元できない。
    ' バブルの OnAction 結線は廃止済みで、利用者がバブルを選択する必要は
    ' もう無い。UserInterfaceOnly:=True なのでマクロ側の描画は従来どおり動く。
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    ws.EnableSelection = 1
    On Error GoTo 0

    On Error Resume Next
    modTelemetry.TrackScreen "chat"
    On Error GoTo 0

    ApplyTheme ws
    FreezeShapePlacement ws   ' 全Shapeを絶対配置に固定(ズレ防止)
    BringFixedToFront ws      ' 固定UIを最前面へ(Z-Order維持)
    modSkin.BeautifyAll ws    ' フォント統一(Yu Gothic UI)+固定クロムに柔らかい影
    mChatBottom = modUINexusDraw.ChatTop(ws)

    Application.ScreenUpdating = True
    ParkFocus                 ' A2/C4: Shape選択解除+アクティブセルpark(白ハンドルを出さない)
End Sub

' AddChatBubble - チャットバブル1件を追加(role:user/ai。戻り値=バブルShape名)。
Public Function AddChatBubble(ByVal role As String, ByVal bodyText As String, _
                              Optional ByVal withActions As Boolean = False, _
                              Optional ByVal thinking As String = "") As String
    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then Exit Function

    If mChatBottom < modUINexusDraw.ChatTop(ws) Then RecalcChatBottom ws

    Dim isUser As Boolean
    isUser = (LCase$(role) = "user")

    Dim chatL As Double, chatW As Double
    chatL = modUINexusDraw.ChatLeft(ws)
    chatW = modUINexusDraw.ChatWidth(ws)
    Dim bubbleW As Double: bubbleW = chatW * BUBBLE_RATIO

    Dim topY As Double: topY = mChatBottom + BUBBLE_GAP
    Dim seq As String: seq = NextSeq(ws)

    ' --- 思考プロセス(AI限定) ---
    If (Not isUser) And LenB(Trim$(thinking)) > 0 Then
        Dim thk As Shape
        Set thk = ws.Shapes.AddShape(1, chatL + 8, topY, bubbleW - 16, 20)   ' 1=四角
        thk.Name = "nx_thk_" & seq
        thk.Line.Visible = 0
        thk.Fill.Visible = 0
        With thk.TextFrame2
            .WordWrap = -1
            .TextRange.Text = "思考プロセス: " & thinking
            .TextRange.Font.Size = 8.5
            .TextRange.Font.Italic = -1
            .AutoSize = 1   ' msoAutoSizeShapeToFitText
        End With
        topY = thk.Top + thk.Height + 4
    End If

    ' --- バブル本体(角丸四角) ---
    Dim bx As Double
    If isUser Then
        bx = chatL + chatW - bubbleW
    Else
        bx = chatL
    End If

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, bx, topY, bubbleW, 30)   ' 5=角丸四角
    shp.Name = "nx_msg_" & IIf(isUser, "u", "a") & "_" & seq
    shp.Adjustments(1) = 0.08   ' 角丸を小さめに(HTMLの12px相当)
    With shp.TextFrame2
        .WordWrap = -1
        .MarginLeft = 12: .MarginRight = 12: .MarginTop = 9: .MarginBottom = 9
        .TextRange.Text = bodyText
        .TextRange.Font.Size = 10.5
        .TextRange.ParagraphFormat.Alignment = 1   ' msoAlignLeft(左揃え。0は無効値)
        .AutoSize = 1
    End With
    If shp.Height < 28 Then shp.Height = 28
    shp.Shadow.Visible = 0

    ' 2026-07-26: アクションは常に「最新のAI回答」に紐づく文脈表示へ変えたため、
    ' 古いバブルをクリックして対象を切り替える操作は廃止した(仕様書§2.2)。

    PaintBubble shp, isUser
    modSkin.StyleBubble shp  ' Yu Gothic UI(バブルはフラット=影は選択時のみ)
    mChatBottom = shp.Top + shp.Height

    CapBubbles ws            ' 古い吹き出しを間引いてShape増殖(32bitクラッシュ)を防ぐ
    ScrollToBottom ws
    BringFixedToFront ws     ' 固定UI(サイドバー/トップバー/アクションバー)を最前面へ
    AddChatBubble = shp.Name
End Function

' UpdateBubbleText - 既に描いたバブルの本文を差し替える。
'
' なぜ必要か:
'   回答を待つ10～20秒(しっかり調べるなら1～2分)のあいだ、画面には
'   「考えています…」が1個あるだけだった。ところが実際には、検索は最初の
'   1～2秒で終わっていて、どの資料に答えがあるかはその時点で分かっている。
'   その一番おいしい情報を、利用者の見ていない旧ホームシートへ書いていた
'   (modUIMain.RenderSourcesPreview)。
'   ここを差し替えられるようにして、「もう見つけてある。いま文章にしている
'   だけ」という状態を待ち時間の主役にする。待たされている時間は変わらない
'   のに、体感はまるで別物になる。
'
'   AutoSizeで高さが変わるため、会話の下端(mChatBottom)は必ず取り直す。
'   取り直さないと、次に置くバブルがこのバブルへ重なる。
Public Sub UpdateBubbleText(ByVal shapeName As String, ByVal newText As String)
    If LenB(shapeName) = 0 Then Exit Sub

    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Dim shp As Shape
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub

    On Error Resume Next
    With shp.TextFrame2
        .WordWrap = -1
        .AutoSize = 0                    ' 一度切ってから入れ直さないと再フィットしない
        .TextRange.Text = newText
        .TextRange.Font.Size = 10.5
        .AutoSize = 1                    ' msoAutoSizeShapeToFitText
    End With
    If shp.Height < 28 Then shp.Height = 28

    RecalcChatBottom ws
    ScrollToBottom ws
    BringFixedToFront ws
    On Error GoTo 0
End Sub

' ChatBottomFor - 会話の現在の下端。modStarter が質問ボタンを積む基準に使う
'   (mChatBottom は Private なので、読み取り専用の窓口だけ開ける)。
Public Function ChatBottomFor(ByVal ws As Worksheet) As Double
    If ws Is Nothing Then Exit Function
    If mChatBottom < modUINexusDraw.ChatTop(ws) Then RecalcChatBottom ws
    ChatBottomFor = mChatBottom
End Function

' 固定UI(ヘッダー+入力欄)を最前面に維持(描画末に必ず呼ぶ)。
Public Sub BringFixedToFront(ByVal ws As Worksheet)
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 7) = "nx_top_" Then shp.ZOrder MSO_BRING_TO_FRONT
    Next shp
    On Error GoTo 0
End Sub

' nx_msg_が上限超過なら古い順に削除(nx_thk_も一緒に)。
Private Sub CapBubbles(ByVal ws As Worksheet)
    Dim names() As String, seqs() As Long
    ReDim names(0 To 255)
    ReDim seqs(0 To 255)
    Dim n As Long: n = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 7) = "nx_msg_" Then
            If n > UBound(names) Then
                ReDim Preserve names(0 To UBound(names) + 256)
                ReDim Preserve seqs(0 To UBound(seqs) + 256)
            End If
            names(n) = shp.Name
            seqs(n) = CLng(Val(Right$(shp.Name, 4)))
            n = n + 1
        End If
    Next shp
    If n <= MAX_BUBBLES Then Exit Sub

    Dim toRemove As Long: toRemove = n - MAX_BUBBLES

    ' seq昇順バブルソート(nは高々数百)
    Dim a As Long, b As Long
    For a = 0 To n - 2
        For b = 0 To n - 2 - a
            If seqs(b) > seqs(b + 1) Then
                Dim ts As Long: ts = seqs(b): seqs(b) = seqs(b + 1): seqs(b + 1) = ts
                Dim tn As String: tn = names(b): names(b) = names(b + 1): names(b + 1) = tn
            End If
        Next b
    Next a

    Dim ri As Long
    For ri = 0 To toRemove - 1
        On Error Resume Next
        ws.Shapes(names(ri)).Delete
        ws.Shapes("nx_thk_" & Format$(seqs(ri), "0000")).Delete
        On Error GoTo 0
    Next ri
End Sub

' ToggleTheme - ライト/ダーク反転+全体再彩色
Public Sub ToggleTheme()
    If CurrentTheme() = "dark" Then
        SaveTheme "light"
    Else
        SaveTheme "dark"
    End If

    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    On Error Resume Next   ' 再彩色が中断しても必ず暗転解除へ到達させる
    ApplyTheme ws
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

' RestoreExcelUI - ネイティブUI復元
Public Sub RestoreExcelUI()
    On Error Resume Next
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    On Error GoTo 0
    On Error Resume Next
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    With ActiveWindow
        .DisplayGridlines = True
        .DisplayHeadings = True
        .DisplayWorkbookTabs = True
        .DisplayHorizontalScrollBar = True
    End With
    ' Ctrl+Z/Ctrl+Yの無効化を解除(引数省略=既定へ戻す)。
    Application.OnKey "^z"
    Application.OnKey "^y"
    ' 砂時計/ステータスバーも念のため既定へ。
    Application.Cursor = -4143   ' xlDefault
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' GoToNexus - 「戻る」共通口。Activate失敗時はRestoreExcelUIで脱出路を残す。
Public Sub GoToNexus(ByVal source As String)
    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Activate
    If Err.Number <> 0 Then
        modLog.LogError "E0801", source, _
            "GoToNexus [ws.Activate失敗" & ChrW(&H2192) & "ネイティブタブ復元で脱出路確保] ws.Visible=" & ws.Visible & _
            " ActiveSheet=" & ThisWorkbook.ActiveSheet.Name & _
            " AppWin=" & Application.Windows.Count & " WbWin=" & ThisWorkbook.Windows.Count, Err.Number
        Err.Clear
        RestoreExcelUI
    End If
    On Error GoTo 0
End Sub

' GoToNativeSheet - GoToNexus同様の脱出路付き遷移をホーム/マイ本棚等にも提供。
Public Sub GoToNativeSheet(ByVal sheetName As String, ByVal source As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Activate
    If Err.Number <> 0 Then
        modLog.LogError "E0801", source, _
            "GoToNativeSheet(" & sheetName & ") [ws.Activate失敗" & ChrW(&H2192) & "ネイティブタブ復元で脱出路確保] ws.Visible=" & ws.Visible & _
            " ActiveSheet=" & ThisWorkbook.ActiveSheet.Name & _
            " AppWin=" & Application.Windows.Count & " WbWin=" & ThisWorkbook.Windows.Count, Err.Number
        Err.Clear
        RestoreExcelUI
    End If
    On Error GoTo 0
End Sub

' InitUIの各Draw*段の失敗記録役。
Private Sub LogDrawStageError(ByVal stageName As String, ByVal ws As Worksheet)
    Dim n As Long: n = Err.Number
    On Error Resume Next
    modLog.LogError "E0801", "modApp.LaunchNexus", _
        stageName & " [段階失敗・以降を継続] ActiveSheet=" & ThisWorkbook.ActiveSheet.Name, n
    On Error GoTo 0
End Sub

' Ctrl+Z/Ctrl+Y無効化(Shapeと隠しDBの整合をUndoが壊すため。復元はRestoreExcelUI)。
Private Sub DisableUndoRedo()
    On Error Resume Next
    Application.OnKey "^z", ""
    Application.OnKey "^y", ""
    On Error GoTo 0
End Sub

' Shape選択解除+アクティブセルpark(スクロール崩壊防止)。
Public Sub ParkFocus()
    On Error Resume Next
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub   ' 別ブックの選択状態を汚さない
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    If ws.Name = NEXUS_SHEET Then
        ' 入力セル(nx_input)へpark=Shape解除+次の入力に即備える
        ws.Range("C" & modUINexusDraw.INPUT_ROW).Select
    Else
        ws.Range("A1").Select   ' マイ本棚/Dashboard等は左上(固定領域)へpark
    End If
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

' 会話履歴を消さずにNexus画面を再描画(手動リフレッシュ用)。
Public Sub Repaint()
    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    On Error Resume Next
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    ActiveWorkbook.AutoSaveOn = False
    If ActiveSheet Is ws Then
        With ActiveWindow
            .DisplayGridlines = False
            .DisplayHeadings = False
            .DisplayWorkbookTabs = False
            .DisplayHorizontalScrollBar = False
            .Zoom = 100
        End With
    End If
    On Error GoTo 0

    ApplyTheme ws            ' 全nx_Shapeを再彩色(ゴースト=前画面の残像を塗り直す)
    FreezeShapePlacement ws  ' 絶対配置に再固定
    BringFixedToFront ws     ' 固定UIを最前面へ
    modSkin.BeautifyAll ws   ' フォント統一+固定クロムに柔らかい影
    Application.ScreenUpdating = True
    ParkFocus
End Sub

' Phase1暫定(Phase2でController移管)。
Public Sub NexusActionStub()
    MsgBox "このボタンは準備中です(次のフェーズで有効になります)。", vbInformation, "Nexus Agent"
End Sub

' 公開ゲッター: 他のNexus画面がテーマ一貫の配色/現在テーマを得る窓口。
Public Function UiColor(ByVal key As String) As Long
    UiColor = ThemeColor(key)
End Function

Public Function UiTheme() As String
    UiTheme = CurrentTheme()
End Function

' FreezeShapePlacement - 全Shapeを絶対配置固定しズレを防ぐ。
Public Sub FreezeShapePlacement(ByVal ws As Worksheet)
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Placement = 3   ' xlFreeFloating
    Next shp
    On Error GoTo 0
End Sub

'描画

'テーマ

Private Function CurrentTheme() As String
    CurrentTheme = "light"
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_UISTATE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(CStr(ws.Cells(i, 1).Value), THEME_KEY, vbTextCompare) = 0 Then
            Dim v As String
            v = LCase$(Trim$(CStr(ws.Cells(i, 2).Value)))
            If LenB(v) > 0 Then CurrentTheme = v   ' スキン名も可(検証はResolveColor側)
            Exit Function
        End If
    Next i
End Function

Private Sub SaveTheme(ByVal themeName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_UISTATE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim r As Long: r = 0
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(CStr(ws.Cells(i, 1).Value), THEME_KEY, vbTextCompare) = 0 Then
            r = i
            Exit For
        End If
    Next i
    If r = 0 Then
        r = lastRow + 1
        If r < 1 Then r = 1
        ws.Cells(r, 1).Value = THEME_KEY
    End If
    ws.Cells(r, 2).Value = themeName
End Sub

' 配色解決はmodSkin.ResolveColorへ委譲。
Private Function ThemeColor(ByVal key As String) As Long
    ThemeColor = modSkin.ResolveColor(key, CurrentTheme())
End Function

' テーマ適用(背景+全nx_Shape再彩色)。
Private Sub ApplyTheme(ByVal ws As Worksheet)
    ws.Cells.Interior.Color = ThemeColor("bg")

    ' 入力欄(C3:K3)は上の一括塗りで消えるため塗り直す(両端B/L列は
    ' あえて無地のまま=入力欄に見せない)。
    On Error Resume Next
    With ws.Range("C" & modUINexusDraw.INPUT_ROW & ":K" & modUINexusDraw.INPUT_ROW)
        .Interior.Color = RGB(255, 255, 255)
        .BorderAround LineStyle:=1, Weight:=2, Color:=ThemeColor("border")
    End With
    On Error GoTo 0

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
        End If
NextShp:
    Next shp
End Sub

Private Sub PaintBubble(ByVal shp As Shape, ByVal isUser As Boolean)
    If isUser Then
        shp.Fill.ForeColor.RGB = ThemeColor("userBubble")
        shp.Line.Visible = 0
        ' §9(Apple風): 自分の発言だけ濃紺の微グラデーションで奥行きを出す。
        modSkin.ApplyGradient shp, ThemeColor("userBubble"), ThemeColor("primary")
    Else
        shp.Fill.ForeColor.RGB = ThemeColor("aiBubble")
        shp.Line.Visible = -1
        shp.Line.ForeColor.RGB = ThemeColor("border")
        shp.Line.Weight = 0.75
    End If
    SetShapeTextColor shp, ThemeColor("text")
End Sub

Public Sub PaintActionButton(ByVal shp As Shape, ByVal kind As String)
    shp.Fill.ForeColor.RGB = ThemeColor("surface")
    shp.Line.Visible = -1
    shp.Line.Weight = 0.75
    Select Case kind
        Case "resolve"
            shp.Line.ForeColor.RGB = RGB(16, 185, 129)
            SetShapeTextColor shp, RGB(16, 185, 129)
        Case "unsure"
            shp.Line.ForeColor.RGB = RGB(245, 158, 11)
            SetShapeTextColor shp, RGB(180, 110, 8)
        Case "bad"
            shp.Line.ForeColor.RGB = RGB(148, 163, 184)
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
End Sub

Private Sub SetShapeTextColor(ByVal shp As Shape, ByVal rgbVal As Long)
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

'util

Private Function GetNexusSheet() As Worksheet
    On Error Resume Next
    Set GetNexusSheet = ThisWorkbook.Worksheets(NEXUS_SHEET)
    On Error GoTo 0
End Function

Private Function GetOrCreateNexusSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = NEXUS_SHEET
        On Error GoTo 0
    End If
    Set GetOrCreateNexusSheet = ws
    Exit Function
Fail:
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup27
FailCleanup27:
    ' Name代入失敗時、既定名のまま孤児化するのを防ぐ。
    If Not ws Is Nothing Then
        On Error Resume Next
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        On Error GoTo 0
    End If
    Set GetOrCreateNexusSheet = Nothing
End Function

Private Sub RemoveNexusShapes(ByVal ws As Worksheet)
    Dim names() As String
    ReDim names(0 To ws.Shapes.count)
    Dim n As Long: n = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 3) = "nx_" Then
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

' バブル連番(既存Shape数から採番)。
Private Function NextSeq(ByVal ws As Worksheet) As String
    ' 個数ではなく既存の最大seqを見る(CapBubbles後は個数が減り、個数+1だと
    ' 既存Shapeと同名衝突する実バグだった)。
    Dim maxN As Long: maxN = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 7) = "nx_msg_" Then
            Dim n As Long: n = CLng(Val(Right$(shp.Name, 4)))
            If n > maxN Then maxN = n
        End If
    Next shp
    NextSeq = Format$(maxN + 1, "0000")
End Function

' 会話領域の最下端を再計算する。固定クロム(nx_top_)以外の全nx_Shapeを見るので、
' バブルだけでなく文脈アクション(nx_act_)や出典チップ(nx_cite_)も自動で考慮され、
' 次のバブルがそれらに重ならない。1ターン描き終わるたびに呼ぶ。
Public Sub RecalcChatBottom(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    mChatBottom = modUINexusDraw.ChatTop(ws)
    Dim shp As Shape
    For Each shp In ws.Shapes
        ' 会話の流れに属するShapeだけを数える。除外ではなく明示の許可リストに
        ' したのは、トースト/ツアー/ヘルプ等の重ね表示(nx_toast/nx_tour_/
        ' nx_help_/nx_peek)が画面下部に出たときに下端が引きずられ、次の
        ' バブルが画面外へ飛ぶ事故を構造的に防ぐため。
        Dim nm As String: nm = shp.Name
        If Left$(nm, 7) = "nx_msg_" Or Left$(nm, 7) = "nx_thk_" _
           Or Left$(nm, 7) = "nx_act_" Or Left$(nm, 8) = "nx_cite_" _
           Or Left$(nm, 8) = "nx_conf_" Or Left$(nm, 6) = "nx_sq_" _
           Or Left$(nm, 10) = "nx_mentor_" Then
            If shp.Top + shp.Height > mChatBottom Then mChatBottom = shp.Top + shp.Height
        End If
    Next shp
End Sub

' 1ターン分(バブル+アクション+出典)を描き終えたあとの締め処理。
Public Sub SettleChat()
    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then Exit Sub
    RecalcChatBottom ws
    ScrollToBottom ws
End Sub

' ClearChat - 会話をクリアする(実機要望: 長い会話をリセットしたい)。
Public Sub ClearChat()
    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Dim names() As String
    ReDim names(0 To ws.Shapes.Count)
    Dim n As Long: n = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        Dim nm As String: nm = shp.Name
        If Left$(nm, 7) = "nx_msg_" Or Left$(nm, 7) = "nx_thk_" _
           Or Left$(nm, 7) = "nx_act_" Or Left$(nm, 8) = "nx_cite_" _
           Or Left$(nm, 8) = "nx_conf_" Or Left$(nm, 6) = "nx_sq_" Then
            names(n) = nm
            n = n + 1
        End If
    Next shp
    Dim i As Long
    For i = 0 To n - 1
        ws.Shapes(names(i)).Delete
    Next i
    mChatBottom = modUINexusDraw.ChatTop(ws)
    On Error GoTo 0
End Sub

' 選択中バブルを強調表示(primary色の太枠)。他は通常枠へ戻す。
Public Sub MarkActiveBubble(ByVal shapeName As String)
    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then Exit Sub

    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 9) = "nx_msg_a_" Then
            If shp.Name = shapeName Then
                shp.Line.Visible = -1
                shp.Line.ForeColor.RGB = ThemeColor("primary")
                shp.Line.Weight = 1.75
                modSkin.ApplySoftShadow shp   ' Active State: 選択中バブルだけ浮遊させる
            Else
                PaintBubble shp, False
                On Error Resume Next
                shp.Shadow.Visible = 0        ' 非選択はフラットへ戻す
                On Error GoTo 0
            End If
        End If
    Next shp
End Sub

' 指定バブルの本文を返す(無ければ"")。
Public Function BubbleTextOf(ByVal shapeName As String) As String
    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    BubbleTextOf = ws.Shapes(shapeName).TextFrame2.TextRange.Text
    On Error GoTo 0
End Function

' 最新AIバブル名を返す(無ければ"")。
Public Function LatestAiBubbleName() As String
    Dim ws As Worksheet
    Set ws = GetNexusSheet()
    If ws Is Nothing Then Exit Function

    Dim bestName As String: bestName = ""
    Dim bestBottom As Double: bestBottom = -1
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 9) = "nx_msg_a_" Then
            If shp.Top + shp.Height > bestBottom Then
                bestBottom = shp.Top + shp.Height
                bestName = shp.Name
            End If
        End If
    Next shp
    LatestAiBubbleName = bestName
End Function

' 最新バブルが見えるところまでスクロールする。Application.GoToはセル選択を
' 伴い保護シート+EnableSelectionと衝突するため、ScrollRowだけを動かす
' (FreezePanes下では下ペインのスクロール位置だけが変わる)。
Private Sub ScrollToBottom(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If Not (ThisWorkbook.ActiveSheet Is ws) Then Exit Sub
    On Error Resume Next
    Dim targetRow As Long
    targetRow = CLng(mChatBottom / 18) - 8
    If targetRow < 5 Then targetRow = 5
    ActiveWindow.ScrollRow = targetRow
    On Error GoTo 0
End Sub
