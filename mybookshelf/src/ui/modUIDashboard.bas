Attribute VB_Name = "modUIDashboard"
Option Explicit

' ============================================================================
' modUIDashboard - 「ダッシュボード」画面(個人統計+バッジ)の構築と描画
'   (MASTER_SPEC §7.6/§8.3/§9)
' ----------------------------------------------------------------------------
' 役割:
'   統計タイル(今月の質問数/🟢自己解決/取り戻した時間/冊数)+バッジ棚+
'   蔵書数バーを表示する「見るだけ」の画面。契約はEnsureLayout/
'   RenderDashboardの2本のみ(closed契約。ボタン等の追加公開関数は作らない)。
'
' 設計判断:
'   ・この画面自体にはユーザー操作(ボタン)が無い(§8.3のASCII図に操作は
'     描かれていない)。タブをクリックして開くだけで最新状態が見えることが
'     望ましいが、本ブックにはワークシートのActivateイベントに反応する
'     シートクラスモジュールが無い(担当ファイルはThisWorkbook.clsのみ)ため、
'     実際の更新タイミングはmodBoot.Boot(起動時)と、modUIMain.OnAskButton /
'     modUIShelf.OnAddFiles等の「統計が変わりうる操作の直後」に限られる
'     (それぞれのモジュールがRefreshBadgesAndDashboard相当の処理で
'     modStats.EvaluateBadges→modUIDashboard.RenderDashboardを呼ぶ)。
'     タブを開いた瞬間の自動更新が無い点は既知の制約としてWave3に申し送る。
'   ・「今月の質問数」: my_statsのカウンタは累計のみで月別集計を持たない
'     (MASTER_SPEC §4のmy_stats列定義に月別キーが無い)。modStats.GetStat
'     の契約もLong一値を返すのみで月別問い合わせはできない。そのため、
'     usage_logシート(event="ask"の行)をこのモジュールが直接読み、
'     timestampの先頭7文字(yyyy-mm)が今月と一致する行数を数える方式にした
'     (modStatsの契約を勝手に拡張せず、UI側で「今月」を絞り込む)。
'   ・バッジの獲得日は my_stats の "badge:<id>" 行の value(取得日文字列)
'     を直接読む。modStats.GetStatはLongしか返せない契約のため、日付文字列
'     の取得にはこの関数を使えない(バッジ取得日はLongで表現できない)。
'     そのためmy_statsシートを直接読む(このモジュールはUI層であり、
'     シートを直接読むこと自体はR4の制約対象外)。
'   ・「蔵書数の推移棒」: my_statsにも履歴的なスナップショットは保持されて
'     いない(現在値の累計のみ)ため、真の時系列推移ではなく「現在の蔵書数」
'     を1本のREPTバーで表す簡易表示にした(データモデルの制約による意図的な
'     単純化。Wave3申し送り)。
' ============================================================================

Private Const REPT_CHAR As String = "■"
Private Const MAX_BAR_LEN As Long = 40

' ----------------------------------------------------------------------------
' EnsureLayout
' ----------------------------------------------------------------------------
Public Sub EnsureLayout()
    Dim ws As Worksheet
    Set ws = GetOrCreateDashSheet()
    If ws Is Nothing Then Exit Sub

    ' uiStep: modUIMain.EnsureLayoutと同じ考え方(2026-07-15 実機E0801対策)。
    Dim uiStep As String
    On Error GoTo Fail

    Application.ScreenUpdating = False

    uiStep = "既存ボタンの削除"
    RemoveManagedShapes ws
    uiStep = "セルのクリア"
    ws.Cells.Clear

    uiStep = "既定フォント設定"
    ws.Cells.Font.Name = "游ゴシック"
    ws.Cells.Font.Size = 11

    ws.Columns("A:H").ColumnWidth = 12

    uiStep = "タイトル行"
    With ws.Range("A1:H1")
        .Merge
        .Value = "📊 ダッシュボード"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = 2039071   ' RGB(31,78,120)
        .Font.Color = 16777215      ' RGB(255,255,255)
        .HorizontalAlignment = -4108   ' xlCenter
    End With
    ws.Rows("1").RowHeight = 24

    uiStep = "サブタイトル行"
    With ws.Range("A2:H2")
        .Merge
        .Value = "あなたの本棚とAI活用の記録"
        .Font.Size = 10
        .Font.Italic = True
        .HorizontalAlignment = -4108
    End With
    ws.Rows("2").RowHeight = 18

    uiStep = "統計タイル枠"
    ws.Rows("4:6").RowHeight = 20
    FormatTile ws, "A4:B6"
    FormatTile ws, "C4:D6"
    FormatTile ws, "E4:F6"
    FormatTile ws, "G4:H6"

    uiStep = "バッジ見出し"
    With ws.Range("A7:H7")
        .Merge
        .Value = "🏅 バッジ"
        .Font.Bold = True
        .Font.Size = 12
    End With
    ws.Rows("7").RowHeight = 20
    ws.Rows("8:15").RowHeight = 18

    uiStep = "育ちぐあい見出し"
    With ws.Range("A16:H16")
        .Merge
        .Value = "📈 本棚の育ちぐあい"
        .Font.Bold = True
        .Font.Size = 12
    End With
    ws.Rows("16").RowHeight = 20

    uiStep = "育ちぐあいバー枠"
    With ws.Range("A17:H18")
        .Merge
        .Font.Size = 10
        .Font.Name = "Consolas"
        .VerticalAlignment = -4108   ' xlCenter
    End With
    ws.Rows("17:18").RowHeight = 18

    Application.ScreenUpdating = True

    uiStep = "ダッシュボードの再描画(RenderDashboard)"
    RenderDashboard
    Exit Sub

Fail:
    Dim origNum As Long, origDesc As String
    origNum = Err.Number
    origDesc = Err.Description
    On Error Resume Next
    Application.ScreenUpdating = True
    On Error GoTo 0
    Err.Raise origNum, "modUIDashboard.EnsureLayout", "[" & uiStep & "] " & origDesc
End Sub

' ----------------------------------------------------------------------------
' RenderDashboard - 統計タイル+バッジ棚+REPT("■")棒グラフ。EvaluateBadges後に呼ばれる。
' ----------------------------------------------------------------------------
Public Sub RenderDashboard()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetDashSheet()
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    Dim askThisMonth As Long
    askThisMonth = MonthlyAskCount()

    Dim solveTotal As Long
    solveTotal = SafeGetStat("selfsolve_total")

    Dim savedMinutes As Long
    savedMinutes = SafeSavedMinutes()

    Dim shelfCount As Long
    shelfCount = SafeShelfCount()

    RenderTile ws, "A4:B6", CStr(askThisMonth), "今月の質問数"
    RenderTile ws, "C4:D6", CStr(solveTotal), "🟢 自己解決した回数"
    RenderTile ws, "E4:F6", FormatMinutes(savedMinutes), "取り戻した時間"
    RenderTile ws, "G4:H6", CStr(shelfCount), "本棚の資料数"

    RenderBadges ws

    Dim barLen As Long
    barLen = shelfCount
    If barLen > MAX_BAR_LEN Then barLen = MAX_BAR_LEN

    ws.Range("A17:H18").Value = "本棚の資料数: " & String$(barLen, REPT_CHAR) & "  (" & shelfCount & "件)"

    Application.ScreenUpdating = True
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

Private Sub FormatTile(ByVal ws As Worksheet, ByVal addr As String)
    With ws.Range(addr)
        .Merge
        .Interior.Color = 15987946   ' 薄い水色
        .Borders.LineStyle = 1
        .HorizontalAlignment = -4108   ' xlCenter
        .VerticalAlignment = -4108
        .WrapText = True
    End With
End Sub

Private Sub RenderTile(ByVal ws As Worksheet, ByVal addr As String, ByVal valueText As String, ByVal labelText As String)
    Dim t As String
    t = valueText & vbLf & labelText

    On Error Resume Next
    ws.Range(addr).Value = t
    On Error GoTo 0

    On Error Resume Next
    ws.Range(addr).Font.Size = 12
    On Error GoTo 0

    On Error Resume Next
    ws.Range(addr).Font.Bold = True
    On Error GoTo 0
End Sub

Private Sub RenderBadges(ByVal ws As Worksheet)
    Dim ids() As String
    Dim titles() As String
    Dim conditions() As String
    BadgeCatalog ids, titles, conditions

    Dim i As Long
    For i = 0 To UBound(ids)
        Dim r As Long
        r = 8 + i
        Dim dt As String
        dt = BadgeDate(ids(i))

        Dim line As String
        Dim earned As Boolean
        earned = (LenB(dt) > 0)
        If earned Then
            line = "🏅 " & titles(i) & " — 獲得しました(" & dt & ")"
        Else
            line = "🔒 " & titles(i) & " — " & conditions(i)
        End If

        With ws.Range("A" & r & ":H" & r)
            .Value = line
            .Font.Size = 10
            If earned Then
                .Font.Color = 0
                .Font.Bold = True
            Else
                .Font.Color = 10921638   ' RGB(166,166,166)
                .Font.Bold = False
            End If
        End With
    Next i
End Sub

' MASTER_SPEC §9のバッジ定義。
Private Sub BadgeCatalog(ByRef ids() As String, ByRef titles() As String, ByRef conditions() As String)
    ids = Split("first_ingest,shelf10,shelf30,first_pack_out,first_pack_in,solve10,solve50,streak7", ",")
    titles = Split("初めての取込,本棚10冊,本棚30冊,初パック共有,初パック取込,自己解決10件,自己解決50件,7日連続利用", ",")
    conditions = Split( _
        "資料を1つ本棚に追加すると獲得|" & _
        "資料を10冊集めると獲得|" & _
        "資料を30冊集めると獲得|" & _
        "資料をパックとして誰かに渡すと獲得|" & _
        "誰かのパックを取り込むと獲得|" & _
        "🟢解決したが10回になると獲得|" & _
        "🟢解決したが50回になると獲得|" & _
        "7日連続で使うと獲得", "|")
End Sub

Private Function BadgeDate(ByVal badgeId As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_STATS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim key As String
    key = "badge:" & badgeId

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim i As Long
    For i = 2 To lastRow
        If StrComp(CStr(ws.Cells(i, 1).Value), key, vbTextCompare) = 0 Then
            BadgeDate = Trim$(CStr(ws.Cells(i, 2).Value))
            Exit Function
        End If
    Next i
End Function

Private Function MonthlyAskCount() As Long
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_USAGE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim monthPrefix As String
    monthPrefix = Format$(Now, "yyyy-mm")

    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 2)).Value

    Dim n As Long
    n = 0
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If Left$(CStr(arr(i, 1)), 7) = monthPrefix Then
            If StrComp(CStr(arr(i, 2)), "ask", vbTextCompare) = 0 Then
                n = n + 1
            End If
        End If
    Next i
    MonthlyAskCount = n
End Function

Private Function SafeGetStat(ByVal key As String) As Long
    On Error Resume Next
    SafeGetStat = modStats.GetStat(key)
    On Error GoTo 0
End Function

Private Function SafeSavedMinutes() As Long
    On Error Resume Next
    SafeSavedMinutes = modStats.SavedMinutesEstimate()
    On Error GoTo 0
End Function

Private Function SafeShelfCount() As Long
    Dim names() As String
    Dim stats() As String
    On Error Resume Next
    SafeShelfCount = modShelf.SourceList(names, stats)
    On Error GoTo 0
End Function

Private Function FormatMinutes(ByVal minutes As Long) As String
    If minutes < 60 Then
        FormatMinutes = minutes & "分"
    Else
        Dim h As Long
        h = minutes \ 60
        Dim m As Long
        m = minutes Mod 60
        If m = 0 Then
            FormatMinutes = h & "時間"
        Else
            FormatMinutes = h & "時間" & m & "分"
        End If
    End If
End Function

Private Function GetDashSheet() As Worksheet
    On Error Resume Next
    Set GetDashSheet = ThisWorkbook.Worksheets(modAppDef.SH_DASH)
    On Error GoTo 0
End Function

Private Function GetOrCreateDashSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_DASH)
    On Error GoTo 0
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = modAppDef.SH_DASH
        On Error GoTo 0
    End If
    Set GetOrCreateDashSheet = ws
    Exit Function
Fail:
    Set GetOrCreateDashSheet = Nothing
End Function

Private Sub RemoveManagedShapes(ByVal ws As Worksheet)
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
