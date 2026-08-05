Attribute VB_Name = "modViewport"
Option Explicit

' ============================================================================
' modViewport - 画面ごとの「見せてよい範囲」を Worksheet.ScrollArea として与える
'               小さな共通部品(2026-08-05 R18-3b・実機第5報②)。
' ----------------------------------------------------------------------------
' なぜ要るのか:
'   実機で「どの画面も右にも下にも無限にスクロールでき、真っ白な余白が延々と
'   続く」と報告された。主因は各画面の構築Subが ws.Cells(=A1:XFD1048576)へ
'   フォント/背景を当てていたこと(R18-3aで是正済み)だが、それだけでは
'   「利用者が行き先の無い場所へ行ける」状態は残る。Excelには「ここから先へは
'   スクロールさせない」という一行の宣言(Worksheet.ScrollArea)があるので、
'   各画面の構築Subの末尾でそれを宣言する。
'
' 設計の鉄則:
'   ・ScrollArea はブックに保存されないセッション限定のプロパティ。本アプリは
'     画面を出すたびに構築Subを必ず通る(冪等な全再構築)ので、そこに置けば
'     再適用の配線は要らない(調査agent2 §2.1/§2.2)。
'   ・範囲は必ず「実際に描いた最大到達点+余白」で決める。きつく締めると
'     ボタンが境界の外に取り残され、憲章§3-1「押せるものは必ず反応する」を
'     破る(見えない=押せない)。迷ったら広い側へ倒す。
'   ・失敗しても画面を落とさない(範囲文字列が不正でも、スクロールが自由な
'     ままになるだけで機能は失われない)。全てOn Error Resume Next配下に置く。
'
' 依存: Worksheet/Range のみ(他モジュールを呼ばない=どの画面からでも安全)。
' ============================================================================

' ----------------------------------------------------------------------------
' ApplyScrollBound - ws.ScrollArea を設定する。boundAddr が空なら制限を外す。
' ----------------------------------------------------------------------------
'   例) modViewport.ApplyScrollBound ws, "A1:L60"
'   ScrollArea はシート保護下でも設定できるが、順序の事故(保護→選択不能→
'   FreezePanes失敗)を避けるため、呼び出し側は「幾何とShapeを置き終えた直後・
'   保護をかける前」に呼ぶこと。
Public Sub ApplyScrollBound(ByVal ws As Worksheet, ByVal boundAddr As String)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.ScrollArea = boundAddr
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' BoundFor - 実測の右下端(pt)を含む最小の "A1:<列><行>" を組み立てる。
' ----------------------------------------------------------------------------
'   rightX/bottomY : 描いた内容の右端・下端(pt。余白を足した値を渡す)
'   maxCol/maxRow  : そのシートで許す上限(列幅・行高が壊れていても暴走しない)
'   列幅・行高はフォントやDPIで変わり、ptから列番号を机上計算することはできない
'   (「9文字幅=何pt」は端末依存)。実際のセル幾何(Left/Width, Top/Height)を
'   読んで決める ―― 本アプリがShape座標を必ずセル幾何から出すのと同じ理由。
Public Function BoundFor(ByVal ws As Worksheet, ByVal rightX As Double, _
                         ByVal bottomY As Double, ByVal maxCol As Long, _
                         ByVal maxRow As Long) As String
    If ws Is Nothing Then Exit Function
    If maxCol < 1 Then maxCol = 1
    If maxRow < 1 Then maxRow = 1

    Dim lastCol As Long: lastCol = maxCol
    Dim lastRow As Long: lastRow = maxRow
    Dim i As Long
    On Error Resume Next
    For i = 1 To maxCol
        If ws.Cells(1, i).Left + ws.Cells(1, i).Width >= rightX Then
            lastCol = i
            Exit For
        End If
    Next i
    For i = 1 To maxRow
        If ws.Cells(i, 1).Top + ws.Cells(i, 1).Height >= bottomY Then
            lastRow = i
            Exit For
        End If
    Next i
    On Error GoTo 0

    BoundFor = "A1:" & ColLetter(lastCol) & lastRow
End Function

' ----------------------------------------------------------------------------
' ColLetter - 列番号(1始まり)を列名へ。1→"A" / 26→"Z" / 27→"AA"。
' ----------------------------------------------------------------------------
'   Excelに聞けば(Cells(1,n).Address)取れるが、範囲文字列の組み立てのたびに
'   COMを往復させたくないうえ、ここだけは実行テストで固定できる純粋な算数
'   なので自前で持つ。範囲外(0以下)は "A" に丸める(壊れた値でRangeを
'   落とさない=呼び出し側の防御を1つ減らす)。
Public Function ColLetter(ByVal colIdx As Long) As String
    Dim n As Long: n = colIdx
    If n < 1 Then n = 1
    Dim s As String
    Do While n > 0
        Dim r As Long: r = (n - 1) Mod 26
        s = Chr$(65 + r) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function
