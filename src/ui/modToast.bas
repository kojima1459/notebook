Attribute VB_Name = "modToast"
Option Explicit

' ============================================================================
' modToast - Toast通知("nx_toast")の描画・待機・自動消去だけを持つモジュール
'   (2026-09-08 R41 B1)。
' ----------------------------------------------------------------------------
' なぜ modSkin から出したか(憲章§4-6):
'   modSkin が残27字(上限30,000まで残り27字)で、waitless Toastへの
'   OnTime自動消去(下記ArmSweep等)が1行も入らなかった。前例は modSkin→
'   modProgressBar(R18-1a)・modUI→modSkin(R11-F1)。移設は事実を保った
'   忠実な移設で、modSkin.ShowToast は同じシグネチャのまま
'   modToast.ShowToast への1行委譲になる(呼び出し元は全て
'   modSkin.ShowToast のままで変更不要)。
'
' 役割:
'   ・ShowToast: MsgBoxの代替(非ブロッキング通知)。画面上部中央に細長Shapeを
'     出し、待って消す(既定)か、すぐ戻る(waitless)。
'   ・waitless(R11-H Med4裁定当時)は「描いたらすぐ戻る」だけで自動消去の
'     仕組みが無く、次のShowToast/PaintProgressの掃除に任せていたため、
'     二度と呼ばれない画面では消えないまま残っていた(R41裁定B1の原因)。
'     ここへ Application.OnTime による自動消去(ArmSweep/CancelSweep/
'     ToastSweepTick)を足す。作法は modViewport.ScheduleRefit/CancelRefit/
'     RefitProcName と同型(予約は1本・ブック名修飾・失敗は非致命として握る)。
'   ・CancelSweep は modUI.RestoreExcelUI(Auto_Close 経路)が
'     modViewport.CancelRefit の直後に必ず呼ぶ(残したまま閉じると
'     時刻到来でブックが開き直る事故と同型)。
' ============================================================================

' モジュールレベル宣言はプロシージャより前に置く(実機VBAの制約)。
Private mSweepAt As Date        ' 予約済み OnTime の時刻(未予約時は無効)
Private mSweepArmed As Boolean  ' 予約が生きているか(mSweepAtの有効フラグ)

' ----------------------------------------------------------------------------
' ShowToast - MsgBoxの代替(非ブロッキング通知)。画面上部中央に細長Shapeを出し、
'   短時間表示して自動で消す。kind: "success"/"error"/"info"。
' ----------------------------------------------------------------------------
' waitless(R11-H Med4): Trueなら1.1秒の待機・削除をせず描いたらすぐ戻る
' (「一言返すだけ」は待たせること自体が害)。2026-09-08(R41 B1): 待って消す
' 側と同じく Application.OnTime による自動消去(ArmSweep)を足した。既定は
' 待って消す。
Public Sub ShowToast(ByVal message As String, Optional ByVal kind As String = "info", _
                     Optional ByVal waitless As Boolean = False)
    On Error Resume Next
    ' 別ブック誤爆ガード: 他の業務Excelを見ている間にToastを描くと、他人の
    ' ブックへShapeを生成して業務データを汚す。自ブックがアクティブな時だけ描く。
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    ' 前の予約を必ず解く(1本の原則。孤児Sweepを作らない)。ガードの【後】に置く:
    ' 描かずに戻る経路で前の予約だけ解くと、残っている前のトーストが二度と
    ' 自動で消えなくなる(レビュー1周目 MINOR-5)。
    CancelSweep
    ws.Shapes("nx_toast").Delete   ' 前のToastが残っていれば消す(孤児防止)

    Dim toastW As Double: toastW = 380
    Dim leftPos As Double: leftPos = 280
    Dim topPos As Double: topPos = 96
    leftPos = ActiveWindow.VisibleRange.Left + (ActiveWindow.VisibleRange.Width - toastW) / 2
    topPos = ActiveWindow.VisibleRange.Top + 92

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, leftPos, topPos, toastW, 34)   ' 5=角丸四角(高さはFitToastHeightで確定)
    shp.Name = "nx_toast"
    shp.Adjustments(1) = 0.35
    shp.Line.Visible = 0
    shp.Placement = 3   ' xlFreeFloating

    ' 種別アイコン(視覚的認知スピード): 成功✅ / 注意⚠️ / 情報💡
    Dim bg As Long, fg As Long, icon As String
    Select Case LCase$(kind)
        ' R47: 白字 on RGB(0,168,89) は 3.11:1 で AA(4.5)未達。10.5pt の通常字なので
        ' large-text 例外も効かない。modSkin が2度「ACCENT は白字と組まない」と
        ' 明文化して各所を直したのに、一番よく見るトーストだけ取り残されていた。
        ' PRIMARY #01675B(白字 6.79:1)へ。
        Case "success": bg = RGB(1, 103, 91): fg = RGB(255, 255, 255): icon = ChrW(&H2705)
        Case "error":   bg = RGB(220, 38, 38):  fg = RGB(255, 255, 255): icon = modEmj.Warn()
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
    modChrome.FitToastHeight shp   ' R30 W2-1: 実測AutoSizeへ転換(icon込み計測)

    modSkin.ApplySoftShadow shp
    shp.ZOrder 0   ' msoBringToFront
    DoEvents
    If waitless Then
        ArmSweep SweepMsFor(message)   ' R41 Fix M3: 待たせない代わりに寿命は読む側の余裕を持つ
        On Error GoTo 0
        Exit Sub
    End If
    ToastWait modChrome.ToastWaitMsFor(message)   ' R29 W2-3: 文字量に応じ可変(3,000-9,000ms)
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
' ArmSweep - waitless Toastの自動消去を Application.OnTime で1本予約する
'   (2026-09-08 R41 B1)。作法は modViewport.ScheduleRefit と同型。
' ----------------------------------------------------------------------------
Public Sub ArmSweep(ByVal ms As Long)
    On Error Resume Next
    CancelSweep
    Dim t As Date: t = Now + CeilSec(ms) / 86400#
    Application.OnTime EarliestTime:=t, Procedure:=SweepProcName()
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub                            ' 予約できないのは非致命(次のShowToastが掃除に来る)
    End If
    mSweepAt = t
    mSweepArmed = True
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' SweepIfDue - 「孤児のトーストだけ掃除する」窓口(R46 B-5)。
' ----------------------------------------------------------------------------
' 実機報告「作業用Excelの案内トーストが一瞬で消える」の対処。
' 原因は modProgressBar.PaintProgress の冒頭が nx_toast を【無条件で】
' Delete していたこと。その理由コメントは「waitless のToastは自分では
' 消えないので」だったが、R41 で ArmSweep を入れた時点でこの前提は失効して
' いる。失効した掃除口を引退させなかったため、古い仕掛けが新しい仕掛けを
' 殺していた。取込バナーは BANNER_SEC=1 秒ごとに塗り直され、そもそも
' 「作業用Excel」ボタンは取込バナーが出ている間しか存在しないので、
' 押した直後のトーストは最大1秒で確実に消される(寿命12,134msの計算自体は
' 正しく、そこまで生き残れなかった)。
' 予約が生きていて期限前なら何もしない(自前の OnTime に回収させる)。
' それ以外(予約が無い=孤児 / 期限切れ)は従来どおり消す。
' 注意: ShowToast 冒頭の Delete は【無条件のまま】にすること。あちらを
' この判定にすると、新しいトーストが古いトーストを消せず二重表示になる。
Public Sub SweepIfDue(ByVal ws As Worksheet)
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    If mSweepArmed Then
        If Now < mSweepAt Then Exit Sub
    End If
    ws.Shapes("nx_toast").Delete
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' SweepMsFor - waitless トーストの寿命(ms)。純ロジック(R41 Fix M3・レビュー
'   1周目 MAJOR-3)。待って消す側の ToastWaitMsFor(読了速度=全角15字/秒・
'   3,000〜9,000ms)は「利用者を止めている時間」なので短めに切ってあるが、
'   waitless は誰も止めておらず、利用者が別の操作(作業用Excelへ移る等)の
'   途中で目に入る前提なので、同じ文で2倍の余裕を持たせる(6,000〜15,000ms)。
'   R19H FA-8「作業用Excelの案内(78字)は消えないまま残す」は、この寿命
'   (約10秒)で置き換える裁定(R41 spec §5)。
' ----------------------------------------------------------------------------
Public Function SweepMsFor(ByVal message As String) As Long
    Const SWEEP_MAX_MS As Long = 15000
    Dim ms As Long: ms = modChrome.ToastWaitMsFor(message) * 2
    If ms > SWEEP_MAX_MS Then ms = SWEEP_MAX_MS
    SweepMsFor = ms
End Function

' ms(ミリ秒)を切り上げ秒数へ(OnTimeの時刻粒度は秒)。
Private Function CeilSec(ByVal ms As Long) As Double
    CeilSec = -Int(-(ms / 1000#))
End Function

' ----------------------------------------------------------------------------
' CancelSweep - 未消化の予約を取り消す。modUI.RestoreExcelUI
'   (modViewport.CancelRefitの直後・Auto_Close経路で必ず通る)から呼ぶ。
'   予約が無い/既に発火済みでも安全(1004は握る)。
' ----------------------------------------------------------------------------
Public Sub CancelSweep()
    On Error Resume Next
    If mSweepArmed Then
        Application.OnTime EarliestTime:=mSweepAt, Procedure:=SweepProcName(), Schedule:=False
        Err.Clear
    End If
    mSweepArmed = False
    On Error GoTo 0
End Sub

' SweepProcName - 予約/解除で必ず同じ文字列になるように1箇所で組む。
Private Function SweepProcName() As String
    SweepProcName = "'" & ThisWorkbook.Name & "'!modToast.ToastSweepTick"
End Function

' ----------------------------------------------------------------------------
' ToastSweepTick - OnTime のコールバック本体。Public 必須(OnTime は
'   Application.Run と同じ遅延バインドで Private を呼べない)。全シートの
'   nx_toast を消す(ClearProgressと同じ全シート横断=孤児を作らない。他ブックが
'   前面でも自ブックのシートを名前で消すだけなので誤爆しない)。
' ----------------------------------------------------------------------------
Public Sub ToastSweepTick()
    On Error Resume Next
    mSweepArmed = False
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Shapes("nx_toast").Delete
    Next ws
    On Error GoTo 0
End Sub
