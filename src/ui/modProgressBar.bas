Attribute VB_Name = "modProgressBar"
Option Explicit

' ============================================================================
' modProgressBar - 進捗バナー("nx_progress")と、その中の2つのボタン
'   (■中断 / 作業用Excel)の描画・撤去だけを持つモジュール(2026-08-05 R18-1a)。
' ----------------------------------------------------------------------------
' なぜ modSkin から出したか(憲章§4-6):
'   modSkin が27,990字(上限30,000まで残り10字)で、R18-1b〜1eの修正が1行も
'   入らなかった。前例は modUI→modSkin のテーマ移設、optOcrCore→optOcrEta。
'   移設は事実を保った忠実な移設で、呼び出し元は modUIMain.ShowProgress /
'   HideProgress と modShelfBatch.ShowIngestBanner の3行だけ書き換えた。
'   Shape名の定数は Public Const にする(Private Const はモジュールを跨いで
'   参照できず、LibreOffice の構文検証も通らない)。
'
' 役割(移設前からの事実):
'   ・ShowToastと違い待機ゼロ(ファイル数×1.1秒の純増を避ける)。更新後
'     DoEvents 1回で再描画。別ブック表示中は何もしない(誤爆ガード)。
'     シート切替時は旧シートのShapeを消して描く。
'   ・cancellable(R15-FixA FA-6): Trueで「中断」を添える(従来は全処理で生え、
'     押しても止まらず壊れたボタン同然=憲章§3-1。今は modShelfBatch.
'     ShowIngestBanner だけがTrueを渡す)。Falseでも既存ボタンは消さない
'     (実況中に消えると押したい瞬間に押せない)。R16-2a: Trueなら
'     modWorkExcel.PaintWorkButton が■中断の左隣へ作業用Excelも添える
'     (座標・生成は同モジュール側。ここは座標計算を持たない)。寿命はバナーと
'     同じで ClearProgress が消す。
' ============================================================================

' R10-5: PaintProgress/ClearProgressが進捗バナーの直前描画シート名を控える
' (モジュールレベル宣言はプロシージャより前に置く。実機VBAの制約)。
Private mProgressSheetName As String

Private Const PROGRESS_NAME As String = "nx_progress"
' R15-6a: 中断ボタンのShape名("nx_progress"と対で出し消す)。
Public Const PROGRESS_CANCEL_NAME As String = "nx_progress_cancel"
' R16-2a: 作業用Excelボタンの名前。ClearProgressの削除対象へ同列で足す。
Public Const PROGRESS_WORK_NAME As String = "nx_progress_work"

' R18-1b: バナー幅の上限。可視幅いっぱいまでは広げない(全画面表示の
' ワイドモニタで1,600pt級の帯になると、かえって文の視線移動が長くなる)。
Private Const BAR_MAX_W As Double = 760
' 右端の余白。cancellable のときは ■中断76 + 間隔6 + 作業用Excel96 + 外側12。
Private Const MARGIN_R_PLAIN As Double = 16
Private Const MARGIN_R_BUTTONS As Double = 190

' ----------------------------------------------------------------------------
' BarWidthFor - バナー幅(pt)の決定。純ロジック(ゴールデンテスト対象)。
' ----------------------------------------------------------------------------
' 2026-08-05(R18-1b): 従来は固定380ptだった。本文の可視幅は
' 380 − 左余白16 − 右余白190 = 174pt しかなく、OCRの実況文(約363pt)は
' 2〜3行に折り返して高さ30ptのピルから溢れ、地色との対比1.05:1で完全に
' 読めなかった(実機第5報①)。#30恒久対策(R11-B)の共通部品
' modUIMain.ViewportWidth はとうに在ったのに、この1箇所だけ適用漏れしていた
' (憲章§4-5)。viewport−16 と 760 の小さい方にすれば本文可視幅は
' 554pt以上になり、現行の文面が1行に収まる。狭い窓でも画面外へはみ出さない。
Public Function BarWidthFor(ByVal viewportW As Double) As Double
    Dim w As Double: w = viewportW - 16
    If w > BAR_MAX_W Then w = BAR_MAX_W
    If w < 300 Then w = 300
    BarWidthFor = w
End Function

' ----------------------------------------------------------------------------
' BarHeightFor - バナー高さ(pt)の決定。純ロジック(ゴールデンテスト対象)。
' ----------------------------------------------------------------------------
' 2026-08-05(R18H FA-2 / B-H1・A-L11): R18-1b で幅を viewport 連動にしたが、
' 高さは30pt(1行)固定のままだった。1,024pt級の狭い窓では本文可視幅が
' 800pt前後まで縮み、OCRの実況文(推定約363pt)より長い文面 ―― 例えば
' 残り時間と終了目安を両方載せた行 ―― が入ったとたん2行に折り返して
' 30ptのピルから溢れ、R18-1b 以前とまったく同じ「読めない実況」に戻る。
' 幅だけを直したのは対症療法だったので、文面が入らないと分かった時点で
' 高さを2行ぶんへ広げる(縮む方向にも追随する=次の短い文で30ptへ戻る)。
'
' 幅の見積りは BarWidthFor と同じ流儀の純関数(TextWidthPt)。実測せず
' 文字数×係数で近似し、【過大側へ倒す】: 早めに2行化するのは「1行ぶん高い
' 帯が出る」だけだが、遅れると文字が切れて読めない(憲章§3-2)。
' bodyW は呼び出し側が barW から左右余白を引いた実際の本文可視幅。
Public Function BarHeightFor(ByVal message As String, ByVal bodyW As Double) As Double
    BarHeightFor = 30
    If bodyW <= 0 Then Exit Function
    If TextWidthPt(message) > bodyW Then BarHeightFor = 46
End Function

' 文字列の表示幅(pt)の近似。Yu Gothic UI 10pt を前提に、半角(ASCII+半角
' カナ)を5.5pt・それ以外(全角)を10.5ptで数える。実フォントの前進幅より
' やや大きめの係数を使うのは、上の「過大側へ倒す」判断そのもの。
Private Function TextWidthPt(ByVal s As String) As Double
    Dim w As Double, i As Long, c As Long
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536
        If c < &H100 Or (c >= &HFF61 And c <= &HFF9F) Then
            w = w + 5.5
        Else
            w = w + 10.5
        End If
    Next i
    TextWidthPt = w
End Function

Public Sub PaintProgress(ByVal message As String, Optional ByVal cancellable As Boolean = False)
    On Error Resume Next
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    ' 2026-07-31(R11-H Med4): waitless のToastは自分では消えないので、進捗
    ' バナーを出すここでも掃除する(ShowToastの先頭と同じ役目)。
    ws.Shapes("nx_toast").Delete

    If LenB(mProgressSheetName) > 0 And mProgressSheetName <> ws.Name Then
        Dim wsOld As Worksheet
        Set wsOld = ThisWorkbook.Worksheets(mProgressSheetName)
        If Not wsOld Is Nothing Then wsOld.Shapes(PROGRESS_NAME).Delete
        If Not wsOld Is Nothing Then wsOld.Shapes(PROGRESS_CANCEL_NAME).Delete
        ' R16H FA-6: 作業用Excelも対で消す(消し忘れると別シートに取り残される)。
        If Not wsOld Is Nothing Then wsOld.Shapes(PROGRESS_WORK_NAME).Delete
    End If
    mProgressSheetName = ws.Name

    Dim barW As Double: barW = BarWidthFor(modUIMain.ViewportWidth())
    Dim leftPos As Double, topPos As Double
    leftPos = ActiveWindow.VisibleRange.Left + (ActiveWindow.VisibleRange.Width - barW) / 2
    If leftPos < ActiveWindow.VisibleRange.Left Then leftPos = ActiveWindow.VisibleRange.Left
    ' R10c(M4): トースト(nx_toast)も同じ+92に出るため、取込完了の瞬間だけ
    ' 2枚が重なり文字が読めなかった。バナーはトースト(高さ34)の下へ
    ' (92+34+余白4=130)。
    topPos = ActiveWindow.VisibleRange.Top + 130

    ' R18H FA-2(ii): 右余白は cancellable ではなく【ボタンが実際に出ているか】で
    ' 決める。cancellable=False の更新(関所・ベクトル化・検索キャッシュ)は
    ' ボタンを消さないので、そこで16ptへ戻すと本文がボタンの下へ潜って読めなく
    ' なっていた(ZOrderで前面へ戻しても、文字が隠れる事実は変わらない)。
    Dim marginR As Double: marginR = MARGIN_R_PLAIN
    If cancellable Then marginR = MARGIN_R_BUTTONS
    If HasShape(ws, PROGRESS_CANCEL_NAME) Or HasShape(ws, PROGRESS_WORK_NAME) Then
        marginR = MARGIN_R_BUTTONS
    End If
    ' R18H FA-2(i): 文面が本文可視幅に収まらなければ2行ぶんへ広げる(生成時も
    ' 更新時も。短い文面へ戻れば30ptへ縮む)。
    Dim barH As Double: barH = BarHeightFor(message, barW - 16 - marginR)

    Dim shp As Shape
    Set shp = ws.Shapes(PROGRESS_NAME)
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(5, leftPos, topPos, barW, barH)   ' 5=角丸四角
        shp.Name = PROGRESS_NAME
        shp.Adjustments(1) = 0.3
        shp.Line.Visible = 0
        shp.Placement = 3   ' xlFreeFloating
        shp.Fill.ForeColor.RGB = RGB(30, 41, 59)
        With shp.TextFrame2
            .WordWrap = -1
            .MarginLeft = 16: .MarginTop = 4: .MarginBottom = 4   ' 右余白は下で
            .TextRange.Font.Name = "Yu Gothic UI"
            .TextRange.Font.Size = 10
            .TextRange.ParagraphFormat.Alignment = 2   ' 中央
            .VerticalAnchor = 3
        End With
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(248, 250, 252)
        modSkin.ApplySoftShadow shp
    Else
        shp.Left = leftPos
        shp.Top = topPos
        ' R18-1b: 窓の大きさは取込の途中でも変わる。幅も毎回入れ直す
        ' (生成時だけだと最初の1回の窓幅に焼き付く)。R18H FA-2: 高さも同様。
        shp.Width = barW
        shp.Height = barH
    End If
    ' R15-FixB(FB-6): 中断ボタンはバナー【内側】右端へ(従来は右外
    ' leftPos+barW+6で低解像度・小窓では画面外へはみ出し押せなかった。
    ' 内側なら可視領域中央のバナーと必ず一緒に見える)。右余白190=従来88+
    ' 作業用Excelボタン96+間隔6(R16H FA-5)。毎回入れ直すのは cancellable が
    ' 呼びごとに変わり得るため。
    shp.TextFrame2.MarginRight = marginR
    shp.TextFrame2.TextRange.Text = message
    shp.ZOrder 0   ' msoBringToFront

    If cancellable Then
        ' R18H FA-2(i): ボタンのY基準は【バナーの下段】。1行(30pt)のときは
        ' 従来と同じ topPos で、2行(46pt)へ広がったときだけ下へ寄る
        ' (上寄せのままだと文の2行目がボタンの下に潜る)。
        Dim btnTop As Double: btnTop = topPos + barH - 30
        PaintCancelButton ws, leftPos + barW - 82, btnTop
        modWorkExcel.PaintWorkButton ws, leftPos + barW - 82, btnTop
        ' R18-1d: 取込バナーが出ている間は砂時計をやめる。busy の合図は
        ' バナーが担っており、砂時計は「押しても無駄」に見えるだけで
        ' ■中断と作業用Excelという【押してほしいボタン】を殺していた。
        ' 復元は modUiLock.Leave が取込完了時にどのみち行う(ClearProgress
        ' では触らない=表示の後始末で他機能のカーソルを奪わない)。
        Application.Cursor = -4143   ' xlDefault
    End If

    ' R18-1c: バナーを最前面にした【あと】で、2つのボタンを必ず前面へ戻す。
    ' cancellable=False の PaintProgress(関所・ベクトル化・検索キャッシュ)は
    ' ボタンを消さずにバナーだけを ZOrder 0 していたため、不透明なバナーの下へ
    ' 両ボタンが埋まり、クリックがバナー(OnActionなし)に吸われて完全に無反応
    ' になっていた(実機第5報①の主犯。押すほど埋まる自己増悪だった)。
    ' cancellable の値に関わらず、在るものは必ず前へ出す(憲章§3-1)。
    BringToFront ws, PROGRESS_WORK_NAME
    BringToFront ws, PROGRESS_CANCEL_NAME
    DoEvents
    On Error GoTo 0
End Sub

' 指定名のShapeが在るか(R18H FA-2(ii)。無ければ False。生成はしない)。
Private Function HasShape(ByVal ws As Worksheet, ByVal shapeName As String) As Boolean
    On Error Resume Next
    Dim s As Shape
    Set s = ws.Shapes(shapeName)
    HasShape = Not (s Is Nothing)
    On Error GoTo 0
End Function

' 指定名のShapeが在れば最前面へ。無ければ何もしない(生成はしない)。
Private Sub BringToFront(ByVal ws As Worksheet, ByVal shapeName As String)
    On Error Resume Next
    Dim s As Shape
    Set s = ws.Shapes(shapeName)
    If Not s Is Nothing Then s.ZOrder 0
    On Error GoTo 0
End Sub

' PaintCancelButton - 進捗バナー内側右端の「中断」(R15-6a・RC3。位置は
'   R15-FixB FB-6 で外側へ→内側へ)。85〜127分の取込を止める手段が無く強制
'   終了しか無かった。止まるのは今の頁の後(OnCancelIngestは印を立てるだけ)。
'   絵文字は使わない(CP932・R13-L6)。
Private Sub PaintCancelButton(ByVal ws As Worksheet, ByVal leftPos As Double, _
                              ByVal topPos As Double)
    On Error Resume Next
    Dim btn As Shape
    Set btn = ws.Shapes(PROGRESS_CANCEL_NAME)
    If btn Is Nothing Then
        Set btn = ws.Shapes.AddShape(5, leftPos, topPos, 76, 30)
        btn.Name = PROGRESS_CANCEL_NAME
        btn.Adjustments(1) = 0.3: btn.Line.Visible = 0: btn.Placement = 3
        btn.Fill.ForeColor.RGB = RGB(120, 32, 40)
        With btn.TextFrame2
            .TextRange.Font.Name = "Yu Gothic UI": .TextRange.Font.Size = 10
            .TextRange.ParagraphFormat.Alignment = 2: .VerticalAnchor = 3
        End With
        btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 241, 242)
        btn.TextFrame2.TextRange.Text = ChrW(&H25A0) & "中断"
        btn.OnAction = "modShelfBatch.OnCancelIngest"
    Else
        btn.Left = leftPos: btn.Top = topPos
    End If
    btn.ZOrder 0
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' ClearProgress - バナーと2つのボタンを消す。
' ----------------------------------------------------------------------------
' 2026-08-05(R18-1e): 直前に描いた1枚(mProgressSheetName)だけでなく、全ての
' ワークシートを横断して消す。従来は「描いたシートを覚えている」ことが前提で、
' 強制終了やVBAリセットで変数が飛ぶとShapeだけがブックに保存されて残り、
' 次に開いたときに押しても何も起きない黒い帯として居座っていた(実機第5報⑪)。
' 消す対象は4つのShape名だけなので、シート枚数ぶんの On Error Resume Next 削除で
' 十分に安く、覚えているかどうかに依存しない。
' カーソル(砂時計)はここでは触らない: 復元は modUiLock.Leave の仕事で、
' 表示の後始末が他機能の状態を書き換えてはならない(憲章§4-4)。
' 2026-08-06(R19-3b): "nx_toast"(modSkin.ShowToastのトースト)も同じ焼き付き
' バグクラスなので対象へ加えた(実機第6報③)。modSkin側はPrivate定数を持たず
' 文字列リテラルで名付けているため、ここも同じリテラルで揃える。
Public Sub ClearProgress()
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Shapes(PROGRESS_NAME).Delete
        ws.Shapes(PROGRESS_CANCEL_NAME).Delete
        ws.Shapes(PROGRESS_WORK_NAME).Delete
        ws.Shapes("nx_toast").Delete
    Next ws
    mProgressSheetName = ""
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' SweepOrphans - 起動時・Hub再構築時の孤児バナー掃除(2026-08-05 R18-1e)。
' ----------------------------------------------------------------------------
' 前回のセッションが強制終了され、進捗バナーが出たままブックが保存されると、
' 次の起動でも黒い帯が残る(ClearProgress を呼ぶ機会が無いまま死んだため)。
' modHub.EnsureHubLayout から1行呼ぶ。取込中(modShelf.IsBusy)は【今出ている
' 本物】を消してしまうので何もしない。表示の掃除が取込を壊してはならないので
' 全体を On Error Resume Next で包む(憲章§4-4)。
Public Sub SweepOrphans()
    On Error Resume Next
    Dim busy As Boolean: busy = False
    busy = modShelf.IsBusy()
    If busy Then Exit Sub
    ClearProgress
    On Error GoTo 0
End Sub
