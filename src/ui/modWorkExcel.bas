Attribute VB_Name = "modWorkExcel"
Option Explicit

' ============================================================================
' modWorkExcel - 取込や入念な質問応答でこのExcelが「応答なし」になっている間、
'   別のExcelプロセスを開いて他の仕事ができるようにする一式(R16-2、2026-08-05)。
' ----------------------------------------------------------------------------
' なぜ要るのか:
'   ChatGPTV(画像PDFのOCR)・ChatGPT(通常のQ&A・入念モード)はいずれも待ち秒数を
'   指定できない同期呼び出しで、応答が返るまでこのブックのExcelプロセスは
'   他の操作を受け付けない(公式仕様。VBA側から制御不能。
'   docs/dev/RIBBON_API_CONFIRMED.md:29)。別プロセスのExcel(excel.exe /x)は
'   このブロックの影響を受けないため、取込や質問の完了を待つ間も別の資料の
'   確認・別の仕事ができるようにする。同一インスタンス内でDoEventsの隙間に
'   他ブックを編集する案内はMS公式が非推奨のため、別プロセス一本に統一する。
'
' 2つの配置(どちらも同じOnAction=OnOpenWorkExcelを指す):
'   (1) 取込バナー内、■中断の左隣(modSkin.PaintProgressのcancellable=True
'       分岐からPaintWorkButtonを1行呼ぶ)。
'   (2) アイドル時のヘルプカード常設導線(modHelpのカードへ1行追加)。
'
' 設計判断:
'   ・OnOpenWorkExcelはmodUiLock.Enter/BlockIfIngestingを意図して呼ばない
'     (modApp.OnAddDocsがmodUiLock.Enterを取込完了まで握るため、通せば必ず
'     弾かれ、取込中に一番使いたいこの機能が使えなくなる)。取込中でも動く
'     べきハンドラとしてtools/vba_lint.pyのONACTION_GUARD_ALLOWLISTに
'     理由付きで登録済み。
'   ・再入ガードはTimerベースの2秒デバウンス(DoEventsのキュー消化で1回の
'     クリックが複数回OnActionを発火させることがある連打対策)。待ちループは
'     しない=前回からの経過だけを見る瞬時判定。modUiLockのようなグローバル
'     ロックは使わない(他の処理と競合する理由が無い独立機能のため)。
'   ・失敗時はMsgBoxを出さない(取込バナー経由の呼び出しでモーダルを出すと
'     取込全体を止めてしまう=憲章§3-1)。Application.StatusBarへの短文表示+
'     modLog.LogError(新コードE0904)に留める。quietFail引数はこのための
'     分岐で、OnOpenWorkExcel(バナー内ボタンからも呼ばれる共有経路)からは
'     常にTrueで渡す。
'   ・DisableProcessWindowsGhosting(user32、引数なし・戻り値なし・表示系)は
'     本リポジトリ初のDeclare(docs/dev/EDGE_CASES.md §1.2に例外条項を追記
'     済み。32bit実機のみを前提とし、LOコンパイル通過を実験で確認済み)。
'     呼ぶとプロセス終了まで解除できない公式仕様のため、config
'     freeze_keep_banner(既定on)でオプトアウトできるようにしている。
' ============================================================================

#If VBA7 Then
    Private Declare PtrSafe Sub DisableProcessWindowsGhosting Lib "user32" ()
#Else
    Private Declare Sub DisableProcessWindowsGhosting Lib "user32" ()
#End If

Private mLastOpenAt As Double          ' Timer値。0=未実行(2秒デバウンス用)
' EnsureNoGhostingの1プロセス1回ガード。onで呼んだ後だけでなく【offと判定した
' ことも】覚える(R16H FB-1 / A-L10)。offのままだと1ページ・1段ごとに
' modConfig.GetBool がconfigシートを引き直す(取込では毎ページ通る経路)。
' 帰結: freeze_keep_banner の変更はExcelの再起動で反映される(docs/30 §11)。
Private mChecked As Boolean

Private Const WORK_BTN_NAME As String = "nx_progress_work"
Private Const WORK_BTN_W As Double = 96

' ----------------------------------------------------------------------------
' OnOpenWorkExcel - 「作業用Excel」ボタンの共通ハンドラ(R16-2a)。取込バナー内
'   ボタンとヘルプカードのボタンの両方が、この1本(引数なし)を指す。
'   modUiLock.Enter/BlockIfIngestingは意図して呼ばない(モジュール冒頭の
'   コメント参照)。2秒以内の連打は無視する(デバウンス)。
' ----------------------------------------------------------------------------
Public Sub OnOpenWorkExcel()
    On Error Resume Next   ' サーキットブレーカー(ARCHITECTURE §4)。失敗しても
                            ' 「作業用Excelが開かないだけ」に留め、他機能を道連れにしない。
    If Debounced() Then DoOpenWorkExcel True
    On Error GoTo 0
End Sub

' 2秒デバウンス。前回の成功呼び出しから2秒以上経っていたときだけTrueを返し、
' 同時に基準時刻を今に更新する。待ちループはしない(瞬時判定)。
Private Function Debounced() As Boolean
    Dim t As Double: t = Timer
    Dim elapsed As Double: elapsed = t - mLastOpenAt
    If elapsed < 0 Then elapsed = elapsed + 86400   ' 深夜0時のTimerロールオーバーガード
    If mLastOpenAt <> 0 And elapsed < 2# Then Exit Function
    mLastOpenAt = t
    Debounced = True
End Function

' 実処理。optGsProc.RunGsAsyncと同じEDR配慮パターン(WScript.Shellが端末
' ポリシーで塞がれている場合の失敗経路を必ず捕捉し、原因をログへ残す)。
' quietFail: Trueなら失敗してもMsgBoxを出さない(StatusBar+ログのみ)。
Private Sub DoOpenWorkExcel(ByVal quietFail As Boolean)
    Dim cmd As String
    cmd = modUtilText.BuildWorkExcelCmd(Application.Path)

    Dim wsh As Object
    Dim failNum As Long, failDesc As String
    On Error GoTo NoRun
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 1, False
    Set wsh = Nothing
    On Error GoTo 0

    On Error Resume Next
    Application.StatusBar = "作業用Excelを起動しました。そちらで仕事ができます"
    On Error GoTo 0
    Exit Sub

NoRun:
    failNum = Err.Number
    failDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、後始末中に起きたエラーは
    ' 呼び出し元へ飛んで本来の原因を上書きする。Resume でハンドラを抜けてから
    ' 後始末する(modUtilText.ReadTextFileUtf8等と同じ作法)。
    Resume WorkExcelFailCleanup
WorkExcelFailCleanup:
    On Error Resume Next
    Application.StatusBar = "作業用Excelを起動できませんでした。しばらくしてから再度お試しください。"
    modLog.LogError "E0904", "modWorkExcel.OnOpenWorkExcel", _
        "WScript.Shell起動失敗: " & failDesc, failNum
    If Not quietFail Then
        MsgBox "作業用Excelを起動できませんでした。" & vbLf & "(コード: E0904)", _
            vbExclamation, modAppDef.APP_NAME
    End If
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' PaintWorkButton - 進捗バナー内、■中断の左隣に作業用Excelボタンを添える
'   (R16-2a)。modSkin.PaintProgressのcancellable=True分岐から1行だけ呼ばれる。
'   座標は中断ボタンの位置(cancelLeftPos/cancelTopPos)を受け取り、その左へ
'   幅+間隔ぶんずらして自分の位置を決める(modSkin側は座標計算を持たない)。
'   削除はmodSkin.ClearProgress側にPROGRESS_WORK_NAMEとして同列に登録済み
'   (Shape名の文字列そのものはWORK_BTN_NAMEがこちら側の一次情報)。
' ----------------------------------------------------------------------------
Public Sub PaintWorkButton(ByVal ws As Worksheet, ByVal cancelLeftPos As Double, _
                           ByVal cancelTopPos As Double)
    On Error Resume Next
    Dim leftPos As Double: leftPos = cancelLeftPos - WORK_BTN_W - 6

    Dim btn As Shape
    Set btn = ws.Shapes(WORK_BTN_NAME)
    If btn Is Nothing Then
        Set btn = ws.Shapes.AddShape(5, leftPos, cancelTopPos, WORK_BTN_W, 30)
        btn.Name = WORK_BTN_NAME
        btn.Adjustments(1) = 0.3: btn.Line.Visible = 0: btn.Placement = 3
        btn.Fill.ForeColor.RGB = RGB(30, 74, 110)
        With btn.TextFrame2
            .WordWrap = -1
            .TextRange.Font.Name = "Yu Gothic UI": .TextRange.Font.Size = 9.5
            .TextRange.ParagraphFormat.Alignment = 2: .VerticalAnchor = 3
            .MarginLeft = 2: .MarginRight = 2
        End With
        btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        btn.TextFrame2.TextRange.Text = ChrW(&HD83D) & ChrW(&HDDD4) & "作業用Excel"
        btn.OnAction = "modWorkExcel.OnOpenWorkExcel"
    Else
        btn.Left = leftPos: btn.Top = cancelTopPos
    End If
    btn.ZOrder 0
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' EnsureNoGhosting - DWMの「応答なし」白画面化(ゴースト化)を抑止する(R16-2b)。
'   config freeze_keep_banner(既定on)がオフなら何もしない。プロセス生存中に
'   最大1回だけAPIを呼ぶ(呼ぶとプロセス終了まで解除不能な公式仕様のため、
'   複数回呼んでも1回目以降は無意味。フラグで無駄な再試行そのものを避ける)。
'   判定そのものも1回きり(mChecked)。offの人はconfigの読み直しだけが毎回
'   残るため、設定変更はExcelの再起動で反映される仕様とする(R16H FB-1)。
'   呼び出しチョークポイントはmodLive.PaintStage(取込・Q&A双方の長時間
'   ブロック直前に必ず通る)。LO実行時はこの経路へ到達しない(UI層のため
'   pure対象外)が、念のためOn Error Resume Nextで自滅しないようにする。
' ----------------------------------------------------------------------------
Public Sub EnsureNoGhosting()
    If mChecked Then Exit Sub
    mChecked = True
    On Error Resume Next
    If modConfig.GetBool("freeze_keep_banner", True) Then
        DisableProcessWindowsGhosting
    End If
    On Error GoTo 0
End Sub
