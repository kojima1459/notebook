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
'   (1) 取込バナー内、■中断の左隣(modProgressBar.PaintProgressのcancellable=True
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
'     取込全体を止めてしまう=憲章§3-1)。modSkin.ShowToast(error・waitless)+
'     modLog.LogError(新コードE0904)に留める(2026-08-10 R27波3-7で
'     Application.StatusBarから移設。ステータスバーは非表示設定のため
'     そこに書いた文字は誰にも届いていなかった)。quietFail引数はこのための
'     分岐で、OnOpenWorkExcel(バナー内ボタンからも呼ばれる共有経路)からは
'     常にTrueで渡す。
'   ・DisableProcessWindowsGhosting(user32、引数なし・戻り値なし・表示系)は
'     本リポジトリ初のDeclare(docs/dev/EDGE_CASES.md §1.2に例外条項を追記
'     済み。32bit実機のみを前提とし、LOコンパイル通過を実験で確認済み)。
'     呼ぶとプロセス終了まで解除できない公式仕様のため、config
'     freeze_keep_banner でオプトイン方式にしている(2026-09-11 R46で既定off。
'     理由は EnsureNoGhosting の直上を参照)。
' ============================================================================

#If VBA7 Then
    Private Declare PtrSafe Sub DisableProcessWindowsGhosting Lib "user32" ()
#Else
    Private Declare Sub DisableProcessWindowsGhosting Lib "user32" ()
#End If

Private mLastOpenAt As Double          ' Timer値。0=未実行(2秒デバウンス用)
' R18H FA-1(A-H1/B-H2): 2段目確認のセッション記憶。mLaunched=このセッションで
' 別Excelを起動済み(どの入口からでも立てる)。mOfferAnswered=2段目に一度答えた。
' 従来は取込前確認が【資料1件ごと】に2段目まで出したため、20件のスキャンPDFを
' 選ぶと同じ質問が20回出た(1回目に「はい」で開いた人にも、「いいえ」と答えた
' 人にも聞き続ける=憲章§3-1/§3-4)。1段目(所要時間の見積り)はファイル単位が
' 正しいのでそのまま残し、2段目だけを1セッション1回に畳む。
Private mLaunched As Boolean
Private mOfferAnswered As Boolean
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

' ----------------------------------------------------------------------------
' OpenWorkExcelNow - デバウンスを通さない起動口(2026-08-05 R18-1f)。
'   取込前確認の2段目(下の OfferBeforeIngest)から呼ばれる。人が
'   モーダルへ「はい」と答えた直後の1回きりで、連打は構造的に起きない。
'   逆にデバウンスを通すと、直前にバナー内のボタンを押していた人だけが
'   2秒の壁で黙って無視される(押したのに何も起きない=憲章§3-1)。
'   失敗しても MsgBox は出さない(quietFail:=True。取込ループの直前で
'   モーダルを重ねない)。記録は DoOpenWorkExcel 側の E0904 が残す。
' ----------------------------------------------------------------------------
Public Sub OpenWorkExcelNow()
    On Error Resume Next
    DoOpenWorkExcel True
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' OfferBeforeIngest - 取込を始める直前の2段目確認(R18-1f→R18H FA-1で移設)。
' ----------------------------------------------------------------------------
' 「はい」で作業用Excel(別プロセスの excel.exe /x)を先に開いてから取込へ入る。
' 「いいえ」はそのまま続行する(どちらを選んでも取込は必ず始まる=この
' ダイアログで取込が止まることはない)。文言は BMP の文字だけで書く: 非BMPの
' 絵文字は CP932 実行文への変換で化け、ダイアログでは化けたまま利用者に届く
' (実機第5報⑤)。ここでは「作業用Excel」と名前で呼ぶ。
' 1セッション1回だけ(FA-1): 既に起動済み(mLaunched)なら聞かずに続行し、
' 一度答えていれば(mOfferAnswered)二度と聞かない。呼び出し元は取込前確認の
' 2段目1箇所だけで、置き場をUI層のここにしたのは【記憶とダイアログと起動を
' 1モジュールに閉じる】ため(取込層に旗を置くと同じ判断が2箇所に散る=§4-5)。
' デバウンスは通さない(DoOpenWorkExcel直呼び)。ここは連打の起きないモーダル
' 直後で、直前に自分でボタンを押していた人が2秒の壁で無視されると
' 「はいを押したのに何も起きない」になる(憲章§3-1)。
' 表示・起動の失敗が取込を壊してはならないので全体を OERN で包む(§4-4)。
Public Sub OfferBeforeIngest()
    On Error Resume Next
    If mLaunched Or mOfferAnswered Then Exit Sub
    mOfferAnswered = True
    Dim msg As String
    msg = "先に作業用Excelを開いてから開始しますか?" & vbLf & _
          "(取込中はこのExcelを操作できません)" & vbLf & vbLf & _
          "[はい] 別のExcelを開いてから取り込みます" & vbLf & _
          "[いいえ] このまま取り込みます" & vbLf & vbLf & _
          "この確認は今回のご利用で1回だけです。"
    If MsgBox(msg, vbQuestion + vbYesNo + &H10000, modAppDef.APP_NAME) = vbYes Then
        OpenWorkExcelNow
        modLog.LogUsage "work_excel_preopen", "", "取込前確認から作業用Excelを起動しました"
    End If
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
    ' R18H FA-1: どの入口(バナー内ボタン/ヘルプカード/取込前確認)から起動しても
    ' 「もう開いてある」事実は同じ。2段目確認はこの旗を見て黙って続行する。
    mLaunched = True
    ' R19-5e(実機第6報⑤): 一文だけ「閉じてください」を足す。開いたまま残った
    ' 作業用Excelは、次に本体xlsmをダブルクリックしたときの【吸い込み先】に
    ' なる(Excelは既定で複数ブックを1プロセスへ結合する。調査⑤ウェブ班4章)。
    ' そうなると本体のOCR同期呼び出しがプロセス全体を止め、作業用Excelごと
    ' 固まる=実機第6報⑤そのもの。ここは起動直後の一番短い経路で、この機能を
    ' 使った人【全員】が必ず一度は通る唯一の場所なので、ここ1箇所だけに置く。
    ' 文字はBMPのみ(非BMPはCP932往復で化ける=実機第5報⑤)。
    '
    ' 2026-08-06(R19H FA-8 / B-H③): 伝える先を Application.StatusBar から
    ' modSkin.ShowToast へ変えた。本アプリは起動時に【ステータスバー自体を
    ' 非表示にしている】ので、ここに書いた文字は誰の目にも入らない ――
    ' つまり「閉じてください」という一番効く一言を、この機能を使った人全員に
    ' 一度も届けられていなかった(§3-3「次の一手を必ず言う」が空振り)。
    ' waitless:=True にするのは2つの理由から: (1) ここは取込へ入る直前かバナー内
    ' ボタンの直後で、1.1秒の待ちを足したくない (2) この文は複数行ぶんあり、1.1秒で
    ' 消えると読み切れない。残った "nx_toast" は取込終了時の
    ' modProgressBar.ClearProgress と modUI.ClearChat が回収する(R19-3a/3b)。
    '
    ' 2026-09-10(R43 波A 1-4・実機報告Q4): 旧文言(84字)は「Excelが2つに
    ' なること」「どちらが作業用か」「どちらを閉じるか」を一度も言っておらず、
    ' 実機で「2個のどっちを消せば?」と聞き返された。この3点を明示する文言
    ' (103字・全角換算91字)へ差し替える(spec_20260910_R43 §1-4)。
    ' modTestsPure43.TestR41SweepMs43 のゴールデン文字列・期待値(12134)も
    ' 同じ変更で更新済み。
    modSkin.ShowToast "作業用のExcelを開きました。いまExcelが2つあります。" & _
        "新しく開いた空の方が作業用です。そちらで仕事をして、終わったら空の方だけ" & _
        "閉じてください。MyBookshelf(この画面)は閉じないでください。", _
        "info", True
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
    ' 2026-08-10(R27波3-7): 失敗の一言を Application.StatusBar から
    ' modSkin.ShowToast へ移す。本アプリは起動時にステータスバー自体を
    ' 非表示にしているので、ここに書いた文字は【誰の目にも入っていなかった】
    ' (同じ理由で成功側の案内を R19H FA-8 でトーストへ移してある。:169-176)。
    ' quietFail(バナー内ボタン・取込前確認からの呼び出し)でもトーストは出す。
    ' モーダルを出さないのは取込を止めないためであって、黙るためではない
    ' ―― 押したのに何も起きなければ利用者には故障と区別が付かない(憲章§3-1)。
    ' waitless:=True: 取込へ入る直前に1.1秒の待ちを足さない(:172と同じ理由)。
    modSkin.ShowToast "作業用Excelを起動できませんでした。" & _
        "しばらくしてから再度お試しください(コード: E0904)。", "error", True
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
'   (R16-2a)。modProgressBar.PaintProgressのcancellable=True分岐から呼ばれる。
'   座標は中断ボタンの位置(cancelLeftPos/cancelTopPos)を受け取り、その左へ
'   幅+間隔ぶんずらして自分の位置を決める(modSkin側は座標計算を持たない)。
'   削除はmodProgressBar.ClearProgress側にPROGRESS_WORK_NAMEとして同列に登録済み
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
        btn.TextFrame2.TextRange.Text = modEmj.WindowIcon() & "作業用Excel"
        btn.OnAction = "modWorkExcel.OnOpenWorkExcel"
    Else
        btn.Left = leftPos: btn.Top = cancelTopPos
    End If
    btn.ZOrder 0
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' EnsureNoGhosting - DWMの「応答なし」白画面化(ゴースト化)を抑止する(R16-2b)。
'   config freeze_keep_banner(2026-09-11 R46で既定off)がオフなら何もしない。
'   【既定をoffにした理由】ゴーストウィンドウはOSが無応答アプリへ用意した唯一の
'   復旧経路で、タスクバーからの復元・右クリック→閉じる・タスクの終了はすべて
'   そこを通る。抑止すると進捗バナーは見え続けるが、その3つが同時に死ぬ
'   (R45実機: 取込中に最小化したらPC再起動以外に手が無く、xlsm のロックも残って
'   再実行が読み取り専用になった)。バナーの見た目より復旧手段を優先する。
'   GetBool の第2引数(configにキーが無いときの値)もFalseにしてある――ここを
'   Trueのままにすると、キーを持たない旧configや壊れたconfigで抑止が復活する。
'   プロセス生存中に
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
    If modConfig.GetBool("freeze_keep_banner", False) Then
        DisableProcessWindowsGhosting
    End If
    On Error GoTo 0
End Sub
