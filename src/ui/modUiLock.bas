Attribute VB_Name = "modUiLock"
Option Explicit

' ============================================================================
' modUiLock - グローバルUIロック(全アクションの単一関所)。盲点A1/D7への恒久対策。
' ----------------------------------------------------------------------------
' 役割:
'   ・二重送信/連打防止: 全てのOnAction入口が Enter() を先頭で呼び、Falseなら
'     即Exitする。処理中(API通信/K-Means/描画)の多重発火をこの1点で封じる。
'   ・re-entrancy防止: 重処理中はDoEventsでイベントループが回るため、ユーザーが
'     別のボタンやタブを押すと処理が入れ子で走り得る(§盲点D7)。mBusyフラグで
'     入れ子の再入を弾く。
'   ・ローディング状態の可視化: ロック取得中は砂時計カーソル+ステータスバーで
'     「処理中」を明示(Webアプリのボタン無効化＝グレーアウトに相当)。
'
' 設計判断:
'   ・状態はモジュールprivateのmBusy 1個(グローバルは1つで十分。画面横断で
'     単一の処理しか走らせない設計)。旧: modApp内のmBusy(送信系のみ)を本ロックへ
'     一本化し、👍👎/解決/照会/Word/ナビ/コピー等の未保護ハンドラも同じ関所を通す。
'   ・Enter/Leaveは必ず対で呼ぶ。呼び出し側は On Error GoTo で異常時も必ず
'     Leaveへ到達させること(ロック取りっぱなしでUIが永久ロックする事故を防ぐ)。
'   ・Leaveはフォーカスpark(modUI.ParkFocus)も行う。全アクション完了時に必ず
'     Shape選択を解除(白い選択ハンドルを消す=盲点C4)し、アクティブセルを安全な
'     位置へ戻す(矢印キーでのスクロール崩壊=盲点A2の防止)。これを1点に集約する
'     ことで、各ハンドラがparkを呼び忘れても必ずparkされる。
'   ・視覚系(カーソル/ステータスバー/park)は全てOn Error Resume Next配下。
'     復帰処理そのものが例外でロック解除を妨げてはならない。
'   ・LibreOfficeは静的コンパイルのみ(実行しない)。Application.Cursor等は
'     コンパイル可能な標準VBAのため構文検査は通る。
' ============================================================================

Private mBusy As Boolean
' 2026-07-28(レビュー M-9): Leave の呼び忘れが1箇所でもあると、全ボタンが
' 【無音・無期限】に死ぬ。利用者にできることが何も無い(再起動しかない)。
' modShelfSync の同種ガードには30分の自動解除があるのに、こちらには無かった。
' 全数調査では現時点で漏れは見つからなかったが、将来の1バグで全滅する
' 構造そのものが危うい。取得時刻を控え、期限を過ぎたロックは無効とみなす。
Private mBusySince As Date
Private Const LOCK_EXPIRY_MIN As Long = 10

' R13-4d: 直近に「はい(中断して終了)」と答えた時刻。開発構成では終了ボタン
' (modApp.OnSaveAndExit)が ThisWorkbook.Close を呼ぶと Workbook_BeforeClose が
' 続けて発火するため、同じ問いが二度出る。二度目は既定が「いいえ」なので、
' 一度承諾した終了がそこで止まって見える。承諾は短時間だけ憶えておく。
Private mCloseOkSince As Date
Private Const CLOSE_OK_SEC As Long = 60

' ----------------------------------------------------------------------------
' Enter - ロック取得を試みる。取得できたらTrue(処理続行可)、既に処理中ならFalse
'         (呼び出し側は即Exitすること)。取得時は砂時計+ステータスバーを表示。
' ----------------------------------------------------------------------------
Public Function Enter() As Boolean
    If mBusy Then
        If Not LockExpired() Then
            Enter = False
            ' 押しても無反応、では利用者は壊れたと判断する。何が起きているかは
            ' 一言でも返す(StatusBar は Nexus 画面では隠れているためトースト)。
            On Error Resume Next
            modSkin.ShowToast "まだ前の処理が動いています。少しお待ちください。", "info"
            On Error GoTo 0
            Exit Function
        End If
        ' 期限切れ。前の処理は Leave に到達せず落ちたとみなして奪い返す。
        On Error Resume Next
        modLog.LogError "E0801", "modUiLock.Enter", _
            "UIロックが" & LOCK_EXPIRY_MIN & "分を超えたため自動解除しました"
        On Error GoTo 0
    End If
    mBusy = True
    mBusySince = Now
    Enter = True
    On Error Resume Next
    Application.Cursor = 2                      ' xlWait(砂時計)
    Application.StatusBar = "処理中です。少々お待ちください..."
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' Leave - ロック解放。カーソル/ステータスバーを既定へ戻し、フォーカスをpark。
'         正常・異常どちらの経路からも必ず呼ぶこと。
' ----------------------------------------------------------------------------
Public Sub Leave()
    mBusy = False
    mBusySince = 0
    On Error Resume Next
    Application.Cursor = -4143                  ' xlDefault
    Application.StatusBar = False
    ' 暗転固定の全クラス対策: 描画ルーチンがScreenUpdating=Falseのままエラー中断しても、
    ' 全アクションの出口である本Leaveで必ず復帰させる(UIロックアウト=脱出不能の防止)。
    Application.ScreenUpdating = True
    On Error GoTo 0
    On Error Resume Next
    modUI.ParkFocus                             ' Shape選択解除+アクティブセルpark(C4/A2)
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' IsBusy - 現在ロック中か(参照用)。
' ----------------------------------------------------------------------------
Public Function IsBusy() As Boolean
    If Not mBusy Then Exit Function
    IsBusy = Not LockExpired()
End Function

' ----------------------------------------------------------------------------
' BlockIfIngesting - 取込・同期の最中なら実況だけ出してTrueを返す(R7 B-2)。
' ----------------------------------------------------------------------------
' 取込は長時間かかるうえ、抽出ループの DoEvents でメッセージが捌かれるため、
' その最中に押されたボタン(OnAction)は【取込の途中から入れ子で】走り出す。
' 実機報告「取込中にボタンを押すとExcelが応答なし」の正体はこれで、
' 画面遷移が取込の掴んでいるCOMオブジェクトやシート状態と噛み合わなくなる。
'
' 画面遷移系の入口はここを最初に通し、busyなら何もせず即Exitする。
' モーダルは出さない(連打のたびにダイアログが積み上がると、それ自体が
' 「応答なし」に見える)。実況行だけで「待てばよい」ことを伝える。
' modUiLock 自身のロック(Enter/Leave)とは別物: あちらは1アクションの二重
' 発火を防ぐもので、こちらは【別プロセスのように長く走る取込】から画面を守る。
Public Function BlockIfIngesting() As Boolean
    Dim busyNow As Boolean
    On Error Resume Next
    busyNow = modShelf.IsBusy()
    If Not busyNow Then busyNow = modShelfSync.IsBusy()
    On Error GoTo 0
    If Not busyNow Then Exit Function

    BlockIfIngesting = True
    On Error Resume Next
    modUIMain.SetStage "処理中です。完了までお待ちください…"
    ' SetStageの実況(ホームのセル/StatusBar/チャットバブル)はマイ本棚系画面
    ' からは不可視なので、どの画面でも見える進捗バナーへ出す。
    ' R10c(M3): ここは R10-1 で ShowToast にしていたが、トーストは表示に
    ' 1.1秒のブロッキング待ちが入る。連打されるほど待ちが積もり、
    ' 「押しても無反応」を直すつもりが「押すほど固まる」を作っていた
    ' (取込中に連打されるのがまさにこの関所)。待ちゼロのバナーへ置換する。
    ' バナーは取込側の次のShowProgress更新かHideProgressで上書き/消去される。
    modUIMain.ShowProgress "取り込み処理が終わるまでお待ちください…"
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' ConfirmCloseDuringIngest - 取込中の終了を、利用者に一度だけ確かめる(R13-4d)。
' ----------------------------------------------------------------------------
' 戻り値 True=閉じてよい / False=閉じてはいけない(呼び出し元は終了を中止)。
'
' R11-A C1 では取込中の終了を【問答無用で拒否】していた。データ保全としては
' 正しいが、GSの本文抽出やOCRは数分かかることがあり、その間は終了ボタンも
' ウィンドウの×も無反応にしか見えない(憲章§3-1「無反応は故障と同義」)。
' 何が起きているかを伝えたうえで、中断して終わる道も残す。既定は「いいえ」
' (vbDefaultButton2)にして、Enterの流し打ちで取込が飛ばないようにする。
'
' busy 判定は BlockIfIngesting と同じ2つ(取込・フォルダ同期)。判定を
' 2箇所で違えないよう、参照する関数はここでも同じものだけを使う。
' 例外は外へ出さない(判断に失敗したら閉じてよい側=利用者の操作を
' 妨げない側へ倒す。ロックが焼き付いてブックを閉じられなくなる方が害が大きい)。
Public Function ConfirmCloseDuringIngest() As Boolean
    ConfirmCloseDuringIngest = True

    Dim busyNow As Boolean
    On Error Resume Next
    busyNow = modShelf.IsBusy()
    If Not busyNow Then busyNow = modShelfSync.IsBusy()
    On Error GoTo 0
    If Not busyNow Then Exit Function

    ' 同じ終了操作の中で二度聞かない(終了ボタン→BeforeClose の連鎖)。
    On Error Resume Next
    If mCloseOkSince <> 0 Then
        If DateDiff("s", mCloseOkSince, Now) < CLOSE_OK_SEC Then Exit Function
    End If
    On Error GoTo 0

    Dim resp As VbMsgBoxResult
    resp = vbNo
    On Error Resume Next
    resp = MsgBox("資料の取込中です。中断して終了しますか?" & vbLf & vbLf & _
        "  [はい] 取込を中断して終了します(取込中の資料は入りません)" & vbLf & _
        "  [いいえ] 終了せず、取込の完了を待ちます", _
        vbYesNo + vbQuestion + vbDefaultButton2, modAppDef.APP_NAME)
    On Error GoTo 0

    ConfirmCloseDuringIngest = (resp = vbYes)
    On Error Resume Next
    If ConfirmCloseDuringIngest Then
        mCloseOkSince = Now
        modLog.LogUsage "close_confirmed_ingesting", "", _
            "取込中の終了を利用者が承諾しました(取込中の資料は入りません)"
    Else
        modLog.LogUsage "close_canceled_ingesting", "", _
            "取込中の終了要求を利用者の選択で中止しました"
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' CancelCloseOk - 「中断して終了」の承諾を取り消す(R13 F5)。
' ----------------------------------------------------------------------------
' 承諾(mCloseOkSince)は「同じ終了操作の中で二度聞かない」ためだけの短命な印で、
' 終了そのものを最後までやり切ったときにしか意味を持たない。ところが承諾の直後に
' 出る保存確認で[キャンセル]を押すと、終了は中止されたのに印だけ60秒生き残る。
' その60秒のあいだにウィンドウの×を押すと、取込中でも確認が一切出ないまま
' 閉じてしまう ―― 一度「やめる」と言った利用者の資料が黙って消える経路になる。
' 終了を取りやめる全ての分岐からここを呼び、承諾を捨てる。
Public Sub CancelCloseOk()
    mCloseOkSince = 0
End Sub

' 取得から LOCK_EXPIRY_MIN 分を過ぎたか(=Leave 漏れの疑い)。
' 時刻が読めないときは期限切れ扱いにする(判断できないロックで
' 全ボタンを殺し続けるより、奪い返す方が害が小さい)。
Private Function LockExpired() As Boolean
    On Error GoTo Expired
    If mBusySince = 0 Then
        LockExpired = True
        Exit Function
    End If
    LockExpired = (DateDiff("n", mBusySince, Now) >= LOCK_EXPIRY_MIN)
    Exit Function
Expired:
    LockExpired = True
End Function
