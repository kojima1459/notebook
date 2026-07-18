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

' ----------------------------------------------------------------------------
' Enter - ロック取得を試みる。取得できたらTrue(処理続行可)、既に処理中ならFalse
'         (呼び出し側は即Exitすること)。取得時は砂時計+ステータスバーを表示。
' ----------------------------------------------------------------------------
Public Function Enter() As Boolean
    If mBusy Then
        Enter = False
        Exit Function
    End If
    mBusy = True
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
    On Error Resume Next
    Application.Cursor = -4143                  ' xlDefault
    Application.StatusBar = False
    On Error GoTo 0
    On Error Resume Next
    modUI.ParkFocus                             ' Shape選択解除+アクティブセルpark(C4/A2)
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' IsBusy - 現在ロック中か(参照用)。
' ----------------------------------------------------------------------------
Public Function IsBusy() As Boolean
    IsBusy = mBusy
End Function
