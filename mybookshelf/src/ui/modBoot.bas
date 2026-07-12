Attribute VB_Name = "modBoot"
Option Explicit

' ============================================================================
' modBoot - Workbook_Open時の起動シーケンス(MASTER_SPEC §7.6)
' ----------------------------------------------------------------------------
' 役割:
'   config確認→初回起動時のpack_author入力→3画面(ホーム/マイ本棚/
'   ダッシュボード)のEnsureLayout→QuickHealthCheck→sync_on_openなら
'   SyncNow→ScheduleAutoSync→内部シート隠蔽、という起動の一本道を
'   このモジュールだけが所有する。ThisWorkbook.clsのWorkbook_Open/
'   Workbook_BeforeCloseはBoot/Auto_Closeへの薄い転送のみで、実際の
'   ロジックはここに置く(ビルド版ではインストーラがThisWorkbookを
'     占有するため、実行時はAuto_Open/Auto_Closeが本命という二重化設計。
'     §7.6末尾の指示どおり)。
'
' 設計判断:
'   ・gBootDoneガード: 2回目以降のBoot呼び出し(Auto_Open経由の二重呼び出し等)
'     では、重い初期化(first-run入力・SyncNow等)を再実行せず、3画面の
'     EnsureLayoutだけをやり直して画面を整える(冪等な「開き直し」)。
'   ・first-run(pack_author入力): config の pack_author が空の場合だけ
'     InputBoxで尋ねる(絵文字は使わない。§12「MsgBox/InputBox文字列では
'     絵文字を使わない」規約の厳守)。空欄のまま閉じた場合は「名称未設定」を
'     既定値として保存し、次回以降は再度尋ねない(毎回聞かれる煩わしさを
'     避ける。あとでconfigシートからいつでも変更できる旨を案内文に含める)。
'   ・3画面のEnsureLayout直後に、RenderShelf/EvaluateBadges/RenderDashboard
'     もあわせて呼ぶ。理由: このブックには各シートのActivateイベントに
'     反応するシートクラスモジュールが割り当てられていない(担当ファイルは
'     ThisWorkbook.clsのみ)。タブを切り替えるだけでは再描画が起きないため、
'     非エンジニアが開いた瞬間に「空っぽの画面」を見せないよう、起動時に
'     一度だけ実データで埋めておく(以降はmodUIMain/modUIShelf側の各操作が
'     完了時に自分で再描画する設計に委ねる)。
'   ・本棚が空のときは、modShelf.TotalChunksを見てmodUIMain.
'     ShowEmptyShelfHintを呼ぶ(§8.1「まず『マイ本棚』タブで資料を1つ
'     追加してみましょう →」常設表示)。「空かどうか」の判定はここ(modBoot)
'     の責務、「どう見せるか」はmodUIMainの責務、と役割を分けている。
'   ・sync_on_openによるSyncNowはOn Error Resume Next(1行スコープ)で
'     保護する(同期に失敗してもブック自体は開けるようにする。§13
'     「フォルダ削除・リネーム」等のエッジケースでBoot全体を落とさない
'     ため)。
'   ・Auto_Closeは必ずmodShelfSync.CancelAutoSyncを呼ぶ(§12「OnTime予約
'     解除はmodBoot.Auto_Closeから必ず呼ぶ(OnTime残存→勝手にExcelが
'     再起動する事故防止)」)。StatusBarも既定に戻す(=False)。
'   ・gBootDoneはPrivateモジュール変数(§7.6契約のPublicはBoot/Auto_Open/
'     Auto_Closeの3本のみ。状態フラグを外部に晒す必要は無い)。
' ============================================================================

Private gBootDone As Boolean

' ----------------------------------------------------------------------------
' Auto_Open - Boot呼び(ガード付き)
' ----------------------------------------------------------------------------
Public Sub Auto_Open()
    Boot
End Sub

' ----------------------------------------------------------------------------
' Boot - V2パターン踏襲の起動シーケンス本体
' ----------------------------------------------------------------------------
Public Sub Boot()
    If gBootDone Then
        On Error Resume Next
        modUIMain.EnsureLayout
        On Error GoTo 0
        On Error Resume Next
        modUIShelf.EnsureLayout
        On Error GoTo 0
        On Error Resume Next
        modUIDashboard.EnsureLayout
        On Error GoTo 0
        Exit Sub
    End If

    On Error GoTo Failed

    ' 1) config確認
    modConfig.EnsureLoaded

    ' 2) first-run: pack_author入力
    EnsureFirstRun

    ' 3) 3画面EnsureLayout(それぞれが内部でRenderShelf/RenderDashboardまで実行する)
    modUIMain.EnsureLayout
    modUIShelf.EnsureLayout
    modUIDashboard.EnsureLayout

    ' Wave4修正: modStats.TouchToday(streak_days/last_used_date更新)を
    ' どこからも呼んでいなかったため、streak7バッジもダッシュボードの
    ' 連続利用日数も永久に更新されない不具合があった。1セッション1回
    ' (gBootDoneで守られたこの初回Boot経路)呼べばTouchTodayの契約
    ' (§7.5: 昨日なら+1/今日なら不変/それ以外リセット)は成立する。
    On Error Resume Next
    modStats.TouchToday
    On Error GoTo 0

    On Error Resume Next
    modStats.EvaluateBadges
    On Error GoTo 0
    On Error Resume Next
    modUIDashboard.RenderDashboard
    On Error GoTo 0

    ' 本棚が空なら回答エリアに常設案内を出す
    Dim isEmpty As Boolean
    isEmpty = False
    On Error Resume Next
    isEmpty = (modShelf.TotalChunks() = 0)
    On Error GoTo 0
    If isEmpty Then
        On Error Resume Next
        modUIMain.ShowEmptyShelfHint
        On Error GoTo 0
    End If

    ' 4) QuickHealthCheck
    Dim warn As String
    warn = modDiag.QuickHealthCheck()
    If LenB(warn) > 0 Then
        MsgBox "起動時の確認で気になる点がありました。" & vbLf & warn & vbLf & _
               "詳しくは診断ボタンで確認できます。", vbExclamation, modAppDef.APP_NAME
    End If

    ' 5) sync_on_openならSyncNow(On Error保護)。起動時の自動同期はユーザー
    '    操作への応答ではないため silent:=True で呼び、完了/未設定時の警告
    '    ダイアログでブックを開いた直後にポップアップを出さない
    '    (Wave4修正: 名前入力直後にE0502警告が必ず出る不具合への対応)。
    If modConfig.GetBool("sync_on_open", True) Then
        On Error Resume Next
        modShelfSync.SyncNow silent:=True
        On Error GoTo 0
    End If

    ' 6) ScheduleAutoSync
    On Error Resume Next
    modShelfSync.ScheduleAutoSync
    On Error GoTo 0

    ' 7) 内部シート隠蔽
    HideInternalSheets

    gBootDone = True
    Exit Sub

Failed:
    modLog.LogError "E0801", "modBoot.Boot", Err.Description
    MsgBox modLog.FriendlyMessage("E0801") & vbLf & "(コード: E0801)", vbCritical, modAppDef.APP_NAME
End Sub

' ----------------------------------------------------------------------------
' Auto_Close - CancelAutoSync(必須)+Application.StatusBar=False
' ----------------------------------------------------------------------------
Public Sub Auto_Close()
    On Error Resume Next
    modShelfSync.CancelAutoSync
    On Error GoTo 0

    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

Private Sub EnsureFirstRun()
    Dim cur As String
    cur = Trim$(modConfig.GetString("pack_author", ""))
    If LenB(cur) > 0 Then Exit Sub

    Dim nm As String
    nm = InputBox("はじめまして。あなたのお名前(またはニックネーム)を教えてください。" & vbCrLf & _
                  "資料をパックとして他の人と共有するとき、作成者として表示されます。" & vbCrLf & _
                  "(あとで config シートからいつでも変更できます)", _
                  modAppDef.APP_NAME & " - はじめの設定")
    nm = Trim$(CStr(nm))
    If LenB(nm) = 0 Then nm = "名称未設定"
    modConfig.SetValue "pack_author", nm
End Sub

' V2 HideInternalSheets踏襲。§4の可視性表のとおり:
'   veryHidden: my_knowledge / my_vectors / ui_state / vba_src
'   hidden    : config / my_manifest / my_stats / usage_log / err_log
Private Sub HideInternalSheets()
    Const VERY_HIDDEN As Long = 2   ' xlSheetVeryHidden
    Const HIDDEN As Long = 0        ' xlSheetHidden

    HideSheetSafely modAppDef.SH_KNOWLEDGE, VERY_HIDDEN
    HideSheetSafely modAppDef.SH_VECTORS, VERY_HIDDEN
    HideSheetSafely modAppDef.SH_UISTATE, VERY_HIDDEN
    HideSheetSafely "vba_src", VERY_HIDDEN

    HideSheetSafely modAppDef.SH_CONFIG, HIDDEN
    HideSheetSafely modAppDef.SH_MANIFEST, HIDDEN
    HideSheetSafely modAppDef.SH_STATS, HIDDEN
    HideSheetSafely modAppDef.SH_USAGE, HIDDEN
    HideSheetSafely modAppDef.SH_ERRLOG, HIDDEN
End Sub

Private Sub HideSheetSafely(ByVal sheetName As String, ByVal visibility As Long)
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Visible = visibility
    On Error GoTo 0
End Sub
