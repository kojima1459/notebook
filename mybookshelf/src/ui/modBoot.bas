Attribute VB_Name = "modBoot"
Option Explicit

' ============================================================================
' modBoot - Workbook_Open時の起動シーケンス(MASTER_SPEC §7.6)
' ----------------------------------------------------------------------------
' 役割:
'   config確認→初回起動時のpack_author入力→3画面(ホーム/マイ本棚/
'   ダッシュボード)のEnsureLayout→QuickHealthCheck→AIリボン利用期限確認
'   (modGateway.RunLimitCheck)→sync_on_openならSyncNow→ScheduleAutoSync
'   →内部シート隠蔽、という起動の一本道を
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

    ' bootStage: 失敗時に「どの段階で落ちたか」をユーザー向けダイアログと
    ' err_logの両方に出すための段階名。実機で初発のエラーが出たとき、
    ' スクリーンショット1枚で原因箇所まで特定できるようにする
    ' (2026-07-15 実機E0801報告への恒久対策。R5「原因の分からないエラーを
    ' ユーザーに見せない」の徹底)。
    Dim bootStage As String
    On Error GoTo Failed

    ' 起動中の大量のシート書換え・Activate/Select中に(将来追加・注入され得る)
    ' イベントの連鎖発火で無限ループに陥らないよう、イベントを抑止する。
    ' 正常終了・異常終了いずれのパスでも必ずTrueへ戻す(死の連鎖防止)。
    On Error Resume Next
    Application.EnableEvents = False
    On Error GoTo Failed

    ' 1) config確認
    bootStage = "設定の読み込み(config)"
    modConfig.EnsureLoaded

    ' 2) first-run: pack_author入力
    bootStage = "はじめの設定(名前の保存)"
    EnsureFirstRun

    ' 3) 3画面EnsureLayout(それぞれが内部でRenderShelf/RenderDashboardまで実行する)。
    ' 2026-07-16: 実機で「1画面の描画中の不具合がアプリ全体の起動を止める」
    ' 事例が立て続けに見つかった(ホーム→マイ本棚→…と直しても次の画面で
    ' 同種の失敗が起きる)。非エンジニアが使う配布物として、1画面の描画に
    ' 問題があってもアプリ自体は必ず開けるべきなので、3画面それぞれを
    ' 独立してOn Error Resume Nextで保護し、失敗はerr_logに詳細(uiStepまで
    ' 埋め込み済みのErr.Description)を記録した上で次の画面へ進む
    ' (該当画面だけ表示が崩れる可能性はあるが、アプリが開けないよりずっと良い)。
    bootStage = "ホーム画面の組み立て"
    On Error Resume Next
    modUIMain.EnsureLayout
    LogBootStageErrorIfAny bootStage
    On Error GoTo Failed

    bootStage = "マイ本棚画面の組み立て"
    On Error Resume Next
    modUIShelf.EnsureLayout
    LogBootStageErrorIfAny bootStage
    On Error GoTo Failed

    bootStage = "ダッシュボード画面の組み立て"
    On Error Resume Next
    modUIDashboard.EnsureLayout
    LogBootStageErrorIfAny bootStage
    On Error GoTo Failed
    bootStage = ""

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
    bootStage = "起動時の健全性チェック"
    Dim warn As String
    warn = modDiag.QuickHealthCheck()
    If LenB(warn) > 0 Then
        MsgBox "起動時の確認で気になる点がありました。" & vbLf & warn & vbLf & _
               "詳しくは診断ボタンで確認できます。", vbExclamation, modAppDef.APP_NAME
    End If

    ' 4.5) AIリボンの利用期限確認(裁定D3の穏当運用)。True=続行不可でも
    '      アプリの起動自体は止めず、丁寧な案内メッセージだけを出す(AIへの
    '      質問など回答系を実行した際に改めて案内される)。mock_llm=TRUE・
    '      config limit_check=FALSE・リボン未検出のときはRunLimitCheck側で
    '      即False(続行可)になるため、ここでの分岐は不要。
    Dim limited As Boolean
    limited = False
    On Error Resume Next
    limited = modGateway.RunLimitCheck()
    On Error GoTo 0
    If limited Then
        MsgBox "AIリボンの利用期限確認により、現在AI機能の利用が制限されています。" & vbLf & _
               "本棚の閲覧や資料の管理は、これまでどおりお使いいただけます。" & vbLf & _
               "AIへの質問などを実行された際には、あらためてご案内します。" & vbLf & _
               "(お手数ですが、AIリボンの利用期限については担当部署へお問い合わせください)", _
               vbInformation, modAppDef.APP_NAME
    End If

    ' 5) sync_on_openならSyncNow(On Error保護)。起動時の自動同期はユーザー
    '    操作への応答ではないため silent:=True で呼び、完了/未設定時の警告
    '    ダイアログでブックを開いた直後にポップアップを出さない
    '    (Wave4修正: 名前入力直後にE0502警告が必ず出る不具合への対応)。
    If modConfig.GetBool("sync_on_open", True) Then
        On Error Resume Next
        modShelfSync.SyncNow silent:=True
        On Error GoTo 0
    Else
        ' sync_on_open=Falseでも、起動時にP2P感謝状・品質報告集計は一度回収する
        ' (sync_on_open=TrueのときはSyncNow内で回収済み=二重回収しない)。
        On Error Resume Next
        modP2P.CollectThanks silent:=True
        modP2P.CollectNoiseVotes silent:=True
        On Error GoTo 0
    End If

    ' 6) ScheduleAutoSync
    On Error Resume Next
    modShelfSync.ScheduleAutoSync
    On Error GoTo 0

    ' 7) 内部シート隠蔽
    bootStage = "内部シートの整理"
    HideInternalSheets

    ' 8) Nexus UI(config nexus_ui=TRUEのとき新SPA UIを起動。失敗しても
    '    旧3画面は生きているため、起動自体は続行する)
    If modConfig.GetBool("nexus_ui", False) Then
        bootStage = "Nexus画面の起動"
        On Error Resume Next
        modApp.LaunchNexus
        On Error GoTo Failed
        bootStage = ""
    End If

    gBootDone = True
    On Error Resume Next
    Application.EnableEvents = True   ' 起動中に抑止したイベントを復帰
    On Error GoTo 0
    Exit Sub

Failed:
    ' EnsureLayout途中で落ちるとScreenUpdating=Falseのまま画面が固まって
    ' 見えるため、必ず戻す(ダイアログより先に)。
    Dim failDesc As String
    failDesc = Err.Description
    Dim failNum As Long
    failNum = Err.Number
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True   ' イベント抑止も必ず復帰(死の連鎖防止)
    On Error GoTo 0
    modLog.LogError "E0801", "modBoot.Boot", _
        "stage=" & bootStage & " err#" & failNum & ": " & failDesc
    ' 「原因の分からないエラー」を出さない: どの段階で何のエラーが起きたかを
    ' ダイアログに含める。利用者はこの画面の写真を送るだけで報告が完結する。
    MsgBox modLog.FriendlyMessage("E0801") & vbLf & vbLf & _
           "失敗した処理: " & bootStage & vbLf & _
           "エラー内容: " & modUtil.SafeLeft(failDesc, 200) & " (#" & failNum & ")" & vbLf & _
           "(コード: E0801)" & vbLf & vbLf & _
           "この画面を撮影して管理者へ送っていただければ、原因を特定できます。", _
           vbCritical, modAppDef.APP_NAME
End Sub

' 3画面EnsureLayoutを個別にOn Error Resume Nextで保護したときの記録役。
' Err.Number<>0のときだけE0801としてerr_logへ書き、Errをクリアする
' (呼び出し側は直後にOn Error GoTo Failedへ戻すので、ここではErrを
' 汚さないよう自分の中だけで完結させる)。
Private Sub LogBootStageErrorIfAny(ByVal stage As String)
    If Err.Number = 0 Then Exit Sub
    Dim n As Long, d As String
    n = Err.Number
    d = Err.Description
    Err.Clear
    On Error Resume Next
    modLog.LogError "E0801", "modBoot.Boot", "stage=" & stage & " err#" & n & ": " & d
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Auto_Close - CancelAutoSync(必須)+Application.StatusBar=False
' ----------------------------------------------------------------------------
Public Sub Auto_Close()
    On Error Resume Next
    modShelfSync.CancelAutoSync
    On Error GoTo 0

    ' Nexusが隠したネイティブUI(リボン等)を必ず復元する
    On Error Resume Next
    modUI.RestoreExcelUI
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
