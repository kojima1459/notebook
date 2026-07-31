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

' 2026-07-28(レビュー H-1): Bootが1回の起動で2回走る経路がある。
'   Workbook_Open → Install が OnTime(+1秒) で Boot を予約
'   → 直後に保存済みの Auto_Open が Boot を同期実行
'   → Boot は共有同期＋起動ジッタ(DoEventsループ)で1秒を超えるため、
'      そのDoEvents中に期限到来済みのOnTime Bootが「割り込んで」発火する。
' gBootDone は Boot の最後で立つので、この時点ではまだ False。つまり
' フル初期化が二重に走り、MsgBoxが2連発したり SyncNow が二重に走ったり、
' AutoSyncTick の予約が追跡不能になったりする。
' gBootDone とは別に「今まさに走っている」フラグが要る。
Private mBootRunning As Boolean

' インストーラが予約した OnTime の時刻を置く場所(vba_src シートのE1)。
' OnTime のキャンセルには予約時刻そのものが要るが、インストーラ側の
' ローカル変数はプロシージャを抜けた時点で消えるため、シートに残す。
' 保存後に書くので .xlsm には残らない(セッション内だけの一時状態で正しい)。
Private Const BOOT_SCHED_SHEET As String = "vba_src"
Private Const BOOT_SCHED_ROW As Long = 1
Private Const BOOT_SCHED_COL As Long = 5

' 軽量マクロ無効ガード(盲点D1)の案内シート名。ビルド時に先頭・可視で作られ、
' マクロ無効で開かれた場合はこのシートがそのまま見える(壊れたUIを見せない)。
' 起動が成功したこの経路(HideInternalSheets)でだけ隠す。ビルド側の
' build_mybookshelf.py の GUARD_SHEET_NAME と文字列を一致させること。
Private Const GUARD_SHEET As String = "はじめにお読みください"

' ----------------------------------------------------------------------------
' Auto_Open - Boot呼び(ガード付き)
' ----------------------------------------------------------------------------
Public Sub Auto_Open()
    Boot
End Sub

' ----------------------------------------------------------------------------
' CancelPendingInstallerBoot - インストーラが予約した OnTime Boot を取り消す。
'
' 自己インストーラ(ThisWorkbookストリーム)は Workbook_Open の中で
' Application.OnTime Now+1秒, "modBoot.Boot" を予約する。VBEへの注入直後に
' 同期でBootを呼ぶと1004になることがあるための逃げなのだが、予約時刻を
' どこにも残していなかったため、誰もキャンセルできなかった。
' 予約時刻は vba_src シートのE1にDouble(シリアル値)で置いてもらい、
' 先に走った側とAuto_Closeがそれを使って取り消す。
'
' OnTime のキャンセルは EarliestTime が一致しないと 1004 になるが、
' 「既に発火済み」も同じく 1004 なので、失敗は握って構わない(取り消す
' 相手がもう居ないだけ)。セルは必ず消してから取り消しにいくので、
' 2回目以降の呼び出しは何もしない。
' ----------------------------------------------------------------------------
Private Sub CancelPendingInstallerBoot()
    On Error Resume Next
    Dim w As Worksheet
    Set w = ThisWorkbook.Worksheets(BOOT_SCHED_SHEET)
    If w Is Nothing Then Exit Sub

    Dim v As Variant
    v = w.Cells(BOOT_SCHED_ROW, BOOT_SCHED_COL).Value
    If IsError(v) Then Exit Sub
    If Not IsNumeric(v) Then Exit Sub
    Dim serial As Double
    serial = CDbl(v)
    If serial <= 0 Then Exit Sub

    w.Cells(BOOT_SCHED_ROW, BOOT_SCHED_COL).ClearContents
    Application.OnTime EarliestTime:=CDate(serial), Procedure:="modBoot.Boot", Schedule:=False
    Err.Clear
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Boot - V2パターン踏襲の起動シーケンス本体
' ----------------------------------------------------------------------------
Public Sub Boot()
    ' インストーラが予約した「+1秒後のBoot」を、先着した側がここで取り消す。
    ' 取り消せないままだと (a) 起動中のDoEvents中に割り込んで二重初期化、
    ' (b) 発火前にブックを閉じると数秒後に勝手に開き直して保存までされる、
    ' という2つの事故になる(レビュー H-1)。
    CancelPendingInstallerBoot

    ' 再入ガード。Boot の途中の DoEvents から Boot が呼ばれても、
    ' 2本目は何もせずに戻る。gBootDone は最後に立つので守りにならない。
    If mBootRunning Then Exit Sub

    If gBootDone Then
        ' 2026-07-28(レビュー H-2): この再実行パスは Hub 化以前のまま
        ' 取り残されており、旧ホームUI(btn_/lbl_)だけを組み直して
        ' 抜けていた。初回パスは EnsureLayout の上から EnsureHubLayout で
        ' 描き替える規約なので、ここだけ通ると Hub タイルの上に旧「質問する」
        ' 画面が重なる。既知の「タブが常時表示されない」「位置がずれる」は
        ' これが最有力(EnsureHubLayout を通ると直るので「再描画で直る」とも
        ' 症状が一致する)。初回パスと同じ順序に揃える。
        On Error Resume Next
        modUIMain.EnsureLayout
        modHub.EnsureHubLayout
        On Error GoTo 0
        On Error Resume Next
        modUIShelf.EnsureLayout
        On Error GoTo 0
        Exit Sub
    End If

    mBootRunning = True

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

    ' 1.5) 端末チェック(config allowed_domain が設定された組織だけ有効)。
    '      許可外の端末では知識を消して案内だけ出す。私物PCへコピーされた
    '      場合の一次防御。設定が空なら何もしない=既定では誰の業務も止めない。
    bootStage = "端末の確認"
    On Error Resume Next
    If Not modGuard.CheckDomain() Then
        ' 2026-07-28(レビュー M-6): 1回の不一致では消さない。VPN未接続や
        ' 一時的なプロファイル不整合でも不一致になり得るのに、帰結が
        ' 予告なしの即ワイプでは誤判定のコストが高すぎる。
        ' 規定回数(既定3回)続けて不一致だったときだけ消す。
        If modGuard.DomainBlockShouldWipe() Then modGuard.WipeKnowledge
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        ' 2026-07-28(レビュー L-8): この脱出だけ gBootDone を立てずに
        ' 抜けていたため、二重Boot(H-1)を踏むと消去と案内が2回走った。
        ' 「起動としては終わった」ので、成功パスと同じくフラグを立てる。
        gBootDone = True
        mBootRunning = False
        modGuard.ShowDomainBlocked
        Exit Sub
    End If
    ' 通ったので不一致の連続回数はリセットする。
    modGuard.ClearDomainBlockStreak

    ' 社内ネットワークに長期間つながっていない端末は知識を失効させる
    ' (7日前から予告あり。黙って消さない)。
    modGuard.EnforceExpiry
    On Error GoTo Failed

    ' 1.7) 初期ナレッジの展開(初回のみ)。同梱パックを本棚へ写すだけなので
    '      チャンク分割もAPI呼び出しも起きない。開いた瞬間から質問できる
    '      状態を作るのが目的で、ここが遅いと意味が無い。
    bootStage = "初期ナレッジの取り込み"
    On Error Resume Next
    modSeed.EnsureSeedLoaded
    LogBootStageErrorIfAny bootStage
    On Error GoTo Failed

    ' 2) first-run: pack_author入力
    bootStage = "はじめの設定(名前の保存)"
    EnsureFirstRun

    ' 2.5) 統計の更新(streak_days/バッジ)。2026-07-30(R3是正・要件A):
    ' 従来はここより後、3画面のEnsureLayout(Hub描画を含む)が終わった後に
    ' TouchToday/EvaluateBadgesを呼んでいた。Hubは明示的に再描画されるまで
    ' そのままなので、初回描画の「連続ログイン」「みんな(今日/今月)」等の
    ' タイルとバッジ棚が1つ前の値のまま固まっていた(R3要件定義書 背景1)。
    ' pack_author入力(直前のEnsureFirstRun)の直後・Hub等のEnsureLayoutより
    ' 前にここへ移すことで、初回描画時点からstreak_daysと獲得済みバッジが
    ' 最新になる。またEvaluateBadgesの先頭にあるwelcomeバッジ(要件C)は
    ' pack_author登録が条件なので、EnsureFirstRunの直後に置くことで
    ' 「名前入力→(同じBoot内で)EvaluateBadges→ポップアップ→Hubのバッジ棚に
    ' 点灯」まで一続きに起きる(要件定義書 要件C)。
    bootStage = "統計の更新(連続利用日数・バッジ)"
    On Error Resume Next
    modStats.TouchToday
    LogBootStageErrorIfAny bootStage
    modStats.EvaluateBadges
    LogBootStageErrorIfAny bootStage
    On Error GoTo Failed

    ' 3) 3画面EnsureLayout(それぞれが内部でRenderShelf/RenderDashboardまで実行する)。
    ' 2026-07-16: 実機で「1画面の描画中の不具合がアプリ全体の起動を止める」
    ' 事例が立て続けに見つかった(ホーム→マイ本棚→…と直しても次の画面で
    ' 同種の失敗が起きる)。非エンジニアが使う配布物として、1画面の描画に
    ' 問題があってもアプリ自体は必ず開けるべきなので、3画面それぞれを
    ' 独立してOn Error Resume Nextで保護し、失敗はerr_logに詳細(uiStepまで
    ' 埋め込み済みのErr.Description)を記録した上で次の画面へ進む
    ' (該当画面だけ表示が崩れる可能性はあるが、アプリが開けないよりずっと良い)。
    ' ホームはHub画面(modHub)へ置き換えた。modUIMain.EnsureLayoutは
    ' mb_question等の名前定義とRenderAnswerの土台を用意する役割が残るため
    ' 先に実行し、その上からHubのレイアウトで描き替える(Hub側が旧btn_/lbl_
    ' Shapeも消すので二重表示にはならない)。
    bootStage = "ホーム画面の組み立て"
    On Error Resume Next
    modUIMain.EnsureLayout
    LogBootStageErrorIfAny bootStage
    On Error GoTo Failed

    ' 2026-07-31(レビュー R8 F3 / R8b B9): 起動ジッタは【共有I/Oを伴う最初の
    ' 処理の直前】へ置く。共有I/Oを最初に行うのは modHub.EnsureHubLayout
    ' (受信箱の描画で modShare.Reachable() を呼ぶ)なので、その1行手前が正しい。
    '
    ' B9: R8では modUIMain.EnsureLayout よりさらに前に置いていたが、この時点では
    ' まだ軽量マクロ無効ガードの案内シートがアクティブで、Nexus画面も
    ' ホーム画面も組み上がっていない。StartupJitter は中で DoEvents を回して
    ' 最大3秒待つため、【利用者にはガードシートが見えていて、しかも操作を
    ' 受け付けてしまう】3秒の窓ができていた。そこでシートを触られると、
    ' 起動処理と競合して原因不明の崩れ方をする。画面が組み上がってから散らす。
    ' F3の要件(最初の共有I/Oより前)は、EnsureHubLayout の直前なので満たす。
    On Error Resume Next
    modChannel.StartupJitter
    On Error GoTo Failed

    On Error Resume Next
    modHub.EnsureHubLayout
    LogBootStageErrorIfAny bootStage
    On Error GoTo Failed

    bootStage = "マイ本棚画面の組み立て"
    On Error Resume Next
    modUIShelf.EnsureLayout
    LogBootStageErrorIfAny bootStage
    On Error GoTo Failed

    ' 「ダッシュボード」シート(SH_DASH)の組み立ては削除した(2026-07-27)。
    ' このシートへ遷移するコードはソース全体に1行も無く、起動のたびに
    ' 誰も到達できない画面を作っていた。統計の表示先はHubとDashboard画面。
    bootStage = ""

    ' Wave4修正: modStats.TouchToday(streak_days/last_used_date更新)を
    ' どこからも呼んでいなかったため、streak7バッジもダッシュボードの
    ' 連続利用日数も永久に更新されない不具合があった。TouchToday/
    ' EvaluateBadgesの呼び出し自体は2.5)へ移した(要件A・上記コメント参照)。
    ' 1セッション1回(gBootDoneで守られたこの初回Boot経路)呼べば
    ' TouchTodayの契約(§7.5: 昨日なら+1/今日なら不変/それ以外リセット)は
    ' 成立する。

    ' 本棚が空なら回答エリアに常設案内を出す
    Dim shelfIsEmpty As Boolean
    shelfIsEmpty = False
    On Error Resume Next
    shelfIsEmpty = (modShelf.TotalChunks() = 0)
    On Error GoTo Failed
    If shelfIsEmpty Then
        On Error Resume Next
        modUIMain.ShowEmptyShelfHint
        On Error GoTo Failed
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
    On Error GoTo Failed
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
    ' 2026-07-28(レビュー M-23): 起動ジッタは【共有I/Oを始める前】に置く。
    ' 従来は SyncNow(その中で感謝状・品質報告の全ファイル走査をする)を
    ' 済ませてから散らしていたため、朝の一斉起動で最も重い処理が
    ' 全員同時刻に走っていた。分散させたい負荷の後で散らしても意味が無い。
    ' 2026-07-31(レビュー R8 F3): その呼び出しは 3) の直前へ移した
    ' (受信箱の描画がさらに前で共有を触っていたため)。1回で足りる。

    If modConfig.GetBool("sync_on_open", True) Then
        On Error Resume Next
        modShelfSync.SyncNow silent:=True
        On Error GoTo Failed
    Else
        ' sync_on_open=Falseでも、起動時にP2P感謝状・品質報告集計は一度回収する
        ' (sync_on_open=TrueのときはSyncNow内で回収済み=二重回収しない)。
        On Error Resume Next
        modP2P.CollectThanks silent:=True
        modP2P.CollectNoiseVotes silent:=True
        On Error GoTo Failed
    End If

    ' 共有知フライホイールの受信(ファイルコピーだけ=API呼び出しなし)。
    ' 本棚への取り込みは利用者がHubのお知らせを押したときだけ行う。
    ' 朝の一斉起動で共有フォルダへ全員が同時に殺到しないよう散らす。
    On Error Resume Next
    ' 旧タグ("pack:<作者名>")で入った部門チャンクの一度きりの掃除。
    ' 修正前のビルドで部門へつないだ端末には、消せない残骸が溜まっている
    ' (レビュー C-1)。掃除して版数を空にし、正しいタグで入れ直させる。
    ' 掃除が走ったことは usage_log("channel_origin_migration")に残る。
    modChannel.MigrateOriginNamespace
    modInsight.CollectInsights

    ' 共有フォルダに到達できたことを記録(端末失効タイマーのリセット)。
    '
    ' 2026-07-28(解説書 §11-10): 条件が「チャンネルが1件以上見つかったとき」
    ' だった。ListChannels は channels\ が無い、または version.txt を持つ
    ' サブフォルダが1つも無ければ空を返すため、
    '   【共有フォルダには正常に到達できているが、まだどの部門も正典を
    '     発行していない】
    ' という状態では TouchReach が呼ばれず、失効カウンタだけが進み続ける。
    ' knowledge_expire_days(既定30)日後に、部門の正典どころか
    ' 【利用者が自分で取り込んだ資料まで】消える。
    ' PoC の初期状態(共有パスは設定済み・正典は未発行)がまさにこの条件。
    '
    ' 失効タイマーが見ているのは「社内ネットワークに繋がっているか」であって
    ' 「正典が発行されているか」ではない。判定は共有フォルダのルートへ
    ' 到達できたかどうかにする。到達性は modShare が1セッション1回だけ
    ' 確かめてキャッシュするので、ここでの追加コストは無い。
    If modShare.Reachable() Then modGuard.TouchReach
    On Error GoTo Failed

    ' 部門チャンネルは「更新があるか」だけ見る(version.txtを読むだけ=軽い)。
    ' 実際の取り込みは埋め込みAPIを使うので、必ず利用者のクリックを待つ。
    ' 起動が遅いことが最大の離脱要因なので、ここで重い処理は絶対にしない。

    ' 6) ScheduleAutoSync
    On Error Resume Next
    modShelfSync.ScheduleAutoSync
    On Error GoTo Failed

    ' 7) 内部シート隠蔽
    bootStage = "内部シートの整理"
    HideInternalSheets
    RemoveOrphanDefaultSheets

    ' 8) Nexus UI(config nexus_ui=TRUEのとき新SPA UIを起動。失敗しても
    '    旧3画面は生きているため、起動自体は続行する)
    If modConfig.GetBool("nexus_ui", False) Then
        bootStage = "Nexus画面の起動"
        On Error Resume Next
        modApp.LaunchNexus
        ' 2026-07-21: このstageだけLogBootStageErrorIfAnyの呼び出しが漏れており、
        ' Nexus画面構築の失敗がerr_logに一切残らない盲点になっていた
        ' (実機で「Nexus画面が白紙」という報告のみでは原因を特定できない
        ' 状態が続いていた恒久対策)。Home/マイ本棚と同じ扱いに揃える。
        LogBootStageErrorIfAny bootStage
        On Error GoTo Failed
        bootStage = ""
    End If

    ' 8.5) 軽量マクロ無効ガード(D1): 起動が成功したので案内シートを隠す
    '      (Nexus起動後=別シートがアクティブな状態で隠すため確実に隠れる)。
    On Error Resume Next
    HideGuardSheet
    On Error GoTo Failed

    ' 8.6) ホットキー登録: Ctrl+Shift+Q=一撃召喚 / Ctrl+Enter=送信(Nexus上のみ発火)。
    '      解除はAuto_Close(既存のCtrl+Z解除と同じライフサイクル)。
    On Error Resume Next
    Application.OnKey "^+q", "modApp.SummonNexus"
    ' 2026-07-28(レビュー L-19): Ctrl+Enter を2通り登録する。
    ' "^~" はメインキーの Enter しか拾わないため、テンキーの Enter で
    ' 送信できなかった。"^{ENTER}" を足して両方拾う。
    ' なお、セル編集中(入力欄に文字を打っている最中)は OnKey が効かず、
    ' 1回目の Ctrl+Enter は「確定」になる。これはExcelの仕様で回避できない
    ' ため、ヒント文は「入力後に Ctrl+Enter」と書き換えてある。
    Application.OnKey "^~", "modApp.HotSend"
    Application.OnKey "^{ENTER}", "modApp.HotSend"
    On Error GoTo Failed

    ' 2026-07-31(レビュー R8 F3): ここから先の再描画では、Hubの受信箱が
    ' 共有フォルダへ実際に問い合わせてよい。起動シーケンス中は
    ' 「押して確認」のプレースホルダで済ませ、共有I/Oを1回も走らせない。
    On Error Resume Next
    modHubStat.AllowShareQueries
    On Error GoTo Failed

    gBootDone = True
    mBootRunning = False
    On Error Resume Next
    Application.EnableEvents = True   ' 起動中に抑止したイベントを復帰
    ' 3画面EnsureLayout/LaunchNexusはResume Next保護下で中断し得るため、暗転
    ' (ScreenUpdating=False)のまま起動完了する経路をここで確実に塞ぐ。
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub

Failed:
    mBootRunning = False   ' 失敗しても必ず降ろす(降ろし忘れると次回が素通りする)
    ' EnsureLayout途中で落ちるとScreenUpdating=Falseのまま画面が固まって
    ' 見えるため、必ず戻す(ダイアログより先に)。
    Dim failDesc As String
    failDesc = Err.Description
    Dim failNum As Long
    failNum = Err.Number
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailedCleanup2
FailedCleanup2:
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
    ' インストーラが予約した「+1秒後のBoot」がまだ残っていたら取り消す。
    ' 残したまま閉じると、数秒後にExcelがこのブックを勝手に開き直し、
    ' インストーラが走って保存までしてしまう(レビュー H-1)。
    CancelPendingInstallerBoot
    On Error GoTo 0

    ' Nexusが隠したネイティブUI(リボン等)を必ず復元する
    On Error Resume Next
    modUI.RestoreExcelUI
    Application.OnKey "^+q"   ' ホットキーも既定へ戻す(残存するとブック閉鎖後にエラー)
    Application.OnKey "^~"
    Application.OnKey "^{ENTER}"
    On Error GoTo 0

    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0

    ' 利用データの送信は終了時に行う(起動時にやると朝の一斉アクセスと
    ' 重なるうえ、起動が遅くなる。終了時なら数秒かかっても業務を止めない)。
    On Error Resume Next
    modTelemetry.Publish
    On Error GoTo 0

    ' Ctrl+Break等でBoot/SyncNow途中のEnableEvents=False焼き付きが起きても、
    ' ブックを閉じれば必ずここで復帰させる(次回セッションへ持ち越さない)。
    On Error Resume Next
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' RunFirstRunPromptEarly - 自己インストーラのWorkbook_Open内、OnTimeでBootを
'   予約する前に同期呼び出しされる(2026-07-21実機対応)。名前入力ダイアログ
'   (EnsureFirstRun内のInputBox)がOnTime経由のBoot内で表示されると、その
'   直後からws.Activateが実機で安定して失敗する現象が全ラウンドのログで
'   一貫して観測されたため、ダイアログの発生タイミングをWorkbook_Openの
'   通常のイベントコンテキスト内に変える対策。EnsureFirstRun自体はpack_author
'   設定済みなら即returnする冪等な関数なので、Boot側の既存呼び出しは
'   フォールバックとしてそのまま残す(二重表示にはならない)。
Public Sub RunFirstRunPromptEarly()
    On Error Resume Next
    modConfig.EnsureLoaded
    EnsureFirstRun
    On Error GoTo 0
End Sub

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

    ' 到達導線が無く中身も描かなくなったので隠す(空の可視シートを残さない)。
    HideSheetSafely modAppDef.SH_DASH, HIDDEN
    HideSheetSafely modAppDef.SH_CONFIG, HIDDEN
    HideSheetSafely modAppDef.SH_MANIFEST, HIDDEN
    HideSheetSafely modAppDef.SH_STATS, HIDDEN
    HideSheetSafely modAppDef.SH_USAGE, HIDDEN
    HideSheetSafely modAppDef.SH_ERRLOG, HIDDEN
End Sub

' 2026-07-21実機対応: 原因未特定の「Sheet2のような既定名の空白シートが
' 混入する」実機報告への当座の後始末。本アプリはビルド時・実行時とも
' "Sheet"+数字という既定名のシートを作らない設計のため、その名前パターンに
' 一致し、かつ本当に空(セル内容もShapeも無い)シートだけを安全に削除する
' (取りこぼしても実害はないので、判定条件は厳しめに倒す。原因特定のため
'  削除前にerr_logへ記録する)。
Private Sub RemoveOrphanDefaultSheets()
    On Error Resume Next
    Dim i As Long
    For i = ThisWorkbook.Worksheets.count To 1 Step -1
        Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(i)
        If ws.Name Like "Sheet[0-9]*" Then
            Dim isEmptySheet As Boolean
            isEmptySheet = (ws.UsedRange.Address = "$A$1") And _
                           (LenB(CStr(ws.Range("A1").Value)) = 0) And (ws.Shapes.count = 0)
            If isEmptySheet Then
                modLog.LogError "E0801", "modBoot.Boot", _
                    "原因不明の既定名シートを検出・削除: " & ws.Name
                Application.DisplayAlerts = False
                ws.Delete
                Application.DisplayAlerts = True
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

' 盲点D1(軽量マクロ無効ガード)の実行時側。起動が成功したのでマクロ有効化の
' 案内シートを隠す(役目を終えた)。アクティブシートは隠せない仕様のため、隠す前に
' 別の可視シートへフォーカスを移してから隠す。マクロ無効で開かれた場合はBoot自体が
' 動かずこの経路に到達しないので、案内は表示されたまま=壊れたUIの代わりに案内が見える。
Private Sub HideGuardSheet()
    On Error Resume Next
    Dim gs As Worksheet
    Set gs = ThisWorkbook.Worksheets(GUARD_SHEET)
    If gs Is Nothing Then Exit Sub
    If ActiveSheet Is gs Then
        Dim other As Worksheet
        For Each other In ThisWorkbook.Worksheets
            If Not (other Is gs) Then
                If other.Visible = -1 Then   ' xlSheetVisible
                    If Not modUI.ActivateSheetRobust(other, "modBoot.Boot") Then _
                        modLog.LogError "E0801", "modBoot.Boot", "案内シート退避のActivateに失敗"
                    Exit For
                End If
            End If
        Next other
    End If
    gs.Visible = 0   ' xlSheetHidden
    On Error GoTo 0
End Sub

Private Sub HideSheetSafely(ByVal sheetName As String, ByVal visibility As Long)
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Visible = visibility
    On Error GoTo 0
End Sub
