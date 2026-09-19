Attribute VB_Name = "modBoard"
Option Explicit

' ============================================================================
' modBoard - チーム連帯ボード(組織全体の節約時間+感謝ベース称号)。疎結合プラグイン。
' ----------------------------------------------------------------------------
' 役割:
'   1) statsビーコン: 自分の統計(感謝受領数/節約時間 日・月・年)を共有フォルダ
'      <nexus_share_path>\board\stats_<自分hash>.txt へ発信する(1人1ファイル上書き
'      =書込競合ゼロ。noise投票と同じ実証済みパターン)。
'   2) 集計: 組織全体の「今日/今月/今年の節約時間」。2026-08-16(R33 W6-1)から
'      【集約スナップショット方式】= 発行者端末だけが全ビーコンを読んで
'      <share>\board\summary.txt を1本書き、他の全端末はそれを1本読むだけ。
'   3) サイドバー常設ウィジェット: 「🌍 みんなの節約」+月間目標プログレスバー。
'      クリックで履歴ポップアップ(個人の直近7日+月/年、組織合計)。
'   4) 称号: 感謝受領数に応じた絶対的な称号(💡5件以上/🌟20件以上)。自己申告不可の
'      P2P感謝状のみが源泉のため偽装不能。modMentorが専門家名の頭に付ける。
'
' 設計判断:
'   ・節約時間の実データは modStats の日付キー(sv:d:yyyymmdd / sv:m:yyyymm /
'     sv:y:yyyy)。✅解決(modAsk.FeedbackGreen)時に1回ぶん(分数は
'     modP2PIo.MinutesPerSelfsolve。config minutes_per_selfsolve、既定15)をBumpする(発火点は
'     modAsk内=多重防止ガードの内側)。日付キー方式なので「日/月/年を跨いだら
'     リセット」はロジック不要で自動成立し、過去キーがそのまま履歴になる。
'   ・ビーコンには自分の dayKey/monKey/yearKey を明記し、集計側は「現在の同じ
'     キーと一致する値だけ」を合算する(古いビーコンの昨日の値を今日に混ぜない)。
'   ・エントリポイント(BootBoard/クリック)はサーキットブレーカーで全障害を握り、
'     共有フォルダ未達なら個人統計のみで描画(fail-soft)。
'   ・Shape名は nx_sb_stat*(サイドバー系prefix)= 既存のZ-Order維持/テーマ再彩色
'     ループが追加コード無しで面倒を見てくれる。modUIは1バイトも変更しない。
'   ・I/OヘルパーはmodMentorと同仕様の軽量複製(モジュール完全隔離の意図的重複)。
'
' 残っている制約(R33 W6-1。旧「本格展開時の設計課題」の後継として正直に残す):
'   ・書けるのは発行者用ブックの端末だけ。発行者が長期不在だと集計は更新されず、
'     SUMMARY_MAX_AGE_HOURS を過ぎたら数字ではなく「集計はまだありません」を
'     出す(古い数字を今日の数字として出し続けない)。運用では発行者用ブックを
'     日常的に開く端末へ1台置くこと。
'   ・1回に読むのは SCAN_CAP 本まで=全社規模では概算(画面とログに母数つきで
'     明示)。ただし窓は日ごとに移動し称号は引き継ぐので、誰も恒久的には
'     外れない(R33H F18。仕組みは modShare.BoardScanStart 直上の節)。
' ============================================================================

Private Const BOARD_SUBDIR As String = "board"
Private Const WIDGET_TOP As Double = 386     ' サイドバー内の縦位置(nav6項目の末端374の下)
' R13-7d: 15分換算の重複定数はmodP2PIo.MinutesPerSelfsolve()へ統合したため、
' ここのPrivate Const(未使用のまま放置されていた複製)は削除した。

' ビーコン集計キャッシュ(セッション内。共有フォルダ再走査はBootBoard時のみ)
Private mOrgDay As Long, mOrgMon As Long, mOrgYear As Long
Private mTitles As Object    ' Dictionary: id(小文字) -> thanks受領数
Private mLoaded As Boolean
' R13-7c: 自分の部(DeptOf)一致ビーコンだけの今月合算と、その元の
' 自分のチーム/部コード。myDeptが空(チーム不明)のときはポップアップへ出さない。
Private mDeptMon As Long
Private mMyDept As String
' BootBoard を通ったか(2026-07-30 レビュー2-C)。起動シーケンス中の
' Hub初期描画から共有フォルダI/Oを走らせないためのゲート。
Private mBooted As Boolean

' ビーコン再発信のスロットル(2026-07-31 R8 F8)と、集計のTTL(同 F9)。
' どちらも Timer(0:00からの経過秒)で持つ。判定は modShareRule.CacheIsFresh
' に寄せてあり、日跨ぎで Timer が0へ戻る場合も「古い」に倒れる。
Private mBeaconAt As Double
Private mAggAt As Double
' 前回発信したビーコンの中身(送信時刻を除く)。R8b B11: 中身が変わったかで
' スロットル間隔を切り替えるために保持する。
Private mLastBeaconData As String
Private Const BEACON_IDLE_SEC As Double = 600#    ' 中身が同じときの最短間隔
Private Const BEACON_FRESH_SEC As Double = 60#    ' 中身が変わったときの最短間隔
Private Const AGG_TTL_SEC As Double = 600#

' 集約スナップショット(2026-08-16 R33 W6-1)。SCAN_CAP=発行者端末が1回に開く
' ビーコンの上限。SUMMARY_MAX_AGE_HOURS=これより古い集計は数字として使わない。
' FORCE_MIN_SEC=「更新」ボタンの下限間隔。TITLE_MIN=称号が付く最小の感謝
' 受領数(TitleForの下段と同じ値。スナップショットへはこれ以上の人だけ載せる)。
Private Const SCAN_CAP As Long = 500
Private Const SUMMARY_MAX_AGE_HOURS As Long = 24
Private Const FORCE_MIN_SEC As Double = 60#
Private Const TITLE_MIN As Long = 5

' 「更新」ボタンの下限間隔の起点(Timer)。
Private mForceAt As Double
' いま画面に出している集計の状態。"ok"(使える)/"stale"(古すぎる)/
' "broken"(読めない)/"none"(まだ誰も書いていない)/""(未取得)。
' 数字を画面へ出してよいのは "ok" のときだけ。
Private mAggState As String
Private mAggStamp As String     ' スナップショットの作成時刻(ISO)
Private mAggUsers As Long       ' 集計に入った人数
Private mAggTotal As Long       ' 在ったビーコンの本数(=母数。R33H F18)
Private mAggApprox As Boolean   ' SCAN_CAPで打ち切った概算か
Private mLoggedState As String  ' usage_logへ同じ状態を毎回積まないための印

' ----------------------------------------------------------------------------
' BootBoard - エントリポイント(LaunchNexus末尾から1行フック)。
'   ローカル処理→(共有到達なら)発信→集計。全障害を内部で握る。
'
' 処理の順番について(2026-07-31 R8 F2・司令塔裁定):
'   F2の本旨は「関所を通らずに共有パスへ触らない」であって
'   「共有が無い端末のローカルUIを止める」ではない。
'   DrawWidget(旧Shapeの掃除)と ShowWeeklySummary(先週の節約時間を
'   ねぎらうToast)は my_stats を読むだけで共有フォルダに一切触らない。
'   ここを Reachable() の後ろに置くと、共有未接続の端末では
'   【週次のねぎらいが永久に出ない・旧Shapeが永久に残る】という
'   到達性とは無関係の機能欠落になる。
'   よって「ローカル処理は無条件 → 最初の共有I/Oの直前でガード」の順にする。
' ----------------------------------------------------------------------------
Public Sub BootBoard()
    On Error Resume Next   ' 安全弁: 本機能の失敗を起動へ絶対に波及させない

    ' R29H F2b: modIntegrity.WarnAtStartupはInitUIより前(modBoot)に走るため
    ' 保留していた案内バブルを、InitUI後に到達する最初の地点(ここ)で出す。
    modIntegrity.FlushPendingBubbles

    ' --- ここから下は共有フォルダに触らない(到達性に関係なく必ず実行する)---
    DrawWidget             ' 旧サイドバーShapeの掃除(ローカルのみ)
    ShowWeeklySummary      ' B-5: 週の初回起動時だけ、先週の節約時間を労いバブル

    ' 起動シーケンスを通り抜けた印。ここから先の再描画では
    ' RefreshBoardTiles が動いてよい(レビュー2-C)。共有へ届くかどうかとは
    ' 別の話なので、ガードより前で立てる。
    mBooted = True

    ' --- ここから下が共有フォルダI/O。届かないならまとめて見送る ---
    ' 発信も集計も全部が共有I/Oで、届かない共有に対してはリトライ
    ' (3回×バックオフ)を全部払ってから空振りする。到達判定は modShare が
    ' 1セッション1回だけキャッシュ付きで行うので追加費用は無い。
    If Not modShare.Reachable() Then Exit Sub
    PublishBeacon
    RefreshBoard
    On Error GoTo 0
End Sub

' 週次サマリー(B-5): その週の初回起動時に1回だけ、先週の個人合計をToastで返す。
' ガードはISO風の年+週番号キー(wk:yyyww)。sv:d:日付キーを7日分読むだけ。
Private Sub ShowWeeklySummary()
    Dim wkKey As String
    wkKey = "wk:" & modUtilText.IsoYear(Date) & Format$(DatePart("ww", Date, vbMonday), "00")
    If modStats.GetStat(wkKey) > 0 Then Exit Sub
    modStats.Bump wkKey

    ' 先週(直近の月曜の7日前～日曜)の個人合計
    Dim mon As Date: mon = Date - Weekday(Date, vbMonday) + 1   ' 今週の月曜
    Dim total As Long, i As Long
    For i = 1 To 7
        ' R12-H-3: 和暦カレンダー端末では Format$(d,"yyyymmdd") が元号年を返す。
        ' 書く側(modAsk)は既に modUtilText.IsoDateCompact なので、読む側が
        ' Format$ のままだと【一致しないキーを引いて常に0分】になっていた。
        total = total + MyMin("d", modUtilText.IsoDateCompact(mon - i))
    Next i
    If total <= 0 Then Exit Sub   ' ゼロ週は何も言わない(空虚な自慢をしない)

    modUI.AddChatBubble "ai", "先週、あなたはこのツールで " & FmtMin(total) & " を節約しました。今週もいいスタートを。"   ' R29H F2b
End Sub

' ----------------------------------------------------------------------------
' TitleFor - 感謝ベース称号(偽装不可)。💡=5件以上 / 🌟=20件以上 / それ以外""。
'   自分自身はローカル統計から、他者はビーコンから引く。
' ----------------------------------------------------------------------------
Public Function TitleFor(ByVal userId As String) As String
    On Error Resume Next
    Dim n As Long: n = -1
    Dim myId As String: myId = modP2P.CurrentUserId()
    If StrComp(userId, myId, vbTextCompare) = 0 Then
        n = modStats.GetStat("thanks_received_total")
    ElseIf mLoaded Then
        ' R33H F30: 書く側と同じ IdKey で引く(LCase だけでは一致しない)。
        Dim k As String: k = modP2PIo.IdKey(userId)
        If mTitles.Exists(k) Then n = CLng(mTitles(k))
    End If
    If n >= 20 Then
        TitleFor = ChrW(&HD83C) & ChrW(&HDF1F) & " "        ' U+1F31F 星
    ElseIf n >= TITLE_MIN Then
        TitleFor = ChrW(&HD83D) & ChrW(&HDCA1) & " "        ' U+1F4A1 電球
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' OnWidgetClick - ウィジェットクリック → 履歴ポップアップ(個人7日+月/年+組織)。
' ----------------------------------------------------------------------------
Public Sub OnWidgetClick()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    HideHistory

    ' R20-4c(実機第7報③④): 共有フォルダが未設定のときは「0分」の空虚な
    ' 履歴を見せるのではなく、その場で設定できるように誘導する。
    Dim shareSet As Boolean
    On Error Resume Next
    shareSet = (LenB(modShare.BasePath()) > 0)
    On Error GoTo Done
    If Not shareSet Then
        modUiLock.Leave
        If MsgBox("共有フォルダが未設定です。今設定しますか?", _
                  vbYesNo + vbQuestion, modAppDef.APP_NAME) = vbYes Then
            On Error Resume Next
            modHelp.OnShareSetup
            On Error GoTo 0
        End If
        Exit Sub
    End If

    ' 2026-07-26: 呼び出し元がHubの統計タイル(ホームシート)へ移ったのに、
    ' ここは常にNexusシートへ描いていた。別シートに描かれるので画面上は
    ' 「押しても何も出ない」状態になる。今アクティブなシートへ描く。
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    If ws Is Nothing Then GoTo Done

    Dim body As String
    body = ChrW(&HD83D) & ChrW(&HDCC8) & " 節約時間レポート" & vbLf & vbLf & _
           "【あなた】" & vbLf & _
           "  今日: " & FmtMin(MyMin("d", modUtilText.IsoDateCompact(Date))) & _
           "  /  今月: " & FmtMin(MyMin("m", modUtilText.IsoYm(Date))) & _
           "  /  今年: " & FmtMin(MyMin("y", modUtilText.IsoYear(Date))) & vbLf & vbLf & _
           "【直近7日の履歴】" & vbLf & History7() & vbLf & _
           "【みんな(組織全体)】" & vbLf & OrgBlockForPopup() & vbLf & vbLf & _
           "※「" & ChrW(&H2705) & "解決した」1回=" & modP2PIo.MinutesPerSelfsolve() & _
           "分の節約として、他者からの感謝と同じしくみで" & vbLf & _
           "  組織に共有・合算されます(集計は発行者用ブックの端末が作るので" & vbLf & _
           "  数十分から半日ほど遅れることがあります)。" & vbLf & _
           "(クリックで閉じる)"

    ' 位置は画面の見えている範囲から中央寄せ(旧サイドバー幅を前提にした
    ' x=260の決め打ちだと、サイドバー廃止後は左に寄りすぎる)。
    Dim popW As Double: popW = 470
    Dim popL As Double: popL = 40
    Dim popT As Double: popT = 90
    On Error Resume Next
    popL = ActiveWindow.VisibleRange.Left + (ActiveWindow.VisibleRange.Width - popW) / 2
    popT = ActiveWindow.VisibleRange.Top + 80
    On Error GoTo Done
    If popL < 8 Then popL = 8

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, popL, popT, popW, 60)
    ' nx_hub_ 接頭辞にしておくと、Hubの再描画(RemoveHubShapes)が
    ' 消し忘れのポップアップを自動で片付けてくれる。
    shp.Name = "nx_hub_hist"
    shp.Adjustments(1) = 0.06
    shp.Line.Visible = -1
    shp.Line.Weight = 1#
    shp.Line.ForeColor.RGB = modUI.UiColor("primary")
    shp.Fill.ForeColor.RGB = modUI.UiColor("surface")
    With shp.TextFrame2
        .WordWrap = -1
        .AutoSize = 1
        .MarginLeft = 16: .MarginRight = 16: .MarginTop = 10: .MarginBottom = 10
        .TextRange.Text = body
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 9.5
    End With
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")
    shp.OnAction = "modBoard.OnHistoryClose"
    shp.Placement = 3
    modSkin.ApplySoftShadow shp
    shp.ZOrder 0
Done:
    modUiLock.Leave
End Sub

Public Sub OnHistoryClose()
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    HideHistory
    On Error GoTo 0
    modUiLock.Leave
End Sub

' ポップアップはアクティブなシートに描かれるため、Hub/チャットの両方から消す。
Private Sub HideHistory()
    On Error Resume Next
    ThisWorkbook.Worksheets("Nexus").Shapes("nx_hub_hist").Delete
    ThisWorkbook.Worksheets("Nexus").Shapes("nx_sb_stat_hist").Delete   ' 旧名の残骸
    ThisWorkbook.Worksheets(modAppDef.SH_HOME).Shapes("nx_hub_hist").Delete
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 内部: ビーコン発信/集計
' ----------------------------------------------------------------------------
' ----------------------------------------------------------------------------
' PublishBeacon - 自分の統計を共有フォルダへ1ファイルで置く(上書き=競合ゼロ)。
'
' 2026-07-31(レビュー R8 F8): Public にして、「解決した」の直後にも呼べるようにした。
'   従来は起動時(BootBoard)にしか発信していなかった。ところが「今日の節約
'   時間」は【その日に「解決した」を押した分】で、押すのは起動よりずっと後になる。
'   つまり全員のビーコンが「今日 0分」の状態で置かれ、翌朝の起動で上書き
'   されるまで誰の画面でも「みんなの節約(今日)」が構造的にほぼ0だった。
'   目玉の指標が常に0では、共有そのものが動いていないように見える。
'
'   ただし「解決した」は1日に何度も押されるので、そのたび共有へ書くと人数×回数の
'   書込みになる。そこでスロットルを掛けるが、間隔は【中身が変わったか】で
'   分ける(2026-07-31 R8b B11):
'     ・中身が前回と同じ  … 10分に1回(定期的な生存通知。急ぐ理由が無い)
'     ・中身が変わった    … 60秒に1回(「解決した」を押した直後。ここが目玉の指標)
'   一律10分だと、起動から10分以内に押した分が丸ごと落ちる。朝いちばんに
'   使う人ほど反映されないという、いちばん見られたくない挙動になっていた。
'   比較には送信時刻(NowStamp)を含めない。含めると毎回「変わった」になる。
' ----------------------------------------------------------------------------
Public Sub PublishBeacon()
    ' 共有へ届かないなら何もしない。届かない共有への書込みリトライ
    ' (3回×バックオフ)を「解決した」のたびに払わせない(R8 F2/F8)。
    If Not modShare.Reachable() Then Exit Sub

    Dim myId As String
    On Error Resume Next
    myId = modP2P.CurrentUserId()
    On Error GoTo 0
    If LenB(myId) = 0 Then Exit Sub

    ' 統計の読み出しは my_stats(ローカル)なので、スロットル判定の前に
    ' 済ませてよい。共有I/Oは下の MkDir/WriteBeacon だけ。
    Dim dk As String: dk = modUtilText.IsoDateCompact(Date)
    Dim mk As String: mk = modUtilText.IsoYm(Date)
    Dim yk As String: yk = modUtilText.IsoYear(Date)
    ' R13-7c: team列を末尾に足す(旧読み手は列数を見ないので無害、新読み手は
    ' modP2PIo.BeaconTeamFieldがUBoundで分岐する)。組み立ては純ロジックとして
    ' modP2PIoへ切り出してあるので、フィールド順の間違いをLOテストで固定できる。
    Dim dataText As String
    dataText = modP2PIo.BeaconDataText(myId, modStats.GetStat("thanks_received_total"), _
        dk, MyMin("d", dk), mk, MyMin("m", mk), yk, MyMin("y", yk), MyTeamCode())

    Dim ttlSec As Double
    If StrComp(dataText, mLastBeaconData, vbBinaryCompare) = 0 Then
        ttlSec = BEACON_IDLE_SEC       ' 中身が同じ = 急がない
    Else
        ttlSec = BEACON_FRESH_SEC      ' 中身が変わった = 早く届ける
    End If
    If modShareRule.CacheIsFresh(mBeaconAt, Timer, ttlSec) Then Exit Sub

    Dim folderPath As String: folderPath = BoardDir()
    If LenB(folderPath) = 0 Then Exit Sub
    On Error Resume Next
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    On Error GoTo 0

    ' スロットルの起点は「実際に書きに行った時刻」。書けたかどうかに関わらず
    ' 進める(届かない共有へ短い間隔で挑み続けないため)。
    mBeaconAt = Timer

    ' 2026-07-31(C4): 「前回送った中身」は【書けたときだけ】更新する。
    ' 送る前に更新すると、書込みに失敗したビーコンが「送信済み」扱いになり、
    ' 次回は中身が同じ = 10分待ちの側へ倒れる。失敗したのに10分間だまって
    ' 再送しない、が一番まずい(その間の「みんなの節約(今日)」は古いまま)。
    ' 失敗のままなら中身は前回と違うままなので、60秒後にもう一度挑戦する。
    If WriteBeacon(folderPath & "stats_" & modUtil.Fnv1a64Hex(myId) & ".txt", _
                   dataText & vbTab & modUtil.NowStamp()) Then
        mLastBeaconData = dataText
    End If
End Sub

' RefreshBoardTiles - 要件D(2026-07-30 R3)。modHub.EnsureHubLayoutが
'   タイル描画前に「みんな(今日/今月)」を鮮度良くするために呼ぶ公開口
'   (RefreshBoard自体はPrivateなのでモジュール外から直接は呼べない)。
'   RefreshBoardは集計をやり直すだけ(mOrgDay等の再計算)で副作用が無い
'   ので、BootBoard(LaunchNexus内=EnsureHubLayoutよりさらに後に走る)から
'   の既存呼び出しと重複しても壊れない(冪等)。呼び出し元(modHub)側で
'   modShare.Reachable()ガード+On Error保護を必ず付けること
'   (共有フォルダ未設定・到達不能時にブロッキングしないため)。
'   2026-07-30(レビュー2-C): BootBoard を通るまでは何もしない。
'   Boot 中の Hub 初期描画(modBoot→EnsureHubLayout)からここが呼ばれると、
'   起動を軽くするための StartupJitter より前に共有フォルダの同期I/Oが
'   走ってしまい、遅い/届かない共有フォルダでは起動そのものが待たされる
'   (LAN外・VPN未接続では Dir がタイムアウトするまで固まる)。
'   初期描画は前回セッションの値(=従来どおり)のままにし、
'   LaunchNexus 内の BootBoard 以降の再描画から鮮度を取りに行く。
Public Sub RefreshBoardTiles()
    If Not mBooted Then Exit Sub
    RefreshBoard
End Sub

' ----------------------------------------------------------------------------
' ForceRefreshBoard - TTLを無視して集計をやり直す(2026-07-31 R8 F9)。
'   ダッシュボードの「更新」ボタンだけがこれを呼ぶ。R8 F9 の裁定は
'   「押しても値が変わらないのは、ボタンが壊れているのと区別が付かない」。
'   R33 W6-1 で下限間隔(FORCE_MIN_SEC=60秒)を足したが、その中で【黙って
'   Exit する】形にしたため、その裁定を真下で破っていた(R33H F19)。
'   直し方は2つ:
'   (a) 60秒以内なら理由を1行トーストで返す(何も起きない、を無くす)。
'   (b) 下限の起点は【実際に集計をやり直したとき】だけ進める。共有が
'       瞬断していた1回で下限を使い切ると、復帰しても60秒押せなかった。
' ----------------------------------------------------------------------------
Public Sub ForceRefreshBoard()
    If modShareRule.CacheIsFresh(mForceAt, Timer, FORCE_MIN_SEC) Then
        On Error Resume Next
        modSkin.ShowToast modShare.BoardForceWaitText(CLng(FORCE_MIN_SEC)), "info", True
        On Error GoTo 0
        Exit Sub
    End If
    mAggAt = 0
    If RefreshBoard() Then mForceAt = Timer
End Sub

' ----------------------------------------------------------------------------
' RefreshBoard - 組織合計を出し直す(集約スナップショット方式)。
'
' 2026-07-31(レビュー R8 F9): 10分のTTLを付けた。Hubは画面を戻るたびに
'   描き直されるので、TTLが無いと「画面を行き来した回数 × 人数」が
'   そのままファイルサーバへの往復になる。値は分単位なので10分古くてよい。
'
' 2026-08-16(R33 W6-1・全社12,000人の前提条件): 全員が全員のビーコンを開く
'   形をやめた。旧実装はファイル数=利用者数(消す処理が無く退職者ぶんも増える)
'   を1本ずつ開いており、12,000人なら1回の集計で12,000オープンが走る。
'   発行者端末(modPublish.CanPublish。共有の寿命管理を任せている modInsightIo
'   と同じ区別)だけが全ビーコンを読み summary.txt を1本書き(BuildSnapshot)、
'   他の全端末はそれを1本読む(LoadSnapshot)。無い/壊れている/古すぎるときも
'   【全件走査へは戻さない】(戻せば消した当のコストが復活する)。
'   TTLとの関係: AGG_TTL_SEC=共有へ【触りに行く間隔】/ SUMMARY_MAX_AGE_*=
'   読めた【中身を信じてよいか】。軸が違うので条件が絡まない。
' ----------------------------------------------------------------------------
'   戻り値(R33H F19): 実際に集計をやり直したか。「更新」ボタンの下限間隔の
'   起点を、走ってもいない回で進めないために要る。
Private Function RefreshBoard() As Boolean
    ' TTL内なら前回の集計値をそのまま使う(mOrgDay等は消さない)。
    If modShareRule.CacheIsFresh(mAggAt, Timer, AGG_TTL_SEC) Then Exit Function

    ' 2026-07-31(R8b B8): ゼロ化・TTLの起点更新は【集計に入れると決めてから】。
    ' 従来はここより前で mOrgDay 等を0にし mAggAt も進めていたため、
    ' 共有が一時的に届かない(BoardDir が空)1回の呼び出しだけで
    '   ・「みんなの節約」が 0分 / 0分 になる
    '   ・他人の称号(mTitles)が全部消える
    ' という表示になり、しかも mAggAt が進んでいるので【10分間そのまま】
    ' 固定された。共有はすぐ復帰しているのに画面だけが壊れて見える。
    ' 集計できないときは、古い値を出し続ける方がずっとましなので何も触らない。
    Dim folderPath As String: folderPath = BoardDir()
    If LenB(folderPath) = 0 Then Exit Function
    If LenB(Dir(folderPath, vbDirectory)) = 0 Then Exit Function

    mAggAt = Timer
    RefreshBoard = True

    Dim isPub As Boolean
    On Error Resume Next
    isPub = modPublish.CanPublish()
    On Error GoTo 0

    If isPub Then
        BuildSnapshot folderPath
    Else
        LoadSnapshot folderPath
    End If
End Function

' BuildSnapshot - 発行者端末だけが通る道。ビーコンを読んで組織合計を出し、
'   結果を summary.txt へ1本書く。1回に開くのは SCAN_CAP 本までで、超える
'   ときは【日ごとに窓をずらして】全員が順に入る(R33H F18。恒久除外を作らない
'   ことが最低条件。仕組みと根拠は modShare.BoardScanStart 直上の節)。
'   500本ぶんの待ちを払うのは発行者1台だけで、他の全端末は1本読みで済む。
Private Sub BuildSnapshot(ByVal folderPath As String)
    Dim dk As String: dk = modUtilText.IsoDateCompact(Date)
    Dim mk As String: mk = modUtilText.IsoYm(Date)
    Dim yk As String: yk = modUtilText.IsoYear(Date)
    mMyDept = modP2PIo.DeptOf(MyTeamCode())   ' R13-7c: config優先→userIdから推定

    Dim orgD As Long, orgM As Long, orgY As Long
    Dim usersN As Long   ' 実際に読めて合算できた人数(開いた本数ではない)
    Dim deptAgg As Object: Set deptAgg = CreateObject("Scripting.Dictionary")
    Dim titleAgg As Object: Set titleAgg = CreateObject("Scripting.Dictionary")

    ' R33H F18: 称号は前回のぶんを引き継いでから今回ぶんで上書きする。窓に
    ' 入らなかった日に称号が消えないため(受領数は減らない量なので上書きで
    ' 事実に合う)。Dir を使うのでビーコンの列挙より【前】に済ませる。
    Dim carried As Long: carried = modShare.BoardCarryTitles(folderPath, titleAgg)

    Dim totalN As Long
    Dim names() As String
    names = Split(modShare.BoardListBeacons(folderPath, SCAN_CAP, totalN), vbLf)
    Dim nFiles As Long
    If LenB(names(0)) > 0 Then nFiles = UBound(names) + 1
    Dim capped As Boolean: capped = (totalN > nFiles)

    Dim i As Long
    For i = 0 To nFiles - 1
        Dim rec As String
        If ReadBeacon(folderPath & names(i), rec) Then
            Dim f() As String: f = Split(rec, vbTab)
            If UBound(f) >= 8 Then
                ' 書き出す行のキーになるので SanitizeId を通す(タブ・改行が
                ' 混じると読む側の列が丸ごとズレる)。数値欄は BoardNum を通す。
                Dim uid As String: uid = modP2PIo.IdKey(f(0))
                If LenB(uid) > 0 Then
                    usersN = usersN + 1
                    ' 称号は TITLE_MIN 未満を載せない(行数が称号持ちの人数で
                    ' 頭打ちになる。読む側の判定は5件/20件のままなので見え方は
                    ' 変わらない)。
                    Dim thN As Long: thN = modShare.BoardNum(f(1))
                    If thN >= TITLE_MIN Then titleAgg(uid) = thN
                    If f(2) = dk Then orgD = orgD + modShare.BoardNum(f(3))   ' 同じ日キーのみ合算
                    If f(6) = yk Then orgY = orgY + modShare.BoardNum(f(7))
                    If f(4) = mk Then
                        orgM = orgM + modShare.BoardNum(f(5))
                        ' 部別の今月合計。旧実装は自分の部だけを数えていたが、
                        ' この1本は他の端末も読むので全部署ぶんを持つ(行数は
                        ' 部の数だけ)。team列の無い旧形式は空が返り対象外。
                        Dim dept As String
                        dept = modP2PIo.DeptOf(modP2PIo.SanitizeId(modP2PIo.BeaconTeamField(f)))
                        If LenB(dept) > 0 Then
                            If deptAgg.Exists(dept) Then
                                deptAgg(dept) = CLng(deptAgg(dept)) + modShare.BoardNum(f(5))
                            Else
                                deptAgg(dept) = modShare.BoardNum(f(5))
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ' 発行者自身の画面にも同じ数字を出す(自分だけ別経路にしない)。
    mOrgDay = orgD: mOrgMon = orgM: mOrgYear = orgY
    Set mTitles = titleAgg
    mLoaded = True
    mDeptMon = 0
    If LenB(mMyDept) > 0 Then
        If deptAgg.Exists(mMyDept) Then mDeptMon = CLng(deptAgg(mMyDept))
    End If
    mAggStamp = modUtil.NowStamp()
    mAggUsers = usersN
    mAggTotal = totalN
    mAggApprox = capped
    mAggState = "ok"

    Dim body As String
    body = modTelemetry.BoardBodyCarry( _
        modShare.BoardHeadText(mAggStamp, dk, orgD, mk, orgM, yk, orgY, usersN, _
                               capped, totalN), deptAgg, titleAgg)

    ' 置き換えの実体(一時ファイル経由の差し替え)は modShare 側。
    Dim myHash As String
    On Error Resume Next
    myHash = modUtil.Fnv1a64Hex(modP2P.CurrentUserId())
    On Error GoTo 0
    Dim wroteOk As Boolean
    If LenB(myHash) > 0 Then wroteOk = modShare.BoardWriteSummary(folderPath, myHash, body)
    On Error Resume Next
    modLog.LogUsage "board_summary", "write", _
        "組織集計のスナップショットを" & IIf(wroteOk, "更新しました", "書けませんでした") & _
        "(この端末は発行者用ブックなので集計を書く側です): " & usersN & "名ぶん" & _
        "/在" & totalN & "本・称号引継" & carried & "件" & _
        IIf(capped, "・1回" & SCAN_CAP & "本の窓(日ごとに移動)", "")
    On Error GoTo 0
End Sub

' LoadSnapshot - 発行者以外の全端末が通る道。summary.txt を1本だけ読む。
'   壊れている/古すぎるときは数字を捨てる(ClearAggregate)。全件走査への
'   フォールバックは【禁止】: それでは12,000オープンが復活する。
Private Sub LoadSnapshot(ByVal folderPath As String)
    ' R33H F17: 先に在るかだけ見る。ReadBeacon の3回×150/300/450ms は
    ' 「他人の書きかけ」を待つための待ちで、【無いことが定常状態になりうる】
    ' summary.txt に払う筋合いは無い(導入直後・発行者不在の全端末が起動時と
    ' TTL満了ごとに900ms払っていた)。
    If LenB(Dir(folderPath & modShare.BOARD_SUMMARY_NAME)) = 0 Then
        ClearAggregate "none"
        Exit Sub
    End If
    Dim rec As String
    If Not ReadBeacon(folderPath & modShare.BOARD_SUMMARY_NAME, rec) Then
        ClearAggregate "none"
        Exit Sub
    End If

    ' R33H F15: 入口は BoardTextStatus(終端行が無ければヘッダが綺麗でもbroken)。
    Dim head As String: head = modShare.BoardHeadLine(rec)
    Dim st As String
    st = modShare.BoardTextStatus(rec, modUtilText.IsoDateTime(Now - SUMMARY_MAX_AGE_HOURS / 24#))
    If StrComp(st, "ok", vbBinaryCompare) <> 0 Then
        ClearAggregate st
        Exit Sub
    End If

    mMyDept = modP2PIo.DeptOf(MyTeamCode())
    mOrgDay = modShare.BoardHeadMin(head, "d", modUtilText.IsoDateCompact(Date))
    mOrgMon = modShare.BoardHeadMin(head, "m", modUtilText.IsoYm(Date))
    mOrgYear = modShare.BoardHeadMin(head, "y", modUtilText.IsoYear(Date))
    mAggStamp = modShare.BoardHeadField(head, 1)
    mAggUsers = modShare.BoardNum(modShare.BoardHeadField(head, 8))
    mAggTotal = modShare.BoardNum(modShare.BoardHeadField(head, 10))
    mAggApprox = (modShare.BoardHeadField(head, 9) = "1")
    Set mTitles = CreateObject("Scripting.Dictionary")
    mLoaded = True
    mDeptMon = modShare.BoardReadRows(rec, mMyDept, mTitles)
    mAggState = "ok"
    LogAggState "read", "組織集計を summary.txt 1本から読みました(" & mAggUsers & "名ぶん" & _
        IIf(mAggApprox, "・概算", "") & " / 集計時点 " & mAggStamp & ")。"
End Sub

' ClearAggregate - 使える集計が無いときの見せ方。数字は0にし理由を残す。
'   古い数字を出し続けない ―― いつの数字か分からない数字がいちばん困る。
'   称号(mTitles)も同時に空にする(中途半端に前回ぶんが残ると、人によって
'   称号が付いたり付かなかったりする理由を誰も説明できない)。
Private Sub ClearAggregate(ByVal stateText As String)
    mOrgDay = 0: mOrgMon = 0: mOrgYear = 0
    mDeptMon = 0
    Set mTitles = CreateObject("Scripting.Dictionary")
    mLoaded = False
    mAggStamp = ""
    mAggUsers = 0
    mAggTotal = 0
    mAggApprox = False
    mAggState = stateText
    LogAggState stateText, modShare.BoardStateText(stateText, SUMMARY_MAX_AGE_HOURS)
End Sub

' 状態が変わったときだけ usage_log へ1行(10分ごとに同じ行を積まない)。
Private Sub LogAggState(ByVal stateText As String, ByVal detail As String)
    If StrComp(stateText, mLoggedState, vbBinaryCompare) = 0 Then Exit Sub
    mLoggedState = stateText
    On Error Resume Next
    modLog.LogUsage "board_summary", stateText, detail
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 旧サイドバーウィジェット(nx_sb_stat*)の掃除だけを行う後方互換スタブ。
' 2026-07-26 再設計でチャット画面のサイドバーを全廃したため置き場所そのものが
' 無くなった。集計値は Hub の統計タイル(みんな今日/今月)へ移し、詳細は
' OnWidgetClick(タイル上の透明Shape)から出す。既存ブックに残る旧Shapeを
' 消すためだけに残す(Public契約と呼び出し元を壊さないため)。
' ----------------------------------------------------------------------------
Public Sub DrawWidget()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Nexus")
    If ws Is Nothing Then Exit Sub
    ws.Shapes("nx_sb_stat_bg").Delete
    ws.Shapes("nx_sb_stat_bar").Delete
    ws.Shapes("nx_sb_stat_fill").Delete
    ws.Shapes("nx_sb_stat_hist").Delete
    On Error GoTo 0
End Sub

' Hubタイルの値へ付ける概算の断り(R33H F18。文言は modShare 側)。
Public Function OrgApproxSuffix() As String
    OrgApproxSuffix = modShare.BoardApproxSuffix(mAggUsers, mAggTotal, mAggApprox)
End Function

' R49(監査 H-H-10 で発覚): Public OrgSummaryText を削除した。
' 「Hub のタイルから呼ぶための公開集計値」とあったが**誰も呼んでいない**。
' Hub タイルは modHubStat:270 が OrgApproxSuffix を使って自前で組み立て、
' ポップアップは OrgBlockForPopup → modShare.BoardOrgBlock を通る。
' つまり同じ文面を作る【3つ目の写し】で、憲章 §4-5(2箇所で違う答えを出す
' 実装を残さない)に反していた。過去の監査報告に名前が出ているだけで
' 孤児Public検査から救済されていた。

' ポップアップの組織ブロック(文面の組み立ては modShare.BoardOrgBlock)。
Private Function OrgBlockForPopup() As String
    OrgBlockForPopup = modShare.BoardOrgBlock(mAggState, FmtMin(mOrgDay), FmtMin(mOrgMon), _
        FmtMin(mOrgYear), modShare.BoardDeptLine(mMyDept, mDeptMon, mAggApprox), _
        mAggStamp, mAggUsers, mAggApprox, _
        SCAN_CAP, SUMMARY_MAX_AGE_HOURS)
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------
Private Function MyMin(ByVal kind As String, ByVal keyPart As String) As Long
    On Error Resume Next
    MyMin = modStats.GetStat("sv:" & kind & ":" & keyPart)
    On Error GoTo 0
End Function

' R13-7c: 自分のチーム/部コードの優先順位。config user_department →
' modP2PIo.TeamCodeOf(自分のuserId)。どちらも取れなければ空(DeptOfも空を
' 返すので、部の合算行そのものがポップアップに出ない)。
'
' R13 F12: user_department は自由記述で、実機には「営業部」のような部署名が
' 入っている。初版はそれを無検証で採用していたため、team列に日本語が乗り、
' DeptOf がその先頭3字を部コードとして扱っていた(他端末と噛み合わず、
' 部別集計が黙って狂う)。規約に合う文字列のときだけ採用し、合わなければ
' 自分のuserIdからの推定へ落とす ―― 「設定されているが規約外」を
' 「未設定」と同じに扱うのが、集計を汚さない唯一の扱い方になる。
Private Function MyTeamCode() As String
    On Error Resume Next
    Dim d As String: d = Trim$(modConfig.GetString("user_department", ""))
    If modP2PIo.IsTeamCode(d) Then
        MyTeamCode = d
        Exit Function
    End If
    MyTeamCode = modP2PIo.TeamCodeOf(modP2P.CurrentUserId())
    On Error GoTo 0
End Function

' 直近7日の個人履歴(日付キーを7回引くだけ。0分の日は「-」)。
Private Function History7() As String
    Dim s As String
    Dim i As Long
    For i = 0 To 6
        Dim d As Date: d = Date - i
        Dim v As Long: v = MyMin("d", modUtilText.IsoDateCompact(d))   ' R12-H-3(同上)
        s = s & "  " & Format$(d, "mm/dd") & ": " & IIf(v > 0, FmtMin(v), "-") & vbLf
    Next i
    History7 = s
End Function

Private Function FmtMin(ByVal m As Long) As String
    If m >= 60 Then
        FmtMin = Format$(m \ 60, "0") & "時間" & IIf(m Mod 60 > 0, (m Mod 60) & "分", "")
    Else
        FmtMin = m & "分"
    End If
End Function

' 2026-07-31(レビュー R8 F2): 自前で nexus_share_path を読んでパスを
' 組み立てていたため、modShare の関所(1セッション1回の到達判定)を
' 素通りしていた。届かない共有に対して起動のたび OS のタイムアウトを
' 払い直すうえ、「modShare は届かないと言っているのに modBoard は
' 触りに行く」という矛盾した状態になる。到達判定とルート解決は
' modShare だけが行う、が本来の設計(modShare 冒頭「唯一性の原則」)。
Private Function BoardDir() As String
    BoardDir = modShare.SubDir(BOARD_SUBDIR)
End Function

' ビーコンI/O(AVロック耐性リトライ。COMは両経路Set=Nothing。modMentorと同仕様の複製)
Private Function WriteBeacon(ByVal filePath As String, ByVal content As String) As Boolean
    Dim attempt As Long
    For attempt = 1 To 3
        If TryW(filePath, content) Then
            ' 2026-07-31(レビュー R8・F7最小実装): 成功/失敗を modShare へ伝える。
            ' 到達OKと判定した後にVPNが切れると、各所のリトライ(3回×バックオフ)を
            ' 延々と払い続ける。連続5回の失敗でそのセッションを見送りへ倒す。
            On Error Resume Next
            modShare.ReportSuccess
            On Error GoTo 0
            WriteBeacon = True
            Exit Function
        End If
        BoardWait 250 * attempt
    Next attempt
    On Error Resume Next
    modShare.ReportFailure "modBoard WriteBeacon(ビーコン発信)"
    On Error GoTo 0
End Function

' UTF-8書き出しの実体は modUtilText.WriteTextFileUtf8(2026-07-31 R11-F2)。
Private Function TryW(ByVal filePath As String, ByVal content As String) As Boolean
    TryW = modUtilText.WriteTextFileUtf8(filePath, content)
End Function

Private Function ReadBeacon(ByVal filePath As String, ByRef outText As String) As Boolean
    Dim attempt As Long
    For attempt = 1 To 3
        If TryR(filePath, outText) Then
            On Error Resume Next
            modShare.ReportSuccess
            On Error GoTo 0
            ReadBeacon = True
            Exit Function
        End If
        BoardWait 150 * attempt
    Next attempt
    ' 2026-07-31(R8b B3): ここで modShare.ReportFailure を呼んではいけない。
    ' 読めないのは【他人が置いた1ファイル】で、書き込み中・AVスキャン中・
    ' 権限違いなど、共有そのものは健全でも普通に起こる。40人いれば5件くらい
    ' 続けて読めない日は珍しくない。それでセッションを到達不可へ降格すると、
    ' 感謝状の送受信・チャンネル同期・正典の発行までまとめて止まり、
    ' 受信箱には「共有フォルダが未設定です」と出る。
    ' 何も設定を変えていない利用者が、原因を突き止めようのない全滅を見る。
    ' 降格の根拠にしてよいのは「自分のファイルを自分で書けない」= WriteBeacon
    ' 側だけ(そちらは残してある)。
End Function

' UTF-8読み取りの実体は modUtilText.ReadTextFileUtf8(2026-07-31 R11-F2)。
Private Function TryR(ByVal filePath As String, ByRef outText As String) As Boolean
    TryR = modUtilText.ReadTextFileUtf8(filePath, outText)
End Function

Private Sub BoardWait(ByVal ms As Long)
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000# < ms
        DoEvents
        If Timer < t0 Then Exit Do
    Loop
End Sub

' modHubの統計タイルから参照する集計値のアクセサ(モジュール変数の直接公開を避ける)。
Public Function OrgMinutesDay() As Long
    OrgMinutesDay = mOrgDay
End Function

Public Function OrgMinutesMon() As Long
    OrgMinutesMon = mOrgMon
End Function
