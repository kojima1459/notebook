Attribute VB_Name = "modBoard"
Option Explicit

' ============================================================================
' modBoard - チーム連帯ボード(組織全体の節約時間+感謝ベース称号)。疎結合プラグイン。
' ----------------------------------------------------------------------------
' 役割:
'   1) statsビーコン: 自分の統計(感謝受領数/節約時間 日・月・年)を共有フォルダ
'      <nexus_share_path>\board\stats_<自分hash>.txt へ発信する(1人1ファイル上書き
'      =書込競合ゼロ。noise投票と同じ実証済みパターン)。
'   2) 集計: 全員のビーコンを読み、組織全体の「今日/今月/今年の節約時間」を算出。
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

    ' --- ここから下は共有フォルダに触らない(到達性に関係なく必ず実行する)---
    DrawWidget             ' 旧サイドバーShapeの掃除(ローカルのみ)
    ShowWeeklySummary      ' B-5: 週の初回起動時だけ、先週の節約時間を労いToast

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

    modSkin.ShowToast "先週、あなたはこのツールで " & FmtMin(total) & " を節約しました。今週もいいスタートを。", "success", True   ' R29H F2
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
        If mTitles.Exists(LCase$(userId)) Then n = CLng(mTitles(LCase$(userId)))
    End If
    If n >= 20 Then
        TitleFor = ChrW(&HD83C) & ChrW(&HDF1F) & " "        ' U+1F31F 星
    ElseIf n >= 5 Then
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
           "【みんな(組織全体)】" & vbLf & _
           "  今日: " & FmtMin(mOrgDay) & "  /  今月: " & FmtMin(mOrgMon) & _
           "  /  今年: " & FmtMin(mOrgYear) & DeptLineForPopup() & vbLf & vbLf & _
           "※「" & ChrW(&H2705) & "解決した」1回=" & modP2PIo.MinutesPerSelfsolve() & _
           "分の節約として、他者からの感謝と同じしくみで" & vbLf & _
           "  組織に共有・合算されます(共有フォルダ経由なので最大10分遅れの目安)。" & vbLf & _
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
'   ダッシュボードの「更新」ボタンだけがこれを呼ぶ。利用者が明示的に
'   更新を押したときに「10分経っていないので前の値です」と返すのは、
'   ボタンが壊れているのと区別が付かない。
' ----------------------------------------------------------------------------
Public Sub ForceRefreshBoard()
    mAggAt = 0
    RefreshBoard
End Sub

' ----------------------------------------------------------------------------
' RefreshBoard - 全員のビーコンを読んで組織合計を出し直す。
'
' 2026-07-31(レビュー R8 F9): 10分のTTLを付けた。
'   ここは board\stats_*.txt を【人数ぶん】開いて読む。Hubは画面を戻る
'   たびに描き直されるので、TTLが無いと「1人が画面を行き来した回数 ×
'   部内の人数」がそのままファイルサーバへの往復になる。数十人のPoCでも
'   午前中に数千回になり得る。値は分単位の集計で、10分古くても実害は無い。
'
'   本来は「誰か1人が集計スナップショットを書き、他は1ファイル読むだけ」に
'   すべきで、1,000人を超える規模ではTTLでは足りない。ただしそれは
'   「誰が書くか」「壊れたスナップショットをどう検出するか」という別設計を
'   伴うため、PoC(数十人)ではTTLで足りると判断した。本格展開時の設計課題。
' ----------------------------------------------------------------------------
Private Sub RefreshBoard()
    ' TTL内なら前回の集計値をそのまま使う(mOrgDay等は消さない)。
    If modShareRule.CacheIsFresh(mAggAt, Timer, AGG_TTL_SEC) Then Exit Sub

    ' 2026-07-31(R8b B8): ゼロ化・TTLの起点更新は【集計に入れると決めてから】。
    ' 従来はここより前で mOrgDay 等を0にし mAggAt も進めていたため、
    ' 共有が一時的に届かない(BoardDir が空)1回の呼び出しだけで
    '   ・「みんなの節約」が 0分 / 0分 になる
    '   ・他人の称号(mTitles)が全部消える
    ' という表示になり、しかも mAggAt が進んでいるので【10分間そのまま】
    ' 固定された。共有はすぐ復帰しているのに画面だけが壊れて見える。
    ' 集計できないときは、古い値を出し続ける方がずっとましなので何も触らない。
    Dim folderPath As String: folderPath = BoardDir()
    If LenB(folderPath) = 0 Then Exit Sub
    If LenB(Dir(folderPath, vbDirectory)) = 0 Then Exit Sub

    mAggAt = Timer
    mOrgDay = 0: mOrgMon = 0: mOrgYear = 0
    mDeptMon = 0
    Set mTitles = CreateObject("Scripting.Dictionary")
    mLoaded = True

    Dim dk As String: dk = modUtilText.IsoDateCompact(Date)
    Dim mk As String: mk = modUtilText.IsoYm(Date)
    Dim yk As String: yk = modUtilText.IsoYear(Date)

    ' R13-7c: 自分の部(先頭3字)を先に確定させておく。config優先、無ければ
    ' 自分のuserIdの末尾チームコードから推定(どちらも不明なら空のまま=
    ' 部の合算行そのものをポップアップへ出さない)。
    mMyDept = modP2PIo.DeptOf(MyTeamCode())

    ' collect-then-process(Dir列挙中に他のDirを呼ばない)
    Dim names() As String: ReDim names(0 To 63)
    Dim nFiles As Long: nFiles = 0
    Dim fn As String: fn = Dir(folderPath & "stats_*.txt")
    Do While LenB(fn) > 0
        If nFiles > UBound(names) Then ReDim Preserve names(0 To UBound(names) + 64)
        names(nFiles) = fn
        nFiles = nFiles + 1
        fn = Dir()
    Loop

    Dim i As Long
    For i = 0 To nFiles - 1
        Dim rec As String
        If ReadBeacon(folderPath & names(i), rec) Then
            Dim f() As String: f = Split(rec, vbTab)
            If UBound(f) >= 8 Then
                Dim uid As String: uid = LCase$(Trim$(f(0)))
                If LenB(uid) > 0 Then
                    ' R13 L-batch: 数値欄は必ず SafeNum を通す。素の CLng(Val(...))
                    ' だと壊れたビーコン1本("99999999999" 等)でオーバーフローし、
                    ' その時点で集計ループごと中断=以降の全員ぶんが欠ける。
                    mTitles(uid) = SafeNum(f(1))                      ' 称号: 感謝受領数
                    If f(2) = dk Then mOrgDay = mOrgDay + SafeNum(f(3))   ' 同じ日キーのみ合算
                    If f(4) = mk Then mOrgMon = mOrgMon + SafeNum(f(5))
                    If f(6) = yk Then mOrgYear = mOrgYear + SafeNum(f(7))

                    ' R13-7c: 自分の部が分かっているときだけ、同じ部のビーコンを
                    ' 追加で合算する(team列が無い旧形式ビーコンはBeaconTeamFieldが
                    ' 空文字を返すので自然に対象外になる)。
                    ' R13 L-batch: 日次の部合算(旧mDeptDay)は溜めるだけで
                    ' どこにも出ていなかったので外した。出す先ができたときに
                    ' 月次と同じ形で足せばよい(使われない状態のまま持たない)。
                    If LenB(mMyDept) > 0 Then
                        Dim beaconTeam As String: beaconTeam = modP2PIo.BeaconTeamField(f)
                        If LenB(beaconTeam) > 0 Then
                            If StrComp(modP2PIo.DeptOf(beaconTeam), mMyDept, vbTextCompare) = 0 Then
                                If f(4) = mk Then mDeptMon = mDeptMon + SafeNum(f(5))
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Sub

' ----------------------------------------------------------------------------
' サイドバーウィジェット描画(nx_sb_stat* = 既存Z-Order/テーマループ管轄)。
' 実機報告(2026-07-22)「今日の節約時間が0分のまま」対策: 起動時に1度しか
' 呼ばれておらず、その後「解決した」を押してもウィジェットが再描画されず
' 表示が固まっていた。Publicにして加算直後にも呼べるようにする。
' ----------------------------------------------------------------------------
' 2026-07-26 再設計: チャット画面のサイドバーを全廃したため、このウィジェットの
' 置き場所そのものが無くなった。集計値は Hub の統計タイル(みんな今日/今月)へ
' 移し、詳細ランキングは modBoard.OnWidgetClick(Hubのタイル上の透明Shape)から
' 出す。既存ブックに残っている旧Shapeを掃除するだけの後方互換スタブとして残す
' (Public契約と呼び出し元を壊さないため)。
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

' Hub の「みんなの節約」タイルから呼ぶための公開集計値。
Public Function OrgSummaryText() As String
    OrgSummaryText = ChrW(&HD83C) & ChrW(&HDF0D) & " みんなの節約時間: 今日 " & _
        FmtMin(mOrgDay) & " / 今月 " & FmtMin(mOrgMon)
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

' 自分の部が分かっていて、かつ今月の合算が1分でもあるときだけ1行足す
' (タイル新設はしない。R13-7c)。
Private Function DeptLineForPopup() As String
    If LenB(mMyDept) = 0 Then Exit Function
    If mDeptMon <= 0 Then Exit Function
    ' R13 L-batch: 60分未満を時間へ丸めると「約0時間」になる。1分でも
    ' 貯まっているから出している行なのに「0」と書くのは、事実としても
    ' 労いとしても間違っている。60分未満は分のまま出す。
    Dim amt As String
    If mDeptMon < 60 Then
        amt = mDeptMon & "分"
    Else
        amt = CLng(Round(mDeptMon / 60, 0)) & "時間"
    End If
    DeptLineForPopup = vbLf & "  部(" & mMyDept & ")で今月 約" & amt
End Function

' ビーコンの数値欄を安全に読む(R13 L-batch)。共有フォルダのファイルは
' 誰でも書ける以上、壊れた値・巨大な値・負の値が来る前提で扱う。
' Long の範囲を超える値で CLng がオーバーフローすると、集計ループが
' その1本で止まり、以降のビーコンが【全部】欠けたまま画面に出てしまう。
' 1億分(=約190年ぶん)を超える値と負の値は、現実の節約時間ではないので
' 0 として捨てる。読めない文字列は Val が 0 を返すのでそのまま 0 になる。
Private Function SafeNum(ByVal s As String) As Long
    On Error GoTo Bad
    Dim v As Double: v = Val(s)
    If v < 0 Or v > 100000000# Then Exit Function
    SafeNum = CLng(v)
    Exit Function
Bad:
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
