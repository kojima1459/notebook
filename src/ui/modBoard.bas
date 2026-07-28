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
'     sv:y:yyyy)。✅解決(modAsk.FeedbackGreen)時に1回15分をBumpする(発火点は
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
Private Const MIN_PER_SOLVE As Long = 15     ' modStats.MINUTES_PER_SELFSOLVEと同値

' ビーコン集計キャッシュ(セッション内。共有フォルダ再走査はBootBoard時のみ)
Private mOrgDay As Long, mOrgMon As Long, mOrgYear As Long
Private mTitles As Object    ' Dictionary: id(小文字) -> thanks受領数
Private mLoaded As Boolean

' ----------------------------------------------------------------------------
' BootBoard - エントリポイント(LaunchNexus末尾から1行フック)。
'   発信→集計→ウィジェット描画。全障害を内部で握る。
' ----------------------------------------------------------------------------
Public Sub BootBoard()
    On Error Resume Next   ' 安全弁: 本機能の失敗を起動へ絶対に波及させない
    PublishBeacon
    RefreshBoard
    DrawWidget
    ShowWeeklySummary      ' B-5: 週の初回起動時だけ、先週の節約時間を労いToast
    On Error GoTo 0
End Sub

' 週次サマリー(B-5): その週の初回起動時に1回だけ、先週の個人合計をToastで返す。
' ガードはISO風の年+週番号キー(wk:yyyww)。sv:d:日付キーを7日分読むだけ。
Private Sub ShowWeeklySummary()
    Dim wkKey As String
    wkKey = "wk:" & Format$(Date, "yyyy") & Format$(DatePart("ww", Date, vbMonday), "00")
    If modStats.GetStat(wkKey) > 0 Then Exit Sub
    modStats.Bump wkKey

    ' 先週(直近の月曜の7日前～日曜)の個人合計
    Dim mon As Date: mon = Date - Weekday(Date, vbMonday) + 1   ' 今週の月曜
    Dim total As Long, i As Long
    For i = 1 To 7
        total = total + MyMin("d", Format$(mon - i, "yyyymmdd"))
    Next i
    If total <= 0 Then Exit Sub   ' ゼロ週は何も言わない(空虚な自慢をしない)

    modSkin.ShowToast "先週、あなたはこのツールで " & FmtMin(total) & " を取り戻しました。今週もいいスタートを。", "success"
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

    ' 2026-07-26: 呼び出し元がHubの統計タイル(ホームシート)へ移ったのに、
    ' ここは常にNexusシートへ描いていた。別シートに描かれるので画面上は
    ' 「押しても何も出ない」状態になる。今アクティブなシートへ描く。
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    If ws Is Nothing Then GoTo Done

    Dim body As String
    body = ChrW(&HD83D) & ChrW(&HDCC8) & " 節約時間レポート" & vbLf & vbLf & _
           "【あなた】" & vbLf & _
           "  今日: " & FmtMin(MyMin("d", Format$(Date, "yyyymmdd"))) & _
           "  /  今月: " & FmtMin(MyMin("m", Format$(Date, "yyyymm"))) & _
           "  /  今年: " & FmtMin(MyMin("y", Format$(Date, "yyyy"))) & vbLf & vbLf & _
           "【直近7日の履歴】" & vbLf & History7() & vbLf & _
           "【みんな(組織全体)】" & vbLf & _
           "  今日: " & FmtMin(mOrgDay) & "  /  今月: " & FmtMin(mOrgMon) & _
           "  /  今年: " & FmtMin(mOrgYear) & vbLf & vbLf & _
           "※「" & ChrW(&H2705) & "解決した」1回=15分の節約として、他者からの感謝と同じP2P機構で" & vbLf & _
           "  組織に共有・合算されます。" & vbLf & "(クリックで閉じる)"

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
Private Sub PublishBeacon()
    Dim folderPath As String: folderPath = BoardDir()
    If LenB(folderPath) = 0 Then Exit Sub
    On Error Resume Next
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    On Error GoTo 0

    Dim myId As String
    On Error Resume Next
    myId = modP2P.CurrentUserId()
    On Error GoTo 0
    If LenB(myId) = 0 Then Exit Sub

    Dim dk As String: dk = Format$(Date, "yyyymmdd")
    Dim mk As String: mk = Format$(Date, "yyyymm")
    Dim yk As String: yk = Format$(Date, "yyyy")
    Dim rowText As String
    rowText = myId & vbTab & modStats.GetStat("thanks_received_total") & vbTab & _
              dk & vbTab & MyMin("d", dk) & vbTab & _
              mk & vbTab & MyMin("m", mk) & vbTab & _
              yk & vbTab & MyMin("y", yk) & vbTab & modUtil.NowStamp()

    WriteBeacon folderPath & "stats_" & modUtil.Fnv1a64Hex(myId) & ".txt", rowText
End Sub

Private Sub RefreshBoard()
    mOrgDay = 0: mOrgMon = 0: mOrgYear = 0
    Set mTitles = CreateObject("Scripting.Dictionary")
    mLoaded = True

    Dim folderPath As String: folderPath = BoardDir()
    If LenB(folderPath) = 0 Then Exit Sub
    If LenB(Dir(folderPath, vbDirectory)) = 0 Then Exit Sub

    Dim dk As String: dk = Format$(Date, "yyyymmdd")
    Dim mk As String: mk = Format$(Date, "yyyymm")
    Dim yk As String: yk = Format$(Date, "yyyy")

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
                    mTitles(uid) = CLng(Val(f(1)))                      ' 称号: 感謝受領数
                    If f(2) = dk Then mOrgDay = mOrgDay + CLng(Val(f(3)))   ' 同じ日キーのみ合算
                    If f(4) = mk Then mOrgMon = mOrgMon + CLng(Val(f(5)))
                    If f(6) = yk Then mOrgYear = mOrgYear + CLng(Val(f(7)))
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

' 直近7日の個人履歴(日付キーを7回引くだけ。0分の日は「-」)。
Private Function History7() As String
    Dim s As String
    Dim i As Long
    For i = 0 To 6
        Dim d As Date: d = Date - i
        Dim v As Long: v = MyMin("d", Format$(d, "yyyymmdd"))
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

Private Function BoardDir() As String
    Dim basePath As String: basePath = modConfig.GetString("nexus_share_path", "")
    If LenB(basePath) = 0 Then Exit Function
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
    BoardDir = basePath & BOARD_SUBDIR & "\"
End Function

' ビーコンI/O(AVロック耐性リトライ。COMは両経路Set=Nothing。modMentorと同仕様の複製)
Private Function WriteBeacon(ByVal filePath As String, ByVal content As String) As Boolean
    Dim attempt As Long
    For attempt = 1 To 3
        If TryW(filePath, content) Then
            WriteBeacon = True
            Exit Function
        End If
        BoardWait 250 * attempt
    Next attempt
End Function

Private Function TryW(ByVal filePath As String, ByVal content As String) As Boolean
    Dim st As Object
    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText content
    st.SaveToFile filePath, 2
    st.Close
    Set st = Nothing
    TryW = True
    Exit Function
Fail:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Set st = Nothing
    On Error GoTo 0
End Function

Private Function ReadBeacon(ByVal filePath As String, ByRef outText As String) As Boolean
    Dim attempt As Long
    For attempt = 1 To 3
        If TryR(filePath, outText) Then
            ReadBeacon = True
            Exit Function
        End If
        BoardWait 150 * attempt
    Next attempt
End Function

Private Function TryR(ByVal filePath As String, ByRef outText As String) As Boolean
    Dim st As Object
    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile filePath
    outText = CStr(st.ReadText(-1))
    st.Close
    Set st = Nothing
    TryR = True
    Exit Function
Fail:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Set st = Nothing
    On Error GoTo 0
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
