Attribute VB_Name = "modChatLog"
Option Explicit

' ============================================================================
' modChatLog - チャット履歴シート("チャット履歴")への質問/回答の記録
' ----------------------------------------------------------------------------
' 役割:
'   ユーザーがCtrl+Fで過去の質問・回答を検索できるよう、可視シート
'   「チャット履歴」へ1往復=1行で追記する(modLog.basの設計を踏襲)。
'
' 設計判断:
'   ・シートは初回書込時にその場で生成する(modLogのEnsureLogSheetと同じ
'     「ビルド時に無くても動作を止めない」保険)。ヘッダはA=質問/B=回答/
'     C=日時/D=モード、可視(err_log/usage_logと異なりユーザーが直接見る
'     前提のため非表示にしない)。
'   ・最新を常に行2へ挿入する(Rows(2).Insert)。ユーザーがシートを開いた
'     瞬間に最新のやり取りが見えるようにするため。
'   ・上限100件: 挿入後、データ行(2行目以降)が100件を超えたら最下行から
'     削除する。無限に増え続けてブックが肥大化するのを防ぐ。
'   ・数式インジェクション対策: 質問/回答の先頭文字が =/+/-/@ のいずれかだと
'     Excelが数式として解釈してしまうため、アポストロフィを前置してテキスト
'     強制する(SafeLeftで長さを絞った後に判定する)。
'   ・WrapText=Falseで書き込む: 折り返し有効のまま長文が入ると行高さが
'     自動拡張され、シートが縦に爆発する(描画サイクルの重さにも波及する)
'     ため明示的に無効化する。
'   ・書込失敗はDebug.Printのみに留め、アプリの処理を止めない
'     (「ログで死なない」。modLog.LogErrorと同じ運用方針)。
'   ・エスケープハッチ: config chat_log_enabled=FALSE で何もしない
'     (既定TRUE)。
' ============================================================================

Private Const SHEET_NAME As String = "チャット履歴"
Private Const MAX_ROWS As Long = 100
Private Const MAX_Q_CHARS As Long = 1000
Private Const MAX_ANS_CHARS As Long = 5000

' ----------------------------------------------------------------------------
' LogTurn - 唯一の公開API。1往復分(質問/回答/モード)を最新行として記録する。
' ----------------------------------------------------------------------------
Public Sub LogTurn(ByVal q As String, ByVal ans As String, ByVal mode As String)
    If Not modConfig.GetBool("chat_log_enabled", True) Then Exit Sub

    On Error GoTo Fail
    Dim ws As Worksheet
    Set ws = EnsureChatSheet()
    If ws Is Nothing Then GoTo Fail

    ws.Rows(2).Insert Shift:=-4121   ' xlShiftDown (名前付き定数への依存を避ける本プロジェクトの慣習に合わせ数値リテラル)

    Dim qs As String, ans2 As String
    qs = modUtilText.SanitizeForCell(modUtil.SafeLeft(q, MAX_Q_CHARS))
    ans2 = modUtilText.SanitizeForCell(modUtil.SafeLeft(ans, MAX_ANS_CHARS))

    With ws.Range(ws.Cells(2, 1), ws.Cells(2, 4))
        .WrapText = False
    End With
    ws.Cells(2, 1).Value = qs
    ws.Cells(2, 2).Value = ans2
    ws.Cells(2, 3).Value = modUtil.NowStamp()
    ws.Cells(2, 4).Value = mode

    TrimToMaxRows ws
    Exit Sub

Fail:
    ' ログで死なない(modLogと同じ運用): 書込失敗はDebug.Printのみに留める。
    Debug.Print "[modChatLog.LogTurn:書込失敗] mode=" & mode & " q=" & Left$(q, 80)
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------
Private Function EnsureChatSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Worksheets.Addは新シートをアクティブ化する仕様のため、そのままだと
        ' 「最初の回答の直後に画面がチャット履歴シートへ勝手に切り替わる」。
        ' 追加前のアクティブシートを覚えて、生成後に必ず戻す。
        Dim prevActive As Object
        On Error Resume Next
        Set prevActive = ThisWorkbook.ActiveSheet
        On Error GoTo 0

        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = SHEET_NAME
        ws.Cells(1, 1).Value = "質問"
        ws.Cells(1, 2).Value = "回答"
        ws.Cells(1, 3).Value = "日時"
        ws.Cells(1, 4).Value = "モード"
        ws.Columns(1).ColumnWidth = 50
        ws.Columns(2).ColumnWidth = 80
        ws.Columns(3).ColumnWidth = 20
        ws.Columns(4).ColumnWidth = 12
        ws.Visible = -1   ' xlSheetVisible (名前付き定数への依存を避けるmodLogの慣習に合わせる)
        On Error GoTo 0

        ' ユーザーが見ていた画面(Nexus/ホーム等)へフォーカスを戻す。
        On Error Resume Next
        If Not prevActive Is Nothing Then prevActive.Activate
        On Error GoTo 0
    End If
    Set EnsureChatSheet = ws
    Exit Function
Fail:
    Set EnsureChatSheet = Nothing
End Function

' データ行(2行目以降)が上限を超えたら、最下行から削除して上限件数に収める。
Private Sub TrimToMaxRows(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim dataRows As Long
    dataRows = lastRow - 1   ' ヘッダ除く
    If dataRows > MAX_ROWS Then
        ' 上限より後ろ(古い方=下側)を一括削除。上限行が (MAX_ROWS + 1)行目、
        ' その次(MAX_ROWS + 2行目)以降が削除対象。
        ws.Range(ws.Cells(MAX_ROWS + 2, 1), ws.Cells(lastRow, 1)).EntireRow.Delete
    End If
End Sub

' ----------------------------------------------------------------------------
' R34 A2: modShare(基盤層)専用のファイルI/Oヘルパーを基盤層内で完結させる
'   置き場所(modUtil/modTelemetryは容量都合・層都合でNG=司令塔裁定)。
' ----------------------------------------------------------------------------
' BakIsStale - fileTimeがnowTimeより24時間(86400秒)より古いか【純関数】。
Public Function BakIsStale(ByVal fileTime As Date, ByVal nowTime As Date) As Boolean
    Dim diffSec As Double: diffSec = (CDbl(nowTime) - CDbl(fileTime)) * 86400#
    BakIsStale = (diffSec > 86400#)
End Function

' BoardSweepStaleBak - folderPath配下のsummary_*.bakのうち24時間より古いものを
'   消す(自分のhashのぶんも含めてよい=書く直前に呼ぶ設計)。
' ----------------------------------------------------------------------------
' R34 F1レビュー(MAJOR-1)の根治: 【集めてから消す】の2周構成にする。
' 初版は Dir の列挙中に Kill を挟み、さらに Err.Clear の位置の都合で
'   ・列挙の途中でディレクトリを書き換える(Dirの列挙状態が壊れる)
'   ・Dir() が失敗すると leaf が更新されず Do While を抜けられない
' という2つの穴があった。作法は modInsightIo.GcOldInsights:294-320 を踏襲:
'   1周目 = 名前だけを配列へ集めて Dir を閉じる
'   2周目 = Err.Clear → FileDateTime → Err.Number=0 を確認 → BakIsStale → Kill
' FileDateTime が落ちた(消えた・権限が無い)ファイルは触らずに次へ進む
' ―― 判定できなかったものを消すのは、古いから消すのとは別のことなので。
' 純関数 BakIsStale は不変。
Public Sub BoardSweepStaleBak(ByVal folderPath As String)
    If LenB(folderPath) = 0 Then Exit Sub
    On Error Resume Next

    ' --- 1周目: 名前を集めきる(この間ディレクトリは書き換えない) ---
    Dim names() As String: ReDim names(0 To 63)
    Dim n As Long
    Dim leaf As String: leaf = Dir(folderPath & "summary_*.bak")
    Do While LenB(leaf) > 0
        If n > UBound(names) Then ReDim Preserve names(0 To UBound(names) + 64)
        names(n) = leaf
        n = n + 1
        leaf = Dir()
        If Err.Number <> 0 Then Exit Do   ' Dir()が落ちたら列挙を打ち切る(無限ループ防止)
    Loop

    ' --- 2周目: 集めた名前だけを見る(Dirの列挙状態にもう依存しない) ---
    Dim nowTime As Date: nowTime = Now
    Dim i As Long
    For i = 0 To n - 1
        Dim full As String: full = folderPath & names(i)
        Dim stamp As Date
        Err.Clear
        stamp = FileDateTime(full)
        If Err.Number = 0 Then
            If BakIsStale(stamp, nowTime) Then Kill full
        End If
        Err.Clear
    Next i
    On Error GoTo 0
End Sub
