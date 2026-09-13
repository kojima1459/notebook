Attribute VB_Name = "modAnalytics"
Option Explicit

' ============================================================================
' modAnalytics - ガバナンス分析用CSVエクスポート(usage_logの行別エンリッチ)
' ----------------------------------------------------------------------------
' 役割:
'   usage_log(§ modAppDef.SH_USAGE)の各行をそのまま横に広げ、誰が(user_id/
'   department)・何を(event/mode/detail)・どれだけ(latency_ms/hit_count/
'   saved_min/exp_activity/exp_thanks)・どの資料クラスタに関連して行ったかを
'   1行=1CSV行として書き出す。ダッシュボードの「📥分析用ログ出力」ボタンが
'   modAnalytics.ExportAnalyticsCsv を直接呼ぶ(modDash側で結線済み)。
'
' 設計判断:
'   ・データ取得は全て既存の公開APIを1回だけ呼び、以後はローカル変数/辞書を
'     使い回す。特に modCluster.SourceClusterMap() はK-Meansを内部で実行する
'     ため、行ループの外で1回だけ呼び出しキャッシュする。
'   ・usage_logの読み取りは他モジュール(modDash.CountUsageEvent等)と同じ
'     慣習で、1データ行のみの場合に .Value がスカラーになる問題を手動配列
'     組立てで回避する。
'   ・CSVはExcelが日本語を正しく開けるよう UTF-8 BOM 付きで書き出す
'     (ADODB.Stream の Charset="utf-8" は既定でBOMを付与する)。各フィールド
'     は CsvField で必ずエスケープする(カンマ/ダブルクォート/改行を含む場合
'     のみ引用符で囲む)。
'   ・保存失敗やダイアログキャンセルはクラッシュさせず、MsgBoxで案内するか
'     静かに終了する。
' ============================================================================

Private Const CSV_HEADER As String = _
    "timestamp,user_id,department,event,mode,detail,latency_ms,hit_count,saved_min,exp_activity,exp_thanks,cluster_id"

' ----------------------------------------------------------------------------
' ExportAnalyticsCsv - usage_logを読み、行ごとにエンリッチしてCSV(UTF-8 BOM付)
'   として書き出す。唯一のPublicメンバー(ダッシュボードのボタンから直接呼ばれる)。
' ----------------------------------------------------------------------------
Public Sub ExportAnalyticsCsv()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_USAGE)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "分析できるログがまだありません。", vbExclamation, modAppDef.APP_NAME
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "分析できるログがまだありません。", vbExclamation, modAppDef.APP_NAME
        Exit Sub
    End If

    Dim prevScreenUpdating As Boolean
    prevScreenUpdating = Application.ScreenUpdating

    On Error GoTo Fail
    Application.ScreenUpdating = False

    Dim rowCount As Long
    rowCount = lastRow - 1

    Dim csvText As String
    csvText = BuildCsvText(ws, lastRow)

    Dim defaultName As String
    defaultName = "MyBookshelf分析ログ_" & Format$(Now, "yyyymmdd") & ".csv"

    Dim savePathVariant As Variant
    savePathVariant = Application.GetSaveAsFilename( _
        InitialFileName:=defaultName, FileFilter:="CSV (*.csv),*.csv")

    If VarType(savePathVariant) = vbBoolean Then
        ' ユーザーがキャンセル(Falseが返る)。静かに終了する。
        Application.ScreenUpdating = prevScreenUpdating
        Exit Sub
    End If

    Dim savePath As String
    savePath = CStr(savePathVariant)

    If Not WriteCsvWithBom(savePath, csvText) Then GoTo Fail

    Application.ScreenUpdating = prevScreenUpdating
    MsgBox "分析ログを書き出しました。" & vbLf & _
        rowCount & "件のログを「" & modUtil.FileNameOf(savePath) & "」として保存しました。", _
        vbInformation, modAppDef.APP_NAME
    Exit Sub

Fail:
    ' Fail:はエラー経由(On Error GoTo Fail)だけでなく、WriteCsvWithBom失敗時の
    ' 明示的な GoTo Fail(エラー未発生)でも来る。両方から安全に通れるよう、
    ' ここでは新たに On Error を書かない(modLog.LogErrorは内部で自己防御済み
    ' の薄いラッパーなので、ここでの追加保護は不要)。
    Dim expErrNum As Long: expErrNum = Err.Number
    Application.ScreenUpdating = prevScreenUpdating
    ' 2026-07-31(R11-E M-5): 失敗が画面のMsgBoxだけに留まり、err_logに痕跡が
    ' 残らなかった(監査1)。原因調査ができない無言の失敗(憲章§4-1)を解消する。
    modLog.LogError "E0801", "modAnalytics.ExportAnalyticsCsv", "CSV書き出し失敗", expErrNum
    MsgBox "分析ログの書き出しに失敗しました。" & vbLf & _
        "保存先フォルダの権限を確認するか、時間を置いてからもう一度お試しください。", _
        vbExclamation, modAppDef.APP_NAME
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' BuildCsvText - usage_logの2行目～lastRow行を読み、ヘッダ+全データ行を
'   CSV(vbCrLf区切り)として組み立てて返す。呼び出し時点で lastRow>=2 が
'   保証されている前提(ExportAnalyticsCsv側で検査済み)。
Private Function BuildCsvText(ByVal ws As Worksheet, ByVal lastRow As Long) As String
    Dim arr As Variant
    arr = ReadUsageRows(ws, lastRow)

    Dim userId As String: userId = SafeUserId()
    Dim department As String: department = SafeDepartment()

    Dim expQuestion As Long: expQuestion = modConfig.GetLong("exp_question", 5)
    Dim expRegister As Long: expRegister = modConfig.GetLong("exp_register", 20)
    Dim expPackShare As Long: expPackShare = modConfig.GetLong("exp_pack_share", 30)
    Dim expThumbup As Long: expThumbup = modConfig.GetLong("exp_thumbup", 10)

    Dim clusterMap As Object
    Set clusterMap = SafeClusterMap()   ' K-Meansは内部で1回だけ実行される

    Dim rowCount As Long
    rowCount = UBound(arr, 1) - LBound(arr, 1) + 1

    Dim lines() As String
    ReDim lines(0 To rowCount)
    lines(0) = CSV_HEADER

    Dim idx As Long
    Dim outRow As Long: outRow = 1
    For idx = LBound(arr, 1) To UBound(arr, 1)
        Dim ts As String: ts = CStr(arr(idx, 1))
        Dim ev As String: ev = CStr(arr(idx, 2))
        Dim md As String: md = CStr(arr(idx, 3))
        Dim dt As String: dt = CStr(arr(idx, 4))
        Dim latencyStr As String: latencyStr = CStr(arr(idx, 5))
        Dim hitStr As String: hitStr = CStr(arr(idx, 6))

        Dim savedMin As Long
        If StrComp(ev, "feedback_green", vbBinaryCompare) = 0 Then
            savedMin = 15
        Else
            savedMin = 0
        End If

        Dim expActivity As Long
        Select Case ev
            Case "ask":         expActivity = expQuestion
            Case "ingest":      expActivity = expRegister
            Case "pack_export": expActivity = expPackShare
            Case Else:          expActivity = 0
        End Select

        Dim expThanks As Long
        If StrComp(ev, "thanks_recv", vbBinaryCompare) = 0 Then
            expThanks = expThumbup
        Else
            expThanks = 0
        End If

        Dim clusterIdText As String: clusterIdText = ""
        Dim srcName As String: srcName = ExtractSourceName(dt)
        If LenB(srcName) > 0 Then
            If clusterMap.Exists(srcName) Then
                clusterIdText = CStr(CLng(clusterMap(srcName)))
            End If
        End If

        lines(outRow) = CsvField(ts) & "," & CsvField(userId) & "," & CsvField(department) & "," & _
            CsvField(ev) & "," & CsvField(md) & "," & CsvField(dt) & "," & _
            CsvField(latencyStr) & "," & CsvField(hitStr) & "," & _
            CsvField(CStr(savedMin)) & "," & CsvField(CStr(expActivity)) & "," & _
            CsvField(CStr(expThanks)) & "," & CsvField(clusterIdText)

        outRow = outRow + 1
    Next idx

    BuildCsvText = Join(lines, vbCrLf)
End Function

' ReadUsageRows - usage_logの2行目～lastRow行(A:F列)をVariant配列で返す。
'   1データ行のみの場合、Range.Valueがスカラーではなく2次元配列で返るよう
'   手動で組み立てる(modDash.CountUsageEvent/modUIDashboard.MonthlyAskCountと
'   同じ慣習)。
Private Function ReadUsageRows(ByVal ws As Worksheet, ByVal lastRow As Long) As Variant
    Dim arr As Variant
    If lastRow = 2 Then
        Dim tmp(1 To 1, 1 To 6) As Variant
        Dim c As Long
        For c = 1 To 6
            tmp(1, c) = ws.Cells(2, c).Value
        Next c
        arr = tmp
    Else
        arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 6)).Value
    End If
    ReadUsageRows = arr
End Function

' ExtractSourceName - detail文字列から資料名を抜き出す。
'   "source=<NAME> ..." 形式(ingest)は "source=" の後～次の半角空白まで。
'   "... src=<NAME>" 形式(thanks_emit/thanks_recv)は "src=" の後～末尾まで。
'   どちらも無ければ空文字("q=..."のask行など、資料に紐付かない行)。
Private Function ExtractSourceName(ByVal detail As String) As String
    Dim p As Long
    p = InStr(1, detail, "source=", vbBinaryCompare)
    If p > 0 Then
        Dim rest As String: rest = Mid$(detail, p + 7)
        ' ingestのdetailは "source=<NAME> chunks=<n> status=<s>"。資料名に半角
        ' 空白が含まれても途中で切れないよう、まず " chunks=" を終端に使う。
        Dim term As Long: term = InStr(rest, " chunks=")
        If term > 0 Then
            ExtractSourceName = Left$(rest, term - 1)
        Else
            Dim sp As Long: sp = InStr(rest, " ")
            If sp > 0 Then
                ExtractSourceName = Left$(rest, sp - 1)
            Else
                ExtractSourceName = rest
            End If
        End If
        Exit Function
    End If

    p = InStr(1, detail, "src=", vbBinaryCompare)
    If p > 0 Then
        ExtractSourceName = Mid$(detail, p + 4)
        Exit Function
    End If

    ExtractSourceName = ""
End Function

' CsvField - CSVの1フィールドをエスケープする。カンマ/ダブルクォート/CR/LFを
'   含む場合のみ引用符で囲み、内部のダブルクォートは2個に複製する。
'   CSVインジェクション対策: Excelで開いた際に先頭文字が「=」「+」「-」「@」/
'   タブ/CRだと数式や制御文字として解釈され、意図しない計算実行や表示崩れの
'   リスクになる(いわゆるCSV Injection / Formula Injection)。既存の引用符
'   エスケープを適用する前に、先頭アポストロフィ「'」を付与して文字列強制する
'   ことでExcel側に「これは数式ではない」と伝える(数値列の先頭「-」符号にも
'   付くが、分析ログの用途では実害がないため特別扱いしない)。
Private Function CsvField(ByVal s As String) As String
    If Len(s) > 0 Then
        Dim leadCh As String: leadCh = Left$(s, 1)
        If leadCh = "=" Or leadCh = "+" Or leadCh = "-" Or leadCh = "@" Or _
           leadCh = Chr$(9) Or leadCh = Chr$(13) Then
            s = "'" & s
        End If
    End If

    If InStr(s, ",") > 0 Or InStr(s, """") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then
        CsvField = """" & Replace(s, """", """""") & """"
    Else
        CsvField = s
    End If
End Function

' WriteCsvWithBom - UTF-8(BOM付)でファイル書き出し。保存先が共有/ネットワーク
'   フォルダのとき、アンチウイルスがミリ秒単位でファイルを掴む(実行時エラー70等)
'   ことに耐えるため、modP2P.WriteUtf8Retryと同じく最大3回・WaitMs(DoEvents待機)で
'   リトライする(SRE監査Phase2.1: ガバナンスCSVの保存先も共有パスになり得るため)。
Private Function WriteCsvWithBom(ByVal filePath As String, ByVal content As String) As Boolean
    Dim attempt As Long
    For attempt = 1 To 3
        If TryWriteCsvOnce(filePath, content) Then
            WriteCsvWithBom = True
            Exit Function
        End If
        CsvRetryWait 250 * attempt   ' 250/500/750ms バックオフ(DoEventsで応答性維持)
    Next attempt
End Function

' UTF-8(BOM付)書き出しの実体は modUtilText.WriteTextFileUtf8(2026-07-31
' R11-F2)。BOMは ADODB.Stream の既定どおり付く=Excelが日本語CSVを文字化け
' せずに開ける、という本機能の前提はそのまま維持される。
Private Function TryWriteCsvOnce(ByVal filePath As String, ByVal content As String) As Boolean
    TryWriteCsvOnce = modUtilText.WriteTextFileUtf8(filePath, content)
End Function

' Timer基準の短時間待機(DoEventsで応答性維持。Sleep API宣言を避けbitness非依存)。
Private Sub CsvRetryWait(ByVal ms As Long)
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000# < ms
        DoEvents
        If Timer < t0 Then Exit Do   ' 深夜0時のTimerロールオーバーガード
    Loop
End Sub

' SafeUserId - modP2P.CurrentUserId()のラッパー(ADSystemInfo失敗時も
'   クラッシュさせない防御)。
Private Function SafeUserId() As String
    On Error Resume Next
    SafeUserId = modP2P.CurrentUserId()
    On Error GoTo 0
End Function

' SafeDepartment - modConfig.GetString("user_department","")のラッパー。
Private Function SafeDepartment() As String
    On Error Resume Next
    SafeDepartment = modConfig.GetString("user_department", "")
    On Error GoTo 0
End Function

' SafeClusterMap - modCluster.SourceClusterMap()のラッパー。万一Nothingが
'   返っても(通常は空Dictionaryが返る契約だが)呼び出し側の.Existsが
'   落ちないよう空Dictionaryへフォールバックする。
Private Function SafeClusterMap() As Object
    On Error Resume Next
    Set SafeClusterMap = modCluster.SourceClusterMap()
    On Error GoTo 0
    If SafeClusterMap Is Nothing Then Set SafeClusterMap = CreateObject("Scripting.Dictionary")
End Function
