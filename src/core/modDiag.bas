Attribute VB_Name = "modDiag"
Option Explicit

' ============================================================================
' modDiag - 自己診断(🩺診断ボタン)
' ----------------------------------------------------------------------------
' 役割:
'   「動かない」時に、どこが問題かを非エンジニアにも分かる言葉で一目瞭然に
'   する。diag_reportシートを毎回作り直し(削除→再生成)、結果を表示する。
'
' 設計判断:
'   ・診断は「壊れていても診断自体は落ちない」ことを最優先する。個々の
'     チェックは On Error で個別に保護し、1つの異常が他のチェックを
'     道連れにしないようにする。
'   ・本モジュールは基盤層(src/core)に置かれるため、機能層モジュール
'     (modShelf/modStats等、まだ存在しないWaveで実装される)へは一切
'     依存しない。本棚統計等は my_knowledge/my_manifest シートを直接
'     読んで自己完結させる(依存ルールR1: 基盤層は上位層を呼ばない)。
'   ・QuickHealthCheckはBoot時に軽量に呼ばれる版。問題なければ""を返す。
' ============================================================================

Private Const DIAG_SHEET As String = "diag_report"

' 「🩺診断」ボタン。diag_reportシートを(再)生成して表示する。
Public Sub RunDiagnostics()
    Dim ws As Worksheet: Set ws = RecreateDiagSheet()
    Dim r As Long: r = 1

    WriteLine ws, r, "■ マイ本棚AI 自己診断  " & modUtil.NowStamp(): r = r + 1
    WriteLine ws, r, "この画面の内容がよく分からない場合は、この画面をスクリーンショットして" & _
                     "管理者に送ってください。": r = r + 2

    ' [バージョン]
    WriteLine ws, r, "[バージョン]": r = r + 1
    WriteCheck ws, r, True, "アプリ名: " & modAppDef.APP_NAME & "  version " & modAppDef.APP_VERSION, "": r = r + 1
    r = r + 1

    ' [設定サマリー]
    WriteLine ws, r, "[設定サマリー]": r = r + 1
    On Error Resume Next
    Dim cfgBuildStamp As String: cfgBuildStamp = modConfig.GetString("build_stamp", "(不明)")
    WriteCheck ws, r, True, "  build_stamp: " & cfgBuildStamp, "": r = r + 1
    On Error GoTo 0

    On Error Resume Next
    Dim cfgMockOn As Boolean: cfgMockOn = modConfig.GetBool("mock_llm", True)
    WriteCheck ws, r, True, "  mock_llm: " & IIf(cfgMockOn, "TRUE", "FALSE"), _
               IIf(cfgMockOn, "本番前にFALSEへ", ""): r = r + 1
    On Error GoTo 0

    On Error Resume Next
    Dim cfgEmbedTransport As String: cfgEmbedTransport = modConfig.GetString("embed_transport", "ribbon")
    WriteCheck ws, r, True, "  embed_transport: " & cfgEmbedTransport, "": r = r + 1
    On Error GoTo 0

    On Error Resume Next
    Dim cfgShelfFolder As String: cfgShelfFolder = modConfig.GetString("shelf_folder", "")
    WriteCheck ws, r, True, "  shelf_folder: " & _
               IIf(LenB(cfgShelfFolder) = 0, "(未設定)", modUtil.SafeLeft(cfgShelfFolder, 80)), "": r = r + 1
    On Error GoTo 0

    On Error Resume Next
    Dim cfgSharePath As String: cfgSharePath = modConfig.GetString("nexus_share_path", "")
    WriteCheck ws, r, True, "  nexus_share_path: " & _
               IIf(LenB(cfgSharePath) = 0, "(未設定)", modUtil.SafeLeft(cfgSharePath, 80)), "": r = r + 1
    On Error GoTo 0

    On Error Resume Next
    Dim cfgRecModel As String: cfgRecModel = modConfig.GetString("recommended_model", "")
    WriteCheck ws, r, True, "  recommended_model: " & cfgRecModel, "": r = r + 1
    On Error GoTo 0

    On Error Resume Next
    Dim cfgQuickModel As String: cfgQuickModel = modConfig.GetString("quick_model", "")
    WriteCheck ws, r, True, "  quick_model: " & cfgQuickModel, "": r = r + 1
    On Error GoTo 0

    On Error Resume Next
    Dim cfgSyncInterval As Long: cfgSyncInterval = modConfig.GetLong("sync_interval_min", 0)
    WriteCheck ws, r, True, "  sync_interval_min: " & cfgSyncInterval & _
               IIf(cfgSyncInterval = 0, "(自動同期オフ)", ""), "": r = r + 1
    On Error GoTo 0
    r = r + 1

    ' [AIリボン]
    WriteLine ws, r, "[AIリボン]": r = r + 1
    Dim mockOn As Boolean: mockOn = modConfig.GetBool("mock_llm", True)
    If mockOn Then
        WriteCheck ws, r, True, _
            "mock_llm=TRUE のため、リボン呼び出しは行わずダミー応答で動作しています。", _
            "本番利用前に config の mock_llm を FALSE にしてください。": r = r + 1
    Else
        Dim ribbonOk As Boolean: ribbonOk = modGateway.RibbonAvailable()
        WriteCheck ws, r, ribbonOk, _
            IIf(ribbonOk, "AIリボン(ChatGPT関数)が見つかりました。", "AIリボンが見つかりません。"), _
            IIf(ribbonOk, "", modLog.FriendlyMessage("E0201")): r = r + 1
    End If
    r = r + 1

    ' [シート存在+行数]
    WriteLine ws, r, "[シート]": r = r + 1
    Dim sheetNames() As String
    sheetNames = Split(modAppDef.SH_HOWTO & "|" & modAppDef.SH_HOME & "|" & modAppDef.SH_SHELF & "|" & _
                        modAppDef.SH_DASH & "|" & modAppDef.SH_CONFIG & "|" & modAppDef.SH_KNOWLEDGE & "|" & _
                        modAppDef.SH_VECTORS & "|" & modAppDef.SH_MANIFEST & "|" & modAppDef.SH_STATS & "|" & _
                        modAppDef.SH_USAGE & "|" & modAppDef.SH_ERRLOG & "|" & modAppDef.SH_UISTATE, "|")
    Dim i As Long
    For i = LBound(sheetNames) To UBound(sheetNames)
        Dim nm As String: nm = sheetNames(i)
        Dim existsFlag As Boolean: existsFlag = SheetExists(nm)
        Dim rowsN As Long: rowsN = 0
        If existsFlag Then rowsN = LastRow(nm)
        WriteCheck ws, r, existsFlag, "  " & nm & " : " & IIf(existsFlag, "存在(" & rowsN & "行)", "見つかりません"), _
                   IIf(existsFlag, "", modLog.FriendlyMessage("E0101")): r = r + 1
    Next i
    r = r + 1

    ' [config必須キー]
    WriteLine ws, r, "[設定(config)]": r = r + 1
    Dim reqKeys() As String
    reqKeys = Split("mock_llm|recommended_model|quick_model|embed_dim|shelf_max_chunks|" & _
                     "topk_quick|topk_deep|answer_language|pack_author", "|")
    Dim cfgExists As Boolean: cfgExists = SheetExists(modAppDef.SH_CONFIG)
    For i = LBound(reqKeys) To UBound(reqKeys)
        Dim k As String: k = reqKeys(i)
        Dim found As Boolean: found = cfgExists And ConfigKeyExists(k)
        WriteCheck ws, r, found, "  " & k & IIf(found, " : 設定あり", " : 見つかりません"), _
                   IIf(found, "", modLog.FriendlyMessage("E0101")): r = r + 1
    Next i
    r = r + 1

    ' [本棚の状態]
    WriteLine ws, r, "[本棚の状態]": r = r + 1
    Dim fileCount As Long, chunkCount As Long, pendingCount As Long
    ShelfCounts fileCount, chunkCount, pendingCount
    WriteCheck ws, r, True, "登録資料数(manifest): " & fileCount & " 件", "": r = r + 1
    WriteCheck ws, r, True, "チャンク数(my_knowledge): " & chunkCount & " 件", "": r = r + 1
    WriteCheck ws, r, (pendingCount = 0), _
        "AIが読める形にまだ変換していないチャンク: " & pendingCount & " 件", _
        IIf(pendingCount = 0, "", "「マイ本棚」タブの" & ChrW(&HD83D) & ChrW(&HDD04) & "同期を押すと続きから変換されます。"): r = r + 1
    r = r + 1

    ' [追加機能(opt)]
    WriteLine ws, r, "[追加機能(opt)]": r = r + 1
    Dim featIds() As String: featIds = Split("tts|vision|markdown|diffdoc", "|")
    For i = LBound(featIds) To UBound(featIds)
        Dim fid As String: fid = featIds(i)
        Dim present As Boolean: present = modFeatures.ModulePresent(fid)
        Dim enabled As Boolean: enabled = modFeatures.FeatureEnabled(fid)
        WriteCheck ws, r, True, "  " & fid & " : モジュール" & IIf(present, "あり", "なし") & _
                   " / 設定" & IIf(enabled, "有効", "無効"), "": r = r + 1
    Next i
    r = r + 1

    ' [直近のエラー]
    WriteLine ws, r, "[直近のエラー(最大5件)]": r = r + 1
    Dim errLines() As String
    Dim errN As Long: errN = RecentErrors(errLines)
    If errN = 0 Then
        WriteCheck ws, r, True, "エラーはありません。", "": r = r + 1
    Else
        For i = 0 To errN - 1
            WriteCheck ws, r, False, "  " & errLines(i), "": r = r + 1
        Next i
    End If
    r = r + 2

    WriteLine ws, r, "以上です。この画面をスクリーンショットして管理者に送ってください。"

    ws.Columns("A").ColumnWidth = 110

    On Error Resume Next
    ws.Activate
    On Error GoTo 0
End Sub

' Boot時の軽量版。問題なければ"" / あれば警告文を返す。実行時例外は出さない。
Public Function QuickHealthCheck() As String
    Dim msgs As String

    If Not SheetExists(modAppDef.SH_CONFIG) Then msgs = msgs & "・configシートが見つかりません。" & vbLf
    If Not SheetExists(modAppDef.SH_KNOWLEDGE) Then msgs = msgs & "・my_knowledgeシートが見つかりません。" & vbLf
    If Not SheetExists(modAppDef.SH_VECTORS) Then msgs = msgs & "・my_vectorsシートが見つかりません。" & vbLf
    If Not SheetExists(modAppDef.SH_MANIFEST) Then msgs = msgs & "・my_manifestシートが見つかりません。" & vbLf

    QuickHealthCheck = msgs
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' 2026-07-28(レビュー L-6): 追加と改名を1つの手続きとして扱う。
' 従来は Delete が失敗しても構わず Add し、直後の Name 代入が
' 「同名シートが既にある」で失敗すると、Sheet1 のような既定名の
' シートが可視のまま残っていた。modBoot.RemoveOrphanDefaultSheets が
' 「原因未特定」と注記して毎起動これを掃除しているが、その発生源が
' ここである可能性が高い。改名に失敗したら、作ったシートを片付けてから
' 既存シートを使い回す(孤児を残さない)。
Private Function RecreateDiagSheet() As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(DIAG_SHEET).Delete
    Err.Clear
    On Error GoTo 0
    Application.DisplayAlerts = True

    Dim ws As Worksheet
    On Error GoTo Fallback
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    ws.Name = DIAG_SHEET
    On Error GoTo 0
    Set RecreateDiagSheet = ws
    Exit Function

Fallback:
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FallbackCleanup2
FallbackCleanup2:
    ' 改名できなかった。作ったばかりのシートは消して、既にある
    ' 診断シートを再利用する(中身は呼び出し側が上書きする)。
    On Error Resume Next
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    Set ws = ThisWorkbook.Worksheets(DIAG_SHEET)
    If Not ws Is Nothing Then ws.Cells.ClearContents
    Err.Clear
    On Error GoTo 0
    Set RecreateDiagSheet = ws
End Function

Private Sub WriteLine(ByVal ws As Worksheet, ByVal r As Long, ByVal text As String)
    ws.Cells(r, 1).Value = modUtil.SafeLeft(text, 500)
End Sub

Private Sub WriteCheck(ByVal ws As Worksheet, ByVal r As Long, ByVal ok As Boolean, _
                       ByVal text As String, ByVal howTo As String)
    Dim mark As String: mark = IIf(ok, ChrW(&H2705) & " ", ChrW(&H26A0) & ChrW(&HFE0F) & " ")
    Dim line As String: line = mark & text
    If LenB(howTo) > 0 Then line = line & "  " & ChrW(&H2192) & " " & howTo
    ws.Cells(r, 1).Value = modUtil.SafeLeft(line, 500)
End Sub

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function

Private Function LastRow(ByVal nm As String) As Long
    On Error GoTo Fail
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(nm)
    LastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Exit Function
Fail:
    LastRow = 0
End Function

Private Function ConfigKeyExists(ByVal key As String) As Boolean
    On Error GoTo Fail
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(modAppDef.SH_CONFIG)
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim i As Long
    For i = 2 To lastR
        If StrComp(CStr(ws.Cells(i, 1).Value), key, vbTextCompare) = 0 Then
            ConfigKeyExists = True
            Exit Function
        End If
    Next i
    Exit Function
Fail:
    ConfigKeyExists = False
End Function

' 資料数(manifest行数)/チャンク数(my_knowledge行数)/未変換件数(embedded<>1)。
' my_knowledge列構成(MASTER_SPEC §4): chunk_id,source,origin,page,summary,
' keywords,full_text,added_at,embedded  ->  embeddedは9列目。
Private Sub ShelfCounts(ByRef fileCount As Long, ByRef chunkCount As Long, ByRef pendingCount As Long)
    fileCount = 0: chunkCount = 0: pendingCount = 0
    On Error GoTo Done

    If SheetExists(modAppDef.SH_MANIFEST) Then
        Dim wsM As Worksheet: Set wsM = ThisWorkbook.Worksheets(modAppDef.SH_MANIFEST)
        Dim lastM As Long: lastM = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
        If lastM >= 2 Then fileCount = lastM - 1
    End If

    If SheetExists(modAppDef.SH_KNOWLEDGE) Then
        Dim wsK As Worksheet: Set wsK = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
        Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
        If lastK >= 2 Then
            chunkCount = lastK - 1
            If lastK = 2 Then
                If CStr(wsK.Cells(2, 9).Value) <> "1" Then pendingCount = pendingCount + 1
            Else
                Dim arr As Variant
                arr = wsK.Range(wsK.Cells(2, 9), wsK.Cells(lastK, 9)).Value
                Dim i As Long
                For i = LBound(arr, 1) To UBound(arr, 1)
                    If CStr(arr(i, 1)) <> "1" Then pendingCount = pendingCount + 1
                Next i
            End If
        End If
    End If
Done:
End Sub

' err_logシート末尾から最大5件を新しい順に取り出す。
Private Function RecentErrors(ByRef lines() As String) As Long
    ReDim lines(0 To 4)
    If Not SheetExists(modAppDef.SH_ERRLOG) Then
        RecentErrors = 0
        Exit Function
    End If

    On Error GoTo Fail
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(modAppDef.SH_ERRLOG)
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then
        RecentErrors = 0
        Exit Function
    End If

    Dim startR As Long: startR = lastR - 4
    If startR < 2 Then startR = 2

    Dim n As Long: n = 0
    Dim i As Long
    For i = lastR To startR Step -1
        Dim ts As String: ts = CStr(ws.Cells(i, 1).Value)
        Dim code As String: code = CStr(ws.Cells(i, 2).Value)
        Dim ctx As String: ctx = CStr(ws.Cells(i, 3).Value)
        lines(n) = ts & "  " & code & "  " & ctx
        n = n + 1
    Next i
    ReDim Preserve lines(0 To n - 1)
    RecentErrors = n
    Exit Function
Fail:
    RecentErrors = 0
End Function

' ----------------------------------------------------------------------------
' RecentErrorsForClipboard - 🩺診断シートの「📋直近のエラーをコピー」ボタン
' (modUIMain.OnCopyRecentErrors)専用。err_logの直近5件を、コード・モジュール/
' 関数名(context)・生エラー番号・HTTPステータス・詳細まで含めて、そのまま
' 開発者へ貼り付けられる整形テキストにして返す。クリップボードへの実書込み
' 自体はUI層(modClip)の責務のため、ここでは文字列を返すだけに留める
' (R1: 基盤層は上位層を呼ばない)。
' ----------------------------------------------------------------------------
Public Function RecentErrorsForClipboard() As String
    Dim blocks() As String
    Dim n As Long: n = RecentErrorsDetailed(blocks)
    If n = 0 Then
        RecentErrorsForClipboard = "直近のエラーはありません。"
        Exit Function
    End If

    Dim out As String
    out = modAppDef.APP_NAME & " v" & modAppDef.APP_VERSION & "  直近のエラー(" & n & "件・新しい順)" & vbCrLf & _
          "========================================" & vbCrLf
    Dim i As Long
    For i = 0 To n - 1
        out = out & blocks(i) & vbCrLf & "----------------------------------------" & vbCrLf
    Next i
    RecentErrorsForClipboard = out
End Function

' err_logシート末尾から最大5件を新しい順に取り出す(コピー用・詳細版)。
' 1件=複数行のブロック文字列(timestamp/code/context/err_number/http_status/detail)。
Private Function RecentErrorsDetailed(ByRef blocks() As String) As Long
    ReDim blocks(0 To 4)
    If Not SheetExists(modAppDef.SH_ERRLOG) Then
        RecentErrorsDetailed = 0
        Exit Function
    End If

    On Error GoTo Fail
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(modAppDef.SH_ERRLOG)
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then
        RecentErrorsDetailed = 0
        Exit Function
    End If

    Dim startR As Long: startR = lastR - 4
    If startR < 2 Then startR = 2

    Dim n As Long: n = 0
    Dim i As Long
    For i = lastR To startR Step -1
        Dim ts As String: ts = CStr(ws.Cells(i, 1).Value)
        Dim code As String: code = CStr(ws.Cells(i, 2).Value)
        Dim ctx As String: ctx = CStr(ws.Cells(i, 3).Value)
        Dim detail As String: detail = CStr(ws.Cells(i, 4).Value)
        ' err_number/http_status(F/G列)は旧バージョンで作られたシートには
        ' 存在しない場合があるため、読み取り自体を1行スコープで保護する。
        Dim errNum As String: errNum = ""
        Dim httpStatus As String: httpStatus = ""
        On Error Resume Next
        errNum = CStr(ws.Cells(i, 6).Value)
        httpStatus = CStr(ws.Cells(i, 7).Value)
        On Error GoTo Fail

        Dim block As String
        block = "[" & ts & "] " & code & "  " & ctx
        If LenB(errNum) > 0 And errNum <> "0" Then block = block & "  (err#" & errNum & ")"
        If LenB(httpStatus) > 0 And httpStatus <> "0" Then block = block & "  (HTTP " & httpStatus & ")"
        block = block & vbCrLf & detail

        blocks(n) = block
        n = n + 1
    Next i
    ReDim Preserve blocks(0 To n - 1)
    RecentErrorsDetailed = n
    Exit Function
Fail:
    RecentErrorsDetailed = 0
End Function
