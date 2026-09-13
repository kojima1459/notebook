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
    Dim cfgMockOn As Boolean: cfgMockOn = modConfig.GetBool("mock_llm", False)
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

    ' [共有フォルダ](R13-7b / R33 W4-4改め W4-5): 診断は実行のたびに実際に
    ' 確かめる(「診断は都度、通常描画はキャッシュを使い回す」の使い分け)。
    '
    ' 【なぜ Reachable() ではなく ProbePath() を呼ぶのか】(2026-08-16 R33 W4-5)
    ' modShare.Reachable は mState が確定していれば即座に前回値を返す
    ' (modShare.bas:71-76。NG→OK の昇格経路も TTL も無い)。そして mState は
    ' 起動シーケンスで必ず解決される(modBoot:279→modHub の
    ' RefreshOrgTilesIfReachable、modBoot:406 の modGuard.TouchReach)。
    ' つまり利用者が🩺診断を押す時点で判定は例外なく確定済みで、ここは
    ' 一度もプローブせずキャッシュを読むだけだった。結果、
    '   「VPN未接続で開く → VPNを繋ぐ → 診断を押す」= 何度押しても
    '   『到達できません → VPN接続をご確認ください』
    ' となり、原因を直した利用者が解決済みのネットワークを疑い続ける。
    ' この画面は「スクリーンショットして管理者に送ってください」と案内する
    ' 画面なので、古い結論がそのまま管理者側の一次情報になっていた。
    '
    ' ProbePath はキャッシュを持たない(modShare.bas:130)ので都度実測になり、
    ' 判定式は設定画面(modHelp.OnShareSetup)と共用の1本のままで増えない。
    ' ResetProbe は呼ばない ―― 呼ぶと診断の副作用で他モジュールが使う
    ' セッションキャッシュまで捨ててしまう(セッションキャッシュの方針は
    ' 変えない、が R13-7b の裁定)。BasePath() が空(=未設定)のときは
    ' ProbePath が即 False を返し、従来どおり「未設定」表示になる。
    WriteLine ws, r, "[共有フォルダ]": r = r + 1
    On Error Resume Next
    Dim shareConfigured As Boolean: shareConfigured = (LenB(cfgSharePath) > 0)
    Dim shareReach As Boolean: shareReach = modShare.ProbePath(modShare.BasePath())
    Dim shareOk As Boolean: shareOk = (Not shareConfigured) Or shareReach
    WriteCheck ws, r, shareOk, _
        "  到達性: " & IIf(Not shareConfigured, "未設定(共有機能は休止中)", _
                            IIf(shareReach, "到達できました", "到達できません")), _
        IIf(shareConfigured And Not shareReach, "ネットワークまたはVPN接続をご確認ください", ""): r = r + 1
    On Error GoTo 0
    r = r + 1

    ' [AIリボン]
    WriteLine ws, r, "[AIリボン]": r = r + 1
    Dim mockOn As Boolean: mockOn = modConfig.GetBool("mock_llm", False)
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
    ' 2026-08-16(R33波5c C-2): nexus_ui を必須キーに追加。R33波4(W4-2)で
    ' Python側TRUE / VBA側FALSE の既定値不一致を揃えたが、キー自体が
    ' config から欠けている端末では既定値に落ちて画面が旧UIになる。
    ' その欠落が診断画面のどこにも出ていなかった(=「見た目が違う」を
    ' 診断票から切り分けられない)ため、ここへ載せる。
    reqKeys = Split("mock_llm|recommended_model|quick_model|embed_dim|shelf_max_chunks|" & _
                     "topk_quick|topk_deep|answer_language|pack_author|nexus_ui", "|")
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

    ' ------------------------------------------------------------------
    ' [余白の敷き詰め] 2026-08-16(R33H F24)
    '   R33 DoD 第8項「条件付き書式が UsedRange を膨らませないことをテストで
    '   固定する」は、実装が【実行時の自己検算】(modBackdrop.ApplyCF が張った
    '   直後に UsedRange を測り、膨らんでいたら剥がして巻き戻す)へ置き換わった
    '   ため、自動テストでは達成できない(FormatConditions は LO では検証不能)。
    '   その代わりに、非エンジニアが「どちらの手が効いたのか」を1画面で
    '   判定できるようにする ―― 成立したら backdrop_cf が、不首尾なら
    '   backdrop_cf_failed が usage_log に1行だけ残る設計なので、その直近1行を
    '   そのまま見せる。実機スモークの確認項目はこの2行を読むだけで済む。
    WriteLine ws, r, "[余白の敷き詰め(条件付き書式)]": r = r + 1
    Dim cfOk As String: cfOk = LastUsageLine("backdrop_cf")
    Dim cfNg As String: cfNg = LastUsageLine("backdrop_cf_failed")
    WriteCheck ws, r, (LenB(cfOk) > 0), _
        "  成立(backdrop_cf): " & IIf(LenB(cfOk) > 0, cfOk, "記録なし"), "": r = r + 1
    WriteCheck ws, r, (LenB(cfNg) = 0), _
        "  不首尾(backdrop_cf_failed): " & IIf(LenB(cfNg) > 0, cfNg, "記録なし"), _
        IIf(LenB(cfNg) > 0, "この端末では条件付き書式が使えていません。" & _
            "背景画像の方(backdrop_failed が無ければ成功)で余白が塗られます。", ""): r = r + 1
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

    ' R11-C(lint許容登録): Activate失敗を致命的にしない(診断結果は既に書き
    ' 終わっているため、前面化に失敗しても続行する)。ログ付き許容続行。
    On Error Resume Next
    Err.Clear
    ws.Activate   ' lint:allow-raw-activate(許容続行・R11-C裁定)
    If Err.Number <> 0 Then
        modLog.LogError "E0801", "modDiag.RunDiagnostics", "[許容続行] 診断シートActivate失敗", Err.Number
    End If
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
' WarnIfReadOnly - このブックが読み取り専用で開かれていたら1度だけ知らせる
'   (2026-08-04 R15-3b・実機第4報 RC9)。
' ----------------------------------------------------------------------------
' 読み取り専用でも取込も質問も普通に動く。ところが結果は1つも残らず、
' 閉じた瞬間に数時間ぶんの取込が消える(憲章§3-5「利用者の資産を失わない」)。
' 症状として現れるのは「昨日入れた資料が今日は無い」だけで、原因に辿り着く
' 手がかりが1つも無い。起動時に1回だけ、目立つ形で伝える。
' 機能は止めない(閲覧・質問はそのまま使えるため。§3-1)。
' 置き場を modBoot ではなく診断側にしたのは、modBoot が30,000字上限まで
' 残り僅かで1行も足せないため(判定内容も「起動時の環境チェック」で
' QuickHealthCheck と同種)。表示・記録まで含めてここで完結させる。
' 例外は外へ出さない(起動を止めない)。
Public Sub WarnIfReadOnly()
    On Error Resume Next
    ' 判定を変数へ受けてから見る: OERN配下でIf条件の評価そのものが失敗すると
    ' 次の文へ進んでしまい、読み取り専用でもないのに警告を出しかねない。
    Dim isReadOnly As Boolean: isReadOnly = False
    isReadOnly = ThisWorkbook.ReadOnly
    If Not isReadOnly Then Exit Sub
    modLog.LogError "E0805", "modDiag.WarnIfReadOnly", _
        "ReadOnly=True name=" & ThisWorkbook.Name
    ' R15-FixB(FB-13・レビューB-L): どのファイルの話なのかを必ず添える。
    ' 読み取り専用になる典型は「zipの中から直接開いた」「共有フォルダの原本を
    ' 誰かが開いている」「Downloadsの保護ビュー」で、いずれも【場所を見れば
    ' 一目で分かる】。名前しか出さないと、同名のコピーを複数持っている人は
    ' どれを開いているのか確かめる手段がなく、案内どおりに直しようがない。
    MsgBox modLog.ReadOnlyWarnMsg() & vbLf & vbLf & _
        "このファイルの場所: " & ThisWorkbook.FullName, _
        vbExclamation, modAppDef.APP_NAME
    On Error GoTo 0
End Sub

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
' ----------------------------------------------------------------------------
' LastUsageLine - usage_log の直近1行を "時刻 詳細" で返す(2026-08-16 R33H F24)。
'   event_name(B列)が一致する【最も新しい】1行だけを見る。見つからなければ ""。
'   下から上へ走るのは、usage_log が数千行になっても数行しか読まないため。
'   例外は外へ出さない(診断画面が1項目のせいで丸ごと出ないことを防ぐ)。
' ----------------------------------------------------------------------------
Private Function LastUsageLine(ByVal eventName As String) As String
    If Not SheetExists(modAppDef.SH_USAGE) Then Exit Function
    On Error GoTo UsgFail
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(modAppDef.SH_USAGE)
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then Exit Function
    Dim i As Long
    For i = lastR To 2 Step -1
        If StrComp(CStr(ws.Cells(i, 2).Value), eventName, vbBinaryCompare) = 0 Then
            LastUsageLine = CStr(ws.Cells(i, 1).Value) & "  " & _
                            modUtil.SafeLeft(CStr(ws.Cells(i, 4).Value), 160)
            Exit Function
        End If
    Next i
    Exit Function
UsgFail:
    LastUsageLine = ""
End Function

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
