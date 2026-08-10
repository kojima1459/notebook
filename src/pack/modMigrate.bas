Attribute VB_Name = "modMigrate"
Option Explicit

' ============================================================================
' modMigrate - 新しい版のブックへ「自分の資産」を引き継ぐ
' ----------------------------------------------------------------------------
' 解決する問題(解説書 §12.3 B3):
'   このアプリは全データをブック内のシートに持つ。つまり新しい .xlsm へ
'   差し替えると、その人が積み上げたものが丸ごと消える。
'     ・本棚の中身(my_knowledge / my_vectors)
'     ・資料台帳(my_manifest) … これが無いと本棚一覧に出ず、
'       フォルダ同期・個別削除の対象からも外れる
'     ・EXP / バッジ / 連続利用日数 / 節約時間(my_stats)
'     ・共有パス・部署・同期フォルダ等の設定(config)
'   PoC 中は不具合修正版を何度も配る。配るたびに全員の資産が消えるなら、
'   誰も本気で資料を入れない。「壊れたら作り直せばいい」で済むのは
'   作った本人だけである。
'
' 方針:
'   1ファイル(.xlsx)に固めて渡す。新しいブックで読み込めば元に戻る。
'   パック(modPack)を流用しない理由は、パックが「他人に配る知識」の器で
'   あって「自分の環境まるごと」ではないため。実際パック経由の移行では
'   origin が self から pack: に変わり、my_manifest も統計も運べない。
'
' 安全のための設計判断:
'   ・config は【上書きしない】。新しいブックに存在するキーだけ、
'     値を引き継ぐ。新版で増えたキーは新版の既定のままにする
'     (古い設定で新機能を殺さないため)。
'   ・config のうち【ビルドが決める値】は絶対に引き継がない。
'     mock_llm / build_stamp / publish_key / azure_embed_* がそれで、
'     ここを運ぶと「利用者用ブックに発行キーが復活する」
'     「本番ビルドにモックが復活する」という事故になる(§12.3 B1/B2)。
'   ・my_stats は【併合】する(キー単位で上書き)。新版が先に書いた
'     バッジ等を消さないため。
'   ・本棚(my_knowledge / my_vectors / my_manifest)は【置換】する。
'     移行先は新品のはずで、混ぜる方が危ない。
'   ・書き出しは PII 走査を通さない。これは配布物ではなく本人の控えで、
'     本人の資料が本人の手元へ戻るだけだから。ただしファイル自体は
'     本棚と同じ機微を持つので、その旨をダイアログで明示する。
' ============================================================================

Private Const MIG_META As String = "mig_meta"
Private Const MIG_FORMAT_VERSION As Long = 1

' 2026-08-10(R27波3-4): 保護されたブックを掴んでもパスワード入力ダイアログを
' 出さないためのダミー(modExtractorExcel と同型)。正しいはずがないので、
' 保護されていれば即エラーになり Failed 経路で扱える。
Private Const DUMMY_PASSWORD As String = "__mybookshelf_no_password__"

' 引き継がない config キー(ビルドが決める値・秘密)。
' 前方一致で判定するので azure_embed_key / azure_embed_url は "azure_" で拾う。
Private Const CFG_SKIP_EXACT As String = "|mock_llm|build_stamp|publish_key|"
Private Const CFG_SKIP_PREFIX As String = "azure_"

' ----------------------------------------------------------------------------
' 引き継ぐ config キーのホワイトリスト(2026-08-01 R12-3-2)
' ----------------------------------------------------------------------------
' なぜ「引き継がないものを除く」から「引き継ぐものだけ」へ変えたか:
'   従来は IsBuildOwnedKey に載っていない全キーを新版へ書き戻していた。
'   その結果、新版で【既定値を直した】キー(例: knowledge_expire_days を
'   30→0 に直した R12-1-1、shelf_max_chunks・chunk_limit の上限統一)が、
'   引き継ぎのたびに旧ブックの値で静かに巻き戻る。直したはずの不具合が
'   移行した人にだけ再発する、という最悪の形の劣化になる。
'   「新版が正しい」を既定に置き、利用者が自分で決めた値だけを運ぶ。
'
' 何を載せたか(configシートの説明文からの判断):
'   (a) 実行時に利用者操作で書き換わるキー = 画面から設定できるもの。
'       modConfig.SetValue の呼び出し元を全数確認して列挙した:
'         nexus_share_path(modHelp)/ pack_author(modBoot)/
'         answer_language(modApp)/ shelf_folder(modShelfSync)/
'         ghostscript_path(optVision)/ active_channel・
'         unsubscribed_channels(modChannel)
'   (b) 端末ごとに人が書き入れる環境値・好み。
'         user_department(分析用の自己申告)/ ghostscript_search_dirs
'         (社内の標準配置先)/ sync_interval_min・sync_on_open(同期の好み)/
'         chat_log_enabled(履歴を残すか)/ telemetry_enabled(送信の可否)/
'         insight_share_enabled(2026-08-01 R12-2-2追加。共有知の送信可否。
'         telemetry_enabledと同じ「送信の可否」トグルであり、載せないと
'         利用者がFALSEにした個人の意思表示が移行のたびにTRUEへ巻き戻る)
'
' 何を載せなかったか(意図的):
'   ・モデル名・effort・topk・チャンク設定・検索方式などのチューニング値は
'     製品の作り手が決める。新版の既定を優先する。
'   ・allowed_domain / knowledge_expire_days / admin_users /
'     domain_wipe_after_n_boots / feedback_mail_to / noise_global_threshold は
'     配布する側(IT・管理者)がビルドまたは配布時に決める運用値。旧端末の値を
'     持ち込むと、新しい運用方針が端末単位で無効化される。
'   ・feature_* / holidays / exp_* も同じ理由で新版の値を使う。
Private Const CFG_KEEP_KEYS As String = _
    "|nexus_share_path|shelf_folder|sync_interval_min|sync_on_open|" & _
    "user_department|pack_author|answer_language|active_channel|" & _
    "unsubscribed_channels|ghostscript_path|ghostscript_search_dirs|" & _
    "chat_log_enabled|telemetry_enabled|insight_share_enabled|"

' ----------------------------------------------------------------------------
' ExportUserData - 引き継ぎファイルを書き出す。
' ----------------------------------------------------------------------------
' 呼び出し側(UI層)が modUiLock を取ってから呼ぶこと。R1(層の向き)により
' 本モジュールからUI層は参照しない。導線は modMigrateUI にある。
Public Sub ExportUserData()
    On Error GoTo Failed

    Dim fd As Object
    Set fd = Application.FileDialog(2)   ' msoFileDialogSaveAs
    fd.Title = "引き継ぎファイルの保存先を選んでください"
    fd.InitialFileName = "MyBookshelf_引き継ぎ_" & Format$(Now, "yyyymmdd-hhnn") & ".xlsx"
    If fd.Show <> -1 Then GoTo Done
    If fd.SelectedItems.count < 1 Then GoTo Done
    Dim savePath As String: savePath = CStr(fd.SelectedItems(1))
    If LCase$(modUtil.ExtOf(savePath)) <> "xlsx" Then savePath = savePath & ".xlsx"

    Application.Cursor = 2
    Application.StatusBar = "引き継ぎファイルを書き出しています..."

    Dim prevAlerts As Boolean: prevAlerts = Application.DisplayAlerts
    Dim prevScreen As Boolean: prevScreen = Application.ScreenUpdating
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim wb As Workbook
    Set wb = Application.Workbooks.Add
    ' 追加直後のブックは既定シートが1枚。名前を付けてメタに使う。
    wb.Worksheets(1).Name = MIG_META
    WriteMeta wb.Worksheets(MIG_META)

    CopySheetInto wb, modAppDef.SH_KNOWLEDGE
    CopySheetInto wb, modAppDef.SH_VECTORS
    CopySheetInto wb, modAppDef.SH_MANIFEST
    CopySheetInto wb, modAppDef.SH_STATS
    CopyConfigInto wb
    CopyInsightInboxInto wb

    wb.SaveAs Filename:=savePath, FileFormat:=51   ' xlOpenXMLWorkbook(.xlsx)
    wb.Close SaveChanges:=False
    Set wb = Nothing

    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    Application.Cursor = -4143
    Application.StatusBar = False

    On Error Resume Next
    modLog.LogUsage "migrate_export", "", modUtil.SafeLeft(savePath, 200)
    On Error GoTo Failed

    MsgBox "引き継ぎファイルを書き出しました。" & vbCrLf & vbCrLf & _
           savePath & vbCrLf & vbCrLf & _
           "新しい版のブックを開いて「引き継ぎファイルを読み込む」を実行すると、" & vbCrLf & _
           "本棚・資料一覧・これまでの記録・設定が元に戻ります。" & vbCrLf & vbCrLf & _
           "※このファイルには本棚の中身がそのまま入っています。" & vbCrLf & _
           "　資料と同じ扱いで保管してください。", _
           vbInformation, modAppDef.APP_NAME

Done:
    On Error Resume Next
    Application.Cursor = -4143
    Application.StatusBar = False
    On Error GoTo 0
    Exit Sub

Failed:
    Dim desc As String: desc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailedCleanup0
FailedCleanup0:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Cursor = -4143
    Application.StatusBar = False
    modLog.LogError "E0801", "modMigrate.ExportUserData", desc
    MsgBox "引き継ぎファイルを書き出せませんでした。" & vbCrLf & _
           "保存先に書き込む権限があるかご確認ください。" & vbCrLf & "(" & desc & ")", _
           vbExclamation, modAppDef.APP_NAME
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' ImportUserData - 引き継ぎファイルを読み込む。
' ----------------------------------------------------------------------------
' 呼び出し側(UI層)が modUiLock を取ってから呼ぶこと(ExportUserData と同じ)。
Public Sub ImportUserData()
    On Error GoTo Failed

    ' 本棚を置き換える操作なので、必ず確認を取る。
    Dim have As Long
    On Error Resume Next
    have = modShelf.TotalChunks()
    On Error GoTo Failed
    If have > 0 Then
        If MsgBox("今この本棚に入っている " & have & " 件は、読み込んだ内容で" & vbCrLf & _
                  "【置き換わります】。よろしいですか?" & vbCrLf & vbCrLf & _
                  "(引き継ぎは、新しい版のブックを開いた直後に行う操作です)", _
                  vbOKCancel + vbExclamation, modAppDef.APP_NAME) <> vbOK Then GoTo Done
    End If

    Dim fd As Object
    Set fd = Application.FileDialog(3)   ' msoFileDialogFilePicker
    fd.AllowMultiSelect = False
    fd.Title = "引き継ぎファイルを選んでください"
    fd.Filters.Clear
    fd.Filters.Add "Excelブック", "*.xlsx"
    If fd.Show <> -1 Then GoTo Done
    If fd.SelectedItems.count < 1 Then GoTo Done

    Application.Cursor = 2
    Application.StatusBar = "引き継ぎファイルを読み込んでいます..."

    ' 2026-08-10(R27波3-4): 書き出し側(ExportUserData:110-132)と対称に
    ' DisplayAlerts を退避して落とす。読み込み側はこれまで素通しで、
    ' (a)開く時のパスワード/読み取り推奨/リンク更新のダイアログ
    ' (b)RestoreSheet のシート差し替えで出る確認
    ' がモーダルで出得た。引き継ぎは「新しい版を開いた直後」に行う操作なので、
    ' ここで止まると利用者は【最初の一手で固まった】としか受け取れない。
    Dim prevAlerts As Boolean: prevAlerts = Application.DisplayAlerts
    Dim prevScreen As Boolean: prevScreen = Application.ScreenUpdating
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ' ダミーPassword/IgnoreReadOnlyRecommended は modExtractorExcel:181-186 と
    ' 同型。保護されたファイルを選ばれても入力を求めず即エラーへ倒す
    ' (ここは他人から受け取ったファイルを開く経路なので必ず要る)。
    Dim wb As Workbook
    Set wb = Application.Workbooks.Open(Filename:=CStr(fd.SelectedItems(1)), _
                                        ReadOnly:=True, UpdateLinks:=0, _
                                        IgnoreReadOnlyRecommended:=True, AddToMru:=False, _
                                        Password:=DUMMY_PASSWORD, _
                                        WriteResPassword:=DUMMY_PASSWORD)

    Dim reason As String
    If Not ValidateMigFile(wb, reason) Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
        Application.DisplayAlerts = prevAlerts
        Application.ScreenUpdating = prevScreen
        Application.Cursor = -4143
        Application.StatusBar = False
        MsgBox "このファイルは引き継ぎファイルではないようです。" & vbCrLf & _
               "(" & reason & ")", vbExclamation, modAppDef.APP_NAME
        GoTo Done
    End If

    ' R12-3-1: 3シートそれぞれの成否と実測行数を受け取る。1枚でも失敗したら
    ' 完了ダイアログの文言を切り替える(「引き継ぎました」と言わない)。
    Dim rowsK As Long, rowsV As Long, rowsM As Long
    Dim allOk As Boolean: allOk = True
    If Not RestoreSheet(wb, modAppDef.SH_KNOWLEDGE, rowsK) Then allOk = False
    If Not RestoreSheet(wb, modAppDef.SH_VECTORS, rowsV) Then allOk = False
    If Not RestoreSheet(wb, modAppDef.SH_MANIFEST, rowsM) Then allOk = False
    MergeStats wb
    MergeConfig wb
    MergeInsightInbox wb

    wb.Close SaveChanges:=False
    Set wb = Nothing
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    Application.Cursor = -4143
    Application.StatusBar = False

    Dim got As Long
    On Error Resume Next
    got = modShelf.TotalChunks()
    ' R12-3-1: 「何件戻したつもりか」ではなく、復元後のシートを数えた実測値を残す。
    modLog.LogUsage "migrate_import", "", "chunks=" & got & _
        " knowledge=" & rowsK & " vectors=" & rowsV & " manifest=" & rowsM & _
        " restored_ok=" & allOk
    On Error GoTo Failed

    If allOk Then
        MsgBox "引き継ぎました。" & vbCrLf & vbCrLf & _
               "本棚: " & got & " 件" & vbCrLf & vbCrLf & _
               "設定は、あなたが決めた項目(共有フォルダ・同期フォルダ・部署名・" & vbCrLf & _
               "作成者名・部門チャンネル等)だけを引き継いでいます" & vbCrLf & _
               "(その他はこの版の初期値のままです)。" & vbCrLf & _
               "画面を描き直すため、一度閉じて開き直してください。", _
               vbInformation, modAppDef.APP_NAME
    Else
        ' 2026-08-10(R27波3-12): 「元のファイル」は選んだ引き継ぎファイルの
        ' ことなのに、「この本棚は元のまま(無事)」と読める文だった。実際は
        ' 本棚の3シートを1枚ずつ置換していく途中で失敗しているので、この本棚は
        ' 既に入れ替わりかけている。何が残っていて何が変わり得たかを分けて言う。
        MsgBox "一部復元できませんでした。" & vbCrLf & _
               "引き継ぎ元のファイルは残っています" & vbCrLf & _
               "(この本棚の一部は入れ替わった可能性があります)。" & vbCrLf & _
               "もう一度お試しください。" & vbCrLf & vbCrLf & _
               "いま入っているのは 本棚: " & got & " 件 です。" & vbCrLf & _
               "続けて同じ結果になるときは、診断ボタンの" & vbCrLf & _
               "「直近のエラーをコピー」で担当者へご連絡ください。", _
               vbExclamation, modAppDef.APP_NAME
    End If

Done:
    On Error Resume Next
    Application.Cursor = -4143
    Application.StatusBar = False
    On Error GoTo 0
    Exit Sub

Failed:
    Dim desc As String: desc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailedCleanup1
FailedCleanup1:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Cursor = -4143
    Application.StatusBar = False
    modLog.LogError "E0801", "modMigrate.ImportUserData", desc
    MsgBox "引き継ぎファイルを読み込めませんでした。" & vbCrLf & "(" & desc & ")", _
           vbExclamation, modAppDef.APP_NAME
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 内部
' ----------------------------------------------------------------------------
Private Sub WriteMeta(ByVal ws As Worksheet)
    ws.Cells(1, 1).Value = "key"
    ws.Cells(1, 2).Value = "value"
    ws.Cells(2, 1).Value = "mig_format_version"
    ws.Cells(2, 2).Value = MIG_FORMAT_VERSION
    ws.Cells(3, 1).Value = "app_version"
    ws.Cells(3, 2).Value = modAppDef.APP_VERSION
    ws.Cells(4, 1).Value = "created_at"
    ws.Cells(4, 2).Value = modUtil.NowStamp()
    ws.Cells(5, 1).Value = "embed_dim"
    ws.Cells(5, 2).Value = modConfig.GetLong("embed_dim", 1536)
End Sub

Private Function ValidateMigFile(ByVal wb As Workbook, ByRef reason As String) As Boolean
    If wb Is Nothing Then
        reason = "ブックを読み込めませんでした"
        Exit Function
    End If
    If Not SheetExistsIn(wb, MIG_META) Then
        reason = "引き継ぎ情報(" & MIG_META & ")が見つかりません"
        Exit Function
    End If
    If Not SheetExistsIn(wb, modAppDef.SH_KNOWLEDGE) Then
        reason = "本棚の中身が入っていません"
        Exit Function
    End If

    ' 2026-08-01(R12-1-7): この版より新しい形式の引き継ぎファイルは断る。
    ' mig_format_version を書いておきながら誰も読んでいなかったため、将来の版で
    ' 列や意味を足したファイルを旧版が黙って部分的に取り込み、「移行したのに
    ' 一部だけ欠けている」という気付けない壊れ方をする(前方互換の穴)。
    ' 旧い形式(小さい番号)は従来どおり受け入れる。
    Dim srcFmt As Long
    srcFmt = CLng(Val(ReadMetaValue(wb.Worksheets(MIG_META), "mig_format_version")))
    If srcFmt > MIG_FORMAT_VERSION Then
        reason = "この引き継ぎファイルは、より新しい版のマイ本棚で作られています" & _
                 "(ファイルの形式=" & srcFmt & " / この版が読めるのは " & _
                 MIG_FORMAT_VERSION & " まで)。" & vbCrLf & _
                 "新しい版のマイ本棚を開いて、そちらで取り込んでください"
        Exit Function
    End If

    ' ベクトルの次元が違うブックへ移すと、検索が全件で次元不一致になる。
    ' 取り込んでから気付くと本棚が使えないので、入口で止める。
    Dim srcDim As Long: srcDim = CLng(Val(ReadMetaValue(wb.Worksheets(MIG_META), "embed_dim")))
    Dim myDim As Long: myDim = modConfig.GetLong("embed_dim", 1536)
    If srcDim > 0 And srcDim <> myDim Then
        reason = "ベクトルの次元数が一致しません(引き継ぎ元=" & srcDim & " / この版=" & myDim & ")"
        Exit Function
    End If
    ValidateMigFile = True
End Function

Private Function ReadMetaValue(ByVal ws As Worksheet, ByVal key As String) As String
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim i As Long
    For i = 1 To lastR
        If StrComp(CStr(ws.Cells(i, 1).Value), key, vbTextCompare) = 0 Then
            ReadMetaValue = CStr(ws.Cells(i, 2).Value)
            Exit Function
        End If
    Next i
End Function

Private Function SheetExistsIn(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExistsIn = Not (ws Is Nothing)
End Function

' このブックのシートを、引き継ぎブックへ値だけコピーする。
Private Sub CopySheetInto(ByVal destWb As Workbook, ByVal sheetName As String)
    On Error Resume Next
    Dim src As Worksheet
    Set src = ThisWorkbook.Worksheets(sheetName)
    If src Is Nothing Then Exit Sub

    Dim dst As Worksheet
    Set dst = destWb.Worksheets.Add(After:=destWb.Worksheets(destWb.Worksheets.count))
    dst.Name = sheetName

    Dim lastR As Long: lastR = src.Cells(src.Rows.count, 1).End(xlUp).row
    Dim lastC As Long: lastC = src.Cells(1, src.Columns.count).End(xlToLeft).Column
    If lastR < 1 Or lastC < 1 Then Exit Sub

    ' 数式ではなく値を運ぶ(移行先で参照が壊れないように)。
    ' full_text は最大32,000字あるので、一括代入は200行ずつに割る。
    Dim rowStart As Long
    For rowStart = 1 To lastR Step 200
        Dim rowsN As Long
        rowsN = lastR - rowStart + 1
        If rowsN > 200 Then rowsN = 200
        dst.Range(dst.Cells(rowStart, 1), dst.Cells(rowStart + rowsN - 1, lastC)).Value = _
            src.Range(src.Cells(rowStart, 1), src.Cells(rowStart + rowsN - 1, lastC)).Value
    Next rowStart
    On Error GoTo 0
End Sub

' config は「引き継いでよいキーだけ」を別シート名で持ち出す。
Private Sub CopyConfigInto(ByVal destWb As Workbook)
    On Error Resume Next
    Dim src As Worksheet
    Set src = ThisWorkbook.Worksheets(modAppDef.SH_CONFIG)
    If src Is Nothing Then Exit Sub

    Dim dst As Worksheet
    Set dst = destWb.Worksheets.Add(After:=destWb.Worksheets(destWb.Worksheets.count))
    dst.Name = "config_user"
    dst.Cells(1, 1).Value = "key"
    dst.Cells(1, 2).Value = "value"

    Dim lastR As Long: lastR = src.Cells(src.Rows.count, 1).End(xlUp).row
    Dim outR As Long: outR = 2
    Dim i As Long
    For i = 2 To lastR
        Dim k As String: k = Trim$(CStr(src.Cells(i, 1).Value))
        If LenB(k) > 0 Then
            If Not IsBuildOwnedKey(k) Then
                dst.Cells(outR, 1).Value = k
                dst.Cells(outR, 2).Value = src.Cells(i, 2).Value
                outR = outR + 1
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

' insight_inbox は「未取込(consumed<>1)の行だけ」を別途持ち出す(2026-08-01
' R12-5-12)。取込済み行は本棚へ既に反映済みで運ぶ理由が無く、
' modInsightIo.TrimConsumedRows(R12-3-5)と同じ「取込前の共有知は消さない」
' 判断に合わせる。列構成は insight_inbox 実シートと同じ(先頭行=ヘッダー)。
Private Sub CopyInsightInboxInto(ByVal destWb As Workbook)
    On Error Resume Next
    Dim src As Worksheet: Set src = ThisWorkbook.Worksheets("insight_inbox")
    If src Is Nothing Then Exit Sub

    Dim lastR As Long: lastR = src.Cells(src.Rows.count, 1).End(xlUp).row
    Dim lastC As Long: lastC = src.Cells(1, src.Columns.count).End(xlToLeft).Column
    If lastC < 1 Then lastC = 10   ' ヘッダーすら読めない異常系の保険(実シート想定は10列)

    Dim dst As Worksheet
    Set dst = destWb.Worksheets.Add(After:=destWb.Worksheets(destWb.Worksheets.count))
    dst.Name = "insight_inbox"
    dst.Range(dst.Cells(1, 1), dst.Cells(1, lastC)).Value = _
        src.Range(src.Cells(1, 1), src.Cells(1, lastC)).Value
    If lastR < 2 Then Exit Sub

    ' 列I(9列目)=consumed。一括読み→フィルタ→書出しで行数に依らず軽く保つ
    ' (TrimConsumedRowsと同じ作法。受信箱は最大500行=R12-3-5のため実質軽い)。
    Dim arr As Variant: arr = src.Range(src.Cells(2, 1), src.Cells(lastR, lastC)).Value
    Dim outR As Long: outR = 2
    Dim i As Long, c As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If CStr(arr(i, 9)) <> "1" Then
            For c = 1 To lastC
                dst.Cells(outR, c).Value = InboxCellValue(arr(i, c), c)
            Next c
            outR = outR + 1
        End If
    Next i
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' InboxCellValue - insight_inbox の1セルを書く直前に通す(R12-H-4)。
' ----------------------------------------------------------------------------
' 受信箱は【他人が書いた文章】が入る唯一のシートで、引き継ぎファイルは
' 端末をまたいで運ばれる。取込経路(modInsightIo)は SanitizeForCell 済みでも、
' 引き継ぎ経路だけが素通りだと、そこが数式インジェクションの抜け道になる
' (R12-2 で塞いだ3経路と同型の穴)。
' 対象は未信頼テキスト列だけに限る: 4=author / 6=question / 7=answer_or_reason
' / 8=source_or_dept(ビルド側 text_cols と同じ集合)。nonce・created_at・
' consumed・selected のような機械が作る列に "'" を足すと、突合や数値判定が
' 壊れるので触らない。
Private Function InboxCellValue(ByVal v As Variant, ByVal col As Long) As Variant
    Select Case col
        Case 4, 6, 7, 8
            InboxCellValue = modUtilText.SanitizeForCell(CStr(v))
        Case Else
            InboxCellValue = v
    End Select
End Function

' ビルドが決める値・秘密は引き継がない(§12.3 B1/B2 の再発防止)。
Private Function IsBuildOwnedKey(ByVal k As String) As Boolean
    If InStr(1, CFG_SKIP_EXACT, "|" & LCase$(k) & "|", vbTextCompare) > 0 Then
        IsBuildOwnedKey = True
        Exit Function
    End If
    IsBuildOwnedKey = (StrComp(Left$(LCase$(k), Len(CFG_SKIP_PREFIX)), CFG_SKIP_PREFIX, vbTextCompare) = 0)
End Function

' 利用者が決めた値として引き継ぐキーか(R12-3-2。判断根拠は CFG_KEEP_KEYS の上)。
' 書き出し側(CopyConfigInto)は従来どおり広めに持ち出す。読み込み側で絞るのは、
' 旧い引き継ぎファイルでも新しい判断基準がそのまま効くようにするため。
Private Function IsUserOwnedKey(ByVal k As String) As Boolean
    IsUserOwnedKey = (InStr(1, CFG_KEEP_KEYS, "|" & LCase$(Trim$(k)) & "|", vbTextCompare) > 0)
End Function

' 引き継ぎブックのシートで、このブックのシートを置き換える。
'
' 2026-08-01(R12-3-1): 戻り値 Boolean + 実測行数へ変えた。
' 従来は関数全体が On Error Resume Next で、書き戻しの途中で失敗しても
' 「引き継ぎました」という完了ダイアログだけが出た(憲章§4-1 無言の失敗禁止)。
' 本棚が半分しか戻っていないのに成功と言うのは、利用者が最も気付けない形の
' データ欠損である。失敗は E0801 で記録し、呼び出し側が文言を切り替える。
'   outRows : 復元後に dst 側で【実際に】数えたデータ行数(見出し行を除く)。
'             書けたつもりの件数ではなく、書けた結果を数える。
Private Function RestoreSheet(ByVal srcWb As Workbook, ByVal sheetName As String, _
                              ByRef outRows As Long) As Boolean
    outRows = 0
    On Error GoTo Fail

    ' 引き継ぎファイル側にそのシートが無いのは失敗ではない(旧版の
    ' 引き継ぎファイルには存在しないシートがあり得る)。何も置き換えない。
    If Not SheetExistsIn(srcWb, sheetName) Then
        RestoreSheet = True
        Exit Function
    End If

    Dim src As Worksheet: Set src = srcWb.Worksheets(sheetName)
    ' 置き換え先が無い(=このブックが壊れている)場合は Worksheets(...) が
    ' 実行時エラー9を出し、下の Fail が記録する。黙って諦めない。
    Dim dst As Worksheet: Set dst = ThisWorkbook.Worksheets(sheetName)

    Dim lastR As Long: lastR = src.Cells(src.Rows.count, 1).End(xlUp).row
    Dim lastC As Long: lastC = src.Cells(1, src.Columns.count).End(xlToLeft).Column
    If lastC < 1 Then
        RestoreSheet = True
        Exit Function
    End If

    ' 既存を消してから入れる(混ぜない)。見出し行は入れ直すので全消しでよい。
    dst.Cells.ClearContents
    If lastR < 1 Then
        RestoreSheet = True
        Exit Function
    End If

    Dim rowStart As Long
    For rowStart = 1 To lastR Step 200
        Dim rowsN As Long
        rowsN = lastR - rowStart + 1
        If rowsN > 200 Then rowsN = 200
        dst.Range(dst.Cells(rowStart, 1), dst.Cells(rowStart + rowsN - 1, lastC)).Value = _
            src.Range(src.Cells(rowStart, 1), src.Cells(rowStart + rowsN - 1, lastC)).Value
    Next rowStart

    outRows = dst.Cells(dst.Rows.count, 1).End(xlUp).row - 1
    If outRows < 0 Then outRows = 0
    ' R12-H-6: 引き継ぎはシートを丸ごと入れ替える。my_vectors を入れ替えたら
    ' セッション内キャッシュは中身ごと別物なので、解放して世代を進める
    ' (件数も先頭/末尾idも偶然一致し得るため、印だけでは足りない)。
    If StrComp(sheetName, modAppDef.SH_VECTORS, vbTextCompare) = 0 Then
        On Error Resume Next
        modVecCache.ResetVecCache
        modVecCache.BumpGeneration
        On Error GoTo Fail
    End If
    RestoreSheet = True
    Exit Function

Fail:
    Dim failNum As Long: failNum = Err.Number
    Dim failDesc As String: failDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かない。記録の前に Resume で
    ' ハンドラを抜ける(modShelf.IngestFile と同じ作法。2026-07-30 実機err#462)。
    Resume FailCleanup
FailCleanup:
    On Error Resume Next
    modLog.LogError "E0801", "modMigrate.ImportUserData", _
        "シート復元に失敗: " & sheetName & " err#" & failNum & ": " & failDesc
    On Error GoTo 0
End Function

' my_stats はキー単位で併合する(新版が先に書いた行を消さない)。
Private Sub MergeStats(ByVal srcWb As Workbook)
    On Error Resume Next
    If Not SheetExistsIn(srcWb, modAppDef.SH_STATS) Then Exit Sub
    Dim src As Worksheet: Set src = srcWb.Worksheets(modAppDef.SH_STATS)
    Dim lastR As Long: lastR = src.Cells(src.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then Exit Sub

    Dim arr As Variant
    arr = src.Range(src.Cells(2, 1), src.Cells(lastR, 2)).Value
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim k As String: k = Trim$(CStr(arr(i, 1)))
        If LenB(k) > 0 Then modStats.SetStatText k, CStr(arr(i, 2))
    Next i
    On Error GoTo 0
End Sub

' config は「この版に存在するキーだけ」値を引き継ぐ。
' 新版で増えたキーは新版の既定のまま(古い設定で新機能を殺さない)。
Private Sub MergeConfig(ByVal srcWb As Workbook)
    On Error Resume Next
    If Not SheetExistsIn(srcWb, "config_user") Then Exit Sub
    Dim src As Worksheet: Set src = srcWb.Worksheets("config_user")
    Dim lastR As Long: lastR = src.Cells(src.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then Exit Sub

    Dim me_ As Worksheet: Set me_ = ThisWorkbook.Worksheets(modAppDef.SH_CONFIG)
    If me_ Is Nothing Then Exit Sub
    Dim myLast As Long: myLast = me_.Cells(me_.Rows.count, 1).End(xlUp).row
    If myLast < 2 Then Exit Sub

    ' この版のキー集合(存在確認用)。
    Dim known As Object: Set known = CreateObject("Scripting.Dictionary")
    Dim myArr As Variant: myArr = me_.Range(me_.Cells(2, 1), me_.Cells(myLast, 1)).Value
    Dim j As Long
    For j = LBound(myArr, 1) To UBound(myArr, 1)
        Dim mk As String: mk = LCase$(Trim$(CStr(myArr(j, 1))))
        If LenB(mk) > 0 Then
            If Not known.Exists(mk) Then known.Add mk, True
        End If
    Next j

    Dim arr As Variant: arr = src.Range(src.Cells(2, 1), src.Cells(lastR, 2)).Value
    Dim i As Long, applied As Long, skipped As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim k As String: k = Trim$(CStr(arr(i, 1)))
        If LenB(k) > 0 Then
            ' R12-3-2: ホワイトリスト方式。ビルドが決める値の除外(従来)に加え、
            ' 「利用者が決めた値」だけを通す。新版の既定値を旧値で巻き戻さない。
            If IsUserOwnedKey(k) And Not IsBuildOwnedKey(k) Then
                If known.Exists(LCase$(k)) Then
                    modConfig.SetValue k, CStr(arr(i, 2))
                    applied = applied + 1
                End If
            Else
                skipped = skipped + 1
            End If
        End If
    Next i
    modLog.LogUsage "migrate_config", "", "applied=" & applied & " kept_new_default=" & skipped
    On Error GoTo 0
End Sub

' insight_inbox は書き出し側(CopyInsightInboxInto)と対称に「未取込の行を
' 追記」する(2026-08-01 R12-5-12)。置換ではない: 移行先ブックが起動後に
' 自分で受信した共有知を消さないため、既存の受信箱に合流させる。
' nonce(A列)が既に存在する行は二重取り込みしない。
Private Sub MergeInsightInbox(ByVal srcWb As Workbook)
    On Error Resume Next
    If Not SheetExistsIn(srcWb, "insight_inbox") Then Exit Sub
    Dim src As Worksheet: Set src = srcWb.Worksheets("insight_inbox")
    Dim srcLastR As Long: srcLastR = src.Cells(src.Rows.count, 1).End(xlUp).row
    If srcLastR < 2 Then Exit Sub

    Dim dst As Worksheet: Set dst = ThisWorkbook.Worksheets("insight_inbox")
    If dst Is Nothing Then Exit Sub
    Dim dstLastR As Long: dstLastR = dst.Cells(dst.Rows.count, 1).End(xlUp).row

    Dim lastC As Long: lastC = src.Cells(1, src.Columns.count).End(xlToLeft).Column
    If lastC < 1 Then lastC = 10

    ' 既存nonce(A列)の集合(重複取り込み防止)。
    Dim known As Object: Set known = CreateObject("Scripting.Dictionary")
    If dstLastR >= 2 Then
        Dim dstArr As Variant: dstArr = dst.Range(dst.Cells(2, 1), dst.Cells(dstLastR, 1)).Value
        Dim j As Long
        For j = LBound(dstArr, 1) To UBound(dstArr, 1)
            Dim dn As String: dn = CStr(dstArr(j, 1))
            If LenB(dn) > 0 Then
                If Not known.Exists(dn) Then known.Add dn, True
            End If
        Next j
    End If

    Dim srcArr As Variant: srcArr = src.Range(src.Cells(2, 1), src.Cells(srcLastR, lastC)).Value
    Dim outR As Long: outR = dstLastR + 1
    Dim added As Long, dup As Long, skipped As Long
    Dim i As Long, c As Long
    For i = LBound(srcArr, 1) To UBound(srcArr, 1)
        Dim nc As String: nc = CStr(srcArr(i, 1))
        If LenB(nc) > 0 And CStr(srcArr(i, 9)) <> "1" Then   ' consumed(念のため防御的に再確認)
            If known.Exists(nc) Then
                dup = dup + 1
            Else
                For c = 1 To lastC
                    dst.Cells(outR, c).Value = InboxCellValue(srcArr(i, c), c)   ' R12-H-4
                Next c
                known.Add nc, True
                outR = outR + 1
                added = added + 1
            End If
        Else
            skipped = skipped + 1
        End If
    Next i
    modLog.LogUsage "migrate_insight", "", "added=" & added & " dup_skip=" & dup & _
        " consumed_skip=" & skipped
    On Error GoTo 0
End Sub
