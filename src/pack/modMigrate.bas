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

' 引き継がない config キー(ビルドが決める値・秘密)。
' 前方一致で判定するので azure_embed_key / azure_embed_url は "azure_" で拾う。
Private Const CFG_SKIP_EXACT As String = "|mock_llm|build_stamp|publish_key|"
Private Const CFG_SKIP_PREFIX As String = "azure_"

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

    Dim prevScreen As Boolean: prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim wb As Workbook
    Set wb = Application.Workbooks.Open(Filename:=CStr(fd.SelectedItems(1)), _
                                        ReadOnly:=True, UpdateLinks:=0)

    Dim reason As String
    If Not ValidateMigFile(wb, reason) Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
        Application.ScreenUpdating = prevScreen
        Application.Cursor = -4143
        Application.StatusBar = False
        MsgBox "このファイルは引き継ぎファイルではないようです。" & vbCrLf & _
               "(" & reason & ")", vbExclamation, modAppDef.APP_NAME
        GoTo Done
    End If

    RestoreSheet wb, modAppDef.SH_KNOWLEDGE
    RestoreSheet wb, modAppDef.SH_VECTORS
    RestoreSheet wb, modAppDef.SH_MANIFEST
    MergeStats wb
    MergeConfig wb

    wb.Close SaveChanges:=False
    Set wb = Nothing
    Application.ScreenUpdating = prevScreen
    Application.Cursor = -4143
    Application.StatusBar = False

    Dim got As Long
    On Error Resume Next
    got = modShelf.TotalChunks()
    modLog.LogUsage "migrate_import", "", "chunks=" & got
    On Error GoTo Failed

    MsgBox "引き継ぎました。" & vbCrLf & vbCrLf & _
           "本棚: " & got & " 件" & vbCrLf & vbCrLf & _
           "設定は、この版に存在する項目だけを引き継いでいます" & vbCrLf & _
           "(新しく増えた設定は初期値のままです)。" & vbCrLf & _
           "画面を描き直すため、一度閉じて開き直してください。", _
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
    Resume FailedCleanup1
FailedCleanup1:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
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

' ビルドが決める値・秘密は引き継がない(§12.3 B1/B2 の再発防止)。
Private Function IsBuildOwnedKey(ByVal k As String) As Boolean
    If InStr(1, CFG_SKIP_EXACT, "|" & LCase$(k) & "|", vbTextCompare) > 0 Then
        IsBuildOwnedKey = True
        Exit Function
    End If
    IsBuildOwnedKey = (StrComp(Left$(LCase$(k), Len(CFG_SKIP_PREFIX)), CFG_SKIP_PREFIX, vbTextCompare) = 0)
End Function

' 引き継ぎブックのシートで、このブックのシートを置き換える。
Private Sub RestoreSheet(ByVal srcWb As Workbook, ByVal sheetName As String)
    On Error Resume Next
    If Not SheetExistsIn(srcWb, sheetName) Then Exit Sub
    Dim src As Worksheet: Set src = srcWb.Worksheets(sheetName)
    Dim dst As Worksheet: Set dst = ThisWorkbook.Worksheets(sheetName)
    If dst Is Nothing Then Exit Sub

    Dim lastR As Long: lastR = src.Cells(src.Rows.count, 1).End(xlUp).row
    Dim lastC As Long: lastC = src.Cells(1, src.Columns.count).End(xlToLeft).Column
    If lastC < 1 Then Exit Sub

    ' 既存を消してから入れる(混ぜない)。見出し行は入れ直すので全消しでよい。
    dst.Cells.ClearContents
    If lastR < 1 Then Exit Sub

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
    Dim i As Long, applied As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim k As String: k = Trim$(CStr(arr(i, 1)))
        If LenB(k) > 0 Then
            If Not IsBuildOwnedKey(k) Then
                If known.Exists(LCase$(k)) Then
                    modConfig.SetValue k, CStr(arr(i, 2))
                    applied = applied + 1
                End If
            End If
        End If
    Next i
    modLog.LogUsage "migrate_config", "", "applied=" & applied
    On Error GoTo 0
End Sub
