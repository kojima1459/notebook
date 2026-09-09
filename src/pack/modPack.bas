Attribute VB_Name = "modPack"
Option Explicit

' ============================================================================
' modPack - ナレッジパックの書き出し/取込(MASTER_SPEC §7.4)
' ----------------------------------------------------------------------------
' 本棚(全体 or 1資料)を新規.xlsx(マクロ無し)へ pack_meta/pack_chunks/
' pack_vectors の3シートで書き出し、取込側は chunk_id の fnv ハッシュ部で
' 重複排除しながら my_knowledge/my_vectors へ追記する(origin="pack:"&作成者)。
'
' 設計判断:
'   ・検証は2段(§7.4): ValidatePack(wb,reason)がWorkbook実物(シート存在+
'     embed_dim一致)を見て、内部で純関数ValidatePackMeta(Long/Long/String
'     のみ=テスト可能)を呼ぶ。自分のembed_dimとの一致はValidatePack側。
'   ・ダイアログ版とパス指定版を分けてある。modPackExport.ExportPackToFile /
'     ImportPackFile はダイアログを出さずに完了でき、部門正典の発行と
'     チャンネル同期がこちらを使う。PII走査だけはどちらの経路でも必ず通す。
' ============================================================================

Private Const COL_ID As Long = 1
Private Const COL_SOURCE As Long = 2
Private Const COL_ORIGIN As Long = 3
Private Const COL_PAGE As Long = 4
Private Const COL_SUMMARY As Long = 5
Private Const COL_KEYWORDS As Long = 6
Private Const COL_FULLTEXT As Long = 7
Private Const COL_ADDED As Long = 8
Private Const COL_EMBEDDED As Long = 9

' ImportPackDialog - ファイル選択→ValidatePack→重複(fnvハッシュ)スキップ
'   しつつmy_knowledge/my_vectorsへ取込。origin="pack:"&作成者。
Public Sub ImportPackDialog()
    Dim fd As Object
    Set fd = Application.FileDialog(3)   ' msoFileDialogFilePicker(名前付き定数は使わない)
    fd.AllowMultiSelect = False
    fd.Title = "取り込むパックファイルを選んでください"
    fd.Filters.Clear
    fd.Filters.Add "Excelブック", "*.xlsx"

    If fd.Show <> -1 Then Exit Sub   ' キャンセル
    If fd.SelectedItems.count < 1 Then Exit Sub
    ImportPackFile CStr(fd.SelectedItems(1)), False
End Sub

' ImportPackFile - パスを直接指定してパックを取り込む(部門チャンネルの
'   自動同期から使う)。silent:=True のときは完了/失敗のダイアログを出さず、
'   取り込み件数だけを返す(起動時の同期でダイアログが出ると業務が止まる)。
'   ダイアログ経由の従来動作は ImportPackDialog がここを silent:=False で呼ぶ。
'
' originOverride (2026-07-28 レビュー C-1):
'   取り込んだ行の origin 列に書く値。空なら従来どおり "pack:"&作者名。
'   部門チャンネルの同期は "channel:"&部門名 を渡す。
'   従来はチャンネル経由でも "pack:"&作者名 が書かれる一方、消す側は
'   "pack:"&部門名 を探していたため、部門名≠作者名である限り削除は常に
'   0件だった(=切替しても前の部門が残り、更新配信で新旧が混ざる)。
'   手渡しパックと部門正典で名前空間を分け、二度と衝突させない。
'
' purgeOrigins (2026-07-28 レビュー H-14):
'   書き込む直前に消す origin タグ("|"区切りで複数可)。
'   「消してから読み込む」と、読み込みに失敗したときに何も残らない。
'   ここまで来た時点でパックの検証もチャンクの読み出しも済んでいるので、
'   消してから書くまでの間に失敗する余地がほぼ無い。
Public Function ImportPackFile(ByVal packPath As String, ByVal silent As Boolean, _
                               Optional ByVal originOverride As String = "", _
                               Optional ByVal purgeOrigins As String = "") As Long
    ' 2026-07-28(レビュー M-10): Open から Close までを必ずハンドラで覆う。
    ' 従来はこの区間が無ハンドラで、pack 内のエラー値セル等で素のVBAエラーが
    ' 出ると【pack.xlsx が開きっぱなし】のまま抜けていた。共有フォルダの
    ' ファイルを掴んだままになるので、発行者が次の版を書けなくなる。
    ' 画面のちらつき(ScreenUpdating)もここで抑える。
    Dim wb As Workbook
    Dim prevScreen As Boolean: prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False

    ' 2026-09-04(R35 F3a・実機第20報): なぜ Password 引数を外したか。
    ' この端末は Workbooks.Open に Password 引数(名前付き引数)を付けると、保護されていない
    ' 普通のファイルでも Open 自体が失敗する(2026-07-29型の端末と同型。
    ' modExtractorExcel が既に踏んで「引数なしで再試行」の保険を持つ壁と同じ)。
    ' ダミーPassword(旧R27波3-3)で「入力を求めない」を実現していたが、
    ' それでは保護ファイルもこの端末型の普通ファイルも同じ Open失敗にしか
    ' ならない。そこで Open の【前】に先頭4バイトで暗号化かどうかを判別し
    ' (modShelfScan.EncryptedFileKind→OpenGateReason)、暗号化/判別不能と
    ' 分かったものだけを開かずに E0701 で倒す。"plain"と分かったものは
    ' Password/WriteResPassword を渡さずに開く(部門チャンネルの自動同期
    ' (起動直後・無人)からも通る経路なので、開いた先で読み取り推奨/
    ' リンク更新のモーダルが出ないよう IgnoreReadOnlyRecommended/UpdateLinks:=0
    ' は維持する)。
    ' modUiLock.AlertsOff/On は UI層のため機能層からは呼べない(R1)。
    ' 同層 modMigrate.ExportUserData(:110-132)と同じ素の退避/復元にする。
    Dim prevAlerts As Boolean: prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    Dim kind As String: kind = modShelfScan.EncryptedFileKind(packPath)
    Dim gate As String: gate = OpenGateReason(kind)
    If gate <> "open" Then
        Application.DisplayAlerts = prevAlerts
        Application.ScreenUpdating = prevScreen
        Dim gateMsg As String
        If gate = "enc" Then
            gateMsg = "このファイルはパスワードで保護されています。保護を外したものを受け取ってください。"
        Else
            gateMsg = "ファイルを読めません。他の人が書き込み中か、Excel ブックではない可能性があります。"
        End If
        modLog.LogError "E0701", "modPack.ImportPackFile", _
            "開く前に判別して中止(kind=[" & kind & "]): " & packPath
        If Not silent Then
            MsgBox gateMsg & vbLf & "(コード: E0701)", vbExclamation, modAppDef.APP_NAME
        End If
        Exit Function
    End If

    On Error GoTo OpenFail
    Set wb = Application.Workbooks.Open( _
        Filename:=packPath, ReadOnly:=True, UpdateLinks:=0, _
        IgnoreReadOnlyRecommended:=True, AddToMru:=False)
    Application.DisplayAlerts = prevAlerts
    On Error GoTo LoadFailed

    Dim reason As String
    If Not ValidatePack(wb, reason) Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0

        Dim badCode As String
        If InStr(reason, "次元") > 0 Then
            badCode = "E0702"
        Else
            badCode = "E0701"
        End If
        modLog.LogError badCode, "modPack.ImportPackFile", reason
        Application.ScreenUpdating = prevScreen
        If Not silent Then
            MsgBox modLog.FriendlyMessage(badCode) & vbLf & "(コード: " & badCode & ")" & vbLf & _
                "詳細: " & reason, vbExclamation, modAppDef.APP_NAME
        End If
        Exit Function
    End If

    Dim authorName As String
    authorName = Trim$(ReadMetaValue(wb.Worksheets("pack_meta"), "author"))
    If LenB(authorName) = 0 Then authorName = "不明"

    ' 2026-07-28(レビュー H-5): 感謝状の宛先に使う作者のID(ADのCN等)。
    ' チャンクの origin には従来どおり表示名を書く(画面に出るのは名前)。
    ' 表示名→IDの対応だけ my_stats に控えておき、modP2P が宛先解決に使う。
    ' 旧いパック(author_id 無し)を取り込んだときは何も記録しないので、
    ' modP2P 側は従来どおり表示名で宛先を作る(過渡期の互換)。
    Dim authorId As String
    authorId = Trim$(ReadMetaValue(wb.Worksheets("pack_meta"), "author_id"))
    If LenB(authorId) > 0 Then
        On Error Resume Next
        modStats.SetStatText "pkauth:" & LCase$(authorName), authorId
        On Error GoTo LoadFailed
    End If

    ' 2026-07-31(レビュー R8 F1): 部門チャンネル経由の取込では、チャンクの
    ' origin が "channel:<部門名>" になる(originOverride)。部門名は人では
    ' ないため、pkauth(表示名→ID)だけでは感謝状の宛先を解決できず、
    ' チャンネル経由の感謝は【構造的に一度も発行されていなかった】。
    ' pkauth と同じ作法で "chauth:<部門名>" → 発行者ID を控える。
    ' キーの綴りは modShareRule.AuthorStatKey が唯一の決定者(読む側の
    ' modP2P.ResolveAuthorId も同じ関数を通す)。
    ' 発行者IDを持たない旧い正典パックでは何も記録しない = 従来どおり感謝なし。
    If LenB(authorId) > 0 Then
        Dim chKey As String: chKey = modShareRule.AuthorStatKey(originOverride)
        If LenB(chKey) > 0 Then
            On Error Resume Next
            modStats.SetStatText chKey, authorId
            On Error GoTo LoadFailed
        End If
    End If

    Dim ids() As String, sources() As String, pages() As Long
    Dim summaries() As String, keywords() As String, fullTexts() As String, vectors() As String
    Dim n As Long
    n = LoadPackChunksAndVectors(wb, ids, sources, pages, summaries, keywords, fullTexts, vectors)

    On Error Resume Next
    wb.Close SaveChanges:=False
    Set wb = Nothing
    On Error GoTo 0
    Application.ScreenUpdating = prevScreen

    If n = 0 Then
        If Not silent Then
            MsgBox "パックの中に取込める資料がありませんでした。" & vbLf & _
                "別のパックファイルでお試しください。", vbExclamation, modAppDef.APP_NAME
        End If
        Exit Function
    End If

    ' ここまででパックの検証もチャンクの読み出しも終わり、ブックも閉じてある。
    ' 「旧版を消す」のはこの位置。これより前で消すと、読み込みに失敗した時に
    ' 旧版も新版も無いという最悪の状態が残る(レビュー H-14)。
    ' R33H M3: 削除の成否を必ず受け取る。F2 は失敗の形を「消し損ね(重複)」から
    ' 【欠落】へ反転させたのに、最頻経路(📡部門チャンネル→SyncChannel→ここ)
    ' だけが outOk を受け取らず、他部門や自分で登録した資料まで消えたあとで
    ' 「N件取込」の成功ダイアログを出していた。self(自分で登録した資料)には
    ' 再取込の導線が1つも無く「取り返しがつく側に倒した」という前提が
    ' 成立しないので、1つでも失敗したらこの取込は失敗として返す。
    Dim purgedCount As Long
    Dim purgeOk As Boolean: purgeOk = True
    If LenB(purgeOrigins) > 0 Then
        Dim tags() As String: tags = Split(purgeOrigins, "|")
        Dim ti As Long
        Dim rmOk As Boolean
        For ti = LBound(tags) To UBound(tags)
            If LenB(Trim$(tags(ti))) > 0 Then
                rmOk = True
                purgedCount = purgedCount + modShelfStore.RemoveRowsByOrigin(Trim$(tags(ti)), rmOk)
                If Not rmOk Then purgeOk = False
            End If
        Next ti
    End If

    Dim importedCount As Long, skippedCount As Long
    ImportChunksDedup ids, sources, pages, summaries, keywords, fullTexts, vectors, n, authorName, _
        originOverride, importedCount, skippedCount

    modStats.Bump "pack_import_total"
    modLog.LogUsage "pack_import", "", "imported=" & importedCount & " skipped=" & skippedCount & _
        " purged=" & purgedCount & " author=" & authorName
    ' R33H M3: 旧版の削除に失敗していたら、取り込めた件数にかかわらず失敗を
    ' 返す(0件=呼び出し元 modChannel.SyncChannel の失敗側。版数を記録しないので
    ' 次に押したときに取り込み直せる)。取り込み自体は上で済ませてある ――
    ' 消えた行を1件でも埋め戻せる方が、何もしないより取り返しがつく。
    If Not purgeOk Then
        modLog.LogError "E0701", "modPack.ImportPackFile", _
            "旧版の削除に失敗(欠落の可能性): imported=" & importedCount & _
            " purged=" & purgedCount & " origins=" & purgeOrigins
        If Not silent Then
            MsgBox "本棚の書き換えに失敗しました。" & vbCrLf & vbCrLf & _
                "ほかの部門や、ご自分で登録した資料まで消えている可能性があります。" & vbCrLf & _
                "マイ本棚を開いてご確認ください(このツールの管理担当にもご連絡ください)。", _
                vbExclamation, modAppDef.APP_NAME
        End If
        Exit Function
    End If

    ' R40 F4(実機報告): 「はじめてのパック取込」の獲得告知を取込直後に出す
    ' (従来は次の無関係な操作まで遅れて出ていた。取得済みなら即Exit)。
    ' 旧版削除の失敗(上の E0701)より後=失敗の直前に祝わない(レビュー R40 m)。
    On Error Resume Next
    modStats.EvaluateBadges
    On Error GoTo 0

    ImportPackFile = importedCount
    If Not silent Then
        MsgBox importedCount & "件取込 / " & skippedCount & "件は既にありました。", _
            vbInformation, modAppDef.APP_NAME
    End If
    Exit Function

LoadFailed:
    ' 検証・読み出しの途中で落ちた。何より先にブックを閉じる
    ' (掴んだままにすると発行者が次の版を書けなくなる)。
    Dim loadDesc As String: loadDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume LoadFailedCleanup1
LoadFailedCleanup1:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set wb = Nothing
    On Error GoTo 0
    Application.ScreenUpdating = prevScreen
    modLog.LogError "E0701", "modPack.ImportPackFile", "パックの読み出しに失敗: " & loadDesc
    If Not silent Then
        MsgBox "パックファイルを読み込めませんでした。" & vbLf & _
            "壊れたファイルでないか確認してください。" & vbLf & "(コード: E0701)", _
            vbExclamation, modAppDef.APP_NAME
    End If
    Exit Function

OpenFail:
    ' Err はハンドラを解いた時点で消える。後始末より先に退避する
    ' (modMigrate/modHub と同じ作法。従来は空文字を記録していた)。
    Dim openDesc As String: openDesc = Err.Description
    On Error GoTo 0
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    modLog.LogError "E0701", "modPack.ImportPackFile", "ファイルを開けません: " & openDesc
    If Not silent Then
        MsgBox "パックファイルを開けませんでした。" & vbLf & _
            "壊れたファイルでないか、パスに間違いがないか確認してください。" & vbLf & "(コード: E0701)", _
            vbExclamation, modAppDef.APP_NAME
    End If
End Function

' ValidatePack - Workbook実物を検査(3シート存在・pack_format_version一致・
'   embed_dim一致)。E0701/E0702の判定材料をreasonに返す。
Public Function ValidatePack(ByVal wb As Workbook, ByRef reason As String) As Boolean
    reason = ""
    If wb Is Nothing Then
        reason = "ブックを読み込めませんでした"
        Exit Function
    End If

    If Not SheetExistsIn(wb, "pack_meta") Or Not SheetExistsIn(wb, "pack_chunks") _
            Or Not SheetExistsIn(wb, "pack_vectors") Then
        reason = "必要なシート(pack_meta/pack_chunks/pack_vectors)が見つかりません"
        Exit Function
    End If

    Dim wsMeta As Worksheet: Set wsMeta = wb.Worksheets("pack_meta")
    Dim fmtVer As Long: fmtVer = CLngSafe(ReadMetaValue(wsMeta, "pack_format_version"))
    Dim ed As Long: ed = CLngSafe(ReadMetaValue(wsMeta, "embed_dim"))

    Dim metaReason As String
    If Not ValidatePackMeta(fmtVer, ed, metaReason) Then
        reason = metaReason
        Exit Function
    End If

    Dim myDim As Long: myDim = modConfig.GetLong("embed_dim", 1536)
    If ed <> myDim Then
        reason = "ベクトルの次元数が一致しません(パック=" & ed & " / このブック=" & myDim & ")"
        Exit Function
    End If

    ValidatePack = True
End Function

' ValidatePackMeta - バージョン/次元の判定ロジックを切り出した純関数(§7.8)。
'   Excelオブジェクトに触れない(Long/Long/Stringのみ)。形式バージョンが
'   対応版か・次元数が妥当か(1以上)だけを見る。「自分のconfig embed_dimと
'   の一致」はValidatePackが別途判定する(上記コメント参照)。
Public Function ValidatePackMeta(ByVal formatVersion As Long, ByVal embedDim As Long, ByRef reason As String) As Boolean
    reason = ""
    If formatVersion <> modAppDef.PACK_FORMAT_VERSION Then
        reason = "パックの形式バージョン(" & formatVersion & ")が対応バージョン(" & _
            modAppDef.PACK_FORMAT_VERSION & ")と一致しません"
        Exit Function
    End If
    If embedDim < 1 Then
        reason = "パックのベクトル次元数(" & embedDim & ")が不正です"
        Exit Function
    End If
    ValidatePackMeta = True
End Function

' OpenGateReason - modShelfScan.EncryptedFileKind の戻り値("plain"/"enc"/"")を
'   「開いてよいか」の3値へ寄せる純関数(R35 F3a・実機第20報)。
'   "open" 以外は Workbooks.Open を1回も呼ばずに倒してよいという意味。
'   modMigrate.ImportUserData も同じ判定にこの関数を使う(同層・pack配下)。
Public Function OpenGateReason(ByVal kind As String) As String
    Select Case kind
    Case "plain"
        OpenGateReason = "open"
    Case "enc"
        OpenGateReason = "enc"
    Case Else
        OpenGateReason = "unreadable"
    End Select
End Function

' 内部ヘルパー: 共通
Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function CLngSafe(ByVal s As String) As Long
    If IsNumeric(s) Then CLngSafe = CLng(Val(s))
End Function


Private Function SheetExistsIn(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExistsIn = Not (ws Is Nothing)
End Function

Private Function ReadMetaValue(ByVal wsMeta As Worksheet, ByVal key As String) As String
    Dim lastR As Long: lastR = wsMeta.Cells(wsMeta.Rows.count, 1).End(xlUp).row
    If lastR < 1 Then Exit Function
    If lastR = 1 Then
        If StrComp(CStr(wsMeta.Cells(1, 1).Value), key, vbTextCompare) = 0 Then
            ReadMetaValue = CStr(wsMeta.Cells(1, 2).Value)
        End If
        Exit Function
    End If

    Dim arr As Variant: arr = wsMeta.Range(wsMeta.Cells(1, 1), wsMeta.Cells(lastR, 2)).Value
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(CStr(arr(i, 1)), key, vbTextCompare) = 0 Then
            ReadMetaValue = CStr(arr(i, 2))
            Exit Function
        End If
    Next i
End Function


' 内部ヘルパー: 取込(ImportPackDialog)
Private Function LoadPackChunksAndVectors(ByVal wb As Workbook, ByRef ids() As String, ByRef sources() As String, _
        ByRef pages() As Long, ByRef summaries() As String, ByRef keywords() As String, _
        ByRef fullTexts() As String, ByRef vectors() As String) As Long
    ReDim ids(0 To 0): ReDim sources(0 To 0): ReDim pages(0 To 0)
    ReDim summaries(0 To 0): ReDim keywords(0 To 0): ReDim fullTexts(0 To 0): ReDim vectors(0 To 0)

    Dim wsC As Worksheet: Set wsC = wb.Worksheets("pack_chunks")
    Dim lastC As Long: lastC = wsC.Cells(wsC.Rows.count, 1).End(xlUp).row
    If lastC < 2 Then Exit Function

    Dim arrC As Variant: arrC = wsC.Range(wsC.Cells(2, 1), wsC.Cells(lastC, 6)).Value
    Dim n As Long: n = UBound(arrC, 1) - LBound(arrC, 1) + 1

    Dim wsV As Worksheet: Set wsV = wb.Worksheets("pack_vectors")
    Dim vecMap As Object: Set vecMap = CreateObject("Scripting.Dictionary")
    Dim lastV As Long: lastV = wsV.Cells(wsV.Rows.count, 1).End(xlUp).row
    If lastV >= 2 Then
        Dim arrV As Variant: arrV = wsV.Range(wsV.Cells(2, 1), wsV.Cells(lastV, 2)).Value
        Dim j As Long
        For j = LBound(arrV, 1) To UBound(arrV, 1)
            Dim vid As String: vid = CStr(arrV(j, 1))
            If LenB(vid) > 0 Then
                If Not vecMap.Exists(vid) Then vecMap.Add vid, CStr(arrV(j, 2))
            End If
        Next j
    End If

    Dim tmpIds() As String: ReDim tmpIds(0 To n - 1)
    Dim tmpSources() As String: ReDim tmpSources(0 To n - 1)
    Dim tmpPages() As Long: ReDim tmpPages(0 To n - 1)
    Dim tmpSummaries() As String: ReDim tmpSummaries(0 To n - 1)
    Dim tmpKeywords() As String: ReDim tmpKeywords(0 To n - 1)
    Dim tmpFullTexts() As String: ReDim tmpFullTexts(0 To n - 1)
    Dim tmpVectors() As String: ReDim tmpVectors(0 To n - 1)
    Dim cnt As Long: cnt = 0

    Dim i As Long
    For i = LBound(arrC, 1) To UBound(arrC, 1)
        Dim cid As String: cid = CStr(arrC(i, 1))
        If LenB(cid) > 0 Then
            tmpIds(cnt) = cid
            tmpSources(cnt) = CStr(arrC(i, 2))
            ' R42 C1(受入PA-01/SR-02): pageがLong超過等の壊れたパックを取り込むと
            ' CLngが素の例外を投げ、呼び出し元がOn Error GoTo 0のままブック残留・
            ' ScreenUpdating未復元になっていた。範囲外はErr.RaiseしLoadFailedで
            ' 正規に倒す(Andは短絡しないのでIsNumericの判定は外側のIfで分ける)。
            Dim pageVal As Variant: pageVal = arrC(i, 3)
            If IsNumeric(pageVal) Then
                If pageVal >= 0 And pageVal <= 999999 And pageVal = Fix(pageVal) Then
                    tmpPages(cnt) = CLng(pageVal)
                Else
                    Err.Raise 6, "modPack", "pack_chunks 行" & (i + 1) & " の page が不正: " & CStr(pageVal)
                End If
            End If
            tmpSummaries(cnt) = CStr(arrC(i, 4))
            tmpKeywords(cnt) = CStr(arrC(i, 5))
            tmpFullTexts(cnt) = CStr(arrC(i, 6))
            If vecMap.Exists(cid) Then
                tmpVectors(cnt) = CStr(vecMap(cid))
            Else
                tmpVectors(cnt) = ""
            End If
            cnt = cnt + 1
        End If
    Next i

    If cnt > 0 Then
        ReDim Preserve tmpIds(0 To cnt - 1)
        ReDim Preserve tmpSources(0 To cnt - 1)
        ReDim Preserve tmpPages(0 To cnt - 1)
        ReDim Preserve tmpSummaries(0 To cnt - 1)
        ReDim Preserve tmpKeywords(0 To cnt - 1)
        ReDim Preserve tmpFullTexts(0 To cnt - 1)
        ReDim Preserve tmpVectors(0 To cnt - 1)
        ids = tmpIds: sources = tmpSources: pages = tmpPages
        summaries = tmpSummaries: keywords = tmpKeywords: fullTexts = tmpFullTexts
        vectors = tmpVectors
    End If
    LoadPackChunksAndVectors = cnt
End Function

Private Sub ImportChunksDedup(ids() As String, sources() As String, pages() As Long, summaries() As String, _
        keywords() As String, fullTexts() As String, vectors() As String, ByVal n As Long, ByVal authorName As String, _
        ByVal originOverride As String, ByRef importedCount As Long, ByRef skippedCount As Long)
    importedCount = 0
    skippedCount = 0

    Dim wsK As Worksheet: Set wsK = EnsureKnowledgeSheet()
    Dim wsV As Worksheet: Set wsV = EnsureVectorSheet()
    If wsK Is Nothing Or wsV Is Nothing Then Exit Sub

    Dim existingHashes As Object: Set existingHashes = BuildExistingHashSet(wsK)

    Dim outK() As Variant: ReDim outK(1 To n, 1 To 9)
    Dim outV() As Variant: ReDim outV(1 To n, 1 To 2)
    Dim vecCount As Long          ' 実際に書くベクトル行数(空ベクトルは書かない)
    Dim addedStamp As String: addedStamp = modUtil.NowStamp()
    ' 部門チャンネル経由なら "channel:"&部門名、手渡しパックなら "pack:"&作者名。
    ' 消す側(modChannel/modShelfStore)が探すタグと必ず同じ値になるよう、
    ' 呼び出し側から受け取った値をそのまま書く(レビュー C-1)。
    Dim originStr As String
    If LenB(Trim$(originOverride)) > 0 Then
        originStr = Trim$(originOverride)
    Else
        originStr = "pack:" & authorName
    End If

    Dim i As Long
    For i = 0 To n - 1
        Dim hashHex As String: hashHex = ExtractHashFromChunkId(ids(i))
        If LenB(hashHex) = 0 Then hashHex = modUtil.Fnv1a64Hex(modUtil.NormalizeForHash(fullTexts(i)))

        If existingHashes.Exists(hashHex) Then
            skippedCount = skippedCount + 1
        Else
            existingHashes.Add hashHex, True
            importedCount = importedCount + 1
            outK(importedCount, COL_ID) = ids(i)
            ' 2026-08-01(R12-2-1): source/summary/keywords/full_text は他者が
            ' 作った資料・部門パック由来の未信頼テキスト。セルへ書く前に
            ' 数式インジェクション対策を通す(セキュリティ監査3)。
            outK(importedCount, COL_SOURCE) = modUtilText.SanitizeForCell(sources(i))
            outK(importedCount, COL_ORIGIN) = originStr
            outK(importedCount, COL_PAGE) = pages(i)
            outK(importedCount, COL_SUMMARY) = modUtilText.SanitizeForCell(summaries(i))
            outK(importedCount, COL_KEYWORDS) = modUtilText.SanitizeForCell(keywords(i))
            outK(importedCount, COL_FULLTEXT) = modUtilText.SanitizeForCell(modUtil.SafeLeft(fullTexts(i), 32000))
            outK(importedCount, COL_ADDED) = addedStamp
            outK(importedCount, COL_EMBEDDED) = IIf(LenB(vectors(i)) > 0, 1, 0)

            ' 2026-07-28(レビュー M-17): ベクトルが空のチャンクは
            ' my_vectors に空行を作らない。作ると、あとで埋め込みが走った
            ' ときに同じ chunk_id の行がもう1本できて恒久的に残る
            ' (単段検索では同じチャンクが topK の枠を2つ占める)。
            If LenB(vectors(i)) > 0 Then
                vecCount = vecCount + 1
                outV(vecCount, 1) = ids(i)
                outV(vecCount, 2) = modUtil.SafeLeft(vectors(i), 32000)
            End If
        End If
    Next i

    ' 2026-07-28(レビュー M-17): 数千行×9列(full_text は最大32,000字)を
    ' 1回の Range 代入で書くと、大型正典の初回接続で実行時エラー7
    ' (メモリが不足しています)になり得る。modShelf.IngestFile は同じ処理を
    ' 既に200行バッチ化して同じ事故を回避済みだったので、こちらも揃える。
    If importedCount > 0 Then
        Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
        If lastK < 1 Then lastK = 1
        Dim firstKRow As Long: firstKRow = lastK + 1
        If firstKRow < 2 Then firstKRow = 2
        WriteRowsBatched wsK, firstKRow, outK, importedCount, 9
    End If

    If vecCount > 0 Then
        Dim lastV As Long: lastV = wsV.Cells(wsV.Rows.count, 1).End(xlUp).row
        If lastV < 1 Then lastV = 1
        Dim firstVRow As Long: firstVRow = lastV + 1
        If firstVRow < 2 Then firstVRow = 2
        WriteRowsBatched wsV, firstVRow, outV, vecCount, 2
        ' R12-H-6: my_vectors を増やしたら検索側のセッションキャッシュは古い。
        ' 件数が変わるので印だけでも作り直されるが、世代を進める場所を
        ' 「ベクトルを書いた所」に統一しておく(検知漏れを型で防ぐ)。
        On Error Resume Next
        modVecCache.BumpGeneration
        On Error GoTo 0
    End If
End Sub

' 配列を200行ずつに割って書き込む(1回の巨大な Range 代入で err#7 になるのを
' 避けるため。modShelf.IngestFile と同じ方式)。
Private Sub WriteRowsBatched(ByVal ws As Worksheet, ByVal firstRow As Long, _
                             ByRef src As Variant, ByVal rowCount As Long, ByVal colCount As Long)
    Const BATCH_ROWS As Long = 200
    Dim start As Long
    For start = 1 To rowCount Step BATCH_ROWS
        Dim nRows As Long
        nRows = rowCount - start + 1
        If nRows > BATCH_ROWS Then nRows = BATCH_ROWS

        Dim chunkArr() As Variant: ReDim chunkArr(1 To nRows, 1 To colCount)
        Dim r As Long, c As Long
        For r = 1 To nRows
            For c = 1 To colCount
                chunkArr(r, c) = src(start + r - 1, c)
            Next c
        Next r

        Dim rowTop As Long: rowTop = firstRow + start - 1
        ws.Range(ws.Cells(rowTop, 1), ws.Cells(rowTop + nRows - 1, colCount)).Value = chunkArr
    Next start
End Sub

' my_knowledge/my_vectors共通: 無ければveryHiddenで新規作成しヘッダを書く
' (modShelf.bas/modEmbed.bas の同名Ensure*Sheetと同じ設計。複製理由は上部コメント参照)。
Private Function EnsureVeryHiddenSheet(ByVal sheetName As String, ByVal hdr As Variant) As Worksheet
    Dim ws As Worksheet: Set ws = GetSheet(sheetName)
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = sheetName
        Dim i As Long
        For i = LBound(hdr) To UBound(hdr)
            ws.Cells(1, i + 1).Value = hdr(i)
        Next i
        On Error Resume Next
        ws.Visible = 2   ' xlSheetVeryHidden
        On Error GoTo 0
    End If
    Set EnsureVeryHiddenSheet = ws
    Exit Function
Fail:
    Set EnsureVeryHiddenSheet = Nothing
End Function

Private Function EnsureKnowledgeSheet() As Worksheet
    Set EnsureKnowledgeSheet = EnsureVeryHiddenSheet(modAppDef.SH_KNOWLEDGE, _
        Array("chunk_id", "source", "origin", "page", "summary", "keywords", "full_text", "added_at", "embedded"))
End Function

Private Function EnsureVectorSheet() As Worksheet
    Set EnsureVectorSheet = EnsureVeryHiddenSheet(modAppDef.SH_VECTORS, Array("chunk_id", "vector_csv"))
End Function

' chunk_id「bs::HASH::pN::cN」のHASH部分を集めたDictionaryを返す
' (modShelf.bas BuildExistingHashSetと同じ設計。複製理由は上部コメント参照)。
Private Function BuildExistingHashSet(ByVal wsK As Worksheet) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    If wsK Is Nothing Then
        Set BuildExistingHashSet = dict
        Exit Function
    End If

    Dim lastK As Long: lastK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then
        Set BuildExistingHashSet = dict
        Exit Function
    End If

    If lastK = 2 Then
        AddHashIfNew dict, ExtractHashFromChunkId(CStr(wsK.Cells(2, 1).Value))
    Else
        Dim ids As Variant: ids = wsK.Range(wsK.Cells(2, 1), wsK.Cells(lastK, 1)).Value
        Dim i As Long
        For i = LBound(ids, 1) To UBound(ids, 1)
            AddHashIfNew dict, ExtractHashFromChunkId(CStr(ids(i, 1)))
        Next i
    End If
    Set BuildExistingHashSet = dict
End Function

Private Sub AddHashIfNew(ByVal dict As Object, ByVal h As String)
    If LenB(h) = 0 Then Exit Sub
    If Not dict.Exists(h) Then dict.Add h, True
End Sub

Private Function ExtractHashFromChunkId(ByVal chunkId As String) As String
    If LenB(chunkId) = 0 Then Exit Function
    Dim parts() As String: parts = Split(chunkId, "::")
    If UBound(parts) - LBound(parts) + 1 < 2 Then Exit Function
    ExtractHashFromChunkId = parts(LBound(parts) + 1)
End Function

