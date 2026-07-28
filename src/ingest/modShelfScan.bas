Attribute VB_Name = "modShelfScan"
Option Explicit

' ========================================
' modShelfScan - フォルダ同期の「見る」層(ディスク列挙 / manifestスコープ読込)
'
' modShelfSync から切り出した、判断を持たない部分。
' ディスクに何があるか・台帳に何が載っているかを読むだけで、
' 取り込むか消すかは一切決めない。差分判定(modShelfSync)と分けることで、
' 「消えたと誤判定した」ときにどちらを疑えばよいかがはっきりする。
'
' 切り出しの理由(2026-07-28): modShelfSync が契約上限30,000字に対し
' 残り994字となり、列挙失敗の検知(レビュー H-7)を入れた時点で
' 次の修正が入らなくなっていた(レビュー I-2)。
' ========================================

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Public Function EnsureTrailingSlash(ByVal path As String) As String
    Dim p As String: p = path
    If Len(p) = 0 Then Exit Function
    If Right$(p, 1) <> "\" And Right$(p, 1) <> "/" Then p = p & "\"
    EnsureTrailingSlash = p
End Function

' pathからファイル名部分を除いたディレクトリ部分(末尾区切り文字含む)を返す。
Public Function DirOf(ByVal path As String) As String
    Dim fn As String: fn = modUtil.FileNameOf(path)
    If Len(fn) >= Len(path) Then
        DirOf = ""
    Else
        DirOf = Left$(path, Len(path) - Len(fn))
    End If
End Function

' folderNorm配下(第1階層のみ・対応拡張子のみ)のファイル名一覧を返す。
' サブフォルダはDir()の既定動作(vbDirectory属性を指定しない)により
' 自動的に除外される。
' enumFailed (2026-07-28 レビュー H-7):
'   最初の Dir$ が失敗したことを呼び出し側へ伝える。従来は
'   On Error Resume Next で握って「0件」と同じ結果になっていたため、
'   アクセス拒否・ネットワーク瞬断・AVブロックが「フォルダが空になった」
'   と解釈され、消失判定が同期スコープの全資料を DeleteSource していた。
'   フォルダ不存在は別途保護済みだが、存在確認を通ったあとの列挙失敗には
'   保護が無かった。「読めなかった」と「無かった」は別物として扱う。
Public Sub EnumFolderFiles(ByVal folderNorm As String, ByRef fileNames() As String, _
                            ByRef fileCount As Long, Optional ByRef enumFailed As Boolean)
    Dim names() As String: ReDim names(0 To 63)
    Dim cnt As Long: cnt = 0
    enumFailed = False

    On Error Resume Next
    Err.Clear
    Dim nm As String: nm = Dir$(folderNorm & "*.*")
    If Err.Number <> 0 Then
        enumFailed = True
        Err.Clear
        On Error GoTo 0
        ReDim fileNames(0 To 0)
        fileCount = 0
        Exit Sub
    End If
    On Error GoTo 0

    Do While LenB(nm) > 0
        If IsSupportedExtLocal(nm) And Not IsExcludedFile(folderNorm, nm) Then
            If cnt > UBound(names) Then ReDim Preserve names(0 To (UBound(names) + 1) * 2 - 1)
            names(cnt) = nm
            cnt = cnt + 1
        End If
        nm = Dir$()
    Loop

    If cnt = 0 Then
        ReDim fileNames(0 To 0)
    Else
        ReDim Preserve names(0 To cnt - 1)
        fileNames = names
    End If
    fileCount = cnt
End Sub

' ----------------------------------------------------------------------------
' IsExcludedFile - 走査から必ず除く files(2026-07-28追加)
' ----------------------------------------------------------------------------
' 実務で確実に踏む2つの地雷を塞ぐ。どちらも「置き場所を間違えた」だけで
' 起きるのに、症状が分かりにくい。
'
' 1) このアプリ自身(.xlsm)
'    対応拡張子に xlsm が含まれるため、本棚フォルダを .xlsm と同じ場所に
'    設定すると、アプリが自分自身を資料として開こうとする。
'    既に開いているブックなので取込は失敗し、毎回同期のたびに
'    err_log へ E0302 が積まれる。運用担当が「同じフォルダでいいですか」と
'    聞いてくるのは自然な発想なので、構造的に禁じるのではなく無視する。
'
' 2) Office の一時ファイル(~$で始まる)
'    誰かが Word/Excel で資料を開いている間、同じフォルダに ~$資料名.docx が
'    作られる。拡張子は docx/xlsx なので対応形式に該当してしまい、
'    中身は本文ではないため取り込むとゴミが本棚に入る。
'    「共有フォルダに置いて、誰かが開いている」は常に起きる状況なので必須。
Public Function IsExcludedFile(ByVal folderNorm As String, ByVal fileName As String) As Boolean
    ' Office一時ファイル(~$資料名.docx 等)
    If Left$(fileName, 2) = "~$" Then
        IsExcludedFile = True
        Exit Function
    End If

    ' このアプリ自身。フルパスで比べる(同名の別ファイルを巻き添えにしない)。
    On Error Resume Next
    Dim me_ As String: me_ = ThisWorkbook.FullName
    On Error GoTo 0
    If LenB(me_) > 0 Then
        If StrComp(folderNorm & fileName, me_, vbTextCompare) = 0 Then
            IsExcludedFile = True
            Exit Function
        End If
    End If

    ' パック(.xlsx)の取り違え防止までは行わない。パックは「📥パック取込」
    ' から明示的に入れるものなので、本棚フォルダに置く運用は想定しない。
End Function

Public Function IsSupportedExtLocal(ByVal fileName As String) As Boolean
    Dim ext As String: ext = modUtil.ExtOf(fileName)
    If LenB(ext) = 0 Then Exit Function
    Dim supported As String: supported = modExtractor.SupportedExts()
    IsSupportedExtLocal = (InStr(1, "," & supported & ",", "," & ext & ",", vbTextCompare) > 0)
End Function

' my_manifestのうち origin=self かつ folderNorm配下の行だけを抜き出す
' (SyncNowの比較・消失検知スコープ)。一括読み取り→配列フィルタで
' ループ内Range直接アクセスを避ける(§12)。
Public Sub LoadManifestScope(ByVal wsM As Worksheet, ByVal folderNorm As String, _
        ByRef outPaths() As String, ByRef outNames() As String, ByRef outModified() As Date, _
        ByRef outSize() As Double, ByRef outStatus() As String, ByRef outCount As Long)
    outCount = 0
    ReDim outPaths(0 To 0)
    ReDim outNames(0 To 0)
    ReDim outModified(0 To 0)
    ReDim outSize(0 To 0)
    ReDim outStatus(0 To 0)
    If wsM Is Nothing Then Exit Sub

    Dim lastM As Long: lastM = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
    If lastM < 2 Then Exit Sub

    Dim arr As Variant: arr = wsM.Range(wsM.Cells(2, 1), wsM.Cells(lastM, 9)).Value
    Dim n As Long: n = UBound(arr, 1) - LBound(arr, 1) + 1

    Dim tmpPaths() As String: ReDim tmpPaths(0 To n - 1)
    Dim tmpNames() As String: ReDim tmpNames(0 To n - 1)
    Dim tmpModified() As Date: ReDim tmpModified(0 To n - 1)
    Dim tmpSize() As Double: ReDim tmpSize(0 To n - 1)
    Dim tmpStatus() As String: ReDim tmpStatus(0 To n - 1)
    Dim cnt As Long: cnt = 0

    Dim folderLower As String: folderLower = LCase$(folderNorm)

    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(CStr(arr(i, 9)), "self", vbTextCompare) = 0 Then
            Dim p As String: p = CStr(arr(i, 1))
            ' 2026-07-28(レビュー H-9): 拡張子でも絞る。
            ' ディスク側の列挙(EnumFolderFiles)は対応拡張子だけを見るのに、
            ' 台帳側のスコープ判定には拡張子の条件が無かった。
            ' スクショ取込は shelf_folder に .jpg を保存して台帳に載せるが、
            ' jpg は対応拡張子ではないためディスク側には現れず、
            ' 「台帳に有りディスクに無い=消失」と誤判定されて翌朝の同期で
            ' 無言削除されていた(ファイルは残るが再取込もされない)。
            ' 両側で同じ拡張子フィルタを使う。
            If IsSupportedExtLocal(p) And LCase$(EnsureTrailingSlash(DirOf(p))) = folderLower Then
                tmpPaths(cnt) = p
                tmpNames(cnt) = CStr(arr(i, 2))

                On Error Resume Next
                tmpModified(cnt) = CDate(arr(i, 3))
                On Error GoTo 0

                If IsNumeric(arr(i, 4)) Then tmpSize(cnt) = CDbl(arr(i, 4))
                tmpStatus(cnt) = CStr(arr(i, 6))
                cnt = cnt + 1
            End If
        End If
    Next i

    If cnt > 0 Then
        ReDim Preserve tmpPaths(0 To cnt - 1)
        ReDim Preserve tmpNames(0 To cnt - 1)
        ReDim Preserve tmpModified(0 To cnt - 1)
        ReDim Preserve tmpSize(0 To cnt - 1)
        ReDim Preserve tmpStatus(0 To cnt - 1)
        outPaths = tmpPaths
        outNames = tmpNames
        outModified = tmpModified
        outSize = tmpSize
        outStatus = tmpStatus
    End If
    outCount = cnt
End Sub

Public Function FindManifestIndexByName(names() As String, ByVal cnt As Long, ByVal fileName As String) As Long
    FindManifestIndexByName = -1
    Dim i As Long
    For i = 0 To cnt - 1
        If StrComp(names(i), fileName, vbTextCompare) = 0 Then
            FindManifestIndexByName = i
            Exit Function
        End If
    Next i
End Function

' shelf_folder自体が見つからない(削除・リネーム・OneDriveオフライン等)ときに、
' そのフォルダ配下由来のmanifest行だけを status="missing" に更新する
' (実削除はしない。§13「フォルダ削除・リネーム(E0502+missing)」)。
Public Sub MarkFolderScopeMissing(ByVal folder As String)
    Dim wsM As Worksheet: Set wsM = GetSheet(modAppDef.SH_MANIFEST)
    If wsM Is Nothing Then Exit Sub

    Dim lastM As Long: lastM = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
    If lastM < 2 Then Exit Sub

    Dim folderLower As String: folderLower = LCase$(EnsureTrailingSlash(folder))

    Dim arr As Variant: arr = wsM.Range(wsM.Cells(2, 1), wsM.Cells(lastM, 9)).Value
    Dim changed As Boolean: changed = False

    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(CStr(arr(i, 9)), "self", vbTextCompare) = 0 Then
            Dim p As String: p = CStr(arr(i, 1))
            If LCase$(EnsureTrailingSlash(DirOf(p))) = folderLower Then
                If CStr(arr(i, 6)) <> "missing" Then
                    arr(i, 6) = "missing"
                    changed = True
                End If
            End If
        End If
    Next i

    If changed Then
        wsM.Range(wsM.Cells(2, 1), wsM.Cells(lastM, 9)).Value = arr
    End If
End Sub

' 失敗時は0を返し呼び出し元(SyncNowの差分判定)を止めないが、共有フォルダの
' アクセス権限・セキュリティソフトのブロック等(エラー52/70想定)を握りつぶさず
' E0801として記録する(🩺診断「直近のエラーをコピー」導線・実機環境の壁対策)。
' エラー53(ファイル未検出。Dir走査直後にファイルが消えた等の一過性競合)は
' 次回同期で自然に解消するためログを汚さない。
Public Function SafeFileLen(ByVal path As String) As Double
    On Error GoTo Fail
    SafeFileLen = CDbl(FileLen(path))
    Exit Function
Fail:
    SafeFileLen = 0
    ' 契約チェック対策: LogErrorのcontext引数は「modX.Y」形式だとYがPublicか
    ' 検証されるため(§7)、Private助手関数の名前ではなく実際のPublic呼び出し
    ' 元(SyncNow)を書き、助手関数名はdetail側に含める。
    ' 2026-07-28(レビュー L-9): Err の退避を On Error Resume Next より前へ。
    ' On Error Resume Next 文そのものが Err をリセットするため、従来は
    ' 番号も説明も空のログしか残らず、権限なのか瞬断なのか判別できなかった。
    Dim errNum As Long: errNum = Err.Number
    Dim errDesc As String: errDesc = Err.Description
    If errNum <> 53 Then
        On Error Resume Next
        modLog.LogError "E0801", "modShelfSync.SyncNow", "SafeFileLen: " & modUtil.SafeLeft(path, 300) & " : " & errDesc, errNum
        On Error GoTo 0
    End If
End Function

Public Function SafeFileDateTime(ByVal path As String) As Date
    On Error GoTo Fail
    SafeFileDateTime = FileDateTime(path)
    Exit Function
Fail:
    SafeFileDateTime = Now
    ' 2026-07-28(レビュー L-9): Err の退避を On Error Resume Next より前へ。
    ' On Error Resume Next 文そのものが Err をリセットするため、従来は
    ' 番号も説明も空のログしか残らず、権限なのか瞬断なのか判別できなかった。
    Dim errNum As Long: errNum = Err.Number
    Dim errDesc As String: errDesc = Err.Description
    If errNum <> 53 Then
        On Error Resume Next
        modLog.LogError "E0801", "modShelfSync.SyncNow", "SafeFileDateTime: " & modUtil.SafeLeft(path, 300) & " : " & errDesc, errNum
        On Error GoTo 0
    End If
End Function
