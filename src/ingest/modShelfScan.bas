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

' ----------------------------------------------------------------------------
' FileHeadHex / EncryptedByHeader - 開く【前】に、パスワードで暗号化された
'   Office文書かどうかを先頭バイトから見分ける(2026-08-16 R33波3 W3-11)。
' ----------------------------------------------------------------------------
' なぜ要るのか(実機で踏み抜いた事故):
'   暗号化ブックに対して Workbooks.Open / Documents.Open を「パスワード引数
'   なし」で呼ぶと、Excel/Word は【パスワード入力モーダル】を出して待つ。
'   無人の自動同期では誰も押さないので、そこで永久に止まる(利用者には
'   「開いた瞬間に固まった」としか見えない)。DisplayAlerts=False では
'   このプロンプトは抑止できない ―― modExtractorWord は既に
'   wdAlertsNone を設定済みなのに同じプロンプトが出る、が実証。
'   一方で「引数なしでの再試行」自体は 2026-07-29 の実機事故(Password 引数を
'   付けると Open 自体が失敗する端末があった)への保険なので消せない。
'   そこで【開く前に判別し、暗号化と分かったものだけ再試行に到達させない】。
'
' 見分けの根拠(推測ではなく確認済み):
'   ・.xlsx/.xlsm/.docx などの OOXML は ZIP コンテナ。先頭は "PK" = 50 4B。
'     (openpyxl で作った .xlsx / zipfile で作った .docx の実測がともに
'      50 4B 03 04 14 00 00 00)
'   ・ECMA-376 の暗号化文書は、中身の ZIP を丸ごと OLE 複合ドキュメントへ
'     包んだもの(EncryptionInfo / EncryptedPackage ストリーム)。したがって
'     先頭は OLE の署名 D0 CF 11 E0 A1 B1 1A E1 になる
'     (olefile.MAGIC と msoffcrypto の判別ロジックで確認)。
'   ・したがって「拡張子は OOXML なのに中身が OLE」= 暗号化、と断じてよい。
'
' 【判別できないものを「暗号化ではない」と言わない】
'   .xls / .doc(旧バイナリ形式)は暗号化の有無にかかわらず常に OLE なので、
'   この方法では区別できない。PDF も暗号化されていても先頭は "%PDF" のまま。
'   これらは "" (=判別不能)を返し、従来どおりの経路(引数なし再試行を含む)
'   へ流す。ここを「OLE だから暗号化」と倒すと、正常な .xls / .doc が
'   1本残らず取り込めなくなる。安全側は【従来どおり】であって【弾く】ではない。
'
' 戻り値は文字列3種:
'   "enc"   … パスワードで暗号化されている(開かずに失敗させてよい)
'   "plain" … 暗号化されていない(素の ZIP)
'   ""      … 判別不能(従来どおりの経路へ流す)
' ----------------------------------------------------------------------------
Public Function EncryptedByHeader(ByVal ext As String, ByVal headHex As String) As String
    Select Case LCase$(Trim$(ext))
    Case "xlsx", "xlsm", "xltx", "xltm", "docx", "docm", "dotx", "dotm", "pptx", "pptm"
        ' OOXML(ZIPコンテナ)の拡張子だけが判定対象。
    Case Else
        Exit Function                      ' "" = 判別不能
    End Select

    Dim h As String: h = UCase$(Trim$(headHex))
    ' 4バイト読めていないファイル(空/途中で切れている)は判定しない。
    If Len(h) < 8 Then Exit Function

    If Left$(h, 4) = "504B" Then
        ' "PK" で始まる = ZIP。Office が作る文書は必ず 50 4B 03 04 だが、
        ' 空アーカイブ(50 4B 05 06)等も ZIP なので2バイトで見る。
        EncryptedByHeader = "plain"
    ElseIf Left$(h, 8) = "D0CF11E0" Then
        EncryptedByHeader = "enc"
    End If
    ' どちらでもない(PDFを .xlsx に改名した等)は "" のまま=判別不能。
End Function

' ファイル先頭4バイトを大文字16進8字で返す。読めなければ ""(判別不能扱い)。
' 例外は外へ出さない ―― 見分けに失敗しても取込そのものは従来どおり続ける。
Public Function FileHeadHex(ByVal path As String) As String
    Dim st As Object
    Dim b As Variant
    Dim n As Long: n = -1

    On Error GoTo HeadFail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1          ' adTypeBinary
    st.Open
    st.LoadFromFile path
    b = st.Read(4)
    st.Close
    Set st = Nothing

    ' 0バイトのファイルでは Read が配列を返さない。UBound がそこで落ちても
    ' n = -1 のまま=空文字を返す(判別不能)。
    On Error Resume Next
    n = UBound(b)
    Err.Clear
    On Error GoTo HeadFail

    Dim s As String
    Dim i As Long
    For i = 0 To n
        s = s & Right$("0" & Hex$(b(i)), 2)
    Next i
    FileHeadHex = UCase$(s)
    Exit Function

HeadFail:
    Resume HeadCleanup
HeadCleanup:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Set st = Nothing
    On Error GoTo 0
    FileHeadHex = ""
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
'   2026-08-05(R18-2f・調査agent1 §3(c)): Err で捕まえられない列挙失敗が
'   残っていた。UNC(\\server\share)では Dir$ が途中で切れても【エラーを
'   返さず】短いリストや空を返すことがあり、その0件が「フォルダが空になった」
'   として消失判定へ流れ、実在するファイルが DeleteSource されていた
'   (復旧不能)。0件でも台帳に当該スコープの資料が残っているなら、
'   「空になった」より「読めなかった」の方が圧倒的にありそうな説明なので、
'   enumFailed 側へ倒す(憲章§3-5: 消える方に倒さない)。
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
        ' R18-2f: 0件かつ台帳に当該スコープの資料あり=列挙失敗とみなす。
        If EnumLooksFailed(cnt, ManifestScopeCount(folderNorm)) Then enumFailed = True
    Else
        ReDim Preserve names(0 To cnt - 1)
        fileNames = names
    End If
    fileCount = cnt
End Sub

' ----------------------------------------------------------------------------
' EnumLooksFailed - 列挙結果を「読めなかった」と見なすか(純ロジック・R18-2f)。
' ----------------------------------------------------------------------------
' diskN=0 かつ manifestN>0 のときだけ True。
'   ・diskN>0 : 少なくとも読めている。途中で切れていても、消失判定は
'     「台帳に有りディスクに無い」ファイル単位で行われるので、ここで倒すと
'     正当な削除(利用者が本当に消したファイル)まで永久に反映されなくなる。
'   ・manifestN=0: そのフォルダの資料をまだ1件も持っていない=消える資料が
'     無いので、失敗と見なす意味が無い(初回の同期で毎回失敗扱いにしない)。
Public Function EnumLooksFailed(ByVal diskN As Long, ByVal manifestN As Long) As Boolean
    EnumLooksFailed = (diskN = 0 And manifestN > 0)
End Function

' folderNorm 配下の manifest スコープ件数(LoadManifestScope と同じ条件)。
' 読めなければ0(=従来どおり消失判定へ進む。判定材料が無いのに倒さない)。
Private Function ManifestScopeCount(ByVal folderNorm As String) As Long
    On Error Resume Next
    Dim wsM As Worksheet: Set wsM = GetSheet(modAppDef.SH_MANIFEST)
    If wsM Is Nothing Then Exit Function
    Dim mPaths() As String, mNames() As String, mModified() As Date
    Dim mSize() As Double, mStatus() As String, mCount As Long
    LoadManifestScope wsM, folderNorm, mPaths, mNames, mModified, mSize, mStatus, mCount
    ManifestScopeCount = mCount
    On Error GoTo 0
End Function

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

' ----------------------------------------------------------------------------
' EnsureManifestTextFormat - my_manifest の文字列列を "@"(文字列)書式へ固定。
' ----------------------------------------------------------------------------
' 2026-08-10(R27波3-5): file_path(1) / file_name(2) / error_note(7) /
' origin(9) には、利用者の資料名・OSのパス・例外メッセージがそのまま入る。
' 書式が「標準」のセルへ "-案件A_見積.pdf" や "=集計.xlsx" を .Value で書くと、
' Excel はそれを【数式】として解釈しようとして実行時エラーになり、
' modShelfStore.UpsertManifestRow ごと落ちる ―― つまり先頭がハイフンや
' 等号のファイル名は、それだけで取込が台帳の1行を書けずに死ぬ。業務資料に
' "-" 始まりの名前は珍しくない。列を "@" にしておけば必ず文字列として入る。
'
' 置き場の判断: 本来は modShelfStore.EnsureManifestSheet の中(同モジュールの
' EnsureKnowledgeSheet が :100-107 でまったく同じことをしているのと対称)だが、
' あちらは30,000字上限まで余裕が無く1行も入らない(憲章§4-6)。同じ ingest 層で
' manifest を読む係である本モジュールへ置き、取込の入口から呼ぶ。
' 冪等: 何度呼んでも同じ結果(EnsureKnowledgeSheet が毎回無条件に張るのと同じ
' 作法)。NumberFormat はブックに保存されるので、一度通れば以後の全経路に効く。
'
' 呼び出し元(台帳へ書き込みが始まる前の入口3箇所。呼び出し側は上限まで余裕が
' 無いモジュールが多いため、注記はこちらへ集約し、あちらは1行だけ置く):
'   ・modShelfBatch.AddFilesResult … ボタンからの取込(利用者が名前を選ぶ経路)
'   ・modShelfSync.SyncNow          … フォルダ同期(起動時の自動同期を含む)
'   ・modMigrate.ImportUserData     … 引き継ぎ(台帳がシートごと差し替わる)
' 資料1件だけを直接 IngestFile へ渡す経路(modVault のナレッジ登録・
' modUIShelf のスクショ取込)は、ファイル名をアプリ側で組み立てており
' 先頭が "-"/"=" になり得ないため呼んでいない。
Public Sub EnsureManifestTextFormat()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = modShelfStore.EnsureManifestSheet()
    If ws Is Nothing Then Exit Sub
    ws.Columns(1).NumberFormat = "@"    ' file_path
    ws.Columns(2).NumberFormat = "@"    ' file_name
    ws.Columns(7).NumberFormat = "@"    ' error_note
    ws.Columns(9).NumberFormat = "@"    ' origin
    Err.Clear
    On Error GoTo 0
End Sub

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
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup9
FailCleanup9:
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
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup10
FailCleanup10:
    If errNum <> 53 Then
        On Error Resume Next
        modLog.LogError "E0801", "modShelfSync.SyncNow", "SafeFileDateTime: " & modUtil.SafeLeft(path, 300) & " : " & errDesc, errNum
        On Error GoTo 0
    End If
End Function
