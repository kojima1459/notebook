Attribute VB_Name = "modIntegrity"
Option Explicit

' ============================================================================
' modIntegrity - データ整合性の観測点(2026-08-05 R18-2b/2d・実機第5報⑧)
' ----------------------------------------------------------------------------
' なぜ要るのか:
'   「昨日入れた資料が今日は0件になっている」という報告に対して、アプリ側に
'   【気付く仕組みが1つも無かった】。表示は my_manifest の chunk_count を
'   そのまま出すだけで実データと突き合わせず、保存が次回に残ったかどうかを
'   確かめる記録も無かった(調査agent1)。憲章§4-2「観測性ファースト」。
'
' ここが持つのは2つだけ:
'   (1) 台帳と実データの突合(ReconcileChunkCount) — マイ本棚のカードが
'       my_knowledge の実行数と食い違ったら、実行数を正として台帳を直す。
'   (2) 保存の履歴と起動時の突合(RecordSaveMark / WarnAtStartup) —
'       保存が成功した時点の (my_knowledge行数, ブックのフルパス) を ui_state に
'       控え、次の起動で行数が減っていたら1回だけ警告する。あわせて、ブックが
'       %TEMP%・Temp1_*.zip 配下(=zipの中から直接開いた/一時展開コピー)なら
'       「保存が次回に残らない場所」であることを起動時に1回伝える。
'
' 置き場が基盤層(src/core)である理由:
'   modBoot は30,000字上限まで残り僅かで判定本体を置けず、modShelfStore /
'   modShelf も同様に逼迫している(憲章§4-6)。判定材料は my_knowledge /
'   my_manifest / ui_state のシートだけで、上位層を1つも呼ばない(R1)。
'   modDiag.WarnIfReadOnly が同じ理由で同じ層に置かれているのと同型。
'
' 判定そのもの(DataShrunk / IsVolatilePath / ReconcileStatText)は Excel に
' 触れない純ロジックへ切り出し、modTestsPure15 がゴールデンで固定する。
' ============================================================================

' ui_state のキー(modState.SaveState/LoadState 経由。作法は nexus_theme と同じ)。
Private Const KEY_LAST_ROWS As String = "integrity_last_rows"
Private Const KEY_LAST_PATH As String = "integrity_last_path"

' 起動時の警告は1セッション1回まで(Bootが二度走っても二度は出さない)。
Private mWarned As Boolean

' ----------------------------------------------------------------------------
' ReconcileStatText - 台帳の統計文字列の chunk_count を actualN へ直す(純)。
' ----------------------------------------------------------------------------
' statText は modShelf.SourceList が組み立てた
' "status|ingested_at|chunk_count|error_note|origin"。3つ目が actualN と
' 一致していれば statText をそのまま返す(=呼び出し側は何もしない合図)。
' error_note に "|" が混じっていても Join で元どおりに戻る(触るのは3つ目だけ)。
' 5つ未満の壊れた文字列は触らない(直せないものを壊さない)。
Public Function ReconcileStatText(ByVal statText As String, ByVal actualN As Long) As String
    ReconcileStatText = statText
    Dim f() As String: f = Split(statText, "|")
    If UBound(f) - LBound(f) + 1 < 5 Then Exit Function
    Dim idx As Long: idx = LBound(f) + 2
    Dim shown As Long: shown = -1
    If IsNumeric(f(idx)) Then shown = CLng(f(idx))
    If shown = actualN Then Exit Function
    f(idx) = CStr(actualN)
    ReconcileStatText = Join(f, "|")
End Function

' ----------------------------------------------------------------------------
' IndexOfName - names(0..n-1) から target の位置を探す(無ければ -1。純ロジック)。
' ----------------------------------------------------------------------------
' R18-2b の突合で my_knowledge の全行に対して呼ばれるため、直前に一致した位置
' (hint)を先に1回だけ試す。my_knowledge は source 単位の連続ブロックで追記
' されるので、ほぼ全ての行がこの1回の比較で決まり、実質 O(行数)で済む
' (毎回先頭から探すと 資料数×行数 の比較になり、2万行の本棚で描画が固まる)。
' hint が範囲外・不一致でも結果は変わらない(速さだけの助言で、正しさは
' 常に線形探索が保証する)。比較は StrComp(vbTextCompare)=大小無視で、
' manifest 側の照合(FindConflictingManifestPath 等)と作法を揃える。
Public Function IndexOfName(ByRef names() As String, ByVal n As Long, _
                            ByVal target As String, ByVal hint As Long) As Long
    IndexOfName = -1
    If n < 1 Then Exit Function
    If LenB(target) = 0 Then Exit Function
    If hint >= 0 And hint < n Then
        If StrComp(names(hint), target, vbTextCompare) = 0 Then
            IndexOfName = hint
            Exit Function
        End If
    End If
    Dim i As Long
    For i = 0 To n - 1
        If StrComp(names(i), target, vbTextCompare) = 0 Then
            IndexOfName = i
            Exit Function
        End If
    Next i
End Function

' ----------------------------------------------------------------------------
' ReconcileChunkCount - (2b の本体)カード1枚ぶんの突合と自動修復。
' ----------------------------------------------------------------------------
' 一致していればシートにも触らずログも書かない(通常はここで即戻る)。
' 食い違うときだけ実行数を正として manifest の5列目を直し、usage_log に
' manifest_fix を1行残す。台帳を直せなくても(シート無し・行が見つからない・
' 書込み失敗)戻り値は実行数に揃うので、少なくとも画面が嘘をつくことは無い。
' 突合の失敗が一覧の描画を止めてはならないので全体を OERN で包む(憲章§4-4)。
Public Function ReconcileChunkCount(ByVal sourceName As String, ByVal statText As String, _
                                    ByVal actualN As Long) As String
    ReconcileChunkCount = statText
    On Error Resume Next
    Dim fixed As String: fixed = ReconcileStatText(statText, actualN)
    If fixed = statText Then Exit Function
    ReconcileChunkCount = fixed

    Dim wsM As Worksheet: Set wsM = GetSheet(modAppDef.SH_MANIFEST)
    If wsM Is Nothing Then Exit Function
    Dim r As Long: r = ManifestRowByName(wsM, sourceName)
    If r < 2 Then Exit Function
    Dim shownTxt As String: shownTxt = CStr(wsM.Cells(r, 5).Value)
    wsM.Cells(r, 5).Value = actualN
    modLog.LogUsage "manifest_fix", "", modUtil.SafeLeft(sourceName, 200) & _
        ": 台帳" & shownTxt & "件 → 実データ" & actualN & "件へ修復"
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' RecordSaveMark - 保存が成功した瞬間の事実を控える(2d)。
' ----------------------------------------------------------------------------
' modShelfBatch.SaveCheckpoint が ThisWorkbook.Save の【直前】に1行呼ぶ
' (R18H FA-3: 後に呼ぶとディスク上のマークだけが常に1世代古くなる)。
' 加えて、意図した削除の直後にも呼ぶ(modShelf.DeleteSource /
' modShelfStore.RemoveRowsByOrigin(Prefix)。保存はしない=削除は保存の合図では
' ない。ui_state はブックと一緒に保存/破棄されるのでどちらでも辻褄が合う)。
' 控えるのは (my_knowledge の行数, ThisWorkbook.FullName) の2つだけ。
' 「保存した」と「次に開いたときに残っていた」は別の事実で、後者を確かめる
' 手段がこれまで1つも無かった(調査agent1 §4: zip直開き・一時展開コピーは
' 保存も成功し ReadOnly でもないため、アプリからは完全に正常に見える)。
' 記録の失敗が保存を壊してはならないので OERN で包む。
Public Sub RecordSaveMark()
    On Error Resume Next
    modState.SaveState KEY_LAST_ROWS, CStr(KnowledgeRowCount())
    modState.SaveState KEY_LAST_PATH, ThisWorkbook.FullName
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' DataShrunk - 「前回保存した時より資料が減ったか」(純ロジック)。
' ----------------------------------------------------------------------------
' prevRows <= 0(記録が無い・壊れている)は判定しない=初回起動や旧ブックから
' の移行で、根拠なく警告を出さない。増えている/同じは当然 False。
Public Function DataShrunk(ByVal prevRows As Long, ByVal curRows As Long) As Boolean
    If prevRows <= 0 Then Exit Function
    DataShrunk = (curRows < prevRows)
End Function

' ----------------------------------------------------------------------------
' IsVolatilePath - 「保存しても次回に残らない場所」か(純ロジック)。
' ----------------------------------------------------------------------------
' zip をダブルクリックして中の xlsm を直接開くと、Windows は
' %TEMP%\Temp1_<zip名>.zip\ のような一時フォルダへ展開してそこを開く。
' 保存は成功し ReadOnly でもないので、アプリからは完全に正常に見えるが、
' 次に開くのは別の展開コピーで、前回の取込は1件も残っていない
' (調査agent1 §5 順位2「検知ゼロ」)。判定材料はパス文字列だけなので、
' 環境に依らない素朴な一致で見る(正規表現も FileSystemObject も使わない
' =管理端末で塞がれても必ず動く)。
'   ・"temp1_"  : zip 直開きの展開フォルダ名(Temp1_, Temp2_ ... と続く)
'   ・".zip\"   : zip の中を直接開いている
'   ・tempA/tempB(=Environ("TEMP")/Environ("TMP") の実値)配下
'
' 2026-08-05(R18H FA-4 / A-M3・B-M5): "\temp\" の部分一致を廃止した。
' D:\temp\約款\ のような【正規の作業フォルダ】や \\share\Temp\ の共有を
' 一時フォルダと誤判定し、正しく保存できている端末に「保存が次回に残らない
' 場所です」という嘘を毎回出していた(憲章§3-3の逆・その機能はもう使われない)。
' 判定は実際の %TEMP%/%TMP% との【前方一致】に限る。環境変数はここでは読まず
' 引数で受け取る(呼び出し元 WarnAtStartup が読む)=この関数は純ロジックの
' ままで、テストが端末の環境変数に左右されない。
' 大文字小文字は無視する(パスの表記は端末ごとに揺れる)。
Public Function IsVolatilePath(ByVal folderPath As String, _
                               Optional ByVal tempA As String = "", _
                               Optional ByVal tempB As String = "") As Boolean
    Dim p As String: p = LCase$(Trim$(folderPath))
    If LenB(p) = 0 Then Exit Function
    If InStr(p, "temp1_") > 0 Then IsVolatilePath = True
    If InStr(p, ".zip\") > 0 Then IsVolatilePath = True
    If UnderDir(p, tempA) Then IsVolatilePath = True
    If UnderDir(p, tempB) Then IsVolatilePath = True
End Function

' p(小文字化済み)が dirPath そのもの、またはその配下か(前方一致・R18H FA-4)。
' 末尾に区切りを補ってから比べる: 補わないと "C:\Temp" が "C:\Temp2\..." にも
' 当たる(1文字違いのフォルダを巻き込む誤検知)。
Private Function UnderDir(ByVal p As String, ByVal dirPath As String) As Boolean
    Dim d As String: d = LCase$(Trim$(dirPath))
    If LenB(d) = 0 Then Exit Function
    If Right$(d, 1) <> "\" Then d = d & "\"
    Dim pp As String: pp = p
    If Right$(pp, 1) <> "\" Then pp = pp & "\"
    If Len(pp) < Len(d) Then Exit Function
    UnderDir = (Left$(pp, Len(d)) = d)
End Function

' Environ の読み出し(失敗しても空文字で返る)。判定そのものは純ロジック側。
Private Function SafeEnv(ByVal keyName As String) As String
    On Error Resume Next
    SafeEnv = Environ$(keyName)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' ShrinkWarnMsg / VolatileWarnMsg - 警告文(純ロジック・BMPの文字だけ)。
' ----------------------------------------------------------------------------
' 非BMPの絵文字は CP932 の実行文へ変換される往復で化け、ダイアログには
' 化けたまま出る(実機第5報⑤)。ここでは記号を一切使わず日本語だけで書く。
' 「何が起きたか」+「次の一手」を必ず両方言う(憲章§3-3)。
Public Function ShrinkWarnMsg(ByVal prevRows As Long, ByVal curRows As Long, _
                              ByVal prevPath As String, ByVal curPath As String) As String
    ShrinkWarnMsg = "前回このアプリを保存したときは資料の中身が " & prevRows & _
        " 件ありましたが、今は " & curRows & " 件になっています。" & vbLf & _
        "別の場所にあるコピーを開いている可能性があります。" & vbLf & vbLf & _
        "前回保存した場所: " & modUtil.SafeLeft(prevPath, 200) & vbLf & _
        "今開いている場所: " & modUtil.SafeLeft(curPath, 200) & vbLf & vbLf & _
        "心当たりが無いときは、この画面を管理者に見せてください。"
End Function

Public Function VolatileWarnMsg(ByVal curPath As String) As String
    VolatileWarnMsg = "保存が次回に残らない場所でこのファイルを開いています。" & vbLf & _
        "取り込んだ資料は、次にこのファイルを開いたときに残りません。" & vbLf & vbLf & _
        "今の場所: " & modUtil.SafeLeft(curPath, 200) & vbLf & vbLf & _
        "配布された zip を右クリックして「すべて展開」し、" & vbLf & _
        "展開されたフォルダの中のファイルを開き直してください。"
End Function

' ----------------------------------------------------------------------------
' WarnAtStartup - 起動時の突合(2d 本体)。modBoot から1行だけ呼ばれる。
' ----------------------------------------------------------------------------
' 順番は「置き場所の警告」が先: 一時フォルダで開いていることが分かっていれば、
' 件数が減っているのは当然の結果で、原因を先に伝えた方が利用者は動ける。
' どちらも1セッション1回まで。機能は止めない(閲覧・質問はそのまま使える)。
' 起動を止めないよう例外は外へ出さない。
'
' 2026-08-05(R18H FA-4 / A-M3・B-M5): (i) 置き場所の判定に渡すのは
' ThisWorkbook.Path(フォルダ)。FullName を渡すと、ファイル名に "temp1_" を
' 含むだけの資料まで一時扱いになる。(ii) 一時フォルダの警告を出しても
' 【減少チェックは打ち切らない】。従来はここで Exit Sub していたため、
' zip直開きの端末では「資料が減っている」という本命の異常を最後まで一度も
' 検知できなかった(2つは別の事実で、片方が他方を隠す理由は無い)。
Public Sub WarnAtStartup()
    On Error Resume Next
    If mWarned Then Exit Sub
    mWarned = True

    Dim curPath As String: curPath = ThisWorkbook.FullName
    Dim bookDir As String: bookDir = ThisWorkbook.Path
    Dim curRows As Long: curRows = KnowledgeRowCount()

    If IsVolatilePath(bookDir, SafeEnv("TEMP"), SafeEnv("TMP")) Then
        modLog.LogUsage "integrity_warn", "volatile_path", modUtil.SafeLeft(curPath, 300)
        MsgBox VolatileWarnMsg(curPath), vbExclamation, modAppDef.APP_NAME
    End If

    Dim prevRows As Long: prevRows = 0
    Dim prevTxt As String: prevTxt = modState.LoadState(KEY_LAST_ROWS, "")
    If IsNumeric(prevTxt) Then prevRows = CLng(Val(prevTxt))
    If Not DataShrunk(prevRows, curRows) Then Exit Sub

    Dim prevPath As String: prevPath = modState.LoadState(KEY_LAST_PATH, "(記録なし)")
    modLog.LogUsage "integrity_warn", "rows_shrunk", _
        "前回" & prevRows & "件 → 今回" & curRows & "件 / 前回パス=" & _
        modUtil.SafeLeft(prevPath, 200)
    MsgBox ShrinkWarnMsg(prevRows, curRows, prevPath, curPath), _
        vbExclamation, modAppDef.APP_NAME
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー(シートを直接読む。上位層は呼ばない=R1)
' ----------------------------------------------------------------------------

' my_knowledge の行数(ヘッダ除く)。シートが無ければ0。
Private Function KnowledgeRowCount() As Long
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then Exit Function
    Dim lastK As Long: lastK = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then Exit Function
    KnowledgeRowCount = lastK - 1
    On Error GoTo 0
End Function

' file_name(2列目)でmanifest行を探す(見つからなければ0)。
Private Function ManifestRowByName(ByVal wsM As Worksheet, ByVal fileName As String) As Long
    Dim lastM As Long: lastM = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
    If lastM < 2 Then Exit Function
    Dim arr As Variant: arr = wsM.Range(wsM.Cells(2, 2), wsM.Cells(lastM, 2)).Value
    If Not IsArray(arr) Then
        If StrComp(CStr(arr), fileName, vbTextCompare) = 0 Then ManifestRowByName = 2
        Exit Function
    End If
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(CStr(arr(i, 1)), fileName, vbTextCompare) = 0 Then
            ManifestRowByName = i + 1
            Exit Function
        End If
    Next i
End Function

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function
