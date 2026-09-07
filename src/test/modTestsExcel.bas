Attribute VB_Name = "modTestsExcel"
Option Explicit

' ============================================================================
' modTestsExcel - Excel専用・🩺診断から実行する簡易E2Eスモークテスト
' ----------------------------------------------------------------------------
' 注意(重要): 本モジュールはExcel専用である。Worksheets/Application/
'   ThisWorkbook等のExcelオブジェクトに触れるため、tools/run_lo_tests.py の
'   モード1(LibreOffice上でRunAllPureTestsを実際に実行する純ロジックテスト)
'   の対象には含めない(PURE_ALLOWLISTに載せない)。ただしモード2(全モジュール
'   の構文コンパイルチェック・実行はしない)には他の全モジュールと同様に
'   含まれるため、本モジュールは常に構文的に正しくなければならない
'   (MASTER_SPEC §11.2)。
'
' 実行方法: .xlsmを開いた状態(=modBoot.Bootが一度走った状態が望ましいが、
'   必須ではない。各所がOn Error Resume Nextで自衛しているため未実行でも
'   クラッシュはしない)で、VBEからRunExcelE2ESmokeTestを直接実行するか、
'   将来 modDiag 側の「🩺診断」ボタンに配線する(本Waveでは modDiag.bas は
'   担当外のため配線コードそのものは追加しない。§7.1 modDiag契約に
'   RunExcelE2ESmokeTestの呼び出しを追加するかはWave4/5判断)。
'
' 何を検証するか(MASTER_SPEC §7.8 Wave3指示どおり):
'   1. ui_state初期化(mode=quickをセット)
'   2. Tempフォルダにダミーtxtを書き出す
'   3. modShelf.IngestFile でその資料を本棚へ取込む
'   4. my_manifest/my_knowledge の行数(件数)が期待どおり増えたことを検証
'   5. modRetrieve.Search でその資料がヒットすることを検証
'   6. modAsk.Answer("...", "quick") が非空で出典表記([本棚:...])を
'      含むことを検証(mock_llm=TRUEのダミー応答は常に出典例を含む設計
'      なので、mock環境でも意味のある確認になる)
'   7. モックの重複排除(pack取込のImportChunksDedupと同じ
'      Fnv1a64Hex(NormalizeForHash(full_text))ベースの仕組みを
'      modShelf.IngestFileが使っている)を、同一内容・別ファイル名の
'      2件目取込で確認する
'   8. 後始末(取込んだテスト資料の削除・一時ファイルの削除・ui_state復元)
'
' modPack(ExportPackDialog/ImportPackDialog)について:
'   この2つはFileDialogのユーザー操作が前提のPublic Sub(MASTER_SPEC §7.4)
'   であり、ヘッドレスに自動実行することはできない(ダイアログを
'   スクリプトで操作する手段を本Waveの担当範囲では持たない)。
'   このため「パック往復+重複スキップ」の確認は、
'     (a) RunExcelE2ESmokeTest内では、modPackの重複排除が使うのと
'         全く同じ判定ロジック(Fnv1a64Hex+NormalizeForHash+
'         chunk_idハッシュ部の一致判定)を modShelf.IngestFile 経由で
'         自動検証し(上記7番)、
'     (b) 実際のExportPackDialog/ImportPackDialogそのもの(ファイル
'         ダイアログ操作を含む)は、別関数 RunExcelPackDialogRoundTrip
'         として人が操作することを前提に用意する(このSubは
'         RunExcelE2ESmokeTestからは自動的に呼ばない。誤って実行すると
'         保存/選択ダイアログが表示され処理がブロックされるため)。
'   本当の意味での「ダイアログ操作込みの受入確認」は
'   docs/40_受入チェックリスト15分.md(§11.3)でカバーする。
' ============================================================================

Private Const TEST_SOURCE_1 As String = "_mybookshelf_e2e_smoke_test.txt"
Private Const TEST_SOURCE_2 As String = "_mybookshelf_e2e_smoke_test_dup.txt"
Private Const TEST_CONTENT As String = _
    "ゾウガメ工房の年次点検チェックリストについてのテスト資料です。" & vbLf & _
    "点検は毎年4月に実施し、ゾウガメ工房の担当者が立ち会います。"
' R39 F004(受入テスト再現): 9字の質問は ambiguous_max_chars(既定10)以下で、本棚に
' 資料が複数あると聞き返し(modClarify)になり「[本棚:」を含まず FAIL した。
' 資料名と目的を含む1文にして、通常回答の経路を通す。
Private Const TEST_QUERY As String = "ゾウガメ工房の年次点検はいつ実施しますか"

' ----------------------------------------------------------------------------
' RunExcelE2ESmokeTest - 自動実行のメインエントリ(ダイアログ操作なし)。
'   結果はmodTestRunnerに集計し、最後にMsgBoxでPASS/FAIL件数を表示する。
' ----------------------------------------------------------------------------
Public Sub RunExcelE2ESmokeTest()
    modTestRunner.ResetTests

    Dim path1 As String, path2 As String
    path1 = TempFilePath(TEST_SOURCE_1)
    path2 = TempFilePath(TEST_SOURCE_2)

    On Error GoTo Cleanup

    ' 0) 後始末忘れ対策: 前回異常終了分が残っていれば先に掃除しておく
    PreCleanUpTestSources

    ' 1) ui_state初期化(mode=quick)。ui_stateシートが無い環境では
    '    スキップし、その旨を1件のCheckとして記録する(致命的ではない)。
    If SetUiStateMode("quick") Then
        modTestRunner.Check "E2E_ui_state初期化_mode_quick", True
    Else
        modTestRunner.Check "[SKIP] E2E_ui_state初期化: シート未検出のためスキップ", True, _
            "ui_stateシートが見つからないため、mode設定はスキップした" & _
            "(modAsk の Answer は質問文とモードを直接引数で受け取るので、この検証には影響しない)"
    End If

    ' 2) Tempにダミーtxtを書き出す
    WriteTextFile path1, TEST_CONTENT
    modTestRunner.Check "E2E_ダミーtxt作成_1件目", FileExistsOnDisk(path1), "path=" & path1

    ' 3) modShelf.IngestFileで取込む(1件目)
    Dim knowledgeBefore As Long: knowledgeBefore = modShelf.TotalChunks()
    Dim manifestNamesBefore() As String, manifestStatsBefore() As String
    Dim manifestCountBefore As Long
    manifestCountBefore = modShelf.SourceList(manifestNamesBefore, manifestStatsBefore)

    Dim status1 As String
    status1 = modShelf.IngestFile(path1, "self")
    modTestRunner.Check "E2E_IngestFile_1件目_status", (status1 = "done" Or status1 = "partial"), _
        "status1=" & status1

    ' 4) 行数検証: my_knowledgeの総チャンク数が1件目の分だけ増えたこと、
    '    my_manifestに1件目のsourceが現れたこと。
    Dim knowledgeAfter1 As Long: knowledgeAfter1 = modShelf.TotalChunks()
    modTestRunner.Check "E2E_行数検証_knowledge増加_1件目", (knowledgeAfter1 > knowledgeBefore), _
        "before=" & knowledgeBefore & " after=" & knowledgeAfter1

    Dim names() As String, stats() As String
    Dim n As Long: n = modShelf.SourceList(names, stats)
    modTestRunner.Check "E2E_行数検証_manifest増加", (n > manifestCountBefore), _
        "before=" & manifestCountBefore & " after=" & n
    modTestRunner.Check "E2E_行数検証_manifestに1件目が現れる", ContainsName(names, n, TEST_SOURCE_1)

    ' 5) modRetrieve.Searchでヒットすることを検証(topKは大きめにして、
    '    本棚に他の資料が多くても取りこぼさないようにする)。
    Dim hits() As Hit
    Dim nHits As Long
    nHits = modRetrieve.Search(TEST_QUERY, 200, hits)
    modTestRunner.Check "E2E_Search_ヒット件数0件超", (nHits > 0), "nHits=" & nHits
    If nHits > 0 Then
        modTestRunner.Check "E2E_Search_取込んだ資料がヒットに含まれる", _
            HitsContainSource(hits, nHits, TEST_SOURCE_1)
    End If

    ' 6) modAsk.Answer("quick")が非空+出典表記を含むことを検証
    Dim ans As String
    ans = modAsk.Answer(TEST_QUERY, "quick")
    modTestRunner.Check "E2E_Answer_quick_非空", (LenB(Trim$(ans)) > 0), "len=" & Len(ans)
    modTestRunner.Check "E2E_Answer_quick_出典表記を含む", (InStr(ans, "[本棚:") > 0), "ans=" & modUtil.SafeLeft(ans, 300)

    ' 6b) 出典アクセサ(Peek View用・添字0始まり)が i=0 で落ちないこと。
    '     2026-07-28(レビュー C-2): 内部配列が ReDim(1 To n) なのにアクセサが
    '     i をそのまま添字に使っていたため、i=0 で実行時エラー9になり、
    '     呼び出し側(modPeek)が On Error で握って「出典チップが1枚も出ない」
    '     という静かな全滅になっていた。呼び出し側と同じ 0 始まりで叩いて
    '     資料名が取れることを確認する(握らずにここで検出させる)。
    Dim hitN As Long: hitN = modAsk.LastHitCount()
    modTestRunner.Check "E2E_出典アクセサ_件数0件超", (hitN > 0), "LastHitCount=" & hitN
    If hitN > 0 Then
        Dim acc0 As String
        Dim accErr As Long
        On Error Resume Next
        Err.Clear
        acc0 = modAsk.LastHitSource(0)
        accErr = Err.Number
        Err.Clear
        On Error GoTo Cleanup
        modTestRunner.Check "E2E_出典アクセサ_LastHitSource(0)が例外を出さない", _
            (accErr = 0), "Err=" & accErr
        modTestRunner.Check "E2E_出典アクセサ_LastHitSource(0)が非空", _
            (LenB(acc0) > 0), "src0=" & acc0
    End If

    ' 7) 重複排除の確認: 全く同じ内容を別ファイル名で2件目として取込むと、
    '    chunk_idのハッシュ部が1件目と一致するためチャンクは追加されない
    '    (modShelf.IngestFile内のBuildExistingHashSet+Fnv1a64Hexによる判定。
    '    modPack.ImportChunksDedupが使う判定ロジックと同一)。
    WriteTextFile path2, TEST_CONTENT
    Dim knowledgeBefore2 As Long: knowledgeBefore2 = modShelf.TotalChunks()
    Dim status2 As String
    status2 = modShelf.IngestFile(path2, "self")
    modTestRunner.Check "E2E_IngestFile_2件目_status", (status2 = "done" Or status2 = "partial"), _
        "status2=" & status2
    Dim knowledgeAfter2 As Long: knowledgeAfter2 = modShelf.TotalChunks()
    modTestRunner.Check "E2E_重複スキップ_同一内容はチャンク追加されない", _
        (knowledgeAfter2 = knowledgeBefore2), _
        "before=" & knowledgeBefore2 & " after=" & knowledgeAfter2 & _
        "(パック取込の重複排除と同じFnv1a64Hexハッシュ判定を検証)"

    ' 2件目もmanifestには載る(chunk_count=0の状態で記録される仕様)ことの確認
    Dim names2() As String, stats2() As String
    Dim n2 As Long: n2 = modShelf.SourceList(names2, stats2)
    modTestRunner.Check "E2E_重複スキップ後もmanifestに2件目が記録される", ContainsName(names2, n2, TEST_SOURCE_2)

    ' 正常系はハンドラ本体(Resume)を跨いで後始末へ入る
    ' (Resume はエラーが起きていないと実行時エラー20になる)。
    GoTo CleanupBody

Cleanup:
    ' Err の内容は Resume でクリアされるので先に控える。
    Dim unexpectedNum As Long: unexpectedNum = Err.Number
    Dim unexpectedDesc As String: unexpectedDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かない(下の後始末が素通りし、
    ' 後始末で起きたエラーが呼び出し元へ飛ぶ)。まずハンドラを抜ける。
    Resume CleanupBody

CleanupBody:
    On Error Resume Next
    If unexpectedNum <> 0 Then
        modTestRunner.Check "E2E_想定外エラー", False, _
            "Err=" & unexpectedNum & ": " & unexpectedDesc
    End If

    ' 8) 後始末: 取込んだテスト資料を削除し、一時ファイルも削除する。
    modShelf.DeleteSource TEST_SOURCE_1
    modShelf.DeleteSource TEST_SOURCE_2
    DeleteFileIfExists path1
    DeleteFileIfExists path2
    On Error GoTo 0

    MsgBox modTestRunner.ReportText(), vbInformation, "マイ本棚AI Excel E2Eスモークテスト結果"
End Sub

' ----------------------------------------------------------------------------
' RunExcelPackDialogRoundTrip - modPackExport.ExportPackDialog/ImportPackDialogを
'   実際に呼び出す、人が操作することを前提にした補助確認(モジュール冒頭の
'   コメント参照)。RunExcelE2ESmokeTestからは呼ばれない。ダイアログが
'   表示されるので、テスト実行者は以下の手順で操作すること:
'     1) 「本棚全体をパックにして書き出しますか?」→いいえ(アクティブ行の
'        資料1件、または任意)を選び、保存先ダイアログで適当な場所に保存する
'     2) 続けて表示される取込ダイアログで、いま保存したファイルを選ぶ
'        (1回目の取込は新規、2回目に同じファイルをもう一度取込むと
'        「重複スキップ」の効果がImportPackDialog自身のMsgBoxで確認できる)
' ----------------------------------------------------------------------------
Public Sub RunExcelPackDialogRoundTrip()
    Dim before As Long: before = modShelf.TotalChunks()

    MsgBox "パック往復確認を開始します。" & vbLf & _
        "次のダイアログで書き出し先を選んでください。" & vbLf & _
        "続いて表示される取込ダイアログでは、いま書き出したファイルを選んでください。", _
        vbInformation, "パック往復確認(手動)"

    modPackExport.ExportPackDialog
    modPack.ImportPackDialog

    Dim afterFirstImport As Long: afterFirstImport = modShelf.TotalChunks()

    Dim again As Long
    again = MsgBox("もう一度同じファイルを取込んで、重複スキップ(0件取込/全件スキップ)を" & vbLf & _
        "確認しますか?", vbYesNo + vbQuestion, "パック往復確認(手動)")
    If again = vbYes Then
        modPack.ImportPackDialog
    End If

    Dim afterSecondImport As Long: afterSecondImport = modShelf.TotalChunks()

    MsgBox "確認用の参考値です(重複スキップが効いていれば2回目の取込でチャンク数は増えないはずです)。" & vbLf & _
        "書き出し前: " & before & vbLf & _
        "1回目取込後: " & afterFirstImport & vbLf & _
        "2回目取込後: " & afterSecondImport, vbInformation, "パック往復確認(手動)・結果"
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

Private Function TempFilePath(ByVal fileName As String) As String
    Dim dir_ As String
    dir_ = Environ$("TEMP")
    If LenB(dir_) = 0 Then dir_ = Environ$("TMP")
    If LenB(dir_) = 0 Then dir_ = "C:\Temp"
    If Right$(dir_, 1) <> "\" Then dir_ = dir_ & "\"
    TempFilePath = dir_ & fileName
End Function

Private Sub WriteTextFile(ByVal path As String, ByVal content As String)
    Dim f As Long: f = FreeFile
    Open path For Output As #f
    Print #f, content
    Close #f
End Sub

Private Function FileExistsOnDisk(ByVal path As String) As Boolean
    On Error Resume Next
    FileExistsOnDisk = (Len(Dir$(path)) > 0)
    On Error GoTo 0
End Function

Private Sub DeleteFileIfExists(ByVal path As String)
    On Error Resume Next
    If Len(Dir$(path)) > 0 Then Kill path
    On Error GoTo 0
End Sub

' 前回の異常終了で本棚に残っている可能性のあるテスト資料を先に片付ける
' (毎回のRunExcelE2ESmokeTestが「本棚がまっさらな状態」を仮定しないための保険。
' DeleteSourceは対象が存在しなくても安全に何もしない設計)。
Private Sub PreCleanUpTestSources()
    On Error Resume Next
    modShelf.DeleteSource TEST_SOURCE_1
    modShelf.DeleteSource TEST_SOURCE_2
    On Error GoTo 0
End Sub

' ui_stateシート(A=key, B=value)の"mode"行を書き換える(無ければ末尾に追記)。
' シート自体が見つからない場合はFalseを返すだけで例外は出さない。
Private Function SetUiStateMode(ByVal modeValue As String) As Boolean
    On Error GoTo NotFound
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_UISTATE)
    On Error GoTo 0

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    Dim foundRow As Long: foundRow = 0
    For i = 1 To lastRow
        If StrComp(CStr(ws.Cells(i, 1).Value), "mode", vbTextCompare) = 0 Then
            foundRow = i
            Exit For
        End If
    Next i

    If foundRow = 0 Then
        foundRow = lastRow + 1
        If foundRow < 1 Then foundRow = 1
        ws.Cells(foundRow, 1).Value = "mode"
    End If
    ws.Cells(foundRow, 2).Value = modeValue

    SetUiStateMode = True
    Exit Function
NotFound:
    SetUiStateMode = False
End Function

Private Function ContainsName(ByRef names() As String, ByVal n As Long, ByVal target As String) As Boolean
    If n < 1 Then Exit Function
    Dim i As Long
    For i = LBound(names) To (LBound(names) + n - 1)
        If StrComp(names(i), target, vbTextCompare) = 0 Then
            ContainsName = True
            Exit Function
        End If
    Next i
End Function

Private Function HitsContainSource(hits() As Hit, ByVal nHits As Long, ByVal target As String) As Boolean
    If nHits < 1 Then Exit Function
    Dim i As Long
    For i = 1 To nHits
        If StrComp(hits(i).source, target, vbTextCompare) = 0 Then
            HitsContainSource = True
            Exit Function
        End If
    Next i
End Function
