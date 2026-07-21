Attribute VB_Name = "optDiffDoc"
Option Explicit

' ============================================================================
' optDiffDoc - 約款差分(opt機能・MASTER_SPEC §7.7)
' ----------------------------------------------------------------------------
' 役割:
'   新旧2ファイルを選ばせ、本文を比較して変更点+影響コメントをAI
'   (確認済みのChatGPT関数・modGateway.CallLLM経由)に分析させ、
'   diff_reportシート(毎回削除→再生成)へ書き出す。この層が無くても製品は
'   成立し、modules.jsonから1行削除するだけで撤去できる(§7.7)。
'
' ============================================================================
' 【重要・設計上の逸脱と根拠(Wave4/最終報告で必ず確認してほしい点)】
'   発注コメントには「modExtractor.ExtractFile×2(契約§7.2通りに呼ぶ)」と
'   あったが、これは MASTER_SPEC §7.7 opt層共通契約「コアモジュールへの
'   参照は基盤層(modAppDef/modConfig/modLog/modUtil/modGateway/modTypes)+
'   modUIMain.SetStageのみ可」と直接矛盾する(modExtractorはsrc/ingest=
'   部品層であり基盤層ではない)。さらに tools/vba_lint.py の
'   check_layer_dependency は opt層からのコア参照を基盤層+SetStageのみに
'   機械的に制限しており、modExtractorを参照すると自担当モジュールで
'   ERRORになる(自己検証必須要件に抵触)。
'   本実装は「§7.7厳守」を優先し、modExtractorには一切依存しない。
'   代わりに本モジュール内で完結する簡易テキスト抽出(txt/md/csvのみ、
'   ADODB.Streamで直接読む)を実装した。docx/doc/pdf/xlsx等の比較は
'   このバージョンでは非対応とし、選択時に丁寧な案内文を出す。
'   これは「対応形式の縮小」という明確な逸脱であり、最終レポートの
'   逸脱提案として報告する(MASTER_SPEC冒頭の指示どおり、独断で契約を
'   拡張解釈して実装だけ先に進めることはしない)。
'   将来この機能拡張が必要になった場合の対処案(いずれか要判断):
'     (a) §7.7の opt層コア参照制限を「基盤層+modUIMain.SetStage+
'         modExtractor(このAPIのみ)」のように緩和し、lintの
'         check_layer_dependencyにopt向け例外を追加する。
'     (b) modExtractorWord/Excel/Acrobat相当の抽出ロジックを本モジュール内に
'         複製する(コード重複は増えるが層分離は保たれる)。
' ============================================================================

' === SIGNATURE ASSUMPTION ===
'   本モジュールが呼ぶリボン関数はChatGPT(modGateway.CallLLM経由)のみで、
'   これはMASTER_SPEC §7.1で仕様確定済み・確認済みの関数のため、未確認の
'   シグネチャ仮定は無い(§7.7コメント「確認済み関数のみ使用」のとおり)。
'   ここで実際に「仮定」しているのは、対応ファイル形式をtxt/md/csvに限定した
'   上記の設計判断そのものである。
'   仕様回答(または層制限の見直し方針)が来たら直す行:
'     ・ExtractPlainText() の対応拡張子判定(If Not (ext = "txt" Or ...))を
'       広げる、または上記(a)/(b)の方針でmodExtractor経由に差し替える。
' ================================================================================

Private Const DIFF_SHEET As String = "diff_report"

Public Function Ping() As Boolean
    Ping = True
End Function

' ----------------------------------------------------------------------------
' CompareTwoDocsDialog - 新旧2ファイル選択→抽出→ChatGPT差分分析→
'   diff_reportシート書き出し(MASTER_SPEC §7.7)。
' ----------------------------------------------------------------------------
Public Sub CompareTwoDocsDialog()
    On Error GoTo Fail

    Dim oldPath As String
    oldPath = PickCompareFile("比較する「旧」ファイルを選んでください")
    If LenB(oldPath) = 0 Then Exit Sub

    Dim newPath As String
    newPath = PickCompareFile("比較する「新」ファイルを選んでください")
    If LenB(newPath) = 0 Then Exit Sub

    Dim oldText As String, newText As String, extractErr As String

    If Not ExtractPlainText(oldPath, oldText, extractErr) Then
        MsgBox "旧ファイル「" & modUtil.FileNameOf(oldPath) & "」を読み込めませんでした。" & vbLf & _
            extractErr, vbExclamation, modAppDef.APP_NAME
        Exit Sub
    End If
    If Not ExtractPlainText(newPath, newText, extractErr) Then
        MsgBox "新ファイル「" & modUtil.FileNameOf(newPath) & "」を読み込めませんでした。" & vbLf & _
            extractErr, vbExclamation, modAppDef.APP_NAME
        Exit Sub
    End If

    If LenB(Trim$(oldText)) = 0 And LenB(Trim$(newText)) = 0 Then
        MsgBox "どちらのファイルも中身が空でした。" & vbLf & "別のファイルでお試しください。", _
            vbExclamation, modAppDef.APP_NAME
        Exit Sub
    End If

    modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDCD1) & " 差分を分析中…"

    Dim maxChars As Long: maxChars = modConfig.GetLong("max_context_chars", 40000)
    Dim halfChars As Long: halfChars = maxChars \ 2
    If halfChars < 500 Then halfChars = 500

    Dim prompt As String
    prompt = BuildDiffPrompt(modUtil.FileNameOf(oldPath), TruncateWithNotice(oldText, halfChars), _
                              modUtil.FileNameOf(newPath), TruncateWithNotice(newText, halfChars))

    Dim mdl As String: mdl = modConfig.GetString("recommended_model", "gpt-5.5")
    Dim eff As String: eff = modConfig.GetString("deep_draft_effort", "medium")
    Dim vrb As String: vrb = modConfig.GetString("deep_draft_verbosity", "high")
    Dim latency As Long

    Dim resp As String
    resp = modGateway.CallLLM(prompt, "diff", eff, vrb, mdl, latency)

    modUIMain.SetStage ""

    If Left$(resp, 5) = "#ERR:" Then
        Dim code As String: code = ExtractErrorCode(resp)
        If LenB(code) = 0 Then code = "E0202"
        MsgBox modLog.FriendlyMessage(code) & vbLf & "(コード: " & code & ")", vbExclamation, modAppDef.APP_NAME
        Exit Sub
    End If

    WriteDiffReportSheet oldPath, newPath, resp

    modLog.LogUsage "diff", "", "old=" & modUtil.FileNameOf(oldPath) & " new=" & modUtil.FileNameOf(newPath), latency

    MsgBox "差分の分析が終わりました。" & vbLf & _
        "詳しい内容は「" & DIFF_SHEET & "」シートを確認してください。", vbInformation, modAppDef.APP_NAME
    Exit Sub

Fail:
    Dim errDesc As String: errDesc = Err.Description
    Err.Clear
    On Error GoTo 0
    On Error Resume Next
    modUIMain.SetStage ""
    On Error GoTo 0
    modLog.LogError "E0602", "optDiffDoc.CompareTwoDocsDialog", errDesc
    MsgBox modLog.FriendlyMessage("E0602") & vbLf & "(コード: E0602)", vbExclamation, modAppDef.APP_NAME
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー(すべてPrivate: optDiffDocの公開契約はPing/CompareTwoDocsDialogのみ)
' ----------------------------------------------------------------------------

Private Function PickCompareFile(ByVal title As String) As String
    Dim fd As Object
    Set fd = Application.FileDialog(3)   ' msoFileDialogFilePicker(名前付き定数は使わない)
    fd.AllowMultiSelect = False
    fd.Title = title
    fd.Filters.Clear
    fd.Filters.Add "テキスト系ファイル", "*.txt;*.md;*.csv"
    fd.Filters.Add "すべてのファイル", "*.*"

    If fd.Show <> -1 Then Exit Function   ' キャンセル
    If fd.SelectedItems.count < 1 Then Exit Function
    PickCompareFile = CStr(fd.SelectedItems(1))
End Function

' このバージョンの対応形式はtxt/md/csvのみ(上部の設計判断コメント参照)。
' ADODB.StreamでUTF-8として読み込む(modExtractor.ExtractPlainTextと同じ
' 技法だが、opt層のコア参照制限のため独立実装している)。
Private Function ExtractPlainText(ByVal path As String, ByRef outText As String, ByRef errDetail As String) As Boolean
    outText = ""
    errDetail = ""

    Dim ext As String: ext = modUtil.ExtOf(path)
    If Not (ext = "txt" Or ext = "md" Or ext = "csv") Then
        errDetail = "このバージョンの約款差分は txt/md/csv 形式のみに対応しています" & vbLf & _
            "(選ばれたファイル: 拡張子=" & ext & ")。他の形式は一度テキストとして書き出してからお試しください。"
        Exit Function
    End If

    Dim st As Object
    On Error GoTo Failed
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2          ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    outText = st.ReadText
    st.Close
    Set st = Nothing

    ExtractPlainText = True
    Exit Function

Failed:
    errDetail = "ファイルを開けませんでした(他のアプリで開いている、または権限がない可能性があります)。"
    If Not st Is Nothing Then
        On Error Resume Next
        st.Close
        On Error GoTo 0
    End If
    Set st = Nothing
    ExtractPlainText = False
End Function

Private Function TruncateWithNotice(ByVal s As String, ByVal lim As Long) As String
    If Len(s) <= lim Then
        TruncateWithNotice = s
    Else
        TruncateWithNotice = Left$(s, lim) & vbLf & "(一部省略)"
    End If
End Function

Private Function BuildDiffPrompt(ByVal oldName As String, ByVal oldText As String, _
                                  ByVal newName As String, ByVal newText As String) As String
    Dim lang As String: lang = modConfig.GetString("answer_language", "日本語")

    Dim sb As String
    sb = "あなたは社内文書の改定差分を確認する専門家です。以下の「旧」と「新」の本文を条文/項目単位で比較し、" & _
         lang & "で、変更点の一覧と、それぞれの変更が業務にどう影響しうるかの短いコメントを付けて出力してください。" & vbLf
    sb = sb & "本文に書かれていない推測はせず、根拠が無い場合は「不明」と述べてください。" & vbLf & vbLf
    sb = sb & "## 旧: " & oldName & vbLf & oldText & vbLf & vbLf
    sb = sb & "## 新: " & newName & vbLf & newText & vbLf
    BuildDiffPrompt = sb
End Function

' "#ERR:E0202:説明..." -> "E0202"
Private Function ExtractErrorCode(ByVal s As String) As String
    If Left$(s, 5) <> "#ERR:" Then Exit Function
    Dim rest As String: rest = Mid$(s, 6)
    Dim p As Long: p = InStr(rest, ":")
    If p = 0 Then
        ExtractErrorCode = rest
    Else
        ExtractErrorCode = Left$(rest, p - 1)
    End If
End Function

Private Sub WriteDiffReportSheet(ByVal oldPath As String, ByVal newPath As String, ByVal diffText As String)
    Dim ws As Worksheet: Set ws = RecreateDiffSheet()

    Dim lines() As String: lines = SplitToLines(diffText)
    Dim nLines As Long: nLines = LinesCount(lines)
    Dim totalRows As Long: totalRows = 5 + nLines

    Dim arr() As Variant: ReDim arr(1 To totalRows, 1 To 1)
    arr(1, 1) = "■ 約款差分レポート  " & modUtil.NowStamp()
    arr(2, 1) = "旧ファイル: " & modUtil.FileNameOf(oldPath)
    arr(3, 1) = "新ファイル: " & modUtil.FileNameOf(newPath)
    arr(4, 1) = ""
    arr(5, 1) = "[変更点と影響コメント]"

    Dim lo As Long: lo = LBound(lines)
    Dim i As Long
    For i = 0 To nLines - 1
        arr(6 + i, 1) = modUtil.SafeLeft(lines(lo + i), 2000)
    Next i

    ws.Range(ws.Cells(1, 1), ws.Cells(totalRows, 1)).Value = arr
    ws.Columns("A").ColumnWidth = 120

    On Error Resume Next
    ws.Activate
    On Error GoTo 0
End Sub

Private Function RecreateDiffSheet() As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(DIFF_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    ws.Name = DIFF_SHEET
    Set RecreateDiffSheet = ws
End Function

Private Function SplitToLines(ByVal s As String) As String()
    Dim norm As String: norm = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)
    SplitToLines = Split(norm, vbLf)
End Function

Private Function LinesCount(arr() As String) As Long
    On Error GoTo Empty0
    LinesCount = UBound(arr) - LBound(arr) + 1
    Exit Function
Empty0:
    LinesCount = 0
End Function
