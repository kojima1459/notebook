Attribute VB_Name = "modDiag"
Option Explicit

' ============================================================================
' modDiag - Self-diagnostics & visible error reporting
' ----------------------------------------------------------------------------
' Goal: when something does not work, show EXACTLY where the problem is so it
' can be fixed fast. Everything is written to a `diag` sheet (auto-created)
' AND summarised in a message box.
'
' Entry points:
'   RunDiagnostics()  - full health check (bound to the 「診断」 button)
'   TestRibbon()      - just the corporate AI ribbon call
'   ReportError(...)  - helper used by every handler's error trap
' ============================================================================

Private Const DIAG_SHEET As String = "diag"

' ----------------------------------------------------------------------------
' Full health check. Safe to run anytime. Never raises.
' ----------------------------------------------------------------------------
Public Sub RunDiagnostics()
    On Error Resume Next
    Dim r As String
    r = "■ InternalNotebookLM v2 自己診断  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbLf
    r = r & "----------------------------------------" & vbLf

    ' 1. Environment
    r = r & "[環境]" & vbLf
    r = r & "  Excel バージョン : " & Application.Version & vbLf
    r = r & "  OS              : " & Application.OperatingSystem & vbLf
    r = r & "  ユーザー名       : " & Environ$("USERNAME") & vbLf
    r = r & vbLf

    ' 2. Sheets present
    r = r & "[シート存在チェック]" & vbLf
    Dim need As Variant
    need = Array("main", "config", "system_prompt", "department", _
                 "manifest", "knowledge_base", "feedback", "usage_log")
    Dim i As Long
    For i = LBound(need) To UBound(need)
        r = r & "  " & PadR(CStr(need(i)), 16) & " : " & _
            IIf(SheetExists(CStr(need(i))), "OK", "★欠落★") & vbLf
    Next i
    r = r & vbLf

    ' 3. Knowledge base
    r = r & "[ナレッジ]" & vbLf
    Dim kb As Long: kb = modKnowledgeBase.RowCount()
    r = r & "  knowledge_base 行数 : " & kb & IIf(kb = 0, "  ★0件★", "") & vbLf
    Dim dept As String: dept = modUserProfile.CurrentDept()
    r = r & "  現在の部署ID        : " & IIf(LenB(dept) = 0, "(未設定)", dept) & vbLf
    Dim tbl As String: tbl = modKnowledgeBase.BuildRouterTable(dept)
    r = r & "  部署で参照可能な行   : " & CountLines(tbl) & vbLf
    r = r & vbLf

    ' 4. Config keys
    r = r & "[設定]" & vbLf
    r = r & "  mock_llm         : " & modConfig.GetBool("mock_llm", False) & _
            IIf(modConfig.GetBool("mock_llm", False), "  ← ダミー応答モード", "  ← 本番(リボン呼出)") & vbLf
    r = r & "  verifier_enabled : " & modConfig.GetBool("verifier_enabled", True) & vbLf
    r = r & "  router_max_chunks: " & modConfig.GetLong("router_max_chunks", 8) & vbLf
    r = r & vbLf

    ' 5. Prompts present
    r = r & "[システムプロンプト]" & vbLf
    r = r & "  router   : " & LenInfo(modPrompts.GetRouter()) & vbLf
    r = r & "  drafter  : " & LenInfo(modPrompts.GetDrafter()) & vbLf
    r = r & "  verifier : " & LenInfo(modPrompts.GetVerifier()) & vbLf
    r = r & vbLf

    ' 6. Ribbon
    r = r & "[社内AIリボン]" & vbLf
    If modConfig.GetBool("mock_llm", False) Then
        r = r & "  (mock_llm=TRUE のためスキップ。本番確認時は FALSE にして再診断)" & vbLf
    Else
        r = r & "  " & RibbonStatus() & vbLf
    End If

    WriteDiagSheet r
    MsgBox r & vbLf & "（詳細は『diag』シートにも出力しました）", vbInformation, "自己診断結果"
End Sub

' ----------------------------------------------------------------------------
' Test the corporate AI ribbon in isolation and report the exact result/error.
' ----------------------------------------------------------------------------
Public Sub TestRibbon()
    Dim msg As String: msg = RibbonStatus()
    WriteDiagSheet "■ リボン接続テスト  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbLf & msg
    MsgBox msg, vbInformation, "リボン接続テスト"
End Sub

Private Function RibbonStatus() As String
    Dim t0 As Double: t0 = Timer
    On Error GoTo Failed
    Dim res As Variant
    res = Application.Run("ChatGPT", "これはテストです。「OK」とだけ返答してください。")
    Dim ms As Long: ms = CLng((Timer - t0) * 1000)
    RibbonStatus = "応答あり (" & ms & " ms): " & Left$(CStr(res), 200)
    Exit Function
Failed:
    RibbonStatus = "★呼び出し失敗★  Err " & Err.Number & ": " & Err.Description & vbLf & _
                   "  → 『ChatGPT』マクロが見つからない可能性。社内AIリボン(アドイン)が" & vbLf & _
                   "     読み込まれているか、関数名が『ChatGPT』で正しいか確認してください。"
End Function

' ----------------------------------------------------------------------------
' Standard error reporter: call from every handler's error trap so the user
' sees module + procedure + error number + description in one place.
' ----------------------------------------------------------------------------
Public Sub ReportError(ByVal where As String, ByVal errNum As Long, _
                       ByVal errDesc As String, Optional ByVal extra As String = "")
    Dim m As String
    m = "■ エラー発生" & vbLf & _
        "場所 : " & where & vbLf & _
        "番号 : " & errNum & vbLf & _
        "内容 : " & errDesc & vbLf
    If LenB(extra) > 0 Then m = m & "補足 : " & extra & vbLf
    m = m & "時刻 : " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    WriteDiagSheet m
    MsgBox m, vbCritical, "エラー (" & where & ")"
End Sub

' ----------------------------------------------------------------------------
' Helpers
' ----------------------------------------------------------------------------
Private Sub WriteDiagSheet(ByVal text As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(DIAG_SHEET)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = DIAG_SHEET
    End If
    Dim r As Long: r = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If r < 1 Then r = 1
    ' Blank separator line then the block
    Dim lines() As String: lines = Split(text, vbLf)
    Dim i As Long
    ws.Cells(r + 1, 1).value = "================================"
    For i = LBound(lines) To UBound(lines)
        ws.Cells(r + 2 + i, 1).value = lines(i)
    Next i
    ws.Columns("A").ColumnWidth = 100
End Sub

Private Function SheetExists(ByVal name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function

Private Function CountLines(ByVal s As String) As Long
    If LenB(s) = 0 Then Exit Function
    CountLines = Len(s) - Len(Replace(s, vbLf, ""))
End Function

Private Function LenInfo(ByVal s As String) As String
    If LenB(s) = 0 Then
        LenInfo = "★空★"
    Else
        LenInfo = CStr(Len(s)) & " 文字"
    End If
End Function

Private Function PadR(ByVal s As String, ByVal n As Long) As String
    If Len(s) >= n Then PadR = s Else PadR = s & Space$(n - Len(s))
End Function
