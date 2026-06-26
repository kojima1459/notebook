Attribute VB_Name = "modBoot"
Option Explicit

' ============================================================================
' modBoot - Workbook_Open entry point (v2 / 社内AIリボン経由)
' ----------------------------------------------------------------------------
' v2 boot sequence (vs v1):
'   1. Read config from `config` sheet (no external file)
'   2. First-run: ask for department, save to in-workbook hidden cell
'   3. Verify knowledge_base sheet has rows
'   4. Render the main UI
' v2 has NO external dependencies:
'   - No Gemini API key
'   - No external index file
'   - No SharePoint sync (Phase 3)
' ============================================================================

Public gReady As Boolean
Public gLastQuestion As String        ' for feedback handlers
Public gLastAnswer As String          ' for feedback handlers
Public gLastSelectedIds As String     ' chunk IDs used in last answer (comma sep)
Public gFollowupMode As Boolean       ' true when next RunQuery should treat as a follow-up
Public gPrevQ As String               ' previous turn's question (legacy / unused)
Public gPrevA As String               ' previous turn's answer (legacy / unused)
Public gHistory As String             ' rolling Q&A history for multi-turn follow-ups
Public Const HISTORY_MAX_TURNS As Long = 6   ' keep last N exchanges as context
Public Const HISTORY_SEP As String = "<<<__TURN__>>>"
Private gBootDone As Boolean          ' guard: Workbook_Open AND Auto_Open both call Boot

' Auto_Open is called by Excel when the workbook is opened interactively.
' We keep BOTH Workbook_Open (in ThisWorkbook) and Auto_Open so the UI still
' builds if one path is suppressed by policy. The guard prevents double work.
Public Sub Auto_Open()
    Boot
End Sub

Public Sub Boot()
    If gBootDone And gReady Then
        On Error Resume Next
        modChatUI.EnsureLayout      ' already initialized; just re-show the UI
        On Error GoTo 0
        Exit Sub
    End If

    On Error GoTo Failed
    modConfig.EnsureLoaded
    modUserProfile.EnsureFirstRun

    If modKnowledgeBase.RowCount() = 0 Then
        MsgBox "knowledge_base シートが空です。前処理データを取り込んでから再度開いてください。", _
               vbExclamation, "InternalNotebookLM v2"
        gReady = False
        Exit Sub
    End If

    ' Mark ready BEFORE building the UI. If EnsureLayout's cosmetic
    ' Activate/Select fails on first open (Workbook_Open timing), we
    ' don't want it to flip gReady back to False.
    gReady = True
    gBootDone = True
    On Error Resume Next
    modChatUI.EnsureLayout
    HideInternalSheets
    On Error GoTo 0
    Exit Sub

Failed:
    MsgBox "起動エラー: " & Err.Description, vbCritical, "InternalNotebookLM v2"
    gReady = False
End Sub

' ----------------------------------------------------------------------------
' Multi-turn follow-up history helpers
' ----------------------------------------------------------------------------
' Append one exchange (question + answer) to gHistory, keeping only the last
' HISTORY_MAX_TURNS turns. Answers are trimmed so the running prompt stays
' bounded even after many follow-ups.
Public Sub AppendHistory(ByVal q As String, ByVal a As String)
    Dim turn As String
    turn = "Q: " & q & vbLf & "A: " & Left$(a, 2500)

    If LenB(gHistory) = 0 Then
        gHistory = turn
    Else
        gHistory = gHistory & HISTORY_SEP & turn
    End If

    ' Trim to last HISTORY_MAX_TURNS
    Dim parts() As String: parts = Split(gHistory, HISTORY_SEP)
    If UBound(parts) + 1 > HISTORY_MAX_TURNS Then
        Dim keep As String, i As Long
        For i = UBound(parts) - HISTORY_MAX_TURNS + 1 To UBound(parts)
            If LenB(keep) > 0 Then keep = keep & HISTORY_SEP
            keep = keep & parts(i)
        Next i
        gHistory = keep
    End If
End Sub

' Build the conversation block to inject into the drafter prompt for follow-ups.
Public Function HistoryBlock() As String
    If LenB(gHistory) = 0 Then Exit Function
    Dim parts() As String: parts = Split(gHistory, HISTORY_SEP)
    Dim out As String, i As Long, n As Long
    For i = LBound(parts) To UBound(parts)
        n = n + 1
        out = out & "【会話" & n & "】" & vbLf & parts(i) & vbLf & vbLf
    Next i
    HistoryBlock = out
End Function

Public Sub ResetHistory()
    gHistory = ""
End Sub

' ============================================================================
' Keep internal sheets hidden at runtime. xlSheetVeryHidden (=2) prevents users
' from unhiding via Format > Sheet > Unhide — VBA access only.
' This is called every Boot() so manual unhiding during a session is re-hidden
' on next open. Complements the build-time sheet_state settings.
' ============================================================================
Private Sub HideInternalSheets()
    Const VERY_HIDDEN As Long = 2   ' xlSheetVeryHidden
    Const HIDDEN      As Long = 0   ' xlSheetHidden
    Dim protect() As Variant
    ' veryHidden: users must never see or edit these
    protect = Array("system_prompt", "knowledge_base", "manifest", _
                    "feedback", "vba_src")
    Dim sheetName As Variant
    For Each sheetName In protect
        On Error Resume Next
        ThisWorkbook.Worksheets(CStr(sheetName)).Visible = VERY_HIDDEN
        On Error GoTo 0
    Next sheetName
    ' hidden (admin can unhide via Format > Sheet > Unhide if needed)
    Dim adminOnly() As Variant
    adminOnly = Array("config", "usage_log")
    For Each sheetName In adminOnly
        On Error Resume Next
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        If Not ws Is Nothing Then
            If ws.Visible = -1 Then ws.Visible = HIDDEN   ' only hide if currently visible
        End If
        Set ws = Nothing
        On Error GoTo 0
    Next sheetName
End Sub
