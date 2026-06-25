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
    On Error GoTo 0
    Exit Sub

Failed:
    MsgBox "起動エラー: " & Err.Description, vbCritical, "InternalNotebookLM v2"
    gReady = False
End Sub
