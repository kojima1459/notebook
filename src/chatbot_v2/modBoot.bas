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

Public Sub Boot()
    On Error GoTo Failed
    modConfig.EnsureLoaded
    modUserProfile.EnsureFirstRun

    If modKnowledgeBase.RowCount() = 0 Then
        MsgBox "knowledge_base シートが空です。前処理データを取り込んでから再度開いてください。", _
               vbExclamation, "InternalNotebookLM v2"
        gReady = False
        Exit Sub
    End If

    gReady = True
    modChatUI.EnsureLayout
    Exit Sub

Failed:
    MsgBox "起動エラー: " & Err.Description, vbCritical, "InternalNotebookLM v2"
    gReady = False
End Sub
