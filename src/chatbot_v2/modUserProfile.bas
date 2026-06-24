Attribute VB_Name = "modUserProfile"
Option Explicit

' ============================================================================
' modUserProfile - Per-user dept and identity
' ----------------------------------------------------------------------------
' v2 saves user profile into the config sheet so distribution is just the
' single xlsm. First-run prompts:
'   - department: pick from department sheet
'   - role: optional (アンダーライター, 営業, 鑑定人, etc.)
' Saved to config rows: department_id, role
' ============================================================================

Public Sub EnsureFirstRun()
    Dim deptId As String: deptId = modConfig.GetString("department_id", "")
    If LenB(deptId) > 0 Then Exit Sub
    PromptDepartment
End Sub

Public Sub PromptDepartment()
    Dim depts As Collection
    Set depts = LoadDeptList()
    If depts.count = 0 Then
        MsgBox "department シートが空です。先に部署マスタを設定してください。", _
               vbExclamation
        Exit Sub
    End If

    Dim choices As String, i As Long
    For i = 1 To depts.count
        choices = choices & i & " : " & depts(i) & vbLf
    Next i

    Dim picked As Variant
    picked = Application.InputBox( _
        Prompt:="所属部署を番号で選択してください:" & vbLf & vbLf & choices, _
        Title:="初回設定 - 部署選択", _
        Type:=1)

    If VarType(picked) = vbBoolean And picked = False Then Exit Sub
    Dim n As Long: n = CLng(picked)
    If n < 1 Or n > depts.count Then
        MsgBox "番号が範囲外です。", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("department")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "department シートが見つかりません。配布ファイルが壊れている可能性があります。", _
               vbCritical, "InternalNotebookLM v2"
        Exit Sub
    End If
    Dim deptRow As Long: deptRow = n + 1   ' +1 for header
    modConfig.SetValue "department_id", CStr(ws.Cells(deptRow, 1).value)
    modConfig.SetValue "department_name", CStr(ws.Cells(deptRow, 2).value)

    Dim role As String
    role = InputBox("役割を入力してください (例: アンダーライター、営業)", "初回設定 - 役割")
    If LenB(role) > 0 Then modConfig.SetValue "role", role

    MsgBox "初回設定が完了しました。" & vbCrLf & vbCrLf & _
           "部署: " & depts(n) & vbCrLf & _
           "役割: " & role, vbInformation
End Sub

Private Function LoadDeptList() As Collection
    Dim col As New Collection
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("department")
    On Error GoTo 0
    If ws Is Nothing Then
        Set LoadDeptList = col
        Exit Function
    End If
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim i As Long
    For i = 2 To lastRow
        Dim deptName As String: deptName = CStr(ws.Cells(i, 2).value)
        If LenB(deptName) > 0 Then col.Add deptName
    Next i
    Set LoadDeptList = col
End Function

Public Function CurrentDept() As String
    CurrentDept = modConfig.GetString("department_id", "")
End Function

Public Function CurrentDeptName() As String
    CurrentDeptName = modConfig.GetString("department_name", "")
End Function

Public Function CurrentRole() As String
    CurrentRole = modConfig.GetString("role", "")
End Function

Public Function CurrentUser() As String
    CurrentUser = Environ$("USERNAME")
End Function
