Attribute VB_Name = "modUserProfile"
Option Explicit

' ============================================================================
' modUserProfile - persisted per-user attributes (department, role, consent)
' ----------------------------------------------------------------------------
' First-time launch prompts for department/role via Application.InputBox.
' Stored in a simple key=value text file under AppData. Never sent to API
' calls directly; only used for usage CSV.
' ============================================================================

Private Const PROFILE_FILE As String = "profile.dat"

Public Function Department() As String
    Department = ReadField("department")
End Function

Public Function Role() As String
    Role = ReadField("role")
End Function

Public Function ShareQuestionConsent() As Boolean
    ShareQuestionConsent = (ReadField("share_question") = "1")
End Function

Public Sub SetShareQuestionConsent(ByVal v As Boolean)
    WriteField "share_question", IIf(v, "1", "0")
End Sub

Public Sub EnsureFirstRun()
    If LenB(Department()) > 0 Then Exit Sub
    Dim dep As String, role As String
    dep = Trim$(CStr(Application.InputBox( _
        "あなたの部署名を入力してください（例: 第一営業部）", "初期設定 (1/2)", , , , , , 2)))
    If LenB(dep) = 0 Then dep = "未設定"
    role = Trim$(CStr(Application.InputBox( _
        "あなたの職位/役割を入力してください（例: 営業担当）", "初期設定 (2/2)", , , , , , 2)))
    If LenB(role) = 0 Then role = "未設定"
    WriteField "department", dep
    WriteField "role", role
End Sub

Private Function ProfilePath() As String
    ProfilePath = modPaths.AppDataRoot() & "\" & PROFILE_FILE
End Function

Private Function ReadField(ByVal key As String) As String
    Dim path As String: path = ProfilePath()
    If Len(Dir$(path)) = 0 Then Exit Function
    Dim fnum As Integer: fnum = FreeFile
    Open path For Input As #fnum
    Dim line As String, eqPos As Long
    Do While Not EOF(fnum)
        Line Input #fnum, line
        eqPos = InStr(line, "=")
        If eqPos > 0 Then
            If LCase$(Trim$(Left$(line, eqPos - 1))) = LCase$(key) Then
                ReadField = Trim$(Mid$(line, eqPos + 1))
                Close #fnum
                Exit Function
            End If
        End If
    Loop
    Close #fnum
End Function

Private Sub WriteField(ByVal key As String, ByVal value As String)
    modPaths.EnsureDir modPaths.AppDataRoot()
    Dim path As String: path = ProfilePath()
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    If Len(Dir$(path)) > 0 Then
        Dim fin As Integer: fin = FreeFile
        Open path For Input As #fin
        Dim line As String, eqPos As Long
        Do While Not EOF(fin)
            Line Input #fin, line
            eqPos = InStr(line, "=")
            If eqPos > 0 Then
                dict(Trim$(Left$(line, eqPos - 1))) = Trim$(Mid$(line, eqPos + 1))
            End If
        Loop
        Close #fin
    End If
    dict(key) = value

    Dim fout As Integer: fout = FreeFile
    Open path For Output As #fout
    Dim k As Variant
    For Each k In dict.Keys
        Print #fout, CStr(k) & "=" & CStr(dict(k))
    Next k
    Close #fout
End Sub
