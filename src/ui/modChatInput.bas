Attribute VB_Name = "modChatInput"
Option Explicit

' ============================================================================
' modChatInput - チャット入力欄(ボーダーレスUserForm ランタイム生成)
' ----------------------------------------------------------------------------
' 役割:
'   Chat画面の下部に固定表示される入力欄。MultiLine TextBox + 送信ボタン。
'   Ctrl+Enterで送信、Enterで改行。ビルドシステムが.frm非対応のため
'   VBComponents.Addでランタイム生成する。
'
' 設計判断:
'   ・BorderStyle=0(fmBorderStyleNone) + 白背景 + 1px相当の境界線(Labelで模倣)
'   ・StartUpPosition=0(Manual) → Chat画面の下部に配置
'   ・モードレス表示(Show False)でChat画面の操作を妨げない
'   ・送信時は modAsk.Answer を呼び出し、結果をChat画面にバブル表示
'   ・フォームは1度生成したら再利用(UnloadせずHide)
'   ・nx_input名前定義セルとの双方向同期(テンプレ流し込み対応)
' ============================================================================

Private Const FORM_NAME As String = "frmChatInputDyn"
Private Const FORM_W As Long = 560
Private Const FORM_H As Long = 80
Private mFormShown As Boolean

' ----------------------------------------------------------------------------
' ShowChatInput - チャット入力欄を表示(冪等)
' ----------------------------------------------------------------------------
Public Sub ShowChatInput()
    On Error GoTo Fail

    ' 既に存在すれば再利用
    If FormExists() Then
        On Error Resume Next
        VBA.UserForms.Add(FORM_NAME).Show False
        On Error GoTo 0
        mFormShown = True
        Exit Sub
    End If

    ' 新規生成
    CleanupForm
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    Dim frm As Object
    Set frm = vbProj.VBComponents.Add(3)  ' vbext_ct_MSForm
    frm.Name = FORM_NAME

    With frm.Properties
        .Item("Width") = FORM_W
        .Item("Height") = FORM_H
        .Item("BorderStyle") = 0
        .Item("BackColor") = &HFFFFFF
        .Item("StartUpPosition") = 0  ' Manual
        .Item("Caption") = ""
    End With

    ' コード注入
    frm.CodeModule.AddFromString BuildInputFormCode()

    ' コントロール
    AddInputControls frm

    ' 表示(モードレス)
    Dim inst As Object
    Set inst = VBA.UserForms.Add(FORM_NAME)

    ' 位置: 画面下部中央
    With inst
        .Left = (Application.Width - FORM_W) / 2
        .Top = Application.Height - FORM_H - 80
    End With
    inst.Show False
    mFormShown = True
    Exit Sub

Fail:
    ' フォールバック: nx_inputセルをアクティブ化
    On Error Resume Next
    ThisWorkbook.Names("nx_input").RefersToRange.Select
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' HideChatInput - 入力欄を非表示
' ----------------------------------------------------------------------------
Public Sub HideChatInput()
    On Error Resume Next
    Dim frm As Object
    Set frm = VBA.UserForms.Add(FORM_NAME)
    frm.Hide
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' コントロール追加
' ----------------------------------------------------------------------------
Private Sub AddInputControls(ByVal frm As Object)
    Dim d As Object
    Set d = frm.Designer

    ' 上部ボーダーライン(1px風)
    Dim border As Object
    Set border = d.Controls.Add("Forms.Label.1", "lblBorder")
    With border
        .Left = 0: .Top = 0
        .Width = FORM_W: .Height = 1
        .BackColor = &HF0ECE2   ' #E2E8F0 (BGR)
        .BackStyle = 1
    End With

    ' テキストボックス(MultiLine)
    Dim txt As Object
    Set txt = d.Controls.Add("Forms.TextBox.1", "txtInput")
    With txt
        .Left = 12: .Top = 10
        .Width = FORM_W - 80: .Height = FORM_H - 24
        .MultiLine = True
        .WordWrap = True
        .ScrollBars = 2  ' fmScrollBarsVertical
        .Font.Size = 11
        .Font.Name = "Yu Gothic UI"
        .BorderStyle = 0  ' fmBorderStyleNone
        .BackColor = &HFFFFFF
    End With

    ' 送信ボタン
    Dim btn As Object
    Set btn = d.Controls.Add("Forms.CommandButton.1", "btnSend")
    With btn
        .Left = FORM_W - 60: .Top = 16
        .Width = 48: .Height = 48
        .Caption = ChrW(&HD83D) & ChrW(&HDCE4)  ' 📤
        .Font.Size = 16
        .BackColor = &HD8991F   ' #1F4E78 (BGR)
        .ForeColor = &HFFFFFF
    End With

    ' ヒントテキスト(下部)
    Dim hint As Object
    Set hint = d.Controls.Add("Forms.Label.1", "lblHint")
    With hint
        .Left = 12: .Top = FORM_H - 14
        .Width = 300: .Height = 12
        .Caption = "Ctrl+Enter で送信 / Enter で改行"
        .Font.Size = 7.5
        .ForeColor = &HB4A394   ' #94A3B8
        .BackStyle = 0
    End With
End Sub

' ----------------------------------------------------------------------------
' フォームコード(イベントハンドラ)
' ----------------------------------------------------------------------------
Private Function BuildInputFormCode() As String
    Dim sb As String
    sb = "Private Sub btnSend_Click()" & vbLf
    sb = sb & "    Call DoSend" & vbLf
    sb = sb & "End Sub" & vbLf & vbLf

    sb = sb & "Private Sub txtInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" & vbLf
    sb = sb & "    If KeyCode = 13 And Shift = 2 Then" & vbLf  ' Enter + Ctrl
    sb = sb & "        KeyCode = 0" & vbLf
    sb = sb & "        Call DoSend" & vbLf
    sb = sb & "    End If" & vbLf
    sb = sb & "End Sub" & vbLf & vbLf

    sb = sb & "Private Sub DoSend()" & vbLf
    sb = sb & "    Dim q As String: q = Trim$(Me.txtInput.Text)" & vbLf
    sb = sb & "    If LenB(q) = 0 Then Exit Sub" & vbLf
    sb = sb & "    Me.txtInput.Text = """" " & vbLf
    sb = sb & "    ' nx_inputセルに同期(既存フローとの互換)" & vbLf
    sb = sb & "    On Error Resume Next" & vbLf
    sb = sb & "    ThisWorkbook.Names(""nx_input"").RefersToRange.Value = q" & vbLf
    sb = sb & "    On Error GoTo 0" & vbLf
    sb = sb & "    ' 回答生成を呼び出し" & vbLf
    sb = sb & "    modApp.OnAsk" & vbLf
    sb = sb & "End Sub" & vbLf & vbLf

    sb = sb & "Private Sub UserForm_Initialize()" & vbLf
    sb = sb & "    ' nx_inputセルにテンプレがあれば流し込み" & vbLf
    sb = sb & "    On Error Resume Next" & vbLf
    sb = sb & "    Dim v As String: v = CStr(ThisWorkbook.Names(""nx_input"").RefersToRange.Value)" & vbLf
    sb = sb & "    If LenB(v) > 0 Then Me.txtInput.Text = v" & vbLf
    sb = sb & "    On Error GoTo 0" & vbLf
    sb = sb & "End Sub" & vbLf

    BuildInputFormCode = sb
End Function

' ----------------------------------------------------------------------------
' ヘルパー
' ----------------------------------------------------------------------------
Private Function FormExists() As Boolean
    On Error Resume Next
    Dim c As Object
    Set c = ThisWorkbook.VBProject.VBComponents(FORM_NAME)
    FormExists = Not c Is Nothing
    On Error GoTo 0
End Function

Private Sub CleanupForm()
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove _
        ThisWorkbook.VBProject.VBComponents(FORM_NAME)
    On Error GoTo 0
End Sub
