Attribute VB_Name = "modGacha"
Option Explicit

' ============================================================================
' modGacha - 今日のワンポイント ポップアップ(ランタイムUserForm生成)
' ----------------------------------------------------------------------------
' 役割:
'   my_knowledgeからランダムに1チャンクを選び、ボーダーレスのポップアップ
'   UserFormで表示する。ビルドシステムが.frm非対応のため、VBComponents.Add
'   でランタイム生成→Show→削除のパターンを採用。
'
' 設計判断:
'   ・BorderStyle=0(fmBorderStyleNone)でフレームレス表示
'   ・StartUpPosition=2(画面中央)
'   ・背景: 丸角風(BackColor白) + 上部にグラデーション風カラーバー(Label)
'   ・閉じる: ×ボタン(Label Click) または Escキー(KeyDown)
'   ・モーダル表示(Show True)でフォーカスを奪う
'   ・生成したVBComponentは表示後に必ずRemove(ゴミ残存防止)
' ============================================================================

Private Const FORM_NAME As String = "frmGachaDyn"
Private Const FORM_W As Long = 380
Private Const FORM_H As Long = 260

' ----------------------------------------------------------------------------
' ShowGacha - ワンポイントポップアップを表示
' ----------------------------------------------------------------------------
Public Sub ShowGacha()
    On Error GoTo Fail

    ' ランダムナレッジ取得
    Dim tip As String
    tip = GetRandomKnowledge()
    If LenB(tip) = 0 Then
        modSkin.ShowToast "まだナレッジが登録されていません。", "info"
        Exit Sub
    End If

    ' 既存の同名フォームを削除(前回のゴミ対策)
    CleanupForm

    ' UserForm生成
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    Dim frm As Object
    Set frm = vbProj.VBComponents.Add(3)  ' vbext_ct_MSForm
    frm.Name = FORM_NAME

    ' フォームプロパティ
    With frm.Properties
        .Item("Width") = FORM_W
        .Item("Height") = FORM_H
        .Item("BorderStyle") = 0       ' fmBorderStyleNone
        .Item("BackColor") = &HFFFFFF  ' 白
        .Item("StartUpPosition") = 2   ' 画面中央
        .Item("Caption") = ""
    End With

    ' コードビハインドにイベントハンドラを注入
    Dim codeMod As Object
    Set codeMod = frm.CodeModule
    codeMod.AddFromString BuildFormCode()

    ' コントロール追加
    AddFormControls frm

    ' 表示(モーダル)
    Dim frmInstance As Object
    Set frmInstance = VBA.UserForms.Add(FORM_NAME)

    ' タグにナレッジテキストをセット(フォーム側で読む)
    frmInstance.Tag = tip

    frmInstance.Show True

    ' 表示後にフォーム削除
    CleanupForm
    Exit Sub

Fail:
    CleanupForm
    ' フォールバック: Toastで表示
    On Error Resume Next
    modSkin.ShowToast ChrW(&HD83D) & ChrW(&HDCA1) & " " & tip, "info"
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' コントロール追加
' ----------------------------------------------------------------------------
Private Sub AddFormControls(ByVal frm As Object)
    Dim designer As Object
    Set designer = frm.Designer

    ' 上部カラーバー(グラデーション風: 紫→青)
    Dim bar As Object
    Set bar = designer.Controls.Add("Forms.Label.1", "lblBar")
    With bar
        .Left = 0: .Top = 0
        .Width = FORM_W: .Height = 6
        .BackColor = &HED8BA7   ' #A78BFA (BGR)
        .BackStyle = 1
    End With

    ' タイトル
    Dim ttl As Object
    Set ttl = designer.Controls.Add("Forms.Label.1", "lblTitle")
    With ttl
        .Left = 20: .Top = 16
        .Width = FORM_W - 60: .Height = 20
        .Caption = ChrW(&HD83C) & ChrW(&HDFB2) & " 今日のワンポイント"
        .Font.Size = 12
        .Font.Bold = True
        .ForeColor = &H7C3AED   ' 紫
        .BackStyle = 0
    End With

    ' 閉じるボタン(×)
    Dim btnX As Object
    Set btnX = designer.Controls.Add("Forms.Label.1", "lblClose")
    With btnX
        .Left = FORM_W - 36: .Top = 12
        .Width = 24: .Height = 24
        .Caption = ChrW(&HD83D) & ChrW(&HDD19)  ' ×
        .Font.Size = 14
        .ForeColor = &H808080
        .BackStyle = 0
        .TextAlign = 2  ' fmTextAlignCenter
    End With

    ' 本文(ナレッジ内容)
    Dim body As Object
    Set body = designer.Controls.Add("Forms.Label.1", "lblBody")
    With body
        .Left = 20: .Top = 48
        .Width = FORM_W - 40: .Height = FORM_H - 100
        .Caption = "(読み込み中...)"
        .Font.Size = 10
        .ForeColor = &H3B291E   ' #1E293B (BGR)
        .WordWrap = True
        .BackStyle = 0
    End With

    ' フッター(出典)
    Dim ftr As Object
    Set ftr = designer.Controls.Add("Forms.Label.1", "lblFooter")
    With ftr
        .Left = 20: .Top = FORM_H - 36
        .Width = FORM_W - 40: .Height = 16
        .Caption = ""
        .Font.Size = 8
        .ForeColor = &HB4A394   ' #94A3B8 (BGR)
        .BackStyle = 0
    End With
End Sub

' ----------------------------------------------------------------------------
' フォームのコードビハインド(イベントハンドラ)
' ----------------------------------------------------------------------------
Private Function BuildFormCode() As String
    Dim sb As String
    sb = "Private Sub UserForm_Initialize()" & vbLf
    sb = sb & "    Me.lblBody.Caption = Me.Tag" & vbLf
    sb = sb & "    Me.lblFooter.Caption = ChrW(&HD83D) & ChrW(&HDCD6) & ""  ナレッジ倉庫からランダム抽出""" & vbLf
    sb = sb & "End Sub" & vbLf & vbLf
    sb = sb & "Private Sub lblClose_Click()" & vbLf
    sb = sb & "    Unload Me" & vbLf
    sb = sb & "End Sub" & vbLf & vbLf
    sb = sb & "Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" & vbLf
    sb = sb & "    If KeyCode = 27 Then Unload Me" & vbLf
    sb = sb & "End Sub" & vbLf & vbLf
    sb = sb & "Private Sub UserForm_Click()" & vbLf
    sb = sb & "    ' 背景クリックでも閉じる(モバイル風UX)" & vbLf
    sb = sb & "End Sub" & vbLf
    BuildFormCode = sb
End Function

' ----------------------------------------------------------------------------
' ランダムナレッジ取得
' ----------------------------------------------------------------------------
Private Function GetRandomKnowledge() As String
    On Error GoTo Fail
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("my_knowledge")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        GetRandomKnowledge = ""
        Exit Function
    End If

    ' ランダム行選択(ヘッダー除く)
    Randomize
    Dim rndRow As Long
    rndRow = Int(Rnd() * (lastRow - 1)) + 2

    ' full_text列(7列目)を取得、200字に截断
    Dim txt As String
    txt = CStr(ws.Cells(rndRow, 7).Value)
    If LenB(txt) = 0 Then
        ' フォールバック: summary列(5列目)
        txt = CStr(ws.Cells(rndRow, 5).Value)
    End If

    If Len(txt) > 200 Then txt = Left$(txt, 200) & "..."
    GetRandomKnowledge = txt
    Exit Function

Fail:
    GetRandomKnowledge = ""
End Function

' ----------------------------------------------------------------------------
' フォーム削除(ゴミ残存防止)
' ----------------------------------------------------------------------------
Private Sub CleanupForm()
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove _
        ThisWorkbook.VBProject.VBComponents(FORM_NAME)
    On Error GoTo 0
End Sub
