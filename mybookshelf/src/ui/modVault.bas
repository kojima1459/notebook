Attribute VB_Name = "modVault"
Option Explicit

' ============================================================================
' modVault - ナレッジ登録フォーム(DOCS_NEXUS_SPEC Phase 2・裁定①)
' ----------------------------------------------------------------------------
' UserForm代替: 「登録フォーム用にデザインされた専用シート」をSPAの画面遷移
' としてアクティブ化する方式(Shape上のテキスト入力はフォーカス制御が不安定な
' ため、入力欄はセルで作る)。登録は既存の取込パイプライン(modShelf.IngestFile)
' を再利用: 入力内容を%TEMP%のUTF-8テキストに書き出して取り込むことで、
' 構造チャンク化・重複排除・バッチ埋め込みまで全て既存の実証済み経路に乗せる。
' ============================================================================

Private Const VAULT_SHEET As String = "VaultInput"
Private Const CELL_TITLE As String = "C6"
Private Const CELL_BODY As String = "C8"
Private Const CELL_TAGS As String = "C18"

' ----------------------------------------------------------------------------
' ShowVaultInput - 登録フォームを描画して表示(SPA遷移)
' ----------------------------------------------------------------------------
Public Sub ShowVaultInput()
    Dim ws As Worksheet
    Set ws = GetOrCreateVaultSheet()
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' 冪等再構築
    RemoveVaultShapes ws
    ws.Cells.Clear
    ws.Cells.Font.Name = "Yu Gothic UI"
    ws.Cells.Interior.Color = RGB(243, 244, 246)

    ws.Columns("A").ColumnWidth = 4
    ws.Columns("B").ColumnWidth = 3
    ws.Columns("C:H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 3

    ' カード風の背景(白・角丸)
    Dim card As Shape
    Set card = ws.Shapes.AddShape(5, 20, 16, 560, 420)
    card.Name = "nxv_card"
    card.Adjustments(1) = 0.04
    card.Fill.ForeColor.RGB = RGB(255, 255, 255)
    card.Line.ForeColor.RGB = RGB(229, 231, 235)
    card.Line.Weight = 0.75
    card.Shadow.Visible = 0
    ' カードを最背面へ(入力セルより奥)
    card.ZOrder 1   ' msoSendToBack

    ' タイトル
    With ws.Range("C2:H2")
        .Merge
        .Value = ChrW(&H2795) & " 新規ナレッジの登録"
        .Font.Size = 14
        .Font.Bold = True
    End With
    With ws.Range("C3:H3")
        .Merge
        .Value = "登録した内容はベクトル化され、社内ナレッジ検索の回答に使われます。"
        .Font.Size = 9
        .Font.Color = RGB(107, 114, 128)
    End With

    ' タイトル入力
    With ws.Range("C5:H5")
        .Merge
        .Value = "タイトル(例: 漁船保険の特例について)"
        .Font.Size = 9.5
        .Font.Bold = True
    End With
    With ws.Range("C6:H6")
        .Merge
        .Interior.Color = RGB(249, 250, 251)
        .Borders.LineStyle = 1
        .Borders.Color = RGB(229, 231, 235)
    End With
    ws.Rows(6).RowHeight = 24

    ' 本文入力
    With ws.Range("C7:H7")
        .Merge
        .Value = "ナレッジ本文(またはAIへの修正指示)"
        .Font.Size = 9.5
        .Font.Bold = True
    End With
    With ws.Range("C8:H16")
        .Merge
        .WrapText = True
        .VerticalAlignment = -4160   ' xlTop
        .Interior.Color = RGB(249, 250, 251)
        .Borders.LineStyle = 1
        .Borders.Color = RGB(229, 231, 235)
    End With

    ' タグ入力
    With ws.Range("C17:H17")
        .Merge
        .Value = "タグ(カンマ区切り。例: 約款解釈,特約)"
        .Font.Size = 9.5
        .Font.Bold = True
    End With
    With ws.Range("C18:H18")
        .Merge
        .Interior.Color = RGB(249, 250, 251)
        .Borders.LineStyle = 1
        .Borders.Color = RGB(229, 231, 235)
    End With
    ws.Rows(18).RowHeight = 24

    ' ボタン(登録=青 / キャンセル=白)
    Dim submitBtn As Shape
    Set submitBtn = ws.Shapes.AddShape(5, 420, 388, 140, 32)
    submitBtn.Name = "nxv_submit"
    submitBtn.Line.Visible = 0
    submitBtn.Fill.ForeColor.RGB = RGB(37, 99, 235)
    With submitBtn.TextFrame2
        .TextRange.Text = "ベクトル化して登録"
        .TextRange.Font.Size = 10.5
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    submitBtn.OnAction = "modVault.OnVaultSubmit"

    Dim cancelBtn As Shape
    Set cancelBtn = ws.Shapes.AddShape(5, 320, 388, 90, 32)
    cancelBtn.Name = "nxv_cancel"
    cancelBtn.Fill.ForeColor.RGB = RGB(255, 255, 255)
    cancelBtn.Line.ForeColor.RGB = RGB(229, 231, 235)
    With cancelBtn.TextFrame2
        .TextRange.Text = "キャンセル"
        .TextRange.Font.Size = 10.5
        .TextRange.Font.Fill.ForeColor.RGB = RGB(17, 24, 39)
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    cancelBtn.OnAction = "modVault.OnVaultCancel"

    ' SPA遷移(表示してアクティブ化・枠線等は非表示)
    ws.Visible = -1   ' xlSheetVisible
    ws.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    On Error GoTo 0
    ws.Range(CELL_TITLE).Select

    Application.ScreenUpdating = True
End Sub

' ----------------------------------------------------------------------------
' OnVaultSubmit - 入力内容を既存取込パイプラインへ流して登録
' ----------------------------------------------------------------------------
Public Sub OnVaultSubmit()
    Dim ws As Worksheet
    Set ws = GetVaultSheet()
    If ws Is Nothing Then Exit Sub

    Dim titleText As String, bodyText As String, tagsText As String
    titleText = Trim$(CStr(ws.Range(CELL_TITLE).Value))
    bodyText = Trim$(CStr(ws.Range(CELL_BODY).Value))
    tagsText = Trim$(CStr(ws.Range(CELL_TAGS).Value))

    If LenB(titleText) = 0 Or LenB(bodyText) = 0 Then
        MsgBox "タイトルと本文を入力してください。", vbExclamation, "Nexus Agent"
        Exit Sub
    End If

    If RegisterKnowledgeText(titleText, TagsHeader(tagsText) & bodyText, tagsText) Then
        MsgBox "ナレッジデータベースに追加され、ベクトル化されました。", vbInformation, "Nexus Agent"
        ClearInputs ws
        CloseVault ws
    Else
        MsgBox "登録に失敗しました。マイ本棚の一覧で状態をご確認ください。", vbExclamation, "Nexus Agent"
    End If
End Sub

Public Sub OnVaultCancel()
    Dim ws As Worksheet
    Set ws = GetVaultSheet()
    If ws Is Nothing Then Exit Sub
    ClearInputs ws
    CloseVault ws
End Sub

' ----------------------------------------------------------------------------
' RegisterKnowledgeText - テキストナレッジを既存パイプラインで登録する共通口
'   (modAppの👎自己学習からも使う)。成功=True(done/partial)。
' ----------------------------------------------------------------------------
Public Function RegisterKnowledgeText(ByVal titleText As String, ByVal bodyText As String, _
                                      ByVal tagsText As String) As Boolean
    On Error GoTo Fail

    Dim tempDir As String: tempDir = Environ$("TEMP")
    If LenB(tempDir) = 0 Then tempDir = Environ$("TMP")
    If LenB(tempDir) = 0 Then Exit Function
    If Right$(tempDir, 1) <> "\" Then tempDir = tempDir & "\"

    Dim filePath As String
    filePath = tempDir & "ナレッジ_" & SanitizeName(titleText) & "_" & _
               Format$(Now, "yyyymmddhhnnss") & ".txt"

    Dim content As String
    content = "【" & titleText & "】" & vbLf & bodyText

    ' UTF-8で保存(modExtractorのtxt読取りがUTF-8のため整合)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2          ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText content
    st.SaveToFile filePath, 2   ' adSaveCreateOverWrite
    st.Close
    Set st = Nothing

    Dim resultStatus As String
    resultStatus = modShelf.IngestFile(filePath, "self")

    ' 一時ファイルは掃除(取込済みなので不要。失敗しても無視)
    On Error Resume Next
    Kill filePath
    On Error GoTo 0

    RegisterKnowledgeText = (resultStatus = "done" Or resultStatus = "partial")
    Exit Function

Fail:
    On Error Resume Next
    modLog.LogError "E0801", "modVault.RegisterKnowledgeText", Err.Description
    On Error GoTo 0
    RegisterKnowledgeText = False
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

Private Function TagsHeader(ByVal tagsText As String) As String
    If LenB(tagsText) = 0 Then Exit Function
    TagsHeader = "タグ: " & tagsText & vbLf & vbLf
End Function

Private Sub ClearInputs(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Range(CELL_TITLE).Value = ""
    ws.Range(CELL_BODY).Value = ""
    ws.Range(CELL_TAGS).Value = ""
    On Error GoTo 0
End Sub

Private Sub CloseVault(ByVal ws As Worksheet)
    On Error Resume Next
    ThisWorkbook.Worksheets("Nexus").Activate
    On Error GoTo 0
    On Error Resume Next
    ws.Visible = 2   ' xlSheetVeryHidden
    On Error GoTo 0
End Sub

Private Function SanitizeName(ByVal s As String) As String
    Dim bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)
    Dim t As String: t = s
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        t = Replace(t, CStr(bad(i)), "_")
    Next i
    SanitizeName = modUtil.SafeLeft(Trim$(t), 60)
End Function

Private Function GetVaultSheet() As Worksheet
    On Error Resume Next
    Set GetVaultSheet = ThisWorkbook.Worksheets(VAULT_SHEET)
    On Error GoTo 0
End Function

Private Function GetOrCreateVaultSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = GetVaultSheet()
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = VAULT_SHEET
        On Error GoTo 0
    End If
    Set GetOrCreateVaultSheet = ws
    Exit Function
Fail:
    Set GetOrCreateVaultSheet = Nothing
End Function

Private Sub RemoveVaultShapes(ByVal ws As Worksheet)
    Dim names() As String
    ReDim names(0 To ws.Shapes.count)
    Dim n As Long: n = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 4) = "nxv_" Then
            names(n) = shp.Name
            n = n + 1
        End If
    Next shp
    Dim i As Long
    For i = 0 To n - 1
        On Error Resume Next
        ws.Shapes(names(i)).Delete
        On Error GoTo 0
    Next i
End Sub
