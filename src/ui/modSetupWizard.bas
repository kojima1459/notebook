Attribute VB_Name = "modSetupWizard"
Option Explicit

' ============================================================================
' modSetupWizard - 部門設定UI + 初回セットアップウィザード(R20-4・実機第7報③④)
' ----------------------------------------------------------------------------
' 背景:
'   「みんなの節約: 未設定」「部門: 未読込」は同じ根(nexus_share_path 空/
'   user_department 空)。かつ旧案内は「config シートを開け」という、タブ
'   常時非表示のこのアプリでは実行不能な手順を指示していた。実際に動く設定
'   UI(部門はここ・共有フォルダはmodHelp.OnShareSetup)へ導線を一本化する。
' 役割:
'   1) OnDeptSetup: Hubプロフィールカード等からの部門設定(常設の再入口)。
'   2) RunFirstRunWizard: 初回のみ(modTour.StartTourIfFirstRunから1回だけ)。
'      部門→共有フォルダの2ステップ、いずれもスキップ可。
' 防衛設計(modHelp/modTourと同型):
'   1. 全Publicエントリの最上部でOn Error Resume Next(失敗をUIに留める)。
'   2. クリックハンドラはmodUiLock.Enter/Leaveで多重発火を防ぐ。
'   3. modBoot/modAsk等の凍結モジュールは一切触らない(層規約R1: src/ui/配下)。
' ============================================================================

Private Const DEPT_KEY As String = "user_department"

' ----------------------------------------------------------------------------
' OnDeptSetup - 部門設定(Hubプロフィールカードのクリック導線・4c常設入口)。
'   選択肢: 1=商品部 / 2=リスコン部 / 0=その他(自由入力)。数値以外・Cancel
'   は何もしない(現状維持。UNCの直打ち等は不要な単純な選択式のため)。
' ----------------------------------------------------------------------------
Public Sub OnDeptSetup()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done
    Dim dept As String
    dept = PromptDept()
    If LenB(dept) = 0 Then GoTo Done
    modConfig.SetValue DEPT_KEY, dept
    On Error Resume Next
    modSkin.ShowToast "部門を設定しました: " & dept, "success"
    modHub.EnsureHubLayout
    On Error GoTo Done
Done:
    modUiLock.Leave
End Sub

' 部門選択の問答(InputBox)。戻り値は確定した部門名(未確定/Cancelは空)。
' 数値解釈だけを分離したDeptFromSelectorが純ロジック(goldens対象)。
Private Function PromptDept() As String
    Dim sel As String
    sel = InputBox( _
        "部門を選んでください。番号を入力してください。" & vbCrLf & _
        "1 = 商品部" & vbCrLf & "2 = リスコン部" & vbCrLf & _
        "0 = その他(自由入力)", "Nexus Agent - 部門設定")
    If LenB(Trim$(sel)) = 0 Then Exit Function

    Dim mapped As String: mapped = DeptFromSelector(sel)
    If LenB(mapped) > 0 Then
        PromptDept = mapped
        Exit Function
    End If

    If Trim$(sel) = "0" Then
        Dim free As String
        free = Trim$(InputBox("部門名を入力してください。", "Nexus Agent - 部門設定"))
        If LenB(free) > 0 Then PromptDept = free
    End If
End Function

' DeptFromSelector - 選択肢の数値文字列→部門名(純関数・ゴールデン対象)。
'   "1"=商品部 / "2"=リスコン部。"0"(自由入力)・その他・空は空文字を返し、
'   呼び出し側(PromptDept)が個別に扱う。前後空白は無視する。
Public Function DeptFromSelector(ByVal sel As String) As String
    Select Case Trim$(sel)
        Case "1": DeptFromSelector = "商品部"
        Case "2": DeptFromSelector = "リスコン部"
        Case Else: DeptFromSelector = ""
    End Select
End Function

' ----------------------------------------------------------------------------
' RunFirstRunWizard - 初回のみ(modTour.StartTourIfFirstRunから呼ぶ想定)。
'   (1)部門 (2)共有フォルダ、いずれもスキップ可。ツアー本編(質問の書き方等
'   3ステップ)より前に完結する別の問答で、ツアー本編自体は変更しない。
'   modBootは凍結のため触らず、こちら側から差し込む(仕様書4d)。
' ----------------------------------------------------------------------------
Public Sub RunFirstRunWizard()
    On Error Resume Next
    Step1Dept
    Step2ShareFolder
    On Error GoTo 0
End Sub

Private Sub Step1Dept()
    On Error Resume Next
    If MsgBox("部門を設定しますか?(あとからでも設定できます)", _
              vbYesNo + vbQuestion, "Nexus Agent - はじめに(1/2)") = vbYes Then
        OnDeptSetup
    End If
    On Error GoTo 0
End Sub

Private Sub Step2ShareFolder()
    On Error Resume Next
    If MsgBox("共有フォルダを設定しますか?(みんなの節約・部門資料が使えます)" & vbCrLf & _
              "あとからHubの案内カードやヘルプ" & ChrW(&H2699) & "でも設定できます。", _
              vbYesNo + vbQuestion, "Nexus Agent - はじめに(2/2)") = vbYes Then
        modHelp.OnShareSetup
    End If
    On Error GoTo 0
End Sub
