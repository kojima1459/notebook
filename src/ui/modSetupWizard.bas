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

' 2026-08-06 R20H FA-9: ウィザード専用の完了フラグ。modTourの
' "nexus_tour_done" はツアー本編(3ステップ)まで含めた完了時にしか立たない
' ため、ウィザード(部門/共有フォルダの2問)には答えた(またはスキップした)の
' に、直後のツアー本編を最後までやり切らずに閉じると、次回起動時に
' StartTourIfFirstRun→RunFirstRunWizard が再度走り、同じ2問が再出現していた
' (実機報告)。ツアーのフラグとは独立に、ウィザード自身の「やった」を覚える。
Private Const WIZARD_DONE_KEY As String = "nexus_wizard_done"

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
        "0 = その他(自由入力)", modAppDef.APP_NAME & " - 部門設定")
    If LenB(Trim$(sel)) = 0 Then Exit Function

    Dim mapped As String: mapped = DeptFromSelector(sel)
    If LenB(mapped) > 0 Then
        PromptDept = mapped
        Exit Function
    End If

    If Trim$(sel) = "0" Then
        Dim free As String
        free = Trim$(InputBox("部門名を入力してください。", modAppDef.APP_NAME & " - 部門設定"))
        If LenB(free) > 0 Then PromptDept = free
        Exit Function
    End If

    ' 2026-08-06 R20H FA-11: ここまで来た(空でも1/2/0でもない)入力は、
    ' 従来は無言で何も起きなかった(無言失敗)。何が起きたかを一言返す。
    On Error Resume Next
    modSkin.ShowToast "1・2・0のいずれかで入力してください。Hubの案内カードからやり直せます", "info"
    On Error GoTo 0
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
'   2026-08-06 R20H FA-9: 冒頭で専用フラグ(nexus_wizard_done)を確認し、
'   既に一度出したなら即撤退。スキップされた場合でも末尾で必ず"1"を保存
'   する(=一度でも問答の機会を提示すれば、答えてもスキップしても二度と
'   出さない。以降の再設定はHubのプロフィールカード=OnDeptSetup/
'   modHelp.OnShareSetupが常設の再入口)。
' ----------------------------------------------------------------------------
Public Sub RunFirstRunWizard()
    On Error Resume Next
    If Not WizardShouldRun(modState.LoadState(WIZARD_DONE_KEY, "")) Then Exit Sub
    Step1Dept
    Step2ShareFolder
    modState.SaveState WIZARD_DONE_KEY, "1"
    ' R35: 方式Bでは Install の Save が無くなり、初回の pack_author/部門/
    ' 共有パス/この印が「保存しない」で失われる。初回ウィザードが走った回
    ' だけここで1回保存する。
    ' 09-03 敵対的レビュー班D BLOCKER(2周目): 読み取り専用(zip直開き・
    ' 共有中の原本を他人が開いている・保護ビュー)ではSaveがダイアログを
    ' 出し、On Error Resume Next はそのダイアログを止めない。ReadOnlyの
    ' ときは無言でスキップする(既存3箇所と同じ作法: modShelfBatch.bas:
    ' 202-203 / modShelfStore.bas:506 / modShelf.bas:451)。マークはSaveの
    ' 直前に書く(失敗してもメモリ上のマークが最新を指すだけで害は無い)。
    If Not ThisWorkbook.ReadOnly Then
        modIntegrity.RecordSaveMark
        ThisWorkbook.Save
    End If
    On Error GoTo 0
End Sub

' WizardShouldRun - 保存済みフラグの値からウィザードをまだ出してよいかを
'   決める純関数(R20H FA-9・ゴールデン対象)。"1"だけが「出さない」で、
'   それ以外(空/未設定/壊れた値)は全て「出す」側に倒す(1回だけ出す契約を
'   フラグの読み取り忘れ・破損の方向に間違えても再出現に倒れないように)。
Public Function WizardShouldRun(ByVal savedFlag As String) As Boolean
    WizardShouldRun = (savedFlag <> "1")
End Function

Private Sub Step1Dept()
    On Error Resume Next
    If MsgBox("部門を設定しますか?(あとからでも設定できます)", _
              vbYesNo + vbQuestion, modAppDef.APP_NAME & " - はじめに(1/2)") = vbYes Then
        OnDeptSetup
    End If
    On Error GoTo 0
End Sub

Private Sub Step2ShareFolder()
    On Error Resume Next
    If MsgBox("共有フォルダを設定しますか?(みんなの節約・部門資料が使えます)" & vbCrLf & _
              "あとからHubの案内カードやヘルプ" & ChrW(&H2699) & "でも設定できます。", _
              vbYesNo + vbQuestion, modAppDef.APP_NAME & " - はじめに(2/2)") = vbYes Then
        modHelp.OnShareSetup
    End If
    On Error GoTo 0
End Sub
