VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ============================================================================
' ThisWorkbook - autorun + self-installer
' ----------------------------------------------------------------------------
' Why this is unusual:
' Some Excel versions silently drop standard modules from a programmatically-
' built vbaProject.bin even when the dir stream is valid. To survive that, we
' carry every standard module's source code on a hidden `vba_src` sheet, and
' if Workbook_Open finds modBoot.Boot unreachable, it rebuilds the standard
' modules from that sheet via the VBProject API. The user just clicks
' "コンテンツの有効化".
'
' Requires the user (or IT) to have ticked
'   トラストセンター → マクロの設定 →
'   「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」
' If not, Setup explains exactly what to enable.
' ============================================================================

Private Sub Workbook_Open()
    ' Try the normal boot path first (works on the second and later opens
    ' after Setup has installed the standard modules).
    On Error Resume Next
    Application.Run "modBoot.Boot"
    Dim bootErr As Long: bootErr = Err.Number
    On Error GoTo 0

    If bootErr = 0 Then Exit Sub

    ' modBoot is missing -> install all standard modules from vba_src.
    ' This is the first-run path. After Setup saves the workbook, every
    ' subsequent open hits the bootErr=0 branch above.
    Setup
End Sub

Public Sub Setup()
    Dim vbProj As Object
    On Error GoTo NoTrust
    Set vbProj = ThisWorkbook.VBProject
    On Error GoTo Failed

    Dim wsSrc As Worksheet
    On Error Resume Next
    Set wsSrc = ThisWorkbook.Worksheets("vba_src")
    On Error GoTo Failed
    If wsSrc Is Nothing Then
        MsgBox "vba_src シートが見つかりません。配布ファイルが壊れている可能性があります。", _
               vbCritical, "セットアップ"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim lastRow As Long
    lastRow = wsSrc.Cells(wsSrc.Rows.count, 1).End(xlUp).row
    If lastRow < 2 Then
        MsgBox "vba_src にソースが入っていません。", vbCritical, "セットアップ"
        GoTo Cleanup
    End If

    Dim installed As Long, replaced As Long, r As Long
    For r = 2 To lastRow
        Dim modName As String: modName = CStr(wsSrc.Cells(r, 1).value)
        Dim modSrc As String:  modSrc = CStr(wsSrc.Cells(r, 3).value)
        If LenB(modName) = 0 Then GoTo NextRow

        ' Remove existing module of the same name (if any), then re-create.
        On Error Resume Next
        Dim existing As Object: Set existing = Nothing
        Set existing = vbProj.VBComponents(modName)
        If Not existing Is Nothing Then
            vbProj.VBComponents.Remove existing
            replaced = replaced + 1
        End If
        Err.Clear
        On Error GoTo Failed

        ' Add a standard module (vbext_ct_StdModule = 1) and load source.
        Dim vbComp As Object
        Set vbComp = vbProj.VBComponents.Add(1)
        vbComp.Name = modName
        If LenB(modSrc) > 0 Then
            vbComp.CodeModule.AddFromString modSrc
        End If
        installed = installed + 1
NextRow:
    Next r

Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Save so modules persist across closes. If the file is read-only or on a
    ' locked share, Save will fail -- don't abort; modules still work this session.
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo Failed

    MsgBox "セットアップ完了" & vbCrLf & _
           "  ・新規追加: " & installed & " モジュール" & vbCrLf & _
           "  ・置き換え: " & replaced & " モジュール" & vbCrLf & vbCrLf & _
           "今すぐボットを起動します。", vbInformation, "セットアップ"

    ' Now that modules exist, run the real Boot.
    On Error Resume Next
    Application.Run "modBoot.Boot"
    On Error GoTo 0
    Exit Sub

NoTrust:
    MsgBox "VBAプロジェクトへのアクセスが許可されていません。" & vbCrLf & vbCrLf & _
           "【設定手順】" & vbCrLf & _
           "1. Excelの [ファイル] → [オプション]" & vbCrLf & _
           "2. [トラストセンター] → [トラストセンターの設定]" & vbCrLf & _
           "3. [マクロの設定]" & vbCrLf & _
           "4. 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェック" & vbCrLf & _
           "5. OKを押してExcelを再起動し、ファイルを再度開く" & vbCrLf & vbCrLf & _
           "再度開いたら自動でセットアップが実行されます。" & vbCrLf & _
           "（または main シートの B3 のオレンジのセルをクリックしてください）", _
           vbCritical, "セットアップ"
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "セットアップ中にエラーが発生しました。" & vbCrLf & _
           "Err " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "原因の可能性:" & vbCrLf & _
           "  ・トラストセンターで「VBAプロジェクトへのアクセスを信頼する」が未設定" & vbCrLf & _
           "  ・ファイルが読み取り専用 (右クリック→プロパティ→読み取り専用のチェック解除)" & vbCrLf & _
           "  ・vba_src シートが破損", vbCritical, "セットアップ"
End Sub
