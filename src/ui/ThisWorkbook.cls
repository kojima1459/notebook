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
' ThisWorkbook - Workbook_Open/Workbook_BeforeCloseの薄い転送のみ(MASTER_SPEC §7.6)
' ----------------------------------------------------------------------------
' 役割:
'   Workbook_Open → modBoot.Boot / Workbook_BeforeClose → modBoot.Auto_Close
'   への転送だけを行う。ロジックは一切ここに置かない。
'
' 設計判断:
'   ・R35で本番のThisWorkbookはWorkbook_Openを持たなくなった(§2-2)
'     ため、実行時はmodBoot.Auto_Open / modBoot.Auto_Closeが本命という
'     二重化設計になっている(§7.6末尾)。このイベントハンドラは
'     「Workbook_Openが正常に動く環境」での保険的な経路として残す。
'   ・§7の公開契約(vba_lint CONTRACT)により、ThisWorkbookはPublicメンバー
'     0件が正(closed契約)。Workbook_Open/Workbook_BeforeCloseは必ず
'     Privateのイベントハンドラとして書く。
'   ・例外はここで握りつぶさず、modBoot.Boot側の Failed ラベルが
'     modLog.LogErrorへ必ず記録する(R5)。
' ============================================================================

Private Sub Workbook_Open()
    modBoot.Boot
End Sub

' R20-1f(実機第7報⑦): 窓の大きさが変わったら、今出ている画面を組み直す。
' このアプリの幾何は全て「今の可視幅・可視高」から決まるのに、再計算の
' 機会が「画面を開いたとき」だけだったため、開いた後に窓を広げると
' 右と下に空白が残り続けていた。実体(デバウンス・再入抑止・OnTimeの
' 後始末)は modViewport 側。ここは転送だけで、失敗しても握りつぶす
' (イベントで例外を出すとExcelがイベントを止めてしまう)。
' 本番ビルド(baked)は _BAKED_THISWORKBOOK_BODY をそのまま焼き込むため、
' そちら(build_mybookshelf.py)にも同じ転送がある。
Private Sub Workbook_WindowResize(ByVal Wn As Window)
    On Error Resume Next
    modViewport.OnWindowResized
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' R13-4d: 取込中の×クリックは、確認してからにする(いいえ=閉じない)。
    ' 本番のThisWorkbookはこのイベントを持たない設計(§2-2)のため、
    ' これは開発構成でだけ効く経路。
    ' 終了ボタン(modApp.OnSaveAndExit)側にも同じ確認が入っている。
    If Not modUiLock.ConfirmCloseDuringIngest() Then
        Cancel = True
        Exit Sub
    End If
    modBoot.Auto_Close
End Sub
