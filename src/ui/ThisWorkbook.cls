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
'   ・ビルド版ではインストーラ(vba_src自己インストーラ)がThisWorkbookを
'     占有するため、実行時はmodBoot.Auto_Open / modBoot.Auto_Closeが本命
'     という二重化設計になっている(§7.6末尾)。このイベントハンドラは
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

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    modBoot.Auto_Close
End Sub
