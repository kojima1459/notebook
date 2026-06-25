VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Cell B3 on the main sheet is the "setup trigger".
' Clicking it calls modBoot.Boot if modules are installed, or
' ThisWorkbook.Setup if they are not yet. This handler lives in the
' vbaProject.bin Document module so it fires even before standard
' modules are installed -- unlike vba_src modules which need Setup() first.
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    If Not Intersect(Target, Me.Range("B3")) Is Nothing Then
        Application.Run "modBoot.Boot"
        If Err.Number <> 0 Then
            Err.Clear
            ThisWorkbook.Setup
        End If
    End If
    On Error GoTo 0
End Sub
