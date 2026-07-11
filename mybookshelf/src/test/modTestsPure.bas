Attribute VB_Name = "modTestsPure"
Option Explicit

Public Sub RunAll()
    Dim r() As String
    r = modUtil.SplitKeepNonEmpty("a,,b, ,c", ",")
    Dim n As Long: n = UBound(r) - LBound(r) + 1
    modTestRunner.Check "experiment_string_array", (n = 3) And (r(0) = "a") And (r(2) = "c"), "n=" & n
End Sub
