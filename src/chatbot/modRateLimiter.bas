Attribute VB_Name = "modRateLimiter"
Option Explicit

' ============================================================================
' modRateLimiter - per-user client-side rate limiting
' ----------------------------------------------------------------------------
' Stores rolling timestamps in %LOCALAPPDATA%\InternalNotebookLM\ratelimit.dat.
' Hour window and day window are independent counters. On exceedance, the
' chat UI must block the send and surface the wait time.
' ============================================================================

Private Const DAT_FILE As String = "ratelimit.dat"

Public Function CanSend(ByRef outRetryAfterSec As Long) As Boolean
    Dim perHour As Long, perDay As Long
    perHour = modConfig.GetLong("rate_limit", "per_hour", 30)
    perDay = modConfig.GetLong("rate_limit", "per_day", 100)

    Dim stamps() As Double
    Dim n As Long
    LoadStamps stamps, n

    Dim now_ As Double: now_ = Now
    Dim hourCutoff As Double, dayCutoff As Double
    hourCutoff = now_ - (1# / 24#)
    dayCutoff = now_ - 1#

    Dim countHour As Long, countDay As Long, oldestHour As Double, oldestDay As Double
    oldestHour = now_: oldestDay = now_
    Dim i As Long
    For i = 0 To n - 1
        If stamps(i) >= dayCutoff Then
            countDay = countDay + 1
            If stamps(i) < oldestDay Then oldestDay = stamps(i)
        End If
        If stamps(i) >= hourCutoff Then
            countHour = countHour + 1
            If stamps(i) < oldestHour Then oldestHour = stamps(i)
        End If
    Next i

    If countHour >= perHour Then
        outRetryAfterSec = CLng((oldestHour + (1# / 24#) - now_) * 86400#)
        If outRetryAfterSec < 1 Then outRetryAfterSec = 1
        CanSend = False
        Exit Function
    End If
    If countDay >= perDay Then
        outRetryAfterSec = CLng((oldestDay + 1# - now_) * 86400#)
        If outRetryAfterSec < 1 Then outRetryAfterSec = 1
        CanSend = False
        Exit Function
    End If
    outRetryAfterSec = 0
    CanSend = True
End Function

Public Sub RecordSend()
    Dim stamps() As Double
    Dim n As Long
    LoadStamps stamps, n
    Dim dayCutoff As Double: dayCutoff = Now - 1#
    ' Compact: drop entries older than a day
    Dim kept() As Double
    ReDim kept(0 To n)
    Dim k As Long: k = 0
    Dim i As Long
    For i = 0 To n - 1
        If stamps(i) >= dayCutoff Then
            kept(k) = stamps(i)
            k = k + 1
        End If
    Next i
    kept(k) = Now
    k = k + 1
    SaveStamps kept, k
End Sub

Private Sub LoadStamps(ByRef out() As Double, ByRef count As Long)
    Dim path As String: path = modPaths.AppDataRoot() & "\" & DAT_FILE
    If Len(Dir$(path)) = 0 Then
        ReDim out(0 To 0)
        count = 0
        Exit Sub
    End If
    Dim fnum As Integer: fnum = FreeFile
    Open path For Binary Access Read As #fnum
    Dim flen As Long: flen = LOF(fnum)
    count = flen \ 8
    If count > 0 Then
        ReDim out(0 To count - 1)
        Get #fnum, 1, out
    Else
        ReDim out(0 To 0)
    End If
    Close #fnum
End Sub

Private Sub SaveStamps(ByRef arr() As Double, ByVal count As Long)
    modPaths.EnsureDir modPaths.AppDataRoot()
    Dim path As String: path = modPaths.AppDataRoot() & "\" & DAT_FILE
    Dim fnum As Integer: fnum = FreeFile
    Open path For Binary Access Write As #fnum
    If count > 0 Then
        Dim trimmed() As Double
        ReDim trimmed(0 To count - 1)
        Dim i As Long
        For i = 0 To count - 1
            trimmed(i) = arr(i)
        Next i
        Put #fnum, 1, trimmed
    End If
    Close #fnum
End Sub
