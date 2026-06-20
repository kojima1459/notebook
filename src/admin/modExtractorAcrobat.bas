Attribute VB_Name = "modExtractorAcrobat"
Option Explicit

' ============================================================================
' modExtractorAcrobat - PDF extraction via Adobe Acrobat COM (fallback)
' ----------------------------------------------------------------------------
' Some約款 PDFs are scans or use complex page structures that defeat Word's
' reflow. When Adobe Acrobat Pro is installed we can drive its COM API,
' which gives per-page text via the AcroPDPage object.
'
' This is a fallback path; modExtractorWord is preferred because Word is
' on every PC, Acrobat is not.
' ============================================================================

Public Function Extract(ByVal path As String) As ExtractedPage()
    Dim app As Object, doc As Object
    On Error GoTo Failed
    Set app = CreateObject("AcroExch.App")
    Set doc = CreateObject("AcroExch.PDDoc")
    If Not doc.Open(path) Then
        Err.Raise vbObjectError + &H7201, "modExtractorAcrobat", "PDDoc.Open failed"
    End If

    Dim n As Long: n = doc.GetNumPages()
    If n < 1 Then
        doc.Close
        Err.Raise vbObjectError + &H7202, "modExtractorAcrobat", "Zero-page PDF"
    End If

    Dim pages() As ExtractedPage
    ReDim pages(0 To n - 1)

    Dim i As Long
    For i = 0 To n - 1
        Dim page As Object, selection As Object, text As String
        Set page = doc.AcquirePage(i)
        Dim numWords As Long
        numWords = page.GetNumWords()
        Dim w As Long
        text = ""
        For w = 0 To numWords - 1
            text = text & page.GetWord(w) & " "
        Next w
        pages(i).page = i + 1
        pages(i).Text = text
    Next i

    doc.Close
    On Error Resume Next
    app.Exit
    On Error GoTo 0
    Extract = pages
    Exit Function

Failed:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close
    If Not app Is Nothing Then app.Exit
    On Error GoTo 0
    Err.Raise vbObjectError + &H7203, "modExtractorAcrobat", _
        "Acrobat extraction failed: " & Err.Description
End Function
