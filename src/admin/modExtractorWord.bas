Attribute VB_Name = "modExtractorWord"
Option Explicit

' ============================================================================
' modExtractorWord - PDF / Word extraction via the Word.Application COM
' ----------------------------------------------------------------------------
' Word 2013+ opens PDFs by "reflowing" them into editable text. Reflow is
' lossy on heavily formatted PDFs (tables, multi-column约款) but recovers
' enough plain text for embedding.
'
' For .docx / .doc files Word just opens them natively.
'
' We extract per-page text by scanning Document.Range until we hit a page
' break, so chunks can be tagged with the correct page number.
' ============================================================================

Public Function Extract(ByVal path As String) As ExtractedPage()
    Dim word As Object
    Set word = CreateObject("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0 ' wdAlertsNone

    Dim doc As Object
    On Error GoTo Failed
    ' ConfirmConversions:=False suppresses the PDF Reflow prompt.
    Set doc = word.Documents.Open( _
        FileName:=path, _
        ConfirmConversions:=False, _
        ReadOnly:=True, _
        AddToRecentFiles:=False, _
        Visible:=False)

    Dim pageCount As Long
    pageCount = doc.ComputeStatistics(2) ' wdStatisticPages
    If pageCount < 1 Then pageCount = 1

    Dim pages() As ExtractedPage
    ReDim pages(0 To pageCount - 1)

    Dim i As Long
    For i = 1 To pageCount
        pages(i - 1).page = i
        pages(i - 1).Text = ExtractPageText(doc, i, pageCount)
    Next i

    doc.Close 0 ' wdDoNotSaveChanges
    word.Quit 0
    Extract = pages
    Exit Function

Failed:
    Dim msg As String: msg = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close 0
    word.Quit 0
    On Error GoTo 0
    Err.Raise vbObjectError + &H7101, "modExtractorWord", _
        "Word extraction failed for " & path & ": " & msg
End Function

Private Function ExtractPageText(ByVal doc As Object, _
                                 ByVal pageNum As Long, _
                                 ByVal totalPages As Long) As String
    ' wdGoToPage = 1, wdGoToAbsolute = 1
    Dim startRange As Object, endRange As Object
    Set startRange = doc.GoTo(What:=1, Which:=1, count:=pageNum)
    If pageNum < totalPages Then
        Set endRange = doc.GoTo(What:=1, Which:=1, count:=pageNum + 1)
        startRange.End = endRange.Start - 1
    Else
        startRange.End = doc.Content.End
    End If
    ExtractPageText = startRange.Text
End Function
