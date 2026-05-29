Attribute VB_Name = "modExtractor"
Option Explicit

' ============================================================================
' modExtractor - text extraction dispatcher
' ----------------------------------------------------------------------------
' Routes by file extension. Week 1 supports .txt only. Week 2 adds:
'   .pdf  -> modExtractorWord (Word's PDF Reflow)
'   .docx -> modExtractorWord
'   .xlsx -> modExtractorExcel
' ============================================================================

Public Type ExtractedPage
    Page As Long
    Text As String
End Type

Public Function ExtractFile(ByVal path As String, ByRef outSource As String) As ExtractedPage()
    Dim ext As String
    ext = LCase$(Mid$(path, InStrRev(path, ".") + 1))
    outSource = Mid$(path, InStrRev(path, "\") + 1)

    Select Case ext
        Case "txt", "md", "csv"
            ExtractFile = ExtractPlainText(path)
        Case "pdf", "docx", "doc"
            Err.Raise vbObjectError + &H7001, "modExtractor", _
                "PDF/Word extraction lands in Week 2 (Word.Application Reflow)."
        Case "xlsx", "xlsm", "xls"
            Err.Raise vbObjectError + &H7002, "modExtractor", _
                "Excel extraction lands in Week 2."
        Case Else
            Err.Raise vbObjectError + &H7003, "modExtractor", _
                "Unsupported file type: " & ext
    End Select
End Function

Private Function ExtractPlainText(ByVal path As String) As ExtractedPage()
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    Dim text As String: text = st.ReadText
    st.Close

    Dim pages(0 To 0) As ExtractedPage
    pages(0).page = 1
    pages(0).Text = text
    ExtractPlainText = pages
End Function
