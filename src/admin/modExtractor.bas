Attribute VB_Name = "modExtractor"
Option Explicit

' ============================================================================
' modExtractor - text extraction dispatcher
' ----------------------------------------------------------------------------
' Routes by file extension to the right specialist module:
'   .txt / .md / .csv   -> plain ADODB.Stream UTF-8 read
'   .pdf                -> modExtractorWord (Word PDF Reflow), with
'                          modExtractorAcrobat as fallback
'   .docx / .doc        -> modExtractorWord
'   .xlsx / .xls / .xlsm-> modExtractorExcel
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
        Case "pdf"
            ExtractFile = ExtractPdfWithFallback(path)
        Case "docx", "doc"
            ExtractFile = modExtractorWord.Extract(path)
        Case "xlsx", "xls", "xlsm"
            ExtractFile = modExtractorExcel.Extract(path)
        Case Else
            Err.Raise vbObjectError + &H7003, "modExtractor", _
                "Unsupported file type: " & ext
    End Select
End Function

Private Function ExtractPdfWithFallback(ByVal path As String) As ExtractedPage()
    On Error GoTo TryAcrobat
    ExtractPdfWithFallback = modExtractorWord.Extract(path)
    Exit Function
TryAcrobat:
    Dim wordErr As String: wordErr = Err.Description
    On Error GoTo NoFallback
    ExtractPdfWithFallback = modExtractorAcrobat.Extract(path)
    Exit Function
NoFallback:
    Err.Raise vbObjectError + &H7004, "modExtractor", _
        "Both Word and Acrobat extraction failed for " & path & _
        " (Word: " & wordErr & ", Acrobat: " & Err.Description & ")"
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
