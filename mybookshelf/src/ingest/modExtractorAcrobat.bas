Attribute VB_Name = "modExtractorAcrobat"
Option Explicit

' ============================================================================
' modExtractorAcrobat - Adobe Acrobat COM経由のPDF抽出(フォールバック)
' ----------------------------------------------------------------------------
' 役割:
'   スキャンPDFや複雑なページ構造を持つ約款PDFはWordのリフローで崩れることが
'   ある。Adobe Acrobat Proが入っている環境ではそのCOM API(AcroPDPage)を
'   使うとページ単位のテキストを得られることがある。
'   Wordは全PC標準搭載だがAcrobatは必ずしも入っていないため、あくまで
'   modExtractorWord失敗時のフォールバック経路として使う。
'
' 流用元: /home/user/notebook/src/admin/modExtractorAcrobat.bas(V2資産・
'   コピー元。変更禁止)。
'
' 適応要件(MASTER_SPEC §7.2):
'   ・V1は Err.Raise で失敗を通知していたが、本モジュールは
'     Boolean 戻り値 + errDetail(ByRef)へ変換する。
'   ・maxPages を超えるページ数のPDFは、超過分を取得せず先頭 maxPages
'     ページのみ処理し、truncated(ByRef)を True にする。
'   ・AcroExch.App / AcroExch.PDDoc は必ず On Error で包み、成功・失敗
'     いずれの経路でも doc.Close / app.Exit を試みる(V1のFailedハンドラの
'     作法を踏襲)。
'   ・§12「文字列連結の長大化は&連鎖でなく配列+Join」に従い、単語列は
'     配列へ集めてから Join で連結する(V1の `text = text & word & " "`
'     というO(n^2)連結を改善)。
' ============================================================================

Public Function Extract(ByVal path As String, ByVal maxPages As Long, _
                        ByRef pages() As ExtractedPage, ByRef truncated As Boolean, _
                        ByRef errDetail As String) As Boolean
    truncated = False

    Dim app As Object, doc As Object

    On Error GoTo Failed
    Set app = CreateObject("AcroExch.App")
    Set doc = CreateObject("AcroExch.PDDoc")
    If Not doc.Open(path) Then
        errDetail = "PDDoc.Open に失敗しました"
        GoTo CleanupFail
    End If

    Dim n As Long: n = doc.GetNumPages()
    If n < 1 Then
        errDetail = "0ページのPDFです"
        GoTo CleanupFail
    End If

    Dim loopCount As Long: loopCount = n
    If loopCount > maxPages Then
        loopCount = maxPages
        truncated = True
    End If

    Dim tmp() As ExtractedPage: ReDim tmp(0 To loopCount - 1)
    Dim i As Long
    For i = 0 To loopCount - 1
        Dim pg As Object
        Set pg = doc.AcquirePage(i)
        tmp(i).page = i + 1
        tmp(i).Text = ExtractPageWords(pg)
        Set pg = Nothing
    Next i

    doc.Close
    On Error Resume Next
    app.Exit
    On Error GoTo 0
    Set doc = Nothing
    Set app = Nothing

    pages = tmp
    Extract = True
    Exit Function

CleanupFail:
    If Not doc Is Nothing Then
        On Error Resume Next
        doc.Close
        On Error GoTo 0
    End If
    If Not app Is Nothing Then
        On Error Resume Next
        app.Exit
        On Error GoTo 0
    End If
    Extract = False
    Exit Function

Failed:
    errDetail = DescribeComError(Err.Number, Err.Description)
    If Not doc Is Nothing Then
        On Error Resume Next
        doc.Close
        On Error GoTo 0
    End If
    If Not app Is Nothing Then
        On Error Resume Next
        app.Exit
        On Error GoTo 0
    End If
    Extract = False
End Function

' 1ページ分の単語をAcrobatから取り出し、配列+Joinでスペース区切り連結する
' (§12: &連鎖の長大化禁止)。
Private Function ExtractPageWords(ByVal pg As Object) As String
    Dim numWords As Long: numWords = pg.GetNumWords()
    If numWords < 1 Then Exit Function

    Dim parts() As String: ReDim parts(0 To numWords - 1)
    Dim w As Long
    For w = 0 To numWords - 1
        parts(w) = pg.GetWord(w)
    Next w
    ExtractPageWords = Join(parts, " ")
End Function

' Mac等COM不可環境向けの丁寧な案内文を生成する(§13)。
Private Function DescribeComError(ByVal errNum As Long, ByVal desc As String) As String
    If errNum = 429 Then
        DescribeComError = "この環境ではAcrobat連携(COM)が利用できません。" & _
            "Mac版ExcelやCOM未対応環境、またはAcrobat Pro未インストールの可能性があります。" & _
            "Windows版Excel+Acrobat Proでお試しください。(詳細: " & desc & ")"
    Else
        DescribeComError = desc
    End If
End Function
