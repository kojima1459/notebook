Attribute VB_Name = "modExtractorAcrobat"
Option Explicit

' 1ページから取り出す単語数の上限(壊れたPDFでの暴走防止)。
Private Const MAX_WORDS_PER_PAGE As Long = 20000

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

    ' 2026-07-28(レビュー M-13): テキスト取り出しは PDDoc の JSObject 経由。
    ' 取得は文書につき1回でよい(ページごとに取り直す必要は無い)。
    Dim js As Object
    On Error Resume Next
    Set js = doc.GetJSObject()
    On Error GoTo Failed
    If js Is Nothing Then
        errDetail = "Acrobat の JavaScript オブジェクトを取得できませんでした" & _
                    "(Acrobat Reader では利用できません。Acrobat Pro が必要です)"
        GoTo CleanupFail
    End If

    Dim tmp() As ExtractedPage: ReDim tmp(0 To loopCount - 1)
    Dim i As Long
    For i = 0 To loopCount - 1
        tmp(i).page = i + 1
        tmp(i).Text = ExtractPageWords(js, i)
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
    Set doc = Nothing
    Set app = Nothing
    Extract = False
    Exit Function

Failed:
    errDetail = DescribeComError(Err.Number, Err.Description)
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailedCleanup0
FailedCleanup0:
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
    Set doc = Nothing
    Set app = Nothing
    Extract = False
End Function

' 1ページ分の単語をAcrobatから取り出し、配列+Joinでスペース区切り連結する
' (§12: &連鎖の長大化禁止)。
Private Function ExtractPageWords(ByVal js As Object, ByVal pageIndex As Long) As String
    ' 2026-07-28(レビュー M-13): PDPage.GetNumWords / GetWord は
    ' Acrobat IAC(COM)に存在しないメソッドで、エラー438が確実に返っていた。
    ' つまり「Word失敗 → Acrobatで救済」という経路は、Acrobat Pro が
    ' 入っている端末でも一度も機能していない死に経路だった。
    ' 正規の手段は JSObject(Acrobat JavaScript)の
    ' getPageNumWords / getPageNthWord を叩くこと。
    If js Is Nothing Then Exit Function

    Dim numWords As Long
    On Error Resume Next
    numWords = CLng(js.getPageNumWords(pageIndex))
    On Error GoTo 0
    If numWords < 1 Then Exit Function
    If numWords > MAX_WORDS_PER_PAGE Then numWords = MAX_WORDS_PER_PAGE

    Dim parts() As String: ReDim parts(0 To numWords - 1)
    Dim w As Long
    On Error Resume Next
    For w = 0 To numWords - 1
        parts(w) = CStr(js.getPageNthWord(pageIndex, w, True))
    Next w
    On Error GoTo 0
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
