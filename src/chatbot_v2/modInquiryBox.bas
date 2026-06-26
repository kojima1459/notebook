Attribute VB_Name = "modInquiryBox"
Option Explicit

' ============================================================================
' modInquiryBox - 「教えてBOX」本社引受部門への照会パケット生成
' ----------------------------------------------------------------------------
' 目的 (本ツールの核心): 営業が雑に本社へ照会して双方の時間を浪費する問題を、
' 「AIが調べ尽くした上で、整理された照会パケットを自動生成」することで解決する。
' 本社は一から聞き返す工程が消え、営業も論点が整理された状態で照会できる。
'
' パケット内容:
'   - 照会者/部署/日時
'   - 営業の元の質問
'   - AIが解釈した意図・前提
'   - AIが調べた範囲(参照した出典)
'   - AIの暫定回答と自信度
'   - 営業が判断に迷っている点(任意入力)
'
' 送信(両対応):
'   1. Outlook が使えれば メール下書きを自動生成して表示 (宛先=config hq_inquiry_email)
'   2. 使えなければ「照会パケット(コピー用)」シートに出力 + クリップボードへコピー
' どちらも失敗しても落ちない。
' ============================================================================

Private Const COPY_SHEET As String = "照会パケット(コピー用)"

Public Sub OpenInquiryBox()
    On Error GoTo Trap

    If LenB(modBoot.gLastQuestion) = 0 Then
        MsgBox "先に質問を送信し、回答を受け取ってから『教えてBOX』を使ってください。", _
               vbInformation, "教えてBOX"
        Exit Sub
    End If

    ' Optional: what is the user unsure about?
    Dim doubt As String
    doubt = InputBox( _
        "本社へ照会する前に — どこに判断が迷っていますか？(任意)" & vbCrLf & _
        "（例: この前提で合っているか / 約款の解釈が複数あり得る / 事例が見当たらない 等）" & vbCrLf & vbCrLf & _
        "空欄でも照会パケットは作成できます。", _
        "教えてBOX - 照会内容の補足")

    Dim body As String: body = BuildPacket(doubt)
    Dim subject As String
    subject = "【AI照会】" & modUserProfile.CurrentDeptName() & " " & modUserProfile.CurrentUser() & _
              " / " & Left$(Replace(modBoot.gLastQuestion, vbLf, " "), 40)

    Dim hqMail As String: hqMail = modConfig.GetString("hq_inquiry_email", "")

    ' --- Try Outlook draft first ---
    If TryOutlookDraft(hqMail, subject, body) Then
        MsgBox "Outlook で照会メールの下書きを開きました。" & vbCrLf & _
               "内容を確認し、必要に応じて加筆して送信してください。", vbInformation, "教えてBOX"
        Exit Sub
    End If

    ' --- Fallback: copy sheet + clipboard ---
    WriteCopySheet subject, body
    Dim copied As Boolean: copied = TryClipboard(body)

    Dim msg As String
    msg = "『" & COPY_SHEET & "』シートに照会パケットを出力しました。" & vbCrLf
    If copied Then msg = msg & "クリップボードにもコピー済みです。" & vbCrLf
    msg = msg & vbCrLf & "Teams / メール等に貼り付けて本社へ送ってください。"
    If LenB(hqMail) > 0 Then msg = msg & vbCrLf & "(照会先: " & hqMail & ")"
    MsgBox msg, vbInformation, "教えてBOX"
    Exit Sub

Trap:
    modDiag.ReportError "modInquiryBox.OpenInquiryBox", Err.Number, Err.Description, _
        "照会パケットの生成中にエラーが発生しました。"
End Sub

' ----------------------------------------------------------------------------
Private Function BuildPacket(ByVal doubt As String) As String
    Dim nl As String: nl = vbCrLf
    Dim s As String
    s = "■ 本社引受部門への照会（AIアシスタント経由）" & nl
    s = s & "────────────────────────────" & nl
    s = s & "照会者　: " & modUserProfile.CurrentUser() & nl
    s = s & "部署　　: " & modUserProfile.CurrentDeptName() & nl
    s = s & "役割　　: " & modUserProfile.CurrentRole() & nl
    s = s & "日時　　: " & Format$(Now, "yyyy-mm-dd hh:nn") & nl
    s = s & "AI自信度: " & DisplayConfidence(modBoot.gLastConfidence) & nl & nl

    s = s & "【営業からの照会】" & nl & modBoot.gLastQuestion & nl & nl

    If LenB(modBoot.gLastIntent) > 0 Then
        s = s & "【AIが解釈した意図】" & nl & modBoot.gLastIntent & nl & nl
    End If
    If LenB(modBoot.gLastAssumptions) > 0 Then
        s = s & "【AIが置いた前提】" & nl & modBoot.gLastAssumptions & nl & nl
    End If

    s = s & "【AIの暫定回答】" & nl & modBoot.gLastAnswer & nl & nl

    If LenB(modBoot.gLastCitations) > 0 Then
        s = s & "【AIが参照した社内ナレッジ(出典)】" & nl & modBoot.gLastCitations & nl & nl
    End If

    If LenB(doubt) > 0 Then
        s = s & "【照会者が判断に迷っている点】" & nl & doubt & nl & nl
    End If

    s = s & "────────────────────────────" & nl
    s = s & "※ 本パケットは社内AIアシスタントが自動生成しました。AIの暫定回答は" & nl
    s = s & "　 必ずしも正確ではありません。上記前提・出典をご確認の上ご回答ください。" & nl
    BuildPacket = s
End Function

Private Function DisplayConfidence(ByVal c As String) As String
    Select Case c
        Case "高": DisplayConfidence = "高（AIは約款根拠で断定できると判断）"
        Case "中": DisplayConfidence = "中（事例・推論が主。要裏付け）"
        Case "低": DisplayConfidence = "低（該当規定が乏しく要確認）"
        Case Else: DisplayConfidence = "（不明）"
    End Select
End Function

' ----------------------------------------------------------------------------
Private Function TryOutlookDraft(ByVal toAddr As String, ByVal subject As String, _
                                 ByVal body As String) As Boolean
    On Error GoTo Fail
    Dim ol As Object: Set ol = CreateObject("Outlook.Application")
    Dim mail As Object: Set mail = ol.CreateItem(0)   ' olMailItem
    If LenB(toAddr) > 0 Then mail.To = toAddr
    mail.subject = subject
    mail.body = body
    mail.Display                                       ' show draft; do NOT auto-send
    TryOutlookDraft = True
    Exit Function
Fail:
    TryOutlookDraft = False
End Function

Private Sub WriteCopySheet(ByVal subject As String, ByVal body As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(COPY_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(1))
        ws.Name = COPY_SHEET
    End If
    ws.Visible = xlSheetVisible
    ws.Cells.Clear
    ws.Columns("A").ColumnWidth = 100
    With ws.Range("A1")
        .value = "↓ この内容を選択してコピーし、Teams/メールで本社へ送ってください（件名: " & subject & "）"
        .Font.Bold = True
        .Font.color = RGB(60, 90, 160)
    End With
    With ws.Range("A3")
        .value = body
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Size = 10
    End With
    ws.Range("A3").RowHeight = 360
    On Error Resume Next
    ws.Activate
    ws.Range("A3").Select
    On Error GoTo 0
End Sub

' Best-effort clipboard copy via DataObject. May be blocked on hardened PCs;
' failure is non-fatal (the copy sheet is the reliable path).
Private Function TryClipboard(ByVal text As String) As Boolean
    On Error GoTo Fail
    Dim dobj As Object
    Set dobj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  ' MSForms.DataObject
    dobj.SetText text
    dobj.PutInClipboard
    TryClipboard = True
    Exit Function
Fail:
    TryClipboard = False
End Function
