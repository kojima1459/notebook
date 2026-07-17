Attribute VB_Name = "modP2P"
Option Explicit

' ============================================================================
' modP2P - Phase 4: P2P 感謝状(EXP交換)+ AD連携ユーザーID
' ----------------------------------------------------------------------------
' 不正防止の核: 「感謝EXP(thumbup)」は "他者のExcelから共有フォルダ経由で
' 感謝状(✅解決由来)を受領した時のみ" 加算する。自分で自分の回答に✅を押しても
' 感謝EXPは一切増えない(自己申告不可)。活動EXP(質問/登録/共有)はローカルで
' 貯まり、感謝EXPと合算してレベルになる(オーナー裁定: 活動+感謝の合算型)。
'
' 仕組み:
'   ・出所: パック取込時に各チャンクの origin が "pack:<作者>" になっている
'     (modPack実装済み)。この作者情報をそのまま「誰へ感謝するか」に使う。
'   ・発行: ✅解決時、直近回答の最上位引用ソースの作者(自分以外)へ、共有フォルダ
'     <nexus_share_path>\thanks\ に感謝状(1件1ファイル・nonce付きTSV)を書く。
'   ・受領: 同期/起動時に thanks\ を走査し、宛先=自分かつ未処理(nonce重複排除)
'     の感謝状ごとに感謝EXP(+exp_thumbup)を加算する。
'
' 実行環境: Windows版Excelのみ。ADSystemInfo/ADODB.Stream/共有フォルダは実機
' 前提のため、本モジュールは実機(2台+共有フォルダ)での結合テストが必要。
' 本リポジトリのlint/LO構文チェックは構文のみ検証(実行検証は不可)。
' ============================================================================

Private Const THANKS_SUBDIR As String = "thanks"   ' サブフォルダ名(区切り"\"は連結時に付与)
Private mUserIdCache As String

' ----------------------------------------------------------------------------
' CurrentUserId - AD(ADSystemInfo)のCN → 失敗時 USERNAME。共有内での宛先/差出人。
' ----------------------------------------------------------------------------
Public Function CurrentUserId() As String
    If LenB(mUserIdCache) > 0 Then
        CurrentUserId = mUserIdCache
        Exit Function
    End If

    Dim uid As String
    On Error Resume Next
    Dim adsi As Object
    Set adsi = CreateObject("ADSystemInfo")
    uid = CnFromDn(CStr(adsi.UserName))   ' 例 "CN=山田 太郎,OU=..,DC=.."
    On Error GoTo 0

    If LenB(uid) = 0 Then uid = Environ$("USERNAME")
    If LenB(uid) = 0 Then uid = "user"
    uid = SanitizeId(uid)

    mUserIdCache = uid
    CurrentUserId = uid
End Function

Private Function CnFromDn(ByVal dn As String) As String
    If LenB(dn) = 0 Then Exit Function
    Dim p As Long: p = InStr(1, dn, "CN=", vbTextCompare)
    If p = 0 Then Exit Function
    Dim rest As String: rest = Mid$(dn, p + 3)
    Dim c As Long: c = InStr(rest, ",")
    If c > 0 Then rest = Left$(rest, c - 1)
    CnFromDn = Trim$(rest)
End Function

Private Function SanitizeId(ByVal s As String) As String
    Dim bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", ",", vbTab, vbCr, vbLf)
    Dim t As String: t = s
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        t = Replace(t, CStr(bad(i)), "_")
    Next i
    SanitizeId = modUtil.SafeLeft(Trim$(t), 64)
End Function

' ----------------------------------------------------------------------------
' EmitThanksForLastAnswer - ✅解決時に modAsk.FeedbackGreen から呼ぶ。
'   直近回答の最上位ソースの作者(自分以外・出所判明)へ感謝状を書き出す。
' ----------------------------------------------------------------------------
Public Sub EmitThanksForLastAnswer()
    On Error Resume Next
    EmitThanks modAsk.LastTopSource()
    On Error GoTo 0
End Sub

Public Sub EmitThanks(ByVal topSource As String)
    On Error GoTo Done
    If LenB(topSource) = 0 Then Exit Sub

    Dim author As String: author = AuthorOfSource(topSource)
    If LenB(author) = 0 Then Exit Sub                          ' 自作/出所不明 → 感謝なし
    If StrComp(author, "不明", vbTextCompare) = 0 Then Exit Sub

    Dim myId As String: myId = CurrentUserId()
    If StrComp(author, myId, vbTextCompare) = 0 Then Exit Sub  ' 自分自身には送らない

    Dim folderPath As String: folderPath = ThanksDir()
    If LenB(folderPath) = 0 Then Exit Sub
    EnsureDir folderPath

    Dim nonce As String: nonce = NewNonce(myId)
    Dim rowText As String
    rowText = nonce & vbTab & myId & vbTab & author & vbTab & _
              SanitizeField(topSource) & vbTab & modUtil.NowStamp()

    WriteUtf8Line folderPath & "thx_" & author & "_" & nonce & ".txt", rowText
Done:
End Sub

' ----------------------------------------------------------------------------
' CollectThanks - 同期/起動時に呼ぶ。宛先=自分の未処理感謝状ごとに感謝EXP加算。
'   戻り値=今回付与した件数(dedupはnonceで実施)。
' ----------------------------------------------------------------------------
Public Function CollectThanks(Optional ByVal silent As Boolean = False) As Long
    On Error GoTo Done
    Dim folderPath As String: folderPath = ThanksDir()
    If LenB(folderPath) = 0 Then Exit Function

    Dim myId As String: myId = CurrentUserId()
    Dim awarded As Long: awarded = 0

    ' 宛先=自分のファイルだけを列挙(ファイル名に宛先を埋め込んでいる)
    Dim fn As String: fn = Dir(folderPath & "thx_" & myId & "_*.txt")
    Do While LenB(fn) > 0
        Dim rec As String: rec = ReadUtf8All(folderPath & fn)
        Dim f() As String: f = Split(rec, vbTab)
        If UBound(f) >= 4 Then
            Dim nonce As String: nonce = f(0)
            Dim toU As String: toU = f(2)
            If StrComp(toU, myId, vbTextCompare) = 0 Then
                If Not SeenNonce(nonce) Then
                    MarkNonce nonce
                    modStats.AddExp "thumbup"
                    modStats.Bump "thanks_received_total"
                    awarded = awarded + 1
                End If
            End If
        End If
        fn = Dir()   ' 次のファイル(このループ内で他のDir()を呼ばないこと)
    Loop

    If awarded > 0 And Not silent Then
        MsgBox awarded & "件の「ありがとう」が届きました。" & vbLf & _
               "あなたが共有したナレッジが、誰かの役に立っています。" & vbLf & _
               "(感謝EXP +" & (awarded * modConfig.GetLong("exp_thumbup", 10)) & ")", _
               vbInformation, modAppDef.APP_NAME
    End If
    CollectThanks = awarded
Done:
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' ソース名 → 作者(origin が "pack:<作者>" のときのみ。自作等は "")
Private Function AuthorOfSource(ByVal srcName As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim r As Long
    For r = 2 To lastR
        If StrComp(CStr(ws.Cells(r, 2).Value), srcName, vbTextCompare) = 0 Then
            Dim origin As String: origin = CStr(ws.Cells(r, 3).Value)
            If LCase$(Left$(origin, 5)) = "pack:" Then AuthorOfSource = Trim$(Mid$(origin, 6))
            Exit Function
        End If
    Next r
End Function

Private Function ThanksDir() As String
    Dim basePath As String: basePath = modConfig.GetString("nexus_share_path", "")
    If LenB(basePath) = 0 Then Exit Function
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
    ThanksDir = basePath & THANKS_SUBDIR & "\"
End Function

Private Sub EnsureDir(ByVal folderPath As String)
    On Error Resume Next
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    On Error GoTo 0
End Sub

Private Function NewNonce(ByVal myId As String) As String
    NewNonce = myId & "-" & Format$(Now, "yyyymmddhhnnss") & "-" & _
               Format$(Int(Timer * 1000) Mod 100000, "00000")
End Function

Private Function SeenNonce(ByVal nonce As String) As Boolean
    SeenNonce = (modStats.GetStat("thx:" & nonce) > 0)
End Function

Private Sub MarkNonce(ByVal nonce As String)
    modStats.Bump "thx:" & nonce
End Sub

Private Function SanitizeField(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace(t, vbTab, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    SanitizeField = t
End Function

Private Sub WriteUtf8Line(ByVal filePath As String, ByVal content As String)
    On Error Resume Next
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2          ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText content
    st.SaveToFile filePath, 2   ' adSaveCreateOverWrite
    st.Close
    On Error GoTo 0
End Sub

Private Function ReadUtf8All(ByVal filePath As String) As String
    On Error Resume Next
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile filePath
    ReadUtf8All = st.ReadText
    st.Close
    On Error GoTo 0
End Function
